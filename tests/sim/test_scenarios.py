"""Deterministic scenarios with known expected outcomes.

These validate the *oracle* as much as the code: each one encodes a situation
we have already reasoned about (or hit in production), so a disagreement here
means the sim is wrong, not the connector — except the strict-xfail marker,
which encodes a known real defect.
"""

from __future__ import annotations

import pytest
from invariants import drain_and_assert
from sim_graph import INBOX, PROCESSED, Driver, Finding


def make(seed: int = 1, batch: int | None = 50, cap: int = 100) -> Driver:
    return Driver(seed, batch, cap=cap)


# 1 ---------------------------------------------------------------------------
def test_same_second_cohort_drains_exactly():
    """600 messages sharing one truncated second, batch 50.

    The old cursor design could not do this (it compared truncated timestamps);
    the seen-set is id-based, so each cycle must publish exactly 50 — even
    though the sim reshuffles same-second listing order every cycle.
    """
    d = make(batch=50)
    try:
        d.sim.deliver(600, same_second=True)
        for _ in range(12):
            summary = d.run_cycle()
            assert (summary.fetched, summary.published) == (50, 50)
        assert len(d.consumer.first_seen) == 600
        assert not d.consumer.duplicates
        drain_and_assert(d)
    finally:
        d.close()


# 2 ---------------------------------------------------------------------------
def test_visibility_lag_arrival_is_not_lost():
    """Mail stamped earlier but visible later — the class the cursor dropped."""
    d = make(batch=10)
    try:
        (fresh,) = d.sim.deliver(1)
        (late,) = d.sim.deliver(1, backdate_s=90.0, lag_cycles=2)

        d.run_cycle()
        assert set(d.consumer.first_seen) == {fresh}, "only the visible one"

        d.run_cycle()  # `late` becomes visible here
        assert late in d.consumer.first_seen, "late-visible older mail must publish"
        drain_and_assert(d)
    finally:
        d.close()


# 3a --------------------------------------------------------------------------
def test_mover_race_transient_404_skips_only_that_message():
    d = make(batch=5)
    try:
        imids = d.sim.deliver(5)
        d.sim.move_out(imids[2], when=("post_fetch", 2), transient_404=True)
        d.grant_moverace_credit(imids[2])

        summary = d.run_cycle()

        assert summary.error is None, "a 404 mid-batch must not fail the cycle"
        assert summary.published == 4
        assert imids[2] not in d.consumer.first_seen
        drain_and_assert(d)
    finally:
        d.close()


# 3b --------------------------------------------------------------------------
def test_mover_race_5xx_aborts_whole_cycle_then_recovers():
    """Documented cost: a non-404 fetch error discards the already-fetched
    prefix too (poll_mailbox is fetch-all-then-publish)."""
    d = make(batch=5)
    try:
        imids = d.sim.deliver(5)
        d.sim.inject("get_email", status=503, mid=d.sim.by_imid(imids[2]).mid)

        first = d.run_cycle()
        assert (first.fetched, first.published) == (0, 0)
        assert first.error_source == "graph"
        assert not d.consumer.first_seen, "nothing published despite 2 successful fetches"

        second = d.run_cycle()
        assert second.published == 5
        assert not d.consumer.duplicates
        drain_and_assert(d)
    finally:
        d.close()


# 3c --------------------------------------------------------------------------
def test_attachment_404_stops_batch_unlike_message_404():
    """Asymmetry worth naming: get_email tolerates 404 (skip), but the same
    404 from get_attachments stops the batch."""
    d = make(batch=5)
    try:
        imids = d.sim.deliver(3, att_spec="normal")
        d.sim.inject("get_attachments", status=404, mid=d.sim.by_imid(imids[1]).mid)

        summary = d.run_cycle()

        assert summary.error_source == "graph"
        assert summary.published == 1, "stop-batch: only the message before the failure"
        assert summary.dropped == 2
        drain_and_assert(d)  # recovers once the fault clears
    finally:
        d.close()


# 4 ---------------------------------------------------------------------------
@pytest.mark.xfail(
    strict=True,
    reason="KNOWN: stop-batch policy lets one unpublishable message starve the backlog",
)
def test_bus_poison_starves_backlog():
    """Desired behaviour: mail behind a permanently unpublishable message still
    gets ingested. Today the batch stops at the poison every cycle, forever."""
    d = make(batch=10)
    try:
        imids = d.sim.deliver(4)
        d.consumer.poison.add(imids[0])

        for _ in range(5):
            d.run_cycle(check=False)

        assert set(imids[1:]) <= set(d.consumer.first_seen), "backlog behind poison starved"
    finally:
        d.close()


def test_i7_detector_fires_on_starvation():
    """Oracle self-test: the starvation invariant must catch the wedge above."""
    d = make(batch=10)
    try:
        imids = d.sim.deliver(4)
        d.consumer.poison.add(imids[0])

        with pytest.raises(Finding, match="starvation"):
            for _ in range(5):
                d.run_cycle()
    finally:
        d.close()


def test_bus_transient_outage_recovers():
    d = make(batch=10)
    try:
        d.sim.deliver(3)
        d.consumer.fail_next = 1

        first = d.run_cycle()
        assert first.error_source == "bus"
        assert first.published == 0

        second = d.run_cycle()
        assert second.published == 3
        assert not d.consumer.duplicates
        drain_and_assert(d)
    finally:
        d.close()


# 5 ---------------------------------------------------------------------------
def test_restart_mid_drain_republishes_within_budget():
    d = make(batch=50)
    try:
        d.sim.deliver(120)
        first = d.run_cycle()
        assert first.published == 50

        d.restart()  # seen-set lost — at-least-once by design
        drain_and_assert(d)

        assert len(d.consumer.first_seen) == 120, "everything ingested exactly once"
        assert sum(d.consumer.duplicates.values()) == 50, "only the pre-restart batch repeats"
        assert d.monitor.snapshot().status != "stale"
    finally:
        d.close()


# 6 ---------------------------------------------------------------------------
def test_move_back_republishes_after_prune():
    d = make(batch=10)
    try:
        (imid,) = d.sim.deliver(1)
        d.run_cycle()
        assert imid in d.consumer.first_seen
        mid = d.sim.by_imid(imid).mid
        assert mid in d.poller._published_ids

        d.sim.move_out(imid)
        quiet = d.run_cycle()
        assert quiet.fetched == 0
        assert mid not in d.poller._published_ids, "prune drops ids that left the folder"

        d.sim.move_back(imid)
        d.grant_moveback_credit(imid)
        again = d.run_cycle()

        assert again.published == 1, "a returned message is treated as new"
        assert d.consumer.duplicates[imid] == 1, "consumer dedupes it"
        drain_and_assert(d)
    finally:
        d.close()


def test_folder_and_mailbox_wiring_is_what_the_poller_asks_for():
    """Guard against the sim drifting from the real call shape."""
    d = make(batch=3)
    try:
        d.sim.deliver(1)
        d.run_cycle()
        assert d.sim.search_calls == 1
        assert d.sim.listed_ids, "listing produced stubs"
        assert all(m.folder == INBOX for m in d.sim.msgs.values())
        assert PROCESSED not in {m.folder for m in d.sim.msgs.values()}
    finally:
        d.close()


# 7 — slow Graph: the heartbeat claim under a probe -----------------------------
def test_long_cycle_stays_live_under_probe():
    """A big slow cycle must not read as wedged.

    Each listed stub and each fetch burns 5s of simulated wall time; the health
    staleness window is 70s (3x the 10s poll interval). A 200-message cycle
    therefore spans ~20 minutes of clock — an orchestrator probing /health
    throughout must never see `stale`, which is only true if the poller beats
    during BOTH the listing and the fetch phases.
    """
    d = make(batch=60)
    try:
        d.sim.deliver(200)
        d.sim.slow_s = 5.0
        summary = d.run_cycle()
        assert summary.published == 60
        assert d.health_probes > 200, "probe actually ran mid-cycle"
        drain_and_assert(d)
    finally:
        d.close()


def test_probe_catches_a_missing_heartbeat():
    """Oracle self-test: with beats disabled the probe must trip."""
    d = make(batch=60)
    try:
        d.sim.deliver(200)
        d.sim.slow_s = 5.0
        d.poller.heartbeat = lambda: None  # simulate the pre-fix code
        with pytest.raises(Finding, match="stale"):
            d.run_cycle()
    finally:
        d.close()


# 8 — partial listing must not prune the seen-set --------------------------------
def test_listing_failure_midway_does_not_lose_the_seen_set():
    """If the listing dies on page N the poller has a partial view; pruning
    against it would forget published mail and republish the world."""
    d = make(batch=10)
    try:
        d.sim.deliver(20)
        d.run_cycle()
        seen_before = set(d.poller._published_ids)
        assert len(seen_before) == 10

        d.sim.inject("search", status=503, after_items=3)
        summary = d.run_cycle()

        assert summary.error_source == "graph"
        assert d.poller._published_ids == seen_before, "partial listing must not prune"
        drain_and_assert(d)
        assert sum(d.consumer.duplicates.values()) == 0, "no republish storm"
    finally:
        d.close()


# 9 — ignore_received_before is a deliberate filter, incl. undated mail ----------
def test_ignore_received_before_filters_old_and_undated_mail():
    """Configured backstop: mail stamped before the cutoff is never ingested —
    and so is mail with no receivedDateTime at all, since Graph's `ge` filter
    excludes nulls. Documented here so the trade-off is explicit."""
    d = make(batch=10)
    try:
        old = d.sim.deliver(1, backdate_s=600.0)[0]
        undated = d.sim.deliver(1, variant="no_received")[0]
        fresh = d.sim.deliver(1)[0]
        d.poller.ignore_received_before = d.clock.now() - __import__("datetime").timedelta(seconds=60)

        d.run_cycle(check=False)

        assert fresh in d.consumer.first_seen
        assert old not in d.consumer.first_seen, "pre-cutoff mail filtered by config"
        assert undated not in d.consumer.first_seen, "null receivedDateTime fails a ge filter"
    finally:
        d.close()
