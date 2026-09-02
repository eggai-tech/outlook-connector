"""The oracle: cycle classification, per-cycle invariants, end-of-run drain.

Invariant numbering follows the plan (I1..I7). Every check either holds or
raises Finding with enough context to replay.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

from sim_graph import INBOX, MAILBOX, Driver, Finding


@dataclass
class Expectation:
    clean: bool
    eligible_mids: set[str]
    expected_unseen: list[str]  # sorted, batch-truncated prefix NOT applied
    prefix_imids: list[str]  # what would be published this cycle, in order
    fail_next_before: int
    bus_failures_before: int
    poison_blocks: bool
    ops_pending: bool
    reasons: list[str] = field(default_factory=list)


def _sorted_unseen(driver: Driver) -> list[str]:
    sim, poller = driver.sim, driver.poller
    now = driver.clock.now()
    eligible = [
        m
        for m in sim.msgs.values()
        if m.folder == INBOX and m.visible_at <= now
    ]
    unseen = [m for m in eligible if m.mid not in poller._published_ids]  # noqa: SLF001
    # mirror the poller's sort; the sim's tie-break decides equal keys, but for
    # expectation we only need the *count* and the batch prefix membership by
    # (second, tiebreak) — reuse the sim's listing order key
    import hashlib

    def tiebreak(mid: str) -> str:
        return hashlib.sha1(f"{sim.seed}:{sim.cycle_no}:{mid}".encode()).hexdigest()

    from sim_graph import T0

    unseen.sort(
        key=lambda m: (m.received_listed is not None, m.received_listed or T0, tiebreak(m.mid))
    )
    return [m.mid for m in unseen]


def snapshot_expectation(driver: Driver) -> Expectation:
    sim, consumer = driver.sim, driver.consumer
    unseen_mids = _sorted_unseen(driver)
    batch = driver.batch if driver.batch is not None else len(unseen_mids)
    prefix_imids = [sim.msgs[mid].imid for mid in unseen_mids[:batch]]
    reasons = []
    if sim.faults:
        reasons.append(f"faults queued: {sim.faults}")
    if sim.pending_ops:
        reasons.append(f"ops pending: {sim.pending_ops}")
    if consumer.fail_next:
        reasons.append(f"fail_next={consumer.fail_next}")
    poison_blocks = any(imid in consumer.poison for imid in prefix_imids)
    if poison_blocks:
        reasons.append("poison in publish prefix")
    if any(sim.msgs[mid].transient_404 for mid in unseen_mids):
        reasons.append("transient 404 armed")
    eligible = {
        m.mid
        for m in sim.msgs.values()
        if m.folder == INBOX and m.visible_at <= driver.clock.now()
    }
    return Expectation(
        clean=not reasons,
        eligible_mids=eligible,
        expected_unseen=unseen_mids,
        prefix_imids=prefix_imids,
        fail_next_before=consumer.fail_next,
        bus_failures_before=consumer.bus_failures,
        poison_blocks=poison_blocks,
        ops_pending=bool(sim.pending_ops),
        reasons=reasons,
    )


def _fail(driver: Driver, name: str, detail: str) -> None:
    trace = "\n".join(driver.log[-30:])
    raise Finding(f"{name}: {detail}\n--- last events ---\n{trace}")


def check_cycle(driver: Driver, pre: Expectation, summary) -> None:
    sim, consumer, poller = driver.sim, driver.consumer, driver.poller
    batch = driver.batch

    # I3 — bounds
    if batch is not None and len(sim.get_email_calls) > batch:
        _fail(driver, "I3", f"get_email calls {len(sim.get_email_calls)} > batch {batch}")
    if sim.search_calls > 1:
        _fail(driver, "I3", f"{sim.search_calls} listings in one cycle")
    if sim.listing_completed:
        expected_pages = math.ceil(len(sim.listed_ids) / 100) if sim.listed_ids else 0
        if sim.list_pages != expected_pages:
            _fail(driver, "I3", f"pages {sim.list_pages} != {expected_pages}")
    for mid in sim.get_attachments_calls:
        m = sim.msgs.get(mid)
        if m is not None and m.att_spec is None:
            _fail(driver, "I3", f"attachments fetched for unflagged {mid}")

    # I1 — clean-cycle exactness
    if pre.clean:
        expect = len(pre.expected_unseen) if batch is None else min(len(pre.expected_unseen), batch)
        if summary.fetched != expect or summary.published != expect or summary.dropped:
            _fail(
                driver,
                "I1",
                f"clean cycle: fetched={summary.fetched} published={summary.published} "
                f"dropped={summary.dropped}, expected {expect}",
            )
        if summary.error is not None:
            _fail(driver, "I1", f"clean cycle errored: {summary.error}")

    # I4 — seen-set hygiene (only when listing completed and no mid-cycle ops)
    if sim.listing_completed and not sim.ops_ran_this_cycle:
        seen = poller._published_ids  # noqa: SLF001
        visible = sim.inbox_visible_mids()
        stray = seen - visible
        if stray:
            _fail(driver, "I4", f"seen-set holds ids not in visible inbox: {stray}")

    # I5 — duplicate budget
    for imid, count in consumer.duplicates.items():
        if count > driver.credits[imid]:
            _fail(
                driver,
                "I5",
                f"{imid} republished {count}x with only {driver.credits[imid]} credits",
            )

    # I6 — health coherence
    snap = driver.monitor.snapshot()
    if snap.status == "stale":
        _fail(driver, "I6", "health stale while cycles progress")
    if summary.error is None and snap.status != "ok":
        _fail(driver, "I6", f"errorless cycle but status={snap.status}")
    if summary.error is not None and snap.status != "degraded":
        _fail(driver, "I6", f"error cycle but status={snap.status}")
    bus_delta = consumer.bus_failures - pre.bus_failures_before
    if summary.error_source == "bus" and bus_delta == 0:
        _fail(driver, "I6", "error attributed to bus without a bus failure")
    if summary.error_source == "graph" and bus_delta > 0:
        _fail(driver, "I6", "error attributed to graph during a bus failure")
    for probe in (snap.graph, snap.bus):
        if probe.last_error and (MAILBOX in probe.last_error or "sim-inbox" in probe.last_error):
            _fail(driver, "I6", f"identity leaked into public error: {probe.last_error}")

    # I7 — starvation detector (poison wedging the head of the queue)
    if (
        summary.published == 0
        and summary.error_source == "bus"
        and pre.fail_next_before == 0
        and pre.poison_blocks
        and len(pre.expected_unseen) > len(consumer.poison)
    ):
        driver.starve_streak += 1
        if driver.starve_streak >= 3:
            _fail(
                driver,
                "I7-Finding(stop-batch starvation)",
                "3 consecutive cycles published nothing while non-poisoned backlog waits",
            )
    else:
        driver.starve_streak = 0


def drain_and_assert(driver: Driver) -> None:
    """I2: clear all faults, drain to quiescence, assert completeness."""
    sim, consumer = driver.sim, driver.consumer
    sim.clear_faults()
    sim.pending_ops.clear()
    consumer.poison.clear()
    consumer.fail_next = 0
    for m in sim.msgs.values():
        m.transient_404 = False

    def backlog() -> int:
        """What the *poller* still has to fetch — not what is un-ingested.

        After a restart the seen-set is empty, so the poller re-walks the whole
        visible folder even though most of it is already ingested; sizing the
        drain from un-ingested count would cut it off mid-walk.
        """
        now = driver.clock.now()
        visible = {
            m.mid for m in sim.msgs.values() if m.folder == INBOX and m.visible_at <= now
        }
        return len(visible - driver.poller._published_ids)  # noqa: SLF001

    def not_yet_visible() -> int:
        now = driver.clock.now()
        return len(
            [m for m in sim.msgs.values() if m.folder == INBOX and m.visible_at > now]
        )

    batch = driver.batch or max(backlog(), 1)
    cap = math.ceil((backlog() + not_yet_visible()) / max(batch, 1)) + 12
    # Two consecutive empty cycles, not one: in slow-Graph mode the clock
    # advances *during* a cycle, so mail that was not yet visible when the
    # listing snapshot was taken can become visible before the cycle ends —
    # "nothing pending" right after a cycle does not mean the next one is empty.
    zero_streak = 0
    for _ in range(cap):
        summary = driver.run_cycle()
        zero_streak = zero_streak + 1 if summary.fetched == 0 else 0
        if zero_streak >= 2 and not_yet_visible() == 0:
            break
    final = driver.run_cycle()
    if final.fetched != 0:
        _fail(driver, "I2", f"no quiescence after drain (fetched={final.fetched})")
    missing = {
        m.imid
        for m in sim.msgs.values()
        if m.folder == INBOX and m.imid not in consumer.first_seen
    }
    if missing:
        _fail(driver, "I2", f"dropped mail never ingested: {missing}")
