"""Simulated Outlook/Graph mailbox, consumer and driver for fuzzing the poller.

Drives the REAL Poller + run_workflow through the duck-typed client seam.
Models the Graph quirks that produced real bugs: second-truncated timestamps
in listings vs microsecond stored values, visibility (replication) lag,
same-second cohorts with per-cycle tie-order reshuffle, mover races between
listing and fetch, transient/permanent 404s, 429/503/network faults.

Deterministic given (seed, event sequence) — no wall clock anywhere.
"""

from __future__ import annotations

import asyncio
import datetime as dt
import hashlib
import math
from collections import Counter
from dataclasses import dataclass

import httpx
from outlook_helper import GraphError
from outlook_helper.schemas import EmailAddress, OutlookAttachment, OutlookBody, OutlookMessage

from outlook_connector.health import HealthMonitor
from outlook_connector.poller import Poller
from outlook_connector.service import run_workflow

UTC = dt.UTC
T0 = dt.datetime(2026, 1, 1, 12, 0, 0, tzinfo=UTC)
POLL_INTERVAL = 10.0
INBOX, PROCESSED = "inbox", "processed"
MAILBOX = "sim-inbox@example.com"

VARIANTS = ("html", "text", "none_body", "unknown_ct", "no_sender", "no_received")
ATT_SPECS = ("normal", "oversize", "item", "no_ct")


class Finding(AssertionError):
    """An invariant violation attributable to the shipped code (not the sim)."""


class SimClock:
    def __init__(self) -> None:
        self.t = T0

    def now(self) -> dt.datetime:
        return self.t

    def advance(self, seconds: float) -> None:
        self.t += dt.timedelta(seconds=seconds)


@dataclass
class SimMessage:
    mid: str  # Graph immutable id — survives moves, dies only on delete
    imid: str  # RFC 822 id — the consumer dedup key
    received_stored: dt.datetime | None  # microsecond precision (Graph's stored value)
    visible_at: dt.datetime  # listings include it only from here (replication lag)
    folder: str | None = INBOX  # None = hard-deleted
    variant: str = "html"
    att_spec: str | None = None
    transient_404: bool = False  # next get_email 404s once (replication)

    @property
    def received_listed(self) -> dt.datetime | None:
        """What listings/filters see: truncated to whole seconds, like Graph."""
        if self.received_stored is None:
            return None
        return self.received_stored.replace(microsecond=0)


class SimGraph:
    """Duck-typed stand-in for outlook_helper.OutlookClient."""

    def __init__(self, seed: int, clock: SimClock, *, max_attachment_bytes: int = 100):
        self.seed = seed
        self.clock = clock
        self.max_attachment_bytes = max_attachment_bytes
        # Slow-Graph mode: each listed stub / full fetch burns wall time, so a
        # big cycle outlives the health staleness window. Only the poller's
        # heartbeats can keep the probe green — which is exactly the claim the
        # fetch-phase heartbeat fix makes.
        self.slow_s = 0.0
        self.probe = None  # called after every simulated delay (mid-cycle /health)
        self.msgs: dict[str, SimMessage] = {}
        self.serial = 0
        self.cycle_no = 0
        self.faults: list[dict] = []
        self.pending_ops: list[dict] = []
        self.ops_ran_this_cycle = False
        self.listing_completed = False
        self.reset_counters()

    # ---- bookkeeping the oracle reads -------------------------------------
    def reset_counters(self) -> None:
        self.search_calls = 0
        self.list_pages = 0
        self.listed_ids: list[str] = []
        self.get_email_calls: list[str] = []
        self.get_attachments_calls: list[str] = []
        self.ops_ran_this_cycle = False
        self.listing_completed = False

    def by_imid(self, imid: str) -> SimMessage:
        for m in self.msgs.values():
            if m.imid == imid:
                return m
        raise KeyError(imid)

    def inbox_visible_mids(self) -> set[str]:
        now = self.clock.now()
        return {
            m.mid
            for m in self.msgs.values()
            if m.folder == INBOX and m.visible_at <= now
        }

    # ---- mutation API (driver/rules) --------------------------------------
    def deliver(
        self,
        count: int = 1,
        *,
        same_second: bool = False,
        backdate_s: float = 0.0,
        lag_cycles: int = 0,
        variant: str = "html",
        att_spec: str | None = None,
    ) -> list[str]:
        imids = []
        base = (self.clock.now() - dt.timedelta(seconds=backdate_s)).replace(microsecond=0)
        for i in range(count):
            self.serial += 1
            n = self.serial
            if variant == "no_received":
                stored = None
            else:
                stored = base if same_second else base + dt.timedelta(seconds=i)
                # distinct sub-second stored values inside a same-second cohort
                stored = stored.replace(microsecond=n % 1_000_000)
            m = SimMessage(
                mid=f"g-{n}",
                imid=f"<imid-{n}@sim>",
                received_stored=stored,
                visible_at=self.clock.now() + dt.timedelta(seconds=POLL_INTERVAL * lag_cycles),
                variant=variant,
                att_spec=att_spec,
            )
            self.msgs[m.mid] = m
            imids.append(m.imid)
        return imids

    def move_out(self, imid: str, *, when="between", transient_404: bool = False) -> None:
        if when == "between":
            self._apply_move(imid, PROCESSED, transient_404)
        else:
            self.pending_ops.append(
                {"kind": "move_out", "imid": imid, "when": when, "t404": transient_404}
            )

    def move_back(self, imid: str) -> None:
        m = self.by_imid(imid)
        if m.folder == PROCESSED:
            m.folder = INBOX

    def delete(self, imid: str, *, when="between") -> None:
        if when == "between":
            self._apply_delete(imid)
        else:
            self.pending_ops.append({"kind": "delete", "imid": imid, "when": when})

    def _apply_move(self, imid: str, folder: str, t404: bool) -> None:
        try:
            m = self.by_imid(imid)
        except KeyError:
            return
        if m.folder == INBOX:
            m.folder = folder
            m.transient_404 = t404

    def _apply_delete(self, imid: str) -> None:
        try:
            self.by_imid(imid).folder = None
        except KeyError:
            pass

    def _run_pending(self, when) -> None:
        remaining = []
        for op in self.pending_ops:
            if op["when"] == when:
                self.ops_ran_this_cycle = True
                if op["kind"] == "move_out":
                    self._apply_move(op["imid"], PROCESSED, op["t404"])
                else:
                    self._apply_delete(op["imid"])
            else:
                remaining.append(op)
        self.pending_ops = remaining

    # ---- fault injection ---------------------------------------------------
    def inject(self, site: str, *, status, mid: str | None = None, sticky: bool = False,
               after_items: int | None = None) -> None:
        self.faults.append(
            {"site": site, "status": status, "mid": mid, "sticky": sticky,
             "after_items": after_items}
        )

    def clear_faults(self) -> None:
        self.faults.clear()

    def _pop_fault(self, site: str, mid: str | None = None) -> dict | None:
        for fault in self.faults:
            if fault["site"] != site:
                continue
            if fault["mid"] is not None and fault["mid"] != mid:
                continue
            if not fault["sticky"]:
                self.faults.remove(fault)
            return fault
        return None

    def _burn(self) -> None:
        if self.slow_s:
            self.clock.advance(self.slow_s)
        if self.probe is not None:
            self.probe()

    @staticmethod
    def _exc(fault: dict) -> Exception:
        if fault["status"] == "net":
            return httpx.ConnectError("sim network down")
        return GraphError(status_code=fault["status"], message=f"sim {fault['status']}")

    # ---- the client surface the poller calls -------------------------------
    def search_email(self, *, folder=None, since=None, oldest_first=False,
                     ids_only=False, **_kw):
        self.search_calls += 1
        fault = self._pop_fault("search")
        seed, cycle = self.seed, self.cycle_no

        def tiebreak(mid: str) -> str:
            return hashlib.sha1(f"{seed}:{cycle}:{mid}".encode()).hexdigest()

        def gen():
            if fault is not None and fault["after_items"] is None:
                raise self._exc(fault)
            now = self.clock.now()
            eligible = [
                m
                for m in self.msgs.values()
                if m.folder == INBOX
                and m.visible_at <= now
                and (
                    since is None
                    or (m.received_listed is not None and m.received_listed >= since)
                )
            ]
            # nulls-first like the poller's own sort; same-second ties reshuffle
            # between cycles (adversarial Graph, deterministic per seed+cycle)
            eligible.sort(
                key=lambda m: (m.received_listed is not None, m.received_listed or T0,
                               tiebreak(m.mid))
            )
            for i, m in enumerate(eligible):
                if fault is not None and fault["after_items"] == i:
                    raise self._exc(fault)
                self.listed_ids.append(m.mid)
                self._burn()
                yield OutlookMessage(id=m.mid, received_at=m.received_listed)
            self.list_pages += math.ceil(len(eligible) / 100) if eligible else 0
            self.listing_completed = True
            self._run_pending("post_listing")

        return gen()

    def get_email(self, mid: str, **_kw) -> OutlookMessage:
        fault = self._pop_fault("get_email", mid=mid)
        if fault is not None:
            raise self._exc(fault)
        self._run_pending(("post_fetch", len(self.get_email_calls)))
        m = self.msgs.get(mid)
        if m is None or m.folder is None:
            raise GraphError(status_code=404, message="ErrorItemNotFound")
        if m.transient_404:
            m.transient_404 = False
            raise GraphError(status_code=404, message="sim replication 404")
        self.get_email_calls.append(mid)
        self._burn()
        return self._full(m)

    def get_attachments(self, mid: str) -> list[OutlookAttachment]:
        fault = self._pop_fault("get_attachments", mid=mid)
        if fault is not None:
            raise self._exc(fault)
        m = self.msgs.get(mid)
        if m is None or m.folder is None:
            raise GraphError(status_code=404, message="ErrorItemNotFound")
        self.get_attachments_calls.append(mid)
        return self._attachments(m)

    # ---- realization --------------------------------------------------------
    def _full(self, m: SimMessage) -> OutlookMessage:
        body = {
            "html": OutlookBody(content_type="html", content=f"<p>{m.mid}</p>"),
            "text": OutlookBody(content_type="text", content=m.mid),
            "none_body": None,
            "unknown_ct": OutlookBody(content_type="weird", content=m.mid),
        }.get(m.variant, OutlookBody(content_type="html", content=f"<p>{m.mid}</p>"))
        return OutlookMessage(
            id=m.mid,
            internet_message_id=m.imid,
            subject=f"subject {m.mid}",
            from_=None if m.variant == "no_sender" else EmailAddress(address="sender@sim"),
            to=[EmailAddress(address=MAILBOX)],
            received_at=m.received_stored,
            body=body,
            has_attachments=m.att_spec is not None,
        )

    def _attachments(self, m: SimMessage) -> list[OutlookAttachment]:
        if m.att_spec is None:
            return []
        cap = self.max_attachment_bytes
        spec = m.att_spec
        if spec == "oversize":
            content = b"x" * (cap + 1)
        elif spec == "item":
            content = None
        else:
            content = b"%PDF"
        return [
            OutlookAttachment(
                id=f"att-{m.mid}",
                name=f"{m.mid}.pdf",
                content_type=None if spec == "no_ct" else "application/pdf",
                size=len(content) if content else 0,
                content=content,
            )
        ]


class SimConsumer:
    """FakeChannel + hlb-style dedup on internet_message_id.

    Also acts as the payload oracle (I8): every published event is checked
    field-by-field against the sim's ground truth, so a mapping regression
    (dropped sender, wrong body slot, lost attachment metadata) fails loudly
    instead of riding along as a well-formed but wrong event.
    """

    def __init__(self, sim: "SimGraph | None" = None) -> None:
        self.sim = sim
        self.first_seen: dict[str, int] = {}
        self.duplicates: Counter[str] = Counter()
        self.publishes: list[str] = []
        self.poison: set[str] = set()
        self.fail_next = 0
        self.bus_failures = 0

    def _verify(self, event) -> None:
        if self.sim is None:
            return
        e = event.data.email
        m = self.sim.msgs.get(e.id)
        if m is None:
            raise Finding(f"I8: published unknown message id {e.id}")
        if event.data.source_mailbox != MAILBOX:
            raise Finding(f"I8: wrong source_mailbox {event.data.source_mailbox}")
        if e.internet_message_id != m.imid:
            raise Finding(f"I8: {e.id} imid {e.internet_message_id} != {m.imid}")
        if e.subject != f"subject {m.mid}":
            raise Finding(f"I8: {e.id} subject {e.subject!r}")
        if e.received_at != m.received_stored:
            raise Finding(f"I8: {e.id} received_at {e.received_at} != {m.received_stored}")

        want_sender = [] if m.variant == "no_sender" else ["sender@sim"]
        if e.from_addresses != want_sender:
            raise Finding(f"I8: {e.id} from_addresses {e.from_addresses} != {want_sender}")
        if e.to_addresses != [MAILBOX]:
            raise Finding(f"I8: {e.id} to_addresses {e.to_addresses}")

        if m.variant == "text":
            want = (None, m.mid)
        elif m.variant in ("none_body", "unknown_ct"):
            want = (None, None)  # unmappable content types drop the body
        else:
            want = (f"<p>{m.mid}</p>", None)
        if (e.body_html, e.body_text) != want:
            raise Finding(
                f"I8: {e.id} variant={m.variant} body=({e.body_html!r}, {e.body_text!r}) "
                f"want {want}"
            )

        if e.has_attachments != (m.att_spec is not None):
            raise Finding(f"I8: {e.id} has_attachments={e.has_attachments} spec={m.att_spec}")
        if m.att_spec is None:
            if e.attachments:
                raise Finding(f"I8: {e.id} carries attachments with no spec")
            return
        if len(e.attachments) != 1:
            raise Finding(f"I8: {e.id} expected 1 attachment, got {len(e.attachments)}")
        att = e.attachments[0]
        if att.file_name != f"{m.mid}.pdf":
            raise Finding(f"I8: {e.id} attachment name {att.file_name!r}")
        want_ct = None if m.att_spec == "no_ct" else "application/pdf"
        if att.content_type != want_ct:
            raise Finding(f"I8: {e.id} attachment content_type {att.content_type!r}")
        cap = self.sim.max_attachment_bytes
        if m.att_spec == "oversize":
            if att.body is not None:
                raise Finding(f"I8: {e.id} oversize attachment kept {len(att.body)}B on the bus")
            if att.size != cap + 1:
                raise Finding(f"I8: {e.id} oversize lost its size metadata ({att.size})")
        elif m.att_spec == "item":
            if att.body is not None:
                raise Finding(f"I8: {e.id} item attachment invented content")
        else:
            if att.body != b"%PDF":
                raise Finding(f"I8: {e.id} attachment body {att.body!r}")

    async def publish(self, event) -> None:
        imid = event.data.email.internet_message_id
        if self.fail_next > 0:
            self.fail_next -= 1
            self.bus_failures += 1
            raise RuntimeError("bus down (transient)")
        if imid in self.poison:
            self.bus_failures += 1
            raise RuntimeError("bus down (poison)")
        self._verify(event)
        self.publishes.append(imid)
        if imid in self.first_seen:
            self.duplicates[imid] += 1
        else:
            self.first_seen[imid] = len(self.publishes)


class Driver:
    """Owns clock, sim, consumer, monitor, poller and the event loop."""

    def __init__(self, seed: int, batch: int | None, *, cap: int = 100):
        self.clock = SimClock()
        self.sim = SimGraph(seed, self.clock, max_attachment_bytes=cap)
        self.consumer = SimConsumer(self.sim)
        self.monitor = HealthMonitor(poll_interval_seconds=POLL_INTERVAL, now=self.clock.now)
        self.batch = batch
        self.cap = cap
        self.loop = asyncio.new_event_loop()
        self.log: list[str] = []
        self.credits: Counter[str] = Counter()  # I5 duplicate budget
        self.starve_streak = 0
        self.health_probes = 0
        self.sim.probe = self._probe_health
        self._new_poller()

    def _probe_health(self) -> None:
        """What an orchestrator's liveness probe sees mid-cycle."""
        self.health_probes += 1
        status = self.monitor.snapshot().status
        if status == "stale":
            raise Finding(
                "I6: /health went stale DURING a cycle that was making progress "
                f"(probe #{self.health_probes}) — a liveness probe would kill a "
                "healthy connector mid-drain"
            )

    def _new_poller(self) -> None:
        self.poller = Poller(
            client=self.sim,
            now=self.clock.now,
            source_folder=INBOX,
            batch_max_messages=self.batch,
            max_attachment_bytes=self.cap,
            heartbeat=self.monitor.beat,
        )
        self.ctx = {
            "poller": self.poller,
            "channel": self.consumer,
            "source_mailbox": MAILBOX,
        }

    def restart(self) -> None:
        for mid in self.poller._published_ids:  # noqa: SLF001 (white-box oracle)
            m = self.sim.msgs.get(mid)
            if m is not None:
                self.credits[m.imid] += 1
        self._new_poller()
        self.log.append("restart")

    def grant_moveback_credit(self, imid: str) -> None:
        if imid in self.consumer.first_seen:
            self.credits[imid] += 1

    def grant_moverace_credit(self, imid: str) -> None:
        self.credits[imid] += 1

    def run_cycle(self, *, check: bool = True):
        from invariants import check_cycle, snapshot_expectation

        self.clock.advance(POLL_INTERVAL)
        self.sim.cycle_no += 1
        pre = snapshot_expectation(self)
        self.sim.reset_counters()
        summary = self.loop.run_until_complete(run_workflow(self.ctx))
        self.monitor.record_cycle(summary)
        self.log.append(
            f"cycle#{self.sim.cycle_no}: clean={pre.clean} fetched={summary.fetched} "
            f"published={summary.published} dropped={summary.dropped} "
            f"err={summary.error_class} src={summary.error_source}"
        )
        if check:
            check_cycle(self, pre, summary)
        return summary

    def close(self) -> None:
        self.loop.close()
