"""Brute-force scenario fuzzer: seeded random event streams, invariants as oracle.

The committed budget is deliberately small and the seeds are fixed, so this is
a fast, deterministic gate that can never go randomly red. One test per seed,
so it also parallelises if you have xdist.

Replay a single failing seed (the failure prints the seed and full trace):

    uv run pytest "tests/sim/test_fuzz.py::test_fuzz_seed[7]" -q -s

Hunt harder on demand — hundreds of fresh scenarios across every core:

    SIM_SEEDS=800 SIM_STEPS=400 SIM_SEED_BASE=1000 \
        uv run --with pytest-xdist pytest tests/sim/test_fuzz.py -n 12 -q

`SIM_ALLOW_POISON=1` additionally arms sticky bus poison; expect it to trip the
starvation invariant (see the strict-xfail in test_scenarios.py) until the
stop-batch policy grows a skip.
"""

from __future__ import annotations

import os
import random

import pytest
from invariants import drain_and_assert
from sim_graph import ATT_SPECS, INBOX, PROCESSED, VARIANTS, Driver, Finding

SEED_BASE = int(os.environ.get("SIM_SEED_BASE", "0"))
SEEDS = [SEED_BASE + i for i in range(int(os.environ.get("SIM_SEEDS", "16")))]
STEPS = int(os.environ.get("SIM_STEPS", "150"))
ALLOW_POISON = os.environ.get("SIM_ALLOW_POISON") == "1"

STEP_WEIGHTS = {
    "cycle": 0.45,
    "deliver": 0.25,
    "move_out": 0.08,
    "graph_fault": 0.06,
    "move_back": 0.04,
    "delete": 0.04,
    "restart": 0.04,
    "bus_outage": 0.04,
    "slow_toggle": 0.02,
}
MAX_LIVE = 80
FAULT_STATUSES = [404, 429, 503, "net"]
FAULT_SITES = ["search", "get_email", "get_attachments"]
WHENS = ["between", "post_listing", ("post_fetch", 1), ("post_fetch", 2)]


def _pick(rng: random.Random) -> str:
    kinds, weights = zip(*STEP_WEIGHTS.items())
    return rng.choices(kinds, weights=weights, k=1)[0]


def _inbox_imids(driver: Driver) -> list[str]:
    return sorted(m.imid for m in driver.sim.msgs.values() if m.folder == INBOX)


def _processed_imids(driver: Driver) -> list[str]:
    return sorted(m.imid for m in driver.sim.msgs.values() if m.folder == PROCESSED)


def run_scenario(seed: int, steps: int = STEPS) -> Driver:
    rng = random.Random(seed)
    driver = Driver(seed, rng.choice([3, 5, 10, None]))
    for _ in range(steps):
        kind = _pick(rng)

        if kind == "cycle":
            driver.run_cycle()

        elif kind == "deliver":
            live = len([m for m in driver.sim.msgs.values() if m.folder == INBOX])
            if live >= MAX_LIVE:
                continue
            driver.sim.deliver(
                rng.randint(1, 5),
                same_second=rng.random() < 0.4,
                backdate_s=rng.choice([0.0, 1.5, 90.0]),
                lag_cycles=rng.choice([0, 0, 1, 2]),
                variant=rng.choice(VARIANTS),
                att_spec=rng.choice([None, None, *ATT_SPECS]),
            )
            driver.log.append("deliver")

        elif kind == "move_out":
            candidates = _inbox_imids(driver)
            if not candidates:
                continue
            imid = rng.choice(candidates)
            when = rng.choice(WHENS)
            t404 = rng.random() < 0.5
            driver.sim.move_out(imid, when=when, transient_404=t404)
            if when != "between":
                driver.grant_moverace_credit(imid)
            driver.log.append(f"move_out {imid} when={when} t404={t404}")

        elif kind == "move_back":
            candidates = _processed_imids(driver)
            if not candidates:
                continue
            imid = rng.choice(candidates)
            driver.sim.move_back(imid)
            driver.grant_moveback_credit(imid)
            driver.log.append(f"move_back {imid}")

        elif kind == "delete":
            candidates = _inbox_imids(driver) + _processed_imids(driver)
            if not candidates:
                continue
            imid = rng.choice(candidates)
            driver.sim.delete(imid, when=rng.choice(WHENS))
            driver.log.append(f"delete {imid}")

        elif kind == "restart":
            driver.restart()

        elif kind == "graph_fault":
            if len(driver.sim.faults) >= 2:
                continue
            site = rng.choice(FAULT_SITES)
            status = rng.choice(FAULT_STATUSES)
            # sometimes fail mid-pagination: the poller has already collected
            # part of the listing when the error hits
            after = rng.choice([None, None, 0, 1, 3]) if site == "search" else None
            driver.sim.inject(site, status=status, sticky=False, after_items=after)
            driver.log.append(f"fault {site} {status} after={after}")

        elif kind == "slow_toggle":
            # slow Graph: burns simulated wall time per listed/fetched message,
            # so cycles outlive the health staleness window unless the poller
            # heartbeats through both phases
            driver.sim.slow_s = rng.choice([0.0, 0.0, 2.0, 5.0])
            driver.log.append(f"slow_s={driver.sim.slow_s}")

        elif kind == "bus_outage":
            if ALLOW_POISON and rng.random() < 0.15:
                candidates = _inbox_imids(driver)
                if candidates:
                    driver.consumer.poison.add(rng.choice(candidates))
                    driver.log.append("poison")
            else:
                driver.consumer.fail_next += rng.randint(1, 2)
                driver.log.append("bus_outage")

    drain_and_assert(driver)
    return driver


@pytest.mark.fuzz
@pytest.mark.parametrize("seed", SEEDS)
def test_fuzz_seed(seed: int):
    driver = None
    try:
        driver = run_scenario(seed)
    except Finding as exc:
        trace = "\n".join(driver.log) if driver else "<no trace>"
        pytest.fail(f"SEED={seed}\n{exc}\n=== full trace ===\n{trace}", pytrace=False)
    except AssertionError as exc:
        trace = "\n".join(driver.log) if driver else "<no trace>"
        pytest.fail(f"SEED={seed} (assertion)\n{exc}\n=== full trace ===\n{trace}", pytrace=False)
    finally:
        if driver is not None:
            driver.close()
