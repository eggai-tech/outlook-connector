# Development

Working on the connector itself. For running it, see [README.md](README.md).

## Repository layout

```
src/outlook_connector/   the connector service
tests/                   its test suite
libs/outlook-helper/     the M365 Graph library, vendored in-tree
development/DESIGN.md    architecture and the bus message contract
development/steps/       design notes kept from the initial build
```

This repository is a [uv workspace](https://docs.astral.sh/uv/concepts/projects/workspaces/)
with two members: the connector at the root, and `outlook-helper` under `libs/`.
They are developed, locked and tested together, and no credentials are needed to
build anything.

## Setup

```sh
just sync         # install both members, with dev dependencies
```

Requires Python 3.13+, [uv](https://docs.astral.sh/uv/) and
[`just`](https://just.systems). `just` on its own lists every recipe.

## Tests

```sh
just test         # connector suite
just test-helper  # outlook-helper suite
just test-all     # both
```

The two suites cannot share a single pytest run: each defines a
`tests/test_config.py` and neither has an `__init__.py`, so collection would
collide on module name. `just test-helper` runs pytest from
`libs/outlook-helper`, giving it its own rootdir and config.

Neither suite touches a live tenant — the helper mocks the Graph HTTP layer
with `respx`, and the connector's tests use fakes.

## Running the connector on the host

`just run` starts your working copy outside Docker against the compose
postgres/redis, which is faster to iterate on than a rebuild:

```sh
docker compose up -d redis postgres   # just the dependencies
just run                              # your working copy against them
```

Use `docker compose up -d redis postgres` rather than `just up` here: `just up`
also starts the containerised connector, and two connectors polling the same
mailboxes will fight over the same mail. If the container is already running,
`docker compose stop connector` first.

## How outlook-helper is wired in

`outlook-helper` used to be a private git dependency fetched over SSH. It is now
in-tree, but **connector code imports it exactly as it would any third-party
package**:

```python
from outlook_helper import OutlookClient, OutlookMessage
```

Never import it by path or relative to `src/`. `pyproject.toml` declares a plain
requirement and only `[tool.uv.sources]` records where it currently lives:

```toml
[project]
dependencies = ["outlook-helper>=0.4.0", ...]

[tool.uv.sources]
outlook-helper = { workspace = true }
```

Keeping those two facts separate is what makes the library's location an
implementation detail. Moving it back out to its own repository is a packaging
change with **no source changes** — delete `[tool.uv.workspace]` and `libs/`,
then point the source at git:

```toml
[tool.uv.sources]
outlook-helper = { git = "https://github.com/eggai-tech/outlook-helper.git", tag = "v0.4.0" }
```

## Releases

The repository ships two products, so tags use distinct prefixes:

| Tag | Releases | Consumed by |
|---|---|---|
| `v*` | the connector image | CI builds and pushes to ghcr.io |
| `outlook-helper-v*` | the library | other projects, via a git subdirectory install |

Nothing forces a library version bump now that it is vendored, so when
`libs/outlook-helper/` changes in a way other projects depend on, bump the
version in `libs/outlook-helper/pyproject.toml` and tag it:

```sh
git tag outlook-helper-v0.4.1 && git push origin outlook-helper-v0.4.1
```

See [libs/outlook-helper/README.md](libs/outlook-helper/README.md) for the
consumer side of that.

## CI

[.github/workflows/ci.yml](.github/workflows/ci.yml) runs both test suites on
every pull request, then builds and pushes the connector image to ghcr.io on
pushes to `main` and on `v*` tags. `main` publishes `:main` and `:sha-*`;
only a non-prerelease `v*` tag publishes `:latest`, so `:latest` always means
"the current release" rather than "the tip of the development branch".
`outlook-helper-v*` tags deliberately do not match the image trigger.
