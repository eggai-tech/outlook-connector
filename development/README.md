# Implementation pieces

The [inbound MVP](../docs/implementation.md) broken into independently shippable
vertical slices, in delivery order. Each piece is testable on its own; you could
ship after Piece 3 and have a working connector.

| # | Piece | Ships | Depends on |
| - | ----- | ----- | ---------- |
| 0 | [Dependency verification](00-dependency-verification.md) | A findings note that de-risks the design | — |
| 1 | [Bus contract](01-bus-contract.md) | Owned Pydantic models + a publish script | 0 |
| 2 | [Service skeleton](02-service-skeleton.md) | A deployable service shell (config + lifecycle) | — |
| 3 | [Inbound polling core](03-inbound-polling-core.md) | The MVP — real mail flows to the bus | 0, 1, 2 |
| 4 | [Attachment metadata](04-attachment-metadata.md) | Attachment metadata on the bus | 3 |

**Ordering rationale:** 0 unblocks design, 1 unblocks downstream consumers, 2
gives a deployable shell, 3 delivers the actual MVP value, 4 is additive.
Pieces 1 and 2 can be built in parallel.
