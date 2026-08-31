# outlook-connector — Implementation Plan

## 1. Components

| Component | Type | Function |
|---|---|---|
| `outlook-helper` | Library | Graph API access. Read, move, delete, send. |
| `outlook-connector` | Service | Ingest loop, extraction, database write, command consumer. |

The library has no database dependency and no bus dependency. A caller can use
it in its own process. The service adds the loop, the storage, and the bus.

## 2. Key decisions

### 2.1 Extraction in one place

The connector parses the MIME one time. All readers get the same result.

**Reason:** A parser has many decisions that are not visible in the data: the
treatment of inline images, of nested messages, and of TNEF. Many parsers give many answers.

Because there is one writer, a shared parser library is not necessary. Readers
copy the schema.

### 2.2 The connector owns the schema

The connector holds the DDL and the migrations. Alembic does the migrations.

Readers use the tables, but must keep them separate from their own migrations.

Put the semantic rules in a `CONTRACT.md` file next to the DDL. The column
names do not show these rules. A new reader must not read the connector source
to learn them. Give an answer for each of these:

- Are inline images rows in the attachment table?
- Is `body_txt` the text part, or a downgrade of the HTML part?
- Does a forwarded message give one row or the rows of its children?
- Is TNEF unpacked before the write?
- Are the headers raw or decoded? Are duplicates kept?

### 2.3 Deterministic key

The key is `(mailbox, internet_message_id)`, with a unique constraint.

**Reason:** The delta cursor and the write are not in one transaction. After a
crash, the connector writes some emails again. A deterministic key makes the
second write an overwrite, not a duplicate. No de-duplication logic is
necessary.

### 2.4 No lifecycle in the connector

The connector never deletes. The mailbox is the source of truth, so a lost row
is always recoverable.

**Reason:** The connector cannot know the retention rules of its readers. A TTL
in the connector is a guess. A delete signal from a reader is a second protocol
and a new failure mode.

### 2.5 Bus for commands and notification

Redis Streams, with consumer groups. Not Redis pub/sub, because pub/sub has no
acknowledgement and no replay.

- One command stream for each mailbox. This keeps the command order correct.
- Commands are fire-and-forget. They are idempotent, so a repeat is safe.
- The notification event contains only the key. Do not put fields in the event.
  A copy of a field in the event becomes stale, and then two sources disagree.

### 2.6 Storage

`BYTEA` columns. PostgreSQL TOASTs large values automatically. A separate blob
table is not necessary.

- `raw` keeps the default `EXTENDED` storage. MIME is base64 text and it
  compresses well.
- Attachment bytes use `SET STORAGE EXTERNAL`. The data is already compressed,
  so a compression attempt is a waste.
- Never use `SELECT *`. This pulls the blob into every metadata query.
- `bytea` has a 1 GB limit for each value and no streaming. A 30 MB attachment
  uses 30 MB of server memory and 30 MB of client memory. Keep the batch size
  small.
- Each insert writes the full bytes to the WAL. The write IO is near two times
  the data size. Replication carries the same load.
- psycopg2 gives a `memoryview`. Cast to `bytes` at the module edge.

## 3. Schema

```
message
  internet_message_id   text     -- part of the key
  mailbox               text     -- part of the key
  provider_msg_id       text     -- Graph ID
  raw                   bytea    -- full original MIME
  subject               text
  body_txt              text
  body_html             text
  headers               jsonb
  received_at           timestamptz
  fetched_at            timestamptz
  parser_version        integer
  UNIQUE (mailbox, internet_message_id)

attachment
  message_id            fk -> message
  filename              text
  content_type          text
  is_inline             boolean
  data                  bytea    -- STORAGE EXTERNAL
```

`parser_version` is necessary. The parse rules will change one time, most
probably for the inline images. The column identifies the old rows. Because
`raw` stays, a re-parse is always possible. The column is cheap now and
expensive later.

## 4. Sequence of operations

### 4.1 Ingest loop

1. Get the delta page from Graph.
2. For each message: get the MIME.
3. Parse the MIME. Unpack TNEF if it is present.
4. Upsert the message row and the attachment rows in one transaction.
5. Publish the key on the notification stream.
6. Save the delta cursor.

Step 4 and step 5 are a dual write. The bus is not in the database
transaction. Publish after the commit. If a reader gets an event and finds no
row, it must retry and not acknowledge.

### 4.2 Command loop

1. Read from the command stream of the mailbox.
2. Apply the command through `outlook-helper`.
3. Acknowledge. Write a log record if there is a failure.

## 5. Parser risks

The standard library covers most of the work:

```python
msg = email.message_from_bytes(raw, policy=policy.default)
for part in msg.iter_attachments():
    ...
```

Note: `policy.default` is necessary. The legacy policy does not decode the
file names correctly.

These conditions need more code:

- **TNEF.** Detect `application/ms-tnef`, `application/vnd.ms-tnef`, the name
  `winmail.dat`, or the signature `0x223E9F78`. Unpack with `tnefparse`.
  Without this, one opaque blob replaces all the true attachments, and there is
  no error.
- **Inline parts.** Signature images and logos are attachments in the MIME.
  Mark them with `is_inline`. Do not discard them, because the decision belongs
  to the reader.
- **Nested `message/rfc822`.** A forwarded email holds its attachments one
  level down. Make the decision one time and record it in `CONTRACT.md`.
- **Reference attachments.** Graph MIME does not contain the bytes of OneDrive
  or SharePoint attachments. If the mailbox has these, the MIME is not the full
  picture.

## 6. Test data

Collect a corpus of `.eml` files. This is the primary test asset. The parser is
a pure function, so a directory of files gives full test coverage with no
Graph access.

## 7. Packaging

- Python package, with a `main()` and a Dockerfile.
- Version tags. Readers pin the version.
- One instance for each mailbox. Each instance has its own configuration.
- No shared instance, and no multi-tenancy.

## 8. Build sequence

1. `outlook-helper`. Graph read, move, delete, send. Pluggable cursor.
2. Parser, with the `.eml` corpus. No database and no bus.
3. Schema, Alembic migrations, and `CONTRACT.md`.
4. Ingest loop, with the notification publish.
5. Command consumer.

Step 1 and step 2 are independent. Both can start immediately.

## 9. Rejected options

| Option | Reason for rejection |
|---|---|
| Key/value schema, readers parse | Many parsers give different results. There is one team and one standard, so one extraction point is better. |
| S3 for the MIME | The requirement says PostgreSQL. `bytea` with TOAST is sufficient at this volume. |
| Redis pub/sub | No acknowledgement, and no replay. |
| Fields in the bus event | The copy becomes stale. The database is one join away. |
| TTL or cleanup in the connector | The connector cannot know the retention rules of its readers. |
| Acknowledgement channel for a delete | A second protocol, and a new failure mode. |
| Attachment parse in each reader | The decisions are silent, and the results diverge. |
