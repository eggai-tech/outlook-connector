# outlook-connector Initial Spec

## 1. Purpose

The outlook-connector is a generic service. It reads a Microsoft 365 mailbox
through the Graph API. It writes each email to a storage backend (v1: PostgreSQL database). 
It also applies mailbox commands (move, delete, send) that it receives on a message bus.

The connector is a reusable component. 
Different projects use it without a change to its code, only with configuration in a file.

## 2. Scope

### 2.1 In scope

- Read one mailbox through the Graph API.
- Extract the full email message one time, at the point of ingest
- Write the Email object, raw MIME (aka .eml), the extracted fields, and the attachments to storage.
- Publish a notification on the bus for each new email.
- Receive and apply mailbox commands from the bus.
- Own the database schema and the schema migrations.

### 2.2 Out of scope

- Business logic of any type.
- The lifecycle of the rows after the connector writes them. Consumers are responsible for cleanup.
- An HTTP API or dashboard. The connector has no read interface.
- Multi-tenancy. One instance reads one mailbox.

## 3. Functional requirements

### Bus

- The connector communicates with other services over an EggAI SDK bus.

### Configuration

- Configuration is done in a config.yaml file.
- All deployment values come from the configuration: the mailbox, the
  credentials, the database URL, and the stream names.
- Secrets are never allowed in the config file, they must come from environment variables.

### Mailbox

- The connector interacts with the Graph API. It is the only component in a solution with 
credentials and access to the mailbox.
- The connector obeys the Graph API rate limits. 

### Ingest

- The connector writes each email one time. The key is the pair
  `(mailbox, message_id)`.
- A repeated ingest of the same email doesn't make a duplicate row.
  The write is an idempotent upsert on the key.

### Extraction

The connector is the only component that parses MIME. It gives all readers one
standard result.

- The connector stores the full original MIME bytes.
- The connector extracts an Email data object with standard fields:
    - message id
    - sender 
    - receiver
    - subject
    - body_text
    - body_html
- Message headers are not relevant.
- The connector extracts each attachment as a separate row. Each row
  has a file name, a content type, and the file bytes.

- The connector decodes RFC 2231 file names correctly.
- The connector records the parser version on each message row. This is a hard coded value in source code.
  The version permits a re-parse of old rows from the stored MIME.

### Commands

- The connector subscribes to a command stream for its mailbox.
- The connector applies these commands: move, delete, and send.
- Commands are idempotent. The key is the `internet_message_id`.
- The connector doesn't send a reply for a command. It writes a log
  record for each failure.
- Commands for one mailbox stay in order. Use one stream for each
  mailbox.

### 3.4 Notification

- The connector publishes an event after each successful write.
- The event contains the key and Email object with standard fields, no MIME, no attachments, no headers. 
Readers can get all other data from the database.

### 3.5 Lifecycle

- The connector doesn't delete rows, it only writes to the consumer's database. 
- The reader that owns the data is responsible for setting and enforcing the retention policy. The connector has no
  knowledge of this.

## 4. Non-functional requirements

- The connector runs as one process for each mailbox.
- The connector needs `Mail.ReadWrite` permission, because it applies
  commands.
