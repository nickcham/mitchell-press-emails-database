# mitchell-press-emails-database

PowerShell script that exports every email from an M365 mailbox (Inbox + Sent Items) into a SQLite database, with an interactive menu for full export, incremental sync, and standalone attachment downloading.

## Prerequisites

- **PowerShell 7+** (or Windows PowerShell 5.1)
- The following modules are auto-installed on first run:
  - `Microsoft.Graph.Authentication`
  - `PSSQLite`

## Usage

### Quick start — interactive browser login (with MFA)

```powershell
.\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db"
```

That's it. A browser window opens, you sign in with your Microsoft 365 account, complete MFA if prompted, and the script runs against your mailbox. No app registration needed.

The interactive menu then appears:

```
============================================
   M365 Mailbox Email Exporter
============================================

  [1] Full Export
      Download ALL emails (first run or full re-scan)

  [2] Incremental Sync
      Download only new/changed emails since last run

  [3] Download Missing Attachments
      Scan DB and download attachments not yet saved to disk

  [4] Build Conversations (Full)
      Rebuild conversations table from all emails in DB

  [5] Build Conversations (Incremental)
      Update only conversations with new/changed emails

  [Q] Quit
```

### Custom attachment folder

```powershell
.\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db" -AttachmentPath "D:\EmailAttachments"
```

### App-only auth (optional — for unattended/service scenarios)

If you need to run this without a user present, set up an [Azure AD App Registration](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) with `Mail.Read` (application) permission + admin consent, then:

```powershell
.\Export-MailboxToSQLite.ps1 `
    -DatabasePath ".\emails.db" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-client-secret" `
    -UserEmail "user@yourdomain.com"
```

## Menu Options

### 1 — Full Export
Downloads every email from Inbox and Sent Items and upserts into the database. Use this for the first run or a complete re-scan. Existing data is updated, not deleted — your `attachments_downloaded` flags and attachment files are preserved.

### 2 — Incremental Sync
Uses the `lastModifiedDateTime` from the previous sync to only fetch emails that are new or changed. Fast for daily/scheduled runs.

### 3 — Download Missing Attachments
A **separate, standalone operation** that:
1. Scans the database for emails where `has_attachments = 1 AND attachments_downloaded = 0`
2. Downloads each attachment from Graph API to disk
3. Records the file path in the `attachments` table
4. Marks the email's `attachments_downloaded` flag to `1`

This means you can sync emails first (option 1 or 2), then come back later and download attachments at your own pace. Re-running option 3 will only fetch what's still missing.

### 4 — Build Conversations (Full)
Reads every email in the database and groups them by `conversation_id` into the `conversations` table. Each row contains the full chronological thread as clean text — this is the table AI uses for KB lookups. No Graph API calls needed; it works entirely from local data.

### 5 — Build Conversations (Incremental)
Only rebuilds conversations where an email's `last_modified` timestamp is newer than the conversation's `last_built` timestamp. Fast for keeping the conversations table current after an incremental email sync.

## Database Schema

### `emails` table

| Column | Type | Description |
|---|---|---|
| `message_id` | TEXT (PK) | Graph API message ID |
| `conversation_id` | TEXT | Conversation thread ID |
| `subject` | TEXT | Email subject line |
| `from_name` | TEXT | Sender display name |
| `from_address` | TEXT | Sender email address |
| `to_recipients` | TEXT | To recipients (semicolon-separated) |
| `cc_recipients` | TEXT | CC recipients |
| `bcc_recipients` | TEXT | BCC recipients |
| `reply_to` | TEXT | Reply-to addresses |
| `sent_datetime` | TEXT | When the email was sent (ISO 8601) |
| `received_datetime` | TEXT | When the email was received (ISO 8601) |
| `has_attachments` | INTEGER | 1 if email has attachments |
| `attachments_downloaded` | INTEGER | 1 if attachments have been saved to disk |
| `importance` | TEXT | low / normal / high |
| `is_read` | INTEGER | 1 if read |
| `is_draft` | INTEGER | 1 if draft |
| `body_content_type` | TEXT | html or text |
| `body_content` | TEXT | Full email body |
| `body_preview` | TEXT | Short body preview |
| `web_link` | TEXT | Outlook web link to the email |
| `folder` | TEXT | Source folder (Inbox / SentItems) |
| `categories` | TEXT | Outlook categories |
| `internet_message_id` | TEXT | RFC 2822 Message-ID header |
| `parent_folder_id` | TEXT | Graph folder ID |
| `created_datetime` | TEXT | When created in mailbox |
| `last_modified` | TEXT | Last modified timestamp |

### `attachments` table

| Column | Type | Description |
|---|---|---|
| `id` | TEXT (PK) | Graph attachment ID |
| `message_id` | TEXT (FK) | References `emails.message_id` |
| `filename` | TEXT | Original filename |
| `content_type` | TEXT | MIME type (e.g. application/pdf) |
| `size_bytes` | INTEGER | File size in bytes |
| `disk_path` | TEXT | Full path to saved file on disk |

### `conversations` table (AI / KB)

| Column | Type | Description |
|---|---|---|
| `conversation_id` | TEXT (PK) | Graph conversation thread ID |
| `subject` | TEXT | Subject from the first email in the thread |
| `participants` | TEXT | All unique email addresses (semicolon-separated) |
| `message_count` | INTEGER | Number of emails in the conversation |
| `has_attachments` | INTEGER | 1 if any email in the thread has attachments |
| `first_message_datetime` | TEXT | Earliest email in the thread |
| `last_message_datetime` | TEXT | Most recent email in the thread |
| `full_thread` | TEXT | Complete conversation as clean text (HTML stripped, chronological) |
| `outlook_link` | TEXT | Outlook web link to the most recent email (for manual review) |
| `last_built` | TEXT | When this conversation row was last rebuilt |
| **AI First-Pass Triage** | | |
| `ai_category` | TEXT | AI-assigned category: `fact`, `how-to`, `info`, `kb`, `rubber-stamp`, `not-relevant` |
| `ai_confidence` | TEXT | AI confidence level: `low`, `medium`, `high` |
| `ai_summary` | TEXT | AI comment on what the conversation is about and why it may be worth preserving |
| `ai_review_datetime` | TEXT | When AI first-pass review was performed |
| **AI Second-Pass Confirmation** | | |
| `ai_kb_confirmed` | INTEGER | 1 = confirmed for KB/RAG, 0 = rejected, NULL = not yet reviewed |
| `ai_kb_confirm_datetime` | TEXT | When AI second-pass confirmation was performed |
| `ai_kb_confirm_notes` | TEXT | AI notes on final KB decision (what to document, how to structure) |

### `sync_log` table

| Column | Type | Description |
|---|---|---|
| `id` | INTEGER (PK) | Auto-increment |
| `sync_type` | TEXT | full / incremental |
| `folder` | TEXT | Inbox or SentItems |
| `started_at` | TEXT | Sync start time (UTC) |
| `completed_at` | TEXT | Sync end time (UTC) |
| `emails_synced` | INTEGER | Number of emails processed |

## Querying the Database

```powershell
Import-Module PSSQLite

# Recent emails
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT subject, from_address, sent_datetime FROM emails ORDER BY sent_datetime DESC LIMIT 10"

# Emails with attachments still pending download
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT subject, from_address FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0"

# Downloaded attachments and their paths
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT e.subject, a.filename, a.disk_path FROM emails e JOIN attachments a ON e.message_id = a.message_id"

# Search conversations (AI/KB table)
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT subject, participants, message_count, last_message_datetime FROM conversations ORDER BY last_message_datetime DESC LIMIT 10"

# Full thread text for a conversation
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT full_thread FROM conversations WHERE subject LIKE '%project update%'"

# Sync history
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT * FROM sync_log ORDER BY completed_at DESC"
```

Or use any SQLite tool (DB Browser for SQLite, `sqlite3` CLI, DBeaver, etc.).
