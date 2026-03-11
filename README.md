# mitchell-press-emails-database

PowerShell script that exports every email from an M365 mailbox (Inbox + Sent Items) into a SQLite database, with an interactive menu for full export, incremental sync, and standalone attachment downloading.

## Prerequisites

- **PowerShell 7+** (or Windows PowerShell 5.1)
- The following modules are auto-installed on first run:
  - `Microsoft.Graph.Authentication`
  - `PSSQLite`

## Azure AD App Registration

1. Go to [Azure Portal > App Registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) and create a new registration.
2. Note the **Application (client) ID** and **Directory (tenant) ID**.
3. Choose your auth method:

| Method | When to use | Permissions needed |
|---|---|---|
| **Interactive (delegated)** | Running as yourself | `Mail.Read` (delegated) |
| **App-only (client secret)** | Unattended / service account | `Mail.Read` (application) + admin consent |

For **app-only** auth, also create a client secret under **Certificates & secrets**.

## Usage

Launch the script — it authenticates, then shows an interactive menu:

```
============================================
   M365 Mailbox Email Exporter
============================================

  [1] Full Export
      Wipe database and download ALL emails

  [2] Incremental Sync
      Download only new/changed emails since last run

  [3] Download Missing Attachments
      Scan DB and download attachments not yet saved to disk

  [Q] Quit
```

### Interactive login (delegated)

```powershell
.\Export-MailboxToSQLite.ps1 `
    -DatabasePath ".\emails.db" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id"
```

### App-only (client credentials)

```powershell
.\Export-MailboxToSQLite.ps1 `
    -DatabasePath ".\emails.db" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-client-secret" `
    -UserEmail "user@yourdomain.com"
```

### Custom attachment folder

```powershell
.\Export-MailboxToSQLite.ps1 `
    -DatabasePath ".\emails.db" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -AttachmentPath "D:\EmailAttachments"
```

## Menu Options

### 1 — Full Export
Deletes all existing data and re-downloads every email from Inbox and Sent Items. Use this for first-time setup or a clean reset.

### 2 — Incremental Sync
Uses the `lastModifiedDateTime` from the previous sync to only fetch emails that are new or changed. Fast for daily/scheduled runs.

### 3 — Download Missing Attachments
A **separate, standalone operation** that:
1. Scans the database for emails where `has_attachments = 1 AND attachments_downloaded = 0`
2. Downloads each attachment from Graph API to disk
3. Records the file path in the `attachments` table
4. Marks the email's `attachments_downloaded` flag to `1`

This means you can sync emails first (option 1 or 2), then come back later and download attachments at your own pace. Re-running option 3 will only fetch what's still missing.

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

# Sync history
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT * FROM sync_log ORDER BY completed_at DESC"
```

Or use any SQLite tool (DB Browser for SQLite, `sqlite3` CLI, DBeaver, etc.).
