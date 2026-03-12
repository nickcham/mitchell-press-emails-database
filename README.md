# mitchell-press-emails-database

PowerShell script that exports every email from an M365 mailbox (Inbox + Sent Items) into CSV files, with an interactive menu for full export, incremental sync, conversation threading, and standalone attachment downloading.

## Prerequisites

- **PowerShell 7+** (or Windows PowerShell 5.1)
- `Microsoft.Graph.Authentication` module (auto-installed on first run)

No other external modules required — uses built-in `Export-Csv`/`Import-Csv`.

## Usage

### Quick start — interactive browser login (with MFA)

```powershell
.\Export-MailboxToCSV.ps1
```

That's it. A browser window opens, you sign in with your Microsoft 365 account, complete MFA if prompted, and the script runs against your mailbox. No app registration needed.

The interactive menu then appears:

```
============================================
   M365 Mailbox Email Exporter (CSV)
============================================

  [1] Full Export
      Download ALL emails (overwrites existing CSVs)

  [2] Incremental Sync
      Download only new/changed since last run

  [3] Download Missing Attachments
      Scan CSVs and download missing attachments

  [4] Build Conversations (Full)
      Rebuild conversations.csv from all emails

  [5] Build Conversations (Incremental)
      Update changed conversations only

  [6] Status & Quick Notes
      Show stats, sync history, custom app guide

  [Q] Quit
```

### Custom output and attachment folders

```powershell
.\Export-MailboxToCSV.ps1 -OutputPath ".\my-export" -AttachmentPath "D:\EmailAttachments"
```

### App-only auth (optional — for unattended/service scenarios)

If you need to run this without a user present, set up an [Azure AD App Registration](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) with `Mail.Read` (application) permission + admin consent, then:

```powershell
.\Export-MailboxToCSV.ps1 `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-client-secret" `
    -UserEmail "user@yourdomain.com"
```

## Output Files

All files are written to the output folder (default: `.\data\`):

| File | Description |
|---|---|
| `emails_inbox.csv` | All Inbox emails |
| `emails_sent.csv` | All Sent Items emails |
| `conversations.csv` | Threaded conversations for AI/KB consumption |
| `sync_log.csv` | Sync history log |

CSVs open directly in Excel, PowerBI, or any data tool.

## Menu Options

### 1 — Full Export
Downloads every email from Inbox and Sent Items and writes to CSV. Use this for the first run or a complete re-scan. Overwrites existing CSVs.

### 2 — Incremental Sync
Uses the `lastModifiedDateTime` from the previous sync to only fetch emails that are new or changed. Existing rows are merged — updated emails are replaced while preserving `attachments_downloaded` flags. New emails are appended.

### 3 — Download Missing Attachments
A **separate, standalone operation** that:
1. Scans both email CSVs for rows where `has_attachments = 1` and `attachments_downloaded = 0`
2. Downloads each attachment from Graph API to disk
3. Marks the email's `attachments_downloaded` flag to `1` in the CSV

Re-running option 3 will only fetch what's still missing.

### 4 — Build Conversations (Full)
Reads every email from both CSVs and groups them by `conversation_id` into `conversations.csv`. Each row contains the full chronological thread as clean text — this is the file AI uses for KB lookups. No Graph API calls needed; it works entirely from local data.

The `full_thread` column is processed for AI/RAG quality:
- **HTML stripping** — `<style>` and `<script>` blocks are removed entirely, block-level tags (`<p>`, `<div>`, `<br>`) are converted to line breaks, all remaining tags are stripped, and HTML entities (`&amp;`, `&nbsp;`, `&quot;`, numeric entities, etc.) are decoded. Paragraph structure is preserved rather than collapsing everything to a single line.
- **Quote deduplication** — each email's body is trimmed at the point where quoted/forwarded content begins (e.g. "--- Original Message ---", "On ... wrote:", `>>>` markers, Outlook's `From:/Sent:` header blocks). This prevents the same content appearing multiple times in the thread and reduces token count for AI consumption.

### 5 — Build Conversations (Incremental)
Only rebuilds conversations where an email's `last_modified` timestamp is newer than the conversation's `last_built` timestamp. Fast for keeping conversations current after an incremental email sync.

### 6 — Status & Quick Notes
Shows email/conversation counts, last sync times, auth info, and a quick guide for setting up a custom Azure AD app registration.

## CSV Columns

### Email CSVs (`emails_inbox.csv` / `emails_sent.csv`)

| Column | Description |
|---|---|
| `message_id` | Graph API message ID (unique key) |
| `conversation_id` | Conversation thread ID |
| `subject` | Email subject line |
| `from_name` | Sender display name |
| `from_address` | Sender email address |
| `to_recipients` | To recipients (semicolon-separated) |
| `cc_recipients` | CC recipients |
| `bcc_recipients` | BCC recipients |
| `reply_to` | Reply-to addresses |
| `sent_datetime` | When the email was sent (ISO 8601) |
| `received_datetime` | When the email was received (ISO 8601) |
| `has_attachments` | 1 if email has attachments |
| `attachments_downloaded` | 1 if attachments have been saved to disk |
| `importance` | low / normal / high |
| `is_read` | 1 if read |
| `is_draft` | 1 if draft |
| `body_content_type` | html or text |
| `body_content` | Full email body (raw) |
| `body_preview` | Short body preview |
| `cleaned_body` | HTML-stripped, quote-deduplicated body text |
| `web_link` | Outlook web link to the email |
| `folder` | Source folder (Inbox / SentItems) |
| `categories` | Outlook categories |
| `internet_message_id` | RFC 2822 Message-ID header |
| `parent_folder_id` | Graph folder ID |
| `created_datetime` | When created in mailbox |
| `last_modified` | Last modified timestamp |

### Conversations CSV (`conversations.csv`)

| Column | Description |
|---|---|
| `conversation_id` | Graph conversation thread ID |
| `subject` | Subject from the first email in the thread |
| `participants` | All unique email addresses (semicolon-separated) |
| `message_count` | Number of emails in the conversation |
| `has_attachments` | 1 if any email in the thread has attachments |
| `first_message_datetime` | Earliest email in the thread |
| `last_message_datetime` | Most recent email in the thread |
| `full_thread` | Complete conversation as clean text (HTML stripped, chronological) |
| `outlook_link` | Outlook web link to the most recent email |
| `last_built` | When this conversation row was last rebuilt |

### Sync Log CSV (`sync_log.csv`)

| Column | Description |
|---|---|
| `sync_type` | full / incremental |
| `folder` | Inbox or SentItems |
| `started_at` | Sync start time (UTC) |
| `completed_at` | Sync end time (UTC) |
| `emails_synced` | Number of emails processed |

## Querying with PowerShell

```powershell
# Recent inbox emails
Import-Csv .\data\emails_inbox.csv | Sort-Object sent_datetime -Descending | Select-Object -First 10 subject, from_address, sent_datetime

# Emails with attachments still pending download
Import-Csv .\data\emails_inbox.csv | Where-Object { $_.has_attachments -eq "1" -and $_.attachments_downloaded -eq "0" } | Select-Object subject, from_address

# Search conversations
Import-Csv .\data\conversations.csv | Where-Object { $_.subject -like '*project update*' } | Select-Object subject, participants, message_count

# Full thread text for a conversation
(Import-Csv .\data\conversations.csv | Where-Object { $_.subject -like '*project update*' }).full_thread

# Sync history
Import-Csv .\data\sync_log.csv | Sort-Object completed_at -Descending
```
