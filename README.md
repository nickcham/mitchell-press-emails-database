# mitchell-press-emails-database

PowerShell script that exports every email from an M365 mailbox (Inbox + Sent Items) into a SQLite database.

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

### Interactive login (delegated)

```powershell
.\Export-MailboxToSQLite.ps1 `
    -DatabasePath ".\emails.db" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id"
```

A browser window will open for you to sign in.

### App-only (client credentials)

```powershell
.\Export-MailboxToSQLite.ps1 `
    -DatabasePath ".\emails.db" `
    -ClientId "your-client-id" `
    -TenantId "your-tenant-id" `
    -ClientSecret "your-client-secret" `
    -UserEmail "user@yourdomain.com"
```

## Database Schema

The SQLite database has a single `emails` table with these columns:

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

## Querying the Database

```powershell
Import-Module PSSQLite
Invoke-SqliteQuery -DataSource ".\emails.db" -Query "SELECT subject, from_address, sent_datetime FROM emails ORDER BY sent_datetime DESC LIMIT 10"
```

Or use any SQLite tool (DB Browser for SQLite, `sqlite3` CLI, DBeaver, etc.).

## Re-running

The script uses `INSERT OR REPLACE`, so re-running it updates existing records and adds new ones without duplicates.
