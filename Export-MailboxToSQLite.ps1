<#
.SYNOPSIS
    M365 Mailbox Email Exporter — interactive menu with Full, Incremental, and Attachment download modes.

.DESCRIPTION
    Connects to Microsoft Graph via interactive browser login (supports MFA),
    then exports emails from Inbox and Sent Items into a SQLite database.
    Auto-creates the database folder (default: .\data\) if it doesn't exist.

    Menu options:
    1) Full Export                — downloads every email (first run or full re-scan)
    2) Incremental Sync          — only fetches new/changed emails since last run
    3) Download Missing Attachments — scans DB for emails not yet downloaded, fetches to disk
    4) Build Conversations (Full) — rebuild conversations table from all emails
    5) Build Conversations (Incremental) — update only changed conversations
    6) Status & Quick Notes       — show syncs, pending, throttling reminders & custom app guide
    Q) Quit

    Email sync and attachment downloading are separate operations. Emails are marked
    with an attachments_downloaded flag so option 3 only fetches what's missing.

    By default, authentication uses interactive browser login — a browser window
    opens, you sign in with your Microsoft 365 account, complete MFA, and the
    script runs against your mailbox. No app registration required for personal use.

    For unattended/service scenarios, you can optionally supply ClientId, TenantId,
    ClientSecret, and UserEmail for app-only (client credentials) auth.

.PARAMETER DatabasePath
    Path to the SQLite database file. Created if it doesn't exist.

.PARAMETER AttachmentPath
    Folder where attachments are saved. Defaults to .\Attachments

.PARAMETER ClientId
    (Optional) Azure AD App Registration Client ID. Only needed for app-only auth.

.PARAMETER TenantId
    (Optional) Azure AD Tenant ID. Only needed for app-only auth.

.PARAMETER ClientSecret
    (Optional) Client Secret for app-only auth.

.PARAMETER UserEmail
    (Optional) Target mailbox email address. Required for app-only auth.

.EXAMPLE
    # Simple — browser login with MFA, default DB path (.\data\emails.db)
    .\Export-MailboxToSQLite.ps1

.EXAMPLE
    # Custom DB path
    .\Export-MailboxToSQLite.ps1 -DatabasePath ".\my-emails.db"

.EXAMPLE
    # App-only (unattended, client credentials)
    .\Export-MailboxToSQLite.ps1 -ClientId "abc" -TenantId "xyz" -ClientSecret "secret" -UserEmail "user@domain.com"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$DatabasePath = ".\data\emails.db",

    [Parameter(Mandatory = $false)]
    [string]$AttachmentPath = ".\Attachments",

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$UserEmail
)

$ErrorActionPreference = "Stop"

# Auto-create database folder if it doesn't exist
$dbFolder = Split-Path $DatabasePath -Parent
if ($dbFolder -and -not (Test-Path $dbFolder)) {
    New-Item -Path $dbFolder -ItemType Directory -Force | Out-Null
    Write-Host "Created folder: $dbFolder" -ForegroundColor Green
}

# ===================================================================
# MENU
# ===================================================================
function Show-Menu {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "   M365 Mailbox Email Exporter" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] Full Export" -ForegroundColor White
    Write-Host "      Download ALL emails (first run or full re-scan)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [2] Incremental Sync" -ForegroundColor White
    Write-Host "      Download only new/changed emails since last run" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [3] Download Missing Attachments" -ForegroundColor White
    Write-Host "      Scan DB and download attachments not yet saved to disk" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [4] Build Conversations (Full)" -ForegroundColor White
    Write-Host "      Rebuild conversations table from all emails in DB" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [5] Build Conversations (Incremental)" -ForegroundColor White
    Write-Host "      Update only conversations with new/changed emails" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [6] Status & Quick Notes" -ForegroundColor Green
    Write-Host "      Show last syncs, pending items, throttling reminders & custom app guide" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [Q] Quit" -ForegroundColor Yellow
    Write-Host ""
    $choice = Read-Host "Select an option"
    return $choice
}

# ===================================================================
# DEPENDENCIES
# ===================================================================
$requiredModules = @("Microsoft.Graph.Authentication", "MySQLite")

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module: $mod ..." -ForegroundColor Yellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -Force
}

# ===================================================================
# AUTHENTICATION
# ===================================================================
function Connect-ToGraph {
    if ($script:ClientSecret) {
        # App-only auth (client credentials) — unattended / service account
        if (-not $script:ClientId -or -not $script:TenantId) {
            throw "ClientId and TenantId are required when using ClientSecret authentication."
        }
        if (-not $script:UserEmail) {
            throw "UserEmail is required when using app-only (ClientSecret) authentication."
        }
        $secureSecret = ConvertTo-SecureString $script:ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($script:ClientId, $secureSecret)
        Connect-MgGraph -TenantId $script:TenantId -ClientSecretCredential $credential -NoWelcome
        $script:basePath = "users/$($script:UserEmail)"
        Write-Host "Authenticated (app-only) for $($script:UserEmail)." -ForegroundColor Green
    }
    else {
        # Interactive browser login — opens browser, you sign in with MFA
        Write-Host "Opening browser for Microsoft 365 sign-in..." -ForegroundColor Cyan
        $connectParams = @{
            Scopes    = @("Mail.Read", "Mail.ReadWrite")
            NoWelcome = $true
        }
        if ($script:TenantId) { $connectParams.TenantId = $script:TenantId }
        if ($script:ClientId) { $connectParams.ClientId = $script:ClientId }
        Connect-MgGraph @connectParams
        $script:basePath = "me"
        Write-Host "Authenticated via browser login." -ForegroundColor Green
    }
}

# ===================================================================
# DATABASE SETUP
# ===================================================================
function Initialize-Database {
    # MySQLite requires explicit DB file creation (unlike PSSQLite which auto-created)
    if (-not (Test-Path $script:DatabasePath)) {
        New-MySQLiteDB -Path $script:DatabasePath
        Write-Host "Created database: $($script:DatabasePath)" -ForegroundColor Cyan
    }

    $createEmailsTable = @"
CREATE TABLE IF NOT EXISTS emails (
    message_id              TEXT PRIMARY KEY,
    conversation_id         TEXT,
    subject                 TEXT,
    from_name               TEXT,
    from_address            TEXT,
    to_recipients           TEXT,
    cc_recipients           TEXT,
    bcc_recipients          TEXT,
    reply_to                TEXT,
    sent_datetime           TEXT,
    received_datetime       TEXT,
    has_attachments         INTEGER,
    attachments_downloaded  INTEGER DEFAULT 0,
    importance              TEXT,
    is_read                 INTEGER,
    is_draft                INTEGER,
    body_content_type       TEXT,
    body_content            TEXT,
    body_preview            TEXT,
    cleaned_body            TEXT,
    web_link                TEXT,
    folder                  TEXT,
    categories              TEXT,
    internet_message_id     TEXT,
    parent_folder_id        TEXT,
    created_datetime        TEXT,
    last_modified           TEXT
);
"@

    $createAttachmentsTable = @"
CREATE TABLE IF NOT EXISTS attachments (
    id              TEXT PRIMARY KEY,
    message_id      TEXT NOT NULL,
    filename        TEXT,
    content_type    TEXT,
    size_bytes      INTEGER,
    disk_path       TEXT,
    FOREIGN KEY (message_id) REFERENCES emails(message_id)
);
"@

    $createSyncTable = @"
CREATE TABLE IF NOT EXISTS sync_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    sync_type       TEXT,
    folder          TEXT,
    started_at      TEXT,
    completed_at    TEXT,
    emails_synced   INTEGER
);
"@

    $createConversationsTable = @"
CREATE TABLE IF NOT EXISTS conversations (
    conversation_id         TEXT PRIMARY KEY,
    subject                 TEXT,
    participants            TEXT,
    message_count           INTEGER,
    has_attachments         INTEGER,
    first_message_datetime  TEXT,
    last_message_datetime   TEXT,
    full_thread             TEXT,
    outlook_link            TEXT,
    last_built              TEXT,

    -- AI first-pass triage
    ai_category             TEXT,
    ai_confidence           TEXT,
    ai_summary              TEXT,
    ai_review_datetime      TEXT,

    -- AI second-pass confirmation
    ai_kb_confirmed         INTEGER,
    ai_kb_confirm_datetime  TEXT,
    ai_kb_confirm_notes     TEXT
);
"@

    Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $createEmailsTable
    Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $createAttachmentsTable
    Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $createConversationsTable
    Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $createSyncTable

    # Migrate existing databases — add columns that may not exist yet
    $migrations = @(
        "ALTER TABLE emails ADD COLUMN attachments_downloaded INTEGER DEFAULT 0;",
        "ALTER TABLE emails ADD COLUMN cleaned_body TEXT;",
        "ALTER TABLE conversations ADD COLUMN outlook_link TEXT;",
        "ALTER TABLE conversations ADD COLUMN ai_category TEXT;",
        "ALTER TABLE conversations ADD COLUMN ai_confidence TEXT;",
        "ALTER TABLE conversations ADD COLUMN ai_summary TEXT;",
        "ALTER TABLE conversations ADD COLUMN ai_review_datetime TEXT;",
        "ALTER TABLE conversations ADD COLUMN ai_kb_confirmed INTEGER;",
        "ALTER TABLE conversations ADD COLUMN ai_kb_confirm_datetime TEXT;",
        "ALTER TABLE conversations ADD COLUMN ai_kb_confirm_notes TEXT;"
    )
    foreach ($sql in $migrations) {
        Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $sql -ErrorAction SilentlyContinue
    }

    # Create indexes for performance
    $indexes = @(
        "CREATE INDEX IF NOT EXISTS idx_emails_conversation_id ON emails(conversation_id);",
        "CREATE INDEX IF NOT EXISTS idx_emails_last_modified ON emails(last_modified);",
        "CREATE INDEX IF NOT EXISTS idx_emails_attachments_pending ON emails(has_attachments, attachments_downloaded);",
        "CREATE INDEX IF NOT EXISTS idx_conversations_ai_review ON conversations(ai_kb_confirmed, ai_confidence);"
    )
    foreach ($sql in $indexes) {
        Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $sql
    }

    Write-Host "Database ready: $($script:DatabasePath)"
}

# ===================================================================
# HELPERS
# ===================================================================
function Format-Recipients {
    param([object[]]$Recipients)
    if (-not $Recipients) { return "" }
    return ($Recipients | ForEach-Object {
        "$($_.emailAddress.name) <$($_.emailAddress.address)>"
    }) -join "; "
}

function Get-LastSyncTime {
    param([string]$Folder)
    $result = Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
        "SELECT MAX(completed_at) AS last_sync FROM sync_log WHERE folder = @folder AND sync_type IN ('full','incremental')" `
        -SqlParameters @{ folder = $Folder }
    if ($result.last_sync) { return $result.last_sync }
    return $null
}

# ===================================================================
# FETCH EMAILS (supports incremental filter + 429 throttling)
# ===================================================================
function Get-Emails {
    param(
        [string]$FolderName,
        [string]$BasePath,
        [string]$SinceDateTime
    )

    $graphFolder = $FolderName  # "Inbox" or "SentItems"
    $fields = 'id,conversationId,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,sentDateTime,receivedDateTime,hasAttachments,importance,isRead,isDraft,body,bodyPreview,uniqueBody,webLink,categories,internetMessageId,parentFolderId,createdDateTime,lastModifiedDateTime'

    $url = "https://graph.microsoft.com/v1.0/$BasePath/mailFolders/$graphFolder/messages"
    $url += "?`$top=100&`$select=$fields"

    if ($SinceDateTime) {
        $filter = "lastModifiedDateTime ge $SinceDateTime"
        $url += "&`$filter=$filter"
        Write-Host "  Incremental filter: modified since $SinceDateTime" -ForegroundColor DarkYellow
    }

    $url += "&`$orderby=lastModifiedDateTime asc"

    $allMessages = @()
    $pageCount = 0

    while ($url) {
        $pageCount++
        Write-Host "  Fetching $FolderName page $pageCount ..." -ForegroundColor Cyan

        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
            if ($response.value) { $allMessages += $response.value }
            $url = $response.'@odata.nextLink'
        } catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                $retry = if ($_.Exception.Response.Headers["Retry-After"]) {
                    [int]$_.Exception.Response.Headers["Retry-After"]
                } else { 60 }
                Write-Host "  429 throttled — sleeping $retry sec..." -ForegroundColor Red
                Start-Sleep -Seconds $retry
                continue
            }
            throw
        }

        Start-Sleep -Milliseconds (Get-Random -Minimum 700 -Maximum 1500)
    }

    return $allMessages
}

# ===================================================================
# SAVE EMAIL (with cleaned_body generation)
# ===================================================================
function Save-Email {
    param(
        [object]$Mail,
        [string]$Folder,
        [string]$DbPath
    )

    $fromName    = if ($Mail.from) { $Mail.from.emailAddress.name } else { "" }
    $fromAddress = if ($Mail.from) { $Mail.from.emailAddress.address } else { "" }

    # Prefer uniqueBody (Graph's quote-stripped version) when available
    $bodyCont = if ($Mail.uniqueBody.content) { $Mail.uniqueBody.content } else { $Mail.body.content }
    $bodyType = if ($Mail.uniqueBody) { $Mail.uniqueBody.contentType } else { $Mail.body.contentType }

    $params = @{
        message_id          = $Mail.id
        conversation_id     = $Mail.conversationId
        subject             = $Mail.subject
        from_name           = $fromName
        from_address        = $fromAddress
        to_recipients       = (Format-Recipients $Mail.toRecipients)
        cc_recipients       = (Format-Recipients $Mail.ccRecipients)
        bcc_recipients      = (Format-Recipients $Mail.bccRecipients)
        reply_to            = (Format-Recipients $Mail.replyTo)
        sent_datetime       = $Mail.sentDateTime
        received_datetime   = $Mail.receivedDateTime
        has_attachments     = [int]$Mail.hasAttachments
        importance          = $Mail.importance
        is_read             = [int]$Mail.isRead
        is_draft            = [int]$Mail.isDraft
        body_content_type   = $bodyType
        body_content        = $bodyCont
        body_preview        = $Mail.bodyPreview
        web_link            = $Mail.webLink
        folder              = $Folder
        categories          = ($Mail.categories -join "; ")
        internet_message_id = $Mail.internetMessageId
        parent_folder_id    = $Mail.parentFolderId
        created_datetime    = $Mail.createdDateTime
        last_modified       = $Mail.lastModifiedDateTime
    }

    # Preserve existing attachments_downloaded flag if the row already exists
    $upsertSQL = @"
INSERT INTO emails (
    message_id, conversation_id, subject, from_name, from_address,
    to_recipients, cc_recipients, bcc_recipients, reply_to,
    sent_datetime, received_datetime, has_attachments, attachments_downloaded,
    importance, is_read, is_draft, body_content_type, body_content, body_preview,
    web_link, folder, categories, internet_message_id,
    parent_folder_id, created_datetime, last_modified
) VALUES (
    @message_id, @conversation_id, @subject, @from_name, @from_address,
    @to_recipients, @cc_recipients, @bcc_recipients, @reply_to,
    @sent_datetime, @received_datetime, @has_attachments,
    COALESCE((SELECT attachments_downloaded FROM emails WHERE message_id = @message_id), 0),
    @importance, @is_read, @is_draft, @body_content_type, @body_content, @body_preview,
    @web_link, @folder, @categories, @internet_message_id,
    @parent_folder_id, @created_datetime, @last_modified
)
ON CONFLICT(message_id) DO UPDATE SET
    conversation_id     = excluded.conversation_id,
    subject             = excluded.subject,
    from_name           = excluded.from_name,
    from_address        = excluded.from_address,
    to_recipients       = excluded.to_recipients,
    cc_recipients       = excluded.cc_recipients,
    bcc_recipients      = excluded.bcc_recipients,
    reply_to            = excluded.reply_to,
    sent_datetime       = excluded.sent_datetime,
    received_datetime   = excluded.received_datetime,
    has_attachments     = excluded.has_attachments,
    importance          = excluded.importance,
    is_read             = excluded.is_read,
    is_draft            = excluded.is_draft,
    body_content_type   = excluded.body_content_type,
    body_content        = excluded.body_content,
    body_preview        = excluded.body_preview,
    web_link            = excluded.web_link,
    folder              = excluded.folder,
    categories          = excluded.categories,
    internet_message_id = excluded.internet_message_id,
    parent_folder_id    = excluded.parent_folder_id,
    created_datetime    = excluded.created_datetime,
    last_modified       = excluded.last_modified;
"@

    try {
        Invoke-MySQLiteQuery -Path $DbPath -Query $upsertSQL -SqlParameters $params
    } catch {
        Write-Host "Save failed for $($Mail.id): $($_.Exception.Message)" -ForegroundColor Red
    }

    # Generate cleaned_body from HTML
    $clean = ConvertFrom-Html -Html $bodyCont
    $clean = Remove-QuotedContent -Body $clean

    try {
        Invoke-MySQLiteQuery -Path $DbPath -Query `
            "UPDATE emails SET cleaned_body = @c WHERE message_id = @id;" `
            -SqlParameters @{ c = $clean; id = $Mail.id }
    } catch {
        Write-Host "Cleaned body update failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ===================================================================
# DOWNLOAD MISSING ATTACHMENTS (standalone operation)
# ===================================================================
function Invoke-DownloadMissingAttachments {
    # Find emails that have attachments but haven't been downloaded yet
    $pending = @(Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
        "SELECT message_id, subject FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0")

    if ($pending.Count -eq 0) {
        Write-Host "`nNo missing attachments found. All caught up!" -ForegroundColor Green
        return
    }

    Write-Host "`nFound $($pending.Count) email(s) with attachments to download." -ForegroundColor Cyan

    # Ensure attachment folder exists
    if (-not (Test-Path $script:AttachmentPath)) {
        New-Item -Path $script:AttachmentPath -ItemType Directory -Force | Out-Null
    }
    Write-Host "Saving to: $($script:AttachmentPath)" -ForegroundColor Cyan

    $downloaded = 0
    $fileCount = 0
    $errors = 0

    foreach ($row in $pending) {
        $downloaded++
        $msgId = $row.message_id
        $subj  = $row.subject
        if ($downloaded % 10 -eq 0 -or $downloaded -eq $pending.Count) {
            Write-Host "  Processing $downloaded / $($pending.Count) ..." -ForegroundColor DarkGray
        }

        try {
            $url = "https://graph.microsoft.com/v1.0/$($script:basePath)/messages/$msgId/attachments"
            $response = Invoke-MgGraphRequest -Method GET -Uri $url

            foreach ($att in $response.value) {
                # Skip reference attachments and items without content
                if ($att.'@odata.type' -eq '#microsoft.graph.referenceAttachment') { continue }
                if (-not $att.contentBytes) { continue }

                # Build safe filename: <message_id_short>_<attachment_id_short>_<original_filename>
                $msgIdShort   = ($msgId -replace '[^a-zA-Z0-9]', '')[0..7] -join ''
                $safeAttId    = ($att.id -replace '[^a-zA-Z0-9]', '')[0..15] -join ''
                $safeFilename = $att.name -replace '[\\/:*?"<>|]', '_'
                $diskFilename = "${msgIdShort}_${safeAttId}_${safeFilename}"
                $diskFullPath  = Join-Path $script:AttachmentPath $diskFilename

                # Save file to disk
                $bytes = [Convert]::FromBase64String($att.contentBytes)
                [System.IO.File]::WriteAllBytes($diskFullPath, $bytes)

                # Record in attachments table (parameterized — no SQL injection)
                $attParams = @{
                    id           = $att.id
                    message_id   = $msgId
                    filename     = $att.name
                    content_type = $att.contentType
                    size_bytes   = $att.size
                    disk_path    = $diskFullPath
                }

                $attSQL = @"
INSERT INTO attachments (id, message_id, filename, content_type, size_bytes, disk_path)
VALUES (@id, @message_id, @filename, @content_type, @size_bytes, @disk_path)
ON CONFLICT(id) DO UPDATE SET
    message_id   = excluded.message_id,
    filename     = excluded.filename,
    content_type = excluded.content_type,
    size_bytes   = excluded.size_bytes,
    disk_path    = excluded.disk_path;
"@
                Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $attSQL -SqlParameters $attParams
                $fileCount++
            }

            # Mark this email as downloaded
            Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
                "UPDATE emails SET attachments_downloaded = 1 WHERE message_id = @mid;" `
                -SqlParameters @{ mid = $msgId }

        } catch {
            $errors++
            Write-Host "    ERROR on '$subj': $($_.Exception.Message)" -ForegroundColor Red
        }

        Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 1800)
    }

    # Summary
    $totalPending = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
        "SELECT COUNT(*) AS cnt FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0").cnt

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   ATTACHMENT DOWNLOAD COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Emails processed:    $downloaded"
    Write-Host "  Files saved:         $fileCount"
    Write-Host "  Errors:              $errors"
    Write-Host "  Still pending:       $totalPending"
    Write-Host "  Attachment folder:   $($script:AttachmentPath)"
    Write-Host ""
}

# ===================================================================
# EMAIL SYNC ENGINE
# ===================================================================
function Invoke-Sync {
    param(
        [string]$Mode  # "full" or "incremental"
    )

    $folders = @("Inbox", "SentItems")
    $totalCount = 0

    foreach ($folder in $folders) {
        Write-Host "`n--- Processing: $folder ---" -ForegroundColor Cyan
        $syncStart = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

        # Determine incremental filter
        $sinceDateTime = $null
        if ($Mode -ne "full") {
            $sinceDateTime = Get-LastSyncTime -Folder $folder
        }

        $emails = Get-Emails -FolderName $folder -BasePath $script:basePath -SinceDateTime $sinceDateTime
        Write-Host "  Found $($emails.Count) emails to process." -ForegroundColor White

        $i = 0
        foreach ($mail in $emails) {
            $i++
            if ($i % 50 -eq 0 -or $i -eq $emails.Count) {
                Write-Host "    Processing $i / $($emails.Count) ..." -ForegroundColor DarkGray
            }
            Save-Email -Mail $mail -Folder $folder -DbPath $script:DatabasePath
        }

        # Log the sync (parameterized — no SQL injection)
        $syncEnd = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $logParams = @{
            sync_type    = $Mode
            folder       = $folder
            started_at   = $syncStart
            completed_at = $syncEnd
            emails_synced = $emails.Count
        }
        Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
            "INSERT INTO sync_log (sync_type, folder, started_at, completed_at, emails_synced) VALUES (@sync_type, @folder, @started_at, @completed_at, @emails_synced);" `
            -SqlParameters $logParams

        $totalCount += $emails.Count
        Write-Host "  Done: $folder ($($emails.Count) emails)" -ForegroundColor Green
    }

    # Reclaim space after bulk operations
    Invoke-MySQLiteQuery -Path $script:DatabasePath -Query "VACUUM;"

    # Summary
    $dbEmailCount = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails").cnt
    $pendingAttach = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
        "SELECT COUNT(*) AS cnt FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0").cnt

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   SYNC COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Mode:                $Mode"
    Write-Host "  Emails processed:    $totalCount"
    Write-Host "  Total in database:   $dbEmailCount"
    Write-Host "  Database file:       $($script:DatabasePath)"
    if ($pendingAttach -gt 0) {
        Write-Host "  Attachments pending: $pendingAttach (use option 3 to download)" -ForegroundColor Yellow
    }
    Write-Host ""
}

# ===================================================================
# HTML-TO-TEXT HELPER
# ===================================================================
function ConvertFrom-Html {
    param([string]$Html)
    if (-not $Html) { return "" }

    $text = $Html
    # Remove style and script blocks entirely (content + tags)
    $text = $text -replace '<style[^>]*>[\s\S]*?</style>', ''
    $text = $text -replace '<script[^>]*>[\s\S]*?</script>', ''
    # Convert <br> and <br/> to newlines
    $text = $text -replace '<br\s*/?>', "`n"
    # Convert block-level closing tags to newlines (paragraphs, divs, headings, list items)
    $text = $text -replace '</(?:p|div|h[1-6]|li|tr|blockquote)>', "`n"
    # Strip all remaining HTML tags
    $text = $text -replace '<[^>]+>', ''
    # Decode common HTML entities
    $text = $text -replace '&nbsp;', ' '
    $text = $text -replace '&amp;', '&'
    $text = $text -replace '&lt;', '<'
    $text = $text -replace '&gt;', '>'
    $text = $text -replace '&quot;', '"'
    $text = $text -replace '&apos;', "'"
    $text = $text -replace '&#(\d+);', { [char][int]$_.Groups[1].Value }
    # Collapse runs of whitespace on each line, but preserve line breaks
    $text = ($text -split "`n" | ForEach-Object { ($_ -replace '\s+', ' ').Trim() } | Where-Object { $_ -ne '' }) -join "`n"
    return $text.Trim()
}

# ===================================================================
# THREAD DEDUPLICATION HELPER
# ===================================================================
function Remove-QuotedContent {
    param([string]$Body)
    if (-not $Body) { return "" }

    # Split into lines and find where quoted/forwarded content starts
    $lines = $Body -split "`n"
    $cutIndex = $lines.Count

    for ($j = 0; $j -lt $lines.Count; $j++) {
        $line = $lines[$j].Trim()
        # Common Outlook quote markers
        if (($line -match '^-{2,}\s*(Original Message|Forwarded message)') -or
            ($line -match '^From:\s+.+@' -and $j -gt 0 -and $lines[$j-1].Trim() -match '^(Sent|Date):') -or
            ($line -match '^_{10,}') -or
            ($line -match '^\s*On .+ wrote:\s*$') -or
            ($line -match '^>{2,}')) {
            $cutIndex = $j
            break
        }
    }

    return ($lines[0..([Math]::Max(0, $cutIndex - 1))] -join "`n").Trim()
}

# ===================================================================
# BUILD CONVERSATIONS TABLE
# ===================================================================
function Build-ConversationRow {
    param([string]$ConvId)

    # Get all emails in this conversation, oldest first (parameterized)
    $emails = @(Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
        "SELECT subject, from_name, from_address, to_recipients, cc_recipients, sent_datetime, body_content, body_preview, has_attachments, web_link FROM emails WHERE conversation_id = @cid ORDER BY sent_datetime ASC" `
        -SqlParameters @{ cid = $ConvId })

    if ($emails.Count -eq 0) { return }

    # Subject from the first email in the thread
    $subject = $emails[0].subject

    # Collect unique participants from all from/to/cc fields
    $allParticipants = @{}
    foreach ($e in $emails) {
        if ($e.from_address) { $allParticipants[$e.from_address.ToLower()] = $e.from_name }
        foreach ($field in @($e.to_recipients, $e.cc_recipients)) {
            if (-not $field) { continue }
            foreach ($entry in ($field -split '; ')) {
                if ($entry -match '<(.+)>') {
                    $allParticipants[$Matches[1].ToLower()] = $entry
                }
            }
        }
    }
    $participants = ($allParticipants.Keys | Sort-Object) -join "; "

    # Build the full thread text for AI consumption
    $threadParts = @()
    foreach ($e in $emails) {
        $body = if ($e.body_content) { $e.body_content } else { $e.body_preview }
        $cleanBody = ConvertFrom-Html -Html $body
        $cleanBody = Remove-QuotedContent -Body $cleanBody

        # Use string concatenation to avoid here-string interpolation risks
        $header = "--- [" + $e.sent_datetime + "] From: " + $e.from_name + " <" + $e.from_address + "> ---"
        $toLine = "To: " + $e.to_recipients
        $ccLine = if ($e.cc_recipients) { "`nCC: " + $e.cc_recipients } else { "" }
        $threadParts += $header + "`n" + $toLine + $ccLine + "`n" + $cleanBody
    }
    $fullThread = $threadParts -join "`n`n"

    $hasAtt = [int](($emails | Where-Object { $_.has_attachments -eq 1 }).Count -gt 0)
    $now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")

    # Use the most recent email's web_link so reviewer can open the conversation
    $outlookLink = ($emails | Where-Object { $_.web_link } | Select-Object -Last 1).web_link

    # Parameterized upsert — no SQL injection
    $convParams = @{
        conversation_id        = $ConvId
        subject                = $subject
        participants           = $participants
        message_count          = $emails.Count
        has_attachments        = $hasAtt
        first_message_datetime = $emails[0].sent_datetime
        last_message_datetime  = $emails[-1].sent_datetime
        full_thread            = $fullThread
        outlook_link           = $outlookLink
        last_built             = $now
    }

    # Upsert: rebuild thread data but preserve existing AI review columns
    $convSQL = @"
INSERT INTO conversations (
    conversation_id, subject, participants, message_count, has_attachments,
    first_message_datetime, last_message_datetime, full_thread, outlook_link, last_built
) VALUES (
    @conversation_id, @subject, @participants, @message_count, @has_attachments,
    @first_message_datetime, @last_message_datetime, @full_thread, @outlook_link, @last_built
)
ON CONFLICT(conversation_id) DO UPDATE SET
    subject                = excluded.subject,
    participants           = excluded.participants,
    message_count          = excluded.message_count,
    has_attachments        = excluded.has_attachments,
    first_message_datetime = excluded.first_message_datetime,
    last_message_datetime  = excluded.last_message_datetime,
    full_thread            = excluded.full_thread,
    outlook_link           = excluded.outlook_link,
    last_built             = excluded.last_built;
"@
    Invoke-MySQLiteQuery -Path $script:DatabasePath -Query $convSQL -SqlParameters $convParams
}

function Invoke-BuildConversations {
    param(
        [string]$Mode  # "full" or "incremental"
    )

    if ($Mode -eq "full") {
        Write-Host "`nBuilding conversations table from ALL emails..." -ForegroundColor Cyan
        $convIds = @(Invoke-MySQLiteQuery -Path $script:DatabasePath -Query `
            "SELECT DISTINCT conversation_id FROM emails WHERE conversation_id IS NOT NULL")
    }
    else {
        # Incremental: find conversations where any email was modified after the conversation was last built
        Write-Host "`nFinding conversations with new/changed emails..." -ForegroundColor Cyan
        $convIds = @(Invoke-MySQLiteQuery -Path $script:DatabasePath -Query @"
SELECT DISTINCT e.conversation_id
FROM emails e
LEFT JOIN conversations c ON e.conversation_id = c.conversation_id
WHERE e.conversation_id IS NOT NULL
  AND (c.last_built IS NULL OR e.last_modified > c.last_built)
"@)
    }

    if ($convIds.Count -eq 0) {
        Write-Host "No conversations to process. All up to date!" -ForegroundColor Green
        return
    }

    Write-Host "  Processing $($convIds.Count) conversation(s)..." -ForegroundColor White

    $i = 0
    foreach ($row in $convIds) {
        $i++
        if ($i % 50 -eq 0 -or $i -eq $convIds.Count) {
            Write-Host "    Building $i / $($convIds.Count) ..." -ForegroundColor DarkGray
        }
        Build-ConversationRow -ConvId $row.conversation_id
    }

    $totalConv = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM conversations").cnt

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   CONVERSATIONS BUILD COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Mode:                $Mode"
    Write-Host "  Conversations built: $($convIds.Count)"
    Write-Host "  Total in table:      $totalConv"
    Write-Host ""
}

# ===================================================================
# STATUS & QUICK NOTES
# ===================================================================
function Show-StatusAndNotes {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "   Quick Status & Brain Reminders" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host ""

    # Auth reminder
    if ($script:ClientSecret) {
        Write-Host "Auth: App-only for $($script:UserEmail)" -ForegroundColor Cyan
    } else {
        Write-Host "Auth: Interactive browser (MFA ok)" -ForegroundColor Cyan
        if ($script:ClientId) {
            Write-Host "  Custom app: $($script:ClientId)" -ForegroundColor Green
        } else {
            Write-Host "  Default SDK app — make your own custom one soon!" -ForegroundColor Yellow
        }
    }
    Write-Host ""

    # Last syncs
    $lastInbox = Get-LastSyncTime -Folder "Inbox"
    $lastSent  = Get-LastSyncTime -Folder "SentItems"
    Write-Host "Last sync times:" -ForegroundColor White
    Write-Host "  Inbox: $($lastInbox ?? 'Never yet')" -ForegroundColor DarkGray
    Write-Host "  Sent:  $($lastSent ?? 'Never yet')" -ForegroundColor DarkGray
    Write-Host ""

    # Stats
    $emailCnt = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails").cnt
    $convCnt = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM conversations").cnt
    $pendingAtt = (Invoke-MySQLiteQuery -Path $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0").cnt
    Write-Host "Emails in DB:        $emailCnt" -ForegroundColor White
    Write-Host "Conversations:       $convCnt" -ForegroundColor White
    Write-Host "Pending attachments: $pendingAtt" -ForegroundColor $(if ($pendingAtt -gt 0) {'Yellow'} else {'Green'})
    Write-Host ""

    Write-Host "Quick Notes:" -ForegroundColor Yellow
    Write-Host "  - 429 throttling is handled automatically (Retry-After header)" -ForegroundColor DarkGray
    Write-Host "  - Conversations table preserves AI review columns on rebuild" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "Custom App Quick Guide:" -ForegroundColor Green
    Write-Host "  1. https://entra.microsoft.com/ -> App registrations -> New" -ForegroundColor Cyan
    Write-Host "     Single tenant, no redirect URI" -ForegroundColor DarkGray
    Write-Host "  2. Add delegated Mail.Read -> Grant admin consent" -ForegroundColor DarkGray
    Write-Host "  3. Copy Client ID -> rerun with -ClientId 'your-id-here'" -ForegroundColor DarkGray
    Write-Host ""

    Read-Host "Press Enter to go back to menu..."
}

# ===================================================================
# MAIN LOOP
# ===================================================================
Connect-ToGraph
Initialize-Database

$running = $true
while ($running) {
    $choice = Show-Menu

    switch ($choice) {
        "1" { Invoke-Sync -Mode "full" }
        "2" { Invoke-Sync -Mode "incremental" }
        "3" { Invoke-DownloadMissingAttachments }
        "4" { Invoke-BuildConversations -Mode "full" }
        "5" { Invoke-BuildConversations -Mode "incremental" }
        "6" { Show-StatusAndNotes }
        "Q" { $running = $false }
        "q" { $running = $false }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
        }
    }
}

Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph. Goodbye!" -ForegroundColor Cyan
