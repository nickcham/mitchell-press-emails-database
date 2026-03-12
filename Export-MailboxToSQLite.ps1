<#
.SYNOPSIS
    M365 Mailbox Email Exporter — interactive menu with Full, Incremental, and Attachment download modes.

.DESCRIPTION
    Connects to Microsoft Graph via interactive browser login (supports MFA),
    then exports emails from Inbox and Sent Items into a SQLite database using MySQLite.

    Menu options:
    1) Full Export                — downloads every email (first run or full re-scan)
    2) Incremental Sync           — only fetches new/changed emails since last run
    3) Download Missing Attachments — scans DB for emails not yet downloaded, fetches to disk
    4) Build Conversations (Full) — rebuild conversations table from all emails
    5) Build Conversations (Incremental) — update only changed conversations
    6) Status & Quick Notes       — syncs, pending, throttling, custom app guide
    Q) Quit

    Uses MySQLite for modern, maintained SQLite operations.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

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

# ===================================================================
# DEPENDENCIES - MySQLite (modern replacement for PSSQLite)
# ===================================================================
$requiredModules = @("Microsoft.Graph.Authentication", "MySQLite")

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module: $mod ..." -ForegroundColor Yellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -Force
}

Write-Host "MySQLite v$((Get-Module MySQLite).Version) loaded — modern SQLite backend ready 💅" -ForegroundColor Green

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
    Write-Host "      Download only new/changed since last run" -ForegroundColor DarkGray
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
    Write-Host "      Show syncs, pending, throttling reminders & custom app guide" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [Q] Quit" -ForegroundColor Yellow
    Write-Host ""
    $choice = Read-Host "Select an option"
    return $choice
}

# ===================================================================
# AUTHENTICATION
# ===================================================================
function Connect-ToGraph {
    if ($script:ClientSecret) {
        if (-not $script:ClientId -or -not $script:TenantId -or -not $script:UserEmail) {
            throw "ClientId, TenantId, and UserEmail required for app-only auth."
        }
        $secureSecret = ConvertTo-SecureString $script:ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($script:ClientId, $secureSecret)
        Connect-MgGraph -TenantId $script:TenantId -ClientSecretCredential $credential -NoWelcome
        $script:basePath = "users/$($script:UserEmail)"
        Write-Host "Authenticated (app-only) for $($script:UserEmail)." -ForegroundColor Green
    } else {
        Write-Host "Opening browser for Microsoft 365 sign-in..." -ForegroundColor Cyan
        $connectParams = @{
            Scopes    = @("Mail.Read", "Mail.ReadWrite")
            NoWelcome = $true
        }
        if ($script:TenantId) { $connectParams.TenantId = $script:TenantId }
        if ($script:ClientId) {
            $connectParams.ClientId = $script:ClientId
            Write-Host "Using custom app registration: $($script:ClientId)" -ForegroundColor Green
        } else {
            Write-Host "Using default Microsoft Graph PowerShell app (consider custom for security)" -ForegroundColor Yellow
        }
        Connect-MgGraph @connectParams
        $script:basePath = "me"
        Write-Host "Authenticated via browser login." -ForegroundColor Green
    }
}

# ===================================================================
# DATABASE SETUP
# ===================================================================
function Initialize-Database {
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
    ai_category             TEXT,
    ai_confidence           TEXT,
    ai_summary              TEXT,
    ai_review_datetime      TEXT,
    ai_kb_confirmed         INTEGER,
    ai_kb_confirm_datetime  TEXT,
    ai_kb_confirm_notes     TEXT
);
"@

    Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $createEmailsTable
    Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $createAttachmentsTable
    Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $createSyncTable
    Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $createConversationsTable

    # Migrations
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
        try { Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $sql } catch {}
    }

    # Indexes
    $indexes = @(
        "CREATE INDEX IF NOT EXISTS idx_emails_conversation_id ON emails(conversation_id);",
        "CREATE INDEX IF NOT EXISTS idx_emails_last_modified ON emails(last_modified);",
        "CREATE INDEX IF NOT EXISTS idx_emails_attachments_pending ON emails(has_attachments, attachments_downloaded);",
        "CREATE INDEX IF NOT EXISTS idx_conversations_ai_review ON conversations(ai_kb_confirmed, ai_confidence);"
    )
    foreach ($sql in $indexes) {
        Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $sql
    }

    Write-Host "Database ready: $($script:DatabasePath)" -ForegroundColor Green
}

# ===================================================================
# HELPERS
# ===================================================================
function Format-Recipients {
    param([object[]]$Recipients)
    if (-not $Recipients) { return "" }
    return ($Recipients | ForEach-Object { "$($_.emailAddress.name) <$($_.emailAddress.address)>" }) -join "; "
}

function Get-LastSyncTime {
    param([string]$Folder)
    $result = Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
        "SELECT MAX(completed_at) AS last_sync FROM sync_log WHERE folder = @folder AND sync_type IN ('full','incremental')" `
        -SqlParameters @{ folder = $Folder }
    return $result.last_sync
}

# ===================================================================
# FETCH EMAILS with throttling
# ===================================================================
function Get-Emails {
    param(
        [string]$FolderName,
        [string]$BasePath,
        [string]$SinceDateTime
    )

    $graphFolder = $FolderName
    $fields = 'id,conversationId,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,sentDateTime,receivedDateTime,hasAttachments,importance,isRead,isDraft,body,bodyPreview,uniqueBody,webLink,categories,internetMessageId,parentFolderId,createdDateTime,lastModifiedDateTime'

    $url = "https://graph.microsoft.com/v1.0/$BasePath/mailFolders/$graphFolder/messages"
    $url += "?`$top=100&`$select=$fields"

    if ($SinceDateTime) {
        $url += "&`$filter=lastModifiedDateTime ge $SinceDateTime"
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
            $messages = $response.value
            if ($messages) { $allMessages += $messages }
            $url = $response.'@odata.nextLink'
        }
        catch {
            if ($_.Exception.Response.StatusCode -eq 429) {
                $retrySec = if ($_.Exception.Response.Headers["Retry-After"]) { [int]$_.Exception.Response.Headers["Retry-After"] } else { 60 }
                Write-Host "  Throttled (429)! Waiting $retrySec sec..." -ForegroundColor Red
                Start-Sleep -Seconds $retrySec
                continue
            }
            throw $_
        }

        Start-Sleep -Milliseconds (Get-Random -Minimum 700 -Maximum 1500)
    }

    return $allMessages
}

# ===================================================================
# SAVE EMAIL + compute cleaned_body
# ===================================================================
function Save-Email {
    param(
        [object]$Mail,
        [string]$Folder,
        [string]$DbPath
    )

    $fromName    = if ($Mail.from) { $Mail.from.emailAddress.name } else { "" }
    $fromAddress = if ($Mail.from) { $Mail.from.emailAddress.address } else { "" }

    $bodyContent = if ($Mail.uniqueBody -and $Mail.uniqueBody.content) { $Mail.uniqueBody.content } else { $Mail.body.content }
    $bodyType    = if ($Mail.uniqueBody) { $Mail.uniqueBody.contentType } else { $Mail.body.contentType }

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
        body_content        = $bodyContent
        body_preview        = $Mail.bodyPreview
        web_link            = $Mail.webLink
        folder              = $Folder
        categories          = ($Mail.categories -join "; ")
        internet_message_id = $Mail.internetMessageId
        parent_folder_id    = $Mail.parentFolderId
        created_datetime    = $Mail.createdDateTime
        last_modified       = $Mail.lastModifiedDateTime
    }

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
        Invoke-MySQLiteQuery -DataSource $DbPath -Query $upsertSQL -SqlParameters $params
    } catch {
        Write-Host "Save failed for $($Mail.id): $($_.Exception.Message)" -ForegroundColor Red
    }

    # Cleaned body
    $cleanBody = ConvertFrom-Html -Html $bodyContent
    $cleanBody = Remove-QuotedContent -Body $cleanBody

    try {
        Invoke-MySQLiteQuery -DataSource $DbPath -Query `
            "UPDATE emails SET cleaned_body = @clean WHERE message_id = @mid;" `
            -SqlParameters @{ clean = $cleanBody; mid = $Mail.id }
    } catch {
        Write-Host "Cleaned body update failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# ===================================================================
# DOWNLOAD MISSING ATTACHMENTS (with throttling & safe names)
# ===================================================================
function Invoke-DownloadMissingAttachments {
    $pending = @(Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
        "SELECT message_id, subject FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0")

    if ($pending.Count -eq 0) {
        Write-Host "`nNo missing attachments found. All caught up!" -ForegroundColor Green
        return
    }

    Write-Host "`nFound $($pending.Count) email(s) with attachments to download." -ForegroundColor Cyan

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
                if ($att.'@odata.type' -eq '#microsoft.graph.referenceAttachment' -or -not $att.contentBytes) { continue }

                $msgIdShort   = ($msgId -replace '[^a-zA-Z0-9]', '')[0..7] -join ''
                $safeAttId    = ($att.id -replace '[^a-zA-Z0-9]', '')[0..15] -join ''
                $safeFilename = $att.name -replace '[\\/:*?"<>|]', '_'
                $diskFilename = if ($safeFilename) { "${msgIdShort}_${safeAttId}_${safeFilename}" } else { "${msgIdShort}_${safeAttId}" }
                $diskFullPath = Join-Path $script:AttachmentPath $diskFilename

                $bytes = [Convert]::FromBase64String($att.contentBytes)
                [System.IO.File]::WriteAllBytes($diskFullPath, $bytes)

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
                Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $attSQL -SqlParameters $attParams
                $fileCount++
            }

            Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
                "UPDATE emails SET attachments_downloaded = 1 WHERE message_id = @mid;" `
                -SqlParameters @{ mid = $msgId }

        } catch {
            $errors++
            Write-Host "    ERROR on '$subj': $($_.Exception.Message)" -ForegroundColor Red
        }

        Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 1800)
    }

    $totalPending = (Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
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

        $syncEnd = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $logParams = @{
            sync_type    = $Mode
            folder       = $folder
            started_at   = $syncStart
            completed_at = $syncEnd
            emails_synced = $emails.Count
        }
        Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
            "INSERT INTO sync_log (sync_type, folder, started_at, completed_at, emails_synced) VALUES (@sync_type, @folder, @started_at, @completed_at, @emails_synced);" `
            -SqlParameters $logParams

        $totalCount += $emails.Count
        Write-Host "  Done: $folder ($($emails.Count) emails)" -ForegroundColor Green
    }

    # Optimize DB
    Write-Host "Optimizing database (VACUUM)..." -ForegroundColor Cyan
    Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query "VACUUM;"

    # Summary
    $dbEmailCount = (Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails").cnt
    $pendingAttach = (Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
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
        Write-Host "  Attachments pending: $pendingAttach (use option 3)" -ForegroundColor Yellow
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
    $text = $text -replace '<style[^>]*>[\s\S]*?</style>', ''
    $text = $text -replace '<script[^>]*>[\s\S]*?</script>', ''
    $text = $text -replace '<br\s*/?>', "`n"
    $text = $text -replace '</(?:p|div|h[1-6]|li|tr|blockquote)>', "`n"
    $text = $text -replace '<[^>]+>', ''
    $text = $text -replace '&nbsp;', ' '
    $text = $text -replace '&amp;', '&'
    $text = $text -replace '&lt;', '<'
    $text = $text -replace '&gt;', '>'
    $text = $text -replace '&quot;', '"'
    $text = $text -replace '&apos;', "'"
    $text = $text -replace '&#(\d+);', { [char][int]$_.Groups[1].Value }
    $text = ($text -split "`n" | ForEach-Object { ($_ -replace '\s+', ' ').Trim() } | Where-Object { $_ -ne '' }) -join "`n"
    return $text.Trim()
}

# ===================================================================
# THREAD DEDUPLICATION HELPER
# ===================================================================
function Remove-QuotedContent {
    param([string]$Body)
    if (-not $Body) { return "" }
    $lines = $Body -split "`n"
    $cutIndex = $lines.Count
    for ($j = 0; $j -lt $lines.Count; $j++) {
        $line = $lines[$j].Trim()
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
# BUILD CONVERSATIONS TABLE (uses cleaned_body)
# ===================================================================
function Build-ConversationRow {
    param([string]$ConvId)

    $emails = @(Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
        "SELECT subject, from_name, from_address, to_recipients, cc_recipients, sent_datetime, cleaned_body, body_preview, has_attachments, web_link FROM emails WHERE conversation_id = @cid ORDER BY sent_datetime ASC" `
        -SqlParameters @{ cid = $ConvId })

    if ($emails.Count -eq 0) { return }

    $subject = $emails[0].subject

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

    $threadParts = @()
    foreach ($e in $emails) {
        $cleanBody = if ($e.cleaned_body) { $e.cleaned_body } else { $e.body_preview }
        $header = "--- [" + $e.sent_datetime + "] From: " + $e.from_name + " <" + $e.from_address + "> ---"
        $toLine = "To: " + $e.to_recipients
        $ccLine = if ($e.cc_recipients) { "`nCC: " + $e.cc_recipients } else { "" }
        $threadParts += $header + "`n" + $toLine + $ccLine + "`n" + $cleanBody
    }
    $fullThread = $threadParts -join "`n`n"

    $hasAtt = [int](($emails | Where-Object { $_.has_attachments -eq 1 }).Count -gt 0)
    $now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $outlookLink = ($emails | Where-Object { $_.web_link } | Select-Object -Last 1).web_link

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
    Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query $convSQL -SqlParameters $convParams
}

function Invoke-BuildConversations {
    param(
        [string]$Mode
    )

    if ($Mode -eq "full") {
        Write-Host "`nBuilding conversations table from ALL emails..." -ForegroundColor Cyan
        $convIds = @(Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query `
            "SELECT DISTINCT conversation_id FROM emails WHERE conversation_id IS NOT NULL")
    } else {
        Write-Host "`nFinding conversations with new/changed emails..." -ForegroundColor Cyan
        $convIds = @(Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query @"
SELECT DISTINCT e.conversation_id
FROM emails e
LEFT JOIN conversations c ON e.conversation_id = c.conversation_id
WHERE e.conversation_id IS NOT NULL
  AND (c.last_built IS NULL OR e.last_modified > c.last_built)
"@
        )
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

    $totalConv = (Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM conversations").cnt

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
# STATUS & QUICK NOTES (with custom app guide & links)
# ===================================================================
function Show-StatusAndNotes {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host "   Quick Status & Brain Reminders 💅" -ForegroundColor Magenta
    Write-Host "============================================" -ForegroundColor Magenta
    Write-Host ""

    if ($script:ClientSecret) {
        Write-Host "Auth: App-only (unattended) for $($script:UserEmail)" -ForegroundColor Cyan
    } else {
        Write-Host "Auth: Interactive browser (MFA ok)" -ForegroundColor Cyan
        if ($script:ClientId) {
            Write-Host "  Using custom app: $($script:ClientId)" -ForegroundColor Green
        } else {
            Write-Host "  Using default Microsoft SDK app — consider custom for security!" -ForegroundColor Yellow
        }
    }
    Write-Host ""

    $lastInbox = Get-LastSyncTime -Folder "Inbox"
    $lastSent  = Get-LastSyncTime -Folder "SentItems"
    Write-Host "Last sync times:" -ForegroundColor White
    Write-Host "  Inbox   → $($lastInbox ?? 'Never')" -ForegroundColor DarkGray
    Write-Host "  Sent    → $($lastSent ?? 'Never')" -ForegroundColor DarkGray
    Write-Host ""

    $totalEmails = (Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails").cnt
    $pendingAtt  = (Invoke-MySQLiteQuery -DataSource $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0").cnt
    Write-Host "Database quick stats:" -ForegroundColor White
    Write-Host "  Emails stored:       $totalEmails" -ForegroundColor DarkGray
    Write-Host "  Attachments pending: $pendingAtt  (run [3] to grab 'em)" -ForegroundColor $(if ($pendingAtt -gt 0) {'Yellow'} else {'Green'})
    Write-Host ""

    Write-Host "Throttling reality check (2026 Graph Mail):" -ForegroundColor Yellow
    Write-Host "  • ~10,000 requests / 10 min per mailbox soft target" -ForegroundColor DarkYellow
    Write-Host "  • Sleeps + 429 handling in place — should stay safe" -ForegroundColor DarkYellow
    Write-Host "  • Full export big mailbox? Be patient" -ForegroundColor DarkYellow
    Write-Host ""

    Write-Host "Custom App Reg Guide — Own Your Shit (CIPP-safe!)" -ForegroundColor Green
    Write-Host "  Default SDK app is sketchy multi-tenant — create your own instead:" -ForegroundColor White
    Write-Host "  1. Go straight to App registrations:" -ForegroundColor DarkGray
    Write-Host "     https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/~/AllApplications" -ForegroundColor Cyan
    Write-Host "     → Click 'New registration'" -ForegroundColor DarkGray
    Write-Host "     Name: 'Nick's Mailbox Exporter' (or whatever slays)" -ForegroundColor DarkGray
    Write-Host "     Accounts: This org only (single tenant)" -ForegroundColor DarkGray
    Write-Host "     Redirect URI: Leave blank" -ForegroundColor DarkGray
    Write-Host "  2. After creation → API permissions → Add → Microsoft Graph → Delegated:" -ForegroundColor DarkGray
    Write-Host "     Mail.Read (add Mail.ReadWrite if needed)" -ForegroundColor DarkGray
    Write-Host "     → Grant admin consent for your tenant" -ForegroundColor DarkGray
    Write-Host "  3. Copy Application (client) ID from Overview" -ForegroundColor DarkGray
    Write-Host "  4. Run script with: -ClientId 'your-new-id-here'" -ForegroundColor DarkGray
    Write-Host "     Example: .\script.ps1 -DatabasePath .\emails.db -ClientId '1234abcd-...'" -ForegroundColor Cyan
    Write-Host "  • To check/manage your new app later:" -ForegroundColor DarkGray
    Write-Host "     https://entra.microsoft.com/#view/Microsoft_AAD_IAM/EnterpriseApplicationsMenuBlade/~/AllApps" -ForegroundColor Cyan
    Write-Host "  • Pre-consent in CIPP/Entra → no user prompts, audit-friendly" -ForegroundColor Yellow
    Write-Host ""

    Read-Host "Press Enter to return to menu..."
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
Write-Host "Disconnected from Microsoft Graph. Goodbye, my king!" -ForegroundColor Cyan
