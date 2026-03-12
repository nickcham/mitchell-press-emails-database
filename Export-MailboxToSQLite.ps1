<#
.SYNOPSIS
    M365 Mailbox Email Exporter — interactive menu with Full, Incremental, and Attachment download modes.

.DESCRIPTION
    Connects to Microsoft Graph via interactive browser login (supports MFA),
    then exports emails from Inbox and Sent Items into a SQLite database.

    Menu options:
    1) Full Export                — downloads every email (first run or full re-scan)
    2) Incremental Sync          — only fetches new/changed emails since last run
    3) Download Missing Attachments — scans DB for emails not yet downloaded, fetches to disk
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
    # Simple — browser login with MFA (most common)
    .\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db"

.EXAMPLE
    # App-only (unattended, client credentials)
    .\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db" -ClientId "abc" -TenantId "xyz" -ClientSecret "secret" -UserEmail "user@domain.com"
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
    Write-Host "  [Q] Quit" -ForegroundColor Yellow
    Write-Host ""
    $choice = Read-Host "Select an option"
    return $choice
}

# ===================================================================
# DEPENDENCIES
# ===================================================================
$requiredModules = @("Microsoft.Graph.Authentication", "PSSQLite")

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module: $mod ..."
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

    Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $createEmailsTable
    Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $createAttachmentsTable
    Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $createSyncTable

    # Add attachments_downloaded column to existing databases that lack it
    try {
        Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
            "ALTER TABLE emails ADD COLUMN attachments_downloaded INTEGER DEFAULT 0;"
    } catch {
        # Column already exists — ignore
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
    $result = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
        "SELECT MAX(completed_at) AS last_sync FROM sync_log WHERE folder = @folder AND sync_type IN ('full','incremental')" `
        -SqlParameters @{ folder = $Folder }
    if ($result.last_sync) { return $result.last_sync }
    return $null
}

# ===================================================================
# FETCH EMAILS (supports incremental filter)
# ===================================================================
function Get-Emails {
    param(
        [string]$FolderName,
        [string]$BasePath,
        [string]$SinceDateTime
    )

    $graphFolder = $FolderName  # "Inbox" or "SentItems"
    $fields = 'id,conversationId,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,sentDateTime,receivedDateTime,hasAttachments,importance,isRead,isDraft,body,bodyPreview,webLink,categories,internetMessageId,parentFolderId,createdDateTime,lastModifiedDateTime'

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
        Write-Host "  Fetching $FolderName page $pageCount ..."
        $response = Invoke-MgGraphRequest -Method GET -Uri $url
        $messages = $response.value
        if ($messages) {
            $allMessages += $messages
        }
        $url = $response.'@odata.nextLink'
    }

    return $allMessages
}

# ===================================================================
# SAVE EMAIL
# ===================================================================
function Save-Email {
    param(
        [object]$Mail,
        [string]$Folder,
        [string]$DbPath
    )

    $fromName    = if ($Mail.from) { $Mail.from.emailAddress.name } else { "" }
    $fromAddress = if ($Mail.from) { $Mail.from.emailAddress.address } else { "" }

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
        body_content_type   = $Mail.body.contentType
        body_content        = $Mail.body.content
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

    Invoke-SqliteQuery -DataSource $DbPath -Query $upsertSQL -SqlParameters $params
}

# ===================================================================
# DOWNLOAD MISSING ATTACHMENTS (standalone operation)
# ===================================================================
function Invoke-DownloadMissingAttachments {
    # Find emails that have attachments but haven't been downloaded yet
    $pending = Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
        "SELECT message_id, subject FROM emails WHERE has_attachments = 1 AND attachments_downloaded = 0"

    if (-not $pending -or $pending.Count -eq 0) {
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

                # Build safe filename: <attachment_id_short>_<original_filename>
                # Uses attachment ID (not message ID) so multiple attachments per email don't collide
                $safeAttId    = ($att.id -replace '[^a-zA-Z0-9]', '')[0..19] -join ''
                $safeFilename = $att.name -replace '[\\/:*?"<>|]', '_'
                $diskFilename = "${safeAttId}_${safeFilename}"
                $diskFullPath  = Join-Path $script:AttachmentPath $diskFilename

                # Save file to disk
                $bytes = [Convert]::FromBase64String($att.contentBytes)
                [System.IO.File]::WriteAllBytes($diskFullPath, $bytes)

                # Record in attachments table
                $attParams = @{
                    id           = $att.id
                    message_id   = $msgId
                    filename     = $att.name
                    content_type = $att.contentType
                    size_bytes   = $att.size
                    disk_path    = $diskFullPath
                }

                $attSQL = @"
INSERT OR REPLACE INTO attachments (id, message_id, filename, content_type, size_bytes, disk_path)
VALUES (@id, @message_id, @filename, @content_type, @size_bytes, @disk_path);
"@
                Invoke-SqliteQuery -DataSource $script:DatabasePath -Query $attSQL -SqlParameters $attParams
                $fileCount++
            }

            # Mark this email as downloaded
            Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
                "UPDATE emails SET attachments_downloaded = 1 WHERE message_id = @mid;" `
                -SqlParameters @{ mid = $msgId }

        } catch {
            $errors++
            Write-Host "    ERROR on '$subj': $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    # Summary
    $totalPending = (Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
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

        # Log the sync
        $syncEnd = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $logParams = @{
            sync_type    = $Mode
            folder       = $folder
            started_at   = $syncStart
            completed_at = $syncEnd
            emails_synced = $emails.Count
        }
        Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
            "INSERT INTO sync_log (sync_type, folder, started_at, completed_at, emails_synced) VALUES (@sync_type, @folder, @started_at, @completed_at, @emails_synced);" `
            -SqlParameters $logParams

        $totalCount += $emails.Count
        Write-Host "  Done: $folder ($($emails.Count) emails)" -ForegroundColor Green
    }

    # Summary
    $dbEmailCount = (Invoke-SqliteQuery -DataSource $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails").cnt
    $pendingAttach = (Invoke-SqliteQuery -DataSource $script:DatabasePath -Query `
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
        "Q" { $running = $false }
        "q" { $running = $false }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
        }
    }
}

Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph. Goodbye!" -ForegroundColor Cyan
