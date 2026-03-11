<#
.SYNOPSIS
    M365 Mailbox Email Exporter — interactive menu with Full, Incremental, and Attachment modes.

.DESCRIPTION
    Connects to Microsoft Graph API and exports emails from Inbox and Sent Items
    into a SQLite database. Offers three operating modes via an interactive menu:

    1) Full Export         — wipes and re-downloads every email
    2) Incremental Sync   — only fetches new/changed emails since last run
    3) Incremental + Attachments — same as #2, plus downloads attachments to disk

.PARAMETER DatabasePath
    Path to the SQLite database file. Created if it doesn't exist.

.PARAMETER ClientId
    Azure AD App Registration Client ID.

.PARAMETER TenantId
    Azure AD Tenant ID.

.PARAMETER ClientSecret
    Client Secret for app-only auth. If omitted, uses interactive (delegated) auth.

.PARAMETER UserEmail
    The mailbox email address to export (required for app-only auth).

.PARAMETER AttachmentPath
    Folder where attachments are saved. Defaults to .\Attachments

.EXAMPLE
    .\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db" -ClientId "abc" -TenantId "xyz"
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$DatabasePath,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [string]$UserEmail,

    [Parameter(Mandatory = $false)]
    [string]$AttachmentPath = ".\Attachments"
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
    Write-Host "      Wipe database and download ALL emails" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [2] Incremental Sync" -ForegroundColor White
    Write-Host "      Download only new/changed emails since last run" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [3] Incremental Sync + Download Attachments" -ForegroundColor White
    Write-Host "      Same as #2, plus save attachments to disk" -ForegroundColor DarkGray
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
        if (-not $script:UserEmail) {
            throw "UserEmail is required when using app-only (ClientSecret) authentication."
        }
        $secureSecret = ConvertTo-SecureString $script:ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($script:ClientId, $secureSecret)
        Connect-MgGraph -TenantId $script:TenantId -ClientSecretCredential $credential -NoWelcome
        $script:basePath = "users/$($script:UserEmail)"
    }
    else {
        Connect-MgGraph -TenantId $script:TenantId -ClientId $script:ClientId -Scopes "Mail.Read", "Mail.ReadWrite" -NoWelcome
        $script:basePath = "me"
    }
    Write-Host "Authenticated to Microsoft Graph." -ForegroundColor Green
}

# ===================================================================
# DATABASE SETUP
# ===================================================================
function Initialize-Database {
    $createEmailsTable = @"
CREATE TABLE IF NOT EXISTS emails (
    message_id          TEXT PRIMARY KEY,
    conversation_id     TEXT,
    subject             TEXT,
    from_name           TEXT,
    from_address        TEXT,
    to_recipients       TEXT,
    cc_recipients       TEXT,
    bcc_recipients      TEXT,
    reply_to            TEXT,
    sent_datetime       TEXT,
    received_datetime   TEXT,
    has_attachments     INTEGER,
    importance          TEXT,
    is_read             INTEGER,
    is_draft            INTEGER,
    body_content_type   TEXT,
    body_content        TEXT,
    body_preview        TEXT,
    web_link            TEXT,
    folder              TEXT,
    categories          TEXT,
    internet_message_id TEXT,
    parent_folder_id    TEXT,
    created_datetime    TEXT,
    last_modified       TEXT
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
        "SELECT MAX(completed_at) AS last_sync FROM sync_log WHERE folder = @folder" `
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

    # For incremental: only fetch emails modified after last sync
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

    $upsertSQL = @"
INSERT OR REPLACE INTO emails (
    message_id, conversation_id, subject, from_name, from_address,
    to_recipients, cc_recipients, bcc_recipients, reply_to,
    sent_datetime, received_datetime, has_attachments, importance,
    is_read, is_draft, body_content_type, body_content, body_preview,
    web_link, folder, categories, internet_message_id,
    parent_folder_id, created_datetime, last_modified
) VALUES (
    @message_id, @conversation_id, @subject, @from_name, @from_address,
    @to_recipients, @cc_recipients, @bcc_recipients, @reply_to,
    @sent_datetime, @received_datetime, @has_attachments, @importance,
    @is_read, @is_draft, @body_content_type, @body_content, @body_preview,
    @web_link, @folder, @categories, @internet_message_id,
    @parent_folder_id, @created_datetime, @last_modified
);
"@

    Invoke-SqliteQuery -DataSource $DbPath -Query $upsertSQL -SqlParameters $params
}

# ===================================================================
# DOWNLOAD ATTACHMENTS
# ===================================================================
function Save-Attachments {
    param(
        [object]$Mail,
        [string]$BasePath,
        [string]$DbPath,
        [string]$AttachFolder
    )

    if (-not $Mail.hasAttachments) { return }

    $url = "https://graph.microsoft.com/v1.0/$BasePath/messages/$($Mail.id)/attachments"
    $response = Invoke-MgGraphRequest -Method GET -Uri $url

    foreach ($att in $response.value) {
        # Skip inline images / reference attachments without content
        if ($att.'@odata.type' -eq '#microsoft.graph.referenceAttachment') { continue }
        if (-not $att.contentBytes) { continue }

        # Build safe filename: <message_id_short>_<original_filename>
        $safeMessageId = ($Mail.id -replace '[^a-zA-Z0-9]', '')[0..19] -join ''
        $safeFilename  = $att.name -replace '[\\/:*?"<>|]', '_'
        $diskFilename  = "${safeMessageId}_${safeFilename}"
        $diskFullPath  = Join-Path $AttachFolder $diskFilename

        # Save file to disk
        $bytes = [Convert]::FromBase64String($att.contentBytes)
        [System.IO.File]::WriteAllBytes($diskFullPath, $bytes)

        # Record in database
        $attParams = @{
            id           = $att.id
            message_id   = $Mail.id
            filename     = $att.name
            content_type = $att.contentType
            size_bytes   = $att.size
            disk_path    = $diskFullPath
        }

        $attSQL = @"
INSERT OR REPLACE INTO attachments (id, message_id, filename, content_type, size_bytes, disk_path)
VALUES (@id, @message_id, @filename, @content_type, @size_bytes, @disk_path);
"@
        Invoke-SqliteQuery -DataSource $DbPath -Query $attSQL -SqlParameters $attParams
    }
}

# ===================================================================
# SYNC ENGINE
# ===================================================================
function Invoke-Sync {
    param(
        [string]$Mode  # "full", "incremental", "incremental+attachments"
    )

    $folders = @("Inbox", "SentItems")
    $totalCount = 0
    $totalAttachments = 0
    $downloadAttachments = ($Mode -eq "incremental+attachments")

    # For full export, wipe existing data
    if ($Mode -eq "full") {
        Write-Host "`nWiping existing data for full export..." -ForegroundColor Yellow
        Invoke-SqliteQuery -DataSource $script:DatabasePath -Query "DELETE FROM attachments;"
        Invoke-SqliteQuery -DataSource $script:DatabasePath -Query "DELETE FROM emails;"
        Invoke-SqliteQuery -DataSource $script:DatabasePath -Query "DELETE FROM sync_log;"
    }

    # Ensure attachment folder exists
    if ($downloadAttachments) {
        if (-not (Test-Path $script:AttachmentPath)) {
            New-Item -Path $script:AttachmentPath -ItemType Directory -Force | Out-Null
        }
        Write-Host "Attachments will be saved to: $($script:AttachmentPath)" -ForegroundColor Cyan
    }

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

            if ($downloadAttachments -and $mail.hasAttachments) {
                Save-Attachments -Mail $mail -BasePath $script:basePath -DbPath $script:DatabasePath -AttachFolder $script:AttachmentPath
                $totalAttachments++
            }
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
    $dbAttachCount = (Invoke-SqliteQuery -DataSource $script:DatabasePath -Query "SELECT COUNT(*) AS cnt FROM attachments").cnt

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   SYNC COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Mode:                $Mode"
    Write-Host "  Emails processed:    $totalCount"
    Write-Host "  Total in database:   $dbEmailCount"
    Write-Host "  Attachments in DB:   $dbAttachCount"
    Write-Host "  Database file:       $($script:DatabasePath)"
    if ($downloadAttachments) {
        Write-Host "  Attachment folder:   $($script:AttachmentPath)"
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
        "3" { Invoke-Sync -Mode "incremental+attachments" }
        "Q" { $running = $false }
        "q" { $running = $false }
        default {
            Write-Host "Invalid option. Please try again." -ForegroundColor Red
        }
    }
}

Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph. Goodbye!" -ForegroundColor Cyan
