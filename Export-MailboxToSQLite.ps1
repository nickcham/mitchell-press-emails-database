<#
.SYNOPSIS
    Exports all emails from an M365 mailbox (Inbox + Sent Items) to a SQLite database.

.DESCRIPTION
    Connects to Microsoft Graph API using app or delegated authentication,
    fetches every email from Inbox and Sent Items, and stores them in a SQLite database.
    The Message ID is used as the primary key.

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

.EXAMPLE
    # Interactive login (delegated permissions)
    .\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db" -ClientId "your-client-id" -TenantId "your-tenant-id"

    # App-only (client credentials)
    .\Export-MailboxToSQLite.ps1 -DatabasePath ".\emails.db" -ClientId "your-client-id" -TenantId "your-tenant-id" -ClientSecret "your-secret" -UserEmail "user@domain.com"
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
    [string]$UserEmail
)

$ErrorActionPreference = "Stop"

# ---------------------------------------------------------------------------
# 1. Install / Import dependencies
# ---------------------------------------------------------------------------
$requiredModules = @("Microsoft.Graph.Authentication", "PSSQLite")

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "Installing module: $mod ..."
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $mod -Force
}

# ---------------------------------------------------------------------------
# 2. Authenticate to Microsoft Graph
# ---------------------------------------------------------------------------
if ($ClientSecret) {
    # App-only authentication (client credentials flow)
    if (-not $UserEmail) {
        throw "UserEmail is required when using app-only (ClientSecret) authentication."
    }
    $secureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($ClientId, $secureSecret)
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $credential -NoWelcome
    $basePath = "users/$UserEmail"
}
else {
    # Delegated (interactive) authentication
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -Scopes "Mail.Read", "Mail.ReadWrite" -NoWelcome
    $basePath = "me"
}

Write-Host "Authenticated to Microsoft Graph."

# ---------------------------------------------------------------------------
# 3. Create SQLite database and table
# ---------------------------------------------------------------------------
$createTableSQL = @"
CREATE TABLE IF NOT EXISTS emails (
    message_id        TEXT PRIMARY KEY,
    conversation_id   TEXT,
    subject           TEXT,
    from_name         TEXT,
    from_address      TEXT,
    to_recipients     TEXT,
    cc_recipients     TEXT,
    bcc_recipients    TEXT,
    reply_to          TEXT,
    sent_datetime     TEXT,
    received_datetime TEXT,
    has_attachments   INTEGER,
    importance        TEXT,
    is_read           INTEGER,
    is_draft          INTEGER,
    body_content_type TEXT,
    body_content      TEXT,
    body_preview      TEXT,
    web_link          TEXT,
    folder            TEXT,
    categories        TEXT,
    internet_message_id TEXT,
    parent_folder_id  TEXT,
    created_datetime  TEXT,
    last_modified     TEXT
);
"@

Invoke-SqliteQuery -DataSource $DatabasePath -Query $createTableSQL
Write-Host "Database ready: $DatabasePath"

# ---------------------------------------------------------------------------
# 4. Helper: format recipients list -> semicolon-separated string
# ---------------------------------------------------------------------------
function Format-Recipients {
    param([object[]]$Recipients)
    if (-not $Recipients) { return "" }
    return ($Recipients | ForEach-Object {
        "$($_.emailAddress.name) <$($_.emailAddress.address)>"
    }) -join "; "
}

# ---------------------------------------------------------------------------
# 5. Helper: fetch all emails from a folder with paging
# ---------------------------------------------------------------------------
function Get-AllEmails {
    param(
        [string]$FolderName,
        [string]$BasePath
    )

    $folderMap = @{
        "Inbox"      = "Inbox"
        "SentItems"  = "SentItems"
    }

    $graphFolder = $folderMap[$FolderName]
    $url = "https://graph.microsoft.com/v1.0/$BasePath/mailFolders/$graphFolder/messages"
    $url += '?$top=100&$select=id,conversationId,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,sentDateTime,receivedDateTime,hasAttachments,importance,isRead,isDraft,body,bodyPreview,webLink,categories,internetMessageId,parentFolderId,createdDateTime,lastModifiedDateTime'

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

# ---------------------------------------------------------------------------
# 6. Helper: upsert a single email into SQLite
# ---------------------------------------------------------------------------
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

# ---------------------------------------------------------------------------
# 7. Main: fetch and store emails
# ---------------------------------------------------------------------------
$folders = @("Inbox", "SentItems")
$totalCount = 0

foreach ($folder in $folders) {
    Write-Host "`nProcessing folder: $folder"
    $emails = Get-AllEmails -FolderName $folder -BasePath $basePath

    Write-Host "  Found $($emails.Count) emails in $folder. Saving to database..."

    $i = 0
    foreach ($mail in $emails) {
        $i++
        if ($i % 100 -eq 0) {
            Write-Host "    Saved $i / $($emails.Count) ..."
        }
        Save-Email -Mail $mail -Folder $folder -DbPath $DatabasePath
    }

    $totalCount += $emails.Count
    Write-Host "  Done with $folder ($($emails.Count) emails)."
}

# ---------------------------------------------------------------------------
# 8. Summary
# ---------------------------------------------------------------------------
$dbCount = (Invoke-SqliteQuery -DataSource $DatabasePath -Query "SELECT COUNT(*) AS cnt FROM emails").cnt
Write-Host "`n=== Export Complete ==="
Write-Host "Total emails fetched: $totalCount"
Write-Host "Total rows in database: $dbCount"
Write-Host "Database file: $DatabasePath"

Disconnect-MgGraph | Out-Null
Write-Host "Disconnected from Microsoft Graph."
