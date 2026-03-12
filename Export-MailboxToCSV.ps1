<#
.SYNOPSIS
    M365 Mailbox Email Exporter — exports to CSV files. No external modules needed (except Graph).

.DESCRIPTION
    Connects to Microsoft Graph via interactive browser login (supports MFA),
    exports emails from Inbox and Sent Items to per-folder CSV files.
    Auto-creates the output folder (default: .\data\).

    Output files:
      .\data\emails_inbox.csv       — all Inbox emails
      .\data\emails_sent.csv        — all Sent Items emails
      .\data\conversations.csv      — threaded conversations for AI/KB
      .\data\sync_log.csv           — sync history

    Menu options:
    1) Full Export       — download all emails (overwrites existing CSVs)
    2) Incremental Sync  — append only new/changed emails
    3) Download Missing Attachments
    4) Build Conversations (Full)
    5) Build Conversations (Incremental)
    6) Status & Quick Notes
    Q) Quit

.PARAMETER OutputPath
    Folder for CSV output. Defaults to .\data

.PARAMETER AttachmentPath
    Folder where attachments are saved. Defaults to .\Attachments

.PARAMETER ClientId
    (Optional) Azure AD App Registration Client ID for app-only auth.

.PARAMETER TenantId
    (Optional) Azure AD Tenant ID for app-only auth.

.PARAMETER ClientSecret
    (Optional) Client Secret for app-only auth.

.PARAMETER UserEmail
    (Optional) Target mailbox email. Required for app-only auth.

.EXAMPLE
    .\Export-MailboxToCSV.ps1

.EXAMPLE
    .\Export-MailboxToCSV.ps1 -OutputPath ".\my-export"

.EXAMPLE
    .\Export-MailboxToCSV.ps1 -ClientId "abc" -TenantId "xyz" -ClientSecret "secret" -UserEmail "user@domain.com"
#>

param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\data",

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

# Auto-create output folder
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    Write-Host "Created folder: $OutputPath" -ForegroundColor Green
}

# CSV file paths
$script:InboxCsv         = Join-Path $OutputPath "emails_inbox.csv"
$script:SentCsv          = Join-Path $OutputPath "emails_sent.csv"
$script:ConversationsCsv = Join-Path $OutputPath "conversations.csv"
$script:SyncLogCsv       = Join-Path $OutputPath "sync_log.csv"

# ===================================================================
# DEPENDENCIES — only Graph auth needed, no SQLite modules
# ===================================================================
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication")) {
    Write-Host "Installing Microsoft.Graph.Authentication ..." -ForegroundColor Yellow
    Install-Module -Name "Microsoft.Graph.Authentication" -Scope CurrentUser -Force -AllowClobber
}
Import-Module "Microsoft.Graph.Authentication" -Force

# ===================================================================
# MENU
# ===================================================================
function Show-Menu {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "   M365 Mailbox Email Exporter (CSV)" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  [1] Full Export" -ForegroundColor White
    Write-Host "      Download ALL emails (overwrites existing CSVs)" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [2] Incremental Sync" -ForegroundColor White
    Write-Host "      Download only new/changed since last run" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [3] Download Missing Attachments" -ForegroundColor White
    Write-Host "      Scan CSVs and download missing attachments" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [4] Build Conversations (Full)" -ForegroundColor White
    Write-Host "      Rebuild conversations.csv from all emails" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [5] Build Conversations (Incremental)" -ForegroundColor White
    Write-Host "      Update changed conversations only" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  [6] Status & Quick Notes" -ForegroundColor Green
    Write-Host "      Show stats, sync history, custom app guide" -ForegroundColor DarkGray
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
            throw "ClientId, TenantId, UserEmail required for app-only auth."
        }
        $secureSecret = ConvertTo-SecureString $script:ClientSecret -AsPlainText -Force
        $credential = New-Object System.Management.Automation.PSCredential($script:ClientId, $secureSecret)
        Connect-MgGraph -TenantId $script:TenantId -ClientSecretCredential $credential -NoWelcome
        $script:basePath = "users/$($script:UserEmail)"
        Write-Host "Authenticated (app-only) for $($script:UserEmail)." -ForegroundColor Green
    }
    else {
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
# HELPERS
# ===================================================================
function Format-Recipients {
    param([object[]]$Recipients)
    if (-not $Recipients) { return "" }
    return ($Recipients | ForEach-Object {
        "$($_.emailAddress.name) <$($_.emailAddress.address)>"
    }) -join "; "
}

function Get-CsvPath {
    param([string]$Folder)
    if ($Folder -eq "Inbox") { return $script:InboxCsv }
    return $script:SentCsv
}

function Get-ExistingIds {
    param([string]$CsvPath)
    if (-not (Test-Path $CsvPath)) { return @{} }
    $ids = @{}
    Import-Csv $CsvPath | ForEach-Object { $ids[$_.message_id] = $true }
    return $ids
}

function Get-LastSyncTime {
    param([string]$Folder)
    if (-not (Test-Path $script:SyncLogCsv)) { return $null }
    $logs = @(Import-Csv $script:SyncLogCsv | Where-Object { $_.folder -eq $Folder -and $_.sync_type -in @('full','incremental') })
    if ($logs.Count -eq 0) { return $null }
    return ($logs | Sort-Object completed_at -Descending | Select-Object -First 1).completed_at
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
# FETCH + WRITE PAGE-BY-PAGE (writes to disk after every 100 emails)
# ===================================================================
function Invoke-FetchAndWrite {
    param(
        [string]$FolderName,
        [string]$BasePath,
        [string]$SinceDateTime,
        [string]$CsvPath,
        [string]$Mode
    )

    $fields = 'id,conversationId,subject,from,toRecipients,ccRecipients,bccRecipients,replyTo,sentDateTime,receivedDateTime,hasAttachments,importance,isRead,isDraft,body,bodyPreview,uniqueBody,webLink,categories,internetMessageId,parentFolderId,createdDateTime,lastModifiedDateTime'

    $url = "https://graph.microsoft.com/v1.0/$BasePath/mailFolders/$FolderName/messages"
    $url += "?`$top=100&`$select=$fields"

    if ($SinceDateTime) {
        $url += "&`$filter=lastModifiedDateTime ge $SinceDateTime"
        Write-Host "  Incremental filter: modified since $SinceDateTime" -ForegroundColor DarkYellow
    }

    $url += "&`$orderby=lastModifiedDateTime asc"

    # For incremental mode, load existing rows into a hashtable for merge
    $existingById = @{}
    if ($Mode -ne "full" -and (Test-Path $CsvPath)) {
        $existingRows = @(Import-Csv $CsvPath)
        foreach ($r in $existingRows) { $existingById[$r.message_id] = $r }
        Write-Host "  Loaded $($existingById.Count) existing rows for merge." -ForegroundColor DarkGray
    }

    # For full mode, delete existing file so first page creates fresh
    if ($Mode -eq "full" -and (Test-Path $CsvPath)) {
        Remove-Item $CsvPath -Force
    }

    $pageCount = 0
    $totalWritten = 0

    while ($url) {
        $pageCount++
        Write-Host "  Fetching $FolderName page $pageCount ..." -ForegroundColor Cyan

        $messages = $null
        try {
            $response = Invoke-MgGraphRequest -Method GET -Uri $url -ErrorAction Stop
            $messages = $response.value
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

        if (-not $messages -or $messages.Count -eq 0) {
            Start-Sleep -Milliseconds (Get-Random -Minimum 700 -Maximum 1500)
            continue
        }

        # Convert this page to row objects
        $pageRows = @()
        foreach ($mail in $messages) {
            $pageRows += Convert-MailToRow -Mail $mail -Folder $FolderName
        }

        if ($Mode -eq "full") {
            # Full: append each page directly to CSV
            $isFirstPage = -not (Test-Path $CsvPath)
            if ($isFirstPage) {
                $pageRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
            } else {
                $pageRows | Export-Csv -Path $CsvPath -Append -NoTypeInformation -Encoding UTF8
            }
        }
        else {
            # Incremental: update the in-memory hashtable, then flush to disk
            foreach ($r in $pageRows) {
                if ($existingById.ContainsKey($r.message_id)) {
                    # Preserve attachments_downloaded flag
                    $r.attachments_downloaded = $existingById[$r.message_id].attachments_downloaded
                }
                $existingById[$r.message_id] = $r
            }
            # Write full merged set to disk after each page
            $existingById.Values | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        }

        $totalWritten += $pageRows.Count
        $fileSize = if (Test-Path $CsvPath) { [math]::Round((Get-Item $CsvPath).Length / 1KB, 1) } else { 0 }
        Write-Host "    Wrote page $pageCount ($($pageRows.Count) emails) -> $CsvPath ($($fileSize) KB on disk, $totalWritten total)" -ForegroundColor Green

        Start-Sleep -Milliseconds (Get-Random -Minimum 700 -Maximum 1500)
    }

    return $totalWritten
}

# ===================================================================
# CONVERT MAIL TO ROW OBJECT
# ===================================================================
function Convert-MailToRow {
    param([object]$Mail, [string]$Folder)

    $fromName    = if ($Mail.from) { $Mail.from.emailAddress.name } else { "" }
    $fromAddress = if ($Mail.from) { $Mail.from.emailAddress.address } else { "" }

    $bodyCont = if ($Mail.uniqueBody.content) { $Mail.uniqueBody.content } else { $Mail.body.content }
    $bodyType = if ($Mail.uniqueBody) { $Mail.uniqueBody.contentType } else { $Mail.body.contentType }

    $cleanBody = ConvertFrom-Html -Html $bodyCont
    $cleanBody = Remove-QuotedContent -Body $cleanBody

    # Sanitize all text fields for CSV safety — replace newlines with literal \n
    # so each email stays on exactly one CSV row. Raw HTML body is dropped entirely
    # (it destroys CSV structure). Use cleaned_body for AI, Graph API for originals.
    $safeCleanBody  = if ($cleanBody)          { $cleanBody -replace "`r`n", '\n' -replace "`n", '\n' -replace "`r", '\n' } else { "" }
    $safePreview    = if ($Mail.bodyPreview)    { $Mail.bodyPreview -replace "`r`n", '\n' -replace "`n", '\n' -replace "`r", '\n' } else { "" }
    $safeSubject    = if ($Mail.subject)        { $Mail.subject -replace "`r`n", ' ' -replace "`n", ' ' -replace "`r", ' ' } else { "" }

    return [PSCustomObject]@{
        message_id             = $Mail.id
        conversation_id        = $Mail.conversationId
        subject                = $safeSubject
        from_name              = $fromName
        from_address           = $fromAddress
        to_recipients          = (Format-Recipients $Mail.toRecipients)
        cc_recipients          = (Format-Recipients $Mail.ccRecipients)
        bcc_recipients         = (Format-Recipients $Mail.bccRecipients)
        reply_to               = (Format-Recipients $Mail.replyTo)
        sent_datetime          = $Mail.sentDateTime
        received_datetime      = $Mail.receivedDateTime
        has_attachments        = [int]$Mail.hasAttachments
        attachments_downloaded = 0
        importance             = $Mail.importance
        is_read                = [int]$Mail.isRead
        is_draft               = [int]$Mail.isDraft
        body_content_type      = $bodyType
        body_preview           = $safePreview
        cleaned_body           = $safeCleanBody
        web_link               = $Mail.webLink
        folder                 = $Folder
        categories             = ($Mail.categories -join "; ")
        internet_message_id    = $Mail.internetMessageId
        parent_folder_id       = $Mail.parentFolderId
        created_datetime       = $Mail.createdDateTime
        last_modified          = $Mail.lastModifiedDateTime
    }
}

# ===================================================================
# EMAIL SYNC ENGINE — writes to disk after every page, not at the end
# ===================================================================
function Invoke-Sync {
    param([string]$Mode)

    $folders = @("Inbox", "SentItems")
    $totalCount = 0

    foreach ($folder in $folders) {
        Write-Host "`n--- Processing: $folder ---" -ForegroundColor Cyan
        $syncStart = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        $csvPath = Get-CsvPath -Folder $folder

        $sinceDateTime = $null
        if ($Mode -ne "full") {
            $sinceDateTime = Get-LastSyncTime -Folder $folder
        }

        # Fetch and write page-by-page — data hits disk after every 100 emails
        $count = Invoke-FetchAndWrite -FolderName $folder -BasePath $script:basePath `
            -SinceDateTime $sinceDateTime -CsvPath $csvPath -Mode $Mode

        # Log the sync
        $syncEnd = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        [PSCustomObject]@{
            sync_type     = $Mode
            folder        = $folder
            started_at    = $syncStart
            completed_at  = $syncEnd
            emails_synced = $count
        } | Export-Csv -Path $script:SyncLogCsv -Append -NoTypeInformation -Encoding UTF8

        $totalCount += $count
        Write-Host "  Done: $folder ($count emails)" -ForegroundColor Green
    }

    # Summary
    $inboxSize = if (Test-Path $script:InboxCsv) { [math]::Round((Get-Item $script:InboxCsv).Length / 1KB, 1) } else { 0 }
    $sentSize  = if (Test-Path $script:SentCsv)  { [math]::Round((Get-Item $script:SentCsv).Length / 1KB, 1)  } else { 0 }

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   SYNC COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Mode:              $Mode"
    Write-Host "  Emails processed:  $totalCount"
    Write-Host "  Inbox CSV:         $($script:InboxCsv) ($inboxSize KB)"
    Write-Host "  Sent CSV:          $($script:SentCsv) ($sentSize KB)"
    Write-Host "  Output folder:     $($script:OutputPath)"
    Write-Host ""
}

# ===================================================================
# DOWNLOAD MISSING ATTACHMENTS
# ===================================================================
function Invoke-DownloadMissingAttachments {
    # Collect pending from both CSVs
    $pending = @()
    foreach ($csvPath in @($script:InboxCsv, $script:SentCsv)) {
        if (-not (Test-Path $csvPath)) { continue }
        $pending += @(Import-Csv $csvPath | Where-Object { $_.has_attachments -eq "1" -and $_.attachments_downloaded -eq "0" })
    }

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
    $completedIds = @{}

    foreach ($row in $pending) {
        $downloaded++
        $msgId = $row.message_id
        if ($downloaded % 10 -eq 0 -or $downloaded -eq $pending.Count) {
            Write-Host "  Processing $downloaded / $($pending.Count) ..." -ForegroundColor DarkGray
        }

        try {
            $url = "https://graph.microsoft.com/v1.0/$($script:basePath)/messages/$msgId/attachments"
            $response = Invoke-MgGraphRequest -Method GET -Uri $url

            foreach ($att in $response.value) {
                if ($att.'@odata.type' -eq '#microsoft.graph.referenceAttachment') { continue }
                if (-not $att.contentBytes) { continue }

                $msgIdShort   = ($msgId -replace '[^a-zA-Z0-9]', '')[0..7] -join ''
                $safeAttId    = ($att.id -replace '[^a-zA-Z0-9]', '')[0..15] -join ''
                $safeFilename = $att.name -replace '[\\/:*?"<>|]', '_'
                $diskFilename = "${msgIdShort}_${safeAttId}_${safeFilename}"
                $diskFullPath = Join-Path $script:AttachmentPath $diskFilename

                $bytes = [Convert]::FromBase64String($att.contentBytes)
                [System.IO.File]::WriteAllBytes($diskFullPath, $bytes)
                $fileCount++
            }

            $completedIds[$msgId] = $true

        } catch {
            $errors++
            Write-Host "    ERROR on '$($row.subject)': $($_.Exception.Message)" -ForegroundColor Red
        }

        Start-Sleep -Milliseconds (Get-Random -Minimum 800 -Maximum 1800)
    }

    # Update CSVs to mark attachments as downloaded
    if ($completedIds.Count -gt 0) {
        foreach ($csvPath in @($script:InboxCsv, $script:SentCsv)) {
            if (-not (Test-Path $csvPath)) { continue }
            $allRows = @(Import-Csv $csvPath)
            foreach ($r in $allRows) {
                if ($completedIds.ContainsKey($r.message_id)) {
                    $r.attachments_downloaded = "1"
                }
            }
            $allRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        }
    }

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   ATTACHMENT DOWNLOAD COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Emails processed:  $downloaded"
    Write-Host "  Files saved:       $fileCount"
    Write-Host "  Errors:            $errors"
    Write-Host "  Attachment folder: $($script:AttachmentPath)"
    Write-Host ""
}

# ===================================================================
# BUILD CONVERSATIONS
# ===================================================================
function Invoke-BuildConversations {
    param([string]$Mode)

    Write-Host "`nBuilding conversations..." -ForegroundColor Cyan

    # Load all emails from both CSVs
    $allEmails = @()
    foreach ($csvPath in @($script:InboxCsv, $script:SentCsv)) {
        if (Test-Path $csvPath) { $allEmails += @(Import-Csv $csvPath) }
    }

    if ($allEmails.Count -eq 0) {
        Write-Host "No emails found. Run a sync first!" -ForegroundColor Yellow
        return
    }

    # Group by conversation_id
    $groups = $allEmails | Where-Object { $_.conversation_id } | Group-Object conversation_id

    # For incremental, load existing conversations and only rebuild changed ones
    $existingConvs = @{}
    if ($Mode -eq "incremental" -and (Test-Path $script:ConversationsCsv)) {
        $existing = @(Import-Csv $script:ConversationsCsv)
        foreach ($c in $existing) { $existingConvs[$c.conversation_id] = $c }
    }

    $now = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    $conversations = @()
    $built = 0

    foreach ($group in $groups) {
        $convId = $group.Name
        $emails = $group.Group | Sort-Object sent_datetime

        # Incremental: skip if no email is newer than last_built
        if ($Mode -eq "incremental" -and $existingConvs.ContainsKey($convId)) {
            $lastBuilt = $existingConvs[$convId].last_built
            $newest = ($emails | Sort-Object last_modified -Descending | Select-Object -First 1).last_modified
            if ($newest -le $lastBuilt) {
                $conversations += $existingConvs[$convId]
                continue
            }
        }

        $built++

        # Subject from first email
        $subject = $emails[0].subject

        # Unique participants
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

        # Build thread text (using \n literals so CSV stays one-row-per-conversation)
        $threadParts = @()
        foreach ($e in $emails) {
            # cleaned_body already has \n literals from email CSV; body_preview may have real newlines
            $body = if ($e.cleaned_body) { $e.cleaned_body } else {
                if ($e.body_preview) { $e.body_preview -replace "`r`n", '\n' -replace "`n", '\n' -replace "`r", '\n' } else { "" }
            }

            $header = "--- [" + $e.sent_datetime + "] From: " + $e.from_name + " <" + $e.from_address + "> ---"
            $toLine = "To: " + $e.to_recipients
            $ccLine = if ($e.cc_recipients) { '\nCC: ' + $e.cc_recipients } else { "" }
            $threadParts += $header + '\n' + $toLine + $ccLine + '\n' + $body
        }
        $fullThread = $threadParts -join '\n\n'

        $hasAtt = [int](($emails | Where-Object { $_.has_attachments -eq "1" }).Count -gt 0)
        $outlookLink = ($emails | Where-Object { $_.web_link } | Select-Object -Last 1).web_link

        $conversations += [PSCustomObject]@{
            conversation_id        = $convId
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

        if ($built % 50 -eq 0) {
            Write-Host "    Built $built conversations ..." -ForegroundColor DarkGray
        }
    }

    $conversations | Export-Csv -Path $script:ConversationsCsv -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "   CONVERSATIONS BUILD COMPLETE" -ForegroundColor Green
    Write-Host "============================================" -ForegroundColor Green
    Write-Host "  Mode:                $Mode"
    Write-Host "  Conversations built: $built"
    Write-Host "  Total:               $($conversations.Count)"
    Write-Host "  File:                $($script:ConversationsCsv)"
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

    $lastInbox = Get-LastSyncTime -Folder "Inbox"
    $lastSent  = Get-LastSyncTime -Folder "SentItems"
    Write-Host "Last sync times:" -ForegroundColor White
    Write-Host "  Inbox: $($lastInbox ?? 'Never yet')" -ForegroundColor DarkGray
    Write-Host "  Sent:  $($lastSent ?? 'Never yet')" -ForegroundColor DarkGray
    Write-Host ""

    $inboxCount = if (Test-Path $script:InboxCsv) { @(Import-Csv $script:InboxCsv).Count } else { 0 }
    $sentCount  = if (Test-Path $script:SentCsv)  { @(Import-Csv $script:SentCsv).Count  } else { 0 }
    $convCount  = if (Test-Path $script:ConversationsCsv) { @(Import-Csv $script:ConversationsCsv).Count } else { 0 }
    Write-Host "Inbox emails:    $inboxCount" -ForegroundColor White
    Write-Host "Sent emails:     $sentCount" -ForegroundColor White
    Write-Host "Conversations:   $convCount" -ForegroundColor White
    Write-Host ""

    Write-Host "Quick Notes:" -ForegroundColor Yellow
    Write-Host "  - 429 throttling is handled automatically (Retry-After header)" -ForegroundColor DarkGray
    Write-Host "  - CSVs open directly in Excel" -ForegroundColor DarkGray
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
