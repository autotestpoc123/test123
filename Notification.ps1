# Notification.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

<#
.SYNOPSIS
English code-review note for function 'Get-Cert'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-Cert {
    param(
        [Parameter(Mandatory = $true)]
        [string]$KeyName
    )

    Add-Type -AssemblyName System.Security
    $certStore = New-Object System.Security.Cryptography.X509Certificates.X509Store('My', 'LocalMachine')
    $certStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    try {
        $certs = $certStore.Certificates |
            Where-Object { $_.Subject -like "*CN=$KeyName*" } |
            Sort-Object NotAfter -Descending
        return ($certs | Select-Object -First 1)
    }
    finally {
        $certStore.Close()
    }
}

<#
.SYNOPSIS
English code-review note for function 'Get-VaultSecret'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-VaultSecret {
    param(
        [Parameter(Mandatory = $true)]
        [string]$VaultServer,
        [Parameter(Mandatory = $true)]
        [string]$VaultEnv,
        [Parameter(Mandatory = $true)]
        [string]$KeyName,
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$SysIdCert,
        [string]$Eonid = '309843'
    )

    $authHeader = @{
        'X-Vault-Namespace' = 'msms/core'
    }
    $certAuthUrl = "$VaultServer/v1/auth/cert/login"
    $keyPathUrl = "$VaultServer/v1/msa/data/secret/$Eonid/$VaultEnv/$KeyName"

    try {
        $certAuthResponse = Invoke-RestMethod -Uri $certAuthUrl -Certificate $SysIdCert -Method Post -Headers $authHeader -UseBasicParsing
    }
    catch {
        throw "Failed to get vault token: $($_.Exception.Message)"
    }

    $vaultClientToken = $certAuthResponse.auth.client_token
    if (-not $vaultClientToken) {
        throw 'Vault authentication succeeded but no client token was returned.'
    }

    $authHeader['X-Vault-Token'] = $vaultClientToken

    try {
        $keyRequestResponse = Invoke-WebRequest -Uri $keyPathUrl -Method Get -Headers $authHeader -UseBasicParsing
    }
    catch {
        throw "Failed to get secret for ${KeyName}: $($_.Exception.Message)"
    }

    if (-not $keyRequestResponse.Content) {
        throw "The secret response for $KeyName was empty."
    }

    $secret = ($keyRequestResponse.Content | ConvertFrom-Json).data.data.$KeyName
    if (-not $secret) {
        throw "The secret value for $KeyName was null."
    }

    return $secret
}

<#
.SYNOPSIS
English code-review note for function 'Send-Mail'.
.DESCRIPTION
Sends notifications using configured transport and security settings.
#>
function Send-Mail {
    param (
        [Parameter(Mandatory = $true)]
        [string]$From,
        [Parameter(Mandatory = $true)]
        [string[]]$To,
        [string]$Cc,
        [Parameter(Mandatory = $true)]
        [string]$Subject,
        [Parameter(Mandatory = $true)]
        [string]$Body,
        [Parameter(Mandatory = $true)]
        [string]$SmtpServer,
        [Parameter(Mandatory = $true)]
        [string]$KeyName,
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert,
        [int]$Port = 2587,
        [string]$LogFilePath
    )

    if (-not $Cert) {
        throw "Certificate not found in LocalMachine\My store for $KeyName."
    }

    $mail = $null
    try {
        $mail = New-Object System.Net.Mail.MailMessage
        $mail.From = $From
        foreach ($recipient in $To) {
            $mail.To.Add($recipient)
        }
        if ($Cc) { $mail.CC.Add($Cc) }
        $mail.Subject = $Subject
        $mail.Body = $Body
        $mail.IsBodyHtml = $true

        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
        $smtp.EnableSsl = $true
        # X509CertificateCollection.Add returns the inserted index. Suppress it:
        # leaking that integer onto the success pipeline makes callers receive
        # an array instead of the single Send-AuditBuMail result object.
        [void]$smtp.ClientCertificates.Add([System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Cert))
        $smtp.Send($mail)
        if ($LogFilePath) {
            Write-Log -LogString "Email sent successfully to $($To -join ', ')" -LogFilePath $LogFilePath
        }
        Write-Host "Email sent successfully to $($To -join ', ')"
    }
    catch {
        if ($LogFilePath) {
            Write-Log -LogString "Failed to send email to $($To -join ', '): $($_.Exception.Message)" -LogFilePath $LogFilePath
        }
        throw "Failed to send email: $($_.Exception.Message)"
    }
    finally {
        if ($mail) {
            $mail.Dispose()
        }
    }
}

<#
.SYNOPSIS
English code-review note for function 'New-HtmlBody'.
.DESCRIPTION
Creates a new object or structure used by subsequent processing steps.
#>
function New-HtmlBody {
    param(
        [string]$TableHtml = '',
        [string]$ViolationContent,
        [string]$NoViolationContent,
        [bool]$HasViolation = $false
    )

    if ($HasViolation) {
        $violationParaHtml = '<p>{{ViolationContent}}</p>'
        return $htmlTemplateNew.Replace('{{ViolationParagraph}}', $violationParaHtml).
            Replace('{{TableSection}}', $TableHtml).
            Replace('{{NoViolationParagraph}}', '').
            Replace('{{ViolationContent}}', $ViolationContent)
    }

    $noViolationParaHtml = '<p>{{NoViolationContent}}</p>'
    return $htmlTemplateNew.Replace('{{ViolationParagraph}}', '').
        Replace('{{TableSection}}', '').
        Replace('{{NoViolationParagraph}}', $noViolationParaHtml).
        Replace('{{NoViolationContent}}', $NoViolationContent)
}

function Send-PreflightNotification {
    param(
        [Parameter(Mandatory)]
        [object]$NotificationConfig,
        [object[]]$MissingItems = @(),
        [object[]]$InvalidItems = @(),
        [Parameter(Mandatory)]
        [string]$Phase,
        [string]$StartDate,
        [string]$EndDate,
        [string]$LogFilePath
    )

    if (-not $NotificationConfig -or -not $NotificationConfig.SmtpServer -or -not $NotificationConfig.Cert) {
        throw "Notification config incomplete: missing SmtpServer or Cert."
    }
    if ($NotificationConfig.OpsTeam.Count -eq 0) {
        throw "Notification config has no OpsTeam recipients."
    }
    if ($MissingItems.Count -eq 0 -and $InvalidItems.Count -eq 0) {
        throw "Send-PreflightNotification called with no MissingItems and no InvalidItems."
    }

    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $missingRendered = @(
        $MissingItems | ForEach-Object {
            "<b>$(& $enc $_.Name)</b>: $(& $enc $_.ExpectedPath) <i>[$(& $enc $_.Source)]</i>"
        }
    )
    $invalidRendered = @(
        $InvalidItems | ForEach-Object {
            "<b>$(& $enc $_.Name)</b>: $(& $enc $_.Error) <i>[$(& $enc $_.Source)]</i>"
        }
    )

    $body = Build-AuditNotificationHtml `
        -Heading "WeCom Audit Preflight Failed - Phase $Phase" `
        -Intro "Date range: $(& $enc $StartDate) - $(& $enc $EndDate)" `
        -Sections @(
            [PSCustomObject]@{ Heading = 'Missing Files';  Items = $missingRendered }
            [PSCustomObject]@{ Heading = 'Invalid Items';  Items = $invalidRendered }
        ) `
        -Footer 'Please prepare the required files and re-trigger the scheduled job.'

    $subject = "[WeCom Audit] Preflight Failed - $Phase blocked ($StartDate - $EndDate)"

    $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    $fromAddress = ([string]$NotificationConfig.From).Trim()
    if ($fromAddress -notmatch $emailPattern) {
        throw "Notification 'From' is not a valid email address: '$fromAddress'. Update Notification.<Env>.From in config to a real email (e.g. 'wecom-audit-qa@yourdomain.com')."
    }

    $validTo = @(
        $NotificationConfig.OpsTeam |
            Where-Object { $_ } |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ -match $emailPattern }
    )
    if ($validTo.Count -eq 0) {
        throw "Notification 'OpsTeam' has no valid email recipients. Configured: $($NotificationConfig.OpsTeam -join ', ')"
    }

    Send-Mail `
        -From $fromAddress `
        -To $validTo `
        -Subject $subject `
        -Body $body `
        -SmtpServer $NotificationConfig.SmtpServer `
        -KeyName $NotificationConfig.CertName `
        -Cert $NotificationConfig.Cert `
        -Port $NotificationConfig.Port `
        -LogFilePath $LogFilePath
}

function Build-AuditNotificationHtml {
    # Contract: Heading / Footer / Sections[].Heading are plain text - this helper
    # HtmlEncodes them. Intro and Sections[].Items are trusted HTML - caller is
    # responsible for encoding any untrusted content inside them.
    param(
        [Parameter(Mandatory)]
        [string]$Heading,
        [string]$Intro,
        [object[]]$Sections = @(),
        [string]$Footer = 'Please follow up and re-trigger the scheduled job once resolved.'
    )

    $encode = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add("<h3>$(& $encode $Heading)</h3>")
    if ($Intro) { $lines.Add("<p>$Intro</p>") }

    foreach ($section in $Sections) {
        $items = @($section.Items)
        if ($items.Count -eq 0) { continue }
        $title = [string]$section.Heading
        $lines.Add("<h4>$(& $encode $title) ($($items.Count))</h4><ul>")
        foreach ($item in $items) {
            $lines.Add("<li>$item</li>")
        }
        $lines.Add("</ul>")
    }

    if ($Footer) { $lines.Add("<p>$(& $encode $Footer)</p>") }
    return ($lines -join "`n")
}

<#
.SYNOPSIS
Sends an HTML email notifying ops that validation failed for an audit run.
.DESCRIPTION
Builds an HTML body grouping missing files into Static, Dynamic, and Unknown
sections (Dynamic entries show their producing task), plus an Unexpected Files
section. Subject is "[WeCom Audit][<ENV>] Validation Failed - <RunId>". From,
OpsTeam are validated against a basic email regex before send;
invalid entries are dropped. Throws if NotificationConfig is incomplete, From is
not a valid address, OpsTeam yields no valid recipients, or both MissingFiles
and UnexpectedFiles are empty.
.PARAMETER NotificationConfig
Resolved notification config (from Resolve-NotificationConfig) - must carry
SmtpServer, Cert, From, OpsTeam, CertName, Port.
.PARAMETER Environment
Environment tag (PROD / QA) shown in the subject line.
.PARAMETER RunId
The run identifier whose validation failed.
.PARAMETER StartDate
Cycle start date (yyyyMMdd) shown in the subject and intro.
.PARAMETER EndDate
Cycle end date (yyyyMMdd) shown in the subject and intro.
.PARAMETER MissingFiles
Missing-file entries from the validation summary. Objects with Name/Source/
ProducedBy are grouped by Source; legacy plain strings are listed under
"Missing Files (unknown)".
.PARAMETER UnexpectedFiles
File names that exist in the source folder but are not in the expected manifest.
.PARAMETER ValidationFolder
Optional - folder actually validated (shown in the intro for triage).
.PARAMETER ValidationReportPath
Optional - path to backup-folder-validation.json.
.PARAMETER SummaryPath
Optional - path to backup-validation-summary.json.
.PARAMETER LogFilePath
Optional log file forwarded to Send-Mail.
.EXAMPLE
Send-ValidationFailureNotification -NotificationConfig $cfg -Environment 'QA' `
    -RunId '20260520_142743' -StartDate '20260506' -EndDate '20260520' `
    -MissingFiles $summary.MissingFiles -UnexpectedFiles $summary.UnexpectedFiles
.NOTES
Dispatched by Invoke-WeComAuditScheduler when AuditValidate exits with code 1
(via Send-ValidationFailureNotificationFromSummary).
#>
function Send-ValidationFailureNotification {
    param(
        [Parameter(Mandatory)]
        [object]$NotificationConfig,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$StartDate,
        [Parameter(Mandatory)]
        [string]$EndDate,
        [object[]]$MissingFiles = @(),
        [string[]]$UnexpectedFiles = @(),
        [string]$ValidationFolder,
        [string]$ValidationReportPath,
        [string]$SummaryPath,
        [string]$LogFilePath
    )

    if (-not $NotificationConfig -or -not $NotificationConfig.SmtpServer -or -not $NotificationConfig.Cert) {
        throw "Notification config incomplete: missing SmtpServer or Cert."
    }
    if ($NotificationConfig.OpsTeam.Count -eq 0) {
        throw "Notification config has no OpsTeam recipients."
    }
    if ($MissingFiles.Count -eq 0 -and $UnexpectedFiles.Count -eq 0) {
        throw "Send-ValidationFailureNotification called with no MissingFiles and no UnexpectedFiles."
    }

    $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    $fromAddress = ([string]$NotificationConfig.From).Trim()
    if ($fromAddress -notmatch $emailPattern) {
        throw "Notification 'From' is not a valid email address: '$fromAddress'."
    }

    $validTo = @(
        $NotificationConfig.OpsTeam |
            Where-Object { $_ } |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ -match $emailPattern }
    )
    if ($validTo.Count -eq 0) {
        throw "Notification 'OpsTeam' has no valid email recipients."
    }

    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $introParts = New-Object 'System.Collections.Generic.List[string]'
    $introParts.Add("Run: $(& $enc $RunId)")
    $introParts.Add("Date range: $(& $enc $StartDate) - $(& $enc $EndDate)")
    if ($ValidationFolder)     { $introParts.Add("Validation folder: $(& $enc $ValidationFolder)") }
    if ($ValidationReportPath) { $introParts.Add("Validation report: $(& $enc $ValidationReportPath)") }
    if ($SummaryPath)          { $introParts.Add("Summary: $(& $enc $SummaryPath)") }
    $intro = ($introParts -join '<br/>')

    $missingNormalized = @(
        $MissingFiles | ForEach-Object {
            if ($_ -is [string]) {
                [PSCustomObject]@{ Name = $_; Source = 'unknown'; ProducedBy = $null }
            }
            else { $_ }
        }
    )

    $staticMissing  = @($missingNormalized | Where-Object { ($_.PSObject.Properties['Source']) -and ($_.Source -eq 'static') })
    $dynamicMissing = @($missingNormalized | Where-Object { ($_.PSObject.Properties['Source']) -and ($_.Source -eq 'dynamic') })
    $otherMissing   = @($missingNormalized | Where-Object {
        $hasSource = $_.PSObject.Properties['Source']
        (-not $hasSource) -or ($_.Source -ne 'static' -and $_.Source -ne 'dynamic')
    })

    $renderMissing = {
        param($entry)
        $namePart = "<b>$(& $enc $entry.Name)</b>"
        if ($entry.PSObject.Properties['ProducedBy'] -and $entry.ProducedBy) {
            $namePart += " <i>(from $(& $enc $entry.ProducedBy))</i>"
        }
        $namePart
    }

    $sections = @(
        [PSCustomObject]@{ Heading = 'Missing Static Files';  Items = @($staticMissing  | ForEach-Object { & $renderMissing $_ }) }
        [PSCustomObject]@{ Heading = 'Missing Dynamic Files'; Items = @($dynamicMissing | ForEach-Object { & $renderMissing $_ }) }
        [PSCustomObject]@{ Heading = 'Missing Files';         Items = @($otherMissing   | ForEach-Object { & $renderMissing $_ }) }
        [PSCustomObject]@{ Heading = 'Unexpected Files';      Items = @($UnexpectedFiles | ForEach-Object { "<b>$(& $enc $_)</b>" }) }
    )

    $body = Build-AuditNotificationHtml `
        -Heading "WeCom Audit Validation Failed - Run $RunId" `
        -Intro $intro `
        -Sections $sections `
        -Footer 'Source folder contents do not match the expected manifest. Resolve the differences and re-run validation.'

    $envTag = ([string]$Environment).ToUpperInvariant()
    $subject = "[WeCom Audit][$envTag] Validation Failed - $RunId ($StartDate - $EndDate)"

    Send-Mail `
        -From $fromAddress `
        -To $validTo `
        -Subject $subject `
        -Body $body `
        -SmtpServer $NotificationConfig.SmtpServer `
        -KeyName $NotificationConfig.CertName `
        -Cert $NotificationConfig.Cert `
        -Port $NotificationConfig.Port `
        -LogFilePath $LogFilePath
}

<#
.SYNOPSIS
Sends an HTML email notifying ops that the archive step failed after a passing
validation.
.DESCRIPTION
Builds an HTML body summarizing the ArchiveResult (Deleted / Failed / Skipped
counts, plus an Aborted reason when present). Footer text is tailored per
ArchiveStatus value (BackupFailed / CleanupAborted / CleanupPartiallyFailed).
Subject is "[WeCom Audit][<ENV>] Archive Failed (<status>) - <RunId>". Same From
and OpsTeam validation rules as Send-ValidationFailureNotification.
.PARAMETER NotificationConfig
Resolved notification config (from Resolve-NotificationConfig) - must carry
SmtpServer, Cert, From, OpsTeam, CertName, Port.
.PARAMETER Environment
Environment tag (PROD / QA) shown in the subject line.
.PARAMETER RunId
The run identifier whose archive step failed.
.PARAMETER StartDate
Cycle start date (yyyyMMdd) shown in the subject and intro.
.PARAMETER EndDate
Cycle end date (yyyyMMdd) shown in the subject and intro.
.PARAMETER ArchiveStatus
ArchiveStatus enum value: BackupFailed / CleanupAborted / CleanupPartiallyFailed
(NoOp / Success / NoSourceFiles / NotAttempted are not failure states and should
not trigger this notification).
.PARAMETER ArchiveResult
Optional ArchiveResult object from the validation summary - properties read are
DeletedCount, FailedCount, SkippedCount, Aborted, AbortReason.
.PARAMETER BackupFolder
Optional - backup folder path shown in the intro for triage.
.PARAMETER SummaryPath
Optional - path to backup-validation-summary.json.
.PARAMETER LogFilePath
Optional log file forwarded to Send-Mail.
.EXAMPLE
Send-ArchiveFailureNotification -NotificationConfig $cfg -Environment 'QA' `
    -RunId '20260520_142743' -StartDate '20260506' -EndDate '20260520' `
    -ArchiveStatus 'CleanupAborted' -ArchiveResult $summary.ArchiveResult
.NOTES
Dispatched by Invoke-WeComAuditScheduler when AuditValidate exits with code 2
(via Send-ArchiveFailureNotificationFromSummary).
#>
function Send-ArchiveFailureNotification {
    param(
        [Parameter(Mandatory)]
        [object]$NotificationConfig,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$StartDate,
        [Parameter(Mandatory)]
        [string]$EndDate,
        [Parameter(Mandatory)]
        [string]$ArchiveStatus,
        [object]$ArchiveResult,
        [string]$BackupFolder,
        [string]$SummaryPath,
        [string]$LogFilePath
    )

    if (-not $NotificationConfig -or -not $NotificationConfig.SmtpServer -or -not $NotificationConfig.Cert) {
        throw "Notification config incomplete: missing SmtpServer or Cert."
    }
    if ($NotificationConfig.OpsTeam.Count -eq 0) {
        throw "Notification config has no OpsTeam recipients."
    }

    $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $fromAddress = ([string]$NotificationConfig.From).Trim()
    if ($fromAddress -notmatch $emailPattern) {
        throw "Notification 'From' is not a valid email address: '$fromAddress'."
    }
    $validTo = @(
        $NotificationConfig.OpsTeam |
            Where-Object { $_ } |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ -match $emailPattern }
    )
    if ($validTo.Count -eq 0) {
        throw "Notification 'OpsTeam' has no valid email recipients."
    }
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $introParts = New-Object 'System.Collections.Generic.List[string]'
    $introParts.Add("Run: $(& $enc $RunId)")
    $introParts.Add("Date range: $(& $enc $StartDate) - $(& $enc $EndDate)")
    $introParts.Add("Archive status: $(& $enc $ArchiveStatus)")
    if ($BackupFolder) { $introParts.Add("Backup folder: $(& $enc $BackupFolder)") }
    if ($SummaryPath)  { $introParts.Add("Summary: $(& $enc $SummaryPath)") }
    $intro = ($introParts -join '<br/>')

    $resultLines = New-Object 'System.Collections.Generic.List[string]'
    if ($ArchiveResult) {
        if ($ArchiveResult.PSObject.Properties['DeletedCount']) { $resultLines.Add("Deleted: $(& $enc $ArchiveResult.DeletedCount)") }
        if ($ArchiveResult.PSObject.Properties['FailedCount'])  { $resultLines.Add("Failed: $(& $enc $ArchiveResult.FailedCount)") }
        if ($ArchiveResult.PSObject.Properties['SkippedCount']) { $resultLines.Add("Skipped: $(& $enc $ArchiveResult.SkippedCount)") }
        if ($ArchiveResult.PSObject.Properties['Aborted'] -and $ArchiveResult.Aborted) {
            $reason = if ($ArchiveResult.PSObject.Properties['AbortReason']) { $ArchiveResult.AbortReason } else { 'unspecified' }
            $resultLines.Add("Aborted: $(& $enc $reason)")
        }
    }

    $sections = @(
        [PSCustomObject]@{ Heading = 'Archive Result'; Items = @($resultLines.ToArray()) }
    )

    $footer = switch ($ArchiveStatus) {
        'BackupFailed'           { 'One or more source files failed to copy. Source files have NOT been deleted. Investigate and re-run validation.' }
        'CleanupAborted'         { 'Source cleanup was aborted by safety check. Files are still in source folder. See log for the abort reason.' }
        'CleanupPartiallyFailed' { 'Some source files could not be deleted. Investigate the residual files and remove manually if appropriate.' }
        default                  { 'Archive step did not complete successfully. Review the summary and validation report.' }
    }

    $body = Build-AuditNotificationHtml `
        -Heading "WeCom Audit Archive Failed - Run $RunId" `
        -Intro $intro `
        -Sections $sections `
        -Footer $footer

    $envTag = ([string]$Environment).ToUpperInvariant()
    $subject = "[WeCom Audit][$envTag] Archive Failed ($ArchiveStatus) - $RunId ($StartDate - $EndDate)"

    Send-Mail `
        -From $fromAddress `
        -To $validTo `
        -Subject $subject `
        -Body $body `
        -SmtpServer $NotificationConfig.SmtpServer `
        -KeyName $NotificationConfig.CertName `
        -Cert $NotificationConfig.Cert `
        -Port $NotificationConfig.Port `
        -LogFilePath $LogFilePath
}

<#
.SYNOPSIS
Sends an HTML email notifying ops that a cycle's Validate + archive completed
successfully.
.DESCRIPTION
Exit code 0 out of Invoke-AuditValidate.ps1 covers three genuinely different
outcomes - files deleted after backup verification, files retained because
source cleanup is disabled, or nothing needed backup at all - and silence on
success is exactly the blind spot this closes for an unattended pipeline that
can delete source files (see PROD SourceCleanup). Whether cleanup was enabled
is taken as an EXPLICIT parameter (Invoke-AuditValidate.ps1's summary carries
SourceCleanupEnabled directly) rather than inferred from ArchiveResult being
non-null: the two usually agree, but inferring a config flag from a side
effect of another script's internal branching is a silent trap for the next
refactor of that script - this function does not guess.
.PARAMETER NotificationConfig
Resolved notification config (from Resolve-NotificationConfig) - must carry
SmtpServer, Cert, From, OpsTeam, CertName, Port.
.PARAMETER Environment
Environment tag (PROD / QA) shown in the subject line.
.PARAMETER RunId
The run identifier that completed.
.PARAMETER StartDate
Cycle start date (yyyyMMdd) shown in the subject and intro.
.PARAMETER EndDate
Cycle end date (yyyyMMdd) shown in the subject and intro.
.PARAMETER ArchiveStatus
ArchiveStatus value for a successful cycle: Success / NoOp / NoSourceFiles.
.PARAMETER SourceCleanupEnabled
Explicit SourceCleanup.Enabled value from the validation summary - drives the
ENABLED/DISABLED wording and which result section is shown. Not inferred.
.PARAMETER ArchiveResult
Optional ArchiveResult object from the validation summary; expected when
SourceCleanupEnabled is true (rendered defensively - a $null ArchiveResult
with SourceCleanupEnabled=true is reported rather than causing a null-ref).
.PARAMETER ExpectedFileCount
Optional - number of files the cycle's manifest expected, for triage context.
.PARAMETER BackupFolder
Optional - backup folder path shown in the intro for triage.
.PARAMETER SummaryPath
Optional - path to backup-validation-summary.json.
.PARAMETER LogFilePath
Optional log file forwarded to Send-Mail.
.EXAMPLE
Send-ValidateCompletionNotification -NotificationConfig $cfg -Environment 'PROD' `
    -RunId '20260520_142743' -StartDate '20260506' -EndDate '20260520' `
    -ArchiveStatus 'Success' -SourceCleanupEnabled $true -ArchiveResult $summary.ArchiveResult
.NOTES
Dispatched once per cycle by Invoke-WeComAuditScheduler when AuditValidate
exits 0 for a run it just executed (via
Send-ValidateCompletionNotificationFromSummary, which de-dupes using an
independent marker file - see that function's own notes on why the summary's
own Notification block is not a reliable dedup source here).
#>
function Send-ValidateCompletionNotification {
    param(
        [Parameter(Mandatory)]
        [object]$NotificationConfig,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$RunId,
        [Parameter(Mandatory)]
        [string]$StartDate,
        [Parameter(Mandatory)]
        [string]$EndDate,
        [Parameter(Mandatory)]
        [string]$ArchiveStatus,
        [Parameter(Mandatory)]
        [bool]$SourceCleanupEnabled,
        [object]$ArchiveResult,
        [int]$ExpectedFileCount,
        [string]$BackupFolder,
        [string]$SummaryPath,
        [string]$LogFilePath
    )

    if (-not $NotificationConfig -or -not $NotificationConfig.SmtpServer -or -not $NotificationConfig.Cert) {
        throw "Notification config incomplete: missing SmtpServer or Cert."
    }
    if ($NotificationConfig.OpsTeam.Count -eq 0) {
        throw "Notification config has no OpsTeam recipients."
    }

    $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'
    $fromAddress = ([string]$NotificationConfig.From).Trim()
    if ($fromAddress -notmatch $emailPattern) {
        throw "Notification 'From' is not a valid email address: '$fromAddress'."
    }
    $validTo = @(
        $NotificationConfig.OpsTeam |
            Where-Object { $_ } |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ -match $emailPattern }
    )
    if ($validTo.Count -eq 0) {
        throw "Notification 'OpsTeam' has no valid email recipients."
    }
    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $introParts = New-Object 'System.Collections.Generic.List[string]'
    $introParts.Add("Run: $(& $enc $RunId)")
    $introParts.Add("Date range: $(& $enc $StartDate) - $(& $enc $EndDate)")
    $introParts.Add("Archive status: $(& $enc $ArchiveStatus)")
    $introParts.Add("Source cleanup: $(if ($SourceCleanupEnabled) { 'ENABLED' } else { 'DISABLED (source files retained)' })")
    if ($PSBoundParameters.ContainsKey('ExpectedFileCount')) { $introParts.Add("Expected files: $(& $enc $ExpectedFileCount)") }
    if ($BackupFolder) { $introParts.Add("Backup folder: $(& $enc $BackupFolder)") }
    if ($SummaryPath)  { $introParts.Add("Summary: $(& $enc $SummaryPath)") }
    $intro = ($introParts -join '<br/>')

    $resultLines = New-Object 'System.Collections.Generic.List[string]'
    if ($SourceCleanupEnabled -and $ArchiveResult) {
        if ($ArchiveResult.PSObject.Properties['DeletedCount']) { $resultLines.Add("Deleted: $(& $enc $ArchiveResult.DeletedCount)") }
        if ($ArchiveResult.PSObject.Properties['FailedCount'])  { $resultLines.Add("Failed: $(& $enc $ArchiveResult.FailedCount)") }
        if ($ArchiveResult.PSObject.Properties['SkippedCount']) { $resultLines.Add("Skipped: $(& $enc $ArchiveResult.SkippedCount)") }
    }
    elseif ($SourceCleanupEnabled) {
        # Defensive: SourceCleanupEnabled=true should normally come with an
        # ArchiveResult, but this function must never null-ref on a mismatch -
        # report the gap instead of crashing mid-notification.
        $resultLines.Add('Source cleanup is enabled, but no archive result details were recorded for this run.')
    }
    else {
        $noCleanupMessage = switch ($ArchiveStatus) {
            'NoOp'          { 'No files required backup this cycle (all already archived or none matched).' }
            'NoSourceFiles' { "No source files were present for this cycle's manifest." }
            default         { 'Source cleanup is disabled for this environment; backed-up files were retained in the source folder.' }
        }
        $resultLines.Add($noCleanupMessage)
    }

    $sections = @(
        [PSCustomObject]@{ Heading = 'Archive Result'; Items = @($resultLines.ToArray()) }
    )

    $body = Build-AuditNotificationHtml `
        -Heading "WeCom Audit Cycle Completed - Run $RunId" `
        -Intro $intro `
        -Sections $sections `
        -Footer 'Validation passed and the archive step completed. No action required.'

    $envTag = ([string]$Environment).ToUpperInvariant()
    $subject = "[WeCom Audit][$envTag] Cycle Completed - $RunId ($StartDate - $EndDate)"

    Send-Mail `
        -From $fromAddress `
        -To $validTo `
        -Subject $subject `
        -Body $body `
        -SmtpServer $NotificationConfig.SmtpServer `
        -KeyName $NotificationConfig.CertName `
        -Cert $NotificationConfig.Cert `
        -Port $NotificationConfig.Port `
        -LogFilePath $LogFilePath
}

<#
.SYNOPSIS
Sends the single deadline-escalation email when a cycle is not complete by the
final check (Thursday 18:00).
.DESCRIPTION
Replaces the retired multi-level reminder system (Sequence / Normal / Final /
LastCall). There is exactly one escalation, sent by the scheduler itself when
invoked with -Escalate and the cycle is still incomplete. Recipients are
OpsTeam plus config EscalationCc (managers). Wording states a fact ("cycle not
completed by deadline"), not a request - this email is the formal record of a
missed cycle.
.PARAMETER NotificationConfig
Resolved notification config (from Resolve-NotificationConfig).
.PARAMETER Environment
Environment tag (PROD / QA) shown in the subject line.
.PARAMETER CycleStartDate
Cycle start date (yyyyMMdd).
.PARAMETER CycleEndDate
Cycle end date (yyyyMMdd) - also the deadline date.
.PARAMETER PendingStage
'Analysis' or 'Validate' - the first stage that is still incomplete.
.PARAMETER MissingItems
Optional preflight missing items (Name, ExpectedPath, Source) for context.
.PARAMETER EscalationCc
Extra Cc addresses (typically managers) from config EscalationCc.
.PARAMETER LogFilePath
Optional log file forwarded to Send-Mail.
#>
function Send-AuditEscalationNotification {
    param(
        [Parameter(Mandatory)]
        [object]$NotificationConfig,
        [Parameter(Mandatory)]
        [string]$Environment,
        [Parameter(Mandatory)]
        [string]$CycleStartDate,
        [Parameter(Mandatory)]
        [string]$CycleEndDate,
        [Parameter(Mandatory)]
        [ValidateSet('Analysis', 'Validate')]
        [string]$PendingStage,
        [ValidateSet('DeadlineMiss', 'RetryExhausted', 'InvariantViolation')]
        [string]$Reason = 'DeadlineMiss',
        [string]$Detail,
        [object[]]$MissingItems = @(),
        [string[]]$EscalationCc = @(),
        [string]$LogFilePath
    )

    if (-not $NotificationConfig -or -not $NotificationConfig.SmtpServer -or -not $NotificationConfig.Cert) {
        throw "Notification config incomplete: missing SmtpServer or Cert."
    }
    if ($NotificationConfig.OpsTeam.Count -eq 0) {
        throw "Notification config has no OpsTeam recipients."
    }

    $enc = { param($s) [System.Net.WebUtility]::HtmlEncode([string]$s) }

    $missingRendered = @(
        $MissingItems | ForEach-Object {
            "<b>$(& $enc $_.Name)</b>: $(& $enc $_.ExpectedPath) <i>[$(& $enc $_.Source)]</i>"
        }
    )

    $sections = @()
    if ($missingRendered.Count -gt 0) {
        $sections += [PSCustomObject]@{ Heading = 'Outstanding Files'; Items = $missingRendered }
    }

    if ($Reason -eq 'RetryExhausted') {
        $introText = "Analysis for cycle $(& $enc $CycleStartDate) - $(& $enc $CycleEndDate) has failed repeatedly and automatic retries have STOPPED. $(& $enc $Detail) A persistent failure like this usually means a deterministic problem (e.g. input log format change), not an infrastructure blip - ENGINEERING investigation of the analysis scripts is required. Ops action: none until engineering confirms a fix, then trigger run-now.cmd."
        $footerText = 'Automatic retries are exhausted for this cycle. This mail is addressed to engineering; operations does not need to act on the source folder.'
        $headingText = "WeCom Audit ESCALATION - Analysis failing repeatedly ($CycleEndDate)"
        $subjectCore = "ESCALATION - Analysis failing repeatedly, engineering required (cycle $CycleEndDate)"
    }
    elseif ($Reason -eq 'InvariantViolation') {
        $introText = "Cycle $(& $enc $CycleStartDate) - $(& $enc $CycleEndDate) cannot continue because persisted analysis state is missing or invalid. $(& $enc $Detail) Engineering investigation is required before rerunning the cycle."
        $footerText = 'The pipeline failed closed before validation or archive. Repair or restore the named run artifacts, then trigger run-now.cmd.'
        $headingText = "WeCom Audit ERROR - Invalid analysis state ($CycleEndDate)"
        $subjectCore = "ERROR - Invalid analysis state, engineering required (cycle $CycleEndDate)"
    }
    else {
        $introText = "Cycle $(& $enc $CycleStartDate) - $(& $enc $CycleEndDate) was not completed by the 18:00 deadline. Pending stage: $(& $enc $PendingStage)."
        $footerText = 'This is the formal deadline-miss record for this cycle. Complete the pending stage (drop the outstanding files into the source folder), or engage engineering if the cycle cannot be recovered.'
        $headingText = "WeCom Audit ESCALATION - Cycle $CycleEndDate not completed"
        $subjectCore = "ESCALATION - Cycle $CycleEndDate not completed (pending: $PendingStage)"
    }

    $body = Build-AuditNotificationHtml `
        -Heading $headingText `
        -Intro $introText `
        -Sections $sections `
        -Footer $footerText

    $envTag = ([string]$Environment).ToUpperInvariant()
    $subject = "[WeCom Audit][$envTag] $subjectCore"

    $emailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

    $fromAddress = ([string]$NotificationConfig.From).Trim()
    if ($fromAddress -notmatch $emailPattern) {
        throw "Notification 'From' is not a valid email address: '$fromAddress'."
    }

    $validTo = @(
        $NotificationConfig.OpsTeam |
            Where-Object { $_ } |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ -match $emailPattern }
    )
    if ($validTo.Count -eq 0) {
        throw "Notification 'OpsTeam' has no valid email recipients."
    }

    $validCc = @(
        @($EscalationCc) |
            Where-Object { $_ } |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_ -match $emailPattern } |
            Select-Object -Unique
    )
    $ccStr = if ($validCc.Count -gt 0) { $validCc -join ',' } else { $null }

    Send-Mail `
        -From $fromAddress `
        -To $validTo `
        -Cc $ccStr `
        -Subject $subject `
        -Body $body `
        -SmtpServer $NotificationConfig.SmtpServer `
        -KeyName $NotificationConfig.CertName `
        -Cert $NotificationConfig.Cert `
        -Port $NotificationConfig.Port `
        -LogFilePath $LogFilePath
}
