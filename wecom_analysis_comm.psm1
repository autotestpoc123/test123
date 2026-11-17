# wecom_analysis_comm.psm1 - facade over modules/internal/*.ps1.
# All functions live in ONE module scope via dot-sourcing: private helpers
# stay visible across files, the export contract below is unchanged, and
# callers keep 'Import-Module .\wecom_analysis_comm.psm1' untouched.
# Internal files contain FUNCTIONS ONLY; every top-level statement is hoisted
# into this preamble so dot-source order below is provably irrelevant.

if (-not ('System.DirectoryServices.Protocols.LdapConnection' -as [type])) {
    Add-Type -AssemblyName System.DirectoryServices.Protocols
}

$script:WeComAuditLogFolderName = 'wecom_audit_log'

$htmlTemplateNew = @"
<html>
<head>
    <style>
        table { border-collapse: collapse; width: auto; }
        th, td { border: 1px solid #ddd; padding: 8px; }
        th { background-color: #f2f2f2; text-align: left; }
        body { font-family: Arial; font-size: 16px; }
    </style>
</head>
<body>
    <div>Hi all,</div>
    {{ViolationParagraph}}

    {{TableSection}}

    {{NoViolationParagraph}}
    <br/>
    <br/>
    <div>Regards,</div>
    <div>COD WeCom Team</div>
</body>
</html>
"@

$internalFiles = @('Core.ps1', 'Config.ps1', 'State.ps1', 'Notification.ps1',
    'Analysis.ps1', 'Archive.ps1', 'SourceCleanup.ps1', 'Identity.ps1')
foreach ($name in $internalFiles) {
    $path = Join-Path $PSScriptRoot "modules\internal\$name"
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "WeCom module deployment is incomplete: missing internal file '$path'."
    }
    . $path
}

Export-ModuleMember -Function Convert-ExactDate, Write-Log, Get-Cert, Get-LogFilePath, Get-VaultSecret, New-LazyLdapConnection, Close-LazyLdapConnection, Send-Mail, New-HtmlBody, New-DateTokenMap, New-AuditTokenMap, Resolve-TemplateText, Get-OptionalObjectPropertyValue, Write-AnalysisSummaryJson, Assert-ConfigInputDirectories, Get-TaskResultByName, Get-TaskSummaryData, Get-TaskSummariesByRunId, Resolve-DynamicSummaryTaskRequirements, Get-BackupValidationConfig, Get-ExpectedBackupFiles, Test-BackupFolderContent, Get-SourceCopyTargets, Format-BackupValidationText, Export-AnalysisReport, New-LdapOrFilter, Split-LdapBatches, Resolve-LdapSearchBase, Get-LdapUserByMail, Get-LdapUserById, Get-WeComAuditLogFolderName, Get-NormalizedFullPath, Test-PathWithinAllowedRoots, Test-SafeToDeleteSourceFile, Remove-SourceFileWithRetry, Invoke-SourceFileCleanup, Test-AllowedRootIsTooBroad, Resolve-SourceCleanupConfig, Assert-SourceCleanupConfig, Assert-TaskNameUniqueness, Resolve-AuditConfigPath, Resolve-AuditOutputRoot, Resolve-AuditInputRoot, Resolve-AuditSourceFolder, Resolve-TaskInputPath, Resolve-NotificationConfig, Get-PreflightFiles, Test-PreflightReady, Rename-MislabeledXlsInputs, Send-PreflightNotification, Send-ValidationFailureNotification, Send-ArchiveFailureNotification, Send-ValidateCompletionNotification, Send-AuditEscalationNotification, Resolve-ScheduleCycle, Resolve-PhaseHandoff, Get-BuMailContentHash, Get-MailLedgerPath, Test-MailLedgerHit, Add-MailLedgerEntry, Add-SentEmailRecord, Test-AnalysisCycleAlreadyComplete, Test-ValidateCycleAlreadyComplete, Send-AuditBuMail, Import-MsimcIdentityMapping, Resolve-MsimcIdentity, Get-MsimcLdapUserByAccount
