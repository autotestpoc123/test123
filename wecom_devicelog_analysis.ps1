param(
    [Parameter(Mandatory = $true)]
    [string]$deviceLogFilePath,
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{8}$')]
    [string]$startDate,
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d{8}$')]
    [string]$endDate,
    [Parameter(Mandatory = $true)]
    [ValidateSet('MSMS', 'MSBIC')]
    [string]$BU,
    [ValidateSet('PROD', 'QA')]
    [string]$env = 'PROD',
    [string]$SummaryOutputPath,
    [string]$TaskOutputDirectory,
    # Sprint 2: mail-ledger parameters (mandatory once wired via Invoke-AuditLog.ps1).
    [Parameter(Mandatory = $true)]
    [string]$TaskName,
    [Parameter(Mandatory = $true)]
    [string]$CycleId,
    [Parameter(Mandatory = $true)]
    [string]$LedgerPath,
    [Parameter(Mandatory = $true)]
    [string]$SentEmailsPath,
    [Parameter(Mandatory = $true)]
    [string]$RunId
)

$parentFolderPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$importModulePath = Join-Path $parentFolderPath "wecom_analysis_comm.psm1"

if (-not (Test-Path $importModulePath)) {
    throw "Module load path not found, please double check!"
}

Import-Module $importModulePath

# Sprint 2.1 (#1): collect ledger content-diff rejections and fail the task at
# the end so operators see an explicit failure rather than a silent-Success run.
# Sub-loops (per BU, per branch) all push into $rejected; each exit point calls
# Assert-NoBuMailRejections before returning.
$rejected = [System.Collections.Generic.List[object]]::new()

function Assert-NoBuMailRejections {
    $count = @($script:rejected).Count
    if ($count -gt 0) {
        $buList = ($script:rejected | ForEach-Object { $_.BU }) -join ', '
        throw "Task '$($script:TaskName)' rejected $count BU send(s) due to content-diff: $buList. Fix source data and rerun with -ForceRerun (engineering only), or use Invoke-BuMailResend.ps1 for per-BU resend."
    }
}

# Generate Export Folder
function Save-AnalysisSummary {
    param(
        [bool]$HasViolation,
        [int]$ViolationDivisionCount,
        [int]$ViolationRecordCount
    )

    $summaryFields = [ordered]@{
        AnalysisType           = 'Device'
        BusinessUnit           = $BU
        StartDate              = $startDate
        EndDate                = $endDate
        HasViolation           = $HasViolation
        ViolationDivisionCount = $ViolationDivisionCount
        ViolationRecordCount   = $ViolationRecordCount
    }

    Write-AnalysisSummaryJson -SummaryOutputPath $SummaryOutputPath -SummaryFields $summaryFields -Depth 4
}

# main procedure
$msmsDivisionScope = @("Private Credit & Equity", "Real Assets", "Global Sales and Marketing")
$msbicDivision = "Fixed Income Division"
# $msbicContacts = "css-wecom@abc.com"
$msViolationCollection = @()
$msbicViolationCollection = @()

$destFoderPath = Export-AnalysisReport -LogFilePath $deviceLogFilePath -TaskOutputDirectory $TaskOutputDirectory -SubFolder 'analyzed' -UseDateSubFolder
$tempLogPath = $deviceLogFilePath
$logFilePath = if ($TaskOutputDirectory) {
    Join-Path $destFoderPath 'task.log'
}
else {
    Get-LogFilePath -Directory $destFoderPath -BaseName "AnalysisLog"
}
$destFilePath = Join-Path $destFoderPath "report.csv"
$Subject = ""

# --set value according to Env
$vaultEnv = 'prod'
$prodid = "cod_wecom_ntfy_prod@abc.com.cn"
$idName = "cod_wecom_ntfy_prod"
$vault_server = "https://vault.srv.ms.com.cn"
$ldapServer = 'cod.ms.com.cn'
$smtp_server = "mta-hub.cod.ms.com.cn"
$domain = 'COD'
$BuContacter = $null

$MsNoviolationRecipients = @("Cynthia.Xu@abc.com.cn", "Jun.Xu@abc.com", "susan.sun@abc.com")
$MsbicBuContacter = "css-wecom@abc.com"
$CcContacter = "cod-wecom-admin@abc.com.cn"

$MSBURecipients = @{
    "Private Credit & Equity"   = @("Cynthia.Xu@abc.com.cn", "Jun.Xu@abc.com", "msim_wecom_weekly_contactstatusreport")
    "Real Assets"               = @("Susan.Sun@abc.com", "ivy.zhou@abc.com.cn", "msim_wecom_weekly_contactstatusreport")
    "Global Sales and Marketing" = @("Matthew.Zhu@abc.com.cn", "msim_wecom_weekly_contactstatusreport")
}

if ($env.ToLower() -eq 'qa') {
    $vaultEnv = "qa"
    $prodid = "wecom_deploy_qa@infradev.abc.com.cn"
    $idName = "wecom_deploy_qa"
    $vault_server = "https://vault.srv.lab.ms.com.cn"
    $ldapServer = 'codqa.lab.ms.com.cn'
    $domain = 'CODQA'
    $smtp_server = "mta-hub.mail.lab.ms.com.cn"
    # Todo need update according to alignment with BUs
    $MsNoviolationRecipients = @("abc.test@dev.abc.com.cn", "yimin.lu06@infradev.abc.com.cn")
    $MSBURecipients = @{
        "Private Credit & Equity"   = @("abc.test@dev.abc.com.cn", "Siyi.Huang@infradev.abc.com.cn")
        "Real Assets"               = @("abc.test@dev.abc.com.cn", "yimin.lu06@infradev.abc.com.cn")
        "Global Sales and Marketing" = @("abc.test@dev.abc.com.cn")
    }
    $CcContacter = "abc.test@dev.abc.com.cn"
    $MsbicBuContacter = "yinmin.lu06@infradev.abc.com.cn"
}

$noViolationContent = "The purpose of this email is to provide information on users in your business unit( BU ) who used unapproved WeCom device(s) for the reporting period listed in the subject line.<br><br/>There were <b>no violations</b> to report this reporting period for your BU.<br><br/>We are only able to distinguish if a user is using WeCom via iPad/ windows/ mac/ and cannot determine if user uses their own iOS mobile to login. This is currently a known limitation for the logs we retrieve."
$DeviceViolationContent = "The purpose of this email is to provide information on users in your business unit( BU ) who used unapproved WeCom device(s) for the reporting period listed in the subject line.<br><br/>We are only able to distinguish if a user is using WeCom via iPad/ windows/ mac/ and cannot determine if user uses their own iOS mobile to login. This is currently a known limitation for the logs we retrieve.<br/>The violation record(s) of this reporting period for your BU as below:<br/>"

if ($BU.ToLower() -eq 'msms') {
    $Subject = "COD WeCom Login to Non-Approved Devices IM BU - Report($startDate - $endDate)"
}
else {
    $Subject = "COD WeCom Login to Non-Approved Devices FID BU - Report($startDate - $endDate)"
}

$violationcounter = 0

try {
    $null = Convert-ExactDate $startDate
    $null = Convert-ExactDate $endDate

    Write-Log "Start to handle device log analysis!" -LogFilePath $logFilePath
    # get system id cert from windows server
    $sysid_cert = Get-Cert -KeyName $prodid
    # get vault secret correctly
    $vault_secret = Get-VaultSecret -VaultServer $vault_server -VaultEnv $vaultEnv -Eonid "309843" -KeyName $prodid -SysIdCert $sysid_cert
    $idSecret = New-Object System.Net.NetworkCredential($idName, $vault_secret, $Domain)
    # build up Ladp lazy connection
    $lazyConn = New-LazyLdapConnection -Server $ldapServer -Port 363 -Credential $idSecret
    # start to do export data
    $deviceData = Import-Csv -Path $tempLogPath -Encoding UTF8
    $filterIds = ($deviceData.Account.ToLower().Trim() | Sort-Object -Unique) -join ';'
    $cNamelookup = Get-LdapUserById -LazyConnection $lazyConn -UserId $filterIds
}
catch {
    Write-Log "Failed during LDAP connection or Export CSV data: $_" -LogFilePath $logFilePath
    return
}
finally {
    if ($lazyConn) {
        Close-LazyLdapConnection -lazy $lazyConn
    }
}

# process Device data log
foreach ($record in $deviceData) {
    $platform = $record.Platform.ToLower().Trim()
    if ($platform -eq 'ios(iphone)') { continue }
    $userId = $record.Account.ToLower().Trim()
    if ($cNamelookup.Valid.ContainsKey($userId)) {
        $division = $cNamelookup.Valid[$userId].Division
        $department = $record.Department.split('/')[1]
        # convert to englist if its mandarin
        $status = if ($record.Status -eq '使用') { 'Used' } else { $record.Status }
        if ($BU -eq 'MSMS' -and $msmsDivisionScope -contains $division) {
            $msViolationCollection += [PSCustomObject]@{
                Time         = $record.Time
                Name         = $record.Name
                Account      = $record.Account
                Department   = $department
                Status       = $status
                'LastUsedOn' = $record.'Last Used on'
                Platform     = $record.Platform
                Division     = $division
            }
            $violationcounter += 1
        }
        elseif ($BU -eq 'MSBIC' -and $division -eq $msbicDivision) {
            $msbicViolationCollection += [PSCustomObject]@{
                Time         = $record.Time
                Name         = $record.Name
                Account      = $record.Account
                Department   = $department
                Status       = $status
                'LastUsedOn' = $record.'Last Used on'
                Platform     = $record.Platform
                Division     = $msbicDivision
            }
            $violationcounter += 1
        }
        else {
            Write-Log "No macthing divsion condition for UserId: ${userId} in Record item: ${record}" -LogFilePath $logFilePath
        }
    }
    else {
        Write-Log "Invalid UserId found: ${userId} in Record item: ${record}" -LogFilePath $logFilePath
    }
}

if ($violationcounter -eq 0) {
    Write-Host "No Violation Usages found" -ForegroundColor Green
    Write-Log "No Violation Usages found" -LogFilePath $LogFilePath
    $htmlBody = New-HtmlBody -TableHtml "" -ViolationContent "" -NoViolationContent $noViolationContent -HasViolation:$false
    Save-AnalysisSummary -HasViolation $false -ViolationDivisionCount 0 -ViolationRecordCount 0

    if ($BU -eq 'MSMS') {
        # noviolation of MSMS
        $BuContacter = $MsNoviolationRecipients
        # Ledger key: 'MSMS-ALL' distinguishes the family-wide summary send (fixed
        # notification list) from per-division sends that use $BuItem as the key.
        $ledgerBu = 'MSMS-ALL'
    }
    else {
        # noviolation of MSBIC
        $BuContacter = $MsbicBuContacter
        # Separate from the per-BU 'MSBIC' key used at line 272 so a later run
        # with real violations does not trigger content-diff Warn on the same key.
        $ledgerBu = 'MSBIC-ALL'
    }

    $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $ledgerBu -RunId $RunId `
        -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
        -From $prodid -To $BuContacter -Cc $CcContacter -Subject $Subject `
        -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port 2587 -LogFilePath $logFilePath
    if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    Assert-NoBuMailRejections

    return
}
else {
    # violation records in MSMS
    if ($msViolationCollection.Count -gt 0) {
        $msmsViolationsByBU = $msViolationCollection | Group-Object 'Division' -AsHashTable
        Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount $msmsViolationsByBU.Keys.Count -ViolationRecordCount $msViolationCollection.Count
        Write-Verbose "MSMS Violation Founds"
        $msViolationCollection | Export-Csv -Path $destFilePath -NoTypeInformation -Encoding UTF8 -Force
        Write-Log "Violation Usage In MSMS found" -LogFilePath $logFilePath
        foreach ($BuItem in $msmsViolationsByBU.Keys + ($MSBURecipients.Keys | Where-Object { $_ -notin $msmsViolationsByBU.Keys })) {
            $BuContacter = $MSBURecipients[$BuItem]
            $hasViolation = $msmsViolationsByBU.ContainsKey($BuItem)
            if ($hasViolation) {
                Write-Verbose $BuItem
                $rowsHtml = $msmsViolationsByBU[$BuItem] | Select-Object @{ Name = 'DateTime(HKT)'; Expression = { "$($_.Time)" } }, Name, Account, Department, Status, LastUsedOn, Platform, Division | ConvertTo-Html -Fragment
                $tableHtml = @"
<table>
    $rowsHtml
</table>
"@
                $htmlBody = New-HtmlBody -TableHtml $tableHtml -ViolationContent $DeviceViolationContent -NoViolationContent "" -HasViolation:$true
                $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $BuItem -RunId $RunId `
                    -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
                    -From $prodid -To $BuContacter -Cc $CcContacter -Subject $Subject `
                    -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port 2587 -LogFilePath $logFilePath
                if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
            }
            else {
                Write-Verbose $BuItem
                Write-Verbose "this BU has no violations, send normal mail"
                $htmlBody = New-HtmlBody -TableHtml "" -ViolationContent "" -NoViolationContent $NoViolationContent -HasViolation:$false
                $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $BuItem -RunId $RunId `
                    -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
                    -From $prodid -To $BuContacter -Cc $CcContacter -Subject $Subject `
                    -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port 2587 -LogFilePath $logFilePath
                if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
            }
        }
    }

    # violation records in MSBIC
    if ($msbicViolationCollection.Count -gt 0) {
        Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount 1 -ViolationRecordCount $msbicViolationCollection.Count
        $msbicViolationCollection | Export-Csv -Path $destFilePath -NoTypeInformation -Encoding UTF8 -Force
        Write-Log "Violation Usage In MSBIC found" -LogFilePath $logFilePath
        $rowsHtml = $msbicViolationCollection | Select-Object @{ Name = 'DateTime(HKT)'; Expression = { "$($_.Time)" } }, Name, Account, Department, Status, LastUsedOn, Platform, Division | ConvertTo-Html -Fragment
        $tableHtml = @"
<table>
    $rowsHtml
</table>
"@
        $htmlBody = New-HtmlBody -TableHtml $tableHtml -ViolationContent $DeviceViolationContent -NoViolationContent "" -HasViolation:$true
        $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU 'MSBIC' -RunId $RunId `
            -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
            -From $prodid -To $MsbicBuContacter -Cc $CcContacter -Subject $Subject `
            -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port 2587 -LogFilePath $logFilePath
        if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    }
}

Assert-NoBuMailRejections

