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
    [ValidateSet('MSMS', 'MSBIC', 'MSIMC')]
    [string]$BU,
    [Parameter(Mandatory = $true)]
    [object]$AnalysisEnvironment,
    [string]$SummaryOutputPath,
    [string]$TaskOutputDirectory,
    [Parameter(Mandatory = $true)]
    [String]$TaskName,
    [Parameter(Mandatory = $true)]
    [string]$CycleId,
    [Parameter(Mandatory = $true)]
    [string]$LedgerPath,
    [Parameter(Mandatory = $true)]
    [string]$sentEmailsPath,
    [Parameter(Mandatory = $true)]
    [string]$RunId,
    [string[]]$ReportRecipients,
    [hashtable]$IdentityMapping,
    [string]$IdentityMappingPath
)

$parentFolderPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$importModulePath = Join-Path $parentFolderPath "wecom_analysis_comm.psm1"

if (-not (Test-Path $importModulePath)) {
    throw "Module load path not found, please double check!"
}

Import-Module $importModulePath
$rejected = [System.Collections.Generic.List[object]]::new()

function Assert-NoBuMailRejections {
    $count = $script:rejected.Count
    if ($count -gt 0) {
        $buList = ($script:rejected | ForEach-Object { $_.BU }) -join ','
        throw "Task '$($script:TaskName)' rejected $count BU send(s) due to content-diff: $buList."
    }
}

# Build ldap query filter
function Build-OrFilter($batches) {
    $classBlock = '(objectClass=user)'
    $cNameClauses = $batches | ForEach-Object { "(cn=$(ConvertTo-LdapFilterEscaped $_))" }
    $cNameBlock = "(|" + ($cNameClauses -join '') + ")"
    return "(&${classBlock}${cNameBlock})"
}

# Batch query ldap with users Id
function Get-LdapUserById {
    param(
        [Parameter(Mandatory = $true)][System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]$lazyConnection,
        [Parameter(Mandatory = $true)]
        [string[]]$UserId,
        [string]$SearchBase,
        [int]$SearchTimeoutSeconds = 30,
        [string[]]$attr = @("sAMAccountName", "division", "legalentity"),
        [int]$BatchSize = 20
    )

    Add-Type -AssemblyName System.DirectoryServices.Protocols

    # force creation / bind
    try {
        $conn = $lazyConnection.Value
    }
    catch {
        throw "failed to create LDAP connection: $($_.Exception.Message)"
    }

    # set timeout
    $conn.Timeout = [TimeSpan]::FromSeconds($SearchTimeoutSeconds)

    # discover defaultNamingContext if SearchBase not provided
    if (-not $SearchBase) {
        $rootReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
            "",
            "(objectClass=*)",
            [System.DirectoryServices.Protocols.SearchScope]::Base,
            "defaultNamingContext"
        )
        $rootResp = $conn.SendRequest($rootReq)
        if ($rootResp.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success) {
            throw "Failed to read RootDSE: $($rootResp.ErrorMessage)"
        }
        $SearchBase = ($rootResp.Entries[0].Attributes["defaultNamingContext"])[0]
    }

    # initial cName lookup
    $CNameLookup = @{}
    $valid = @{}
    $invalid = New-Object System.Collections.Generic.List[string]

    foreach ($batch in ($UserId | ForEach-Object -Begin { $tmp = @() } -Process {
                $tmp += $_
                if ($tmp.count -gt $BatchSize) { , $tmp; $tmp = @() }
            } -End { if ($tmp.count) { , $tmp } })) {
        $batchIds = $batch.split(';').Trim()

        # build filter
        $filter = Build-OrFilter $batchIds
        Write-Verbose $filter

        $searchReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
            $SearchBase,
            $filter,
            [System.DirectoryServices.Protocols.SearchScope]::Subtree,
            $attr
        )

        $searchResp = $conn.SendRequest($searchReq)

        if (-not $valid) { $valid = @{} }

        $foundCnNames = @()

        foreach ($entry in $searchResp.Entries) {
            $info = [PSCustomObject]@{
                SamAccountName = $null
                Division       = $null
                LegalEntity    = $null
            }

            if ($entry.Attributes["sAMAccountName"]) {
                $info.SamAccountName = [string]$entry.Attributes["sAMAccountName"][0]
                $foundCnNames += $info.SamAccountName
            }

            if ($entry.Attributes["division"]) {
                $info.Division = [string]$entry.Attributes["division"][0]
            }

            if ($entry.Attributes["legalentity"]) {
                $info.LegalEntity = [string]$entry.Attributes["legalentity"][0]
            }

            if ($info.SamAccountName) {
                $CNameLookup[$info.SamAccountName.ToLower()] = $info
            }
        }

        foreach ($id in $batchIds) {
            if (-not ($foundCnNames -contains $id)) {
                $invalid.Add($id)
            }
        }
    }

    return @{
        Valid   = $CNameLookup
        Invalid = $invalid
    }
}

function Save-AnalysisSummary {
    param(
        [bool]$HasViolation,
        [int]$ViolationDivisionCount,
        [int]$ViolationRecordCount = 0,
        [int]$InputRecordCount = 0,
        [int]$AnalyzedRecordCount = 0,
        [int]$SkippedIdentityCount = 0
    )

    if (-not $SummaryOutputPath) {
        return
    }

    Write-AnalysisSummaryJson -SummaryOutputPath $SummaryOutputPath -SummaryFields ([ordered]@{
            AnalysisType           = 'Device'
            BusinessUnit           = $BU
            StartDate              = $startDate
            EndDate                = $endDate
            HasViolation           = $HasViolation
            ViolationRecordCount   = $ViolationRecordCount
            InputRecordCount       = $InputRecordCount
            AnalyzedRecordCount    = $AnalyzedRecordCount
            SkippedIdentityCount   = $SkippedIdentityCount
            ViolationDivisionCount = $ViolationDivisionCount
        }) -Depth 4
}


# main procedure
$msmsDivisionScope = @("Private Credit & Equity", "Real Assets", "Global Sales and Marketing")
$msbicDivision = "Fixed Income Division"
$usedStatusToken = -join (@(0x4F7F,0x7528) | ForEach-Object { [char]$_ })
$msViolationCollection = @()
$msbicViolationCollection = @()

$destFolderPath = Export-AnalysisReport -LogFilePath $deviceLogFilePath -TaskOutputDirectory $TaskOutputDirectory -SubFolder 'analyzed' -UseDateSubFolder
$tempLogPath = $deviceLogFilePath
$logFilePath = if ($TaskOutputDirectory) {
    Join-Path $destFolderPath 'task.log'
}
else {
    Get-LogFilePath -Directory $destFolderPath -BaseName "AnalysisLog"
}
$destFilePath = Join-Path $destFolderPath "report.csv"
$Subject = ""

# --resolve environment settings from config
$vaultEnv = $AnalysisEnvironment.VaultEnv
$prodid = $AnalysisEnvironment.ProdId
$idName = $AnalysisEnvironment.IdName
$vault_server = $AnalysisEnvironment.VaultServer
$ldapServer = $AnalysisEnvironment.LdapServer
$smtp_server = $AnalysisEnvironment.SmtpServer
$smtp_port = $AnalysisEnvironment.Port
$domain = $AnalysisEnvironment.Domain
$BuContacter = $null

$MsNoviolationRecipients = $AnalysisEnvironment.DeviceNoviolationRecipients
$MsbicBuContacter = $AnalysisEnvironment.MsbicBuContacter
$CcContacter = $AnalysisEnvironment.Cc
$MSBURecipients = $AnalysisEnvironment.DeviceBURecipients

# SEC-005 (accepted risk): violation notification emails intentionally include
# employee PII (names, accounts, departments) in the table below. This data is
# already registered in the firm DLM inventory, recipients are controlled BU
# contacts, and masking is not required. See SECURITY_REVIEW.md, SEC-005.
$noViolationContent = "The purpose of this email is to provide information on users in your business unit(BU) that have logged into non-approved
                        WeCom device(s) for the reporting period listed in the subject line.<br><br/> `
                        There were <b>no violations</b> to report this reporting period for your BU.<br><br/> `
                        We are only able to distinguish if a user is using WeCom via iPad/ windows/ mac/ android, but not a
                        their own IOS mobile to login. This is currently a known limitation for the logs which vendor is no
                       "
$DeviceViolationContent = "The purpose of this email is to provide information on users in your business unit(BU) that have
                        WeCom device(s) for the reporting period listed in the subject line.<br><br/> `
                        We are only able to distinguish if a user is using WeCom via iPad/ windows/ mac/ android, but not a
                        their own IOS mobile to login. This is currently a known limitation for the logs which vendor is no
                       "
$endDateDisplay = [datetime]::ParseExact($endDate, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture).AddDays(-1).ToString('yyyy-MM-dd') 

if ($BU -eq 'MSMS') {
    $Subject = "COD WeCom Login to Non-Approved Devices IM BU - Report(${startDate} - ${endDateDisplay})"
}
elseif ($BU -eq 'MSIMC') {
    $Subject = "COD WeCom Login to Non-Approved Devices MSIMC BU - Report(${startDate} - ${endDateDisplay})"
}
else {
    $Subject = "COD WeCom Login to Non-Approved Devices FID BU - Report(${startDate} - ${endDateDisplay})"
}


$violationcounter = 0
$lazyConn = $null

try {
    $null = Convert-ExactDate $startDate
    $null = Convert-ExactDate $endDate

    Write-Log "Start to handle device log analysis!" -LogFilePath $logFilePath
    # get system id cert from windows server
    $sysid_cert = Get-Cert -KeyName $prodid
    # get vault secret correctly
    $vault_secret = Get-VaultSecret -VaultEnv $vaultEnv -KeyName $prodid -sysIdCert $sysid_cert -VaultServer $vault_server
    $id_secret = New-Object System.Net.NetworkCredential($idName, $vault_secret, $Domain)
    #build up Ladp lazy connection
    $lazyConn = New-LazyLdapConnection -Server $ldapServer -Port 636 -Credential $id_secret
    # start to do export data
    $deviceData = @(Import-Csv -Path $tempLogPath -Encoding UTF8)
    if ($deviceData.Count -eq 0) {
        throw "Input device log file is empty: $tempLogPath"
    }
    $missingPlatform = @($deviceData | Where-Object { -not $_.Platform -or -not $_.Platform.Trim() })
    if ($missingPlatform.Count -gt 0) {
        throw "Device log contains $($missingPlatform.Count) record(s) without Platform."
    }
    $filterIds = @($deviceData | ForEach-Object { $_.Account.ToLower().Trim() } | Where-Object { $_ } | Sort-Object -Unique) -join ';'
    if ($BU -eq 'MSIMC') {
        if (-not $IdentityMapping) {
            if (-not $IdentityMappingPath) { $IdentityMappingPath = Join-Path (Split-Path $parentFolderPath -Parent) 'convert_table.xlsx' }
            $IdentityMapping = Import-MsimcIdentityMapping -Path $IdentityMappingPath
        }
        $resolutions = Resolve-MsimcIdentity -Identities @($filterIds -split ';') -Kind Account -Mapping $IdentityMapping
        $accounts = @($resolutions.Values | ForEach-Object { $_.CorporateAccount } | Sort-Object -Unique)
        $corporateLookup = Get-MsimcLdapUserByAccount -LazyConnection $lazyConn -Accounts $accounts
        $cNamelookup = @{ Valid = @{}; Invalid = @() }
        foreach ($originalAccount in $resolutions.Keys) {
            $corporateAccount = $resolutions[$originalAccount].CorporateAccount
            if ($corporateLookup.Valid.ContainsKey($corporateAccount)) {
                $cNamelookup.Valid[$originalAccount] = $corporateLookup.Valid[$corporateAccount]
            }
            else { $cNamelookup.Invalid += $originalAccount }
        }
    }
    else { $cNamelookup = Get-LdapUserById -LazyConnection $lazyConn -UserId $filterIds }
}
catch {
    Write-Log "Failed during LDAP connection or Export CSV data: $_" -LogFilePath $logFilePath
    throw
}
finally {
    if ($lazyConn) {
        Close-LazyLdapConnection -lazy $lazyConn
    }
}

if ($BU -eq 'MSIMC') {
    $normalizedRecipients = @(
        @($ReportRecipients) |
            ForEach-Object { if ($_) { $_.Trim().ToLowerInvariant() } } |
            Where-Object { $_ } |
            Select-Object -Unique
    )
    if ($normalizedRecipients.Count -eq 0) {
        throw 'MSIMC device analysis requires at least one report recipient.'
    }
    $msimcViolations = New-Object System.Collections.Generic.List[object]
    $skippedIdentityCount = 0
    $analyzedRecordCount = 0
    foreach ($record in @($deviceData | Sort-Object Time, Account, Platform)) {
        $platform = $record.Platform.Trim().ToLowerInvariant()
        if ($platform -eq 'ios(iphone)') { continue }
        $userId = $record.Account.Trim().ToLowerInvariant()
        if (-not $cNamelookup.Valid.ContainsKey($userId)) {
            $skippedIdentityCount++
            Write-Log "MSIMC invalid account skipped: $userId" -LogFilePath $logFilePath
            continue
        }
        $analyzedRecordCount++
        $department = if ($record.Department -and $record.Department.Contains('/')) { $record.Department.Split('/')[1] } else { $record.Department }
        $msimcViolations.Add([PSCustomObject]@{
            Time       = $record.Time
            Name       = $record.Name
            Account    = $record.Account
            Department = $department
            Status     = if ($record.Status -eq $usedStatusToken) { 'Used' } else { $record.Status }
            LastUsedOn = $record.'Last Used on'
            Platform   = $record.Platform
            Division   = [string]$cNamelookup.Valid[$userId].Division
        })
    }
    if ($analyzedRecordCount -eq 0) {
        throw "MSIMC device analysis could not identify any account. SkippedIdentityCount=$skippedIdentityCount."
    }
    $knownDivisionCount = @($msimcViolations | Where-Object { $_.Division } | Select-Object -ExpandProperty Division -Unique).Count
    $coverageNote = "<br/>Coverage: input=$($deviceData.Count); analyzed=$analyzedRecordCount; skippedIdentity=$skippedIdentityCount.<br/>"
    if ($msimcViolations.Count -eq 0) {
        Save-AnalysisSummary -HasViolation $false -ViolationDivisionCount 0 -ViolationRecordCount 0 -InputRecordCount $deviceData.Count -AnalyzedRecordCount $analyzedRecordCount -SkippedIdentityCount $skippedIdentityCount
        $htmlBody = New-HtmlBody -TableHtml '' -ViolationContent '' -NoViolationContent ($noViolationContent + $coverageNote) -HasViolation:$false
    }
    else {
        $msimcViolations | Export-Csv -Path $destFilePath -NoTypeInformation -Encoding UTF8 -Force
        $rowsHtml = $msimcViolations | Select-Object @{ Name = 'DateTime(HKT)'; Expression = { $_.Time } }, Name, Account, Department, Status, LastUsedOn, Platform, Division | ConvertTo-Html -Fragment 
        $htmlBody = New-HtmlBody -TableHtml "<table>$rowsHtml</table>" -ViolationContent ($DeviceViolationContent + $coverageNote) -noViolationContent '' -HasViolation:$true
        Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount $knownDivisionCount -ViolationRecordCount $msimcViolations.Count -InputRecordCount $deviceData.Count -AnalyzedRecordCount $analyzedRecordCount -SkippedIdentityCount $skippedIdentityCount
    }
    $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU 'MSIMC' -RunId $RunId `
        -LedgerPath $LedgerPath -SentEmailsPath $sentEmailsPath `
        -From $prodid -To @($normalizedRecipients) -Cc ($CcContacter -join ';') -Subject $Subject `
        -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
    if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    Assert-NoBuMailRejections
    return
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
        $status = if ($record.Status -eq $usedStatusToken) { 'Used' } else { $record.Status }
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
    Save-AnalysisSummary -HasViolation $false -ViolationDivisionCount 0

    if ($BU -eq 'MSMS') {
        # noviolation of MSMS
        $BuContacter = $MsNoviolationRecipients
        $ledgerBu = 'MSMS-ALL'
    }
    else {
        # noviolation of MSBIC
        $BuContacter = $MsbicBuContacter
        $ledgerBu = 'MSBIC-ALL'
    }

    $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $ledgerBu -RunId $RunId `
        -LedgerPath $LedgerPath -SentEmailsPath $sentEmailsPath `
        -From $prodid -To $BuContacter -Cc $CcContacter -Subject $Subject `
        -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
    if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    Assert-NoBuMailRejections
    return
}
else {
    # violation records in MSMS
    if ($msViolationCollection.Count -gt 0) {
        $msmsViolationsByBU = $msViolationCollection | Group-Object 'Division' -AsHashTable
        Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount 3
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
                    -LedgerPath $LedgerPath -SentEmailsPath $sentEmailsPath `
                    -From $prodid -To $BuContacter -Cc $CcContacter -Subject $Subject `
                    -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
                if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
            }
            else {
                Write-Verbose $BuItem
                Write-Verbose "this BU has no violations, send normal mail"
                $htmlBody = New-HtmlBody -TableHtml "" -ViolationContent "" -NoViolationContent $NoViolationContent -HasViolation:$false
                $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $BuItem -RunId $RunId `
                    -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
                    -From $prodid -To $BuContacter -Cc $CcContacter -Subject $Subject `
                    -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
                if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
            }
        }
    }

    # violation records in MSBIC
    if ($msbicViolationCollection.Count -gt 0) {
        Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount 1 -AnalysisType 'Device'
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
            -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
        if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    }
}

Assert-NoBuMailRejections
