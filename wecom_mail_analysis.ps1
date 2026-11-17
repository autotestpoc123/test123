param(
    [Parameter(Mandatory = $true)]
    [string]$mailLogFilePath,
    [Parameter(Mandatory = $true)]
    [string]$startDate,
    [Parameter(Mandatory = $true)]
    [string]$endDate,
    [Parameter(Mandatory = $true)]
    [object]$AnalysisEnvironment,
    [string]$SummaryOutputPath,
    [string]$TaskOutputDirectory,
    [Parameter(Mandatory = $true)]
    [string]$TaskName,
    [Parameter(Mandatory = $true)]
    [string]$CycleId,
    [Parameter(Mandatory = $true)]
    [string]$LedgerPath,
    [Parameter(Mandatory = $true)]
    [string]$SentEmailsPath,
    [Parameter(Mandatory = $true)]
    [string]$RunId,
    [string]$BU,
    [string]$WeComSenderDomain,
    [string[]]$ReportRecipients,
    [hashtable]$IdentityMapping,
    [string]$IdentityMappingPath
)

$parentFolderPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$importModulePath = Join-Path $parentFolderPath 'wecom_analysis_comm.psm1'

if (-not (Test-Path $importModulePath)) {
    throw 'Module load path not found, please double check!'
}

Import-Module $importModulePath -Force
$script:rejected = [System.Collections.Generic.List[object]]::new()
$script:TaskName = $TaskName

function Assert-NoBuMailRejections {
    $count = $script:rejected.Count
    if ($count -gt 0) {
        $buList = ($script:rejected | ForEach-Object { $_.BU }) -join ', '
        throw "Task '$($script:TaskName)' rejected $count BU send due to content-diff: $buList."
    }
}

#Build Ldap query filte
function Build-OrFilter($batches) {
    $classBlock = '(|(objectClass=user)(objectClass=group))'
    $mailClauses = $batches | ForEach-Object { "(mail=$(ConvertTo-LdapFilterEscaped $_))" }
    $mailBlock = "(|" + ($mailClauses -join '') + ")"
    return "(&${classBlock}${mailBlock})"
}

#Batch query Ldap with users mail addresses
function Get-LdapUserByMail {
    param(
        [Parameter(Mandatory = $true)][System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]$LazyConnection,
        [Parameter(Mandatory = $true)]
        [string[]]$MailAddresses,
        [string]$SearchBase,
        [int]$SearchTimeoutSeconds = 30,
        [string[]]$attr = @("sAMAccountName", "mail", "division"),
        [int]$BatchSize = 20
    )

    Add-Type -AssemblyName System.DirectoryServices.Protocols
    #force creation /bind
    try {
        $conn = $lazyconnection.Value
    }
    catch {
        throw "failed to create LDAP Connection: $($_.Exception.Message)"
    }
    # set timeout
    $conn.Timeout = [TimeSpan]::FromSeconds($SearchTimeoutSeconds)

    # discover defaultNamingContext if SearchBase not provided
    if (-not $SearchBase) {
        $rootReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
            "", "(objectClass=*)",
            [System.DirectoryServices.Protocols.SearchScope]::Base,
            "defaultNamingContext"
        )
        $rootResp = $conn.SendRequest($rootReq)
        if ($rootResp.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success) {
            throw "Failed to read RootDSE: $($rootResp.ErrorMessage)"
        }
        $SearchBase = ($rootResp.Entries[0].Attributes["defaultNamingContext"])[0]
    }

    #initial maill lookup
    $mailLookup = @{}
    $valid = @{}
    $invalid = New-Object System.Collections.Generic.List[string]

    foreach ($batch in @($MailAddresses | ForEach-Object -Begin { $tmp = @() } -Process {
                $tmp += $_
                if ($tmp.count -gt $BatchSize) { , $tmp; $tmp = @() }
            } -End { if ($tmp.count -gt 0) { , $tmp } })) {

        $batchAddress = $batch.split(';').Trim()
        #build filter
        $filter = Build-OrFilter $batchAddress

        $searchReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
            $SearchBase,
            $filter,
            [System.DirectoryServices.Protocols.SearchScope]::Subtree,
            $attr
        )

        $searchResp = $conn.SendRequest($searchReq)

        if (-not $valid) { $valid = @{} }
        $foundMails = @()
        foreach ($entry in $searchResp.Entries) {
            $info = [PSCustomObject]@{
                SamAccountName = $null
                Mail           = $null
                Division       = $null
            }

            if ($entry.Attributes["mail"]) {
                $info.Mail = [string]$entry.Attributes["mail"][0].ToLower()
                $foundMails += $info.Mail
            }

            if ($entry.Attributes["sAMAccountName"]) {
                $info.SamAccountName = [string]$entry.Attributes["sAMAccountName"][0]
            }

            if ($entry.Attributes["division"]) {
                $info.Division = [string]$entry.Attributes["division"][0]
            }

            if ($info.Mail) {
                $mailLookup[$info.Mail.ToLower()] = $info
            }
        }

        foreach ($addr in $batchAddress) {
            if (-not ($foundMails -contains $addr)) {
                $invalid.Add($addr)
            }
        }
    }
    return @{
        Valid   = $mailLookup
        Invalid = $invalid
    }
}

function Save-AnalysisSummary {
    param(
        [bool]$HasViolation,
        [int]$ViolationDivisionCount,
        [string]$AnalysisType,
        [int]$ViolationRecordCount = 0,
        [int]$InputRecordCount = 0,
        [int]$AnalyzedRecordCount = 0,
        [int]$SkippedSourceCount = 0,
        [int]$SkippedIdentityCount = 0
    )

    if (-not $SummaryOutputPath) {
        return
    }

    $summary = [ordered]@{
        AnalysisType           = 'Mail'
        StartDate              = $startDate
        EndDate                = $endDate
        HasViolation           = $HasViolation
        ViolationDivisionCount = $ViolationDivisionCount
        ViolationRecordCount   = $ViolationRecordCount
        InputRecordCount       = $InputRecordCount
        AnalyzedRecordCount    = $AnalyzedRecordCount
        SkippedSourceCount     = $SkippedSourceCount
        SkippedIdentityCount   = $SkippedIdentityCount
    }
    if ($BU) {
        $summary.BusinessUnit = $BU
    }

    ([pscustomobject]$summary) | ConvertTo-Json -Depth 4 | Set-Content -Path $SummaryOutputPath -Encoding UTF8
}


#main procedure
$logFilePath = $null
$violationFlag = $false
$grouped = $null
$violationCollection = @()
$violationFile = "report.csv"
$contactDivision = @("Private Credit & Equity", "Real Assets", "Global Sales and Marketing")

# --resolve environment settings from config
$vaultEnv = $AnalysisEnvironment.VaultEnv
$prodid = $AnalysisEnvironment.ProdId
$idName = $AnalysisEnvironment.IdName
$vault_server = $AnalysisEnvironment.VaultServer
$internalDomain = $AnalysisEnvironment.InternalDomainPattern
$txmailFilter = $AnalysisEnvironment.TxmailFilterPattern
$codDomain = $AnalysisEnvironment.CodDomainSuffix
$ldapServer = $AnalysisEnvironment.LdapServer
$smtp_server = $AnalysisEnvironment.SmtpServer
$smtp_port = $AnalysisEnvironment.Port
$domain = $AnalysisEnvironment.Domain

$CcContacter = $AnalysisEnvironment.Cc
$NoViolationRecipients = $AnalysisEnvironment.MailNoviolationRecipients
$BURecipients = $AnalysisEnvironment.MailBURecipients

Write-Verbose $ldapServer
Write-Verbose $vault_Server
Write-Verbose $prodid
Write-Verbose $domain

try {
    $null = Convert-ExactDate $startDate
    $null = Convert-ExactDate $endDate
}
catch {
    throw 'Invalid date format. Both startDate and endDate must be in yyyyMMdd format.'
}

$destFolder = Export-AnalysisReport -LogFilePath $mailLogFilePath -TaskOutputDirectory $TaskOutputDirectory -SubFolder 'analyzed' -UseDateSubFolder
$logFilePath = if ($TaskOutputDirectory) {
    Join-Path $destFolder 'task.log'
}
else {
    Get-LogFilePath -Directory $destFolder -BaseName "AnalysisLog"
}

$endDateDisplay = [datetime]::ParseExact($endDate, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture).AddDays(-1).ToString('yyyy-MM-dd')

$isMsimc = ($BU -eq 'MSIMC')
if ($isMsimc) {
    $subject = "COD WeCom MSIMC Mail Data Leakage Manual Review - from ${startDate} to ${endDateDisplay}"
}
else {
    $subject = "COD WeCom Mail Data Leakage Manual Review - from ${startDate} to ${endDateDisplay}"
}
# SEC-005 (accepted risk): violation notification emails intentionally include
# employee PII (names, accounts, email subjects) in the table below. This data
# is already registered in the firm DLM inventory, recipients are controlled
# BU contacts, and masking is not required. See SECURITY_REVIEW.md, SEC-005.
$mailNoViolationContent = "The purpose of this email is to provide information on users in your business unit(BU) that have sent emails from WeCom Mail to both a non-MS email address and an internal MS email address during the reporting period.<br><br/>
                           This manual review identifies instances where emails were sent from WeCom Mail to both a non-MS email address and an internal MS email address during the reporting period.<br><br/>
                           There were <b>no violations</b> to report this reporting period for your BU.<br><br/>"

$mailViolationContent = "The purpose of this email is to provide information on users in your business unit(BU) that have sent emails from WeCom Mail to both a non-MS email address and an internal MS email address during the reporting period.<br><br/>
                           This manual review identifies instances where emails were sent from WeCom Mail to both a non-MS email address and an internal MS email address during the reporting period.<br><br/>
                           The violation record(s) of this reporting period for your BU as below:"
try {
    $csvData = @(Import-Csv -Path $mailLogFilePath -Encoding UTF8 -Delimiter ',')
}
catch {
    throw "Failed to import csv file: $mailLogFilePath"
}

if (-not $csvData) {
    throw "Input mail log file is empty: $mailLogFilePath"
}

if ($isMsimc) {
    if (-not $WeComSenderDomain -or $WeComSenderDomain -notmatch '^[A-Za-z0-9.-]+$') {
        throw 'MSIMC mail analysis requires a valid WeComSenderDomain.'
    }
    $normalizedRecipients = @(
        @($ReportRecipients) |
            ForEach-Object { if ($_) { $_.Trim().ToLowerInvariant() } } |
            Where-Object { $_ } |
            Select-Object -Unique
    )
    if ($normalizedRecipients.Count -eq 0) {
        throw 'MSIMC mail analysis requires at least one report recipient.'
    }

    function Split-MailAddresses {
        param([string]$Value)
        if (-not $Value) { return @() }
        return @(
            $Value -split '[;, ]+' |
                ForEach-Object { $_.Trim().Trim('<>').ToLowerInvariant() } |
                Where-Object { $_ -match '^[^@\s]+@[^@\s]+\.[^@\s]+$' } |
                Select-Object -Unique
        )
    }

    $mailTime = -join (@(0x65F6,0x95F4) | ForEach-Object { [char]$_ })
    $mailSubject = -join (@(0x4E3B,0x9898) | ForEach-Object { [char]$_ })
    $mailSender = -join (@(0x53D1,0x4EF6,0x4EBA) | ForEach-Object { [char]$_ })
    $mailReceiver = -join (@(0x6536,0x4EF6,0x4EBA) | ForEach-Object { [char]$_ })
    $mailStatus = -join (@(0x72B6,0x6001) | ForEach-Object { [char]$_ })
    $mailBounced = -join (@(0x5DF2,0x9000,0x4FE1) | ForEach-Object { [char]$_ })
    $senderDomainPattern = '@' + [regex]::Escape($WeComSenderDomain) + '$'
    $inputRecordCount = @($csvData).Count
    $skippedSourceCount = @($csvData | Where-Object { ([string]$_.$mailSender).Trim() -notmatch $senderDomainPattern }).Count
    $scopedRows = @($csvData | Where-Object { ([string]$_.$mailSender).Trim() -match $senderDomainPattern })
    if ($scopedRows.Count -eq 0) {
        throw "MSIMC mail log has no senders in domain '@$WeComSenderDomain'."
    }
    if (-not $IdentityMapping) {
        if (-not $IdentityMappingPath) { $IdentityMappingPath = Join-Path (Split-Path $parentFolderPath -Parent) 'convert_table.xlsx' }
        $IdentityMapping = Import-MsimcIdentityMapping -Path $IdentityMappingPath
    }
    $senderAddressesToResolve = @(
        $scopedRows | ForEach-Object { Split-MailAddresses $_.$mailSender } |
            Where-Object { $_ -match $senderDomainPattern } |
            Select-Object -Unique
    )
    $senderResolutions = Resolve-MsimcIdentity -Identities $senderAddressesToResolve -Kind Mail -Mapping $IdentityMapping -MailDomainSuffix $codDomain
    $senderCandidates = @($senderResolutions.Values | ForEach-Object { $_.CorporateEmail } | Sort-Object -Unique)
    $receiverCandidates = @(
        $scopedRows | ForEach-Object { if ($_.$mailReceiver) { $_.$mailReceiver -split ';' } } |
            ForEach-Object { $_.Trim().ToLowerInvariant() } |
            Where-Object { $_ -and $_ -match $internalDomain } |
            Select-Object -Unique
    )

    $sysid_cert = Get-Cert -KeyName $prodid
    $vault_secret = Get-VaultSecret -VaultEnv $vaultEnv -KeyName $prodid -sysIdCert $sysid_cert -VaultServer $vault_server
    $id_secret = New-Object System.Net.NetworkCredential($idName, $vault_secret, $Domain)
    $lazyConn = New-LazyLdapConnection -Server $ldapServer -Port 636 -Credential $id_secret
    try {
        $recevierslookup = if ($receiverCandidates.Count -gt 0) {
            Get-LdapUserByMail -LazyConnection $lazyConn -MailAddresses ($receiverCandidates -join ';')
        }
        else {
            @{ Valid = @{}; Invalid = @() }
        }
        $sendersLookup = Get-LdapUserByMail -LazyConnection $lazyConn -MailAddresses ($senderCandidates -join ';')
    }
    finally {
        Close-LazyLdapConnection -lazy $lazyConn
    }


    $grouped = $scopedRows | Group-Object -Property $mailTime, $mailSubject, $mailSender | ForEach-Object {
        $item = New-Object psobject
        $item | Add-Member NoteProperty $mailTime $_.Group[0].$mailTime
        $item | Add-Member NoteProperty $mailSubject $_.Group[0].$mailSubject
        $item | Add-Member NoteProperty $mailSender $_.Group[0].$mailSender
        $item | Add-Member NoteProperty $mailReceiver ((@($_.Group | ForEach-Object { if ($_.$mailReceiver) { $_.$mailReceiver -split ';' } }) | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ';')
        $item | Add-Member NoteProperty $mailStatus $(if ($_.Group | Where-Object { $_.$mailStatus -match $mailBounced }) {'Partially Delivered'} else {'Delivered Successfully'})
        $item
    }

    $skippedIdentityCount = 0
    $analyzedRecordCount = 0
    foreach ($item in @($grouped | Sort-Object $mailTime, $mailSender, $mailSubject)) {
        $senderAddresses = @(Split-MailAddresses $item.$mailSender | Where-Object { $_ -match $senderDomainPattern })
        if ($senderAddresses.Count -eq 0) { continue }
        $mappedSender = $senderResolutions[$senderAddresses[0]].CorporateEmail
        $senderIdentity = $sendersLookup.Valid[$mappedSender]
        if (-not $senderIdentity) {
            $skippedIdentityCount++
            Write-Log -LogString "MSIMC sender identity skipped: $($item.$mailSender)" -LogFilePath $logFilePath
            continue
        }
        $analyzedRecordCount++
        $addresses = @($item.$mailReceiver -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $external = @($addresses | Where-Object { $_ -notmatch $internalDomain })
        $internal = @($addresses | Where-Object { $_ -match $internalDomain })
        $invalidInternal = @($internal | Where-Object { @($recevierslookup.Invalid) -contains $_ })
        if ($external.Count -gt 0 -and $internal.Count -gt 0 -and $invalidInternal.Count -eq $internal.Count) {
            $violationCollection += [PSCustomObject]@{
                DateTime   = $item.$mailTime
                SendType   = 'Send'
                Subject    = $item.$mailSubject
                Sender     = $item.$mailSender
                Recipients = ($addresses -join ';')
                Status     = $item.$mailStatus
                Division   = [string]$senderIdentity.Division
            }
        }
    }

    if ($analyzedRecordCount -eq 0) {
        throw "MSIMC mail analysis could not map any sender identity. SkippedIdentityCount=$skippedIdentityCount."
    }

    $violationCollection = @($violationCollection | Sort-Object DateTime, Sender, Subject, Recipients)
    $knownDivisionCount = @($violationCollection | Where-Object { $_.Division } | Select-Object -ExpandProperty Division -Unique).Count
    $destFilePath = Join-Path $destFolder $violationFile
    $coverageNote = "<br/>Coverage: input=$inputRecordCount; analyzed=$analyzedRecordCount; skippedSource=$skippedSourceCount; skippedIdentity=$skippedIdentityCount."
    if ($violationCollection.Count -eq 0) {
        Save-AnalysisSummary -HasViolation $false -ViolationDivisionCount 0 -ViolationRecordCount 0 -InputRecordCount $inputRecordCount -AnalyzedRecordCount $analyzedRecordCount -SkippedSourceCount $skippedSourceCount -SkippedIdentityCount $skippedIdentityCount
        $htmlBody = New-HtmlBody -TableHtml '' -ViolationContent '' -NoViolationContent ($mailNoViolationContent + $coverageNote) -HasViolation:$false
    }
    else {
        $violationCollection | Export-Csv -Path $destFilePath -NoTypeInformation -Encoding UTF8 -Force
        $rowsHtml = $violationCollection | Select-Object DateTime, Subject, Sender, Recipients, Status, Division | ConvertTo-Html -Fragment 
        $htmlBody = New-HtmlBody -TableHtml "<table>$rowsHtml</table>" -ViolationContent ($mailViolationContent + $coverageNote) -NoViolationContent '' -HasViolation:$true
        Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount $knownDivisionCount -ViolationRecordCount $violationCollection.Count -InputRecordCount $inputRecordCount -AnalyzedRecordCount $analyzedRecordCount -SkippedSourceCount $skippedSourceCount -SkippedIdentityCount $skippedIdentityCount
    }
    $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU 'MSIMC' -RunId $RunId `
        -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
        -From $prodid -To @($normalizedRecipients) -Cc ($CcContacter -join ';') -Subject $subject `
        -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
    if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    Assert-NoBuMailRejections
    return
}

$legacyTime = -join (@(0x65F6,0x95F4) | ForEach-Object { [char]$_ })
$legacySubject = -join (@(0x4E3B,0x9898) | ForEach-Object { [char]$_ })
$legacySender = -join (@(0x53D1,0x4EF6,0x4EBA) | ForEach-Object { [char]$_ })
$legacyReceiver = -join (@(0x6536,0x4EF6,0x4EBA) | ForEach-Object { [char]$_ })
$legacyStatus = -join (@(0x72B6,0x6001) | ForEach-Object { [char]$_ })
$legacyBounced = -join (@(0x5DF2,0x9000,0x4FE1) | ForEach-Object { [char]$_ })
$receviers = $csvData | Select-Object -ExpandProperty $legacyReceiver | ForEach-Object { $_.Trim().Tolower() } | Where-Object{ $_ -match $internalDomain } | Select-Object -Unique
$recevierAddrs = $receviers -join ';'
$senders = $csvData | Select-Object -ExpandProperty $legacySender | ForEach-Object { $_.Trim().Tolower() } | Where-Object { $_ -match $internalDomain } | Select-Object -Unique
$senderAddrs = ($senders | ForEach-Object { $_.Trim().ToLower() -replace '@.*$', $codDomain }) -join ';'
if (($null -eq $receviers) -or ($null -eq $senders)) { throw 'Receivers or Senders addresses do not match expected domain pattern' }
$sysid_cert = Get-Cert -KeyName $prodid
$vault_secret = Get-VaultSecret -VaultEnv $vaultEnv -KeyName $prodid -sysIdCert $sysid_cert -VaultServer $vault_server
$id_secret = New-Object System.Net.NetworkCredential($idName, $vault_secret, $Domain)
$lazyConn = New-LazyLdapConnection -Server $ldapServer -Port 636 -Credential $id_secret
$recevierslookup = Get-LdapUserByMail -LazyConnection $lazyConn -MailAddresses $recevierAddrs
$sendersLookup = Get-LdapUserByMail -LazyConnection $lazyConn -MailAddresses $senderAddrs
Close-LazyLdapConnection -lazy $lazyConn
$grouped = $csvData | Group-Object -Property $legacyTime, $legacySubject, $legacySender | ForEach-Object {
    $item = New-Object psobject
    $item | Add-Member NoteProperty $legacyTime $_.Group[0].$legacyTime
    $item | Add-Member NoteProperty $legacySubject $_.Group[0].$legacySubject
    $item | Add-Member NoteProperty $legacySender $_.Group[0].$legacySender
    $item | Add-Member NoteProperty $legacyReceiver (($_.Group | Select-Object -ExpandProperty $legacyReceiver | Where-Object{ $_ -and $_.Trim() -ne ''} | Sort-Object -Unique) -join ';')
    $item | Add-Member NoteProperty $legacyStatus $(if ($_.Group | Where-Object { $_.$legacyStatus -match $legacyBounced }) {'Partially Delivered'} else {'Delivered Successfully'})
    $item
}
foreach ($item in $grouped) {
    if (-not $item.$legacyReceiver -or -not $item.$legacyReceiver.Contains(';')) { continue }
    $groupRec = @($item.$legacyReceiver -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $internalAddr = @($groupRec | Where-Object { $_ -match $internalDomain })
    $hasExternal = @($groupRec | Where-Object { $_ -notmatch $internalDomain }).Count -gt 0
    $invalidInternals = @($internalAddr | Where-Object { $recevierslookup.Invalid.Contains($_) })
    if ($hasExternal -and $invalidInternals.Count -gt 0 -and $invalidInternals.Count -eq $internalAddr.Count -and $item.$legacySender -match $txmailFilter) {
        $wecomAddr = $item.$legacySender -replace '@.*$', $codDomain
        $mailAddr = $sendersLookup.Valid[$wecomAddr]
        if ($mailAddr -and $mailAddr.Division -and $contactDivision -contains $mailAddr.Division) {
            $violationFlag = $true
            $violationCollection += [PSCustomObject]@{ DateTime = $item.$legacyTime; SendType = 'Send'; Subject = $item.$legacySubject; sender = $item.$legacySender; Recipients = $item.$legacyReceiver; Status = $item.$legacyStatus; Division = $mailAddr.Division}
        }
    }
}

$destFilePath = Join-Path $destFolder $violationFile

$violationCollection = @($violationCollection)
if ((-not $violationFlag) -or $violationCollection.Count -eq 0) {
    Write-Host 'No violation happened' -ForegroundColor Green
    if ($violationFlag -and $violationCollection.Count -eq 0) {
        Write-Log -LogString 'Potential violation pattern found, but no records matched configured division scope. Treat as no violation.' -LogFilePath $logFilePath
    }
    Write-Host "No violation happened" -ForegroundColor Green
    Write-Log -LogString "Completed mail log analysis without violation founded" -LogFilePath $logFilePath
    Save-AnalysisSummary -HasViolation $false -violationDivisionCount 0 -AnalysisType 'Mail'
    $htmlBody = New-HtmlBody -TableHtml "" -ViolationContent "" -NoViolationContent $mailNoViolationContent -HasViolation:$false
    $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU 'ALL' -RunId $RunId `
        -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
        -From $prodid -To $NoViolationRecipients -Cc ($ccContacter -join ';') -Subject $subject `
        -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
    if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    Assert-NoBuMailRejections
    return
}

Write-Host "mail violations file path: $destFilePath"
Write-Log -LogString "Completed mail log analysis with violation founded" -LogFilePath $logFilePath
$violationCollection | Export-Csv -Path $destFilePath -NoTypeInformation -Encoding UTF8 -Force
$violationsByBU = $violationCollection | Group-Object 'Division' -AsHashTable
Save-AnalysisSummary -HasViolation $true -violationDivisionCount 3 -AnalysisType 'Mail'
foreach ($bu in $BURecipients.keys) {
    $BuContacter = $BURecipients[$bu]
    $hasViolation = $violationsByBU.ContainsKey($bu)
    if ($hasViolation) {
        write-verbose $bu
        $rowsHtml = $violationsByBU[$bu] | Select-Object `
            'DateTime', 'Subject', 'Sender', 'Recipients', 'Status', 'Division' | ConvertTo-Html -Fragment
        $tableHtml = @"
<table>
    $rowsHtml
</table>
"@
        $htmlBody = New-HtmlBody -TableHtml $tableHtml -ViolationContent $mailViolationContent -NoViolationContent "" -HasViolation:$true
        $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $bu -RunId $RunId `
            -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
            -From $prodid -To $BuContacter -Cc ($CcContacter -join ';') -Subject $subject `
            -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
        if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    }
    else {
        Write-Verbose $bu
        Write-verbose "this BU has no violations, send normal mail"
        $htmlBody = New-HtmlBody -TableHtml "" -ViolationContent "" -NoViolationContent $mailNoViolationContent -HasViolation:$false
        $sendResult = Send-AuditBuMail -Cycle $CycleId -TaskName $TaskName -BU $bu -RunId $RunId `
            -LedgerPath $LedgerPath -SentEmailsPath $SentEmailsPath `
            -From $prodid -To $BuContacter -Cc ($CcContacter -join ';') -Subject $subject `
            -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert $sysid_cert -Port $smtp_port -LogFilePath $logFilePath
        if ($sendResult.Result -eq 'Rejected') { [void]$rejected.Add($sendResult) }
    }
}

Assert-NoBuMailRejections
