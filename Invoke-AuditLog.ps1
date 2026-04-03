param(
    [Parameter(Mandatory = $true)]
    [string]$mailLogFilePath,
    [Parameter(Mandatory = $true)]
    [string]$startDate,
    [Parameter(Mandatory = $true)]
    [string]$endDate,
    [string]$env = 'QA',
    [string]$SummaryOutputPath,
    [string]$TaskOutputDirectory
)

$parentFolderPath = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$importModulePath = Join-Path $parentFolderPath 'wecom_analysis_comm.psm1'

if (-not (Test-Path $importModulePath)) {
    throw 'Module load path not found, please double check!'
}

Import-Module $importModulePath -Force

function Export-AnalysisReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    if ($TaskOutputDirectory) {
        if (-not (Test-Path $TaskOutputDirectory)) {
            New-Item -Path $TaskOutputDirectory -ItemType Directory -Force | Out-Null
        }
        return $TaskOutputDirectory
    }

    $baseFolder = Split-Path -Parent $LogFilePath
    $reportFolder = Join-Path $baseFolder 'AnalysisReport'
    if (-not (Test-Path $reportFolder)) {
        New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
    }
    return $reportFolder
}

function Resolve-MailColumns {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$CsvData
    )

    $propertyNames = @($CsvData[0].PSObject.Properties.Name)
    $cnTime = ([string][char]0x65F6) + [char]0x95F4
    $cnSendTime = ([string][char]0x53D1) + [char]0x9001 + [char]0x65F6 + [char]0x95F4
    $cnSubject = ([string][char]0x4E3B) + [char]0x9898
    $cnMailSubject = ([string][char]0x90AE) + [char]0x4EF6 + [char]0x4E3B + [char]0x9898
    $cnSender = ([string][char]0x53D1) + [char]0x4EF6 + [char]0x4EBA
    $cnSendUser = ([string][char]0x53D1) + [char]0x9001 + [char]0x4EBA
    $cnRecipient = ([string][char]0x6536) + [char]0x4EF6 + [char]0x4EBA
    $cnReceiveUser = ([string][char]0x63A5) + [char]0x6536 + [char]0x4EBA
    $cnStatus = ([string][char]0x72B6) + [char]0x6001
    $cnDeliveryStatus = ([string][char]0x6295) + [char]0x9012 + [char]0x72B6 + [char]0x6001

    $aliasMap = @{
        Time      = @($cnTime, $cnSendTime, 'Time', 'DateTime', 'SentTime')
        Subject   = @($cnSubject, $cnMailSubject, 'Subject', 'Title')
        Sender    = @($cnSender, $cnSendUser, 'Sender', 'From')
        Recipient = @($cnRecipient, $cnReceiveUser, 'Recipients', 'Recipient', 'To')
        Status    = @($cnStatus, $cnDeliveryStatus, 'Status', 'DeliveryStatus', 'Result')
    }

    $resolved = @{}
    foreach ($logicalName in $aliasMap.Keys) {
        $column = $propertyNames | Where-Object { $_ -in $aliasMap[$logicalName] } | Select-Object -First 1
        if (-not $column) {
            throw "Unable to resolve required mail log column: $logicalName. Available columns: $($propertyNames -join ', ')"
        }

        $resolved[$logicalName] = $column
    }

    return $resolved
}

function Get-RowValue {
    param(
        [Parameter(Mandatory = $true)]
        [psobject]$Row,
        [Parameter(Mandatory = $true)]
        [string]$ColumnName
    )

    return [string]$Row.PSObject.Properties[$ColumnName].Value
}

function Save-AnalysisSummary {
    param(
        [bool]$HasViolation,
        [int]$ViolationDivisionCount,
        [string]$AnalysisType
    )

    if (-not $SummaryOutputPath) {
        return
    }

    $summary = [PSCustomObject]@{
        AnalysisType           = 'Mail'
        StartDate              = $startDate
        EndDate                = $endDate
        HasViolation           = $HasViolation
        ViolationDivisionCount = $ViolationDivisionCount
    }

    $summary | ConvertTo-Json -Depth 4 | Set-Content -Path $SummaryOutputPath -Encoding UTF8
}

switch ($env) {
    'PROD' {
        $smtp_server = 'mailhost.ms.com'
        $ccContacter = @('ling.gu@infradev.mocktest.com.cn')
        $prodid = 'wecom-audit-prod'
        $sysid_cert = 'wecom-audit-prod-cert'
        $internalDomain = [regex]'@mocktest\.com$'
        $txmailFilter = [regex]'txmail'
        $codDomain = '@mocktest.com'
        $contactDivision = @(
            'Morgan Stanley China',
            'Morgan Stanley International Bank (China) Limited',
            'Fixed Income Division'
        )
        $BURecipients = @{
            'Morgan Stanley China' = @('ling.gu@infradev.mocktest.com.cn')
            'Morgan Stanley International Bank (China) Limited' = @('ling.gu@infradev.mocktest.com.cn')
            'Fixed Income Division' = @('ling.gu@infradev.mocktest.com.cn')
        }
        $noViolationRecipients = @('ling.gu@infradev.mocktest.com.cn')
    }
    'QA' {
        $smtp_server = 'mailhost.ms.com'
        $ccContacter = @('ling.gu@infradev.mocktest.com.cn')
        $prodid = 'wecom-audit-qa'
        $sysid_cert = 'wecom-audit-qa-cert'
        $internalDomain = [regex]'@mocktest\.com$'
        $txmailFilter = [regex]'txmail'
        $codDomain = '@mocktest.com'
        $contactDivision = @(
            'Morgan Stanley China',
            'Morgan Stanley International Bank (China) Limited',
            'Fixed Income Division'
        )
        $BURecipients = @{
            'Morgan Stanley China' = @('ling.gu@infradev.mocktest.com.cn')
            'Morgan Stanley International Bank (China) Limited' = @('ling.gu@infradev.mocktest.com.cn')
            'Fixed Income Division' = @('ling.gu@infradev.mocktest.com.cn')
        }
        $noViolationRecipients = @('ling.gu@infradev.mocktest.com.cn')
    }
    default {
        throw "Unknown env: $env"
    }
}

Write-Verbose $prodid
Write-Verbose $codDomain

try {
    $null = Convert-ExactDate $startDate
    $null = Convert-ExactDate $endDate
}
catch {
    throw 'Invalid date format. Both startDate and endDate must be in yyyyMMdd format.'
}

$destFolder = Export-AnalysisReport -LogFilePath $mailLogFilePath
$logFilePath = if ($TaskOutputDirectory) {
    Join-Path $destFolder 'task.log'
}
else {
    Get-LogFilePath -Directory $destFolder -BaseName 'AnalysisLog'
}
$subject = "COD WeCom Mail Data Leakage Manual Review - from ${startDate} to ${endDate}"
$mailNoViolationContent = 'The purpose of this email is to provide information on users in your business unit(BU) that have sent emails externally via WeCom Mail for the reporting period. There were <b>no violations</b> to report this reporting period for your BU.<br/><br/>'
$mailViolationContent = 'The purpose of this email is to provide information on users in your business unit(BU) that have sent emails externally via WeCom Mail for the reporting period. Violations were sent from WeCom Mail to both a non-MS email and an invalid MS email, as these cases are not allowed. The violation record(s) of this reporting period for your BU as below:'

try {
    $csvData = @(Import-Csv -Path $mailLogFilePath -Encoding UTF8 -Delimiter ',')
}
catch {
    throw "Failed to import csv file: $mailLogFilePath"
}

if (-not $csvData) {
    throw "Input mail log file is empty: $mailLogFilePath"
}

$columnMap = Resolve-MailColumns -CsvData $csvData

$receivers = @(
    $csvData |
        ForEach-Object { Get-RowValue -Row $_ -ColumnName $columnMap.Recipient } |
        ForEach-Object {
            if ($_ -is [string]) {
                $_ -split ';|,| ' | Where-Object { $_ }
            }
        } |
        ForEach-Object { $_.Trim().ToLower() } |
        Sort-Object -Unique
)

$senders = @(
    $csvData |
        ForEach-Object { Get-RowValue -Row $_ -ColumnName $columnMap.Sender } |
        ForEach-Object {
            if ($_ -is [string]) {
                $_ -split ';|,| ' | Where-Object { $_ }
            }
        } |
        ForEach-Object { $_.Trim().ToLower() } |
        Sort-Object -Unique
)

$receiverAddrs = @($receivers | Where-Object { $_ -match $internalDomain })
$senderAddrs = @($senders | Where-Object { $_ -match $txmailFilter } | ForEach-Object { $_ -replace '@.*$', $codDomain })

$ldapConn = [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]::new({
    $server = 'localhost'
    $ldap = [System.DirectoryServices.Protocols.LdapConnection]::new($server)
    $ldap.AuthType = [System.DirectoryServices.Protocols.AuthType]::Negotiate
    $ldap.SessionOptions.ProtocolVersion = 3
    $ldap.Bind()
    return $ldap
})

$recevierslookup = Get-LdapUserByMail -lazyConnection $ldapConn -mailAdds $receiverAddrs
$sendersLookup = Get-LdapUserByMail -lazyConnection $ldapConn -mailAdds $senderAddrs

if (@($receivers).Count -eq 0 -or @($senders).Count -eq 0) {
    Write-Log -LogString 'Receivers or senders are empty after normalization.' -LogFilePath $logFilePath
    Write-Host 'Receivers or senders are empty after normalization, please check the mail addresses.' -ForegroundColor Yellow
    Save-AnalysisSummary -HasViolation $false -ViolationDivisionCount 0 -AnalysisType 'Mail'
    return
}

$groupedMap = @{}
foreach ($row in $csvData) {
    $timeValue = Get-RowValue -Row $row -ColumnName $columnMap.Time
    $subjectValue = Get-RowValue -Row $row -ColumnName $columnMap.Subject
    $senderValue = ((Get-RowValue -Row $row -ColumnName $columnMap.Sender) -split ';|,| ' | Where-Object { $_ } | Select-Object -First 1) -as [string]
    $recipientList = @((Get-RowValue -Row $row -ColumnName $columnMap.Recipient) -split ';|,| ' | Where-Object { $_ } | ForEach-Object { $_.Trim().ToLower() })
    $statusValue = Get-RowValue -Row $row -ColumnName $columnMap.Status
    if (-not $senderValue) {
        continue
    }

    $senderValue = $senderValue.Trim().ToLower()
    $groupKey = "$timeValue|$subjectValue|$senderValue"
    if (-not $groupedMap.ContainsKey($groupKey)) {
        $groupedMap[$groupKey] = [PSCustomObject]@{
            Time       = $timeValue
            Subject    = $subjectValue
            Sender     = $senderValue
            Recipients = @()
            Status     = $statusValue
        }
    }

    $groupedMap[$groupKey].Recipients += $recipientList
}

$violationFlag = $false
$violationCollection = @()
$destFilePath = Join-Path $destFolder 'report.csv'

foreach ($item in $groupedMap.Values) {
    $item.Recipients = @($item.Recipients | Sort-Object -Unique)
    $groupRec = @($item.Recipients)
    if ($groupRec.Count -eq 0) {
        continue
    }

    $nonInternal = @($groupRec | Where-Object { $_ -notmatch $internalDomain })
    $internalAddr = @($groupRec | Where-Object { $_ -match $internalDomain })
    if (@($nonInternal).Count -gt 0) {
        $invalidInternals = @($internalAddr | Where-Object { $recevierslookup.Invalid.Contains($_) })
        Write-Log -LogString "The records contain external domain are : $($groupRec -join ';')" -LogFilePath $logFilePath

        if (@($internalAddr).Count -gt 0 -and @($invalidInternals).Count -eq @($internalAddr).Count) {
            $violationFlag = $true

            if ($item.Sender -match $txmailFilter) {
                $wecomAddr = $item.Sender -replace '@.*$', $codDomain
                if ($sendersLookup.Valid[$wecomAddr]) {
                    $division = $sendersLookup.Valid[$wecomAddr].Division
                    if ($division -and $contactDivision -contains $division) {
                        $violationCollection += [PSCustomObject]@{
                            DateTime   = $item.Time
                            Subject    = $item.Subject
                            Sender     = $item.Sender
                            Recipients = ($item.Recipients -join ';')
                            Status     = $item.Status
                            Division   = $division
                        }
                    }
                }
            }
        }
    }
}

if (-not $violationFlag) {
    Write-Host 'No violation happened' -ForegroundColor Green
    Write-Log -LogString 'Completed mail log analysis without violation founded' -LogFilePath $logFilePath
    Save-AnalysisSummary -HasViolation $false -ViolationDivisionCount 0 -AnalysisType 'Mail'
    $htmlBody = New-HtmlBody -TableHtml '' -ViolationContent '' -NoViolationContent $mailNoViolationContent -HasViolation $false
    Send-Mail -From $prodid -To $noViolationRecipients -Cc ($ccContacter -join ';') -Subject $subject `
        -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert (Get-Cert -KeyName $sysid_cert) -Port 2587 -LogFilePath $logFilePath
    return
}

Write-Host "mail violations file path: $destFilePath"
Write-Log -LogString 'Completed mail log analysis with violation founded' -LogFilePath $logFilePath
$violationCollection | Export-Csv -Path $destFilePath -NoTypeInformation -Encoding UTF8 -Force
$violationsByBU = $violationCollection | Group-Object 'Division' -AsHashTable
Save-AnalysisSummary -HasViolation $true -ViolationDivisionCount $violationsByBU.Keys.Count -AnalysisType 'Mail'

foreach ($bu in $BURecipients.Keys) {
    $BuContacter = $BURecipients[$bu]
    $hasViolation = $violationsByBU.ContainsKey($bu)
    if ($hasViolation) {
        Write-Verbose $bu
        $rowsHtml = $violationsByBU[$bu] | Select-Object `
            'DateTime', 'Subject', 'Sender', 'Recipients', 'Status', 'Division' | ConvertTo-Html -Fragment
        $tableHtml = @"
<table>
    $rowsHtml
</table>
"@
        $htmlBody = New-HtmlBody -TableHtml $tableHtml -ViolationContent $mailViolationContent -NoViolationContent '' -HasViolation $true
        Send-Mail -From $prodid -To $BuContacter -Cc ($ccContacter -join ';') -Subject $subject `
            -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert (Get-Cert -KeyName $sysid_cert) -Port 2587 -LogFilePath $logFilePath
    }
    else {
        Write-Verbose $bu
        Write-Verbose 'this BU has no violations, send normal mail'
        $htmlBody = New-HtmlBody -TableHtml '' -ViolationContent '' -NoViolationContent $mailNoViolationContent -HasViolation $false
        Send-Mail -From $prodid -To $BuContacter -Cc ($ccContacter -join ';') -Subject $subject `
            -Body $htmlBody -SmtpServer $smtp_server -KeyName $prodid -Cert (Get-Cert -KeyName $sysid_cert) -Port 2587 -LogFilePath $logFilePath
    }
}
