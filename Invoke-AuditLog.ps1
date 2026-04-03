if (-not ('System.DirectoryServices.Protocols.LdapConnection' -as [type])) {
    Add-Type -AssemblyName System.DirectoryServices.Protocols
}

<#
.SYNOPSIS
English code-review note for function 'Convert-ExactDate'.
.DESCRIPTION
Converts input data into a normalized output format used by the workflow.
#>
function Convert-ExactDate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DateText
    )

    try {
        return [datetime]::ParseExact($DateText, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
    }
    catch {
        throw "Invalid date format '$DateText'. Expected yyyyMMdd."
    }
}

<#
.SYNOPSIS
English code-review note for function 'Write-Log'.
.DESCRIPTION
Writes workflow artifacts to disk for traceability and downstream consumption.
#>
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogString,
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    $time = Get-Date -Format 'yyyy/MM/dd HH:mm:ss'
    "$time - $LogString" | Out-File -FilePath $LogFilePath -Width 1024 -Append -Encoding UTF8
}

<#
.SYNOPSIS
English code-review note for function 'Get-LogFilePath'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-LogFilePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Directory,
        [Parameter(Mandatory = $true)]
        [string]$BaseName
    )

    if (-not (Test-Path -Path $Directory)) {
        New-Item -ItemType Directory -Force -Path $Directory | Out-Null
    }

    $logDate = Get-Date -Format 'yyyyMMdd_HHmmss'
    return (Join-Path $Directory "$BaseName.$logDate.log")
}

<#
.SYNOPSIS
English code-review note for function 'New-DateTokenMap'.
.DESCRIPTION
Creates a new object or structure used by subsequent processing steps.
#>
function New-DateTokenMap {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate
    )

    $startDateValue = Convert-ExactDate $StartDate
    $endDateValue = Convert-ExactDate $EndDate

    return @{
        startDate            = $StartDate
        endDate              = $EndDate
        startDateMMdd        = $StartDate.Substring($StartDate.Length - 4)
        endDateMMdd          = $EndDate.Substring($EndDate.Length - 4)
        endDatePlus1         = $endDateValue.AddDays(1).ToString('yyyyMMdd')
        endDatePlus1MMdd     = $endDateValue.AddDays(1).ToString('MMdd')
        startDate_EndDate    = "${StartDate}_${EndDate}"
        startDateDashEndDate = "${StartDate}-${EndDate}"
    }
}

<#
.SYNOPSIS
English code-review note for function 'Resolve-TemplateText'.
.DESCRIPTION
Resolves runtime values from configuration, tokens, and current execution context.
#>
function Resolve-TemplateText {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Template,
        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens
    )

    $resolved = $Template
    foreach ($key in $Tokens.Keys) {
        $resolved = $resolved.Replace("{$key}", [string]$Tokens[$key])
    }

    return $resolved
}

<#
.SYNOPSIS
Validates configured input paths before task execution starts.
.DESCRIPTION
Checks InputRoot and all configured task InputDirectory values (after token resolution)
and throws one aggregated, readable error if any path is missing.
#>
function Assert-ConfigInputDirectories {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens,
        [Parameter(Mandatory = $true)]
        [string]$ConfigPath
    )

    $issues = New-Object 'System.Collections.Generic.List[string]'
    $resolvedInputRoot = if ($Tokens.ContainsKey('inputRoot')) { [string]$Tokens.inputRoot } else { $null }

    if (-not $resolvedInputRoot) {
        $issues.Add("InputRoot is empty. Set 'InputRoot' in config or WECOM_AUDIT_INPUT_ROOT.")
    }
    elseif (-not (Test-Path -LiteralPath $resolvedInputRoot -PathType Container)) {
        $issues.Add("InputRoot directory does not exist: $resolvedInputRoot")
    }

    $checkedDirectories = @{}
    foreach ($task in @($Config.Tasks)) {
        if (-not $task.ContainsKey('InputDirectory') -or -not $task.InputDirectory) {
            continue
        }

        $taskName = if ($task.ContainsKey('Name') -and $task.Name) { [string]$task.Name } else { '<unnamed-task>' }
        $rawInputDirectory = [string]$task.InputDirectory
        $resolvedInputDirectory = Resolve-TemplateText -Template $rawInputDirectory -Tokens $Tokens

        if (-not $resolvedInputDirectory) {
            $issues.Add("Task '$taskName' has empty InputDirectory after token resolution (raw: '$rawInputDirectory').")
            continue
        }

        if ($checkedDirectories.ContainsKey($resolvedInputDirectory)) {
            continue
        }

        if (-not (Test-Path -LiteralPath $resolvedInputDirectory -PathType Container)) {
            $issues.Add("Task '$taskName' InputDirectory does not exist: $resolvedInputDirectory (raw: '$rawInputDirectory').")
        }
        $checkedDirectories[$resolvedInputDirectory] = $true
    }

    if ($issues.Count -gt 0) {
        $details = $issues | ForEach-Object { " - $_" }
        throw ("Configuration pre-check failed for '$ConfigPath':`n" + ($details -join [Environment]::NewLine))
    }
}

<#
.SYNOPSIS
English code-review note for function 'Get-TaskResultByName'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-TaskResultByName {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$TaskResults,
        [Parameter(Mandatory = $true)]
        [string]$TaskName
    )

    return @($TaskResults | Where-Object { $_.Name -eq $TaskName } | Select-Object -First 1)[0]
}

<#
.SYNOPSIS
English code-review note for function 'Get-TaskSummaryData'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-TaskSummaryData {
    param(
        [Parameter(Mandatory = $true)]
        [object]$TaskResult
    )

    if (-not $TaskResult) {
        return $null
    }

    if ($TaskResult.Status -ne 'completed') {
        return $null
    }

    if (-not $TaskResult.SummaryPath -or -not (Test-Path $TaskResult.SummaryPath -PathType Leaf)) {
        return $null
    }

    return Get-Content -LiteralPath $TaskResult.SummaryPath -Raw | ConvertFrom-Json
}

<#
.SYNOPSIS
English code-review note for function 'ConvertTo-BackupStaticRule'.
.DESCRIPTION
Provides a reusable workflow helper for audit processing.
#>
function ConvertTo-BackupStaticRule {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,
        [string[]]$DefaultWeeks
    )

    if ($Item -is [string]) {
        return [PSCustomObject]@{
            Template       = [string]$Item
            Source         = 'generated'
            Required       = $true
            AppliesToWeeks = @($DefaultWeeks)
            Description    = $null
        }
    }

    $template = if ($null -ne $Item.Template -and [string]$Item.Template) {
        [string]$Item.Template
    }
    elseif ($null -ne $Item.Name -and [string]$Item.Name) {
        [string]$Item.Name
    }
    else {
        throw 'Static backup validation rule must define Template.'
    }

    $appliesToWeeks = if ($null -ne $Item.AppliesToWeeks -and @($Item.AppliesToWeeks).Count -gt 0) {
        @([string[]]$Item.AppliesToWeeks)
    }
    else {
        @($DefaultWeeks)
    }

    return [PSCustomObject]@{
        Template       = $template
        Source         = if ($null -ne $Item.Source -and [string]$Item.Source) { [string]$Item.Source } else { 'generated' }
        Required       = if ($null -ne $Item.Required) { [bool]$Item.Required } else { $true }
        AppliesToWeeks = $appliesToWeeks
        Description    = if ($null -ne $Item.Description) { [string]$Item.Description } else { $null }
    }
}

<#
.SYNOPSIS
English code-review note for function 'ConvertTo-BackupDynamicRule'.
.DESCRIPTION
Provides a reusable workflow helper for audit processing.
#>
function ConvertTo-BackupDynamicRule {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Item,
        [string[]]$DefaultWeeks
    )

    if ($Item -is [string]) {
        throw 'Dynamic backup validation rule must define BaseName and SummaryTaskName.'
    }

    $baseName = if ($null -ne $Item.BaseName -and [string]$Item.BaseName) {
        [string]$Item.BaseName
    }
    elseif ($null -ne $Item.Template -and [string]$Item.Template) {
        [string]$Item.Template
    }
    else {
        throw 'Dynamic backup validation rule must define BaseName.'
    }

    if ($null -eq $Item.SummaryTaskName -or -not [string]$Item.SummaryTaskName) {
        throw "Dynamic backup validation rule '$baseName' must define SummaryTaskName."
    }

    $appliesToWeeks = if ($null -ne $Item.AppliesToWeeks -and @($Item.AppliesToWeeks).Count -gt 0) {
        @([string[]]$Item.AppliesToWeeks)
    }
    else {
        @($DefaultWeeks)
    }

    return [PSCustomObject]@{
        BaseName        = $baseName
        SummaryTaskName = [string]$Item.SummaryTaskName
        Source          = if ($null -ne $Item.Source -and [string]$Item.Source) { [string]$Item.Source } else { 'generated' }
        Required        = if ($null -ne $Item.Required) { [bool]$Item.Required } else { $true }
        AppliesToWeeks  = $appliesToWeeks
        Description     = if ($null -ne $Item.Description) { [string]$Item.Description } else { $null }
    }
}

<#
.SYNOPSIS
English code-review note for function 'Get-BackupValidationConfig'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-BackupValidationConfig {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    $validationNode = $null
    $rulesNode = $null
    $enforceFailure = $false

    if ($Config.ContainsKey('BackupValidation') -and $Config.BackupValidation) {
        $validationNode = $Config.BackupValidation
        $enforceFailure = if ($validationNode.ContainsKey('EnforceFailure')) {
            [bool]$validationNode.EnforceFailure
        }
        elseif ($validationNode.ContainsKey('EnforceBackupValidation')) {
            [bool]$validationNode.EnforceBackupValidation
        }
        else {
            $false
        }

        if ($validationNode.ContainsKey('Rules') -and $validationNode.Rules) {
            $rulesNode = $validationNode.Rules
        }
        else {
            $rulesNode = $validationNode
        }
    }
    elseif ($Config.ContainsKey('BackupValidationRules') -and $Config.BackupValidationRules) {
        $rulesNode = $Config.BackupValidationRules
        $enforceFailure = if ($Config.ContainsKey('EnforceBackupValidation')) { [bool]$Config.EnforceBackupValidation } else { $false }
    }
    else {
        return $null
    }

    function Get-RuleItems {
        param(
            [Parameter(Mandatory = $true)]
            [object]$Node,
            [Parameter(Mandatory = $true)]
            [string]$PropertyName
        )

        $value = $null
        if ($Node -is [hashtable]) {
            if (-not $Node.ContainsKey($PropertyName)) {
                return @()
            }

            $value = $Node[$PropertyName]
        }
        else {
            $property = $Node.PSObject.Properties[$PropertyName]
            if (-not $property) {
                return @()
            }

            $value = $property.Value
        }

        if ($null -eq $value) {
            return @()
        }

        return @($value)
    }

    $staticRules = New-Object 'System.Collections.Generic.List[object]'
    $dynamicRules = New-Object 'System.Collections.Generic.List[object]'

    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'CommonFiles')) {
        [void]$staticRules.Add((ConvertTo-BackupStaticRule -Item $item -DefaultWeeks @()))
    }
    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'CommonFixedFiles')) {
        [void]$staticRules.Add((ConvertTo-BackupStaticRule -Item $item -DefaultWeeks @()))
    }
    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'TwoWeekFiles')) {
        [void]$staticRules.Add((ConvertTo-BackupStaticRule -Item $item -DefaultWeeks @('2')))
    }
    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'TwoWeekFixedFiles')) {
        [void]$staticRules.Add((ConvertTo-BackupStaticRule -Item $item -DefaultWeeks @('2')))
    }
    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'FourWeekFiles')) {
        [void]$staticRules.Add((ConvertTo-BackupStaticRule -Item $item -DefaultWeeks @('4')))
    }
    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'FourWeekFixedFiles')) {
        [void]$staticRules.Add((ConvertTo-BackupStaticRule -Item $item -DefaultWeeks @('4')))
    }
    foreach ($item in (Get-RuleItems -Node $rulesNode -PropertyName 'DynamicFiles')) {
        [void]$dynamicRules.Add((ConvertTo-BackupDynamicRule -Item $item -DefaultWeeks @()))
    }

    return [PSCustomObject]([ordered]@{
        EnforceFailure = $enforceFailure
        StaticRules    = @($staticRules.ToArray())
        DynamicRules   = @($dynamicRules.ToArray())
    })
}

<#
.SYNOPSIS
English code-review note for function 'Get-ExpectedMessageFiles'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-ExpectedMessageFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaseName,
        [object]$SummaryData
    )

    if (-not $SummaryData -or -not $SummaryData.HasViolation) {
        return @($BaseName)
    }

    $count = [int]$SummaryData.ViolationDivisionCount
    if ($count -le 0) {
        return @($BaseName)
    }

    $extension = [System.IO.Path]::GetExtension($BaseName)
    $baseWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($BaseName)
    $files = New-Object 'System.Collections.Generic.List[string]'
    foreach ($index in 1..$count) {
        $files.Add(('{0}_{1}{2}' -f $baseWithoutExtension, $index, $extension))
    }

    return @($files)
}

<#
.SYNOPSIS
English code-review note for function 'Get-ExpectedBackupFiles'.
.DESCRIPTION
Retrieves computed or existing values required by the audit pipeline.
#>
function Get-ExpectedBackupFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory = $true)]
        [hashtable]$DateTokens,
        [Parameter(Mandatory = $true)]
        [object]$BackupValidationConfig,
        [hashtable]$TaskSummaries = @{}
    )

    $expected = New-Object 'System.Collections.Generic.List[string]'

    foreach ($rule in @($BackupValidationConfig.StaticRules)) {
        if (-not $rule.Required) {
            continue
        }

        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) {
            continue
        }

        $expected.Add((Resolve-TemplateText -Template ([string]$rule.Template) -Tokens $DateTokens))
    }

    foreach ($rule in @($BackupValidationConfig.DynamicRules)) {
        if (-not $rule.Required) {
            continue
        }

        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) {
            continue
        }

        $baseName = Resolve-TemplateText -Template ([string]$rule.BaseName) -Tokens $DateTokens
        $summaryData = if ($TaskSummaries.ContainsKey([string]$rule.SummaryTaskName)) { $TaskSummaries[[string]$rule.SummaryTaskName] } else { $null }
        foreach ($name in (Get-ExpectedMessageFiles -BaseName $baseName -SummaryData $summaryData)) {
            $expected.Add($name)
        }
    }

    return @($expected)
}

<#
.SYNOPSIS
English code-review note for function 'Test-BackupFolderContent'.
.DESCRIPTION
Validates current state and returns comparison results for audit checks.
#>
function Test-BackupFolderContent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,
        [Parameter(Mandatory = $true)]
        [string[]]$ExpectedFiles
    )

    $actualFiles = @(
        Get-ChildItem -LiteralPath $BackupFolder -File |
            Select-Object -ExpandProperty Name
    )
    $missingFiles = @($ExpectedFiles | Where-Object { $actualFiles -notcontains $_ })
    $unexpectedFiles = @($actualFiles | Where-Object { $ExpectedFiles -notcontains $_ })

    return [PSCustomObject]@{
        ExpectedFiles   = @($ExpectedFiles)
        ActualFiles     = @($actualFiles)
        MissingFiles    = @($missingFiles)
        UnexpectedFiles = @($unexpectedFiles)
        Passed          = ($missingFiles.Count -eq 0 -and $unexpectedFiles.Count -eq 0)
    }
}

<#
.SYNOPSIS
English code-review note for function 'Format-BackupValidationText'.
.DESCRIPTION
Formats data into a human-readable representation for review output.
#>
function Format-BackupValidationText {
    param(
        [Parameter(Mandatory = $true)]
        [object]$ValidationResult,
        [Parameter(Mandatory = $true)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder
    )

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('Backup Validation Report')
    $lines.Add("Week Cycle: $CurrentRunWeeks")
    $lines.Add("Backup Folder: $BackupFolder")
    $lines.Add("Passed: $($ValidationResult.Passed)")
    if ($ValidationResult.PSObject.Properties['ValidationMode'] -and $ValidationResult.ValidationMode) {
        $lines.Add("Validation Mode: $($ValidationResult.ValidationMode)")
    }
    if ($ValidationResult.PSObject.Properties['MergedRunIds'] -and @($ValidationResult.MergedRunIds).Count -gt 0) {
        $lines.Add("Merged Runs: $($ValidationResult.MergedRunIds -join ', ')")
    }
    $lines.Add('')

    $lines.Add('Missing Files:')
    if (@($ValidationResult.MissingFiles).Count -eq 0) {
        $lines.Add('  (none)')
    }
    else {
        foreach ($file in $ValidationResult.MissingFiles) {
            $lines.Add("  - $file")
        }
    }

    $lines.Add('')
    $lines.Add('Unexpected Files:')
    if (@($ValidationResult.UnexpectedFiles).Count -eq 0) {
        $lines.Add('  (none)')
    }
    else {
        foreach ($file in $ValidationResult.UnexpectedFiles) {
            $lines.Add("  - $file")
        }
    }

    $lines.Add('')
    $lines.Add('Expected Files:')
    foreach ($file in $ValidationResult.ExpectedFiles) {
        $lines.Add("  - $file")
    }

    $lines.Add('')
    $lines.Add('Actual Files:')
    foreach ($file in $ValidationResult.ActualFiles) {
        $lines.Add("  - $file")
    }

    return ($lines -join [Environment]::NewLine)
}

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
English code-review note for function 'New-LazyLdapConnection'.
.DESCRIPTION
Creates a new object or structure used by subsequent processing steps.
#>
function New-LazyLdapConnection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [int]$Port = 636,
        [switch]$UseSsl = $true,
        [int]$TimeoutSeconds = 30,
        [Parameter(Mandatory = $true)]
        [System.Net.NetworkCredential]$Credential
    )

    $server = $Server
    $port = [int]$Port
    $useSsl = $UseSsl
    $timeout = [int]$TimeoutSeconds
    $credential = $Credential

    $factory = [System.Func[System.DirectoryServices.Protocols.LdapConnection]] {
        $identifier = [System.DirectoryServices.Protocols.LdapDirectoryIdentifier]::new($server, $port)
        $conn = [System.DirectoryServices.Protocols.LdapConnection]::new(
            $identifier,
            $credential,
            [System.DirectoryServices.Protocols.AuthType]::Negotiate
        )
        $conn.SessionOptions.ProtocolVersion = 3
        if ($useSsl) {
            $conn.SessionOptions.SecureSocketLayer = $true
        }

        $conn.Timeout = [TimeSpan]::FromSeconds($timeout)
        try {
            $conn.Bind()
        }
        catch {
            $conn.Dispose()
            throw "LDAP bind failed: $($_.Exception.Message)"
        }

        return $conn
    }

    return [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]::new(
        $factory,
        [System.Threading.LazyThreadSafetyMode]::PublicationOnly
    )
}

<#
.SYNOPSIS
Builds an LDAP OR filter for a list of values.
.DESCRIPTION
Generates an LDAP filter fragment by combining object class and attribute clauses.
#>
function New-LdapOrFilter {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Values,
        [Parameter(Mandatory = $true)]
        [string]$AttributeName,
        [string]$ObjectClassFilter = '(objectClass=user)'
    )

    $clauses = @(
        $Values |
            ForEach-Object { $_ } |
            Where-Object { $_ } |
            ForEach-Object { "($AttributeName=$_)"} 
    )

    if ($clauses.Count -eq 0) {
        throw 'New-LdapOrFilter requires at least one non-empty value.'
    }

    return "(&${ObjectClassFilter}(|$($clauses -join '')))"
}

<#
.SYNOPSIS
Splits values into fixed-size batches.
.DESCRIPTION
Returns an array-of-arrays used by LDAP query functions to control request size.
#>
function Split-LdapBatches {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Values,
        [int]$BatchSize = 20
    )

    if ($BatchSize -lt 1) {
        throw 'BatchSize must be greater than 0.'
    }

    $batches = New-Object System.Collections.Generic.List[object]
    $current = New-Object System.Collections.Generic.List[string]

    foreach ($value in $Values) {
        if (-not $value) {
            continue
        }

        $current.Add([string]$value)
        if ($current.Count -ge $BatchSize) {
            $batches.Add(@($current.ToArray()))
            $current.Clear()
        }
    }

    if ($current.Count -gt 0) {
        $batches.Add(@($current.ToArray()))
    }

    return @($batches.ToArray())
}

<#
.SYNOPSIS
Resolves LDAP search base from RootDSE when not provided.
.DESCRIPTION
Uses defaultNamingContext as the search base for subtree queries.
#>
function Resolve-LdapSearchBase {
    param(
        [Parameter(Mandatory = $true)]
        [System.DirectoryServices.Protocols.LdapConnection]$Connection,
        [string]$SearchBase
    )

    if ($SearchBase) {
        return $SearchBase
    }

    $rootReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
        '',
        '(objectClass=*)',
        [System.DirectoryServices.Protocols.SearchScope]::Base,
        'defaultNamingContext'
    )
    $rootResp = $Connection.SendRequest($rootReq)
    if ($rootResp.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success) {
        throw "Failed to read RootDSE: $($rootResp.ErrorMessage)"
    }

    return [string](($rootResp.Entries[0].Attributes['defaultNamingContext'])[0])
}

<#
.SYNOPSIS
Performs LDAP lookups by mail address.
.DESCRIPTION
Returns a hashtable with Valid mail mappings and Invalid addresses.
#>
function Get-LdapUserByMail {
    param(
        [Parameter(Mandatory = $true)]
        [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]$LazyConnection,
        [Parameter(Mandatory = $true)]
        [string[]]$MailAdds,
        [string]$SearchBase,
        [int]$SearchTimeoutSeconds = 30,
        [string[]]$Attributes = @('sAMAccountName', 'mail', 'division'),
        [int]$BatchSize = 20
    )

    try {
        $conn = $LazyConnection.Value
    }
    catch {
        throw "Failed to create LDAP connection: $($_.Exception.Message)"
    }

    $conn.Timeout = [TimeSpan]::FromSeconds($SearchTimeoutSeconds)
    $resolvedBase = Resolve-LdapSearchBase -Connection $conn -SearchBase $SearchBase

    $mailLookup = @{}
    $invalid = New-Object System.Collections.Generic.List[string]
    $normalized = @(
        $MailAdds |
            ForEach-Object {
                if ($_ -is [string]) {
                    $_.Trim().ToLower()
                }
            } |
            Where-Object { $_ }
    )

    foreach ($batch in Split-LdapBatches -Values $normalized -BatchSize $BatchSize) {
        $filter = New-LdapOrFilter -Values $batch -AttributeName 'mail' -ObjectClassFilter '(|(objectClass=user)(objectClass=group))'
        $searchReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
            $resolvedBase,
            $filter,
            [System.DirectoryServices.Protocols.SearchScope]::Subtree,
            $Attributes
        )

        $searchResp = $conn.SendRequest($searchReq)
        $foundMails = @()

        foreach ($entry in $searchResp.Entries) {
            $info = [PSCustomObject]@{
                SamAccountName = $null
                Mail           = $null
                Division       = $null
            }

            if ($entry.Attributes['mail']) {
                $info.Mail = [string]$entry.Attributes['mail'][0].ToLower()
                $foundMails += $info.Mail
            }
            if ($entry.Attributes['sAMAccountName']) {
                $info.SamAccountName = [string]$entry.Attributes['sAMAccountName'][0]
            }
            if ($entry.Attributes['division']) {
                $info.Division = [string]$entry.Attributes['division'][0]
            }

            if ($info.Mail) {
                $mailLookup[$info.Mail] = $info
            }
        }

        foreach ($address in $batch) {
            if (-not ($foundMails -contains $address)) {
                $invalid.Add($address)
            }
        }
    }

    return @{
        Valid   = $mailLookup
        Invalid = $invalid
    }
}

<#
.SYNOPSIS
Performs LDAP lookups by account ID.
.DESCRIPTION
Returns a hashtable with Valid account mappings and Invalid account IDs.
#>
function Get-LdapUserById {
    param(
        [Parameter(Mandatory = $true)]
        [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]$LazyConnection,
        [Parameter(Mandatory = $true)]
        [string[]]$UserId,
        [string]$SearchBase,
        [int]$SearchTimeoutSeconds = 30,
        [string[]]$Attributes = @('sAMAccountName', 'division', 'legalentity'),
        [int]$BatchSize = 20
    )

    try {
        $conn = $LazyConnection.Value
    }
    catch {
        throw "failed to create LDAP connection: $($_.Exception.Message)"
    }

    $conn.Timeout = [TimeSpan]::FromSeconds($SearchTimeoutSeconds)
    $resolvedBase = Resolve-LdapSearchBase -Connection $conn -SearchBase $SearchBase

    $lookup = @{}
    $invalid = New-Object System.Collections.Generic.List[string]
    $normalized = @(
        $UserId |
            ForEach-Object {
                if ($_ -is [string]) {
                    $_ -split ';' | ForEach-Object { $_.Trim().ToLower() }
                }
            } |
            Where-Object { $_ } |
            Select-Object -Unique
    )

    foreach ($batch in Split-LdapBatches -Values $normalized -BatchSize $BatchSize) {
        $filter = New-LdapOrFilter -Values $batch -AttributeName 'cn' -ObjectClassFilter '(objectClass=user)'
        $searchReq = [System.DirectoryServices.Protocols.SearchRequest]::new(
            $resolvedBase,
            $filter,
            [System.DirectoryServices.Protocols.SearchScope]::Subtree,
            $Attributes
        )
        $searchResp = $conn.SendRequest($searchReq)
        $foundCnNames = @()

        foreach ($entry in $searchResp.Entries) {
            $info = [PSCustomObject]@{
                SamAccountName = $null
                Division       = $null
                LegalEntity    = $null
            }

            if ($entry.Attributes['sAMAccountName']) {
                $info.SamAccountName = [string]$entry.Attributes['sAMAccountName'][0]
                $foundCnNames += $info.SamAccountName.ToLower()
            }
            if ($entry.Attributes['division']) {
                $info.Division = [string]$entry.Attributes['division'][0]
            }
            if ($entry.Attributes['legalentity']) {
                $info.LegalEntity = [string]$entry.Attributes['legalentity'][0]
            }

            if ($info.SamAccountName) {
                $lookup[$info.SamAccountName.ToLower()] = $info
            }
        }

        foreach ($id in $batch) {
            if (-not ($foundCnNames -contains $id)) {
                $invalid.Add($id)
            }
        }
    }

    return @{
        Valid   = $lookup
        Invalid = $invalid
    }
}

<#
.SYNOPSIS
English code-review note for function 'Close-LazyLdapConnection'.
.DESCRIPTION
Releases resources created earlier in the workflow to avoid leaks.
#>
function Close-LazyLdapConnection {
    param(
        [Parameter(Mandatory = $true)]
        [System.Lazy[System.DirectoryServices.Protocols.LdapConnection]]$Lazy
    )

    if ($Lazy.IsValueCreated) {
        try {
            $Lazy.Value.Dispose()
        }
        catch {
            throw "LDAP dispose issue: $($_.Exception.Message)"
        }
    }
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
        [Parameter(Mandatory = $true)]
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
        $mail.CC.Add($Cc)
        $mail.Subject = $Subject
        $mail.Body = $Body
        $mail.IsBodyHtml = $true

        $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)
        $smtp.EnableSsl = $true
        $smtp.ClientCertificates.Add([System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Cert))
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

Export-ModuleMember -Function Convert-ExactDate, Write-Log, Get-Cert, Get-LogFilePath, Get-VaultSecret, New-LazyLdapConnection, Close-LazyLdapConnection, Send-Mail, New-HtmlBody, New-DateTokenMap, Resolve-TemplateText, Assert-ConfigInputDirectories, Get-TaskResultByName, Get-TaskSummaryData, Get-BackupValidationConfig, Get-ExpectedBackupFiles, Test-BackupFolderContent, Format-BackupValidationText, New-LdapOrFilter, Split-LdapBatches, Resolve-LdapSearchBase, Get-LdapUserByMail, Get-LdapUserById
