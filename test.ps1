if (-not ('System.DirectoryServices.Protocols.LdapConnection' -as [type])) {
    Add-Type -AssemblyName System.DirectoryServices.Protocols
}

$script:WeComAuditLogFolderName = 'wecom_audit_log'

<#
.SYNOPSIS
Returns the canonical subfolder name used by all entry scripts under LogRoot.
#>
function Get-WeComAuditLogFolderName {
    return $script:WeComAuditLogFolderName
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
Retrieves an optional property value from dictionary-like or object inputs.
.DESCRIPTION
Supports case-insensitive lookup for hashtable-like inputs and returns $null
when the property does not exist.
#>
function Get-OptionalObjectPropertyValue {
    param(
        [Parameter()]
        [object]$InputObject,
        [Parameter(Mandatory = $true)]
        [string]$PropertyName
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        if ($InputObject.Contains($PropertyName)) {
            return $InputObject[$PropertyName]
        }

        foreach ($key in $InputObject.Keys) {
            if ([string]$key -ieq $PropertyName) {
                return $InputObject[$key]
            }
        }
    }

    $property = $InputObject.PSObject.Properties[$PropertyName]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

<#
.SYNOPSIS
Writes analysis summary data to JSON file.
.DESCRIPTION
Keeps one shared implementation for summary JSON persistence used by
mail/device analyzers.
#>
function Write-AnalysisSummaryJson {
    param(
        [string]$SummaryOutputPath,
        [Parameter(Mandatory = $true)]
        [hashtable]$SummaryFields,
        [int]$Depth = 4
    )

    if (-not $SummaryOutputPath) {
        return
    }

    ([PSCustomObject]$SummaryFields) |
        ConvertTo-Json -Depth $Depth |
        Set-Content -Path $SummaryOutputPath -Encoding UTF8
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
            ReadyBy        = 'Validate'
        }
    }

    $template = if ($null -ne $Item.Template -and [string]$Item.Template) {
        [string]$Item.Template
    }
    elseif ($null -ne $Item.File -and [string]$Item.File) {
        [string]$Item.File
    }
    elseif ($null -ne $Item.Name -and [string]$Item.Name) {
        [string]$Item.Name
    }
    else {
        throw 'Static backup validation rule must define Template or File.'
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
        ReadyBy        = if ($null -ne $Item.ReadyBy -and [string]$Item.ReadyBy) { [string]$Item.ReadyBy } else { 'Validate' }
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
function Export-AnalysisReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LogFilePath,
        [string]$TaskOutputDirectory,
        [string]$SubFolder = 'AnalysisReport',
        [switch]$UseDateSubFolder
    )

    $baseFolder = if ($TaskOutputDirectory) {
        if (-not (Test-Path -LiteralPath $TaskOutputDirectory)) {
            New-Item -Path $TaskOutputDirectory -ItemType Directory -Force | Out-Null
        }
        $TaskOutputDirectory
    }
    else {
        Split-Path -Parent $LogFilePath
    }

    $reportFolder = Join-Path $baseFolder $SubFolder

    if ($UseDateSubFolder) {
        $timestamp = Get-Date -Format 'yyyy_MM_dd'
        $reportFolder = Join-Path $reportFolder $timestamp
    }

    if (-not (Test-Path -LiteralPath $reportFolder)) {
        New-Item -Path $reportFolder -ItemType Directory -Force | Out-Null
    }

    return $reportFolder
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

<#
.SYNOPSIS
Normalizes a path for safe string-based prefix comparison.
.DESCRIPTION
Uses System.IO.Path.GetFullPath to resolve '..' and duplicate separators without
triggering filesystem access (safe on UNC). Trims trailing separators so the
result can be compared by appending a single DirectorySeparatorChar.
#>
function Get-NormalizedFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $normalized = [System.IO.Path]::GetFullPath($Path)
    return $normalized.TrimEnd(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
}

<#
.SYNOPSIS
Returns $true only if $Path sits strictly underneath one of $AllowedRoots.
.DESCRIPTION
Prefix match uses OrdinalIgnoreCase and appends a separator to prevent prefix
collision (e.g. 'C:\data' must not accept 'C:\dataX\...').
Returns $false for empty / null AllowedRoots (fail-closed).
#>
function Test-PathWithinAllowedRoots {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string[]]$AllowedRoots
    )

    if (-not $AllowedRoots -or @($AllowedRoots).Count -eq 0) {
        return $false
    }

    $normalizedTarget = Get-NormalizedFullPath $Path
    foreach ($root in $AllowedRoots) {
        if (-not $root) { continue }

        $normalizedRoot = Get-NormalizedFullPath $root
        $rootWithSeparator = $normalizedRoot + [System.IO.Path]::DirectorySeparatorChar
        if ($normalizedTarget.StartsWith(
                $rootWithSeparator,
                [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

<#
.SYNOPSIS
Four-layer safety check before deleting a source file that has been backed up.
.DESCRIPTION
Returns a result object: Safe (bool) and Reason (string).
Checks performed in order:
  1. Source and backup exist as Leaf files.
  2. Source resides within one of the configured AllowedRoots.
  3. Source is not a reparse point (symlink/junction).
  4. Source and backup SHA256 match (detects in-flight modification / corrupt backup).
#>
function Test-SafeToDeleteSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,
        [Parameter(Mandatory = $true)]
        [string]$BackupPath,
        [string[]]$AllowedRoots
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source file not found (already deleted or moved)' }
    }
    if (-not (Test-Path -LiteralPath $BackupPath -PathType Leaf)) {
        return [PSCustomObject]@{ Safe = $false; Reason = "backup file not found: $BackupPath" }
    }
    if (-not (Test-PathWithinAllowedRoots -Path $SourcePath -AllowedRoots $AllowedRoots)) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source path is not within SourceDeletionAllowedRoots' }
    }

    $sourceItem = Get-Item -LiteralPath $SourcePath -Force
    if ($sourceItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source is a reparse point (symlink/junction), refusing to delete' }
    }

    $sourceHash = (Get-FileHash -LiteralPath $SourcePath -Algorithm SHA256).Hash
    $backupHash = (Get-FileHash -LiteralPath $BackupPath -Algorithm SHA256).Hash
    if ($sourceHash -ne $backupHash) {
        return [PSCustomObject]@{ Safe = $false; Reason = 'source file hash does not match backup (source modified or backup corrupt)' }
    }

    return [PSCustomObject]@{ Safe = $true; Reason = 'passed all safety checks' }
}

<#
.SYNOPSIS
Deletes a file with bounded retry suitable for transient NAS errors.
.DESCRIPTION
Returns a result object with Success, Error, and Attempts. Retries on IOException
and general failures (network blips) with a fixed delay. Does not retry on
permission errors that are clearly non-transient — but PowerShell's Remove-Item
does not always distinguish these cleanly, so we treat all failures as retryable
and rely on caller's log to surface patterns.
#>
function Remove-SourceFileWithRetry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            Remove-Item -LiteralPath $Path -Force
            return [PSCustomObject]@{ Success = $true; Error = $null; Attempts = $attempt }
        }
        catch {
            $errorMessage = $_.Exception.Message
            if ($attempt -eq $MaxAttempts) {
                return [PSCustomObject]@{ Success = $false; Error = $errorMessage; Attempts = $attempt }
            }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

<#
.SYNOPSIS
Orchestrates run-level source file cleanup with safety checks and retries.
.DESCRIPTION
Accepts a list of pending deletion items (each with SourcePath, BackupPath,
TaskName), evaluates each through Test-SafeToDeleteSourceFile, deletes the
safe ones with retry, and returns a structured summary suitable for embedding
into run-summary.json.
Additionally asserts the backup folder exists and is non-empty before
attempting any deletion — if that sanity check fails, the entire cleanup
is aborted and marked Skipped.
#>
function Invoke-SourceFileCleanup {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$PendingDeletions,
        [Parameter(Mandatory = $true)]
        [string]$BackupFolder,
        [string[]]$AllowedRoots,
        [string]$LogFilePath
    )

    $details = New-Object 'System.Collections.Generic.List[object]'
    $deletedCount = 0
    $skippedCount = 0
    $failedCount = 0
    $aborted = $false
    $abortReason = $null

    if (-not (Test-Path -LiteralPath $BackupFolder -PathType Container)) {
        $aborted = $true
        $abortReason = "backup folder not found: $BackupFolder"
    }
    $uniquePending = New-Object 'System.Collections.Generic.List[object]'
    $duplicatePending = New-Object 'System.Collections.Generic.List[object]'
    $seenSourcePaths = @{}
    foreach ($item in $PendingDeletions) {
        $key = (Get-NormalizedFullPath ([string]$item.SourcePath)).ToLowerInvariant()
        if ($seenSourcePaths.ContainsKey($key)) {
            $duplicatePending.Add([PSCustomObject]@{
                Item       = $item
                FirstOwner = $seenSourcePaths[$key]
            })
        }
        else {
            $seenSourcePaths[$key] = [string]$item.TaskName
            $uniquePending.Add($item)
        }
    }

    if ($uniquePending.Count -gt 0 -and (Test-Path -LiteralPath $BackupFolder -PathType Container)) {
        $actualBackupCount = @(Get-ChildItem -LiteralPath $BackupFolder -File -ErrorAction SilentlyContinue).Count
        if ($actualBackupCount -lt $uniquePending.Count) {
            $aborted = $true
            $abortReason = "backup folder has $actualBackupCount files, expected at least $($uniquePending.Count) unique source(s) — aborting cleanup to avoid data loss"
        }
    }

    if ($aborted) {
        if ($LogFilePath) {
            Write-Log -LogString "SourceCleanup aborted: $abortReason" -LogFilePath $LogFilePath
        }
        return [PSCustomObject]@{
            Attempted    = $false
            Aborted      = $true
            AbortReason  = $abortReason
            TotalCount   = @($PendingDeletions).Count
            UniqueCount  = $uniquePending.Count
            DeletedCount = 0
            SkippedCount = 0
            FailedCount  = 0
            Details      = @()
        }
    }

    foreach ($dup in $duplicatePending) {
        $details.Add([PSCustomObject]@{
            TaskName   = [string]$dup.Item.TaskName
            SourcePath = [string]$dup.Item.SourcePath
            Status     = 'deduplicated'
            Reason     = "shared source already owned by task '$($dup.FirstOwner)'"
            Attempts   = 0
        })
    }

    foreach ($item in $uniquePending) {
        $sourcePath = [string]$item.SourcePath
        $backupPath = [string]$item.BackupPath
        $taskName = [string]$item.TaskName

        $safetyCheck = Test-SafeToDeleteSourceFile -SourcePath $sourcePath -BackupPath $backupPath -AllowedRoots $AllowedRoots
        if (-not $safetyCheck.Safe) {
            $details.Add([PSCustomObject]@{
                TaskName   = $taskName
                SourcePath = $sourcePath
                Status     = 'skipped'
                Reason     = $safetyCheck.Reason
                Attempts   = 0
            })
            $skippedCount++
            if ($LogFilePath) {
                Write-Log -LogString "SourceCleanup skipped '$sourcePath' (task '$taskName'): $($safetyCheck.Reason)" -LogFilePath $LogFilePath
            }
            continue
        }

        $deleteResult = Remove-SourceFileWithRetry -Path $sourcePath
        if ($deleteResult.Success) {
            $details.Add([PSCustomObject]@{
                TaskName   = $taskName
                SourcePath = $sourcePath
                Status     = 'deleted'
                Reason     = $null
                Attempts   = $deleteResult.Attempts
            })
            $deletedCount++
            if ($LogFilePath) {
                Write-Log -LogString "SourceCleanup deleted '$sourcePath' (task '$taskName') after $($deleteResult.Attempts) attempt(s)" -LogFilePath $LogFilePath
            }
        }
        else {
            $details.Add([PSCustomObject]@{
                TaskName   = $taskName
                SourcePath = $sourcePath
                Status     = 'failed'
                Reason     = $deleteResult.Error
                Attempts   = $deleteResult.Attempts
            })
            $failedCount++
            if ($LogFilePath) {
                Write-Log -LogString "SourceCleanup failed '$sourcePath' (task '$taskName') after $($deleteResult.Attempts) attempt(s): $($deleteResult.Error)" -LogFilePath $LogFilePath
            }
        }
    }

    return [PSCustomObject]@{
        Attempted    = $true
        Aborted      = $false
        AbortReason  = $null
        TotalCount   = @($PendingDeletions).Count
        UniqueCount  = $uniquePending.Count
        DeletedCount = $deletedCount
        SkippedCount = $skippedCount
        FailedCount  = $failedCount
        Details      = @($details.ToArray())
    }
}

<#
.SYNOPSIS
Reports whether an allowed-root path is dangerously broad.
.DESCRIPTION
Returns TooBroad=$true when the path equals a filesystem drive root (e.g. 'C:\')
or a UNC share root (e.g. '\\host\share') — anything strictly underneath would
implicitly include unrelated business data. Caller decides whether to warn or fail.
#>
function Test-AllowedRootIsTooBroad {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Root
    )

    $separators = @(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $normalized = [System.IO.Path]::GetFullPath($Root).TrimEnd($separators)
    $pathRoot = [System.IO.Path]::GetPathRoot($normalized).TrimEnd($separators)

    if ($normalized -ieq $pathRoot) {
        return [PSCustomObject]@{
            TooBroad = $true
            Reason   = "allowed root '$Root' resolves to a filesystem drive or UNC share root"
        }
    }
    return [PSCustomObject]@{ TooBroad = $false; Reason = $null }
}

<#
.SYNOPSIS
Resolves SourceCleanup configuration from nested or legacy config shape.
.DESCRIPTION
Preferred shape:
    SourceCleanup = @{ Enabled = $true; AllowedRoots = @('\\host\share\folder') }
Legacy shape (still supported):
    SourceDeletionAllowedRoots = @('\\host\share\folder')   # implies Enabled=$true
Returns @{ Enabled; AllowedRoots; ConfigShape } where ConfigShape is one of
'nested', 'legacy', or 'absent'. Absent shape yields Enabled=$true with empty
AllowedRoots so the asserter fails closed.
#>
function Resolve-SourceCleanupConfig {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config
    )

    $enabled = $true
    $allowedRoots = @()
    $shape = 'absent'

    if ($Config.ContainsKey('SourceCleanup') -and $Config.SourceCleanup) {
        $shape = 'nested'
        $sub = $Config.SourceCleanup
        if ($sub -isnot [hashtable]) {
            throw "Config 'SourceCleanup' must be a hashtable, got $($sub.GetType().Name)."
        }
        if ($sub.ContainsKey('Enabled')) {
            $enabled = [bool]$sub.Enabled
        }
        if ($sub.ContainsKey('AllowedRoots') -and $sub.AllowedRoots) {
            $allowedRoots = @(
                [string[]]$sub.AllowedRoots |
                    Where-Object { $_ } |
                    ForEach-Object { [string]$_ }
            )
        }
    }
    elseif ($Config.ContainsKey('SourceDeletionAllowedRoots') -and $Config.SourceDeletionAllowedRoots) {
        $shape = 'legacy'
        $allowedRoots = @(
            [string[]]$Config.SourceDeletionAllowedRoots |
                Where-Object { $_ } |
                ForEach-Object { [string]$_ }
        )
    }

    return [PSCustomObject]@{
        Enabled      = $enabled
        AllowedRoots = $allowedRoots
        ConfigShape  = $shape
    }
}

<#
.SYNOPSIS
Startup hard-fail validator for SourceCleanup configuration.
.DESCRIPTION
When Enabled, asserts:
  1. AllowedRoots is non-empty (otherwise cleanup would silently 100% skip).
  2. Every enabled-task input directory is covered by at least one allowed root.
  3. No protected root (BackupRoot/LogRoot/OutputRoot) sits underneath any allowed
     root — would let cleanup walk into backup/log territory.
Emits Write-Warning (does not throw) when an allowed root is too broad
(filesystem/share root). Throws on any hard violation; returns silently otherwise.
When Enabled=$false, all checks are skipped — cleanup will simply not run.
#>
function Assert-SourceCleanupConfig {
    param(
        [Parameter(Mandatory = $true)]
        [bool]$Enabled,
        [string[]]$AllowedRoots,
        [string[]]$EnabledInputDirectories,
        [string[]]$ProtectedRoots
    )

    if (-not $Enabled) { return }

    if (-not $AllowedRoots -or @($AllowedRoots).Count -eq 0) {
        throw "SourceCleanup is enabled but AllowedRoots is empty. This would cause every deletion to be skipped silently. Either populate AllowedRoots or set Enabled = `$false."
    }

    if ($EnabledInputDirectories) {
        $uncovered = New-Object 'System.Collections.Generic.List[string]'
        foreach ($dir in $EnabledInputDirectories) {
            if (-not $dir) { continue }
            if (-not (Test-PathWithinAllowedRoots -Path $dir -AllowedRoots $AllowedRoots)) {
                $uncovered.Add($dir)
            }
        }
        if ($uncovered.Count -gt 0) {
            $sample = ($uncovered | Select-Object -First 5) -join '; '
            throw "SourceCleanup AllowedRoots does not cover $($uncovered.Count) enabled task input director(ies): $sample. Either extend AllowedRoots or disable cleanup."
        }
    }

    if ($ProtectedRoots) {
        foreach ($protected in $ProtectedRoots) {
            if (-not $protected) { continue }
            if (Test-PathWithinAllowedRoots -Path $protected -AllowedRoots $AllowedRoots) {
                throw "SourceCleanup AllowedRoots would expose protected path '$protected' to deletion. Move backup/log roots outside the source cleanup whitelist."
            }
        }
    }

    foreach ($root in $AllowedRoots) {
        $broad = Test-AllowedRootIsTooBroad -Root $root
        if ($broad.TooBroad) {
            Write-Warning "SourceCleanup: $($broad.Reason). Consider narrowing the whitelist to a business-specific subdirectory."
        }
    }
}

<#
.SYNOPSIS
Asserts every configured task has a unique Name.
.DESCRIPTION
Two tasks sharing a Name collide on tasks/<safe-token>/ output folder, summary
and report files, and on dynamic-file lookup keys in BackupValidation. This is a
hard error caught at startup rather than tolerated as an overwrite.
#>
function Assert-TaskNameUniqueness {
    param(
        [Parameter(Mandatory = $true)]
        [object[]]$Tasks
    )

    $seen = @{}
    $duplicates = New-Object 'System.Collections.Generic.List[string]'
    foreach ($task in $Tasks) {
        $name = [string]$task.Name
        if (-not $name) { continue }
        if ($seen.ContainsKey($name)) {
            $duplicates.Add($name)
        }
        else {
            $seen[$name] = $true
        }
    }
    if ($duplicates.Count -gt 0) {
        $list = ($duplicates | Select-Object -Unique) -join ', '
        throw "Configuration error: duplicate task Name(s) detected: $list. Task names must be unique to avoid output folder collision."
    }
}

function Resolve-ScheduleCycle {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [string]$StartDateOverride,
        [string]$ForceCurrentRunWeeks
    )

    $warnings = New-Object 'System.Collections.Generic.List[object]'

    if (-not $Config.ContainsKey('ScheduleAnchor') -or -not $Config.ScheduleAnchor) {
        throw "Config must define 'ScheduleAnchor' (a Thursday in yyyyMMdd format) to enable scheduled execution."
    }

    $anchorStr = [string]$Config.ScheduleAnchor
    $anchor = [DateTime]::ParseExact($anchorStr, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
    if ($anchor.DayOfWeek -ne [DayOfWeek]::Thursday) {
        throw "ScheduleAnchor '$anchorStr' is not a Thursday. Anchor must be a Thursday to align with the biweekly schedule."
    }

    $today = (Get-Date).Date
    $daysFromAnchor = ($today - $anchor).Days

    if ($daysFromAnchor -lt 0) {
        throw "Today ($($today.ToString('yyyyMMdd'))) is before ScheduleAnchor ($($anchor.ToString('yyyyMMdd'))). Cannot compute cycle."
    }

    if ($today.DayOfWeek -ne [DayOfWeek]::Thursday) {
        $warnings.Add(@{ Code = 'NotThursday'; Message = "Today ($($today.ToString('yyyyMMdd')), $($today.DayOfWeek)) is not a Thursday. Scheduled runs should occur on Thursdays." })
    }

    if ($daysFromAnchor % 14 -ne 0) {
        $warnings.Add(@{ Code = 'OffCycle'; Message = "Today is not on a scheduled biweekly cycle (offset $($daysFromAnchor % 14) days from anchor). Use -StartDate to override." })
    }

    $cycleIndex = [Math]::Floor($daysFromAnchor / 14)
    $anchorEndDate = $anchor.AddDays($cycleIndex * 14).ToString('yyyyMMdd')
    $anchorStartDate = $anchor.AddDays($cycleIndex * 14 - 14).ToString('yyyyMMdd')

    $currentRunWeeks = if ($ForceCurrentRunWeeks) {
        $ForceCurrentRunWeeks
    }
    else {
        if ($cycleIndex % 2 -eq 0) { '2' } else { '4' }
    }

    $isOverride = $false
    if ($StartDateOverride) {
        $startDateParsed = [DateTime]::ParseExact($StartDateOverride, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture)
        $computedStartDate = $StartDateOverride
        $computedEndDate = $startDateParsed.AddDays(14).ToString('yyyyMMdd')
        $isOverride = $true

        if ($computedStartDate -ne $anchorStartDate -or $computedEndDate -ne $anchorEndDate) {
            $warnings.Add(@{
                Code    = 'DateOverrideMismatch'
                Message = "-StartDate override produces date range $computedStartDate-$computedEndDate, which differs from the anchor-derived cycle $anchorStartDate-$anchorEndDate. CurrentRunWeeks=$currentRunWeeks is still derived from today's cycleIndex=$cycleIndex, not the overridden dates. Use -ForceCurrentRunWeeks if the cycle type should also differ."
            })
        }
    }
    else {
        $computedStartDate = $anchorStartDate
        $computedEndDate = $anchorEndDate
    }

    return [PSCustomObject]@{
        Anchor          = $anchor
        CycleIndex      = $cycleIndex
        StartDate       = $computedStartDate
        EndDate         = $computedEndDate
        CurrentRunWeeks = $currentRunWeeks
        IsOverride      = $isOverride
        Warnings        = @($warnings.ToArray())
    }
}

function Resolve-PhaseHandoff {
    param(
        [Parameter(Mandatory)]
        [string]$RunsRoot,
        [Parameter(Mandatory)]
        [string]$ExpectedStartDate,
        [Parameter(Mandatory)]
        [string]$ExpectedEndDate,
        [string]$ExpectedRunStatus = 'Success'
    )

    $pointerPath = [System.IO.Path]::Combine($RunsRoot, 'latest-run.json')

    if (-not (Test-Path -LiteralPath $pointerPath -PathType Leaf)) {
        throw "HANDOFF_NOT_FOUND: latest-run.json not found at '$pointerPath'. Phase Analysis must complete successfully before Phase Validate can run."
    }

    $pointer = Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json

    $runId = $null
    if ($pointer.PSObject.Properties['RunId']) {
        $runId = [string]$pointer.RunId
    }
    if (-not $runId) {
        throw "HANDOFF_NO_RUNID: latest-run.json exists but does not contain a RunId. Phase Analysis may not have completed successfully."
    }

    $pointerStartDate = if ($pointer.PSObject.Properties['StartDate']) { [string]$pointer.StartDate } else { $null }
    $pointerEndDate = if ($pointer.PSObject.Properties['EndDate']) { [string]$pointer.EndDate } else { $null }

    if ($pointerStartDate -ne $ExpectedStartDate -or $pointerEndDate -ne $ExpectedEndDate) {
        throw "HANDOFF_DATE_MISMATCH: expected $ExpectedStartDate-$ExpectedEndDate, found $pointerStartDate-$pointerEndDate (RunId=$runId). This may indicate a different analysis run overwrote latest-run.json between Phase 1 and Phase 2."
    }

    $runStatus = if ($pointer.PSObject.Properties['RunStatus']) { [string]$pointer.RunStatus } else { $null }
    if ($runStatus -ne $ExpectedRunStatus) {
        throw "HANDOFF_STATUS_MISMATCH: RunStatus='$runStatus' (RunId=$runId), expected '$ExpectedRunStatus'. Re-run Phase Analysis or fix the underlying failure before proceeding."
    }

    return [PSCustomObject]@{
        RunId     = $runId
        RunStatus = $runStatus
    }
}

function Resolve-AuditConfigPath {
    param(
        [string]$ConfigPath,
        [string]$ScriptRoot
    )

    if ($ConfigPath) { return $ConfigPath }

    if ($env:WECOM_AUDIT_CONFIG_PATH) {
        return [string]$env:WECOM_AUDIT_CONFIG_PATH
    }

    if ($ScriptRoot) {
        return Join-Path $ScriptRoot 'analysis_task.config.psd1'
    }

    throw "No config file could be resolved. Provide -ConfigPath or set WECOM_AUDIT_CONFIG_PATH."
}

function Resolve-AuditOutputRoot {
    param(
        [string]$OutputRoot,
        [hashtable]$Config,
        [string]$ConfigPath
    )

    $folderName = Get-WeComAuditLogFolderName

    if ($OutputRoot) { return $OutputRoot }

    if ($env:WECOM_AUDIT_LOG_ROOT) {
        return [System.IO.Path]::Combine($env:WECOM_AUDIT_LOG_ROOT, $folderName)
    }

    if ($Config -and $Config.ContainsKey('LogRoot') -and $Config.LogRoot) {
        return [System.IO.Path]::Combine([string]$Config.LogRoot, $folderName)
    }

    if ($ConfigPath) {
        return Split-Path $ConfigPath -Parent
    }

    throw "Cannot resolve output root. Provide -OutputRoot, set WECOM_AUDIT_LOG_ROOT, or ensure config contains LogRoot."
}

function Resolve-AuditInputRoot {
    param(
        [hashtable]$Config
    )

    if ($env:WECOM_AUDIT_INPUT_ROOT) {
        return [string]$env:WECOM_AUDIT_INPUT_ROOT
    }

    if ($Config -and $Config.ContainsKey('InputRoot') -and $Config.InputRoot) {
        return [string]$Config.InputRoot
    }

    return 'C:\addin_deploy_cert'
}

function Resolve-TaskInputPath {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Task,
        [Parameter(Mandatory = $true)]
        [hashtable]$Tokens
    )

    if ($Task.ContainsKey('InputPath') -and $Task.InputPath) {
        return Resolve-TemplateText -Template ([string]$Task.InputPath) -Tokens $Tokens
    }

    if (-not $Task.ContainsKey('InputDirectory') -or -not $Task.InputDirectory) {
        throw "Task '$($Task.Name)' must define InputPath or InputDirectory."
    }
    if (-not $Task.ContainsKey('FileNamePattern') -or -not $Task.FileNamePattern) {
        throw "Task '$($Task.Name)' must define FileNamePattern when InputDirectory is used."
    }

    $inputDirectory = Resolve-TemplateText -Template ([string]$Task.InputDirectory) -Tokens $Tokens
    $fileNamePattern = Resolve-TemplateText -Template ([string]$Task.FileNamePattern) -Tokens $Tokens

    if (-not (Test-Path $inputDirectory -PathType Container)) {
        throw "Task '$($Task.Name)' input directory not found: $inputDirectory"
    }

    $matchedFiles = @(
        Get-ChildItem -LiteralPath $inputDirectory -File |
            Where-Object { $_.Name -like $fileNamePattern }
    )

    if ($matchedFiles.Count -eq 0) {
        throw "Task '$($Task.Name)' did not match any file in '$inputDirectory' with pattern '$fileNamePattern'."
    }
    if ($matchedFiles.Count -gt 1) {
        $matchedNames = $matchedFiles | Select-Object -ExpandProperty Name
        throw "Task '$($Task.Name)' matched multiple files for pattern '$fileNamePattern': $($matchedNames -join ', ')"
    }

    return $matchedFiles[0].FullName
}

function Resolve-NotificationConfig {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [string]$Environment
    )

    if (-not $Config.ContainsKey('Notification') -or -not $Config.Notification) {
        return $null
    }

    $notifConfig = $Config.Notification
    if (-not $notifConfig.ContainsKey($Environment) -or -not $notifConfig[$Environment]) {
        return $null
    }

    $envConfig = $notifConfig[$Environment]

    $cert = $null
    $certName = if ($envConfig.ContainsKey('CertName') -and $envConfig.CertName) { [string]$envConfig.CertName } else { $null }
    if ($certName) {
        try { $cert = Get-Cert -KeyName $certName } catch { $cert = $null }
    }

    return [PSCustomObject]@{
        SmtpServer   = if ($envConfig.ContainsKey('SmtpServer')) { [string]$envConfig.SmtpServer } else { $null }
        Port         = if ($envConfig.ContainsKey('Port')) { [int]$envConfig.Port } else { 2587 }
        From         = if ($envConfig.ContainsKey('From')) { [string]$envConfig.From } else { $null }
        CertName     = $certName
        Cert         = $cert
        OpsTeam      = if ($envConfig.ContainsKey('OpsTeam')) { @($envConfig.OpsTeam) } else { @() }
        CcRecipients = if ($envConfig.ContainsKey('CcRecipients')) { @($envConfig.CcRecipients) } else { @() }
    }
}

function Get-PreflightFiles {
    param(
        [Parameter(Mandatory)]
        [object]$BackupValidationConfig,
        [Parameter(Mandatory)]
        [string]$Phase,
        [Parameter(Mandatory)]
        [string]$CurrentRunWeeks,
        [Parameter(Mandatory)]
        [hashtable]$DateTokens
    )

    $files = New-Object 'System.Collections.Generic.List[object]'

    foreach ($rule in @($BackupValidationConfig.StaticRules)) {
        if (-not $rule.Required) { continue }
        if ($rule.ReadyBy -ne $Phase) { continue }
        if (@($rule.AppliesToWeeks).Count -gt 0 -and $rule.AppliesToWeeks -notcontains $CurrentRunWeeks) { continue }

        $resolvedName = Resolve-TemplateText -Template ([string]$rule.Template) -Tokens $DateTokens
        $files.Add([PSCustomObject]@{
            Name         = $resolvedName
            ResolvedPath = $resolvedName
            ReadyBy      = [string]$rule.ReadyBy
            Source       = 'BackupValidationRules'
        })
    }

    return @($files.ToArray())
}

function Test-PreflightReady {
    param(
        [Parameter(Mandatory)]
        [hashtable]$Config,
        [Parameter(Mandatory)]
        [hashtable]$DateTokens,
        [Parameter(Mandatory)]
        [string]$Phase,
        [Parameter(Mandatory)]
        [string]$CurrentRunWeeks,
        [string]$BackupFolder,
        [string]$SourceFolder
    )

    $missingItems = New-Object 'System.Collections.Generic.List[object]'
    $invalidItems = New-Object 'System.Collections.Generic.List[object]'
    $readyItems = New-Object 'System.Collections.Generic.List[object]'

    if ($Phase -eq 'Analysis' -and $Config.Tasks) {
        foreach ($task in @($Config.Tasks)) {
            if (-not $task.Enabled) { continue }
            $taskName = [string]$task.Name

            try {
                $resolvedPath = Resolve-TaskInputPath -Task $task -Tokens $DateTokens
                if (Test-Path -LiteralPath $resolvedPath -PathType Leaf) {
                    $readyItems.Add([PSCustomObject]@{
                        Name         = $taskName
                        ExpectedPath = $resolvedPath
                        Source       = 'task-input'
                    })
                }
                else {
                    $missingItems.Add([PSCustomObject]@{
                        Name         = $taskName
                        ExpectedPath = $resolvedPath
                        Source       = 'task-input'
                    })
                }
            }
            catch {
                $invalidItems.Add([PSCustomObject]@{
                    Name         = $taskName
                    ExpectedPath = $null
                    Source       = 'task-input'
                    Error        = $_.Exception.Message
                })
            }
        }
    }

    $backupValidationConfig = Get-BackupValidationConfig -Config $Config
    if ($backupValidationConfig) {
        $preflightFiles = Get-PreflightFiles -BackupValidationConfig $backupValidationConfig -Phase $Phase -CurrentRunWeeks $CurrentRunWeeks -DateTokens $DateTokens

        foreach ($pf in $preflightFiles) {
            $checkDir = if ($Phase -eq 'Validate' -and $BackupFolder) {
                $BackupFolder
            }
            elseif ($Phase -eq 'Analysis' -and $SourceFolder) {
                $SourceFolder
            }
            else { $null }

            if (-not $checkDir) { continue }
            $fullPath = Join-Path $checkDir $pf.ResolvedPath

            if (Test-Path -LiteralPath $fullPath -PathType Leaf) {
                $readyItems.Add([PSCustomObject]@{
                    Name         = $pf.Name
                    ExpectedPath = $fullPath
                    Source       = 'fixed-file'
                })
            }
            else {
                $missingItems.Add([PSCustomObject]@{
                    Name         = $pf.Name
                    ExpectedPath = $fullPath
                    Source       = 'fixed-file'
                })
            }
        }
    }

    return [PSCustomObject]@{
        AllReady      = ($missingItems.Count -eq 0 -and $invalidItems.Count -eq 0)
        MissingItems  = @($missingItems.ToArray())
        InvalidItems  = @($invalidItems.ToArray())
        ReadyItems    = @($readyItems.ToArray())
    }
}

function Send-PreflightNotification {
    param(
        [Parameter(Mandatory)]
        [object]$NotificationConfig,
        [Parameter(Mandatory)]
        [object[]]$MissingItems,
        [object[]]$InvalidItems,
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

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add("<h3>WeCom Audit Preflight Failed - Phase $Phase</h3>")
    $lines.Add("<p>Date range: $StartDate - $EndDate</p>")

    if ($MissingItems.Count -gt 0) {
        $lines.Add("<h4>Missing Files ($($MissingItems.Count))</h4><ul>")
        foreach ($item in $MissingItems) {
            $lines.Add("<li><b>$($item.Name)</b>: $($item.ExpectedPath) <i>[$($item.Source)]</i></li>")
        }
        $lines.Add("</ul>")
    }

    if ($InvalidItems -and $InvalidItems.Count -gt 0) {
        $lines.Add("<h4>Invalid Items ($($InvalidItems.Count))</h4><ul>")
        foreach ($item in $InvalidItems) {
            $lines.Add("<li><b>$($item.Name)</b>: $($item.Error) <i>[$($item.Source)]</i></li>")
        }
        $lines.Add("</ul>")
    }

    $lines.Add("<p>Please prepare the required files and re-trigger the scheduled job.</p>")

    $subject = "[WeCom Audit] Preflight Failed - $Phase blocked ($StartDate - $EndDate)"
    $body = $lines -join "`n"
    $ccStr = ($NotificationConfig.CcRecipients | Where-Object { $_ }) -join ','
    if (-not $ccStr) { $ccStr = $NotificationConfig.OpsTeam[0] }

    Send-Mail `
        -From $NotificationConfig.From `
        -To $NotificationConfig.OpsTeam `
        -Cc $ccStr `
        -Subject $subject `
        -Body $body `
        -SmtpServer $NotificationConfig.SmtpServer `
        -KeyName $NotificationConfig.CertName `
        -Cert $NotificationConfig.Cert `
        -Port $NotificationConfig.Port `
        -LogFilePath $LogFilePath
}

Export-ModuleMember -Function Convert-ExactDate, Write-Log, Get-Cert, Get-LogFilePath, Get-VaultSecret, New-LazyLdapConnection, Close-LazyLdapConnection, Send-Mail, New-HtmlBody, New-DateTokenMap, Resolve-TemplateText, Get-OptionalObjectPropertyValue, Write-AnalysisSummaryJson, Assert-ConfigInputDirectories, Get-TaskResultByName, Get-TaskSummaryData, Get-BackupValidationConfig, Get-ExpectedBackupFiles, Test-BackupFolderContent, Format-BackupValidationText, Export-AnalysisReport, New-LdapOrFilter, Split-LdapBatches, Resolve-LdapSearchBase, Get-LdapUserByMail, Get-LdapUserById, Get-WeComAuditLogFolderName, Get-NormalizedFullPath, Test-PathWithinAllowedRoots, Test-SafeToDeleteSourceFile, Remove-SourceFileWithRetry, Invoke-SourceFileCleanup, Test-AllowedRootIsTooBroad, Resolve-SourceCleanupConfig, Assert-SourceCleanupConfig, Assert-TaskNameUniqueness, Resolve-AuditConfigPath, Resolve-AuditOutputRoot, Resolve-AuditInputRoot, Resolve-TaskInputPath, Resolve-NotificationConfig, Get-PreflightFiles, Test-PreflightReady, Send-PreflightNotification, Resolve-ScheduleCycle, Resolve-PhaseHandoff
