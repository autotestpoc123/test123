# Config.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

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

    # Dynamic rules describe message files produced AFTER Phase 1 by ops; they are
    # semantically ReadyBy='Validate'. Config may override but typically should not.
    $readyBy = if ($null -ne $Item.ReadyBy -and [string]$Item.ReadyBy) {
        [string]$Item.ReadyBy
    }
    else {
        'Validate'
    }

    return [PSCustomObject]@{
        BaseName        = $baseName
        SummaryTaskName = [string]$Item.SummaryTaskName
        Source          = if ($null -ne $Item.Source -and [string]$Item.Source) { [string]$Item.Source } else { 'generated' }
        Required        = if ($null -ne $Item.Required) { [bool]$Item.Required } else { $true }
        AppliesToWeeks  = $appliesToWeeks
        ReadyBy         = $readyBy
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

<#
.SYNOPSIS
Resolves the folder where source audit files are staged for analysis and validation.
.DESCRIPTION
Returns the configured source folder used by preflight, analysis input discovery,
source-mode validation, and archive copy targets. Resolution order is:
WECOM_AUDIT_SOURCE_FOLDER environment variable, then Config.SourceFolder. If
neither is set this throws rather than silently falling back to a legacy
location - source files are staged in a dedicated folder, so a wrong target
would make every expected file look "missing".
.PARAMETER Config
Imported audit configuration hashtable. Must contain SourceFolder unless the
WECOM_AUDIT_SOURCE_FOLDER environment variable is set.
.EXAMPLE
PS> Resolve-AuditSourceFolder -Config $config
Returns Config.SourceFolder, e.g. C:\addin_deploy_cert\wecom_audit_log\source.
.NOTES
Fail-fast by design: an unconfigured SourceFolder is a deployment error, not a
case to paper over with a default.
#>
function Resolve-AuditSourceFolder {
    param(
        [hashtable]$Config
    )

    if ($env:WECOM_AUDIT_SOURCE_FOLDER) {
        return [string]$env:WECOM_AUDIT_SOURCE_FOLDER
    }

    if ($Config -and $Config.ContainsKey('SourceFolder') -and $Config.SourceFolder) {
        return [string]$Config.SourceFolder
    }

    throw "SourceFolder is not configured. Set 'SourceFolder' in config (e.g. 'C:\addin_deploy_cert\wecom_audit_log\source') or the WECOM_AUDIT_SOURCE_FOLDER environment variable."
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
