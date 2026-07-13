# Core.ps1 - dot-sourced by wecom_analysis_comm.psm1 (single module scope).
# FUNCTIONS ONLY: no top-level statements in internal files (load-order-free).
# Moved verbatim from the monolith - see Verify-ModuleSplit.ps1 hash parity.

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
Builds the full template-token map (date tokens plus resolved roots).
.DESCRIPTION
Single entry point that combines New-DateTokenMap with the resolved InputRoot and
SourceFolder, so every script gets an identical token set. Centralizing this
removes the per-script "resolve roots then bolt the tokens on" ritual and the
drift risk it carries - a missed SourceFolder key would leave '{SourceFolder}'
unresolved inside task InputDirectory paths.
.PARAMETER Config
Imported audit configuration hashtable.
.PARAMETER StartDate
Cycle start date (yyyyMMdd).
.PARAMETER EndDate
Cycle end date (yyyyMMdd).
.EXAMPLE
PS> $tokens = New-AuditTokenMap -Config $config -StartDate '20260514' -EndDate '20260528'
PS> $tokens.SourceFolder
.NOTES
Throws (via Resolve-AuditSourceFolder) when SourceFolder is not configured.
#>
function New-AuditTokenMap {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Config,
        [Parameter(Mandatory = $true)]
        [string]$StartDate,
        [Parameter(Mandatory = $true)]
        [string]$EndDate
    )

    $tokens = New-DateTokenMap -StartDate $StartDate -EndDate $EndDate
    $tokens.InputRoot    = Resolve-AuditInputRoot   -Config $Config
    $tokens.SourceFolder = Resolve-AuditSourceFolder -Config $Config
    return $tokens
}
