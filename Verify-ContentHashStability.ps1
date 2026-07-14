#Requires -Version 5.1
<#
.SYNOPSIS
Static verifier: BU email Subject and Body construction must be deterministic.

.DESCRIPTION
Sprint-1 gate for the ledger-based BU email dedup design. The ledger relies
on SHA256(Subject + Body) being stable across runs of the same input data.
This script scans:

  1. Subject literals in wecom_mail_analysis.ps1 / wecom_devicelog_analysis.ps1
     Every $variable inside the literal must be in the allow-list
     (startDate, endDate, BU, env).

  2. The New-HtmlBody function body in modules/internal/Notification.ps1.
     Must not contain non-deterministic tokens (Get-Date, GUID, Random,
     hostname, PID, etc.).

  3. Assignments to body-related variables ($htmlBody, $tableHtml,
     $violationContent, $noViolationContent, ...) in the two analysis
     scripts. Same non-deterministic-token check.

  4. Diagnostic pass: for every Send-Mail call site in the two analysis
     scripts, report which variables are bound to -Subject and -Body so
     reviewers can trace them back manually.

Exit codes:
  0 = PASS (no findings above Info level, or Warn without -FailOnWarning)
  1 = FAIL (non-deterministic token in mail path, or -FailOnWarning + Warn)

.PARAMETER RepoRoot
Directory containing wecom_*_analysis.ps1 and wecom_analysis_comm.psm1.
Defaults to the parent of this tools directory.

.PARAMETER FailOnWarning
Treat Warn findings as failures.

.EXAMPLE
.\Verify-ContentHashStability.ps1

.EXAMPLE
.\Verify-ContentHashStability.ps1 -FailOnWarning
#>
[CmdletBinding()]
param(
    [string]$RepoRoot,
    [switch]$FailOnWarning
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $RepoRoot) {
    $RepoRoot = if ($PSScriptRoot) { Split-Path $PSScriptRoot -Parent } else { (Get-Location).Path }
}

$mailScript   = Join-Path $RepoRoot 'wecom_mail_analysis.ps1'
$deviceScript = Join-Path $RepoRoot 'wecom_devicelog_analysis.ps1'
$notificationPath = Join-Path $RepoRoot 'modules\internal\Notification.ps1'

foreach ($p in @($mailScript, $deviceScript, $notificationPath)) {
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
        throw "Required file not found: $p"
    }
}

# Non-deterministic token patterns and why each matters.
$badTokens = [ordered]@{
    'Get-Date'             = 'current time'
    '\[DateTime\]::Now'    = 'current time'
    '\[DateTime\]::UtcNow' = 'current time'
    '\[System\.Guid\]'     = 'random GUID'
    'New-Guid'             = 'random GUID'
    'Get-Random'           = 'random number'
    '\$PID\b'              = 'process ID varies per run'
    '\$env:COMPUTERNAME'   = 'hostname varies per machine'
    '\$env:USERNAME'       = 'user varies per invocation'
}

$allowedSubjectVars = @('startDate', 'endDate', 'BU', 'env')

$findings = [System.Collections.Generic.List[pscustomobject]]::new()

function Add-Finding {
    param(
        [ValidateSet('Info','Warn','Fail')][string]$Level,
        [string]$File,
        [int]$Line,
        [string]$Category,
        [string]$Detail
    )
    [void]$findings.Add([pscustomobject]@{
        Level    = $Level
        File     = (Split-Path -Leaf $File)
        Line     = $Line
        Category = $Category
        Detail   = $Detail
    })
}

function Test-LineRangeForBadTokens {
    param(
        [string]$File,
        [int]$FromLine,
        [int]$ToLine,
        [string]$Category
    )
    $lines = Get-Content -LiteralPath $File
    $from = [Math]::Max(1, $FromLine)
    $to   = [Math]::Min($lines.Count, $ToLine)
    for ($i = $from; $i -le $to; $i++) {
        $line = $lines[$i - 1]
        foreach ($tok in $badTokens.Keys) {
            if ($line -match $tok) {
                Add-Finding -Level 'Fail' -File $File -Line $i -Category $Category `
                    -Detail ("Contains '{0}' ({1}): {2}" -f $tok, $badTokens[$tok], $line.Trim())
            }
        }
    }
}

# ---------- 1. Subject literals ----------

$subjectRegex = '\$([Ss])ubject\s*=\s*(?:"([^"]*)"|''([^'']*)'')'

foreach ($script in @($mailScript, $deviceScript)) {
    $lines = Get-Content -LiteralPath $script
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $m = [regex]::Match($line, $subjectRegex)
        if (-not $m.Success) { continue }

        $literal = if ($m.Groups[2].Value) { $m.Groups[2].Value } else { $m.Groups[3].Value }
        $lineNo  = $i + 1

        if ([string]::IsNullOrEmpty($literal)) {
            Add-Finding -Level 'Info' -File $script -Line $lineNo -Category 'Subject' `
                -Detail 'Empty init (no analysis)'
            continue
        }

        $vars = [regex]::Matches($literal, '\$(?:\{([^}]+)\}|(\w+))') |
                ForEach-Object {
                    if ($_.Groups[1].Value) { $_.Groups[1].Value } else { $_.Groups[2].Value }
                } |
                Sort-Object -Unique

        $unknown = @($vars | Where-Object { $_ -notin $allowedSubjectVars })
        if ($unknown.Count -gt 0) {
            Add-Finding -Level 'Warn' -File $script -Line $lineNo -Category 'Subject' `
                -Detail ("Uses variable(s) not in allow-list [{0}]: {1}" -f `
                    ($allowedSubjectVars -join ','), ($unknown -join ', '))
        }
        else {
            $shown = if ($vars) { ($vars -join ',') } else { '(no vars)' }
            Add-Finding -Level 'Info' -File $script -Line $lineNo -Category 'Subject' `
                -Detail ("OK: uses only [{0}]" -f $shown)
        }
    }
}

# ---------- 2. New-HtmlBody function body ----------

$tokens = $null; $parseErrors = $null
$notificationAst = [System.Management.Automation.Language.Parser]::ParseFile(
    $notificationPath, [ref]$tokens, [ref]$parseErrors)
if (@($parseErrors).Count -gt 0) {
    throw "Notification module has parse errors: $($parseErrors[0].Message)"
}
$htmlBodyFunctions = @($notificationAst.FindAll({ param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
        $node.Name -eq 'New-HtmlBody'
    }, $true))

if ($htmlBodyFunctions.Count -eq 1) {
    $hbStart = $htmlBodyFunctions[0].Extent.StartLineNumber
    $hbEnd = $htmlBodyFunctions[0].Extent.EndLineNumber
    Test-LineRangeForBadTokens -File $notificationPath -FromLine $hbStart -ToLine $hbEnd `
        -Category 'New-HtmlBody'
    Add-Finding -Level 'Info' -File $notificationPath -Line $hbStart -Category 'New-HtmlBody' `
        -Detail ("Scanned lines {0}..{1}" -f $hbStart, $hbEnd)
}
else {
    Add-Finding -Level 'Fail' -File $notificationPath -Line 0 -Category 'New-HtmlBody' `
        -Detail "Expected exactly one New-HtmlBody function, found $($htmlBodyFunctions.Count)"
}

# ---------- 3. Body-related variable assignments in analysis scripts ----------

$contentVars = @(
    'noViolationContent', 'violationContent',
    'mailViolationContent', 'mailNoViolationContent',
    'DeviceViolationContent', 'tableHtml', 'htmlBody', 'rowsHtml'
)
$contentVarRegex = '\$(' + ($contentVars -join '|') + ')\s*='

foreach ($script in @($mailScript, $deviceScript)) {
    $lines = Get-Content -LiteralPath $script
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line -notmatch $contentVarRegex) { continue }
        foreach ($tok in $badTokens.Keys) {
            if ($line -match $tok) {
                Add-Finding -Level 'Fail' -File $script -Line ($i + 1) -Category 'BodyVar' `
                    -Detail ("Assignment contains '{0}' ({1}): {2}" -f `
                        $tok, $badTokens[$tok], $line.Trim())
            }
        }
    }
}

# ---------- 4. Send-Mail call sites: report -Subject / -Body bindings ----------

foreach ($script in @($mailScript, $deviceScript)) {
    $lines = Get-Content -LiteralPath $script
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $match = [regex]::Match($lines[$i], '\b(Send-Mail|Send-AuditBuMail)\b')
        if (-not $match.Success) { continue }
        $callName = $match.Groups[1].Value
        # Accumulate up to 8 continuation lines (backtick-terminated)
        $sb = New-Object 'System.Text.StringBuilder'
        for ($j = 0; $j -le 8 -and ($i + $j) -lt $lines.Count; $j++) {
            [void]$sb.AppendLine($lines[$i + $j])
            if ($lines[$i + $j] -notmatch '`\s*$') { break }
        }
        $blob = $sb.ToString()
        $subj = if ($blob -match '-Subject\s+(\S+)') { $Matches[1] } else { '?' }
        $body = if ($blob -match '-Body\s+(\S+)')    { $Matches[1] } else { '?' }
        Add-Finding -Level 'Info' -File $script -Line ($i + 1) -Category $callName `
            -Detail ("Binds -Subject={0} -Body={1}" -f $subj, $body)
    }
}

# ---------- Report ----------

$fail = @($findings | Where-Object { $_.Level -eq 'Fail' }).Count
$warn = @($findings | Where-Object { $_.Level -eq 'Warn' }).Count
$info = @($findings | Where-Object { $_.Level -eq 'Info' }).Count

Write-Host ''
Write-Host '=== ContentHash Stability Verification ===' -ForegroundColor Cyan
Write-Host ("Fails: {0}    Warns: {1}    Info: {2}" -f $fail, $warn, $info)
Write-Host ''

$sortOrder = @{ Fail = 0; Warn = 1; Info = 2 }
$findings |
    Sort-Object @{Expression = { $sortOrder[$_.Level] }}, File, Line |
    Format-Table -AutoSize Level, File, Line, Category, Detail |
    Out-String |
    Write-Host

if ($fail -gt 0) {
    Write-Host 'RESULT: FAIL - non-deterministic content detected in mail path' -ForegroundColor Red
    exit 1
}
if ($warn -gt 0 -and $FailOnWarning) {
    Write-Host 'RESULT: WARN (treated as FAIL due to -FailOnWarning)' -ForegroundColor Yellow
    exit 1
}
if ($warn -gt 0) {
    Write-Host 'RESULT: WARN - review flagged items before enabling ContentHash-based ledger' -ForegroundColor Yellow
    exit 0
}
Write-Host 'RESULT: PASS - Subject and Body construction is deterministic. ContentHash will be stable.' -ForegroundColor Green
exit 0
