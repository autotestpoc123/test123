# Historical MSIMC identifiers are translated before directory validation.
function Import-MsimcIdentityMapping {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "MSIMC identity mapping file not found: $Path"
    }
    if (-not (Get-Command Import-Excel -ErrorAction SilentlyContinue)) {
        Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'ImportExcel') -ErrorAction Stop
    }
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    $rows = @(Import-Excel -Path $Path -ErrorAction Stop)
    if ($rows.Count -eq 0) { throw "MSIMC identity mapping is empty: $Path" }
    foreach ($column in @('User Name', 'WeCom Account ID', 'COD Login ID', 'WeComMail', 'CODMail')) {
        if ($column -notin $rows[0].PSObject.Properties.Name) {
            throw "MSIMC identity mapping is missing column '$column': $Path"
        }
    }
    $mapping = @{ Accounts = @{}; Mails = @{}; Path = [IO.Path]::GetFullPath($Path); Hash = $hash }
    foreach ($row in $rows) {
        $account = ([string]$row.'WeCom Account ID').Trim().ToLowerInvariant()
        $mail = ([string]$row.WeComMail).Trim().ToLowerInvariant()
        $targetAccount = ([string]$row.'COD Login ID').Trim().ToLowerInvariant()
        $targetMail = ([string]$row.CODMail).Trim().ToLowerInvariant()
        if (-not ($account -or $mail -or $targetAccount -or $targetMail)) { continue }
        if (-not $account -or -not $targetAccount -or $targetAccount -match '[@\\/]' -or
            $mail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$' -or
            $targetMail -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            throw "MSIMC identity mapping contains an invalid account/email row for '$account'."
        }
        $identity = [pscustomobject]@{ CorporateAccount = $targetAccount; CorporateEmail = $targetMail }
        foreach ($index in @(@{ Lookup = $mapping.Accounts; Key = $account }, @{ Lookup = $mapping.Mails; Key = $mail })) {
            if ($index.Lookup.ContainsKey($index.Key)) {
                $existing = $index.Lookup[$index.Key]
                if ($existing.CorporateAccount -ne $targetAccount -or $existing.CorporateEmail -ne $targetMail) {
                    throw "MSIMC identity mapping conflicts for '$($index.Key)'."
                }
            }
            $index.Lookup[$index.Key] = $identity
        }
    }
    if ($mapping.Accounts.Count -eq 0) { throw "MSIMC identity mapping has no usable rows: $Path" }
    if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ne $hash) {
        throw 'MSIMC identity mapping changed while it was being read.'
    }
    return $mapping
}

function Resolve-MsimcIdentity {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Identities,
        [Parameter(Mandatory = $true)][ValidateSet('Account', 'Mail')][string]$Kind,
        [Parameter(Mandatory = $true)][hashtable]$Mapping,
        [string]$MailDomainSuffix
    )

    $lookup = if ($Kind -eq 'Account') { $Mapping.Accounts } else { $Mapping.Mails }
    $resolved = @{}
    foreach ($value in $Identities) {
        $key = ([string]$value).Trim().ToLowerInvariant()
        if (-not $key -or $resolved.ContainsKey($key)) { continue }
        $targetAccount = $null
        $targetMail = $null
        $method = 'StandardRule'
        if ($lookup.ContainsKey($key)) {
            $targetAccount = $lookup[$key].CorporateAccount
            $targetMail = $lookup[$key].CorporateEmail
            $method = 'LegacyMapping'
        }
        elseif ($Kind -eq 'Account') { $targetAccount = $key }
        else { $targetMail = $key -replace '@.*$', $MailDomainSuffix }
        $resolved[$key] = [pscustomobject]@{
            CorporateAccount = $targetAccount; CorporateEmail = $targetMail; ResolutionMethod = $method
        }
    }
    return $resolved
}

function Get-MsimcLdapUserByAccount {
    param(
        [Parameter(Mandatory = $true)]$LazyConnection,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Accounts,
        [string]$SearchBase,
        [ValidateRange(1, 1000)][int]$BatchSize = 20
    )

    $valid = @{}
    $invalid = [System.Collections.Generic.List[string]]::new()
    $unique = @($Accounts | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($unique.Count -eq 0) { return @{ Valid = $valid; Invalid = $invalid } }
    $connection = $LazyConnection.Value
    if (-not $SearchBase) {
        $request = [System.DirectoryServices.Protocols.SearchRequest]::new(
            '', '(objectClass=*)', [System.DirectoryServices.Protocols.SearchScope]::Base, @('defaultNamingContext'))
        $response = $connection.SendRequest($request)
        if ($response.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success) { throw 'MSIMC LDAP RootDSE query failed.' }
        $SearchBase = [string]$response.Entries[0].Attributes['defaultNamingContext'][0]
        if (-not $SearchBase) { throw 'MSIMC LDAP defaultNamingContext is empty.' }
    }
    for ($offset = 0; $offset -lt $unique.Count; $offset += $BatchSize) {
        $last = [Math]::Min($offset + $BatchSize - 1, $unique.Count - 1)
        $batch = @($unique[$offset..$last])
        $clauses = foreach ($account in $batch) { '(sAMAccountName={0})' -f (ConvertTo-LdapFilterEscaped $account) }
        $filter = '(&(objectClass=user)(|{0}))' -f ($clauses -join '')
        $request = [System.DirectoryServices.Protocols.SearchRequest]::new(
            $SearchBase, $filter, [System.DirectoryServices.Protocols.SearchScope]::Subtree,
            @('sAMAccountName', 'mail', 'division', 'legalentity'))
        $response = $connection.SendRequest($request)
        if ($response.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success) {
            throw "MSIMC LDAP account query failed: $($response.ResultCode)"
        }
        foreach ($entry in $response.Entries) {
            $identity = [ordered]@{ SamAccountName = $null; Mail = $null; Division = $null; LegalEntity = $null }
            foreach ($pair in @(@('sAMAccountName', 'SamAccountName'), @('mail', 'Mail'), @('division', 'Division'), @('legalentity', 'LegalEntity'))) {
                if ($entry.Attributes[$pair[0]]) { $identity[$pair[1]] = [string]$entry.Attributes[$pair[0]][0] }
            }
            if ($identity.SamAccountName) { $valid[$identity.SamAccountName.ToLowerInvariant()] = [pscustomobject]$identity }
        }
        foreach ($account in $batch) { if (-not $valid.ContainsKey($account)) { $invalid.Add($account) } }
    }
    return @{ Valid = $valid; Invalid = $invalid }
}
