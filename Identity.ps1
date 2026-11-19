function Get-MsimcLdapUserByAccount {
    param(
        [Parameter(Mandatory = $true)]$LazyConnection,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Accounts,
        [string]$SearchBase,
        [int]$SearchTimeoutSeconds=30,
        [ValidateRange(1, 1000)][int]$BatchSize = 20
    )

    Add-Type -AssemblyName System.DirectoryServices.Protocols
    $lookup = @{
        Valid = @{}
        SamAccountNameByCn = @{}
        Invalid = [System.Collections.Generic.List[string]]::new()
    }
    $unique = @($Accounts | ForEach-Object { $_.Trim().ToLowerInvariant() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($unique.Count -eq 0) { return $lookup }
    $connection = $LazyConnection.Value
    $connection.Timeout = [TimeSpan]::FromSeconds($SearchTimeoutSeconds)
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
        # COD Login ID and unmapped device accounts both identify users by CN.
        $usersByCn = @{}
        $clauses = foreach ($cn in $batch) {
            $usersByCn[$cn] = [System.Collections.Generic.List[object]]::new()
            '(cn={0})' -f (ConvertTo-LdapFilterEscaped $cn)
        }
        $filter = '(&(objectClass=user)(|{0}))' -f ($clauses -join '')
        $request = [System.DirectoryServices.Protocols.SearchRequest]::new(
            $SearchBase, $filter, [System.DirectoryServices.Protocols.SearchScope]::Subtree,
            [string[]]@('sAMAccountName', 'cn', 'mail', 'division', 'legalentity'))
        $response = $connection.SendRequest($request)
        if ($response.ResultCode -ne [System.DirectoryServices.Protocols.ResultCode]::Success) {
            throw "MSIMC LDAP account query failed: $($response.ResultCode)"
        }
        foreach ($entry in $response.Entries) {
            $identity = [ordered]@{ SamAccountName = $null; CN = $null; Mail = $null; Division = $null; LegalEntity = $null }
            foreach ($pair in @(@('sAMAccountName', 'SamAccountName'), @('cn', 'CN'), @('mail', 'Mail'), @('division', 'Division'), @('legalentity', 'LegalEntity'))) {
                if ($entry.Attributes[$pair[0]]) { $identity[$pair[1]] = [string]$entry.Attributes[$pair[0]][0] }
            }
            if (-not $identity.CN) { continue }
            $matchedCn = $identity.CN.Trim().ToLowerInvariant()
            if ($usersByCn.ContainsKey($matchedCn)) {
                $usersByCn[$matchedCn].Add([pscustomobject]$identity)
            }
        }
        foreach ($cn in $batch) {
            # Different OUs may contain users with the same CN.
            $users = $usersByCn[$cn]
            if ($users.Count -ne 1 -or [string]::IsNullOrWhiteSpace($users[0].SamAccountName)) {
                $lookup.Invalid.Add($cn)
                continue
            }
            $samAccountName = $users[0].SamAccountName.Trim().ToLowerInvariant()
            $lookup.Valid[$samAccountName] = $users[0]
            $lookup.SamAccountNameByCn[$cn] = $samAccountName
        }
    }
    return $lookup
}
