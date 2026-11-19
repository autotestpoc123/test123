  if ($BU -eq 'MSIMC') {
        if (-not $IdentityMapping) {
            if (-not $IdentityMappingPath) { $IdentityMappingPath = Join-Path (Split-Path $parentFolderPath -Parent) 'convert_table.xlsx' }
            $IdentityMapping = Import-MsimcIdentityMapping -Path $IdentityMappingPath
        }
        $resolutions = Resolve-MsimcIdentity -Identities @($filterIds -split ';') -Kind Account -Mapping $IdentityMapping
        $accounts = @($resolutions.Values | ForEach-Object { $_.CorporateAccount } | Sort-Object -Unique)
        $corporateLookup = Get-MsimcLdapUserByAccount -LazyConnection $lazyConn -Accounts $accounts
        $deviceUsersByAccount = @{}
        foreach ($originalAccount in $resolutions.Keys) {
            $cn = $resolutions[$originalAccount].CorporateAccount
            $samAccountName = $corporateLookup.SamAccountNameByCn[$cn]
            if ($samAccountName) {
                $deviceUsersByAccount[$originalAccount] = $corporateLookup.Valid[$samAccountName]
            }
        }
    }
    else { $cNamelookup = Get-LdapUserById -LazyConnection $lazyConn -UserId $filterIds }
}
