$username = import-csv c:\temp\list2.csv
$results = @()
foreach ($user in $username.Name) {
    $userinfo = get-aduser -identity $user -properties LastLogonDate
    $results += [PSCustomObject]@{
        Name = $userinfo.Name
        LastLogonDate = $userinfo.LastLogonDate
        UserPrincipalName = $userinfo.UserPrincipalName
        Enabled = $userinfo.Enabled
        DistinguishedName = $userinfo.DistinguishedName
    }
}
$results | export-csv c:\temp\userLLD.csv -notype