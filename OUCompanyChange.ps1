$OU = "OU=HOPCo,DC=corp,DC=hopco,DC=com"
$users = Get-ADUser -SearchBase $OU -SearchScope Subtree -Filter * -Properties Company, SamAccountName, DistinguishedName

foreach ($user in $users) {
    $currentValue = $user.Company
    $sam = $user.SamAccountName
if ($currentValue -not $NewCompanyValue) {Set-ADUser -Identity $user.DistinguishedName -Company $NewCompanyValue}
}