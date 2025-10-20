# Install-module azureAD
# import-module azureAD

connect-azuread

$uname = Read-Host "Please enter the username"

$olddc = Read-Host "Enter old DC IP"
$newdc = Read-Host "Enter new Domain DC IP"

$azureadid = get-azureaduser -SearchString $uname | select objectID
$azureadupn = get-azureaduser -ObjectId $azureadid.ObjectId | select UserPrincipalName
$user = Get-ADUser -Identity $uname
$guid = $user.ObjectGUID
$immutableId = [System.Convert]::ToBase64String($guid.ToByteArray())

set-azureaduser -ObjectId $azureadid.ObjectId -ImmutableID $immutableID