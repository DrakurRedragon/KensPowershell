# Install-module azureAD
# import-module azureAD

connect-azuread

$uname = Read-Host "Please enter the username"

$olddc = Read-Host "Enter old DC IP"
$newdc = Read-Host "Enter new Domain DC IP"

$azureadid = get-azureaduser -SearchString $uname | select objectID
$azureadupn = get-azureaduser -ObjectId $azureadid.ObjectId | select UserPrincipalName
$user = Get-ADUser -Identity eroman
$guid = $user.ObjectGUID
$immutableId = [System.Convert]::ToBase64String($guid.ToByteArray())

set-azureaduser -ObjectId $azureadid.ObjectId -ImmutableID $immutableID

graph scope user.readwrite.all
Update-MgUser -UserId "user@domain.com" -OnPremisesImmutableId $null

Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/tcaltune@handcenters.com" -Body @{onPremisesImmutableId = $null}
Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/Users/tcaltune@handcenters.com?select=id,displayName,onPremisesImmutableID"

