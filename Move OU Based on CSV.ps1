$targetOU = "OU=Moved to Premier,OU=No Sync to Premier Tenant,DC=HANDCENTERS,DC=COM"
$users = Import-Csv -Path "C:\temp\movelist.csv"
$logFile = "C:\temp\moveOU.log"
$movedcsv = @()

foreach ($user in $users) {
    $UserPrincipal = get-aduser -Identity $user.Name -Properties UserPrincipalName
    $userPrincipalName = $UserPrincipal.UserPrincipalName
    $currentOU = (Get-ADUser -Identity $user.Name).DistinguishedName

    if ($currentOU -ne $targetOU) {
        Move-ADObject -Identity $user.Name -TargetPath $targetOU
        Write-Host "Moved user $($user.Name) to $targetOU."
        Add-Content -Path $logFile -Value "Moved user $($user.Name) from $currentOU with UPN $userPrincipalName to $targetOU on $(Get-Date)."
        $movedcsv += [PSCustomObject]@{
            Name               = $user.Name
            UserPrincipalName  = $userPrincipalName
            CurrentOU          = $currentOU
            TargetOU           = $targetOU
            Date               = Get-Date
        }
    }
    else {
        Write-Host "User $($user.Name) is already in the target OU."
        Add-Content -Path $logFile -Value "User $($user.Name) is already in the target OU $targetOU with UPN $userPrincipalName on $(Get-Date)."
    }
}

$movedcsv | Export-Csv -Path "C:\temp\movedusers.csv" -NoTypeInformation