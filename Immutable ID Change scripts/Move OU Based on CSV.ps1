# These scripts are designed for the following purpose:  Two disparate AD domains sync up to a single 365 tenant.  These will move the accounts to an unsynced OU, restore the account and blank out the immutable ID, then move it to a synced OU on the other domain.
# Requirements: The accounts in the losing tenant and gaining tenant MUST HAVE THE SAME UPN.  Once this is run you MUST run a FULL SYNC from ADSync.
# This is the first script.  It requires a list of accounts to move in a CSV.  It will output all moved accounts to a CSV.  This file is needed for step 2.

$targetOU = "OU=Non-Synced OU,DC=CORP,DC=COM"
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