# These scripts are designed for the following purpose:  Two disparate AD domains sync up to a single 365 tenant.  These will move the accounts to an unsynced OU, restore the account and blank out the immutable ID, then move it to a synced OU on the other domain.
# Requirements: The accounts in the losing tenant and gaining tenant MUST HAVE THE SAME UPN.  FULL ADSync MUST BE RUN after this script is run.
# This is the 3rd script.  It takes the CSV from the 2nd and moves the accounts in the gaining domain to a syncing OU.

$targetOU = "OU=Synced OU,DC=CORP,DC=COM"
$users = Import-Csv -Path "C:\temp\movedusers.csv"
$logFile = "C:\temp\moveOUPrem.log"
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

$movedcsv | Export-Csv -Path "C:\temp\finalizedusers.csv" -NoTypeInformation