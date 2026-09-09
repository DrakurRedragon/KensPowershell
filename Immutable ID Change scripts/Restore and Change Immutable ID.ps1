# These scripts are designed for the following purpose:  Two disparate AD domains sync up to a single 365 tenant.  These will move the accounts to an unsynced OU, restore the account and blank out the immutable ID, then move it to a synced OU on the other domain.
# Requirements: The accounts in the losing tenant and gaining tenant MUST HAVE THE SAME UPN.  ADSync MUST BE RUN before this script is run.
# This is the second script.  It takes the CSV from the first, then uses Graph to restore the user (since it gets soft deleted) as a cloud account, then wips the Immutable ID so that the gaining tenant can take it over.  It will restore accounts as cloud with their previous passwords.  If there are failures, they MAY need to be restored manually.  This will provide an export that will feed into the 3rd script.
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"
$logFile = "C:\temp\ImmutableClear.log"

foreach ($user in (import-csv c:\temp\movedusers.csv)) {
    $UserPrincipal = $user.UserPrincipalName
    $usertorestore = get-mgdirectorydeleteditemasuser | Where-Object {$_.Mail -eq $UserPrincipal}
    Restore-MgDirectoryDeletedItem -DirectoryObjectId $usertorestore.Id
    Start-Sleep -seconds 2
    Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/Users/$UserPrincipal" -Body @{onPremisesImmutableID = $null}
    Write-Host "Restored user $($user.Name) and cleared ImmutableID for $UserPrincipal."
    Add-Content -Path $logFile -Value "Restored user $($user.Name) with UPN $UserPrincipal and cleared ImmutableID on $(Get-Date)."
}
