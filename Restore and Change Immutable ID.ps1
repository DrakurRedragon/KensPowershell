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
