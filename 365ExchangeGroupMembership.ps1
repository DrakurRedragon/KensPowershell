# This script crawls a 365 environment and pulls distribution groups and 365 group members into excel workbooks, with separate sheets for each group.  It will check and install the ImportExcel and ExchangeOnlineManagement powershell modules.  You do need to have an account with the appropriate reader permissions in the tenant you are connecting to.

# Change the Following Logfile destination entries:

$365logfilename = "c:\Temp\m365group.xlsx"
$distrologfilename = "C:\Temp\DistroGroup.xlsx"

# Checks for and installs ExchangeOnlineManagement and ImportExcel

$module = get-InstalledModule
$name = "ImportExcel"

if ($module.name -like $name)
{
Write-Host "$name is Installed" -ForegroundColor Green
}
else
{
Write-Host "$name was not Installed.  Installing" -ForegroundColor Red
Install-Module $name -scope CurrentUser
Write-Host "$name is now Installed." -ForegroundColor Green
}

$module2 = get-Module
if ($module2.name -like $name)
{
Write-Host "$name is Imported" -ForegroundColor Green
}
else
{
Write-Host "$name was not imported.  Importing" -ForegroundColor Red
Import-Module $name
Write-Host "$name is now imported." -ForegroundColor Green
}

$name = "ExchangeOnlineManagement"

if ($module.name -like $name)
{
Write-Host "$name is Installed" -ForegroundColor Green
}
else
{
Write-Host "$name was not Installed.  Installing" -ForegroundColor Red
Install-Module $name -scope CurrentUser
Write-Host "$name is now Installed." -ForegroundColor Green
}

if ($module2.name -like $name)
{
Write-Host "$name is Imported" -ForegroundColor Green
}
else
{
Write-Host "$name was not imported.  Importing" -ForegroundColor Red
Import-Module $name
Write-Host "$name is now imported." -ForegroundColor Green
}

Connect-ExchangeOnline

# This starts the 365 groups.  It will export each group with its members as a separate sheet in an excel workbook.  The filename and path is listed above at the beginning.

$groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName

foreach ($group in $groups) {

    $members = Get-UnifiedGroupLinks -Identity $group.Identity -LinkType Members -ResultSize Unlimited

    if ($members) {
    $memberarray =@()
        foreach ($member in $members) {
            $displayName = $member.DisplayName
            $upn = $member.PrimarySmtpAddress  # For most mail-enabled objects, this aligns with UPN for users

            $memberArray = $memberarray += "$displayName,$upn" 
            
        }
    $memberarray | Export-Excel -WorksheetName "$group" -Path $365logfilename -BoldTopRow -AutoSize -Title "$group" -TitleSize 16
    } else {
        "(no members)" | Export-Excel -WorksheetName "$group" -Path $365logfilename -BoldTopRow -AutoSize -Title "$group" -TitleSize 16
    }
}

# This starts the distro groups.  Same as with the 365 groups.

$distrogroups = Get-DistributionGroup -ResultSize Unlimited | Sort-Object DisplayName

foreach ($distrogroup in $distrogroups) {

    $members = Get-DistributionGroupMember -Identity $distrogroup.Name -ResultSize Unlimited

    if ($members) {
    $memberarray =@()
        foreach ($member in $members) {
            $displayName = $member.DisplayName
            $upn = $member.PrimarySmtpAddress  # For most mail-enabled objects, this aligns with UPN for users

            $memberArray = $memberarray += "$displayName,$upn" 
            
        }
    $memberarray | Export-Excel -WorksheetName "$distrogroup" -Path $distrologfilename -BoldTopRow -AutoSize -Title "$distrogroup" -TitleSize 16
    } else {
        "(no members)" | Export-Excel -WorksheetName "$distrogroup" -Path $distrologfilename -BoldTopRow -AutoSize -Title "$distrogroup" -TitleSize 16
    }
}

Disconnect-ExchangeOnline -Confirm:$false

Write-Host "Done. M365 Output written to $365logfilename and Distro Output written to $distrologfilename" -ForegroundColor Green