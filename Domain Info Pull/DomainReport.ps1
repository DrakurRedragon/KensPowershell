# This is designed as a quick information gather on a domain from a DC.  It pulls and creates files for: Every server, OS count, GPOs individually in their own HTML files, accounts login scripts and accounts without them, accounts that have roaming profiles, computers and users that haven't checked in in over 90 days, and every member of a group containing Admin in the name

# Define the folder path
$folderPath = "C:\Temp\DomainReport"

# Check if the folder exists.  A second check will create the subfolder for GPO
if (-not (Test-Path -Path $folderPath -PathType Container)) {
    try {
        # Create the folder
        New-Item -Path $folderPath -ItemType Directory -Force | Out-Null
        Write-Host "Folder created: $folderPath"
    }
    catch {
        Write-Host "Error creating folder: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "Folder already exists: $folderPath"
}
$folderPath2 = $folderpath + "\GPO"

# Check if the folder exists
if (-not (Test-Path -Path $folderPath2 -PathType Container)) {
    try {
        # Create the folder
        New-Item -Path $folderPath2 -ItemType Directory -Force | Out-Null
        Write-Host "Folder created: $folderPath2"
    }
    catch {
        Write-Host "Error creating folder: $($_.Exception.Message)" -ForegroundColor Red
    }
}
else {
    Write-Host "Folder already exists: $folderPath2"
}

# All servers in AD
get-adcomputer -filter {OperatingSystem -like "*Server*"} -Properties Name, OperatingSystem, IPV4Address, Enabled | sort -Property Name | export-csv -path $folderPath\server.csv
Write-Host "Server file created"

# All OS in AD by count
Get-ADComputer -Filter {Enabled -eq "True"} -Properties * | Group-Object -Property OperatingSystem,OperatingSystemVersion | Select-Object Name,Count | Sort-Object Name | Export-CSV -Path $folderPath\oscount.csv
Write-Host "OS Count file created"

# GPO Reporting.  Creates HTML Files for each GPO and markes which ones are disabled as well
$gpos = get-gpo -all
foreach ($gpo in $gpos)
{
if ($gpo.GpoStatus -like "*Disable*")
{$filename = "_Disabled " + $gpo.DisplayName}
else
{$filename = $gpo.DisplayName}
get-gporeport -guid $gpo.id -ReportType html -path $folderPath\GPO\$filename.html
}
Write-Host "GPO files created"

# Creates 2 CSVs for users with login scripts and users without
$users = get-aduser -filter *
$scriptpath = @()
$scriptpath2 = @()
foreach($user in $users)
{
$scriptpath += get-aduser -Identity $user.DistinguishedName -Properties Name,ScriptPath,Enabled,DistinguishedName | where -Property ScriptPath -NotLike '' | Select Name,Enabled,Scriptpath,DistinguishedName
}
$scriptpath | Export-CSV -Path $folderPath\Login.csv
foreach($user in $users)
{
$scriptpath2 += get-aduser -Identity $user.DistinguishedName -Properties Name,ScriptPath,Enabled,DistinguishedName | where -Property ScriptPath -Like '' | Select Name,Enabled,Scriptpath,DistinguishedName
}
$scriptpath2 | Export-CSV -Path $folderPath\NoLogin.csv
Write-Host "Login and No Login script files created"

# Creates a CSV of all computers that haven't checked into AD in over 90 days
$DaysInactive = 90
$time = (Get-Date).Adddays(-($DaysInactive))
$stalecomputers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time -and Enabled -eq "True"} -properties * | Select Name,Enabled,LastLogonDate
$stalecomputers | Export-CSV -path "$folderPath\StaleComputer.csv"
Write-Host "Stale Computer file created"

# Creates a CSV of all Users that haven't logged in in the past 90 days
$DaysInactive = 90
$time = (Get-Date).Adddays(-($DaysInactive))
$staleusers = Get-ADUser -Filter {LastLogonTimeStamp -lt $time -and Enabled -eq "True"} -properties * | Select Name,Enabled,LastLogonDate
$staleusers | Export-CSV -path "$folderPath\StaleUser.csv"
Write-Host "Stale User file created"

# Generates a simple text file with all groups that say Admin and their members from local AD
$groups = get-adgroup -filter {Name -like "*Admin*"}
$members = @()
foreach ($group in $groups)
{
$members += "============="
$members += $group.Name
$members += "============="
$members += Get-ADGroupMember -Identity $group.Name | Select Name,samAccountName
}
$members >> $folderPath\adminusers.txt
Write-Host "Admin Group file created"

# This is designed to pull every user in the domain, then provide users with Roaming Profiles. 
$users = get-aduser -filter *
$profilepath = @()
foreach($user in $users)
{
$profilepath += get-aduser -Identity $user.DistinguishedName -Properties Name,ProfilePath,Enabled,DistinguishedName | where -Property ProfilePath -NotLike '' | Select Name,Enabled,Profilepath,DistinguishedName
}
$profilepath | Export-CSV -Path $folderPath\ProfilePath.csv
Write-Host "Roaming Profile file created"