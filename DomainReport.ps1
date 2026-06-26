# This is designed as a quick information gather on a domain from a DC.  It pulls and creates files for: Every server, OS count, GPOs individually in their own HTML files, login scripts and accounts withou them, computers and users that haven't checked in in over 90 days, and every member of a group containing Admin in the name

get-adcomputer -filter {OperatingSystem -like "*Server*"} -Properties Name, OperatingSystem, IPV4Address, Enabled | sort -Property Name | export-csv -path c:\Temp\server.csv

Get-ADComputer -Filter {Enabled -eq "True"} -Properties * | Group-Object -Property OperatingSystem,OperatingSystemVersion | Select-Object Name,Count | Sort-Object Name | Export-CSV -Path c:\Temp\oscount.csv

$gpos = get-gpo -all
foreach ($gpo in $gpos)
{
if ($gpo.GpoStatus -like "*Disable*")
{$filename = "_Disabled " + $gpo.DisplayName}
else
{$filename = $gpo.DisplayName}
get-gporeport -guid $gpo.id -ReportType html -path c:\temp\GPO\$filename.html
}

$users = get-aduser -filter *
$scriptpath = @()
$scriptpath2 = @()
foreach($user in $users)
{
$scriptpath += get-aduser -Identity $user.DistinguishedName -Properties Name,ScriptPath,Enabled,DistinguishedName | where -Property ScriptPath -NotLike '' | Select Name,Enabled,Scriptpath,DistinguishedName
}
$scriptpath | Export-CSV -Path C:\Temp\Login.csv
foreach($user in $users)
{
$scriptpath2 += get-aduser -Identity $user.DistinguishedName -Properties Name,ScriptPath,Enabled,DistinguishedName | where -Property ScriptPath -Like '' | Select Name,Enabled,Scriptpath,DistinguishedName
}
$scriptpath2 | Export-CSV -Path C:\Temp\NoLogin.csv

$DaysInactive = 90
$time = (Get-Date).Adddays(-($DaysInactive))
$stalecomputers = Get-ADComputer -Filter {LastLogonTimeStamp -lt $time -and Enabled -eq "True"} -properties * | Select Name,Enabled,LastLogonDate
$stalecomputers | Export-CSV -path "c:\Temp\StaleComputer.csv"

$DaysInactive = 90
$time = (Get-Date).Adddays(-($DaysInactive))
$staleusers = Get-ADUser -Filter {LastLogonTimeStamp -lt $time -and Enabled -eq "True"} -properties * | Select Name,Enabled,LastLogonDate
$staleusers | Export-CSV -path "c:\Temp\StaleUser.csv"

$groups = get-adgroup -filter {Name -like "*Admin*"}
$members = @()
foreach ($group in $groups)
{
$members += "============="
$members += $group.Name
$members += "============="
$members += Get-ADGroupMember -Identity $group.Name | Select Name,samAccountName
}
$members >> c:\temp\adminusers.txt
