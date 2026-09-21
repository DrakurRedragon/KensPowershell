# This is designed to pull every user in the domain, then filter them into two files, one without login scripts and one with login scripts. 
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
