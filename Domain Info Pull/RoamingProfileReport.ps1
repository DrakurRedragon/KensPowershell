# This is designed to pull every user in the domain, then provide users with Roaming Profiles. 
$users = get-aduser -filter *
$profilepath = @()
foreach($user in $users)
{
$profilepath += get-aduser -Identity $user.DistinguishedName -Properties Name,ProfilePath,Enabled,DistinguishedName | where -Property ProfilePath -NotLike '' | Select Name,Enabled,Profilepath,DistinguishedName
}
$profilepath | Export-CSV -Path C:\Temp\ProfilePath.csv