# This is designed to pull every user in the domain, then provide users with Home Directories set in AD. 
$users = get-aduser -filter *
$homedrive = @()
foreach($user in $users)
{
$homedrive += get-aduser -Identity $user.DistinguishedName -Properties Name,HomeDrive,HomeDirectory,Enabled,DistinguishedName | where -Property HomeDirectory -NotLike '' | Select Name,Enabled,DistinguishedName,HomeDrive,HomeDirectory
}
$homedrive | Export-CSV -Path C:\Temp\HomeDir.csv