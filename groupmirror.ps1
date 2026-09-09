# Using Graph a graph commandlet, this pulls the members of a group based on the ID.  Change this in the first line.  The second line is for the receiving group.  It will then put every single member of the first group into the second.
$memberlist = get-mggroupmember -groupID bab0be4c-9d02-4c97-80ec-32e82b264e79 -all
$groupid = "9c951fe6-af0d-4444-a89d-69f99b2ae951"
foreach ($user in $memberlist)
{
	$userid = $user.id

	Invoke-MGGraphRequest -Method POST -URI "https://graph.microsoft.com/v1.0/groups/$groupid/members/`$ref" -ContentType application/json -Body @{"@odata.id"= "https://graph.microsoft.com/v1.0/directoryObjects/$userid"}
}
