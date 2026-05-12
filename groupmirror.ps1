$memberlist = get-mggroupmember -groupID bae545c7-0844-4cc8-8c02-e7ff6d868048 -all
$groupid = "2f53cffb-d130-44af-91c1-aa9b2aa44c8f"
foreach ($user in $memberlist)
{
	$userid = $user.id

	Invoke-MGGraphRequest -Method POST -URI "https://graph.microsoft.com/v1.0/groups/$groupid/members/`$ref" -ContentType application/json -Body @{"@odata.id"= "https://graph.microsoft.com/v1.0/directoryObjects/$userid"}
}
