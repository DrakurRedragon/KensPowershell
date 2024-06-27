connect-mggraph -scopes policy.read.all,directory.read.all,RoleManagement.Read.All

$var1 = Get-MgIdentityConditionalAccessPolicy

foreach ($1 in $var1) 
{
	write-host "Name: " $1.DisplayName
	write-host "State: " $1.state
	Write-Host "--------"
	foreach ($appid in $1.Conditions.Applications.IncludeApplications)
	{	
		try
		{
			$app1 = get-mgserviceprincipalbyappid -appid $($1.Conditions.Applications.IncludeApplications) -ea stop
			write-host "Included Applications: " $app1.DisplayName
		}
		catch
		{
			write-host "Included Applications: " $1.Conditions.Applications.IncludeApplications
		}
	}
	write-host "Apps:" $1.Conditions.ClientAppTypes
	Write-Host "--------"
	write-host "Include Locations: " $1.Conditions.Locations.includeLocations
	write-host "Exclude Locations: " $1.Conditions.Locations.excludeLocations
	write-host "Exclude Groups: " $1.Conditions.Users.ExcludeGroups
	$exrole=@()
	foreach ($exroleid in $1.Conditions.Users.ExcludeRoles)
	{
		try
		{
			$exrole = $exrole += get-mgdirectoryroletemplate -DirectoryRoleTemplateId $exroleid -ea stop | Select -ExpandProperty DisplayName
		}
		catch
		{
			$exrole = $exrole += $exroleid
		}
	}
	write-host "Excluded Roles: " $exrole
	$exuser=@()
	foreach ($exuserid in $1.Conditions.Users.ExcludeUsers)
	{
		try
		{
			$exuser = $exuser += get-mguser -userid $exuserid -ea stop | Select -ExpandProperty DisplayName
		}
		catch
		{
			$exuser = $exuser += $exuserid
		}
	}
	write-host "Excluded Users: " $exuser
	write-host "Include Groups: " $1.Conditions.Users.ExcludeGroups
	$inrole=@()
	foreach ($inroleid in $1.Conditions.Users.IncludeRoles)
	{
		try
		{
			$inrole = $inrole += get-mgdirectoryroletemplate -DirectoryRoleTemplateId $inroleid -ea stop | Select -ExpandProperty DisplayName
		}
		catch
		{
			$inrole = $inrole += $inroleid
		}
	}
	Write-Host "Include Roles: " $inrole
	$inuser=@()
	foreach ($inuserid in $1.Conditions.Users.IncludeUsers)
	{
		try
		{
			$inuser = $inuser += get-mguser -userid $inuserid -ea stop | Select -ExpandProperty DisplayName
		}
		catch
		{
			$inuser = $inuser += $inuserid
		}
	}
	Write-Host "Include Users: " $inuser
Write-Host "--------"
Write-host "Action (Built In): " $1.GrantControls.BuiltInControls
Write-host "Authentication Strength: " $1.GrantControls.AuthenticationStrength.DisplayName
write-host " "
write-host " "
}


get-mgapplicationbyappid -appid 41b23e61-6c1e-4545-b367-cd054e0ed4b4
