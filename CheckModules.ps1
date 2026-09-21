$module = get-InstalledModule
$name = "ImportExcel"
if ($module.name -like $name)
{
Write-Output "$name is Installed"
}
else
{
Write-Output "$name was not Installed.  Installing"
Install-Module $name -scope CurrentUser
Write-Output "$name is now Installed."
}
$module = get-Module
if ($module.name -like $name)
{
Write-Output "$name is Imported"
}
else
{
Write-Output "$name was not imported.  Importing"
Import-Module $name
Write-Output "$name is now imported."
}