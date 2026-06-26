# This pulls every single GPO in a domain, creating an HTML file for each one.
$gpos = get-gpo -all
foreach ($gpo in $gpos)
{
if ($gpo.GpoStatus -like "*Disable*")
{$filename = "_Disabled " + $gpo.DisplayName}
else
{$filename = $gpo.DisplayName}
get-gporeport -guid $gpo.id -ReportType html -path c:\temp\GPO\$filename.html
}
