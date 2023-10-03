Read-Host "This will assist in the generation of results for Standards Audit."
Read-Host "207.1.2 Ensure Modern Authentication for Exchange Online is Enabled"
$org = Read-Host "Please enter the organization: "
$cred = Get-Credential
Connect-ExchangeOnline -Organization $org -Credential $cred
get-organizationconfig | format-table -auto name, OAuth*
Read-Host "216.2.5 Ensure Office 365 SharePoint Infected files are disallowed for download."
$adminsite = Read-Host "Please enter the prefix for the Sharepoint admin site: "
Connect-SPOService -url https://$adminsite-admin.sharepoint.com
Get-SPOTenant | Select-Object DisallowInfectedFileDownload
