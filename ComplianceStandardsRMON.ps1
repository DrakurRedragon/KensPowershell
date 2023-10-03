Read-Host "This will assist in the generation of results for Standards Audit."
Read-Host "207.1.2 Ensure Modern Authentication for Exchange Online is Enabled and 239.4.9 Ensure Bastic Authentication for Exchange online is disabled"
$org = Read-Host "Please enter the organization: "
Connect-ExchangeOnline -Organization $org
get-organizationconfig | format-table -auto name, OAuth*
Get-User -ResultSize Unlimited | Select-Object UserPrincipalName, AuthenticationPolicy
Read-Host "216.2.5 Ensure Office 365 SharePoint Infected files are disallowed for download."
$adminsite = Read-Host "Please enter the prefix for the Sharepoint admin site: "
Connect-SPOService -url https://$adminsite-admin.sharepoint.com
Get-SPOTenant | Select-Object DisallowInfectedFileDownload
Read-Host "230.3.8 Ensure that Microsoft Teams is using OneDrive for Business and SharePoint for meeting recordings"
Connect-MicrosoftTeams
Get-CsTeamsMeetingPolicy -identity global | fl RecordingStorageMode
Read-Host "234.4.4 Ensure Automatic Forwarding Options are disabled"
Connect-ExchangeOnline -Organization $org
Get-RemoteDomain Default | fl AllowedOOFType, AutoForwardEnabled
