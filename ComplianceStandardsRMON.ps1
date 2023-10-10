Write-Host "This will assist in the generation of results for Standards Audit."
Start-Sleep -Seconds 1.5
Write-Host "207.1.2 Ensure Modern Authentication for Exchange Online is Enabled and 239.4.9 Ensure Basic Authentication for Exchange online is disabled"
Start-Sleep -Seconds 1.5
$org = Read-Host "Please enter the organization: "
Connect-ExchangeOnline -Organization $org
get-organizationconfig | format-table -auto name, OAuth*
Get-User -ResultSize Unlimited | Select-Object UserPrincipalName, AuthenticationPolicy
Read-Host -Prompt "Press any key to continue"
Write-Host "4.17 SMTP AUTH SHALL be disabled."
Get-TransportConfig | Format-List SmtpClientAuthenticationDisabled
Read-Host -Prompt "Press any key to continue"
Write-Host "216.2.5 Ensure Office 365 SharePoint Infected files are disallowed for download."
Start-Sleep -Seconds 1.5
$adminsite = Read-Host "Please enter the prefix for the Sharepoint admin site: "
Connect-SPOService -url https://$adminsite-admin.sharepoint.com
Get-SPOTenant | Select-Object DisallowInfectedFileDownload
Read-Host -Prompt "Press any key to continue"
Write-Host "230.3.8 Ensure that Microsoft Teams is using OneDrive for Business and SharePoint for meeting recordings"
Start-Sleep -Seconds 1.5
Connect-MicrosoftTeams
Get-CsTeamsMeetingPolicy -identity global | fl RecordingStorageMode
Read-Host -Prompt "Press any key to continue"
Write-Host "234.4.4 Ensure Automatic Forwarding Options are disabled"
Start-Sleep -Seconds 1.5
Connect-ExchangeOnline -Organization $org
Get-RemoteDomain Default | fl AllowedOOFType, AutoForwardEnabled
Read-Host -Prompt "Press any key to continue"
Write-Host "244.4.15 Ensure Mailtips are enabled for end users"
Start-Sleep -Seconds 1.5
Get-OrganizationConfig |Select-Object MailTipsAllTipsEnabled, MailTipsExternalRecipientsTipsEnabled, MailTipsGroupMetricsEnabled, MailTipsLargeAudienceThreshold
Read-Host -Prompt "Press any key to continue"
Write-Host "250. 6.4 Ensure external storage providers available in Outlook on the Web are restricted"
Start-Sleep -Seconds 1.5
Get-OwaMailboxPolicy | Format-Table Name, AdditionalStorageProvidersAvailable
Read-Host -Prompt "Press any key to continue"
