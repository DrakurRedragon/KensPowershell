# This is the start of the Script.  It will specify al of the connections and establish the variables.
# Make sure you have the following modules installed to be able to run this script.
# Install-Module -Name Microsoft.Online.SharePoint.PowerShell -scope CurrentUser
# install-module msonline -scope CurrentUser
# install-module ExchangeOnlineManagement -scope CurrentUser
# install-module microsoftteams -scope CurrentUser
# install-module azuread -scope currentuser

#


Write-Output "You will prompted to log in several times.  Use a Global Administrator account for the login."
#Read-Host "The script will pause to make sure you have access to the Global Adminstrator Account.  Press Enter when ready."
Write-Output "Connecting to Azure AD...This may take a few moments."
Connect-AzureAD
Write-Output "Connected"
Write-Output "Connecting to Microsoft Teams...This may take a few moments."
Connect-MicrosoftTeams
Write-Output "Connected"
Write-Output "Connecting to Exchange Online...This may take a few moments."
Connect-ExchangeOnline -ShowBanner:$false
Write-Output "Connected"
Write-Output "Connecting to MSOL Service...This may take a few moments."
Write-Output "Connected"
Connect-MsolService
Write-Output "Connected"
Write-Output "Connecting to Information Protection...This may take a few moments."
Connect-IPPSSession
Write-Output "Connected"
$adminsite = Read-Host "Please enter the prefix for the Sharepoint admin site: "
Write-Output "Connecting to Sharepoint Online Admin...This may take a few moments."
Connect-SPOService -url https://$adminsite-admin.sharepoint.com
Write-Output "Connected"
Write-Output "We will need to gather information to move forward; these will assist in various places within the script."
$Dom = Read-Host "Enter the FQDN: "

# Create a new instance/object of MS Word
$MSWord = New-Object -ComObject Word.Application

# Make MS Word visible
$MSWord.Visible = $True

# Add a new document
$mydoc = $MSWord.Documents.Add()

# Create a reference to the current document so we can begin adding text
$myText = $MSWord.Selection


# =============================================
# It starts with the 1.x series of questions for Azure Active Directory
# =============================================
# Questions not in this script: 1.1.7-15
# =============================================

Write-Output "Microsoft 365 Verification Script for Standards Alignment"

$myText.Font.Bold = 1
$myText.Style = 'Title'
$myText.TypeText("Microsoft 365 Verification Script for Standards Alignment")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$myText.Font.Bold = 1
$myText.Style = 'Heading 1'
$myText.TypeText("1.x Series")
$myText.TypeParagraph()
$myText.Font.Bold = 0

Write-Output "Verifying 1.1.1.1, 1.1.2.1 for MFA, 1.1.1.2 for Persistent Browser Session and 1.1.6 for legacy authentication."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("Verifying 1.1.1.1, 1.1.2.1 for MFA, 1.1.1.2 for Persistent Browser Session and 1.1.6 for legacy authentication")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-AzureADMSConditionalAccessPolicy

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify that the policies exist: MFA for Administrators, MFA for All Users, and Block Legacy Authentication.  They should be in an Enabled State.  You will need to verify the persistent browser session manually from the portal the first time.")
$myText.TypeParagraph()

Write-Output "Verifying Global Administrators, 1.1.3.1 and Cloud Only admin accounts, 1.1.3.2"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("Verifying Global Administrators, 1.1.3.1 and Cloud Only admin accounts, 1.1.3.2")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$role = Get-AzureADDirectoryRole | Where-Object {$_.displayName -eq 'Global Administrator'}
$var1 = Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId | Select-Object -property UserPrincipalName,DirSyncEnabled
foreach($usern in $var1) { 
$myText.Style = 'Normal'
$myText.typeText("Account: $($usern.UserPrincipalName). ")
$myText.TypeParagraph()
if ($usern.DirSyncEnabled -eq $null)
{$dirsync = "Yes"}
else
{$dirsync = "No"}
$myText.typeText("Cloud Only: $($dirsync).")
$myText.TypeParagraph()
}

Write-Output "Verifying SSPR is enabled, 1.1.4"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("Verifying SSPR is enabled, 1.1.4")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-AzureADMSAuthorizationPolicy | Select-Object -Property AllowedToUseSSPR

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("If it says true, then SSPR is enabled.")
$myText.TypeParagraph()


$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.7 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.8 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.9 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.10 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.11 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.12 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.13 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.14 is not supported.")
$myText.TypeParagraph()
$myText.Style = 'Heading 2'
$myText.TypeText("1.1.15 is not supported.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

# =============================================
# This is a continuation of the 1.x questions, Modern Authentication.
# =============================================
# This section is complete.
# =============================================

Write-Output "1.2 Ensure Modern Authentication for Exchange Online is Enabled and 4.9 Ensure Basic Authentication for Exchange online is disabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("1.2 Ensure Modern Authentication for Exchange Online is Enabled and 4.9 Ensure Basic Authentication for Exchange online is disabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-organizationconfig | select-object -property name,OAuth*

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify the Policy is set to true")
$myText.TypeParagraph()

$var1 = Get-User -ResultSize Unlimited | Select-Object UserPrincipalName, AuthenticationPolicy | Where-Object -Property AuthenticationPolicy -ne $null
$myText.Style = 'Normal'
$myText.TypeText("List of users with different policies:")
$myText.TypeParagraph()
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify no user has a different authentication policy set.  If it doesn't return anything, then no user has that property set.")
$myText.TypeParagraph()

Write-Output "1.3 Ensure Modern Authentication for SharePoint Applications is required."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("1.3 Ensure Modern Authentication for SharePoint Applications is required.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object LegacyAuthProtocolsEnabled

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("This should be False, otherwise legacy protocols are allowed")
$myText.TypeParagraph()

Write-Output "1.4 Microsoft 365 Passwords are set not to expire."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("1.4 Microsoft 365 Passwords are set not to expire.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-msolpasswordpolicy -domain $Dom | Select-Object ValidityPeriod

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("2147483647 is the desired value.  This means there is no expiration")
$myText.TypeParagraph()

Write-Output "1.5 Password Minimum Requirements, 12+ characters (This will only work on a local domain controller.)"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("1.5 Password Minimum Requirements, 12+ characters (This will only work on a local domain controller.  This not scriptable and will need to be verified, as DCs should not have Office products installed.)")
$myText.TypeParagraph()
$myText.Font.Bold = 0


# =============================================
# This is the 2.x series of questions, Application Permissions
# =============================================

$myText.Font.Bold = 1
$myText.Style = 'Heading 1'
$myText.TypeText("2.x Series")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.1 is not supported.")
$myText.TypeParagraph()

Write-Output "2.2 Ensure Calendar Details sharing with External Users is disabled."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.2 Ensure Calendar Details sharing with External Users is disabled.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-SharingPolicy | Where-object -property Default -eq True | Select-Object -Property Domains

$myText.Style = 'Normal'
$myText.typeText("$($var1.Domains -join ', ')")
$myText.TypeParagraph()
$myText.typeText("Look for Anonymous: in the Domains list.  If it is present and has any properties, external sharing is not disabled.")
$myText.TypeParagraph()

Write-Output "2.3 Ensure Defender for Office 365 SafeLinks for Office Applications is enabled."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.3 Ensure Defender for Office 365 SafeLinks for Office Applications is enabled.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-safelinkspolicy | select-object -property ID,EnableSafeLinksforEmail,EnableSafeLinksforTeams,EnableSafeLinksforOffice

foreach($safelink in $var1) { 
$myText.Style = 'Normal'
$myText.typeText("Policy ID: $($safelink.Id). ")
$myText.TypeParagraph()
$myText.typeText("Safelinks for Office: $($safelink.EnableSafelinksforOffice).")
$myText.TypeParagraph()
} 
$myText.typeText("Find their primary policy, usually the one with the client's name, and verify the setting is true.")
$myText.TypeParagraph()

Write-Output "2.4 Ensure Defender for Office 365 for SharePoint, OneDrive, and Microsoft Teams is Enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.4 Ensure Defender for Office 365 for SharePoint, OneDrive, and Microsoft Teams is Enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-atppolicyforo365 | select-object -property ID,EnableATPForSPOTeamsODB,EnableSafeDocs

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify that both settings are True")
$myText.TypeParagraph()

Write-Output "2.5 Ensure Office 365 SharePoint Infected files are disallowed for download."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.5 Ensure Office 365 SharePoint Infected files are disallowed for download.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object DisallowInfectedFileDownload

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify the setting is True")
$myText.TypeParagraph()

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.6 Ensure user consent to apps accessing company data on their behalf is not allowed is not supported.")
$myText.TypeParagraph()

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.7 Ensure the admin consent workflow is enabled")
$myText.TypeParagraph()

Write-Output "2.8 Ensure users installing Outlook add-ins is not allowed."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.8 Ensure users installing Outlook add-ins is not allowed.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-RoleAssignmentPolicy | Where-Object -property Name -eq "Default Role Assignment Policy" | Select-Object -ExpandProperty AssignedRoles

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify that My Custom Apps, My Marketplace Apps, and My ReadWriteMailbox Apps are NOT present in the list.")
$myText.TypeParagraph()

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.9 Ensure users installing Word, Excel, Powerpoint is not allowed.")
$myText.TypeParagraph()

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.10 Ensure Internal Phishing for Forms is enabled.")
$myText.TypeParagraph() 

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("2.11 Ensure that Sways cannot be shared with people outside the organization")
$myText.TypeParagraph()

# =============================================
# This is the 3.x series of questions, Data Management
# =============================================
# This section is complete.
# =============================================

$myText.Font.Bold = 1
$myText.Style = 'Heading 1'
$myText.TypeText("3.x Series")
$myText.TypeParagraph()
$myText.Font.Bold = 0

Write-Output "3.1 Ensure the customer lockbox feature is enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.1 Ensure the customer lockbox feature is enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-organizationconfig | select-object -property CustomerLockboxEnabled

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify this setting is True")
$myText.TypeParagraph()

Write-Output "3.2 Ensure SharePoint Online data classification policies are set up and used"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.2 Ensure SharePoint Online data classification policies are set up and used")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property EnableAIPIntegration

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify this setting is true")
$myText.TypeParagraph()

Write-Output "3.3 Ensure external users are not allowed in Skype or Teams"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.3 Ensure external users are not allowed in Skype or Teams")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsClientConfiguration | Select-Object -Property AllowGuestUser

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify this setting is false")
$myText.TypeParagraph()

Write-Output "3.4 Ensure DLP policies are enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.4 Ensure DLP policies are enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-DlpCompliancePolicy | Select-Object DisplayName,Enabled

foreach($dlppol in $var1) 
{
	$myText.Style = 'Normal'
	$myText.typeText("Display Name: $($dlppol.DisplayName). ")
	$myText.TypeParagraph()
	$myText.typeText("Enabled: $($dlppol.Enabled).")
	$myText.TypeParagraph()
} 
$myText.typeText("Verify that policies exist and are enabled.")
$myText.TypeParagraph()

Write-Output "3.5 Ensure DLP policies are enabled for Microsoft Teams"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.5 Ensure DLP policies are enabled for Microsoft Teams")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-DlpCompliancePolicy | Select-Object DisplayName,Enabled,Workload

foreach($dlppol in $var1) 
{
	$myText.Style = 'Normal'
	$myText.typeText("Display Name: $($dlppol.DisplayName). ")
	$myText.TypeParagraph()
	$myText.typeText("Enabled: $($dlppol.Enabled).  Enabled For: $($dlppol.Workload)")
	$myText.TypeParagraph()
} 
$myText.typeText("Verify that policies are enabled for Teams in the Enabled For list.")
$myText.TypeParagraph()

Write-Output "3.6 Ensure that external users cannot share files, folders, and sites they do not own"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.6 Ensure that external users cannot share files, folders, and sites they do not own")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property PreventExternalUsersFromResharing

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify this is true")
$myText.TypeParagraph()

Write-Output "3.7 Ensure external file sharing in Teams is enabled for only approved cloud storage services"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.7 Ensure external file sharing in Teams is enabled for only approved cloud storage services")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsClientConfiguration | Select-Object -Property AllowDropbox,AllowBox,AllowGoogleDrive,AllowShareFile,AllowEgnyte

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify that only the methods approved by company policy are listed as true")
$myText.TypeParagraph()

Write-Output "3.8 Ensure that Microsoft Teams is using OneDrive for Business and SharePoint for meeting recordings"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("3.8 Ensure that Microsoft Teams is using OneDrive for Business and SharePoint for meeting recordings")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsMeetingPolicy -identity global | Select-Object -Property RecordingStorageMode

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify this is not Stream")
$myText.TypeParagraph()

# =============================================
# This is the 4.x series of questions, Email Security
# =============================================
# This section is complete.
# =============================================

$myText.Font.Bold = 1
$myText.Style = 'Heading 1'
$myText.TypeText("4.x Series")
$myText.TypeParagraph()
$myText.Font.Bold = 0

Write-Output "4.1 Ensure the Common Attachment Types Filter is enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.1 Ensure the Common Attachment Types Filter is enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-MalwareFilterPolicy | Select-Object -Property ID,FileTypeAction,FileTypes

foreach($malfil in $var1) 
{
	$myText.Style = 'Heading 3' 
	$myText.typeText("Display Name: $($malfil.Id). ")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.typeText("Action: $($malfil.FileTypeAction)")
	$myText.TypeParagraph()
	$myText.typeText("File Types: $($malfil.FileTypes)")
	$myText.TypeParagraph()
} 
$myText.typeText("Verify that policies are enabled and that the common attachment types are listed.")
$myText.TypeParagraph()

Write-Output "4.2 Ensure Exchange Online Spam Policies are set correctly"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.2 Ensure Exchange Online Spam Policies are set correctly")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-hostedoutboundspamfilterpolicy | Select-Object -Property ID,BccSuspiciousOutboundMail,BccSuspiciousOutboundAdditionalRecipients,NotifyOutboundSpam,NotifyOutboundSpamRecipients

$myText.Style = 'Normal'
$myText.typeText("Policy ID: $($var1.id)")
$myText.TypeParagraph()
$myText.typeText("Notify: $($var1.NotifyOutboundSpam) OR BCC: $($var1.BccSuspiciousOutboundMail)")
$myText.TypeParagraph()
$myText.typeText("Email addresses being Notified: $($var1.NotifyOutboundSpamRecipients -join ", ")")
$myText.TypeParagraph()
$myText.typeText("Email addresses being BCCed: $($var1.BccSuspiciousOutboundAdditionalRecipients -join ", ")")
$myText.TypeParagraph()
$myText.typeText("Verify that preferably Notify is set to True and an address is listed in Notified.  If BCC is set to true and an address is available, this would be accepable, if only marginally.")
$myText.TypeParagraph()

Write-Output "4.2.1 IP Allow Lists are NOT present for Spam"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.2.1 IP Allow Lists are NOT present for Spam")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-HostedConnectionFilterPolicy | Select-Object Name,IPAllowList

foreach ($ipallow in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.typeText("$($ipallow.Name)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.typeText("IP List: $($ipallow.IPAllowList -join ", ")")
	$myText.TypeParagraph()
}
$myText.typeText("Verify no IP Allow Lists exists.  The IP List should be blank for all rules")
$myText.TypeParagraph()

Write-Output "4.3 Ensure mail transport rules do not forward email to external domains"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.3 Ensure mail transport rules do not forward email to external domains")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-TransportRule | Where-Object { $_.BlindCopyTo -ne $null -or $_.CopyTo -ne $null -or $_.RedirectMessageTo -ne $null -or $_.AddToRecipients -ne $null }  | Select-Object -Property Name,isValid,State,Mode,BlindCopyTo,CopyTo,RedirectMessageTo,AddToRecipients
if ($var1 -eq $null)
{
	$myText.typeText("No auto forward Transport Rules exist")
	$myText.TypeParagraph()
	$myText.typeText("This is considered in compliance.")
	$myText.TypeParagraph()
}
else
{
	foreach ($forrule in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($forrule.Name)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("isValid: $($forrule.isValid)")
		$myText.TypeParagraph()
		$myText.typeText("State: $($forrule.State)")
		$myText.TypeParagraph()
		$myText.typeText("Mode: $($forrule.Mode)")
		$myText.TypeParagraph()
		$myText.typeText("BlindCopyTo: $($forrule.BlindCopyTo)")
		$myText.TypeParagraph()
		$myText.typeText("CopyTo: $($forrule.CopyTo)")
		$myText.TypeParagraph()
		$myText.typeText("RedirectMessageTo: $($forrule.RedirectMessageTo)")
		$myText.TypeParagraph()
		$myText.typeText("AddToRecipients: $($forrule.AddToRecipients)")
		$myText.TypeParagraph()
	}
	$myText.TypeParagraph()
	$myText.typeText("Verify that these rules do not send outside the organization.")
	$myText.TypeParagraph()
}

Write-Output "4.4 Ensure Automatic Forwarding Options are disabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.4 Ensure Automatic Forwarding Options are disabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-RemoteDomain Default | Select-Object -Property AllowedOOFType, AutoForwardEnabled

$myText.Style = 'Normal'
$myText.typeText("Is Auto Forward Enabled? $($var1.AutoForwardEnabled).  If enabled, where is it allowed? $($var1.AllowedOOFType)")
$myText.TypeParagraph()
$myText.typeText("Ideally, Auto Forward should be disabled (False).  If it is enabled (True), then it should not be allowed External.")
$myText.TypeParagraph()

Write-Output "4.5 Ensure mail transport rules do not whitelist specific domains"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.5 Ensure mail transport rules do not whitelist specific domains")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-TransportRule | Where-Object { $_.SenderDomainIs -ne $null -or $_.From -ne $null -or $_.FromAddressMatchesPatterns -ne $null -or $_.FromAddressContainsWords -ne $null }  | Select-Object -Property Name,isValid,State,Mode,From,SenderDomainIs,FromAddressMatchesPatterns,FromAddressContainsWords,Description
if ($var1 -eq $null)
{
	$myText.typeText("No domain or sender address based Transport Rules exist")
	$myText.TypeParagraph()
	$myText.typeText("This is considered in compliance.")
	$myText.TypeParagraph()
}
else
{
	foreach ($forrule in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($forrule.Name)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("isValid: $($forrule.isValid)")
		$myText.TypeParagraph()
		$myText.typeText("State: $($forrule.State)")
		$myText.TypeParagraph()
		$myText.typeText("Mode: $($forrule.Mode)")
		$myText.TypeParagraph()
		$myText.typeText("From: $($forrule.From)")
		$myText.TypeParagraph()
		$myText.typeText("SenderDomainIs: $($forrule.SenderDomainIs)")
		$myText.TypeParagraph()
		$myText.typeText("FromAddressMatchesPatterns: $($forrule.FromAddressMatchesPatterns)")
		$myText.TypeParagraph()
		$myText.typeText("FromAddressContainsWords: $($forrule.FromAddressContainsWords)")
		$myText.TypeParagraph()
		$myText.typeText("Description: $($forrule.Description)")
		$myText.TypeParagraph()		
	}
	$myText.TypeParagraph()
	$myText.typeText("Verify that these rules do not whitelist or blanket allow entire domains.")
	$myText.TypeParagraph()
}

Write-Output "4.6 Ensure the Client Rules Forwarding Block is enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.6 Ensure the Client Rules Forwarding Block is enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-TransportRule | Where-Object { $_.SentToScope -eq "NotInOrganizaiton" -and $_.FromScope -eq "InOrganizaation" -and $_.MessageTypeMatches -eq "AutoForward"} | Select-Object -Property Name,isValid,State,Mode,SentToScope,FromScope,MessageTypeMatches,Description

if ($var1 -eq $null)
{
	$myText.typeText("The Client Rules Autoforward Transport Rule does not exist")
	$myText.TypeParagraph()
	$myText.typeText("This is considered out of compliance.")
	$myText.TypeParagraph()
}
else
{
	foreach ($forrule in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($forrule.Name)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("isValid: $($forrule.isValid)")
		$myText.TypeParagraph()
		$myText.typeText("State: $($forrule.State)")
		$myText.TypeParagraph()
		$myText.typeText("Mode: $($forrule.Mode)")
		$myText.TypeParagraph()
		$myText.typeText("Description: $($forrule.SentToScope)")
		$myText.TypeParagraph()
		$myText.typeText("State: $($forrule.FromScope)")
		$myText.TypeParagraph()
		$myText.typeText("Mode: $($forrule.MessageTypeMatches)")
		$myText.TypeParagraph()
		$myText.typeText("Description: $($forrule.Description)")
		$myText.TypeParagraph()			
	}
	$myText.TypeParagraph()
	$myText.typeText("Verify that the rule exists and performs some kind of block action.")
	$myText.TypeParagraph()
}

Write-Output "4.7 Ensure the Advanced Threat Protection Safe Links policy is enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.7 Ensure the Advanced Threat Protection Safe Links policy is enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-safelinkspolicy | select-object -property ID,AllowClickThrough,EnableForInternalSenders

foreach ($47sl in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($47sl.id)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("Enabled for Internal Senders: $($47sl.EnableForInternalSenders)")
		$myText.TypeParagraph()
		$myText.typeText("Allow Click Through: $($47sl.AllowclickThrough)")
		$myText.TypeParagraph()
	}

$myText.TypeParagraph()
$myText.typeText("Verify that their primary rule is enabled for internal senders and DOES NOT allow click through.")
$myText.TypeParagraph()

Write-Output "4.8 Ensure the Advanced Threat Protection Safe Attachments policy is enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.8 Ensure the Advanced Threat Protection Safe Attachments policy is enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SafeAttachmentRule

foreach ($48atpsa in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($48atpsa.Name)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("Enabled: $($48atpsa.State)")
		$myText.TypeParagraph()
		$myText.typeText("Safe Attachment Policy: $($48atpsa.SafeAttachmentPolicy)")
		$myText.TypeParagraph()
	}

$myText.TypeParagraph()
$myText.typeText("Verify that rules exist and at least one is enabled with an assigned policy.")
$myText.TypeParagraph()

Write-Output "4.9 Ensure Basic Authentication for Exchange online is disabled" >> $filename

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.9 Ensure Basic Authentication for Exchange online is disabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = get-organizationconfig | select-object -property name,OAuth*

$myText.Style = 'Normal'
$myText.typeText("$($var1)")
$myText.TypeParagraph()
$myText.typeText("Verify the Policy is set to true")
$myText.TypeParagraph()

$var1 = Get-User -ResultSize Unlimited | Select-Object UserPrincipalName, AuthenticationPolicy | Where-Object -Property AuthenticationPolicy -ne $null

if ($var1 -eq $null)
{
	$myText.typeText("No mailboxes found with separate authentication policies.")
	$myText.TypeParagraph()
	$myText.typeText("If the policy above is true, this is in compliance.  Otherwise, the policy needs to be changed to true and no other action needs to be taken.")
	$myText.TypeParagraph()
}
else
{
	foreach ($49 in $var1)
		{
			$myText.Style = 'Heading 3'
			$myText.typeText("Name: $($49.UserPrincipalName)")
			$myText.TypeParagraph()
			$myText.Style = 'Normal'
			$myText.typeText("Enabled: $($49.AuthenticationPolicy)")
			$myText.TypeParagraph()
		}
	$myText.typeText("The above mailboxes have separate authentication policies.  They will need to be verified that they do not specifically allow legacy authentication to be considered in compliance.")
	$myText.TypeParagraph()
}


Write-Output "4.10 Ensure that an anti-phishing policy has been created"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.10 Ensure that an anti-phishing policy has been created")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-AntiPhishPolicy | Select-Object -property ID,Enabled

foreach ($410 in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($410.Id)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("Enabled: $($410.Enabled)")
		$myText.TypeParagraph()
	}

$myText.TypeParagraph()
$myText.typeText("Verify that rules exist and at least one is enabled ..")
$myText.TypeParagraph()

Write-Output "4.11 Ensure that DKIM is enabled for all Exchange Online Domains"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.11 Ensure that DKIM is enabled for all Exchange Online Domains")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-DkimSigningConfig

foreach ($411 in $var1)
{
	$myText.Style = 'Normal'
	$myText.typeText("$($411.domain) - $($411.Enabled)")
	$myText.TypeParagraph()
}

$myText.TypeParagraph()
$myText.typeText("Verify that each is set to true for Exchange Online email sending domains.")
$myText.TypeParagraph()

Write-Output "4.12 Ensure SPF Records are published for Exchange domains."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.12 Ensure SPF Records are published for Exchange domains.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$myText.Style = 'Normal'
$myText.typeText("$(Resolve-DnsName -Name $Dom -Type TXT -Server 1.1.1.1 -ErrorAction SilentlyContinue | Where {$_.Strings -Match "v=spf"} | Select -ExpandProperty Strings)")
$myText.TypeParagraph()

Write-Output "Verify there is an entry for SPF."


$addDomyn = Read-Host -Prompt "Are there additional domains?  Y or N?"
if ($addDomyn -eq 'y')
{
    do {
        $addDom = Read-Host -Prompt "Enter additional mail domains to search."
		$myText.Style = 'Normal'
		$myText.typeText("$($addDom)")
		$myText.TypeParagraph()
		$myText.typeText("$(Resolve-DnsName -Name $addDom -Type TXT -Server 1.1.1.1 -ErrorAction SilentlyContinue | Where {$_.Strings -Match "v=spf"} | Select -ExpandProperty Strings)")
		$myText.TypeParagraph()
        $quit = Read-Host -Prompt "Do you have any other domains to search?  Y or N?"
        } while ($quit -eq 'y')
}

$myText.TypeParagraph()
$myText.typeText("If the string is empty, then the record does not exist.  This is a failure.")
$myText.TypeParagraph()

Write-Output "4.13 Ensure DMARC Records for all Exchange Online domains are published"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.13 Ensure DMARC Records for all Exchange Online domains are published")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$myText.Style = 'Normal'
$myText.typeText("$($Dom)")
$myText.TypeParagraph()
$myText.typeText("$(Resolve-DnsName -Name _dmarc.$Dom -Type TXT -Server 1.1.1.1 -ErrorAction SilentlyContinue | Select Strings)")
$myText.TypeParagraph()

Write-Output "Verify there is an entry for v=DMARC1." >> $filename
$addDomyn = Read-Host -Prompt "Are there additional domains?  Y or N?"
if ($addDomyn -eq 'y')
{
    do {
        $addDom = Read-Host -Prompt "Enter additional mail domains to search."
		$myText.Style = 'Normal'
		$myText.typeText("$($addDom)")
		$myText.TypeParagraph()
		$myText.typeText("$(Resolve-DnsName -Name _dmarc.$addDom -Type TXT -Server 1.1.1.1 -ErrorAction SilentlyContinue | Select Strings)")
		$myText.TypeParagraph()
        $quit = Read-Host -Prompt "Do you have any other domains to search?  Y or N?"
        } while ($quit -eq 'y')
}

$myText.TypeParagraph()
$myText.typeText("If the string is empty, then the record does not exist.  This is a failure.")
$myText.TypeParagraph()

Write-Output "4.14 Ensure notifications for internal users sending malware is Enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.14 Ensure notifications for internal users sending malware is Enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-MalwareFilterPolicy | Select-Object -Property ID,EnableInternalSenderAdminNotifications,InternalSenderAdminAddress

foreach ($414 in $var1)
	{
		$myText.Style = 'Heading 3'
		$myText.typeText("Name: $($414.Id)")
		$myText.TypeParagraph()
		$myText.Style = 'Normal'
		$myText.typeText("Enabled: $($414.EnableInternalSenderAdminNotifications)")
		$myText.TypeParagraph()
		$myText.typeText("Notification Addresses: $($414.InternalSenderAdminAddress)")
		$myText.TypeParagraph()
	}

$myText.TypeParagraph()
$myText.typeText("Verify that there is a policy with the value enabled that is sending to a valid email address that we monitor.")
$myText.TypeParagraph()

Write-Output "4.15 Ensure Mailtips are enabled for end users"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.15 Ensure Mailtips are enabled for end users")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-OrganizationConfig | Select-Object MailTipsAllTipsEnabled, MailTipsExternalRecipientsTipsEnabled, MailTipsGroupMetricsEnabled, MailTipsLargeAudienceThreshold

$myText.Style = 'Normal'
$myText.typeText("Mail tips Enabled? $($var1.MailTipsAllTipsEnabled)")
$myText.TypeParagraph()
$myText.typeText("Mail tips Enabled for external recipients? $($var1.MailTipsExternalRecipientsTipsEnabled)")
$myText.TypeParagraph()
$myText.typeText("Mail tips Group Metrics enabled? $($var1.MailTipsGroupMetricsEnabled)")
$myText.TypeParagraph()
$myText.typeText("Mail tips Large audience Threshold? $($var1.MailTipsLargeAudienceThreshold)")
$myText.TypeParagraph()
$myText.TypeParagraph()
$myText.typeText("If the first is true, they are in compliance with this policy.  The second should also be set to true, but this isn't required.")
$myText.TypeParagraph()

Write-Output "4.16 Set spam policy to quarantine messages - not deliver to junk mail"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.16 Set spam policy to quarantine messages - not deliver to junk mail")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-HostedContentFilterRule

foreach ($416 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($416.Name)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Enabled: $($416.State)")
	$myText.TypeParagraph()
	$myText.TypeText("Policy: $($416.HostedContentFilterPolicy)")
	$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("Verify that a rule exists that is enabled with a Policy set.  Verify the policy contents below.")
$myText.TypeParagraph()

$var1 = Get-HostedContentFilterPolicy | Select-Object -property ID,SpamAction,PhishSpamAction,BulkSpamAction,HighConfidencePhishAction

foreach ($4162 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($4162.ID)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Spam Action: $($4162.SpamAction)")
	$myText.TypeParagraph()
	$myText.TypeText("Phish Spam Action: $($4162.PhishSpamAction)")
	$myText.TypeParagraph()
	$myText.TypeText("Bulk Spam Action: $($4162.BulkSpamAction)")
	$myText.TypeParagraph()
	$myText.TypeText("High Confidence Phish Action: $($4162.HighConfidencePhishAction)")
	$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("Find the matching policy for the enabled rule.  Make sure that none of the values are Deliver to Junk Mail.  They should be at least Quarantine if not Reject.")
$myText.TypeParagraph()

Write-Output "4.17 SMTP AUTH SHALL be disabled."

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.17 SMTP AUTH SHALL be disabled.")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-TransportConfig | Select-Object -Property SmtpClientAuthenticationDisabled

$myText.Style = 'Normal'
$myText.typeText("Smtp Client Authentication Disabled: $($var1.SmtpClientAuthenticationDisabled )")
$myText.TypeParagraph()
$myText.typeText("Verify this setting is true.")
$myText.TypeParagraph()

Write-Output "4.18 Email Encryption Policy is Configured"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.18 Email Encryption Policy is Configured")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-IRMConfiguration | Select-Object -Property AzureRMSLicensingEnabled

$myText.Style = 'Normal'
$myText.typeText("Is RMS Licensing enabled? $($var1.AzureRMSLicensingEnabled)")
$myText.TypeParagraph()
$myText.typeText("Verify this setting is true, then verify that the Transport Rules exist below.")
$myText.TypeParagraph()

$var1 = Get-TransportRule | Select-Object -Property Name,IsValid,State,Mode,ApplyRightsProtectionTemplate,ApplyRightsProtectionCustomizationTemplate | Where-Object {$_.ApplyRightsProtectionTemplate -ne $null -or $_.ApplyRightsProtectionCustomizationTemplate -ne $null}

foreach ($418 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($418.Name)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Is it Valid, enabled, and enforced? $($418.isValid), $($418.State), $($418.Mode)")
	$myText.TypeParagraph()
	$myText.TypeText("Rights Protection Template: $($418.ApplyRightsProtectionTemplate)")
	$myText.TypeParagraph()
	$myText.TypeText("Customization Template: $($418.ApplyRightsProtectionCustomizationTemplate)")
	$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("Verify that the policies have one of the two items set for encryption.  If they are set to encrypt or the customization template is set to encrypt, then this is in compliance.")
$myText.TypeParagraph()

Write-Output "4.19 Enhanced Filtering is turned on for 3rd Party email filtering tools"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("4.19 Enhanced Filtering is turned on for 3rd Party email filtering tools")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-InboundConnector | Select-Object -Property Name,EFSkipLastIP,EFSkipIPs | Where-Object { $_.Name -ne "Inbound On-Premises Connector" }

foreach ($419 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($419.Name)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Enhanced Filtering Skip Last IP setting: $($419.EFSkipLastIP)")
	$myText.TypeParagraph()
	$myText.TypeText("Enhanced filtering skip IP List: $($419.EFSkipIPs -join ", ")")
	$myText.TypeParagraph()
}

$myText.TypeParagraph()
$myText.typeText("For each of the inbound connectors, either the setting should be true or IPs should exist in the list.  This will make sure that all of the protections view any mail coming in using those connectors as external mail and scan accordingly.  If the information above is blank, no connector other than the standard On-Premises connector exists, and this is considered in compliance.")
$myText.TypeParagraph()

# =============================================
# This is the 5.x series of questions, Auditing
# =============================================
# The majority of the requirements are covered in the Weekly Audit Script that pulls all of the information from the audit logs.
# =============================================

$myText.Font.Bold = 1
$myText.Style = 'Heading 1'
$myText.TypeText("5.x Series")
$myText.TypeParagraph()
$myText.Font.Bold = 0

Write-Output "5.1 Ensure Microsoft 365 audit log search is Enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("5.1 Ensure Microsoft 365 audit log search is Enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-AdminAuditLogConfig | Select-Object -Property AdminAuditLogEnabled

$myText.Style = 'Normal'
$myText.typeText("Is this setting turned on? $($var1.AdminAuditLogEnabled)")
$myText.TypeParagraph()
$myText.typeText("Verify that the setting is set to True.")
$myText.TypeParagraph()

Write-Output "5.2 Ensure mailbox auditing for all users is Enabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("5.2 Ensure mailbox auditing for all users is Enabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.AuditEnabled -ne $true -and ($_.RecipientTypeDetails -ne "UserMailbox" -or $_.RecipientTypeDetails -ne "SharedMailbox")}

foreach ($52 in $var1)
{
$myText.Style = 'Normal'
$myText.typeText("$($52.Name) - $($52.Alias)")
$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("The mailboxes above are not audited.")
$myText.TypeParagraph()


# =============================================
# This is the 6.x Series of Questions, Storage 
# =============================================
# This section is complete.                    
# =============================================

Write-Output "6.1 Ensure document sharing is being controlled by domains with allow list or deny list"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.1 Ensure document sharing is being controlled by domains with allow list or deny list")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property SharingAllowedDomainList, SharingBlockedDomainList, SharingdomainRestrictionMode

$myText.Style = 'Normal'
$myText.typeText("Restriction Mode: $($var1.SharingdomainRestrictionMode)")
$myText.TypeParagraph()
$myText.typeText("Allowed Domain List: $($var1.SharingAllowedDomainList -join ", ")")
$myText.TypeParagraph()
$myText.typeText("Blocked Domain List: $($var1.SharingBlockedDomainList -join ", ")")
$myText.TypeParagraph()
$myText.TypeParagraph()
$myText.typeText("Verify that the restriction mode is set to something, then verify that the list that corresponds to the mode set has entries.")
$myText.TypeParagraph()

Write-Output "6.2 Block OneDrive for Business sync from unmanaged devices"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.2 Block OneDrive for Business sync from unmanaged devices")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenantSyncClientRestriction | Select-Object -Property TenantRestrictionEnabled, AllowedDomainList

$myText.Style = 'Normal'
$myText.typeText("Restriction Enabled: $($var1.TenantRestrictionEnabled)")
$myText.TypeParagraph()
$myText.typeText("Allowed Domains: $($var1.AllowedDomainList)")
$myText.TypeParagraph()
$myText.TypeParagraph()
$myText.typeText("Verify that the restriction is enabled and that the allowed domains is limited to the managed domains.")
$myText.TypeParagraph()

Write-Output "6.3 Ensure expiration time for external sharing links is set"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.3 Ensure expiration time for external sharing links is set")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property ExternalUserExpirationRequired, ExternalUserExpireInDays

$myText.Style = 'Normal'
$myText.typeText("Expiration required: $($var1.ExternalUserExpirationRequired)")
$myText.TypeParagraph()
$myText.typeText("Days set by default: $($var1.ExternalUserExpireInDays)")
$myText.TypeParagraph()
$myText.TypeParagraph()
$myText.typeText("Verify that expiration is required and that the days are set to an actual amount.")
$myText.TypeParagraph()

Write-Output "6.4 Ensure external storage providers available in Outlook on the Web are restricted"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.4 Ensure external storage providers available in Outlook on the Web are restricted")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-OwaMailboxPolicy | Select-Object -Property Name,AdditionalStorageProvidersAvailable

foreach ($64 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($64.Name)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("$($64.AdditionalStorageProvidersAvailable)")
	$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("For the in place rules, verify that this value is not True, unless there is a business need to allow it.")
$myText.TypeParagraph()

Write-Output "6.5 File and Folder Links default Sharing Settings SHALL be set to 'Specific People'"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.5 File and Folder Links default Sharing Settings SHALL be set to 'Specific People'")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property DefaultSharingLinkType

$myText.Style = 'Normal'
$myText.typeText("Default Link type: $($var1.DefaultSharingLinkType)")
$myText.TypeParagraph()
$myText.typeText("Verify that the setting is NOT Anonymous Access.")
$myText.TypeParagraph()

Write-Output "6.6.1 Anyone Links SHOULD be turned Off"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.6.1 Anyone Links SHOULD be turned Off")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property SharingCapability

$myText.Style = 'Normal'
$myText.typeText("Sharing Capability: $($var1.SharingCapability)")
$myText.TypeParagraph()
$myText.typeText("If the Value is ExternalUserandGuestSharing, it allows Anyone links.")
$myText.TypeParagraph()

Write-Output "6.6.2 If Enabled, Anyone Links SHOULD be set to View by default"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.6.2 If Enabled, Anyone Links SHOULD be set to View by default")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOTenant | Select-Object -Property FileAnonymousLinkType, FolderAnonymousLinkType

$myText.Style = 'Normal'
$myText.typeText("File Default: $($var1.FileAnonymousLinkType)")
$myText.TypeParagraph()
$myText.typeText("Folder Default: $($var1.FolderAnonymousLinkType)")
$myText.TypeParagraph()
$myText.typeText("Both of theses settings shoould be View.")
$myText.TypeParagraph()

Write-Output "6.7 Users SHALL Be Prevented from Running Custom Scripts"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("6.7 Users SHALL Be Prevented from Running Custom Scripts")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-SPOSite -identity https://$adminsite.sharepoint.com | Select-Object -Property DenyAddAndCustomizePages

$myText.Style = 'Normal'
$myText.typeText("$($var1.DenyAddAndCustomizePages)")
$myText.TypeParagraph()
$myText.typeText("This setting should be Enabled.  That means it is denying the custom scripts from running.")
$myText.TypeParagraph()

# =============================================
# This is the 8.x series of questions for Teams
# =============================================
# This section is complete.                    
# =============================================

Write-Output "8.1 Private channels shall be utilized to restrict access to sensitive information"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.1 Private channels shall be utilized to restrict access to sensitive information")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-Team | Where-Object -Property Visibility -eq "Public"

foreach ($81 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($81.DisplayName)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("$($81.Visibility)")
	$myText.TypeParagraph()
	$myText.TypeText("$($81.Description)")
	$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("Above is a list of all public Teams channels.  Make sure that they are not sensitive.")
$myText.TypeParagraph()

Write-Output "8.2 External Participants SHOULD NOT be allowed to request control"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.2 External Participants SHOULD NOT be allowed to request control")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsMeetingPolicy | Where-Object -Property Identity -eq "Global" | Select-Object -Property AllowExternalParticipantGiveRequestControl

$myText.Style = 'Normal'
$myText.typeText("Can External Participants request control: $($var1.AllowExternalParticipantGiveRequestControl)")
$myText.TypeParagraph()
$myText.typeText("This should be set to false.")
$myText.TypeParagraph()

Write-Output "8.3 Anonymous Users SHALL NOT be enabled to start meetings"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.3 Anonymous Users SHALL NOT be enabled to start meetings")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsMeetingPolicy | Where-Object -Property Identity -eq "Global" | Select-Object -Property AllowAnonymousUsersToStartMeeting

$myText.Style = 'Normal'
$myText.typeText("Can External Participants start meetings: $($var1.AllowAnonymousUsersToStartMeeting)")
$myText.TypeParagraph()
$myText.typeText("This should be set to false.")
$myText.TypeParagraph()

Write-Output "8.4 Automatic Admittance to meetings SHOULD be restricted"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.4 Automatic Admittance to meetings SHOULD be restricted")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsMeetingPolicy | Where-Object -Property Identity -eq "Global" | Select-Object -Property AutoAdmittedUsers

$myText.Style = 'Normal'
$myText.typeText("Who will be automatically admitted: $($var1.AutoAdmittedUsers)")
$myText.TypeParagraph()
$myText.typeText("This should be set to EveryoneInSameAndFederatedCompany.")
$myText.TypeParagraph()

Write-Output "8.5 External User Access SHALL be restricted"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.5 External User Access SHALL be restricted")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CSTenantFederationConfiguration | Select-Object -Property AllowedDomains

$myText.Style = 'Normal'
$myText.typeText("List of allowed domains: $($var1.AllowedDomains -join ", ")")
$myText.TypeParagraph()
$myText.typeText("This should have domains listed in it, otherwise all external domains are authorized.")
$myText.TypeParagraph()

Write-Output "8.6.1 Unmanaged Users SHALL NOT be able to initiatie contact with internal users"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.6.1 Unmanaged Users SHALL NOT be able to initiatie contact with internal users")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CSTenantFederationConfiguration | Select-Object -Property AllowTeamsConsumerInbound

$myText.Style = 'Normal'
$myText.typeText("Personal Teams Users can message corporate users first: $($var1.AllowTeamsConsumerInbound)")
$myText.TypeParagraph()
$myText.typeText("This setting should be false.")
$myText.TypeParagraph()

Write-Output "8.6.2 Unmanaged Users SHOULD not be able to be contacted by internal users"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.6.2 Unmanaged Users SHOULD not be able to be contacted by internal users")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CSTenantFederationConfiguration | Select-Object -Property AllowTeamsConsumer

$myText.Style = 'Normal'
$myText.typeText("Corporate users can message teams personal users: $($var1.AllowTeamsConsumer)")
$myText.TypeParagraph()
$myText.typeText("This setting should be false.")
$myText.TypeParagraph()

Write-Output "8.7 Contact with Skype Users SHALL be Disabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.7 Contact with Skype Users SHALL be Disabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CSTenantFederationConfiguration | Select-Object -Property AllowPublicUSers

$myText.Style = 'Normal'
$myText.typeText("Skype users allowed: $($var1.AllowPublicUSers)")
$myText.TypeParagraph()
$myText.typeText("This setting should be false.")
$myText.TypeParagraph()

Write-Output "8.8 Teams Email Integration SHALL be disabled"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.8 Teams Email Integration SHALL be disabled")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsClientConfiguration | Select-Object -Property AllowEmailIntoChannel

$myText.Style = 'Normal'
$myText.typeText("Allowed to Email into channels: $($var1.AllowEmailIntoChannel)")
$myText.TypeParagraph()
$myText.typeText("This setting should be false.")
$myText.TypeParagraph()

Write-Output "8.9 Only Approved Apps should be installed"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.9 Only Approved Apps should be installed")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsAppPermissionPolicy

foreach ($89 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($89.Name)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Default: $($89.DefaultCatalogAppsType)")
	$myText.TypeParagraph()
	$myText.TypeText("Apps: $($89.DefaultCatalogApps -join ", ")")
	$myText.TypeParagraph()
	$myText.TypeText("Global: $($89.GlobalCatalogAppsType)")
	$myText.TypeParagraph()
	$myText.TypeText("Apps: $($89.GlobalCatalogApps -join ", ")")
	$myText.TypeParagraph()
	$myText.TypeText("Private: $($89.PrivateCatalogAppsType)")
	$myText.TypeParagraph()
	$myText.TypeText("Apps: $($89.PrivateCatalogApps -join ", ")")
	$myText.TypeParagraph()
}

$myText.TypeParagraph()
$myText.typeText("The Apps Type should be in AllowedAppList with apps listed specifically in the setting.")
$myText.TypeParagraph()


Write-Output "8.10 Only the Meeting Organizer should be able to record live events"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.10 Only the Meeting Organizer should be able to record live events")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsMeetingBroadcastPolicy | Select-Object -Property Identity,BroadcastRecordingMode

foreach ($810 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($810.Identity)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Default: $($810.BroadcastRecordingMode)")
	$myText.TypeParagraph()
}
$myText.TypeParagraph()
$myText.typeText("The setting should be in UserOverride")
$myText.TypeParagraph()

Write-Output "8.11 Creation of new Teams Channels is restricted"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.11 Creation of new Teams Channels is restricted")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-CsTeamsChannelsPolicy | Select-Object -Property Identity,AllowOrgWideTeamCreation,AllowSharedChannelCreation,AllowPrivateChannelCreation

foreach ($811 in $var1)
{
	$myText.Style = 'Heading 3'
	$myText.TypeText("$($811.Identity)")
	$myText.TypeParagraph()
	$myText.Style = 'Normal'
	$myText.TypeText("Org Wide creation: $($811.AllowOrgWideTeamCreation)")
	$myText.TypeParagraph()
	$myText.TypeText("Shared creation: $($811.AllowSharedChannelCreation)")
	$myText.TypeParagraph()
	$myText.TypeText("Private creation: $($811.AllowPrivateChannelCreation)")
	$myText.TypeParagraph()
}

$myText.TypeParagraph()
$myText.typeText("The above settings should be false.  This will make it so that only administrators or specified indviduals can create teams.")
$myText.TypeParagraph()

Write-Output "8.12 Teams Channels have an expiration policy"

$myText.Font.Bold = 1
$myText.Style = 'Heading 2'
$myText.TypeText("8.12 Teams Channels have an expiration policy")
$myText.TypeParagraph()
$myText.Font.Bold = 0

$var1 = Get-AzureADMSGroupLifecyclePolicy

if ($var1 -eq $null)
{
$myText.typeText("No lifecycle policy exists.  This is considered a failure.")
$myText.TypeParagraph()
}
else
{
$myText.typeText("$($var1)")
$myText.TypeParagraph()
}

$myText.TypeParagraph()
$myText.typeText("This is the end of the file.  The script has completed successfully.  Please go to myITProcess and input the results, and make sure you follow the instructions for disposition of this file.")
$myText.TypeParagraph()

Write-Output "Script completed.  Saving File."
$fileprefix = Read-Host "Please enter the two digit month followed by the client code.  So East Coast Lumber for March would be '03 ECL'."

# Save and quit
$filename = 'C:\Temp\'+$fileprefix+' M365 Alignment.docx'
$saveFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatDocumentDefault
$mydoc.SaveAs([ref][system.object]$filename, [ref]$saveFormat)
$mydoc.Close()
$MSWord.Quit()
# Clean up Com object
$null =
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$MSWord)
Remove-Variable MSWord
Remove-Variable var1
