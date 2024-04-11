Write-Output "Running the Log Gathering Script.  Logging into the environment.  There will be several prompts."

Write-Output "Logging into Microsoft Graph.  If you have not run this before, log in as a global administrator to consent to the Graph permissions."
Connect-MgGraph -ContextScope Process -Scopes "User.Read.all", "Group.ReadWrite.All", "IdentityRiskEvent.Read.All", "IdentityRiskyUser.Read.All", "AuditLog.Read.All", "Directory.Read.All" -NoWelcome
Write-Output "Logging into Exchange Online"
Connect-ExchangeOnline -ShowBanner:$false
Write-Output "Logging into Azure AD"
Connect-AzureAD > $null
Write-Output "All Services connected"
Write-Output "Now please specify the file name.  Format should be: YY-MM-DD Clientcode.  So '24-04-05 ECL' would be for East Coast Lumber, run on the 5th of April 2024.  It will append the .xlsx automatically."
$logfilename = Read-Host "Filename: " 
$logfilename = $logfilename+".xlsx"

$filterDate = (Get-Date).AddDays(-7)
$AuditLogs = Get-MgAuditLogDirectoryAudit | Where-Object {$_.ActivityDateTime -GT $filterDate} | Select-Object ActivityDateTime, ActivitydisplayName, Category, @{N='User'; E={$null}}, @{N='Target'; E={$null}}, OperationType, Result, ResultReason, InitiatedBy, TargetResources
 
foreach($log in $AuditLogs){
 
    $log.User = $log.InitiatedBy.User.UserPrincipalName
    $log.Target = $log.TargetResources.UserPrincipalName
}

Get-MgRiskDetection|Where { $_.ActivityDateTime -GT $filterDate} | Select-Object -property ActivityDateTime,Activity,UserDisplayName,UserPrincipalName,RiskState,RiskDetail,Location | Export-Excel -WorksheetName "5.3 Risky Users" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.3 Ensure the Azure AD 'Risky sign-ins' report is reviewed at least (weekly)" -TitleSize 16
$AuditLogs | Select-Object ActivityDateTime, ActivitydisplayName, Category, User, Target, OperationType, Result, ResultReason | Export-Excel -WorksheetName "5.4, 5.5, 5.6" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.4 Ensure the Application Usage report, (5.5) the SSPR Reset Activity, and (5.6) User role group role changes are reviewed weekly" -TitleSize 16
Get-Mailbox -ResultSize Unlimited -Filter "ForwardingAddress -like '*' -or ForwardingSmtpAddress -like '*'" | Select-Object Name,ForwardingAddress,ForwardingSmtpAddress | Export-Excel -WorksheetName "5.7 Mail Forwarding" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.7 Ensure mail forwarding rules are reviewed at least (weekly)" -TitleSize 16
$forwardrules = get-mailbox | foreach { get-inboxRule –mailbox $_.alias } | where-object { ( $_.forwardAsAttachmentTo –ne $NULL ) –or ( $_.forwardTo –ne $NULL ) –or ( $_.redirectTo –ne $NULL ) }
$forwardArray = @()
$forwardrules | foreach { $forwardarray = $forwardarray += get-inboxrule -identity $_.identity | Select-Object -property Identity,Name,@{name = 'Forwarded To'; expression = { $_.ForwardTo -join ', ' }},@{name = 'Forwarded As Attachment To'; expression = { $_.ForwardasAttachmentTo -join ', '}},@{name = 'RedirectTo'; expression = { $_.RedirectTo -join ', '}}}
$forwardArray | Export-Excel -WorksheetName "5.7 Mail Forwarding 2" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.7 Ensure mail forwarding rules are reviewed at least (weekly)" -TitleSize 16


$StartDate=(((Get-Date).AddDays(-7))).Date
$EndDate=Get-Date


$OutputCSV=".\NonOwner-Mailbox-Access-Report_$((Get-Date -format yyyy-MMM-dd-ddd` hh-mm` tt).ToString()).csv"
$IntervalTimeInMinutes=1440    #$IntervalTimeInMinutes=Read-Host Enter interval time period '(in minutes)'
$CurrentStart=$StartDate
$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
$Operation='ApplyRecord','Copy','Create','FolderBind','HardDelete','MessageBind','Move','MoveToDeletedItem','RecordDelete','SendAs','SendOnBehalf','SoftDelete','Update','UpdateCalendarDelegation','UpdateFolderPermissions','UpdateInboxRules'


#Check whether CurrentEnd exceeds EndDate(checks for 1st iteration)
if($CurrentEnd -gt $EndDate)
{
 $CurrentEnd=$EndDate
}

$AggregateResults = 0
$CurrentResult= @()
$CurrentResultCount=0
$NonOwnerAccess=0
Write-Host `nRetrieving audit log from $StartDate to $EndDate... -ForegroundColor Yellow

while($true)
{
 #Write-Host Retrieving audit log between StartDate $CurrentStart to EndDate $CurrentEnd ******* IntervalTime $IntervalTimeInMinutes minutes
 if($CurrentStart -eq $CurrentEnd)
 {
  Write-Host Start and end time are same.Please enter different time range -ForegroundColor Red
  Exit
 }


 #Getting Non-Owner mailbox access for a given time range
 else
 {
  $Results=Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd -Operations $Operations -SessionId s -SessionCommand ReturnLargeSet -ResultSize 5000
 }
 $AllAuditData=@()
 $AllAudits=
 foreach($Result in $Results)
 {
  $AuditData=$Result.auditdata | ConvertFrom-Json

  #Remove owner access
  if($AuditData.LogonType -eq 0)
  {
   continue
  }

  #Filter for external access
  if(($IncludeExternalAccess -eq $false) -and ($AuditData.ExternalAccess -eq $true))
  {
   continue
  }

  #Processing non-owner mailbox access records
  if(($AuditData.LogonUserSId -ne $AuditData.MailboxOwnerSid) -or ((($AuditData.Operation -eq "SendAs") -or ($AuditData.Operation -eq "SendOnBehalf")) -and ($AuditData.UserType -eq 0)))
  {
   $AuditData.CreationTime=(Get-Date($AuditData.CreationTime)).ToLocalTime()
   if($AuditData.LogonType -eq 1)
   {
    $LogonType="Administrator"
   }
   elseif($AuditData.LogonType -eq 2)
   {
    $LogonType="Delegated"
   }
   else
   {
    $LogonType="Microsoft datacenter"
   }
   if($AuditData.Operation -eq "SendAs")
   {
    $AccessedMB=$AuditData.SendAsUserSMTP
    $AccessedBy=$AuditData.MailboxOwnerUPN
   }
   elseif($AuditData.Operation -eq "SendOnBehalf")
   {
    $AccessedMB=$AuditData.SendOnBehalfOfUserSmtp
    $AccessedBy=$AuditData.MailboxOwnerUPN
   }
   else
   {
    $AccessedMB=$AuditData.MailboxOwnerUPN
    $AccessedBy=$AuditData.UserId
   }
   if($AccessedMB -eq $AccessedBy)
   {
    Continue
   }
  $NonOwnerAccess++
  $AllAudits=@{'Access Time'=$AuditData.CreationTime;'Accessed by'=$AccessedBy;'Performed Operation'=$AuditData.Operation;'Accessed Mailbox'=$AccessedMB;'Logon Type'=$LogonType;'Result Status'=$AuditData.ResultStatus;'External Access'=$AuditData.ExternalAccess}
  $AllAuditData= New-Object PSObject -Property $AllAudits
  $AllAuditData | Sort 'Access Time','Accessed by' | select 'Access Time','Logon Type','Accessed by','Performed Operation','Accessed Mailbox','Result Status','External Access' | Export-Excel -WorksheetName "5.8 Non Owner Access" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.8 Mailbox access by owner" -TitleSize 16
 }
 }
 #$CurrentResult += $Results
 $currentResultCount=$CurrentResultCount+($Results.count)
 $AggregateResults +=$Results.count
 if(($CurrentResultCount -eq 50000) -or ($Results.count -lt 5000))
 {
  if($CurrentResultCount -eq 50000)
  {
   Write-Host Retrieved max record for the current range.Proceeding further may cause data loss or rerun the script with reduced time interval. -ForegroundColor Red
   $Confirm=Read-Host `nAre you sure you want to continue? [Y] Yes [N] No
   if($Confirm -notmatch "[Y]")
   {
    Write-Host Please rerun the script with reduced time interval -ForegroundColor Red
    Exit
   }
   else
   {
    Write-Host Proceeding audit log collection with data loss
   }
  }
  #Check for last iteration
  if(($CurrentEnd -eq $EndDate))
  {
   break
  }
  [DateTime]$CurrentStart=$CurrentEnd
  #Break loop if start date exceeds current date(There will be no data)
  if($CurrentStart -gt (Get-Date))
  {
   break
  }
  [DateTime]$CurrentEnd=$CurrentStart.AddMinutes($IntervalTimeInMinutes)
  if($CurrentEnd -gt $EndDate)
  {
   $CurrentEnd=$EndDate
  }

  $CurrentResultCount=0
  $CurrentResult = @()
 }
}
Get-MailDetailATPReport | Where-Object -Property EventType -ne "Message Passed" | Export-Excel -WorksheetName "5.9 Malware" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.9 Ensure the Malware Detections report is reviewed at least (weekly)" -TitleSize 16
Get-MgAuditLogProvisioning|Where { $_.ActivityDateTime -GT $filterDate} | Export-Excel -WorksheetName "5.10 Provisioning" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.10 Ensure the Account Provisioning Activity report is reviewed at least (weekly)" -TitleSize 16

$DirectoryRoles = Get-AzureADDirectoryRole | Select DisplayName, ObjectId # This command dumps all of the Azure AD roles that have ever been assigned into a variable
$drresults = @()
foreach($role in $DirectoryRoles){ # This foreach loop writes the current role it's checking assignments for to the screen then get's a list of users with that role and dumps it into a CSV, then repeats for all remaining Azure AD roles.

    $drresults = $drresults += Get-AzureADDirectoryRoleMember -ObjectId $role.ObjectId | Select @{Name="Azure AD Role"; Expression={$role.DisplayName}}, DisplayName, UserPrincipalName 

}
$drresults | Export-Excel -WorksheetName "5.11 Administrators" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.11 Ensure non-global administrator role group assignments are reviewed at least (weekly)" -TitleSize 16
get-SpoofMailReport -startdate $filterdate -enddate $enddate | Export-Excel -WorksheetName "5.12 Spoofed Domains" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.12 Ensure the spoofed domains report is review weekly" -TitleSize 16
get-blockedsenderaddress | Export-Excel -WorksheetName "5.14 Spam Restricted" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.14 Ensure the report of users who have had their email privileges restricted due to spamming is reviewed" -TitleSize 16
get-azureaduser | where-object -property UserType -EQ Guest | Select-Object -property DisplayName,Mail,MailNickName,AccountEnabled,UserState,ShowInAddressList | Export-Excel -WorksheetName "5.15 Guest Users" -Path c:\Temp\$logfilename -BoldTopRow -AutoSize -Title "5.15 (L1) Ensure Guest Users are reviewed at least (biweekly)" -TitleSize 16
