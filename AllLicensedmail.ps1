connect-exchangeonline
#Create the new dynamic distro
New-DynamicDistributionGroup -Name "OSG All Users" -Alias "OSGAllUsers" -PrimarySMTPAddress "osgallusers@osgpc.com" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox' -and AccountDisabled -eq $false)}
#Set the values for hidden and sender authentication enabled
Set-DynamicDistributionGroup -identity "OSG All Users" -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox' -and AccountDisabled -eq $false -and Company -eq 'OSG')}
