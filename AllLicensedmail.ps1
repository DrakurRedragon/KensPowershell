New-DynamicDistributionGroup -Name "orthoct.hoppercomms" -Alias "orthoct.hoppercomms" -PrimarySMTPAddress "orthoct.hoppercomms@myorthoct.com" -RecipientFilter {(RecipientTypeDetails -eq 'UserMailbox' -and AccountDisabled -eq $false)}
Set-DynamicDistributionGroup -identity "orthoct.hoppercomms" -HiddenFromAddressListsEnabled $true -RequireSenderAuthenticationEnabled $false
