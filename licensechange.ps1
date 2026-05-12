$usercsv = import-csv -path "c:\temp\user.csv"
foreach ($i in $usercsv)
{
    $usern = $i.UserPrincipalName

    $user = invoke-MGGraphRequest -Method GET -URI https://graph.microsoft.com/v1.0/users/$usern

    $userid = $user.id

    $groupid = "2f53cffb-d130-44af-91c1-aa9b2aa44c8f"

    Invoke-MGGraphRequest -Method POST -URI "https://graph.microsoft.com/v1.0/groups/$groupid/members/`$ref" -ContentType application/json -Body @{"@odata.id"= "https://graph.microsoft.com/v1.0/directoryObjects/$userid"}
    
    # 6fd2c87f-b296-42f0-b197-1e91e994b900 is E3
    # 32b47245-eb31-44fc-b945-a8b1576c439f is Purview

    $licenseparam = 
    @{
    "addLicenses"= @()
    "removeLicenses"= 
        @(
            "6fd2c87f-b296-42f0-b197-1e91e994b900",
            "32b47245-eb31-44fc-b945-a8b1576c439f"     
        )
    }

    Invoke-MgGraphRequest -Method POST -URI "https://graph.microsoft.com/v1.0/users/$usern/assignLicense" -ContentType application/json -Body $licenseparam
}
