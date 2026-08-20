$profilePath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

Get-ChildItem $profilePath | ForEach-Object {
    $profileName = $_.PSChildName
    $profileKey  = $_.PSPath
    $rootProps = Get-ItemProperty -Path $profileKey -ErrorAction SilentlyContinue
    $childProps = Get-ChildItem $profileKey -Recurse | ForEach-Object {
        Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
    }
    # Pull all property values from this profile's subkeys
    $subValues = Get-ChildItem $profileKey -Recurse | ForEach-Object {
        Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
    }

    $match = $subValues | Where-Object {
        ($_."Account Name" -like "*orthopaedicspecailty.com*") -or ($_."Email" -like "*orthopaedicspecailty.com*")
    }

    if ($match) {
        Write-Host "Deleting profile: $profileName"
        Remove-Item -Path $profileKey -Recurse -Force
    }
}