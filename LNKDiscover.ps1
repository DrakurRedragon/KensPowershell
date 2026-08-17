# This will scan a root with recursion for every .lnk file, which are folder level shortcuts, and return the name, location, and the target path for each.

# Change the Root Path of where to start

$rootpath = "E:\DFSRoot\FolderRedirect\UserProfiles"

# Change this to the path for your CSV:

$logpath = "C:\Temp\results.csv"

# The script will run, silently skipping all errors

$lnkfiles = Get-ChildItem -LiteralPath $rootpath -Filter *.lnk -Recurse -File | Where-Object { $_.FullName -notmatch '\\AppData\\' }

$shell = New-Object -ComObject Wscript.shell
$results = New-Object System.Collections.Generic.List[object]

foreach ($file in $lnkfiles){
    $shortcut = $shell.CreateShortcut($file.FullName)
    $target = $shortcut.TargetPath
    $name = $file.name
    $loc = $file.FullName
    $results.Add([PSCustomObject]@{
        Name = $name
        TargetPath = $target
        Location=$loc
    })
}

$results | Export-CSV -Path $logpath -NoTypeInformation