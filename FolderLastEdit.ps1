$topPath = "E:\PREMIERDC01\PT"
$folderResults = @()

# Check subfolders
Get-ChildItem -Path $topPath | Where-Object { $_.PSIsContainer } | ForEach-Object {
    $newest = Get-ChildItem -Path $_.FullName -Recurse -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($newest) {
        "{0} - {1}" -f $_.Name, $newest.LastWriteTime
        $folderResults += New-Object PSObject -Property @{
            Folder = $_.Name
            NewestDate = $newest.LastWriteTime
        }
    } else {
        "{0} - (no files found)" -f $_.Name
    }
}

# Check files directly in the top-level folder
$topLevelNewest = Get-ChildItem -Path $topPath -ErrorAction SilentlyContinue | Where-Object { -not $_.PSIsContainer } | Sort-Object LastWriteTime -Descending | Select-Object -First 1

if ($topLevelNewest) {
    "(top-level files) - {0}" -f $topLevelNewest.LastWriteTime
    $folderResults += New-Object PSObject -Property @{
        Folder = "(top-level files)"
        NewestDate = $topLevelNewest.LastWriteTime
    }
}

# Determine overall newest across subfolders AND top-level files
$overallNewest = $folderResults | Sort-Object NewestDate -Descending | Select-Object -First 1

if ($overallNewest) {
    Write-Host ""
    Write-Host ("Newest overall: {0} - {1}" -f $overallNewest.Folder, $overallNewest.NewestDate) -ForegroundColor Green
}