<#

.PARAMETER RootPath
    Root of the tree to scan, e.g. E:\DFSRoot\FolderRedirect\UserProfiles

.PARAMETER CsvPath
    Path to the CSV containing TargetPath -> URL mappings.

.PARAMETER LogPath
    Where to write a CSV log of every action taken (or that would be taken
    in -DryRun mode). Defaults to .\ShortcutConversion_log.csv

.PARAMETER DryRun
    Default is $true. When true, NOTHING is deleted or moved - the script
    only reports what it WOULD do, to the console and the log CSV. Re-run
    with -DryRun:$false once you've reviewed the log and are confident.

.EXAMPLE
    # Safe first pass - see what would happen, nothing is changed
    .\ConvertLnk.ps1 -RootPath X:\Wherever -CsvPath Y:\Convert.csv
    add -DryRun:$false to the end to run for real

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RootPath,

    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [string]$LogPath = ".\ShortcutConversion_log.csv",

    [bool]$DryRun = $true
)

if (-not (Test-Path -LiteralPath $RootPath)) {
    throw "RootPath '$RootPath' is not reachable."
}
if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CsvPath '$CsvPath' was not found."
}

if ($DryRun) {
    Write-Host "=== DRY RUN MODE - no files will be changed, moved, or deleted ===" -ForegroundColor Yellow
    Write-Host "    Re-run with -DryRun:`$false once you've reviewed the log.`n" -ForegroundColor Yellow
}

$csv = Import-Csv -Path $CsvPath

if (-not $csv -or $csv.Count -eq 0) {
    throw "CSV at '$CsvPath' has no rows."
}

$firstRow = $csv[0]
if (-not ($firstRow.PSObject.Properties.Name -contains 'TargetPath')) {
    throw "CSV must contain a 'TargetPath' column. Found columns: $($firstRow.PSObject.Properties.Name -join ', ')"
}

$urlColumn = $null
foreach ($prop in $firstRow.PSObject.Properties.Name) {
    if ($prop -eq 'TargetPath') { continue }
    $sampleValue = ($csv | Where-Object { $_.$prop -match '^https?://' } | Select-Object -First 1).$prop
    if ($sampleValue) {
        $urlColumn = $prop
        break
    }
}

if (-not $urlColumn) {
    throw "Could not auto-detect a URL column in the CSV (looked for values starting with http:// or https://). Columns found: $($firstRow.PSObject.Properties.Name -join ', ')"
}

Write-Host "Using CSV column '$urlColumn' as the URL source.`n" -ForegroundColor Cyan

function Get-NormalizedPath {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    return $Path.Trim().TrimEnd('\').ToLowerInvariant()
}

$targetLookup = @{}
foreach ($row in $csv) {
    $key = Get-NormalizedPath $row.TargetPath
    if ($key -and -not $targetLookup.ContainsKey($key)) {
        $targetLookup[$key] = $row.$urlColumn
    }
}

Write-Host "Loaded $($targetLookup.Count) unique target path(s) from CSV.`n" -ForegroundColor Cyan

# --- Find candidate .lnk files under Desktop folders ---------------------------------------------------
Write-Host "Scanning '$RootPath' for .lnk files under Desktop folders..." -ForegroundColor Cyan

$allLnk = Get-ChildItem -Recurse -LiteralPath $RootPath -File -Filter *.lnk -ErrorAction SilentlyContinue
$desktopLnk = $allLnk | Where-Object { $_.DirectoryName -imatch '(^|\\)Desktop(\\|$)' }

Write-Host "Found $($desktopLnk.Count) .lnk file(s) under Desktop folders (out of $($allLnk.Count) total .lnk files scanned).`n" -ForegroundColor Cyan

# --- Process ---------------------------------------------------
$shell = New-Object -ComObject WScript.Shell
$log = New-Object System.Collections.Generic.List[object]
$i = 0
$total = $desktopLnk.Count

foreach ($file in $desktopLnk) {
    $i++
    if ($i % 25 -eq 0 -or $i -eq $total) {
        Write-Progress -Activity "Processing Desktop shortcuts" -Status "$i of $total" -PercentComplete (($i / [math]::Max($total,1)) * 100)
    }

    $logEntry = [PSCustomObject]@{
        LnkPath     = $file.FullName
        Target      = $null
        Action      = $null
        Destination = $null
        Note        = $null
    }

    try {
        $shortcut = $shell.CreateShortcut($file.FullName)
        $target = $shortcut.TargetPath
        $logEntry.Target = $target

        $key = Get-NormalizedPath $target
        $matchedUrl = if ($key -and $targetLookup.ContainsKey($key)) { $targetLookup[$key] } else { $null }

        if ($matchedUrl) {
            # --- MATCH: convert to .url and delete the .lnk ---
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
            $urlFilePath = Join-Path -Path $file.DirectoryName -ChildPath "$baseName.url"

            $logEntry.Action = 'Convert'
            $logEntry.Destination = $urlFilePath
            $logEntry.Note = "Matched URL: $matchedUrl"

            if (-not $DryRun) {
                $urlContent = "[InternetShortcut]`r`nURL=$matchedUrl`r`n"
                Set-Content -LiteralPath $urlFilePath -Value $urlContent -Encoding ASCII -Force
                Remove-Item -LiteralPath $file.FullName -Force
            }
        }
        else {
            # --- NO MATCH: move to OldShortcuts one level up from Desktop ---
            # Find the "Desktop" segment in the path and take everything before it
            $parts = $file.DirectoryName -split '\\'
            $desktopIndex = -1
            for ($p = 0; $p -lt $parts.Count; $p++) {
                if ($parts[$p] -ieq 'Desktop') { $desktopIndex = $p; break }
            }

            if ($desktopIndex -lt 1) {
                $logEntry.Action = 'Skip'
                $logEntry.Note = "Could not resolve a parent-of-Desktop folder for this path."
                $log.Add($logEntry)
                continue
            }

            $desktopParent = ($parts[0..($desktopIndex - 1)]) -join '\'
            $oldShortcutsFolder = Join-Path -Path $desktopParent -ChildPath 'OldShortcuts'
            $destPath = Join-Path -Path $oldShortcutsFolder -ChildPath $file.Name

            # Avoid collisions if a file with the same name already exists there
            if (Test-Path -LiteralPath $destPath) {
                $base = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
                $ext = [System.IO.Path]::GetExtension($file.Name)
                $counter = 1
                do {
                    $destPath = Join-Path -Path $oldShortcutsFolder -ChildPath "$base ($counter)$ext"
                    $counter++
                } while (Test-Path -LiteralPath $destPath)
            }

            $logEntry.Action = 'Move'
            $logEntry.Destination = $destPath
            $logEntry.Note = if ([string]::IsNullOrWhiteSpace($target)) { "No target / unresolved shortcut" } else { "No matching TargetPath in CSV" }

            if (-not $DryRun) {
                if (-not (Test-Path -LiteralPath $oldShortcutsFolder)) {
                    New-Item -Path $oldShortcutsFolder -ItemType Directory -Force | Out-Null
                }
                Move-Item -LiteralPath $file.FullName -Destination $destPath -Force
            }
        }
    }
    catch {
        $logEntry.Action = 'Error'
        $logEntry.Note = $_.Exception.Message
    }

    $log.Add($logEntry)
}

Write-Progress -Activity "Processing Desktop shortcuts" -Completed

# --- Write log / summary ---------------------------------------------------
$log | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

$converted = ($log | Where-Object { $_.Action -eq 'Convert' }).Count
$moved     = ($log | Where-Object { $_.Action -eq 'Move' }).Count
$skipped   = ($log | Where-Object { $_.Action -eq 'Skip' }).Count
$errors    = ($log | Where-Object { $_.Action -eq 'Error' }).Count

Write-Host ""
Write-Host "Done. $(if ($DryRun) { '(DRY RUN - nothing was actually changed)' })" -ForegroundColor Green
Write-Host "  Matched / converted to .url : $converted"
Write-Host "  No match / moved to OldShortcuts: $moved"
Write-Host "  Skipped (unresolvable path) : $skipped"
Write-Host "  Errors                      : $errors"
Write-Host "  Full log written to         : $LogPath"