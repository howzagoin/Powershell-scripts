# ── CONFIGURATION ─────────────────────────────────────────────────────────
$SiteURL    = "https://fbaint.sharepoint.com/sites/Marketing"
$FolderName = "Mode 2.0"
$ParentPath = "sites/Marketing/Shared Documents/Photography"
$BatchSize  = 10000            # size per get-pnp batch
$LogCsv     = "Mode2.0_RestoreLog.csv"
$LocalFolder= "C:\Mode 2.0 restored"  # optional local directory

# ── CREATE LOCAL FOLDER (optional) ───────────────────────────────────────
if (-not (Test-Path $LocalFolder)) {
    New-Item –Path $LocalFolder –ItemType Directory | Out-Null
}

# ── CONNECT ───────────────────────────────────────────────────────────────
Connect-PnPOnline -Url $SiteURL -Interactive

# ── FETCH & RESTORE IN BATCHES ─────────────────────────────────────────
$log = @()
$moreItems = $true

while ($moreItems) {
    try {
        $items = Get-PnPRecycleBinItem -RowLimit $BatchSize -FirstStage
        $second = Get-PnPRecycleBinItem -RowLimit $BatchSize -SecondStage
        $all = $items + $second
    }
    catch {
        Write-Host "Error fetching items: $_" -ForegroundColor Red
        break
    }

    if (-not $all) {
        $moreItems = $false
        break
    }

    $filtered = $all | Where-Object {
        ($_.ItemType -eq "Folder" -and $_.Title -eq $FolderName -and $_.DirName -eq $ParentPath) -or
        ($_.DirName -like "$ParentPath/$FolderName/*")
    }

    foreach ($it in $filtered) {
        try {
            Restore-PnPRecycleBinItem -Identity $it -Force
            Write-Host "Restored: $($it.Title) at $($it.DirName)" -ForegroundColor Green

            $log += [PSCustomObject]@{
                ID          = $it.Id
                Title       = $it.Title
                DirName     = $it.DirName
                DeletedDate = $it.DeletedDate
                RestoredAt  = (Get-Date)
            }
        }
        catch {
            Write-Host "Failed to restore ID $($it.Id): $_" -ForegroundColor Yellow
        }
    }

    # throttle-friendly pause
    Start-Sleep -Seconds 5

    # exit if batch didn't fill
    if ($all.Count -lt $BatchSize) { $moreItems = $false }
}

# ── WRITE LOG ─────────────────────────────────────────────────────────────
if ($log.Count) {
    $log | Export-Csv -Path $LogCsv -NoTypeInformation
    Write-Host "Restore log saved to $LogCsv" -ForegroundColor Cyan
}
else {
    Write-Host "No matching items found to restore." -ForegroundColor Yellow
}
