

#clear TPM - uses the owner authorization value stored in the registry instead of specifying a value or using a value in a file. You can read more on this at https://docs.microsoft.com/en-us/powershell/module/trustedplatformmodule/clear-tpm?view=win10-ps
# Define the path for the backup file
$backupFilePath = "C:\TPMBackup\TpmOwnerAuthBackup.txt"

# Create the backup directory if it does not exist
if (-not (Test-Path "C:\TPMBackup")) {
    New-Item -Path "C:\TPMBackup" -ItemType Directory
}

# Backup the TPM owner authorization
try {
    $tpm = Get-Tpm
    if ($tpm.TpmPresent -eq $true -and $tpm.TpmReady -eq $true) {
        Backup-TpmOwnerAuth -Path $backupFilePath
        Write-Output "TPM owner authorization has been backed up to $backupFilePath"
    } else {
        Write-Output "TPM is not present or not ready."
        exit
    }
} catch {
    Write-Output "Failed to backup TPM owner authorization. Error: $_"
    exit
}

# Clear the TPM
try {
    $confirmation = Read-Host "Are you sure you want to clear the TPM? This action cannot be undone. Type 'Yes' to confirm"
    if ($confirmation -eq "Yes") {
        Clear-Tpm
        Write-Output "TPM has been cleared. Your computer will restart now."
        Restart-Computer -Force
    } else {
        Write-Output "TPM clear operation canceled."
    }
} catch {
    Write-Output "Failed to clear the TPM. Error: $_"
}


# Restart the device to apply changes
Write-Output "Restarting the device to apply changes..."
Restart-Computer -Force
