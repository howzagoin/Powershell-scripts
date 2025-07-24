# ---------------------------------------------------------
# Author: Tim MacLatchy
# Date: 16-07-2025
# License: MIT
# Description: List all calendar folders (default + custom) in a mailbox and show who has access
# ---------------------------------------------------------

# Connect as admin
Connect-ExchangeOnline -UserPrincipalName "timothy.maclatchy@journebrands.com"






# Robust calendar scan for all users (based on TenantAudit copy 2.ps1)
$mailboxes = Get-Mailbox -ResultSize Unlimited
$total = $mailboxes.Count
$i = 0
foreach ($mb in $mailboxes) {
    $i++
    Write-Progress -Activity "Enumerating Shared Calendars" -Status "Processing $($mb.UserPrincipalName) ($i of $total)" -PercentComplete (($i / $total) * 100)
    $folders = Get-MailboxFolderStatistics -Identity $mb.UserPrincipalName -ErrorAction SilentlyContinue
    $calendarFolders = $folders | Where-Object {
        $_.FolderPath -match "^/Calendar" -or $_.FolderType -like '*Calendar*'
    }
    foreach ($folder in $calendarFolders) {
        $cleanPath = $folder.FolderPath.TrimStart("/").Replace("/", "\")
        if ($cleanPath -eq "\\") { $cleanPath = "Calendar" }
        # Skip problematic folders
        if ($folder.Name -eq "Calendar Logging" -or [string]::IsNullOrWhiteSpace($cleanPath)) {
            continue
        }
        $folderIdentity = "${mb.UserPrincipalName}:\$cleanPath"
        try {
            $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
            $sharedPerms = $permissions | Where-Object { $_.User -notin @('Default','Anonymous') -and $_.AccessRights -ne 'None' }
            if ($sharedPerms) {
                Write-Host "--- Mailbox: $($mb.UserPrincipalName) | Folder: $($folder.Name) | Path: $folderIdentity | Class: $($folder.FolderClass) ---"
                Write-Host "Permissions for ${folderIdentity}:"
                foreach ($entry in $sharedPerms) {
                    Write-Host "User: $($entry.User) | Access Rights: $($entry.AccessRights)"
                }
                Write-Host ""
            }
        } catch {
            Write-Host "Failed to get permissions for $folderIdentity. Error: $_"
        }
    }
}
Write-Progress -Activity "Enumerating Shared Calendars" -Completed

# Disconnect
Disconnect-ExchangeOnline -Confirm:$false
