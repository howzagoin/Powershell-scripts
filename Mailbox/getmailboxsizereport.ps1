# --- Function: Get Mailbox Size Report for All Users ---
function Get-AllMailboxSizeReport {
    param(
        [string]$ExportPath = $("MailboxSizeReport_" + (Get-Date -Format 'yyyy-MM-dd_HHmmss') + ".csv")
    )
    $results = @()
    $mailboxes = Get-Mailbox -ResultSize Unlimited
    $mbCount = 0
    foreach ($mb in $mailboxes) {
        $stats = Get-MailboxStatistics -Identity $mb.UserPrincipalName
        $result = [PSCustomObject]@{
            DisplayName           = $mb.DisplayName
            UserPrincipalName     = $mb.UserPrincipalName
            MailboxType           = $mb.RecipientTypeDetails
            PrimarySMTPAddress    = $mb.PrimarySMTPAddress
            ArchiveStatus         = if (($mb.ArchiveDatabase -eq $null) -and ($mb.ArchiveDatabaseGuid -eq $mb.ArchiveGuid)) { "Disabled" } else { "Active" }
            ItemCount             = $stats.ItemCount
            TotalSize             = $stats.TotalItemSize.ToString().Split('(')[0].Trim()
            TotalSizeBytes        = ($stats.TotalItemSize -replace "(.*\\()|,| [a-z]*\\)", "")
            DeletedItemCount      = $stats.DeletedItemCount
            DeletedItemSize       = $stats.TotalDeletedItemSize
            IssueWarningQuota     = $mb.IssueWarningQuota -replace "\\(.*",""
            ProhibitSendQuota     = $mb.ProhibitSendQuota -replace "\\(.*",""
            ProhibitSendReceiveQuota = $mb.ProhibitSendReceiveQuota -replace "\\(.*",""
        }
        $results += $result
        $mbCount++
        Write-Progress -Activity "Processing mailboxes" -Status "Mailbox $mbCount: $($mb.DisplayName)" -PercentComplete (($mbCount / $mailboxes.Count) * 100)
    }
    $results | Export-Csv -Path $ExportPath -NoTypeInformation
    Write-Host "Mailbox size report exported to $ExportPath" -ForegroundColor Green
    return $ExportPath
}