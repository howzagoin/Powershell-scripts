# Test script for dialog functions

function Get-SaveFileDialog {
    param(
        [string]$InitialDirectory = [Environment]::GetFolderPath('Desktop'),
        [string]$Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
        [string]$DefaultFileName = "SharePointAudit.xlsx",
        [string]$Title = "Save SharePoint Audit Report"
    )
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.InitialDirectory = $InitialDirectory
        $SaveFileDialog.Filter = $Filter
        $SaveFileDialog.FileName = $DefaultFileName
        $SaveFileDialog.Title = $Title
        $SaveFileDialog.DefaultExt = "xlsx"
        $SaveFileDialog.AddExtension = $true
        
        $result = $SaveFileDialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $SaveFileDialog.FileName
        } else {
            return $null
        }
    }
    catch {
        Write-Host "[Warning] Could not show save dialog. Using default filename in current directory." -ForegroundColor Yellow
        return $DefaultFileName
    }
}

function Get-SearchTermsFromUser {
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic
        $searchTerms = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter search terms for the audit (optional - leave blank for 'AllSites'):",
            "SharePoint Audit Search Terms",
            "AllSites"
        )
        if ([string]::IsNullOrWhiteSpace($searchTerms)) {
            return "AllSites"
        }
        return $searchTerms.Trim()
    }
    catch {
        Write-Host "[Warning] Could not show input dialog. Using default search term." -ForegroundColor Yellow
        return "AllSites"
    }
}

# Test the dialogs
Write-Host "Testing Search Terms Dialog..." -ForegroundColor Cyan
$searchTerms = Get-SearchTermsFromUser
Write-Host "Search Terms: $searchTerms" -ForegroundColor Green

$tenantName = "TestTenant"
$dateStr = Get-Date -Format yyyyMMdd_HHmmss
$cleanSearchTerms = $searchTerms -replace '[^\w\s-]', '' -replace '\s+', '_'
$defaultFileName = "TenantAudit-$cleanSearchTerms-$tenantName-$dateStr.xlsx"

Write-Host "Testing Save File Dialog..." -ForegroundColor Cyan
$filePath = Get-SaveFileDialog -DefaultFileName $defaultFileName

if ($filePath) {
    Write-Host "Selected file path: $filePath" -ForegroundColor Green
} else {
    Write-Host "Dialog was cancelled" -ForegroundColor Yellow
}
