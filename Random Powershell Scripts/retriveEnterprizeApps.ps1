<#
.SYNOPSIS
  Streams non-Microsoft Enterprise Apps to console. Zero stalling.

.AUTHOR
  Tim MacLatchy

.DATE
  15-07-2025

.LICENSE
  MIT License
#>

# ------------------- Module Check & Import (diagnostics) -------------------
Write-Host "[INFO] Checking for Microsoft.Graph module..." -ForegroundColor Yellow
if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) {
    Write-Host "[INFO] Microsoft.Graph module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force
}
Write-Host "[INFO] Importing Microsoft.Graph module..." -ForegroundColor Yellow
Import-Module Microsoft.Graph -Force
Write-Host "[INFO] Microsoft.Graph module imported." -ForegroundColor Green

# ------------------- Connect (force fresh login, diagnostics) -------------------
Write-Host "[INFO] Disconnecting any existing Graph session..." -ForegroundColor Yellow
Disconnect-MgGraph -ErrorAction SilentlyContinue
Write-Host "[INFO] Connecting to Microsoft Graph (force fresh login)..." -ForegroundColor Yellow
try {
    Connect-MgGraph -Scopes "Application.Read.All", "Directory.Read.All" -ErrorAction Stop
    Write-Host "[INFO] Connected to Microsoft Graph." -ForegroundColor Green
} catch {
    Write-Error "[ERROR] Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit 1
}

# ------------------- Enumerate App Registrations and Enterprise Apps -------------------

Write-Host "[INFO] Retrieving all application registrations using Get-MgApplication..." -ForegroundColor Cyan
try {
    $apps = Get-MgApplication -All
    $idpApps = $apps | Where-Object { $_.SignInAudience -eq 'AzureADMyOrg' -or $_.SignInAudience -eq 'AzureADMultipleOrgs' }
    $totalIdpApps = $idpApps.Count
    Write-Host ("`nTotal applications using this tenant as their Identity Provider: {0}" -f $totalIdpApps) -ForegroundColor Green
    $idpApps | ForEach-Object { Write-Host ("- $($_.DisplayName) [$($_.AppId)] (SignInAudience: $($_.SignInAudience))") }
} catch {
    Write-Error "[ERROR] Failed to retrieve applications: $($_.Exception.Message)"
}

Write-Host "[INFO] Retrieving all enterprise applications (service principals) using Get-MgServicePrincipal..." -ForegroundColor Cyan
try {
    $sps = Get-MgServicePrincipal -All
    $spsWithUsers = @()
    foreach ($sp in $sps) {
        $users = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
        if ($users -and $users.Count -gt 0) {
            $spsWithUsers += $sp
        }
    }
    $totalSPsWithUsers = $spsWithUsers.Count
    Write-Host ("Total enterprise applications with assigned users: {0}" -f $totalSPsWithUsers) -ForegroundColor Green
    # Optional: List service principal names with users
    # $spsWithUsers | ForEach-Object { Write-Host ("- $($_.DisplayName) [$($_.AppId)]") }
    Write-Host "✅ Done!" -ForegroundColor Green
} catch {
    Write-Error "[ERROR] Failed to retrieve service principals or user assignments: $($_.Exception.Message)"
}
