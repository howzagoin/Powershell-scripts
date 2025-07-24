# Test_EnterpriseApps_FINAL.ps1
# Return ALL Enterprise Applications that are NOT owned by Microsoft
# Author: Tim MacLatchy – 2025-07-16 – MIT

# 1 ── Connect
Connect-MgGraph -Scopes 'Application.Read.All','Directory.Read.All' -NoWelcome | Out-Null

# 2 ── Microsoft’s *global* tenant Id (constant)
$microsoftTenantId = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'   # NEVER changes

# 3 ── Fetch all service principals
Write-Progress -Activity 'Enterprise Apps' -Status 'Querying…' -PercentComplete 0
$all = Get-MgServicePrincipal -All

# 4 ── Filter: Enterprise + NOT Microsoft-owned (+ optional name filter)
$enterpriseApps = $all | Where-Object {
    $_.ServicePrincipalType -eq 'Application' -and
    $_.AppOwnerOrganizationId -ne $microsoftTenantId -and
    $_.DisplayName -notmatch '^Microsoft'                    # extra safety
}

# 5 ── Output
$enterpriseApps | Select-Object DisplayName, Id, AppId, Homepage, PublisherName, CreatedDateTime |
    Sort-Object DisplayName |
    Format-Table -AutoSize

Write-Host "`nEnterprise Applications found: $($enterpriseApps.Count)" -ForegroundColor Green

# 6 ── Optional CSV
$enterpriseApps | Export-Csv -Path "EnterpriseApps_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv" -NoTypeInformation

# 7 ── Disconnect
Disconnect-MgGraph -ErrorAction SilentlyContinue