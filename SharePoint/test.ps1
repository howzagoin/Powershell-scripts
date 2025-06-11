Write-Host "Starting test"
Connect-MgGraph -ClientId '278b9af9-888d-4344-93bb-769bdd739249' -TenantId 'ca0711e2-e703-4f4e-9099-17d97863211c' -CertificateThumbprint 'B0AF0EF7659EA83D3140844F4BF89CCBB9413DBA' -NoWelcome -ErrorAction Stop
$ctx = Get-MgContext
Write-Host "AuthType: $($ctx.AuthType)"