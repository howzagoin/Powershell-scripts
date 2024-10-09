Connect-AzureAD

$allUsers = Get-AzureADUser -All $true

$activeUsers = $allUsers | Where-Object { $_.AccountEnabled -eq $true -and $_.UserPrincipalName -notlike "*#EXT#*" }

$activeUsers | Select-Object DisplayName, UserPrincipalName, AccountEnabled | Export-Csv -Path "C:\Temp\ActiveUsers.csv" -NoTypeInformation

Disconnect-AzureAD