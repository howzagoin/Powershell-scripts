$accessToken = (Get-MgContext).AccessToken
$headers = @{
    Authorization = "Bearer $accessToken"
}
$endpoint = "https://graph.microsoft.com/v1.0/deviceManagement/securityRecommendations"

try {
    $response = Invoke-RestMethod -Uri $endpoint -Method GET -Headers $headers
    Write-Host "API Response:" -ForegroundColor Green
    $response
}
catch {
    Write-Error "Error calling Microsoft Graph API: $_"
}
