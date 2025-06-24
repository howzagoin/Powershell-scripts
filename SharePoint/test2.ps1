$apiKey = "4KgveJYUOWZ5mio5lhR"
$authString = "$apiKey:X"
$encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
$headers = @{ Authorization = "Basic $encodedAuth" }
