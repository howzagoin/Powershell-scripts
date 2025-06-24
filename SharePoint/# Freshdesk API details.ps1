# Freshdesk API key and domain
$apiKey = "4KgveJYUOWZ5mio5lhR"
$freshdeskDomain = "itsupport-journebrands.freshdesk.com"

# Encode the API key for authorization
$authString = "$apiKey:X"
$encodedAuth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($authString))
$headers = @{ Authorization = "Basic $encodedAuth" }

# Function to retrieve all deleted contacts with pagination
function Get-DeletedContacts {
    $contacts = @()
    $page = 1
    do {
        $uri = "https://$freshdeskDomain/api/v2/contacts?state=deleted&page=$page&per_page=100"
        try {
            $response = Invoke-RestMethod -Method GET -Uri $uri -Headers $headers
            if ($response) {
                $contacts += $response
                $page++
            } else {
                break
            }
        } catch {
            Write-Warning "Failed to retrieve contacts on page ${page}: $_"
            break
        }
    } while ($response.Count -eq 100)
    return $contacts
}

# Function to perform a hard delete on a contact by ID
function HardDelete-ContactById {
    param ([int]$ContactId)
    $uri = "https://$freshdeskDomain/api/v2/contacts/$ContactId/hard_delete"
    try {
        Invoke-RestMethod -Method DELETE -Uri $uri -Headers $headers
        Write-Host "Permanently deleted contact ID $ContactId"
    } catch {
        Write-Warning "Failed to hard delete contact ID ${ContactId}: $_"
    }
}

# Main script execution
$allDeletedContacts = Get-DeletedContacts
Write-Host "Total deleted contacts to permanently delete: $($allDeletedContacts.Count)"

foreach ($contact in $allDeletedContacts) {
    HardDelete-ContactById -ContactId $contact.id
}
