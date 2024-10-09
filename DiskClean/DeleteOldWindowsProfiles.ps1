# Delete old User Profiles Script

# List of accounts for which profiles must not be deleted
$ExcludedUsers = "Public", "Default", "witadmin"
$RunOnServers = $false
[int]$MaximumProfileAge = 30 # Profiles older than this will be deleted

# Get the current logged-in user's username
$currentUser = (Get-WmiObject -Class Win32_ComputerSystem).UserName.Split("\")[-1]

# Add the current user to the exclusion list
$ExcludedUsers += $currentUser

$osInfo = Get-CimInstance -ClassName Win32_OperatingSystem

if ($RunOnServers -eq $true -or $osInfo.ProductType -eq 1) {

    $profiles = Get-WMIObject -Class Win32_UserProfile | Where-Object {
        (!$_.Special -and $_.Loaded -eq $false) # Only consider non-special, unloaded profiles
    }

    # Output array to hold profile details for logging purposes and listing
    $profilesToDelete = @()

    # Initialize progress bar
    $totalProfiles = $profiles.Count
    $counter = 0

    foreach ($profile in $profiles) {
        $counter++

        $username = $profile.LocalPath.Replace("C:\Users\", "") # Extract username from the local path

        # Skip AzureAD profiles or if the username is in the exclusion list
        if (!($ExcludedUsers -contains $username) -and $username -notmatch "^AzureAD\\") {
            try {
                # Get the LastWriteTime of the UsrClass.dat file
                $lastWriteTime = (Get-ChildItem -Path "$($profile.LocalPath)\AppData\Local\Microsoft\Windows\UsrClass.dat" -Force).LastWriteTime

                # Check if the profile is older than the maximum allowed age
                if ($lastWriteTime -lt (Get-Date).AddDays(-$MaximumProfileAge)) {
                    # Add the profile to the list of profiles to delete
                    $profilesToDelete += [PSCustomObject]@{
                        SID           = $profile.SID
                        LastUseTime   = $profile.LastUseTime
                        LastWriteTime = $lastWriteTime
                        LocalPath     = $profile.LocalPath
                        Username      = $username
                    }
                }
            } catch {
                # Error handling
                Write-Host "Error occurred while processing profile for $username: $_" -ForegroundColor Red
            }
        }
    }

    # Display the profiles that will be deleted
    if ($profilesToDelete.Count -gt 0) {
        Write-Host "The following profiles are scheduled for deletion:" -ForegroundColor Cyan
        $profilesToDelete | Sort-Object LocalPath | Format-Table

        # Prompt the user for confirmation before deleting
        $confirmation = Read-Host "Do you want to proceed with the deletion? (Yes/No)"

        if ($confirmation -eq "Yes") {
            # Proceed with deletion
            $counter = 0

            foreach ($profile in $profilesToDelete) {
                $counter++

                # Update the progress bar for each profile being processed
                Write-Progress -Activity "Deleting old profiles" -Status "Processing profile $counter of $($profilesToDelete.Count)" -PercentComplete (($counter / $profilesToDelete.Count) * 100)

                # Remove the user profile
                try {
                    $profile | Remove-WmiObject
                    Write-Host "Profile for $($profile.Username) removed. LastWriteTime: $($profile.LastWriteTime)" -ForegroundColor Green
                } catch {
                    Write-Host "Error occurred while deleting profile for $($profile.Username): $_" -ForegroundColor Red
                }
            }

            # Complete the progress bar
            Write-Progress -Activity "Deleting old profiles" -Status "Completed" -PercentComplete 100 -Completed

            Write-Host "Profile deletion completed." -ForegroundColor Green
        } else {
            Write-Host "Profile deletion canceled." -ForegroundColor Yellow
        }
    } else {
        Write-Host "No profiles found that meet the deletion criteria." -ForegroundColor Cyan
    }
} else {
    Write-Host "This script is not configured to run on non-server operating systems. Exiting." -ForegroundColor Red
}
