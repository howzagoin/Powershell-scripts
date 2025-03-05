# Function to connect to SharePoint Online with MFA (Web login)
Function Connect-SharePointWithMFA {
    param (
        [string]$AdminUrl
    )
    try {
        Write-Output "Connecting to SharePoint Online at $AdminUrl..."

        # Use PnP PowerShell with Web Login (MFA)
        Connect-PnPOnline -Url $AdminUrl -UseWebLogin

        Write-Output "Successfully connected to SharePoint Online."
    }
    catch {
        Write-Error "Failed to connect to SharePoint Online. Please check your credentials and try again."
        throw
    }
}

# Function to find large files in SharePoint libraries
Function Find-LargeFiles {
    param (
        [string]$TenantAdminURL,
        [string]$CSVFilePath
    )

    # Delete the output report if it exists
    if (Test-Path $CSVFilePath) {
        Write-Output "Deleting existing output file at $CSVFilePath..."
        Remove-Item $CSVFilePath
    }

    # Check if the Temp directory exists and create if not
    $tempDirectory = [System.IO.Path]::GetDirectoryName($CSVFilePath)
    if (-not (Test-Path $tempDirectory)) {
        Write-Output "Creating directory $tempDirectory..."
        New-Item -ItemType Directory -Path $tempDirectory -Force
    }

    Write-Output "Connecting to SharePoint tenant admin URL: $TenantAdminURL..."
    $SiteCollections = Get-PnPTenantSite | Where-Object {
        $_.URL -like '*/sites*' -and $_.Template -NotIn @(
            "SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", 
            "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1"
        )
    }

    $ExcludedLists = @(
        "Form Templates", "Preservation Hold Library", "Site Assets", 
        "Pages", "Site Pages", "Images", "Site Collection Documents", 
        "Site Collection Images", "Style Library"
    )

    $SiteCounter = 1
    ForEach ($Site in $SiteCollections) {
        Write-Progress -Id 1 -Activity "Processing Site Collections" `
            -Status "Processing Site: $($Site.URL) ($SiteCounter of $($SiteCollections.Count))" `
            -PercentComplete (($SiteCounter / $SiteCollections.Count) * 100)

        Connect-SharePointWithMFA -AdminUrl $Site.URL

        $DocumentLibraries = Get-PnPList | Where-Object {
            $_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.ItemCount -gt 0
        }

        $ListCounter = 1
        ForEach ($List in $DocumentLibraries) {
            $global:counter = 0
            $FileData = @()

            Write-Progress -Id 2 -ParentId 1 -Activity "Processing Document Libraries" `
                -Status "Processing Library: $($List.Title) ($ListCounter of $($DocumentLibraries.Count))" `
                -PercentComplete (($ListCounter / $DocumentLibraries.Count) * 100)

            $Files = Get-PnPListItem -List $List -Fields FileLeafRef,FileRef,SMTotalFileStreamSize `
                -PageSize 500 -ScriptBlock {
                    Param($items)
                    $global:counter += $items.Count
                    if ($List.ItemCount -gt 0) {
                        Write-Progress -Id 3 -ParentId 2 `
                            -PercentComplete ($global:counter / ($List.ItemCount) * 100) `
                            -Activity "Getting List Items of '$($List.Title)'" `
                            -Status "Processing Items $global:counter to $($List.ItemCount)"
                    }
                } | Where-Object {
                    ($_.FileSystemObjectType -eq "File") `
                    -and ($_.FieldValues.SMTotalFileStreamSize / 1MB -gt 100)
                }

            ForEach ($File in $Files) {
                $FileData += [PSCustomObject][ordered]@{
                    Library = $List.Title
                    FileName = $File.FieldValues.FileLeafRef
                    URL = $File.FieldValues.FileRef
                    Size = [math]::Round(($File.FieldValues.SMTotalFileStreamSize / 1MB), 2)
                }
            }

            $FileData | Sort-Object Size -Descending | Export-Csv -Path $CSVFilePath `
                -NoTypeInformation -Append
            $ListCounter++
        }

        $SiteCounter++
    }

    Write-Output "Large file report saved to $CSVFilePath."
}

# Main Script Execution
Function Main {
    $TenantAdminURL = "https://company-admin.sharepoint.com"
    $CSVFilePath = "C:\Temp\LargeFiles-WGC.csv"

    # Connect to SharePoint Online with MFA (Web login)
    Connect-SharePointWithMFA -AdminUrl $TenantAdminURL

    # Find large files
    Find-LargeFiles -TenantAdminURL $TenantAdminURL -CSVFilePath $CSVFilePath
}

# Call the Main function to start the script
Main
