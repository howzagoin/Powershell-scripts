#Config Variables
$TenantAdminURL = "https://company-admin.sharepoint.com/"
$CSVFilePath = "C:\Temp\LargeFiles-WGC.csv"
  
#Connect to Admin Center using PnP Online
Connect-PnPOnline -Url $TenantAdminURL -Interactive
 
#Delete the Output Report, if exists
if (Test-Path $CSVFilePath) { Remove-Item $CSVFilePath }
 
#Get All Site collections - Exclude: Seach Center, Redirect site, Mysite Host, App Catalog, Content Type Hub, eDiscovery and Bot Sites
$SiteCollections = Get-PnPTenantSite | Where { $_.URL -like '*/sites*' -and $_.Template -NotIn ("SRCHCEN#0", "REDIRECTSITE#0", "SPSMSITEHOST#0", "APPCATALOG#0", "POINTPUBLISHINGHUB#0", "EDISC#0", "STS#-1")}
 
#Get All Large Lists from the Web - Exclude Hidden and certain lists
$ExcludedLists = @("Form Templates", "Preservation Hold Library","Site Assets", "Pages", "Site Pages", "Images",
                        "Site Collection Documents", "Site Collection Images","Style Library")
 
$SiteCounter = 1  
#Loop through each site collection
ForEach($Site in $SiteCollections)
{   
    #Display a Progress bar
    Write-Progress -id 1 -Activity "Processing Site Collections" -Status "Processing Site: $($Site.URL)' ($SiteCounter of $($SiteCollections.Count))" -PercentComplete (($SiteCounter / $SiteCollections.Count) * 100)
  
    #Connect to the site
    Connect-PnPOnline -Url $Site.URL -Interactive
 
    #Get all document libraries
    $DocumentLibraries = Get-PnPList | Where-Object {$_.BaseType -eq "DocumentLibrary" -and $_.Hidden -eq $False -and $_.Title -notin $ExcludedLists -and $_.ItemCount -gt 0}
 
    $ListCounter = 1
    #Iterate through document libraries
    ForEach ($List in $DocumentLibraries)
    {
        $global:counter = 0
        $FileData = @()
 
        Write-Progress -id 2 -ParentId 1 -Activity "Processing Document Libraries" -Status "Processing Document Library: $($List.Title)' ($ListCounter of $($DocumentLibraries.Count))" -PercentComplete (($ListCounter / $DocumentLibraries.Count) * 100)
 
        #Get All Files of the library with size > 100MB
        $Files = Get-PnPListItem -List $List -Fields FileLeafRef,FileRef,SMTotalFileStreamSize -PageSize 500 -ScriptBlock { Param($items) $global:counter += $items.Count; Write-Progress -Id 3 -parentId 2 -PercentComplete ($global:Counter / ($List.ItemCount) * 100) -Activity "Getting List Items of '$($List.Title)'" -Status "Processing Items $global:Counter to $($List.ItemCount)";} | Where {($_.FileSystemObjectType -eq "File") -and ($_.FieldValues.SMTotalFileStreamSize/1MB -gt 100)}
 
        #Collect data from each files
        ForEach ($File in $Files)
        {
            $FileData += [PSCustomObject][ordered]@{
                Library      = $List.Title
                FileName  = $File.FieldValues.FileLeafRef
                URL            = $File.FieldValues.FileRef
                Size            = [math]::Round(($File.FieldValues.SMTotalFileStreamSize/1MB),2)
            }
        }
 
        #Export Files data to CSV File
        $FileData | Sort-object Size -Descending
        $FileData | Export-Csv -Path $CSVFilePath -NoTypeInformation -Append
        $ListCounter++
        #Write-Progress -Activity "Completed Processing List $($List.Title)" -Completed -id 2
 
    }
    $SiteCounter++
}