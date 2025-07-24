#region metadata# Script Metadata
# Author: Tim MacLatchy
# Date: 17-07-2025
# License: MIT
# Description: Audits Microsoft 365 tenant users, mailboxes, groups, permissions, calendar sharing, enterprise apps, licenses, domains, and exports all results to a formatted Excel workbook.
# Steps:
#   1. Verify and install required modules
#   2. Authenticate to Microsoft Graph and Exchange Online
#   3. Retrieve users, mailboxes, groups, permissions, calendar, app, license, and domain data
#   4. Export all data to a formatted Excel file
#endregion

#region 1. Initialization & Modules
Write-Progress -Activity 'Initialization & Modules' -Status 'Loading...' -PercentComplete 0
$ErrorActionPreference = 'Stop'
$WarningPreference     = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms

$global:AuditStats = [ordered]@{
    UsersProcessed    = 0
    RulesProcessed    = 0
    ErrorsEncountered = 0
    StartTime         = Get-Date
}

$modules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Users',
    'Microsoft.Graph.Groups',
    'Microsoft.Graph.Identity.DirectoryManagement',
    'Microsoft.Graph.Applications',
    'ExchangeOnlineManagement',
    'ImportExcel'
)

foreach ($m in $modules) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        try {
            Install-Module -Name $m -Scope CurrentUser -Force -WarningAction SilentlyContinue
        } catch {
            Write-Warning ("Failed to install module " + $m + ": " + $_.Exception.Message)
        }
    }
    try {
        Import-Module $m -Force
    } catch {
        Write-Warning ("Failed to import module " + $m + ": " + $_.Exception.Message)
    }
}

# Ensure ExchangeOnlineManagement is available before continuing
if (-not (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement')) {
    Write-Host 'ERROR: ExchangeOnlineManagement module is not installed. Please install it and try again.' -ForegroundColor Red
    exit 1
}
Write-Progress -Activity 'Initialization & Modules' -Completed
#endregion

# For individual user audits, use TenantAudit-User.ps1 instead
Write-Host "Starting full tenant audit..." -ForegroundColor Cyan
$selectedUsers = @()
#endregion

#region Helper Functions
Write-Progress -Activity 'Helper Functions' -Status 'Defining...' -PercentComplete 0
function Get-UserDirectoryRoles {
    param(
        [Parameter(Mandatory)]
        [string]$UserId
    )
    $roleNames = @()
    # Get all directory roles in tenant
    $allRoles = Get-MgDirectoryRole -ErrorAction SilentlyContinue
    foreach ($role in $allRoles) {
        try {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -ErrorAction SilentlyContinue
            if ($members) {
                foreach ($member in $members) {
                    if ($member.Id -eq $UserId) {
                        $roleNames += $role.DisplayName
                    }
                }
            }
        } catch {}
    }
    if ($roleNames.Count -eq 0) { return 'None' }
    $sortedRoles = $roleNames | Sort-Object -Unique
    return ($sortedRoles -join '; ')
}
Write-Progress -Activity 'Helper Functions' -Completed
#endregion

#region 2. Choose Excel file path
Write-Progress -Activity 'Excel File Picker' -Status 'Selecting file...' -PercentComplete 0
Add-Type -AssemblyName System.Windows.Forms
$sd = New-Object System.Windows.Forms.SaveFileDialog
$sd.Title = 'Save Tenant Audit Report'
$sd.Filter = 'Excel Workbook (*.xlsx)|*.xlsx'
$sd.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
$sd.FileName = "M365_TenantAudit_{0:yyyyMMdd_HHmmss}.xlsx" -f (Get-Date)
if ($sd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host 'Audit cancelled.' -ForegroundColor Yellow
    return
}
$excelFile = $sd.FileName
Write-Progress -Activity 'Excel File Picker' -Completed
#endregion

#region 3. Connect to Graph & Exchange (one login)
Write-Progress -Activity 'Connect to Graph & Exchange' -Status 'Connecting...' -PercentComplete 0
Connect-MgGraph -Scopes @(
    'User.Read.All','Group.Read.All','Directory.Read.All',
    'AuditLog.Read.All','Mail.Read','MailboxSettings.Read',
    'Policy.Read.All','Application.Read.All'
) -NoWelcome
$loginForm = New-Object System.Windows.Forms.Form
$loginForm.TopMost = $true
$loginForm.Show()
$ctx = Get-MgContext
$exchangeConnected = $false
try {
    Connect-ExchangeOnline -UserPrincipalName $ctx.Account -ShowBanner:$false -ErrorAction Stop
    $exchangeConnected = $true
} catch {
    Write-Host 'ERROR: Exchange Online not connected. Mailbox data will be unavailable. Please check your credentials and network connection.' -ForegroundColor Red
    $loginForm.Close()
    exit 1
}
$loginForm.Close()
Write-Progress -Activity 'Connect to Graph & Exchange' -Completed
#endregion

#region 4. Users and Licenses
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Users & Licenses' -PercentComplete 10
Write-Host 'Retrieving licenses...'
$licenses    = Get-MgSubscribedSku -All -ErrorAction Stop
Write-Host 'Licenses retrieved.'
$userResults = @()
if ($selectedUsers.Count -gt 0) {
    Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status "Retrieving selected user(s)..." -PercentComplete 15
    $allUsers = @()
    foreach ($sel in $selectedUsers) {
        $user = Get-MgUser -Filter "UserPrincipalName eq '$sel' or Mail eq '$sel' or DisplayName eq '$sel'" -Property Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,UserType,JobTitle,MobilePhone,Department,Country,CreatedDateTime,Mail,MailNickname,OtherMails,ProxyAddresses -ErrorAction SilentlyContinue
        if ($user) { $allUsers += $user }
        else { Write-Host "User not found: $sel" -ForegroundColor Yellow }
    }
    if ($allUsers.Count -eq 0) {
        Write-Host "No valid users found. Exiting." -ForegroundColor Yellow
        exit 1
    }
    Write-Host "Selected users retrieved: $($allUsers.Count)"
} else {
    Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status "Retrieving all users..." -PercentComplete 15
    $allUsers = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,UserType,JobTitle,MobilePhone,Department,Country,CreatedDateTime,Mail,MailNickname,OtherMails,ProxyAddresses -ErrorAction Stop
    Write-Host "Users retrieved: $($allUsers.Count)"
}
foreach ($u in $allUsers) {
    $currentUserIndex = [array]::IndexOf($allUsers, $u) + 1
    $userPct = [math]::Round(($currentUserIndex / $allUsers.Count) * 100, 1)
    Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status "Scanning user $currentUserIndex of $($allUsers.Count): $($u.DisplayName) <$($u.UserPrincipalName)>" -PercentComplete $userPct
    $assigned = if ($u.AssignedLicenses) {
        ($u.AssignedLicenses | ForEach-Object {
            ($licenses | Where-Object SkuId -eq $_.SkuId).SkuPartNumber
        }) -join '; '
    } else { 'None' }
    $params = @{ UserId = $u.Id }
    $roles = Get-UserDirectoryRoles @params
    # ...existing code...
    $mb = Get-Mailbox -Identity $u.UserPrincipalName -ErrorAction SilentlyContinue
    if ($mb) {
        try {
            $mbStats = Get-EXOMailboxStatistics -Identity $u.UserPrincipalName -ErrorAction Stop |
                Select-Object TotalItemSize, ItemCount, LastLogonTime, LastLogoffTime, DisplayName
        } catch {
            Write-Host "Error retrieving mailbox statistics for $($u.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
            $global:AuditStats.ErrorsEncountered++
            $mbStats = $null
        }
    } else {
        $mbStats = $null
    }

    $mailboxSizeGB = 0
    if ($mbStats -and $mbStats.TotalItemSize) {
        $sizeString = $mbStats.TotalItemSize.ToString()
        if ($sizeString -match '\(([0-9,]+) bytes\)') {
            $bytesStr = $matches[1] -replace ',',''
            $mailboxSizeGB = [math]::Round([double]$bytesStr / 1GB, 2)
        }
    }

    $archiveSizeGB = 0
    if ($mb) {
        if ($mb.ArchiveStatus -eq 'Active') {
            try {
                $archiveStats = Get-EXOMailboxStatistics -Identity $u.UserPrincipalName -Archive -ErrorAction Stop |
                    Select-Object TotalItemSize, ItemCount, LastLogonTime, LastLogoffTime, DisplayName
            } catch {
                Write-Host "Error retrieving archive statistics for $($u.UserPrincipalName): $($_.Exception.Message)" -ForegroundColor Red
                $archiveStats = $null
            }
        } else {
            $archiveStats = $null
        }
    } else {
        $archiveStats = $null
    }

    if ($archiveStats -and $archiveStats.TotalItemSize) {
        $sizeString = $archiveStats.TotalItemSize.ToString()
        if ($sizeString -match '\(([0-9,]+) bytes\)') {
            $bytesStr = $matches[1] -replace ',',''
            $archiveSizeGB = [math]::Round([double]$bytesStr / 1GB, 2)
        }
    }
    $userResults += [PSCustomObject]@{
        DisplayName       = $u.DisplayName
        UserPrincipalName = $u.UserPrincipalName
        AccountEnabled    = $u.AccountEnabled
        UserType          = $u.UserType
        JobTitle          = $u.JobTitle
        MobilePhone       = $u.MobilePhone
        Department        = $u.Department
        Country           = $u.Country
        CreatedDate       = $u.CreatedDateTime
        AssignedLicenses  = $assigned
        UserRoles         = $roles
        MailboxType       = if ($mb) { $mb.RecipientTypeDetails } else { '' }
        MailboxSizeGB     = if ($mbStats -and $mbStats.TotalItemSize) {
            [double]$mailboxSizeGB
        } else {
            ''
        }
        ArchiveSizeGB     = if ($archiveStats -and $archiveStats.TotalItemSize) {
            [double]$archiveSizeGB
        } else {
            ''
        }
        Mail              = $u.Mail
        MailNickname      = $u.MailNickname
        OtherMails        = if ($u.OtherMails) { $u.OtherMails -join '; ' } else { '' }
        ProxyAddresses    = if ($u.ProxyAddresses -is [array]) {
            $u.ProxyAddresses -join '; '
        } elseif ($u.ProxyAddresses -is [string]) {
            $u.ProxyAddresses
        } elseif ($u.ProxyAddresses) {
            $u.ProxyAddresses.ToString()
        } else {
            ''
        }
    }
    $global:AuditStats.UsersProcessed++
}
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Users & Licenses completed' -PercentComplete 25
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Starting Groups & Members enumeration...' -PercentComplete 30
#endregion

#region 7. Groups & GroupMembers 
Write-Progress -Activity 'Groups & Members Enumeration' -Status 'Starting group enumeration...' -PercentComplete 0
$allGroups = $null
Write-Host 'Retrieving groups...'
$allGroups = Get-MgGroup -All -Property Id,DisplayName,Mail,Description,GroupTypes,ResourceProvisioningOptions,MailEnabled,SecurityEnabled,OnPremisesSyncEnabled -ErrorAction Stop
Write-Host "Groups retrieved: $($allGroups.Count)"
$groupResults = @()
$groupMemberResults = @()
$uniqueGroups = $allGroups | Group-Object Id | ForEach-Object { $_.Group[0] }
$totalGroups = $uniqueGroups.Count
for ($i = 0; $i -lt $totalGroups; $i++) {
    $g   = $uniqueGroups[$i]
    $pct = [math]::Round((($i+1)/$totalGroups)*100, 1)
    Write-Progress -Activity 'Groups & Members Enumeration' -Status "Scanning group $($i+1) of ${totalGroups}: $($g.DisplayName)" -PercentComplete $pct
    # Group type logic
    $gt = if ($g.ResourceProvisioningOptions -contains 'Team') {
        'Teams'
    } elseif ($g.GroupTypes -contains 'SharePoint') {
        'SharePoint'
    } elseif ($g.GroupTypes -contains 'DynamicMembership') {
        'Dynamic'
    } elseif ($g.MailEnabled -and -not $g.SecurityEnabled) {
        'Distribution'
    } elseif ($g.MailEnabled -and $g.SecurityEnabled) {
        'Mail-Enabled Security'
    } elseif ($g.GroupTypes -contains 'Unified') {
        'Microsoft 365'
    } elseif ($g.OnPremisesSyncEnabled) {
        'OnPrem AD'
    } elseif ($g.SecurityEnabled) {
        'Security'
    } else {
        'Other'
    }
    $members = @( )
    $owners = @( )
    try {
        $members = @( Get-MgGroupMember -GroupId $g.Id -All -ErrorAction SilentlyContinue )
    } catch {
        $members = @()
    }
    try {
        $owners  = @( Get-MgGroupOwner  -GroupId $g.Id -All -ErrorAction SilentlyContinue )
    } catch {
        $owners = @()
    }
    $ownerNames = $owners | ForEach-Object { $_.AdditionalProperties['displayName'] ?? $_.DisplayName }
    $groupResults += [PSCustomObject]@{
        GroupName        = $g.DisplayName
        GroupType        = $gt
        EmailAddress     = $g.Mail
        GroupDescription = $g.Description
        MemberCount      = $members.Count
        OwnerCount       = $owners.Count
        OwnerNames       = $ownerNames -join '; '
    }
    foreach ($m in $members) {
        $dn = $m.AdditionalProperties['displayName'] ?? $m.DisplayName
        $upn= $m.AdditionalProperties['userPrincipalName'] ?? $m.UserPrincipalName
        $groupMemberResults += [PSCustomObject]@{
            GroupName  = $g.DisplayName
            GroupType  = $gt
            MemberName = $dn
            MemberUPN  = $upn
            MemberType = 'Member'
        }
    }
    foreach ($o in $owners) {
        $dn = $o.AdditionalProperties['displayName'] ?? $o.DisplayName
        $upn= $o.AdditionalProperties['userPrincipalName'] ?? $o.UserPrincipalName
        $groupMemberResults += [PSCustomObject]@{
            GroupName  = $g.DisplayName
            GroupType  = $gt
            MemberName = $dn
            MemberUPN  = $upn
            MemberType = 'Owner'
        }
    }
}
Write-Progress -Activity 'Groups & Members Enumeration' -Status 'Completed group and member enumeration.' -PercentComplete 100
#endregion

#region 8. Mailbox Rules 
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Starting Mailbox Rules...' -PercentComplete 45
$mailboxRules = @()
if ($exchangeConnected) {
     for ($i = 0; $i -lt $allUsers.Count; $i++) {
     $u = $allUsers[$i]
     $mbRulePct = [math]::Round((($i+1)/$allUsers.Count)*100, 1)
     Write-Progress -Activity 'Mailbox Rules' -Status "Scanning mailbox rules for user $($i+1) of $($allUsers.Count): $($u.DisplayName) <$($u.UserPrincipalName)>" -PercentComplete $mbRulePct
     try {
         $rules = Get-InboxRule -Mailbox $u.UserPrincipalName -ErrorAction SilentlyContinue
     } catch {
         $rules = $null
     }
     try {
         $oof = Get-MailboxAutoReplyConfiguration -Identity $u.UserPrincipalName -ErrorAction SilentlyContinue
     } catch {
         $oof = $null
     }
     if ($rules) {
         foreach ($rule in $rules) {
             $mailboxRules += [PSCustomObject]@{
                 MailboxOwner = $u.UserPrincipalName
                 RuleName     = $rule.Name
                 Enabled      = $rule.Enabled
                 ForwardTo    = if ($rule.ForwardTo) { ($rule.ForwardTo | ForEach-Object { $_.ToString() }) -join '; ' } else { '' }
                 RedirectTo   = if ($rule.RedirectTo) { ($rule.RedirectTo | ForEach-Object { $_.ToString() }) -join '; ' } else { '' }
                 Description  = $rule.Description
                 Priority     = $rule.Priority
                 From         = if ($rule.From) { ($rule.From | ForEach-Object { $_.ToString() }) -join '; ' } else { '' }
                 SentTo       = if ($rule.SentTo) { ($rule.SentTo | ForEach-Object { $_.ToString() }) -join '; ' } else { '' }
                 Conditions   = if ($rule.Conditions) { ($rule.Conditions | ForEach-Object { $_.ToString() }) -join '; ' } else { '' }
                 Actions      = if ($rule.Actions) { ($rule.Actions | ForEach-Object { $_.ToString() }) -join '; ' } else { '' }
             }
         }
     }
     if ($oof -and $oof.AutomaticRepliesEnabled) {
         $mailboxRules += [PSCustomObject]@{
             MailboxOwner = $u.UserPrincipalName
             RuleType     = 'OutOfOffice'
             OOFMessage   = $oof.ReplyMessage
             OOFStartTime = $oof.StartTime
             OOFEndTime   = $oof.EndTime
         }
     }
     }
     $global:AuditStats.RulesProcessed = $mailboxRules.Count
     Write-Progress -Activity 'Mailbox Rules' -Status 'Completed' -PercentComplete 100
 } else {
     Write-Warning "Skipping mailbox rules: Exchange Online not connected."
 }
Write-Progress -Activity 'Mailbox Rules' -Completed
#endregion

#region 9. DelegatedMailboxes
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Processing Delegated Mailboxes...' -PercentComplete 60
$delegatedMailboxes = @()
 for ($i = 0; $i -lt $allUsers.Count; $i++) {
     $u = $allUsers[$i]
     $delPct = [math]::Round((($i+1)/$allUsers.Count)*100, 1)
     Write-Progress -Activity 'DelegatedMailboxes' -Status "Scanning delegated mailboxes for user $($i+1) of $($allUsers.Count): $($u.DisplayName) <$($u.UserPrincipalName)>" -PercentComplete $delPct
     $mb = Get-Mailbox -Identity $u.UserPrincipalName -ErrorAction SilentlyContinue
     if ($mb) {
         Get-MailboxPermission -Identity $u.UserPrincipalName -ErrorAction SilentlyContinue |
             Where-Object { $_.User -notlike 'NT AUTHORITY*' -and $_.User -notlike 'S-1-*' } | ForEach-Object {
                 $delegatedMailboxes += [PSCustomObject]@{
                     MailboxOwner   = $u.UserPrincipalName
                     DelegateUser   = $_.User
                     MailboxType    = $mb.RecipientTypeDetails
                     AccessRights   = ($_.AccessRights -join ', ')
                     DelegationType = $(if ($_.IsInherited) { 'Inherited' } else { 'Direct' })
                 }
             }
     }
 }
Write-Progress -Activity 'DelegatedMailboxes' -Completed
#endregion

#region 10. Calendars (only shared with others)
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Processing Calendar Permissions...' -PercentComplete 70
$calendars = @()
 for ($i = 0; $i -lt $allUsers.Count; $i++) {
    $u = $allUsers[$i]
    $calPct = [math]::Round((($i+1)/$allUsers.Count)*100, 1)
    Write-Progress -Activity 'Calendars' -Status "Scanning calendars for user $($i+1) of $($allUsers.Count): $($u.DisplayName) <$($u.UserPrincipalName)>" -PercentComplete $calPct
    $mb = Get-Mailbox -Identity $u.UserPrincipalName -ErrorAction SilentlyContinue
    if ($mb) {
        try {
            $folders = Get-MailboxFolderStatistics -Identity $u.UserPrincipalName -ErrorAction SilentlyContinue
            $calendarFolders = $folders | Where-Object { $_.FolderPath -match "^/Calendar" -or $_.FolderType -like '*Calendar*' }
            foreach ($folder in $calendarFolders) {
                $cleanPath = $folder.FolderPath.TrimStart("/").Replace("/", "\")
                if ($cleanPath -eq "\\") { $cleanPath = "Calendar" }
                # Skip problematic folders
                if ($folder.Name -eq "Calendar Logging" -or [string]::IsNullOrWhiteSpace($cleanPath)) { continue }
                $folderIdentity = "$($u.UserPrincipalName):\$cleanPath"
                try {
                    $perms = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
                    if ($perms) {
                        $filteredPermissions = $perms | Where-Object { 
                            # Include permissions that are not Default/Anonymous AND have actual rights
                            $_.User.DisplayName -notin @('Default','Anonymous') -and 
                            ($_.AccessRights | Where-Object { $_ -notin @('None','AvailabilityOnly') }).Count -gt 0
                        }
                        if ($filteredPermissions.Count -gt 0) {
                            foreach ($perm in $filteredPermissions) {
                                $calendars += [PSCustomObject]@{
                                    CalendarOwner   = $u.UserPrincipalName
                                    CalendarName    = $folder.Name
                                    SharedWithUser  = $perm.User.ToString()
                                    AccessRights    = ($perm.AccessRights -join ', ')
                                    MailboxType     = $mb.RecipientTypeDetails
                                    DelegationType  = if ($perm.SharingPermissionFlags) { 'Shared' } else { 'Delegated' }
                                    FolderPath      = $cleanPath
                                    ItemCount       = $folder.ItemsInFolder
                                }
                            }
                        }
                    }
                } catch {}
            }
        } catch {}
    }
}
Write-Progress -Activity 'Calendars' -Completed
#endregion

#region 11. Enterprise Applications (non-Microsoft)
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Processing Enterprise Applications...' -PercentComplete 80
### Enterprise Apps tab: sort by DisplayName, output each assigned user on a separate row, and show app owners
Write-Progress -Activity 'Enterprise Apps' -Status 'Querying service principals...' -PercentComplete 0
$microsoftTenantId = 'f8cdef31-a31e-4b4a-93e4-5f571e91255a'
$allSp = Get-MgServicePrincipal -All
$filteredApps = $allSp | Where-Object {
    $_.ServicePrincipalType -eq 'Application' -and
    $_.AppOwnerOrganizationId -ne $microsoftTenantId -and
    $_.DisplayName -notmatch '^Microsoft'
}
$appResults = @()
$sortedApps = $filteredApps | Sort-Object DisplayName
for ($i = 0; $i -lt $sortedApps.Count; $i++) {
    $app = $sortedApps[$i]
    $pct = [math]::Round((($i+1)/$sortedApps.Count)*100, 1)
    Write-Progress -Activity 'Enterprise Apps' -Status "Processing app $($i+1) of $($sortedApps.Count): $($app.DisplayName)" -PercentComplete $pct
    $assignedUsers = @()
    $owners = @()
    try {
        $assignedUsers = Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $app.Id -ErrorAction SilentlyContinue | ForEach-Object {
            $_.PrincipalDisplayName
        }
    } catch {}
    try {
        $owners = Get-MgServicePrincipalOwner -ServicePrincipalId $app.Id -ErrorAction SilentlyContinue | ForEach-Object {
            $_.DisplayName
        }
    } catch {}
    if ($assignedUsers.Count -eq 0) {
        $appResults += [PSCustomObject]@{
            DisplayName   = $app.DisplayName
            Homepage      = $app.Homepage
            LoginUrl      = $app.LoginUrl
            LogoutUrl     = $app.LogoutUrl
            AssignedUser  = ''
            AppOwner      = ($owners -join '; ')
        }
    } else {
        foreach ($user in $assignedUsers) {
            $appResults += [PSCustomObject]@{
                DisplayName   = $app.DisplayName
                Homepage      = $app.Homepage
                LoginUrl      = $app.LoginUrl
                LogoutUrl     = $app.LogoutUrl
                AssignedUser  = $user
                AppOwner      = ($owners -join '; ')
            }
        }
    }
}
Write-Progress -Activity 'Enterprise Apps' -Status 'Completed' -PercentComplete 100
Write-Progress -Activity 'Enterprise Apps' -Completed
#endregion

#region 12. Licenses, Domains, SummaryNotes
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Processing Licenses & Domains...' -PercentComplete 90
$licenseSummary = @()
for ($i = 0; $i -lt $licenses.Count; $i++) {
    $lic = $licenses[$i]
    $pct = [math]::Round((($i+1)/$licenses.Count)*100, 1)
    Write-Progress -Activity 'Licenses' -Status "Processing license $($i+1) of $($licenses.Count): $($lic.SkuPartNumber)" -PercentComplete $pct
    $used = $lic.ConsumedUnits
    $total = $lic.PrepaidUnits.Enabled
    $free = $total - $used
    $licenseSummary += [PSCustomObject]@{
        SkuPartNumber = $lic.SkuPartNumber
        SkuName       = $lic.SkuPartNumber
        Used          = $used
        Free          = $free
        Total         = $total
    }
}
Write-Progress -Activity 'Licenses' -Status 'Completed' -PercentComplete 100

Write-Progress -Activity 'Domains' -Status 'Starting...' -PercentComplete 0
$allDomains = Get-MgDomain -All
$totalDomains = $allDomains.Count
$domainResults = @()
for ($i = 0; $i -lt $allDomains.Count; $i++) {
    $d = $allDomains[$i]
    $pct = [math]::Round((($i+1)/$allDomains.Count)*100, 1)
    Write-Progress -Activity 'Domains' -Status "Processing domain $($i+1) of $($allDomains.Count): $($d.Id)" -PercentComplete $pct
    $domainResults += [PSCustomObject]@{
        DomainName  = $d.Id
        IsVerified  = $d.IsVerified
        IsDefault   = $d.IsDefault
    }
}
Write-Progress -Activity 'Domains' -Status 'Completed' -PercentComplete 100

Write-Progress -Activity 'SummaryNotes' -Status 'Starting...' -PercentComplete 0
$endTime = Get-Date
$duration = $endTime - $global:AuditStats.StartTime
$summaryNotes = @()
$summaryNotes += [PSCustomObject]@{ Note = "Audit completed on $($endTime.ToString('yyyy-MM-dd HH:mm:ss')) by $($ctx.Account)" }
$summaryNotes += [PSCustomObject]@{ Note = "Start Time: $($global:AuditStats.StartTime)" }
$summaryNotes += [PSCustomObject]@{ Note = "End Time: $($endTime)" }
$summaryNotes += [PSCustomObject]@{ Note = "Duration: $($duration.ToString('hh\:mm\:ss'))" }
$summaryNotes += [PSCustomObject]@{ Note = "Users processed: $($global:AuditStats.UsersProcessed)" }
$summaryNotes += [PSCustomObject]@{ Note = "Rules processed: $($global:AuditStats.RulesProcessed)" }
$summaryNotes += [PSCustomObject]@{ Note = "Errors encountered: $($global:AuditStats.ErrorsEncountered)" }
$summaryNotes += [PSCustomObject]@{ Note = "User+Mailbox count: $($userResults.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "Group count: $($groupResults.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "GroupMembers count: $($groupMemberResults.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "MailboxRules count: $($mailboxRules.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "DelegatedMailboxes count: $($delegatedMailboxes.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "Calendars count: $($calendars.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "EnterpriseApps count: $($appResults.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "Licenses count: $($licenseSummary.Count)" }
$summaryNotes += [PSCustomObject]@{ Note = "Domains count: $($domainResults.Count)" }
Write-Progress -Activity 'SummaryNotes' -Status 'Completed' -PercentComplete 100
#endregion

#region 13. Export to Excel & Formatting

Write-Host "Exporting results to Excel..."
Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Status 'Exporting to Excel...' -PercentComplete 95

$pkg = Open-ExcelPackage -Path $excelFile -Create

if ($selectedUsers.Count -eq 1 -and $allUsers.Count -eq 1) {
    # Single user: compile all info into one worksheet
    $singleUser = $allUsers[0]
    $singleTabData = @()
    $singleTabData += [PSCustomObject]@{ Section = 'User'; Data = $userResults | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'Groups'; Data = $groupResults | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'GroupMembers'; Data = $groupMemberResults | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'MailboxRules'; Data = $mailboxRules | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'DelegatedMailboxes'; Data = $delegatedMailboxes | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'Calendars'; Data = $calendars | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'EnterpriseApps'; Data = $appResults | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'Licenses'; Data = $licenseSummary | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'Domains'; Data = $domainResults | ConvertTo-Json -Compress }
    $singleTabData += [PSCustomObject]@{ Section = 'SummaryNotes'; Data = $summaryNotes | ConvertTo-Json -Compress }
    $pkg = $singleTabData | Export-Excel -ExcelPackage $pkg -WorksheetName "User_$($singleUser.UserPrincipalName)" -TableStyle "Medium1" -AutoSize -AutoFilter -BoldTopRow -PassThru
    Close-ExcelPackage $pkg -Show
    Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Completed
    return
}


# If multiple users selected in option 2, create a worksheet per user with only their data
if ($selectedUsers.Count -gt 1 -and $allUsers.Count -gt 1) {
    $worksheetCounter = 1
    foreach ($u in $allUsers) {
        $userUpn = $u.UserPrincipalName
        $userTabData = [ordered]@{
            'User'              = $userResults | Where-Object { $_.UserPrincipalName -eq $userUpn }
            'Groups'            = $groupResults | Where-Object { ($groupMemberResults | Where-Object { $_.MemberUPN -eq $userUpn }).GroupName -contains $_.GroupName }
            'GroupMembers'      = $groupMemberResults | Where-Object { $_.MemberUPN -eq $userUpn }
            'MailboxRules'      = $mailboxRules | Where-Object { $_.MailboxOwner -eq $userUpn }
            'DelegatedMailboxes'= $delegatedMailboxes | Where-Object { $_.MailboxOwner -eq $userUpn -or $_.DelegateUser -eq $userUpn }
            'Calendars'         = $calendars | Where-Object { $_.CalendarOwner -eq $userUpn }
            'EnterpriseApps'    = $appResults | Where-Object { $_.AssignedUser -eq $userUpn }
            'Licenses'          = $licenseSummary # Licenses are tenant-wide, include all
            'Domains'           = $domainResults  # Domains are tenant-wide, include all
            'SummaryNotes'      = $summaryNotes   # Summary is tenant-wide, include all
        }
        # Create a copy of keys before iteration
        $tabKeys = @($userTabData.Keys)
        foreach ($tab in $tabKeys) {
            if (-not $userTabData[$tab] -or $userTabData[$tab].Count -eq 0) {
                $userTabData[$tab] = @([PSCustomObject]@{ Info = 'No data found' })
            }
        }
        # Generate a short, unique identifier for the user
        $shortName = $u.DisplayName.Split(' ')[0] + $worksheetCounter.ToString()
        
        # Process each data section one at a time
        foreach ($entry in $userTabData.GetEnumerator()) {
            # Keep worksheet names short and simple
            $wsName = "$shortName-$($entry.Key)"
            if ($wsName.Length -gt 31) { # Excel worksheet name length limit
                $wsName = $wsName.Substring(0, 31)
            }
            
            # Export the data to Excel
            $pkg = $entry.Value | Export-Excel -ExcelPackage $pkg -WorksheetName $wsName `
                -TableStyle ("Medium$($worksheetCounter % 21 + 1)") -AutoSize -AutoFilter -BoldTopRow -PassThru
        }
        $worksheetCounter++
    }
    Close-ExcelPackage $pkg -Show
    Write-Progress -Activity 'Processing Microsoft 365 Tenant' -Completed
    return
}

# Usual multi-tab export for all/multiple users (all users or option 1)
# Define the data sets and export order
$orderedDataSets = [ordered]@{
    'Users'                   = $userResults
    'Groups'                  = $groupResults
    'GroupMembers'            = $groupMemberResults
    'MailboxRules'            = $mailboxRules
    'DelegatedMailboxes'      = $delegatedMailboxes
    'Calendars'               = $calendars
    'EnterpriseApps'          = $appResults
    'Licenses'                = $licenseSummary
    'Domains'                 = $domainResults
    'SummaryNotes'            = $summaryNotes
}

# Add 'No data found' to any empty tab
$dataSetKeys = @($orderedDataSets.Keys)
foreach ($k in $dataSetKeys) {
    if (-not $orderedDataSets[$k] -or $orderedDataSets[$k].Count -eq 0) {
        $orderedDataSets[$k] = @([PSCustomObject]@{ Info = 'No data found' })
    }
}

# Export each worksheet
$worksheetCounter = 1
foreach ($name in $orderedDataSets.Keys) {
    $pct = [math]::Round((($worksheetCounter)/$orderedDataSets.Keys.Count)*100, 1)
    Write-Progress -Activity 'Export to Excel & Formatting' -Status "Exporting worksheet $worksheetCounter of $($orderedDataSets.Keys.Count): $name" -PercentComplete $pct
    
    # Create a new worksheet for each data set
    try {
        $pkg = $orderedDataSets[$name] | Export-Excel -ExcelPackage $pkg -WorksheetName $name `
            -TableStyle "Medium$($worksheetCounter % 21 + 1)" -AutoSize -AutoFilter -BoldTopRow -PassThru
    }
    catch {
        Write-Warning "Error exporting worksheet '$name': $($_.Exception.Message)"
        continue
    }
    $worksheetCounter++
}

# Custom formatting for Groups worksheet: remove duplicate group type headings and extra rows
if ($pkg.Workbook.Worksheets['Groups']) {
    $wsG = $pkg.Workbook.Worksheets['Groups']
    $wsG.Cells.AutoFitColumns()
}

# Highlight owners in GroupMembers
if ($pkg.Workbook.Worksheets['GroupMembers']) {
    $wsGM = $pkg.Workbook.Worksheets['GroupMembers']
    for ($i = 2; $i -le $wsGM.Dimension.End.Row; $i++) {
        $isOwner = $wsGM.Cells[$i, 5].Text -eq 'True'
        $color = if ($isOwner) { [System.Drawing.Color]::LightYellow } else { [System.Drawing.Color]::White }
        $wsGM.Row($i).Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsGM.Row($i).Style.Fill.BackgroundColor.SetColor($color)
    }
}

# Format mailbox sizes in Users sheet (where mailbox data is stored)
if ($pkg.Workbook.Worksheets['Users']) {
    $wsU = $pkg.Workbook.Worksheets['Users']
    
    # Find the column indices for MailboxSizeGB and ArchiveSizeGB
    $columnHeaders = 1..$wsU.Dimension.End.Column | ForEach-Object {
        [PSCustomObject]@{
            Index = $_
            Name = $wsU.Cells[1,$_].Text
        }
    }
    
    $mailboxSizeCol = ($columnHeaders | Where-Object { $_.Name -eq 'MailboxSizeGB' }).Index
    $archiveSizeCol = ($columnHeaders | Where-Object { $_.Name -eq 'ArchiveSizeGB' }).Index
    
    if ($mailboxSizeCol) {
        Write-Host "Formatting mailbox size column $mailboxSizeCol"
        $wsU.Column($mailboxSizeCol).Style.Numberformat.Format = "#,##0.00"
        # Ensure numeric values
        2..$wsU.Dimension.End.Row | ForEach-Object {
            $cell = $wsU.Cells[$_, $mailboxSizeCol]
            $originalValue = $cell.Value
            if ($null -ne $originalValue -and $originalValue -ne '') {
                try {
                    $numericValue = [double]$originalValue
                    $cell.Value = $numericValue
                    Write-Host "Row ${_}: Converted $originalValue to $numericValue"
                } catch {
                    Write-Warning "Failed to convert value '$originalValue' to number in row ${_}"
                }
            }
        }
    }
    if ($archiveSizeCol) {
        Write-Host "Formatting archive size column $archiveSizeCol"
        $wsU.Column($archiveSizeCol).Style.Numberformat.Format = "#,##0.00"
        # Ensure numeric values
        2..$wsU.Dimension.End.Row | ForEach-Object {
            $cell = $wsU.Cells[$_, $archiveSizeCol]
            $originalValue = $cell.Value
            if ($null -ne $originalValue -and $originalValue -ne '') {
                try {
                    $numericValue = [double]$originalValue
                    $cell.Value = $numericValue
                    Write-Host "Row ${_}: Converted $originalValue to $numericValue"
                } catch {
                    Write-Warning "Failed to convert value '$originalValue' to number in row ${_}"
                }
            }
        }
    }
    
    $wsU.Cells.AutoFitColumns()
}

Close-ExcelPackage $pkg -Show
Write-Progress -Activity 'Export to Excel & Formatting' -Status 'Completed' -PercentComplete 100
#endregion

#region 14. Cleanup
Write-Progress -Activity 'Cleanup' -Status 'Cleaning up...' -PercentComplete 0
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
} catch {
    # Ignore cleanup errors
}
Write-Progress -Activity 'Cleanup' -Completed

# Print summary after cleanup
# Detailed audit summary output
$endTime = Get-Date
$duration = $endTime - $global:AuditStats.StartTime
Write-Host ("`nAudit complete. Results saved to $excelFile")
Write-Host ("Duration: $($duration.ToString('hh\:mm\:ss'))")
Write-Host ("Users processed: $($global:AuditStats.UsersProcessed)")
Write-Host ("Rules processed: $($global:AuditStats.RulesProcessed)")
Write-Host ("Errors encountered: $($global:AuditStats.ErrorsEncountered)")
Write-Host ("User+Mailbox count: $($userResults.Count)")
Write-Host ("Group count: $($groupResults.Count)")
Write-Host ("GroupMembers count: $($groupMemberResults.Count)")
Write-Host ("MailboxRules count: $($mailboxRules.Count)")
Write-Host ("DelegatedMailboxes count: $($delegatedMailboxes.Count)")
Write-Host ("Calendars count: $($calendars.Count)")
Write-Host ("EnterpriseApps count: $($appResults.Count)")
Write-Host ("Licenses count: $($licenseSummary.Count)")
Write-Host ("Domains count: $($domainResults.Count)")
#endregion