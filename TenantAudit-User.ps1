
# TenantAudit-User.ps1
# Standalone script for auditing selected user(s) in Microsoft 365 tenant
# Usage: .\TenantAudit-User.ps1 -UserPrincipalNames user1@domain.com,user2@domain.com [-ExcelFile output.xlsx]

param(
    [Parameter(Mandatory)]
    [string[]]$UserPrincipalNames,
    [string]$ExcelFile = "UserAuditReport.xlsx"
)

# Import required modules
Import-Module ImportExcel -ErrorAction Stop
Import-Module Microsoft.Graph -ErrorAction Stop
Import-Module ExchangeOnlineManagement -ErrorAction Stop

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes @(
    'User.Read.All','Group.Read.All','Directory.Read.All',
    'AuditLog.Read.All','Mail.Read','MailboxSettings.Read',
    'Policy.Read.All','Application.Read.All'
) -NoWelcome -ErrorAction Stop

# Connect to Exchange Online
Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
try {
    Connect-ExchangeOnline -ShowBanner:$false
    $exchangeConnected = $true
    Write-Host "Connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Warning "Failed to connect to Exchange Online: $($_.Exception.Message)"
    $exchangeConnected = $false
    throw "Exchange Online connection required for mailbox operations"
}

# Collect user objects
$selectedUsers = @()
foreach ($upn in $UserPrincipalNames) {
    $user = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
    if ($user) { $selectedUsers += $user }
    else { Write-Warning "User not found: $upn" }
}

if (-not $selectedUsers) {
    Write-Error "No valid users found. Exiting."
    exit 1
}


# Collect additional data for each selected user
$userResults = $selectedUsers
$groupMemberResults = @()
$mailboxRules = @()
$delegatedMailboxes = @()
$calendars = @()
$appResults = @()

foreach ($user in $selectedUsers) {
    $userUpn = $user.UserPrincipalName

    # Group Memberships
    $userGroups = Get-MgUserMemberOf -UserId $userUpn -ErrorAction SilentlyContinue | Where-Object { $_.AdditionalProperties["@odata.type"] -eq '#microsoft.graph.group' }
    foreach ($group in $userGroups) {
        $groupMemberResults += [PSCustomObject]@{
            MemberUPN = $userUpn
            GroupName = $group.AdditionalProperties["displayName"]
            GroupType = $group.AdditionalProperties["groupTypes"] -join ','
            MemberType = 'User'
        }
    }

    # Mailbox Rules (ExchangeOnlineManagement required)
    try {
        $rules = Get-InboxRule -Mailbox $userUpn -ErrorAction SilentlyContinue
        foreach ($rule in $rules) {
            $mailboxRules += [PSCustomObject]@{
                MailboxOwner = $userUpn
                RuleName = $rule.Name
                Enabled = $rule.Enabled
                ForwardTo = ($rule.ForwardTo | ForEach-Object { $_.Name }) -join ','
                RedirectTo = ($rule.RedirectTo | ForEach-Object { $_.Name }) -join ','
                Description = $rule.Description
                Priority = $rule.Priority
            }
        }
    } catch {}

    # Delegated Mailboxes (user is owner or delegate)
    try {
        $delegates = Get-MailboxPermission -Identity $userUpn -ErrorAction SilentlyContinue | Where-Object { $_.User -notin @('NT AUTHORITY\SELF','NT AUTHORITY\SYSTEM','S-1-5-32-544') }
        foreach ($del in $delegates) {
            $delegatedMailboxes += [PSCustomObject]@{
                MailboxOwner = $userUpn
                DelegateUser = $del.User
                AccessRights = $del.AccessRights -join ','
                IsInherited = $del.IsInherited
                DenyRights = $del.Deny
            }
        }
        $asDelegate = Get-MailboxPermission -User $userUpn -ErrorAction SilentlyContinue
        foreach ($del in $asDelegate) {
            $delegatedMailboxes += [PSCustomObject]@{
                MailboxOwner = $del.Identity
                DelegateUser = $userUpn
                AccessRights = $del.AccessRights -join ','
                IsInherited = $del.IsInherited
                DenyRights = $del.Deny
            }
        }
    } catch {}

    # Calendar Permissions
    try {
        $folders = Get-MailboxFolderStatistics -Identity $userUpn -ErrorAction SilentlyContinue | Where-Object { $_.FolderType -eq 'Calendar' }
        foreach ($folder in $folders) {
            $folderIdentity = '{0}:{1}' -f $userUpn, $folder.FolderPath
            $perms = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction SilentlyContinue
            foreach ($perm in $perms) {
                if ($perm.User -notin @('Default','Anonymous')) {
                    $calendars += [PSCustomObject]@{
                        CalendarOwner = $userUpn
                        CalendarName = $folder.Name
                        SharedWithUser = $perm.User
                        AccessRights = $perm.AccessRights -join ','
                        MailboxType = $folder.FolderType
                        FolderPath = $folder.FolderPath
                        ItemCount = $folder.ItemsInFolder
                    }
                }
            }
        }
    } catch {}

    # Enterprise Applications (assignments)
    try {
        $spAssignments = Get-MgUserAppRoleAssignment -UserId $userUpn -ErrorAction SilentlyContinue
        foreach ($app in $spAssignments) {
            $appResults += [PSCustomObject]@{
                DisplayName = $app.AdditionalProperties["resourceDisplayName"]
                Homepage = ''
                LoginUrl = ''
                LogoutUrl = ''
                AssignedUser = $userUpn
                AppOwner = ''
            }
        }
    } catch {}
}

# Create Excel package
$pkg = Open-ExcelPackage -Path $ExcelFile -Create

foreach ($user in $selectedUsers) {
    $userUpn = $user.UserPrincipalName
    $safeUsername = ($user.DisplayName -replace '[\\/?*\[\]:"]', '_').Substring(0, [Math]::Min(31, $user.DisplayName.Length))
    Write-Host "Processing worksheet for user: $safeUsername"
    $wsUser = $pkg.Workbook.Worksheets.Add($safeUsername)
    $row = 1

    # User Profile Section
    $wsUser.Cells[$row, 1].Value = "=== USER PROFILE ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2

    # Add user details
    foreach ($prop in $user.PSObject.Properties) {
        $wsUser.Cells[$row, 1].Value = $prop.Name
        $wsUser.Cells[$row, 2].Value = $prop.Value
        $row++
    }
    $row += 2

    # Group Memberships Section
    $wsUser.Cells[$row, 1].Value = "=== GROUP MEMBERSHIPS ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("GroupName", "GroupType", "MemberType")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $userGroups = $groupMemberResults | Where-Object { $_.MemberUPN -eq $userUpn }
    if ($userGroups) {
        foreach ($group in $userGroups) {
            $wsUser.Cells[$row, 1].Value = $group.GroupName
            $wsUser.Cells[$row, 2].Value = $group.GroupType
            $wsUser.Cells[$row, 3].Value = $group.MemberType
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No group memberships found"
        $row++
    }
    $row += 2

    # Mailbox Rules Section
    $wsUser.Cells[$row, 1].Value = "=== MAILBOX RULES ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("RuleName", "Enabled", "ForwardTo", "RedirectTo", "Description", "Priority")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $userRules = $mailboxRules | Where-Object { $_.MailboxOwner -eq $userUpn }
    if ($userRules) {
        foreach ($rule in $userRules) {
            $wsUser.Cells[$row, 1].Value = $rule.RuleName
            $wsUser.Cells[$row, 2].Value = $rule.Enabled
            $wsUser.Cells[$row, 3].Value = $rule.ForwardTo
            $wsUser.Cells[$row, 4].Value = $rule.RedirectTo
            $wsUser.Cells[$row, 5].Value = $rule.Description
            $wsUser.Cells[$row, 6].Value = $rule.Priority
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No mailbox rules found"
        $row++
    }
    $row += 2

    # Delegated Mailboxes Section (Owner)
    $wsUser.Cells[$row, 1].Value = "=== DELEGATED MAILBOXES (USER IS OWNER) ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("DelegateUser", "AccessRights", "IsInherited", "DenyRights")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $delegatedByUser = $delegatedMailboxes | Where-Object { $_.MailboxOwner -eq $userUpn }
    if ($delegatedByUser) {
        foreach ($delegation in $delegatedByUser) {
            $wsUser.Cells[$row, 1].Value = $delegation.DelegateUser
            $wsUser.Cells[$row, 2].Value = $delegation.AccessRights
            $wsUser.Cells[$row, 3].Value = $delegation.IsInherited
            $wsUser.Cells[$row, 4].Value = $delegation.DenyRights
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No delegated mailboxes owned by this user"
        $row++
    }
    $row += 2

    # Delegated Mailboxes Section (Delegate)
    $wsUser.Cells[$row, 1].Value = "=== DELEGATED MAILBOXES (USER IS DELEGATE) ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("MailboxOwner", "AccessRights", "IsInherited", "DenyRights")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $delegatedToUser = $delegatedMailboxes | Where-Object { $_.DelegateUser -eq $userUpn }
    if ($delegatedToUser) {
        foreach ($delegation in $delegatedToUser) {
            $wsUser.Cells[$row, 1].Value = $delegation.MailboxOwner
            $wsUser.Cells[$row, 2].Value = $delegation.AccessRights
            $wsUser.Cells[$row, 3].Value = $delegation.IsInherited
            $wsUser.Cells[$row, 4].Value = $delegation.DenyRights
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No mailboxes delegated to this user"
        $row++
    }
    $row += 2

    # Calendar Permissions Section (Shared With User)
    $wsUser.Cells[$row, 1].Value = "=== CALENDARS SHARED WITH THIS USER ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("CalendarOwner", "CalendarName", "SharedWithUser", "AccessRights", "MailboxType", "FolderPath", "ItemCount")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $sharedWithUser = $calendars | Where-Object { $_.SharedWithUser -eq $userUpn }
    if ($sharedWithUser) {
        foreach ($cal in $sharedWithUser) {
            $wsUser.Cells[$row, 1].Value = $cal.CalendarOwner
            $wsUser.Cells[$row, 2].Value = $cal.CalendarName
            $wsUser.Cells[$row, 3].Value = $cal.SharedWithUser
            $wsUser.Cells[$row, 4].Value = $cal.AccessRights
            $wsUser.Cells[$row, 5].Value = $cal.MailboxType
            $wsUser.Cells[$row, 6].Value = $cal.FolderPath
            $wsUser.Cells[$row, 7].Value = $cal.ItemCount
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No calendars shared with this user"
        $row++
    }
    $row += 2

    # Calendar Permissions Section (User Shares)
    $wsUser.Cells[$row, 1].Value = "=== CALENDARS THIS USER SHARES ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("SharedWithUser", "AccessRights", "FolderPath")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $sharedByUser = $calendars | Where-Object { $_.CalendarOwner -eq $userUpn }
    if ($sharedByUser) {
        foreach ($cal in $sharedByUser) {
            $wsUser.Cells[$row, 1].Value = $cal.SharedWithUser
            $wsUser.Cells[$row, 2].Value = $cal.AccessRights
            $wsUser.Cells[$row, 3].Value = $cal.FolderPath
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No calendars shared by this user"
        $row++
    }
    $row += 2

    # Enterprise Applications Section
    $wsUser.Cells[$row, 1].Value = "=== ENTERPRISE APPLICATIONS ==="
    $wsUser.Cells[$row, 1, $row, 6].Merge = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Font.Bold = $true
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.PatternType = 'Solid'
    $wsUser.Cells[$row, 1, $row, 6].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightSteelBlue)
    $row += 2
    $headers = @("DisplayName", "Homepage", "LoginUrl", "LogoutUrl", "AppOwner")
    1..$headers.Count | ForEach-Object { $wsUser.Cells[$row, $_].Value = $headers[$_ - 1]; $wsUser.Cells[$row, $_].Style.Font.Bold = $true }
    $row++
    $userApps = $appResults | Where-Object { $_.AssignedUser -eq $userUpn }
    if ($userApps) {
        foreach ($app in $userApps) {
            $wsUser.Cells[$row, 1].Value = $app.DisplayName
            $wsUser.Cells[$row, 2].Value = $app.Homepage
            $wsUser.Cells[$row, 3].Value = $app.LoginUrl
            $wsUser.Cells[$row, 4].Value = $app.LogoutUrl
            $wsUser.Cells[$row, 5].Value = $app.AppOwner
            $row++
        }
    } else {
        $wsUser.Cells[$row, 1].Value = "No enterprise applications found"
        $row++
    }
    $row += 2
}

# Save and close
Close-ExcelPackage $pkg -Show
Write-Host "User audit report exported to $ExcelFile" -ForegroundColor Green
