# Reconnect with explicit scopes if needed
Connect-MgGraph -Scopes "Directory.Read.All", "RoleManagement.Read.Directory" -NoWelcome

# Get the current user object
$userUPN = "timothy.maclatchy@journebrands.com"
$user = Get-MgUser -UserId $userUPN

# Get all directory roles
$allRoles = Get-MgDirectoryRole

# Find assigned roles for this user
$assignedRoles = @()
foreach ($role in $allRoles) {
    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All
    if ($members.Id -contains $user.Id) {
        $assignedRoles += $role
    }
}

if ($assignedRoles) {
    Write-Host "✅ Roles assigned to ${userUPN}:" -ForegroundColor Green
    $assignedRoles | Select-Object DisplayName, Id
} else {
    Write-Host "⚠️ No directory roles found for $userUPN." -ForegroundColor Yellow
}
