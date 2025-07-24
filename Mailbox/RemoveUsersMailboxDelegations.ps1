<#
.SYNOPSIS
    Remove all mailbox delegations (FullAccess, SendAs, and SendOnBehalf) for a specified user
    from all mailboxes in Exchange Online.

.DESCRIPTION
    •   Checks for and installs required PowerShell modules automatically
    •   Prompts for the user email address to remove delegations from
    •   Efficiently finds and removes all three types of mailbox delegations:
        - Full Access permissions
        - Send As permissions  
        - Send On Behalf permissions
    •   Provides detailed progress feedback and error reporting
    •   Optimized to only check mailboxes where permissions actually exist
#>

# --- Initialize and Set Security Protocol ---
Write-Host "=== Exchange Online Delegation Removal Script ===" -ForegroundColor Cyan
Write-Host "This script will remove ALL mailbox delegations for a specified user." -ForegroundColor Yellow
Write-Host ""

try {
    # Ensure modern security protocol is used for connections
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
    Write-Host "✓ Security protocol set to TLS 1.2" -ForegroundColor Green
} catch {
    Write-Warning "⚠ Could not set TLS 1.2. This may cause issues on older systems."
}

# Allow this script to run regardless of system policy, for this session only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
Write-Host "✓ Execution policy bypassed for this session" -ForegroundColor Green

# --- Check and Install Required Modules ---
Write-Host "`n--- Checking Required Modules ---" -ForegroundColor Cyan

$requiredModules = @(
    @{Name = 'ExchangeOnlineManagement'; DisplayName = 'Exchange Online Management'}
)

foreach ($module in $requiredModules) {
    Write-Host "Checking for $($module.DisplayName) module..." -NoNewline
    
    if (Get-Module -ListAvailable $module.Name) {
        Write-Host " ✓ Found" -ForegroundColor Green
    } else {
        Write-Host " ✗ Not found - Installing..." -ForegroundColor Yellow
        
        try {
            # Ensure NuGet provider is available
            if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                Write-Host "  Installing NuGet package provider..." -ForegroundColor Yellow
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop
            }
            
            # Install the required module
            Write-Host "  Installing $($module.DisplayName) module..." -ForegroundColor Yellow
            Install-Module $module.Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            Write-Host "  ✓ Successfully installed $($module.DisplayName)" -ForegroundColor Green
            
        } catch {
            Write-Error "✗ Failed to install $($module.DisplayName) module: $($_.Exception.Message)"
            Write-Host "Please install the module manually using: Install-Module $($module.Name)" -ForegroundColor Red
            return
        }
    }
}

# Import the Exchange Online module
Write-Host "Importing Exchange Online Management module..." -NoNewline
try {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Write-Host " ✓ Imported" -ForegroundColor Green
} catch {
    Write-Error "✗ Failed to import Exchange Online Management module: $($_.Exception.Message)"
    return
}

# --- Connect to Exchange Online ---
Write-Host "`n--- Connecting to Exchange Online ---" -ForegroundColor Cyan
try {
    Write-Host "Initiating connection to Exchange Online..." -ForegroundColor Yellow
    Write-Host "Please authenticate when prompted." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "✓ Successfully connected to Exchange Online" -ForegroundColor Green
} catch {
    Write-Error "✗ Failed to connect to Exchange Online: $($_.Exception.Message)"
    Write-Host "Please check your credentials and ensure you have the necessary permissions." -ForegroundColor Red
    return
}

# --- Get User Input ---
Write-Host "`n--- User Input ---" -ForegroundColor Cyan
do {
    $userToRemove = Read-Host "Enter the email address of the user to remove all delegations from"
    
    # Validate email format
    if ($userToRemove -notlike '*@*' -or $userToRemove -notlike '*.*') {
        Write-Warning "⚠ '$userToRemove' does not appear to be a valid email address."
        Write-Host "Please enter a valid email address (e.g., user@domain.com)" -ForegroundColor Yellow
        $userToRemove = $null
    }
} while (-not $userToRemove)

Write-Host "`n=== Processing Delegations for: $userToRemove ===" -ForegroundColor Green
Write-Host "This may take several minutes depending on the number of mailboxes..." -ForegroundColor Yellow

# --- Remove Full Access Permissions ---
Write-Host "`n[1/3] Checking for Full Access permissions..." -ForegroundColor Cyan
$fullAccessCount = 0

try {
    $fullAccessMailboxes = Get-EXOMailbox -ResultSize Unlimited | Get-EXOMailboxPermission | Where-Object { 
        ($_.User -eq $userToRemove) -and 
        ($_.AccessRights -match "FullAccess") -and 
        ($_.IsInherited -eq $false) 
    }

    if ($fullAccessMailboxes) {
        foreach ($permission in $fullAccessMailboxes) {
            $mailboxId = $permission.Identity
            Write-Host "  Removing Full Access from: $mailboxId" -ForegroundColor Yellow
            
            try {
                Remove-MailboxPermission -Identity $mailboxId -User $userToRemove -AccessRights FullAccess -Confirm:$false -InheritanceType All -ErrorAction Stop
                Write-Host "    ✓ Successfully removed" -ForegroundColor Green
                $fullAccessCount++
            } catch {
                Write-Warning "    ✗ Failed to remove Full Access from '$mailboxId': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "  ✓ No Full Access permissions found for this user" -ForegroundColor Green
    }
} catch {
    Write-Warning "⚠ Error checking Full Access permissions: $($_.Exception.Message)"
}

# --- Remove Send As Permissions ---
Write-Host "`n[2/3] Checking for Send As permissions..." -ForegroundColor Cyan
$sendAsCount = 0

try {
    $sendAsMailboxes = Get-EXOMailbox -ResultSize Unlimited | Get-RecipientPermission -Trustee $userToRemove -ErrorAction SilentlyContinue

    if ($sendAsMailboxes) {
        foreach ($permission in $sendAsMailboxes) {
            $mailboxId = $permission.Identity
            Write-Host "  Removing Send As from: $mailboxId" -ForegroundColor Yellow
            
            try {
                Remove-RecipientPermission -Identity $mailboxId -Trustee $userToRemove -AccessRights SendAs -Confirm:$false -ErrorAction Stop
                Write-Host "    ✓ Successfully removed" -ForegroundColor Green
                $sendAsCount++
            } catch {
                Write-Warning "    ✗ Failed to remove Send As from '$mailboxId': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "  ✓ No Send As permissions found for this user" -ForegroundColor Green
    }
} catch {
    Write-Warning "⚠ Error checking Send As permissions: $($_.Exception.Message)"
}

# --- Remove Send On Behalf Permissions ---
Write-Host "`n[3/3] Checking for Send On Behalf permissions..." -ForegroundColor Cyan
$sendOnBehalfCount = 0

try {
    # Use filter to efficiently find mailboxes with Send On Behalf permissions
    $sendOnBehalfMailboxes = Get-EXOMailbox -ResultSize Unlimited -Filter "GrantSendOnBehalfTo -eq '$userToRemove'"

    if ($sendOnBehalfMailboxes) {
        foreach ($mailbox in $sendOnBehalfMailboxes) {
            $mailboxId = $mailbox.PrimarySmtpAddress
            Write-Host "  Removing Send On Behalf from: $mailboxId" -ForegroundColor Yellow
            
            try {
                # Use the @{remove=...} syntax to remove a value from a multi-valued property
                Set-Mailbox -Identity $mailboxId -GrantSendOnBehalfTo @{remove = $userToRemove} -ErrorAction Stop
                Write-Host "    ✓ Successfully removed" -ForegroundColor Green
                $sendOnBehalfCount++
            } catch {
                Write-Warning "    ✗ Failed to remove Send On Behalf from '$mailboxId': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Host "  ✓ No Send On Behalf permissions found for this user" -ForegroundColor Green
    }
} catch {
    Write-Warning "⚠ Error checking Send On Behalf permissions: $($_.Exception.Message)"
}

# --- Summary Report ---
Write-Host "`n=== SUMMARY REPORT ===" -ForegroundColor Cyan
Write-Host "User processed: $userToRemove" -ForegroundColor White
Write-Host "Delegations removed:" -ForegroundColor White
Write-Host "  • Full Access permissions: $fullAccessCount" -ForegroundColor Green
Write-Host "  • Send As permissions: $sendAsCount" -ForegroundColor Green
Write-Host "  • Send On Behalf permissions: $sendOnBehalfCount" -ForegroundColor Green
Write-Host "  • Total delegations removed: $($fullAccessCount + $sendAsCount + $sendOnBehalfCount)" -ForegroundColor Green

if (($fullAccessCount + $sendAsCount + $sendOnBehalfCount) -eq 0) {
    Write-Host "`n✓ No delegations were found for this user - they may have already been removed." -ForegroundColor Yellow
} else {
    Write-Host "`n✓ All delegations have been successfully processed!" -ForegroundColor Green
}

# --- Disconnect and Cleanup ---
Write-Host "`n--- Disconnecting from Exchange Online ---" -ForegroundColor Cyan
try {
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "✓ Successfully disconnected from Exchange Online" -ForegroundColor Green
} catch {
    Write-Warning "⚠ Error during disconnection: $($_.Exception.Message)"
}

Write-Host "`n=== Script Completed ===" -ForegroundColor Cyan
Write-Host "Review any warnings above for delegations that could not be removed." -ForegroundColor Yellow
Write-Host "Press any key to exit..." -ForegroundColor Gray
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")