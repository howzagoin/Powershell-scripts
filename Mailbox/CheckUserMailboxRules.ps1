# ================================================
# Author: Tim MacLatchy
# Date: 16/10/2024
# License: MIT License
# Description: Script to log in with MFA, check user's mailbox rules,
# including forwarding/redirection rules, and allow disabling/removing rules.
# ================================================

# Function to check for required modules and install if missing
function Check-RequiredModules {
    $modules = @("ExchangeOnlineManagement")
    foreach ($module in $modules) {
        if (-not (Get-Module -ListAvailable -Name ${module})) {
            Install-Module -Name ${module} -Force -AllowClobber
        }
    }
}

# Function for logging in with MFA
function Login-WithMFA {
    $adminEmail = Read-Host "Enter the admin email address"
    Write-Host "Logging in as admin with MFA..."
    try {
        Connect-ExchangeOnline -UserPrincipalName ${adminEmail} -ShowProgress $true
        Write-Host "Login successful." -ForegroundColor Green
    } catch {
        Write-Host "Login failed. Please check credentials or MFA method." -ForegroundColor Red
        throw $_
    }
}

# Function to get and display user's inbox and redirection rules
function Get-MailboxRules {
    param ($emailAddress)

    Write-Host "`nFetching all rules for ${emailAddress}..."

    try {
        # Fetch all inbox rules and redirection settings
        $rules = Get-InboxRule -Mailbox ${emailAddress} | Sort-Object Date -Descending

        $redirectRule = Get-Mailbox -Identity ${emailAddress} | Select-Object -ExpandProperty ForwardingAddress
        $redirectRuleEnabled = Get-Mailbox -Identity ${emailAddress} | Select-Object -ExpandProperty DeliverToMailboxAndForward

        if (${rules}.Count -eq 0 -and -not ${redirectRule}) {
            Write-Host "No inbox rules or redirection rules found for ${emailAddress}." -ForegroundColor Yellow
            return
        }

        Write-Host "`nRules found for ${emailAddress}:"

        # Display Inbox Rules
        if (${rules}.Count -gt 0) {
            Write-Host "`nInbox Rules:"
            $i = 1
            foreach ($rule in ${rules}) {
                Write-Host "$i. Rule Name: ${rule.Name}, Created: ${rule.Date}, Enabled: ${rule.Enabled}" -ForegroundColor Cyan
                $i++
            }
        }

        # Display Redirection/Forwarding Rules
        if (${redirectRule}) {
            Write-Host "`nForwarding/Redirection Rule:"
            $redirectStatus = if (${redirectRuleEnabled}) { "Enabled" } else { "Disabled" }
            Write-Host "$i. Forwarding Address: ${redirectRule}, Status: ${redirectStatus}" -ForegroundColor Cyan
        }

        return ${rules}, ${redirectRule}
    } catch {
        Write-Host "Error fetching rules for ${emailAddress}." -ForegroundColor Red
        throw $_
    }
}

# Function to disable or remove a rule
function Modify-MailboxRules {
    param ($emailAddress, $rules, $redirectRule)

    while ($true) {
        $choice = Read-Host "`nWould you like to (D)isable, (R)emove, (M)odify a rule, or (E)xit?"
        if ($choice -match '^[DREMdrem]$') {
            if ($choice -match '[Ee]') {
                break
            }

            $ruleNumber = Read-Host "Enter the rule number to modify (Inbox Rule or Redirection)"
            if ($ruleNumber -match '^\d+$' -and ${ruleNumber} -le ${rules}.Count) {
                $selectedRule = ${rules}[${ruleNumber} - 1]
                $ruleName = ${selectedRule}.Name

                if ($choice -match '[Dd]') {
                    Write-Host "`nDisabling rule: ${ruleName}..."
                    Disable-InboxRule -Identity ${ruleName} -Mailbox ${emailAddress}
                    Write-Host "Rule disabled." -ForegroundColor Green
                } elseif ($choice -match '[Rr]') {
                    Write-Host "`nRemoving rule: ${ruleName}..."
                    Remove-InboxRule -Identity ${ruleName} -Mailbox ${emailAddress}
                    Write-Host "Rule removed." -ForegroundColor Green
                } elseif ($choice -match '[Mm]') {
                    Write-Host "`nModifying rule: ${ruleName}..."
                    # Modify logic here (example could be changing conditions, etc.)
                }

                # Refresh the list of rules after modification
                $rules = Get-MailboxRules -emailAddress ${emailAddress}
                continue
            } elseif (${ruleNumber} -eq ${rules}.Count + 1 -and ${redirectRule}) {
                if ($choice -match '[Dd]') {
                    Write-Host "`nDisabling forwarding/redirection rule..."
                    Set-Mailbox -Identity ${emailAddress} -ForwardingAddress $null -DeliverToMailboxAndForward $false
                    Write-Host "Forwarding rule disabled." -ForegroundColor Green
                } elseif ($choice -match '[Rr]') {
                    Write-Host "`nRemoving forwarding/redirection rule..."
                    Set-Mailbox -Identity ${emailAddress} -ForwardingAddress $null
                    Write-Host "Forwarding rule removed." -ForegroundColor Green
                }

                # Refresh the list of rules after modification
                $rules = Get-MailboxRules -emailAddress ${emailAddress}
                continue
            } else {
                Write-Host "Invalid rule number. Please enter a valid number." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Invalid choice. Please choose D, R, M, or E." -ForegroundColor Yellow
        }
    }
}

# Main function
function Main {
    Check-RequiredModules
    Login-WithMFA

    $emailAddress = Read-Host "Enter the user's email address"
    $rules, $redirectRule = Get-MailboxRules -emailAddress ${emailAddress}

    if ((${rules} -ne $null -and ${rules}.Count -gt 0) -or ${redirectRule}) {
        Modify-MailboxRules -emailAddress ${emailAddress} -rules ${rules} -redirectRule ${redirectRule}
    } else {
        Write-Host "No rules to modify."
    }

    Disconnect-ExchangeOnline -Confirm:$false
}

# Run the script
Main
