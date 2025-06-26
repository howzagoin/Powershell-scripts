<#
.SYNOPSIS
  Clean up broken Chocolatey install and reinstall Chocolatey 2.4.3.0
.NOTES
  Run from an elevated PowerShell 7 x64 session.
#>

function Backup-Item {
    param(
        [Parameter(Mandatory)][string]$Path
    )
    if (Test-Path $Path) {
        $Backup = "$Path.backup_$(Get-Date -f yyyyMMdd_HHmmss)"
        Write-Host "🗄️  Backing up '$Path' → '$Backup'"
        Copy-Item $Path $Backup -Recurse -Force
        Write-Host "✅ Backup complete"
    }
}

function Remove-IfExists {
    param(
        [Parameter(Mandatory)][string]$Path
    )
    if (Test-Path $Path) {
        Write-Host "🧹 Removing '$Path'"
        Remove-Item $Path -Recurse -Force
    }
}

function Ensure-Path {
    param(
        [Parameter(Mandatory)][string]$NewPath
    )
    $current = [Environment]::GetEnvironmentVariable('Path','Machine')
    if ($current -notmatch [Regex]::Escape($NewPath)) {
        Write-Host "➕ Adding $NewPath to system PATH"
        [Environment]::SetEnvironmentVariable('Path', "$current;$NewPath", 'Machine')
    }
}

Write-Host "🔍  Checking for stray Chocolatey artefacts..."
$chocoDir   = "C:\ProgramData\chocolatey"
$regKey32   = "HKLM:\SOFTWARE\Chocolatey"
$regKey64   = "HKLM:\SOFTWARE\WOW6432Node\Chocolatey"

Backup-Item -Path $chocoDir
Remove-IfExists $chocoDir

foreach ($key in $regKey32,$regKey64) {
    if (Test-Path $key) {
        Write-Host "🧹 Removing registry key $key"
        Remove-Item -Path $key -Recurse -Force
    }
}

Write-Host "`n⬇️  Downloading Chocolatey MSI..."
$msiUrl  = "https://github.com/chocolatey/choco/releases/download/2.4.3/chocolatey-2.4.3.0.msi"
$msiPath = "$env:TEMP\choco-2.4.3.0.msi"
Invoke-WebRequest $msiUrl -OutFile $msiPath

Write-Host "⚙️  Running MSI *silently* with full log (may prompt for UAC)..."
$logPath = "$env:TEMP\choco-install.log"
Start-Process "msiexec.exe" -Wait -Verb RunAs -ArgumentList @(
    "/i `"$msiPath`"",
    "/qn",
    "/l*v `"$logPath`""
)

Write-Host "`n📄 Installer log saved to $logPath"
if (-not (Test-Path "$chocoDir\bin\choco.exe")) {
    Write-Warning "❌ choco.exe not found in expected location. Check the log above."
    return
}

Ensure-Path "$chocoDir\bin"

Write-Host "`n🎉 Chocolatey reinstall complete!"
Write-Host "   -> Close *all* PowerShell/Command windows and open a new one."
Write-Host "   -> Verify with: choco --version"
