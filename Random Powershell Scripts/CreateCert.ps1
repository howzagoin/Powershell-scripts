<#
.SYNOPSIS
Creates a self‑signed certificate (PFX + CER), uploads the public key to
an Azure AD app (by OBJECT ID), and trusts the certificate locally.

.NOTES
Author  : ChatGPT
Updated : 02‑07‑2025
Fixes   : Correct X509Store constructor (uses StoreName/StoreLocation enums)
#>

#region ─── Helper Functions ────────────────────────────────────────────────────────────────

function New-AppCertificate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $CertName,
        [Parameter(Mandatory)][string] $OutputFolder,
        [Parameter(Mandatory)][SecureString] $PfxPassword,
        [Parameter()][ValidateRange(1,2)][int] $ValidYears = 1
    )
    try {
        if (-not (Test-Path $OutputFolder)) {
            New-Item -ItemType Directory -Path $OutputFolder -ErrorAction Stop | Out-Null
        }

        $notAfter = (Get-Date).AddYears($ValidYears)
        Write-Host "▶ Creating certificate '$CertName' (expires $(Get-Date $notAfter -Format 'dd-MM-yyyy'))..."

        $cert = New-SelfSignedCertificate `
                    -Subject "CN=$CertName" `
                    -KeySpec Signature -KeyLength 2048 -KeyExportPolicy Exportable `
                    -CertStoreLocation 'Cert:\CurrentUser\My' `
                    -NotAfter $notAfter `
                    -KeyUsage DigitalSignature,KeyEncipherment `
                    -TextExtension @('2.5.29.37={text}1.3.6.1.5.5.7.3.3') `
                    -ErrorAction Stop

        $stamp   = Get-Date -Format 'dd-MM-yyyy'
        $pfxPath = Join-Path $OutputFolder "$CertName-$stamp.pfx"
        $cerPath = Join-Path $OutputFolder "$CertName-$stamp.cer"

        Export-PfxCertificate -Cert $cert -FilePath $pfxPath -Password $PfxPassword -Force -ErrorAction Stop | Out-Null
        Export-Certificate    -Cert $cert -FilePath $cerPath -Force -ErrorAction Stop | Out-Null

        Write-Host "✔ PFX exported to $pfxPath"
        Write-Host "✔ CER exported to $cerPath"

        [PSCustomObject]@{
            Certificate = $cert
            PfxPath     = $pfxPath
            CerPath     = $cerPath
        }
    }
    catch { throw "Failed to create/export certificate: $($_.Exception.Message)" }
}

function Add-CertificateToWindows {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$CerPath)

    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)

        $storeLocation = if ($isAdmin) {
            [System.Security.Cryptography.X509Certificates.StoreLocation]::LocalMachine
        } else {
            Write-Warning "Not running as Administrator – importing certificate to CurrentUser\\Root only."
            [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
        }

        $storeName = [System.Security.Cryptography.X509Certificates.StoreName]::Root
        $store     = New-Object System.Security.Cryptography.X509Certificates.X509Store($storeName, $storeLocation)
        $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)

        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($CerPath)
        if (-not ($store.Certificates | Where-Object { $_.Thumbprint -eq $cert.Thumbprint })) {
            $store.Add($cert)
            Write-Host "✔ Certificate added to Trusted Root store ($storeLocation)."
        } else {
            Write-Host "ℹ Certificate already present in Trusted Root store ($storeLocation)."
        }
        $store.Close()
    }
    catch { throw "Failed to add certificate to Windows store: $($_.Exception.Message)" }
}

function Add-CertificateToAzureApp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$AppObjectId,   # OBJECT ID
        [Parameter(Mandatory)][string]$CerPath,
        [Parameter()][int]$ValidYears = 1
    )

    begin {
        foreach ($m in 'Microsoft.Graph','Microsoft.Graph.Applications') {
            if (-not (Get-Module $m -ListAvailable)) {
                Write-Host "▶ Installing $m ..."
                Install-Module $m -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            }
            Import-Module $m -ErrorAction Stop
        }

        if (-not (Get-MgContext) -or (Get-MgContext).TenantId -ne $TenantId) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
            Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All' -ErrorAction Stop
        }
    }

    process {
        try {
            $app   = Get-MgApplication -ApplicationId $AppObjectId -Property keyCredentials -ErrorAction Stop
            $start = (Get-Date).ToUniversalTime().AddSeconds(-5)
            $end   = $start.AddYears([math]::Min($ValidYears,2)).AddSeconds(-10)
            $bytes = [IO.File]::ReadAllBytes($CerPath)

            $newKey = [Microsoft.Graph.PowerShell.Models.MicrosoftGraphKeyCredential]::new()
            $newKey.Type          = 'AsymmetricX509Cert'
            $newKey.Usage         = 'Verify'
            $newKey.Key           = $bytes
            $newKey.StartDateTime = $start
            $newKey.EndDateTime   = $end
            $newKey.DisplayName   = "CN=" + (Get-Item $CerPath).BaseName

            $merged = @($app.KeyCredentials) + $newKey

            Write-Host "▶ Uploading certificate to Azure AD app (Object ID $AppObjectId) ..."
            Update-MgApplication -ApplicationId $AppObjectId -KeyCredentials $merged -ErrorAction Stop
            Write-Host "✔ Certificate successfully added to Azure AD application."
        }
        catch { throw "Failed to upload certificate: $($_.Exception.Message)" }
    }
}

#endregion

function Main {

    #─── USER CONFIG ───────────────────────────────────────────────────────────────
    $tenantId     = 'ca0711e2-e703-4f4e-9099-17d97863211c'   # Directory (tenant) ID
    $appObjectId  = 'b453f5ac-28a7-44fd-96ae-eb5e7299c784'   # OBJECT ID of the app
    $certName     = 'SharePointAdmin-GraphCert'
    $outputFolder = "$env:USERPROFILE\Downloads\AppCerts"
    $pfxPlainPwd  = ''   # leave blank → auto‑generate secure password
    #───────────────────────────────────────────────────────────────────────────────

    if ([string]::IsNullOrWhiteSpace($pfxPlainPwd)) {
        $pfxPlainPwd = [guid]::NewGuid().ToString('N') + '!a1'
        Write-Warning "No PFX password supplied. Generated password: $pfxPlainPwd"
    }
    $securePwd = ConvertTo-SecureString $pfxPlainPwd -AsPlainText -Force

    $certInfo = New-AppCertificate -CertName $certName -OutputFolder $outputFolder `
                                   -PfxPassword $securePwd -ValidYears 1

    Add-CertificateToAzureApp -TenantId $tenantId -AppObjectId $appObjectId `
                              -CerPath $certInfo.CerPath -ValidYears 1

    Add-CertificateToWindows -CerPath $certInfo.CerPath
}

Main
