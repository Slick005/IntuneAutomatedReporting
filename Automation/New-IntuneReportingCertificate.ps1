<#

.SYNOPSIS
Creates a self-signed certificate for the Intune Automated Reporting app registration.

.DESCRIPTION
Generates an RSA key pair and a self-signed certificate, then writes two files to OutputPath:

  <Name>.cer  Public certificate. Upload it to the app registration
              (New-IntuneReportingTenant.ps1 -CertificatePath).
  <Name>.pfx  Certificate with private key, protected by PfxPassword. Import it into the
              Automation Account (Certificates) or the machine that runs the reports.

The certificate is built with .NET APIs, so the script runs on Windows, macOS and Linux
with PowerShell 7, and on Windows PowerShell 5.1 with .NET Framework 4.7.2 or later.

.PARAMETER Name
Base name of the output files. Defaults to IntuneAutomatedReporting.

.PARAMETER Subject
Certificate subject. Defaults to CN=<Name>.

.PARAMETER ValidityMonths
How long the certificate is valid, in months (1 to 60). Defaults to 12.

.PARAMETER KeyLength
RSA key size in bits. Defaults to 2048.

.PARAMETER OutputPath
Folder for the .cer and .pfx files. Created if it does not exist. Defaults to the current folder.

.PARAMETER PfxPassword
Password that protects the .pfx file. When omitted a random password is generated and
returned in the output object's PfxPassword property.

.PARAMETER Force
Overwrite existing .cer and .pfx files with the same name.

.EXAMPLE
$cert = .\New-IntuneReportingCertificate.ps1 -OutputPath .\certs
.\New-IntuneReportingTenant.ps1 -TenantId contoso.onmicrosoft.com -CertificatePath $cert.CerPath

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
param
(
    [ValidatePattern('^[A-Za-z0-9._-]+$')]
    [string]$Name = 'IntuneAutomatedReporting',

    [string]$Subject,

    [ValidateRange(1, 60)]
    [int]$ValidityMonths = 12,

    [ValidateSet(2048, 3072, 4096)]
    [int]$KeyLength = 2048,

    [string]$OutputPath = (Get-Location).Path,

    [securestring]$PfxPassword,

    [switch]$Force
)

$ErrorActionPreference = 'Stop'

####################################################

function New-RandomPassword {

    param
    (
        [int]$Length = 32
    )

    $chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789!#%+-=?@_'.ToCharArray()
    $bytes = New-Object byte[] $Length
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try {
        $rng.GetBytes($bytes)
    }
    finally {
        $rng.Dispose()
    }

    $password = New-Object System.Security.SecureString
    foreach ($b in $bytes) {
        $password.AppendChar($chars[$b % $chars.Length])
    }
    $password.MakeReadOnly()
    return $password
}

####################################################

if (-not $Subject) {
    $Subject = "CN=$Name"
}

$cerPath = Join-Path $OutputPath "$Name.cer"
$pfxPath = Join-Path $OutputPath "$Name.pfx"

foreach ($path in $cerPath, $pfxPath) {
    if ((Test-Path -LiteralPath $path) -and -not $Force) {
        throw "File '$path' already exists. Use -Force to overwrite it."
    }
}

if (-not $PSCmdlet.ShouldProcess($OutputPath, "Create self-signed certificate '$Subject' ($Name.cer, $Name.pfx)")) {
    return
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

# .NET file APIs resolve relative paths against the process folder, not the PowerShell location
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path
$cerPath = Join-Path $OutputPath "$Name.cer"
$pfxPath = Join-Path $OutputPath "$Name.pfx"

$passwordGenerated = $false
if (-not $PfxPassword) {
    $PfxPassword = New-RandomPassword
    $passwordGenerated = $true
}

$rsa = [System.Security.Cryptography.RSA]::Create($KeyLength)
try {
    $request = New-Object System.Security.Cryptography.X509Certificates.CertificateRequest(
        $Subject,
        $rsa,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256,
        [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)

    $request.CertificateExtensions.Add(
        (New-Object System.Security.Cryptography.X509Certificates.X509BasicConstraintsExtension($false, $false, 0, $true)))
    $request.CertificateExtensions.Add(
        (New-Object System.Security.Cryptography.X509Certificates.X509KeyUsageExtension(
            [System.Security.Cryptography.X509Certificates.X509KeyUsageFlags]'DigitalSignature, KeyEncipherment', $true)))

    $ekus = New-Object System.Security.Cryptography.OidCollection
    [void]$ekus.Add((New-Object System.Security.Cryptography.Oid('1.3.6.1.5.5.7.3.2', 'Client Authentication')))
    $request.CertificateExtensions.Add(
        (New-Object System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension($ekus, $false)))

    $notBefore = [DateTimeOffset]::UtcNow.AddMinutes(-5)
    $notAfter = $notBefore.AddMonths($ValidityMonths)
    $certificate = $request.CreateSelfSigned($notBefore, $notAfter)

    try {
        $cerBytes = $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
        $pfxBytes = $certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $PfxPassword)

        [System.IO.File]::WriteAllBytes($cerPath, $cerBytes)
        [System.IO.File]::WriteAllBytes($pfxPath, $pfxBytes)

        [pscustomobject]@{
            Subject           = $certificate.Subject
            Thumbprint        = $certificate.Thumbprint
            NotBefore         = $certificate.NotBefore
            NotAfter          = $certificate.NotAfter
            CerPath           = $cerPath
            PfxPath           = $pfxPath
            PfxPassword       = $PfxPassword
            PasswordGenerated = $passwordGenerated
        }
    }
    finally {
        $certificate.Dispose()
    }
}
finally {
    $rsa.Dispose()
}
