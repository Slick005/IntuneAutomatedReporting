<#

.SYNOPSIS
Creates or updates the Entra app registration used by the Intune Automated Reporting runbooks.

.DESCRIPTION
Automates step 1 of the set up ("Graph API Registration"). The script:

  1. Connects to Microsoft Graph with the Microsoft.Graph PowerShell SDK (delegated, as an admin).
  2. Creates a single-tenant app registration, or reuses the one with the same display name.
  3. Adds the Microsoft Graph application permissions the runbooks need.
  4. Creates the service principal and grants admin consent for those permissions.
  5. Optionally uploads a certificate (see New-IntuneReportingCertificate.ps1).
  6. Optionally creates a client secret and returns it as the 'GraphApi' credential the
     runbooks read with Get-AutomationPSCredential.

Running the script again is safe: existing permissions, consent and certificates are kept
and only what is missing is added.

Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Applications modules, and an
account that can create applications and grant admin consent (Cloud Application
Administrator plus Privileged Role Administrator, or Global Administrator).

.PARAMETER TenantId
Tenant ID or verified domain (for example contoso.onmicrosoft.com).

.PARAMETER DisplayName
Display name of the app registration. Defaults to 'Intune Automated Reporting'.

.PARAMETER Permissions
Microsoft Graph application permissions to grant. Defaults to the read-only permissions the
scripts in the Scripts folder use.

.PARAMETER CertificatePath
Path to a .cer file to add to the app registration's certificates.

.PARAMETER CreateClientSecret
Create a new client secret. The runbooks authenticate with a client secret stored in the
Automation Account credential 'GraphApi' (user name = client ID, password = secret).

.PARAMETER SecretValidityMonths
How long a new client secret is valid, in months (1 to 24). Defaults to 12.

.EXAMPLE
.\New-IntuneReportingTenant.ps1 -TenantId contoso.onmicrosoft.com -CreateClientSecret

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
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId,

    [ValidateNotNullOrEmpty()]
    [string]$DisplayName = 'Intune Automated Reporting',

    [ValidateNotNullOrEmpty()]
    [string[]]$Permissions = @(
        'DeviceManagementApps.Read.All',
        'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All',
        'Group.Read.All',
        'User.Read.All'
    ),

    [string]$CertificatePath,

    [switch]$CreateClientSecret,

    [ValidateRange(1, 24)]
    [int]$SecretValidityMonths = 12
)

$ErrorActionPreference = 'Stop'

$graphAppId = '00000003-0000-0000-c000-000000000000'
$adminScopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')

####################################################

function Invoke-WithRetry {

    # A new service principal can take a few seconds to replicate before it accepts role assignments
    param
    (
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = 5,
        [int]$DelaySeconds = 5
    )

    for ($attempt = 1; ; $attempt++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($attempt -ge $MaxAttempts) {
                throw
            }
            Write-Verbose "Attempt $attempt failed: $($_.Exception.Message). Retrying in $DelaySeconds seconds."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

####################################################

# Load the certificate before touching the tenant so a bad path fails fast
$certificate = $null
if ($CertificatePath) {
    if (-not (Test-Path -LiteralPath $CertificatePath -PathType Leaf)) {
        throw "Certificate file '$CertificatePath' was not found."
    }
    $certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2(
        (Resolve-Path -LiteralPath $CertificatePath).Path)
}

# Connect, reusing an existing session for the same tenant when it already has the scopes
$context = Get-MgContext
$hasScopes = $context -and -not ($adminScopes | Where-Object { $context.Scopes -notcontains $_ })
$sameTenant = $context -and ($context.TenantId -eq $TenantId -or $context.Account -like "*@$TenantId")
if (-not ($hasScopes -and $sameTenant)) {
    Connect-MgGraph -TenantId $TenantId -Scopes $adminScopes -NoWelcome
    $context = Get-MgContext
}
$TenantId = $context.TenantId

# Resolve permission names to Microsoft Graph app role IDs
$graphSp = Get-MgServicePrincipal -Filter "appId eq '$graphAppId'"
if (-not $graphSp) {
    throw "The Microsoft Graph service principal was not found in tenant '$TenantId'."
}

$roles = foreach ($permission in $Permissions) {
    $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $permission -and $_.AllowedMemberTypes -contains 'Application' }
    if (-not $role) {
        throw "'$permission' is not a Microsoft Graph application permission."
    }
    $role
}

# Find or create the app registration
$escapedName = $DisplayName -replace "'", "''"
$app = @(Get-MgApplication -Filter "displayName eq '$escapedName'")
if ($app.Count -gt 1) {
    throw "Found $($app.Count) app registrations named '$DisplayName'. Rename or remove the duplicates, or pass -DisplayName."
}
$app = $app | Select-Object -First 1

$graphAccess = @{
    ResourceAppId  = $graphAppId
    ResourceAccess = @($roles | ForEach-Object { @{ Id = $_.Id; Type = 'Role' } })
}

if (-not $app) {
    if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Create app registration')) {
        return
    }
    $app = New-MgApplication -DisplayName $DisplayName -SignInAudience 'AzureADMyOrg' -RequiredResourceAccess @($graphAccess)
    Write-Verbose "Created app registration '$DisplayName' ($($app.AppId))."
}
else {
    Write-Verbose "Using existing app registration '$DisplayName' ($($app.AppId))."

    # Keep permissions on other APIs and any extra Graph permissions already configured
    $otherAccess = @($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -ne $graphAppId })
    $existingGraph = @($app.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $graphAppId } |
        ForEach-Object { $_.ResourceAccess })
    $missing = @($roles | Where-Object { $existingGraph.Id -notcontains $_.Id })

    if ($missing.Count -gt 0 -and $PSCmdlet.ShouldProcess($DisplayName, "Add Graph permissions: $($missing.Value -join ', ')")) {
        $mergedGraph = @{
            ResourceAppId  = $graphAppId
            ResourceAccess = @($existingGraph | ForEach-Object { @{ Id = $_.Id; Type = $_.Type } }) +
                @($missing | ForEach-Object { @{ Id = $_.Id; Type = 'Role' } })
        }
        Update-MgApplication -ApplicationId $app.Id -RequiredResourceAccess (@($otherAccess) + @($mergedGraph))
    }
}

# Upload the certificate unless the app already has it
if ($certificate) {
    $existingKeys = @($app.KeyCredentials)
    $alreadyAdded = $existingKeys | Where-Object {
        $_.CustomKeyIdentifier -and ([System.BitConverter]::ToString($_.CustomKeyIdentifier) -replace '-', '') -eq $certificate.Thumbprint
    }

    if ($alreadyAdded) {
        Write-Verbose "Certificate $($certificate.Thumbprint) is already on the app registration."
    }
    elseif ($PSCmdlet.ShouldProcess($DisplayName, "Add certificate $($certificate.Thumbprint)")) {
        # Graph replaces the whole collection, so existing keys are sent back by KeyId to keep them
        $keys = @($existingKeys | ForEach-Object {
                @{
                    KeyId               = $_.KeyId
                    Type                = $_.Type
                    Usage               = $_.Usage
                    DisplayName         = $_.DisplayName
                    CustomKeyIdentifier = $_.CustomKeyIdentifier
                }
            })
        $keys += @{
            Type                = 'AsymmetricX509Cert'
            Usage               = 'Verify'
            Key                 = $certificate.RawData
            DisplayName         = $certificate.Subject
            CustomKeyIdentifier = $certificate.GetCertHash()
        }
        Update-MgApplication -ApplicationId $app.Id -KeyCredentials $keys
    }
}

# Find or create the service principal
$sp = Get-MgServicePrincipal -Filter "appId eq '$($app.AppId)'"
if (-not $sp -and $PSCmdlet.ShouldProcess($DisplayName, 'Create service principal')) {
    $sp = New-MgServicePrincipal -AppId $app.AppId
}

# Grant admin consent by assigning each app role to the service principal
if ($sp) {
    $assigned = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -All |
        Where-Object { $_.ResourceId -eq $graphSp.Id })

    foreach ($role in $roles) {
        if ($assigned.AppRoleId -contains $role.Id) {
            continue
        }
        if ($PSCmdlet.ShouldProcess($DisplayName, "Grant admin consent for $($role.Value)")) {
            Invoke-WithRetry {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -PrincipalId $sp.Id -ResourceId $graphSp.Id -AppRoleId $role.Id
            } | Out-Null
        }
    }
}

# Create the client secret the runbooks use
$credential = $null
$secretExpires = $null
if ($CreateClientSecret -and $PSCmdlet.ShouldProcess($DisplayName, 'Create client secret')) {
    $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential @{
        DisplayName = 'Intune Automated Reporting runbooks'
        EndDateTime = (Get-Date).ToUniversalTime().AddMonths($SecretValidityMonths)
    }
    $credential = New-Object System.Management.Automation.PSCredential(
        $app.AppId, (ConvertTo-SecureString -String $secret.SecretText -AsPlainText -Force))
    $secretExpires = $secret.EndDateTime
}

[pscustomobject]@{
    TenantId              = $TenantId
    DisplayName           = $DisplayName
    ClientId              = $app.AppId
    ApplicationObjectId   = $app.Id
    ServicePrincipalId    = $sp.Id
    Permissions           = @($roles.Value)
    CertificateThumbprint = if ($certificate) { $certificate.Thumbprint } else { $null }
    GraphApiCredential    = $credential
    SecretExpires         = $secretExpires
}
