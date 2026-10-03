# Intune Automated Reporting
Azure Automation runbook scripts to grab Intune data and output CSV to Storage Account containers.

## Set up

Setting up the runbook scripts to automatically grab Intune data ouput CSVs to Azure blob storage accounts can be done in six simple steps.

1. Graph API Registration
1. Create a Resource Group
1. Create a Storage Account
1. Create an Automation Account
1. Set up Automation Runbooks
1. Import data into PowerBI

Check out the step-by-step guide to learn more.

### Automate the Graph API registration

The scripts in the `Automation` folder replace the manual app registration in step 1. They need PowerShell 7 (or Windows PowerShell 5.1) and the Microsoft Graph PowerShell SDK, run as an admin who can create app registrations and grant admin consent (Cloud Application Administrator plus Privileged Role Administrator, or Global Administrator).

```powershell
Install-Module Microsoft.Graph.Authentication, Microsoft.Graph.Applications -Scope CurrentUser

# Create the app registration, grant admin consent and create the client secret the runbooks use
$app = .\Automation\New-IntuneReportingTenant.ps1 -TenantId contoso.onmicrosoft.com -CreateClientSecret
$app | Format-List TenantId, ClientId, Permissions, SecretExpires
```

`New-IntuneReportingTenant.ps1` creates a single-tenant app registration named *Intune Automated Reporting* (change it with `-DisplayName`), adds the read-only Microsoft Graph application permissions the runbooks use, creates the service principal and grants admin consent. Running it again is safe: it reuses the app with the same name and only adds what is missing.

| Permission | Used by |
| --- | --- |
| DeviceManagementApps.Read.All | AppInstallStates.ps1 |
| DeviceManagementConfiguration.Read.All | DeviceConfigurationGroups.ps1, DeviceConfigurationStates.ps1, ServicingRings.ps1 |
| DeviceManagementManagedDevices.Read.All | Antivirus.ps1, Malware.ps1, ServicingRings.ps1 |
| Group.Read.All | DeviceConfigurationGroups.ps1, DeviceConfigurationStates.ps1, ServicingRings.ps1 |
| User.Read.All | AppInstallStates.ps1, ServicingRings.ps1 |

With `-CreateClientSecret` the output's `GraphApiCredential` holds the client ID as user name and the new secret as password, which is exactly the `GraphApi` credential the runbooks read with `Get-AutomationPSCredential`. Store it in your Automation Account (requires the Az.Automation module):

```powershell
New-AzAutomationCredential -ResourceGroupName <resource group> -AutomationAccountName <automation account> `
    -Name GraphApi -Value $app.GraphApiCredential
New-AzAutomationVariable -ResourceGroupName <resource group> -AutomationAccountName <automation account> `
    -Name tenantID -Value $app.TenantId -Encrypted $false
```

The secret is only shown once, so store it before closing the session. Run the script with `-CreateClientSecret` again to add a new secret before the old one expires (`SecretExpires`).

#### Certificate authentication

To authenticate with a certificate instead of (or as well as) a secret, create one with `New-IntuneReportingCertificate.ps1` and upload it to the app:

```powershell
$cert = .\Automation\New-IntuneReportingCertificate.ps1 -OutputPath .\certs -ValidityMonths 12
.\Automation\New-IntuneReportingTenant.ps1 -TenantId contoso.onmicrosoft.com -CertificatePath $cert.CerPath
```

This writes `IntuneAutomatedReporting.cer` (public key, uploaded to the app) and `IntuneAutomatedReporting.pfx` (private key) to the output folder. Pass `-PfxPassword` to choose the .pfx password; otherwise one is generated and returned in `$cert.PfxPassword`. Import the .pfx into the Automation Account's certificates, or into the certificate store of the machine that connects with `Connect-MgGraph -ClientId <client ID> -TenantId <tenant> -CertificateThumbprint $cert.Thumbprint`. Keep the .pfx out of source control.

Both scripts support `-WhatIf` to preview the changes and `-Verbose` for details. See `Get-Help .\Automation\New-IntuneReportingTenant.ps1 -Full` for every parameter.

#### Running the tests

The Pester tests mock every Microsoft Graph call, so they run without a tenant or the Graph modules:

```powershell
Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser
Invoke-Pester ./Tests -Output Detailed
```

The tests need PowerShell 7.

## Create effective visuals in PowerBI

### Applications

This visual uses the following scripts.
* **AppInstallStates.ps1**
* **ServicingRings.ps1**

![Applications visualization](/images/appsbydepartment.png)

### Servicing

The following visuals use the **ServicingRings.ps1** script.

![Servicing visualization](/images/servicingoverview.png)

<br>

![Historical servicing visualization](/images/historicalservicingcomparison.png)

