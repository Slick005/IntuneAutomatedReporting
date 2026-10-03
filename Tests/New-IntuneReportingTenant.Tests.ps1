#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }

BeforeAll {
    $TenantScript = Join-Path $PSScriptRoot '..' 'Automation' 'New-IntuneReportingTenant.ps1'
    $CertScript = Join-Path $PSScriptRoot '..' 'Automation' 'New-IntuneReportingCertificate.ps1'
    $FxGraphAppId = '00000003-0000-0000-c000-000000000000'

    # Stubs with the parameters the script uses, so the tests run without the Microsoft.Graph
    # modules installed. Every test mocks them; Pester copies these signatures into the mocks.
    $stubs = @{
        'Get-MgContext'                           = ''
        'Connect-MgGraph'                         = '[string]$TenantId, [string[]]$Scopes, [switch]$NoWelcome'
        'Get-MgServicePrincipal'                  = '[string]$Filter'
        'New-MgServicePrincipal'                  = '[string]$AppId'
        'Get-MgApplication'                       = '[string]$Filter'
        'New-MgApplication'                       = '[string]$DisplayName, [string]$SignInAudience, [object[]]$RequiredResourceAccess'
        'Update-MgApplication'                    = '[string]$ApplicationId, [object[]]$RequiredResourceAccess, [object[]]$KeyCredentials'
        'Add-MgApplicationPassword'               = '[string]$ApplicationId, [hashtable]$PasswordCredential'
        'Get-MgServicePrincipalAppRoleAssignment' = '[string]$ServicePrincipalId, [switch]$All'
        'New-MgServicePrincipalAppRoleAssignment' = '[string]$ServicePrincipalId, [string]$PrincipalId, [string]$ResourceId, [string]$AppRoleId'
    }
    foreach ($name in $stubs.Keys) {
        if (-not (Get-Command $name -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$name" -Value ([scriptblock]::Create("param($($stubs[$name]))"))
        }
    }

    # Fixtures use an Fx prefix: mock bodies run while the script under test is on the call stack, and
    # PowerShell variable lookup is dynamic, so a name like $roles or $sp would pick up the script's own.
    function New-Role ([string]$Value) {
        [pscustomobject]@{ Id = [guid]::NewGuid().ToString(); Value = $Value; AllowedMemberTypes = @('Application') }
    }

    $FxRoles = @{}
    foreach ($value in 'DeviceManagementApps.Read.All', 'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All', 'Group.Read.All', 'User.Read.All', 'Mail.Send') {
        $FxRoles[$value] = New-Role $value
    }
    # A delegated-only scope with the same shape, which must never be picked as an app permission
    $FxDelegatedOnly = [pscustomobject]@{ Id = [guid]::NewGuid().ToString(); Value = 'Delegated.Only'; AllowedMemberTypes = @('User') }

    $FxGraphSp = [pscustomobject]@{
        Id       = 'graph-sp-id'
        AppId    = $FxGraphAppId
        AppRoles = @($FxRoles.Values) + $FxDelegatedOnly
    }
    $FxDefaultPermissions = @('DeviceManagementApps.Read.All', 'DeviceManagementConfiguration.Read.All',
        'DeviceManagementManagedDevices.Read.All', 'Group.Read.All', 'User.Read.All')
}

Describe 'New-IntuneReportingTenant.ps1' {

    BeforeEach {
        $FxNewApp = [pscustomobject]@{
            Id                     = 'app-object-id'
            AppId                  = 'app-client-id'
            DisplayName            = 'Intune Automated Reporting'
            RequiredResourceAccess = @()
            KeyCredentials         = @()
        }
        $FxSp = [pscustomobject]@{ Id = 'app-sp-id'; AppId = 'app-client-id' }

        Mock Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-guid'; Account = 'admin@contoso.onmicrosoft.com'; Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All') } }
        Mock Connect-MgGraph { }
        Mock Get-MgServicePrincipal { $FxGraphSp } -ParameterFilter { $Filter -eq "appId eq '00000003-0000-0000-c000-000000000000'" }
        Mock Get-MgServicePrincipal { $null } -ParameterFilter { $Filter -eq "appId eq 'app-client-id'" }
        Mock New-MgServicePrincipal { $FxSp }
        Mock Get-MgApplication { }
        Mock New-MgApplication { $FxNewApp }
        Mock Update-MgApplication { }
        Mock Get-MgServicePrincipalAppRoleAssignment { }
        Mock New-MgServicePrincipalAppRoleAssignment { [pscustomobject]@{ Id = 'assignment' } }
        Mock Add-MgApplicationPassword { [pscustomobject]@{ SecretText = 's3cr3t-value'; EndDateTime = [datetime]'2027-10-03' } }
        Mock Start-Sleep { }
    }

    Context 'Connecting to Microsoft Graph' {

        It 'reuses an existing session for the same tenant' {
            & $TenantScript -TenantId 'contoso.onmicrosoft.com' | Out-Null
            Should -Invoke Connect-MgGraph -Times 0 -Exactly
        }

        It 'connects with the admin scopes when there is no session' {
            $state = @{ Calls = 0 }
            Mock Get-MgContext {
                $state.Calls++
                if ($state.Calls -gt 1) { [pscustomobject]@{ TenantId = 'tenant-guid'; Scopes = @() } }
            }

            $result = & $TenantScript -TenantId 'contoso.onmicrosoft.com'

            Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter {
                $TenantId -eq 'contoso.onmicrosoft.com' -and
                $Scopes -contains 'Application.ReadWrite.All' -and $Scopes -contains 'AppRoleAssignment.ReadWrite.All'
            }
            $result.TenantId | Should -Be 'tenant-guid'
        }

        It 'reconnects when the session is for another tenant' {
            Mock Get-MgContext { [pscustomobject]@{ TenantId = 'other-tenant'; Account = 'admin@fabrikam.com'; Scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All') } }
            & $TenantScript -TenantId 'contoso.onmicrosoft.com' | Out-Null
            Should -Invoke Connect-MgGraph -Times 1 -Exactly
        }

        It 'reconnects when the session is missing a scope' {
            Mock Get-MgContext { [pscustomobject]@{ TenantId = 'tenant-guid'; Scopes = @('Application.ReadWrite.All') } }
            & $TenantScript -TenantId 'tenant-guid' | Out-Null
            Should -Invoke Connect-MgGraph -Times 1 -Exactly
        }
    }

    Context 'New app registration' {

        It 'creates a single-tenant app with the default Graph application permissions' {
            & $TenantScript -TenantId 'tenant-guid' | Out-Null

            Should -Invoke New-MgApplication -Times 1 -Exactly -ParameterFilter {
                $access = $RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000003-0000-0000-c000-000000000000' }
                $DisplayName -eq 'Intune Automated Reporting' -and
                $SignInAudience -eq 'AzureADMyOrg' -and
                @($access.ResourceAccess).Count -eq 5 -and
                -not ($access.ResourceAccess | Where-Object { $_.Type -ne 'Role' })
            }
        }

        It 'creates the service principal and grants admin consent for every permission' {
            & $TenantScript -TenantId 'tenant-guid' | Out-Null

            Should -Invoke New-MgServicePrincipal -Times 1 -Exactly -ParameterFilter { $AppId -eq 'app-client-id' }
            foreach ($permission in $FxDefaultPermissions) {
                $roleId = $FxRoles[$permission].Id
                Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 1 -Exactly -ParameterFilter {
                    $ServicePrincipalId -eq 'app-sp-id' -and $PrincipalId -eq 'app-sp-id' -and
                    $ResourceId -eq 'graph-sp-id' -and $AppRoleId -eq $roleId
                }
            }
        }

        It 'returns the app details' {
            $result = & $TenantScript -TenantId 'tenant-guid' -DisplayName 'Intune Automated Reporting'

            $result.ClientId | Should -Be 'app-client-id'
            $result.ApplicationObjectId | Should -Be 'app-object-id'
            $result.ServicePrincipalId | Should -Be 'app-sp-id'
            $result.TenantId | Should -Be 'tenant-guid'
            $result.Permissions | Should -Be $FxDefaultPermissions
            $result.GraphApiCredential | Should -BeNullOrEmpty
            $result.CertificateThumbprint | Should -BeNullOrEmpty
        }

        It 'uses a custom display name and permission list' {
            $result = & $TenantScript -TenantId 'tenant-guid' -DisplayName 'Reporting Test' -Permissions 'User.Read.All'

            Should -Invoke New-MgApplication -Times 1 -Exactly -ParameterFilter { $DisplayName -eq 'Reporting Test' }
            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 1 -Exactly
            $result.Permissions | Should -Be @('User.Read.All')
        }

        It 'retries a role assignment while the new service principal replicates' {
            $state = @{ Attempts = 0 }
            Mock New-MgServicePrincipalAppRoleAssignment {
                $state.Attempts++
                if ($state.Attempts -eq 1) { throw 'Resource does not exist or one of its queried reference-property objects are not present.' }
                [pscustomobject]@{ Id = 'assignment' }
            }

            & $TenantScript -TenantId 'tenant-guid' -Permissions 'User.Read.All' | Out-Null

            $state.Attempts | Should -Be 2
            Should -Invoke Start-Sleep -Times 1 -Exactly
        }

        It 'fails after the retries are used up' {
            Mock New-MgServicePrincipalAppRoleAssignment { throw 'still not there' }
            { & $TenantScript -TenantId 'tenant-guid' -Permissions 'User.Read.All' } | Should -Throw '*still not there*'
            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 5 -Exactly
        }
    }

    Context 'Existing app registration' {

        BeforeEach {
            $FxExistingApp = [pscustomobject]@{
                Id                     = 'app-object-id'
                AppId                  = 'app-client-id'
                DisplayName            = 'Intune Automated Reporting'
                RequiredResourceAccess = @(
                    [pscustomobject]@{
                        ResourceAppId  = '00000003-0000-0000-c000-000000000000'
                        ResourceAccess = @(
                            [pscustomobject]@{ Id = $FxRoles['User.Read.All'].Id; Type = 'Role' },
                            [pscustomobject]@{ Id = $FxRoles['Mail.Send'].Id; Type = 'Role' }
                        )
                    },
                    [pscustomobject]@{
                        ResourceAppId  = 'other-api'
                        ResourceAccess = @([pscustomobject]@{ Id = 'other-role'; Type = 'Role' })
                    }
                )
                KeyCredentials         = @()
            }
            Mock Get-MgApplication { $FxExistingApp }
            Mock Get-MgServicePrincipal { $FxSp } -ParameterFilter { $Filter -eq "appId eq 'app-client-id'" }
            Mock Get-MgServicePrincipalAppRoleAssignment {
                [pscustomobject]@{ ResourceId = 'graph-sp-id'; AppRoleId = $FxRoles['User.Read.All'].Id },
                [pscustomobject]@{ ResourceId = 'other-sp-id'; AppRoleId = $FxRoles['Group.Read.All'].Id }
            }
        }

        It 'does not create a second app or service principal' {
            & $TenantScript -TenantId 'tenant-guid' | Out-Null
            Should -Invoke New-MgApplication -Times 0 -Exactly
            Should -Invoke New-MgServicePrincipal -Times 0 -Exactly
        }

        It 'adds only the missing permissions and keeps the existing ones' {
            & $TenantScript -TenantId 'tenant-guid' | Out-Null

            Should -Invoke Update-MgApplication -Times 1 -Exactly -ParameterFilter {
                $graph = $RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq '00000003-0000-0000-c000-000000000000' }
                $other = $RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq 'other-api' }
                $ids = @($graph.ResourceAccess | ForEach-Object { $_.Id })
                $ApplicationId -eq 'app-object-id' -and
                $ids.Count -eq 6 -and
                $ids -contains $FxRoles['Mail.Send'].Id -and
                ($ids | Select-Object -Unique).Count -eq 6 -and
                $other.ResourceAccess[0].Id -eq 'other-role'
            }
        }

        It 'grants consent only for roles that are not assigned on Graph yet' {
            & $TenantScript -TenantId 'tenant-guid' | Out-Null

            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 4 -Exactly
            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 0 -Exactly -ParameterFilter { $AppRoleId -eq $FxRoles['User.Read.All'].Id }
            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 1 -Exactly -ParameterFilter { $AppRoleId -eq $FxRoles['Group.Read.All'].Id }
        }

        It 'changes nothing when everything is already in place' {
            Mock Get-MgServicePrincipalAppRoleAssignment {
                $FxDefaultPermissions | ForEach-Object { [pscustomobject]@{ ResourceId = 'graph-sp-id'; AppRoleId = $FxRoles[$_].Id } }
            }
            $FxExistingApp.RequiredResourceAccess[0].ResourceAccess = @(
                $FxDefaultPermissions | ForEach-Object { [pscustomobject]@{ Id = $FxRoles[$_].Id; Type = 'Role' } })

            & $TenantScript -TenantId 'tenant-guid' | Out-Null

            Should -Invoke Update-MgApplication -Times 0 -Exactly
            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 0 -Exactly
        }

        It 'escapes quotes in the display name filter' {
            & $TenantScript -TenantId 'tenant-guid' -DisplayName "Erik's Reporting" | Out-Null
            Should -Invoke Get-MgApplication -Times 1 -Exactly -ParameterFilter { $Filter -eq "displayName eq 'Erik''s Reporting'" }
        }

        It 'stops when several apps share the display name' {
            Mock Get-MgApplication { $FxExistingApp, $FxExistingApp }
            { & $TenantScript -TenantId 'tenant-guid' } | Should -Throw '*Found 2 app registrations*'
            Should -Invoke Update-MgApplication -Times 0 -Exactly
        }
    }

    Context 'Certificate' {

        BeforeAll {
            $FxCert = & $CertScript -OutputPath (Join-Path $TestDrive 'certs') -ValidityMonths 1
        }

        It 'uploads the certificate to a new app' {
            $result = & $TenantScript -TenantId 'tenant-guid' -CertificatePath $FxCert.CerPath

            Should -Invoke Update-MgApplication -Times 1 -Exactly -ParameterFilter {
                $key = @($KeyCredentials)
                $key.Count -eq 1 -and
                $key[0].Type -eq 'AsymmetricX509Cert' -and $key[0].Usage -eq 'Verify' -and
                ([System.BitConverter]::ToString($key[0].CustomKeyIdentifier) -replace '-', '') -eq $FxCert.Thumbprint -and
                $key[0].Key.Length -gt 0
            }
            $result.CertificateThumbprint | Should -Be $FxCert.Thumbprint
        }

        It 'keeps existing certificates when adding a new one' {
            $FxNewApp.KeyCredentials = @([pscustomobject]@{
                    KeyId = 'old-key'; Type = 'AsymmetricX509Cert'; Usage = 'Verify'; DisplayName = 'CN=Old'
                    CustomKeyIdentifier = [byte[]](1, 2, 3)
                })

            & $TenantScript -TenantId 'tenant-guid' -CertificatePath $FxCert.CerPath | Out-Null

            Should -Invoke Update-MgApplication -Times 1 -Exactly -ParameterFilter {
                @($KeyCredentials).Count -eq 2 -and $KeyCredentials[0].KeyId -eq 'old-key' -and -not $KeyCredentials[0].Key
            }
        }

        It 'does not upload a certificate the app already has' {
            $FxNewApp.KeyCredentials = @([pscustomobject]@{
                    KeyId = 'same-key'; Type = 'AsymmetricX509Cert'; Usage = 'Verify'
                    CustomKeyIdentifier = [System.Convert]::FromHexString($FxCert.Thumbprint)
                })

            & $TenantScript -TenantId 'tenant-guid' -CertificatePath $FxCert.CerPath | Out-Null

            Should -Invoke Update-MgApplication -Times 0 -Exactly
        }

        It 'fails before connecting when the certificate file is missing' {
            { & $TenantScript -TenantId 'tenant-guid' -CertificatePath (Join-Path $TestDrive 'missing.cer') } |
                Should -Throw '*was not found*'
            Should -Invoke Connect-MgGraph -Times 0 -Exactly
            Should -Invoke Get-MgContext -Times 0 -Exactly
        }
    }

    Context 'Client secret' {

        It 'returns the GraphApi credential the runbooks use' {
            $result = & $TenantScript -TenantId 'tenant-guid' -CreateClientSecret -SecretValidityMonths 6

            Should -Invoke Add-MgApplicationPassword -Times 1 -Exactly -ParameterFilter {
                $ApplicationId -eq 'app-object-id' -and
                $PasswordCredential.EndDateTime -gt (Get-Date).ToUniversalTime().AddMonths(5) -and
                $PasswordCredential.EndDateTime -lt (Get-Date).ToUniversalTime().AddMonths(7)
            }
            $result.GraphApiCredential.UserName | Should -Be 'app-client-id'
            $result.GraphApiCredential.GetNetworkCredential().Password | Should -Be 's3cr3t-value'
            $result.SecretExpires | Should -Be ([datetime]'2027-10-03')
        }

        It 'does not create a secret unless asked' {
            & $TenantScript -TenantId 'tenant-guid' | Out-Null
            Should -Invoke Add-MgApplicationPassword -Times 0 -Exactly
        }
    }

    Context 'Validation and WhatIf' {

        It 'rejects an unknown permission before changing anything' {
            { & $TenantScript -TenantId 'tenant-guid' -Permissions 'Not.A.Permission' } | Should -Throw "*'Not.A.Permission' is not a Microsoft Graph application permission*"
            Should -Invoke New-MgApplication -Times 0 -Exactly
        }

        It 'rejects a delegated-only permission' {
            { & $TenantScript -TenantId 'tenant-guid' -Permissions 'Delegated.Only' } | Should -Throw '*is not a Microsoft Graph application permission*'
        }

        It 'fails when the Microsoft Graph service principal is missing' {
            Mock Get-MgServicePrincipal { $null } -ParameterFilter { $Filter -eq "appId eq '00000003-0000-0000-c000-000000000000'" }
            { & $TenantScript -TenantId 'tenant-guid' } | Should -Throw '*Microsoft Graph service principal was not found*'
        }

        It 'makes no changes with -WhatIf' {
            & $TenantScript -TenantId 'tenant-guid' -CreateClientSecret -WhatIf | Out-Null

            Should -Invoke New-MgApplication -Times 0 -Exactly
            Should -Invoke Update-MgApplication -Times 0 -Exactly
            Should -Invoke New-MgServicePrincipal -Times 0 -Exactly
            Should -Invoke New-MgServicePrincipalAppRoleAssignment -Times 0 -Exactly
            Should -Invoke Add-MgApplicationPassword -Times 0 -Exactly
        }
    }
}

Describe 'New-IntuneReportingCertificate.ps1' {

    It 'writes a .cer and a password-protected .pfx for the same certificate' {
        $out = Join-Path $TestDrive 'pair'
        $password = ConvertTo-SecureString 'P@ssw0rd-for-tests' -AsPlainText -Force

        $result = & $CertScript -OutputPath $out -Name 'Reporting' -PfxPassword $password -ValidityMonths 6

        $result.CerPath | Should -Exist
        $result.PfxPath | Should -Exist
        $result.Subject | Should -Be 'CN=Reporting'
        $result.PasswordGenerated | Should -BeFalse

        $cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($result.CerPath)
        $pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($result.PfxPath, $password)
        $cer.Thumbprint | Should -Be $result.Thumbprint
        $pfx.Thumbprint | Should -Be $result.Thumbprint
        $cer.HasPrivateKey | Should -BeFalse
        $pfx.HasPrivateKey | Should -BeTrue
        $cer.NotAfter | Should -BeGreaterThan (Get-Date).AddMonths(5)
        $cer.NotAfter | Should -BeLessThan (Get-Date).AddMonths(7)
    }

    It 'creates a client authentication certificate with the requested key size' {
        $result = & $CertScript -OutputPath (Join-Path $TestDrive 'eku') -KeyLength 3072 -Subject 'CN=Custom Subject'
        $cer = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($result.CerPath)

        $cer.Subject | Should -Be 'CN=Custom Subject'
        ([System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPublicKey($cer).ExportParameters($false).Modulus.Length * 8) | Should -Be 3072
        $eku = $cer.Extensions | Where-Object { $_ -is [System.Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension] }
        $eku.EnhancedKeyUsages.Value | Should -Contain '1.3.6.1.5.5.7.3.2'
    }

    It 'generates and returns a PFX password when none is given' {
        $result = & $CertScript -OutputPath (Join-Path $TestDrive 'generated')

        $result.PasswordGenerated | Should -BeTrue
        $result.PfxPassword | Should -BeOfType [securestring]
        $pfx = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($result.PfxPath, $result.PfxPassword)
        $pfx.HasPrivateKey | Should -BeTrue
    }

    It 'refuses to overwrite existing files unless -Force is used' {
        $out = Join-Path $TestDrive 'overwrite'
        $first = & $CertScript -OutputPath $out

        { & $CertScript -OutputPath $out } | Should -Throw '*already exists*'

        $second = & $CertScript -OutputPath $out -Force
        $second.Thumbprint | Should -Not -Be $first.Thumbprint
    }

    It 'writes relative output paths under the current location' {
        Push-Location $TestDrive
        try {
            $result = & $CertScript -OutputPath 'relative'
            $result.CerPath | Should -Be (Join-Path $TestDrive 'relative' 'IntuneAutomatedReporting.cer')
            $result.CerPath | Should -Exist
        }
        finally {
            Pop-Location
        }
    }

    It 'writes nothing with -WhatIf' {
        $out = Join-Path $TestDrive 'whatif'
        & $CertScript -OutputPath $out -WhatIf | Should -BeNullOrEmpty
        $out | Should -Not -Exist
    }
}
