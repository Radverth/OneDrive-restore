#Requires -Version 7.0
<#
.SYNOPSIS
    Stage 1 - one-time Azure AD app registration and certificate setup.

.DESCRIPTION
    Creates the app-only identity that stages 2-5 authenticate with. Run once,
    as a Global Administrator, on the workstation that will run the toolkit -
    the certificate's private key stays in that account's certificate store.
#>

Set-StrictMode -Version Latest
Import-Module (Join-Path $PSScriptRoot 'Common.psm1') -DisableNameChecking

# Well-known appId of the Microsoft Graph service principal.
$script:GraphAppId = '00000003-0000-0000-c000-000000000000'

# Application (not delegated) permissions the toolkit needs.
$script:RequiredPermission = @(
    'Files.ReadWrite.All'
    'Sites.ReadWrite.All'
    'User.Read.All'
)

# Office 365 SharePoint Online. Only needed to download recycle bin contents:
# Graph cannot restore a OneDrive for Business item, so that path goes through
# SharePoint REST, which needs its own application permission.
$script:SharePointAppId = '00000003-0000-0ff1-ce00-000000000000'
$script:SharePointPermission = @('Sites.FullControl.All')

function Test-AppRegistrationPrerequisite {
    [CmdletBinding()]
    param()

    $ok = Test-ToolkitPrerequisite -RequiredModule @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')

    if (-not $IsWindows) {
        Write-ToolkitLog 'Certificate creation uses New-SelfSignedCertificate and the Windows certificate store, so stage 1 must run on Windows.' -Level ERROR
        Write-ToolkitLog 'On other platforms, create the app registration and upload a certificate manually, then fill in TenantId/AppId/CertificateThumbprint in config/toolkit-config.json.' -Level WARN
        $ok = $false
    }

    return $ok
}

function New-ToolkitCertificate {
    <#
    .SYNOPSIS
        Creates the self-signed signing certificate and exports its public .cer.
    .OUTPUTS
        PSCustomObject with Certificate, Thumbprint, Subject, PublicKeyPath.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Subject,
        [ValidateRange(1, 5)][int]$ValidYears = 2
    )

    $ctx = Initialize-ToolkitEnvironment

    Write-ToolkitLog ('Creating self-signed certificate "{0}" valid for {1} year(s)...' -f $Subject, $ValidYears) -Level INFO

    $certificate = New-SelfSignedCertificate `
        -Subject ('CN={0}' -f $Subject) `
        -CertStoreLocation 'Cert:\CurrentUser\My' `
        -KeyExportPolicy Exportable `
        -KeySpec Signature `
        -KeyLength 2048 `
        -KeyAlgorithm RSA `
        -HashAlgorithm SHA256 `
        -NotAfter (Get-Date).AddYears($ValidYears) `
        -ErrorAction Stop

    $publicKeyPath = Join-Path $ctx.ConfigDir ('{0}.cer' -f ($Subject -replace '[^A-Za-z0-9._-]', '-'))
    Export-Certificate -Cert $certificate -FilePath $publicKeyPath -Force | Out-Null

    Write-ToolkitLog ('Certificate created. Thumbprint {0}, expires {1:yyyy-MM-dd}.' -f $certificate.Thumbprint, $certificate.NotAfter) -Level SUCCESS
    Write-ToolkitLog ('Public key exported to {0}' -f $publicKeyPath) -Level INFO

    return [pscustomobject]@{
        Certificate   = $certificate
        Thumbprint    = $certificate.Thumbprint
        Subject       = $Subject
        PublicKeyPath = $publicKeyPath
        NotAfter      = $certificate.NotAfter
    }
}

function New-ResourceAccessEntry {
    <#
    .SYNOPSIS
        Maps permission names onto the application-role IDs a resource exposes.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ServicePrincipal,
        [Parameter(Mandatory)][string[]]$Permission
    )

    $access = @()
    foreach ($name in $Permission) {
        $role = $ServicePrincipal.AppRoles |
            Where-Object { $_.Value -eq $name -and $_.AllowedMemberTypes -contains 'Application' } |
            Select-Object -First 1

        if (-not $role) {
            throw ('{0} does not expose an application role called {1}.' -f $ServicePrincipal.DisplayName, $name)
        }
        $access += @{ id = $role.Id; type = 'Role' }
    }
    return $access
}

function New-ToolkitAppRegistration {
    <#
    .SYNOPSIS
        Creates the application object with the certificate already attached.
    .DESCRIPTION
        The certificate goes in as part of New-MgApplication rather than through
        New-MgApplicationKeyCredential: the addKey action that cmdlet calls needs
        proof-of-possession signed with an existing credential, which a brand-new
        application does not have. New-MgApplicationKeyCredential is the right
        cmdlet for later certificate rollover (see Add-ToolkitAppCertificate).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$DisplayName,
        [Parameter(Mandatory)]$Certificate,
        [Parameter(Mandatory)]$GraphServicePrincipal,
        $SharePointServicePrincipal
    )

    $requiredResourceAccess = @(@{
        ResourceAppId  = $script:GraphAppId
        ResourceAccess = (New-ResourceAccessEntry -ServicePrincipal $GraphServicePrincipal -Permission $script:RequiredPermission)
    })

    if ($SharePointServicePrincipal) {
        $requiredResourceAccess += @{
            ResourceAppId  = $script:SharePointAppId
            ResourceAccess = (New-ResourceAccessEntry -ServicePrincipal $SharePointServicePrincipal -Permission $script:SharePointPermission)
        }
    }

    $keyCredential = @{
        Type  = 'AsymmetricX509Cert'
        Usage = 'Verify'
        Key   = $Certificate.GetRawCertData()
        DisplayName = ('CN={0}' -f $Certificate.Subject)
    }

    Write-ToolkitLog ('Creating application registration "{0}"...' -f $DisplayName) -Level INFO

    $application = New-MgApplication `
        -DisplayName $DisplayName `
        -SignInAudience 'AzureADMyOrg' `
        -KeyCredentials @($keyCredential) `
        -RequiredResourceAccess $requiredResourceAccess `
        -ErrorAction Stop

    Write-ToolkitLog ('Application created. AppId {0} (object id {1}).' -f $application.AppId, $application.Id) -Level SUCCESS
    return $application
}

function Add-ToolkitAppCertificate {
    <#
    .SYNOPSIS
        Adds a new certificate to an existing registration (rollover before expiry).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ApplicationObjectId,
        [Parameter(Mandatory)]$Certificate
    )

    $existing = Get-MgApplication -ApplicationId $ApplicationObjectId -ErrorAction Stop
    $keyCredentials = @()
    foreach ($key in $existing.KeyCredentials) {
        $keyCredentials += @{
            Type        = $key.Type
            Usage       = $key.Usage
            Key         = $key.Key
            DisplayName = $key.DisplayName
        }
    }
    $keyCredentials += @{
        Type        = 'AsymmetricX509Cert'
        Usage       = 'Verify'
        Key         = $Certificate.GetRawCertData()
        DisplayName = ('CN={0}' -f $Certificate.Subject)
    }

    Update-MgApplication -ApplicationId $ApplicationObjectId -KeyCredentials $keyCredentials -ErrorAction Stop
    Write-ToolkitLog ('Certificate {0} added to existing registration.' -f $Certificate.Thumbprint) -Level SUCCESS
}

function Grant-ToolkitAdminConsent {
    <#
    .SYNOPSIS
        Grants admin consent by assigning each application role to the app's service principal.
    .OUTPUTS
        PSCustomObject with Granted / Failed permission lists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ServicePrincipal,
        [Parameter(Mandatory)]$ResourceServicePrincipal,
        [string[]]$Permission = $script:RequiredPermission
    )

    $granted = @()
    $failed = @()

    $existingAssignment = @()
    try {
        $existingAssignment = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipal.Id -All -ErrorAction Stop)
    }
    catch {
        Write-ToolkitLog ('Could not read existing role assignments: {0}' -f $_.Exception.Message) -Level WARN
    }

    foreach ($permission in $Permission) {
        $role = $ResourceServicePrincipal.AppRoles |
            Where-Object { $_.Value -eq $permission -and $_.AllowedMemberTypes -contains 'Application' } |
            Select-Object -First 1

        if (-not $role) {
            $failed += $permission
            Write-ToolkitLog ('No application role found for {0}.' -f $permission) -Level ERROR
            continue
        }

        if ($existingAssignment | Where-Object { $_.AppRoleId -eq $role.Id -and $_.ResourceId -eq $ResourceServicePrincipal.Id }) {
            Write-ToolkitLog ('{0} is already consented.' -f $permission) -Level INFO
            $granted += $permission
            continue
        }

        try {
            New-MgServicePrincipalAppRoleAssignment `
                -ServicePrincipalId $ServicePrincipal.Id `
                -PrincipalId $ServicePrincipal.Id `
                -ResourceId $ResourceServicePrincipal.Id `
                -AppRoleId $role.Id `
                -ErrorAction Stop | Out-Null

            $granted += $permission
            Write-ToolkitLog ('Consent granted for {0}.' -f $permission) -Level SUCCESS
        }
        catch {
            $failed += $permission
            Write-ToolkitLog ('Consent for {0} failed: {1}' -f $permission, $_.Exception.Message) -Level ERROR
        }
    }

    return [pscustomobject]@{ Granted = $granted; Failed = $failed }
}

function Invoke-AppRegistrationSetup {
    <#
    .SYNOPSIS
        Menu option 1 - registers the app, creates the certificate, grants consent.
    #>
    [CmdletBinding()]
    param()

    Write-ToolkitHeader 'Stage 1 - Azure AD App Registration & Certificate (one-time)'

    if (-not (Test-AppRegistrationPrerequisite)) { return }

    $config = Get-ToolkitConfig

    Write-Host ''
    Write-Host 'This step signs you in interactively as a Global Administrator to create' -ForegroundColor Gray
    Write-Host 'the app registration. Every later stage uses the certificate instead.' -ForegroundColor Gray
    Write-Host ''

    $tenantId = Read-ToolkitValue -Prompt 'Tenant ID (or tenant domain)' -Default ([string]$config['TenantId'])
    $appName = Read-ToolkitValue -Prompt 'App registration display name' -Default 'OneDrive Repair Toolkit'
    $years = [int](Read-ToolkitValue -Prompt 'Certificate validity in years (1-2)' -Default '2')
    if ($years -lt 1 -or $years -gt 5) { $years = 2 }

    try {
        Import-Module Microsoft.Graph.Applications -ErrorAction Stop

        Write-ToolkitLog 'Signing in interactively for the bootstrap step...' -Level INFO
        Connect-MgGraph -TenantId $tenantId `
            -Scopes 'Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All', 'Directory.ReadWrite.All' `
            -NoWelcome -ErrorAction Stop

        $graphSp = Get-MgServicePrincipal -Filter ("appId eq '{0}'" -f $script:GraphAppId) -ErrorAction Stop
        if (-not $graphSp) { throw 'Could not resolve the Microsoft Graph service principal in this tenant.' }

        # Downloading recycle bin contents means restoring items, and Graph cannot
        # restore for OneDrive for Business - that path needs SharePoint REST and
        # therefore a SharePoint application permission. It is a broad permission,
        # so it is opt-in rather than granted by default.
        Write-Host ''
        Write-Host '  Downloading recycle bin CONTENTS needs one extra permission:' -ForegroundColor Gray
        Write-Host ('    {0} on Office 365 SharePoint Online' -f ($script:SharePointPermission -join ', ')) -ForegroundColor Gray
        Write-Host '  Without it everything else still works, including the recycle bin' -ForegroundColor Gray
        Write-Host '  inventory - only restoring and downloading deleted files needs it.' -ForegroundColor Gray
        Write-Host ''
        $wantSharePoint = Confirm-ToolkitAction -Prompt 'Include the SharePoint permission?'

        $sharePointSp = $null
        if ($wantSharePoint) {
            $sharePointSp = Get-MgServicePrincipal -Filter ("appId eq '{0}'" -f $script:SharePointAppId) -ErrorAction SilentlyContinue |
                Select-Object -First 1
            if (-not $sharePointSp) {
                Write-ToolkitLog 'The SharePoint Online service principal was not found in this tenant; continuing with Graph permissions only.' -Level WARN
                $wantSharePoint = $false
            }
        }

        $existingApp = Get-MgApplication -Filter ("displayName eq '{0}'" -f ($appName -replace "'", "''")) -ErrorAction SilentlyContinue |
            Select-Object -First 1

        $certInfo = New-ToolkitCertificate -Subject $appName -ValidYears $years

        if ($existingApp) {
            Write-ToolkitLog ('An application called "{0}" already exists (AppId {1}).' -f $appName, $existingApp.AppId) -Level WARN
            if (Confirm-ToolkitAction -Prompt 'Add the new certificate to that existing registration instead of creating a new one?' -DefaultYes) {
                Add-ToolkitAppCertificate -ApplicationObjectId $existingApp.Id -Certificate $certInfo.Certificate

                # An app registered before the recycle bin download existed will not
                # list the SharePoint resource yet, so add it on the way through.
                if ($sharePointSp -and -not ($existingApp.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $script:SharePointAppId })) {
                    $updated = @()
                    foreach ($entry in $existingApp.RequiredResourceAccess) {
                        $updated += @{
                            ResourceAppId  = $entry.ResourceAppId
                            ResourceAccess = @($entry.ResourceAccess | ForEach-Object { @{ id = $_.Id; type = $_.Type } })
                        }
                    }
                    $updated += @{
                        ResourceAppId  = $script:SharePointAppId
                        ResourceAccess = (New-ResourceAccessEntry -ServicePrincipal $sharePointSp -Permission $script:SharePointPermission)
                    }
                    Update-MgApplication -ApplicationId $existingApp.Id -RequiredResourceAccess $updated -ErrorAction Stop
                    Write-ToolkitLog 'Added the SharePoint permission to the existing registration.' -Level SUCCESS
                }

                $application = Get-MgApplication -ApplicationId $existingApp.Id -ErrorAction Stop
            }
            else {
                $appName = Read-ToolkitValue -Prompt 'New display name to use instead'
                $application = New-ToolkitAppRegistration -DisplayName $appName -Certificate $certInfo.Certificate `
                    -GraphServicePrincipal $graphSp -SharePointServicePrincipal $sharePointSp
            }
        }
        else {
            $application = New-ToolkitAppRegistration -DisplayName $appName -Certificate $certInfo.Certificate `
                -GraphServicePrincipal $graphSp -SharePointServicePrincipal $sharePointSp
        }

        $servicePrincipal = Get-MgServicePrincipal -Filter ("appId eq '{0}'" -f $application.AppId) -ErrorAction SilentlyContinue |
            Select-Object -First 1

        if (-not $servicePrincipal) {
            Write-ToolkitLog 'Creating the service principal for the new application...' -Level INFO
            # Directory replication can lag application creation by a few seconds.
            $servicePrincipal = $null
            for ($attempt = 1; $attempt -le 5 -and -not $servicePrincipal; $attempt++) {
                try {
                    $servicePrincipal = New-MgServicePrincipal -AppId $application.AppId -ErrorAction Stop
                }
                catch {
                    Write-ToolkitLog ('Service principal creation attempt {0} failed, retrying: {1}' -f $attempt, $_.Exception.Message) -Level WARN -NoConsole
                    Start-Sleep -Seconds ([math]::Min(5 * $attempt, 20))
                }
            }
            if (-not $servicePrincipal) { throw 'Could not create the service principal for the application.' }
        }

        $consent = Grant-ToolkitAdminConsent -ServicePrincipal $servicePrincipal -ResourceServicePrincipal $graphSp

        if ($sharePointSp) {
            $sharePointConsent = Grant-ToolkitAdminConsent -ServicePrincipal $servicePrincipal `
                -ResourceServicePrincipal $sharePointSp -Permission $script:SharePointPermission
            $consent.Granted += $sharePointConsent.Granted
            $consent.Failed += $sharePointConsent.Failed
        }

        $config['TenantId'] = $tenantId
        $config['AppId'] = $application.AppId
        $config['CertificateThumbprint'] = $certInfo.Thumbprint
        $config['CertificateSubject'] = $certInfo.Subject
        $config['CertificatePath'] = $certInfo.PublicKeyPath
        Save-ToolkitConfig -Config $config | Out-Null

        Write-Host ''
        Write-ToolkitHeader 'Registration summary'
        Write-Host ('  Tenant ID   : {0}' -f $tenantId)
        Write-Host ('  App ID      : {0}' -f $application.AppId)
        Write-Host ('  Thumbprint  : {0}' -f $certInfo.Thumbprint)
        Write-Host ('  Cert expiry : {0:yyyy-MM-dd}' -f $certInfo.NotAfter)
        Write-Host ('  Consented   : {0}' -f (($consent.Granted -join ', ') -replace '^$', 'none'))

        if ($consent.Failed.Count -gt 0) {
            Write-Host ''
            Write-ToolkitLog ('Admin consent could not be granted for: {0}' -f ($consent.Failed -join ', ')) -Level ERROR
            Write-ToolkitLog 'This is usually conditional access or a missing privileged role. Grant consent manually here:' -Level WARN
            Write-Host ('  https://login.microsoftonline.com/{0}/adminconsent?client_id={1}' -f $tenantId, $application.AppId) -ForegroundColor Cyan
        }
        else {
            Write-Host ''
            Write-ToolkitLog 'Admin consent granted for all required permissions.' -Level SUCCESS
        }

        Write-ToolkitLog 'Configuration saved. Waiting 20s for directory replication before verifying...' -Level INFO
        Start-Sleep -Seconds 20

        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose 'No interactive session to disconnect.' }
        (Get-ToolkitContext).GraphConnected = $false

        if (Connect-ToolkitGraph -TenantId $tenantId -AppId $application.AppId -CertificateThumbprint $certInfo.Thumbprint -Force) {
            Write-ToolkitLog 'Verified: app-only certificate authentication works.' -Level SUCCESS
        }
        else {
            Write-ToolkitLog 'App-only sign-in did not work yet. Consent/replication can take a few minutes - try menu option 2 shortly.' -Level WARN
        }
    }
    catch {
        Write-ToolkitLog ('App registration failed: {0}' -f $_.Exception.Message) -Level ERROR
    }
}

Export-ModuleMember -Function @(
    'Invoke-AppRegistrationSetup'
    'New-ToolkitCertificate'
    'New-ToolkitAppRegistration'
    'Add-ToolkitAppCertificate'
    'Grant-ToolkitAdminConsent'
    'New-ResourceAccessEntry'
    'Test-AppRegistrationPrerequisite'
)
