function Connect-SPOFactory {
    <#
    .SYNOPSIS
        Establishes a connection to SharePoint Online with multiple authentication methods and token caching.

    .DESCRIPTION
        Enterprise-grade connection management for MSP environments supporting multiple SharePoint Online tenants.
        Supports Interactive, DeviceCode, Certificate, ManagedIdentity, and StoredCredential authentication.
        Caches authentication tokens in the current PowerShell session for reuse.

    .PARAMETER TenantUrl
        SharePoint admin center URL (https://tenant-admin.sharepoint.com)

    .PARAMETER ClientName
        MSP client identifier for connection isolation and logging

    .PARAMETER Interactive
        Use interactive browser-based authentication (default if no other method specified)

    .PARAMETER DeviceCode
        Use device code flow for authentication (useful for remote/restricted environments)

    .PARAMETER Certificate
        Use certificate-based authentication (requires CertificateThumbprint and ClientId)

    .PARAMETER CertificateThumbprint
        Certificate thumbprint for certificate-based authentication

    .PARAMETER ClientId
        Azure AD application ID for app-only authentication

    .PARAMETER ManagedIdentity
        Use Azure Managed Identity for authentication (for Azure-hosted scenarios)

    .PARAMETER StoredCredential
        Name of stored credential in SecretStore

    .PARAMETER Credential
        PSCredential object for direct authentication

    .PARAMETER Force
        Force new connection even if one exists

    .PARAMETER SkipCertificateCheck
        Skip certificate validation (not recommended for production)

    .EXAMPLE
        Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" -ClientName "Contoso" -Interactive
        
        Connects using interactive browser authentication.

    .EXAMPLE
        Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" -ClientName "Contoso" -DeviceCode
        
        Connects using device code flow.

    .EXAMPLE
        Connect-SPOFactory -TenantUrl "https://contoso-admin.sharepoint.com" -ClientName "Contoso" -Certificate -CertificateThumbprint "1234567890" -ClientId "app-guid"
        
        Connects using certificate-based app-only authentication.

    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
        Requires: PnP.PowerShell 2.0+
    #>

    [CmdletBinding(DefaultParameterSetName = 'Interactive')]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^https://[a-zA-Z0-9-]+-admin\.sharepoint\.com$')]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientName,
        
        [Parameter(ParameterSetName = 'Interactive')]
        [switch]$Interactive,
        
        [Parameter(ParameterSetName = 'DeviceCode')]
        [switch]$DeviceCode,
        
        [Parameter(ParameterSetName = 'Certificate')]
        [switch]$Certificate,
        
        [Parameter(ParameterSetName = 'Certificate', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$CertificateThumbprint,
        
        [Parameter(ParameterSetName = 'Certificate', Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,
        
        [Parameter(ParameterSetName = 'ManagedIdentity')]
        [switch]$ManagedIdentity,
        
        [Parameter(ParameterSetName = 'StoredCredential')]
        [ValidateNotNullOrEmpty()]
        [string]$StoredCredential,
        
        [Parameter(ParameterSetName = 'Credential')]
        [System.Management.Automation.PSCredential]$Credential,
        
        [Parameter()]
        [switch]$Force,
        
        [Parameter()]
        [switch]$SkipCertificateCheck
    )

    begin {
        $startTime = Get-Date
        Write-SPOFactoryLog -Message "Initiating connection to $TenantUrl for client: $ClientName" -Level Info -ClientName $ClientName
        
        # Initialize connection cache if not exists
        if (-not $script:SPOConnections) {
            $script:SPOConnections = @{}
        }
        
        # Check for existing connection
        if ($script:SPOConnections.ContainsKey($ClientName) -and -not $Force) {
            $existingConnection = $script:SPOConnections[$ClientName]
            
            # Test if connection is still valid
            try {
                Connect-PnPOnline -Url $existingConnection.TenantUrl -ReturnConnection -ErrorAction Stop | Out-Null
                Write-SPOFactoryLog -Message "Using existing connection for client: $ClientName" -Level Info -ClientName $ClientName
                
                return @{
                    Success = $true
                    ClientName = $ClientName
                    TenantUrl = $existingConnection.TenantUrl
                    AuthMethod = $existingConnection.AuthMethod
                    ConnectedAt = $existingConnection.ConnectedAt
                    CachedConnection = $true
                    Message = "Using cached connection"
                }
            }
            catch {
                Write-SPOFactoryLog -Message "Existing connection invalid, establishing new connection" -Level Warning -ClientName $ClientName
                $script:SPOConnections.Remove($ClientName)
            }
        }
    }

    process {
        try {
            $connectionParams = @{
                ReturnConnection = $true
                ErrorAction = 'Stop'
            }
            
            # Add certificate check parameter if specified
            if ($SkipCertificateCheck) {
                $connectionParams['SkipTenantAdminCheck'] = $true
            }
            
            # Determine authentication method and connect
            $authMethod = switch ($PSCmdlet.ParameterSetName) {
                'Interactive' {
                    Write-SPOFactoryLog -Message "Using interactive authentication" -Level Info -ClientName $ClientName
                    $connectionParams['Url'] = $TenantUrl
                    $connectionParams['Interactive'] = $true
                    'Interactive'
                }
                
                'DeviceCode' {
                    Write-SPOFactoryLog -Message "Using device code authentication" -Level Info -ClientName $ClientName
                    Write-Host "To sign in, use a web browser to open the page https://microsoft.com/devicelogin" -ForegroundColor Yellow
                    $connectionParams['Url'] = $TenantUrl
                    $connectionParams['DeviceLogin'] = $true
                    'DeviceCode'
                }
                
                'Certificate' {
                    Write-SPOFactoryLog -Message "Using certificate authentication" -Level Info -ClientName $ClientName
                    $connectionParams['Url'] = $TenantUrl
                    $connectionParams['ClientId'] = $ClientId
                    $connectionParams['Thumbprint'] = $CertificateThumbprint
                    
                    # Extract tenant from URL
                    if ($TenantUrl -match 'https://([^-]+)-admin\.sharepoint\.com') {
                        $connectionParams['Tenant'] = $matches[1] + '.onmicrosoft.com'
                    }
                    'Certificate'
                }
                
                'ManagedIdentity' {
                    Write-SPOFactoryLog -Message "Using managed identity authentication" -Level Info -ClientName $ClientName
                    $connectionParams['Url'] = $TenantUrl
                    $connectionParams['ManagedIdentity'] = $true
                    'ManagedIdentity'
                }
                
                'StoredCredential' {
                    Write-SPOFactoryLog -Message "Using stored credential: $StoredCredential" -Level Info -ClientName $ClientName
                    
                    # Retrieve credential from SecretStore
                    try {
                        $cred = Get-Secret -Name $StoredCredential -AsPlainText -ErrorAction Stop
                        $connectionParams['Url'] = $TenantUrl
                        $connectionParams['Credentials'] = $cred
                    }
                    catch {
                        # Fallback to credential manager
                        $cred = Get-StoredCredential -Target $StoredCredential
                        if (-not $cred) {
                            throw "Stored credential '$StoredCredential' not found"
                        }
                        $connectionParams['Url'] = $TenantUrl
                        $connectionParams['Credentials'] = $cred
                    }
                    'StoredCredential'
                }
                
                'Credential' {
                    Write-SPOFactoryLog -Message "Using provided credential" -Level Info -ClientName $ClientName
                    $connectionParams['Url'] = $TenantUrl
                    $connectionParams['Credentials'] = $Credential
                    'Credential'
                }
                
                default {
                    # Default to Interactive if no parameter set specified
                    Write-SPOFactoryLog -Message "No authentication method specified, defaulting to interactive" -Level Info -ClientName $ClientName
                    $connectionParams['Url'] = $TenantUrl
                    $connectionParams['Interactive'] = $true
                    'Interactive'
                }
            }
            
            # Establish connection
            Write-SPOFactoryLog -Message "Establishing connection to SharePoint Online..." -Level Info -ClientName $ClientName
            $connection = Connect-PnPOnline @connectionParams
            
            # Verify connection
            $context = Get-PnPContext
            if (-not $context) {
                throw "Failed to establish PnP context"
            }
            
            # Get additional connection info
            $web = Get-PnPWeb -Includes Title, Url
            $currentUser = Get-PnPProperty -ClientObject $context.Web -Property CurrentUser
            
            # Cache connection information
            $connectionInfo = @{
                ClientName = $ClientName
                TenantUrl = $TenantUrl
                AuthMethod = $authMethod
                Connection = $connection
                Context = $context
                ConnectedAt = Get-Date
                ConnectedAs = $currentUser.Email
                WebTitle = $web.Title
                WebUrl = $web.Url
                SessionId = [guid]::NewGuid().ToString()
            }
            
            $script:SPOConnections[$ClientName] = $connectionInfo
            
            # Set as current connection
            $script:CurrentSPOConnection = $ClientName
            
            # Log success
            $duration = (Get-Date) - $startTime
            Write-SPOFactoryLog -Message "Successfully connected to $TenantUrl as $($currentUser.Email)" -Level Info -ClientName $ClientName
            Write-SPOFactoryLog -Message "Connection established in $($duration.TotalSeconds) seconds" -Level Debug -ClientName $ClientName
            
            # Display connection info
            Write-Host "`nSuccessfully connected to SharePoint Online" -ForegroundColor Green
            Write-Host "  Tenant: $TenantUrl" -ForegroundColor Gray
            Write-Host "  Client: $ClientName" -ForegroundColor Gray
            Write-Host "  User: $($currentUser.Email)" -ForegroundColor Gray
            Write-Host "  Method: $authMethod" -ForegroundColor Gray
            Write-Host "  Session: $($connectionInfo.SessionId)" -ForegroundColor Gray
            
            # Return connection details
            return @{
                Success = $true
                ClientName = $ClientName
                TenantUrl = $TenantUrl
                AuthMethod = $authMethod
                ConnectedAs = $currentUser.Email
                ConnectedAt = $connectionInfo.ConnectedAt
                SessionId = $connectionInfo.SessionId
                WebTitle = $web.Title
                WebUrl = $web.Url
                Message = "Connection established successfully"
            }
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-SPOFactoryLog -Message "Connection failed: $errorMessage" -Level Error -ClientName $ClientName -Exception $_
            
            # Provide helpful error messages
            $helpMessage = switch -Regex ($errorMessage) {
                'AADSTS50076' { "Multi-factor authentication is required. Use -Interactive or -DeviceCode parameter." }
                'AADSTS700016' { "Application not found. Verify ClientId and app registration." }
                'AADSTS70001' { "Application disabled. Contact your Azure AD administrator." }
                'certificate' { "Certificate not found. Verify thumbprint and certificate store." }
                'tenant' { "Invalid tenant URL. Use format: https://tenant-admin.sharepoint.com" }
                'credentials' { "Invalid credentials. Verify username and password." }
                default { "Check your authentication parameters and network connectivity." }
            }
            
            Write-Host "`nConnection failed: $errorMessage" -ForegroundColor Red
            Write-Host "Suggestion: $helpMessage" -ForegroundColor Yellow
            
            throw "Failed to connect to SharePoint Online: $errorMessage"
        }
    }

    end {
        # Export connection for use in other functions
        if ($script:SPOConnections.ContainsKey($ClientName)) {
            $env:SPOFactoryCurrentClient = $ClientName
            Write-SPOFactoryLog -Message "Connection cached and set as current for client: $ClientName" -Level Debug -ClientName $ClientName
        }
    }
}

# Helper function to get stored credentials (Windows Credential Manager)
function Get-StoredCredential {
    param([string]$Target)
    
    try {
        Add-Type -AssemblyName System.Runtime.InteropServices
        $sig = @'
[DllImport("Advapi32.dll", SetLastError = true, EntryPoint = "CredReadW", CharSet = CharSet.Unicode)]
public static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);

[DllImport("Advapi32.dll", SetLastError = true)]
public static extern bool CredFree([In] IntPtr cred);
'@
        
        $type = Add-Type -MemberDefinition $sig -Name Win32Utils -Namespace CredManager -PassThru -ErrorAction Stop
        
        $credPtr = [IntPtr]::Zero
        $success = $type::CredRead($Target, 1, 0, [ref]$credPtr)
        
        if ($success) {
            $credObject = [System.Runtime.InteropServices.Marshal]::PtrToStructure($credPtr, [Type]"CREDENTIAL")
            $username = $credObject.UserName
            $password = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($credObject.CredentialBlob)
            
            $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)
            
            [void]$type::CredFree($credPtr)
            return $credential
        }
    }
    catch {
        Write-SPOFactoryLog -Message "Failed to retrieve stored credential: $_" -Level Warning
    }
    
    return $null
}