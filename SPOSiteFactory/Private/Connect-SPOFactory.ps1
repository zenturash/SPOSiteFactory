function Connect-SPOFactory {
    <#
    .SYNOPSIS
        Establishes a connection to SharePoint Online with MSP multi-tenant support.

    .DESCRIPTION
        This function manages connections to SharePoint Online tenants in MSP environments.
        It supports multiple authentication methods, connection pooling, credential vaulting,
        and tenant isolation for managing hundreds of client tenants.

    .PARAMETER TenantUrl
        The SharePoint Online tenant URL (e.g., https://contoso.sharepoint.com or https://contoso-admin.sharepoint.com)

    .PARAMETER ClientName
        The client name for MSP tenant identification and isolation

    .PARAMETER AuthMethod
        Authentication method to use for connection

    .PARAMETER SaveCredential
        Saves the credential to the secure vault for future use

    .PARAMETER ClientId
        Application (client) ID for app-only authentication

    .PARAMETER CertificateThumbprint
        Certificate thumbprint for certificate-based authentication

    .PARAMETER TenantId
        Azure AD tenant ID (GUID)

    .PARAMETER Credential
        PSCredential object for interactive authentication

    .PARAMETER UseConnectionPool
        Reuses existing connections when possible to improve performance

    .PARAMETER Force
        Forces a new connection even if one already exists

    .EXAMPLE
        Connect-SPOFactory -TenantUrl "https://contoso.sharepoint.com" -ClientName "Contoso Corp" -AuthMethod StoredCredential

    .EXAMPLE
        Connect-SPOFactory -TenantUrl "https://contoso.sharepoint.com" -ClientName "Contoso Corp" -AuthMethod Certificate -ClientId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "ABC123..."

    .EXAMPLE
        Connect-SPOFactory -TenantUrl "https://contoso.sharepoint.com" -ClientName "Contoso Corp" -AuthMethod Interactive -SaveCredential

    .NOTES
        Requires PnP.PowerShell 2.0+ and Microsoft.PowerShell.SecretManagement modules
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({
            if ($_ -match '^https://[a-zA-Z0-9.-]+\.sharepoint\.(com|us|de|cn)/?') {
                $true
            } else {
                throw "TenantUrl must be a valid SharePoint Online URL (e.g., https://contoso.sharepoint.com)"
            }
        })]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Interactive', 'Certificate', 'AppOnly', 'StoredCredential', 'DeviceCode')]
        [string]$AuthMethod = 'StoredCredential',
        
        [Parameter(Mandatory = $false)]
        [switch]$SaveCredential,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false)]
        [string]$CertificateThumbprint,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [PSCredential]$Credential,
        
        [Parameter(Mandatory = $false)]
        [switch]$UseConnectionPool = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        Write-PSFMessage -Level Verbose -Message "Initiating SPOFactory connection for client: $ClientName"
        
        # Normalize tenant URL
        $TenantUrl = $TenantUrl.TrimEnd('/')
        $connectionKey = "$ClientName|$TenantUrl"
        
        # Validate prerequisites
        if (-not (Get-Module -Name PnP.PowerShell -ListAvailable)) {
            throw "PnP.PowerShell module is required but not installed. Please run: Install-Module -Name PnP.PowerShell"
        }
    }

    process {
        try {
            # Check for existing connection in pool
            if ($UseConnectionPool -and -not $Force -and $script:SPOFactoryConnectionPool.ContainsKey($connectionKey)) {
                $existingConnection = $script:SPOFactoryConnectionPool[$connectionKey]
                
                # Validate existing connection
                if (Test-SPOFactoryConnection -ConnectionInfo $existingConnection) {
                    Write-PSFMessage -Level Host -Message "Reusing existing connection for $ClientName"
                    
                    # Switch to existing connection
                    Connect-PnPOnline -Connection $existingConnection.Connection -ReturnConnection
                    
                    # Update last used timestamp
                    $script:SPOFactoryConnectionPool[$connectionKey].LastUsed = Get-Date
                    $script:SPOFactoryConnections[$ClientName] = $script:SPOFactoryConnectionPool[$connectionKey]
                    
                    return $script:SPOFactoryConnectionPool[$connectionKey]
                }
                else {
                    Write-PSFMessage -Level Warning -Message "Existing connection for $ClientName is no longer valid. Creating new connection."
                    $script:SPOFactoryConnectionPool.Remove($connectionKey)
                }
            }

            # Prepare connection parameters
            $connectionParams = @{
                Url = $TenantUrl
                ReturnConnection = $true
            }

            Write-PSFMessage -Level Verbose -Message "Connecting to $TenantUrl using $AuthMethod authentication"

            # Handle different authentication methods
            switch ($AuthMethod) {
                'StoredCredential' {
                    $storedCred = Get-SPOFactoryCredential -ClientName $ClientName
                    if ($storedCred) {
                        $connectionParams.Credentials = $storedCred
                        Write-PSFMessage -Level Verbose -Message "Using stored credential for $ClientName"
                    }
                    else {
                        Write-PSFMessage -Level Warning -Message "No stored credential found for $ClientName. Falling back to interactive authentication."
                        $connectionParams.Interactive = $true
                    }
                }

                'Interactive' {
                    $connectionParams.Interactive = $true
                    
                    if ($SaveCredential) {
                        Write-PSFMessage -Level Host -Message "Note: Credentials will be saved to secure vault after successful authentication"
                    }
                }

                'Certificate' {
                    if (-not $ClientId -or -not $CertificateThumbprint) {
                        throw "Certificate authentication requires both ClientId and CertificateThumbprint parameters"
                    }
                    
                    $connectionParams.ClientId = $ClientId
                    $connectionParams.Thumbprint = $CertificateThumbprint
                    
                    if ($TenantId) {
                        $connectionParams.Tenant = $TenantId
                    }
                }

                'AppOnly' {
                    if (-not $ClientId) {
                        throw "App-only authentication requires ClientId parameter"
                    }
                    
                    # Try to get client secret from vault
                    $clientSecret = Get-SPOFactorySecret -ClientName $ClientName -SecretType 'ClientSecret'
                    if ($clientSecret) {
                        $connectionParams.ClientId = $ClientId
                        $connectionParams.ClientSecret = $clientSecret
                        
                        if ($TenantId) {
                            $connectionParams.Tenant = $TenantId
                        }
                    }
                    else {
                        throw "No client secret found in vault for $ClientName. Please store the secret using Set-SPOFactorySecret"
                    }
                }

                'DeviceCode' {
                    $connectionParams.DeviceLogin = $true
                    
                    if ($ClientId) {
                        $connectionParams.ClientId = $ClientId
                    }
                }
            }

            # Add custom timeout
            if ($script:SPOFactoryConfig.ConnectionTimeout) {
                $connectionParams.ConnectionTimeout = $script:SPOFactoryConfig.ConnectionTimeout
            }

            # Attempt connection with retry logic
            $connection = Invoke-SPOFactoryCommand -ScriptBlock {
                Connect-PnPOnline @connectionParams
            } -ErrorMessage "Failed to connect to SharePoint Online for client $ClientName" -MaxRetries $script:SPOFactoryConfig.RetryAttempts

            if (-not $connection) {
                throw "Connection attempt failed for $ClientName"
            }

            # Create connection info object
            $connectionInfo = @{
                ClientName = $ClientName
                TenantUrl = $TenantUrl
                Connection = $connection
                AuthMethod = $AuthMethod
                Connected = Get-Date
                LastUsed = Get-Date
                ConnectionId = [System.Guid]::NewGuid()
                Region = Get-SPOFactoryTenantRegion -TenantUrl $TenantUrl
            }

            # Save connection to pools
            $script:SPOFactoryConnectionPool[$connectionKey] = $connectionInfo
            $script:SPOFactoryConnections[$ClientName] = $connectionInfo

            # Save credential if requested and using interactive auth
            if ($SaveCredential -and $AuthMethod -eq 'Interactive' -and $Credential) {
                try {
                    Set-SPOFactoryCredential -ClientName $ClientName -Credential $Credential
                    Write-PSFMessage -Level Host -Message "Credential saved to secure vault for $ClientName"
                }
                catch {
                    Write-PSFMessage -Level Warning -Message "Failed to save credential for $ClientName`: $_"
                }
            }

            # Log successful connection
            Write-PSFMessage -Level Host -Message "Successfully connected to SharePoint Online for client: $ClientName"
            Write-SPOFactoryLog -Message "Connection established for client $ClientName to $TenantUrl using $AuthMethod" -Level Info -ClientName $ClientName

            # Update tenant registry
            Update-SPOFactoryTenantRegistry -ClientName $ClientName -TenantUrl $TenantUrl -LastConnected (Get-Date)

            return $connectionInfo
        }
        catch {
            $errorMsg = "Failed to connect to SharePoint Online for client $ClientName`: $_"
            Write-PSFMessage -Level Error -Message $errorMsg
            Write-SPOFactoryLog -Message $errorMsg -Level Error -ClientName $ClientName
            throw
        }
    }

    end {
        # Clean up old connections from pool if needed
        if ($script:SPOFactoryConnectionPool.Count -gt $script:SPOFactoryConfig.MaxConcurrentConnections) {
            Remove-SPOFactoryStaleConnections
        }
    }
}

function Disconnect-SPOFactory {
    <#
    .SYNOPSIS
        Disconnects from SharePoint Online for a specific client or all clients.

    .DESCRIPTION
        Safely disconnects from SharePoint Online tenants with proper cleanup
        for MSP multi-tenant environments.

    .PARAMETER ClientName
        The client name to disconnect. If not specified, disconnects all clients.

    .PARAMETER All
        Disconnects from all active client connections.

    .EXAMPLE
        Disconnect-SPOFactory -ClientName "Contoso Corp"

    .EXAMPLE
        Disconnect-SPOFactory -All
    #>

    [CmdletBinding(DefaultParameterSetName = 'ByClient')]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'ByClient')]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'All')]
        [switch]$All
    )

    process {
        try {
            if ($All) {
                Write-PSFMessage -Level Host -Message "Disconnecting from all SharePoint Online tenants..."
                
                $disconnectedCount = 0
                foreach ($client in $script:SPOFactoryConnections.Keys) {
                    try {
                        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
                            Disconnect-PnPOnline -Connection $script:SPOFactoryConnections[$client].Connection
                        }
                        
                        Write-SPOFactoryLog -Message "Disconnected from tenant" -Level Info -ClientName $client
                        $disconnectedCount++
                    }
                    catch {
                        Write-PSFMessage -Level Warning -Message "Failed to disconnect from $client`: $_"
                    }
                }
                
                # Clear all connections
                $script:SPOFactoryConnections.Clear()
                $script:SPOFactoryConnectionPool.Clear()
                
                Write-PSFMessage -Level Host -Message "Disconnected from $disconnectedCount SharePoint Online tenants"
            }
            elseif ($ClientName) {
                if ($script:SPOFactoryConnections.ContainsKey($ClientName)) {
                    try {
                        if (Get-PnPConnection -ErrorAction SilentlyContinue) {
                            Disconnect-PnPOnline -Connection $script:SPOFactoryConnections[$ClientName].Connection
                        }
                        
                        $connectionKey = "$ClientName|$($script:SPOFactoryConnections[$ClientName].TenantUrl)"
                        $script:SPOFactoryConnections.Remove($ClientName)
                        $script:SPOFactoryConnectionPool.Remove($connectionKey)
                        
                        Write-SPOFactoryLog -Message "Disconnected from tenant" -Level Info -ClientName $ClientName
                        Write-PSFMessage -Level Host -Message "Disconnected from SharePoint Online for client: $ClientName"
                    }
                    catch {
                        Write-PSFMessage -Level Warning -Message "Failed to disconnect from $ClientName`: $_"
                        throw
                    }
                }
                else {
                    Write-PSFMessage -Level Warning -Message "No active connection found for client: $ClientName"
                }
            }
            else {
                # Disconnect current connection if any
                if (Get-PnPConnection -ErrorAction SilentlyContinue) {
                    Disconnect-PnPOnline
                    Write-PSFMessage -Level Host -Message "Disconnected from current SharePoint Online session"
                }
            }
        }
        catch {
            Write-PSFMessage -Level Error -Message "Error during disconnect operation: $_"
            throw
        }
    }
}

function Test-SPOFactoryConnection {
    <#
    .SYNOPSIS
        Tests the validity of a SharePoint Online connection.

    .DESCRIPTION
        Validates that a connection to SharePoint Online is still active and responsive.

    .PARAMETER ConnectionInfo
        Connection information object to test.

    .PARAMETER ClientName
        Client name to test connection for.

    .EXAMPLE
        Test-SPOFactoryConnection -ClientName "Contoso Corp"

    .EXAMPLE
        Test-SPOFactoryConnection -ConnectionInfo $connectionInfo
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false, ParameterSetName = 'ByInfo')]
        [hashtable]$ConnectionInfo,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'ByClient')]
        [string]$ClientName
    )

    process {
        try {
            if ($ClientName) {
                if (-not $script:SPOFactoryConnections.ContainsKey($ClientName)) {
                    return $false
                }
                $ConnectionInfo = $script:SPOFactoryConnections[$ClientName]
            }

            if (-not $ConnectionInfo) {
                return $false
            }

            # Test connection by trying to get tenant properties
            $connection = Get-PnPConnection -Connection $ConnectionInfo.Connection -ErrorAction SilentlyContinue
            if (-not $connection) {
                return $false
            }

            # Quick validation - try to get web properties
            $null = Get-PnPWeb -Connection $ConnectionInfo.Connection -ErrorAction Stop
            
            return $true
        }
        catch {
            Write-PSFMessage -Level Debug -Message "Connection test failed: $_"
            return $false
        }
    }
}

function Remove-SPOFactoryStaleConnections {
    <#
    .SYNOPSIS
        Removes stale connections from the connection pool.

    .DESCRIPTION
        Cleans up old or inactive connections to maintain optimal performance
        and stay within connection limits.

    .PARAMETER MaxAge
        Maximum age in minutes for connections before they're considered stale.

    .EXAMPLE
        Remove-SPOFactoryStaleConnections -MaxAge 30
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [int]$MaxAge = 60
    )

    process {
        $cutoffTime = (Get-Date).AddMinutes(-$MaxAge)
        $removedCount = 0
        
        $staleConnections = $script:SPOFactoryConnectionPool.Keys | Where-Object {
            $script:SPOFactoryConnectionPool[$_].LastUsed -lt $cutoffTime
        }
        
        foreach ($connectionKey in $staleConnections) {
            try {
                $connectionInfo = $script:SPOFactoryConnectionPool[$connectionKey]
                
                if (Get-PnPConnection -Connection $connectionInfo.Connection -ErrorAction SilentlyContinue) {
                    Disconnect-PnPOnline -Connection $connectionInfo.Connection
                }
                
                $script:SPOFactoryConnectionPool.Remove($connectionKey)
                $script:SPOFactoryConnections.Remove($connectionInfo.ClientName)
                $removedCount++
                
                Write-PSFMessage -Level Debug -Message "Removed stale connection for $($connectionInfo.ClientName)"
            }
            catch {
                Write-PSFMessage -Level Warning -Message "Failed to remove stale connection: $_"
            }
        }
        
        if ($removedCount -gt 0) {
            Write-PSFMessage -Level Verbose -Message "Removed $removedCount stale connections from pool"
        }
    }
}