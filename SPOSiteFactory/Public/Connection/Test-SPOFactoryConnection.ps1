function Test-SPOFactoryConnection {
    <#
    .SYNOPSIS
        Tests the current SharePoint Online connection status.

    .DESCRIPTION
        Verifies if there is an active connection to SharePoint Online for the specified client.
        Returns detailed connection information including authentication method and permissions.

    .PARAMETER ClientName
        The MSP client name to test. If not specified, tests the current connection.

    .PARAMETER Detailed
        Return detailed connection information including permissions and context.

    .EXAMPLE
        Test-SPOFactoryConnection -ClientName "Contoso"
        
        Tests if there is an active connection for the Contoso client.

    .EXAMPLE
        Test-SPOFactoryConnection -Detailed
        
        Returns detailed information about the current connection.

    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
    #>

    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ClientName,
        
        [Parameter()]
        [switch]$Detailed
    )

    process {
        # If no client specified, use current connection
        if (-not $ClientName) {
            if ($script:CurrentSPOConnection) {
                $ClientName = $script:CurrentSPOConnection
            }
            elseif ($env:SPOFactoryCurrentClient) {
                $ClientName = $env:SPOFactoryCurrentClient
            }
            else {
                Write-SPOFactoryLog -Message "No active connection found" -Level Warning
                return @{
                    IsConnected = $false
                    Message = "No active SharePoint connection"
                }
            }
        }
        
        # Check if connection exists in cache
        if (-not $script:SPOConnections -or -not $script:SPOConnections.ContainsKey($ClientName)) {
            Write-SPOFactoryLog -Message "No cached connection found for client: $ClientName" -Level Warning
            return @{
                IsConnected = $false
                ClientName = $ClientName
                Message = "No connection found for client: $ClientName"
            }
        }
        
        $connectionInfo = $script:SPOConnections[$ClientName]
        
        try {
            # Test the actual connection
            Write-SPOFactoryLog -Message "Testing connection for client: $ClientName" -Level Debug -ClientName $ClientName
            
            # Try to get current web to verify connection
            $web = Get-PnPWeb -ErrorAction Stop
            
            if (-not $web) {
                throw "Unable to retrieve web context"
            }
            
            # Build result object
            $result = @{
                IsConnected = $true
                ClientName = $ClientName
                TenantUrl = $connectionInfo.TenantUrl
                AuthMethod = $connectionInfo.AuthMethod
                ConnectedAs = $connectionInfo.ConnectedAs
                ConnectedAt = $connectionInfo.ConnectedAt
                SessionId = $connectionInfo.SessionId
                SessionDuration = (Get-Date) - $connectionInfo.ConnectedAt
                WebUrl = $web.Url
                WebTitle = $web.Title
                Message = "Connection is active"
            }
            
            # Add detailed information if requested
            if ($Detailed) {
                try {
                    # Get current user details
                    $currentUser = Get-PnPProperty -ClientObject (Get-PnPContext).Web -Property CurrentUser
                    
                    # Get site permissions
                    $permissions = @()
                    try {
                        $sitePermissions = Get-PnPSiteCollectionAdmin -ErrorAction SilentlyContinue
                        if ($sitePermissions) {
                            $permissions += "SiteCollectionAdmin"
                        }
                    }
                    catch {
                        # Not a site collection admin
                    }
                    
                    # Get tenant info
                    $tenantInfo = @{}
                    try {
                        $tenant = Get-PnPTenant -ErrorAction SilentlyContinue
                        if ($tenant) {
                            $tenantInfo = @{
                                DisplayName = $tenant.DisplayName
                                TenantId = $tenant.TenantId
                            }
                            $permissions += "TenantAdmin"
                        }
                    }
                    catch {
                        # Not a tenant admin
                    }
                    
                    # Add detailed properties
                    $result['UserDetails'] = @{
                        Email = $currentUser.Email
                        Title = $currentUser.Title
                        Id = $currentUser.Id
                        LoginName = $currentUser.LoginName
                    }
                    
                    $result['Permissions'] = $permissions
                    $result['TenantInfo'] = $tenantInfo
                    
                    # Get connection limits
                    $result['Limits'] = @{
                        MaxSites = 2000000  # SharePoint Online limit
                        MaxListItems = 30000000  # List item limit
                        MaxFileSize = 250  # GB
                    }
                    
                    # Get PnP version
                    $pnpModule = Get-Module -Name PnP.PowerShell
                    if ($pnpModule) {
                        $result['PnPVersion'] = $pnpModule.Version.ToString()
                    }
                }
                catch {
                    Write-SPOFactoryLog -Message "Could not retrieve detailed information: $_" -Level Warning -ClientName $ClientName
                }
            }
            
            Write-SPOFactoryLog -Message "Connection test successful for client: $ClientName" -Level Info -ClientName $ClientName
            
            # Display connection status
            Write-Host "`nConnection Status: " -NoNewline
            Write-Host "Connected" -ForegroundColor Green
            Write-Host "  Client: $ClientName" -ForegroundColor Gray
            Write-Host "  Tenant: $($connectionInfo.TenantUrl)" -ForegroundColor Gray
            Write-Host "  User: $($connectionInfo.ConnectedAs)" -ForegroundColor Gray
            Write-Host "  Duration: $($result.SessionDuration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
            
            if ($Detailed -and $result.Permissions) {
                Write-Host "  Permissions: $($result.Permissions -join ', ')" -ForegroundColor Gray
            }
            
            return $result
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-SPOFactoryLog -Message "Connection test failed: $errorMessage" -Level Warning -ClientName $ClientName
            
            # Remove invalid connection from cache
            $script:SPOConnections.Remove($ClientName)
            if ($script:CurrentSPOConnection -eq $ClientName) {
                $script:CurrentSPOConnection = $null
                $env:SPOFactoryCurrentClient = $null
            }
            
            Write-Host "`nConnection Status: " -NoNewline
            Write-Host "Disconnected" -ForegroundColor Red
            Write-Host "  Client: $ClientName" -ForegroundColor Gray
            Write-Host "  Error: $errorMessage" -ForegroundColor Gray
            Write-Host "  Action: Connection removed from cache" -ForegroundColor Yellow
            
            return @{
                IsConnected = $false
                ClientName = $ClientName
                TenantUrl = $connectionInfo.TenantUrl
                Error = $errorMessage
                Message = "Connection is no longer valid"
            }
        }
    }
}