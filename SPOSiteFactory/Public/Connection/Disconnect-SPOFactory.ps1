function Disconnect-SPOFactory {
    <#
    .SYNOPSIS
        Disconnects from SharePoint Online and clears cached connections.

    .DESCRIPTION
        Closes the SharePoint Online connection for the specified client and removes cached authentication tokens.
        Can disconnect specific clients or all cached connections.

    .PARAMETER ClientName
        The MSP client name to disconnect. Use 'All' to disconnect all clients.

    .PARAMETER Force
        Suppress confirmation prompts.

    .EXAMPLE
        Disconnect-SPOFactory -ClientName "Contoso"
        
        Disconnects the Contoso client connection.

    .EXAMPLE
        Disconnect-SPOFactory -ClientName "All" -Force
        
        Disconnects all cached connections without confirmation.

    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter()]
        [switch]$Force
    )

    process {
        # Check if connection cache exists
        if (-not $script:SPOConnections) {
            Write-SPOFactoryLog -Message "No active connections found" -Level Warning
            Write-Host "No active connections to disconnect" -ForegroundColor Yellow
            return
        }
        
        # Handle disconnect all
        if ($ClientName -eq 'All') {
            $clientsToDisconnect = $script:SPOConnections.Keys | ForEach-Object { $_ }
            
            if ($clientsToDisconnect.Count -eq 0) {
                Write-Host "No active connections found" -ForegroundColor Yellow
                return
            }
            
            if ($PSCmdlet.ShouldProcess("$($clientsToDisconnect.Count) client connections", "Disconnect")) {
                if (-not $Force) {
                    $confirmation = Read-Host "Are you sure you want to disconnect all $($clientsToDisconnect.Count) connections? (Y/N)"
                    if ($confirmation -ne 'Y') {
                        Write-Host "Disconnect cancelled" -ForegroundColor Yellow
                        return
                    }
                }
                
                foreach ($client in $clientsToDisconnect) {
                    try {
                        # Disconnect PnP session
                        Disconnect-PnPOnline -ErrorAction SilentlyContinue
                        
                        # Remove from cache
                        $script:SPOConnections.Remove($client)
                        
                        Write-SPOFactoryLog -Message "Disconnected client: $client" -Level Info -ClientName $client
                        Write-Host "Disconnected: $client" -ForegroundColor Gray
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Error disconnecting client ${client}: $_" -Level Warning -ClientName $client
                    }
                }
                
                # Clear current connection
                if ($script:CurrentSPOConnection) {
                    $script:CurrentSPOConnection = $null
                }
                
                $env:SPOFactoryCurrentClient = $null
                
                Write-Host "`nAll connections disconnected successfully" -ForegroundColor Green
            }
        }
        else {
            # Disconnect specific client
            if (-not $script:SPOConnections.ContainsKey($ClientName)) {
                Write-SPOFactoryLog -Message "No active connection found for client: $ClientName" -Level Warning
                Write-Host "No active connection found for client: $ClientName" -ForegroundColor Yellow
                return
            }
            
            if ($PSCmdlet.ShouldProcess($ClientName, "Disconnect SharePoint connection")) {
                try {
                    $connectionInfo = $script:SPOConnections[$ClientName]
                    
                    # Disconnect PnP session
                    Disconnect-PnPOnline -ErrorAction SilentlyContinue
                    
                    # Remove from cache
                    $script:SPOConnections.Remove($ClientName)
                    
                    # Clear current connection if it matches
                    if ($script:CurrentSPOConnection -eq $ClientName) {
                        $script:CurrentSPOConnection = $null
                        $env:SPOFactoryCurrentClient = $null
                    }
                    
                    Write-SPOFactoryLog -Message "Disconnected from $($connectionInfo.TenantUrl)" -Level Info -ClientName $ClientName
                    
                    Write-Host "`nDisconnected from SharePoint Online" -ForegroundColor Green
                    Write-Host "  Client: $ClientName" -ForegroundColor Gray
                    Write-Host "  Tenant: $($connectionInfo.TenantUrl)" -ForegroundColor Gray
                    Write-Host "  Session Duration: $((Get-Date) - $connectionInfo.ConnectedAt)" -ForegroundColor Gray
                }
                catch {
                    Write-SPOFactoryLog -Message "Error disconnecting: $_" -Level Error -ClientName $ClientName
                    Write-Host "Error disconnecting: $_" -ForegroundColor Red
                }
            }
        }
    }

    end {
        # Check remaining connections
        if ($script:SPOConnections -and $script:SPOConnections.Count -gt 0) {
            Write-Host "`nRemaining active connections: $($script:SPOConnections.Count)" -ForegroundColor Cyan
            $script:SPOConnections.Keys | ForEach-Object {
                Write-Host "  - $_" -ForegroundColor Gray
            }
        }
        else {
            Write-Host "`nAll connections closed" -ForegroundColor Gray
        }
    }
}