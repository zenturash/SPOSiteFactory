function Get-SPOFactoryConnection {
    <#
    .SYNOPSIS
        Gets information about cached SharePoint Online connections.

    .DESCRIPTION
        Retrieves information about active SharePoint Online connections cached in the current session.
        Can return all connections or filter by client name.

    .PARAMETER ClientName
        Optional client name to filter. If not specified, returns all connections.

    .PARAMETER Current
        Returns only the current active connection.

    .EXAMPLE
        Get-SPOFactoryConnection
        
        Lists all cached connections.

    .EXAMPLE
        Get-SPOFactoryConnection -ClientName "Contoso"
        
        Gets connection information for the Contoso client.

    .EXAMPLE
        Get-SPOFactoryConnection -Current
        
        Gets the current active connection.

    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
    #>

    [CmdletBinding(DefaultParameterSetName = 'All')]
    param(
        [Parameter(ParameterSetName = 'Specific')]
        [string]$ClientName,
        
        [Parameter(ParameterSetName = 'Current')]
        [switch]$Current
    )

    process {
        # Check if any connections exist
        if (-not $script:SPOConnections -or $script:SPOConnections.Count -eq 0) {
            Write-Host "No active SharePoint connections found" -ForegroundColor Yellow
            return
        }
        
        # Handle current connection request
        if ($Current) {
            if (-not $script:CurrentSPOConnection) {
                Write-Host "No current connection set" -ForegroundColor Yellow
                return
            }
            
            $ClientName = $script:CurrentSPOConnection
        }
        
        # Get specific or all connections
        $connectionsToReturn = @()
        
        if ($ClientName) {
            # Get specific client connection
            if ($script:SPOConnections.ContainsKey($ClientName)) {
                $conn = $script:SPOConnections[$ClientName]
                $connectionsToReturn += [PSCustomObject]@{
                    ClientName = $conn.ClientName
                    TenantUrl = $conn.TenantUrl
                    AuthMethod = $conn.AuthMethod
                    ConnectedAs = $conn.ConnectedAs
                    ConnectedAt = $conn.ConnectedAt
                    Duration = (Get-Date) - $conn.ConnectedAt
                    SessionId = $conn.SessionId
                    IsCurrent = ($script:CurrentSPOConnection -eq $conn.ClientName)
                }
            }
            else {
                Write-Host "No connection found for client: $ClientName" -ForegroundColor Yellow
                return
            }
        }
        else {
            # Get all connections
            foreach ($key in $script:SPOConnections.Keys) {
                $conn = $script:SPOConnections[$key]
                $connectionsToReturn += [PSCustomObject]@{
                    ClientName = $conn.ClientName
                    TenantUrl = $conn.TenantUrl
                    AuthMethod = $conn.AuthMethod
                    ConnectedAs = $conn.ConnectedAs
                    ConnectedAt = $conn.ConnectedAt
                    Duration = (Get-Date) - $conn.ConnectedAt
                    SessionId = $conn.SessionId
                    IsCurrent = ($script:CurrentSPOConnection -eq $conn.ClientName)
                }
            }
        }
        
        # Display connections
        if ($connectionsToReturn.Count -gt 0) {
            Write-Host "`nActive SharePoint Connections:" -ForegroundColor Cyan
            Write-Host "==============================" -ForegroundColor Cyan
            
            foreach ($conn in $connectionsToReturn) {
                $marker = if ($conn.IsCurrent) { "[*]" } else { "   " }
                Write-Host "`n$marker Client: $($conn.ClientName)" -ForegroundColor $(if ($conn.IsCurrent) { 'Green' } else { 'White' })
                Write-Host "    Tenant: $($conn.TenantUrl)" -ForegroundColor Gray
                Write-Host "    User: $($conn.ConnectedAs)" -ForegroundColor Gray
                Write-Host "    Method: $($conn.AuthMethod)" -ForegroundColor Gray
                Write-Host "    Connected: $($conn.ConnectedAt.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Gray
                Write-Host "    Duration: $($conn.Duration.ToString('hh\:mm\:ss'))" -ForegroundColor Gray
                Write-Host "    Session: $($conn.SessionId)" -ForegroundColor DarkGray
            }
            
            if ($connectionsToReturn.Count -gt 1) {
                Write-Host "`n[*] = Current connection" -ForegroundColor DarkGray
            }
        }
        
        return $connectionsToReturn
    }
}

# Alias for convenience
Set-Alias -Name Get-SPOConnection -Value Get-SPOFactoryConnection