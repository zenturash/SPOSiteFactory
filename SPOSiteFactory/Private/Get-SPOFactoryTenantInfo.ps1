function Get-SPOFactoryTenantInfo {
    <#
    .SYNOPSIS
        Retrieves comprehensive tenant information for MSP client management.

    .DESCRIPTION
        Gathers detailed information about SharePoint Online tenants including
        configuration, health status, licensing, and MSP-specific metadata.

    .PARAMETER ClientName
        The client name to get tenant information for

    .PARAMETER TenantUrl
        SharePoint Online tenant URL

    .PARAMETER IncludeHealth
        Include tenant health and service status information

    .PARAMETER IncludeLicensing
        Include licensing and quota information

    .PARAMETER IncludeConfiguration
        Include tenant configuration settings

    .PARAMETER UseCache
        Use cached tenant information if available

    .EXAMPLE
        Get-SPOFactoryTenantInfo -ClientName "Contoso Corp"

    .EXAMPLE
        Get-SPOFactoryTenantInfo -ClientName "Contoso Corp" -IncludeHealth -IncludeLicensing
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeHealth,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeLicensing,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeConfiguration,
        
        [Parameter(Mandatory = $false)]
        [switch]$UseCache = $true
    )

    begin {
        Write-SPOFactoryLog -Message "Retrieving tenant information for $ClientName" -Level Info -ClientName $ClientName -Category 'System'
        
        # Check cache first
        if ($UseCache) {
            $cacheKey = "TenantInfo-$ClientName"
            $cachedInfo = Get-SPOFactoryCacheItem -Key $cacheKey -MaxAge (New-TimeSpan -Hours 1)
            if ($cachedInfo) {
                Write-SPOFactoryLog -Message "Using cached tenant information for $ClientName" -Level Debug -ClientName $ClientName -Category 'System'
                return $cachedInfo
            }
        }
    }

    process {
        try {
            # Initialize tenant info object
            $tenantInfo = @{
                ClientName = $ClientName
                CollectedAt = Get-Date
                BasicInfo = @{}
                Health = @{}
                Licensing = @{}
                Configuration = @{}
                MSPMetadata = @{}
            }

            # Ensure we have a connection
            $connection = $script:SPOFactoryConnections[$ClientName]
            if (-not $connection) {
                throw "No active connection found for client: $ClientName"
            }

            # Get basic tenant information
            $tenantInfo.BasicInfo = Get-SPOFactoryBasicTenantInfo -ClientName $ClientName

            # Get health information if requested
            if ($IncludeHealth) {
                $tenantInfo.Health = Get-SPOFactoryTenantHealth -ClientName $ClientName
            }

            # Get licensing information if requested
            if ($IncludeLicensing) {
                $tenantInfo.Licensing = Get-SPOFactoryTenantLicensing -ClientName $ClientName
            }

            # Get configuration information if requested
            if ($IncludeConfiguration) {
                $tenantInfo.Configuration = Get-SPOFactoryTenantConfiguration -ClientName $ClientName
            }

            # Add MSP metadata
            $tenantInfo.MSPMetadata = Get-SPOFactoryMSPMetadata -ClientName $ClientName

            # Cache the results
            if ($UseCache) {
                $cacheKey = "TenantInfo-$ClientName"
                Set-SPOFactoryCacheItem -Key $cacheKey -Value $tenantInfo -ExpiresIn (New-TimeSpan -Hours 1)
            }

            Write-SPOFactoryLog -Message "Successfully retrieved tenant information for $ClientName" -Level Info -ClientName $ClientName -Category 'System'
            return $tenantInfo
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve tenant information for $ClientName`: $_" -Level Error -ClientName $ClientName -Category 'System' -Exception $_.Exception
            throw
        }
    }
}

function Get-SPOFactoryBasicTenantInfo {
    [CmdletBinding()]
    param([string]$ClientName)

    return Invoke-SPOFactoryCommand -ScriptBlock {
        $tenant = Get-PnPTenant -ErrorAction Stop
        $web = Get-PnPWeb -ErrorAction Stop
        
        @{
            DisplayName = $tenant.DisplayName
            TenantUrl = $web.Url
            PrimaryDomain = $tenant.DefaultSiteCollectionOwner
            Region = Get-SPOFactoryTenantRegion -TenantUrl $web.Url
            CreatedDate = $tenant.CreatedTime
            LastModified = $tenant.LastModified
            TenantId = $tenant.TenantId
            ComplianceAttribute = $tenant.ComplianceAttribute
            GeoLocation = $tenant.GeoLocation
        }
    } -ClientName $ClientName -Category 'System' -ErrorMessage "Failed to get basic tenant information" -PassThru
}

function Get-SPOFactoryTenantHealth {
    [CmdletBinding()]
    param([string]$ClientName)

    $healthInfo = @{
        OverallStatus = 'Unknown'
        ServiceStatus = @{}
        LastChecked = Get-Date
        Issues = @()
        Warnings = @()
    }

    try {
        # Check basic connectivity
        $connectivityTest = Invoke-SPOFactoryCommand -ScriptBlock {
            try {
                $web = Get-PnPWeb -ErrorAction Stop
                return @{ Status = 'Healthy'; ResponseTime = (Measure-Command { Get-PnPWeb }).TotalMilliseconds }
            }
            catch {
                return @{ Status = 'Unhealthy'; Error = $_.Exception.Message }
            }
        } -ClientName $ClientName -Category 'System' -SuppressErrors -PassThru

        $healthInfo.ServiceStatus.Connectivity = $connectivityTest

        # Check storage quota
        $storageInfo = Invoke-SPOFactoryCommand -ScriptBlock {
            try {
                $tenant = Get-PnPTenant
                $storageQuota = $tenant.StorageQuota
                $storageUsed = $tenant.StorageQuotaUsed
                $storagePercentUsed = if ($storageQuota -gt 0) { ($storageUsed / $storageQuota) * 100 } else { 0 }
                
                return @{
                    Status = if ($storagePercentUsed -gt 90) { 'Critical' } elseif ($storagePercentUsed -gt 75) { 'Warning' } else { 'Healthy' }
                    QuotaGB = [math]::Round($storageQuota / 1024, 2)
                    UsedGB = [math]::Round($storageUsed / 1024, 2)
                    PercentUsed = [math]::Round($storagePercentUsed, 2)
                }
            }
            catch {
                return @{ Status = 'Unknown'; Error = $_.Exception.Message }
            }
        } -ClientName $ClientName -Category 'System' -SuppressErrors -PassThru

        $healthInfo.ServiceStatus.Storage = $storageInfo

        # Determine overall status
        $statuses = $healthInfo.ServiceStatus.Values | ForEach-Object { $_.Status }
        if ($statuses -contains 'Critical') {
            $healthInfo.OverallStatus = 'Critical'
        } elseif ($statuses -contains 'Warning') {
            $healthInfo.OverallStatus = 'Warning'
        } elseif ($statuses -contains 'Unhealthy') {
            $healthInfo.OverallStatus = 'Unhealthy'
        } elseif ($statuses -contains 'Healthy') {
            $healthInfo.OverallStatus = 'Healthy'
        }

        # Collect issues and warnings
        foreach ($service in $healthInfo.ServiceStatus.Keys) {
            $serviceStatus = $healthInfo.ServiceStatus[$service]
            if ($serviceStatus.Status -in @('Critical', 'Unhealthy')) {
                $healthInfo.Issues += "$service is $($serviceStatus.Status)"
            } elseif ($serviceStatus.Status -eq 'Warning') {
                $healthInfo.Warnings += "$service has warnings"
            }
        }
    }
    catch {
        $healthInfo.OverallStatus = 'Error'
        $healthInfo.Issues += "Health check failed: $_"
    }

    return $healthInfo
}

function Get-SPOFactoryTenantLicensing {
    [CmdletBinding()]
    param([string]$ClientName)

    return Invoke-SPOFactoryCommand -ScriptBlock {
        try {
            $tenant = Get-PnPTenant
            
            @{
                Plan = $tenant.SharingCapability
                StorageQuotaGB = [math]::Round($tenant.StorageQuota / 1024, 2)
                StorageUsedGB = [math]::Round($tenant.StorageQuotaUsed / 1024, 2)
                StorageAvailableGB = [math]::Round(($tenant.StorageQuota - $tenant.StorageQuotaUsed) / 1024, 2)
                UserCodeMaximumLevel = $tenant.UserCodeMaximumLevel
                CompatibilityRange = $tenant.CompatibilityRange
                MaxSiteStorageQuotaGB = [math]::Round($tenant.PersonalSiteStorageQuota / 1024, 2)
            }
        }
        catch {
            @{
                Error = $_.Exception.Message
                Status = 'Failed to retrieve licensing information'
            }
        }
    } -ClientName $ClientName -Category 'System' -ErrorMessage "Failed to get tenant licensing information" -PassThru
}

function Get-SPOFactoryTenantConfiguration {
    [CmdletBinding()]
    param([string]$ClientName)

    return Invoke-SPOFactoryCommand -ScriptBlock {
        try {
            $tenant = Get-PnPTenant
            
            @{
                SharingCapability = $tenant.SharingCapability
                DefaultSharingLinkType = $tenant.DefaultSharingLinkType
                PreventExternalUsersFromResharing = $tenant.PreventExternalUsersFromResharing
                RequireAnonymousLinksExpireInDays = $tenant.RequireAnonymousLinksExpireInDays
                FileAnonymousLinkType = $tenant.FileAnonymousLinkType
                FolderAnonymousLinkType = $tenant.FolderAnonymousLinkType
                ExternalUserExpirationRequired = $tenant.ExternalUserExpirationRequired
                ExternalUserExpireInDays = $tenant.ExternalUserExpireInDays
                ConditionalAccessPolicy = $tenant.ConditionalAccessPolicy
                AllowedDomainListForSyncClient = $tenant.AllowedDomainListForSyncClient
                BlockedDomainListForSyncClient = $tenant.BlockedDomainListForSyncClient
                DenyAddAndCustomizePages = $tenant.DenyAddAndCustomizePages
                PublicCdnAllowedFileTypes = $tenant.PublicCdnAllowedFileTypes
                PublicCdnEnabled = $tenant.PublicCdnEnabled
            }
        }
        catch {
            @{
                Error = $_.Exception.Message
                Status = 'Failed to retrieve configuration information'
            }
        }
    } -ClientName $ClientName -Category 'System' -ErrorMessage "Failed to get tenant configuration" -PassThru
}

function Get-SPOFactoryMSPMetadata {
    [CmdletBinding()]
    param([string]$ClientName)

    try {
        # Get client configuration
        $clientConfig = Get-SPOFactoryClientConfig -ClientName $ClientName -IncludeDefaults

        # Get connection info
        $connectionInfo = $script:SPOFactoryConnections[$ClientName]

        # Get last activity from tenant registry
        $tenantRegistry = Get-SPOFactoryTenantRegistry -ClientName $ClientName

        return @{
            MSPConfiguration = @{
                DefaultBaseline = $clientConfig.DefaultBaseline
                CustomSettings = $clientConfig.Keys | Where-Object { $_ -notlike '_*' }
            }
            ConnectionInfo = if ($connectionInfo) {
                @{
                    LastConnected = $connectionInfo.Connected
                    LastUsed = $connectionInfo.LastUsed
                    AuthMethod = $connectionInfo.AuthMethod
                    Region = $connectionInfo.Region
                    ConnectionId = $connectionInfo.ConnectionId
                }
            } else { @{} }
            ActivityInfo = if ($tenantRegistry) {
                @{
                    LastActivity = $tenantRegistry.LastActivity
                    TotalOperations = $tenantRegistry.TotalOperations
                    LastHealthCheck = $tenantRegistry.LastHealthCheck
                    LastSecurityScan = $tenantRegistry.LastSecurityScan
                }
            } else { @{} }
            MSPSettings = @{
                MonitoringEnabled = $true
                AlertsConfigured = $clientConfig.AlertEmail -ne $null
                BackupRetention = $script:SPOFactoryConstants.LogRetentionDays
                ComplianceTracking = $script:SPOFactoryConfig.EnableAuditLog
            }
        }
    }
    catch {
        return @{
            Error = $_.Exception.Message
            Status = 'Failed to retrieve MSP metadata'
        }
    }
}

function Get-SPOFactoryTenantRegion {
    <#
    .SYNOPSIS
        Determines the SharePoint Online region from tenant URL.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantUrl
    )

    switch -Regex ($TenantUrl.ToLower()) {
        '\.sharepoint\.com' { return 'Global' }
        '\.sharepoint\.us' { return 'GCC' }
        '\.sharepoint-mil\.us' { return 'GCCH' }
        '\.sharepoint\.de' { return 'Germany' }
        '\.sharepoint\.cn' { return 'China' }
        default { return 'Unknown' }
    }
}

function Update-SPOFactoryTenantRegistry {
    <#
    .SYNOPSIS
        Updates tenant registry with activity information.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $false)]
        [DateTime]$LastConnected,
        
        [Parameter(Mandatory = $false)]
        [string]$Activity,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Metadata
    )

    try {
        $registryPath = Join-Path $script:SPOFactoryConfig.ConfigPath "TenantRegistry"
        if (-not (Test-Path $registryPath)) {
            New-Item -Path $registryPath -ItemType Directory -Force | Out-Null
        }

        $tenantRegistryFile = Join-Path $registryPath "$ClientName.json"
        
        # Load existing registry or create new
        $tenantRegistry = if (Test-Path $tenantRegistryFile) {
            Get-Content $tenantRegistryFile -Raw | ConvertFrom-Json -AsHashtable
        } else {
            @{
                ClientName = $ClientName
                TenantUrl = $TenantUrl
                FirstSeen = Get-Date
                TotalOperations = 0
                Activities = @()
            }
        }

        # Update registry
        if ($TenantUrl) { $tenantRegistry.TenantUrl = $TenantUrl }
        if ($LastConnected) { $tenantRegistry.LastConnected = $LastConnected }
        if ($Activity) {
            $tenantRegistry.LastActivity = Get-Date
            $tenantRegistry.TotalOperations++
            
            # Add activity to history (keep last 100)
            if (-not $tenantRegistry.Activities) { $tenantRegistry.Activities = @() }
            $tenantRegistry.Activities += @{
                Timestamp = Get-Date
                Activity = $Activity
                Metadata = $Metadata
            }
            
            if ($tenantRegistry.Activities.Count -gt 100) {
                $tenantRegistry.Activities = $tenantRegistry.Activities[-100..-1]
            }
        }

        $tenantRegistry.LastUpdated = Get-Date
        
        # Save registry
        $tenantRegistry | ConvertTo-Json -Depth 5 | Out-File -FilePath $tenantRegistryFile -Encoding UTF8
        
        Write-SPOFactoryLog -Message "Updated tenant registry for $ClientName" -Level Debug -ClientName $ClientName -Category 'System'
    }
    catch {
        Write-SPOFactoryLog -Message "Failed to update tenant registry for $ClientName`: $_" -Level Warning -ClientName $ClientName -Category 'System'
    }
}

function Get-SPOFactoryTenantRegistry {
    <#
    .SYNOPSIS
        Retrieves tenant registry information.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName
    )

    try {
        $registryPath = Join-Path $script:SPOFactoryConfig.ConfigPath "TenantRegistry"
        $tenantRegistryFile = Join-Path $registryPath "$ClientName.json"
        
        if (Test-Path $tenantRegistryFile) {
            return Get-Content $tenantRegistryFile -Raw | ConvertFrom-Json -AsHashtable
        } else {
            return $null
        }
    }
    catch {
        Write-SPOFactoryLog -Message "Failed to get tenant registry for $ClientName`: $_" -Level Warning -ClientName $ClientName -Category 'System'
        return $null
    }
}

# Simple cache implementation for tenant information
$script:SPOFactoryCache = @{}

function Get-SPOFactoryCacheItem {
    [CmdletBinding()]
    param(
        [string]$Key,
        [timespan]$MaxAge
    )

    if ($script:SPOFactoryCache.ContainsKey($Key)) {
        $cacheItem = $script:SPOFactoryCache[$Key]
        if ((Get-Date) - $cacheItem.CreatedAt -lt $MaxAge) {
            return $cacheItem.Value
        } else {
            $script:SPOFactoryCache.Remove($Key)
        }
    }
    return $null
}

function Set-SPOFactoryCacheItem {
    [CmdletBinding()]
    param(
        [string]$Key,
        $Value,
        [timespan]$ExpiresIn
    )

    $script:SPOFactoryCache[$Key] = @{
        Value = $Value
        CreatedAt = Get-Date
        ExpiresAt = (Get-Date).Add($ExpiresIn)
    }
}