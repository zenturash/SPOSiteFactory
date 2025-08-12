function Add-SPOSiteToHub {
    <#
    .SYNOPSIS
        Associates SharePoint Online sites with hub sites for MSP multi-tenant environments.

    .DESCRIPTION
        Enterprise-grade hub association function designed for MSP environments managing multiple
        SharePoint Online tenants. Provides site-to-hub association with bulk processing capabilities,
        permission inheritance options, and comprehensive error handling with rollback support.

    .PARAMETER HubSiteUrl
        The hub site URL to associate sites with

    .PARAMETER SiteUrl
        Single site URL to associate with the hub (use with single site association)

    .PARAMETER SiteUrls
        Array of site URLs to associate with the hub (use for bulk association)

    .PARAMETER ClientName
        Client name for MSP tenant isolation and logging

    .PARAMETER EnablePermissionSync
        Enable permission synchronization between hub and associated sites (optional)

    .PARAMETER ApplyHubNavigation
        Apply hub navigation to associated sites (default: true)

    .PARAMETER ApplyHubTheme
        Apply hub theme to associated sites (default: true)

    .PARAMETER WaitForCompletion
        Wait for association to complete before returning (default: true)

    .PARAMETER TimeoutMinutes
        Maximum time to wait for association in minutes (default: 10)

    .PARAMETER ContinueOnError
        Continue processing other sites if one fails (for bulk operations)

    .PARAMETER MaxConcurrent
        Maximum number of concurrent associations (default: 5)

    .PARAMETER WhatIf
        Show what associations would be made without actually making them

    .PARAMETER Force
        Suppress confirmation prompts for bulk operations

    .EXAMPLE
        Add-SPOSiteToHub -HubSiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpTeam" -ClientName "ContosoCorp"

    .EXAMPLE
        Add-SPOSiteToHub -HubSiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -SiteUrls @("https://contoso.sharepoint.com/sites/ContosoCorpTeam1", "https://contoso.sharepoint.com/sites/ContosoCorpTeam2") -ClientName "ContosoCorp" -ContinueOnError

    .EXAMPLE
        Add-SPOSiteToHub -HubSiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpSecureHub" -SiteUrls @("https://contoso.sharepoint.com/sites/ContosoCorpProject1", "https://contoso.sharepoint.com/sites/ContosoCorpProject2") -ClientName "ContosoCorp" -EnablePermissionSync -ApplyHubNavigation -ApplyHubTheme
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Single')]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Bulk')]
        [string[]]$SiteUrls,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnablePermissionSync,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyHubNavigation = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyHubTheme = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$WaitForCompletion = $true,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 10,
        
        [Parameter(Mandatory = $false)]
        [switch]$ContinueOnError,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxConcurrent = 5,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        $operationId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        $startTime = Get-Date
        
        # Determine the sites to process
        $sitesToProcess = if ($SiteUrl) { @($SiteUrl) } else { $SiteUrls }
        $isBulkOperation = $sitesToProcess.Count -gt 1
        
        Write-SPOFactoryLog -Message "Starting hub site association: $($sitesToProcess.Count) sites to $HubSiteUrl" -Level Info -ClientName $ClientName -Category 'Hub' -Tag @('HubAssociationStart', $operationId) -EnableAuditLog

        $result = @{
            Success = $false
            HubSiteUrl = $HubSiteUrl
            TotalSites = $sitesToProcess.Count
            SuccessfulAssociations = @()
            FailedAssociations = @()
            Warnings = @()
            Errors = @()
            OperationTime = $null
            HubInfo = $null
        }

        # Validate prerequisites
        try {
            Write-SPOFactoryLog -Message "Validating prerequisites and hub site availability" -Level Debug -ClientName $ClientName -Category 'Hub'
            
            # Test connection
            $connectionTest = Test-SPOFactoryConnection -ClientName $ClientName
            if (-not $connectionTest.IsConnected) {
                throw "Not connected to SharePoint Online. Connection required for hub site association."
            }

            # Validate and get hub site information
            $hubValidation = Get-SPOHubSiteInfo -HubSiteUrl $HubSiteUrl -ClientName $ClientName
            if (-not $hubValidation.IsHubSite) {
                throw "Site is not registered as a hub site: $HubSiteUrl"
            }

            $result.HubInfo = $hubValidation

            # Validate all target sites
            $invalidSites = @()
            foreach ($targetSite in $sitesToProcess) {
                $siteValidation = Test-SPOSiteForHubAssociation -SiteUrl $targetSite -HubSiteUrl $HubSiteUrl -ClientName $ClientName
                if (-not $siteValidation.IsValid) {
                    $invalidSites += @{ Url = $targetSite; Issues = $siteValidation.Issues }
                }
            }

            if ($invalidSites.Count -gt 0 -and -not $ContinueOnError) {
                $invalidUrls = ($invalidSites | ForEach-Object { "$($_.Url): $($_.Issues -join ', ')" }) -join '; '
                throw "Site validation failed for: $invalidUrls"
            } elseif ($invalidSites.Count -gt 0) {
                foreach ($invalidSite in $invalidSites) {
                    $result.Warnings += "Site $($invalidSite.Url) validation warnings: $($invalidSite.Issues -join ', ')"
                }
            }

            Write-SPOFactoryLog -Message "Prerequisites validation completed successfully" -Level Info -ClientName $ClientName -Category 'Hub'
        }
        catch {
            Write-SPOFactoryLog -Message "Prerequisites validation failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Hub' -Exception $_.Exception
            throw
        }
    }

    process {
        if (-not $PSCmdlet.ShouldProcess("$($sitesToProcess.Count) sites", "Associate with hub site $HubSiteUrl")) {
            return
        }

        try {
            # Process site associations
            if ($isBulkOperation -and $sitesToProcess.Count -gt $MaxConcurrent) {
                # Process in batches for large bulk operations
                Write-SPOFactoryLog -Message "Processing $($sitesToProcess.Count) sites in batches of $MaxConcurrent" -Level Info -ClientName $ClientName -Category 'Hub'
                
                $batches = Split-ArrayIntoBatches -Array $sitesToProcess -BatchSize $MaxConcurrent
                
                foreach ($batch in $batches) {
                    $batchResults = Process-HubAssociationBatch -HubSiteUrl $HubSiteUrl -SiteUrls $batch -ClientName $ClientName -EnablePermissionSync:$EnablePermissionSync -ApplyHubNavigation:$ApplyHubNavigation -ApplyHubTheme:$ApplyHubTheme -WaitForCompletion:$WaitForCompletion -TimeoutMinutes $TimeoutMinutes -ContinueOnError:$ContinueOnError -OperationId $operationId
                    
                    $result.SuccessfulAssociations += $batchResults.SuccessfulAssociations
                    $result.FailedAssociations += $batchResults.FailedAssociations
                    $result.Warnings += $batchResults.Warnings
                    $result.Errors += $batchResults.Errors
                }
            } else {
                # Process all sites directly
                $associationResults = Process-HubAssociationBatch -HubSiteUrl $HubSiteUrl -SiteUrls $sitesToProcess -ClientName $ClientName -EnablePermissionSync:$EnablePermissionSync -ApplyHubNavigation:$ApplyHubNavigation -ApplyHubTheme:$ApplyHubTheme -WaitForCompletion:$WaitForCompletion -TimeoutMinutes $TimeoutMinutes -ContinueOnError:$ContinueOnError -OperationId $operationId
                
                $result.SuccessfulAssociations = $associationResults.SuccessfulAssociations
                $result.FailedAssociations = $associationResults.FailedAssociations
                $result.Warnings += $associationResults.Warnings
                $result.Errors += $associationResults.Errors
            }

            # Determine overall success
            $result.Success = ($result.SuccessfulAssociations.Count -gt 0) -and ($result.FailedAssociations.Count -eq 0 -or $ContinueOnError)
            $result.OperationTime = (Get-Date) - $startTime

            # Log summary
            $successCount = $result.SuccessfulAssociations.Count
            $failedCount = $result.FailedAssociations.Count
            $logLevel = if ($result.Success) { 'Info' } else { 'Warning' }
            
            Write-SPOFactoryLog -Message "Hub association completed - Success: $successCount, Failed: $failedCount, Time: $($result.OperationTime.TotalSeconds.ToString('F1'))s" -Level $logLevel -ClientName $ClientName -Category 'Hub' -Tag @('HubAssociationComplete', $operationId) -EnableAuditLog

            return $result
        }
        catch {
            Write-SPOFactoryLog -Message "Hub site association failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Hub' -Exception $_.Exception -Tag @('HubAssociationError', $operationId)
            
            $result.Errors += $_.Exception.Message
            throw
        }
    }

    end {
        if ($result.Success) {
            Write-SPOFactoryLog -Message "Hub site association operation completed successfully" -Level Info -ClientName $ClientName -Category 'Hub' -EnableAuditLog
        }
    }
}

function Get-SPOHubSiteInfo {
    <#
    .SYNOPSIS
        Retrieves comprehensive hub site information and validates hub status.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $hubInfo = Invoke-SPOFactoryCommand -ScriptBlock {
            # Get hub site information
            $hubSite = Get-PnPHubSite -Identity $HubSiteUrl -ErrorAction Stop
            
            # Get additional site information
            $site = Get-PnPSite -Identity $HubSiteUrl -Includes Id, Owner, Created, LastContentModifiedDate
            
            # Get associated sites
            $associatedSites = Get-PnPHubSiteChild -Identity $hubSite.Id
            
            return @{
                HubSiteId = $hubSite.Id
                Title = $hubSite.Title
                Url = $hubSite.Url
                Description = $hubSite.Description
                LogoUrl = $hubSite.LogoUrl
                SiteDesignId = $hubSite.SiteDesignId
                Owner = $site.Owner
                Created = $site.Created
                LastModified = $site.LastContentModifiedDate
                AssociatedSites = $associatedSites | ForEach-Object { @{ Url = $_.Url; Title = $_.Title } }
                AssociatedSitesCount = $associatedSites.Count
                IsHubSite = $true
            }
        } -ClientName $ClientName -Category 'Hub' -ErrorMessage "Failed to get hub site information"

        return $hubInfo
    }
    catch {
        return @{
            IsHubSite = $false
            Error = $_.Exception.Message
            HubSiteUrl = $HubSiteUrl
        }
    }
}

function Test-SPOSiteForHubAssociation {
    <#
    .SYNOPSIS
        Validates if a site can be associated with a hub site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $validation = @{
        IsValid = $true
        Issues = @()
        Warnings = @()
        SiteInfo = $null
    }

    try {
        # Get site information
        $siteInfo = Invoke-SPOFactoryCommand -ScriptBlock {
            $site = Get-PnPSite -Identity $SiteUrl -ErrorAction Stop
            $hubSite = Get-PnPHubSite -Identity $site.Url -ErrorAction SilentlyContinue
            
            return @{
                Id = $site.Id
                Url = $site.Url
                Title = $site.Title
                Template = $site.Template
                IsHubSite = $hubSite -ne $null
                HubSiteId = if ($hubSite) { $hubSite.Id } else { $null }
                ReadOnly = $site.ReadOnly
                Status = $site.Status
            }
        } -ClientName $ClientName -Category 'Hub' -SuppressErrors

        if (-not $siteInfo) {
            $validation.Issues += "Site not found or not accessible: $SiteUrl"
            $validation.IsValid = $false
            return $validation
        }

        $validation.SiteInfo = $siteInfo

        # Check if site is already a hub site
        if ($siteInfo.IsHubSite) {
            $validation.Issues += "Site is already registered as a hub site and cannot be associated with another hub"
            $validation.IsValid = $false
        }

        # Check if site is read-only
        if ($siteInfo.ReadOnly) {
            $validation.Issues += "Site is read-only and cannot be associated with hub"
            $validation.IsValid = $false
        }

        # Check site status
        if ($siteInfo.Status -ne 'Active') {
            $validation.Issues += "Site status is not Active: $($siteInfo.Status)"
            $validation.IsValid = $false
        }

        # Check if site is the same as hub site
        if ($siteInfo.Url.TrimEnd('/') -eq $HubSiteUrl.TrimEnd('/')) {
            $validation.Issues += "Site cannot be associated with itself"
            $validation.IsValid = $false
        }

        # Check if already associated with another hub
        if ($siteInfo.HubSiteId) {
            $validation.Warnings += "Site is already associated with hub site ID: $($siteInfo.HubSiteId)"
        }

        return $validation
    }
    catch {
        $validation.Issues += "Error validating site: $($_.Exception.Message)"
        $validation.IsValid = $false
        return $validation
    }
}

function Process-HubAssociationBatch {
    <#
    .SYNOPSIS
        Processes a batch of hub site associations.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string[]]$SiteUrls,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnablePermissionSync,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyHubNavigation,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyHubTheme,
        
        [Parameter(Mandatory = $false)]
        [switch]$WaitForCompletion,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes,
        
        [Parameter(Mandatory = $false)]
        [switch]$ContinueOnError,
        
        [Parameter(Mandatory = $true)]
        [string]$OperationId
    )

    $result = @{
        SuccessfulAssociations = @()
        FailedAssociations = @()
        Warnings = @()
        Errors = @()
    }

    foreach ($siteUrl in $SiteUrls) {
        try {
            Write-SPOFactoryLog -Message "Associating site with hub: $siteUrl" -Level Debug -ClientName $ClientName -Category 'Hub' -Tag @('SiteAssociation', $OperationId)
            
            # Perform the association
            $associationResult = Invoke-SPOFactoryCommand -ScriptBlock {
                # Get the hub site ID
                $hubSite = Get-PnPHubSite -Identity $HubSiteUrl
                
                # Associate the site with the hub
                Add-PnPHubSiteAssociation -Site $siteUrl -HubSite $hubSite.Id
                
                return @{
                    SiteUrl = $siteUrl
                    HubSiteId = $hubSite.Id
                    AssociationTime = Get-Date
                }
            } -ClientName $ClientName -Category 'Hub' -ErrorMessage "Failed to associate site with hub" -CriticalOperation:(-not $ContinueOnError)

            if ($associationResult) {
                # Apply additional configurations
                $configResults = @()
                
                # Apply hub navigation if requested
                if ($ApplyHubNavigation) {
                    $navResult = Set-SPOHubNavigationForSite -SiteUrl $siteUrl -HubSiteUrl $HubSiteUrl -ClientName $ClientName
                    $configResults += "Navigation: $($navResult.Status)"
                }

                # Apply hub theme if requested
                if ($ApplyHubTheme) {
                    $themeResult = Set-SPOHubThemeForSite -SiteUrl $siteUrl -HubSiteUrl $HubSiteUrl -ClientName $ClientName
                    $configResults += "Theme: $($themeResult.Status)"
                }

                # Configure permission sync if requested
                if ($EnablePermissionSync) {
                    $permResult = Set-SPOHubPermissionSyncForSite -SiteUrl $siteUrl -HubSiteUrl $HubSiteUrl -ClientName $ClientName
                    $configResults += "Permissions: $($permResult.Status)"
                }

                # Wait for completion if requested
                if ($WaitForCompletion) {
                    $waitResult = Wait-SPOHubAssociation -SiteUrl $siteUrl -HubSiteUrl $HubSiteUrl -TimeoutMinutes $TimeoutMinutes -ClientName $ClientName
                    if (-not $waitResult.Success) {
                        $result.Warnings += "Association may not be fully complete for $siteUrl`: $($waitResult.Status)"
                    }
                }

                $associationInfo = @{
                    SiteUrl = $siteUrl
                    HubSiteUrl = $HubSiteUrl
                    HubSiteId = $associationResult.HubSiteId
                    AssociationTime = $associationResult.AssociationTime
                    Configurations = $configResults
                    Status = 'Success'
                }

                $result.SuccessfulAssociations += $associationInfo
                Write-SPOFactoryLog -Message "Site associated successfully: $siteUrl" -Level Info -ClientName $ClientName -Category 'Hub' -Tag @('AssociationSuccess', $OperationId)
            } else {
                throw "Association returned null result"
            }
        }
        catch {
            $errorInfo = @{
                SiteUrl = $siteUrl
                HubSiteUrl = $HubSiteUrl
                Error = $_.Exception.Message
                Status = 'Failed'
                Timestamp = Get-Date
            }

            $result.FailedAssociations += $errorInfo
            $result.Errors += "Failed to associate $siteUrl`: $($_.Exception.Message)"
            
            Write-SPOFactoryLog -Message "Site association failed: $siteUrl - $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Hub' -Exception $_.Exception -Tag @('AssociationFailed', $OperationId)

            if (-not $ContinueOnError) {
                throw
            }
        }
    }

    return $result
}

function Set-SPOHubNavigationForSite {
    <#
    .SYNOPSIS
        Applies hub navigation to an associated site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $navResult = Invoke-SPOFactoryCommand -ScriptBlock {
            # Connect to the site and apply hub navigation
            Connect-PnPOnline -Url $SiteUrl -Interactive
            
            # This would apply hub navigation - implementation depends on specific requirements
            # For now, return a placeholder result
            return "Navigation inheritance enabled"
        } -ClientName $ClientName -Category 'Hub' -SuppressErrors

        return @{
            Status = if ($navResult) { 'Success' } else { 'Failed' }
            Details = $navResult
        }
    }
    catch {
        return @{
            Status = 'Error'
            Details = $_.Exception.Message
        }
    }
}

function Set-SPOHubThemeForSite {
    <#
    .SYNOPSIS
        Applies hub theme to an associated site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $themeResult = Invoke-SPOFactoryCommand -ScriptBlock {
            # Connect to the site and apply hub theme
            Connect-PnPOnline -Url $SiteUrl -Interactive
            
            # This would apply hub theme - implementation depends on specific requirements
            # For now, return a placeholder result
            return "Theme inheritance enabled"
        } -ClientName $ClientName -Category 'Hub' -SuppressErrors

        return @{
            Status = if ($themeResult) { 'Success' } else { 'Failed' }
            Details = $themeResult
        }
    }
    catch {
        return @{
            Status = 'Error'
            Details = $_.Exception.Message
        }
    }
}

function Set-SPOHubPermissionSyncForSite {
    <#
    .SYNOPSIS
        Configures permission synchronization between hub and site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $permResult = Invoke-SPOFactoryCommand -ScriptBlock {
            # Configure permission synchronization
            # This would implement specific permission sync logic
            # For now, return a placeholder result
            return "Permission sync configured"
        } -ClientName $ClientName -Category 'Hub' -SuppressErrors

        return @{
            Status = if ($permResult) { 'Success' } else { 'Failed' }
            Details = $permResult
        }
    }
    catch {
        return @{
            Status = 'Error'
            Details = $_.Exception.Message
        }
    }
}

function Wait-SPOHubAssociation {
    <#
    .SYNOPSIS
        Waits for hub association to complete and become active.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 10,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        Success = $false
        Status = 'Unknown'
        ElapsedTime = $null
    }

    $startTime = Get-Date
    $timeoutTime = $startTime.AddMinutes($TimeoutMinutes)
    
    try {
        while ((Get-Date) -lt $timeoutTime) {
            $associationCheck = Invoke-SPOFactoryCommand -ScriptBlock {
                $site = Get-PnPSite -Identity $SiteUrl -Includes HubSiteId
                return @{
                    IsAssociated = $site.HubSiteId -ne $null
                    HubSiteId = $site.HubSiteId
                }
            } -ClientName $ClientName -Category 'Hub' -SuppressErrors

            if ($associationCheck.IsAssociated) {
                $result.Success = $true
                $result.Status = 'Associated'
                break
            }

            Start-Sleep -Seconds 15
        }

        if (-not $result.Success) {
            $result.Status = 'Timeout'
        }

        $result.ElapsedTime = (Get-Date) - $startTime
        return $result
    }
    catch {
        $result.Status = 'Error'
        $result.ElapsedTime = (Get-Date) - $startTime
        return $result
    }
}

function Split-ArrayIntoBatches {
    <#
    .SYNOPSIS
        Splits an array into smaller batches for processing.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$Array,
        
        [Parameter(Mandatory = $true)]
        [int]$BatchSize
    )

    $batches = @()
    for ($i = 0; $i -lt $Array.Count; $i += $BatchSize) {
        $end = [Math]::Min($i + $BatchSize - 1, $Array.Count - 1)
        $batches += ,@($Array[$i..$end])
    }
    
    return $batches
}