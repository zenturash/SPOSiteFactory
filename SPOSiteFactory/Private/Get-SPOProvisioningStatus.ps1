function Get-SPOProvisioningStatus {
    <#
    .SYNOPSIS
        Retrieves detailed SharePoint Online site provisioning status for MSP monitoring.

    .DESCRIPTION
        Comprehensive provisioning status function designed for MSP environments managing
        multiple SharePoint Online tenants. Provides detailed status information, error
        detection, and provisioning progress tracking.

    .PARAMETER SiteUrl
        The SharePoint site URL to check status for

    .PARAMETER ClientName
        Client name for MSP tenant isolation and logging

    .PARAMETER IncludeDetailedInfo
        Include additional detailed information about the site

    .PARAMETER CheckSubsites
        Also check subsites if any exist

    .PARAMETER ValidateFeatures
        Validate that required features are activated

    .PARAMETER RetryOnFailure
        Retry status check if initial attempt fails

    .EXAMPLE
        Get-SPOProvisioningStatus -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpTeam" -ClientName "ContosoCorp"

    .EXAMPLE
        $status = Get-SPOProvisioningStatus -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -ClientName "ContosoCorp" -IncludeDetailedInfo

    .EXAMPLE
        Get-SPOProvisioningStatus -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpComm" -ClientName "ContosoCorp" -ValidateFeatures -RetryOnFailure
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeDetailedInfo,
        
        [Parameter(Mandatory = $false)]
        [switch]$CheckSubsites,
        
        [Parameter(Mandatory = $false)]
        [switch]$ValidateFeatures,
        
        [Parameter(Mandatory = $false)]
        [switch]$RetryOnFailure
    )

    begin {
        Write-SPOFactoryLog -Message "Checking provisioning status for: $SiteUrl" -Level Debug -ClientName $ClientName -Category 'Provisioning'
        
        $statusResult = @{
            SiteUrl = $SiteUrl
            IsAvailable = $false
            Status = 'Unknown'
            Details = ''
            HasError = $false
            ErrorMessage = ''
            Site = $null
            ProvisioningStage = 'Unknown'
            LastChecked = Get-Date
            Features = @()
            Subsites = @()
            HealthChecks = @()
            Performance = @{
                ResponseTime = $null
                LoadTime = $null
            }
        }

        $attempts = if ($RetryOnFailure) { 3 } else { 1 }
        $currentAttempt = 0
    }

    process {
        while ($currentAttempt -lt $attempts) {
            $currentAttempt++
            
            try {
                Write-SPOFactoryLog -Message "Status check attempt $currentAttempt/$attempts" -Level Debug -ClientName $ClientName -Category 'Provisioning'
                
                $startTime = Get-Date

                # Primary availability check
                $siteCheckResult = Test-SPOSiteAvailability -SiteUrl $SiteUrl -ClientName $ClientName
                
                if ($siteCheckResult.IsAvailable) {
                    $statusResult.IsAvailable = $true
                    $statusResult.Status = 'Available'
                    $statusResult.Site = $siteCheckResult.Site
                    $statusResult.ProvisioningStage = 'Completed'
                    
                    # Get detailed site information
                    if ($IncludeDetailedInfo) {
                        $detailedInfo = Get-SPOSiteDetailedInfo -SiteUrl $SiteUrl -ClientName $ClientName
                        $statusResult.Site = Merge-SPOSiteInfo -BaseInfo $statusResult.Site -DetailedInfo $detailedInfo
                    }

                    # Validate features if requested
                    if ($ValidateFeatures) {
                        $statusResult.Features = Test-SPOSiteFeatures -SiteUrl $SiteUrl -ClientName $ClientName
                    }

                    # Check subsites if requested
                    if ($CheckSubsites) {
                        $statusResult.Subsites = Get-SPOSiteSubsites -SiteUrl $SiteUrl -ClientName $ClientName
                    }

                    # Perform health checks
                    $statusResult.HealthChecks = Test-SPOSiteHealth -SiteUrl $SiteUrl -ClientName $ClientName

                } elseif ($siteCheckResult.IsProvisioning) {
                    $statusResult.Status = 'Provisioning'
                    $statusResult.Details = $siteCheckResult.ProvisioningDetails
                    $statusResult.ProvisioningStage = $siteCheckResult.Stage
                    
                } elseif ($siteCheckResult.HasError) {
                    $statusResult.HasError = $true
                    $statusResult.ErrorMessage = $siteCheckResult.ErrorMessage
                    $statusResult.Status = 'Error'
                    $statusResult.ProvisioningStage = 'Failed'
                    
                } else {
                    $statusResult.Status = 'NotFound'
                    $statusResult.Details = 'Site does not exist or is not accessible'
                    $statusResult.ProvisioningStage = 'NotStarted'
                }

                # Calculate performance metrics
                $statusResult.Performance.ResponseTime = (Get-Date) - $startTime

                # Log status result
                Write-SPOFactoryLog -Message "Site status: $($statusResult.Status) | Stage: $($statusResult.ProvisioningStage)" -Level Debug -ClientName $ClientName -Category 'Provisioning'

                # Success - break retry loop
                break

            }
            catch {
                Write-SPOFactoryLog -Message "Error checking provisioning status (Attempt $currentAttempt): $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception

                $statusResult.HasError = $true
                $statusResult.ErrorMessage = $_.Exception.Message
                $statusResult.Status = 'Error'

                # If this is the last attempt or we're not retrying, break
                if ($currentAttempt -ge $attempts) {
                    break
                }

                # Wait before retry
                Start-Sleep -Seconds (2 * $currentAttempt)
            }
        }

        return $statusResult
    }
}

function Test-SPOSiteAvailability {
    <#
    .SYNOPSIS
        Tests if a SharePoint site is available and accessible.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        IsAvailable = $false
        IsProvisioning = $false
        HasError = $false
        ErrorMessage = ''
        Site = $null
        ProvisioningDetails = ''
        Stage = 'Unknown'
    }

    try {
        # Try to get site using PnP PowerShell
        $site = Invoke-SPOFactoryCommand -ScriptBlock {
            $siteInfo = Get-PnPSite -Identity $SiteUrl -ErrorAction Stop
            $webInfo = Get-PnPWeb -ErrorAction Stop
            
            return @{
                Id = $siteInfo.Id
                Title = $webInfo.Title
                Url = $siteInfo.Url
                Owner = $siteInfo.Owner
                Created = $siteInfo.Created
                LastContentModifiedDate = $siteInfo.LastContentModifiedDate
                Template = $webInfo.WebTemplate
                Language = $webInfo.Language
                Status = $siteInfo.Status
                StorageQuota = $siteInfo.StorageQuota
                StorageUsage = $siteInfo.StorageUsage
                IsAvailable = $true
            }
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        if ($site) {
            $result.IsAvailable = $true
            $result.Site = $site
            $result.Stage = 'Completed'
        } else {
            # Site not found - check if it's still provisioning
            $provisioningCheck = Test-SPOSiteProvisioningStatus -SiteUrl $SiteUrl -ClientName $ClientName
            
            if ($provisioningCheck.IsProvisioning) {
                $result.IsProvisioning = $true
                $result.ProvisioningDetails = $provisioningCheck.Details
                $result.Stage = $provisioningCheck.Stage
            } else {
                $result.Stage = 'NotFound'
            }
        }

        return $result
    }
    catch {
        # Analyze the error to determine if it's provisioning-related
        $errorMessage = $_.Exception.Message.ToLower()
        
        if ($errorMessage -match 'provisioning|creating|initializing') {
            $result.IsProvisioning = $true
            $result.ProvisioningDetails = "Site appears to be provisioning"
            $result.Stage = 'Provisioning'
        } elseif ($errorMessage -match 'not found|404|does not exist') {
            $result.Stage = 'NotFound'
        } elseif ($errorMessage -match 'access denied|unauthorized|403') {
            $result.HasError = $true
            $result.ErrorMessage = "Access denied to site"
            $result.Stage = 'AccessDenied'
        } elseif ($errorMessage -match 'timeout') {
            $result.HasError = $true
            $result.ErrorMessage = "Timeout accessing site"
            $result.Stage = 'Timeout'
        } else {
            $result.HasError = $true
            $result.ErrorMessage = $_.Exception.Message
            $result.Stage = 'Error'
        }

        return $result
    }
}

function Test-SPOSiteProvisioningStatus {
    <#
    .SYNOPSIS
        Checks if a site is in provisioning status.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        IsProvisioning = $false
        Details = ''
        Stage = 'Unknown'
    }

    try {
        # Try different approaches to detect provisioning
        
        # Method 1: Check using SPO cmdlets if available
        $spoCheck = Invoke-SPOFactoryCommand -ScriptBlock {
            # This would use SPO cmdlets if available
            # For now, return null to indicate not available
            return $null
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        # Method 2: HTTP status check
        try {
            $response = Invoke-WebRequest -Uri $SiteUrl -Method Head -UseBasicParsing -TimeoutSec 10 -ErrorAction SilentlyContinue
            
            if ($response.StatusCode -eq 503) {
                $result.IsProvisioning = $true
                $result.Details = "Site returning 503 Service Unavailable (likely provisioning)"
                $result.Stage = 'Provisioning'
            }
        }
        catch {
            # Web request failed - could indicate provisioning
            if ($_.Exception.Message -match '503|service unavailable|provisioning') {
                $result.IsProvisioning = $true
                $result.Details = "Web request indicates provisioning status"
                $result.Stage = 'Provisioning'
            }
        }

        return $result
    }
    catch {
        return $result
    }
}

function Get-SPOSiteDetailedInfo {
    <#
    .SYNOPSIS
        Retrieves comprehensive site information.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $detailedInfo = Invoke-SPOFactoryCommand -ScriptBlock {
            $web = Get-PnPWeb -Includes Lists,Features,UserCustomActions,Navigation
            $site = Get-PnPSite -Includes Features,UserCustomActions,EventReceivers
            
            return @{
                Web = @{
                    Description = $web.Description
                    MasterUrl = $web.MasterUrl
                    CustomMasterUrl = $web.CustomMasterUrl
                    AlternateCssUrl = $web.AlternateCssUrl
                    SiteLogoUrl = $web.SiteLogoUrl
                    QuickLaunchEnabled = $web.QuickLaunchEnabled
                    TreeViewEnabled = $web.TreeViewEnabled
                    UIVersion = $web.UIVersion
                    Configuration = $web.Configuration
                    Features = $web.Features | Select-Object DisplayName, DefinitionId
                    Lists = $web.Lists | Select-Object Title, BaseType, ItemCount, Hidden
                    Navigation = @{
                        TopNavigationBar = $web.Navigation.TopNavigationBar
                        QuickLaunch = $web.Navigation.QuickLaunch
                    }
                }
                Site = @{
                    Features = $site.Features | Select-Object DisplayName, DefinitionId
                    EventReceivers = $site.EventReceivers | Select-Object ReceiverName, ReceiverUrl, ReceiverAssembly
                    UserCustomActions = $site.UserCustomActions | Select-Object Name, Location, ScriptBlock, ScriptSrc
                    ServerRelativeUrl = $site.ServerRelativeUrl
                    Url = $site.Url
                    ReadOnly = $site.ReadOnly
                    ShareByEmailEnabled = $site.ShareByEmailEnabled
                }
            }
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        return $detailedInfo
    }
    catch {
        Write-SPOFactoryLog -Message "Error retrieving detailed site info: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
        return @{}
    }
}

function Test-SPOSiteFeatures {
    <#
    .SYNOPSIS
        Tests and validates SharePoint site features.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $featureResults = Invoke-SPOFactoryCommand -ScriptBlock {
            $siteFeatures = Get-PnPFeature -Scope Site
            $webFeatures = Get-PnPFeature -Scope Web
            
            return @{
                SiteFeatures = $siteFeatures | Select-Object DisplayName, DefinitionId, Scope
                WebFeatures = $webFeatures | Select-Object DisplayName, DefinitionId, Scope
                TotalCount = $siteFeatures.Count + $webFeatures.Count
                CriticalFeatures = @{
                    SharePointServerPublishing = ($siteFeatures | Where-Object { $_.DefinitionId -eq "f6924d36-2fa8-4f0b-b16d-06b7250180fa" }) -ne $null
                    DocumentSets = ($siteFeatures | Where-Object { $_.DefinitionId -eq "3bae86a2-776d-499d-9db8-fa4cdc7884f8" }) -ne $null
                    WorkflowTask = ($webFeatures | Where-Object { $_.DefinitionId -eq "57311b7a-9afd-4ff0-866e-9393ad6647b1" }) -ne $null
                }
            }
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        return $(if ($featureResults) { $featureResults } else { @{ SiteFeatures = @(); WebFeatures = @(); TotalCount = 0; CriticalFeatures = @{} } })
    }
    catch {
        Write-SPOFactoryLog -Message "Error checking site features: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
        return @{ SiteFeatures = @(); WebFeatures = @(); TotalCount = 0; CriticalFeatures = @{} }
    }
}

function Get-SPOSiteSubsites {
    <#
    .SYNOPSIS
        Retrieves subsites information.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $subsites = Invoke-SPOFactoryCommand -ScriptBlock {
            Get-PnPSubWeb -Recurse | Select-Object Title, Url, Created, WebTemplate, Language
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        return $(if ($subsites) { $subsites } else { @() })
    }
    catch {
        Write-SPOFactoryLog -Message "Error retrieving subsites: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
        return @()
    }
}

function Test-SPOSiteHealth {
    <#
    .SYNOPSIS
        Performs basic health checks on the site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $healthChecks = @()

    try {
        $healthResults = Invoke-SPOFactoryCommand -ScriptBlock {
            $web = Get-PnPWeb -Includes Lists
            $lists = $web.Lists
            
            return @{
                HasDocumentLibrary = ($lists | Where-Object { $_.BaseTemplate -eq 101 }).Count -gt 0
                HasSystemLists = ($lists | Where-Object { $_.Hidden -eq $true }).Count -gt 0
                ListCount = $lists.Count
                WebTitle = $web.Title
                WebId = $web.Id
                HasCustomization = $web.CustomMasterUrl -ne $web.MasterUrl
            }
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        if ($healthResults) {
            $healthChecks += @{ Check = 'DocumentLibrary'; Status = if ($healthResults.HasDocumentLibrary) { 'Pass' } else { 'Warning' }; Details = "Document library availability" }
            $healthChecks += @{ Check = 'SystemLists'; Status = if ($healthResults.HasSystemLists) { 'Pass' } else { 'Warning' }; Details = "System lists availability" }
            $healthChecks += @{ Check = 'ListCount'; Status = if ($healthResults.ListCount -gt 0) { 'Pass' } else { 'Warning' }; Details = "Total lists: $($healthResults.ListCount)" }
            $healthChecks += @{ Check = 'WebTitle'; Status = if ($healthResults.WebTitle) { 'Pass' } else { 'Warning' }; Details = "Web title configured" }
        } else {
            $healthChecks += @{ Check = 'BasicAccess'; Status = 'Error'; Details = 'Unable to perform health checks' }
        }

        return $healthChecks
    }
    catch {
        $healthChecks += @{ Check = 'HealthCheck'; Status = 'Error'; Details = "Error: $($_.Exception.Message)" }
        return $healthChecks
    }
}

function Merge-SPOSiteInfo {
    <#
    .SYNOPSIS
        Merges base site info with detailed information.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$BaseInfo,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$DetailedInfo
    )

    if (-not $DetailedInfo -or $DetailedInfo.Count -eq 0) {
        return $BaseInfo
    }

    $merged = $BaseInfo.Clone()
    
    foreach ($key in $DetailedInfo.Keys) {
        $merged[$key] = $DetailedInfo[$key]
    }

    return $merged
}