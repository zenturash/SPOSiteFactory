function Test-SPOSiteUrl {
    <#
    .SYNOPSIS
        Validates SharePoint Online site URLs for MSP multi-tenant environments.

    .DESCRIPTION
        Comprehensive URL validation function designed for MSP environments managing multiple
        SharePoint Online tenants. Validates URL format, availability, compliance with MSP
        naming conventions, and checks for conflicts with existing sites.

    .PARAMETER SiteUrl
        The SharePoint site URL to validate

    .PARAMETER ClientName
        Client name for MSP tenant isolation and naming validation

    .PARAMETER SiteType
        Type of site (TeamSite, CommunicationSite, HubSite) for specific validation rules

    .PARAMETER TenantUrl
        Base tenant URL for validation context

    .PARAMETER CheckAvailability
        Verify that the URL is available (not already in use)

    .PARAMETER AllowExisting
        Allow validation to pass even if site already exists

    .PARAMETER MSPNamingConvention
        Enforce MSP naming conventions (e.g., /sites/ClientName-SiteName format)

    .EXAMPLE
        Test-SPOSiteUrl -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpTeam" -ClientName "ContosoCorp" -SiteType "TeamSite"

    .EXAMPLE
        Test-SPOSiteUrl -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -ClientName "ContosoCorp" -SiteType "HubSite" -CheckAvailability

    .EXAMPLE
        $urlValidation = Test-SPOSiteUrl -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpComm" -ClientName "ContosoCorp" -SiteType "CommunicationSite" -MSPNamingConvention
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('TeamSite', 'CommunicationSite', 'HubSite')]
        [string]$SiteType = 'TeamSite',
        
        [Parameter(Mandatory = $false)]
        [string]$TenantUrl,
        
        [Parameter(Mandatory = $false)]
        [switch]$CheckAvailability,
        
        [Parameter(Mandatory = $false)]
        [switch]$AllowExisting,
        
        [Parameter(Mandatory = $false)]
        [switch]$MSPNamingConvention
    )

    begin {
        Write-SPOFactoryLog -Message "Starting URL validation for: $SiteUrl" -Level Debug -ClientName $ClientName -Category 'Provisioning'
        
        $validationResult = @{
            IsValid = $false
            IsAvailable = $false
            ValidationErrors = @()
            ValidationWarnings = @()
            SuggestedUrl = $null
            MSPCompliant = $false
            ClientIsolated = $false
        }

        # Get tenant context if not provided
        if (-not $TenantUrl) {
            try {
                $tenantInfo = Get-SPOFactoryTenantInfo -ClientName $ClientName
                $TenantUrl = $tenantInfo.TenantUrl
            }
            catch {
                Write-SPOFactoryLog -Message "Could not determine tenant URL, proceeding with basic validation" -Level Warning -ClientName $ClientName -Category 'Provisioning'
            }
        }
    }

    process {
        try {
            # Basic URL format validation
            if (-not (Test-SPOFactoryUrlFormat -Url $SiteUrl)) {
                $validationResult.ValidationErrors += "Invalid URL format"
                return $validationResult
            }

            # Parse URL components
            $uri = [System.Uri]$SiteUrl
            $sitePath = $uri.LocalPath
            $siteCollection = $sitePath -replace '^/sites/', ''

            Write-SPOFactoryLog -Message "Validating URL components - Host: $($uri.Host), Path: $sitePath" -Level Debug -ClientName $ClientName -Category 'Provisioning'

            # Validate URL scheme and structure
            if ($uri.Scheme -ne 'https') {
                $validationResult.ValidationErrors += "URL must use HTTPS scheme"
            }

            if (-not $uri.Host.EndsWith('.sharepoint.com')) {
                $validationResult.ValidationErrors += "URL must be a valid SharePoint Online URL"
            }

            if (-not $sitePath.StartsWith('/sites/')) {
                $validationResult.ValidationErrors += "URL must be a site collection under /sites/"
            }

            # Validate site collection name
            $siteNameValidation = Test-SPOFactorySiteCollectionName -Name $siteCollection -ClientName $ClientName -SiteType $SiteType
            if (-not $siteNameValidation.IsValid) {
                $validationResult.ValidationErrors += $siteNameValidation.Errors
            }

            # MSP naming convention validation
            if ($MSPNamingConvention -and $ClientName) {
                $mspValidation = Test-SPOFactoryMSPNaming -SiteCollection $siteCollection -ClientName $ClientName -SiteType $SiteType
                $validationResult.MSPCompliant = $mspValidation.IsCompliant
                $validationResult.ClientIsolated = $mspValidation.IsClientIsolated
                
                if (-not $mspValidation.IsCompliant) {
                    $validationResult.ValidationWarnings += $mspValidation.Warnings
                    $validationResult.SuggestedUrl = $mspValidation.SuggestedUrl
                }
            }

            # Check for reserved names and conflicts
            $reservedNames = @('admin', 'api', 'www', 'mail', 'ftp', 'localhost', 'test', 'dev', 'staging', 'prod', 'production')
            if ($siteCollection.ToLower() -in $reservedNames) {
                $validationResult.ValidationErrors += "Site collection name '$siteCollection' is reserved and cannot be used"
            }

            # Length validation
            if ($siteCollection.Length -gt 64) {
                $validationResult.ValidationErrors += "Site collection name exceeds maximum length of 64 characters"
            }

            if ($siteCollection.Length -lt 3) {
                $validationResult.ValidationErrors += "Site collection name must be at least 3 characters long"
            }

            # Character validation
            if ($siteCollection -match '[^a-zA-Z0-9\-]') {
                $validationResult.ValidationErrors += "Site collection name contains invalid characters. Only letters, numbers, and hyphens are allowed"
            }

            if ($siteCollection.StartsWith('-') -or $siteCollection.EndsWith('-')) {
                $validationResult.ValidationErrors += "Site collection name cannot start or end with a hyphen"
            }

            # Tenant-specific validation
            if ($TenantUrl) {
                $expectedHost = ([System.Uri]$TenantUrl).Host
                if ($uri.Host -ne $expectedHost) {
                    $validationResult.ValidationErrors += "URL host '$($uri.Host)' does not match tenant host '$expectedHost'"
                }
            }

            # Check availability if requested
            if ($CheckAvailability) {
                Write-SPOFactoryLog -Message "Checking site URL availability" -Level Debug -ClientName $ClientName -Category 'Provisioning'
                
                $availabilityResult = Test-SPOSiteUrlAvailability -SiteUrl $SiteUrl -ClientName $ClientName
                $validationResult.IsAvailable = $availabilityResult.IsAvailable
                
                if (-not $availabilityResult.IsAvailable -and -not $AllowExisting) {
                    $validationResult.ValidationErrors += "Site URL is already in use"
                }
                
                if ($availabilityResult.AlternativeUrls) {
                    $validationResult.SuggestedUrl = $availabilityResult.AlternativeUrls[0]
                }
            }

            # Set overall validation status
            $validationResult.IsValid = ($validationResult.ValidationErrors.Count -eq 0)

            # Log validation results
            if ($validationResult.IsValid) {
                Write-SPOFactoryLog -Message "URL validation successful" -Level Info -ClientName $ClientName -Category 'Provisioning'
            } else {
                Write-SPOFactoryLog -Message "URL validation failed: $($validationResult.ValidationErrors -join '; ')" -Level Warning -ClientName $ClientName -Category 'Provisioning'
            }

            if ($validationResult.ValidationWarnings.Count -gt 0) {
                Write-SPOFactoryLog -Message "URL validation warnings: $($validationResult.ValidationWarnings -join '; ')" -Level Warning -ClientName $ClientName -Category 'Provisioning'
            }

            return $validationResult
        }
        catch {
            Write-SPOFactoryLog -Message "Error during URL validation: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
            
            $validationResult.ValidationErrors += "Validation process failed: $($_.Exception.Message)"
            return $validationResult
        }
    }
}

function Test-SPOFactoryUrlFormat {
    <#
    .SYNOPSIS
        Tests basic URL format validity.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Url
    )

    try {
        $uri = [System.Uri]$Url
        return $uri.IsAbsoluteUri
    }
    catch {
        return $false
    }
}

function Test-SPOFactorySiteCollectionName {
    <#
    .SYNOPSIS
        Validates site collection name according to SharePoint rules.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$SiteType
    )

    $validation = @{
        IsValid = $true
        Errors = @()
        Warnings = @()
    }

    # Check for empty or null
    if ([string]::IsNullOrWhiteSpace($Name)) {
        $validation.IsValid = $false
        $validation.Errors += "Site collection name cannot be empty"
        return $validation
    }

    # Character set validation
    if ($Name -notmatch '^[a-zA-Z0-9\-]+$') {
        $validation.IsValid = $false
        $validation.Errors += "Site collection name can only contain letters, numbers, and hyphens"
    }

    # Start/end character validation
    if ($Name -match '^-' -or $Name -match '-$') {
        $validation.IsValid = $false
        $validation.Errors += "Site collection name cannot start or end with a hyphen"
    }

    # Double hyphen validation
    if ($Name -match '--') {
        $validation.IsValid = $false
        $validation.Errors += "Site collection name cannot contain consecutive hyphens"
    }

    # Length validation
    if ($Name.Length -lt 3) {
        $validation.IsValid = $false
        $validation.Errors += "Site collection name must be at least 3 characters long"
    }

    if ($Name.Length -gt 64) {
        $validation.IsValid = $false
        $validation.Errors += "Site collection name cannot exceed 64 characters"
    }

    return $validation
}

function Test-SPOFactoryMSPNaming {
    <#
    .SYNOPSIS
        Validates MSP naming conventions for client isolation.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteCollection,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$SiteType
    )

    $validation = @{
        IsCompliant = $false
        IsClientIsolated = $false
        Warnings = @()
        SuggestedUrl = $null
    }

    # Expected format: ClientName-SiteName
    $expectedPrefix = "$ClientName-"
    $validation.IsClientIsolated = $SiteCollection.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)
    
    if ($validation.IsClientIsolated) {
        $validation.IsCompliant = $true
    } else {
        $validation.Warnings += "Site does not follow MSP naming convention. Expected format: $expectedPrefix[SiteName]"
        
        # Generate suggested URL
        $siteName = $SiteCollection
        if ($SiteType) {
            $siteName = "$SiteCollection$SiteType"
        }
        $validation.SuggestedUrl = "https://tenant.sharepoint.com/sites/$expectedPrefix$siteName"
    }

    return $validation
}

function Test-SPOSiteUrlAvailability {
    <#
    .SYNOPSIS
        Checks if a SharePoint site URL is available.
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
        AlternativeUrls = @()
        ExistingSite = $null
    }

    try {
        # Use PnP PowerShell to check site existence
        $existingSite = Invoke-SPOFactoryCommand -ScriptBlock {
            Get-PnPSite -Identity $SiteUrl -ErrorAction SilentlyContinue
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        if ($existingSite) {
            $result.ExistingSite = @{
                Title = $existingSite.Title
                Owner = $existingSite.Owner
                CreatedDate = $existingSite.Created
                LastModified = $existingSite.LastContentModifiedDate
            }
            $result.IsAvailable = $false
            
            # Generate alternative URLs
            $uri = [System.Uri]$SiteUrl
            $sitePath = $uri.LocalPath -replace '^/sites/', ''
            $baseUrl = "$($uri.Scheme)://$($uri.Host)/sites"
            
            for ($i = 2; $i -le 5; $i++) {
                $alternativeUrl = "$baseUrl/$sitePath$i"
                $result.AlternativeUrls += $alternativeUrl
            }
        } else {
            $result.IsAvailable = $true
        }

        Write-SPOFactoryLog -Message "URL availability check completed - Available: $($result.IsAvailable)" -Level Debug -ClientName $ClientName -Category 'Provisioning'
        
        return $result
    }
    catch {
        Write-SPOFactoryLog -Message "Error checking URL availability: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
        
        # Assume available if we can't check
        $result.IsAvailable = $true
        return $result
    }
}