function Set-SPOSiteSecurityBaseline {
    <#
    .SYNOPSIS
        Applies comprehensive security baselines to SharePoint Online sites for MSP environments.

    .DESCRIPTION
        Enterprise security baseline function designed for MSP environments managing multiple
        SharePoint Online tenants. Loads security configurations from baseline files and applies
        tenant-level and site-level security settings, including document library configuration,
        external sharing restrictions, and audit settings.

    .PARAMETER SiteUrl
        The SharePoint site URL to apply security baseline to

    .PARAMETER BaselineName
        Name of the security baseline to apply (MSPStandard, MSPSecure)

    .PARAMETER ClientName
        Client name for MSP tenant isolation and logging

    .PARAMETER CustomBaselinePath
        Path to custom baseline JSON file (optional)

    .PARAMETER ApplyToTenant
        Also apply tenant-level settings from baseline

    .PARAMETER ApplyToSite
        Apply site-level settings from baseline (default: true)

    .PARAMETER ConfigureDocumentLibraries
        Configure document libraries with security settings (default: true)

    .PARAMETER EnableAuditing
        Enable audit logging for the site (default: true)

    .PARAMETER WhatIf
        Show what changes would be made without applying them

    .PARAMETER Force
        Apply settings without confirmation prompts

    .EXAMPLE
        Set-SPOSiteSecurityBaseline -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpTeam" -BaselineName "MSPStandard" -ClientName "ContosoCorp"

    .EXAMPLE
        Set-SPOSiteSecurityBaseline -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -BaselineName "MSPSecure" -ClientName "ContosoCorp" -ApplyToTenant

    .EXAMPLE
        Set-SPOSiteSecurityBaseline -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpComm" -CustomBaselinePath "C:\Baselines\Custom.json" -ClientName "ContosoCorp" -WhatIf
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('MSPStandard', 'MSPSecure')]
        [string]$BaselineName = 'MSPStandard',
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$CustomBaselinePath,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyToTenant,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplyToSite = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$ConfigureDocumentLibraries = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnableAuditing = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        $operationId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        Write-SPOFactoryLog -Message "Starting security baseline application: $BaselineName" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('BaselineStart', $operationId)

        $result = @{
            Success = $false
            BaselineName = $BaselineName
            AppliedSettings = @()
            FailedSettings = @()
            Warnings = @()
            TenantSettings = @()
            SiteSettings = @()
            DocumentLibrarySettings = @()
            AuditSettings = @()
        }

        # Load baseline configuration
        try {
            if ($CustomBaselinePath -and (Test-Path $CustomBaselinePath)) {
                Write-SPOFactoryLog -Message "Loading custom baseline from: $CustomBaselinePath" -Level Debug -ClientName $ClientName -Category 'Security'
                $baseline = Get-Content $CustomBaselinePath -Raw | ConvertFrom-Json
            } else {
                Write-SPOFactoryLog -Message "Loading standard baseline: $BaselineName" -Level Debug -ClientName $ClientName -Category 'Security'
                $baseline = Get-SPOFactoryBaseline -BaselineName $BaselineName
            }

            if (-not $baseline) {
                throw "Unable to load baseline configuration"
            }

            Write-SPOFactoryLog -Message "Baseline loaded successfully: $($baseline.name) v$($baseline.version)" -Level Info -ClientName $ClientName -Category 'Security'
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to load baseline configuration: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Security' -Exception $_.Exception
            throw
        }
    }

    process {
        try {
            # Validate baseline before applying
            $validationResult = Test-SPOFactoryBaseline -Baseline $baseline -ClientName $ClientName
            if (-not $validationResult.IsValid) {
                Write-SPOFactoryLog -Message "Baseline validation failed: $($validationResult.Issues -join '; ')" -Level Error -ClientName $ClientName -Category 'Security'
                throw "Baseline validation failed"
            }

            # Apply tenant-level settings if requested
            if ($ApplyToTenant -and $baseline.tenantSettings) {
                Write-SPOFactoryLog -Message "Applying tenant-level security settings" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('TenantSettings', $operationId)
                
                if ($PSCmdlet.ShouldProcess("Tenant", "Apply tenant security settings")) {
                    $tenantResult = Set-SPOTenantSecuritySettings -Settings $baseline.tenantSettings -ClientName $ClientName -Force:$Force
                    $result.TenantSettings = $tenantResult.AppliedSettings
                    $result.FailedSettings += $tenantResult.FailedSettings
                }
            }

            # Apply site-level settings
            if ($ApplyToSite -and $baseline.siteSettings) {
                Write-SPOFactoryLog -Message "Applying site-level security settings" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('SiteSettings', $operationId)
                
                if ($PSCmdlet.ShouldProcess($SiteUrl, "Apply site security settings")) {
                    $siteResult = Set-SPOSiteSecuritySettings -SiteUrl $SiteUrl -Settings $baseline.siteSettings -SecuritySettings $baseline.securitySettings -ClientName $ClientName -Force:$Force
                    $result.SiteSettings = $siteResult.AppliedSettings
                    $result.FailedSettings += $siteResult.FailedSettings
                }
            }

            # Configure document libraries
            if ($ConfigureDocumentLibraries) {
                Write-SPOFactoryLog -Message "Configuring document library security settings" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('DocumentLibraries', $operationId)
                
                if ($PSCmdlet.ShouldProcess($SiteUrl, "Configure document library security")) {
                    $libraryResult = Set-SPODocumentLibrarySecuritySettings -SiteUrl $SiteUrl -Baseline $baseline -ClientName $ClientName -Force:$Force
                    $result.DocumentLibrarySettings = $libraryResult.AppliedSettings
                    $result.FailedSettings += $libraryResult.FailedSettings
                }
            }

            # Enable auditing
            if ($EnableAuditing -and $baseline.securitySettings.enableAuditLog) {
                Write-SPOFactoryLog -Message "Configuring audit logging" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('AuditSettings', $operationId)
                
                if ($PSCmdlet.ShouldProcess($SiteUrl, "Enable audit logging")) {
                    $auditResult = Set-SPOSiteAuditSettings -SiteUrl $SiteUrl -Baseline $baseline -ClientName $ClientName -Force:$Force
                    $result.AuditSettings = $auditResult.AppliedSettings
                    $result.FailedSettings += $auditResult.FailedSettings
                }
            }

            # Compile all applied settings
            $result.AppliedSettings = $result.TenantSettings + $result.SiteSettings + $result.DocumentLibrarySettings + $result.AuditSettings

            # Determine overall success
            $result.Success = ($result.FailedSettings.Count -eq 0) -or ($result.AppliedSettings.Count -gt $result.FailedSettings.Count)

            # Log final results
            $successCount = $result.AppliedSettings.Count
            $failedCount = $result.FailedSettings.Count
            $logLevel = if ($result.Success) { 'Info' } else { 'Warning' }
            
            Write-SPOFactoryLog -Message "Security baseline application completed - Success: $successCount, Failed: $failedCount" -Level $logLevel -ClientName $ClientName -Category 'Security' -Tag @('BaselineComplete', $operationId) -EnableAuditLog

            return $result
        }
        catch {
            Write-SPOFactoryLog -Message "Error applying security baseline: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Security' -Exception $_.Exception -Tag @('BaselineError', $operationId)
            throw
        }
    }
}

function Get-SPOFactoryBaseline {
    <#
    .SYNOPSIS
        Loads baseline configuration from module data files.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$BaselineName
    )

    try {
        $moduleRoot = $script:SPOFactoryConfig.ModuleRoot
        $baselinePath = Join-Path $moduleRoot "Data\Baselines\$BaselineName.json"
        
        if (-not (Test-Path $baselinePath)) {
            throw "Baseline file not found: $baselinePath"
        }

        $baseline = Get-Content $baselinePath -Raw | ConvertFrom-Json
        return $baseline
    }
    catch {
        throw "Failed to load baseline '$BaselineName': $($_.Exception.Message)"
    }
}

function Test-SPOFactoryBaseline {
    <#
    .SYNOPSIS
        Validates baseline configuration for consistency and completeness.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$Baseline,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $validation = @{
        IsValid = $true
        Issues = @()
        Warnings = @()
    }

    try {
        # Check required properties
        $requiredProperties = @('name', 'version', 'tenantSettings', 'siteSettings', 'securitySettings')
        foreach ($property in $requiredProperties) {
            if (-not $Baseline.PSObject.Properties.Name -contains $property) {
                $validation.Issues += "Missing required property: $property"
                $validation.IsValid = $false
            }
        }

        # Validate against schema if available
        if ($Baseline.validationRules) {
            foreach ($requiredSetting in $Baseline.validationRules.requiredSettings) {
                if (-not $Baseline.siteSettings.PSObject.Properties.Name -contains $requiredSetting) {
                    $validation.Issues += "Missing required site setting: $requiredSetting"
                    $validation.IsValid = $false
                }
            }

            # Check for conflicting settings
            if ($Baseline.validationRules.conflictSettings) {
                foreach ($conflictKey in $Baseline.validationRules.conflictSettings.PSObject.Properties.Name) {
                    $conflictRule = $Baseline.validationRules.conflictSettings.$conflictKey
                    
                    foreach ($conflictsWith in $conflictRule.conflictsWith) {
                        if ($Baseline.tenantSettings.$conflictKey -and $Baseline.tenantSettings.$conflictsWith) {
                            $validation.Warnings += "Potential conflict: $conflictKey conflicts with $conflictsWith - $($conflictRule.reason)"
                        }
                    }
                }
            }
        }

        return $validation
    }
    catch {
        $validation.IsValid = $false
        $validation.Issues += "Error validating baseline: $($_.Exception.Message)"
        return $validation
    }
}

function Set-SPOTenantSecuritySettings {
    <#
    .SYNOPSIS
        Applies tenant-level security settings.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$Settings,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $result = @{
        AppliedSettings = @()
        FailedSettings = @()
    }

    try {
        # Apply sharing settings
        if ($Settings.sharingCapability) {
            $settingResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPTenant -SharingCapability $Settings.sharingCapability
                return "SharingCapability set to $($Settings.sharingCapability)"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($settingResult) {
                $result.AppliedSettings += @{ Setting = 'SharingCapability'; Value = $Settings.sharingCapability; Result = $settingResult }
            } else {
                $result.FailedSettings += @{ Setting = 'SharingCapability'; Value = $Settings.sharingCapability; Error = 'Failed to apply' }
            }
        }

        # Apply default sharing link type
        if ($Settings.defaultSharingLinkType) {
            $settingResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPTenant -DefaultSharingLinkType $Settings.defaultSharingLinkType
                return "DefaultSharingLinkType set to $($Settings.defaultSharingLinkType)"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($settingResult) {
                $result.AppliedSettings += @{ Setting = 'DefaultSharingLinkType'; Value = $Settings.defaultSharingLinkType; Result = $settingResult }
            } else {
                $result.FailedSettings += @{ Setting = 'DefaultSharingLinkType'; Value = $Settings.defaultSharingLinkType; Error = 'Failed to apply' }
            }
        }

        # Apply external user expiration
        if ($Settings.externalUserExpireInDays) {
            $settingResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPTenant -ExternalUserExpireInDays $Settings.externalUserExpireInDays
                return "ExternalUserExpireInDays set to $($Settings.externalUserExpireInDays)"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($settingResult) {
                $result.AppliedSettings += @{ Setting = 'ExternalUserExpireInDays'; Value = $Settings.externalUserExpireInDays; Result = $settingResult }
            } else {
                $result.FailedSettings += @{ Setting = 'ExternalUserExpireInDays'; Value = $Settings.externalUserExpireInDays; Error = 'Failed to apply' }
            }
        }

        return $result
    }
    catch {
        $result.FailedSettings += @{ Setting = 'TenantSettings'; Error = $_.Exception.Message }
        return $result
    }
}

function Set-SPOSiteSecuritySettings {
    <#
    .SYNOPSIS
        Applies site-level security settings.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [PSObject]$Settings,
        
        [Parameter(Mandatory = $false)]
        [PSObject]$SecuritySettings,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $result = @{
        AppliedSettings = @()
        FailedSettings = @()
    }

    try {
        # Apply site sharing settings
        if ($Settings.sharingCapability) {
            $settingResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPSite -Identity $SiteUrl -SharingCapability $Settings.sharingCapability
                return "Site SharingCapability set to $($Settings.sharingCapability)"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($settingResult) {
                $result.AppliedSettings += @{ Setting = 'SiteSharingCapability'; Value = $Settings.sharingCapability; Result = $settingResult }
            } else {
                $result.FailedSettings += @{ Setting = 'SiteSharingCapability'; Value = $Settings.sharingCapability; Error = 'Failed to apply' }
            }
        }

        # Disable custom pages if specified
        if ($Settings.denyAddAndCustomizePages -eq $true) {
            $settingResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPSite -Identity $SiteUrl -DenyAddAndCustomizePages
                return "DenyAddAndCustomizePages enabled"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($settingResult) {
                $result.AppliedSettings += @{ Setting = 'DenyAddAndCustomizePages'; Value = $true; Result = $settingResult }
            } else {
                $result.FailedSettings += @{ Setting = 'DenyAddAndCustomizePages'; Value = $true; Error = 'Failed to apply' }
            }
        }

        # Apply member sharing restrictions
        if ($Settings.membersCanShare -eq $false) {
            $settingResult = Invoke-SPOFactoryCommand -ScriptBlock {
                $web = Get-PnPWeb
                $web.MembersCanShare = $false
                $web.Update()
                Invoke-PnPQuery
                return "MembersCanShare disabled"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($settingResult) {
                $result.AppliedSettings += @{ Setting = 'MembersCanShare'; Value = $false; Result = $settingResult }
            } else {
                $result.FailedSettings += @{ Setting = 'MembersCanShare'; Value = $false; Error = 'Failed to apply' }
            }
        }

        return $result
    }
    catch {
        $result.FailedSettings += @{ Setting = 'SiteSettings'; Error = $_.Exception.Message }
        return $result
    }
}

function Set-SPODocumentLibrarySecuritySettings {
    <#
    .SYNOPSIS
        Configures document library security settings.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [PSObject]$Baseline,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $result = @{
        AppliedSettings = @()
        FailedSettings = @()
    }

    try {
        # Get all document libraries
        $libraries = Invoke-SPOFactoryCommand -ScriptBlock {
            Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }
        } -ClientName $ClientName -Category 'Security' -SuppressErrors

        foreach ($library in $libraries) {
            Write-SPOFactoryLog -Message "Configuring security for library: $($library.Title)" -Level Debug -ClientName $ClientName -Category 'Security'

            # Enable versioning
            $versionResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPList -Identity $library.Id -EnableVersioning $true -MajorVersions 50 -MinorVersions 10
                return "Versioning enabled for $($library.Title)"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($versionResult) {
                $result.AppliedSettings += @{ Setting = 'Versioning'; Library = $library.Title; Result = $versionResult }
            } else {
                $result.FailedSettings += @{ Setting = 'Versioning'; Library = $library.Title; Error = 'Failed to enable versioning' }
            }

            # Configure content approval if specified
            if ($Baseline.governanceSettings.requireContentApproval) {
                $approvalResult = Invoke-SPOFactoryCommand -ScriptBlock {
                    Set-PnPList -Identity $library.Id -EnableContentTypes $true -EnableModeration $true
                    return "Content approval enabled for $($library.Title)"
                } -ClientName $ClientName -Category 'Security' -SuppressErrors

                if ($approvalResult) {
                    $result.AppliedSettings += @{ Setting = 'ContentApproval'; Library = $library.Title; Result = $approvalResult }
                } else {
                    $result.FailedSettings += @{ Setting = 'ContentApproval'; Library = $library.Title; Error = 'Failed to enable content approval' }
                }
            }

            # Disable DefaultItemOpenInBrowser for security
            $browserResult = Invoke-SPOFactoryCommand -ScriptBlock {
                Set-PnPList -Identity $library.Id -DefaultItemOpenInBrowser $false
                return "DefaultItemOpenInBrowser disabled for $($library.Title)"
            } -ClientName $ClientName -Category 'Security' -SuppressErrors

            if ($browserResult) {
                $result.AppliedSettings += @{ Setting = 'DefaultItemOpenInBrowser'; Library = $library.Title; Result = $browserResult }
            } else {
                $result.FailedSettings += @{ Setting = 'DefaultItemOpenInBrowser'; Library = $library.Title; Error = 'Failed to disable DefaultItemOpenInBrowser' }
            }
        }

        return $result
    }
    catch {
        $result.FailedSettings += @{ Setting = 'DocumentLibrarySettings'; Error = $_.Exception.Message }
        return $result
    }
}

function Set-SPOSiteAuditSettings {
    <#
    .SYNOPSIS
        Configures audit logging settings for the site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [PSObject]$Baseline,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    $result = @{
        AppliedSettings = @()
        FailedSettings = @()
    }

    try {
        # Enable audit logging
        $auditResult = Invoke-SPOFactoryCommand -ScriptBlock {
            $site = Get-PnPSite
            $site.Audit.AuditFlags = [Microsoft.SharePoint.Client.AuditMaskType]::All
            $site.Audit.Update()
            $site.Update()
            Invoke-PnPQuery
            return "Audit logging enabled with all flags"
        } -ClientName $ClientName -Category 'Security' -SuppressErrors

        if ($auditResult) {
            $result.AppliedSettings += @{ Setting = 'AuditLogging'; Result = $auditResult }
        } else {
            $result.FailedSettings += @{ Setting = 'AuditLogging'; Error = 'Failed to enable audit logging' }
        }

        Write-SPOFactoryLog -Message "Audit configuration completed for site" -Level Info -ClientName $ClientName -Category 'Security' -EnableAuditLog

        return $result
    }
    catch {
        $result.FailedSettings += @{ Setting = 'AuditSettings'; Error = $_.Exception.Message }
        return $result
    }
}