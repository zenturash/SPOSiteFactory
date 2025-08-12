function New-SPOSite {
    <#
    .SYNOPSIS
        Creates SharePoint Online sites with comprehensive MSP features and M365 Group integration.

    .DESCRIPTION
        Enterprise-grade site creation function designed for MSP environments managing multiple
        SharePoint Online tenants. Supports both TeamSite and CommunicationSite creation with
        Microsoft 365 Group integration, security baselines, Office file handling configuration,
        and comprehensive error handling with rollback capabilities.

    .PARAMETER SiteUrl
        The SharePoint site URL to create

    .PARAMETER Title
        The title for the site

    .PARAMETER Description
        Description for the site (optional)

    .PARAMETER Owner
        Primary owner email address for the site

    .PARAMETER SiteType
        Type of site to create (TeamSite or CommunicationSite)

    .PARAMETER ClientName
        Client name for MSP tenant isolation and naming conventions

    .PARAMETER SecurityBaseline
        Security baseline to apply (MSPStandard, MSPSecure, or custom path)

    .PARAMETER Language
        Language ID for the site (default: 1033 for English)

    .PARAMETER TimeZone
        Time zone ID for the site (default: 13 for Eastern Time)

    .PARAMETER Template
        Site template to use (auto-selected based on SiteType if not specified)

    .PARAMETER HubSiteUrl
        Hub site URL to associate this site with (optional)

    .PARAMETER CreateM365Group
        Create Microsoft 365 Group for team sites (default: true for TeamSite)

    .PARAMETER GroupAlias
        Alias for the Microsoft 365 Group (auto-generated if not specified)

    .PARAMETER GroupMembers
        Array of email addresses for group members (optional)

    .PARAMETER GroupOwners
        Array of email addresses for additional group owners (optional)

    .PARAMETER SiteDesignId
        Site design to apply to the site (optional)

    .PARAMETER WaitForCompletion
        Wait for site creation to complete before returning (default: true)

    .PARAMETER TimeoutMinutes
        Maximum time to wait for site creation in minutes (default: 30)

    .PARAMETER ApplySecurityBaseline
        Apply security baseline after creation (default: true)

    .PARAMETER ConfigureOfficeFileHandling
        Configure Office file handling settings for security (default: true)

    .PARAMETER EnableAuditing
        Enable comprehensive audit logging (default: true)

    .PARAMETER WhatIf
        Show what would be created without actually creating it

    .PARAMETER Force
        Suppress confirmation prompts

    .EXAMPLE
        New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpTeam" -Title "Contoso Team Site" -Owner "admin@contoso.com" -SiteType "TeamSite" -ClientName "ContosoCorp"

    .EXAMPLE
        New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpComm" -Title "Contoso Communication Site" -Owner "admin@contoso.com" -SiteType "CommunicationSite" -ClientName "ContosoCorp" -HubSiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub"

    .EXAMPLE
        New-SPOSite -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpProject" -Title "Project Site" -Owner "pm@contoso.com" -SiteType "TeamSite" -ClientName "ContosoCorp" -GroupMembers @("user1@contoso.com", "user2@contoso.com") -SecurityBaseline "MSPSecure"
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $true)]
        [string]$Owner,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('TeamSite', 'CommunicationSite')]
        [string]$SiteType,
        
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$SecurityBaseline = 'MSPStandard',
        
        [Parameter(Mandatory = $false)]
        [int]$Language = 1033,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeZone = 13,
        
        [Parameter(Mandatory = $false)]
        [string]$Template,
        
        [Parameter(Mandatory = $false)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [switch]$CreateM365Group,
        
        [Parameter(Mandatory = $false)]
        [string]$GroupAlias,
        
        [Parameter(Mandatory = $false)]
        [string[]]$GroupMembers,
        
        [Parameter(Mandatory = $false)]
        [string[]]$GroupOwners,
        
        [Parameter(Mandatory = $false)]
        [string]$SiteDesignId,
        
        [Parameter(Mandatory = $false)]
        [switch]$WaitForCompletion = $true,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 30,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplySecurityBaseline = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$ConfigureOfficeFileHandling = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnableAuditing = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        $operationId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        $startTime = Get-Date
        
        Write-SPOFactoryLog -Message "Starting site creation: $SiteUrl ($SiteType)" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('SiteStart', $operationId) -EnableAuditLog

        $result = @{
            Success = $false
            SiteUrl = $SiteUrl
            SiteType = $SiteType
            Site = $null
            M365Group = $null
            CreationTime = $null
            SecurityBaseline = @{
                Applied = $false
                Results = @()
            }
            HubAssociation = $null
            Errors = @()
            Warnings = @()
            RollbackActions = @()
        }

        # Set default values based on site type
        if (-not $Template) {
            $Template = switch ($SiteType) {
                'TeamSite' { 'GROUP#0' }
                'CommunicationSite' { 'SITEPAGEPUBLISHING#0' }
            }
        }

        # Default M365 Group creation for team sites
        if (-not $PSBoundParameters.ContainsKey('CreateM365Group')) {
            $CreateM365Group = ($SiteType -eq 'TeamSite')
        }

        # Generate group alias if not provided
        if ($CreateM365Group -and -not $GroupAlias) {
            $uri = [System.Uri]$SiteUrl
            $sitePath = $uri.LocalPath -replace '^/sites/', ''
            $GroupAlias = $sitePath -replace '[^a-zA-Z0-9]', ''
            
            if ($GroupAlias.Length -gt 64) {
                $GroupAlias = $GroupAlias.Substring(0, 64)
            }
        }

        # Validate prerequisites
        try {
            Write-SPOFactoryLog -Message "Validating prerequisites and parameters" -Level Debug -ClientName $ClientName -Category 'Provisioning'
            
            # Test connection
            $connectionTest = Test-SPOFactoryConnection -ClientName $ClientName
            if (-not $connectionTest.IsConnected) {
                throw "Not connected to SharePoint Online. Connection required for site creation."
            }

            # Validate URL format and availability
            $urlValidation = Test-SPOSiteUrl -SiteUrl $SiteUrl -ClientName $ClientName -SiteType $SiteType -CheckAvailability -MSPNamingConvention
            if (-not $urlValidation.IsValid) {
                throw "URL validation failed: $($urlValidation.ValidationErrors -join '; ')"
            }

            if (-not $urlValidation.IsAvailable) {
                if ($Force) {
                    Write-SPOFactoryLog -Message "Site URL already exists but Force specified, continuing" -Level Warning -ClientName $ClientName -Category 'Provisioning'
                } else {
                    throw "Site URL is already in use: $SiteUrl"
                }
            }

            # Validate owner email
            if ($Owner -notmatch '^[^@]+@[^@]+\.[^@]+$') {
                throw "Invalid owner email address format: $Owner"
            }

            # Validate hub site if specified
            if ($HubSiteUrl) {
                $hubValidation = Test-SPOHubSiteAvailability -HubSiteUrl $HubSiteUrl -ClientName $ClientName
                if (-not $hubValidation.IsAvailable) {
                    throw "Hub site is not available or not registered as hub: $HubSiteUrl"
                }
            }

            Write-SPOFactoryLog -Message "Prerequisites validation completed successfully" -Level Info -ClientName $ClientName -Category 'Provisioning'
        }
        catch {
            Write-SPOFactoryLog -Message "Prerequisites validation failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
            throw
        }
    }

    process {
        if (-not $PSCmdlet.ShouldProcess($SiteUrl, "Create SharePoint Site ($SiteType)")) {
            return
        }

        try {
            # Step 1: Create Microsoft 365 Group if required
            if ($CreateM365Group) {
                Write-SPOFactoryLog -Message "Creating Microsoft 365 Group: $GroupAlias" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('GroupCreation', $operationId)
                
                $groupCreationResult = New-SPOFactoryM365Group -GroupAlias $GroupAlias -DisplayName $Title -Description $Description -Owner $Owner -Members $GroupMembers -AdditionalOwners $GroupOwners -ClientName $ClientName
                
                if ($groupCreationResult.Success) {
                    $result.M365Group = $groupCreationResult.Group
                    $result.RollbackActions += @{ Action = 'DeleteM365Group'; GroupId = $groupCreationResult.Group.Id }
                    Write-SPOFactoryLog -Message "Microsoft 365 Group created successfully: $($groupCreationResult.Group.Id)" -Level Info -ClientName $ClientName -Category 'Provisioning' -EnableAuditLog
                } else {
                    throw "Failed to create Microsoft 365 Group: $($groupCreationResult.Error)"
                }
            }

            # Step 2: Create the site collection
            Write-SPOFactoryLog -Message "Creating $SiteType site collection" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('SiteCreation', $operationId)
            
            $siteCreationResult = switch ($SiteType) {
                'TeamSite' {
                    New-SPOFactoryTeamSite -SiteUrl $SiteUrl -Title $Title -Description $Description -Owner $Owner -Template $Template -Language $Language -TimeZone $TimeZone -GroupId ($result.M365Group.Id) -ClientName $ClientName
                }
                'CommunicationSite' {
                    New-SPOFactoryCommunicationSite -SiteUrl $SiteUrl -Title $Title -Description $Description -Owner $Owner -Template $Template -Language $Language -TimeZone $TimeZone -ClientName $ClientName
                }
            }

            if (-not $siteCreationResult.Success) {
                throw "Site creation failed: $($siteCreationResult.Error)"
            }

            $result.Site = $siteCreationResult.Site
            $result.RollbackActions += @{ Action = 'DeleteSite'; Url = $SiteUrl }

            Write-SPOFactoryLog -Message "$SiteType created successfully" -Level Info -ClientName $ClientName -Category 'Provisioning' -EnableAuditLog

            # Step 3: Wait for site to be fully provisioned
            if ($WaitForCompletion) {
                Write-SPOFactoryLog -Message "Waiting for site provisioning to complete" -Level Info -ClientName $ClientName -Category 'Provisioning'
                
                $waitResult = Wait-SPOSiteCreation -SiteUrl $SiteUrl -TimeoutMinutes $TimeoutMinutes -ClientName $ClientName -ExpectedTitle $Title -ExpectedOwner $Owner -ShowProgress
                
                if (-not $waitResult.Success) {
                    Write-SPOFactoryLog -Message "Site provisioning did not complete within timeout: $($waitResult.FinalStatus)" -Level Warning -ClientName $ClientName -Category 'Provisioning'
                    
                    if ($waitResult.TimedOut) {
                        $result.Warnings += "Site provisioning timed out but may still complete"
                    } else {
                        throw "Site provisioning failed: $($waitResult.FinalStatus)"
                    }
                }
            }

            # Step 4: Apply site design if specified
            if ($SiteDesignId) {
                Write-SPOFactoryLog -Message "Applying site design: $SiteDesignId" -Level Info -ClientName $ClientName -Category 'Provisioning'
                
                $designResult = Invoke-SPOFactoryCommand -ScriptBlock {
                    Connect-PnPOnline -Url $SiteUrl -Interactive
                    $design = Invoke-PnPSiteDesign -Identity $SiteDesignId
                    return "Site design applied: $SiteDesignId"
                } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

                if ($designResult) {
                    Write-SPOFactoryLog -Message $designResult -Level Info -ClientName $ClientName -Category 'Provisioning'
                } else {
                    $result.Warnings += "Failed to apply site design: $SiteDesignId"
                }
            }

            # Step 5: Configure Office file handling for security
            if ($ConfigureOfficeFileHandling) {
                Write-SPOFactoryLog -Message "Configuring Office file handling settings" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('OfficeFileHandling', $operationId)
                
                $officeConfigResult = Set-SPOOfficeFileHandling -SiteUrl $SiteUrl -ClientName $ClientName
                if (-not $officeConfigResult.Success) {
                    $result.Warnings += "Failed to configure Office file handling: $($officeConfigResult.Error)"
                }
            }

            # Step 6: Apply security baseline
            if ($ApplySecurityBaseline) {
                Write-SPOFactoryLog -Message "Applying security baseline: $SecurityBaseline" -Level Info -ClientName $ClientName -Category 'Security' -Tag @('SecurityBaseline', $operationId)
                
                try {
                    $securityResult = Set-SPOSiteSecurityBaseline -SiteUrl $SiteUrl -BaselineName $SecurityBaseline -ClientName $ClientName -ApplyToSite -ConfigureDocumentLibraries -EnableAuditing:$EnableAuditing -Force:$Force
                    
                    $result.SecurityBaseline.Applied = $securityResult.Success
                    $result.SecurityBaseline.Results = $securityResult
                    
                    if ($securityResult.Success) {
                        Write-SPOFactoryLog -Message "Security baseline applied successfully" -Level Info -ClientName $ClientName -Category 'Security' -EnableAuditLog
                    } else {
                        $result.Warnings += "Security baseline application had issues: $($securityResult.FailedSettings.Count) failed settings"
                    }
                }
                catch {
                    $result.Warnings += "Failed to apply security baseline: $($_.Exception.Message)"
                    Write-SPOFactoryLog -Message "Security baseline application failed: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Security' -Exception $_.Exception
                }
            }

            # Step 7: Associate with hub site if specified
            if ($HubSiteUrl) {
                Write-SPOFactoryLog -Message "Associating site with hub: $HubSiteUrl" -Level Info -ClientName $ClientName -Category 'Hub' -Tag @('HubAssociation', $operationId)
                
                try {
                    $hubAssociationResult = Add-SPOSiteToHub -HubSiteUrl $HubSiteUrl -SiteUrl $SiteUrl -ClientName $ClientName
                    
                    if ($hubAssociationResult.Success) {
                        $result.HubAssociation = $hubAssociationResult
                        Write-SPOFactoryLog -Message "Site associated with hub successfully" -Level Info -ClientName $ClientName -Category 'Hub' -EnableAuditLog
                    } else {
                        $result.Warnings += "Failed to associate with hub site: $($hubAssociationResult.Error)"
                    }
                }
                catch {
                    $result.Warnings += "Error associating with hub site: $($_.Exception.Message)"
                }
            }

            # Mark as successful
            $result.Success = $true
            $result.CreationTime = (Get-Date) - $startTime
            
            Write-SPOFactoryLog -Message "Site creation completed successfully in $($result.CreationTime.TotalMinutes.ToString('F1')) minutes" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('SiteComplete', $operationId) -EnableAuditLog

            return $result
        }
        catch {
            Write-SPOFactoryLog -Message "Site creation failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception -Tag @('SiteError', $operationId)
            
            $result.Errors += $_.Exception.Message

            # Attempt rollback if requested and not in WhatIf mode
            if (-not $WhatIfPreference -and $result.RollbackActions.Count -gt 0) {
                Write-SPOFactoryLog -Message "Attempting rollback of partial site creation" -Level Warning -ClientName $ClientName -Category 'Provisioning'
                
                try {
                    $rollbackResult = Invoke-SPOSiteRollback -RollbackActions $result.RollbackActions -ClientName $ClientName
                    Write-SPOFactoryLog -Message "Rollback completed: $($rollbackResult.RollbackActions.Count) actions performed" -Level Info -ClientName $ClientName -Category 'Provisioning'
                }
                catch {
                    Write-SPOFactoryLog -Message "Rollback failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
                    $result.Errors += "Rollback failed: $($_.Exception.Message)"
                }
            }

            throw
        }
    }

    end {
        if ($result.Success) {
            Write-SPOFactoryLog -Message "Site creation operation completed successfully" -Level Info -ClientName $ClientName -Category 'Provisioning' -EnableAuditLog
        }
        
        return $result
    }
}

function New-SPOFactoryM365Group {
    <#
    .SYNOPSIS
        Creates Microsoft 365 Group for SharePoint sites.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$GroupAlias,
        
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $true)]
        [string]$Owner,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Members,
        
        [Parameter(Mandatory = $false)]
        [string[]]$AdditionalOwners,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        Success = $false
        Group = $null
        Error = $null
    }

    try {
        # Create the Microsoft 365 Group
        $groupResult = Invoke-SPOFactoryCommand -ScriptBlock {
            $groupParams = @{
                DisplayName = $DisplayName
                Alias = $GroupAlias
                Owner = $Owner
                IsPrivate = $true  # Private by default for MSP security
            }

            if ($Description) {
                $groupParams.Description = $Description
            }

            $group = New-PnPMicrosoft365Group @groupParams
            return $group
        } -ClientName $ClientName -Category 'Provisioning' -ErrorMessage "Failed to create Microsoft 365 Group"

        if ($groupResult) {
            $result.Group = $groupResult
            $result.Success = $true

            # Add additional owners if specified
            if ($AdditionalOwners -and $AdditionalOwners.Count -gt 0) {
                foreach ($additionalOwner in $AdditionalOwners) {
                    try {
                        Invoke-SPOFactoryCommand -ScriptBlock {
                            Add-PnPMicrosoft365GroupOwner -Identity $groupResult.Id -Users $additionalOwner
                        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Failed to add additional owner $additionalOwner`: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning'
                    }
                }
            }

            # Add members if specified
            if ($Members -and $Members.Count -gt 0) {
                foreach ($member in $Members) {
                    try {
                        Invoke-SPOFactoryCommand -ScriptBlock {
                            Add-PnPMicrosoft365GroupMember -Identity $groupResult.Id -Users $member
                        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Failed to add member $member`: $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning'
                    }
                }
            }
        }

        return $result
    }
    catch {
        $result.Error = $_.Exception.Message
        return $result
    }
}

function New-SPOFactoryTeamSite {
    <#
    .SYNOPSIS
        Creates SharePoint team site with M365 Group integration.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $true)]
        [string]$Owner,
        
        [Parameter(Mandatory = $false)]
        [string]$Template = 'GROUP#0',
        
        [Parameter(Mandatory = $false)]
        [int]$Language = 1033,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeZone = 13,
        
        [Parameter(Mandatory = $false)]
        [string]$GroupId,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        Success = $false
        Site = $null
        Error = $null
    }

    try {
        $siteResult = Invoke-SPOFactoryCommand -ScriptBlock {
            $siteParams = @{
                Type = 'TeamSite'
                Url = $SiteUrl
                Title = $Title
                Owner = $Owner
                Lcid = $Language
                TimeZone = $TimeZone
            }

            if ($Description) {
                $siteParams.Description = $Description
            }

            if ($GroupId) {
                $siteParams.GroupId = $GroupId
            }

            $site = New-PnPSite @siteParams
            return $site
        } -ClientName $ClientName -Category 'Provisioning' -ErrorMessage "Failed to create team site"

        if ($siteResult) {
            $result.Site = $siteResult
            $result.Success = $true
        }

        return $result
    }
    catch {
        $result.Error = $_.Exception.Message
        return $result
    }
}

function New-SPOFactoryCommunicationSite {
    <#
    .SYNOPSIS
        Creates SharePoint communication site.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $true)]
        [string]$Title,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $true)]
        [string]$Owner,
        
        [Parameter(Mandatory = $false)]
        [string]$Template = 'SITEPAGEPUBLISHING#0',
        
        [Parameter(Mandatory = $false)]
        [int]$Language = 1033,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeZone = 13,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        Success = $false
        Site = $null
        Error = $null
    }

    try {
        $siteResult = Invoke-SPOFactoryCommand -ScriptBlock {
            $siteParams = @{
                Type = 'CommunicationSite'
                Url = $SiteUrl
                Title = $Title
                Owner = $Owner
                Lcid = $Language
                TimeZone = $TimeZone
            }

            if ($Description) {
                $siteParams.Description = $Description
            }

            $site = New-PnPSite @siteParams
            return $site
        } -ClientName $ClientName -Category 'Provisioning' -ErrorMessage "Failed to create communication site"

        if ($siteResult) {
            $result.Site = $siteResult
            $result.Success = $true
        }

        return $result
    }
    catch {
        $result.Error = $_.Exception.Message
        return $result
    }
}

function Test-SPOHubSiteAvailability {
    <#
    .SYNOPSIS
        Tests if a hub site is available and registered.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$HubSiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $hubTest = Invoke-SPOFactoryCommand -ScriptBlock {
            $hubSites = Get-PnPHubSite -Identity $HubSiteUrl -ErrorAction SilentlyContinue
            return $hubSites -ne $null
        } -ClientName $ClientName -Category 'Hub' -SuppressErrors

        return @{
            IsAvailable = $hubTest
            HubSiteUrl = $HubSiteUrl
        }
    }
    catch {
        return @{
            IsAvailable = $false
            HubSiteUrl = $HubSiteUrl
            Error = $_.Exception.Message
        }
    }
}

function Set-SPOOfficeFileHandling {
    <#
    .SYNOPSIS
        Configures Office file handling settings for security.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        Success = $false
        AppliedSettings = @()
        Error = $null
    }

    try {
        # Enable Office file handling feature and set DefaultItemOpenInBrowser to false
        $officeResult = Invoke-SPOFactoryCommand -ScriptBlock {
            # Enable the feature: 8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295 (Office Web Apps)
            Enable-PnPFeature -Identity "8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295" -Scope Web -Force -ErrorAction SilentlyContinue

            # Configure document libraries
            $lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 }
            $configuredLibraries = @()
            
            foreach ($list in $lists) {
                Set-PnPList -Identity $list.Id -DefaultItemOpenInBrowser $false
                $configuredLibraries += $list.Title
            }

            return @{
                FeatureEnabled = $true
                ConfiguredLibraries = $configuredLibraries
            }
        } -ClientName $ClientName -Category 'Security' -SuppressErrors

        if ($officeResult) {
            $result.Success = $true
            $result.AppliedSettings += "Office Web Apps feature enabled"
            $result.AppliedSettings += "DefaultItemOpenInBrowser set to false for $($officeResult.ConfiguredLibraries.Count) libraries"
        }

        return $result
    }
    catch {
        $result.Error = $_.Exception.Message
        return $result
    }
}

function Invoke-SPOSiteRollback {
    <#
    .SYNOPSIS
        Performs rollback operations for failed site creation.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$RollbackActions,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $rollbackResult = @{
        Success = $true
        RollbackActions = @()
        Errors = @()
    }

    # Process rollback actions in reverse order
    $sortedActions = $RollbackActions | Sort-Object { 
        switch ($_.Action) {
            'DeleteSite' { 1 }
            'DeleteM365Group' { 2 }
            default { 3 }
        }
    }

    foreach ($action in $sortedActions) {
        try {
            switch ($action.Action) {
                'DeleteSite' {
                    Write-SPOFactoryLog -Message "Rolling back site creation: $($action.Url)" -Level Info -ClientName $ClientName -Category 'Provisioning'
                    
                    $deleteResult = Invoke-SPOFactoryCommand -ScriptBlock {
                        Remove-PnPSite -Identity $action.Url -Force
                        return "Site deleted: $($action.Url)"
                    } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

                    if ($deleteResult) {
                        $rollbackResult.RollbackActions += @{ Action = 'DeleteSite'; Status = 'Success'; Details = $deleteResult }
                    } else {
                        $rollbackResult.Errors += "Failed to delete site: $($action.Url)"
                    }
                }
                
                'DeleteM365Group' {
                    Write-SPOFactoryLog -Message "Rolling back M365 Group creation: $($action.GroupId)" -Level Info -ClientName $ClientName -Category 'Provisioning'
                    
                    $deleteGroupResult = Invoke-SPOFactoryCommand -ScriptBlock {
                        Remove-PnPMicrosoft365Group -Identity $action.GroupId -Force
                        return "M365 Group deleted: $($action.GroupId)"
                    } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

                    if ($deleteGroupResult) {
                        $rollbackResult.RollbackActions += @{ Action = 'DeleteM365Group'; Status = 'Success'; Details = $deleteGroupResult }
                    } else {
                        $rollbackResult.Errors += "Failed to delete M365 Group: $($action.GroupId)"
                    }
                }
            }
        }
        catch {
            $rollbackResult.Errors += "Rollback error for $($action.Action): $($_.Exception.Message)"
            Write-SPOFactoryLog -Message "Rollback error for $($action.Action): $($_.Exception.Message)" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
        }
    }

    $rollbackResult.Success = $rollbackResult.Errors.Count -eq 0
    return $rollbackResult
}