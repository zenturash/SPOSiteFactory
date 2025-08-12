function New-SPOHubSite {
    <#
    .SYNOPSIS
        Creates SharePoint Online hub sites with comprehensive MSP security and compliance features.

    .DESCRIPTION
        Enterprise-grade hub site creation function designed for MSP environments managing multiple
        SharePoint Online tenants. Provides complete hub site provisioning with security baselines,
        client isolation, audit trails, and comprehensive error handling with rollback capabilities.

    .PARAMETER SiteUrl
        The SharePoint hub site URL to create

    .PARAMETER Title
        The title for the hub site

    .PARAMETER Description
        Description for the hub site (optional)

    .PARAMETER Owner
        Primary owner email address for the hub site

    .PARAMETER ClientName
        Client name for MSP tenant isolation and naming conventions

    .PARAMETER SecurityBaseline
        Security baseline to apply (MSPStandard, MSPSecure, or custom path)

    .PARAMETER Template
        Site template to use (default: SITEPAGEPUBLISHING#0 for communication site)

    .PARAMETER Language
        Language ID for the site (default: 1033 for English)

    .PARAMETER TimeZone
        Time zone ID for the site (default: 13 for Eastern Time)

    .PARAMETER HubSiteDesignId
        Site design to apply to the hub site (optional)

    .PARAMETER AssociatedSites
        Array of existing site URLs to associate with the hub (optional)

    .PARAMETER Administrators
        Additional administrators for the hub site (optional)

    .PARAMETER WaitForCompletion
        Wait for hub site creation to complete before returning (default: true)

    .PARAMETER TimeoutMinutes
        Maximum time to wait for site creation in minutes (default: 30)

    .PARAMETER ApplySecurityBaseline
        Apply security baseline after creation (default: true)

    .PARAMETER EnableAuditing
        Enable comprehensive audit logging (default: true)

    .PARAMETER WhatIf
        Show what would be created without actually creating it

    .PARAMETER Force
        Suppress confirmation prompts

    .EXAMPLE
        New-SPOHubSite -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -Title "Contoso Corp Hub" -Owner "admin@contoso.com" -ClientName "ContosoCorp"

    .EXAMPLE
        New-SPOHubSite -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpSecureHub" -Title "Contoso Secure Hub" -Owner "admin@contoso.com" -ClientName "ContosoCorp" -SecurityBaseline "MSPSecure" -AssociatedSites @("https://contoso.sharepoint.com/sites/ContosoCorpTeam1", "https://contoso.sharepoint.com/sites/ContosoCorpTeam2")

    .EXAMPLE
        New-SPOHubSite -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpDeptHub" -Title "Department Hub" -Owner "dept@contoso.com" -ClientName "ContosoCorp" -HubSiteDesignId "12345678-1234-1234-1234-123456789012" -WhatIf
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
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
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$SecurityBaseline = 'MSPStandard',
        
        [Parameter(Mandatory = $false)]
        [string]$Template = 'SITEPAGEPUBLISHING#0',
        
        [Parameter(Mandatory = $false)]
        [int]$Language = 1033,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeZone = 13,
        
        [Parameter(Mandatory = $false)]
        [string]$HubSiteDesignId,
        
        [Parameter(Mandatory = $false)]
        [string[]]$AssociatedSites,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Administrators,
        
        [Parameter(Mandatory = $false)]
        [switch]$WaitForCompletion = $true,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 30,
        
        [Parameter(Mandatory = $false)]
        [switch]$ApplySecurityBaseline = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$EnableAuditing = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        $operationId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        $startTime = Get-Date
        
        Write-SPOFactoryLog -Message "Starting hub site creation: $SiteUrl" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('HubSiteStart', $operationId) -EnableAuditLog

        $result = @{
            Success = $false
            SiteUrl = $SiteUrl
            HubSiteId = $null
            Site = $null
            CreationTime = $null
            SecurityBaseline = @{
                Applied = $false
                Results = @()
            }
            AssociatedSites = @()
            Errors = @()
            Warnings = @()
            RollbackActions = @()
        }

        # Validate prerequisites
        try {
            Write-SPOFactoryLog -Message "Validating prerequisites and parameters" -Level Debug -ClientName $ClientName -Category 'Provisioning'
            
            # Test connection
            $connectionTest = Test-SPOFactoryConnection -ClientName $ClientName
            if (-not $connectionTest.IsConnected) {
                throw "Not connected to SharePoint Online. Connection required for hub site creation."
            }

            # Validate URL format and availability
            $urlValidation = Test-SPOSiteUrl -SiteUrl $SiteUrl -ClientName $ClientName -SiteType 'HubSite' -CheckAvailability -MSPNamingConvention
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

            # Validate template
            $validTemplates = @('SITEPAGEPUBLISHING#0', 'TEAM#0', 'STS#3')
            if ($Template -notin $validTemplates) {
                Write-SPOFactoryLog -Message "Warning: Template '$Template' may not be suitable for hub sites" -Level Warning -ClientName $ClientName -Category 'Provisioning'
            }

            Write-SPOFactoryLog -Message "Prerequisites validation completed successfully" -Level Info -ClientName $ClientName -Category 'Provisioning'
        }
        catch {
            Write-SPOFactoryLog -Message "Prerequisites validation failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception
            throw
        }
    }

    process {
        if (-not $PSCmdlet.ShouldProcess($SiteUrl, "Create SharePoint Hub Site")) {
            return
        }

        try {
            # Step 1: Create the site collection
            Write-SPOFactoryLog -Message "Creating site collection" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('SiteCreation', $operationId)
            
            $siteCreationResult = Invoke-SPOFactoryCommand -ScriptBlock {
                $siteParams = @{
                    Url = $SiteUrl
                    Title = $Title
                    Owner = $Owner
                    Template = $Template
                    Lcid = $Language
                    TimeZone = $TimeZone
                }

                if ($Description) {
                    $siteParams.Description = $Description
                }

                # Create the site
                $site = New-PnPSite @siteParams
                return $site
            } -ClientName $ClientName -Category 'Provisioning' -ErrorMessage "Failed to create site collection" -CriticalOperation

            if (-not $siteCreationResult) {
                throw "Site collection creation returned null result"
            }

            $result.Site = $siteCreationResult
            $result.RollbackActions += @{ Action = 'DeleteSite'; Url = $SiteUrl }

            Write-SPOFactoryLog -Message "Site collection created successfully" -Level Info -ClientName $ClientName -Category 'Provisioning' -EnableAuditLog

            # Step 2: Wait for site to be fully provisioned
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

            # Step 3: Register as hub site
            Write-SPOFactoryLog -Message "Registering site as hub site" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('HubRegistration', $operationId)
            
            $hubRegistrationResult = Invoke-SPOFactoryCommand -ScriptBlock {
                # Connect to specific site for hub registration
                Connect-PnPOnline -Url $SiteUrl -Interactive
                
                # Register as hub site
                $hubSite = Register-PnPHubSite -Site $SiteUrl
                return $hubSite
            } -ClientName $ClientName -Category 'Provisioning' -ErrorMessage "Failed to register hub site" -CriticalOperation

            if ($hubRegistrationResult) {
                $result.HubSiteId = $hubRegistrationResult.HubSiteId
                $result.RollbackActions += @{ Action = 'UnregisterHubSite'; HubSiteId = $hubRegistrationResult.HubSiteId }
                Write-SPOFactoryLog -Message "Hub site registered successfully with ID: $($hubRegistrationResult.HubSiteId)" -Level Info -ClientName $ClientName -Category 'Provisioning' -EnableAuditLog
            }

            # Step 4: Apply additional administrators
            if ($Administrators -and $Administrators.Count -gt 0) {
                Write-SPOFactoryLog -Message "Adding additional administrators: $($Administrators -join ', ')" -Level Info -ClientName $ClientName -Category 'Provisioning'
                
                foreach ($admin in $Administrators) {
                    try {
                        $addAdminResult = Invoke-SPOFactoryCommand -ScriptBlock {
                            Add-PnPSiteCollectionAdmin -Owners $admin
                            return "Added administrator: $admin"
                        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

                        if ($addAdminResult) {
                            Write-SPOFactoryLog -Message $addAdminResult -Level Debug -ClientName $ClientName -Category 'Provisioning'
                        } else {
                            $result.Warnings += "Failed to add administrator: $admin"
                        }
                    }
                    catch {
                        $result.Warnings += "Error adding administrator $admin`: $($_.Exception.Message)"
                    }
                }
            }

            # Step 5: Apply site design if specified
            if ($HubSiteDesignId) {
                Write-SPOFactoryLog -Message "Applying site design: $HubSiteDesignId" -Level Info -ClientName $ClientName -Category 'Provisioning'
                
                $designResult = Invoke-SPOFactoryCommand -ScriptBlock {
                    Invoke-PnPSiteDesign -Identity $HubSiteDesignId
                    return "Site design applied: $HubSiteDesignId"
                } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

                if ($designResult) {
                    Write-SPOFactoryLog -Message $designResult -Level Info -ClientName $ClientName -Category 'Provisioning'
                } else {
                    $result.Warnings += "Failed to apply site design: $HubSiteDesignId"
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

            # Step 7: Associate existing sites if specified
            if ($AssociatedSites -and $AssociatedSites.Count -gt 0) {
                Write-SPOFactoryLog -Message "Associating sites with hub: $($AssociatedSites.Count) sites" -Level Info -ClientName $ClientName -Category 'Hub' -Tag @('SiteAssociation', $operationId)
                
                foreach ($associateSiteUrl in $AssociatedSites) {
                    try {
                        $associationResult = Add-SPOSiteToHub -HubSiteUrl $SiteUrl -SiteUrl $associateSiteUrl -ClientName $ClientName
                        
                        if ($associationResult.Success) {
                            $result.AssociatedSites += $associateSiteUrl
                            Write-SPOFactoryLog -Message "Site associated successfully: $associateSiteUrl" -Level Info -ClientName $ClientName -Category 'Hub'
                        } else {
                            $result.Warnings += "Failed to associate site: $associateSiteUrl"
                        }
                    }
                    catch {
                        $result.Warnings += "Error associating site $associateSiteUrl`: $($_.Exception.Message)"
                    }
                }
            }

            # Mark as successful
            $result.Success = $true
            $result.CreationTime = (Get-Date) - $startTime
            
            Write-SPOFactoryLog -Message "Hub site creation completed successfully in $($result.CreationTime.TotalMinutes.ToString('F1')) minutes" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('HubSiteComplete', $operationId) -EnableAuditLog

            return $result
        }
        catch {
            Write-SPOFactoryLog -Message "Hub site creation failed: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception -Tag @('HubSiteError', $operationId)
            
            $result.Errors += $_.Exception.Message

            # Attempt rollback if requested and not in WhatIf mode
            if (-not $WhatIfPreference -and $result.RollbackActions.Count -gt 0) {
                Write-SPOFactoryLog -Message "Attempting rollback of partial hub site creation" -Level Warning -ClientName $ClientName -Category 'Provisioning'
                
                try {
                    $rollbackResult = Invoke-SPOHubSiteRollback -RollbackActions $result.RollbackActions -ClientName $ClientName
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
            Write-SPOFactoryLog -Message "Hub site creation operation completed successfully" -Level Info -ClientName $ClientName -Category 'Provisioning' -EnableAuditLog
        }
        
        return $result
    }
}

function Test-SPOFactoryConnection {
    <#
    .SYNOPSIS
        Tests SharePoint Online connection for the client.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    try {
        $testResult = Invoke-SPOFactoryCommand -ScriptBlock {
            $context = Get-PnPContext
            return @{
                IsConnected = $context -ne $null
                Url = $context.Url
                User = $context.ExecutingWebRequest.Credentials
            }
        } -ClientName $ClientName -Category 'Connection' -SuppressErrors

        return @{
            IsConnected = $testResult.IsConnected
            Details = $testResult
        }
    }
    catch {
        return @{
            IsConnected = $false
            Error = $_.Exception.Message
        }
    }
}

function Invoke-SPOHubSiteRollback {
    <#
    .SYNOPSIS
        Performs rollback operations for failed hub site creation.
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
            'UnregisterHubSite' { 1 }
            'DeleteSite' { 2 }
            default { 3 }
        }
    }

    foreach ($action in $sortedActions) {
        try {
            switch ($action.Action) {
                'UnregisterHubSite' {
                    Write-SPOFactoryLog -Message "Rolling back hub site registration: $($action.HubSiteId)" -Level Info -ClientName $ClientName -Category 'Provisioning'
                    
                    $unregisterResult = Invoke-SPOFactoryCommand -ScriptBlock {
                        Unregister-PnPHubSite -Site $action.HubSiteId
                        return "Hub site unregistered: $($action.HubSiteId)"
                    } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

                    if ($unregisterResult) {
                        $rollbackResult.RollbackActions += @{ Action = 'UnregisterHubSite'; Status = 'Success'; Details = $unregisterResult }
                    } else {
                        $rollbackResult.Errors += "Failed to unregister hub site: $($action.HubSiteId)"
                    }
                }
                
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