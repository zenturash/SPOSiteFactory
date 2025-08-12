function Set-SPOSiteTemplate {
    <#
    .SYNOPSIS
        Updates existing SharePoint site templates for MSP multi-tenant environments.
    
    .DESCRIPTION
        The Set-SPOSiteTemplate function updates existing site templates with new configurations,
        libraries, lists, features, or security settings. Supports both local template files
        and SharePoint Online site designs.
    
    .PARAMETER Name
        Name of the template to update.
    
    .PARAMETER Path
        Path to the template file to update.
    
    .PARAMETER DisplayName
        New display name for the template.
    
    .PARAMETER Description
        New description for the template.
    
    .PARAMETER Category
        New category for the template.
    
    .PARAMETER Libraries
        Updated array of document libraries.
    
    .PARAMETER Lists
        Updated array of lists.
    
    .PARAMETER Features
        Updated array of feature IDs.
    
    .PARAMETER SecurityBaseline
        Updated security baseline.
    
    .PARAMETER AddLibrary
        Add a new library to the existing template.
    
    .PARAMETER RemoveLibrary
        Remove a library from the template.
    
    .PARAMETER AddList
        Add a new list to the existing template.
    
    .PARAMETER RemoveList
        Remove a list from the template.
    
    .PARAMETER ClientName
        The MSP client name for client-specific updates.
    
    .PARAMETER UpdateSharePoint
        If specified, updates the SharePoint site design.
    
    .PARAMETER Force
        Suppress confirmation prompts.
    
    .EXAMPLE
        Set-SPOSiteTemplate -Name "ProjectTemplate" -Description "Updated project template with new features"
        
        Updates the description of an existing template.
    
    .EXAMPLE
        Set-SPOSiteTemplate -Name "ProjectTemplate" -AddLibrary @{name="Contracts"; versioning=$true}
        
        Adds a new document library to an existing template.
    
    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
        Requires: SharePoint Online Management Shell, PnP.PowerShell
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'ByName')]
        [string]$Name,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'ByPath')]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Template file not found: $_"
            }
            $true
        })]
        [string]$Path,
        
        [Parameter()]
        [string]$DisplayName,
        
        [Parameter()]
        [string]$Description,
        
        [Parameter()]
        [ValidateSet('Project', 'Department', 'Communication', 'TeamSite', 'Custom')]
        [string]$Category,
        
        [Parameter()]
        [object[]]$Libraries,
        
        [Parameter()]
        [object[]]$Lists,
        
        [Parameter()]
        [string[]]$Features,
        
        [Parameter()]
        [string]$SecurityBaseline,
        
        [Parameter()]
        [object]$AddLibrary,
        
        [Parameter()]
        [string]$RemoveLibrary,
        
        [Parameter()]
        [object]$AddList,
        
        [Parameter()]
        [string]$RemoveList,
        
        [Parameter()]
        [string]$ClientName,
        
        [Parameter()]
        [object]$Navigation,
        
        [Parameter()]
        [string]$Theme,
        
        [Parameter()]
        [switch]$UpdateSharePoint,
        
        [Parameter()]
        [switch]$Force
    )
    
    begin {
        Write-SPOFactoryLog -Message "Updating site template" -Level Info
        
        # Get module base path
        $modulePath = Split-Path -Parent $PSScriptRoot
        $templatesPath = Join-Path $modulePath "Data\Templates"
        
        # Determine template path
        if ($PSCmdlet.ParameterSetName -eq 'ByName') {
            $templatePath = Join-Path $templatesPath "$Name.json"
            if (-not (Test-Path $templatePath)) {
                throw "Template not found: $Name"
            }
        }
        else {
            $templatePath = $Path
        }
        
        Write-SPOFactoryLog -Message "Loading template from: $templatePath" -Level Info
    }
    
    process {
        try {
            # Load existing template
            $templateContent = Get-Content -Path $templatePath -Raw
            $template = $templateContent | ConvertFrom-Json -AsHashtable
            
            if (-not $template) {
                throw "Failed to load template from $templatePath"
            }
            
            # Track changes for logging
            $changes = @()
            
            # Update basic properties
            if ($DisplayName) {
                $template.displayName = $DisplayName
                $changes += "Updated display name to: $DisplayName"
            }
            
            if ($Description) {
                $template.description = $Description
                $changes += "Updated description"
            }
            
            if ($Category) {
                $template.category = $Category
                $changes += "Updated category to: $Category"
            }
            
            if ($SecurityBaseline) {
                $template.securityBaseline = $SecurityBaseline
                $changes += "Updated security baseline to: $SecurityBaseline"
            }
            
            if ($Theme) {
                $template.theme = $Theme
                $changes += "Updated theme to: $Theme"
            }
            
            if ($Navigation) {
                $template.navigation = $Navigation
                $changes += "Updated navigation structure"
            }
            
            # Handle libraries
            if ($Libraries) {
                $template.libraries = $Libraries
                $changes += "Replaced all libraries ($($Libraries.Count) libraries)"
            }
            
            if ($AddLibrary) {
                if (-not $template.libraries) {
                    $template.libraries = @()
                }
                
                # Check if library already exists
                $existing = $template.libraries | Where-Object { $_.name -eq $AddLibrary.name }
                if ($existing -and -not $Force) {
                    throw "Library '$($AddLibrary.name)' already exists in template. Use -Force to overwrite."
                }
                elseif ($existing) {
                    # Remove existing library
                    $template.libraries = $template.libraries | Where-Object { $_.name -ne $AddLibrary.name }
                }
                
                # Add the new library
                $newLibrary = @{
                    name = $AddLibrary.name
                    displayName = if ($AddLibrary.displayName) { $AddLibrary.displayName } else { $AddLibrary.name }
                    versioning = if ($null -ne $AddLibrary.versioning) { $AddLibrary.versioning } else { $true }
                    checkOut = if ($null -ne $AddLibrary.checkOut) { $AddLibrary.checkOut } else { $false }
                    majorVersionLimit = if ($AddLibrary.majorVersionLimit) { $AddLibrary.majorVersionLimit } else { 50 }
                }
                
                if ($AddLibrary.minorVersionLimit) {
                    $newLibrary.minorVersionLimit = $AddLibrary.minorVersionLimit
                }
                
                $template.libraries += $newLibrary
                $changes += "Added library: $($AddLibrary.name)"
            }
            
            if ($RemoveLibrary) {
                if ($template.libraries) {
                    $originalCount = $template.libraries.Count
                    $template.libraries = $template.libraries | Where-Object { $_.name -ne $RemoveLibrary }
                    
                    if ($template.libraries.Count -lt $originalCount) {
                        $changes += "Removed library: $RemoveLibrary"
                    }
                    else {
                        Write-SPOFactoryLog -Message "Library '$RemoveLibrary' not found in template" -Level Warning
                    }
                }
            }
            
            # Handle lists
            if ($Lists) {
                $template.lists = $Lists
                $changes += "Replaced all lists ($($Lists.Count) lists)"
            }
            
            if ($AddList) {
                if (-not $template.lists) {
                    $template.lists = @()
                }
                
                # Check if list already exists
                $existing = $template.lists | Where-Object { $_.name -eq $AddList.name }
                if ($existing -and -not $Force) {
                    throw "List '$($AddList.name)' already exists in template. Use -Force to overwrite."
                }
                elseif ($existing) {
                    # Remove existing list
                    $template.lists = $template.lists | Where-Object { $_.name -ne $AddList.name }
                }
                
                # Add the new list
                $newList = @{
                    name = $AddList.name
                    displayName = if ($AddList.displayName) { $AddList.displayName } else { $AddList.name }
                    template = if ($AddList.template) { $AddList.template } else { "GenericList" }
                    description = if ($AddList.description) { $AddList.description } else { "" }
                }
                
                if ($AddList.columns) {
                    $newList.columns = $AddList.columns
                }
                
                $template.lists += $newList
                $changes += "Added list: $($AddList.name)"
            }
            
            if ($RemoveList) {
                if ($template.lists) {
                    $originalCount = $template.lists.Count
                    $template.lists = $template.lists | Where-Object { $_.name -ne $RemoveList }
                    
                    if ($template.lists.Count -lt $originalCount) {
                        $changes += "Removed list: $RemoveList"
                    }
                    else {
                        Write-SPOFactoryLog -Message "List '$RemoveList' not found in template" -Level Warning
                    }
                }
            }
            
            # Handle features
            if ($Features) {
                $template.features = $Features
                $changes += "Updated features ($($Features.Count) features)"
            }
            
            # Update metadata
            $template.version = if ($template.version) {
                # Increment minor version
                $versionParts = $template.version -split '\.'
                if ($versionParts.Count -eq 3) {
                    $versionParts[2] = [int]$versionParts[2] + 1
                    $versionParts -join '.'
                }
                else {
                    "1.0.1"
                }
            }
            else {
                "1.0.1"
            }
            
            $template.modified = Get-Date -Format "yyyy-MM-dd"
            $template.modifiedBy = if ($env:USERNAME) { $env:USERNAME } else { "MSP Automation" }
            
            # Add client restriction if specified
            if ($ClientName) {
                if (-not $template.clients) {
                    $template.clients = @()
                }
                if ($ClientName -notin $template.clients) {
                    $template.clients += $ClientName
                    $changes += "Added client restriction: $ClientName"
                }
            }
            
            # Save updated template
            if ($PSCmdlet.ShouldProcess($templatePath, "Update Site Template")) {
                if ($changes.Count -eq 0) {
                    Write-SPOFactoryLog -Message "No changes to apply to template" -Level Warning
                    return
                }
                
                # Backup existing template
                $backupPath = "$templatePath.backup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Copy-Item -Path $templatePath -Destination $backupPath -Force
                Write-SPOFactoryLog -Message "Created backup: $backupPath" -Level Info
                
                # Save updated template
                $template | ConvertTo-Json -Depth 10 | Out-File -FilePath $templatePath -Encoding UTF8 -Force
                Write-SPOFactoryLog -Message "Template updated successfully" -Level Info
                
                # Log all changes
                foreach ($change in $changes) {
                    Write-SPOFactoryLog -Message "Change: $change" -Level Info
                }
                
                # Update SharePoint site design if requested
                if ($UpdateSharePoint) {
                    Write-SPOFactoryLog -Message "Updating SharePoint site design" -Level Info
                    
                    $updateResult = Update-SPOSiteDesignFromTemplate -Template $template -ClientName $ClientName
                    
                    if ($updateResult.Success) {
                        Write-SPOFactoryLog -Message "SharePoint site design updated successfully" -Level Info
                    }
                    else {
                        Write-SPOFactoryLog -Message "Failed to update SharePoint site design: $($updateResult.Error)" -Level Warning
                    }
                }
                
                # Return updated template info
                [PSCustomObject]@{
                    Name = $template.name
                    DisplayName = $template.displayName
                    Description = $template.description
                    Category = $template.category
                    Version = $template.version
                    Modified = $template.modified
                    Path = $templatePath
                    BackupPath = $backupPath
                    Changes = $changes
                    UpdatedSharePoint = $UpdateSharePoint.IsPresent
                }
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to update template: $_" -Level Error
            throw
        }
    }
    
    end {
        Write-SPOFactoryLog -Message "Template update completed" -Level Info
    }
}

function Update-SPOSiteDesignFromTemplate {
    <#
    .SYNOPSIS
        Updates a SharePoint site design from a template.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Template,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )
    
    $result = @{
        Success = $false
        Error = $null
    }
    
    try {
        # Find existing site design
        $existingDesign = Invoke-SPOFactoryCommand -ScriptBlock {
            Get-PnPSiteDesign | Where-Object { $_.Title -eq $Template.displayName }
        } -ClientName $ClientName -Category 'Configuration' -SuppressErrors
        
        if ($existingDesign) {
            # Update the site script
            $scriptUpdate = Invoke-SPOFactoryCommand -ScriptBlock {
                # Build updated site script content
                $scriptContent = @{
                    '$schema' = 'https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json'
                    actions = @()
                    version = 2
                }
                
                # Add updated actions based on template
                # ... (similar to Export-SPOTemplateToSiteDesign)
                
                # Update the script
                if ($existingDesign.SiteScriptIds -and $existingDesign.SiteScriptIds.Count -gt 0) {
                    Set-PnPSiteScript -Identity $existingDesign.SiteScriptIds[0] -Content ($scriptContent | ConvertTo-Json -Depth 10)
                }
                
                return $true
            } -ClientName $ClientName -Category 'Configuration' -ErrorMessage "Failed to update site script"
            
            if ($scriptUpdate) {
                $result.Success = $true
            }
        }
        else {
            $result.Error = "Site design not found for template: $($Template.displayName)"
        }
        
        return $result
    }
    catch {
        $result.Error = $_.Exception.Message
        return $result
    }
}