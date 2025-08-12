function New-SPOSiteTemplate {
    <#
    .SYNOPSIS
        Creates new SharePoint site templates for MSP multi-tenant environments.
    
    .DESCRIPTION
        The New-SPOSiteTemplate function creates custom site templates that define site structure,
        libraries, lists, features, and security settings. Templates enable consistent site
        provisioning across multiple MSP clients with standardized configurations.
    
    .PARAMETER Name
        Name of the template (no spaces, used as identifier).
    
    .PARAMETER DisplayName
        Display name for the template.
    
    .PARAMETER Description
        Description of what the template provides.
    
    .PARAMETER Category
        Category of the template (Project, Department, Communication, TeamSite, Custom).
    
    .PARAMETER BaseTemplate
        SharePoint base template to use (GROUP#0, SITEPAGEPUBLISHING#0, STS#3).
    
    .PARAMETER Libraries
        Array of document libraries to create with the template.
    
    .PARAMETER Lists
        Array of lists to create with the template.
    
    .PARAMETER Features
        Array of feature IDs to activate.
    
    .PARAMETER SecurityBaseline
        Security baseline to apply (MSPStandard, MSPSecure, or custom).
    
    .PARAMETER ClientName
        The MSP client name for client-specific templates.
    
    .PARAMETER Navigation
        Navigation structure to apply.
    
    .PARAMETER Theme
        Theme to apply to sites created from this template.
    
    .PARAMETER SiteScripts
        Array of site script IDs to include.
    
    .PARAMETER OutputPath
        Path where the template file should be saved.
    
    .PARAMETER ExportToSharePoint
        If specified, exports the template as a SharePoint site design.
    
    .PARAMETER Force
        Overwrite existing template with the same name.
    
    .EXAMPLE
        New-SPOSiteTemplate -Name "ProjectTemplate" -DisplayName "Project Site Template" -Category "Project" -BaseTemplate "GROUP#0" -SecurityBaseline "MSPSecure"
        
        Creates a new project site template with secure baseline.
    
    .EXAMPLE
        $libraries = @(
            @{name="Project Docs"; versioning=$true; checkOut=$true},
            @{name="Deliverables"; versioning=$true}
        )
        New-SPOSiteTemplate -Name "CustomProject" -DisplayName "Custom Project Template" -Libraries $libraries -ClientName "Contoso"
        
        Creates a custom project template with specific document libraries.
    
    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
        Requires: SharePoint Online Management Shell, PnP.PowerShell
    #>
    
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^[a-zA-Z0-9]+$')]
        [string]$Name,
        
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('Project', 'Department', 'Communication', 'TeamSite', 'Custom')]
        [string]$Category,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('GROUP#0', 'SITEPAGEPUBLISHING#0', 'STS#3')]
        [string]$BaseTemplate = 'GROUP#0',
        
        [Parameter(Mandatory = $false)]
        [object[]]$Libraries,
        
        [Parameter(Mandatory = $false)]
        [object[]]$Lists,
        
        [Parameter(Mandatory = $false)]
        [string[]]$Features,
        
        [Parameter(Mandatory = $false)]
        [string]$SecurityBaseline = 'MSPStandard',
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [object]$Navigation,
        
        [Parameter(Mandatory = $false)]
        [string]$Theme,
        
        [Parameter(Mandatory = $false)]
        [string[]]$SiteScripts,
        
        [Parameter(Mandatory = $false)]
        [string]$OutputPath,
        
        [Parameter(Mandatory = $false)]
        [switch]$ExportToSharePoint,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )
    
    begin {
        Write-SPOFactoryLog -Message "Creating new site template: $Name" -Level Info
        
        # Get module base path
        $modulePath = Split-Path -Parent $PSScriptRoot
        $templatesPath = Join-Path $modulePath "Data\Templates"
        
        # Create templates directory if it doesn't exist
        if (-not (Test-Path $templatesPath)) {
            New-Item -Path $templatesPath -ItemType Directory -Force | Out-Null
        }
        
        # Set default output path if not specified
        if (-not $OutputPath) {
            $OutputPath = Join-Path $templatesPath "$Name.json"
        }
        
        # Check if template already exists
        if ((Test-Path $OutputPath) -and -not $Force) {
            throw "Template already exists at $OutputPath. Use -Force to overwrite."
        }
    }
    
    process {
        try {
            # Build template structure
            $template = @{
                name = $Name
                displayName = $DisplayName
                description = if ($Description) { $Description } else { "$DisplayName template for SharePoint sites" }
                category = $Category
                baseTemplate = $BaseTemplate
                version = "1.0.0"
                created = Get-Date -Format "yyyy-MM-dd"
                author = if ($env:USERNAME) { $env:USERNAME } else { "MSP Automation" }
                securityBaseline = $SecurityBaseline
            }
            
            # Add client-specific settings if specified
            if ($ClientName) {
                $template.clients = @($ClientName)
                Write-SPOFactoryLog -Message "Template restricted to client: $ClientName" -Level Info
            }
            
            # Add document libraries
            if ($Libraries) {
                $template.libraries = @()
                foreach ($lib in $Libraries) {
                    $library = @{
                        name = $lib.name
                        displayName = if ($lib.displayName) { $lib.displayName } else { $lib.name }
                        versioning = if ($null -ne $lib.versioning) { $lib.versioning } else { $true }
                        checkOut = if ($null -ne $lib.checkOut) { $lib.checkOut } else { $false }
                        majorVersionLimit = if ($lib.majorVersionLimit) { $lib.majorVersionLimit } else { 50 }
                    }
                    
                    if ($lib.minorVersionLimit) {
                        $library.minorVersionLimit = $lib.minorVersionLimit
                    }
                    
                    if ($lib.contentTypes) {
                        $library.contentTypes = $lib.contentTypes
                    }
                    
                    $template.libraries += $library
                }
                Write-SPOFactoryLog -Message "Added $($template.libraries.Count) document libraries to template" -Level Info
            }
            else {
                # Add default document library
                $template.libraries = @(
                    @{
                        name = "Documents"
                        displayName = "Documents"
                        versioning = $true
                        checkOut = $false
                        majorVersionLimit = 50
                    }
                )
            }
            
            # Add lists
            if ($Lists) {
                $template.lists = @()
                foreach ($lst in $Lists) {
                    $list = @{
                        name = $lst.name
                        displayName = if ($lst.displayName) { $lst.displayName } else { $lst.name }
                        template = if ($lst.template) { $lst.template } else { "GenericList" }
                        description = if ($lst.description) { $lst.description } else { "" }
                    }
                    
                    if ($lst.columns) {
                        $list.columns = $lst.columns
                    }
                    
                    if ($lst.views) {
                        $list.views = $lst.views
                    }
                    
                    $template.lists += $list
                }
                Write-SPOFactoryLog -Message "Added $($template.lists.Count) lists to template" -Level Info
            }
            
            # Add features
            if ($Features) {
                $template.features = $Features
            }
            else {
                # Add default Office file handling feature
                $template.features = @('8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295')
            }
            Write-SPOFactoryLog -Message "Added $($template.features.Count) features to template" -Level Info
            
            # Add navigation structure
            if ($Navigation) {
                $template.navigation = $Navigation
                Write-SPOFactoryLog -Message "Added navigation structure to template" -Level Info
            }
            
            # Add theme
            if ($Theme) {
                $template.theme = $Theme
                Write-SPOFactoryLog -Message "Added theme to template: $Theme" -Level Info
            }
            
            # Add site scripts
            if ($SiteScripts) {
                $template.siteScripts = $SiteScripts
                Write-SPOFactoryLog -Message "Added $($SiteScripts.Count) site scripts to template" -Level Info
            }
            
            # Add permissions structure
            $template.permissions = @{
                breakInheritance = $false
                copyRoleAssignments = $true
                clearSubscopes = $false
            }
            
            # Add default settings
            $template.settings = @{
                regionalSettings = @{
                    timeZone = 13  # Eastern Time
                    locale = 1033  # English US
                    calendarType = 1
                    workDays = 62  # Monday through Friday
                    workDayStartHour = 480  # 8:00 AM
                    workDayEndHour = 1020  # 5:00 PM
                }
                sharingSettings = @{
                    sharingCapability = "ExternalUserAndGuestSharing"
                    defaultSharingLinkType = "Internal"
                    defaultLinkPermission = "View"
                    requireAnonymousLinksExpireInDays = 30
                }
            }
            
            # Save template to file
            if ($PSCmdlet.ShouldProcess($OutputPath, "Create Site Template")) {
                $template | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
                Write-SPOFactoryLog -Message "Template saved to: $OutputPath" -Level Info
                
                # Export to SharePoint if requested
                if ($ExportToSharePoint) {
                    Write-SPOFactoryLog -Message "Exporting template to SharePoint as site design" -Level Info
                    
                    $exportResult = Export-SPOTemplateToSiteDesign -Template $template -ClientName $ClientName
                    
                    if ($exportResult.Success) {
                        Write-SPOFactoryLog -Message "Template exported to SharePoint with ID: $($exportResult.SiteDesignId)" -Level Info
                    }
                    else {
                        Write-SPOFactoryLog -Message "Failed to export template to SharePoint: $($exportResult.Error)" -Level Warning
                    }
                }
                
                # Return template object
                [PSCustomObject]@{
                    Name = $Name
                    DisplayName = $DisplayName
                    Description = $template.description
                    Category = $Category
                    BaseTemplate = $BaseTemplate
                    Path = $OutputPath
                    Created = Get-Date
                    ExportedToSharePoint = $ExportToSharePoint.IsPresent
                }
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to create template: $_" -Level Error
            throw
        }
    }
    
    end {
        Write-SPOFactoryLog -Message "Template creation completed" -Level Info
    }
}

function Export-SPOTemplateToSiteDesign {
    <#
    .SYNOPSIS
        Exports a template to SharePoint as a site design.
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
        SiteDesignId = $null
        SiteScriptId = $null
        Error = $null
    }
    
    try {
        # Build site script from template
        $siteScript = @{
            '$schema' = 'https://developer.microsoft.com/json-schemas/sp/site-design-script-actions.schema.json'
            actions = @()
            bindings = @()
            version = 1
        }
        
        # Add library creation actions
        if ($Template.libraries) {
            foreach ($lib in $Template.libraries) {
                $siteScript.actions += @{
                    verb = 'createSPList'
                    listName = $lib.name
                    templateType = 101  # Document Library
                    subactions = @(
                        @{
                            verb = 'setTitle'
                            title = $lib.displayName
                        }
                    )
                }
                
                if ($lib.versioning) {
                    $siteScript.actions += @{
                        verb = 'setSPListVersioning'
                        listName = $lib.name
                        enableVersioning = $true
                        majorVersionLimit = $lib.majorVersionLimit
                    }
                }
            }
        }
        
        # Add list creation actions
        if ($Template.lists) {
            foreach ($lst in $Template.lists) {
                $templateType = switch ($lst.template) {
                    'GenericList' { 100 }
                    'TasksList' { 107 }
                    'Events' { 106 }
                    'Announcements' { 104 }
                    'Contacts' { 105 }
                    default { 100 }
                }
                
                $siteScript.actions += @{
                    verb = 'createSPList'
                    listName = $lst.name
                    templateType = $templateType
                    subactions = @(
                        @{
                            verb = 'setTitle'
                            title = $lst.displayName
                        },
                        @{
                            verb = 'setDescription'
                            description = $lst.description
                        }
                    )
                }
            }
        }
        
        # Add theme action if specified
        if ($Template.theme) {
            $siteScript.actions += @{
                verb = 'applyTheme'
                themeName = $Template.theme
            }
        }
        
        # Create the site script in SharePoint
        $scriptResult = Invoke-SPOFactoryCommand -ScriptBlock {
            $scriptJson = $siteScript | ConvertTo-Json -Depth 10
            $script = Add-PnPSiteScript -Title "$($Template.displayName) Script" -Content $scriptJson -Description $Template.description
            return $script
        } -ClientName $ClientName -Category 'Configuration' -ErrorMessage "Failed to create site script"
        
        if ($scriptResult) {
            $result.SiteScriptId = $scriptResult.Id
            
            # Create the site design
            $designResult = Invoke-SPOFactoryCommand -ScriptBlock {
                $webTemplate = switch ($Template.baseTemplate) {
                    'GROUP#0' { '64' }  # Team site
                    'SITEPAGEPUBLISHING#0' { '68' }  # Communication site
                    'STS#3' { '1' }  # Blank site
                    default { '64' }
                }
                
                $design = Add-PnPSiteDesign -Title $Template.displayName -SiteScriptIds $scriptResult.Id -WebTemplate $webTemplate -Description $Template.description
                return $design
            } -ClientName $ClientName -Category 'Configuration' -ErrorMessage "Failed to create site design"
            
            if ($designResult) {
                $result.SiteDesignId = $designResult.Id
                $result.Success = $true
            }
        }
        
        return $result
    }
    catch {
        $result.Error = $_.Exception.Message
        return $result
    }
}