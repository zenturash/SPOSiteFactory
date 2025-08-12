function Get-SPOSiteTemplate {
    <#
    .SYNOPSIS
        Retrieves SharePoint site templates for MSP multi-tenant environments.
    
    .DESCRIPTION
        The Get-SPOSiteTemplate function retrieves site templates from the Data/Templates directory
        or from SharePoint Online. Templates define site structure, libraries, features, and settings
        for consistent site provisioning across MSP clients.
    
    .PARAMETER Name
        Name of the specific template to retrieve.
    
    .PARAMETER Path
        Path to a custom template file.
    
    .PARAMETER ClientName
        The MSP client name for multi-tenant scenarios.
    
    .PARAMETER BuiltIn
        If specified, retrieves built-in SharePoint templates.
    
    .PARAMETER Online
        If specified, retrieves templates from SharePoint Online.
    
    .PARAMETER Category
        Filter templates by category (Project, Department, Communication, etc.).
    
    .EXAMPLE
        Get-SPOSiteTemplate -Name "ProjectSite"
        
        Retrieves the ProjectSite template from the local template library.
    
    .EXAMPLE
        Get-SPOSiteTemplate -BuiltIn
        
        Retrieves all built-in SharePoint site templates.
    
    .EXAMPLE
        Get-SPOSiteTemplate -Category "Department" -ClientName "Contoso"
        
        Retrieves all department templates for the Contoso client.
    
    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
        Requires: SharePoint Online Management Shell, PnP.PowerShell
    #>
    
    [CmdletBinding(DefaultParameterSetName = 'Local')]
    param(
        [Parameter(ParameterSetName = 'Local')]
        [Parameter(ParameterSetName = 'Online')]
        [string]$Name,
        
        [Parameter(ParameterSetName = 'File')]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Template file not found: $_"
            }
            if ($_ -notmatch '\.(json|xml)$') {
                throw "Template file must be JSON or XML format"
            }
            $true
        })]
        [string]$Path,
        
        [Parameter()]
        [string]$ClientName,
        
        [Parameter(ParameterSetName = 'BuiltIn')]
        [switch]$BuiltIn,
        
        [Parameter(ParameterSetName = 'Online')]
        [switch]$Online,
        
        [Parameter()]
        [ValidateSet('Project', 'Department', 'Communication', 'TeamSite', 'Custom')]
        [string]$Category
    )
    
    begin {
        Write-SPOFactoryLog -Message "Retrieving site templates" -Level Info
        
        $templates = @()
        
        # Get module base path
        $modulePath = Split-Path -Parent $PSScriptRoot
        $templatesPath = Join-Path $modulePath "Data\Templates"
        
        # Create templates directory if it doesn't exist
        if (-not (Test-Path $templatesPath)) {
            New-Item -Path $templatesPath -ItemType Directory -Force | Out-Null
            Write-SPOFactoryLog -Message "Created templates directory: $templatesPath" -Level Info
        }
    }
    
    process {
        try {
            if ($BuiltIn) {
                # Get built-in SharePoint templates
                Write-SPOFactoryLog -Message "Retrieving built-in SharePoint templates" -Level Info
                
                $builtInTemplates = @(
                    @{
                        Name = 'TeamSite'
                        Id = 'GROUP#0'
                        DisplayName = 'Team Site'
                        Description = 'Create a team site with Microsoft 365 Group'
                        Category = 'TeamSite'
                        Type = 'Built-In'
                    },
                    @{
                        Name = 'CommunicationSite'
                        Id = 'SITEPAGEPUBLISHING#0'
                        DisplayName = 'Communication Site'
                        Description = 'Create a communication site for broadcasting information'
                        Category = 'Communication'
                        Type = 'Built-In'
                    },
                    @{
                        Name = 'BlankSite'
                        Id = 'STS#3'
                        DisplayName = 'Blank Site'
                        Description = 'Create a blank site with no pre-configured content'
                        Category = 'Custom'
                        Type = 'Built-In'
                    }
                )
                
                if ($Category) {
                    $builtInTemplates = $builtInTemplates | Where-Object { $_.Category -eq $Category }
                }
                
                if ($Name) {
                    $builtInTemplates = $builtInTemplates | Where-Object { $_.Name -eq $Name }
                }
                
                foreach ($template in $builtInTemplates) {
                    $templates += [PSCustomObject]$template
                }
            }
            elseif ($Online) {
                # Get templates from SharePoint Online
                Write-SPOFactoryLog -Message "Retrieving templates from SharePoint Online" -Level Info
                
                $onlineTemplates = Invoke-SPOFactoryCommand -ScriptBlock {
                    $siteDesigns = Get-PnPSiteDesign
                    $siteScripts = Get-PnPSiteScript
                    
                    $templates = @()
                    foreach ($design in $siteDesigns) {
                        $templates += @{
                            Name = $design.Title
                            Id = $design.Id
                            DisplayName = $design.Title
                            Description = $design.Description
                            WebTemplate = $design.WebTemplate
                            SiteScriptIds = $design.SiteScriptIds
                            Type = 'Online'
                            Version = $design.Version
                        }
                    }
                    
                    return $templates
                } -ClientName $ClientName -Category 'Configuration' -SuppressErrors
                
                if ($onlineTemplates) {
                    if ($Name) {
                        $onlineTemplates = $onlineTemplates | Where-Object { $_.Name -eq $Name }
                    }
                    
                    foreach ($template in $onlineTemplates) {
                        $templates += [PSCustomObject]$template
                    }
                }
            }
            elseif ($Path) {
                # Load template from file
                Write-SPOFactoryLog -Message "Loading template from file: $Path" -Level Info
                
                $templateContent = Get-Content -Path $Path -Raw
                
                if ($Path -match '\.json$') {
                    $template = $templateContent | ConvertFrom-Json
                }
                else {
                    $template = [xml]$templateContent
                }
                
                $templates += [PSCustomObject]@{
                    Name = $template.name
                    DisplayName = $template.displayName
                    Description = $template.description
                    Category = $template.category
                    Type = 'File'
                    Path = $Path
                    Configuration = $template
                }
            }
            else {
                # Get local templates from Data/Templates directory
                Write-SPOFactoryLog -Message "Retrieving local templates from $templatesPath" -Level Info
                
                # Create default templates if they don't exist
                $defaultTemplates = @(
                    @{
                        name = 'StandardTeamSite'
                        displayName = 'Standard Team Site'
                        description = 'Standard team site with document libraries and lists'
                        category = 'TeamSite'
                        baseTemplate = 'GROUP#0'
                        libraries = @(
                            @{
                                name = 'Documents'
                                versioning = $true
                                checkOut = $false
                                majorVersionLimit = 50
                            },
                            @{
                                name = 'Site Assets'
                                versioning = $true
                                checkOut = $false
                            }
                        )
                        features = @('8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295')
                        securityBaseline = 'MSPStandard'
                    },
                    @{
                        name = 'ProjectSite'
                        displayName = 'Project Site'
                        description = 'Project collaboration site with task tracking'
                        category = 'Project'
                        baseTemplate = 'STS#3'
                        libraries = @(
                            @{
                                name = 'Project Documents'
                                versioning = $true
                                checkOut = $true
                                majorVersionLimit = 10
                                minorVersionLimit = 5
                            },
                            @{
                                name = 'Deliverables'
                                versioning = $true
                                checkOut = $false
                            }
                        )
                        lists = @(
                            @{
                                name = 'Project Tasks'
                                template = 'TasksList'
                            },
                            @{
                                name = 'Project Calendar'
                                template = 'Events'
                            }
                        )
                        features = @('8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295')
                        securityBaseline = 'MSPSecure'
                    },
                    @{
                        name = 'DepartmentSite'
                        displayName = 'Department Site'
                        description = 'Department collaboration and information sharing'
                        category = 'Department'
                        baseTemplate = 'SITEPAGEPUBLISHING#0'
                        libraries = @(
                            @{
                                name = 'Department Documents'
                                versioning = $true
                                checkOut = $false
                            },
                            @{
                                name = 'Policies'
                                versioning = $true
                                checkOut = $true
                            }
                        )
                        features = @('8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295')
                        securityBaseline = 'MSPStandard'
                    },
                    @{
                        name = 'CommunicationSite'
                        displayName = 'Communication Site'
                        description = 'Corporate communication and announcements'
                        category = 'Communication'
                        baseTemplate = 'SITEPAGEPUBLISHING#0'
                        libraries = @(
                            @{
                                name = 'Site Pages'
                                versioning = $true
                                checkOut = $false
                            },
                            @{
                                name = 'Site Assets'
                                versioning = $false
                                checkOut = $false
                            }
                        )
                        features = @('8A4B8DE2-6FD8-41e9-923C-C7C3C00F8295')
                        securityBaseline = 'MSPStandard'
                    }
                )
                
                # Create default template files if they don't exist
                foreach ($defaultTemplate in $defaultTemplates) {
                    $templateFilePath = Join-Path $templatesPath "$($defaultTemplate.name).json"
                    
                    if (-not (Test-Path $templateFilePath)) {
                        $defaultTemplate | ConvertTo-Json -Depth 10 | Out-File -FilePath $templateFilePath -Encoding UTF8
                        Write-SPOFactoryLog -Message "Created default template: $($defaultTemplate.name)" -Level Info
                    }
                }
                
                # Load all templates from directory
                $templateFiles = Get-ChildItem -Path $templatesPath -Filter "*.json" -ErrorAction SilentlyContinue
                
                foreach ($file in $templateFiles) {
                    try {
                        $templateContent = Get-Content -Path $file.FullName -Raw
                        $template = $templateContent | ConvertFrom-Json
                        
                        # Apply filters
                        if ($Name -and $template.name -ne $Name) {
                            continue
                        }
                        
                        if ($Category -and $template.category -ne $Category) {
                            continue
                        }
                        
                        if ($ClientName) {
                            # Check if template has client-specific settings
                            if ($template.clients -and $ClientName -notin $template.clients) {
                                continue
                            }
                        }
                        
                        $templates += [PSCustomObject]@{
                            Name = $template.name
                            DisplayName = $template.displayName
                            Description = $template.description
                            Category = $template.category
                            BaseTemplate = $template.baseTemplate
                            Type = 'Local'
                            Path = $file.FullName
                            Configuration = $template
                        }
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Failed to load template from $($file.Name): $_" -Level Warning
                    }
                }
            }
            
            # Return templates
            if ($templates.Count -eq 0) {
                Write-SPOFactoryLog -Message "No templates found matching criteria" -Level Warning
            }
            else {
                Write-SPOFactoryLog -Message "Found $($templates.Count) templates" -Level Info
            }
            
            return $templates
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve templates: $_" -Level Error
            throw
        }
    }
    
    end {
        Write-SPOFactoryLog -Message "Template retrieval completed" -Level Info
    }
}