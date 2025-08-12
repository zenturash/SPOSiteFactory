function New-SPOSiteFromConfig {
    <#
    .SYNOPSIS
        Creates SharePoint sites from a configuration file or object.
    
    .DESCRIPTION
        The New-SPOSiteFromConfig function creates multiple SharePoint sites based on a JSON configuration file
        or hashtable object. It supports hub sites, team sites, communication sites, and automatic hub associations.
        Designed for MSP environments with multi-tenant support and bulk provisioning capabilities.
    
    .PARAMETER ConfigPath
        Path to a JSON configuration file containing site definitions.
    
    .PARAMETER Configuration
        A hashtable or PSCustomObject containing site configuration.
    
    .PARAMETER ClientName
        The MSP client name for multi-tenant scenarios. Overrides the client specified in configuration.
    
    .PARAMETER ValidateOnly
        If specified, validates the configuration without creating sites.
    
    .PARAMETER SkipExisting
        If specified, skips sites that already exist instead of throwing an error.
    
    .PARAMETER MaxConcurrent
        Maximum number of concurrent site creation operations. Default is 5.
    
    .PARAMETER GenerateReport
        If specified, generates a detailed report of the provisioning operation.
    
    .PARAMETER ReportPath
        Path where the provisioning report should be saved.
    
    .EXAMPLE
        New-SPOSiteFromConfig -ConfigPath "C:\Config\sites.json" -ClientName "Contoso"
        
        Creates sites from a JSON configuration file for the Contoso client.
    
    .EXAMPLE
        $config = @{
            hubSite = @{
                title = "Corporate Hub"
                url = "corp-hub"
                securityBaseline = "MSPSecure"
            }
            sites = @(
                @{
                    title = "Finance"
                    url = "finance"
                    type = "TeamSite"
                    joinHub = $true
                }
            )
        }
        New-SPOSiteFromConfig -Configuration $config -GenerateReport
        
        Creates sites from a hashtable configuration and generates a report.
    
    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
        Requires: SharePoint Online Management Shell, PnP.PowerShell
    #>
    
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'File')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'File')]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Configuration file not found: $_"
            }
            if ($_ -notmatch '\.json$') {
                throw "Configuration file must be a JSON file"
            }
            $true
        })]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'Object')]
        [ValidateNotNull()]
        [object]$Configuration,
        
        [Parameter(Mandatory = $false)]
        [ValidatePattern('^[a-zA-Z0-9-]+$')]
        [string]$ClientName,
        
        [Parameter()]
        [switch]$ValidateOnly,
        
        [Parameter()]
        [switch]$SkipExisting,
        
        [Parameter()]
        [ValidateRange(1, 10)]
        [int]$MaxConcurrent = 5,
        
        [Parameter()]
        [switch]$GenerateReport,
        
        [Parameter()]
        [string]$ReportPath = "$env:TEMP\SPOProvisioning_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    )
    
    begin {
        Write-SPOFactoryLog -Message "Starting configuration-based site provisioning" -Level Info
        
        # Initialize report structure
        $report = @{
            StartTime = Get-Date
            Configuration = $null
            ValidationResults = @()
            ProvisioningResults = @()
            Summary = @{
                TotalSites = 0
                Successful = 0
                Failed = 0
                Skipped = 0
            }
        }
        
        # Load configuration
        try {
            if ($PSCmdlet.ParameterSetName -eq 'File') {
                Write-SPOFactoryLog -Message "Loading configuration from file: $ConfigPath" -Level Info
                $configContent = Get-Content -Path $ConfigPath -Raw
                $config = $configContent | ConvertFrom-Json -AsHashtable
            }
            else {
                Write-SPOFactoryLog -Message "Using provided configuration object" -Level Info
                $config = $Configuration
                if ($Configuration -is [string]) {
                    $config = $Configuration | ConvertFrom-Json -AsHashtable
                }
            }
            
            $report.Configuration = $config
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to load configuration: $_" -Level Error
            throw "Configuration loading failed: $_"
        }
        
        # Override client name if specified
        if ($ClientName) {
            Write-SPOFactoryLog -Message "Overriding client name with: $ClientName" -Level Info
            if ($config -is [hashtable]) {
                $config['client'] = $ClientName
            }
        }
        elseif ($config.client) {
            $ClientName = $config.client
        }
        else {
            Write-SPOFactoryLog -Message "No client name specified, using default" -Level Warning
            $ClientName = "Default"
        }
    }
    
    process {
        # Validate configuration structure
        Write-SPOFactoryLog -Message "Validating configuration structure" -Level Info
        $validationErrors = @()
        
        # Check for required elements
        if (-not $config.sites -and -not $config.hubSite) {
            $validationErrors += "Configuration must contain either 'sites' or 'hubSite' element"
        }
        
        # Validate hub site if present
        if ($config.hubSite) {
            if (-not $config.hubSite.title) {
                $validationErrors += "Hub site must have a title"
            }
            if (-not $config.hubSite.url) {
                $validationErrors += "Hub site must have a URL"
            }
        }
        
        # Validate sites array
        if ($config.sites) {
            $siteIndex = 0
            foreach ($site in $config.sites) {
                if (-not $site.title) {
                    $validationErrors += "Site at index $siteIndex must have a title"
                }
                if (-not $site.url) {
                    $validationErrors += "Site at index $siteIndex must have a URL"
                }
                if ($site.type -and $site.type -notin @('TeamSite', 'CommunicationSite')) {
                    $validationErrors += "Site at index $siteIndex has invalid type: $($site.type)"
                }
                $siteIndex++
            }
        }
        
        # Check for validation errors
        if ($validationErrors.Count -gt 0) {
            $report.ValidationResults = $validationErrors
            Write-SPOFactoryLog -Message "Configuration validation failed with $($validationErrors.Count) errors" -Level Error
            foreach ($error in $validationErrors) {
                Write-SPOFactoryLog -Message "Validation error: $error" -Level Error
            }
            
            if ($GenerateReport) {
                $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $ReportPath -Encoding UTF8
            }
            
            throw "Configuration validation failed. Check the log for details."
        }
        
        Write-SPOFactoryLog -Message "Configuration validation successful" -Level Info
        
        # If ValidateOnly flag is set, return validation results
        if ($ValidateOnly) {
            Write-SPOFactoryLog -Message "ValidateOnly flag set, skipping site creation" -Level Info
            $validationResult = [PSCustomObject]@{
                IsValid = $true
                Client = $ClientName
                HubSite = if ($config.hubSite) { $config.hubSite.title } else { $null }
                SiteCount = if ($config.sites) { $config.sites.Count } else { 0 }
                ValidationTime = Get-Date
            }
            return $validationResult
        }
        
        # Process hub site creation first if specified
        $hubSiteUrl = $null
        if ($config.hubSite) {
            Write-SPOFactoryLog -Message "Creating hub site: $($config.hubSite.title)" -Level Info
            
            try {
                $hubParams = @{
                    Title = $config.hubSite.title
                    Url = $config.hubSite.url
                    ClientName = $ClientName
                }
                
                # Add optional parameters
                if ($config.hubSite.description) {
                    $hubParams['Description'] = $config.hubSite.description
                }
                if ($config.hubSite.securityBaseline) {
                    $hubParams['SecurityBaseline'] = $config.hubSite.securityBaseline
                }
                if ($config.hubSite.owners) {
                    $hubParams['Owners'] = $config.hubSite.owners
                }
                if ($config.hubSite.applySecurityImmediately) {
                    $hubParams['ApplySecurityImmediately'] = $true
                }
                
                if ($PSCmdlet.ShouldProcess("$($config.hubSite.title)", "Create Hub Site")) {
                    $hubResult = New-SPOHubSite @hubParams -ErrorAction Stop
                    $hubSiteUrl = $hubResult.Url
                    
                    $report.ProvisioningResults += @{
                        Type = 'HubSite'
                        Title = $config.hubSite.title
                        Url = $hubSiteUrl
                        Status = 'Success'
                        Message = 'Hub site created successfully'
                        Timestamp = Get-Date
                    }
                    $report.Summary.Successful++
                    
                    Write-SPOFactoryLog -Message "Hub site created successfully: $hubSiteUrl" -Level Info
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                if ($errorMessage -match 'already exists' -and $SkipExisting) {
                    Write-SPOFactoryLog -Message "Hub site already exists, skipping: $($config.hubSite.url)" -Level Warning
                    $report.ProvisioningResults += @{
                        Type = 'HubSite'
                        Title = $config.hubSite.title
                        Url = $config.hubSite.url
                        Status = 'Skipped'
                        Message = 'Site already exists'
                        Timestamp = Get-Date
                    }
                    $report.Summary.Skipped++
                    
                    # Try to get the existing hub site URL
                    try {
                        $existingHub = Get-PnPHubSite | Where-Object { $_.Title -eq $config.hubSite.title }
                        if ($existingHub) {
                            $hubSiteUrl = $existingHub.SiteUrl
                        }
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Could not retrieve existing hub site URL" -Level Warning
                    }
                }
                else {
                    Write-SPOFactoryLog -Message "Failed to create hub site: $errorMessage" -Level Error
                    $report.ProvisioningResults += @{
                        Type = 'HubSite'
                        Title = $config.hubSite.title
                        Url = $config.hubSite.url
                        Status = 'Failed'
                        Message = $errorMessage
                        Timestamp = Get-Date
                    }
                    $report.Summary.Failed++
                    
                    if (-not $SkipExisting) {
                        throw
                    }
                }
            }
        }
        
        # Process sites array
        if ($config.sites) {
            $report.Summary.TotalSites = $config.sites.Count
            
            # Create runspace pool for parallel processing
            $runspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxConcurrent)
            $runspacePool.Open()
            
            $jobs = @()
            
            foreach ($site in $config.sites) {
                $scriptBlock = {
                    param($site, $ClientName, $hubSiteUrl, $SkipExisting)
                    
                    try {
                        # Build parameters for site creation
                        $siteParams = @{
                            Title = $site.title
                            Url = $site.url
                            ClientName = $ClientName
                        }
                        
                        # Determine site type
                        if ($site.type -eq 'CommunicationSite') {
                            $siteParams['CommunicationSite'] = $true
                        }
                        else {
                            $siteParams['TeamSite'] = $true
                        }
                        
                        # Add optional parameters
                        if ($site.description) {
                            $siteParams['Description'] = $site.description
                        }
                        if ($site.securityBaseline) {
                            $siteParams['SecurityBaseline'] = $site.securityBaseline
                        }
                        if ($site.owners) {
                            $siteParams['Owners'] = $site.owners
                        }
                        if ($site.members) {
                            $siteParams['Members'] = $site.members
                        }
                        
                        # Add hub association if specified
                        if ($site.joinHub -and $hubSiteUrl) {
                            $siteParams['HubSiteUrl'] = $hubSiteUrl
                        }
                        elseif ($site.hubSite) {
                            $siteParams['HubSiteUrl'] = $site.hubSite
                        }
                        
                        # Create the site
                        $result = New-SPOSite @siteParams -ErrorAction Stop
                        
                        return @{
                            Type = $site.type
                            Title = $site.title
                            Url = $result.Url
                            Status = 'Success'
                            Message = 'Site created successfully'
                            Timestamp = Get-Date
                        }
                    }
                    catch {
                        $errorMessage = $_.Exception.Message
                        if ($errorMessage -match 'already exists' -and $SkipExisting) {
                            return @{
                                Type = $site.type
                                Title = $site.title
                                Url = $site.url
                                Status = 'Skipped'
                                Message = 'Site already exists'
                                Timestamp = Get-Date
                            }
                        }
                        else {
                            return @{
                                Type = $site.type
                                Title = $site.title
                                Url = $site.url
                                Status = 'Failed'
                                Message = $errorMessage
                                Timestamp = Get-Date
                            }
                        }
                    }
                }
                
                # Create and start runspace
                $powershell = [powershell]::Create().AddScript($scriptBlock)
                $powershell.AddArgument($site).AddArgument($ClientName).AddArgument($hubSiteUrl).AddArgument($SkipExisting) | Out-Null
                $powershell.RunspacePool = $runspacePool
                
                $jobs += [PSCustomObject]@{
                    PowerShell = $powershell
                    Handle = $powershell.BeginInvoke()
                    Site = $site
                }
                
                Write-SPOFactoryLog -Message "Started provisioning job for site: $($site.title)" -Level Info
            }
            
            # Wait for all jobs to complete
            Write-SPOFactoryLog -Message "Waiting for $($jobs.Count) provisioning jobs to complete" -Level Info
            $completedJobs = 0
            
            while ($jobs.Handle.IsCompleted -contains $false) {
                $completed = ($jobs.Handle.IsCompleted -eq $true).Count
                if ($completed -gt $completedJobs) {
                    $completedJobs = $completed
                    Write-Progress -Activity "Creating Sites" -Status "$completedJobs of $($jobs.Count) completed" `
                        -PercentComplete (($completedJobs / $jobs.Count) * 100)
                }
                Start-Sleep -Milliseconds 500
            }
            
            Write-Progress -Activity "Creating Sites" -Completed
            
            # Collect results
            foreach ($job in $jobs) {
                try {
                    $result = $job.PowerShell.EndInvoke($job.Handle)
                    $report.ProvisioningResults += $result
                    
                    switch ($result.Status) {
                        'Success' { $report.Summary.Successful++ }
                        'Failed' { $report.Summary.Failed++ }
                        'Skipped' { $report.Summary.Skipped++ }
                    }
                    
                    Write-SPOFactoryLog -Message "Site '$($result.Title)' provisioning status: $($result.Status)" -Level Info
                }
                catch {
                    Write-SPOFactoryLog -Message "Error collecting results for site '$($job.Site.title)': $_" -Level Error
                    $report.ProvisioningResults += @{
                        Type = $job.Site.type
                        Title = $job.Site.title
                        Url = $job.Site.url
                        Status = 'Failed'
                        Message = $_.Exception.Message
                        Timestamp = Get-Date
                    }
                    $report.Summary.Failed++
                }
                finally {
                    $job.PowerShell.Dispose()
                }
            }
            
            # Clean up runspace pool
            $runspacePool.Close()
            $runspacePool.Dispose()
        }
        
        # Complete report
        $report.EndTime = Get-Date
        $report.Duration = $report.EndTime - $report.StartTime
        
        # Generate and save report if requested
        if ($GenerateReport) {
            Write-SPOFactoryLog -Message "Generating provisioning report: $ReportPath" -Level Info
            $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $ReportPath -Encoding UTF8
            
            Write-Host "`nProvisioning Report Summary:" -ForegroundColor Cyan
            Write-Host "  Total Sites: $($report.Summary.TotalSites)" -ForegroundColor White
            Write-Host "  Successful: $($report.Summary.Successful)" -ForegroundColor Green
            Write-Host "  Failed: $($report.Summary.Failed)" -ForegroundColor Red
            Write-Host "  Skipped: $($report.Summary.Skipped)" -ForegroundColor Yellow
            Write-Host "  Report saved to: $ReportPath" -ForegroundColor Gray
        }
        
        # Return summary object
        [PSCustomObject]@{
            Client = $ClientName
            TotalSites = $report.Summary.TotalSites
            Successful = $report.Summary.Successful
            Failed = $report.Summary.Failed
            Skipped = $report.Summary.Skipped
            Duration = $report.Duration
            ReportPath = if ($GenerateReport) { $ReportPath } else { $null }
        }
    }
    
    end {
        Write-SPOFactoryLog -Message "Configuration-based site provisioning completed" -Level Info
    }
}