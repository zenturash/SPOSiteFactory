function New-SPOBulkSites {
    <#
    .SYNOPSIS
        Creates multiple SharePoint sites in bulk with parallel processing.
    
    .DESCRIPTION
        The New-SPOBulkSites function creates multiple SharePoint sites efficiently using parallel processing.
        It supports both sequential and parallel execution, progress tracking, and comprehensive error handling.
        Designed for MSP environments to provision large numbers of sites across multiple clients.
    
    .PARAMETER Sites
        An array of site objects containing site configuration details.
    
    .PARAMETER ConfigPath
        Path to a CSV or JSON file containing site definitions.
    
    .PARAMETER ClientName
        The MSP client name for multi-tenant scenarios.
    
    .PARAMETER BatchSize
        Number of sites to process in each batch. Default is 10.
    
    .PARAMETER Parallel
        If specified, creates sites in parallel for improved performance.
    
    .PARAMETER ThrottleLimit
        Maximum number of concurrent operations when using parallel processing. Default is 5.
    
    .PARAMETER ContinueOnError
        If specified, continues processing remaining sites even if some fail.
    
    .PARAMETER RetryFailedSites
        If specified, automatically retries failed site creations.
    
    .PARAMETER MaxRetries
        Maximum number of retry attempts for failed sites. Default is 2.
    
    .PARAMETER GenerateReport
        If specified, generates a detailed report of the bulk operation.
    
    .PARAMETER ReportPath
        Path where the bulk operation report should be saved.
    
    .EXAMPLE
        New-SPOBulkSites -ConfigPath "C:\Sites\bulk-sites.csv" -ClientName "Contoso" -Parallel
        
        Creates sites from a CSV file using parallel processing.
    
    .EXAMPLE
        $sites = @(
            @{Title="Site1"; Url="site1"; Type="TeamSite"},
            @{Title="Site2"; Url="site2"; Type="CommunicationSite"}
        )
        New-SPOBulkSites -Sites $sites -BatchSize 5 -GenerateReport
        
        Creates sites from an array with batch processing and generates a report.
    
    .NOTES
        Author: MSP Automation Team
        Version: 1.0.0
        Requires: SharePoint Online Management Shell, PnP.PowerShell
    #>
    
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Array')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Array', ValueFromPipeline = $true)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Sites,
        
        [Parameter(Mandatory = $true, ParameterSetName = 'File')]
        [ValidateScript({
            if (-not (Test-Path $_)) {
                throw "Configuration file not found: $_"
            }
            if ($_ -notmatch '\.(csv|json)$') {
                throw "Configuration file must be CSV or JSON format"
            }
            $true
        })]
        [string]$ConfigPath,
        
        [Parameter(Mandatory = $false)]
        [ValidatePattern('^[a-zA-Z0-9-]+$')]
        [string]$ClientName,
        
        [Parameter()]
        [ValidateRange(1, 100)]
        [int]$BatchSize = 10,
        
        [Parameter()]
        [switch]$Parallel,
        
        [Parameter()]
        [ValidateRange(1, 20)]
        [int]$ThrottleLimit = 5,
        
        [Parameter()]
        [switch]$ContinueOnError,
        
        [Parameter()]
        [switch]$RetryFailedSites,
        
        [Parameter()]
        [ValidateRange(1, 5)]
        [int]$MaxRetries = 2,
        
        [Parameter()]
        [switch]$GenerateReport,
        
        [Parameter()]
        [string]$ReportPath = "$env:TEMP\SPOBulkSites_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    )
    
    begin {
        Write-SPOFactoryLog -Message "Starting bulk site creation" -Level Info
        
        # Initialize tracking variables
        $script:totalSites = 0
        $script:processedSites = 0
        $script:successfulSites = 0
        $script:failedSites = 0
        $script:retriedSites = 0
        
        # Initialize report structure
        $report = @{
            StartTime = Get-Date
            Configuration = @{
                ClientName = $ClientName
                BatchSize = $BatchSize
                Parallel = $Parallel.IsPresent
                ThrottleLimit = $ThrottleLimit
                ContinueOnError = $ContinueOnError.IsPresent
                RetryFailedSites = $RetryFailedSites.IsPresent
                MaxRetries = $MaxRetries
            }
            Batches = @()
            Sites = @()
            Summary = @{
                TotalSites = 0
                Successful = 0
                Failed = 0
                Retried = 0
                Duration = $null
            }
        }
        
        # Load sites from file if specified
        if ($PSCmdlet.ParameterSetName -eq 'File') {
            Write-SPOFactoryLog -Message "Loading sites from file: $ConfigPath" -Level Info
            
            try {
                if ($ConfigPath -match '\.csv$') {
                    $Sites = Import-Csv -Path $ConfigPath
                    Write-SPOFactoryLog -Message "Loaded $($Sites.Count) sites from CSV" -Level Info
                }
                elseif ($ConfigPath -match '\.json$') {
                    $jsonContent = Get-Content -Path $ConfigPath -Raw
                    $Sites = $jsonContent | ConvertFrom-Json
                    Write-SPOFactoryLog -Message "Loaded $($Sites.Count) sites from JSON" -Level Info
                }
            }
            catch {
                Write-SPOFactoryLog -Message "Failed to load sites from file: $_" -Level Error
                throw "Failed to load configuration file: $_"
            }
        }
        
        # Validate sites array
        if (-not $Sites -or $Sites.Count -eq 0) {
            throw "No sites provided for bulk creation"
        }
        
        $script:totalSites = $Sites.Count
        $report.Summary.TotalSites = $script:totalSites
        
        Write-SPOFactoryLog -Message "Preparing to create $($script:totalSites) sites" -Level Info
        
        # Set default client name if not specified
        if (-not $ClientName) {
            $ClientName = "Default"
            Write-SPOFactoryLog -Message "No client name specified, using default" -Level Warning
        }
    }
    
    process {
        # Function to create a single site
        function New-SingleSite {
            param($SiteConfig, $ClientName, $RetryCount = 0)
            
            try {
                # Build parameters for site creation
                $siteParams = @{
                    Title = $SiteConfig.Title
                    Url = $SiteConfig.Url
                    ClientName = $ClientName
                }
                
                # Determine site type
                if ($SiteConfig.Type -eq 'CommunicationSite' -or $SiteConfig.Template -eq 'SITEPAGEPUBLISHING#0') {
                    $siteParams['CommunicationSite'] = $true
                }
                else {
                    $siteParams['TeamSite'] = $true
                }
                
                # Add optional parameters
                if ($SiteConfig.Description) {
                    $siteParams['Description'] = $SiteConfig.Description
                }
                if ($SiteConfig.SecurityBaseline) {
                    $siteParams['SecurityBaseline'] = $SiteConfig.SecurityBaseline
                }
                if ($SiteConfig.Owners) {
                    $owners = if ($SiteConfig.Owners -is [string]) {
                        $SiteConfig.Owners -split ';|,'
                    } else {
                        $SiteConfig.Owners
                    }
                    $siteParams['Owners'] = $owners
                }
                if ($SiteConfig.Members) {
                    $members = if ($SiteConfig.Members -is [string]) {
                        $SiteConfig.Members -split ';|,'
                    } else {
                        $SiteConfig.Members
                    }
                    $siteParams['Members'] = $members
                }
                if ($SiteConfig.HubSiteUrl) {
                    $siteParams['HubSiteUrl'] = $SiteConfig.HubSiteUrl
                }
                
                # Create the site
                $startTime = Get-Date
                $result = New-SPOSite @siteParams -ErrorAction Stop
                $duration = (Get-Date) - $startTime
                
                return @{
                    Title = $SiteConfig.Title
                    Url = $result.Url
                    Type = $SiteConfig.Type
                    Status = 'Success'
                    Message = 'Site created successfully'
                    Duration = $duration
                    RetryCount = $RetryCount
                    Timestamp = Get-Date
                }
            }
            catch {
                $errorMessage = $_.Exception.Message
                
                # Check if site already exists
                if ($errorMessage -match 'already exists') {
                    return @{
                        Title = $SiteConfig.Title
                        Url = $SiteConfig.Url
                        Type = $SiteConfig.Type
                        Status = 'AlreadyExists'
                        Message = 'Site already exists'
                        Duration = $null
                        RetryCount = $RetryCount
                        Timestamp = Get-Date
                    }
                }
                
                # Return failure result
                return @{
                    Title = $SiteConfig.Title
                    Url = $SiteConfig.Url
                    Type = $SiteConfig.Type
                    Status = 'Failed'
                    Message = $errorMessage
                    Duration = $null
                    RetryCount = $RetryCount
                    Timestamp = Get-Date
                    Error = $_
                }
            }
        }
        
        # Process sites in batches
        $batches = for ($i = 0; $i -lt $Sites.Count; $i += $BatchSize) {
            $endIndex = [Math]::Min($i + $BatchSize - 1, $Sites.Count - 1)
            ,@($Sites[$i..$endIndex])
        }
        
        Write-SPOFactoryLog -Message "Processing $($batches.Count) batches of sites" -Level Info
        
        $batchNumber = 0
        foreach ($batch in $batches) {
            $batchNumber++
            $batchStartTime = Get-Date
            
            Write-SPOFactoryLog -Message "Processing batch $batchNumber of $($batches.Count) ($(

.Count) sites)" -Level Info
            Write-Progress -Activity "Bulk Site Creation" -Status "Processing batch $batchNumber of $($batches.Count)" `
                -PercentComplete (($batchNumber - 1) / $batches.Count * 100)
            
            $batchReport = @{
                BatchNumber = $batchNumber
                StartTime = $batchStartTime
                Sites = @()
                Successful = 0
                Failed = 0
            }
            
            if ($Parallel) {
                # Parallel processing using runspaces
                Write-SPOFactoryLog -Message "Using parallel processing with throttle limit: $ThrottleLimit" -Level Info
                
                $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
                $runspacePool.Open()
                
                $jobs = @()
                
                foreach ($site in $batch) {
                    $scriptBlock = {
                        param($SiteConfig, $ClientName, $MaxRetries, $RetryFailedSites)
                        
                        # Load required modules in runspace
                        Import-Module SPOSiteFactory -ErrorAction SilentlyContinue
                        
                        # Function to create site (defined inline for runspace)
                        function New-SingleSiteInRunspace {
                            param($SiteConfig, $ClientName)
                            
                            try {
                                $siteParams = @{
                                    Title = $SiteConfig.Title
                                    Url = $SiteConfig.Url
                                    ClientName = $ClientName
                                }
                                
                                if ($SiteConfig.Type -eq 'CommunicationSite') {
                                    $siteParams['CommunicationSite'] = $true
                                } else {
                                    $siteParams['TeamSite'] = $true
                                }
                                
                                if ($SiteConfig.Description) { $siteParams['Description'] = $SiteConfig.Description }
                                if ($SiteConfig.SecurityBaseline) { $siteParams['SecurityBaseline'] = $SiteConfig.SecurityBaseline }
                                if ($SiteConfig.Owners) { $siteParams['Owners'] = $SiteConfig.Owners -split ';|,' }
                                if ($SiteConfig.Members) { $siteParams['Members'] = $SiteConfig.Members -split ';|,' }
                                if ($SiteConfig.HubSiteUrl) { $siteParams['HubSiteUrl'] = $SiteConfig.HubSiteUrl }
                                
                                $result = New-SPOSite @siteParams -ErrorAction Stop
                                
                                return @{
                                    Title = $SiteConfig.Title
                                    Url = $result.Url
                                    Type = $SiteConfig.Type
                                    Status = 'Success'
                                    Message = 'Site created successfully'
                                    Timestamp = Get-Date
                                }
                            }
                            catch {
                                return @{
                                    Title = $SiteConfig.Title
                                    Url = $SiteConfig.Url
                                    Type = $SiteConfig.Type
                                    Status = 'Failed'
                                    Message = $_.Exception.Message
                                    Timestamp = Get-Date
                                }
                            }
                        }
                        
                        # Attempt site creation with retries
                        $retryCount = 0
                        $result = $null
                        
                        do {
                            $result = New-SingleSiteInRunspace -SiteConfig $SiteConfig -ClientName $ClientName
                            
                            if ($result.Status -eq 'Failed' -and $RetryFailedSites -and $retryCount -lt $MaxRetries) {
                                $retryCount++
                                Start-Sleep -Seconds (2 * $retryCount)  # Exponential backoff
                            }
                            else {
                                break
                            }
                        } while ($retryCount -le $MaxRetries)
                        
                        $result['RetryCount'] = $retryCount
                        return $result
                    }
                    
                    $powershell = [powershell]::Create().AddScript($scriptBlock)
                    $powershell.AddArgument($site).AddArgument($ClientName).AddArgument($MaxRetries).AddArgument($RetryFailedSites) | Out-Null
                    $powershell.RunspacePool = $runspacePool
                    
                    $jobs += [PSCustomObject]@{
                        PowerShell = $powershell
                        Handle = $powershell.BeginInvoke()
                        Site = $site
                    }
                }
                
                # Wait for all jobs in batch to complete
                while ($jobs.Handle.IsCompleted -contains $false) {
                    Start-Sleep -Milliseconds 500
                }
                
                # Collect results
                foreach ($job in $jobs) {
                    try {
                        $result = $job.PowerShell.EndInvoke($job.Handle)
                        $batchReport.Sites += $result
                        $report.Sites += $result
                        
                        if ($result.Status -eq 'Success') {
                            $batchReport.Successful++
                            $script:successfulSites++
                        }
                        else {
                            $batchReport.Failed++
                            $script:failedSites++
                            
                            if (-not $ContinueOnError -and $result.Status -eq 'Failed') {
                                Write-SPOFactoryLog -Message "Stopping bulk operation due to failure (ContinueOnError not set)" -Level Error
                                throw "Site creation failed for $($result.Title): $($result.Message)"
                            }
                        }
                        
                        if ($result.RetryCount -gt 0) {
                            $script:retriedSites++
                        }
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Error collecting results for site '$($job.Site.Title)': $_" -Level Error
                        $batchReport.Failed++
                        $script:failedSites++
                    }
                    finally {
                        $job.PowerShell.Dispose()
                    }
                }
                
                # Clean up runspace pool
                $runspacePool.Close()
                $runspacePool.Dispose()
            }
            else {
                # Sequential processing
                Write-SPOFactoryLog -Message "Using sequential processing" -Level Info
                
                foreach ($site in $batch) {
                    if ($PSCmdlet.ShouldProcess($site.Title, "Create Site")) {
                        $result = New-SingleSite -SiteConfig $site -ClientName $ClientName
                        
                        # Handle retries if enabled
                        if ($result.Status -eq 'Failed' -and $RetryFailedSites) {
                            $retryCount = 0
                            while ($retryCount -lt $MaxRetries -and $result.Status -eq 'Failed') {
                                $retryCount++
                                Write-SPOFactoryLog -Message "Retrying site creation for '$($site.Title)' (attempt $retryCount of $MaxRetries)" -Level Warning
                                Start-Sleep -Seconds (2 * $retryCount)  # Exponential backoff
                                $result = New-SingleSite -SiteConfig $site -ClientName $ClientName -RetryCount $retryCount
                            }
                            
                            if ($retryCount -gt 0) {
                                $script:retriedSites++
                            }
                        }
                        
                        $batchReport.Sites += $result
                        $report.Sites += $result
                        
                        if ($result.Status -eq 'Success' -or $result.Status -eq 'AlreadyExists') {
                            $batchReport.Successful++
                            $script:successfulSites++
                        }
                        else {
                            $batchReport.Failed++
                            $script:failedSites++
                            
                            if (-not $ContinueOnError) {
                                Write-SPOFactoryLog -Message "Stopping bulk operation due to failure (ContinueOnError not set)" -Level Error
                                throw "Site creation failed for $($result.Title): $($result.Message)"
                            }
                        }
                    }
                    
                    $script:processedSites++
                    
                    # Update progress for individual sites within batch
                    $overallProgress = (($batchNumber - 1) * $BatchSize + $batch.IndexOf($site) + 1) / $script:totalSites * 100
                    Write-Progress -Activity "Bulk Site Creation" `
                        -Status "Batch $batchNumber of $($batches.Count) - Site '$($site.Title)'" `
                        -PercentComplete $overallProgress
                }
            }
            
            $batchReport.EndTime = Get-Date
            $batchReport.Duration = $batchReport.EndTime - $batchStartTime
            $report.Batches += $batchReport
            
            Write-SPOFactoryLog -Message "Batch $batchNumber completed - Successful: $($batchReport.Successful), Failed: $($batchReport.Failed)" -Level Info
        }
        
        Write-Progress -Activity "Bulk Site Creation" -Completed
    }
    
    end {
        # Calculate final statistics
        $report.EndTime = Get-Date
        $report.Summary.Duration = $report.EndTime - $report.StartTime
        $report.Summary.Successful = $script:successfulSites
        $report.Summary.Failed = $script:failedSites
        $report.Summary.Retried = $script:retriedSites
        
        # Generate and save report if requested
        if ($GenerateReport) {
            Write-SPOFactoryLog -Message "Generating bulk operation report: $ReportPath" -Level Info
            $report | ConvertTo-Json -Depth 10 | Out-File -FilePath $ReportPath -Encoding UTF8
        }
        
        # Display summary
        Write-Host "`nBulk Site Creation Summary:" -ForegroundColor Cyan
        Write-Host "  Total Sites: $($script:totalSites)" -ForegroundColor White
        Write-Host "  Successful: $($script:successfulSites)" -ForegroundColor Green
        Write-Host "  Failed: $($script:failedSites)" -ForegroundColor $(if ($script:failedSites -gt 0) { 'Red' } else { 'Gray' })
        if ($script:retriedSites -gt 0) {
            Write-Host "  Retried: $($script:retriedSites)" -ForegroundColor Yellow
        }
        Write-Host "  Duration: $($report.Summary.Duration.ToString('mm\:ss'))" -ForegroundColor Gray
        
        if ($GenerateReport) {
            Write-Host "  Report saved to: $ReportPath" -ForegroundColor Gray
        }
        
        Write-SPOFactoryLog -Message "Bulk site creation completed - Total: $($script:totalSites), Successful: $($script:successfulSites), Failed: $($script:failedSites)" -Level Info
        
        # Return summary object
        [PSCustomObject]@{
            TotalSites = $script:totalSites
            Successful = $script:successfulSites
            Failed = $script:failedSites
            Retried = $script:retriedSites
            Duration = $report.Summary.Duration
            Batches = $batches.Count
            ReportPath = if ($GenerateReport) { $ReportPath } else { $null }
        }
    }
}