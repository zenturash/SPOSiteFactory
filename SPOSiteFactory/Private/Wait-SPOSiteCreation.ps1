function Wait-SPOSiteCreation {
    <#
    .SYNOPSIS
        Waits for SharePoint Online site creation to complete with comprehensive monitoring.

    .DESCRIPTION
        Enterprise-grade function for waiting on SharePoint Online site provisioning completion
        in MSP environments. Provides detailed monitoring, timeout handling, progress reporting,
        and proper error handling for multiple tenant scenarios.

    .PARAMETER SiteUrl
        The SharePoint site URL to monitor for completion

    .PARAMETER TimeoutMinutes
        Maximum time to wait for site creation in minutes (default: 30)

    .PARAMETER PollingInterval
        Interval between status checks in seconds (default: 30)

    .PARAMETER ClientName
        Client name for MSP tenant isolation and logging

    .PARAMETER ExpectedTitle
        Expected site title for validation once site is available

    .PARAMETER ExpectedOwner
        Expected site owner for validation once site is available

    .PARAMETER ShowProgress
        Display progress bar during wait operation

    .PARAMETER FailOnTimeout
        Throw error if timeout is reached (default: true)

    .PARAMETER ValidateConfiguration
        Perform additional validation once site is available

    .EXAMPLE
        Wait-SPOSiteCreation -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpTeam" -ClientName "ContosoCorp" -TimeoutMinutes 45

    .EXAMPLE
        $result = Wait-SPOSiteCreation -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpHub" -ClientName "ContosoCorp" -ExpectedTitle "Contoso Hub" -ShowProgress

    .EXAMPLE
        Wait-SPOSiteCreation -SiteUrl "https://contoso.sharepoint.com/sites/ContosoCorpComm" -ClientName "ContosoCorp" -ValidateConfiguration -FailOnTimeout:$false
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [int]$TimeoutMinutes = 30,
        
        [Parameter(Mandatory = $false)]
        [int]$PollingInterval = 30,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$ExpectedTitle,
        
        [Parameter(Mandatory = $false)]
        [string]$ExpectedOwner,
        
        [Parameter(Mandatory = $false)]
        [switch]$ShowProgress,
        
        [Parameter(Mandatory = $false)]
        [switch]$FailOnTimeout = $true,
        
        [Parameter(Mandatory = $false)]
        [switch]$ValidateConfiguration
    )

    begin {
        $operationId = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)
        $startTime = Get-Date
        $timeoutTime = $startTime.AddMinutes($TimeoutMinutes)
        $attempt = 0
        $maxAttempts = [Math]::Ceiling($TimeoutMinutes * 60 / $PollingInterval)

        Write-SPOFactoryLog -Message "Starting site creation monitoring for: $SiteUrl (Timeout: $TimeoutMinutes minutes)" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('WaitStart', $operationId)

        $result = @{
            Success = $false
            SiteAvailable = $false
            TimedOut = $false
            ElapsedTime = $null
            AttemptCount = 0
            Site = $null
            ValidationResults = @()
            ProvisioningStages = @()
            FinalStatus = 'Unknown'
        }

        # Initialize progress tracking
        if ($ShowProgress) {
            Write-Progress -Id 1 -Activity "Waiting for site creation" -Status "Starting monitoring..." -PercentComplete 0
        }
    }

    process {
        try {
            while ((Get-Date) -lt $timeoutTime) {
                $attempt++
                $result.AttemptCount = $attempt
                $elapsedMinutes = [Math]::Round(((Get-Date) - $startTime).TotalMinutes, 1)
                $percentComplete = [Math]::Min(($elapsedMinutes / $TimeoutMinutes) * 100, 95)

                Write-SPOFactoryLog -Message "Checking site availability - Attempt $attempt/$maxAttempts (Elapsed: $elapsedMinutes min)" -Level Debug -ClientName $ClientName -Category 'Provisioning' -Tag @('WaitCheck', $operationId)

                if ($ShowProgress) {
                    Write-Progress -Id 1 -Activity "Waiting for site creation" -Status "Checking site availability (Attempt $attempt)" -PercentComplete $percentComplete
                }

                # Get detailed provisioning status
                $provisioningStatus = Get-SPOProvisioningStatus -SiteUrl $SiteUrl -ClientName $ClientName

                # Track provisioning stages
                $stageInfo = @{
                    Timestamp = Get-Date
                    Attempt = $attempt
                    Status = $provisioningStatus.Status
                    IsAvailable = $provisioningStatus.IsAvailable
                    Details = $provisioningStatus.Details
                }
                $result.ProvisioningStages += $stageInfo

                # Check if site is available
                if ($provisioningStatus.IsAvailable) {
                    Write-SPOFactoryLog -Message "Site is available, performing validation checks" -Level Info -ClientName $ClientName -Category 'Provisioning' -Tag @('SiteAvailable', $operationId)
                    
                    $result.SiteAvailable = $true
                    $result.Site = $provisioningStatus.Site

                    # Validate site if requested
                    if ($ValidateConfiguration -or $ExpectedTitle -or $ExpectedOwner) {
                        $validationResult = Test-SPOSiteConfiguration -SiteUrl $SiteUrl -ExpectedTitle $ExpectedTitle -ExpectedOwner $ExpectedOwner -ClientName $ClientName
                        $result.ValidationResults = $validationResult

                        if (-not $validationResult.IsValid) {
                            Write-SPOFactoryLog -Message "Site validation failed: $($validationResult.Issues -join '; ')" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Tag @('ValidationFailed', $operationId)
                        }
                    }

                    # Wait a bit more for full provisioning
                    Write-SPOFactoryLog -Message "Site available, waiting additional 60 seconds for full provisioning" -Level Debug -ClientName $ClientName -Category 'Provisioning'
                    Start-Sleep -Seconds 60

                    # Final availability check
                    $finalCheck = Test-SPOSiteFinalAvailability -SiteUrl $SiteUrl -ClientName $ClientName
                    if ($finalCheck.IsFullyProvisioned) {
                        $result.Success = $true
                        $result.FinalStatus = 'Completed'
                        break
                    } else {
                        Write-SPOFactoryLog -Message "Site partially available, continuing to wait" -Level Debug -ClientName $ClientName -Category 'Provisioning'
                        $result.FinalStatus = 'PartiallyAvailable'
                    }
                } else {
                    # Log current status
                    $statusMessage = "Site not yet available - Status: $($provisioningStatus.Status)"
                    if ($provisioningStatus.Details) {
                        $statusMessage += " | Details: $($provisioningStatus.Details)"
                    }
                    Write-SPOFactoryLog -Message $statusMessage -Level Debug -ClientName $ClientName -Category 'Provisioning' -Tag @('WaitContinue', $operationId)
                    $result.FinalStatus = $provisioningStatus.Status
                }

                # Check for provisioning errors
                if ($provisioningStatus.HasError) {
                    Write-SPOFactoryLog -Message "Site provisioning error detected: $($provisioningStatus.ErrorMessage)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Tag @('ProvisioningError', $operationId)
                    $result.FinalStatus = 'Error'
                    break
                }

                # Wait before next check
                Start-Sleep -Seconds $PollingInterval
            }

            # Handle timeout
            if (-not $result.Success -and (Get-Date) -ge $timeoutTime) {
                $result.TimedOut = $true
                $result.FinalStatus = 'TimedOut'
                Write-SPOFactoryLog -Message "Site creation monitoring timed out after $TimeoutMinutes minutes" -Level Warning -ClientName $ClientName -Category 'Provisioning' -Tag @('WaitTimeout', $operationId)
            }

        }
        catch {
            Write-SPOFactoryLog -Message "Error during site creation monitoring: $($_.Exception.Message)" -Level Error -ClientName $ClientName -Category 'Provisioning' -Exception $_.Exception -Tag @('WaitError', $operationId)
            $result.FinalStatus = 'Error'
            
            if ($FailOnTimeout) {
                throw
            }
        }
        finally {
            # Calculate final elapsed time
            $result.ElapsedTime = (Get-Date) - $startTime

            # Close progress bar
            if ($ShowProgress) {
                if ($result.Success) {
                    Write-Progress -Id 1 -Activity "Site creation complete" -Status "Success" -PercentComplete 100
                } else {
                    Write-Progress -Id 1 -Activity "Site creation monitoring ended" -Status $result.FinalStatus -PercentComplete 100
                }
                Start-Sleep -Seconds 1
                Write-Progress -Id 1 -Activity "Complete" -Completed
            }

            # Log final result
            $finalMessage = "Site creation monitoring completed - Status: $($result.FinalStatus) | Elapsed: $($result.ElapsedTime.ToString('mm\:ss')) | Attempts: $($result.AttemptCount)"
            $logLevel = if ($result.Success) { 'Info' } else { 'Warning' }
            Write-SPOFactoryLog -Message $finalMessage -Level $logLevel -ClientName $ClientName -Category 'Provisioning' -Tag @('WaitComplete', $operationId)

            # Handle failure conditions
            if (-not $result.Success -and $FailOnTimeout) {
                if ($result.TimedOut) {
                    $errorMessage = "Site creation timed out after $TimeoutMinutes minutes. Site URL: $SiteUrl"
                } else {
                    $errorMessage = "Site creation monitoring failed. Status: $($result.FinalStatus). Site URL: $SiteUrl"
                }
                
                throw [System.TimeoutException]::new($errorMessage)
            }
        }
    }

    end {
        return $result
    }
}

function Test-SPOSiteConfiguration {
    <#
    .SYNOPSIS
        Validates site configuration against expected values.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ExpectedTitle,
        
        [Parameter(Mandatory = $false)]
        [string]$ExpectedOwner,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $validation = @{
        IsValid = $true
        Issues = @()
        Checks = @()
    }

    try {
        $site = Invoke-SPOFactoryCommand -ScriptBlock {
            Get-PnPSite -Identity $SiteUrl -Includes Owner
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        if (-not $site) {
            $validation.IsValid = $false
            $validation.Issues += "Unable to retrieve site information"
            return $validation
        }

        # Title validation
        if ($ExpectedTitle -and $site.Title -ne $ExpectedTitle) {
            $validation.Issues += "Site title mismatch. Expected: '$ExpectedTitle', Actual: '$($site.Title)'"
            $validation.IsValid = $false
        }
        $validation.Checks += @{ Property = 'Title'; Expected = $ExpectedTitle; Actual = $site.Title; Valid = ($site.Title -eq $ExpectedTitle) }

        # Owner validation  
        if ($ExpectedOwner -and $site.Owner -ne $ExpectedOwner) {
            $validation.Issues += "Site owner mismatch. Expected: '$ExpectedOwner', Actual: '$($site.Owner)'"
            $validation.IsValid = $false
        }
        $validation.Checks += @{ Property = 'Owner'; Expected = $ExpectedOwner; Actual = $site.Owner; Valid = ($site.Owner -eq $ExpectedOwner) }

        # Basic health checks
        if ($site.Status -ne 'Active') {
            $validation.Issues += "Site status is not Active: $($site.Status)"
            $validation.IsValid = $false
        }
        $validation.Checks += @{ Property = 'Status'; Expected = 'Active'; Actual = $site.Status; Valid = ($site.Status -eq 'Active') }

        return $validation
    }
    catch {
        $validation.IsValid = $false
        $validation.Issues += "Error validating site configuration: $($_.Exception.Message)"
        return $validation
    }
}

function Test-SPOSiteFinalAvailability {
    <#
    .SYNOPSIS
        Performs comprehensive final availability check.
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    $result = @{
        IsFullyProvisioned = $false
        AvailabilityChecks = @()
        Issues = @()
    }

    try {
        # Check site accessibility
        $siteCheck = Invoke-SPOFactoryCommand -ScriptBlock {
            $web = Get-PnPWeb -ErrorAction Stop
            return @{
                Title = $web.Title
                Id = $web.Id
                ServerRelativeUrl = $web.ServerRelativeUrl
                Created = $web.Created
                IsAvailable = $true
            }
        } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

        if ($siteCheck) {
            $result.AvailabilityChecks += @{ Check = 'WebAccess'; Status = 'Success'; Details = "Site web accessible" }
            
            # Check document library
            $libCheck = Invoke-SPOFactoryCommand -ScriptBlock {
                $lists = Get-PnPList -ErrorAction SilentlyContinue
                return $lists | Where-Object { $_.BaseTemplate -eq 101 } | Select-Object -First 1
            } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

            if ($libCheck) {
                $result.AvailabilityChecks += @{ Check = 'DocumentLibrary'; Status = 'Success'; Details = "Document library available" }
            } else {
                $result.AvailabilityChecks += @{ Check = 'DocumentLibrary'; Status = 'Warning'; Details = "Document library not found" }
                $result.Issues += "Document library not available"
            }

            # Check permissions
            $permCheck = Invoke-SPOFactoryCommand -ScriptBlock {
                $perms = Get-PnPWeb -Includes HasUniqueRoleAssignments,RoleAssignments -ErrorAction SilentlyContinue
                return $perms.RoleAssignments.Count -gt 0
            } -ClientName $ClientName -Category 'Provisioning' -SuppressErrors

            if ($permCheck) {
                $result.AvailabilityChecks += @{ Check = 'Permissions'; Status = 'Success'; Details = "Permissions configured" }
            } else {
                $result.AvailabilityChecks += @{ Check = 'Permissions'; Status = 'Warning'; Details = "Permissions may not be fully configured" }
            }

            # Determine if fully provisioned (allow some warnings)
            $criticalFailures = $result.AvailabilityChecks | Where-Object { $_.Status -eq 'Error' }
            $result.IsFullyProvisioned = ($criticalFailures.Count -eq 0)

        } else {
            $result.AvailabilityChecks += @{ Check = 'WebAccess'; Status = 'Error'; Details = "Unable to access site web" }
            $result.Issues += "Site web not accessible"
        }

        return $result
    }
    catch {
        $result.Issues += "Error checking final availability: $($_.Exception.Message)"
        $result.AvailabilityChecks += @{ Check = 'FinalCheck'; Status = 'Error'; Details = $_.Exception.Message }
        return $result
    }
}