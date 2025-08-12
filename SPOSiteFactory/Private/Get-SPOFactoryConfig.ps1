function Get-SPOFactoryConfig {
    <#
    .SYNOPSIS
        Retrieves SPOSiteFactory configuration with MSP multi-tenant support.

    .DESCRIPTION
        Gets configuration settings for SPOSiteFactory module with support for
        global MSP settings, client-specific configurations, and inherited defaults.
        Supports hierarchical configuration management for enterprise MSP environments.

    .PARAMETER ClientName
        Specific client name to retrieve configuration for

    .PARAMETER ConfigType
        Type of configuration to retrieve

    .PARAMETER Setting
        Specific setting name to retrieve

    .PARAMETER IncludeDefaults
        Include default values when retrieving client-specific configuration

    .PARAMETER AsHashtable
        Return configuration as hashtable instead of PSObject

    .EXAMPLE
        Get-SPOFactoryConfig

    .EXAMPLE
        Get-SPOFactoryConfig -ClientName "Contoso Corp"

    .EXAMPLE
        Get-SPOFactoryConfig -ClientName "Contoso Corp" -Setting "DefaultBaseline"

    .EXAMPLE
        Get-SPOFactoryConfig -ConfigType "Global" -AsHashtable
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Global', 'Client', 'Template', 'Security', 'All')]
        [string]$ConfigType = 'All',
        
        [Parameter(Mandatory = $false)]
        [string]$Setting,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeDefaults,
        
        [Parameter(Mandatory = $false)]
        [switch]$AsHashtable
    )

    begin {
        Write-SPOFactoryLog -Message "Retrieving SPOSiteFactory configuration" -Level Debug -ClientName $ClientName -Category 'Configuration'
        
        $configResult = @{}
    }

    process {
        try {
            # Get global configuration
            if ($ConfigType -in @('Global', 'All')) {
                $globalConfig = Get-SPOFactoryGlobalConfig
                if ($ConfigType -eq 'Global') {
                    $configResult = $globalConfig
                } else {
                    $configResult['Global'] = $globalConfig
                }
            }

            # Get client-specific configuration
            if ($ClientName -and $ConfigType -in @('Client', 'All')) {
                $clientConfig = Get-SPOFactoryClientConfig -ClientName $ClientName -IncludeDefaults:$IncludeDefaults
                if ($ConfigType -eq 'Client') {
                    $configResult = $clientConfig
                } else {
                    $configResult['Clients'] = @{}
                    $configResult['Clients'][$ClientName] = $clientConfig
                }
            }

            # Get template configurations
            if ($ConfigType -in @('Template', 'All')) {
                $templateConfig = Get-SPOFactoryTemplateConfig
                if ($ConfigType -eq 'Template') {
                    $configResult = $templateConfig
                } else {
                    $configResult['Templates'] = $templateConfig
                }
            }

            # Get security configurations
            if ($ConfigType -in @('Security', 'All')) {
                $securityConfig = Get-SPOFactorySecurityConfig
                if ($ConfigType -eq 'Security') {
                    $configResult = $securityConfig
                } else {
                    $configResult['Security'] = $securityConfig
                }
            }

            # Return specific setting if requested
            if ($Setting) {
                $configResult = Get-SPOFactoryConfigSetting -Configuration $configResult -Setting $Setting -ClientName $ClientName
            }

            # Return as hashtable or PSObject
            if ($AsHashtable) {
                return $configResult
            } else {
                return [PSCustomObject]$configResult
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve configuration: $_" -Level Error -ClientName $ClientName -Category 'Configuration' -Exception $_.Exception
            throw
        }
    }
}

function Set-SPOFactoryConfig {
    <#
    .SYNOPSIS
        Sets SPOSiteFactory configuration with validation and backup.

    .DESCRIPTION
        Updates configuration settings for SPOSiteFactory module with support for
        global MSP settings, client-specific configurations, and template management.
        Includes validation, backup, and rollback capabilities.

    .PARAMETER ClientName
        Specific client name to set configuration for

    .PARAMETER ConfigType
        Type of configuration to set

    .PARAMETER Setting
        Configuration setting name

    .PARAMETER Value
        Configuration value to set

    .PARAMETER Configuration
        Complete configuration object to set

    .PARAMETER Force
        Force setting without confirmation

    .PARAMETER BackupCurrent
        Create backup of current configuration before changes

    .PARAMETER ValidateOnly
        Only validate the configuration without setting it

    .EXAMPLE
        Set-SPOFactoryConfig -Setting "LogPath" -Value "C:\Logs\SPOFactory"

    .EXAMPLE
        Set-SPOFactoryConfig -ClientName "Contoso Corp" -Setting "DefaultBaseline" -Value "CustomBaseline"

    .EXAMPLE
        Set-SPOFactoryConfig -ClientName "Contoso Corp" -Configuration $clientConfig -BackupCurrent
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Global', 'Client', 'Template', 'Security')]
        [string]$ConfigType = 'Global',
        
        [Parameter(Mandatory = $false)]
        [string]$Setting,
        
        [Parameter(Mandatory = $false)]
        $Value,
        
        [Parameter(Mandatory = $false)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force,
        
        [Parameter(Mandatory = $false)]
        [switch]$BackupCurrent,
        
        [Parameter(Mandatory = $false)]
        [switch]$ValidateOnly
    )

    begin {
        Write-SPOFactoryLog -Message "Setting SPOSiteFactory configuration" -Level Info -ClientName $ClientName -Category 'Configuration'
        
        if ($BackupCurrent -and -not $ValidateOnly) {
            $backupPath = New-SPOFactoryConfigBackup -ClientName $ClientName -ConfigType $ConfigType
            Write-SPOFactoryLog -Message "Configuration backed up to: $backupPath" -Level Info -ClientName $ClientName -Category 'Configuration'
        }
    }

    process {
        try {
            # Validate inputs
            if (-not $Configuration -and (-not $Setting -or $null -eq $Value)) {
                throw "Either Configuration parameter or both Setting and Value parameters must be provided"
            }

            # Get current configuration for validation
            $currentConfig = Get-SPOFactoryConfig -ClientName $ClientName -ConfigType $ConfigType -AsHashtable

            # Prepare new configuration
            $newConfig = if ($Configuration) {
                $Configuration.Clone()
            } else {
                $currentConfig.Clone()
                $newConfig[$Setting] = $Value
                $newConfig
            }

            # Validate configuration
            $validationResult = Test-SPOFactoryConfig -Configuration $newConfig -ConfigType $ConfigType -ClientName $ClientName

            if (-not $validationResult.IsValid) {
                $errorMessage = "Configuration validation failed: $($validationResult.Errors -join '; ')"
                Write-SPOFactoryLog -Message $errorMessage -Level Error -ClientName $ClientName -Category 'Configuration'
                throw $errorMessage
            }

            if ($ValidateOnly) {
                Write-SPOFactoryLog -Message "Configuration validation passed" -Level Info -ClientName $ClientName -Category 'Configuration'
                return $validationResult
            }

            # Apply configuration changes
            if ($PSCmdlet.ShouldProcess("SPOSiteFactory Configuration", "Update $ConfigType configuration")) {
                switch ($ConfigType) {
                    'Global' {
                        Set-SPOFactoryGlobalConfig -Configuration $newConfig
                    }
                    'Client' {
                        if (-not $ClientName) {
                            throw "ClientName parameter is required for Client configuration type"
                        }
                        Set-SPOFactoryClientConfig -ClientName $ClientName -Configuration $newConfig
                    }
                    'Template' {
                        Set-SPOFactoryTemplateConfig -Configuration $newConfig
                    }
                    'Security' {
                        Set-SPOFactorySecurityConfig -Configuration $newConfig
                    }
                }

                # Log successful configuration update
                $changeDescription = if ($Setting) {
                    "Updated setting '$Setting' to '$Value'"
                } else {
                    "Updated complete $ConfigType configuration"
                }
                
                Write-SPOFactoryLog -Message "Configuration updated: $changeDescription" -Level Info -ClientName $ClientName -Category 'Configuration' -EnableAuditLog
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to set configuration: $_" -Level Error -ClientName $ClientName -Category 'Configuration' -Exception $_.Exception
            throw
        }
    }
}

function Get-SPOFactoryGlobalConfig {
    <#
    .SYNOPSIS
        Retrieves global MSP configuration settings.
    #>

    [CmdletBinding()]
    param()

    process {
        try {
            $globalConfigPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Global.json"
            
            if (Test-Path $globalConfigPath) {
                $savedConfig = Get-Content $globalConfigPath -Raw | ConvertFrom-Json -AsHashtable
                
                # Merge with script defaults, giving priority to saved values
                $mergedConfig = $script:SPOFactoryConfig.Clone()
                foreach ($key in $savedConfig.Keys) {
                    $mergedConfig[$key] = $savedConfig[$key]
                }
                
                return $mergedConfig
            } else {
                # Return script defaults
                return $script:SPOFactoryConfig.Clone()
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve global configuration, using defaults: $_" -Level Warning -Category 'Configuration'
            return $script:SPOFactoryConfig.Clone()
        }
    }
}

function Set-SPOFactoryGlobalConfig {
    <#
    .SYNOPSIS
        Sets global MSP configuration settings.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    process {
        try {
            $globalConfigPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Global.json"
            
            # Update script configuration
            foreach ($key in $Configuration.Keys) {
                $script:SPOFactoryConfig[$key] = $Configuration[$key]
            }
            
            # Save to file
            $Configuration | ConvertTo-Json -Depth 5 | Out-File -FilePath $globalConfigPath -Encoding UTF8
            
            Write-SPOFactoryLog -Message "Global configuration saved to: $globalConfigPath" -Level Debug -Category 'Configuration'
        }
        catch {
            throw "Failed to save global configuration: $_"
        }
    }
}

function Get-SPOFactoryClientConfig {
    <#
    .SYNOPSIS
        Retrieves client-specific configuration with inheritance.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [switch]$IncludeDefaults
    )

    process {
        try {
            $clientConfigPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Tenants\$ClientName.json"
            $clientConfig = @{}
            
            if (Test-Path $clientConfigPath) {
                $clientConfig = Get-Content $clientConfigPath -Raw | ConvertFrom-Json -AsHashtable
            }
            
            if ($IncludeDefaults) {
                # Start with global defaults
                $globalConfig = Get-SPOFactoryGlobalConfig
                $mergedConfig = $globalConfig.Clone()
                
                # Override with client-specific settings
                foreach ($key in $clientConfig.Keys) {
                    $mergedConfig[$key] = $clientConfig[$key]
                }
                
                return $mergedConfig
            } else {
                return $clientConfig
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve client configuration for $ClientName`: $_" -Level Warning -ClientName $ClientName -Category 'Configuration'
            
            if ($IncludeDefaults) {
                return Get-SPOFactoryGlobalConfig
            } else {
                return @{}
            }
        }
    }
}

function Set-SPOFactoryClientConfig {
    <#
    .SYNOPSIS
        Sets client-specific configuration.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    process {
        try {
            $tenantsPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Tenants"
            if (-not (Test-Path $tenantsPath)) {
                New-Item -Path $tenantsPath -ItemType Directory -Force | Out-Null
            }
            
            $clientConfigPath = Join-Path $tenantsPath "$ClientName.json"
            
            # Add metadata
            $configWithMetadata = $Configuration.Clone()
            $configWithMetadata['_metadata'] = @{
                ClientName = $ClientName
                LastModified = Get-Date
                ModuleVersion = $script:SPOFactoryConstants.ModuleVersion
                ConfigVersion = $script:SPOFactoryConstants.ConfigVersion
            }
            
            # Save configuration
            $configWithMetadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $clientConfigPath -Encoding UTF8
            
            Write-SPOFactoryLog -Message "Client configuration saved for: $ClientName" -Level Debug -ClientName $ClientName -Category 'Configuration'
        }
        catch {
            throw "Failed to save client configuration for $ClientName`: $_"
        }
    }
}

function Get-SPOFactoryTemplateConfig {
    <#
    .SYNOPSIS
        Retrieves template configurations.
    #>

    [CmdletBinding()]
    param()

    process {
        try {
            $templatesPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Templates"
            $templates = @{}
            
            if (Test-Path $templatesPath) {
                $templateFiles = Get-ChildItem -Path $templatesPath -Filter "*.json"
                foreach ($templateFile in $templateFiles) {
                    $templateName = [System.IO.Path]::GetFileNameWithoutExtension($templateFile.Name)
                    $templates[$templateName] = Get-Content $templateFile.FullName -Raw | ConvertFrom-Json -AsHashtable
                }
            }
            
            return $templates
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve template configurations: $_" -Level Warning -Category 'Configuration'
            return @{}
        }
    }
}

function Set-SPOFactoryTemplateConfig {
    <#
    .SYNOPSIS
        Sets template configurations.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    process {
        try {
            $templatesPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Templates"
            if (-not (Test-Path $templatesPath)) {
                New-Item -Path $templatesPath -ItemType Directory -Force | Out-Null
            }
            
            foreach ($templateName in $Configuration.Keys) {
                $templatePath = Join-Path $templatesPath "$templateName.json"
                $Configuration[$templateName] | ConvertTo-Json -Depth 5 | Out-File -FilePath $templatePath -Encoding UTF8
            }
            
            Write-SPOFactoryLog -Message "Template configurations saved" -Level Debug -Category 'Configuration'
        }
        catch {
            throw "Failed to save template configurations: $_"
        }
    }
}

function Get-SPOFactorySecurityConfig {
    <#
    .SYNOPSIS
        Retrieves security-related configurations.
    #>

    [CmdletBinding()]
    param()

    process {
        try {
            $securityConfigPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Security.json"
            
            $defaultSecurityConfig = @{
                EncryptCredentials = $true
                AuditAllOperations = $true
                LogRetentionDays = 90
                RequireSecureConnection = $true
                AllowedRegions = @('Global', 'GCC')
                MaxConcurrentConnections = 50
                ConnectionTimeout = 300
                CertificateValidation = $true
            }
            
            if (Test-Path $securityConfigPath) {
                $savedConfig = Get-Content $securityConfigPath -Raw | ConvertFrom-Json -AsHashtable
                
                # Merge with defaults
                foreach ($key in $defaultSecurityConfig.Keys) {
                    if (-not $savedConfig.ContainsKey($key)) {
                        $savedConfig[$key] = $defaultSecurityConfig[$key]
                    }
                }
                
                return $savedConfig
            } else {
                return $defaultSecurityConfig
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve security configuration, using defaults: $_" -Level Warning -Category 'Configuration'
            return @{
                EncryptCredentials = $true
                AuditAllOperations = $true
            }
        }
    }
}

function Set-SPOFactorySecurityConfig {
    <#
    .SYNOPSIS
        Sets security-related configurations.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration
    )

    process {
        try {
            $securityConfigPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Security.json"
            
            # Add metadata
            $configWithMetadata = $Configuration.Clone()
            $configWithMetadata['_metadata'] = @{
                LastModified = Get-Date
                ModuleVersion = $script:SPOFactoryConstants.ModuleVersion
                ConfigVersion = $script:SPOFactoryConstants.ConfigVersion
            }
            
            $configWithMetadata | ConvertTo-Json -Depth 5 | Out-File -FilePath $securityConfigPath -Encoding UTF8
            
            Write-SPOFactoryLog -Message "Security configuration saved" -Level Info -Category 'Security' -EnableAuditLog
        }
        catch {
            throw "Failed to save security configuration: $_"
        }
    }
}

function Test-SPOFactoryConfig {
    <#
    .SYNOPSIS
        Validates SPOSiteFactory configuration.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory = $false)]
        [string]$ConfigType = 'Global',
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    process {
        $validationResult = @{
            IsValid = $true
            Errors = @()
            Warnings = @()
        }

        try {
            # Validate required settings based on config type
            switch ($ConfigType) {
                'Global' {
                    $requiredSettings = @('LogPath', 'ConfigPath', 'CredentialVault')
                    foreach ($setting in $requiredSettings) {
                        if (-not $Configuration.ContainsKey($setting) -or [string]::IsNullOrWhiteSpace($Configuration[$setting])) {
                            $validationResult.Errors += "Required setting '$setting' is missing or empty"
                            $validationResult.IsValid = $false
                        }
                    }
                    
                    # Validate numeric settings
                    $numericSettings = @{
                        'MaxConcurrentConnections' = @{ Min = 1; Max = 100 }
                        'ConnectionTimeout' = @{ Min = 30; Max = 600 }
                        'RetryAttempts' = @{ Min = 1; Max = 10 }
                        'BatchSize' = @{ Min = 1; Max = 1000 }
                    }
                    
                    foreach ($setting in $numericSettings.Keys) {
                        if ($Configuration.ContainsKey($setting)) {
                            $value = $Configuration[$setting]
                            $min = $numericSettings[$setting].Min
                            $max = $numericSettings[$setting].Max
                            
                            if ($value -lt $min -or $value -gt $max) {
                                $validationResult.Errors += "Setting '$setting' must be between $min and $max"
                                $validationResult.IsValid = $false
                            }
                        }
                    }
                }
                
                'Client' {
                    if (-not $ClientName) {
                        $validationResult.Errors += "ClientName is required for client configuration validation"
                        $validationResult.IsValid = $false
                    }
                    
                    # Validate client-specific settings
                    if ($Configuration.ContainsKey('TenantUrl')) {
                        if ($Configuration.TenantUrl -notmatch '^https://[a-zA-Z0-9.-]+\.sharepoint\.(com|us|de|cn)/?') {
                            $validationResult.Errors += "Invalid TenantUrl format"
                            $validationResult.IsValid = $false
                        }
                    }
                }
                
                'Security' {
                    # Validate security settings
                    if ($Configuration.ContainsKey('AllowedRegions')) {
                        $validRegions = @('Global', 'GCC', 'GCCH', 'DoD')
                        foreach ($region in $Configuration.AllowedRegions) {
                            if ($region -notin $validRegions) {
                                $validationResult.Warnings += "Unknown region: $region"
                            }
                        }
                    }
                }
            }

            # Validate paths exist
            $pathSettings = @('LogPath', 'ConfigPath')
            foreach ($setting in $pathSettings) {
                if ($Configuration.ContainsKey($setting)) {
                    $path = $Configuration[$setting]
                    if (-not (Test-Path $path)) {
                        try {
                            New-Item -Path $path -ItemType Directory -Force | Out-Null
                            $validationResult.Warnings += "Created missing directory: $path"
                        }
                        catch {
                            $validationResult.Errors += "Cannot create or access path '$path': $_"
                            $validationResult.IsValid = $false
                        }
                    }
                }
            }
        }
        catch {
            $validationResult.Errors += "Configuration validation error: $_"
            $validationResult.IsValid = $false
        }

        return $validationResult
    }
}

function Get-SPOFactoryConfigSetting {
    <#
    .SYNOPSIS
        Retrieves a specific configuration setting with hierarchy support.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Configuration,
        
        [Parameter(Mandatory = $true)]
        [string]$Setting,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientName
    )

    process {
        # Try client-specific first
        if ($ClientName -and $Configuration.ContainsKey('Clients') -and $Configuration.Clients.ContainsKey($ClientName)) {
            if ($Configuration.Clients[$ClientName].ContainsKey($Setting)) {
                return $Configuration.Clients[$ClientName][$Setting]
            }
        }
        
        # Try global configuration
        if ($Configuration.ContainsKey('Global') -and $Configuration.Global.ContainsKey($Setting)) {
            return $Configuration.Global[$Setting]
        }
        
        # Try direct access
        if ($Configuration.ContainsKey($Setting)) {
            return $Configuration[$Setting]
        }
        
        return $null
    }
}

function New-SPOFactoryConfigBackup {
    <#
    .SYNOPSIS
        Creates backup of current configuration.
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$ConfigType = 'Global'
    )

    process {
        try {
            $backupPath = Join-Path $script:SPOFactoryConfig.ConfigPath "Backups"
            if (-not (Test-Path $backupPath)) {
                New-Item -Path $backupPath -ItemType Directory -Force | Out-Null
            }
            
            $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $backupFileName = if ($ClientName) {
                "Config_${ConfigType}_${ClientName}_${timestamp}.json"
            } else {
                "Config_${ConfigType}_${timestamp}.json"
            }
            
            $backupFilePath = Join-Path $backupPath $backupFileName
            
            # Get current configuration
            $currentConfig = Get-SPOFactoryConfig -ClientName $ClientName -ConfigType $ConfigType -AsHashtable
            
            # Save backup
            $currentConfig | ConvertTo-Json -Depth 5 | Out-File -FilePath $backupFilePath -Encoding UTF8
            
            return $backupFilePath
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to create configuration backup: $_" -Level Warning -ClientName $ClientName -Category 'Configuration'
            throw
        }
    }
}