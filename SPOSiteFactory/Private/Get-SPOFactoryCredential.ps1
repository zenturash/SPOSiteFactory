function Get-SPOFactoryCredential {
    <#
    .SYNOPSIS
        Retrieves stored credentials for SharePoint Online clients in MSP environments.

    .DESCRIPTION
        Securely retrieves stored credentials from the configured secret vault for
        MSP client tenants. Supports multiple authentication types and provides
        fallback mechanisms for credential retrieval.

    .PARAMETER ClientName
        The client name to retrieve credentials for

    .PARAMETER AuthType
        Type of authentication credentials to retrieve

    .PARAMETER VaultName
        Override the default secret vault name

    .PARAMETER NoCache
        Skip credential caching

    .EXAMPLE
        Get-SPOFactoryCredential -ClientName "Contoso Corp"

    .EXAMPLE
        Get-SPOFactoryCredential -ClientName "Contoso Corp" -AuthType "Certificate"
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Password', 'Certificate', 'ClientSecret', 'Token')]
        [string]$AuthType = 'Password',
        
        [Parameter(Mandatory = $false)]
        [string]$VaultName = $script:SPOFactoryConfig.CredentialVault,
        
        [Parameter(Mandatory = $false)]
        [switch]$NoCache
    )

    begin {
        Write-SPOFactoryLog -Message "Retrieving credentials for client: $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
        
        # Check if credential is cached (unless NoCache is specified)
        if (-not $NoCache -and $script:SPOFactoryCredentialCache) {
            $cacheKey = "$ClientName|$AuthType"
            if ($script:SPOFactoryCredentialCache.ContainsKey($cacheKey)) {
                $cachedCred = $script:SPOFactoryCredentialCache[$cacheKey]
                if ($cachedCred.ExpiresAt -gt (Get-Date)) {
                    Write-SPOFactoryLog -Message "Using cached credential for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                    return $cachedCred.Credential
                }
            }
        }
    }

    process {
        try {
            # Check if SecretManagement is available
            if (-not (Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable)) {
                Write-SPOFactoryLog -Message "SecretManagement module not available, cannot retrieve stored credentials" -Level Warning -ClientName $ClientName -Category 'Security'
                return $null
            }

            # Check if vault exists
            $vault = Get-SecretVault -Name $VaultName -ErrorAction SilentlyContinue
            if (-not $vault) {
                Write-SPOFactoryLog -Message "Secret vault '$VaultName' not found" -Level Warning -ClientName $ClientName -Category 'Security'
                return $null
            }

            # Construct secret name
            $secretName = Get-SPOFactorySecretName -ClientName $ClientName -AuthType $AuthType

            # Retrieve credential from vault
            $credential = $null
            switch ($AuthType) {
                'Password' {
                    $credential = Get-Secret -Name $secretName -Vault $VaultName -AsPlainText:$false -ErrorAction SilentlyContinue
                    if ($credential -and $credential -is [PSCredential]) {
                        Write-SPOFactoryLog -Message "Retrieved password credential for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                    }
                }
                
                'Certificate' {
                    $certInfo = Get-Secret -Name $secretName -Vault $VaultName -AsPlainText:$true -ErrorAction SilentlyContinue
                    if ($certInfo) {
                        # Parse certificate information
                        $certData = $certInfo | ConvertFrom-Json
                        $credential = @{
                            Thumbprint = $certData.Thumbprint
                            ClientId = $certData.ClientId
                            TenantId = $certData.TenantId
                            Certificate = $certData.Certificate
                        }
                        Write-SPOFactoryLog -Message "Retrieved certificate credential for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                    }
                }
                
                'ClientSecret' {
                    $secretValue = Get-Secret -Name $secretName -Vault $VaultName -AsPlainText:$true -ErrorAction SilentlyContinue
                    if ($secretValue) {
                        $secretData = $secretValue | ConvertFrom-Json
                        $credential = @{
                            ClientId = $secretData.ClientId
                            ClientSecret = $secretData.ClientSecret
                            TenantId = $secretData.TenantId
                        }
                        Write-SPOFactoryLog -Message "Retrieved client secret credential for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                    }
                }
                
                'Token' {
                    $tokenValue = Get-Secret -Name $secretName -Vault $VaultName -AsPlainText:$true -ErrorAction SilentlyContinue
                    if ($tokenValue) {
                        $tokenData = $tokenValue | ConvertFrom-Json
                        $credential = @{
                            AccessToken = $tokenData.AccessToken
                            RefreshToken = $tokenData.RefreshToken
                            ExpiresAt = [DateTime]$tokenData.ExpiresAt
                            TokenType = $tokenData.TokenType
                        }
                        Write-SPOFactoryLog -Message "Retrieved token credential for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                    }
                }
            }

            if ($credential) {
                # Cache the credential if caching is enabled
                if (-not $NoCache) {
                    if (-not $script:SPOFactoryCredentialCache) {
                        $script:SPOFactoryCredentialCache = @{}
                    }
                    
                    $cacheKey = "$ClientName|$AuthType"
                    $cacheExpiry = (Get-Date).AddMinutes(30) # Cache for 30 minutes
                    
                    $script:SPOFactoryCredentialCache[$cacheKey] = @{
                        Credential = $credential
                        ExpiresAt = $cacheExpiry
                        RetrievedAt = Get-Date
                    }
                }

                Write-SPOFactoryLog -Message "Successfully retrieved credential for $ClientName" -Level Info -ClientName $ClientName -Category 'Security'
                return $credential
            } else {
                Write-SPOFactoryLog -Message "No credential found for $ClientName with auth type $AuthType" -Level Warning -ClientName $ClientName -Category 'Security'
                return $null
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve credential for $ClientName`: $_" -Level Error -ClientName $ClientName -Category 'Security' -Exception $_.Exception
            return $null
        }
    }
}

function Set-SPOFactoryCredential {
    <#
    .SYNOPSIS
        Stores credentials securely for SharePoint Online clients in MSP environments.

    .DESCRIPTION
        Securely stores credentials in the configured secret vault for MSP client tenants.
        Supports multiple authentication types with proper encryption and metadata.

    .PARAMETER ClientName
        The client name to store credentials for

    .PARAMETER Credential
        The credential object to store

    .PARAMETER AuthType
        Type of authentication credentials to store

    .PARAMETER Thumbprint
        Certificate thumbprint for certificate-based authentication

    .PARAMETER ClientId
        Azure AD Application (Client) ID

    .PARAMETER ClientSecret
        Azure AD Application Client Secret

    .PARAMETER TenantId
        Azure AD Tenant ID

    .PARAMETER VaultName
        Override the default secret vault name

    .PARAMETER Description
        Description for the stored credential

    .PARAMETER ExpiresAt
        Expiration date for the credential

    .EXAMPLE
        Set-SPOFactoryCredential -ClientName "Contoso Corp" -Credential $cred

    .EXAMPLE
        Set-SPOFactoryCredential -ClientName "Contoso Corp" -AuthType "Certificate" -Thumbprint "ABC123..." -ClientId "12345..." -TenantId "67890..."
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Password')]
        [PSCredential]$Credential,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Password', 'Certificate', 'ClientSecret', 'Token')]
        [string]$AuthType = 'Password',
        
        [Parameter(Mandatory = $false, ParameterSetName = 'Certificate')]
        [string]$Thumbprint,
        
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        
        [Parameter(Mandatory = $false, ParameterSetName = 'ClientSecret')]
        [string]$ClientSecret,
        
        [Parameter(Mandatory = $false)]
        [string]$TenantId,
        
        [Parameter(Mandatory = $false)]
        [string]$VaultName = $script:SPOFactoryConfig.CredentialVault,
        
        [Parameter(Mandatory = $false)]
        [string]$Description,
        
        [Parameter(Mandatory = $false)]
        [DateTime]$ExpiresAt
    )

    begin {
        Write-SPOFactoryLog -Message "Storing credentials for client: $ClientName" -Level Info -ClientName $ClientName -Category 'Security' -EnableAuditLog
    }

    process {
        try {
            # Validate SecretManagement availability
            if (-not (Get-Module -Name Microsoft.PowerShell.SecretManagement -ListAvailable)) {
                throw "SecretManagement module is not available"
            }

            # Validate vault exists
            $vault = Get-SecretVault -Name $VaultName -ErrorAction SilentlyContinue
            if (-not $vault) {
                throw "Secret vault '$VaultName' not found"
            }

            # Construct secret name
            $secretName = Get-SPOFactorySecretName -ClientName $ClientName -AuthType $AuthType

            if ($PSCmdlet.ShouldProcess("Secret '$secretName'", "Store credential in vault '$VaultName'")) {
                $secretValue = $null
                $metadata = @{
                    ClientName = $ClientName
                    AuthType = $AuthType
                    CreatedAt = Get-Date
                    CreatedBy = $env:USERNAME
                    ModuleVersion = $script:SPOFactoryConstants.ModuleVersion
                    Description = $Description
                }

                if ($ExpiresAt) {
                    $metadata.ExpiresAt = $ExpiresAt
                }

                switch ($AuthType) {
                    'Password' {
                        if (-not $Credential) {
                            throw "Credential parameter is required for Password auth type"
                        }
                        $secretValue = $Credential
                        $metadata.Username = $Credential.UserName
                    }
                    
                    'Certificate' {
                        if (-not $Thumbprint) {
                            throw "Thumbprint parameter is required for Certificate auth type"
                        }
                        
                        $certData = @{
                            Thumbprint = $Thumbprint
                            ClientId = $ClientId
                            TenantId = $TenantId
                        }
                        
                        $secretValue = ($certData | ConvertTo-Json)
                        $metadata.Thumbprint = $Thumbprint
                        $metadata.ClientId = $ClientId
                    }
                    
                    'ClientSecret' {
                        if (-not $ClientSecret -or -not $ClientId) {
                            throw "ClientSecret and ClientId parameters are required for ClientSecret auth type"
                        }
                        
                        $secretData = @{
                            ClientId = $ClientId
                            ClientSecret = $ClientSecret
                            TenantId = $TenantId
                        }
                        
                        $secretValue = ($secretData | ConvertTo-Json)
                        $metadata.ClientId = $ClientId
                    }
                    
                    'Token' {
                        # Token handling would be implemented based on specific requirements
                        throw "Token auth type storage is not yet implemented"
                    }
                }

                # Store the secret
                Set-Secret -Name $secretName -Secret $secretValue -Vault $VaultName -Metadata $metadata

                # Clear cached credential if it exists
                if ($script:SPOFactoryCredentialCache) {
                    $cacheKey = "$ClientName|$AuthType"
                    if ($script:SPOFactoryCredentialCache.ContainsKey($cacheKey)) {
                        $script:SPOFactoryCredentialCache.Remove($cacheKey)
                    }
                }

                Write-SPOFactoryLog -Message "Successfully stored credential for $ClientName (Type: $AuthType)" -Level Info -ClientName $ClientName -Category 'Security' -EnableAuditLog
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to store credential for $ClientName`: $_" -Level Error -ClientName $ClientName -Category 'Security' -Exception $_.Exception -EnableAuditLog
            throw
        }
    }
}

function Remove-SPOFactoryCredential {
    <#
    .SYNOPSIS
        Removes stored credentials for SharePoint Online clients.

    .DESCRIPTION
        Securely removes stored credentials from the secret vault and clears any cached credentials.

    .PARAMETER ClientName
        The client name to remove credentials for

    .PARAMETER AuthType
        Type of authentication credentials to remove

    .PARAMETER VaultName
        Override the default secret vault name

    .PARAMETER Force
        Force removal without confirmation

    .EXAMPLE
        Remove-SPOFactoryCredential -ClientName "Contoso Corp"

    .EXAMPLE
        Remove-SPOFactoryCredential -ClientName "Contoso Corp" -AuthType "Certificate" -Force
    #>

    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Password', 'Certificate', 'ClientSecret', 'Token', 'All')]
        [string]$AuthType = 'All',
        
        [Parameter(Mandatory = $false)]
        [string]$VaultName = $script:SPOFactoryConfig.CredentialVault,
        
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    begin {
        Write-SPOFactoryLog -Message "Removing credentials for client: $ClientName" -Level Warning -ClientName $ClientName -Category 'Security' -EnableAuditLog
    }

    process {
        try {
            # Determine which auth types to remove
            $authTypesToRemove = if ($AuthType -eq 'All') {
                @('Password', 'Certificate', 'ClientSecret', 'Token')
            } else {
                @($AuthType)
            }

            foreach ($currentAuthType in $authTypesToRemove) {
                $secretName = Get-SPOFactorySecretName -ClientName $ClientName -AuthType $currentAuthType

                if ($Force -or $PSCmdlet.ShouldProcess("Secret '$secretName'", "Remove credential from vault '$VaultName'")) {
                    try {
                        # Check if secret exists
                        $secret = Get-Secret -Name $secretName -Vault $VaultName -ErrorAction SilentlyContinue
                        if ($secret) {
                            Remove-Secret -Name $secretName -Vault $VaultName
                            Write-SPOFactoryLog -Message "Removed $currentAuthType credential for $ClientName" -Level Info -ClientName $ClientName -Category 'Security' -EnableAuditLog
                        }

                        # Clear from cache
                        if ($script:SPOFactoryCredentialCache) {
                            $cacheKey = "$ClientName|$currentAuthType"
                            if ($script:SPOFactoryCredentialCache.ContainsKey($cacheKey)) {
                                $script:SPOFactoryCredentialCache.Remove($cacheKey)
                            }
                        }
                    }
                    catch {
                        Write-SPOFactoryLog -Message "Failed to remove $currentAuthType credential for $ClientName`: $_" -Level Warning -ClientName $ClientName -Category 'Security'
                    }
                }
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Error removing credentials for $ClientName`: $_" -Level Error -ClientName $ClientName -Category 'Security' -Exception $_.Exception
            throw
        }
    }
}

function Get-SPOFactorySecret {
    <#
    .SYNOPSIS
        Retrieves specific secrets from the vault for MSP operations.

    .DESCRIPTION
        Helper function to retrieve specific types of secrets for MSP operations
        such as client secrets, API keys, and other sensitive configuration data.

    .PARAMETER ClientName
        The client name to retrieve secret for

    .PARAMETER SecretType
        Type of secret to retrieve

    .PARAMETER SecretName
        Specific secret name if not following standard naming

    .EXAMPLE
        Get-SPOFactorySecret -ClientName "Contoso Corp" -SecretType "APIKey"
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('APIKey', 'ClientSecret', 'WebhookSecret', 'EncryptionKey')]
        [string]$SecretType,
        
        [Parameter(Mandatory = $false)]
        [string]$SecretName
    )

    process {
        try {
            if (-not $SecretName) {
                $SecretName = "SPOFactory-$ClientName-$SecretType"
            }

            $secret = Get-Secret -Name $SecretName -Vault $script:SPOFactoryConfig.CredentialVault -AsPlainText -ErrorAction SilentlyContinue
            
            if ($secret) {
                Write-SPOFactoryLog -Message "Retrieved $SecretType secret for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                return $secret
            } else {
                Write-SPOFactoryLog -Message "No $SecretType secret found for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
                return $null
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to retrieve $SecretType secret for $ClientName`: $_" -Level Warning -ClientName $ClientName -Category 'Security'
            return $null
        }
    }
}

function Set-SPOFactorySecret {
    <#
    .SYNOPSIS
        Stores specific secrets in the vault for MSP operations.

    .DESCRIPTION
        Helper function to store specific types of secrets for MSP operations
        with proper naming and metadata.

    .PARAMETER ClientName
        The client name to store secret for

    .PARAMETER SecretType
        Type of secret to store

    .PARAMETER SecretValue
        The secret value to store

    .PARAMETER SecretName
        Specific secret name if not following standard naming

    .PARAMETER ExpiresAt
        When the secret expires

    .EXAMPLE
        Set-SPOFactorySecret -ClientName "Contoso Corp" -SecretType "APIKey" -SecretValue "secret123"
    #>

    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $true)]
        [ValidateSet('APIKey', 'ClientSecret', 'WebhookSecret', 'EncryptionKey')]
        [string]$SecretType,
        
        [Parameter(Mandatory = $true)]
        [string]$SecretValue,
        
        [Parameter(Mandatory = $false)]
        [string]$SecretName,
        
        [Parameter(Mandatory = $false)]
        [DateTime]$ExpiresAt
    )

    process {
        try {
            if (-not $SecretName) {
                $SecretName = "SPOFactory-$ClientName-$SecretType"
            }

            if ($PSCmdlet.ShouldProcess("Secret '$SecretName'", "Store $SecretType secret")) {
                $metadata = @{
                    ClientName = $ClientName
                    SecretType = $SecretType
                    CreatedAt = Get-Date
                    CreatedBy = $env:USERNAME
                    ModuleVersion = $script:SPOFactoryConstants.ModuleVersion
                }

                if ($ExpiresAt) {
                    $metadata.ExpiresAt = $ExpiresAt
                }

                Set-Secret -Name $SecretName -Secret $SecretValue -Vault $script:SPOFactoryConfig.CredentialVault -Metadata $metadata

                Write-SPOFactoryLog -Message "Stored $SecretType secret for $ClientName" -Level Info -ClientName $ClientName -Category 'Security' -EnableAuditLog
            }
        }
        catch {
            Write-SPOFactoryLog -Message "Failed to store $SecretType secret for $ClientName`: $_" -Level Error -ClientName $ClientName -Category 'Security' -Exception $_.Exception
            throw
        }
    }
}

function Get-SPOFactorySecretName {
    <#
    .SYNOPSIS
        Generates standardized secret names for MSP credential storage.

    .DESCRIPTION
        Creates consistent secret names based on client name and authentication type
        for organized credential management in MSP environments.

    .PARAMETER ClientName
        The client name

    .PARAMETER AuthType
        The authentication type

    .EXAMPLE
        Get-SPOFactorySecretName -ClientName "Contoso Corp" -AuthType "Password"
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $true)]
        [string]$AuthType
    )

    # Clean client name for use in secret name
    $cleanClientName = $ClientName -replace '[^\w\-]', '-'
    
    return "SPOFactory-$cleanClientName-$AuthType"
}

function Clear-SPOFactoryCredentialCache {
    <#
    .SYNOPSIS
        Clears cached credentials from memory.

    .DESCRIPTION
        Removes cached credentials to force fresh retrieval from the vault.
        Useful for security purposes or when credentials have been updated.

    .PARAMETER ClientName
        Specific client to clear cache for

    .PARAMETER AuthType
        Specific auth type to clear cache for

    .EXAMPLE
        Clear-SPOFactoryCredentialCache

    .EXAMPLE
        Clear-SPOFactoryCredentialCache -ClientName "Contoso Corp"
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientName,
        
        [Parameter(Mandatory = $false)]
        [string]$AuthType
    )

    process {
        if (-not $script:SPOFactoryCredentialCache) {
            return
        }

        if ($ClientName -and $AuthType) {
            $cacheKey = "$ClientName|$AuthType"
            if ($script:SPOFactoryCredentialCache.ContainsKey($cacheKey)) {
                $script:SPOFactoryCredentialCache.Remove($cacheKey)
                Write-SPOFactoryLog -Message "Cleared credential cache for $ClientName ($AuthType)" -Level Debug -ClientName $ClientName -Category 'Security'
            }
        }
        elseif ($ClientName) {
            $keysToRemove = $script:SPOFactoryCredentialCache.Keys | Where-Object { $_ -like "$ClientName|*" }
            foreach ($key in $keysToRemove) {
                $script:SPOFactoryCredentialCache.Remove($key)
            }
            Write-SPOFactoryLog -Message "Cleared all credential cache for $ClientName" -Level Debug -ClientName $ClientName -Category 'Security'
        }
        else {
            $script:SPOFactoryCredentialCache.Clear()
            Write-SPOFactoryLog -Message "Cleared all credential cache" -Level Debug -Category 'Security'
        }
    }
}