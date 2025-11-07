<#
.SYNOPSIS
    Common functions module for Intune PowerShell scripts.

.DESCRIPTION
    This module provides shared functions used across multiple Intune PowerShell scripts,
    including Graph API authentication, retry logic, and common utilities.

.NOTES
    Version: 1.0
    Author: Intune Scripts Collection
#>

# Function to invoke Graph API with retry logic
# Handles transient failures (429, 503, etc.) with exponential backoff
function Invoke-GraphApiWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Uri,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory=$false)]
        [string]$Method = "Get",
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$InitialDelaySeconds = 2
    )
    
    $retryCount = 0
    $delaySeconds = $InitialDelaySeconds
    
    while ($retryCount -le $MaxRetries) {
        try {
            $response = Invoke-RestMethod -Uri $Uri -Headers $Headers -Method $Method -ErrorAction Stop
            return $response
        }
        catch {
            $statusCode = $_.Exception.Response.StatusCode.value__
            
            # Check if this is a retryable error (429 Too Many Requests, 503 Service Unavailable, 502 Bad Gateway, 504 Gateway Timeout)
            if ($statusCode -in @(429, 502, 503, 504) -and $retryCount -lt $MaxRetries) {
                $retryCount++
                
                # Check for Retry-After header (429 Too Many Requests)
                if ($statusCode -eq 429 -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfter = $_.Exception.Response.Headers['Retry-After']
                    if ($retryAfter -match '^\d+$') {
                        $delaySeconds = [int]$retryAfter
                        Write-Verbose "Graph API call rate limited. Using Retry-After header value: $delaySeconds seconds (attempt $retryCount of $MaxRetries)..."
                    }
                    else {
                        Write-Verbose "Graph API call failed with status $statusCode. Retrying in $delaySeconds seconds (attempt $retryCount of $MaxRetries)..."
                    }
                }
                else {
                    Write-Verbose "Graph API call failed with status $statusCode. Retrying in $delaySeconds seconds (attempt $retryCount of $MaxRetries)..."
                }
                
                Start-Sleep -Seconds $delaySeconds
                # Exponential backoff: double the delay for next retry (unless Retry-After was used)
                if ($statusCode -ne 429 -or -not $_.Exception.Response.Headers['Retry-After']) {
                    $delaySeconds = $delaySeconds * 2
                }
            }
            else {
                # Not a retryable error or max retries reached
                throw
            }
        }
    }
}

# Function to get all pages from a paginated Graph API response
function Get-GraphApiAllPages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Uri,
        
        [Parameter(Mandatory=$true)]
        [hashtable]$Headers,
        
        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,
        
        [Parameter(Mandatory=$false)]
        [int]$InitialDelaySeconds = 2
    )
    
    $allItems = @()
    $currentUri = $Uri
    $pageCount = 1
    
    do {
        Write-Progress -Activity "Retrieving paginated results" -Status "Fetching page $pageCount" -PercentComplete -1
        
        $response = Invoke-GraphApiWithRetry -Uri $currentUri -Headers $Headers -Method Get -MaxRetries $MaxRetries -InitialDelaySeconds $InitialDelaySeconds
        
        if ($null -ne $response -and $null -ne $response.value) {
            $allItems += $response.value
        }
        
        if ($response.'@odata.nextLink') {
            $currentUri = $response.'@odata.nextLink'
            $pageCount++
        }
        else {
            $currentUri = $null
        }
    } while ($null -ne $currentUri)
    
    Write-Progress -Activity "Retrieving paginated results" -Completed
    
    return $allItems
}

# Function to convert single object or array to array
function ConvertTo-Array {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline=$true, Mandatory=$true)]
        $InputObject
    )
    
    if ($null -eq $InputObject) { 
        return @() 
    }
    if ($InputObject -is [Array]) { 
        return $InputObject 
    }
    return @($InputObject)
}

# Function to get Graph API access token
function Get-GraphAccessToken {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$AppId,
        
        [Parameter(Mandatory=$true)]
        [string]$TenantId,
        
        [Parameter(Mandatory=$true)]
        [string]$ClientSecret
    )
    
    try {
        $body = @{
            grant_type    = "client_credentials"
            scope         = "https://graph.microsoft.com/.default"
            client_id     = $AppId
            client_secret = $ClientSecret
        }
        
        $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method Post -Body $body
        return $tokenResponse.access_token
    }
    catch {
        Write-Error "Failed to obtain access token: $($_.Exception.Message)"
        throw
    }
}

# Function to validate GUID format
function Test-GuidFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Guid
    )
    
    return $Guid -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
}

# Function to validate email format
function Test-EmailFormat {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Email
    )
    
    return $Email -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
}

# Function to sanitize input for OData queries
function Remove-ODataInjectionChars {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputString
    )
    
    # Remove potentially dangerous characters from OData filter
    return $InputString -replace "[';]", ""
}

# Function to validate and sanitize OData filter input
function Test-ODataFilterInput {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputString,
        
        [Parameter(Mandatory=$false)]
        [switch]$AllowWildcards
    )
    
    # Remove potentially dangerous characters from OData filter
    $sanitized = $InputString -replace "[';]", ""
    
    if (-not $AllowWildcards) {
        $sanitized = $sanitized -replace "[*%]", ""
    }
    
    if ([string]::IsNullOrWhiteSpace($sanitized)) {
        throw "Input string cannot be empty after sanitization."
    }
    
    return $sanitized
}

# Function to get Intune log path
function Get-IntuneLogPath {
    [CmdletBinding()]
    param()
    
    $logPath = Join-Path $env:ProgramData "Microsoft\IntuneManagementExtension\Logs"
    
    if (-not (Test-Path -Path $logPath)) {
        try {
            New-Item -Path $logPath -ItemType Directory -Force | Out-Null
        }
        catch {
            Write-Warning "Could not create log directory: $logPath"
        }
    }
    
    return $logPath
}

# Function to write standardized error message
function Write-StandardError {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$ScriptName,
        
        [Parameter(Mandatory=$true)]
        [string]$Operation,
        
        [Parameter(Mandatory=$false)]
        [string]$ObjectName = "",
        
        [Parameter(Mandatory=$false)]
        [Exception]$Exception
    )
    
    $errorMessage = "$ScriptName`: $Operation"
    if ($ObjectName) {
        $errorMessage += " on object '$ObjectName'"
    }
    
    if ($Exception) {
        $errorMessage += " - $($Exception.Message)"
    }
    
    Write-Error $errorMessage
    
    if ($Exception -and $PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($Exception.ScriptStackTrace)"
        if ($Exception.InnerException) {
            Write-Verbose "Inner Exception: $($Exception.InnerException.Message)"
        }
    }
}

# Export module members
Export-ModuleMember -Function Invoke-GraphApiWithRetry, Get-GraphAccessToken, Test-GuidFormat, Test-EmailFormat, Remove-ODataInjectionChars, Get-GraphApiAllPages, ConvertTo-Array, Test-ODataFilterInput, Get-IntuneLogPath, Write-StandardError

