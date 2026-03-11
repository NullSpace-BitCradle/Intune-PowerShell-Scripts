#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Remediation script for Office update channel configuration.

.DESCRIPTION
    This script configures Microsoft Office to use the Semi-Annual update channel
    and updates it to the latest version. This is the remediation script that pairs
    with Detect-OfficeUpdateChannel.ps1. It should be run when Office is not on
    the Semi-Annual channel or not on the latest version.

.EXAMPLE
    .\Set-OfficeUpdateChannel.ps1

.EXAMPLE
    .\Set-OfficeUpdateChannel.ps1 -WhatIf

.NOTES
    Requires: Administrator privileges
    Requires: Office Click-to-Run installation
    Exit Codes:
    - 0: Remediation successful (Office configured to Semi-Annual channel)
    - 1: Remediation failed or error occurred
#>

[CmdletBinding(SupportsShouldProcess=$true)]
# Define script parameters
param()

# Main script execution wrapped in try-catch for error handling
try
{
    # Check if running as administrator
    # Office configuration changes require administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "This script requires administrator privileges. Please run as administrator."
        exit 1
    }

    # Check if Office Click-to-Run is installed
    # Verify the Office configuration registry path exists
    $officeConfigPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
    if (-not (Test-Path -Path $officeConfigPath)) {
        Write-Error "Office Click-to-Run is not installed or configuration path not found."
        exit 1
    }

    Write-Host "Starting Office update channel remediation..." -ForegroundColor Yellow

    # Function to invoke Office API with retry logic
    # Handles transient failures (429, 503, etc.) with exponential backoff
    function Invoke-OfficeApiWithRetry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory=$true)]
            [string]$Uri,
            
            [Parameter(Mandatory=$false)]
            [int]$MaxRetries = 3,
            
            [Parameter(Mandatory=$false)]
            [int]$InitialDelaySeconds = 2
        )
        
        $retryCount = 0
        $delaySeconds = $InitialDelaySeconds
        
        while ($retryCount -le $MaxRetries) {
            try {
                $response = Invoke-RestMethod -Uri $Uri -ErrorAction Stop
                return $response
            }
            catch {
                $statusCode = $_.Exception.Response.StatusCode.value__
                
                # Check if this is a retryable error (429 Too Many Requests, 503 Service Unavailable, 502 Bad Gateway, 504 Gateway Timeout)
                if ($statusCode -in @(429, 502, 503, 504) -and $retryCount -lt $MaxRetries) {
                    $retryCount++
                    Write-Verbose "Office API call failed with status $statusCode. Retrying in $delaySeconds seconds (attempt $retryCount of $MaxRetries)..."
                    Start-Sleep -Seconds $delaySeconds
                    # Exponential backoff: double the delay for next retry
                    $delaySeconds = $delaySeconds * 2
                }
                else {
                    # Not a retryable error or max retries reached
                    throw
                }
            }
        }
    }
    
    # Query Microsoft's Office release information service to get the latest available versions
    # This provides up-to-date information about all Office update channels and their latest versions
    Write-Host "Querying Microsoft Office release information..." -ForegroundColor Yellow
    $CloudVersionInfo = Invoke-OfficeApiWithRetry -Uri 'https://clients.config.office.net/releases/v1.0/OfficeReleases'
    
    # Find the Semi-Annual channel information
    # Filter the cloud version information to find the Semi-Annual channel
    $semiAnnualChannel = $CloudVersionInfo | Where-Object { $_.channelId -eq "SemiAnnual" }
    
    if ($null -eq $semiAnnualChannel) {
        Write-Error "Could not find Semi-Annual channel information from Microsoft."
        exit 1
    }

    # Get the Semi-Annual channel CDN base URL
    # This URL is used to configure Office to use the Semi-Annual update channel
    $semiAnnualCDN = $semiAnnualChannel.OfficeVersions.cdnBaseURL
    $latestVersion = $semiAnnualChannel.latestversion

    Write-Host "Semi-Annual channel latest version: $latestVersion" -ForegroundColor Cyan
    Write-Host "Semi-Annual channel CDN: $semiAnnualCDN" -ForegroundColor Cyan

    # Read the current Office configuration from the registry
    # Get the current CDN base URL to see what channel is currently configured
    $currentChannel = Get-ItemPropertyValue -Path $officeConfigPath -Name "CDNBaseUrl" -ErrorAction SilentlyContinue
    
    # Check if Office is already configured for Semi-Annual channel
    if ($currentChannel -eq $semiAnnualCDN) {
        Write-Host "Office is already configured for Semi-Annual channel." -ForegroundColor Green
        
        # Check if Office is on the latest version
        $currentVersion = Get-ItemPropertyValue -Path $officeConfigPath -Name "VersionToReport" -ErrorAction SilentlyContinue
        if ($currentVersion -eq $latestVersion) {
            Write-Host "Office is already on the latest version ($latestVersion)." -ForegroundColor Green
            exit 0
        }
        else {
            Write-Host "Office is on version $currentVersion but latest is $latestVersion. Updating..." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "Current channel: $currentChannel" -ForegroundColor Yellow
        
        if ($PSCmdlet.ShouldProcess("Office CDNBaseUrl registry value", "Configure Office to use Semi-Annual channel ($semiAnnualCDN)")) {
            Write-Host "Configuring Office to use Semi-Annual channel..." -ForegroundColor Yellow

            # Set the CDN base URL to Semi-Annual channel
            # This configures Office to use the Semi-Annual update channel
            Set-ItemProperty -Path $officeConfigPath -Name "CDNBaseUrl" -Value $semiAnnualCDN -Force -ErrorAction Stop
            Write-Host "Office update channel set to Semi-Annual." -ForegroundColor Green
        }
    }

    # Trigger Office update to get the latest version
    # Use Office Click-to-Run update mechanism to update Office
    if ($PSCmdlet.ShouldProcess("Office installation", "Trigger update to latest version ($latestVersion)")) {
        Write-Host "Triggering Office update to latest version..." -ForegroundColor Yellow

        # Check if OfficeClickToRun.exe exists (Office update executable)
        $officeUpdatePath = "${env:ProgramFiles}\Microsoft Office\Office16\OfficeClickToRun.exe"
        if (-not (Test-Path -Path $officeUpdatePath)) {
            # Try alternative path for 64-bit Office on 32-bit system
            $officeUpdatePath = "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OfficeClickToRun.exe"
        }

        if (Test-Path -Path $officeUpdatePath) {
            # Run Office update in the background
            # /update user: This updates Office for the current user
            # /update user displaylevel=false: Runs update silently
            Write-Host "Starting Office update process..." -ForegroundColor Yellow
            Start-Process -FilePath $officeUpdatePath -ArgumentList "/update", "user", "displaylevel=false" -NoNewWindow -Wait -ErrorAction Stop
            Write-Host "Office update process completed." -ForegroundColor Green
        }
        else {
            Write-Warning "OfficeClickToRun.exe not found. Office may update automatically on next check."
        }
    }

    # Verify the remediation was successful
    # Wait a moment for registry values to update
    Start-Sleep -Seconds 5
    
    # Check the current configuration
    $finalChannel = Get-ItemPropertyValue -Path $officeConfigPath -Name "CDNBaseUrl" -ErrorAction SilentlyContinue
    $finalVersion = Get-ItemPropertyValue -Path $officeConfigPath -Name "VersionToReport" -ErrorAction SilentlyContinue

    if ($finalChannel -eq $semiAnnualCDN) {
        Write-Host "Office is now configured for Semi-Annual channel." -ForegroundColor Green
        Write-Host "Current version: $finalVersion" -ForegroundColor Cyan
        Write-Host "Latest version: $latestVersion" -ForegroundColor Cyan
        
        # Note: Version may not update immediately, but channel is configured correctly
        if ($finalVersion -eq $latestVersion) {
            Write-Host "Office is on the latest version." -ForegroundColor Green
        }
        else {
            Write-Host "Office update is in progress. Version will update to $latestVersion when complete." -ForegroundColor Yellow
        }
        
        exit 0
    }
    else {
        Write-Error "Remediation failed. Office channel is still set to: $finalChannel"
        exit 1
    }
}
# Error handling block - catches exceptions during registry modifications or API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Set-OfficeUpdateChannel: Failed to set Office update channel - $errMsg"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    # Exit with code 1 to indicate an error occurred
    exit 1
}

