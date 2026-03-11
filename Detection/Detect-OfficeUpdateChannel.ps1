<#
.SYNOPSIS
    Detection script for Office update channel and version verification.

.DESCRIPTION
    This script verifies if Microsoft Office is using the Semi-Annual update channel
    and is on the latest version. It reads the local Office configuration from the registry
    and compares it with the latest available version from Microsoft's cloud service.

.EXAMPLE
    .\Detect-OfficeUpdateChannel.ps1

.NOTES
    Exit Codes:
    - 0: Office is on Semi-Annual channel and latest version
    - 1: Office is not on Semi-Annual channel or not on latest version, or error occurred
#>

[CmdletBinding()]
param()

# Main script execution wrapped in try-catch for error handling
try 
{
    # Read the currently installed Office version from the registry
    # This value is reported by Office Click-to-Run installation
    $ReportedVersion = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "VersionToReport"
    
    # Read the CDN (Content Delivery Network) base URL from the registry
    # This URL indicates which update channel Office is configured to use
    # Select-Object -Last 1 ensures we get the most recent value if multiple exist
    $Channel = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -Name "CDNBaseUrl" | Select-Object -Last 1
    
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
    $CloudVersionInfo = Invoke-OfficeApiWithRetry -Uri 'https://clients.config.office.net/releases/v1.0/OfficeReleases'
    
    # Filter the cloud version information to find the channel that matches the local CDN URL
    # This identifies which update channel the local Office installation is using
    $UsedChannel = $CloudVersionInfo | Where-Object { $_.OfficeVersions.cdnBaseURL -eq $Channel }
    
    # Check if channel was found in cloud version information
    if ($null -eq $UsedChannel) {
        Write-Error "Could not determine Office update channel from cloud service. Channel URL: $Channel"
        exit 1
    }
    
    # Handle array results - ensure we have a single channel result
    if ($UsedChannel -is [Array]) {
        if ($UsedChannel.Count -gt 1) {
            Write-Warning "Multiple channels found matching CDN URL. Using first result."
        }
        $UsedChannel = $UsedChannel[0]
    }
    
    # Check if the detected channel is Semi-Annual and if the installed version matches the latest version
    # Both conditions must be true for the check to pass
    if (($UsedChannel.channelId -eq "SemiAnnual") -and ($UsedChannel.latestversion -eq $ReportedVersion)) {
        # Office is on Semi-Annual channel and latest version - compliance check passes
        Write-Host "Currently using the latest version of Office in the $($UsedChannel.channelId) Channel: $($ReportedVersion)"
        exit 0
    }
    else {
        # Office is either not on Semi-Annual channel or not on the latest version
        # Exit code 1 indicates non-compliance
        Write-Host "Not using Semi-Annual channel. Detected channel is the $($UsedChannel.channelId) Channel."
        exit 1   
    }
}
# Error handling block - catches exceptions during registry reads or API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Detect-OfficeUpdateChannel: Failed to detect Office update channel - $errMsg"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    # Exit with code 1 to indicate an error occurred
    exit 1
}
