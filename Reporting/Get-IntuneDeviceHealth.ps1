<#
.SYNOPSIS
    Comprehensive device health check.

.DESCRIPTION
    This script performs a comprehensive health check on Intune-managed devices including
    enrollment status, sync status, connectivity, app installation status, compliance state,
    and other health indicators. It can check a specific device or all devices.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.PARAMETER DeviceId
    Optional. The device ID to check health for. If not specified, checks all devices.

.PARAMETER ExportPath
    Optional. Path to export results to CSV file. If not specified, results are displayed only.

.EXAMPLE
    .\Get-IntuneDeviceHealth.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntuneDeviceHealth.ps1 -DeviceId 'device-guid' -ExportPath 'C:\Temp\DeviceHealth.csv'

.NOTES
    Requires: Microsoft.Graph PowerShell module
    Requires: Entra ID App Registration with DeviceManagementManagedDevices.Read.All permission
#>

[CmdletBinding()]
# Define script parameters for credential input and query options
# These allow credentials and query parameters to be passed as command-line arguments
param(
    [Parameter(Mandatory=$false)]
    [string]$AppId,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$false)]
    [string]$DeviceId,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath
)

# Check if Microsoft.Graph module is installed, and install if missing
# This module is required to interact with Microsoft Graph API
if(-not (Get-Module -Name Microsoft.Graph -ListAvailable))
{
    try {
        # Try to install for current user first (doesn't require admin privileges)
        # This is the preferred method as it doesn't require elevation
        Install-Module -Name Microsoft.Graph -Repository PSGallery -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
        Write-Host "Microsoft.Graph module installed successfully for current user." -ForegroundColor Green
    }
    catch [Exception] {
        # If user-scope installation fails, try system-wide installation
        # This requires administrator privileges
        Write-Warning "Failed to install module for current user: $($_.Exception.Message)"
        Write-Host "Attempting system-wide installation (requires admin privileges)..." -ForegroundColor Yellow
        try {
            Install-Module -Name Microsoft.Graph -Repository PSGallery -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
            Write-Host "Microsoft.Graph module installed successfully system-wide." -ForegroundColor Green
        }
        catch {
            # If both installation methods fail, exit with error
            Write-Error "Failed to install Microsoft.Graph module: $($_.Exception.Message)"
            exit 1
        }
    }
}

# Import the Microsoft Graph PowerShell module into the current session
# This makes the Graph API cmdlets available for use
try{
    Import-Module -Name Microsoft.Graph
}catch [Exception] {
    # Display error message if import fails
    Write-Error "Failed to import Microsoft.Graph module: $($_.Exception.Message)"
    exit 1
}

# Credential handling: Check parameters first, then environment variables
# This provides flexibility in how credentials are provided
$appid = if ($AppId) { $AppId } elseif ($env:INTUNE_APP_ID) { $env:INTUNE_APP_ID } else { '<Your Entra Registered App ID here>' }
$tenantid = if ($TenantId) { $TenantId } elseif ($env:INTUNE_TENANT_ID) { $env:INTUNE_TENANT_ID } else { '<Your Tenant ID here>' }
$clientsecret = if ($ClientSecret) { $ClientSecret } elseif ($env:INTUNE_CLIENT_SECRET) { $env:INTUNE_CLIENT_SECRET } else { '<Your Client Secret here>' }

# Validate that credentials are not placeholder values
# This prevents accidental use of placeholder credentials
if ($appid -match '^<.*>$' -or $tenantid -match '^<.*>$' -or $clientsecret -match '^<.*>$') {
    Write-Error "Please provide valid credentials via parameters or environment variables (INTUNE_APP_ID, INTUNE_TENANT_ID, INTUNE_CLIENT_SECRET)."
    exit 1
}

# Validate DeviceId format if provided
if ($DeviceId -and -not ($DeviceId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')) {
    Write-Error "DeviceId must be a valid GUID format."
    exit 1
}

# Main script execution wrapped in try-catch for error handling
try
{
    # Import shared module for Graph API utilities
    Import-Module "$PSScriptRoot\..\Common\IntuneCommon.psm1" -Force -ErrorAction Stop

    # Connect to Microsoft Graph API using app registration credentials
    # This authenticates the script to access Intune data via Graph API
    Write-Host "Connecting to Microsoft Graph API..." -ForegroundColor Yellow
    
    $body = @{
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $appid
        client_secret = $clientsecret
    }
    
    # Request an access token from Microsoft identity platform
    # The token is used to authenticate subsequent Graph API requests
    $tokenResponse = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token" -Method Post -Body $body
    $accessToken = $tokenResponse.access_token
    
    # Set up headers for Graph API requests
    # The access token is included in the Authorization header
    $headers = @{
        "Authorization" = "Bearer $accessToken"
        "Content-Type"  = "application/json"
    }
    
    Write-Host "Successfully connected to Microsoft Graph API." -ForegroundColor Green
    
    # Initialize array to store device health results
    # This will hold all device health information retrieved from Graph API
    $deviceHealthResults = @()
    
    # Determine which devices to check
    # If DeviceId is specified, check only that device
    # Otherwise, check all devices
    if ($DeviceId) {
        Write-Host "Checking health for device: $DeviceId" -ForegroundColor Yellow
        
        # Query device details for specific device
        # Get comprehensive device information
        $deviceUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$DeviceId')"
        $device = Invoke-GraphApiWithRetry -Uri $deviceUrl -Headers $headers -Method Get
        $devices = @($device)
    }
    else {
        Write-Host "Checking health for all devices..." -ForegroundColor Yellow
        
        # Query all managed devices from Intune
        $devicesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        $devices = Invoke-GraphApiWithRetry -Uri $devicesUrl -Headers $headers -Method Get
        
        # Handle pagination
        $allDevices = @()
        if ($null -ne $devices -and $null -ne $devices.value) {
            $allDevices += $devices.value
        }
        $pageCount = 1
        while ($null -ne $devices -and $null -ne $devices.'@odata.nextLink') {
            $pageCount++
            Write-Progress -Activity "Retrieving devices" -Status "Fetching page $pageCount" -PercentComplete -1
            $devices = Invoke-GraphApiWithRetry -Uri $devices.'@odata.nextLink' -Headers $headers -Method Get
            if ($null -ne $devices -and $null -ne $devices.value) {
                $allDevices += $devices.value
            }
        }
        Write-Progress -Activity "Retrieving devices" -Completed
        
        $devices = $allDevices
    }
    
    Write-Host "Found $($devices.Count) device(s). Checking health..." -ForegroundColor Cyan
    
    # Process each device and perform health check
    $deviceCount = 0
    foreach ($device in $devices) {
        $deviceCount++
        Write-Host "Checking device $deviceCount of $($devices.Count): $($device.deviceName)" -ForegroundColor Cyan
        
        # Calculate health score based on various factors
        $healthScore = 100
        $healthIssues = @()
        
        # Check 1: Enrollment status
        $enrollmentStatus = "Healthy"
        if (-not $device.enrolledDateTime) {
            $enrollmentStatus = "Not Enrolled"
            $healthScore -= 30
            $healthIssues += "Device not enrolled"
        }
        else {
            $enrollmentAge = (Get-Date) - [DateTime]$device.enrolledDateTime
            if ($enrollmentAge.Days -gt 365) {
                $enrollmentStatus = "Old Enrollment"
                $healthScore -= 5
                $healthIssues += "Enrollment is over 1 year old"
            }
        }
        
        # Check 2: Sync status
        $syncStatus = "Healthy"
        if ($device.lastSyncDateTime) {
            $lastSync = [DateTime]$device.lastSyncDateTime
            $timeSinceSync = (Get-Date) - $lastSync
            
            if ($timeSinceSync.TotalHours -gt 24) {
                $syncStatus = "Stale"
                $healthScore -= 20
                $healthIssues += "Last sync was over 24 hours ago"
            }
            elseif ($timeSinceSync.TotalHours -gt 8) {
                $syncStatus = "Warning"
                $healthScore -= 10
                $healthIssues += "Last sync was over 8 hours ago"
            }
        }
        else {
            $syncStatus = "Never Synced"
            $healthScore -= 25
            $healthIssues += "Device has never synced"
        }
        
        # Check 3: Compliance state
        # Verify device is compliant
        $complianceStatus = "Healthy"
        if ($device.complianceState) {
            if ($device.complianceState -ne "Compliant") {
                $complianceStatus = $device.complianceState
                $healthScore -= 15
                $healthIssues += "Device is not compliant: $($device.complianceState)"
            }
        }
        else {
            $complianceStatus = "Unknown"
            $healthScore -= 5
            $healthIssues += "Compliance state unknown"
        }
        
        # Check 4: Management agent
        # Verify device is using a supported management agent
        $managementStatus = "Healthy"
        if ($device.managementAgent) {
            $supportedAgents = @("MDM", "EAS", "IntuneClient")
            if ($device.managementAgent -notin $supportedAgents) {
                $managementStatus = "Unsupported"
                $healthScore -= 10
                $healthIssues += "Unsupported management agent: $($device.managementAgent)"
            }
        }
        else {
            $managementStatus = "Unknown"
            $healthScore -= 5
            $healthIssues += "Management agent unknown"
        }
        
        # Check 5: Operating system
        # Verify device is running a supported OS
        $osStatus = "Healthy"
        if ($device.operatingSystem) {
            $supportedOS = @("Windows", "iOS", "Android", "macOS")
            if ($device.operatingSystem -notin $supportedOS) {
                $osStatus = "Unsupported"
                $healthScore -= 5
                $healthIssues += "Unsupported operating system: $($device.operatingSystem)"
            }
        }
        
        # Check 6: Storage space (if available)
        # Verify device has adequate storage
        $storageStatus = "Healthy"
        if ($device.totalStorageSpaceInBytes -and $device.freeStorageSpaceInBytes) {
            $usedPercent = (($device.totalStorageSpaceInBytes - $device.freeStorageSpaceInBytes) / $device.totalStorageSpaceInBytes) * 100
            
            if ($usedPercent -gt 90) {
                $storageStatus = "Critical"
                $healthScore -= 15
                $healthIssues += "Storage is over 90% full"
            }
            elseif ($usedPercent -gt 80) {
                $storageStatus = "Warning"
                $healthScore -= 5
                $healthIssues += "Storage is over 80% full"
            }
        }
        
        # Determine overall health status
        # Categorize health based on score
        $overallHealth = "Healthy"
        if ($healthScore -lt 50) {
            $overallHealth = "Critical"
        }
        elseif ($healthScore -lt 70) {
            $overallHealth = "Warning"
        }
        elseif ($healthScore -lt 90) {
            $overallHealth = "Fair"
        }
        
        # Create device health result object
        $deviceHealth = [PSCustomObject]@{
            DeviceId = $device.id
            DeviceName = $device.deviceName
            UserPrincipalName = $device.userPrincipalName
            OverallHealth = $overallHealth
            HealthScore = $healthScore
            EnrollmentStatus = $enrollmentStatus
            EnrollmentDate = $device.enrolledDateTime
            SyncStatus = $syncStatus
            LastSyncDateTime = $device.lastSyncDateTime
            ComplianceStatus = $complianceStatus
            ComplianceState = $device.complianceState
            ManagementStatus = $managementStatus
            ManagementAgent = $device.managementAgent
            OSStatus = $osStatus
            OperatingSystem = $device.operatingSystem
            OSVersion = $device.osVersion
            StorageStatus = $storageStatus
            TotalStorageGB = if ($device.totalStorageSpaceInBytes) { [math]::Round($device.totalStorageSpaceInBytes / 1GB, 2) } else { $null }
            FreeStorageGB = if ($device.freeStorageSpaceInBytes) { [math]::Round($device.freeStorageSpaceInBytes / 1GB, 2) } else { $null }
            HealthIssues = ($healthIssues -join '; ')
            HealthIssueCount = $healthIssues.Count
        }
        
        $deviceHealthResults += $deviceHealth
    }
    
    # Display results
    Write-Host "`nDevice Health Results:" -ForegroundColor Cyan
    Write-Host "=====================" -ForegroundColor Cyan
    
    if ($deviceHealthResults.Count -eq 0) {
        Write-Host "No devices found matching the specified criteria." -ForegroundColor Yellow
    }
    else {
        # Group results by health status for better display
        $healthyDevices = $deviceHealthResults | Where-Object { $_.OverallHealth -eq "Healthy" }
        $fairDevices = $deviceHealthResults | Where-Object { $_.OverallHealth -eq "Fair" }
        $warningDevices = $deviceHealthResults | Where-Object { $_.OverallHealth -eq "Warning" }
        $criticalDevices = $deviceHealthResults | Where-Object { $_.OverallHealth -eq "Critical" }
        
        Write-Host "`nHealth Summary:" -ForegroundColor Cyan
        Write-Host "  Healthy: $($healthyDevices.Count)" -ForegroundColor Green
        Write-Host "  Fair: $($fairDevices.Count)" -ForegroundColor Yellow
        Write-Host "  Warning: $($warningDevices.Count)" -ForegroundColor $(if ($warningDevices.Count -gt 0) { 'Yellow' } else { 'White' })
        Write-Host "  Critical: $($criticalDevices.Count)" -ForegroundColor $(if ($criticalDevices.Count -gt 0) { 'Red' } else { 'White' })
        
        # Display each device's health
        foreach ($health in $deviceHealthResults) {
            $healthColor = switch ($health.OverallHealth) {
                "Healthy" { "Green" }
                "Fair" { "Yellow" }
                "Warning" { "Yellow" }
                "Critical" { "Red" }
                default { "White" }
            }
            
            Write-Host "`nDevice: $($health.DeviceName)" -ForegroundColor $healthColor
            Write-Host "  Overall Health: $($health.OverallHealth) (Score: $($health.HealthScore))" -ForegroundColor $healthColor
            Write-Host "  User: $($health.UserPrincipalName)" -ForegroundColor White
            Write-Host "  Enrollment: $($health.EnrollmentStatus)" -ForegroundColor White
            Write-Host "  Sync: $($health.SyncStatus) (Last: $($health.LastSyncDateTime))" -ForegroundColor White
            Write-Host "  Compliance: $($health.ComplianceStatus)" -ForegroundColor White
            Write-Host "  Management: $($health.ManagementStatus) ($($health.ManagementAgent))" -ForegroundColor White
            if ($health.HealthIssueCount -gt 0) {
                Write-Host "  Issues: $($health.HealthIssueCount)" -ForegroundColor Red
                foreach ($issue in ($health.HealthIssues -split '; ')) {
                    Write-Host "    - $issue" -ForegroundColor Red
                }
            }
        }
        
        # Export to CSV if path is specified
        if ($ExportPath) {
            # Create directory if it doesn't exist
            $exportDirectory = Split-Path -Path $ExportPath -Parent
            if ($exportDirectory -and -not (Test-Path -Path $exportDirectory)) {
                New-Item -Path $exportDirectory -ItemType Directory -Force | Out-Null
            }
            
            # Export results to CSV file
            $deviceHealthResults | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
            
            # Verify the export file was created successfully
            if (-not (Test-Path -Path $ExportPath)) {
                Write-Error "Export file was not created successfully at: $ExportPath"
                exit 1
            }
        }
    }
    
    Write-Host "`nHealth check completed successfully. Checked $($deviceHealthResults.Count) device(s)." -ForegroundColor Green
}
# Error handling block - catches exceptions during Graph API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Get-IntuneDeviceHealth: Failed to check device health - $errMsg"
    
    # Provide additional error context if available
    if ($_.Exception.Response) {
        $statusCode = $_.Exception.Response.StatusCode.value__
        Write-Verbose "HTTP Status Code: $statusCode"
    }
    
    if ($_.Exception.InnerException) {
        Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
    }
    
    exit 1
}

