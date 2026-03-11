<#
.SYNOPSIS
    Exports comprehensive device information to CSV/JSON.

.DESCRIPTION
    This script connects to Microsoft Graph API and collects comprehensive device information
    including device details, compliance status, app installations, policy assignments, and
    other Intune-related information. The collected data is exported to CSV or JSON format.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.PARAMETER DeviceId
    Optional. The device ID to generate report for. If not specified, generates report for all devices.

.PARAMETER OutputPath
    Optional. Path to export report file. If not specified, exports to $env:TEMP\IntuneDeviceReport.csv.

.PARAMETER Format
    Optional. Export format: "CSV" or "JSON". Default is "CSV".

.EXAMPLE
    .\Export-IntuneDeviceReport.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Export-IntuneDeviceReport.ps1 -DeviceId 'device-guid' -Format 'JSON' -OutputPath 'C:\Reports\DeviceReport.json'

.NOTES
    Requires: Microsoft.Graph PowerShell module
    Requires: Entra ID App Registration with DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All, and DeviceManagementApps.Read.All permissions
#>

[CmdletBinding()]
# Define script parameters for credential input and export options
# These allow credentials and export parameters to be passed as command-line arguments
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
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("CSV", "JSON")]
    [string]$Format = "CSV"
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

# Determine output path if not specified
# If not specified, create a timestamped file in $env:TEMP
if (-not $OutputPath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $extension = if ($Format -eq "JSON") { "json" } else { "csv" }
    $OutputPath = Join-Path $env:TEMP "IntuneDeviceReport_$timestamp.$extension"
}

# Create output directory if it doesn't exist
# This ensures the directory exists before exporting the report
$exportDirectory = Split-Path -Path $OutputPath -Parent
if ($exportDirectory -and -not (Test-Path -Path $exportDirectory)) {
    New-Item -Path $exportDirectory -ItemType Directory -Force | Out-Null
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
    
    # Initialize array to store device reports
    # This will hold all device information retrieved from Graph API
    $deviceReports = @()
    
    # Determine which devices to query
    # If DeviceId is specified, query only that device
    # Otherwise, query all devices
    if ($DeviceId) {
        Write-Host "Generating report for device: $DeviceId" -ForegroundColor Yellow
        
        # Query device details for specific device
        # Get comprehensive device information
        $deviceUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$DeviceId')"
        $device = Invoke-GraphApiWithRetry -Uri $deviceUrl -Headers $headers -Method Get
        
        if ($null -eq $device) {
            Write-Error "Device not found or API returned null response for device ID: $DeviceId"
            exit 1
        }
        
        $devices = @($device)
    }
    else {
        Write-Host "Generating report for all devices..." -ForegroundColor Yellow
        
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
    
    Write-Host "Found $($devices.Count) device(s). Collecting information..." -ForegroundColor Cyan
    
    # Process each device and build comprehensive report
    # Collect all relevant information for each device
    $deviceCount = 0
    foreach ($device in $devices) {
        $deviceCount++
        Write-Host "Processing device $deviceCount of $($devices.Count): $($device.deviceName)" -ForegroundColor Cyan
        
        # Query device-specific compliance policy assignments
        # Get compliance policies assigned to this specific device
        $deviceCompliancePolicies = @()
        try {
            $deviceCompliancePoliciesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$($device.id)')/deviceCompliancePolicyStates"
            $deviceCompliancePoliciesResponse = Invoke-GraphApiWithRetry -Uri $deviceCompliancePoliciesUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue
            if ($deviceCompliancePoliciesResponse -and $deviceCompliancePoliciesResponse.value) {
                $deviceCompliancePolicies = $deviceCompliancePoliciesResponse.value
            }
        }
        catch {
            Write-Verbose "Could not retrieve compliance policy states for device $($device.id): $($_.Exception.Message)"
        }
        
        # Query device-specific configuration profile assignments
        # Get configuration profiles assigned to this specific device
        $deviceConfigProfiles = @()
        try {
            $deviceConfigProfilesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$($device.id)')/deviceConfigurationStates"
            $deviceConfigProfilesResponse = Invoke-GraphApiWithRetry -Uri $deviceConfigProfilesUrl -Headers $headers -Method Get -ErrorAction SilentlyContinue
            if ($deviceConfigProfilesResponse -and $deviceConfigProfilesResponse.value) {
                $deviceConfigProfiles = $deviceConfigProfilesResponse.value
            }
        }
        catch {
            Write-Verbose "Could not retrieve configuration profile states for device $($device.id): $($_.Exception.Message)"
        }
        
        # Create device report object with all available information
        $deviceReport = [PSCustomObject]@{
            DeviceId = $device.id
            DeviceName = $device.deviceName
            UserPrincipalName = $device.userPrincipalName
            UserDisplayName = $device.userDisplayName
            EnrollmentDateTime = $device.enrolledDateTime
            LastSyncDateTime = $device.lastSyncDateTime
            ComplianceState = $device.complianceState
            ManagementAgent = $device.managementAgent
            OperatingSystem = $device.operatingSystem
            OSVersion = $device.osVersion
            Manufacturer = $device.manufacturer
            Model = $device.model
            SerialNumber = $device.serialNumber
            TotalStorageSpaceInBytes = $device.totalStorageSpaceInBytes
            FreeStorageSpaceInBytes = $device.freeStorageSpaceInBytes
            ManagedDeviceName = $device.managedDeviceName
            DeviceEnrollmentType = $device.deviceEnrollmentType
            IsSupervised = $device.isSupervised
            IsEncrypted = $device.isEncrypted
            JailBroken = $device.jailBroken
            PhoneNumber = $device.phoneNumber
            IMEI = $device.imei
            EASDeviceId = $device.easDeviceId
            ExchangeAccessState = $device.exchangeAccessState
            ExchangeAccessStateReason = $device.exchangeAccessStateReason
            CompliancePolicies = ($deviceCompliancePolicies | ForEach-Object { $_.displayName }) -join '; '
            ConfigurationProfiles = ($deviceConfigProfiles | ForEach-Object { $_.displayName }) -join '; '
            ReportGeneratedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        }
        
        $deviceReports += $deviceReport
    }
    
    # Export report based on specified format
    # Export to CSV or JSON format as requested
    Write-Host "`nExporting report to $Format format..." -ForegroundColor Yellow
    
    if ($Format -eq "JSON") {
        # Export to JSON format
        # Convert device reports to JSON and save to file
        $deviceReports | ConvertTo-Json -Depth 10 | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Host "Report exported to: $OutputPath" -ForegroundColor Green
    }
    else {
        # Export to CSV format
        # Convert device reports to CSV and save to file
        $deviceReports | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: $OutputPath" -ForegroundColor Green
    }
    
    # Verify the export file was created successfully
    if (-not (Test-Path -Path $OutputPath)) {
        Write-Error "Export file was not created successfully at: $OutputPath"
        exit 1
    }
    
    # Verify the file has content
    $fileInfo = Get-Item -Path $OutputPath
    if ($fileInfo.Length -eq 0) {
        Write-Warning "Export file was created but appears to be empty."
    }
    
    # Display summary
    Write-Host "`nReport Summary:" -ForegroundColor Cyan
    Write-Host "===============" -ForegroundColor Cyan
    Write-Host "Devices included: $($deviceReports.Count)" -ForegroundColor White
    Write-Host "Output format: $Format" -ForegroundColor White
    Write-Host "Output path: $OutputPath" -ForegroundColor White
    
    Write-Host "`nReport generation completed successfully." -ForegroundColor Green
    exit 0
}
# Error handling block - catches exceptions during Graph API calls or export
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Export-IntuneDeviceReport: Failed to export device report - $errMsg"
    
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

