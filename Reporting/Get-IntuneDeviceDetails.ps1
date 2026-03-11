<#
.SYNOPSIS
    Retrieves detailed information about Intune-managed devices.

.DESCRIPTION
    This script connects to Microsoft Graph API and retrieves comprehensive information
    about Intune-managed devices including enrollment details, last sync time, management
    state, hardware information, and device configuration. It can query a specific device
    by ID or retrieve all devices.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.PARAMETER DeviceId
    Optional. The device ID to query. If not specified, retrieves all devices.

.PARAMETER DeviceName
    Optional. The device name to search for. Partial matches are supported.

.PARAMETER ExportPath
    Optional. Path to export results to CSV file. If not specified, results are displayed only.

.EXAMPLE
    .\Get-IntuneDeviceDetails.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntuneDeviceDetails.ps1 -DeviceId 'device-guid' -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntuneDeviceDetails.ps1 -DeviceName 'LAPTOP-01' -ExportPath 'C:\Temp\DeviceDetails.csv'

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
    [string]$DeviceName,
    
    [Parameter(Mandatory=$false)]
    [string]$ExportPath,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 10)]
    [int]$MaxRetries = 3,
    
    [Parameter(Mandatory=$false)]
    [ValidateRange(1, 60)]
    [int]$InitialDelaySeconds = 2
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
if ($DeviceId) {
    if (-not ($DeviceId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')) {
        Write-Error "DeviceId must be a valid GUID format."
        exit 1
    }
}

# Sanitize DeviceName input for OData query
if ($DeviceName) {
    # Remove potentially dangerous characters from OData filter
    $DeviceName = $DeviceName -replace "[';]", ""
    if ([string]::IsNullOrWhiteSpace($DeviceName)) {
        Write-Error "DeviceName cannot be empty after sanitization."
        exit 1
    }
}

# Validate ExportPath if provided
if ($ExportPath) {
    # Check if the directory exists or can be created
    $exportDirectory = Split-Path -Path $ExportPath -Parent
    if ($exportDirectory -and -not (Test-Path -Path $exportDirectory)) {
        try {
            New-Item -Path $exportDirectory -ItemType Directory -Force | Out-Null
            Write-Verbose "Created export directory: $exportDirectory"
        }
        catch {
            Write-Error "Cannot create export directory: $exportDirectory. Error: $($_.Exception.Message)"
            exit 1
        }
    }
    
    # Check if file path is writable (if file exists) or directory is writable
    if (Test-Path -Path $ExportPath) {
        try {
            $testFile = [System.IO.File]::OpenWrite($ExportPath)
            $testFile.Close()
        }
        catch {
            Write-Error "Export path is not writable: $ExportPath. Error: $($_.Exception.Message)"
            exit 1
        }
    }
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
    
    # Initialize array to store device details
    # This will hold all device information retrieved from Graph API
    $deviceDetails = @()
    
    # Determine which devices to query based on provided parameters
    # If DeviceId is specified, query only that device
    # If DeviceName is specified, search for devices matching that name
    # Otherwise, query all devices
    if ($DeviceId) {
        Write-Host "Querying details for device: $DeviceId" -ForegroundColor Yellow
        
        # Query device details for specific device
        # The Graph API endpoint returns comprehensive device information
        $deviceUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$DeviceId')"
        $device = Invoke-GraphApiWithRetry -Uri $deviceUrl -Headers $headers -Method Get
        
        if ($null -eq $device) {
            Write-Error "Device not found or API returned null response for device ID: $DeviceId"
            exit 1
        }
        
        # Create device detail object
        $deviceDetail = [PSCustomObject]@{
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
        }
        
        $deviceDetails += $deviceDetail
    }
    elseif ($DeviceName) {
        Write-Host "Searching for devices with name: $DeviceName" -ForegroundColor Yellow
        
        # Query all devices and filter by name
        # Filter devices by device name (supports partial matches)
        $devicesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=contains(deviceName,'$DeviceName')"
        $devices = Invoke-GraphApiWithRetry -Uri $devicesUrl -Headers $headers -Method Get
        
        # Handle null or empty response
        if ($null -eq $devices -or $null -eq $devices.value) {
            Write-Warning "No devices found matching name: $DeviceName"
            $devices = @{ value = @() }
        }
        
        # Process each matching device
        foreach ($device in $devices.value) {
            $deviceDetail = [PSCustomObject]@{
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
            }
            
            $deviceDetails += $deviceDetail
        }
    }
    else {
        Write-Host "Querying details for all devices..." -ForegroundColor Yellow
        
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
        
        # Process each device
        foreach ($device in $allDevices) {
            $deviceDetail = [PSCustomObject]@{
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
            }
            
            $deviceDetails += $deviceDetail
        }
    }
    
    # Display results
    Write-Host "`nDevice Details:" -ForegroundColor Cyan
    Write-Host "===============" -ForegroundColor Cyan
    
    if ($deviceDetails.Count -eq 0) {
        Write-Host "No devices found matching the specified criteria." -ForegroundColor Yellow
    }
    else {
        # Display each device's details
        foreach ($detail in $deviceDetails) {
            Write-Host "`nDevice: $($detail.DeviceName)" -ForegroundColor Green
            Write-Host "  Device ID: $($detail.DeviceId)" -ForegroundColor White
            Write-Host "  User: $($detail.UserPrincipalName) ($($detail.UserDisplayName))" -ForegroundColor White
            Write-Host "  Operating System: $($detail.OperatingSystem) $($detail.OSVersion)" -ForegroundColor White
            Write-Host "  Manufacturer: $($detail.Manufacturer)" -ForegroundColor White
            Write-Host "  Model: $($detail.Model)" -ForegroundColor White
            Write-Host "  Serial Number: $($detail.SerialNumber)" -ForegroundColor White
            Write-Host "  Enrollment Date: $($detail.EnrollmentDateTime)" -ForegroundColor White
            Write-Host "  Last Sync: $($detail.LastSyncDateTime)" -ForegroundColor White
            Write-Host "  Compliance State: $($detail.ComplianceState)" -ForegroundColor $(if ($detail.ComplianceState -eq 'Compliant') { 'Green' } else { 'Red' })
            Write-Host "  Management Agent: $($detail.ManagementAgent)" -ForegroundColor White
            Write-Host "  Enrollment Type: $($detail.DeviceEnrollmentType)" -ForegroundColor White
            Write-Host "  Supervised: $($detail.IsSupervised)" -ForegroundColor White
            Write-Host "  Encrypted: $($detail.IsEncrypted)" -ForegroundColor White
            if ($detail.TotalStorageSpaceInBytes) {
                $totalGB = [math]::Round($detail.TotalStorageSpaceInBytes / 1GB, 2)
                $freeGB = [math]::Round($detail.FreeStorageSpaceInBytes / 1GB, 2)
                Write-Host "  Storage: $freeGB GB free of $totalGB GB total" -ForegroundColor White
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
            $deviceDetails | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
            
            # Verify the export file was created successfully
            if (-not (Test-Path -Path $ExportPath)) {
                Write-Error "Export file was not created successfully at: $ExportPath"
                exit 1
            }
        }
    }
    
    Write-Host "`nQuery completed successfully. Found $($deviceDetails.Count) device(s)." -ForegroundColor Green
}
# Error handling block - catches exceptions during Graph API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Get-IntuneDeviceDetails: Failed to retrieve device details - $errMsg"
    
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

