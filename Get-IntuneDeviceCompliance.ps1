<#
.SYNOPSIS
    Retrieves compliance status for Intune-managed devices.

.DESCRIPTION
    This script connects to Microsoft Graph API and retrieves compliance status information
    for Intune-managed devices. It can query compliance for a specific device, all devices,
    or devices for a specific user. The script displays compliance policy assignments,
    compliance state, and any non-compliance reasons.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.PARAMETER DeviceId
    Optional. The device ID to query compliance for. If not specified, retrieves all devices.

.PARAMETER UserPrincipalName
    Optional. The user principal name to query devices for. If specified, retrieves compliance
    for all devices owned by this user.

.PARAMETER ExportPath
    Optional. Path to export results to CSV file. If not specified, results are displayed only.

.EXAMPLE
    .\Get-IntuneDeviceCompliance.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntuneDeviceCompliance.ps1 -DeviceId 'device-guid' -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntuneDeviceCompliance.ps1 -UserPrincipalName 'user@domain.com' -ExportPath 'C:\Temp\ComplianceReport.csv'

.NOTES
    Requires: Microsoft.Graph PowerShell module
    Requires: Entra ID App Registration with DeviceManagementManagedDevices.Read.All and DeviceManagementConfiguration.Read.All permissions
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
    [string]$UserPrincipalName,
    
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
if ($DeviceId -and -not ($DeviceId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')) {
    Write-Error "DeviceId must be a valid GUID format."
    exit 1
}

# Validate UserPrincipalName format if provided
if ($UserPrincipalName -and -not ($UserPrincipalName -match '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')) {
    Write-Error "UserPrincipalName must be a valid email address format."
    exit 1
}

# Sanitize UserPrincipalName input for OData query
if ($UserPrincipalName) {
    # Remove potentially dangerous characters from OData filter
    $UserPrincipalName = $UserPrincipalName -replace "[';]", ""
    if ([string]::IsNullOrWhiteSpace($UserPrincipalName)) {
        Write-Error "UserPrincipalName cannot be empty after sanitization."
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
    Import-Module "$PSScriptRoot\IntuneCommon.psm1" -Force -ErrorAction Stop

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
    
    # Initialize array to store compliance results
    # This will hold all compliance information retrieved from Graph API
    $complianceResults = @()
    
    # Determine which devices to query based on provided parameters
    # If DeviceId is specified, query only that device
    # If UserPrincipalName is specified, query all devices for that user
    # Otherwise, query all devices
    if ($DeviceId) {
        Write-Host "Querying compliance for device: $DeviceId" -ForegroundColor Yellow
        
        # Query device compliance status for specific device
        # The Graph API endpoint returns compliance state and policy assignments
        $deviceComplianceUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$DeviceId')?`$expand=detectedApps"
        $device = Invoke-GraphApiWithRetry -Uri $deviceComplianceUrl -Headers $headers -Method Get
        
        if ($null -eq $device) {
            Write-Error "Device not found or API returned null response for device ID: $DeviceId"
            exit 1
        }
        
        # Query compliance policies for this device
        # Get all compliance policies that apply to this device
        $compliancePoliciesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies"
        $compliancePolicies = Invoke-GraphApiWithRetry -Uri $compliancePoliciesUrl -Headers $headers -Method Get
        
        # Handle null or empty response
        $compliancePolicyNames = @()
        if ($null -ne $compliancePolicies -and $null -ne $compliancePolicies.value) {
            $compliancePolicyNames = $compliancePolicies.value | ForEach-Object { $_.displayName }
        }
        
        # Create compliance result object
        $complianceResult = [PSCustomObject]@{
            DeviceId = $device.id
            DeviceName = $device.deviceName
            UserPrincipalName = $device.userPrincipalName
            ComplianceState = $device.complianceState
            LastSyncDateTime = $device.lastSyncDateTime
            OSVersion = $device.osVersion
            ManagementAgent = $device.managementAgent
            CompliancePolicies = ($compliancePolicyNames) -join '; '
        }
        
        $complianceResults += $complianceResult
    }
    elseif ($UserPrincipalName) {
        Write-Host "Querying compliance for devices owned by: $UserPrincipalName" -ForegroundColor Yellow
        
        # Query all devices for the specified user
        # Filter devices by user principal name
        $devicesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=userPrincipalName eq '$UserPrincipalName'"
        $devices = Invoke-GraphApiWithRetry -Uri $devicesUrl -Headers $headers -Method Get
        
        # Handle null or empty response
        if ($null -eq $devices -or $null -eq $devices.value) {
            Write-Warning "No devices found for user: $UserPrincipalName"
            $devices = @{ value = @() }
        }
        
        # Query compliance policies
        $compliancePoliciesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies"
        $compliancePolicies = Invoke-GraphApiWithRetry -Uri $compliancePoliciesUrl -Headers $headers -Method Get
        
        # Handle null or empty response
        $compliancePolicyNames = @()
        if ($null -ne $compliancePolicies -and $null -ne $compliancePolicies.value) {
            $compliancePolicyNames = $compliancePolicies.value | ForEach-Object { $_.displayName }
        }
        
        foreach ($device in $devices.value) {
            $complianceResult = [PSCustomObject]@{
                DeviceId = $device.id
                DeviceName = $device.deviceName
                UserPrincipalName = $device.userPrincipalName
                ComplianceState = $device.complianceState
                LastSyncDateTime = $device.lastSyncDateTime
                OSVersion = $device.osVersion
                ManagementAgent = $device.managementAgent
                CompliancePolicies = ($compliancePolicyNames) -join '; '
            }
            
            $complianceResults += $complianceResult
        }
    }
    else {
        Write-Host "Querying compliance for all devices..." -ForegroundColor Yellow
        
        # Query all managed devices from Intune
        $devicesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        $devices = Invoke-GraphApiWithRetry -Uri $devicesUrl -Headers $headers -Method Get
        
        # Handle null or empty response and pagination
        $allDevices = @()
        if ($null -ne $devices -and $null -ne $devices.value) {
            $allDevices += $devices.value
        }
        
        # Query compliance policies
        $compliancePoliciesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies"
        $compliancePolicies = Invoke-GraphApiWithRetry -Uri $compliancePoliciesUrl -Headers $headers -Method Get
        
        # Extract policy names from response
        $compliancePolicyNames = @()
        if ($null -ne $compliancePolicies -and $null -ne $compliancePolicies.value) {
            $compliancePolicyNames = $compliancePolicies.value | ForEach-Object { $_.displayName }
        }
        
        # Handle pagination for devices
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
        
        foreach ($device in $allDevices) {
            $complianceResult = [PSCustomObject]@{
                DeviceId = $device.id
                DeviceName = $device.deviceName
                UserPrincipalName = $device.userPrincipalName
                ComplianceState = $device.complianceState
                LastSyncDateTime = $device.lastSyncDateTime
                OSVersion = $device.osVersion
                ManagementAgent = $device.managementAgent
                CompliancePolicies = ($compliancePolicyNames) -join '; '
            }
            
            $complianceResults += $complianceResult
        }
    }
    
    # Display results
    Write-Host "`nCompliance Results:" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    
    if ($complianceResults.Count -eq 0) {
        Write-Host "No devices found matching the specified criteria." -ForegroundColor Yellow
    }
    else {
        # Display each compliance result
        foreach ($result in $complianceResults) {
            Write-Host "`nDevice: $($result.DeviceName)" -ForegroundColor Green
            Write-Host "  Device ID: $($result.DeviceId)" -ForegroundColor White
            Write-Host "  User: $($result.UserPrincipalName)" -ForegroundColor White
            Write-Host "  Compliance State: $($result.ComplianceState)" -ForegroundColor $(if ($result.ComplianceState -eq 'Compliant') { 'Green' } else { 'Red' })
            Write-Host "  Last Sync: $($result.LastSyncDateTime)" -ForegroundColor White
            Write-Host "  OS Version: $($result.OSVersion)" -ForegroundColor White
            Write-Host "  Management Agent: $($result.ManagementAgent)" -ForegroundColor White
        }
        
        # Export to CSV if path is specified
        if ($ExportPath) {
            # Validate export path (should have been validated earlier, but double-check)
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
            
            # Export results to CSV file
            try {
                $complianceResults | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
                Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
                
                # Verify the export file was created successfully
                if (-not (Test-Path -Path $ExportPath)) {
                    Write-Error "Export file was not created successfully at: $ExportPath"
                    exit 1
                }
            }
            catch {
                Write-Error "Get-IntuneDeviceCompliance: Failed to export results to CSV - $($_.Exception.Message)"
                exit 1
            }
        }
    }
    
    Write-Host "`nQuery completed successfully." -ForegroundColor Green
}
# Error handling block - catches exceptions during Graph API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Get-IntuneDeviceCompliance: Failed to retrieve device compliance information - $errMsg"
    
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

