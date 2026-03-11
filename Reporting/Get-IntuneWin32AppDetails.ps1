<#
.SYNOPSIS
    Retrieves detailed information about Win32 app deployments.

.DESCRIPTION
    This script connects to Microsoft Graph API and retrieves detailed information about
    Win32 app deployments in Intune including app installation status, deployment progress,
    installation errors, and app assignment details. It can query a specific app or all apps.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.PARAMETER IntuneAppId
    Optional. The Intune app ID to query. If not specified, retrieves all Win32 apps.

.PARAMETER DeviceId
    Optional. The device ID to query app installation status for. If specified, shows
    app installation status for that device.

.PARAMETER ExportPath
    Optional. Path to export results to CSV file. If not specified, results are displayed only.

.EXAMPLE
    .\Get-IntuneWin32AppDetails.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntuneWin32AppDetails.ps1 -IntuneAppId 'app-guid' -DeviceId 'device-guid' -ExportPath 'C:\Temp\AppDetails.csv'

.NOTES
    Requires: Microsoft.Graph PowerShell module
    Requires: Entra ID App Registration with DeviceManagementApps.Read.All and DeviceManagementManagedDevices.Read.All permissions
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
    [string]$IntuneAppId,
    
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

# Validate IntuneAppId format if provided
if ($IntuneAppId -and -not ($IntuneAppId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')) {
    Write-Error "IntuneAppId must be a valid GUID format."
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
    
    # Initialize array to store app details
    # This will hold all app information retrieved from Graph API
    $appDetails = @()
    
    # Determine which apps to query
    # If IntuneAppId is specified, query only that app
    # Otherwise, query all Win32 apps
    if ($IntuneAppId) {
        Write-Host "Querying details for app: $IntuneAppId" -ForegroundColor Yellow
        
        # Query app details for specific app
        # Get comprehensive app information
        $appUrl = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps('$IntuneAppId')"
        try {
            $app = Invoke-GraphApiWithRetry -Uri $appUrl -Headers $headers -Method Get
            
            if ($null -eq $app) {
                Write-Error "App not found or API returned null response for app ID: $IntuneAppId"
                exit 1
            }
            
            # Check if this is a Win32 app
            # Win32 apps have a specific OData type
            if ($app.'@odata.type' -eq '#microsoft.graph.win32LobApp') {
                # Get app assignments
                # Get all assignments for this app
                $assignmentsUrl = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps('$IntuneAppId')/assignments"
                try {
                    $assignments = Invoke-GraphApiWithRetry -Uri $assignmentsUrl -Headers $headers -Method Get
                    if ($null -eq $assignments -or $null -eq $assignments.value) {
                        $assignments = @{ value = @() }
                    }
                }
                catch {
                    $assignments = @{ value = @() }
                    Write-Verbose "Could not retrieve assignments for app: $($_.Exception.Message)"
                }
                
                # Create app detail object
                $appDetail = [PSCustomObject]@{
                    AppId = $app.id
                    AppName = $app.displayName
                    AppDescription = $app.description
                    AppVersion = $app.version
                    AppType = "Win32"
                    Publisher = $app.publisher
                    InstallCommandLine = $app.installCommandLine
                    UninstallCommandLine = $app.uninstallCommandLine
                    MinimumSupportedOperatingSystem = $app.minimumSupportedOperatingSystem
                    AssignmentCount = $assignments.value.Count
                    Assignments = ($assignments.value | ForEach-Object { $_.target.'@odata.type' }) -join '; '
                }
                
                $appDetails += $appDetail
            }
            else {
                Write-Warning "App $IntuneAppId is not a Win32 app. Type: $($app.'@odata.type')"
            }
        }
        catch {
            Write-Error "Failed to retrieve app details: $($_.Exception.Message)"
        }
    }
    else {
        Write-Host "Querying details for all Win32 apps..." -ForegroundColor Yellow
        
        # Query all Win32 apps from Intune
        $appsUrl = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps?`$filter=isof('microsoft.graph.win32LobApp')"
        $apps = Invoke-GraphApiWithRetry -Uri $appsUrl -Headers $headers -Method Get
        
        # Handle null response and pagination
        $allApps = @()
        if ($null -ne $apps -and $null -ne $apps.value) {
            $allApps += $apps.value
        }
        $pageCount = 1
        while ($null -ne $apps -and $null -ne $apps.'@odata.nextLink') {
            $pageCount++
            Write-Progress -Activity "Retrieving Win32 apps" -Status "Fetching page $pageCount" -PercentComplete -1
            $apps = Invoke-GraphApiWithRetry -Uri $apps.'@odata.nextLink' -Headers $headers -Method Get
            if ($null -ne $apps -and $null -ne $apps.value) {
                $allApps += $apps.value
            }
        }
        Write-Progress -Activity "Retrieving Win32 apps" -Completed
        
        Write-Host "Found $($allApps.Count) Win32 app(s). Processing..." -ForegroundColor Cyan
        
        # Process each app and get its details
        $appCount = 0
        foreach ($app in $allApps) {
            $appCount++
            Write-Progress -Activity "Processing Win32 apps" -Status "Processing app $appCount of $($allApps.Count): $($app.displayName)" -PercentComplete (($appCount / $allApps.Count) * 100)
            Write-Host "Processing app $appCount of $($allApps.Count): $($app.displayName)" -ForegroundColor Cyan
            
            # Get app assignments
            # Get all assignments for this app
            $assignmentsUrl = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps('$($app.id)')/assignments"
            try {
                $assignments = Invoke-GraphApiWithRetry -Uri $assignmentsUrl -Headers $headers -Method Get
                if ($null -eq $assignments -or $null -eq $assignments.value) {
                    $assignments = @{ value = @() }
                }
            }
            catch {
                $assignments = @{ value = @() }
                Write-Verbose "Could not retrieve assignments for app: $($_.Exception.Message)"
            }
            
            # Create app detail object
            $appDetail = [PSCustomObject]@{
                AppId = $app.id
                AppName = $app.displayName
                AppDescription = $app.description
                AppVersion = $app.version
                AppType = "Win32"
                Publisher = $app.publisher
                InstallCommandLine = $app.installCommandLine
                UninstallCommandLine = $app.uninstallCommandLine
                MinimumSupportedOperatingSystem = $app.minimumSupportedOperatingSystem
                AssignmentCount = $assignments.value.Count
                Assignments = ($assignments.value | ForEach-Object { $_.target.'@odata.type' }) -join '; '
            }
            
            $appDetails += $appDetail
        }
        Write-Progress -Activity "Processing Win32 apps" -Completed
    }
    
    # If DeviceId is specified, get app installation status for that device
    if ($DeviceId) {
        Write-Host "`nQuerying app installation status for device: $DeviceId" -ForegroundColor Yellow
        
        # Query device app installation status
        $deviceAppsUrl = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices('$DeviceId')/detectedApps"
        try {
            $deviceApps = Invoke-GraphApiWithRetry -Uri $deviceAppsUrl -Headers $headers -Method Get
            
            # Handle null or empty response
            if ($null -ne $deviceApps -and $null -ne $deviceApps.value) {
                Write-Host "Found $($deviceApps.value.Count) app(s) detected on device." -ForegroundColor Cyan
                
                # Match detected apps with Win32 apps
                foreach ($appDetail in $appDetails) {
                    $detectedApp = $deviceApps.value | Where-Object { $_.displayName -eq $appDetail.AppName }
                    if ($detectedApp) {
                        $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallStatus" -Value "Installed" -Force
                        $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallVersion" -Value $detectedApp.version -Force
                    }
                    else {
                        $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallStatus" -Value "Not Installed" -Force
                        $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallVersion" -Value $null -Force
                    }
                }
            }
            else {
                Write-Warning "No apps detected on device or API returned null response"
                foreach ($appDetail in $appDetails) {
                    $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallStatus" -Value "Unknown" -Force
                    $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallVersion" -Value $null -Force
                }
            }
        }
        catch {
            Write-Warning "Could not retrieve app installation status for device: $($_.Exception.Message)"
            foreach ($appDetail in $appDetails) {
                $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallStatus" -Value "Unknown" -Force
                $appDetail | Add-Member -MemberType NoteProperty -Name "DeviceInstallVersion" -Value $null -Force
            }
        }
    }
    
    # Display results
    Write-Host "`nWin32 App Details:" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan
    
    if ($appDetails.Count -eq 0) {
        Write-Host "No Win32 apps found matching the specified criteria." -ForegroundColor Yellow
    }
    else {
        # Display each app's details
        foreach ($detail in $appDetails) {
            Write-Host "`nApp: $($detail.AppName)" -ForegroundColor Green
            Write-Host "  App ID: $($detail.AppId)" -ForegroundColor White
            Write-Host "  Publisher: $($detail.Publisher)" -ForegroundColor White
            Write-Host "  Version: $($detail.AppVersion)" -ForegroundColor White
            Write-Host "  Install Command: $($detail.InstallCommandLine)" -ForegroundColor White
            Write-Host "  Assignment Count: $($detail.AssignmentCount)" -ForegroundColor White
            if ($detail.DeviceInstallStatus) {
                Write-Host "  Device Install Status: $($detail.DeviceInstallStatus)" -ForegroundColor $(if ($detail.DeviceInstallStatus -eq 'Installed') { 'Green' } else { 'Yellow' })
                if ($detail.DeviceInstallVersion) {
                    Write-Host "  Device Install Version: $($detail.DeviceInstallVersion)" -ForegroundColor White
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
            $appDetails | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
            
            # Verify the export file was created successfully
            if (-not (Test-Path -Path $ExportPath)) {
                Write-Error "Export file was not created successfully at: $ExportPath"
                exit 1
            }
        }
    }
    
    Write-Host "`nQuery completed successfully. Found $($appDetails.Count) Win32 app(s)." -ForegroundColor Green
}
# Error handling block - catches exceptions during Graph API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Get-IntuneWin32AppDetails: Failed to retrieve Win32 app details - $errMsg"
    
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

