<#
.SYNOPSIS
    Retrieves all applications published in Microsoft Intune and exports their assignment details.

.DESCRIPTION
    This script retrieves all applications published in Microsoft Intune and exports their
    assignment details to a CSV file. It connects to Microsoft Graph API using app registration
    credentials and retrieves information about each app including its type, version, assignments,
    and target groups.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.EXAMPLE
    .\Get-IntuneAllAppsAssignmentDetails.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.NOTES
    Requires: Microsoft.Graph PowerShell module
    Requires: Entra ID App Registration with appropriate Graph API permissions
#>

[CmdletBinding()]
# Define script parameters for credential input
# These allow credentials to be passed as command-line arguments
param(
    [Parameter(Mandatory=$false)]
    [string]$AppId,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientSecret
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

# Define the path where the CSV export file will be saved
# Use $env:TEMP for cross-system compatibility
$ExportCSVpath = Join-Path $env:TEMP "Get-IntuneAllAppsAssignmentDetails.csv"

# Extract the directory path from the full file path
$exportDirectory = Split-Path -Path $ExportCSVpath -Parent

# Check if the export directory exists, create it if it doesn't
# This prevents errors when trying to write the CSV file
if (-not (Test-Path -Path $exportDirectory)) {
    try {
        # Create the directory if it doesn't exist
        New-Item -Path $exportDirectory -ItemType Directory -Force | Out-Null
        Write-Host "Created export directory: $exportDirectory" -ForegroundColor Green
    }
    catch {
        # Exit if directory creation fails
        Write-Error "Failed to create export directory: $($_.Exception.Message)"
        exit 1
    }
}

# Credential handling: Check parameters first, then environment variables, then use placeholders
# This provides flexibility in how credentials are provided
$appid = if ($AppId) { $AppId } elseif ($env:INTUNE_APP_ID) { $env:INTUNE_APP_ID } else { '<Your Entra Registered App ID here>' }
$tenantid = if ($TenantId) { $TenantId } elseif ($env:INTUNE_TENANT_ID) { $env:INTUNE_TENANT_ID } else { '<Your Tenant ID here>' }
$secret = if ($ClientSecret) { $ClientSecret } elseif ($env:INTUNE_CLIENT_SECRET) { $env:INTUNE_CLIENT_SECRET } else { '<Your Entra Registered App Client Secret here>' }

# Validate that credentials are not placeholder values
# This prevents accidental use of placeholder credentials
if ($appid -match '^<.*>$' -or $tenantid -match '^<.*>$' -or $secret -match '^<.*>$') {
    Write-Error "Please provide valid credentials via parameters or environment variables (INTUNE_APP_ID, INTUNE_TENANT_ID, INTUNE_CLIENT_SECRET)."
    exit 1
}
 
# Prepare the OAuth2 token request body
# This uses the client credentials flow (app-only authentication)
$body =  @{
    Grant_Type    = "client_credentials"  # OAuth2 grant type for app-only authentication
    Scope         = "https://graph.microsoft.com/.default"  # Request all default permissions
    Client_Id     = $appid  # Application ID
    Client_Secret = $secret  # Application secret
}
 
# Request an access token from Microsoft Identity Platform
# This token will be used to authenticate Graph API requests
$connection = Invoke-RestMethod `
    -Uri https://login.microsoftonline.com/$tenantid/oauth2/v2.0/token `
    -Method POST `
    -Body $body

# Extract the access token from the response
$token = $connection.access_token

# Convert the token to a SecureString for use with Connect-MgGraph
# This helps protect the token in memory
$accessToken = ConvertTo-SecureString -String $token -AsPlainText -Force

# Connect to Microsoft Graph using the access token
# -NoWelcome suppresses the welcome message
Connect-MgGraph -AccessToken $accessToken -NoWelcome

### Main Section ###

# Define the Graph API version and resource endpoint
# Using Beta version to access the latest features
$graphApiVersion = "Beta"
$Resource = "deviceAppManagement/mobileApps"  # Endpoint for retrieving mobile apps

# Construct the full URI with query parameters
# $expand parameter includes related entities (Assignments, Category, Version) in the response
$uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$expand=Assignments,Category,Version"

# Retrieve all apps from Intune
try{
    # Make the Graph API request to get all mobile apps
    # .Value property contains the array of apps
    $response = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType PSObject
    
    # Handle null or empty response
    if ($null -eq $response -or $null -eq $response.Value) {
        Write-Warning "No apps found or API returned null response"
        $Apps = @()
    }
    else {
        $Apps = $response.Value
    }
}catch {
    # Error handling for API request failures
    $ex = $_.Exception
    if ($ex.Response) {
        Write-Host "Request to $Uri failed with HTTP Status $([int]$ex.Response.StatusCode) $($ex.Response.StatusDescription)" -ForegroundColor Red
        # Read and display error response body
        try {
            $errorResponse = $ex.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($errorResponse)
            $reader.BaseStream.Position = 0
            $reader.DiscardBufferedData()
            $responseBody = $reader.ReadToEnd()
            Write-Host "Response content:`n$responseBody" -ForegroundColor Red
        }
        catch {
            Write-Verbose "Could not read error response body: $($_.Exception.Message)"
        }
        Write-Error "Get-IntuneAllAppsAssignmentDetails: Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    }
    else {
        Write-Error "Get-IntuneAllAppsAssignmentDetails: Request to $Uri failed - $($ex.Exception.Message)"
    }
    Write-Host
    exit 1
}

# Display the total number of apps found
if ($Apps -and $Apps.Count -gt 0) {
    Write-Host "Number of Apps found is: $($Apps.Count)" -ForegroundColor Cyan
}
else {
    Write-Host "No apps found." -ForegroundColor Yellow
    exit 0
}

# Create an ArrayList to store the output data
# ArrayList is used for better performance when adding items dynamically
$Output=New-Object System.Collections.ArrayList

# Process each app retrieved from Intune
$appCount = 0
foreach ($App in $Apps) 
{
    $appCount++
    # Display progress indicator
    Write-Progress -Activity "Processing Intune Apps" -Status "Processing app $appCount of $($Apps.Count): $($App.displayName)" -PercentComplete (($appCount / $Apps.Count) * 100)
    
    # Display the current app name being processed
    Write-host "App Name: $($App.displayName)" -ForegroundColor Yellow
    
    # Convert the OData type identifier to a human-readable app type name
    # The @odata.type property contains the full type name from Graph API
    $AppType = switch($App.'@odata.type'){
                        "#microsoft.graph.androidLobApp" { 'Android LOB' }
                        "#microsoft.graph.androidStoreApp" { 'Google Play' }
                        "#microsoft.graph.androidManagedStoreWebApp" { 'Android Managed Store Web App' }
                        "#microsoft.graph.managedAndroidLobApp" { 'Managed Android LOB' }
                        "#microsoft.graph.managedAndroidStoreApp" { 'Managed Google Play' }
                        "#microsoft.graph.androidForWorkApp" { 'Android for Work' }
                        "#microsoft.graph.iosLobApp" { 'iOS LOB' }
                        "#microsoft.graph.iosStoreApp" { 'iOS Store' }
                        "#microsoft.graph.managedIOSLobApp" { 'Managed iOS LOB' }
                        "#microsoft.graph.managedIOSStoreApp" { 'Managed iOS Store' }
                        "#microsoft.graph.iosVppApp" { 'iOS VPP' }
                        "#microsoft.graph.iosVppEBook" { 'iOS VPP Ebook' }
                        "#microsoft.graph.iosiPadOSWebClip" { 'iOS Weblink' }
                        "#microsoft.graph.webApp" { 'Web Link' }
                        "#microsoft.graph.microsoftStoreForBusiness" { 'Microsoft Store' }
                        "#microsoft.graph.winGetApp" { 'New Microsoft Store' }
                        "#microsoft.graph.windowsStoreApp" { 'Windows Store App' }
                        "#microsoft.graph.windowsPhoneXAP" { 'Windows Phone XAP' }
                        "#microsoft.graph.win32LobApp" { 'Win32' }
                        "#microsoft.graph.windowsAppX" { 'Windows AppX' }
                        "#microsoft.graph.windowsMicrosoftEdgeApp" { 'Intune Windows Built-in Edge App' }
                        "#microsoft.graph.windowsMobileMSI" { 'Windows MSI' }
                        "#microsoft.graph.windowsUniversalAppX" { 'Windows Universl AppX' }
                        "#microsoft.graph.windowsUniversalAppXContainedApp" { 'Windows Universal AppX Contained App' }
                        "#microsoft.graph.windowsWebApp"  { 'Windows Web App' }
                        "#microsoft.graph.officeSuiteApp" { 'Intune Windows Built-in M365 App for Enterprise' }
                        "#microsoft.graph.macOSDmgApp" { 'MacOS DMG' }
                        "#microsoft.graph.macOSLobApp" { 'MacOS LOB' }
                        "#microsoft.graph.macOSMdatpApp" { 'MacOS MDATP' }
                        "#microsoft.graph.macOSMicrosoftDefenderApp" { 'MacOS Defender' }
                        "#microsoft.graph.macOSMicrosoftEdgeApp" { 'MacOS Edge' }
                        "#microsoft.graph.macOSOfficeSuiteApp" { 'Intune MacOS Built-in M365 App for Enterprise' }
                        "#microsoft.graph.macOSPkgApp" { 'MacOS PKG' }
                        "#microsoft.graph.macOsVppApp" { 'MacOS VPP' }
                        default { 'Unknown' }  # Fallback for unknown app types
    }

    # Display the app type
    Write-host "App Type: $AppType" -ForegroundColor Cyan

    # Get the app version (may be null for some app types)
    $AppVer = $App.displayVersion
    Write-host "App Version: $AppVer"

    # Check if the app has any assignments
    # Apps can have zero, one, or multiple assignments
    If(($null -eq $App.assignments) -or ($App.assignments -eq "") -or ($App.assignments.count -lt 1))
    {
        # App has no assignments - add to output with "Not Assigned"
        Write-Host "No assignments for this app"
        $GroupName = "Not Assigned"
        # Add a row to the output array with assignment information
        $Output.Add( (New-Object -TypeName PSObject -Property @{"Name"="$($App.displayName)"; "Platform" = "$AppType"; "Version" = "$AppVer"; "Group" = "$GroupName"; "Assignment" = "$null"} ) )
    } 
    else 
    {
        # App has one or more assignments - process each assignment
        foreach($assignment in $App.assignments)
        {
            # Display the assignment intent (Available or Required)
            write-host "Assignment intent: $($assignment.intent)"
 
            # Determine the target group name based on assignment target type
            If ($($assignment.target.'@odata.type') -like "*allLicensedUsersAssignmentTarget"){
                # App is assigned to all licensed users
                Write-Host "Published to All Users"
                $GroupName = "All Users"
            } elseif($($assignment.target.'@odata.type') -like "*allDevicesAssignmentTarget"){
                # App is assigned to all devices
                Write-Host "Published to All Devices"
                $GroupName = "All Devices"
            }
            else
            {
                # App is assigned to a specific Azure AD group
                # Need to look up the group display name using the Group ID
                write-host "Group ID: $($assignment.target.GroupID)"
                # Query Graph API to get the group's display name
                try {
                    $group = Get-MgGroup -GroupId $assignment.target.GroupID -ErrorAction Stop
                    if ($null -ne $group -and $null -ne $group.DisplayName) {
                        $GroupName = $group.DisplayName
                    }
                    else {
                        $GroupName = $assignment.target.GroupID
                        Write-Warning "Could not retrieve group name for Group ID: $($assignment.target.GroupID)"
                    }
                }
                catch {
                    $GroupName = $assignment.target.GroupID
                    Write-Verbose "Failed to retrieve group name for Group ID: $($assignment.target.GroupID) - $($_.Exception.Message)"
                }
            }

            # Display the resolved group name
            Write-Host "Group Name: $GroupName"
            
            # Add a row to the output array with all app and assignment details
            $Output.Add( (New-Object -TypeName PSObject -Property @{"Name"="$($App.displayName)"; "Platform" = "$AppType"; "Version" = "$AppVer"; "Group" = "$GroupName"; "Assignment" = "$($assignment.intent)"} ) )
        }
    }
}
Write-Progress -Activity "Processing Intune Apps" -Completed

# Export the collected data to CSV file
# Select-Object filters to only the columns we want in the CSV
$output | Select-Object Name,Platform,Version,Assignment,Group | Export-CSV -Path $ExportCSVpath -Encoding utf8 -NoTypeInformation

# Disconnect from Microsoft Graph to clean up the session
Disconnect-MgGraph
