<#
.SYNOPSIS
    Tests connectivity to Microsoft Intune and Microsoft 365 endpoints required for Intune management.

.DESCRIPTION
    This script tests connectivity to all mandatory Intune and Microsoft 365 Common endpoints
    required for Intune device management functionality. It queries Microsoft's Office 365 IP
    Address and URL web service to obtain the latest endpoint URLs for Service Area "MEM"
    (Microsoft Endpoint Manager) and "M365 Common", then tests connectivity to those endpoints
    from the client device. The script handles proxy configurations and provides detailed
    output about which endpoints are accessible.

.PARAMETER ExportPath
    Optional path to export test results to a JSON file.
    If specified, test results including network configuration and endpoint status will be saved.

.EXAMPLE
    .\Test-IntuneConnectivity.ps1

.EXAMPLE
    .\Test-IntuneConnectivity.ps1 -ExportPath "C:\Temp\ConnectivityTest.json"

.NOTES
    Prerequisites:
    - PowerShell 3.0 or later
    - PS script execution policy: Bypass (or appropriate policy to allow script execution)
    - Does not require elevation (runs with current user permissions)
    - Internet connectivity to reach Microsoft endpoints
    
    The script:
    - Queries Microsoft's endpoint web service for the latest endpoint URLs
    - Tests connectivity to mandatory Intune and M365 Common endpoints
    - Handles proxy configurations automatically
    - Displays network configuration information for troubleshooting
    - Returns $true if all mandatory endpoints are accessible, $false otherwise
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ExportPath
)

Function Get-ProxySettings {
    <#
    .SYNOPSIS
        Retrieves the current Windows HTTP proxy configuration.
    
    .DESCRIPTION
        This function checks the system's winHTTP proxy settings using netsh command.
        It determines if the system is using a proxy server or has direct internet access.
        Returns the proxy server address or "NoProxy" if direct access is configured.
    
    .OUTPUTS
        String: Proxy server address (with http:// prefix) or "NoProxy"
    #>
    
    # Display status message to user
    Write-Host "Checking winHTTP proxy settings..." -ForegroundColor Yellow
    
    # Initialize proxy server variable with default value
    $ProxyServer = "NoProxy"
    
    # Execute netsh command to retrieve winHTTP proxy configuration
    # This command shows the current proxy settings for Windows HTTP services
    $winHTTP = netsh winhttp show proxy
    
    # Parse the output to find the proxy server line
    # Select-String searches for lines containing "server" (case-insensitive)
    $Proxy = $winHTTP | Select-String server
    
    # Extract the proxy server address from the output
    # TrimStart removes the "Proxy Server(s) :  " prefix to get just the address
    $ProxyServer = $Proxy.ToString().TrimStart("Proxy Server(s) :  ")

    # Check if the output indicates direct access (no proxy configured)
    if ($ProxyServer -eq "Direct access (no proxy server).") {
        # Set to "NoProxy" to indicate direct internet access
        $ProxyServer = "NoProxy"
        Write-Host "Access Type : DIRECT"
    }

    # Check if a proxy server is configured and ensure it has http:// prefix
    # This is required for the Invoke-WebRequest cmdlet to use the proxy correctly
    if ( ($ProxyServer -ne "NoProxy") -and (-not($ProxyServer.StartsWith("http://")))) {
        # Display proxy configuration to user
        Write-Host "Access Type : PROXY"
        Write-Host "Proxy Server List :" $ProxyServer
        # Add http:// prefix if not present (required for proxy usage)
        $ProxyServer = "http://" + $ProxyServer
    }
    
    # Return the proxy server address or "NoProxy"
    return $ProxyServer
}

Function Get-M365CommonEndpointList {
    <#
    .SYNOPSIS
        Retrieves the list of Microsoft 365 Common endpoints required for Intune connectivity.
    
    .DESCRIPTION
        This function queries Microsoft's Office 365 IP Address and URL web service to get
        the latest list of M365 Common endpoints. These endpoints are required for various
        Microsoft 365 services to function properly. The function filters and processes
        the endpoint data for connectivity testing.
    
    .OUTPUTS
        Array: Processed endpoint list with category and URL information
    #>
    
    # Function to invoke endpoint API with retry logic
    # Handles transient failures (429, 503, etc.) with exponential backoff
    function Invoke-EndpointApiWithRetry {
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
                    Write-Verbose "Endpoint API call failed with status $statusCode. Retrying in $delaySeconds seconds (attempt $retryCount of $MaxRetries)..."
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
    
    # Query Microsoft's Office 365 endpoint web service for Common service area endpoints
    # The service provides up-to-date endpoint URLs that may change over time
    # ServiceAreas=Common: Requests endpoints for common M365 services
    # clientrequestid: Unique GUID for tracking the request
    try {
        $endpointListM365 = (Invoke-EndpointApiWithRetry -Uri ("https://endpoints.office.com/endpoints/WorldWide?ServiceAreas=Common`&clientrequestid=" + ([GUID]::NewGuid()).Guid)) | Where-Object { $_.ServiceArea -eq "Common" -and $_.urls }
    }
    catch {
        Write-Error "Get-M365CommonEndpointList: Failed to retrieve M365 Common endpoints - $($_.Exception.Message)"
        return @()
    }

    # Create a mapping of endpoint IDs to human-readable categories
    # This helps identify what each endpoint is used for during testing
    # mandatory = $true indicates these endpoints are required for Intune to function
    [PsObject[]]$endpointListCategoriesM365 = @()
    $endpointListCategoriesM365 += [PsObject]@{id = 56; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 59; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 78; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 83; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 84; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 125; category = 'M365 Common'; mandatory = $true }
    $endpointListCategoriesM365 += [PsObject]@{id = 156; category = 'M365 Common'; mandatory = $true }
    
    # Process the endpoint list and create a structured output object
    # This combines endpoint data with category information for easier processing
    [PsObject[]]$endpointRequestListM365 = @()
    for ($i = 0; $i -lt $endpointListM365.Count; $i++) {
        # Match each endpoint with its category based on ID
        # Create an object with id, category, urls, and mandatory flag
        $endpointRequestListM365 += [PsObject]@{ id = $endpointListM365[$i].id; category = ($endpointListCategoriesM365 | Where-Object { $_.id -eq $endpointListM365[$i].id }).category; urls = $endpointListM365[$i].urls; mandatory = ($endpointListCategoriesM365 | Where-Object { $_.id -eq $endpointListM365[$i].id }).mandatory }
    }

    # Clean up URL list: Remove wildcard prefixes and ensure uniqueness
    # Wildcard prefixes (*.) are not useful for direct connectivity testing
    for ($i = 0; $i -lt $endpointRequestListM365.Count; $i++) {
        for ($j = 0; $j -lt $endpointRequestListM365[$i].urls.Count; $j++) {
            # Remove *. prefix from URLs (e.g., "*.example.com" becomes "example.com")
            $targetUrl = $endpointRequestListM365[$i].urls[$j].replace('*.', '')
            $endpointRequestListM365[$i].urls[$j] = $targetUrl
        }
        # Remove duplicate URLs and sort alphabetically for consistency
        $endpointRequestListM365[$i].urls = $endpointRequestListM365[$i].urls | Sort-Object -Unique
    }
    
    # Return the processed endpoint list
    return $endpointRequestListM365
}

Function Get-IntuneEndpointList {
    <#
    .SYNOPSIS
        Retrieves the list of Microsoft Endpoint Manager (Intune) endpoints required for connectivity.
    
    .DESCRIPTION
        This function queries Microsoft's Office 365 IP Address and URL web service to get
        the latest list of MEM (Microsoft Endpoint Manager) endpoints. These endpoints are
        specifically required for Intune device management functionality. The function
        filters and processes the endpoint data for connectivity testing.
    
    .OUTPUTS
        Array: Processed endpoint list with category and URL information
    #>
    
    # Query Microsoft's Office 365 endpoint web service for MEM (Microsoft Endpoint Manager) service area endpoints
    # ServiceAreas=MEM: Requests endpoints specifically for Intune/MEM services
    # clientrequestid: Unique GUID for tracking the request
    try {
        $endpointList = (Invoke-EndpointApiWithRetry -Uri ("https://endpoints.office.com/endpoints/WorldWide?ServiceAreas=MEM`&clientrequestid=" + ([GUID]::NewGuid()).Guid)) | Where-Object { $_.ServiceArea -eq "MEM" -and $_.urls }
    }
    catch {
        Write-Error "Get-IntuneEndpointList: Failed to retrieve Intune endpoints - $($_.Exception.Message)"
        return @()
    }

    # Create a mapping of endpoint IDs to human-readable categories
    # mandatory = $true: Endpoints required for basic Intune functionality
    # mandatory = $false: Optional endpoints (e.g., platform-specific or feature-specific)
    [PsObject[]]$endpointListCategories = @()
    $endpointListCategories += [PsObject]@{id = 163; category = 'Global'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 164; category = 'Delivery Optimization'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 165; category = 'NTP Sync'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 169; category = 'Windows Notifications & Store'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 170; category = 'Scripts & Win32 Apps'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 171; category = 'Push Notifications'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 172; category = 'Delivery Optimization'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 173; category = 'Autopilot Self-deploy'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 178; category = 'Apple Device Management'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 179; category = 'Android (AOSP) Device Management'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 181; category = 'Remote Help'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 182; category = 'Collect Diagnostics'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 186; category = 'Microsoft Azure attestation - Windows 11 only'; mandatory = $true }
    $endpointListCategories += [PsObject]@{id = 187; category = 'Android Remote Help'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 188; category = 'Remote Help GCC Dependency'; mandatory = $false }
    $endpointListCategories += [PsObject]@{id = 189; category = 'Feature Flighting'; mandatory = $false }
    
    # Process the endpoint list and create a structured output object
    # This combines endpoint data with category information for easier processing
    [PsObject[]]$endpointRequestList = @()
    for ($i = 0; $i -lt $endpointList.Count; $i++) {
        # Match each endpoint with its category based on ID
        # Create an object with id, category, urls, and mandatory flag
        $endpointRequestList += [PsObject]@{ id = $endpointList[$i].id; category = ($endpointListCategories | Where-Object { $_.id -eq $endpointList[$i].id }).category; urls = $endpointList[$i].urls; mandatory = ($endpointListCategories | Where-Object { $_.id -eq $endpointList[$i].id }).mandatory }
    }

    # Clean up URL list: Remove wildcard prefixes and ensure uniqueness
    # Wildcard prefixes (*.) are not useful for direct connectivity testing
    for ($i = 0; $i -lt $endpointRequestList.Count; $i++) {
        for ($j = 0; $j -lt $endpointRequestList[$i].urls.Count; $j++) {
            # Remove *. prefix from URLs (e.g., "*.example.com" becomes "example.com")
            $targetUrl = $endpointRequestList[$i].urls[$j].replace('*.', '')
            $endpointRequestList[$i].urls[$j] = $targetUrl
        }
        # Remove duplicate URLs and sort alphabetically for consistency
        $endpointRequestList[$i].urls = $endpointRequestList[$i].urls | Sort-Object -Unique
    }
    
    # Return the processed endpoint list
    return $endpointRequestList
}

Function Test-DeviceIntuneConnectivity {
    <#
    .SYNOPSIS
        Tests connectivity to all required Intune and Microsoft 365 endpoints.
    
    .DESCRIPTION
        This function tests connectivity to mandatory Intune and M365 Common endpoints
        from the device. It checks each endpoint using HTTP requests and reports
        success or failure. The function handles proxy configurations and provides
        detailed output about which endpoints are accessible.
    
    .OUTPUTS
        Boolean: $true if all mandatory endpoints are accessible, $false otherwise
    #>
    
    # Store the original error action preference to restore it later
    # This prevents the function from permanently changing the global error handling behavior
    $originalErrorAction = $ErrorActionPreference
    
    # Set error action to SilentlyContinue to prevent errors from stopping the test
    # This allows the script to continue testing other endpoints even if some fail
    $ErrorActionPreference = 'SilentlyContinue'
    
    # Initialize variables for tracking test results
    $TestFailed = $false  # Flag to track if any mandatory endpoint failed
    $ProxyServer = Get-ProxySettings  # Get proxy configuration
    $endpointListM365Common = Get-M365CommonEndpointList  # Get M365 Common endpoints
    $endpointListIntune = Get-IntuneEndpointList  # Get Intune-specific endpoints
    $script:failedEndpointList = @{}  # Hash table to store failed endpoints by category (script-level for export)
    $script:passedEndpointList = @{}  # Hash table to store passed endpoints by category (script-level for export)

    # Validate that endpoints were retrieved successfully
    if ($null -eq $endpointListM365Common -or $endpointListM365Common.Count -eq 0) {
        Write-Warning "No M365 Common endpoints retrieved. Connectivity test may be incomplete."
    }
    if ($null -eq $endpointListIntune -or $endpointListIntune.Count -eq 0) {
        Write-Warning "No Intune endpoints retrieved. Connectivity test may be incomplete."
    }
    if (($null -eq $endpointListM365Common -or $endpointListM365Common.Count -eq 0) -and 
        ($null -eq $endpointListIntune -or $endpointListIntune.Count -eq 0)) {
        Write-Error "Failed to retrieve any endpoints. Cannot proceed with connectivity test."
        return $false
    }

    # Display start message to user
    Write-Host "Starting Connectivity Check..." -ForegroundColor Yellow

    # Test M365 Common endpoints (required for various M365 services)
    foreach ($endpoint in $endpointListM365Common) 
    {        
        # Only test mandatory endpoints (skip optional ones)
        if ($endpoint.mandatory -eq $true) 
        {  
            # Display the category being tested
            Write-Host "Checking Category: ..." $endpoint.category -ForegroundColor Yellow
            
            # Test each URL in the endpoint category
            foreach ($url in $endpoint.urls) {
                # Test connectivity based on proxy configuration
                try {
                    if ($ProxyServer -eq "NoProxy") 
                    {
                        # Direct connection (no proxy) - test endpoint directly
                        # -UseBasicParsing: Faster, doesn't parse HTML (we only need status code)
                        # -TimeoutSec: Set timeout to prevent hanging
                        $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop).StatusCode
                    }
                    else 
                    {
                        # Proxy connection - test endpoint through proxy server
                        $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing -Proxy $ProxyServer -TimeoutSec 10 -ErrorAction Stop).StatusCode
                    }
                }
                catch {
                    # Connection failed - set TestResult to indicate failure
                    $TestResult = $null
                }
                
                # Check if the connection was successful (HTTP 200 = OK)
                # Also check if TestResult is null (connection failed)
                if ($TestResult -eq 200) 
                {
                    # Track passed endpoints for export
                    if (-not $script:passedEndpointList.ContainsKey($endpoint.category)) {
                        $script:passedEndpointList.Add($endpoint.category, @())
                    }
                    $script:passedEndpointList[$endpoint.category] += $url
                    
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        Write-Host "Connection to " $url ".............. Succeeded (needed for Asia & Pacific tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        Write-Host "Connection to " $url ".............. Succeeded (needed for Europe tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        Write-Host "Connection to " $url ".............. Succeeded (needed for North America tenants only)." -ForegroundColor Green 
                    }
                    else {
                        Write-Host "Connection to " $url ".............. Succeeded." -ForegroundColor Green 
                    }
                }
                else 
                {
                    # Connection failed - mark test as failed and record the failed endpoint
                    # TestResult is either null (exception) or not 200 (non-success status)
                    $TestFailed = $true
                    
                    # Add failed endpoint to the failed list by category
                    # This groups failures by category for easier troubleshooting
                    if ($script:failedEndpointList.ContainsKey($endpoint.category)) {
                        # Category already exists - add URL to existing array
                        $script:failedEndpointList[$endpoint.category] += $url
                    }
                    else {
                        # New category - create new entry with URL
                        $script:failedEndpointList.Add($endpoint.category, @($url))
                    }
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        Write-Host "Connection to " $url ".............. Failed (needed for Asia & Pacific tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        Write-Host "Connection to " $url ".............. Failed (needed for Europe tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        Write-Host "Connection to " $url ".............. Failed (needed for North America tenants only)." -ForegroundColor Red 
                    }
                    else {
                        Write-Host "Connection to " $url ".............. Failed." -ForegroundColor Red 
                    }
                }
            }
        }
        else 
        {
            #Write-Host "Skipping Category: ..." $endpoint.category -ForegroundColor Yellow
        }
    }

    foreach ($endpoint in $endpointListIntune) 
    {        
        if ($endpoint.mandatory -eq $true) 
        {
            Write-Host "Checking Category: ..." $endpoint.category -ForegroundColor Yellow
            foreach ($url in $endpoint.urls) {
                try {
                    if ($ProxyServer -eq "NoProxy") {
                        $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing -TimeoutSec 10 -ErrorAction Stop).StatusCode
                    }
                    else {
                        $TestResult = (Invoke-WebRequest -uri $url -UseBasicParsing -Proxy $ProxyServer -TimeoutSec 10 -ErrorAction Stop).StatusCode
                    }
                }
                catch {
                    $TestResult = $null
                }
                if ($TestResult -eq 200) {
                    # Track passed endpoints for export
                    if (-not $script:passedEndpointList.ContainsKey($endpoint.category)) {
                        $script:passedEndpointList.Add($endpoint.category, @())
                    }
                    $script:passedEndpointList[$endpoint.category] += $url
                    
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        Write-Host "Connection to " $url ".............. Succeeded (needed for Asia & Pacific tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        Write-Host "Connection to " $url ".............. Succeeded (needed for Europe tenants only)." -ForegroundColor Green 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        Write-Host "Connection to " $url ".............. Succeeded (needed for North America tenants only)." -ForegroundColor Green 
                    }
                    else {
                        Write-Host "Connection to " $url ".............. Succeeded." -ForegroundColor Green 
                    }
                }
                else {
                    # Connection failed - TestResult is either null (exception) or not 200
                    $TestFailed = $true
                    if ($script:failedEndpointList.ContainsKey($endpoint.category)) {
                        $script:failedEndpointList[$endpoint.category] += $url
                    }
                    else {
                        $script:failedEndpointList.Add($endpoint.category, @($url))
                    }
                    if (($url.StartsWith('approdimedata') -or ($url.StartsWith("intunemaape13") -or $url.StartsWith("intunemaape17") -or $url.StartsWith("intunemaape18") -or $url.StartsWith("intunemaape19")))) {
                        Write-Host "Connection to " $url ".............. Failed (needed for Asia & Pacific tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('euprodimedata') -or ($url.StartsWith("intunemaape7") -or $url.StartsWith("intunemaape8") -or $url.StartsWith("intunemaape9") -or $url.StartsWith("intunemaape10") -or $url.StartsWith("intunemaape11") -or $url.StartsWith("intunemaape12")))) {
                        Write-Host "Connection to " $url ".............. Failed (needed for Europe tenants only)." -ForegroundColor Red 
                    }
                    elseif (($url.StartsWith('naprodimedata') -or ($url.StartsWith("intunemaape1") -or $url.StartsWith("intunemaape2") -or $url.StartsWith("intunemaape3") -or $url.StartsWith("intunemaape4") -or $url.StartsWith("intunemaape5") -or $url.StartsWith("intunemaape6")))) {
                        Write-Host "Connection to " $url ".............. Failed (needed for North America tenants only)." -ForegroundColor Red 
                    }
                    else {
                        Write-Host "Connection to " $url ".............. Failed." -ForegroundColor Red 
                    }
                }
            }
        }
        else 
        {
            #Write-Host "Skipping Category: ..." $endpoint.category -ForegroundColor Yellow
        }
    }

    # Summary: Display failed endpoints if any failures occurred
    if ($TestFailed) 
    {
        # Display summary of failed endpoints grouped by category
        Write-Host "Test failed. Please check the following URLs:" -ForegroundColor Red
        foreach ($failedEndpoint in $script:failedEndpointList.Keys) {
            # Display category name
            Write-Host $failedEndpoint -ForegroundColor Red
            # Display each failed URL in that category
            foreach ($failedUrl in $script:failedEndpointList[$failedEndpoint]) {
                Write-Host $failedUrl -ForegroundColor Red
            }
        }
    }
    
    # Restore the original error action preference
    # This ensures the function doesn't permanently change error handling behavior
    $ErrorActionPreference = $originalErrorAction
    
    # Return test result: $true if all mandatory endpoints passed, $false if any failed
    if ($TestFailed) {
        Write-Host "Test-DeviceIntuneConnectivity completed with failures." -ForegroundColor Yellow -BackgroundColor Black
        return $false
    }
    else {
        Write-Host "Test-DeviceIntuneConnectivity completed successfully." -ForegroundColor Green -BackgroundColor Black
        return $true
    }
}

### Main Execution Section ###

# Execute the connectivity test function
# This tests all mandatory Intune and M365 Common endpoints
$connectivityResult = Test-DeviceIntuneConnectivity

# Display network configuration information
# This helps troubleshoot connectivity issues by showing the device's network setup
$NetworkConfiguration = @()

# Get network configuration for all network interfaces
# This includes IP addresses, gateways, DNS servers, and network profiles
Get-NetIPConfiguration | ForEach-Object {
    # Create a custom object with network configuration details
    # Handle null values gracefully to prevent errors
    $NetworkConfiguration += New-Object PSObject -Property @{
        InterfaceAlias = $_.InterfaceAlias  # Network adapter name
        ProfileName = if($null -ne $_.NetProfile.Name){$_.NetProfile.Name}else{""}  # Network profile (Domain/Private/Public)
        IPv4Address = if($null -ne $_.IPv4Address){$_.IPv4Address}else{""}  # IPv4 address
        IPv6Address = if($null -ne $_.IPv6Address){$_.IPv6Address}else{""}  # IPv6 address
        IPv4DefaultGateway = if($null -ne $_.IPv4DefaultGateway){$_.IPv4DefaultGateway.NextHop}else{""}  # IPv4 gateway
        IPv6DefaultGateway = if($null -ne $_.IPv6DefaultGateway){$_.IPv6DefaultGateway.NextHop}else{""}  # IPv6 gateway
        DNSServer = if($null -ne $_.DNSServer){$_.DNSServer.ServerAddresses}else{""}  # DNS server addresses
    }
}

# Display the network configuration in a formatted table
# This provides a clear view of the device's network setup for troubleshooting
$NetworkConfiguration | Format-Table -AutoSize

# Export results to file if path is specified
if ($ExportPath) {
    try {
        # Create directory if it doesn't exist
        $exportDirectory = Split-Path -Path $ExportPath -Parent
        if ($exportDirectory -and -not (Test-Path -Path $exportDirectory)) {
            New-Item -Path $exportDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Create export object with test results
        $exportData = @{
            TestDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            TestResult = if ($connectivityResult) { "Passed" } else { "Failed" }
            NetworkConfiguration = $NetworkConfiguration
        }
        
        # Add endpoint test results if available
        if ($script:passedEndpointList -and $script:passedEndpointList.Count -gt 0) {
            $exportData["PassedEndpoints"] = $script:passedEndpointList
        }
        if ($script:failedEndpointList -and $script:failedEndpointList.Count -gt 0) {
            $exportData["FailedEndpoints"] = $script:failedEndpointList
        }
        
        # Export to JSON format
        $exportData | ConvertTo-Json -Depth 10 | Out-File -FilePath $ExportPath -Encoding UTF8
        Write-Host "`nTest results exported to: $ExportPath" -ForegroundColor Green
        
        # Verify the export file was created successfully
        if (-not (Test-Path -Path $ExportPath)) {
            Write-Warning "Export file was not created successfully at: $ExportPath"
        }
    }
    catch {
        Write-Warning "Failed to export test results: $($_.Exception.Message)"
    }
}

# Exit with appropriate exit code based on connectivity test results
# Exit code 0 = All mandatory endpoints accessible
# Exit code 1 = One or more mandatory endpoints failed
if ($connectivityResult) {
    exit 0
} else {
    exit 1
}
