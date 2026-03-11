<#
.SYNOPSIS
    Retrieves all Intune policy assignments (configuration profiles, compliance policies).

.DESCRIPTION
    This script connects to Microsoft Graph API and retrieves all Intune policy assignments
    including configuration profiles and compliance policies. It shows which policies are
    assigned to which groups or users, assignment filters, and assignment intent.

.PARAMETER AppId
    The Application (Client) ID of your Entra ID registered app. Can also be provided via
    INTUNE_APP_ID environment variable.

.PARAMETER TenantId
    Your Azure AD Tenant ID. Can also be provided via INTUNE_TENANT_ID environment variable.

.PARAMETER ClientSecret
    The Client Secret of your Entra ID registered app. Can also be provided via
    INTUNE_CLIENT_SECRET environment variable.

.PARAMETER PolicyType
    Optional. Filter by policy type: "Configuration" or "Compliance". If not specified, retrieves both.

.PARAMETER ExportPath
    Optional. Path to export results to CSV file. If not specified, results are displayed only.

.EXAMPLE
    .\Get-IntunePolicyAssignments.ps1 -AppId 'your-app-id' -TenantId 'your-tenant-id' -ClientSecret 'your-secret'

.EXAMPLE
    .\Get-IntunePolicyAssignments.ps1 -PolicyType "Configuration" -ExportPath 'C:\Temp\PolicyAssignments.csv'

.NOTES
    Requires: Microsoft.Graph PowerShell module
    Requires: Entra ID App Registration with DeviceManagementConfiguration.Read.All permission
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
    [ValidateSet("Configuration", "Compliance", "All")]
    [string]$PolicyType = "All",
    
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
    
    # Initialize array to store policy assignment results
    $policyAssignments = @()
    
    # Query configuration profiles if requested
    # Configuration profiles define device settings and restrictions
    if ($PolicyType -eq "All" -or $PolicyType -eq "Configuration") {
        Write-Host "`nQuerying configuration profiles..." -ForegroundColor Yellow
        
        # Query all configuration profiles from Intune
        $configProfilesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations"
        $configProfiles = Invoke-GraphApiWithRetry -Uri $configProfilesUrl -Headers $headers -Method Get
        
        # Handle null or empty response
        $allConfigProfiles = @()
        if ($null -ne $configProfiles -and $null -ne $configProfiles.value) {
            $allConfigProfiles += $configProfiles.value
        }
        $pageCount = 1
        
        # Check if there are more pages of results
        while ($null -ne $configProfiles -and $null -ne $configProfiles.'@odata.nextLink') {
            $pageCount++
            Write-Progress -Activity "Retrieving configuration profiles" -Status "Fetching page $pageCount" -PercentComplete -1
            $configProfiles = Invoke-GraphApiWithRetry -Uri $configProfiles.'@odata.nextLink' -Headers $headers -Method Get
            if ($null -ne $configProfiles -and $null -ne $configProfiles.value) {
                $allConfigProfiles += $configProfiles.value
            }
        }
        Write-Progress -Activity "Retrieving configuration profiles" -Completed
        
        # Process each configuration profile
        $profileCount = 0
        foreach ($configProfile in $allConfigProfiles) {
            $profileCount++
            Write-Progress -Activity "Processing configuration profiles" -Status "Processing profile $profileCount of $($allConfigProfiles.Count): $($configProfile.displayName)" -PercentComplete (($profileCount / $allConfigProfiles.Count) * 100)
            Write-Host "  Processing configuration profile: $($configProfile.displayName)" -ForegroundColor Cyan
            
            # Get assignments for this configuration profile
            $assignmentsUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations('$($configProfile.id)')/assignments"
            try {
                $assignments = Invoke-GraphApiWithRetry -Uri $assignmentsUrl -Headers $headers -Method Get
                
                # Handle null or empty response
                if ($null -eq $assignments -or $null -eq $assignments.value) {
                    # No assignments found, create entry with no assignments
                    $policyAssignment = [PSCustomObject]@{
                        PolicyType = "Configuration"
                        PolicyId = $configProfile.id
                        PolicyName = $configProfile.displayName
                        PolicyDescription = $configProfile.description
                        AssignmentId = $null
                        TargetType = "None"
                        TargetId = $null
                        TargetName = "No assignments"
                        Intent = $null
                        Source = $null
                    }
                    $policyAssignments += $policyAssignment
                    continue
                }
                
                foreach ($assignment in $assignments.value) {
                    # Get target group information
                    $targetGroupId = $assignment.target.groupId
                    $targetName = $assignment.target.'@odata.type'
                    
                    # Resolve group name if it's a group assignment
                    if ($targetGroupId) {
                        try {
                            $groupUrl = "https://graph.microsoft.com/v1.0/groups/$targetGroupId"
                            $group = Invoke-GraphApiWithRetry -Uri $groupUrl -Headers $headers -Method Get
                            if ($null -ne $group -and $null -ne $group.displayName) {
                                $targetName = $group.displayName
                            }
                            else {
                                $targetName = $targetGroupId
                                Write-Verbose "Could not retrieve group name for Group ID: $targetGroupId"
                            }
                        }
                        catch {
                            $targetName = $targetGroupId
                            Write-Verbose "Failed to retrieve group name for Group ID: $targetGroupId - $($_.Exception.Message)"
                        }
                    }
                    
                    # Create policy assignment object
                    $policyAssignment = [PSCustomObject]@{
                        PolicyType = "Configuration"
                        PolicyId = $configProfile.id
                        PolicyName = $configProfile.displayName
                        PolicyDescription = $configProfile.description
                        AssignmentId = $assignment.id
                        TargetType = $assignment.target.'@odata.type'
                        TargetId = $targetGroupId
                        TargetName = $targetName
                        Intent = $assignment.intent
                        Source = $assignment.source
                    }
                    
                    $policyAssignments += $policyAssignment
                }
            }
            catch {
                # If API call failed, create entry with error information
                Write-Verbose "Failed to retrieve assignments for configuration profile '$($configProfile.displayName)': $($_.Exception.Message)"
                $policyAssignment = [PSCustomObject]@{
                    PolicyType = "Configuration"
                    PolicyId = $configProfile.id
                    PolicyName = $configProfile.displayName
                    PolicyDescription = $configProfile.description
                    AssignmentId = $null
                    TargetType = "Error"
                    TargetId = $null
                    TargetName = "Error retrieving assignments: $($_.Exception.Message)"
                    Intent = $null
                    Source = $null
                }
                
                $policyAssignments += $policyAssignment
            }
        }
        Write-Progress -Activity "Processing configuration profiles" -Completed
    }
    
    # Query compliance policies if requested
    # Compliance policies define requirements devices must meet
    if ($PolicyType -eq "All" -or $PolicyType -eq "Compliance") {
        Write-Host "`nQuerying compliance policies..." -ForegroundColor Yellow
        
        # Query all compliance policies from Intune
        $compliancePoliciesUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies"
        $compliancePolicies = Invoke-GraphApiWithRetry -Uri $compliancePoliciesUrl -Headers $headers -Method Get
        
        # Handle null or empty response
        $allCompliancePolicies = @()
        if ($null -ne $compliancePolicies -and $null -ne $compliancePolicies.value) {
            $allCompliancePolicies += $compliancePolicies.value
        }
        $pageCount = 1
        
        # Check if there are more pages of results
        while ($null -ne $compliancePolicies -and $null -ne $compliancePolicies.'@odata.nextLink') {
            $pageCount++
            Write-Progress -Activity "Retrieving compliance policies" -Status "Fetching page $pageCount" -PercentComplete -1
            $compliancePolicies = Invoke-GraphApiWithRetry -Uri $compliancePolicies.'@odata.nextLink' -Headers $headers -Method Get
            if ($null -ne $compliancePolicies -and $null -ne $compliancePolicies.value) {
                $allCompliancePolicies += $compliancePolicies.value
            }
        }
        Write-Progress -Activity "Retrieving compliance policies" -Completed
        
        # Process each compliance policy and get its assignments
        $policyCount = 0
        foreach ($policy in $allCompliancePolicies) {
            $policyCount++
            Write-Progress -Activity "Processing compliance policies" -Status "Processing policy $policyCount of $($allCompliancePolicies.Count): $($policy.displayName)" -PercentComplete (($policyCount / $allCompliancePolicies.Count) * 100)
            Write-Host "  Processing compliance policy: $($policy.displayName)" -ForegroundColor Cyan
            
            # Get assignments for this compliance policy
            $assignmentsUrl = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies('$($policy.id)')/assignments"
            try {
                $assignments = Invoke-GraphApiWithRetry -Uri $assignmentsUrl -Headers $headers -Method Get
                
                # Handle null or empty response
                if ($null -eq $assignments -or $null -eq $assignments.value) {
                    # No assignments found, create entry with no assignments
                    $policyAssignment = [PSCustomObject]@{
                        PolicyType = "Compliance"
                        PolicyId = $policy.id
                        PolicyName = $policy.displayName
                        PolicyDescription = $policy.description
                        AssignmentId = $null
                        TargetType = "None"
                        TargetId = $null
                        TargetName = "No assignments"
                        Intent = $null
                        Source = $null
                    }
                    $policyAssignments += $policyAssignment
                    continue
                }
                
                foreach ($assignment in $assignments.value) {
                    # Get target group information
                    $targetGroupId = $assignment.target.groupId
                    $targetName = $assignment.target.'@odata.type'
                    
                    # Resolve group name if it's a group assignment
                    if ($targetGroupId) {
                        try {
                            $groupUrl = "https://graph.microsoft.com/v1.0/groups/$targetGroupId"
                            $group = Invoke-GraphApiWithRetry -Uri $groupUrl -Headers $headers -Method Get
                            if ($null -ne $group -and $null -ne $group.displayName) {
                                $targetName = $group.displayName
                            }
                            else {
                                $targetName = $targetGroupId
                                Write-Verbose "Could not retrieve group name for Group ID: $targetGroupId"
                            }
                        }
                        catch {
                            $targetName = $targetGroupId
                            Write-Verbose "Failed to retrieve group name for Group ID: $targetGroupId - $($_.Exception.Message)"
                        }
                    }
                    
                    # Create policy assignment object
                    $policyAssignment = [PSCustomObject]@{
                        PolicyType = "Compliance"
                        PolicyId = $policy.id
                        PolicyName = $policy.displayName
                        PolicyDescription = $policy.description
                        AssignmentId = $assignment.id
                        TargetType = $assignment.target.'@odata.type'
                        TargetId = $targetGroupId
                        TargetName = $targetName
                        Intent = $assignment.intent
                        Source = $assignment.source
                    }
                    
                    $policyAssignments += $policyAssignment
                }
            }
            catch {
                # If API call failed, create entry with error information
                Write-Verbose "Failed to retrieve assignments for compliance policy '$($policy.displayName)': $($_.Exception.Message)"
                $policyAssignment = [PSCustomObject]@{
                    PolicyType = "Compliance"
                    PolicyId = $policy.id
                    PolicyName = $policy.displayName
                    PolicyDescription = $policy.description
                    AssignmentId = $null
                    TargetType = "Error"
                    TargetId = $null
                    TargetName = "Error retrieving assignments: $($_.Exception.Message)"
                    Intent = $null
                    Source = $null
                }
                
                $policyAssignments += $policyAssignment
            }
        }
        Write-Progress -Activity "Processing compliance policies" -Completed
    }
    
    # Display results
    Write-Host "`nPolicy Assignments:" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    
    if ($policyAssignments.Count -eq 0) {
        Write-Host "No policy assignments found." -ForegroundColor Yellow
    }
    else {
        # Group results by policy type for better display
        $configAssignments = $policyAssignments | Where-Object { $_.PolicyType -eq "Configuration" }
        $complianceAssignments = $policyAssignments | Where-Object { $_.PolicyType -eq "Compliance" }
        
        if ($configAssignments.Count -gt 0) {
            Write-Host "`nConfiguration Profiles: $($configAssignments.Count) assignment(s)" -ForegroundColor Green
            foreach ($assignment in $configAssignments) {
                Write-Host "  Policy: $($assignment.PolicyName)" -ForegroundColor White
                Write-Host "    Target: $($assignment.TargetName) ($($assignment.TargetType))" -ForegroundColor Gray
                Write-Host "    Intent: $($assignment.Intent)" -ForegroundColor Gray
            }
        }
        
        if ($complianceAssignments.Count -gt 0) {
            Write-Host "`nCompliance Policies: $($complianceAssignments.Count) assignment(s)" -ForegroundColor Green
            foreach ($assignment in $complianceAssignments) {
                Write-Host "  Policy: $($assignment.PolicyName)" -ForegroundColor White
                Write-Host "    Target: $($assignment.TargetName) ($($assignment.TargetType))" -ForegroundColor Gray
                Write-Host "    Intent: $($assignment.Intent)" -ForegroundColor Gray
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
            $policyAssignments | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "`nResults exported to: $ExportPath" -ForegroundColor Green
            
            # Verify the export file was created successfully
            if (-not (Test-Path -Path $ExportPath)) {
                Write-Error "Export file was not created successfully at: $ExportPath"
                exit 1
            }
        }
    }
    
    Write-Host "`nQuery completed successfully. Found $($policyAssignments.Count) policy assignment(s)." -ForegroundColor Green
}
# Error handling block - catches exceptions during Graph API calls
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Get-IntunePolicyAssignments: Failed to retrieve policy assignments - $errMsg"
    
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

