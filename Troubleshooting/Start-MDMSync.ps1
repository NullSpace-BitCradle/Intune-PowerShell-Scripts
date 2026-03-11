<#
.SYNOPSIS
    Starts an MDM sync session using Windows Management APIs.

.DESCRIPTION
    This script creates an MDM (Mobile Device Management) sync session and waits for it to complete.
    It uses the Windows.Management.MdmSessionManager API to trigger a sync between the device
    and the MDM service (e.g., Microsoft Intune). The script waits up to the specified timeout
    for the sync to complete and provides status updates during the process.

.PARAMETER TimeoutSeconds
    Optional. Maximum time in seconds to wait for sync to complete. Default is 60 seconds.

.PARAMETER CheckIntervalSeconds
    Optional. Interval in seconds between status checks. Default is 5 seconds.

.EXAMPLE
    .\Start-MDMSync.ps1

.EXAMPLE
    .\Start-MDMSync.ps1 -TimeoutSeconds 120 -CheckIntervalSeconds 10

.NOTES
    Requires: Windows 10/11 with MDM enrollment
    Exit Codes:
    - 0: Sync completed successfully
    - 1: Sync failed or ended with error state
    - 2: Sync timeout (didn't complete within timeout period)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [int]$TimeoutSeconds = 60,
    
    [Parameter(Mandatory=$false)]
    [int]$CheckIntervalSeconds = 5
)

# Main script execution wrapped in try-catch for error handling
try {
    # Load the Windows.Management.MdmSessionManager type from Windows Runtime
    # This API is used to create and manage MDM sync sessions
    # Out-Null suppresses the output from the type loading
    [Windows.Management.MdmSessionManager, Windows.Management, ContentType = WindowsRuntime] | Out-Null
    
    # Attempt to create a new MDM sync session
    # TryCreateSession() returns a session object if successful, or null if it fails
    Write-Host "Attempting to create MDM sync session..." -ForegroundColor Yellow
    $MDMSession = [Windows.Management.MdmSessionManager]::TryCreateSession()
    
    # Validate that the session was created successfully
    # If null, the device may not be enrolled in MDM or there's a configuration issue
    if ($null -eq $MDMSession) {
        Write-Error "Failed to create MDM session. Device may not be enrolled in MDM."
        exit 1
    }

    # Start the MDM sync session asynchronously
    # StartAsync() initiates the sync process but doesn't wait for completion
    Write-Host "MDM session created. Starting sync..." -ForegroundColor Yellow
    $MDMSession.StartAsync() | Out-Null

    # Wait for sync to complete with timeout protection
    # Calculate maximum number of iterations based on timeout and check interval
    $maxIterations = [math]::Ceiling($TimeoutSeconds / $CheckIntervalSeconds)
    $MDMSyncTimeout = 0  # Counter for iterations
    $lastState = $null
    $unchangedCount = 0

    # Loop until sync completes or timeout is reached
    # Check session state at specified intervals
    do {
        # Wait for the specified interval before checking status again
        Start-Sleep -Seconds $CheckIntervalSeconds
        # Increment the iteration counter
        $MDMSyncTimeout += 1
        
        # Check if session state has changed
        if ($MDMSession.State -eq $lastState) {
            $unchangedCount++
            if ($unchangedCount -gt 5) {
                Write-Warning "Session state unchanged for extended period. Current state: $($MDMSession.State)"
                break
            }
        }
        else {
            $unchangedCount = 0
        }
        $lastState = $MDMSession.State
        
        # Display progress to the user
        Write-Host "Waiting for sync to complete... ($($MDMSyncTimeout * $CheckIntervalSeconds)/$TimeoutSeconds seconds) - State: $($MDMSession.State)" -ForegroundColor Cyan
    } while (($MDMSession.State -ne "Completed") -and ($MDMSyncTimeout -lt $maxIterations))

    # Check the final status of the sync session
    if ($MDMSession.State -eq "Completed") {
        # Sync completed successfully
        Write-Host "MDM sync completed successfully." -ForegroundColor Green
        exit 0
    }
    elseif ($MDMSession.State -eq "Error" -or $MDMSession.State -eq "Failed") {
        # Sync ended in an explicit error state
        Write-Error "MDM sync failed with state: $($MDMSession.State)"
        exit 1
    }
    elseif ($MDMSyncTimeout -ge $maxIterations) {
        # Sync did not complete within the timeout period
        # This may indicate network issues or MDM service problems
        Write-Warning "MDM sync did not complete within the timeout period. Current state: $($MDMSession.State)"
        exit 2
    }
    else {
        # Sync ended in an unexpected state (not Completed, Error, Failed, or timeout)
        Write-Warning "MDM sync ended with unexpected state: $($MDMSession.State)"
        exit 1
    }
}
# Error handling block - catches exceptions during session creation or sync process
catch {
    # Display the error message with script name for context
    Write-Error "Start-MDMSync: Failed to start MDM sync - $($_.Exception.Message)"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    exit 1
}
