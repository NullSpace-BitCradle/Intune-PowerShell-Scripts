#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Collects Intune-related logs from device.

.DESCRIPTION
    This script collects Intune-related logs from the device including logs from
    Intune Management Extension, MDM enrollment, device management, and other
    Intune-related components. The logs are collected and optionally exported
    to a specified location.

.PARAMETER OutputPath
    Optional. Path to export collected logs. If not specified, logs are collected
    to a timestamped folder in $env:TEMP\IntuneLogs.

.PARAMETER IncludeEventLogs
    Optional switch. If specified, includes Windows Event Logs related to Intune.

.EXAMPLE
    .\Get-IntuneDeviceLogs.ps1

.EXAMPLE
    .\Get-IntuneDeviceLogs.ps1 -OutputPath 'C:\Logs\Intune' -IncludeEventLogs

.NOTES
    Requires: Administrator privileges
    Exit Codes:
    - 0: Log collection successful
    - 1: Log collection failed or error occurred
#>

[CmdletBinding()]
# Define script parameters for output path and options
# These allow customization of log collection behavior
param(
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [switch]$IncludeEventLogs
)

# Main script execution wrapped in try-catch for error handling
try
{
    # Check if running as administrator
    # Log collection requires administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "This script requires administrator privileges. Please run as administrator."
        exit 1
    }

    Write-Host "Collecting Intune-related logs..." -ForegroundColor Yellow
    Write-Host "===================================" -ForegroundColor Yellow
    
    # Check for required tools
    Write-Verbose "Checking for required tools..."
    $requiredTools = @('reg', 'wevtutil')
    $missingTools = @()
    foreach ($tool in $requiredTools) {
        if (-not (Get-Command $tool -ErrorAction SilentlyContinue)) {
            $missingTools += $tool
        }
    }
    if ($missingTools.Count -gt 0) {
        Write-Warning "Some required tools are not available: $($missingTools -join ', '). Some log collection features may be unavailable."
    }
    
    # Determine output path
    # If not specified, create a timestamped folder in $env:TEMP\IntuneLogs
    if (-not $OutputPath) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $OutputPath = Join-Path $env:TEMP "IntuneLogs\$timestamp"
    }
    
    # Create output directory if it doesn't exist
    # This ensures the directory exists before collecting logs
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $OutputPath" -ForegroundColor Green
    }
    
    Write-Host "Output path: $OutputPath" -ForegroundColor Cyan
    
    # Initialize log collection summary
    # This tracks which logs were successfully collected
    $collectedLogs = @()
    $failedLogs = @()
    
    # Collect Intune Management Extension logs
    # These logs contain information about Intune Management Extension operations
    Write-Host "`n[1/6] Collecting Intune Management Extension logs..." -ForegroundColor Cyan
    Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting Intune Management Extension logs..." -PercentComplete 0
    try {
        $imeLogPath = "${env:ProgramData}\Microsoft\IntuneManagementExtension\Logs"
        if (Test-Path -Path $imeLogPath) {
            $imeOutputPath = Join-Path $OutputPath "IntuneManagementExtension"
            New-Item -Path $imeOutputPath -ItemType Directory -Force | Out-Null
            
            # Get log files and check sizes before copying
            $logFiles = Get-ChildItem -Path $imeLogPath -Recurse -File -ErrorAction SilentlyContinue
            $totalSize = ($logFiles | Measure-Object -Property Length -Sum).Sum / 1MB
            if ($totalSize -gt 100) {
                Write-Warning "Intune Management Extension logs total size is $([math]::Round($totalSize, 2)) MB. This may take a while to copy."
            }
            
            # Copy all log files from IME log directory
            # These logs contain detailed information about IME operations
            Copy-Item -Path "$imeLogPath\*" -Destination $imeOutputPath -Recurse -Force -ErrorAction Stop
            $collectedLogs += "Intune Management Extension logs"
            Write-Host "  Successfully collected Intune Management Extension logs." -ForegroundColor Green
        }
        else {
            Write-Host "  Intune Management Extension log path not found: $imeLogPath" -ForegroundColor Yellow
        }
    }
    catch {
        $failedLogs += "Intune Management Extension logs: $($_.Exception.Message)"
        Write-Host "  Failed to collect Intune Management Extension logs: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting Intune Management Extension logs..." -Completed
    
    # Collect MDM enrollment and device management logs
    # Note: Both enrollment and management logs are stored in the same directory
    # We collect them once and label appropriately to avoid duplication
    Write-Host "`n[2/6] Collecting MDM enrollment and device management logs..." -ForegroundColor Cyan
    Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting MDM enrollment and device management logs..." -PercentComplete 20
    try {
        $mdmLogPath = "${env:ProgramData}\Microsoft\Windows\DeviceManagement\Logs"
        if (Test-Path -Path $mdmLogPath) {
            $mdmOutputPath = Join-Path $OutputPath "MDMEnrollmentAndManagement"
            New-Item -Path $mdmOutputPath -ItemType Directory -Force | Out-Null
            
            # Get log files and check sizes before copying
            $logFiles = Get-ChildItem -Path $mdmLogPath -Recurse -File -ErrorAction SilentlyContinue
            $totalSize = ($logFiles | Measure-Object -Property Length -Sum).Sum / 1MB
            if ($totalSize -gt 100) {
                Write-Warning "MDM logs total size is $([math]::Round($totalSize, 2)) MB. This may take a while to copy."
            }
            
            # Copy all log files from MDM log directory
            # These logs contain both enrollment and device management information
            # Both are stored in the same directory, so we collect them together
            Copy-Item -Path "$mdmLogPath\*" -Destination $mdmOutputPath -Recurse -Force -ErrorAction Stop
            $collectedLogs += "MDM enrollment and device management logs"
            Write-Host "  Successfully collected MDM enrollment and device management logs." -ForegroundColor Green
        }
        else {
            Write-Host "  MDM log path not found: $mdmLogPath" -ForegroundColor Yellow
        }
    }
    catch {
        $failedLogs += "MDM enrollment and device management logs: $($_.Exception.Message)"
        Write-Host "  Failed to collect MDM logs: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting MDM enrollment and device management logs..." -Completed
    
    # Collect enrollment registry information
    # Export registry keys related to enrollment for troubleshooting
    Write-Host "`n[3/6] Collecting enrollment registry information..." -ForegroundColor Cyan
    Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting enrollment registry information..." -PercentComplete 40
    try {
        $regOutputPath = Join-Path $OutputPath "Registry"
        New-Item -Path $regOutputPath -ItemType Directory -Force | Out-Null
        
        # Export enrollment registry keys
        # These registry keys contain enrollment configuration and status
        $enrollmentPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
        if (Test-Path -Path $enrollmentPath) {
            if (Get-Command reg -ErrorAction SilentlyContinue) {
                $regFile = Join-Path $regOutputPath "Enrollments.reg"
                reg export "HKLM\SOFTWARE\Microsoft\Enrollments" $regFile /y | Out-Null
                if (Test-Path -Path $regFile) {
                    $collectedLogs += "Enrollment registry information"
                    Write-Host "  Successfully collected enrollment registry information." -ForegroundColor Green
                }
            }
            else {
                Write-Host "  reg.exe not available. Skipping registry export." -ForegroundColor Yellow
            }
        }
        
        # Note: MDM registry keys are the same as enrollment keys, so we don't need to export twice
    }
    catch {
        $failedLogs += "Enrollment registry information: $($_.Exception.Message)"
        Write-Host "  Failed to collect enrollment registry information: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting enrollment registry information..." -Completed
    
    # Collect Windows Event Logs if requested
    # These logs contain Windows events related to Intune and MDM
    if ($IncludeEventLogs) {
        Write-Host "`n[4/6] Collecting Windows Event Logs..." -ForegroundColor Cyan
        Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting Windows Event Logs..." -PercentComplete 60
        try {
            $eventLogOutputPath = Join-Path $OutputPath "EventLogs"
            New-Item -Path $eventLogOutputPath -ItemType Directory -Force | Out-Null
            
            # Export relevant event logs
            # These event logs contain Windows events related to Intune operations
            $eventLogs = @("Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider", 
                          "Microsoft-Windows-User Device Registration/Admin",
                          "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin")
            
            if (Get-Command wevtutil -ErrorAction SilentlyContinue) {
                foreach ($logName in $eventLogs) {
                    try {
                        # Check if event log exists
                        if (Get-WinEvent -ListLog $logName -ErrorAction SilentlyContinue) {
                            $logFileName = $logName -replace '[\\/]', '_'
                            $logFile = Join-Path $eventLogOutputPath "$logFileName.evtx"
                            
                            # Export event log
                            wevtutil epl $logName $logFile | Out-Null
                            if (Test-Path -Path $logFile) {
                                Write-Host "  Exported event log: $logName" -ForegroundColor Green
                            }
                        }
                    }
                    catch {
                        Write-Verbose "Could not export event log: $logName - $($_.Exception.Message)"
                        Write-Host "  Could not export event log: $logName" -ForegroundColor Yellow
                    }
                }
            }
            else {
                Write-Host "  wevtutil.exe not available. Skipping event log export." -ForegroundColor Yellow
            }
            
            $collectedLogs += "Windows Event Logs"
            Write-Host "  Successfully collected Windows Event Logs." -ForegroundColor Green
        }
        catch {
            $failedLogs += "Windows Event Logs: $($_.Exception.Message)"
            Write-Host "  Failed to collect Windows Event Logs: $($_.Exception.Message)" -ForegroundColor Red
        }
        Write-Progress -Activity "Collecting Intune Logs" -Status "Collecting Windows Event Logs..." -Completed
    }
    
    # Create log collection summary
    # This provides a summary of what was collected
    Write-Host "`n[5/6] Creating log collection summary..." -ForegroundColor Cyan
    Write-Progress -Activity "Collecting Intune Logs" -Status "Creating log collection summary..." -PercentComplete 80
    try {
        $summaryPath = Join-Path $OutputPath "LogCollectionSummary.txt"
        $summary = @"
Intune Log Collection Summary
=============================
Collection Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Output Path: $OutputPath

Collected Logs:
$($collectedLogs -join "`n")

Failed Collections:
$($failedLogs -join "`n")

System Information:
- Computer Name: $env:COMPUTERNAME
- OS Version: $((Get-CimInstance Win32_OperatingSystem).Version)
- OS Build: $((Get-CimInstance Win32_OperatingSystem).BuildNumber)
"@
        
        $summary | Out-File -FilePath $summaryPath -Encoding UTF8
        $collectedLogs += "Log collection summary"
        Write-Host "  Successfully created log collection summary." -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to create log collection summary: $($_.Exception.Message)" -ForegroundColor Red
    }
    Write-Progress -Activity "Collecting Intune Logs" -Status "Creating log collection summary..." -Completed
    
    # Display summary
    Write-Host "`n===================================" -ForegroundColor Yellow
    Write-Host "Log Collection Summary" -ForegroundColor Cyan
    Write-Host "===================================" -ForegroundColor Yellow
    Write-Host "Output Path: $OutputPath" -ForegroundColor White
    Write-Host "Collected: $($collectedLogs.Count) log type(s)" -ForegroundColor Green
    if ($failedLogs.Count -gt 0) {
        Write-Host "Failed: $($failedLogs.Count) collection(s)" -ForegroundColor Red
        foreach ($failed in $failedLogs) {
            Write-Host "  - $failed" -ForegroundColor Red
        }
    }
    
    Write-Host "`nLog collection completed successfully." -ForegroundColor Green
    exit 0
}
# Error handling block - catches exceptions during log collection
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Get-IntuneDeviceLogs: Failed to collect Intune logs - $errMsg"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    exit 1
}

