#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Enables smart card logon enforcement on Windows devices.

.DESCRIPTION
    This script enables smart card logon enforcement by modifying registry keys that control
    smart card authentication behavior. It sets registry values to require smart card for
    interactive logon and enables automatic lock on smart card removal. The script also enables
    the Smart Card Policy Service.

.PARAMETER KeepLogs
    Keeps the log file after script execution instead of removing it.
    Useful for troubleshooting purposes.

.EXAMPLE
    .\Enable-SmartCardLogonEnforcement.ps1

.EXAMPLE
    .\Enable-SmartCardLogonEnforcement.ps1 -WhatIf

.EXAMPLE
    .\Enable-SmartCardLogonEnforcement.ps1 -KeepLogs

.NOTES
    Requires: Administrator privileges
    Modifies: HKLM registry keys
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Mandatory=$false)]
    [switch]$KeepLogs
)

# Architecture check: If running as a 32-bit process on an x64 system, re-launch as a 64-bit process
# This ensures registry modifications are made to the correct registry hive (64-bit vs 32-bit)
# PROCESSOR_ARCHITEW6432 is set on 64-bit systems when running 32-bit processes
if ($env:PROCESSOR_ARCHITEW6432 -and $env:PROCESSOR_ARCHITEW6432 -ne "ARM64")
{
    # Check if the 64-bit PowerShell executable exists in the SysNative path
    # SysNative is a special path that allows 32-bit processes to access 64-bit system files
    if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe")
    {
        # Re-launch the script in 64-bit PowerShell
        # -ExecutionPolicy bypass: Allows script execution regardless of execution policy
        # -NoProfile: Starts PowerShell without loading user profile (faster startup)
        # -File: Specifies the script file to execute
        & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypass -NoProfile -File "$PSCommandPath"
        # Exit with the same exit code as the re-launched script
        exit $LASTEXITCODE
    }
}
    
# Logging Preparation Section
# Set up variables for logging script execution to a log file

# Define the application name for log file naming
$AppName = "Enable_SmartCardLogon_Enforcement"

# Construct the log file name using Intune naming convention
$Log_FileName = "win32-$AppName.log"

# Define the log file path (Intune Management Extension logs directory)
# Use environment variable for cross-system compatibility
$Log_Path = Join-Path $env:ProgramData "Microsoft\IntuneManagementExtension\Logs"

# Construct the full path to the log file
$TestPath = "$Log_Path\$Log_FileName"

# Define a section line separator for log readability (asterisks repeated 10 times)
$SectionLine="* * "*10

# Check if the log file exists, create it if it doesn't
# This ensures logging can proceed even if the file doesn't exist
If(!(Test-Path $TestPath))
{
    # Create the log file in the Intune logs directory
    New-Item -Path $Log_Path -Name $Log_FileName -ItemType "File" -Force
}

# Define a logging function to write timestamped messages to the log file
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Message  # The message to write to the log
    )
    # Get current date and time formatted as "DayOfWeek MM/dd/yyyy HH:mm:ss"
    $timestamp = Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss"
    # Append the timestamped message to the log file
    Add-Content -Path $TestPath -Value "$timestamp : $Message"
}

# Start logging - write initial log entries
# This log file can also be used by Intune Management Extension for detection
Write-Log "Begin..."
Write-Log $SectionLine

# Section 1: Enable Smart Card Logon Enforcement
# This section modifies the registry to require smart card for interactive logon

try
{
    # Check if the System policies registry path exists
    # This path contains policy settings including smart card enforcement
    if (-NOT (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System')) {
        # Registry path doesn't exist - create it
        Write-Log "Reg path not found, hence creating Reg path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Force -ErrorAction SilentlyContinue | Out-Null
    }
    
    # Set the scforceoption registry value to 1
    # Value 0 = Allow password logon (smart card not required)
    # Value 1 = Require smart card for interactive logon
    if ($PSCmdlet.ShouldProcess('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\scforceoption', 'Set value to 1 (require smart card for interactive logon)')) {
        Write-Log "Reg path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System found, updating Reg key scforceoption with value 1"
        New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System' -Name 'scforceoption' -Value 1 -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Reg key value updated"
    }
}
catch 
{
    # Log any errors that occur during registry modification with enhanced details
    $errMsg = $_.Exception.Message
    Write-Log "Error: $errMsg"
    if ($_.Exception.InnerException) {
        Write-Log "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    Write-Log "Stack Trace: $($_.ScriptStackTrace)"
}

# Section 2: Enable Auto-lock upon Smart Card Removal
# This section modifies the registry to automatically lock the screen when smart card is removed

try
{
    # Check if the Winlogon registry path exists
    # This path contains Windows logon configuration settings
    if (-NOT (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon')) {
        # Registry path doesn't exist - create it
        Write-Log "Reg path not found, hence creating Reg path HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
        New-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Force -ErrorAction SilentlyContinue | Out-Null
    }
    
    # Set the ScRemoveOption registry value to 1
    # Value 0 = Do not lock workstation when smart card is removed
    # Value 1 = Lock workstation when smart card is removed
    if ($PSCmdlet.ShouldProcess('HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\ScRemoveOption', 'Set value to 1 (lock workstation on smart card removal)')) {
        Write-Log "Reg path HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon found, updating Reg key ScRemoveOption with value 1"
        New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'ScRemoveOption' -Value 1 -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Reg key value updated"
    }
}
catch 
{
    # Log any errors that occur during registry modification with enhanced details
    $errMsg = $_.Exception.Message
    Write-Log "Error: $errMsg"
    if ($_.Exception.InnerException) {
        Write-Log "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    Write-Log "Stack Trace: $($_.ScriptStackTrace)"
}

# Section 3: Enable Smart Card Policy Service
# This section enables the Smart Card Policy Service by setting its startup type to Automatic

try
{
    # Check if the Smart Card Policy Service registry path exists
    # This path contains service configuration for SCPolicySvc
    if (-NOT (Test-Path -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Services\SCPolicySvc')) {
        # Registry path doesn't exist - create it
        Write-Log "Reg path not found, hence creating Reg path HKLM:\SYSTEM\CurrentControlSet\Services\SCPolicySvc"
        New-Item -Path 'HKLM:\SYSTEM\CurrentControlSet\Services\SCPolicySvc' -Force -ErrorAction SilentlyContinue | Out-Null
    }
    
    # Set the Start registry value to 2 (Automatic)
    # Service startup types: 0=Boot, 1=System, 2=Auto, 3=Manual, 4=Disabled
    # Setting to 2 (Automatic) ensures the service starts automatically on boot
    if ($PSCmdlet.ShouldProcess('HKLM:\SYSTEM\CurrentControlSet\Services\SCPolicySvc', 'Set Start value to 2 (Automatic) and enable SCPolicySvc')) {
        Write-Log "Reg path HKLM:\SYSTEM\CurrentControlSet\Services\SCPolicySvc found, updating Reg key Start with value 2"
        New-ItemProperty -LiteralPath 'HKLM:\SYSTEM\CurrentControlSet\Services\SCPolicySvc' -Name 'Start' -Value 2 -Force -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Reg key value updated"

        # Also set the service startup type using Set-Service to ensure immediate effect
        try {
            Set-Service -Name "SCPolicySvc" -StartupType Automatic -ErrorAction SilentlyContinue
            Write-Log "Service startup type set to Automatic using Set-Service"
        }
        catch {
            Write-Log "Warning: Could not set service startup type using Set-Service: $($_.Exception.Message)"
        }
    }
}
catch 
{
    # Log any errors that occur during registry modification with enhanced details
    $errMsg = $_.Exception.Message
    Write-Log "Error: $errMsg"
    if ($_.Exception.InnerException) {
        Write-Log "Inner Exception: $($_.Exception.InnerException.Message)"
    }
    Write-Log "Stack Trace: $($_.ScriptStackTrace)"
}

# Cleanup: Remove the log file after script execution
# This is done to clean up temporary log files
# -ErrorAction SilentlyContinue prevents errors if the file doesn't exist
# Only remove if -KeepLogs is not specified
if (-not $WhatIfPreference -and -not $KeepLogs) {
    $logFilePath = Join-Path $Log_Path $Log_FileName
    Remove-Item -Path $logFilePath -Force -ErrorAction SilentlyContinue
}
elseif ($KeepLogs) {
    $logFilePath = Join-Path $Log_Path $Log_FileName
    Write-Log "Log file kept at: $logFilePath"
}

# Log completion and exit with success code
Write-Log "Script completed successfully."
exit 0
