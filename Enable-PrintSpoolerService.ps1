#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Enables the Print Spooler service and sets it to start automatically.

.DESCRIPTION
    This script enables the Print Spooler service by setting its startup type to Automatic
    and starting the service if it's not already running. This is typically used to restore
    printing functionality after the service has been disabled for security reasons.

.EXAMPLE
    .\Enable-PrintSpoolerService.ps1

.EXAMPLE
    .\Enable-PrintSpoolerService.ps1 -WhatIf

.NOTES
    Requires: Administrator privileges
#>

[CmdletBinding(SupportsShouldProcess=$true)]
param()

# Main script execution wrapped in try-catch for error handling
try {
    # Check if running as administrator
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "This script requires administrator privileges. Please run as administrator."
        exit 1
    }

    # Check if the service exists
    $service = Get-Service -Name "Spooler" -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-Error "Print Spooler service not found on this system."
        exit 1
    }

    Write-Host "Current Print Spooler service status: $($service.Status)" -ForegroundColor Yellow
    Write-Host "Current startup type: $($service.StartType)" -ForegroundColor Yellow

    # Check if the service is currently running
    if ($service.Status -eq "Running")
    {
        # Service is already running - just set the startup type to Automatic
        # This ensures the service will start automatically on system boot
        if ($PSCmdlet.ShouldProcess("Print Spooler service", "Set startup type to Automatic")) {
            Write-Host "Print Spooler service is already running. Setting startup type to Automatic..." -ForegroundColor Yellow
            Set-Service -name "Spooler" -startupType "Automatic" -ErrorAction Stop
            Write-Host "Print Spooler startup type set to Automatic successfully." -ForegroundColor Green
        }
    }
    else
    {
        # Service is not running - need to set startup type and then start it
        # First, set the startup type to Automatic so it will start on boot
        if ($PSCmdlet.ShouldProcess("Print Spooler service", "Set startup type to Automatic and start service")) {
            Write-Host "Setting Print Spooler startup type to Automatic..." -ForegroundColor Yellow
            Set-Service -name "Spooler" -startupType "Automatic" -ErrorAction Stop
            Write-Host "Print Spooler startup type set to Automatic successfully." -ForegroundColor Green

            # Wait 10 seconds to allow the service configuration to be applied
            # This ensures the service is ready to be started
            Write-Host "Waiting for service configuration to be applied..." -ForegroundColor Yellow
            Start-Sleep -Seconds 10

            # Start the Print Spooler service
            # This makes the service immediately available without requiring a reboot
            Write-Host "Starting Print Spooler service..." -ForegroundColor Yellow
            Start-Service -Name "Spooler" -ErrorAction Stop
            Write-Host "Print Spooler service started successfully." -ForegroundColor Green
        }
    }

    # Verify the change
    $service = Get-Service -Name "Spooler"
    Write-Host "Final status: $($service.Status)" -ForegroundColor Cyan
    Write-Host "Final startup type: $($service.StartType)" -ForegroundColor Cyan
} catch {
    Write-Error "Enable-PrintSpoolerService: Failed to enable Print Spooler service - $($_.Exception.Message)"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    exit 1
}
