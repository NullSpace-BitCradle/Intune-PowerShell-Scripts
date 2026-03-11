#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Disables the Print Spooler service to mitigate PrintNightmare vulnerabilities.

.DESCRIPTION
    This script stops the Print Spooler service if it's running and sets its startup type to Disabled.
    Requires administrator privileges.

.EXAMPLE
    .\Disable-PrintSpoolerService.ps1

.EXAMPLE
    .\Disable-PrintSpoolerService.ps1 -WhatIf

.NOTES
    Requires: Administrator privileges
#>

[CmdletBinding(SupportsShouldProcess=$true)]
# Define script parameters
param()

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

    if ($service.Status -eq "Running") {
        if ($PSCmdlet.ShouldProcess("Print Spooler service", "Stop service")) {
            Write-Host "Stopping Print Spooler service..." -ForegroundColor Yellow
            Stop-Service -Name "Spooler" -Force -ErrorAction Stop
            Write-Host "Print Spooler service stopped successfully." -ForegroundColor Green
        }
    }
    else {
        Write-Host "Print Spooler service is not running." -ForegroundColor Yellow
    }

    if ($PSCmdlet.ShouldProcess("Print Spooler service", "Set startup type to Disabled")) {
        Write-Host "Setting Print Spooler startup type to Disabled..." -ForegroundColor Yellow
        Set-Service -Name "Spooler" -StartupType Disabled -ErrorAction Stop
        Write-Host "Print Spooler startup type set to Disabled successfully." -ForegroundColor Green

        # Verify the change
        $service = Get-Service -Name "Spooler"
        Write-Host "Final status: $($service.Status)" -ForegroundColor Cyan
        Write-Host "Final startup type: $($service.StartType)" -ForegroundColor Cyan
    }
} catch {
    Write-Error "Disable-PrintSpoolerService: Failed to disable Print Spooler service - $($_.Exception.Message)"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    exit 1
}
