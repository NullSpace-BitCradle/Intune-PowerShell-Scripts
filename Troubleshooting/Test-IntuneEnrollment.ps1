<#
.SYNOPSIS
    Tests if device is properly enrolled in Intune.

.DESCRIPTION
    This script checks the local device to verify if it is properly enrolled in Microsoft Intune.
    It checks enrollment status, MDM authority, enrollment date, and management state.
    The script can be run locally on the device to verify enrollment status.

.EXAMPLE
    .\Test-IntuneEnrollment.ps1

.NOTES
    Exit Codes:
    - 0: Device is properly enrolled in Intune
    - 1: Device is not enrolled or enrollment issues detected
#>

[CmdletBinding()]
param()

# Main script execution wrapped in try-catch for error handling
try
{
    Write-Host "Testing Intune enrollment status..." -ForegroundColor Yellow
    Write-Host "====================================" -ForegroundColor Yellow
    
    # Initialize enrollment status variables
    # These will track various enrollment indicators
    $isEnrolled = $false
    $enrollmentIssues = @()
    $enrollmentInfo = @{}
    
    # Check 1: Verify MDM enrollment in registry
    # The registry path contains MDM enrollment information
    Write-Host "`n[1/5] Checking MDM enrollment registry..." -ForegroundColor Cyan
    $mdmEnrollmentPath = "HKLM:\SOFTWARE\Microsoft\Enrollments"
    
    if (Test-Path -Path $mdmEnrollmentPath) {
        # Get all enrollment subkeys
        # Each enrollment has a GUID subkey
        $enrollmentKeys = Get-ChildItem -Path $mdmEnrollmentPath -ErrorAction SilentlyContinue
        
        if ($enrollmentKeys) {
            Write-Host "  Found $($enrollmentKeys.Count) enrollment key(s)." -ForegroundColor Green
            
            # Check each enrollment for Intune/MDM authority
            foreach ($key in $enrollmentKeys) {
                $providerId = Get-ItemPropertyValue -Path $key.PSPath -Name "ProviderID" -ErrorAction SilentlyContinue
                $enrollmentType = Get-ItemPropertyValue -Path $key.PSPath -Name "EnrollmentType" -ErrorAction SilentlyContinue
                
                # Check if this is an Intune/MDM enrollment
                # Intune uses specific provider IDs and enrollment types
                if ($providerId -and $enrollmentType) {
                    Write-Host "  Provider ID: $providerId" -ForegroundColor White
                    Write-Host "  Enrollment Type: $enrollmentType" -ForegroundColor White
                    
                    # Check for Intune-specific indicators
                    if ($providerId -like "*Intune*" -or $providerId -like "*Microsoft*") {
                        $isEnrolled = $true
                        $enrollmentInfo["ProviderID"] = $providerId
                        $enrollmentInfo["EnrollmentType"] = $enrollmentType
                    }
                }
            }
        }
        else {
            $enrollmentIssues += "No enrollment keys found in registry"
            Write-Host "  No enrollment keys found." -ForegroundColor Red
        }
    }
    else {
        $enrollmentIssues += "MDM enrollment registry path not found"
        Write-Host "  MDM enrollment registry path not found." -ForegroundColor Red
    }
    
    # Check 2: Verify MDM authority
    # Check the MDM authority setting which indicates the management service
    Write-Host "`n[2/5] Checking MDM authority..." -ForegroundColor Cyan
    
    try {
        # Get MDM authority from registry
        # This indicates which MDM service is managing the device
        $mdmAuthority = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Enrollments" -Name "MDMAuthority" -ErrorAction SilentlyContinue
        
        if ($mdmAuthority) {
            Write-Host "  MDM Authority: $mdmAuthority" -ForegroundColor Green
            $enrollmentInfo["MDMAuthority"] = $mdmAuthority
            
            # Check if MDM authority is Intune
            if ($mdmAuthority -eq "Intune" -or $mdmAuthority -like "*Intune*") {
                $isEnrolled = $true
            }
            else {
                $enrollmentIssues += "MDM Authority is not Intune: $mdmAuthority"
                Write-Host "  Warning: MDM Authority is not Intune." -ForegroundColor Yellow
            }
        }
        else {
            # Try alternative method to find MDM authority
            $enrollmentKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Enrollments" -ErrorAction SilentlyContinue
            foreach ($key in $enrollmentKeys) {
                $providerPath = Join-Path $key.PSPath "DMClient\Provider"
                if (Test-Path -Path $providerPath) {
                    $providerId = Get-ItemPropertyValue -Path $providerPath -Name "ProviderID" -ErrorAction SilentlyContinue
                    if ($providerId) {
                        Write-Host "  Provider ID: $providerId" -ForegroundColor Green
                        $enrollmentInfo["ProviderID"] = $providerId
                        if ($providerId -like "*Intune*" -or $providerId -like "*Microsoft*") {
                            $isEnrolled = $true
                        }
                    }
                }
            }
            
            if (-not $mdmAuthority) {
                $enrollmentIssues += "MDM Authority not found"
                Write-Host "  MDM Authority not found." -ForegroundColor Red
            }
        }
    }
    catch {
        $enrollmentIssues += "Error checking MDM authority: $($_.Exception.Message)"
        Write-Host "  Error checking MDM authority." -ForegroundColor Red
    }
    
    # Check 3: Verify enrollment date
    # Check when the device was enrolled
    Write-Host "`n[3/5] Checking enrollment date..." -ForegroundColor Cyan
    try {
        $enrollmentDate = Get-ItemPropertyValue -Path "HKLM:\SOFTWARE\Microsoft\Enrollments\*\FirstSyncTime" -ErrorAction SilentlyContinue
        
        if ($enrollmentDate) {
            Write-Host "  Enrollment Date: $enrollmentDate" -ForegroundColor Green
            $enrollmentInfo["EnrollmentDate"] = $enrollmentDate
        }
        else {
            # Try alternative method
            $enrollmentKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Enrollments" -ErrorAction SilentlyContinue
            foreach ($key in $enrollmentKeys) {
                $firstSync = Get-ItemPropertyValue -Path $key.PSPath -Name "FirstSyncTime" -ErrorAction SilentlyContinue
                if ($firstSync) {
                    Write-Host "  Enrollment Date: $firstSync" -ForegroundColor Green
                    $enrollmentInfo["EnrollmentDate"] = $firstSync
                    break
                }
            }
            
            if (-not $enrollmentInfo["EnrollmentDate"]) {
                $enrollmentIssues += "Enrollment date not found"
                Write-Host "  Enrollment date not found." -ForegroundColor Yellow
            }
        }
    }
    catch {
        $enrollmentIssues += "Error checking enrollment date: $($_.Exception.Message)"
        Write-Host "  Error checking enrollment date." -ForegroundColor Red
    }
    
    # Check 4: Verify device is managed
    # Check if the device is actually being managed by MDM
    Write-Host "`n[4/5] Checking device management status..." -ForegroundColor Cyan
    try {
        # Check for MDM enrollment status
        $isManaged = $false
        
        # Check registry for management status
        $enrollmentKeys = Get-ChildItem -Path "HKLM:\SOFTWARE\Microsoft\Enrollments" -ErrorAction SilentlyContinue
        foreach ($key in $enrollmentKeys) {
            $isEnrolledValue = Get-ItemPropertyValue -Path $key.PSPath -Name "IsEnrolled" -ErrorAction SilentlyContinue
            if ($isEnrolledValue -eq 1 -or $isEnrolledValue -eq $true) {
                $isManaged = $true
                Write-Host "  Device is enrolled: $isEnrolledValue" -ForegroundColor Green
                $enrollmentInfo["IsEnrolled"] = $isEnrolledValue
                break
            }
        }
        
        if (-not $isManaged) {
            $enrollmentIssues += "Device management status not confirmed"
            Write-Host "  Device management status not confirmed." -ForegroundColor Yellow
        }
    }
    catch {
        $enrollmentIssues += "Error checking management status: $($_.Exception.Message)"
        Write-Host "  Error checking management status." -ForegroundColor Red
    }
    
    # Check 5: Verify Intune Management Extension (if applicable)
    # Check if Intune Management Extension is installed and running
    Write-Host "`n[5/5] Checking Intune Management Extension..." -ForegroundColor Cyan
    try {
        $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue
        
        if ($imeService) {
            Write-Host "  Intune Management Extension service found." -ForegroundColor Green
            Write-Host "  Service Status: $($imeService.Status)" -ForegroundColor $(if ($imeService.Status -eq 'Running') { 'Green' } else { 'Yellow' })
            $enrollmentInfo["IMEServiceStatus"] = $imeService.Status
            
            if ($imeService.Status -ne 'Running') {
                $enrollmentIssues += "Intune Management Extension service is not running"
            }
        }
        else {
            Write-Host "  Intune Management Extension service not found (may not be required)." -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "  Intune Management Extension service check skipped (may not be installed)." -ForegroundColor Yellow
    }
    
    # Summary
    Write-Host "`n====================================" -ForegroundColor Yellow
    Write-Host "Enrollment Test Summary" -ForegroundColor Cyan
    Write-Host "====================================" -ForegroundColor Yellow
    
    if ($isEnrolled) {
        Write-Host "`nResult: Device appears to be enrolled in Intune." -ForegroundColor Green
        
        # Display enrollment information
        Write-Host "`nEnrollment Information:" -ForegroundColor Cyan
        foreach ($key in $enrollmentInfo.Keys) {
            Write-Host "  $key : $($enrollmentInfo[$key])" -ForegroundColor White
        }
        
        # Display any issues found
        if ($enrollmentIssues.Count -gt 0) {
            Write-Host "`nIssues Found:" -ForegroundColor Yellow
            foreach ($issue in $enrollmentIssues) {
                Write-Host "  - $issue" -ForegroundColor Yellow
            }
        }
        
        exit 0
    }
    else {
        Write-Host "`nResult: Device does not appear to be enrolled in Intune." -ForegroundColor Red
        
        # Display issues found
        if ($enrollmentIssues.Count -gt 0) {
            Write-Host "`nIssues Found:" -ForegroundColor Red
            foreach ($issue in $enrollmentIssues) {
                Write-Host "  - $issue" -ForegroundColor Red
            }
        }
        
        exit 1
    }
}
# Error handling block - catches exceptions during enrollment checks
catch
{
    # Capture and display the error message with script name for context
    $errMsg = $_.Exception.Message
    Write-Error "Test-IntuneEnrollment: Failed to test Intune enrollment - $errMsg"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    exit 1
}

