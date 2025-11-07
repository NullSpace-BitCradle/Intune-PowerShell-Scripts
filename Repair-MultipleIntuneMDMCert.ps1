#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Remediation script for multiple Intune MDM Device CA certificates.

.DESCRIPTION
    This script removes duplicate Intune MDM Device CA certificates from the Local Machine
    certificate store, keeping only the most recent one. This is the remediation script
    that pairs with Detect-MultipleIntuneMDMCert.ps1. It should be run when multiple
    certificates are detected.

.PARAMETER WhatIf
    Shows what would happen if the script runs. The script is not run.
    Displays which certificates would be removed without actually removing them.

.EXAMPLE
    .\Repair-MultipleIntuneMDMCert.ps1

.EXAMPLE
    .\Repair-MultipleIntuneMDMCert.ps1 -WhatIf

.NOTES
    Requires: Administrator privileges
    Exit Codes:
    - 0: Remediation successful (duplicates removed or no duplicates found)
    - 1: Remediation failed or error occurred
#>

[CmdletBinding(SupportsShouldProcess=$true)]
# Define script parameters
param(
    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# Main script execution wrapped in try-catch for error handling
try
{
    # Check if running as administrator
    # Certificate store modifications require administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Error "This script requires administrator privileges. Please run as administrator."
        exit 1
    }

    # Validate certificate store is accessible
    $certStorePath = "Cert:\LocalMachine\My"
    if (-not (Test-Path -Path $certStorePath)) {
        Write-Error "Certificate store not accessible: $certStorePath"
        exit 1
    }
    
    # Query the Local Machine Personal certificate store for certificates issued by Microsoft Intune MDM Device CA
    # The certificate store path Cert:\LocalMachine\My contains certificates in the Personal store
    # Filter results to only include certificates where the Issuer matches "CN=Microsoft Intune MDM Device CA"
    $certificates = Get-ChildItem -Path $certStorePath -ErrorAction Stop | Where-Object {$_.Issuer -eq "CN=Microsoft Intune MDM Device CA"}
    
    # Determine certificate count - handle both single object and array cases
    # When a single certificate is found, PowerShell returns an object, not an array
    $certificateCount = if ($certificates -is [Array]) { 
        $certificates.Count 
    } 
    else { 
        if ($certificates) { 1 } else { 0 } 
    }
    
    Write-Host "Found $certificateCount Intune MDM Device CA certificate(s)." -ForegroundColor Yellow
    
    # Check if there are multiple certificates (the condition that requires remediation)
    if($certificateCount -gt 1)
    {
        Write-Host "Multiple certificates detected. Removing duplicates..." -ForegroundColor Yellow
        
        # Sort certificates by NotBefore date (most recent first)
        # This ensures we keep the newest certificate and remove older duplicates
        $sortedCertificates = $certificates | Sort-Object -Property NotBefore -Descending
        
        # Keep the first (most recent) certificate
        $certificateToKeep = $sortedCertificates[0]
        Write-Host "Keeping most recent certificate (Thumbprint: $($certificateToKeep.Thumbprint), NotBefore: $($certificateToKeep.NotBefore))" -ForegroundColor Green
        
        # Remove all other (older) certificates
        # Start from index 1 to skip the certificate we're keeping
        for ($i = 1; $i -lt $sortedCertificates.Count; $i++)
        {
            $certToRemove = $sortedCertificates[$i]
            
            if ($WhatIf) {
                Write-Host "What if: Would remove duplicate certificate (Thumbprint: $($certToRemove.Thumbprint), NotBefore: $($certToRemove.NotBefore))" -ForegroundColor Yellow
            }
            else {
                Write-Host "Removing duplicate certificate (Thumbprint: $($certToRemove.Thumbprint), NotBefore: $($certToRemove.NotBefore))" -ForegroundColor Yellow
                
                # Remove the certificate from the certificate store
                # This permanently deletes the certificate
                Remove-Item -Path $certToRemove.PSPath -Force -ErrorAction Stop
                
                Write-Host "Successfully removed certificate with Thumbprint: $($certToRemove.Thumbprint)" -ForegroundColor Green
            }
        }
        
        if ($WhatIf) {
            Write-Host "`nWhat if: Would keep certificate (Thumbprint: $($certificateToKeep.Thumbprint))" -ForegroundColor Green
            Write-Host "What if: Would remove $($sortedCertificates.Count - 1) duplicate certificate(s)" -ForegroundColor Yellow
            exit 0
        }
        
        # Verify the remediation was successful
        # Query the certificate store again to confirm only one certificate remains
        $remainingCertificates = Get-ChildItem -Path $certStorePath -ErrorAction Stop | Where-Object {$_.Issuer -eq "CN=Microsoft Intune MDM Device CA"}
        
        # Determine certificate count - handle both single object and array cases
        $remainingCount = if ($remainingCertificates -is [Array]) { 
            $remainingCertificates.Count 
        } 
        else { 
            if ($remainingCertificates) { 1 } else { 0 } 
        }
        
        if ($remainingCount -eq 1)
        {
            # Get the certificate object (handle both array and single object cases)
            $remainingCert = if ($remainingCertificates -is [Array]) { 
                $remainingCertificates[0] 
            } else { 
                $remainingCertificates 
            }
            Write-Host "Remediation successful. One certificate remains (Thumbprint: $($remainingCert.Thumbprint))." -ForegroundColor Green
            exit 0
        }
        else
        {
            Write-Error "Remediation failed. Expected 1 certificate but found $remainingCount."
            exit 1
        }
    }
    else
    {
        # Zero or one certificate found - no remediation needed
        # This is the expected state, so we exit successfully
        Write-Host "No remediation needed. Expected number of certificates found ($certificateCount)." -ForegroundColor Green
        exit 0
    }
}
# Error handling block - catches any exceptions during certificate operations
catch 
{
    # Capture the error message for logging
    $errMsg = $_.Exception.Message
    # Display the error message to the console with script name for context
    Write-Error "Repair-MultipleIntuneMDMCert: Failed to repair certificates - $errMsg"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    # Exit with code 1 to indicate an error occurred
    exit 1
}

