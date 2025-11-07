<#
.SYNOPSIS
    Detection script for multiple Intune MDM Device CA certificates.

.DESCRIPTION
    This script checks the Local Machine certificate store for multiple Intune MDM Device CA certificates.
    If more than one certificate is found, it indicates a remediation is needed.
    Used as a detection script in Intune compliance policies or remediation scripts.

.EXAMPLE
    .\Detect-MultipleIntuneMDMCert.ps1

.NOTES
    Exit Codes:
    - 0: No remediation needed (0 or 1 certificate found)
    - 1: Remediation needed (more than 1 certificate found) or error occurred
#>

[CmdletBinding()]
param()

# Main script execution wrapped in try-catch for error handling
try
{
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
    $certCount = if ($certificates -is [Array]) { 
        $certificates.Count 
    } 
    else { 
        if ($certificates) { 1 } else { 0 } 
    }
    
    # Check if more than one certificate was found
    # Multiple certificates indicate a problem that needs remediation
    if($certCount -gt 1)
    {
        # Multiple certificates detected - remediation is required
        # Exit code 1 signals to Intune that remediation is needed
        Write-Output "Remediation needed"
        exit 1
    }
    else
    {
        # Zero or one certificate found - this is the expected state
        # Exit code 0 signals to Intune that no remediation is needed
        Write-Output "Remediation not needed"
        exit 0
    }
}
# Error handling block - catches any exceptions during certificate query or processing
catch 
{
    # Capture the error message for logging
    $errMsg = $_.Exception.Message
    # Display the error message to the console with script name for context
    Write-Error "Detect-MultipleIntuneMDMCert: Failed to query certificate store - $errMsg"
    if ($PSBoundParameters.ContainsKey('Verbose')) {
        Write-Verbose "Stack Trace: $($_.ScriptStackTrace)"
        if ($_.Exception.InnerException) {
            Write-Verbose "Inner Exception: $($_.Exception.InnerException.Message)"
        }
    }
    # Exit with code 1 to indicate an error occurred (treats errors as requiring remediation)
    exit 1
}
