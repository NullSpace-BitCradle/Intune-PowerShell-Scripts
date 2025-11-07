# Intune PowerShell Scripts

A comprehensive collection of PowerShell scripts for managing, troubleshooting, and reporting on Microsoft Intune deployments. These scripts provide detection, remediation, reporting, and management capabilities for Intune-managed devices and policies.

## 📋 Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Script Categories](#script-categories)
  - [Detection & Remediation](#detection--remediation)
  - [Device Management](#device-management)
  - [Reporting & Analysis](#reporting--analysis)
  - [Service Management](#service-management)
  - [Connectivity & Testing](#connectivity--testing)
- [Common Module](#common-module)
- [Usage Examples](#usage-examples)
- [Exit Codes](#exit-codes)
- [Error Handling](#error-handling)
- [Security Considerations](#security-considerations)
- [Contributing](#contributing)
- [Script Index](#script-index)

## Overview

This repository contains production-ready PowerShell scripts designed for Microsoft Intune administrators. The scripts are organized into logical categories and follow PowerShell best practices, including:

- Comprehensive error handling
- Standardized exit codes
- Detailed logging and verbose output
- Input validation
- WhatIf support for modification scripts
- Progress indicators for long-running operations

## Prerequisites

### General Requirements

- **Windows 10/11** or **Windows Server 2016+**
- **PowerShell 5.1** or later (PowerShell 7+ recommended)
- **Administrator privileges** (required for scripts that modify system settings)

### For Graph API Scripts

Scripts that interact with Microsoft Graph API require:

- **Microsoft.Graph PowerShell Module** (automatically installed if missing)
- **Entra ID App Registration** with the following Graph API permissions:
  - `DeviceManagementManagedDevices.Read.All`
  - `DeviceManagementConfiguration.Read.All`
  - `DeviceManagementApps.Read.All`
  - `Group.Read.All` (for group name resolution)

### App Registration Setup

1. Navigate to [Azure Portal](https://portal.azure.com) → **Entra ID** → **App registrations**
2. Create a new app registration or use an existing one
3. Create a **Client Secret** and note the value (it won't be shown again)
4. Grant the required API permissions listed above
5. Note the **Application (client) ID** and **Directory (tenant) ID**

## Installation

1. Clone or download this repository:

   ```powershell
   git clone https://github.com/yourusername/Intune-PowerShell-Scripts.git
   cd Intune-PowerShell-Scripts
   ```

2. Import the common module (optional, but recommended):

   ```powershell
   Import-Module .\IntuneCommon.psm1
   ```

3. Configure credentials using environment variables (recommended):

   ```powershell
   $env:INTUNE_APP_ID = "your-app-id"
   $env:INTUNE_TENANT_ID = "your-tenant-id"
   $env:INTUNE_CLIENT_SECRET = "your-client-secret"
   ```

## Quick Start

### Example: Get Device Details

```powershell
.\Get-IntuneDeviceDetails.ps1 `
    -AppId $env:INTUNE_APP_ID `
    -TenantId $env:INTUNE_TENANT_ID `
    -ClientSecret $env:INTUNE_CLIENT_SECRET `
    -ExportPath "C:\Reports\Devices.csv"
```

### Example: Test Device Enrollment

```powershell
.\Test-IntuneEnrollment.ps1
```

### Example: Check Device Health

```powershell
.\Get-IntuneDeviceHealth.ps1 `
    -AppId $env:INTUNE_APP_ID `
    -TenantId $env:INTUNE_TENANT_ID `
    -ClientSecret $env:INTUNE_CLIENT_SECRET `
    -ExportPath "C:\Reports\HealthReport.csv"
```

## Script Categories

### Detection & Remediation

Scripts designed for use in Intune compliance policies and remediation scripts.

#### `Detect-MultipleIntuneMDMCert.ps1`

Detects multiple Intune MDM Device CA certificates in the Local Machine certificate store. Used as a detection script in Intune compliance policies.

**Usage:**

```powershell
.\Detect-MultipleIntuneMDMCert.ps1
```

**Exit Codes:**

- `0`: No remediation needed (0 or 1 certificate found)
- `1`: Remediation needed (more than 1 certificate found) or error occurred

**Parameters:**

- None

---

#### `Repair-MultipleIntuneMDMCert.ps1`

Remediation script that removes duplicate Intune MDM Device CA certificates, keeping only the most recent one.

**Usage:**

```powershell
.\Repair-MultipleIntuneMDMCert.ps1 [-WhatIf]
```

**Exit Codes:**

- `0`: Success (remediation completed or not needed)
- `1`: Error occurred

**Parameters:**

- `-WhatIf`: Shows what would be done without making changes

**Requirements:** Administrator privileges

---

#### `Detect-OfficeUpdateChannel.ps1`

Verifies if Microsoft Office is using the Semi-Annual update channel and is on the latest version. Used as a detection script in Intune compliance policies.

**Usage:**

```powershell
.\Detect-OfficeUpdateChannel.ps1
```

**Exit Codes:**

- `0`: Office is on Semi-Annual channel and latest version
- `1`: Office is not on Semi-Annual channel or not on latest version

**Parameters:**

- None

**Requirements:** Office Click-to-Run installation

---

#### `Set-OfficeUpdateChannel.ps1`

Remediation script that configures Microsoft Office to use the Semi-Annual update channel and triggers an update to the latest version.

**Usage:**

```powershell
.\Set-OfficeUpdateChannel.ps1 [-WhatIf]
```

**Exit Codes:**

- `0`: Success (Office configured and updated)
- `1`: Error occurred

**Parameters:**

- `-WhatIf`: Shows what would be done without making changes

**Requirements:** Administrator privileges, Office Click-to-Run installation

---

### Device Management

Scripts for managing and monitoring Intune-enrolled devices.

#### `Get-IntuneDeviceDetails.ps1`

Retrieves comprehensive information about Intune-managed devices including enrollment details, hardware information, compliance state, and management agent.

**Usage:**

```powershell
# Get all devices
.\Get-IntuneDeviceDetails.ps1 -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Get specific device
.\Get-IntuneDeviceDetails.ps1 -DeviceId 'device-guid' -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Search by device name
.\Get-IntuneDeviceDetails.ps1 -DeviceName 'LAPTOP-*' -ExportPath 'C:\Reports\Devices.csv'

# Export to CSV
.\Get-IntuneDeviceDetails.ps1 -ExportPath 'C:\Reports\Devices.csv' -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'
```

**Parameters:**

- `-DeviceId`: Specific device ID to query
- `-DeviceName`: Device name filter (supports wildcards)
- `-ExportPath`: Path to export CSV file
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

---

#### `Get-IntuneDeviceCompliance.ps1`

Retrieves compliance status for Intune-managed devices, including compliance policy assignments and compliance state.

**Usage:**

```powershell
# Get compliance for all devices
.\Get-IntuneDeviceCompliance.ps1 -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Get compliance for specific device
.\Get-IntuneDeviceCompliance.ps1 -DeviceId 'device-guid' -ExportPath 'C:\Reports\Compliance.csv'

# Get compliance for user's devices
.\Get-IntuneDeviceCompliance.ps1 -UserPrincipalName 'user@domain.com' -ExportPath 'C:\Reports\UserCompliance.csv'
```

**Parameters:**

- `-DeviceId`: Specific device ID to query
- `-UserPrincipalName`: User principal name to filter devices
- `-ExportPath`: Path to export CSV file
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

---

#### `Get-IntuneDeviceHealth.ps1`

Performs comprehensive health checks on Intune-managed devices including enrollment status, sync status, compliance state, management agent, and storage space.

**Usage:**

```powershell
# Check health for all devices
.\Get-IntuneDeviceHealth.ps1 -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Check health for specific device
.\Get-IntuneDeviceHealth.ps1 -DeviceId 'device-guid' -ExportPath 'C:\Reports\HealthReport.csv'
```

**Parameters:**

- `-DeviceId`: Specific device ID to query
- `-ExportPath`: Path to export CSV file
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

**Health Check Factors:**

- Enrollment status and age
- Last sync time
- Compliance state
- Management agent type
- Operating system version
- Storage space (if available)

---

#### `Test-IntuneEnrollment.ps1`

Tests if a device is properly enrolled in Intune by checking enrollment status, MDM authority, enrollment date, and Intune Management Extension service.

**Usage:**

```powershell
.\Test-IntuneEnrollment.ps1 [-Verbose]
```

**Exit Codes:**

- `0`: Device is properly enrolled
- `1`: Device is not enrolled or enrollment issues detected

**Parameters:**

- `-Verbose`: Display detailed enrollment information

---

#### `Start-MDMSync.ps1`

Initiates an MDM sync session between the device and Intune using Windows Management APIs.

**Usage:**

```powershell
# Default timeout (60 seconds)
.\Start-MDMSync.ps1

# Custom timeout and check interval
.\Start-MDMSync.ps1 -TimeoutSeconds 120 -CheckIntervalSeconds 10
```

**Exit Codes:**

- `0`: Sync completed successfully
- `1`: Sync failed or ended with error state
- `2`: Sync timeout (didn't complete within timeout period)

**Parameters:**

- `-TimeoutSeconds`: Maximum time in seconds to wait for sync (default: 60)
- `-CheckIntervalSeconds`: Interval in seconds between status checks (default: 5)

**Requirements:** Windows 10/11 with MDM enrollment

---

### Reporting & Analysis

Scripts for generating reports and analyzing Intune deployments.

#### `Get-IntuneAllAppsAssignmentDetails.ps1`

Retrieves all applications published in Microsoft Intune and exports their assignment details to a CSV file.

**Usage:**

```powershell
.\Get-IntuneAllAppsAssignmentDetails.ps1 `
    -AppId 'app-id' `
    -TenantId 'tenant-id' `
    -ClientSecret 'secret' `
    -ExportCSVpath 'C:\Reports\AppAssignments.csv'
```

**Parameters:**

- `-ExportCSVpath`: Path to export CSV file (default: `C:\Temp\Get-IntuneAllAppsAssignmentDetails.csv`)
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

---

#### `Get-IntunePolicyAssignments.ps1`

Retrieves all Intune policy assignments including configuration profiles and compliance policies with their target groups.

**Usage:**

```powershell
# Get all policy assignments
.\Get-IntunePolicyAssignments.ps1 -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Get only configuration profiles
.\Get-IntunePolicyAssignments.ps1 -PolicyType "Configuration" -ExportPath 'C:\Reports\ConfigProfiles.csv'

# Get only compliance policies
.\Get-IntunePolicyAssignments.ps1 -PolicyType "Compliance" -ExportPath 'C:\Reports\CompliancePolicies.csv'
```

**Parameters:**

- `-PolicyType`: Type of policies to retrieve ("All", "Configuration", or "Compliance") (default: "All")
- `-ExportPath`: Path to export CSV file
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

---

#### `Get-IntuneWin32AppDetails.ps1`

Retrieves detailed information about Win32 app deployments including installation status, assignment details, and device installation status.

**Usage:**

```powershell
# Get all Win32 apps
.\Get-IntuneWin32AppDetails.ps1 -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Get specific app
.\Get-IntuneWin32AppDetails.ps1 -IntuneAppId 'app-guid' -ExportPath 'C:\Reports\AppDetails.csv'

# Get app with device installation status
.\Get-IntuneWin32AppDetails.ps1 -IntuneAppId 'app-guid' -DeviceId 'device-guid' -ExportPath 'C:\Reports\AppStatus.csv'
```

**Parameters:**

- `-IntuneAppId`: Specific Win32 app ID to query
- `-DeviceId`: Device ID to check installation status
- `-ExportPath`: Path to export CSV file
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

---

#### `Export-IntuneDeviceReport.ps1`

Exports comprehensive device information to CSV or JSON format including device details, compliance status, policy assignments, and configuration profile states.

**Usage:**

```powershell
# Export all devices to CSV
.\Export-IntuneDeviceReport.ps1 -AppId 'app-id' -TenantId 'tenant-id' -ClientSecret 'secret'

# Export specific device to JSON
.\Export-IntuneDeviceReport.ps1 `
    -DeviceId 'device-guid' `
    -Format 'JSON' `
    -OutputPath 'C:\Reports\DeviceReport.json' `
    -AppId 'app-id' `
    -TenantId 'tenant-id' `
    -ClientSecret 'secret'
```

**Parameters:**

- `-DeviceId`: Specific device ID to export
- `-Format`: Export format ("CSV" or "JSON") (default: "CSV")
- `-OutputPath`: Path to export file (default: timestamped file in `$env:TEMP\IntuneReports`)
- `-AppId`: Entra ID Application ID (or use environment variable)
- `-TenantId`: Entra ID Tenant ID (or use environment variable)
- `-ClientSecret`: Entra ID Client Secret (or use environment variable)

---

#### `Get-IntuneDeviceLogs.ps1`

Collects Intune-related logs from the device including Intune Management Extension logs, MDM enrollment registry information, and Windows Event Logs.

**Usage:**

```powershell
# Collect logs to default location
.\Get-IntuneDeviceLogs.ps1

# Collect logs with Event Logs to custom location
.\Get-IntuneDeviceLogs.ps1 -OutputPath 'C:\Logs\Intune' -IncludeEventLogs
```

**Parameters:**

- `-OutputPath`: Path to save collected logs (default: timestamped folder in `$env:TEMP\IntuneLogs`)
- `-IncludeEventLogs`: Include Windows Event Logs in collection

**Requirements:** Administrator privileges

**Collected Logs:**

- Intune Management Extension logs
- MDM enrollment registry information
- Windows Event Logs (if `-IncludeEventLogs` is specified)

---

### Service Management

Scripts for managing Windows services related to Intune and security.

#### `Disable-PrintSpoolerService.ps1`

Disables the Print Spooler service to mitigate PrintNightmare vulnerabilities.

**Usage:**

```powershell
.\Disable-PrintSpoolerService.ps1 [-WhatIf]
```

**Exit Codes:**

- `0`: Success (service disabled)
- `1`: Error occurred

**Parameters:**

- `-WhatIf`: Shows what would be done without making changes

**Requirements:** Administrator privileges

**Note:** This will disable printing functionality on the device.

---

#### `Enable-PrintSpoolerService.ps1`

Enables the Print Spooler service and sets it to start automatically.

**Usage:**

```powershell
.\Enable-PrintSpoolerService.ps1 [-WhatIf]
```

**Exit Codes:**

- `0`: Success (service enabled)
- `1`: Error occurred

**Parameters:**

- `-WhatIf`: Shows what would be done without making changes

**Requirements:** Administrator privileges

---

#### `Disable-SmartCardLogonEnforcement.ps1`

Disables smart card logon enforcement by modifying registry keys and disabling the Smart Card Policy Service.

**Usage:**

```powershell
.\Disable-SmartCardLogonEnforcement.ps1 [-WhatIf]
```

**Exit Codes:**

- `0`: Success (smart card enforcement disabled)
- `1`: Error occurred

**Parameters:**

- `-WhatIf`: Shows what would be done without making changes

**Requirements:** Administrator privileges

**Note:** This allows password-based logon and prevents auto-lock on smart card removal.

---

#### `Enable-SmartCardLogonEnforcement.ps1`

Enables smart card logon enforcement by modifying registry keys and enabling the Smart Card Policy Service.

**Usage:**

```powershell
.\Enable-SmartCardLogonEnforcement.ps1 [-WhatIf]
```

**Exit Codes:**

- `0`: Success (smart card enforcement enabled)
- `1`: Error occurred

**Parameters:**

- `-WhatIf`: Shows what would be done without making changes

**Requirements:** Administrator privileges

**Note:** This enforces smart card authentication and enables auto-lock on smart card removal.

---

### Connectivity & Testing

Scripts for testing connectivity and enrollment status.

#### `Test-IntuneConnectivity.ps1`

Tests connectivity to Microsoft Intune and Microsoft 365 endpoints, displays network configuration information, and exports results to CSV.

**Usage:**

```powershell
# Test connectivity
.\Test-IntuneConnectivity.ps1

# Test with export
.\Test-IntuneConnectivity.ps1 -ExportPath 'C:\Reports\Connectivity.csv'
```

**Parameters:**

- `-ExportPath`: Path to export CSV file with test results

**Tested Endpoints:**

- Microsoft 365 Common endpoints (authentication, identity)
- Intune/MEM endpoints (device management, enrollment)

---

## Common Module

### `IntuneCommon.psm1`

A PowerShell module containing shared functions used across multiple scripts:

- **`Invoke-GraphApiWithRetry`**: Invokes Graph API requests with automatic retry logic and exponential backoff
- **`Get-GraphAccessToken`**: Retrieves OAuth2 access token for Microsoft Graph API
- **`Test-GuidFormat`**: Validates GUID format
- **`Test-EmailFormat`**: Validates email address format
- **`Remove-ODataInjectionChars`**: Sanitizes input for OData queries

**Usage:**

```powershell
Import-Module .\IntuneCommon.psm1
```

The module is automatically used by scripts that require Graph API access. You can also use it directly in your own scripts.

---

## Usage Examples

### Example 1: Generate Comprehensive Device Report

```powershell
# Export all devices to CSV with full details
.\Export-IntuneDeviceReport.ps1 `
    -AppId $env:INTUNE_APP_ID `
    -TenantId $env:INTUNE_TENANT_ID `
    -ClientSecret $env:INTUNE_CLIENT_SECRET `
    -OutputPath "C:\Reports\AllDevices_$(Get-Date -Format 'yyyyMMdd').csv"
```

### Example 2: Check Compliance for Specific User

```powershell
# Get compliance status for all devices owned by a user
.\Get-IntuneDeviceCompliance.ps1 `
    -UserPrincipalName "john.doe@contoso.com" `
    -ExportPath "C:\Reports\JohnDoe_Compliance.csv" `
    -AppId $env:INTUNE_APP_ID `
    -TenantId $env:INTUNE_TENANT_ID `
    -ClientSecret $env:INTUNE_CLIENT_SECRET
```

### Example 3: Monitor Device Health

```powershell
# Check health for all devices and export results
.\Get-IntuneDeviceHealth.ps1 `
    -ExportPath "C:\Reports\DeviceHealth_$(Get-Date -Format 'yyyyMMdd').csv" `
    -AppId $env:INTUNE_APP_ID `
    -TenantId $env:INTUNE_TENANT_ID `
    -ClientSecret $env:INTUNE_CLIENT_SECRET
```

### Example 4: Remediate Office Update Channel

```powershell
# First, detect the issue
.\Detect-OfficeUpdateChannel.ps1

# If detection returns exit code 1, run remediation
if ($LASTEXITCODE -eq 1) {
    .\Set-OfficeUpdateChannel.ps1
}
```

### Example 5: Collect Logs for Troubleshooting

```powershell
# Collect all Intune-related logs
.\Get-IntuneDeviceLogs.ps1 `
    -OutputPath "C:\Logs\Intune_$(Get-Date -Format 'yyyyMMdd_HHmmss')" `
    -IncludeEventLogs
```

---

## Exit Codes

All scripts follow standard exit code conventions:

| Exit Code | Meaning |
|-----------|---------|
| `0` | Success / No action needed |
| `1` | Failure / Action needed / Error occurred |
| `2` | Timeout or specific error condition (where applicable) |

### Exit Code Usage

- **Detection scripts**: `0` = compliant, `1` = non-compliant or error
- **Remediation scripts**: `0` = success, `1` = error
- **Reporting scripts**: `0` = success, `1` = error
- **Management scripts**: `0` = success, `1` = error, `2` = timeout

---

## Error Handling

All scripts include comprehensive error handling:

- **Try-catch blocks** for exception handling
- **Standardized error messages** with script name prefixes
- **Verbose output** for debugging (use `-Verbose` parameter)
- **Stack traces** in verbose mode
- **Null response handling** for API calls
- **Input validation** for parameters
- **Progress indicators** for long-running operations

### Error Message Format

Error messages follow a consistent format:

```text
ScriptName: Failed to perform action - Error details
```

### Verbose Output

Use the `-Verbose` parameter to get detailed information:

```powershell
.\Get-IntuneDeviceDetails.ps1 -DeviceId 'guid' -Verbose
```

---

## Security Considerations

### Credentials

- **Never hardcode credentials** in scripts
- Use **environment variables** or **parameters** for sensitive information
- Consider using **Azure Key Vault** for production environments
- **Rotate client secrets** regularly

### Permissions

- Scripts that modify system settings require **administrator privileges**
- Graph API scripts require **appropriate API permissions** in Entra ID
- Use **principle of least privilege** when assigning permissions

### Registry Modifications

- Some scripts modify registry keys
- **Always review changes** before deployment
- **Test in non-production** environments first
- **Backup registry** before making changes

### Service Management Security

- Service management scripts can **impact system functionality**
- Use `-WhatIf` parameter to preview changes
- **Test thoroughly** before deploying to production

### Network Security

- Scripts communicate with Microsoft Graph API over HTTPS
- Ensure **firewall rules** allow access to Graph API endpoints
- Use **Test-IntuneConnectivity.ps1** to verify network connectivity

---

## Contributing

When contributing to this repository:

1. **Follow PowerShell best practices**:
   - Use approved verbs (Get, Set, Test, Start, etc.)
   - Follow verb-noun naming convention
   - Include comprehensive help documentation

2. **Code quality**:
   - Include comprehensive header documentation
   - Add inline comments explaining functionality
   - Implement proper error handling
   - Use appropriate exit codes
   - Add input validation

3. **Testing**:
   - Test scripts thoroughly before committing
   - Test with `-WhatIf` parameter for modification scripts
   - Test error scenarios
   - Verify exit codes

4. **Documentation**:
   - Update README.md with new scripts
   - Include usage examples
   - Document all parameters
   - Document exit codes

5. **Common functions**:
   - Use `IntuneCommon.psm1` for shared functionality
   - Add new common functions when appropriate
   - Keep functions focused and reusable

---

## Script Index

| Script Name | Category | Requires Admin | Requires Graph API | Description |
|------------|----------|----------------|---------------------|-------------|
| `Detect-MultipleIntuneMDMCert.ps1` | Detection | No | No | Detects multiple Intune MDM certificates |
| `Repair-MultipleIntuneMDMCert.ps1` | Remediation | Yes | No | Removes duplicate Intune MDM certificates |
| `Detect-OfficeUpdateChannel.ps1` | Detection | No | No | Detects Office update channel |
| `Set-OfficeUpdateChannel.ps1` | Remediation | Yes | No | Sets Office to Semi-Annual channel |
| `Get-IntuneDeviceDetails.ps1` | Reporting | No | Yes | Retrieves device details |
| `Get-IntuneDeviceCompliance.ps1` | Reporting | No | Yes | Retrieves device compliance status |
| `Get-IntuneDeviceHealth.ps1` | Reporting | No | Yes | Performs device health checks |
| `Test-IntuneEnrollment.ps1` | Testing | No | No | Tests device enrollment status |
| `Start-MDMSync.ps1` | Management | No | No | Initiates MDM sync session |
| `Get-IntuneAllAppsAssignmentDetails.ps1` | Reporting | No | Yes | Retrieves app assignment details |
| `Get-IntunePolicyAssignments.ps1` | Reporting | No | Yes | Retrieves policy assignments |
| `Get-IntuneWin32AppDetails.ps1` | Reporting | No | Yes | Retrieves Win32 app details |
| `Export-IntuneDeviceReport.ps1` | Reporting | No | Yes | Exports comprehensive device reports |
| `Get-IntuneDeviceLogs.ps1` | Reporting | Yes | No | Collects Intune-related logs |
| `Disable-PrintSpoolerService.ps1` | Service Management | Yes | No | Disables Print Spooler service |
| `Enable-PrintSpoolerService.ps1` | Service Management | Yes | No | Enables Print Spooler service |
| `Disable-SmartCardLogonEnforcement.ps1` | Service Management | Yes | No | Disables smart card logon enforcement |
| `Enable-SmartCardLogonEnforcement.ps1` | Service Management | Yes | No | Enables smart card logon enforcement |
| `Test-IntuneConnectivity.ps1` | Testing | No | No | Tests connectivity to Intune endpoints |
| `IntuneCommon.psm1` | Module | No | No | Common functions module |

---

## License

This collection is provided as-is for use in managing Microsoft Intune deployments. Use at your own risk.

## Support

For issues or questions:

1. Review script header documentation for usage examples
2. Check script exit codes and error messages
3. Verify prerequisites and permissions
4. Review logs for detailed error information
5. Use `-Verbose` parameter for detailed output

---

**Last Updated:** 2024
