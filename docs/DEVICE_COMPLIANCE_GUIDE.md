# Device Compliance Reporting Guide

This guide provides comprehensive documentation for the Device Compliance Reporter module, which generates compliance reports from Microsoft Intune using the Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Azure Setup](#azure-setup)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Configuration](#configuration)
- [API Reference](#api-reference)
- [Usage Examples](#usage-examples)
- [Error Handling](#error-handling)
- [Advanced Features](#advanced-features)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Performance Considerations](#performance-considerations)

## Overview

The Device Compliance Reporter is a TypeScript module that provides a robust solution for:

- Authenticating with Microsoft Graph API using Azure AD credentials
- Fetching device compliance status from Microsoft Intune
- Generating comprehensive compliance reports and statistics
- Exporting data to CSV and JSON formats
- Handling errors with automatic retry logic
- Filtering devices by operating system and compliance state

### Key Features

1. **Type-Safe**: Full TypeScript support with comprehensive type definitions
2. **Reliable**: Automatic retry logic with exponential backoff
3. **Flexible**: Multiple filtering and configuration options
4. **Scalable**: Handles large device inventories with pagination
5. **Production-Ready**: Comprehensive error handling and logging

## Prerequisites

Before using the Device Compliance Reporter, ensure you have:

### Software Requirements

- Node.js 18.0.0 or higher
- npm or yarn package manager
- TypeScript 5.0 or higher (dev dependency)

### Azure Requirements

- An Azure AD (Entra ID) tenant
- Global Administrator or Intune Administrator access
- Microsoft Intune subscription with enrolled devices

### Knowledge Requirements

- Basic understanding of TypeScript/JavaScript
- Familiarity with Microsoft Graph API
- Understanding of Azure AD app registrations

## Azure Setup

### Step 1: Create Azure AD App Registration

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure the application:
   - **Name**: `Intune Compliance Reporter` (or your preferred name)
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Leave blank (not needed for service principal)
5. Click **Register**

### Step 2: Note Application Details

After registration, note the following from the Overview page:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 3: Create Client Secret

1. In your app registration, go to **Certificates & secrets**
2. Click **New client secret**
3. Add a description: `Compliance Reporter Secret`
4. Select expiration period (recommended: 12-24 months)
5. Click **Add**
6. **IMPORTANT**: Copy the secret value immediately (you won't see it again)

### Step 4: Configure API Permissions

1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions** (not Delegated)
5. Add the following permissions:

   **Required Permissions:**
   - `DeviceManagementManagedDevices.Read.All` - Read managed device properties
   - `DeviceManagementConfiguration.Read.All` - Read device configuration

   **Optional Permissions (for extended features):**
   - `User.Read.All` - Read user profiles
   - `Directory.Read.All` - Read directory data

6. Click **Add permissions**
7. Click **Grant admin consent for [Your Organization]**
8. Confirm the consent grant

### Step 5: Verify Permissions

Ensure the following:
- All required permissions show "Granted for [Your Organization]"
- The app registration has no warning icons
- Admin consent has been successfully granted

## Installation

### 1. Clone or Download the Project

```bash
git clone <repository-url>
cd intunereporting
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Configure Environment Variables

```bash
# Copy the example environment file
cp .env.example .env

# Edit with your credentials
nano .env
```

Add your Azure AD credentials to `.env`:

```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
```

### 4. Build the Project

```bash
npm run build
```

## Quick Start

### Simplest Usage

```typescript
import { createReporterFromEnv } from './src/reports/device-compliance';

// Uses environment variables automatically
const reporter = createReporterFromEnv();

// Generate report
const result = await reporter.generateReport({
  outputDir: './reports'
});

console.log(`Compliance: ${result.summary.compliancePercentage}%`);
```

### Run the Example Script

```bash
# Using ts-node
npm run dev

# Or run compiled version
npm run build && npm start
```

## Configuration

### AuthConfig

Authentication configuration for Microsoft Graph API.

```typescript
interface AuthConfig {
  tenantId: string;      // Azure AD tenant ID
  clientId: string;      // Application (client) ID
  clientSecret: string;  // Client secret value
}
```

### ReportOptions

Options for customizing report generation.

```typescript
interface ReportOptions {
  outputDir?: string;           // Output directory for files
  includeDetails?: boolean;     // Include device details (default: true)
  osFilter?: string;            // Filter by OS (e.g., "Windows", "iOS")
  complianceFilter?: string;    // Filter by state (e.g., "compliant")
  maxRetries?: number;          // Max retry attempts (default: 3)
  retryDelay?: number;          // Retry delay in ms (default: 1000)
}
```

## API Reference

### DeviceComplianceReporter Class

Main class for generating compliance reports.

#### Constructor

```typescript
constructor(authConfig: AuthConfig)
```

Creates a new DeviceComplianceReporter instance.

**Parameters:**
- `authConfig` - Authentication configuration

**Throws:**
- `Error` if configuration is invalid

**Example:**
```typescript
const reporter = new DeviceComplianceReporter({
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret'
});
```

#### generateReport()

```typescript
async generateReport(options?: ReportOptions): Promise<ReportResult>
```

Generates a comprehensive compliance report.

**Parameters:**
- `options` - Optional report configuration

**Returns:**
- `Promise<ReportResult>` - Report result with summary and devices

**Throws:**
- `Error` if report generation fails

**Example:**
```typescript
const result = await reporter.generateReport({
  outputDir: './reports',
  includeDetails: true,
  maxRetries: 5
});
```

#### printSummary() (Static)

```typescript
static printSummary(summary: ComplianceSummary): void
```

Prints formatted compliance summary to console.

**Parameters:**
- `summary` - Compliance summary to print

**Example:**
```typescript
DeviceComplianceReporter.printSummary(result.summary);
```

### Helper Functions

#### createReporterFromEnv()

```typescript
function createReporterFromEnv(): DeviceComplianceReporter
```

Creates a reporter instance from environment variables.

**Environment Variables Required:**
- `AZURE_TENANT_ID`
- `AZURE_CLIENT_ID`
- `AZURE_CLIENT_SECRET`

**Returns:**
- `DeviceComplianceReporter` instance

**Throws:**
- `Error` if environment variables are missing

**Example:**
```typescript
const reporter = createReporterFromEnv();
```

#### generateComplianceReport()

```typescript
async function generateComplianceReport(
  authConfig: AuthConfig,
  outputDir?: string
): Promise<ReportResult>
```

Quick function to generate a compliance report.

**Parameters:**
- `authConfig` - Authentication configuration
- `outputDir` - Optional output directory

**Returns:**
- `Promise<ReportResult>` - Report result

**Example:**
```typescript
const result = await generateComplianceReport(
  { tenantId, clientId, clientSecret },
  './reports'
);
```

## Usage Examples

### Example 1: Basic Report Generation

```typescript
import { DeviceComplianceReporter } from './src/reports/device-compliance';

const reporter = new DeviceComplianceReporter({
  tenantId: process.env.AZURE_TENANT_ID!,
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!
});

const result = await reporter.generateReport({
  outputDir: './reports',
  includeDetails: true
});

console.log(`Total Devices: ${result.summary.totalDevices}`);
console.log(`Compliant: ${result.summary.compliantDevices}`);
console.log(`Compliance Rate: ${result.summary.compliancePercentage}%`);
```

### Example 2: Filter by Operating System

```typescript
// Windows devices only
const windowsResult = await reporter.generateReport({
  outputDir: './reports/windows',
  osFilter: 'Windows'
});

// iOS devices only
const iosResult = await reporter.generateReport({
  outputDir: './reports/ios',
  osFilter: 'iOS'
});

// Android devices only
const androidResult = await reporter.generateReport({
  outputDir: './reports/android',
  osFilter: 'Android'
});
```

### Example 3: Filter by Compliance State

```typescript
// Get only non-compliant devices
const nonCompliantResult = await reporter.generateReport({
  outputDir: './reports/non-compliant',
  complianceFilter: 'noncompliant',
  includeDetails: true
});

// List non-compliant devices
console.log('Non-Compliant Devices:');
nonCompliantResult.devices.forEach(device => {
  console.log(`- ${device.deviceName} (${device.userPrincipalName})`);
});
```

### Example 4: Custom Analysis

```typescript
const result = await reporter.generateReport({ includeDetails: true });

// Find devices not checked in recently
const staleDevices = result.devices.filter(device => {
  const lastReported = new Date(device.lastReportedDateTime);
  const daysSince = (Date.now() - lastReported.getTime()) / (1000 * 60 * 60 * 24);
  return daysSince > 7;
});

console.log(`Devices not checked in for 7+ days: ${staleDevices.length}`);

// Analyze by manufacturer
const byManufacturer = result.devices.reduce((acc, device) => {
  const mfg = device.manufacturer || 'Unknown';
  if (!acc[mfg]) acc[mfg] = { total: 0, compliant: 0 };
  acc[mfg].total++;
  if (device.complianceState === 'compliant') acc[mfg].compliant++;
  return acc;
}, {} as Record<string, { total: number; compliant: number }>);
```

### Example 5: Error Handling with Retry

```typescript
try {
  const result = await reporter.generateReport({
    outputDir: './reports',
    maxRetries: 5,        // Retry up to 5 times
    retryDelay: 2000      // Wait 2 seconds between retries
  });

  console.log('Report generated successfully');
} catch (error) {
  if (error instanceof Error) {
    console.error('Failed to generate report:', error.message);

    // Handle specific errors
    if (error.message.includes('Authentication failed')) {
      console.error('Check your Azure AD credentials');
    } else if (error.message.includes('Failed after')) {
      console.error('Max retries exceeded. Check network connectivity');
    }
  }
}
```

### Example 6: Scheduled Reports

```typescript
import { CronJob } from 'cron';

// Generate report daily at 8 AM
const job = new CronJob('0 8 * * *', async () => {
  console.log('Generating daily compliance report...');

  try {
    const result = await reporter.generateReport({
      outputDir: './reports/daily'
    });

    // Send summary via email or notification system
    sendNotification(result.summary);
  } catch (error) {
    console.error('Failed to generate scheduled report:', error);
  }
});

job.start();
```

## Error Handling

The module implements comprehensive error handling:

### Authentication Errors

```typescript
try {
  const reporter = new DeviceComplianceReporter(config);
  const result = await reporter.generateReport();
} catch (error) {
  if (error.message.includes('Authentication failed')) {
    // Check credentials
    console.error('Verify your tenant ID, client ID, and client secret');
  }
}
```

### Network/API Errors

Automatically retried with exponential backoff:

```typescript
const result = await reporter.generateReport({
  maxRetries: 5,      // Default: 3
  retryDelay: 1000    // Default: 1000ms, increases exponentially
});
```

### Validation Errors

Input validation with helpful error messages:

```typescript
try {
  const reporter = new DeviceComplianceReporter({
    tenantId: '',  // Invalid: empty
    clientId: 'valid-id',
    clientSecret: 'valid-secret'
  });
} catch (error) {
  console.error(error.message); // "Tenant ID is required"
}
```

## Advanced Features

### Pagination Handling

The module automatically handles pagination for large device inventories:

```typescript
// Automatically fetches all pages
const result = await reporter.generateReport();
console.log(`Retrieved ${result.devices.length} devices`);
```

### CSV Export Customization

CSV files include proper escaping for special characters:

```typescript
// Devices with commas, quotes, or newlines in names are properly escaped
const result = await reporter.generateReport({
  outputDir: './reports'
});
// CSV file at: ./reports/device-compliance-[timestamp].csv
```

### Summary Statistics

Detailed compliance statistics:

```typescript
const result = await reporter.generateReport();

console.log('Overall Statistics:');
console.log(`Total: ${result.summary.totalDevices}`);
console.log(`Compliant: ${result.summary.compliantDevices}`);
console.log(`Non-Compliant: ${result.summary.nonCompliantDevices}`);
console.log(`In Grace Period: ${result.summary.inGracePeriodDevices}`);
console.log(`Unknown/Error: ${result.summary.unknownDevices + result.summary.errorDevices}`);

console.log('\nBy Operating System:');
for (const [os, stats] of Object.entries(result.summary.byOperatingSystem)) {
  console.log(`${os}: ${stats.compliant}/${stats.total} (${stats.compliancePercentage}%)`);
}
```

## Best Practices

### 1. Secure Credential Management

```typescript
// ✅ GOOD: Use environment variables
const reporter = createReporterFromEnv();

// ❌ BAD: Hardcode credentials
const reporter = new DeviceComplianceReporter({
  tenantId: 'hardcoded-value',
  // ...
});
```

### 2. Error Handling

```typescript
// ✅ GOOD: Comprehensive error handling
try {
  const result = await reporter.generateReport();
  DeviceComplianceReporter.printSummary(result.summary);
} catch (error) {
  console.error('Report generation failed:', error);
  // Notify administrators, log to monitoring system, etc.
}
```

### 3. Resource Cleanup

```typescript
// ✅ GOOD: Generate report once, reuse data
const result = await reporter.generateReport({ includeDetails: true });
processDevices(result.devices);
generateCharts(result.summary);
sendNotifications(result.summary);
```

### 4. Rate Limiting

```typescript
// ✅ GOOD: Use retry logic for rate limits
const result = await reporter.generateReport({
  maxRetries: 5,
  retryDelay: 2000
});
```

### 5. Filtering for Performance

```typescript
// ✅ GOOD: Filter at API level
const result = await reporter.generateReport({
  osFilter: 'Windows'
});

// ❌ LESS EFFICIENT: Filter after fetching
const allDevices = await reporter.generateReport();
const windowsDevices = allDevices.devices.filter(d => d.operatingSystem === 'Windows');
```

## Troubleshooting

### Issue: Authentication Failed

**Symptoms:** Error message containing "Authentication failed"

**Solutions:**
1. Verify credentials in `.env` file
2. Check that client secret hasn't expired
3. Ensure app registration exists in Azure AD
4. Verify tenant ID is correct

### Issue: Insufficient Permissions

**Symptoms:** Error about missing permissions or access denied

**Solutions:**
1. Verify API permissions are configured
2. Ensure admin consent has been granted
3. Check that required permissions include:
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementConfiguration.Read.All`

### Issue: No Devices Returned

**Symptoms:** Report shows 0 devices

**Solutions:**
1. Verify devices are enrolled in Intune
2. Check app has correct permissions
3. Ensure service account has Intune access
4. Try removing filters (osFilter, complianceFilter)

### Issue: Rate Limiting

**Symptoms:** Errors after multiple rapid requests

**Solutions:**
1. Increase retry delay: `retryDelay: 5000`
2. Add delays between multiple reports
3. Consider caching results

### Issue: Timeout Errors

**Symptoms:** Requests timing out

**Solutions:**
1. Check network connectivity
2. Increase max retries
3. Check Microsoft Graph API status

## Performance Considerations

### Large Device Inventories

For organizations with 1000+ devices:

```typescript
// Use pagination (handled automatically)
const result = await reporter.generateReport({
  includeDetails: true
});

// Process in batches if needed
const batchSize = 100;
for (let i = 0; i < result.devices.length; i += batchSize) {
  const batch = result.devices.slice(i, i + batchSize);
  await processBatch(batch);
}
```

### Memory Usage

For very large reports:

```typescript
// Option 1: Don't include details if not needed
const result = await reporter.generateReport({
  includeDetails: false  // Only get summary
});

// Option 2: Process and discard
const result = await reporter.generateReport({
  outputDir: './reports'  // Write to files
});
// Files are saved, can process later
```

### Network Optimization

```typescript
// Optimize retry settings based on your network
const result = await reporter.generateReport({
  maxRetries: 3,        // Fewer retries for fast networks
  retryDelay: 500       // Shorter delay for fast networks
});
```

## Security Notes

1. **Never commit `.env` file** - Add to `.gitignore`
2. **Rotate secrets regularly** - Set expiration when creating
3. **Use least privilege** - Only grant required permissions
4. **Monitor usage** - Check Azure AD sign-in logs
5. **Use Azure Key Vault** - For production environments

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Intune Device Management](https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Microsoft Graph Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference)

## Support

For issues, questions, or contributions:
1. Check this guide's troubleshooting section
2. Review example code in `src/reports/examples/`
3. Check test file for additional usage patterns
4. Open an issue in the repository

---

**Version:** 1.0.0
**Last Updated:** 2026-02-04
