# Windows RT Sideloading Key Reports - Configuration Manager Migration Guide

This guide provides comprehensive documentation for the Windows RT Sideloading Key Reports module, which replicates Configuration Manager reports 35 and 36 using Microsoft Graph API for Microsoft Intune.

## Table of Contents

- [Overview](#overview)
- [Configuration Manager Reports Implemented](#configuration-manager-reports-implemented)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
- [Usage Examples](#usage-examples)
- [Report Details](#report-details)
- [Alert System](#alert-system)
- [Export Formats](#export-formats)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Migration Notes](#migration-notes)

## Overview

The Windows RT Sideloading Key Reports module provides comprehensive tracking and analysis of Windows RT sideloading keys in your Intune environment. This module helps you:

- Monitor sideloading key usage and activation counts
- Track key expiration and identify keys nearing expiration
- Generate alerts for high usage, expiring, and unused keys
- Identify revoked or expired keys
- Export detailed reports in multiple formats (JSON, CSV, HTML)

### Key Features

1. **Complete Key Tracking**: Monitor all aspects of sideloading keys including activation counts, status, and expiration
2. **Automated Alerts**: Built-in alerting for keys requiring attention (expiring, high usage, unused)
3. **Flexible Filtering**: Filter keys by status, activation counts, and usage percentages
4. **Multiple Export Formats**: Export reports as JSON, CSV, or HTML
5. **Type-Safe**: Full TypeScript support with comprehensive interfaces
6. **Production-Ready**: Includes retry logic, error handling, and comprehensive logging

## Configuration Manager Reports Implemented

### Report 35: Windows RT Sideloading Keys Detailed Status

This report provides per-key detailed information including:
- Key value (masked for security)
- Total activations and remaining activations
- Last updated date and days since update
- Current status (active/expired/revoked)
- Usage percentage
- Expiration warnings
- High usage indicators

**Configuration Manager Equivalent:**
```
Report Name: Windows RT sideloading key detailed status
Report Path: \Device Management\Windows RT Sideloading
Report ID: 35
```

### Report 36: Windows RT Sideloading Keys Summary

This report provides aggregated statistics including:
- Total keys count
- Keys grouped by status (active, expired, revoked)
- Usage statistics across all keys
- Most and least used keys
- Alerts for keys requiring attention

**Configuration Manager Equivalent:**
```
Report Name: Windows RT sideloading key summary
Report Path: \Device Management\Windows RT Sideloading
Report ID: 36
```

## Prerequisites

### Software Requirements

- Node.js 18.0.0 or higher
- TypeScript 5.0 or higher
- Microsoft Graph Client SDK

### Azure/Intune Requirements

- Azure AD (Entra ID) tenant
- Microsoft Intune subscription
- Global Administrator or Intune Administrator access
- Azure AD app registration with appropriate permissions

### Required Microsoft Graph API Permissions

The following **Application permissions** are required:

- `DeviceManagementApps.Read.All` - Read Microsoft Intune apps and policies
- `DeviceManagementApps.ReadWrite.All` - (Optional) For write operations

## Installation

### Step 1: Install Dependencies

```bash
npm install @microsoft/microsoft-graph-client
npm install --save-dev @types/node typescript
```

### Step 2: Configure Azure AD App Registration

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations**
3. Create or select your app registration
4. Add the required API permissions (see above)
5. Grant admin consent for the permissions
6. Create a client secret

### Step 3: Configure Environment Variables

Create a `.env` file with your credentials:

```env
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret
```

### Step 4: Import the Module

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from './src/types';
```

## Quick Start

### Basic Usage

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';

// Initialize Graph client
const graphClient = Client.init({
  authProvider: (done) => {
    // Your authentication logic here
    done(null, accessToken);
  },
});

// Create config
const config: AppConfig = {
  // Your configuration
};

// Initialize report
const report = new WindowsRTSideloadingReports(graphClient, config);

// Get summary of all keys
const summary = await report.getSideloadingKeysSummary();
console.log(`Total Keys: ${summary.totalKeys}`);
console.log(`Active Keys: ${summary.activeKeysCount}`);
console.log(`Expiring Keys: ${summary.alerts.expiringKeysCount}`);

// Get detailed status for a specific key
const detailedStatus = await report.getSideloadingKeyDetailedStatus('key-id');
console.log(`Key Status: ${detailedStatus.status}`);
console.log(`Usage: ${detailedStatus.usagePercentage}%`);
```

## API Reference

### Class: WindowsRTSideloadingReports

Main class for Windows RT sideloading key reporting.

#### Methods

##### `getAllSideloadingKeys(): Promise<SideloadingKey[]>`

Retrieves all sideloading keys from Microsoft Graph API.

**Returns:** Array of sideloading keys

**Example:**
```typescript
const keys = await report.getAllSideloadingKeys();
console.log(`Found ${keys.length} sideloading keys`);
```

##### `getSideloadingKeyById(keyId: string): Promise<SideloadingKey>`

Retrieves a specific sideloading key by its ID.

**Parameters:**
- `keyId` - The ID of the sideloading key

**Returns:** Sideloading key details

**Example:**
```typescript
const key = await report.getSideloadingKeyById('abc-123-def');
console.log(`Key: ${key.displayName}`);
```

##### `getSideloadingKeyDetailedStatus(keyId: string, thresholds?: AlertThresholds): Promise<SideloadingKeyDetailedStatus>`

**Report 35: Windows RT Sideloading Keys Detailed Status**

Generates detailed status information for a specific sideloading key.

**Parameters:**
- `keyId` - The ID of the sideloading key
- `thresholds` - Optional custom alert thresholds

**Returns:** Detailed key status information

**Example:**
```typescript
const status = await report.getSideloadingKeyDetailedStatus('key-001');

console.log(`Key ID: ${status.keyId}`);
console.log(`Display Name: ${status.displayName}`);
console.log(`Status: ${status.status}`);
console.log(`Activations Used: ${status.activationsUsed} / ${status.totalActivations}`);
console.log(`Usage Percentage: ${status.usagePercentage}%`);
console.log(`Days Since Update: ${status.lastUpdatedDaysAgo}`);

if (status.isHighUsage) {
  console.log('⚠️ WARNING: Key is at high usage!');
}

if (status.isNearExpiration) {
  console.log('⚠️ WARNING: Key is near expiration!');
}
```

##### `getSideloadingKeysSummary(filter?: SideloadingKeyFilter, thresholds?: AlertThresholds): Promise<SideloadingKeysSummary>`

**Report 36: Windows RT Sideloading Keys Summary**

Generates an aggregated summary of all sideloading keys.

**Parameters:**
- `filter` - Optional filter criteria
- `thresholds` - Optional custom alert thresholds

**Returns:** Aggregated key status summary

**Example:**
```typescript
const summary = await report.getSideloadingKeysSummary();

console.log(`Total Keys: ${summary.totalKeys}`);
console.log(`Active Keys: ${summary.activeKeysCount}`);
console.log(`Expired Keys: ${summary.expiredKeysCount}`);
console.log(`Revoked Keys: ${summary.revokedKeysCount}`);

console.log('\nUsage Statistics:');
console.log(`Total Activations: ${summary.usageStatistics.totalActivationsAcrossAllKeys}`);
console.log(`Average per Key: ${summary.usageStatistics.averageActivationsPerKey}`);
console.log(`Average Usage: ${summary.usageStatistics.averageUsagePercentage}%`);

if (summary.usageStatistics.mostUsedKey) {
  console.log(`\nMost Used: ${summary.usageStatistics.mostUsedKey.displayName} (${summary.usageStatistics.mostUsedKey.activations} activations)`);
}

console.log('\nAlerts:');
console.log(`Expiring Keys: ${summary.alerts.expiringKeysCount}`);
console.log(`High Usage Keys: ${summary.alerts.highUsageKeysCount}`);
console.log(`Unused Keys: ${summary.alerts.unusedKeysCount}`);
```

##### `generateComprehensiveReport(filter?: SideloadingKeyFilter, thresholds?: AlertThresholds): Promise<ReportData>`

Generates a comprehensive report including both summary and detailed statuses for all keys.

**Parameters:**
- `filter` - Optional filter criteria
- `thresholds` - Optional custom alert thresholds

**Returns:** Complete report data with metadata

**Example:**
```typescript
const report = await report.generateComprehensiveReport();

console.log(`Report Generated: ${report.metadata.generatedAt}`);
console.log(`Total Keys: ${report.summary.totalKeys}`);
console.log(`Active Keys: ${report.summary.activeKeys}`);
```

##### `exportReport(reportData: ReportData, options: ReportExportOptions): Promise<string>`

Exports report data to a file in the specified format.

**Parameters:**
- `reportData` - The report data to export
- `options` - Export options (format, output directory, etc.)

**Returns:** Path to the exported file

**Example:**
```typescript
const reportData = await report.generateComprehensiveReport();

const outputPath = await report.exportReport(reportData, {
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true,
});

console.log(`Report exported to: ${outputPath}`);
```

### Interfaces

#### SideloadingKey

Represents a Windows RT sideloading key from Microsoft Graph API.

```typescript
interface SideloadingKey {
  id: string;
  displayName?: string;
  value?: string;
  totalActivation: number;
  lastUpdatedDateTime?: string;
  description?: string;
}
```

#### SideloadingKeyDetailedStatus

Detailed status information for a sideloading key.

```typescript
interface SideloadingKeyDetailedStatus {
  keyId: string;
  keyValue: string;  // Masked for security
  displayName: string;
  totalActivations: number;
  activationsUsed: number;
  activationsRemaining: number;
  lastUpdatedDate: string;
  lastUpdatedDaysAgo: number;
  status: 'active' | 'expired' | 'revoked' | 'unknown';
  statusReason: string;
  expirationWarning: boolean;
  usagePercentage: number;
  description: string;
  isNearExpiration: boolean;
  isHighUsage: boolean;
}
```

#### SideloadingKeysSummary

Aggregated summary of all sideloading keys.

```typescript
interface SideloadingKeysSummary {
  totalKeys: number;
  activeKeysCount: number;
  expiredKeysCount: number;
  revokedKeysCount: number;
  unknownKeysCount: number;
  keysByStatus: Array<{
    status: 'active' | 'expired' | 'revoked' | 'unknown';
    count: number;
    percentage: number;
    keys: string[];
  }>;
  usageStatistics: {
    totalActivationsAcrossAllKeys: number;
    averageActivationsPerKey: number;
    mostUsedKey: {
      keyId: string;
      displayName: string;
      activations: number;
    } | null;
    leastUsedKey: {
      keyId: string;
      displayName: string;
      activations: number;
    } | null;
    averageUsagePercentage: number;
  };
  alerts: {
    expiringKeysCount: number;
    highUsageKeysCount: number;
    unusedKeysCount: number;
    expiringKeys: string[];
    highUsageKeys: string[];
    unusedKeys: string[];
  };
  lastUpdatedDateTime: string;
}
```

## Usage Examples

### Example 1: Get Summary Report with Custom Thresholds

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';

const report = new WindowsRTSideloadingReports(graphClient, config);

// Get summary with custom alert thresholds
const summary = await report.getSideloadingKeysSummary(undefined, {
  expirationWarningDays: 60,  // Warn 60 days before expiration
  highUsagePercentage: 75,    // Alert at 75% usage
  lowUsagePercentage: 5,      // Alert if usage < 5%
});

// Display alerts
if (summary.alerts.expiringKeysCount > 0) {
  console.log('\n⚠️ EXPIRING KEYS:');
  summary.alerts.expiringKeys.forEach(key => {
    console.log(`  - ${key}`);
  });
}

if (summary.alerts.highUsageKeysCount > 0) {
  console.log('\n⚠️ HIGH USAGE KEYS:');
  summary.alerts.highUsageKeys.forEach(key => {
    console.log(`  - ${key}`);
  });
}

if (summary.alerts.unusedKeysCount > 0) {
  console.log('\n💡 UNUSED KEYS:');
  summary.alerts.unusedKeys.forEach(key => {
    console.log(`  - ${key}`);
  });
}
```

### Example 2: Filter Active Keys and Export to CSV

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';

const report = new WindowsRTSideloadingReports(graphClient, config);

// Get only active keys
const summary = await report.getSideloadingKeysSummary({
  status: 'active',
  includeExpired: false,
  includeRevoked: false,
});

console.log(`Active Keys: ${summary.activeKeysCount}`);

// Generate comprehensive report
const reportData = await report.generateComprehensiveReport({
  status: 'active',
});

// Export to CSV
const csvPath = await report.exportReport(reportData, {
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true,
});

console.log(`Report exported to: ${csvPath}`);
```

### Example 3: Monitor High Usage Keys

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';

const report = new WindowsRTSideloadingReports(graphClient, config);

// Get keys with usage >= 80%
const summary = await report.getSideloadingKeysSummary({
  minUsagePercentage: 80,
});

console.log(`\nHigh Usage Keys (>= 80%):`);
console.log(`Total: ${summary.totalKeys}`);

// Get detailed status for each high usage key
const keys = await report.getAllSideloadingKeys();

for (const key of keys) {
  const status = await report.getSideloadingKeyDetailedStatus(key.id);

  if (status.usagePercentage >= 80) {
    console.log(`\n${status.displayName}:`);
    console.log(`  Usage: ${status.usagePercentage}%`);
    console.log(`  Activations: ${status.activationsUsed} / ${status.totalActivations}`);
    console.log(`  Remaining: ${status.activationsRemaining}`);
    console.log(`  Status: ${status.status}`);

    if (status.usagePercentage >= 95) {
      console.log(`  ⚠️ CRITICAL: Key nearly exhausted!`);
    }
  }
}
```

### Example 4: Detailed Key Analysis

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';

const report = new WindowsRTSideloadingReports(graphClient, config);

// Analyze a specific key
const keyId = 'your-key-id';
const status = await report.getSideloadingKeyDetailedStatus(keyId);

console.log('=== SIDELOADING KEY ANALYSIS ===');
console.log(`\nKey Information:`);
console.log(`  ID: ${status.keyId}`);
console.log(`  Name: ${status.displayName}`);
console.log(`  Value: ${status.keyValue}`);  // Masked
console.log(`  Description: ${status.description}`);

console.log(`\nActivation Details:`);
console.log(`  Total Activations: ${status.totalActivations}`);
console.log(`  Activations Used: ${status.activationsUsed}`);
console.log(`  Activations Remaining: ${status.activationsRemaining}`);
console.log(`  Usage Percentage: ${status.usagePercentage}%`);

console.log(`\nStatus Information:`);
console.log(`  Status: ${status.status}`);
console.log(`  Reason: ${status.statusReason}`);
console.log(`  Last Updated: ${status.lastUpdatedDate}`);
console.log(`  Days Since Update: ${status.lastUpdatedDaysAgo}`);

console.log(`\nAlerts:`);
console.log(`  Near Expiration: ${status.isNearExpiration ? 'YES ⚠️' : 'No'}`);
console.log(`  High Usage: ${status.isHighUsage ? 'YES ⚠️' : 'No'}`);
console.log(`  Expiration Warning: ${status.expirationWarning ? 'YES ⚠️' : 'No'}`);

// Recommendations
console.log(`\nRecommendations:`);
if (status.isNearExpiration) {
  console.log(`  - Consider renewing this key soon`);
}
if (status.isHighUsage) {
  console.log(`  - Prepare a replacement key`);
  console.log(`  - Monitor activation usage closely`);
}
if (status.usagePercentage < 10 && status.lastUpdatedDaysAgo > 90) {
  console.log(`  - This key may not be in use`);
  console.log(`  - Consider reviewing if this key is still needed`);
}
```

### Example 5: Export Reports in Multiple Formats

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';

const report = new WindowsRTSideloadingReports(graphClient, config);

// Generate comprehensive report
const reportData = await report.generateComprehensiveReport();

// Export to all formats
const formats: Array<'json' | 'csv' | 'html'> = ['json', 'csv', 'html'];

for (const format of formats) {
  const outputPath = await report.exportReport(reportData, {
    format,
    outputDir: './reports',
    includeTimestamp: true,
  });

  console.log(`${format.toUpperCase()} report exported to: ${outputPath}`);
}

console.log('\nAll reports exported successfully!');
```

### Example 6: Automated Monitoring Script

```typescript
import { WindowsRTSideloadingReports } from './src/reports/configmgr/windows-rt-sideloading-reports';

async function monitorSideloadingKeys() {
  const report = new WindowsRTSideloadingReports(graphClient, config);

  // Get summary with strict thresholds
  const summary = await report.getSideloadingKeysSummary(undefined, {
    expirationWarningDays: 30,
    highUsagePercentage: 85,
    lowUsagePercentage: 5,
  });

  console.log(`\n===== SIDELOADING KEY MONITORING REPORT =====`);
  console.log(`Generated: ${summary.lastUpdatedDateTime}`);
  console.log(`\nOverview:`);
  console.log(`  Total Keys: ${summary.totalKeys}`);
  console.log(`  Active: ${summary.activeKeysCount}`);
  console.log(`  Expired: ${summary.expiredKeysCount}`);
  console.log(`  Revoked: ${summary.revokedKeysCount}`);

  // Check for critical alerts
  const hasAlerts =
    summary.alerts.expiringKeysCount > 0 ||
    summary.alerts.highUsageKeysCount > 0;

  if (hasAlerts) {
    console.log(`\n⚠️ ALERTS REQUIRE ATTENTION ⚠️`);

    if (summary.alerts.expiringKeysCount > 0) {
      console.log(`\nExpiring Keys (${summary.alerts.expiringKeysCount}):`);
      summary.alerts.expiringKeys.forEach(key => {
        console.log(`  - ${key}`);
      });
    }

    if (summary.alerts.highUsageKeysCount > 0) {
      console.log(`\nHigh Usage Keys (${summary.alerts.highUsageKeysCount}):`);
      summary.alerts.highUsageKeys.forEach(key => {
        console.log(`  - ${key}`);
      });
    }

    // Send email notification (pseudo-code)
    // await sendAlertEmail(summary);
  } else {
    console.log(`\n✅ No alerts - all keys are healthy`);
  }

  // Export report for records
  const reportData = await report.generateComprehensiveReport();
  await report.exportReport(reportData, {
    format: 'json',
    outputDir: './reports/archive',
    includeTimestamp: true,
  });
}

// Run monitoring
monitorSideloadingKeys();
```

## Report Details

### Key Status Types

| Status | Description | Criteria |
|--------|-------------|----------|
| `active` | Key is active and available for use | Key is within validity period and under activation limit |
| `expired` | Key has expired | Key has exceeded 365 days or reached 5000 activations |
| `revoked` | Key has been revoked | Key value is null or empty |
| `unknown` | Status cannot be determined | Insufficient data to determine status |

### Activation Limits

- **Maximum Activations**: 5,000 per key (Windows RT standard)
- **High Usage Threshold**: 80% (4,000 activations) - default, configurable
- **Low Usage Threshold**: 10% (500 activations) - default, configurable

### Expiration Warnings

Keys are flagged for expiration based on:
- **Time-based**: Keys older than 335 days (30 days before 1-year expiration)
- **Usage-based**: Keys at or above 80% activation usage
- **Configurable**: Adjust thresholds via `AlertThresholds` parameter

## Alert System

The module includes a comprehensive alert system that monitors:

### Expiring Keys
- Keys approaching their expiration date (default: 30 days)
- Based on `lastUpdatedDateTime` field
- Configurable via `expirationWarningDays` threshold

### High Usage Keys
- Keys approaching activation limits (default: 80%)
- Calculated as `(activationsUsed / totalActivations) * 100`
- Configurable via `highUsagePercentage` threshold

### Unused Keys
- Keys with minimal usage (default: < 10%)
- May indicate unused or forgotten keys
- Configurable via `lowUsagePercentage` threshold

### Customizing Alert Thresholds

```typescript
const customThresholds = {
  expirationWarningDays: 45,    // Warn 45 days before expiration
  highUsagePercentage: 90,      // Alert at 90% usage
  lowUsagePercentage: 3,        // Alert if usage < 3%
};

const summary = await report.getSideloadingKeysSummary(undefined, customThresholds);
```

## Export Formats

### JSON Format
- Complete structured data
- Ideal for programmatic processing
- Includes all metadata and nested objects

```typescript
await report.exportReport(reportData, {
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true,
});
```

### CSV Format
- Flat structure suitable for Excel/spreadsheets
- Easy to import into other systems
- Best for tabular data analysis

```typescript
await report.exportReport(reportData, {
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true,
});
```

### HTML Format
- Human-readable format
- Styled for easy viewing in browsers
- Includes charts and visualizations

```typescript
await report.exportReport(reportData, {
  format: 'html',
  outputDir: './reports',
  includeTimestamp: true,
});
```

## Best Practices

### 1. Regular Monitoring

Run reports on a scheduled basis to track key usage trends:

```typescript
// Example: Daily monitoring
setInterval(async () => {
  const summary = await report.getSideloadingKeysSummary();

  if (summary.alerts.expiringKeysCount > 0 || summary.alerts.highUsageKeysCount > 0) {
    // Send alert notification
    console.log('Action required: Keys need attention');
  }
}, 24 * 60 * 60 * 1000); // Run daily
```

### 2. Key Lifecycle Management

- Review unused keys quarterly
- Replace keys before they reach 90% usage
- Renew keys at least 30 days before expiration
- Document key purpose in the `description` field

### 3. Security Considerations

- Key values are automatically masked in reports
- Store reports securely if they contain sensitive information
- Regularly audit key access and usage
- Revoke unused or compromised keys immediately

### 4. Performance Optimization

- Cache summary data for frequently accessed reports
- Use filters to reduce data retrieval for large key sets
- Schedule heavy reports during off-peak hours

### 5. Data Retention

- Archive historical reports for compliance
- Maintain at least 12 months of key usage data
- Export reports regularly for backup purposes

## Troubleshooting

### Common Issues

#### Issue: No keys returned

**Possible Causes:**
- Insufficient API permissions
- No sideloading keys configured in Intune
- API authentication failure

**Solution:**
```typescript
// Verify permissions and authentication
const keys = await report.getAllSideloadingKeys();
console.log(`Keys found: ${keys.length}`);

// If zero, check:
// 1. API permissions in Azure AD
// 2. Sideloading keys exist in Intune portal
// 3. Access token is valid
```

#### Issue: Incorrect usage percentages

**Possible Cause:**
- Incorrect maximum activation value
- Data synchronization delay

**Solution:**
```typescript
// Verify activation counts
const status = await report.getSideloadingKeyDetailedStatus(keyId);
console.log(`Activations: ${status.activationsUsed} / ${status.totalActivations}`);
console.log(`Percentage: ${status.usagePercentage}%`);

// Expected: 5000 total activations for Windows RT keys
```

#### Issue: Status showing as 'unknown'

**Possible Cause:**
- Missing or incomplete key data from API
- Key in transitional state

**Solution:**
```typescript
// Check raw key data
const key = await report.getSideloadingKeyById(keyId);
console.log('Raw key data:', key);

// Verify all required fields are present
```

#### Issue: Export fails

**Possible Cause:**
- Insufficient file system permissions
- Invalid output directory
- Disk space issues

**Solution:**
```typescript
// Check directory exists and is writable
import * as fs from 'fs';

const outputDir = './reports';
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// Try export again
await report.exportReport(reportData, {
  format: 'json',
  outputDir,
  includeTimestamp: true,
});
```

### Debugging

Enable detailed logging to troubleshoot issues:

```typescript
// Set log level to debug in config
const config: AppConfig = {
  // ... other config
  logging: {
    level: 'debug',  // Changed from 'info' or 'error'
    console: true,
    // ... other logging config
  },
};
```

## Migration Notes

### Migrating from Configuration Manager

When migrating from Configuration Manager Reports 35 and 36:

#### Data Mapping

| Configuration Manager | Microsoft Graph API | Notes |
|----------------------|---------------------|-------|
| Sideloading Key ID | `id` | Direct mapping |
| Key Value | `value` | Masked in reports for security |
| Total Activations | `totalActivation` | Same concept, different property name |
| Last Updated | `lastUpdatedDateTime` | ISO 8601 format in Graph API |
| Display Name | `displayName` | Direct mapping |

#### Feature Parity

✅ **Implemented:**
- Key status tracking
- Activation count monitoring
- Expiration detection
- Usage percentage calculations
- Summary statistics
- Export to multiple formats

⚠️ **Differences:**
- Configuration Manager may include additional custom fields
- Report scheduling handled differently (use external scheduler)
- Alert notifications require custom implementation

#### Migration Checklist

- [ ] Verify API permissions in Azure AD
- [ ] Test reports with sample keys
- [ ] Compare output with Configuration Manager reports
- [ ] Set up automated scheduling (if needed)
- [ ] Configure alert thresholds to match current policies
- [ ] Export historical data from Configuration Manager
- [ ] Document any custom fields or calculations
- [ ] Train users on new reporting interface

### Extending Functionality

The module can be extended to add custom functionality:

```typescript
// Example: Custom alert handler
class CustomSideloadingReports extends WindowsRTSideloadingReports {
  async getSideloadingKeysSummary(
    filter?: SideloadingKeyFilter,
    thresholds?: AlertThresholds
  ): Promise<SideloadingKeysSummary> {
    const summary = await super.getSideloadingKeysSummary(filter, thresholds);

    // Add custom alerting
    if (summary.alerts.expiringKeysCount > 0) {
      await this.sendAlertEmail(summary);
    }

    return summary;
  }

  private async sendAlertEmail(summary: SideloadingKeysSummary): Promise<void> {
    // Custom email notification logic
  }
}
```

## Additional Resources

### Microsoft Documentation
- [Microsoft Graph API - Sideloading Keys](https://learn.microsoft.com/en-us/graph/api/resources/intune-onboarding-sideloadingkey)
- [Windows RT Sideloading Overview](https://learn.microsoft.com/en-us/windows/application-management/sideload-apps-in-windows-10)
- [Intune App Management](https://learn.microsoft.com/en-us/mem/intune/apps/)

### Related Modules
- Device Hardware Reports
- Device Action Reports
- Mobile Device Inventory Reports

### Support
For issues, questions, or feature requests, please refer to the main project repository.

---

**Last Updated:** 2024-02-06
**Version:** 1.0.0
**Module:** Windows RT Sideloading Reports
