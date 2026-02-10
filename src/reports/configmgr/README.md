# Configuration Manager Device Management Reports

Comprehensive orchestrator for all 37 Configuration Manager device management reports, modernized for Microsoft Intune using Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Complete Report Catalog](#complete-report-catalog)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Examples](#usage-examples)
- [Report Categories](#report-categories)
- [API Reference](#api-reference)
- [Migration Guide](#migration-guide)
- [Troubleshooting](#troubleshooting)

## Overview

This module provides a unified orchestrator that replicates all 37 Configuration Manager device management reports using modern Microsoft Graph API endpoints. It offers:

- **Unified API**: Single interface for all 37 reports
- **Batch Execution**: Run multiple reports in parallel
- **Dashboard Generation**: Comprehensive metrics across all reports
- **Export Formats**: JSON, CSV, and HTML output
- **Scheduling Support**: Integrate with task schedulers
- **Progress Tracking**: Monitor batch execution progress
- **Error Handling**: Robust error handling and logging

## Complete Report Catalog

### All 37 Reports at a Glance

| # | Report Name | Category | Supported |
|---|-------------|----------|-----------|
| 1 | All corporate-owned mobile devices | Mobile Inventory | ✅ |
| 2 | All mobile device clients | Mobile Inventory | ✅ |
| 3 | Certificate issues (Windows CE) | Windows CE | ❌ Legacy |
| 4 | Client deployment failure (Windows CE) | Windows CE | ❌ Legacy |
| 5 | Client deployment status (Windows CE) | Windows CE | ❌ Legacy |
| 6 | Client deployment success (Windows CE) | Windows CE | ❌ Legacy |
| 7 | Communication issues (Windows CE) | Windows CE | ❌ Legacy |
| 8 | ActiveSync mailbox policy compliance | Exchange ActiveSync | ✅ |
| 9 | Devices by display configurations | Hardware | ✅ |
| 10 | Devices by operating system | Hardware | ✅ |
| 11 | Devices by program memory | Hardware | ✅ |
| 12 | Devices by storage memory | Hardware | ✅ |
| 13 | Health information (Windows CE) | Windows CE | ❌ Legacy |
| 14 | Health summary (Windows CE) | Windows CE | ❌ Legacy |
| 15 | Inactive mobile devices (Exchange) | Exchange ActiveSync | ✅ |
| 16 | Health attestation report | Security | ✅ |
| 17 | Devices enrolled per user (detailed) | Enrollment | ✅ |
| 18 | Devices in specific category | Device Category | ✅ |
| 19 | Local client issues (Windows CE) | Windows CE | ❌ Legacy |
| 20 | Mobile device client information | Mobile Inventory | ✅ |
| 21 | Mobile device compliance details | Exchange ActiveSync | ✅ |
| 22 | Mobile devices by OS | Mobile Inventory | ✅ |
| 23 | Jailbroken or rooted devices | Security | ✅ |
| 24 | Unmanaged devices (enrollment failed) | Device Category | ✅ |
| 25 | Devices with specific free memory | Hardware | ✅ |
| 26 | Devices with specific free storage | Hardware | ✅ |
| 27 | Certificate renewal issues | Security | ✅ |
| 28 | Devices with low free memory | Hardware | ✅ |
| 29 | Devices with low free storage | Hardware | ✅ |
| 30 | Device counts per user (summary) | Enrollment | ✅ |
| 31 | Pending retire and wipe requests | Device Actions | ✅ |
| 32 | Recently enrolled devices | Device Actions | ✅ |
| 33 | Recently wiped devices | Device Actions | ✅ |
| 34 | Settings summary (Exchange) | Exchange ActiveSync | ✅ |
| 35 | Windows RT sideloading keys | Sideloading | ✅ |
| 36 | Windows RT sideloading keys status | Sideloading | ✅ |

**Summary**: 29 supported reports, 8 legacy Windows CE reports (unsupported)

## Installation

```bash
# Install dependencies
npm install

# Import the orchestrator
import { ConfigMgrDeviceManagementReports } from './src/reports/configmgr';
```

## Quick Start

### Basic Usage

```typescript
import { ConfigMgrDeviceManagementReports } from './src/reports/configmgr';
import { Client } from '@microsoft/microsoft-graph-client';

// Initialize Graph client and config
const graphClient = Client.init({/* auth config */});
const config = {/* app config */};

// Create orchestrator
const reports = new ConfigMgrDeviceManagementReports(graphClient, config);

// List all available reports
const allReports = reports.getAllReports();
console.log(`Available reports: ${allReports.length}`);

// Execute a specific report
const report1 = await reports.getReportByNumber(1);
console.log(`Report 1: ${report1.data.length} corporate devices found`);

// Generate comprehensive dashboard
const dashboard = await reports.generateDashboard();
console.log(`Total devices: ${dashboard.summary.totalDevices}`);
console.log(`Compliant: ${dashboard.summary.compliantDevices}`);
```

### Export All Reports

```typescript
// Export all supported reports to JSON
const results = await reports.exportAllReports({
  outputFormat: 'json',
  outputDirectory: './output/reports',
  parallel: true,
  maxParallel: 5
});

console.log(`Generated ${results.successCount} reports successfully`);
console.log(`Total execution time: ${results.summary.executionTimeMs}ms`);
```

## Usage Examples

### Example 1: Corporate Device Inventory

```typescript
// Report 1: All corporate-owned mobile devices
const corporateDevices = await reports.getReportByNumber(1, {
  operatingSystem: 'iOS'
});

console.log(`Corporate iOS devices: ${corporateDevices.data.length}`);
corporateDevices.data.forEach(device => {
  console.log(`- ${device.deviceName} (${device.model})`);
});
```

### Example 2: Security Audit

```typescript
// Report 23: Jailbroken/rooted devices
const securityReport = await reports.getReportByNumber(23);

if (securityReport.data.length > 0) {
  console.log('⚠️  Security Alert: Compromised devices detected');
  securityReport.data.forEach(device => {
    console.log(`- ${device.deviceName}: ${device.riskLevel} risk`);
  });
}
```

### Example 3: Hardware Analysis

```typescript
// Reports 9-12: Complete hardware analysis
const hardwareReports = await Promise.all([
  reports.getReportByNumber(9),  // Display configurations
  reports.getReportByNumber(10), // OS distribution
  reports.getReportByNumber(11), // Memory ranges
  reports.getReportByNumber(12), // Storage ranges
]);

console.log('Hardware Analysis:');
hardwareReports.forEach((report, index) => {
  console.log(`Report ${index + 9}: ${report.data.length} records`);
});
```

### Example 4: Batch Export by Category

```typescript
// Export all security reports
const securityReports = reports.getReportsByCategory(
  ReportCategory.SECURITY
);

const reportNumbers = securityReports.map(r => r.reportNumber);
const results = await reports.exportReports(reportNumbers, {
  outputFormat: 'html',
  outputDirectory: './security-audit',
  includeTimestamp: true
});

console.log(`Exported ${results.successCount} security reports`);
```

### Example 5: User Enrollment Analysis

```typescript
// Report 17 & 30: User enrollment analysis
const userDevices = await reports.getReportByNumber(17, {
  userId: 'user@contoso.com'
});

const allUserCounts = await reports.getReportByNumber(30);

console.log(`User has ${userDevices.data[0].totalDevices} devices`);
console.log(`Total users: ${allUserCounts.data.length}`);
```

### Example 6: Device Action Monitoring

```typescript
// Reports 31-33: Device action tracking
const actionReports = await Promise.all([
  reports.getReportByNumber(31), // Pending retire/wipe
  reports.getReportByNumber(32), // Recently enrolled
  reports.getReportByNumber(33), // Recently wiped
]);

console.log(`Pending actions: ${actionReports[0].data.length}`);
console.log(`New enrollments (7d): ${actionReports[1].data.length}`);
console.log(`Recent wipes (30d): ${actionReports[2].data.length}`);
```

## Report Categories

### Mobile Device Inventory (Reports 1, 2, 20, 22)

Track and analyze all mobile devices in your Intune tenant.

```typescript
// Get all supported reports in this category
const inventoryReports = reports.getReportsByCategory(
  ReportCategory.MOBILE_INVENTORY
);

// Execute all inventory reports
for (const report of inventoryReports) {
  const data = await reports.getReportByNumber(report.reportNumber);
  console.log(`${report.reportName}: ${data.data.length} records`);
}
```

**Modern Intune Equivalent**: Devices > All devices

### Device Hardware (Reports 9-12, 25-26, 28-29)

Analyze hardware configurations, memory, and storage.

```typescript
// Find devices with low storage
const lowStorage = await reports.getReportByNumber(29, {
  thresholdGB: 10
});

console.log(`Devices with < 10GB free: ${lowStorage.data.length}`);
```

**Modern Intune Equivalent**: Devices > All devices > Hardware properties

### Exchange ActiveSync (Reports 8, 15, 21, 34)

Monitor Exchange ActiveSync device compliance and activity.

```typescript
// Check inactive devices
const inactiveDevices = await reports.getReportByNumber(15, {
  inactiveDays: 30
});

console.log(`Inactive devices (30d): ${inactiveDevices.data.length}`);
```

**Modern Intune Equivalent**: Devices > Compliance policies

### Device Security (Reports 16, 23, 27)

Security and compliance monitoring.

```typescript
// Security audit: jailbroken + certificate issues
const [jailbroken, certIssues] = await Promise.all([
  reports.getReportByNumber(23),
  reports.getReportByNumber(27)
]);

console.log(`Security Issues:`);
console.log(`- Jailbroken: ${jailbroken.data.length}`);
console.log(`- Certificate issues: ${certIssues.data.length}`);
```

**Modern Intune Equivalent**: Devices > Monitor > Device health

### Device Enrollment (Reports 17, 30)

Track device enrollment per user.

```typescript
// Enrollment analysis
const enrollmentSummary = await reports.getReportByNumber(30);

const topUsers = enrollmentSummary.data
  .sort((a, b) => b.totalDeviceCount - a.totalDeviceCount)
  .slice(0, 10);

console.log('Top 10 users by device count:');
topUsers.forEach(user => {
  console.log(`${user.displayName}: ${user.totalDeviceCount} devices`);
});
```

**Modern Intune Equivalent**: Reports > Device enrollment

### Device Actions (Reports 31-33)

Monitor device lifecycle actions.

```typescript
// Track recent activity
const recentActivity = await reports.getReportByNumber(32, {
  dateRange: {
    startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
    endDate: new Date()
  }
});

console.log(`New enrollments this week: ${recentActivity.data.length}`);
```

**Modern Intune Equivalent**: Devices > All devices > Device actions

### Device Category (Reports 18, 24)

Manage device categorization and enrollment failures.

```typescript
// Get devices in a specific category
const categoryDevices = await reports.getReportByNumber(18, {
  categoryId: 'executive-devices'
});

console.log(`Executive devices: ${categoryDevices.data.length}`);
```

**Modern Intune Equivalent**: Devices > Device categories

### Windows RT Sideloading (Reports 35-36)

Manage Windows RT sideloading keys.

```typescript
// Check sideloading key status
const keyStatus = await reports.getReportByNumber(36);

console.log(`Active keys: ${keyStatus.data.filter(k => k.status === 'active').length}`);
```

**Modern Intune Equivalent**: Tenant administration > Sideloading keys

### Windows CE Legacy (Reports 3-7, 13-14, 19)

**Status**: Not supported (Windows CE is end-of-life)

```typescript
// Attempting to run Windows CE reports returns migration guidance
const ceReport = await reports.getReportByNumber(3);

console.log(ceReport.summary.warning);
// "This report is not supported in modern Intune"

console.log(ceReport.summary.migrationNotes);
// "Migrate to Windows 10 IoT, Android, or iOS devices"
```

## API Reference

### ConfigMgrDeviceManagementReports Class

#### Constructor

```typescript
constructor(graphClient: Client, config: AppConfig)
```

#### Methods

##### getAllReports()

Returns metadata for all 37 reports.

```typescript
getAllReports(): ConfigMgrReportMetadata[]
```

##### getReportByNumber()

Execute a specific report by number (1-37).

```typescript
async getReportByNumber(
  reportNumber: number,
  options?: Record<string, any>
): Promise<ReportData>
```

**Parameters**:
- `reportNumber`: Report number (1-37)
- `options`: Report-specific options (filters, parameters)

##### getReportsByCategory()

Get all reports in a specific category.

```typescript
getReportsByCategory(category: ReportCategory): ConfigMgrReportMetadata[]
```

##### exportAllReports()

Export all supported reports in batch.

```typescript
async exportAllReports(
  options?: ReportExecutionOptions
): Promise<BatchExecutionResult>
```

**Options**:
```typescript
interface ReportExecutionOptions {
  outputFormat?: 'json' | 'csv' | 'html';
  outputDirectory?: string;
  includeTimestamp?: boolean;
  parallel?: boolean;
  maxParallel?: number;
  filters?: Record<string, any>;
}
```

##### exportReports()

Export specific reports.

```typescript
async exportReports(
  reportNumbers: number[],
  options?: ReportExecutionOptions
): Promise<BatchExecutionResult>
```

##### generateDashboard()

Generate comprehensive dashboard with all metrics.

```typescript
async generateDashboard(): Promise<ConfigMgrDashboard>
```

**Returns**:
```typescript
interface ConfigMgrDashboard {
  generated: string;
  tenantId: string;
  summary: {
    totalDevices: number;
    totalUsers: number;
    corporateDevices: number;
    compliantDevices: number;
    activeDevices: number;
    // ... more metrics
  };
  byCategory: Record<ReportCategory, any>;
  topInsights: InsightItem[];
  recommendations: string[];
}
```

##### getSupportedReports()

Get only supported reports (excludes Windows CE).

```typescript
getSupportedReports(): ConfigMgrReportMetadata[]
```

##### getUnsupportedReports()

Get legacy/unsupported reports.

```typescript
getUnsupportedReports(): ConfigMgrReportMetadata[]
```

### Report Metadata

```typescript
interface ConfigMgrReportMetadata {
  reportNumber: number;
  reportName: string;
  description: string;
  category: ReportCategory;
  parameters?: string[];
  modernIntuneEquivalent: string;
  graphApiEndpoints: string[];
  isSupported: boolean;
  migrationNotes?: string;
}
```

## Migration Guide

### From Configuration Manager to Intune

#### Report Mapping

| ConfigMgr Report | Intune Equivalent | Notes |
|-----------------|-------------------|-------|
| All corporate devices | Devices > Filter by Owner | Direct mapping |
| Mobile devices by OS | Reports > OS distribution | Built-in report |
| Jailbroken devices | Compliance > Non-compliant | Filter by jailbreak |
| Pending wipe requests | Devices > Device actions | Action status |

#### Migration Steps

1. **Identify Required Reports**
   ```typescript
   const requiredReports = [1, 2, 20, 22, 23, 31];
   const metadata = requiredReports.map(n =>
     reports.getReportMetadata(n)
   );
   ```

2. **Test Report Execution**
   ```typescript
   for (const reportNum of requiredReports) {
     try {
       const data = await reports.getReportByNumber(reportNum);
       console.log(`✅ Report ${reportNum}: ${data.data.length} records`);
     } catch (error) {
       console.log(`❌ Report ${reportNum}: ${error.message}`);
     }
   }
   ```

3. **Schedule Regular Execution**
   ```typescript
   // Use cron or task scheduler
   const schedule = {
     reportNumbers: [1, 2, 20, 22],
     cronExpression: '0 0 * * *', // Daily at midnight
     outputFormat: 'html',
     enabled: true
   };

   reports.scheduleReports(schedule);
   ```

### Windows CE Migration

Windows CE devices (Reports 3-7, 13-14, 19) are **not supported** in Intune.

**Recommended Migration Path**:

1. **For Industrial/Embedded Scenarios**:
   - Migrate to Windows 10 IoT Enterprise
   - Migrate to Windows 10 IoT Core

2. **For Mobile Device Scenarios**:
   - Migrate to iOS devices
   - Migrate to Android Enterprise devices

3. **For General Purpose**:
   - Upgrade to Windows 10/11

**Resources**:
- [Windows IoT Documentation](https://docs.microsoft.com/windows/iot-core/)
- [Android Enterprise Migration Guide](https://docs.microsoft.com/mem/intune/enrollment/android-enterprise-overview)

## Troubleshooting

### Common Issues

#### Issue: "Report execution failed"

**Solution**:
```typescript
try {
  const report = await reports.getReportByNumber(1);
} catch (error) {
  console.error('Report failed:', error.message);

  // Check report metadata
  const metadata = reports.getReportMetadata(1);
  console.log('Is supported?', metadata?.isSupported);

  // Verify Graph API permissions
  console.log('Required endpoints:', metadata?.graphApiEndpoints);
}
```

#### Issue: "Graph API permission denied"

**Required Permissions**:
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `User.Read.All`

**Solution**:
```bash
# Grant admin consent in Azure Portal
# Application > API permissions > Grant admin consent
```

#### Issue: "Batch export timeout"

**Solution**:
```typescript
// Reduce parallel execution
const results = await reports.exportAllReports({
  parallel: true,
  maxParallel: 3, // Reduce from default 5
  outputFormat: 'json'
});
```

#### Issue: "Memory issues with large datasets"

**Solution**:
```typescript
// Process reports sequentially
const results = await reports.exportReports(
  [1, 2, 20, 22],
  {
    parallel: false, // Sequential execution
    outputFormat: 'csv' // More memory efficient
  }
);
```

### Debug Mode

Enable verbose logging:

```typescript
import { Logger } from '../../core/logger';

const logger = Logger.getInstance();
logger.setLevel('debug');

// Now all report operations will log detailed info
const report = await reports.getReportByNumber(1);
```

### Performance Optimization

```typescript
// Cache frequently accessed data
const dashboard = await reports.generateDashboard();
// Dashboard aggregates multiple reports efficiently

// Use filters to reduce data volume
const filteredReport = await reports.getReportByNumber(1, {
  operatingSystem: 'iOS',
  complianceState: 'compliant'
});

// Batch export with parallel execution
const results = await reports.exportAllReports({
  parallel: true,
  maxParallel: 5, // Adjust based on API throttling
  outputDirectory: './output'
});
```

## Advanced Features

### Custom Report Filters

```typescript
// Apply advanced filters
const report = await reports.getReportByNumber(1, {
  ownerType: 'company',
  operatingSystem: 'iOS',
  minLastSyncDays: 0,
  maxLastSyncDays: 7,
  complianceState: 'compliant'
});
```

### Progress Tracking

```typescript
// Track batch execution progress
const reportNumbers = [1, 2, 3, 4, 5];
let completed = 0;

for (const num of reportNumbers) {
  const result = await reports.getReportByNumber(num);
  completed++;
  console.log(`Progress: ${completed}/${reportNumbers.length}`);
}
```

### Custom Output Formatting

```typescript
// Export with custom formatting
const report = await reports.getReportByNumber(1);

// Convert to custom format
const customFormat = {
  reportTitle: report.metadata.reportName,
  generatedDate: report.metadata.generatedAt,
  devices: report.data.map(d => ({
    name: d.deviceName,
    os: d.operatingSystem,
    user: d.userPrincipalName
  }))
};

fs.writeFileSync('./custom-report.json', JSON.stringify(customFormat, null, 2));
```

## Support

For issues, questions, or contributions:

- **Documentation**: See `/docs` directory
- **Examples**: See `/src/reports/configmgr/examples`
- **Tests**: See `/src/reports/configmgr/*.test.ts`

## License

Copyright (c) 2024 Intune Reporting System
