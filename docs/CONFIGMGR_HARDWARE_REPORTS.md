# Configuration Manager Hardware Reports for Intune

This guide provides comprehensive documentation for the Device Hardware Reports module, which replicates Configuration Manager hardware inventory reports using Microsoft Graph API for Intune-managed devices.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Supported Reports](#supported-reports)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
- [Usage Examples](#usage-examples)
- [Microsoft Graph API Endpoints](#microsoft-graph-api-endpoints)
- [Data Models](#data-models)
- [Export Formats](#export-formats)
- [Filtering and Thresholds](#filtering-and-thresholds)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Performance Considerations](#performance-considerations)

## Overview

The Device Hardware Reports module provides a comprehensive solution for hardware inventory and analysis of Intune-managed mobile devices. It replicates the functionality of Configuration Manager hardware reports, adapted for the Microsoft Intune cloud-based MDM environment.

### Key Features

- **Complete Hardware Inventory**: Track device specifications including memory, storage, display configurations, and OS versions
- **Memory Analysis**: Categorize devices by RAM capacity and identify low-memory devices
- **Storage Management**: Monitor storage utilization and identify devices with low available space
- **OS Distribution**: Analyze operating system versions across your device fleet
- **Display Configuration**: Categorize devices by display capabilities
- **Custom Thresholds**: Configure memory and storage thresholds based on your requirements
- **Multiple Export Formats**: Export reports in JSON, CSV, and HTML formats
- **TypeScript Support**: Full type safety with comprehensive TypeScript interfaces
- **Automated Alerts**: Identify devices requiring attention based on configurable thresholds

### ConfigMgr Reports Replicated

This module replicates the following Configuration Manager reports:

| Report # | ConfigMgr Report Name | Implementation Method |
|----------|----------------------|----------------------|
| 9 | Count of mobile devices by display configurations | `getDisplayConfigurationSummary()` |
| 10 | Count of mobile devices by operating system | `getOSDistributionSummary()` |
| 11 | Count of mobile devices by program memory | `getMemoryRangeSummary()` |
| 12 | Count of mobile devices by storage memory | `getStorageRangeSummary()` |
| 25 | Mobile devices with specific free program memory | `getDevicesWithSpecificFreeMemory()` |
| 26 | Mobile devices with specific free removable storage | `getDevicesWithSpecificFreeStorage()` |
| 28 | Mobile devices with low free program memory | `getDevicesWithLowMemory()` |
| 29 | Mobile devices with low free removable storage | `getDevicesWithLowStorage()` |

## Prerequisites

### Software Requirements

- **Node.js**: Version 18.0.0 or higher
- **npm/yarn**: Latest stable version
- **TypeScript**: Version 5.0 or higher (dev dependency)

### Azure Requirements

- **Azure AD Tenant**: Active Azure AD (Entra ID) tenant
- **Microsoft Intune**: Active Intune subscription with enrolled devices
- **Admin Access**: Global Administrator or Intune Administrator role
- **App Registration**: Azure AD app registration with appropriate Graph API permissions

### Required Permissions

The following Microsoft Graph API permissions are required:

**Application Permissions** (recommended for automation):
- `DeviceManagementManagedDevices.Read.All` - Read managed device properties
- `DeviceManagementConfiguration.Read.All` - Read device configuration

**Delegated Permissions** (for interactive scenarios):
- `DeviceManagementManagedDevices.Read` - Read managed device properties

### Knowledge Requirements

- Basic understanding of TypeScript/JavaScript
- Familiarity with Microsoft Graph API
- Understanding of Azure AD app registrations
- Knowledge of device management concepts (memory, storage, OS)

## Supported Reports

### Report 9: Display Configuration Summary

**Description**: Groups mobile devices by display configuration categories based on device tier (memory).

**Use Cases**:
- Identify device capabilities for app deployment
- Plan UI/UX strategies for different device tiers
- Understand device quality distribution

**Method**: `getDisplayConfigurationSummary()`

**Output**: Array of display configuration summaries with device counts and percentages

**Categories**:
- Standard (< 2GB RAM)
- Enhanced (2-4GB RAM)
- High (4-6GB RAM)
- Premium (6-8GB RAM)
- Ultra (8GB+ RAM)

### Report 10: OS Distribution

**Description**: Counts devices by operating system and version.

**Use Cases**:
- Plan OS upgrade strategies
- Identify unsupported OS versions
- Track OS adoption rates
- Security patch compliance

**Method**: `getOSDistributionSummary()`

**Output**: Array of OS distributions with version breakdowns

### Report 11: Memory Range Summary

**Description**: Categorizes devices by program memory (RAM) ranges.

**Use Cases**:
- Assess device capabilities for app deployments
- Identify devices unsuitable for memory-intensive applications
- Plan device refresh cycles
- Budget planning for hardware upgrades

**Method**: `getMemoryRangeSummary()`

**Output**: Device counts categorized into memory ranges

**Ranges**:
- < 1 GB
- 1-2 GB
- 2-4 GB
- 4-6 GB
- 6-8 GB
- 8-16 GB
- \> 16 GB

### Report 12: Storage Range Summary

**Description**: Categorizes devices by total storage capacity.

**Use Cases**:
- Plan content and app distribution strategies
- Identify devices with limited storage
- Optimize app deployment based on storage availability
- Device procurement planning

**Method**: `getStorageRangeSummary()`

**Output**: Device counts categorized into storage ranges

**Ranges**:
- < 16 GB
- 16-32 GB
- 32-64 GB
- 64-128 GB
- 128-256 GB
- 256-512 GB
- 512 GB - 1 TB
- \> 1 TB

### Report 25: Devices with Specific Free Memory

**Description**: Finds devices with free RAM within a specified range.

**Use Cases**:
- Identify devices suitable for specific workloads
- Target devices for memory-intensive app deployments
- Performance troubleshooting
- Capacity planning

**Method**: `getDevicesWithSpecificFreeMemory(minMB, maxMB?)`

**Parameters**:
- `minMB`: Minimum free memory in MB
- `maxMB`: Maximum free memory in MB (optional)

### Report 26: Devices with Specific Free Storage

**Description**: Finds devices with free storage within a specified range.

**Use Cases**:
- Target devices for large app/content deployments
- Identify devices needing storage cleanup
- Storage capacity planning
- User communication for storage management

**Method**: `getDevicesWithSpecificFreeStorage(minGB, maxGB?)`

**Parameters**:
- `minGB`: Minimum free storage in GB
- `maxGB`: Maximum free storage in GB (optional)

### Report 28: Low Memory Devices

**Description**: Identifies devices with critically low free program memory.

**Use Cases**:
- Proactive performance issue detection
- User communication for device cleanup
- Device refresh planning
- Service desk prioritization

**Method**: `getDevicesWithLowMemory(thresholdMB?)`

**Parameters**:
- `thresholdMB`: Memory threshold in MB (default: 512MB)

**Alert Levels**:
- Critical: < 256 MB free
- Warning: 256-512 MB free

### Report 29: Low Storage Devices

**Description**: Identifies devices with critically low free storage.

**Use Cases**:
- Prevent app deployment failures
- Proactive user notification
- Storage cleanup campaigns
- Device health monitoring

**Method**: `getDevicesWithLowStorage(thresholdGB?)`

**Parameters**:
- `thresholdGB`: Storage threshold in GB (default: 10GB)

**Alert Levels**:
- Critical: < 5 GB free
- Warning: 5-10 GB free

## Installation

### Step 1: Azure AD App Registration

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure:
   - **Name**: `Intune Hardware Reports`
   - **Supported account types**: Single tenant
   - **Redirect URI**: Not required
5. Click **Register**

### Step 2: Note Application Details

From the app registration Overview page, record:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 3: Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `Hardware Reports Secret`
4. Expiration: 12-24 months (recommended)
5. Click **Add**
6. **Copy the secret value immediately** (it won't be shown again)

### Step 4: Configure API Permissions

1. Go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph**
4. Choose **Application permissions**
5. Add permissions:
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementConfiguration.Read.All`
6. Click **Grant admin consent for [Your Organization]**

### Step 5: Install Dependencies

```bash
# Clone the repository
git clone <repository-url>
cd intunereporting

# Install dependencies
npm install

# Build the project
npm run build
```

### Step 6: Configure Environment

Create a `.env` file:

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

## Quick Start

### Basic Usage

```typescript
import { DeviceHardwareReports } from './reports/configmgr/device-hardware-reports';
import { GraphAuth } from './auth/graph-auth';
import { AppConfig } from './types';

// Initialize
const config: AppConfig = {
  authentication: {
    tenantId: process.env.AZURE_TENANT_ID!,
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    authMethod: 'clientSecret',
  },
  // ... other config
};

const auth = new GraphAuth(config);
const graphClient = await auth.getAuthenticatedClient();
const reporter = new DeviceHardwareReports(graphClient, config);

// Get all device hardware info
const devices = await reporter.getAllDeviceHardwareInfo();
console.log(`Total devices: ${devices.length}`);

// Get OS distribution
const osDistribution = await reporter.getOSDistributionSummary();
console.log('OS Distribution:', osDistribution);

// Find low storage devices
const lowStorageDevices = await reporter.getDevicesWithLowStorage(10);
console.log(`Devices with <10GB free: ${lowStorageDevices.length}`);
```

### Generate Comprehensive Report

```typescript
// Generate complete hardware report
const report = await reporter.generateComprehensiveReport();

// Export to JSON
await reporter.exportReport(report, {
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true,
});

// Export to CSV
await reporter.exportReport(report, {
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true,
});

// Export to HTML
await reporter.exportReport(report, {
  format: 'html',
  outputDir: './reports',
  includeTimestamp: true,
});
```

### Filter by Operating System

```typescript
// Get only iOS devices
const iosDevices = await reporter.getAllDeviceHardwareInfo({
  operatingSystem: 'iOS',
});

// Get memory distribution for iOS
const iosMemory = await reporter.getMemoryRangeSummary(iosDevices);

// Get storage distribution for iOS
const iosStorage = await reporter.getStorageRangeSummary(iosDevices);
```

## API Reference

### Class: DeviceHardwareReports

#### Constructor

```typescript
constructor(graphClient: Client, config: AppConfig)
```

**Parameters**:
- `graphClient`: Authenticated Microsoft Graph client
- `config`: Application configuration object

#### Methods

##### getAllDeviceHardwareInfo()

```typescript
async getAllDeviceHardwareInfo(filter?: HardwareReportFilter): Promise<DeviceHardwareInfo[]>
```

Retrieves complete hardware information for all managed devices.

**Parameters**:
- `filter` (optional): Filter criteria for devices

**Returns**: Array of device hardware information

**Example**:
```typescript
const devices = await reporter.getAllDeviceHardwareInfo();
const iosDevices = await reporter.getAllDeviceHardwareInfo({ operatingSystem: 'iOS' });
```

##### getDisplayConfigurationSummary()

```typescript
async getDisplayConfigurationSummary(devices?: DeviceHardwareInfo[]): Promise<DisplayConfigurationSummary[]>
```

Report #9: Groups devices by display configuration categories.

**Parameters**:
- `devices` (optional): Pre-fetched device list

**Returns**: Array of display configuration summaries

##### getOSDistributionSummary()

```typescript
async getOSDistributionSummary(devices?: DeviceHardwareInfo[]): Promise<OSDistributionSummary[]>
```

Report #10: Groups devices by operating system and version.

**Parameters**:
- `devices` (optional): Pre-fetched device list

**Returns**: Array of OS distribution summaries

##### getMemoryRangeSummary()

```typescript
async getMemoryRangeSummary(devices?: DeviceHardwareInfo[]): Promise<MemoryRangeSummary[]>
```

Report #11: Categorizes devices by program memory ranges.

**Parameters**:
- `devices` (optional): Pre-fetched device list

**Returns**: Array of memory range summaries

##### getStorageRangeSummary()

```typescript
async getStorageRangeSummary(devices?: DeviceHardwareInfo[]): Promise<StorageRangeSummary[]>
```

Report #12: Categorizes devices by storage capacity ranges.

**Parameters**:
- `devices` (optional): Pre-fetched device list

**Returns**: Array of storage range summaries

##### getDevicesWithSpecificFreeMemory()

```typescript
async getDevicesWithSpecificFreeMemory(
  minFreeMemoryMB: number,
  maxFreeMemoryMB?: number,
  devices?: DeviceHardwareInfo[]
): Promise<DeviceMemoryThreshold[]>
```

Report #25: Finds devices with free memory within specified range.

**Parameters**:
- `minFreeMemoryMB`: Minimum free memory in MB
- `maxFreeMemoryMB` (optional): Maximum free memory in MB
- `devices` (optional): Pre-fetched device list

**Returns**: Array of devices matching criteria

**Examples**:
```typescript
// Devices with at least 2GB free
const devices = await reporter.getDevicesWithSpecificFreeMemory(2048);

// Devices with 2-4GB free
const devices = await reporter.getDevicesWithSpecificFreeMemory(2048, 4096);
```

##### getDevicesWithSpecificFreeStorage()

```typescript
async getDevicesWithSpecificFreeStorage(
  minFreeStorageGB: number,
  maxFreeStorageGB?: number,
  devices?: DeviceHardwareInfo[]
): Promise<DeviceStorageThreshold[]>
```

Report #26: Finds devices with free storage within specified range.

**Parameters**:
- `minFreeStorageGB`: Minimum free storage in GB
- `maxFreeStorageGB` (optional): Maximum free storage in GB
- `devices` (optional): Pre-fetched device list

**Returns**: Array of devices matching criteria

**Examples**:
```typescript
// Devices with at least 50GB free
const devices = await reporter.getDevicesWithSpecificFreeStorage(50);

// Devices with 50-200GB free
const devices = await reporter.getDevicesWithSpecificFreeStorage(50, 200);
```

##### getDevicesWithLowMemory()

```typescript
async getDevicesWithLowMemory(
  thresholdMB: number = 512,
  devices?: DeviceHardwareInfo[]
): Promise<DeviceMemoryThreshold[]>
```

Report #28: Identifies devices with low free program memory.

**Parameters**:
- `thresholdMB`: Memory threshold in MB (default: 512)
- `devices` (optional): Pre-fetched device list

**Returns**: Array of devices below threshold

**Example**:
```typescript
// Find devices with <512MB free (default)
const lowMemory = await reporter.getDevicesWithLowMemory();

// Find devices with <1GB free
const lowMemory = await reporter.getDevicesWithLowMemory(1024);
```

##### getDevicesWithLowStorage()

```typescript
async getDevicesWithLowStorage(
  thresholdGB: number = 10,
  devices?: DeviceHardwareInfo[]
): Promise<DeviceStorageThreshold[]>
```

Report #29: Identifies devices with low free storage.

**Parameters**:
- `thresholdGB`: Storage threshold in GB (default: 10)
- `devices` (optional): Pre-fetched device list

**Returns**: Array of devices below threshold (sorted by free storage ascending)

**Example**:
```typescript
// Find devices with <10GB free (default)
const lowStorage = await reporter.getDevicesWithLowStorage();

// Find devices with <20GB free
const lowStorage = await reporter.getDevicesWithLowStorage(20);
```

##### generateComprehensiveReport()

```typescript
async generateComprehensiveReport(filter?: HardwareReportFilter): Promise<ReportData>
```

Generates a comprehensive hardware report with all sections.

**Parameters**:
- `filter` (optional): Filter criteria

**Returns**: Complete report data with metadata and summary

**Includes**:
- Display configuration summary
- OS distribution
- Memory range analysis
- Storage range analysis
- Low memory/storage device counts
- Top device models by memory/storage
- Average statistics

##### exportReport()

```typescript
async exportReport(reportData: ReportData, options: ReportExportOptions): Promise<string>
```

Exports report to specified format.

**Parameters**:
- `reportData`: Report data to export
- `options`: Export configuration

**Returns**: Path to exported file

**Supported Formats**: JSON, CSV, HTML

## Microsoft Graph API Endpoints

The module uses the following Microsoft Graph API endpoints:

### Primary Endpoint

```
GET /deviceManagement/managedDevices
```

**Query Parameters**:
- `$select`: Specifies properties to retrieve
- `$filter`: OData filter for querying specific devices
- `$top`: Number of results per page (default: 999)

### Properties Retrieved

| Property | Description | Unit |
|----------|-------------|------|
| `id` | Device ID | - |
| `deviceName` | Device name | - |
| `manufacturer` | Device manufacturer | - |
| `model` | Device model | - |
| `serialNumber` | Serial number | - |
| `userPrincipalName` | User email | - |
| `enrolledDateTime` | Enrollment date | ISO 8601 |
| `lastSyncDateTime` | Last sync date | ISO 8601 |
| `operatingSystem` | OS name | - |
| `osVersion` | OS version | - |
| `totalStorageSpaceInBytes` | Total storage | Bytes |
| `freeStorageSpaceInBytes` | Free storage | Bytes |
| `physicalMemoryInBytes` | Total RAM | Bytes |
| `managementAgent` | Management type | - |
| `complianceState` | Compliance status | - |

### API Call Example

```http
GET https://graph.microsoft.com/v1.0/deviceManagement/managedDevices
  ?$select=id,deviceName,operatingSystem,osVersion,
           physicalMemoryInBytes,totalStorageSpaceInBytes,
           freeStorageSpaceInBytes,manufacturer,model
  &$top=999
```

### Pagination Handling

The module automatically handles pagination for large device inventories using the `@odata.nextLink` property.

## Data Models

### TypeScript Interfaces

#### DeviceHardwareInfo

Complete hardware information for a device.

```typescript
interface DeviceHardwareInfo {
  deviceId: string;
  deviceName: string;
  manufacturer: string;
  model: string;
  serialNumber: string;
  userPrincipalName: string;
  enrolledDateTime: string;
  lastSyncDateTime: string;
  operatingSystem: OperatingSystemInfo;
  display?: DisplayConfiguration;
  memory: MemoryConfiguration;
  storage: StorageConfiguration;
  managementAgent: string;
  complianceState: string;
}
```

#### MemoryConfiguration

Memory-related information.

```typescript
interface MemoryConfiguration {
  totalMemoryInBytes: number;
  totalMemoryInMB: number;
  totalMemoryInGB: number;
  freeMemoryInBytes?: number;
  freeMemoryInMB?: number;
  freeMemoryPercentage?: number;
}
```

#### StorageConfiguration

Storage-related information.

```typescript
interface StorageConfiguration {
  totalStorageInBytes: number;
  totalStorageInMB: number;
  totalStorageInGB: number;
  freeStorageInBytes: number;
  freeStorageInMB: number;
  freeStorageInGB: number;
  usedStorageInBytes: number;
  usedStoragePercentage: number;
  freeStoragePercentage: number;
}
```

#### OperatingSystemInfo

Operating system information.

```typescript
interface OperatingSystemInfo {
  operatingSystem: string;
  osVersion: string;
  osPlatform: string;
  buildNumber?: string;
}
```

#### HardwareReportFilter

Filter criteria for hardware reports.

```typescript
interface HardwareReportFilter {
  operatingSystem?: string;
  manufacturer?: string;
  model?: string;
  minMemoryMB?: number;
  maxMemoryMB?: number;
  minStorageGB?: number;
  maxStorageGB?: number;
  complianceState?: string;
}
```

## Export Formats

### JSON Export

Full-fidelity export with all data structures preserved.

**Use Cases**:
- Data integration with other systems
- Programmatic processing
- Archival storage

**Example**:
```typescript
await reporter.exportReport(report, {
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true,
});
```

**Output Structure**:
```json
{
  "metadata": {
    "reportName": "device-hardware-reports",
    "generatedAt": "2024-02-06T10:00:00Z",
    "recordCount": 150
  },
  "data": [...],
  "summary": {...}
}
```

### CSV Export

Tabular format suitable for spreadsheet applications.

**Use Cases**:
- Excel/Google Sheets analysis
- Database imports
- Quick data review

**Columns**:
- Device Name
- Manufacturer
- Model
- Operating System
- OS Version
- Total Memory (GB)
- Total Storage (GB)
- Free Storage (GB)
- Free Storage %
- Compliance State
- Last Sync Date

### HTML Export

Formatted report for viewing in web browsers.

**Use Cases**:
- Executive summaries
- Team sharing
- Email reports
- Web publishing

**Features**:
- Responsive design
- Sortable tables
- Visual styling
- Print-friendly

## Filtering and Thresholds

### Filter Options

#### Operating System Filter

```typescript
const devices = await reporter.getAllDeviceHardwareInfo({
  operatingSystem: 'iOS',
});
```

Supported values: `iOS`, `Android`, `Windows`, `macOS`

#### Manufacturer Filter

```typescript
const devices = await reporter.getAllDeviceHardwareInfo({
  manufacturer: 'Apple',
});
```

#### Model Filter

```typescript
const devices = await reporter.getAllDeviceHardwareInfo({
  model: 'iPhone 14',
});
```

#### Compliance State Filter

```typescript
const devices = await reporter.getAllDeviceHardwareInfo({
  complianceState: 'compliant',
});
```

Values: `compliant`, `noncompliant`, `inGracePeriod`, `unknown`

#### Combined Filters

```typescript
const devices = await reporter.getAllDeviceHardwareInfo({
  operatingSystem: 'iOS',
  manufacturer: 'Apple',
  complianceState: 'compliant',
});
```

### Custom Thresholds

#### Platform-Specific Thresholds

Different device types have different requirements:

**Mobile Devices (iOS/Android)**:
```typescript
const mobileLowMemory = await reporter.getDevicesWithLowMemory(1024); // 1GB
const mobileLowStorage = await reporter.getDevicesWithLowStorage(5); // 5GB
```

**Windows Devices**:
```typescript
const windowsLowMemory = await reporter.getDevicesWithLowMemory(2048); // 2GB
const windowsLowStorage = await reporter.getDevicesWithLowStorage(20); // 20GB
```

**Tablets/High-End Devices**:
```typescript
const tabletLowMemory = await reporter.getDevicesWithLowMemory(1536); // 1.5GB
const tabletLowStorage = await reporter.getDevicesWithLowStorage(10); // 10GB
```

#### Unit Conversions

**Memory**:
- 1 GB = 1024 MB
- 512 MB = 0.5 GB
- 2048 MB = 2 GB

**Storage**:
- 1 GB = 1024 MB
- 1 TB = 1024 GB

## Best Practices

### Performance Optimization

1. **Cache Device Data**: Fetch all devices once, then pass the cached list to subsequent methods
   ```typescript
   const devices = await reporter.getAllDeviceHardwareInfo();
   const memory = await reporter.getMemoryRangeSummary(devices);
   const storage = await reporter.getStorageRangeSummary(devices);
   ```

2. **Use Filters Early**: Apply filters at the API level when possible
   ```typescript
   const iosDevices = await reporter.getAllDeviceHardwareInfo({
     operatingSystem: 'iOS',
   });
   ```

3. **Schedule Off-Peak**: Run comprehensive reports during off-peak hours

### Data Accuracy

1. **Check Last Sync Date**: Devices that haven't synced recently may have stale data
   ```typescript
   devices.filter(d => {
     const lastSync = new Date(d.lastSyncDateTime);
     const daysSince = (Date.now() - lastSync.getTime()) / (1000 * 60 * 60 * 24);
     return daysSince <= 7; // Only devices synced in last week
   });
   ```

2. **Memory Estimation**: Graph API doesn't provide real-time free memory for mobile devices
   - The module estimates free memory at ~40% of total
   - Use for trending, not absolute values

3. **Storage Data**: Storage data is accurate and updated during device sync

### Security

1. **Secure Credentials**: Never commit credentials to source control
2. **Use Key Vault**: Store secrets in Azure Key Vault for production
3. **Rotate Secrets**: Regularly rotate client secrets
4. **Principle of Least Privilege**: Grant only required permissions

### Error Handling

1. **Retry Logic**: The module includes automatic retry with exponential backoff
2. **Validate Data**: Check for null/undefined values before processing
3. **Log Errors**: Enable logging for troubleshooting
4. **Graceful Degradation**: Handle missing device properties gracefully

### Reporting Cadence

**Daily**:
- Low memory/storage alerts
- Critical device issues

**Weekly**:
- OS distribution trends
- Memory/storage utilization

**Monthly**:
- Comprehensive hardware inventory
- Device refresh planning
- Capacity forecasting

## Troubleshooting

### Common Issues

#### No Devices Returned

**Problem**: `getAllDeviceHardwareInfo()` returns empty array

**Solutions**:
1. Verify Graph API permissions are granted
2. Check that devices are enrolled in Intune
3. Ensure app registration has admin consent
4. Verify authentication credentials

#### Authentication Failures

**Problem**: Error authenticating with Graph API

**Solutions**:
1. Verify tenant ID, client ID, and client secret
2. Check that client secret hasn't expired
3. Ensure API permissions are granted
4. Verify admin consent is completed

#### Rate Limiting

**Problem**: HTTP 429 (Too Many Requests) errors

**Solutions**:
1. Module includes automatic retry logic
2. Reduce concurrent requests
3. Implement request throttling
4. Schedule reports during off-peak hours

#### Missing Hardware Data

**Problem**: Some devices missing memory/storage data

**Solutions**:
1. Check device sync status
2. Verify device is fully enrolled
3. Some device types don't report all properties
4. Check Graph API documentation for property availability

#### Inaccurate Free Memory

**Problem**: Free memory values seem incorrect

**Explanation**:
- Graph API doesn't provide real-time free memory for most mobile devices
- Module estimates free memory at ~40% of total
- Use total memory for absolute comparisons
- Use free memory estimates for trending only

### Debugging

Enable detailed logging:

```typescript
const config: AppConfig = {
  // ...
  logging: {
    level: 'debug',
    console: true,
    file: './logs/hardware-debug.log',
  },
};
```

Check Graph API responses:

```typescript
try {
  const devices = await reporter.getAllDeviceHardwareInfo();
} catch (error) {
  console.error('Graph API Error:', error);
  // Check error.statusCode, error.message, error.body
}
```

### Getting Help

1. Check the [examples file](../src/reports/examples/device-hardware-reports-examples.ts)
2. Review [unit tests](../src/reports/configmgr/device-hardware-reports.test.ts)
3. Consult [Microsoft Graph API documentation](https://docs.microsoft.com/graph/api/intune-devices-manageddevice-list)
4. Review module source code for implementation details

## Performance Considerations

### API Call Optimization

**Device Count Impact**:
- < 100 devices: < 2 seconds
- 100-1000 devices: 5-15 seconds
- 1000-5000 devices: 30-60 seconds
- \> 5000 devices: 1-3 minutes

**Pagination**:
- Graph API returns max 1000 devices per page
- Module automatically handles pagination
- Large tenants may require multiple API calls

### Memory Usage

**Estimated Memory Requirements**:
- 100 devices: ~5 MB
- 1,000 devices: ~50 MB
- 10,000 devices: ~500 MB

**Optimization Tips**:
1. Process in batches for very large device counts
2. Export to file instead of keeping all data in memory
3. Use streaming for CSV exports

### Network Considerations

**Bandwidth**:
- ~5 KB per device on average
- 1,000 devices ≈ 5 MB transfer

**Latency**:
- Depends on network connection to Microsoft Graph
- Retry logic handles transient network issues
- Consider regional Azure endpoints for optimal performance

### Caching Strategy

```typescript
// Cache devices for multiple report types
const devices = await reporter.getAllDeviceHardwareInfo();

// Reuse cached data
const osDistribution = await reporter.getOSDistributionSummary(devices);
const memoryRanges = await reporter.getMemoryRangeSummary(devices);
const storageRanges = await reporter.getStorageRangeSummary(devices);
const lowMemory = await reporter.getDevicesWithLowMemory(512, devices);
const lowStorage = await reporter.getDevicesWithLowStorage(10, devices);

// Single API call, multiple reports generated
```

## Advanced Usage

### Scheduled Reports

Run hardware reports on a schedule using cron:

```typescript
import { CronJob } from 'cron';

// Daily low storage report at 8 AM
const job = new CronJob('0 8 * * *', async () => {
  const reporter = await initializeReporter();
  const lowStorage = await reporter.getDevicesWithLowStorage(10);

  if (lowStorage.length > 0) {
    // Send alert email
    await sendAlertEmail(lowStorage);
  }
});

job.start();
```

### Custom Metrics

Calculate custom hardware metrics:

```typescript
const devices = await reporter.getAllDeviceHardwareInfo();

// Average storage utilization
const avgUtilization = devices.reduce((sum, d) =>
  sum + d.storage.usedStoragePercentage, 0
) / devices.length;

// Devices below organization standard (e.g., < 4GB RAM)
const belowStandard = devices.filter(d =>
  d.memory.totalMemoryInGB < 4
);

// Storage capacity trend
const totalStorage = devices.reduce((sum, d) =>
  sum + d.storage.totalStorageInGB, 0
);
const avgStorage = totalStorage / devices.length;
```

### Integration with Monitoring Systems

Export metrics to monitoring platforms:

```typescript
// Prometheus-style metrics
const metrics = {
  devices_total: devices.length,
  devices_low_memory: lowMemory.length,
  devices_low_storage: lowStorage.length,
  avg_memory_gb: avgMemory,
  avg_storage_gb: avgStorage,
};

// Send to monitoring system
await sendMetrics(metrics);
```

## Examples

For comprehensive usage examples, see:
- [Usage Examples](../src/reports/examples/device-hardware-reports-examples.ts)

The examples file includes 12 detailed examples covering:
1. Basic device hardware retrieval
2. Display configuration summary
3. OS distribution analysis
4. Memory range categorization
5. Storage range categorization
6. Specific free memory filtering
7. Specific free storage filtering
8. Low memory device alerts
9. Low storage device alerts
10. Comprehensive report generation
11. Platform-filtered reports
12. Custom threshold configuration

## Changelog

### Version 1.0.0 (2024-02-06)

**Initial Release**:
- Implemented all 8 Configuration Manager hardware reports
- Full TypeScript support with comprehensive type definitions
- Export to JSON, CSV, and HTML formats
- Automatic pagination handling
- Retry logic with exponential backoff
- Filtering by OS, manufacturer, model, compliance state
- Configurable memory and storage thresholds
- Comprehensive unit tests
- Detailed documentation and examples

## License

MIT License - See LICENSE file for details

## Support

For issues, questions, or contributions:
1. Check the troubleshooting section
2. Review the examples
3. Consult the unit tests
4. Refer to Microsoft Graph API documentation

---

**Last Updated**: 2024-02-06
**Module Version**: 1.0.0
**Compatibility**: Node.js 18+, TypeScript 5+
