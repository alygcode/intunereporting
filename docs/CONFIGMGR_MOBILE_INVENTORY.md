# Configuration Manager Mobile Device Inventory Reports for Intune

Complete API reference and implementation guide for mobile device inventory reports that replicate Configuration Manager functionality using Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Configuration Manager Reports Implemented](#configuration-manager-reports-implemented)
- [Microsoft Graph API Endpoints](#microsoft-graph-api-endpoints)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [API Reference](#api-reference)
  - [Class: MobileDeviceInventoryReports](#class-mobiledeviceinventoryreports)
  - [Interfaces](#interfaces)
  - [Enums](#enums)
- [Usage Examples](#usage-examples)
- [Export Formats](#export-formats)
- [Filtering and Sorting](#filtering-and-sorting)
- [Pagination](#pagination)
- [Error Handling](#error-handling)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)

## Overview

The Mobile Device Inventory Reports module provides comprehensive reporting capabilities for Intune-managed mobile devices, replicating key Configuration Manager reports using Microsoft Graph API. This module enables organizations to:

- Track corporate-owned mobile devices
- Monitor device client communication status
- Analyze operating system distribution
- Identify inactive or problematic devices
- Generate detailed inventory reports

### Key Features

- **ConfigMgr Compatibility**: Replicates 4 key Configuration Manager mobile device reports
- **Comprehensive Data**: Collects detailed device information including hardware, OS, user, and management data
- **Communication Monitoring**: Tracks device sync status and identifies connectivity issues
- **Multiple Export Formats**: Export reports as JSON, CSV, or HTML
- **Advanced Filtering**: Filter devices by OS, ownership type, compliance state, and more
- **Pagination Support**: Handle large datasets efficiently
- **TypeScript Support**: Fully typed interfaces for all data structures

## Configuration Manager Reports Implemented

### Report 1: All Corporate-Owned Mobile Devices
**Purpose**: Displays all mobile devices that are marked as corporate/company-owned.

**Graph API Equivalent**:
```typescript
GET /deviceManagement/managedDevices?$filter=managedDeviceOwnerType eq 'company'
```

**Key Data Points**:
- Device name, manufacturer, model
- Operating system and version
- Serial number, IMEI, phone number
- User information
- Enrollment profile
- Last sync date

### Report 2: All Mobile Device Clients
**Purpose**: Displays information about all mobile device clients managed by Intune MDM, excluding devices managed only through Exchange ActiveSync connector.

**Graph API Equivalent**:
```typescript
GET /deviceManagement/managedDevices?$filter=managementAgent ne 'eas'
```

**Key Data Points**:
- All device hardware information
- Management agent type
- Ownership classification
- Compliance state
- Enrollment details

### Report 20: Mobile Device Client Information
**Purpose**: Shows mobile devices with management client installed and verifies management point communication status.

**Graph API Equivalent**:
```typescript
GET /deviceManagement/managedDevices
// Calculate communication status based on lastSyncDateTime
```

**Key Data Points**:
- Device identification
- Last contact date/time
- Days since last contact
- Communication status (Healthy/Warning/Critical)
- Active/inactive status

### Report 22: Mobile Devices by Operating System
**Purpose**: Displays mobile devices organized by operating system with version distribution.

**Graph API Equivalent**:
```typescript
GET /deviceManagement/managedDevices?$select=operatingSystem,osVersion,deviceName
// Group by OS and version
```

**Key Data Points**:
- Operating system distribution
- Version breakdown per OS
- Device count and percentages
- Latest version identification

## Microsoft Graph API Endpoints

### Primary Endpoints

#### Get All Managed Devices
```
GET /deviceManagement/managedDevices
```

#### Filter by Ownership Type
```
GET /deviceManagement/managedDevices?$filter=managedDeviceOwnerType eq 'company'
GET /deviceManagement/managedDevices?$filter=managedDeviceOwnerType eq 'personal'
```

#### Exclude Exchange Connector Devices
```
GET /deviceManagement/managedDevices?$filter=managementAgent ne 'eas'
```

#### Filter by Operating System
```
GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'iOS'
GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'Android'
GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'Windows'
```

#### Select Specific Fields
```
GET /deviceManagement/managedDevices?$select=id,deviceName,operatingSystem,osVersion,lastSyncDateTime,managementAgent
```

### Required Permissions

**Delegated (work or school account)**:
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementManagedDevices.ReadWrite.All`

**Application**:
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementManagedDevices.ReadWrite.All`

## Installation

### Prerequisites

```bash
npm install @microsoft/microsoft-graph-client
npm install @azure/identity
```

### Import the Module

```typescript
import { MobileDeviceInventoryReports } from './reports/configmgr/mobile-device-inventory-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from './types';
```

## Quick Start

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { MobileDeviceInventoryReports } from './reports/configmgr/mobile-device-inventory-reports';

// Initialize Graph client (authentication setup required)
const graphClient = Client.init({
  authProvider: (done) => {
    done(null, accessToken);
  },
});

// Create report instance
const report = new MobileDeviceInventoryReports(graphClient, config);

// Get all corporate-owned devices
const corporateDevices = await report.getAllCorporateOwnedDevices();
console.log(`Found ${corporateDevices.length} corporate devices`);

// Get OS distribution
const osDistribution = await report.getMobileDevicesByOperatingSystem();
osDistribution.forEach(os => {
  console.log(`${os.operatingSystem}: ${os.deviceCount} devices (${os.percentage}%)`);
});
```

## API Reference

### Class: MobileDeviceInventoryReports

Main class providing mobile device inventory reporting capabilities.

#### Constructor

```typescript
constructor(graphClient: Client, config: AppConfig)
```

**Parameters**:
- `graphClient`: Microsoft Graph API client instance
- `config`: Application configuration object

#### Methods

##### execute()

Executes the default mobile device inventory report.

```typescript
async execute(): Promise<ReportData>
```

**Returns**: Complete report with metadata, device data, and summary statistics.

**Example**:
```typescript
const reportData = await report.execute();
console.log(`Total devices: ${reportData.summary.totalDevices}`);
```

---

##### getAllCorporateOwnedDevices()

Retrieves all corporate-owned mobile devices (ConfigMgr Report #1).

```typescript
async getAllCorporateOwnedDevices(
  options?: MobileDeviceFilterOptions & MobileDeviceSortOptions
): Promise<CorporateMobileDevice[]>
```

**Parameters**:
- `options` (optional): Filter and sort options

**Returns**: Array of corporate mobile devices with detailed information.

**Example**:
```typescript
// Get all corporate devices
const corporateDevices = await report.getAllCorporateOwnedDevices();

// Filter by iOS devices only
const iOSCorporate = await report.getAllCorporateOwnedDevices({
  operatingSystem: 'iOS'
});

// Filter and sort
const sortedDevices = await report.getAllCorporateOwnedDevices({
  complianceState: 'compliant',
  field: 'lastSyncDateTime',
  direction: 'desc'
});
```

---

##### getAllMobileDeviceClients()

Retrieves all mobile device clients, excluding Exchange connector devices (ConfigMgr Report #2).

```typescript
async getAllMobileDeviceClients(
  options?: MobileDeviceFilterOptions & MobileDeviceSortOptions
): Promise<MobileDeviceInfo[]>
```

**Parameters**:
- `options` (optional): Filter and sort options

**Returns**: Array of mobile device information.

**Example**:
```typescript
// Get all mobile device clients
const allClients = await report.getAllMobileDeviceClients();

// Filter by Android devices
const androidClients = await report.getAllMobileDeviceClients({
  operatingSystem: 'Android'
});

// Get supervised devices only
const supervisedDevices = await report.getAllMobileDeviceClients({
  isSupervised: true
});
```

---

##### getMobileDeviceClientInformation()

Retrieves mobile device client information with communication status (ConfigMgr Report #20).

```typescript
async getMobileDeviceClientInformation(
  options?: MobileDeviceFilterOptions & MobileDeviceSortOptions
): Promise<MobileDeviceClientInfo[]>
```

**Parameters**:
- `options` (optional): Filter and sort options

**Returns**: Array of device client information with communication status.

**Communication Status Categories**:
- **Healthy**: Synced within 24 hours
- **Warning**: Synced between 1-7 days ago
- **Critical**: Not synced in 7+ days

**Example**:
```typescript
const clientInfo = await report.getMobileDeviceClientInformation();

// Filter devices with critical communication status
const criticalDevices = clientInfo.filter(
  d => d.communicationStatus === CommunicationStatus.CRITICAL
);

console.log(`${criticalDevices.length} devices need attention`);
```

---

##### getMobileDevicesByOperatingSystem()

Retrieves mobile devices grouped by operating system with version distribution (ConfigMgr Report #22).

```typescript
async getMobileDevicesByOperatingSystem(
  options?: MobileDeviceFilterOptions
): Promise<OSDistribution[]>
```

**Parameters**:
- `options` (optional): Filter options

**Returns**: Array of OS distribution data with version breakdown.

**Example**:
```typescript
const osDistribution = await report.getMobileDevicesByOperatingSystem();

osDistribution.forEach(os => {
  console.log(`\n${os.operatingSystem}: ${os.deviceCount} devices`);
  os.versions.forEach(version => {
    console.log(`  - Version ${version.version}: ${version.deviceCount} devices`);
  });
});
```

---

##### getAllMobileDevices()

Retrieves all mobile devices without filtering.

```typescript
async getAllMobileDevices(
  options?: MobileDeviceFilterOptions
): Promise<MobileDeviceInfo[]>
```

**Parameters**:
- `options` (optional): Filter options

**Returns**: Array of all mobile devices.

**Example**:
```typescript
const allDevices = await report.getAllMobileDevices();
```

---

##### getDevicesByOwnerType()

Retrieves devices filtered by ownership type.

```typescript
async getDevicesByOwnerType(
  ownerType: DeviceOwnershipType
): Promise<MobileDeviceInfo[]>
```

**Parameters**:
- `ownerType`: Device ownership type (company, personal, unknown)

**Returns**: Array of devices with specified ownership type.

**Example**:
```typescript
const corporateDevices = await report.getDevicesByOwnerType(
  DeviceOwnershipType.COMPANY
);
const personalDevices = await report.getDevicesByOwnerType(
  DeviceOwnershipType.PERSONAL
);
```

---

##### getDevicesByOS()

Retrieves devices filtered by operating system.

```typescript
async getDevicesByOS(operatingSystem: string): Promise<MobileDeviceInfo[]>
```

**Parameters**:
- `operatingSystem`: Operating system name (iOS, Android, Windows, etc.)

**Returns**: Array of devices with specified OS.

**Example**:
```typescript
const iosDevices = await report.getDevicesByOS('iOS');
const androidDevices = await report.getDevicesByOS('Android');
```

---

##### getDevicesByManagementAgent()

Retrieves devices filtered by management agent type.

```typescript
async getDevicesByManagementAgent(
  managementAgent: ManagementAgentType
): Promise<MobileDeviceInfo[]>
```

**Parameters**:
- `managementAgent`: Management agent type

**Returns**: Array of devices with specified management agent.

**Example**:
```typescript
const mdmDevices = await report.getDevicesByManagementAgent(
  ManagementAgentType.INTUNE_MDM
);
```

---

##### getInactiveDevices()

Identifies devices that haven't synced within specified days.

```typescript
async getInactiveDevices(
  devices: MobileDeviceInfo[],
  daysThreshold: number = 30
): Promise<MobileDeviceInfo[]>
```

**Parameters**:
- `devices`: Array of devices to analyze
- `daysThreshold`: Number of days without sync to consider inactive (default: 30)

**Returns**: Array of inactive devices.

**Example**:
```typescript
const allDevices = await report.getAllMobileDevices();
const inactiveDevices = await report.getInactiveDevices(allDevices, 30);
console.log(`${inactiveDevices.length} devices inactive for 30+ days`);
```

---

##### getActiveDevices()

Identifies devices that have synced within specified days.

```typescript
getActiveDevices(
  devices: MobileDeviceInfo[],
  daysThreshold: number = 7
): MobileDeviceInfo[]
```

**Parameters**:
- `devices`: Array of devices to analyze
- `daysThreshold`: Number of days to consider active (default: 7)

**Returns**: Array of active devices.

**Example**:
```typescript
const allDevices = await report.getAllMobileDevices();
const activeDevices = report.getActiveDevices(allDevices, 7);
console.log(`${activeDevices.length} devices active in last 7 days`);
```

---

##### paginate()

Applies pagination to device results.

```typescript
paginate<T>(data: T[], options: PaginationOptions): PaginatedResult<T>
```

**Parameters**:
- `data`: Array of data to paginate
- `options`: Pagination options (page number, page size)

**Returns**: Paginated result with metadata.

**Example**:
```typescript
const allDevices = await report.getAllMobileDevices();
const page1 = report.paginate(allDevices, { page: 1, pageSize: 50 });

console.log(`Page ${page1.pagination.currentPage} of ${page1.pagination.totalPages}`);
console.log(`Showing ${page1.data.length} devices`);
```

---

##### exportReport()

Exports report data to file in specified format.

```typescript
async exportReport(
  reportData: ReportData,
  options: ExportOptions
): Promise<string>
```

**Parameters**:
- `reportData`: Report data to export
- `options`: Export options (format, directory, timestamp)

**Returns**: Path to exported file.

**Example**:
```typescript
const reportData = await report.execute();

// Export as JSON
const jsonPath = await report.exportReport(reportData, {
  format: 'json',
  outputDirectory: './reports',
  includeTimestamp: true
});

// Export as CSV
const csvPath = await report.exportReport(reportData, {
  format: 'csv',
  outputDirectory: './reports',
  includeTimestamp: true
});

// Export as HTML
const htmlPath = await report.exportReport(reportData, {
  format: 'html',
  outputDirectory: './reports',
  includeTimestamp: true
});
```

---

### Convenience Functions

Quick helper functions for common report scenarios.

##### generateCorporateDevicesReport()

Generates a complete corporate devices report.

```typescript
async function generateCorporateDevicesReport(
  graphClient: Client,
  config: AppConfig
): Promise<ReportData>
```

**Example**:
```typescript
import { generateCorporateDevicesReport } from './mobile-device-inventory-reports';

const report = await generateCorporateDevicesReport(graphClient, config);
console.log(`Total: ${report.summary.totalCorporateDevices}`);
```

---

##### generateOSDistributionReport()

Generates a complete OS distribution report.

```typescript
async function generateOSDistributionReport(
  graphClient: Client,
  config: AppConfig
): Promise<ReportData>
```

**Example**:
```typescript
import { generateOSDistributionReport } from './mobile-device-inventory-reports';

const report = await generateOSDistributionReport(graphClient, config);
```

---

##### generateDeviceClientCommunicationReport()

Generates a complete device client communication report.

```typescript
async function generateDeviceClientCommunicationReport(
  graphClient: Client,
  config: AppConfig
): Promise<ReportData>
```

**Example**:
```typescript
import { generateDeviceClientCommunicationReport } from './mobile-device-inventory-reports';

const report = await generateDeviceClientCommunicationReport(graphClient, config);
console.log(`Healthy: ${report.summary.healthyDevices}`);
console.log(`Critical: ${report.summary.criticalDevices}`);
```

---

### Interfaces

#### MobileDeviceInfo

Basic mobile device information.

```typescript
interface MobileDeviceInfo {
  deviceId: string;
  deviceName: string;
  operatingSystem: string;
  osVersion: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  imei?: string;
  meid?: string;
  phoneNumber?: string;
  ownerType: DeviceOwnershipType;
  managementAgent: ManagementAgentType;
  enrolledDateTime?: Date;
  lastSyncDateTime?: Date;
  complianceState?: string;
  userPrincipalName?: string;
  userDisplayName?: string;
  emailAddress?: string;
  azureADDeviceId?: string;
  managedDeviceOwnerType?: string;
  deviceEnrollmentType?: string;
  isSupervised?: boolean;
  isEncrypted?: boolean;
  jailBroken?: string;
}
```

---

#### CorporateMobileDevice

Corporate-owned mobile device with additional details.

```typescript
interface CorporateMobileDevice extends MobileDeviceInfo {
  purchaseDate?: Date;
  warrantyExpirationDate?: Date;
  assetTag?: string;
  departmentName?: string;
  costCenter?: string;
  enrollmentProfileName?: string;
}
```

---

#### MobileDeviceClientInfo

Mobile device client information with communication status.

```typescript
interface MobileDeviceClientInfo extends MobileDeviceInfo {
  communicationStatus: CommunicationStatus;
  lastContactDateTime?: Date;
  daysSinceLastContact?: number;
  isActive: boolean;
  managementPointUrl?: string;
  clientVersion?: string;
  deviceRegistrationState?: string;
  exchangeAccessState?: string;
  exchangeAccessStateReason?: string;
}
```

---

#### OSDistribution

Operating system distribution data.

```typescript
interface OSDistribution {
  operatingSystem: string;
  deviceCount: number;
  percentage: number;
  versions: OSVersionDistribution[];
}
```

---

#### OSVersionDistribution

OS version distribution data.

```typescript
interface OSVersionDistribution {
  version: string;
  deviceCount: number;
  percentage: number;
  isLatest?: boolean;
  isSupportedVersion?: boolean;
}
```

---

#### MobileDeviceFilterOptions

Filter options for device queries.

```typescript
interface MobileDeviceFilterOptions {
  ownerType?: DeviceOwnershipType;
  operatingSystem?: string;
  managementAgent?: ManagementAgentType;
  complianceState?: string;
  minLastSyncDays?: number;
  maxLastSyncDays?: number;
  isSupervised?: boolean;
  includeInactive?: boolean;
}
```

---

#### MobileDeviceSortOptions

Sort options for device results.

```typescript
interface MobileDeviceSortOptions {
  field: 'deviceName' | 'operatingSystem' | 'lastSyncDateTime' | 'enrolledDateTime' | 'userPrincipalName';
  direction: 'asc' | 'desc';
}
```

---

#### PaginationOptions

Pagination configuration.

```typescript
interface PaginationOptions {
  page: number;
  pageSize: number;
}
```

---

#### PaginatedResult

Paginated result with metadata.

```typescript
interface PaginatedResult<T> {
  data: T[];
  pagination: {
    currentPage: number;
    pageSize: number;
    totalPages: number;
    totalRecords: number;
    hasNextPage: boolean;
    hasPreviousPage: boolean;
  };
}
```

---

#### ExportOptions

Export configuration options.

```typescript
interface ExportOptions {
  format: 'json' | 'csv' | 'html';
  outputDirectory: string;
  includeTimestamp?: boolean;
}
```

---

### Enums

#### DeviceOwnershipType

Device ownership classification.

```typescript
enum DeviceOwnershipType {
  UNKNOWN = 'unknown',
  COMPANY = 'company',
  PERSONAL = 'personal',
}
```

---

#### ManagementAgentType

Device management agent type.

```typescript
enum ManagementAgentType {
  MDM = 'mdm',
  EAS = 'eas',
  INTUNE_MDM_AND_EAS = 'intuneClientAndMdmAndEas',
  INTUNE_MDM = 'intuneMdm',
  EASMDM = 'easMdm',
  CONFIGMGR = 'configurationManagerClient',
  CONFIGMGR_MDM = 'configurationManagerClientMdm',
  CONFIGMGR_MDM_EAS = 'configurationManagerClientMdmEas',
  UNKNOWN = 'unknown',
  JAMF = 'jamf',
  GOOGLE_CLOUD = 'googleCloudDevicePolicyController',
}
```

---

#### CommunicationStatus

Device communication health status.

```typescript
enum CommunicationStatus {
  HEALTHY = 'healthy',    // Synced within 24 hours
  WARNING = 'warning',    // Synced 1-7 days ago
  CRITICAL = 'critical',  // Not synced in 7+ days
  UNKNOWN = 'unknown',    // Never synced or no data
}
```

---

## Usage Examples

### Example 1: Corporate Device Inventory

Generate a complete inventory of corporate-owned devices.

```typescript
const report = new MobileDeviceInventoryReports(graphClient, config);

// Get all corporate devices
const corporateDevices = await report.getAllCorporateOwnedDevices();

// Group by OS
const byOS = new Map<string, number>();
corporateDevices.forEach(device => {
  byOS.set(device.operatingSystem, (byOS.get(device.operatingSystem) || 0) + 1);
});

console.log('Corporate Devices by OS:');
byOS.forEach((count, os) => {
  console.log(`  ${os}: ${count}`);
});

// Export to CSV
const reportData = {
  metadata: {
    reportName: 'corporate-devices',
    generatedAt: new Date().toISOString(),
    generatedBy: 'Intune Reporting',
    recordCount: corporateDevices.length,
  },
  data: corporateDevices,
  summary: { total: corporateDevices.length },
};

await report.exportReport(reportData, {
  format: 'csv',
  outputDirectory: './reports',
  includeTimestamp: true,
});
```

---

### Example 2: Identify Problematic Devices

Find devices with communication issues or that are inactive.

```typescript
const report = new MobileDeviceInventoryReports(graphClient, config);

// Get device client information with communication status
const clientInfo = await report.getMobileDeviceClientInformation();

// Filter by communication status
const criticalDevices = clientInfo.filter(
  d => d.communicationStatus === CommunicationStatus.CRITICAL
);

const warningDevices = clientInfo.filter(
  d => d.communicationStatus === CommunicationStatus.WARNING
);

console.log(`Critical devices (7+ days): ${criticalDevices.length}`);
console.log(`Warning devices (1-7 days): ${warningDevices.length}`);

// Show devices needing immediate attention
criticalDevices.forEach(device => {
  console.log(`\n${device.deviceName}`);
  console.log(`  User: ${device.userPrincipalName}`);
  console.log(`  Last Sync: ${device.lastContactDateTime?.toLocaleString()}`);
  console.log(`  Days Since Contact: ${device.daysSinceLastContact}`);
});
```

---

### Example 3: OS Distribution Analysis

Analyze operating system distribution across mobile devices.

```typescript
const report = new MobileDeviceInventoryReports(graphClient, config);

// Get OS distribution
const osDistribution = await report.getMobileDevicesByOperatingSystem();

console.log('Mobile Device OS Distribution:\n');

osDistribution.forEach(os => {
  console.log(`${os.operatingSystem}: ${os.deviceCount} devices (${os.percentage.toFixed(2)}%)`);

  // Show version breakdown
  console.log('  Versions:');
  os.versions.forEach(version => {
    console.log(`    ${version.version}: ${version.deviceCount} (${version.percentage.toFixed(2)}%)`);
  });
  console.log();
});

// Export to HTML for visualization
const reportData = {
  metadata: {
    reportName: 'os-distribution',
    generatedAt: new Date().toISOString(),
    generatedBy: 'Intune Reporting',
    recordCount: osDistribution.length,
  },
  data: osDistribution,
  summary: {
    totalOS: osDistribution.length,
    totalDevices: osDistribution.reduce((sum, os) => sum + os.deviceCount, 0),
  },
};

await report.exportReport(reportData, {
  format: 'html',
  outputDirectory: './reports',
  includeTimestamp: true,
});
```

---

### Example 4: Advanced Filtering

Use advanced filtering to find specific devices.

```typescript
const report = new MobileDeviceInventoryReports(graphClient, config);

// Get compliant iOS corporate devices, sorted by last sync
const devices = await report.getAllCorporateOwnedDevices({
  operatingSystem: 'iOS',
  complianceState: 'compliant',
  field: 'lastSyncDateTime',
  direction: 'desc',
});

console.log(`Found ${devices.length} compliant iOS corporate devices\n`);

// Show most recently synced devices
devices.slice(0, 10).forEach((device, index) => {
  console.log(`${index + 1}. ${device.deviceName}`);
  console.log(`   Model: ${device.model}`);
  console.log(`   OS: ${device.osVersion}`);
  console.log(`   Last Sync: ${device.lastSyncDateTime?.toLocaleString()}`);
  console.log();
});
```

---

### Example 5: Pagination for Large Datasets

Handle large numbers of devices using pagination.

```typescript
const report = new MobileDeviceInventoryReports(graphClient, config);

// Get all devices
const allDevices = await report.getAllMobileDevices();

// Paginate results
const pageSize = 50;
const totalPages = Math.ceil(allDevices.length / pageSize);

console.log(`Total devices: ${allDevices.length}`);
console.log(`Pages: ${totalPages} (${pageSize} per page)\n`);

// Process each page
for (let pageNum = 1; pageNum <= totalPages; pageNum++) {
  const page = report.paginate(allDevices, {
    page: pageNum,
    pageSize,
  });

  console.log(`\nPage ${page.pagination.currentPage}:`);
  console.log(`  Devices: ${page.data.length}`);
  console.log(`  Has Next: ${page.pagination.hasNextPage}`);

  // Process devices on this page
  page.data.forEach(device => {
    // Process device...
  });
}
```

---

### Example 6: Complete Workflow

Complete workflow from data collection to export.

```typescript
const report = new MobileDeviceInventoryReports(graphClient, config);

// Step 1: Collect data
console.log('Collecting device data...');
const allDevices = await report.getAllMobileDevices();
const corporateDevices = await report.getAllCorporateOwnedDevices();
const clientInfo = await report.getMobileDeviceClientInformation();
const osDistribution = await report.getMobileDevicesByOperatingSystem();

// Step 2: Analyze
const activeDevices = report.getActiveDevices(allDevices, 7);
const inactiveDevices = await report.getInactiveDevices(allDevices, 30);
const healthyDevices = clientInfo.filter(
  d => d.communicationStatus === CommunicationStatus.HEALTHY
);

// Step 3: Generate summary
const summary = {
  totalDevices: allDevices.length,
  corporateDevices: corporateDevices.length,
  personalDevices: allDevices.length - corporateDevices.length,
  activeDevices: activeDevices.length,
  inactiveDevices: inactiveDevices.length,
  healthyDevices: healthyDevices.length,
  operatingSystems: osDistribution.map(os => ({
    name: os.operatingSystem,
    count: os.deviceCount,
  })),
};

// Step 4: Export reports
const reportData = {
  metadata: {
    reportName: 'mobile-device-inventory-complete',
    generatedAt: new Date().toISOString(),
    generatedBy: 'Intune Reporting',
    recordCount: allDevices.length,
  },
  data: allDevices,
  summary,
};

// Export in multiple formats
await report.exportReport(reportData, {
  format: 'json',
  outputDirectory: './reports',
  includeTimestamp: true,
});

await report.exportReport(reportData, {
  format: 'csv',
  outputDirectory: './reports',
  includeTimestamp: true,
});

await report.exportReport(reportData, {
  format: 'html',
  outputDirectory: './reports',
  includeTimestamp: true,
});

console.log('Reports exported successfully!');
```

---

## Export Formats

### JSON Export

Exports complete report data as JSON with full metadata and nested structures.

**Features**:
- Preserves all data types
- Maintains nested objects and arrays
- Includes full metadata
- Pretty-printed for readability

**Example**:
```json
{
  "metadata": {
    "reportName": "mobile-device-inventory",
    "generatedAt": "2024-02-06T10:00:00Z",
    "generatedBy": "Intune Reporting Dashboard",
    "recordCount": 150
  },
  "data": [
    {
      "deviceId": "device-001",
      "deviceName": "iPhone 14 Pro",
      "operatingSystem": "iOS",
      "osVersion": "17.2.1"
    }
  ],
  "summary": {
    "totalDevices": 150,
    "corporateDevices": 100
  }
}
```

---

### CSV Export

Exports data as comma-separated values, flattened for spreadsheet analysis.

**Features**:
- Flattened data structure
- Metadata as comments
- Compatible with Excel, Google Sheets
- Easy to import into databases

**Example**:
```csv
# Report: mobile-device-inventory
# Generated: 2024-02-06T10:00:00Z
# Records: 150

deviceId,deviceName,operatingSystem,osVersion,ownerType,lastSyncDateTime
device-001,iPhone 14 Pro,iOS,17.2.1,company,2024-02-06T08:00:00Z
device-002,Galaxy S23,Android,14.0,company,2024-02-06T07:30:00Z
```

---

### HTML Export

Exports data as styled HTML with interactive tables and charts.

**Features**:
- Professional styling
- Responsive design
- Summary cards
- Sortable tables
- Print-friendly

**Example Output**:
- Metadata section with report details
- Summary statistics in cards
- Full data table with all records
- Automatic formatting of dates and values

---

## Filtering and Sorting

### Available Filters

| Filter | Type | Description | Example |
|--------|------|-------------|---------|
| `ownerType` | `DeviceOwnershipType` | Filter by ownership | `company`, `personal` |
| `operatingSystem` | `string` | Filter by OS | `iOS`, `Android`, `Windows` |
| `managementAgent` | `ManagementAgentType` | Filter by agent | `mdm`, `eas` |
| `complianceState` | `string` | Filter by compliance | `compliant`, `noncompliant` |
| `isSupervised` | `boolean` | Filter supervised devices | `true`, `false` |
| `minLastSyncDays` | `number` | Minimum days since sync | `7` |
| `maxLastSyncDays` | `number` | Maximum days since sync | `30` |

### Sort Options

| Field | Description |
|-------|-------------|
| `deviceName` | Sort by device name |
| `operatingSystem` | Sort by OS |
| `lastSyncDateTime` | Sort by last sync date |
| `enrolledDateTime` | Sort by enrollment date |
| `userPrincipalName` | Sort by user |

**Directions**: `asc` (ascending), `desc` (descending)

### Filter Examples

```typescript
// iOS devices only
const iosDevices = await report.getAllMobileDeviceClients({
  operatingSystem: 'iOS'
});

// Compliant corporate devices
const compliantCorporate = await report.getAllCorporateOwnedDevices({
  complianceState: 'compliant'
});

// Supervised devices
const supervised = await report.getAllMobileDeviceClients({
  isSupervised: true
});

// Combined filters with sorting
const devices = await report.getAllMobileDeviceClients({
  operatingSystem: 'Android',
  ownerType: DeviceOwnershipType.COMPANY,
  complianceState: 'compliant',
  field: 'lastSyncDateTime',
  direction: 'desc'
});
```

---

## Pagination

### Basic Pagination

```typescript
const allDevices = await report.getAllMobileDevices();

// Get first page (50 devices)
const page1 = report.paginate(allDevices, {
  page: 1,
  pageSize: 50
});

console.log(`Page ${page1.pagination.currentPage} of ${page1.pagination.totalPages}`);
console.log(`Showing ${page1.data.length} devices`);
```

### Pagination Properties

```typescript
interface PaginationMetadata {
  currentPage: number;      // Current page number
  pageSize: number;         // Items per page
  totalPages: number;       // Total number of pages
  totalRecords: number;     // Total number of records
  hasNextPage: boolean;     // True if more pages available
  hasPreviousPage: boolean; // True if previous pages exist
}
```

### Iterating Through Pages

```typescript
const allDevices = await report.getAllMobileDevices();
const pageSize = 100;

let currentPage = 1;
let hasMore = true;

while (hasMore) {
  const page = report.paginate(allDevices, {
    page: currentPage,
    pageSize
  });

  // Process devices on this page
  page.data.forEach(device => {
    console.log(device.deviceName);
  });

  hasMore = page.pagination.hasNextPage;
  currentPage++;
}
```

---

## Error Handling

### Common Errors

#### Authentication Errors

```typescript
try {
  const devices = await report.getAllMobileDevices();
} catch (error) {
  if (error.statusCode === 401) {
    console.error('Authentication failed. Check credentials.');
  } else if (error.statusCode === 403) {
    console.error('Insufficient permissions. Check Graph API permissions.');
  }
}
```

#### Rate Limiting

```typescript
try {
  const devices = await report.getAllMobileDevices();
} catch (error) {
  if (error.statusCode === 429) {
    console.error('Rate limit exceeded. Retry after delay.');
    // Implement exponential backoff
  }
}
```

#### Network Errors

```typescript
try {
  const devices = await report.getAllMobileDevices();
} catch (error) {
  if (error.code === 'ETIMEDOUT') {
    console.error('Request timed out. Check network connection.');
  }
}
```

### Best Practices

1. **Always use try-catch blocks**
   ```typescript
   try {
     const devices = await report.getAllMobileDevices();
   } catch (error) {
     console.error('Error fetching devices:', error);
     // Handle error appropriately
   }
   ```

2. **Check for empty results**
   ```typescript
   const devices = await report.getAllMobileDevices();
   if (devices.length === 0) {
     console.log('No devices found');
     return;
   }
   ```

3. **Validate data before processing**
   ```typescript
   devices.forEach(device => {
     if (device.lastSyncDateTime) {
       // Process sync date
     } else {
       console.warn(`Device ${device.deviceName} has no sync date`);
     }
   });
   ```

---

## Best Practices

### Performance Optimization

1. **Use filtering at the API level**
   ```typescript
   // Good - filter at API
   const iosDevices = await report.getAllMobileDevices({
     operatingSystem: 'iOS'
   });

   // Bad - fetch all then filter
   const allDevices = await report.getAllMobileDevices();
   const iosDevices = allDevices.filter(d => d.operatingSystem === 'iOS');
   ```

2. **Use pagination for large datasets**
   ```typescript
   const pageSize = 100;
   const page1 = report.paginate(allDevices, { page: 1, pageSize });
   ```

3. **Cache frequently accessed data**
   ```typescript
   // Cache devices list
   let cachedDevices: MobileDeviceInfo[] | null = null;
   let cacheTime: number = 0;
   const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

   async function getDevices() {
     const now = Date.now();
     if (cachedDevices && (now - cacheTime) < CACHE_TTL) {
       return cachedDevices;
     }

     cachedDevices = await report.getAllMobileDevices();
     cacheTime = now;
     return cachedDevices;
   }
   ```

### Security Best Practices

1. **Never log sensitive data**
   ```typescript
   // Bad
   console.log('Device:', device);

   // Good
   console.log('Device ID:', device.deviceId);
   ```

2. **Use least privilege permissions**
   - Request only `DeviceManagementManagedDevices.Read.All`
   - Avoid `ReadWrite` permissions unless necessary

3. **Implement proper authentication**
   ```typescript
   // Use secure token storage
   // Implement token refresh
   // Use MSAL for authentication
   ```

### Reporting Best Practices

1. **Include meaningful metadata**
   ```typescript
   const reportData = {
     metadata: {
       reportName: 'device-inventory',
       generatedAt: new Date().toISOString(),
       generatedBy: 'Automated System',
       recordCount: devices.length,
       parameters: {
         ownerType: 'company',
         dateRange: '2024-01-01 to 2024-02-01'
       }
     },
     data: devices,
     summary: {...}
   };
   ```

2. **Generate multiple export formats**
   ```typescript
   // Export JSON for programmatic access
   await report.exportReport(reportData, { format: 'json', ... });

   // Export CSV for spreadsheet analysis
   await report.exportReport(reportData, { format: 'csv', ... });

   // Export HTML for viewing
   await report.exportReport(reportData, { format: 'html', ... });
   ```

3. **Schedule regular reports**
   ```typescript
   // Run daily at 6 AM
   cron.schedule('0 6 * * *', async () => {
     const reportData = await report.execute();
     await report.exportReport(reportData, {...});
   });
   ```

---

## Troubleshooting

### Issue: No Devices Returned

**Possible Causes**:
- Insufficient permissions
- No devices enrolled in Intune
- Incorrect filter parameters

**Solution**:
```typescript
// Check permissions
console.log('Checking permissions...');

// Try without filters
const allDevices = await report.getAllMobileDevices();
console.log(`Total devices: ${allDevices.length}`);

// Check specific filter
const iosDevices = await report.getAllMobileDevices({
  operatingSystem: 'iOS'
});
console.log(`iOS devices: ${iosDevices.length}`);
```

---

### Issue: Rate Limiting Errors

**Symptoms**: Error 429 from Graph API

**Solution**:
```typescript
// The module includes automatic retry with exponential backoff
// Reduce concurrent requests if needed

// Process in batches
const batchSize = 10;
for (let i = 0; i < deviceIds.length; i += batchSize) {
  const batch = deviceIds.slice(i, i + batchSize);
  await Promise.all(batch.map(id => processDevice(id)));
  await new Promise(resolve => setTimeout(resolve, 1000)); // Wait 1 second
}
```

---

### Issue: Missing Device Data

**Symptoms**: Undefined or null values for device properties

**Solution**:
```typescript
// Always check for undefined/null
const deviceName = device.deviceName || 'Unknown';
const manufacturer = device.manufacturer || 'Unknown';

// Use optional chaining
const syncDate = device.lastSyncDateTime?.toLocaleString() || 'Never';

// Validate before processing
if (device.lastSyncDateTime) {
  const daysSince = calculateDaysSince(device.lastSyncDateTime);
}
```

---

### Issue: Export Fails

**Symptoms**: Error when exporting reports

**Solution**:
```typescript
// Ensure output directory exists
const fs = require('fs');
const outputDir = './reports';
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// Check for write permissions
try {
  await report.exportReport(reportData, {
    format: 'json',
    outputDirectory: outputDir,
    includeTimestamp: true
  });
} catch (error) {
  console.error('Export failed:', error.message);
}
```

---

## Support and Contributing

### Getting Help

- Review the [examples file](../src/reports/examples/mobile-device-inventory-examples.ts)
- Check the [test file](../src/reports/configmgr/mobile-device-inventory-reports.test.ts) for usage patterns
- Consult [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/api/intune-devices-manageddevice-list)

### Reporting Issues

When reporting issues, include:
1. Code snippet showing the issue
2. Error messages and stack traces
3. Environment details (Node.js version, dependencies)
4. Expected vs actual behavior

### Contributing

Contributions are welcome! Please:
1. Follow the existing code style
2. Add tests for new features
3. Update documentation
4. Submit pull requests with clear descriptions

---

## License

This module is part of the Intune Reporting Dashboard project.

---

## Changelog

### Version 1.0.0
- Initial release
- Implemented 4 ConfigMgr-style reports
- Added support for JSON, CSV, and HTML exports
- Comprehensive filtering and sorting
- Pagination support
- Full TypeScript interfaces
- Complete test coverage
- Usage examples

---

## Additional Resources

- [Microsoft Graph API - Managed Devices](https://docs.microsoft.com/en-us/graph/api/intune-devices-manageddevice-list)
- [Intune Device Management](https://docs.microsoft.com/en-us/mem/intune/remote-actions/)
- [Configuration Manager Reports](https://docs.microsoft.com/en-us/mem/configmgr/core/servers/manage/introduction-to-reporting)
- [Graph API Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference)

---

**Last Updated**: February 6, 2024
