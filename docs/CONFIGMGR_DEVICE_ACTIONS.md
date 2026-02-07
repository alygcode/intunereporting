# Configuration Manager Device Action Reports for Intune

This guide provides comprehensive documentation for the Device Action Reports module, which replicates Configuration Manager reports for Intune device lifecycle actions using Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Configuration Manager Reports Mapping](#configuration-manager-reports-mapping)
- [API Reference](#api-reference)
- [Usage Examples](#usage-examples)
- [Microsoft Graph API Endpoints](#microsoft-graph-api-endpoints)
- [Data Models](#data-models)
- [Export Formats](#export-formats)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Performance Considerations](#performance-considerations)

## Overview

The Device Action Reports module provides Intune equivalents for Configuration Manager's device action and lifecycle reports. It tracks device enrollment, pending actions, and wipe operations to help administrators monitor device management lifecycle.

### Implemented Reports

This module implements the following Configuration Manager reports for Intune:

1. **Report #31**: Pending retire and wipe request for mobile devices
2. **Report #32**: Recently enrolled and assigned mobile devices
3. **Report #33**: Recently wiped mobile devices

### Key Features

- **Complete Lifecycle Tracking**: Monitor devices from enrollment through retirement
- **Action History**: Track pending and completed device management actions
- **Assignment Verification**: Check successful site/group assignment for new devices
- **Wipe Analytics**: Monitor wipe success rates and duration
- **Multiple Export Formats**: JSON, CSV, and HTML output
- **Type-Safe**: Full TypeScript support with comprehensive interfaces
- **Pagination Support**: Handles large device inventories
- **Error Handling**: Robust error handling with detailed logging

## Prerequisites

### Software Requirements

- Node.js 18.0.0 or higher
- TypeScript 5.0 or higher
- npm or yarn package manager

### Azure Requirements

- Azure AD (Entra ID) tenant
- Intune Administrator or Global Administrator access
- Microsoft Intune subscription with enrolled devices

### Required API Permissions

Your Azure AD app registration needs the following Microsoft Graph permissions:

**Application Permissions (Required):**
- `DeviceManagementManagedDevices.Read.All` - Read managed device properties
- `DeviceManagementConfiguration.Read.All` - Read device configuration and actions

**Application Permissions (Optional, for extended features):**
- `User.Read.All` - Read user profiles for assignment details
- `Group.Read.All` - Read group information for assignment details

## Installation

### 1. Install Dependencies

```bash
npm install @microsoft/microsoft-graph-client @azure/identity json2csv
npm install --save-dev @types/node typescript
```

### 2. Configure Environment Variables

Create a `.env` file in your project root:

```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
```

### 3. Import the Module

```typescript
import {
  DeviceActionReports,
  createDeviceActionReports,
  generateQuickReport,
} from './src/reports/configmgr/device-action-reports';
```

## Quick Start

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { DeviceActionReports } from './device-action-reports';

// Create authenticated Graph client
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID!,
  process.env.AZURE_CLIENT_ID!,
  process.env.AZURE_CLIENT_SECRET!
);

const graphClient = Client.initWithMiddleware({ authProvider });

// Create reports instance
const reports = new DeviceActionReports(graphClient);

// Generate comprehensive report
const report = await reports.generateDeviceActionReport();

console.log(`Pending Actions: ${report.summary.totalPendingActions}`);
console.log(`Recent Enrollments: ${report.summary.totalRecentEnrollments}`);
console.log(`Recent Wipes: ${report.summary.totalRecentWipes}`);
```

### Quick Report Helper

```typescript
import { generateQuickReport } from './device-action-reports';

// Generate 7-day report
const report = await generateQuickReport(graphClient, 7);

// Export to HTML
const reports = createDeviceActionReports(graphClient);
const html = reports.exportToHTML(report);
await fs.writeFile('report.html', html);
```

## Configuration Manager Reports Mapping

### Report #31: Pending Retire and Wipe Requests

**Configuration Manager**: Shows all mobile devices with pending wipe or retire requests.

**Intune Implementation**:
```typescript
const pendingActions = await reports.getPendingRetireWipeRequests();
```

**Data Returned**:
- Device information (name, OS, version)
- User details
- Action type (retire, wipe, factory reset)
- Action status (pending, in progress, completed, failed)
- Requested date/time
- Ownership type

**Use Cases**:
- Monitor pending device cleanup actions
- Identify stalled retire/wipe operations
- Track device retirement timeline

### Report #32: Recently Enrolled and Assigned Mobile Devices

**Configuration Manager**: Lists newly enrolled devices with assignment status.

**Intune Implementation**:
```typescript
const enrolledDevices = await reports.getRecentlyEnrolledDevices({
  dateRange: {
    startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});
```

**Data Returned**:
- Device details (name, model, serial number)
- Enrollment date and age
- User information
- Compliance state
- Assignment status (assigned/not assigned)
- Assigned users and groups
- Management agent information

**Use Cases**:
- Verify new device enrollments
- Ensure proper policy assignment
- Monitor enrollment trends
- Identify unassigned devices

### Report #33: Recently Wiped Mobile Devices

**Configuration Manager**: Shows completed wipe operations for mobile devices.

**Intune Implementation**:
```typescript
const wipedDevices = await reports.getRecentlyWipedDevices({
  dateRange: {
    startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});
```

**Data Returned**:
- Device information
- Wipe type (wipe, retire, factory reset)
- Wipe status (success, failed, partial)
- Request and completion timestamps
- Wipe duration
- Error details (if failed)
- Initiated by user

**Use Cases**:
- Audit device wipe operations
- Monitor wipe success rates
- Identify failed wipe operations
- Track wipe performance metrics

## API Reference

### DeviceActionReports Class

#### Constructor

```typescript
constructor(graphClient: Client)
```

Creates a new instance of DeviceActionReports.

**Parameters**:
- `graphClient`: Authenticated Microsoft Graph client

#### Methods

##### getPendingRetireWipeRequests()

```typescript
async getPendingRetireWipeRequests(
  options?: DeviceActionReportOptions
): Promise<PendingDeviceAction[]>
```

Retrieves all pending retire and wipe requests.

**Parameters**:
- `options`: Optional filtering options

**Returns**: Array of pending device actions

**Example**:
```typescript
const pending = await reports.getPendingRetireWipeRequests({
  actionType: ['retire', 'wipe'],
  status: ['pending'],
  top: 100,
});
```

##### getRecentlyEnrolledDevices()

```typescript
async getRecentlyEnrolledDevices(
  options?: DeviceActionReportOptions
): Promise<RecentlyEnrolledDevice[]>
```

Retrieves recently enrolled devices (default: last 7 days).

**Parameters**:
- `options`: Optional filtering and date range options

**Returns**: Array of recently enrolled devices

**Example**:
```typescript
const enrolled = await reports.getRecentlyEnrolledDevices({
  dateRange: {
    startDate: new Date('2024-01-01'),
    endDate: new Date('2024-01-31'),
  },
  operatingSystem: ['iOS', 'Android'],
});
```

##### getRecentlyWipedDevices()

```typescript
async getRecentlyWipedDevices(
  options?: DeviceActionReportOptions
): Promise<RecentlyWipedDevice[]>
```

Retrieves recently wiped devices (default: last 30 days).

**Parameters**:
- `options`: Optional filtering and date range options

**Returns**: Array of recently wiped devices

**Example**:
```typescript
const wiped = await reports.getRecentlyWipedDevices({
  dateRange: {
    startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});
```

##### generateDeviceActionReport()

```typescript
async generateDeviceActionReport(
  options?: DeviceActionReportOptions
): Promise<DeviceActionReport>
```

Generates a comprehensive report including all sections.

**Parameters**:
- `options`: Optional filtering and configuration options

**Returns**: Complete device action report with summary

**Example**:
```typescript
const report = await reports.generateDeviceActionReport({
  dateRange: {
    startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});
```

##### Export Methods

```typescript
exportToJSON(report: DeviceActionReport): string
exportPendingActionsToCSV(actions: PendingDeviceAction[]): string
exportEnrollmentsToCSV(enrollments: RecentlyEnrolledDevice[]): string
exportWipesToCSV(wipes: RecentlyWipedDevice[]): string
exportToHTML(report: DeviceActionReport): string
```

Export reports to various formats.

## Microsoft Graph API Endpoints

The module uses the following Microsoft Graph API endpoints:

### Get Managed Devices

```
GET /deviceManagement/managedDevices
```

**Query Parameters**:
- `$select`: Specific properties to retrieve
- `$filter`: Filter criteria (enrolledDateTime, operatingSystem, etc.)
- `$top`: Limit number of results

**Properties Used**:
- `id`, `deviceName`, `userPrincipalName`, `userDisplayName`
- `operatingSystem`, `osVersion`
- `manufacturer`, `model`, `serialNumber`
- `enrolledDateTime`, `lastSyncDateTime`
- `complianceState`, `managementAgent`
- `managedDeviceOwnerType`, `deviceEnrollmentType`
- `azureADDeviceId`, `retireAfterDateTime`

### Get Device Actions

```
GET /deviceManagement/managedDevices/{id}/deviceManagementTroubleshootingEvents
```

Retrieves action history and troubleshooting events for a device.

### Date Filtering

```
GET /deviceManagement/managedDevices?$filter=enrolledDateTime gt {date}
```

Filter devices by enrollment date or other date-based criteria.

## Data Models

### DeviceActionReportOptions

```typescript
interface DeviceActionReportOptions {
  dateRange?: {
    startDate: Date;
    endDate: Date;
  };
  actionType?: DeviceActionType[];
  status?: DeviceActionStatus[];
  ownershipType?: DeviceOwnershipType[];
  operatingSystem?: string[];
  top?: number;
}
```

### PendingDeviceAction

```typescript
interface PendingDeviceAction {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  operatingSystem: string;
  osVersion: string;
  actionType: DeviceActionType;
  actionStatus: DeviceActionStatus;
  actionRequestedDateTime: string;
  actionInitiatedBy?: string;
  lastSyncDateTime: string;
  retireAfterDateTime?: string;
  managedDeviceOwnerType: DeviceOwnershipType;
  estimatedCompletionTime?: string;
  actionDetails?: string;
}
```

### RecentlyEnrolledDevice

```typescript
interface RecentlyEnrolledDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  operatingSystem: string;
  osVersion: string;
  manufacturer: string;
  model: string;
  serialNumber: string;
  enrolledDateTime: string;
  lastSyncDateTime: string;
  complianceState: string;
  managementAgent: string;
  managedDeviceOwnerType: DeviceOwnershipType;
  enrollmentType: string;
  azureADDeviceId: string;
  daysSinceEnrollment: number;
  isAssigned: boolean;
  assignedUsers?: string[];
  assignedGroups?: string[];
}
```

### RecentlyWipedDevice

```typescript
interface RecentlyWipedDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  operatingSystem: string;
  osVersion: string;
  wipeRequestedDateTime: string;
  wipeCompletedDateTime: string;
  wipeInitiatedBy?: string;
  wipeType: 'wipe' | 'retire' | 'factoryReset';
  managedDeviceOwnerType: DeviceOwnershipType;
  wipeDurationMinutes: number;
  wipeStatus: 'success' | 'failed' | 'partial';
  wipeErrorCode?: string;
  wipeErrorMessage?: string;
  lastKnownLocation?: string;
}
```

### DeviceActionReport

```typescript
interface DeviceActionReport {
  summary: DeviceActionSummary;
  pendingActions: PendingDeviceAction[];
  recentEnrollments: RecentlyEnrolledDevice[];
  recentWipes: RecentlyWipedDevice[];
  reportPeriod: {
    startDate: string;
    endDate: string;
  };
  metadata: {
    reportName: string;
    generatedAt: string;
    generatedBy: string;
    recordCount: number;
    parameters?: Record<string, any>;
  };
}
```

## Export Formats

### JSON Export

```typescript
const report = await reports.generateDeviceActionReport();
const json = reports.exportToJSON(report);
await fs.writeFile('report.json', json);
```

**Output**: Formatted JSON with all report data and metadata.

### CSV Export

```typescript
// Export pending actions
const pending = await reports.getPendingRetireWipeRequests();
const csv = reports.exportPendingActionsToCSV(pending);
await fs.writeFile('pending-actions.csv', csv);

// Export enrollments
const enrolled = await reports.getRecentlyEnrolledDevices();
const enrollCsv = reports.exportEnrollmentsToCSV(enrolled);

// Export wipes
const wiped = await reports.getRecentlyWipedDevices();
const wipeCsv = reports.exportWipesToCSV(wiped);
```

**Output**: Comma-separated values with headers, suitable for Excel.

### HTML Export

```typescript
const report = await reports.generateDeviceActionReport();
const html = reports.exportToHTML(report);
await fs.writeFile('report.html', html);
```

**Output**: Styled HTML report with:
- Summary statistics cards
- Sortable data tables
- Professional formatting
- Print-friendly layout

## Usage Examples

### Example 1: Monitor Pending Actions

```typescript
const reports = new DeviceActionReports(graphClient);
const pending = await reports.getPendingRetireWipeRequests();

// Filter for actions older than 7 days
const staleActions = pending.filter((action) => {
  const requested = new Date(action.actionRequestedDateTime);
  const daysOld = (Date.now() - requested.getTime()) / (1000 * 60 * 60 * 24);
  return daysOld > 7;
});

console.log(`Stale pending actions: ${staleActions.length}`);
```

### Example 2: Weekly Enrollment Report

```typescript
const reports = new DeviceActionReports(graphClient);

// Get last week's enrollments
const enrolled = await reports.getRecentlyEnrolledDevices({
  dateRange: {
    startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});

// Group by operating system
const byOS = enrolled.reduce((acc, device) => {
  acc[device.operatingSystem] = (acc[device.operatingSystem] || 0) + 1;
  return acc;
}, {} as Record<string, number>);

console.log('Weekly enrollments by OS:', byOS);
```

### Example 3: Wipe Success Monitoring

```typescript
const reports = new DeviceActionReports(graphClient);
const wiped = await reports.getRecentlyWipedDevices();

// Calculate success rate
const successful = wiped.filter((d) => d.wipeStatus === 'success').length;
const successRate = (successful / wiped.length) * 100;

// Calculate average duration
const avgDuration =
  wiped.reduce((sum, d) => sum + d.wipeDurationMinutes, 0) / wiped.length;

console.log(`Wipe success rate: ${successRate.toFixed(1)}%`);
console.log(`Average duration: ${avgDuration.toFixed(1)} minutes`);
```

### Example 4: Comprehensive Monthly Report

```typescript
const reports = new DeviceActionReports(graphClient);

// Generate monthly report
const report = await reports.generateDeviceActionReport({
  dateRange: {
    startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});

// Export to all formats
const json = reports.exportToJSON(report);
await fs.writeFile('monthly-report.json', json);

const html = reports.exportToHTML(report);
await fs.writeFile('monthly-report.html', html);
```

### Example 5: Filter by Ownership Type

```typescript
const reports = new DeviceActionReports(graphClient);

// Corporate devices only
const corporateEnrollments = await reports.getRecentlyEnrolledDevices({
  ownershipType: ['company'],
  dateRange: {
    startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});

// Personal devices only
const personalEnrollments = await reports.getRecentlyEnrolledDevices({
  ownershipType: ['personal'],
  dateRange: {
    startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
    endDate: new Date(),
  },
});

console.log(`Corporate: ${corporateEnrollments.length}`);
console.log(`Personal: ${personalEnrollments.length}`);
```

## Best Practices

### 1. Date Range Selection

- Use appropriate date ranges to limit data volume
- Default ranges: 7 days for enrollments, 30 days for wipes
- Consider your organization's device turnover rate

```typescript
// Good: Specific date range
const options = {
  dateRange: {
    startDate: new Date('2024-01-01'),
    endDate: new Date('2024-01-31'),
  },
};

// Avoid: Very large date ranges
// This may timeout or return excessive data
```

### 2. Filtering

- Apply filters to reduce API calls and processing time
- Filter by OS, ownership type, or action status
- Use `top` parameter to limit results

```typescript
const options: DeviceActionReportOptions = {
  operatingSystem: ['iOS', 'Android'],
  ownershipType: ['company'],
  top: 500,
};
```

### 3. Error Handling

- Always wrap API calls in try-catch blocks
- Handle rate limiting gracefully
- Log errors for troubleshooting

```typescript
try {
  const report = await reports.generateDeviceActionReport();
} catch (error) {
  console.error('Report generation failed:', error);
  // Implement retry logic or notification
}
```

### 4. Performance

- Use parallel requests for independent operations
- Cache results when generating multiple reports
- Limit unnecessary data fetching

```typescript
// Good: Parallel execution
const [pending, enrolled, wiped] = await Promise.all([
  reports.getPendingRetireWipeRequests(),
  reports.getRecentlyEnrolledDevices(),
  reports.getRecentlyWipedDevices(),
]);

// Avoid: Sequential execution
// Takes 3x longer
```

### 5. Data Retention

- Export and archive reports regularly
- Store in appropriate format (JSON for processing, HTML for viewing)
- Implement automated reporting schedules

## Troubleshooting

### Common Issues

#### 1. Authentication Failures

**Problem**: "Unauthorized" or "Invalid credentials" errors

**Solution**:
- Verify client ID, tenant ID, and client secret
- Check API permissions are granted
- Ensure admin consent is provided

```typescript
// Test authentication separately
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID!,
  process.env.AZURE_CLIENT_ID!,
  process.env.AZURE_CLIENT_SECRET!
);

const token = await credential.getToken('https://graph.microsoft.com/.default');
console.log('Auth successful:', !!token);
```

#### 2. Empty Results

**Problem**: Reports return no data

**Solution**:
- Verify devices exist in the specified date range
- Check filter criteria aren't too restrictive
- Ensure devices are properly enrolled in Intune

```typescript
// Debug: Check total device count
const allDevices = await graphClient
  .api('/deviceManagement/managedDevices')
  .top(1)
  .get();

console.log('Total devices:', allDevices['@odata.count']);
```

#### 3. Rate Limiting

**Problem**: "429 Too Many Requests" errors

**Solution**:
- Implement retry logic with exponential backoff
- Reduce request frequency
- Use pagination efficiently

```typescript
// Implement retry logic
async function retryableRequest<T>(fn: () => Promise<T>, retries = 3): Promise<T> {
  try {
    return await fn();
  } catch (error: any) {
    if (error.statusCode === 429 && retries > 0) {
      const delay = Math.pow(2, 3 - retries) * 1000;
      await new Promise((resolve) => setTimeout(resolve, delay));
      return retryableRequest(fn, retries - 1);
    }
    throw error;
  }
}
```

#### 4. Incomplete Wipe Data

**Problem**: Missing wipe action data

**Solution**:
- Wipe actions may not always be tracked in troubleshooting events
- Check device's `retireAfterDateTime` property
- Some actions may not persist after device removal

#### 5. Assignment Data Unavailable

**Problem**: Assignment information not available

**Solution**:
- Assignment tracking requires additional Graph API calls
- May need `Group.Read.All` permission
- Some enrollment types don't support assignment data

## Performance Considerations

### API Call Optimization

1. **Batch Requests**: Use parallel requests for independent operations
2. **Selective Properties**: Request only needed properties with `$select`
3. **Pagination**: Handle large result sets efficiently
4. **Caching**: Cache device data when generating multiple reports

### Memory Management

- Process large datasets in chunks
- Stream data when possible
- Clear references to large objects after use

### Timeouts

- Set appropriate timeouts for API calls
- Implement timeout handlers
- Consider async processing for large reports

### Example: Optimized Large-Scale Report

```typescript
async function generateOptimizedReport() {
  const reports = new DeviceActionReports(graphClient);

  // Limit date range
  const options: DeviceActionReportOptions = {
    dateRange: {
      startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
      endDate: new Date(),
    },
    top: 1000, // Limit results
  };

  // Execute in parallel
  const [pending, enrolled, wiped] = await Promise.all([
    reports.getPendingRetireWipeRequests(options),
    reports.getRecentlyEnrolledDevices(options),
    reports.getRecentlyWipedDevices(options),
  ]);

  return { pending, enrolled, wiped };
}
```

## Additional Resources

### Microsoft Documentation

- [Microsoft Graph API - Managed Devices](https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [Intune Device Actions](https://learn.microsoft.com/en-us/mem/intune/remote-actions/)
- [Device Enrollment](https://learn.microsoft.com/en-us/mem/intune/enrollment/)

### Related Modules

- `device-hardware-reports.ts` - Hardware inventory reports
- `mobile-device-inventory-reports.ts` - Mobile device inventory
- `intune-enrollment-tracking-reports.ts` - Enrollment tracking

### Support

For issues, questions, or contributions:
- Review the unit tests in `device-action-reports.test.ts`
- Check example implementations in `examples/device-action-reports-examples.ts`
- Consult the main Intune Reporting Guide

## License

This module is part of the Intune Reporting Dashboard project.

## Version History

- **1.0.0** (2024-02-06): Initial release
  - Report #31: Pending retire and wipe requests
  - Report #32: Recently enrolled devices
  - Report #33: Recently wiped devices
  - JSON, CSV, and HTML export formats
  - Comprehensive test suite
  - Usage examples and documentation
