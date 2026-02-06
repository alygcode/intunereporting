# Device Compliance Report Module

## Overview

The Device Compliance Report module provides comprehensive device compliance reporting capabilities for Microsoft Intune using the Microsoft Graph API. This module enables you to monitor, analyze, and report on device compliance status across your organization.

## Features

### Core Capabilities

1. **Compliance Policies Management**
   - Retrieve all device compliance policies
   - Get policy settings and assignments
   - Filter policies by platform

2. **Device Compliance Status**
   - Get compliance status per policy
   - Retrieve device state summaries
   - Track compliance percentages

3. **Detailed Device Information**
   - Get comprehensive device compliance details
   - View policy states for each device
   - Access setting-level compliance information

4. **Non-Compliance Analysis**
   - Identify non-compliant devices
   - Get detailed reasons for non-compliance
   - Receive remediation recommendations

5. **Compliance Trends**
   - Track compliance over time
   - Identify compliance improvements or declines
   - View historical compliance data

6. **Multiple Export Formats**
   - JSON format for programmatic access
   - CSV format for spreadsheet analysis
   - HTML format for visual reporting

## Installation

```bash
npm install @microsoft/microsoft-graph-client
npm install @azure/identity
```

## Quick Start

```typescript
import { DeviceComplianceReport } from './reports/device-compliance-report';
import { Client } from '@microsoft/microsoft-graph-client';

// Initialize Graph client (see Authentication section)
const graphClient = await createGraphClient();

// Create report instance
const report = new DeviceComplianceReport(graphClient, config);

// Execute full compliance report
const reportData = await report.execute();

console.log(`Compliance: ${reportData.summary.compliancePercentage}%`);
console.log(`Non-Compliant Devices: ${reportData.summary.nonCompliantDevices}`);
```

## Authentication

### Using Client Credentials

```typescript
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { Client } from '@microsoft/microsoft-graph-client';

const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID,
  process.env.AZURE_CLIENT_ID,
  process.env.AZURE_CLIENT_SECRET
);

const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

const graphClient = Client.initWithMiddleware({ authProvider });
```

### Required Permissions

Your Azure AD application needs the following Microsoft Graph API permissions:

- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementApps.Read.All`

## API Reference

### DeviceComplianceReport Class

#### Constructor

```typescript
constructor(graphClient: Client, config: AppConfig)
```

**Parameters:**
- `graphClient`: Microsoft Graph API client instance
- `config`: Application configuration object

#### Methods

##### execute()

Executes the complete compliance report, gathering all compliance data.

```typescript
async execute(): Promise<ReportData>
```

**Returns:** Complete report data including metadata, devices, policies, and summary

**Example:**
```typescript
const reportData = await report.execute();
```

##### getAllCompliancePolicies()

Retrieves all device compliance policies from Intune.

```typescript
async getAllCompliancePolicies(options?: ComplianceFilterOptions): Promise<DeviceCompliancePolicy[]>
```

**Parameters:**
- `options` (optional): Filter options for policies

**Returns:** Array of compliance policies

**Example:**
```typescript
const policies = await report.getAllCompliancePolicies();
console.log(`Found ${policies.length} policies`);

// Filter by platform
const windowsPolicies = await report.getAllCompliancePolicies({
  platform: CompliancePlatform.WINDOWS_10
});
```

##### getAllPolicyDeviceStateSummaries()

Retrieves device state summaries for all compliance policies.

```typescript
async getAllPolicyDeviceStateSummaries(): Promise<PolicyDeviceStateSummary[]>
```

**Returns:** Array of policy device state summaries with compliance statistics

**Example:**
```typescript
const summaries = await report.getAllPolicyDeviceStateSummaries();

for (const summary of summaries) {
  console.log(`${summary.policyName}: ${summary.compliancePercentage}%`);
}
```

##### getPolicyDeviceStatuses()

Gets device statuses for a specific compliance policy.

```typescript
async getPolicyDeviceStatuses(policyId: string): Promise<DevicePolicyComplianceStatus[]>
```

**Parameters:**
- `policyId`: ID of the compliance policy

**Returns:** Array of device compliance statuses for the specified policy

**Example:**
```typescript
const statuses = await report.getPolicyDeviceStatuses('policy-id-123');
```

##### getAllDeviceComplianceDetails()

Retrieves detailed compliance information for all managed devices.

```typescript
async getAllDeviceComplianceDetails(options?: ComplianceFilterOptions): Promise<DeviceComplianceDetails[]>
```

**Parameters:**
- `options` (optional): Filter options for devices

**Returns:** Array of detailed device compliance information

**Example:**
```typescript
const devices = await report.getAllDeviceComplianceDetails();

// Filter by compliance state
const nonCompliantDevices = await report.getAllDeviceComplianceDetails({
  complianceState: ComplianceState.NON_COMPLIANT
});
```

##### getDeviceComplianceDetails()

Gets detailed compliance information for a specific device.

```typescript
async getDeviceComplianceDetails(device: any): Promise<DeviceComplianceDetails>
```

**Parameters:**
- `device`: Device object from Graph API

**Returns:** Detailed compliance information for the device

##### getDevicePolicyStates()

Gets compliance policy states for a specific device.

```typescript
async getDevicePolicyStates(deviceId: string): Promise<DevicePolicyState[]>
```

**Parameters:**
- `deviceId`: ID of the device

**Returns:** Array of policy states for the device

**Example:**
```typescript
const policyStates = await report.getDevicePolicyStates('device-id-123');

for (const state of policyStates) {
  console.log(`${state.policyName}: ${state.state}`);
  console.log(`  Settings: ${state.settingCount}`);
}
```

##### getNonCompliantDevicesWithReasons()

Retrieves all non-compliant devices with detailed reasons for non-compliance.

```typescript
async getNonCompliantDevicesWithReasons(options?: ComplianceFilterOptions): Promise<NonCompliantDevice[]>
```

**Parameters:**
- `options` (optional): Filter options

**Returns:** Array of non-compliant devices with detailed reasons

**Example:**
```typescript
const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

for (const device of nonCompliantDevices) {
  console.log(`${device.deviceName}:`);

  for (const policy of device.nonCompliantPolicies) {
    console.log(`  Policy: ${policy.policyName}`);

    for (const setting of policy.nonCompliantSettings) {
      console.log(`    - ${setting.settingName}: ${setting.errorDescription}`);

      if (setting.remediationActions) {
        console.log(`      Remediation: ${setting.remediationActions[0]}`);
      }
    }
  }
}
```

##### getComplianceTrends()

Generates compliance trends over a specified time period.

```typescript
async getComplianceTrends(days: number = 30): Promise<ComplianceTrends | null>
```

**Parameters:**
- `days`: Number of days to analyze (default: 30)

**Returns:** Compliance trends data or null if unavailable

**Example:**
```typescript
const trends = await report.getComplianceTrends(30);

console.log(`Trend: ${trends.trend}`);
console.log(`Average Compliance: ${trends.averageCompliancePercentage}%`);
console.log(`Change: ${trends.complianceChange}%`);
```

## Type Definitions

### DeviceCompliancePolicy

```typescript
interface DeviceCompliancePolicy {
  id: string;
  displayName: string;
  description?: string;
  platform: CompliancePlatform;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  version: number;
  assignmentCount: number;
  settingsCount: number;
  scheduledActionsCount: number;
  isAssigned: boolean;
}
```

### PolicyDeviceStateSummary

```typescript
interface PolicyDeviceStateSummary {
  policyId: string;
  policyName: string;
  compliantDeviceCount: number;
  nonCompliantDeviceCount: number;
  errorDeviceCount: number;
  conflictDeviceCount: number;
  notApplicableDeviceCount: number;
  inGracePeriodCount: number;
  unknownDeviceCount: number;
  remediatedDeviceCount: number;
  notAssignedDeviceCount: number;
  totalDeviceCount: number;
  compliancePercentage: number;
}
```

### DeviceComplianceDetails

```typescript
interface DeviceComplianceDetails {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  osVersion: string;
  overallComplianceState: ComplianceState;
  lastSyncDateTime: Date;
  policies: DevicePolicyState[];
  compliancePoliciesCount: number;
  compliantPoliciesCount: number;
  nonCompliantPoliciesCount: number;
  isManaged: boolean;
  isSupervised?: boolean;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  imei?: string;
}
```

### NonCompliantDevice

```typescript
interface NonCompliantDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  osVersion: string;
  complianceState: ComplianceState;
  nonCompliantPolicies: NonCompliantPolicyInfo[];
  totalNonCompliantSettings: number;
  lastSyncDateTime: Date;
  gracePeriodExpirationDateTime?: Date;
  deviceModel?: string;
  manufacturer?: string;
  userEmail?: string;
}
```

### ComplianceTrends

```typescript
interface ComplianceTrends {
  startDate: Date;
  endDate: Date;
  dataPoints: ComplianceTrendDataPoint[];
  averageCompliancePercentage: number;
  complianceChange: number;
  trend: 'improving' | 'declining' | 'stable';
  peakCompliance: ComplianceTrendDataPoint;
  lowestCompliance: ComplianceTrendDataPoint;
}
```

### Enumerations

#### ComplianceState

```typescript
enum ComplianceState {
  COMPLIANT = 'compliant',
  NON_COMPLIANT = 'noncompliant',
  IN_GRACE_PERIOD = 'inGracePeriod',
  CONFIG_MANAGER = 'configManager',
  ERROR = 'error',
  UNKNOWN = 'unknown',
  CONFLICT = 'conflict',
  NOT_ASSIGNED = 'notAssigned'
}
```

#### CompliancePlatform

```typescript
enum CompliancePlatform {
  ANDROID = 'android',
  ANDROID_WORK_PROFILE = 'androidWorkProfile',
  ANDROID_AOSP = 'androidAOSP',
  IOS = 'iOS',
  MAC_OS = 'macOS',
  WINDOWS_10 = 'windows10',
  WINDOWS_81 = 'windows81',
  WINDOWS_PHONE_81 = 'windowsPhone81'
}
```

## Usage Examples

### Example 1: Basic Compliance Report

```typescript
const report = new DeviceComplianceReport(graphClient, config);
const reportData = await report.execute();

console.log(`Total Devices: ${reportData.summary.totalDevices}`);
console.log(`Compliance: ${reportData.summary.compliancePercentage}%`);

// Export to files
const outputDir = './reports';
await OutputFormatter.format(reportData, 'json', outputDir);
await OutputFormatter.format(reportData, 'csv', outputDir);
await OutputFormatter.format(reportData, 'html', outputDir);
```

### Example 2: Identify Non-Compliant Devices

```typescript
const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

for (const device of nonCompliantDevices) {
  console.log(`Device: ${device.deviceName} (${device.userPrincipalName})`);
  console.log(`Total Issues: ${device.totalNonCompliantSettings}`);

  for (const policy of device.nonCompliantPolicies) {
    console.log(`\n  Policy: ${policy.policyName}`);

    for (const setting of policy.nonCompliantSettings) {
      console.log(`    Issue: ${setting.settingName}`);
      console.log(`    Description: ${setting.errorDescription}`);

      if (setting.remediationActions && setting.remediationActions.length > 0) {
        console.log(`    Remediation:`);
        setting.remediationActions.forEach(action => {
          console.log(`      - ${action}`);
        });
      }
    }
  }
}
```

### Example 3: Monitor Compliance Trends

```typescript
const trends = await report.getComplianceTrends(30);

if (trends) {
  console.log(`\nCompliance Trend: ${trends.trend}`);
  console.log(`Average: ${trends.averageCompliancePercentage}%`);
  console.log(`Change: ${trends.complianceChange > 0 ? '+' : ''}${trends.complianceChange}%`);

  console.log(`\nPeak Compliance:`);
  console.log(`  Date: ${trends.peakCompliance.date.toLocaleDateString()}`);
  console.log(`  Rate: ${trends.peakCompliance.compliancePercentage}%`);

  console.log(`\nLowest Compliance:`);
  console.log(`  Date: ${trends.lowestCompliance.date.toLocaleDateString()}`);
  console.log(`  Rate: ${trends.lowestCompliance.compliancePercentage}%`);
}
```

### Example 4: Platform-Specific Reports

```typescript
// Get Windows-only compliance data
const windowsPolicies = await report.getAllCompliancePolicies({
  platform: CompliancePlatform.WINDOWS_10
});

const allDevices = await report.getAllDeviceComplianceDetails();
const windowsDevices = allDevices.filter(d =>
  d.platform.toLowerCase().includes('windows')
);

const windowsCompliant = windowsDevices.filter(d =>
  d.overallComplianceState === ComplianceState.COMPLIANT
).length;

console.log(`Windows Devices: ${windowsDevices.length}`);
console.log(`Compliant: ${windowsCompliant}`);
console.log(`Compliance Rate: ${Math.round((windowsCompliant / windowsDevices.length) * 100)}%`);
```

### Example 5: Critical Issues Detection

```typescript
const reportData = await report.execute();

if (reportData.summary.criticalIssues && reportData.summary.criticalIssues.length > 0) {
  console.log('\nCritical Issues Detected:\n');

  for (const issue of reportData.summary.criticalIssues) {
    console.log(`[${issue.severity.toUpperCase()}] ${issue.issueType}`);
    console.log(`  ${issue.description}`);
    console.log(`  Affected: ${issue.affectedCount}`);
    console.log(`  Recommendation: ${issue.recommendation}`);

    if (issue.relatedPolicyName) {
      console.log(`  Related Policy: ${issue.relatedPolicyName}`);
    }

    console.log('');
  }
}
```

## Error Handling

The module includes comprehensive error handling:

```typescript
try {
  const report = new DeviceComplianceReport(graphClient, config);
  const reportData = await report.execute();

  // Process report data
  console.log('Report generated successfully');
} catch (error) {
  if (error.statusCode === 403) {
    console.error('Insufficient permissions to access compliance data');
  } else if (error.statusCode === 429) {
    console.error('Rate limit exceeded, please retry later');
  } else {
    console.error('Error generating report:', error.message);
  }
}
```

## Graph API Endpoints Used

The module interacts with the following Microsoft Graph API endpoints:

| Endpoint | Purpose |
|----------|---------|
| `/deviceManagement/deviceCompliancePolicies` | Retrieve all compliance policies |
| `/deviceManagement/deviceCompliancePolicies/{id}/deviceStateSummary` | Get policy device state summary |
| `/deviceManagement/deviceCompliancePolicies/{id}/deviceStatuses` | Get device statuses for a policy |
| `/deviceManagement/managedDevices` | Retrieve all managed devices |
| `/deviceManagement/managedDevices/{id}/deviceCompliancePolicyStates` | Get device policy states |
| `/deviceManagement/detectedApps` | Get discovered applications (for context) |

## Performance Considerations

### Batch Processing

The module uses batch processing for large device collections:

```typescript
// Processes devices in batches of 50 to avoid timeout
const batchSize = 50;
for (let i = 0; i < devices.length; i += batchSize) {
  const batch = devices.slice(i, Math.min(i + batchSize, devices.length));
  // Process batch
}
```

### Pagination

Automatic pagination handling for large result sets:

```typescript
// Automatically handles pagination using BaseReport.getAllPages()
const response = await this.graphClient
  .api('/deviceManagement/managedDevices')
  .top(999)
  .get();

const allDevices = await this.getAllPages(response);
```

### Retry Logic

Built-in retry logic with exponential backoff:

```typescript
// Automatically retries failed requests
const response = await this.retryGraphCall(() =>
  this.graphClient.api('/deviceManagement/deviceCompliancePolicies').get()
);
```

## Best Practices

1. **Regular Monitoring**: Schedule compliance reports to run daily or weekly
2. **Filter When Possible**: Use filter options to reduce data transfer
3. **Cache Results**: Cache report data to avoid excessive API calls
4. **Handle Rate Limits**: Implement retry logic for rate limit errors
5. **Export Reports**: Save reports to files for historical tracking
6. **Monitor Trends**: Use trend analysis to identify compliance patterns
7. **Act on Critical Issues**: Address critical issues highlighted in the summary

## Troubleshooting

### Common Issues

#### Issue: "Insufficient Permissions"

**Solution:** Ensure your Azure AD application has the required permissions:
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementConfiguration.Read.All`

#### Issue: "Rate Limit Exceeded"

**Solution:** The module includes retry logic, but for very large tenants, consider:
- Running reports during off-peak hours
- Using filters to reduce data volume
- Implementing exponential backoff

#### Issue: "No Data Returned"

**Solution:** Verify:
- Devices are enrolled in Intune
- Compliance policies are assigned
- Devices have synced recently

## Contributing

To contribute to this module:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## License

This module is part of the Intune Reporting Dashboard project.

## Support

For issues and questions:
- Check the [examples](../src/reports/examples/device-compliance-report-example.ts)
- Review the [API documentation](https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-devicecompliancepolicy)
- Open an issue on GitHub

## Changelog

### Version 1.0.0
- Initial release
- Complete compliance policy retrieval
- Device compliance status tracking
- Non-compliant device analysis
- Compliance trends reporting
- Multiple export formats
- Comprehensive error handling
- Retry logic with exponential backoff
- Batch processing for large datasets
- Critical issues detection
- Remediation recommendations
