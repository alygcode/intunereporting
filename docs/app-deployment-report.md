# Application Deployment Report Module

Comprehensive application deployment reporting for Microsoft Intune using Graph API.

## Overview

The Application Deployment Report module provides detailed insights into application deployments across your Intune-managed devices, including:

- **Managed Apps Inventory**: All apps (iOS, Android, Windows, macOS, Web)
- **Installation Status**: Per-app deployment statistics
- **Device-level Details**: Installation status for each device
- **Failed Installations**: Detailed error analysis with troubleshooting links
- **Update Compliance**: Track app versions and update requirements
- **Multiple Export Formats**: JSON, CSV, and HTML output

## Features

### 1. Managed Apps Retrieval

Get all managed applications across all platforms with comprehensive metadata:

```typescript
import { AppDeploymentReport, AppPlatform } from './reports/app-deployment-report';

const report = new AppDeploymentReport(graphClient, config);

// Get all apps
const allApps = await report.getAllManagedApps();

// Filter by platform
const iosApps = await report.getAllManagedApps({ platform: AppPlatform.IOS });
const androidApps = await report.getAllManagedApps({ platform: AppPlatform.ANDROID });
const windowsApps = await report.getAllManagedApps({ platform: AppPlatform.WINDOWS });

// Filter by publisher
const microsoftApps = await report.getAllManagedApps({ publisher: 'Microsoft' });

// Get specific app
const app = await report.getManagedAppById('app-id-here');
```

### 2. Installation Status Per App

Track installation metrics for each application:

```typescript
// Get all installation summaries
const summaries = await report.getAllAppInstallSummaries();

// Get summary for specific app
const summary = await report.getAppInstallSummary('app-id', 'App Name');

// Summary includes:
// - installedDeviceCount
// - failedDeviceCount
// - pendingInstallDeviceCount
// - notInstalledDeviceCount
// - installedUserCount
// - failedUserCount
```

### 3. Device-level Deployment Details

Get granular device-level installation status:

```typescript
// Get all device statuses
const deviceStatuses = await report.getAllAppDeviceStatuses();

// Get device statuses for specific app
const appDeviceStatuses = await report.getAppDeviceStatuses('app-id', 'App Name');

// Get user-level statuses for specific app
const userStatuses = await report.getAppUserStatuses('app-id', 'App Name');

// Each device status includes:
// - deviceName, deviceId
// - userName, userPrincipalName
// - platform, osVersion
// - installState (installed, failed, pending, etc.)
// - errorCode, installStateDetail
// - lastSyncDateTime
```

### 4. Failed Installations with Error Details

Identify and troubleshoot failed app installations:

```typescript
// Get all failed installations
const failures = await report.getFailedInstallations();

// Get failures for specific app
const appFailures = await report.getFailedInstallationsByApp('app-id', 'App Name');

// Each failure includes:
// - appName, deviceName, userName
// - errorCode, errorDescription
// - failureReason (human-readable)
// - troubleshootingLink (Microsoft Docs)
// - installStateDetail
// - lastSyncDateTime
```

Common error codes and their meanings:

| Error Code | Description |
|------------|-------------|
| `0x80073CFF` | Package not found |
| `0x87D1041C` | Network connection required |
| `0x87D1041D` | Insufficient storage space |
| `0x87D13B7D` | Device not compliant |
| `0x87D13BA2` | Minimum OS version not met |
| `0x8007065E` | Authentication error |
| `0x80072EE7` | Server name could not be resolved |

### 5. App Update Compliance

Track app versions and identify devices requiring updates:

```typescript
// Get update compliance for all apps
const compliance = await report.getAppUpdateCompliance();

// Get compliance for specific app
const appCompliance = await report.getAppUpdateComplianceById('app-id');

// Each compliance record includes:
// - currentVersion, latestVersion
// - devicesOnCurrentVersion
// - devicesOnOlderVersion
// - devicesRequiringUpdate
// - updateCompliancePercentage
// - deviceDetails (devices needing updates)
```

### 6. Export to Multiple Formats

Export reports in JSON, CSV, or HTML format:

```typescript
// Export complete deployment report
const jsonPath = await report.exportDeploymentReport({
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true
});

const csvPath = await report.exportDeploymentReport({
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true
});

const htmlPath = await report.exportDeploymentReport({
  format: 'html',
  outputDir: './reports',
  includeTimestamp: true
});

// Export specific data sets
await report.exportInstallSummaries({ format: 'csv', outputDir: './reports' });
await report.exportFailedInstallations({ format: 'csv', outputDir: './reports' });
await report.exportUpdateCompliance({ format: 'csv', outputDir: './reports' });
```

## Graph API Endpoints Used

The module uses the following Microsoft Graph API endpoints:

- `GET /deviceAppManagement/mobileApps` - Get all managed apps
- `GET /deviceAppManagement/mobileApps/{id}` - Get specific app
- `GET /deviceAppManagement/mobileApps/{id}/installSummary` - Get installation summary
- `GET /deviceAppManagement/mobileApps/{id}/deviceStatuses` - Get device-level statuses
- `GET /deviceAppManagement/mobileApps/{id}/userStatuses` - Get user-level statuses

## Required Permissions

Your Azure AD application needs the following Microsoft Graph API permissions:

**Application Permissions** (for service/daemon apps):
- `DeviceManagementApps.Read.All`
- `DeviceManagementManagedDevices.Read.All`

**Delegated Permissions** (for user-context apps):
- `DeviceManagementApps.Read.All`
- `DeviceManagementManagedDevices.Read.All`

## Usage Examples

### Complete Deployment Report

```typescript
import { AppDeploymentReport } from './reports/app-deployment-report';

const report = new AppDeploymentReport(graphClient, config);

// Execute complete report
const reportData = await report.execute();

console.log('Deployment Statistics:');
console.log(`Total Apps: ${reportData.summary.statistics.totalApps}`);
console.log(`Total Devices: ${reportData.summary.statistics.totalDevices}`);
console.log(`Success Rate: ${reportData.summary.statistics.successRate}%`);
console.log(`Failure Rate: ${reportData.summary.statistics.failureRate}%`);
```

### Filter by Platform

```typescript
// Get iOS apps only
const iosApps = await report.getAllManagedApps({
  platform: AppPlatform.IOS
});

// Get Android apps from specific publisher
const googleApps = await report.getAllManagedApps({
  platform: AppPlatform.ANDROID,
  publisher: 'Google'
});
```

### Analyze Failed Installations

```typescript
const failures = await report.getFailedInstallations({ maxResults: 50 });

// Group by app
const failuresByApp = new Map();
failures.forEach(failure => {
  failuresByApp.set(
    failure.appName,
    (failuresByApp.get(failure.appName) || 0) + 1
  );
});

// Display top failed apps
console.log('Top Failed Apps:');
Array.from(failuresByApp.entries())
  .sort((a, b) => b[1] - a[1])
  .slice(0, 10)
  .forEach(([appName, count]) => {
    console.log(`${appName}: ${count} failures`);
  });
```

### Monitor Update Compliance

```typescript
const compliance = await report.getAppUpdateCompliance();

// Find apps with low compliance
const lowCompliance = compliance.filter(c => c.updateCompliancePercentage < 80);

console.log('Apps with Low Update Compliance:');
lowCompliance.forEach(app => {
  console.log(`${app.appName}: ${app.updateCompliancePercentage}%`);
  console.log(`  Devices requiring update: ${app.devicesRequiringUpdate}`);
});
```

## TypeScript Types

### Core Types

```typescript
interface ManagedAppInfo {
  id: string;
  displayName: string;
  publisher: string;
  platform: AppPlatform;
  appType: string;
  version?: string;
  bundleId?: string;
  packageId?: string;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  // ... more properties
}

interface AppInstallSummary {
  appId: string;
  appName: string;
  installedDeviceCount: number;
  failedDeviceCount: number;
  pendingInstallDeviceCount: number;
  notApplicableDeviceCount: number;
  notInstalledDeviceCount: number;
  installedUserCount: number;
  failedUserCount: number;
}

interface AppDeviceStatus {
  id: string;
  appId: string;
  appName: string;
  deviceId: string;
  deviceName: string;
  userName?: string;
  userPrincipalName?: string;
  platform?: string;
  osVersion?: string;
  installState: InstallState;
  errorCode?: string;
  lastSyncDateTime?: Date;
}

interface FailedInstallation {
  appId: string;
  appName: string;
  deviceId: string;
  deviceName: string;
  errorCode?: string;
  errorDescription?: string;
  failureReason?: string;
  troubleshootingLink?: string;
  lastSyncDateTime?: Date;
}

interface AppUpdateCompliance {
  appId: string;
  appName: string;
  currentVersion: string;
  devicesOnCurrentVersion: number;
  devicesOnOlderVersion: number;
  devicesRequiringUpdate: number;
  updateCompliancePercentage: number;
  deviceDetails: AppUpdateDeviceDetail[];
}
```

### Enums

```typescript
enum AppPlatform {
  IOS = 'iOS',
  ANDROID = 'Android',
  WINDOWS = 'Windows',
  MACOS = 'macOS',
  WEB = 'Web',
  UNKNOWN = 'Unknown'
}

enum InstallState {
  INSTALLED = 'installed',
  FAILED = 'failed',
  NOT_INSTALLED = 'notInstalled',
  AVAILABLE = 'available',
  PENDING = 'pending',
  UNINSTALLING = 'uninstalling',
  DOWNLOADING = 'downloading',
  UNKNOWN = 'unknown'
}
```

## Error Handling

The module includes comprehensive error handling:

```typescript
try {
  const report = new AppDeploymentReport(graphClient, config);
  const reportData = await report.execute();
  // Process report data
} catch (error) {
  console.error('Error generating deployment report:', error);
  // Handle error appropriately
}
```

All methods include:
- **Retry logic** with exponential backoff for transient errors
- **Pagination** support for large data sets
- **Partial results** on partial failures (returns what was successfully retrieved)
- **Detailed logging** of errors and warnings

## Performance Considerations

### Batch Processing

The module processes apps in batches to avoid rate limiting:

```typescript
// Limit number of apps processed
const deviceStatuses = await report.getAllAppDeviceStatuses({
  maxResults: 50  // Process only 50 apps
});
```

### Parallel Execution

Multiple data sources are fetched in parallel for better performance:

```typescript
const [apps, summaries, statuses, failures] = await Promise.all([
  report.getAllManagedApps(),
  report.getAllAppInstallSummaries(),
  report.getAllAppDeviceStatuses({ maxResults: 20 }),
  report.getFailedInstallations({ maxResults: 20 })
]);
```

### Rate Limiting

The module automatically handles Microsoft Graph API rate limits:
- Automatic retry with exponential backoff
- Built-in delays between batch operations
- Respects HTTP 429 (Too Many Requests) responses

## Best Practices

1. **Start with summaries**: Use `getAllAppInstallSummaries()` first to get an overview
2. **Filter appropriately**: Use filter options to reduce data volume
3. **Limit results**: Use `maxResults` option for device-level queries
4. **Export regularly**: Schedule regular exports for historical tracking
5. **Monitor failures**: Set up alerts based on failure thresholds
6. **Track compliance**: Regular compliance checks help identify update issues

## Integration Examples

### Scheduled Reporting

```typescript
import cron from 'node-cron';

// Run daily at 2 AM
cron.schedule('0 2 * * *', async () => {
  const report = new AppDeploymentReport(graphClient, config);

  await report.exportDeploymentReport({
    format: 'html',
    outputDir: './daily-reports',
    includeTimestamp: true
  });

  console.log('Daily deployment report generated');
});
```

### Alert on High Failure Rate

```typescript
const summaries = await report.getAllAppInstallSummaries();

summaries.forEach(summary => {
  const total = summary.installedDeviceCount + summary.failedDeviceCount;
  const failureRate = total > 0 ? (summary.failedDeviceCount / total) * 100 : 0;

  if (failureRate > 20) {
    console.warn(`Alert: ${summary.appName} has ${failureRate}% failure rate`);
    // Send alert notification
  }
});
```

### Export to Database

```typescript
const reportData = await report.execute();

// Save to database
await db.deploymentReports.create({
  timestamp: new Date(),
  statistics: reportData.summary.statistics,
  data: reportData.data
});
```

## Troubleshooting

### No Data Returned

- Verify API permissions are granted and consented
- Check authentication token is valid
- Ensure tenant has managed apps deployed

### Slow Performance

- Use `maxResults` option to limit data volume
- Fetch only required data (e.g., summaries instead of device statuses)
- Run reports during off-peak hours

### Rate Limiting Errors

- Module handles rate limiting automatically
- For very large tenants, consider scheduling reports to run over longer periods
- Use batch processing with appropriate delays

## Support and Documentation

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Intune App Management](https://docs.microsoft.com/en-us/mem/intune/apps/)
- [App Installation Error Codes](https://docs.microsoft.com/en-us/mem/intune/apps/app-install-error-codes)
- [Troubleshooting App Installation](https://docs.microsoft.com/en-us/mem/intune/apps/troubleshoot-app-install)

## License

See the main project LICENSE file.
