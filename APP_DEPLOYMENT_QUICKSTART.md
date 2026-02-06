# Application Deployment Report - Quick Start Guide

## Installation

```bash
npm install
```

## Basic Usage

### 1. Import the Module

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { AppDeploymentReport, AppPlatform, InstallState } from './src/reports/app-deployment-report';
import { AppConfig } from './src/types';
```

### 2. Initialize the Report

```typescript
// Initialize Graph Client (assumes authentication is configured)
const graphClient = Client.init({
  authProvider: yourAuthProvider
});

// Create config
const config: AppConfig = {
  // Your configuration here
};

// Create report instance
const report = new AppDeploymentReport(graphClient, config);
```

### 3. Run Complete Report

```typescript
// Get everything at once
const reportData = await report.execute();

console.log('Report Statistics:', reportData.summary?.statistics);
console.log('Total Records:', reportData.metadata.recordCount);
```

## Common Use Cases

### Get All Apps

```typescript
// Get all managed apps
const allApps = await report.getAllManagedApps();
console.log(`Total apps: ${allApps.length}`);

// Filter by platform
const iosApps = await report.getAllManagedApps({
  platform: AppPlatform.IOS
});

// Filter by publisher
const msApps = await report.getAllManagedApps({
  publisher: 'Microsoft'
});
```

### Get Installation Status

```typescript
// Get installation summaries for all apps
const summaries = await report.getAllAppInstallSummaries();

summaries.forEach(summary => {
  console.log(`${summary.appName}:`);
  console.log(`  Installed: ${summary.installedDeviceCount}`);
  console.log(`  Failed: ${summary.failedDeviceCount}`);
  console.log(`  Pending: ${summary.pendingInstallDeviceCount}`);
});
```

### Get Failed Installations

```typescript
// Get all failures
const failures = await report.getFailedInstallations();

console.log(`Total failures: ${failures.length}`);

failures.forEach(failure => {
  console.log(`\n${failure.appName} on ${failure.deviceName}`);
  console.log(`Error: ${failure.errorDescription}`);
  console.log(`Troubleshoot: ${failure.troubleshootingLink}`);
});
```

### Get App Assignments

```typescript
// Get all assignments
const assignments = await report.getAllAppAssignments();

// Count by intent
const required = assignments.filter(a => a.intent === 'required');
const available = assignments.filter(a => a.intent === 'available');

console.log(`Required: ${required.length}`);
console.log(`Available: ${available.length}`);

// Get assignments for specific app
const appAssignments = await report.getAppAssignments('app-id');
const stats = await report.getAppAssignmentStats('app-id');

console.log('Assignment Stats:', stats);
```

### Check Update Compliance

```typescript
// Get update compliance for all apps
const compliance = await report.getAppUpdateCompliance();

// Find apps needing updates
const needsUpdate = compliance.filter(c => c.devicesRequiringUpdate > 0);

console.log(`Apps needing updates: ${needsUpdate.length}`);

needsUpdate.forEach(app => {
  console.log(`\n${app.appName}`);
  console.log(`  Current version: ${app.currentVersion}`);
  console.log(`  Devices needing update: ${app.devicesRequiringUpdate}`);
  console.log(`  Compliance: ${app.updateCompliancePercentage}%`);
});
```

### Export Reports

```typescript
// Export complete report as CSV
const csvPath = await report.exportDeploymentReport({
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true
});
console.log(`CSV report: ${csvPath}`);

// Export only failures
const failuresPath = await report.exportFailedInstallations({
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true
});

// Export assignments
const assignmentsPath = await report.exportAppAssignments({
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true
});
```

## Filtering Options

All methods support flexible filtering:

```typescript
const options = {
  platform: AppPlatform.IOS,        // iOS, ANDROID, WINDOWS, MACOS, WEB
  appName: 'Office',                // Partial name match
  publisher: 'Microsoft',           // Partial publisher match
  deviceId: 'device-123',           // Specific device
  userName: 'user@company.com',     // Specific user
  installState: InstallState.FAILED, // Filter by install state
  maxResults: 50                    // Limit results
};

const apps = await report.getAllManagedApps(options);
```

## Platform Types

```typescript
enum AppPlatform {
  IOS = 'iOS',
  ANDROID = 'Android',
  WINDOWS = 'Windows',
  MACOS = 'macOS',
  WEB = 'Web',
  UNKNOWN = 'Unknown'
}
```

## Install States

```typescript
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

## Assignment Intents

```typescript
enum AssignmentIntent {
  REQUIRED = 'required',                            // Must install
  AVAILABLE = 'available',                          // Available in portal
  UNINSTALL = 'uninstall',                         // Should remove
  AVAILABLE_WITHOUT_ENROLLMENT = 'availableWithoutEnrollment'
}
```

## Performance Tips

### 1. Limit Results for Large Tenants

```typescript
// Get first 20 apps only
const apps = await report.getAllManagedApps({ maxResults: 20 });

// Get device statuses for first 10 apps
const statuses = await report.getAllAppDeviceStatuses({ maxResults: 10 });
```

### 2. Use Platform Filtering

```typescript
// Process one platform at a time
const iosApps = await report.getAllManagedApps({ platform: AppPlatform.IOS });
const iosStatuses = await report.getAllAppDeviceStatuses({ platform: AppPlatform.IOS });
```

### 3. Export Large Datasets

```typescript
// Instead of loading all data in memory, export directly
await report.exportDeploymentReport({
  format: 'csv',
  outputDir: './reports'
});
```

## Error Handling

The module includes built-in retry logic and error handling:

```typescript
try {
  const reportData = await report.execute();
  console.log('Success!', reportData.summary);
} catch (error) {
  console.error('Report failed:', error.message);
  // Module will have logged detailed errors
}
```

## Running Examples

Run the included examples:

```typescript
import examples from './src/reports/examples/app-deployment-examples';

// Run specific example
await examples.example1_CompleteDeploymentReport(graphClient, config);
await examples.example2_GetManagedApps(graphClient, config);
await examples.example5_GetFailedInstallations(graphClient, config);

// Run all examples
await examples.runAllExamples(graphClient, config);
```

## Common Scenarios

### Daily Failure Report

```typescript
async function dailyFailureReport() {
  const report = new AppDeploymentReport(graphClient, config);

  // Get failures
  const failures = await report.getFailedInstallations();

  // Export to CSV for email
  const path = await report.exportFailedInstallations({
    format: 'csv',
    outputDir: './reports/daily',
    includeTimestamp: true
  });

  // Send email with CSV (your email logic here)
  console.log(`Failures exported to: ${path}`);
}
```

### Weekly Compliance Report

```typescript
async function weeklyComplianceReport() {
  const report = new AppDeploymentReport(graphClient, config);

  // Get update compliance
  const compliance = await report.getAppUpdateCompliance();

  // Calculate statistics
  const totalDevices = compliance.reduce((sum, c) =>
    sum + c.devicesOnCurrentVersion + c.devicesOnOlderVersion, 0);
  const upToDate = compliance.reduce((sum, c) =>
    sum + c.devicesOnCurrentVersion, 0);
  const complianceRate = Math.round((upToDate / totalDevices) * 100);

  console.log(`Overall update compliance: ${complianceRate}%`);

  // Export report
  await report.exportUpdateCompliance({
    format: 'html',
    outputDir: './reports/weekly',
    includeTimestamp: true
  });
}
```

### App Deployment Audit

```typescript
async function auditAppDeployment(appId: string) {
  const report = new AppDeploymentReport(graphClient, config);

  // Get app details
  const app = await report.getManagedAppById(appId);
  console.log(`Auditing: ${app.displayName}`);

  // Get assignments
  const assignments = await report.getAppAssignments(appId);
  console.log(`Assignments: ${assignments.length}`);

  // Get installation status
  const summary = await report.getAppInstallSummary(appId);
  console.log(`Installed: ${summary?.installedDeviceCount}`);
  console.log(`Failed: ${summary?.failedDeviceCount}`);

  // Get failures
  const failures = await report.getFailedInstallationsByApp(appId);
  console.log(`Failures to investigate: ${failures.length}`);

  // Check update compliance
  const compliance = await report.getAppUpdateComplianceById(appId);
  if (compliance) {
    console.log(`Update compliance: ${compliance.updateCompliancePercentage}%`);
  }
}
```

## Troubleshooting

### "Rate limit exceeded" errors

The module includes automatic retry logic, but for very large tenants:

```typescript
// Process in smaller batches
const apps = await report.getAllManagedApps({ maxResults: 10 });

// Add delays between calls if needed
for (const app of apps) {
  const statuses = await report.getAppDeviceStatuses(app.id);
  await new Promise(resolve => setTimeout(resolve, 1000)); // 1 second delay
}
```

### "Timeout" errors

```typescript
// Reduce maxResults
const statuses = await report.getAllAppDeviceStatuses({ maxResults: 5 });

// Or process one platform at a time
const platforms = [AppPlatform.IOS, AppPlatform.ANDROID, AppPlatform.WINDOWS];
for (const platform of platforms) {
  const apps = await report.getAllManagedApps({ platform });
  console.log(`${platform}: ${apps.length} apps`);
}
```

### Missing data

Some data may not be available for all apps:
- Install summaries require app assignments
- Device statuses require app installations
- Update compliance requires version information

```typescript
const summary = await report.getAppInstallSummary(appId);
if (!summary) {
  console.log('No installation summary available (app may not be assigned)');
}
```

## Next Steps

1. Review the full implementation summary: `APP_DEPLOYMENT_IMPLEMENTATION_SUMMARY.md`
2. Check the examples file: `src/reports/examples/app-deployment-examples.ts`
3. Run unit tests: `npm test`
4. Customize filtering and export options for your needs

## Support

For issues or questions:
1. Check the implementation summary for detailed documentation
2. Review the examples for common use cases
3. Check Microsoft Graph API documentation for endpoint details

## File Locations

- **Main Module**: `/home/user/intunereporting/src/reports/app-deployment-report.ts`
- **Examples**: `/home/user/intunereporting/src/reports/examples/app-deployment-examples.ts`
- **Tests**: `/home/user/intunereporting/src/reports/app-deployment-report.test.ts`
- **Summary**: `/home/user/intunereporting/APP_DEPLOYMENT_IMPLEMENTATION_SUMMARY.md`
- **Quick Start**: `/home/user/intunereporting/APP_DEPLOYMENT_QUICKSTART.md` (this file)
