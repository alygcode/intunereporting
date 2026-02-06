# Application Deployment Report - Quick Start Guide

Get started with the Application Deployment Report module in 5 minutes.

## Prerequisites

1. **Azure AD App Registration** with permissions:
   - `DeviceManagementApps.Read.All`
   - `DeviceManagementManagedDevices.Read.All`

2. **Node.js** version 14 or higher

3. **Microsoft Graph Client** configured and authenticated

## Installation

The module is included in the Intune Reporting project. No additional installation required.

## Quick Start

### Step 1: Import the Module

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { AppDeploymentReport, AppPlatform } from './reports/app-deployment-report';
import { AppConfig } from './types';
```

### Step 2: Initialize Graph Client

```typescript
// Using client credentials
import { createClientCredentialsAuth, AuthFlowType } from './auth/graph-auth';

const authManager = createClientCredentialsAuth({
  clientId: process.env.CLIENT_ID!,
  clientSecret: process.env.CLIENT_SECRET!,
  tenantId: process.env.TENANT_ID!
});

const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

// Create Graph client
const graphClient = Client.init({
  authProvider: (done) => {
    done(null, token);
  }
});
```

### Step 3: Create Report Instance

```typescript
const config: AppConfig = {
  // Your app configuration
  authentication: {
    tenantId: process.env.TENANT_ID!,
    clientId: process.env.CLIENT_ID!,
    clientSecret: process.env.CLIENT_SECRET!,
    authMethod: 'clientSecret'
  },
  reports: { enabled: ['app-deployment-report'], disabled: [], settings: {} },
  output: {
    defaultFormat: 'json',
    directory: './reports',
    includeTimestamp: true,
    compression: false
  },
  scheduler: { enabled: false, timezone: 'UTC', schedules: [] },
  logging: { level: 'info', file: 'app.log', console: true, maxSize: '10m', maxFiles: 5 }
};

const report = new AppDeploymentReport(graphClient, config);
```

### Step 4: Run Your First Report

```typescript
// Get complete deployment report
const reportData = await report.execute();

console.log(`Total Apps: ${reportData.summary.statistics.totalApps}`);
console.log(`Total Devices: ${reportData.summary.statistics.totalDevices}`);
console.log(`Success Rate: ${reportData.summary.statistics.successRate}%`);
```

## Common Use Cases

### 1. Get All Managed Apps

```typescript
const allApps = await report.getAllManagedApps();

console.log(`Total Apps: ${allApps.length}`);
allApps.forEach(app => {
  console.log(`${app.displayName} (${app.platform})`);
});
```

### 2. Get iOS Apps Only

```typescript
const iosApps = await report.getAllManagedApps({
  platform: AppPlatform.IOS
});

console.log(`iOS Apps: ${iosApps.length}`);
```

### 3. Get Installation Summary

```typescript
const summaries = await report.getAllAppInstallSummaries();

summaries.forEach(summary => {
  console.log(`\n${summary.appName}`);
  console.log(`  Installed: ${summary.installedDeviceCount}`);
  console.log(`  Failed: ${summary.failedDeviceCount}`);
  console.log(`  Pending: ${summary.pendingInstallDeviceCount}`);
});
```

### 4. Find Failed Installations

```typescript
const failures = await report.getFailedInstallations();

console.log(`Total Failures: ${failures.length}\n`);

failures.forEach(failure => {
  console.log(`App: ${failure.appName}`);
  console.log(`Device: ${failure.deviceName}`);
  console.log(`Error: ${failure.errorDescription}`);
  console.log(`Troubleshooting: ${failure.troubleshootingLink}\n`);
});
```

### 5. Check Update Compliance

```typescript
const compliance = await report.getAppUpdateCompliance();

compliance.forEach(app => {
  if (app.updateCompliancePercentage < 100) {
    console.log(`${app.appName}: ${app.updateCompliancePercentage}% compliance`);
    console.log(`  ${app.devicesRequiringUpdate} devices need update`);
  }
});
```

### 6. Export to File

```typescript
// Export as JSON
await report.exportDeploymentReport({
  format: 'json',
  outputDir: './reports',
  includeTimestamp: true
});

// Export as CSV
await report.exportDeploymentReport({
  format: 'csv',
  outputDir: './reports',
  includeTimestamp: true
});

// Export as HTML
await report.exportDeploymentReport({
  format: 'html',
  outputDir: './reports',
  includeTimestamp: true
});
```

## Complete Example

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { AppDeploymentReport, AppPlatform } from './reports/app-deployment-report';
import { createClientCredentialsAuth, AuthFlowType } from './auth/graph-auth';
import * as dotenv from 'dotenv';

dotenv.config();

async function main() {
  try {
    // 1. Authenticate
    const authManager = createClientCredentialsAuth({
      clientId: process.env.CLIENT_ID!,
      clientSecret: process.env.CLIENT_SECRET!,
      tenantId: process.env.TENANT_ID!
    });

    const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

    // 2. Create Graph client
    const graphClient = Client.init({
      authProvider: (done) => done(null, token)
    });

    // 3. Create report
    const report = new AppDeploymentReport(graphClient, config);

    // 4. Get all managed apps
    console.log('Fetching managed apps...');
    const apps = await report.getAllManagedApps();
    console.log(`Found ${apps.length} apps\n`);

    // 5. Get installation summaries
    console.log('Fetching installation summaries...');
    const summaries = await report.getAllAppInstallSummaries();
    console.log(`Retrieved ${summaries.length} summaries\n`);

    // 6. Find failed installations
    console.log('Checking for failures...');
    const failures = await report.getFailedInstallations({ maxResults: 50 });
    console.log(`Found ${failures.length} failures\n`);

    if (failures.length > 0) {
      console.log('Top Failed Apps:');
      const failuresByApp = new Map();
      failures.forEach(f => {
        failuresByApp.set(f.appName, (failuresByApp.get(f.appName) || 0) + 1);
      });

      Array.from(failuresByApp.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 5)
        .forEach(([app, count]) => {
          console.log(`  ${app}: ${count} failures`);
        });
    }

    // 7. Export reports
    console.log('\nExporting reports...');
    await report.exportDeploymentReport({
      format: 'html',
      outputDir: './reports',
      includeTimestamp: true
    });
    console.log('✓ Report exported successfully!');

  } catch (error) {
    console.error('Error:', error);
    process.exit(1);
  }
}

main();
```

## Environment Variables

Create a `.env` file:

```env
# Azure AD Authentication
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret

# Optional: Output Configuration
OUTPUT_DIR=./reports
LOG_LEVEL=info
```

## Running the Example

```bash
# Install dependencies
npm install

# Run the example
npm run example:app-deployment

# Or with ts-node
npx ts-node examples/app-deployment-example.ts
```

## Filtering Options

### Filter by Platform

```typescript
// iOS only
const iosApps = await report.getAllManagedApps({
  platform: AppPlatform.IOS
});

// Android only
const androidApps = await report.getAllManagedApps({
  platform: AppPlatform.ANDROID
});

// Windows only
const windowsApps = await report.getAllManagedApps({
  platform: AppPlatform.WINDOWS
});
```

### Filter by Publisher

```typescript
const microsoftApps = await report.getAllManagedApps({
  publisher: 'Microsoft'
});
```

### Filter by App Name

```typescript
const outlookApps = await report.getAllManagedApps({
  appName: 'Outlook'
});
```

### Combined Filters

```typescript
const googleAndroidApps = await report.getAllManagedApps({
  platform: AppPlatform.ANDROID,
  publisher: 'Google'
});
```

### Limit Results

```typescript
// Get device statuses for first 20 apps only
const deviceStatuses = await report.getAllAppDeviceStatuses({
  maxResults: 20
});
```

## Performance Tips

1. **Start with summaries** - They're fast and give you an overview
2. **Use filters** - Reduce data volume with platform/publisher filters
3. **Limit results** - Use `maxResults` for device-level queries
4. **Run during off-peak** - Large queries should run during off-peak hours
5. **Cache results** - Store results locally to avoid repeated API calls

## Troubleshooting

### Authentication Errors

```
Error: Failed to authenticate
```

**Solution**: Verify your credentials and API permissions in Azure AD.

### No Data Returned

```
Total Apps: 0
```

**Solution**:
- Ensure apps are deployed in Intune
- Verify API permissions are granted and admin consented
- Check authentication token is valid

### Rate Limiting

```
Error: 429 Too Many Requests
```

**Solution**: The module handles this automatically with retry logic. For very large tenants:
- Use `maxResults` to limit data
- Run reports during off-peak hours
- Increase delays between batch operations

### Timeout Errors

```
Error: Request timeout
```

**Solution**:
- Reduce `maxResults` value
- Process fewer apps at a time
- Check network connectivity

## Next Steps

- Read the [Full Documentation](./app-deployment-report.md)
- Check out [Usage Examples](../src/reports/examples/app-deployment-examples.ts)
- Explore [Graph API Documentation](https://docs.microsoft.com/en-us/graph/)

## Support

For issues or questions:
- Review the [full documentation](./app-deployment-report.md)
- Check [Microsoft Graph API docs](https://docs.microsoft.com/en-us/graph/)
- See [Intune troubleshooting guide](https://docs.microsoft.com/en-us/mem/intune/apps/troubleshoot-app-install)

## License

See the main project LICENSE file.
