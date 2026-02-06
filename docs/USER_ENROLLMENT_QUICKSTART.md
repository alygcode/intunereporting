# User Enrollment Report - Quick Start Guide

Get started with the User Enrollment Report module in 5 minutes.

## Prerequisites

1. **Azure AD App Registration** with the following permissions:
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementServiceConfig.Read.All`
   - `User.Read.All`
   - `Directory.Read.All`

2. **Environment Variables**:
   ```bash
   export AZURE_TENANT_ID="your-tenant-id"
   export AZURE_CLIENT_ID="your-app-client-id"
   export AZURE_CLIENT_SECRET="your-app-secret"
   ```

3. **Node.js Dependencies**:
   ```bash
   npm install @microsoft/microsoft-graph-client @azure/identity
   ```

## 5-Minute Quick Start

### Step 1: Create Authentication

Create a file `auth.ts`:

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

export function createGraphClient(): Client {
  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID!,
    process.env.AZURE_CLIENT_ID!,
    process.env.AZURE_CLIENT_SECRET!
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });

  return Client.initWithMiddleware({ authProvider });
}
```

### Step 2: Generate Your First Report

Create `enrollment-report.ts`:

```typescript
import { createGraphClient } from './auth';
import { UserEnrollmentReport } from './src/reports/user-enrollment-report';
import { AppConfig } from './src/types';
import * as fs from 'fs';

async function main() {
  // Create Graph client
  const graphClient = createGraphClient();

  // Create config
  const config: AppConfig = {
    tenantId: process.env.AZURE_TENANT_ID!,
    clientId: process.env.AZURE_CLIENT_ID!,
    clientSecret: process.env.AZURE_CLIENT_SECRET!,
    outputDir: './reports'
  };

  // Create report
  const report = new UserEnrollmentReport(graphClient, config);

  // Generate report
  console.log('Generating enrollment report...');
  const result = await report.execute({
    includeTrends: true,
    includeFailures: true,
    includePlatformStats: true,
    trendDays: 30
  });

  // Print summary
  console.log('\n📊 Enrollment Summary:');
  console.log(`Total Devices: ${result.summary?.totalDevices || 0}`);
  console.log(`Total Users: ${result.summary?.totalUsers || 0}`);
  console.log(`Success Rate: ${result.summary?.successRate || 0}%`);
  console.log(`Failed Enrollments: ${result.summary?.failedEnrollments || 0}`);

  // Save reports
  const outputDir = './reports';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  fs.writeFileSync(`${outputDir}/enrollment.json`, report.exportAsJSON(result));
  fs.writeFileSync(`${outputDir}/enrollment.csv`, report.exportAsCSV(result));
  fs.writeFileSync(`${outputDir}/enrollment.html`, report.exportAsHTML(result));

  console.log('\n✅ Reports saved to ./reports/');
}

main().catch(console.error);
```

### Step 3: Run the Report

```bash
npx ts-node enrollment-report.ts
```

## Common Tasks

### Get Enrollment Statistics

```typescript
const report = new UserEnrollmentReport(graphClient, config);

const endDate = new Date();
const startDate = new Date();
startDate.setDate(startDate.getDate() - 30);

const stats = await report.getEnrollmentStatistics(startDate, endDate);

console.log('Last 30 Days:');
console.log(`  Enrolled: ${stats.enrolledLast30Days}`);
console.log(`  Success Rate: ${stats.enrollmentSuccessRate}%`);
console.log(`  Avg Devices/User: ${stats.averageDevicesPerUser.toFixed(2)}`);
```

### Get User's Devices

```typescript
const userId = 'user-id-here';
const devices = await report.getUserDevices(userId);

devices.forEach(device => {
  console.log(`${device.deviceName} (${device.platform})`);
});
```

### Get Enrollment Failures

```typescript
const failures = await report.getEnrollmentFailures(30, 50);

console.log(`Found ${failures.length} failures`);

failures.forEach(failure => {
  console.log(`\n${failure.userPrincipalName}`);
  console.log(`  Category: ${failure.failureCategory}`);
  console.log(`  Reason: ${failure.failureReason}`);
  console.log(`  Steps to fix:`);
  failure.troubleshootingSteps.slice(0, 3).forEach((step, i) => {
    console.log(`    ${i + 1}. ${step}`);
  });
});
```

### Track Device Ownership

```typescript
const ownership = await report.getDeviceOwnershipBreakdown();

console.log('Device Ownership:');
console.log(`  Corporate: ${ownership.corporate.count} (${ownership.corporate.percentage}%)`);
console.log(`  Personal: ${ownership.personal.count} (${ownership.personal.percentage}%)`);
console.log(`  Unknown: ${ownership.unknown.count} (${ownership.unknown.percentage}%)`);
```

### Generate Enrollment Trends

```typescript
const trends = await report.getEnrollmentTrends(30);

console.log('Last 7 Days:');
trends.slice(-7).forEach(trend => {
  console.log(`${trend.date}:`);
  console.log(`  Total: ${trend.totalEnrollments}`);
  console.log(`  Success: ${trend.successfulEnrollments}`);
  console.log(`  Failed: ${trend.failedEnrollments}`);
});
```

### Get Platform Statistics

```typescript
const platformStats = await report.getDeviceEnrollmentByPlatform();

platformStats.forEach(stat => {
  console.log(`\n${stat.platform}:`);
  console.log(`  Total: ${stat.totalDevices}`);
  console.log(`  Last 7d: ${stat.enrolledLast7Days}`);
  console.log(`  Top Model: ${stat.topModels[0]?.model || 'N/A'}`);
});
```

## Output Formats

### JSON Export

```typescript
const jsonReport = report.exportAsJSON(result);
fs.writeFileSync('report.json', jsonReport);
```

Includes complete data structure with all details.

### CSV Export

```typescript
const csvReport = report.exportAsCSV(result);
fs.writeFileSync('report.csv', csvReport);
```

Ideal for Excel and data analysis.

### HTML Export

```typescript
const htmlReport = report.exportAsHTML(result);
fs.writeFileSync('report.html', htmlReport);
```

Beautiful formatted report for sharing with stakeholders.

## Error Handling

```typescript
try {
  const report = new UserEnrollmentReport(graphClient, config);
  const result = await report.execute();
} catch (error: any) {
  if (error.statusCode === 401) {
    console.error('Authentication failed. Check your credentials.');
  } else if (error.statusCode === 403) {
    console.error('Insufficient permissions. Check API permissions.');
  } else if (error.statusCode === 429) {
    console.error('Rate limited. The module will automatically retry.');
  } else {
    console.error('Error:', error.message);
  }
}
```

## Filtering Options

### Filter by Time Period

```typescript
await report.execute({
  startDate: new Date('2024-01-01'),
  endDate: new Date('2024-12-31'),
  includeTrends: true,
  trendDays: 30
});
```

### Filter by Platform

```typescript
await report.execute({
  platformFilter: ['iOS', 'Android'],
  includePlatformStats: true
});
```

### Filter by User

```typescript
await report.execute({
  userFilter: ['user-id-1', 'user-id-2'],
  includeUserDetails: true
});
```

## Best Practices

1. **Start Small**: Begin with basic statistics, then add more features
2. **Cache Data**: Store results to avoid repeated API calls
3. **Schedule Reports**: Run reports during off-peak hours
4. **Monitor Failures**: Set up alerts for high failure rates
5. **Regular Reviews**: Generate weekly/monthly reports for trend analysis

## Next Steps

1. Explore more examples: `src/reports/examples/user-enrollment-report-example.ts`
2. Read full documentation: `docs/USER_ENROLLMENT_REPORT.md`
3. Customize report options for your needs
4. Set up automated report generation
5. Integrate with your monitoring systems

## Troubleshooting

### "No devices found"

- Verify devices are enrolled in Intune
- Check API permissions
- Ensure your service principal has access

### "Authentication failed"

- Verify environment variables are set
- Check tenant ID, client ID, and secret
- Ensure app registration is configured correctly

### "Rate limiting errors"

- The module handles this automatically with retry logic
- Reduce concurrent requests if issues persist
- Spread out report generation times

## Getting Help

- Full API Reference: `docs/USER_ENROLLMENT_REPORT.md`
- Examples: `src/reports/examples/user-enrollment-report-example.ts`
- TypeScript Types: Defined in `src/reports/user-enrollment-report.ts`

## Complete Working Example

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { UserEnrollmentReport } from './src/reports/user-enrollment-report';
import { AppConfig } from './src/types';
import * as fs from 'fs';
import * as dotenv from 'dotenv';

dotenv.config();

async function generateEnrollmentReport() {
  try {
    // 1. Setup authentication
    const credential = new ClientSecretCredential(
      process.env.AZURE_TENANT_ID!,
      process.env.AZURE_CLIENT_ID!,
      process.env.AZURE_CLIENT_SECRET!
    );

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default']
    });

    const graphClient = Client.initWithMiddleware({ authProvider });

    // 2. Create config
    const config: AppConfig = {
      tenantId: process.env.AZURE_TENANT_ID!,
      clientId: process.env.AZURE_CLIENT_ID!,
      clientSecret: process.env.AZURE_CLIENT_SECRET!,
      outputDir: './reports'
    };

    // 3. Create report instance
    const report = new UserEnrollmentReport(graphClient, config);

    // 4. Generate comprehensive report
    console.log('Generating enrollment report...');
    const result = await report.execute({
      includeTrends: true,
      includeFailures: true,
      includePlatformStats: true,
      includeUserDetails: false,
      trendDays: 30,
      maxFailures: 100
    });

    // 5. Display summary
    console.log('\n📊 Enrollment Summary:');
    console.log('='.repeat(50));
    console.log(`Total Devices: ${result.summary?.totalDevices || 0}`);
    console.log(`Total Users: ${result.summary?.totalUsers || 0}`);
    console.log(`Success Rate: ${result.summary?.successRate || 0}%`);
    console.log(`Failed Enrollments: ${result.summary?.failedEnrollments || 0}`);

    // 6. Save reports in multiple formats
    const outputDir = './reports';
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
    }

    fs.writeFileSync(
      `${outputDir}/enrollment-report.json`,
      report.exportAsJSON(result)
    );

    fs.writeFileSync(
      `${outputDir}/enrollment-report.csv`,
      report.exportAsCSV(result)
    );

    fs.writeFileSync(
      `${outputDir}/enrollment-report.html`,
      report.exportAsHTML(result)
    );

    console.log('\n✅ Reports saved successfully!');
    console.log(`   JSON: ${outputDir}/enrollment-report.json`);
    console.log(`   CSV:  ${outputDir}/enrollment-report.csv`);
    console.log(`   HTML: ${outputDir}/enrollment-report.html`);

  } catch (error: any) {
    console.error('❌ Error generating report:', error.message);

    if (error.statusCode === 401) {
      console.error('   Authentication failed. Check your credentials.');
    } else if (error.statusCode === 403) {
      console.error('   Insufficient permissions. Verify API permissions.');
    }
  }
}

// Run the report
generateEnrollmentReport();
```

Save this as `quick-start.ts` and run:

```bash
npx ts-node quick-start.ts
```

That's it! You now have a complete enrollment report in JSON, CSV, and HTML formats.
