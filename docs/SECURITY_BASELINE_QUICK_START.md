# Security Baseline Report - Quick Start Guide

## Installation & Setup

```bash
# Install dependencies
npm install

# Configure environment variables
cp .env.example .env
# Edit .env with your Azure AD credentials
```

## Quick Examples

### 1. Generate Full Report

```typescript
import { SecurityBaselineReport } from './reports/security-baseline-report';

const report = new SecurityBaselineReport(graphClient, config);
const data = await report.execute();
console.log(data.summary);
```

### 2. Get All Baselines

```typescript
const baselines = await report.getAllSecurityBaselines();
console.log(`Found ${baselines.length} baselines`);
```

### 3. Check Device Compliance

```typescript
const intents = await report.getSecurityBaselineIntents();
const deviceStates = await report.getBaselineCompliancePerDevice(intents[0].id);

const compliant = deviceStates.filter(d => d.state === 'compliant').length;
console.log(`${compliant}/${deviceStates.length} devices compliant`);
```

### 4. Find Non-Compliant Settings

```typescript
const issues = await report.getNonCompliantSettings();
console.log(`Found ${issues.length} non-compliant settings`);

issues.forEach(issue => {
  console.log(`${issue.deviceName}: ${issue.settingName}`);
  console.log(`  Expected: ${issue.expectedValue}`);
  console.log(`  Current: ${issue.currentValue}`);
});
```

### 5. Export to Multiple Formats

```typescript
const comprehensiveReport = await report.generateComprehensiveReport({
  includeDeviceDetails: true,
  includeSettingDetails: true
});

const files = await report.exportReport(comprehensiveReport, {
  outputDir: './reports',
  formats: ['json', 'csv', 'html']
});

console.log('Exported:', files);
```

## Run Examples

```bash
# Show all examples
ts-node src/reports/examples/security-baseline-example.ts

# Run specific example
ts-node src/reports/examples/security-baseline-example.ts 1
```

## Key Methods

| Method | Description |
|--------|-------------|
| `getAllSecurityBaselines()` | Get all security baseline templates |
| `getSecurityBaselineIntents()` | Get applied baselines (intents) |
| `getBaselineCompliancePerDevice(intentId)` | Get device-level compliance |
| `getBaselineDeviceStateSummary(intentId)` | Get summary statistics |
| `getBaselineSettingStates(intentId)` | Get setting-level details |
| `getNonCompliantSettings()` | Get all non-compliant settings |
| `getSecurityScoreTrends(days)` | Analyze trends over time |
| `generateComprehensiveReport(options)` | Generate full report |
| `exportReport(report, options)` | Export to JSON/CSV/HTML |

## Report Options

```typescript
interface SecurityBaselineReportOptions {
  includeDeviceDetails?: boolean;      // Include device-level data
  includeSettingDetails?: boolean;     // Include setting-level data
  templateId?: string;                 // Filter by template
  intentId?: string;                   // Filter by intent
  includeTrends?: boolean;             // Include trend analysis
  trendDays?: number;                  // Days for trends (default: 30)
  outputDir?: string;                  // Export directory
  formats?: ['json' | 'csv' | 'html']; // Export formats
}
```

## Common Use Cases

### Weekly Compliance Report

```typescript
async function weeklyReport() {
  const report = new SecurityBaselineReport(graphClient, config);
  const data = await report.generateComprehensiveReport({
    includeDeviceDetails: true,
    includeSettingDetails: false
  });

  await report.exportReport(data, {
    outputDir: './reports/weekly',
    formats: ['html', 'csv']
  });
}
```

### Remediation List

```typescript
async function getRemediationList() {
  const report = new SecurityBaselineReport(graphClient, config);
  const issues = await report.getNonCompliantSettings();

  await report.exportNonCompliantSettingsToCSV(
    issues,
    './reports/remediation-list.csv'
  );
}
```

### Compliance Dashboard Data

```typescript
async function getDashboardData() {
  const report = new SecurityBaselineReport(graphClient, config);
  const intents = await report.getSecurityBaselineIntents();

  const dashboardData = [];
  for (const intent of intents) {
    const summary = await report.getBaselineDeviceStateSummary(intent.id);
    dashboardData.push({
      baseline: intent.displayName,
      compliant: summary.compliantCount,
      nonCompliant: summary.nonCompliantCount,
      total: (summary.compliantCount || 0) + (summary.nonCompliantCount || 0)
    });
  }

  return dashboardData;
}
```

## Graph API Endpoints

- `/deviceManagement/templates` - Get baseline templates
- `/deviceManagement/intents` - Get baseline intents
- `/deviceManagement/intents/{id}/deviceStates` - Get device states
- `/deviceManagement/intents/{id}/deviceStateSummary` - Get summary
- `/deviceManagement/intents/{id}/deviceSettingStateSummaries` - Get settings

## Required Permissions

- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementManagedDevices.Read.All`

## Troubleshooting

**No data returned?**
- Check if baselines are configured in Intune
- Verify devices are enrolled and assigned
- Ensure proper API permissions

**Rate limiting errors?**
- Module includes automatic retry logic
- Consider adding delays between batch operations

**Authentication fails?**
- Verify `.env` credentials
- Check admin consent for API permissions

## Next Steps

1. Review the [full documentation](./SECURITY_BASELINE_REPORT.md)
2. Run the example scripts
3. Customize for your use case
4. Schedule automated reports

## Support

For detailed documentation, see: `docs/SECURITY_BASELINE_REPORT.md`
