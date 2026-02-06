# Security Baseline Report Module

Comprehensive security baseline reporting module for Microsoft Intune using Graph API.

## Overview

The Security Baseline Report module provides detailed insights into your organization's security baseline compliance across all managed devices. It leverages the Microsoft Graph API to retrieve, analyze, and export security baseline data in multiple formats.

## Features

- **Get All Security Baselines**: Retrieve all available security baseline templates
- **Get Baseline Intents**: List all applied security baselines (intents)
- **Device Compliance**: Get compliance status for each device against security baselines
- **Baseline Setting States**: Retrieve detailed setting-level compliance information
- **Non-Compliant Settings**: Identify all non-compliant settings across the organization
- **Security Score Trends**: Analyze security posture over time
- **Multiple Export Formats**: Export reports to JSON, CSV, and HTML
- **Comprehensive Error Handling**: Automatic retry logic with exponential backoff
- **Pagination Support**: Handle large datasets efficiently

## Prerequisites

- Node.js >= 16.0.0
- Microsoft Azure AD application with appropriate permissions
- Microsoft Intune subscription

## Required Azure AD Permissions

Your Azure AD application must have the following Microsoft Graph API permissions:

- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementConfiguration.ReadWrite.All` (for write operations)
- `DeviceManagementManagedDevices.Read.All`

## Installation

```bash
npm install
```

## Configuration

Create a `.env` file in the project root with your Azure AD credentials:

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

## Usage

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { SecurityBaselineReport } from './reports/security-baseline-report';
import { AppConfig } from './types';

// Initialize Graph Client (see authentication section)
const graphClient = initializeGraphClient();

// Create report instance
const config: AppConfig = getDefaultConfig();
const report = new SecurityBaselineReport(graphClient, config);

// Execute comprehensive report
const reportData = await report.execute();
console.log(reportData);
```

### Get All Security Baselines

```typescript
const baselines = await report.getAllSecurityBaselines();

baselines.forEach(baseline => {
  console.log(`${baseline.displayName} (v${baseline.version})`);
  console.log(`Platform: ${baseline.platformType}`);
  console.log(`ID: ${baseline.id}`);
});
```

### Get Baseline Compliance Per Device

```typescript
// First, get an intent ID
const intents = await report.getSecurityBaselineIntents();
const intentId = intents[0].id;

// Get device compliance for this baseline
const deviceStates = await report.getBaselineCompliancePerDevice(intentId);

// Analyze compliance
const compliantCount = deviceStates.filter(d => d.state === 'compliant').length;
const nonCompliantCount = deviceStates.filter(d => d.state === 'nonCompliant').length;

console.log(`Compliant: ${compliantCount}`);
console.log(`Non-Compliant: ${nonCompliantCount}`);
```

### Get Device State Summary

```typescript
const summary = await report.getBaselineDeviceStateSummary(intentId);

console.log('Summary:');
console.log(`- Compliant: ${summary.compliantCount}`);
console.log(`- Non-Compliant: ${summary.nonCompliantCount}`);
console.log(`- Error: ${summary.errorCount}`);
console.log(`- Unknown: ${summary.unknownCount}`);
```

### Get Non-Compliant Settings

```typescript
const nonCompliantSettings = await report.getNonCompliantSettings();

nonCompliantSettings.forEach(setting => {
  console.log(`\nSetting: ${setting.settingName}`);
  console.log(`Device: ${setting.deviceName}`);
  console.log(`User: ${setting.userPrincipalName}`);
  console.log(`Baseline: ${setting.baselineName}`);
  console.log(`Expected: ${setting.expectedValue}`);
  console.log(`Current: ${setting.currentValue}`);
});
```

### Get Security Score Trends

```typescript
// Get trends for the last 30 days
const trends = await report.getSecurityScoreTrends(30);

trends.forEach(trend => {
  console.log(`${trend.baselineName}: ${trend.compliancePercentage}% compliant`);
});
```

### Generate Comprehensive Report

```typescript
const comprehensiveReport = await report.generateComprehensiveReport({
  includeDeviceDetails: true,
  includeSettingDetails: true,
  includeTrends: true,
  trendDays: 30
});

// Print formatted summary
SecurityBaselineReport.printSummary(comprehensiveReport);
```

### Export to Multiple Formats

```typescript
const comprehensiveReport = await report.generateComprehensiveReport();

const exportedFiles = await report.exportReport(comprehensiveReport, {
  outputDir: './reports',
  formats: ['json', 'csv', 'html']
});

console.log('Exported files:');
console.log(`JSON: ${exportedFiles.json}`);
console.log(`CSV: ${exportedFiles.csv}`);
console.log(`HTML: ${exportedFiles.html}`);
```

### Export Non-Compliant Settings to CSV

```typescript
const nonCompliantSettings = await report.getNonCompliantSettings();

await report.exportNonCompliantSettingsToCSV(
  nonCompliantSettings,
  './reports/non-compliant-settings.csv'
);
```

### Filter by Specific Template

```typescript
// Get specific template ID
const templates = await report.getAllSecurityBaselines();
const templateId = templates[0].id;

// Generate report for this template only
const filteredReport = await report.generateComprehensiveReport({
  templateId: templateId,
  includeDeviceDetails: true
});
```

### Filter by Specific Intent

```typescript
// Get specific intent ID
const intents = await report.getSecurityBaselineIntents();
const intentId = intents[0].id;

// Generate report for this intent only
const filteredReport = await report.generateComprehensiveReport({
  intentId: intentId,
  includeDeviceDetails: true,
  includeSettingDetails: true
});
```

## API Reference

### Class: SecurityBaselineReport

#### Methods

##### `execute(): Promise<ReportData>`

Executes the comprehensive security baseline report.

**Returns:** Complete report data including all baselines, compliance, and trends.

##### `getAllSecurityBaselines(): Promise<SecurityBaselineTemplate[]>`

Get all security baseline templates from Intune.

**Returns:** Array of security baseline templates.

**Throws:** Error if API call fails.

##### `getSecurityBaselineIntents(templateId?: string): Promise<SecurityBaselineIntent[]>`

Get all security baseline intents (applied baselines).

**Parameters:**
- `templateId` (optional): Filter by template ID

**Returns:** Array of security baseline intents.

**Throws:** Error if API call fails.

##### `getBaselineCompliancePerDevice(intentId: string): Promise<SecurityBaselineDeviceState[]>`

Get baseline compliance per device for a specific intent.

**Parameters:**
- `intentId`: The security baseline intent ID

**Returns:** Array of device states.

**Throws:** Error if API call fails or intentId is invalid.

##### `getBaselineDeviceStateSummary(intentId: string): Promise<SecurityBaselineStateSummary>`

Get device state summary for a specific intent.

**Parameters:**
- `intentId`: The security baseline intent ID

**Returns:** Summary of device states.

**Throws:** Error if API call fails or intentId is invalid.

##### `getBaselineSettingStates(intentId: string): Promise<SecurityBaselineSettingState[]>`

Get baseline setting states for a specific intent.

**Parameters:**
- `intentId`: The security baseline intent ID

**Returns:** Array of setting states.

**Throws:** Error if API call fails or intentId is invalid.

##### `getNonCompliantSettings(): Promise<NonCompliantSetting[]>`

Get all non-compliant settings across all security baselines.

**Returns:** Array of non-compliant settings with device and baseline information.

**Throws:** Error if API calls fail.

##### `getSecurityScoreTrends(days?: number): Promise<SecurityScoreTrendData[]>`

Get security score trends over time.

**Parameters:**
- `days` (optional): Number of days to analyze (default: 30)

**Returns:** Array of trend data points.

##### `generateComprehensiveReport(options?: SecurityBaselineReportOptions): Promise<SecurityBaselineComplianceReport>`

Generate a comprehensive security baseline report.

**Parameters:**
- `options` (optional): Report generation options

**Returns:** Complete security baseline compliance report.

##### `exportReport(report: SecurityBaselineComplianceReport, options?: SecurityBaselineReportOptions): Promise<{ [format: string]: string }>`

Export security baseline report to specified formats.

**Parameters:**
- `report`: The report to export
- `options` (optional): Export options including output directory and formats

**Returns:** Object with file paths for each exported format.

##### `exportNonCompliantSettingsToCSV(settings: NonCompliantSetting[], outputPath: string): Promise<void>`

Export non-compliant settings to CSV.

**Parameters:**
- `settings`: Non-compliant settings to export
- `outputPath`: Output file path

##### `exportTrendsToJSON(trends: SecurityScoreTrendData[], outputPath: string): Promise<void>`

Export security score trends to JSON.

**Parameters:**
- `trends`: Trend data to export
- `outputPath`: Output file path

##### `static printSummary(report: SecurityBaselineComplianceReport): void`

Print a formatted security baseline summary to console.

**Parameters:**
- `report`: Security baseline report to print

## Types

### SecurityBaselineTemplate

```typescript
interface SecurityBaselineTemplate {
  id: string;
  displayName: string;
  description?: string;
  version?: string;
  templateType?: string;
  templateSubtype?: string;
  publishedDateTime?: string;
  platformType?: string;
}
```

### SecurityBaselineIntent

```typescript
interface SecurityBaselineIntent {
  id: string;
  displayName: string;
  description?: string;
  templateId?: string;
  lastModifiedDateTime?: string;
  createdDateTime?: string;
  roleScopeTagIds?: string[];
  isAssigned?: boolean;
}
```

### SecurityBaselineDeviceState

```typescript
interface SecurityBaselineDeviceState {
  id: string;
  deviceId?: string;
  deviceDisplayName?: string;
  userPrincipalName?: string;
  state?: SecurityBaselineComplianceState;
  lastReportedDateTime?: string;
  userName?: string;
  deviceModel?: string;
  platform?: string;
  settingStatesCount?: number;
}
```

### SecurityBaselineStateSummary

```typescript
interface SecurityBaselineStateSummary {
  unknownCount?: number;
  notApplicableCount?: number;
  compliantCount?: number;
  remediatedCount?: number;
  nonCompliantCount?: number;
  errorCount?: number;
  conflictCount?: number;
  notAssignedCount?: number;
}
```

### NonCompliantSetting

```typescript
interface NonCompliantSetting {
  settingName: string;
  deviceName: string;
  userPrincipalName: string;
  currentValue: string;
  expectedValue: string;
  baselineName: string;
  errorCode?: string;
  lastReportedDateTime?: string;
}
```

### SecurityScoreTrendData

```typescript
interface SecurityScoreTrendData {
  date: string;
  totalDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  compliancePercentage: number;
  baselineName: string;
}
```

### SecurityBaselineReportOptions

```typescript
interface SecurityBaselineReportOptions {
  includeDeviceDetails?: boolean;
  includeSettingDetails?: boolean;
  templateId?: string;
  intentId?: string;
  includeTrends?: boolean;
  trendDays?: number;
  outputDir?: string;
  formats?: Array<'json' | 'csv' | 'html'>;
}
```

## Graph API Endpoints Used

The module uses the following Microsoft Graph API endpoints:

- `GET /deviceManagement/templates` - Get security baseline templates
- `GET /deviceManagement/intents` - Get security baseline intents
- `GET /deviceManagement/intents/{id}/deviceStates` - Get device states for an intent
- `GET /deviceManagement/intents/{id}/deviceStateSummary` - Get device state summary
- `GET /deviceManagement/intents/{id}/deviceSettingStateSummaries` - Get setting states

## Error Handling

The module includes comprehensive error handling:

- **Automatic Retry Logic**: Failed API calls are automatically retried up to 3 times with exponential backoff
- **Rate Limiting**: Handles 429 (Too Many Requests) errors with appropriate delays
- **Graceful Degradation**: Continues processing even if some intents fail
- **Detailed Error Messages**: Provides clear error messages for troubleshooting

## Pagination

The module automatically handles pagination for large datasets using the `@odata.nextLink` property from Graph API responses. All methods that return arrays automatically fetch all pages.

## Performance Considerations

- **Parallel Processing**: Where possible, the module makes parallel API calls to improve performance
- **Selective Data Fetching**: Use options to include only the data you need
- **Caching**: Consider implementing caching for frequently accessed data
- **Rate Limiting**: Be mindful of Microsoft Graph API rate limits

## Examples

See the examples directory for complete working examples:

```bash
# Run all examples (shows menu)
ts-node src/reports/examples/security-baseline-example.ts

# Run specific example
ts-node src/reports/examples/security-baseline-example.ts 1
```

Available examples:
1. Comprehensive Report
2. Get All Baselines
3. Get Device Compliance
4. Get Non-Compliant Settings
5. Get Security Trends
6. Generate and Export
7. Filter by Template
8. Export Non-Compliant CSV

## Troubleshooting

### Authentication Errors

**Error:** "Authentication failed"

**Solution:** Verify your Azure AD credentials in the `.env` file and ensure your application has the required permissions.

### Permission Errors

**Error:** "Forbidden" or "Access Denied"

**Solution:** Ensure your Azure AD application has been granted the required Microsoft Graph API permissions and admin consent has been provided.

### No Data Returned

**Issue:** No baselines or intents found

**Solution:**
- Verify that security baselines are configured in your Intune tenant
- Check that devices are enrolled and assigned to baselines
- Ensure your application has read permissions

### Rate Limiting

**Error:** "429 Too Many Requests"

**Solution:** The module automatically handles rate limiting with retry logic. If you continue to experience issues, consider adding delays between large batch operations.

## Best Practices

1. **Run Reports During Off-Peak Hours**: Generate large reports during off-peak hours to minimize impact
2. **Use Filtering**: When possible, filter by specific templates or intents to reduce API calls
3. **Export Regularly**: Schedule regular exports to track trends over time
4. **Monitor Non-Compliant Settings**: Focus on non-compliant settings for remediation efforts
5. **Archive Reports**: Keep historical reports to track security posture improvements

## Contributing

Contributions are welcome! Please see the main project CONTRIBUTING.md for guidelines.

## License

MIT License - See LICENSE file for details.

## Support

For issues, questions, or contributions, please open an issue on the project repository.

## Related Documentation

- [Microsoft Graph API - Security Baselines](https://docs.microsoft.com/en-us/graph/api/resources/intune-deviceintent-devicemanagementintent)
- [Microsoft Intune Security Baselines](https://docs.microsoft.com/en-us/mem/intune/protect/security-baselines)
- [Graph API Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference)

## Changelog

### Version 1.0.0
- Initial release
- Support for all security baseline operations
- Multiple export formats (JSON, CSV, HTML)
- Comprehensive error handling and retry logic
- Pagination support
- Security score trends (placeholder for historical data)
