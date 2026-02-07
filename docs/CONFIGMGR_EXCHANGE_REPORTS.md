# Configuration Manager Exchange ActiveSync Reports Guide

This guide provides comprehensive documentation for the Exchange ActiveSync Reports module, which replicates Configuration Manager Exchange ActiveSync reports for Microsoft Intune using the Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Azure Setup](#azure-setup)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Reports Overview](#reports-overview)
- [API Reference](#api-reference)
- [Usage Examples](#usage-examples)
- [TypeScript Interfaces](#typescript-interfaces)
- [Export Formats](#export-formats)
- [Error Handling](#error-handling)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Performance Considerations](#performance-considerations)

## Overview

The Exchange ActiveSync Reports module provides a comprehensive solution for monitoring and reporting on Exchange ActiveSync mobile devices in Microsoft Intune, replicating the functionality of Configuration Manager reports.

### Key Features

1. **Configuration Manager Report Compatibility**: Replicates ConfigMgr reports 8, 15, 21, and 34
2. **Type-Safe**: Full TypeScript support with comprehensive type definitions
3. **Reliable**: Automatic retry logic with exponential backoff
4. **Flexible**: Multiple filtering and configuration options
5. **Scalable**: Handles large device inventories with pagination
6. **Multi-Format Export**: JSON, CSV, and HTML output formats
7. **Production-Ready**: Comprehensive error handling and logging

### Replicated Configuration Manager Reports

This module implements the following Configuration Manager mobile device reports:

| Report # | Report Name | Description |
|----------|-------------|-------------|
| 8 | Compliance status of default ActiveSync mailbox policy | Shows compliance with Exchange ActiveSync policy |
| 15 | Inactive mobile devices (Exchange) | Devices not connected to Exchange Server for X days |
| 21 | Mobile device compliance details (Exchange) | Per-device compliance details |
| 34 | Settings summary for mobile devices (Exchange) | Device count by policy settings |

## Prerequisites

### Software Requirements

- Node.js 18.0.0 or higher
- npm or yarn package manager
- TypeScript 5.0 or higher

### Azure Requirements

- Azure AD (Entra ID) tenant
- Global Administrator or Intune Administrator access
- Microsoft Intune subscription with enrolled mobile devices
- Exchange ActiveSync integration configured

### Knowledge Requirements

- Basic understanding of TypeScript/JavaScript
- Familiarity with Microsoft Graph API
- Understanding of Exchange ActiveSync concepts
- Knowledge of Azure AD app registrations

## Azure Setup

### Step 1: Create Azure AD App Registration

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Configure the application:
   - **Name**: `Intune Exchange ActiveSync Reporter`
   - **Supported account types**: Accounts in this organizational directory only
   - **Redirect URI**: Leave blank
5. Click **Register**

### Step 2: Note Application Details

From the Overview page, copy:
- **Application (client) ID**
- **Directory (tenant) ID**

### Step 3: Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Add description: `Exchange ActiveSync Reporter Secret`
4. Select expiration: 12-24 months (recommended)
5. Click **Add**
6. **Copy the secret value immediately**

### Step 4: Configure API Permissions

Add the following Microsoft Graph Application permissions:

**Required Permissions:**
- `DeviceManagementManagedDevices.Read.All` - Read managed device properties
- `DeviceManagementConfiguration.Read.All` - Read device configuration
- `DeviceManagementApps.Read.All` - Read device compliance policies

**Optional Permissions:**
- `User.Read.All` - Read user profiles
- `Directory.Read.All` - Read directory data

### Step 5: Grant Admin Consent

1. Click **Grant admin consent for [Your Organization]**
2. Confirm the consent grant
3. Verify all permissions show "Granted"

## Installation

### 1. Clone the Repository

```bash
git clone <repository-url>
cd intunereporting
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Configure Environment Variables

Create a `.env` file:

```bash
cp .env.example .env
```

Add your Azure credentials:

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

### 4. Build the Project

```bash
npm run build
```

## Quick Start

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { ClientSecretCredential } from '@azure/identity';
import { ExchangeActiveSyncReports } from './src/reports/configmgr/exchange-activesync-reports';

// Create credentials
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID!,
  process.env.AZURE_CLIENT_ID!,
  process.env.AZURE_CLIENT_SECRET!
);

// Create auth provider
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

// Create Graph client
const graphClient = Client.initWithMiddleware({ authProvider });

// Create report instance
const config = { /* your config */ };
const report = new ExchangeActiveSyncReports(graphClient, config);

// Generate complete report
const reportData = await report.execute();
console.log(`Total Devices: ${reportData.summary.totalMobileDevices}`);
console.log(`Compliance: ${reportData.summary.compliancePercentage}%`);
```

### Run Example Scripts

```bash
# Run all examples
npm run examples:exchange-activesync

# Run specific example
ts-node src/reports/examples/exchange-activesync-examples.ts
```

## Reports Overview

### Report 8: Compliance Status of Default ActiveSync Mailbox Policy

Shows compliance status summary for ActiveSync mailbox policies.

**Key Metrics:**
- Total devices per policy
- Allowed vs. blocked vs. quarantined devices
- Compliance percentage
- Platform distribution
- Access state breakdown

**Use Cases:**
- Monitor policy compliance
- Identify blocking issues
- Track platform-specific compliance

```typescript
const policyCompliance = await report.getActiveSyncPolicyCompliance();

for (const policy of policyCompliance) {
  console.log(`Policy: ${policy.policyName}`);
  console.log(`  Compliance: ${policy.compliancePercentage}%`);
  console.log(`  Allowed: ${policy.allowedDevices}`);
  console.log(`  Blocked: ${policy.blockedDevices}`);
}
```

### Report 15: Inactive Mobile Devices (Exchange)

Identifies devices that haven't synced with Exchange Server within a specified timeframe.

**Key Metrics:**
- Days since last sync
- Exchange access state
- Compliance state
- User information
- Device details

**Use Cases:**
- Identify stale devices
- Clean up device inventory
- Detect sync issues
- User follow-up

```typescript
// Find devices inactive for 30+ days
const inactiveDevices = await report.getInactiveMobileDevices(30);

for (const device of inactiveDevices) {
  console.log(`${device.deviceName}: ${device.daysSinceLastSync} days`);
  console.log(`  User: ${device.userPrincipalName}`);
  console.log(`  Last Sync: ${device.lastSyncDateTime}`);
}
```

### Report 21: Mobile Device Compliance Details (Exchange)

Provides detailed compliance information for each mobile device.

**Key Metrics:**
- Exchange access state and reason
- Compliance policy status
- Individual setting compliance
- Device actions history
- EAS activation status

**Use Cases:**
- Troubleshoot compliance issues
- Review policy application
- Audit device settings
- Track device actions

```typescript
const deviceDetails = await report.getAllMobileDeviceComplianceDetails();

for (const device of deviceDetails) {
  console.log(`Device: ${device.deviceName}`);
  console.log(`  Access State: ${device.exchangeAccessState}`);
  console.log(`  Access Reason: ${device.exchangeAccessStateReason}`);
  console.log(`  Policies: ${device.compliancePolicies.length}`);

  for (const policy of device.compliancePolicies) {
    console.log(`    ${policy.policyName}: ${policy.complianceState}`);
  }
}
```

### Report 34: Settings Summary for Mobile Devices (Exchange)

Aggregates settings application across all mobile devices.

**Key Metrics:**
- Device count per setting
- Enabled vs. disabled vs. not configured
- Setting category (Exchange, Compliance, Configuration)
- Platform distribution
- Affected policies

**Use Cases:**
- Analyze setting adoption
- Identify misconfigured settings
- Policy coverage analysis
- Platform-specific settings

```typescript
const settingsSummary = await report.getMobileDeviceSettingsSummary();

for (const setting of settingsSummary) {
  console.log(`Setting: ${setting.settingName}`);
  console.log(`  Category: ${setting.settingCategory}`);
  console.log(`  Devices: ${setting.deviceCount}`);
  console.log(`  Enabled: ${setting.enabledCount} (${setting.percentage}%)`);
  console.log(`  Policies: ${setting.affectedPolicies.join(', ')}`);
}
```

## API Reference

### Class: ExchangeActiveSyncReports

Main class for generating Exchange ActiveSync reports.

#### Constructor

```typescript
constructor(graphClient: Client, config: AppConfig)
```

**Parameters:**
- `graphClient`: Microsoft Graph client instance
- `config`: Application configuration

#### Methods

##### execute()

Generates a complete Exchange ActiveSync report with all sections.

```typescript
async execute(): Promise<ReportData>
```

**Returns:** Complete report data including metadata, data, and summary

**Example:**
```typescript
const reportData = await report.execute();
```

##### getActiveSyncPolicyCompliance()

Gets compliance status for ActiveSync mailbox policies (Report 8).

```typescript
async getActiveSyncPolicyCompliance(): Promise<ActiveSyncPolicyComplianceSummary[]>
```

**Returns:** Array of policy compliance summaries

**Example:**
```typescript
const summaries = await report.getActiveSyncPolicyCompliance();
```

##### getInactiveMobileDevices()

Finds mobile devices that haven't synced within specified days (Report 15).

```typescript
async getInactiveMobileDevices(
  inactiveDaysThreshold: number = 30,
  options?: ExchangeActiveSyncFilterOptions
): Promise<InactiveMobileDevice[]>
```

**Parameters:**
- `inactiveDaysThreshold`: Days of inactivity (default: 30)
- `options`: Optional filter options

**Returns:** Array of inactive devices

**Example:**
```typescript
// Find devices inactive for 60+ days
const inactive = await report.getInactiveMobileDevices(60);

// With filters
const inactiveiOS = await report.getInactiveMobileDevices(30, {
  platform: 'iOS',
  exchangeAccessState: ExchangeAccessState.ALLOWED
});
```

##### getAllMobileDeviceComplianceDetails()

Gets detailed compliance information for all mobile devices (Report 21).

```typescript
async getAllMobileDeviceComplianceDetails(
  options?: ExchangeActiveSyncFilterOptions
): Promise<MobileDeviceComplianceDetails[]>
```

**Parameters:**
- `options`: Optional filter options

**Returns:** Array of device compliance details

**Example:**
```typescript
// Get all devices
const allDevices = await report.getAllMobileDeviceComplianceDetails();

// Filter by access state
const blockedDevices = await report.getAllMobileDeviceComplianceDetails({
  exchangeAccessState: ExchangeAccessState.BLOCKED
});

// Filter by platform and compliance
const nonCompliantAndroid = await report.getAllMobileDeviceComplianceDetails({
  platform: 'Android',
  complianceState: 'noncompliant'
});
```

##### getMobileDeviceSettingsSummary()

Gets aggregated settings summary across mobile devices (Report 34).

```typescript
async getMobileDeviceSettingsSummary(): Promise<MobileDeviceSettingsSummary[]>
```

**Returns:** Array of settings summaries

**Example:**
```typescript
const settings = await report.getMobileDeviceSettingsSummary();

// Find settings with low compliance
const lowCompliance = settings.filter(s =>
  s.percentage < 70 && s.deviceCount > 5
);
```

##### Export Methods

Export report data in various formats.

```typescript
async exportToJson(
  reportData: ReportData,
  outputDir: string,
  includeTimestamp?: boolean
): Promise<string>

async exportToCsv(
  reportData: ReportData,
  outputDir: string,
  includeTimestamp?: boolean
): Promise<string>

async exportToHtml(
  reportData: ReportData,
  outputDir: string,
  includeTimestamp?: boolean
): Promise<string>
```

**Example:**
```typescript
const reportData = await report.execute();

// Export to all formats
await report.exportToJson(reportData, './reports', true);
await report.exportToCsv(reportData, './reports', true);
await report.exportToHtml(reportData, './reports', true);
```

## TypeScript Interfaces

### ExchangeAccessState

Enumeration of Exchange access states.

```typescript
enum ExchangeAccessState {
  ALLOWED = 'allowed',
  BLOCKED = 'blocked',
  QUARANTINED = 'quarantined',
  UNKNOWN = 'unknown',
  NONE = 'none'
}
```

### ExchangeAccessStateReason

Enumeration of Exchange access state reasons.

```typescript
enum ExchangeAccessStateReason {
  NONE = 'none',
  UNKNOWN = 'unknown',
  EXCHANGE_GLOBAL_RULE = 'exchangeGlobalRule',
  EXCHANGE_INDIVIDUAL_RULE = 'exchangeIndividualRule',
  EXCHANGE_DEVICE_RULE = 'exchangeDeviceRule',
  EXCHANGE_UPGRADE = 'exchangeUpgrade',
  EXCHANGE_MAILBOX_POLICY = 'exchangeMailboxPolicy',
  OTHER = 'other',
  COMPLIANT = 'compliant',
  NOT_COMPLIANT = 'notCompliant',
  NOT_ENROLLED = 'notEnrolled',
  UNKNOWN_LOCATION = 'unknownLocation',
  MFA_REQUIRED = 'mfaRequired',
  AZURE_AD_BLOCK_DUE_TO_ACCESS_POLICY = 'azureADBlockDueToAccessPolicy',
  COMPROMISED_PASSWORD = 'compromisedPassword',
  DEVICE_NOT_KNOWN_WITH_MANAGED_APP = 'deviceNotKnownWithManagedApp'
}
```

### ActiveSyncPolicyComplianceSummary

Summary of ActiveSync mailbox policy compliance.

```typescript
interface ActiveSyncPolicyComplianceSummary {
  policyId: string;
  policyName: string;
  totalDevices: number;
  allowedDevices: number;
  blockedDevices: number;
  quarantinedDevices: number;
  unknownDevices: number;
  compliancePercentage: number;
  lastUpdated: Date;
  devicesByPlatform: Record<string, number>;
  devicesByAccessState: Record<ExchangeAccessState, number>;
}
```

### InactiveMobileDevice

Information about inactive mobile devices.

```typescript
interface InactiveMobileDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  emailAddress: string;
  operatingSystem: string;
  osVersion: string;
  model: string;
  manufacturer: string;
  lastSyncDateTime: Date;
  daysSinceLastSync: number;
  exchangeAccessState: ExchangeAccessState;
  exchangeAccessStateReason: ExchangeAccessStateReason;
  exchangeLastSuccessfulSyncDateTime?: Date;
  enrolledDateTime?: Date;
  complianceState: string;
  managementAgent: string;
  isSupervised?: boolean;
  serialNumber?: string;
  imei?: string;
}
```

### MobileDeviceComplianceDetails

Detailed compliance information for a mobile device.

```typescript
interface MobileDeviceComplianceDetails {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  emailAddress: string;
  operatingSystem: string;
  osVersion: string;
  model: string;
  manufacturer: string;
  exchangeAccessState: ExchangeAccessState;
  exchangeAccessStateReason: ExchangeAccessStateReason;
  exchangeLastSuccessfulSyncDateTime?: Date;
  complianceState: string;
  lastSyncDateTime: Date;
  enrolledDateTime?: Date;
  activeSyncId?: string;
  easDeviceId?: string;
  easActivated: boolean;
  easActivationDateTime?: Date;
  mailboxPolicies: MailboxPolicyAssignment[];
  compliancePolicies: CompliancePolicyStatus[];
  deviceActions: DeviceAction[];
  isManaged: boolean;
  isSupervised?: boolean;
  jailBroken?: string;
}
```

### MobileDeviceSettingsSummary

Summary of settings application across devices.

```typescript
interface MobileDeviceSettingsSummary {
  settingName: string;
  settingDescription: string;
  settingCategory: 'exchange' | 'compliance' | 'configuration';
  deviceCount: number;
  enabledCount: number;
  disabledCount: number;
  notConfiguredCount: number;
  percentage: number;
  affectedPolicies: string[];
  platforms: Record<string, number>;
}
```

### ExchangeActiveSyncFilterOptions

Filter options for queries.

```typescript
interface ExchangeActiveSyncFilterOptions {
  platform?: string;
  exchangeAccessState?: ExchangeAccessState;
  inactiveDaysThreshold?: number;
  userPrincipalName?: string;
  includeDetails?: boolean;
  complianceState?: string;
}
```

## Usage Examples

### Example 1: Monitor Policy Compliance

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);
const summaries = await report.getActiveSyncPolicyCompliance();

for (const summary of summaries) {
  if (summary.compliancePercentage < 90) {
    console.log(`⚠️ Low compliance: ${summary.policyName}`);
    console.log(`  Compliance: ${summary.compliancePercentage}%`);
    console.log(`  Blocked: ${summary.blockedDevices}`);
    console.log(`  Quarantined: ${summary.quarantinedDevices}`);
  }
}
```

### Example 2: Identify and Clean Up Inactive Devices

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);

// Find devices inactive for 90+ days
const veryInactive = await report.getInactiveMobileDevices(90);

console.log(`Found ${veryInactive.length} devices inactive for 90+ days`);

// Generate report for cleanup
const cleanupReport = veryInactive.map(device => ({
  deviceName: device.deviceName,
  user: device.userPrincipalName,
  lastSync: device.lastSyncDateTime,
  daysInactive: device.daysSinceLastSync,
  action: 'Recommend retirement'
}));

// Export for review
await report.exportToCsv(
  { metadata: {}, data: cleanupReport, summary: {} },
  './reports/cleanup',
  true
);
```

### Example 3: Troubleshoot Blocked Devices

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);

// Get all blocked devices
const blockedDevices = await report.getAllMobileDeviceComplianceDetails({
  exchangeAccessState: ExchangeAccessState.BLOCKED
});

console.log(`${blockedDevices.length} blocked devices found\n`);

// Analyze block reasons
const blockReasons = blockedDevices.reduce((acc, device) => {
  const reason = device.exchangeAccessStateReason;
  acc[reason] = (acc[reason] || 0) + 1;
  return acc;
}, {} as Record<string, number>);

console.log('Block Reasons:');
for (const [reason, count] of Object.entries(blockReasons)) {
  console.log(`  ${reason}: ${count}`);
}

// Review non-compliant policies
for (const device of blockedDevices) {
  console.log(`\n${device.deviceName} (${device.userPrincipalName})`);
  console.log(`  Reason: ${device.exchangeAccessStateReason}`);

  const nonCompliantPolicies = device.compliancePolicies.filter(
    p => p.complianceState.toLowerCase() !== 'compliant'
  );

  for (const policy of nonCompliantPolicies) {
    console.log(`  Policy: ${policy.policyName} - ${policy.complianceState}`);
    console.log(`    Non-compliant settings: ${policy.settingStates.length}`);
  }
}
```

### Example 4: Analyze Setting Compliance

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);
const settings = await report.getMobileDeviceSettingsSummary();

// Find settings with low compliance
const problemSettings = settings.filter(s =>
  s.percentage < 75 && s.deviceCount > 10
);

console.log(`Found ${problemSettings.length} settings with <75% compliance:\n`);

for (const setting of problemSettings) {
  console.log(`Setting: ${setting.settingName}`);
  console.log(`  Description: ${setting.settingDescription}`);
  console.log(`  Compliance: ${setting.percentage}%`);
  console.log(`  Enabled: ${setting.enabledCount}/${setting.deviceCount}`);
  console.log(`  Policies: ${setting.affectedPolicies.join(', ')}`);

  console.log('  Platform breakdown:');
  for (const [platform, count] of Object.entries(setting.platforms)) {
    console.log(`    ${platform}: ${count}`);
  }

  console.log('');
}
```

### Example 5: Platform-Specific Analysis

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);

// Analyze iOS devices
const iosDevices = await report.getAllMobileDeviceComplianceDetails({
  platform: 'iOS'
});

const iosStats = {
  total: iosDevices.length,
  compliant: iosDevices.filter(d => d.complianceState === 'compliant').length,
  easActivated: iosDevices.filter(d => d.easActivated).length,
  supervised: iosDevices.filter(d => d.isSupervised).length,
  allowed: iosDevices.filter(d => d.exchangeAccessState === ExchangeAccessState.ALLOWED).length
};

console.log('iOS Device Statistics:');
console.log(`  Total: ${iosStats.total}`);
console.log(`  Compliant: ${iosStats.compliant} (${Math.round((iosStats.compliant / iosStats.total) * 100)}%)`);
console.log(`  EAS Activated: ${iosStats.easActivated}`);
console.log(`  Supervised: ${iosStats.supervised}`);
console.log(`  Exchange Allowed: ${iosStats.allowed}`);

// Repeat for Android
const androidDevices = await report.getAllMobileDeviceComplianceDetails({
  platform: 'Android'
});

// ... similar analysis
```

### Example 6: Compliance Trend Monitoring

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);

// Get current compliance
const currentReport = await report.execute();
const currentCompliance = currentReport.summary.compliancePercentage;

// Store for trend tracking
const trendData = {
  date: new Date(),
  compliancePercentage: currentCompliance,
  totalDevices: currentReport.summary.totalMobileDevices,
  activeDevices: currentReport.summary.activeDevices,
  inactiveDevices: currentReport.summary.inactiveDevices,
  allowedDevices: currentReport.summary.allowedDevices,
  blockedDevices: currentReport.summary.blockedDevices
};

// Save to database or file for historical tracking
console.log('Compliance Snapshot:');
console.log(`  Date: ${trendData.date.toISOString()}`);
console.log(`  Compliance: ${trendData.compliancePercentage}%`);
console.log(`  Active: ${trendData.activeDevices}`);
console.log(`  Inactive: ${trendData.inactiveDevices}`);
console.log(`  Blocked: ${trendData.blockedDevices}`);
```

### Example 7: Export Multi-Format Reports

```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);
const reportData = await report.execute();

const timestamp = new Date().toISOString().split('T')[0];
const outputDir = `./reports/exchange-activesync/${timestamp}`;

// Create output directory
await fs.mkdir(outputDir, { recursive: true });

// Export in all formats
const jsonPath = await report.exportToJson(reportData, outputDir, true);
const csvPath = await report.exportToCsv(reportData, outputDir, true);
const htmlPath = await report.exportToHtml(reportData, outputDir, true);

console.log('Reports exported:');
console.log(`  JSON: ${jsonPath}`);
console.log(`  CSV: ${csvPath}`);
console.log(`  HTML: ${htmlPath}`);

// Generate individual report sections
const inactiveDevices = await report.getInactiveMobileDevices(30);
const inactiveReport = {
  metadata: {
    reportName: 'Inactive Devices',
    generatedAt: new Date().toISOString(),
    recordCount: inactiveDevices.length
  },
  data: inactiveDevices
};

await report.exportToCsv(inactiveReport, `${outputDir}/inactive`, true);
```

### Example 8: Scheduled Reporting

```typescript
import { CronJob } from 'cron';

// Schedule daily report at 8 AM
const job = new CronJob('0 8 * * *', async () => {
  try {
    const report = new ExchangeActiveSyncReports(graphClient, config);
    const reportData = await report.execute();

    const summary = reportData.summary as any;

    // Check for critical issues
    const criticalIssues = [];

    if (summary.compliancePercentage < 80) {
      criticalIssues.push(`Low compliance: ${summary.compliancePercentage}%`);
    }

    if (summary.blockedDevices > summary.totalMobileDevices * 0.1) {
      criticalIssues.push(`High blocked device count: ${summary.blockedDevices}`);
    }

    if (summary.inactiveDevices > summary.totalMobileDevices * 0.2) {
      criticalIssues.push(`High inactive device count: ${summary.inactiveDevices}`);
    }

    // Send notification if issues found
    if (criticalIssues.length > 0) {
      await sendAlertEmail({
        subject: 'Exchange ActiveSync Compliance Alert',
        body: `Critical issues detected:\n${criticalIssues.join('\n')}`,
        reportData
      });
    }

    // Export daily report
    const outputDir = './reports/daily';
    await report.exportToHtml(reportData, outputDir, true);

    console.log(`Daily report completed: ${new Date().toISOString()}`);
  } catch (error) {
    console.error('Scheduled report failed:', error);
  }
});

job.start();
```

## Export Formats

### JSON Export

Complete report data in JSON format.

**Features:**
- Full data structure preservation
- Easy programmatic access
- Machine-readable
- Compact or pretty-printed

**Example:**
```typescript
await report.exportToJson(reportData, './reports', true);
```

**Output:**
```json
{
  "metadata": {
    "reportName": "exchange-activesync-reports",
    "generatedAt": "2024-02-06T10:00:00Z",
    "recordCount": 150
  },
  "data": [...],
  "summary": {
    "totalMobileDevices": 150,
    "compliancePercentage": 92.5,
    ...
  }
}
```

### CSV Export

Tabular data suitable for Excel and data analysis.

**Features:**
- Excel-compatible
- Easy data analysis
- Pivot table support
- Simple filtering and sorting

**Example:**
```typescript
await report.exportToCsv(reportData, './reports', true);
```

**Output:**
```csv
type,deviceName,userPrincipalName,complianceState,exchangeAccessState
inactive-device,iPhone 14 Pro,user@contoso.com,compliant,allowed
inactive-device,Galaxy S23,user2@contoso.com,noncompliant,blocked
```

### HTML Export

Formatted HTML report with styling.

**Features:**
- Human-readable
- Formatted tables
- Sortable columns
- Print-friendly
- Email-compatible

**Example:**
```typescript
await report.exportToHtml(reportData, './reports', true);
```

## Error Handling

The module includes comprehensive error handling:

### Automatic Retry Logic

Failed API calls are automatically retried with exponential backoff.

```typescript
// Configured in BaseReport
protected async retryGraphCall<T>(
  apiCall: () => Promise<T>,
  maxRetries: number = 3,
  baseDelay: number = 1000
): Promise<T>
```

### Error Types

Common errors and handling:

1. **Authentication Errors**
   - Invalid credentials
   - Expired tokens
   - Insufficient permissions

2. **API Errors**
   - Rate limiting (429)
   - Server errors (500+)
   - Not found (404)

3. **Data Errors**
   - Missing required fields
   - Invalid data types
   - Unexpected responses

### Error Handling Example

```typescript
try {
  const report = new ExchangeActiveSyncReports(graphClient, config);
  const reportData = await report.execute();

  // Process report
} catch (error: any) {
  if (error.statusCode === 429) {
    console.error('Rate limited. Retry after:', error.retryAfter);
  } else if (error.statusCode === 401) {
    console.error('Authentication failed. Check credentials.');
  } else if (error.statusCode >= 500) {
    console.error('Server error. Retry later.');
  } else {
    console.error('Unexpected error:', error.message);
  }

  // Log for debugging
  logger.error('Report generation failed', {
    error: error.message,
    stack: error.stack,
    statusCode: error.statusCode
  });
}
```

## Best Practices

### 1. Authentication Management

```typescript
// Use environment variables for credentials
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID!,
  process.env.AZURE_CLIENT_ID!,
  process.env.AZURE_CLIENT_SECRET!
);

// Rotate secrets regularly
// Monitor expiration dates
// Use Key Vault for production
```

### 2. Performance Optimization

```typescript
// Batch API calls when possible
const [policies, devices, settings] = await Promise.all([
  report.getActiveSyncPolicyCompliance(),
  report.getAllMobileDeviceComplianceDetails(),
  report.getMobileDeviceSettingsSummary()
]);

// Use filters to reduce data volume
const recentDevices = await report.getInactiveMobileDevices(7, {
  platform: 'iOS',
  exchangeAccessState: ExchangeAccessState.ALLOWED
});

// Cache results for repeated queries
const cache = new Map();
const cacheKey = 'compliance-summary';
if (!cache.has(cacheKey)) {
  const summary = await report.getActiveSyncPolicyCompliance();
  cache.set(cacheKey, summary);
}
```

### 3. Data Validation

```typescript
// Validate report data
function validateReportData(data: any): boolean {
  if (!data || !data.metadata || !data.summary) {
    return false;
  }

  if (data.summary.totalMobileDevices < 0) {
    return false;
  }

  return true;
}

const reportData = await report.execute();
if (!validateReportData(reportData)) {
  throw new Error('Invalid report data');
}
```

### 4. Logging

```typescript
import { Logger } from './core/logger';

const logger = Logger.getInstance();

// Log important events
logger.info('Starting Exchange ActiveSync report generation');

// Log warnings
if (inactiveDevices.length > totalDevices * 0.3) {
  logger.warn('High inactive device count', {
    inactive: inactiveDevices.length,
    total: totalDevices,
    percentage: (inactiveDevices.length / totalDevices) * 100
  });
}

// Log errors with context
try {
  await report.execute();
} catch (error) {
  logger.error('Report generation failed', {
    error: error.message,
    timestamp: new Date().toISOString()
  });
}
```

### 5. Resource Management

```typescript
// Implement cleanup
class ExchangeActiveSyncReportRunner {
  private report: ExchangeActiveSyncReports;

  async run() {
    try {
      const reportData = await this.report.execute();
      await this.processReport(reportData);
    } finally {
      // Cleanup resources
      await this.cleanup();
    }
  }

  private async cleanup() {
    // Clear caches
    // Close connections
    // Release memory
  }
}
```

## Troubleshooting

### Common Issues

#### 1. Authentication Failures

**Symptom:** 401 Unauthorized errors

**Solutions:**
- Verify credentials in .env file
- Check app registration in Azure Portal
- Ensure admin consent is granted
- Verify API permissions are correct
- Check if client secret has expired

```bash
# Test authentication
ts-node -e "
  import { ClientSecretCredential } from '@azure/identity';
  const cred = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID!,
    process.env.AZURE_CLIENT_ID!,
    process.env.AZURE_CLIENT_SECRET!
  );
  await cred.getToken('https://graph.microsoft.com/.default');
  console.log('Authentication successful');
"
```

#### 2. No Devices Returned

**Symptom:** Empty device arrays

**Possible Causes:**
- No devices enrolled in Intune
- No Exchange ActiveSync devices
- Incorrect filters applied
- Missing permissions

**Solutions:**
```typescript
// Check total managed devices
const allDevices = await graphClient
  .api('/deviceManagement/managedDevices')
  .select('id,deviceName')
  .get();
console.log(`Total managed devices: ${allDevices.value.length}`);

// Check EAS-specific devices
const easDevices = await graphClient
  .api('/deviceManagement/managedDevices')
  .filter('easActivated eq true')
  .get();
console.log(`EAS activated devices: ${easDevices.value.length}`);
```

#### 3. Rate Limiting

**Symptom:** 429 Too Many Requests

**Solutions:**
- Implement exponential backoff (already included)
- Reduce concurrent requests
- Use batch endpoints when available
- Implement request throttling

```typescript
// Custom retry with longer delays
const reportWithRetry = async (maxAttempts = 5) => {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      return await report.execute();
    } catch (error: any) {
      if (error.statusCode === 429 && attempt < maxAttempts) {
        const delay = Math.pow(2, attempt) * 1000;
        console.log(`Rate limited, retrying in ${delay}ms...`);
        await new Promise(resolve => setTimeout(resolve, delay));
      } else {
        throw error;
      }
    }
  }
};
```

#### 4. Performance Issues

**Symptom:** Slow report generation

**Solutions:**
- Use filters to reduce data volume
- Implement pagination properly
- Cache frequently accessed data
- Run reports during off-peak hours

```typescript
// Optimize with filters
const filteredDevices = await report.getAllMobileDeviceComplianceDetails({
  platform: 'iOS',
  complianceState: 'noncompliant'
});

// Process in batches
const batchSize = 50;
for (let i = 0; i < devices.length; i += batchSize) {
  const batch = devices.slice(i, i + batchSize);
  await processBatch(batch);
}
```

#### 5. Incomplete Data

**Symptom:** Missing device properties

**Possible Causes:**
- Device not fully enrolled
- Properties not available for device type
- Sync pending
- Data not yet replicated

**Solutions:**
```typescript
// Check for missing data
const devicesWithMissingData = devices.filter(d =>
  !d.lastSyncDateTime || !d.exchangeAccessState
);

if (devicesWithMissingData.length > 0) {
  logger.warn(`${devicesWithMissingData.length} devices have incomplete data`);

  // Retry later or request sync
  for (const device of devicesWithMissingData) {
    logger.debug('Device with missing data', {
      deviceId: device.deviceId,
      deviceName: device.deviceName,
      missingFields: Object.keys(device).filter(k => !device[k])
    });
  }
}
```

## Performance Considerations

### API Rate Limits

Microsoft Graph API has the following limits:
- **Per app per tenant**: 10,000 requests per 10 minutes
- **Per user per tenant**: 2,000 requests per second

### Optimization Strategies

1. **Use Selective Queries**
```typescript
// Good: Select only needed properties
.select(['id', 'deviceName', 'complianceState'])

// Bad: Get all properties
.get()
```

2. **Implement Caching**
```typescript
class CachedExchangeActiveSyncReports extends ExchangeActiveSyncReports {
  private cache = new Map<string, { data: any; timestamp: number }>();
  private cacheDuration = 5 * 60 * 1000; // 5 minutes

  async getActiveSyncPolicyCompliance() {
    const cacheKey = 'policy-compliance';
    const cached = this.cache.get(cacheKey);

    if (cached && Date.now() - cached.timestamp < this.cacheDuration) {
      return cached.data;
    }

    const data = await super.getActiveSyncPolicyCompliance();
    this.cache.set(cacheKey, { data, timestamp: Date.now() });
    return data;
  }
}
```

3. **Batch Processing**
```typescript
// Process devices in parallel batches
const batchSize = 20;
const batches = [];

for (let i = 0; i < devices.length; i += batchSize) {
  batches.push(devices.slice(i, i + batchSize));
}

const results = await Promise.all(
  batches.map(batch => processBatch(batch))
);
```

4. **Progressive Loading**
```typescript
// Load data progressively
async function* loadDevices() {
  let nextLink: string | undefined;

  do {
    const response = await graphClient
      .api(nextLink || '/deviceManagement/managedDevices')
      .get();

    yield response.value;
    nextLink = response['@odata.nextLink'];
  } while (nextLink);
}

// Process as data arrives
for await (const deviceBatch of loadDevices()) {
  await processDevices(deviceBatch);
}
```

### Memory Management

```typescript
// Use streaming for large exports
import { createWriteStream } from 'fs';
import { pipeline } from 'stream/promises';

async function exportLargeReport(devices: any[]) {
  const writeStream = createWriteStream('report.csv');

  // Write header
  writeStream.write('deviceName,user,compliance\n');

  // Stream data
  for (const device of devices) {
    writeStream.write(
      `${device.deviceName},${device.userPrincipalName},${device.complianceState}\n`
    );
  }

  writeStream.end();
}
```

## Testing

### Run Unit Tests

```bash
# Run all tests
npm test

# Run Exchange ActiveSync tests only
npm test -- exchange-activesync-reports.test.ts

# Run with coverage
npm run test:coverage
```

### Test Structure

```typescript
describe('ExchangeActiveSyncReports', () => {
  describe('getActiveSyncPolicyCompliance', () => {
    it('should retrieve policy compliance status', async () => {
      const summaries = await report.getActiveSyncPolicyCompliance();
      expect(summaries).toBeDefined();
      expect(Array.isArray(summaries)).toBe(true);
    });
  });
});
```

## Support and Resources

### Microsoft Graph API Documentation
- [Managed Devices](https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [Device Compliance Policies](https://learn.microsoft.com/en-us/graph/api/resources/intune-deviceconfig-devicecompliancepolicy)
- [Exchange ActiveSync](https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-devicemanagement)

### Configuration Manager Reports Reference
- [Report 8: Compliance status of default ActiveSync mailbox policy](https://learn.microsoft.com/en-us/mem/configmgr/compliance/deploy-use/monitor-compliance-settings)
- [Report 15: Inactive mobile devices](https://learn.microsoft.com/en-us/mem/configmgr/mdm/deploy-use/monitor-mobile-devices)
- [Report 21: Mobile device compliance details](https://learn.microsoft.com/en-us/mem/configmgr/compliance/deploy-use/create-compliance-policy)
- [Report 34: Settings summary for mobile devices](https://learn.microsoft.com/en-us/mem/configmgr/compliance/deploy-use/monitor-compliance-settings)

### Community and Support
- GitHub Issues: Report bugs and request features
- Microsoft Tech Community: Exchange ActiveSync discussions
- Stack Overflow: `microsoft-graph` and `intune` tags

## License

[Your License Here]

## Contributing

Contributions are welcome! Please read our contributing guidelines before submitting pull requests.

---

**Last Updated:** 2024-02-06
**Version:** 1.0.0
