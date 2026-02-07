# Exchange ActiveSync Reports - Quick Reference

## Class: ExchangeActiveSyncReports

### Import
```typescript
import {
  ExchangeActiveSyncReports,
  ExchangeAccessState,
  ExchangeAccessStateReason,
  ActiveSyncPolicyComplianceSummary,
  InactiveMobileDevice,
  MobileDeviceComplianceDetails,
  MobileDeviceSettingsSummary
} from './src/reports/configmgr/exchange-activesync-reports';
```

---

## Main Methods

### 1. execute()
**Purpose**: Run all Exchange ActiveSync reports
```typescript
async execute(): Promise<ReportData>
```

**Returns**: Complete report data with all 4 reports combined

**Example**:
```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);
const data = await report.execute();
```

---

### 2. getActiveSyncPolicyCompliance()
**Purpose**: Report 8 - Compliance status of default ActiveSync mailbox policy
```typescript
async getActiveSyncPolicyCompliance(): Promise<ActiveSyncPolicyComplianceSummary[]>
```

**Returns**: Array of policy compliance summaries

**Example**:
```typescript
const summaries = await report.getActiveSyncPolicyCompliance();
summaries.forEach(summary => {
  console.log(`${summary.policyName}: ${summary.compliancePercentage}%`);
  console.log(`  Allowed: ${summary.allowedDevices}`);
  console.log(`  Blocked: ${summary.blockedDevices}`);
});
```

---

### 3. getInactiveMobileDevices()
**Purpose**: Report 15 - Inactive mobile devices
```typescript
async getInactiveMobileDevices(
  inactiveDaysThreshold: number = 30,
  options?: ExchangeActiveSyncFilterOptions
): Promise<InactiveMobileDevice[]>
```

**Parameters**:
- `inactiveDaysThreshold`: Number of days (default: 30)
- `options`: Optional filters (platform, exchangeAccessState, userPrincipalName)

**Returns**: Array of inactive devices

**Examples**:
```typescript
// Get devices inactive for 30+ days
const inactive = await report.getInactiveMobileDevices(30);

// Get inactive iOS devices only
const inactiveiOS = await report.getInactiveMobileDevices(30, {
  platform: 'iOS'
});

// Get inactive blocked devices
const inactiveBlocked = await report.getInactiveMobileDevices(30, {
  exchangeAccessState: ExchangeAccessState.BLOCKED
});
```

---

### 4. getAllMobileDeviceComplianceDetails()
**Purpose**: Report 21 - Mobile device compliance details
```typescript
async getAllMobileDeviceComplianceDetails(
  options?: ExchangeActiveSyncFilterOptions
): Promise<MobileDeviceComplianceDetails[]>
```

**Parameters**:
- `options`: Optional filters

**Returns**: Array of detailed device compliance information

**Example**:
```typescript
// Get all device details
const details = await report.getAllMobileDeviceComplianceDetails();

// Get non-compliant devices only
const nonCompliant = await report.getAllMobileDeviceComplianceDetails({
  complianceState: 'noncompliant'
});

details.forEach(device => {
  console.log(`${device.deviceName}:`);
  console.log(`  User: ${device.userPrincipalName}`);
  console.log(`  Compliance: ${device.complianceState}`);
  console.log(`  Policies: ${device.compliancePolicies.length}`);
});
```

---

### 5. getMobileDeviceSettingsSummary()
**Purpose**: Report 34 - Settings summary for mobile devices
```typescript
async getMobileDeviceSettingsSummary(): Promise<MobileDeviceSettingsSummary[]>
```

**Returns**: Array of setting summaries

**Example**:
```typescript
const settings = await report.getMobileDeviceSettingsSummary();
settings.forEach(setting => {
  console.log(`${setting.settingName}:`);
  console.log(`  Devices: ${setting.deviceCount}`);
  console.log(`  Enabled: ${setting.enabledCount} (${setting.percentage}%)`);
  console.log(`  Policies: ${setting.affectedPolicies.join(', ')}`);
});
```

---

## Export Methods

### exportToJson()
```typescript
async exportToJson(
  reportData: ReportData,
  outputDir: string,
  includeTimestamp: boolean = true
): Promise<string>
```

**Example**:
```typescript
const data = await report.execute();
const jsonPath = await report.exportToJson(data, './output');
console.log(`Saved to: ${jsonPath}`);
```

### exportToCsv()
```typescript
async exportToCsv(
  reportData: ReportData,
  outputDir: string,
  includeTimestamp: boolean = true
): Promise<string>
```

**Example**:
```typescript
const csvPath = await report.exportToCsv(data, './output');
```

### exportToHtml()
```typescript
async exportToHtml(
  reportData: ReportData,
  outputDir: string,
  includeTimestamp: boolean = true
): Promise<string>
```

**Example**:
```typescript
const htmlPath = await report.exportToHtml(data, './output');
```

---

## TypeScript Types

### Enums

#### ExchangeAccessState
```typescript
enum ExchangeAccessState {
  ALLOWED = 'allowed',
  BLOCKED = 'blocked',
  QUARANTINED = 'quarantined',
  UNKNOWN = 'unknown',
  NONE = 'none'
}
```

#### ExchangeAccessStateReason
```typescript
enum ExchangeAccessStateReason {
  NONE = 'none',
  UNKNOWN = 'unknown',
  EXCHANGE_GLOBAL_RULE = 'exchangeGlobalRule',
  EXCHANGE_INDIVIDUAL_RULE = 'exchangeIndividualRule',
  EXCHANGE_DEVICE_RULE = 'exchangeDeviceRule',
  EXCHANGE_UPGRADE = 'exchangeUpgrade',
  EXCHANGE_MAILBOX_POLICY = 'exchangeMailboxPolicy',
  COMPLIANT = 'compliant',
  NOT_COMPLIANT = 'notCompliant',
  NOT_ENROLLED = 'notEnrolled',
  // ... more values
}
```

---

### Interfaces

#### ActiveSyncPolicyComplianceSummary
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

#### InactiveMobileDevice
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

#### MobileDeviceComplianceDetails
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

#### MobileDeviceSettingsSummary
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

#### ExchangeActiveSyncFilterOptions
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

---

## Common Patterns

### Pattern 1: Get All Reports and Export
```typescript
const report = new ExchangeActiveSyncReports(graphClient, config);
const data = await report.execute();

// Export to all formats
await report.exportToJson(data, './reports');
await report.exportToCsv(data, './reports');
await report.exportToHtml(data, './reports');
```

### Pattern 2: Find Problem Devices
```typescript
// Get inactive devices
const inactive = await report.getInactiveMobileDevices(30);

// Get blocked devices
const blocked = await report.getAllMobileDeviceComplianceDetails({
  exchangeAccessState: ExchangeAccessState.BLOCKED
});

// Get non-compliant devices
const nonCompliant = await report.getAllMobileDeviceComplianceDetails({
  complianceState: 'noncompliant'
});
```

### Pattern 3: Policy Analysis
```typescript
const policies = await report.getActiveSyncPolicyCompliance();

// Find policies with low compliance
const lowCompliance = policies.filter(p =>
  p.compliancePercentage < 80 && p.totalDevices > 5
);

// Get most problematic policy
const worst = policies.reduce((prev, curr) =>
  prev.compliancePercentage < curr.compliancePercentage ? prev : curr
);
```

### Pattern 4: Platform-Specific Reports
```typescript
// iOS devices only
const iosInactive = await report.getInactiveMobileDevices(30, {
  platform: 'iOS'
});

const iosDetails = await report.getAllMobileDeviceComplianceDetails({
  platform: 'iOS'
});

// Android devices only
const androidInactive = await report.getInactiveMobileDevices(30, {
  platform: 'Android'
});
```

### Pattern 5: User-Specific Reports
```typescript
// Specific user
const userDevices = await report.getAllMobileDeviceComplianceDetails({
  userPrincipalName: 'user@contoso.com'
});

// Check if user has inactive devices
const userInactive = await report.getInactiveMobileDevices(30, {
  userPrincipalName: 'user@contoso.com'
});
```

---

## Error Handling

All methods use graceful degradation:

```typescript
try {
  const data = await report.execute();
  // Process data
} catch (error) {
  // Critical error - couldn't connect to Graph API
  console.error('Failed to generate report:', error);
}

// Methods return partial results on error
const inactive = await report.getInactiveMobileDevices(30);
// Returns empty array [] if error occurs
```

---

## Performance Tips

1. **Use Filtering**: Filter at the source to reduce data processing
   ```typescript
   // Good
   const iosDevices = await report.getInactiveMobileDevices(30, {
     platform: 'iOS'
   });

   // Less efficient
   const allDevices = await report.getInactiveMobileDevices(30);
   const iosDevices = allDevices.filter(d => d.operatingSystem === 'iOS');
   ```

2. **Batch Processing**: The module automatically processes devices in batches of 30

3. **Parallel Execution**: Use `Promise.all()` for independent reports
   ```typescript
   const [inactive, settings] = await Promise.all([
     report.getInactiveMobileDevices(30),
     report.getMobileDeviceSettingsSummary()
   ]);
   ```

4. **Export Once**: Generate report once, export to multiple formats
   ```typescript
   const data = await report.execute();
   await Promise.all([
     report.exportToJson(data, './output'),
     report.exportToCsv(data, './output'),
     report.exportToHtml(data, './output')
   ]);
   ```

---

## Graph API Permissions Required

```
Application Permissions:
- DeviceManagementManagedDevices.Read.All
- DeviceManagementConfiguration.Read.All
```

---

## Files

- **Implementation**: `src/reports/configmgr/exchange-activesync-reports.ts`
- **Tests**: `src/reports/configmgr/exchange-activesync-reports.test.ts`
- **Examples**: `src/reports/examples/exchange-activesync-examples.ts`
- **Docs**: `docs/CONFIGMGR_EXCHANGE_IMPL.md`

---

**Quick Start**: See `src/reports/examples/exchange-activesync-examples.ts` for working examples.
