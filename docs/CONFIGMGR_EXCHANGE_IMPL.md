# Exchange ActiveSync Reports - Implementation Guide

This document provides technical implementation details for the Exchange ActiveSync Reports module, including architecture, code structure, Graph API integration, and extension patterns.

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Module Structure](#module-structure)
- [Microsoft Graph API Integration](#microsoft-graph-api-integration)
- [Implementation Details](#implementation-details)
- [Data Flow](#data-flow)
- [Testing Strategy](#testing-strategy)
- [Performance Optimizations](#performance-optimizations)
- [Extending the Module](#extending-the-module)
- [Code Walkthrough](#code-walkthrough)

## Architecture Overview

### Design Principles

The Exchange ActiveSync Reports module follows these key design principles:

1. **Separation of Concerns**: Report generation logic is separated from data fetching and formatting
2. **Single Responsibility**: Each method handles one specific report type
3. **DRY (Don't Repeat Yourself)**: Common patterns are abstracted into helper methods
4. **Type Safety**: Comprehensive TypeScript types ensure compile-time safety
5. **Error Resilience**: Automatic retry logic and graceful degradation
6. **Extensibility**: Easy to add new reports or customize existing ones

### Component Hierarchy

```
BaseReport (Abstract Class)
    │
    ├── Graph Client Integration
    ├── Retry Logic & Pagination
    └── Metadata Creation
        │
        └── ExchangeActiveSyncReports (Concrete Implementation)
            │
            ├── Report 8: Policy Compliance Status
            ├── Report 15: Inactive Devices
            ├── Report 21: Device Compliance Details
            ├── Report 34: Settings Summary
            │
            ├── Helper Methods
            ├── Filtering Logic
            └── Export Methods
```

## Module Structure

### File Organization

```
src/reports/configmgr/
├── exchange-activesync-reports.ts      # Main implementation
├── exchange-activesync-reports.test.ts # Unit tests
│
src/reports/examples/
├── exchange-activesync-examples.ts     # Usage examples
│
docs/
├── CONFIGMGR_EXCHANGE_REPORTS.md       # User guide
└── CONFIGMGR_EXCHANGE_IMPL.md          # This file
```

### Class Structure

```typescript
export class ExchangeActiveSyncReports extends BaseReport {
  // Properties
  name: string
  description: string
  category: string
  enabled: boolean

  // Main execution method
  execute(): Promise<ReportData>

  // Report 8: Policy Compliance
  getActiveSyncPolicyCompliance(): Promise<ActiveSyncPolicyComplianceSummary[]>
  private getPolicyComplianceSummary(policy, devices): Promise<ActiveSyncPolicyComplianceSummary>
  private getUnassignedDevicesSummary(devices): ActiveSyncPolicyComplianceSummary

  // Report 15: Inactive Devices
  getInactiveMobileDevices(threshold, options?): Promise<InactiveMobileDevice[]>
  private mapToInactiveDevice(device, days): InactiveMobileDevice

  // Report 21: Device Compliance Details
  getAllMobileDeviceComplianceDetails(options?): Promise<MobileDeviceComplianceDetails[]>
  private getDeviceComplianceDetails(device): Promise<MobileDeviceComplianceDetails>
  private getDeviceCompliancePolicies(deviceId): Promise<CompliancePolicyStatus[]>
  private getDeviceActions(deviceId): Promise<DeviceAction[]>

  // Report 34: Settings Summary
  getMobileDeviceSettingsSummary(): Promise<MobileDeviceSettingsSummary[]>

  // Helper methods
  private mapExchangeAccessState(state): ExchangeAccessState
  private mapExchangeAccessStateReason(reason): ExchangeAccessStateReason
  private categorizeSettings(settingName): 'exchange' | 'compliance' | 'configuration'
  private getSettingDescription(settingName): string
  private matchesInactiveDeviceFilter(device, options): boolean
  private matchesComplianceDetailsFilter(device, options): boolean
  private generateExchangeActiveSyncSummary(data): ExchangeActiveSyncReportSummary

  // Export methods
  exportToJson(reportData, outputDir, includeTimestamp?): Promise<string>
  exportToCsv(reportData, outputDir, includeTimestamp?): Promise<string>
  exportToHtml(reportData, outputDir, includeTimestamp?): Promise<string>
}
```

## Microsoft Graph API Integration

### Required API Permissions

The module requires the following Microsoft Graph Application permissions:

```
DeviceManagementManagedDevices.Read.All
DeviceManagementConfiguration.Read.All
```

### Graph API Endpoints Used

#### 1. Managed Devices Endpoint

**Purpose**: Fetch all managed mobile devices with Exchange data

```typescript
GET /deviceManagement/managedDevices
```

**Query Parameters**:
- `$select`: Specific fields to retrieve
- `$filter`: Filter by management agent (eas, easMdm, mdm)
- `$top`: Limit results per page (999)

**Example Implementation**:

```typescript
const response = await this.graphClient
  .api('/deviceManagement/managedDevices')
  .select([
    'id',
    'deviceName',
    'userPrincipalName',
    'emailAddress',
    'operatingSystem',
    'osVersion',
    'model',
    'manufacturer',
    'lastSyncDateTime',
    'exchangeAccessState',
    'exchangeAccessStateReason',
    'exchangeLastSuccessfulSyncDateTime',
    'enrolledDateTime',
    'complianceState',
    'managementAgent',
    'isSupervised',
    'serialNumber',
    'imei',
    'easActivated',
    'easActivationDateTime',
    'activeSyncId',
    'easDeviceId',
    'jailBroken'
  ])
  .filter(`managementAgent eq 'eas' or managementAgent eq 'easMdm' or managementAgent eq 'mdm'`)
  .top(999)
  .get();
```

#### 2. Compliance Policies Endpoint

**Purpose**: Retrieve all device compliance policies

```typescript
GET /deviceManagement/deviceCompliancePolicies
```

**Example Implementation**:

```typescript
const policiesResponse = await this.graphClient
  .api('/deviceManagement/deviceCompliancePolicies')
  .select(['id', 'displayName'])
  .top(999)
  .get();
```

#### 3. Policy Device Statuses Endpoint

**Purpose**: Get device statuses for a specific compliance policy

```typescript
GET /deviceManagement/deviceCompliancePolicies/{id}/deviceStatuses
```

**Example Implementation**:

```typescript
const statusResponse = await this.graphClient
  .api(`/deviceManagement/deviceCompliancePolicies/${policy.id}/deviceStatuses`)
  .select(['id', 'status', 'platform'])
  .top(999)
  .get();
```

#### 4. Device Compliance Policy States Endpoint

**Purpose**: Get compliance policy states for a specific device

```typescript
GET /deviceManagement/managedDevices/{id}/deviceCompliancePolicyStates
```

**Example Implementation**:

```typescript
const response = await this.graphClient
  .api(`/deviceManagement/managedDevices/${deviceId}/deviceCompliancePolicyStates`)
  .select(['id', 'displayName', 'state', 'lastReportedDateTime', 'settingStates'])
  .expand('settingStates')
  .get();
```

### Pagination Handling

All Graph API calls use the inherited `getAllPages<T>()` method from `BaseReport`:

```typescript
protected async getAllPages<T>(initialResponse: any): Promise<T[]> {
  let results: T[] = initialResponse.value || [];
  let nextLink = initialResponse['@odata.nextLink'];

  while (nextLink) {
    try {
      const response = await this.graphClient.api(nextLink).get();
      results = results.concat(response.value || []);
      nextLink = response['@odata.nextLink'];
    } catch (error) {
      logger.error('Error fetching next page', error);
      break;
    }
  }

  return results;
}
```

### Retry Logic

All API calls are wrapped with retry logic using `retryGraphCall()`:

```typescript
protected async retryGraphCall<T>(
  apiCall: () => Promise<T>,
  maxRetries: number = 3,
  baseDelay: number = 1000
): Promise<T> {
  for (let attempt = 0; attempt < maxRetries; attempt++) {
    try {
      return await apiCall();
    } catch (error: any) {
      if (attempt === maxRetries - 1) throw error;

      const statusCode = error?.statusCode;
      if (statusCode === 429 || statusCode >= 500) {
        const delay = baseDelay * Math.pow(2, attempt);
        logger.warn(`Rate limited or server error, retrying in ${delay}ms`,
          { attempt, statusCode });
        await this.sleep(delay);
      } else {
        throw error;
      }
    }
  }
}
```

## Implementation Details

### Report 8: ActiveSync Policy Compliance Status

**Purpose**: Show compliance summary for each ActiveSync mailbox policy

**Implementation Flow**:

1. Fetch all managed devices with Exchange data
2. Filter devices by management agent (eas, easMdm)
3. Retrieve all compliance policies
4. For each policy:
   - Get device statuses
   - Match devices with statuses
   - Aggregate by Exchange access state
   - Calculate compliance percentage
5. Generate summary for unassigned devices
6. Return array of policy compliance summaries

**Key Code**:

```typescript
async getActiveSyncPolicyCompliance(): Promise<ActiveSyncPolicyComplianceSummary[]> {
  const summaries: ActiveSyncPolicyComplianceSummary[] = [];

  // Get all devices with Exchange data
  const response = await this.retryGraphCall(() =>
    this.graphClient
      .api('/deviceManagement/managedDevices')
      .select([...fields])
      .filter(`managementAgent eq 'eas' or managementAgent eq 'easMdm' or easActivated eq true`)
      .top(999)
      .get()
  );

  const devices = await this.getAllPages(response);

  // Get compliance policies
  const policiesResponse = await this.retryGraphCall(() =>
    this.graphClient
      .api('/deviceManagement/deviceCompliancePolicies')
      .select(['id', 'displayName'])
      .top(999)
      .get()
  );

  const policies = await this.getAllPages(policiesResponse);

  // Process each policy
  for (const policy of policies) {
    const summary = await this.getPolicyComplianceSummary(policy, devices);
    summaries.push(summary);
  }

  // Add unassigned devices summary
  const unassignedSummary = this.getUnassignedDevicesSummary(devices);
  if (unassignedSummary.totalDevices > 0) {
    summaries.push(unassignedSummary);
  }

  return summaries;
}
```

### Report 15: Inactive Mobile Devices

**Purpose**: Identify devices that haven't synced within a specified threshold

**Implementation Flow**:

1. Fetch all managed mobile devices
2. Calculate days since last sync for each device
3. Filter devices exceeding the inactive threshold
4. Apply additional filters (platform, access state, user)
5. Sort by days inactive (descending)
6. Return inactive devices list

**Key Algorithm**:

```typescript
async getInactiveMobileDevices(
  inactiveDaysThreshold: number = 30,
  options?: ExchangeActiveSyncFilterOptions
): Promise<InactiveMobileDevice[]> {
  const inactiveDevices: InactiveMobileDevice[] = [];
  const devices = await this.getAllPages(response);

  const now = Date.now();
  const thresholdMs = inactiveDaysThreshold * 24 * 60 * 60 * 1000;

  for (const device of devices) {
    const lastSyncDate = device.lastSyncDateTime
      ? new Date(device.lastSyncDateTime)
      : device.exchangeLastSuccessfulSyncDateTime
        ? new Date(device.exchangeLastSuccessfulSyncDateTime)
        : null;

    if (!lastSyncDate) {
      // Device has never synced - check enrollment date
      const daysSinceEnrollment = device.enrolledDateTime
        ? Math.floor((now - new Date(device.enrolledDateTime).getTime()) / (1000 * 60 * 60 * 24))
        : 9999;

      if (daysSinceEnrollment >= inactiveDaysThreshold) {
        inactiveDevices.push(this.mapToInactiveDevice(device, daysSinceEnrollment));
      }
    } else {
      const msSinceLastSync = now - lastSyncDate.getTime();
      const daysSinceLastSync = Math.floor(msSinceLastSync / (1000 * 60 * 60 * 24));

      if (msSinceLastSync >= thresholdMs) {
        if (this.matchesInactiveDeviceFilter(device, options)) {
          inactiveDevices.push(this.mapToInactiveDevice(device, daysSinceLastSync));
        }
      }
    }
  }

  // Sort by days inactive (descending)
  inactiveDevices.sort((a, b) => b.daysSinceLastSync - a.daysSinceLastSync);

  return inactiveDevices;
}
```

### Report 21: Mobile Device Compliance Details

**Purpose**: Provide detailed compliance information for each mobile device

**Implementation Flow**:

1. Fetch all managed mobile devices
2. Process devices in batches (30 devices per batch)
3. For each device:
   - Get compliance policies and their states
   - Get device action history
   - Get mailbox policy assignments (if available)
   - Combine into comprehensive details object
4. Apply filters if provided
5. Return detailed device compliance array

**Batch Processing Pattern**:

```typescript
async getAllMobileDeviceComplianceDetails(
  options?: ExchangeActiveSyncFilterOptions
): Promise<MobileDeviceComplianceDetails[]> {
  const deviceDetails: MobileDeviceComplianceDetails[] = [];
  const devices = await this.getAllPages(response);

  const batchSize = 30;
  for (let i = 0; i < devices.length; i += batchSize) {
    const batch = devices.slice(i, Math.min(i + batchSize, devices.length));

    const batchPromises = batch.map(async (device) => {
      try {
        const details = await this.getDeviceComplianceDetails(device);
        if (this.matchesComplianceDetailsFilter(details, options)) {
          return details;
        }
      } catch (error) {
        logger.warn(`Failed to fetch compliance details for device ${device.id}`, error);
      }
      return null;
    });

    const batchResults = await Promise.all(batchPromises);
    deviceDetails.push(...batchResults.filter(d => d !== null));
  }

  return deviceDetails;
}
```

### Report 34: Settings Summary for Mobile Devices

**Purpose**: Aggregate setting configurations across all mobile devices

**Implementation Flow**:

1. Get all compliance policies
2. For each policy:
   - Get device statuses
   - Sample devices to retrieve settings
   - Extract setting states
3. Aggregate settings across all policies:
   - Track device count per setting
   - Count enabled/disabled/not configured
   - Track platforms using each setting
   - Track policies containing each setting
4. Calculate percentages
5. Sort by device count (descending)
6. Return settings summary array

**Settings Aggregation Pattern**:

```typescript
async getMobileDeviceSettingsSummary(): Promise<MobileDeviceSettingsSummary[]> {
  const summaries: MobileDeviceSettingsSummary[] = [];
  const settingsMap = new Map<string, {
    description: string;
    category: 'exchange' | 'compliance' | 'configuration';
    deviceCount: number;
    enabledCount: number;
    disabledCount: number;
    notConfiguredCount: number;
    policies: Set<string>;
    platforms: Record<string, number>;
  }>();

  // Process each policy and aggregate settings
  for (const policy of policies) {
    const deviceStatuses = await this.getAllPages(statusResponse);
    const sampleSize = Math.min(10, statuses.length);

    for (let i = 0; i < sampleSize; i++) {
      const policyStates = await this.getDevicePolicyStates(deviceId, policy.id);

      for (const setting of policyStates.settingStates) {
        const settingName = setting.setting || setting.settingName;

        if (!settingsMap.has(settingName)) {
          settingsMap.set(settingName, {
            description: this.getSettingDescription(settingName),
            category: this.categorizeSettings(settingName),
            deviceCount: 0,
            enabledCount: 0,
            disabledCount: 0,
            notConfiguredCount: 0,
            policies: new Set(),
            platforms: {}
          });
        }

        const settingData = settingsMap.get(settingName)!;
        settingData.deviceCount++;
        settingData.policies.add(policy.displayName);

        // Update state counts based on setting state
        this.updateSettingStateCounts(settingData, setting.state);
      }
    }
  }

  // Convert map to array and calculate percentages
  for (const [settingName, data] of settingsMap.entries()) {
    summaries.push({
      settingName,
      settingDescription: data.description,
      settingCategory: data.category,
      deviceCount: data.deviceCount,
      enabledCount: data.enabledCount,
      disabledCount: data.disabledCount,
      notConfiguredCount: data.notConfiguredCount,
      percentage: (data.enabledCount / data.deviceCount) * 100,
      affectedPolicies: Array.from(data.policies),
      platforms: data.platforms
    });
  }

  return summaries.sort((a, b) => b.deviceCount - a.deviceCount);
}
```

## Data Flow

### Complete Report Execution Flow

```
execute()
    │
    ├─> getActiveSyncPolicyCompliance()
    │       │
    │       ├─> Fetch managed devices (EAS only)
    │       ├─> Fetch compliance policies
    │       ├─> For each policy:
    │       │       ├─> Get device statuses
    │       │       └─> Calculate compliance metrics
    │       └─> Generate unassigned summary
    │
    ├─> getInactiveMobileDevices(30)
    │       │
    │       ├─> Fetch all managed mobile devices
    │       ├─> Calculate days since last sync
    │       ├─> Filter by threshold
    │       └─> Sort by days inactive
    │
    ├─> getAllMobileDeviceComplianceDetails()
    │       │
    │       ├─> Fetch all managed mobile devices
    │       ├─> Process in batches (30 devices)
    │       └─> For each device:
    │               ├─> Get compliance policies
    │               ├─> Get device actions
    │               └─> Combine into details object
    │
    ├─> getMobileDeviceSettingsSummary()
    │       │
    │       ├─> Fetch compliance policies
    │       ├─> For each policy:
    │       │       ├─> Get device statuses
    │       │       └─> Sample devices for settings
    │       └─> Aggregate settings data
    │
    └─> generateExchangeActiveSyncSummary()
            │
            └─> Combine all data into summary object
```

### Type Mapping Flow

```
Graph API Response (any)
    │
    ├─> mapExchangeAccessState()
    │       │
    │       └─> ExchangeAccessState enum
    │
    ├─> mapExchangeAccessStateReason()
    │       │
    │       └─> ExchangeAccessStateReason enum
    │
    ├─> mapToInactiveDevice()
    │       │
    │       └─> InactiveMobileDevice interface
    │
    └─> categorizeSettings()
            │
            └─> 'exchange' | 'compliance' | 'configuration'
```

## Testing Strategy

### Unit Tests Structure

The test suite (`exchange-activesync-reports.test.ts`) follows this structure:

```typescript
describe('ExchangeActiveSyncReports', () => {
  // Setup and teardown
  beforeEach(() => {
    // Initialize mocks
  });

  afterEach(() => {
    // Cleanup
  });

  describe('getActiveSyncPolicyCompliance', () => {
    it('should fetch policy compliance summaries');
    it('should handle devices without policies');
    it('should calculate compliance percentages correctly');
    it('should aggregate by platform');
    it('should handle errors gracefully');
  });

  describe('getInactiveMobileDevices', () => {
    it('should identify inactive devices');
    it('should respect threshold parameter');
    it('should handle devices that never synced');
    it('should apply filters correctly');
    it('should sort by days inactive');
  });

  describe('getAllMobileDeviceComplianceDetails', () => {
    it('should fetch device compliance details');
    it('should process devices in batches');
    it('should include policy states');
    it('should include device actions');
    it('should apply filters');
  });

  describe('getMobileDeviceSettingsSummary', () => {
    it('should aggregate settings across policies');
    it('should categorize settings correctly');
    it('should calculate percentages');
    it('should track affected policies');
  });

  describe('Helper Methods', () => {
    it('should map exchange access states correctly');
    it('should map access state reasons correctly');
    it('should categorize settings');
    it('should match filters');
  });
});
```

### Mocking Strategy

```typescript
// Mock Graph Client
const mockGraphClient = {
  api: jest.fn().mockReturnThis(),
  select: jest.fn().mockReturnThis(),
  filter: jest.fn().mockReturnThis(),
  top: jest.fn().mockReturnThis(),
  expand: jest.fn().mockReturnThis(),
  get: jest.fn()
};

// Mock Responses
const mockDevicesResponse = {
  value: [mockDevice1, mockDevice2, mockDevice3],
  '@odata.nextLink': null
};

const mockPoliciesResponse = {
  value: [mockPolicy1, mockPolicy2],
  '@odata.nextLink': null
};

// Setup mock behavior
beforeEach(() => {
  mockGraphClient.get.mockImplementation((endpoint) => {
    if (endpoint.includes('managedDevices')) {
      return Promise.resolve(mockDevicesResponse);
    }
    if (endpoint.includes('deviceCompliancePolicies')) {
      return Promise.resolve(mockPoliciesResponse);
    }
    return Promise.resolve({ value: [] });
  });
});
```

### Test Coverage Goals

- **Line Coverage**: > 90%
- **Branch Coverage**: > 85%
- **Function Coverage**: 100%
- **Statement Coverage**: > 90%

## Performance Optimizations

### 1. Batch Processing

Process devices in batches to prevent memory issues and rate limiting:

```typescript
const batchSize = 30;
for (let i = 0; i < devices.length; i += batchSize) {
  const batch = devices.slice(i, Math.min(i + batchSize, devices.length));
  const batchPromises = batch.map(async (device) => {
    // Process device
  });
  const batchResults = await Promise.all(batchPromises);
}
```

### 2. Parallel Execution

Execute independent operations in parallel:

```typescript
const [
  policyCompliance,
  inactiveDevices,
  deviceDetails,
  settingsSummary
] = await Promise.all([
  this.getActiveSyncPolicyCompliance(),
  this.getInactiveMobileDevices(30),
  this.getAllMobileDeviceComplianceDetails(),
  this.getMobileDeviceSettingsSummary()
]);
```

### 3. Field Selection

Only request necessary fields to reduce payload size:

```typescript
.select([
  'id',
  'deviceName',
  'userPrincipalName',
  // Only include required fields
])
```

### 4. Sampling for Settings

Sample a subset of devices for settings analysis:

```typescript
const sampleSize = Math.min(10, statuses.length);
for (let i = 0; i < sampleSize; i++) {
  // Process sample device
}
```

### 5. Caching Opportunities

Consider caching for frequently accessed data:

```typescript
// Cache compliance policies (valid for 1 hour)
private cachedPolicies: { data: any[], timestamp: number } | null = null;

async getCachedPolicies() {
  const now = Date.now();
  const cacheTimeout = 60 * 60 * 1000; // 1 hour

  if (this.cachedPolicies && (now - this.cachedPolicies.timestamp) < cacheTimeout) {
    return this.cachedPolicies.data;
  }

  const policies = await this.fetchPolicies();
  this.cachedPolicies = { data: policies, timestamp: now };
  return policies;
}
```

## Extending the Module

### Adding a New Report

To add a new Exchange ActiveSync report:

1. **Define Types**:

```typescript
export interface NewReportData {
  // Define your data structure
  reportField1: string;
  reportField2: number;
  // ...
}
```

2. **Add Method**:

```typescript
async getNewReport(parameters?: any): Promise<NewReportData[]> {
  logger.info('Fetching new report data');
  const results: NewReportData[] = [];

  try {
    // Fetch data from Graph API
    const response = await this.retryGraphCall(() =>
      this.graphClient
        .api('/deviceManagement/...')
        .select([...])
        .get()
    );

    const data = await this.getAllPages(response);

    // Process data
    for (const item of data) {
      results.push({
        reportField1: item.field1,
        reportField2: item.field2
      });
    }

    logger.info(`Generated ${results.length} report items`);
    return results;
  } catch (error) {
    logger.error('Error generating new report', error);
    return results;
  }
}
```

3. **Update Execute Method**:

```typescript
async execute(): Promise<ReportData> {
  const [
    // existing reports...
    newReport
  ] = await Promise.all([
    // existing calls...
    this.getNewReport()
  ]);

  // Add to data array
  const data = [
    // existing data...
    ...newReport.map(r => ({ type: 'new-report', ...r }))
  ];

  return { metadata, data, summary };
}
```

4. **Add Tests**:

```typescript
describe('getNewReport', () => {
  it('should fetch new report data', async () => {
    // Test implementation
  });
});
```

### Adding Custom Filters

To add custom filtering capabilities:

1. **Extend Filter Interface**:

```typescript
export interface ExchangeActiveSyncFilterOptions {
  // Existing filters...
  customFilter?: string;
  dateRange?: { start: Date; end: Date };
}
```

2. **Implement Filter Logic**:

```typescript
private matchesCustomFilter(device: any, options?: ExchangeActiveSyncFilterOptions): boolean {
  if (!options) return true;

  if (options.customFilter) {
    // Implement custom filter logic
    if (!device.customField?.includes(options.customFilter)) {
      return false;
    }
  }

  if (options.dateRange) {
    const deviceDate = new Date(device.dateField);
    if (deviceDate < options.dateRange.start || deviceDate > options.dateRange.end) {
      return false;
    }
  }

  return true;
}
```

### Adding Custom Transformations

To add custom data transformations:

```typescript
private transformDeviceData(device: any): CustomDeviceFormat {
  return {
    // Standard fields
    id: device.id,
    name: device.deviceName,

    // Custom calculations
    riskScore: this.calculateRiskScore(device),
    category: this.categorizeDevice(device),

    // Custom formatting
    displayName: `${device.deviceName} (${device.operatingSystem})`,
    formattedDate: this.formatDate(device.lastSyncDateTime)
  };
}

private calculateRiskScore(device: any): number {
  let score = 0;

  if (device.exchangeAccessState === 'blocked') score += 50;
  if (device.jailBroken === 'true') score += 30;
  if (!device.easActivated) score += 20;

  return score;
}
```

## Code Walkthrough

### Key Method: getActiveSyncPolicyCompliance()

Let's walk through the complete flow of the policy compliance method:

```typescript
async getActiveSyncPolicyCompliance(): Promise<ActiveSyncPolicyComplianceSummary[]> {
  logger.info('Fetching ActiveSync policy compliance status');
  const summaries: ActiveSyncPolicyComplianceSummary[] = [];

  try {
    // STEP 1: Fetch all devices with Exchange ActiveSync enabled
    // Filter by management agent to only get EAS-enabled devices
    const response = await this.retryGraphCall(() =>
      this.graphClient
        .api('/deviceManagement/managedDevices')
        .select([
          'id',
          'deviceName',
          'operatingSystem',
          'exchangeAccessState',
          'exchangeAccessStateReason',
          'easActivated',
          'complianceState'
        ])
        // Filter for Exchange ActiveSync devices
        .filter(`managementAgent eq 'eas' or managementAgent eq 'easMdm' or easActivated eq true`)
        .top(999)
        .get()
    );

    // STEP 2: Handle pagination to get all devices
    const devices = await this.getAllPages(response);
    logger.info(`Processing ${devices.length} Exchange ActiveSync devices`);

    // STEP 3: Fetch all compliance policies
    const policiesResponse = await this.retryGraphCall(() =>
      this.graphClient
        .api('/deviceManagement/deviceCompliancePolicies')
        .select(['id', 'displayName'])
        .top(999)
        .get()
    );

    const policies = await this.getAllPages(policiesResponse);

    // STEP 4: Process each policy to generate compliance summary
    for (const policy of policies) {
      try {
        // Get detailed summary for this policy
        const summary = await this.getPolicyComplianceSummary(policy, devices);
        summaries.push(summary);
      } catch (error) {
        logger.warn(`Failed to get compliance summary for policy ${policy.id}`, error);
        // Continue processing other policies even if one fails
      }
    }

    // STEP 5: Generate summary for devices without explicit policy assignment
    const unassignedSummary = this.getUnassignedDevicesSummary(devices);
    if (unassignedSummary.totalDevices > 0) {
      summaries.push(unassignedSummary);
    }

    logger.info(`Generated ${summaries.length} policy compliance summaries`);
    return summaries;

  } catch (error) {
    logger.error('Error fetching ActiveSync policy compliance', error);
    return summaries; // Return partial results instead of throwing
  }
}
```

**Key Design Decisions**:

1. **Graceful Degradation**: Returns partial results on error instead of throwing
2. **Retry Logic**: Uses `retryGraphCall()` for resilience
3. **Pagination**: Uses `getAllPages()` to handle large datasets
4. **Comprehensive Logging**: Logs at each major step for observability
5. **Include Unassigned**: Captures devices without explicit policy assignments

### Key Method: getPolicyComplianceSummary()

```typescript
private async getPolicyComplianceSummary(
  policy: any,
  allDevices: any[]
): Promise<ActiveSyncPolicyComplianceSummary> {
  // STEP 1: Get device statuses for this specific policy
  const statusResponse = await this.retryGraphCall(() =>
    this.graphClient
      .api(`/deviceManagement/deviceCompliancePolicies/${policy.id}/deviceStatuses`)
      .select(['id', 'status', 'platform'])
      .top(999)
      .get()
  );

  const deviceStatuses = await this.getAllPages(statusResponse);

  // STEP 2: Initialize aggregation structures
  const devicesByAccessState: Record<ExchangeAccessState, number> = {
    [ExchangeAccessState.ALLOWED]: 0,
    [ExchangeAccessState.BLOCKED]: 0,
    [ExchangeAccessState.QUARANTINED]: 0,
    [ExchangeAccessState.UNKNOWN]: 0,
    [ExchangeAccessState.NONE]: 0
  };

  const devicesByPlatform: Record<string, number> = {};

  // STEP 3: Match devices with policy statuses
  // Only count devices that are both in our EAS device list AND have this policy
  const matchedDevices = allDevices.filter(device =>
    deviceStatuses.some(status => status.id === device.id)
  );

  // STEP 4: Aggregate statistics
  for (const device of matchedDevices) {
    // Map string to enum for type safety
    const accessState = this.mapExchangeAccessState(device.exchangeAccessState);
    devicesByAccessState[accessState]++;

    // Track by platform
    const platform = device.operatingSystem || 'Unknown';
    devicesByPlatform[platform] = (devicesByPlatform[platform] || 0) + 1;
  }

  // STEP 5: Calculate metrics
  const totalDevices = matchedDevices.length;
  const allowedDevices = devicesByAccessState[ExchangeAccessState.ALLOWED];
  const compliancePercentage = totalDevices > 0
    ? Math.round((allowedDevices / totalDevices) * 100 * 100) / 100
    : 0;

  // STEP 6: Return structured summary
  return {
    policyId: policy.id,
    policyName: policy.displayName || 'Unknown',
    totalDevices,
    allowedDevices,
    blockedDevices: devicesByAccessState[ExchangeAccessState.BLOCKED],
    quarantinedDevices: devicesByAccessState[ExchangeAccessState.QUARANTINED],
    unknownDevices: devicesByAccessState[ExchangeAccessState.UNKNOWN] +
                    devicesByAccessState[ExchangeAccessState.NONE],
    compliancePercentage,
    lastUpdated: new Date(),
    devicesByPlatform,
    devicesByAccessState
  };
}
```

**Key Techniques**:

1. **Device Matching**: Joins devices with policy statuses by ID
2. **Type-Safe Aggregation**: Uses Record types with enums for counts
3. **Percentage Calculation**: Rounds to 2 decimal places
4. **Unknown Handling**: Combines UNKNOWN and NONE states
5. **Default Values**: Handles missing data gracefully

### Helper Method: mapExchangeAccessState()

```typescript
private mapExchangeAccessState(state: string): ExchangeAccessState {
  if (!state) return ExchangeAccessState.NONE;

  const stateLower = state.toLowerCase();

  // Pattern matching for fuzzy string matching
  if (stateLower.includes('allow')) return ExchangeAccessState.ALLOWED;
  if (stateLower.includes('block')) return ExchangeAccessState.BLOCKED;
  if (stateLower.includes('quarantine')) return ExchangeAccessState.QUARANTINED;
  if (stateLower.includes('unknown')) return ExchangeAccessState.UNKNOWN;

  return ExchangeAccessState.NONE;
}
```

**Design Rationale**:

1. **Defensive**: Handles null/undefined input
2. **Fuzzy Matching**: Uses `includes()` instead of exact match for resilience
3. **Case Insensitive**: Normalizes to lowercase
4. **Default Fallback**: Returns NONE for unrecognized states

## Best Practices

### 1. Error Handling

Always use try-catch blocks and log errors:

```typescript
try {
  const result = await this.fetchData();
  return result;
} catch (error) {
  logger.error('Descriptive error message', error);
  return []; // Return empty array instead of throwing
}
```

### 2. Logging

Use appropriate log levels:

```typescript
logger.info('High-level operation started');   // Info
logger.debug('Detailed processing info');      // Debug
logger.warn('Non-critical issue occurred');    // Warning
logger.error('Critical error occurred', error); // Error
```

### 3. Type Safety

Always define interfaces for data structures:

```typescript
// Good: Type-safe
interface DeviceData {
  id: string;
  name: string;
}

function processDevice(device: DeviceData) {
  // Compiler ensures device has required properties
}

// Bad: No type safety
function processDevice(device: any) {
  // No compile-time checks
}
```

### 4. Null Safety

Handle null/undefined values:

```typescript
// Good: Safe handling
const name = device.deviceName || 'Unknown';
const syncDate = device.lastSyncDateTime
  ? new Date(device.lastSyncDateTime)
  : null;

// Bad: Unsafe
const name = device.deviceName;  // Could be undefined
const syncDate = new Date(device.lastSyncDateTime); // Could throw
```

### 5. Performance

Minimize API calls:

```typescript
// Good: Single call with pagination
const response = await this.graphClient.api('/devices').top(999).get();
const allDevices = await this.getAllPages(response);

// Bad: Multiple calls
for (const id of deviceIds) {
  const device = await this.graphClient.api(`/devices/${id}`).get();
}
```

## Troubleshooting

### Common Issues

#### 1. Rate Limiting (429 Error)

**Symptom**: "TooManyRequests" error from Graph API

**Solution**: The retry logic automatically handles this with exponential backoff

```typescript
// Adjust retry parameters if needed
await this.retryGraphCall(apiCall, 5, 2000); // 5 retries, 2s base delay
```

#### 2. Insufficient Permissions

**Symptom**: "Forbidden" or "Unauthorized" errors

**Solution**: Verify app registration has required permissions:
- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementConfiguration.Read.All`

#### 3. Empty Results

**Symptom**: No devices returned

**Debug Steps**:
1. Check filter criteria
2. Verify devices have Exchange ActiveSync enabled
3. Check device management agent values

```typescript
// Add debugging
logger.debug('Filter criteria:', {
  managementAgent: ['eas', 'easMdm', 'mdm'],
  easActivated: true
});
```

#### 4. Memory Issues with Large Datasets

**Symptom**: Out of memory errors with many devices

**Solution**: Adjust batch size or process incrementally

```typescript
// Reduce batch size
const batchSize = 10; // Instead of 30

// Or process and export incrementally
for (const batch of batches) {
  const results = await processBatch(batch);
  await exportBatch(results);
}
```

## Conclusion

The Exchange ActiveSync Reports module demonstrates several key software engineering principles:

- **Abstraction**: Inherits from BaseReport for common functionality
- **Composition**: Combines multiple smaller methods into comprehensive reports
- **Type Safety**: Uses TypeScript for compile-time safety
- **Resilience**: Automatic retries and graceful degradation
- **Observability**: Comprehensive logging throughout
- **Testability**: Clear separation of concerns enables easy testing
- **Extensibility**: Easy to add new reports or customize existing ones

By following the patterns and practices outlined in this document, you can effectively extend and maintain the Exchange ActiveSync reporting functionality.

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/overview)
- [Intune Device Management](https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [Exchange ActiveSync Overview](https://docs.microsoft.com/en-us/exchange/clients/exchange-activesync/)
- [TypeScript Best Practices](https://www.typescriptlang.org/docs/handbook/declaration-files/do-s-and-don-ts.html)

---

**Document Version**: 1.0
**Last Updated**: 2024-02-07
**Maintained By**: Intune Reporting Team
