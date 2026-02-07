# Windows RT Sideloading Reports - Implementation Guide

This document provides detailed implementation guidance for the Windows RT Sideloading Reports module, including technical architecture, API integration details, and advanced usage patterns.

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Microsoft Graph API Integration](#microsoft-graph-api-integration)
- [Implementation Details](#implementation-details)
- [Type System](#type-system)
- [Key Status Determination Logic](#key-status-determination-logic)
- [Alert System Implementation](#alert-system-implementation)
- [Testing Strategy](#testing-strategy)
- [Performance Considerations](#performance-considerations)
- [Error Handling](#error-handling)
- [Advanced Usage Patterns](#advanced-usage-patterns)
- [Extending the Module](#extending-the-module)

## Architecture Overview

### Class Hierarchy

```
BaseReport (abstract)
  └── WindowsRTSideloadingReports
```

### Module Structure

```
src/reports/configmgr/
├── windows-rt-sideloading-reports.ts       # Main implementation
└── windows-rt-sideloading-reports.test.ts  # Comprehensive tests

docs/
├── CONFIGMGR_WINDOWS_RT_SIDELOADING.md     # User guide
└── CONFIGMGR_WINDOWS_RT_IMPL.md            # Implementation guide (this file)
```

### Core Components

1. **Data Retrieval Layer**: Interacts with Microsoft Graph API
2. **Processing Layer**: Analyzes and transforms raw data
3. **Reporting Layer**: Generates formatted reports
4. **Export Layer**: Outputs data in multiple formats
5. **Alert System**: Monitors and flags key health issues

## Microsoft Graph API Integration

### Primary Endpoints

#### Get All Sideloading Keys

```http
GET https://graph.microsoft.com/v1.0/deviceAppManagement/sideLoadingKeys
```

**Query Parameters:**
- `$select`: Fields to retrieve
- `$top`: Number of records to return (max 999)
- `$skip`: Pagination offset

**Implementation:**
```typescript
async getAllSideloadingKeys(): Promise<SideloadingKey[]> {
  const response = await this.retryGraphCall(() =>
    this.graphClient
      .api('/deviceAppManagement/sideLoadingKeys')
      .select([
        'id',
        'displayName',
        'value',
        'totalActivation',
        'lastUpdatedDateTime',
        'description',
      ])
      .top(999)
      .get()
  );

  return await this.getAllPages<SideloadingKey>(response);
}
```

#### Get Single Sideloading Key

```http
GET https://graph.microsoft.com/v1.0/deviceAppManagement/sideLoadingKeys/{id}
```

**Implementation:**
```typescript
async getSideloadingKeyById(keyId: string): Promise<SideloadingKey> {
  const key = await this.retryGraphCall(() =>
    this.graphClient
      .api(`/deviceAppManagement/sideLoadingKeys/${keyId}`)
      .select([
        'id',
        'displayName',
        'value',
        'totalActivation',
        'lastUpdatedDateTime',
        'description',
      ])
      .get()
  );

  return key;
}
```

### API Response Structure

**Raw API Response:**
```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#deviceAppManagement/sideLoadingKeys",
  "@odata.count": 6,
  "value": [
    {
      "id": "key-001",
      "displayName": "Production Sideloading Key",
      "value": "ABCD-EFGH-IJKL-MNOP-QRST-UVWX-YZ12-3456",
      "totalActivation": 1200,
      "lastUpdatedDateTime": "2024-01-15T10:30:00Z",
      "description": "Primary production sideloading key"
    }
  ],
  "@odata.nextLink": "https://graph.microsoft.com/v1.0/deviceAppManagement/sideLoadingKeys?$skip=999"
}
```

### Pagination Handling

The module implements automatic pagination to handle large datasets:

```typescript
private async getAllPages<T>(initialResponse: any): Promise<T[]> {
  let results = initialResponse.value || [];
  let nextLink = initialResponse['@odata.nextLink'];

  while (nextLink) {
    const response = await this.retryGraphCall(() =>
      this.graphClient.api(nextLink).get()
    );
    results = results.concat(response.value || []);
    nextLink = response['@odata.nextLink'];
  }

  return results;
}
```

### Retry Logic

Built-in retry mechanism for transient failures:

```typescript
// Inherited from BaseReport
protected async retryGraphCall<T>(
  operation: () => Promise<T>,
  maxRetries: number = 3,
  delayMs: number = 1000
): Promise<T> {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      return await operation();
    } catch (error) {
      if (attempt === maxRetries) throw error;

      // Exponential backoff
      await new Promise(resolve =>
        setTimeout(resolve, delayMs * Math.pow(2, attempt - 1))
      );
    }
  }

  throw new Error('Max retries exceeded');
}
```

## Implementation Details

### Report 35: Detailed Status Implementation

**Method:** `getSideloadingKeyDetailedStatus(keyId, thresholds?)`

**Processing Steps:**

1. **Fetch Key Data**
   ```typescript
   const key = await this.getSideloadingKeyById(keyId);
   ```

2. **Determine Status**
   ```typescript
   const status = this.determineKeyStatus(key);
   const statusReason = this.getStatusReason(key, status);
   ```

3. **Calculate Metrics**
   ```typescript
   const maxActivations = 5000; // Windows RT standard
   const activationsUsed = key.totalActivation || 0;
   const activationsRemaining = Math.max(0, maxActivations - activationsUsed);
   const usagePercentage = maxActivations > 0
     ? Math.round((activationsUsed / maxActivations) * 100 * 100) / 100
     : 0;
   ```

4. **Calculate Time Metrics**
   ```typescript
   const lastUpdated = key.lastUpdatedDateTime
     ? new Date(key.lastUpdatedDateTime)
     : new Date();
   const now = new Date();
   const daysSinceUpdate = Math.floor(
     (now.getTime() - lastUpdated.getTime()) / (1000 * 60 * 60 * 24)
   );
   ```

5. **Apply Thresholds**
   ```typescript
   const expirationWarningDays = thresholds?.expirationWarningDays ?? 30;
   const highUsagePercentage = thresholds?.highUsagePercentage ?? 80;

   const isNearExpiration = status === 'active' &&
     daysSinceUpdate > 365 - expirationWarningDays;
   const isHighUsage = usagePercentage >= highUsagePercentage;
   ```

6. **Build Response**
   ```typescript
   return {
     keyId: key.id,
     keyValue: this.maskKeyValue(key.value || 'N/A'),
     displayName: key.displayName || `Key ${key.id}`,
     totalActivations: maxActivations,
     activationsUsed,
     activationsRemaining,
     lastUpdatedDate: lastUpdated.toISOString(),
     lastUpdatedDaysAgo: daysSinceUpdate,
     status,
     statusReason,
     expirationWarning: isNearExpiration,
     usagePercentage,
     description: key.description || 'No description',
     isNearExpiration,
     isHighUsage,
   };
   ```

### Report 36: Summary Implementation

**Method:** `getSideloadingKeysSummary(filter?, thresholds?)`

**Processing Steps:**

1. **Fetch All Keys**
   ```typescript
   const keys = await this.getAllSideloadingKeys();
   ```

2. **Apply Filters**
   ```typescript
   let filteredKeys = keys;
   if (filter) {
     filteredKeys = this.applyFilters(keys, filter);
   }
   ```

3. **Categorize by Status**
   ```typescript
   const statusCounts = new Map<SideloadingKeyStatus, SideloadingKey[]>();
   filteredKeys.forEach((key) => {
     const status = this.determineKeyStatus(key);
     if (!statusCounts.has(status)) {
       statusCounts.set(status, []);
     }
     statusCounts.get(status)!.push(key);
   });
   ```

4. **Calculate Statistics**
   ```typescript
   const totalActivationsAcrossAllKeys = filteredKeys.reduce(
     (sum, key) => sum + (key.totalActivation || 0),
     0
   );

   const averageActivationsPerKey = totalKeys > 0
     ? Math.round((totalActivationsAcrossAllKeys / totalKeys) * 100) / 100
     : 0;
   ```

5. **Identify Extremes**
   ```typescript
   const sortedByUsage = [...filteredKeys].sort(
     (a, b) => (b.totalActivation || 0) - (a.totalActivation || 0)
   );

   const mostUsedKey = sortedByUsage[0];
   const leastUsedKey = sortedByUsage[sortedByUsage.length - 1];
   ```

6. **Generate Alerts**
   ```typescript
   const expiringKeys: string[] = [];
   const highUsageKeys: string[] = [];
   const unusedKeys: string[] = [];

   for (const key of filteredKeys) {
     // Check expiration
     if (status === 'active' && daysSinceUpdate > 365 - expirationWarningDays) {
       expiringKeys.push(key.displayName || key.id);
     }

     // Check high usage
     const usagePercent = (activations / maxActivations) * 100;
     if (usagePercent >= highUsagePercentage) {
       highUsageKeys.push(key.displayName || key.id);
     }

     // Check low usage
     if (usagePercent <= lowUsagePercentage) {
       unusedKeys.push(key.displayName || key.id);
     }
   }
   ```

## Type System

### Core Interfaces

#### SideloadingKey

Raw data structure from Microsoft Graph API:

```typescript
interface SideloadingKey {
  id: string;                      // Unique identifier
  displayName?: string;            // Human-readable name
  value?: string;                  // Key value (product key)
  totalActivation: number;         // Current activation count
  lastUpdatedDateTime?: string;    // ISO 8601 timestamp
  description?: string;            // Key description
  '@odata.type'?: string;         // OData type metadata
}
```

#### SideloadingKeyDetailedStatus

Processed detailed information:

```typescript
interface SideloadingKeyDetailedStatus {
  keyId: string;                   // Unique identifier
  keyValue: string;                // Masked key value
  displayName: string;             // Display name
  totalActivations: number;        // Maximum activations (5000)
  activationsUsed: number;         // Current usage
  activationsRemaining: number;    // Remaining activations
  lastUpdatedDate: string;         // ISO 8601 timestamp
  lastUpdatedDaysAgo: number;      // Days since last update
  status: SideloadingKeyStatus;    // Current status
  statusReason: string;            // Human-readable reason
  expirationWarning: boolean;      // Expiration flag
  usagePercentage: number;         // Usage percentage (0-100)
  description: string;             // Key description
  createdDateTime?: string;        // Creation timestamp
  isNearExpiration: boolean;       // Near expiration flag
  isHighUsage: boolean;            // High usage flag
}
```

#### SideloadingKeysSummary

Aggregated statistics:

```typescript
interface SideloadingKeysSummary {
  totalKeys: number;
  activeKeysCount: number;
  expiredKeysCount: number;
  revokedKeysCount: number;
  unknownKeysCount: number;
  keysByStatus: Array<{
    status: SideloadingKeyStatus;
    count: number;
    percentage: number;
    keys: string[];
  }>;
  usageStatistics: {
    totalActivationsAcrossAllKeys: number;
    averageActivationsPerKey: number;
    mostUsedKey: {
      keyId: string;
      displayName: string;
      activations: number;
    } | null;
    leastUsedKey: {
      keyId: string;
      displayName: string;
      activations: number;
    } | null;
    averageUsagePercentage: number;
  };
  alerts: {
    expiringKeysCount: number;
    highUsageKeysCount: number;
    unusedKeysCount: number;
    expiringKeys: string[];
    highUsageKeys: string[];
    unusedKeys: string[];
  };
  lastUpdatedDateTime: string;
}
```

### Type Aliases

```typescript
type SideloadingKeyStatus = 'active' | 'expired' | 'revoked' | 'unknown';
```

### Filter Types

```typescript
interface SideloadingKeyFilter {
  status?: SideloadingKeyStatus;
  minActivations?: number;
  maxActivations?: number;
  minUsagePercentage?: number;
  maxUsagePercentage?: number;
  includeExpired?: boolean;
  includeRevoked?: boolean;
}
```

### Alert Configuration

```typescript
interface AlertThresholds {
  expirationWarningDays?: number;    // Default: 30
  highUsagePercentage?: number;      // Default: 80
  lowUsagePercentage?: number;       // Default: 10
}
```

## Key Status Determination Logic

### Status Decision Tree

```typescript
private determineKeyStatus(key: SideloadingKey): SideloadingKeyStatus {
  // 1. Check expiration by date
  if (key.lastUpdatedDateTime) {
    const daysSinceUpdate = this.calculateDaysSince(key.lastUpdatedDateTime);
    if (daysSinceUpdate > 365) {
      return 'expired';
    }
  }

  // 2. Check expiration by activation count
  const maxActivations = 5000;
  if ((key.totalActivation || 0) >= maxActivations) {
    return 'expired';
  }

  // 3. Check if revoked (no key value)
  if (!key.value) {
    return 'revoked';
  }

  // 4. Default to active
  return 'active';
}
```

### Status Reasons

```typescript
private getStatusReason(
  key: SideloadingKey,
  status: SideloadingKeyStatus
): string {
  switch (status) {
    case 'active':
      return 'Key is currently active and available for use';

    case 'expired':
      if (key.totalActivation && key.totalActivation >= 5000) {
        return 'Key has reached maximum activation limit (5000)';
      }
      if (key.lastUpdatedDateTime) {
        const daysSinceUpdate = this.calculateDaysSince(key.lastUpdatedDateTime);
        return `Key has expired (${daysSinceUpdate} days since last update)`;
      }
      return 'Key has expired';

    case 'revoked':
      return 'Key has been revoked and is no longer valid';

    case 'unknown':
    default:
      return 'Key status could not be determined';
  }
}
```

### Key Value Masking

Security implementation to protect sensitive key values:

```typescript
private maskKeyValue(value: string): string {
  if (!value || value === 'N/A') {
    return 'N/A';
  }

  // Show first 4 and last 4 characters
  if (value.length <= 8) {
    return '****';
  }

  const start = value.substring(0, 4);
  const end = value.substring(value.length - 4);
  const masked = '*'.repeat(Math.min(value.length - 8, 20));

  return `${start}${masked}${end}`;
}
```

**Examples:**
- Input: `ABCD-EFGH-IJKL-MNOP-QRST-UVWX-YZ12-3456`
- Output: `ABCD********************3456`

## Alert System Implementation

### Alert Thresholds

Default values defined as class constants:

```typescript
class WindowsRTSideloadingReports extends BaseReport {
  private readonly DEFAULT_EXPIRATION_WARNING_DAYS = 30;
  private readonly DEFAULT_HIGH_USAGE_PERCENTAGE = 80;
  private readonly DEFAULT_LOW_USAGE_PERCENTAGE = 10;
}
```

### Alert Detection Logic

#### Expiring Keys Detection

```typescript
// Check for keys approaching expiration
if (status === 'active' && key.lastUpdatedDateTime) {
  const daysSinceUpdate = this.calculateDaysSince(key.lastUpdatedDateTime);

  if (daysSinceUpdate > 365 - expirationWarningDays) {
    expiringKeys.push(key.displayName || key.id);
  }
}
```

**Logic:**
- Only active keys are checked
- Warning triggers when: `daysSinceUpdate > (365 - threshold)`
- Default: Alert 30 days before 1-year expiration

#### High Usage Detection

```typescript
const maxActivations = 5000;
const usagePercent = (key.totalActivation / maxActivations) * 100;

if (usagePercent >= highUsagePercentage) {
  highUsageKeys.push(key.displayName || key.id);
}
```

**Logic:**
- Calculated as percentage: `(used / max) * 100`
- Default threshold: 80%
- Triggers when: `usagePercent >= threshold`

#### Unused Keys Detection

```typescript
const maxActivations = 5000;
const usagePercent = (key.totalActivation / maxActivations) * 100;

if (usagePercent <= lowUsagePercentage) {
  unusedKeys.push(key.displayName || key.id);
}
```

**Logic:**
- Same calculation as high usage
- Default threshold: 10%
- Identifies potentially abandoned keys

### Custom Alert Thresholds

Users can override default thresholds:

```typescript
const summary = await report.getSideloadingKeysSummary(undefined, {
  expirationWarningDays: 45,    // Warn 45 days before expiration
  highUsagePercentage: 90,      // Alert at 90% usage
  lowUsagePercentage: 5,        // Alert if usage < 5%
});
```

## Testing Strategy

### Test Coverage

The module includes comprehensive unit tests covering:

1. **API Integration Tests**
   - Fetching all keys
   - Fetching specific keys
   - Pagination handling

2. **Status Calculation Tests**
   - Active key detection
   - Expired key detection (by date)
   - Expired key detection (by activation count)
   - Revoked key detection

3. **Metrics Calculation Tests**
   - Usage percentage calculation
   - Activation remaining calculation
   - Days since update calculation

4. **Alert System Tests**
   - Expiring keys detection
   - High usage keys detection
   - Unused keys detection
   - Custom threshold application

5. **Filter Tests**
   - Status filtering
   - Activation count filtering
   - Usage percentage filtering
   - Combined filters

6. **Error Handling Tests**
   - API failures
   - Missing data handling
   - Invalid key IDs

### Mock Implementation

```typescript
class MockGraphClient {
  private responses: Map<string, any> = new Map();

  api(endpoint: string) {
    return {
      select: (fields: string[]) => this,
      top: (value: number) => this,
      get: async () => this.responses.get(endpoint) || { value: [] },
    };
  }
}
```

### Test Execution

```bash
# Run tests
npm test windows-rt-sideloading-reports.test.ts

# Run with coverage
npm test -- --coverage windows-rt-sideloading-reports.test.ts

# Run specific test suite
npm test -- -t "getSideloadingKeyDetailedStatus"
```

## Performance Considerations

### Caching Strategy

For improved performance with large key sets:

```typescript
class CachedSideloadingReports extends WindowsRTSideloadingReports {
  private keyCache: Map<string, SideloadingKey> = new Map();
  private cacheExpiry: number = 5 * 60 * 1000; // 5 minutes
  private lastFetch: number = 0;

  async getAllSideloadingKeys(): Promise<SideloadingKey[]> {
    const now = Date.now();

    if (now - this.lastFetch < this.cacheExpiry && this.keyCache.size > 0) {
      return Array.from(this.keyCache.values());
    }

    const keys = await super.getAllSideloadingKeys();

    this.keyCache.clear();
    keys.forEach(key => this.keyCache.set(key.id, key));
    this.lastFetch = now;

    return keys;
  }
}
```

### Batch Processing

For processing large numbers of keys:

```typescript
async function processKeysInBatches(
  report: WindowsRTSideloadingReports,
  batchSize: number = 10
): Promise<SideloadingKeyDetailedStatus[]> {
  const keys = await report.getAllSideloadingKeys();
  const results: SideloadingKeyDetailedStatus[] = [];

  for (let i = 0; i < keys.length; i += batchSize) {
    const batch = keys.slice(i, i + batchSize);
    const batchResults = await Promise.all(
      batch.map(key => report.getSideloadingKeyDetailedStatus(key.id))
    );
    results.push(...batchResults);
  }

  return results;
}
```

### Query Optimization

Minimize API calls by requesting only needed fields:

```typescript
// Good: Only request needed fields
.select(['id', 'displayName', 'totalActivation'])

// Avoid: Requesting all fields
.select(['*'])
```

## Error Handling

### Error Types

```typescript
// API errors
try {
  const keys = await report.getAllSideloadingKeys();
} catch (error) {
  if (error.statusCode === 401) {
    // Authentication error
    console.error('Authentication failed - check credentials');
  } else if (error.statusCode === 403) {
    // Permission error
    console.error('Insufficient permissions');
  } else if (error.statusCode === 429) {
    // Rate limiting
    console.error('Rate limit exceeded - retry later');
  } else {
    // Generic error
    console.error('API error:', error.message);
  }
}
```

### Graceful Degradation

```typescript
async getSideloadingKeyDetailedStatus(
  keyId: string
): Promise<SideloadingKeyDetailedStatus> {
  try {
    const key = await this.getSideloadingKeyById(keyId);
    // Process key...
  } catch (error) {
    logger.error('Failed to fetch key details', { keyId, error });

    // Return minimal status
    return {
      keyId,
      keyValue: 'N/A',
      displayName: 'Unknown',
      totalActivations: 0,
      activationsUsed: 0,
      activationsRemaining: 0,
      lastUpdatedDate: new Date().toISOString(),
      lastUpdatedDaysAgo: 0,
      status: 'unknown',
      statusReason: 'Failed to retrieve key information',
      expirationWarning: false,
      usagePercentage: 0,
      description: 'Error retrieving key',
      isNearExpiration: false,
      isHighUsage: false,
    };
  }
}
```

### Logging

The module uses comprehensive logging for debugging:

```typescript
logger.info('Fetching all sideloading keys');
logger.info('Sideloading keys retrieved', { count: keys.length });
logger.error('Failed to fetch sideloading keys', error);
logger.warn('Key status could not be determined', { keyId });
```

## Advanced Usage Patterns

### Pattern 1: Scheduled Monitoring

```typescript
import { WindowsRTSideloadingReports } from './windows-rt-sideloading-reports';
import cron from 'node-cron';

// Run daily at 8 AM
cron.schedule('0 8 * * *', async () => {
  const report = new WindowsRTSideloadingReports(graphClient, config);
  const summary = await report.getSideloadingKeysSummary();

  // Check for critical alerts
  if (summary.alerts.expiringKeysCount > 0 ||
      summary.alerts.highUsageKeysCount > 0) {
    await sendAlertEmail(summary);
  }

  // Archive report
  await report.exportReport(
    { metadata: {}, data: summary, summary: {} },
    { format: 'json', outputDir: './reports/archive', includeTimestamp: true }
  );
});
```

### Pattern 2: Key Lifecycle Automation

```typescript
async function manageKeyLifecycle(
  report: WindowsRTSideloadingReports
): Promise<void> {
  const keys = await report.getAllSideloadingKeys();

  for (const key of keys) {
    const status = await report.getSideloadingKeyDetailedStatus(key.id);

    // Auto-renew keys approaching expiration
    if (status.isNearExpiration && !status.expirationWarning) {
      await scheduleKeyRenewal(key);
    }

    // Flag high usage keys for replacement
    if (status.usagePercentage >= 90) {
      await createReplacementKeyRequest(key);
    }

    // Archive unused keys
    if (status.usagePercentage < 5 && status.lastUpdatedDaysAgo > 180) {
      await archiveUnusedKey(key);
    }
  }
}
```

### Pattern 3: Compliance Reporting

```typescript
async function generateComplianceReport(
  report: WindowsRTSideloadingReports
): Promise<ComplianceReport> {
  const summary = await report.getSideloadingKeysSummary();

  return {
    timestamp: new Date().toISOString(),
    totalKeys: summary.totalKeys,
    activeKeys: summary.activeKeysCount,
    expiredKeys: summary.expiredKeysCount,
    revokedKeys: summary.revokedKeysCount,
    compliance: {
      allKeysValid: summary.expiredKeysCount === 0 && summary.revokedKeysCount === 0,
      noExpiringKeys: summary.alerts.expiringKeysCount === 0,
      healthyUsage: summary.alerts.highUsageKeysCount === 0,
    },
    recommendations: generateRecommendations(summary),
  };
}
```

### Pattern 4: Multi-Format Reporting

```typescript
async function exportAllFormats(
  report: WindowsRTSideloadingReports
): Promise<string[]> {
  const reportData = await report.generateComprehensiveReport();
  const formats: Array<'json' | 'csv' | 'html'> = ['json', 'csv', 'html'];

  const exportPromises = formats.map(format =>
    report.exportReport(reportData, {
      format,
      outputDir: './reports',
      includeTimestamp: true,
    })
  );

  return await Promise.all(exportPromises);
}
```

## Extending the Module

### Custom Status Logic

```typescript
class CustomSideloadingReports extends WindowsRTSideloadingReports {
  protected determineKeyStatus(key: SideloadingKey): SideloadingKeyStatus {
    // Custom status logic
    if (this.isKeyInGracePeriod(key)) {
      return 'active'; // Override expiration
    }

    return super.determineKeyStatus(key);
  }

  private isKeyInGracePeriod(key: SideloadingKey): boolean {
    // Custom grace period logic
    const daysSinceUpdate = this.calculateDaysSince(key.lastUpdatedDateTime);
    return daysSinceUpdate <= 395; // 30-day grace period after 1 year
  }
}
```

### Additional Metrics

```typescript
class EnhancedSideloadingReports extends WindowsRTSideloadingReports {
  async getKeyTrends(keyId: string): Promise<KeyTrends> {
    const key = await this.getSideloadingKeyById(keyId);

    // Fetch historical data (custom implementation)
    const history = await this.getKeyHistory(keyId);

    return {
      keyId,
      displayName: key.displayName,
      activationsPerDay: this.calculateDailyActivations(history),
      projectedExpiration: this.projectExpirationDate(key, history),
      usageTrend: this.analyzeUsageTrend(history),
    };
  }

  private calculateDailyActivations(history: any[]): number {
    // Custom calculation logic
    return 0;
  }

  private projectExpirationDate(key: SideloadingKey, history: any[]): Date {
    // Predictive analysis
    return new Date();
  }

  private analyzeUsageTrend(history: any[]): 'increasing' | 'decreasing' | 'stable' {
    // Trend analysis
    return 'stable';
  }
}
```

### Custom Export Formats

```typescript
class ExtendedSideloadingReports extends WindowsRTSideloadingReports {
  async exportToPDF(reportData: ReportData, outputPath: string): Promise<string> {
    // Custom PDF export implementation
    const pdfGenerator = new PDFGenerator();
    return await pdfGenerator.generate(reportData, outputPath);
  }

  async exportToExcel(reportData: ReportData, outputPath: string): Promise<string> {
    // Custom Excel export implementation
    const excelGenerator = new ExcelGenerator();
    return await excelGenerator.generate(reportData, outputPath);
  }
}
```

## Integration Examples

### Integration with Azure Functions

```typescript
import { AzureFunction, Context, HttpRequest } from '@azure/functions';
import { WindowsRTSideloadingReports } from './windows-rt-sideloading-reports';

const httpTrigger: AzureFunction = async function (
  context: Context,
  req: HttpRequest
): Promise<void> {
  try {
    const report = new WindowsRTSideloadingReports(graphClient, config);
    const summary = await report.getSideloadingKeysSummary();

    context.res = {
      status: 200,
      body: summary,
      headers: {
        'Content-Type': 'application/json',
      },
    };
  } catch (error) {
    context.res = {
      status: 500,
      body: { error: error.message },
    };
  }
};

export default httpTrigger;
```

### Integration with Power BI

```typescript
async function exportForPowerBI(
  report: WindowsRTSideloadingReports
): Promise<void> {
  const summary = await report.getSideloadingKeysSummary();

  // Flatten data for Power BI
  const powerBIData = {
    keys: summary.keysByStatus.flatMap(status =>
      status.keys.map(key => ({
        name: key,
        status: status.status,
        count: 1,
        timestamp: new Date().toISOString(),
      }))
    ),
    metrics: {
      totalKeys: summary.totalKeys,
      activeKeys: summary.activeKeysCount,
      expiredKeys: summary.expiredKeysCount,
      timestamp: new Date().toISOString(),
    },
  };

  // Export to CSV for Power BI ingestion
  await fs.writeFile(
    './powerbi-data.json',
    JSON.stringify(powerBIData, null, 2)
  );
}
```

## Best Practices Summary

1. **Always use retry logic** for API calls to handle transient failures
2. **Implement caching** for frequently accessed data
3. **Mask sensitive data** (key values) in reports and logs
4. **Use custom thresholds** to match organizational policies
5. **Monitor performance** with logging and metrics
6. **Handle errors gracefully** with fallback values
7. **Test thoroughly** with comprehensive unit and integration tests
8. **Document customizations** when extending the module
9. **Version control reports** for audit and compliance
10. **Automate monitoring** with scheduled tasks

## References

- [Microsoft Graph API - Sideloading Keys](https://learn.microsoft.com/en-us/graph/api/resources/intune-onboarding-sideloadingkey)
- [TypeScript Best Practices](https://www.typescriptlang.org/docs/handbook/declaration-files/do-s-and-don-ts.html)
- [Node.js Error Handling](https://nodejs.org/api/errors.html)

---

**Document Version:** 1.0.0
**Last Updated:** 2024-02-07
**Module Version:** 1.0.0
