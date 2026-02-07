# Mobile Device Inventory Reports - Implementation Guide

This document provides technical implementation details, architecture notes, and developer guidance for the Mobile Device Inventory Reports module.

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Implementation Details](#implementation-details)
- [Configuration Manager Report Mapping](#configuration-manager-report-mapping)
- [Microsoft Graph API Integration](#microsoft-graph-api-integration)
- [Data Flow](#data-flow)
- [Code Structure](#code-structure)
- [TypeScript Type System](#typescript-type-system)
- [Testing Strategy](#testing-strategy)
- [Performance Considerations](#performance-considerations)
- [Error Handling](#error-handling)
- [Extending the Module](#extending-the-module)
- [Deployment Notes](#deployment-notes)

## Architecture Overview

### Module Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                  MobileDeviceInventoryReports               │
│                    (Main Report Class)                      │
├─────────────────────────────────────────────────────────────┤
│  - extends BaseReport                                       │
│  - implements Report interface                              │
│  - uses Graph Client for API calls                          │
│  - provides ConfigMgr-compatible reports                    │
└─────────────────────────────────────────────────────────────┘
                            │
                            │ uses
                            ↓
┌─────────────────────────────────────────────────────────────┐
│                Microsoft Graph Client                       │
├─────────────────────────────────────────────────────────────┤
│  - Authentication handling                                  │
│  - API call execution                                       │
│  - Retry logic with exponential backoff                     │
│  - Pagination support                                       │
└─────────────────────────────────────────────────────────────┘
                            │
                            │ queries
                            ↓
┌─────────────────────────────────────────────────────────────┐
│            /deviceManagement/managedDevices                 │
│                  (Graph API Endpoint)                       │
├─────────────────────────────────────────────────────────────┤
│  - Returns managed device data                              │
│  - Supports OData filtering                                 │
│  - Supports field selection                                 │
│  - Handles pagination automatically                         │
└─────────────────────────────────────────────────────────────┘
```

### Class Hierarchy

```
Report (interface)
    ↓
BaseReport (abstract class)
    ↓
MobileDeviceInventoryReports (concrete implementation)
```

## Implementation Details

### File Location

**Main Implementation**: `/src/reports/configmgr/mobile-device-inventory-reports.ts`

**Supporting Files**:
- Tests: `/src/reports/configmgr/mobile-device-inventory-reports.test.ts`
- Examples: `/src/reports/examples/mobile-device-inventory-examples.ts`
- Documentation: `/docs/CONFIGMGR_MOBILE_INVENTORY.md`
- Implementation Docs: `/docs/CONFIGMGR_MOBILE_INVENTORY_IMPL.md` (this file)

### Dependencies

```typescript
// External dependencies
import { Client } from '@microsoft/microsoft-graph-client';

// Internal dependencies
import { BaseReport } from '../base-report';
import { ReportData, AppConfig } from '../../types';
import { Logger } from '../../core/logger';
import { OutputFormatter } from '../../formatters/output-formatter';
```

### Key Components

1. **Enums**: Define device ownership types, management agents, and communication status
2. **Interfaces**: Define data structures for devices, filters, sorting, and pagination
3. **Main Class**: Implements all report methods and helper functions
4. **Convenience Functions**: Quick report generation helpers
5. **Export Methods**: Support for JSON, CSV, and HTML output

## Configuration Manager Report Mapping

### Report 1: All Corporate-Owned Mobile Devices

**ConfigMgr Query Equivalent**:
```sql
SELECT
    v_GS_MOBILE_DEVICE_GENERAL.DeviceID,
    v_GS_MOBILE_DEVICE_GENERAL.DeviceName,
    v_GS_MOBILE_DEVICE_GENERAL.OperatingSystem,
    v_GS_MOBILE_DEVICE_GENERAL.Manufacturer,
    v_GS_MOBILE_DEVICE_GENERAL.Model
FROM v_GS_MOBILE_DEVICE_GENERAL
WHERE v_GS_MOBILE_DEVICE_GENERAL.OwnerType = 'Company'
```

**Graph API Implementation**:
```typescript
async getAllCorporateOwnedDevices() {
  const filter = "managedDeviceOwnerType eq 'company'";

  return await this.graphClient
    .api('/deviceManagement/managedDevices')
    .filter(filter)
    .select([...fields])
    .get();
}
```

**Data Mapping**:
| ConfigMgr Field | Graph API Field | Type |
|----------------|-----------------|------|
| `DeviceID` | `id` | string |
| `DeviceName` | `deviceName` | string |
| `OperatingSystem` | `operatingSystem` | string |
| `OSVersion` | `osVersion` | string |
| `OwnerType` | `managedDeviceOwnerType` | enum |

### Report 2: All Mobile Device Clients

**ConfigMgr Query Equivalent**:
```sql
SELECT *
FROM v_GS_MOBILE_DEVICE_CLIENT
WHERE ManagementAgent <> 'EAS'
```

**Graph API Implementation**:
```typescript
async getAllMobileDeviceClients() {
  const filter = "managementAgent ne 'eas'";

  return await this.graphClient
    .api('/deviceManagement/managedDevices')
    .filter(filter)
    .get();
}
```

**Exclusion Logic**: Filters out devices managed only through Exchange ActiveSync connector, ensuring only MDM-managed devices are included.

### Report 20: Mobile Device Client Information

**ConfigMgr Data Sources**:
- `v_ClientHealthState` - Client health status
- `v_CH_ClientSummary` - Client communication history

**Graph API Implementation**:
```typescript
async getMobileDeviceClientInformation() {
  const devices = await this.getAllMobileDeviceClients();

  // Calculate communication status
  return devices.map(device => ({
    ...device,
    communicationStatus: this.determineCommunicationStatus(
      this.calculateDaysSince(device.lastSyncDateTime)
    ),
    isActive: daysSinceLastContact <= 30
  }));
}
```

**Communication Status Calculation**:
```typescript
private determineCommunicationStatus(daysSinceLastContact?: number): CommunicationStatus {
  if (!daysSinceLastContact) return CommunicationStatus.UNKNOWN;
  if (daysSinceLastContact <= 1) return CommunicationStatus.HEALTHY;
  if (daysSinceLastContact <= 7) return CommunicationStatus.WARNING;
  return CommunicationStatus.CRITICAL;
}
```

### Report 22: Mobile Devices by Operating System

**ConfigMgr Query Equivalent**:
```sql
SELECT
    OperatingSystem,
    COUNT(*) as DeviceCount
FROM v_GS_MOBILE_DEVICE_GENERAL
GROUP BY OperatingSystem
```

**Graph API Implementation**:
```typescript
async getMobileDevicesByOperatingSystem() {
  const devices = await this.getAllMobileDeviceClients();

  // Group by OS in memory
  const osGroups = new Map<string, MobileDeviceInfo[]>();
  devices.forEach(device => {
    const os = device.operatingSystem;
    if (!osGroups.has(os)) osGroups.set(os, []);
    osGroups.get(os)!.push(device);
  });

  // Calculate distribution
  return this.buildOSDistribution(osGroups, devices.length);
}
```

## Microsoft Graph API Integration

### Endpoint Details

**Primary Endpoint**: `GET /deviceManagement/managedDevices`

**API Version**: v1.0 (stable)

**Response Format**: JSON with OData metadata

### Authentication

**Required Permissions**:
- Delegated: `DeviceManagementManagedDevices.Read.All`
- Application: `DeviceManagementManagedDevices.Read.All`

**Authentication Flow**:
```typescript
// Application uses client credentials
const clientSecretCredential = new ClientSecretCredential(
  tenantId,
  clientId,
  clientSecret
);

// Initialize Graph client
const graphClient = Client.initWithMiddleware({
  authProvider: new TokenCredentialAuthenticationProvider(
    clientSecretCredential,
    { scopes: ['https://graph.microsoft.com/.default'] }
  )
});
```

### OData Query Support

**Filtering**:
```typescript
// Single filter
$filter=operatingSystem eq 'iOS'

// Multiple filters
$filter=operatingSystem eq 'iOS' and managedDeviceOwnerType eq 'company'

// Exclusion filter
$filter=managementAgent ne 'eas'
```

**Field Selection**:
```typescript
$select=id,deviceName,operatingSystem,osVersion,lastSyncDateTime
```

**Top/Skip (Not used - pagination handled automatically)**:
```typescript
$top=100
$skip=100
```

### Pagination Handling

**Graph API Pagination**:
```typescript
protected async getAllPages<T>(initialResponse: any): Promise<T[]> {
  let results: T[] = initialResponse.value || [];
  let nextLink = initialResponse['@odata.nextLink'];

  while (nextLink) {
    const response = await this.graphClient.api(nextLink).get();
    results = results.concat(response.value || []);
    nextLink = response['@odata.nextLink'];
  }

  return results;
}
```

**Automatic Pagination**: The `getAllPages` method from `BaseReport` automatically follows `@odata.nextLink` to retrieve all results.

### Retry Logic

**Implementation**:
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
        await this.sleep(delay);
      } else {
        throw error;
      }
    }
  }
}
```

**Retry Scenarios**:
- HTTP 429 (Too Many Requests) - Exponential backoff
- HTTP 5xx (Server Errors) - Exponential backoff
- Network timeouts - Exponential backoff

## Data Flow

### Request Flow

```
┌────────────────┐
│ Report Method  │
│ (e.g., get...  │
│  Corporate     │
│  Devices)      │
└────────┬───────┘
         │
         │ 1. Build OData filter
         ↓
┌────────────────┐
│ Build Filter   │
│ buildOData     │
│ Filter()       │
└────────┬───────┘
         │
         │ 2. Execute Graph API call
         ↓
┌────────────────┐
│ retryGraphCall │
│ with           │
│ exponential    │
│ backoff        │
└────────┬───────┘
         │
         │ 3. Retrieve all pages
         ↓
┌────────────────┐
│ getAllPages()  │
│ handles        │
│ pagination     │
└────────┬───────┘
         │
         │ 4. Map to TypeScript interfaces
         ↓
┌────────────────┐
│ Map raw data   │
│ to typed       │
│ interfaces     │
└────────┬───────┘
         │
         │ 5. Apply in-memory filtering/sorting
         ↓
┌────────────────┐
│ Apply client-  │
│ side filters   │
│ & sorting      │
└────────┬───────┘
         │
         │ 6. Return results
         ↓
┌────────────────┐
│ Return typed   │
│ device array   │
└────────────────┘
```

### Data Transformation

**Raw Graph Response**:
```json
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#deviceManagement/managedDevices",
  "value": [
    {
      "id": "device-001",
      "deviceName": "iPhone 14 Pro",
      "operatingSystem": "iOS",
      "osVersion": "17.2.1",
      "managedDeviceOwnerType": "company",
      "managementAgent": "mdm",
      "lastSyncDateTime": "2024-02-06T08:00:00Z"
    }
  ],
  "@odata.nextLink": "https://..."
}
```

**Transformed TypeScript Object**:
```typescript
{
  deviceId: 'device-001',
  deviceName: 'iPhone 14 Pro',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  ownerType: DeviceOwnershipType.COMPANY,
  managementAgent: ManagementAgentType.MDM,
  lastSyncDateTime: new Date('2024-02-06T08:00:00Z')
}
```

### Aggregation Pipeline

For OS distribution reports:

```
Raw Devices
    ↓
Group by OS
    ↓
Group by Version (within OS)
    ↓
Calculate Counts & Percentages
    ↓
Sort by Device Count
    ↓
Return OSDistribution[]
```

## Code Structure

### File Organization

```
src/reports/configmgr/mobile-device-inventory-reports.ts
│
├── Enums (lines 34-68)
│   ├── DeviceOwnershipType
│   ├── ManagementAgentType
│   └── CommunicationStatus
│
├── Interfaces (lines 70-201)
│   ├── MobileDeviceInfo
│   ├── CorporateMobileDevice
│   ├── MobileDeviceClientInfo
│   ├── OSDistribution
│   ├── OSVersionDistribution
│   ├── MobileDeviceFilterOptions
│   ├── MobileDeviceSortOptions
│   ├── PaginationOptions
│   ├── PaginatedResult
│   └── ExportOptions
│
├── Main Class: MobileDeviceInventoryReports (lines 203-930)
│   │
│   ├── Constructor & Properties (lines 212-220)
│   │
│   ├── execute() - Default Report (lines 225-252)
│   │
│   ├── Report Methods (lines 254-584)
│   │   ├── getAllCorporateOwnedDevices() - Report 1
│   │   ├── getAllMobileDeviceClients() - Report 2
│   │   ├── getMobileDeviceClientInformation() - Report 20
│   │   └── getMobileDevicesByOperatingSystem() - Report 22
│   │
│   ├── Utility Methods (lines 586-762)
│   │   ├── getAllMobileDevices()
│   │   ├── getDevicesByOwnerType()
│   │   ├── getDevicesByOS()
│   │   ├── getDevicesByManagementAgent()
│   │   ├── getInactiveDevices()
│   │   ├── getActiveDevices()
│   │   ├── paginate()
│   │   └── exportReport()
│   │
│   └── Private Helper Methods (lines 764-929)
│       ├── buildODataFilter()
│       ├── mapOwnerType()
│       ├── mapManagementAgent()
│       ├── calculateDaysSince()
│       ├── determineCommunicationStatus()
│       ├── compareDevices()
│       ├── groupByOwnership()
│       ├── groupByOS()
│       └── groupByManagementAgent()
│
└── Convenience Functions (lines 932-1023)
    ├── generateCorporateDevicesReport()
    ├── generateOSDistributionReport()
    └── generateDeviceClientCommunicationReport()
```

### Method Categories

**Public Report Methods**: Core ConfigMgr-compatible reports
- `getAllCorporateOwnedDevices()`
- `getAllMobileDeviceClients()`
- `getMobileDeviceClientInformation()`
- `getMobileDevicesByOperatingSystem()`

**Public Utility Methods**: Helper methods for common operations
- `getAllMobileDevices()`
- `getDevicesByOwnerType()`
- `getDevicesByOS()`
- `getActiveDevices()`
- `getInactiveDevices()`
- `paginate()`
- `exportReport()`

**Private Helper Methods**: Internal implementation details
- `buildODataFilter()`
- `mapOwnerType()`
- `mapManagementAgent()`
- `calculateDaysSince()`
- `determineCommunicationStatus()`
- `compareDevices()`
- `groupByOwnership()`
- `groupByOS()`
- `groupByManagementAgent()`

## TypeScript Type System

### Type Hierarchy

```
MobileDeviceInfo (base interface)
    ↓
CorporateMobileDevice (extends MobileDeviceInfo)
    ↓
    Additional fields:
    - purchaseDate
    - warrantyExpirationDate
    - assetTag
    - departmentName
    - costCenter
    - enrollmentProfileName

MobileDeviceInfo (base interface)
    ↓
MobileDeviceClientInfo (extends MobileDeviceInfo)
    ↓
    Additional fields:
    - communicationStatus
    - lastContactDateTime
    - daysSinceLastContact
    - isActive
    - managementPointUrl
    - clientVersion
```

### Enum Type Safety

**DeviceOwnershipType**:
```typescript
export enum DeviceOwnershipType {
  UNKNOWN = 'unknown',
  COMPANY = 'company',
  PERSONAL = 'personal',
}

// Usage ensures type safety
const type: DeviceOwnershipType = DeviceOwnershipType.COMPANY;
// Compiler error if invalid value:
// const invalid: DeviceOwnershipType = 'invalid'; // ❌ Error
```

**ManagementAgentType**:
```typescript
export enum ManagementAgentType {
  MDM = 'mdm',
  EAS = 'eas',
  INTUNE_MDM = 'intuneMdm',
  // ... etc
}
```

**CommunicationStatus**:
```typescript
export enum CommunicationStatus {
  HEALTHY = 'healthy',
  WARNING = 'warning',
  CRITICAL = 'critical',
  UNKNOWN = 'unknown',
}
```

### Generic Type Support

**Pagination**:
```typescript
paginate<T>(data: T[], options: PaginationOptions): PaginatedResult<T>

// Type-safe usage
const devices: MobileDeviceInfo[] = await report.getAllMobileDevices();
const page: PaginatedResult<MobileDeviceInfo> = report.paginate(devices, {...});
```

**getAllPages**:
```typescript
protected async getAllPages<T>(initialResponse: any): Promise<T[]>

// Type inference
const devices = await this.getAllPages<ManagedDevice>(response);
```

### Optional Properties

**Interface Design**:
```typescript
interface MobileDeviceInfo {
  deviceId: string;              // Required
  deviceName: string;            // Required
  operatingSystem: string;       // Required
  manufacturer?: string;         // Optional
  serialNumber?: string;         // Optional
  lastSyncDateTime?: Date;       // Optional
}
```

**Safe Access**:
```typescript
// Optional chaining
const manufacturer = device.manufacturer?.toUpperCase() ?? 'UNKNOWN';

// Null coalescing
const syncDate = device.lastSyncDateTime ?? new Date();
```

## Testing Strategy

### Test Coverage

**File**: `/src/reports/configmgr/mobile-device-inventory-reports.test.ts`

**Coverage Areas**:
1. All public report methods
2. Filtering and sorting
3. Pagination
4. Data aggregation
5. Error handling
6. Edge cases
7. Convenience functions

### Test Structure

```typescript
describe('MobileDeviceInventoryReports', () => {
  let report: MobileDeviceInventoryReports;
  let mockClient: MockGraphClient;
  let config: AppConfig;

  beforeEach(() => {
    // Setup test environment
  });

  describe('getAllCorporateOwnedDevices', () => {
    it('should retrieve all corporate-owned mobile devices', async () => {
      // Test implementation
    });
  });

  // ... more test suites
});
```

### Mock Data Strategy

**Mock Graph Client**:
```typescript
class MockGraphClient {
  private responses: Map<string, any> = new Map();

  api(endpoint: string) {
    return {
      filter: (filter: string) => this,
      select: (fields: string[]) => this,
      get: async () => {
        // Return appropriate mock data based on filter
      }
    };
  }
}
```

**Mock Device Data**:
```typescript
const mockDevice1 = {
  id: 'device-001',
  deviceName: 'iPhone 14 Pro',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  managedDeviceOwnerType: 'company',
  // ... more fields
};
```

### Test Categories

**1. Basic Functionality Tests**:
```typescript
it('should retrieve all corporate-owned mobile devices')
it('should retrieve all mobile device clients')
it('should group devices by operating system')
```

**2. Filtering Tests**:
```typescript
it('should filter by operating system')
it('should filter by compliance state')
it('should filter by supervised status')
```

**3. Data Transformation Tests**:
```typescript
it('should parse device data correctly')
it('should calculate communication status')
it('should calculate days since last contact')
```

**4. Pagination Tests**:
```typescript
it('should paginate results correctly')
it('should handle last page correctly')
it('should handle empty results')
```

**5. Error Handling Tests**:
```typescript
it('should handle API errors gracefully')
it('should handle empty device list')
it('should handle devices with missing optional fields')
```

### Running Tests

```bash
# Run all tests
npm test

# Run specific test file
npm test mobile-device-inventory-reports.test.ts

# Run with coverage
npm run test:coverage

# Run in watch mode
npm run test:watch
```

## Performance Considerations

### API Call Optimization

**1. Field Selection**:
```typescript
// Good - only request needed fields
.select([
  'id',
  'deviceName',
  'operatingSystem',
  'osVersion',
  'lastSyncDateTime'
])

// Bad - request all fields
.get() // Returns all fields (larger payload)
```

**2. Server-Side Filtering**:
```typescript
// Good - filter at API level
const filter = "managedDeviceOwnerType eq 'company'";
await this.graphClient
  .api('/deviceManagement/managedDevices')
  .filter(filter)
  .get();

// Bad - fetch all and filter client-side
const allDevices = await this.graphClient
  .api('/deviceManagement/managedDevices')
  .get();
const corporate = allDevices.filter(d => d.ownerType === 'company');
```

**3. Pagination**:
```typescript
// Automatic pagination with getAllPages
const response = await this.graphClient.api('/deviceManagement/managedDevices').get();
const allDevices = await this.getAllPages<Device>(response);
```

### Memory Management

**Large Dataset Handling**:
```typescript
// Stream processing for very large datasets
async function* streamDevices() {
  let nextLink = '/deviceManagement/managedDevices';

  while (nextLink) {
    const response = await graphClient.api(nextLink).get();
    yield* response.value;
    nextLink = response['@odata.nextLink'];
  }
}

// Process in batches
for await (const device of streamDevices()) {
  processDevice(device);
}
```

**Client-Side Pagination**:
```typescript
// Use pagination to limit memory footprint
const page = report.paginate(allDevices, { page: 1, pageSize: 100 });
// Only loads 100 devices into memory at a time
```

### Caching Strategy

**Implementation**:
```typescript
class CachedMobileDeviceInventoryReports extends MobileDeviceInventoryReports {
  private cache: Map<string, { data: any; timestamp: number }> = new Map();
  private cacheTTL = 5 * 60 * 1000; // 5 minutes

  async getAllMobileDevices(options?: MobileDeviceFilterOptions) {
    const cacheKey = JSON.stringify(options);
    const cached = this.cache.get(cacheKey);

    if (cached && (Date.now() - cached.timestamp) < this.cacheTTL) {
      return cached.data;
    }

    const data = await super.getAllMobileDevices(options);
    this.cache.set(cacheKey, { data, timestamp: Date.now() });
    return data;
  }
}
```

### Rate Limiting

**Built-in Retry Logic**:
```typescript
// Exponential backoff automatically handles rate limiting
protected async retryGraphCall<T>(
  apiCall: () => Promise<T>,
  maxRetries: number = 3,
  baseDelay: number = 1000
): Promise<T>
```

**Throttling for Batch Operations**:
```typescript
// Process in controlled batches
const batchSize = 10;
const delayBetweenBatches = 1000; // 1 second

for (let i = 0; i < items.length; i += batchSize) {
  const batch = items.slice(i, i + batchSize);
  await Promise.all(batch.map(item => processItem(item)));

  if (i + batchSize < items.length) {
    await new Promise(resolve => setTimeout(resolve, delayBetweenBatches));
  }
}
```

## Error Handling

### Error Types

**1. Authentication Errors**:
```typescript
// HTTP 401 - Unauthorized
// HTTP 403 - Forbidden (insufficient permissions)

catch (error) {
  if (error.statusCode === 401) {
    logger.error('Authentication failed');
    // Trigger token refresh
  } else if (error.statusCode === 403) {
    logger.error('Insufficient permissions');
    // Request additional permissions
  }
}
```

**2. Rate Limiting Errors**:
```typescript
// HTTP 429 - Too Many Requests

// Handled automatically by retryGraphCall
// Implements exponential backoff
```

**3. Server Errors**:
```typescript
// HTTP 5xx - Server errors

// Handled automatically by retryGraphCall
// Retries up to 3 times with exponential backoff
```

**4. Network Errors**:
```typescript
catch (error) {
  if (error.code === 'ETIMEDOUT' || error.code === 'ECONNRESET') {
    logger.error('Network error', error);
    // Retry or fail gracefully
  }
}
```

### Error Propagation

**Pattern**:
```typescript
async getAllCorporateOwnedDevices() {
  try {
    // Implementation
  } catch (error) {
    logger.error('Failed to fetch corporate devices', error);
    throw error; // Propagate to caller
  }
}
```

**Caller Handling**:
```typescript
try {
  const devices = await report.getAllCorporateOwnedDevices();
} catch (error) {
  console.error('Error:', error.message);
  // Display user-friendly message
  // Log to monitoring system
  // Trigger alert if critical
}
```

### Validation

**Input Validation**:
```typescript
paginate<T>(data: T[], options: PaginationOptions): PaginatedResult<T> {
  if (!data || !Array.isArray(data)) {
    throw new Error('Data must be an array');
  }

  if (options.page < 1) {
    throw new Error('Page number must be >= 1');
  }

  if (options.pageSize < 1) {
    throw new Error('Page size must be >= 1');
  }

  // Implementation
}
```

**Data Validation**:
```typescript
private mapOwnerType(ownerType: string | undefined): DeviceOwnershipType {
  if (!ownerType) return DeviceOwnershipType.UNKNOWN;

  const normalized = ownerType.toLowerCase();
  if (normalized.includes('company')) return DeviceOwnershipType.COMPANY;
  if (normalized.includes('personal')) return DeviceOwnershipType.PERSONAL;

  // Default to UNKNOWN for unrecognized values
  return DeviceOwnershipType.UNKNOWN;
}
```

## Extending the Module

### Adding New Reports

**Step 1: Define Interface**:
```typescript
export interface CustomDeviceReport extends MobileDeviceInfo {
  customField1: string;
  customField2: number;
}
```

**Step 2: Implement Method**:
```typescript
async getCustomDeviceReport(
  options?: MobileDeviceFilterOptions
): Promise<CustomDeviceReport[]> {
  logger.info('Fetching custom device report');

  try {
    const filter = this.buildODataFilter(options);

    const response = await this.retryGraphCall(() =>
      this.graphClient
        .api('/deviceManagement/managedDevices')
        .filter(filter)
        .select([...fields])
        .get()
    );

    const devices = await this.getAllPages<any>(response);

    return devices.map(device => ({
      // Map fields
    }));
  } catch (error) {
    logger.error('Failed to fetch custom report', error);
    throw error;
  }
}
```

**Step 3: Add Tests**:
```typescript
describe('getCustomDeviceReport', () => {
  it('should retrieve custom device report', async () => {
    const report = await report.getCustomDeviceReport();
    expect(report).toBeDefined();
  });
});
```

### Adding New Filters

**Step 1: Extend Interface**:
```typescript
export interface MobileDeviceFilterOptions {
  // ... existing filters
  newFilter?: string;
}
```

**Step 2: Update buildODataFilter**:
```typescript
private buildODataFilter(options: MobileDeviceFilterOptions): string {
  const filters: string[] = [];

  // ... existing filters

  if (options.newFilter) {
    filters.push(`newField eq '${options.newFilter}'`);
  }

  return filters.join(' and ');
}
```

### Adding New Export Formats

**Step 1: Update Interface**:
```typescript
export interface ExportOptions {
  format: 'json' | 'csv' | 'html' | 'xlsx'; // Add new format
  outputDirectory: string;
  includeTimestamp?: boolean;
}
```

**Step 2: Implement in OutputFormatter**:
```typescript
// In src/formatters/output-formatter.ts
static async format(
  data: ReportData,
  format: 'json' | 'csv' | 'html' | 'xlsx',
  outputDir: string,
  includeTimestamp: boolean
): Promise<string> {
  switch (format) {
    case 'xlsx':
      return this.formatAsExcel(data, outputDir, includeTimestamp);
    // ... other formats
  }
}
```

### Adding Custom Aggregations

**Example: Device Count by Manufacturer**:
```typescript
async getDevicesByManufacturer(): Promise<Map<string, number>> {
  const devices = await this.getAllMobileDevices();

  const manufacturerCount = new Map<string, number>();
  devices.forEach(device => {
    const manufacturer = device.manufacturer || 'Unknown';
    manufacturerCount.set(
      manufacturer,
      (manufacturerCount.get(manufacturer) || 0) + 1
    );
  });

  return manufacturerCount;
}
```

## Deployment Notes

### Environment Configuration

**Required Environment Variables**:
```bash
# Azure AD App Registration
TENANT_ID=your-tenant-id
CLIENT_ID=your-client-id
CLIENT_SECRET=your-client-secret

# Graph API
GRAPH_API_VERSION=v1.0
GRAPH_API_ENDPOINT=https://graph.microsoft.com

# Output Configuration
OUTPUT_DIRECTORY=/var/reports
INCLUDE_TIMESTAMP=true

# Logging
LOG_LEVEL=info
LOG_FILE=/var/logs/intune-reporting.log
```

### Azure AD App Registration

**Required API Permissions**:
1. Navigate to Azure Portal → Azure Active Directory → App registrations
2. Select your app or create new app registration
3. Go to "API permissions"
4. Add permissions:
   - Microsoft Graph → Application permissions → `DeviceManagementManagedDevices.Read.All`
5. Grant admin consent

**Authentication Setup**:
```typescript
import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

const credential = new ClientSecretCredential(
  process.env.TENANT_ID!,
  process.env.CLIENT_ID!,
  process.env.CLIENT_SECRET!
);

const authProvider = new TokenCredentialAuthenticationProvider(
  credential,
  { scopes: ['https://graph.microsoft.com/.default'] }
);

const graphClient = Client.initWithMiddleware({ authProvider });
```

### Production Considerations

**1. Logging**:
```typescript
// Configure production logging
const logger = Logger.getInstance({
  level: 'info',
  file: '/var/logs/mobile-device-inventory.log',
  maxSize: '100m',
  maxFiles: 10
});
```

**2. Error Monitoring**:
```typescript
// Integrate with monitoring service
process.on('unhandledRejection', (error) => {
  logger.error('Unhandled promise rejection', error);
  // Send to monitoring service (e.g., Sentry, AppInsights)
});
```

**3. Performance Monitoring**:
```typescript
// Track execution time
const startTime = Date.now();
const devices = await report.getAllMobileDevices();
const executionTime = Date.now() - startTime;

logger.info('Report execution time', {
  method: 'getAllMobileDevices',
  executionTime,
  deviceCount: devices.length
});
```

**4. Scheduled Execution**:
```typescript
import cron from 'node-cron';

// Run daily at 6 AM
cron.schedule('0 6 * * *', async () => {
  try {
    logger.info('Starting scheduled mobile device inventory report');
    const report = new MobileDeviceInventoryReports(graphClient, config);
    const data = await report.execute();

    await report.exportReport(data, {
      format: 'json',
      outputDirectory: '/var/reports',
      includeTimestamp: true
    });

    logger.info('Scheduled report completed successfully');
  } catch (error) {
    logger.error('Scheduled report failed', error);
  }
});
```

### Security Best Practices

**1. Credential Management**:
```typescript
// Use Azure Key Vault for secrets
import { SecretClient } from '@azure/keyvault-secrets';
import { DefaultAzureCredential } from '@azure/identity';

const credential = new DefaultAzureCredential();
const client = new SecretClient(vaultUrl, credential);

const clientSecret = await client.getSecret('client-secret');
```

**2. Least Privilege Access**:
- Request only `Read.All` permissions (not `ReadWrite.All`)
- Use service principals with limited scope
- Rotate client secrets regularly

**3. Data Protection**:
```typescript
// Sanitize sensitive data before logging
function sanitizeDeviceData(device: MobileDeviceInfo) {
  return {
    deviceId: device.deviceId,
    operatingSystem: device.operatingSystem,
    // Exclude: phoneNumber, imei, meid, serialNumber
  };
}

logger.info('Device processed', sanitizeDeviceData(device));
```

**4. Audit Logging**:
```typescript
// Log all report executions
logger.audit({
  action: 'REPORT_EXECUTION',
  reportName: 'mobile-device-inventory',
  user: context.user,
  timestamp: new Date().toISOString(),
  recordCount: devices.length
});
```

### Health Checks

**Implementation**:
```typescript
export async function healthCheck(): Promise<HealthCheckResult> {
  try {
    // Test Graph API connectivity
    await graphClient.api('/deviceManagement/managedDevices')
      .top(1)
      .get();

    return {
      status: 'healthy',
      timestamp: new Date().toISOString(),
      services: {
        graphApi: 'up',
        authentication: 'ok'
      }
    };
  } catch (error) {
    return {
      status: 'unhealthy',
      timestamp: new Date().toISOString(),
      error: error.message
    };
  }
}
```

---

## Summary

This implementation guide provides comprehensive technical details for working with the Mobile Device Inventory Reports module. Key takeaways:

1. **Architecture**: Built on BaseReport, integrates with Microsoft Graph API
2. **ConfigMgr Compatibility**: Replicates 4 key ConfigMgr reports using Graph API
3. **Type Safety**: Comprehensive TypeScript interfaces and enums
4. **Performance**: Optimized with field selection, server-side filtering, and pagination
5. **Testing**: Full unit test coverage with mock data
6. **Extensibility**: Well-structured for adding new reports and features
7. **Production Ready**: Includes error handling, logging, monitoring, and security

For usage examples and API reference, see the main documentation: [CONFIGMGR_MOBILE_INVENTORY.md](/docs/CONFIGMGR_MOBILE_INVENTORY.md)

---

**Last Updated**: February 7, 2026
**Module Version**: 1.0.0
**Author**: Intune Reporting Dashboard Team
