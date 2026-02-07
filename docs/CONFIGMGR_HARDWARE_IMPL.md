# Configuration Manager Hardware Reports - Implementation Guide

This document provides detailed implementation guidance for the Device Hardware Reports module, including architecture, design patterns, customization options, and integration strategies.

## Table of Contents

- [Architecture Overview](#architecture-overview)
- [Implementation Details](#implementation-details)
- [Data Flow](#data-flow)
- [Report Implementation](#report-implementation)
- [Customization Guide](#customization-guide)
- [Integration Patterns](#integration-patterns)
- [Testing Strategy](#testing-strategy)
- [Performance Optimization](#performance-optimization)
- [Extension Points](#extension-points)
- [Code Examples](#code-examples)

## Architecture Overview

### Module Structure

```
src/reports/configmgr/
├── device-hardware-reports.ts      # Main implementation
├── device-hardware-reports.test.ts # Unit tests
└── examples/
    └── device-hardware-reports-examples.ts

src/reports/base-report.ts          # Base class
src/formatters/output-formatter.ts  # Export functionality
src/auth/graph-auth.ts              # Authentication
```

### Class Hierarchy

```
BaseReport (Abstract)
    ├── Provides: retryGraphCall(), getAllPages(), createMetadata()
    ├── Handles: Pagination, retry logic, error handling
    │
    └── DeviceHardwareReports (Concrete)
        ├── Core Methods: getAllDeviceHardwareInfo()
        ├── Report Methods: 8 ConfigMgr report equivalents
        ├── Helper Methods: Filtering, categorization, statistics
        └── Export Methods: Multi-format export capability
```

### Design Patterns

#### 1. Repository Pattern
- Separates data access from business logic
- Graph API calls isolated in `getAllDeviceHardwareInfo()`
- Enables mocking and testing

#### 2. Template Method Pattern
- Base class defines common report structure
- Subclasses implement specific report logic
- Consistent error handling and pagination

#### 3. Strategy Pattern
- Different export formats (JSON, CSV, HTML)
- Configurable threshold strategies
- Platform-specific filtering

#### 4. Builder Pattern
- Comprehensive report generation
- Incremental data aggregation
- Optional filtering at each stage

## Implementation Details

### Report Implementations

#### Report 9: Display Configuration Summary

**Implementation**: `getDisplayConfigurationSummary()`

**Algorithm**:
1. Fetch all devices or use cached list
2. Group devices by display tier (based on RAM as proxy)
3. Calculate device counts and percentages
4. Sort by device count descending
5. Return categorized summary

**Display Categories**:
```typescript
private categorizeDisplayByMemory(memoryGB: number): string {
  if (memoryGB < 2) return 'Standard (< 2GB RAM)';
  if (memoryGB < 4) return 'Enhanced (2-4GB RAM)';
  if (memoryGB < 6) return 'High (4-6GB RAM)';
  if (memoryGB < 8) return 'Premium (6-8GB RAM)';
  return 'Ultra (8GB+ RAM)';
}
```

**Rationale**: Mobile devices don't expose direct display configuration via Graph API. RAM is used as a proxy for device tier, which correlates with display quality.

#### Report 10: OS Distribution Summary

**Implementation**: `getOSDistributionSummary()`

**Algorithm**:
1. Fetch all devices
2. Group by OS name and version (combined key)
3. Calculate counts and percentages per OS/version
4. Sort by device count descending
5. Return distribution array

**Key Features**:
- Tracks both OS platform and version
- Identifies version fragmentation
- Supports compliance tracking

**Data Structure**:
```typescript
interface OSDistributionSummary {
  operatingSystem: string;    // e.g., "iOS"
  osVersion: string;           // e.g., "17.2.1"
  deviceCount: number;         // Device count
  percentage: number;          // Percentage of total
  devices: string[];           // Device names
}
```

#### Report 11: Memory Range Summary

**Implementation**: `getMemoryRangeSummary()`

**Algorithm**:
1. Define memory ranges in MB
2. Categorize each device into appropriate range
3. Count devices per range
4. Calculate percentages
5. Return summary for all ranges (including empty)

**Memory Ranges**:
```typescript
const ranges = [
  { min: 0, max: 1024, label: '< 1 GB' },
  { min: 1024, max: 2048, label: '1-2 GB' },
  { min: 2048, max: 4096, label: '2-4 GB' },
  { min: 4096, max: 6144, label: '4-6 GB' },
  { min: 6144, max: 8192, label: '6-8 GB' },
  { min: 8192, max: 16384, label: '8-16 GB' },
  { min: 16384, max: Infinity, label: '> 16 GB' },
];
```

**Use Case**: Identify device capabilities for memory-intensive app deployments.

#### Report 12: Storage Range Summary

**Implementation**: `getStorageRangeSummary()`

**Algorithm**: Similar to memory ranges, but for storage capacity

**Storage Ranges**:
```typescript
const ranges = [
  { min: 0, max: 16, label: '< 16 GB' },
  { min: 16, max: 32, label: '16-32 GB' },
  { min: 32, max: 64, label: '32-64 GB' },
  { min: 64, max: 128, label: '64-128 GB' },
  { min: 128, max: 256, label: '128-256 GB' },
  { min: 256, max: 512, label: '256-512 GB' },
  { min: 512, max: 1024, label: '512 GB - 1 TB' },
  { min: 1024, max: Infinity, label: '> 1 TB' },
];
```

#### Report 25: Specific Free Memory

**Implementation**: `getDevicesWithSpecificFreeMemory()`

**Algorithm**:
1. Fetch all devices
2. Estimate free memory (40% of total)
3. Filter devices within specified range
4. Return matching devices

**Important Note**: Graph API doesn't provide real-time free memory for mobile devices. The implementation uses a 40% estimation as a proxy. This is suitable for trending but not absolute measurements.

**Estimation Logic**:
```typescript
const totalMemoryMB = device.memory.totalMemoryInMB;
const estimatedFreeMemoryMB = Math.round(totalMemoryMB * 0.4);
const estimatedFreePercentage = 40;
```

#### Report 26: Specific Free Storage

**Implementation**: `getDevicesWithSpecificFreeStorage()`

**Algorithm**:
1. Fetch all devices
2. Extract actual free storage from Graph API
3. Filter devices within specified range
4. Return matching devices with storage details

**Accuracy**: Storage data is accurate and updated during device sync.

**Data Available**:
- Total storage (accurate)
- Free storage (accurate)
- Used storage (calculated)
- Free percentage (calculated)

#### Report 28: Low Memory Devices

**Implementation**: `getDevicesWithLowMemory()`

**Algorithm**:
1. Uses `getDevicesWithSpecificFreeMemory()` internally
2. Sets range from 0 to threshold
3. Returns devices below threshold

**Default Threshold**: 512 MB (configurable)

**Use Cases**:
- Performance troubleshooting
- Proactive user communication
- Device refresh planning

#### Report 29: Low Storage Devices

**Implementation**: `getDevicesWithLowStorage()`

**Algorithm**:
1. Fetch all devices
2. Filter devices with free storage < threshold
3. Sort by free storage ascending (most critical first)
4. Return sorted list

**Default Threshold**: 10 GB (configurable)

**Sorting**: Results are sorted with most critical devices first, enabling prioritized action.

## Data Flow

### Complete Report Generation Flow

```
┌─────────────────────────────────────────────────────────┐
│ 1. Authentication                                       │
│    - GraphAuth.getAuthenticatedClient()                │
│    - Azure AD OAuth token acquisition                  │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│ 2. Initialize Reporter                                  │
│    - new DeviceHardwareReports(client, config)         │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│ 3. Fetch Device Data                                    │
│    - GET /deviceManagement/managedDevices              │
│    - Select required properties                         │
│    - Handle pagination (@odata.nextLink)               │
│    - Retry on failures (exponential backoff)           │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│ 4. Transform Data                                       │
│    - Parse raw Graph API response                      │
│    - Calculate memory in MB/GB                         │
│    - Calculate storage in MB/GB                        │
│    - Compute percentages and utilization               │
│    - Map to DeviceHardwareInfo interface               │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│ 5. Generate Reports                                     │
│    - Display configuration summary                      │
│    - OS distribution analysis                           │
│    - Memory range categorization                        │
│    - Storage range categorization                       │
│    - Low memory device identification                   │
│    - Low storage device identification                  │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│ 6. Aggregate Summary                                    │
│    - Calculate totals and averages                      │
│    - Identify top device models                         │
│    - Generate alert counts                              │
│    - Compile comprehensive summary                      │
└─────────────────┬───────────────────────────────────────┘
                  │
                  ▼
┌─────────────────────────────────────────────────────────┐
│ 7. Export Report                                        │
│    - Format as JSON/CSV/HTML                           │
│    - Write to output directory                         │
│    - Include timestamp if configured                    │
└─────────────────────────────────────────────────────────┘
```

### API Call Sequence

```typescript
// Single-report flow
getAllDeviceHardwareInfo()
    ├─→ retryGraphCall()
    │   ├─→ graphClient.api('/deviceManagement/managedDevices')
    │   └─→ Exponential backoff on failures
    │
    ├─→ getAllPages()
    │   └─→ Follow @odata.nextLink until complete
    │
    └─→ Transform to DeviceHardwareInfo[]

// Comprehensive report flow (optimized)
getAllDeviceHardwareInfo()  // Single API call
    ├─→ getDisplayConfigurationSummary(cachedDevices)
    ├─→ getOSDistributionSummary(cachedDevices)
    ├─→ getMemoryRangeSummary(cachedDevices)
    ├─→ getStorageRangeSummary(cachedDevices)
    ├─→ getDevicesWithLowMemory(512, cachedDevices)
    └─→ getDevicesWithLowStorage(10, cachedDevices)
```

## Customization Guide

### Adding Custom Memory Ranges

**Scenario**: Your organization wants different memory categorization.

**Implementation**:

```typescript
// Create a custom method
async getCustomMemoryRangeSummary(
  devices?: DeviceHardwareInfo[]
): Promise<MemoryRangeSummary[]> {
  if (!devices) {
    devices = await this.getAllDeviceHardwareInfo();
  }

  // Define custom ranges
  const customRanges = [
    { min: 0, max: 512, label: 'Very Low (< 512MB)' },
    { min: 512, max: 2048, label: 'Low (0.5-2GB)' },
    { min: 2048, max: 8192, label: 'Medium (2-8GB)' },
    { min: 8192, max: Infinity, label: 'High (> 8GB)' },
  ];

  // Group devices by range (same logic as getMemoryRangeSummary)
  const rangeGroups = new Map<string, DeviceHardwareInfo[]>();

  devices.forEach((device) => {
    const memoryMB = device.memory.totalMemoryInMB;
    for (const range of customRanges) {
      if (memoryMB >= range.min && memoryMB < range.max) {
        if (!rangeGroups.has(range.label)) {
          rangeGroups.set(range.label, []);
        }
        rangeGroups.get(range.label)!.push(device);
        break;
      }
    }
  });

  // Convert to summary array
  const totalDevices = devices.length;
  const summary: MemoryRangeSummary[] = [];

  customRanges.forEach((range) => {
    const groupDevices = rangeGroups.get(range.label) || [];
    summary.push({
      range: range.label,
      rangeInMB: `${range.min}-${range.max === Infinity ? '∞' : range.max} MB`,
      deviceCount: groupDevices.length,
      percentage: totalDevices > 0
        ? Math.round((groupDevices.length / totalDevices) * 100 * 100) / 100
        : 0,
      devices: groupDevices.map((d) => d.deviceName),
    });
  });

  return summary;
}
```

### Adding Custom Storage Thresholds by OS

**Scenario**: Different storage thresholds for iOS vs Android vs Windows.

**Implementation**:

```typescript
async getDevicesWithLowStorageByPlatform(): Promise<{
  ios: DeviceStorageThreshold[];
  android: DeviceStorageThreshold[];
  windows: DeviceStorageThreshold[];
}> {
  const allDevices = await this.getAllDeviceHardwareInfo();

  // Separate by platform
  const iosDevices = allDevices.filter(d =>
    d.operatingSystem.operatingSystem.includes('iOS')
  );
  const androidDevices = allDevices.filter(d =>
    d.operatingSystem.operatingSystem.includes('Android')
  );
  const windowsDevices = allDevices.filter(d =>
    d.operatingSystem.operatingSystem.includes('Windows')
  );

  // Apply platform-specific thresholds
  return {
    ios: await this.getDevicesWithLowStorage(5, iosDevices),      // 5GB for iOS
    android: await this.getDevicesWithLowStorage(8, androidDevices), // 8GB for Android
    windows: await this.getDevicesWithLowStorage(20, windowsDevices), // 20GB for Windows
  };
}
```

### Adding Device Age Analysis

**Scenario**: Combine hardware data with enrollment date to identify old devices.

**Implementation**:

```typescript
interface DeviceAgeAnalysis {
  deviceId: string;
  deviceName: string;
  manufacturer: string;
  model: string;
  enrolledDate: Date;
  ageInMonths: number;
  memory: number;
  storage: number;
  needsRefresh: boolean;
}

async getDeviceAgeAnalysis(): Promise<DeviceAgeAnalysis[]> {
  const devices = await this.getAllDeviceHardwareInfo();
  const now = new Date();

  return devices.map(device => {
    const enrolled = new Date(device.enrolledDateTime);
    const ageInMonths = Math.floor(
      (now.getTime() - enrolled.getTime()) / (1000 * 60 * 60 * 24 * 30)
    );

    // Define refresh criteria: > 36 months old AND (< 4GB RAM OR < 64GB storage)
    const needsRefresh =
      ageInMonths > 36 &&
      (device.memory.totalMemoryInGB < 4 || device.storage.totalStorageInGB < 64);

    return {
      deviceId: device.deviceId,
      deviceName: device.deviceName,
      manufacturer: device.manufacturer,
      model: device.model,
      enrolledDate: enrolled,
      ageInMonths,
      memory: device.memory.totalMemoryInGB,
      storage: device.storage.totalStorageInGB,
      needsRefresh,
    };
  }).sort((a, b) => b.ageInMonths - a.ageInMonths);
}
```

### Adding Manufacturer-Specific Reports

**Scenario**: Generate report for specific manufacturer with custom metrics.

**Implementation**:

```typescript
async getManufacturerHardwareReport(manufacturer: string): Promise<{
  manufacturer: string;
  totalDevices: number;
  models: Map<string, number>;
  averageMemory: number;
  averageStorage: number;
  oldestDevice: DeviceHardwareInfo;
  newestDevice: DeviceHardwareInfo;
}> {
  const devices = await this.getAllDeviceHardwareInfo({
    manufacturer,
  });

  // Count devices per model
  const models = new Map<string, number>();
  devices.forEach(d => {
    models.set(d.model, (models.get(d.model) || 0) + 1);
  });

  // Calculate averages
  const avgMemory = devices.reduce((sum, d) =>
    sum + d.memory.totalMemoryInGB, 0
  ) / devices.length;

  const avgStorage = devices.reduce((sum, d) =>
    sum + d.storage.totalStorageInGB, 0
  ) / devices.length;

  // Find oldest and newest
  const sortedByDate = [...devices].sort((a, b) =>
    new Date(a.enrolledDateTime).getTime() - new Date(b.enrolledDateTime).getTime()
  );

  return {
    manufacturer,
    totalDevices: devices.length,
    models,
    averageMemory: Math.round(avgMemory * 100) / 100,
    averageStorage: Math.round(avgStorage * 100) / 100,
    oldestDevice: sortedByDate[0],
    newestDevice: sortedByDate[sortedByDate.length - 1],
  };
}
```

## Integration Patterns

### Pattern 1: Automated Monitoring Dashboard

**Use Case**: Real-time dashboard showing hardware health.

**Implementation**:

```typescript
import { DeviceHardwareReports } from './reports/configmgr/device-hardware-reports';

class HardwareMonitoringDashboard {
  private reporter: DeviceHardwareReports;
  private refreshInterval: number = 3600000; // 1 hour

  constructor(reporter: DeviceHardwareReports) {
    this.reporter = reporter;
  }

  async start() {
    // Initial load
    await this.updateMetrics();

    // Periodic refresh
    setInterval(() => this.updateMetrics(), this.refreshInterval);
  }

  private async updateMetrics() {
    try {
      const report = await this.reporter.generateComprehensiveReport();

      const metrics = {
        timestamp: new Date().toISOString(),
        totalDevices: report.summary.totalDevices,
        lowMemoryDevices: report.summary.lowMemoryDeviceCount,
        lowStorageDevices: report.summary.lowStorageDeviceCount,
        averageMemory: report.summary.averageMemoryGB,
        averageStorage: report.summary.averageStorageGB,
        averageFreeStorage: report.summary.averageFreeStoragePercentage,
        osPlatforms: report.summary.osDistribution.length,
      };

      // Send to monitoring system
      await this.sendToMonitoring(metrics);

      // Check alert thresholds
      if (metrics.lowStorageDevices > 10) {
        await this.sendAlert('High number of low storage devices', metrics);
      }

      if (metrics.lowMemoryDevices > 5) {
        await this.sendAlert('Multiple devices with low memory', metrics);
      }
    } catch (error) {
      console.error('Dashboard update failed:', error);
    }
  }

  private async sendToMonitoring(metrics: any) {
    // Implementation depends on monitoring platform
    // Examples: Prometheus, Datadog, Application Insights
  }

  private async sendAlert(message: string, context: any) {
    // Send to alerting system
    // Examples: PagerDuty, OpsGenie, email, Teams webhook
  }
}
```

### Pattern 2: Scheduled Report Delivery

**Use Case**: Weekly hardware reports emailed to IT team.

**Implementation**:

```typescript
import { CronJob } from 'cron';
import { createTransport } from 'nodemailer';

class ScheduledHardwareReports {
  private reporter: DeviceHardwareReports;
  private emailTransporter: any;

  constructor(reporter: DeviceHardwareReports) {
    this.reporter = reporter;
    this.emailTransporter = createTransport({
      // Email configuration
    });
  }

  startScheduledReports() {
    // Weekly comprehensive report (Monday 8 AM)
    new CronJob('0 8 * * 1', async () => {
      await this.sendWeeklyReport();
    }).start();

    // Daily low storage alert (Every day 9 AM)
    new CronJob('0 9 * * *', async () => {
      await this.sendLowStorageAlert();
    }).start();

    // Monthly device inventory (1st of month, 10 AM)
    new CronJob('0 10 1 * *', async () => {
      await this.sendMonthlyInventory();
    }).start();
  }

  private async sendWeeklyReport() {
    const report = await this.reporter.generateComprehensiveReport();

    // Export to HTML
    const htmlPath = await this.reporter.exportReport(report, {
      format: 'html',
      outputDir: './reports/weekly',
      includeTimestamp: true,
    });

    // Send email
    await this.emailTransporter.sendMail({
      to: 'it-team@company.com',
      subject: 'Weekly Hardware Inventory Report',
      html: fs.readFileSync(htmlPath, 'utf8'),
    });
  }

  private async sendLowStorageAlert() {
    const lowStorage = await this.reporter.getDevicesWithLowStorage(10);

    if (lowStorage.length > 0) {
      const criticalDevices = lowStorage.filter(d => d.freeStorageInGB < 5);

      const emailBody = `
        <h2>Low Storage Alert</h2>
        <p>Total devices with low storage: ${lowStorage.length}</p>
        <p>Critical devices (<5GB): ${criticalDevices.length}</p>

        <h3>Most Critical Devices:</h3>
        <table>
          <tr>
            <th>Device</th>
            <th>User</th>
            <th>Free Storage</th>
            <th>Total Storage</th>
          </tr>
          ${criticalDevices.slice(0, 10).map(d => `
            <tr>
              <td>${d.deviceName}</td>
              <td>${d.userPrincipalName}</td>
              <td>${d.freeStorageInGB} GB</td>
              <td>${d.totalStorageInGB} GB</td>
            </tr>
          `).join('')}
        </table>
      `;

      await this.emailTransporter.sendMail({
        to: 'helpdesk@company.com',
        subject: `ALERT: ${criticalDevices.length} devices critically low on storage`,
        html: emailBody,
      });
    }
  }

  private async sendMonthlyInventory() {
    const report = await this.reporter.generateComprehensiveReport();

    // Export all formats
    const jsonPath = await this.reporter.exportReport(report, {
      format: 'json',
      outputDir: './reports/monthly',
      includeTimestamp: true,
    });

    const csvPath = await this.reporter.exportReport(report, {
      format: 'csv',
      outputDir: './reports/monthly',
      includeTimestamp: true,
    });

    const htmlPath = await this.reporter.exportReport(report, {
      format: 'html',
      outputDir: './reports/monthly',
      includeTimestamp: true,
    });

    // Archive to cloud storage
    await this.archiveToCloud([jsonPath, csvPath, htmlPath]);

    // Send executive summary
    await this.emailTransporter.sendMail({
      to: 'executives@company.com',
      subject: 'Monthly Device Hardware Inventory',
      html: fs.readFileSync(htmlPath, 'utf8'),
      attachments: [
        { filename: 'report.json', path: jsonPath },
        { filename: 'report.csv', path: csvPath },
      ],
    });
  }

  private async archiveToCloud(files: string[]) {
    // Upload to Azure Blob Storage, AWS S3, etc.
  }
}
```

### Pattern 3: ServiceNow Integration

**Use Case**: Create ServiceNow incidents for devices with critical issues.

**Implementation**:

```typescript
import axios from 'axios';

class ServiceNowIntegration {
  private reporter: DeviceHardwareReports;
  private snowUrl: string;
  private snowAuth: string;

  constructor(reporter: DeviceHardwareReports, snowUrl: string, auth: string) {
    this.reporter = reporter;
    this.snowUrl = snowUrl;
    this.snowAuth = auth;
  }

  async checkAndCreateIncidents() {
    // Find critical devices
    const lowStorage = await this.reporter.getDevicesWithLowStorage(5);
    const lowMemory = await this.reporter.getDevicesWithLowMemory(256);

    // Create incidents for critical devices
    for (const device of lowStorage) {
      await this.createIncident({
        short_description: `Device ${device.deviceName} critically low on storage`,
        description: `
          Device: ${device.deviceName}
          User: ${device.userPrincipalName}
          Free Storage: ${device.freeStorageInGB} GB (${device.freeStoragePercentage.toFixed(2)}%)
          Total Storage: ${device.totalStorageInGB} GB

          Action Required: Contact user to free up storage space or consider device refresh.
        `,
        urgency: device.freeStorageInGB < 2 ? 1 : 2, // High urgency if < 2GB
        category: 'Hardware',
        subcategory: 'Storage',
        assigned_to: 'helpdesk-group',
      });
    }

    for (const device of lowMemory) {
      await this.createIncident({
        short_description: `Device ${device.deviceName} critically low on memory`,
        description: `
          Device: ${device.deviceName}
          User: ${device.userPrincipalName}
          Estimated Free Memory: ${device.freeMemoryInMB} MB
          Total Memory: ${device.totalMemoryInMB} MB

          Action Required: Performance issues likely. Consider device refresh.
        `,
        urgency: 2,
        category: 'Hardware',
        subcategory: 'Memory',
        assigned_to: 'helpdesk-group',
      });
    }
  }

  private async createIncident(incident: any) {
    try {
      await axios.post(
        `${this.snowUrl}/api/now/table/incident`,
        incident,
        {
          headers: {
            'Authorization': this.snowAuth,
            'Content-Type': 'application/json',
          },
        }
      );
    } catch (error) {
      console.error('Failed to create ServiceNow incident:', error);
    }
  }
}
```

## Testing Strategy

### Unit Testing Approach

The module includes comprehensive unit tests covering:

1. **Data Retrieval Tests**
   - Verify Graph API calls
   - Test pagination handling
   - Validate data transformation

2. **Report Generation Tests**
   - Test each report method independently
   - Verify calculations and categorizations
   - Check sorting and filtering

3. **Filter Tests**
   - Test each filter option
   - Test combined filters
   - Verify filter edge cases

4. **Threshold Tests**
   - Test default thresholds
   - Test custom thresholds
   - Verify threshold boundary conditions

5. **Error Handling Tests**
   - API failures
   - Empty datasets
   - Missing device properties
   - Invalid parameters

### Mock Data Strategy

```typescript
// Mock device with all properties
const mockDevice = {
  id: 'device-001',
  deviceName: 'iPhone 14 Pro',
  manufacturer: 'Apple Inc.',
  model: 'iPhone 14 Pro',
  serialNumber: 'SN001',
  userPrincipalName: 'user@company.com',
  enrolledDateTime: '2024-01-15T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:00:00Z',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  totalStorageSpaceInBytes: 256 * 1024 * 1024 * 1024,
  freeStorageSpaceInBytes: 128 * 1024 * 1024 * 1024,
  physicalMemoryInBytes: 6 * 1024 * 1024 * 1024,
  managementAgent: 'mdm',
  complianceState: 'compliant',
};
```

### Integration Testing

```typescript
describe('Integration Tests', () => {
  it('should generate complete report end-to-end', async () => {
    // Real Graph API call (requires auth)
    const auth = new GraphAuth(config);
    const client = await auth.getAuthenticatedClient();
    const reporter = new DeviceHardwareReports(client, config);

    const report = await reporter.generateComprehensiveReport();

    expect(report.metadata).toBeDefined();
    expect(report.data.length).toBeGreaterThan(0);
    expect(report.summary.totalDevices).toBe(report.data.length);
  });
});
```

### Performance Testing

```typescript
describe('Performance Tests', () => {
  it('should handle 10,000 devices efficiently', async () => {
    const largeDataset = generateMockDevices(10000);

    const startTime = performance.now();
    const summary = await reporter.getMemoryRangeSummary(largeDataset);
    const duration = performance.now() - startTime;

    expect(duration).toBeLessThan(1000); // Should complete in < 1 second
    expect(summary.length).toBeGreaterThan(0);
  });
});
```

## Performance Optimization

### Optimization Techniques

#### 1. Data Caching

```typescript
class OptimizedDeviceHardwareReports extends DeviceHardwareReports {
  private deviceCache: DeviceHardwareInfo[] | null = null;
  private cacheTimestamp: Date | null = null;
  private cacheDuration: number = 5 * 60 * 1000; // 5 minutes

  async getAllDeviceHardwareInfo(
    filter?: HardwareReportFilter
  ): Promise<DeviceHardwareInfo[]> {
    const now = new Date();

    // Return cached data if fresh
    if (
      this.deviceCache &&
      this.cacheTimestamp &&
      now.getTime() - this.cacheTimestamp.getTime() < this.cacheDuration
    ) {
      return this.applyFiltersToCache(this.deviceCache, filter);
    }

    // Fetch fresh data
    const devices = await super.getAllDeviceHardwareInfo(filter);

    // Update cache
    this.deviceCache = devices;
    this.cacheTimestamp = now;

    return devices;
  }

  clearCache() {
    this.deviceCache = null;
    this.cacheTimestamp = null;
  }
}
```

#### 2. Batch Processing

```typescript
async processLargeDeviceSet(
  batchSize: number = 1000
): Promise<ReportData[]> {
  const allDevices = await this.getAllDeviceHardwareInfo();
  const reports: ReportData[] = [];

  for (let i = 0; i < allDevices.length; i += batchSize) {
    const batch = allDevices.slice(i, i + batchSize);
    const report = await this.generateComprehensiveReport();
    report.data = batch; // Override with batch
    reports.push(report);
  }

  return reports;
}
```

#### 3. Parallel Report Generation

```typescript
async generateAllReportsInParallel(
  devices: DeviceHardwareInfo[]
): Promise<any> {
  const [
    displayConfig,
    osDistribution,
    memoryRanges,
    storageRanges,
    lowMemory,
    lowStorage,
  ] = await Promise.all([
    this.getDisplayConfigurationSummary(devices),
    this.getOSDistributionSummary(devices),
    this.getMemoryRangeSummary(devices),
    this.getStorageRangeSummary(devices),
    this.getDevicesWithLowMemory(512, devices),
    this.getDevicesWithLowStorage(10, devices),
  ]);

  return {
    displayConfig,
    osDistribution,
    memoryRanges,
    storageRanges,
    lowMemory,
    lowStorage,
  };
}
```

## Extension Points

### Adding New Report Types

The module is designed for extension. Here's how to add a new report:

#### Example: Device Battery Health Report

```typescript
interface BatteryHealthInfo {
  deviceId: string;
  deviceName: string;
  batteryHealthPercentage: number;
  batteryLevel: number;
  chargeCycles: number;
  estimatedRuntimeInMinutes: number;
  needsReplacement: boolean;
}

// Extend the class
export class ExtendedDeviceHardwareReports extends DeviceHardwareReports {
  /**
   * New Report: Battery Health Analysis
   */
  async getBatteryHealthReport(): Promise<BatteryHealthInfo[]> {
    // This would require additional Graph API endpoints
    // Currently not available in /managedDevices but shown as example

    const devices = await this.getAllDeviceHardwareInfo();

    // Would fetch battery data from additional API call
    // For demo purposes, showing structure

    return devices.map(device => ({
      deviceId: device.deviceId,
      deviceName: device.deviceName,
      batteryHealthPercentage: 85, // Would come from API
      batteryLevel: 75, // Current charge level
      chargeCycles: 350, // Total charge cycles
      estimatedRuntimeInMinutes: 480, // 8 hours
      needsReplacement: false,
    }));
  }

  /**
   * New Report: Devices needing battery replacement
   */
  async getDevicesNeedingBatteryReplacement(
    healthThreshold: number = 80
  ): Promise<BatteryHealthInfo[]> {
    const batteryReport = await this.getBatteryHealthReport();

    return batteryReport.filter(
      device => device.batteryHealthPercentage < healthThreshold
    );
  }
}
```

### Custom Export Formats

Add support for new export formats:

```typescript
class CustomExportReports extends DeviceHardwareReports {
  /**
   * Export to Excel with formatting
   */
  async exportToExcel(
    reportData: ReportData,
    outputPath: string
  ): Promise<string> {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();

    // Summary sheet
    const summarySheet = workbook.addWorksheet('Summary');
    summarySheet.addRow(['Metric', 'Value']);
    summarySheet.addRow(['Total Devices', reportData.summary.totalDevices]);
    summarySheet.addRow(['Average Memory (GB)', reportData.summary.averageMemoryGB]);
    summarySheet.addRow(['Average Storage (GB)', reportData.summary.averageStorageGB]);

    // Devices sheet
    const devicesSheet = workbook.addWorksheet('Devices');
    devicesSheet.addRow([
      'Device Name', 'Manufacturer', 'Model', 'OS', 'Memory (GB)', 'Storage (GB)'
    ]);

    reportData.data.forEach((device: any) => {
      devicesSheet.addRow([
        device.deviceName,
        device.manufacturer,
        device.model,
        `${device.operatingSystem.operatingSystem} ${device.operatingSystem.osVersion}`,
        device.memory.totalMemoryInGB,
        device.storage.totalStorageInGB,
      ]);
    });

    await workbook.xlsx.writeFile(outputPath);
    return outputPath;
  }

  /**
   * Export to PDF
   */
  async exportToPDF(
    reportData: ReportData,
    outputPath: string
  ): Promise<string> {
    const PDFDocument = require('pdfkit');
    const fs = require('fs');

    const doc = new PDFDocument();
    doc.pipe(fs.createWriteStream(outputPath));

    // Title
    doc.fontSize(20).text('Hardware Inventory Report', { align: 'center' });
    doc.moveDown();

    // Summary
    doc.fontSize(14).text('Summary');
    doc.fontSize(10);
    doc.text(`Total Devices: ${reportData.summary.totalDevices}`);
    doc.text(`Average Memory: ${reportData.summary.averageMemoryGB} GB`);
    doc.text(`Average Storage: ${reportData.summary.averageStorageGB} GB`);

    doc.end();
    return outputPath;
  }
}
```

## Code Examples

### Complete Implementation Example

```typescript
import { DeviceHardwareReports } from './reports/configmgr/device-hardware-reports';
import { GraphAuth } from './auth/graph-auth';

async function completeHardwareReportingExample() {
  // 1. Initialize
  const config = {
    authentication: {
      tenantId: process.env.AZURE_TENANT_ID!,
      clientId: process.env.AZURE_CLIENT_ID!,
      clientSecret: process.env.AZURE_CLIENT_SECRET!,
      authMethod: 'clientSecret' as const,
    },
    // ... rest of config
  };

  const auth = new GraphAuth(config);
  const graphClient = await auth.getAuthenticatedClient();
  const reporter = new DeviceHardwareReports(graphClient, config);

  // 2. Fetch all device data once
  console.log('Fetching device hardware data...');
  const devices = await reporter.getAllDeviceHardwareInfo();
  console.log(`Found ${devices.length} devices`);

  // 3. Generate all reports using cached data
  console.log('Generating reports...');

  const displaySummary = await reporter.getDisplayConfigurationSummary(devices);
  console.log(`Display configurations: ${displaySummary.length}`);

  const osSummary = await reporter.getOSDistributionSummary(devices);
  console.log(`OS versions: ${osSummary.length}`);

  const memoryRanges = await reporter.getMemoryRangeSummary(devices);
  console.log(`Memory ranges analyzed`);

  const storageRanges = await reporter.getStorageRangeSummary(devices);
  console.log(`Storage ranges analyzed`);

  // 4. Identify critical devices
  const lowMemory = await reporter.getDevicesWithLowMemory(512, devices);
  console.log(`Devices with low memory: ${lowMemory.length}`);

  const lowStorage = await reporter.getDevicesWithLowStorage(10, devices);
  console.log(`Devices with low storage: ${lowStorage.length}`);

  // 5. Generate comprehensive report
  const comprehensiveReport = await reporter.generateComprehensiveReport();

  // 6. Export in all formats
  console.log('Exporting reports...');

  await reporter.exportReport(comprehensiveReport, {
    format: 'json',
    outputDir: './reports',
    includeTimestamp: true,
  });

  await reporter.exportReport(comprehensiveReport, {
    format: 'csv',
    outputDir: './reports',
    includeTimestamp: true,
  });

  await reporter.exportReport(comprehensiveReport, {
    format: 'html',
    outputDir: './reports',
    includeTimestamp: true,
  });

  console.log('Reports generated successfully!');

  // 7. Return summary for display
  return {
    totalDevices: devices.length,
    lowMemoryCount: lowMemory.length,
    lowStorageCount: lowStorage.length,
    avgMemory: comprehensiveReport.summary.averageMemoryGB,
    avgStorage: comprehensiveReport.summary.averageStorageGB,
  };
}

// Execute
completeHardwareReportingExample()
  .then(summary => console.log('Summary:', summary))
  .catch(error => console.error('Error:', error));
```

## Summary

The Device Hardware Reports module provides a robust, extensible foundation for hardware inventory management in Intune environments. Key architectural decisions:

1. **Separation of Concerns**: Data fetching, processing, and reporting are clearly separated
2. **Caching Support**: Pass device arrays to avoid redundant API calls
3. **Type Safety**: Comprehensive TypeScript interfaces ensure correctness
4. **Extensibility**: Easy to add custom reports and integrations
5. **Error Handling**: Retry logic and graceful degradation
6. **Performance**: Optimized for large device inventories
7. **Testing**: Comprehensive test coverage with mocking support

For more information:
- [Usage Documentation](./CONFIGMGR_HARDWARE_REPORTS.md)
- [Examples](../src/reports/examples/device-hardware-reports-examples.ts)
- [Unit Tests](../src/reports/configmgr/device-hardware-reports.test.ts)
- [Microsoft Graph API Docs](https://docs.microsoft.com/graph/api/intune-devices-manageddevice-list)

---

**Document Version**: 1.0.0
**Last Updated**: 2024-02-07
**Compatibility**: Node.js 18+, TypeScript 5+
