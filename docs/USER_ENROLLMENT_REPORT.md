# User Enrollment Report Module

Comprehensive user enrollment reporting for Microsoft Intune using the Graph API.

## Overview

The User Enrollment Report module provides detailed insights into device enrollment across your Intune tenant, including user-specific enrollment data, failure analysis, trend reporting, and device ownership tracking.

## Features

### 1. Get All Enrolled Users
- Fetch all users with Intune licenses
- Retrieve user profile information (department, job title, etc.)
- Filter by license assignment
- Support for large user bases with pagination

### 2. Get Devices Per User
- Retrieve managed devices for specific users
- Get owned devices (registered + managed)
- Filter by user principal name (UPN)
- Support for multiple device types and platforms

### 3. Get Enrollment Status and Methods
- Track enrollment methods (User Enrollment, Device Enrollment Program, etc.)
- Monitor management states (Managed, Supervised, etc.)
- Analyze ownership types (Corporate, Personal, BYOD)
- Platform-specific enrollment breakdown

### 4. Get Enrollment Failures and Troubleshooting
- Automatic failure categorization
- Detailed error code mapping
- Step-by-step troubleshooting guides
- Correlation IDs for tracking

### 5. Generate Enrollment Trend Reports
- Daily enrollment trends over configurable periods
- Success vs failure rate tracking
- Platform breakdown over time
- Unique user count tracking

### 6. Track User Device Ownership
- Corporate vs Personal device tracking
- Platform-specific ownership analysis
- Per-user ownership summary
- Owned devices inventory

## Graph API Endpoints Used

| Endpoint | Purpose |
|----------|---------|
| `GET /users` | Fetch all enrolled users with licenses |
| `GET /users/{id}/ownedDevices` | Get devices owned by a specific user |
| `GET /users/{id}/managedDevices` | Get managed devices for a specific user |
| `GET /deviceManagement/managedDevices` | Get all managed devices in the tenant |
| `GET /deviceManagement/managedDevices?$filter=userPrincipalName eq '{upn}'` | Get devices by user principal name |
| `GET /deviceManagement/troubleshootingEvents` | Get enrollment failures and troubleshooting events |

## Installation

```bash
npm install @microsoft/microsoft-graph-client @azure/identity
```

## Quick Start

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { UserEnrollmentReport } from './reports/user-enrollment-report';
import { AppConfig } from './types';

// Create Graph client (see authentication section)
const graphClient = createGraphClient();

// Create config
const config: AppConfig = {
  tenantId: 'your-tenant-id',
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  outputDir: './reports'
};

// Create report instance
const report = new UserEnrollmentReport(graphClient, config);

// Generate comprehensive report
const result = await report.execute({
  includeTrends: true,
  includeFailures: true,
  includePlatformStats: true,
  trendDays: 30
});

// Export in different formats
const jsonReport = report.exportAsJSON(result);
const csvReport = report.exportAsCSV(result);
const htmlReport = report.exportAsHTML(result);
```

## Authentication

### Using Client Credentials (Service Principal)

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';

function createGraphClient(): Client {
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

### Required Permissions

The following Microsoft Graph API permissions are required:

- `DeviceManagementManagedDevices.Read.All` - Read managed device information
- `DeviceManagementServiceConfig.Read.All` - Read device management configuration
- `User.Read.All` - Read all users
- `Directory.Read.All` - Read directory data

## API Reference

### Main Report Methods

#### `execute(options?: UserEnrollmentReportOptions): Promise<ReportData>`

Generate a comprehensive enrollment report.

**Options:**
```typescript
interface UserEnrollmentReportOptions {
  startDate?: Date;              // Start date for filtering (optional)
  endDate?: Date;                // End date for filtering (optional)
  includeTrends?: boolean;       // Include trend analysis (default: true)
  includeFailures?: boolean;     // Include failure analysis (default: true)
  includePlatformStats?: boolean; // Include platform stats (default: true)
  includeUserDetails?: boolean;  // Include user details (default: false)
  trendDays?: number;            // Days to analyze trends (default: 30)
  maxFailures?: number;          // Max failures to include (default: 100)
  platformFilter?: string[];     // Filter by platforms
  userFilter?: string[];         // Filter by user IDs
}
```

#### `getAllEnrolledUsers(): Promise<any[]>`

Fetch all users with Intune licenses assigned.

**Returns:** Array of user objects with profile information.

```typescript
const users = await report.getAllEnrolledUsers();
console.log(`Found ${users.length} enrolled users`);
```

#### `getUserDevices(userId: string): Promise<EnrolledDevice[]>`

Get all managed devices for a specific user.

**Parameters:**
- `userId`: Azure AD user ID

**Returns:** Array of enrolled device objects.

```typescript
const devices = await report.getUserDevices('user-id-here');
devices.forEach(device => {
  console.log(`${device.deviceName} - ${device.platform}`);
});
```

#### `getUserDevicesByUPN(userPrincipalName: string): Promise<EnrolledDevice[]>`

Get managed devices by user principal name using filter.

**Parameters:**
- `userPrincipalName`: User's UPN (email address)

**Returns:** Array of enrolled device objects.

```typescript
const devices = await report.getUserDevicesByUPN('user@contoso.com');
```

#### `getUserOwnedDevices(userId: string): Promise<any[]>`

Get both registered and managed devices owned by a user.

**Parameters:**
- `userId`: Azure AD user ID

**Returns:** Array of owned device objects.

```typescript
const ownedDevices = await report.getUserOwnedDevices('user-id-here');
```

### Enrollment Status Methods

#### `getUserEnrollmentStatus(userId: string): Promise<UserEnrollmentStatus>`

Get enrollment status for a specific user.

**Returns:** User enrollment status object.

```typescript
const status = await report.getUserEnrollmentStatus('user-id-here');
console.log(`User: ${status.displayName}`);
console.log(`Devices: ${status.enrolledDeviceCount}`);
console.log(`Failures: ${status.failureCount}`);
```

#### `getComprehensiveUserEnrollmentStatus(userId: string): Promise<UserEnrollmentStatus>`

Get comprehensive enrollment status including owned devices.

```typescript
const status = await report.getComprehensiveUserEnrollmentStatus('user-id-here');
console.log(`Managed Devices: ${status.devices.length}`);
console.log(`Owned Devices: ${status.ownedDevices?.length || 0}`);
```

#### `getEnrollmentStatusSummary(): Promise<EnrollmentStatusSummary>`

Get enrollment status summary across all managed devices.

```typescript
const summary = await report.getEnrollmentStatusSummary();
console.log('By Management State:', summary.byManagementState);
console.log('By Owner Type:', summary.byOwnerType);
```

#### `getEnrollmentMethodsBreakdown(): Promise<EnrollmentMethodsBreakdown>`

Get breakdown of enrollment methods used.

```typescript
const methods = await report.getEnrollmentMethodsBreakdown();
Object.entries(methods).forEach(([method, data]) => {
  console.log(`${method}: ${data.count} devices (${data.percentage}%)`);
});
```

### Statistics Methods

#### `getEnrollmentStatistics(startDate: Date, endDate: Date): Promise<EnrollmentStatistics>`

Get enrollment statistics for a specific date range.

```typescript
const endDate = new Date();
const startDate = new Date();
startDate.setDate(startDate.getDate() - 30);

const stats = await report.getEnrollmentStatistics(startDate, endDate);
console.log(`Total Devices: ${stats.totalDevices}`);
console.log(`Success Rate: ${stats.enrollmentSuccessRate}%`);
```

#### `getDeviceEnrollmentByPlatform(): Promise<PlatformEnrollmentStats[]>`

Get device enrollment statistics by platform.

```typescript
const platformStats = await report.getDeviceEnrollmentByPlatform();
platformStats.forEach(stat => {
  console.log(`${stat.platform}: ${stat.totalDevices} devices`);
  console.log(`  Top Model: ${stat.topModels[0]?.model}`);
});
```

### Failure Analysis Methods

#### `getEnrollmentFailures(days?: number, maxResults?: number): Promise<EnrollmentFailure[]>`

Get enrollment failures with categorization and troubleshooting.

**Parameters:**
- `days`: Number of days to look back (default: 30)
- `maxResults`: Maximum failures to return (default: 100)

```typescript
const failures = await report.getEnrollmentFailures(30, 50);
failures.forEach(failure => {
  console.log(`User: ${failure.userPrincipalName}`);
  console.log(`Category: ${failure.failureCategory}`);
  console.log(`Reason: ${failure.failureReason}`);
  console.log(`Troubleshooting:`);
  failure.troubleshootingSteps.forEach(step => {
    console.log(`  - ${step}`);
  });
});
```

#### `getEnrollmentTroubleshootingEvents(startDate?, endDate?, userId?): Promise<any[]>`

Get enrollment troubleshooting events with enhanced details.

```typescript
const events = await report.getEnrollmentTroubleshootingEvents(
  new Date('2024-01-01'),
  new Date(),
  'user-id-here'
);
```

#### `getUserTroubleshootingEvents(userId: string, days?: number): Promise<any[]>`

Get troubleshooting events for a specific user.

```typescript
const events = await report.getUserTroubleshootingEvents('user-id-here', 90);
```

### Trend Analysis Methods

#### `getEnrollmentTrends(days?: number): Promise<EnrollmentTrendDataPoint[]>`

Generate enrollment trends over time.

**Parameters:**
- `days`: Number of days to analyze (default: 30)

```typescript
const trends = await report.getEnrollmentTrends(30);
trends.forEach(trend => {
  console.log(`${trend.date}:`);
  console.log(`  Total: ${trend.totalEnrollments}`);
  console.log(`  Successful: ${trend.successfulEnrollments}`);
  console.log(`  Failed: ${trend.failedEnrollments}`);
});
```

### Device Ownership Methods

#### `getDeviceOwnershipBreakdown(): Promise<OwnershipBreakdown>`

Get device ownership breakdown (Corporate vs Personal).

```typescript
const ownership = await report.getDeviceOwnershipBreakdown();
console.log(`Corporate: ${ownership.corporate.count} (${ownership.corporate.percentage}%)`);
console.log(`Personal: ${ownership.personal.count} (${ownership.personal.percentage}%)`);
console.log('By Platform:', ownership.byPlatform);
```

#### `getUserDeviceOwnership(userId: string): Promise<UserDeviceOwnership>`

Get device ownership summary for a specific user.

```typescript
const ownership = await report.getUserDeviceOwnership('user-id-here');
console.log(`Corporate Devices: ${ownership.corporateDevices.length}`);
console.log(`Personal Devices: ${ownership.personalDevices.length}`);
```

### Export Methods

#### `exportAsJSON(reportData: ReportData): string`

Export report data as JSON.

```typescript
const jsonReport = report.exportAsJSON(result);
fs.writeFileSync('enrollment-report.json', jsonReport);
```

#### `exportAsCSV(reportData: ReportData): string`

Export report data as CSV.

```typescript
const csvReport = report.exportAsCSV(result);
fs.writeFileSync('enrollment-report.csv', csvReport);
```

#### `exportAsHTML(reportData: ReportData): string`

Export report data as HTML.

```typescript
const htmlReport = report.exportAsHTML(result);
fs.writeFileSync('enrollment-report.html', htmlReport);
```

## TypeScript Types

### EnrollmentStatistics

```typescript
interface EnrollmentStatistics {
  totalDevices: number;
  totalUsers: number;
  enrolledLast24Hours: number;
  enrolledLast7Days: number;
  enrolledLast30Days: number;
  enrolledLast90Days: number;
  averageDevicesPerUser: number;
  enrollmentSuccessRate: number;
  totalEnrollmentAttempts: number;
  failedEnrollments: number;
  pendingEnrollments: number;
  lastUpdated: string;
}
```

### UserEnrollmentStatus

```typescript
interface UserEnrollmentStatus {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  enrolledDeviceCount: number;
  devices: EnrolledDevice[];
  firstEnrollmentDate?: string;
  lastEnrollmentDate?: string;
  hasEnrollmentFailures: boolean;
  failureCount: number;
  enrollmentRestrictions?: string[];
}
```

### EnrolledDevice

```typescript
interface EnrolledDevice {
  deviceId: string;
  deviceName: string;
  platform: string;
  osVersion: string;
  model?: string;
  manufacturer?: string;
  enrollmentDate: string;
  enrollmentType: string;
  managementState: string;
  complianceState: string;
  lastSyncDateTime: string;
  isSupervised?: boolean;
  serialNumber?: string;
}
```

### EnrollmentFailure

```typescript
interface EnrollmentFailure {
  failureId: string;
  userId: string;
  userPrincipalName: string;
  userDisplayName?: string;
  deviceName?: string;
  deviceId?: string;
  platform: string;
  osVersion?: string;
  failureDateTime: string;
  errorCode: string;
  errorMessage: string;
  failureCategory: EnrollmentFailureCategory;
  failureReason: string;
  troubleshootingSteps: string[];
  correlationId?: string;
  additionalDetails?: Record<string, any>;
}
```

### EnrollmentFailureCategory

```typescript
enum EnrollmentFailureCategory {
  Authentication = 'Authentication',
  Authorization = 'Authorization',
  Licensing = 'Licensing',
  DeviceLimit = 'Device Limit',
  PlatformRestriction = 'Platform Restriction',
  NetworkConnectivity = 'Network Connectivity',
  CertificateIssue = 'Certificate Issue',
  PolicyConflict = 'Policy Conflict',
  DeviceConfiguration = 'Device Configuration',
  Unknown = 'Unknown'
}
```

## Error Code Mapping

The module includes comprehensive error code mapping with troubleshooting steps:

| Error Code | Category | Description |
|------------|----------|-------------|
| 80180002 | Licensing | User does not have an Intune license |
| 80180014 | Device Limit | User has reached maximum device limit |
| 80180018 | Platform Restriction | Device platform not allowed |
| 0x80180026 | Authentication | Authentication failed |
| 0xcaa9001f | Certificate Issue | Certificate validation failed |
| 0x80072ee7 | Network Connectivity | Network connectivity issues |
| 0x80180012 | Authorization | User not authorized to enroll |
| 0x80180013 | Policy Conflict | Conflicting policies |
| 0x87d13ba2 | Device Configuration | Configuration/profile installation failed |

Each error includes detailed troubleshooting steps.

## Error Handling

The module includes comprehensive error handling:

```typescript
import { Logger } from '../core/logger';

try {
  const report = new UserEnrollmentReport(graphClient, config);
  const result = await report.execute();
} catch (error) {
  if (error.statusCode === 429) {
    // Rate limiting - automatically retried with exponential backoff
  } else if (error.statusCode >= 500) {
    // Server error - automatically retried
  } else {
    // Other error - log and handle appropriately
    console.error('Error generating report:', error);
  }
}
```

## Pagination Support

All methods that fetch data from Graph API include automatic pagination:

```typescript
// Automatically handles pagination for large result sets
const devices = await report.fetchAllManagedDevices();
// Returns all devices, not just the first page
```

## Performance Considerations

### Parallel Data Fetching

The module fetches data in parallel where possible:

```typescript
const [devices, failures] = await Promise.all([
  this.fetchAllManagedDevices(),
  this.fetchEnrollmentFailures(30)
]);
```

### Retry Logic

Automatic retry with exponential backoff for rate limiting and server errors:

```typescript
protected async retryGraphCall<T>(
  apiCall: () => Promise<T>,
  maxRetries: number = 3,
  baseDelay: number = 1000
): Promise<T>
```

## Examples

See `/src/reports/examples/user-enrollment-report-example.ts` for comprehensive usage examples:

1. Basic Enrollment Report
2. Enrollment Statistics by Period
3. Get All Enrolled Users
4. Get User Devices
5. User Enrollment Status
6. Enrollment Failures
7. Enrollment Trends
8. Platform Statistics
9. Enrollment Methods
10. Device Ownership
11. User Troubleshooting
12. Get Devices by UPN

## Running Examples

```bash
# Install dependencies
npm install

# Set environment variables
export AZURE_TENANT_ID="your-tenant-id"
export AZURE_CLIENT_ID="your-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
export TEST_USER_ID="user-id-for-testing"
export TEST_USER_UPN="user@contoso.com"

# Run specific example
npx ts-node src/reports/examples/user-enrollment-report-example.ts 1

# Run all examples
npx ts-node src/reports/examples/user-enrollment-report-example.ts
```

## Common Use Cases

### 1. Monitor New Enrollments

```typescript
const stats = await report.getEnrollmentStatistics(
  new Date(Date.now() - 24 * 60 * 60 * 1000),
  new Date()
);
console.log(`New enrollments in last 24h: ${stats.enrolledLast24Hours}`);
```

### 2. Identify Enrollment Issues

```typescript
const failures = await report.getEnrollmentFailures(7);
const byCategory = failures.reduce((acc, f) => {
  acc[f.failureCategory] = (acc[f.failureCategory] || 0) + 1;
  return acc;
}, {});
console.log('Failures by category:', byCategory);
```

### 3. Track BYOD vs Corporate Devices

```typescript
const ownership = await report.getDeviceOwnershipBreakdown();
console.log(`Corporate: ${ownership.corporate.percentage}%`);
console.log(`Personal: ${ownership.personal.percentage}%`);
```

### 4. Analyze Platform Trends

```typescript
const trends = await report.getEnrollmentTrends(30);
const latestTrend = trends[trends.length - 1];
console.log('Latest platform breakdown:', latestTrend.platformBreakdown);
```

### 5. Audit User Device Count

```typescript
const users = await report.getAllEnrolledUsers();
for (const user of users) {
  const status = await report.getUserEnrollmentStatus(user.id);
  if (status.enrolledDeviceCount > 5) {
    console.log(`${user.displayName} has ${status.enrolledDeviceCount} devices`);
  }
}
```

## Troubleshooting

### Issue: No data returned

**Solution:** Verify API permissions and ensure devices are enrolled in Intune.

### Issue: Rate limiting (429 errors)

**Solution:** The module automatically retries with exponential backoff. If issues persist, reduce concurrent requests.

### Issue: Authentication failures

**Solution:** Verify credentials and ensure the service principal has the required Graph API permissions.

### Issue: Missing user data

**Solution:** Ensure users have Intune licenses assigned and devices are properly enrolled.

## Best Practices

1. **Use pagination wisely** - The module handles it automatically, but be aware of large datasets
2. **Cache results** - Don't fetch the same data multiple times unnecessarily
3. **Filter early** - Use Graph API filters to reduce data transfer
4. **Handle errors** - Always wrap calls in try-catch blocks
5. **Monitor rate limits** - Spread out requests for large tenants
6. **Use parallel requests** - Fetch independent data in parallel for better performance

## Contributing

Contributions are welcome! Please ensure:
- TypeScript types are properly defined
- Error handling is comprehensive
- Documentation is updated
- Examples are provided for new features

## License

MIT

## Support

For issues and questions:
- GitHub Issues: [Link to repository]
- Documentation: This file
- Examples: `/src/reports/examples/user-enrollment-report-example.ts`
