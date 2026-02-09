# Configuration Manager Device Category and Enrollment Reports - Implementation Guide

## Overview

This document provides comprehensive implementation guidance for the Device Category and Enrollment Reports module, which replicates Configuration Manager functionality in Microsoft Intune environments.

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Report 18: Device Category Management](#report-18-device-category-management)
- [Report 24: Enrollment Failure Detection](#report-24-enrollment-failure-detection)
- [Implementation Details](#implementation-details)
- [API Endpoints](#api-endpoints)
- [Usage Examples](#usage-examples)
- [Troubleshooting](#troubleshooting)
- [Best Practices](#best-practices)

## Architecture

### Module Structure

```
device-category-enrollment-reports.ts
├── Type Definitions
│   ├── DeviceCategory
│   ├── CategorizedDevice
│   ├── UnmanagedDevice
│   ├── EnrollmentFailureAnalysis
│   └── Filter Options
├── Helper Functions
│   ├── determineEnrollmentFailureType()
│   ├── determineCertificateStatus()
│   ├── getTroubleshootingSteps()
│   └── getRecommendation()
└── DeviceCategoryEnrollmentReports Class
    ├── Category Management Methods
    ├── Enrollment Failure Methods
    └── Export Methods
```

### Key Components

1. **Device Category Management**: List and analyze devices by category
2. **Enrollment Failure Detection**: Identify and categorize enrollment issues
3. **Failure Analysis**: Comprehensive analysis with recommendations
4. **Export Capabilities**: JSON, CSV, and HTML formats

## Report 18: Device Category Management

### Overview

Report 18 provides comprehensive device categorization capabilities, allowing administrators to list, filter, and analyze devices by their assigned categories.

### Features

- List all device categories
- Get devices by category name or ID
- Generate category summary with statistics
- Platform and owner type breakdowns
- Compliance tracking per category
- Enrollment age analysis

### API Methods

#### `getAllDeviceCategories()`

Retrieves all device categories defined in the tenant.

**Returns**: `Promise<DeviceCategory[]>`

**Example**:
```typescript
const report = new DeviceCategoryEnrollmentReports(graphClient, config);
const categories = await report.getAllDeviceCategories();

console.log(`Found ${categories.length} categories:`);
categories.forEach(cat => {
  console.log(`- ${cat.displayName}`);
  console.log(`  ID: ${cat.id}`);
  console.log(`  Description: ${cat.description || 'N/A'}`);
});
```

**Output**:
```
Found 5 categories:
- Corporate
  ID: cat-001
  Description: Corporate-owned devices
- Personal
  ID: cat-002
  Description: Personal devices (BYOD)
- Kiosk
  ID: cat-003
  Description: Kiosk and shared devices
- Development
  ID: cat-004
  Description: Developer devices
- Test
  ID: cat-005
  Description: Test and staging devices
```

#### `getDevicesByCategory(categoryName: string)`

Retrieves all devices in a specific category.

**Parameters**:
- `categoryName`: The display name of the category

**Returns**: `Promise<CategorizedDevice[]>`

**Example**:
```typescript
const devices = await report.getDevicesByCategory('Corporate');

console.log(`\nDevices in Corporate category: ${devices.length}`);
devices.forEach(device => {
  console.log(`\n  Device: ${device.deviceName}`);
  console.log(`  User: ${device.userPrincipalName}`);
  console.log(`  Platform: ${device.platform} ${device.osVersion}`);
  console.log(`  Compliance: ${device.complianceState}`);
  console.log(`  Management State: ${device.managementState}`);
  console.log(`  Enrolled: ${new Date(device.enrolledDateTime).toLocaleDateString()}`);
  console.log(`  Last Sync: ${new Date(device.lastSyncDateTime).toLocaleDateString()}`);
});
```

**Output**:
```
Devices in Corporate category: 142

  Device: iPhone 14 Pro - John Doe
  User: john.doe@contoso.com
  Platform: iOS 17.2.1
  Compliance: Compliant
  Management State: Managed
  Enrolled: 2024-01-15
  Last Sync: 2024-02-06

  Device: Samsung Galaxy S23 - Jane Smith
  User: jane.smith@contoso.com
  Platform: Android 14.0
  Compliance: Compliant
  Management State: Managed
  Enrolled: 2024-01-20
  Last Sync: 2024-02-06
```

#### `getDevicesByCategoryId(categoryId: string)`

Retrieves devices by category ID.

**Parameters**:
- `categoryId`: The unique identifier of the category

**Returns**: `Promise<CategorizedDevice[]>`

**Example**:
```typescript
const devices = await report.getDevicesByCategoryId('cat-001');
console.log(`Found ${devices.length} devices in category`);
```

#### `getCategorySummary()`

Generates comprehensive summary statistics for all categories.

**Returns**: `Promise<CategorySummary[]>`

**Example**:
```typescript
const summaries = await report.getCategorySummary();

summaries.forEach(summary => {
  console.log(`\nCategory: ${summary.categoryName}`);
  console.log(`  Total Devices: ${summary.totalDevices}`);
  console.log(`  Compliant: ${summary.compliantDevices} (${Math.round(summary.compliantDevices / summary.totalDevices * 100)}%)`);
  console.log(`  Non-Compliant: ${summary.nonCompliantDevices}`);
  console.log(`  Average Enrollment Age: ${summary.averageEnrollmentAge} days`);

  console.log(`  Platform Breakdown:`);
  Object.entries(summary.platformBreakdown).forEach(([platform, count]) => {
    if (count > 0) {
      console.log(`    ${platform}: ${count}`);
    }
  });

  console.log(`  Owner Type Breakdown:`);
  Object.entries(summary.ownerTypeBreakdown).forEach(([type, count]) => {
    if (count > 0) {
      console.log(`    ${type}: ${count}`);
    }
  });
});
```

**Output**:
```
Category: Corporate
  Total Devices: 142
  Compliant: 138 (97%)
  Non-Compliant: 4
  Average Enrollment Age: 45 days
  Platform Breakdown:
    iOS: 58
    Android: 47
    Windows: 32
    macOS: 5
  Owner Type Breakdown:
    Company: 142

Category: Personal
  Total Devices: 87
  Compliant: 79 (91%)
  Non-Compliant: 8
  Average Enrollment Age: 62 days
  Platform Breakdown:
    iOS: 42
    Android: 45
  Owner Type Breakdown:
    Personal: 87
```

### Export Options

Export category reports in multiple formats:

```typescript
// Export to JSON
const jsonPath = await report.exportCategoryReportToJSON('Corporate', './reports');
console.log(`JSON report: ${jsonPath}`);

// Export to CSV
const csvPath = await report.exportCategoryReportToCSV('Corporate', './reports');
console.log(`CSV report: ${csvPath}`);

// Export to HTML
const htmlPath = await report.exportCategoryReportToHTML('Corporate', './reports');
console.log(`HTML report: ${htmlPath}`);
```

## Report 24: Enrollment Failure Detection

### Overview

Report 24 identifies and categorizes devices with enrollment failures, providing comprehensive troubleshooting guidance and actionable recommendations.

### Failure Types

The module detects the following enrollment failure types:

1. **Certificate Failure**: Expired, invalid, or missing certificates
2. **Network Failure**: Connectivity issues to Intune endpoints
3. **Policy Failure**: Policy application errors
4. **Authentication Failure**: Identity and authentication issues
5. **Assignment Failure**: Failed to receive initial assignments
6. **Sync Failure**: Extended periods without sync
7. **Management State Error**: Management agent errors
8. **Enrollment Blocked**: Blocked by restrictions

### Certificate Status Detection

Certificate status is automatically determined:

- **Valid**: Certificate is valid and not expiring soon
- **Expiring**: Certificate expires within 30 days
- **Expired**: Certificate has expired
- **Not Issued**: No certificate issued
- **Invalid**: Certificate data is invalid
- **Unknown**: Certificate status cannot be determined

### Severity Levels

Each failure is assigned a severity level:

- **Critical**: Certificate or authentication failures (immediate action required)
- **High**: Assignment failures or recent enrollment failures
- **Medium**: Policy, network, or sync failures
- **Low**: Other issues with lower impact

### API Methods

#### `getUnmanagedDevicesEnrollmentFailed(filter?: EnrollmentFailureFilter)`

Identifies unmanaged devices with enrollment failures.

**Parameters**:
- `filter` (optional): Filter criteria for failures

**Returns**: `Promise<UnmanagedDevice[]>`

**Example**:
```typescript
// Get all unmanaged devices
const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

console.log(`Found ${unmanagedDevices.length} unmanaged devices\n`);

unmanagedDevices.forEach((device, index) => {
  console.log(`${index + 1}. Device: ${device.deviceName || 'Unknown'}`);
  console.log(`   User: ${device.userPrincipalName}`);
  console.log(`   Platform: ${device.platform} ${device.osVersion || ''}`);
  console.log(`   Failure Type: ${device.failureType}`);
  console.log(`   Category: ${device.failureCategory}`);
  console.log(`   Severity: ${device.severity}`);
  console.log(`   Certificate Status: ${device.certificateStatus}`);
  console.log(`   Days Since Enrollment: ${device.daysSinceEnrollment || 'Unknown'}`);
  console.log(`   Management State: ${device.managementState}`);
  console.log(`   Failure Reason: ${device.failureReason}`);
  console.log(`   Recommendation: ${device.recommendation}`);
  console.log(`\n   Troubleshooting Steps:`);
  device.troubleshootingSteps.slice(0, 5).forEach((step, i) => {
    console.log(`     ${i + 1}. ${step}`);
  });
  console.log('');
});
```

**Output**:
```
Found 12 unmanaged devices

1. Device: Samsung Galaxy S23
   User: jane.smith@contoso.com
   Platform: Android 14.0
   Failure Type: Certificate Failure
   Category: Infrastructure
   Severity: Critical
   Certificate Status: Expired
   Days Since Enrollment: 47
   Management State: Unhealthy
   Failure Reason: Device certificate has expired or is invalid, preventing secure communication with management server (Management State: unhealthy)
   Recommendation: CRITICAL: Immediate action required: Review and fix certificate infrastructure. Consider implementing automated certificate renewal.

   Troubleshooting Steps:
     1. Review device enrollment status in Microsoft Intune admin center
     2. Check user license assignments in Azure AD
     3. Verify enrollment restrictions are not blocking the device
     4. Verify NDES/SCEP connector is running and accessible
     5. Check certificate template configuration in ADCS

2. Device: iPad Pro
   User: john.doe@contoso.com
   Platform: iOS 17.2.1
   Failure Type: Assignment Failure
   Category: Configuration
   Severity: High
   Certificate Status: Valid
   Days Since Enrollment: 5
   Management State: Discovered
   Failure Reason: Device enrolled but failed to receive initial assignment from management server (Management State: discovered)
   Recommendation: Review group assignments and enrollment profiles. Verify scope tags are correctly configured.

   Troubleshooting Steps:
     1. Review device enrollment status in Microsoft Intune admin center
     2. Check user license assignments in Azure AD
     3. Verify enrollment restrictions are not blocking the device
     4. Verify enrollment profile assignments
     5. Check device group memberships
```

**Filtering Examples**:

```typescript
// Get only critical failures
const criticalFailures = await report.getUnmanagedDevicesEnrollmentFailed({
  severity: ['Critical']
});

// Get certificate-related failures
const certFailures = await report.getUnmanagedDevicesEnrollmentFailed({
  failureType: [EnrollmentFailureType.CertificateFailure]
});

// Get recent failures (within 7 days)
const recentFailures = await report.getUnmanagedDevicesEnrollmentFailed({
  maxDaysSinceEnrollment: 7
});

// Get iOS certificate failures
const iosCertFailures = await report.getUnmanagedDevicesEnrollmentFailed({
  platform: ['ios'],
  failureType: [EnrollmentFailureType.CertificateFailure]
});
```

#### `analyzeEnrollmentFailures(filter?: EnrollmentFailureFilter)`

Generates comprehensive enrollment failure analysis with statistics and recommendations.

**Parameters**:
- `filter` (optional): Filter criteria for analysis

**Returns**: `Promise<EnrollmentFailureAnalysis>`

**Example**:
```typescript
const analysis = await report.analyzeEnrollmentFailures();

console.log('=== ENROLLMENT FAILURE ANALYSIS ===\n');
console.log(`Total Failures: ${analysis.totalFailures}`);
console.log(`Affected Users: ${analysis.affectedUsers}`);
console.log(`Average Days Since Failure: ${analysis.averageDaysSinceFailure}`);

console.log('\n--- Failures by Type ---');
Object.entries(analysis.failuresByType).forEach(([type, count]) => {
  if (count > 0) {
    console.log(`${type}: ${count} (${Math.round(count / analysis.totalFailures * 100)}%)`);
  }
});

console.log('\n--- Failures by Category ---');
Object.entries(analysis.failuresByCategory).forEach(([category, count]) => {
  console.log(`${category}: ${count} (${Math.round(count / analysis.totalFailures * 100)}%)`);
});

console.log('\n--- Failures by Severity ---');
Object.entries(analysis.failuresBySeverity).forEach(([severity, count]) => {
  console.log(`${severity}: ${count} (${Math.round(count / analysis.totalFailures * 100)}%)`);
});

console.log('\n--- Failures by Platform ---');
Object.entries(analysis.failuresByPlatform).forEach(([platform, count]) => {
  console.log(`${platform}: ${count}`);
});

console.log('\n--- Issue Breakdown ---');
console.log(`Certificate Issues: ${analysis.certificateIssues}`);
console.log(`Network Issues: ${analysis.networkIssues}`);
console.log(`Policy Issues: ${analysis.policyIssues}`);
console.log(`Authentication Issues: ${analysis.authenticationIssues}`);

if (analysis.criticalDevices.length > 0) {
  console.log(`\n--- Critical Devices (${analysis.criticalDevices.length}) ---`);
  analysis.criticalDevices.forEach(device => {
    console.log(`- ${device.deviceName}: ${device.failureType}`);
  });
}

console.log('\n--- RECOMMENDED ACTIONS ---');
analysis.recommendedActions.forEach((action, index) => {
  console.log(`${index + 1}. ${action}`);
});
```

**Output**:
```
=== ENROLLMENT FAILURE ANALYSIS ===

Total Failures: 23
Affected Users: 19
Average Days Since Failure: 18

--- Failures by Type ---
Certificate Failure: 8 (35%)
Assignment Failure: 5 (22%)
Sync Failure: 4 (17%)
Policy Failure: 3 (13%)
Authentication Failure: 2 (9%)
Network Failure: 1 (4%)

--- Failures by Category ---
Infrastructure: 10 (43%)
Configuration: 8 (35%)
Connectivity: 5 (22%)

--- Failures by Severity ---
Critical: 10 (43%)
High: 6 (26%)
Medium: 5 (22%)
Low: 2 (9%)

--- Failures by Platform ---
Android: 9
iOS: 8
Windows: 4
macOS: 2

--- Issue Breakdown ---
Certificate Issues: 8
Network Issues: 1
Policy Issues: 3
Authentication Issues: 2

--- Critical Devices (10) ---
- Samsung Galaxy S23: Certificate Failure
- Pixel 7 Pro: Certificate Failure
- iPhone 14: Authentication Failure
- iPad Air: Certificate Failure
- Galaxy Tab S8: Authentication Failure

--- RECOMMENDED ACTIONS ---
1. URGENT: 8 devices have certificate issues. Review NDES/SCEP infrastructure immediately.
2. 2 devices have authentication failures. Review Azure AD and Conditional Access policies.
3. 1 devices experiencing network connectivity issues. Verify firewall rules and proxy settings.
4. 3 devices failing policy application. Review policy assignments and configurations.
5. Primary issue: Certificate Failure (8 devices). Focus remediation efforts here.
6. Review troubleshooting steps for each device and implement recommended fixes.
7. Monitor enrollment health regularly to catch issues early.
```

#### `getDevicesWithFailedAssignment()`

Identifies devices that enrolled but failed to receive initial assignments.

**Returns**: `Promise<DeviceAssignmentFailure[]>`

**Example**:
```typescript
const assignmentFailures = await report.getDevicesWithFailedAssignment();

console.log(`Found ${assignmentFailures.length} devices with assignment failures\n`);

assignmentFailures.forEach(device => {
  console.log(`Device: ${device.deviceName}`);
  console.log(`User: ${device.userPrincipalName}`);
  console.log(`Platform: ${device.platform}`);
  console.log(`Assignment Type: ${device.assignmentType}`);
  console.log(`Failure Reason: ${device.failureReason}`);
  console.log(`Expected Assignment: ${device.expectedAssignment}`);
  console.log(`Current Assignment: ${device.currentAssignment}`);
  console.log(`Retry Count: ${device.retryCount}`);
  console.log(`Last Attempt: ${device.lastAttemptDate}`);
  console.log('\nTroubleshooting Steps:');
  device.troubleshootingSteps.forEach((step, i) => {
    console.log(`  ${i + 1}. ${step}`);
  });
  console.log('');
});
```

**Output**:
```
Found 5 devices with assignment failures

Device: iPad Pro
User: john.doe@contoso.com
Platform: iOS
Assignment Type: Initial Assignment
Failure Reason: Device enrolled but failed to receive initial assignment from management server
Expected Assignment: Management policies and configurations
Current Assignment: None
Retry Count: 5
Last Attempt: 2024-02-01T10:00:00Z

Troubleshooting Steps:
  1. Review device enrollment status in Microsoft Intune admin center
  2. Check user license assignments in Azure AD
  3. Verify enrollment restrictions are not blocking the device
  4. Verify enrollment profile assignments
  5. Check device group memberships
  6. Review assignment filters configuration
  7. Validate scope tags are correctly applied
```

### Export Options

Export enrollment failure reports in multiple formats:

```typescript
// Export analysis to JSON
const jsonPath = await report.exportEnrollmentFailureReportToJSON(
  { severity: ['Critical', 'High'] },
  './reports'
);

// Export device list to CSV
const csvPath = await report.exportEnrollmentFailureReportToCSV(
  { failureType: [EnrollmentFailureType.CertificateFailure] },
  './reports'
);

// Export comprehensive HTML report
const htmlPath = await report.exportEnrollmentFailureReportToHTML(
  undefined,
  './reports'
);

console.log('Reports exported:');
console.log(`- JSON: ${jsonPath}`);
console.log(`- CSV: ${csvPath}`);
console.log(`- HTML: ${htmlPath}`);
```

## Implementation Details

### Enrollment Failure Detection Logic

The module uses sophisticated logic to detect and categorize enrollment failures:

#### 1. Certificate Failure Detection

```typescript
if (device.certificateExpirationDateTime) {
  const certExpiry = new Date(device.certificateExpirationDateTime);
  if (certExpiry < new Date()) {
    return EnrollmentFailureType.CertificateFailure;
  }
}
```

#### 2. Sync Failure Detection

```typescript
if (enrolled && lastSync) {
  const daysSinceSync = Math.floor(
    (new Date().getTime() - lastSync.getTime()) / (1000 * 60 * 60 * 24)
  );
  if (daysSinceSync > 7) {
    return EnrollmentFailureType.SyncFailure;
  }
}
```

#### 3. Assignment Failure Detection

```typescript
if (enrolled && !lastSync) {
  const daysSinceEnrollment = Math.floor(
    (new Date().getTime() - enrolled.getTime()) / (1000 * 60 * 60 * 24)
  );
  if (daysSinceEnrollment > 1) {
    return EnrollmentFailureType.AssignmentFailure;
  }
}
```

#### 4. Management State Analysis

```typescript
const managementState = device.managementState?.toLowerCase() || '';

if (managementState === 'unhealthy' || managementState === 'managedwitherrors') {
  return EnrollmentFailureType.ManagementStateError;
}

if (managementState === 'policyapplyfailed') {
  return EnrollmentFailureType.PolicyFailure;
}
```

### Troubleshooting Recommendations

The module provides context-aware troubleshooting steps based on failure type:

#### Certificate Failure Steps

1. Verify NDES/SCEP connector is running
2. Check certificate template configuration
3. Validate certificate profile assignments
4. Review certificate connector logs
5. Ensure device can reach CA endpoints
6. Check for expired or revoked certificates
7. Verify certificate renewal settings

#### Network Failure Steps

1. Test connectivity to Intune endpoints
2. Verify firewall rules
3. Check proxy configuration
4. Validate DNS resolution
5. Review network logs
6. Ensure stable internet connection

#### Policy Failure Steps

1. Review assigned policies for conflicts
2. Check configuration profile compatibility
3. Validate policy assignments
4. Review policy application logs
5. Check for policy syntax errors
6. Verify device meets requirements

#### Authentication Failure Steps

1. Verify user credentials
2. Check Azure AD device registration
3. Review Conditional Access policies
4. Validate MFA configuration
5. Check for expired passwords
6. Review Azure AD sign-in logs

#### Assignment Failure Steps

1. Verify enrollment profile assignments
2. Check device group memberships
3. Review assignment filters
4. Validate scope tags
5. Check for assignment timing issues
6. Review enrollment status page settings

## API Endpoints

### Microsoft Graph API Endpoints Used

| Endpoint | Purpose | Method |
|----------|---------|--------|
| `/deviceManagement/deviceCategories` | Get all categories | GET |
| `/deviceManagement/deviceCategories/{id}` | Get category by ID | GET |
| `/deviceManagement/managedDevices` | Get all managed devices | GET |
| `/deviceManagement/managedDevices?$filter=` | Get filtered devices | GET |

### Query Parameters

- `$select`: Specify fields to return
- `$filter`: Filter results
- `$top`: Limit number of results
- `$orderby`: Sort results

### Example API Calls

```typescript
// Get all categories
GET /deviceManagement/deviceCategories
  ?$select=id,displayName,description,lastModifiedDateTime
  &$top=999

// Get devices by category
GET /deviceManagement/managedDevices
  ?$filter=deviceCategoryDisplayName eq 'Corporate'
  &$select=id,deviceName,userPrincipalName,operatingSystem,managementState
  &$top=999

// Get devices with enrollment details
GET /deviceManagement/managedDevices
  ?$select=id,deviceName,enrolledDateTime,lastSyncDateTime,managementState,
           enrollmentState,certificateExpirationDateTime,complianceState
  &$top=999
```

## Usage Examples

### Complete Workflow Example

```typescript
import { DeviceCategoryEnrollmentReports } from './device-category-enrollment-reports';
import { Client } from '@microsoft/microsoft-graph-client';

// Initialize
const graphClient = Client.init({
  authProvider: (done) => {
    done(null, accessToken);
  }
});

const config = {
  auth: {
    clientId: 'your-client-id',
    clientSecret: 'your-secret',
    tenantId: 'your-tenant-id'
  },
  output: {
    directory: './reports',
    format: 'json'
  }
};

const report = new DeviceCategoryEnrollmentReports(graphClient, config);

// 1. Generate Category Summary
console.log('=== CATEGORY ANALYSIS ===\n');
const summaries = await report.getCategorySummary();

summaries.forEach(summary => {
  console.log(`${summary.categoryName}: ${summary.totalDevices} devices`);
  console.log(`  Compliant: ${summary.compliantDevices}`);
  console.log(`  Platforms: ${Object.keys(summary.platformBreakdown).join(', ')}`);
});

// 2. Analyze Enrollment Failures
console.log('\n=== ENROLLMENT FAILURES ===\n');
const analysis = await report.analyzeEnrollmentFailures();

console.log(`Total Failures: ${analysis.totalFailures}`);
console.log(`Critical Devices: ${analysis.criticalDevices.length}`);
console.log(`Certificate Issues: ${analysis.certificateIssues}`);

// 3. Get Critical Devices
if (analysis.criticalDevices.length > 0) {
  console.log('\nCritical Devices Requiring Immediate Action:');

  for (const device of analysis.criticalDevices) {
    console.log(`\n${device.deviceName} (${device.userPrincipalName})`);
    console.log(`  Issue: ${device.failureType}`);
    console.log(`  Recommendation: ${device.recommendation}`);

    // Export individual device report
    const deviceReport = {
      metadata: {
        reportName: `Critical Device: ${device.deviceName}`,
        generatedAt: new Date().toISOString()
      },
      data: [device]
    };
  }
}

// 4. Export Comprehensive Reports
console.log('\n=== EXPORTING REPORTS ===\n');

// Export by category
for (const summary of summaries) {
  if (summary.totalDevices > 0) {
    const path = await report.exportCategoryReportToHTML(
      summary.categoryName,
      './reports/categories'
    );
    console.log(`Exported: ${path}`);
  }
}

// Export enrollment failures
const failurePath = await report.exportEnrollmentFailureReportToHTML(
  { severity: ['Critical', 'High'] },
  './reports/failures'
);
console.log(`Exported failures: ${failurePath}`);

// 5. Generate Action Plan
console.log('\n=== ACTION PLAN ===\n');
analysis.recommendedActions.forEach((action, index) => {
  console.log(`${index + 1}. ${action}`);
});
```

### Automated Monitoring Script

```typescript
// Daily enrollment health monitoring
async function monitorEnrollmentHealth() {
  const report = new DeviceCategoryEnrollmentReports(graphClient, config);

  // Get current state
  const analysis = await report.analyzeEnrollmentFailures();

  // Check thresholds
  const criticalThreshold = 5;
  const failureRateThreshold = 0.05; // 5%

  if (analysis.criticalDevices.length >= criticalThreshold) {
    console.log('⚠️ ALERT: High number of critical failures detected!');
    console.log(`Critical devices: ${analysis.criticalDevices.length}`);

    // Send alert notification
    await sendAlertEmail({
      subject: 'Critical Enrollment Failures Detected',
      body: `${analysis.criticalDevices.length} devices require immediate attention`,
      devices: analysis.criticalDevices
    });
  }

  // Export daily report
  const reportPath = await report.exportEnrollmentFailureReportToHTML(
    undefined,
    `./reports/daily/${new Date().toISOString().split('T')[0]}`
  );

  console.log(`Daily report saved: ${reportPath}`);
}

// Run daily
setInterval(monitorEnrollmentHealth, 24 * 60 * 60 * 1000);
```

## Troubleshooting

### Common Issues

#### Issue: No devices returned for category

**Symptoms**: `getDevicesByCategory()` returns empty array

**Solutions**:
1. Verify category name is spelled correctly (case-sensitive)
2. Check if devices are actually assigned to the category
3. Confirm Graph API permissions are sufficient
4. Review filter syntax in API call

#### Issue: Certificate status shows as "Unknown"

**Symptoms**: All devices show `CertificateStatus.Unknown`

**Solutions**:
1. Verify SCEP/NDES is properly configured
2. Check if certificate profiles are assigned
3. Confirm devices have enrolled with certificate
4. Review certificate connector logs

#### Issue: High number of "Unknown" failure types

**Symptoms**: Many devices categorized as `EnrollmentFailureType.Unknown`

**Solutions**:
1. Check device logs for detailed error codes
2. Review managementState values
3. Verify lastSyncDateTime is being captured
4. Update failure detection logic if needed

### Performance Optimization

For large tenants with thousands of devices:

```typescript
// Use pagination for large datasets
async function getLargeDeviceList() {
  const pageSize = 100;
  let allDevices = [];
  let hasMore = true;
  let skipToken = '';

  while (hasMore) {
    const response = await graphClient
      .api('/deviceManagement/managedDevices')
      .top(pageSize)
      .skipToken(skipToken)
      .get();

    allDevices = allDevices.concat(response.value);

    if (response['@odata.nextLink']) {
      skipToken = response['@odata.nextLink'].split('$skiptoken=')[1];
    } else {
      hasMore = false;
    }
  }

  return allDevices;
}
```

## Best Practices

### 1. Regular Monitoring

- Run enrollment failure analysis daily
- Set up automated alerts for critical failures
- Track trends over time
- Monitor certificate expiration dates

### 2. Proactive Management

- Review devices approaching certificate expiration
- Address assignment failures within 24 hours
- Implement automated remediation where possible
- Maintain up-to-date enrollment documentation

### 3. Reporting

- Export reports before major changes
- Keep historical reports for trend analysis
- Share reports with relevant stakeholders
- Document remediation actions taken

### 4. Category Management

- Use consistent naming conventions
- Document category purposes
- Regularly review category assignments
- Keep categories aligned with business needs

### 5. Security

- Protect exported reports (may contain sensitive data)
- Use appropriate Graph API permissions (minimum required)
- Implement audit logging for report access
- Secure API credentials properly

## Conclusion

The Device Category and Enrollment Reports module provides comprehensive capabilities for managing device categories and identifying enrollment failures in Intune environments. By following this implementation guide, administrators can effectively monitor, troubleshoot, and resolve device management issues.

For additional support or questions, please refer to:
- Microsoft Graph API documentation
- Intune troubleshooting documentation
- Module source code and inline documentation
- Unit tests for usage examples
