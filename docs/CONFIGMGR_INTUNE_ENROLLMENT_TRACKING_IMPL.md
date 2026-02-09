# Configuration Manager Intune Enrollment Tracking Reports - Implementation Guide

## Overview

This document provides comprehensive implementation details for the **Intune Enrollment Tracking Reports** module, which replicates Configuration Manager enrollment tracking reports using Microsoft Graph API.

## Table of Contents

- [Reports Implemented](#reports-implemented)
- [Architecture](#architecture)
- [TypeScript Interfaces](#typescript-interfaces)
- [Core Methods](#core-methods)
- [Features](#features)
- [Usage Examples](#usage-examples)
- [API Reference](#api-reference)
- [Testing](#testing)
- [Best Practices](#best-practices)

## Reports Implemented

### Report 17: List of Devices Enrolled Per User in Microsoft Intune

**Purpose**: Provides a detailed list of all devices enrolled by each user with comprehensive device information.

**Key Information**:
- User details (name, UPN, department, job title)
- Complete device inventory per user
- Platform breakdown (iOS, Android, Windows, macOS, Linux)
- Enrollment method analysis
- Owner type distribution (Corporate vs. Personal)
- Enrollment timeline and device age metrics

### Report 30: Number of Devices Enrolled Per User in Microsoft Intune

**Purpose**: Provides aggregated enrollment counts per user with statistical breakdowns.

**Key Information**:
- Total device count per user
- Platform-specific counts
- Corporate vs. Personal device counts
- Compliance status counts
- Active vs. Inactive device counts
- Enrollment trends over time

## Architecture

### Microsoft Graph API Endpoints

The implementation uses the following Graph API endpoints:

```typescript
// Get all managed devices
GET /deviceManagement/managedDevices

// Get user details
GET /users/{id}
GET /users?$filter=userPrincipalName eq '{upn}'

// Get user's managed devices
GET /users/{id}/managedDevices

// Get user licensing information
GET /users/{id}?$select=assignedLicenses
```

### Data Flow

```
┌─────────────────────┐
│   Graph API Call    │
│  (Managed Devices)  │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Group by User ID   │
│  or User UPN        │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Fetch User Details │
│  (Parallel calls)   │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Calculate Stats    │
│  & Breakdowns       │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Apply Filters &    │
│  Sorting            │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│  Return Report Data │
└─────────────────────┘
```

## TypeScript Interfaces

### Core Data Structures

#### UserEnrolledDevice

```typescript
interface UserEnrolledDevice {
  deviceId: string;
  deviceName: string;
  platform: string;                    // iOS, Android, Windows, macOS, Linux
  osVersion: string;
  model?: string;
  manufacturer?: string;
  enrollmentDate: string;
  enrollmentType: string;              // UserEnrollment, DeviceEnrollment, etc.
  enrollmentMethod: string;
  managementState: string;
  complianceState: string;
  lastSyncDateTime: string;
  ownerType: string;                   // Corporate, Personal
  serialNumber?: string;
  imei?: string;
  isSupervised?: boolean;
  deviceCategory?: string;
  enrollmentProfileName?: string;
}
```

#### UserDeviceList

```typescript
interface UserDeviceList {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  email?: string;
  totalDevices: number;
  devices: UserEnrolledDevice[];
  devicesByPlatform: PlatformBreakdown;
  devicesByEnrollmentMethod: EnrollmentMethodBreakdown;
  devicesByOwnerType: OwnerTypeBreakdown;
  firstEnrollmentDate?: string;
  lastEnrollmentDate?: string;
  averageDeviceAge: number;            // in days
}
```

#### UserEnrollmentCount

```typescript
interface UserEnrollmentCount {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  email?: string;
  totalDeviceCount: number;
  iosDeviceCount: number;
  androidDeviceCount: number;
  windowsDeviceCount: number;
  macOsDeviceCount: number;
  linuxDeviceCount: number;
  otherDeviceCount: number;
  corporateDeviceCount: number;
  personalDeviceCount: number;
  compliantDeviceCount: number;
  nonCompliantDeviceCount: number;
  activeDeviceCount: number;           // Synced in last 30 days
  inactiveDeviceCount: number;
  enrollmentTrend: EnrollmentTrend;
  lastEnrollmentDate?: string;
  firstEnrollmentDate?: string;
}
```

#### UserLicensingInfo

```typescript
interface UserLicensingInfo {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  isLicensed: boolean;
  assignedLicenses: string[];          // SKU IDs
  intuneLicense: boolean;
  ems: boolean;                        // Enterprise Mobility + Security
  microsoft365: boolean;
  accountEnabled: boolean;
}
```

#### DeviceLimitViolation

```typescript
interface DeviceLimitViolation {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  deviceCount: number;
  deviceLimit: number;
  excessDevices: number;
  violationSeverity: 'Low' | 'Medium' | 'High' | 'Critical';
  isLicensed: boolean;
  devices: UserEnrolledDevice[];
}
```

#### UnusualEnrollmentPattern

```typescript
interface UnusualEnrollmentPattern {
  patternType: 'RapidEnrollment' | 'BulkEnrollment' | 'UnusualLocation' |
               'AfterHoursEnrollment' | 'SuspiciousDevice';
  severity: 'Low' | 'Medium' | 'High' | 'Critical';
  userId?: string;
  userPrincipalName?: string;
  displayName?: string;
  deviceCount: number;
  timeframe: string;
  description: string;
  affectedDevices: string[];
  detectedAt: string;
  recommendation: string;
}
```

### Filter Interfaces

#### EnrollmentTrackingFilters

```typescript
interface EnrollmentTrackingFilters {
  userPrincipalName?: string;          // Filter by specific user UPN
  userId?: string;                     // Filter by user ID
  department?: string;                 // Filter by department
  startDate?: Date;                    // Enrollment start date
  endDate?: Date;                      // Enrollment end date
  platform?: string[];                 // Filter by platforms
  enrollmentMethod?: string[];         // Filter by enrollment methods
  ownerType?: 'Corporate' | 'Personal' | 'All';
  complianceState?: 'Compliant' | 'NonCompliant' | 'All';
  minDeviceCount?: number;             // Minimum devices per user
  maxDeviceCount?: number;             // Maximum devices per user
  includeInactiveUsers?: boolean;
  sortBy?: 'deviceCount' | 'userName' | 'lastEnrollment' | 'department';
  sortOrder?: 'asc' | 'desc';
  top?: number;                        // Limit results
}
```

## Core Methods

### IntuneEnrollmentTrackingReports Class

#### Primary Report Methods

##### getDevicesPerUserReport()

```typescript
async getDevicesPerUserReport(
  options?: EnrollmentTrackingReportOptions
): Promise<ReportData>
```

**Description**: Generates Report 17 - detailed list of devices enrolled per user.

**Returns**: Complete report with user device lists, including device details and breakdowns.

##### getEnrollmentCountPerUserReport()

```typescript
async getEnrollmentCountPerUserReport(
  options?: EnrollmentTrackingReportOptions
): Promise<ReportData>
```

**Description**: Generates Report 30 - aggregated enrollment counts per user.

**Returns**: Report with enrollment counts, platform breakdowns, and statistics.

#### Additional User Methods

##### getDevicesEnrolledPerUser()

```typescript
async getDevicesEnrolledPerUser(
  userIdOrUpn: string
): Promise<UserDeviceList | null>
```

**Description**: Gets all devices enrolled by a specific user.

**Parameters**:
- `userIdOrUpn`: User ID or User Principal Name (supports both)

**Returns**: Device list for the user, or null if user not found.

**Example**:
```typescript
// By UPN
const devices = await report.getDevicesEnrolledPerUser('john.doe@contoso.com');

// By User ID
const devices = await report.getDevicesEnrolledPerUser('a1b2c3d4-e5f6-7890');
```

##### getAllUsersWithDeviceCounts()

```typescript
async getAllUsersWithDeviceCounts(
  options?: EnrollmentTrackingReportOptions
): Promise<UserEnrollmentCount[]>
```

**Description**: Retrieves enrollment counts for all users with enrolled devices.

**Returns**: Array of user enrollment counts.

**Example**:
```typescript
const allUsers = await report.getAllUsersWithDeviceCounts();

// Filter by department
const itUsers = await report.getAllUsersWithDeviceCounts({
  filters: { department: 'IT' }
});
```

##### getUsersWithMultipleDevices()

```typescript
async getUsersWithMultipleDevices(
  threshold?: number,
  options?: EnrollmentTrackingReportOptions
): Promise<UserEnrollmentCount[]>
```

**Description**: Identifies users who have enrolled more than the specified number of devices.

**Parameters**:
- `threshold`: Minimum device count (default: 1)
- `options`: Additional filtering options

**Returns**: Users exceeding the threshold.

**Example**:
```typescript
// Find users with 3+ devices
const multiDevice = await report.getUsersWithMultipleDevices(3);

// Find users with 5+ devices in IT
const itMulti = await report.getUsersWithMultipleDevices(5, {
  filters: { department: 'IT' }
});
```

##### getUserEnrollmentDetails()

```typescript
async getUserEnrollmentDetails(
  userIdOrUpn: string
): Promise<{
  user: UserDeviceList | null;
  enrollmentCount: UserEnrollmentCount | null;
  licensing: UserLicensingInfo | null;
  violatesLimit: boolean;
  limitViolation?: DeviceLimitViolation;
}>
```

**Description**: Retrieves comprehensive enrollment details for a user, including licensing validation and limit checking.

**Returns**: Complete user enrollment profile with violation detection.

**Example**:
```typescript
const details = await report.getUserEnrollmentDetails('john.doe@contoso.com');

console.log(`Total devices: ${details.enrollmentCount?.totalDeviceCount}`);
console.log(`Licensed: ${details.licensing?.isLicensed}`);
console.log(`Violates limit: ${details.violatesLimit}`);

if (details.limitViolation) {
  console.log(`Excess devices: ${details.limitViolation.excessDevices}`);
  console.log(`Severity: ${details.limitViolation.violationSeverity}`);
}
```

##### getEnrollmentStatisticsByUser()

```typescript
async getEnrollmentStatisticsByUser(
  options?: EnrollmentTrackingReportOptions
): Promise<EnrollmentStatistics>
```

**Description**: Generates comprehensive enrollment statistics aggregated by user.

**Returns**: Statistical analysis including averages, distributions, and trends.

**Example**:
```typescript
const stats = await report.getEnrollmentStatisticsByUser();

console.log(`Total users: ${stats.totalUsers}`);
console.log(`Average devices/user: ${stats.averageDevicesPerUser}`);
console.log(`Users violating limit: ${stats.usersViolatingLimit}`);
console.log(`Compliance rate: ${stats.complianceOverview.complianceRate}%`);
```

##### detectUnusualEnrollmentPatterns()

```typescript
async detectUnusualEnrollmentPatterns(
  options?: EnrollmentTrackingReportOptions
): Promise<UnusualEnrollmentPattern[]>
```

**Description**: Analyzes enrollment data to detect suspicious or unusual patterns.

**Detection Types**:
1. **Rapid Enrollment**: 5+ devices enrolled within 24 hours
2. **Bulk Enrollment**: 20+ devices enrolled on same day
3. **After-Hours Enrollment**: Enrollments outside business hours (10 PM - 6 AM UTC)
4. **Excessive Devices**: Users with 2x+ the device limit
5. **Suspicious Devices**: Devices with no name, never synced, or long-term non-compliant

**Returns**: Array of detected patterns sorted by severity.

**Example**:
```typescript
const patterns = await report.detectUnusualEnrollmentPatterns();

patterns.forEach(pattern => {
  console.log(`[${pattern.severity}] ${pattern.patternType}`);
  console.log(`Description: ${pattern.description}`);
  console.log(`Recommendation: ${pattern.recommendation}`);
  console.log(`Affected devices: ${pattern.affectedDevices.length}`);
});
```

## Features

### 1. Per-User Device Listing

**Capability**: Complete device inventory for each user

**Data Provided**:
- Device hardware details (manufacturer, model, serial number)
- OS and version information
- Enrollment date and method
- Management and compliance states
- Last sync timestamp
- Ownership type

**Use Cases**:
- User device audits
- Compliance verification
- Enrollment method analysis
- Device lifecycle tracking

### 2. Aggregated Counts by User

**Capability**: Statistical rollup of device enrollments

**Metrics Provided**:
- Total device count
- Platform-specific counts
- Corporate vs. Personal counts
- Compliant vs. Non-compliant counts
- Active vs. Inactive counts
- Enrollment trends (7/30/90/365 days)

**Use Cases**:
- Capacity planning
- License allocation
- Policy compliance monitoring
- Trend analysis

### 3. Device Limit Analysis

**Capability**: Automatic detection of users exceeding device limits

**Features**:
- Configurable device limit (default: 15)
- Violation severity calculation
- Excess device identification
- Licensed vs. unlicensed user tracking

**Severity Levels**:
- **Low**: 0-50% over limit
- **Medium**: 50-100% over limit
- **High**: 100-200% over limit (2x limit)
- **Critical**: 200%+ over limit (3x+ limit)

**Use Cases**:
- Policy enforcement
- Resource optimization
- Cost management
- Security compliance

### 4. Licensing Validation

**Capability**: Verification of user licensing for Intune enrollment

**License Types Detected**:
- Intune Plan 1
- Enterprise Mobility + Security (EMS) E3
- Enterprise Mobility + Security (EMS) E5
- Microsoft 365 E3
- Microsoft 365 E5

**Validation Checks**:
- User has assigned licenses
- Intune-specific license present
- EMS subscription active
- Microsoft 365 subscription active
- Account enabled status

**Use Cases**:
- License compliance audits
- Cost optimization
- Unauthorized enrollment detection
- License allocation planning

### 5. Enrollment Trend Tracking

**Capability**: Time-based enrollment analysis

**Time Periods**:
- Last 7 days
- Last 30 days
- Last 90 days
- Last 365 days

**Analysis**:
- Enrollment velocity
- Growth trends
- Seasonal patterns
- Department-specific trends

**Use Cases**:
- Capacity forecasting
- Deployment planning
- Onboarding metrics
- Adoption tracking

### 6. Unusual Pattern Detection

**Capability**: Automated anomaly detection for security and compliance

**Pattern Types**:

#### Rapid Enrollment
- **Trigger**: 5+ devices in 24 hours
- **Severity**: High/Critical
- **Risk**: Unauthorized bulk enrollment, compromised account

#### Bulk Enrollment
- **Trigger**: 20+ devices on same day
- **Severity**: Medium/Critical
- **Risk**: Unplanned mass deployment, script-based enrollment

#### After-Hours Enrollment
- **Trigger**: Enrollment outside 6 AM - 10 PM UTC
- **Severity**: Low/High
- **Risk**: Unauthorized access, global deployment without approval

#### Excessive Devices
- **Trigger**: 2x+ device limit
- **Severity**: High/Critical
- **Risk**: Policy violation, resource waste, security risk

#### Suspicious Devices
- **Trigger**: No device name, never synced, or long-term non-compliant
- **Severity**: Medium
- **Risk**: Orphaned devices, failed enrollments, stale records

**Use Cases**:
- Security monitoring
- Policy compliance
- Fraud detection
- Operational hygiene

### 7. Export Formats

**Supported Formats**:

#### JSON
```typescript
const reportData = await report.getEnrollmentCountPerUserReport();
const json = report.exportAsJSON(reportData);
```

#### CSV
```typescript
const reportData = await report.getEnrollmentCountPerUserReport();
const csv = report.exportAsCSV(reportData);
// Headers: User Principal Name, Display Name, Department, Total Devices, iOS, Android, etc.
```

#### HTML
```typescript
const reportData = await report.getEnrollmentCountPerUserReport();
const html = report.exportAsHTML(reportData);
// Styled, interactive HTML report with tables and summary cards
```

**Features**:
- Automatic format detection based on data type
- Responsive HTML layout
- Platform-specific color coding
- Summary statistics in all formats

## Usage Examples

### Example 1: Basic Enrollment Report

```typescript
import { IntuneEnrollmentTrackingReports } from './reports/configmgr/intune-enrollment-tracking-reports';
import { Client } from '@microsoft/microsoft-graph-client';

// Initialize
const graphClient = Client.init(/* auth config */);
const config = { deviceEnrollmentLimit: 15 };
const report = new IntuneEnrollmentTrackingReports(graphClient, config);

// Generate enrollment count report
const enrollmentReport = await report.getEnrollmentCountPerUserReport();

console.log(`Total users: ${enrollmentReport.summary.totalUsers}`);
console.log(`Total devices: ${enrollmentReport.summary.totalDevices}`);
console.log(`Avg devices/user: ${enrollmentReport.summary.averageDevicesPerUser}`);

// Export to CSV
const csv = report.exportAsCSV(enrollmentReport);
await fs.writeFile('enrollment-report.csv', csv);
```

### Example 2: Filter by Department

```typescript
// Get all devices for IT department
const itReport = await report.getDevicesPerUserReport({
  filters: {
    department: 'IT',
    sortBy: 'deviceCount',
    sortOrder: 'desc'
  }
});

console.log('IT Department Device Enrollment:');
itReport.data.forEach(user => {
  console.log(`${user.displayName}: ${user.totalDevices} devices`);
  console.log(`  iOS: ${user.devicesByPlatform.iOS}`);
  console.log(`  Android: ${user.devicesByPlatform.Android}`);
  console.log(`  Windows: ${user.devicesByPlatform.Windows}`);
});
```

### Example 3: Identify Users Exceeding Device Limits

```typescript
// Find users with more than 10 devices
const powerUsers = await report.getUsersWithMultipleDevices(10);

console.log(`Found ${powerUsers.length} users with 10+ devices`);

for (const user of powerUsers) {
  console.log(`\n${user.displayName} (${user.userPrincipalName})`);
  console.log(`  Total devices: ${user.totalDeviceCount}`);
  console.log(`  Corporate: ${user.corporateDeviceCount}`);
  console.log(`  Personal: ${user.personalDeviceCount}`);
  console.log(`  Compliant: ${user.compliantDeviceCount}`);
  console.log(`  Active: ${user.activeDeviceCount}`);
}
```

### Example 4: User Enrollment Audit

```typescript
// Get detailed enrollment info for specific user
const userEmail = 'john.doe@contoso.com';
const details = await report.getUserEnrollmentDetails(userEmail);

console.log(`User Enrollment Audit: ${details.enrollmentCount?.displayName}`);
console.log(`\nLicensing:`);
console.log(`  Licensed: ${details.licensing?.isLicensed}`);
console.log(`  Intune License: ${details.licensing?.intuneLicense}`);
console.log(`  EMS: ${details.licensing?.ems}`);
console.log(`  M365: ${details.licensing?.microsoft365}`);

console.log(`\nDevice Summary:`);
console.log(`  Total: ${details.enrollmentCount?.totalDeviceCount}`);
console.log(`  iOS: ${details.enrollmentCount?.iosDeviceCount}`);
console.log(`  Android: ${details.enrollmentCount?.androidDeviceCount}`);
console.log(`  Windows: ${details.enrollmentCount?.windowsDeviceCount}`);

console.log(`\nCompliance:`);
console.log(`  Compliant: ${details.enrollmentCount?.compliantDeviceCount}`);
console.log(`  Non-compliant: ${details.enrollmentCount?.nonCompliantDeviceCount}`);

if (details.violatesLimit) {
  console.log(`\n⚠️  LIMIT VIOLATION`);
  console.log(`  Severity: ${details.limitViolation?.violationSeverity}`);
  console.log(`  Device limit: ${details.limitViolation?.deviceLimit}`);
  console.log(`  Current count: ${details.limitViolation?.deviceCount}`);
  console.log(`  Excess: ${details.limitViolation?.excessDevices}`);
}

console.log(`\nDevice List:`);
details.user?.devices.forEach(device => {
  console.log(`  - ${device.deviceName} (${device.platform})`);
  console.log(`    Enrolled: ${device.enrollmentDate}`);
  console.log(`    Compliance: ${device.complianceState}`);
  console.log(`    Owner: ${device.ownerType}`);
});
```

### Example 5: Platform-Specific Analysis

```typescript
// Get iOS devices only
const iosReport = await report.getDevicesPerUserReport({
  filters: {
    platform: ['iOS'],
    complianceState: 'Compliant'
  }
});

console.log('Compliant iOS Devices by User:');
iosReport.data.forEach(user => {
  console.log(`${user.displayName}: ${user.totalDevices} iOS devices`);
});

// Get Android Enterprise enrollments
const androidReport = await report.getDevicesPerUserReport({
  filters: {
    platform: ['Android'],
    enrollmentMethod: ['androidEnterprise']
  }
});

console.log('\nAndroid Enterprise Enrollments:');
androidReport.data.forEach(user => {
  console.log(`${user.displayName}: ${user.totalDevices} Android Enterprise devices`);
});
```

### Example 6: Enrollment Statistics Dashboard

```typescript
// Generate comprehensive statistics
const stats = await report.getEnrollmentStatisticsByUser();

console.log('=== ENROLLMENT STATISTICS DASHBOARD ===\n');

console.log('Overall Metrics:');
console.log(`  Total users: ${stats.totalUsers}`);
console.log(`  Total devices: ${stats.totalDevices}`);
console.log(`  Average devices/user: ${stats.averageDevicesPerUser}`);
console.log(`  Median devices/user: ${stats.medianDevicesPerUser}`);

console.log('\nUser Distribution:');
console.log(`  Single device: ${stats.usersWithSingleDevice}`);
console.log(`  Multiple devices: ${stats.usersWithMultipleDevices}`);
console.log(`  Violating limit: ${stats.usersViolatingLimit}`);

console.log('\nPlatform Distribution:');
console.log(`  iOS: ${stats.platformDistribution.iOS}`);
console.log(`  Android: ${stats.platformDistribution.Android}`);
console.log(`  Windows: ${stats.platformDistribution.Windows}`);
console.log(`  macOS: ${stats.platformDistribution.macOS}`);

console.log('\nCompliance:');
console.log(`  Compliant: ${stats.complianceOverview.totalCompliant}`);
console.log(`  Non-compliant: ${stats.complianceOverview.totalNonCompliant}`);
console.log(`  Rate: ${stats.complianceOverview.complianceRate}%`);

console.log('\nEnrollment Trends:');
console.log(`  Last 7 days: ${stats.enrollmentTrends.last7Days}`);
console.log(`  Last 30 days: ${stats.enrollmentTrends.last30Days}`);
console.log(`  Last 90 days: ${stats.enrollmentTrends.last90Days}`);
console.log(`  Last year: ${stats.enrollmentTrends.last365Days}`);

console.log('\nTop Departments:');
stats.departmentBreakdown.slice(0, 5).forEach((dept, i) => {
  console.log(`  ${i + 1}. ${dept.department}`);
  console.log(`     Users: ${dept.userCount}`);
  console.log(`     Devices: ${dept.deviceCount}`);
  console.log(`     Avg: ${dept.averageDevicesPerUser}`);
});

console.log('\nTop Users by Device Count:');
stats.topUsersByDeviceCount.slice(0, 10).forEach((user, i) => {
  console.log(`  ${i + 1}. ${user.displayName}: ${user.totalDeviceCount} devices`);
});
```

### Example 7: Security Anomaly Detection

```typescript
// Detect unusual enrollment patterns
const patterns = await report.detectUnusualEnrollmentPatterns();

console.log('=== SECURITY ANOMALY REPORT ===\n');
console.log(`Total patterns detected: ${patterns.length}\n`);

// Group by severity
const critical = patterns.filter(p => p.severity === 'Critical');
const high = patterns.filter(p => p.severity === 'High');
const medium = patterns.filter(p => p.severity === 'Medium');
const low = patterns.filter(p => p.severity === 'Low');

if (critical.length > 0) {
  console.log('🚨 CRITICAL ISSUES:');
  critical.forEach(p => {
    console.log(`  Type: ${p.patternType}`);
    console.log(`  Description: ${p.description}`);
    console.log(`  Affected: ${p.deviceCount} devices`);
    console.log(`  Recommendation: ${p.recommendation}`);
    if (p.userPrincipalName) {
      console.log(`  User: ${p.userPrincipalName}`);
    }
    console.log('');
  });
}

if (high.length > 0) {
  console.log('⚠️  HIGH PRIORITY:');
  high.forEach(p => {
    console.log(`  ${p.patternType}: ${p.description}`);
    console.log(`  Action: ${p.recommendation}\n`);
  });
}

if (medium.length > 0) {
  console.log('ℹ️  MEDIUM PRIORITY:');
  medium.forEach(p => {
    console.log(`  ${p.patternType}: ${p.description}\n`);
  });
}

// Export detailed pattern report
const patternReport = {
  metadata: {
    generatedAt: new Date().toISOString(),
    totalPatterns: patterns.length,
    severityBreakdown: {
      critical: critical.length,
      high: high.length,
      medium: medium.length,
      low: low.length
    }
  },
  patterns
};

await fs.writeFile(
  'anomaly-report.json',
  JSON.stringify(patternReport, null, 2)
);
```

### Example 8: License Compliance Audit

```typescript
// Find unlicensed users with enrolled devices
const allUsers = await report.getAllUsersWithDeviceCounts();

const unlicensedEnrollments = [];

for (const user of allUsers) {
  const details = await report.getUserEnrollmentDetails(user.userPrincipalName);

  if (details.licensing && !details.licensing.isLicensed) {
    unlicensedEnrollments.push({
      user: user.displayName,
      upn: user.userPrincipalName,
      deviceCount: user.totalDeviceCount,
      accountEnabled: details.licensing.accountEnabled
    });
  }
}

console.log('=== LICENSE COMPLIANCE AUDIT ===\n');
console.log(`Unlicensed users with devices: ${unlicensedEnrollments.length}\n`);

unlicensedEnrollments.forEach(item => {
  console.log(`User: ${item.user}`);
  console.log(`  Email: ${item.upn}`);
  console.log(`  Devices: ${item.deviceCount}`);
  console.log(`  Account enabled: ${item.accountEnabled}`);
  console.log('');
});
```

### Example 9: Date Range Filtering

```typescript
// Get enrollments from last quarter
const threeMonthsAgo = new Date();
threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);

const q4Report = await report.getEnrollmentCountPerUserReport({
  filters: {
    startDate: threeMonthsAgo,
    endDate: new Date(),
    sortBy: 'lastEnrollment',
    sortOrder: 'desc'
  }
});

console.log('Last Quarter Enrollments:');
q4Report.data.forEach(user => {
  console.log(`${user.displayName}: ${user.totalDeviceCount} devices`);
  console.log(`  Last enrollment: ${user.lastEnrollmentDate}`);
});
```

### Example 10: Corporate vs. Personal Device Analysis

```typescript
// Analyze corporate vs personal device distribution
const stats = await report.getEnrollmentStatisticsByUser();

const totalCorporate = stats.platformDistribution.iOS +
                       stats.platformDistribution.Android +
                       stats.platformDistribution.Windows +
                       stats.platformDistribution.macOS;

console.log('=== DEVICE OWNERSHIP ANALYSIS ===\n');

// Get detailed breakdown
const allUsers = await report.getAllUsersWithDeviceCounts();

let totalCorp = 0;
let totalPersonal = 0;

allUsers.forEach(user => {
  totalCorp += user.corporateDeviceCount;
  totalPersonal += user.personalDeviceCount;
});

console.log(`Total corporate devices: ${totalCorp}`);
console.log(`Total personal devices: ${totalPersonal}`);
console.log(`Corporate percentage: ${((totalCorp / (totalCorp + totalPersonal)) * 100).toFixed(2)}%`);

// Find users with only personal devices
const personalOnly = allUsers.filter(u =>
  u.personalDeviceCount > 0 && u.corporateDeviceCount === 0
);

console.log(`\nUsers with only personal devices: ${personalOnly.length}`);
personalOnly.slice(0, 10).forEach(user => {
  console.log(`  ${user.displayName}: ${user.personalDeviceCount} personal devices`);
});
```

## API Reference

### Class: IntuneEnrollmentTrackingReports

**Extends**: `BaseReport`

#### Constructor

```typescript
constructor(graphClient: Client, config: AppConfig)
```

#### Properties

| Property | Type | Description |
|----------|------|-------------|
| `name` | `string` | Report name identifier |
| `description` | `string` | Report description |
| `category` | `string` | Report category (Configuration Manager Reports) |
| `enabled` | `boolean` | Report enabled status |

#### Methods

| Method | Parameters | Return Type | Description |
|--------|-----------|-------------|-------------|
| `execute()` | `options?: EnrollmentTrackingReportOptions` | `Promise<ReportData>` | Execute default enrollment report |
| `getDevicesPerUserReport()` | `options?: EnrollmentTrackingReportOptions` | `Promise<ReportData>` | Report 17: Device list per user |
| `getEnrollmentCountPerUserReport()` | `options?: EnrollmentTrackingReportOptions` | `Promise<ReportData>` | Report 30: Enrollment count per user |
| `getDevicesEnrolledPerUser()` | `userIdOrUpn: string` | `Promise<UserDeviceList \| null>` | Get devices for specific user |
| `getAllUsersWithDeviceCounts()` | `options?: EnrollmentTrackingReportOptions` | `Promise<UserEnrollmentCount[]>` | Get all user enrollment counts |
| `getUsersWithMultipleDevices()` | `threshold?: number, options?: EnrollmentTrackingReportOptions` | `Promise<UserEnrollmentCount[]>` | Get users exceeding threshold |
| `getUserEnrollmentDetails()` | `userIdOrUpn: string` | `Promise<UserEnrollmentDetails>` | Get comprehensive user details |
| `getEnrollmentStatisticsByUser()` | `options?: EnrollmentTrackingReportOptions` | `Promise<EnrollmentStatistics>` | Get enrollment statistics |
| `detectUnusualEnrollmentPatterns()` | `options?: EnrollmentTrackingReportOptions` | `Promise<UnusualEnrollmentPattern[]>` | Detect anomalies |
| `generateEnrollmentStatistics()` | `filters?: EnrollmentTrackingFilters` | `Promise<EnrollmentTrackingStatistics>` | Generate overall statistics |
| `exportAsJSON()` | `reportData: ReportData` | `string` | Export report as JSON |
| `exportAsCSV()` | `reportData: ReportData` | `string` | Export report as CSV |
| `exportAsHTML()` | `reportData: ReportData` | `string` | Export report as HTML |

## Testing

### Unit Tests

The module includes comprehensive unit tests covering:

1. **Report Generation**
   - Device list per user (Report 17)
   - Enrollment count per user (Report 30)
   - Statistics generation

2. **Filtering and Sorting**
   - Platform filtering
   - Department filtering
   - Date range filtering
   - Device count filtering
   - Multi-field sorting

3. **User Methods**
   - Get devices by UPN
   - Get devices by user ID
   - Multiple device detection
   - Enrollment details retrieval

4. **Pattern Detection**
   - Rapid enrollment detection
   - Bulk enrollment detection
   - After-hours detection
   - Excessive device detection
   - Suspicious device detection

5. **Export Functions**
   - JSON export
   - CSV export
   - HTML export

6. **Error Handling**
   - Graph API errors
   - Empty result sets
   - Missing user data

### Running Tests

```bash
# Run all tests
npm test intune-enrollment-tracking-reports.test.ts

# Run with coverage
npm test -- --coverage intune-enrollment-tracking-reports.test.ts

# Run specific test suite
npm test -- --grep "getDevicesPerUserReport"
```

### Test Coverage

- **Statements**: 95%+
- **Branches**: 90%+
- **Functions**: 95%+
- **Lines**: 95%+

## Best Practices

### 1. Performance Optimization

```typescript
// ✅ Good: Use filters to reduce data volume
const report = await reports.getEnrollmentCountPerUserReport({
  filters: {
    department: 'IT',
    top: 100
  }
});

// ❌ Bad: Fetch all data then filter in memory
const allData = await reports.getAllUsersWithDeviceCounts();
const itData = allData.filter(u => u.department === 'IT');
```

### 2. Error Handling

```typescript
// ✅ Good: Handle errors gracefully
try {
  const details = await report.getUserEnrollmentDetails(userId);
  if (!details.user) {
    console.warn(`User ${userId} not found or has no devices`);
    return;
  }
  // Process details
} catch (error) {
  logger.error('Failed to get enrollment details', error);
  // Fallback logic
}

// ❌ Bad: No error handling
const details = await report.getUserEnrollmentDetails(userId);
const deviceCount = details.user.totalDevices; // May throw if user is null
```

### 3. Batch Processing

```typescript
// ✅ Good: Process users in batches
const allUsers = await report.getAllUsersWithDeviceCounts();
const batchSize = 10;

for (let i = 0; i < allUsers.length; i += batchSize) {
  const batch = allUsers.slice(i, i + batchSize);

  await Promise.all(batch.map(async user => {
    const details = await report.getUserEnrollmentDetails(user.userPrincipalName);
    // Process details
  }));

  // Rate limiting delay
  await new Promise(resolve => setTimeout(resolve, 1000));
}

// ❌ Bad: Process all users simultaneously
await Promise.all(allUsers.map(user =>
  report.getUserEnrollmentDetails(user.userPrincipalName)
)); // May hit API rate limits
```

### 4. Caching Results

```typescript
// ✅ Good: Cache results for repeated access
class EnrollmentReportService {
  private cache = new Map<string, { data: any; timestamp: number }>();
  private cacheTTL = 5 * 60 * 1000; // 5 minutes

  async getUserDevices(userId: string): Promise<UserDeviceList | null> {
    const cached = this.cache.get(userId);

    if (cached && Date.now() - cached.timestamp < this.cacheTTL) {
      return cached.data;
    }

    const data = await report.getDevicesEnrolledPerUser(userId);
    this.cache.set(userId, { data, timestamp: Date.now() });

    return data;
  }
}
```

### 5. Logging and Monitoring

```typescript
// ✅ Good: Comprehensive logging
logger.info('Starting enrollment pattern detection');

const startTime = Date.now();
const patterns = await report.detectUnusualEnrollmentPatterns();
const duration = Date.now() - startTime;

logger.info('Pattern detection complete', {
  patternsFound: patterns.length,
  durationMs: duration,
  criticalCount: patterns.filter(p => p.severity === 'Critical').length
});

patterns.forEach(pattern => {
  if (pattern.severity === 'Critical' || pattern.severity === 'High') {
    logger.warn('Unusual enrollment pattern detected', {
      type: pattern.patternType,
      severity: pattern.severity,
      affectedDevices: pattern.deviceCount
    });
  }
});
```

### 6. Configuration Management

```typescript
// ✅ Good: Externalize configuration
const config: AppConfig = {
  clientId: process.env.AZURE_CLIENT_ID!,
  tenantId: process.env.AZURE_TENANT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  deviceEnrollmentLimit: parseInt(process.env.DEVICE_LIMIT || '15'),
  rapidEnrollmentThreshold: parseInt(process.env.RAPID_THRESHOLD || '5'),
  rapidEnrollmentWindow: parseInt(process.env.RAPID_WINDOW || '24')
};

// ❌ Bad: Hardcoded values
const limit = 15; // What if this changes?
```

### 7. Security Considerations

```typescript
// ✅ Good: Validate and sanitize inputs
async function getUserReport(userInput: string): Promise<UserDeviceList | null> {
  // Validate UPN format
  const upnRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

  if (!upnRegex.test(userInput) && !uuidRegex.test(userInput)) {
    throw new Error('Invalid user identifier format');
  }

  return await report.getDevicesEnrolledPerUser(userInput);
}

// ✅ Good: Implement least privilege
// Only request necessary permissions in Graph API calls
const response = await graphClient
  .api('/deviceManagement/managedDevices')
  .select(['id', 'deviceName', 'userPrincipalName']) // Only needed fields
  .top(100) // Limit results
  .get();
```

---

## Summary

The **Intune Enrollment Tracking Reports** module provides comprehensive enrollment tracking and analysis capabilities that replicate and extend Configuration Manager functionality using Microsoft Graph API.

**Key Capabilities**:
- Detailed per-user device inventories
- Aggregated enrollment statistics
- Device limit monitoring and violation detection
- User licensing validation
- Enrollment trend analysis
- Security anomaly detection
- Multiple export formats

**Primary Use Cases**:
- User device audits
- License compliance verification
- Policy enforcement
- Security monitoring
- Capacity planning
- Trend analysis
- Operational reporting

**Integration Points**:
- Microsoft Graph API
- Azure AD/Entra ID
- Microsoft Intune
- Configuration Manager (legacy compatibility)

For additional support or questions, refer to the main project documentation or contact the development team.
