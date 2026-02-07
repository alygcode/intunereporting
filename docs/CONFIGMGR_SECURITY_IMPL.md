# Configuration Manager Device Security Reports for Intune

This guide provides comprehensive documentation for the Device Security Reports module, which replicates Configuration Manager security reports for Intune environments using Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Configuration Manager Reports Mapping](#configuration-manager-reports-mapping)
- [API Reference](#api-reference)
- [Usage Examples](#usage-examples)
- [Microsoft Graph API Endpoints](#microsoft-graph-api-endpoints)
- [Data Models](#data-models)
- [Security Risk Analysis](#security-risk-analysis)
- [Remediation Guidance](#remediation-guidance)
- [Export Formats](#export-formats)
- [Best Practices](#best-practices)
- [Troubleshooting](#troubleshooting)
- [Performance Considerations](#performance-considerations)

## Overview

The Device Security Reports module provides Intune equivalents for Configuration Manager's critical security reports. It enables comprehensive security monitoring for mobile devices, certificate management, and Windows device health attestation.

### Implemented Reports

This module implements the following Configuration Manager reports for Intune:

1. **Report #23**: Mobile devices that are jailbroken or rooted - Security risk device detection
2. **Report #27**: Mobile devices with certificate renewal issues - Certificate expiration and renewal failures
3. **Report #16**: List of devices by Health Attestation state - Windows device health attestation attributes

### Key Features

- **Jailbreak/Root Detection**: Identify compromised iOS and Android devices with risk scoring
- **Certificate Monitoring**: Track certificate expiration, renewal failures, and compliance
- **Health Attestation**: Monitor Windows 10/11 security features (BitLocker, Secure Boot, TPM, etc.)
- **Risk Assessment**: Automated risk level calculation and security scoring
- **Remediation Guidance**: Actionable recommendations for each security issue
- **Multiple Export Formats**: JSON, CSV, and HTML output with visual risk indicators
- **Type-Safe**: Full TypeScript support with comprehensive interfaces
- **Pagination Support**: Handles large device inventories efficiently
- **Advanced Filtering**: Filter by OS, compliance state, risk level, severity, and more

## Prerequisites

### Software Requirements

- Node.js 18.0.0 or higher
- TypeScript 5.0 or higher
- npm or yarn package manager

### Azure Requirements

- Azure AD (Entra ID) tenant
- Intune Administrator or Global Administrator access
- Microsoft Intune subscription with enrolled devices
- Devices enrolled in Intune (mobile devices, Windows 10/11)

### Required API Permissions

Your Azure AD app registration needs the following Microsoft Graph permissions:

**Application Permissions (Required):**
- `DeviceManagementManagedDevices.Read.All` - Read managed device properties and health state
- `DeviceManagementConfiguration.Read.All` - Read device configuration and compliance states

**Application Permissions (Optional, for extended features):**
- `User.Read.All` - Read user profiles for detailed reporting
- `Directory.Read.All` - Read directory data for organizational context

## Installation

### 1. Install Dependencies

```bash
npm install @microsoft/microsoft-graph-client @azure/identity json2csv
npm install --save-dev @types/node typescript vitest
```

### 2. Configure Environment Variables

Create a `.env` file in your project root:

```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
```

### 3. Import the Module

```typescript
import {
  DeviceSecurityReports,
  JailbrokenDevice,
  CertificateIssue,
  HealthAttestationState,
  ReportFilters,
} from './src/reports/configmgr/device-security-reports';
```

## Quick Start

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { DeviceSecurityReports } from './device-security-reports';

// Create authenticated Graph client
const credential = new ClientSecretCredential(
  process.env.AZURE_TENANT_ID!,
  process.env.AZURE_CLIENT_ID!,
  process.env.AZURE_CLIENT_SECRET!
);

const authProvider = {
  getAccessToken: async () => {
    const token = await credential.getToken('https://graph.microsoft.com/.default');
    return token.token;
  },
};

const graphClient = Client.initWithMiddleware({
  authProvider: {
    getAccessToken: authProvider.getAccessToken,
  },
});

// Create security reports instance
const securityReports = new DeviceSecurityReports(graphClient);

// Generate all three security reports
const [jailbreakReport, certReport, healthReport] = await Promise.all([
  securityReports.generateJailbreakReport(),
  securityReports.generateCertificateRenewalReport(),
  securityReports.generateHealthAttestationReport(),
]);

console.log('=== Security Summary ===');
console.log(`Jailbroken Devices: ${jailbreakReport.summary.jailbrokenDevices}`);
console.log(`Certificate Issues: ${certReport.summary.devicesWithIssues}`);
console.log(`Health Attestation Devices: ${healthReport.summary.totalDevices}`);
```

## Configuration Manager Reports Mapping

### Report #23: Mobile Devices That Are Jailbroken or Rooted

**Configuration Manager**: Displays all mobile devices that have been jailbroken (iOS) or rooted (Android), presenting a significant security risk.

**Intune Implementation**:
```typescript
const jailbrokenDevices = await securityReports.getJailbrokenOrRootedDevices();
// Alternative method name:
const jailbrokenDevices = await securityReports.getJailbrokenDevices();
```

**Microsoft Graph API Used**:
- `GET /deviceManagement/managedDevices?$filter=jailBroken eq true`

**Data Returned**:
- Device identification (ID, name, model, manufacturer)
- User information (UPN, email)
- Operating system and version
- Jailbreak/root status
- **Risk Level**: Critical, High, Medium, or Low
- **Risk Score**: Numerical score 0-100
- Compliance state
- Management state
- Last sync and enrollment dates
- **Security Notes**: Detailed risk assessment findings

**Security Risk Scoring Algorithm**:

The module automatically calculates risk levels based on multiple factors:

1. **Base Risk** (40 points): Device is jailbroken/rooted
2. **Compliance Impact** (20 points): Device is non-compliant
3. **Platform Risk**:
   - iOS jailbroken: +15 points
   - Android rooted: +10 points
4. **Device Age** (15 points): Older devices with jailbreak = higher risk
5. **Sync Status** (10 points): Devices not syncing regularly = higher risk

**Risk Levels**:
- **Critical** (80-100): Immediate action required
- **High** (60-79): Urgent attention needed
- **Medium** (40-59): Monitor closely
- **Low** (0-39): Standard monitoring

**Use Cases**:
- Identify security-compromised devices in your environment
- Generate risk-based prioritization for remediation
- Monitor security compliance for mobile devices
- Track trends in jailbreak/root detection over time
- Block access for high-risk devices
- Generate security compliance reports for auditing

**Security Notes Generated**:
- "Device is jailbroken/rooted - high security risk"
- "Non-compliant due to jailbreak/root status"
- "Device has not synced recently - potential management bypass"
- "Older device with jailbreak - elevated security risk"

### Report #27: Mobile Devices with Certificate Renewal Issues

**Configuration Manager**: Shows mobile devices experiencing certificate expiration or renewal failures, which can impact connectivity and security.

**Intune Implementation**:
```typescript
const certificateIssues = await securityReports.getDevicesWithCertificateIssues();
// Alternative method name:
const certificateIssues = await securityReports.getCertificateRenewalIssues();
```

**Microsoft Graph API Used**:
- `GET /deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'`
- `GET /deviceManagement/managedDevices/{id}/deviceComplianceDeviceStatuses`
- `GET /deviceManagement/managedDevices/{id}/deviceConfigurationStates`

**Data Returned**:
- Device and user information
- **Certificate Type**: Device, User, WiFi, VPN, SCEP, PKCS, or Other
- Certificate issuer and subject
- **Expiration Date**: When certificate expires
- **Days Until Expiration**: Countdown to expiration
- **Status**: Expired, Expiring, Renewal Failed, Invalid, Revoked, or Valid
- **Severity**: Critical, High, Medium, or Low
- Issue description and context
- **Remediation Action**: Specific steps to resolve
- Last renewal attempt timestamp
- Renewal error messages
- Certificate thumbprint
- Compliance state

**Certificate Issue Severity Levels**:

1. **Critical**:
   - Certificate already expired
   - Certificate renewal failed multiple times
   - Certificate revoked
   - Expires within 7 days

2. **High**:
   - Certificate expires within 8-14 days
   - Renewal attempt failed once
   - Invalid certificate detected

3. **Medium**:
   - Certificate expires within 15-30 days
   - First renewal attempt pending

4. **Low**:
   - Certificate expires in 31-60 days
   - Valid certificate with upcoming renewal

**Certificate Types Monitored**:

- **Device Certificates**: Device authentication certificates
- **User Certificates**: User authentication certificates
- **WiFi Certificates**: Wireless network authentication
- **VPN Certificates**: VPN connection authentication
- **SCEP Certificates**: Simple Certificate Enrollment Protocol
- **PKCS Certificates**: Public Key Cryptography Standards
- **Other**: Custom or specialized certificates

**Use Cases**:
- Proactive certificate expiration monitoring
- Identify and resolve renewal failures
- Prevent connectivity disruptions due to expired certificates
- Track certificate lifecycle management
- Generate alerts for critical certificate issues
- Compliance reporting for certificate policies

**Remediation Actions Generated**:
- "Renew certificate immediately - expired"
- "Troubleshoot SCEP renewal failure - check CA connectivity"
- "Redeploy certificate profile to device"
- "Verify user enrollment for user certificate"
- "Check network connectivity for automatic renewal"
- "Manual certificate installation may be required"

### Report #16: List of Devices by Health Attestation State

**Configuration Manager**: Displays Windows 10/11 devices with their hardware and firmware security attributes verified through health attestation.

**Intune Implementation**:
```typescript
const healthDevices = await securityReports.getDevicesByHealthAttestationState();
```

**Microsoft Graph API Used**:
- `GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'Windows'`
- `GET /deviceManagement/managedDevices/{id}/deviceHealthAttestationState`

**Health Attestation Properties**:

**Core Security Features**:
- **Health Attestation Supported**: Whether device supports health attestation
- **BitLocker Status**: Enabled, Disabled, or Unknown
- **Secure Boot**: Boot integrity verification
- **Code Integrity**: OS code signing enforcement
- **Boot Debugging**: Kernel debugging status (should be disabled)
- **Kernel Debugging**: OS kernel debugging status (should be disabled)
- **Test Signing**: Test-signed drivers allowed (should be disabled)
- **ELAM (Early Launch Anti-Malware)**: Anti-malware boot protection
- **Virtual Secure Mode (VSM)**: Credential Guard and other virtualization-based security
- **TPM Version**: Trusted Platform Module version (1.2 or 2.0)
- **PCR Hash Algorithm**: Platform Configuration Register hash algorithm
- **Boot Manager Version**: Windows Boot Manager version
- **Code Integrity Check Version**: CI policy version

**Health Score Calculation**:

Each device receives a health score (0-100) based on security feature compliance:

**Scoring Matrix**:
- BitLocker Enabled: +15 points
- Secure Boot Enabled: +15 points
- Code Integrity Enabled: +10 points
- ELAM Enabled: +10 points
- Virtual Secure Mode Enabled: +15 points
- TPM 2.0 Present: +10 points
- Boot Debugging Disabled: +10 points
- Kernel Debugging Disabled: +5 points
- Test Signing Disabled: +10 points (Critical)
- Health Attestation Supported: +0 points (required baseline)

**Health Score Ranges**:
- **Excellent** (90-100): Full security compliance
- **Good** (70-89): Minor improvements needed
- **Fair** (50-69): Multiple security gaps
- **Poor** (30-49): Significant security concerns
- **Critical** (0-29): Immediate remediation required

**Critical Security Issues Detected**:

The module automatically identifies and reports critical security issues:

1. **Test Signing Enabled** (Critical)
   - Allows unsigned or test-signed drivers
   - Major security vulnerability
   - Recommendation: Disable test signing via Group Policy or MDM

2. **BitLocker Disabled** (High)
   - Data at rest not encrypted
   - Risk of data theft if device lost/stolen
   - Recommendation: Enable BitLocker via Intune policy

3. **Secure Boot Disabled** (High)
   - Boot process not protected
   - Vulnerable to bootkit malware
   - Recommendation: Enable Secure Boot in UEFI/BIOS

4. **Boot/Kernel Debugging Enabled** (High)
   - Kernel debugging active
   - Potential for system exploitation
   - Recommendation: Disable kernel debugging

5. **Code Integrity Disabled** (Medium)
   - Driver signing not enforced
   - Risk of malicious drivers
   - Recommendation: Enable code integrity via policies

6. **VSM/Credential Guard Disabled** (Medium)
   - Advanced credential protection not active
   - Recommendation: Enable VSM if hardware supports

7. **TPM 1.2 or Missing** (Medium)
   - Older or no TPM
   - Limited security features
   - Recommendation: Upgrade to TPM 2.0 hardware

**Data Returned**:
- Device identification and user information
- Operating system details
- All health attestation properties
- **Health Score**: 0-100 numerical score
- **Health Score Range**: Excellent, Good, Fair, Poor, or Critical
- **Security Issues**: List of identified security problems
- **Recommendations**: Specific remediation actions

**Use Cases**:
- Monitor Windows 10/11 security feature adoption
- Identify devices not meeting security baselines
- Track BitLocker deployment status
- Verify Secure Boot and TPM usage
- Detect devices with debugging enabled
- Generate compliance reports for security standards
- Prioritize remediation efforts based on health scores

**Security Recommendations Generated**:
- "Enable BitLocker encryption for data protection"
- "Enable Secure Boot to protect boot process"
- "Disable test signing (CRITICAL SECURITY RISK)"
- "Disable boot/kernel debugging"
- "Enable Code Integrity to enforce driver signing"
- "Enable Virtual Secure Mode for Credential Guard"
- "Upgrade to TPM 2.0 hardware for enhanced security"
- "Enable ELAM for boot-time malware protection"

## API Reference

### Class: DeviceSecurityReports

#### Constructor

```typescript
constructor(graphClient: Client)
```

Creates a new instance of DeviceSecurityReports.

**Parameters**:
- `graphClient` - Authenticated Microsoft Graph client

#### Methods

##### getJailbrokenOrRootedDevices()

```typescript
async getJailbrokenOrRootedDevices(
  filters?: ReportFilters
): Promise<JailbrokenDevice[]>
```

Retrieves all jailbroken (iOS) or rooted (Android) devices with risk assessment.

**Parameters**:
- `filters` (optional) - Report filters to apply

**Returns**: Array of jailbroken device records

**Example**:
```typescript
const devices = await securityReports.getJailbrokenOrRootedDevices({
  riskLevel: ['Critical', 'High'],
  complianceState: ['noncompliant']
});
```

##### getJailbrokenDevices()

Alias for `getJailbrokenOrRootedDevices()`.

##### getDevicesWithCertificateIssues()

```typescript
async getDevicesWithCertificateIssues(
  filters?: ReportFilters
): Promise<CertificateIssue[]>
```

Retrieves devices with certificate expiration or renewal issues.

**Parameters**:
- `filters` (optional) - Report filters to apply

**Returns**: Array of certificate issue records

**Example**:
```typescript
const issues = await securityReports.getDevicesWithCertificateIssues({
  severity: ['Critical', 'High'],
  complianceState: ['noncompliant']
});
```

##### getCertificateRenewalIssues()

Alias for `getDevicesWithCertificateIssues()`.

##### getDevicesByHealthAttestationState()

```typescript
async getDevicesByHealthAttestationState(
  filters?: ReportFilters
): Promise<HealthAttestationState[]>
```

Retrieves Windows devices with health attestation information and security scoring.

**Parameters**:
- `filters` (optional) - Report filters to apply

**Returns**: Array of health attestation state records

**Example**:
```typescript
const devices = await securityReports.getDevicesByHealthAttestationState({
  operatingSystem: ['Windows'],
  complianceState: ['noncompliant']
});

// Filter for poor health
const poorHealth = devices.filter(d =>
  d.healthScoreRange === 'Poor' || d.healthScoreRange === 'Critical'
);
```

##### generateJailbreakReport()

```typescript
async generateJailbreakReport(
  filters?: ReportFilters
): Promise<{
  devices: JailbrokenDevice[];
  summary: JailbreakReportSummary;
}>
```

Generates comprehensive jailbreak detection report with summary statistics.

**Returns**: Object containing devices array and summary

**Example**:
```typescript
const report = await securityReports.generateJailbreakReport();
console.log(`Found ${report.summary.jailbrokenDevices} jailbroken devices`);
console.log(`Risk Score: ${report.summary.riskAssessment.overallRiskScore}/100`);
```

##### generateCertificateRenewalReport()

```typescript
async generateCertificateRenewalReport(
  filters?: ReportFilters
): Promise<{
  issues: CertificateIssue[];
  summary: CertificateReportSummary;
}>
```

Generates comprehensive certificate renewal issues report with summary.

**Returns**: Object containing issues array and summary

**Example**:
```typescript
const report = await securityReports.generateCertificateRenewalReport();
console.log(`Devices with issues: ${report.summary.devicesWithIssues}`);
console.log(`Expired certificates: ${report.summary.expiredCertificates}`);
```

##### generateHealthAttestationReport()

```typescript
async generateHealthAttestationReport(
  filters?: ReportFilters
): Promise<{
  devices: HealthAttestationState[];
  summary: HealthAttestationReportSummary;
}>
```

Generates comprehensive health attestation report with summary statistics.

**Returns**: Object containing devices array and summary

**Example**:
```typescript
const report = await securityReports.generateHealthAttestationReport();
console.log(`Total Devices: ${report.summary.totalDevices}`);
console.log(`BitLocker Enabled: ${report.summary.bitLockerPercentage}%`);
console.log(`Average Health Score: ${report.summary.averageHealthScore}/100`);
```

### Report Filters

```typescript
interface ReportFilters {
  /** Filter by operating system */
  operatingSystem?: string[];
  /** Filter by compliance state */
  complianceState?: string[];
  /** Filter by risk level */
  riskLevel?: ('Critical' | 'High' | 'Medium' | 'Low')[];
  /** Filter by severity */
  severity?: ('Critical' | 'High' | 'Medium' | 'Low')[];
  /** Filter by user principal names */
  userPrincipalNames?: string[];
  /** Filter by date range */
  dateFrom?: Date;
  dateTo?: Date;
}
```

## Usage Examples

### Example 1: Identify Critical Security Risks

```typescript
import { DeviceSecurityReports } from './device-security-reports';

const securityReports = new DeviceSecurityReports(graphClient);

// Get critical jailbroken devices
const criticalDevices = await securityReports.getJailbrokenOrRootedDevices({
  riskLevel: ['Critical'],
  complianceState: ['noncompliant']
});

console.log(`\n=== CRITICAL JAILBROKEN DEVICES ===`);
criticalDevices.forEach(device => {
  console.log(`\nDevice: ${device.deviceName}`);
  console.log(`User: ${device.userPrincipalName}`);
  console.log(`Platform: ${device.operatingSystem} ${device.osVersion}`);
  console.log(`Risk Score: ${device.riskScore}/100`);
  console.log(`Compliance: ${device.complianceState}`);
  console.log(`Security Notes:`);
  device.securityNotes.forEach(note => console.log(`  - ${note}`));
});

// Block access for critical devices (example)
for (const device of criticalDevices) {
  console.log(`\n⚠️ ACTION REQUIRED: Block device ${device.deviceName}`);
  console.log(`   Recommendation: Initiate remote wipe and require re-enrollment`);
}
```

### Example 2: Certificate Expiration Dashboard

```typescript
const certReport = await securityReports.generateCertificateRenewalReport();

console.log(`\n=== CERTIFICATE EXPIRATION DASHBOARD ===`);
console.log(`Total Devices Monitored: ${certReport.summary.totalDevices}`);
console.log(`Devices with Issues: ${certReport.summary.devicesWithIssues}`);
console.log(`\nExpiration Status:`);
console.log(`  Already Expired: ${certReport.summary.expiredCertificates}`);
console.log(`  Expiring Soon (30 days): ${certReport.summary.expiringSoonCertificates}`);
console.log(`  Renewal Failures: ${certReport.summary.failedRenewals}`);

console.log(`\nBy Severity:`);
console.log(`  Critical: ${certReport.summary.bySeverity.critical}`);
console.log(`  High: ${certReport.summary.bySeverity.high}`);
console.log(`  Medium: ${certReport.summary.bySeverity.medium}`);
console.log(`  Low: ${certReport.summary.bySeverity.low}`);

console.log(`\nBy Certificate Type:`);
Object.entries(certReport.summary.byCertificateType).forEach(([type, count]) => {
  console.log(`  ${type}: ${count}`);
});

// Export critical issues
const criticalIssues = certReport.issues.filter(i => i.severity === 'Critical');
await fs.writeFile(
  './reports/critical-cert-issues.json',
  JSON.stringify(criticalIssues, null, 2)
);
```

### Example 3: Health Attestation Compliance Report

```typescript
const healthReport = await securityReports.generateHealthAttestationReport();

console.log(`\n=== WINDOWS HEALTH ATTESTATION COMPLIANCE ===`);
console.log(`Total Windows Devices: ${healthReport.summary.totalDevices}`);
console.log(`Health Attestation Supported: ${healthReport.summary.supportedDevices}`);
console.log(`Average Health Score: ${healthReport.summary.averageHealthScore}/100`);

console.log(`\nBitLocker Status:`);
console.log(`  Enabled: ${healthReport.summary.bitLockerEnabled} (${healthReport.summary.bitLockerPercentage}%)`);
console.log(`  Disabled: ${healthReport.summary.bitLockerDisabled}`);

console.log(`\nSecure Boot Status:`);
console.log(`  Enabled: ${healthReport.summary.secureBootEnabled} (${healthReport.summary.secureBootPercentage}%)`);
console.log(`  Disabled: ${healthReport.summary.secureBootDisabled}`);

console.log(`\nHealth Score Distribution:`);
console.log(`  Excellent (90-100): ${healthReport.summary.excellentHealthDevices}`);
console.log(`  Good (70-89): ${healthReport.summary.goodHealthDevices}`);
console.log(`  Fair (50-69): ${healthReport.summary.fairHealthDevices}`);
console.log(`  Poor (30-49): ${healthReport.summary.poorHealthDevices}`);
console.log(`  Critical (0-29): ${healthReport.summary.criticalHealthDevices}`);

console.log(`\nCritical Security Issues:`);
healthReport.summary.criticalIssues.forEach(issue => {
  console.log(`\n[${issue.severity}] ${issue.description}`);
  console.log(`  Affected Devices: ${issue.affectedDeviceCount}`);
  console.log(`  Recommendation: ${issue.recommendation}`);
});
```

### Example 4: Comprehensive Security Dashboard

```typescript
async function generateSecurityDashboard() {
  const securityReports = new DeviceSecurityReports(graphClient);

  console.log('Generating comprehensive security dashboard...\n');

  // Generate all reports in parallel
  const [jailbreakReport, certReport, healthReport] = await Promise.all([
    securityReports.generateJailbreakReport(),
    securityReports.generateCertificateRenewalReport(),
    securityReports.generateHealthAttestationReport()
  ]);

  const dashboard = {
    generatedAt: new Date().toISOString(),

    jailbreakSecurity: {
      totalJailbroken: jailbreakReport.summary.jailbrokenDevices,
      criticalRisk: jailbreakReport.summary.riskAssessment.criticalRiskDevices,
      highRisk: jailbreakReport.summary.riskAssessment.highRiskDevices,
      overallRiskScore: jailbreakReport.summary.riskAssessment.overallRiskScore,
      iosJailbroken: jailbreakReport.summary.byPlatform.iOS.jailbroken,
      androidRooted: jailbreakReport.summary.byPlatform.Android.rooted,
      recommendations: jailbreakReport.summary.riskAssessment.recommendations
    },

    certificateSecurity: {
      devicesWithIssues: certReport.summary.devicesWithIssues,
      expiredCertificates: certReport.summary.expiredCertificates,
      expiringSoon: certReport.summary.expiringSoonCertificates,
      failedRenewals: certReport.summary.failedRenewals,
      criticalIssues: certReport.summary.bySeverity.critical
    },

    healthAttestation: {
      totalDevices: healthReport.summary.totalDevices,
      averageHealthScore: healthReport.summary.averageHealthScore,
      bitLockerPercentage: healthReport.summary.bitLockerPercentage,
      secureBootPercentage: healthReport.summary.secureBootPercentage,
      criticalHealthDevices: healthReport.summary.criticalHealthDevices,
      testSigningEnabled: healthReport.summary.testSigningEnabled
    },

    alerts: []
  };

  // Generate security alerts
  if (dashboard.jailbreakSecurity.criticalRisk > 0) {
    dashboard.alerts.push({
      severity: 'Critical',
      category: 'Jailbreak/Root',
      message: `${dashboard.jailbreakSecurity.criticalRisk} devices with critical jailbreak risk detected`,
      action: 'Review and block access immediately'
    });
  }

  if (dashboard.certificateSecurity.expiredCertificates > 0) {
    dashboard.alerts.push({
      severity: 'Critical',
      category: 'Certificates',
      message: `${dashboard.certificateSecurity.expiredCertificates} expired certificates detected`,
      action: 'Renew certificates immediately to prevent connectivity loss'
    });
  }

  if (dashboard.healthAttestation.criticalHealthDevices > 0) {
    dashboard.alerts.push({
      severity: 'High',
      category: 'Health Attestation',
      message: `${dashboard.healthAttestation.criticalHealthDevices} devices with critical health issues`,
      action: 'Review health attestation failures and remediate'
    });
  }

  if (healthReport.summary.testSigningEnabled > 0) {
    dashboard.alerts.push({
      severity: 'Critical',
      category: 'Test Signing',
      message: `${healthReport.summary.testSigningEnabled} devices have test signing enabled`,
      action: 'URGENT: Disable test signing immediately - major security vulnerability'
    });
  }

  // Export dashboard
  await fs.writeFile(
    './reports/security-dashboard.json',
    JSON.stringify(dashboard, null, 2)
  );

  console.log('\n=== SECURITY DASHBOARD ===');
  console.log(JSON.stringify(dashboard, null, 2));

  return dashboard;
}

// Execute
await generateSecurityDashboard();
```

### Example 5: Platform-Specific Analysis

```typescript
// iOS Jailbreak Analysis
const iosDevices = await securityReports.getJailbrokenOrRootedDevices({
  operatingSystem: ['iOS']
});

console.log(`\n=== iOS JAILBREAK ANALYSIS ===`);
console.log(`Total iOS Jailbroken: ${iosDevices.length}`);

const iosByVersion = iosDevices.reduce((acc, device) => {
  acc[device.osVersion] = (acc[device.osVersion] || 0) + 1;
  return acc;
}, {} as Record<string, number>);

console.log('\nBy iOS Version:');
Object.entries(iosByVersion).forEach(([version, count]) => {
  console.log(`  iOS ${version}: ${count}`);
});

// Android Root Analysis
const androidDevices = await securityReports.getJailbrokenOrRootedDevices({
  operatingSystem: ['Android']
});

console.log(`\n=== ANDROID ROOT ANALYSIS ===`);
console.log(`Total Android Rooted: ${androidDevices.length}`);

const androidByManufacturer = androidDevices.reduce((acc, device) => {
  acc[device.manufacturer] = (acc[device.manufacturer] || 0) + 1;
  return acc;
}, {} as Record<string, number>);

console.log('\nBy Manufacturer:');
Object.entries(androidByManufacturer).forEach(([mfr, count]) => {
  console.log(`  ${mfr}: ${count}`);
});
```

### Example 6: Scheduled Security Monitoring

```typescript
import * as nodemailer from 'nodemailer';

async function scheduledSecurityReport() {
  const securityReports = new DeviceSecurityReports(graphClient);
  const timestamp = new Date().toISOString().split('T')[0];
  const reportDir = `./reports/${timestamp}`;

  await fs.mkdir(reportDir, { recursive: true });

  // Generate all reports
  const [jailbreakReport, certReport, healthReport] = await Promise.all([
    securityReports.generateJailbreakReport(),
    securityReports.generateCertificateRenewalReport(),
    securityReports.generateHealthAttestationReport()
  ]);

  // Save reports
  await Promise.all([
    fs.writeFile(
      `${reportDir}/jailbreak-report.json`,
      JSON.stringify(jailbreakReport, null, 2)
    ),
    fs.writeFile(
      `${reportDir}/certificate-report.json`,
      JSON.stringify(certReport, null, 2)
    ),
    fs.writeFile(
      `${reportDir}/health-attestation-report.json`,
      JSON.stringify(healthReport, null, 2)
    )
  ]);

  // Check for critical issues
  const criticalIssues = [];

  if (jailbreakReport.summary.riskAssessment.criticalRiskDevices > 0) {
    criticalIssues.push(
      `${jailbreakReport.summary.riskAssessment.criticalRiskDevices} devices with critical jailbreak risk`
    );
  }

  if (certReport.summary.expiredCertificates > 0) {
    criticalIssues.push(
      `${certReport.summary.expiredCertificates} expired certificates`
    );
  }

  if (healthReport.summary.testSigningEnabled > 0) {
    criticalIssues.push(
      `${healthReport.summary.testSigningEnabled} devices with test signing enabled (CRITICAL)`
    );
  }

  // Send email alert if critical issues found
  if (criticalIssues.length > 0) {
    const transporter = nodemailer.createTransport({
      host: 'smtp.office365.com',
      port: 587,
      secure: false,
      auth: {
        user: process.env.EMAIL_USER,
        pass: process.env.EMAIL_PASS
      }
    });

    await transporter.sendMail({
      from: process.env.EMAIL_USER,
      to: 'security-team@contoso.com',
      subject: `[CRITICAL] Security Issues Detected - ${timestamp}`,
      html: `
        <h2>Critical Security Issues Detected</h2>
        <p>The following critical security issues were found in today's scan:</p>
        <ul>
          ${criticalIssues.map(issue => `<li style="color: red;"><strong>${issue}</strong></li>`).join('')}
        </ul>
        <p>Please review the attached reports for detailed information and take immediate action.</p>
        <p><strong>Generated:</strong> ${new Date().toLocaleString()}</p>
      `,
      attachments: [
        { path: `${reportDir}/jailbreak-report.json` },
        { path: `${reportDir}/certificate-report.json` },
        { path: `${reportDir}/health-attestation-report.json` }
      ]
    });

    console.log('✅ Critical security alert sent to security team');
  } else {
    console.log('✅ No critical security issues found');
  }

  console.log(`✅ Security reports saved to ${reportDir}`);
}

// Run daily at 8 AM
setInterval(scheduledSecurityReport, 24 * 60 * 60 * 1000);
```

## Microsoft Graph API Endpoints

### Jailbreak Detection

**Primary Endpoint**:
```
GET /deviceManagement/managedDevices?$filter=jailBroken eq true
```

**Query Parameters**:
- `$filter`: Filter for jailbroken devices
- `$select`: Select specific properties
- `$top`: Limit results per page (max 999)

**Response Properties Used**:
- `id`, `deviceName`, `userPrincipalName`, `emailAddress`
- `operatingSystem`, `osVersion`, `model`, `manufacturer`
- `jailBroken`, `complianceState`, `managementState`
- `lastSyncDateTime`, `enrolledDateTime`

### Certificate Monitoring

**Primary Endpoints**:
```
GET /deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'
GET /deviceManagement/managedDevices/{id}/deviceComplianceDeviceStatuses
GET /deviceManagement/managedDevices/{id}/deviceConfigurationStates
```

**Certificate-Related Compliance Settings**:
- `certificate.required`
- `certificate.expiration`
- `scep.renewal.status`
- `pkcs.renewal.status`

### Health Attestation

**Primary Endpoint**:
```
GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'Windows'
```

**Health Attestation Properties** (from `deviceHealthAttestationState`):
- `healthAttestationSupportedStatus`
- `bitLockerStatus`
- `secureBoot`
- `codeIntegrity`
- `bootDebugging`
- `operatingSystemKernelDebugging`
- `testSigning`
- `earlyLaunchAntiMalwareDriverProtection`
- `virtualSecureMode`
- `tpmVersion`
- `pcrHashAlgorithm`
- `bootManagerVersion`
- `codeIntegrityCheckVersion`

## Data Models

### JailbrokenDevice

```typescript
interface JailbrokenDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userEmail: string;
  operatingSystem: string;
  osVersion: string;
  model: string;
  manufacturer: string;
  isJailBroken: boolean;
  riskLevel: 'Critical' | 'High' | 'Medium' | 'Low';
  riskScore: number; // 0-100
  complianceState: string;
  lastSyncDateTime: Date;
  enrolledDateTime: Date;
  managementState: string;
  securityNotes: string[];
}
```

### CertificateIssue

```typescript
interface CertificateIssue {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userEmail: string;
  operatingSystem: string;
  certificateType: 'Device' | 'User' | 'WiFi' | 'VPN' | 'SCEP' | 'PKCS' | 'Other';
  issuer: string;
  subject: string;
  expirationDate: Date;
  daysUntilExpiration: number;
  status: 'Expired' | 'Expiring' | 'Renewal Failed' | 'Invalid' | 'Revoked' | 'Valid';
  severity: 'Critical' | 'High' | 'Medium' | 'Low';
  issueDescription: string;
  remediationAction: string;
  lastRenewalAttempt?: Date;
  renewalError?: string;
  thumbprint?: string;
  complianceState: string;
}
```

### HealthAttestationState

```typescript
interface HealthAttestationState {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userEmail: string;
  operatingSystem: string;
  osVersion: string;
  model: string;
  manufacturer: string;
  lastSyncDateTime: Date;
  complianceState: string;

  // Health Attestation Properties
  healthAttestationSupported: boolean;
  bitLockerStatus: 'Enabled' | 'Disabled' | 'Unknown';
  secureBootEnabled: boolean;
  codeIntegrityEnabled: boolean;
  bootDebuggingEnabled: boolean;
  kernelDebuggingEnabled: boolean;
  testSigningEnabled: boolean;
  elamEnabled: boolean;
  virtualSecureModeEnabled: boolean;
  tpmVersion?: string;
  pcrHashAlgorithm?: string;
  bootManagerVersion?: string;
  codeIntegrityCheckVersion?: string;

  // Health Score
  healthScore: number; // 0-100
  healthScoreRange: 'Excellent' | 'Good' | 'Fair' | 'Poor' | 'Critical';

  // Security Analysis
  securityIssues: string[];
  recommendations: string[];
}
```

## Security Risk Analysis

### Jailbreak Risk Scoring

The module uses a comprehensive risk scoring algorithm:

```typescript
function calculateJailbreakRiskScore(device: any): number {
  let score = 40; // Base score for being jailbroken

  // Compliance impact
  if (device.complianceState === 'noncompliant') {
    score += 20;
  }

  // Platform-specific risk
  if (device.operatingSystem?.toLowerCase().includes('ios')) {
    score += 15; // iOS jailbreak generally higher risk
  } else if (device.operatingSystem?.toLowerCase().includes('android')) {
    score += 10; // Android root
  }

  // Device age (older jailbroken devices = higher risk)
  const enrolledDate = new Date(device.enrolledDateTime);
  const daysSinceEnrollment = Math.floor(
    (Date.now() - enrolledDate.getTime()) / (1000 * 60 * 60 * 24)
  );

  if (daysSinceEnrollment > 365) {
    score += 15; // Old jailbroken device
  } else if (daysSinceEnrollment > 180) {
    score += 10;
  } else if (daysSinceEnrollment > 90) {
    score += 5;
  }

  // Sync status (not syncing = potential management bypass)
  const lastSync = new Date(device.lastSyncDateTime);
  const daysSinceSync = Math.floor(
    (Date.now() - lastSync.getTime()) / (1000 * 60 * 60 * 24)
  );

  if (daysSinceSync > 7) {
    score += 10; // Not syncing regularly
  } else if (daysSinceSync > 3) {
    score += 5;
  }

  return Math.min(score, 100);
}
```

### Health Score Calculation

```typescript
function calculateHealthScore(attestation: any): number {
  let score = 0;

  if (attestation.bitLockerStatus === 'enabled') score += 15;
  if (attestation.secureBoot === 'enabled') score += 15;
  if (attestation.codeIntegrity === 'enabled') score += 10;
  if (attestation.earlyLaunchAntiMalwareDriverProtection === 'enabled') score += 10;
  if (attestation.virtualSecureMode === 'enabled') score += 15;
  if (attestation.tpmVersion === '2.0') score += 10;
  if (attestation.bootDebugging === 'disabled') score += 10;
  if (attestation.operatingSystemKernelDebugging === 'disabled') score += 5;
  if (attestation.testSigning === 'disabled') score += 10;

  return score;
}
```

## Remediation Guidance

### Jailbroken/Rooted Devices

**Critical Risk Devices (Score 80-100)**:
1. Immediately block device access to corporate resources
2. Initiate remote wipe if device contains sensitive data
3. Require user to restore device to factory settings
4. Re-enroll device after verification
5. Educate user on jailbreak/root security risks

**High Risk Devices (Score 60-79)**:
1. Restrict access to sensitive apps and data
2. Increase monitoring frequency
3. Schedule meeting with user to discuss security concerns
4. Consider conditional access policies
5. Plan for device replacement or wipe

**Medium/Low Risk Devices**:
1. Monitor closely
2. Apply conditional access policies
3. User education
4. Consider compliance policy enforcement

### Certificate Issues

**Expired Certificates (Critical)**:
1. Renew certificate immediately
2. Verify certificate authority connectivity
3. Redeploy certificate profile
4. Check for SCEP/PKCS configuration errors
5. Manual certificate installation if automated renewal fails

**Expiring Soon (High)**:
1. Initiate renewal process
2. Monitor renewal status
3. Prepare for manual intervention if needed
4. Verify auto-renewal settings

**Renewal Failures (High)**:
1. Check network connectivity
2. Verify CA availability and accessibility
3. Review SCEP/PKCS profile configuration
4. Check device enrollment status
5. Review certificate templates and permissions

### Health Attestation Issues

**Test Signing Enabled (Critical)**:
```powershell
# Disable test signing via PowerShell
bcdedit /set testsigning off

# Or via Intune remediation script
if ((Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Config' -Name VulnerableDriverBlocklistEnable).VulnerableDriverBlocklistEnable -ne 1) {
  Set-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\CI\Config' -Name VulnerableDriverBlocklistEnable -Value 1
  Restart-Computer
}
```

**BitLocker Disabled (High)**:
1. Deploy Intune BitLocker encryption policy
2. Verify TPM is available and functional
3. Check for hardware compatibility
4. Monitor encryption progress
5. Verify recovery key backup to Azure AD

**Secure Boot Disabled (High)**:
1. Enable Secure Boot in UEFI/BIOS settings
2. Update firmware if needed
3. Verify hardware compatibility
4. Deploy UEFI configuration via Intune
5. Remote assistance for BIOS access if needed

**Debugging Enabled (High)**:
```powershell
# Disable kernel debugging
bcdedit /debug off
bcdedit /set bootdebug off
```

## Export Formats

### JSON Export

Full detailed data with all properties:

```typescript
await fs.writeFile(
  './reports/jailbreak-report.json',
  JSON.stringify(jailbreakReport, null, 2)
);
```

### CSV Export

Flattened data suitable for Excel analysis:

```typescript
import { Parser } from 'json2csv';

const fields = [
  'deviceName',
  'userPrincipalName',
  'operatingSystem',
  'riskLevel',
  'riskScore',
  'complianceState'
];

const parser = new Parser({ fields });
const csv = parser.parse(jailbreakReport.devices);

await fs.writeFile('./reports/jailbreak-report.csv', csv);
```

### HTML Export

Formatted HTML with visual indicators:

```typescript
function generateHTML(report: any): string {
  return `
<!DOCTYPE html>
<html>
<head>
  <title>Security Report</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    .critical { background-color: #dc3545; color: white; }
    .high { background-color: #fd7e14; color: white; }
    .medium { background-color: #ffc107; }
    .low { background-color: #28a745; color: white; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #007bff; color: white; }
  </style>
</head>
<body>
  <h1>Jailbreak Detection Report</h1>
  <p>Generated: ${new Date().toLocaleString()}</p>

  <h2>Summary</h2>
  <p>Total Jailbroken Devices: ${report.summary.jailbrokenDevices}</p>
  <p>Overall Risk Score: ${report.summary.riskAssessment.overallRiskScore}/100</p>

  <table>
    <thead>
      <tr>
        <th>Device</th>
        <th>User</th>
        <th>OS</th>
        <th>Risk Level</th>
        <th>Score</th>
        <th>Compliance</th>
      </tr>
    </thead>
    <tbody>
      ${report.devices.map(d => `
        <tr class="${d.riskLevel.toLowerCase()}">
          <td>${d.deviceName}</td>
          <td>${d.userPrincipalName}</td>
          <td>${d.operatingSystem} ${d.osVersion}</td>
          <td>${d.riskLevel}</td>
          <td>${d.riskScore}</td>
          <td>${d.complianceState}</td>
        </tr>
      `).join('')}
    </tbody>
  </table>
</body>
</html>
  `;
}

await fs.writeFile('./reports/report.html', generateHTML(jailbreakReport));
```

## Best Practices

### 1. Schedule Regular Reports

Run security reports on a regular schedule:

```typescript
// Daily security scan at 8 AM
const schedule = require('node-schedule');

schedule.scheduleJob('0 8 * * *', async () => {
  await generateSecurityDashboard();
});
```

### 2. Implement Alerting

Set up automated alerts for critical issues:

```typescript
if (report.summary.riskAssessment.criticalRiskDevices > 0) {
  await sendAlertEmail({
    subject: '[CRITICAL] Jailbroken Devices Detected',
    body: `Found ${report.summary.riskAssessment.criticalRiskDevices} critical risk devices`
  });
}
```

### 3. Track Trends Over Time

Store historical data for trend analysis:

```typescript
const historicalData = {
  date: new Date().toISOString(),
  jailbrokenCount: report.summary.jailbrokenDevices,
  riskScore: report.summary.riskAssessment.overallRiskScore
};

await db.insert('security_history', historicalData);
```

### 4. Integrate with SIEM

Export data to your Security Information and Event Management system:

```typescript
// Send to Splunk, Azure Sentinel, etc.
await sendToSIEM({
  eventType: 'intune.security.jailbreak',
  severity: 'critical',
  data: report
});
```

### 5. Use Conditional Access

Combine with Azure AD Conditional Access:

```typescript
// Block access for jailbroken devices
if (device.riskLevel === 'Critical') {
  await blockDeviceAccess(device.deviceId);
  await sendUserNotification(device.userPrincipalName);
}
```

### 6. Implement Remediation Workflows

Automate remediation actions:

```typescript
for (const device of criticalDevices) {
  // Create ServiceNow ticket
  await createTicket({
    title: `Critical Jailbreak: ${device.deviceName}`,
    description: `Device ${device.deviceName} is jailbroken with risk score ${device.riskScore}`,
    priority: 'P1',
    assignee: 'security-team'
  });

  // Block device
  await blockDevice(device.deviceId);

  // Notify user
  await sendEmail({
    to: device.userEmail,
    subject: 'Urgent: Your device has been blocked',
    body: `Your device ${device.deviceName} has been detected as jailbroken...`
  });
}
```

## Troubleshooting

### Issue: No jailbroken devices returned but you know they exist

**Possible Causes**:
1. Device hasn't synced recently
2. Graph API permissions insufficient
3. Jailbreak detection not triggered yet

**Solutions**:
```typescript
// Check all devices, not just jailbroken
const allDevices = await graphClient
  .api('/deviceManagement/managedDevices')
  .get();

console.log(`Total devices: ${allDevices.value.length}`);

// Check permissions
const permissions = await graphClient
  .api('/me/oauth2PermissionGrants')
  .get();

console.log('Granted permissions:', permissions);

// Force device sync
await graphClient
  .api(`/deviceManagement/managedDevices/${deviceId}/syncDevice`)
  .post({});
```

### Issue: Certificate issues not detected

**Possible Causes**:
1. Certificate compliance policies not deployed
2. Device not enrolled in certificate-based authentication
3. Certificate state not reported to Intune

**Solutions**:
```typescript
// Check device compliance policies
const policies = await graphClient
  .api('/deviceManagement/deviceCompliancePolicies')
  .get();

// Check device configuration profiles
const profiles = await graphClient
  .api('/deviceManagement/deviceConfigurations')
  .filter("contains(displayName, 'certificate')")
  .get();
```

### Issue: Health attestation not available

**Possible Causes**:
1. Device doesn't support health attestation (requires Windows 10/11 with TPM)
2. Health attestation service not enabled
3. Device not yet reporting health state

**Solutions**:
```typescript
// Check Windows devices specifically
const windowsDevices = await graphClient
  .api('/deviceManagement/managedDevices')
  .filter("operatingSystem eq 'Windows'")
  .select('id,deviceName,operatingSystem,osVersion')
  .get();

// Check each device for health attestation support
for (const device of windowsDevices.value) {
  const healthState = await graphClient
    .api(`/deviceManagement/managedDevices/${device.id}`)
    .select('deviceHealthAttestationState')
    .get();

  console.log(`${device.deviceName}: ${healthState.deviceHealthAttestationState?.healthAttestationSupportedStatus || 'Not supported'}`);
}
```

### Issue: Rate limiting errors

**Solution**:
The module includes automatic retry with exponential backoff, but you can adjust:

```typescript
private async executeWithRetry<T>(
  operation: () => Promise<T>,
  maxRetries: number = 5, // Increase retries
  baseDelay: number = 2000  // Increase delay
): Promise<T> {
  // ... implementation with retry logic
}
```

## Performance Considerations

### 1. Pagination

The module automatically handles pagination for large datasets:

```typescript
// Fetches all pages automatically
const allDevices = await securityReports.getJailbrokenOrRootedDevices();
```

### 2. Parallel Requests

Generate multiple reports in parallel:

```typescript
const [jailbreakReport, certReport, healthReport] = await Promise.all([
  securityReports.generateJailbreakReport(),
  securityReports.generateCertificateRenewalReport(),
  securityReports.generateHealthAttestationReport()
]);
```

### 3. Selective Properties

Only request needed properties to reduce data transfer:

```typescript
const response = await graphClient
  .api('/deviceManagement/managedDevices')
  .select('id,deviceName,jailBroken,complianceState')
  .get();
```

### 4. Caching

Cache results for frequently accessed data:

```typescript
const cache = new Map<string, any>();
const CACHE_TTL = 5 * 60 * 1000; // 5 minutes

async function getCachedDevices(): Promise<any[]> {
  const cached = cache.get('devices');
  if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
    return cached.data;
  }

  const data = await securityReports.getJailbrokenOrRootedDevices();
  cache.set('devices', { data, timestamp: Date.now() });
  return data;
}
```

### 5. Batch Processing

For large environments (10,000+ devices), process in batches:

```typescript
async function processLargeEnvironment() {
  const batchSize = 1000;
  let skip = 0;
  let hasMore = true;

  while (hasMore) {
    const response = await graphClient
      .api('/deviceManagement/managedDevices')
      .filter('jailBroken eq true')
      .top(batchSize)
      .skip(skip)
      .get();

    // Process batch
    await processBatch(response.value);

    skip += batchSize;
    hasMore = !!response['@odata.nextLink'];
  }
}
```

---

## Additional Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/overview)
- [Intune Device Management](https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [Device Health Attestation](https://docs.microsoft.com/en-us/windows/security/threat-protection/protect-high-value-assets-by-controlling-the-health-of-windows-10-based-devices)
- [Azure AD Conditional Access](https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/overview)

## Support

For issues, questions, or contributions, please contact the Intune Reporting System team or create an issue in the repository.

---

**Version**: 1.0.0
**Last Updated**: February 2026
**Author**: Intune Reporting System Team
