/**
 * Device Security Reports Usage Examples
 *
 * This file demonstrates how to use the Device Security Reports module
 * to replicate Configuration Manager security reports in Intune.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import {
  DeviceSecurityReports,
  JailbrokenDevice,
  CertificateIssue,
  HealthAttestationState,
  ReportFilters,
} from '../device-security-reports';
import * as fs from 'fs/promises';

/**
 * Example 1: Basic Security Report Generation
 *
 * Demonstrates how to generate all three security reports using the
 * Configuration Manager-compatible method names.
 */
async function example1_BasicSecurityReports() {
  console.log('\n=== Example 1: Basic Security Reports ===\n');

  // Setup authentication
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

  // Report #23: Jailbroken or Rooted Devices
  console.log('Generating Report #23: Jailbroken or Rooted Devices...');
  const jailbrokenDevices = await securityReports.getJailbrokenOrRootedDevices();
  console.log(`Found ${jailbrokenDevices.length} jailbroken or rooted devices`);

  // Report #27: Certificate Renewal Issues
  console.log('\nGenerating Report #27: Certificate Renewal Issues...');
  const certificateIssues = await securityReports.getDevicesWithCertificateIssues();
  console.log(`Found ${certificateIssues.length} devices with certificate issues`);

  // Report #16: Health Attestation State
  console.log('\nGenerating Report #16: Health Attestation State...');
  const healthDevices = await securityReports.getDevicesByHealthAttestationState();
  console.log(`Found ${healthDevices.length} devices with health attestation data`);
}

/**
 * Example 2: Jailbreak Detection with Risk Analysis
 *
 * Report #23: Mobile devices that are jailbroken or rooted
 */
async function example2_JailbreakDetection(graphClient: Client) {
  console.log('\n=== Example 2: Jailbreak Detection with Risk Analysis ===\n');

  const securityReports = new DeviceSecurityReports(graphClient);

  // Get all jailbroken devices
  const jailbrokenDevices = await securityReports.getJailbrokenOrRootedDevices();

  console.log(`Total Jailbroken/Rooted Devices: ${jailbrokenDevices.length}\n`);

  // Group by risk level
  const riskGroups = {
    Critical: jailbrokenDevices.filter(d => d.riskLevel === 'Critical'),
    High: jailbrokenDevices.filter(d => d.riskLevel === 'High'),
    Medium: jailbrokenDevices.filter(d => d.riskLevel === 'Medium'),
    Low: jailbrokenDevices.filter(d => d.riskLevel === 'Low'),
  };

  console.log('Risk Level Distribution:');
  console.log(`  Critical: ${riskGroups.Critical.length}`);
  console.log(`  High: ${riskGroups.High.length}`);
  console.log(`  Medium: ${riskGroups.Medium.length}`);
  console.log(`  Low: ${riskGroups.Low.length}`);

  // Display critical risk devices
  if (riskGroups.Critical.length > 0) {
    console.log('\n⚠️ CRITICAL RISK DEVICES - IMMEDIATE ACTION REQUIRED:');
    riskGroups.Critical.forEach(device => {
      console.log(`\n  Device: ${device.deviceName}`);
      console.log(`  User: ${device.userPrincipalName}`);
      console.log(`  Platform: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`  Risk Score: ${device.riskScore}/100`);
      console.log(`  Compliance: ${device.complianceState}`);
      console.log(`  Security Notes:`);
      device.securityNotes.forEach(note => console.log(`    - ${note}`));
    });
  }

  // Platform breakdown
  const iosDevices = jailbrokenDevices.filter(d =>
    d.operatingSystem.toLowerCase().includes('ios')
  );
  const androidDevices = jailbrokenDevices.filter(d =>
    d.operatingSystem.toLowerCase().includes('android')
  );

  console.log('\nPlatform Distribution:');
  console.log(`  iOS Jailbroken: ${iosDevices.length}`);
  console.log(`  Android Rooted: ${androidDevices.length}`);
}

/**
 * Example 3: Certificate Renewal Monitoring
 *
 * Report #27: Mobile devices with certificate renewal issues
 */
async function example3_CertificateMonitoring(graphClient: Client) {
  console.log('\n=== Example 3: Certificate Renewal Monitoring ===\n');

  const securityReports = new DeviceSecurityReports(graphClient);

  // Get all certificate issues
  const issues = await securityReports.getDevicesWithCertificateIssues();

  console.log(`Total Devices with Certificate Issues: ${issues.length}\n`);

  // Group by severity
  const severityGroups = {
    Critical: issues.filter(i => i.severity === 'Critical'),
    High: issues.filter(i => i.severity === 'High'),
    Medium: issues.filter(i => i.severity === 'Medium'),
    Low: issues.filter(i => i.severity === 'Low'),
  };

  console.log('Severity Distribution:');
  console.log(`  Critical: ${severityGroups.Critical.length}`);
  console.log(`  High: ${severityGroups.High.length}`);
  console.log(`  Medium: ${severityGroups.Medium.length}`);
  console.log(`  Low: ${severityGroups.Low.length}`);

  // Group by certificate type
  const typeGroups = issues.reduce((acc, issue) => {
    acc[issue.certificateType] = (acc[issue.certificateType] || 0) + 1;
    return acc;
  }, {} as Record<string, number>);

  console.log('\nCertificate Type Distribution:');
  Object.entries(typeGroups).forEach(([type, count]) => {
    console.log(`  ${type}: ${count}`);
  });

  // Display critical issues
  if (severityGroups.Critical.length > 0) {
    console.log('\n🔴 CRITICAL CERTIFICATE ISSUES:');
    severityGroups.Critical.forEach(issue => {
      console.log(`\n  Device: ${issue.deviceName}`);
      console.log(`  User: ${issue.userPrincipalName}`);
      console.log(`  Certificate Type: ${issue.certificateType}`);
      console.log(`  Status: ${issue.status}`);
      console.log(`  Expiration: ${issue.expirationDate.toLocaleDateString()}`);
      console.log(`  Days Until Expiration: ${issue.daysUntilExpiration}`);
      console.log(`  Issue: ${issue.issueDescription}`);
      console.log(`  Remediation: ${issue.remediationAction}`);
    });
  }

  // Expiration timeline
  const expired = issues.filter(i => i.status === 'Expired');
  const expiring7Days = issues.filter(
    i => i.daysUntilExpiration >= 0 && i.daysUntilExpiration <= 7
  );
  const expiring30Days = issues.filter(
    i => i.daysUntilExpiration > 7 && i.daysUntilExpiration <= 30
  );

  console.log('\nExpiration Timeline:');
  console.log(`  Already Expired: ${expired.length}`);
  console.log(`  Expiring within 7 days: ${expiring7Days.length}`);
  console.log(`  Expiring within 30 days: ${expiring30Days.length}`);
}

/**
 * Example 4: Windows Health Attestation Analysis
 *
 * Report #16: List of devices by Health Attestation state
 */
async function example4_HealthAttestation(graphClient: Client) {
  console.log('\n=== Example 4: Windows Health Attestation Analysis ===\n');

  const securityReports = new DeviceSecurityReports(graphClient);

  // Get health attestation report
  const report = await securityReports.generateHealthAttestationReport();

  console.log(`Total Windows Devices: ${report.summary.totalDevices}`);
  console.log(`Health Attestation Supported: ${report.summary.supportedDevices}\n`);

  // BitLocker statistics
  console.log('BitLocker Status:');
  console.log(
    `  Enabled: ${report.summary.bitLockerEnabled} (${report.summary.bitLockerPercentage}%)`
  );
  console.log(`  Disabled: ${report.summary.bitLockerDisabled}`);

  // Secure Boot statistics
  console.log('\nSecure Boot Status:');
  console.log(
    `  Enabled: ${report.summary.secureBootEnabled} (${report.summary.secureBootPercentage}%)`
  );
  console.log(`  Disabled: ${report.summary.secureBootDisabled}`);

  // Code Integrity statistics
  console.log('\nCode Integrity Status:');
  console.log(
    `  Enabled: ${report.summary.codeIntegrityEnabled} (${report.summary.codeIntegrityPercentage}%)`
  );
  console.log(`  Disabled: ${report.summary.codeIntegrityDisabled}`);

  // Health score distribution
  console.log('\nHealth Score Distribution:');
  console.log(`  Excellent (90-100): ${report.summary.excellentHealthDevices}`);
  console.log(`  Good (70-89): ${report.summary.goodHealthDevices}`);
  console.log(`  Fair (50-69): ${report.summary.fairHealthDevices}`);
  console.log(`  Poor (30-49): ${report.summary.poorHealthDevices}`);
  console.log(`  Critical (0-29): ${report.summary.criticalHealthDevices}`);
  console.log(`  Average: ${report.summary.averageHealthScore}/100`);

  // Critical security issues
  if (report.summary.criticalIssues.length > 0) {
    console.log('\n🚨 CRITICAL SECURITY ISSUES:');
    report.summary.criticalIssues.forEach(issue => {
      console.log(`\n  [${issue.severity}] ${issue.description}`);
      console.log(`  Affected Devices: ${issue.affectedDeviceCount}`);
      console.log(`  Recommendation: ${issue.recommendation}`);
    });
  }

  // TPM version distribution
  console.log('\nTPM Version Distribution:');
  Object.entries(report.summary.tpmVersions).forEach(([version, count]) => {
    console.log(`  TPM ${version}: ${count}`);
  });

  // Identify devices needing attention
  const poorHealthDevices = report.devices.filter(
    d => d.healthScoreRange === 'Poor' || d.healthScoreRange === 'Critical'
  );

  if (poorHealthDevices.length > 0) {
    console.log(`\n⚠️ Devices Requiring Immediate Attention: ${poorHealthDevices.length}`);
    console.log('\nTop 5 Devices with Lowest Health Scores:');

    const sorted = [...poorHealthDevices].sort((a, b) => a.healthScore - b.healthScore);
    sorted.slice(0, 5).forEach((device, index) => {
      console.log(`\n  ${index + 1}. ${device.deviceName}`);
      console.log(`     User: ${device.userPrincipalName}`);
      console.log(`     Health Score: ${device.healthScore}/100 (${device.healthScoreRange})`);
      console.log(`     BitLocker: ${device.bitLockerStatus}`);
      console.log(`     Secure Boot: ${device.secureBootEnabled ? 'Enabled' : 'Disabled'}`);
      console.log(`     Issues: ${device.securityIssues.length}`);
      if (device.securityIssues.length > 0) {
        console.log(`     Top Issues:`);
        device.securityIssues.slice(0, 3).forEach(issue => {
          console.log(`       - ${issue}`);
        });
      }
    });
  }
}

/**
 * Example 5: Comprehensive Security Dashboard
 *
 * Combines all three reports into a single dashboard
 */
async function example5_SecurityDashboard(graphClient: Client) {
  console.log('\n=== Example 5: Comprehensive Security Dashboard ===\n');

  const securityReports = new DeviceSecurityReports(graphClient);

  console.log('Generating comprehensive security dashboard...\n');

  // Generate all reports in parallel
  const [jailbreakReport, certReport, healthReport] = await Promise.all([
    securityReports.generateJailbreakReport(),
    securityReports.generateCertificateRenewalReport(),
    securityReports.generateHealthAttestationReport(),
  ]);

  // Create dashboard object
  const dashboard = {
    generatedAt: new Date().toISOString(),

    jailbreakSecurity: {
      totalJailbroken: jailbreakReport.summary.jailbrokenDevices,
      criticalRisk: jailbreakReport.summary.riskAssessment.criticalRiskDevices,
      highRisk: jailbreakReport.summary.riskAssessment.highRiskDevices,
      overallRiskScore: jailbreakReport.summary.riskAssessment.overallRiskScore,
      iosJailbroken: jailbreakReport.summary.byPlatform.iOS.jailbroken,
      androidRooted: jailbreakReport.summary.byPlatform.Android.rooted,
    },

    certificateSecurity: {
      devicesWithIssues: certReport.summary.devicesWithIssues,
      expiredCertificates: certReport.summary.expiredCertificates,
      expiringSoon: certReport.summary.expiringSoonCertificates,
      failedRenewals: certReport.summary.failedRenewals,
      criticalIssues: certReport.summary.bySeverity.critical,
    },

    healthAttestation: {
      totalDevices: healthReport.summary.totalDevices,
      averageHealthScore: healthReport.summary.averageHealthScore,
      bitLockerPercentage: healthReport.summary.bitLockerPercentage,
      secureBootPercentage: healthReport.summary.secureBootPercentage,
      criticalHealthDevices: healthReport.summary.criticalHealthDevices,
      testSigningEnabled: healthReport.summary.testSigningEnabled,
    },

    alerts: [] as any[],
  };

  // Generate security alerts
  if (dashboard.jailbreakSecurity.criticalRisk > 0) {
    dashboard.alerts.push({
      severity: 'Critical',
      category: 'Jailbreak/Root',
      message: `${dashboard.jailbreakSecurity.criticalRisk} devices with critical jailbreak risk`,
      action: 'Review and block access immediately',
    });
  }

  if (dashboard.certificateSecurity.expiredCertificates > 0) {
    dashboard.alerts.push({
      severity: 'Critical',
      category: 'Certificates',
      message: `${dashboard.certificateSecurity.expiredCertificates} expired certificates detected`,
      action: 'Renew certificates immediately',
    });
  }

  if (dashboard.healthAttestation.criticalHealthDevices > 0) {
    dashboard.alerts.push({
      severity: 'High',
      category: 'Health Attestation',
      message: `${dashboard.healthAttestation.criticalHealthDevices} devices with critical health issues`,
      action: 'Review and remediate health failures',
    });
  }

  if (dashboard.healthAttestation.testSigningEnabled > 0) {
    dashboard.alerts.push({
      severity: 'Critical',
      category: 'Test Signing',
      message: `${dashboard.healthAttestation.testSigningEnabled} devices have test signing enabled`,
      action: 'URGENT: Disable test signing immediately',
    });
  }

  // Display dashboard
  console.log('=== SECURITY DASHBOARD ===\n');
  console.log(`Generated: ${new Date(dashboard.generatedAt).toLocaleString()}\n`);

  console.log('JAILBREAK SECURITY:');
  console.log(`  Total Jailbroken: ${dashboard.jailbreakSecurity.totalJailbroken}`);
  console.log(`  Critical Risk: ${dashboard.jailbreakSecurity.criticalRisk}`);
  console.log(`  High Risk: ${dashboard.jailbreakSecurity.highRisk}`);
  console.log(`  Overall Risk Score: ${dashboard.jailbreakSecurity.overallRiskScore}/100`);

  console.log('\nCERTIFICATE SECURITY:');
  console.log(`  Devices with Issues: ${dashboard.certificateSecurity.devicesWithIssues}`);
  console.log(`  Expired Certificates: ${dashboard.certificateSecurity.expiredCertificates}`);
  console.log(`  Expiring Soon: ${dashboard.certificateSecurity.expiringSoon}`);
  console.log(`  Failed Renewals: ${dashboard.certificateSecurity.failedRenewals}`);

  console.log('\nHEALTH ATTESTATION:');
  console.log(`  Total Devices: ${dashboard.healthAttestation.totalDevices}`);
  console.log(`  Average Health Score: ${dashboard.healthAttestation.averageHealthScore}/100`);
  console.log(`  BitLocker Enabled: ${dashboard.healthAttestation.bitLockerPercentage}%`);
  console.log(`  Secure Boot Enabled: ${dashboard.healthAttestation.secureBootPercentage}%`);
  console.log(`  Critical Health Devices: ${dashboard.healthAttestation.criticalHealthDevices}`);

  if (dashboard.alerts.length > 0) {
    console.log('\n🚨 SECURITY ALERTS:');
    dashboard.alerts.forEach(alert => {
      console.log(`\n  [${alert.severity}] ${alert.category}`);
      console.log(`  ${alert.message}`);
      console.log(`  Action: ${alert.action}`);
    });
  } else {
    console.log('\n✅ No critical security alerts');
  }

  // Export dashboard
  await fs.writeFile(
    './reports/security-dashboard.json',
    JSON.stringify(dashboard, null, 2)
  );

  console.log('\n✅ Dashboard exported to ./reports/security-dashboard.json');
}

/**
 * Example 6: Advanced Filtering
 *
 * Demonstrates filtering capabilities across all report types
 */
async function example6_AdvancedFiltering(graphClient: Client) {
  console.log('\n=== Example 6: Advanced Filtering ===\n');

  const securityReports = new DeviceSecurityReports(graphClient);

  // Filter 1: Critical jailbroken devices that are non-compliant
  console.log('Filter 1: Critical jailbroken devices (non-compliant)');
  const criticalJailbroken = await securityReports.getJailbrokenOrRootedDevices({
    riskLevel: ['Critical'],
    complianceState: ['noncompliant'],
  });
  console.log(`  Found: ${criticalJailbroken.length} devices\n`);

  // Filter 2: iOS-only jailbroken devices
  console.log('Filter 2: iOS jailbroken devices only');
  const iosJailbroken = await securityReports.getJailbrokenOrRootedDevices({
    operatingSystem: ['iOS'],
  });
  console.log(`  Found: ${iosJailbroken.length} devices\n`);

  // Filter 3: Critical certificate issues
  console.log('Filter 3: Critical certificate issues');
  const criticalCerts = await securityReports.getDevicesWithCertificateIssues({
    severity: ['Critical', 'High'],
  });
  console.log(`  Found: ${criticalCerts.length} issues\n`);

  // Filter 4: Windows devices with non-compliant health
  console.log('Filter 4: Non-compliant Windows devices');
  const nonCompliantHealth = await securityReports.getDevicesByHealthAttestationState({
    operatingSystem: ['Windows'],
    complianceState: ['noncompliant'],
  });
  console.log(`  Found: ${nonCompliantHealth.length} devices\n`);

  // Filter 5: Specific users
  console.log('Filter 5: Security issues for specific users');
  const userFilters: ReportFilters = {
    userPrincipalNames: ['user1@contoso.com', 'user2@contoso.com'],
  };

  const [userJailbreak, userCerts, userHealth] = await Promise.all([
    securityReports.getJailbrokenOrRootedDevices(userFilters),
    securityReports.getDevicesWithCertificateIssues(userFilters),
    securityReports.getDevicesByHealthAttestationState(userFilters),
  ]);

  console.log(`  Jailbroken devices: ${userJailbreak.length}`);
  console.log(`  Certificate issues: ${userCerts.length}`);
  console.log(`  Health attestation devices: ${userHealth.length}`);
}

/**
 * Main execution function
 */
async function main() {
  console.log('==========================================================');
  console.log('Device Security Reports - Configuration Manager Reports');
  console.log('==========================================================');

  try {
    // Setup authentication
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

    // Run examples
    await example1_BasicSecurityReports();
    // await example2_JailbreakDetection(graphClient);
    // await example3_CertificateMonitoring(graphClient);
    // await example4_HealthAttestation(graphClient);
    // await example5_SecurityDashboard(graphClient);
    // await example6_AdvancedFiltering(graphClient);

    console.log('\n==========================================================');
    console.log('Examples completed successfully!');
    console.log('==========================================================\n');
  } catch (error) {
    console.error('Error running examples:', error);
    process.exit(1);
  }
}

// Export functions for individual use
export {
  example1_BasicSecurityReports,
  example2_JailbreakDetection,
  example3_CertificateMonitoring,
  example4_HealthAttestation,
  example5_SecurityDashboard,
  example6_AdvancedFiltering,
};

// Run if executed directly
if (require.main === module) {
  main();
}
