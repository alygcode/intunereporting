/**
 * Device Compliance Report - Usage Examples
 *
 * This file demonstrates various ways to use the DeviceComplianceReport module
 * to generate comprehensive compliance reports from Microsoft Intune.
 *
 * @module device-compliance-report-example
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { ClientSecretCredential } from '@azure/identity';
import { DeviceComplianceReport, CompliancePlatform, ComplianceState } from '../device-compliance-report';
import { OutputFormatter } from '../../formatters/output-formatter';
import { AppConfig } from '../../types';
import * as path from 'path';
import * as fs from 'fs/promises';

// ============================================================================
// Example 1: Basic Compliance Report
// ============================================================================

/**
 * Generate a basic compliance report with all data
 */
export async function example1_BasicComplianceReport() {
  console.log('\n=== Example 1: Basic Compliance Report ===\n');

  try {
    // Initialize Graph client
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();

    // Create report instance
    const report = new DeviceComplianceReport(graphClient, config);

    // Execute the report
    console.log('Generating comprehensive compliance report...');
    const reportData = await report.execute();

    console.log('\nReport Summary:');
    console.log(`  Total Devices: ${reportData.summary.totalDevices}`);
    console.log(`  Compliant: ${reportData.summary.compliantDevices} (${reportData.summary.compliancePercentage}%)`);
    console.log(`  Non-Compliant: ${reportData.summary.nonCompliantDevices}`);
    console.log(`  Errors: ${reportData.summary.errorDevices}`);
    console.log(`  Total Policies: ${reportData.summary.totalPolicies}`);
    console.log(`  Assigned Policies: ${reportData.summary.assignedPolicies}`);

    // Display platform breakdown
    console.log('\nCompliance by Platform:');
    for (const [platform, stats] of Object.entries(reportData.summary.byPlatform)) {
      console.log(`  ${platform}:`);
      console.log(`    Total: ${stats.total}`);
      console.log(`    Compliant: ${stats.compliant} (${stats.compliancePercentage}%)`);
      console.log(`    Non-Compliant: ${stats.nonCompliant}`);
    }

    // Display critical issues
    if (reportData.summary.criticalIssues && reportData.summary.criticalIssues.length > 0) {
      console.log('\nCritical Issues:');
      for (const issue of reportData.summary.criticalIssues) {
        console.log(`  [${issue.severity.toUpperCase()}] ${issue.description}`);
        console.log(`    Recommendation: ${issue.recommendation}`);
      }
    }

    // Export to all formats
    const outputDir = './reports/compliance';
    await fs.mkdir(outputDir, { recursive: true });

    const jsonPath = await OutputFormatter.format(reportData, 'json', outputDir, true);
    console.log(`\nJSON report saved to: ${jsonPath}`);

    const csvPath = await OutputFormatter.format(reportData, 'csv', outputDir, true);
    console.log(`CSV report saved to: ${csvPath}`);

    const htmlPath = await OutputFormatter.format(reportData, 'html', outputDir, true);
    console.log(`HTML report saved to: ${htmlPath}`);

    return reportData;
  } catch (error) {
    console.error('Error generating compliance report:', error);
    throw error;
  }
}

// ============================================================================
// Example 2: Get All Compliance Policies
// ============================================================================

/**
 * Retrieve and display all device compliance policies
 */
export async function example2_GetAllCompliancePolicies() {
  console.log('\n=== Example 2: Get All Compliance Policies ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Fetching all compliance policies...');
    const policies = await report.getAllCompliancePolicies();

    console.log(`\nFound ${policies.length} compliance policies:\n`);

    for (const policy of policies) {
      console.log(`Policy: ${policy.displayName}`);
      console.log(`  ID: ${policy.id}`);
      console.log(`  Platform: ${policy.platform}`);
      console.log(`  Version: ${policy.version}`);
      console.log(`  Assignments: ${policy.assignmentCount}`);
      console.log(`  Settings: ${policy.settingsCount}`);
      console.log(`  Created: ${policy.createdDateTime.toLocaleDateString()}`);
      console.log(`  Modified: ${policy.lastModifiedDateTime.toLocaleDateString()}`);
      console.log(`  Is Assigned: ${policy.isAssigned ? 'Yes' : 'No'}`);

      if (policy.description) {
        console.log(`  Description: ${policy.description}`);
      }

      console.log('');
    }

    // Group by platform
    const byPlatform = policies.reduce((acc, policy) => {
      acc[policy.platform] = (acc[policy.platform] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Policies by Platform:');
    for (const [platform, count] of Object.entries(byPlatform)) {
      console.log(`  ${platform}: ${count}`);
    }

    // Identify unassigned policies
    const unassigned = policies.filter(p => !p.isAssigned);
    if (unassigned.length > 0) {
      console.log(`\n⚠️  Warning: ${unassigned.length} ${unassigned.length === 1 ? 'policy is' : 'policies are'} not assigned:`);
      for (const policy of unassigned) {
        console.log(`  - ${policy.displayName}`);
      }
    }

    return policies;
  } catch (error) {
    console.error('Error fetching compliance policies:', error);
    throw error;
  }
}

// ============================================================================
// Example 3: Get Policy Device State Summaries
// ============================================================================

/**
 * Get compliance statistics for each policy
 */
export async function example3_GetPolicyDeviceStateSummaries() {
  console.log('\n=== Example 3: Policy Device State Summaries ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Fetching policy device state summaries...');
    const summaries = await report.getAllPolicyDeviceStateSummaries();

    console.log(`\nFound summaries for ${summaries.length} policies:\n`);

    // Sort by compliance percentage (lowest first to highlight issues)
    const sortedSummaries = summaries.sort((a, b) => a.compliancePercentage - b.compliancePercentage);

    for (const summary of sortedSummaries) {
      console.log(`Policy: ${summary.policyName}`);
      console.log(`  Total Devices: ${summary.totalDeviceCount}`);
      console.log(`  Compliant: ${summary.compliantDeviceCount} (${summary.compliancePercentage}%)`);
      console.log(`  Non-Compliant: ${summary.nonCompliantDeviceCount}`);
      console.log(`  In Grace Period: ${summary.inGracePeriodCount}`);
      console.log(`  Errors: ${summary.errorDeviceCount}`);
      console.log(`  Conflicts: ${summary.conflictDeviceCount}`);
      console.log(`  Not Applicable: ${summary.notApplicableDeviceCount}`);

      // Highlight issues
      if (summary.compliancePercentage < 80 && summary.totalDeviceCount > 5) {
        console.log(`  ⚠️  LOW COMPLIANCE RATE!`);
      }

      console.log('');
    }

    // Overall statistics
    const totalDevices = summaries.reduce((sum, s) => sum + s.totalDeviceCount, 0);
    const totalCompliant = summaries.reduce((sum, s) => sum + s.compliantDeviceCount, 0);
    const avgCompliance = totalDevices > 0 ? Math.round((totalCompliant / totalDevices) * 100 * 100) / 100 : 0;

    console.log('Overall Statistics:');
    console.log(`  Average Compliance: ${avgCompliance}%`);
    console.log(`  Total Device-Policy Assignments: ${totalDevices}`);
    console.log(`  Total Compliant: ${totalCompliant}`);

    return summaries;
  } catch (error) {
    console.error('Error fetching policy summaries:', error);
    throw error;
  }
}

// ============================================================================
// Example 4: Get Non-Compliant Devices with Reasons
// ============================================================================

/**
 * Retrieve non-compliant devices with detailed reasons
 */
export async function example4_GetNonCompliantDevicesWithReasons() {
  console.log('\n=== Example 4: Non-Compliant Devices with Reasons ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Fetching non-compliant devices with detailed reasons...');
    const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

    console.log(`\nFound ${nonCompliantDevices.length} non-compliant devices:\n`);

    for (const device of nonCompliantDevices.slice(0, 10)) { // Show first 10
      console.log(`Device: ${device.deviceName}`);
      console.log(`  User: ${device.userPrincipalName}`);
      console.log(`  Platform: ${device.platform} ${device.osVersion}`);
      console.log(`  Model: ${device.manufacturer} ${device.deviceModel}`);
      console.log(`  Compliance State: ${device.complianceState}`);
      console.log(`  Last Sync: ${device.lastSyncDateTime.toLocaleString()}`);

      if (device.gracePeriodExpirationDateTime) {
        const daysLeft = Math.ceil((device.gracePeriodExpirationDateTime.getTime() - Date.now()) / (1000 * 60 * 60 * 24));
        console.log(`  Grace Period: ${daysLeft} days remaining`);
      }

      console.log(`  Non-Compliant Policies: ${device.nonCompliantPolicies.length}`);

      for (const policy of device.nonCompliantPolicies) {
        console.log(`\n  Policy: ${policy.policyName}`);
        console.log(`    Non-Compliant Settings: ${policy.settingCount}`);

        for (const setting of policy.nonCompliantSettings.slice(0, 3)) { // Show first 3 settings
          console.log(`\n    Setting: ${setting.settingName}`);
          console.log(`      State: ${setting.state}`);

          if (setting.currentValue) {
            console.log(`      Current Value: ${setting.currentValue}`);
          }

          if (setting.errorDescription) {
            console.log(`      Error: ${setting.errorDescription}`);
          }

          if (setting.remediationActions && setting.remediationActions.length > 0) {
            console.log(`      Remediation:`);
            for (const action of setting.remediationActions) {
              console.log(`        - ${action}`);
            }
          }
        }

        if (policy.nonCompliantSettings.length > 3) {
          console.log(`\n    ... and ${policy.nonCompliantSettings.length - 3} more settings`);
        }
      }

      console.log('\n' + '─'.repeat(80) + '\n');
    }

    if (nonCompliantDevices.length > 10) {
      console.log(`... and ${nonCompliantDevices.length - 10} more devices\n`);
    }

    // Summary statistics
    const totalNonCompliantSettings = nonCompliantDevices.reduce((sum, d) => sum + d.totalNonCompliantSettings, 0);
    const avgSettingsPerDevice = nonCompliantDevices.length > 0
      ? Math.round(totalNonCompliantSettings / nonCompliantDevices.length)
      : 0;

    console.log('Non-Compliance Summary:');
    console.log(`  Total Non-Compliant Devices: ${nonCompliantDevices.length}`);
    console.log(`  Total Non-Compliant Settings: ${totalNonCompliantSettings}`);
    console.log(`  Average Settings per Device: ${avgSettingsPerDevice}`);

    // Group by platform
    const byPlatform = nonCompliantDevices.reduce((acc, device) => {
      acc[device.platform] = (acc[device.platform] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('\nNon-Compliant Devices by Platform:');
    for (const [platform, count] of Object.entries(byPlatform)) {
      console.log(`  ${platform}: ${count}`);
    }

    return nonCompliantDevices;
  } catch (error) {
    console.error('Error fetching non-compliant devices:', error);
    throw error;
  }
}

// ============================================================================
// Example 5: Get Device Compliance Details
// ============================================================================

/**
 * Get detailed compliance information for all devices
 */
export async function example5_GetDeviceComplianceDetails() {
  console.log('\n=== Example 5: Device Compliance Details ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Fetching detailed device compliance information...');
    const devices = await report.getAllDeviceComplianceDetails();

    console.log(`\nProcessed ${devices.length} devices:\n`);

    // Show details for first 5 devices
    for (const device of devices.slice(0, 5)) {
      console.log(`Device: ${device.deviceName}`);
      console.log(`  ID: ${device.deviceId}`);
      console.log(`  User: ${device.userPrincipalName}`);
      console.log(`  Platform: ${device.platform} ${device.osVersion}`);
      console.log(`  Overall State: ${device.overallComplianceState}`);
      console.log(`  Last Sync: ${device.lastSyncDateTime.toLocaleString()}`);
      console.log(`  Managed: ${device.isManaged ? 'Yes' : 'No'}`);

      if (device.manufacturer && device.model) {
        console.log(`  Device: ${device.manufacturer} ${device.model}`);
      }

      if (device.serialNumber) {
        console.log(`  Serial: ${device.serialNumber}`);
      }

      console.log(`  Policies Applied: ${device.compliancePoliciesCount}`);
      console.log(`  Compliant: ${device.compliantPoliciesCount}`);
      console.log(`  Non-Compliant: ${device.nonCompliantPoliciesCount}`);

      if (device.policies.length > 0) {
        console.log('\n  Policy Details:');
        for (const policy of device.policies) {
          console.log(`    - ${policy.policyName}: ${policy.state} (${policy.settingCount} settings)`);
        }
      }

      console.log('');
    }

    // Statistics
    const compliantDevices = devices.filter(d => d.overallComplianceState === ComplianceState.COMPLIANT).length;
    const nonCompliantDevices = devices.filter(d => d.overallComplianceState === ComplianceState.NON_COMPLIANT).length;
    const compliancePercentage = devices.length > 0
      ? Math.round((compliantDevices / devices.length) * 100 * 100) / 100
      : 0;

    console.log('Device Compliance Statistics:');
    console.log(`  Total Devices: ${devices.length}`);
    console.log(`  Compliant: ${compliantDevices} (${compliancePercentage}%)`);
    console.log(`  Non-Compliant: ${nonCompliantDevices}`);

    return devices;
  } catch (error) {
    console.error('Error fetching device compliance details:', error);
    throw error;
  }
}

// ============================================================================
// Example 6: Get Compliance Trends
// ============================================================================

/**
 * Generate compliance trends over time
 */
export async function example6_GetComplianceTrends() {
  console.log('\n=== Example 6: Compliance Trends Over Time ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Generating compliance trends for the last 30 days...');
    const trends = await report.getComplianceTrends(30);

    if (!trends) {
      console.log('No trend data available');
      return null;
    }

    console.log(`\nTrend Analysis (${trends.startDate.toLocaleDateString()} to ${trends.endDate.toLocaleDateString()}):\n`);

    console.log(`Overall Trend: ${trends.trend.toUpperCase()}`);
    console.log(`Average Compliance: ${trends.averageCompliancePercentage}%`);
    console.log(`Compliance Change: ${trends.complianceChange > 0 ? '+' : ''}${trends.complianceChange}%`);

    console.log(`\nPeak Compliance:`);
    console.log(`  Date: ${trends.peakCompliance.date.toLocaleDateString()}`);
    console.log(`  Compliance: ${trends.peakCompliance.compliancePercentage}%`);
    console.log(`  Compliant Devices: ${trends.peakCompliance.compliantCount}`);

    console.log(`\nLowest Compliance:`);
    console.log(`  Date: ${trends.lowestCompliance.date.toLocaleDateString()}`);
    console.log(`  Compliance: ${trends.lowestCompliance.compliancePercentage}%`);
    console.log(`  Compliant Devices: ${trends.lowestCompliance.compliantCount}`);

    // Show recent trend (last 7 days)
    console.log(`\nRecent Trend (Last 7 Days):`);
    const recentDataPoints = trends.dataPoints.slice(-7);

    for (const dataPoint of recentDataPoints) {
      console.log(`  ${dataPoint.date.toLocaleDateString()}: ${dataPoint.compliancePercentage}% ` +
                  `(${dataPoint.compliantCount}/${dataPoint.totalCount})`);
    }

    // Trend visualization (simple text-based)
    console.log(`\nCompliance Trend Chart:`);
    console.log('  100% │');

    const chartData = trends.dataPoints.slice(-14); // Last 14 days
    for (let i = 100; i >= 0; i -= 10) {
      const line = chartData.map(dp => {
        const val = dp.compliancePercentage;
        if (Math.abs(val - i) < 5) return '●';
        return ' ';
      }).join('');

      console.log(`  ${i.toString().padStart(3)}% │ ${line}`);
    }

    console.log('    0% └' + '─'.repeat(chartData.length));
    console.log('       ' + chartData.map((_, i) => i % 2 === 0 ? '│' : ' ').join(''));

    return trends;
  } catch (error) {
    console.error('Error generating compliance trends:', error);
    throw error;
  }
}

// ============================================================================
// Example 7: Filtered Compliance Report (Windows Only)
// ============================================================================

/**
 * Generate a filtered compliance report for Windows devices only
 */
export async function example7_FilteredComplianceReport() {
  console.log('\n=== Example 7: Filtered Compliance Report (Windows Only) ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Fetching Windows compliance policies...');
    const windowsPolicies = await report.getAllCompliancePolicies({
      platform: CompliancePlatform.WINDOWS_10
    });

    console.log(`Found ${windowsPolicies.length} Windows 10 policies\n`);

    console.log('Fetching Windows device compliance details...');
    const allDevices = await report.getAllDeviceComplianceDetails();
    const windowsDevices = allDevices.filter(d =>
      d.platform.toLowerCase().includes('windows')
    );

    console.log(`Found ${windowsDevices.length} Windows devices\n`);

    const compliantWindows = windowsDevices.filter(d =>
      d.overallComplianceState === ComplianceState.COMPLIANT
    ).length;
    const compliancePercentage = windowsDevices.length > 0
      ? Math.round((compliantWindows / windowsDevices.length) * 100 * 100) / 100
      : 0;

    console.log('Windows Device Compliance:');
    console.log(`  Total: ${windowsDevices.length}`);
    console.log(`  Compliant: ${compliantWindows} (${compliancePercentage}%)`);
    console.log(`  Non-Compliant: ${windowsDevices.length - compliantWindows}`);

    // Export Windows-specific report
    const outputDir = './reports/compliance/windows';
    await fs.mkdir(outputDir, { recursive: true });

    const reportData = {
      metadata: {
        reportName: 'Windows Device Compliance Report',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: windowsDevices.length,
        parameters: { platform: 'Windows 10' }
      },
      data: windowsDevices,
      summary: {
        totalDevices: windowsDevices.length,
        compliantDevices: compliantWindows,
        compliancePercentage,
        totalPolicies: windowsPolicies.length
      }
    };

    const jsonPath = await OutputFormatter.format(reportData, 'json', outputDir, true);
    console.log(`\nWindows compliance report saved to: ${jsonPath}`);

    return { windowsPolicies, windowsDevices };
  } catch (error) {
    console.error('Error generating Windows compliance report:', error);
    throw error;
  }
}

// ============================================================================
// Example 8: Export Non-Compliant Devices Report
// ============================================================================

/**
 * Generate and export a focused report on non-compliant devices
 */
export async function example8_ExportNonCompliantDevicesReport() {
  console.log('\n=== Example 8: Export Non-Compliant Devices Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new DeviceComplianceReport(graphClient, config);

    console.log('Fetching non-compliant devices...');
    const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

    // Transform data for export
    const exportData = nonCompliantDevices.map(device => ({
      deviceName: device.deviceName,
      userPrincipalName: device.userPrincipalName,
      userEmail: device.userEmail || 'N/A',
      platform: device.platform,
      osVersion: device.osVersion,
      manufacturer: device.manufacturer || 'N/A',
      model: device.deviceModel || 'N/A',
      complianceState: device.complianceState,
      nonCompliantPoliciesCount: device.nonCompliantPolicies.length,
      totalNonCompliantSettings: device.totalNonCompliantSettings,
      lastSyncDateTime: device.lastSyncDateTime.toLocaleString(),
      gracePeriodExpiration: device.gracePeriodExpirationDateTime?.toLocaleString() || 'N/A',
      topIssues: device.nonCompliantPolicies
        .flatMap(p => p.nonCompliantSettings.slice(0, 3))
        .map(s => s.settingName)
        .join('; ')
    }));

    const reportData = {
      metadata: {
        reportName: 'Non-Compliant Devices Report',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: exportData.length
      },
      data: exportData,
      summary: {
        totalNonCompliantDevices: nonCompliantDevices.length,
        totalNonCompliantSettings: nonCompliantDevices.reduce((sum, d) => sum + d.totalNonCompliantSettings, 0)
      }
    };

    const outputDir = './reports/compliance/non-compliant';
    await fs.mkdir(outputDir, { recursive: true });

    const csvPath = await OutputFormatter.format(reportData, 'csv', outputDir, true);
    console.log(`CSV report saved to: ${csvPath}`);

    const htmlPath = await OutputFormatter.format(reportData, 'html', outputDir, true);
    console.log(`HTML report saved to: ${htmlPath}`);

    console.log(`\nGenerated non-compliant devices report with ${exportData.length} devices`);

    return reportData;
  } catch (error) {
    console.error('Error exporting non-compliant devices report:', error);
    throw error;
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Creates a Graph API client with authentication
 */
async function createGraphClient(): Promise<Client> {
  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID || '',
    process.env.AZURE_CLIENT_ID || '',
    process.env.AZURE_CLIENT_SECRET || ''
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default']
  });

  return Client.initWithMiddleware({ authProvider });
}

/**
 * Gets default configuration
 */
function getDefaultConfig(): AppConfig {
  return {
    authentication: {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET,
      authMethod: 'clientSecret'
    },
    reports: {
      enabled: ['device-compliance-report'],
      disabled: [],
      settings: {}
    },
    output: {
      defaultFormat: 'json',
      directory: './reports',
      includeTimestamp: true,
      compression: false
    },
    scheduler: {
      enabled: false,
      timezone: 'UTC',
      schedules: []
    },
    logging: {
      level: 'info',
      file: './logs/app.log',
      console: true,
      maxSize: '10m',
      maxFiles: 5
    }
  };
}

// ============================================================================
// Main Example Runner
// ============================================================================

/**
 * Run all examples
 */
export async function runAllExamples() {
  console.log('\n' + '='.repeat(80));
  console.log('DEVICE COMPLIANCE REPORT - EXAMPLES');
  console.log('='.repeat(80));

  try {
    // Example 1: Basic report
    await example1_BasicComplianceReport();

    // Example 2: Get policies
    await example2_GetAllCompliancePolicies();

    // Example 3: Policy summaries
    await example3_GetPolicyDeviceStateSummaries();

    // Example 4: Non-compliant devices
    await example4_GetNonCompliantDevicesWithReasons();

    // Example 5: Device details
    await example5_GetDeviceComplianceDetails();

    // Example 6: Trends
    await example6_GetComplianceTrends();

    // Example 7: Filtered report
    await example7_FilteredComplianceReport();

    // Example 8: Export non-compliant
    await example8_ExportNonCompliantDevicesReport();

    console.log('\n' + '='.repeat(80));
    console.log('ALL EXAMPLES COMPLETED SUCCESSFULLY');
    console.log('='.repeat(80) + '\n');
  } catch (error) {
    console.error('\nExample execution failed:', error);
    throw error;
  }
}

// Run examples if executed directly
if (require.main === module) {
  runAllExamples()
    .then(() => process.exit(0))
    .catch(() => process.exit(1));
}
