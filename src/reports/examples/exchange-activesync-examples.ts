/**
 * Exchange ActiveSync Reports - Usage Examples
 *
 * This file demonstrates various ways to use the ExchangeActiveSyncReports module
 * to generate comprehensive Exchange ActiveSync reports from Microsoft Intune.
 *
 * @module exchange-activesync-examples
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { ClientSecretCredential } from '@azure/identity';
import {
  ExchangeActiveSyncReports,
  ExchangeAccessState,
  ExchangeAccessStateReason
} from '../configmgr/exchange-activesync-reports';
import { OutputFormatter } from '../../formatters/output-formatter';
import { AppConfig } from '../../types';
import * as path from 'path';
import * as fs from 'fs/promises';

// ============================================================================
// Example 1: Complete Exchange ActiveSync Report
// ============================================================================

/**
 * Generate a comprehensive Exchange ActiveSync report with all sections
 */
export async function example1_CompleteExchangeActiveSyncReport() {
  console.log('\n=== Example 1: Complete Exchange ActiveSync Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    console.log('Generating comprehensive Exchange ActiveSync report...');
    const reportData = await report.execute();

    const summary = reportData.summary as any;
    console.log('\nReport Summary:');
    console.log(`  Total Mobile Devices: ${summary.totalMobileDevices}`);
    console.log(`  Active Devices: ${summary.activeDevices}`);
    console.log(`  Inactive Devices: ${summary.inactiveDevices}`);
    console.log(`  Allowed Devices: ${summary.allowedDevices}`);
    console.log(`  Blocked Devices: ${summary.blockedDevices}`);
    console.log(`  Quarantined Devices: ${summary.quarantinedDevices}`);
    console.log(`  EAS Activated Devices: ${summary.easActivatedDevices}`);
    console.log(`  Compliance Percentage: ${summary.compliancePercentage}%`);
    console.log(`  Average Days Since Last Sync: ${summary.averageDaysSinceLastSync}`);

    console.log('\nDevices by Platform:');
    for (const [platform, count] of Object.entries(summary.devicesByPlatform)) {
      console.log(`  ${platform}: ${count}`);
    }

    console.log('\nDevices by Access State:');
    for (const [state, count] of Object.entries(summary.devicesByAccessState)) {
      console.log(`  ${state}: ${count}`);
    }

    // Export to all formats
    const outputDir = './reports/exchange-activesync';
    await fs.mkdir(outputDir, { recursive: true });

    const jsonPath = await report.exportToJson(reportData, outputDir, true);
    console.log(`\nJSON report saved to: ${jsonPath}`);

    const csvPath = await report.exportToCsv(reportData, outputDir, true);
    console.log(`CSV report saved to: ${csvPath}`);

    const htmlPath = await report.exportToHtml(reportData, outputDir, true);
    console.log(`HTML report saved to: ${htmlPath}`);

    return reportData;
  } catch (error) {
    console.error('Error generating Exchange ActiveSync report:', error);
    throw error;
  }
}

// ============================================================================
// Example 2: ActiveSync Policy Compliance Status (Report 8)
// ============================================================================

/**
 * Get compliance status of ActiveSync mailbox policies
 */
export async function example2_GetActiveSyncPolicyCompliance() {
  console.log('\n=== Example 2: ActiveSync Policy Compliance Status ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    console.log('Fetching ActiveSync policy compliance status...');
    const policyCompliance = await report.getActiveSyncPolicyCompliance();

    console.log(`\nFound ${policyCompliance.length} policy compliance summaries:\n`);

    for (const policy of policyCompliance) {
      console.log(`Policy: ${policy.policyName}`);
      console.log(`  Total Devices: ${policy.totalDevices}`);
      console.log(`  Allowed: ${policy.allowedDevices} (${policy.compliancePercentage}%)`);
      console.log(`  Blocked: ${policy.blockedDevices}`);
      console.log(`  Quarantined: ${policy.quarantinedDevices}`);
      console.log(`  Unknown: ${policy.unknownDevices}`);
      console.log(`  Last Updated: ${policy.lastUpdated.toLocaleString()}`);

      console.log('\n  Devices by Platform:');
      for (const [platform, count] of Object.entries(policy.devicesByPlatform)) {
        console.log(`    ${platform}: ${count}`);
      }

      console.log('\n  Devices by Access State:');
      for (const [state, count] of Object.entries(policy.devicesByAccessState)) {
        console.log(`    ${state}: ${count}`);
      }

      console.log('');
    }

    // Identify policies with low compliance
    const lowCompliancePolicies = policyCompliance.filter(p =>
      p.compliancePercentage < 80 && p.totalDevices > 5
    );

    if (lowCompliancePolicies.length > 0) {
      console.log('⚠️  Policies with Low Compliance:');
      for (const policy of lowCompliancePolicies) {
        console.log(`  - ${policy.policyName}: ${policy.compliancePercentage}%`);
      }
    }

    return policyCompliance;
  } catch (error) {
    console.error('Error fetching policy compliance:', error);
    throw error;
  }
}

// ============================================================================
// Example 3: Inactive Mobile Devices (Report 15)
// ============================================================================

/**
 * Identify devices that haven't synced within the last 30 days
 */
export async function example3_GetInactiveMobileDevices() {
  console.log('\n=== Example 3: Inactive Mobile Devices ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    const inactiveDaysThreshold = 30;
    console.log(`Fetching devices inactive for more than ${inactiveDaysThreshold} days...`);

    const inactiveDevices = await report.getInactiveMobileDevices(inactiveDaysThreshold);

    console.log(`\nFound ${inactiveDevices.length} inactive devices:\n`);

    // Show top 10 most inactive devices
    for (const device of inactiveDevices.slice(0, 10)) {
      console.log(`Device: ${device.deviceName}`);
      console.log(`  User: ${device.userPrincipalName}`);
      console.log(`  Platform: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`  Model: ${device.manufacturer} ${device.model}`);
      console.log(`  Last Sync: ${device.lastSyncDateTime.toLocaleString()}`);
      console.log(`  Days Since Last Sync: ${device.daysSinceLastSync}`);
      console.log(`  Exchange Access State: ${device.exchangeAccessState}`);
      console.log(`  Compliance State: ${device.complianceState}`);

      if (device.serialNumber) {
        console.log(`  Serial Number: ${device.serialNumber}`);
      }

      console.log('');
    }

    if (inactiveDevices.length > 10) {
      console.log(`... and ${inactiveDevices.length - 10} more devices\n`);
    }

    // Statistics
    const byPlatform = inactiveDevices.reduce((acc, device) => {
      acc[device.operatingSystem] = (acc[device.operatingSystem] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Inactive Devices by Platform:');
    for (const [platform, count] of Object.entries(byPlatform)) {
      console.log(`  ${platform}: ${count}`);
    }

    // Calculate average inactivity
    const avgInactivity = inactiveDevices.length > 0
      ? Math.round(inactiveDevices.reduce((sum, d) => sum + d.daysSinceLastSync, 0) / inactiveDevices.length)
      : 0;
    console.log(`\nAverage Inactivity: ${avgInactivity} days`);

    // Most inactive device
    if (inactiveDevices.length > 0) {
      const mostInactive = inactiveDevices[0];
      console.log(`\nMost Inactive Device: ${mostInactive.deviceName}`);
      console.log(`  Days Inactive: ${mostInactive.daysSinceLastSync}`);
    }

    return inactiveDevices;
  } catch (error) {
    console.error('Error fetching inactive devices:', error);
    throw error;
  }
}

// ============================================================================
// Example 4: Mobile Device Compliance Details (Report 21)
// ============================================================================

/**
 * Get detailed compliance information for all mobile devices
 */
export async function example4_GetMobileDeviceComplianceDetails() {
  console.log('\n=== Example 4: Mobile Device Compliance Details ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    console.log('Fetching detailed device compliance information...');
    const deviceDetails = await report.getAllMobileDeviceComplianceDetails();

    console.log(`\nProcessed ${deviceDetails.length} devices:\n`);

    // Show details for first 5 devices
    for (const device of deviceDetails.slice(0, 5)) {
      console.log(`Device: ${device.deviceName}`);
      console.log(`  User: ${device.userPrincipalName}`);
      console.log(`  Email: ${device.emailAddress}`);
      console.log(`  Platform: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`  Model: ${device.manufacturer} ${device.model}`);
      console.log(`  Exchange Access State: ${device.exchangeAccessState}`);
      console.log(`  Exchange Access Reason: ${device.exchangeAccessStateReason}`);
      console.log(`  Compliance State: ${device.complianceState}`);
      console.log(`  Last Sync: ${device.lastSyncDateTime.toLocaleString()}`);
      console.log(`  EAS Activated: ${device.easActivated ? 'Yes' : 'No'}`);

      if (device.easActivated && device.easActivationDateTime) {
        console.log(`  EAS Activation Date: ${device.easActivationDateTime.toLocaleString()}`);
      }

      console.log(`  Compliance Policies: ${device.compliancePolicies.length}`);
      for (const policy of device.compliancePolicies) {
        console.log(`    - ${policy.policyName}: ${policy.complianceState}`);
        console.log(`      Settings: ${policy.settingStates.length}`);

        // Show first 3 settings
        for (const setting of policy.settingStates.slice(0, 3)) {
          console.log(`        • ${setting.settingName}: ${setting.state}`);
          if (setting.currentValue) {
            console.log(`          Current: ${setting.currentValue}`);
          }
        }

        if (policy.settingStates.length > 3) {
          console.log(`        ... and ${policy.settingStates.length - 3} more settings`);
        }
      }

      if (device.isSupervised !== undefined) {
        console.log(`  Supervised: ${device.isSupervised ? 'Yes' : 'No'}`);
      }

      if (device.jailBroken) {
        console.log(`  Jailbroken: ${device.jailBroken}`);
      }

      console.log('');
    }

    // Statistics
    const compliantCount = deviceDetails.filter(d =>
      d.complianceState.toLowerCase() === 'compliant'
    ).length;

    const easActivatedCount = deviceDetails.filter(d => d.easActivated).length;

    console.log('Device Compliance Statistics:');
    console.log(`  Total Devices: ${deviceDetails.length}`);
    console.log(`  Compliant: ${compliantCount} (${Math.round((compliantCount / deviceDetails.length) * 100)}%)`);
    console.log(`  EAS Activated: ${easActivatedCount}`);

    // Group by Exchange Access State
    const byAccessState = deviceDetails.reduce((acc, device) => {
      acc[device.exchangeAccessState] = (acc[device.exchangeAccessState] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('\nDevices by Exchange Access State:');
    for (const [state, count] of Object.entries(byAccessState)) {
      console.log(`  ${state}: ${count}`);
    }

    return deviceDetails;
  } catch (error) {
    console.error('Error fetching device compliance details:', error);
    throw error;
  }
}

// ============================================================================
// Example 5: Settings Summary for Mobile Devices (Report 34)
// ============================================================================

/**
 * Get aggregated settings summary across all mobile devices
 */
export async function example5_GetMobileDeviceSettingsSummary() {
  console.log('\n=== Example 5: Mobile Device Settings Summary ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    console.log('Fetching mobile device settings summary...');
    const settingsSummary = await report.getMobileDeviceSettingsSummary();

    console.log(`\nFound ${settingsSummary.length} settings:\n`);

    // Show top 20 most common settings
    for (const setting of settingsSummary.slice(0, 20)) {
      console.log(`Setting: ${setting.settingName}`);
      console.log(`  Description: ${setting.settingDescription}`);
      console.log(`  Category: ${setting.settingCategory}`);
      console.log(`  Devices: ${setting.deviceCount}`);
      console.log(`  Enabled: ${setting.enabledCount} (${setting.percentage}%)`);
      console.log(`  Disabled: ${setting.disabledCount}`);
      console.log(`  Not Configured: ${setting.notConfiguredCount}`);

      if (setting.affectedPolicies.length > 0) {
        console.log(`  Affected Policies: ${setting.affectedPolicies.join(', ')}`);
      }

      console.log('  Platform Distribution:');
      for (const [platform, count] of Object.entries(setting.platforms)) {
        console.log(`    ${platform}: ${count}`);
      }

      console.log('');
    }

    if (settingsSummary.length > 20) {
      console.log(`... and ${settingsSummary.length - 20} more settings\n`);
    }

    // Group by category
    const byCategory = settingsSummary.reduce((acc, setting) => {
      acc[setting.settingCategory] = (acc[setting.settingCategory] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Settings by Category:');
    for (const [category, count] of Object.entries(byCategory)) {
      console.log(`  ${category}: ${count}`);
    }

    // Find settings with low compliance
    const lowComplianceSettings = settingsSummary.filter(s =>
      s.percentage < 70 && s.deviceCount > 5
    );

    if (lowComplianceSettings.length > 0) {
      console.log('\n⚠️  Settings with Low Compliance:');
      for (const setting of lowComplianceSettings.slice(0, 10)) {
        console.log(`  - ${setting.settingName}: ${setting.percentage}% (${setting.deviceCount} devices)`);
      }
    }

    return settingsSummary;
  } catch (error) {
    console.error('Error fetching settings summary:', error);
    throw error;
  }
}

// ============================================================================
// Example 6: Filtered Report - Blocked Devices Only
// ============================================================================

/**
 * Generate a report focusing on blocked Exchange ActiveSync devices
 */
export async function example6_GetBlockedDevicesReport() {
  console.log('\n=== Example 6: Blocked Devices Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    console.log('Fetching blocked devices...');
    const allDevices = await report.getAllMobileDeviceComplianceDetails({
      exchangeAccessState: ExchangeAccessState.BLOCKED
    });

    console.log(`\nFound ${allDevices.length} blocked devices:\n`);

    for (const device of allDevices) {
      console.log(`Device: ${device.deviceName}`);
      console.log(`  User: ${device.userPrincipalName}`);
      console.log(`  Platform: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`  Block Reason: ${device.exchangeAccessStateReason}`);
      console.log(`  Compliance State: ${device.complianceState}`);
      console.log(`  Last Sync: ${device.lastSyncDateTime.toLocaleString()}`);

      if (device.compliancePolicies.length > 0) {
        console.log(`  Non-Compliant Policies:`);
        for (const policy of device.compliancePolicies) {
          if (policy.complianceState.toLowerCase() !== 'compliant') {
            console.log(`    - ${policy.policyName}: ${policy.complianceState}`);
          }
        }
      }

      console.log('');
    }

    // Group by block reason
    const byReason = allDevices.reduce((acc, device) => {
      acc[device.exchangeAccessStateReason] = (acc[device.exchangeAccessStateReason] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Blocked Devices by Reason:');
    for (const [reason, count] of Object.entries(byReason)) {
      console.log(`  ${reason}: ${count}`);
    }

    // Export blocked devices report
    const outputDir = './reports/exchange-activesync/blocked';
    await fs.mkdir(outputDir, { recursive: true });

    const reportData = {
      metadata: {
        reportName: 'Blocked Exchange ActiveSync Devices',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: allDevices.length
      },
      data: allDevices,
      summary: {
        totalBlockedDevices: allDevices.length,
        blockReasons: byReason
      }
    };

    const csvPath = await OutputFormatter.format(reportData, 'csv', outputDir, true);
    console.log(`\nBlocked devices report saved to: ${csvPath}`);

    return allDevices;
  } catch (error) {
    console.error('Error fetching blocked devices:', error);
    throw error;
  }
}

// ============================================================================
// Example 7: Inactive Devices with Multiple Thresholds
// ============================================================================

/**
 * Compare inactive devices across different time thresholds
 */
export async function example7_CompareInactivityThresholds() {
  console.log('\n=== Example 7: Compare Inactivity Thresholds ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    const thresholds = [7, 14, 30, 60, 90];

    console.log('Analyzing device inactivity across multiple thresholds...\n');

    const results = [];
    for (const threshold of thresholds) {
      const devices = await report.getInactiveMobileDevices(threshold);
      results.push({
        threshold,
        count: devices.length,
        devices
      });

      console.log(`Devices inactive for ${threshold}+ days: ${devices.length}`);
    }

    console.log('\nInactivity Trend:');
    console.log('Days  │ Count │ Bar');
    console.log('──────┼───────┼────────────────────────────────');

    const maxCount = Math.max(...results.map(r => r.count));
    for (const result of results) {
      const barLength = Math.round((result.count / maxCount) * 30);
      const bar = '█'.repeat(barLength);
      console.log(`${result.threshold.toString().padStart(3)}+  │ ${result.count.toString().padStart(5)} │ ${bar}`);
    }

    // Identify devices at risk (inactive 7-30 days)
    const recentlyInactive = results[2].devices.filter(d => d.daysSinceLastSync < 30);

    console.log(`\n⚠️  Devices Recently Inactive (30+ days): ${recentlyInactive.length}`);
    console.log('These devices may need attention to prevent further inactivity.\n');

    return results;
  } catch (error) {
    console.error('Error comparing inactivity thresholds:', error);
    throw error;
  }
}

// ============================================================================
// Example 8: Platform-Specific Exchange ActiveSync Report
// ============================================================================

/**
 * Generate Exchange ActiveSync report for iOS devices only
 */
export async function example8_GetIOSExchangeReport() {
  console.log('\n=== Example 8: iOS Exchange ActiveSync Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = getDefaultConfig();
    const report = new ExchangeActiveSyncReports(graphClient, config);

    console.log('Fetching iOS device compliance details...');
    const allDevices = await report.getAllMobileDeviceComplianceDetails({
      platform: 'iOS'
    });

    console.log(`\nFound ${allDevices.length} iOS devices\n`);

    // Statistics
    const compliantCount = allDevices.filter(d =>
      d.complianceState.toLowerCase() === 'compliant'
    ).length;

    const allowedCount = allDevices.filter(d =>
      d.exchangeAccessState === ExchangeAccessState.ALLOWED
    ).length;

    const easActivatedCount = allDevices.filter(d => d.easActivated).length;
    const supervisedCount = allDevices.filter(d => d.isSupervised).length;

    console.log('iOS Device Statistics:');
    console.log(`  Total iOS Devices: ${allDevices.length}`);
    console.log(`  Compliant: ${compliantCount} (${Math.round((compliantCount / allDevices.length) * 100)}%)`);
    console.log(`  Exchange Allowed: ${allowedCount} (${Math.round((allowedCount / allDevices.length) * 100)}%)`);
    console.log(`  EAS Activated: ${easActivatedCount}`);
    console.log(`  Supervised: ${supervisedCount}`);

    // OS Version distribution
    const byVersion = allDevices.reduce((acc, device) => {
      acc[device.osVersion] = (acc[device.osVersion] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('\niOS Version Distribution:');
    for (const [version, count] of Object.entries(byVersion).sort(([a], [b]) => b.localeCompare(a))) {
      const percentage = Math.round((count / allDevices.length) * 100);
      console.log(`  iOS ${version}: ${count} (${percentage}%)`);
    }

    // Check for inactive iOS devices
    const inactiveDevices = await report.getInactiveMobileDevices(30, {
      platform: 'iOS'
    });

    console.log(`\nInactive iOS Devices (30+ days): ${inactiveDevices.length}`);

    // Export iOS-specific report
    const outputDir = './reports/exchange-activesync/ios';
    await fs.mkdir(outputDir, { recursive: true });

    const reportData = {
      metadata: {
        reportName: 'iOS Exchange ActiveSync Report',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: allDevices.length,
        parameters: { platform: 'iOS' }
      },
      data: allDevices,
      summary: {
        totalDevices: allDevices.length,
        compliantDevices: compliantCount,
        allowedDevices: allowedCount,
        easActivatedDevices: easActivatedCount,
        supervisedDevices: supervisedCount,
        inactiveDevices: inactiveDevices.length,
        versionDistribution: byVersion
      }
    };

    const jsonPath = await OutputFormatter.format(reportData, 'json', outputDir, true);
    console.log(`\niOS Exchange ActiveSync report saved to: ${jsonPath}`);

    return { allDevices, inactiveDevices };
  } catch (error) {
    console.error('Error generating iOS Exchange report:', error);
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
      enabled: ['exchange-activesync-reports'],
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
  console.log('EXCHANGE ACTIVESYNC REPORTS - EXAMPLES');
  console.log('='.repeat(80));

  try {
    // Example 1: Complete report
    await example1_CompleteExchangeActiveSyncReport();

    // Example 2: Policy compliance
    await example2_GetActiveSyncPolicyCompliance();

    // Example 3: Inactive devices
    await example3_GetInactiveMobileDevices();

    // Example 4: Device compliance details
    await example4_GetMobileDeviceComplianceDetails();

    // Example 5: Settings summary
    await example5_GetMobileDeviceSettingsSummary();

    // Example 6: Blocked devices
    await example6_GetBlockedDevicesReport();

    // Example 7: Inactivity thresholds
    await example7_CompareInactivityThresholds();

    // Example 8: iOS-specific report
    await example8_GetIOSExchangeReport();

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
