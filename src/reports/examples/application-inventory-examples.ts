/**
 * Application Inventory Report - Usage Examples
 *
 * This file contains comprehensive examples demonstrating how to use the
 * ApplicationInventoryReport class and its various features.
 *
 * @module application-inventory-examples
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';
import {
  ApplicationInventoryReport,
  ManagedAppType,
  AppInstallationStatus,
  AppProtectionStatus
} from '../application-inventory';

// ===========================
// Example 1: Complete Inventory Report
// ===========================

/**
 * Generates a complete application inventory report with all data
 */
export async function example1_CompleteInventoryReport(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 1: Generating Complete Application Inventory Report');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const result = await report.execute();

    console.log('\nReport Metadata:');
    console.log(`- Report Name: ${result.metadata.reportName}`);
    console.log(`- Generated At: ${result.metadata.generatedAt}`);
    console.log(`- Total Records: ${result.metadata.recordCount}`);
    console.log(`- Duration: ${result.metadata.parameters?.duration}`);

    console.log('\nOverview:');
    console.log(`- Total Managed Apps: ${result.summary.overview.totalManagedApps}`);
    console.log(`- Total Discovered Apps: ${result.summary.overview.totalDiscoveredApps}`);
    console.log(`- Total Devices: ${result.summary.overview.totalDevices}`);
    console.log(`- Total Protection Policies: ${result.summary.overview.totalProtectionPolicies}`);
    console.log(`- Total Configuration Policies: ${result.summary.overview.totalConfigurationPolicies}`);

    console.log('\nApps by Platform:');
    console.log(`- iOS: ${result.summary.appsByPlatform.iOS}`);
    console.log(`- Android: ${result.summary.appsByPlatform.Android}`);
    console.log(`- Windows: ${result.summary.appsByPlatform.Windows}`);
    console.log(`- macOS: ${result.summary.appsByPlatform.macOS}`);
    console.log(`- Web: ${result.summary.appsByPlatform.Web}`);

    console.log('\nInstallation Summary:');
    console.log(`- Installed: ${result.summary.installationSummary.installed}`);
    console.log(`- Failed: ${result.summary.installationSummary.failed}`);
    console.log(`- Pending: ${result.summary.installationSummary.pending}`);
    console.log(`- Not Installed: ${result.summary.installationSummary.notInstalled}`);

    console.log('\nProtection Summary:');
    console.log(`- Compliant: ${result.summary.protectionSummary.compliant}`);
    console.log(`- Non-Compliant: ${result.summary.protectionSummary.nonCompliant}`);
    console.log(`- Not Applicable: ${result.summary.protectionSummary.notApplicable}`);

    console.log('\nTop 10 Most Deployed Apps:');
    result.summary.insights.topDiscoveredApps.forEach((app, index) => {
      console.log(`  ${index + 1}. ${app.name} (v${app.version}) - ${app.deviceCount} devices`);
    });

    if (result.summary.insights.topFailedApps.length > 0) {
      console.log('\nTop 10 Apps with Most Failures:');
      result.summary.insights.topFailedApps.forEach((app, index) => {
        console.log(`  ${index + 1}. ${app.name} - ${app.failureCount} failures`);
      });
    }

    console.log('\n' + '='.repeat(70));
    console.log('Report completed successfully!\n');
  } catch (error) {
    console.error('Error generating report:', error);
  }
}

// ===========================
// Example 2: Managed Apps Inventory
// ===========================

/**
 * Retrieves and displays all managed applications
 */
export async function example2_ManagedAppsInventory(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 2: Managed Apps Inventory');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const managedApps = await report.getManagedApps();

    console.log(`\nFound ${managedApps.length} managed applications\n`);

    // Display first 10 apps with details
    managedApps.slice(0, 10).forEach((app, index) => {
      console.log(`${index + 1}. ${app.displayName}`);
      console.log(`   Publisher: ${app.publisher}`);
      console.log(`   Type: ${app.appType}`);
      console.log(`   Version: ${app.version || 'N/A'}`);
      console.log(`   Assignments: ${app.assignmentCount}`);

      if (app.installSummary) {
        console.log(`   Install Status:`);
        console.log(`     - Installed: ${app.installSummary.installedDeviceCount} devices`);
        console.log(`     - Failed: ${app.installSummary.failedDeviceCount} devices`);
        console.log(`     - Pending: ${app.installSummary.pendingInstallDeviceCount} devices`);
      }

      console.log(`   Created: ${app.createdDateTime.toLocaleDateString()}`);
      console.log('');
    });

    // Group by platform
    const byPlatform = managedApps.reduce((acc, app) => {
      acc[app.appType] = (acc[app.appType] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Apps by Platform:');
    Object.entries(byPlatform).forEach(([platform, count]) => {
      console.log(`  ${platform}: ${count}`);
    });

  } catch (error) {
    console.error('Error fetching managed apps:', error);
  }
}

// ===========================
// Example 3: Filter iOS Apps Only
// ===========================

/**
 * Retrieves only iOS applications
 */
export async function example3_FilterIOSApps(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 3: Filter iOS Apps Only');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const iosApps = await report.getManagedApps({
      platform: ManagedAppType.IOS
    });

    console.log(`\nFound ${iosApps.length} iOS applications\n`);

    iosApps.slice(0, 15).forEach((app, index) => {
      console.log(`${index + 1}. ${app.displayName}`);
      console.log(`   Publisher: ${app.publisher}`);
      console.log(`   Bundle ID: ${app.bundleId || 'N/A'}`);
      console.log(`   Version: ${app.version || 'N/A'}`);

      if (app.installSummary) {
        const totalDevices = app.installSummary.installedDeviceCount +
          app.installSummary.failedDeviceCount +
          app.installSummary.pendingInstallDeviceCount;
        const successRate = totalDevices > 0
          ? ((app.installSummary.installedDeviceCount / totalDevices) * 100).toFixed(1)
          : '0';
        console.log(`   Success Rate: ${successRate}% (${app.installSummary.installedDeviceCount}/${totalDevices})`);
      }
      console.log('');
    });

  } catch (error) {
    console.error('Error fetching iOS apps:', error);
  }
}

// ===========================
// Example 4: Discovered Apps
// ===========================

/**
 * Retrieves all discovered applications on managed devices
 */
export async function example4_DiscoveredApps(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 4: Discovered Apps on Devices');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const discoveredApps = await report.getDiscoveredApps();

    console.log(`\nFound ${discoveredApps.length} discovered applications\n`);

    // Sort by device count (most deployed first)
    const sortedApps = discoveredApps.sort((a, b) => b.deviceCount - a.deviceCount);

    console.log('Top 20 Most Deployed Apps:\n');
    sortedApps.slice(0, 20).forEach((app, index) => {
      console.log(`${index + 1}. ${app.displayName} (v${app.version})`);
      console.log(`   Installed on: ${app.deviceCount} devices`);
      console.log(`   Platform: ${app.platform || 'Unknown'}`);

      if (app.sizeInByte) {
        const sizeMB = (app.sizeInByte / 1024 / 1024).toFixed(2);
        console.log(`   Size: ${sizeMB} MB`);
      }
      console.log('');
    });

    // Show apps installed on 50+ devices
    const widelyDeployed = discoveredApps.filter(app => app.deviceCount >= 50);
    console.log(`\nApps installed on 50+ devices: ${widelyDeployed.length}`);

  } catch (error) {
    console.error('Error fetching discovered apps:', error);
  }
}

// ===========================
// Example 5: Find Specific Apps
// ===========================

/**
 * Finds specific apps by name across all devices
 */
export async function example5_FindSpecificApps(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 5: Find Specific Apps (Microsoft, Adobe)');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    // Find Microsoft apps
    const microsoftApps = await report.getManagedApps({
      publisher: 'Microsoft'
    });

    console.log(`\nFound ${microsoftApps.length} Microsoft apps:\n`);

    microsoftApps.slice(0, 10).forEach((app, index) => {
      console.log(`${index + 1}. ${app.displayName}`);
      console.log(`   Type: ${app.appType}`);
      console.log(`   Version: ${app.version || 'N/A'}`);
      if (app.installSummary) {
        console.log(`   Installed on: ${app.installSummary.installedDeviceCount} devices`);
      }
      console.log('');
    });

    // Find discovered Adobe apps
    const discoveredApps = await report.getDiscoveredApps({
      appName: 'Adobe'
    });

    console.log(`\nFound ${discoveredApps.length} Adobe apps (discovered):\n`);

    discoveredApps.forEach((app, index) => {
      console.log(`${index + 1}. ${app.displayName} (v${app.version})`);
      console.log(`   Devices: ${app.deviceCount}`);
      console.log('');
    });

  } catch (error) {
    console.error('Error finding specific apps:', error);
  }
}

// ===========================
// Example 6: App Installation Status
// ===========================

/**
 * Retrieves and analyzes app installation status
 */
export async function example6_AppInstallationStatus(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 6: App Installation Status Analysis');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const installStatus = await report.getAppInstallationStatus();

    console.log(`\nTotal installation records: ${installStatus.length}\n`);

    // Count by status
    const byStatus = installStatus.reduce((acc, status) => {
      acc[status.status] = (acc[status.status] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Installation Status Summary:');
    Object.entries(byStatus).forEach(([status, count]) => {
      console.log(`  ${status}: ${count}`);
    });

    // Show failed installations
    const failed = installStatus.filter(s => s.status === AppInstallationStatus.FAILED);

    if (failed.length > 0) {
      console.log(`\n\nFailed Installations (${failed.length} total):\n`);

      failed.slice(0, 15).forEach((status, index) => {
        console.log(`${index + 1}. ${status.appName}`);
        console.log(`   Device: ${status.deviceName}`);
        console.log(`   User: ${status.userPrincipalName}`);
        console.log(`   Error Code: ${status.errorCode || 'N/A'}`);
        console.log(`   OS Version: ${status.osVersion || 'N/A'}`);
        console.log('');
      });

      // Group failures by app
      const failuresByApp = failed.reduce((acc, status) => {
        acc[status.appName] = (acc[status.appName] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

      const sortedFailures = Object.entries(failuresByApp)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);

      console.log('Top 10 Apps with Most Failures:');
      sortedFailures.forEach(([appName, count], index) => {
        console.log(`  ${index + 1}. ${appName}: ${count} failures`);
      });
    }

  } catch (error) {
    console.error('Error fetching installation status:', error);
  }
}

// ===========================
// Example 7: App Protection Policies
// ===========================

/**
 * Retrieves and displays app protection policies
 */
export async function example7_AppProtectionPolicies(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 7: App Protection Policies');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const policies = await report.getAppProtectionPolicies();

    console.log(`\nFound ${policies.length} app protection policies\n`);

    // Group by platform
    const byPlatform = {
      iOS: policies.filter(p => p.platform === 'iOS'),
      Android: policies.filter(p => p.platform === 'android'),
      Windows: policies.filter(p => p.platform === 'windows')
    };

    console.log('iOS Protection Policies:\n');
    byPlatform.iOS.forEach((policy, index) => {
      console.log(`${index + 1}. ${policy.displayName}`);
      console.log(`   Description: ${policy.description || 'N/A'}`);
      console.log(`   Apps: ${policy.appCount}`);
      console.log(`   Assignments: ${policy.assignmentCount}`);
      console.log(`   Created: ${policy.createdDateTime.toLocaleDateString()}`);
      console.log(`   Last Modified: ${policy.lastModifiedDateTime.toLocaleDateString()}`);
      console.log('');
    });

    console.log('\nAndroid Protection Policies:\n');
    byPlatform.Android.forEach((policy, index) => {
      console.log(`${index + 1}. ${policy.displayName}`);
      console.log(`   Description: ${policy.description || 'N/A'}`);
      console.log(`   Apps: ${policy.appCount}`);
      console.log(`   Assignments: ${policy.assignmentCount}`);
      console.log(`   Created: ${policy.createdDateTime.toLocaleDateString()}`);
      console.log('');
    });

    if (byPlatform.Windows.length > 0) {
      console.log('\nWindows Information Protection Policies:\n');
      byPlatform.Windows.forEach((policy, index) => {
        console.log(`${index + 1}. ${policy.displayName}`);
        console.log(`   Description: ${policy.description || 'N/A'}`);
        console.log(`   Assignments: ${policy.assignmentCount}`);
        console.log('');
      });
    }

  } catch (error) {
    console.error('Error fetching app protection policies:', error);
  }
}

// ===========================
// Example 8: Protection Policy Status
// ===========================

/**
 * Retrieves and analyzes app protection policy compliance status
 */
export async function example8_ProtectionPolicyStatus(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 8: App Protection Policy Compliance Status');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const statuses = await report.getAppProtectionPolicyStatus();

    console.log(`\nTotal protection status records: ${statuses.length}\n`);

    // Count by status
    const byStatus = statuses.reduce((acc, status) => {
      acc[status.status] = (acc[status.status] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Protection Status Summary:');
    Object.entries(byStatus).forEach(([status, count]) => {
      console.log(`  ${status}: ${count}`);
    });

    // Show non-compliant users
    const nonCompliant = statuses.filter(s => s.status === AppProtectionStatus.NON_COMPLIANT);

    if (nonCompliant.length > 0) {
      console.log(`\n\nNon-Compliant Users (${nonCompliant.length} total):\n`);

      nonCompliant.slice(0, 20).forEach((status, index) => {
        console.log(`${index + 1}. ${status.userPrincipalName}`);
        console.log(`   Policy: ${status.policyName}`);
        console.log(`   Platform: ${status.platform}`);
        if (status.lastCheckInDateTime) {
          console.log(`   Last Check-in: ${status.lastCheckInDateTime.toLocaleString()}`);
        }
        console.log('');
      });

      // Group by policy
      const byPolicy = nonCompliant.reduce((acc, status) => {
        acc[status.policyName] = (acc[status.policyName] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

      console.log('Non-Compliance by Policy:');
      Object.entries(byPolicy).forEach(([policy, count]) => {
        console.log(`  ${policy}: ${count} non-compliant users`);
      });
    }

  } catch (error) {
    console.error('Error fetching protection policy status:', error);
  }
}

// ===========================
// Example 9: App Configuration Policies
// ===========================

/**
 * Retrieves and displays app configuration policies
 */
export async function example9_AppConfigurationPolicies(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 9: App Configuration Policies');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const policies = await report.getAppConfigurationPolicies();

    console.log(`\nFound ${policies.length} app configuration policies\n`);

    policies.forEach((policy, index) => {
      console.log(`${index + 1}. ${policy.displayName}`);
      console.log(`   Description: ${policy.description || 'N/A'}`);
      console.log(`   Targeted Apps: ${policy.targetedMobileApps}`);
      console.log(`   Assignments: ${policy.assignmentCount}`);
      console.log(`   Has Payload: ${policy.hasPayload ? 'Yes' : 'No'}`);
      console.log(`   Created: ${policy.createdDateTime.toLocaleDateString()}`);
      console.log(`   Last Modified: ${policy.lastModifiedDateTime.toLocaleDateString()}`);
      console.log('');
    });

    // Summary stats
    const totalTargetedApps = policies.reduce((sum, p) => sum + p.targetedMobileApps, 0);
    const totalAssignments = policies.reduce((sum, p) => sum + p.assignmentCount, 0);
    const withPayload = policies.filter(p => p.hasPayload).length;

    console.log('Configuration Summary:');
    console.log(`  Total Policies: ${policies.length}`);
    console.log(`  Total Targeted Apps: ${totalTargetedApps}`);
    console.log(`  Total Assignments: ${totalAssignments}`);
    console.log(`  Policies with Payload: ${withPayload}`);

  } catch (error) {
    console.error('Error fetching app configuration policies:', error);
  }
}

// ===========================
// Example 10: Configuration Status
// ===========================

/**
 * Retrieves and analyzes app configuration status
 */
export async function example10_ConfigurationStatus(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 10: App Configuration Status Analysis');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const statuses = await report.getAppConfigurationStatus();

    console.log(`\nTotal configuration status records: ${statuses.length}\n`);

    // Count by status
    const byStatus = statuses.reduce((acc, status) => {
      acc[status.status] = (acc[status.status] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('Configuration Status Summary:');
    Object.entries(byStatus).forEach(([status, count]) => {
      console.log(`  ${status}: ${count}`);
    });

    // Show errors
    const errors = statuses.filter(s => s.status === 'error');

    if (errors.length > 0) {
      console.log(`\n\nConfiguration Errors (${errors.length} total):\n`);

      errors.slice(0, 20).forEach((status, index) => {
        console.log(`${index + 1}. ${status.policyName}`);
        console.log(`   Device: ${status.deviceName}`);
        console.log(`   User: ${status.userPrincipalName}`);
        console.log(`   Platform: ${status.platform}`);
        console.log(`   Error Code: ${status.errorCode || 'N/A'}`);
        if (status.lastReportedDateTime) {
          console.log(`   Last Reported: ${status.lastReportedDateTime.toLocaleString()}`);
        }
        console.log('');
      });

      // Group errors by policy
      const errorsByPolicy = errors.reduce((acc, status) => {
        acc[status.policyName] = (acc[status.policyName] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

      console.log('Errors by Policy:');
      Object.entries(errorsByPolicy)
        .sort((a, b) => b[1] - a[1])
        .forEach(([policy, count]) => {
          console.log(`  ${policy}: ${count} errors`);
        });
    }

    // Show pending configurations
    const pending = statuses.filter(s => s.status === 'pending');
    if (pending.length > 0) {
      console.log(`\n\nPending Configurations: ${pending.length}`);
    }

  } catch (error) {
    console.error('Error fetching configuration status:', error);
  }
}

// ===========================
// Example 11: Export Report Data
// ===========================

/**
 * Generates a report and shows how to export the data
 */
export async function example11_ExportReportData(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 11: Export Report Data');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  try {
    const result = await report.execute();

    // Export to JSON
    const jsonData = JSON.stringify(result, null, 2);
    console.log(`\nJSON Export: ${jsonData.length} characters`);

    // In a real application, you would write to a file:
    // import fs from 'fs';
    // fs.writeFileSync('application-inventory.json', jsonData);

    // Show data structure
    console.log('\nReport Structure:');
    console.log(`- Metadata: ${Object.keys(result.metadata).length} fields`);
    console.log(`- Data Records: ${result.data.length}`);
    console.log(`- Summary Sections: ${Object.keys(result.summary).length}`);

    // Show record types
    const recordTypes = result.data.reduce((acc, record: any) => {
      acc[record.type] = (acc[record.type] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);

    console.log('\nRecord Types:');
    Object.entries(recordTypes).forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });

    // You can also export specific sections
    const managedApps = result.data.filter((r: any) => r.type === 'managed-app');
    console.log(`\nManaged Apps can be exported separately: ${managedApps.length} records`);

  } catch (error) {
    console.error('Error exporting report:', error);
  }
}

// ===========================
// Example 12: Real-time Monitoring
// ===========================

/**
 * Demonstrates continuous monitoring of app inventory
 */
export async function example12_RealTimeMonitoring(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('Example 12: Real-time Application Monitoring');
  console.log('='.repeat(70));

  const report = new ApplicationInventoryReport(graphClient, config);

  let previousFailureCount = 0;
  let previousNonCompliantCount = 0;

  // In a real application, this would run continuously
  // For demo purposes, we'll run it once
  try {
    console.log('\nMonitoring application inventory...\n');

    const [installStatus, protectionStatus] = await Promise.all([
      report.getAppInstallationStatus(),
      report.getAppProtectionPolicyStatus()
    ]);

    // Check for new failures
    const currentFailures = installStatus.filter(
      s => s.status === AppInstallationStatus.FAILED
    );

    console.log(`Installation Failures: ${currentFailures.length}`);
    if (currentFailures.length > previousFailureCount) {
      console.log(`  ⚠️  New failures detected: ${currentFailures.length - previousFailureCount}`);
    }

    // Check for non-compliant devices
    const currentNonCompliant = protectionStatus.filter(
      s => s.status === AppProtectionStatus.NON_COMPLIANT
    );

    console.log(`Non-Compliant Devices: ${currentNonCompliant.length}`);
    if (currentNonCompliant.length > previousNonCompliantCount) {
      console.log(`  ⚠️  New non-compliant devices: ${currentNonCompliant.length - previousNonCompliantCount}`);
    }

    // Get managed apps with high failure rates
    const managedApps = await report.getManagedApps();
    const problematicApps = managedApps.filter(app => {
      if (!app.installSummary) return false;
      const total = app.installSummary.installedDeviceCount + app.installSummary.failedDeviceCount;
      if (total === 0) return false;
      const failureRate = (app.installSummary.failedDeviceCount / total) * 100;
      return failureRate > 20; // More than 20% failure rate
    });

    if (problematicApps.length > 0) {
      console.log(`\n⚠️  Apps with high failure rates (>20%):`);
      problematicApps.forEach(app => {
        const total = app.installSummary!.installedDeviceCount + app.installSummary!.failedDeviceCount;
        const failureRate = ((app.installSummary!.failedDeviceCount / total) * 100).toFixed(1);
        console.log(`  - ${app.displayName}: ${failureRate}% failure rate`);
      });
    }

    console.log('\nMonitoring complete. Run this periodically for continuous monitoring.');

  } catch (error) {
    console.error('Error during monitoring:', error);
  }
}

// ===========================
// Main Function
// ===========================

/**
 * Runs all examples
 */
export async function runAllExamples(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('APPLICATION INVENTORY REPORT - ALL EXAMPLES');
  console.log('='.repeat(70) + '\n');

  const examples = [
    { name: 'Complete Inventory Report', fn: example1_CompleteInventoryReport },
    { name: 'Managed Apps Inventory', fn: example2_ManagedAppsInventory },
    { name: 'Filter iOS Apps', fn: example3_FilterIOSApps },
    { name: 'Discovered Apps', fn: example4_DiscoveredApps },
    { name: 'Find Specific Apps', fn: example5_FindSpecificApps },
    { name: 'App Installation Status', fn: example6_AppInstallationStatus },
    { name: 'App Protection Policies', fn: example7_AppProtectionPolicies },
    { name: 'Protection Policy Status', fn: example8_ProtectionPolicyStatus },
    { name: 'App Configuration Policies', fn: example9_AppConfigurationPolicies },
    { name: 'Configuration Status', fn: example10_ConfigurationStatus },
    { name: 'Export Report Data', fn: example11_ExportReportData },
    { name: 'Real-time Monitoring', fn: example12_RealTimeMonitoring }
  ];

  for (let i = 0; i < examples.length; i++) {
    const example = examples[i];
    console.log(`\nRunning Example ${i + 1}/${examples.length}: ${example.name}...`);

    try {
      await example.fn(graphClient, config);
      console.log(`✓ Example ${i + 1} completed\n`);
    } catch (error) {
      console.error(`✗ Example ${i + 1} failed:`, error);
    }

    // Add delay between examples to avoid rate limiting
    if (i < examples.length - 1) {
      console.log('Waiting 2 seconds before next example...');
      await new Promise(resolve => setTimeout(resolve, 2000));
    }
  }

  console.log('\n' + '='.repeat(70));
  console.log('ALL EXAMPLES COMPLETED');
  console.log('='.repeat(70) + '\n');
}
