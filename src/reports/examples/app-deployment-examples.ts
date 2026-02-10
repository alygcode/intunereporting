/**
 * Application Deployment Report - Usage Examples
 *
 * This file demonstrates various usage scenarios for the Application Deployment Report module.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { AppDeploymentReport, AppPlatform, InstallState } from '../app-deployment-report';
import { AppConfig } from '../../types';

/**
 * Example 1: Basic usage - Get complete deployment report
 */
export async function example1_CompleteDeploymentReport(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 1: Complete Deployment Report ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    // Execute complete report
    const reportData = await report.execute();

    console.log('Report Metadata:');
    console.log(`  Report Name: ${reportData.metadata.reportName}`);
    console.log(`  Generated At: ${reportData.metadata.generatedAt}`);
    console.log(`  Total Records: ${reportData.metadata.recordCount}`);
    console.log(`  Duration: ${reportData.metadata.parameters?.duration}`);

    console.log('\nDeployment Statistics:');
    const stats = reportData.summary?.statistics;
    if (stats) {
      console.log(`  Total Apps: ${stats.totalApps}`);
      console.log(`  Total Devices: ${stats.totalDevices}`);
      console.log(`  Total Installations: ${stats.totalInstallations}`);
      console.log(`  Success Rate: ${stats.successRate}%`);
      console.log(`  Failure Rate: ${stats.failureRate}%`);

      console.log('\n  Apps by Platform:');
      Object.entries(stats.appsByPlatform).forEach(([platform, count]) => {
        if (count > 0) {
          console.log(`    ${platform}: ${count}`);
        }
      });

      if (stats.topFailedApps.length > 0) {
        console.log('\n  Top Failed Apps:');
        stats.topFailedApps.forEach((app, index) => {
          console.log(`    ${index + 1}. ${app.appName} (${app.failureCount} failures)`);
        });
      }
    }

    console.log('\n✓ Complete deployment report generated successfully!\n');
  } catch (error) {
    console.error('Error generating deployment report:', error);
  }
}

/**
 * Example 2: Get all managed apps with filtering
 */
export async function example2_GetManagedApps(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 2: Get Managed Apps ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    // Get all apps
    console.log('Fetching all managed apps...');
    const allApps = await report.getAllManagedApps();
    console.log(`Total Apps: ${allApps.length}\n`);

    // Get iOS apps only
    console.log('Fetching iOS apps...');
    const iosApps = await report.getAllManagedApps({ platform: AppPlatform.IOS });
    console.log(`iOS Apps: ${iosApps.length}`);
    iosApps.slice(0, 5).forEach(app => {
      console.log(`  - ${app.displayName} (${app.publisher})`);
    });

    // Get Android apps only
    console.log('\nFetching Android apps...');
    const androidApps = await report.getAllManagedApps({ platform: AppPlatform.ANDROID });
    console.log(`Android Apps: ${androidApps.length}`);
    androidApps.slice(0, 5).forEach(app => {
      console.log(`  - ${app.displayName} (${app.publisher})`);
    });

    // Get Windows apps only
    console.log('\nFetching Windows apps...');
    const windowsApps = await report.getAllManagedApps({ platform: AppPlatform.WINDOWS });
    console.log(`Windows Apps: ${windowsApps.length}`);
    windowsApps.slice(0, 5).forEach(app => {
      console.log(`  - ${app.displayName} (${app.publisher})`);
    });

    // Filter by publisher
    console.log('\nFetching Microsoft apps...');
    const microsoftApps = await report.getAllManagedApps({ publisher: 'Microsoft' });
    console.log(`Microsoft Apps: ${microsoftApps.length}`);
    microsoftApps.slice(0, 5).forEach(app => {
      console.log(`  - ${app.displayName}`);
    });

    console.log('\n✓ Managed apps retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching managed apps:', error);
  }
}

/**
 * Example 3: Get installation summaries for all apps
 */
export async function example3_GetInstallationSummaries(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 3: Installation Summaries ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    console.log('Fetching installation summaries...');
    const summaries = await report.getAllAppInstallSummaries();

    console.log(`\nInstallation Summary for ${summaries.length} apps:\n`);

    // Display summary for each app
    summaries.forEach((summary, index) => {
      if (index < 10) { // Show first 10
        console.log(`${index + 1}. ${summary.appName}`);
        console.log(`   Installed Devices: ${summary.installedDeviceCount}`);
        console.log(`   Failed Devices: ${summary.failedDeviceCount}`);
        console.log(`   Pending Devices: ${summary.pendingInstallDeviceCount}`);
        console.log(`   Not Installed Devices: ${summary.notInstalledDeviceCount}`);

        const total = summary.installedDeviceCount + summary.failedDeviceCount +
                     summary.pendingInstallDeviceCount;
        const successRate = total > 0
          ? Math.round((summary.installedDeviceCount / total) * 100)
          : 0;
        console.log(`   Success Rate: ${successRate}%\n`);
      }
    });

    // Overall statistics
    const totalInstalled = summaries.reduce((sum, s) => sum + s.installedDeviceCount, 0);
    const totalFailed = summaries.reduce((sum, s) => sum + s.failedDeviceCount, 0);
    const totalPending = summaries.reduce((sum, s) => sum + s.pendingInstallDeviceCount, 0);

    console.log('Overall Statistics:');
    console.log(`  Total Installed: ${totalInstalled}`);
    console.log(`  Total Failed: ${totalFailed}`);
    console.log(`  Total Pending: ${totalPending}`);

    console.log('\n✓ Installation summaries retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching installation summaries:', error);
  }
}

/**
 * Example 4: Get device-level deployment status
 */
export async function example4_GetDeviceStatuses(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 4: Device-level Deployment Status ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    console.log('Fetching device statuses (this may take a while)...');
    const deviceStatuses = await report.getAllAppDeviceStatuses({ maxResults: 20 });

    console.log(`\nRetrieved ${deviceStatuses.length} device status records\n`);

    // Group by install state
    const byState = new Map<InstallState, number>();
    deviceStatuses.forEach(status => {
      byState.set(status.installState, (byState.get(status.installState) || 0) + 1);
    });

    console.log('Status Breakdown:');
    byState.forEach((count, state) => {
      console.log(`  ${state}: ${count}`);
    });

    // Show sample device statuses
    console.log('\nSample Device Statuses:');
    deviceStatuses.slice(0, 5).forEach((status, index) => {
      console.log(`\n${index + 1}. ${status.appName}`);
      console.log(`   Device: ${status.deviceName}`);
      console.log(`   User: ${status.userPrincipalName || 'N/A'}`);
      console.log(`   Platform: ${status.platform || 'N/A'}`);
      console.log(`   OS Version: ${status.osVersion || 'N/A'}`);
      console.log(`   Install State: ${status.installState}`);
      if (status.errorCode) {
        console.log(`   Error Code: ${status.errorCode}`);
      }
    });

    console.log('\n✓ Device statuses retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching device statuses:', error);
  }
}

/**
 * Example 5: Get failed installations with error details
 */
export async function example5_GetFailedInstallations(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 5: Failed Installations ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    console.log('Analyzing failed installations...');
    const failures = await report.getFailedInstallations({ maxResults: 20 });

    console.log(`\nFound ${failures.length} failed installations\n`);

    if (failures.length === 0) {
      console.log('No failed installations found!');
      return;
    }

    // Group by app
    const failuresByApp = new Map<string, number>();
    failures.forEach(failure => {
      failuresByApp.set(failure.appName, (failuresByApp.get(failure.appName) || 0) + 1);
    });

    console.log('Failures by App:');
    Array.from(failuresByApp.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .forEach(([appName, count], index) => {
        console.log(`  ${index + 1}. ${appName}: ${count} failures`);
      });

    // Show detailed failure information
    console.log('\nDetailed Failure Information:');
    failures.slice(0, 5).forEach((failure, index) => {
      console.log(`\n${index + 1}. ${failure.appName}`);
      console.log(`   Device: ${failure.deviceName}`);
      console.log(`   User: ${failure.userPrincipalName || 'N/A'}`);
      console.log(`   Platform: ${failure.platform || 'N/A'}`);
      console.log(`   OS Version: ${failure.osVersion || 'N/A'}`);
      console.log(`   Error Code: ${failure.errorCode || 'N/A'}`);
      console.log(`   Error Description: ${failure.errorDescription || 'N/A'}`);
      console.log(`   Failure Reason: ${failure.failureReason || 'N/A'}`);
      console.log(`   Troubleshooting: ${failure.troubleshootingLink || 'N/A'}`);
      console.log(`   Last Sync: ${failure.lastSyncDateTime?.toLocaleString() || 'N/A'}`);
    });

    console.log('\n✓ Failed installations analyzed successfully!\n');
  } catch (error) {
    console.error('Error analyzing failed installations:', error);
  }
}

/**
 * Example 6: Get app update compliance
 */
export async function example6_GetUpdateCompliance(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 6: App Update Compliance ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    console.log('Analyzing app update compliance...');
    const compliance = await report.getAppUpdateCompliance();

    console.log(`\nUpdate compliance for ${compliance.length} apps\n`);

    // Show apps requiring updates
    const appsRequiringUpdates = compliance.filter(c => c.devicesRequiringUpdate > 0);

    if (appsRequiringUpdates.length > 0) {
      console.log('Apps with Devices Requiring Updates:');
      appsRequiringUpdates
        .sort((a, b) => b.devicesRequiringUpdate - a.devicesRequiringUpdate)
        .slice(0, 10)
        .forEach((app, index) => {
          console.log(`\n${index + 1}. ${app.appName}`);
          console.log(`   Current Version: ${app.currentVersion}`);
          console.log(`   Devices on Current Version: ${app.devicesOnCurrentVersion}`);
          console.log(`   Devices on Older Versions: ${app.devicesOnOlderVersion}`);
          console.log(`   Devices Requiring Update: ${app.devicesRequiringUpdate}`);
          console.log(`   Update Compliance: ${app.updateCompliancePercentage}%`);
        });
    } else {
      console.log('All apps are up to date!');
    }

    // Overall compliance statistics
    const totalDevices = compliance.reduce((sum, c) =>
      sum + c.devicesOnCurrentVersion + c.devicesOnOlderVersion, 0);
    const devicesUpToDate = compliance.reduce((sum, c) => sum + c.devicesOnCurrentVersion, 0);
    const overallCompliance = totalDevices > 0
      ? Math.round((devicesUpToDate / totalDevices) * 100)
      : 100;

    console.log('\n\nOverall Update Compliance:');
    console.log(`  Total Devices: ${totalDevices}`);
    console.log(`  Devices Up to Date: ${devicesUpToDate}`);
    console.log(`  Overall Compliance: ${overallCompliance}%`);

    console.log('\n✓ Update compliance analyzed successfully!\n');
  } catch (error) {
    console.error('Error analyzing update compliance:', error);
  }
}

/**
 * Example 7: Export deployment report to different formats
 */
export async function example7_ExportReports(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 7: Export Deployment Reports ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);
    const outputDir = './reports/output';

    // Export complete report as JSON
    console.log('Exporting complete report as JSON...');
    const jsonPath = await report.exportDeploymentReport({
      format: 'json',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ JSON report saved to: ${jsonPath}`);

    // Export complete report as CSV
    console.log('\nExporting complete report as CSV...');
    const csvPath = await report.exportDeploymentReport({
      format: 'csv',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ CSV report saved to: ${csvPath}`);

    // Export complete report as HTML
    console.log('\nExporting complete report as HTML...');
    const htmlPath = await report.exportDeploymentReport({
      format: 'html',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ HTML report saved to: ${htmlPath}`);

    // Export only failed installations
    console.log('\nExporting failed installations...');
    const failuresPath = await report.exportFailedInstallations({
      format: 'csv',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ Failed installations report saved to: ${failuresPath}`);

    // Export only update compliance
    console.log('\nExporting update compliance...');
    const compliancePath = await report.exportUpdateCompliance({
      format: 'csv',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ Update compliance report saved to: ${compliancePath}`);

    // Export installation summaries
    console.log('\nExporting installation summaries...');
    const summariesPath = await report.exportInstallSummaries({
      format: 'csv',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ Installation summaries report saved to: ${summariesPath}`);

    // Export app assignments
    console.log('\nExporting app assignments...');
    const assignmentsPath = await report.exportAppAssignments({
      format: 'csv',
      outputDir,
      includeTimestamp: true
    });
    console.log(`✓ App assignments report saved to: ${assignmentsPath}`);

    console.log('\n✓ All reports exported successfully!\n');
  } catch (error) {
    console.error('Error exporting reports:', error);
  }
}

/**
 * Example 8: Get app assignment details
 */
export async function example8_GetAppAssignments(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 8: App Assignment Details ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    console.log('Fetching app assignments...');
    const assignments = await report.getAllAppAssignments({ maxResults: 50 });

    console.log(`\nFound ${assignments.length} app assignments\n`);

    // Group by intent
    const byIntent = new Map<string, number>();
    assignments.forEach(assignment => {
      byIntent.set(assignment.intent, (byIntent.get(assignment.intent) || 0) + 1);
    });

    console.log('Assignments by Intent:');
    byIntent.forEach((count, intent) => {
      console.log(`  ${intent}: ${count}`);
    });

    // Group by target type
    const byTargetType = new Map<string, number>();
    assignments.forEach(assignment => {
      byTargetType.set(assignment.targetType, (byTargetType.get(assignment.targetType) || 0) + 1);
    });

    console.log('\nAssignments by Target Type:');
    byTargetType.forEach((count, targetType) => {
      console.log(`  ${targetType}: ${count}`);
    });

    // Show sample assignments
    console.log('\nSample Assignments:');
    assignments.slice(0, 5).forEach((assignment, index) => {
      console.log(`\n${index + 1}. ${assignment.appName}`);
      console.log(`   Intent: ${assignment.intent}`);
      console.log(`   Target Type: ${assignment.targetType}`);
      console.log(`   Target Name: ${assignment.targetName}`);
      console.log(`   Include/Exclude: ${assignment.includeExclude}`);
    });

    // Count include vs exclude
    const includeCount = assignments.filter(a => a.includeExclude === 'include').length;
    const excludeCount = assignments.filter(a => a.includeExclude === 'exclude').length;

    console.log('\n\nAssignment Types:');
    console.log(`  Include Assignments: ${includeCount}`);
    console.log(`  Exclude Assignments: ${excludeCount}`);

    console.log('\n✓ App assignments retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching app assignments:', error);
  }
}

/**
 * Example 9: Get deployment details for a specific app
 */
export async function example9_SpecificAppDeployment(
  graphClient: Client,
  config: AppConfig,
  appId: string
): Promise<void> {
  console.log('=== Example 9: Specific App Deployment ===\n');

  try {
    const report = new AppDeploymentReport(graphClient, config);

    // Get app details
    console.log('Fetching app details...');
    const app = await report.getManagedAppById(appId);
    console.log(`\nApp: ${app.displayName}`);
    console.log(`Publisher: ${app.publisher}`);
    console.log(`Platform: ${app.platform}`);
    console.log(`Version: ${app.version || 'N/A'}`);

    // Get installation summary
    console.log('\nFetching installation summary...');
    const summary = await report.getAppInstallSummary(appId, app.displayName);
    if (summary) {
      console.log(`Installed Devices: ${summary.installedDeviceCount}`);
      console.log(`Failed Devices: ${summary.failedDeviceCount}`);
      console.log(`Pending Devices: ${summary.pendingInstallDeviceCount}`);
    }

    // Get app assignments
    console.log('\nFetching app assignments...');
    const assignments = await report.getAppAssignments(appId, app.displayName);
    console.log(`Total assignments: ${assignments.length}`);

    if (assignments.length > 0) {
      console.log('\nAssignment Details:');
      assignments.forEach((assignment, index) => {
        console.log(`\n${index + 1}. Intent: ${assignment.intent}`);
        console.log(`   Target: ${assignment.targetName} (${assignment.targetType})`);
        console.log(`   Type: ${assignment.includeExclude}`);
      });

      // Get assignment statistics
      const stats = await report.getAppAssignmentStats(appId);
      console.log('\nAssignment Statistics:');
      console.log(`  Total Assignments: ${stats.totalAssignments}`);
      console.log(`  Required: ${stats.requiredAssignments}`);
      console.log(`  Available: ${stats.availableAssignments}`);
      console.log(`  Uninstall: ${stats.uninstallAssignments}`);
      console.log(`  Group Assignments: ${stats.groupAssignments}`);
      console.log(`  All Users: ${stats.allUsersAssignments}`);
      console.log(`  All Devices: ${stats.allDevicesAssignments}`);
    }

    // Get device statuses
    console.log('\nFetching device statuses...');
    const deviceStatuses = await report.getAppDeviceStatuses(appId, app.displayName);
    console.log(`Total device records: ${deviceStatuses.length}`);

    // Get failed installations
    console.log('\nFetching failed installations...');
    const failures = await report.getFailedInstallationsByApp(appId, app.displayName);
    console.log(`Failed installations: ${failures.length}`);

    if (failures.length > 0) {
      console.log('\nFailed Installation Details:');
      failures.slice(0, 5).forEach((failure, index) => {
        console.log(`\n${index + 1}. Device: ${failure.deviceName}`);
        console.log(`   User: ${failure.userPrincipalName || 'N/A'}`);
        console.log(`   Error: ${failure.errorDescription || 'N/A'}`);
      });
    }

    // Get update compliance
    console.log('\nFetching update compliance...');
    const compliance = await report.getAppUpdateComplianceById(appId);
    if (compliance) {
      console.log(`Current Version: ${compliance.currentVersion}`);
      console.log(`Devices on Current Version: ${compliance.devicesOnCurrentVersion}`);
      console.log(`Devices on Older Versions: ${compliance.devicesOnOlderVersion}`);
      console.log(`Update Compliance: ${compliance.updateCompliancePercentage}%`);
    }

    console.log('\n✓ App deployment details retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching app deployment details:', error);
  }
}

/**
 * Example 10: Run all examples
 */
export async function runAllExamples(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('\n' + '='.repeat(60));
  console.log('APPLICATION DEPLOYMENT REPORT - ALL EXAMPLES');
  console.log('='.repeat(60) + '\n');

  await example1_CompleteDeploymentReport(graphClient, config);
  await example2_GetManagedApps(graphClient, config);
  await example3_GetInstallationSummaries(graphClient, config);
  await example4_GetDeviceStatuses(graphClient, config);
  await example5_GetFailedInstallations(graphClient, config);
  await example6_GetUpdateCompliance(graphClient, config);
  await example7_ExportReports(graphClient, config);
  await example8_GetAppAssignments(graphClient, config);

  console.log('\n' + '='.repeat(60));
  console.log('ALL EXAMPLES COMPLETED SUCCESSFULLY');
  console.log('='.repeat(60) + '\n');
}

// Export all examples
export default {
  example1_CompleteDeploymentReport,
  example2_GetManagedApps,
  example3_GetInstallationSummaries,
  example4_GetDeviceStatuses,
  example5_GetFailedInstallations,
  example6_GetUpdateCompliance,
  example7_ExportReports,
  example8_GetAppAssignments,
  example9_SpecificAppDeployment,
  runAllExamples
};
