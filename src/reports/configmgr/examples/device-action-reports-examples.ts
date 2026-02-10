/**
 * Usage Examples for Device Action Reports
 *
 * This file demonstrates various ways to use the Device Action Reports module
 * to generate Configuration Manager equivalent reports for Intune device actions.
 *
 * @module DeviceActionReportsExamples
 */

import {
  DeviceActionReports,
  createDeviceActionReports,
  generateQuickReport,
  DeviceActionReport,
  PendingDeviceAction,
  RecentlyEnrolledDevice,
  RecentlyWipedDevice,
  DeviceActionReportOptions,
} from '../device-action-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import * as dotenv from 'dotenv';
import * as fs from 'fs/promises';
import * as path from 'path';

// Load environment variables
dotenv.config();

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Create authenticated Graph client
 */
function createGraphClient(): Client {
  const credential = new ClientSecretCredential(
    process.env.AZURE_TENANT_ID || '',
    process.env.AZURE_CLIENT_ID || '',
    process.env.AZURE_CLIENT_SECRET || ''
  );

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });

  return Client.initWithMiddleware({
    authProvider,
  });
}

/**
 * Save report to file
 */
async function saveReport(
  filename: string,
  content: string,
  outputDir: string = './reports'
): Promise<string> {
  await fs.mkdir(outputDir, { recursive: true });
  const filePath = path.join(outputDir, filename);
  await fs.writeFile(filePath, content, 'utf-8');
  console.log(`Report saved to: ${filePath}`);
  return filePath;
}

// ============================================================================
// Example 1: Basic Usage - Get Pending Retire and Wipe Requests
// ============================================================================

async function example1PendingActions(): Promise<void> {
  console.log('\n=== Example 1: Get Pending Retire and Wipe Requests ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Get all pending retire and wipe requests
    const pendingActions = await reports.getPendingRetireWipeRequests();

    console.log(`Total pending actions: ${pendingActions.length}`);

    // Display first 5 pending actions
    if (pendingActions.length > 0) {
      console.log('\nFirst 5 pending actions:');
      pendingActions.slice(0, 5).forEach((action, index) => {
        console.log(`\n${index + 1}. ${action.deviceName}`);
        console.log(`   User: ${action.userPrincipalName}`);
        console.log(`   Action: ${action.actionType}`);
        console.log(`   Status: ${action.actionStatus}`);
        console.log(`   Requested: ${action.actionRequestedDateTime}`);
        console.log(`   OS: ${action.operatingSystem} ${action.osVersion}`);
      });
    }

    // Export to CSV
    const csv = reports.exportPendingActionsToCSV(pendingActions);
    await saveReport('pending-actions.csv', csv);
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 2: Recently Enrolled Devices (Last 7 Days)
// ============================================================================

async function example2RecentEnrollments(): Promise<void> {
  console.log('\n=== Example 2: Recently Enrolled Devices (Last 7 Days) ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Get devices enrolled in last 7 days (default)
    const enrolledDevices = await reports.getRecentlyEnrolledDevices();

    console.log(`Total recently enrolled devices: ${enrolledDevices.length}`);

    // Group by operating system
    const byOS: Record<string, number> = {};
    enrolledDevices.forEach((device) => {
      byOS[device.operatingSystem] = (byOS[device.operatingSystem] || 0) + 1;
    });

    console.log('\nEnrollments by Operating System:');
    Object.entries(byOS)
      .sort((a, b) => b[1] - a[1])
      .forEach(([os, count]) => {
        console.log(`  ${os}: ${count} devices`);
      });

    // Show assignment status
    const assignedCount = enrolledDevices.filter((d) => d.isAssigned).length;
    console.log(`\nAssignment Status:`);
    console.log(`  Assigned: ${assignedCount}`);
    console.log(`  Not Assigned: ${enrolledDevices.length - assignedCount}`);

    // Display recently enrolled devices
    if (enrolledDevices.length > 0) {
      console.log('\nRecently enrolled devices:');
      enrolledDevices.slice(0, 5).forEach((device, index) => {
        console.log(`\n${index + 1}. ${device.deviceName}`);
        console.log(`   User: ${device.userPrincipalName}`);
        console.log(`   Model: ${device.manufacturer} ${device.model}`);
        console.log(`   Enrolled: ${device.enrolledDateTime}`);
        console.log(`   Days Since Enrollment: ${device.daysSinceEnrollment}`);
        console.log(`   Compliance: ${device.complianceState}`);
        console.log(`   Assigned: ${device.isAssigned ? 'Yes' : 'No'}`);
      });
    }

    // Export to CSV
    const csv = reports.exportEnrollmentsToCSV(enrolledDevices);
    await saveReport('recent-enrollments-7days.csv', csv);
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 3: Custom Date Range - Last 30 Days
// ============================================================================

async function example3CustomDateRange(): Promise<void> {
  console.log('\n=== Example 3: Custom Date Range (Last 30 Days) ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Define custom date range (last 30 days)
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 30);

    const options: DeviceActionReportOptions = {
      dateRange: { startDate, endDate },
      top: 100, // Limit to 100 devices
    };

    // Get enrollments for the custom date range
    const enrolledDevices = await reports.getRecentlyEnrolledDevices(options);

    console.log(`Devices enrolled in last 30 days: ${enrolledDevices.length}`);
    console.log(`Date range: ${startDate.toISOString()} to ${endDate.toISOString()}`);

    // Calculate average days since enrollment
    const avgDays =
      enrolledDevices.reduce((sum, d) => sum + d.daysSinceEnrollment, 0) /
      (enrolledDevices.length || 1);

    console.log(`\nAverage days since enrollment: ${avgDays.toFixed(1)}`);
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 4: Recently Wiped Devices
// ============================================================================

async function example4RecentWipes(): Promise<void> {
  console.log('\n=== Example 4: Recently Wiped Devices ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Get wiped devices (default: last 30 days)
    const wipedDevices = await reports.getRecentlyWipedDevices();

    console.log(`Total wiped devices: ${wipedDevices.length}`);

    if (wipedDevices.length > 0) {
      // Calculate statistics
      const successfulWipes = wipedDevices.filter((d) => d.wipeStatus === 'success').length;
      const failedWipes = wipedDevices.filter((d) => d.wipeStatus === 'failed').length;
      const avgDuration =
        wipedDevices.reduce((sum, d) => sum + d.wipeDurationMinutes, 0) /
        wipedDevices.length;

      console.log('\nWipe Statistics:');
      console.log(`  Successful: ${successfulWipes}`);
      console.log(`  Failed: ${failedWipes}`);
      console.log(`  Success Rate: ${((successfulWipes / wipedDevices.length) * 100).toFixed(1)}%`);
      console.log(`  Average Duration: ${avgDuration.toFixed(1)} minutes`);

      // Group by wipe type
      const byType: Record<string, number> = {};
      wipedDevices.forEach((device) => {
        byType[device.wipeType] = (byType[device.wipeType] || 0) + 1;
      });

      console.log('\nWipes by Type:');
      Object.entries(byType).forEach(([type, count]) => {
        console.log(`  ${type}: ${count}`);
      });

      // Display recent wipes
      console.log('\nRecent wipes:');
      wipedDevices.slice(0, 5).forEach((device, index) => {
        console.log(`\n${index + 1}. ${device.deviceName}`);
        console.log(`   User: ${device.userPrincipalName}`);
        console.log(`   OS: ${device.operatingSystem} ${device.osVersion}`);
        console.log(`   Type: ${device.wipeType}`);
        console.log(`   Status: ${device.wipeStatus}`);
        console.log(`   Duration: ${device.wipeDurationMinutes} minutes`);
        console.log(`   Completed: ${device.wipeCompletedDateTime}`);
      });
    }

    // Export to CSV
    const csv = reports.exportWipesToCSV(wipedDevices);
    await saveReport('recent-wipes.csv', csv);
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 5: Comprehensive Report with All Sections
// ============================================================================

async function example5ComprehensiveReport(): Promise<void> {
  console.log('\n=== Example 5: Comprehensive Device Action Report ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Generate comprehensive report with default options (last 30 days)
    const report = await reports.generateDeviceActionReport();

    console.log('=== Device Action Report Summary ===\n');
    console.log(`Report Generated: ${report.metadata.generatedAt}`);
    console.log(`Report Period: ${report.reportPeriod.startDate} to ${report.reportPeriod.endDate}`);
    console.log(`Total Records: ${report.metadata.recordCount}\n`);

    console.log('Summary Statistics:');
    console.log(`  Pending Actions: ${report.summary.totalPendingActions}`);
    console.log(`  Recent Enrollments: ${report.summary.totalRecentEnrollments}`);
    console.log(`  Recent Wipes: ${report.summary.totalRecentWipes}`);
    console.log(`  Wipe Success Rate: ${report.summary.wipeSuccessRate.toFixed(1)}%`);
    console.log(
      `  Avg Wipe Duration: ${report.summary.averageWipeDurationMinutes.toFixed(1)} minutes`
    );

    // Display actions by type
    console.log('\nActions by Type:');
    Object.entries(report.summary.actionsByType).forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });

    // Display enrollments by OS
    console.log('\nEnrollments by Operating System:');
    Object.entries(report.summary.enrollmentsByOS).forEach(([os, count]) => {
      console.log(`  ${os}: ${count}`);
    });

    // Export to different formats
    console.log('\nExporting reports...');

    // Export to JSON
    const json = reports.exportToJSON(report);
    await saveReport('device-action-report.json', json);

    // Export to HTML
    const html = reports.exportToHTML(report);
    await saveReport('device-action-report.html', html);

    console.log('\nAll reports generated successfully!');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 6: Filtered Reports - Specific Operating System
// ============================================================================

async function example6FilterByOS(): Promise<void> {
  console.log('\n=== Example 6: Filter by Operating System ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Filter for iOS devices only
    const options: DeviceActionReportOptions = {
      operatingSystem: ['iOS'],
      dateRange: {
        startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
        endDate: new Date(),
      },
    };

    const enrolledDevices = await reports.getRecentlyEnrolledDevices(options);

    console.log(`iOS devices enrolled in last 7 days: ${enrolledDevices.length}`);

    if (enrolledDevices.length > 0) {
      console.log('\niOS device models:');
      const models = new Map<string, number>();
      enrolledDevices.forEach((device) => {
        const model = `${device.manufacturer} ${device.model}`;
        models.set(model, (models.get(model) || 0) + 1);
      });

      Array.from(models.entries())
        .sort((a, b) => b[1] - a[1])
        .forEach(([model, count]) => {
          console.log(`  ${model}: ${count}`);
        });
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 7: Filter by Ownership Type
// ============================================================================

async function example7FilterByOwnership(): Promise<void> {
  console.log('\n=== Example 7: Filter by Ownership Type ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Filter for corporate devices
    const corporateOptions: DeviceActionReportOptions = {
      ownershipType: ['company'],
      dateRange: {
        startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
        endDate: new Date(),
      },
    };

    const corporateDevices = await reports.getRecentlyEnrolledDevices(corporateOptions);

    console.log(`Corporate devices enrolled in last 30 days: ${corporateDevices.length}`);

    // Filter for personal devices
    const personalOptions: DeviceActionReportOptions = {
      ownershipType: ['personal'],
      dateRange: corporateOptions.dateRange,
    };

    const personalDevices = await reports.getRecentlyEnrolledDevices(personalOptions);

    console.log(`Personal devices enrolled in last 30 days: ${personalDevices.length}`);

    // Calculate percentages
    const total = corporateDevices.length + personalDevices.length;
    if (total > 0) {
      console.log('\nOwnership Distribution:');
      console.log(`  Corporate: ${((corporateDevices.length / total) * 100).toFixed(1)}%`);
      console.log(`  Personal: ${((personalDevices.length / total) * 100).toFixed(1)}%`);
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 8: Quick Report Helper Function
// ============================================================================

async function example8QuickReport(): Promise<void> {
  console.log('\n=== Example 8: Quick Report Helper Function ===\n');

  try {
    const graphClient = createGraphClient();

    // Use the quick helper function for 7-day report
    const report = await generateQuickReport(graphClient, 7);

    console.log('Quick 7-Day Report:');
    console.log(`  Pending Actions: ${report.summary.totalPendingActions}`);
    console.log(`  Recent Enrollments: ${report.summary.totalRecentEnrollments}`);
    console.log(`  Recent Wipes: ${report.summary.totalRecentWipes}`);

    // Export to JSON
    const reports = createDeviceActionReports(graphClient);
    const json = reports.exportToJSON(report);
    await saveReport('quick-report-7days.json', json);
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 9: Monitor Device Lifecycle
// ============================================================================

async function example9DeviceLifecycle(): Promise<void> {
  console.log('\n=== Example 9: Device Lifecycle Monitoring ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Get all three report types
    const [pendingActions, enrollments, wipes] = await Promise.all([
      reports.getPendingRetireWipeRequests(),
      reports.getRecentlyEnrolledDevices({
        dateRange: {
          startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
          endDate: new Date(),
        },
      }),
      reports.getRecentlyWipedDevices({
        dateRange: {
          startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
          endDate: new Date(),
        },
      }),
    ]);

    console.log('=== Device Lifecycle Summary (Last 30 Days) ===\n');
    console.log(`New Enrollments: ${enrollments.length}`);
    console.log(`Pending Actions: ${pendingActions.length}`);
    console.log(`Completed Wipes: ${wipes.length}`);

    // Calculate net change
    const netChange = enrollments.length - wipes.length;
    console.log(`\nNet Device Change: ${netChange > 0 ? '+' : ''}${netChange}`);

    // Analyze trends
    if (enrollments.length > 0) {
      const avgDaysSinceEnrollment =
        enrollments.reduce((sum, d) => sum + d.daysSinceEnrollment, 0) / enrollments.length;
      console.log(`\nAverage enrollment age: ${avgDaysSinceEnrollment.toFixed(1)} days`);
    }

    if (wipes.length > 0) {
      const avgWipeDuration =
        wipes.reduce((sum, d) => sum + d.wipeDurationMinutes, 0) / wipes.length;
      console.log(`Average wipe duration: ${avgWipeDuration.toFixed(1)} minutes`);
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 10: Advanced Filtering and Analysis
// ============================================================================

async function example10AdvancedAnalysis(): Promise<void> {
  console.log('\n=== Example 10: Advanced Filtering and Analysis ===\n');

  try {
    const graphClient = createGraphClient();
    const reports = new DeviceActionReports(graphClient);

    // Generate comprehensive report
    const report = await reports.generateDeviceActionReport({
      dateRange: {
        startDate: new Date(Date.now() - 90 * 24 * 60 * 60 * 1000), // 90 days
        endDate: new Date(),
      },
    });

    console.log('=== Advanced Analysis (Last 90 Days) ===\n');

    // Analyze enrollment patterns by day of week
    const enrollmentsByDay: Record<string, number> = {
      Sunday: 0,
      Monday: 0,
      Tuesday: 0,
      Wednesday: 0,
      Thursday: 0,
      Friday: 0,
      Saturday: 0,
    };

    const dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

    report.recentEnrollments.forEach((device) => {
      const date = new Date(device.enrolledDateTime);
      const day = dayNames[date.getDay()];
      enrollmentsByDay[day]++;
    });

    console.log('Enrollments by Day of Week:');
    Object.entries(enrollmentsByDay)
      .sort((a, b) => b[1] - a[1])
      .forEach(([day, count]) => {
        console.log(`  ${day}: ${count}`);
      });

    // Analyze wipe patterns
    if (report.recentWipes.length > 0) {
      const wipesByMonth: Record<string, number> = {};
      report.recentWipes.forEach((device) => {
        const date = new Date(device.wipeCompletedDateTime);
        const month = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
        wipesByMonth[month] = (wipesByMonth[month] || 0) + 1;
      });

      console.log('\nWipes by Month:');
      Object.entries(wipesByMonth)
        .sort()
        .forEach(([month, count]) => {
          console.log(`  ${month}: ${count}`);
        });
    }

    // Identify potential issues
    console.log('\n=== Potential Issues ===');

    // Pending actions older than 7 days
    const stalePendingActions = report.pendingActions.filter((action) => {
      const requestedDate = new Date(action.actionRequestedDateTime);
      const daysOld = (Date.now() - requestedDate.getTime()) / (1000 * 60 * 60 * 24);
      return daysOld > 7;
    });

    if (stalePendingActions.length > 0) {
      console.log(`\n⚠ Stale pending actions (>7 days): ${stalePendingActions.length}`);
      stalePendingActions.slice(0, 3).forEach((action) => {
        console.log(`  - ${action.deviceName} (${action.actionType})`);
      });
    }

    // Devices enrolled but not synced recently
    const unsyncedDevices = report.recentEnrollments.filter((device) => {
      const lastSync = new Date(device.lastSyncDateTime);
      const daysSinceSync = (Date.now() - lastSync.getTime()) / (1000 * 60 * 60 * 24);
      return daysSinceSync > 7;
    });

    if (unsyncedDevices.length > 0) {
      console.log(`\n⚠ Enrolled devices not synced in 7+ days: ${unsyncedDevices.length}`);
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Main Function
// ============================================================================

async function main(): Promise<void> {
  const examples = [
    { name: 'Pending Retire and Wipe Requests', fn: example1PendingActions },
    { name: 'Recently Enrolled Devices (7 Days)', fn: example2RecentEnrollments },
    { name: 'Custom Date Range (30 Days)', fn: example3CustomDateRange },
    { name: 'Recently Wiped Devices', fn: example4RecentWipes },
    { name: 'Comprehensive Report', fn: example5ComprehensiveReport },
    { name: 'Filter by Operating System', fn: example6FilterByOS },
    { name: 'Filter by Ownership Type', fn: example7FilterByOwnership },
    { name: 'Quick Report Helper', fn: example8QuickReport },
    { name: 'Device Lifecycle Monitoring', fn: example9DeviceLifecycle },
    { name: 'Advanced Analysis', fn: example10AdvancedAnalysis },
  ];

  console.log('=======================================================');
  console.log('   Device Action Reports - Configuration Manager');
  console.log('   Equivalent Reports for Intune');
  console.log('=======================================================\n');

  // Check if specific example is requested
  const exampleNum = process.argv[2];

  if (exampleNum) {
    const num = parseInt(exampleNum, 10);
    if (num >= 1 && num <= examples.length) {
      console.log(`Running Example ${num}: ${examples[num - 1].name}\n`);
      await examples[num - 1].fn();
      return;
    } else {
      console.log(`Invalid example number. Choose 1-${examples.length}`);
      return;
    }
  }

  // Display available examples
  console.log('Available Examples:\n');
  examples.forEach((example, index) => {
    console.log(`  ${index + 1}. ${example.name}`);
  });
  console.log('\nUsage:');
  console.log('  ts-node device-action-reports-examples.ts [example-number]');
  console.log('\nExample:');
  console.log('  ts-node device-action-reports-examples.ts 1');
  console.log('  ts-node device-action-reports-examples.ts 5\n');

  // Uncomment to run a specific example by default
  // await example1PendingActions();
}

// Run if executed directly
if (require.main === module) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}

// Export all examples
export {
  example1PendingActions,
  example2RecentEnrollments,
  example3CustomDateRange,
  example4RecentWipes,
  example5ComprehensiveReport,
  example6FilterByOS,
  example7FilterByOwnership,
  example8QuickReport,
  example9DeviceLifecycle,
  example10AdvancedAnalysis,
};
