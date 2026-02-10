/**
 * Example Usage of User Enrollment Report
 *
 * This file demonstrates various ways to use the User Enrollment Report
 * to generate comprehensive enrollment analytics from Microsoft Intune.
 *
 * Features demonstrated:
 * - Get all enrolled users
 * - Get devices per user
 * - Get enrollment status and methods
 * - Get enrollment failures and troubleshooting
 * - Generate enrollment trend reports
 * - Track user device ownership
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { UserEnrollmentReport, UserEnrollmentReportOptions } from '../user-enrollment-report';
import { AppConfig } from '../../types';
import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';

// Load environment variables
dotenv.config();

/**
 * Helper function to create authenticated Graph client
 */
function createGraphClient(): Client {
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
 * Helper function to create report instance
 */
function createReport(): UserEnrollmentReport {
  const graphClient = createGraphClient();
  const config: AppConfig = {
    tenantId: process.env.AZURE_TENANT_ID || '',
    clientId: process.env.AZURE_CLIENT_ID || '',
    clientSecret: process.env.AZURE_CLIENT_SECRET || '',
    outputDir: './reports/enrollment'
  };

  return new UserEnrollmentReport(graphClient, config);
}

/**
 * Helper function to save report to file
 */
function saveReport(filename: string, content: string): void {
  const dir = './reports/enrollment';
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  fs.writeFileSync(path.join(dir, filename), content);
  console.log(`Report saved to: ${path.join(dir, filename)}`);
}

/**
 * Example 1: Basic enrollment report with all features
 */
async function example1BasicEnrollmentReport(): Promise<void> {
  console.log('\n=== Example 1: Basic Enrollment Report ===\n');

  try {
    const report = createReport();

    const options: UserEnrollmentReportOptions = {
      includeTrends: true,
      includeFailures: true,
      includePlatformStats: true,
      includeUserDetails: false,
      trendDays: 30,
      maxFailures: 100
    };

    console.log('Generating comprehensive enrollment report...');
    const result = await report.execute(options);

    console.log('\n📊 Enrollment Statistics:');
    console.log(`Total Devices: ${result.summary?.totalDevices || 0}`);
    console.log(`Total Users: ${result.summary?.totalUsers || 0}`);
    console.log(`Success Rate: ${result.summary?.successRate || 0}%`);
    console.log(`Failed Enrollments: ${result.summary?.failedEnrollments || 0}`);

    // Save reports in different formats
    saveReport('enrollment-report.json', report.exportAsJSON(result));
    saveReport('enrollment-report.csv', report.exportAsCSV(result));
    saveReport('enrollment-report.html', report.exportAsHTML(result));

    console.log('\n✅ Report generation complete!');
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 2: Get enrollment statistics for a specific time period
 */
async function example2EnrollmentStatisticsByPeriod(): Promise<void> {
  console.log('\n=== Example 2: Enrollment Statistics by Period ===\n');

  try {
    const report = createReport();

    // Last 30 days
    const endDate = new Date();
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 30);

    console.log(`Fetching statistics from ${startDate.toLocaleDateString()} to ${endDate.toLocaleDateString()}...`);
    const stats = await report.getEnrollmentStatistics(startDate, endDate);

    console.log('\n📈 Enrollment Statistics (Last 30 Days):');
    console.log(`Total Devices: ${stats.totalDevices}`);
    console.log(`Total Users: ${stats.totalUsers}`);
    console.log(`Enrolled Last 24h: ${stats.enrolledLast24Hours}`);
    console.log(`Enrolled Last 7d: ${stats.enrolledLast7Days}`);
    console.log(`Enrolled Last 30d: ${stats.enrolledLast30Days}`);
    console.log(`Average Devices/User: ${stats.averageDevicesPerUser.toFixed(2)}`);
    console.log(`Success Rate: ${stats.enrollmentSuccessRate}%`);
    console.log(`Failed Enrollments: ${stats.failedEnrollments}`);
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 3: Get all enrolled users
 */
async function example3GetAllEnrolledUsers(): Promise<void> {
  console.log('\n=== Example 3: Get All Enrolled Users ===\n');

  try {
    const report = createReport();

    console.log('Fetching all enrolled users...');
    const users = await report.getAllEnrolledUsers();

    console.log(`\n👥 Total Enrolled Users: ${users.length}`);

    if (users.length > 0) {
      console.log('\nFirst 10 enrolled users:');
      users.slice(0, 10).forEach((user, index) => {
        console.log(`${index + 1}. ${user.displayName} (${user.userPrincipalName})`);
        if (user.department) console.log(`   Department: ${user.department}`);
        if (user.jobTitle) console.log(`   Job Title: ${user.jobTitle}`);
      });
    }

    // Save to JSON
    saveReport('enrolled-users.json', JSON.stringify(users, null, 2));
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 4: Get devices for a specific user
 */
async function example4GetUserDevices(): Promise<void> {
  console.log('\n=== Example 4: Get User Devices ===\n');

  try {
    const report = createReport();

    // Replace with actual user ID or use first user from your tenant
    const userId = process.env.TEST_USER_ID || '';

    if (!userId) {
      console.log('⚠️  Set TEST_USER_ID environment variable to test this example');
      return;
    }

    console.log(`Fetching devices for user: ${userId}...`);
    const devices = await report.getUserDevices(userId);

    console.log(`\n📱 Total Devices: ${devices.length}`);

    if (devices.length > 0) {
      console.log('\nDevice Details:');
      devices.forEach((device, index) => {
        console.log(`\n${index + 1}. ${device.deviceName}`);
        console.log(`   Platform: ${device.platform}`);
        console.log(`   OS Version: ${device.osVersion}`);
        console.log(`   Model: ${device.model || 'N/A'}`);
        console.log(`   Manufacturer: ${device.manufacturer || 'N/A'}`);
        console.log(`   Enrollment Date: ${device.enrollmentDate}`);
        console.log(`   Enrollment Type: ${device.enrollmentType}`);
        console.log(`   Management State: ${device.managementState}`);
        console.log(`   Compliance State: ${device.complianceState}`);
        console.log(`   Last Sync: ${device.lastSyncDateTime}`);
      });
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 5: Get user enrollment status with comprehensive details
 */
async function example5GetUserEnrollmentStatus(): Promise<void> {
  console.log('\n=== Example 5: User Enrollment Status ===\n');

  try {
    const report = createReport();

    const userId = process.env.TEST_USER_ID || '';

    if (!userId) {
      console.log('⚠️  Set TEST_USER_ID environment variable to test this example');
      return;
    }

    console.log(`Fetching comprehensive enrollment status for user: ${userId}...`);
    const status = await report.getComprehensiveUserEnrollmentStatus(userId);

    console.log('\n👤 User Information:');
    console.log(`Name: ${status.displayName}`);
    console.log(`UPN: ${status.userPrincipalName}`);
    console.log(`Department: ${status.department || 'N/A'}`);
    console.log(`Job Title: ${status.jobTitle || 'N/A'}`);

    console.log('\n📊 Enrollment Summary:');
    console.log(`Enrolled Devices: ${status.enrolledDeviceCount}`);
    console.log(`First Enrollment: ${status.firstEnrollmentDate || 'N/A'}`);
    console.log(`Last Enrollment: ${status.lastEnrollmentDate || 'N/A'}`);
    console.log(`Has Failures: ${status.hasEnrollmentFailures ? 'Yes' : 'No'}`);
    console.log(`Failure Count: ${status.failureCount}`);

    if (status.ownedDevices && status.ownedDevices.length > 0) {
      console.log(`\n🔑 Owned Devices: ${status.ownedDevices.length}`);
    }

    // Save detailed status
    saveReport(`user-${userId}-enrollment-status.json`, JSON.stringify(status, null, 2));
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 6: Get enrollment failures with troubleshooting
 */
async function example6GetEnrollmentFailures(): Promise<void> {
  console.log('\n=== Example 6: Enrollment Failures ===\n');

  try {
    const report = createReport();

    console.log('Fetching enrollment failures from last 30 days...');
    const failures = await report.getEnrollmentFailures(30, 50);

    console.log(`\n❌ Total Failures: ${failures.length}`);

    if (failures.length > 0) {
      console.log('\nRecent Enrollment Failures:');

      // Group by category
      const byCategory = failures.reduce((acc, failure) => {
        const cat = failure.failureCategory;
        if (!acc[cat]) acc[cat] = [];
        acc[cat].push(failure);
        return acc;
      }, {} as Record<string, typeof failures>);

      console.log('\n📊 Failures by Category:');
      Object.entries(byCategory).forEach(([category, fails]) => {
        console.log(`- ${category}: ${fails.length}`);
      });

      console.log('\n🔍 Top 5 Recent Failures:');
      failures.slice(0, 5).forEach((failure, index) => {
        console.log(`\n${index + 1}. ${failure.userPrincipalName} - ${failure.platform}`);
        console.log(`   Date: ${new Date(failure.failureDateTime).toLocaleString()}`);
        console.log(`   Category: ${failure.failureCategory}`);
        console.log(`   Error Code: ${failure.errorCode}`);
        console.log(`   Reason: ${failure.failureReason}`);
        console.log(`   Troubleshooting Steps:`);
        failure.troubleshootingSteps.slice(0, 3).forEach((step, i) => {
          console.log(`     ${i + 1}. ${step}`);
        });
      });

      // Save failures report
      saveReport('enrollment-failures.json', JSON.stringify(failures, null, 2));
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 7: Generate enrollment trends
 */
async function example7EnrollmentTrends(): Promise<void> {
  console.log('\n=== Example 7: Enrollment Trends ===\n');

  try {
    const report = createReport();

    console.log('Generating enrollment trends for last 30 days...');
    const trends = await report.getEnrollmentTrends(30);

    console.log(`\n📈 Trend Data Points: ${trends.length}`);

    if (trends.length > 0) {
      console.log('\nLast 7 Days Trend:');
      trends.slice(-7).forEach(trend => {
        console.log(`\n${trend.date}:`);
        console.log(`  Total Enrollments: ${trend.totalEnrollments}`);
        console.log(`  Successful: ${trend.successfulEnrollments}`);
        console.log(`  Failed: ${trend.failedEnrollments}`);
        console.log(`  Unique Users: ${trend.uniqueUsers}`);
        console.log(`  Platform Breakdown: iOS:${trend.platformBreakdown.iOS}, Android:${trend.platformBreakdown.Android}, Windows:${trend.platformBreakdown.Windows}`);
      });

      // Calculate summary
      const totalEnrollments = trends.reduce((sum, t) => sum + t.totalEnrollments, 0);
      const totalSuccessful = trends.reduce((sum, t) => sum + t.successfulEnrollments, 0);
      const successRate = totalEnrollments > 0 ? (totalSuccessful / totalEnrollments * 100).toFixed(2) : 0;

      console.log(`\n📊 Summary (Last 30 Days):`);
      console.log(`Total Enrollments: ${totalEnrollments}`);
      console.log(`Successful: ${totalSuccessful}`);
      console.log(`Success Rate: ${successRate}%`);

      // Save trends
      saveReport('enrollment-trends.json', JSON.stringify(trends, null, 2));
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 8: Platform-specific enrollment statistics
 */
async function example8PlatformStatistics(): Promise<void> {
  console.log('\n=== Example 8: Platform Statistics ===\n');

  try {
    const report = createReport();

    console.log('Fetching platform-specific enrollment statistics...');
    const platformStats = await report.getDeviceEnrollmentByPlatform();

    console.log(`\n📱 Platforms: ${platformStats.length}`);

    platformStats.forEach(stat => {
      console.log(`\n${stat.platform}:`);
      console.log(`  Total Devices: ${stat.totalDevices}`);
      console.log(`  Enrolled Last 24h: ${stat.enrolledLast24Hours}`);
      console.log(`  Enrolled Last 7d: ${stat.enrolledLast7Days}`);
      console.log(`  Enrolled Last 30d: ${stat.enrolledLast30Days}`);

      if (stat.topModels.length > 0) {
        console.log(`  Top Models:`);
        stat.topModels.slice(0, 3).forEach(model => {
          console.log(`    - ${model.model}: ${model.count} (${model.percentage}%)`);
        });
      }

      if (stat.osVersionDistribution.length > 0) {
        console.log(`  Top OS Versions:`);
        stat.osVersionDistribution.slice(0, 3).forEach(os => {
          console.log(`    - ${os.osVersion}: ${os.count} (${os.percentage}%)`);
        });
      }
    });

    // Save platform stats
    saveReport('platform-statistics.json', JSON.stringify(platformStats, null, 2));
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 9: Enrollment methods breakdown
 */
async function example9EnrollmentMethods(): Promise<void> {
  console.log('\n=== Example 9: Enrollment Methods ===\n');

  try {
    const report = createReport();

    console.log('Analyzing enrollment methods...');
    const methods = await report.getEnrollmentMethodsBreakdown();

    console.log('\n📋 Enrollment Methods Breakdown:');

    Object.entries(methods)
      .sort((a, b) => b[1].count - a[1].count)
      .forEach(([method, data]) => {
        console.log(`\n${method}:`);
        console.log(`  Count: ${data.count}`);
        console.log(`  Percentage: ${data.percentage}%`);
        if (data.devices.length > 0) {
          console.log(`  Example Devices: ${data.devices.slice(0, 3).join(', ')}`);
        }
      });

    // Get enrollment status summary
    console.log('\n\n📊 Enrollment Status Summary:');
    const statusSummary = await report.getEnrollmentStatusSummary();

    console.log(`Total Devices: ${statusSummary.totalDevices}`);

    console.log('\nBy Management State:');
    Object.entries(statusSummary.byManagementState).forEach(([state, count]) => {
      console.log(`  - ${state}: ${count}`);
    });

    console.log('\nBy Owner Type:');
    Object.entries(statusSummary.byOwnerType).forEach(([type, count]) => {
      console.log(`  - ${type}: ${count}`);
    });

    // Save methods breakdown
    saveReport('enrollment-methods.json', JSON.stringify({ methods, statusSummary }, null, 2));
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 10: Device ownership tracking
 */
async function example10DeviceOwnership(): Promise<void> {
  console.log('\n=== Example 10: Device Ownership Tracking ===\n');

  try {
    const report = createReport();

    console.log('Analyzing device ownership...');
    const ownership = await report.getDeviceOwnershipBreakdown();

    console.log('\n🏢 Device Ownership Breakdown:');
    console.log(`Total Devices: ${ownership.totalDevices}`);
    console.log(`\nCorporate Devices: ${ownership.corporate.count} (${ownership.corporate.percentage}%)`);
    console.log(`Personal Devices: ${ownership.personal.count} (${ownership.personal.percentage}%)`);
    console.log(`Unknown: ${ownership.unknown.count} (${ownership.unknown.percentage}%)`);

    console.log('\n📱 Ownership by Platform:');
    Object.entries(ownership.byPlatform).forEach(([platform, counts]) => {
      console.log(`\n${platform}:`);
      console.log(`  Corporate: ${counts.corporate}`);
      console.log(`  Personal: ${counts.personal}`);
      console.log(`  Unknown: ${counts.unknown}`);
    });

    // If user ID is available, get user-specific ownership
    const userId = process.env.TEST_USER_ID;
    if (userId) {
      console.log(`\n\n👤 User Device Ownership (${userId}):`);
      const userOwnership = await report.getUserDeviceOwnership(userId);
      console.log(`Total Devices: ${userOwnership.totalDevices}`);
      console.log(`Corporate: ${userOwnership.corporateDevices.length}`);
      console.log(`Personal: ${userOwnership.personalDevices.length}`);
    }

    // Save ownership data
    saveReport('device-ownership.json', JSON.stringify(ownership, null, 2));
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 11: Troubleshooting events for a specific user
 */
async function example11UserTroubleshooting(): Promise<void> {
  console.log('\n=== Example 11: User Troubleshooting Events ===\n');

  try {
    const report = createReport();

    const userId = process.env.TEST_USER_ID || '';

    if (!userId) {
      console.log('⚠️  Set TEST_USER_ID environment variable to test this example');
      return;
    }

    console.log(`Fetching troubleshooting events for user: ${userId}...`);
    const events = await report.getUserTroubleshootingEvents(userId, 90);

    console.log(`\n🔧 Total Troubleshooting Events: ${events.length}`);

    if (events.length > 0) {
      console.log('\nRecent Events:');
      events.slice(0, 10).forEach((event, index) => {
        console.log(`\n${index + 1}. ${event.eventName || 'Unknown Event'}`);
        console.log(`   Date: ${new Date(event.eventDateTime).toLocaleString()}`);
        console.log(`   Correlation ID: ${event.correlationId || 'N/A'}`);
        if (event.troubleshootingErrorDetails) {
          console.log(`   Error Details: ${JSON.stringify(event.troubleshootingErrorDetails)}`);
        }
      });

      saveReport(`user-${userId}-troubleshooting.json`, JSON.stringify(events, null, 2));
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 12: Get devices by user principal name
 */
async function example12GetDevicesByUPN(): Promise<void> {
  console.log('\n=== Example 12: Get Devices by UPN ===\n');

  try {
    const report = createReport();

    const upn = process.env.TEST_USER_UPN || '';

    if (!upn) {
      console.log('⚠️  Set TEST_USER_UPN environment variable to test this example');
      return;
    }

    console.log(`Fetching devices for UPN: ${upn}...`);
    const devices = await report.getUserDevicesByUPN(upn);

    console.log(`\n📱 Total Devices: ${devices.length}`);

    if (devices.length > 0) {
      devices.forEach((device, index) => {
        console.log(`\n${index + 1}. ${device.deviceName}`);
        console.log(`   Platform: ${device.platform}`);
        console.log(`   Enrollment Type: ${device.enrollmentType}`);
        console.log(`   Compliance: ${device.complianceState}`);
      });
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Main function to run examples
 */
async function main(): Promise<void> {
  const examples = [
    { name: 'Basic Enrollment Report', fn: example1BasicEnrollmentReport },
    { name: 'Enrollment Statistics by Period', fn: example2EnrollmentStatisticsByPeriod },
    { name: 'Get All Enrolled Users', fn: example3GetAllEnrolledUsers },
    { name: 'Get User Devices', fn: example4GetUserDevices },
    { name: 'User Enrollment Status', fn: example5GetUserEnrollmentStatus },
    { name: 'Enrollment Failures', fn: example6GetEnrollmentFailures },
    { name: 'Enrollment Trends', fn: example7EnrollmentTrends },
    { name: 'Platform Statistics', fn: example8PlatformStatistics },
    { name: 'Enrollment Methods', fn: example9EnrollmentMethods },
    { name: 'Device Ownership', fn: example10DeviceOwnership },
    { name: 'User Troubleshooting', fn: example11UserTroubleshooting },
    { name: 'Get Devices by UPN', fn: example12GetDevicesByUPN }
  ];

  console.log('User Enrollment Report - Examples');
  console.log('==================================\n');

  // Check if a specific example is requested
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

  // Show available examples
  console.log('Available examples:');
  examples.forEach((example, index) => {
    console.log(`${index + 1}. ${example.name}`);
  });
  console.log('\nUsage: ts-node user-enrollment-report-example.ts [example-number]');
  console.log('Example: ts-node user-enrollment-report-example.ts 1\n');
  console.log('\nEnvironment variables needed:');
  console.log('- AZURE_TENANT_ID');
  console.log('- AZURE_CLIENT_ID');
  console.log('- AZURE_CLIENT_SECRET');
  console.log('- TEST_USER_ID (optional, for user-specific examples)');
  console.log('- TEST_USER_UPN (optional, for UPN-based examples)\n');
}

// Run if executed directly
if (require.main === module) {
  main().catch(console.error);
}

export {
  example1BasicEnrollmentReport,
  example2EnrollmentStatisticsByPeriod,
  example3GetAllEnrolledUsers,
  example4GetUserDevices,
  example5GetUserEnrollmentStatus,
  example6GetEnrollmentFailures,
  example7EnrollmentTrends,
  example8PlatformStatistics,
  example9EnrollmentMethods,
  example10DeviceOwnership,
  example11UserTroubleshooting,
  example12GetDevicesByUPN
};
