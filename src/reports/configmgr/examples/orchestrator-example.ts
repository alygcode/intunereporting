/**
 * ConfigMgr Device Management Reports Orchestrator - Usage Examples
 *
 * This file demonstrates how to use the master orchestrator to execute
 * all 37 Configuration Manager device management reports.
 *
 * @module orchestrator-examples
 */

import { Client } from '@microsoft/microsoft-graph-client';
import {
  ConfigMgrDeviceManagementReports,
  ReportCategory,
  createConfigMgrReports,
  executeReport,
  exportAllSupportedReports,
} from '../index';
import { AppConfig } from '../../../types';
import * as fs from 'fs';
import * as path from 'path';

// ============================================================================
// Example Configuration
// ============================================================================

/**
 * Sample configuration for the orchestrator
 */
const sampleConfig: AppConfig = {
  authentication: {
    tenantId: 'your-tenant-id',
    clientId: 'your-client-id',
    clientSecret: 'your-client-secret',
    authMethod: 'clientSecret',
  },
  reports: {
    enabled: ['*'],
    disabled: [],
    settings: {},
  },
  output: {
    defaultFormat: 'json',
    directory: './output/configmgr-reports',
    includeTimestamp: true,
    compression: false,
  },
  scheduler: {
    enabled: false,
    timezone: 'UTC',
    schedules: [],
  },
  logging: {
    level: 'info',
    file: './logs/configmgr-reports.log',
    console: true,
    maxSize: '10m',
    maxFiles: 5,
  },
};

// ============================================================================
// Example 1: Basic Setup and Report Discovery
// ============================================================================

async function example1_BasicSetup(graphClient: Client) {
  console.log('\n=== Example 1: Basic Setup and Report Discovery ===\n');

  // Create orchestrator instance
  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  // List all available reports
  console.log('📋 All Available Reports:');
  const allReports = reports.getAllReports();
  console.log(`Total reports: ${allReports.length}\n`);

  // Show first 5 reports
  allReports.slice(0, 5).forEach(report => {
    console.log(`Report ${report.reportNumber}: ${report.reportName}`);
    console.log(`  Category: ${report.category}`);
    console.log(`  Supported: ${report.isSupported ? '✅' : '❌'}`);
    console.log(`  Intune Equivalent: ${report.modernIntuneEquivalent}\n`);
  });

  // Count supported vs unsupported reports
  const supportedReports = reports.getSupportedReports();
  const unsupportedReports = reports.getUnsupportedReports();

  console.log(`\n📊 Report Summary:`);
  console.log(`  Supported: ${supportedReports.length}`);
  console.log(`  Legacy/Unsupported: ${unsupportedReports.length}`);
}

// ============================================================================
// Example 2: Execute Individual Reports
// ============================================================================

async function example2_ExecuteIndividualReports(graphClient: Client) {
  console.log('\n=== Example 2: Execute Individual Reports ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  // Report 1: All corporate-owned mobile devices
  console.log('📱 Report 1: Corporate-Owned Mobile Devices');
  const corporateDevices = await reports.getReportByNumber(1);
  console.log(`Found ${corporateDevices.data.length} corporate devices`);
  console.log(`Generated: ${corporateDevices.metadata.generatedAt}\n`);

  // Report 22: Mobile devices by operating system
  console.log('💻 Report 22: Devices by Operating System');
  const osDist = await reports.getReportByNumber(22);
  console.log(`OS categories: ${osDist.data.length}`);
  osDist.data.forEach((os: any) => {
    console.log(`  ${os.operatingSystem}: ${os.deviceCount} devices (${os.percentage}%)`);
  });

  // Report 23: Jailbroken or rooted devices
  console.log('\n🔒 Report 23: Security - Jailbroken Devices');
  const jailbrokenDevices = await reports.getReportByNumber(23);
  if (jailbrokenDevices.data.length > 0) {
    console.log(`⚠️  Warning: ${jailbrokenDevices.data.length} compromised devices found!`);
    jailbrokenDevices.data.slice(0, 3).forEach((device: any) => {
      console.log(`  - ${device.deviceName} (${device.riskLevel} risk)`);
    });
  } else {
    console.log('✅ No jailbroken devices detected');
  }
}

// ============================================================================
// Example 3: Execute Reports with Filters
// ============================================================================

async function example3_ExecuteWithFilters(graphClient: Client) {
  console.log('\n=== Example 3: Execute Reports with Filters ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  // Report 1 with iOS filter
  console.log('🍎 Corporate iOS Devices Only:');
  const iosDevices = await reports.getReportByNumber(1, {
    operatingSystem: 'iOS',
  });
  console.log(`iOS devices: ${iosDevices.data.length}\n`);

  // Report 28: Devices with low memory (custom threshold)
  console.log('💾 Devices with Low Memory (< 1GB):');
  const lowMemoryDevices = await reports.getReportByNumber(28, {
    thresholdMB: 1024, // 1GB
  });
  console.log(`Devices below threshold: ${lowMemoryDevices.data.length}`);
  if (lowMemoryDevices.data.length > 0) {
    lowMemoryDevices.data.slice(0, 3).forEach((device: any) => {
      console.log(`  - ${device.deviceName}: ${device.totalMemoryInMB}MB RAM`);
    });
  }

  // Report 15: Inactive devices (custom period)
  console.log('\n📊 Inactive Devices (> 60 days):');
  const inactiveDevices = await reports.getReportByNumber(15, {
    inactiveDays: 60,
  });
  console.log(`Inactive devices: ${inactiveDevices.data.length}`);
}

// ============================================================================
// Example 4: Execute Reports by Category
// ============================================================================

async function example4_ExecuteByCategory(graphClient: Client) {
  console.log('\n=== Example 4: Execute Reports by Category ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  // Get all security reports
  console.log('🔐 Executing All Security Reports:');
  const securityReports = reports.getReportsByCategory(ReportCategory.SECURITY);

  for (const reportMeta of securityReports) {
    console.log(`\nReport ${reportMeta.reportNumber}: ${reportMeta.reportName}`);
    try {
      const reportData = await reports.getReportByNumber(reportMeta.reportNumber);
      console.log(`  ✅ Success: ${reportData.data.length} records`);
    } catch (error) {
      console.log(`  ❌ Failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  // Get all hardware reports
  console.log('\n\n🖥️  Executing All Hardware Reports:');
  const hardwareReports = reports.getReportsByCategory(ReportCategory.HARDWARE);

  for (const reportMeta of hardwareReports) {
    console.log(`\nReport ${reportMeta.reportNumber}: ${reportMeta.reportName}`);
    try {
      const reportData = await reports.getReportByNumber(reportMeta.reportNumber);
      console.log(`  ✅ Success: ${reportData.data.length} records`);
    } catch (error) {
      console.log(`  ❌ Failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }
}

// ============================================================================
// Example 5: Batch Export Multiple Reports
// ============================================================================

async function example5_BatchExport(graphClient: Client) {
  console.log('\n=== Example 5: Batch Export Multiple Reports ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  // Export specific reports
  console.log('📦 Exporting Reports 1, 2, 20, 22 to JSON...');
  const reportNumbers = [1, 2, 20, 22];

  const results = await reports.exportReports(reportNumbers, {
    outputFormat: 'json',
    outputDirectory: './output/batch-export',
    includeTimestamp: true,
    parallel: true,
    maxParallel: 2,
  });

  console.log(`\n📊 Batch Export Results:`);
  console.log(`  Total: ${results.totalReports}`);
  console.log(`  Success: ${results.successCount}`);
  console.log(`  Failed: ${results.failureCount}`);
  console.log(`  Duration: ${results.summary.executionTimeMs}ms`);
  console.log(`  Total Records: ${results.summary.totalRecords}`);

  // Show individual results
  console.log('\n📋 Individual Report Results:');
  results.results.forEach(result => {
    const status = result.success ? '✅' : '❌';
    console.log(`  ${status} ${result.reportName}`);
    if (result.success) {
      console.log(`      Output: ${result.outputPath}`);
      console.log(`      Records: ${result.recordCount}`);
      console.log(`      Duration: ${result.duration}ms`);
    } else {
      console.log(`      Error: ${result.error}`);
    }
  });
}

// ============================================================================
// Example 6: Export All Supported Reports
// ============================================================================

async function example6_ExportAllReports(graphClient: Client) {
  console.log('\n=== Example 6: Export All Supported Reports ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  console.log('🚀 Exporting all 29 supported reports...');
  console.log('This may take several minutes...\n');

  const startTime = Date.now();

  const results = await reports.exportAllReports({
    outputFormat: 'json',
    outputDirectory: './output/all-reports',
    includeTimestamp: true,
    parallel: true,
    maxParallel: 5,
  });

  const duration = Date.now() - startTime;

  console.log('\n✅ All Reports Export Complete!');
  console.log(`\n📊 Summary:`);
  console.log(`  Total Reports: ${results.totalReports}`);
  console.log(`  Successful: ${results.successCount}`);
  console.log(`  Failed: ${results.failureCount}`);
  console.log(`  Total Duration: ${duration}ms (${(duration / 1000).toFixed(2)}s)`);
  console.log(`  Total Records: ${results.summary.totalRecords}`);

  // Show failed reports if any
  if (results.failureCount > 0) {
    console.log('\n❌ Failed Reports:');
    results.results
      .filter(r => !r.success)
      .forEach(r => {
        console.log(`  - ${r.reportName}: ${r.error}`);
      });
  }

  // Calculate average execution time
  const avgDuration = results.results.reduce((sum, r) => sum + r.duration, 0) / results.results.length;
  console.log(`\n⚡ Performance:`);
  console.log(`  Average per report: ${avgDuration.toFixed(0)}ms`);
  console.log(`  Records per second: ${(results.summary.totalRecords / (duration / 1000)).toFixed(0)}`);
}

// ============================================================================
// Example 7: Export Reports in Different Formats
// ============================================================================

async function example7_MultipleFormats(graphClient: Client) {
  console.log('\n=== Example 7: Export Reports in Different Formats ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  const reportNumber = 1; // Corporate devices report

  // Export as JSON
  console.log('📄 Exporting as JSON...');
  await reports.exportReports([reportNumber], {
    outputFormat: 'json',
    outputDirectory: './output/formats/json',
  });

  // Export as CSV
  console.log('📊 Exporting as CSV...');
  await reports.exportReports([reportNumber], {
    outputFormat: 'csv',
    outputDirectory: './output/formats/csv',
  });

  // Export as HTML
  console.log('🌐 Exporting as HTML...');
  await reports.exportReports([reportNumber], {
    outputFormat: 'html',
    outputDirectory: './output/formats/html',
  });

  console.log('\n✅ All formats exported successfully!');
  console.log('Check ./output/formats/ directory for results');
}

// ============================================================================
// Example 8: Generate Comprehensive Dashboard
// ============================================================================

async function example8_GenerateDashboard(graphClient: Client) {
  console.log('\n=== Example 8: Generate Comprehensive Dashboard ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  console.log('🎯 Generating dashboard...');
  console.log('This aggregates data from multiple reports...\n');

  const dashboard = await reports.generateDashboard();

  // Display summary
  console.log('📊 Dashboard Summary:');
  console.log(`  Generated: ${dashboard.generated}`);
  console.log(`  Tenant ID: ${dashboard.tenantId}`);

  console.log('\n📱 Device Statistics:');
  console.log(`  Total Devices: ${dashboard.summary.totalDevices}`);
  console.log(`  Total Users: ${dashboard.summary.totalUsers}`);
  console.log(`  Corporate Devices: ${dashboard.summary.corporateDevices}`);
  console.log(`  Personal Devices: ${dashboard.summary.personalDevices}`);

  console.log('\n✅ Compliance:');
  console.log(`  Compliant: ${dashboard.summary.compliantDevices}`);
  console.log(`  Non-Compliant: ${dashboard.summary.nonCompliantDevices}`);
  const complianceRate = (dashboard.summary.compliantDevices / dashboard.summary.totalDevices * 100).toFixed(1);
  console.log(`  Compliance Rate: ${complianceRate}%`);

  console.log('\n🔄 Activity:');
  console.log(`  Active (7d): ${dashboard.summary.activeDevices}`);
  console.log(`  Inactive: ${dashboard.summary.inactiveDevices}`);

  // Display top insights
  if (dashboard.topInsights.length > 0) {
    console.log('\n⚠️  Top Insights:');
    dashboard.topInsights.forEach(insight => {
      const icon = insight.severity === 'critical' ? '🔴' :
                   insight.severity === 'warning' ? '🟡' : '🔵';
      console.log(`  ${icon} ${insight.message} (${insight.count})`);
    });
  }

  // Display recommendations
  if (dashboard.recommendations.length > 0) {
    console.log('\n💡 Recommendations:');
    dashboard.recommendations.forEach((rec, index) => {
      console.log(`  ${index + 1}. ${rec}`);
    });
  }

  // Save dashboard to file
  const dashboardPath = './output/dashboard.json';
  fs.mkdirSync(path.dirname(dashboardPath), { recursive: true });
  fs.writeFileSync(dashboardPath, JSON.stringify(dashboard, null, 2));
  console.log(`\n💾 Dashboard saved to: ${dashboardPath}`);
}

// ============================================================================
// Example 9: Sequential vs Parallel Execution
// ============================================================================

async function example9_SequentialVsParallel(graphClient: Client) {
  console.log('\n=== Example 9: Sequential vs Parallel Execution ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);
  const reportNumbers = [1, 2, 20, 22, 23];

  // Sequential execution
  console.log('⏱️  Sequential Execution:');
  const seqStart = Date.now();
  const seqResults = await reports.exportReports(reportNumbers, {
    outputFormat: 'json',
    outputDirectory: './output/sequential',
    parallel: false,
  });
  const seqDuration = Date.now() - seqStart;
  console.log(`  Duration: ${seqDuration}ms`);
  console.log(`  Success: ${seqResults.successCount}/${seqResults.totalReports}`);

  // Parallel execution
  console.log('\n⚡ Parallel Execution:');
  const parStart = Date.now();
  const parResults = await reports.exportReports(reportNumbers, {
    outputFormat: 'json',
    outputDirectory: './output/parallel',
    parallel: true,
    maxParallel: 3,
  });
  const parDuration = Date.now() - parStart;
  console.log(`  Duration: ${parDuration}ms`);
  console.log(`  Success: ${parResults.successCount}/${parResults.totalReports}`);

  // Compare
  const speedup = (seqDuration / parDuration).toFixed(2);
  const improvement = (((seqDuration - parDuration) / seqDuration) * 100).toFixed(1);

  console.log('\n📊 Performance Comparison:');
  console.log(`  Speedup: ${speedup}x faster`);
  console.log(`  Time Saved: ${improvement}% reduction`);
  console.log(`  Absolute Savings: ${seqDuration - parDuration}ms`);
}

// ============================================================================
// Example 10: Error Handling and Recovery
// ============================================================================

async function example10_ErrorHandling(graphClient: Client) {
  console.log('\n=== Example 10: Error Handling and Recovery ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  // Try to execute an unsupported report (Windows CE)
  console.log('❌ Attempting unsupported report (Windows CE):');
  try {
    const ceReport = await reports.getReportByNumber(3);
    console.log(`Status: ${ceReport.summary.isSupported ? 'Supported' : 'Not Supported'}`);
    console.log(`Warning: ${ceReport.summary.warning}`);
    console.log(`Migration Notes: ${ceReport.summary.migrationNotes}`);
  } catch (error) {
    console.log(`Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }

  // Try invalid report number
  console.log('\n❌ Attempting invalid report number (999):');
  try {
    await reports.getReportByNumber(999);
  } catch (error) {
    console.log(`Caught error: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }

  // Batch execution with error recovery
  console.log('\n📦 Batch execution with mixed valid/invalid reports:');
  const mixedReports = [1, 3, 999, 22]; // Mix of valid, unsupported, and invalid

  const results = await Promise.allSettled(
    mixedReports.map(num => reports.getReportByNumber(num))
  );

  results.forEach((result, index) => {
    const reportNum = mixedReports[index];
    if (result.status === 'fulfilled') {
      console.log(`  ✅ Report ${reportNum}: ${result.value.data.length} records`);
    } else {
      console.log(`  ❌ Report ${reportNum}: ${result.reason.message}`);
    }
  });
}

// ============================================================================
// Example 11: Convenience Functions
// ============================================================================

async function example11_ConvenienceFunctions(graphClient: Client) {
  console.log('\n=== Example 11: Convenience Functions ===\n');

  // Create orchestrator using convenience function
  console.log('🎯 Using createConfigMgrReports():');
  const reports = createConfigMgrReports(graphClient, sampleConfig);
  console.log('Orchestrator created\n');

  // Execute single report using convenience function
  console.log('📄 Using executeReport():');
  const report = await executeReport(graphClient, sampleConfig, 1);
  console.log(`Report data: ${report.data.length} records\n`);

  // Export all reports using convenience function
  console.log('📦 Using exportAllSupportedReports():');
  const results = await exportAllSupportedReports(graphClient, sampleConfig, {
    outputFormat: 'json',
    outputDirectory: './output/convenience',
    parallel: true,
    maxParallel: 3,
  });
  console.log(`Exported ${results.successCount}/${results.totalReports} reports`);
}

// ============================================================================
// Example 12: Complete Workflow
// ============================================================================

async function example12_CompleteWorkflow(graphClient: Client) {
  console.log('\n=== Example 12: Complete Reporting Workflow ===\n');

  const reports = new ConfigMgrDeviceManagementReports(graphClient, sampleConfig);

  console.log('🎯 Step 1: Discover available reports');
  const allReports = reports.getAllReports();
  const supportedReports = reports.getSupportedReports();
  console.log(`  Total: ${allReports.length} | Supported: ${supportedReports.length}\n`);

  console.log('🎯 Step 2: Execute key security reports');
  const securityReports = reports.getReportsByCategory(ReportCategory.SECURITY);
  console.log(`  Found ${securityReports.length} security reports`);

  for (const reportMeta of securityReports) {
    const data = await reports.getReportByNumber(reportMeta.reportNumber);
    console.log(`  ✅ ${reportMeta.reportName}: ${data.data.length} records`);
  }

  console.log('\n🎯 Step 3: Generate dashboard');
  const dashboard = await reports.generateDashboard();
  console.log(`  Devices: ${dashboard.summary.totalDevices}`);
  console.log(`  Compliance: ${(dashboard.summary.compliantDevices / dashboard.summary.totalDevices * 100).toFixed(1)}%`);
  console.log(`  Insights: ${dashboard.topInsights.length}`);

  console.log('\n🎯 Step 4: Export all reports');
  const exportResults = await reports.exportAllReports({
    outputFormat: 'json',
    outputDirectory: './output/workflow',
    parallel: true,
    maxParallel: 5,
  });

  console.log(`  Success: ${exportResults.successCount}/${exportResults.totalReports}`);
  console.log(`  Duration: ${(exportResults.summary.executionTimeMs / 1000).toFixed(2)}s`);

  console.log('\n✅ Complete workflow finished!');
  console.log('📊 Reports available in: ./output/workflow/');
}

// ============================================================================
// Main Execution
// ============================================================================

async function main() {
  console.log('╔═══════════════════════════════════════════════════════════╗');
  console.log('║  ConfigMgr Device Management Reports - Usage Examples    ║');
  console.log('║  Orchestrator for all 37 Reports                          ║');
  console.log('╚═══════════════════════════════════════════════════════════╝');

  // Note: In a real implementation, you would initialize the Graph client
  // with proper authentication. This is a placeholder.
  const graphClient = Client.init({
    authProvider: (done) => {
      // Your auth implementation here
      done(null, 'access-token');
    },
  });

  try {
    // Run examples
    await example1_BasicSetup(graphClient);
    await example2_ExecuteIndividualReports(graphClient);
    await example3_ExecuteWithFilters(graphClient);
    await example4_ExecuteByCategory(graphClient);
    await example5_BatchExport(graphClient);
    await example6_ExportAllReports(graphClient);
    await example7_MultipleFormats(graphClient);
    await example8_GenerateDashboard(graphClient);
    await example9_SequentialVsParallel(graphClient);
    await example10_ErrorHandling(graphClient);
    await example11_ConvenienceFunctions(graphClient);
    await example12_CompleteWorkflow(graphClient);

    console.log('\n✅ All examples completed successfully!');
  } catch (error) {
    console.error('\n❌ Error running examples:', error);
  }
}

// Run if executed directly
if (require.main === module) {
  main().catch(console.error);
}

// Export examples for individual testing
export {
  example1_BasicSetup,
  example2_ExecuteIndividualReports,
  example3_ExecuteWithFilters,
  example4_ExecuteByCategory,
  example5_BatchExport,
  example6_ExportAllReports,
  example7_MultipleFormats,
  example8_GenerateDashboard,
  example9_SequentialVsParallel,
  example10_ErrorHandling,
  example11_ConvenienceFunctions,
  example12_CompleteWorkflow,
};
