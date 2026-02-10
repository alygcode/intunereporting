/**
 * Device Hardware Reports - Usage Examples
 *
 * This file demonstrates comprehensive usage of the DeviceHardwareReports module
 * which replicates Configuration Manager hardware reports for Intune.
 *
 * @module device-hardware-reports-examples
 */

import { DeviceHardwareReports } from '../configmgr/device-hardware-reports';
import { GraphAuth } from '../../auth/graph-auth';
import { AppConfig } from '../../types';
import * as dotenv from 'dotenv';
import * as path from 'path';
import * as fs from 'fs';

// Load environment variables
dotenv.config();

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Get configuration from environment variables
 */
function getConfig(): AppConfig {
  return {
    authentication: {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET || '',
      authMethod: 'clientSecret',
    },
    reports: {
      enabled: ['device-hardware-reports'],
      disabled: [],
      settings: {},
    },
    output: {
      defaultFormat: 'json',
      directory: './reports',
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
      file: './logs/hardware-reports.log',
      console: true,
      maxSize: '10m',
      maxFiles: 5,
    },
  };
}

/**
 * Initialize the hardware reports instance
 */
async function initializeReporter(): Promise<DeviceHardwareReports> {
  const config = getConfig();
  const auth = new GraphAuth(config);
  const graphClient = await auth.getAuthenticatedClient();

  return new DeviceHardwareReports(graphClient, config);
}

/**
 * Print formatted section header
 */
function printHeader(title: string): void {
  console.log('\n' + '='.repeat(80));
  console.log(`  ${title}`);
  console.log('='.repeat(80) + '\n');
}

/**
 * Save report to file
 */
function saveToFile(data: any, filename: string): void {
  const outputDir = './reports/examples';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  const filepath = path.join(outputDir, filename);
  fs.writeFileSync(filepath, JSON.stringify(data, null, 2));
  console.log(`✓ Report saved to: ${filepath}\n`);
}

// ============================================================================
// Example 1: Get All Device Hardware Information
// ============================================================================

/**
 * Example 1: Retrieve all device hardware information
 *
 * This example demonstrates how to fetch complete hardware inventory
 * for all managed devices in your Intune tenant.
 */
async function example1_GetAllDeviceHardware(): Promise<void> {
  printHeader('Example 1: Get All Device Hardware Information');

  try {
    const reporter = await initializeReporter();

    console.log('Fetching all device hardware information...\n');
    const devices = await reporter.getAllDeviceHardwareInfo();

    console.log(`Total devices found: ${devices.length}\n`);

    // Display first 5 devices
    devices.slice(0, 5).forEach((device, index) => {
      console.log(`${index + 1}. ${device.deviceName}`);
      console.log(`   Manufacturer: ${device.manufacturer}`);
      console.log(`   Model: ${device.model}`);
      console.log(`   OS: ${device.operatingSystem.operatingSystem} ${device.operatingSystem.osVersion}`);
      console.log(`   Memory: ${device.memory.totalMemoryInGB} GB`);
      console.log(`   Storage: ${device.storage.totalStorageInGB} GB (${device.storage.freeStoragePercentage}% free)`);
      console.log(`   Compliance: ${device.complianceState}`);
      console.log('');
    });

    saveToFile({ totalDevices: devices.length, devices }, 'example1-all-devices.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 2: Display Configuration Summary (Report 9)
// ============================================================================

/**
 * Example 2: Count of mobile devices by display configurations
 *
 * ConfigMgr Report #9: Shows device count grouped by display settings.
 * For mobile devices, this categorizes by device tier based on RAM.
 */
async function example2_DisplayConfigurationSummary(): Promise<void> {
  printHeader('Example 2: Display Configuration Summary (Report 9)');

  try {
    const reporter = await initializeReporter();

    console.log('Generating display configuration summary...\n');
    const summary = await reporter.getDisplayConfigurationSummary();

    console.log('Display Configuration Distribution:\n');
    summary.forEach((config, index) => {
      console.log(`${index + 1}. ${config.displayResolution}`);
      console.log(`   Devices: ${config.deviceCount} (${config.percentage}%)`);
      console.log(`   Examples: ${config.devices.slice(0, 3).join(', ')}`);
      console.log('');
    });

    saveToFile(summary, 'example2-display-config.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 3: OS Distribution Report (Report 10)
// ============================================================================

/**
 * Example 3: Count of mobile devices by operating system
 *
 * ConfigMgr Report #10: Device count grouped by OS and version.
 * Provides insights into the OS landscape of your mobile fleet.
 */
async function example3_OSDistribution(): Promise<void> {
  printHeader('Example 3: OS Distribution Report (Report 10)');

  try {
    const reporter = await initializeReporter();

    console.log('Generating OS distribution summary...\n');
    const summary = await reporter.getOSDistributionSummary();

    console.log('Operating System Distribution:\n');
    summary.forEach((os, index) => {
      console.log(`${index + 1}. ${os.operatingSystem} ${os.osVersion}`);
      console.log(`   Devices: ${os.deviceCount} (${os.percentage}%)`);
      console.log(`   Examples: ${os.devices.slice(0, 3).join(', ')}`);
      console.log('');
    });

    // Calculate totals
    const totalDevices = summary.reduce((sum, os) => sum + os.deviceCount, 0);
    console.log(`\nTotal Devices: ${totalDevices}`);
    console.log(`Unique OS Versions: ${summary.length}\n`);

    saveToFile(summary, 'example3-os-distribution.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 4: Memory Range Summary (Report 11)
// ============================================================================

/**
 * Example 4: Count of mobile devices by program memory
 *
 * ConfigMgr Report #11: Device count categorized by RAM ranges.
 * Helps identify device capabilities and plan for app deployments.
 */
async function example4_MemoryRanges(): Promise<void> {
  printHeader('Example 4: Memory Range Summary (Report 11)');

  try {
    const reporter = await initializeReporter();

    console.log('Generating memory range summary...\n');
    const summary = await reporter.getMemoryRangeSummary();

    console.log('Memory Distribution:\n');
    summary.forEach((range, index) => {
      console.log(`${index + 1}. ${range.range}`);
      console.log(`   Range: ${range.rangeInMB}`);
      console.log(`   Devices: ${range.deviceCount} (${range.percentage}%)`);
      if (range.deviceCount > 0) {
        console.log(`   Examples: ${range.devices.slice(0, 3).join(', ')}`);
      }
      console.log('');
    });

    // Identify low memory devices
    const lowMemoryRanges = summary.filter(r => r.range.includes('< 2 GB') || r.range.includes('1-2 GB'));
    const lowMemoryCount = lowMemoryRanges.reduce((sum, r) => sum + r.deviceCount, 0);
    console.log(`\nDevices with <2GB RAM: ${lowMemoryCount}`);
    console.log('Consider these devices when deploying memory-intensive apps.\n');

    saveToFile(summary, 'example4-memory-ranges.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 5: Storage Range Summary (Report 12)
// ============================================================================

/**
 * Example 5: Count of mobile devices by storage memory
 *
 * ConfigMgr Report #12: Device count categorized by storage capacity.
 * Essential for planning app and content distribution.
 */
async function example5_StorageRanges(): Promise<void> {
  printHeader('Example 5: Storage Range Summary (Report 12)');

  try {
    const reporter = await initializeReporter();

    console.log('Generating storage range summary...\n');
    const summary = await reporter.getStorageRangeSummary();

    console.log('Storage Distribution:\n');
    summary.forEach((range, index) => {
      console.log(`${index + 1}. ${range.range}`);
      console.log(`   Range: ${range.rangeInGB}`);
      console.log(`   Devices: ${range.deviceCount} (${range.percentage}%)`);
      if (range.deviceCount > 0) {
        console.log(`   Examples: ${range.devices.slice(0, 3).join(', ')}`);
      }
      console.log('');
    });

    // Identify high and low capacity devices
    const lowStorageRanges = summary.filter(r => r.range.includes('< 16') || r.range.includes('16-32'));
    const highStorageRanges = summary.filter(r => r.range.includes('> 1 TB') || r.range.includes('512 GB'));

    const lowStorageCount = lowStorageRanges.reduce((sum, r) => sum + r.deviceCount, 0);
    const highStorageCount = highStorageRanges.reduce((sum, r) => sum + r.deviceCount, 0);

    console.log(`\nLow Capacity Devices (<32GB): ${lowStorageCount}`);
    console.log(`High Capacity Devices (>512GB): ${highStorageCount}\n`);

    saveToFile(summary, 'example5-storage-ranges.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 6: Devices with Specific Free Memory (Report 25)
// ============================================================================

/**
 * Example 6: Mobile devices with specific free program memory
 *
 * ConfigMgr Report #25: Find devices with free RAM within a specific range.
 * Useful for identifying devices suitable for specific workloads.
 */
async function example6_SpecificFreeMemory(): Promise<void> {
  printHeader('Example 6: Devices with Specific Free Memory (Report 25)');

  try {
    const reporter = await initializeReporter();

    // Find devices with 2GB-4GB free memory
    const minMemoryMB = 2048; // 2GB
    const maxMemoryMB = 4096; // 4GB

    console.log(`Finding devices with ${minMemoryMB / 1024}GB - ${maxMemoryMB / 1024}GB free memory...\n`);
    const devices = await reporter.getDevicesWithSpecificFreeMemory(minMemoryMB, maxMemoryMB);

    console.log(`Found ${devices.length} devices:\n`);

    devices.slice(0, 10).forEach((device, index) => {
      console.log(`${index + 1}. ${device.deviceName}`);
      console.log(`   Manufacturer: ${device.manufacturer} ${device.model}`);
      console.log(`   Total Memory: ${device.totalMemoryInMB} MB (${(device.totalMemoryInMB / 1024).toFixed(2)} GB)`);
      console.log(`   Free Memory: ${device.freeMemoryInMB} MB (~${device.freeMemoryPercentage}% free)`);
      console.log(`   User: ${device.userPrincipalName}`);
      console.log(`   Compliance: ${device.complianceState}`);
      console.log('');
    });

    saveToFile({ criteria: { minMemoryMB, maxMemoryMB }, deviceCount: devices.length, devices }, 'example6-specific-memory.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 7: Devices with Specific Free Storage (Report 26)
// ============================================================================

/**
 * Example 7: Mobile devices with specific free removable storage
 *
 * ConfigMgr Report #26: Find devices with free storage within a specific range.
 * Critical for managing storage-intensive app deployments.
 */
async function example7_SpecificFreeStorage(): Promise<void> {
  printHeader('Example 7: Devices with Specific Free Storage (Report 26)');

  try {
    const reporter = await initializeReporter();

    // Find devices with 50GB-200GB free storage
    const minStorageGB = 50;
    const maxStorageGB = 200;

    console.log(`Finding devices with ${minStorageGB}GB - ${maxStorageGB}GB free storage...\n`);
    const devices = await reporter.getDevicesWithSpecificFreeStorage(minStorageGB, maxStorageGB);

    console.log(`Found ${devices.length} devices:\n`);

    devices.slice(0, 10).forEach((device, index) => {
      console.log(`${index + 1}. ${device.deviceName}`);
      console.log(`   Manufacturer: ${device.manufacturer} ${device.model}`);
      console.log(`   Total Storage: ${device.totalStorageInGB} GB`);
      console.log(`   Free Storage: ${device.freeStorageInGB} GB (${device.freeStoragePercentage.toFixed(2)}%)`);
      console.log(`   Used Storage: ${device.usedStorageInGB} GB`);
      console.log(`   User: ${device.userPrincipalName}`);
      console.log(`   Last Sync: ${device.lastSyncDateTime}`);
      console.log('');
    });

    saveToFile({ criteria: { minStorageGB, maxStorageGB }, deviceCount: devices.length, devices }, 'example7-specific-storage.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 8: Low Memory Devices Alert (Report 28)
// ============================================================================

/**
 * Example 8: Mobile devices with low free program memory
 *
 * ConfigMgr Report #28: Identify devices with critically low free RAM.
 * Essential for proactive performance management.
 */
async function example8_LowMemoryDevices(): Promise<void> {
  printHeader('Example 8: Low Memory Devices Alert (Report 28)');

  try {
    const reporter = await initializeReporter();

    // Find devices with less than 512MB free memory
    const thresholdMB = 512;

    console.log(`Finding devices with less than ${thresholdMB}MB free memory...\n`);
    const devices = await reporter.getDevicesWithLowMemory(thresholdMB);

    console.log(`⚠️  ALERT: ${devices.length} devices with low memory detected!\n`);

    if (devices.length > 0) {
      console.log('Critical Devices:\n');
      devices.slice(0, 15).forEach((device, index) => {
        console.log(`${index + 1}. ${device.deviceName}`);
        console.log(`   Manufacturer: ${device.manufacturer} ${device.model}`);
        console.log(`   Total Memory: ${device.totalMemoryInMB} MB`);
        console.log(`   Free Memory: ${device.freeMemoryInMB} MB (⚠️ ${device.freeMemoryPercentage}%)`);
        console.log(`   User: ${device.userPrincipalName}`);
        console.log(`   Last Sync: ${device.lastSyncDateTime}`);
        console.log(`   Compliance: ${device.complianceState}`);
        console.log('   ⚡ ACTION REQUIRED: Consider memory cleanup or device upgrade');
        console.log('');
      });

      console.log('\nRecommendations:');
      console.log('1. Contact users to clear unnecessary apps/data');
      console.log('2. Review installed applications for memory usage');
      console.log('3. Consider device refresh for persistent low memory issues');
      console.log('4. Monitor these devices closely for performance problems\n');
    } else {
      console.log('✓ All devices have adequate free memory.\n');
    }

    saveToFile({ thresholdMB, deviceCount: devices.length, devices }, 'example8-low-memory-alert.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 9: Low Storage Devices Alert (Report 29)
// ============================================================================

/**
 * Example 9: Mobile devices with low free removable storage
 *
 * ConfigMgr Report #29: Identify devices with critically low free storage.
 * Prevents storage-related issues and app deployment failures.
 */
async function example9_LowStorageDevices(): Promise<void> {
  printHeader('Example 9: Low Storage Devices Alert (Report 29)');

  try {
    const reporter = await initializeReporter();

    // Find devices with less than 10GB free storage
    const thresholdGB = 10;

    console.log(`Finding devices with less than ${thresholdGB}GB free storage...\n`);
    const devices = await reporter.getDevicesWithLowStorage(thresholdGB);

    console.log(`⚠️  ALERT: ${devices.length} devices with low storage detected!\n`);

    if (devices.length > 0) {
      console.log('Critical Devices (sorted by available storage):\n');
      devices.slice(0, 15).forEach((device, index) => {
        const urgency = device.freeStorageInGB < 5 ? '🔴 CRITICAL' : '⚠️  WARNING';
        console.log(`${index + 1}. ${device.deviceName} - ${urgency}`);
        console.log(`   Manufacturer: ${device.manufacturer} ${device.model}`);
        console.log(`   Total Storage: ${device.totalStorageInGB} GB`);
        console.log(`   Free Storage: ${device.freeStorageInGB} GB (${device.freeStoragePercentage.toFixed(2)}%)`);
        console.log(`   Used Storage: ${device.usedStorageInGB} GB`);
        console.log(`   User: ${device.userPrincipalName}`);
        console.log(`   Last Sync: ${device.lastSyncDateTime}`);
        console.log('   ⚡ ACTION REQUIRED: Storage cleanup needed');
        console.log('');
      });

      // Statistics
      const criticalDevices = devices.filter(d => d.freeStorageInGB < 5);
      const warningDevices = devices.filter(d => d.freeStorageInGB >= 5);

      console.log('\nStorage Alert Summary:');
      console.log(`🔴 Critical (<5GB): ${criticalDevices.length} devices`);
      console.log(`⚠️  Warning (5-${thresholdGB}GB): ${warningDevices.length} devices`);
      console.log('');

      console.log('Recommendations:');
      console.log('1. Immediate action required for critical devices (<5GB)');
      console.log('2. Contact users to remove unnecessary files/apps');
      console.log('3. Review cloud storage options for file offloading');
      console.log('4. Consider storage optimization policies');
      console.log('5. Pause non-essential app deployments to these devices\n');
    } else {
      console.log('✓ All devices have adequate free storage.\n');
    }

    saveToFile({ thresholdGB, deviceCount: devices.length, devices }, 'example9-low-storage-alert.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 10: Comprehensive Hardware Report
// ============================================================================

/**
 * Example 10: Generate comprehensive hardware report
 *
 * Combines all reports into a single comprehensive analysis.
 * Includes summary statistics, trends, and actionable insights.
 */
async function example10_ComprehensiveReport(): Promise<void> {
  printHeader('Example 10: Comprehensive Hardware Report');

  try {
    const reporter = await initializeReporter();

    console.log('Generating comprehensive hardware report...\n');
    const report = await reporter.generateComprehensiveReport();

    console.log('=== HARDWARE INVENTORY REPORT ===\n');
    console.log(`Generated: ${report.summary.generatedAt}`);
    console.log(`Total Devices: ${report.summary.totalDevices}\n`);

    console.log('--- Operating System Distribution ---');
    report.summary.osDistribution.slice(0, 5).forEach(os => {
      console.log(`  ${os.operatingSystem} ${os.osVersion}: ${os.deviceCount} (${os.percentage.toFixed(2)}%)`);
    });

    console.log('\n--- Memory Analysis ---');
    console.log(`  Average Memory: ${report.summary.averageMemoryGB} GB`);
    console.log(`  Low Memory Devices: ${report.summary.lowMemoryDeviceCount}`);
    console.log('\n  Top Memory Configurations:');
    report.summary.topMemoryModels.slice(0, 5).forEach((model, i) => {
      console.log(`  ${i + 1}. ${model.manufacturer} ${model.model}: ${model.value} GB (${model.count} devices)`);
    });

    console.log('\n--- Storage Analysis ---');
    console.log(`  Average Storage: ${report.summary.averageStorageGB} GB`);
    console.log(`  Average Free Space: ${report.summary.averageFreeStoragePercentage}%`);
    console.log(`  Low Storage Devices: ${report.summary.lowStorageDeviceCount}`);
    console.log('\n  Top Storage Configurations:');
    report.summary.topStorageModels.slice(0, 5).forEach((model, i) => {
      console.log(`  ${i + 1}. ${model.manufacturer} ${model.model}: ${model.value} GB (${model.count} devices)`);
    });

    console.log('\n--- Alerts & Recommendations ---');
    if (report.summary.lowMemoryDeviceCount > 0) {
      console.log(`  ⚠️  ${report.summary.lowMemoryDeviceCount} devices with low memory`);
    }
    if (report.summary.lowStorageDeviceCount > 0) {
      console.log(`  ⚠️  ${report.summary.lowStorageDeviceCount} devices with low storage`);
    }
    if (report.summary.lowMemoryDeviceCount === 0 && report.summary.lowStorageDeviceCount === 0) {
      console.log('  ✓ No critical hardware issues detected');
    }

    console.log('\n--- Display Configuration Categories ---');
    report.summary.displayConfigurations.forEach(config => {
      console.log(`  ${config.displayResolution}: ${config.deviceCount} (${config.percentage.toFixed(2)}%)`);
    });

    console.log('\n');

    // Export to multiple formats
    console.log('Exporting report to multiple formats...\n');

    // JSON
    const jsonPath = await reporter.exportReport(report, {
      format: 'json',
      outputDir: './reports/hardware',
      includeTimestamp: true,
    });
    console.log(`✓ JSON report: ${jsonPath}`);

    // CSV
    const csvPath = await reporter.exportReport(report, {
      format: 'csv',
      outputDir: './reports/hardware',
      includeTimestamp: true,
    });
    console.log(`✓ CSV report: ${csvPath}`);

    // HTML
    const htmlPath = await reporter.exportReport(report, {
      format: 'html',
      outputDir: './reports/hardware',
      includeTimestamp: true,
    });
    console.log(`✓ HTML report: ${htmlPath}\n`);

    saveToFile(report, 'example10-comprehensive-report.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 11: Filtered Hardware Report (iOS Devices Only)
// ============================================================================

/**
 * Example 11: Generate filtered hardware report for iOS devices
 *
 * Demonstrates filtering capabilities for platform-specific analysis.
 */
async function example11_FilteredReport(): Promise<void> {
  printHeader('Example 11: Filtered Hardware Report (iOS Devices)');

  try {
    const reporter = await initializeReporter();

    console.log('Generating hardware report for iOS devices only...\n');

    const devices = await reporter.getAllDeviceHardwareInfo({
      operatingSystem: 'iOS',
    });

    console.log(`iOS Devices Found: ${devices.length}\n`);

    if (devices.length > 0) {
      // Memory distribution for iOS
      const memoryRanges = await reporter.getMemoryRangeSummary(devices);
      console.log('iOS Memory Distribution:');
      memoryRanges.forEach(range => {
        if (range.deviceCount > 0) {
          console.log(`  ${range.range}: ${range.deviceCount} devices (${range.percentage.toFixed(2)}%)`);
        }
      });

      // Storage distribution for iOS
      const storageRanges = await reporter.getStorageRangeSummary(devices);
      console.log('\niOS Storage Distribution:');
      storageRanges.forEach(range => {
        if (range.deviceCount > 0) {
          console.log(`  ${range.range}: ${range.deviceCount} devices (${range.percentage.toFixed(2)}%)`);
        }
      });

      // Low storage iOS devices
      const lowStorageDevices = await reporter.getDevicesWithLowStorage(10, devices);
      console.log(`\niOS Devices with Low Storage: ${lowStorageDevices.length}`);

      console.log('\n');
    } else {
      console.log('No iOS devices found in the inventory.\n');
    }

    saveToFile({ platform: 'iOS', deviceCount: devices.length, devices }, 'example11-ios-filtered.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Example 12: Custom Thresholds for Alerts
// ============================================================================

/**
 * Example 12: Custom threshold configuration for different device types
 *
 * Demonstrates advanced usage with custom thresholds based on device
 * characteristics and use cases.
 */
async function example12_CustomThresholds(): Promise<void> {
  printHeader('Example 12: Custom Thresholds for Different Device Types');

  try {
    const reporter = await initializeReporter();

    console.log('Analyzing devices with custom thresholds...\n');

    const allDevices = await reporter.getAllDeviceHardwareInfo();

    // Different thresholds for different device types
    console.log('=== MOBILE DEVICES (iOS/Android) ===');
    const mobileDevices = allDevices.filter(d =>
      d.operatingSystem.operatingSystem.includes('iOS') ||
      d.operatingSystem.operatingSystem.includes('Android')
    );

    const mobileLowMemory = await reporter.getDevicesWithLowMemory(1024, mobileDevices); // 1GB for mobile
    const mobileLowStorage = await reporter.getDevicesWithLowStorage(5, mobileDevices); // 5GB for mobile

    console.log(`Total Mobile Devices: ${mobileDevices.length}`);
    console.log(`Low Memory (<1GB free): ${mobileLowMemory.length}`);
    console.log(`Low Storage (<5GB free): ${mobileLowStorage.length}\n`);

    console.log('=== WINDOWS DEVICES ===');
    const windowsDevices = allDevices.filter(d =>
      d.operatingSystem.operatingSystem.includes('Windows')
    );

    const windowsLowMemory = await reporter.getDevicesWithLowMemory(2048, windowsDevices); // 2GB for Windows
    const windowsLowStorage = await reporter.getDevicesWithLowStorage(20, windowsDevices); // 20GB for Windows

    console.log(`Total Windows Devices: ${windowsDevices.length}`);
    console.log(`Low Memory (<2GB free): ${windowsLowMemory.length}`);
    console.log(`Low Storage (<20GB free): ${windowsLowStorage.length}\n`);

    // Summary
    const results = {
      mobile: {
        total: mobileDevices.length,
        lowMemory: mobileLowMemory.length,
        lowStorage: mobileLowStorage.length,
        thresholds: { memoryMB: 1024, storageGB: 5 },
      },
      windows: {
        total: windowsDevices.length,
        lowMemory: windowsLowMemory.length,
        lowStorage: windowsLowStorage.length,
        thresholds: { memoryMB: 2048, storageGB: 20 },
      },
    };

    console.log('=== ALERT SUMMARY ===');
    console.log(`Total Devices with Issues: ${mobileLowMemory.length + mobileLowStorage.length + windowsLowMemory.length + windowsLowStorage.length}`);
    console.log('\nPlatform-Specific Alerts:');
    console.log(`  Mobile: ${mobileLowMemory.length + mobileLowStorage.length} issues`);
    console.log(`  Windows: ${windowsLowMemory.length + windowsLowStorage.length} issues\n`);

    saveToFile(results, 'example12-custom-thresholds.json');
  } catch (error) {
    console.error('Error:', error);
  }
}

// ============================================================================
// Main Execution
// ============================================================================

/**
 * Run all examples
 */
async function runAllExamples(): Promise<void> {
  console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
  console.log('║                                                                          ║');
  console.log('║         Device Hardware Reports - Comprehensive Examples Suite          ║');
  console.log('║                                                                          ║');
  console.log('║  Replicating Configuration Manager Hardware Reports for Intune          ║');
  console.log('║                                                                          ║');
  console.log('╚══════════════════════════════════════════════════════════════════════════╝\n');

  const examples = [
    { name: 'Example 1: Get All Device Hardware', fn: example1_GetAllDeviceHardware },
    { name: 'Example 2: Display Configuration Summary', fn: example2_DisplayConfigurationSummary },
    { name: 'Example 3: OS Distribution Report', fn: example3_OSDistribution },
    { name: 'Example 4: Memory Range Summary', fn: example4_MemoryRanges },
    { name: 'Example 5: Storage Range Summary', fn: example5_StorageRanges },
    { name: 'Example 6: Specific Free Memory', fn: example6_SpecificFreeMemory },
    { name: 'Example 7: Specific Free Storage', fn: example7_SpecificFreeStorage },
    { name: 'Example 8: Low Memory Alert', fn: example8_LowMemoryDevices },
    { name: 'Example 9: Low Storage Alert', fn: example9_LowStorageDevices },
    { name: 'Example 10: Comprehensive Report', fn: example10_ComprehensiveReport },
    { name: 'Example 11: Filtered Report (iOS)', fn: example11_FilteredReport },
    { name: 'Example 12: Custom Thresholds', fn: example12_CustomThresholds },
  ];

  for (const example of examples) {
    try {
      await example.fn();
    } catch (error) {
      console.error(`Error in ${example.name}:`, error);
    }
  }

  console.log('\n╔══════════════════════════════════════════════════════════════════════════╗');
  console.log('║                     All Examples Completed!                              ║');
  console.log('╚══════════════════════════════════════════════════════════════════════════╝\n');
}

/**
 * Run a specific example
 */
async function runExample(exampleNumber: number): Promise<void> {
  const examples = [
    example1_GetAllDeviceHardware,
    example2_DisplayConfigurationSummary,
    example3_OSDistribution,
    example4_MemoryRanges,
    example5_StorageRanges,
    example6_SpecificFreeMemory,
    example7_SpecificFreeStorage,
    example8_LowMemoryDevices,
    example9_LowStorageDevices,
    example10_ComprehensiveReport,
    example11_FilteredReport,
    example12_CustomThresholds,
  ];

  if (exampleNumber < 1 || exampleNumber > examples.length) {
    console.error(`Invalid example number. Choose between 1 and ${examples.length}`);
    return;
  }

  await examples[exampleNumber - 1]();
}

// ============================================================================
// CLI Interface
// ============================================================================

if (require.main === module) {
  const args = process.argv.slice(2);

  if (args.length === 0) {
    // Run all examples
    runAllExamples().catch(console.error);
  } else {
    // Run specific example
    const exampleNum = parseInt(args[0]);
    runExample(exampleNum).catch(console.error);
  }
}

// ============================================================================
// Exports
// ============================================================================

export {
  example1_GetAllDeviceHardware,
  example2_DisplayConfigurationSummary,
  example3_OSDistribution,
  example4_MemoryRanges,
  example5_StorageRanges,
  example6_SpecificFreeMemory,
  example7_SpecificFreeStorage,
  example8_LowMemoryDevices,
  example9_LowStorageDevices,
  example10_ComprehensiveReport,
  example11_FilteredReport,
  example12_CustomThresholds,
  runAllExamples,
  runExample,
};
