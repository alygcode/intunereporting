/**
 * Mobile Device Inventory Reports - Usage Examples
 *
 * This file demonstrates various usage scenarios for the Mobile Device Inventory Reports module.
 * These examples show how to use the ConfigMgr-style mobile device reports with Intune.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import {
  MobileDeviceInventoryReports,
  DeviceOwnershipType,
  ManagementAgentType,
  CommunicationStatus,
} from '../configmgr/mobile-device-inventory-reports';
import { AppConfig } from '../../types';

/**
 * Example 1: Get all corporate-owned mobile devices
 *
 * This example demonstrates how to retrieve all mobile devices that are
 * corporate/company-owned, which replicates ConfigMgr Report #1.
 */
export async function example1_GetCorporateOwnedDevices(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 1: All Corporate-Owned Mobile Devices ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get all corporate-owned devices
    console.log('Fetching corporate-owned mobile devices...');
    const corporateDevices = await report.getAllCorporateOwnedDevices();

    console.log(`\nTotal corporate devices: ${corporateDevices.length}\n`);

    // Display summary by operating system
    const osCounts = new Map<string, number>();
    corporateDevices.forEach(device => {
      osCounts.set(device.operatingSystem, (osCounts.get(device.operatingSystem) || 0) + 1);
    });

    console.log('Corporate Devices by Operating System:');
    osCounts.forEach((count, os) => {
      console.log(`  ${os}: ${count}`);
    });

    // Display sample devices
    console.log('\nSample Corporate Devices:');
    corporateDevices.slice(0, 5).forEach((device, index) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   OS: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`   Manufacturer: ${device.manufacturer || 'Unknown'}`);
      console.log(`   Model: ${device.model || 'Unknown'}`);
      console.log(`   Serial Number: ${device.serialNumber || 'N/A'}`);
      console.log(`   User: ${device.userPrincipalName || 'N/A'}`);
      console.log(`   Last Sync: ${device.lastSyncDateTime?.toLocaleString() || 'Never'}`);
      console.log(`   Compliance: ${device.complianceState || 'Unknown'}`);
      console.log(`   Enrollment Profile: ${device.enrollmentProfileName || 'N/A'}`);
    });

    // Filter by specific OS
    console.log('\n\nFetching iOS corporate devices only...');
    const iOSCorporateDevices = await report.getAllCorporateOwnedDevices({
      operatingSystem: 'iOS',
    });
    console.log(`iOS Corporate Devices: ${iOSCorporateDevices.length}`);

    console.log('\n✓ Corporate-owned devices retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching corporate-owned devices:', error);
  }
}

/**
 * Example 2: Get all mobile device clients (excluding Exchange connector)
 *
 * This example demonstrates how to retrieve all mobile device clients managed
 * by Intune MDM, excluding Exchange ActiveSync connector devices.
 * This replicates ConfigMgr Report #2.
 */
export async function example2_GetAllMobileDeviceClients(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 2: All Mobile Device Clients ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get all mobile device clients (excluding Exchange connector)
    console.log('Fetching all mobile device clients...');
    const mobileClients = await report.getAllMobileDeviceClients();

    console.log(`\nTotal mobile device clients: ${mobileClients.length}\n`);

    // Group by ownership type
    const ownershipCounts = new Map<string, number>();
    mobileClients.forEach(device => {
      ownershipCounts.set(device.ownerType, (ownershipCounts.get(device.ownerType) || 0) + 1);
    });

    console.log('Devices by Ownership Type:');
    ownershipCounts.forEach((count, type) => {
      console.log(`  ${type}: ${count}`);
    });

    // Group by management agent
    const agentCounts = new Map<string, number>();
    mobileClients.forEach(device => {
      agentCounts.set(device.managementAgent, (agentCounts.get(device.managementAgent) || 0) + 1);
    });

    console.log('\nDevices by Management Agent:');
    agentCounts.forEach((count, agent) => {
      console.log(`  ${agent}: ${count}`);
    });

    // Display devices by platform
    console.log('\n\nDevices by Platform:');
    const platforms = ['iOS', 'Android', 'Windows'];
    for (const platform of platforms) {
      const platformDevices = mobileClients.filter(d => d.operatingSystem === platform);
      console.log(`\n${platform}: ${platformDevices.length} devices`);
      platformDevices.slice(0, 3).forEach((device, index) => {
        console.log(`  ${index + 1}. ${device.deviceName} (${device.osVersion})`);
      });
    }

    // Filter by compliance state
    console.log('\n\nCompliance Status:');
    const compliantDevices = mobileClients.filter(d => d.complianceState === 'compliant');
    const nonCompliantDevices = mobileClients.filter(d => d.complianceState === 'noncompliant');
    console.log(`  Compliant: ${compliantDevices.length}`);
    console.log(`  Non-Compliant: ${nonCompliantDevices.length}`);

    console.log('\n✓ Mobile device clients retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching mobile device clients:', error);
  }
}

/**
 * Example 3: Get mobile device client information with communication status
 *
 * This example demonstrates how to retrieve mobile device client information
 * with management point communication status, which replicates ConfigMgr Report #20.
 */
export async function example3_GetDeviceClientInformation(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 3: Mobile Device Client Information ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get device client information with communication status
    console.log('Fetching device client information...');
    const clientInfo = await report.getMobileDeviceClientInformation();

    console.log(`\nTotal devices: ${clientInfo.length}\n`);

    // Group by communication status
    const statusCounts = new Map<string, number>();
    clientInfo.forEach(device => {
      statusCounts.set(device.communicationStatus, (statusCounts.get(device.communicationStatus) || 0) + 1);
    });

    console.log('Devices by Communication Status:');
    statusCounts.forEach((count, status) => {
      console.log(`  ${status}: ${count}`);
    });

    // Show active vs inactive devices
    const activeDevices = clientInfo.filter(d => d.isActive);
    const inactiveDevices = clientInfo.filter(d => !d.isActive);

    console.log('\nDevice Activity:');
    console.log(`  Active (synced within 30 days): ${activeDevices.length}`);
    console.log(`  Inactive (not synced in 30+ days): ${inactiveDevices.length}`);

    // Display healthy devices
    console.log('\n\nHealthy Devices (synced within 24 hours):');
    const healthyDevices = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.HEALTHY);
    healthyDevices.slice(0, 5).forEach((device, index) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   User: ${device.userPrincipalName || 'N/A'}`);
      console.log(`   OS: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`   Last Contact: ${device.lastContactDateTime?.toLocaleString() || 'Never'}`);
      console.log(`   Days Since Contact: ${device.daysSinceLastContact || 'N/A'}`);
      console.log(`   Status: ${device.communicationStatus}`);
    });

    // Display devices with communication issues
    console.log('\n\nDevices with Communication Issues (7+ days since last sync):');
    const criticalDevices = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.CRITICAL);
    criticalDevices.slice(0, 5).forEach((device, index) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   User: ${device.userPrincipalName || 'N/A'}`);
      console.log(`   Last Contact: ${device.lastContactDateTime?.toLocaleString() || 'Never'}`);
      console.log(`   Days Since Contact: ${device.daysSinceLastContact || 'N/A'}`);
      console.log(`   Status: ${device.communicationStatus}`);
      console.log(`   ⚠️ Action Required: Check device connectivity`);
    });

    console.log('\n✓ Device client information retrieved successfully!\n');
  } catch (error) {
    console.error('Error fetching device client information:', error);
  }
}

/**
 * Example 4: Get mobile devices grouped by operating system
 *
 * This example demonstrates how to retrieve mobile devices grouped by
 * operating system with version distribution, which replicates ConfigMgr Report #22.
 */
export async function example4_GetDevicesByOperatingSystem(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 4: Mobile Devices by Operating System ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get OS distribution
    console.log('Analyzing OS distribution...');
    const osDistribution = await report.getMobileDevicesByOperatingSystem();

    console.log(`\nOperating Systems Found: ${osDistribution.length}\n`);

    // Display OS distribution
    console.log('Operating System Distribution:');
    osDistribution.forEach((os, index) => {
      console.log(`\n${index + 1}. ${os.operatingSystem}`);
      console.log(`   Device Count: ${os.deviceCount}`);
      console.log(`   Percentage: ${os.percentage.toFixed(2)}%`);
      console.log(`   Versions Found: ${os.versions.length}`);

      // Display version breakdown
      console.log('\n   Version Distribution:');
      os.versions.forEach((version, vIndex) => {
        console.log(`     ${vIndex + 1}. Version ${version.version}`);
        console.log(`        Devices: ${version.deviceCount}`);
        console.log(`        Percentage: ${version.percentage.toFixed(2)}% of ${os.operatingSystem} devices`);
      });
    });

    // Create a visual chart (text-based)
    console.log('\n\nOS Distribution Chart:');
    const maxBarLength = 50;
    const maxDeviceCount = Math.max(...osDistribution.map(os => os.deviceCount));
    osDistribution.forEach(os => {
      const barLength = Math.round((os.deviceCount / maxDeviceCount) * maxBarLength);
      const bar = '█'.repeat(barLength);
      console.log(`${os.operatingSystem.padEnd(15)} ${bar} ${os.deviceCount} (${os.percentage.toFixed(1)}%)`);
    });

    // Filter by specific OS
    console.log('\n\nGetting detailed Android distribution...');
    const androidDistribution = await report.getMobileDevicesByOperatingSystem({
      operatingSystem: 'Android',
    });

    if (androidDistribution.length > 0) {
      const android = androidDistribution[0];
      console.log(`\nAndroid Version Distribution:`);
      android.versions.forEach((version, index) => {
        console.log(`  ${index + 1}. Android ${version.version}: ${version.deviceCount} devices`);
      });
    }

    console.log('\n✓ OS distribution analysis completed successfully!\n');
  } catch (error) {
    console.error('Error analyzing OS distribution:', error);
  }
}

/**
 * Example 5: Filter and sort devices with advanced options
 *
 * This example demonstrates advanced filtering and sorting capabilities.
 */
export async function example5_AdvancedFilteringAndSorting(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 5: Advanced Filtering and Sorting ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get iOS devices sorted by last sync date
    console.log('Fetching iOS devices sorted by last sync date...');
    const iosDevices = await report.getAllMobileDeviceClients({
      operatingSystem: 'iOS',
      field: 'lastSyncDateTime',
      direction: 'desc',
    });

    console.log(`\niOS Devices: ${iosDevices.length}\n`);
    console.log('Most Recently Synced iOS Devices:');
    iosDevices.slice(0, 5).forEach((device, index) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   Model: ${device.model || 'Unknown'}`);
      console.log(`   OS Version: ${device.osVersion}`);
      console.log(`   Last Sync: ${device.lastSyncDateTime?.toLocaleString() || 'Never'}`);
    });

    // Get supervised devices only
    console.log('\n\nFetching supervised devices...');
    const supervisedDevices = await report.getAllMobileDeviceClients({
      isSupervised: true,
    });
    console.log(`Supervised Devices: ${supervisedDevices.length}`);

    // Get compliant corporate devices
    console.log('\nFetching compliant corporate devices...');
    const compliantCorporate = await report.getAllCorporateOwnedDevices({
      complianceState: 'compliant',
    });
    console.log(`Compliant Corporate Devices: ${compliantCorporate.length}`);

    // Get devices by specific management agent
    console.log('\nFetching devices managed by Intune MDM...');
    const intuneMdmDevices = await report.getDevicesByManagementAgent(ManagementAgentType.INTUNE_MDM);
    console.log(`Intune MDM Managed Devices: ${intuneMdmDevices.length}`);

    // Get devices sorted by enrollment date
    console.log('\n\nFetching recently enrolled devices...');
    const recentlyEnrolled = await report.getAllMobileDeviceClients({
      field: 'enrolledDateTime',
      direction: 'desc',
    });

    console.log('\nMost Recently Enrolled Devices:');
    recentlyEnrolled.slice(0, 5).forEach((device, index) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   OS: ${device.operatingSystem}`);
      console.log(`   User: ${device.userPrincipalName || 'N/A'}`);
      console.log(`   Enrolled: ${device.enrolledDateTime?.toLocaleString() || 'Unknown'}`);
    });

    console.log('\n✓ Advanced filtering and sorting completed successfully!\n');
  } catch (error) {
    console.error('Error with advanced filtering:', error);
  }
}

/**
 * Example 6: Identify inactive and problematic devices
 *
 * This example demonstrates how to identify devices that may need attention.
 */
export async function example6_IdentifyProblematicDevices(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 6: Identify Problematic Devices ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get all devices
    console.log('Analyzing all devices for issues...');
    const allDevices = await report.getAllMobileDevices();

    // Get inactive devices (not synced in 30+ days)
    const inactiveDevices = await report.getInactiveDevices(allDevices, 30);
    console.log(`\nInactive Devices (30+ days): ${inactiveDevices.length}`);

    if (inactiveDevices.length > 0) {
      console.log('\nTop 10 Longest Inactive Devices:');
      inactiveDevices
        .sort((a, b) => {
          const daysA = a.lastSyncDateTime ? report['calculateDaysSince'](a.lastSyncDateTime) : 999999;
          const daysB = b.lastSyncDateTime ? report['calculateDaysSince'](b.lastSyncDateTime) : 999999;
          return daysB - daysA;
        })
        .slice(0, 10)
        .forEach((device, index) => {
          const daysSinceSync = device.lastSyncDateTime
            ? Math.ceil((new Date().getTime() - device.lastSyncDateTime.getTime()) / (1000 * 60 * 60 * 24))
            : 'Never';
          console.log(`\n${index + 1}. ${device.deviceName}`);
          console.log(`   User: ${device.userPrincipalName || 'N/A'}`);
          console.log(`   Last Sync: ${device.lastSyncDateTime?.toLocaleString() || 'Never'}`);
          console.log(`   Days Since Sync: ${daysSinceSync}`);
          console.log(`   ⚠️ Consider removing from management`);
        });
    }

    // Get very active devices (synced in last 7 days)
    const activeDevices = report.getActiveDevices(allDevices, 7);
    console.log(`\n\nActive Devices (last 7 days): ${activeDevices.length}`);

    // Get device client info to identify communication issues
    const clientInfo = await report.getMobileDeviceClientInformation();

    const warningDevices = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.WARNING);
    const criticalDevices = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.CRITICAL);

    console.log('\n\nDevices Needing Attention:');
    console.log(`  Warning Status (1-7 days): ${warningDevices.length}`);
    console.log(`  Critical Status (7+ days): ${criticalDevices.length}`);

    if (criticalDevices.length > 0) {
      console.log('\n\nCritical Devices (Immediate Attention Required):');
      criticalDevices.slice(0, 10).forEach((device, index) => {
        console.log(`\n${index + 1}. ${device.deviceName}`);
        console.log(`   User: ${device.userPrincipalName || 'N/A'}`);
        console.log(`   OS: ${device.operatingSystem}`);
        console.log(`   Days Since Contact: ${device.daysSinceLastContact}`);
        console.log(`   🚨 Urgent: Device may be lost, stolen, or turned off`);
      });
    }

    // Summary statistics
    console.log('\n\nDevice Health Summary:');
    console.log(`  Total Devices: ${allDevices.length}`);
    console.log(`  Healthy: ${clientInfo.filter(d => d.communicationStatus === CommunicationStatus.HEALTHY).length}`);
    console.log(`  Warning: ${warningDevices.length}`);
    console.log(`  Critical: ${criticalDevices.length}`);
    console.log(`  Inactive (30+ days): ${inactiveDevices.length}`);

    const healthPercentage = ((clientInfo.filter(d => d.communicationStatus === CommunicationStatus.HEALTHY).length / allDevices.length) * 100).toFixed(2);
    console.log(`\nOverall Health: ${healthPercentage}%`);

    console.log('\n✓ Problematic device analysis completed successfully!\n');
  } catch (error) {
    console.error('Error analyzing problematic devices:', error);
  }
}

/**
 * Example 7: Export reports to different formats
 *
 * This example demonstrates how to export mobile device inventory reports
 * to JSON, CSV, and HTML formats.
 */
export async function example7_ExportReports(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 7: Export Mobile Device Reports ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);
    const outputDir = './reports/output';

    // Export complete mobile device inventory as JSON
    console.log('Exporting complete mobile device inventory as JSON...');
    const reportData = await report.execute();
    const jsonPath = await report.exportReport(reportData, {
      format: 'json',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ JSON report saved to: ${jsonPath}`);

    // Export corporate devices as CSV
    console.log('\nExporting corporate devices as CSV...');
    const corporateDevices = await report.getAllCorporateOwnedDevices();
    const corporateReportData = {
      metadata: {
        reportName: 'corporate-mobile-devices',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: corporateDevices.length,
      },
      data: corporateDevices,
      summary: {
        totalDevices: corporateDevices.length,
      },
    };
    const csvPath = await report.exportReport(corporateReportData, {
      format: 'csv',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ CSV report saved to: ${csvPath}`);

    // Export OS distribution as HTML
    console.log('\nExporting OS distribution as HTML...');
    const osDistribution = await report.getMobileDevicesByOperatingSystem();
    const osReportData = {
      metadata: {
        reportName: 'mobile-devices-os-distribution',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: osDistribution.length,
      },
      data: osDistribution,
      summary: {
        totalOperatingSystems: osDistribution.length,
        totalDevices: osDistribution.reduce((sum, os) => sum + os.deviceCount, 0),
      },
    };
    const htmlPath = await report.exportReport(osReportData, {
      format: 'html',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ HTML report saved to: ${htmlPath}`);

    // Export device client communication status
    console.log('\nExporting device client communication status...');
    const clientInfo = await report.getMobileDeviceClientInformation();
    const clientReportData = {
      metadata: {
        reportName: 'mobile-device-client-communication',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: clientInfo.length,
      },
      data: clientInfo,
      summary: {
        totalDevices: clientInfo.length,
        healthyDevices: clientInfo.filter(d => d.communicationStatus === CommunicationStatus.HEALTHY).length,
        warningDevices: clientInfo.filter(d => d.communicationStatus === CommunicationStatus.WARNING).length,
        criticalDevices: clientInfo.filter(d => d.communicationStatus === CommunicationStatus.CRITICAL).length,
      },
    };
    const commPath = await report.exportReport(clientReportData, {
      format: 'csv',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ Communication status report saved to: ${commPath}`);

    console.log('\n✓ All reports exported successfully!\n');
  } catch (error) {
    console.error('Error exporting reports:', error);
  }
}

/**
 * Example 8: Pagination and large dataset handling
 *
 * This example demonstrates how to handle large datasets using pagination.
 */
export async function example8_PaginationExample(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 8: Pagination and Large Dataset Handling ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Get all devices
    console.log('Fetching all devices...');
    const allDevices = await report.getAllMobileDevices();
    console.log(`Total devices: ${allDevices.length}`);

    // Apply pagination
    const pageSize = 10;
    const totalPages = Math.ceil(allDevices.length / pageSize);

    console.log(`\nPaginating results: ${pageSize} devices per page, ${totalPages} total pages\n`);

    // Show first page
    const page1 = report.paginate(allDevices, { page: 1, pageSize });
    console.log('Page 1:');
    console.log(`  Current Page: ${page1.pagination.currentPage}`);
    console.log(`  Page Size: ${page1.pagination.pageSize}`);
    console.log(`  Total Pages: ${page1.pagination.totalPages}`);
    console.log(`  Total Records: ${page1.pagination.totalRecords}`);
    console.log(`  Has Next Page: ${page1.pagination.hasNextPage}`);
    console.log(`  Has Previous Page: ${page1.pagination.hasPreviousPage}`);

    console.log('\n  Devices on this page:');
    page1.data.forEach((device, index) => {
      console.log(`    ${index + 1}. ${device.deviceName} (${device.operatingSystem})`);
    });

    // Show last page
    if (totalPages > 1) {
      const lastPage = report.paginate(allDevices, { page: totalPages, pageSize });
      console.log(`\n\nPage ${totalPages} (Last Page):`);
      console.log(`  Devices on this page: ${lastPage.data.length}`);
      lastPage.data.forEach((device, index) => {
        console.log(`    ${index + 1}. ${device.deviceName} (${device.operatingSystem})`);
      });
    }

    // Demonstrate iterating through all pages
    console.log('\n\nIterating through all pages...');
    for (let pageNum = 1; pageNum <= Math.min(3, totalPages); pageNum++) {
      const page = report.paginate(allDevices, { page: pageNum, pageSize });
      console.log(`\nPage ${pageNum}: ${page.data.length} devices`);
    }

    console.log('\n✓ Pagination example completed successfully!\n');
  } catch (error) {
    console.error('Error with pagination:', error);
  }
}

/**
 * Example 9: Generate convenience reports
 *
 * This example demonstrates using the convenience functions to quickly
 * generate specific reports.
 */
export async function example9_ConvenienceReports(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 9: Convenience Report Functions ===\n');

  try {
    // Import convenience functions
    const {
      generateCorporateDevicesReport,
      generateOSDistributionReport,
      generateDeviceClientCommunicationReport,
    } = await import('../configmgr/mobile-device-inventory-reports');

    // Generate corporate devices report
    console.log('Generating corporate devices report...');
    const corporateReport = await generateCorporateDevicesReport(graphClient, config);
    console.log(`\nCorporate Devices Report:`);
    console.log(`  Total Corporate Devices: ${corporateReport.summary.totalCorporateDevices}`);
    console.log(`  Operating Systems:`, corporateReport.summary.byOS);

    // Generate OS distribution report
    console.log('\n\nGenerating OS distribution report...');
    const osReport = await generateOSDistributionReport(graphClient, config);
    console.log(`\nOS Distribution Report:`);
    console.log(`  Total Operating Systems: ${osReport.summary.totalOperatingSystems}`);
    console.log(`  Total Devices: ${osReport.summary.totalDevices}`);

    // Generate device client communication report
    console.log('\n\nGenerating device client communication report...');
    const commReport = await generateDeviceClientCommunicationReport(graphClient, config);
    console.log(`\nDevice Client Communication Report:`);
    console.log(`  Total Devices: ${commReport.summary.totalDevices}`);
    console.log(`  Healthy Devices: ${commReport.summary.healthyDevices}`);
    console.log(`  Warning Devices: ${commReport.summary.warningDevices}`);
    console.log(`  Critical Devices: ${commReport.summary.criticalDevices}`);
    console.log(`  Active Devices: ${commReport.summary.activeDevices}`);
    console.log(`  Inactive Devices: ${commReport.summary.inactiveDevices}`);

    console.log('\n✓ Convenience reports generated successfully!\n');
  } catch (error) {
    console.error('Error generating convenience reports:', error);
  }
}

/**
 * Example 10: Complete workflow - Analysis, filtering, and export
 *
 * This example demonstrates a complete workflow from data collection to export.
 */
export async function example10_CompleteWorkflow(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('=== Example 10: Complete Workflow ===\n');

  try {
    const report = new MobileDeviceInventoryReports(graphClient, config);

    // Step 1: Collect all device data
    console.log('Step 1: Collecting all device data...');
    const allDevices = await report.getAllMobileDevices();
    console.log(`✓ Collected ${allDevices.length} devices\n`);

    // Step 2: Analyze OS distribution
    console.log('Step 2: Analyzing OS distribution...');
    const osDistribution = await report.getMobileDevicesByOperatingSystem();
    console.log(`✓ Found ${osDistribution.length} operating systems\n`);

    // Step 3: Identify corporate vs personal devices
    console.log('Step 3: Identifying corporate vs personal devices...');
    const corporateDevices = await report.getAllCorporateOwnedDevices();
    const personalDevices = await report.getDevicesByOwnerType(DeviceOwnershipType.PERSONAL);
    console.log(`✓ Corporate: ${corporateDevices.length}, Personal: ${personalDevices.length}\n`);

    // Step 4: Check device health
    console.log('Step 4: Checking device health...');
    const clientInfo = await report.getMobileDeviceClientInformation();
    const healthyCount = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.HEALTHY).length;
    const issueCount = clientInfo.length - healthyCount;
    console.log(`✓ Healthy: ${healthyCount}, Issues: ${issueCount}\n`);

    // Step 5: Generate summary report
    console.log('Step 5: Generating summary report...');
    const summaryData = {
      metadata: {
        reportName: 'mobile-device-inventory-summary',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Intune Reporting Dashboard',
        recordCount: allDevices.length,
      },
      data: allDevices,
      summary: {
        totalDevices: allDevices.length,
        corporateDevices: corporateDevices.length,
        personalDevices: personalDevices.length,
        healthyDevices: healthyCount,
        devicesWithIssues: issueCount,
        operatingSystems: osDistribution.map(os => ({
          name: os.operatingSystem,
          count: os.deviceCount,
          percentage: os.percentage,
        })),
      },
    };
    console.log(`✓ Summary report generated\n`);

    // Step 6: Export to all formats
    console.log('Step 6: Exporting to all formats...');
    const outputDir = './reports/output';

    const jsonPath = await report.exportReport(summaryData, {
      format: 'json',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ JSON: ${jsonPath}`);

    const csvPath = await report.exportReport(summaryData, {
      format: 'csv',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ CSV: ${csvPath}`);

    const htmlPath = await report.exportReport(summaryData, {
      format: 'html',
      outputDirectory: outputDir,
      includeTimestamp: true,
    });
    console.log(`✓ HTML: ${htmlPath}`);

    // Step 7: Display final summary
    console.log('\n\n=== Final Summary ===');
    console.log(`Total Mobile Devices: ${summaryData.summary.totalDevices}`);
    console.log(`Corporate Devices: ${summaryData.summary.corporateDevices}`);
    console.log(`Personal Devices: ${summaryData.summary.personalDevices}`);
    console.log(`Healthy Devices: ${summaryData.summary.healthyDevices}`);
    console.log(`Devices with Issues: ${summaryData.summary.devicesWithIssues}`);
    console.log('\nOperating System Distribution:');
    summaryData.summary.operatingSystems.forEach(os => {
      console.log(`  ${os.name}: ${os.count} (${os.percentage.toFixed(2)}%)`);
    });

    console.log('\n✓ Complete workflow finished successfully!\n');
  } catch (error) {
    console.error('Error in complete workflow:', error);
  }
}

/**
 * Run all examples
 */
export async function runAllExamples(
  graphClient: Client,
  config: AppConfig
): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('MOBILE DEVICE INVENTORY REPORTS - ALL EXAMPLES');
  console.log('='.repeat(70) + '\n');

  await example1_GetCorporateOwnedDevices(graphClient, config);
  await example2_GetAllMobileDeviceClients(graphClient, config);
  await example3_GetDeviceClientInformation(graphClient, config);
  await example4_GetDevicesByOperatingSystem(graphClient, config);
  await example5_AdvancedFilteringAndSorting(graphClient, config);
  await example6_IdentifyProblematicDevices(graphClient, config);
  await example7_ExportReports(graphClient, config);
  await example8_PaginationExample(graphClient, config);
  await example9_ConvenienceReports(graphClient, config);
  await example10_CompleteWorkflow(graphClient, config);

  console.log('\n' + '='.repeat(70));
  console.log('ALL EXAMPLES COMPLETED SUCCESSFULLY');
  console.log('='.repeat(70) + '\n');
}

// Export all examples
export default {
  example1_GetCorporateOwnedDevices,
  example2_GetAllMobileDeviceClients,
  example3_GetDeviceClientInformation,
  example4_GetDevicesByOperatingSystem,
  example5_AdvancedFilteringAndSorting,
  example6_IdentifyProblematicDevices,
  example7_ExportReports,
  example8_PaginationExample,
  example9_ConvenienceReports,
  example10_CompleteWorkflow,
  runAllExamples,
};
