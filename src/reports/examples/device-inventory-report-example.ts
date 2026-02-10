/**
 * Device Inventory Report - Usage Examples
 *
 * This file demonstrates various ways to use the comprehensive Device Inventory Report module
 * for Microsoft Intune. It shows different configuration options and use cases.
 *
 * Prerequisites:
 * - Azure AD app registration with appropriate Microsoft Graph permissions
 * - Environment variables configured (see .env.example)
 *
 * Required permissions:
 * - DeviceManagementManagedDevices.Read.All
 * - DeviceManagementApps.Read.All
 * - DeviceManagementConfiguration.Read.All
 *
 * @module device-inventory-report-example
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import * as dotenv from 'dotenv';
import * as path from 'path';
import {
  DeviceInventoryReport,
  DeviceInventoryOptions,
  generateBasicDeviceInventory,
  generateComprehensiveDeviceInventory,
} from '../device-inventory-report';
import { AppConfig } from '../../types';

// Load environment variables
dotenv.config();

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Create authenticated Microsoft Graph client
 */
async function createGraphClient(): Promise<Client> {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      'Missing required environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET'
    );
  }

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });

  return Client.initWithMiddleware({ authProvider });
}

/**
 * Create mock app configuration
 */
function createMockConfig(): AppConfig {
  return {
    authentication: {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET,
      authMethod: 'clientSecret',
    },
    reports: {
      enabled: ['device-inventory'],
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
      file: './logs/intune-reports.log',
      console: true,
      maxSize: '10m',
      maxFiles: 5,
    },
  };
}

// ============================================================================
// Example 1: Basic Device Inventory
// ============================================================================

/**
 * Generate a basic device inventory report with minimal options
 */
async function example1_BasicInventory(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 1: Basic Device Inventory');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    console.log('Generating basic device inventory report...');

    const reportData = await generateBasicDeviceInventory(graphClient, config);

    console.log('\nReport Summary:');
    console.log(`- Total Devices: ${reportData.metadata.recordCount}`);
    console.log(`- Generated At: ${reportData.metadata.generatedAt}`);

    if (reportData.summary) {
      console.log('\nBy Operating System:');
      Object.entries(reportData.summary.byOS as Record<string, number>).forEach(([os, count]) => {
        console.log(`  ${os}: ${count}`);
      });

      console.log('\nBy Compliance:');
      Object.entries(reportData.summary.byCompliance as Record<string, number>).forEach(
        ([state, count]) => {
          console.log(`  ${state}: ${count}`);
        }
      );
    }

    console.log('\nFirst 5 devices:');
    reportData.data.slice(0, 5).forEach((device: any, index: number) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   OS: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`   User: ${device.userDisplayName}`);
      console.log(`   Compliance: ${device.complianceState}`);
    });

    console.log('\n✓ Example 1 completed successfully!');
  } catch (error) {
    console.error('Example 1 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 2: Comprehensive Inventory with Hardware Details
// ============================================================================

/**
 * Generate comprehensive inventory including hardware details
 */
async function example2_ComprehensiveWithHardware(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 2: Comprehensive Inventory with Hardware Details');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const options: DeviceInventoryOptions = {
      includeHardware: true,
      includeSoftware: false,
      includeHealthAttestation: false,
      outputFormats: ['json', 'html'],
      outputDirectory: './reports/device-inventory',
    };

    console.log('Generating comprehensive device inventory with hardware details...');

    const report = new DeviceInventoryReport(graphClient, config);
    const reportData = await report.executeWithOptions(options);

    DeviceInventoryReport.printSummary(reportData);

    if (reportData.hardwareInventory && reportData.hardwareInventory.length > 0) {
      console.log('Hardware Inventory Sample (first 3 devices):');
      reportData.hardwareInventory.slice(0, 3).forEach((hw, index) => {
        console.log(`\n${index + 1}. ${hw.deviceName}`);
        console.log(`   Manufacturer: ${hw.manufacturer || 'N/A'}`);
        console.log(`   Model: ${hw.model || 'N/A'}`);
        console.log(`   Serial: ${hw.serialNumber || 'N/A'}`);
        console.log(`   Storage: ${hw.totalStorageGB || 'N/A'} GB (${hw.usedStoragePercentage || 'N/A'}% used)`);
        console.log(`   Memory: ${hw.totalMemoryGB || 'N/A'} GB`);
        console.log(`   Encrypted: ${hw.isEncrypted ? 'Yes' : 'No'}`);
      });
    }

    console.log('\n✓ Example 2 completed successfully!');
  } catch (error) {
    console.error('Example 2 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 3: Software Inventory Report
// ============================================================================

/**
 * Generate inventory with software/application details
 */
async function example3_SoftwareInventory(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 3: Software Inventory Report');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const options: DeviceInventoryOptions = {
      includeHardware: false,
      includeSoftware: true,
      maxDevicesForSoftware: 10, // Limit to 10 devices for demonstration
      includeHealthAttestation: false,
      outputFormats: ['json', 'csv'],
      outputDirectory: './reports/software-inventory',
    };

    console.log('Generating software inventory (limited to 10 devices)...');
    console.log('Note: Software inventory can be time-consuming for large device counts.');

    const report = new DeviceInventoryReport(graphClient, config);
    const reportData = await report.executeWithOptions(options);

    console.log('\nSoftware Inventory Summary:');
    console.log(`- Devices analyzed: ${reportData.softwareInventory?.length || 0}`);

    if (reportData.softwareInventory && reportData.softwareInventory.length > 0) {
      const totalApps = reportData.softwareInventory.reduce(
        (sum, si) => sum + si.totalApps,
        0
      );
      console.log(`- Total applications detected: ${totalApps}`);

      console.log('\nSample Device Software (first device):');
      const firstDevice = reportData.softwareInventory[0];
      console.log(`Device: ${firstDevice.deviceName}`);
      console.log(`OS: ${firstDevice.operatingSystem} ${firstDevice.osVersion}`);
      console.log(`Total Apps: ${firstDevice.totalApps}`);

      if (firstDevice.applications.length > 0) {
        console.log('\nTop 10 Applications:');
        firstDevice.applications.slice(0, 10).forEach((app, index) => {
          console.log(`${index + 1}. ${app.displayName} (${app.version})`);
        });
      }
    }

    console.log('\n✓ Example 3 completed successfully!');
  } catch (error) {
    console.error('Example 3 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 4: Platform-Specific Report (Windows Only)
// ============================================================================

/**
 * Generate inventory for a specific platform (Windows)
 */
async function example4_WindowsDevicesOnly(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 4: Windows Devices Inventory');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const options: DeviceInventoryOptions = {
      platformFilter: 'Windows',
      includeHardware: true,
      includeHealthAttestation: true, // Health attestation is Windows-specific
      includeSoftware: false,
      outputFormats: ['json', 'html'],
      outputDirectory: './reports/windows-devices',
    };

    console.log('Generating Windows device inventory with health attestation...');

    const report = new DeviceInventoryReport(graphClient, config);
    const reportData = await report.executeWithOptions(options);

    console.log('\nWindows Device Summary:');
    console.log(`- Total Windows Devices: ${reportData.metadata.totalDevices}`);

    // Show version distribution
    const windowsSummary = reportData.summary.byOperatingSystem['Windows'];
    if (windowsSummary) {
      console.log('\nWindows Version Distribution:');
      Object.entries(windowsSummary.versions)
        .sort(([, a], [, b]) => b - a)
        .slice(0, 5)
        .forEach(([version, count]) => {
          console.log(`  ${version}: ${count} devices`);
        });
    }

    // Show health attestation sample
    if (reportData.healthAttestation && reportData.healthAttestation.length > 0) {
      console.log(`\nHealth Attestation Data Available: ${reportData.healthAttestation.length} devices`);
      console.log('\nSample Health Attestation (first device):');
      const firstAttestation = reportData.healthAttestation[0];
      console.log(`  Device: ${firstAttestation.deviceName}`);
      console.log(`  Secure Boot: ${firstAttestation.secureBoot || 'N/A'}`);
      console.log(`  BitLocker Status: ${firstAttestation.bitLockerStatus || 'N/A'}`);
      console.log(`  Code Integrity: ${firstAttestation.codeIntegrity || 'N/A'}`);
      console.log(`  TPM Version: ${firstAttestation.tpmVersion || 'N/A'}`);
    }

    console.log('\n✓ Example 4 completed successfully!');
  } catch (error) {
    console.error('Example 4 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 5: Corporate-Owned Devices Only
// ============================================================================

/**
 * Generate inventory for corporate-owned devices only
 */
async function example5_CorporateDevices(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 5: Corporate-Owned Devices Inventory');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const options: DeviceInventoryOptions = {
      ownershipFilter: 'company',
      includeHardware: true,
      includeSoftware: false,
      includeHealthAttestation: false,
      outputFormats: ['json', 'csv'],
      outputDirectory: './reports/corporate-devices',
    };

    console.log('Generating corporate device inventory...');

    const report = new DeviceInventoryReport(graphClient, config);
    const reportData = await report.executeWithOptions(options);

    console.log('\nCorporate Device Summary:');
    console.log(`- Total Corporate Devices: ${reportData.metadata.totalDevices}`);

    console.log('\nBy Operating System:');
    Object.entries(reportData.summary.byOperatingSystem).forEach(([os, summary]) => {
      console.log(`  ${os}: ${summary.totalDevices} devices`);
    });

    console.log('\nBy Enrollment Type:');
    Object.entries(reportData.summary.byEnrollmentType).forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });

    console.log('\n✓ Example 5 completed successfully!');
  } catch (error) {
    console.error('Example 5 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 6: Non-Compliant Devices Report
// ============================================================================

/**
 * Generate inventory for non-compliant devices
 */
async function example6_NonCompliantDevices(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 6: Non-Compliant Devices Report');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const options: DeviceInventoryOptions = {
      complianceFilter: 'noncompliant',
      includeHardware: false,
      includeSoftware: false,
      includeHealthAttestation: false,
      outputFormats: ['json', 'html', 'csv'],
      outputDirectory: './reports/non-compliant',
    };

    console.log('Generating non-compliant device report...');

    const report = new DeviceInventoryReport(graphClient, config);
    const reportData = await report.executeWithOptions(options);

    console.log('\nNon-Compliant Device Summary:');
    console.log(`- Total Non-Compliant Devices: ${reportData.metadata.totalDevices}`);

    console.log('\nBy Operating System:');
    Object.entries(reportData.summary.byOperatingSystem).forEach(([os, summary]) => {
      console.log(`  ${os}: ${summary.totalDevices} devices`);
    });

    console.log('\nNon-Compliant Devices:');
    reportData.devices.slice(0, 10).forEach((device, index) => {
      console.log(`\n${index + 1}. ${device.deviceName}`);
      console.log(`   User: ${device.userDisplayName || 'N/A'}`);
      console.log(`   OS: ${device.operatingSystem} ${device.osVersion}`);
      console.log(`   Last Sync: ${device.lastSyncDateTime ? new Date(device.lastSyncDateTime).toLocaleDateString() : 'Never'}`);
    });

    console.log('\n✓ Example 6 completed successfully!');
  } catch (error) {
    console.error('Example 6 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 7: Complete Comprehensive Report
// ============================================================================

/**
 * Generate the most comprehensive report with all options enabled
 */
async function example7_CompleteComprehensiveReport(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 7: Complete Comprehensive Device Inventory Report');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const options: DeviceInventoryOptions = {
      includeHardware: true,
      includeSoftware: true,
      maxDevicesForSoftware: 25, // Limit software inventory for performance
      includeHealthAttestation: true,
      outputFormats: ['json', 'html', 'csv'],
      outputDirectory: './reports/comprehensive',
    };

    console.log('Generating complete comprehensive device inventory...');
    console.log('Note: This may take several minutes depending on device count.');

    const report = new DeviceInventoryReport(graphClient, config);
    const reportData = await report.executeWithOptions(options);

    // Use the built-in summary printer
    DeviceInventoryReport.printSummary(reportData);

    console.log('\nReport Components Generated:');
    console.log(`✓ Device Inventory: ${reportData.devices.length} devices`);
    console.log(`✓ Hardware Inventory: ${reportData.hardwareInventory?.length || 0} devices`);
    console.log(`✓ Software Inventory: ${reportData.softwareInventory?.length || 0} devices`);
    console.log(`✓ Ownership Details: ${reportData.ownershipDetails?.length || 0} devices`);
    console.log(`✓ Health Attestation: ${reportData.healthAttestation?.length || 0} devices`);

    console.log('\n✓ Example 7 completed successfully!');
  } catch (error) {
    console.error('Example 7 failed:', error);
    throw error;
  }
}

// ============================================================================
// Example 8: Using Individual Methods
// ============================================================================

/**
 * Demonstrate using individual report methods for custom workflows
 */
async function example8_IndividualMethods(): Promise<void> {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 8: Using Individual Report Methods');
  console.log('='.repeat(70) + '\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new DeviceInventoryReport(graphClient, config);

    // Get Windows devices
    console.log('1. Fetching Windows devices...');
    const windowsDevices = await report.getDevicesByPlatform('Windows');
    console.log(`   Found ${windowsDevices.length} Windows devices`);

    // Get corporate devices
    console.log('\n2. Fetching corporate-owned devices...');
    const corporateDevices = await report.getDevicesByOwnershipType('company');
    console.log(`   Found ${corporateDevices.length} corporate devices`);

    // Get non-compliant devices
    console.log('\n3. Fetching non-compliant devices...');
    const nonCompliantDevices = await report.getDevicesByComplianceState('noncompliant');
    console.log(`   Found ${nonCompliantDevices.length} non-compliant devices`);

    // Get specific device details
    if (windowsDevices.length > 0) {
      const firstDeviceId = windowsDevices[0].id;
      console.log(`\n4. Fetching detailed info for device: ${windowsDevices[0].deviceName}`);

      const deviceDetails = await report.getDeviceById(firstDeviceId);
      console.log(`   Device ID: ${deviceDetails.id}`);
      console.log(`   Serial: ${deviceDetails.serialNumber || 'N/A'}`);
      console.log(`   Enrolled: ${deviceDetails.enrolledDateTime ? new Date(deviceDetails.enrolledDateTime).toLocaleDateString() : 'N/A'}`);

      // Get software inventory for this device
      console.log('\n5. Fetching software inventory for this device...');
      const softwareInventory = await report.getDeviceSoftwareInventory(firstDeviceId);
      if (softwareInventory) {
        console.log(`   Found ${softwareInventory.totalApps} applications`);
        if (softwareInventory.applications.length > 0) {
          console.log('\n   Top 5 applications:');
          softwareInventory.applications.slice(0, 5).forEach((app, index) => {
            console.log(`   ${index + 1}. ${app.displayName} (${app.version})`);
          });
        }
      }
    }

    console.log('\n✓ Example 8 completed successfully!');
  } catch (error) {
    console.error('Example 8 failed:', error);
    throw error;
  }
}

// ============================================================================
// Main Function - Run All Examples
// ============================================================================

/**
 * Main function to run all examples
 */
async function main(): Promise<void> {
  console.log('\n');
  console.log('╔════════════════════════════════════════════════════════════════════╗');
  console.log('║   Device Inventory Report - Comprehensive Usage Examples          ║');
  console.log('╚════════════════════════════════════════════════════════════════════╝');

  // Check for required environment variables
  if (
    !process.env.AZURE_TENANT_ID ||
    !process.env.AZURE_CLIENT_ID ||
    !process.env.AZURE_CLIENT_SECRET
  ) {
    console.error('\n❌ Error: Missing required environment variables');
    console.error('\nPlease set the following environment variables:');
    console.error('  - AZURE_TENANT_ID');
    console.error('  - AZURE_CLIENT_ID');
    console.error('  - AZURE_CLIENT_SECRET');
    console.error('\nYou can create a .env file based on .env.example\n');
    process.exit(1);
  }

  const examples = [
    { name: 'Basic Inventory', fn: example1_BasicInventory },
    { name: 'Comprehensive with Hardware', fn: example2_ComprehensiveWithHardware },
    { name: 'Software Inventory', fn: example3_SoftwareInventory },
    { name: 'Windows Devices Only', fn: example4_WindowsDevicesOnly },
    { name: 'Corporate Devices', fn: example5_CorporateDevices },
    { name: 'Non-Compliant Devices', fn: example6_NonCompliantDevices },
    { name: 'Complete Comprehensive', fn: example7_CompleteComprehensiveReport },
    { name: 'Individual Methods', fn: example8_IndividualMethods },
  ];

  // Get example number from command line argument
  const exampleArg = process.argv[2];

  if (exampleArg) {
    const exampleNum = parseInt(exampleArg, 10);
    if (exampleNum >= 1 && exampleNum <= examples.length) {
      const example = examples[exampleNum - 1];
      console.log(`\nRunning Example ${exampleNum}: ${example.name}\n`);
      await example.fn();
    } else {
      console.error(`\n❌ Invalid example number. Please choose 1-${examples.length}\n`);
      process.exit(1);
    }
  } else {
    // Run all examples
    console.log('\nRunning all examples...\n');
    console.log('Note: You can run a specific example by passing its number as an argument.');
    console.log(`Usage: npm run example [1-${examples.length}]\n`);

    for (let i = 0; i < examples.length; i++) {
      const example = examples[i];
      console.log(`\nRunning Example ${i + 1}/${examples.length}: ${example.name}`);
      try {
        await example.fn();
      } catch (error) {
        console.error(`Example ${i + 1} failed:`, error);
        // Continue with other examples
      }

      // Add delay between examples
      if (i < examples.length - 1) {
        console.log('\nWaiting 2 seconds before next example...');
        await new Promise((resolve) => setTimeout(resolve, 2000));
      }
    }
  }

  console.log('\n' + '='.repeat(70));
  console.log('All examples completed!');
  console.log('='.repeat(70) + '\n');
}

// ============================================================================
// Run Examples
// ============================================================================

if (require.main === module) {
  main()
    .then(() => {
      console.log('\n✓ Examples completed successfully!\n');
      process.exit(0);
    })
    .catch((error) => {
      console.error('\n❌ Examples failed:', error);
      process.exit(1);
    });
}

// Export examples for external use
export {
  example1_BasicInventory,
  example2_ComprehensiveWithHardware,
  example3_SoftwareInventory,
  example4_WindowsDevicesOnly,
  example5_CorporateDevices,
  example6_NonCompliantDevices,
  example7_CompleteComprehensiveReport,
  example8_IndividualMethods,
};
