/**
 * Example Usage of Device Compliance Reporter
 *
 * This file demonstrates various ways to use the Device Compliance Reporter
 * to generate compliance reports from Microsoft Intune.
 */

import {
  DeviceComplianceReporter,
  createReporterFromEnv,
  generateComplianceReport,
  AuthConfig,
  ReportOptions
} from '../device-compliance';
import * as dotenv from 'dotenv';

// Load environment variables
dotenv.config();

/**
 * Example 1: Basic usage with environment variables
 */
async function example1BasicUsage(): Promise<void> {
  console.log('\n=== Example 1: Basic Usage ===\n');

  try {
    // Create reporter from environment variables
    const reporter = createReporterFromEnv();

    // Generate report
    const result = await reporter.generateReport({
      outputDir: './reports',
      includeDetails: true
    });

    // Print summary
    DeviceComplianceReporter.printSummary(result.summary);

    console.log('Report files generated:');
    console.log('- JSON:', result.files.json);
    console.log('- CSV:', result.files.csv);
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 2: Using explicit credentials
 */
async function example2ExplicitCredentials(): Promise<void> {
  console.log('\n=== Example 2: Explicit Credentials ===\n');

  try {
    const config: AuthConfig = {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET || ''
    };

    const reporter = new DeviceComplianceReporter(config);

    const result = await reporter.generateReport({
      outputDir: './reports',
      includeDetails: true,
      maxRetries: 5,
      retryDelay: 2000
    });

    console.log(`Total devices: ${result.summary.totalDevices}`);
    console.log(`Compliance rate: ${result.summary.compliancePercentage}%`);
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 3: Filtering by operating system
 */
async function example3FilterByOS(): Promise<void> {
  console.log('\n=== Example 3: Filter by Operating System ===\n');

  try {
    const reporter = createReporterFromEnv();

    // Generate report for Windows devices only
    const windowsResult = await reporter.generateReport({
      outputDir: './reports/windows',
      osFilter: 'Windows',
      includeDetails: true
    });

    console.log('Windows Devices:');
    console.log(`- Total: ${windowsResult.summary.totalDevices}`);
    console.log(`- Compliant: ${windowsResult.summary.compliantDevices}`);
    console.log(`- Compliance Rate: ${windowsResult.summary.compliancePercentage}%`);
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 4: Filtering by compliance state
 */
async function example4FilterByCompliance(): Promise<void> {
  console.log('\n=== Example 4: Filter by Compliance State ===\n');

  try {
    const reporter = createReporterFromEnv();

    // Generate report for non-compliant devices only
    const nonCompliantResult = await reporter.generateReport({
      outputDir: './reports/non-compliant',
      complianceFilter: 'noncompliant',
      includeDetails: true
    });

    console.log('Non-Compliant Devices:');
    console.log(`- Total: ${nonCompliantResult.summary.totalDevices}`);

    if (nonCompliantResult.devices.length > 0) {
      console.log('\nFirst 5 non-compliant devices:');
      nonCompliantResult.devices.slice(0, 5).forEach((device, index) => {
        console.log(`${index + 1}. ${device.deviceName} (${device.userPrincipalName})`);
      });
    }
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 5: Using the quick helper function
 */
async function example5QuickHelper(): Promise<void> {
  console.log('\n=== Example 5: Quick Helper Function ===\n');

  try {
    const config: AuthConfig = {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET || ''
    };

    const result = await generateComplianceReport(config, './reports');

    DeviceComplianceReporter.printSummary(result.summary);
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 6: Custom report processing
 */
async function example6CustomProcessing(): Promise<void> {
  console.log('\n=== Example 6: Custom Report Processing ===\n');

  try {
    const reporter = createReporterFromEnv();

    const result = await reporter.generateReport({
      includeDetails: true
    });

    // Custom analysis: Find devices not checked in for 7+ days
    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

    const staleDevices = result.devices.filter(device => {
      const lastReported = new Date(device.lastReportedDateTime);
      return lastReported < sevenDaysAgo;
    });

    console.log(`Devices not checked in for 7+ days: ${staleDevices.length}`);

    if (staleDevices.length > 0) {
      console.log('\nStale devices:');
      staleDevices.slice(0, 5).forEach(device => {
        console.log(`- ${device.deviceName}: ${device.lastReportedDateTime}`);
      });
    }

    // Custom analysis: Compliance by manufacturer
    const byManufacturer: Record<string, { total: number; compliant: number }> = {};

    result.devices.forEach(device => {
      const manufacturer = device.manufacturer || 'Unknown';
      if (!byManufacturer[manufacturer]) {
        byManufacturer[manufacturer] = { total: 0, compliant: 0 };
      }
      byManufacturer[manufacturer].total++;
      if (device.complianceState.toLowerCase() === 'compliant') {
        byManufacturer[manufacturer].compliant++;
      }
    });

    console.log('\nCompliance by Manufacturer:');
    Object.entries(byManufacturer)
      .sort((a, b) => b[1].total - a[1].total)
      .slice(0, 5)
      .forEach(([manufacturer, stats]) => {
        const percentage = ((stats.compliant / stats.total) * 100).toFixed(1);
        console.log(`- ${manufacturer}: ${stats.compliant}/${stats.total} (${percentage}%)`);
      });
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 7: Multiple reports with different filters
 */
async function example7MultipleReports(): Promise<void> {
  console.log('\n=== Example 7: Multiple Reports ===\n');

  try {
    const reporter = createReporterFromEnv();

    const options: ReportOptions[] = [
      { outputDir: './reports/all', includeDetails: true },
      { outputDir: './reports/windows', osFilter: 'Windows', includeDetails: true },
      { outputDir: './reports/ios', osFilter: 'iOS', includeDetails: true },
      { outputDir: './reports/android', osFilter: 'Android', includeDetails: true }
    ];

    console.log('Generating multiple reports...\n');

    for (const option of options) {
      const result = await reporter.generateReport(option);
      const filterName = option.osFilter || 'All Devices';
      console.log(`${filterName}: ${result.summary.totalDevices} devices, ${result.summary.compliancePercentage}% compliant`);
    }

    console.log('\nAll reports generated successfully!');
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Example 8: Error handling demonstration
 */
async function example8ErrorHandling(): Promise<void> {
  console.log('\n=== Example 8: Error Handling ===\n');

  try {
    // This will fail due to invalid credentials
    const config: AuthConfig = {
      tenantId: 'invalid',
      clientId: 'invalid',
      clientSecret: 'invalid'
    };

    const reporter = new DeviceComplianceReporter(config);
    await reporter.generateReport();
  } catch (error) {
    if (error instanceof Error) {
      console.log('Caught expected error:');
      console.log(`- Type: ${error.name}`);
      console.log(`- Message: ${error.message}`);
    }
  }

  // Demonstrate retry logic with valid credentials
  try {
    const reporter = createReporterFromEnv();

    const result = await reporter.generateReport({
      maxRetries: 5,
      retryDelay: 1000
    });

    console.log('\nSuccessfully generated report with retry logic enabled');
    console.log(`Total devices: ${result.summary.totalDevices}`);
  } catch (error) {
    console.error('Error:', error);
  }
}

/**
 * Main function to run all examples
 */
async function main(): Promise<void> {
  const examples = [
    { name: 'Basic Usage', fn: example1BasicUsage },
    { name: 'Explicit Credentials', fn: example2ExplicitCredentials },
    { name: 'Filter by OS', fn: example3FilterByOS },
    { name: 'Filter by Compliance', fn: example4FilterByCompliance },
    { name: 'Quick Helper', fn: example5QuickHelper },
    { name: 'Custom Processing', fn: example6CustomProcessing },
    { name: 'Multiple Reports', fn: example7MultipleReports },
    { name: 'Error Handling', fn: example8ErrorHandling }
  ];

  console.log('Device Compliance Reporter - Examples');
  console.log('=====================================\n');

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

  // Run all examples
  console.log('Available examples:');
  examples.forEach((example, index) => {
    console.log(`${index + 1}. ${example.name}`);
  });
  console.log('\nUsage: ts-node device-compliance-example.ts [example-number]');
  console.log('Example: ts-node device-compliance-example.ts 1\n');

  // Uncomment to run a specific example by default
  // await example1BasicUsage();
}

// Run if executed directly
if (require.main === module) {
  main().catch(console.error);
}

export {
  example1BasicUsage,
  example2ExplicitCredentials,
  example3FilterByOS,
  example4FilterByCompliance,
  example5QuickHelper,
  example6CustomProcessing,
  example7MultipleReports,
  example8ErrorHandling
};
