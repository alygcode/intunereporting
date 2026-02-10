/**
 * Example usage of the Security Baseline Report
 *
 * This file demonstrates how to use the SecurityBaselineReport class
 * to generate comprehensive security baseline compliance reports.
 */

import { ClientSecretCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { SecurityBaselineReport } from '../security-baseline-report';
import { AppConfig } from '../../types';
import * as dotenv from 'dotenv';
import * as path from 'path';

// Load environment variables
dotenv.config();

/**
 * Initialize Microsoft Graph Client
 */
async function initializeGraphClient(): Promise<Client> {
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
 * Get default app configuration
 */
function getDefaultConfig(): AppConfig {
  return {
    authentication: {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET || '',
      authMethod: 'clientSecret',
    },
    reports: {
      enabled: ['security-baseline-report'],
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
      file: 'logs/intune-reports.log',
      console: true,
      maxSize: '10m',
      maxFiles: 5,
    },
  };
}

/**
 * Example 1: Generate a comprehensive security baseline report
 */
async function example1_ComprehensiveReport() {
  console.log('\n=== Example 1: Comprehensive Security Baseline Report ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    // Execute the full report
    const reportData = await report.execute();

    console.log('Report generated successfully!');
    console.log(`Total records: ${reportData.metadata.recordCount}`);
    console.log('\nSummary:');
    console.log(JSON.stringify(reportData.summary, null, 2));

    // Optionally save to file
    const outputPath = path.join('./reports', 'security-baseline-full-report.json');
    // await fs.writeFile(outputPath, JSON.stringify(reportData, null, 2));
    // console.log(`\nReport saved to: ${outputPath}`);
  } catch (error) {
    console.error('Error generating report:', error);
  }
}

/**
 * Example 2: Get all security baselines
 */
async function example2_GetAllBaselines() {
  console.log('\n=== Example 2: Get All Security Baselines ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    const baselines = await report.getAllSecurityBaselines();

    console.log(`Found ${baselines.length} security baseline templates:\n`);

    baselines.forEach((baseline, idx) => {
      console.log(`${idx + 1}. ${baseline.displayName}`);
      console.log(`   ID: ${baseline.id}`);
      console.log(`   Version: ${baseline.version || 'N/A'}`);
      console.log(`   Platform: ${baseline.platformType || 'N/A'}`);
      console.log(`   Description: ${baseline.description || 'N/A'}\n`);
    });
  } catch (error) {
    console.error('Error fetching baselines:', error);
  }
}

/**
 * Example 3: Get baseline compliance per device
 */
async function example3_GetDeviceCompliance() {
  console.log('\n=== Example 3: Get Baseline Compliance Per Device ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    // First, get the intents to find an intent ID
    const intents = await report.getSecurityBaselineIntents();

    if (intents.length === 0) {
      console.log('No security baseline intents found.');
      return;
    }

    // Use the first intent as an example
    const firstIntent = intents[0];
    console.log(`Analyzing baseline: ${firstIntent.displayName}`);
    console.log(`Intent ID: ${firstIntent.id}\n`);

    // Get device compliance for this baseline
    const deviceStates = await report.getBaselineCompliancePerDevice(firstIntent.id);

    console.log(`Total devices: ${deviceStates.length}\n`);

    // Group by compliance state
    const byState: { [key: string]: number } = {};
    deviceStates.forEach((device) => {
      const state = device.state || 'unknown';
      byState[state] = (byState[state] || 0) + 1;
    });

    console.log('Compliance breakdown:');
    Object.entries(byState).forEach(([state, count]) => {
      console.log(`  ${state}: ${count}`);
    });

    // Show some example devices
    console.log('\nExample non-compliant devices (first 5):');
    const nonCompliant = deviceStates
      .filter((d) => d.state === 'nonCompliant')
      .slice(0, 5);

    nonCompliant.forEach((device, idx) => {
      console.log(`\n${idx + 1}. ${device.deviceDisplayName || 'Unknown Device'}`);
      console.log(`   User: ${device.userPrincipalName || 'Unknown'}`);
      console.log(`   Model: ${device.deviceModel || 'Unknown'}`);
      console.log(`   Last Reported: ${device.lastReportedDateTime || 'Never'}`);
    });
  } catch (error) {
    console.error('Error getting device compliance:', error);
  }
}

/**
 * Example 4: Get non-compliant settings
 */
async function example4_GetNonCompliantSettings() {
  console.log('\n=== Example 4: Get Non-Compliant Settings ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    const nonCompliantSettings = await report.getNonCompliantSettings();

    console.log(`Found ${nonCompliantSettings.length} non-compliant settings\n`);

    // Group by baseline
    const byBaseline: { [key: string]: number } = {};
    nonCompliantSettings.forEach((setting) => {
      byBaseline[setting.baselineName] = (byBaseline[setting.baselineName] || 0) + 1;
    });

    console.log('Non-compliant settings by baseline:');
    Object.entries(byBaseline)
      .sort((a, b) => b[1] - a[1])
      .forEach(([baseline, count]) => {
        console.log(`  ${baseline}: ${count} issues`);
      });

    // Show top 10 non-compliant settings
    console.log('\nTop 10 non-compliant settings:');
    nonCompliantSettings.slice(0, 10).forEach((setting, idx) => {
      console.log(`\n${idx + 1}. ${setting.settingName}`);
      console.log(`   Device: ${setting.deviceName}`);
      console.log(`   User: ${setting.userPrincipalName}`);
      console.log(`   Baseline: ${setting.baselineName}`);
      console.log(`   Expected: ${setting.expectedValue}`);
      console.log(`   Current: ${setting.currentValue}`);
      if (setting.errorCode) {
        console.log(`   Error Code: ${setting.errorCode}`);
      }
    });
  } catch (error) {
    console.error('Error getting non-compliant settings:', error);
  }
}

/**
 * Example 5: Get security score trends
 */
async function example5_GetSecurityTrends() {
  console.log('\n=== Example 5: Get Security Score Trends ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    const trends = await report.getSecurityScoreTrends(30);

    console.log(`Generated ${trends.length} trend data points\n`);

    trends.forEach((trend) => {
      console.log(`\nBaseline: ${trend.baselineName}`);
      console.log(`Date: ${trend.date}`);
      console.log(`Compliance: ${trend.compliancePercentage.toFixed(2)}%`);
      console.log(
        `Devices: ${trend.compliantDevices}/${trend.totalDevices} compliant`
      );
    });

    // Calculate average compliance
    if (trends.length > 0) {
      const avgCompliance =
        trends.reduce((sum, t) => sum + t.compliancePercentage, 0) / trends.length;
      console.log(`\nAverage Compliance: ${avgCompliance.toFixed(2)}%`);
    }
  } catch (error) {
    console.error('Error getting security trends:', error);
  }
}

/**
 * Example 6: Generate and export comprehensive report
 */
async function example6_GenerateAndExport() {
  console.log('\n=== Example 6: Generate and Export Comprehensive Report ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    // Generate comprehensive report with all options
    const comprehensiveReport = await report.generateComprehensiveReport({
      includeDeviceDetails: true,
      includeSettingDetails: true,
      includeTrends: true,
      trendDays: 30,
    });

    // Print summary
    console.log('Comprehensive Report Generated!\n');
    SecurityBaselineReport.printSummary(comprehensiveReport);

    // Export to multiple formats
    const exportedFiles = await report.exportReport(comprehensiveReport, {
      outputDir: './reports',
      formats: ['json', 'csv', 'html'],
    });

    console.log('\nExported Files:');
    Object.entries(exportedFiles).forEach(([format, filePath]) => {
      console.log(`  ${format.toUpperCase()}: ${filePath}`);
    });
  } catch (error) {
    console.error('Error generating and exporting report:', error);
  }
}

/**
 * Example 7: Filter by specific baseline template
 */
async function example7_FilterByTemplate() {
  console.log('\n=== Example 7: Filter by Specific Baseline Template ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    // Get all templates first
    const templates = await report.getAllSecurityBaselines();

    if (templates.length === 0) {
      console.log('No baseline templates found.');
      return;
    }

    // Use the first template as an example
    const templateId = templates[0].id;
    console.log(`Filtering by template: ${templates[0].displayName}`);
    console.log(`Template ID: ${templateId}\n`);

    // Generate report for this template only
    const filteredReport = await report.generateComprehensiveReport({
      templateId: templateId,
      includeDeviceDetails: true,
      includeSettingDetails: true,
    });

    console.log('Filtered Report Generated!\n');
    console.log(`Intents for this template: ${filteredReport.intents.length}`);
    console.log(`Devices affected: ${filteredReport.metadata.totalDevices}`);
    console.log(
      `Compliant devices: ${filteredReport.summary.compliantCount || 0}`
    );
    console.log(
      `Non-compliant devices: ${filteredReport.summary.nonCompliantCount || 0}`
    );
  } catch (error) {
    console.error('Error generating filtered report:', error);
  }
}

/**
 * Example 8: Export non-compliant settings to CSV
 */
async function example8_ExportNonCompliantToCSV() {
  console.log('\n=== Example 8: Export Non-Compliant Settings to CSV ===\n');

  try {
    const graphClient = await initializeGraphClient();
    const config = getDefaultConfig();
    const report = new SecurityBaselineReport(graphClient, config);

    const nonCompliantSettings = await report.getNonCompliantSettings();

    if (nonCompliantSettings.length === 0) {
      console.log('No non-compliant settings found. Great job!');
      return;
    }

    const outputPath = path.join('./reports', 'non-compliant-settings.csv');
    await report.exportNonCompliantSettingsToCSV(nonCompliantSettings, outputPath);

    console.log(`Exported ${nonCompliantSettings.length} non-compliant settings`);
    console.log(`CSV file saved to: ${outputPath}`);
  } catch (error) {
    console.error('Error exporting non-compliant settings:', error);
  }
}

/**
 * Main function - runs all examples
 */
async function main() {
  console.log('=================================================');
  console.log('  Security Baseline Report - Usage Examples');
  console.log('=================================================');

  const examples = [
    { name: 'Comprehensive Report', fn: example1_ComprehensiveReport },
    { name: 'Get All Baselines', fn: example2_GetAllBaselines },
    { name: 'Get Device Compliance', fn: example3_GetDeviceCompliance },
    { name: 'Get Non-Compliant Settings', fn: example4_GetNonCompliantSettings },
    { name: 'Get Security Trends', fn: example5_GetSecurityTrends },
    { name: 'Generate and Export', fn: example6_GenerateAndExport },
    { name: 'Filter by Template', fn: example7_FilterByTemplate },
    { name: 'Export Non-Compliant CSV', fn: example8_ExportNonCompliantToCSV },
  ];

  // Get example number from command line, or run all
  const exampleNum = process.argv[2] ? parseInt(process.argv[2]) : null;

  if (exampleNum && exampleNum >= 1 && exampleNum <= examples.length) {
    // Run specific example
    const example = examples[exampleNum - 1];
    console.log(`\nRunning Example ${exampleNum}: ${example.name}\n`);
    await example.fn();
  } else {
    // Run all examples (or show menu)
    console.log('\nAvailable Examples:');
    examples.forEach((example, idx) => {
      console.log(`  ${idx + 1}. ${example.name}`);
    });
    console.log('\nUsage:');
    console.log('  ts-node src/reports/examples/security-baseline-example.ts [number]');
    console.log('\nExample:');
    console.log('  ts-node src/reports/examples/security-baseline-example.ts 1');
    console.log('  (runs the Comprehensive Report example)');
    console.log('\nTo run all examples, omit the number (not recommended for production).');
  }
}

// Run main function if executed directly
if (require.main === module) {
  main().catch((error) => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}

export {
  example1_ComprehensiveReport,
  example2_GetAllBaselines,
  example3_GetDeviceCompliance,
  example4_GetNonCompliantSettings,
  example5_GetSecurityTrends,
  example6_GenerateAndExport,
  example7_FilterByTemplate,
  example8_ExportNonCompliantToCSV,
};
