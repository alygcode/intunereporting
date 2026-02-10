/**
 * Policy Assignment Report - Usage Examples
 *
 * This file demonstrates various ways to use the PolicyAssignmentReport class
 * to generate comprehensive policy reports for Microsoft Intune.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { PolicyAssignmentReport, PolicyReportOptions } from '../policy-assignment-report';
import { AppConfig } from '../../types';

// ============================================================================
// Authentication Setup
// ============================================================================

/**
 * Create and authenticate Microsoft Graph client
 */
async function createGraphClient(): Promise<Client> {
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
 * Create a mock config object
 */
function createMockConfig(): AppConfig {
  return {
    authentication: {
      tenantId: process.env.AZURE_TENANT_ID || '',
      clientId: process.env.AZURE_CLIENT_ID || '',
      clientSecret: process.env.AZURE_CLIENT_SECRET || '',
      authMethod: 'clientSecret'
    },
    reports: {
      enabled: ['policy-assignment-report'],
      disabled: [],
      settings: {}
    },
    output: {
      defaultFormat: 'json',
      directory: './reports',
      includeTimestamp: true,
      compression: false
    },
    scheduler: {
      enabled: false,
      timezone: 'UTC',
      schedules: []
    },
    logging: {
      level: 'info',
      file: './logs/intune-reports.log',
      console: true,
      maxSize: '10m',
      maxFiles: 5
    }
  };
}

// ============================================================================
// Example 1: Basic Policy Assignment Report
// ============================================================================

/**
 * Generate a basic policy assignment report
 */
export async function example1_BasicReport() {
  console.log('\n=== Example 1: Basic Policy Assignment Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    console.log('Generating basic policy assignment report...');

    const result = await report.execute({
      includeDeviceStatus: true,
      includeUserStatus: true,
      includeConflicts: true,
      includeSuccessRates: true
    });

    console.log('\nReport Summary:');
    console.log('─'.repeat(60));
    console.log(`Total Policies: ${result.summary.totalPolicies}`);
    console.log(`Policies with Assignments: ${result.summary.policiesWithAssignments}`);
    console.log(`Total Assignments: ${result.summary.totalAssignments}`);
    console.log(`Overall Success Rate: ${result.summary.overallSuccessRate}%`);
    console.log(`Total Conflicts: ${result.summary.totalConflicts}`);
    console.log('─'.repeat(60));

    console.log('\n✓ Basic report generated successfully!');
  } catch (error) {
    console.error('Error generating basic report:', error);
    throw error;
  }
}

// ============================================================================
// Example 2: Export to Multiple Formats
// ============================================================================

/**
 * Generate report and export to JSON, CSV, and HTML
 */
export async function example2_ExportToMultipleFormats() {
  console.log('\n=== Example 2: Export to Multiple Formats ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    const outputDir = './reports/policy-assignments';

    console.log('Generating and exporting policy assignment report...');
    console.log(`Output directory: ${outputDir}`);

    const result = await report.execute({
      includeDeviceStatus: true,
      includeUserStatus: true,
      includeConflicts: true,
      includeSuccessRates: true,
      outputDir
    });

    console.log('\n✓ Reports exported successfully!');
    console.log('Files generated:');
    console.log(`  - JSON report with full data`);
    console.log(`  - CSV report for spreadsheet analysis`);
    console.log(`  - HTML report for easy viewing`);

    return result;
  } catch (error) {
    console.error('Error exporting reports:', error);
    throw error;
  }
}

// ============================================================================
// Example 3: Focused Report (Specific Policy)
// ============================================================================

/**
 * Generate report for a specific policy
 */
export async function example3_SpecificPolicyReport() {
  console.log('\n=== Example 3: Specific Policy Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    // Replace with actual policy ID
    const policyId = 'your-policy-id-here';

    console.log(`Generating report for policy: ${policyId}`);

    const result = await report.execute({
      policyId,
      includeDeviceStatus: true,
      includeUserStatus: true,
      includeConflicts: false, // Skip conflicts for single policy
      includeSuccessRates: true
    });

    if (result.data && result.data.length > 0) {
      const data = result.data[0] as any;
      const policy = data.policies[0];

      console.log('\nPolicy Details:');
      console.log('─'.repeat(60));
      console.log(`Name: ${policy.displayName}`);
      console.log(`Platform: ${policy.platformType}`);
      console.log(`Assignments: ${policy.assignmentCount}`);
      console.log(`Device Statuses: ${policy.deviceStatusCount}`);
      console.log(`User Statuses: ${policy.userStatusCount}`);
      console.log('─'.repeat(60));
    }

    console.log('\n✓ Specific policy report generated successfully!');

    return result;
  } catch (error) {
    console.error('Error generating specific policy report:', error);
    throw error;
  }
}

// ============================================================================
// Example 4: Platform-Filtered Report
// ============================================================================

/**
 * Generate report filtered by platform
 */
export async function example4_PlatformFilteredReport() {
  console.log('\n=== Example 4: Platform-Filtered Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    const platform = 'Windows';

    console.log(`Generating report for ${platform} policies...`);

    const result = await report.execute({
      platformFilter: platform,
      includeDeviceStatus: true,
      includeUserStatus: false,
      includeConflicts: true,
      includeSuccessRates: true
    });

    console.log(`\n${platform} Policies Summary:`);
    console.log('─'.repeat(60));
    console.log(`Total ${platform} Policies: ${result.summary.totalPolicies}`);
    console.log(`Success Rate: ${result.summary.overallSuccessRate}%`);
    console.log('─'.repeat(60));

    console.log('\n✓ Platform-filtered report generated successfully!');

    return result;
  } catch (error) {
    console.error('Error generating platform-filtered report:', error);
    throw error;
  }
}

// ============================================================================
// Example 5: Minimal Report (Summary Only)
// ============================================================================

/**
 * Generate minimal report with summary statistics only
 */
export async function example5_MinimalReport() {
  console.log('\n=== Example 5: Minimal Report (Summary Only) ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    console.log('Generating minimal summary report...');

    const result = await report.execute({
      includeDeviceStatus: false,
      includeUserStatus: false,
      includeConflicts: false,
      includeSuccessRates: false
    });

    console.log('\nQuick Summary:');
    console.log('─'.repeat(60));
    console.log(`Total Policies: ${result.summary.totalPolicies}`);
    console.log(`With Assignments: ${result.summary.policiesWithAssignments}`);
    console.log(`Without Assignments: ${result.summary.policiesWithoutAssignments}`);
    console.log(`Total Assignments: ${result.summary.totalAssignments}`);
    console.log('─'.repeat(60));

    console.log('\nPolicies by Platform:');
    Object.entries(result.summary.policiesByPlatform).forEach(([platform, count]) => {
      console.log(`  ${platform}: ${count}`);
    });

    console.log('\n✓ Minimal report generated successfully!');

    return result;
  } catch (error) {
    console.error('Error generating minimal report:', error);
    throw error;
  }
}

// ============================================================================
// Example 6: Conflict Analysis Report
// ============================================================================

/**
 * Generate report focused on policy conflicts
 */
export async function example6_ConflictAnalysisReport() {
  console.log('\n=== Example 6: Conflict Analysis Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    console.log('Analyzing policy conflicts...');

    const result = await report.execute({
      includeDeviceStatus: false,
      includeUserStatus: false,
      includeConflicts: true,
      includeSuccessRates: false
    });

    if (result.data && result.data.length > 0) {
      const data = result.data[0] as any;
      const conflicts = data.conflicts;

      console.log(`\nFound ${conflicts.length} policy conflicts`);

      if (conflicts.length > 0) {
        console.log('\nConflict Details:');
        console.log('─'.repeat(60));

        conflicts.forEach((conflict: any, index: number) => {
          console.log(`\nConflict ${index + 1}:`);
          console.log(`  Affected Devices: ${conflict.affectedDevicesCount}`);
          console.log(`  Conflicting Policies:`);
          conflict.conflictingPolicyNames.forEach((name: string) => {
            console.log(`    - ${name}`);
          });
          if (conflict.description) {
            console.log(`  Description: ${conflict.description}`);
          }
        });

        console.log('─'.repeat(60));
      } else {
        console.log('\n✓ No policy conflicts detected!');
      }
    }

    console.log('\n✓ Conflict analysis completed successfully!');

    return result;
  } catch (error) {
    console.error('Error analyzing conflicts:', error);
    throw error;
  }
}

// ============================================================================
// Example 7: Success Rate Analysis
// ============================================================================

/**
 * Generate report focused on deployment success rates
 */
export async function example7_SuccessRateAnalysis() {
  console.log('\n=== Example 7: Success Rate Analysis ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    console.log('Analyzing policy deployment success rates...');

    const result = await report.execute({
      includeDeviceStatus: true,
      includeUserStatus: true,
      includeConflicts: false,
      includeSuccessRates: true
    });

    if (result.data && result.data.length > 0) {
      const data = result.data[0] as any;
      const successRates = data.successRates;

      console.log('\nDeployment Success Rates:');
      console.log('─'.repeat(80));
      console.log(
        'Policy Name'.padEnd(40) +
        'Success'.padEnd(12) +
        'Failure'.padEnd(12) +
        'Rate'.padEnd(10)
      );
      console.log('─'.repeat(80));

      successRates.forEach((rate: any) => {
        const name = rate.policyName.length > 38
          ? rate.policyName.substring(0, 35) + '...'
          : rate.policyName;

        console.log(
          name.padEnd(40) +
          `${rate.successCount}/${rate.totalTargets}`.padEnd(12) +
          rate.failureCount.toString().padEnd(12) +
          `${rate.successRate}%`.padEnd(10)
        );
      });

      console.log('─'.repeat(80));

      // Identify policies with low success rates
      const lowSuccessRates = successRates.filter((r: any) => r.successRate < 80 && r.totalTargets > 0);

      if (lowSuccessRates.length > 0) {
        console.log('\n⚠ Policies with Success Rate < 80%:');
        lowSuccessRates.forEach((rate: any) => {
          console.log(`  - ${rate.policyName}: ${rate.successRate}%`);
        });
      }
    }

    console.log('\n✓ Success rate analysis completed!');

    return result;
  } catch (error) {
    console.error('Error analyzing success rates:', error);
    throw error;
  }
}

// ============================================================================
// Example 8: Complete Comprehensive Report
// ============================================================================

/**
 * Generate complete comprehensive report with all features
 */
export async function example8_ComprehensiveReport() {
  console.log('\n=== Example 8: Complete Comprehensive Report ===\n');

  try {
    const graphClient = await createGraphClient();
    const config = createMockConfig();

    const report = new PolicyAssignmentReport(graphClient, config);

    const outputDir = './reports/comprehensive';

    console.log('Generating comprehensive policy assignment report...');
    console.log('This may take several minutes for large tenants...\n');

    const options: PolicyReportOptions = {
      includeDeviceStatus: true,
      includeUserStatus: true,
      includeConflicts: true,
      includeSuccessRates: true,
      outputDir,
      maxRetries: 5
    };

    const startTime = Date.now();
    const result = await report.execute(options);
    const duration = ((Date.now() - startTime) / 1000).toFixed(2);

    console.log('\n' + '═'.repeat(80));
    console.log('COMPREHENSIVE POLICY ASSIGNMENT REPORT');
    console.log('═'.repeat(80));

    console.log('\nExecution Summary:');
    console.log(`  Duration: ${duration} seconds`);
    console.log(`  Report Generated: ${new Date(result.summary.generatedAt).toLocaleString()}`);

    console.log('\nPolicy Overview:');
    console.log(`  Total Policies: ${result.summary.totalPolicies}`);
    console.log(`  With Assignments: ${result.summary.policiesWithAssignments}`);
    console.log(`  Without Assignments: ${result.summary.policiesWithoutAssignments}`);

    console.log('\nDeployment Statistics:');
    console.log(`  Total Assignments: ${result.summary.totalAssignments}`);
    console.log(`  Device Deployments: ${result.summary.totalDeviceDeployments}`);
    console.log(`  User Deployments: ${result.summary.totalUserDeployments}`);

    console.log('\nSuccess Metrics:');
    console.log(`  Overall Success Rate: ${result.summary.overallSuccessRate}%`);
    console.log(`  Overall Failure Rate: ${result.summary.overallFailureRate}%`);

    console.log('\nConflicts:');
    console.log(`  Total Conflicts Detected: ${result.summary.totalConflicts}`);

    console.log('\nPolicies by Platform:');
    Object.entries(result.summary.policiesByPlatform)
      .sort(([, a], [, b]) => (b as number) - (a as number))
      .forEach(([platform, count]) => {
        console.log(`  ${platform}: ${count}`);
      });

    console.log('\n' + '═'.repeat(80));
    console.log('✓ Comprehensive report generated successfully!');
    console.log(`\nReports saved to: ${outputDir}`);
    console.log('  - JSON: Full data export');
    console.log('  - CSV: Spreadsheet-friendly format');
    console.log('  - HTML: Interactive web report');
    console.log('═'.repeat(80) + '\n');

    return result;
  } catch (error) {
    console.error('Error generating comprehensive report:', error);
    throw error;
  }
}

// ============================================================================
// Main Function - Run All Examples
// ============================================================================

/**
 * Run all examples
 */
async function main() {
  console.log('\n' + '═'.repeat(80));
  console.log('POLICY ASSIGNMENT REPORT - EXAMPLES');
  console.log('═'.repeat(80));

  try {
    // Uncomment the examples you want to run:

    // await example1_BasicReport();
    // await example2_ExportToMultipleFormats();
    // await example3_SpecificPolicyReport();
    // await example4_PlatformFilteredReport();
    // await example5_MinimalReport();
    // await example6_ConflictAnalysisReport();
    // await example7_SuccessRateAnalysis();
    await example8_ComprehensiveReport();

    console.log('\n✓ All examples completed successfully!\n');
    process.exit(0);
  } catch (error) {
    console.error('\n✗ Error running examples:', error);
    process.exit(1);
  }
}

// Run if executed directly
if (require.main === module) {
  main();
}

// Export all examples
export {
  createGraphClient,
  createMockConfig
};
