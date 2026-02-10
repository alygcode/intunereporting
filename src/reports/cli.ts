#!/usr/bin/env node

/**
 * Command Line Interface for Device Compliance Reporter
 *
 * This CLI tool provides an easy way to generate device compliance reports
 * from the command line with various options and filters.
 *
 * Usage:
 *   ts-node src/reports/cli.ts [options]
 *   npm run report -- [options]
 */

import { program } from 'commander';
import * as dotenv from 'dotenv';
import * as path from 'path';
import {
  DeviceComplianceReporter,
  createReporterFromEnv,
  AuthConfig,
  ReportOptions
} from './device-compliance';

// Load environment variables
dotenv.config();

/**
 * Main CLI function
 */
async function main(): Promise<void> {
  program
    .name('intune-compliance-report')
    .description('Generate device compliance reports from Microsoft Intune')
    .version('1.0.0');

  program
    .command('generate')
    .description('Generate a device compliance report')
    .option('-o, --output <directory>', 'Output directory for reports', './reports')
    .option('-t, --tenant <id>', 'Azure AD tenant ID (overrides env)')
    .option('-c, --client <id>', 'Azure AD client ID (overrides env)')
    .option('-s, --secret <secret>', 'Azure AD client secret (overrides env)')
    .option('--os-filter <os>', 'Filter by operating system (e.g., Windows, iOS, Android)')
    .option('--compliance-filter <state>', 'Filter by compliance state (e.g., compliant, noncompliant)')
    .option('--no-details', 'Exclude device details from output')
    .option('--max-retries <number>', 'Maximum retry attempts', '3')
    .option('--retry-delay <ms>', 'Retry delay in milliseconds', '1000')
    .option('--json-only', 'Generate only JSON output')
    .option('--csv-only', 'Generate only CSV output')
    .option('--summary-only', 'Print summary only (no file output)')
    .option('-v, --verbose', 'Enable verbose logging')
    .action(async (options) => {
      try {
        if (options.verbose) {
          console.log('Options:', options);
        }

        // Create auth config
        let authConfig: AuthConfig;

        if (options.tenant && options.client && options.secret) {
          // Use command line provided credentials
          authConfig = {
            tenantId: options.tenant,
            clientId: options.client,
            clientSecret: options.secret
          };
          if (options.verbose) {
            console.log('Using credentials from command line arguments');
          }
        } else {
          // Use environment variables
          authConfig = {
            tenantId: process.env.AZURE_TENANT_ID || '',
            clientId: process.env.AZURE_CLIENT_ID || '',
            clientSecret: process.env.AZURE_CLIENT_SECRET || ''
          };
          if (options.verbose) {
            console.log('Using credentials from environment variables');
          }
        }

        // Validate credentials
        if (!authConfig.tenantId || !authConfig.clientId || !authConfig.clientSecret) {
          console.error('Error: Missing authentication credentials');
          console.error('Provide credentials via:');
          console.error('  1. Environment variables: AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET');
          console.error('  2. Command line: --tenant <id> --client <id> --secret <secret>');
          process.exit(1);
        }

        // Create reporter
        const reporter = new DeviceComplianceReporter(authConfig);

        // Build report options
        const reportOptions: ReportOptions = {
          outputDir: options.summaryOnly ? undefined : options.output,
          includeDetails: options.details !== false,
          osFilter: options.osFilter,
          complianceFilter: options.complianceFilter,
          maxRetries: parseInt(options.maxRetries, 10),
          retryDelay: parseInt(options.retryDelay, 10)
        };

        if (options.verbose) {
          console.log('Report options:', reportOptions);
        }

        // Generate report
        console.log('Generating compliance report...\n');
        const startTime = Date.now();

        const result = await reporter.generateReport(reportOptions);

        const duration = ((Date.now() - startTime) / 1000).toFixed(2);
        console.log(`\nReport generated in ${duration} seconds\n`);

        // Print summary
        DeviceComplianceReporter.printSummary(result.summary);

        // Display file paths
        if (!options.summaryOnly) {
          console.log('Report files:');
          if (result.files.json && !options.csvOnly) {
            console.log(`  JSON: ${result.files.json}`);
          }
          if (result.files.csv && !options.jsonOnly) {
            console.log(`  CSV: ${result.files.csv}`);
          }
        }

        process.exit(0);
      } catch (error) {
        console.error('\nError generating report:');
        console.error(error instanceof Error ? error.message : String(error));

        if (options.verbose && error instanceof Error && error.stack) {
          console.error('\nStack trace:');
          console.error(error.stack);
        }

        process.exit(1);
      }
    });

  program
    .command('test-connection')
    .description('Test connection to Microsoft Graph API')
    .option('-t, --tenant <id>', 'Azure AD tenant ID (overrides env)')
    .option('-c, --client <id>', 'Azure AD client ID (overrides env)')
    .option('-s, --secret <secret>', 'Azure AD client secret (overrides env)')
    .option('-v, --verbose', 'Enable verbose logging')
    .action(async (options) => {
      try {
        console.log('Testing connection to Microsoft Graph API...\n');

        // Create auth config
        let authConfig: AuthConfig;

        if (options.tenant && options.client && options.secret) {
          authConfig = {
            tenantId: options.tenant,
            clientId: options.client,
            clientSecret: options.secret
          };
        } else {
          authConfig = {
            tenantId: process.env.AZURE_TENANT_ID || '',
            clientId: process.env.AZURE_CLIENT_ID || '',
            clientSecret: process.env.AZURE_CLIENT_SECRET || ''
          };
        }

        // Validate credentials
        if (!authConfig.tenantId || !authConfig.clientId || !authConfig.clientSecret) {
          console.error('Error: Missing authentication credentials');
          process.exit(1);
        }

        if (options.verbose) {
          console.log('Tenant ID:', authConfig.tenantId);
          console.log('Client ID:', authConfig.clientId);
          console.log('Client Secret:', authConfig.clientSecret.substring(0, 4) + '****');
          console.log();
        }

        // Create reporter and test
        const reporter = new DeviceComplianceReporter(authConfig);

        // This will attempt to authenticate
        await (reporter as any).ensureAuthenticated();

        console.log('✓ Successfully connected to Microsoft Graph API');
        console.log('✓ Authentication successful');
        console.log('✓ Credentials are valid\n');

        console.log('You can now generate reports using the "generate" command.');

        process.exit(0);
      } catch (error) {
        console.error('\n✗ Connection test failed');
        console.error(error instanceof Error ? error.message : String(error));

        if (options.verbose && error instanceof Error && error.stack) {
          console.error('\nStack trace:');
          console.error(error.stack);
        }

        console.error('\nTroubleshooting:');
        console.error('  1. Verify your Azure AD credentials');
        console.error('  2. Check that the app registration exists');
        console.error('  3. Ensure required API permissions are granted');
        console.error('  4. Verify admin consent has been given');

        process.exit(1);
      }
    });

  program
    .command('list-filters')
    .description('List available filter options')
    .action(() => {
      console.log('Available Operating System Filters:');
      console.log('  --os-filter Windows    - Windows devices');
      console.log('  --os-filter iOS        - iOS/iPadOS devices');
      console.log('  --os-filter Android    - Android devices');
      console.log('  --os-filter macOS      - macOS devices');
      console.log('  --os-filter Linux      - Linux devices');
      console.log();
      console.log('Available Compliance State Filters:');
      console.log('  --compliance-filter compliant      - Compliant devices only');
      console.log('  --compliance-filter noncompliant   - Non-compliant devices only');
      console.log('  --compliance-filter ingraceperiod  - Devices in grace period');
      console.log('  --compliance-filter unknown        - Unknown compliance state');
      console.log('  --compliance-filter error          - Devices with errors');
      console.log();
      console.log('Example Usage:');
      console.log('  # Generate report for non-compliant Windows devices');
      console.log('  npm run report -- generate --os-filter Windows --compliance-filter noncompliant');
      console.log();
      console.log('  # Generate iOS compliance report with verbose output');
      console.log('  npm run report -- generate --os-filter iOS -v');
      console.log();
      console.log('  # Test connection');
      console.log('  npm run report -- test-connection');
    });

  program
    .command('examples')
    .description('Show usage examples')
    .action(() => {
      console.log('Device Compliance Reporter - Usage Examples\n');
      console.log('=' .repeat(60));
      console.log();

      console.log('1. Generate basic report:');
      console.log('   npm run report -- generate');
      console.log();

      console.log('2. Generate report with custom output directory:');
      console.log('   npm run report -- generate --output ./my-reports');
      console.log();

      console.log('3. Filter by operating system:');
      console.log('   npm run report -- generate --os-filter Windows');
      console.log('   npm run report -- generate --os-filter iOS');
      console.log();

      console.log('4. Filter by compliance state:');
      console.log('   npm run report -- generate --compliance-filter noncompliant');
      console.log();

      console.log('5. Combine filters:');
      console.log('   npm run report -- generate --os-filter Android --compliance-filter noncompliant');
      console.log();

      console.log('6. Summary only (no files):');
      console.log('   npm run report -- generate --summary-only');
      console.log();

      console.log('7. With retry configuration:');
      console.log('   npm run report -- generate --max-retries 5 --retry-delay 2000');
      console.log();

      console.log('8. Test connection:');
      console.log('   npm run report -- test-connection');
      console.log();

      console.log('9. Verbose output:');
      console.log('   npm run report -- generate -v');
      console.log();

      console.log('10. Using explicit credentials:');
      console.log('    npm run report -- generate \\');
      console.log('      --tenant "your-tenant-id" \\');
      console.log('      --client "your-client-id" \\');
      console.log('      --secret "your-secret"');
      console.log();

      console.log('For more information, see the documentation in docs/DEVICE_COMPLIANCE_GUIDE.md');
    });

  // Parse arguments
  program.parse(process.argv);

  // Show help if no command provided
  if (!process.argv.slice(2).length) {
    program.outputHelp();
  }
}

// Run main function
if (require.main === module) {
  main().catch(error => {
    console.error('Fatal error:', error);
    process.exit(1);
  });
}

export { main };
