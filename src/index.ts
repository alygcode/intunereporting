#!/usr/bin/env node

import { Command } from 'commander';
import { ReportOrchestrator } from './core/orchestrator';
import { ConfigManager } from './core/config';
import { Logger } from './core/logger';
import { SchedulerManager } from './core/scheduler';
import * as dotenv from 'dotenv';
import * as path from 'path';
import chalk from 'chalk';

// Load environment variables
dotenv.config();

const logger = Logger.getInstance();
const program = new Command();

/**
 * Main CLI Application for Intune Reporting Dashboard
 */
async function main() {
  program
    .name('intune-reports')
    .description('Intune Reporting Dashboard - Generate and schedule comprehensive Intune reports')
    .version('1.0.0');

  // Initialize command - sets up configuration
  program
    .command('init')
    .description('Initialize the reporting dashboard with default configuration')
    .option('-o, --output <path>', 'Output directory for configuration file', '.')
    .action(async (options) => {
      try {
        const configPath = path.join(options.output, 'intune-reports.config.yaml');
        await ConfigManager.createDefaultConfig(configPath);
        console.log(chalk.green(`✓ Configuration file created at: ${configPath}`));
        console.log(chalk.blue('Please update the configuration with your Azure AD credentials.'));
      } catch (error) {
        logger.error('Failed to initialize configuration', error);
        console.error(chalk.red(`Error: ${(error as Error).message}`));
        process.exit(1);
      }
    });

  // Run command - executes reports
  program
    .command('run')
    .description('Run one or more reports')
    .option('-c, --config <path>', 'Path to configuration file', './intune-reports.config.yaml')
    .option('-r, --reports <reports...>', 'Specific reports to run (comma-separated)', [])
    .option('-f, --format <format>', 'Output format (json, csv, html, all)', 'json')
    .option('-o, --output <path>', 'Output directory for reports', './reports')
    .option('-a, --all', 'Run all available reports', false)
    .action(async (options) => {
      try {
        logger.info('Starting report generation', { options });

        const config = await ConfigManager.load(options.config);
        const orchestrator = new ReportOrchestrator(config);

        const reports = options.all ? 'all' : options.reports;
        const formats = options.format === 'all'
          ? ['json', 'csv', 'html']
          : [options.format];

        console.log(chalk.blue('Initializing Intune Reporting Dashboard...'));
        await orchestrator.initialize();

        console.log(chalk.blue(`Running reports: ${reports === 'all' ? 'ALL' : reports.join(', ')}`));
        const results = await orchestrator.runReports(reports, formats, options.output);

        console.log(chalk.green('\n✓ Report generation completed successfully!\n'));
        console.log(chalk.bold('Results:'));
        results.forEach(result => {
          const status = result.success ? chalk.green('✓') : chalk.red('✗');
          console.log(`  ${status} ${result.reportName}: ${result.outputPath || result.error}`);
        });

        logger.info('Report generation completed', { results });
      } catch (error) {
        logger.error('Failed to run reports', error);
        console.error(chalk.red(`Error: ${(error as Error).message}`));
        process.exit(1);
      }
    });

  // Schedule command - sets up scheduled reporting
  program
    .command('schedule')
    .description('Start the scheduler for automated report generation')
    .option('-c, --config <path>', 'Path to configuration file', './intune-reports.config.yaml')
    .option('-d, --daemon', 'Run as daemon in background', false)
    .action(async (options) => {
      try {
        logger.info('Starting scheduler', { options });

        const config = await ConfigManager.load(options.config);
        const orchestrator = new ReportOrchestrator(config);
        await orchestrator.initialize();

        const scheduler = new SchedulerManager(orchestrator, config);

        console.log(chalk.blue('Starting report scheduler...'));
        scheduler.start();

        console.log(chalk.green('✓ Scheduler started successfully!'));
        console.log(chalk.blue('\nScheduled jobs:'));
        config.scheduler.schedules.forEach(schedule => {
          console.log(`  - ${schedule.name}: ${schedule.cron} (${schedule.reports.join(', ')})`);
        });

        if (!options.daemon) {
          console.log(chalk.yellow('\nPress Ctrl+C to stop the scheduler.\n'));

          // Keep the process running
          process.on('SIGINT', () => {
            console.log(chalk.yellow('\nStopping scheduler...'));
            scheduler.stop();
            console.log(chalk.green('✓ Scheduler stopped.'));
            process.exit(0);
          });
        }
      } catch (error) {
        logger.error('Failed to start scheduler', error);
        console.error(chalk.red(`Error: ${(error as Error).message}`));
        process.exit(1);
      }
    });

  // List command - shows available reports
  program
    .command('list')
    .description('List all available reports')
    .option('-c, --config <path>', 'Path to configuration file', './intune-reports.config.yaml')
    .action(async (options) => {
      try {
        const config = await ConfigManager.load(options.config);
        const orchestrator = new ReportOrchestrator(config);
        await orchestrator.initialize();

        const reports = orchestrator.listAvailableReports();

        console.log(chalk.bold('\nAvailable Reports:\n'));
        reports.forEach(report => {
          console.log(chalk.blue(`  ${report.name}`));
          console.log(`    Description: ${report.description}`);
          console.log(`    Category: ${report.category}`);
          console.log(`    Enabled: ${report.enabled ? chalk.green('Yes') : chalk.red('No')}`);
          console.log();
        });
      } catch (error) {
        logger.error('Failed to list reports', error);
        console.error(chalk.red(`Error: ${(error as Error).message}`));
        process.exit(1);
      }
    });

  // Status command - shows system status
  program
    .command('status')
    .description('Show system status and health checks')
    .option('-c, --config <path>', 'Path to configuration file', './intune-reports.config.yaml')
    .action(async (options) => {
      try {
        const config = await ConfigManager.load(options.config);
        const orchestrator = new ReportOrchestrator(config);

        console.log(chalk.bold('\nIntune Reporting Dashboard - Status\n'));
        console.log(chalk.blue('Initializing and checking connectivity...'));

        await orchestrator.initialize();
        const status = await orchestrator.getStatus();

        console.log(chalk.bold('\nAuthentication:'));
        console.log(`  Status: ${status.authenticated ? chalk.green('✓ Connected') : chalk.red('✗ Not Connected')}`);
        console.log(`  Tenant: ${status.tenantId || 'N/A'}`);

        console.log(chalk.bold('\nConfiguration:'));
        console.log(`  Reports Available: ${status.reportsAvailable}`);
        console.log(`  Reports Enabled: ${status.reportsEnabled}`);
        console.log(`  Scheduler: ${status.schedulerEnabled ? chalk.green('Enabled') : chalk.yellow('Disabled')}`);

        console.log(chalk.bold('\nOutput:'));
        console.log(`  Default Format: ${status.defaultFormat}`);
        console.log(`  Output Directory: ${status.outputDirectory}`);

        console.log();
      } catch (error) {
        logger.error('Failed to get status', error);
        console.error(chalk.red(`Error: ${(error as Error).message}`));
        process.exit(1);
      }
    });

  // Validate command - validates configuration
  program
    .command('validate')
    .description('Validate configuration file')
    .option('-c, --config <path>', 'Path to configuration file', './intune-reports.config.yaml')
    .action(async (options) => {
      try {
        console.log(chalk.blue(`Validating configuration: ${options.config}`));
        const config = await ConfigManager.load(options.config);
        const validation = ConfigManager.validate(config);

        if (validation.valid) {
          console.log(chalk.green('✓ Configuration is valid!'));
        } else {
          console.log(chalk.red('✗ Configuration has errors:'));
          validation.errors.forEach(error => {
            console.log(chalk.red(`  - ${error}`));
          });
          process.exit(1);
        }
      } catch (error) {
        logger.error('Failed to validate configuration', error);
        console.error(chalk.red(`Error: ${(error as Error).message}`));
        process.exit(1);
      }
    });

  await program.parseAsync(process.argv);
}

// Run the CLI application
main().catch(error => {
  logger.error('Unhandled error in main application', error);
  console.error(chalk.red(`Fatal error: ${error.message}`));
  process.exit(1);
});
