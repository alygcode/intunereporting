import { ClientSecretCredential, ClientCertificateCredential } from '@azure/identity';
import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { AppConfig, Report, ReportResult, SystemStatus, ReportData } from '../types';
import { Logger } from './logger';
import { OutputFormatter } from '../formatters/output-formatter';
import * as fs from 'fs/promises';
import * as path from 'path';

// Import available reports
import { DeviceInventoryReport } from '../reports/device-inventory';
import { ComplianceSummaryReport } from '../reports/compliance-summary';
import { ApplicationInventoryReport } from '../reports/application-inventory';
import { UserDevicesReport } from '../reports/user-devices';
import { PolicyAssignmentsReport } from '../reports/policy-assignments';

const logger = Logger.getInstance();

/**
 * Report Orchestrator - Main coordinator for all report generation
 * Handles authentication, report execution, and output formatting
 */
export class ReportOrchestrator {
  private config: AppConfig;
  private graphClient?: Client;
  private reports: Map<string, Report>;
  private initialized: boolean = false;

  constructor(config: AppConfig) {
    this.config = config;
    this.reports = new Map();
  }

  /**
   * Initialize the orchestrator - authenticate and register reports
   */
  async initialize(): Promise<void> {
    if (this.initialized) {
      return;
    }

    logger.info('Initializing Report Orchestrator');

    try {
      // Authenticate to Microsoft Graph
      await this.authenticate();

      // Register available reports
      this.registerReports();

      this.initialized = true;
      logger.info('Report Orchestrator initialized successfully');
    } catch (error) {
      logger.error('Failed to initialize Report Orchestrator', error);
      throw error;
    }
  }

  /**
   * Authenticate to Microsoft Graph API
   */
  private async authenticate(): Promise<void> {
    logger.info('Authenticating to Microsoft Graph');

    const { tenantId, clientId, clientSecret, authMethod, certificatePath, certificateThumbprint } =
      this.config.authentication;

    let credential;

    switch (authMethod) {
      case 'clientSecret':
        if (!clientSecret) {
          throw new Error('Client secret is required for clientSecret authentication');
        }
        credential = new ClientSecretCredential(tenantId, clientId, clientSecret);
        break;

      case 'certificate':
        if (!certificatePath || !certificateThumbprint) {
          throw new Error('Certificate path and thumbprint are required for certificate authentication');
        }
        const certificateContent = await fs.readFile(certificatePath, 'utf-8');
        credential = new ClientCertificateCredential(
          tenantId,
          clientId,
          certificateContent,
          certificateThumbprint
        );
        break;

      default:
        throw new Error(`Unsupported authentication method: ${authMethod}`);
    }

    const authProvider = new TokenCredentialAuthenticationProvider(credential, {
      scopes: ['https://graph.microsoft.com/.default'],
    });

    this.graphClient = Client.initWithMiddleware({ authProvider });

    // Test authentication
    try {
      await this.graphClient.api('/organization').get();
      logger.info('Successfully authenticated to Microsoft Graph');
    } catch (error) {
      logger.error('Failed to authenticate to Microsoft Graph', error);
      throw new Error('Authentication failed. Please check your credentials.');
    }
  }

  /**
   * Register all available reports
   */
  private registerReports(): void {
    if (!this.graphClient) {
      throw new Error('Graph client not initialized');
    }

    const reportClasses = [
      DeviceInventoryReport,
      ComplianceSummaryReport,
      ApplicationInventoryReport,
      UserDevicesReport,
      PolicyAssignmentsReport,
    ];

    reportClasses.forEach((ReportClass) => {
      const report = new ReportClass(this.graphClient!, this.config);
      this.reports.set(report.name, report);
      logger.debug(`Registered report: ${report.name}`);
    });

    logger.info(`Registered ${this.reports.size} reports`);
  }

  /**
   * Run one or more reports
   */
  async runReports(
    reportNames: string | string[],
    formats: string[],
    outputDir: string
  ): Promise<ReportResult[]> {
    if (!this.initialized) {
      throw new Error('Orchestrator not initialized. Call initialize() first.');
    }

    // Ensure output directory exists
    await fs.mkdir(outputDir, { recursive: true });

    const reportsToRun = this.resolveReports(reportNames);
    const results: ReportResult[] = [];

    logger.info(`Running ${reportsToRun.length} reports`, {
      reports: reportsToRun.map((r) => r.name),
      formats,
    });

    for (const report of reportsToRun) {
      // Skip disabled reports
      if (!report.enabled) {
        logger.warn(`Skipping disabled report: ${report.name}`);
        continue;
      }

      logger.info(`Executing report: ${report.name}`);
      const startTime = Date.now();

      try {
        // Execute the report
        const reportData = await report.execute();
        const duration = Date.now() - startTime;

        // Generate output in requested formats
        for (const format of formats) {
          const outputPath = await OutputFormatter.format(
            reportData,
            format as 'json' | 'csv' | 'html',
            outputDir,
            this.config.output.includeTimestamp
          );

          results.push({
            reportName: report.name,
            success: true,
            format,
            outputPath,
            duration,
            recordCount: reportData.data.length,
          });

          logger.info(`Report generated: ${report.name}`, {
            format,
            outputPath,
            recordCount: reportData.data.length,
            duration,
          });
        }
      } catch (error) {
        const duration = Date.now() - startTime;
        logger.error(`Failed to generate report: ${report.name}`, error);

        results.push({
          reportName: report.name,
          success: false,
          format: formats.join(','),
          error: (error as Error).message,
          duration,
        });
      }
    }

    return results;
  }

  /**
   * Resolve which reports to run based on input
   */
  private resolveReports(reportNames: string | string[]): Report[] {
    if (reportNames === 'all') {
      return Array.from(this.reports.values()).filter(
        (report) =>
          this.config.reports.enabled.includes(report.name) ||
          (this.config.reports.enabled.length === 0 && !this.config.reports.disabled.includes(report.name))
      );
    }

    const names = Array.isArray(reportNames) ? reportNames : [reportNames];
    const reports: Report[] = [];

    for (const name of names) {
      const report = this.reports.get(name);
      if (!report) {
        logger.warn(`Report not found: ${name}`);
        continue;
      }
      reports.push(report);
    }

    return reports;
  }

  /**
   * List all available reports
   */
  listAvailableReports(): Report[] {
    return Array.from(this.reports.values());
  }

  /**
   * Get system status and health information
   */
  async getStatus(): Promise<SystemStatus> {
    const authenticated = this.initialized && !!this.graphClient;
    let tenantId: string | undefined;

    if (authenticated && this.graphClient) {
      try {
        const org = await this.graphClient.api('/organization').get();
        tenantId = org.value[0]?.id;
      } catch (error) {
        logger.error('Failed to get organization info', error);
      }
    }

    const allReports = Array.from(this.reports.values());
    const enabledReports = allReports.filter(
      (r) =>
        this.config.reports.enabled.includes(r.name) ||
        (this.config.reports.enabled.length === 0 && !this.config.reports.disabled.includes(r.name))
    );

    return {
      authenticated,
      tenantId,
      reportsAvailable: allReports.length,
      reportsEnabled: enabledReports.length,
      schedulerEnabled: this.config.scheduler?.enabled || false,
      defaultFormat: this.config.output.defaultFormat,
      outputDirectory: this.config.output.directory,
    };
  }

  /**
   * Get the Graph client instance
   */
  getGraphClient(): Client {
    if (!this.graphClient) {
      throw new Error('Graph client not initialized');
    }
    return this.graphClient;
  }
}
