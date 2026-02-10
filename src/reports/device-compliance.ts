/**
 * Device Compliance Report Generator for Microsoft Intune
 *
 * This module provides functionality to generate device compliance reports
 * from Microsoft Intune using the Microsoft Graph API.
 *
 * Features:
 * - Authentication with Microsoft Graph API
 * - Fetch device compliance status
 * - Generate compliance summary reports
 * - Export to CSV and JSON formats
 * - Comprehensive error handling and retry logic
 *
 * @module device-compliance
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import * as fs from 'fs/promises';
import * as path from 'path';

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Configuration for authenticating with Microsoft Graph API
 */
export interface AuthConfig {
  /** Azure AD tenant ID */
  tenantId: string;
  /** Application (client) ID */
  clientId: string;
  /** Client secret for authentication */
  clientSecret: string;
}

/**
 * Device compliance status information from Intune
 */
export interface DeviceCompliance {
  /** Device ID */
  id: string;
  /** Device name */
  deviceName: string;
  /** User principal name of the device owner */
  userPrincipalName: string;
  /** Operating system of the device */
  operatingSystem: string;
  /** OS version */
  osVersion: string;
  /** Compliance state (compliant, noncompliant, etc.) */
  complianceState: string;
  /** Last time the device checked in */
  lastReportedDateTime: string;
  /** Device manufacturer */
  manufacturer?: string;
  /** Device model */
  model?: string;
  /** Whether the device is managed */
  isManaged: boolean;
  /** Device enrollment type */
  deviceEnrollmentType?: string;
}

/**
 * Summary statistics for device compliance
 */
export interface ComplianceSummary {
  /** Total number of devices */
  totalDevices: number;
  /** Number of compliant devices */
  compliantDevices: number;
  /** Number of non-compliant devices */
  nonCompliantDevices: number;
  /** Number of devices with unknown status */
  unknownDevices: number;
  /** Number of devices in grace period */
  inGracePeriodDevices: number;
  /** Number of devices with configuration errors */
  errorDevices: number;
  /** Compliance percentage */
  compliancePercentage: number;
  /** Summary by operating system */
  byOperatingSystem: Record<string, OSComplianceSummary>;
  /** Report generation timestamp */
  generatedAt: string;
}

/**
 * Compliance summary per operating system
 */
export interface OSComplianceSummary {
  total: number;
  compliant: number;
  nonCompliant: number;
  compliancePercentage: number;
}

/**
 * Options for report generation
 */
export interface ReportOptions {
  /** Output directory for reports */
  outputDir?: string;
  /** Whether to include detailed device information */
  includeDetails?: boolean;
  /** Filter by operating system */
  osFilter?: string;
  /** Filter by compliance state */
  complianceFilter?: string;
  /** Maximum number of retry attempts */
  maxRetries?: number;
  /** Delay between retries in milliseconds */
  retryDelay?: number;
}

/**
 * Result of report generation
 */
export interface ReportResult {
  /** Summary of compliance data */
  summary: ComplianceSummary;
  /** Detailed device compliance data */
  devices: DeviceCompliance[];
  /** Paths to generated files */
  files: {
    csv?: string;
    json?: string;
  };
}

// ============================================================================
// Main Class
// ============================================================================

/**
 * Device Compliance Report Generator
 *
 * Handles authentication, data fetching, and report generation for
 * Intune device compliance information.
 *
 * @example
 * ```typescript
 * const config: AuthConfig = {
 *   tenantId: 'your-tenant-id',
 *   clientId: 'your-client-id',
 *   clientSecret: 'your-client-secret'
 * };
 *
 * const reporter = new DeviceComplianceReporter(config);
 * const result = await reporter.generateReport({
 *   outputDir: './reports',
 *   includeDetails: true
 * });
 *
 * console.log(`Compliance: ${result.summary.compliancePercentage}%`);
 * ```
 */
export class DeviceComplianceReporter {
  private client: Client | null = null;
  private authConfig: AuthConfig;
  private readonly defaultRetries = 3;
  private readonly defaultRetryDelay = 1000; // 1 second

  /**
   * Creates a new DeviceComplianceReporter instance
   *
   * @param authConfig - Authentication configuration for Microsoft Graph API
   */
  constructor(authConfig: AuthConfig) {
    this.validateAuthConfig(authConfig);
    this.authConfig = authConfig;
  }

  // ==========================================================================
  // Authentication
  // ==========================================================================

  /**
   * Validates the authentication configuration
   *
   * @param config - Authentication configuration to validate
   * @throws Error if configuration is invalid
   */
  private validateAuthConfig(config: AuthConfig): void {
    if (!config.tenantId || config.tenantId.trim() === '') {
      throw new Error('Tenant ID is required');
    }
    if (!config.clientId || config.clientId.trim() === '') {
      throw new Error('Client ID is required');
    }
    if (!config.clientSecret || config.clientSecret.trim() === '') {
      throw new Error('Client secret is required');
    }
  }

  /**
   * Initializes and authenticates the Microsoft Graph client
   *
   * @throws Error if authentication fails
   */
  private async authenticate(): Promise<void> {
    try {
      const credential = new ClientSecretCredential(
        this.authConfig.tenantId,
        this.authConfig.clientId,
        this.authConfig.clientSecret
      );

      const authProvider = new TokenCredentialAuthenticationProvider(
        credential,
        {
          scopes: ['https://graph.microsoft.com/.default']
        }
      );

      this.client = Client.initWithMiddleware({
        authProvider: authProvider
      });

      // Test the connection
      await this.client.api('/me').get();
    } catch (error) {
      throw new Error(`Authentication failed: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Ensures the client is authenticated
   *
   * @throws Error if client initialization fails
   */
  private async ensureAuthenticated(): Promise<Client> {
    if (!this.client) {
      await this.authenticate();
    }
    if (!this.client) {
      throw new Error('Failed to initialize Graph client');
    }
    return this.client;
  }

  // ==========================================================================
  // Data Fetching with Retry Logic
  // ==========================================================================

  /**
   * Executes a function with retry logic
   *
   * @param fn - Async function to execute
   * @param maxRetries - Maximum number of retry attempts
   * @param retryDelay - Delay between retries in milliseconds
   * @returns Result of the function execution
   * @throws Error if all retry attempts fail
   */
  private async executeWithRetry<T>(
    fn: () => Promise<T>,
    maxRetries: number = this.defaultRetries,
    retryDelay: number = this.defaultRetryDelay
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await fn();
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        console.warn(
          `Attempt ${attempt}/${maxRetries} failed: ${lastError.message}`
        );

        if (attempt < maxRetries) {
          // Exponential backoff
          const delay = retryDelay * Math.pow(2, attempt - 1);
          console.log(`Retrying in ${delay}ms...`);
          await this.sleep(delay);
        }
      }
    }

    throw new Error(
      `Failed after ${maxRetries} attempts. Last error: ${lastError?.message}`
    );
  }

  /**
   * Fetches all managed devices from Intune
   *
   * @param options - Report options including retry configuration
   * @returns Array of device compliance information
   */
  private async fetchDevices(options: ReportOptions): Promise<DeviceCompliance[]> {
    const client = await this.ensureAuthenticated();
    const devices: DeviceCompliance[] = [];

    const fetchPage = async (url?: string): Promise<void> => {
      const request = url
        ? client.api(url)
        : client.api('/deviceManagement/managedDevices')
            .select([
              'id',
              'deviceName',
              'userPrincipalName',
              'operatingSystem',
              'osVersion',
              'complianceState',
              'lastSyncDateTime',
              'manufacturer',
              'model',
              'managementAgent',
              'deviceEnrollmentType'
            ].join(','))
            .top(999); // Maximum page size

      const response = await request.get();

      if (response.value && Array.isArray(response.value)) {
        for (const device of response.value) {
          const compliance: DeviceCompliance = {
            id: device.id || 'unknown',
            deviceName: device.deviceName || 'Unknown Device',
            userPrincipalName: device.userPrincipalName || 'Unknown User',
            operatingSystem: device.operatingSystem || 'Unknown',
            osVersion: device.osVersion || 'Unknown',
            complianceState: device.complianceState || 'unknown',
            lastReportedDateTime: device.lastSyncDateTime || new Date().toISOString(),
            manufacturer: device.manufacturer,
            model: device.model,
            isManaged: device.managementAgent !== 'unknown',
            deviceEnrollmentType: device.deviceEnrollmentType
          };

          // Apply filters if specified
          if (options.osFilter && !compliance.operatingSystem.toLowerCase().includes(options.osFilter.toLowerCase())) {
            continue;
          }
          if (options.complianceFilter && compliance.complianceState.toLowerCase() !== options.complianceFilter.toLowerCase()) {
            continue;
          }

          devices.push(compliance);
        }
      }

      // Handle pagination
      if (response['@odata.nextLink']) {
        await fetchPage(response['@odata.nextLink']);
      }
    };

    await this.executeWithRetry(
      () => fetchPage(),
      options.maxRetries || this.defaultRetries,
      options.retryDelay || this.defaultRetryDelay
    );

    return devices;
  }

  // ==========================================================================
  // Report Generation
  // ==========================================================================

  /**
   * Generates a compliance summary from device data
   *
   * @param devices - Array of device compliance information
   * @returns Compliance summary statistics
   */
  private generateSummary(devices: DeviceCompliance[]): ComplianceSummary {
    const summary: ComplianceSummary = {
      totalDevices: devices.length,
      compliantDevices: 0,
      nonCompliantDevices: 0,
      unknownDevices: 0,
      inGracePeriodDevices: 0,
      errorDevices: 0,
      compliancePercentage: 0,
      byOperatingSystem: {},
      generatedAt: new Date().toISOString()
    };

    // Count compliance states
    for (const device of devices) {
      const state = device.complianceState.toLowerCase();

      switch (state) {
        case 'compliant':
          summary.compliantDevices++;
          break;
        case 'noncompliant':
          summary.nonCompliantDevices++;
          break;
        case 'ingraceperiod':
          summary.inGracePeriodDevices++;
          break;
        case 'error':
          summary.errorDevices++;
          break;
        default:
          summary.unknownDevices++;
      }

      // Track by OS
      const os = device.operatingSystem;
      if (!summary.byOperatingSystem[os]) {
        summary.byOperatingSystem[os] = {
          total: 0,
          compliant: 0,
          nonCompliant: 0,
          compliancePercentage: 0
        };
      }

      summary.byOperatingSystem[os].total++;
      if (state === 'compliant') {
        summary.byOperatingSystem[os].compliant++;
      } else if (state === 'noncompliant') {
        summary.byOperatingSystem[os].nonCompliant++;
      }
    }

    // Calculate percentages
    if (summary.totalDevices > 0) {
      summary.compliancePercentage =
        Math.round((summary.compliantDevices / summary.totalDevices) * 100 * 100) / 100;

      for (const os in summary.byOperatingSystem) {
        const osData = summary.byOperatingSystem[os];
        osData.compliancePercentage =
          osData.total > 0
            ? Math.round((osData.compliant / osData.total) * 100 * 100) / 100
            : 0;
      }
    }

    return summary;
  }

  /**
   * Generates a comprehensive compliance report
   *
   * @param options - Report generation options
   * @returns Report result including summary and device data
   */
  async generateReport(options: ReportOptions = {}): Promise<ReportResult> {
    try {
      console.log('Fetching device compliance data from Intune...');
      const devices = await this.fetchDevices(options);

      console.log(`Retrieved ${devices.length} devices`);
      const summary = this.generateSummary(devices);

      const result: ReportResult = {
        summary,
        devices: options.includeDetails !== false ? devices : [],
        files: {}
      };

      // Export to files if output directory is specified
      if (options.outputDir) {
        await this.ensureDirectoryExists(options.outputDir);

        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const baseFileName = `device-compliance-${timestamp}`;

        // Export to JSON
        const jsonPath = path.join(options.outputDir, `${baseFileName}.json`);
        await this.exportToJSON(result, jsonPath);
        result.files.json = jsonPath;

        // Export to CSV
        const csvPath = path.join(options.outputDir, `${baseFileName}.csv`);
        await this.exportToCSV(devices, csvPath);
        result.files.csv = csvPath;

        console.log(`Reports generated:`);
        console.log(`  JSON: ${jsonPath}`);
        console.log(`  CSV: ${csvPath}`);
      }

      return result;
    } catch (error) {
      throw new Error(
        `Failed to generate compliance report: ${this.getErrorMessage(error)}`
      );
    }
  }

  // ==========================================================================
  // Export Functions
  // ==========================================================================

  /**
   * Exports report data to JSON format
   *
   * @param result - Report result to export
   * @param filePath - Output file path
   */
  private async exportToJSON(result: ReportResult, filePath: string): Promise<void> {
    try {
      const json = JSON.stringify(result, null, 2);
      await fs.writeFile(filePath, json, 'utf-8');
    } catch (error) {
      throw new Error(`Failed to export JSON: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Exports device data to CSV format
   *
   * @param devices - Array of device compliance information
   * @param filePath - Output file path
   */
  private async exportToCSV(devices: DeviceCompliance[], filePath: string): Promise<void> {
    try {
      const headers = [
        'Device ID',
        'Device Name',
        'User',
        'Operating System',
        'OS Version',
        'Compliance State',
        'Last Reported',
        'Manufacturer',
        'Model',
        'Is Managed',
        'Enrollment Type'
      ];

      const rows = devices.map(device => [
        this.escapeCsvValue(device.id),
        this.escapeCsvValue(device.deviceName),
        this.escapeCsvValue(device.userPrincipalName),
        this.escapeCsvValue(device.operatingSystem),
        this.escapeCsvValue(device.osVersion),
        this.escapeCsvValue(device.complianceState),
        this.escapeCsvValue(device.lastReportedDateTime),
        this.escapeCsvValue(device.manufacturer || ''),
        this.escapeCsvValue(device.model || ''),
        device.isManaged ? 'Yes' : 'No',
        this.escapeCsvValue(device.deviceEnrollmentType || '')
      ]);

      const csv = [
        headers.join(','),
        ...rows.map(row => row.join(','))
      ].join('\n');

      await fs.writeFile(filePath, csv, 'utf-8');
    } catch (error) {
      throw new Error(`Failed to export CSV: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Escapes a value for CSV format
   *
   * @param value - Value to escape
   * @returns Escaped CSV value
   */
  private escapeCsvValue(value: string): string {
    if (value.includes(',') || value.includes('"') || value.includes('\n')) {
      return `"${value.replace(/"/g, '""')}"`;
    }
    return value;
  }

  /**
   * Ensures a directory exists, creating it if necessary
   *
   * @param dirPath - Directory path to ensure exists
   */
  private async ensureDirectoryExists(dirPath: string): Promise<void> {
    try {
      await fs.access(dirPath);
    } catch {
      await fs.mkdir(dirPath, { recursive: true });
    }
  }

  // ==========================================================================
  // Utility Functions
  // ==========================================================================

  /**
   * Extracts error message from various error types
   *
   * @param error - Error object
   * @returns Error message string
   */
  private getErrorMessage(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }
    if (typeof error === 'string') {
      return error;
    }
    if (error && typeof error === 'object' && 'message' in error) {
      return String(error.message);
    }
    return 'Unknown error occurred';
  }

  /**
   * Utility function to sleep for a specified duration
   *
   * @param ms - Milliseconds to sleep
   */
  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Prints a formatted compliance summary to console
   *
   * @param summary - Compliance summary to print
   */
  static printSummary(summary: ComplianceSummary): void {
    console.log('\n=== Device Compliance Summary ===');
    console.log(`Generated: ${new Date(summary.generatedAt).toLocaleString()}`);
    console.log(`\nTotal Devices: ${summary.totalDevices}`);
    console.log(`Compliant: ${summary.compliantDevices} (${summary.compliancePercentage}%)`);
    console.log(`Non-Compliant: ${summary.nonCompliantDevices}`);
    console.log(`In Grace Period: ${summary.inGracePeriodDevices}`);
    console.log(`Unknown/Error: ${summary.unknownDevices + summary.errorDevices}`);

    console.log('\n=== Compliance by Operating System ===');
    for (const [os, data] of Object.entries(summary.byOperatingSystem)) {
      console.log(`\n${os}:`);
      console.log(`  Total: ${data.total}`);
      console.log(`  Compliant: ${data.compliant} (${data.compliancePercentage}%)`);
      console.log(`  Non-Compliant: ${data.nonCompliant}`);
    }
    console.log('\n================================\n');
  }
}

// ============================================================================
// Standalone Usage Functions
// ============================================================================

/**
 * Creates a device compliance reporter from environment variables
 *
 * Expects the following environment variables:
 * - AZURE_TENANT_ID
 * - AZURE_CLIENT_ID
 * - AZURE_CLIENT_SECRET
 *
 * @returns DeviceComplianceReporter instance
 * @throws Error if required environment variables are missing
 */
export function createReporterFromEnv(): DeviceComplianceReporter {
  const config: AuthConfig = {
    tenantId: process.env.AZURE_TENANT_ID || '',
    clientId: process.env.AZURE_CLIENT_ID || '',
    clientSecret: process.env.AZURE_CLIENT_SECRET || ''
  };

  return new DeviceComplianceReporter(config);
}

/**
 * Quick function to generate a compliance report
 *
 * @param authConfig - Authentication configuration
 * @param outputDir - Directory for output files
 * @returns Report result
 */
export async function generateComplianceReport(
  authConfig: AuthConfig,
  outputDir?: string
): Promise<ReportResult> {
  const reporter = new DeviceComplianceReporter(authConfig);
  return await reporter.generateReport({ outputDir, includeDetails: true });
}

// ============================================================================
// Example Usage (when run as a script)
// ============================================================================

/**
 * Example usage when file is run directly
 */
async function main(): Promise<void> {
  try {
    // Load configuration from environment variables
    const reporter = createReporterFromEnv();

    // Generate report
    const result = await reporter.generateReport({
      outputDir: './reports',
      includeDetails: true,
      maxRetries: 3,
      retryDelay: 1000
    });

    // Print summary
    DeviceComplianceReporter.printSummary(result.summary);

    console.log('Report generation completed successfully!');
    process.exit(0);
  } catch (error) {
    console.error('Error generating report:', error);
    process.exit(1);
  }
}

// Run main function if executed directly
if (require.main === module) {
  main();
}
