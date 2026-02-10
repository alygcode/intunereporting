/**
 * Security Baseline Report Generator for Microsoft Intune
 *
 * This module provides comprehensive functionality to generate security baseline reports
 * from Microsoft Intune using the Microsoft Graph API.
 *
 * Features:
 * - Get all security baselines (templates)
 * - Get baseline compliance per device
 * - Get baseline setting states
 * - Get non-compliant settings
 * - Get security score trends
 * - Export to JSON, CSV, and HTML formats
 * - Comprehensive error handling and retry logic
 * - Pagination support for large datasets
 *
 * @module security-baseline-report
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { BaseReport } from './base-report';
import { ReportData, AppConfig } from '../types';
import { Logger } from '../core/logger';
import { OutputFormatter } from '../formatters/output-formatter';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Security baseline template information
 */
export interface SecurityBaselineTemplate {
  /** Template ID */
  id: string;
  /** Display name of the template */
  displayName: string;
  /** Description of the template */
  description?: string;
  /** Template version */
  version?: string;
  /** Template type */
  templateType?: string;
  /** Template subtype */
  templateSubtype?: string;
  /** Published date */
  publishedDateTime?: string;
  /** Platform type (Windows10, etc.) */
  platformType?: string;
}

/**
 * Security baseline intent (applied baseline)
 */
export interface SecurityBaselineIntent {
  /** Intent ID */
  id: string;
  /** Display name */
  displayName: string;
  /** Description */
  description?: string;
  /** Template ID this is based on */
  templateId?: string;
  /** Last modified date */
  lastModifiedDateTime?: string;
  /** Creation date */
  createdDateTime?: string;
  /** Role scope tag IDs */
  roleScopeTagIds?: string[];
  /** Is assigned */
  isAssigned?: boolean;
}

/**
 * Device state for a security baseline
 */
export interface SecurityBaselineDeviceState {
  /** Device state ID */
  id: string;
  /** Device ID */
  deviceId?: string;
  /** Device display name */
  deviceDisplayName?: string;
  /** User principal name */
  userPrincipalName?: string;
  /** Compliance state */
  state?: SecurityBaselineComplianceState;
  /** Last reported date */
  lastReportedDateTime?: string;
  /** User name */
  userName?: string;
  /** Device model */
  deviceModel?: string;
  /** Platform type */
  platform?: string;
  /** Setting states count */
  settingStatesCount?: number;
}

/**
 * Device state summary for a security baseline
 */
export interface SecurityBaselineStateSummary {
  /** Unknown count */
  unknownCount?: number;
  /** Not applicable count */
  notApplicableCount?: number;
  /** Compliant count */
  compliantCount?: number;
  /** Remediated count */
  remediatedCount?: number;
  /** Non-compliant count */
  nonCompliantCount?: number;
  /** Error count */
  errorCount?: number;
  /** Conflict count */
  conflictCount?: number;
  /** Not assigned count */
  notAssignedCount?: number;
}

/**
 * Security baseline compliance state
 */
export type SecurityBaselineComplianceState =
  | 'unknown'
  | 'notApplicable'
  | 'compliant'
  | 'remediated'
  | 'nonCompliant'
  | 'error'
  | 'conflict'
  | 'notAssigned';

/**
 * Security baseline setting state per device
 */
export interface SecurityBaselineSettingState {
  /** Setting ID */
  id: string;
  /** Setting definition ID */
  settingDefinitionId?: string;
  /** Setting name */
  settingName?: string;
  /** Setting instance ID */
  settingInstanceId?: string;
  /** Device ID */
  deviceId?: string;
  /** Device name */
  deviceName?: string;
  /** User ID */
  userId?: string;
  /** User name */
  userName?: string;
  /** User principal name */
  userPrincipalName?: string;
  /** Current value */
  currentValue?: string;
  /** Expected value */
  expectedValue?: string;
  /** State */
  state?: SecurityBaselineComplianceState;
  /** Error code */
  errorCode?: string;
  /** Sources */
  sources?: Array<{
    id?: string;
    displayName?: string;
  }>;
}

/**
 * Non-compliant setting information
 */
export interface NonCompliantSetting {
  /** Setting name */
  settingName: string;
  /** Device name */
  deviceName: string;
  /** User principal name */
  userPrincipalName: string;
  /** Current value */
  currentValue: string;
  /** Expected value */
  expectedValue: string;
  /** Baseline name */
  baselineName: string;
  /** Error code */
  errorCode?: string;
  /** Last reported */
  lastReportedDateTime?: string;
}

/**
 * Security score trend data point
 */
export interface SecurityScoreTrendData {
  /** Date */
  date: string;
  /** Total devices */
  totalDevices: number;
  /** Compliant devices */
  compliantDevices: number;
  /** Non-compliant devices */
  nonCompliantDevices: number;
  /** Compliance percentage */
  compliancePercentage: number;
  /** Baseline name */
  baselineName: string;
}

/**
 * Comprehensive security baseline compliance report
 */
export interface SecurityBaselineComplianceReport {
  /** Report metadata */
  metadata: {
    generatedAt: string;
    totalBaselines: number;
    totalIntents: number;
    totalDevices: number;
  };
  /** All baseline templates */
  templates: SecurityBaselineTemplate[];
  /** All baseline intents (applied baselines) */
  intents: SecurityBaselineIntent[];
  /** Compliance per device across all baselines */
  deviceCompliance: SecurityBaselineDeviceState[];
  /** Non-compliant settings */
  nonCompliantSettings: NonCompliantSetting[];
  /** Summary statistics */
  summary: SecurityBaselineStateSummary;
  /** Security score trends */
  trends?: SecurityScoreTrendData[];
}

/**
 * Options for security baseline report generation
 */
export interface SecurityBaselineReportOptions {
  /** Include device-level details */
  includeDeviceDetails?: boolean;
  /** Include setting-level details */
  includeSettingDetails?: boolean;
  /** Filter by specific baseline template ID */
  templateId?: string;
  /** Filter by specific intent ID */
  intentId?: string;
  /** Include security score trends */
  includeTrends?: boolean;
  /** Number of days for trend analysis */
  trendDays?: number;
  /** Output directory for exported reports */
  outputDir?: string;
  /** Export formats */
  formats?: Array<'json' | 'csv' | 'html'>;
}

// ============================================================================
// Main Class
// ============================================================================

/**
 * Security Baseline Report Generator
 *
 * Provides comprehensive reporting on Microsoft Intune security baselines,
 * including compliance status, device states, and non-compliant settings.
 *
 * @example
 * ```typescript
 * const report = new SecurityBaselineReport(graphClient, config);
 * const result = await report.execute();
 *
 * // Or get specific data
 * const baselines = await report.getAllSecurityBaselines();
 * const compliance = await report.getBaselineCompliancePerDevice('intentId');
 * const nonCompliant = await report.getNonCompliantSettings();
 * ```
 */
export class SecurityBaselineReport extends BaseReport {
  name = 'security-baseline-report';
  description = 'Comprehensive security baseline compliance report for Microsoft Intune';
  category = 'Security';
  enabled = true;

  constructor(graphClient: Client, config: AppConfig) {
    super(graphClient, config);
  }

  // ==========================================================================
  // Main Report Execution
  // ==========================================================================

  /**
   * Executes the comprehensive security baseline report
   *
   * @returns Complete report data including all baselines, compliance, and trends
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Security Baseline Report');

    try {
      const report = await this.generateComprehensiveReport();

      const data = this.flattenReportForExport(report);

      const summary = {
        totalBaselines: report.metadata.totalBaselines,
        totalIntents: report.metadata.totalIntents,
        totalDevices: report.metadata.totalDevices,
        compliantDevices: report.summary.compliantCount || 0,
        nonCompliantDevices: report.summary.nonCompliantCount || 0,
        errorDevices: report.summary.errorCount || 0,
        compliancePercentage:
          report.metadata.totalDevices > 0
            ? (
                ((report.summary.compliantCount || 0) / report.metadata.totalDevices) *
                100
              ).toFixed(2) + '%'
            : '0%',
        nonCompliantSettingsCount: report.nonCompliantSettings.length,
      };

      logger.info('Security Baseline Report completed', {
        baselines: report.metadata.totalBaselines,
        devices: report.metadata.totalDevices,
      });

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute Security Baseline Report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Core API Methods
  // ==========================================================================

  /**
   * Get all security baseline templates from Intune
   *
   * Retrieves all available security baseline templates that can be applied
   * to devices in the organization.
   *
   * @returns Array of security baseline templates
   * @throws Error if API call fails
   */
  async getAllSecurityBaselines(): Promise<SecurityBaselineTemplate[]> {
    logger.info('Fetching all security baseline templates');

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/templates')
          .filter(
            "templateType eq 'securityBaseline' or templateSubtype eq 'securityBaseline'"
          )
          .select([
            'id',
            'displayName',
            'description',
            'version',
            'templateType',
            'templateSubtype',
            'publishedDateTime',
            'platformType',
          ])
          .get()
      );

      const templates = await this.getAllPages<SecurityBaselineTemplate>(response);

      logger.info(`Retrieved ${templates.length} security baseline templates`);
      return templates;
    } catch (error) {
      logger.error('Failed to fetch security baseline templates', error);
      throw new Error(`Failed to get security baselines: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Get all security baseline intents (applied baselines)
   *
   * Retrieves all security baseline intents that have been configured
   * and assigned in the organization.
   *
   * @param templateId - Optional filter by template ID
   * @returns Array of security baseline intents
   * @throws Error if API call fails
   */
  async getSecurityBaselineIntents(
    templateId?: string
  ): Promise<SecurityBaselineIntent[]> {
    logger.info('Fetching security baseline intents', { templateId });

    try {
      let apiCall = this.graphClient
        .api('/deviceManagement/intents')
        .filter("templateId ne null")
        .select([
          'id',
          'displayName',
          'description',
          'templateId',
          'lastModifiedDateTime',
          'createdDateTime',
          'roleScopeTagIds',
          'isAssigned',
        ]);

      if (templateId) {
        apiCall = apiCall.filter(`templateId eq '${templateId}'`);
      }

      const response = await this.retryGraphCall(() => apiCall.get());

      const intents = await this.getAllPages<SecurityBaselineIntent>(response);

      logger.info(`Retrieved ${intents.length} security baseline intents`);
      return intents;
    } catch (error) {
      logger.error('Failed to fetch security baseline intents', error);
      throw new Error(`Failed to get baseline intents: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Get baseline compliance per device for a specific intent
   *
   * Retrieves detailed device state information showing how each device
   * complies with a specific security baseline.
   *
   * @param intentId - The security baseline intent ID
   * @returns Array of device states
   * @throws Error if API call fails or intentId is invalid
   */
  async getBaselineCompliancePerDevice(
    intentId: string
  ): Promise<SecurityBaselineDeviceState[]> {
    if (!intentId) {
      throw new Error('Intent ID is required');
    }

    logger.info('Fetching baseline compliance per device', { intentId });

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/intents/${intentId}/deviceStates`)
          .select([
            'id',
            'deviceId',
            'deviceDisplayName',
            'userPrincipalName',
            'state',
            'lastReportedDateTime',
            'userName',
            'deviceModel',
            'platform',
          ])
          .get()
      );

      const deviceStates = await this.getAllPages<SecurityBaselineDeviceState>(response);

      logger.info(`Retrieved ${deviceStates.length} device states for intent ${intentId}`);
      return deviceStates;
    } catch (error) {
      logger.error('Failed to fetch device states', { intentId, error });
      throw new Error(
        `Failed to get device compliance for intent ${intentId}: ${this.getErrorMessage(error)}`
      );
    }
  }

  /**
   * Get device state summary for a specific intent
   *
   * Retrieves aggregate statistics showing overall compliance state
   * for a security baseline.
   *
   * @param intentId - The security baseline intent ID
   * @returns Summary of device states
   * @throws Error if API call fails or intentId is invalid
   */
  async getBaselineDeviceStateSummary(
    intentId: string
  ): Promise<SecurityBaselineStateSummary> {
    if (!intentId) {
      throw new Error('Intent ID is required');
    }

    logger.info('Fetching baseline device state summary', { intentId });

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/intents/${intentId}/deviceStateSummary`)
          .get()
      );

      const summary: SecurityBaselineStateSummary = {
        unknownCount: response.unknownCount || 0,
        notApplicableCount: response.notApplicableCount || 0,
        compliantCount: response.compliantCount || 0,
        remediatedCount: response.remediatedCount || 0,
        nonCompliantCount: response.nonCompliantCount || 0,
        errorCount: response.errorCount || 0,
        conflictCount: response.conflictCount || 0,
        notAssignedCount: response.notAssignedCount || 0,
      };

      logger.info('Retrieved device state summary', { intentId, summary });
      return summary;
    } catch (error) {
      logger.error('Failed to fetch device state summary', { intentId, error });
      throw new Error(
        `Failed to get state summary for intent ${intentId}: ${this.getErrorMessage(error)}`
      );
    }
  }

  /**
   * Get baseline setting states for a specific intent
   *
   * Retrieves detailed setting-level information for a security baseline,
   * showing which specific settings are non-compliant on which devices.
   *
   * @param intentId - The security baseline intent ID
   * @returns Array of setting states
   * @throws Error if API call fails or intentId is invalid
   */
  async getBaselineSettingStates(
    intentId: string
  ): Promise<SecurityBaselineSettingState[]> {
    if (!intentId) {
      throw new Error('Intent ID is required');
    }

    logger.info('Fetching baseline setting states', { intentId });

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/intents/${intentId}/deviceSettingStateSummaries`)
          .expand('settingStates')
          .get()
      );

      const summaries = await this.getAllPages<any>(response);

      // Flatten setting states from all summaries
      const settingStates: SecurityBaselineSettingState[] = [];

      for (const summary of summaries) {
        if (summary.settingStates && Array.isArray(summary.settingStates)) {
          for (const state of summary.settingStates) {
            settingStates.push({
              id: state.id || '',
              settingDefinitionId: summary.settingDefinitionId,
              settingName: summary.settingName || state.settingName,
              settingInstanceId: state.settingInstanceId,
              deviceId: state.deviceId,
              deviceName: state.deviceName,
              userId: state.userId,
              userName: state.userName,
              userPrincipalName: state.userPrincipalName,
              currentValue: state.currentValue,
              expectedValue: summary.intendedValue || state.intendedValue,
              state: state.state,
              errorCode: state.errorCode,
              sources: state.sources,
            });
          }
        }
      }

      logger.info(`Retrieved ${settingStates.length} setting states for intent ${intentId}`);
      return settingStates;
    } catch (error) {
      logger.error('Failed to fetch baseline setting states', { intentId, error });
      // Return empty array if endpoint not supported or fails
      logger.warn('Setting states may not be available for this baseline');
      return [];
    }
  }

  /**
   * Get all non-compliant settings across all security baselines
   *
   * Aggregates all non-compliant settings from all security baselines,
   * providing a comprehensive view of security issues.
   *
   * @returns Array of non-compliant settings with device and baseline information
   * @throws Error if API calls fail
   */
  async getNonCompliantSettings(): Promise<NonCompliantSetting[]> {
    logger.info('Fetching non-compliant settings across all baselines');

    try {
      const nonCompliantSettings: NonCompliantSetting[] = [];

      // Get all intents
      const intents = await this.getSecurityBaselineIntents();

      // Process each intent
      for (const intent of intents) {
        try {
          // Get device states for this intent
          const deviceStates = await this.getBaselineCompliancePerDevice(intent.id);

          // Filter non-compliant devices
          const nonCompliantDevices = deviceStates.filter(
            (state) => state.state === 'nonCompliant' || state.state === 'error'
          );

          if (nonCompliantDevices.length === 0) {
            continue;
          }

          // Try to get detailed setting states
          const settingStates = await this.getBaselineSettingStates(intent.id);

          // Match setting states with non-compliant devices
          for (const deviceState of nonCompliantDevices) {
            const deviceSettings = settingStates.filter(
              (setting) => setting.deviceId === deviceState.deviceId
            );

            if (deviceSettings.length > 0) {
              // Add each non-compliant setting
              for (const setting of deviceSettings) {
                if (
                  setting.state === 'nonCompliant' ||
                  setting.state === 'error' ||
                  setting.state === 'conflict'
                ) {
                  nonCompliantSettings.push({
                    settingName: setting.settingName || 'Unknown Setting',
                    deviceName: deviceState.deviceDisplayName || 'Unknown Device',
                    userPrincipalName: deviceState.userPrincipalName || 'Unknown User',
                    currentValue: setting.currentValue || 'Not Set',
                    expectedValue: setting.expectedValue || 'Unknown',
                    baselineName: intent.displayName,
                    errorCode: setting.errorCode,
                    lastReportedDateTime: deviceState.lastReportedDateTime,
                  });
                }
              }
            } else {
              // If no detailed settings available, add a summary entry
              nonCompliantSettings.push({
                settingName: 'Multiple Settings',
                deviceName: deviceState.deviceDisplayName || 'Unknown Device',
                userPrincipalName: deviceState.userPrincipalName || 'Unknown User',
                currentValue: 'Non-Compliant',
                expectedValue: 'Compliant',
                baselineName: intent.displayName,
                lastReportedDateTime: deviceState.lastReportedDateTime,
              });
            }
          }
        } catch (error) {
          logger.warn(`Failed to process intent ${intent.id}`, error);
          // Continue with next intent
        }
      }

      logger.info(`Found ${nonCompliantSettings.length} non-compliant settings`);
      return nonCompliantSettings;
    } catch (error) {
      logger.error('Failed to fetch non-compliant settings', error);
      throw new Error(`Failed to get non-compliant settings: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Get security score trends over time
   *
   * Analyzes security baseline compliance over a period of time to show
   * trends and improvements or degradations in security posture.
   *
   * Note: This is a simulated trend based on current state. For actual historical
   * data, you would need to store snapshots over time.
   *
   * @param days - Number of days to analyze (default: 30)
   * @returns Array of trend data points
   */
  async getSecurityScoreTrends(days: number = 30): Promise<SecurityScoreTrendData[]> {
    logger.info('Generating security score trends', { days });

    try {
      const trends: SecurityScoreTrendData[] = [];
      const intents = await this.getSecurityBaselineIntents();

      // For each intent, get current compliance state
      for (const intent of intents) {
        try {
          const summary = await this.getBaselineDeviceStateSummary(intent.id);

          const totalDevices =
            (summary.compliantCount || 0) +
            (summary.nonCompliantCount || 0) +
            (summary.errorCount || 0) +
            (summary.unknownCount || 0);

          if (totalDevices === 0) {
            continue;
          }

          const compliancePercentage =
            ((summary.compliantCount || 0) / totalDevices) * 100;

          // Current state
          trends.push({
            date: new Date().toISOString().split('T')[0],
            totalDevices,
            compliantDevices: summary.compliantCount || 0,
            nonCompliantDevices: summary.nonCompliantCount || 0,
            compliancePercentage: Math.round(compliancePercentage * 100) / 100,
            baselineName: intent.displayName,
          });

          // Note: Historical data would require stored snapshots
          // This is a placeholder for trend visualization
        } catch (error) {
          logger.warn(`Failed to get trend data for intent ${intent.id}`, error);
        }
      }

      logger.info(`Generated ${trends.length} trend data points`);
      return trends;
    } catch (error) {
      logger.error('Failed to generate security score trends', error);
      throw new Error(`Failed to get security trends: ${this.getErrorMessage(error)}`);
    }
  }

  // ==========================================================================
  // Comprehensive Report Generation
  // ==========================================================================

  /**
   * Generate a comprehensive security baseline report
   *
   * Combines all security baseline data into a single comprehensive report
   * including templates, intents, device compliance, and non-compliant settings.
   *
   * @param options - Report generation options
   * @returns Complete security baseline compliance report
   */
  async generateComprehensiveReport(
    options: SecurityBaselineReportOptions = {}
  ): Promise<SecurityBaselineComplianceReport> {
    logger.info('Generating comprehensive security baseline report', options);

    try {
      // Fetch all templates
      const templates = await this.getAllSecurityBaselines();

      // Fetch all intents
      let intents = await this.getSecurityBaselineIntents(options.templateId);

      // Filter by specific intent if requested
      if (options.intentId) {
        intents = intents.filter((intent) => intent.id === options.intentId);
      }

      // Collect device compliance across all intents
      const allDeviceStates: SecurityBaselineDeviceState[] = [];
      const combinedSummary: SecurityBaselineStateSummary = {
        unknownCount: 0,
        notApplicableCount: 0,
        compliantCount: 0,
        remediatedCount: 0,
        nonCompliantCount: 0,
        errorCount: 0,
        conflictCount: 0,
        notAssignedCount: 0,
      };

      if (options.includeDeviceDetails !== false) {
        for (const intent of intents) {
          try {
            const deviceStates = await this.getBaselineCompliancePerDevice(intent.id);
            allDeviceStates.push(...deviceStates);

            // Aggregate summary
            const summary = await this.getBaselineDeviceStateSummary(intent.id);
            combinedSummary.unknownCount! += summary.unknownCount || 0;
            combinedSummary.notApplicableCount! += summary.notApplicableCount || 0;
            combinedSummary.compliantCount! += summary.compliantCount || 0;
            combinedSummary.remediatedCount! += summary.remediatedCount || 0;
            combinedSummary.nonCompliantCount! += summary.nonCompliantCount || 0;
            combinedSummary.errorCount! += summary.errorCount || 0;
            combinedSummary.conflictCount! += summary.conflictCount || 0;
            combinedSummary.notAssignedCount! += summary.notAssignedCount || 0;
          } catch (error) {
            logger.warn(`Failed to get device states for intent ${intent.id}`, error);
          }
        }
      }

      // Get non-compliant settings if requested
      let nonCompliantSettings: NonCompliantSetting[] = [];
      if (options.includeSettingDetails !== false) {
        nonCompliantSettings = await this.getNonCompliantSettings();
      }

      // Get security trends if requested
      let trends: SecurityScoreTrendData[] | undefined;
      if (options.includeTrends) {
        trends = await this.getSecurityScoreTrends(options.trendDays || 30);
      }

      // Get unique device count
      const uniqueDevices = new Set(allDeviceStates.map((d) => d.deviceId));

      const report: SecurityBaselineComplianceReport = {
        metadata: {
          generatedAt: new Date().toISOString(),
          totalBaselines: templates.length,
          totalIntents: intents.length,
          totalDevices: uniqueDevices.size,
        },
        templates,
        intents,
        deviceCompliance: allDeviceStates,
        nonCompliantSettings,
        summary: combinedSummary,
        trends,
      };

      logger.info('Comprehensive report generated successfully', {
        templates: templates.length,
        intents: intents.length,
        devices: uniqueDevices.size,
        nonCompliantSettings: nonCompliantSettings.length,
      });

      return report;
    } catch (error) {
      logger.error('Failed to generate comprehensive report', error);
      throw new Error(
        `Failed to generate comprehensive report: ${this.getErrorMessage(error)}`
      );
    }
  }

  // ==========================================================================
  // Export Methods
  // ==========================================================================

  /**
   * Export security baseline report to specified formats
   *
   * @param report - The report to export
   * @param options - Export options including output directory and formats
   * @returns Object with file paths for each exported format
   */
  async exportReport(
    report: SecurityBaselineComplianceReport,
    options: SecurityBaselineReportOptions = {}
  ): Promise<{ [format: string]: string }> {
    logger.info('Exporting security baseline report', options);

    if (!options.outputDir) {
      throw new Error('Output directory is required for export');
    }

    const formats = options.formats || ['json', 'csv', 'html'];
    const exportedFiles: { [format: string]: string } = {};

    const reportData: ReportData = {
      metadata: this.createMetadata(
        'Security Baseline Report',
        report.deviceCompliance.length,
        {
          totalBaselines: report.metadata.totalBaselines,
          totalIntents: report.metadata.totalIntents,
          totalDevices: report.metadata.totalDevices,
        }
      ),
      data: this.flattenReportForExport(report),
      summary: {
        totalBaselines: report.metadata.totalBaselines,
        totalIntents: report.metadata.totalIntents,
        totalDevices: report.metadata.totalDevices,
        compliantDevices: report.summary.compliantCount,
        nonCompliantDevices: report.summary.nonCompliantCount,
        errorDevices: report.summary.errorCount,
        nonCompliantSettingsCount: report.nonCompliantSettings.length,
      },
    };

    for (const format of formats) {
      try {
        const filePath = await OutputFormatter.format(
          reportData,
          format as 'json' | 'csv' | 'html',
          options.outputDir,
          true
        );
        exportedFiles[format] = filePath;
        logger.info(`Exported ${format.toUpperCase()} report to ${filePath}`);
      } catch (error) {
        logger.error(`Failed to export ${format} format`, error);
        throw new Error(
          `Failed to export ${format} format: ${this.getErrorMessage(error)}`
        );
      }
    }

    return exportedFiles;
  }

  /**
   * Export non-compliant settings to CSV
   *
   * Creates a detailed CSV report specifically for non-compliant settings,
   * making it easy to identify and remediate security issues.
   *
   * @param settings - Non-compliant settings to export
   * @param outputPath - Output file path
   */
  async exportNonCompliantSettingsToCSV(
    settings: NonCompliantSetting[],
    outputPath: string
  ): Promise<void> {
    logger.info('Exporting non-compliant settings to CSV', { outputPath });

    try {
      const reportData: ReportData = {
        metadata: this.createMetadata('Non-Compliant Security Settings', settings.length),
        data: settings,
        summary: {
          totalNonCompliantSettings: settings.length,
          uniqueDevices: new Set(settings.map((s) => s.deviceName)).size,
          uniqueBaselines: new Set(settings.map((s) => s.baselineName)).size,
        },
      };

      await OutputFormatter.format(reportData, 'csv', outputPath, false);

      logger.info(`Non-compliant settings exported to ${outputPath}`);
    } catch (error) {
      logger.error('Failed to export non-compliant settings', error);
      throw new Error(
        `Failed to export non-compliant settings: ${this.getErrorMessage(error)}`
      );
    }
  }

  /**
   * Export security score trends to JSON
   *
   * @param trends - Trend data to export
   * @param outputPath - Output file path
   */
  async exportTrendsToJSON(
    trends: SecurityScoreTrendData[],
    outputPath: string
  ): Promise<void> {
    logger.info('Exporting security score trends to JSON', { outputPath });

    try {
      const reportData: ReportData = {
        metadata: this.createMetadata('Security Score Trends', trends.length),
        data: trends,
        summary: {
          totalDataPoints: trends.length,
          averageCompliance:
            trends.reduce((sum, t) => sum + t.compliancePercentage, 0) / trends.length || 0,
        },
      };

      await OutputFormatter.format(reportData, 'json', outputPath, false);

      logger.info(`Trends exported to ${outputPath}`);
    } catch (error) {
      logger.error('Failed to export trends', error);
      throw new Error(`Failed to export trends: ${this.getErrorMessage(error)}`);
    }
  }

  // ==========================================================================
  // Utility Methods
  // ==========================================================================

  /**
   * Flatten comprehensive report for export
   *
   * Converts the nested report structure into a flat array suitable
   * for CSV and tabular display.
   *
   * @param report - Comprehensive report
   * @returns Flattened data array
   */
  private flattenReportForExport(report: SecurityBaselineComplianceReport): any[] {
    const flattened: any[] = [];

    for (const device of report.deviceCompliance) {
      // Find the intent this device belongs to
      const intent = report.intents.find((i) =>
        report.deviceCompliance.some((d) => d.id === device.id)
      );

      // Find non-compliant settings for this device
      const deviceNonCompliantSettings = report.nonCompliantSettings.filter(
        (s) => s.deviceName === device.deviceDisplayName
      );

      flattened.push({
        deviceName: device.deviceDisplayName || 'Unknown',
        userPrincipalName: device.userPrincipalName || 'Unknown',
        complianceState: device.state || 'Unknown',
        platform: device.platform || 'Unknown',
        deviceModel: device.deviceModel || 'Unknown',
        baselineName: intent?.displayName || 'Unknown',
        lastReportedDateTime: device.lastReportedDateTime
          ? new Date(device.lastReportedDateTime).toLocaleString()
          : 'Never',
        nonCompliantSettingsCount: deviceNonCompliantSettings.length,
      });
    }

    return flattened;
  }

  /**
   * Extract error message from various error types
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
   * Print a formatted security baseline summary to console
   *
   * @param report - Security baseline report to print
   */
  static printSummary(report: SecurityBaselineComplianceReport): void {
    console.log('\n========================================');
    console.log('  Security Baseline Compliance Report');
    console.log('========================================\n');
    console.log(`Generated: ${new Date(report.metadata.generatedAt).toLocaleString()}\n`);

    console.log('--- Overview ---');
    console.log(`Total Baselines: ${report.metadata.totalBaselines}`);
    console.log(`Applied Baselines (Intents): ${report.metadata.totalIntents}`);
    console.log(`Total Devices: ${report.metadata.totalDevices}\n`);

    console.log('--- Compliance Summary ---');
    console.log(`Compliant Devices: ${report.summary.compliantCount}`);
    console.log(`Non-Compliant Devices: ${report.summary.nonCompliantCount}`);
    console.log(`Error Devices: ${report.summary.errorCount}`);
    console.log(`Unknown/Other: ${report.summary.unknownCount}\n`);

    const totalActive =
      (report.summary.compliantCount || 0) +
      (report.summary.nonCompliantCount || 0) +
      (report.summary.errorCount || 0);

    if (totalActive > 0) {
      const compliancePercentage =
        ((report.summary.compliantCount || 0) / totalActive) * 100;
      console.log(`Overall Compliance: ${compliancePercentage.toFixed(2)}%\n`);
    }

    console.log('--- Non-Compliant Settings ---');
    console.log(`Total Non-Compliant Settings: ${report.nonCompliantSettings.length}\n`);

    if (report.nonCompliantSettings.length > 0) {
      console.log('Top 10 Non-Compliant Settings:');
      const topSettings = report.nonCompliantSettings.slice(0, 10);
      topSettings.forEach((setting, idx) => {
        console.log(`  ${idx + 1}. ${setting.settingName}`);
        console.log(`     Device: ${setting.deviceName}`);
        console.log(`     Baseline: ${setting.baselineName}`);
        console.log(`     Expected: ${setting.expectedValue}, Current: ${setting.currentValue}\n`);
      });
    }

    if (report.trends && report.trends.length > 0) {
      console.log('--- Security Score Trends ---');
      report.trends.forEach((trend) => {
        console.log(`${trend.baselineName}:`);
        console.log(`  Date: ${trend.date}`);
        console.log(`  Compliance: ${trend.compliancePercentage}%`);
        console.log(
          `  Devices: ${trend.compliantDevices}/${trend.totalDevices} compliant\n`
        );
      });
    }

    console.log('========================================\n');
  }
}

// ============================================================================
// Standalone Usage Example
// ============================================================================

/**
 * Example usage when file is executed directly
 */
async function main(): Promise<void> {
  // This is an example of how to use the SecurityBaselineReport class
  // In practice, you would initialize this through the main application
  console.log('Security Baseline Report Module');
  console.log('This module should be used through the main Intune Reporting Dashboard');
  console.log('See documentation for usage examples');
}

// Run main function if executed directly
if (require.main === module) {
  main();
}

export default SecurityBaselineReport;
