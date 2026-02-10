/**
 * Windows RT Sideloading Reports for Intune
 *
 * Replicates Configuration Manager Windows RT sideloading key reports using Microsoft Graph API
 * Provides detailed tracking and analysis of Windows RT sideloading keys
 *
 * @module WindowsRTSideloadingReports
 * @category Reports
 */

import { BaseReport } from '../base-report';
import { ReportData } from '../../types';
import { Logger } from '../../core/logger';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';
import { OutputFormatter } from '../../formatters/output-formatter';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Status of a sideloading key
 */
export type SideloadingKeyStatus = 'active' | 'expired' | 'revoked' | 'unknown';

/**
 * Sideloading key information from Microsoft Graph API
 */
export interface SideloadingKey {
  id: string;
  displayName?: string;
  value?: string;
  totalActivation: number;
  lastUpdatedDateTime?: string;
  description?: string;
  '@odata.type'?: string;
}

/**
 * Detailed sideloading key status information
 */
export interface SideloadingKeyDetailedStatus {
  keyId: string;
  keyValue: string;
  displayName: string;
  totalActivations: number;
  activationsUsed: number;
  activationsRemaining: number;
  lastUpdatedDate: string;
  lastUpdatedDaysAgo: number;
  status: SideloadingKeyStatus;
  statusReason: string;
  expirationWarning: boolean;
  usagePercentage: number;
  description: string;
  createdDateTime?: string;
  isNearExpiration: boolean;
  isHighUsage: boolean;
}

/**
 * Sideloading keys summary statistics
 */
export interface SideloadingKeysSummary {
  totalKeys: number;
  activeKeysCount: number;
  expiredKeysCount: number;
  revokedKeysCount: number;
  unknownKeysCount: number;
  keysByStatus: {
    status: SideloadingKeyStatus;
    count: number;
    percentage: number;
    keys: string[];
  }[];
  usageStatistics: {
    totalActivationsAcrossAllKeys: number;
    averageActivationsPerKey: number;
    mostUsedKey: {
      keyId: string;
      displayName: string;
      activations: number;
    } | null;
    leastUsedKey: {
      keyId: string;
      displayName: string;
      activations: number;
    } | null;
    averageUsagePercentage: number;
  };
  alerts: {
    expiringKeysCount: number;
    highUsageKeysCount: number;
    unusedKeysCount: number;
    expiringKeys: string[];
    highUsageKeys: string[];
    unusedKeys: string[];
  };
  lastUpdatedDateTime: string;
}

/**
 * Filter options for sideloading key reports
 */
export interface SideloadingKeyFilter {
  status?: SideloadingKeyStatus;
  minActivations?: number;
  maxActivations?: number;
  minUsagePercentage?: number;
  maxUsagePercentage?: number;
  includeExpired?: boolean;
  includeRevoked?: boolean;
}

/**
 * Export options for reports
 */
export interface ReportExportOptions {
  format: 'json' | 'csv' | 'html';
  outputDir: string;
  includeTimestamp?: boolean;
}

/**
 * Alert thresholds for sideloading keys
 */
export interface AlertThresholds {
  expirationWarningDays?: number;
  highUsagePercentage?: number;
  lowUsagePercentage?: number;
}

// ============================================================================
// Windows RT Sideloading Reports Class
// ============================================================================

/**
 * Windows RT Sideloading Reports
 *
 * Provides comprehensive reporting capabilities for Windows RT sideloading keys.
 * Replicates Configuration Manager sideloading key reports using Microsoft Graph API.
 *
 * Configuration Manager Reports Implemented:
 * - Report 35: Windows RT Sideloading Keys Detailed Status
 * - Report 36: Windows RT Sideloading Keys Summary
 */
export class WindowsRTSideloadingReports extends BaseReport {
  name = 'windows-rt-sideloading-reports';
  description = 'Windows RT sideloading key tracking and analysis';
  category = 'Device Management';
  enabled = true;

  // Default alert thresholds
  private readonly DEFAULT_EXPIRATION_WARNING_DAYS = 30;
  private readonly DEFAULT_HIGH_USAGE_PERCENTAGE = 80;
  private readonly DEFAULT_LOW_USAGE_PERCENTAGE = 10;

  constructor(graphClient: Client, config: AppConfig) {
    super(graphClient, config);
  }

  /**
   * Execute the complete sideloading key report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Windows RT Sideloading Reports');

    try {
      const summary = await this.getSideloadingKeysSummary();

      logger.info('Windows RT Sideloading Reports completed', {
        totalKeys: summary.totalKeys,
      });

      return {
        metadata: this.createMetadata(this.name, summary.totalKeys),
        data: summary,
        summary: {
          totalKeys: summary.totalKeys,
          activeKeys: summary.activeKeysCount,
          alerts: summary.alerts.expiringKeysCount + summary.alerts.highUsageKeysCount,
        },
      };
    } catch (error) {
      logger.error('Failed to execute Windows RT Sideloading Reports', error);
      throw error;
    }
  }

  /**
   * Get all sideloading keys from Microsoft Graph API
   *
   * @returns Array of sideloading keys
   */
  async getAllSideloadingKeys(): Promise<SideloadingKey[]> {
    logger.info('Fetching all sideloading keys');

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/sideLoadingKeys')
          .select([
            'id',
            'displayName',
            'value',
            'totalActivation',
            'lastUpdatedDateTime',
            'description',
          ])
          .top(999)
          .get()
      );

      const keys = await this.getAllPages<SideloadingKey>(response);

      logger.info('Sideloading keys retrieved', { count: keys.length });
      return keys;
    } catch (error) {
      logger.error('Failed to fetch sideloading keys', error);
      throw error;
    }
  }

  /**
   * Get a specific sideloading key by ID
   *
   * @param keyId The ID of the sideloading key
   * @returns Sideloading key details
   */
  async getSideloadingKeyById(keyId: string): Promise<SideloadingKey> {
    logger.info('Fetching sideloading key by ID', { keyId });

    try {
      const key = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/sideLoadingKeys/${keyId}`)
          .select([
            'id',
            'displayName',
            'value',
            'totalActivation',
            'lastUpdatedDateTime',
            'description',
          ])
          .get()
      );

      logger.info('Sideloading key retrieved', { keyId });
      return key;
    } catch (error) {
      logger.error('Failed to fetch sideloading key', { keyId, error });
      throw error;
    }
  }

  /**
   * Report 35: Windows RT Sideloading Keys Detailed Status
   *
   * Get detailed status information for a specific sideloading key
   *
   * @param keyId The ID of the sideloading key
   * @param thresholds Optional alert thresholds
   * @returns Detailed key status information
   */
  async getSideloadingKeyDetailedStatus(
    keyId: string,
    thresholds?: AlertThresholds
  ): Promise<SideloadingKeyDetailedStatus> {
    logger.info('Generating detailed status for sideloading key', { keyId });

    const key = await this.getSideloadingKeyById(keyId);

    // Determine key status
    const status = this.determineKeyStatus(key);
    const statusReason = this.getStatusReason(key, status);

    // Calculate usage metrics
    const totalActivations = key.totalActivation || 0;
    const maxActivations = 5000; // Windows RT default max activations
    const activationsUsed = totalActivations;
    const activationsRemaining = Math.max(0, maxActivations - activationsUsed);
    const usagePercentage = maxActivations > 0
      ? Math.round((activationsUsed / maxActivations) * 100 * 100) / 100
      : 0;

    // Calculate time metrics
    const lastUpdated = key.lastUpdatedDateTime
      ? new Date(key.lastUpdatedDateTime)
      : new Date();
    const now = new Date();
    const daysSinceUpdate = Math.floor(
      (now.getTime() - lastUpdated.getTime()) / (1000 * 60 * 60 * 24)
    );

    // Apply thresholds
    const expirationWarningDays = thresholds?.expirationWarningDays
      ?? this.DEFAULT_EXPIRATION_WARNING_DAYS;
    const highUsagePercentage = thresholds?.highUsagePercentage
      ?? this.DEFAULT_HIGH_USAGE_PERCENTAGE;

    // Check for warnings
    const expirationWarning = status === 'active' && daysSinceUpdate > 365 - expirationWarningDays;
    const isNearExpiration = expirationWarning;
    const isHighUsage = usagePercentage >= highUsagePercentage;

    const detailedStatus: SideloadingKeyDetailedStatus = {
      keyId: key.id,
      keyValue: this.maskKeyValue(key.value || 'N/A'),
      displayName: key.displayName || `Key ${key.id}`,
      totalActivations: maxActivations,
      activationsUsed,
      activationsRemaining,
      lastUpdatedDate: lastUpdated.toISOString(),
      lastUpdatedDaysAgo: daysSinceUpdate,
      status,
      statusReason,
      expirationWarning,
      usagePercentage,
      description: key.description || 'No description',
      isNearExpiration,
      isHighUsage,
    };

    logger.info('Detailed status generated', {
      keyId,
      status,
      usagePercentage,
      isNearExpiration,
      isHighUsage,
    });

    return detailedStatus;
  }

  /**
   * Report 36: Windows RT Sideloading Keys Summary
   *
   * Get aggregated summary of all sideloading keys
   *
   * @param filter Optional filter criteria
   * @param thresholds Optional alert thresholds
   * @returns Aggregated key status summary
   */
  async getSideloadingKeysSummary(
    filter?: SideloadingKeyFilter,
    thresholds?: AlertThresholds
  ): Promise<SideloadingKeysSummary> {
    logger.info('Generating sideloading keys summary');

    const keys = await this.getAllSideloadingKeys();

    // Apply filters
    let filteredKeys = keys;
    if (filter) {
      filteredKeys = this.applyFilters(keys, filter);
    }

    // Count keys by status
    const statusCounts = new Map<SideloadingKeyStatus, SideloadingKey[]>();
    filteredKeys.forEach((key) => {
      const status = this.determineKeyStatus(key);
      if (!statusCounts.has(status)) {
        statusCounts.set(status, []);
      }
      statusCounts.get(status)!.push(key);
    });

    // Build status breakdown
    const totalKeys = filteredKeys.length;
    const keysByStatus = Array.from(statusCounts.entries()).map(([status, statusKeys]) => ({
      status,
      count: statusKeys.length,
      percentage: totalKeys > 0
        ? Math.round((statusKeys.length / totalKeys) * 100 * 100) / 100
        : 0,
      keys: statusKeys.map((k) => k.displayName || k.id),
    }));

    // Calculate usage statistics
    const totalActivationsAcrossAllKeys = filteredKeys.reduce(
      (sum, key) => sum + (key.totalActivation || 0),
      0
    );
    const averageActivationsPerKey = totalKeys > 0
      ? Math.round((totalActivationsAcrossAllKeys / totalKeys) * 100) / 100
      : 0;

    // Find most and least used keys
    const sortedByUsage = [...filteredKeys].sort(
      (a, b) => (b.totalActivation || 0) - (a.totalActivation || 0)
    );

    const mostUsedKey = sortedByUsage.length > 0
      ? {
          keyId: sortedByUsage[0].id,
          displayName: sortedByUsage[0].displayName || sortedByUsage[0].id,
          activations: sortedByUsage[0].totalActivation || 0,
        }
      : null;

    const leastUsedKey = sortedByUsage.length > 0
      ? {
          keyId: sortedByUsage[sortedByUsage.length - 1].id,
          displayName: sortedByUsage[sortedByUsage.length - 1].displayName ||
            sortedByUsage[sortedByUsage.length - 1].id,
          activations: sortedByUsage[sortedByUsage.length - 1].totalActivation || 0,
        }
      : null;

    // Calculate average usage percentage
    const maxActivations = 5000;
    const averageUsagePercentage = totalKeys > 0
      ? Math.round((totalActivationsAcrossAllKeys / (totalKeys * maxActivations)) * 100 * 100) / 100
      : 0;

    // Generate alerts
    const expirationWarningDays = thresholds?.expirationWarningDays
      ?? this.DEFAULT_EXPIRATION_WARNING_DAYS;
    const highUsagePercentage = thresholds?.highUsagePercentage
      ?? this.DEFAULT_HIGH_USAGE_PERCENTAGE;
    const lowUsagePercentage = thresholds?.lowUsagePercentage
      ?? this.DEFAULT_LOW_USAGE_PERCENTAGE;

    const expiringKeys: string[] = [];
    const highUsageKeys: string[] = [];
    const unusedKeys: string[] = [];

    for (const key of filteredKeys) {
      const status = this.determineKeyStatus(key);

      // Check for expiring keys
      if (status === 'active' && key.lastUpdatedDateTime) {
        const lastUpdated = new Date(key.lastUpdatedDateTime);
        const now = new Date();
        const daysSinceUpdate = Math.floor(
          (now.getTime() - lastUpdated.getTime()) / (1000 * 60 * 60 * 24)
        );

        if (daysSinceUpdate > 365 - expirationWarningDays) {
          expiringKeys.push(key.displayName || key.id);
        }
      }

      // Check for high usage
      const activations = key.totalActivation || 0;
      const usagePercent = (activations / maxActivations) * 100;

      if (usagePercent >= highUsagePercentage) {
        highUsageKeys.push(key.displayName || key.id);
      }

      // Check for unused keys
      if (usagePercent <= lowUsagePercentage) {
        unusedKeys.push(key.displayName || key.id);
      }
    }

    const summary: SideloadingKeysSummary = {
      totalKeys,
      activeKeysCount: statusCounts.get('active')?.length || 0,
      expiredKeysCount: statusCounts.get('expired')?.length || 0,
      revokedKeysCount: statusCounts.get('revoked')?.length || 0,
      unknownKeysCount: statusCounts.get('unknown')?.length || 0,
      keysByStatus,
      usageStatistics: {
        totalActivationsAcrossAllKeys,
        averageActivationsPerKey,
        mostUsedKey,
        leastUsedKey,
        averageUsagePercentage,
      },
      alerts: {
        expiringKeysCount: expiringKeys.length,
        highUsageKeysCount: highUsageKeys.length,
        unusedKeysCount: unusedKeys.length,
        expiringKeys,
        highUsageKeys,
        unusedKeys,
      },
      lastUpdatedDateTime: new Date().toISOString(),
    };

    logger.info('Sideloading keys summary generated', {
      totalKeys,
      activeKeys: summary.activeKeysCount,
      expiredKeys: summary.expiredKeysCount,
      alertsCount: expiringKeys.length + highUsageKeys.length,
    });

    return summary;
  }

  /**
   * Export report to specified format
   *
   * @param reportData Report data to export
   * @param options Export options
   * @returns Path to exported file
   */
  async exportReport(reportData: ReportData, options: ReportExportOptions): Promise<string> {
    logger.info('Exporting sideloading key report', { format: options.format });

    const outputPath = await OutputFormatter.format(
      reportData,
      options.format,
      options.outputDir,
      options.includeTimestamp !== false
    );

    logger.info('Sideloading key report exported', { path: outputPath });
    return outputPath;
  }

  /**
   * Generate comprehensive sideloading key report with all sections
   *
   * @param filter Optional filter criteria
   * @param thresholds Optional alert thresholds
   * @returns Complete sideloading key report data
   */
  async generateComprehensiveReport(
    filter?: SideloadingKeyFilter,
    thresholds?: AlertThresholds
  ): Promise<ReportData> {
    logger.info('Generating comprehensive sideloading key report');

    const summary = await this.getSideloadingKeysSummary(filter, thresholds);
    const keys = await this.getAllSideloadingKeys();

    // Get detailed status for all keys
    const detailedStatuses = await Promise.all(
      keys.map((key) => this.getSideloadingKeyDetailedStatus(key.id, thresholds))
    );

    return {
      metadata: this.createMetadata(this.name, summary.totalKeys, { filter, thresholds }),
      data: {
        summary,
        detailedStatuses,
        keys,
      },
      summary: {
        totalKeys: summary.totalKeys,
        activeKeys: summary.activeKeysCount,
        expiredKeys: summary.expiredKeysCount,
        alerts: summary.alerts,
        generatedAt: new Date().toISOString(),
      },
    };
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Determine the status of a sideloading key
   */
  private determineKeyStatus(key: SideloadingKey): SideloadingKeyStatus {
    // Check if key is expired based on last updated date
    if (key.lastUpdatedDateTime) {
      const lastUpdated = new Date(key.lastUpdatedDateTime);
      const now = new Date();
      const daysSinceUpdate = Math.floor(
        (now.getTime() - lastUpdated.getTime()) / (1000 * 60 * 60 * 24)
      );

      // Keys typically expire after 1 year
      if (daysSinceUpdate > 365) {
        return 'expired';
      }
    }

    // Check if key has reached max activations
    const maxActivations = 5000;
    const activations = key.totalActivation || 0;

    if (activations >= maxActivations) {
      return 'expired';
    }

    // Check if key value is null or empty (might be revoked)
    if (!key.value) {
      return 'revoked';
    }

    // Default to active
    return 'active';
  }

  /**
   * Get status reason for a key
   */
  private getStatusReason(key: SideloadingKey, status: SideloadingKeyStatus): string {
    switch (status) {
      case 'active':
        return 'Key is currently active and available for use';

      case 'expired':
        if (key.totalActivation && key.totalActivation >= 5000) {
          return 'Key has reached maximum activation limit (5000)';
        }
        if (key.lastUpdatedDateTime) {
          const lastUpdated = new Date(key.lastUpdatedDateTime);
          const now = new Date();
          const daysSinceUpdate = Math.floor(
            (now.getTime() - lastUpdated.getTime()) / (1000 * 60 * 60 * 24)
          );
          return `Key has expired (${daysSinceUpdate} days since last update)`;
        }
        return 'Key has expired';

      case 'revoked':
        return 'Key has been revoked and is no longer valid';

      case 'unknown':
      default:
        return 'Key status could not be determined';
    }
  }

  /**
   * Mask sensitive key value for security
   */
  private maskKeyValue(value: string): string {
    if (!value || value === 'N/A') {
      return 'N/A';
    }

    // Show first 4 and last 4 characters, mask the rest
    if (value.length <= 8) {
      return '****';
    }

    const start = value.substring(0, 4);
    const end = value.substring(value.length - 4);
    const masked = '*'.repeat(Math.min(value.length - 8, 20));

    return `${start}${masked}${end}`;
  }

  /**
   * Apply filters to sideloading keys list
   */
  private applyFilters(
    keys: SideloadingKey[],
    filter: SideloadingKeyFilter
  ): SideloadingKey[] {
    let filtered = keys;

    if (filter.status) {
      filtered = filtered.filter((key) => this.determineKeyStatus(key) === filter.status);
    }

    if (filter.minActivations !== undefined) {
      filtered = filtered.filter((key) => (key.totalActivation || 0) >= filter.minActivations!);
    }

    if (filter.maxActivations !== undefined) {
      filtered = filtered.filter((key) => (key.totalActivation || 0) <= filter.maxActivations!);
    }

    if (filter.minUsagePercentage !== undefined) {
      const maxActivations = 5000;
      filtered = filtered.filter((key) => {
        const usagePercent = ((key.totalActivation || 0) / maxActivations) * 100;
        return usagePercent >= filter.minUsagePercentage!;
      });
    }

    if (filter.maxUsagePercentage !== undefined) {
      const maxActivations = 5000;
      filtered = filtered.filter((key) => {
        const usagePercent = ((key.totalActivation || 0) / maxActivations) * 100;
        return usagePercent <= filter.maxUsagePercentage!;
      });
    }

    if (filter.includeExpired === false) {
      filtered = filtered.filter((key) => this.determineKeyStatus(key) !== 'expired');
    }

    if (filter.includeRevoked === false) {
      filtered = filtered.filter((key) => this.determineKeyStatus(key) !== 'revoked');
    }

    return filtered;
  }
}

// ============================================================================
// Exports
// ============================================================================

export default WindowsRTSideloadingReports;
