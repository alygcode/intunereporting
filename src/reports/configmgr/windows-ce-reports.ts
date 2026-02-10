/**
 * Windows CE Device Reports for Configuration Manager
 *
 * IMPORTANT: Windows CE is a legacy platform that is NOT SUPPORTED in modern Microsoft Intune.
 * Microsoft ended support for Windows CE/Windows Mobile in 2019.
 *
 * This module provides placeholder implementations that return empty results with warnings
 * and migration guidance to modern Windows 10/11 devices.
 *
 * Configuration Manager Reports (Legacy - Not Applicable):
 * 3. Certificate issues on mobile devices (Windows CE) - Not supported
 * 4. Client deployment failure (Windows CE) - Not supported
 * 5. Client deployment status details (Windows CE) - Not supported
 * 6. Client deployment success (Windows CE) - Not supported
 * 7. Communication issues on mobile devices (Windows CE) - Not supported
 * 13. Health information for mobile devices (Windows CE) - Not supported
 * 14. Health summary for mobile devices (Windows CE) - Not supported
 * 19. Local client issues on mobile devices (Windows CE) - Not supported
 *
 * Migration Path:
 * Organizations still using Windows CE devices should migrate to:
 * - Windows 10 IoT Enterprise/Core for embedded/industrial scenarios
 * - Windows 11 for modern desktop experiences
 * - Android or iOS for mobile device scenarios
 *
 * See: docs/CONFIGMGR_WINDOWS_CE_MIGRATION.md for detailed migration guidance
 *
 * @module windows-ce-reports
 * @deprecated Windows CE is not supported in Microsoft Intune
 */

import { BaseReport } from '../base-report';
import { ReportData, AppConfig } from '../../types';
import { Client } from '@microsoft/microsoft-graph-client';
import { Logger } from '../../core/logger';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Platform support status
 */
export enum PlatformSupportStatus {
  NOT_SUPPORTED = 'not_supported',
  END_OF_LIFE = 'end_of_life',
  DEPRECATED = 'deprecated',
}

/**
 * Migration recommendation for unsupported platforms
 */
export interface MigrationRecommendation {
  currentPlatform: string;
  supportStatus: PlatformSupportStatus;
  endOfLifeDate: string;
  recommendedPlatforms: string[];
  migrationPath: string;
  additionalResources: string[];
  warning: string;
}

/**
 * Empty report result with warning
 */
export interface UnsupportedPlatformReport {
  platform: string;
  reportType: string;
  supportStatus: PlatformSupportStatus;
  warning: string;
  deviceCount: number;
  devices: never[];
  migrationRecommendation: MigrationRecommendation;
  generatedAt: Date;
}

/**
 * Filter options (not applicable for Windows CE)
 */
export interface WindowsCEFilterOptions {
  // No filters applicable for unsupported platform
}

/**
 * Export options
 */
export interface ExportOptions {
  format: 'json' | 'csv' | 'html';
  outputDirectory: string;
  includeTimestamp?: boolean;
}

// ============================================================================
// Constants
// ============================================================================

/**
 * Windows CE end-of-life information
 */
const WINDOWS_CE_EOL_INFO = {
  platform: 'Windows CE / Windows Mobile',
  endOfLifeDate: '2019-01-08',
  endOfSupportDate: '2019-01-08',
  finalVersion: 'Windows Embedded Compact 2013',
  microsoftAnnouncementUrl: 'https://support.microsoft.com/en-us/lifecycle/search?alpha=Windows%20Embedded%20Compact',
};

/**
 * Migration paths from Windows CE
 */
const MIGRATION_PATHS = {
  embedded: {
    target: 'Windows 10 IoT Enterprise / Windows 10 IoT Core',
    scenarios: ['Industrial automation', 'Kiosks', 'Point-of-sale systems', 'Medical devices'],
    url: 'https://www.microsoft.com/en-us/windows/windows-10-iot',
  },
  desktop: {
    target: 'Windows 10 / Windows 11',
    scenarios: ['General-purpose computing', 'Business workstations', 'Enterprise devices'],
    url: 'https://www.microsoft.com/en-us/windows/business',
  },
  mobile: {
    target: 'iOS / Android with Microsoft Intune',
    scenarios: ['Mobile workforce', 'Field service', 'BYOD programs'],
    url: 'https://www.microsoft.com/en-us/microsoft-365/enterprise-mobility-security',
  },
};

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Windows CE Reports (Unsupported Platform)
 *
 * This class provides placeholder implementations for Windows CE reports.
 * All methods return empty results with warnings and migration guidance.
 *
 * Windows CE is NOT supported in Microsoft Intune and reached end-of-life in January 2019.
 */
export class WindowsCEReports extends BaseReport {
  name = 'windows-ce-reports';
  description = 'Windows CE reports (platform not supported - returns empty results with migration guidance)';
  category = 'ConfigMgr-Style Reports (Legacy)';
  enabled = true;

  constructor(graphClient: Client, config: AppConfig) {
    super(graphClient, config);
  }

  /**
   * Execute the default Windows CE report
   * Returns warning about unsupported platform
   */
  async execute(): Promise<ReportData> {
    logger.warn('Windows CE reports requested - platform not supported in Intune');

    const migrationInfo = this.getMigrationRecommendation();

    return {
      metadata: this.createMetadata(this.name, 0, {
        warning: 'Windows CE is not supported in Microsoft Intune',
        endOfLife: WINDOWS_CE_EOL_INFO.endOfLifeDate,
      }),
      data: {
        platform: 'Windows CE',
        supportStatus: PlatformSupportStatus.NOT_SUPPORTED,
        message: 'Windows CE is not supported in Microsoft Intune. This platform reached end-of-life on January 8, 2019.',
        migrationRecommendation: migrationInfo,
      },
      summary: {
        platformSupported: false,
        deviceCount: 0,
        migrationRequired: true,
      },
    };
  }

  // ==========================================================================
  // Report 3: Certificate Issues on Mobile Devices (Windows CE)
  // ==========================================================================

  /**
   * Get certificate issues on Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getCertificateIssues(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Certificate issues report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Certificate Issues', {
      specificGuidance: 'Certificate management is not available for Windows CE devices in Intune. ' +
        'Migrate to modern Windows, iOS, or Android devices for certificate-based authentication.',
    });
  }

  // ==========================================================================
  // Report 4: Client Deployment Failure (Windows CE)
  // ==========================================================================

  /**
   * Get client deployment failures on Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getClientDeploymentFailures(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Client deployment failures report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Client Deployment Failures', {
      specificGuidance: 'Client deployment is not available for Windows CE devices in Intune. ' +
        'Modern device management is available for Windows 10/11, iOS, and Android.',
    });
  }

  // ==========================================================================
  // Report 5: Client Deployment Status Details (Windows CE)
  // ==========================================================================

  /**
   * Get client deployment status for Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getClientDeploymentStatus(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Client deployment status report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Client Deployment Status', {
      specificGuidance: 'Deployment status tracking is not available for Windows CE devices in Intune. ' +
        'Deploy modern Windows 10/11 devices for comprehensive deployment reporting.',
    });
  }

  // ==========================================================================
  // Report 6: Client Deployment Success (Windows CE)
  // ==========================================================================

  /**
   * Get successful client deployments on Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getClientDeploymentSuccess(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Client deployment success report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Client Deployment Success', {
      specificGuidance: 'Deployment success tracking is not available for Windows CE devices in Intune. ' +
        'Modern deployment tracking is available for supported platforms.',
    });
  }

  // ==========================================================================
  // Report 7: Communication Issues on Mobile Devices (Windows CE)
  // ==========================================================================

  /**
   * Get communication issues on Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getCommunicationIssues(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Communication issues report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Communication Issues', {
      specificGuidance: 'Communication monitoring is not available for Windows CE devices in Intune. ' +
        'Modern device communication is supported on Windows 10/11, iOS, and Android.',
    });
  }

  // ==========================================================================
  // Report 13: Health Information for Mobile Devices (Windows CE)
  // ==========================================================================

  /**
   * Get health information for Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getHealthInformation(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Health information report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Device Health Information', {
      specificGuidance: 'Health attestation and monitoring is not available for Windows CE devices in Intune. ' +
        'Comprehensive health reporting is available for modern Windows, iOS, and Android devices.',
    });
  }

  // ==========================================================================
  // Report 14: Health Summary for Mobile Devices (Windows CE)
  // ==========================================================================

  /**
   * Get health summary for Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getHealthSummary(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Health summary report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Device Health Summary', {
      specificGuidance: 'Health summary reporting is not available for Windows CE devices in Intune. ' +
        'Advanced health analytics are available for supported modern platforms.',
    });
  }

  // ==========================================================================
  // Report 19: Local Client Issues on Mobile Devices (Windows CE)
  // ==========================================================================

  /**
   * Get local client issues on Windows CE devices
   *
   * Returns empty results with warning - Windows CE not supported
   *
   * @param options - Filter options (not applicable)
   * @returns Empty report with migration guidance
   */
  async getLocalClientIssues(
    options?: WindowsCEFilterOptions
  ): Promise<UnsupportedPlatformReport> {
    logger.warn('Local client issues report requested for Windows CE - platform not supported');

    return this.createUnsupportedReport('Local Client Issues', {
      specificGuidance: 'Client issue tracking is not available for Windows CE devices in Intune. ' +
        'Modern troubleshooting and diagnostics are available for supported platforms.',
    });
  }

  // ==========================================================================
  // Helper Methods
  // ==========================================================================

  /**
   * Create an unsupported platform report
   */
  private createUnsupportedReport(
    reportType: string,
    options?: { specificGuidance?: string }
  ): UnsupportedPlatformReport {
    const baseWarning = `${reportType} report is not available for Windows CE devices. ` +
      'Windows CE reached end-of-life on January 8, 2019 and is not supported in Microsoft Intune.';

    const fullWarning = options?.specificGuidance
      ? `${baseWarning} ${options.specificGuidance}`
      : baseWarning;

    return {
      platform: 'Windows CE',
      reportType,
      supportStatus: PlatformSupportStatus.NOT_SUPPORTED,
      warning: fullWarning,
      deviceCount: 0,
      devices: [],
      migrationRecommendation: this.getMigrationRecommendation(),
      generatedAt: new Date(),
    };
  }

  /**
   * Get migration recommendation for Windows CE devices
   */
  getMigrationRecommendation(): MigrationRecommendation {
    return {
      currentPlatform: 'Windows CE / Windows Mobile',
      supportStatus: PlatformSupportStatus.END_OF_LIFE,
      endOfLifeDate: WINDOWS_CE_EOL_INFO.endOfLifeDate,
      recommendedPlatforms: [
        'Windows 10 IoT Enterprise - For embedded and industrial scenarios',
        'Windows 10 IoT Core - For small-footprint IoT devices',
        'Windows 10/11 - For general-purpose and business workstations',
        'iOS with Microsoft Intune - For mobile workforce scenarios',
        'Android with Microsoft Intune - For mobile workforce scenarios',
      ],
      migrationPath: this.buildMigrationPathGuidance(),
      additionalResources: [
        'Windows CE End-of-Life: ' + WINDOWS_CE_EOL_INFO.microsoftAnnouncementUrl,
        'Windows 10 IoT: ' + MIGRATION_PATHS.embedded.url,
        'Microsoft Intune: ' + MIGRATION_PATHS.mobile.url,
        'Migration Guide: docs/CONFIGMGR_WINDOWS_CE_MIGRATION.md',
      ],
      warning: 'CRITICAL: Windows CE devices cannot be managed by Microsoft Intune. ' +
        'Immediate migration to supported platforms is required for continued device management, ' +
        'security updates, and compliance requirements.',
    };
  }

  /**
   * Build detailed migration path guidance
   */
  private buildMigrationPathGuidance(): string {
    return `
MIGRATION PATH FROM WINDOWS CE:

1. EMBEDDED/INDUSTRIAL SCENARIOS:
   Target Platform: ${MIGRATION_PATHS.embedded.target}
   Use Cases: ${MIGRATION_PATHS.embedded.scenarios.join(', ')}
   Learn More: ${MIGRATION_PATHS.embedded.url}

2. DESKTOP/WORKSTATION SCENARIOS:
   Target Platform: ${MIGRATION_PATHS.desktop.target}
   Use Cases: ${MIGRATION_PATHS.desktop.scenarios.join(', ')}
   Learn More: ${MIGRATION_PATHS.desktop.url}

3. MOBILE WORKFORCE SCENARIOS:
   Target Platform: ${MIGRATION_PATHS.mobile.target}
   Use Cases: ${MIGRATION_PATHS.mobile.scenarios.join(', ')}
   Learn More: ${MIGRATION_PATHS.mobile.url}

NEXT STEPS:
1. Inventory all Windows CE devices in your environment
2. Identify use cases and application requirements
3. Select appropriate target platform(s) based on use case
4. Plan application migration/modernization
5. Deploy and test new devices
6. Migrate users and retire Windows CE devices

For detailed migration planning, see: docs/CONFIGMGR_WINDOWS_CE_MIGRATION.md
    `.trim();
  }

  /**
   * Get platform end-of-life information
   */
  getEndOfLifeInfo() {
    return {
      ...WINDOWS_CE_EOL_INFO,
      yearsOutOfSupport: new Date().getFullYear() - new Date(WINDOWS_CE_EOL_INFO.endOfLifeDate).getFullYear(),
      securityRisk: 'CRITICAL',
      complianceRisk: 'HIGH',
      recommendation: 'Immediate migration required',
    };
  }
}

// ============================================================================
// Convenience Functions
// ============================================================================

/**
 * Generate Windows CE platform status report
 */
export async function generateWindowsCEPlatformReport(
  graphClient: Client,
  config: AppConfig
): Promise<ReportData> {
  const report = new WindowsCEReports(graphClient, config);
  return await report.execute();
}

/**
 * Get migration guidance for Windows CE
 */
export function getWindowsCEMigrationGuidance(
  graphClient: Client,
  config: AppConfig
): MigrationRecommendation {
  const report = new WindowsCEReports(graphClient, config);
  return report.getMigrationRecommendation();
}

// Export all
export default WindowsCEReports;
