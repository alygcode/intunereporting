/**
 * Health Attestation Reports for Microsoft Intune
 *
 * Replicates Configuration Manager report: "List of devices by Health Attestation state"
 * Displays devices with Health Attestation Service attributes including:
 * - BitLocker status
 * - Secure Boot status
 * - Code Integrity
 * - Windows Defender status
 * - TPM version and attestation
 *
 * Microsoft Graph API endpoints used:
 * - GET /deviceManagement/managedDevices
 * - deviceHealthAttestationState property
 *
 * @module health-attestation-reports
 */

import { BaseReport } from '../base-report';
import { ReportData } from '../../types';
import { Logger } from '../../core/logger';
import { OutputFormatter } from '../../formatters/output-formatter';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Health Attestation Status
 */
export enum HealthAttestationStatus {
  SUPPORTED = 'supported',
  NOT_SUPPORTED = 'notSupported',
  UNKNOWN = 'unknown'
}

/**
 * BitLocker Status
 */
export enum BitLockerStatus {
  ENABLED = 'enabled',
  DISABLED = 'disabled',
  UNKNOWN = 'unknown'
}

/**
 * Attestation State
 */
export enum AttestationState {
  ENABLED = 'enabled',
  DISABLED = 'disabled',
  NOT_APPLICABLE = 'notApplicable',
  UNKNOWN = 'unknown'
}

/**
 * PCR Hash Algorithm
 */
export enum PcrHashAlgorithm {
  SHA1 = 'sha1',
  SHA256 = 'sha256',
  UNKNOWN = 'unknown'
}

/**
 * Health Score Range
 */
export enum HealthScoreRange {
  EXCELLENT = 'excellent', // 90-100
  GOOD = 'good',          // 75-89
  FAIR = 'fair',          // 50-74
  POOR = 'poor',          // 25-49
  CRITICAL = 'critical'   // 0-24
}

/**
 * Complete Device Health Attestation State
 * Based on Microsoft Graph API deviceHealthAttestationState
 */
export interface DeviceHealthAttestationState {
  // Device identification
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  osVersion: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  lastSyncDateTime: Date;

  // Health Attestation Support
  healthAttestationSupportedStatus: HealthAttestationStatus;

  // BitLocker
  bitLockerStatus: BitLockerStatus;

  // Boot Manager
  bootManagerVersion?: string;
  bootManagerSecurityVersion?: string;
  bootAppSecurityVersion?: string;

  // Code Integrity
  codeIntegrityCheckVersion?: string;
  codeIntegrity: AttestationState;
  codeIntegrityPolicy?: string;

  // Secure Boot
  secureBoot: AttestationState;
  secureBootConfigurationPolicyFingerPrint?: string;

  // Debugging
  bootDebugging: AttestationState;
  operatingSystemKernelDebugging: AttestationState;
  testSigning: AttestationState;

  // Safe Mode and Windows PE
  safeMode: AttestationState;
  windowsPE: AttestationState;

  // Early Launch Anti-Malware (ELAM)
  earlyLaunchAntiMalwareDriverProtection: AttestationState;

  // Virtual Secure Mode (VSM) / Credential Guard
  virtualSecureMode: AttestationState;

  // TPM (Trusted Platform Module)
  tpmVersion?: string;

  // PCR (Platform Configuration Registers)
  pcrHashAlgorithm: PcrHashAlgorithm;
  pcr0?: string;

  // Revision Lists
  bootRevisionListInfo?: string;
  operatingSystemRevListInfo?: string;

  // Attestation timestamps
  attestationIdentityKey?: string;
  issuedDateTime?: Date;

  // Health Score (calculated)
  healthScore: number;
  healthScoreRange: HealthScoreRange;

  // Issues found
  securityIssues: string[];
  recommendations: string[];
}

/**
 * Health Attestation Summary Statistics
 */
export interface HealthAttestationSummary {
  totalDevices: number;

  // Support Status
  supportedDevices: number;
  notSupportedDevices: number;
  unknownSupportDevices: number;

  // BitLocker
  bitLockerEnabled: number;
  bitLockerDisabled: number;
  bitLockerUnknown: number;
  bitLockerPercentage: number;

  // Secure Boot
  secureBootEnabled: number;
  secureBootDisabled: number;
  secureBootPercentage: number;

  // Code Integrity
  codeIntegrityEnabled: number;
  codeIntegrityDisabled: number;
  codeIntegrityPercentage: number;

  // Debugging States
  bootDebuggingEnabled: number;
  kernelDebuggingEnabled: number;
  testSigningEnabled: number;

  // ELAM
  elamEnabled: number;
  elamDisabled: number;
  elamPercentage: number;

  // Virtual Secure Mode
  vsmEnabled: number;
  vsmDisabled: number;
  vsmPercentage: number;

  // Health Scores
  averageHealthScore: number;
  excellentHealthDevices: number;
  goodHealthDevices: number;
  fairHealthDevices: number;
  poorHealthDevices: number;
  criticalHealthDevices: number;

  // TPM versions
  tpmVersions: Record<string, number>;

  // PCR Hash Algorithms
  pcrHashAlgorithms: Record<string, number>;

  // Platform breakdown
  byPlatform: Record<string, PlatformHealthSummary>;

  // Critical findings
  criticalIssues: HealthCriticalIssue[];

  // Generated timestamp
  generatedAt: Date;
}

/**
 * Platform-specific health summary
 */
export interface PlatformHealthSummary {
  total: number;
  averageHealthScore: number;
  bitLockerEnabled: number;
  secureBootEnabled: number;
  codeIntegrityEnabled: number;
  vsmEnabled: number;
}

/**
 * Critical health issue
 */
export interface HealthCriticalIssue {
  issueType: 'bitlocker_disabled' | 'secure_boot_disabled' | 'code_integrity_disabled' |
             'debugging_enabled' | 'test_signing_enabled' | 'elam_disabled' |
             'vsm_disabled' | 'attestation_not_supported';
  severity: 'critical' | 'high' | 'medium' | 'low';
  description: string;
  affectedDeviceCount: number;
  affectedPercentage: number;
  recommendation: string;
}

/**
 * Filter options for health attestation queries
 */
export interface HealthAttestationFilterOptions {
  platform?: string;
  bitLockerStatus?: BitLockerStatus;
  secureBootEnabled?: boolean;
  codeIntegrityEnabled?: boolean;
  minHealthScore?: number;
  maxHealthScore?: number;
  healthScoreRange?: HealthScoreRange;
  attestationSupported?: boolean;
  deviceName?: string;
  userPrincipalName?: string;
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Health Attestation Reports
 *
 * Comprehensive device health attestation reporting with:
 * - Device health attestation state from all managed devices
 * - BitLocker, Secure Boot, Code Integrity status
 * - TPM version and attestation details
 * - Health score calculation
 * - Filtering and export capabilities
 * - Security recommendations
 */
export class HealthAttestationReports extends BaseReport {
  name = 'health-attestation-reports';
  description = 'Device health attestation state including BitLocker, Secure Boot, Code Integrity, and TPM information';
  category = 'Security';
  enabled = true;

  /**
   * Execute the complete health attestation report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Health Attestation Report');

    try {
      const startTime = Date.now();

      // Fetch all device health attestation states
      const deviceHealthStates = await this.getAllDeviceHealthAttestationStates();

      // Generate summary statistics
      const summary = this.generateHealthAttestationSummary(deviceHealthStates);

      const duration = Date.now() - startTime;

      logger.info('Health Attestation Report completed', {
        duration: `${duration}ms`,
        devices: deviceHealthStates.length,
        averageHealthScore: summary.averageHealthScore
      });

      return {
        metadata: this.createMetadata(this.name, deviceHealthStates.length, {
          duration: `${duration}ms`,
          averageHealthScore: summary.averageHealthScore
        }),
        data: deviceHealthStates.map(d => ({ type: 'health-attestation', ...d })),
        summary
      };
    } catch (error) {
      logger.error('Failed to execute Health Attestation Report', error);
      throw error;
    }
  }

  /**
   * Get all device health attestation states
   */
  async getAllDeviceHealthAttestationStates(
    options?: HealthAttestationFilterOptions
  ): Promise<DeviceHealthAttestationState[]> {
    logger.info('Fetching device health attestation states');
    const deviceHealthStates: DeviceHealthAttestationState[] = [];

    try {
      // Fetch managed devices with health attestation state
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'operatingSystem',
            'osVersion',
            'manufacturer',
            'model',
            'serialNumber',
            'lastSyncDateTime',
            'deviceHealthAttestationState'
          ])
          .filter(`managementAgent eq 'mdm' or managementAgent eq 'easMdm' or managementAgent eq 'intuneClient'`)
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(response);
      logger.info(`Processing health attestation for ${devices.length} devices`);

      // Process devices in batches
      const batchSize = 50;
      for (let i = 0; i < devices.length; i += batchSize) {
        const batch = devices.slice(i, Math.min(i + batchSize, devices.length));

        const batchPromises = batch.map(async (device) => {
          try {
            const healthState = await this.parseDeviceHealthAttestation(device);

            // Apply filters if provided
            if (this.matchesHealthFilter(healthState, options)) {
              return healthState;
            }
          } catch (error) {
            logger.warn(`Failed to parse health attestation for device ${device.id}`, error);
          }
          return null;
        });

        const batchResults = await Promise.all(batchPromises);
        deviceHealthStates.push(...batchResults.filter(d => d !== null) as DeviceHealthAttestationState[]);
      }

      logger.info(`Processed ${deviceHealthStates.length} device health attestation states`);
      return deviceHealthStates;
    } catch (error) {
      logger.error('Error fetching device health attestation states', error);
      throw new Error(`Failed to fetch device health attestation: ${error.message}`);
    }
  }

  /**
   * Get health attestation state for a specific device
   */
  async getDeviceHealthAttestation(deviceId: string): Promise<DeviceHealthAttestationState | null> {
    logger.info(`Fetching health attestation for device ${deviceId}`);

    try {
      const device = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}`)
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'operatingSystem',
            'osVersion',
            'manufacturer',
            'model',
            'serialNumber',
            'lastSyncDateTime',
            'deviceHealthAttestationState'
          ])
          .get()
      );

      return await this.parseDeviceHealthAttestation(device);
    } catch (error) {
      logger.error(`Failed to fetch health attestation for device ${deviceId}`, error);
      return null;
    }
  }

  /**
   * Get devices with critical health issues
   */
  async getDevicesWithCriticalHealthIssues(): Promise<DeviceHealthAttestationState[]> {
    logger.info('Fetching devices with critical health issues');

    const allDevices = await this.getAllDeviceHealthAttestationStates();

    return allDevices.filter(device => {
      // Consider critical if health score is below 50 or has critical security issues
      return device.healthScore < 50 || device.securityIssues.length >= 3;
    });
  }

  /**
   * Get devices by health score range
   */
  async getDevicesByHealthScore(scoreRange: HealthScoreRange): Promise<DeviceHealthAttestationState[]> {
    logger.info(`Fetching devices with ${scoreRange} health score`);

    const allDevices = await this.getAllDeviceHealthAttestationStates();

    return allDevices.filter(device => device.healthScoreRange === scoreRange);
  }

  /**
   * Get devices with BitLocker disabled
   */
  async getDevicesWithBitLockerDisabled(): Promise<DeviceHealthAttestationState[]> {
    logger.info('Fetching devices with BitLocker disabled');

    return await this.getAllDeviceHealthAttestationStates({
      bitLockerStatus: BitLockerStatus.DISABLED
    });
  }

  /**
   * Get devices with Secure Boot disabled
   */
  async getDevicesWithSecureBootDisabled(): Promise<DeviceHealthAttestationState[]> {
    logger.info('Fetching devices with Secure Boot disabled');

    return await this.getAllDeviceHealthAttestationStates({
      secureBootEnabled: false
    });
  }

  /**
   * Export health attestation report to file
   */
  async exportReport(
    format: 'json' | 'csv' | 'html',
    outputDir: string,
    options?: HealthAttestationFilterOptions
  ): Promise<string> {
    logger.info(`Exporting health attestation report as ${format}`);

    const reportData = await this.execute();

    // Apply filters if provided
    if (options) {
      const filteredData = reportData.data.filter(device =>
        this.matchesHealthFilter(device as any, options)
      );
      reportData.data = filteredData;
      reportData.metadata.recordCount = filteredData.length;
    }

    const outputPath = await OutputFormatter.format(
      reportData,
      format,
      outputDir,
      true
    );

    logger.info(`Report exported to: ${outputPath}`);
    return outputPath;
  }

  // ============================================================================
  // Private Helper Methods
  // ============================================================================

  /**
   * Parse device health attestation from Graph API response
   */
  private async parseDeviceHealthAttestation(device: any): Promise<DeviceHealthAttestationState> {
    const hasState = device.deviceHealthAttestationState &&
                     Object.keys(device.deviceHealthAttestationState).length > 0;

    const attestationState = hasState ? device.deviceHealthAttestationState : {};

    // Parse attestation data
    const healthState: DeviceHealthAttestationState = {
      // Device identification
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      platform: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      manufacturer: device.manufacturer,
      model: device.model,
      serialNumber: device.serialNumber,
      lastSyncDateTime: device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : new Date(),

      // Health Attestation Support
      healthAttestationSupportedStatus: this.parseHealthAttestationStatus(
        attestationState.healthAttestationSupportedStatus
      ),

      // BitLocker
      bitLockerStatus: this.parseBitLockerStatus(attestationState.bitLockerStatus),

      // Boot Manager
      bootManagerVersion: attestationState.bootManagerVersion,
      bootManagerSecurityVersion: attestationState.bootManagerSecurityVersion,
      bootAppSecurityVersion: attestationState.bootAppSecurityVersion,

      // Code Integrity
      codeIntegrityCheckVersion: attestationState.codeIntegrityCheckVersion,
      codeIntegrity: this.parseAttestationState(attestationState.codeIntegrity),
      codeIntegrityPolicy: attestationState.codeIntegrityPolicy,

      // Secure Boot
      secureBoot: this.parseAttestationState(attestationState.secureBoot),
      secureBootConfigurationPolicyFingerPrint: attestationState.secureBootConfigurationPolicyFingerPrint,

      // Debugging
      bootDebugging: this.parseAttestationState(attestationState.bootDebugging),
      operatingSystemKernelDebugging: this.parseAttestationState(attestationState.operatingSystemKernelDebugging),
      testSigning: this.parseAttestationState(attestationState.testSigning),

      // Safe Mode and Windows PE
      safeMode: this.parseAttestationState(attestationState.safeMode),
      windowsPE: this.parseAttestationState(attestationState.windowsPE),

      // ELAM
      earlyLaunchAntiMalwareDriverProtection: this.parseAttestationState(
        attestationState.earlyLaunchAntiMalwareDriverProtection
      ),

      // Virtual Secure Mode
      virtualSecureMode: this.parseAttestationState(attestationState.virtualSecureMode),

      // TPM
      tpmVersion: attestationState.tpmVersion,

      // PCR
      pcrHashAlgorithm: this.parsePcrHashAlgorithm(attestationState.pcrHashAlgorithm),
      pcr0: attestationState.pcr0,

      // Revision Lists
      bootRevisionListInfo: attestationState.bootRevisionListInfo,
      operatingSystemRevListInfo: attestationState.operatingSystemRevListInfo,

      // Attestation metadata
      attestationIdentityKey: attestationState.attestationIdentityKey,
      issuedDateTime: attestationState.issuedDateTime ? new Date(attestationState.issuedDateTime) : undefined,

      // Health Score (calculated below)
      healthScore: 0,
      healthScoreRange: HealthScoreRange.CRITICAL,

      // Issues and recommendations (calculated below)
      securityIssues: [],
      recommendations: []
    };

    // Calculate health score and identify issues
    this.calculateHealthScore(healthState);
    this.identifySecurityIssues(healthState);
    this.generateRecommendations(healthState);

    return healthState;
  }

  /**
   * Calculate health score based on security posture
   * Score: 0-100 based on security features enabled/disabled
   */
  private calculateHealthScore(healthState: DeviceHealthAttestationState): void {
    let score = 0;
    let maxScore = 0;

    // Only calculate if attestation is supported
    if (healthState.healthAttestationSupportedStatus === HealthAttestationStatus.NOT_SUPPORTED) {
      healthState.healthScore = 0;
      healthState.healthScoreRange = HealthScoreRange.CRITICAL;
      return;
    }

    // BitLocker (20 points)
    maxScore += 20;
    if (healthState.bitLockerStatus === BitLockerStatus.ENABLED) {
      score += 20;
    }

    // Secure Boot (20 points)
    maxScore += 20;
    if (healthState.secureBoot === AttestationState.ENABLED) {
      score += 20;
    }

    // Code Integrity (15 points)
    maxScore += 15;
    if (healthState.codeIntegrity === AttestationState.ENABLED) {
      score += 15;
    }

    // Boot Debugging DISABLED (10 points - good if disabled)
    maxScore += 10;
    if (healthState.bootDebugging === AttestationState.DISABLED ||
        healthState.bootDebugging === AttestationState.NOT_APPLICABLE) {
      score += 10;
    }

    // Kernel Debugging DISABLED (10 points - good if disabled)
    maxScore += 10;
    if (healthState.operatingSystemKernelDebugging === AttestationState.DISABLED ||
        healthState.operatingSystemKernelDebugging === AttestationState.NOT_APPLICABLE) {
      score += 10;
    }

    // Test Signing DISABLED (10 points - good if disabled)
    maxScore += 10;
    if (healthState.testSigning === AttestationState.DISABLED ||
        healthState.testSigning === AttestationState.NOT_APPLICABLE) {
      score += 10;
    }

    // ELAM (10 points)
    maxScore += 10;
    if (healthState.earlyLaunchAntiMalwareDriverProtection === AttestationState.ENABLED) {
      score += 10;
    }

    // Virtual Secure Mode (5 points)
    maxScore += 5;
    if (healthState.virtualSecureMode === AttestationState.ENABLED) {
      score += 5;
    }

    // Calculate percentage
    healthState.healthScore = maxScore > 0 ? Math.round((score / maxScore) * 100) : 0;

    // Determine health score range
    if (healthState.healthScore >= 90) {
      healthState.healthScoreRange = HealthScoreRange.EXCELLENT;
    } else if (healthState.healthScore >= 75) {
      healthState.healthScoreRange = HealthScoreRange.GOOD;
    } else if (healthState.healthScore >= 50) {
      healthState.healthScoreRange = HealthScoreRange.FAIR;
    } else if (healthState.healthScore >= 25) {
      healthState.healthScoreRange = HealthScoreRange.POOR;
    } else {
      healthState.healthScoreRange = HealthScoreRange.CRITICAL;
    }
  }

  /**
   * Identify security issues based on health attestation state
   */
  private identifySecurityIssues(healthState: DeviceHealthAttestationState): void {
    const issues: string[] = [];

    if (healthState.healthAttestationSupportedStatus === HealthAttestationStatus.NOT_SUPPORTED) {
      issues.push('Health attestation is not supported on this device');
    }

    if (healthState.bitLockerStatus === BitLockerStatus.DISABLED) {
      issues.push('BitLocker encryption is disabled');
    }

    if (healthState.secureBoot === AttestationState.DISABLED) {
      issues.push('Secure Boot is disabled');
    }

    if (healthState.codeIntegrity === AttestationState.DISABLED) {
      issues.push('Code Integrity is disabled');
    }

    if (healthState.bootDebugging === AttestationState.ENABLED) {
      issues.push('Boot debugging is enabled (security risk)');
    }

    if (healthState.operatingSystemKernelDebugging === AttestationState.ENABLED) {
      issues.push('Kernel debugging is enabled (security risk)');
    }

    if (healthState.testSigning === AttestationState.ENABLED) {
      issues.push('Test signing is enabled (security risk)');
    }

    if (healthState.safeMode === AttestationState.ENABLED) {
      issues.push('Device is in Safe Mode');
    }

    if (healthState.windowsPE === AttestationState.ENABLED) {
      issues.push('Device is running Windows PE');
    }

    if (healthState.earlyLaunchAntiMalwareDriverProtection === AttestationState.DISABLED) {
      issues.push('Early Launch Anti-Malware (ELAM) is disabled');
    }

    if (healthState.virtualSecureMode === AttestationState.DISABLED) {
      issues.push('Virtual Secure Mode (Credential Guard) is disabled');
    }

    healthState.securityIssues = issues;
  }

  /**
   * Generate security recommendations based on identified issues
   */
  private generateRecommendations(healthState: DeviceHealthAttestationState): void {
    const recommendations: string[] = [];

    if (healthState.healthAttestationSupportedStatus === HealthAttestationStatus.NOT_SUPPORTED) {
      recommendations.push('Upgrade to Windows 10/11 with TPM 2.0 support');
      recommendations.push('Ensure device has a compatible Trusted Platform Module (TPM)');
    }

    if (healthState.bitLockerStatus === BitLockerStatus.DISABLED) {
      recommendations.push('Enable BitLocker drive encryption');
      recommendations.push('Verify TPM is available and properly configured');
    }

    if (healthState.secureBoot === AttestationState.DISABLED) {
      recommendations.push('Enable Secure Boot in UEFI/BIOS settings');
      recommendations.push('Ensure device supports UEFI firmware (not legacy BIOS)');
    }

    if (healthState.codeIntegrity === AttestationState.DISABLED) {
      recommendations.push('Enable Code Integrity protection');
      recommendations.push('Apply Windows Defender Application Control (WDAC) policies');
    }

    if (healthState.bootDebugging === AttestationState.ENABLED) {
      recommendations.push('Disable boot debugging (bcdedit /set debug off)');
    }

    if (healthState.operatingSystemKernelDebugging === AttestationState.ENABLED) {
      recommendations.push('Disable kernel debugging');
    }

    if (healthState.testSigning === AttestationState.ENABLED) {
      recommendations.push('Disable test signing (bcdedit /set testsigning off)');
    }

    if (healthState.earlyLaunchAntiMalwareDriverProtection === AttestationState.DISABLED) {
      recommendations.push('Enable Windows Defender Antivirus');
      recommendations.push('Ensure ELAM driver is properly loaded at boot');
    }

    if (healthState.virtualSecureMode === AttestationState.DISABLED) {
      recommendations.push('Enable Credential Guard (Virtual Secure Mode)');
      recommendations.push('Verify hardware virtualization support (VT-x/AMD-V)');
    }

    if (recommendations.length === 0) {
      recommendations.push('Device has a strong security posture');
      recommendations.push('Continue monitoring health attestation status');
    }

    healthState.recommendations = recommendations;
  }

  /**
   * Generate summary statistics for health attestation
   */
  private generateHealthAttestationSummary(
    deviceHealthStates: DeviceHealthAttestationState[]
  ): HealthAttestationSummary {
    const totalDevices = deviceHealthStates.length;

    // Initialize counters
    let supportedDevices = 0;
    let notSupportedDevices = 0;
    let unknownSupportDevices = 0;

    let bitLockerEnabled = 0;
    let bitLockerDisabled = 0;
    let bitLockerUnknown = 0;

    let secureBootEnabled = 0;
    let secureBootDisabled = 0;

    let codeIntegrityEnabled = 0;
    let codeIntegrityDisabled = 0;

    let bootDebuggingEnabled = 0;
    let kernelDebuggingEnabled = 0;
    let testSigningEnabled = 0;

    let elamEnabled = 0;
    let elamDisabled = 0;

    let vsmEnabled = 0;
    let vsmDisabled = 0;

    let excellentHealthDevices = 0;
    let goodHealthDevices = 0;
    let fairHealthDevices = 0;
    let poorHealthDevices = 0;
    let criticalHealthDevices = 0;

    const tpmVersions: Record<string, number> = {};
    const pcrHashAlgorithms: Record<string, number> = {};
    const byPlatform: Record<string, PlatformHealthSummary> = {};

    let totalHealthScore = 0;

    // Process each device
    for (const device of deviceHealthStates) {
      // Support status
      if (device.healthAttestationSupportedStatus === HealthAttestationStatus.SUPPORTED) {
        supportedDevices++;
      } else if (device.healthAttestationSupportedStatus === HealthAttestationStatus.NOT_SUPPORTED) {
        notSupportedDevices++;
      } else {
        unknownSupportDevices++;
      }

      // BitLocker
      if (device.bitLockerStatus === BitLockerStatus.ENABLED) {
        bitLockerEnabled++;
      } else if (device.bitLockerStatus === BitLockerStatus.DISABLED) {
        bitLockerDisabled++;
      } else {
        bitLockerUnknown++;
      }

      // Secure Boot
      if (device.secureBoot === AttestationState.ENABLED) {
        secureBootEnabled++;
      } else if (device.secureBoot === AttestationState.DISABLED) {
        secureBootDisabled++;
      }

      // Code Integrity
      if (device.codeIntegrity === AttestationState.ENABLED) {
        codeIntegrityEnabled++;
      } else if (device.codeIntegrity === AttestationState.DISABLED) {
        codeIntegrityDisabled++;
      }

      // Debugging
      if (device.bootDebugging === AttestationState.ENABLED) {
        bootDebuggingEnabled++;
      }
      if (device.operatingSystemKernelDebugging === AttestationState.ENABLED) {
        kernelDebuggingEnabled++;
      }
      if (device.testSigning === AttestationState.ENABLED) {
        testSigningEnabled++;
      }

      // ELAM
      if (device.earlyLaunchAntiMalwareDriverProtection === AttestationState.ENABLED) {
        elamEnabled++;
      } else if (device.earlyLaunchAntiMalwareDriverProtection === AttestationState.DISABLED) {
        elamDisabled++;
      }

      // VSM
      if (device.virtualSecureMode === AttestationState.ENABLED) {
        vsmEnabled++;
      } else if (device.virtualSecureMode === AttestationState.DISABLED) {
        vsmDisabled++;
      }

      // Health score ranges
      switch (device.healthScoreRange) {
        case HealthScoreRange.EXCELLENT:
          excellentHealthDevices++;
          break;
        case HealthScoreRange.GOOD:
          goodHealthDevices++;
          break;
        case HealthScoreRange.FAIR:
          fairHealthDevices++;
          break;
        case HealthScoreRange.POOR:
          poorHealthDevices++;
          break;
        case HealthScoreRange.CRITICAL:
          criticalHealthDevices++;
          break;
      }

      totalHealthScore += device.healthScore;

      // TPM versions
      if (device.tpmVersion) {
        tpmVersions[device.tpmVersion] = (tpmVersions[device.tpmVersion] || 0) + 1;
      }

      // PCR Hash Algorithms
      const pcrAlgo = device.pcrHashAlgorithm;
      pcrHashAlgorithms[pcrAlgo] = (pcrHashAlgorithms[pcrAlgo] || 0) + 1;

      // Platform breakdown
      if (!byPlatform[device.platform]) {
        byPlatform[device.platform] = {
          total: 0,
          averageHealthScore: 0,
          bitLockerEnabled: 0,
          secureBootEnabled: 0,
          codeIntegrityEnabled: 0,
          vsmEnabled: 0
        };
      }

      const platformSummary = byPlatform[device.platform];
      platformSummary.total++;
      platformSummary.averageHealthScore += device.healthScore;

      if (device.bitLockerStatus === BitLockerStatus.ENABLED) {
        platformSummary.bitLockerEnabled++;
      }
      if (device.secureBoot === AttestationState.ENABLED) {
        platformSummary.secureBootEnabled++;
      }
      if (device.codeIntegrity === AttestationState.ENABLED) {
        platformSummary.codeIntegrityEnabled++;
      }
      if (device.virtualSecureMode === AttestationState.ENABLED) {
        platformSummary.vsmEnabled++;
      }
    }

    // Calculate platform averages
    for (const platform in byPlatform) {
      const summary = byPlatform[platform];
      summary.averageHealthScore = summary.total > 0
        ? Math.round((summary.averageHealthScore / summary.total) * 100) / 100
        : 0;
    }

    // Calculate percentages
    const bitLockerPercentage = totalDevices > 0
      ? Math.round((bitLockerEnabled / totalDevices) * 100 * 100) / 100
      : 0;
    const secureBootPercentage = totalDevices > 0
      ? Math.round((secureBootEnabled / totalDevices) * 100 * 100) / 100
      : 0;
    const codeIntegrityPercentage = totalDevices > 0
      ? Math.round((codeIntegrityEnabled / totalDevices) * 100 * 100) / 100
      : 0;
    const elamPercentage = totalDevices > 0
      ? Math.round((elamEnabled / totalDevices) * 100 * 100) / 100
      : 0;
    const vsmPercentage = totalDevices > 0
      ? Math.round((vsmEnabled / totalDevices) * 100 * 100) / 100
      : 0;
    const averageHealthScore = totalDevices > 0
      ? Math.round((totalHealthScore / totalDevices) * 100) / 100
      : 0;

    // Identify critical issues
    const criticalIssues = this.identifyCriticalIssues({
      totalDevices,
      bitLockerEnabled,
      bitLockerDisabled,
      secureBootEnabled,
      secureBootDisabled,
      codeIntegrityEnabled,
      codeIntegrityDisabled,
      bootDebuggingEnabled,
      kernelDebuggingEnabled,
      testSigningEnabled,
      elamDisabled,
      vsmDisabled,
      notSupportedDevices
    });

    return {
      totalDevices,
      supportedDevices,
      notSupportedDevices,
      unknownSupportDevices,
      bitLockerEnabled,
      bitLockerDisabled,
      bitLockerUnknown,
      bitLockerPercentage,
      secureBootEnabled,
      secureBootDisabled,
      secureBootPercentage,
      codeIntegrityEnabled,
      codeIntegrityDisabled,
      codeIntegrityPercentage,
      bootDebuggingEnabled,
      kernelDebuggingEnabled,
      testSigningEnabled,
      elamEnabled,
      elamDisabled,
      elamPercentage,
      vsmEnabled,
      vsmDisabled,
      vsmPercentage,
      averageHealthScore,
      excellentHealthDevices,
      goodHealthDevices,
      fairHealthDevices,
      poorHealthDevices,
      criticalHealthDevices,
      tpmVersions,
      pcrHashAlgorithms,
      byPlatform,
      criticalIssues,
      generatedAt: new Date()
    };
  }

  /**
   * Identify critical health issues across all devices
   */
  private identifyCriticalIssues(stats: any): HealthCriticalIssue[] {
    const issues: HealthCriticalIssue[] = [];
    const totalDevices = stats.totalDevices;

    if (totalDevices === 0) return issues;

    // BitLocker disabled
    if (stats.bitLockerDisabled > 0) {
      const percentage = Math.round((stats.bitLockerDisabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'bitlocker_disabled',
        severity: percentage > 50 ? 'critical' : percentage > 25 ? 'high' : 'medium',
        description: `${stats.bitLockerDisabled} devices have BitLocker disabled`,
        affectedDeviceCount: stats.bitLockerDisabled,
        affectedPercentage: percentage,
        recommendation: 'Enable BitLocker on all devices to protect data at rest. Deploy BitLocker policies through Intune.'
      });
    }

    // Secure Boot disabled
    if (stats.secureBootDisabled > 0) {
      const percentage = Math.round((stats.secureBootDisabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'secure_boot_disabled',
        severity: percentage > 50 ? 'critical' : percentage > 25 ? 'high' : 'medium',
        description: `${stats.secureBootDisabled} devices have Secure Boot disabled`,
        affectedDeviceCount: stats.secureBootDisabled,
        affectedPercentage: percentage,
        recommendation: 'Enable Secure Boot in UEFI firmware settings to prevent unauthorized boot loaders.'
      });
    }

    // Code Integrity disabled
    if (stats.codeIntegrityDisabled > 0) {
      const percentage = Math.round((stats.codeIntegrityDisabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'code_integrity_disabled',
        severity: percentage > 50 ? 'critical' : percentage > 25 ? 'high' : 'medium',
        description: `${stats.codeIntegrityDisabled} devices have Code Integrity disabled`,
        affectedDeviceCount: stats.codeIntegrityDisabled,
        affectedPercentage: percentage,
        recommendation: 'Enable Code Integrity to ensure only trusted code runs on devices.'
      });
    }

    // Boot debugging enabled
    if (stats.bootDebuggingEnabled > 0) {
      const percentage = Math.round((stats.bootDebuggingEnabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'debugging_enabled',
        severity: 'high',
        description: `${stats.bootDebuggingEnabled} devices have boot debugging enabled`,
        affectedDeviceCount: stats.bootDebuggingEnabled,
        affectedPercentage: percentage,
        recommendation: 'Disable boot debugging on production devices (bcdedit /set debug off).'
      });
    }

    // Kernel debugging enabled
    if (stats.kernelDebuggingEnabled > 0) {
      const percentage = Math.round((stats.kernelDebuggingEnabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'debugging_enabled',
        severity: 'high',
        description: `${stats.kernelDebuggingEnabled} devices have kernel debugging enabled`,
        affectedDeviceCount: stats.kernelDebuggingEnabled,
        affectedPercentage: percentage,
        recommendation: 'Disable kernel debugging on production devices.'
      });
    }

    // Test signing enabled
    if (stats.testSigningEnabled > 0) {
      const percentage = Math.round((stats.testSigningEnabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'test_signing_enabled',
        severity: 'critical',
        description: `${stats.testSigningEnabled} devices have test signing enabled`,
        affectedDeviceCount: stats.testSigningEnabled,
        affectedPercentage: percentage,
        recommendation: 'Disable test signing immediately (bcdedit /set testsigning off). This is a critical security vulnerability.'
      });
    }

    // ELAM disabled
    if (stats.elamDisabled > 0) {
      const percentage = Math.round((stats.elamDisabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'elam_disabled',
        severity: percentage > 50 ? 'high' : 'medium',
        description: `${stats.elamDisabled} devices have Early Launch Anti-Malware disabled`,
        affectedDeviceCount: stats.elamDisabled,
        affectedPercentage: percentage,
        recommendation: 'Enable Windows Defender or another ELAM-compatible antivirus solution.'
      });
    }

    // VSM disabled
    if (stats.vsmDisabled > 0) {
      const percentage = Math.round((stats.vsmDisabled / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'vsm_disabled',
        severity: 'medium',
        description: `${stats.vsmDisabled} devices have Virtual Secure Mode (Credential Guard) disabled`,
        affectedDeviceCount: stats.vsmDisabled,
        affectedPercentage: percentage,
        recommendation: 'Enable Credential Guard for enhanced credential protection on supported devices.'
      });
    }

    // Attestation not supported
    if (stats.notSupportedDevices > 0) {
      const percentage = Math.round((stats.notSupportedDevices / totalDevices) * 100 * 100) / 100;
      issues.push({
        issueType: 'attestation_not_supported',
        severity: percentage > 25 ? 'high' : 'medium',
        description: `${stats.notSupportedDevices} devices do not support health attestation`,
        affectedDeviceCount: stats.notSupportedDevices,
        affectedPercentage: percentage,
        recommendation: 'Consider upgrading devices to Windows 10/11 with TPM 2.0 support for health attestation capabilities.'
      });
    }

    // Sort by severity and affected count
    issues.sort((a, b) => {
      const severityOrder = { critical: 0, high: 1, medium: 2, low: 3 };
      const severityDiff = severityOrder[a.severity] - severityOrder[b.severity];
      if (severityDiff !== 0) return severityDiff;
      return b.affectedDeviceCount - a.affectedDeviceCount;
    });

    return issues;
  }

  /**
   * Parse health attestation status
   */
  private parseHealthAttestationStatus(status: string): HealthAttestationStatus {
    if (!status) return HealthAttestationStatus.UNKNOWN;

    const statusLower = status.toLowerCase();
    if (statusLower.includes('support') && !statusLower.includes('not')) {
      return HealthAttestationStatus.SUPPORTED;
    }
    if (statusLower.includes('not') || statusLower.includes('unsupport')) {
      return HealthAttestationStatus.NOT_SUPPORTED;
    }
    return HealthAttestationStatus.UNKNOWN;
  }

  /**
   * Parse BitLocker status
   */
  private parseBitLockerStatus(status: string): BitLockerStatus {
    if (!status) return BitLockerStatus.UNKNOWN;

    const statusLower = status.toLowerCase();
    if (statusLower.includes('enable') || statusLower === 'on') {
      return BitLockerStatus.ENABLED;
    }
    if (statusLower.includes('disable') || statusLower === 'off') {
      return BitLockerStatus.DISABLED;
    }
    return BitLockerStatus.UNKNOWN;
  }

  /**
   * Parse generic attestation state
   */
  private parseAttestationState(state: string): AttestationState {
    if (!state) return AttestationState.UNKNOWN;

    const stateLower = state.toLowerCase();
    if (stateLower.includes('enable') || stateLower === 'on' || stateLower === 'true') {
      return AttestationState.ENABLED;
    }
    if (stateLower.includes('disable') || stateLower === 'off' || stateLower === 'false') {
      return AttestationState.DISABLED;
    }
    if (stateLower.includes('notapplicable') || stateLower.includes('not applicable')) {
      return AttestationState.NOT_APPLICABLE;
    }
    return AttestationState.UNKNOWN;
  }

  /**
   * Parse PCR hash algorithm
   */
  private parsePcrHashAlgorithm(algorithm: string): PcrHashAlgorithm {
    if (!algorithm) return PcrHashAlgorithm.UNKNOWN;

    const algoLower = algorithm.toLowerCase();
    if (algoLower.includes('sha256') || algoLower === 'sha-256') {
      return PcrHashAlgorithm.SHA256;
    }
    if (algoLower.includes('sha1') || algoLower === 'sha-1') {
      return PcrHashAlgorithm.SHA1;
    }
    return PcrHashAlgorithm.UNKNOWN;
  }

  /**
   * Check if device matches filter criteria
   */
  private matchesHealthFilter(
    device: DeviceHealthAttestationState,
    options?: HealthAttestationFilterOptions
  ): boolean {
    if (!options) return true;

    if (options.platform && device.platform.toLowerCase() !== options.platform.toLowerCase()) {
      return false;
    }

    if (options.bitLockerStatus && device.bitLockerStatus !== options.bitLockerStatus) {
      return false;
    }

    if (options.secureBootEnabled !== undefined) {
      const secureBootEnabled = device.secureBoot === AttestationState.ENABLED;
      if (secureBootEnabled !== options.secureBootEnabled) {
        return false;
      }
    }

    if (options.codeIntegrityEnabled !== undefined) {
      const codeIntegrityEnabled = device.codeIntegrity === AttestationState.ENABLED;
      if (codeIntegrityEnabled !== options.codeIntegrityEnabled) {
        return false;
      }
    }

    if (options.minHealthScore !== undefined && device.healthScore < options.minHealthScore) {
      return false;
    }

    if (options.maxHealthScore !== undefined && device.healthScore > options.maxHealthScore) {
      return false;
    }

    if (options.healthScoreRange && device.healthScoreRange !== options.healthScoreRange) {
      return false;
    }

    if (options.attestationSupported !== undefined) {
      const isSupported = device.healthAttestationSupportedStatus === HealthAttestationStatus.SUPPORTED;
      if (isSupported !== options.attestationSupported) {
        return false;
      }
    }

    if (options.deviceName &&
        !device.deviceName.toLowerCase().includes(options.deviceName.toLowerCase())) {
      return false;
    }

    if (options.userPrincipalName &&
        !device.userPrincipalName.toLowerCase().includes(options.userPrincipalName.toLowerCase())) {
      return false;
    }

    return true;
  }
}

// ============================================================================
// Export for use in other modules
// ============================================================================

export default HealthAttestationReports;
