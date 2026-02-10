/**
 * Security and Compliance Reporting for Microsoft Intune
 *
 * This module provides comprehensive security and compliance reporting including:
 * - Windows Update compliance status
 * - Security baseline compliance
 * - Microsoft Defender ATP status
 * - Encryption status (BitLocker for Windows, FileVault for macOS)
 * - Threat detection and remediation reports
 *
 * @module security-compliance
 * @author Intune Reporting System
 * @version 1.0.0
 */

import { Client } from '@microsoft/microsoft-graph-client';

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Windows Update compliance status for a device
 */
export interface WindowsUpdateCompliance {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  osVersion: string;
  lastSyncDateTime: Date;
  complianceState: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown';
  pendingUpdates: number;
  failedUpdates: number;
  securityUpdatesStatus: UpdateStatus;
  qualityUpdatesStatus: UpdateStatus;
  featureUpdatesStatus: UpdateStatus;
  lastUpdateCheckDateTime: Date;
  daysUntilUpdateDeadline: number;
}

/**
 * Update status details
 */
export interface UpdateStatus {
  state: 'upToDate' | 'pending' | 'failed' | 'installing' | 'unknown';
  availableVersion?: string;
  installedVersion?: string;
  errorCode?: string;
  errorDescription?: string;
}

/**
 * Security baseline compliance information
 */
export interface SecurityBaselineCompliance {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  baselineName: string;
  baselineVersion: string;
  state: 'compliant' | 'noncompliant' | 'conflict' | 'error' | 'unknown';
  settingsCount: number;
  compliantSettingsCount: number;
  noncompliantSettingsCount: number;
  conflictSettingsCount: number;
  errorSettingsCount: number;
  lastReportedDateTime: Date;
  noncompliantSettings?: SecurityBaselineSetting[];
}

/**
 * Individual security baseline setting
 */
export interface SecurityBaselineSetting {
  settingName: string;
  settingCategory: string;
  currentValue: string;
  expectedValue: string;
  state: 'compliant' | 'noncompliant' | 'conflict' | 'error';
  remediationAction?: string;
}

/**
 * Microsoft Defender ATP status
 */
export interface DefenderATPStatus {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  onboardingState: 'onboarded' | 'canBeOnboarded' | 'unsupported' | 'unknown';
  healthState: 'healthy' | 'inactive' | 'misconfigured' | 'unknown';
  riskScore: 'none' | 'informational' | 'low' | 'medium' | 'high';
  exposureScore: number;
  antivirusStatus: AntivirusStatus;
  firewallStatus: FirewallStatus;
  realTimeProtectionEnabled: boolean;
  networkInspectionEnabled: boolean;
  quickScanOverdue: boolean;
  fullScanOverdue: boolean;
  signatureOutOfDate: boolean;
  lastQuickScanDateTime?: Date;
  lastFullScanDateTime?: Date;
  lastReportedDateTime: Date;
}

/**
 * Antivirus status details
 */
export interface AntivirusStatus {
  enabled: boolean;
  updated: boolean;
  productVersion: string;
  engineVersion: string;
  signatureVersion: string;
  lastSignatureUpdateDateTime?: Date;
}

/**
 * Firewall status details
 */
export interface FirewallStatus {
  domainProfile: 'enabled' | 'disabled' | 'unknown';
  privateProfile: 'enabled' | 'disabled' | 'unknown';
  publicProfile: 'enabled' | 'disabled' | 'unknown';
}

/**
 * Encryption status for a device
 */
export interface EncryptionStatus {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: 'Windows' | 'macOS' | 'iOS' | 'Android' | 'Linux';
  encryptionMethod: string;
  encryptionState: 'encrypted' | 'notEncrypted' | 'encrypting' | 'decrypting' | 'suspended' | 'unknown';
  encryptionCompliance: 'compliant' | 'noncompliant' | 'unknown';
  // Windows BitLocker specific
  bitLockerStatus?: BitLockerStatus;
  // macOS FileVault specific
  fileVaultStatus?: FileVaultStatus;
  lastReportedDateTime: Date;
}

/**
 * BitLocker encryption status (Windows)
 */
export interface BitLockerStatus {
  protectionStatus: 'protected' | 'unprotected' | 'unknown';
  encryptionMethod: 'aes128' | 'aes256' | 'xtsAes128' | 'xtsAes256' | 'unknown';
  recoveryKeyBackupStatus: 'backedUp' | 'notBackedUp' | 'unknown';
  tpmPresent: boolean;
  tpmActivated: boolean;
  volumes: BitLockerVolume[];
}

/**
 * BitLocker volume information
 */
export interface BitLockerVolume {
  volumeType: 'operatingSystem' | 'fixedData' | 'removableData';
  mountPoint: string;
  encryptionPercentage: number;
  protectionStatus: 'protected' | 'unprotected';
}

/**
 * FileVault encryption status (macOS)
 */
export interface FileVaultStatus {
  enabled: boolean;
  recoveryKeyType: 'personal' | 'institutional' | 'none';
  recoveryKeyEscrowed: boolean;
  encryptionInProgress: boolean;
  deferralCount?: number;
}

/**
 * Threat detection report
 */
export interface ThreatDetectionReport {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  threatCount: number;
  activeThreatCount: number;
  remediatedThreatCount: number;
  threats: ThreatDetails[];
  lastScanDateTime?: Date;
  lastReportedDateTime: Date;
}

/**
 * Individual threat details
 */
export interface ThreatDetails {
  threatId: string;
  threatName: string;
  category: 'malware' | 'virus' | 'trojan' | 'ransomware' | 'spyware' | 'adware' | 'pua' | 'other';
  severity: 'low' | 'medium' | 'high' | 'critical';
  state: 'active' | 'remediated' | 'quarantined' | 'removed' | 'allowed' | 'blocked';
  detectionDateTime: Date;
  remediationDateTime?: Date;
  affectedFiles: string[];
  remediationAction?: 'remove' | 'quarantine' | 'clean' | 'allow' | 'block' | 'userDefined';
  executionState?: 'blocked' | 'running' | 'notRunning';
}

/**
 * Aggregated security compliance summary
 */
export interface SecurityComplianceSummary {
  totalDevices: number;
  reportGeneratedDate: Date;

  // Windows Update Summary
  updateCompliance: {
    compliant: number;
    noncompliant: number;
    unknown: number;
    devicesWithPendingUpdates: number;
    devicesWithFailedUpdates: number;
  };

  // Security Baseline Summary
  baselineCompliance: {
    compliant: number;
    noncompliant: number;
    unknown: number;
    baselines: Map<string, { compliant: number; noncompliant: number }>;
  };

  // Defender ATP Summary
  defenderATP: {
    onboarded: number;
    notOnboarded: number;
    healthy: number;
    unhealthy: number;
    highRiskDevices: number;
    antivirusDisabled: number;
    signatureOutOfDate: number;
  };

  // Encryption Summary
  encryption: {
    encrypted: number;
    notEncrypted: number;
    encrypting: number;
    unknown: number;
    bitLockerCompliant: number;
    fileVaultCompliant: number;
  };

  // Threat Summary
  threats: {
    devicesWithThreats: number;
    totalThreats: number;
    activeThreats: number;
    remediatedThreats: number;
    criticalThreats: number;
    highSeverityThreats: number;
  };
}

/**
 * Filter options for security compliance reports
 */
export interface SecurityComplianceFilters {
  deviceIds?: string[];
  userPrincipalNames?: string[];
  complianceState?: ('compliant' | 'noncompliant' | 'unknown')[];
  platforms?: ('Windows' | 'macOS' | 'iOS' | 'Android')[];
  riskLevel?: ('none' | 'low' | 'medium' | 'high')[];
  dateFrom?: Date;
  dateTo?: Date;
}

/**
 * Export format options
 */
export type ExportFormat = 'json' | 'csv' | 'html' | 'pdf';

// ============================================================================
// Security Compliance Reporting Class
// ============================================================================

/**
 * Main class for security and compliance reporting
 */
export class SecurityComplianceReporter {
  private graphClient: Client;

  /**
   * Initialize the security compliance reporter
   * @param graphClient - Authenticated Microsoft Graph client
   */
  constructor(graphClient: Client) {
    this.graphClient = graphClient;
  }

  // ==========================================================================
  // Windows Update Compliance
  // ==========================================================================

  /**
   * Get Windows Update compliance status for all devices
   * @param filters - Optional filters to apply
   * @returns Array of Windows Update compliance information
   */
  async getWindowsUpdateCompliance(
    filters?: SecurityComplianceFilters
  ): Promise<WindowsUpdateCompliance[]> {
    try {
      const devices: WindowsUpdateCompliance[] = [];

      // Get device compliance policies
      let endpoint = '/deviceManagement/deviceCompliancePolicies';
      const compliancePolicies = await this.graphClient
        .api(endpoint)
        .get();

      // Get managed devices
      endpoint = '/deviceManagement/managedDevices';
      let query = this.graphClient.api(endpoint)
        .select([
          'id',
          'deviceName',
          'userPrincipalName',
          'osVersion',
          'lastSyncDateTime',
          'complianceState',
          'operatingSystem'
        ].join(','))
        .filter("operatingSystem eq 'Windows'");

      const managedDevices = await query.get();

      // For each Windows device, get update compliance details
      for (const device of managedDevices.value || []) {
        try {
          // Get Windows update states
          const updateStates = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/windowsProtectionState`)
            .get()
            .catch(() => null);

          // Get device compliance details
          const complianceDetails = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/deviceCompliancePolicyStates`)
            .get()
            .catch(() => ({ value: [] }));

          const windowsUpdateData: WindowsUpdateCompliance = {
            deviceId: device.id,
            deviceName: device.deviceName || 'Unknown',
            userPrincipalName: device.userPrincipalName || 'Unknown',
            osVersion: device.osVersion || 'Unknown',
            lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
            complianceState: device.complianceState || 'unknown',
            pendingUpdates: updateStates?.productStatus === 'notUpToDate' ? 1 : 0,
            failedUpdates: 0,
            securityUpdatesStatus: this.parseUpdateStatus(updateStates?.securityUpdatesStatus),
            qualityUpdatesStatus: this.parseUpdateStatus(updateStates?.qualityUpdatesStatus),
            featureUpdatesStatus: this.parseUpdateStatus(updateStates?.featureUpdatesStatus),
            lastUpdateCheckDateTime: new Date(updateStates?.lastCheckDateTime || Date.now()),
            daysUntilUpdateDeadline: this.calculateDaysUntilDeadline(updateStates),
          };

          // Apply filters
          if (this.matchesFilters(windowsUpdateData, filters)) {
            devices.push(windowsUpdateData);
          }
        } catch (error) {
          console.error(`Error processing device ${device.id}:`, error);
        }
      }

      return devices;
    } catch (error) {
      console.error('Error fetching Windows Update compliance:', error);
      throw new Error(`Failed to fetch Windows Update compliance: ${error.message}`);
    }
  }

  /**
   * Get detailed Windows Update status for a specific device
   * @param deviceId - The device ID
   * @returns Windows Update compliance information
   */
  async getDeviceWindowsUpdateStatus(deviceId: string): Promise<WindowsUpdateCompliance> {
    try {
      const device = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}`)
        .select('id,deviceName,userPrincipalName,osVersion,lastSyncDateTime,complianceState')
        .get();

      const updateStates = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/windowsProtectionState`)
        .get()
        .catch(() => null);

      return {
        deviceId: device.id,
        deviceName: device.deviceName || 'Unknown',
        userPrincipalName: device.userPrincipalName || 'Unknown',
        osVersion: device.osVersion || 'Unknown',
        lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
        complianceState: device.complianceState || 'unknown',
        pendingUpdates: updateStates?.productStatus === 'notUpToDate' ? 1 : 0,
        failedUpdates: 0,
        securityUpdatesStatus: this.parseUpdateStatus(updateStates?.securityUpdatesStatus),
        qualityUpdatesStatus: this.parseUpdateStatus(updateStates?.qualityUpdatesStatus),
        featureUpdatesStatus: this.parseUpdateStatus(updateStates?.featureUpdatesStatus),
        lastUpdateCheckDateTime: new Date(updateStates?.lastCheckDateTime || Date.now()),
        daysUntilUpdateDeadline: this.calculateDaysUntilDeadline(updateStates),
      };
    } catch (error) {
      console.error(`Error fetching Windows Update status for device ${deviceId}:`, error);
      throw new Error(`Failed to fetch Windows Update status: ${error.message}`);
    }
  }

  // ==========================================================================
  // Security Baseline Compliance
  // ==========================================================================

  /**
   * Get security baseline compliance for all devices
   * @param filters - Optional filters to apply
   * @returns Array of security baseline compliance information
   */
  async getSecurityBaselineCompliance(
    filters?: SecurityComplianceFilters
  ): Promise<SecurityBaselineCompliance[]> {
    try {
      const complianceData: SecurityBaselineCompliance[] = [];

      // Get security baselines
      const baselines = await this.graphClient
        .api('/deviceManagement/templates')
        .filter("isof('microsoft.graph.securityBaselineTemplate')")
        .get();

      // Get devices
      const devices = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select('id,deviceName,userPrincipalName')
        .get();

      // For each device and baseline, get compliance status
      for (const device of devices.value || []) {
        for (const baseline of baselines.value || []) {
          try {
            // Get device's baseline compliance state
            const complianceState = await this.graphClient
              .api(`/deviceManagement/managedDevices/${device.id}/deviceConfigurationStates`)
              .filter(`templateId eq '${baseline.id}'`)
              .get()
              .catch(() => ({ value: [] }));

            if (complianceState.value && complianceState.value.length > 0) {
              const state = complianceState.value[0];

              const baselineCompliance: SecurityBaselineCompliance = {
                deviceId: device.id,
                deviceName: device.deviceName || 'Unknown',
                userPrincipalName: device.userPrincipalName || 'Unknown',
                baselineName: baseline.displayName || 'Unknown Baseline',
                baselineVersion: baseline.version || '1.0',
                state: state.state || 'unknown',
                settingsCount: state.settingCount || 0,
                compliantSettingsCount: state.compliantSettingCount || 0,
                noncompliantSettingsCount: state.nonCompliantSettingCount || 0,
                conflictSettingsCount: state.conflictSettingCount || 0,
                errorSettingsCount: state.errorSettingCount || 0,
                lastReportedDateTime: new Date(state.lastReportedDateTime || Date.now()),
              };

              // Get noncompliant settings details if available
              if (state.state === 'noncompliant') {
                baselineCompliance.noncompliantSettings = await this.getNoncompliantSettings(
                  device.id,
                  baseline.id
                );
              }

              if (this.matchesFilters(baselineCompliance, filters)) {
                complianceData.push(baselineCompliance);
              }
            }
          } catch (error) {
            console.error(`Error processing baseline ${baseline.id} for device ${device.id}:`, error);
          }
        }
      }

      return complianceData;
    } catch (error) {
      console.error('Error fetching security baseline compliance:', error);
      throw new Error(`Failed to fetch security baseline compliance: ${error.message}`);
    }
  }

  /**
   * Get noncompliant security baseline settings for a device
   * @param deviceId - The device ID
   * @param baselineId - The baseline template ID
   * @returns Array of noncompliant settings
   */
  private async getNoncompliantSettings(
    deviceId: string,
    baselineId: string
  ): Promise<SecurityBaselineSetting[]> {
    try {
      const settings = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/deviceConfigurationStates`)
        .filter(`templateId eq '${baselineId}'`)
        .expand('settingStates')
        .get();

      const noncompliantSettings: SecurityBaselineSetting[] = [];

      if (settings.value && settings.value[0]?.settingStates) {
        for (const setting of settings.value[0].settingStates) {
          if (setting.state === 'noncompliant' || setting.state === 'error') {
            noncompliantSettings.push({
              settingName: setting.setting || 'Unknown',
              settingCategory: setting.settingCategory || 'General',
              currentValue: setting.currentValue || 'Not set',
              expectedValue: setting.expectedValue || 'Not specified',
              state: setting.state,
              remediationAction: this.getRemediationAction(setting),
            });
          }
        }
      }

      return noncompliantSettings;
    } catch (error) {
      console.error('Error fetching noncompliant settings:', error);
      return [];
    }
  }

  // ==========================================================================
  // Microsoft Defender ATP Status
  // ==========================================================================

  /**
   * Get Microsoft Defender ATP status for all devices
   * @param filters - Optional filters to apply
   * @returns Array of Defender ATP status information
   */
  async getDefenderATPStatus(
    filters?: SecurityComplianceFilters
  ): Promise<DefenderATPStatus[]> {
    try {
      const defenderStatuses: DefenderATPStatus[] = [];

      // Get managed devices
      const devices = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select('id,deviceName,userPrincipalName,operatingSystem')
        .filter("operatingSystem eq 'Windows' or operatingSystem eq 'macOS'")
        .get();

      for (const device of devices.value || []) {
        try {
          // Get Windows protection state (includes Defender info)
          const protectionState = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/windowsProtectionState`)
            .get()
            .catch(() => null);

          if (protectionState) {
            const defenderStatus: DefenderATPStatus = {
              deviceId: device.id,
              deviceName: device.deviceName || 'Unknown',
              userPrincipalName: device.userPrincipalName || 'Unknown',
              onboardingState: this.mapOnboardingState(protectionState.malwareProtectionEnabled),
              healthState: this.mapHealthState(protectionState),
              riskScore: this.calculateRiskScore(protectionState),
              exposureScore: this.calculateExposureScore(protectionState),
              antivirusStatus: {
                enabled: protectionState.malwareProtectionEnabled || false,
                updated: !protectionState.antivirusSignatureOutOfDate,
                productVersion: protectionState.productStatus || 'Unknown',
                engineVersion: protectionState.engineVersion || 'Unknown',
                signatureVersion: protectionState.antivirusSignatureVersion || 'Unknown',
                lastSignatureUpdateDateTime: protectionState.lastSignatureUpdateDateTime
                  ? new Date(protectionState.lastSignatureUpdateDateTime)
                  : undefined,
              },
              firewallStatus: {
                domainProfile: protectionState.firewallDomainProfile || 'unknown',
                privateProfile: protectionState.firewallPrivateProfile || 'unknown',
                publicProfile: protectionState.firewallPublicProfile || 'unknown',
              },
              realTimeProtectionEnabled: protectionState.realTimeProtectionEnabled || false,
              networkInspectionEnabled: protectionState.networkInspectionSystemEnabled || false,
              quickScanOverdue: protectionState.quickScanOverdue || false,
              fullScanOverdue: protectionState.fullScanOverdue || false,
              signatureOutOfDate: protectionState.antivirusSignatureOutOfDate || false,
              lastQuickScanDateTime: protectionState.lastQuickScanDateTime
                ? new Date(protectionState.lastQuickScanDateTime)
                : undefined,
              lastFullScanDateTime: protectionState.lastFullScanDateTime
                ? new Date(protectionState.lastFullScanDateTime)
                : undefined,
              lastReportedDateTime: new Date(protectionState.lastReportedDateTime || Date.now()),
            };

            if (this.matchesFilters(defenderStatus, filters)) {
              defenderStatuses.push(defenderStatus);
            }
          }
        } catch (error) {
          console.error(`Error processing Defender status for device ${device.id}:`, error);
        }
      }

      return defenderStatuses;
    } catch (error) {
      console.error('Error fetching Defender ATP status:', error);
      throw new Error(`Failed to fetch Defender ATP status: ${error.message}`);
    }
  }

  /**
   * Get Defender ATP status for a specific device
   * @param deviceId - The device ID
   * @returns Defender ATP status information
   */
  async getDeviceDefenderATPStatus(deviceId: string): Promise<DefenderATPStatus> {
    try {
      const device = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}`)
        .select('id,deviceName,userPrincipalName')
        .get();

      const protectionState = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/windowsProtectionState`)
        .get();

      return {
        deviceId: device.id,
        deviceName: device.deviceName || 'Unknown',
        userPrincipalName: device.userPrincipalName || 'Unknown',
        onboardingState: this.mapOnboardingState(protectionState.malwareProtectionEnabled),
        healthState: this.mapHealthState(protectionState),
        riskScore: this.calculateRiskScore(protectionState),
        exposureScore: this.calculateExposureScore(protectionState),
        antivirusStatus: {
          enabled: protectionState.malwareProtectionEnabled || false,
          updated: !protectionState.antivirusSignatureOutOfDate,
          productVersion: protectionState.productStatus || 'Unknown',
          engineVersion: protectionState.engineVersion || 'Unknown',
          signatureVersion: protectionState.antivirusSignatureVersion || 'Unknown',
          lastSignatureUpdateDateTime: protectionState.lastSignatureUpdateDateTime
            ? new Date(protectionState.lastSignatureUpdateDateTime)
            : undefined,
        },
        firewallStatus: {
          domainProfile: protectionState.firewallDomainProfile || 'unknown',
          privateProfile: protectionState.firewallPrivateProfile || 'unknown',
          publicProfile: protectionState.firewallPublicProfile || 'unknown',
        },
        realTimeProtectionEnabled: protectionState.realTimeProtectionEnabled || false,
        networkInspectionEnabled: protectionState.networkInspectionSystemEnabled || false,
        quickScanOverdue: protectionState.quickScanOverdue || false,
        fullScanOverdue: protectionState.fullScanOverdue || false,
        signatureOutOfDate: protectionState.antivirusSignatureOutOfDate || false,
        lastQuickScanDateTime: protectionState.lastQuickScanDateTime
          ? new Date(protectionState.lastQuickScanDateTime)
          : undefined,
        lastFullScanDateTime: protectionState.lastFullScanDateTime
          ? new Date(protectionState.lastFullScanDateTime)
          : undefined,
        lastReportedDateTime: new Date(protectionState.lastReportedDateTime || Date.now()),
      };
    } catch (error) {
      console.error(`Error fetching Defender ATP status for device ${deviceId}:`, error);
      throw new Error(`Failed to fetch Defender ATP status: ${error.message}`);
    }
  }

  // ==========================================================================
  // Encryption Status (BitLocker & FileVault)
  // ==========================================================================

  /**
   * Get encryption status for all devices
   * @param filters - Optional filters to apply
   * @returns Array of encryption status information
   */
  async getEncryptionStatus(
    filters?: SecurityComplianceFilters
  ): Promise<EncryptionStatus[]> {
    try {
      const encryptionStatuses: EncryptionStatus[] = [];

      // Get managed devices
      const devices = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select('id,deviceName,userPrincipalName,operatingSystem,isEncrypted')
        .get();

      for (const device of devices.value || []) {
        try {
          const platform = this.normalizePlatform(device.operatingSystem);
          let encryptionStatus: EncryptionStatus;

          if (platform === 'Windows') {
            // Get BitLocker status
            const bitLockerState = await this.getBitLockerStatus(device.id);
            encryptionStatus = {
              deviceId: device.id,
              deviceName: device.deviceName || 'Unknown',
              userPrincipalName: device.userPrincipalName || 'Unknown',
              platform: 'Windows',
              encryptionMethod: bitLockerState?.encryptionMethod || 'Unknown',
              encryptionState: this.mapEncryptionState(device.isEncrypted, bitLockerState),
              encryptionCompliance: this.evaluateEncryptionCompliance(bitLockerState),
              bitLockerStatus: bitLockerState,
              lastReportedDateTime: new Date(),
            };
          } else if (platform === 'macOS') {
            // Get FileVault status
            const fileVaultState = await this.getFileVaultStatus(device.id);
            encryptionStatus = {
              deviceId: device.id,
              deviceName: device.deviceName || 'Unknown',
              userPrincipalName: device.userPrincipalName || 'Unknown',
              platform: 'macOS',
              encryptionMethod: 'FileVault',
              encryptionState: fileVaultState?.enabled ? 'encrypted' : 'notEncrypted',
              encryptionCompliance: fileVaultState?.enabled ? 'compliant' : 'noncompliant',
              fileVaultStatus: fileVaultState,
              lastReportedDateTime: new Date(),
            };
          } else {
            // Other platforms
            encryptionStatus = {
              deviceId: device.id,
              deviceName: device.deviceName || 'Unknown',
              userPrincipalName: device.userPrincipalName || 'Unknown',
              platform: platform as any,
              encryptionMethod: 'Platform Default',
              encryptionState: device.isEncrypted ? 'encrypted' : 'notEncrypted',
              encryptionCompliance: device.isEncrypted ? 'compliant' : 'noncompliant',
              lastReportedDateTime: new Date(),
            };
          }

          if (this.matchesFilters(encryptionStatus, filters)) {
            encryptionStatuses.push(encryptionStatus);
          }
        } catch (error) {
          console.error(`Error processing encryption status for device ${device.id}:`, error);
        }
      }

      return encryptionStatuses;
    } catch (error) {
      console.error('Error fetching encryption status:', error);
      throw new Error(`Failed to fetch encryption status: ${error.message}`);
    }
  }

  /**
   * Get BitLocker status for a Windows device
   * @param deviceId - The device ID
   * @returns BitLocker status information
   */
  private async getBitLockerStatus(deviceId: string): Promise<BitLockerStatus | undefined> {
    try {
      const encryptionState = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/windowsProtectionState`)
        .get()
        .catch(() => null);

      if (!encryptionState) return undefined;

      // Get BitLocker recovery keys
      const recoveryKeys = await this.graphClient
        .api(`/informationProtection/bitlocker/recoveryKeys`)
        .filter(`deviceId eq '${deviceId}'`)
        .get()
        .catch(() => ({ value: [] }));

      return {
        protectionStatus: encryptionState.isEncrypted ? 'protected' : 'unprotected',
        encryptionMethod: this.mapBitLockerEncryptionMethod(encryptionState.encryptionMethod),
        recoveryKeyBackupStatus: recoveryKeys.value?.length > 0 ? 'backedUp' : 'notBackedUp',
        tpmPresent: true, // Assume TPM is present for modern Windows devices
        tpmActivated: true,
        volumes: [
          {
            volumeType: 'operatingSystem',
            mountPoint: 'C:',
            encryptionPercentage: encryptionState.isEncrypted ? 100 : 0,
            protectionStatus: encryptionState.isEncrypted ? 'protected' : 'unprotected',
          },
        ],
      };
    } catch (error) {
      console.error('Error fetching BitLocker status:', error);
      return undefined;
    }
  }

  /**
   * Get FileVault status for a macOS device
   * @param deviceId - The device ID
   * @returns FileVault status information
   */
  private async getFileVaultStatus(deviceId: string): Promise<FileVaultStatus | undefined> {
    try {
      // Get device configuration states
      const configStates = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/deviceConfigurationStates`)
        .get()
        .catch(() => ({ value: [] }));

      // Look for FileVault configuration
      const fileVaultConfig = configStates.value?.find(
        (config: any) => config.displayName?.toLowerCase().includes('filevault')
      );

      if (fileVaultConfig) {
        return {
          enabled: fileVaultConfig.state === 'compliant',
          recoveryKeyType: 'institutional',
          recoveryKeyEscrowed: fileVaultConfig.state === 'compliant',
          encryptionInProgress: fileVaultConfig.state === 'pending',
          deferralCount: 0,
        };
      }

      return {
        enabled: false,
        recoveryKeyType: 'none',
        recoveryKeyEscrowed: false,
        encryptionInProgress: false,
      };
    } catch (error) {
      console.error('Error fetching FileVault status:', error);
      return undefined;
    }
  }

  // ==========================================================================
  // Threat Detection Reports
  // ==========================================================================

  /**
   * Get threat detection reports for all devices
   * @param filters - Optional filters to apply
   * @returns Array of threat detection reports
   */
  async getThreatDetectionReports(
    filters?: SecurityComplianceFilters
  ): Promise<ThreatDetectionReport[]> {
    try {
      const threatReports: ThreatDetectionReport[] = [];

      // Get devices with detected threats
      const devices = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select('id,deviceName,userPrincipalName')
        .get();

      for (const device of devices.value || []) {
        try {
          // Get detected malware for the device
          const detectedMalware = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/detectedApps`)
            .filter("detectionState ne 'notDetected'")
            .get()
            .catch(() => ({ value: [] }));

          // Get Windows protection state for additional threat info
          const protectionState = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/windowsProtectionState`)
            .get()
            .catch(() => null);

          const threats: ThreatDetails[] = [];
          let activeThreatCount = 0;
          let remediatedThreatCount = 0;

          for (const malware of detectedMalware.value || []) {
            const threat: ThreatDetails = {
              threatId: malware.id || `threat-${Date.now()}`,
              threatName: malware.displayName || 'Unknown Threat',
              category: this.categorizeThreat(malware.displayName),
              severity: this.assessThreatSeverity(malware),
              state: this.mapThreatState(malware.detectionState),
              detectionDateTime: new Date(malware.detectionDateTime || Date.now()),
              remediationDateTime: malware.remediationDateTime
                ? new Date(malware.remediationDateTime)
                : undefined,
              affectedFiles: malware.affectedFiles || [],
              remediationAction: this.determineThreatRemediationAction(malware),
              executionState: malware.isActive ? 'running' : 'notRunning',
            };

            threats.push(threat);

            if (threat.state === 'active') {
              activeThreatCount++;
            } else if (threat.state === 'remediated' || threat.state === 'removed') {
              remediatedThreatCount++;
            }
          }

          const threatReport: ThreatDetectionReport = {
            deviceId: device.id,
            deviceName: device.deviceName || 'Unknown',
            userPrincipalName: device.userPrincipalName || 'Unknown',
            threatCount: threats.length,
            activeThreatCount,
            remediatedThreatCount,
            threats,
            lastScanDateTime: protectionState?.lastQuickScanDateTime
              ? new Date(protectionState.lastQuickScanDateTime)
              : undefined,
            lastReportedDateTime: new Date(),
          };

          // Only include devices with threats or include all based on filters
          if (threats.length > 0 || !filters) {
            if (this.matchesFilters(threatReport, filters)) {
              threatReports.push(threatReport);
            }
          }
        } catch (error) {
          console.error(`Error processing threats for device ${device.id}:`, error);
        }
      }

      return threatReports;
    } catch (error) {
      console.error('Error fetching threat detection reports:', error);
      throw new Error(`Failed to fetch threat detection reports: ${error.message}`);
    }
  }

  /**
   * Get threat detection report for a specific device
   * @param deviceId - The device ID
   * @returns Threat detection report
   */
  async getDeviceThreatReport(deviceId: string): Promise<ThreatDetectionReport> {
    try {
      const device = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}`)
        .select('id,deviceName,userPrincipalName')
        .get();

      const detectedMalware = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/detectedApps`)
        .filter("detectionState ne 'notDetected'")
        .get()
        .catch(() => ({ value: [] }));

      const protectionState = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/windowsProtectionState`)
        .get()
        .catch(() => null);

      const threats: ThreatDetails[] = [];
      let activeThreatCount = 0;
      let remediatedThreatCount = 0;

      for (const malware of detectedMalware.value || []) {
        const threat: ThreatDetails = {
          threatId: malware.id || `threat-${Date.now()}`,
          threatName: malware.displayName || 'Unknown Threat',
          category: this.categorizeThreat(malware.displayName),
          severity: this.assessThreatSeverity(malware),
          state: this.mapThreatState(malware.detectionState),
          detectionDateTime: new Date(malware.detectionDateTime || Date.now()),
          remediationDateTime: malware.remediationDateTime
            ? new Date(malware.remediationDateTime)
            : undefined,
          affectedFiles: malware.affectedFiles || [],
          remediationAction: this.determineThreatRemediationAction(malware),
          executionState: malware.isActive ? 'running' : 'notRunning',
        };

        threats.push(threat);

        if (threat.state === 'active') {
          activeThreatCount++;
        } else if (threat.state === 'remediated' || threat.state === 'removed') {
          remediatedThreatCount++;
        }
      }

      return {
        deviceId: device.id,
        deviceName: device.deviceName || 'Unknown',
        userPrincipalName: device.userPrincipalName || 'Unknown',
        threatCount: threats.length,
        activeThreatCount,
        remediatedThreatCount,
        threats,
        lastScanDateTime: protectionState?.lastQuickScanDateTime
          ? new Date(protectionState.lastQuickScanDateTime)
          : undefined,
        lastReportedDateTime: new Date(),
      };
    } catch (error) {
      console.error(`Error fetching threat report for device ${deviceId}:`, error);
      throw new Error(`Failed to fetch threat report: ${error.message}`);
    }
  }

  // ==========================================================================
  // Comprehensive Security Summary
  // ==========================================================================

  /**
   * Generate a comprehensive security compliance summary
   * @param filters - Optional filters to apply
   * @returns Security compliance summary
   */
  async getSecurityComplianceSummary(
    filters?: SecurityComplianceFilters
  ): Promise<SecurityComplianceSummary> {
    try {
      // Fetch all compliance data in parallel
      const [
        updateCompliance,
        baselineCompliance,
        defenderStatus,
        encryptionStatus,
        threatReports,
      ] = await Promise.all([
        this.getWindowsUpdateCompliance(filters),
        this.getSecurityBaselineCompliance(filters),
        this.getDefenderATPStatus(filters),
        this.getEncryptionStatus(filters),
        this.getThreatDetectionReports(filters),
      ]);

      // Calculate update compliance summary
      const updateSummary = {
        compliant: updateCompliance.filter(d => d.complianceState === 'compliant').length,
        noncompliant: updateCompliance.filter(d => d.complianceState === 'noncompliant').length,
        unknown: updateCompliance.filter(d => d.complianceState === 'unknown').length,
        devicesWithPendingUpdates: updateCompliance.filter(d => d.pendingUpdates > 0).length,
        devicesWithFailedUpdates: updateCompliance.filter(d => d.failedUpdates > 0).length,
      };

      // Calculate baseline compliance summary
      const baselineMap = new Map<string, { compliant: number; noncompliant: number }>();
      for (const baseline of baselineCompliance) {
        const existing = baselineMap.get(baseline.baselineName) || { compliant: 0, noncompliant: 0 };
        if (baseline.state === 'compliant') {
          existing.compliant++;
        } else {
          existing.noncompliant++;
        }
        baselineMap.set(baseline.baselineName, existing);
      }

      const baselineSummary = {
        compliant: baselineCompliance.filter(d => d.state === 'compliant').length,
        noncompliant: baselineCompliance.filter(d => d.state === 'noncompliant').length,
        unknown: baselineCompliance.filter(d => d.state === 'unknown').length,
        baselines: baselineMap,
      };

      // Calculate Defender ATP summary
      const defenderSummary = {
        onboarded: defenderStatus.filter(d => d.onboardingState === 'onboarded').length,
        notOnboarded: defenderStatus.filter(d => d.onboardingState !== 'onboarded').length,
        healthy: defenderStatus.filter(d => d.healthState === 'healthy').length,
        unhealthy: defenderStatus.filter(d => d.healthState !== 'healthy').length,
        highRiskDevices: defenderStatus.filter(d => d.riskScore === 'high').length,
        antivirusDisabled: defenderStatus.filter(d => !d.antivirusStatus.enabled).length,
        signatureOutOfDate: defenderStatus.filter(d => d.signatureOutOfDate).length,
      };

      // Calculate encryption summary
      const encryptionSummary = {
        encrypted: encryptionStatus.filter(d => d.encryptionState === 'encrypted').length,
        notEncrypted: encryptionStatus.filter(d => d.encryptionState === 'notEncrypted').length,
        encrypting: encryptionStatus.filter(d => d.encryptionState === 'encrypting').length,
        unknown: encryptionStatus.filter(d => d.encryptionState === 'unknown').length,
        bitLockerCompliant: encryptionStatus.filter(
          d => d.platform === 'Windows' && d.encryptionCompliance === 'compliant'
        ).length,
        fileVaultCompliant: encryptionStatus.filter(
          d => d.platform === 'macOS' && d.encryptionCompliance === 'compliant'
        ).length,
      };

      // Calculate threat summary
      const totalThreats = threatReports.reduce((sum, report) => sum + report.threatCount, 0);
      const activeThreats = threatReports.reduce((sum, report) => sum + report.activeThreatCount, 0);
      const remediatedThreats = threatReports.reduce(
        (sum, report) => sum + report.remediatedThreatCount,
        0
      );
      const criticalThreats = threatReports.reduce(
        (sum, report) =>
          sum + report.threats.filter(t => t.severity === 'critical').length,
        0
      );
      const highSeverityThreats = threatReports.reduce(
        (sum, report) =>
          sum + report.threats.filter(t => t.severity === 'high').length,
        0
      );

      const threatSummary = {
        devicesWithThreats: threatReports.filter(r => r.threatCount > 0).length,
        totalThreats,
        activeThreats,
        remediatedThreats,
        criticalThreats,
        highSeverityThreats,
      };

      // Get total unique devices across all reports
      const allDeviceIds = new Set<string>();
      updateCompliance.forEach(d => allDeviceIds.add(d.deviceId));
      baselineCompliance.forEach(d => allDeviceIds.add(d.deviceId));
      defenderStatus.forEach(d => allDeviceIds.add(d.deviceId));
      encryptionStatus.forEach(d => allDeviceIds.add(d.deviceId));
      threatReports.forEach(d => allDeviceIds.add(d.deviceId));

      return {
        totalDevices: allDeviceIds.size,
        reportGeneratedDate: new Date(),
        updateCompliance: updateSummary,
        baselineCompliance: baselineSummary,
        defenderATP: defenderSummary,
        encryption: encryptionSummary,
        threats: threatSummary,
      };
    } catch (error) {
      console.error('Error generating security compliance summary:', error);
      throw new Error(`Failed to generate security compliance summary: ${error.message}`);
    }
  }

  // ==========================================================================
  // Export Functions
  // ==========================================================================

  /**
   * Export security compliance data to various formats
   * @param data - The data to export
   * @param format - The export format
   * @param filename - The output filename
   * @returns Exported data as string or buffer
   */
  async exportSecurityCompliance(
    data: any,
    format: ExportFormat,
    filename?: string
  ): Promise<string | Buffer> {
    switch (format) {
      case 'json':
        return this.exportToJSON(data);
      case 'csv':
        return this.exportToCSV(data);
      case 'html':
        return this.exportToHTML(data);
      case 'pdf':
        return this.exportToPDF(data);
      default:
        throw new Error(`Unsupported export format: ${format}`);
    }
  }

  /**
   * Export data to JSON format
   */
  private exportToJSON(data: any): string {
    return JSON.stringify(data, null, 2);
  }

  /**
   * Export data to CSV format
   */
  private exportToCSV(data: any): string {
    if (Array.isArray(data) && data.length > 0) {
      const headers = Object.keys(data[0]);
      const rows = data.map(item =>
        headers.map(header => {
          const value = item[header];
          if (value === null || value === undefined) return '';
          if (typeof value === 'object') return JSON.stringify(value);
          return `"${String(value).replace(/"/g, '""')}"`;
        }).join(',')
      );
      return [headers.join(','), ...rows].join('\n');
    }
    return '';
  }

  /**
   * Export data to HTML format
   */
  private exportToHTML(data: any): string {
    const timestamp = new Date().toLocaleString();
    let html = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Security Compliance Report</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1 { color: #0078d4; }
    table { border-collapse: collapse; width: 100%; margin: 20px 0; }
    th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
    th { background-color: #0078d4; color: white; }
    tr:nth-child(even) { background-color: #f2f2f2; }
    .summary { background-color: #e7f3ff; padding: 15px; border-radius: 5px; margin: 20px 0; }
    .compliant { color: #107c10; font-weight: bold; }
    .noncompliant { color: #d13438; font-weight: bold; }
    .warning { color: #ff8c00; font-weight: bold; }
  </style>
</head>
<body>
  <h1>Security and Compliance Report</h1>
  <p>Generated: ${timestamp}</p>
  <div class="summary">
    <h2>Report Data</h2>
    <pre>${JSON.stringify(data, null, 2)}</pre>
  </div>
</body>
</html>`;
    return html;
  }

  /**
   * Export data to PDF format (placeholder - requires PDF library)
   */
  private exportToPDF(data: any): Buffer {
    // This is a placeholder. In production, you would use a library like pdfkit or puppeteer
    throw new Error('PDF export requires additional dependencies (pdfkit or puppeteer)');
  }

  // ==========================================================================
  // Helper Methods
  // ==========================================================================

  /**
   * Parse update status from protection state
   */
  private parseUpdateStatus(status: any): UpdateStatus {
    if (!status) {
      return { state: 'unknown' };
    }

    return {
      state: status.state || 'unknown',
      availableVersion: status.availableVersion,
      installedVersion: status.installedVersion,
      errorCode: status.errorCode,
      errorDescription: status.errorDescription,
    };
  }

  /**
   * Calculate days until update deadline
   */
  private calculateDaysUntilDeadline(protectionState: any): number {
    if (!protectionState?.deadlineDate) return -1;
    const deadline = new Date(protectionState.deadlineDate);
    const today = new Date();
    const diffTime = deadline.getTime() - today.getTime();
    const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    return diffDays;
  }

  /**
   * Get remediation action for a noncompliant setting
   */
  private getRemediationAction(setting: any): string {
    return `Configure ${setting.setting} to ${setting.expectedValue}`;
  }

  /**
   * Map onboarding state
   */
  private mapOnboardingState(enabled: boolean): DefenderATPStatus['onboardingState'] {
    return enabled ? 'onboarded' : 'canBeOnboarded';
  }

  /**
   * Map health state from protection state
   */
  private mapHealthState(protectionState: any): DefenderATPStatus['healthState'] {
    if (!protectionState) return 'unknown';

    const issues = [
      !protectionState.malwareProtectionEnabled,
      protectionState.antivirusSignatureOutOfDate,
      !protectionState.realTimeProtectionEnabled,
      protectionState.quickScanOverdue,
    ];

    const issueCount = issues.filter(Boolean).length;

    if (issueCount === 0) return 'healthy';
    if (issueCount >= 2) return 'misconfigured';
    return 'inactive';
  }

  /**
   * Calculate risk score based on protection state
   */
  private calculateRiskScore(protectionState: any): DefenderATPStatus['riskScore'] {
    if (!protectionState) return 'unknown';

    let score = 0;
    if (!protectionState.malwareProtectionEnabled) score += 3;
    if (protectionState.antivirusSignatureOutOfDate) score += 2;
    if (!protectionState.realTimeProtectionEnabled) score += 2;
    if (protectionState.fullScanOverdue) score += 1;
    if (protectionState.quickScanOverdue) score += 1;

    if (score === 0) return 'none';
    if (score <= 2) return 'low';
    if (score <= 4) return 'medium';
    return 'high';
  }

  /**
   * Calculate exposure score
   */
  private calculateExposureScore(protectionState: any): number {
    if (!protectionState) return 0;

    let score = 0;
    if (!protectionState.malwareProtectionEnabled) score += 30;
    if (protectionState.antivirusSignatureOutOfDate) score += 20;
    if (!protectionState.realTimeProtectionEnabled) score += 25;
    if (!protectionState.networkInspectionSystemEnabled) score += 15;
    if (protectionState.quickScanOverdue) score += 5;
    if (protectionState.fullScanOverdue) score += 5;

    return Math.min(score, 100);
  }

  /**
   * Normalize platform name
   */
  private normalizePlatform(platform: string): string {
    if (!platform) return 'Unknown';
    const lower = platform.toLowerCase();
    if (lower.includes('windows')) return 'Windows';
    if (lower.includes('mac') || lower.includes('osx')) return 'macOS';
    if (lower.includes('ios')) return 'iOS';
    if (lower.includes('android')) return 'Android';
    if (lower.includes('linux')) return 'Linux';
    return platform;
  }

  /**
   * Map encryption state
   */
  private mapEncryptionState(
    isEncrypted: boolean,
    encryptionDetails: any
  ): EncryptionStatus['encryptionState'] {
    if (!encryptionDetails) {
      return isEncrypted ? 'encrypted' : 'notEncrypted';
    }

    if (encryptionDetails.encryptionInProgress) return 'encrypting';
    if (encryptionDetails.protectionStatus === 'protected') return 'encrypted';
    if (encryptionDetails.protectionStatus === 'unprotected') return 'notEncrypted';

    return isEncrypted ? 'encrypted' : 'notEncrypted';
  }

  /**
   * Evaluate encryption compliance
   */
  private evaluateEncryptionCompliance(
    encryptionDetails: any
  ): EncryptionStatus['encryptionCompliance'] {
    if (!encryptionDetails) return 'unknown';

    if (encryptionDetails.protectionStatus === 'protected' &&
        encryptionDetails.recoveryKeyBackupStatus === 'backedUp') {
      return 'compliant';
    }

    return 'noncompliant';
  }

  /**
   * Map BitLocker encryption method
   */
  private mapBitLockerEncryptionMethod(method: string): BitLockerStatus['encryptionMethod'] {
    if (!method) return 'unknown';
    const lower = method.toLowerCase();
    if (lower.includes('aes256') || lower.includes('xts-aes256')) return 'xtsAes256';
    if (lower.includes('aes128') || lower.includes('xts-aes128')) return 'xtsAes128';
    return 'unknown';
  }

  /**
   * Categorize threat type
   */
  private categorizeThreat(threatName: string): ThreatDetails['category'] {
    if (!threatName) return 'other';
    const lower = threatName.toLowerCase();
    if (lower.includes('ransomware')) return 'ransomware';
    if (lower.includes('trojan')) return 'trojan';
    if (lower.includes('virus')) return 'virus';
    if (lower.includes('spyware')) return 'spyware';
    if (lower.includes('adware')) return 'adware';
    if (lower.includes('pua') || lower.includes('potentially unwanted')) return 'pua';
    return 'malware';
  }

  /**
   * Assess threat severity
   */
  private assessThreatSeverity(malware: any): ThreatDetails['severity'] {
    const severity = malware.severity?.toLowerCase() || '';
    if (severity.includes('critical')) return 'critical';
    if (severity.includes('high')) return 'high';
    if (severity.includes('medium')) return 'medium';
    return 'low';
  }

  /**
   * Map threat state
   */
  private mapThreatState(detectionState: string): ThreatDetails['state'] {
    if (!detectionState) return 'active';
    const lower = detectionState.toLowerCase();
    if (lower.includes('remediated')) return 'remediated';
    if (lower.includes('quarantined')) return 'quarantined';
    if (lower.includes('removed')) return 'removed';
    if (lower.includes('allowed')) return 'allowed';
    if (lower.includes('blocked')) return 'blocked';
    return 'active';
  }

  /**
   * Determine threat remediation action
   */
  private determineThreatRemediationAction(malware: any): ThreatDetails['remediationAction'] {
    const category = this.categorizeThreat(malware.displayName);
    const severity = this.assessThreatSeverity(malware);

    if (severity === 'critical' || category === 'ransomware') return 'remove';
    if (severity === 'high') return 'quarantine';
    if (category === 'pua') return 'allow';
    return 'clean';
  }

  /**
   * Check if data matches filters
   */
  private matchesFilters(data: any, filters?: SecurityComplianceFilters): boolean {
    if (!filters) return true;

    // Device ID filter
    if (filters.deviceIds && filters.deviceIds.length > 0) {
      if (!filters.deviceIds.includes(data.deviceId)) return false;
    }

    // User principal name filter
    if (filters.userPrincipalNames && filters.userPrincipalNames.length > 0) {
      if (!filters.userPrincipalNames.includes(data.userPrincipalName)) return false;
    }

    // Compliance state filter
    if (filters.complianceState && filters.complianceState.length > 0) {
      const state = data.complianceState || data.state || data.encryptionCompliance;
      if (!filters.complianceState.includes(state)) return false;
    }

    // Platform filter
    if (filters.platforms && filters.platforms.length > 0) {
      if (data.platform && !filters.platforms.includes(data.platform)) return false;
    }

    // Risk level filter
    if (filters.riskLevel && filters.riskLevel.length > 0) {
      if (data.riskScore && !filters.riskLevel.includes(data.riskScore)) return false;
    }

    // Date range filter
    if (filters.dateFrom || filters.dateTo) {
      const reportDate = data.lastReportedDateTime || data.lastSyncDateTime || new Date();
      if (filters.dateFrom && reportDate < filters.dateFrom) return false;
      if (filters.dateTo && reportDate > filters.dateTo) return false;
    }

    return true;
  }
}

// ============================================================================
// Usage Examples
// ============================================================================

/**
 * Example usage of the SecurityComplianceReporter
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { SecurityComplianceReporter } from './security-compliance';
 *
 * // Initialize Graph client (with authentication)
 * const graphClient = Client.init({
 *   authProvider: (done) => {
 *     done(null, accessToken);
 *   }
 * });
 *
 * // Create reporter instance
 * const reporter = new SecurityComplianceReporter(graphClient);
 *
 * // Get Windows Update compliance
 * const updateCompliance = await reporter.getWindowsUpdateCompliance();
 * console.log(`Devices with pending updates: ${updateCompliance.filter(d => d.pendingUpdates > 0).length}`);
 *
 * // Get security baseline compliance
 * const baselineCompliance = await reporter.getSecurityBaselineCompliance();
 * console.log(`Noncompliant devices: ${baselineCompliance.filter(d => d.state === 'noncompliant').length}`);
 *
 * // Get Defender ATP status
 * const defenderStatus = await reporter.getDefenderATPStatus();
 * console.log(`High-risk devices: ${defenderStatus.filter(d => d.riskScore === 'high').length}`);
 *
 * // Get encryption status
 * const encryptionStatus = await reporter.getEncryptionStatus();
 * console.log(`Encrypted devices: ${encryptionStatus.filter(d => d.encryptionState === 'encrypted').length}`);
 *
 * // Get threat detection reports
 * const threatReports = await reporter.getThreatDetectionReports();
 * console.log(`Devices with active threats: ${threatReports.filter(r => r.activeThreatCount > 0).length}`);
 *
 * // Get comprehensive summary
 * const summary = await reporter.getSecurityComplianceSummary();
 * console.log('Security Compliance Summary:', summary);
 *
 * // Export to JSON
 * const jsonData = await reporter.exportSecurityCompliance(summary, 'json');
 * console.log(jsonData);
 *
 * // Use filters
 * const filters: SecurityComplianceFilters = {
 *   complianceState: ['noncompliant'],
 *   platforms: ['Windows'],
 *   dateFrom: new Date('2024-01-01'),
 * };
 *
 * const filteredCompliance = await reporter.getWindowsUpdateCompliance(filters);
 * console.log(`Filtered results: ${filteredCompliance.length}`);
 * ```
 */

export default SecurityComplianceReporter;
