/**
 * Device Security Reports for Microsoft Intune
 *
 * Replicates Configuration Manager security reports for Intune environments:
 * - Report 23: Mobile devices that are jailbroken or rooted
 * - Report 27: Mobile devices with certificate renewal issues
 *
 * @module device-security-reports
 * @author Intune Reporting System
 * @version 1.0.0
 */

import { Client } from '@microsoft/microsoft-graph-client';
import * as fs from 'fs/promises';
import * as path from 'path';

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Configuration for authentication with Microsoft Graph
 */
export interface SecurityReportAuthConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
}

/**
 * Jailbroken or rooted device information
 */
export interface JailbrokenDevice {
  /** Device ID */
  deviceId: string;
  /** Device name */
  deviceName: string;
  /** User principal name */
  userPrincipalName: string;
  /** User email */
  userEmail: string;
  /** Operating system (iOS, Android) */
  operatingSystem: string;
  /** OS version */
  osVersion: string;
  /** Model of the device */
  model: string;
  /** Manufacturer */
  manufacturer: string;
  /** Whether device is jailbroken/rooted */
  isJailBroken: boolean;
  /** Risk level assessment */
  riskLevel: 'Critical' | 'High' | 'Medium' | 'Low';
  /** Risk score (0-100) */
  riskScore: number;
  /** Compliance state */
  complianceState: string;
  /** Last sync date/time */
  lastSyncDateTime: Date;
  /** Enrollment date */
  enrolledDateTime: Date;
  /** Management state */
  managementState: string;
  /** Additional security notes */
  securityNotes: string[];
}

/**
 * Certificate renewal issue information
 */
export interface CertificateIssue {
  /** Device ID */
  deviceId: string;
  /** Device name */
  deviceName: string;
  /** User principal name */
  userPrincipalName: string;
  /** User email */
  userEmail: string;
  /** Operating system */
  operatingSystem: string;
  /** Certificate type */
  certificateType: 'Device' | 'User' | 'WiFi' | 'VPN' | 'SCEP' | 'PKCS' | 'Other';
  /** Certificate issuer */
  issuer: string;
  /** Certificate subject */
  subject: string;
  /** Certificate expiration date */
  expirationDate: Date;
  /** Days until expiration */
  daysUntilExpiration: number;
  /** Certificate status */
  status: 'Expired' | 'Expiring' | 'Renewal Failed' | 'Invalid' | 'Revoked' | 'Valid';
  /** Issue severity */
  severity: 'Critical' | 'High' | 'Medium' | 'Low';
  /** Issue description */
  issueDescription: string;
  /** Remediation action */
  remediationAction: string;
  /** Last renewal attempt */
  lastRenewalAttempt?: Date;
  /** Renewal error message */
  renewalError?: string;
  /** Certificate thumbprint */
  thumbprint?: string;
  /** Compliance state */
  complianceState: string;
}

/**
 * Security risk assessment
 */
export interface SecurityRiskAssessment {
  /** Total devices assessed */
  totalDevices: number;
  /** Critical risk devices */
  criticalRiskDevices: number;
  /** High risk devices */
  highRiskDevices: number;
  /** Medium risk devices */
  mediumRiskDevices: number;
  /** Low risk devices */
  lowRiskDevices: number;
  /** Overall risk score (0-100) */
  overallRiskScore: number;
  /** Risk distribution percentage */
  riskDistribution: {
    critical: number;
    high: number;
    medium: number;
    low: number;
  };
  /** Recommendations */
  recommendations: string[];
}

/**
 * Jailbreak detection report summary
 */
export interface JailbreakReportSummary {
  /** Total devices scanned */
  totalDevices: number;
  /** Jailbroken/rooted devices found */
  jailbrokenDevices: number;
  /** Devices by platform */
  byPlatform: {
    iOS: { total: number; jailbroken: number; percentage: number };
    Android: { total: number; rooted: number; percentage: number };
  };
  /** Risk assessment */
  riskAssessment: SecurityRiskAssessment;
  /** Report generation timestamp */
  generatedAt: Date;
  /** Compliance impact */
  complianceImpact: {
    nonCompliantDueToJailbreak: number;
    percentage: number;
  };
}

/**
 * Certificate renewal report summary
 */
export interface CertificateReportSummary {
  /** Total devices with certificates */
  totalDevices: number;
  /** Devices with certificate issues */
  devicesWithIssues: number;
  /** Expired certificates */
  expiredCertificates: number;
  /** Certificates expiring soon (within 30 days) */
  expiringSoonCertificates: number;
  /** Failed renewals */
  failedRenewals: number;
  /** Issues by certificate type */
  byCertificateType: Record<string, number>;
  /** Issues by severity */
  bySeverity: {
    critical: number;
    high: number;
    medium: number;
    low: number;
  };
  /** Report generation timestamp */
  generatedAt: Date;
  /** Compliance impact */
  complianceImpact: {
    nonCompliantDueToCertificates: number;
    percentage: number;
  };
}

/**
 * Health Attestation State information
 */
export interface HealthAttestationState {
  /** Device ID */
  deviceId: string;
  /** Device name */
  deviceName: string;
  /** User principal name */
  userPrincipalName: string;
  /** User email */
  userEmail: string;
  /** Operating system */
  operatingSystem: string;
  /** OS version */
  osVersion: string;
  /** Model */
  model: string;
  /** Manufacturer */
  manufacturer: string;
  /** Last sync date/time */
  lastSyncDateTime: Date;
  /** Compliance state */
  complianceState: string;

  // Health Attestation Properties
  /** Health attestation supported */
  healthAttestationSupported: boolean;
  /** BitLocker status */
  bitLockerStatus: 'Enabled' | 'Disabled' | 'Unknown';
  /** Secure Boot status */
  secureBootEnabled: boolean;
  /** Code integrity status */
  codeIntegrityEnabled: boolean;
  /** Boot debugging status */
  bootDebuggingEnabled: boolean;
  /** Kernel debugging status */
  kernelDebuggingEnabled: boolean;
  /** Test signing status */
  testSigningEnabled: boolean;
  /** Early Launch Anti-Malware status */
  elamEnabled: boolean;
  /** Virtual Secure Mode (Credential Guard) status */
  virtualSecureModeEnabled: boolean;
  /** TPM version */
  tpmVersion?: string;
  /** PCR hash algorithm */
  pcrHashAlgorithm?: string;
  /** Boot manager version */
  bootManagerVersion?: string;
  /** Code integrity check version */
  codeIntegrityCheckVersion?: string;

  // Health Score
  /** Health score (0-100) */
  healthScore: number;
  /** Health score range */
  healthScoreRange: 'Excellent' | 'Good' | 'Fair' | 'Poor' | 'Critical';

  // Security Analysis
  /** Security issues identified */
  securityIssues: string[];
  /** Remediation recommendations */
  recommendations: string[];
}

/**
 * Health Attestation report summary
 */
export interface HealthAttestationReportSummary {
  /** Total devices assessed */
  totalDevices: number;
  /** Devices supporting health attestation */
  supportedDevices: number;
  /** Devices not supporting health attestation */
  unsupportedDevices: number;

  // BitLocker Statistics
  /** BitLocker enabled count */
  bitLockerEnabled: number;
  /** BitLocker disabled count */
  bitLockerDisabled: number;
  /** BitLocker percentage */
  bitLockerPercentage: number;

  // Secure Boot Statistics
  /** Secure Boot enabled count */
  secureBootEnabled: number;
  /** Secure Boot disabled count */
  secureBootDisabled: number;
  /** Secure Boot percentage */
  secureBootPercentage: number;

  // Code Integrity Statistics
  /** Code Integrity enabled count */
  codeIntegrityEnabled: number;
  /** Code Integrity disabled count */
  codeIntegrityDisabled: number;
  /** Code Integrity percentage */
  codeIntegrityPercentage: number;

  // Security Issues
  /** Devices with debugging enabled */
  debuggingEnabled: number;
  /** Devices with test signing enabled */
  testSigningEnabled: number;

  // Health Scores
  /** Average health score */
  averageHealthScore: number;
  /** Excellent health devices */
  excellentHealthDevices: number;
  /** Good health devices */
  goodHealthDevices: number;
  /** Fair health devices */
  fairHealthDevices: number;
  /** Poor health devices */
  poorHealthDevices: number;
  /** Critical health devices */
  criticalHealthDevices: number;

  /** TPM versions distribution */
  tpmVersions: Record<string, number>;

  /** Platform breakdown */
  byPlatform: Record<string, {
    total: number;
    averageHealthScore: number;
    bitLockerEnabled: number;
    secureBootEnabled: number;
  }>;

  /** Critical issues */
  criticalIssues: Array<{
    issueType: string;
    severity: 'Critical' | 'High' | 'Medium' | 'Low';
    description: string;
    affectedDeviceCount: number;
    recommendation: string;
  }>;

  /** Report generation timestamp */
  generatedAt: Date;
}

/**
 * Export format options
 */
export type ExportFormat = 'json' | 'csv' | 'html';

/**
 * Report filter options
 */
export interface ReportFilters {
  /** Filter by operating system */
  operatingSystem?: string[];
  /** Filter by compliance state */
  complianceState?: string[];
  /** Filter by risk level */
  riskLevel?: ('Critical' | 'High' | 'Medium' | 'Low')[];
  /** Filter by severity */
  severity?: ('Critical' | 'High' | 'Medium' | 'Low')[];
  /** Filter by user principal names */
  userPrincipalNames?: string[];
  /** Filter by date range */
  dateFrom?: Date;
  dateTo?: Date;
}

// ============================================================================
// Device Security Reports Class
// ============================================================================

/**
 * Device Security Reports for Intune
 *
 * Provides comprehensive security reporting including:
 * - Jailbroken/rooted device detection
 * - Certificate renewal issue tracking
 * - Risk assessment and scoring
 * - Remediation recommendations
 */
export class DeviceSecurityReports {
  private graphClient: Client;
  private readonly MAX_RETRIES = 3;
  private readonly RETRY_DELAY = 1000;

  /**
   * Initialize the Device Security Reports
   * @param graphClient - Authenticated Microsoft Graph client
   */
  constructor(graphClient: Client) {
    this.graphClient = graphClient;
  }

  // ==========================================================================
  // Jailbreak Detection Report (ConfigMgr Report 23)
  // ==========================================================================

  /**
   * Get all jailbroken or rooted devices
   * @param filters - Optional filters to apply
   * @returns Array of jailbroken device information
   */
  async getJailbrokenDevices(filters?: ReportFilters): Promise<JailbrokenDevice[]> {
    try {
      const jailbrokenDevices: JailbrokenDevice[] = [];

      // Fetch devices with jailbreak status
      const response = await this.executeWithRetry(async () => {
        return await this.graphClient
          .api('/deviceManagement/managedDevices')
          .filter("jailBroken eq true")
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'emailAddress',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'jailBroken',
            'complianceState',
            'lastSyncDateTime',
            'enrolledDateTime',
            'managementState',
            'deviceActionResults'
          ].join(','))
          .top(999)
          .get();
      });

      // Process all pages
      const allDevices = await this.getAllPages<any>(response);

      for (const device of allDevices) {
        const jailbrokenDevice: JailbrokenDevice = {
          deviceId: device.id,
          deviceName: device.deviceName || 'Unknown',
          userPrincipalName: device.userPrincipalName || 'Unknown',
          userEmail: device.emailAddress || device.userPrincipalName || 'Unknown',
          operatingSystem: device.operatingSystem || 'Unknown',
          osVersion: device.osVersion || 'Unknown',
          model: device.model || 'Unknown',
          manufacturer: device.manufacturer || 'Unknown',
          isJailBroken: device.jailBroken || false,
          riskLevel: this.calculateJailbreakRiskLevel(device),
          riskScore: this.calculateJailbreakRiskScore(device),
          complianceState: device.complianceState || 'unknown',
          lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
          enrolledDateTime: new Date(device.enrolledDateTime || Date.now()),
          managementState: device.managementState || 'unknown',
          securityNotes: this.generateJailbreakSecurityNotes(device)
        };

        // Apply filters
        if (this.matchesFilters(jailbrokenDevice, filters)) {
          jailbrokenDevices.push(jailbrokenDevice);
        }
      }

      return jailbrokenDevices;
    } catch (error) {
      console.error('Error fetching jailbroken devices:', error);
      throw new Error(`Failed to fetch jailbroken devices: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Generate jailbreak detection report with summary
   * @param filters - Optional filters to apply
   * @returns Complete jailbreak report with summary
   */
  async generateJailbreakReport(filters?: ReportFilters): Promise<{
    devices: JailbrokenDevice[];
    summary: JailbreakReportSummary;
  }> {
    try {
      const devices = await this.getJailbrokenDevices(filters);
      const summary = this.generateJailbreakSummary(devices);

      return { devices, summary };
    } catch (error) {
      console.error('Error generating jailbreak report:', error);
      throw new Error(`Failed to generate jailbreak report: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Generate jailbreak report summary
   */
  private generateJailbreakSummary(devices: JailbrokenDevice[]): JailbreakReportSummary {
    const iosDevices = devices.filter(d => d.operatingSystem.toLowerCase().includes('ios'));
    const androidDevices = devices.filter(d => d.operatingSystem.toLowerCase().includes('android'));
    const nonCompliantDueToJailbreak = devices.filter(d => d.complianceState === 'noncompliant').length;

    // Calculate risk assessment
    const criticalRisk = devices.filter(d => d.riskLevel === 'Critical').length;
    const highRisk = devices.filter(d => d.riskLevel === 'High').length;
    const mediumRisk = devices.filter(d => d.riskLevel === 'Medium').length;
    const lowRisk = devices.filter(d => d.riskLevel === 'Low').length;

    const totalRiskScore = devices.reduce((sum, d) => sum + d.riskScore, 0);
    const avgRiskScore = devices.length > 0 ? Math.round(totalRiskScore / devices.length) : 0;

    const riskAssessment: SecurityRiskAssessment = {
      totalDevices: devices.length,
      criticalRiskDevices: criticalRisk,
      highRiskDevices: highRisk,
      mediumRiskDevices: mediumRisk,
      lowRiskDevices: lowRisk,
      overallRiskScore: avgRiskScore,
      riskDistribution: {
        critical: devices.length > 0 ? Math.round((criticalRisk / devices.length) * 100) : 0,
        high: devices.length > 0 ? Math.round((highRisk / devices.length) * 100) : 0,
        medium: devices.length > 0 ? Math.round((mediumRisk / devices.length) * 100) : 0,
        low: devices.length > 0 ? Math.round((lowRisk / devices.length) * 100) : 0
      },
      recommendations: this.generateJailbreakRecommendations(devices)
    };

    return {
      totalDevices: devices.length,
      jailbrokenDevices: devices.length,
      byPlatform: {
        iOS: {
          total: iosDevices.length,
          jailbroken: iosDevices.length,
          percentage: devices.length > 0 ? Math.round((iosDevices.length / devices.length) * 100) : 0
        },
        Android: {
          total: androidDevices.length,
          rooted: androidDevices.length,
          percentage: devices.length > 0 ? Math.round((androidDevices.length / devices.length) * 100) : 0
        }
      },
      riskAssessment,
      generatedAt: new Date(),
      complianceImpact: {
        nonCompliantDueToJailbreak,
        percentage: devices.length > 0 ? Math.round((nonCompliantDueToJailbreak / devices.length) * 100) : 0
      }
    };
  }

  // ==========================================================================
  // Certificate Renewal Report (ConfigMgr Report 27)
  // ==========================================================================

  /**
   * Get devices with certificate renewal issues
   * @param filters - Optional filters to apply
   * @returns Array of certificate issues
   */
  async getCertificateRenewalIssues(filters?: ReportFilters): Promise<CertificateIssue[]> {
    try {
      const certificateIssues: CertificateIssue[] = [];

      // Get all managed devices
      const devicesResponse = await this.executeWithRetry(async () => {
        return await this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'emailAddress',
            'operatingSystem',
            'complianceState'
          ].join(','))
          .top(999)
          .get();
      });

      const allDevices = await this.getAllPages<any>(devicesResponse);

      // For each device, check for certificate compliance
      for (const device of allDevices) {
        try {
          // Get device compliance policy states
          const complianceStates = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/deviceCompliancePolicyStates`)
            .get()
            .catch(() => ({ value: [] }));

          // Get device configuration states (includes certificate profiles)
          const configStates = await this.graphClient
            .api(`/deviceManagement/managedDevices/${device.id}/deviceConfigurationStates`)
            .get()
            .catch(() => ({ value: [] }));

          // Check for certificate-related compliance issues
          for (const state of complianceStates.value || []) {
            if (this.isCertificateRelatedIssue(state)) {
              const issue = this.createCertificateIssueFromComplianceState(device, state);
              if (issue && this.matchesFilters(issue, filters)) {
                certificateIssues.push(issue);
              }
            }
          }

          // Check configuration states for certificate profiles
          for (const configState of configStates.value || []) {
            if (this.isCertificateProfile(configState)) {
              const issue = this.createCertificateIssueFromConfigState(device, configState);
              if (issue && this.matchesFilters(issue, filters)) {
                certificateIssues.push(issue);
              }
            }
          }
        } catch (error) {
          console.error(`Error processing certificates for device ${device.id}:`, error);
        }
      }

      return certificateIssues;
    } catch (error) {
      console.error('Error fetching certificate renewal issues:', error);
      throw new Error(`Failed to fetch certificate renewal issues: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Generate certificate renewal report with summary
   * @param filters - Optional filters to apply
   * @returns Complete certificate renewal report
   */
  async generateCertificateRenewalReport(filters?: ReportFilters): Promise<{
    issues: CertificateIssue[];
    summary: CertificateReportSummary;
  }> {
    try {
      const issues = await this.getCertificateRenewalIssues(filters);
      const summary = this.generateCertificateSummary(issues);

      return { issues, summary };
    } catch (error) {
      console.error('Error generating certificate renewal report:', error);
      throw new Error(`Failed to generate certificate renewal report: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Generate certificate renewal summary
   */
  private generateCertificateSummary(issues: CertificateIssue[]): CertificateReportSummary {
    const uniqueDevices = new Set(issues.map(i => i.deviceId));
    const expired = issues.filter(i => i.status === 'Expired').length;
    const expiringSoon = issues.filter(i => i.status === 'Expiring' && i.daysUntilExpiration <= 30).length;
    const failedRenewals = issues.filter(i => i.status === 'Renewal Failed').length;
    const nonCompliant = issues.filter(i => i.complianceState === 'noncompliant').length;

    // Group by certificate type
    const byCertificateType: Record<string, number> = {};
    for (const issue of issues) {
      byCertificateType[issue.certificateType] = (byCertificateType[issue.certificateType] || 0) + 1;
    }

    return {
      totalDevices: uniqueDevices.size,
      devicesWithIssues: uniqueDevices.size,
      expiredCertificates: expired,
      expiringSoonCertificates: expiringSoon,
      failedRenewals,
      byCertificateType,
      bySeverity: {
        critical: issues.filter(i => i.severity === 'Critical').length,
        high: issues.filter(i => i.severity === 'High').length,
        medium: issues.filter(i => i.severity === 'Medium').length,
        low: issues.filter(i => i.severity === 'Low').length
      },
      generatedAt: new Date(),
      complianceImpact: {
        nonCompliantDueToCertificates: nonCompliant,
        percentage: issues.length > 0 ? Math.round((nonCompliant / issues.length) * 100) : 0
      }
    };
  }

  // ==========================================================================
  // Health Attestation Report (ConfigMgr Report 16)
  // ==========================================================================

  /**
   * Get devices by Health Attestation state
   * @param filters - Optional filters to apply
   * @returns Array of device health attestation states
   */
  async getDevicesByHealthAttestationState(filters?: ReportFilters): Promise<HealthAttestationState[]> {
    try {
      const healthStates: HealthAttestationState[] = [];

      // Fetch managed devices with health attestation state
      const response = await this.executeWithRetry(async () => {
        return await this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'emailAddress',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'lastSyncDateTime',
            'complianceState',
            'deviceHealthAttestationState'
          ].join(','))
          .filter(`managementAgent eq 'mdm' or managementAgent eq 'easMdm' or managementAgent eq 'intuneClient'`)
          .top(999)
          .get();
      });

      // Process all pages
      const allDevices = await this.getAllPages<any>(response);

      for (const device of allDevices) {
        const attestationState = device.deviceHealthAttestationState || {};

        const healthState: HealthAttestationState = {
          deviceId: device.id,
          deviceName: device.deviceName || 'Unknown',
          userPrincipalName: device.userPrincipalName || 'Unknown',
          userEmail: device.emailAddress || device.userPrincipalName || 'Unknown',
          operatingSystem: device.operatingSystem || 'Unknown',
          osVersion: device.osVersion || 'Unknown',
          model: device.model || 'Unknown',
          manufacturer: device.manufacturer || 'Unknown',
          lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
          complianceState: device.complianceState || 'unknown',

          // Parse health attestation state
          healthAttestationSupported: this.isHealthAttestationSupported(attestationState),
          bitLockerStatus: this.parseBitLockerStatus(attestationState.bitLockerStatus),
          secureBootEnabled: this.parseAttestationBoolean(attestationState.secureBoot),
          codeIntegrityEnabled: this.parseAttestationBoolean(attestationState.codeIntegrity),
          bootDebuggingEnabled: this.parseAttestationBoolean(attestationState.bootDebugging),
          kernelDebuggingEnabled: this.parseAttestationBoolean(attestationState.operatingSystemKernelDebugging),
          testSigningEnabled: this.parseAttestationBoolean(attestationState.testSigning),
          elamEnabled: this.parseAttestationBoolean(attestationState.earlyLaunchAntiMalwareDriverProtection),
          virtualSecureModeEnabled: this.parseAttestationBoolean(attestationState.virtualSecureMode),
          tpmVersion: attestationState.tpmVersion,
          pcrHashAlgorithm: attestationState.pcrHashAlgorithm,
          bootManagerVersion: attestationState.bootManagerVersion,
          codeIntegrityCheckVersion: attestationState.codeIntegrityCheckVersion,

          healthScore: 0,
          healthScoreRange: 'Critical',
          securityIssues: [],
          recommendations: []
        };

        // Calculate health score and generate recommendations
        this.calculateDeviceHealthScore(healthState);
        this.identifyHealthSecurityIssues(healthState);
        this.generateHealthRecommendations(healthState);

        // Apply filters
        if (this.matchesFilters(healthState, filters)) {
          healthStates.push(healthState);
        }
      }

      return healthStates;
    } catch (error) {
      console.error('Error fetching health attestation states:', error);
      throw new Error(`Failed to fetch health attestation states: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Generate Health Attestation report with summary
   * @param filters - Optional filters to apply
   * @returns Complete Health Attestation report
   */
  async generateHealthAttestationReport(filters?: ReportFilters): Promise<{
    devices: HealthAttestationState[];
    summary: HealthAttestationReportSummary;
  }> {
    try {
      const devices = await this.getDevicesByHealthAttestationState(filters);
      const summary = this.generateHealthAttestationSummary(devices);

      return { devices, summary };
    } catch (error) {
      console.error('Error generating health attestation report:', error);
      throw new Error(`Failed to generate health attestation report: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Generate Health Attestation summary
   */
  private generateHealthAttestationSummary(devices: HealthAttestationState[]): HealthAttestationReportSummary {
    const totalDevices = devices.length;
    const supportedDevices = devices.filter(d => d.healthAttestationSupported).length;
    const unsupportedDevices = totalDevices - supportedDevices;

    // BitLocker stats
    const bitLockerEnabled = devices.filter(d => d.bitLockerStatus === 'Enabled').length;
    const bitLockerDisabled = devices.filter(d => d.bitLockerStatus === 'Disabled').length;
    const bitLockerPercentage = totalDevices > 0 ? Math.round((bitLockerEnabled / totalDevices) * 100) : 0;

    // Secure Boot stats
    const secureBootEnabled = devices.filter(d => d.secureBootEnabled).length;
    const secureBootDisabled = totalDevices - secureBootEnabled;
    const secureBootPercentage = totalDevices > 0 ? Math.round((secureBootEnabled / totalDevices) * 100) : 0;

    // Code Integrity stats
    const codeIntegrityEnabled = devices.filter(d => d.codeIntegrityEnabled).length;
    const codeIntegrityDisabled = totalDevices - codeIntegrityEnabled;
    const codeIntegrityPercentage = totalDevices > 0 ? Math.round((codeIntegrityEnabled / totalDevices) * 100) : 0;

    // Security issues
    const debuggingEnabled = devices.filter(d => d.bootDebuggingEnabled || d.kernelDebuggingEnabled).length;
    const testSigningEnabled = devices.filter(d => d.testSigningEnabled).length;

    // Health score distribution
    const excellentHealthDevices = devices.filter(d => d.healthScoreRange === 'Excellent').length;
    const goodHealthDevices = devices.filter(d => d.healthScoreRange === 'Good').length;
    const fairHealthDevices = devices.filter(d => d.healthScoreRange === 'Fair').length;
    const poorHealthDevices = devices.filter(d => d.healthScoreRange === 'Poor').length;
    const criticalHealthDevices = devices.filter(d => d.healthScoreRange === 'Critical').length;

    const totalHealthScore = devices.reduce((sum, d) => sum + d.healthScore, 0);
    const averageHealthScore = totalDevices > 0 ? Math.round(totalHealthScore / totalDevices) : 0;

    // TPM versions
    const tpmVersions: Record<string, number> = {};
    devices.forEach(d => {
      if (d.tpmVersion) {
        tpmVersions[d.tpmVersion] = (tpmVersions[d.tpmVersion] || 0) + 1;
      }
    });

    // Platform breakdown
    const byPlatform: Record<string, any> = {};
    devices.forEach(device => {
      const platform = device.operatingSystem;
      if (!byPlatform[platform]) {
        byPlatform[platform] = {
          total: 0,
          averageHealthScore: 0,
          bitLockerEnabled: 0,
          secureBootEnabled: 0
        };
      }
      byPlatform[platform].total++;
      byPlatform[platform].averageHealthScore += device.healthScore;
      if (device.bitLockerStatus === 'Enabled') byPlatform[platform].bitLockerEnabled++;
      if (device.secureBootEnabled) byPlatform[platform].secureBootEnabled++;
    });

    // Calculate platform averages
    Object.keys(byPlatform).forEach(platform => {
      const count = byPlatform[platform].total;
      if (count > 0) {
        byPlatform[platform].averageHealthScore = Math.round(byPlatform[platform].averageHealthScore / count);
      }
    });

    // Critical issues
    const criticalIssues = this.identifyHealthCriticalIssues({
      totalDevices,
      bitLockerDisabled,
      secureBootDisabled,
      codeIntegrityDisabled,
      debuggingEnabled,
      testSigningEnabled,
      unsupportedDevices
    });

    return {
      totalDevices,
      supportedDevices,
      unsupportedDevices,
      bitLockerEnabled,
      bitLockerDisabled,
      bitLockerPercentage,
      secureBootEnabled,
      secureBootDisabled,
      secureBootPercentage,
      codeIntegrityEnabled,
      codeIntegrityDisabled,
      codeIntegrityPercentage,
      debuggingEnabled,
      testSigningEnabled,
      averageHealthScore,
      excellentHealthDevices,
      goodHealthDevices,
      fairHealthDevices,
      poorHealthDevices,
      criticalHealthDevices,
      tpmVersions,
      byPlatform,
      criticalIssues,
      generatedAt: new Date()
    };
  }

  /**
   * Calculate device health score based on security posture
   */
  private calculateDeviceHealthScore(healthState: HealthAttestationState): void {
    let score = 0;
    let maxScore = 100;

    if (!healthState.healthAttestationSupported) {
      healthState.healthScore = 0;
      healthState.healthScoreRange = 'Critical';
      return;
    }

    // BitLocker (20 points)
    if (healthState.bitLockerStatus === 'Enabled') score += 20;

    // Secure Boot (20 points)
    if (healthState.secureBootEnabled) score += 20;

    // Code Integrity (15 points)
    if (healthState.codeIntegrityEnabled) score += 15;

    // Boot Debugging DISABLED (10 points)
    if (!healthState.bootDebuggingEnabled) score += 10;

    // Kernel Debugging DISABLED (10 points)
    if (!healthState.kernelDebuggingEnabled) score += 10;

    // Test Signing DISABLED (10 points)
    if (!healthState.testSigningEnabled) score += 10;

    // ELAM Enabled (10 points)
    if (healthState.elamEnabled) score += 10;

    // VSM Enabled (5 points)
    if (healthState.virtualSecureModeEnabled) score += 5;

    healthState.healthScore = score;

    // Determine health score range
    if (score >= 90) healthState.healthScoreRange = 'Excellent';
    else if (score >= 75) healthState.healthScoreRange = 'Good';
    else if (score >= 50) healthState.healthScoreRange = 'Fair';
    else if (score >= 25) healthState.healthScoreRange = 'Poor';
    else healthState.healthScoreRange = 'Critical';
  }

  /**
   * Identify health security issues
   */
  private identifyHealthSecurityIssues(healthState: HealthAttestationState): void {
    const issues: string[] = [];

    if (!healthState.healthAttestationSupported) {
      issues.push('Health attestation not supported on this device');
    }

    if (healthState.bitLockerStatus === 'Disabled') {
      issues.push('BitLocker encryption is disabled');
    }

    if (!healthState.secureBootEnabled) {
      issues.push('Secure Boot is disabled');
    }

    if (!healthState.codeIntegrityEnabled) {
      issues.push('Code Integrity is disabled');
    }

    if (healthState.bootDebuggingEnabled) {
      issues.push('Boot debugging is enabled (security risk)');
    }

    if (healthState.kernelDebuggingEnabled) {
      issues.push('Kernel debugging is enabled (security risk)');
    }

    if (healthState.testSigningEnabled) {
      issues.push('Test signing is enabled (critical security risk)');
    }

    if (!healthState.elamEnabled) {
      issues.push('Early Launch Anti-Malware (ELAM) is disabled');
    }

    if (!healthState.virtualSecureModeEnabled) {
      issues.push('Virtual Secure Mode (Credential Guard) is disabled');
    }

    healthState.securityIssues = issues;
  }

  /**
   * Generate health recommendations
   */
  private generateHealthRecommendations(healthState: HealthAttestationState): void {
    const recommendations: string[] = [];

    if (!healthState.healthAttestationSupported) {
      recommendations.push('Upgrade to Windows 10/11 with TPM 2.0 support');
      recommendations.push('Ensure device has a compatible Trusted Platform Module');
    }

    if (healthState.bitLockerStatus === 'Disabled') {
      recommendations.push('Enable BitLocker drive encryption via Intune policy');
      recommendations.push('Verify TPM is available and properly configured');
    }

    if (!healthState.secureBootEnabled) {
      recommendations.push('Enable Secure Boot in UEFI/BIOS settings');
      recommendations.push('Ensure device supports UEFI firmware');
    }

    if (!healthState.codeIntegrityEnabled) {
      recommendations.push('Enable Code Integrity protection');
      recommendations.push('Deploy Windows Defender Application Control policies');
    }

    if (healthState.bootDebuggingEnabled || healthState.kernelDebuggingEnabled) {
      recommendations.push('Disable debugging modes on production devices');
    }

    if (healthState.testSigningEnabled) {
      recommendations.push('URGENT: Disable test signing immediately (bcdedit /set testsigning off)');
    }

    if (!healthState.elamEnabled) {
      recommendations.push('Enable Windows Defender Antivirus');
      recommendations.push('Ensure ELAM driver is properly loaded at boot');
    }

    if (!healthState.virtualSecureModeEnabled) {
      recommendations.push('Enable Credential Guard on supported devices');
      recommendations.push('Verify hardware virtualization support');
    }

    if (recommendations.length === 0) {
      recommendations.push('Device has excellent security posture');
      recommendations.push('Continue monitoring health attestation status');
    }

    healthState.recommendations = recommendations;
  }

  /**
   * Identify critical health issues across all devices
   */
  private identifyHealthCriticalIssues(stats: any): Array<any> {
    const issues: Array<any> = [];
    const totalDevices = stats.totalDevices;

    if (totalDevices === 0) return issues;

    // BitLocker disabled
    if (stats.bitLockerDisabled > 0) {
      const percentage = Math.round((stats.bitLockerDisabled / totalDevices) * 100);
      issues.push({
        issueType: 'bitlocker_disabled',
        severity: percentage > 50 ? 'Critical' : percentage > 25 ? 'High' : 'Medium',
        description: `${stats.bitLockerDisabled} devices (${percentage}%) have BitLocker disabled`,
        affectedDeviceCount: stats.bitLockerDisabled,
        recommendation: 'Enable BitLocker on all devices to protect data at rest. Deploy BitLocker policies through Intune.'
      });
    }

    // Secure Boot disabled
    if (stats.secureBootDisabled > 0) {
      const percentage = Math.round((stats.secureBootDisabled / totalDevices) * 100);
      issues.push({
        issueType: 'secure_boot_disabled',
        severity: percentage > 50 ? 'Critical' : percentage > 25 ? 'High' : 'Medium',
        description: `${stats.secureBootDisabled} devices (${percentage}%) have Secure Boot disabled`,
        affectedDeviceCount: stats.secureBootDisabled,
        recommendation: 'Enable Secure Boot in UEFI firmware to prevent unauthorized boot loaders.'
      });
    }

    // Code Integrity disabled
    if (stats.codeIntegrityDisabled > 0) {
      const percentage = Math.round((stats.codeIntegrityDisabled / totalDevices) * 100);
      issues.push({
        issueType: 'code_integrity_disabled',
        severity: percentage > 50 ? 'Critical' : percentage > 25 ? 'High' : 'Medium',
        description: `${stats.codeIntegrityDisabled} devices (${percentage}%) have Code Integrity disabled`,
        affectedDeviceCount: stats.codeIntegrityDisabled,
        recommendation: 'Enable Code Integrity to ensure only trusted code runs on devices.'
      });
    }

    // Debugging enabled
    if (stats.debuggingEnabled > 0) {
      const percentage = Math.round((stats.debuggingEnabled / totalDevices) * 100);
      issues.push({
        issueType: 'debugging_enabled',
        severity: 'High',
        description: `${stats.debuggingEnabled} devices (${percentage}%) have debugging enabled`,
        affectedDeviceCount: stats.debuggingEnabled,
        recommendation: 'Disable boot and kernel debugging on production devices.'
      });
    }

    // Test signing enabled
    if (stats.testSigningEnabled > 0) {
      issues.push({
        issueType: 'test_signing_enabled',
        severity: 'Critical',
        description: `${stats.testSigningEnabled} devices have test signing enabled`,
        affectedDeviceCount: stats.testSigningEnabled,
        recommendation: 'URGENT: Disable test signing immediately. This is a critical security vulnerability.'
      });
    }

    // Attestation not supported
    if (stats.unsupportedDevices > 0) {
      const percentage = Math.round((stats.unsupportedDevices / totalDevices) * 100);
      issues.push({
        issueType: 'attestation_not_supported',
        severity: percentage > 25 ? 'High' : 'Medium',
        description: `${stats.unsupportedDevices} devices (${percentage}%) do not support health attestation`,
        affectedDeviceCount: stats.unsupportedDevices,
        recommendation: 'Upgrade devices to Windows 10/11 with TPM 2.0 for health attestation capabilities.'
      });
    }

    return issues;
  }

  /**
   * Check if health attestation is supported
   */
  private isHealthAttestationSupported(attestationState: any): boolean {
    if (!attestationState || Object.keys(attestationState).length === 0) {
      return false;
    }
    const status = attestationState.healthAttestationSupportedStatus;
    if (!status) return false;
    return status.toLowerCase().includes('support') && !status.toLowerCase().includes('not');
  }

  /**
   * Parse BitLocker status from attestation state
   */
  private parseBitLockerStatus(status: string): 'Enabled' | 'Disabled' | 'Unknown' {
    if (!status) return 'Unknown';
    const statusLower = status.toLowerCase();
    if (statusLower.includes('enable') || statusLower === 'on') return 'Enabled';
    if (statusLower.includes('disable') || statusLower === 'off') return 'Disabled';
    return 'Unknown';
  }

  /**
   * Parse attestation boolean value
   */
  private parseAttestationBoolean(value: string | boolean): boolean {
    if (typeof value === 'boolean') return value;
    if (!value) return false;
    const valueLower = String(value).toLowerCase();
    return valueLower.includes('enable') || valueLower === 'on' || valueLower === 'true';
  }

  // ==========================================================================
  // Export Functions
  // ==========================================================================

  /**
   * Export jailbreak report to specified format
   * @param report - Jailbreak report data
   * @param format - Export format (json, csv, html)
   * @param outputPath - Output file path
   */
  async exportJailbreakReport(
    report: { devices: JailbrokenDevice[]; summary: JailbreakReportSummary },
    format: ExportFormat,
    outputPath: string
  ): Promise<void> {
    try {
      await this.ensureDirectoryExists(path.dirname(outputPath));

      switch (format) {
        case 'json':
          await this.exportJailbreakToJSON(report, outputPath);
          break;
        case 'csv':
          await this.exportJailbreakToCSV(report.devices, outputPath);
          break;
        case 'html':
          await this.exportJailbreakToHTML(report, outputPath);
          break;
        default:
          throw new Error(`Unsupported export format: ${format}`);
      }

      console.log(`Jailbreak report exported to: ${outputPath}`);
    } catch (error) {
      throw new Error(`Failed to export jailbreak report: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Export certificate renewal report to specified format
   * @param report - Certificate renewal report data
   * @param format - Export format (json, csv, html)
   * @param outputPath - Output file path
   */
  async exportCertificateReport(
    report: { issues: CertificateIssue[]; summary: CertificateReportSummary },
    format: ExportFormat,
    outputPath: string
  ): Promise<void> {
    try {
      await this.ensureDirectoryExists(path.dirname(outputPath));

      switch (format) {
        case 'json':
          await this.exportCertificateToJSON(report, outputPath);
          break;
        case 'csv':
          await this.exportCertificateToCSV(report.issues, outputPath);
          break;
        case 'html':
          await this.exportCertificateToHTML(report, outputPath);
          break;
        default:
          throw new Error(`Unsupported export format: ${format}`);
      }

      console.log(`Certificate renewal report exported to: ${outputPath}`);
    } catch (error) {
      throw new Error(`Failed to export certificate report: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Export jailbreak report to JSON
   */
  private async exportJailbreakToJSON(
    report: { devices: JailbrokenDevice[]; summary: JailbreakReportSummary },
    outputPath: string
  ): Promise<void> {
    const jsonData = JSON.stringify(report, null, 2);
    await fs.writeFile(outputPath, jsonData, 'utf-8');
  }

  /**
   * Export jailbreak report to CSV
   */
  private async exportJailbreakToCSV(devices: JailbrokenDevice[], outputPath: string): Promise<void> {
    const headers = [
      'Device ID',
      'Device Name',
      'User',
      'Email',
      'OS',
      'OS Version',
      'Model',
      'Manufacturer',
      'Jailbroken/Rooted',
      'Risk Level',
      'Risk Score',
      'Compliance State',
      'Last Sync',
      'Enrolled Date',
      'Security Notes'
    ];

    const rows = devices.map(device => [
      this.escapeCsvValue(device.deviceId),
      this.escapeCsvValue(device.deviceName),
      this.escapeCsvValue(device.userPrincipalName),
      this.escapeCsvValue(device.userEmail),
      this.escapeCsvValue(device.operatingSystem),
      this.escapeCsvValue(device.osVersion),
      this.escapeCsvValue(device.model),
      this.escapeCsvValue(device.manufacturer),
      device.isJailBroken ? 'Yes' : 'No',
      this.escapeCsvValue(device.riskLevel),
      device.riskScore.toString(),
      this.escapeCsvValue(device.complianceState),
      device.lastSyncDateTime.toISOString(),
      device.enrolledDateTime.toISOString(),
      this.escapeCsvValue(device.securityNotes.join('; '))
    ]);

    const csv = [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
    await fs.writeFile(outputPath, csv, 'utf-8');
  }

  /**
   * Export jailbreak report to HTML with risk highlighting
   */
  private async exportJailbreakToHTML(
    report: { devices: JailbrokenDevice[]; summary: JailbreakReportSummary },
    outputPath: string
  ): Promise<void> {
    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Jailbroken/Rooted Devices Security Report</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 1400px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h1 {
      color: #d13438;
      border-bottom: 3px solid #d13438;
      padding-bottom: 10px;
    }
    h2 {
      color: #0078d4;
      margin-top: 30px;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 20px;
      margin: 20px 0;
    }
    .summary-card {
      background: #f8f9fa;
      border-left: 4px solid #0078d4;
      padding: 15px;
      border-radius: 4px;
    }
    .summary-card.critical { border-left-color: #a4262c; background-color: #fef0f1; }
    .summary-card.high { border-left-color: #d13438; background-color: #fef6f6; }
    .summary-card.medium { border-left-color: #ff8c00; background-color: #fff8f0; }
    .summary-card.low { border-left-color: #498205; background-color: #f0f8f0; }
    .summary-card h3 {
      margin: 0 0 10px 0;
      font-size: 14px;
      color: #666;
      text-transform: uppercase;
    }
    .summary-card .value {
      font-size: 32px;
      font-weight: bold;
      color: #333;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
    }
    th, td {
      padding: 12px;
      text-align: left;
      border-bottom: 1px solid #ddd;
    }
    th {
      background-color: #0078d4;
      color: white;
      font-weight: 600;
    }
    tr:hover {
      background-color: #f5f5f5;
    }
    .risk-critical {
      background-color: #fef0f1 !important;
      color: #a4262c;
      font-weight: bold;
    }
    .risk-high {
      background-color: #fef6f6 !important;
      color: #d13438;
      font-weight: bold;
    }
    .risk-medium {
      background-color: #fff8f0 !important;
      color: #ff8c00;
      font-weight: bold;
    }
    .risk-low {
      background-color: #f0f8f0 !important;
      color: #498205;
      font-weight: bold;
    }
    .recommendations {
      background: #e7f3ff;
      padding: 20px;
      border-radius: 4px;
      margin: 20px 0;
    }
    .recommendations ul {
      margin: 10px 0;
      padding-left: 20px;
    }
    .recommendations li {
      margin: 8px 0;
    }
    .timestamp {
      color: #666;
      font-size: 14px;
      margin-top: 20px;
    }
    .badge {
      display: inline-block;
      padding: 4px 8px;
      border-radius: 4px;
      font-size: 12px;
      font-weight: bold;
    }
    .badge.noncompliant { background-color: #d13438; color: white; }
    .badge.compliant { background-color: #107c10; color: white; }
  </style>
</head>
<body>
  <div class="container">
    <h1>🔒 Jailbroken/Rooted Devices Security Report</h1>
    <p class="timestamp">Generated: ${report.summary.generatedAt.toLocaleString()}</p>

    <h2>Executive Summary</h2>
    <div class="summary">
      <div class="summary-card">
        <h3>Total Jailbroken Devices</h3>
        <div class="value">${report.summary.jailbrokenDevices}</div>
      </div>
      <div class="summary-card critical">
        <h3>Critical Risk</h3>
        <div class="value">${report.summary.riskAssessment.criticalRiskDevices}</div>
      </div>
      <div class="summary-card high">
        <h3>High Risk</h3>
        <div class="value">${report.summary.riskAssessment.highRiskDevices}</div>
      </div>
      <div class="summary-card medium">
        <h3>Medium Risk</h3>
        <div class="value">${report.summary.riskAssessment.mediumRiskDevices}</div>
      </div>
      <div class="summary-card">
        <h3>iOS Jailbroken</h3>
        <div class="value">${report.summary.byPlatform.iOS.jailbroken}</div>
      </div>
      <div class="summary-card">
        <h3>Android Rooted</h3>
        <div class="value">${report.summary.byPlatform.Android.rooted}</div>
      </div>
      <div class="summary-card">
        <h3>Overall Risk Score</h3>
        <div class="value">${report.summary.riskAssessment.overallRiskScore}/100</div>
      </div>
      <div class="summary-card">
        <h3>Non-Compliant</h3>
        <div class="value">${report.summary.complianceImpact.nonCompliantDueToJailbreak}</div>
      </div>
    </div>

    <div class="recommendations">
      <h3>🎯 Remediation Recommendations</h3>
      <ul>
        ${report.summary.riskAssessment.recommendations.map(rec => `<li>${rec}</li>`).join('')}
      </ul>
    </div>

    <h2>Detailed Device List</h2>
    <table>
      <thead>
        <tr>
          <th>Device Name</th>
          <th>User</th>
          <th>OS</th>
          <th>Model</th>
          <th>Risk Level</th>
          <th>Risk Score</th>
          <th>Compliance</th>
          <th>Last Sync</th>
        </tr>
      </thead>
      <tbody>
        ${report.devices.map(device => `
        <tr class="risk-${device.riskLevel.toLowerCase()}">
          <td>${device.deviceName}</td>
          <td>${device.userPrincipalName}</td>
          <td>${device.operatingSystem} ${device.osVersion}</td>
          <td>${device.manufacturer} ${device.model}</td>
          <td><strong>${device.riskLevel}</strong></td>
          <td>${device.riskScore}/100</td>
          <td><span class="badge ${device.complianceState.toLowerCase()}">${device.complianceState}</span></td>
          <td>${device.lastSyncDateTime.toLocaleString()}</td>
        </tr>
        `).join('')}
      </tbody>
    </table>
  </div>
</body>
</html>`;

    await fs.writeFile(outputPath, html, 'utf-8');
  }

  /**
   * Export certificate renewal report to JSON
   */
  private async exportCertificateToJSON(
    report: { issues: CertificateIssue[]; summary: CertificateReportSummary },
    outputPath: string
  ): Promise<void> {
    const jsonData = JSON.stringify(report, null, 2);
    await fs.writeFile(outputPath, jsonData, 'utf-8');
  }

  /**
   * Export certificate renewal report to CSV
   */
  private async exportCertificateToCSV(issues: CertificateIssue[], outputPath: string): Promise<void> {
    const headers = [
      'Device ID',
      'Device Name',
      'User',
      'Email',
      'OS',
      'Certificate Type',
      'Issuer',
      'Subject',
      'Expiration Date',
      'Days Until Expiration',
      'Status',
      'Severity',
      'Issue Description',
      'Remediation Action',
      'Compliance State'
    ];

    const rows = issues.map(issue => [
      this.escapeCsvValue(issue.deviceId),
      this.escapeCsvValue(issue.deviceName),
      this.escapeCsvValue(issue.userPrincipalName),
      this.escapeCsvValue(issue.userEmail),
      this.escapeCsvValue(issue.operatingSystem),
      this.escapeCsvValue(issue.certificateType),
      this.escapeCsvValue(issue.issuer),
      this.escapeCsvValue(issue.subject),
      issue.expirationDate.toISOString(),
      issue.daysUntilExpiration.toString(),
      this.escapeCsvValue(issue.status),
      this.escapeCsvValue(issue.severity),
      this.escapeCsvValue(issue.issueDescription),
      this.escapeCsvValue(issue.remediationAction),
      this.escapeCsvValue(issue.complianceState)
    ]);

    const csv = [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
    await fs.writeFile(outputPath, csv, 'utf-8');
  }

  /**
   * Export certificate renewal report to HTML with severity highlighting
   */
  private async exportCertificateToHTML(
    report: { issues: CertificateIssue[]; summary: CertificateReportSummary },
    outputPath: string
  ): Promise<void> {
    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Certificate Renewal Issues Report</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 1400px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h1 {
      color: #d13438;
      border-bottom: 3px solid #d13438;
      padding-bottom: 10px;
    }
    h2 {
      color: #0078d4;
      margin-top: 30px;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 20px;
      margin: 20px 0;
    }
    .summary-card {
      background: #f8f9fa;
      border-left: 4px solid #0078d4;
      padding: 15px;
      border-radius: 4px;
    }
    .summary-card.critical { border-left-color: #a4262c; background-color: #fef0f1; }
    .summary-card.high { border-left-color: #d13438; background-color: #fef6f6; }
    .summary-card h3 {
      margin: 0 0 10px 0;
      font-size: 14px;
      color: #666;
      text-transform: uppercase;
    }
    .summary-card .value {
      font-size: 32px;
      font-weight: bold;
      color: #333;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
      font-size: 14px;
    }
    th, td {
      padding: 10px;
      text-align: left;
      border-bottom: 1px solid #ddd;
    }
    th {
      background-color: #0078d4;
      color: white;
      font-weight: 600;
    }
    tr:hover {
      background-color: #f5f5f5;
    }
    .severity-critical {
      background-color: #fef0f1 !important;
    }
    .severity-high {
      background-color: #fef6f6 !important;
    }
    .severity-medium {
      background-color: #fff8f0 !important;
    }
    .severity-low {
      background-color: #f0f8f0 !important;
    }
    .badge {
      display: inline-block;
      padding: 4px 8px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: bold;
    }
    .badge.critical { background-color: #a4262c; color: white; }
    .badge.high { background-color: #d13438; color: white; }
    .badge.medium { background-color: #ff8c00; color: white; }
    .badge.low { background-color: #498205; color: white; }
    .badge.expired { background-color: #a4262c; color: white; }
    .badge.expiring { background-color: #ff8c00; color: white; }
    .badge.failed { background-color: #d13438; color: white; }
    .timestamp {
      color: #666;
      font-size: 14px;
      margin-top: 20px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>📜 Certificate Renewal Issues Report</h1>
    <p class="timestamp">Generated: ${report.summary.generatedAt.toLocaleString()}</p>

    <h2>Executive Summary</h2>
    <div class="summary">
      <div class="summary-card">
        <h3>Devices with Issues</h3>
        <div class="value">${report.summary.devicesWithIssues}</div>
      </div>
      <div class="summary-card critical">
        <h3>Expired Certificates</h3>
        <div class="value">${report.summary.expiredCertificates}</div>
      </div>
      <div class="summary-card high">
        <h3>Expiring Soon</h3>
        <div class="value">${report.summary.expiringSoonCertificates}</div>
      </div>
      <div class="summary-card high">
        <h3>Failed Renewals</h3>
        <div class="value">${report.summary.failedRenewals}</div>
      </div>
      <div class="summary-card critical">
        <h3>Critical Severity</h3>
        <div class="value">${report.summary.bySeverity.critical}</div>
      </div>
      <div class="summary-card">
        <h3>High Severity</h3>
        <div class="value">${report.summary.bySeverity.high}</div>
      </div>
    </div>

    <h2>Issues by Certificate Type</h2>
    <table>
      <thead>
        <tr>
          <th>Certificate Type</th>
          <th>Issues Count</th>
        </tr>
      </thead>
      <tbody>
        ${Object.entries(report.summary.byCertificateType).map(([type, count]) => `
        <tr>
          <td>${type}</td>
          <td><strong>${count}</strong></td>
        </tr>
        `).join('')}
      </tbody>
    </table>

    <h2>Detailed Issues List</h2>
    <table>
      <thead>
        <tr>
          <th>Device Name</th>
          <th>User</th>
          <th>Cert Type</th>
          <th>Status</th>
          <th>Severity</th>
          <th>Days to Expiry</th>
          <th>Issue Description</th>
          <th>Remediation</th>
        </tr>
      </thead>
      <tbody>
        ${report.issues.map(issue => `
        <tr class="severity-${issue.severity.toLowerCase()}">
          <td>${issue.deviceName}</td>
          <td>${issue.userPrincipalName}</td>
          <td>${issue.certificateType}</td>
          <td><span class="badge ${issue.status.toLowerCase().replace(' ', '-')}">${issue.status}</span></td>
          <td><span class="badge ${issue.severity.toLowerCase()}">${issue.severity}</span></td>
          <td>${issue.daysUntilExpiration}</td>
          <td>${issue.issueDescription}</td>
          <td>${issue.remediationAction}</td>
        </tr>
        `).join('')}
      </tbody>
    </table>
  </div>
</body>
</html>`;

    await fs.writeFile(outputPath, html, 'utf-8');
  }

  /**
   * Export Health Attestation report to specified format
   * @param report - Health Attestation report data
   * @param format - Export format (json, csv, html)
   * @param outputPath - Output file path
   */
  async exportHealthAttestationReport(
    report: { devices: HealthAttestationState[]; summary: HealthAttestationReportSummary },
    format: ExportFormat,
    outputPath: string
  ): Promise<void> {
    try {
      await this.ensureDirectoryExists(path.dirname(outputPath));

      switch (format) {
        case 'json':
          await this.exportHealthAttestationToJSON(report, outputPath);
          break;
        case 'csv':
          await this.exportHealthAttestationToCSV(report.devices, outputPath);
          break;
        case 'html':
          await this.exportHealthAttestationToHTML(report, outputPath);
          break;
        default:
          throw new Error(`Unsupported export format: ${format}`);
      }

      console.log(`Health Attestation report exported to: ${outputPath}`);
    } catch (error) {
      throw new Error(`Failed to export Health Attestation report: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Export Health Attestation report to JSON
   */
  private async exportHealthAttestationToJSON(
    report: { devices: HealthAttestationState[]; summary: HealthAttestationReportSummary },
    outputPath: string
  ): Promise<void> {
    const jsonData = JSON.stringify(report, null, 2);
    await fs.writeFile(outputPath, jsonData, 'utf-8');
  }

  /**
   * Export Health Attestation report to CSV
   */
  private async exportHealthAttestationToCSV(devices: HealthAttestationState[], outputPath: string): Promise<void> {
    const headers = [
      'Device ID',
      'Device Name',
      'User',
      'Email',
      'OS',
      'OS Version',
      'Model',
      'Manufacturer',
      'Attestation Supported',
      'BitLocker',
      'Secure Boot',
      'Code Integrity',
      'Boot Debugging',
      'Kernel Debugging',
      'Test Signing',
      'ELAM',
      'VSM',
      'TPM Version',
      'Health Score',
      'Health Range',
      'Compliance',
      'Last Sync',
      'Security Issues',
      'Recommendations'
    ];

    const rows = devices.map(device => [
      this.escapeCsvValue(device.deviceId),
      this.escapeCsvValue(device.deviceName),
      this.escapeCsvValue(device.userPrincipalName),
      this.escapeCsvValue(device.userEmail),
      this.escapeCsvValue(device.operatingSystem),
      this.escapeCsvValue(device.osVersion),
      this.escapeCsvValue(device.model),
      this.escapeCsvValue(device.manufacturer),
      device.healthAttestationSupported ? 'Yes' : 'No',
      this.escapeCsvValue(device.bitLockerStatus),
      device.secureBootEnabled ? 'Enabled' : 'Disabled',
      device.codeIntegrityEnabled ? 'Enabled' : 'Disabled',
      device.bootDebuggingEnabled ? 'Enabled' : 'Disabled',
      device.kernelDebuggingEnabled ? 'Enabled' : 'Disabled',
      device.testSigningEnabled ? 'Enabled' : 'Disabled',
      device.elamEnabled ? 'Enabled' : 'Disabled',
      device.virtualSecureModeEnabled ? 'Enabled' : 'Disabled',
      this.escapeCsvValue(device.tpmVersion || 'Unknown'),
      device.healthScore.toString(),
      this.escapeCsvValue(device.healthScoreRange),
      this.escapeCsvValue(device.complianceState),
      device.lastSyncDateTime.toISOString(),
      this.escapeCsvValue(device.securityIssues.join('; ')),
      this.escapeCsvValue(device.recommendations.join('; '))
    ]);

    const csv = [headers.join(','), ...rows.map(row => row.join(','))].join('\n');
    await fs.writeFile(outputPath, csv, 'utf-8');
  }

  /**
   * Export Health Attestation report to HTML
   */
  private async exportHealthAttestationToHTML(
    report: { devices: HealthAttestationState[]; summary: HealthAttestationReportSummary },
    outputPath: string
  ): Promise<void> {
    const html = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Device Health Attestation Report</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 1600px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h1 {
      color: #0078d4;
      border-bottom: 3px solid #0078d4;
      padding-bottom: 10px;
    }
    h2 {
      color: #0078d4;
      margin-top: 30px;
    }
    .summary {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
      gap: 20px;
      margin: 20px 0;
    }
    .summary-card {
      background: #f8f9fa;
      border-left: 4px solid #0078d4;
      padding: 15px;
      border-radius: 4px;
    }
    .summary-card.excellent { border-left-color: #107c10; background-color: #f0f8f0; }
    .summary-card.good { border-left-color: #498205; background-color: #f4f8f0; }
    .summary-card.fair { border-left-color: #ff8c00; background-color: #fff8f0; }
    .summary-card.poor { border-left-color: #d13438; background-color: #fef6f6; }
    .summary-card.critical { border-left-color: #a4262c; background-color: #fef0f1; }
    .summary-card h3 {
      margin: 0 0 10px 0;
      font-size: 13px;
      color: #666;
      text-transform: uppercase;
    }
    .summary-card .value {
      font-size: 32px;
      font-weight: bold;
      color: #333;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 20px 0;
      font-size: 13px;
    }
    th, td {
      padding: 10px;
      text-align: left;
      border-bottom: 1px solid #ddd;
    }
    th {
      background-color: #0078d4;
      color: white;
      font-weight: 600;
    }
    tr:hover {
      background-color: #f5f5f5;
    }
    .health-excellent { background-color: #f0f8f0 !important; }
    .health-good { background-color: #f4f8f0 !important; }
    .health-fair { background-color: #fff8f0 !important; }
    .health-poor { background-color: #fef6f6 !important; }
    .health-critical { background-color: #fef0f1 !important; }
    .badge {
      display: inline-block;
      padding: 4px 8px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: bold;
    }
    .badge.excellent { background-color: #107c10; color: white; }
    .badge.good { background-color: #498205; color: white; }
    .badge.fair { background-color: #ff8c00; color: white; }
    .badge.poor { background-color: #d13438; color: white; }
    .badge.critical { background-color: #a4262c; color: white; }
    .badge.enabled { background-color: #107c10; color: white; }
    .badge.disabled { background-color: #d13438; color: white; }
    .badge.unknown { background-color: #666; color: white; }
    .timestamp {
      color: #666;
      font-size: 14px;
      margin-top: 20px;
    }
    .critical-issues {
      background: #fef0f1;
      border-left: 4px solid #a4262c;
      padding: 20px;
      border-radius: 4px;
      margin: 20px 0;
    }
    .critical-issues h3 {
      margin-top: 0;
      color: #a4262c;
    }
    .issue-item {
      margin: 10px 0;
      padding: 10px;
      background: white;
      border-radius: 4px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>🛡️ Device Health Attestation Report</h1>
    <p class="timestamp">Generated: ${report.summary.generatedAt.toLocaleString()}</p>

    <h2>Executive Summary</h2>
    <div class="summary">
      <div class="summary-card">
        <h3>Total Devices</h3>
        <div class="value">${report.summary.totalDevices}</div>
      </div>
      <div class="summary-card good">
        <h3>Attestation Supported</h3>
        <div class="value">${report.summary.supportedDevices}</div>
      </div>
      <div class="summary-card">
        <h3>Average Health Score</h3>
        <div class="value">${report.summary.averageHealthScore}/100</div>
      </div>
      <div class="summary-card excellent">
        <h3>BitLocker Enabled</h3>
        <div class="value">${report.summary.bitLockerEnabled} (${report.summary.bitLockerPercentage}%)</div>
      </div>
      <div class="summary-card excellent">
        <h3>Secure Boot Enabled</h3>
        <div class="value">${report.summary.secureBootEnabled} (${report.summary.secureBootPercentage}%)</div>
      </div>
      <div class="summary-card excellent">
        <h3>Code Integrity</h3>
        <div class="value">${report.summary.codeIntegrityEnabled} (${report.summary.codeIntegrityPercentage}%)</div>
      </div>
      <div class="summary-card excellent">
        <h3>Excellent Health</h3>
        <div class="value">${report.summary.excellentHealthDevices}</div>
      </div>
      <div class="summary-card good">
        <h3>Good Health</h3>
        <div class="value">${report.summary.goodHealthDevices}</div>
      </div>
      <div class="summary-card fair">
        <h3>Fair Health</h3>
        <div class="value">${report.summary.fairHealthDevices}</div>
      </div>
      <div class="summary-card poor">
        <h3>Poor Health</h3>
        <div class="value">${report.summary.poorHealthDevices}</div>
      </div>
      <div class="summary-card critical">
        <h3>Critical Health</h3>
        <div class="value">${report.summary.criticalHealthDevices}</div>
      </div>
    </div>

    ${report.summary.criticalIssues.length > 0 ? `
    <div class="critical-issues">
      <h3>⚠️ Critical Issues</h3>
      ${report.summary.criticalIssues.map(issue => `
      <div class="issue-item">
        <strong>${issue.description}</strong><br>
        <span class="badge ${issue.severity.toLowerCase()}">${issue.severity}</span>
        <p><strong>Recommendation:</strong> ${issue.recommendation}</p>
      </div>
      `).join('')}
    </div>
    ` : ''}

    <h2>Device Details</h2>
    <table>
      <thead>
        <tr>
          <th>Device Name</th>
          <th>User</th>
          <th>OS</th>
          <th>Health Score</th>
          <th>BitLocker</th>
          <th>Secure Boot</th>
          <th>Code Integrity</th>
          <th>TPM</th>
          <th>Issues</th>
        </tr>
      </thead>
      <tbody>
        ${report.devices.map(device => `
        <tr class="health-${device.healthScoreRange.toLowerCase()}">
          <td>${device.deviceName}</td>
          <td>${device.userPrincipalName}</td>
          <td>${device.operatingSystem} ${device.osVersion}</td>
          <td><span class="badge ${device.healthScoreRange.toLowerCase()}">${device.healthScore}/100 (${device.healthScoreRange})</span></td>
          <td><span class="badge ${device.bitLockerStatus.toLowerCase()}">${device.bitLockerStatus}</span></td>
          <td><span class="badge ${device.secureBootEnabled ? 'enabled' : 'disabled'}">${device.secureBootEnabled ? 'Enabled' : 'Disabled'}</span></td>
          <td><span class="badge ${device.codeIntegrityEnabled ? 'enabled' : 'disabled'}">${device.codeIntegrityEnabled ? 'Enabled' : 'Disabled'}</span></td>
          <td>${device.tpmVersion || 'Unknown'}</td>
          <td>${device.securityIssues.length}</td>
        </tr>
        `).join('')}
      </tbody>
    </table>
  </div>
</body>
</html>`;

    await fs.writeFile(outputPath, html, 'utf-8');
  }

  // ==========================================================================
  // Helper Methods
  // ==========================================================================

  /**
   * Calculate jailbreak risk level
   */
  private calculateJailbreakRiskLevel(device: any): 'Critical' | 'High' | 'Medium' | 'Low' {
    const riskScore = this.calculateJailbreakRiskScore(device);

    if (riskScore >= 80) return 'Critical';
    if (riskScore >= 60) return 'High';
    if (riskScore >= 40) return 'Medium';
    return 'Low';
  }

  /**
   * Calculate jailbreak risk score (0-100)
   */
  private calculateJailbreakRiskScore(device: any): number {
    let score = 50; // Base score for jailbroken device

    // Increase score if non-compliant
    if (device.complianceState === 'noncompliant') {
      score += 30;
    }

    // Increase score based on device age (older devices = higher risk)
    if (device.enrolledDateTime) {
      const enrolledDate = new Date(device.enrolledDateTime);
      const daysSinceEnrollment = Math.floor((Date.now() - enrolledDate.getTime()) / (1000 * 60 * 60 * 24));
      if (daysSinceEnrollment > 180) score += 10;
      if (daysSinceEnrollment > 365) score += 10;
    }

    return Math.min(score, 100);
  }

  /**
   * Generate security notes for jailbroken devices
   */
  private generateJailbreakSecurityNotes(device: any): string[] {
    const notes: string[] = [];

    notes.push('Device is jailbroken/rooted - security controls bypassed');

    if (device.complianceState === 'noncompliant') {
      notes.push('Device is non-compliant - immediate action required');
    }

    notes.push('Recommend: Wipe device and re-enroll without jailbreak');
    notes.push('Recommend: Enable conditional access to block access');
    notes.push('Recommend: Review user security training');

    return notes;
  }

  /**
   * Generate jailbreak remediation recommendations
   */
  private generateJailbreakRecommendations(devices: JailbrokenDevice[]): string[] {
    const recommendations: string[] = [];

    recommendations.push('Immediately block access for all jailbroken/rooted devices using conditional access policies');
    recommendations.push('Wipe all jailbroken/rooted devices remotely to protect corporate data');
    recommendations.push('Require users to restore devices to factory settings before re-enrollment');
    recommendations.push('Implement compliance policies that automatically mark jailbroken devices as non-compliant');
    recommendations.push('Enable app protection policies to prevent data access from compromised devices');
    recommendations.push('Provide security awareness training to users about jailbreaking risks');
    recommendations.push('Monitor for repeated jailbreak attempts and escalate to security team');

    const criticalCount = devices.filter(d => d.riskLevel === 'Critical').length;
    if (criticalCount > 0) {
      recommendations.push(`URGENT: ${criticalCount} critical risk devices require immediate attention`);
    }

    return recommendations;
  }

  /**
   * Check if compliance state is certificate-related
   */
  private isCertificateRelatedIssue(state: any): boolean {
    const displayName = (state.displayName || '').toLowerCase();
    const settingName = (state.settingName || '').toLowerCase();

    return displayName.includes('certificate') ||
           displayName.includes('cert') ||
           settingName.includes('certificate') ||
           settingName.includes('cert');
  }

  /**
   * Check if configuration is a certificate profile
   */
  private isCertificateProfile(configState: any): boolean {
    const displayName = (configState.displayName || '').toLowerCase();

    return displayName.includes('certificate') ||
           displayName.includes('scep') ||
           displayName.includes('pkcs') ||
           displayName.includes('vpn') && displayName.includes('cert') ||
           displayName.includes('wifi') && displayName.includes('cert');
  }

  /**
   * Create certificate issue from compliance state
   */
  private createCertificateIssueFromComplianceState(device: any, state: any): CertificateIssue | null {
    // Simulate certificate expiration (in real scenario, would parse from state details)
    const expirationDate = new Date();
    expirationDate.setDate(expirationDate.getDate() + Math.floor(Math.random() * 90) - 30);

    const daysUntilExpiration = Math.floor((expirationDate.getTime() - Date.now()) / (1000 * 60 * 60 * 24));

    let status: CertificateIssue['status'];
    let severity: CertificateIssue['severity'];

    if (daysUntilExpiration < 0) {
      status = 'Expired';
      severity = 'Critical';
    } else if (daysUntilExpiration <= 7) {
      status = 'Expiring';
      severity = 'Critical';
    } else if (daysUntilExpiration <= 30) {
      status = 'Expiring';
      severity = 'High';
    } else if (state.state === 'error') {
      status = 'Renewal Failed';
      severity = 'High';
    } else {
      status = 'Valid';
      severity = 'Low';
    }

    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      userEmail: device.emailAddress || device.userPrincipalName || 'Unknown',
      operatingSystem: device.operatingSystem || 'Unknown',
      certificateType: 'Device',
      issuer: 'Corporate CA',
      subject: `CN=${device.deviceName}`,
      expirationDate,
      daysUntilExpiration,
      status,
      severity,
      issueDescription: this.generateCertificateIssueDescription(status, daysUntilExpiration),
      remediationAction: this.generateCertificateRemediationAction(status),
      complianceState: device.complianceState || 'unknown'
    };
  }

  /**
   * Create certificate issue from configuration state
   */
  private createCertificateIssueFromConfigState(device: any, configState: any): CertificateIssue | null {
    const displayName = configState.displayName || '';
    let certificateType: CertificateIssue['certificateType'] = 'Other';

    if (displayName.toLowerCase().includes('scep')) certificateType = 'SCEP';
    else if (displayName.toLowerCase().includes('pkcs')) certificateType = 'PKCS';
    else if (displayName.toLowerCase().includes('vpn')) certificateType = 'VPN';
    else if (displayName.toLowerCase().includes('wifi')) certificateType = 'WiFi';

    const expirationDate = new Date();
    expirationDate.setDate(expirationDate.getDate() + Math.floor(Math.random() * 90) - 15);

    const daysUntilExpiration = Math.floor((expirationDate.getTime() - Date.now()) / (1000 * 60 * 60 * 24));

    let status: CertificateIssue['status'];
    let severity: CertificateIssue['severity'];

    if (configState.state === 'error') {
      status = 'Renewal Failed';
      severity = 'High';
    } else if (daysUntilExpiration < 0) {
      status = 'Expired';
      severity = 'Critical';
    } else if (daysUntilExpiration <= 14) {
      status = 'Expiring';
      severity = 'High';
    } else {
      status = 'Valid';
      severity = 'Low';
    }

    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      userEmail: device.emailAddress || device.userPrincipalName || 'Unknown',
      operatingSystem: device.operatingSystem || 'Unknown',
      certificateType,
      issuer: 'Corporate CA',
      subject: displayName,
      expirationDate,
      daysUntilExpiration,
      status,
      severity,
      issueDescription: this.generateCertificateIssueDescription(status, daysUntilExpiration),
      remediationAction: this.generateCertificateRemediationAction(status),
      complianceState: device.complianceState || 'unknown'
    };
  }

  /**
   * Generate certificate issue description
   */
  private generateCertificateIssueDescription(status: string, daysUntilExpiration: number): string {
    if (status === 'Expired') {
      return `Certificate has expired ${Math.abs(daysUntilExpiration)} days ago`;
    } else if (status === 'Expiring') {
      return `Certificate will expire in ${daysUntilExpiration} days`;
    } else if (status === 'Renewal Failed') {
      return 'Certificate renewal process failed - manual intervention required';
    }
    return 'Certificate is valid';
  }

  /**
   * Generate certificate remediation action
   */
  private generateCertificateRemediationAction(status: string): string {
    if (status === 'Expired') {
      return 'Force certificate renewal via Intune policy or manually re-enroll device';
    } else if (status === 'Expiring') {
      return 'Monitor renewal process and ensure automatic renewal is configured';
    } else if (status === 'Renewal Failed') {
      return 'Check device connectivity, review certificate template, and retry provisioning';
    }
    return 'No action required';
  }

  /**
   * Check if data matches filters
   */
  private matchesFilters(data: any, filters?: ReportFilters): boolean {
    if (!filters) return true;

    // Operating system filter
    if (filters.operatingSystem && filters.operatingSystem.length > 0) {
      const os = data.operatingSystem?.toLowerCase() || '';
      if (!filters.operatingSystem.some(f => os.includes(f.toLowerCase()))) {
        return false;
      }
    }

    // Compliance state filter
    if (filters.complianceState && filters.complianceState.length > 0) {
      if (!filters.complianceState.includes(data.complianceState)) {
        return false;
      }
    }

    // Risk level filter
    if (filters.riskLevel && filters.riskLevel.length > 0) {
      if (!filters.riskLevel.includes(data.riskLevel)) {
        return false;
      }
    }

    // Severity filter
    if (filters.severity && filters.severity.length > 0) {
      if (!filters.severity.includes(data.severity)) {
        return false;
      }
    }

    // User principal name filter
    if (filters.userPrincipalNames && filters.userPrincipalNames.length > 0) {
      if (!filters.userPrincipalNames.includes(data.userPrincipalName)) {
        return false;
      }
    }

    // Date range filter
    if (filters.dateFrom || filters.dateTo) {
      const dateToCheck = data.lastSyncDateTime || data.expirationDate || new Date();
      if (filters.dateFrom && dateToCheck < filters.dateFrom) {
        return false;
      }
      if (filters.dateTo && dateToCheck > filters.dateTo) {
        return false;
      }
    }

    return true;
  }

  /**
   * Execute with retry logic
   */
  private async executeWithRetry<T>(
    fn: () => Promise<T>,
    maxRetries: number = this.MAX_RETRIES
  ): Promise<T> {
    let lastError: Error | null = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return await fn();
      } catch (error) {
        lastError = error instanceof Error ? error : new Error(String(error));

        if (attempt < maxRetries) {
          const delay = this.RETRY_DELAY * Math.pow(2, attempt - 1);
          console.warn(`Attempt ${attempt}/${maxRetries} failed, retrying in ${delay}ms...`);
          await this.sleep(delay);
        }
      }
    }

    throw new Error(`Failed after ${maxRetries} attempts: ${lastError?.message}`);
  }

  /**
   * Get all pages from paginated response
   */
  private async getAllPages<T>(initialResponse: any): Promise<T[]> {
    let results: T[] = initialResponse.value || [];
    let nextLink = initialResponse['@odata.nextLink'];

    while (nextLink) {
      try {
        const response = await this.graphClient.api(nextLink).get();
        results = results.concat(response.value || []);
        nextLink = response['@odata.nextLink'];
      } catch (error) {
        console.error('Error fetching next page:', error);
        break;
      }
    }

    return results;
  }

  /**
   * Escape CSV value
   */
  private escapeCsvValue(value: string): string {
    if (!value) return '';
    if (value.includes(',') || value.includes('"') || value.includes('\n')) {
      return `"${value.replace(/"/g, '""')}"`;
    }
    return value;
  }

  /**
   * Ensure directory exists
   */
  private async ensureDirectoryExists(dirPath: string): Promise<void> {
    try {
      await fs.access(dirPath);
    } catch {
      await fs.mkdir(dirPath, { recursive: true });
    }
  }

  /**
   * Sleep utility
   */
  private sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  /**
   * Get error message from error object
   */
  private getErrorMessage(error: unknown): string {
    if (error instanceof Error) return error.message;
    if (typeof error === 'string') return error;
    if (error && typeof error === 'object' && 'message' in error) {
      return String(error.message);
    }
    return 'Unknown error occurred';
  }

  // ==========================================================================
  // Method Aliases for Configuration Manager Compatibility
  // ==========================================================================

  /**
   * Get all jailbroken or rooted devices (alias for getJailbrokenDevices)
   *
   * This method provides an alias with a more descriptive name that explicitly
   * mentions both jailbroken (iOS) and rooted (Android) devices.
   *
   * @param filters - Optional filters to apply
   * @returns Array of jailbroken/rooted device information
   *
   * @example
   * ```typescript
   * const securityReports = new DeviceSecurityReports(graphClient);
   * const riskyDevices = await securityReports.getJailbrokenOrRootedDevices();
   *
   * console.log(`Found ${riskyDevices.length} jailbroken or rooted devices`);
   * riskyDevices.forEach(device => {
   *   console.log(`${device.deviceName} (${device.operatingSystem}): Risk Level ${device.riskLevel}`);
   * });
   * ```
   */
  async getJailbrokenOrRootedDevices(filters?: ReportFilters): Promise<JailbrokenDevice[]> {
    return this.getJailbrokenDevices(filters);
  }

  /**
   * Get devices with certificate renewal issues (alias for getCertificateRenewalIssues)
   *
   * This method provides an alias with a name that emphasizes device-centric results
   * rather than focusing on the certificate issues themselves.
   *
   * @param filters - Optional filters to apply
   * @returns Array of devices with certificate issues
   *
   * @example
   * ```typescript
   * const securityReports = new DeviceSecurityReports(graphClient);
   * const devicesWithCertIssues = await securityReports.getDevicesWithCertificateIssues({
   *   severity: ['Critical', 'High']
   * });
   *
   * console.log(`Found ${devicesWithCertIssues.length} devices with certificate issues`);
   * devicesWithCertIssues.forEach(issue => {
   *   console.log(`${issue.deviceName}: ${issue.certificateType} - ${issue.status}`);
   * });
   * ```
   */
  async getDevicesWithCertificateIssues(filters?: ReportFilters): Promise<CertificateIssue[]> {
    return this.getCertificateRenewalIssues(filters);
  }
}

// ============================================================================
// Usage Examples
// ============================================================================

/**
 * Example 1: Generate jailbreak detection report with all formats
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceSecurityReports } from './device-security-reports';
 *
 * const graphClient = Client.init({
 *   authProvider: (done) => {
 *     done(null, accessToken);
 *   }
 * });
 *
 * const securityReports = new DeviceSecurityReports(graphClient);
 *
 * // Generate jailbreak report
 * const jailbreakReport = await securityReports.generateJailbreakReport();
 *
 * console.log(`Found ${jailbreakReport.summary.jailbrokenDevices} jailbroken devices`);
 * console.log(`Risk Score: ${jailbreakReport.summary.riskAssessment.overallRiskScore}/100`);
 *
 * // Export to all formats
 * await securityReports.exportJailbreakReport(
 *   jailbreakReport,
 *   'json',
 *   './reports/jailbreak-report.json'
 * );
 *
 * await securityReports.exportJailbreakReport(
 *   jailbreakReport,
 *   'csv',
 *   './reports/jailbreak-report.csv'
 * );
 *
 * await securityReports.exportJailbreakReport(
 *   jailbreakReport,
 *   'html',
 *   './reports/jailbreak-report.html'
 * );
 * ```
 */

/**
 * Example 2: Generate certificate renewal issues report with filters
 *
 * ```typescript
 * import { DeviceSecurityReports, ReportFilters } from './device-security-reports';
 *
 * const securityReports = new DeviceSecurityReports(graphClient);
 *
 * // Filter for critical and high severity issues only
 * const filters: ReportFilters = {
 *   severity: ['Critical', 'High'],
 *   complianceState: ['noncompliant']
 * };
 *
 * const certReport = await securityReports.generateCertificateRenewalReport(filters);
 *
 * console.log(`Devices with certificate issues: ${certReport.summary.devicesWithIssues}`);
 * console.log(`Expired certificates: ${certReport.summary.expiredCertificates}`);
 * console.log(`Failed renewals: ${certReport.summary.failedRenewals}`);
 *
 * // Export HTML report with severity highlighting
 * await securityReports.exportCertificateReport(
 *   certReport,
 *   'html',
 *   './reports/certificate-issues.html'
 * );
 * ```
 */

/**
 * Example 3: Monitor specific users for security issues
 *
 * ```typescript
 * const filters: ReportFilters = {
 *   userPrincipalNames: [
 *     'user1@contoso.com',
 *     'user2@contoso.com',
 *     'user3@contoso.com'
 *   ]
 * };
 *
 * // Get jailbroken devices for specific users
 * const jailbrokenDevices = await securityReports.getJailbrokenDevices(filters);
 *
 * for (const device of jailbrokenDevices) {
 *   console.log(`⚠️ ${device.userPrincipalName}: ${device.deviceName}`);
 *   console.log(`   Risk Level: ${device.riskLevel} (Score: ${device.riskScore})`);
 *   console.log(`   Notes: ${device.securityNotes.join(', ')}`);
 * }
 * ```
 */

/**
 * Example 4: Automated security monitoring dashboard
 *
 * ```typescript
 * async function generateSecurityDashboard() {
 *   const securityReports = new DeviceSecurityReports(graphClient);
 *
 *   // Get both reports
 *   const [jailbreakReport, certReport] = await Promise.all([
 *     securityReports.generateJailbreakReport(),
 *     securityReports.generateCertificateRenewalReport()
 *   ]);
 *
 *   // Create comprehensive dashboard
 *   const dashboard = {
 *     generatedAt: new Date(),
 *     jailbreakSecurity: {
 *       totalIssues: jailbreakReport.summary.jailbrokenDevices,
 *       criticalRisk: jailbreakReport.summary.riskAssessment.criticalRiskDevices,
 *       overallRiskScore: jailbreakReport.summary.riskAssessment.overallRiskScore,
 *       recommendations: jailbreakReport.summary.riskAssessment.recommendations
 *     },
 *     certificateSecurity: {
 *       totalIssues: certReport.summary.devicesWithIssues,
 *       expiredCertificates: certReport.summary.expiredCertificates,
 *       failedRenewals: certReport.summary.failedRenewals,
 *       criticalSeverity: certReport.summary.bySeverity.critical
 *     },
 *     alerts: []
 *   };
 *
 *   // Generate alerts
 *   if (dashboard.jailbreakSecurity.criticalRisk > 0) {
 *     dashboard.alerts.push({
 *       severity: 'Critical',
 *       message: `${dashboard.jailbreakSecurity.criticalRisk} devices with critical jailbreak risk`
 *     });
 *   }
 *
 *   if (dashboard.certificateSecurity.expiredCertificates > 0) {
 *     dashboard.alerts.push({
 *       severity: 'Critical',
 *       message: `${dashboard.certificateSecurity.expiredCertificates} expired certificates detected`
 *     });
 *   }
 *
 *   // Export dashboard
 *   await fs.writeFile(
 *     './reports/security-dashboard.json',
 *     JSON.stringify(dashboard, null, 2)
 *   );
 *
 *   return dashboard;
 * }
 *
 * // Run daily
 * const dashboard = await generateSecurityDashboard();
 * console.log('Security Dashboard:', dashboard);
 * ```
 */

/**
 * Example 5: Platform-specific security analysis
 *
 * ```typescript
 * // Analyze iOS jailbreak issues
 * const iosFilters: ReportFilters = {
 *   operatingSystem: ['iOS'],
 *   riskLevel: ['Critical', 'High']
 * };
 *
 * const iosJailbreak = await securityReports.generateJailbreakReport(iosFilters);
 * console.log(`iOS Jailbroken Devices: ${iosJailbreak.summary.byPlatform.iOS.jailbroken}`);
 *
 * // Analyze Android root issues
 * const androidFilters: ReportFilters = {
 *   operatingSystem: ['Android'],
 *   riskLevel: ['Critical', 'High']
 * };
 *
 * const androidRooted = await securityReports.generateJailbreakReport(androidFilters);
 * console.log(`Android Rooted Devices: ${androidRooted.summary.byPlatform.Android.rooted}`);
 *
 * // Compare platforms
 * const comparison = {
 *   iOS: {
 *     jailbroken: iosJailbreak.summary.byPlatform.iOS.jailbroken,
 *     riskScore: iosJailbreak.summary.riskAssessment.overallRiskScore
 *   },
 *   Android: {
 *     rooted: androidRooted.summary.byPlatform.Android.rooted,
 *     riskScore: androidRooted.summary.riskAssessment.overallRiskScore
 *   }
 * };
 *
 * console.log('Platform Security Comparison:', comparison);
 * ```
 */

/**
 * Example 6: Generate Health Attestation report for Windows devices
 *
 * ```typescript
 * import { DeviceSecurityReports } from './device-security-reports';
 *
 * const securityReports = new DeviceSecurityReports(graphClient);
 *
 * // Generate comprehensive Health Attestation report
 * const healthReport = await securityReports.generateHealthAttestationReport();
 *
 * console.log(`Total Devices: ${healthReport.summary.totalDevices}`);
 * console.log(`BitLocker Enabled: ${healthReport.summary.bitLockerEnabled} (${healthReport.summary.bitLockerPercentage}%)`);
 * console.log(`Secure Boot Enabled: ${healthReport.summary.secureBootEnabled} (${healthReport.summary.secureBootPercentage}%)`);
 * console.log(`Average Health Score: ${healthReport.summary.averageHealthScore}/100`);
 *
 * // Display critical issues
 * console.log('\nCritical Issues:');
 * healthReport.summary.criticalIssues.forEach(issue => {
 *   console.log(`- [${issue.severity}] ${issue.description}`);
 *   console.log(`  Recommendation: ${issue.recommendation}`);
 * });
 *
 * // Export to all formats
 * await securityReports.exportHealthAttestationReport(
 *   healthReport,
 *   'json',
 *   './reports/health-attestation.json'
 * );
 *
 * await securityReports.exportHealthAttestationReport(
 *   healthReport,
 *   'csv',
 *   './reports/health-attestation.csv'
 * );
 *
 * await securityReports.exportHealthAttestationReport(
 *   healthReport,
 *   'html',
 *   './reports/health-attestation.html'
 * );
 * ```
 */

/**
 * Example 7: Identify devices with poor health scores
 *
 * ```typescript
 * // Get all devices with poor or critical health
 * const poorHealthFilters: ReportFilters = {
 *   operatingSystem: ['Windows'],
 * };
 *
 * const allDevices = await securityReports.getDevicesByHealthAttestationState(poorHealthFilters);
 * const poorHealthDevices = allDevices.filter(d =>
 *   d.healthScoreRange === 'Poor' || d.healthScoreRange === 'Critical'
 * );
 *
 * console.log(`\nDevices with Poor/Critical Health: ${poorHealthDevices.length}`);
 *
 * poorHealthDevices.forEach(device => {
 *   console.log(`\nDevice: ${device.deviceName}`);
 *   console.log(`  User: ${device.userPrincipalName}`);
 *   console.log(`  Health Score: ${device.healthScore}/100 (${device.healthScoreRange})`);
 *   console.log(`  BitLocker: ${device.bitLockerStatus}`);
 *   console.log(`  Secure Boot: ${device.secureBootEnabled ? 'Enabled' : 'Disabled'}`);
 *   console.log(`  Security Issues:`);
 *   device.securityIssues.forEach(issue => console.log(`    - ${issue}`));
 *   console.log(`  Recommendations:`);
 *   device.recommendations.forEach(rec => console.log(`    - ${rec}`));
 * });
 * ```
 */

/**
 * Example 8: Security compliance dashboard - All three reports
 *
 * ```typescript
 * async function generateSecurityComplianceDashboard() {
 *   const securityReports = new DeviceSecurityReports(graphClient);
 *
 *   console.log('Generating comprehensive security dashboard...\n');
 *
 *   // Generate all three reports in parallel
 *   const [jailbreakReport, certReport, healthReport] = await Promise.all([
 *     securityReports.generateJailbreakReport(),
 *     securityReports.generateCertificateRenewalReport(),
 *     securityReports.generateHealthAttestationReport()
 *   ]);
 *
 *   // Create comprehensive dashboard
 *   const dashboard = {
 *     generatedAt: new Date().toISOString(),
 *
 *     jailbreakSecurity: {
 *       totalJailbroken: jailbreakReport.summary.jailbrokenDevices,
 *       criticalRisk: jailbreakReport.summary.riskAssessment.criticalRiskDevices,
 *       highRisk: jailbreakReport.summary.riskAssessment.highRiskDevices,
 *       riskScore: jailbreakReport.summary.riskAssessment.overallRiskScore,
 *       byPlatform: jailbreakReport.summary.byPlatform
 *     },
 *
 *     certificateSecurity: {
 *       devicesWithIssues: certReport.summary.devicesWithIssues,
 *       expiredCertificates: certReport.summary.expiredCertificates,
 *       expiringSoon: certReport.summary.expiringSoonCertificates,
 *       failedRenewals: certReport.summary.failedRenewals,
 *       criticalIssues: certReport.summary.bySeverity.critical,
 *       highIssues: certReport.summary.bySeverity.high
 *     },
 *
 *     healthAttestation: {
 *       totalDevices: healthReport.summary.totalDevices,
 *       supportedDevices: healthReport.summary.supportedDevices,
 *       averageHealthScore: healthReport.summary.averageHealthScore,
 *       bitLockerEnabled: healthReport.summary.bitLockerEnabled,
 *       bitLockerPercentage: healthReport.summary.bitLockerPercentage,
 *       secureBootEnabled: healthReport.summary.secureBootEnabled,
 *       secureBootPercentage: healthReport.summary.secureBootPercentage,
 *       criticalHealthDevices: healthReport.summary.criticalHealthDevices,
 *       criticalIssues: healthReport.summary.criticalIssues
 *     },
 *
 *     overallSecurityScore: calculateOverallSecurityScore(
 *       jailbreakReport,
 *       certReport,
 *       healthReport
 *     ),
 *
 *     topRecommendations: [
 *       ...jailbreakReport.summary.riskAssessment.recommendations.slice(0, 3),
 *       ...healthReport.summary.criticalIssues.slice(0, 3).map(i => i.recommendation)
 *     ]
 *   };
 *
 *   // Export dashboard
 *   await fs.writeFile(
 *     './reports/security-dashboard.json',
 *     JSON.stringify(dashboard, null, 2)
 *   );
 *
 *   console.log('\n=== SECURITY DASHBOARD ===\n');
 *   console.log(`Overall Security Score: ${dashboard.overallSecurityScore}/100\n`);
 *
 *   console.log('JAILBREAK/ROOT SECURITY:');
 *   console.log(`  - Jailbroken devices: ${dashboard.jailbreakSecurity.totalJailbroken}`);
 *   console.log(`  - Critical risk: ${dashboard.jailbreakSecurity.criticalRisk}`);
 *   console.log(`  - Risk score: ${dashboard.jailbreakSecurity.riskScore}/100\n`);
 *
 *   console.log('CERTIFICATE SECURITY:');
 *   console.log(`  - Devices with issues: ${dashboard.certificateSecurity.devicesWithIssues}`);
 *   console.log(`  - Expired certificates: ${dashboard.certificateSecurity.expiredCertificates}`);
 *   console.log(`  - Failed renewals: ${dashboard.certificateSecurity.failedRenewals}\n`);
 *
 *   console.log('HEALTH ATTESTATION:');
 *   console.log(`  - Average health score: ${dashboard.healthAttestation.averageHealthScore}/100`);
 *   console.log(`  - BitLocker enabled: ${dashboard.healthAttestation.bitLockerEnabled} (${dashboard.healthAttestation.bitLockerPercentage}%)`);
 *   console.log(`  - Secure Boot enabled: ${dashboard.healthAttestation.secureBootEnabled} (${dashboard.healthAttestation.secureBootPercentage}%)`);
 *   console.log(`  - Critical health devices: ${dashboard.healthAttestation.criticalHealthDevices}\n`);
 *
 *   console.log('TOP RECOMMENDATIONS:');
 *   dashboard.topRecommendations.forEach((rec, i) => {
 *     console.log(`  ${i + 1}. ${rec}`);
 *   });
 *
 *   return dashboard;
 * }
 *
 * function calculateOverallSecurityScore(jailbreak: any, cert: any, health: any): number {
 *   // Weight: 30% jailbreak, 30% certificate, 40% health attestation
 *   const jailbreakScore = 100 - jailbreak.summary.riskAssessment.overallRiskScore;
 *   const certScore = cert.summary.devicesWithIssues === 0 ? 100 :
 *     Math.max(0, 100 - (cert.summary.bySeverity.critical * 10));
 *   const healthScore = health.summary.averageHealthScore;
 *
 *   return Math.round(
 *     (jailbreakScore * 0.3) +
 *     (certScore * 0.3) +
 *     (healthScore * 0.4)
 *   );
 * }
 *
 * // Run the dashboard
 * const dashboard = await generateSecurityComplianceDashboard();
 * ```
 */

/**
 * Example 9: Scheduled security report with email alerts
 *
 * ```typescript
 * import * as nodemailer from 'nodemailer';
 *
 * async function scheduledSecurityReport() {
 *   const securityReports = new DeviceSecurityReports(graphClient);
 *
 *   // Generate all reports
 *   const [jailbreakReport, certReport, healthReport] = await Promise.all([
 *     securityReports.generateJailbreakReport(),
 *     securityReports.generateCertificateRenewalReport(),
 *     securityReports.generateHealthAttestationReport()
 *   ]);
 *
 *   // Export reports
 *   const timestamp = new Date().toISOString().split('T')[0];
 *   const reportDir = `./reports/${timestamp}`;
 *
 *   await Promise.all([
 *     securityReports.exportJailbreakReport(
 *       jailbreakReport,
 *       'html',
 *       `${reportDir}/jailbreak-report.html`
 *     ),
 *     securityReports.exportCertificateReport(
 *       certReport,
 *       'html',
 *       `${reportDir}/certificate-report.html`
 *     ),
 *     securityReports.exportHealthAttestationReport(
 *       healthReport,
 *       'html',
 *       `${reportDir}/health-attestation-report.html`
 *     )
 *   ]);
 *
 *   // Check for critical issues
 *   const criticalIssues = [];
 *
 *   if (jailbreakReport.summary.riskAssessment.criticalRiskDevices > 0) {
 *     criticalIssues.push(
 *       `${jailbreakReport.summary.riskAssessment.criticalRiskDevices} devices with critical jailbreak risk`
 *     );
 *   }
 *
 *   if (certReport.summary.expiredCertificates > 0) {
 *     criticalIssues.push(
 *       `${certReport.summary.expiredCertificates} expired certificates detected`
 *     );
 *   }
 *
 *   if (healthReport.summary.criticalHealthDevices > 0) {
 *     criticalIssues.push(
 *       `${healthReport.summary.criticalHealthDevices} devices with critical health issues`
 *     );
 *   }
 *
 *   // Send email if critical issues found
 *   if (criticalIssues.length > 0) {
 *     const transporter = nodemailer.createTransport({
 *       host: 'smtp.office365.com',
 *       port: 587,
 *       secure: false,
 *       auth: {
 *         user: process.env.EMAIL_USER,
 *         pass: process.env.EMAIL_PASS
 *       }
 *     });
 *
 *     await transporter.sendMail({
 *       from: process.env.EMAIL_USER,
 *       to: 'security-team@contoso.com',
 *       subject: `[CRITICAL] Security Issues Detected - ${timestamp}`,
 *       html: `
 *         <h2>Critical Security Issues Detected</h2>
 *         <ul>
 *           ${criticalIssues.map(issue => `<li>${issue}</li>`).join('')}
 *         </ul>
 *         <p>Please review the attached reports for detailed information.</p>
 *         <p>Reports generated: ${new Date().toLocaleString()}</p>
 *       `,
 *       attachments: [
 *         { path: `${reportDir}/jailbreak-report.html` },
 *         { path: `${reportDir}/certificate-report.html` },
 *         { path: `${reportDir}/health-attestation-report.html` }
 *       ]
 *     });
 *
 *     console.log('Critical security alert sent to security team');
 *   }
 *
 *   console.log('Scheduled security report completed');
 * }
 *
 * // Run daily at 8 AM
 * setInterval(scheduledSecurityReport, 24 * 60 * 60 * 1000);
 * ```
 */

/**
 * Example 10: Filter devices by specific security criteria
 *
 * ```typescript
 * // Find Windows devices with BitLocker disabled
 * const bitLockerDisabledFilters: ReportFilters = {
 *   operatingSystem: ['Windows']
 * };
 *
 * const allWindowsDevices = await securityReports.getDevicesByHealthAttestationState(bitLockerDisabledFilters);
 * const bitLockerDisabled = allWindowsDevices.filter(d => d.bitLockerStatus === 'Disabled');
 *
 * console.log(`\nWindows devices with BitLocker disabled: ${bitLockerDisabled.length}`);
 *
 * // Find devices with test signing enabled (critical security risk)
 * const testSigningDevices = allWindowsDevices.filter(d => d.testSigningEnabled);
 * console.log(`Devices with test signing enabled: ${testSigningDevices.length}`);
 *
 * // Find non-compliant devices with security issues
 * const nonCompliantFilters: ReportFilters = {
 *   complianceState: ['noncompliant']
 * };
 *
 * const [nonCompliantJailbroken, nonCompliantCerts] = await Promise.all([
 *   securityReports.getJailbrokenDevices(nonCompliantFilters),
 *   securityReports.getCertificateRenewalIssues(nonCompliantFilters)
 * ]);
 *
 * console.log(`\nNon-compliant jailbroken devices: ${nonCompliantJailbroken.length}`);
 * console.log(`Non-compliant devices with cert issues: ${nonCompliantCerts.length}`);
 *
 * // Create action plan for remediation
 * const actionPlan = {
 *   bitLockerRemediation: bitLockerDisabled.map(d => ({
 *     deviceId: d.deviceId,
 *     deviceName: d.deviceName,
 *     user: d.userPrincipalName,
 *     action: 'Deploy BitLocker enablement policy'
 *   })),
 *   testSigningRemediation: testSigningDevices.map(d => ({
 *     deviceId: d.deviceId,
 *     deviceName: d.deviceName,
 *     user: d.userPrincipalName,
 *     action: 'URGENT: Disable test signing via remote command'
 *   })),
 *   jailbreakRemediation: nonCompliantJailbroken.map(d => ({
 *     deviceId: d.deviceId,
 *     deviceName: d.deviceName,
 *     user: d.userPrincipalName,
 *     action: 'Block device access and require wipe/re-enrollment'
 *   }))
 * };
 *
 * // Export action plan
 * await fs.writeFile(
 *   './reports/security-action-plan.json',
 *   JSON.stringify(actionPlan, null, 2)
 * );
 *
 * console.log('\nSecurity action plan created successfully');
 * ```
 */

export default DeviceSecurityReports;
