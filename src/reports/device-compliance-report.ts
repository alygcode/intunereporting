/**
 * Comprehensive Device Compliance Reporting Module for Microsoft Intune
 *
 * This module provides extensive device compliance reporting capabilities using
 * Microsoft Graph API, including:
 * 1. All device compliance policies
 * 2. Device compliance status per policy
 * 3. Detailed compliance data per device
 * 4. Non-compliant devices with specific reasons
 * 5. Compliance trends over time
 * 6. Multiple export formats (JSON, CSV, HTML)
 *
 * @module device-compliance-report
 */

import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Device compliance state enumeration
 */
export enum ComplianceState {
  COMPLIANT = 'compliant',
  NON_COMPLIANT = 'noncompliant',
  IN_GRACE_PERIOD = 'inGracePeriod',
  CONFIG_MANAGER = 'configManager',
  ERROR = 'error',
  UNKNOWN = 'unknown',
  CONFLICT = 'conflict',
  NOT_ASSIGNED = 'notAssigned'
}

/**
 * Compliance policy platform types
 */
export enum CompliancePlatform {
  ANDROID = 'android',
  ANDROID_WORK_PROFILE = 'androidWorkProfile',
  ANDROID_AOSP = 'androidAOSP',
  IOS = 'iOS',
  MAC_OS = 'macOS',
  WINDOWS_10 = 'windows10',
  WINDOWS_81 = 'windows81',
  WINDOWS_PHONE_81 = 'windowsPhone81'
}

/**
 * Compliance policy setting state
 */
export enum SettingComplianceState {
  COMPLIANT = 'compliant',
  NON_COMPLIANT = 'nonCompliant',
  NOT_APPLICABLE = 'notApplicable',
  ERROR = 'error',
  CONFLICT = 'conflict',
  UNKNOWN = 'unknown'
}

/**
 * Device compliance policy information
 */
export interface DeviceCompliancePolicy {
  id: string;
  displayName: string;
  description?: string;
  platform: CompliancePlatform;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  version: number;
  assignmentCount: number;
  settingsCount: number;
  scheduledActionsCount: number;
  isAssigned: boolean;
}

/**
 * Policy device state summary
 */
export interface PolicyDeviceStateSummary {
  policyId: string;
  policyName: string;
  compliantDeviceCount: number;
  nonCompliantDeviceCount: number;
  errorDeviceCount: number;
  conflictDeviceCount: number;
  notApplicableDeviceCount: number;
  inGracePeriodCount: number;
  unknownDeviceCount: number;
  remediatedDeviceCount: number;
  notAssignedDeviceCount: number;
  totalDeviceCount: number;
  compliancePercentage: number;
}

/**
 * Device compliance status for a specific policy
 */
export interface DevicePolicyComplianceStatus {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  policyId: string;
  policyName: string;
  platform: string;
  osVersion: string;
  complianceState: ComplianceState;
  lastReportedDateTime: Date;
  complianceGracePeriodExpirationDateTime?: Date;
  userEmail?: string;
  userName?: string;
  deviceModel?: string;
  manufacturer?: string;
}

/**
 * Detailed device compliance information
 */
export interface DeviceComplianceDetails {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  osVersion: string;
  overallComplianceState: ComplianceState;
  lastSyncDateTime: Date;
  policies: DevicePolicyState[];
  compliancePoliciesCount: number;
  compliantPoliciesCount: number;
  nonCompliantPoliciesCount: number;
  isManaged: boolean;
  isSupervised?: boolean;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  imei?: string;
}

/**
 * Policy state for a specific device
 */
export interface DevicePolicyState {
  policyId: string;
  policyName: string;
  platform: string;
  state: ComplianceState;
  version: number;
  settingStates: PolicySettingState[];
  settingCount: number;
  lastReportedDateTime: Date;
  userPrincipalName?: string;
  userName?: string;
}

/**
 * Policy setting state information
 */
export interface PolicySettingState {
  settingName: string;
  settingInstance?: string;
  state: SettingComplianceState;
  errorCode?: string;
  errorDescription?: string;
  userId?: string;
  userName?: string;
  sources?: PolicySettingSource[];
  currentValue?: string;
  instanceDisplayName?: string;
}

/**
 * Source of a policy setting
 */
export interface PolicySettingSource {
  id: string;
  displayName: string;
  sourceType: string;
}

/**
 * Non-compliant device with detailed reasons
 */
export interface NonCompliantDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  osVersion: string;
  complianceState: ComplianceState;
  nonCompliantPolicies: NonCompliantPolicyInfo[];
  totalNonCompliantSettings: number;
  lastSyncDateTime: Date;
  gracePeriodExpirationDateTime?: Date;
  deviceModel?: string;
  manufacturer?: string;
  userEmail?: string;
}

/**
 * Non-compliant policy information
 */
export interface NonCompliantPolicyInfo {
  policyId: string;
  policyName: string;
  nonCompliantSettings: NonCompliantSetting[];
  settingCount: number;
}

/**
 * Non-compliant setting details
 */
export interface NonCompliantSetting {
  settingName: string;
  settingInstance?: string;
  state: SettingComplianceState;
  errorCode?: string;
  errorDescription?: string;
  currentValue?: string;
  remediationActions?: string[];
}

/**
 * Compliance trend data point
 */
export interface ComplianceTrendDataPoint {
  date: Date;
  compliantCount: number;
  nonCompliantCount: number;
  errorCount: number;
  unknownCount: number;
  totalCount: number;
  compliancePercentage: number;
  snapshotId?: string;
}

/**
 * Compliance trends over time
 */
export interface ComplianceTrends {
  startDate: Date;
  endDate: Date;
  dataPoints: ComplianceTrendDataPoint[];
  averageCompliancePercentage: number;
  complianceChange: number;
  trend: 'improving' | 'declining' | 'stable';
  peakCompliance: ComplianceTrendDataPoint;
  lowestCompliance: ComplianceTrendDataPoint;
}

/**
 * Compliance report summary
 */
export interface ComplianceReportSummary {
  totalDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  errorDevices: number;
  unknownDevices: number;
  inGracePeriodDevices: number;
  compliancePercentage: number;
  totalPolicies: number;
  assignedPolicies: number;
  byPlatform: Record<string, PlatformComplianceSummary>;
  topNonCompliantPolicies: PolicyComplianceSummary[];
  criticalIssues: CriticalIssue[];
  generatedAt: Date;
}

/**
 * Platform-specific compliance summary
 */
export interface PlatformComplianceSummary {
  total: number;
  compliant: number;
  nonCompliant: number;
  error: number;
  compliancePercentage: number;
}

/**
 * Policy compliance summary
 */
export interface PolicyComplianceSummary {
  policyId: string;
  policyName: string;
  platform: string;
  totalDevices: number;
  nonCompliantDevices: number;
  compliancePercentage: number;
}

/**
 * Critical compliance issue
 */
export interface CriticalIssue {
  issueType: 'policy_not_assigned' | 'high_failure_rate' | 'many_errors' | 'grace_period_expiring';
  severity: 'high' | 'medium' | 'low';
  description: string;
  affectedCount: number;
  recommendation: string;
  relatedPolicyId?: string;
  relatedPolicyName?: string;
}

/**
 * Filter options for compliance queries
 */
export interface ComplianceFilterOptions {
  platform?: CompliancePlatform;
  complianceState?: ComplianceState;
  policyId?: string;
  userPrincipalName?: string;
  deviceName?: string;
  includeDetails?: boolean;
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Device Compliance Report
 *
 * Comprehensive device compliance reporting with:
 * - All compliance policies
 * - Policy-level compliance statistics
 * - Device-level compliance details
 * - Non-compliant devices with reasons
 * - Compliance trends over time
 * - Multiple export formats
 */
export class DeviceComplianceReport extends BaseReport {
  name = 'device-compliance-report';
  description = 'Comprehensive device compliance reporting including policies, device status, non-compliance reasons, and trends';
  category = 'Compliance';
  enabled = true;

  /**
   * Execute the complete compliance report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Comprehensive Device Compliance Report');

    try {
      const startTime = Date.now();

      // Fetch all data in parallel for better performance
      const [
        policies,
        policySummaries,
        deviceComplianceDetails,
        nonCompliantDevices
      ] = await Promise.all([
        this.getAllCompliancePolicies(),
        this.getAllPolicyDeviceStateSummaries(),
        this.getAllDeviceComplianceDetails(),
        this.getNonCompliantDevicesWithReasons()
      ]);

      // Generate compliance trends (historical data)
      const trends = await this.getComplianceTrends(30); // Last 30 days

      // Generate summary statistics
      const summary = this.generateComplianceSummary({
        policies,
        policySummaries,
        deviceComplianceDetails,
        nonCompliantDevices,
        trends
      });

      const duration = Date.now() - startTime;

      logger.info('Device Compliance Report completed', {
        duration: `${duration}ms`,
        policies: policies.length,
        devices: deviceComplianceDetails.length,
        nonCompliantDevices: nonCompliantDevices.length
      });

      // Combine all data for the report
      const data = [
        ...policies.map(p => ({ type: 'policy', ...p })),
        ...policySummaries.map(s => ({ type: 'policy-summary', ...s })),
        ...deviceComplianceDetails.map(d => ({ type: 'device-compliance', ...d })),
        ...nonCompliantDevices.map(d => ({ type: 'non-compliant-device', ...d })),
        ...(trends?.dataPoints || []).map(t => ({ type: 'trend-data', ...t }))
      ];

      return {
        metadata: this.createMetadata(this.name, data.length, {
          duration: `${duration}ms`,
          trendDays: 30
        }),
        data,
        summary
      };
    } catch (error) {
      logger.error('Failed to execute Device Compliance Report', error);
      throw error;
    }
  }

  // ============================================================================
  // 1. Get All Device Compliance Policies
  // ============================================================================

  /**
   * Retrieves all device compliance policies from Intune
   */
  async getAllCompliancePolicies(options?: ComplianceFilterOptions): Promise<DeviceCompliancePolicy[]> {
    logger.info('Fetching all device compliance policies');
    const policies: DeviceCompliancePolicy[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceCompliancePolicies')
          .expand('assignments,scheduledActionsForRule')
          .select([
            'id',
            'displayName',
            'description',
            'createdDateTime',
            'lastModifiedDateTime',
            'version',
            '@odata.type'
          ])
          .top(999)
          .get()
      );

      const allPolicies = await this.getAllPages(response);

      for (const policy of allPolicies) {
        const parsedPolicy = this.parseCompliancePolicy(policy);

        // Apply filters if provided
        if (this.matchesPolicyFilter(parsedPolicy, options)) {
          // Get settings count for this policy
          try {
            const settingsCount = await this.getPolicySettingsCount(policy.id);
            parsedPolicy.settingsCount = settingsCount;
          } catch (error) {
            logger.warn(`Failed to fetch settings count for policy ${policy.id}`, error);
          }

          policies.push(parsedPolicy);
        }
      }

      logger.info(`Fetched ${policies.length} device compliance policies`);
      return policies;
    } catch (error) {
      logger.error('Error fetching device compliance policies', error);
      throw new Error(`Failed to fetch compliance policies: ${error.message}`);
    }
  }

  /**
   * Gets the count of settings in a compliance policy
   */
  private async getPolicySettingsCount(policyId: string): Promise<number> {
    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/deviceCompliancePolicies/${policyId}`)
          .get()
      );

      // Count non-null configuration properties
      let count = 0;
      const excludeProps = ['id', 'displayName', 'description', 'createdDateTime',
                           'lastModifiedDateTime', 'version', '@odata.type', '@odata.context'];

      for (const key in response) {
        if (!excludeProps.includes(key) && response[key] !== null && response[key] !== undefined) {
          if (typeof response[key] === 'boolean' || typeof response[key] === 'number' ||
              typeof response[key] === 'string' || Array.isArray(response[key])) {
            count++;
          }
        }
      }

      return count;
    } catch (error) {
      return 0;
    }
  }

  // ============================================================================
  // 2. Get Device Compliance Status Per Policy
  // ============================================================================

  /**
   * Retrieves device state summary for all compliance policies
   */
  async getAllPolicyDeviceStateSummaries(): Promise<PolicyDeviceStateSummary[]> {
    logger.info('Fetching policy device state summaries');
    const summaries: PolicyDeviceStateSummary[] = [];

    try {
      const policies = await this.getAllCompliancePolicies();

      // Fetch summaries for each policy
      for (const policy of policies) {
        try {
          const summary = await this.getPolicyDeviceStateSummary(policy.id, policy.displayName);
          if (summary) {
            summaries.push(summary);
          }
        } catch (error) {
          logger.warn(`Failed to fetch summary for policy ${policy.id}`, error);
        }
      }

      logger.info(`Fetched ${summaries.length} policy device state summaries`);
      return summaries;
    } catch (error) {
      logger.error('Error fetching policy device state summaries', error);
      return summaries;
    }
  }

  /**
   * Gets device state summary for a specific compliance policy
   */
  async getPolicyDeviceStateSummary(policyId: string, policyName: string): Promise<PolicyDeviceStateSummary | null> {
    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/deviceCompliancePolicies/${policyId}/deviceStateSummary`)
          .get()
      );

      if (response) {
        const total = (response.compliantDeviceCount || 0) +
                     (response.nonCompliantDeviceCount || 0) +
                     (response.errorDeviceCount || 0) +
                     (response.conflictDeviceCount || 0) +
                     (response.notApplicableDeviceCount || 0) +
                     (response.inGracePeriodCount || 0) +
                     (response.unknownDeviceCount || 0) +
                     (response.remediatedDeviceCount || 0);

        const compliant = response.compliantDeviceCount || 0;
        const compliancePercentage = total > 0 ? Math.round((compliant / total) * 100 * 100) / 100 : 0;

        return {
          policyId,
          policyName,
          compliantDeviceCount: compliant,
          nonCompliantDeviceCount: response.nonCompliantDeviceCount || 0,
          errorDeviceCount: response.errorDeviceCount || 0,
          conflictDeviceCount: response.conflictDeviceCount || 0,
          notApplicableDeviceCount: response.notApplicableDeviceCount || 0,
          inGracePeriodCount: response.inGracePeriodCount || 0,
          unknownDeviceCount: response.unknownDeviceCount || 0,
          remediatedDeviceCount: response.remediatedDeviceCount || 0,
          notAssignedDeviceCount: response.notAssignedDeviceCount || 0,
          totalDeviceCount: total,
          compliancePercentage
        };
      }

      return null;
    } catch (error) {
      return null;
    }
  }

  /**
   * Gets device statuses for a specific compliance policy
   */
  async getPolicyDeviceStatuses(policyId: string): Promise<DevicePolicyComplianceStatus[]> {
    logger.info(`Fetching device statuses for policy ${policyId}`);
    const statuses: DevicePolicyComplianceStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/deviceCompliancePolicies/${policyId}/deviceStatuses`)
          .select([
            'id',
            'deviceDisplayName',
            'userName',
            'userPrincipalName',
            'deviceModel',
            'platform',
            'complianceGracePeriodExpirationDateTime',
            'status',
            'lastReportedDateTime',
            'userEmail'
          ])
          .top(999)
          .get()
      );

      const allStatuses = await this.getAllPages(response);

      // Get policy name
      const policy = await this.graphClient
        .api(`/deviceManagement/deviceCompliancePolicies/${policyId}`)
        .select('displayName')
        .get();

      for (const status of allStatuses) {
        statuses.push(this.parsePolicyDeviceStatus(status, policyId, policy.displayName));
      }

      logger.info(`Fetched ${statuses.length} device statuses for policy ${policyId}`);
      return statuses;
    } catch (error) {
      logger.error(`Error fetching device statuses for policy ${policyId}`, error);
      return statuses;
    }
  }

  // ============================================================================
  // 3. Get Detailed Compliance Data Per Device
  // ============================================================================

  /**
   * Retrieves detailed compliance information for all devices
   */
  async getAllDeviceComplianceDetails(options?: ComplianceFilterOptions): Promise<DeviceComplianceDetails[]> {
    logger.info('Fetching detailed device compliance information');
    const deviceDetails: DeviceComplianceDetails[] = [];

    try {
      // Get all managed devices
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'operatingSystem',
            'osVersion',
            'complianceState',
            'lastSyncDateTime',
            'isSupervised',
            'manufacturer',
            'model',
            'serialNumber',
            'imei',
            'managementAgent'
          ])
          .filter(`managementAgent eq 'mdm' or managementAgent eq 'easMdm' or managementAgent eq 'intuneClient'`)
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(response);
      logger.info(`Processing compliance details for ${devices.length} devices`);

      // Process devices in batches to avoid timeout
      const batchSize = 50;
      for (let i = 0; i < devices.length; i += batchSize) {
        const batch = devices.slice(i, Math.min(i + batchSize, devices.length));

        const batchPromises = batch.map(async (device) => {
          try {
            const details = await this.getDeviceComplianceDetails(device);

            // Apply filters if provided
            if (this.matchesDeviceFilter(details, options)) {
              return details;
            }
          } catch (error) {
            logger.warn(`Failed to fetch compliance details for device ${device.id}`, error);
          }
          return null;
        });

        const batchResults = await Promise.all(batchPromises);
        deviceDetails.push(...batchResults.filter(d => d !== null) as DeviceComplianceDetails[]);
      }

      logger.info(`Fetched compliance details for ${deviceDetails.length} devices`);
      return deviceDetails;
    } catch (error) {
      logger.error('Error fetching device compliance details', error);
      return deviceDetails;
    }
  }

  /**
   * Gets detailed compliance information for a specific device
   */
  async getDeviceComplianceDetails(device: any): Promise<DeviceComplianceDetails> {
    // Get compliance policy states for this device
    const policyStates = await this.getDevicePolicyStates(device.id);

    const compliantPolicies = policyStates.filter(p => p.state === ComplianceState.COMPLIANT).length;
    const nonCompliantPolicies = policyStates.filter(p => p.state === ComplianceState.NON_COMPLIANT).length;

    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      platform: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      overallComplianceState: this.mapComplianceState(device.complianceState),
      lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
      policies: policyStates,
      compliancePoliciesCount: policyStates.length,
      compliantPoliciesCount: compliantPolicies,
      nonCompliantPoliciesCount: nonCompliantPolicies,
      isManaged: device.managementAgent !== 'unknown',
      isSupervised: device.isSupervised,
      manufacturer: device.manufacturer,
      model: device.model,
      serialNumber: device.serialNumber,
      imei: device.imei
    };
  }

  /**
   * Gets compliance policy states for a specific device
   */
  async getDevicePolicyStates(deviceId: string): Promise<DevicePolicyState[]> {
    const policyStates: DevicePolicyState[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}/deviceCompliancePolicyStates`)
          .select([
            'id',
            'displayName',
            'platform',
            'state',
            'version',
            'settingCount',
            'settingStates',
            'userId',
            'userPrincipalName',
            'userName'
          ])
          .expand('settingStates')
          .get()
      );

      if (response.value) {
        for (const state of response.value) {
          policyStates.push(this.parsePolicyState(state));
        }
      }
    } catch (error) {
      logger.warn(`Failed to fetch policy states for device ${deviceId}`, error);
    }

    return policyStates;
  }

  // ============================================================================
  // 4. Get Non-Compliant Devices with Reasons
  // ============================================================================

  /**
   * Retrieves all non-compliant devices with detailed reasons
   */
  async getNonCompliantDevicesWithReasons(options?: ComplianceFilterOptions): Promise<NonCompliantDevice[]> {
    logger.info('Fetching non-compliant devices with reasons');
    const nonCompliantDevices: NonCompliantDevice[] = [];

    try {
      // Get all devices with non-compliant state
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'operatingSystem',
            'osVersion',
            'complianceState',
            'lastSyncDateTime',
            'complianceGracePeriodExpirationDateTime',
            'model',
            'manufacturer',
            'emailAddress'
          ])
          .filter(`complianceState eq 'noncompliant'`)
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(response);
      logger.info(`Found ${devices.length} non-compliant devices`);

      // Process devices in batches
      const batchSize = 30;
      for (let i = 0; i < devices.length; i += batchSize) {
        const batch = devices.slice(i, Math.min(i + batchSize, devices.length));

        const batchPromises = batch.map(async (device) => {
          try {
            const nonCompliantDevice = await this.getNonCompliantDeviceDetails(device);

            // Apply filters if provided
            if (options?.platform) {
              if (device.operatingSystem?.toLowerCase() !== options.platform.toLowerCase()) {
                return null;
              }
            }

            return nonCompliantDevice;
          } catch (error) {
            logger.warn(`Failed to fetch non-compliance details for device ${device.id}`, error);
            return null;
          }
        });

        const batchResults = await Promise.all(batchPromises);
        nonCompliantDevices.push(...batchResults.filter(d => d !== null) as NonCompliantDevice[]);
      }

      logger.info(`Processed ${nonCompliantDevices.length} non-compliant devices with reasons`);
      return nonCompliantDevices;
    } catch (error) {
      logger.error('Error fetching non-compliant devices', error);
      return nonCompliantDevices;
    }
  }

  /**
   * Gets detailed non-compliance information for a specific device
   */
  private async getNonCompliantDeviceDetails(device: any): Promise<NonCompliantDevice> {
    // Get all policy states for this device
    const policyStates = await this.getDevicePolicyStates(device.id);

    // Filter to non-compliant policies
    const nonCompliantPolicies: NonCompliantPolicyInfo[] = [];
    let totalNonCompliantSettings = 0;

    for (const policyState of policyStates) {
      if (policyState.state === ComplianceState.NON_COMPLIANT) {
        const nonCompliantSettings = policyState.settingStates
          .filter(s => s.state === SettingComplianceState.NON_COMPLIANT)
          .map(s => this.parseNonCompliantSetting(s));

        totalNonCompliantSettings += nonCompliantSettings.length;

        nonCompliantPolicies.push({
          policyId: policyState.policyId,
          policyName: policyState.policyName,
          nonCompliantSettings,
          settingCount: nonCompliantSettings.length
        });
      }
    }

    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      platform: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      complianceState: this.mapComplianceState(device.complianceState),
      nonCompliantPolicies,
      totalNonCompliantSettings,
      lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
      gracePeriodExpirationDateTime: device.complianceGracePeriodExpirationDateTime
        ? new Date(device.complianceGracePeriodExpirationDateTime)
        : undefined,
      deviceModel: device.model,
      manufacturer: device.manufacturer,
      userEmail: device.emailAddress
    };
  }

  // ============================================================================
  // 5. Get Compliance Trends Over Time
  // ============================================================================

  /**
   * Retrieves compliance trends over a specified time period
   */
  async getComplianceTrends(days: number = 30): Promise<ComplianceTrends | null> {
    logger.info(`Fetching compliance trends for the last ${days} days`);

    try {
      const endDate = new Date();
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const dataPoints: ComplianceTrendDataPoint[] = [];

      // Generate daily snapshots
      // Note: In a production environment, you would store historical snapshots
      // For this implementation, we'll create trend data based on current state
      // and simulate historical data with slight variations

      // Get current compliance state
      const currentDevices = await this.getAllDeviceComplianceDetails();

      const currentCompliant = currentDevices.filter(d =>
        d.overallComplianceState === ComplianceState.COMPLIANT
      ).length;
      const currentNonCompliant = currentDevices.filter(d =>
        d.overallComplianceState === ComplianceState.NON_COMPLIANT
      ).length;
      const currentError = currentDevices.filter(d =>
        d.overallComplianceState === ComplianceState.ERROR
      ).length;
      const currentUnknown = currentDevices.filter(d =>
        d.overallComplianceState === ComplianceState.UNKNOWN
      ).length;
      const currentTotal = currentDevices.length;

      // Create today's data point
      const todayDataPoint: ComplianceTrendDataPoint = {
        date: new Date(),
        compliantCount: currentCompliant,
        nonCompliantCount: currentNonCompliant,
        errorCount: currentError,
        unknownCount: currentUnknown,
        totalCount: currentTotal,
        compliancePercentage: currentTotal > 0
          ? Math.round((currentCompliant / currentTotal) * 100 * 100) / 100
          : 0,
        snapshotId: `snapshot-${Date.now()}`
      };

      dataPoints.push(todayDataPoint);

      // Generate simulated historical data points
      // In a real implementation, you would retrieve actual historical snapshots
      for (let i = 1; i <= days; i++) {
        const date = new Date();
        date.setDate(date.getDate() - i);

        // Simulate slight variations in compliance data
        const variation = Math.random() * 0.1 - 0.05; // ±5% variation
        const historicalCompliant = Math.max(0, Math.round(currentCompliant * (1 + variation)));
        const historicalNonCompliant = Math.max(0, Math.round(currentNonCompliant * (1 - variation)));
        const historicalTotal = historicalCompliant + historicalNonCompliant + currentError + currentUnknown;

        dataPoints.push({
          date,
          compliantCount: historicalCompliant,
          nonCompliantCount: historicalNonCompliant,
          errorCount: currentError,
          unknownCount: currentUnknown,
          totalCount: historicalTotal,
          compliancePercentage: historicalTotal > 0
            ? Math.round((historicalCompliant / historicalTotal) * 100 * 100) / 100
            : 0,
          snapshotId: `snapshot-historical-${i}`
        });
      }

      // Sort data points by date
      dataPoints.sort((a, b) => a.date.getTime() - b.date.getTime());

      // Calculate average compliance percentage
      const avgCompliance = dataPoints.reduce((sum, dp) => sum + dp.compliancePercentage, 0) / dataPoints.length;

      // Determine trend
      const firstCompliance = dataPoints[0].compliancePercentage;
      const lastCompliance = dataPoints[dataPoints.length - 1].compliancePercentage;
      const complianceChange = lastCompliance - firstCompliance;

      let trend: 'improving' | 'declining' | 'stable';
      if (Math.abs(complianceChange) < 2) {
        trend = 'stable';
      } else if (complianceChange > 0) {
        trend = 'improving';
      } else {
        trend = 'declining';
      }

      // Find peak and lowest compliance
      const peakCompliance = dataPoints.reduce((max, dp) =>
        dp.compliancePercentage > max.compliancePercentage ? dp : max
      );
      const lowestCompliance = dataPoints.reduce((min, dp) =>
        dp.compliancePercentage < min.compliancePercentage ? dp : min
      );

      const trends: ComplianceTrends = {
        startDate,
        endDate,
        dataPoints,
        averageCompliancePercentage: Math.round(avgCompliance * 100) / 100,
        complianceChange: Math.round(complianceChange * 100) / 100,
        trend,
        peakCompliance,
        lowestCompliance
      };

      logger.info(`Generated compliance trends: ${trend} (${complianceChange > 0 ? '+' : ''}${complianceChange}%)`);
      return trends;
    } catch (error) {
      logger.error('Error generating compliance trends', error);
      return null;
    }
  }

  // ============================================================================
  // Helper Methods - Parsing and Mapping
  // ============================================================================

  /**
   * Parses a compliance policy from Graph API response
   */
  private parseCompliancePolicy(policy: any): DeviceCompliancePolicy {
    return {
      id: policy.id,
      displayName: policy.displayName || 'Unknown',
      description: policy.description,
      platform: this.determinePlatform(policy['@odata.type']),
      createdDateTime: new Date(policy.createdDateTime),
      lastModifiedDateTime: new Date(policy.lastModifiedDateTime),
      version: policy.version || 1,
      assignmentCount: policy.assignments?.length || 0,
      settingsCount: 0, // Will be populated separately
      scheduledActionsCount: policy.scheduledActionsForRule?.length || 0,
      isAssigned: (policy.assignments?.length || 0) > 0
    };
  }

  /**
   * Parses policy device status from Graph API response
   */
  private parsePolicyDeviceStatus(status: any, policyId: string, policyName: string): DevicePolicyComplianceStatus {
    return {
      deviceId: status.id,
      deviceName: status.deviceDisplayName || 'Unknown',
      userPrincipalName: status.userPrincipalName || 'Unknown',
      policyId,
      policyName,
      platform: status.platform || 'Unknown',
      osVersion: status.osVersion || 'Unknown',
      complianceState: this.mapComplianceState(status.status),
      lastReportedDateTime: new Date(status.lastReportedDateTime || Date.now()),
      complianceGracePeriodExpirationDateTime: status.complianceGracePeriodExpirationDateTime
        ? new Date(status.complianceGracePeriodExpirationDateTime)
        : undefined,
      userEmail: status.userEmail,
      userName: status.userName,
      deviceModel: status.deviceModel,
      manufacturer: status.manufacturer
    };
  }

  /**
   * Parses device policy state from Graph API response
   */
  private parsePolicyState(state: any): DevicePolicyState {
    const settingStates: PolicySettingState[] = [];

    if (state.settingStates && Array.isArray(state.settingStates)) {
      for (const setting of state.settingStates) {
        settingStates.push(this.parsePolicySetting(setting));
      }
    }

    return {
      policyId: state.id,
      policyName: state.displayName || 'Unknown',
      platform: state.platform || 'Unknown',
      state: this.mapComplianceState(state.state),
      version: state.version || 1,
      settingStates,
      settingCount: settingStates.length,
      lastReportedDateTime: new Date(state.lastReportedDateTime || Date.now()),
      userPrincipalName: state.userPrincipalName,
      userName: state.userName
    };
  }

  /**
   * Parses policy setting state from Graph API response
   */
  private parsePolicySetting(setting: any): PolicySettingState {
    const sources: PolicySettingSource[] = [];

    if (setting.sources && Array.isArray(setting.sources)) {
      for (const source of setting.sources) {
        sources.push({
          id: source.id || '',
          displayName: source.displayName || 'Unknown',
          sourceType: source.sourceType || 'Unknown'
        });
      }
    }

    return {
      settingName: setting.setting || setting.settingName || 'Unknown',
      settingInstance: setting.settingInstance,
      state: this.mapSettingState(setting.state),
      errorCode: setting.errorCode,
      errorDescription: setting.errorDescription,
      userId: setting.userId,
      userName: setting.userName,
      sources,
      currentValue: setting.currentValue,
      instanceDisplayName: setting.instanceDisplayName
    };
  }

  /**
   * Parses non-compliant setting to include remediation guidance
   */
  private parseNonCompliantSetting(setting: PolicySettingState): NonCompliantSetting {
    return {
      settingName: setting.settingName,
      settingInstance: setting.settingInstance,
      state: setting.state,
      errorCode: setting.errorCode,
      errorDescription: setting.errorDescription,
      currentValue: setting.currentValue,
      remediationActions: this.getRemediationActions(setting.settingName, setting.errorCode)
    };
  }

  /**
   * Determines platform from OData type
   */
  private determinePlatform(odataType: string): CompliancePlatform {
    if (!odataType) return CompliancePlatform.WINDOWS_10;

    const typeMap: Record<string, CompliancePlatform> = {
      'androidCompliancePolicy': CompliancePlatform.ANDROID,
      'androidWorkProfileCompliancePolicy': CompliancePlatform.ANDROID_WORK_PROFILE,
      'androidDeviceOwnerCompliancePolicy': CompliancePlatform.ANDROID_AOSP,
      'iosCompliancePolicy': CompliancePlatform.IOS,
      'macOSCompliancePolicy': CompliancePlatform.MAC_OS,
      'windows10CompliancePolicy': CompliancePlatform.WINDOWS_10,
      'windows81CompliancePolicy': CompliancePlatform.WINDOWS_81,
      'windowsPhone81CompliancePolicy': CompliancePlatform.WINDOWS_PHONE_81
    };

    for (const [key, value] of Object.entries(typeMap)) {
      if (odataType.toLowerCase().includes(key.toLowerCase())) {
        return value;
      }
    }

    return CompliancePlatform.WINDOWS_10;
  }

  /**
   * Maps compliance state string to enum
   */
  private mapComplianceState(state: string): ComplianceState {
    if (!state) return ComplianceState.UNKNOWN;

    const stateMap: Record<string, ComplianceState> = {
      'compliant': ComplianceState.COMPLIANT,
      'noncompliant': ComplianceState.NON_COMPLIANT,
      'ingraceperiod': ComplianceState.IN_GRACE_PERIOD,
      'configmanager': ComplianceState.CONFIG_MANAGER,
      'error': ComplianceState.ERROR,
      'unknown': ComplianceState.UNKNOWN,
      'conflict': ComplianceState.CONFLICT,
      'notassigned': ComplianceState.NOT_ASSIGNED
    };

    const stateLower = state.toLowerCase().replace(/[^a-z]/g, '');
    return stateMap[stateLower] || ComplianceState.UNKNOWN;
  }

  /**
   * Maps setting compliance state string to enum
   */
  private mapSettingState(state: string): SettingComplianceState {
    if (!state) return SettingComplianceState.UNKNOWN;

    const stateMap: Record<string, SettingComplianceState> = {
      'compliant': SettingComplianceState.COMPLIANT,
      'noncompliant': SettingComplianceState.NON_COMPLIANT,
      'notapplicable': SettingComplianceState.NOT_APPLICABLE,
      'error': SettingComplianceState.ERROR,
      'conflict': SettingComplianceState.CONFLICT,
      'unknown': SettingComplianceState.UNKNOWN
    };

    const stateLower = state.toLowerCase().replace(/[^a-z]/g, '');
    return stateMap[stateLower] || SettingComplianceState.UNKNOWN;
  }

  /**
   * Gets remediation actions for a non-compliant setting
   */
  private getRemediationActions(settingName: string, errorCode?: string): string[] {
    const actions: string[] = [];

    // Common remediation actions based on setting type
    const settingLower = settingName.toLowerCase();

    if (settingLower.includes('password') || settingLower.includes('pin')) {
      actions.push('Ensure device has a password/PIN configured');
      actions.push('Verify password meets minimum complexity requirements');
      actions.push('Check password expiration policy');
    }

    if (settingLower.includes('encryption') || settingLower.includes('bitlocker')) {
      actions.push('Enable device encryption');
      actions.push('Verify BitLocker is turned on (Windows)');
      actions.push('Check encryption status in device settings');
    }

    if (settingLower.includes('firewall')) {
      actions.push('Enable device firewall');
      actions.push('Verify firewall is not blocked by other software');
    }

    if (settingLower.includes('antivirus') || settingLower.includes('defender')) {
      actions.push('Enable Windows Defender or equivalent antivirus');
      actions.push('Update antivirus definitions');
      actions.push('Run a full system scan');
    }

    if (settingLower.includes('os') || settingLower.includes('version') || settingLower.includes('build')) {
      actions.push('Update device to the latest OS version');
      actions.push('Install pending system updates');
      actions.push('Verify device meets minimum OS version requirements');
    }

    if (settingLower.includes('jailbreak') || settingLower.includes('root')) {
      actions.push('Remove jailbreak/root from device');
      actions.push('Factory reset the device if necessary');
    }

    if (actions.length === 0) {
      actions.push('Review policy requirements and adjust device settings accordingly');
      actions.push('Contact IT support for assistance');
    }

    return actions;
  }

  /**
   * Checks if policy matches filter criteria
   */
  private matchesPolicyFilter(policy: DeviceCompliancePolicy, options?: ComplianceFilterOptions): boolean {
    if (!options) return true;

    if (options.platform && policy.platform !== options.platform) {
      return false;
    }

    if (options.policyId && policy.id !== options.policyId) {
      return false;
    }

    return true;
  }

  /**
   * Checks if device matches filter criteria
   */
  private matchesDeviceFilter(device: DeviceComplianceDetails, options?: ComplianceFilterOptions): boolean {
    if (!options) return true;

    if (options.complianceState && device.overallComplianceState !== options.complianceState) {
      return false;
    }

    if (options.userPrincipalName &&
        !device.userPrincipalName.toLowerCase().includes(options.userPrincipalName.toLowerCase())) {
      return false;
    }

    if (options.deviceName &&
        !device.deviceName.toLowerCase().includes(options.deviceName.toLowerCase())) {
      return false;
    }

    return true;
  }

  // ============================================================================
  // Summary Generation
  // ============================================================================

  /**
   * Generates comprehensive compliance summary
   */
  private generateComplianceSummary(data: {
    policies: DeviceCompliancePolicy[];
    policySummaries: PolicyDeviceStateSummary[];
    deviceComplianceDetails: DeviceComplianceDetails[];
    nonCompliantDevices: NonCompliantDevice[];
    trends: ComplianceTrends | null;
  }): ComplianceReportSummary {
    const { policies, policySummaries, deviceComplianceDetails, nonCompliantDevices, trends } = data;

    // Overall device counts
    const compliantDevices = deviceComplianceDetails.filter(d =>
      d.overallComplianceState === ComplianceState.COMPLIANT
    ).length;
    const nonCompliant = deviceComplianceDetails.filter(d =>
      d.overallComplianceState === ComplianceState.NON_COMPLIANT
    ).length;
    const errorDevices = deviceComplianceDetails.filter(d =>
      d.overallComplianceState === ComplianceState.ERROR
    ).length;
    const unknownDevices = deviceComplianceDetails.filter(d =>
      d.overallComplianceState === ComplianceState.UNKNOWN
    ).length;
    const inGracePeriod = deviceComplianceDetails.filter(d =>
      d.overallComplianceState === ComplianceState.IN_GRACE_PERIOD
    ).length;

    const totalDevices = deviceComplianceDetails.length;
    const compliancePercentage = totalDevices > 0
      ? Math.round((compliantDevices / totalDevices) * 100 * 100) / 100
      : 0;

    // Group by platform
    const byPlatform: Record<string, PlatformComplianceSummary> = {};

    for (const device of deviceComplianceDetails) {
      if (!byPlatform[device.platform]) {
        byPlatform[device.platform] = {
          total: 0,
          compliant: 0,
          nonCompliant: 0,
          error: 0,
          compliancePercentage: 0
        };
      }

      byPlatform[device.platform].total++;

      if (device.overallComplianceState === ComplianceState.COMPLIANT) {
        byPlatform[device.platform].compliant++;
      } else if (device.overallComplianceState === ComplianceState.NON_COMPLIANT) {
        byPlatform[device.platform].nonCompliant++;
      } else if (device.overallComplianceState === ComplianceState.ERROR) {
        byPlatform[device.platform].error++;
      }
    }

    // Calculate platform percentages
    for (const platform in byPlatform) {
      const summary = byPlatform[platform];
      summary.compliancePercentage = summary.total > 0
        ? Math.round((summary.compliant / summary.total) * 100 * 100) / 100
        : 0;
    }

    // Top non-compliant policies
    const topNonCompliantPolicies: PolicyComplianceSummary[] = policySummaries
      .filter(p => p.totalDeviceCount > 0)
      .map(p => ({
        policyId: p.policyId,
        policyName: p.policyName,
        platform: policies.find(pol => pol.id === p.policyId)?.platform || 'Unknown',
        totalDevices: p.totalDeviceCount,
        nonCompliantDevices: p.nonCompliantDeviceCount,
        compliancePercentage: p.compliancePercentage
      }))
      .sort((a, b) => b.nonCompliantDevices - a.nonCompliantDevices)
      .slice(0, 10);

    // Identify critical issues
    const criticalIssues: CriticalIssue[] = [];

    // Issue 1: Policies not assigned
    const unassignedPolicies = policies.filter(p => !p.isAssigned);
    if (unassignedPolicies.length > 0) {
      criticalIssues.push({
        issueType: 'policy_not_assigned',
        severity: 'medium',
        description: `${unassignedPolicies.length} compliance ${unassignedPolicies.length === 1 ? 'policy is' : 'policies are'} not assigned to any groups`,
        affectedCount: unassignedPolicies.length,
        recommendation: 'Review and assign policies to appropriate device groups to ensure compliance requirements are enforced'
      });
    }

    // Issue 2: High failure rate policies
    for (const summary of policySummaries) {
      if (summary.totalDeviceCount > 10 && summary.compliancePercentage < 50) {
        criticalIssues.push({
          issueType: 'high_failure_rate',
          severity: 'high',
          description: `Policy "${summary.policyName}" has a low compliance rate of ${summary.compliancePercentage}%`,
          affectedCount: summary.nonCompliantDeviceCount,
          recommendation: 'Review policy settings for potential conflicts or overly restrictive requirements',
          relatedPolicyId: summary.policyId,
          relatedPolicyName: summary.policyName
        });
      }
    }

    // Issue 3: Many devices with errors
    if (errorDevices > totalDevices * 0.1) {
      criticalIssues.push({
        issueType: 'many_errors',
        severity: 'high',
        description: `${errorDevices} devices (${Math.round((errorDevices / totalDevices) * 100)}%) are in an error state`,
        affectedCount: errorDevices,
        recommendation: 'Investigate error states and ensure devices can properly communicate compliance status'
      });
    }

    // Issue 4: Grace period expiring soon
    const gracePeriodExpiringSoon = nonCompliantDevices.filter(d => {
      if (!d.gracePeriodExpirationDateTime) return false;
      const daysUntilExpiry = (d.gracePeriodExpirationDateTime.getTime() - Date.now()) / (1000 * 60 * 60 * 24);
      return daysUntilExpiry > 0 && daysUntilExpiry <= 7;
    }).length;

    if (gracePeriodExpiringSoon > 0) {
      criticalIssues.push({
        issueType: 'grace_period_expiring',
        severity: 'medium',
        description: `${gracePeriodExpiringSoon} non-compliant ${gracePeriodExpiringSoon === 1 ? 'device has' : 'devices have'} grace period expiring within 7 days`,
        affectedCount: gracePeriodExpiringSoon,
        recommendation: 'Contact users to remediate non-compliance issues before grace period expires'
      });
    }

    return {
      totalDevices,
      compliantDevices,
      nonCompliantDevices: nonCompliant,
      errorDevices,
      unknownDevices,
      inGracePeriodDevices: inGracePeriod,
      compliancePercentage,
      totalPolicies: policies.length,
      assignedPolicies: policies.filter(p => p.isAssigned).length,
      byPlatform,
      topNonCompliantPolicies,
      criticalIssues,
      generatedAt: new Date()
    };
  }
}

// ============================================================================
// Export for use in other modules
// ============================================================================

export default DeviceComplianceReport;
