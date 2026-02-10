/**
 * Exchange ActiveSync Reporting Module for Microsoft Intune
 *
 * This module replicates Configuration Manager Exchange ActiveSync reports for Intune:
 * - Report 8: Compliance status of default ActiveSync mailbox policy
 * - Report 15: Inactive mobile devices (Exchange)
 * - Report 21: Mobile device compliance details (Exchange)
 * - Report 34: Settings summary for mobile devices (Exchange)
 *
 * Uses Microsoft Graph API endpoints:
 * - GET /deviceManagement/managedDevices
 * - GET /deviceManagement/deviceCompliancePolicies
 * - GET /deviceManagement/deviceCompliancePolicies/{id}/deviceStatuses
 *
 * @module exchange-activesync-reports
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
 * Exchange access state enumeration
 */
export enum ExchangeAccessState {
  ALLOWED = 'allowed',
  BLOCKED = 'blocked',
  QUARANTINED = 'quarantined',
  UNKNOWN = 'unknown',
  NONE = 'none'
}

/**
 * Exchange access state reason enumeration
 */
export enum ExchangeAccessStateReason {
  NONE = 'none',
  UNKNOWN = 'unknown',
  EXCHANGE_GLOBAL_RULE = 'exchangeGlobalRule',
  EXCHANGE_INDIVIDUAL_RULE = 'exchangeIndividualRule',
  EXCHANGE_DEVICE_RULE = 'exchangeDeviceRule',
  EXCHANGE_UPGRADE = 'exchangeUpgrade',
  EXCHANGE_MAILBOX_POLICY = 'exchangeMailboxPolicy',
  OTHER = 'other',
  COMPLIANT = 'compliant',
  NOT_COMPLIANT = 'notCompliant',
  NOT_ENROLLED = 'notEnrolled',
  UNKNOWN_LOCATION = 'unknownLocation',
  MFA_REQUIRED = 'mfaRequired',
  AZURE_AD_BLOCK_DUE_TO_ACCESS_POLICY = 'azureADBlockDueToAccessPolicy',
  COMPROMISED_PASSWORD = 'compromisedPassword',
  DEVICE_NOT_KNOWN_WITH_MANAGED_APP = 'deviceNotKnownWithManagedApp'
}

/**
 * Exchange device type enumeration
 */
export enum ExchangeDeviceType {
  SMART_PHONE = 'smartphone',
  TABLET = 'tablet',
  UNKNOWN = 'unknown'
}

/**
 * ActiveSync mailbox policy compliance summary
 */
export interface ActiveSyncPolicyComplianceSummary {
  policyId: string;
  policyName: string;
  totalDevices: number;
  allowedDevices: number;
  blockedDevices: number;
  quarantinedDevices: number;
  unknownDevices: number;
  compliancePercentage: number;
  lastUpdated: Date;
  devicesByPlatform: Record<string, number>;
  devicesByAccessState: Record<ExchangeAccessState, number>;
}

/**
 * Inactive mobile device information
 */
export interface InactiveMobileDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  emailAddress: string;
  operatingSystem: string;
  osVersion: string;
  model: string;
  manufacturer: string;
  lastSyncDateTime: Date;
  daysSinceLastSync: number;
  exchangeAccessState: ExchangeAccessState;
  exchangeAccessStateReason: ExchangeAccessStateReason;
  exchangeLastSuccessfulSyncDateTime?: Date;
  enrolledDateTime?: Date;
  complianceState: string;
  managementAgent: string;
  isSupervised?: boolean;
  serialNumber?: string;
  imei?: string;
}

/**
 * Mobile device compliance details for Exchange
 */
export interface MobileDeviceComplianceDetails {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  emailAddress: string;
  operatingSystem: string;
  osVersion: string;
  model: string;
  manufacturer: string;
  exchangeAccessState: ExchangeAccessState;
  exchangeAccessStateReason: ExchangeAccessStateReason;
  exchangeLastSuccessfulSyncDateTime?: Date;
  complianceState: string;
  lastSyncDateTime: Date;
  enrolledDateTime?: Date;
  activeSyncId?: string;
  easDeviceId?: string;
  easActivated: boolean;
  easActivationDateTime?: Date;
  mailboxPolicies: MailboxPolicyAssignment[];
  compliancePolicies: CompliancePolicyStatus[];
  deviceActions: DeviceAction[];
  isManaged: boolean;
  isSupervised?: boolean;
  jailBroken?: string;
}

/**
 * Mailbox policy assignment
 */
export interface MailboxPolicyAssignment {
  policyId: string;
  policyName: string;
  assignedDateTime: Date;
  isDefault: boolean;
}

/**
 * Compliance policy status for a device
 */
export interface CompliancePolicyStatus {
  policyId: string;
  policyName: string;
  complianceState: string;
  lastReportedDateTime: Date;
  settingStates: PolicySettingStatus[];
}

/**
 * Policy setting status
 */
export interface PolicySettingStatus {
  settingName: string;
  state: string;
  errorCode?: string;
  errorDescription?: string;
  currentValue?: string;
}

/**
 * Device action history
 */
export interface DeviceAction {
  actionName: string;
  actionState: string;
  startDateTime: Date;
  lastUpdatedDateTime: Date;
  initiatedBy?: string;
}

/**
 * Settings summary for mobile devices
 */
export interface MobileDeviceSettingsSummary {
  settingName: string;
  settingDescription: string;
  settingCategory: 'exchange' | 'compliance' | 'configuration';
  deviceCount: number;
  enabledCount: number;
  disabledCount: number;
  notConfiguredCount: number;
  percentage: number;
  affectedPolicies: string[];
  platforms: Record<string, number>;
}

/**
 * Exchange ActiveSync report summary
 */
export interface ExchangeActiveSyncReportSummary {
  totalMobileDevices: number;
  activeDevices: number;
  inactiveDevices: number;
  allowedDevices: number;
  blockedDevices: number;
  quarantinedDevices: number;
  easActivatedDevices: number;
  compliancePercentage: number;
  averageDaysSinceLastSync: number;
  devicesByPlatform: Record<string, number>;
  devicesByAccessState: Record<string, number>;
  devicesByAccessStateReason: Record<string, number>;
  topInactiveThreshold: number;
  generatedAt: Date;
}

/**
 * Filter options for Exchange ActiveSync queries
 */
export interface ExchangeActiveSyncFilterOptions {
  platform?: string;
  exchangeAccessState?: ExchangeAccessState;
  inactiveDaysThreshold?: number;
  userPrincipalName?: string;
  includeDetails?: boolean;
  complianceState?: string;
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Exchange ActiveSync Reports
 *
 * Provides comprehensive Exchange ActiveSync reporting capabilities:
 * 1. Compliance status of default ActiveSync mailbox policy
 * 2. Inactive mobile devices (Exchange)
 * 3. Mobile device compliance details (Exchange)
 * 4. Settings summary for mobile devices (Exchange)
 */
export class ExchangeActiveSyncReports extends BaseReport {
  name = 'exchange-activesync-reports';
  description = 'Exchange ActiveSync reporting including policy compliance, inactive devices, device details, and settings summary';
  category = 'Exchange ActiveSync';
  enabled = true;

  /**
   * Execute the complete Exchange ActiveSync report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Exchange ActiveSync Reports');

    try {
      const startTime = Date.now();

      // Fetch all Exchange ActiveSync data
      const [
        policyCompliance,
        inactiveDevices,
        deviceDetails,
        settingsSummary
      ] = await Promise.all([
        this.getActiveSyncPolicyCompliance(),
        this.getInactiveMobileDevices(30), // Default: 30 days
        this.getAllMobileDeviceComplianceDetails(),
        this.getMobileDeviceSettingsSummary()
      ]);

      // Generate summary
      const summary = this.generateExchangeActiveSyncSummary({
        policyCompliance,
        inactiveDevices,
        deviceDetails,
        settingsSummary
      });

      const duration = Date.now() - startTime;

      logger.info('Exchange ActiveSync Reports completed', {
        duration: `${duration}ms`,
        totalDevices: deviceDetails.length,
        inactiveDevices: inactiveDevices.length
      });

      // Combine all data for the report
      const data = [
        ...policyCompliance.map(p => ({ type: 'policy-compliance', ...p })),
        ...inactiveDevices.map(d => ({ type: 'inactive-device', ...d })),
        ...deviceDetails.map(d => ({ type: 'device-details', ...d })),
        ...settingsSummary.map(s => ({ type: 'settings-summary', ...s }))
      ];

      return {
        metadata: this.createMetadata(this.name, data.length, {
          duration: `${duration}ms`,
          inactiveDaysThreshold: 30
        }),
        data,
        summary
      };
    } catch (error) {
      logger.error('Failed to execute Exchange ActiveSync Reports', error);
      throw error;
    }
  }

  // ============================================================================
  // Report 8: Compliance Status of Default ActiveSync Mailbox Policy
  // ============================================================================

  /**
   * Gets compliance status summary for ActiveSync mailbox policies
   */
  async getActiveSyncPolicyCompliance(): Promise<ActiveSyncPolicyComplianceSummary[]> {
    logger.info('Fetching ActiveSync policy compliance status');
    const summaries: ActiveSyncPolicyComplianceSummary[] = [];

    try {
      // Get all managed devices with Exchange data
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'operatingSystem',
            'exchangeAccessState',
            'exchangeAccessStateReason',
            'easActivated',
            'complianceState'
          ])
          .filter(`managementAgent eq 'eas' or managementAgent eq 'easMdm' or easActivated eq true`)
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(response);
      logger.info(`Processing ${devices.length} Exchange ActiveSync devices`);

      // Get compliance policies
      const policiesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceCompliancePolicies')
          .select(['id', 'displayName'])
          .top(999)
          .get()
      );

      const policies = await this.getAllPages(policiesResponse);

      // Process each policy
      for (const policy of policies) {
        try {
          const summary = await this.getPolicyComplianceSummary(policy, devices);
          summaries.push(summary);
        } catch (error) {
          logger.warn(`Failed to get compliance summary for policy ${policy.id}`, error);
        }
      }

      // Add a summary for devices without specific policy assignment
      const unassignedSummary = this.getUnassignedDevicesSummary(devices);
      if (unassignedSummary.totalDevices > 0) {
        summaries.push(unassignedSummary);
      }

      logger.info(`Generated ${summaries.length} policy compliance summaries`);
      return summaries;
    } catch (error) {
      logger.error('Error fetching ActiveSync policy compliance', error);
      return summaries;
    }
  }

  /**
   * Gets compliance summary for a specific policy
   */
  private async getPolicyComplianceSummary(
    policy: any,
    allDevices: any[]
  ): Promise<ActiveSyncPolicyComplianceSummary> {
    // Get device statuses for this policy
    const statusResponse = await this.retryGraphCall(() =>
      this.graphClient
        .api(`/deviceManagement/deviceCompliancePolicies/${policy.id}/deviceStatuses`)
        .select(['id', 'status', 'platform'])
        .top(999)
        .get()
    );

    const deviceStatuses = await this.getAllPages(statusResponse);

    // Aggregate statistics
    const devicesByAccessState: Record<ExchangeAccessState, number> = {
      [ExchangeAccessState.ALLOWED]: 0,
      [ExchangeAccessState.BLOCKED]: 0,
      [ExchangeAccessState.QUARANTINED]: 0,
      [ExchangeAccessState.UNKNOWN]: 0,
      [ExchangeAccessState.NONE]: 0
    };

    const devicesByPlatform: Record<string, number> = {};

    // Match devices with statuses
    const matchedDevices = allDevices.filter(device =>
      deviceStatuses.some(status => status.id === device.id)
    );

    for (const device of matchedDevices) {
      const accessState = this.mapExchangeAccessState(device.exchangeAccessState);
      devicesByAccessState[accessState]++;

      const platform = device.operatingSystem || 'Unknown';
      devicesByPlatform[platform] = (devicesByPlatform[platform] || 0) + 1;
    }

    const totalDevices = matchedDevices.length;
    const allowedDevices = devicesByAccessState[ExchangeAccessState.ALLOWED];
    const compliancePercentage = totalDevices > 0
      ? Math.round((allowedDevices / totalDevices) * 100 * 100) / 100
      : 0;

    return {
      policyId: policy.id,
      policyName: policy.displayName || 'Unknown',
      totalDevices,
      allowedDevices,
      blockedDevices: devicesByAccessState[ExchangeAccessState.BLOCKED],
      quarantinedDevices: devicesByAccessState[ExchangeAccessState.QUARANTINED],
      unknownDevices: devicesByAccessState[ExchangeAccessState.UNKNOWN] +
                      devicesByAccessState[ExchangeAccessState.NONE],
      compliancePercentage,
      lastUpdated: new Date(),
      devicesByPlatform,
      devicesByAccessState
    };
  }

  /**
   * Gets summary for devices without policy assignment
   */
  private getUnassignedDevicesSummary(devices: any[]): ActiveSyncPolicyComplianceSummary {
    const devicesByAccessState: Record<ExchangeAccessState, number> = {
      [ExchangeAccessState.ALLOWED]: 0,
      [ExchangeAccessState.BLOCKED]: 0,
      [ExchangeAccessState.QUARANTINED]: 0,
      [ExchangeAccessState.UNKNOWN]: 0,
      [ExchangeAccessState.NONE]: 0
    };

    const devicesByPlatform: Record<string, number> = {};

    for (const device of devices) {
      const accessState = this.mapExchangeAccessState(device.exchangeAccessState);
      devicesByAccessState[accessState]++;

      const platform = device.operatingSystem || 'Unknown';
      devicesByPlatform[platform] = (devicesByPlatform[platform] || 0) + 1;
    }

    const totalDevices = devices.length;
    const allowedDevices = devicesByAccessState[ExchangeAccessState.ALLOWED];
    const compliancePercentage = totalDevices > 0
      ? Math.round((allowedDevices / totalDevices) * 100 * 100) / 100
      : 0;

    return {
      policyId: 'unassigned',
      policyName: 'Default Exchange ActiveSync Policy (Unassigned)',
      totalDevices,
      allowedDevices,
      blockedDevices: devicesByAccessState[ExchangeAccessState.BLOCKED],
      quarantinedDevices: devicesByAccessState[ExchangeAccessState.QUARANTINED],
      unknownDevices: devicesByAccessState[ExchangeAccessState.UNKNOWN] +
                      devicesByAccessState[ExchangeAccessState.NONE],
      compliancePercentage,
      lastUpdated: new Date(),
      devicesByPlatform,
      devicesByAccessState
    };
  }

  // ============================================================================
  // Report 15: Inactive Mobile Devices (Exchange)
  // ============================================================================

  /**
   * Gets inactive mobile devices that haven't synced within specified days
   */
  async getInactiveMobileDevices(
    inactiveDaysThreshold: number = 30,
    options?: ExchangeActiveSyncFilterOptions
  ): Promise<InactiveMobileDevice[]> {
    logger.info(`Fetching inactive mobile devices (threshold: ${inactiveDaysThreshold} days)`);
    const inactiveDevices: InactiveMobileDevice[] = [];

    try {
      // Get all managed mobile devices
      const response = await this.retryGraphCall(() =>
        this.graphClient
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
            'exchangeAccessState',
            'exchangeAccessStateReason',
            'exchangeLastSuccessfulSyncDateTime',
            'enrolledDateTime',
            'complianceState',
            'managementAgent',
            'isSupervised',
            'serialNumber',
            'imei'
          ])
          .filter(`managementAgent eq 'eas' or managementAgent eq 'easMdm' or managementAgent eq 'mdm'`)
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(response);
      logger.info(`Processing ${devices.length} devices for inactivity check`);

      const now = Date.now();
      const thresholdMs = inactiveDaysThreshold * 24 * 60 * 60 * 1000;

      for (const device of devices) {
        const lastSyncDate = device.lastSyncDateTime
          ? new Date(device.lastSyncDateTime)
          : device.exchangeLastSuccessfulSyncDateTime
            ? new Date(device.exchangeLastSuccessfulSyncDateTime)
            : null;

        if (!lastSyncDate) {
          // Device has never synced
          const daysSinceEnrollment = device.enrolledDateTime
            ? Math.floor((now - new Date(device.enrolledDateTime).getTime()) / (1000 * 60 * 60 * 24))
            : 9999;

          if (daysSinceEnrollment >= inactiveDaysThreshold) {
            inactiveDevices.push(this.mapToInactiveDevice(device, daysSinceEnrollment));
          }
        } else {
          const msSinceLastSync = now - lastSyncDate.getTime();
          const daysSinceLastSync = Math.floor(msSinceLastSync / (1000 * 60 * 60 * 24));

          if (msSinceLastSync >= thresholdMs) {
            // Apply additional filters if provided
            if (this.matchesInactiveDeviceFilter(device, options)) {
              inactiveDevices.push(this.mapToInactiveDevice(device, daysSinceLastSync));
            }
          }
        }
      }

      // Sort by days since last sync (descending)
      inactiveDevices.sort((a, b) => b.daysSinceLastSync - a.daysSinceLastSync);

      logger.info(`Found ${inactiveDevices.length} inactive devices`);
      return inactiveDevices;
    } catch (error) {
      logger.error('Error fetching inactive mobile devices', error);
      return inactiveDevices;
    }
  }

  /**
   * Maps device to inactive device structure
   */
  private mapToInactiveDevice(device: any, daysSinceLastSync: number): InactiveMobileDevice {
    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      emailAddress: device.emailAddress || device.userPrincipalName || 'Unknown',
      operatingSystem: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      model: device.model || 'Unknown',
      manufacturer: device.manufacturer || 'Unknown',
      lastSyncDateTime: device.lastSyncDateTime
        ? new Date(device.lastSyncDateTime)
        : device.exchangeLastSuccessfulSyncDateTime
          ? new Date(device.exchangeLastSuccessfulSyncDateTime)
          : new Date(0),
      daysSinceLastSync,
      exchangeAccessState: this.mapExchangeAccessState(device.exchangeAccessState),
      exchangeAccessStateReason: this.mapExchangeAccessStateReason(device.exchangeAccessStateReason),
      exchangeLastSuccessfulSyncDateTime: device.exchangeLastSuccessfulSyncDateTime
        ? new Date(device.exchangeLastSuccessfulSyncDateTime)
        : undefined,
      enrolledDateTime: device.enrolledDateTime ? new Date(device.enrolledDateTime) : undefined,
      complianceState: device.complianceState || 'unknown',
      managementAgent: device.managementAgent || 'unknown',
      isSupervised: device.isSupervised,
      serialNumber: device.serialNumber,
      imei: device.imei
    };
  }

  // ============================================================================
  // Report 21: Mobile Device Compliance Details (Exchange)
  // ============================================================================

  /**
   * Gets detailed compliance information for all mobile devices
   */
  async getAllMobileDeviceComplianceDetails(
    options?: ExchangeActiveSyncFilterOptions
  ): Promise<MobileDeviceComplianceDetails[]> {
    logger.info('Fetching mobile device compliance details');
    const deviceDetails: MobileDeviceComplianceDetails[] = [];

    try {
      // Get all managed mobile devices
      const response = await this.retryGraphCall(() =>
        this.graphClient
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
            'exchangeAccessState',
            'exchangeAccessStateReason',
            'exchangeLastSuccessfulSyncDateTime',
            'complianceState',
            'lastSyncDateTime',
            'enrolledDateTime',
            'activeSyncId',
            'easDeviceId',
            'easActivated',
            'easActivationDateTime',
            'managementAgent',
            'isSupervised',
            'jailBroken'
          ])
          .filter(`managementAgent eq 'eas' or managementAgent eq 'easMdm' or managementAgent eq 'mdm'`)
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(response);
      logger.info(`Processing compliance details for ${devices.length} devices`);

      // Process devices in batches
      const batchSize = 30;
      for (let i = 0; i < devices.length; i += batchSize) {
        const batch = devices.slice(i, Math.min(i + batchSize, devices.length));

        const batchPromises = batch.map(async (device) => {
          try {
            const details = await this.getDeviceComplianceDetails(device);

            // Apply filters if provided
            if (this.matchesComplianceDetailsFilter(details, options)) {
              return details;
            }
          } catch (error) {
            logger.warn(`Failed to fetch compliance details for device ${device.id}`, error);
          }
          return null;
        });

        const batchResults = await Promise.all(batchPromises);
        deviceDetails.push(...batchResults.filter(d => d !== null) as MobileDeviceComplianceDetails[]);
      }

      logger.info(`Fetched compliance details for ${deviceDetails.length} devices`);
      return deviceDetails;
    } catch (error) {
      logger.error('Error fetching mobile device compliance details', error);
      return deviceDetails;
    }
  }

  /**
   * Gets detailed compliance information for a specific device
   */
  private async getDeviceComplianceDetails(device: any): Promise<MobileDeviceComplianceDetails> {
    // Get mailbox policies (simulated - not directly available in Graph API)
    const mailboxPolicies: MailboxPolicyAssignment[] = [];

    // Get compliance policies
    const compliancePolicies = await this.getDeviceCompliancePolicies(device.id);

    // Get device actions
    const deviceActions = await this.getDeviceActions(device.id);

    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      emailAddress: device.emailAddress || device.userPrincipalName || 'Unknown',
      operatingSystem: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      model: device.model || 'Unknown',
      manufacturer: device.manufacturer || 'Unknown',
      exchangeAccessState: this.mapExchangeAccessState(device.exchangeAccessState),
      exchangeAccessStateReason: this.mapExchangeAccessStateReason(device.exchangeAccessStateReason),
      exchangeLastSuccessfulSyncDateTime: device.exchangeLastSuccessfulSyncDateTime
        ? new Date(device.exchangeLastSuccessfulSyncDateTime)
        : undefined,
      complianceState: device.complianceState || 'unknown',
      lastSyncDateTime: new Date(device.lastSyncDateTime || Date.now()),
      enrolledDateTime: device.enrolledDateTime ? new Date(device.enrolledDateTime) : undefined,
      activeSyncId: device.activeSyncId,
      easDeviceId: device.easDeviceId,
      easActivated: device.easActivated || false,
      easActivationDateTime: device.easActivationDateTime
        ? new Date(device.easActivationDateTime)
        : undefined,
      mailboxPolicies,
      compliancePolicies,
      deviceActions,
      isManaged: device.managementAgent !== 'unknown',
      isSupervised: device.isSupervised,
      jailBroken: device.jailBroken
    };
  }

  /**
   * Gets compliance policies for a device
   */
  private async getDeviceCompliancePolicies(deviceId: string): Promise<CompliancePolicyStatus[]> {
    const policies: CompliancePolicyStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}/deviceCompliancePolicyStates`)
          .select([
            'id',
            'displayName',
            'state',
            'lastReportedDateTime',
            'settingStates'
          ])
          .expand('settingStates')
          .get()
      );

      if (response.value) {
        for (const policy of response.value) {
          const settingStates: PolicySettingStatus[] = [];

          if (policy.settingStates && Array.isArray(policy.settingStates)) {
            for (const setting of policy.settingStates) {
              settingStates.push({
                settingName: setting.setting || setting.settingName || 'Unknown',
                state: setting.state || 'unknown',
                errorCode: setting.errorCode,
                errorDescription: setting.errorDescription,
                currentValue: setting.currentValue
              });
            }
          }

          policies.push({
            policyId: policy.id,
            policyName: policy.displayName || 'Unknown',
            complianceState: policy.state || 'unknown',
            lastReportedDateTime: new Date(policy.lastReportedDateTime || Date.now()),
            settingStates
          });
        }
      }
    } catch (error) {
      logger.warn(`Failed to fetch compliance policies for device ${deviceId}`, error);
    }

    return policies;
  }

  /**
   * Gets device actions for a device
   */
  private async getDeviceActions(deviceId: string): Promise<DeviceAction[]> {
    const actions: DeviceAction[] = [];

    try {
      // Note: This endpoint may not be available in all Graph API versions
      // Using device compliance actions as a proxy
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}`)
          .select(['deviceActionResults'])
          .get()
      );

      if (response.deviceActionResults && Array.isArray(response.deviceActionResults)) {
        for (const action of response.deviceActionResults) {
          actions.push({
            actionName: action.actionName || 'Unknown',
            actionState: action.actionState || 'unknown',
            startDateTime: new Date(action.startDateTime || Date.now()),
            lastUpdatedDateTime: new Date(action.lastUpdatedDateTime || Date.now()),
            initiatedBy: action.initiatedBy
          });
        }
      }
    } catch (error) {
      // Device actions endpoint may not be available - this is not critical
      logger.debug(`Device actions not available for device ${deviceId}`);
    }

    return actions;
  }

  // ============================================================================
  // Report 34: Settings Summary for Mobile Devices (Exchange)
  // ============================================================================

  /**
   * Gets aggregated settings summary across mobile devices
   */
  async getMobileDeviceSettingsSummary(): Promise<MobileDeviceSettingsSummary[]> {
    logger.info('Fetching mobile device settings summary');
    const summaries: MobileDeviceSettingsSummary[] = [];

    try {
      // Get all compliance policies
      const policiesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceCompliancePolicies')
          .select(['id', 'displayName'])
          .top(999)
          .get()
      );

      const policies = await this.getAllPages(policiesResponse);

      // Track settings across all policies and devices
      const settingsMap = new Map<string, {
        description: string;
        category: 'exchange' | 'compliance' | 'configuration';
        deviceCount: number;
        enabledCount: number;
        disabledCount: number;
        notConfiguredCount: number;
        policies: Set<string>;
        platforms: Record<string, number>;
      }>();

      // Process each policy
      for (const policy of policies) {
        try {
          const deviceStatuses = await this.retryGraphCall(() =>
            this.graphClient
              .api(`/deviceManagement/deviceCompliancePolicies/${policy.id}/deviceStatuses`)
              .select(['id', 'status', 'platform'])
              .top(999)
              .get()
          );

          const statuses = await this.getAllPages(deviceStatuses);

          // Get detailed settings for a sample of devices
          const sampleSize = Math.min(10, statuses.length);
          for (let i = 0; i < sampleSize; i++) {
            const deviceId = statuses[i].id;
            const platform = statuses[i].platform || 'Unknown';

            try {
              const policyStates = await this.retryGraphCall(() =>
                this.graphClient
                  .api(`/deviceManagement/managedDevices/${deviceId}/deviceCompliancePolicyStates`)
                  .filter(`id eq '${policy.id}'`)
                  .select(['settingStates'])
                  .expand('settingStates')
                  .get()
              );

              if (policyStates.value && policyStates.value[0]?.settingStates) {
                for (const setting of policyStates.value[0].settingStates) {
                  const settingName = setting.setting || setting.settingName || 'Unknown';

                  if (!settingsMap.has(settingName)) {
                    settingsMap.set(settingName, {
                      description: this.getSettingDescription(settingName),
                      category: this.categorizeSettings(settingName),
                      deviceCount: 0,
                      enabledCount: 0,
                      disabledCount: 0,
                      notConfiguredCount: 0,
                      policies: new Set(),
                      platforms: {}
                    });
                  }

                  const settingData = settingsMap.get(settingName)!;
                  settingData.deviceCount++;
                  settingData.policies.add(policy.displayName);
                  settingData.platforms[platform] = (settingData.platforms[platform] || 0) + 1;

                  // Categorize setting state
                  const state = (setting.state || '').toLowerCase();
                  if (state === 'compliant' || state.includes('enabled') || state.includes('true')) {
                    settingData.enabledCount++;
                  } else if (state === 'noncompliant' || state.includes('disabled') || state.includes('false')) {
                    settingData.disabledCount++;
                  } else {
                    settingData.notConfiguredCount++;
                  }
                }
              }
            } catch (error) {
              logger.debug(`Failed to fetch settings for device ${deviceId}`);
            }
          }
        } catch (error) {
          logger.warn(`Failed to process policy ${policy.id}`, error);
        }
      }

      // Convert map to array of summaries
      for (const [settingName, data] of settingsMap.entries()) {
        const total = data.deviceCount;
        const percentage = total > 0
          ? Math.round((data.enabledCount / total) * 100 * 100) / 100
          : 0;

        summaries.push({
          settingName,
          settingDescription: data.description,
          settingCategory: data.category,
          deviceCount: total,
          enabledCount: data.enabledCount,
          disabledCount: data.disabledCount,
          notConfiguredCount: data.notConfiguredCount,
          percentage,
          affectedPolicies: Array.from(data.policies),
          platforms: data.platforms
        });
      }

      // Sort by device count (descending)
      summaries.sort((a, b) => b.deviceCount - a.deviceCount);

      logger.info(`Generated ${summaries.length} settings summaries`);
      return summaries;
    } catch (error) {
      logger.error('Error fetching mobile device settings summary', error);
      return summaries;
    }
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Maps Exchange access state string to enum
   */
  private mapExchangeAccessState(state: string): ExchangeAccessState {
    if (!state) return ExchangeAccessState.NONE;

    const stateLower = state.toLowerCase();
    if (stateLower.includes('allow')) return ExchangeAccessState.ALLOWED;
    if (stateLower.includes('block')) return ExchangeAccessState.BLOCKED;
    if (stateLower.includes('quarantine')) return ExchangeAccessState.QUARANTINED;
    if (stateLower.includes('unknown')) return ExchangeAccessState.UNKNOWN;

    return ExchangeAccessState.NONE;
  }

  /**
   * Maps Exchange access state reason string to enum
   */
  private mapExchangeAccessStateReason(reason: string): ExchangeAccessStateReason {
    if (!reason) return ExchangeAccessStateReason.NONE;

    const reasonLower = reason.toLowerCase().replace(/[^a-z]/g, '');

    const reasonMap: Record<string, ExchangeAccessStateReason> = {
      'none': ExchangeAccessStateReason.NONE,
      'unknown': ExchangeAccessStateReason.UNKNOWN,
      'exchangeglobalrule': ExchangeAccessStateReason.EXCHANGE_GLOBAL_RULE,
      'exchangeindividualrule': ExchangeAccessStateReason.EXCHANGE_INDIVIDUAL_RULE,
      'exchangedevicerule': ExchangeAccessStateReason.EXCHANGE_DEVICE_RULE,
      'exchangeupgrade': ExchangeAccessStateReason.EXCHANGE_UPGRADE,
      'exchangemailboxpolicy': ExchangeAccessStateReason.EXCHANGE_MAILBOX_POLICY,
      'compliant': ExchangeAccessStateReason.COMPLIANT,
      'notcompliant': ExchangeAccessStateReason.NOT_COMPLIANT,
      'notenrolled': ExchangeAccessStateReason.NOT_ENROLLED,
      'unknownlocation': ExchangeAccessStateReason.UNKNOWN_LOCATION,
      'mfarequired': ExchangeAccessStateReason.MFA_REQUIRED,
      'azureadblockduetoaccesspolicy': ExchangeAccessStateReason.AZURE_AD_BLOCK_DUE_TO_ACCESS_POLICY,
      'compromisedpassword': ExchangeAccessStateReason.COMPROMISED_PASSWORD,
      'devicenotknownwithmanagedapp': ExchangeAccessStateReason.DEVICE_NOT_KNOWN_WITH_MANAGED_APP
    };

    return reasonMap[reasonLower] || ExchangeAccessStateReason.OTHER;
  }

  /**
   * Categorizes setting by type
   */
  private categorizeSettings(settingName: string): 'exchange' | 'compliance' | 'configuration' {
    const nameLower = settingName.toLowerCase();

    if (nameLower.includes('exchange') || nameLower.includes('activesync') ||
        nameLower.includes('email') || nameLower.includes('mailbox')) {
      return 'exchange';
    }

    if (nameLower.includes('compliance') || nameLower.includes('policy') ||
        nameLower.includes('required') || nameLower.includes('minimum')) {
      return 'compliance';
    }

    return 'configuration';
  }

  /**
   * Gets human-readable description for a setting
   */
  private getSettingDescription(settingName: string): string {
    const descriptions: Record<string, string> = {
      'passwordRequired': 'Require a password to unlock mobile devices',
      'passwordMinimumLength': 'Minimum password length requirement',
      'passwordMinutesOfInactivityBeforeLock': 'Minutes of inactivity before device locks',
      'passwordMinimumCharacterSetCount': 'Minimum character sets in password',
      'passwordPreviousPasswordBlockCount': 'Number of previous passwords to prevent reuse',
      'passwordExpirationDays': 'Days until password expires',
      'passwordRequiredType': 'Type of password required',
      'deviceComplianceChecked': 'Device compliance policy is evaluated',
      'osMinimumVersion': 'Minimum operating system version required',
      'osMaximumVersion': 'Maximum operating system version allowed',
      'securityBlockJailbrokenDevices': 'Block jailbroken or rooted devices',
      'storageRequireEncryption': 'Require device storage encryption',
      'requireAppVerify': 'Require app verification on device',
      'requireSafetyNetAttestationBasicIntegrity': 'Require SafetyNet basic integrity',
      'requireGooglePlayServices': 'Require Google Play Services',
      'requireUpToDateSecurityProviders': 'Require up-to-date security providers',
      'requireCompanyPortalAppIntegrity': 'Require Company Portal app integrity'
    };

    return descriptions[settingName] || `${settingName} setting`;
  }

  /**
   * Matches device against inactive device filter
   */
  private matchesInactiveDeviceFilter(device: any, options?: ExchangeActiveSyncFilterOptions): boolean {
    if (!options) return true;

    if (options.platform && device.operatingSystem !== options.platform) {
      return false;
    }

    if (options.exchangeAccessState &&
        this.mapExchangeAccessState(device.exchangeAccessState) !== options.exchangeAccessState) {
      return false;
    }

    if (options.userPrincipalName &&
        !device.userPrincipalName?.toLowerCase().includes(options.userPrincipalName.toLowerCase())) {
      return false;
    }

    return true;
  }

  /**
   * Matches device against compliance details filter
   */
  private matchesComplianceDetailsFilter(
    device: MobileDeviceComplianceDetails,
    options?: ExchangeActiveSyncFilterOptions
  ): boolean {
    if (!options) return true;

    if (options.platform && device.operatingSystem !== options.platform) {
      return false;
    }

    if (options.exchangeAccessState && device.exchangeAccessState !== options.exchangeAccessState) {
      return false;
    }

    if (options.complianceState && device.complianceState !== options.complianceState) {
      return false;
    }

    if (options.userPrincipalName &&
        !device.userPrincipalName.toLowerCase().includes(options.userPrincipalName.toLowerCase())) {
      return false;
    }

    return true;
  }

  /**
   * Generates Exchange ActiveSync summary
   */
  private generateExchangeActiveSyncSummary(data: {
    policyCompliance: ActiveSyncPolicyComplianceSummary[];
    inactiveDevices: InactiveMobileDevice[];
    deviceDetails: MobileDeviceComplianceDetails[];
    settingsSummary: MobileDeviceSettingsSummary[];
  }): ExchangeActiveSyncReportSummary {
    const { policyCompliance, inactiveDevices, deviceDetails, settingsSummary } = data;

    // Aggregate device counts
    const devicesByPlatform: Record<string, number> = {};
    const devicesByAccessState: Record<string, number> = {};
    const devicesByAccessStateReason: Record<string, number> = {};

    let easActivatedCount = 0;
    let totalDaysSinceLastSync = 0;

    for (const device of deviceDetails) {
      // By platform
      devicesByPlatform[device.operatingSystem] =
        (devicesByPlatform[device.operatingSystem] || 0) + 1;

      // By access state
      const accessState = device.exchangeAccessState;
      devicesByAccessState[accessState] = (devicesByAccessState[accessState] || 0) + 1;

      // By access state reason
      const reason = device.exchangeAccessStateReason;
      devicesByAccessStateReason[reason] = (devicesByAccessStateReason[reason] || 0) + 1;

      // EAS activation
      if (device.easActivated) {
        easActivatedCount++;
      }
    }

    // Calculate average days since last sync
    for (const device of inactiveDevices) {
      totalDaysSinceLastSync += device.daysSinceLastSync;
    }
    const averageDaysSinceLastSync = inactiveDevices.length > 0
      ? Math.round(totalDaysSinceLastSync / inactiveDevices.length)
      : 0;

    // Count compliant devices
    const compliantDevices = deviceDetails.filter(d =>
      d.exchangeAccessState === ExchangeAccessState.ALLOWED &&
      d.complianceState.toLowerCase() === 'compliant'
    ).length;

    const compliancePercentage = deviceDetails.length > 0
      ? Math.round((compliantDevices / deviceDetails.length) * 100 * 100) / 100
      : 0;

    return {
      totalMobileDevices: deviceDetails.length,
      activeDevices: deviceDetails.length - inactiveDevices.length,
      inactiveDevices: inactiveDevices.length,
      allowedDevices: devicesByAccessState[ExchangeAccessState.ALLOWED] || 0,
      blockedDevices: devicesByAccessState[ExchangeAccessState.BLOCKED] || 0,
      quarantinedDevices: devicesByAccessState[ExchangeAccessState.QUARANTINED] || 0,
      easActivatedDevices: easActivatedCount,
      compliancePercentage,
      averageDaysSinceLastSync,
      devicesByPlatform,
      devicesByAccessState,
      devicesByAccessStateReason,
      topInactiveThreshold: 30,
      generatedAt: new Date()
    };
  }

  // ============================================================================
  // Export Methods
  // ============================================================================

  /**
   * Export report to JSON format
   */
  async exportToJson(
    reportData: ReportData,
    outputDir: string,
    includeTimestamp: boolean = true
  ): Promise<string> {
    return await OutputFormatter.format(reportData, 'json', outputDir, includeTimestamp);
  }

  /**
   * Export report to CSV format
   */
  async exportToCsv(
    reportData: ReportData,
    outputDir: string,
    includeTimestamp: boolean = true
  ): Promise<string> {
    return await OutputFormatter.format(reportData, 'csv', outputDir, includeTimestamp);
  }

  /**
   * Export report to HTML format
   */
  async exportToHtml(
    reportData: ReportData,
    outputDir: string,
    includeTimestamp: boolean = true
  ): Promise<string> {
    return await OutputFormatter.format(reportData, 'html', outputDir, includeTimestamp);
  }
}

// ============================================================================
// Export for use in other modules
// ============================================================================

export default ExchangeActiveSyncReports;
