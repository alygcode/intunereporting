/**
 * Microsoft Intune Configuration Profile Reporting Solution
 *
 * This module provides comprehensive reporting capabilities for Intune configuration profiles including:
 * - Profile assignment status
 * - Profile deployment success/failure
 * - Device configuration compliance
 * - Policy conflicts detection
 * - Profile per-device status
 */

import { Client } from '@microsoft/microsoft-graph-client';

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Configuration profile types
 */
export enum ProfileType {
  DeviceConfiguration = 'deviceConfiguration',
  SettingsCatalog = 'settingsCatalog',
  Administrative = 'administrativeTemplate',
  Custom = 'custom'
}

/**
 * Profile assignment status
 */
export enum AssignmentStatus {
  Assigned = 'assigned',
  NotAssigned = 'notAssigned',
  Excluded = 'excluded',
  Pending = 'pending'
}

/**
 * Deployment status
 */
export enum DeploymentStatus {
  Success = 'success',
  Failed = 'failed',
  Pending = 'pending',
  NotApplicable = 'notApplicable',
  Error = 'error'
}

/**
 * Compliance status
 */
export enum ComplianceStatus {
  Compliant = 'compliant',
  NonCompliant = 'nonCompliant',
  Error = 'error',
  Conflict = 'conflict',
  NotApplicable = 'notApplicable'
}

/**
 * Configuration profile base interface
 */
export interface ConfigurationProfile {
  id: string;
  displayName: string;
  description?: string;
  profileType: ProfileType;
  platformType: string;
  createdDateTime: string;
  lastModifiedDateTime: string;
  version: number;
  assignments: ProfileAssignment[];
}

/**
 * Profile assignment details
 */
export interface ProfileAssignment {
  id: string;
  target: AssignmentTarget;
  intent?: string;
}

/**
 * Assignment target (groups, users, devices)
 */
export interface AssignmentTarget {
  targetType: 'group' | 'allUsers' | 'allDevices';
  groupId?: string;
  groupName?: string;
  deviceAndAppManagementAssignmentFilterId?: string;
  deviceAndAppManagementAssignmentFilterType?: 'include' | 'exclude';
}

/**
 * Profile assignment status report
 */
export interface ProfileAssignmentReport {
  profileId: string;
  profileName: string;
  profileType: ProfileType;
  totalAssignments: number;
  assignmentDetails: AssignmentDetail[];
  targetedDevicesCount: number;
  targetedUsersCount: number;
  generatedAt: string;
}

/**
 * Individual assignment detail
 */
export interface AssignmentDetail {
  assignmentId: string;
  targetType: string;
  targetName: string;
  targetId?: string;
  status: AssignmentStatus;
  filterName?: string;
  filterType?: string;
}

/**
 * Profile deployment report
 */
export interface ProfileDeploymentReport {
  profileId: string;
  profileName: string;
  profileType: ProfileType;
  totalDevices: number;
  successCount: number;
  failedCount: number;
  pendingCount: number;
  errorCount: number;
  notApplicableCount: number;
  successRate: number;
  deploymentDetails: DeviceDeploymentStatus[];
  generatedAt: string;
}

/**
 * Device-specific deployment status
 */
export interface DeviceDeploymentStatus {
  deviceId: string;
  deviceName: string;
  userName?: string;
  status: DeploymentStatus;
  lastReportedDateTime: string;
  errorCode?: string;
  errorMessage?: string;
  osVersion?: string;
  complianceState?: string;
}

/**
 * Device configuration compliance report
 */
export interface ConfigurationComplianceReport {
  profileId: string;
  profileName: string;
  totalDevices: number;
  compliantCount: number;
  nonCompliantCount: number;
  errorCount: number;
  conflictCount: number;
  notApplicableCount: number;
  complianceRate: number;
  deviceComplianceDetails: DeviceComplianceDetail[];
  settingComplianceSummary: SettingComplianceSummary[];
  generatedAt: string;
}

/**
 * Device compliance details
 */
export interface DeviceComplianceDetail {
  deviceId: string;
  deviceName: string;
  userName?: string;
  complianceStatus: ComplianceStatus;
  lastReportedDateTime: string;
  nonCompliantSettings?: SettingStatus[];
  errorSettings?: SettingStatus[];
}

/**
 * Setting-level status
 */
export interface SettingStatus {
  settingName: string;
  currentValue?: string;
  expectedValue?: string;
  status: ComplianceStatus;
  errorCode?: string;
  errorMessage?: string;
}

/**
 * Setting compliance summary across all devices
 */
export interface SettingComplianceSummary {
  settingName: string;
  compliantCount: number;
  nonCompliantCount: number;
  errorCount: number;
  complianceRate: number;
}

/**
 * Policy conflict detection report
 */
export interface PolicyConflictReport {
  deviceId: string;
  deviceName: string;
  userName?: string;
  totalConflicts: number;
  conflicts: PolicyConflict[];
  affectedSettings: string[];
  generatedAt: string;
}

/**
 * Individual policy conflict
 */
export interface PolicyConflict {
  settingName: string;
  conflictingProfiles: ConflictingProfile[];
  winningProfile?: ConflictingProfile;
  resolutionReason?: string;
}

/**
 * Conflicting profile information
 */
export interface ConflictingProfile {
  profileId: string;
  profileName: string;
  profileType: ProfileType;
  settingValue: string;
  priority?: number;
}

/**
 * Per-device profile status report
 */
export interface PerDeviceProfileReport {
  deviceId: string;
  deviceName: string;
  userName?: string;
  osVersion?: string;
  lastSyncDateTime?: string;
  totalProfiles: number;
  successfulProfiles: number;
  failedProfiles: number;
  pendingProfiles: number;
  profileStatuses: ProfileStatusDetail[];
  generatedAt: string;
}

/**
 * Profile status for a specific device
 */
export interface ProfileStatusDetail {
  profileId: string;
  profileName: string;
  profileType: ProfileType;
  deploymentStatus: DeploymentStatus;
  complianceStatus: ComplianceStatus;
  lastReportedDateTime: string;
  errorCode?: string;
  errorMessage?: string;
  settingsApplied?: number;
  settingsFailed?: number;
}

/**
 * Aggregated reporting options
 */
export interface ReportingOptions {
  includeDetails?: boolean;
  filterByStatus?: DeploymentStatus[];
  filterByPlatform?: string[];
  startDate?: Date;
  endDate?: Date;
  top?: number;
  skip?: number;
}

// ============================================================================
// Configuration Profile Reporting Service
// ============================================================================

export class ConfigurationProfileReportingService {
  private graphClient: Client;

  constructor(graphClient: Client) {
    this.graphClient = graphClient;
  }

  // ==========================================================================
  // 1. Profile Assignment Status Reporting
  // ==========================================================================

  /**
   * Get assignment status for all configuration profiles
   */
  async getAllProfileAssignments(options?: ReportingOptions): Promise<ProfileAssignmentReport[]> {
    try {
      const profiles = await this.getAllConfigurationProfiles();
      const reports: ProfileAssignmentReport[] = [];

      for (const profile of profiles) {
        const report = await this.getProfileAssignmentStatus(profile.id, options);
        reports.push(report);
      }

      return reports;
    } catch (error) {
      throw new Error(`Failed to get profile assignments: ${error.message}`);
    }
  }

  /**
   * Get assignment status for a specific configuration profile
   */
  async getProfileAssignmentStatus(
    profileId: string,
    options?: ReportingOptions
  ): Promise<ProfileAssignmentReport> {
    try {
      // Get profile details
      const profile = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}`)
        .get();

      // Get assignments
      const assignments = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}/assignments`)
        .get();

      // Process assignment details
      const assignmentDetails: AssignmentDetail[] = [];
      let targetedDevicesCount = 0;
      let targetedUsersCount = 0;

      for (const assignment of assignments.value || []) {
        const detail = await this.processAssignment(assignment);
        assignmentDetails.push(detail);

        // Count targets
        if (assignment.target['@odata.type'] === '#microsoft.graph.allDevicesAssignmentTarget') {
          targetedDevicesCount += await this.getDeviceCount();
        } else if (assignment.target['@odata.type'] === '#microsoft.graph.allLicensedUsersAssignmentTarget') {
          targetedUsersCount += await this.getUserCount();
        } else if (assignment.target.groupId) {
          const groupMembers = await this.getGroupMemberCount(assignment.target.groupId);
          targetedUsersCount += groupMembers.users;
          targetedDevicesCount += groupMembers.devices;
        }
      }

      return {
        profileId: profile.id,
        profileName: profile.displayName,
        profileType: this.mapProfileType(profile['@odata.type']),
        totalAssignments: assignmentDetails.length,
        assignmentDetails,
        targetedDevicesCount,
        targetedUsersCount,
        generatedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new Error(`Failed to get assignment status for profile ${profileId}: ${error.message}`);
    }
  }

  /**
   * Process individual assignment
   */
  private async processAssignment(assignment: any): Promise<AssignmentDetail> {
    const targetType = assignment.target['@odata.type'];
    let targetName = 'Unknown';
    let status = AssignmentStatus.Assigned;

    if (targetType === '#microsoft.graph.allDevicesAssignmentTarget') {
      targetName = 'All Devices';
    } else if (targetType === '#microsoft.graph.allLicensedUsersAssignmentTarget') {
      targetName = 'All Users';
    } else if (assignment.target.groupId) {
      try {
        const group = await this.graphClient
          .api(`/groups/${assignment.target.groupId}`)
          .select('displayName')
          .get();
        targetName = group.displayName;
      } catch {
        targetName = `Group (${assignment.target.groupId})`;
      }
    }

    // Get filter information if present
    let filterName: string | undefined;
    if (assignment.target.deviceAndAppManagementAssignmentFilterId) {
      try {
        const filter = await this.graphClient
          .api(`/deviceManagement/assignmentFilters/${assignment.target.deviceAndAppManagementAssignmentFilterId}`)
          .get();
        filterName = filter.displayName;
      } catch {
        filterName = 'Filter Applied';
      }
    }

    return {
      assignmentId: assignment.id,
      targetType: targetType.replace('#microsoft.graph.', ''),
      targetName,
      targetId: assignment.target.groupId,
      status,
      filterName,
      filterType: assignment.target.deviceAndAppManagementAssignmentFilterType
    };
  }

  // ==========================================================================
  // 2. Profile Deployment Success/Failure Reporting
  // ==========================================================================

  /**
   * Get deployment status for a specific configuration profile
   */
  async getProfileDeploymentStatus(
    profileId: string,
    options?: ReportingOptions
  ): Promise<ProfileDeploymentReport> {
    try {
      // Get profile details
      const profile = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}`)
        .get();

      // Get device status overview
      const deviceStatusOverview = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}/deviceStatusOverview`)
        .get();

      // Get detailed device statuses
      const deviceStatuses = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}/deviceStatuses`)
        .top(options?.top || 999)
        .get();

      // Process device statuses
      const deploymentDetails: DeviceDeploymentStatus[] = [];
      let successCount = 0;
      let failedCount = 0;
      let pendingCount = 0;
      let errorCount = 0;
      let notApplicableCount = 0;

      for (const deviceStatus of deviceStatuses.value || []) {
        const detail = await this.processDeviceDeploymentStatus(deviceStatus);
        deploymentDetails.push(detail);

        // Count statuses
        switch (detail.status) {
          case DeploymentStatus.Success:
            successCount++;
            break;
          case DeploymentStatus.Failed:
            failedCount++;
            break;
          case DeploymentStatus.Pending:
            pendingCount++;
            break;
          case DeploymentStatus.Error:
            errorCount++;
            break;
          case DeploymentStatus.NotApplicable:
            notApplicableCount++;
            break;
        }
      }

      const totalDevices = deploymentDetails.length;
      const successRate = totalDevices > 0 ? (successCount / totalDevices) * 100 : 0;

      return {
        profileId: profile.id,
        profileName: profile.displayName,
        profileType: this.mapProfileType(profile['@odata.type']),
        totalDevices,
        successCount,
        failedCount,
        pendingCount,
        errorCount,
        notApplicableCount,
        successRate: Math.round(successRate * 100) / 100,
        deploymentDetails,
        generatedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new Error(`Failed to get deployment status for profile ${profileId}: ${error.message}`);
    }
  }

  /**
   * Get deployment status for all profiles
   */
  async getAllProfileDeploymentStatuses(
    options?: ReportingOptions
  ): Promise<ProfileDeploymentReport[]> {
    try {
      const profiles = await this.getAllConfigurationProfiles();
      const reports: ProfileDeploymentReport[] = [];

      for (const profile of profiles) {
        try {
          const report = await this.getProfileDeploymentStatus(profile.id, options);
          reports.push(report);
        } catch (error) {
          console.error(`Error processing profile ${profile.id}:`, error.message);
        }
      }

      return reports;
    } catch (error) {
      throw new Error(`Failed to get all deployment statuses: ${error.message}`);
    }
  }

  /**
   * Process device deployment status
   */
  private async processDeviceDeploymentStatus(deviceStatus: any): Promise<DeviceDeploymentStatus> {
    let deviceName = 'Unknown Device';
    let userName: string | undefined;
    let osVersion: string | undefined;

    // Try to get device details
    if (deviceStatus.deviceDisplayName) {
      deviceName = deviceStatus.deviceDisplayName;
    }

    if (deviceStatus.userPrincipalName) {
      userName = deviceStatus.userPrincipalName;
    }

    // Map status
    let status = DeploymentStatus.Pending;
    if (deviceStatus.status === 'compliant') {
      status = DeploymentStatus.Success;
    } else if (deviceStatus.status === 'nonCompliant') {
      status = DeploymentStatus.Failed;
    } else if (deviceStatus.status === 'error') {
      status = DeploymentStatus.Error;
    } else if (deviceStatus.status === 'notApplicable') {
      status = DeploymentStatus.NotApplicable;
    }

    return {
      deviceId: deviceStatus.id,
      deviceName,
      userName,
      status,
      lastReportedDateTime: deviceStatus.lastReportedDateTime,
      errorCode: deviceStatus.errorCode,
      errorMessage: deviceStatus.errorMessage,
      osVersion,
      complianceState: deviceStatus.complianceGracePeriodExpirationDateTime
    };
  }

  // ==========================================================================
  // 3. Device Configuration Compliance Reporting
  // ==========================================================================

  /**
   * Get compliance report for a specific configuration profile
   */
  async getConfigurationComplianceReport(
    profileId: string,
    options?: ReportingOptions
  ): Promise<ConfigurationComplianceReport> {
    try {
      // Get profile details
      const profile = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}`)
        .get();

      // Get device statuses with settings details
      const deviceStatuses = await this.graphClient
        .api(`/deviceManagement/deviceConfigurations/${profileId}/deviceStatuses`)
        .expand('deviceConfiguration')
        .top(options?.top || 999)
        .get();

      // Process compliance details
      const deviceComplianceDetails: DeviceComplianceDetail[] = [];
      const settingComplianceMap = new Map<string, { compliant: number; nonCompliant: number; error: number }>();

      let compliantCount = 0;
      let nonCompliantCount = 0;
      let errorCount = 0;
      let conflictCount = 0;
      let notApplicableCount = 0;

      for (const deviceStatus of deviceStatuses.value || []) {
        const detail = await this.processDeviceComplianceDetail(deviceStatus);
        deviceComplianceDetails.push(detail);

        // Count compliance states
        switch (detail.complianceStatus) {
          case ComplianceStatus.Compliant:
            compliantCount++;
            break;
          case ComplianceStatus.NonCompliant:
            nonCompliantCount++;
            break;
          case ComplianceStatus.Error:
            errorCount++;
            break;
          case ComplianceStatus.Conflict:
            conflictCount++;
            break;
          case ComplianceStatus.NotApplicable:
            notApplicableCount++;
            break;
        }

        // Aggregate setting compliance
        if (detail.nonCompliantSettings) {
          for (const setting of detail.nonCompliantSettings) {
            const stats = settingComplianceMap.get(setting.settingName) || { compliant: 0, nonCompliant: 0, error: 0 };
            stats.nonCompliant++;
            settingComplianceMap.set(setting.settingName, stats);
          }
        }

        if (detail.errorSettings) {
          for (const setting of detail.errorSettings) {
            const stats = settingComplianceMap.get(setting.settingName) || { compliant: 0, nonCompliant: 0, error: 0 };
            stats.error++;
            settingComplianceMap.set(setting.settingName, stats);
          }
        }
      }

      // Create setting compliance summary
      const settingComplianceSummary: SettingComplianceSummary[] = [];
      const totalDevices = deviceComplianceDetails.length;

      settingComplianceMap.forEach((stats, settingName) => {
        const compliant = totalDevices - stats.nonCompliant - stats.error;
        const complianceRate = totalDevices > 0 ? (compliant / totalDevices) * 100 : 0;

        settingComplianceSummary.push({
          settingName,
          compliantCount: compliant,
          nonCompliantCount: stats.nonCompliant,
          errorCount: stats.error,
          complianceRate: Math.round(complianceRate * 100) / 100
        });
      });

      const complianceRate = totalDevices > 0 ? (compliantCount / totalDevices) * 100 : 0;

      return {
        profileId: profile.id,
        profileName: profile.displayName,
        totalDevices,
        compliantCount,
        nonCompliantCount,
        errorCount,
        conflictCount,
        notApplicableCount,
        complianceRate: Math.round(complianceRate * 100) / 100,
        deviceComplianceDetails,
        settingComplianceSummary: settingComplianceSummary.sort((a, b) => a.complianceRate - b.complianceRate),
        generatedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new Error(`Failed to get compliance report for profile ${profileId}: ${error.message}`);
    }
  }

  /**
   * Process device compliance detail
   */
  private async processDeviceComplianceDetail(deviceStatus: any): Promise<DeviceComplianceDetail> {
    // Map compliance status
    let complianceStatus = ComplianceStatus.NotApplicable;
    if (deviceStatus.status === 'compliant') {
      complianceStatus = ComplianceStatus.Compliant;
    } else if (deviceStatus.status === 'nonCompliant') {
      complianceStatus = ComplianceStatus.NonCompliant;
    } else if (deviceStatus.status === 'error') {
      complianceStatus = ComplianceStatus.Error;
    } else if (deviceStatus.status === 'conflict') {
      complianceStatus = ComplianceStatus.Conflict;
    }

    const nonCompliantSettings: SettingStatus[] = [];
    const errorSettings: SettingStatus[] = [];

    // Process setting states if available
    if (deviceStatus.settingStates) {
      for (const settingState of deviceStatus.settingStates) {
        const setting: SettingStatus = {
          settingName: settingState.setting || settingState.settingName,
          currentValue: settingState.currentValue,
          expectedValue: settingState.desiredValue,
          status: this.mapSettingStatus(settingState.state),
          errorCode: settingState.errorCode,
          errorMessage: settingState.errorDescription
        };

        if (setting.status === ComplianceStatus.NonCompliant) {
          nonCompliantSettings.push(setting);
        } else if (setting.status === ComplianceStatus.Error) {
          errorSettings.push(setting);
        }
      }
    }

    return {
      deviceId: deviceStatus.id,
      deviceName: deviceStatus.deviceDisplayName || 'Unknown Device',
      userName: deviceStatus.userPrincipalName,
      complianceStatus,
      lastReportedDateTime: deviceStatus.lastReportedDateTime,
      nonCompliantSettings: nonCompliantSettings.length > 0 ? nonCompliantSettings : undefined,
      errorSettings: errorSettings.length > 0 ? errorSettings : undefined
    };
  }

  // ==========================================================================
  // 4. Policy Conflicts Detection
  // ==========================================================================

  /**
   * Detect policy conflicts for a specific device
   */
  async detectPolicyConflicts(deviceId: string): Promise<PolicyConflictReport> {
    try {
      // Get device details
      const device = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}`)
        .select('deviceName,userPrincipalName')
        .get();

      // Get all configuration policies applied to the device
      const deviceConfigurations = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/deviceConfigurationStates`)
        .expand('settingStates')
        .get();

      // Analyze conflicts
      const settingMap = new Map<string, ConflictingProfile[]>();

      for (const config of deviceConfigurations.value || []) {
        if (config.settingStates) {
          for (const settingState of config.settingStates) {
            const settingName = settingState.setting || settingState.settingName;

            if (!settingMap.has(settingName)) {
              settingMap.set(settingName, []);
            }

            const profiles = settingMap.get(settingName)!;
            profiles.push({
              profileId: config.id,
              profileName: config.displayName,
              profileType: ProfileType.DeviceConfiguration,
              settingValue: settingState.currentValue || settingState.desiredValue,
              priority: config.priority
            });
          }
        }
      }

      // Identify actual conflicts (same setting, different values)
      const conflicts: PolicyConflict[] = [];
      const affectedSettings: string[] = [];

      settingMap.forEach((profiles, settingName) => {
        if (profiles.length > 1) {
          // Check if there are different values
          const uniqueValues = new Set(profiles.map(p => p.settingValue));

          if (uniqueValues.size > 1) {
            // This is a conflict
            const sortedProfiles = profiles.sort((a, b) => (a.priority || 999) - (b.priority || 999));

            conflicts.push({
              settingName,
              conflictingProfiles: sortedProfiles,
              winningProfile: sortedProfiles[0],
              resolutionReason: 'Resolved by policy priority'
            });

            affectedSettings.push(settingName);
          }
        }
      });

      return {
        deviceId: device.id,
        deviceName: device.deviceName,
        userName: device.userPrincipalName,
        totalConflicts: conflicts.length,
        conflicts,
        affectedSettings,
        generatedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new Error(`Failed to detect policy conflicts for device ${deviceId}: ${error.message}`);
    }
  }

  /**
   * Detect policy conflicts across all devices
   */
  async detectAllPolicyConflicts(options?: ReportingOptions): Promise<PolicyConflictReport[]> {
    try {
      // Get all managed devices
      const devices = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select('id,deviceName')
        .top(options?.top || 100)
        .get();

      const conflictReports: PolicyConflictReport[] = [];

      for (const device of devices.value || []) {
        try {
          const report = await this.detectPolicyConflicts(device.id);
          if (report.totalConflicts > 0) {
            conflictReports.push(report);
          }
        } catch (error) {
          console.error(`Error detecting conflicts for device ${device.id}:`, error.message);
        }
      }

      return conflictReports;
    } catch (error) {
      throw new Error(`Failed to detect all policy conflicts: ${error.message}`);
    }
  }

  // ==========================================================================
  // 5. Profile Per-Device Status Reporting
  // ==========================================================================

  /**
   * Get all profile statuses for a specific device
   */
  async getPerDeviceProfileStatus(deviceId: string): Promise<PerDeviceProfileReport> {
    try {
      // Get device details
      const device = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}`)
        .select('deviceName,userPrincipalName,osVersion,lastSyncDateTime')
        .get();

      // Get all configuration states for the device
      const configStates = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/deviceConfigurationStates`)
        .get();

      // Process profile statuses
      const profileStatuses: ProfileStatusDetail[] = [];
      let successfulProfiles = 0;
      let failedProfiles = 0;
      let pendingProfiles = 0;

      for (const configState of configStates.value || []) {
        const detail = await this.processProfileStatusDetail(configState);
        profileStatuses.push(detail);

        // Count statuses
        if (detail.deploymentStatus === DeploymentStatus.Success) {
          successfulProfiles++;
        } else if (detail.deploymentStatus === DeploymentStatus.Failed) {
          failedProfiles++;
        } else if (detail.deploymentStatus === DeploymentStatus.Pending) {
          pendingProfiles++;
        }
      }

      return {
        deviceId: device.id,
        deviceName: device.deviceName,
        userName: device.userPrincipalName,
        osVersion: device.osVersion,
        lastSyncDateTime: device.lastSyncDateTime,
        totalProfiles: profileStatuses.length,
        successfulProfiles,
        failedProfiles,
        pendingProfiles,
        profileStatuses,
        generatedAt: new Date().toISOString()
      };
    } catch (error) {
      throw new Error(`Failed to get per-device profile status for device ${deviceId}: ${error.message}`);
    }
  }

  /**
   * Get profile status across all devices
   */
  async getAllDevicesProfileStatus(options?: ReportingOptions): Promise<PerDeviceProfileReport[]> {
    try {
      // Get all managed devices
      const devices = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select('id,deviceName')
        .top(options?.top || 100)
        .get();

      const reports: PerDeviceProfileReport[] = [];

      for (const device of devices.value || []) {
        try {
          const report = await this.getPerDeviceProfileStatus(device.id);
          reports.push(report);
        } catch (error) {
          console.error(`Error getting profile status for device ${device.id}:`, error.message);
        }
      }

      return reports;
    } catch (error) {
      throw new Error(`Failed to get all devices profile status: ${error.message}`);
    }
  }

  /**
   * Process profile status detail
   */
  private async processProfileStatusDetail(configState: any): Promise<ProfileStatusDetail> {
    // Map deployment status
    let deploymentStatus = DeploymentStatus.Pending;
    if (configState.state === 'compliant') {
      deploymentStatus = DeploymentStatus.Success;
    } else if (configState.state === 'nonCompliant' || configState.state === 'error') {
      deploymentStatus = DeploymentStatus.Failed;
    }

    // Map compliance status
    let complianceStatus = ComplianceStatus.NotApplicable;
    if (configState.state === 'compliant') {
      complianceStatus = ComplianceStatus.Compliant;
    } else if (configState.state === 'nonCompliant') {
      complianceStatus = ComplianceStatus.NonCompliant;
    } else if (configState.state === 'error') {
      complianceStatus = ComplianceStatus.Error;
    } else if (configState.state === 'conflict') {
      complianceStatus = ComplianceStatus.Conflict;
    }

    // Count settings
    let settingsApplied = 0;
    let settingsFailed = 0;

    if (configState.settingStates) {
      for (const setting of configState.settingStates) {
        if (setting.state === 'compliant') {
          settingsApplied++;
        } else if (setting.state === 'nonCompliant' || setting.state === 'error') {
          settingsFailed++;
        }
      }
    }

    return {
      profileId: configState.id,
      profileName: configState.displayName,
      profileType: ProfileType.DeviceConfiguration,
      deploymentStatus,
      complianceStatus,
      lastReportedDateTime: configState.lastReportedDateTime || new Date().toISOString(),
      errorCode: configState.errorCode,
      errorMessage: configState.errorDescription,
      settingsApplied: settingsApplied > 0 ? settingsApplied : undefined,
      settingsFailed: settingsFailed > 0 ? settingsFailed : undefined
    };
  }

  // ==========================================================================
  // Utility Methods
  // ==========================================================================

  /**
   * Get all configuration profiles
   */
  private async getAllConfigurationProfiles(): Promise<any[]> {
    try {
      const profiles = await this.graphClient
        .api('/deviceManagement/deviceConfigurations')
        .get();

      return profiles.value || [];
    } catch (error) {
      throw new Error(`Failed to get configuration profiles: ${error.message}`);
    }
  }

  /**
   * Map profile type from OData type
   */
  private mapProfileType(odataType: string): ProfileType {
    if (odataType.includes('windows10CustomConfiguration')) {
      return ProfileType.Custom;
    } else if (odataType.includes('windows10GeneralConfiguration')) {
      return ProfileType.DeviceConfiguration;
    } else if (odataType.includes('administrativeTemplate')) {
      return ProfileType.Administrative;
    }
    return ProfileType.DeviceConfiguration;
  }

  /**
   * Map setting status
   */
  private mapSettingStatus(state: string): ComplianceStatus {
    switch (state?.toLowerCase()) {
      case 'compliant':
        return ComplianceStatus.Compliant;
      case 'noncompliant':
        return ComplianceStatus.NonCompliant;
      case 'error':
        return ComplianceStatus.Error;
      case 'conflict':
        return ComplianceStatus.Conflict;
      default:
        return ComplianceStatus.NotApplicable;
    }
  }

  /**
   * Get total device count
   */
  private async getDeviceCount(): Promise<number> {
    try {
      const result = await this.graphClient
        .api('/deviceManagement/managedDevices/$count')
        .get();
      return result || 0;
    } catch {
      return 0;
    }
  }

  /**
   * Get total user count
   */
  private async getUserCount(): Promise<number> {
    try {
      const result = await this.graphClient
        .api('/users/$count')
        .header('ConsistencyLevel', 'eventual')
        .get();
      return result || 0;
    } catch {
      return 0;
    }
  }

  /**
   * Get group member count
   */
  private async getGroupMemberCount(groupId: string): Promise<{ users: number; devices: number }> {
    try {
      const members = await this.graphClient
        .api(`/groups/${groupId}/members`)
        .select('@odata.type')
        .get();

      let users = 0;
      let devices = 0;

      for (const member of members.value || []) {
        if (member['@odata.type'] === '#microsoft.graph.user') {
          users++;
        } else if (member['@odata.type'] === '#microsoft.graph.device') {
          devices++;
        }
      }

      return { users, devices };
    } catch {
      return { users: 0, devices: 0 };
    }
  }

  // ==========================================================================
  // Export Methods
  // ==========================================================================

  /**
   * Export profile assignment report to JSON
   */
  exportAssignmentReportToJSON(report: ProfileAssignmentReport): string {
    return JSON.stringify(report, null, 2);
  }

  /**
   * Export deployment report to JSON
   */
  exportDeploymentReportToJSON(report: ProfileDeploymentReport): string {
    return JSON.stringify(report, null, 2);
  }

  /**
   * Export compliance report to JSON
   */
  exportComplianceReportToJSON(report: ConfigurationComplianceReport): string {
    return JSON.stringify(report, null, 2);
  }

  /**
   * Export conflict report to JSON
   */
  exportConflictReportToJSON(report: PolicyConflictReport): string {
    return JSON.stringify(report, null, 2);
  }

  /**
   * Export per-device report to JSON
   */
  exportPerDeviceReportToJSON(report: PerDeviceProfileReport): string {
    return JSON.stringify(report, null, 2);
  }

  /**
   * Export report to CSV format
   */
  exportReportToCSV(data: any[], headers: string[]): string {
    const csvRows: string[] = [];

    // Add header
    csvRows.push(headers.join(','));

    // Add data rows
    for (const row of data) {
      const values = headers.map(header => {
        const value = row[header];
        // Escape commas and quotes
        const escaped = ('' + value).replace(/"/g, '""');
        return `"${escaped}"`;
      });
      csvRows.push(values.join(','));
    }

    return csvRows.join('\n');
  }
}

// ============================================================================
// Example Usage
// ============================================================================

/**
 * Example: Initialize and use the reporting service
 */
export async function exampleUsage(graphClient: Client) {
  const reportingService = new ConfigurationProfileReportingService(graphClient);

  // 1. Get profile assignment status
  console.log('Fetching profile assignments...');
  const assignmentReports = await reportingService.getAllProfileAssignments();
  console.log(`Found ${assignmentReports.length} profiles with assignments`);

  // 2. Get deployment status for a specific profile
  console.log('\nFetching deployment status...');
  if (assignmentReports.length > 0) {
    const deploymentReport = await reportingService.getProfileDeploymentStatus(
      assignmentReports[0].profileId
    );
    console.log(`Profile: ${deploymentReport.profileName}`);
    console.log(`Success Rate: ${deploymentReport.successRate}%`);
    console.log(`Successful: ${deploymentReport.successCount}`);
    console.log(`Failed: ${deploymentReport.failedCount}`);
  }

  // 3. Get compliance report
  console.log('\nFetching compliance report...');
  if (assignmentReports.length > 0) {
    const complianceReport = await reportingService.getConfigurationComplianceReport(
      assignmentReports[0].profileId
    );
    console.log(`Compliance Rate: ${complianceReport.complianceRate}%`);
    console.log(`Compliant Devices: ${complianceReport.compliantCount}`);
    console.log(`Non-Compliant Devices: ${complianceReport.nonCompliantCount}`);
  }

  // 4. Detect policy conflicts
  console.log('\nDetecting policy conflicts...');
  const conflictReports = await reportingService.detectAllPolicyConflicts({ top: 10 });
  console.log(`Found ${conflictReports.length} devices with conflicts`);

  for (const report of conflictReports) {
    console.log(`\nDevice: ${report.deviceName}`);
    console.log(`Total Conflicts: ${report.totalConflicts}`);
    for (const conflict of report.conflicts) {
      console.log(`  - ${conflict.settingName}: ${conflict.conflictingProfiles.length} conflicting profiles`);
    }
  }

  // 5. Get per-device profile status
  console.log('\nFetching per-device status...');
  const deviceReports = await reportingService.getAllDevicesProfileStatus({ top: 5 });

  for (const report of deviceReports) {
    console.log(`\nDevice: ${report.deviceName}`);
    console.log(`Total Profiles: ${report.totalProfiles}`);
    console.log(`Successful: ${report.successfulProfiles}`);
    console.log(`Failed: ${report.failedProfiles}`);
  }

  // Export examples
  console.log('\nExporting reports...');
  if (assignmentReports.length > 0) {
    const json = reportingService.exportAssignmentReportToJSON(assignmentReports[0]);
    console.log('Assignment report exported to JSON');
  }
}

export default ConfigurationProfileReportingService;
