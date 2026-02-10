/**
 * Policy Assignment Report Generator for Microsoft Intune
 *
 * This module provides comprehensive reporting functionality for Microsoft Intune
 * configuration policies, including assignments, deployment status, conflicts,
 * and success/failure rates using the Microsoft Graph API.
 *
 * Features:
 * - Fetch all configuration policies
 * - Get policy assignments (users/groups/devices)
 * - Get policy deployment status (device and user statuses)
 * - Detect and report policy conflicts
 * - Calculate success and failure rates
 * - Export to multiple formats (JSON, CSV, HTML)
 * - Comprehensive error handling and retry logic
 * - Automatic pagination support
 *
 * @module policy-assignment-report
 */

import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';
import * as fs from 'fs/promises';
import * as path from 'path';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Configuration policy information
 */
export interface PolicyConfiguration {
  /** Policy ID */
  id: string;
  /** Display name of the policy */
  displayName: string;
  /** Policy description */
  description: string;
  /** Policy platform type */
  platformType: string;
  /** Policy technology type */
  technologies: string;
  /** Creation date */
  createdDateTime: string;
  /** Last modified date */
  lastModifiedDateTime: string;
  /** Number of assignments */
  assignmentCount: number;
  /** Number of device statuses */
  deviceStatusCount: number;
  /** Number of user statuses */
  userStatusCount: number;
  /** Policy version */
  version?: number;
}

/**
 * Policy assignment information
 */
export interface PolicyAssignment {
  /** Assignment ID */
  id: string;
  /** Policy ID */
  policyId: string;
  /** Policy name */
  policyName: string;
  /** Target group ID */
  targetGroupId?: string;
  /** Target group name */
  targetGroupName?: string;
  /** Assignment intent (apply, remove, etc.) */
  intent?: string;
  /** Assignment source (direct, policySets, etc.) */
  source?: string;
  /** Assignment target type (group, allUsers, allDevices) */
  targetType: string;
  /** Include/exclude filter */
  filterType?: string;
}

/**
 * Device deployment status for a policy
 */
export interface DeviceDeploymentStatus {
  /** Device ID */
  deviceId: string;
  /** Device name */
  deviceName: string;
  /** User principal name */
  userPrincipalName: string;
  /** Policy ID */
  policyId: string;
  /** Policy name */
  policyName: string;
  /** Status (success, error, conflict, notApplicable, etc.) */
  status: string;
  /** Compliance state */
  complianceState?: string;
  /** Last reported date */
  lastReportedDateTime: string;
  /** Platform (Windows, iOS, Android, etc.) */
  platform?: string;
  /** OS version */
  osVersion?: string;
}

/**
 * User deployment status for a policy
 */
export interface UserDeploymentStatus {
  /** User ID */
  userId: string;
  /** User principal name */
  userPrincipalName: string;
  /** Policy ID */
  policyId: string;
  /** Policy name */
  policyName: string;
  /** Status (success, error, conflict, notApplicable, etc.) */
  status: string;
  /** Last reported date */
  lastReportedDateTime: string;
  /** Number of devices affected */
  devicesCount?: number;
}

/**
 * Policy conflict information
 */
export interface PolicyConflict {
  /** Conflict ID */
  id: string;
  /** Conflicting policy IDs */
  conflictingPolicyIds: string[];
  /** Conflicting policy names */
  conflictingPolicyNames: string[];
  /** Affected devices count */
  affectedDevicesCount: number;
  /** Conflict setting name */
  conflictingSettingName?: string;
  /** Conflict description */
  description?: string;
}

/**
 * Success/failure rate statistics
 */
export interface PolicySuccessRate {
  /** Policy ID */
  policyId: string;
  /** Policy name */
  policyName: string;
  /** Total targets (devices or users) */
  totalTargets: number;
  /** Successful deployments */
  successCount: number;
  /** Failed deployments */
  failureCount: number;
  /** Conflicts */
  conflictCount: number;
  /** Pending deployments */
  pendingCount: number;
  /** Not applicable */
  notApplicableCount: number;
  /** Success percentage */
  successRate: number;
  /** Failure percentage */
  failureRate: number;
}

/**
 * Overall policy assignment summary
 */
export interface PolicyAssignmentSummary {
  /** Total policies */
  totalPolicies: number;
  /** Policies with assignments */
  policiesWithAssignments: number;
  /** Policies without assignments */
  policiesWithoutAssignments: number;
  /** Total assignments */
  totalAssignments: number;
  /** Total device deployments */
  totalDeviceDeployments: number;
  /** Total user deployments */
  totalUserDeployments: number;
  /** Total conflicts detected */
  totalConflicts: number;
  /** Overall success rate */
  overallSuccessRate: number;
  /** Overall failure rate */
  overallFailureRate: number;
  /** Policies by platform */
  policiesByPlatform: Record<string, number>;
  /** Report generation timestamp */
  generatedAt: string;
}

/**
 * Options for policy assignment report generation
 */
export interface PolicyReportOptions {
  /** Include device deployment status */
  includeDeviceStatus?: boolean;
  /** Include user deployment status */
  includeUserStatus?: boolean;
  /** Include conflict analysis */
  includeConflicts?: boolean;
  /** Include success/failure rates */
  includeSuccessRates?: boolean;
  /** Filter by policy ID */
  policyId?: string;
  /** Filter by platform */
  platformFilter?: string;
  /** Maximum number of retry attempts */
  maxRetries?: number;
  /** Output directory for exports */
  outputDir?: string;
}

/**
 * Complete policy assignment report data
 */
export interface PolicyAssignmentReportData {
  /** Policy configurations */
  policies: PolicyConfiguration[];
  /** Policy assignments */
  assignments: PolicyAssignment[];
  /** Device deployment statuses */
  deviceStatuses: DeviceDeploymentStatus[];
  /** User deployment statuses */
  userStatuses: UserDeploymentStatus[];
  /** Policy conflicts */
  conflicts: PolicyConflict[];
  /** Success/failure rates */
  successRates: PolicySuccessRate[];
  /** Summary statistics */
  summary: PolicyAssignmentSummary;
}

// ============================================================================
// Main Class
// ============================================================================

/**
 * Policy Assignment Report Generator
 *
 * Comprehensive reporting for Intune configuration policies including
 * assignments, deployment status, conflicts, and success metrics.
 *
 * @example
 * ```typescript
 * const report = new PolicyAssignmentReport(graphClient, config);
 * const data = await report.execute();
 * await report.exportToCSV(data, './reports');
 * await report.exportToHTML(data, './reports');
 * ```
 */
export class PolicyAssignmentReport extends BaseReport {
  name = 'policy-assignment-report';
  description = 'Comprehensive policy assignment and deployment status report';
  category = 'Policies';
  enabled = true;

  /**
   * Execute the policy assignment report
   *
   * @param options - Report generation options
   * @returns Complete report data
   */
  async execute(options: PolicyReportOptions = {}): Promise<ReportData> {
    logger.info('Executing Policy Assignment Report');

    try {
      // Fetch all configuration policies
      const policies = await this.getAllConfigurationPolicies(options);
      logger.info(`Retrieved ${policies.length} configuration policies`);

      // Fetch assignments for all policies
      const assignments = await this.getAllPolicyAssignments(policies);
      logger.info(`Retrieved ${assignments.length} policy assignments`);

      // Initialize arrays for optional data
      let deviceStatuses: DeviceDeploymentStatus[] = [];
      let userStatuses: UserDeploymentStatus[] = [];
      let conflicts: PolicyConflict[] = [];
      let successRates: PolicySuccessRate[] = [];

      // Fetch device deployment status if requested
      if (options.includeDeviceStatus !== false) {
        deviceStatuses = await this.getAllDeviceStatuses(policies);
        logger.info(`Retrieved ${deviceStatuses.length} device deployment statuses`);
      }

      // Fetch user deployment status if requested
      if (options.includeUserStatus !== false) {
        userStatuses = await this.getAllUserStatuses(policies);
        logger.info(`Retrieved ${userStatuses.length} user deployment statuses`);
      }

      // Fetch policy conflicts if requested
      if (options.includeConflicts !== false) {
        conflicts = await this.getPolicyConflicts();
        logger.info(`Retrieved ${conflicts.length} policy conflicts`);
      }

      // Calculate success/failure rates if requested
      if (options.includeSuccessRates !== false) {
        successRates = this.calculateSuccessRates(policies, deviceStatuses, userStatuses);
        logger.info(`Calculated success rates for ${successRates.length} policies`);
      }

      // Generate summary
      const summary = this.generateSummary(
        policies,
        assignments,
        deviceStatuses,
        userStatuses,
        conflicts,
        successRates
      );

      // Prepare report data
      const reportData: PolicyAssignmentReportData = {
        policies,
        assignments,
        deviceStatuses,
        userStatuses,
        conflicts,
        successRates,
        summary
      };

      // Export if output directory specified
      if (options.outputDir) {
        await this.exportReports(reportData, options.outputDir);
      }

      logger.info('Policy Assignment Report completed successfully');

      return {
        metadata: this.createMetadata(this.name, policies.length, options),
        data: [reportData],
        summary
      };
    } catch (error) {
      logger.error('Failed to execute Policy Assignment Report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Data Fetching Methods
  // ==========================================================================

  /**
   * Fetch all configuration policies from Intune
   *
   * @param options - Report options
   * @returns Array of policy configurations
   */
  async getAllConfigurationPolicies(
    options: PolicyReportOptions
  ): Promise<PolicyConfiguration[]> {
    logger.debug('Fetching configuration policies');

    const policies: PolicyConfiguration[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceConfigurations')
          .select([
            'id',
            'displayName',
            'description',
            'createdDateTime',
            'lastModifiedDateTime',
            'version'
          ])
          .top(999)
          .get()
      );

      const allPolicies = await this.getAllPages(response);

      for (const policy of allPolicies) {
        // Apply filters
        if (options.policyId && policy.id !== options.policyId) {
          continue;
        }

        // Get additional policy details
        const policyDetails = await this.getPolicyDetails(policy.id);

        policies.push({
          id: policy.id,
          displayName: policy.displayName || 'Unnamed Policy',
          description: policy.description || 'No description',
          platformType: this.determinePlatformType(policy),
          technologies: this.determineTechnologies(policy),
          createdDateTime: policy.createdDateTime || new Date().toISOString(),
          lastModifiedDateTime: policy.lastModifiedDateTime || new Date().toISOString(),
          assignmentCount: policyDetails.assignmentCount || 0,
          deviceStatusCount: policyDetails.deviceStatusCount || 0,
          userStatusCount: policyDetails.userStatusCount || 0,
          version: policy.version
        });
      }

      return policies;
    } catch (error) {
      logger.error('Error fetching configuration policies', error);
      throw error;
    }
  }

  /**
   * Get additional details for a specific policy
   *
   * @param policyId - Policy ID
   * @returns Policy detail counts
   */
  private async getPolicyDetails(
    policyId: string
  ): Promise<{ assignmentCount: number; deviceStatusCount: number; userStatusCount: number }> {
    try {
      const [assignmentsResponse, deviceStatusResponse, userStatusResponse] = await Promise.all([
        this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policyId}/assignments`)
            .get()
            .catch(() => ({ value: [] }))
        ),
        this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policyId}/deviceStatuses`)
            .top(1)
            .get()
            .catch(() => ({ value: [] }))
        ),
        this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policyId}/userStatuses`)
            .top(1)
            .get()
            .catch(() => ({ value: [] }))
        )
      ]);

      return {
        assignmentCount: assignmentsResponse.value?.length || 0,
        deviceStatusCount: deviceStatusResponse['@odata.count'] || 0,
        userStatusCount: userStatusResponse['@odata.count'] || 0
      };
    } catch (error) {
      logger.warn(`Failed to get details for policy ${policyId}`, error);
      return { assignmentCount: 0, deviceStatusCount: 0, userStatusCount: 0 };
    }
  }

  /**
   * Fetch all policy assignments
   *
   * @param policies - Array of policies to fetch assignments for
   * @returns Array of policy assignments
   */
  async getAllPolicyAssignments(
    policies: PolicyConfiguration[]
  ): Promise<PolicyAssignment[]> {
    logger.debug('Fetching policy assignments');

    const assignments: PolicyAssignment[] = [];

    for (const policy of policies) {
      try {
        const response = await this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policy.id}/assignments`)
            .expand('target')
            .get()
        );

        const policyAssignments = response.value || [];

        for (const assignment of policyAssignments) {
          const targetType = this.determineTargetType(assignment.target);

          // Get group name if it's a group assignment
          let targetGroupName = undefined;
          if (assignment.target?.groupId) {
            targetGroupName = await this.getGroupName(assignment.target.groupId);
          }

          assignments.push({
            id: assignment.id,
            policyId: policy.id,
            policyName: policy.displayName,
            targetGroupId: assignment.target?.groupId,
            targetGroupName,
            intent: assignment.intent,
            source: assignment.source,
            targetType,
            filterType: assignment.target?.deviceAndAppManagementAssignmentFilterType
          });
        }
      } catch (error) {
        logger.warn(`Failed to fetch assignments for policy ${policy.displayName}`, error);
      }
    }

    return assignments;
  }

  /**
   * Fetch device deployment statuses for all policies
   *
   * @param policies - Array of policies
   * @returns Array of device deployment statuses
   */
  async getAllDeviceStatuses(
    policies: PolicyConfiguration[]
  ): Promise<DeviceDeploymentStatus[]> {
    logger.debug('Fetching device deployment statuses');

    const statuses: DeviceDeploymentStatus[] = [];

    for (const policy of policies) {
      try {
        const response = await this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policy.id}/deviceStatuses`)
            .select([
              'id',
              'deviceDisplayName',
              'userName',
              'status',
              'complianceGracePeriodExpirationDateTime',
              'lastReportedDateTime',
              'platform',
              'osVersion'
            ])
            .top(999)
            .get()
        );

        const deviceStatuses = await this.getAllPages(response);

        for (const status of deviceStatuses) {
          statuses.push({
            deviceId: status.id || 'unknown',
            deviceName: status.deviceDisplayName || 'Unknown Device',
            userPrincipalName: status.userName || 'Unknown User',
            policyId: policy.id,
            policyName: policy.displayName,
            status: status.status || 'unknown',
            complianceState: status.complianceGracePeriodExpirationDateTime,
            lastReportedDateTime: status.lastReportedDateTime || new Date().toISOString(),
            platform: status.platform,
            osVersion: status.osVersion
          });
        }
      } catch (error) {
        logger.warn(`Failed to fetch device statuses for policy ${policy.displayName}`, error);
      }
    }

    return statuses;
  }

  /**
   * Fetch user deployment statuses for all policies
   *
   * @param policies - Array of policies
   * @returns Array of user deployment statuses
   */
  async getAllUserStatuses(
    policies: PolicyConfiguration[]
  ): Promise<UserDeploymentStatus[]> {
    logger.debug('Fetching user deployment statuses');

    const statuses: UserDeploymentStatus[] = [];

    for (const policy of policies) {
      try {
        const response = await this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policy.id}/userStatuses`)
            .select([
              'id',
              'userPrincipalName',
              'status',
              'lastReportedDateTime',
              'devicesCount'
            ])
            .top(999)
            .get()
        );

        const userStatuses = await this.getAllPages(response);

        for (const status of userStatuses) {
          statuses.push({
            userId: status.id || 'unknown',
            userPrincipalName: status.userPrincipalName || 'Unknown User',
            policyId: policy.id,
            policyName: policy.displayName,
            status: status.status || 'unknown',
            lastReportedDateTime: status.lastReportedDateTime || new Date().toISOString(),
            devicesCount: status.devicesCount || 0
          });
        }
      } catch (error) {
        logger.warn(`Failed to fetch user statuses for policy ${policy.displayName}`, error);
      }
    }

    return statuses;
  }

  /**
   * Fetch policy conflicts from Intune
   *
   * @returns Array of policy conflicts
   */
  async getPolicyConflicts(): Promise<PolicyConflict[]> {
    logger.debug('Fetching policy conflicts');

    const conflicts: PolicyConflict[] = [];

    try {
      // Try to fetch conflict summary
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceConfigurationConflictSummary')
          .get()
          .catch(() => ({ value: [] }))
      );

      const conflictData = await this.getAllPages(response);

      for (const conflict of conflictData) {
        const policyIds = conflict.contributingSettings?.map((s: any) => s.settingName) || [];
        const policyNames = await this.getPolicyNames(policyIds);

        conflicts.push({
          id: conflict.id || `conflict-${conflicts.length}`,
          conflictingPolicyIds: policyIds,
          conflictingPolicyNames: policyNames,
          affectedDevicesCount: conflict.conflictedDeviceConfigurations?.length || 0,
          conflictingSettingName: conflict.settingName,
          description: conflict.description || 'Policy conflict detected'
        });
      }
    } catch (error) {
      logger.warn('Failed to fetch policy conflicts (may not be available)', error);
      // Conflicts may not be available in all tenants/licenses
    }

    return conflicts;
  }

  // ==========================================================================
  // Calculation Methods
  // ==========================================================================

  /**
   * Calculate success and failure rates for policies
   *
   * @param policies - Array of policies
   * @param deviceStatuses - Array of device statuses
   * @param userStatuses - Array of user statuses
   * @returns Array of success rate statistics
   */
  calculateSuccessRates(
    policies: PolicyConfiguration[],
    deviceStatuses: DeviceDeploymentStatus[],
    userStatuses: UserDeploymentStatus[]
  ): PolicySuccessRate[] {
    logger.debug('Calculating success rates');

    const rates: PolicySuccessRate[] = [];

    for (const policy of policies) {
      const policyDeviceStatuses = deviceStatuses.filter(s => s.policyId === policy.id);
      const policyUserStatuses = userStatuses.filter(s => s.policyId === policy.id);

      // Use device statuses if available, otherwise use user statuses
      const statuses = policyDeviceStatuses.length > 0 ? policyDeviceStatuses : policyUserStatuses;

      const successCount = statuses.filter(s =>
        s.status.toLowerCase() === 'success' || s.status.toLowerCase() === 'compliant'
      ).length;

      const failureCount = statuses.filter(s =>
        s.status.toLowerCase() === 'error' || s.status.toLowerCase() === 'failed' || s.status.toLowerCase() === 'noncompliant'
      ).length;

      const conflictCount = statuses.filter(s =>
        s.status.toLowerCase() === 'conflict'
      ).length;

      const pendingCount = statuses.filter(s =>
        s.status.toLowerCase() === 'pending' || s.status.toLowerCase() === 'notassigned'
      ).length;

      const notApplicableCount = statuses.filter(s =>
        s.status.toLowerCase() === 'notapplicable'
      ).length;

      const totalTargets = statuses.length;
      const successRate = totalTargets > 0 ? Math.round((successCount / totalTargets) * 100 * 100) / 100 : 0;
      const failureRate = totalTargets > 0 ? Math.round((failureCount / totalTargets) * 100 * 100) / 100 : 0;

      rates.push({
        policyId: policy.id,
        policyName: policy.displayName,
        totalTargets,
        successCount,
        failureCount,
        conflictCount,
        pendingCount,
        notApplicableCount,
        successRate,
        failureRate
      });
    }

    return rates;
  }

  /**
   * Generate summary statistics for the report
   */
  private generateSummary(
    policies: PolicyConfiguration[],
    assignments: PolicyAssignment[],
    deviceStatuses: DeviceDeploymentStatus[],
    userStatuses: UserDeploymentStatus[],
    conflicts: PolicyConflict[],
    successRates: PolicySuccessRate[]
  ): PolicyAssignmentSummary {
    logger.debug('Generating summary statistics');

    const policiesWithAssignments = policies.filter(p => p.assignmentCount > 0).length;
    const policiesWithoutAssignments = policies.length - policiesWithAssignments;

    // Calculate overall success rate
    const totalTargets = successRates.reduce((sum, rate) => sum + rate.totalTargets, 0);
    const totalSuccess = successRates.reduce((sum, rate) => sum + rate.successCount, 0);
    const totalFailure = successRates.reduce((sum, rate) => sum + rate.failureCount, 0);

    const overallSuccessRate = totalTargets > 0
      ? Math.round((totalSuccess / totalTargets) * 100 * 100) / 100
      : 0;

    const overallFailureRate = totalTargets > 0
      ? Math.round((totalFailure / totalTargets) * 100 * 100) / 100
      : 0;

    // Count policies by platform
    const policiesByPlatform: Record<string, number> = {};
    for (const policy of policies) {
      const platform = policy.platformType;
      policiesByPlatform[platform] = (policiesByPlatform[platform] || 0) + 1;
    }

    return {
      totalPolicies: policies.length,
      policiesWithAssignments,
      policiesWithoutAssignments,
      totalAssignments: assignments.length,
      totalDeviceDeployments: deviceStatuses.length,
      totalUserDeployments: userStatuses.length,
      totalConflicts: conflicts.length,
      overallSuccessRate,
      overallFailureRate,
      policiesByPlatform,
      generatedAt: new Date().toISOString()
    };
  }

  // ==========================================================================
  // Helper Methods
  // ==========================================================================

  /**
   * Determine the platform type from a policy object
   */
  private determinePlatformType(policy: any): string {
    const odataType = policy['@odata.type'];

    if (!odataType) return 'Unknown';

    if (odataType.includes('windows')) return 'Windows';
    if (odataType.includes('ios')) return 'iOS';
    if (odataType.includes('android')) return 'Android';
    if (odataType.includes('macOS')) return 'macOS';

    return 'Multi-Platform';
  }

  /**
   * Determine the technologies from a policy object
   */
  private determineTechnologies(policy: any): string {
    const odataType = policy['@odata.type'];

    if (!odataType) return 'Unknown';

    if (odataType.includes('mdm')) return 'MDM';
    if (odataType.includes('mam')) return 'MAM';

    return 'Configuration';
  }

  /**
   * Determine the target type from an assignment target
   */
  private determineTargetType(target: any): string {
    if (!target) return 'Unknown';

    const odataType = target['@odata.type'];

    if (odataType?.includes('allLicensedUsersAssignmentTarget')) return 'All Users';
    if (odataType?.includes('allDevicesAssignmentTarget')) return 'All Devices';
    if (odataType?.includes('groupAssignmentTarget')) return 'Group';
    if (odataType?.includes('exclusionGroupAssignmentTarget')) return 'Exclusion Group';

    return 'Other';
  }

  /**
   * Get group name by ID
   */
  private async getGroupName(groupId: string): Promise<string> {
    try {
      const group = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/groups/${groupId}`)
          .select('displayName')
          .get()
      );
      return group.displayName || 'Unknown Group';
    } catch (error) {
      logger.warn(`Failed to fetch group name for ${groupId}`, error);
      return 'Unknown Group';
    }
  }

  /**
   * Get policy names from policy IDs
   */
  private async getPolicyNames(policyIds: string[]): Promise<string[]> {
    const names: string[] = [];

    for (const policyId of policyIds) {
      try {
        const policy = await this.retryGraphCall(() =>
          this.graphClient
            .api(`/deviceManagement/deviceConfigurations/${policyId}`)
            .select('displayName')
            .get()
        );
        names.push(policy.displayName || 'Unknown Policy');
      } catch (error) {
        names.push('Unknown Policy');
      }
    }

    return names;
  }

  // ==========================================================================
  // Export Methods
  // ==========================================================================

  /**
   * Export all reports in multiple formats
   */
  private async exportReports(
    data: PolicyAssignmentReportData,
    outputDir: string
  ): Promise<void> {
    logger.info('Exporting reports');

    // Ensure output directory exists
    await fs.mkdir(outputDir, { recursive: true });

    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');

    // Export to JSON
    await this.exportToJSON(data, path.join(outputDir, `policy-assignment-${timestamp}.json`));

    // Export to CSV
    await this.exportToCSV(data, path.join(outputDir, `policy-assignment-${timestamp}.csv`));

    // Export to HTML
    await this.exportToHTML(data, path.join(outputDir, `policy-assignment-${timestamp}.html`));

    logger.info('Reports exported successfully');
  }

  /**
   * Export report data to JSON format
   */
  async exportToJSON(data: PolicyAssignmentReportData, filePath: string): Promise<void> {
    try {
      const json = JSON.stringify(data, null, 2);
      await fs.writeFile(filePath, json, 'utf-8');
      logger.info(`JSON report exported to ${filePath}`);
    } catch (error) {
      logger.error('Failed to export JSON report', error);
      throw error;
    }
  }

  /**
   * Export report data to CSV format
   */
  async exportToCSV(data: PolicyAssignmentReportData, filePath: string): Promise<void> {
    try {
      let csv = '';

      // Policies CSV
      csv += '=== CONFIGURATION POLICIES ===\n';
      csv += 'Policy ID,Policy Name,Description,Platform,Created Date,Modified Date,Assignments,Device Statuses,User Statuses\n';

      for (const policy of data.policies) {
        csv += [
          this.escapeCsv(policy.id),
          this.escapeCsv(policy.displayName),
          this.escapeCsv(policy.description),
          this.escapeCsv(policy.platformType),
          this.escapeCsv(new Date(policy.createdDateTime).toLocaleDateString()),
          this.escapeCsv(new Date(policy.lastModifiedDateTime).toLocaleDateString()),
          policy.assignmentCount,
          policy.deviceStatusCount,
          policy.userStatusCount
        ].join(',') + '\n';
      }

      csv += '\n=== ASSIGNMENTS ===\n';
      csv += 'Policy Name,Target Type,Target Group,Intent,Source\n';

      for (const assignment of data.assignments) {
        csv += [
          this.escapeCsv(assignment.policyName),
          this.escapeCsv(assignment.targetType),
          this.escapeCsv(assignment.targetGroupName || 'N/A'),
          this.escapeCsv(assignment.intent || 'N/A'),
          this.escapeCsv(assignment.source || 'N/A')
        ].join(',') + '\n';
      }

      if (data.successRates.length > 0) {
        csv += '\n=== SUCCESS RATES ===\n';
        csv += 'Policy Name,Total Targets,Success,Failure,Conflicts,Pending,Not Applicable,Success Rate %,Failure Rate %\n';

        for (const rate of data.successRates) {
          csv += [
            this.escapeCsv(rate.policyName),
            rate.totalTargets,
            rate.successCount,
            rate.failureCount,
            rate.conflictCount,
            rate.pendingCount,
            rate.notApplicableCount,
            rate.successRate,
            rate.failureRate
          ].join(',') + '\n';
        }
      }

      await fs.writeFile(filePath, csv, 'utf-8');
      logger.info(`CSV report exported to ${filePath}`);
    } catch (error) {
      logger.error('Failed to export CSV report', error);
      throw error;
    }
  }

  /**
   * Export report data to HTML format
   */
  async exportToHTML(data: PolicyAssignmentReportData, filePath: string): Promise<void> {
    try {
      const html = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Policy Assignment Report</title>
  <style>
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f5f5f5; padding: 20px; }
    .container { max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    h1 { color: #0078d4; margin-bottom: 10px; font-size: 28px; }
    h2 { color: #333; margin-top: 30px; margin-bottom: 15px; font-size: 22px; border-bottom: 2px solid #0078d4; padding-bottom: 8px; }
    h3 { color: #555; margin-top: 20px; margin-bottom: 10px; font-size: 18px; }
    .metadata { color: #666; margin-bottom: 30px; font-size: 14px; }
    .summary-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(250px, 1fr)); gap: 20px; margin: 20px 0; }
    .summary-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 20px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .summary-card.success { background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); }
    .summary-card.warning { background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); }
    .summary-card.info { background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%); }
    .summary-card h3 { color: white; font-size: 14px; margin-bottom: 10px; opacity: 0.9; }
    .summary-card .value { font-size: 32px; font-weight: bold; }
    .summary-card .label { font-size: 12px; opacity: 0.8; margin-top: 5px; }
    table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 14px; }
    thead { background: #0078d4; color: white; }
    th { padding: 12px; text-align: left; font-weight: 600; }
    td { padding: 12px; border-bottom: 1px solid #e0e0e0; }
    tbody tr:hover { background: #f9f9f9; }
    .status-badge { display: inline-block; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: 600; }
    .status-success { background: #d4edda; color: #155724; }
    .status-error { background: #f8d7da; color: #721c24; }
    .status-warning { background: #fff3cd; color: #856404; }
    .status-info { background: #d1ecf1; color: #0c5460; }
    .platform-badge { background: #e7f3ff; color: #0078d4; padding: 4px 8px; border-radius: 4px; font-size: 12px; }
    .progress-bar { width: 100%; height: 20px; background: #e0e0e0; border-radius: 10px; overflow: hidden; }
    .progress-fill { height: 100%; background: linear-gradient(90deg, #11998e 0%, #38ef7d 100%); transition: width 0.3s; }
  </style>
</head>
<body>
  <div class="container">
    <h1>Policy Assignment Report</h1>
    <div class="metadata">
      Generated: ${new Date(data.summary.generatedAt).toLocaleString()}<br>
      Total Policies: ${data.summary.totalPolicies}
    </div>

    <h2>Summary</h2>
    <div class="summary-grid">
      <div class="summary-card">
        <h3>Total Policies</h3>
        <div class="value">${data.summary.totalPolicies}</div>
        <div class="label">${data.summary.policiesWithAssignments} with assignments</div>
      </div>
      <div class="summary-card success">
        <h3>Overall Success Rate</h3>
        <div class="value">${data.summary.overallSuccessRate}%</div>
        <div class="label">Deployment success</div>
      </div>
      <div class="summary-card warning">
        <h3>Total Conflicts</h3>
        <div class="value">${data.summary.totalConflicts}</div>
        <div class="label">Policy conflicts detected</div>
      </div>
      <div class="summary-card info">
        <h3>Total Assignments</h3>
        <div class="value">${data.summary.totalAssignments}</div>
        <div class="label">Across all policies</div>
      </div>
    </div>

    <h2>Configuration Policies</h2>
    <table>
      <thead>
        <tr>
          <th>Policy Name</th>
          <th>Platform</th>
          <th>Assignments</th>
          <th>Device Statuses</th>
          <th>Created</th>
          <th>Modified</th>
        </tr>
      </thead>
      <tbody>
        ${data.policies.map(p => `
          <tr>
            <td><strong>${this.escapeHtml(p.displayName)}</strong><br><small>${this.escapeHtml(p.description)}</small></td>
            <td><span class="platform-badge">${this.escapeHtml(p.platformType)}</span></td>
            <td>${p.assignmentCount}</td>
            <td>${p.deviceStatusCount}</td>
            <td>${new Date(p.createdDateTime).toLocaleDateString()}</td>
            <td>${new Date(p.lastModifiedDateTime).toLocaleDateString()}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>

    ${data.successRates.length > 0 ? `
    <h2>Success Rates</h2>
    <table>
      <thead>
        <tr>
          <th>Policy Name</th>
          <th>Total Targets</th>
          <th>Success</th>
          <th>Failure</th>
          <th>Success Rate</th>
          <th>Progress</th>
        </tr>
      </thead>
      <tbody>
        ${data.successRates.map(r => `
          <tr>
            <td>${this.escapeHtml(r.policyName)}</td>
            <td>${r.totalTargets}</td>
            <td><span class="status-badge status-success">${r.successCount}</span></td>
            <td><span class="status-badge status-error">${r.failureCount}</span></td>
            <td><strong>${r.successRate}%</strong></td>
            <td>
              <div class="progress-bar">
                <div class="progress-fill" style="width: ${r.successRate}%"></div>
              </div>
            </td>
          </tr>
        `).join('')}
      </tbody>
    </table>
    ` : ''}

    ${data.conflicts.length > 0 ? `
    <h2>Policy Conflicts</h2>
    <table>
      <thead>
        <tr>
          <th>Conflict ID</th>
          <th>Conflicting Policies</th>
          <th>Affected Devices</th>
          <th>Description</th>
        </tr>
      </thead>
      <tbody>
        ${data.conflicts.map(c => `
          <tr>
            <td>${this.escapeHtml(c.id)}</td>
            <td>${c.conflictingPolicyNames.map(n => this.escapeHtml(n)).join(', ')}</td>
            <td>${c.affectedDevicesCount}</td>
            <td>${this.escapeHtml(c.description || 'N/A')}</td>
          </tr>
        `).join('')}
      </tbody>
    </table>
    ` : ''}
  </div>
</body>
</html>
      `;

      await fs.writeFile(filePath, html, 'utf-8');
      logger.info(`HTML report exported to ${filePath}`);
    } catch (error) {
      logger.error('Failed to export HTML report', error);
      throw error;
    }
  }

  /**
   * Escape CSV value
   */
  private escapeCsv(value: string | number | undefined): string {
    if (value === undefined || value === null) return '';
    const str = String(value);
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }

  /**
   * Escape HTML special characters
   */
  private escapeHtml(text: string): string {
    const map: Record<string, string> = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
  }
}

// ============================================================================
// Export
// ============================================================================

export default PolicyAssignmentReport;
