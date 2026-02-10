/**
 * Configuration Manager Device Action Reports for Intune
 *
 * This module replicates Configuration Manager reports for device actions including:
 * - Pending retire and wipe requests
 * - Recently enrolled and assigned mobile devices
 * - Recently wiped mobile devices
 *
 * Uses Microsoft Graph API to track device lifecycle actions.
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { Parser } from 'json2csv';

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Device action types
 */
export type DeviceActionType =
  | 'retire'
  | 'wipe'
  | 'remoteLock'
  | 'resetPasscode'
  | 'rebootNow'
  | 'delete'
  | 'factoryReset';

/**
 * Device action status
 */
export type DeviceActionStatus =
  | 'pending'
  | 'inProgress'
  | 'completed'
  | 'failed'
  | 'notSupported'
  | 'cancelled';

/**
 * Device ownership type
 */
export type DeviceOwnershipType = 'company' | 'personal' | 'unknown';

/**
 * Pending device action
 */
export interface PendingDeviceAction {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  operatingSystem: string;
  osVersion: string;
  actionType: DeviceActionType;
  actionStatus: DeviceActionStatus;
  actionRequestedDateTime: string;
  actionInitiatedBy?: string;
  lastSyncDateTime: string;
  retireAfterDateTime?: string;
  managedDeviceOwnerType: DeviceOwnershipType;
  estimatedCompletionTime?: string;
  actionDetails?: string;
}

/**
 * Recently enrolled device
 */
export interface RecentlyEnrolledDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  operatingSystem: string;
  osVersion: string;
  manufacturer: string;
  model: string;
  serialNumber: string;
  enrolledDateTime: string;
  lastSyncDateTime: string;
  complianceState: string;
  managementAgent: string;
  managedDeviceOwnerType: DeviceOwnershipType;
  enrollmentType: string;
  azureADDeviceId: string;
  daysSinceEnrollment: number;
  isAssigned: boolean;
  assignedUsers?: string[];
  assignedGroups?: string[];
}

/**
 * Recently wiped device
 */
export interface RecentlyWipedDevice {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  operatingSystem: string;
  osVersion: string;
  wipeRequestedDateTime: string;
  wipeCompletedDateTime: string;
  wipeInitiatedBy?: string;
  wipeType: 'wipe' | 'retire' | 'factoryReset';
  managedDeviceOwnerType: DeviceOwnershipType;
  wipeDurationMinutes: number;
  wipeStatus: 'success' | 'failed' | 'partial';
  wipeErrorCode?: string;
  wipeErrorMessage?: string;
  lastKnownLocation?: string;
}

/**
 * Report options for filtering
 */
export interface DeviceActionReportOptions {
  dateRange?: {
    startDate: Date;
    endDate: Date;
  };
  actionType?: DeviceActionType[];
  status?: DeviceActionStatus[];
  ownershipType?: DeviceOwnershipType[];
  operatingSystem?: string[];
  top?: number;
}

/**
 * Report summary statistics
 */
export interface DeviceActionSummary {
  totalPendingActions: number;
  totalRecentEnrollments: number;
  totalRecentWipes: number;
  actionsByType: Record<DeviceActionType, number>;
  actionsByStatus: Record<DeviceActionStatus, number>;
  enrollmentsByOS: Record<string, number>;
  wipesByOS: Record<string, number>;
  averageWipeDurationMinutes: number;
  wipeSuccessRate: number;
  generatedAt: string;
}

/**
 * Complete device action report
 */
export interface DeviceActionReport {
  summary: DeviceActionSummary;
  pendingActions: PendingDeviceAction[];
  recentEnrollments: RecentlyEnrolledDevice[];
  recentWipes: RecentlyWipedDevice[];
  reportPeriod: {
    startDate: string;
    endDate: string;
  };
  metadata: {
    reportName: string;
    generatedAt: string;
    generatedBy: string;
    recordCount: number;
    parameters?: Record<string, any>;
  };
}

// ============================================================================
// Device Action Reports Class
// ============================================================================

/**
 * Device Action Reports - Configuration Manager equivalent reports
 */
export class DeviceActionReports {
  private graphClient: Client;

  constructor(graphClient: Client) {
    this.graphClient = graphClient;
  }

  // ==========================================================================
  // Report #31: Pending Retire and Wipe Requests
  // ==========================================================================

  /**
   * Get pending retire and wipe requests for mobile devices
   * Displays all devices with pending wipe/retire actions
   */
  async getPendingRetireWipeRequests(
    options: DeviceActionReportOptions = {}
  ): Promise<PendingDeviceAction[]> {
    try {
      // Fetch all managed devices
      const devices = await this.fetchManagedDevices(options);

      const pendingActions: PendingDeviceAction[] = [];

      for (const device of devices) {
        // Check for retire after date
        if (device.retireAfterDateTime) {
          const retireDate = new Date(device.retireAfterDateTime);
          const now = new Date();

          if (retireDate <= now) {
            pendingActions.push(this.mapToPendingAction(device, 'retire'));
          }
        }

        // Check device actions for pending operations
        const deviceActions = await this.getDeviceActions(device.id);

        for (const action of deviceActions) {
          if (this.isActionPending(action)) {
            const actionType = this.mapActionType(action.actionName);
            if (actionType === 'retire' || actionType === 'wipe') {
              pendingActions.push(
                this.mapToPendingActionFromDeviceAction(device, action)
              );
            }
          }
        }
      }

      // Apply filters
      return this.filterPendingActions(pendingActions, options);
    } catch (error) {
      console.error('Error fetching pending retire/wipe requests:', error);
      throw new Error(`Failed to fetch pending actions: ${error}`);
    }
  }

  // ==========================================================================
  // Report #32: Recently Enrolled and Assigned Mobile Devices
  // ==========================================================================

  /**
   * Get recently enrolled mobile devices with assignment information
   * Default: last 7 days, configurable
   */
  async getRecentlyEnrolledDevices(
    options: DeviceActionReportOptions = {}
  ): Promise<RecentlyEnrolledDevice[]> {
    try {
      // Default to last 7 days if no date range specified
      const dateRange = options.dateRange || {
        startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
        endDate: new Date(),
      };

      // Fetch devices enrolled in the date range
      const filter = `enrolledDateTime ge ${dateRange.startDate.toISOString()}`;
      const devices = await this.fetchManagedDevices({ ...options, dateRange });

      const enrolledDevices: RecentlyEnrolledDevice[] = [];

      for (const device of devices) {
        if (!device.enrolledDateTime) continue;

        const enrolledDate = new Date(device.enrolledDateTime);
        if (enrolledDate >= dateRange.startDate && enrolledDate <= dateRange.endDate) {
          // Get assignment information
          const assignments = await this.getDeviceAssignments(device.id);

          enrolledDevices.push({
            deviceId: device.id,
            deviceName: device.deviceName || 'Unknown',
            userPrincipalName: device.userPrincipalName || 'N/A',
            userDisplayName: device.userDisplayName || 'N/A',
            operatingSystem: device.operatingSystem || 'Unknown',
            osVersion: device.osVersion || 'Unknown',
            manufacturer: device.manufacturer || 'Unknown',
            model: device.model || 'Unknown',
            serialNumber: device.serialNumber || 'N/A',
            enrolledDateTime: device.enrolledDateTime,
            lastSyncDateTime: device.lastSyncDateTime || 'Never',
            complianceState: device.complianceState || 'Unknown',
            managementAgent: device.managementAgent || 'Unknown',
            managedDeviceOwnerType: this.mapOwnershipType(device.managedDeviceOwnerType),
            enrollmentType: device.deviceEnrollmentType || 'Unknown',
            azureADDeviceId: device.azureADDeviceId || 'N/A',
            daysSinceEnrollment: this.calculateDaysSince(device.enrolledDateTime),
            isAssigned: assignments.users.length > 0 || assignments.groups.length > 0,
            assignedUsers: assignments.users,
            assignedGroups: assignments.groups,
          });
        }
      }

      return this.filterRecentEnrollments(enrolledDevices, options);
    } catch (error) {
      console.error('Error fetching recently enrolled devices:', error);
      throw new Error(`Failed to fetch enrolled devices: ${error}`);
    }
  }

  // ==========================================================================
  // Report #33: Recently Wiped Mobile Devices
  // ==========================================================================

  /**
   * Get recently wiped mobile devices
   * Shows completed wipe/retire operations
   */
  async getRecentlyWipedDevices(
    options: DeviceActionReportOptions = {}
  ): Promise<RecentlyWipedDevice[]> {
    try {
      // Default to last 30 days
      const dateRange = options.dateRange || {
        startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
        endDate: new Date(),
      };

      const wipedDevices: RecentlyWipedDevice[] = [];

      // Fetch all devices to check their action history
      const devices = await this.fetchManagedDevices(options);

      for (const device of devices) {
        const deviceActions = await this.getDeviceActions(device.id);

        for (const action of deviceActions) {
          const actionType = this.mapActionType(action.actionName);

          if (
            (actionType === 'wipe' || actionType === 'retire' || actionType === 'factoryReset') &&
            action.actionState === 'done'
          ) {
            const completedDate = new Date(action.lastUpdatedDateTime);

            if (completedDate >= dateRange.startDate && completedDate <= dateRange.endDate) {
              wipedDevices.push(
                this.mapToWipedDevice(device, action, dateRange)
              );
            }
          }
        }
      }

      return this.filterRecentWipes(wipedDevices, options);
    } catch (error) {
      console.error('Error fetching recently wiped devices:', error);
      throw new Error(`Failed to fetch wiped devices: ${error}`);
    }
  }

  // ==========================================================================
  // Comprehensive Report Generation
  // ==========================================================================

  /**
   * Generate comprehensive device action report
   * Combines all three report types with summary statistics
   */
  async generateDeviceActionReport(
    options: DeviceActionReportOptions = {}
  ): Promise<DeviceActionReport> {
    try {
      const dateRange = options.dateRange || {
        startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
        endDate: new Date(),
      };

      // Fetch all reports in parallel for efficiency
      const [pendingActions, recentEnrollments, recentWipes] = await Promise.all([
        this.getPendingRetireWipeRequests(options),
        this.getRecentlyEnrolledDevices(options),
        this.getRecentlyWipedDevices(options),
      ]);

      // Generate summary statistics
      const summary = this.generateSummary(
        pendingActions,
        recentEnrollments,
        recentWipes
      );

      return {
        summary,
        pendingActions,
        recentEnrollments,
        recentWipes,
        reportPeriod: {
          startDate: dateRange.startDate.toISOString(),
          endDate: dateRange.endDate.toISOString(),
        },
        metadata: {
          reportName: 'Device Action Report',
          generatedAt: new Date().toISOString(),
          generatedBy: 'Intune Reporting - Device Action Reports',
          recordCount:
            pendingActions.length + recentEnrollments.length + recentWipes.length,
          parameters: options,
        },
      };
    } catch (error) {
      console.error('Error generating device action report:', error);
      throw new Error(`Failed to generate report: ${error}`);
    }
  }

  // ==========================================================================
  // Export Functions
  // ==========================================================================

  /**
   * Export report to JSON format
   */
  exportToJSON(report: DeviceActionReport): string {
    return JSON.stringify(report, null, 2);
  }

  /**
   * Export pending actions to CSV
   */
  exportPendingActionsToCSV(actions: PendingDeviceAction[]): string {
    if (actions.length === 0) {
      return 'No pending actions found';
    }

    const parser = new Parser({
      fields: [
        'deviceName',
        'userPrincipalName',
        'operatingSystem',
        'actionType',
        'actionStatus',
        'actionRequestedDateTime',
        'retireAfterDateTime',
        'managedDeviceOwnerType',
        'lastSyncDateTime',
      ],
    });

    return parser.parse(actions);
  }

  /**
   * Export recent enrollments to CSV
   */
  exportEnrollmentsToCSV(enrollments: RecentlyEnrolledDevice[]): string {
    if (enrollments.length === 0) {
      return 'No recent enrollments found';
    }

    const parser = new Parser({
      fields: [
        'deviceName',
        'userPrincipalName',
        'operatingSystem',
        'manufacturer',
        'model',
        'enrolledDateTime',
        'daysSinceEnrollment',
        'complianceState',
        'managedDeviceOwnerType',
        'isAssigned',
      ],
    });

    return parser.parse(enrollments);
  }

  /**
   * Export recent wipes to CSV
   */
  exportWipesToCSV(wipes: RecentlyWipedDevice[]): string {
    if (wipes.length === 0) {
      return 'No recent wipes found';
    }

    const parser = new Parser({
      fields: [
        'deviceName',
        'userPrincipalName',
        'operatingSystem',
        'wipeType',
        'wipeRequestedDateTime',
        'wipeCompletedDateTime',
        'wipeDurationMinutes',
        'wipeStatus',
        'managedDeviceOwnerType',
      ],
    });

    return parser.parse(wipes);
  }

  /**
   * Export complete report to HTML
   */
  exportToHTML(report: DeviceActionReport): string {
    return `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Device Action Report</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
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
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        h2 {
            color: #106ebe;
            margin-top: 30px;
        }
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }
        .summary-card {
            background: #e3f2fd;
            padding: 15px;
            border-radius: 4px;
            border-left: 4px solid #0078d4;
        }
        .summary-card h3 {
            margin: 0 0 5px 0;
            font-size: 14px;
            color: #666;
        }
        .summary-card p {
            margin: 0;
            font-size: 24px;
            font-weight: bold;
            color: #0078d4;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        th {
            background: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover {
            background: #f5f5f5;
        }
        .status-pending { color: #ff9800; font-weight: bold; }
        .status-completed { color: #4caf50; font-weight: bold; }
        .status-failed { color: #f44336; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Device Action Report</h1>
        <p><strong>Generated:</strong> ${report.metadata.generatedAt}</p>
        <p><strong>Period:</strong> ${report.reportPeriod.startDate} to ${report.reportPeriod.endDate}</p>

        <div class="summary">
            <div class="summary-card">
                <h3>Pending Actions</h3>
                <p>${report.summary.totalPendingActions}</p>
            </div>
            <div class="summary-card">
                <h3>Recent Enrollments</h3>
                <p>${report.summary.totalRecentEnrollments}</p>
            </div>
            <div class="summary-card">
                <h3>Recent Wipes</h3>
                <p>${report.summary.totalRecentWipes}</p>
            </div>
            <div class="summary-card">
                <h3>Wipe Success Rate</h3>
                <p>${report.summary.wipeSuccessRate.toFixed(1)}%</p>
            </div>
        </div>

        <h2>Pending Retire and Wipe Requests</h2>
        ${this.generatePendingActionsTable(report.pendingActions)}

        <h2>Recently Enrolled Devices</h2>
        ${this.generateEnrollmentsTable(report.recentEnrollments)}

        <h2>Recently Wiped Devices</h2>
        ${this.generateWipesTable(report.recentWipes)}
    </div>
</body>
</html>
    `;
  }

  // ==========================================================================
  // Private Helper Methods
  // ==========================================================================

  /**
   * Fetch managed devices from Graph API
   */
  private async fetchManagedDevices(
    options: DeviceActionReportOptions
  ): Promise<any[]> {
    try {
      let request = this.graphClient
        .api('/deviceManagement/managedDevices')
        .select([
          'id',
          'deviceName',
          'userPrincipalName',
          'userDisplayName',
          'operatingSystem',
          'osVersion',
          'manufacturer',
          'model',
          'serialNumber',
          'enrolledDateTime',
          'lastSyncDateTime',
          'complianceState',
          'managementAgent',
          'managedDeviceOwnerType',
          'deviceEnrollmentType',
          'azureADDeviceId',
          'retireAfterDateTime',
        ]);

      // Apply filters
      const filters: string[] = [];

      if (options.ownershipType && options.ownershipType.length > 0) {
        const ownershipFilters = options.ownershipType.map(
          (type) => `managedDeviceOwnerType eq '${type}'`
        );
        filters.push(`(${ownershipFilters.join(' or ')})`);
      }

      if (options.operatingSystem && options.operatingSystem.length > 0) {
        const osFilters = options.operatingSystem.map(
          (os) => `operatingSystem eq '${os}'`
        );
        filters.push(`(${osFilters.join(' or ')})`);
      }

      if (filters.length > 0) {
        request = request.filter(filters.join(' and '));
      }

      if (options.top) {
        request = request.top(options.top);
      } else {
        request = request.top(999);
      }

      const response = await request.get();
      return this.getAllPages(response);
    } catch (error) {
      console.error('Error fetching managed devices:', error);
      throw error;
    }
  }

  /**
   * Get all pages from paginated Graph API response
   */
  private async getAllPages(initialResponse: any): Promise<any[]> {
    let results = initialResponse.value || [];
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
   * Get device actions for a specific device
   */
  private async getDeviceActions(deviceId: string): Promise<any[]> {
    try {
      const response = await this.graphClient
        .api(`/deviceManagement/managedDevices/${deviceId}/deviceManagementTroubleshootingEvents`)
        .top(100)
        .get();

      return response.value || [];
    } catch (error) {
      // Device actions might not be available for all devices
      return [];
    }
  }

  /**
   * Get device assignments (users and groups)
   */
  private async getDeviceAssignments(
    deviceId: string
  ): Promise<{ users: string[]; groups: string[] }> {
    try {
      // In a real implementation, this would fetch group memberships
      // For now, return empty arrays
      return { users: [], groups: [] };
    } catch (error) {
      return { users: [], groups: [] };
    }
  }

  /**
   * Check if action is pending
   */
  private isActionPending(action: any): boolean {
    const pendingStates = ['pending', 'notStarted', 'actionInProgress'];
    return pendingStates.includes(action.actionState?.toLowerCase());
  }

  /**
   * Map Graph API action name to DeviceActionType
   */
  private mapActionType(actionName: string): DeviceActionType {
    const actionMap: Record<string, DeviceActionType> = {
      retire: 'retire',
      wipe: 'wipe',
      remoteLock: 'remoteLock',
      resetPasscode: 'resetPasscode',
      rebootNow: 'rebootNow',
      delete: 'delete',
      factoryReset: 'factoryReset',
    };

    return actionMap[actionName?.toLowerCase()] || 'wipe';
  }

  /**
   * Map device ownership type
   */
  private mapOwnershipType(ownerType: string): DeviceOwnershipType {
    const lowerType = ownerType?.toLowerCase();
    if (lowerType === 'company' || lowerType === 'corporate') return 'company';
    if (lowerType === 'personal' || lowerType === 'byod') return 'personal';
    return 'unknown';
  }

  /**
   * Map device to pending action
   */
  private mapToPendingAction(
    device: any,
    actionType: DeviceActionType
  ): PendingDeviceAction {
    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'N/A',
      userDisplayName: device.userDisplayName || 'N/A',
      operatingSystem: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      actionType,
      actionStatus: 'pending',
      actionRequestedDateTime: device.retireAfterDateTime || new Date().toISOString(),
      lastSyncDateTime: device.lastSyncDateTime || 'Never',
      retireAfterDateTime: device.retireAfterDateTime,
      managedDeviceOwnerType: this.mapOwnershipType(device.managedDeviceOwnerType),
    };
  }

  /**
   * Map device action to pending action
   */
  private mapToPendingActionFromDeviceAction(
    device: any,
    action: any
  ): PendingDeviceAction {
    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'N/A',
      userDisplayName: device.userDisplayName || 'N/A',
      operatingSystem: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      actionType: this.mapActionType(action.actionName),
      actionStatus: this.mapActionStatus(action.actionState),
      actionRequestedDateTime: action.actionDateTime || new Date().toISOString(),
      lastSyncDateTime: device.lastSyncDateTime || 'Never',
      managedDeviceOwnerType: this.mapOwnershipType(device.managedDeviceOwnerType),
      actionDetails: action.actionDetails,
    };
  }

  /**
   * Map action state to status
   */
  private mapActionStatus(actionState: string): DeviceActionStatus {
    const stateMap: Record<string, DeviceActionStatus> = {
      pending: 'pending',
      notStarted: 'pending',
      actionInProgress: 'inProgress',
      done: 'completed',
      failed: 'failed',
      notSupported: 'notSupported',
      cancelled: 'cancelled',
    };

    return stateMap[actionState?.toLowerCase()] || 'pending';
  }

  /**
   * Map device action to wiped device
   */
  private mapToWipedDevice(
    device: any,
    action: any,
    dateRange: { startDate: Date; endDate: Date }
  ): RecentlyWipedDevice {
    const requestedDate = new Date(action.actionDateTime);
    const completedDate = new Date(action.lastUpdatedDateTime);
    const durationMinutes = Math.floor(
      (completedDate.getTime() - requestedDate.getTime()) / 1000 / 60
    );

    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'N/A',
      userDisplayName: device.userDisplayName || 'N/A',
      operatingSystem: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      wipeRequestedDateTime: action.actionDateTime,
      wipeCompletedDateTime: action.lastUpdatedDateTime,
      wipeType: this.mapActionType(action.actionName) as any,
      managedDeviceOwnerType: this.mapOwnershipType(device.managedDeviceOwnerType),
      wipeDurationMinutes: durationMinutes > 0 ? durationMinutes : 0,
      wipeStatus: action.actionState === 'done' ? 'success' : 'failed',
      wipeErrorCode: action.errorCode,
      wipeErrorMessage: action.errorMessage,
    };
  }

  /**
   * Calculate days since a date
   */
  private calculateDaysSince(dateString: string): number {
    const date = new Date(dateString);
    const now = new Date();
    const diff = now.getTime() - date.getTime();
    return Math.floor(diff / (1000 * 60 * 60 * 24));
  }

  /**
   * Filter pending actions based on options
   */
  private filterPendingActions(
    actions: PendingDeviceAction[],
    options: DeviceActionReportOptions
  ): PendingDeviceAction[] {
    let filtered = actions;

    if (options.actionType && options.actionType.length > 0) {
      filtered = filtered.filter((a) => options.actionType!.includes(a.actionType));
    }

    if (options.status && options.status.length > 0) {
      filtered = filtered.filter((a) => options.status!.includes(a.actionStatus));
    }

    if (options.top) {
      filtered = filtered.slice(0, options.top);
    }

    return filtered;
  }

  /**
   * Filter recent enrollments based on options
   */
  private filterRecentEnrollments(
    enrollments: RecentlyEnrolledDevice[],
    options: DeviceActionReportOptions
  ): RecentlyEnrolledDevice[] {
    let filtered = enrollments;

    if (options.top) {
      filtered = filtered.slice(0, options.top);
    }

    return filtered.sort(
      (a, b) =>
        new Date(b.enrolledDateTime).getTime() -
        new Date(a.enrolledDateTime).getTime()
    );
  }

  /**
   * Filter recent wipes based on options
   */
  private filterRecentWipes(
    wipes: RecentlyWipedDevice[],
    options: DeviceActionReportOptions
  ): RecentlyWipedDevice[] {
    let filtered = wipes;

    if (options.top) {
      filtered = filtered.slice(0, options.top);
    }

    return filtered.sort(
      (a, b) =>
        new Date(b.wipeCompletedDateTime).getTime() -
        new Date(a.wipeCompletedDateTime).getTime()
    );
  }

  /**
   * Generate summary statistics
   */
  private generateSummary(
    pendingActions: PendingDeviceAction[],
    recentEnrollments: RecentlyEnrolledDevice[],
    recentWipes: RecentlyWipedDevice[]
  ): DeviceActionSummary {
    // Count actions by type
    const actionsByType = {} as Record<DeviceActionType, number>;
    pendingActions.forEach((action) => {
      actionsByType[action.actionType] = (actionsByType[action.actionType] || 0) + 1;
    });

    // Count actions by status
    const actionsByStatus = {} as Record<DeviceActionStatus, number>;
    pendingActions.forEach((action) => {
      actionsByStatus[action.actionStatus] =
        (actionsByStatus[action.actionStatus] || 0) + 1;
    });

    // Count enrollments by OS
    const enrollmentsByOS: Record<string, number> = {};
    recentEnrollments.forEach((device) => {
      enrollmentsByOS[device.operatingSystem] =
        (enrollmentsByOS[device.operatingSystem] || 0) + 1;
    });

    // Count wipes by OS
    const wipesByOS: Record<string, number> = {};
    recentWipes.forEach((device) => {
      wipesByOS[device.operatingSystem] = (wipesByOS[device.operatingSystem] || 0) + 1;
    });

    // Calculate average wipe duration
    const totalDuration = recentWipes.reduce(
      (sum, wipe) => sum + wipe.wipeDurationMinutes,
      0
    );
    const averageWipeDurationMinutes =
      recentWipes.length > 0 ? totalDuration / recentWipes.length : 0;

    // Calculate wipe success rate
    const successfulWipes = recentWipes.filter((w) => w.wipeStatus === 'success').length;
    const wipeSuccessRate =
      recentWipes.length > 0 ? (successfulWipes / recentWipes.length) * 100 : 0;

    return {
      totalPendingActions: pendingActions.length,
      totalRecentEnrollments: recentEnrollments.length,
      totalRecentWipes: recentWipes.length,
      actionsByType,
      actionsByStatus,
      enrollmentsByOS,
      wipesByOS,
      averageWipeDurationMinutes,
      wipeSuccessRate,
      generatedAt: new Date().toISOString(),
    };
  }

  /**
   * Generate HTML table for pending actions
   */
  private generatePendingActionsTable(actions: PendingDeviceAction[]): string {
    if (actions.length === 0) {
      return '<p>No pending actions found.</p>';
    }

    let html = '<table><thead><tr>';
    html += '<th>Device Name</th>';
    html += '<th>User</th>';
    html += '<th>OS</th>';
    html += '<th>Action Type</th>';
    html += '<th>Status</th>';
    html += '<th>Requested</th>';
    html += '<th>Ownership</th>';
    html += '</tr></thead><tbody>';

    actions.forEach((action) => {
      html += '<tr>';
      html += `<td>${action.deviceName}</td>`;
      html += `<td>${action.userPrincipalName}</td>`;
      html += `<td>${action.operatingSystem}</td>`;
      html += `<td>${action.actionType}</td>`;
      html += `<td class="status-${action.actionStatus}">${action.actionStatus}</td>`;
      html += `<td>${new Date(action.actionRequestedDateTime).toLocaleString()}</td>`;
      html += `<td>${action.managedDeviceOwnerType}</td>`;
      html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
  }

  /**
   * Generate HTML table for enrollments
   */
  private generateEnrollmentsTable(enrollments: RecentlyEnrolledDevice[]): string {
    if (enrollments.length === 0) {
      return '<p>No recent enrollments found.</p>';
    }

    let html = '<table><thead><tr>';
    html += '<th>Device Name</th>';
    html += '<th>User</th>';
    html += '<th>OS</th>';
    html += '<th>Model</th>';
    html += '<th>Enrolled</th>';
    html += '<th>Days</th>';
    html += '<th>Assigned</th>';
    html += '</tr></thead><tbody>';

    enrollments.forEach((device) => {
      html += '<tr>';
      html += `<td>${device.deviceName}</td>`;
      html += `<td>${device.userPrincipalName}</td>`;
      html += `<td>${device.operatingSystem}</td>`;
      html += `<td>${device.manufacturer} ${device.model}</td>`;
      html += `<td>${new Date(device.enrolledDateTime).toLocaleString()}</td>`;
      html += `<td>${device.daysSinceEnrollment}</td>`;
      html += `<td>${device.isAssigned ? 'Yes' : 'No'}</td>`;
      html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
  }

  /**
   * Generate HTML table for wipes
   */
  private generateWipesTable(wipes: RecentlyWipedDevice[]): string {
    if (wipes.length === 0) {
      return '<p>No recent wipes found.</p>';
    }

    let html = '<table><thead><tr>';
    html += '<th>Device Name</th>';
    html += '<th>User</th>';
    html += '<th>OS</th>';
    html += '<th>Wipe Type</th>';
    html += '<th>Completed</th>';
    html += '<th>Duration (min)</th>';
    html += '<th>Status</th>';
    html += '</tr></thead><tbody>';

    wipes.forEach((wipe) => {
      html += '<tr>';
      html += `<td>${wipe.deviceName}</td>`;
      html += `<td>${wipe.userPrincipalName}</td>`;
      html += `<td>${wipe.operatingSystem}</td>`;
      html += `<td>${wipe.wipeType}</td>`;
      html += `<td>${new Date(wipe.wipeCompletedDateTime).toLocaleString()}</td>`;
      html += `<td>${wipe.wipeDurationMinutes}</td>`;
      html += `<td class="status-${wipe.wipeStatus}">${wipe.wipeStatus}</td>`;
      html += '</tr>';
    });

    html += '</tbody></table>';
    return html;
  }
}

// ============================================================================
// Utility Functions
// ============================================================================

/**
 * Create DeviceActionReports instance from Graph client
 */
export function createDeviceActionReports(graphClient: Client): DeviceActionReports {
  return new DeviceActionReports(graphClient);
}

/**
 * Quick report generation with default options
 */
export async function generateQuickReport(
  graphClient: Client,
  days: number = 7
): Promise<DeviceActionReport> {
  const reports = new DeviceActionReports(graphClient);
  return reports.generateDeviceActionReport({
    dateRange: {
      startDate: new Date(Date.now() - days * 24 * 60 * 60 * 1000),
      endDate: new Date(),
    },
  });
}

// Export default
export default DeviceActionReports;
