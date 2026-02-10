/**
 * Application Deployment Reporting Module for Microsoft Intune
 *
 * This module provides comprehensive application deployment reporting capabilities using
 * Microsoft Graph API, including:
 * 1. All managed apps (iOS, Android, Windows, macOS, Web)
 * 2. App installation status per app and per device
 * 3. App deployment details per device
 * 4. App assignment details (groups, users, devices)
 * 5. Failed installations with error details
 * 6. App update compliance and version tracking
 * 7. Export methods for different formats (JSON, CSV, HTML)
 *
 * @module app-deployment-report
 */

import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';
import { OutputFormatter } from '../formatters/output-formatter';

const logger = Logger.getInstance();

// ===========================
// Type Definitions
// ===========================

/**
 * Platform types for applications
 */
export enum AppPlatform {
  IOS = 'iOS',
  ANDROID = 'Android',
  WINDOWS = 'Windows',
  MACOS = 'macOS',
  WEB = 'Web',
  UNKNOWN = 'Unknown'
}

/**
 * App installation state
 */
export enum InstallState {
  INSTALLED = 'installed',
  FAILED = 'failed',
  NOT_INSTALLED = 'notInstalled',
  AVAILABLE = 'available',
  PENDING = 'pending',
  UNINSTALLING = 'uninstalling',
  DOWNLOADING = 'downloading',
  UNKNOWN = 'unknown'
}

/**
 * App assignment intent
 */
export enum AssignmentIntent {
  REQUIRED = 'required',
  AVAILABLE = 'available',
  UNINSTALL = 'uninstall',
  AVAILABLE_WITHOUT_ENROLLMENT = 'availableWithoutEnrollment'
}

/**
 * Managed application information
 */
export interface ManagedAppInfo {
  id: string;
  displayName: string;
  publisher: string;
  platform: AppPlatform;
  appType: string;
  version?: string;
  bundleId?: string;
  packageId?: string;
  productId?: string;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  description?: string;
  developer?: string;
  isFeatured: boolean;
  largeIcon?: string;
  privacyInformationUrl?: string;
  informationUrl?: string;
  owner?: string;
  notes?: string;
}

/**
 * App installation summary
 */
export interface AppInstallSummary {
  appId: string;
  appName: string;
  installedDeviceCount: number;
  failedDeviceCount: number;
  pendingInstallDeviceCount: number;
  notApplicableDeviceCount: number;
  notInstalledDeviceCount: number;
  installedUserCount: number;
  failedUserCount: number;
  pendingInstallUserCount: number;
  notApplicableUserCount: number;
  notInstalledUserCount: number;
}

/**
 * App installation status per device
 */
export interface AppDeviceStatus {
  id: string;
  appId: string;
  appName: string;
  deviceId: string;
  deviceName: string;
  userName?: string;
  userPrincipalName?: string;
  platform?: string;
  osVersion?: string;
  osDescription?: string;
  installState: InstallState;
  installStateDetail?: string;
  errorCode?: string;
  lastSyncDateTime?: Date;
  displayVersion?: string;
  installDate?: Date;
}

/**
 * App installation status per user
 */
export interface AppUserStatus {
  id: string;
  appId: string;
  appName: string;
  userName: string;
  userPrincipalName: string;
  installedDeviceCount: number;
  failedDeviceCount: number;
  notInstalledDeviceCount: number;
  userEmail?: string;
}

/**
 * Failed installation with error details
 */
export interface FailedInstallation {
  appId: string;
  appName: string;
  appVersion?: string;
  deviceId: string;
  deviceName: string;
  userName?: string;
  userPrincipalName?: string;
  platform?: string;
  osVersion?: string;
  errorCode?: string;
  errorDescription?: string;
  installState: string;
  installStateDetail?: string;
  lastSyncDateTime?: Date;
  failureReason?: string;
  troubleshootingLink?: string;
}

/**
 * App assignment information
 */
export interface AppAssignment {
  id: string;
  appId: string;
  appName: string;
  intent: AssignmentIntent;
  targetType: 'group' | 'allUsers' | 'allDevices';
  targetId?: string;
  targetName?: string;
  includeExclude: 'include' | 'exclude';
  filterType?: string;
  filterMode?: 'include' | 'exclude';
  settings?: any;
}

/**
 * App update compliance information
 */
export interface AppUpdateCompliance {
  appId: string;
  appName: string;
  currentVersion: string;
  latestVersion?: string;
  devicesOnCurrentVersion: number;
  devicesOnOlderVersion: number;
  devicesRequiringUpdate: number;
  updateCompliancePercentage: number;
  lastUpdateCheckDate?: Date;
  deviceDetails: AppUpdateDeviceDetail[];
}

/**
 * Device-level update compliance detail
 */
export interface AppUpdateDeviceDetail {
  deviceId: string;
  deviceName: string;
  currentVersion: string;
  requiresUpdate: boolean;
  userName?: string;
  lastSyncDateTime?: Date;
}

/**
 * Deployment statistics
 */
export interface DeploymentStatistics {
  totalApps: number;
  appsByPlatform: Record<AppPlatform, number>;
  totalDevices: number;
  totalUsers: number;
  totalInstallations: number;
  successfulInstallations: number;
  failedInstallations: number;
  pendingInstallations: number;
  successRate: number;
  failureRate: number;
  topFailedApps: Array<{ appName: string; failureCount: number }>;
  deploymentsByPlatform: Record<string, number>;
  totalAssignments: number;
  assignmentsByIntent: Record<AssignmentIntent, number>;
  assignmentsByTargetType: Record<'group' | 'allUsers' | 'allDevices', number>;
  appsWithAssignments: number;
  appsWithoutAssignments: number;
}

/**
 * Filter options for deployment queries
 */
export interface DeploymentFilterOptions {
  platform?: AppPlatform;
  appName?: string;
  publisher?: string;
  deviceId?: string;
  userName?: string;
  installState?: InstallState;
  includeSystemApps?: boolean;
  dateFrom?: Date;
  dateTo?: Date;
  maxResults?: number;
}

/**
 * Export options
 */
export interface ExportOptions {
  format: 'json' | 'csv' | 'html';
  outputDir: string;
  includeTimestamp?: boolean;
  includeMetadata?: boolean;
  includeSummary?: boolean;
}

// ===========================
// Main Report Class
// ===========================

/**
 * Application Deployment Report
 *
 * Provides comprehensive application deployment reporting with:
 * - All managed apps across platforms (iOS, Android, Windows, macOS, Web)
 * - Installation status per app and per device
 * - Deployment details per device
 * - App assignment details (groups, all users, all devices)
 * - Failed installations with error details and troubleshooting links
 * - Update compliance tracking and version management
 * - Multiple export formats (JSON, CSV, HTML)
 */
export class AppDeploymentReport extends BaseReport {
  name = 'app-deployment-report';
  description = 'Comprehensive application deployment reporting including installation status, assignments, failures, and update compliance';
  category = 'Applications';
  enabled = true;

  /**
   * Execute the complete app deployment report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Application Deployment Report');

    try {
      const startTime = Date.now();

      // Fetch all deployment data in parallel for better performance
      const [
        managedApps,
        installSummaries,
        deviceStatuses,
        failedInstallations,
        assignments
      ] = await Promise.all([
        this.getAllManagedApps(),
        this.getAllAppInstallSummaries(),
        this.getAllAppDeviceStatuses(),
        this.getFailedInstallations(),
        this.getAllAppAssignments()
      ]);

      // Get update compliance (requires processing app versions)
      const updateCompliance = await this.getAppUpdateCompliance(managedApps, deviceStatuses);

      // Generate deployment statistics
      const statistics = this.generateDeploymentStatistics({
        managedApps,
        installSummaries,
        deviceStatuses,
        failedInstallations,
        assignments
      });

      const duration = Date.now() - startTime;

      logger.info('Application Deployment Report completed', {
        duration: `${duration}ms`,
        managedApps: managedApps.length,
        installSummaries: installSummaries.length,
        deviceStatuses: deviceStatuses.length,
        failedInstallations: failedInstallations.length,
        assignments: assignments.length
      });

      // Combine all data for the report
      const data = [
        ...managedApps.map(app => ({ type: 'managed-app', ...app })),
        ...installSummaries.map(summary => ({ type: 'install-summary', ...summary })),
        ...deviceStatuses.map(status => ({ type: 'device-status', ...status })),
        ...failedInstallations.map(failure => ({ type: 'failed-installation', ...failure })),
        ...updateCompliance.map(compliance => ({ type: 'update-compliance', ...compliance })),
        ...assignments.map(assignment => ({ type: 'app-assignment', ...assignment }))
      ];

      return {
        metadata: this.createMetadata(this.name, data.length, {
          duration: `${duration}ms`,
          statistics
        }),
        data,
        summary: {
          statistics,
          reportGeneratedAt: new Date().toISOString()
        }
      };
    } catch (error) {
      logger.error('Failed to execute Application Deployment Report', error);
      throw error;
    }
  }

  // ===========================
  // 1. Get All Managed Apps
  // ===========================

  /**
   * Retrieves all managed applications from Intune
   * Supports iOS, Android, Windows, macOS, and Web apps
   */
  async getAllManagedApps(options?: DeploymentFilterOptions): Promise<ManagedAppInfo[]> {
    logger.info('Fetching all managed apps');
    const managedApps: ManagedAppInfo[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/mobileApps')
          .select([
            'id',
            'displayName',
            'description',
            'publisher',
            'createdDateTime',
            'lastModifiedDateTime',
            '@odata.type',
            'displayVersion',
            'bundleId',
            'packageId',
            'productId',
            'developer',
            'owner',
            'notes',
            'isFeatured',
            'largeIcon',
            'privacyInformationUrl',
            'informationUrl'
          ])
          .top(999)
          .get()
      );

      const apps = await this.getAllPages(response);

      for (const app of apps) {
        const managedApp = this.parseManagedApp(app);

        // Apply filters if provided
        if (this.matchesAppFilter(managedApp, options)) {
          managedApps.push(managedApp);
        }
      }

      logger.info(`Fetched ${managedApps.length} managed apps`);
      return managedApps;
    } catch (error: any) {
      logger.error('Error fetching managed apps', error);
      throw new Error(`Failed to fetch managed apps: ${error.message}`);
    }
  }

  /**
   * Get a specific managed app by ID
   */
  async getManagedAppById(appId: string): Promise<ManagedAppInfo> {
    logger.info(`Fetching managed app: ${appId}`);

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileApps/${appId}`)
          .get()
      );

      return this.parseManagedApp(response);
    } catch (error: any) {
      logger.error(`Error fetching app ${appId}`, error);
      throw new Error(`Failed to fetch app: ${error.message}`);
    }
  }

  // ===========================
  // 2. App Installation Status Per App
  // ===========================

  /**
   * Gets installation summary for all apps
   */
  async getAllAppInstallSummaries(options?: DeploymentFilterOptions): Promise<AppInstallSummary[]> {
    logger.info('Fetching app installation summaries');
    const summaries: AppInstallSummary[] = [];

    try {
      const apps = await this.getAllManagedApps(options);

      // Fetch install summaries in batches to avoid throttling
      const batchSize = 10;
      for (let i = 0; i < apps.length; i += batchSize) {
        const batch = apps.slice(i, i + batchSize);
        const batchSummaries = await Promise.allSettled(
          batch.map(app => this.getAppInstallSummary(app.id, app.displayName))
        );

        batchSummaries.forEach((result) => {
          if (result.status === 'fulfilled' && result.value) {
            summaries.push(result.value);
          }
        });

        // Small delay between batches to avoid rate limiting
        if (i + batchSize < apps.length) {
          await this.sleep(500);
        }
      }

      logger.info(`Fetched ${summaries.length} install summaries`);
      return summaries;
    } catch (error: any) {
      logger.error('Error fetching install summaries', error);
      return summaries; // Return partial results
    }
  }

  /**
   * Gets installation summary for a specific app
   */
  async getAppInstallSummary(appId: string, appName?: string): Promise<AppInstallSummary | null> {
    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileApps/${appId}/installSummary`)
          .get()
      );

      if (response) {
        return {
          appId,
          appName: appName || 'Unknown',
          installedDeviceCount: response.installedDeviceCount || 0,
          failedDeviceCount: response.failedDeviceCount || 0,
          pendingInstallDeviceCount: response.pendingInstallDeviceCount || 0,
          notApplicableDeviceCount: response.notApplicableDeviceCount || 0,
          notInstalledDeviceCount: response.notInstalledDeviceCount || 0,
          installedUserCount: response.installedUserCount || 0,
          failedUserCount: response.failedUserCount || 0,
          pendingInstallUserCount: response.pendingInstallUserCount || 0,
          notApplicableUserCount: response.notApplicableUserCount || 0,
          notInstalledUserCount: response.notInstalledUserCount || 0
        };
      }

      return null;
    } catch (error: any) {
      logger.warn(`Failed to fetch install summary for app ${appId}`, error);
      return null;
    }
  }

  // ===========================
  // 3. App Deployment Details Per Device
  // ===========================

  /**
   * Gets device-level installation status for all apps
   */
  async getAllAppDeviceStatuses(options?: DeploymentFilterOptions): Promise<AppDeviceStatus[]> {
    logger.info('Fetching app device statuses');
    const deviceStatuses: AppDeviceStatus[] = [];

    try {
      const apps = await this.getAllManagedApps(options);
      const maxApps = options?.maxResults || 50; // Limit to avoid timeout
      const appsToProcess = apps.slice(0, maxApps);

      logger.info(`Processing device statuses for ${appsToProcess.length} apps`);

      // Process apps in batches
      const batchSize = 5;
      for (let i = 0; i < appsToProcess.length; i += batchSize) {
        const batch = appsToProcess.slice(i, i + batchSize);
        const batchStatuses = await Promise.allSettled(
          batch.map(app => this.getAppDeviceStatuses(app.id, app.displayName))
        );

        batchStatuses.forEach((result) => {
          if (result.status === 'fulfilled') {
            deviceStatuses.push(...result.value);
          }
        });

        // Delay between batches
        if (i + batchSize < appsToProcess.length) {
          await this.sleep(1000);
        }
      }

      logger.info(`Fetched ${deviceStatuses.length} device status records`);
      return deviceStatuses;
    } catch (error: any) {
      logger.error('Error fetching device statuses', error);
      return deviceStatuses; // Return partial results
    }
  }

  /**
   * Gets device installation statuses for a specific app
   */
  async getAppDeviceStatuses(appId: string, appName?: string): Promise<AppDeviceStatus[]> {
    const statuses: AppDeviceStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileApps/${appId}/deviceStatuses`)
          .select([
            'id',
            'deviceId',
            'deviceName',
            'userName',
            'userPrincipalName',
            'platform',
            'osVersion',
            'osDescription',
            'installState',
            'installStateDetail',
            'errorCode',
            'lastSyncDateTime',
            'displayVersion'
          ])
          .top(999)
          .get()
      );

      const deviceStatuses = await this.getAllPages(response);

      for (const status of deviceStatuses) {
        statuses.push({
          id: status.id,
          appId,
          appName: appName || 'Unknown',
          deviceId: status.deviceId || 'Unknown',
          deviceName: status.deviceName || 'Unknown',
          userName: status.userName,
          userPrincipalName: status.userPrincipalName,
          platform: status.platform,
          osVersion: status.osVersion,
          osDescription: status.osDescription,
          installState: this.parseInstallState(status.installState),
          installStateDetail: status.installStateDetail,
          errorCode: status.errorCode,
          lastSyncDateTime: status.lastSyncDateTime ? new Date(status.lastSyncDateTime) : undefined,
          displayVersion: status.displayVersion
        });
      }

      return statuses;
    } catch (error: any) {
      logger.warn(`Failed to fetch device statuses for app ${appId}`, error);
      return statuses;
    }
  }

  /**
   * Gets user-level installation statuses for a specific app
   */
  async getAppUserStatuses(appId: string, appName?: string): Promise<AppUserStatus[]> {
    const statuses: AppUserStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileApps/${appId}/userStatuses`)
          .select([
            'id',
            'userName',
            'userPrincipalName',
            'installedDeviceCount',
            'failedDeviceCount',
            'notInstalledDeviceCount',
            'userEmail'
          ])
          .top(999)
          .get()
      );

      const userStatuses = await this.getAllPages(response);

      for (const status of userStatuses) {
        statuses.push({
          id: status.id,
          appId,
          appName: appName || 'Unknown',
          userName: status.userName || 'Unknown',
          userPrincipalName: status.userPrincipalName || 'Unknown',
          installedDeviceCount: status.installedDeviceCount || 0,
          failedDeviceCount: status.failedDeviceCount || 0,
          notInstalledDeviceCount: status.notInstalledDeviceCount || 0,
          userEmail: status.userEmail
        });
      }

      return statuses;
    } catch (error: any) {
      logger.warn(`Failed to fetch user statuses for app ${appId}`, error);
      return statuses;
    }
  }

  // ===========================
  // 4. App Assignment Details
  // ===========================

  /**
   * Gets all app assignments across all managed apps
   */
  async getAllAppAssignments(options?: DeploymentFilterOptions): Promise<AppAssignment[]> {
    logger.info('Fetching app assignments');
    const assignments: AppAssignment[] = [];

    try {
      const apps = await this.getAllManagedApps(options);
      const maxApps = options?.maxResults || 100; // Limit to avoid timeout
      const appsToProcess = apps.slice(0, maxApps);

      logger.info(`Processing assignments for ${appsToProcess.length} apps`);

      // Process apps in batches to avoid rate limiting
      const batchSize = 10;
      for (let i = 0; i < appsToProcess.length; i += batchSize) {
        const batch = appsToProcess.slice(i, i + batchSize);
        const batchAssignments = await Promise.allSettled(
          batch.map(app => this.getAppAssignments(app.id, app.displayName))
        );

        batchAssignments.forEach((result) => {
          if (result.status === 'fulfilled') {
            assignments.push(...result.value);
          }
        });

        // Small delay between batches
        if (i + batchSize < appsToProcess.length) {
          await this.sleep(500);
        }
      }

      logger.info(`Fetched ${assignments.length} app assignments`);
      return assignments;
    } catch (error: any) {
      logger.error('Error fetching app assignments', error);
      return assignments; // Return partial results
    }
  }

  /**
   * Gets assignments for a specific app
   */
  async getAppAssignments(appId: string, appName?: string): Promise<AppAssignment[]> {
    const assignments: AppAssignment[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileApps/${appId}/assignments`)
          .expand('target')
          .select([
            'id',
            'intent',
            'target',
            'settings'
          ])
          .get()
      );

      const assignmentData = response.value || [];

      for (const assignment of assignmentData) {
        assignments.push(this.parseAppAssignment(assignment, appId, appName));
      }

      return assignments;
    } catch (error: any) {
      logger.warn(`Failed to fetch assignments for app ${appId}`, error);
      return assignments;
    }
  }

  /**
   * Gets assignment statistics for a specific app
   */
  async getAppAssignmentStats(appId: string): Promise<{
    totalAssignments: number;
    requiredAssignments: number;
    availableAssignments: number;
    uninstallAssignments: number;
    groupAssignments: number;
    allUsersAssignments: number;
    allDevicesAssignments: number;
  }> {
    const assignments = await this.getAppAssignments(appId);

    return {
      totalAssignments: assignments.length,
      requiredAssignments: assignments.filter(a => a.intent === AssignmentIntent.REQUIRED).length,
      availableAssignments: assignments.filter(a => a.intent === AssignmentIntent.AVAILABLE).length,
      uninstallAssignments: assignments.filter(a => a.intent === AssignmentIntent.UNINSTALL).length,
      groupAssignments: assignments.filter(a => a.targetType === 'group').length,
      allUsersAssignments: assignments.filter(a => a.targetType === 'allUsers').length,
      allDevicesAssignments: assignments.filter(a => a.targetType === 'allDevices').length
    };
  }

  // ===========================
  // 5. Failed Installations with Error Details
  // ===========================

  /**
   * Gets all failed installations across all apps
   */
  async getFailedInstallations(options?: DeploymentFilterOptions): Promise<FailedInstallation[]> {
    logger.info('Fetching failed installations');
    const failedInstallations: FailedInstallation[] = [];

    try {
      const deviceStatuses = await this.getAllAppDeviceStatuses(options);

      // Filter for failed installations
      for (const status of deviceStatuses) {
        if (status.installState === InstallState.FAILED) {
          const failedInstall: FailedInstallation = {
            appId: status.appId,
            appName: status.appName,
            appVersion: status.displayVersion,
            deviceId: status.deviceId,
            deviceName: status.deviceName,
            userName: status.userName,
            userPrincipalName: status.userPrincipalName,
            platform: status.platform,
            osVersion: status.osVersion,
            errorCode: status.errorCode,
            errorDescription: this.getErrorDescription(status.errorCode),
            installState: status.installState,
            installStateDetail: status.installStateDetail,
            lastSyncDateTime: status.lastSyncDateTime,
            failureReason: this.determineFailureReason(status.errorCode, status.installStateDetail),
            troubleshootingLink: this.getTroubleshootingLink(status.errorCode)
          };

          failedInstallations.push(failedInstall);
        }
      }

      logger.info(`Found ${failedInstallations.length} failed installations`);
      return failedInstallations;
    } catch (error: any) {
      logger.error('Error fetching failed installations', error);
      return failedInstallations; // Return partial results
    }
  }

  /**
   * Gets failed installations for a specific app
   */
  async getFailedInstallationsByApp(appId: string, appName?: string): Promise<FailedInstallation[]> {
    logger.info(`Fetching failed installations for app: ${appId}`);

    try {
      const deviceStatuses = await this.getAppDeviceStatuses(appId, appName);
      const failedInstallations: FailedInstallation[] = [];

      for (const status of deviceStatuses) {
        if (status.installState === InstallState.FAILED) {
          failedInstallations.push({
            appId: status.appId,
            appName: status.appName,
            appVersion: status.displayVersion,
            deviceId: status.deviceId,
            deviceName: status.deviceName,
            userName: status.userName,
            userPrincipalName: status.userPrincipalName,
            platform: status.platform,
            osVersion: status.osVersion,
            errorCode: status.errorCode,
            errorDescription: this.getErrorDescription(status.errorCode),
            installState: status.installState,
            installStateDetail: status.installStateDetail,
            lastSyncDateTime: status.lastSyncDateTime,
            failureReason: this.determineFailureReason(status.errorCode, status.installStateDetail),
            troubleshootingLink: this.getTroubleshootingLink(status.errorCode)
          });
        }
      }

      return failedInstallations;
    } catch (error: any) {
      logger.error(`Error fetching failed installations for app ${appId}`, error);
      throw error;
    }
  }

  // ===========================
  // 6. App Update Compliance
  // ===========================

  /**
   * Gets update compliance information for all apps
   */
  async getAppUpdateCompliance(
    apps?: ManagedAppInfo[],
    deviceStatuses?: AppDeviceStatus[]
  ): Promise<AppUpdateCompliance[]> {
    logger.info('Analyzing app update compliance');
    const updateCompliance: AppUpdateCompliance[] = [];

    try {
      const managedApps = apps || await this.getAllManagedApps();
      const statuses = deviceStatuses || await this.getAllAppDeviceStatuses();

      // Group statuses by app
      const statusesByApp = new Map<string, AppDeviceStatus[]>();
      for (const status of statuses) {
        if (!statusesByApp.has(status.appId)) {
          statusesByApp.set(status.appId, []);
        }
        statusesByApp.get(status.appId)!.push(status);
      }

      // Analyze each app
      for (const app of managedApps) {
        const appStatuses = statusesByApp.get(app.id) || [];

        if (appStatuses.length === 0 || !app.version) {
          continue; // Skip apps with no deployment or version info
        }

        // Find all versions in use
        const versionCounts = new Map<string, number>();
        const deviceDetails: AppUpdateDeviceDetail[] = [];

        for (const status of appStatuses) {
          if (status.installState === InstallState.INSTALLED && status.displayVersion) {
            const version = status.displayVersion;
            versionCounts.set(version, (versionCounts.get(version) || 0) + 1);

            deviceDetails.push({
              deviceId: status.deviceId,
              deviceName: status.deviceName,
              currentVersion: version,
              requiresUpdate: version !== app.version,
              userName: status.userName,
              lastSyncDateTime: status.lastSyncDateTime
            });
          }
        }

        // Calculate compliance metrics
        const totalDevices = deviceDetails.length;
        const devicesOnCurrentVersion = deviceDetails.filter(d => d.currentVersion === app.version).length;
        const devicesOnOlderVersion = totalDevices - devicesOnCurrentVersion;
        const updateCompliancePercentage = totalDevices > 0
          ? Math.round((devicesOnCurrentVersion / totalDevices) * 100)
          : 100;

        updateCompliance.push({
          appId: app.id,
          appName: app.displayName,
          currentVersion: app.version,
          latestVersion: app.version,
          devicesOnCurrentVersion,
          devicesOnOlderVersion,
          devicesRequiringUpdate: devicesOnOlderVersion,
          updateCompliancePercentage,
          lastUpdateCheckDate: new Date(),
          deviceDetails: deviceDetails.filter(d => d.requiresUpdate) // Only include devices needing update
        });
      }

      logger.info(`Analyzed update compliance for ${updateCompliance.length} apps`);
      return updateCompliance;
    } catch (error: any) {
      logger.error('Error analyzing update compliance', error);
      return updateCompliance; // Return partial results
    }
  }

  /**
   * Gets update compliance for a specific app
   */
  async getAppUpdateComplianceById(appId: string): Promise<AppUpdateCompliance | null> {
    try {
      const app = await this.getManagedAppById(appId);
      const deviceStatuses = await this.getAppDeviceStatuses(appId, app.displayName);
      const compliance = await this.getAppUpdateCompliance([app], deviceStatuses);

      return compliance.length > 0 ? compliance[0] : null;
    } catch (error: any) {
      logger.error(`Error getting update compliance for app ${appId}`, error);
      return null;
    }
  }

  // ===========================
  // 7. Export Methods
  // ===========================

  /**
   * Export all deployment data to specified format
   */
  async exportDeploymentReport(options: ExportOptions): Promise<string> {
    logger.info(`Exporting deployment report to ${options.format}`);

    try {
      const reportData = await this.execute();

      const outputPath = await OutputFormatter.format(
        reportData,
        options.format,
        options.outputDir,
        options.includeTimestamp ?? true
      );

      logger.info(`Deployment report exported to: ${outputPath}`);
      return outputPath;
    } catch (error: any) {
      logger.error('Error exporting deployment report', error);
      throw new Error(`Failed to export report: ${error.message}`);
    }
  }

  /**
   * Export app installation summaries
   */
  async exportInstallSummaries(options: ExportOptions): Promise<string> {
    logger.info('Exporting install summaries');

    try {
      const summaries = await this.getAllAppInstallSummaries();

      const reportData: ReportData = {
        metadata: this.createMetadata('app-install-summaries', summaries.length),
        data: summaries,
        summary: {
          totalApps: summaries.length,
          totalSuccessfulInstalls: summaries.reduce((sum, s) => sum + s.installedDeviceCount, 0),
          totalFailedInstalls: summaries.reduce((sum, s) => sum + s.failedDeviceCount, 0)
        }
      };

      const outputPath = await OutputFormatter.format(
        reportData,
        options.format,
        options.outputDir,
        options.includeTimestamp ?? true
      );

      logger.info(`Install summaries exported to: ${outputPath}`);
      return outputPath;
    } catch (error: any) {
      logger.error('Error exporting install summaries', error);
      throw error;
    }
  }

  /**
   * Export failed installations
   */
  async exportFailedInstallations(options: ExportOptions): Promise<string> {
    logger.info('Exporting failed installations');

    try {
      const failures = await this.getFailedInstallations();

      const reportData: ReportData = {
        metadata: this.createMetadata('failed-installations', failures.length),
        data: failures,
        summary: {
          totalFailures: failures.length,
          uniqueApps: new Set(failures.map(f => f.appId)).size,
          uniqueDevices: new Set(failures.map(f => f.deviceId)).size
        }
      };

      const outputPath = await OutputFormatter.format(
        reportData,
        options.format,
        options.outputDir,
        options.includeTimestamp ?? true
      );

      logger.info(`Failed installations exported to: ${outputPath}`);
      return outputPath;
    } catch (error: any) {
      logger.error('Error exporting failed installations', error);
      throw error;
    }
  }

  /**
   * Export update compliance data
   */
  async exportUpdateCompliance(options: ExportOptions): Promise<string> {
    logger.info('Exporting update compliance');

    try {
      const compliance = await this.getAppUpdateCompliance();

      const reportData: ReportData = {
        metadata: this.createMetadata('app-update-compliance', compliance.length),
        data: compliance,
        summary: {
          totalApps: compliance.length,
          averageCompliance: compliance.length > 0
            ? Math.round(compliance.reduce((sum, c) => sum + c.updateCompliancePercentage, 0) / compliance.length)
            : 0,
          totalDevicesRequiringUpdate: compliance.reduce((sum, c) => sum + c.devicesRequiringUpdate, 0)
        }
      };

      const outputPath = await OutputFormatter.format(
        reportData,
        options.format,
        options.outputDir,
        options.includeTimestamp ?? true
      );

      logger.info(`Update compliance exported to: ${outputPath}`);
      return outputPath;
    } catch (error: any) {
      logger.error('Error exporting update compliance', error);
      throw error;
    }
  }

  /**
   * Export app assignment data
   */
  async exportAppAssignments(options: ExportOptions): Promise<string> {
    logger.info('Exporting app assignments');

    try {
      const assignments = await this.getAllAppAssignments();

      // Calculate assignment statistics
      const assignmentsByIntent = {
        required: assignments.filter(a => a.intent === AssignmentIntent.REQUIRED).length,
        available: assignments.filter(a => a.intent === AssignmentIntent.AVAILABLE).length,
        uninstall: assignments.filter(a => a.intent === AssignmentIntent.UNINSTALL).length,
        availableWithoutEnrollment: assignments.filter(
          a => a.intent === AssignmentIntent.AVAILABLE_WITHOUT_ENROLLMENT
        ).length
      };

      const assignmentsByTargetType = {
        group: assignments.filter(a => a.targetType === 'group').length,
        allUsers: assignments.filter(a => a.targetType === 'allUsers').length,
        allDevices: assignments.filter(a => a.targetType === 'allDevices').length
      };

      const reportData: ReportData = {
        metadata: this.createMetadata('app-assignments', assignments.length),
        data: assignments,
        summary: {
          totalAssignments: assignments.length,
          uniqueApps: new Set(assignments.map(a => a.appId)).size,
          assignmentsByIntent,
          assignmentsByTargetType,
          includeAssignments: assignments.filter(a => a.includeExclude === 'include').length,
          excludeAssignments: assignments.filter(a => a.includeExclude === 'exclude').length
        }
      };

      const outputPath = await OutputFormatter.format(
        reportData,
        options.format,
        options.outputDir,
        options.includeTimestamp ?? true
      );

      logger.info(`App assignments exported to: ${outputPath}`);
      return outputPath;
    } catch (error: any) {
      logger.error('Error exporting app assignments', error);
      throw error;
    }
  }

  // ===========================
  // Helper Methods
  // ===========================

  /**
   * Parse managed app from Graph API response
   */
  private parseManagedApp(app: any): ManagedAppInfo {
    const platform = this.determineAppPlatform(app['@odata.type']);

    return {
      id: app.id,
      displayName: app.displayName || 'Unknown',
      publisher: app.publisher || 'Unknown',
      platform,
      appType: app['@odata.type'] || 'unknown',
      version: app.displayVersion || app.versionNumber,
      bundleId: app.bundleId,
      packageId: app.packageId,
      productId: app.productId,
      createdDateTime: new Date(app.createdDateTime),
      lastModifiedDateTime: new Date(app.lastModifiedDateTime),
      description: app.description,
      developer: app.developer,
      isFeatured: app.isFeatured || false,
      largeIcon: app.largeIcon?.value,
      privacyInformationUrl: app.privacyInformationUrl,
      informationUrl: app.informationUrl,
      owner: app.owner,
      notes: app.notes
    };
  }

  /**
   * Determine app platform from OData type
   */
  private determineAppPlatform(odataType: string): AppPlatform {
    if (!odataType) return AppPlatform.UNKNOWN;

    const typeMap: Record<string, AppPlatform> = {
      'iosStoreApp': AppPlatform.IOS,
      'iosVppApp': AppPlatform.IOS,
      'iosLobApp': AppPlatform.IOS,
      'managedIOSStoreApp': AppPlatform.IOS,
      'managedIOSLobApp': AppPlatform.IOS,
      'androidManagedApp': AppPlatform.ANDROID,
      'androidStoreApp': AppPlatform.ANDROID,
      'androidForWorkApp': AppPlatform.ANDROID,
      'managedAndroidStoreApp': AppPlatform.ANDROID,
      'managedAndroidLobApp': AppPlatform.ANDROID,
      'windowsMobileMSI': AppPlatform.WINDOWS,
      'win32LobApp': AppPlatform.WINDOWS,
      'windowsUniversalAppX': AppPlatform.WINDOWS,
      'officeSuiteApp': AppPlatform.WINDOWS,
      'windowsMicrosoftEdgeApp': AppPlatform.WINDOWS,
      'macOSLobApp': AppPlatform.MACOS,
      'macOSOfficeSuiteApp': AppPlatform.MACOS,
      'macOSMicrosoftEdgeApp': AppPlatform.MACOS,
      'webApp': AppPlatform.WEB
    };

    for (const [key, value] of Object.entries(typeMap)) {
      if (odataType.toLowerCase().includes(key.toLowerCase())) {
        return value;
      }
    }

    return AppPlatform.UNKNOWN;
  }

  /**
   * Parse install state from API response
   */
  private parseInstallState(installState: string): InstallState {
    if (!installState) return InstallState.UNKNOWN;

    const stateLower = installState.toLowerCase();

    if (stateLower.includes('installed') && !stateLower.includes('not')) {
      return InstallState.INSTALLED;
    }
    if (stateLower.includes('failed')) {
      return InstallState.FAILED;
    }
    if (stateLower.includes('notinstalled') || stateLower.includes('not installed')) {
      return InstallState.NOT_INSTALLED;
    }
    if (stateLower.includes('available')) {
      return InstallState.AVAILABLE;
    }
    if (stateLower.includes('pending')) {
      return InstallState.PENDING;
    }
    if (stateLower.includes('uninstalling')) {
      return InstallState.UNINSTALLING;
    }
    if (stateLower.includes('downloading')) {
      return InstallState.DOWNLOADING;
    }

    return InstallState.UNKNOWN;
  }

  /**
   * Parse app assignment from Graph API response
   */
  private parseAppAssignment(assignment: any, appId: string, appName?: string): AppAssignment {
    const target = assignment.target || {};
    const odataType = target['@odata.type'] || '';

    // Determine target type
    let targetType: 'group' | 'allUsers' | 'allDevices' = 'group';
    let targetId: string | undefined;
    let targetName: string | undefined;

    if (odataType.includes('allLicensedUsersAssignmentTarget')) {
      targetType = 'allUsers';
      targetName = 'All Users';
    } else if (odataType.includes('allDevicesAssignmentTarget')) {
      targetType = 'allDevices';
      targetName = 'All Devices';
    } else if (odataType.includes('groupAssignmentTarget')) {
      targetType = 'group';
      targetId = target.groupId;
      // Note: Group name would require additional API call to /groups/{id}
      targetName = target.groupId ? `Group ${target.groupId}` : 'Unknown Group';
    }

    // Determine include/exclude
    const includeExclude: 'include' | 'exclude' = odataType.includes('exclusion') ? 'exclude' : 'include';

    // Parse intent
    const intent = this.parseAssignmentIntent(assignment.intent);

    return {
      id: assignment.id || `${appId}_${targetId || targetType}`,
      appId,
      appName: appName || 'Unknown',
      intent,
      targetType,
      targetId,
      targetName,
      includeExclude,
      filterType: target.deviceAndAppManagementAssignmentFilterType,
      filterMode: target.deviceAndAppManagementAssignmentFilterId ? 'include' : undefined,
      settings: assignment.settings
    };
  }

  /**
   * Parse assignment intent from API response
   */
  private parseAssignmentIntent(intent: string): AssignmentIntent {
    if (!intent) return AssignmentIntent.AVAILABLE;

    const intentLower = intent.toLowerCase();

    if (intentLower.includes('required')) {
      return AssignmentIntent.REQUIRED;
    }
    if (intentLower.includes('uninstall')) {
      return AssignmentIntent.UNINSTALL;
    }
    if (intentLower.includes('availablewithoutenrollment')) {
      return AssignmentIntent.AVAILABLE_WITHOUT_ENROLLMENT;
    }

    return AssignmentIntent.AVAILABLE;
  }

  /**
   * Get human-readable error description
   */
  private getErrorDescription(errorCode?: string): string {
    if (!errorCode) return 'Unknown error';

    const errorMap: Record<string, string> = {
      '0x80073CFF': 'App installation failed - Package not found',
      '0x87D1041C': 'App installation failed - Network connection required',
      '0x87D1041D': 'App installation failed - Insufficient storage space',
      '0x87D13B7D': 'App installation failed - Device not compliant',
      '0x87D13B9E': 'App installation failed - Device encryption not enabled',
      '0x87D13BA2': 'App installation failed - Minimum OS version not met',
      '0x87D13BA3': 'App installation failed - Maximum OS version exceeded',
      '0x87D13BA4': 'App installation failed - Device manufacturer restriction',
      '0x87D13BA5': 'App installation failed - Device model restriction',
      '0x8007065E': 'App installation failed - Authentication error',
      '0x80070057': 'App installation failed - Invalid parameter',
      '0x800706BA': 'App installation failed - RPC server unavailable',
      '0x80072EE7': 'App installation failed - Server name could not be resolved',
      '0x80072F8F': 'App installation failed - Certificate validation error'
    };

    return errorMap[errorCode] || `Error code: ${errorCode}`;
  }

  /**
   * Determine failure reason from error code and state detail
   */
  private determineFailureReason(errorCode?: string, stateDetail?: string): string {
    const reasons: string[] = [];

    if (errorCode) {
      const description = this.getErrorDescription(errorCode);
      if (description !== `Error code: ${errorCode}`) {
        reasons.push(description);
      }
    }

    if (stateDetail) {
      reasons.push(stateDetail);
    }

    return reasons.length > 0 ? reasons.join('; ') : 'Unknown failure reason';
  }

  /**
   * Get troubleshooting link for error code
   */
  private getTroubleshootingLink(errorCode?: string): string {
    if (!errorCode) return 'https://docs.microsoft.com/en-us/mem/intune/apps/troubleshoot-app-install';

    // Return specific troubleshooting documentation based on error patterns
    if (errorCode.startsWith('0x87D1')) {
      return 'https://docs.microsoft.com/en-us/mem/intune/apps/app-install-error-codes';
    }
    if (errorCode.startsWith('0x8007')) {
      return 'https://docs.microsoft.com/en-us/windows/win32/debug/system-error-codes';
    }

    return 'https://docs.microsoft.com/en-us/mem/intune/apps/troubleshoot-app-install';
  }

  /**
   * Check if app matches filter criteria
   */
  private matchesAppFilter(app: ManagedAppInfo, options?: DeploymentFilterOptions): boolean {
    if (!options) return true;

    if (options.platform && app.platform !== options.platform) {
      return false;
    }

    if (options.publisher && !app.publisher.toLowerCase().includes(options.publisher.toLowerCase())) {
      return false;
    }

    if (options.appName && !app.displayName.toLowerCase().includes(options.appName.toLowerCase())) {
      return false;
    }

    return true;
  }

  /**
   * Generate deployment statistics
   */
  private generateDeploymentStatistics(data: {
    managedApps: ManagedAppInfo[];
    installSummaries: AppInstallSummary[];
    deviceStatuses: AppDeviceStatus[];
    failedInstallations: FailedInstallation[];
    assignments: AppAssignment[];
  }): DeploymentStatistics {
    const { managedApps, installSummaries, deviceStatuses, failedInstallations, assignments } = data;

    // Count apps by platform
    const appsByPlatform: Record<AppPlatform, number> = {
      [AppPlatform.IOS]: 0,
      [AppPlatform.ANDROID]: 0,
      [AppPlatform.WINDOWS]: 0,
      [AppPlatform.MACOS]: 0,
      [AppPlatform.WEB]: 0,
      [AppPlatform.UNKNOWN]: 0
    };

    for (const app of managedApps) {
      appsByPlatform[app.platform]++;
    }

    // Calculate installation metrics
    const totalInstallations = installSummaries.reduce((sum, s) =>
      sum + s.installedDeviceCount + s.failedDeviceCount + s.pendingInstallDeviceCount, 0);
    const successfulInstallations = installSummaries.reduce((sum, s) => sum + s.installedDeviceCount, 0);
    const failedInstallations_count = installSummaries.reduce((sum, s) => sum + s.failedDeviceCount, 0);
    const pendingInstallations = installSummaries.reduce((sum, s) => sum + s.pendingInstallDeviceCount, 0);

    const successRate = totalInstallations > 0
      ? Math.round((successfulInstallations / totalInstallations) * 100)
      : 0;
    const failureRate = totalInstallations > 0
      ? Math.round((failedInstallations_count / totalInstallations) * 100)
      : 0;

    // Top failed apps
    const failuresByApp = new Map<string, number>();
    for (const failure of failedInstallations) {
      failuresByApp.set(failure.appName, (failuresByApp.get(failure.appName) || 0) + 1);
    }

    const topFailedApps = Array.from(failuresByApp.entries())
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([appName, failureCount]) => ({ appName, failureCount }));

    // Unique devices and users
    const uniqueDevices = new Set(deviceStatuses.map(s => s.deviceId)).size;
    const uniqueUsers = new Set(deviceStatuses.map(s => s.userPrincipalName).filter(u => u)).size;

    // Deployments by platform
    const deploymentsByPlatform: Record<string, number> = {};
    for (const status of deviceStatuses) {
      const platform = status.platform || 'Unknown';
      deploymentsByPlatform[platform] = (deploymentsByPlatform[platform] || 0) + 1;
    }

    // Assignment statistics
    const assignmentsByIntent: Record<AssignmentIntent, number> = {
      [AssignmentIntent.REQUIRED]: assignments.filter(a => a.intent === AssignmentIntent.REQUIRED).length,
      [AssignmentIntent.AVAILABLE]: assignments.filter(a => a.intent === AssignmentIntent.AVAILABLE).length,
      [AssignmentIntent.UNINSTALL]: assignments.filter(a => a.intent === AssignmentIntent.UNINSTALL).length,
      [AssignmentIntent.AVAILABLE_WITHOUT_ENROLLMENT]: assignments.filter(
        a => a.intent === AssignmentIntent.AVAILABLE_WITHOUT_ENROLLMENT
      ).length
    };

    const assignmentsByTargetType: Record<'group' | 'allUsers' | 'allDevices', number> = {
      group: assignments.filter(a => a.targetType === 'group').length,
      allUsers: assignments.filter(a => a.targetType === 'allUsers').length,
      allDevices: assignments.filter(a => a.targetType === 'allDevices').length
    };

    // Apps with/without assignments
    const appsWithAssignmentsSet = new Set(assignments.map(a => a.appId));
    const appsWithAssignments = appsWithAssignmentsSet.size;
    const appsWithoutAssignments = managedApps.length - appsWithAssignments;

    return {
      totalApps: managedApps.length,
      appsByPlatform,
      totalDevices: uniqueDevices,
      totalUsers: uniqueUsers,
      totalInstallations,
      successfulInstallations,
      failedInstallations: failedInstallations_count,
      pendingInstallations,
      successRate,
      failureRate,
      topFailedApps,
      deploymentsByPlatform,
      totalAssignments: assignments.length,
      assignmentsByIntent,
      assignmentsByTargetType,
      appsWithAssignments,
      appsWithoutAssignments
    };
  }
}

// Export for use in other modules
export default AppDeploymentReport;
