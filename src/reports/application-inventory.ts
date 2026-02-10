/**
 * Comprehensive Application Inventory Reporting Solution
 *
 * This module provides extensive application inventory reporting capabilities using
 * Microsoft Graph API, including:
 * 1. Managed app inventory
 * 2. Discovered apps on devices
 * 3. App installation status
 * 4. App protection policies status
 * 5. App configuration status
 *
 * @module application-inventory
 */

import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

// ===========================
// Type Definitions
// ===========================

/**
 * Managed application types in Intune
 */
export enum ManagedAppType {
  IOS = 'iOS',
  ANDROID = 'Android',
  WINDOWS = 'Windows',
  MACOS = 'macOS',
  WEB = 'Web',
  OTHER = 'Other'
}

/**
 * Application installation status
 */
export enum AppInstallationStatus {
  INSTALLED = 'installed',
  FAILED = 'failed',
  NOT_INSTALLED = 'notInstalled',
  PENDING = 'pending',
  AVAILABLE = 'available',
  UNINSTALLING = 'uninstalling',
  UNKNOWN = 'unknown'
}

/**
 * App protection policy status
 */
export enum AppProtectionStatus {
  COMPLIANT = 'compliant',
  NON_COMPLIANT = 'nonCompliant',
  NOT_APPLICABLE = 'notApplicable',
  UNKNOWN = 'unknown'
}

/**
 * Managed application information
 */
export interface ManagedApp {
  id: string;
  displayName: string;
  publisher: string;
  appType: ManagedAppType;
  version?: string;
  bundleId?: string;
  packageId?: string;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  description?: string;
  developer?: string;
  owner?: string;
  notes?: string;
  isFeatured: boolean;
  assignmentCount: number;
  installSummary?: AppInstallSummary;
}

/**
 * Discovered application on devices
 */
export interface DiscoveredApp {
  id: string;
  displayName: string;
  version: string;
  sizeInByte?: number;
  deviceCount: number;
  publisher?: string;
  platform?: string;
  detectedDateTime?: Date;
}

/**
 * Device-specific app information
 */
export interface DeviceAppInfo {
  deviceId: string;
  deviceName: string;
  userPrincipalName?: string;
  platform?: string;
  osVersion?: string;
  installedDateTime?: Date;
  lastSyncDateTime?: Date;
}

/**
 * App installation status details
 */
export interface AppInstallStatus {
  appId: string;
  appName: string;
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  status: AppInstallationStatus;
  installState?: string;
  installStateDetail?: string;
  errorCode?: string;
  lastSyncDateTime?: Date;
  osVersion?: string;
}

/**
 * App installation summary
 */
export interface AppInstallSummary {
  installedDeviceCount: number;
  failedDeviceCount: number;
  pendingInstallDeviceCount: number;
  notApplicableDeviceCount: number;
  notInstalledDeviceCount: number;
  installedUserCount: number;
  failedUserCount: number;
}

/**
 * App protection policy information
 */
export interface AppProtectionPolicy {
  id: string;
  displayName: string;
  description?: string;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  platform: 'android' | 'iOS' | 'windows';
  policyType: string;
  appCount: number;
  assignmentCount: number;
}

/**
 * App protection policy status per user
 */
export interface AppProtectionPolicyStatus {
  policyId: string;
  policyName: string;
  userPrincipalName: string;
  userId: string;
  platform: string;
  status: AppProtectionStatus;
  lastCheckInDateTime?: Date;
}

/**
 * App configuration policy
 */
export interface AppConfigurationPolicy {
  id: string;
  displayName: string;
  description?: string;
  createdDateTime: Date;
  lastModifiedDateTime: Date;
  targetedMobileApps: number;
  assignmentCount: number;
  hasPayload: boolean;
}

/**
 * App configuration status per device
 */
export interface AppConfigurationStatus {
  policyId: string;
  policyName: string;
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  status: 'success' | 'error' | 'pending' | 'notApplicable';
  lastReportedDateTime?: Date;
  errorCode?: string;
}

/**
 * Filter options for application inventory queries
 */
export interface AppInventoryFilterOptions {
  platform?: ManagedAppType;
  publisher?: string;
  appName?: string;
  deviceId?: string;
  userPrincipalName?: string;
  installationStatus?: AppInstallationStatus;
  includeSystemApps?: boolean;
  minDeviceCount?: number;
}

// ===========================
// Main Report Class
// ===========================

/**
 * Application Inventory Report
 *
 * Provides comprehensive application inventory reporting with:
 * - Managed apps inventory
 * - Discovered apps on devices
 * - App installation status
 * - App protection policies
 * - App configuration policies
 */
export class ApplicationInventoryReport extends BaseReport {
  name = 'application-inventory';
  description = 'Comprehensive inventory of all applications including managed apps, discovered apps, installation status, protection policies, and configuration policies';
  category = 'Applications';
  enabled = true;

  /**
   * Execute the complete application inventory report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Comprehensive Application Inventory Report');

    try {
      const startTime = Date.now();

      // Fetch all data in parallel for better performance
      const [
        managedApps,
        discoveredApps,
        installationStatus,
        protectionPolicies,
        protectionStatus,
        configurationPolicies,
        configurationStatus
      ] = await Promise.all([
        this.getManagedApps(),
        this.getDiscoveredApps(),
        this.getAppInstallationStatus(),
        this.getAppProtectionPolicies(),
        this.getAppProtectionPolicyStatus(),
        this.getAppConfigurationPolicies(),
        this.getAppConfigurationStatus()
      ]);

      // Generate summary statistics
      const summary = this.generateSummary({
        managedApps,
        discoveredApps,
        installationStatus,
        protectionPolicies,
        protectionStatus,
        configurationPolicies,
        configurationStatus
      });

      const duration = Date.now() - startTime;

      logger.info('Application Inventory Report completed', {
        duration: `${duration}ms`,
        managedApps: managedApps.length,
        discoveredApps: discoveredApps.length,
        installationRecords: installationStatus.length
      });

      // Combine all data for the report
      const data = [
        ...managedApps.map(app => ({ type: 'managed-app', ...app })),
        ...discoveredApps.map(app => ({ type: 'discovered-app', ...app })),
        ...installationStatus.map(status => ({ type: 'installation-status', ...status })),
        ...protectionPolicies.map(policy => ({ type: 'protection-policy', ...policy })),
        ...protectionStatus.map(status => ({ type: 'protection-status', ...status })),
        ...configurationPolicies.map(policy => ({ type: 'configuration-policy', ...policy })),
        ...configurationStatus.map(status => ({ type: 'configuration-status', ...status }))
      ];

      return {
        metadata: this.createMetadata(this.name, data.length, { duration: `${duration}ms` }),
        data,
        summary
      };
    } catch (error) {
      logger.error('Failed to execute Application Inventory Report', error);
      throw error;
    }
  }

  // ===========================
  // 1. Managed App Inventory
  // ===========================

  /**
   * Retrieves all managed applications from Intune
   */
  async getManagedApps(options?: AppInventoryFilterOptions): Promise<ManagedApp[]> {
    logger.info('Fetching managed apps');
    const managedApps: ManagedApp[] = [];

    try {
      // Fetch all mobile apps with expanded data
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/mobileApps')
          .expand('assignments')
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
            'developer',
            'owner',
            'notes',
            'isFeatured'
          ])
          .top(999)
          .get()
      );

      const apps = await this.getAllPages(response);

      // Process each app
      for (const app of apps) {
        const managedApp = this.parseManagedApp(app);

        // Apply filters if provided
        if (this.matchesFilter(managedApp, options)) {
          // Fetch install summary for each app
          try {
            const installSummary = await this.getAppInstallSummary(app.id);
            managedApp.installSummary = installSummary;
          } catch (error) {
            logger.warn(`Failed to fetch install summary for app ${app.id}`, error);
          }

          managedApps.push(managedApp);
        }
      }

      logger.info(`Fetched ${managedApps.length} managed apps`);
      return managedApps;
    } catch (error) {
      logger.error('Error fetching managed apps', error);
      throw new Error(`Failed to fetch managed apps: ${error.message}`);
    }
  }

  /**
   * Gets installation summary for a specific app
   */
  private async getAppInstallSummary(appId: string): Promise<AppInstallSummary | undefined> {
    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileApps/${appId}/installSummary`)
          .get()
      );

      if (response) {
        return {
          installedDeviceCount: response.installedDeviceCount || 0,
          failedDeviceCount: response.failedDeviceCount || 0,
          pendingInstallDeviceCount: response.pendingInstallDeviceCount || 0,
          notApplicableDeviceCount: response.notApplicableDeviceCount || 0,
          notInstalledDeviceCount: response.notInstalledDeviceCount || 0,
          installedUserCount: response.installedUserCount || 0,
          failedUserCount: response.failedUserCount || 0
        };
      }
    } catch (error) {
      // Install summary might not be available for all apps
      return undefined;
    }
  }

  // ===========================
  // 2. Discovered Apps on Devices
  // ===========================

  /**
   * Retrieves all discovered applications across all managed devices
   */
  async getDiscoveredApps(options?: AppInventoryFilterOptions): Promise<DiscoveredApp[]> {
    logger.info('Fetching discovered apps');
    const discoveredAppsMap = new Map<string, DiscoveredApp>();

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/detectedApps')
          .select([
            'id',
            'displayName',
            'version',
            'sizeInByte',
            'deviceCount',
            'platform'
          ])
          .top(999)
          .get()
      );

      const detectedApps = await this.getAllPages(response);

      for (const app of detectedApps) {
        const discoveredApp = this.parseDiscoveredApp(app);

        // Apply filters
        if (this.matchesDiscoveredAppFilter(discoveredApp, options)) {
          const key = `${discoveredApp.displayName}_${discoveredApp.version}`;

          if (discoveredAppsMap.has(key)) {
            const existing = discoveredAppsMap.get(key)!;
            existing.deviceCount += discoveredApp.deviceCount;
          } else {
            discoveredAppsMap.set(key, discoveredApp);
          }
        }
      }

      const discoveredApps = Array.from(discoveredAppsMap.values());
      logger.info(`Fetched ${discoveredApps.length} discovered apps`);
      return discoveredApps;
    } catch (error) {
      logger.error('Error fetching discovered apps', error);
      throw new Error(`Failed to fetch discovered apps: ${error.message}`);
    }
  }

  // ===========================
  // 3. App Installation Status
  // ===========================

  /**
   * Retrieves installation status for managed apps across all devices
   */
  async getAppInstallationStatus(options?: AppInventoryFilterOptions): Promise<AppInstallStatus[]> {
    logger.info('Fetching app installation status');
    const installStatuses: AppInstallStatus[] = [];

    try {
      // Get all managed devices
      const devicesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'userPrincipalName',
            'operatingSystem',
            'osVersion'
          ])
          .top(999)
          .get()
      );

      const devices = await this.getAllPages(devicesResponse);
      logger.info(`Processing installation status for ${devices.length} devices`);

      // Limit to first 100 devices to avoid timeout for very large tenants
      const deviceSample = devices.slice(0, 100);

      // For each device, get app installation status
      for (const device of deviceSample) {
        try {
          const deviceStatuses = await this.getDeviceAppInstallStatus(device);
          installStatuses.push(...deviceStatuses);
        } catch (error) {
          logger.warn(`Failed to fetch app status for device ${device.id}`, error);
        }
      }

      logger.info(`Fetched ${installStatuses.length} installation status records`);
      return installStatuses;
    } catch (error) {
      logger.error('Error fetching app installation status', error);
      // Return partial results rather than failing completely
      return installStatuses;
    }
  }

  /**
   * Gets app installation status for a specific device
   */
  private async getDeviceAppInstallStatus(device: any): Promise<AppInstallStatus[]> {
    const statuses: AppInstallStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${device.id}/detectedApps`)
          .select([
            'id',
            'displayName',
            'version'
          ])
          .top(50)
          .get()
      );

      if (response.value) {
        for (const app of response.value) {
          statuses.push({
            appId: app.id,
            appName: app.displayName || 'Unknown',
            deviceId: device.id,
            deviceName: device.deviceName || 'Unknown',
            userPrincipalName: device.userPrincipalName || 'Unknown',
            status: AppInstallationStatus.INSTALLED,
            osVersion: device.osVersion
          });
        }
      }
    } catch (error) {
      // Skip devices that can't be queried
    }

    return statuses;
  }

  // ===========================
  // 4. App Protection Policies Status
  // ===========================

  /**
   * Retrieves all app protection policies
   */
  async getAppProtectionPolicies(): Promise<AppProtectionPolicy[]> {
    logger.info('Fetching app protection policies');
    const policies: AppProtectionPolicy[] = [];

    try {
      // Get iOS app protection policies
      const iosPolicies = await this.getIOSAppProtectionPolicies();
      policies.push(...iosPolicies);

      // Get Android app protection policies
      const androidPolicies = await this.getAndroidAppProtectionPolicies();
      policies.push(...androidPolicies);

      // Get Windows Information Protection policies
      const windowsPolicies = await this.getWindowsInformationProtectionPolicies();
      policies.push(...windowsPolicies);

      logger.info(`Fetched ${policies.length} app protection policies`);
      return policies;
    } catch (error) {
      logger.error('Error fetching app protection policies', error);
      return policies;
    }
  }

  /**
   * Gets iOS app protection policies
   */
  private async getIOSAppProtectionPolicies(): Promise<AppProtectionPolicy[]> {
    const policies: AppProtectionPolicy[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/iosManagedAppProtections')
          .expand('apps,assignments')
          .select([
            'id',
            'displayName',
            'description',
            'createdDateTime',
            'lastModifiedDateTime'
          ])
          .get()
      );

      if (response.value) {
        for (const policy of response.value) {
          policies.push(this.parseAppProtectionPolicy(policy, 'iOS'));
        }
      }
    } catch (error) {
      logger.warn('Error fetching iOS app protection policies', error);
    }

    return policies;
  }

  /**
   * Gets Android app protection policies
   */
  private async getAndroidAppProtectionPolicies(): Promise<AppProtectionPolicy[]> {
    const policies: AppProtectionPolicy[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/androidManagedAppProtections')
          .expand('apps,assignments')
          .select([
            'id',
            'displayName',
            'description',
            'createdDateTime',
            'lastModifiedDateTime'
          ])
          .get()
      );

      if (response.value) {
        for (const policy of response.value) {
          policies.push(this.parseAppProtectionPolicy(policy, 'android'));
        }
      }
    } catch (error) {
      logger.warn('Error fetching Android app protection policies', error);
    }

    return policies;
  }

  /**
   * Gets Windows Information Protection policies
   */
  private async getWindowsInformationProtectionPolicies(): Promise<AppProtectionPolicy[]> {
    const policies: AppProtectionPolicy[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/windowsInformationProtectionPolicies')
          .expand('assignments')
          .select([
            'id',
            'displayName',
            'description',
            'createdDateTime',
            'lastModifiedDateTime'
          ])
          .get()
      );

      if (response.value) {
        for (const policy of response.value) {
          policies.push(this.parseAppProtectionPolicy(policy, 'windows'));
        }
      }
    } catch (error) {
      logger.warn('Error fetching Windows Information Protection policies', error);
    }

    return policies;
  }

  /**
   * Gets app protection policy status for all users
   */
  async getAppProtectionPolicyStatus(): Promise<AppProtectionPolicyStatus[]> {
    logger.info('Fetching app protection policy status');
    const statuses: AppProtectionPolicyStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/managedAppStatuses')
          .get()
      );

      if (response.value) {
        for (const status of response.value) {
          const parsedStatus = this.parseAppProtectionStatus(status);
          if (parsedStatus) {
            statuses.push(parsedStatus);
          }
        }
      }

      logger.info(`Fetched ${statuses.length} app protection status records`);
      return statuses;
    } catch (error) {
      logger.warn('Error fetching app protection policy status', error);
      return statuses;
    }
  }

  // ===========================
  // 5. App Configuration Status
  // ===========================

  /**
   * Retrieves all app configuration policies
   */
  async getAppConfigurationPolicies(): Promise<AppConfigurationPolicy[]> {
    logger.info('Fetching app configuration policies');
    const policies: AppConfigurationPolicy[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceAppManagement/mobileAppConfigurations')
          .expand('assignments')
          .select([
            'id',
            'displayName',
            'description',
            'createdDateTime',
            'lastModifiedDateTime',
            'targetedMobileApps'
          ])
          .get()
      );

      if (response.value) {
        for (const config of response.value) {
          policies.push(this.parseAppConfigurationPolicy(config));
        }
      }

      logger.info(`Fetched ${policies.length} app configuration policies`);
      return policies;
    } catch (error) {
      logger.error('Error fetching app configuration policies', error);
      return policies;
    }
  }

  /**
   * Gets app configuration status for all devices
   */
  async getAppConfigurationStatus(): Promise<AppConfigurationStatus[]> {
    logger.info('Fetching app configuration status');
    const statuses: AppConfigurationStatus[] = [];

    try {
      const policies = await this.getAppConfigurationPolicies();

      // For each policy, get device statuses
      for (const policy of policies) {
        try {
          const policyStatuses = await this.getAppConfigurationStatusByPolicy(policy.id);
          statuses.push(...policyStatuses);
        } catch (error) {
          logger.warn(`Failed to fetch status for configuration policy ${policy.id}`, error);
        }
      }

      logger.info(`Fetched ${statuses.length} app configuration status records`);
      return statuses;
    } catch (error) {
      logger.error('Error fetching app configuration status', error);
      return statuses;
    }
  }

  /**
   * Gets app configuration status for a specific policy
   */
  private async getAppConfigurationStatusByPolicy(policyId: string): Promise<AppConfigurationStatus[]> {
    const statuses: AppConfigurationStatus[] = [];

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceAppManagement/mobileAppConfigurations/${policyId}/deviceStatuses`)
          .select([
            'id',
            'deviceDisplayName',
            'userName',
            'status',
            'lastReportedDateTime',
            'platform'
          ])
          .top(100)
          .get()
      );

      if (response.value) {
        for (const status of response.value) {
          statuses.push(this.parseAppConfigurationStatus(status, policyId));
        }
      }
    } catch (error) {
      // Some policies might not have status available
    }

    return statuses;
  }

  // ===========================
  // Helper Methods
  // ===========================

  /**
   * Parses a managed app from Graph API response
   */
  private parseManagedApp(app: any): ManagedApp {
    const appType = this.determineAppType(app['@odata.type']);

    return {
      id: app.id,
      displayName: app.displayName || 'Unknown',
      publisher: app.publisher || 'Unknown',
      appType,
      version: app.displayVersion || app.bundleId || app.packageId,
      bundleId: app.bundleId,
      packageId: app.packageId,
      createdDateTime: new Date(app.createdDateTime),
      lastModifiedDateTime: new Date(app.lastModifiedDateTime),
      description: app.description,
      developer: app.developer,
      owner: app.owner,
      notes: app.notes,
      isFeatured: app.isFeatured || false,
      assignmentCount: app.assignments?.length || 0
    };
  }

  /**
   * Parses a discovered app from Graph API response
   */
  private parseDiscoveredApp(app: any): DiscoveredApp {
    return {
      id: app.id,
      displayName: app.displayName || 'Unknown',
      version: app.version || 'Unknown',
      sizeInByte: app.sizeInByte,
      deviceCount: app.deviceCount || 0,
      platform: app.platform,
      detectedDateTime: app.detectedDateTime ? new Date(app.detectedDateTime) : undefined
    };
  }

  /**
   * Parses app protection policy from Graph API response
   */
  private parseAppProtectionPolicy(policy: any, platform: 'iOS' | 'android' | 'windows'): AppProtectionPolicy {
    return {
      id: policy.id,
      displayName: policy.displayName || 'Unknown',
      description: policy.description,
      createdDateTime: new Date(policy.createdDateTime),
      lastModifiedDateTime: new Date(policy.lastModifiedDateTime),
      platform,
      policyType: policy['@odata.type'] || 'unknown',
      appCount: policy.apps?.length || 0,
      assignmentCount: policy.assignments?.length || 0
    };
  }

  /**
   * Parses app protection status from Graph API response
   */
  private parseAppProtectionStatus(status: any): AppProtectionPolicyStatus | null {
    if (!status.id) return null;

    return {
      policyId: status.id,
      policyName: status.displayName || 'Unknown',
      userPrincipalName: status.userPrincipalName || 'Unknown',
      userId: status.userId || 'Unknown',
      platform: status.platform || 'Unknown',
      status: this.mapProtectionStatus(status.status),
      lastCheckInDateTime: status.lastCheckInDateTime ? new Date(status.lastCheckInDateTime) : undefined
    };
  }

  /**
   * Parses app configuration policy from Graph API response
   */
  private parseAppConfigurationPolicy(config: any): AppConfigurationPolicy {
    return {
      id: config.id,
      displayName: config.displayName || 'Unknown',
      description: config.description,
      createdDateTime: new Date(config.createdDateTime),
      lastModifiedDateTime: new Date(config.lastModifiedDateTime),
      targetedMobileApps: config.targetedMobileApps?.length || 0,
      assignmentCount: config.assignments?.length || 0,
      hasPayload: !!config.payloadJson || !!config.encodedSettingXml
    };
  }

  /**
   * Parses app configuration status from Graph API response
   */
  private parseAppConfigurationStatus(status: any, policyId: string): AppConfigurationStatus {
    return {
      policyId,
      policyName: status.displayName || 'Unknown',
      deviceId: status.id,
      deviceName: status.deviceDisplayName || 'Unknown',
      userPrincipalName: status.userName || 'Unknown',
      platform: status.platform || 'Unknown',
      status: this.mapConfigurationStatus(status.status),
      lastReportedDateTime: status.lastReportedDateTime ? new Date(status.lastReportedDateTime) : undefined,
      errorCode: status.errorCode
    };
  }

  /**
   * Determines app type from OData type
   */
  private determineAppType(odataType: string): ManagedAppType {
    if (!odataType) return ManagedAppType.OTHER;

    const typeMap: Record<string, ManagedAppType> = {
      'iosStoreApp': ManagedAppType.IOS,
      'iosVppApp': ManagedAppType.IOS,
      'iosLobApp': ManagedAppType.IOS,
      'androidManagedApp': ManagedAppType.ANDROID,
      'androidStoreApp': ManagedAppType.ANDROID,
      'androidForWorkApp': ManagedAppType.ANDROID,
      'windowsMobileMSI': ManagedAppType.WINDOWS,
      'win32LobApp': ManagedAppType.WINDOWS,
      'officeSuiteApp': ManagedAppType.WINDOWS,
      'macOSLobApp': ManagedAppType.MACOS,
      'macOSOfficeSuiteApp': ManagedAppType.MACOS,
      'webApp': ManagedAppType.WEB
    };

    for (const [key, value] of Object.entries(typeMap)) {
      if (odataType.toLowerCase().includes(key.toLowerCase())) {
        return value;
      }
    }

    return ManagedAppType.OTHER;
  }

  /**
   * Maps protection status to AppProtectionStatus enum
   */
  private mapProtectionStatus(status: string): AppProtectionStatus {
    if (!status) return AppProtectionStatus.UNKNOWN;

    const statusLower = status.toLowerCase();

    if (statusLower.includes('compliant') && !statusLower.includes('non')) {
      return AppProtectionStatus.COMPLIANT;
    }
    if (statusLower.includes('noncompliant') || statusLower.includes('non-compliant')) {
      return AppProtectionStatus.NON_COMPLIANT;
    }
    if (statusLower.includes('notapplicable') || statusLower.includes('not applicable')) {
      return AppProtectionStatus.NOT_APPLICABLE;
    }

    return AppProtectionStatus.UNKNOWN;
  }

  /**
   * Maps configuration status to standard status string
   */
  private mapConfigurationStatus(status: string): 'success' | 'error' | 'pending' | 'notApplicable' {
    if (!status) return 'notApplicable';

    const statusLower = status.toLowerCase();

    if (statusLower.includes('success') || statusLower.includes('compliant')) {
      return 'success';
    }
    if (statusLower.includes('error') || statusLower.includes('failed')) {
      return 'error';
    }
    if (statusLower.includes('pending')) {
      return 'pending';
    }

    return 'notApplicable';
  }

  /**
   * Checks if a managed app matches the filter criteria
   */
  private matchesFilter(app: ManagedApp, options?: AppInventoryFilterOptions): boolean {
    if (!options) return true;

    if (options.platform && app.appType !== options.platform) {
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
   * Checks if a discovered app matches the filter criteria
   */
  private matchesDiscoveredAppFilter(app: DiscoveredApp, options?: AppInventoryFilterOptions): boolean {
    if (!options) return true;

    if (options.appName && !app.displayName.toLowerCase().includes(options.appName.toLowerCase())) {
      return false;
    }

    if (options.minDeviceCount && app.deviceCount < options.minDeviceCount) {
      return false;
    }

    return true;
  }

  /**
   * Generates summary statistics for the inventory report
   */
  private generateSummary(data: {
    managedApps: ManagedApp[];
    discoveredApps: DiscoveredApp[];
    installationStatus: AppInstallStatus[];
    protectionPolicies: AppProtectionPolicy[];
    protectionStatus: AppProtectionPolicyStatus[];
    configurationPolicies: AppConfigurationPolicy[];
    configurationStatus: AppConfigurationStatus[];
  }): Record<string, any> {
    // Count apps by platform
    const appsByPlatform = {
      iOS: data.managedApps.filter(a => a.appType === ManagedAppType.IOS).length,
      Android: data.managedApps.filter(a => a.appType === ManagedAppType.ANDROID).length,
      Windows: data.managedApps.filter(a => a.appType === ManagedAppType.WINDOWS).length,
      macOS: data.managedApps.filter(a => a.appType === ManagedAppType.MACOS).length,
      Web: data.managedApps.filter(a => a.appType === ManagedAppType.WEB).length,
      Other: data.managedApps.filter(a => a.appType === ManagedAppType.OTHER).length
    };

    // Count installation statuses
    const installationSummary = {
      installed: data.installationStatus.filter(s => s.status === AppInstallationStatus.INSTALLED).length,
      failed: data.installationStatus.filter(s => s.status === AppInstallationStatus.FAILED).length,
      pending: data.installationStatus.filter(s => s.status === AppInstallationStatus.PENDING).length,
      notInstalled: data.installationStatus.filter(s => s.status === AppInstallationStatus.NOT_INSTALLED).length
    };

    // Count protection statuses
    const protectionSummary = {
      compliant: data.protectionStatus.filter(s => s.status === AppProtectionStatus.COMPLIANT).length,
      nonCompliant: data.protectionStatus.filter(s => s.status === AppProtectionStatus.NON_COMPLIANT).length,
      notApplicable: data.protectionStatus.filter(s => s.status === AppProtectionStatus.NOT_APPLICABLE).length
    };

    // Count configuration statuses
    const configurationSummary = {
      success: data.configurationStatus.filter(s => s.status === 'success').length,
      error: data.configurationStatus.filter(s => s.status === 'error').length,
      pending: data.configurationStatus.filter(s => s.status === 'pending').length,
      notApplicable: data.configurationStatus.filter(s => s.status === 'notApplicable').length
    };

    // Count unique devices
    const uniqueDevices = new Set<string>();
    data.installationStatus.forEach(status => uniqueDevices.add(status.deviceId));

    // Top apps by device count
    const topDiscoveredApps = data.discoveredApps
      .sort((a, b) => b.deviceCount - a.deviceCount)
      .slice(0, 10)
      .map(app => ({
        name: app.displayName,
        version: app.version,
        deviceCount: app.deviceCount
      }));

    // Apps with most failures
    const appsWithFailures = data.installationStatus
      .filter(s => s.status === AppInstallationStatus.FAILED)
      .reduce((acc, status) => {
        acc[status.appName] = (acc[status.appName] || 0) + 1;
        return acc;
      }, {} as Record<string, number>);

    const topFailedApps = Object.entries(appsWithFailures)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([name, count]) => ({ name, failureCount: count }));

    return {
      overview: {
        totalManagedApps: data.managedApps.length,
        totalDiscoveredApps: data.discoveredApps.length,
        totalDevices: uniqueDevices.size,
        totalProtectionPolicies: data.protectionPolicies.length,
        totalConfigurationPolicies: data.configurationPolicies.length
      },
      appsByPlatform,
      installationSummary,
      protectionSummary,
      configurationSummary,
      protectionPoliciesByPlatform: {
        iOS: data.protectionPolicies.filter(p => p.platform === 'iOS').length,
        Android: data.protectionPolicies.filter(p => p.platform === 'android').length,
        Windows: data.protectionPolicies.filter(p => p.platform === 'windows').length
      },
      insights: {
        topDiscoveredApps,
        topFailedApps,
        totalAssignments: data.managedApps.reduce((sum, app) => sum + app.assignmentCount, 0),
        appsWithProtectionPolicies: data.protectionPolicies.reduce((sum, p) => sum + p.appCount, 0)
      }
    };
  }
}

// Export for use in other modules
export default ApplicationInventoryReport;
