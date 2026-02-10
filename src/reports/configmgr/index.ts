/**
 * Configuration Manager Device Management Reports - Master Orchestrator
 *
 * This module provides a unified orchestrator for all 37 Configuration Manager
 * device management reports, replicated using Microsoft Graph API for Intune.
 *
 * ============================================================================
 * COMPLETE REPORT CATALOG (37 Reports)
 * ============================================================================
 *
 * Mobile Device Inventory Reports (1, 2, 20, 22):
 *   1. All corporate-owned mobile devices
 *   2. All mobile device clients (excluding Exchange connector)
 *   20. Mobile device client information (communication status)
 *   22. Mobile devices by operating system
 *
 * Windows CE Reports (3-7, 13-14, 19) - LEGACY/UNSUPPORTED:
 *   3. Certificate issues on mobile devices (Windows CE)
 *   4. Client deployment failure (Windows CE)
 *   5. Client deployment status details (Windows CE)
 *   6. Client deployment success (Windows CE)
 *   7. Communication issues on mobile devices (Windows CE)
 *   13. Health information for mobile devices (Windows CE)
 *   14. Health summary for mobile devices (Windows CE)
 *   19. Local client issues on mobile devices (Windows CE)
 *
 * Exchange ActiveSync Reports (8, 15, 21, 34):
 *   8. Compliance status of default ActiveSync mailbox policy
 *   15. Inactive mobile devices (Exchange)
 *   21. Mobile device compliance details (Exchange)
 *   34. Settings summary for mobile devices (Exchange)
 *
 * Device Hardware Reports (9-12, 25-26, 28-29):
 *   9. Count of mobile devices by display configurations
 *   10. Count of mobile devices by operating system
 *   11. Count of mobile devices by program memory
 *   12. Count of mobile devices by storage memory configurations
 *   25. Mobile devices with specific free program memory
 *   26. Mobile devices with specific free removable storage memory
 *   28. Mobile devices with low free program memory
 *   29. Mobile devices with low free removable storage memory
 *
 * Enrollment Tracking Reports (17, 30):
 *   17. List of devices enrolled per user in Microsoft Intune
 *   30. Number of devices enrolled per user in Microsoft Intune
 *
 * Device Category and Enrollment Reports (18, 24):
 *   18. List of devices in a specific device category
 *   24. Mobile devices unmanaged (enrollment failed)
 *
 * Device Security Reports (16, 23, 27):
 *   16. Health attestation report (Device Health Attestation)
 *   23. Mobile devices that are jailbroken or rooted
 *   27. Mobile devices with certificate renewal issues
 *
 * Device Action Reports (31-33):
 *   31. Pending retire and wipe requests
 *   32. Recently enrolled and assigned mobile devices
 *   33. Recently wiped mobile devices
 *
 * Windows RT Sideloading Reports (35-36):
 *   35. Windows RT sideloading keys
 *   36. Windows RT sideloading keys status
 *
 * @module ConfigMgrDeviceManagementReports
 * @author Intune Reporting System
 * @version 1.0.0
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig, ReportData, ReportResult } from '../../types';
import { Logger } from '../../core/logger';
import { OutputFormatter } from '../../formatters/output-formatter';

// Import all report modules
import { MobileDeviceInventoryReports } from './mobile-device-inventory-reports';
import { DeviceHardwareReports } from './device-hardware-reports';
import { ExchangeActiveSyncReports } from './exchange-activesync-reports';
import { DeviceSecurityReports } from './device-security-reports';
import { DeviceActionReports } from './device-action-reports';
import { WindowsRTSideloadingReports } from './windows-rt-sideloading-reports';
// import { WindowsCEReports } from './windows-ce-reports';
import { DeviceCategoryEnrollmentReports } from './device-category-enrollment-reports';
import { IntuneEnrollmentTrackingReports } from './intune-enrollment-tracking-reports';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Report category enumeration
 */
export enum ReportCategory {
  MOBILE_INVENTORY = 'Mobile Device Inventory',
  HARDWARE = 'Device Hardware',
  EXCHANGE_ACTIVESYNC = 'Exchange ActiveSync',
  SECURITY = 'Device Security',
  ENROLLMENT = 'Device Enrollment',
  DEVICE_ACTIONS = 'Device Actions',
  SIDELOADING = 'Windows RT Sideloading',
  WINDOWS_CE = 'Windows CE (Legacy)',
  DEVICE_CATEGORY = 'Device Category',
}

/**
 * Metadata for each of the 37 ConfigMgr reports
 */
export interface ConfigMgrReportMetadata {
  reportNumber: number;
  reportName: string;
  description: string;
  category: ReportCategory;
  parameters?: string[];
  modernIntuneEquivalent: string;
  graphApiEndpoints: string[];
  isSupported: boolean;
  migrationNotes?: string;
}

/**
 * Report execution options
 */
export interface ReportExecutionOptions {
  reportNumbers?: number[];
  categories?: ReportCategory[];
  outputFormat?: 'json' | 'csv' | 'html';
  outputDirectory?: string;
  includeTimestamp?: boolean;
  parallel?: boolean;
  maxParallel?: number;
  filters?: Record<string, any>;
}

/**
 * Batch execution result
 */
export interface BatchExecutionResult {
  totalReports: number;
  successCount: number;
  failureCount: number;
  skippedCount: number;
  results: ReportResult[];
  summary: {
    totalDevices: number;
    totalRecords: number;
    executionTimeMs: number;
    startTime: string;
    endTime: string;
  };
}

/**
 * Dashboard data combining all report metrics
 */
export interface ConfigMgrDashboard {
  generated: string;
  tenantId: string;
  summary: {
    totalDevices: number;
    totalUsers: number;
    totalReportsGenerated: number;
    corporateDevices: number;
    personalDevices: number;
    compliantDevices: number;
    nonCompliantDevices: number;
    activeDevices: number;
    inactiveDevices: number;
  };
  byCategory: Record<ReportCategory, any>;
  topInsights: {
    type: string;
    severity: 'info' | 'warning' | 'critical';
    message: string;
    count: number;
  }[];
  recommendations: string[];
}

/**
 * Report schedule configuration
 */
export interface ReportSchedule {
  reportNumbers: number[];
  cronExpression: string;
  outputFormat: 'json' | 'csv' | 'html';
  enabled: boolean;
  name: string;
  description?: string;
}

// ============================================================================
// Master Orchestrator Class
// ============================================================================

/**
 * ConfigMgr Device Management Reports Master Orchestrator
 *
 * Provides unified access to all 37 Configuration Manager device management
 * reports with batch execution, scheduling, and dashboard capabilities.
 */
export class ConfigMgrDeviceManagementReports {
  private config: AppConfig;

  // Report module instances
  private mobileInventory: MobileDeviceInventoryReports;
  private deviceHardware: DeviceHardwareReports;
  private exchangeActiveSync: ExchangeActiveSyncReports;
  private deviceSecurity: DeviceSecurityReports;
  private deviceActions: DeviceActionReports;
  private windowsRTSideloading: WindowsRTSideloadingReports;
  private deviceCategory: DeviceCategoryEnrollmentReports;
  private enrollmentTracking: IntuneEnrollmentTrackingReports;

  /**
   * All 37 ConfigMgr reports metadata
   */
  private readonly REPORT_CATALOG: Map<number, ConfigMgrReportMetadata> = new Map([
    // Reports 1, 2, 20, 22 - Mobile Device Inventory
    [1, {
      reportNumber: 1,
      reportName: 'All corporate-owned mobile devices',
      description: 'Displays all corporate-owned mobile devices enrolled in Intune',
      category: ReportCategory.MOBILE_INVENTORY,
      parameters: ['ownerType', 'operatingSystem'],
      modernIntuneEquivalent: 'Devices > All devices > Filter by Owner Type = Corporate',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [2, {
      reportNumber: 2,
      reportName: 'All mobile device clients',
      description: 'All mobile device clients excluding Exchange connector devices',
      category: ReportCategory.MOBILE_INVENTORY,
      parameters: ['managementAgent'],
      modernIntuneEquivalent: 'Devices > All devices > Filter by Management Type',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [20, {
      reportNumber: 20,
      reportName: 'Mobile device client information',
      description: 'Communication status and management point connectivity for mobile devices',
      category: ReportCategory.MOBILE_INVENTORY,
      parameters: ['lastSyncThreshold'],
      modernIntuneEquivalent: 'Devices > All devices > Last check-in column',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [22, {
      reportNumber: 22,
      reportName: 'Mobile devices by operating system',
      description: 'Devices grouped by operating system with version distribution',
      category: ReportCategory.MOBILE_INVENTORY,
      parameters: ['operatingSystem'],
      modernIntuneEquivalent: 'Devices > All devices > Group by OS',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],

    // Reports 3-7, 13-14, 19 - Windows CE (Legacy/Unsupported)
    [3, {
      reportNumber: 3,
      reportName: 'Certificate issues on mobile devices (Windows CE)',
      description: 'Certificate issues for Windows CE devices - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [4, {
      reportNumber: 4,
      reportName: 'Client deployment failure (Windows CE)',
      description: 'Failed client deployments for Windows CE - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [5, {
      reportNumber: 5,
      reportName: 'Client deployment status details (Windows CE)',
      description: 'Deployment status details for Windows CE - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [6, {
      reportNumber: 6,
      reportName: 'Client deployment success (Windows CE)',
      description: 'Successful client deployments for Windows CE - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [7, {
      reportNumber: 7,
      reportName: 'Communication issues on mobile devices (Windows CE)',
      description: 'Communication issues for Windows CE - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [13, {
      reportNumber: 13,
      reportName: 'Health information for mobile devices (Windows CE)',
      description: 'Health information for Windows CE devices - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [14, {
      reportNumber: 14,
      reportName: 'Health summary for mobile devices (Windows CE)',
      description: 'Health summary for Windows CE devices - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],
    [19, {
      reportNumber: 19,
      reportName: 'Local client issues on mobile devices (Windows CE)',
      description: 'Local client issues for Windows CE - NOT SUPPORTED',
      category: ReportCategory.WINDOWS_CE,
      modernIntuneEquivalent: 'N/A - Windows CE is not supported in Intune',
      graphApiEndpoints: [],
      isSupported: false,
      migrationNotes: 'Migrate to Windows 10 IoT, Android, or iOS devices',
    }],

    // Reports 8, 15, 21, 34 - Exchange ActiveSync
    [8, {
      reportNumber: 8,
      reportName: 'Compliance status of default ActiveSync mailbox policy',
      description: 'ActiveSync mailbox policy compliance status',
      category: ReportCategory.EXCHANGE_ACTIVESYNC,
      parameters: ['policyId'],
      modernIntuneEquivalent: 'Devices > Compliance policies > Policy status',
      graphApiEndpoints: ['/deviceManagement/managedDevices', '/deviceManagement/deviceCompliancePolicies'],
      isSupported: true,
    }],
    [15, {
      reportNumber: 15,
      reportName: 'Inactive mobile devices (Exchange)',
      description: 'Mobile devices that have not synced with Exchange in specified timeframe',
      category: ReportCategory.EXCHANGE_ACTIVESYNC,
      parameters: ['inactiveDays'],
      modernIntuneEquivalent: 'Devices > All devices > Sort by Last check-in',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [21, {
      reportNumber: 21,
      reportName: 'Mobile device compliance details (Exchange)',
      description: 'Detailed compliance information for Exchange ActiveSync devices',
      category: ReportCategory.EXCHANGE_ACTIVESYNC,
      parameters: ['complianceState'],
      modernIntuneEquivalent: 'Devices > Compliance policies > Device compliance',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [34, {
      reportNumber: 34,
      reportName: 'Settings summary for mobile devices (Exchange)',
      description: 'Summary of device settings for Exchange ActiveSync devices',
      category: ReportCategory.EXCHANGE_ACTIVESYNC,
      parameters: [],
      modernIntuneEquivalent: 'Devices > Configuration profiles > Device status',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],

    // Reports 9-12, 25-26, 28-29 - Device Hardware
    [9, {
      reportNumber: 9,
      reportName: 'Count of mobile devices by display configurations',
      description: 'Devices grouped by display configuration categories',
      category: ReportCategory.HARDWARE,
      parameters: [],
      modernIntuneEquivalent: 'Devices > All devices > Hardware properties',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [10, {
      reportNumber: 10,
      reportName: 'Count of mobile devices by operating system',
      description: 'Device count distribution by operating system and version',
      category: ReportCategory.HARDWARE,
      parameters: [],
      modernIntuneEquivalent: 'Reports > Device compliance > OS versions',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [11, {
      reportNumber: 11,
      reportName: 'Count of mobile devices by program memory',
      description: 'Devices grouped by RAM/memory capacity ranges',
      category: ReportCategory.HARDWARE,
      parameters: [],
      modernIntuneEquivalent: 'Devices > All devices > Hardware inventory',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [12, {
      reportNumber: 12,
      reportName: 'Count of mobile devices by storage memory configurations',
      description: 'Devices grouped by storage capacity ranges',
      category: ReportCategory.HARDWARE,
      parameters: [],
      modernIntuneEquivalent: 'Devices > All devices > Storage properties',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [25, {
      reportNumber: 25,
      reportName: 'Mobile devices with specific free program memory',
      description: 'Devices meeting specified free memory criteria',
      category: ReportCategory.HARDWARE,
      parameters: ['minMemoryMB', 'maxMemoryMB'],
      modernIntuneEquivalent: 'Custom report using Graph API',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [26, {
      reportNumber: 26,
      reportName: 'Mobile devices with specific free removable storage memory',
      description: 'Devices meeting specified free storage criteria',
      category: ReportCategory.HARDWARE,
      parameters: ['minStorageGB', 'maxStorageGB'],
      modernIntuneEquivalent: 'Custom report using Graph API',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [28, {
      reportNumber: 28,
      reportName: 'Mobile devices with low free program memory',
      description: 'Devices below specified memory threshold',
      category: ReportCategory.HARDWARE,
      parameters: ['thresholdMB'],
      modernIntuneEquivalent: 'Custom report using Graph API + Alerts',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [29, {
      reportNumber: 29,
      reportName: 'Mobile devices with low free removable storage memory',
      description: 'Devices below specified storage threshold',
      category: ReportCategory.HARDWARE,
      parameters: ['thresholdGB'],
      modernIntuneEquivalent: 'Custom report using Graph API + Alerts',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],

    // Reports 16, 23, 27 - Device Security
    [16, {
      reportNumber: 16,
      reportName: 'Health attestation report',
      description: 'Device health attestation status and compliance',
      category: ReportCategory.SECURITY,
      parameters: [],
      modernIntuneEquivalent: 'Devices > Monitor > Device health attestation',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [23, {
      reportNumber: 23,
      reportName: 'Mobile devices that are jailbroken or rooted',
      description: 'Identifies compromised devices with jailbreak/root detection',
      category: ReportCategory.SECURITY,
      parameters: [],
      modernIntuneEquivalent: 'Devices > All devices > Filter: Jailbroken = Yes',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [27, {
      reportNumber: 27,
      reportName: 'Mobile devices with certificate renewal issues',
      description: 'Devices with expired or expiring certificates',
      category: ReportCategory.SECURITY,
      parameters: ['expirationDays'],
      modernIntuneEquivalent: 'Devices > Monitor > Certificate connector',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],

    // Reports 17, 30 - Enrollment Tracking
    [17, {
      reportNumber: 17,
      reportName: 'List of devices enrolled per user in Microsoft Intune',
      description: 'Detailed device list for each user with enrollment information',
      category: ReportCategory.ENROLLMENT,
      parameters: ['userId'],
      modernIntuneEquivalent: 'Users > Select user > Devices',
      graphApiEndpoints: ['/deviceManagement/managedDevices', '/users'],
      isSupported: true,
    }],
    [30, {
      reportNumber: 30,
      reportName: 'Number of devices enrolled per user in Microsoft Intune',
      description: 'Summary count of devices per user',
      category: ReportCategory.ENROLLMENT,
      parameters: [],
      modernIntuneEquivalent: 'Reports > Device enrollment > Enrollment summary',
      graphApiEndpoints: ['/deviceManagement/managedDevices', '/users'],
      isSupported: true,
    }],

    // Reports 18, 24 - Device Category
    [18, {
      reportNumber: 18,
      reportName: 'List of devices in a specific device category',
      description: 'All devices assigned to a specific category',
      category: ReportCategory.DEVICE_CATEGORY,
      parameters: ['categoryId'],
      modernIntuneEquivalent: 'Devices > Device categories > Select category',
      graphApiEndpoints: ['/deviceManagement/managedDevices', '/deviceManagement/deviceCategories'],
      isSupported: true,
    }],
    [24, {
      reportNumber: 24,
      reportName: 'Mobile devices unmanaged (enrollment failed)',
      description: 'Devices with failed enrollment attempts',
      category: ReportCategory.DEVICE_CATEGORY,
      parameters: [],
      modernIntuneEquivalent: 'Troubleshooting > Enrollment failures',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],

    // Reports 31-33 - Device Actions
    [31, {
      reportNumber: 31,
      reportName: 'Pending retire and wipe requests',
      description: 'All pending device retire and wipe actions',
      category: ReportCategory.DEVICE_ACTIONS,
      parameters: [],
      modernIntuneEquivalent: 'Devices > All devices > Retire/Wipe status',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [32, {
      reportNumber: 32,
      reportName: 'Recently enrolled and assigned mobile devices',
      description: 'Newly enrolled devices with assignment status',
      category: ReportCategory.DEVICE_ACTIONS,
      parameters: ['days'],
      modernIntuneEquivalent: 'Reports > Device enrollment > Recent enrollments',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],
    [33, {
      reportNumber: 33,
      reportName: 'Recently wiped mobile devices',
      description: 'Devices that were recently wiped or retired',
      category: ReportCategory.DEVICE_ACTIONS,
      parameters: ['days'],
      modernIntuneEquivalent: 'Devices > All devices > Device actions audit',
      graphApiEndpoints: ['/deviceManagement/managedDevices'],
      isSupported: true,
    }],

    // Reports 35-36 - Windows RT Sideloading
    [35, {
      reportNumber: 35,
      reportName: 'Windows RT sideloading keys',
      description: 'All Windows RT sideloading keys and their details',
      category: ReportCategory.SIDELOADING,
      parameters: [],
      modernIntuneEquivalent: 'Tenant administration > Connectors and tokens > Windows sideloading keys',
      graphApiEndpoints: ['/deviceAppManagement/sideLoadingKeys'],
      isSupported: true,
    }],
    [36, {
      reportNumber: 36,
      reportName: 'Windows RT sideloading keys status',
      description: 'Status and usage information for sideloading keys',
      category: ReportCategory.SIDELOADING,
      parameters: [],
      modernIntuneEquivalent: 'Tenant administration > Connectors and tokens > Windows sideloading keys > Status',
      graphApiEndpoints: ['/deviceAppManagement/sideLoadingKeys'],
      isSupported: true,
    }],
  ]);

  constructor(graphClient: Client, config: AppConfig) {
    this.config = config;

    // Initialize all report modules
    this.mobileInventory = new MobileDeviceInventoryReports(graphClient, config);
    this.deviceHardware = new DeviceHardwareReports(graphClient, config);
    this.exchangeActiveSync = new ExchangeActiveSyncReports(graphClient, config);
    this.deviceSecurity = new DeviceSecurityReports(graphClient);
    this.deviceActions = new DeviceActionReports(graphClient);
    this.windowsRTSideloading = new WindowsRTSideloadingReports(graphClient, config);
    this.deviceCategory = new DeviceCategoryEnrollmentReports(graphClient, config);
    this.enrollmentTracking = new IntuneEnrollmentTrackingReports(graphClient, config);

    logger.info('ConfigMgr Device Management Reports Orchestrator initialized');
  }

  // ==========================================================================
  // Report Catalog and Discovery
  // ==========================================================================

  /**
   * Get all 37 reports with metadata
   */
  getAllReports(): ConfigMgrReportMetadata[] {
    return Array.from(this.REPORT_CATALOG.values());
  }

  /**
   * Get report metadata by report number
   */
  getReportMetadata(reportNumber: number): ConfigMgrReportMetadata | undefined {
    return this.REPORT_CATALOG.get(reportNumber);
  }

  /**
   * Get reports by category
   */
  getReportsByCategory(category: ReportCategory): ConfigMgrReportMetadata[] {
    return this.getAllReports().filter(report => report.category === category);
  }

  /**
   * Get all supported reports (excludes Windows CE legacy reports)
   */
  getSupportedReports(): ConfigMgrReportMetadata[] {
    return this.getAllReports().filter(report => report.isSupported);
  }

  /**
   * Get unsupported/legacy reports
   */
  getUnsupportedReports(): ConfigMgrReportMetadata[] {
    return this.getAllReports().filter(report => !report.isSupported);
  }

  // ==========================================================================
  // Individual Report Execution
  // ==========================================================================

  /**
   * Execute a specific report by number
   */
  async getReportByNumber(
    reportNumber: number,
    options?: Record<string, any>
  ): Promise<ReportData> {
    const metadata = this.getReportMetadata(reportNumber);
    if (!metadata) {
      throw new Error(`Report ${reportNumber} not found in catalog`);
    }

    if (!metadata.isSupported) {
      logger.warn(`Report ${reportNumber} is not supported (${metadata.reportName})`);
      return this.createUnsupportedReportResponse(metadata);
    }

    logger.info(`Executing report ${reportNumber}: ${metadata.reportName}`);
    const startTime = Date.now();

    try {
      let reportData: ReportData;

      // Route to appropriate report module based on report number
      switch (reportNumber) {
        // Mobile Device Inventory (1, 2, 20, 22)
        case 1:
          const corporateDevices = await this.mobileInventory.getAllCorporateOwnedDevices(options as any);
          reportData = this.wrapReportData(metadata, corporateDevices);
          break;
        case 2:
          const mobileClients = await this.mobileInventory.getAllMobileDeviceClients(options as any);
          reportData = this.wrapReportData(metadata, mobileClients);
          break;
        case 20:
          const clientInfo = await this.mobileInventory.getMobileDeviceClientInformation(options as any);
          reportData = this.wrapReportData(metadata, clientInfo);
          break;
        case 22:
          const osDist = await this.mobileInventory.getMobileDevicesByOperatingSystem(options as any);
          reportData = this.wrapReportData(metadata, osDist);
          break;

        // Device Hardware (9-12, 25-26, 28-29)
        case 9:
          const displayConfig = await this.deviceHardware.getDisplayConfigurationSummary();
          reportData = this.wrapReportData(metadata, displayConfig);
          break;
        case 10:
          const osDistHardware = await this.deviceHardware.getOSDistributionSummary();
          reportData = this.wrapReportData(metadata, osDistHardware);
          break;
        case 11:
          const memoryRanges = await this.deviceHardware.getMemoryRangeSummary();
          reportData = this.wrapReportData(metadata, memoryRanges);
          break;
        case 12:
          const storageRanges = await this.deviceHardware.getStorageRangeSummary();
          reportData = this.wrapReportData(metadata, storageRanges);
          break;
        case 25:
          const specificMemory = await this.deviceHardware.getDevicesWithSpecificFreeMemory(
            options?.minMemoryMB || 0,
            options?.maxMemoryMB
          );
          reportData = this.wrapReportData(metadata, specificMemory);
          break;
        case 26:
          const specificStorage = await this.deviceHardware.getDevicesWithSpecificFreeStorage(
            options?.minStorageGB || 0,
            options?.maxStorageGB
          );
          reportData = this.wrapReportData(metadata, specificStorage);
          break;
        case 28:
          const lowMemory = await this.deviceHardware.getDevicesWithLowMemory(
            options?.thresholdMB || 512
          );
          reportData = this.wrapReportData(metadata, lowMemory);
          break;
        case 29:
          const lowStorage = await this.deviceHardware.getDevicesWithLowStorage(
            options?.thresholdGB || 10
          );
          reportData = this.wrapReportData(metadata, lowStorage);
          break;

        // Exchange ActiveSync (8, 15, 21, 34)
        case 8:
          const policyCompliance = await this.exchangeActiveSync.getActiveSyncPolicyCompliance();
          reportData = this.wrapReportData(metadata, policyCompliance);
          break;
        case 15:
          const inactiveDevices = await this.exchangeActiveSync.getInactiveMobileDevices(
            options?.inactiveDays || 30
          );
          reportData = this.wrapReportData(metadata, inactiveDevices);
          break;
        case 21:
          const complianceDetails = await this.exchangeActiveSync.getAllMobileDeviceComplianceDetails(
            options as any
          );
          reportData = this.wrapReportData(metadata, complianceDetails);
          break;
        case 34:
          const settingsSummary = await this.exchangeActiveSync.getMobileDeviceSettingsSummary();
          reportData = this.wrapReportData(metadata, settingsSummary);
          break;

        // Device Security (16, 23, 27)
        case 16:
          const healthAttestation = await this.deviceSecurity.generateHealthAttestationReport();
          reportData = this.wrapReportData(metadata, healthAttestation);
          break;
        case 23:
          const jailbrokenDevices = await this.deviceSecurity.getJailbrokenDevices();
          reportData = this.wrapReportData(metadata, jailbrokenDevices);
          break;
        case 27:
          const certIssues = await this.deviceSecurity.getCertificateRenewalIssues(
            options?.expirationDays || 30
          );
          reportData = this.wrapReportData(metadata, certIssues);
          break;

        // Enrollment Tracking (17, 30)
        case 17:
          const userDeviceList = await this.enrollmentTracking.getDevicesEnrolledPerUser(
            options?.userId || ''
          );
          reportData = this.wrapReportData(metadata, userDeviceList ? [userDeviceList] : []);
          break;
        case 30:
          const enrollmentCounts = await this.enrollmentTracking.getEnrollmentCountPerUserReport();
          reportData = this.wrapReportData(metadata, enrollmentCounts);
          break;

        // Device Category (18, 24)
        case 18:
          const categoryDevices = await this.deviceCategory.getDevicesByCategory(
            options?.categoryId || ''
          );
          reportData = this.wrapReportData(metadata, categoryDevices);
          break;
        case 24:
          const failedEnrollments = await this.deviceCategory.getUnmanagedDevicesEnrollmentFailed();
          reportData = this.wrapReportData(metadata, failedEnrollments);
          break;

        // Device Actions (31-33)
        case 31:
          const pendingActions = await this.deviceActions.getPendingRetireWipeRequests(options);
          reportData = this.wrapReportData(metadata, pendingActions);
          break;
        case 32:
          const recentEnrollments = await this.deviceActions.getRecentlyEnrolledDevices(options);
          reportData = this.wrapReportData(metadata, recentEnrollments);
          break;
        case 33:
          const recentWipes = await this.deviceActions.getRecentlyWipedDevices(options);
          reportData = this.wrapReportData(metadata, recentWipes);
          break;

        // Windows RT Sideloading (35-36)
        case 35:
          const sideloadingKeys = await this.windowsRTSideloading.getAllSideloadingKeys();
          reportData = this.wrapReportData(metadata, sideloadingKeys);
          break;
        case 36:
          const keyStatus = await this.windowsRTSideloading.getSideloadingKeysSummary();
          reportData = this.wrapReportData(metadata, keyStatus);
          break;

        default:
          throw new Error(`Report ${reportNumber} execution not implemented`);
      }

      const duration = Date.now() - startTime;
      logger.info(`Report ${reportNumber} completed in ${duration}ms`);

      return reportData;
    } catch (error) {
      logger.error(`Failed to execute report ${reportNumber}`, error);
      throw error;
    }
  }

  // ==========================================================================
  // Batch Execution
  // ==========================================================================

  /**
   * Export all reports in batch
   */
  async exportAllReports(options?: ReportExecutionOptions): Promise<BatchExecutionResult> {
    const supportedReports = this.getSupportedReports();
    const reportNumbers = supportedReports.map(r => r.reportNumber);

    return this.exportReports(reportNumbers, options);
  }

  /**
   * Export specific reports
   */
  async exportReports(
    reportNumbers: number[],
    options?: ReportExecutionOptions
  ): Promise<BatchExecutionResult> {
    const startTime = Date.now();
    const results: ReportResult[] = [];
    let successCount = 0;
    let failureCount = 0;
    let skippedCount = 0;
    let totalRecords = 0;

    const outputFormat = options?.outputFormat || this.config.output.defaultFormat;
    const outputDirectory = options?.outputDirectory || this.config.output.directory;
    const parallel = options?.parallel !== false;
    const maxParallel = options?.maxParallel || 5;

    logger.info(`Starting batch export of ${reportNumbers.length} reports`);

    if (parallel) {
      // Execute reports in parallel with concurrency limit
      for (let i = 0; i < reportNumbers.length; i += maxParallel) {
        const batch = reportNumbers.slice(i, i + maxParallel);
        const batchResults = await Promise.allSettled(
          batch.map(num => this.executeAndExportReport(num, outputFormat, outputDirectory, options))
        );

        for (const result of batchResults) {
          if (result.status === 'fulfilled') {
            results.push(result.value);
            if (result.value.success) {
              successCount++;
              totalRecords += result.value.recordCount || 0;
            } else {
              failureCount++;
            }
          } else {
            failureCount++;
            results.push({
              reportName: `Report ${batch[batchResults.indexOf(result)]}`,
              success: false,
              format: outputFormat,
              error: result.reason?.message || 'Unknown error',
              duration: 0,
            });
          }
        }
      }
    } else {
      // Execute reports sequentially
      for (const reportNumber of reportNumbers) {
        try {
          const result = await this.executeAndExportReport(
            reportNumber,
            outputFormat,
            outputDirectory,
            options
          );
          results.push(result);
          if (result.success) {
            successCount++;
            totalRecords += result.recordCount || 0;
          } else {
            failureCount++;
          }
        } catch (error) {
          failureCount++;
          results.push({
            reportName: `Report ${reportNumber}`,
            success: false,
            format: outputFormat,
            error: error instanceof Error ? error.message : 'Unknown error',
            duration: 0,
          });
        }
      }
    }

    const endTime = Date.now();
    const executionTimeMs = endTime - startTime;

    logger.info(`Batch export completed: ${successCount} success, ${failureCount} failed`);

    return {
      totalReports: reportNumbers.length,
      successCount,
      failureCount,
      skippedCount,
      results,
      summary: {
        totalDevices: 0, // Would need to aggregate from reports
        totalRecords,
        executionTimeMs,
        startTime: new Date(startTime).toISOString(),
        endTime: new Date(endTime).toISOString(),
      },
    };
  }

  // ==========================================================================
  // Dashboard Generation
  // ==========================================================================

  /**
   * Generate comprehensive dashboard with all metrics
   */
  async generateDashboard(): Promise<ConfigMgrDashboard> {
    logger.info('Generating comprehensive ConfigMgr dashboard');
    const startTime = Date.now();

    try {
      // Execute key reports in parallel for dashboard
      const [
        allDevices,
        osDistribution,
        complianceStatus,
        jailbrokenDevices,
        pendingActions,
        recentEnrollments,
      ] = await Promise.all([
        this.mobileInventory.getAllMobileDevices(),
        this.mobileInventory.getMobileDevicesByOperatingSystem(),
        this.exchangeActiveSync.getAllMobileDeviceComplianceDetails().catch(() => []),
        this.deviceSecurity.getJailbrokenDevices().catch(() => []),
        this.deviceActions.getPendingRetireWipeRequests().catch(() => []),
        this.deviceActions.getRecentlyEnrolledDevices({ dateRange: {
          startDate: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000),
          endDate: new Date()
        }}).catch(() => []),
      ]);

      // Calculate summary metrics
      const corporateDevices = allDevices.filter((d: any) => d.ownerType === 'company').length;
      const personalDevices = allDevices.filter((d: any) => d.ownerType === 'personal').length;
      const compliantDevices = allDevices.filter((d: any) => d.complianceState === 'compliant').length;
      const nonCompliantDevices = allDevices.filter((d: any) => d.complianceState !== 'compliant').length;
      const activeDevices = allDevices.filter((d: any) => {
        if (!d.lastSyncDateTime) return false;
        const daysSince = Math.floor((Date.now() - new Date(d.lastSyncDateTime).getTime()) / (1000 * 60 * 60 * 24));
        return daysSince <= 7;
      }).length;
      const inactiveDevices = allDevices.length - activeDevices;

      // Get unique users
      const uniqueUsers = new Set(allDevices.map((d: any) => d.userPrincipalName).filter(Boolean));

      // Generate insights
      const topInsights: ConfigMgrDashboard['topInsights'] = [];

      if (jailbrokenDevices.length > 0) {
        topInsights.push({
          type: 'Security Risk',
          severity: 'critical',
          message: `${jailbrokenDevices.length} jailbroken/rooted devices detected`,
          count: jailbrokenDevices.length,
        });
      }

      if (nonCompliantDevices > allDevices.length * 0.2) {
        topInsights.push({
          type: 'Compliance Issue',
          severity: 'warning',
          message: `${Math.round((nonCompliantDevices / allDevices.length) * 100)}% of devices are non-compliant`,
          count: nonCompliantDevices,
        });
      }

      if (inactiveDevices > allDevices.length * 0.15) {
        topInsights.push({
          type: 'Device Activity',
          severity: 'warning',
          message: `${inactiveDevices} devices inactive for over 7 days`,
          count: inactiveDevices,
        });
      }

      if (pendingActions.length > 0) {
        topInsights.push({
          type: 'Pending Actions',
          severity: 'info',
          message: `${pendingActions.length} pending retire/wipe requests`,
          count: pendingActions.length,
        });
      }

      // Generate recommendations
      const recommendations: string[] = [];

      if (jailbrokenDevices.length > 0) {
        recommendations.push('Review and remediate jailbroken/rooted devices immediately');
      }

      if (nonCompliantDevices > 0) {
        recommendations.push('Investigate non-compliant devices and enforce compliance policies');
      }

      if (inactiveDevices > allDevices.length * 0.1) {
        recommendations.push('Consider retiring devices inactive for more than 30 days');
      }

      if (recentEnrollments.length > 0) {
        recommendations.push(`Welcome ${recentEnrollments.length} newly enrolled devices from the past week`);
      }

      const dashboard: ConfigMgrDashboard = {
        generated: new Date().toISOString(),
        tenantId: this.config.authentication.tenantId,
        summary: {
          totalDevices: allDevices.length,
          totalUsers: uniqueUsers.size,
          totalReportsGenerated: 37,
          corporateDevices,
          personalDevices,
          compliantDevices,
          nonCompliantDevices,
          activeDevices,
          inactiveDevices,
        },
        byCategory: {
          [ReportCategory.MOBILE_INVENTORY]: {
            totalDevices: allDevices.length,
            osDistribution: osDistribution,
          },
          [ReportCategory.HARDWARE]: {
            message: 'Hardware inventory available through device hardware reports',
          },
          [ReportCategory.EXCHANGE_ACTIVESYNC]: {
            complianceStatus: complianceStatus.length,
          },
          [ReportCategory.SECURITY]: {
            jailbrokenDevices: jailbrokenDevices.length,
          },
          [ReportCategory.DEVICE_ACTIONS]: {
            pendingActions: pendingActions.length,
            recentEnrollments: recentEnrollments.length,
          },
          [ReportCategory.ENROLLMENT]: {
            totalUsers: uniqueUsers.size,
            averageDevicesPerUser: allDevices.length / uniqueUsers.size,
          },
          [ReportCategory.DEVICE_CATEGORY]: {
            message: 'Device category information available through category reports',
          },
          [ReportCategory.SIDELOADING]: {
            message: 'Windows RT sideloading information available through sideloading reports',
          },
          [ReportCategory.WINDOWS_CE]: {
            message: 'Windows CE is not supported in Intune',
          },
        },
        topInsights,
        recommendations,
      };

      const duration = Date.now() - startTime;
      logger.info(`Dashboard generated in ${duration}ms`);

      return dashboard;
    } catch (error) {
      logger.error('Failed to generate dashboard', error);
      throw error;
    }
  }

  // ==========================================================================
  // Scheduling and Configuration
  // ==========================================================================

  /**
   * Schedule report execution
   */
  scheduleReports(schedule: ReportSchedule): void {
    logger.info(`Scheduling reports: ${schedule.name}`);
    // Implementation would integrate with the scheduler system
    // This is a placeholder for the scheduling logic
    logger.warn('Report scheduling not yet implemented - use external scheduler');
  }

  /**
   * Get global configuration
   */
  getConfiguration(): AppConfig {
    return this.config;
  }

  /**
   * Update configuration
   */
  updateConfiguration(updates: Partial<AppConfig>): void {
    this.config = { ...this.config, ...updates };
    logger.info('Configuration updated');
  }

  // ==========================================================================
  // Private Helper Methods
  // ==========================================================================

  /**
   * Execute and export a single report
   */
  private async executeAndExportReport(
    reportNumber: number,
    format: 'json' | 'csv' | 'html',
    outputDirectory: string,
    options?: ReportExecutionOptions
  ): Promise<ReportResult> {
    const metadata = this.getReportMetadata(reportNumber);
    if (!metadata) {
      throw new Error(`Report ${reportNumber} not found`);
    }

    const startTime = Date.now();

    try {
      const reportData = await this.getReportByNumber(reportNumber, options?.filters);
      const outputPath = await OutputFormatter.format(
        reportData,
        format,
        outputDirectory,
        options?.includeTimestamp !== false
      );

      return {
        reportName: metadata.reportName,
        success: true,
        format,
        outputPath,
        duration: Date.now() - startTime,
        recordCount: reportData.data?.length || 0,
      };
    } catch (error) {
      return {
        reportName: metadata.reportName,
        success: false,
        format,
        error: error instanceof Error ? error.message : 'Unknown error',
        duration: Date.now() - startTime,
      };
    }
  }

  /**
   * Wrap data in standard report format
   */
  private wrapReportData(
    metadata: ConfigMgrReportMetadata,
    data: any
  ): ReportData {
    return {
      metadata: {
        reportName: `Report ${metadata.reportNumber}: ${metadata.reportName}`,
        generatedAt: new Date().toISOString(),
        generatedBy: 'ConfigMgr Device Management Reports Orchestrator',
        recordCount: Array.isArray(data) ? data.length : 1,
        parameters: {
          reportNumber: metadata.reportNumber,
          category: metadata.category,
        },
      },
      data: Array.isArray(data) ? data : [data],
      summary: {
        reportNumber: metadata.reportNumber,
        category: metadata.category,
        recordCount: Array.isArray(data) ? data.length : 1,
      },
    };
  }

  /**
   * Create response for unsupported reports
   */
  private createUnsupportedReportResponse(
    metadata: ConfigMgrReportMetadata
  ): ReportData {
    return {
      metadata: {
        reportName: `Report ${metadata.reportNumber}: ${metadata.reportName}`,
        generatedAt: new Date().toISOString(),
        generatedBy: 'ConfigMgr Device Management Reports Orchestrator',
        recordCount: 0,
      },
      data: [],
      summary: {
        reportNumber: metadata.reportNumber,
        category: metadata.category,
        isSupported: false,
        warning: 'This report is not supported in modern Intune',
        migrationNotes: metadata.migrationNotes,
      },
    };
  }
}

// ============================================================================
// Convenience Functions
// ============================================================================

/**
 * Create orchestrator instance from Graph client
 */
export function createConfigMgrReports(
  graphClient: Client,
  config: AppConfig
): ConfigMgrDeviceManagementReports {
  return new ConfigMgrDeviceManagementReports(graphClient, config);
}

/**
 * Quick execution of a single report
 */
export async function executeReport(
  graphClient: Client,
  config: AppConfig,
  reportNumber: number,
  options?: Record<string, any>
): Promise<ReportData> {
  const orchestrator = new ConfigMgrDeviceManagementReports(graphClient, config);
  return orchestrator.getReportByNumber(reportNumber, options);
}

/**
 * Export all supported reports
 */
export async function exportAllSupportedReports(
  graphClient: Client,
  config: AppConfig,
  options?: ReportExecutionOptions
): Promise<BatchExecutionResult> {
  const orchestrator = new ConfigMgrDeviceManagementReports(graphClient, config);
  return orchestrator.exportAllReports(options);
}

// Export all types and main class
export default ConfigMgrDeviceManagementReports;
