/**
 * Device Category Reports Module for Microsoft Intune
 *
 * This module replicates Configuration Manager device category reports for Intune environments.
 * It provides comprehensive device categorization, enrollment failure tracking, and troubleshooting
 * capabilities using the Microsoft Graph API.
 *
 * Configuration Manager Report Equivalents:
 * - Report 18: List of devices in a specific device category
 * - Report 24: Mobile devices unmanaged (enrollment failed)
 *
 * Features:
 * - Device category management and listing
 * - Devices by category with full details
 * - Enrollment failure detection and analysis
 * - Management state monitoring
 * - Certificate and enrollment issues identification
 * - Comprehensive troubleshooting recommendations
 * - Multiple export formats (JSON, CSV, HTML)
 * - Pagination support for large datasets
 * - Retry logic with exponential backoff
 *
 * @module device-category-reports
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { BaseReport } from '../base-report';
import { ReportData, AppConfig } from '../../types';
import { Logger } from '../../core/logger';
import { OutputFormatter } from '../../formatters/output-formatter';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Device category information
 */
export interface DeviceCategory {
  id: string;
  displayName: string;
  description?: string;
  deviceCount?: number;
  lastModifiedDateTime?: string;
  roleScopeTagIds?: string[];
}

/**
 * Device information with category
 */
export interface CategorizedDevice {
  deviceId: string;
  deviceName: string;
  deviceCategoryId?: string;
  deviceCategoryDisplayName?: string;
  userPrincipalName?: string;
  userDisplayName?: string;
  userId?: string;
  platform: string;
  osVersion: string;
  model?: string;
  manufacturer?: string;
  serialNumber?: string;
  enrolledDateTime: string;
  lastSyncDateTime: string;
  managementState: string;
  complianceState: string;
  ownerType: string;
  enrollmentType?: string;
  isSupervised?: boolean;
  azureADDeviceId?: string;
  ethernetMacAddress?: string;
  wiFiMacAddress?: string;
}

/**
 * Device category summary statistics
 */
export interface DeviceCategorySummary {
  totalCategories: number;
  totalCategorizedDevices: number;
  totalUncategorizedDevices: number;
  categoriesWithDevices: number;
  emptyCategories: number;
  categoryBreakdown: CategoryBreakdown[];
  topCategories: CategoryCount[];
  platformDistribution: Record<string, number>;
  lastUpdated: string;
}

/**
 * Category breakdown details
 */
export interface CategoryBreakdown {
  categoryId: string;
  categoryName: string;
  deviceCount: number;
  percentage: number;
  platforms: Record<string, number>;
  complianceStats: {
    compliant: number;
    nonCompliant: number;
    unknown: number;
  };
  managementStates: Record<string, number>;
}

/**
 * Category device count
 */
export interface CategoryCount {
  categoryName: string;
  count: number;
  percentage: number;
}

/**
 * Unmanaged device (enrollment failed)
 */
export interface UnmanagedDevice {
  deviceId?: string;
  deviceName?: string;
  userPrincipalName: string;
  userId: string;
  userDisplayName?: string;
  platform: string;
  osVersion?: string;
  model?: string;
  manufacturer?: string;
  serialNumber?: string;
  enrollmentDate?: string;
  lastContactDate?: string;
  failureType: EnrollmentFailureType;
  failureReason: string;
  errorCode?: string;
  managementState?: string;
  enrollmentState?: string;
  certificateStatus?: CertificateStatus;
  daysSinceEnrollment?: number;
  troubleshootingSteps: string[];
  additionalDetails?: Record<string, any>;
}

/**
 * Enrollment failure types
 */
export enum EnrollmentFailureType {
  EnrollmentFailed = 'Enrollment Failed',
  SiteAssignmentFailed = 'Site Assignment Failed',
  ManagementStateFailed = 'Management State Failed',
  CertificateIssueFailed = 'Certificate Issue Failed',
  NoSuccessfulSync = 'No Successful Sync',
  RegistrationOnly = 'Registration Only - Not Managed',
  PolicyApplyFailed = 'Policy Apply Failed',
  Unknown = 'Unknown'
}

/**
 * Certificate status for enrolled devices
 */
export enum CertificateStatus {
  Valid = 'Valid',
  Expired = 'Expired',
  NotIssued = 'Not Issued',
  Revoked = 'Revoked',
  PendingIssuance = 'Pending Issuance',
  Unknown = 'Unknown'
}

/**
 * Unmanaged devices summary
 */
export interface UnmanagedDevicesSummary {
  totalUnmanagedDevices: number;
  failureTypeBreakdown: Record<EnrollmentFailureType, number>;
  platformBreakdown: Record<string, number>;
  affectedUsers: number;
  oldestFailure?: string;
  newestFailure?: string;
  averageDaysSinceEnrollment: number;
  certificateIssues: number;
  managementStateIssues: number;
  syncIssues: number;
  lastUpdated: string;
}

/**
 * Options for device category reports
 */
export interface DeviceCategoryReportOptions {
  categoryId?: string;
  categoryName?: string;
  includeUncategorized?: boolean;
  platformFilter?: string[];
  complianceFilter?: string[];
  managementStateFilter?: string[];
  includeDetails?: boolean;
  limit?: number;
}

/**
 * Options for unmanaged devices report
 */
export interface UnmanagedDevicesReportOptions {
  startDate?: Date;
  endDate?: Date;
  platformFilter?: string[];
  failureTypeFilter?: EnrollmentFailureType[];
  includeOldDevices?: boolean;
  daysSinceEnrollmentThreshold?: number;
  limit?: number;
}

/**
 * Complete device category report
 */
export interface DeviceCategoryReport {
  reportDate: string;
  summary: DeviceCategorySummary;
  categories: DeviceCategory[];
  categorizedDevices: CategorizedDevice[];
  uncategorizedDevices: CategorizedDevice[];
}

/**
 * Complete unmanaged devices report
 */
export interface UnmanagedDevicesReport {
  reportDate: string;
  summary: UnmanagedDevicesSummary;
  unmanagedDevices: UnmanagedDevice[];
  recommendations: string[];
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Determine the enrollment failure type based on device state
 */
function determineFailureType(device: any): EnrollmentFailureType {
  const managementState = device.managementState?.toLowerCase() || '';
  const enrollmentState = device.enrollmentState?.toLowerCase() || '';
  const lastSync = device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : null;
  const enrolled = device.enrolledDateTime ? new Date(device.enrolledDateTime) : null;

  // Check for management state issues
  if (managementState === 'managedwitherrors' || managementState === 'error') {
    return EnrollmentFailureType.ManagementStateFailed;
  }

  // Check for registration-only state (not fully managed)
  if (managementState === 'registered' || enrollmentState === 'registered') {
    return EnrollmentFailureType.RegistrationOnly;
  }

  // Check for policy application failures
  if (managementState === 'policyapplyfailed') {
    return EnrollmentFailureType.PolicyApplyFailed;
  }

  // Check for certificate issues
  if (device.certificateExpirationDateTime) {
    const certExpiry = new Date(device.certificateExpirationDateTime);
    if (certExpiry < new Date()) {
      return EnrollmentFailureType.CertificateIssueFailed;
    }
  }

  // Check for no successful sync after enrollment
  if (enrolled && lastSync) {
    const daysSinceEnrollment = Math.floor((new Date().getTime() - enrolled.getTime()) / (1000 * 60 * 60 * 24));
    const daysSinceSync = Math.floor((new Date().getTime() - lastSync.getTime()) / (1000 * 60 * 60 * 24));

    if (daysSinceEnrollment > 1 && daysSinceSync > 7) {
      return EnrollmentFailureType.NoSuccessfulSync;
    }
  }

  // Check for site assignment failure (no last sync)
  if (enrolled && !lastSync) {
    return EnrollmentFailureType.SiteAssignmentFailed;
  }

  // Default to enrollment failed
  if (managementState === 'failed' || enrollmentState === 'failed') {
    return EnrollmentFailureType.EnrollmentFailed;
  }

  return EnrollmentFailureType.Unknown;
}

/**
 * Get troubleshooting steps for a specific failure type
 */
function getTroubleshootingSteps(failureType: EnrollmentFailureType, errorCode?: string): string[] {
  const baseSteps = [
    'Review the device enrollment status in Microsoft Intune admin center',
    'Check the user\'s license assignment in Azure AD',
    'Verify device enrollment restrictions are not blocking the device',
    'Review Conditional Access policies that may affect enrollment'
  ];

  const specificSteps: Record<EnrollmentFailureType, string[]> = {
    [EnrollmentFailureType.EnrollmentFailed]: [
      'Verify network connectivity between device and Intune service endpoints',
      'Check that required Intune URLs are accessible from the device',
      'Ensure device meets minimum OS version requirements',
      'Review enrollment failure details in Intune troubleshooting blade',
      'Attempt re-enrollment after addressing any identified issues',
      'Check for conflicting MDM enrollment (e.g., existing enrollment)'
    ],
    [EnrollmentFailureType.SiteAssignmentFailed]: [
      'Verify the device successfully completed initial enrollment',
      'Check Intune service connectivity from the device',
      'Review Azure AD device registration status',
      'Ensure device can reach configuration.microsoft.com',
      'Check for network proxy or firewall blocking Intune endpoints',
      'Verify device time and date settings are accurate',
      'Review device enrollment profile assignments'
    ],
    [EnrollmentFailureType.ManagementStateFailed]: [
      'Check device management agent service status',
      'Verify Intune Management Extension is installed and running',
      'Review device event logs for management agent errors',
      'Ensure device can communicate with Intune service',
      'Check for policy conflicts or corrupt policy cache',
      'Consider retiring and re-enrolling the device',
      'Verify device has sufficient disk space and resources'
    ],
    [EnrollmentFailureType.CertificateIssueFailed]: [
      'Verify NDES/SCEP configuration is correct',
      'Check certificate connector service is running',
      'Review certificate profiles in Intune admin center',
      'Ensure device can reach certificate authority endpoints',
      'Check certificate template permissions in AD CS',
      'Review NDES connector logs for certificate request failures',
      'Verify the certificate has not expired',
      'Check device certificate store for corrupt certificates'
    ],
    [EnrollmentFailureType.NoSuccessfulSync]: [
      'Verify device has active internet connectivity',
      'Check Intune service health status in Microsoft 365 admin center',
      'Ensure device is powered on and not in sleep mode',
      'Review device firewall settings for Intune endpoints',
      'Check if device maximum sync interval has been exceeded',
      'Trigger manual sync from Company Portal or Settings app',
      'Review device management agent logs for sync errors',
      'Verify device certificate is valid and not expired'
    ],
    [EnrollmentFailureType.RegistrationOnly]: [
      'Review device enrollment restrictions and limitations',
      'Check if device platform supports full management',
      'Verify user has appropriate licenses for device management',
      'Ensure enrollment profile is configured for full management',
      'Check for MAM-only policies preventing full MDM enrollment',
      'Review device type restrictions in enrollment settings',
      'Consider configuring proper MDM enrollment for the device'
    ],
    [EnrollmentFailureType.PolicyApplyFailed]: [
      'Review assigned policies for conflicts or errors',
      'Check device compliance policy configuration',
      'Verify configuration profiles are compatible with device OS',
      'Review policy application logs on the device',
      'Check for corrupt policy cache on device',
      'Ensure device has resources to apply policies',
      'Consider reassigning policies to the device',
      'Review policy assignment scope and filters'
    ],
    [EnrollmentFailureType.Unknown]: [
      'Review device and user enrollment history in Intune',
      'Check Microsoft Intune troubleshooting blade for details',
      'Verify device appears in Azure AD and Intune device lists',
      'Review audit logs for enrollment-related events',
      'Check for service health issues in Microsoft 365 admin center',
      'Contact Microsoft support with device details and error codes'
    ]
  };

  const steps = [...baseSteps, ...(specificSteps[failureType] || specificSteps[EnrollmentFailureType.Unknown])];

  // Add error-code specific steps if available
  if (errorCode) {
    steps.unshift(`Research error code ${errorCode} in Microsoft documentation`);
  }

  return steps;
}

/**
 * Determine certificate status from device data
 */
function determineCertificateStatus(device: any): CertificateStatus {
  if (!device.certificateExpirationDateTime) {
    return CertificateStatus.NotIssued;
  }

  const certExpiry = new Date(device.certificateExpirationDateTime);
  const now = new Date();

  if (certExpiry < now) {
    return CertificateStatus.Expired;
  }

  const daysUntilExpiry = Math.floor((certExpiry.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));

  if (daysUntilExpiry < 30) {
    return CertificateStatus.PendingIssuance; // Approaching expiry
  }

  return CertificateStatus.Valid;
}

/**
 * Calculate days since enrollment
 */
function calculateDaysSinceEnrollment(enrolledDateTime?: string): number | undefined {
  if (!enrolledDateTime) return undefined;

  const enrolled = new Date(enrolledDateTime);
  const now = new Date();
  return Math.floor((now.getTime() - enrolled.getTime()) / (1000 * 60 * 60 * 24));
}

/**
 * Format device for category report
 */
function formatCategorizedDevice(device: any): CategorizedDevice {
  return {
    deviceId: device.id,
    deviceName: device.deviceName || 'Unknown',
    deviceCategoryId: device.deviceCategoryDisplayName ? device.deviceCategoryDisplayName : undefined,
    deviceCategoryDisplayName: device.deviceCategoryDisplayName || 'Uncategorized',
    userPrincipalName: device.userPrincipalName,
    userDisplayName: device.userDisplayName,
    userId: device.userId,
    platform: device.operatingSystem || 'Unknown',
    osVersion: device.osVersion || 'Unknown',
    model: device.model,
    manufacturer: device.manufacturer,
    serialNumber: device.serialNumber,
    enrolledDateTime: device.enrolledDateTime || 'Unknown',
    lastSyncDateTime: device.lastSyncDateTime || 'Never',
    managementState: device.managementState || 'Unknown',
    complianceState: device.complianceState || 'Unknown',
    ownerType: device.managedDeviceOwnerType || 'Unknown',
    enrollmentType: device.enrollmentType,
    isSupervised: device.isSupervised,
    azureADDeviceId: device.azureADDeviceId,
    ethernetMacAddress: device.ethernetMacAddress,
    wiFiMacAddress: device.wiFiMacAddress
  };
}

/**
 * Check if device is considered unmanaged
 */
function isDeviceUnmanaged(device: any, options: UnmanagedDevicesReportOptions): boolean {
  const managementState = device.managementState?.toLowerCase() || '';
  const enrollmentState = device.enrollmentState?.toLowerCase() || '';
  const lastSync = device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : null;
  const enrolled = device.enrolledDateTime ? new Date(device.enrolledDateTime) : null;

  // Check for explicit failure states
  if (managementState === 'failed' || managementState === 'managedwitherrors' || managementState === 'error') {
    return true;
  }

  // Check for registration-only (not fully managed)
  if (managementState === 'registered' && enrollmentState !== 'enrolled') {
    return true;
  }

  // Check for sync issues
  if (enrolled && lastSync) {
    const daysSinceSync = Math.floor((new Date().getTime() - lastSync.getTime()) / (1000 * 60 * 60 * 24));
    if (daysSinceSync > (options.daysSinceEnrollmentThreshold || 7)) {
      return true;
    }
  }

  // Check for enrollment without sync
  if (enrolled && !lastSync) {
    const daysSinceEnrollment = Math.floor((new Date().getTime() - enrolled.getTime()) / (1000 * 60 * 60 * 24));
    if (daysSinceEnrollment > 1) {
      return true;
    }
  }

  return false;
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Device Category Reports - Configuration Manager report replication for Intune
 *
 * Provides comprehensive device categorization and enrollment failure reporting
 * that replicates Configuration Manager functionality in Intune environments.
 */
export class DeviceCategoryReports extends BaseReport {
  name = 'Device Category Reports';
  description = 'Configuration Manager-style device category and enrollment failure reports';
  category = 'ConfigMgr Compatibility';
  enabled = true;

  constructor(graphClient: Client, config: AppConfig) {
    super(graphClient, config);
  }

  /**
   * Execute the default report (device category summary)
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Device Category Reports');
    const report = await this.generateDeviceCategoryReport();

    return {
      metadata: this.createMetadata('Device Category Report', report.categorizedDevices.length),
      data: report.categorizedDevices,
      summary: report.summary
    };
  }

  // ============================================================================
  // Report 18: List of devices in a specific device category
  // ============================================================================

  /**
   * Get all device categories available in the tenant
   */
  async getAllCategories(): Promise<DeviceCategory[]> {
    logger.info('Fetching all device categories');

    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/deviceCategories')
          .select(['id', 'displayName', 'description', 'lastModifiedDateTime', 'roleScopeTagIds'])
          .top(999)
          .get();
      });

      const categories: DeviceCategory[] = await this.getAllPages(response);

      logger.info(`Found ${categories.length} device categories`);
      return categories;
    } catch (error) {
      logger.error('Error fetching device categories', error);
      throw new Error(`Failed to fetch device categories: ${error}`);
    }
  }

  /**
   * Get devices by category ID
   */
  async getDevicesByCategory(categoryId: string, options?: DeviceCategoryReportOptions): Promise<CategorizedDevice[]> {
    logger.info(`Fetching devices for category: ${categoryId}`);

    try {
      let filter = `deviceCategoryDisplayName ne null`;

      const response = await this.retryGraphCall(async () => {
        let request = this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'deviceCategoryDisplayName',
            'userPrincipalName',
            'userDisplayName',
            'userId',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'serialNumber',
            'enrolledDateTime',
            'lastSyncDateTime',
            'managementState',
            'complianceState',
            'managedDeviceOwnerType',
            'enrollmentType',
            'isSupervised',
            'azureADDeviceId',
            'ethernetMacAddress',
            'wiFiMacAddress'
          ])
          .filter(filter)
          .top(options?.limit || 999);

        return await request.get();
      });

      let devices: any[] = await this.getAllPages(response);

      // Filter by category display name (since we can't filter by category ID directly)
      const category = await this.getCategoryById(categoryId);
      if (category) {
        devices = devices.filter(d => d.deviceCategoryDisplayName === category.displayName);
      }

      // Apply additional filters
      if (options?.platformFilter && options.platformFilter.length > 0) {
        devices = devices.filter(d =>
          options.platformFilter!.includes(d.operatingSystem?.toLowerCase() || '')
        );
      }

      if (options?.complianceFilter && options.complianceFilter.length > 0) {
        devices = devices.filter(d =>
          options.complianceFilter!.includes(d.complianceState?.toLowerCase() || '')
        );
      }

      if (options?.managementStateFilter && options.managementStateFilter.length > 0) {
        devices = devices.filter(d =>
          options.managementStateFilter!.includes(d.managementState?.toLowerCase() || '')
        );
      }

      const categorizedDevices = devices.map(formatCategorizedDevice);

      logger.info(`Found ${categorizedDevices.length} devices in category`);
      return categorizedDevices;
    } catch (error) {
      logger.error('Error fetching devices by category', error);
      throw new Error(`Failed to fetch devices by category: ${error}`);
    }
  }

  /**
   * Get devices by category name
   */
  async getDevicesByCategoryName(categoryName: string, options?: DeviceCategoryReportOptions): Promise<CategorizedDevice[]> {
    logger.info(`Fetching devices for category name: ${categoryName}`);

    try {
      const filter = `deviceCategoryDisplayName eq '${categoryName}'`;

      const response = await this.retryGraphCall(async () => {
        let request = this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'deviceCategoryDisplayName',
            'userPrincipalName',
            'userDisplayName',
            'userId',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'serialNumber',
            'enrolledDateTime',
            'lastSyncDateTime',
            'managementState',
            'complianceState',
            'managedDeviceOwnerType',
            'enrollmentType',
            'isSupervised',
            'azureADDeviceId',
            'ethernetMacAddress',
            'wiFiMacAddress'
          ])
          .filter(filter)
          .top(options?.limit || 999);

        return await request.get();
      });

      let devices: any[] = await this.getAllPages(response);

      // Apply additional filters
      if (options?.platformFilter && options.platformFilter.length > 0) {
        devices = devices.filter(d =>
          options.platformFilter!.includes(d.operatingSystem?.toLowerCase() || '')
        );
      }

      if (options?.complianceFilter && options.complianceFilter.length > 0) {
        devices = devices.filter(d =>
          options.complianceFilter!.includes(d.complianceState?.toLowerCase() || '')
        );
      }

      if (options?.managementStateFilter && options.managementStateFilter.length > 0) {
        devices = devices.filter(d =>
          options.managementStateFilter!.includes(d.managementState?.toLowerCase() || '')
        );
      }

      const categorizedDevices = devices.map(formatCategorizedDevice);

      logger.info(`Found ${categorizedDevices.length} devices in category '${categoryName}'`);
      return categorizedDevices;
    } catch (error) {
      logger.error('Error fetching devices by category name', error);
      throw new Error(`Failed to fetch devices by category name: ${error}`);
    }
  }

  /**
   * Get a specific category by ID
   */
  async getCategoryById(categoryId: string): Promise<DeviceCategory | null> {
    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/deviceManagement/deviceCategories/${categoryId}`)
          .get();
      });

      return {
        id: response.id,
        displayName: response.displayName,
        description: response.description,
        lastModifiedDateTime: response.lastModifiedDateTime,
        roleScopeTagIds: response.roleScopeTagIds
      };
    } catch (error) {
      logger.error(`Error fetching category ${categoryId}`, error);
      return null;
    }
  }

  /**
   * Get uncategorized devices
   */
  async getUncategorizedDevices(options?: DeviceCategoryReportOptions): Promise<CategorizedDevice[]> {
    logger.info('Fetching uncategorized devices');

    try {
      const filter = `deviceCategoryDisplayName eq null`;

      const response = await this.retryGraphCall(async () => {
        let request = this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'deviceCategoryDisplayName',
            'userPrincipalName',
            'userDisplayName',
            'userId',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'serialNumber',
            'enrolledDateTime',
            'lastSyncDateTime',
            'managementState',
            'complianceState',
            'managedDeviceOwnerType',
            'enrollmentType',
            'isSupervised',
            'azureADDeviceId',
            'ethernetMacAddress',
            'wiFiMacAddress'
          ])
          .filter(filter)
          .top(options?.limit || 999);

        return await request.get();
      });

      let devices: any[] = await this.getAllPages(response);

      // Apply additional filters
      if (options?.platformFilter && options.platformFilter.length > 0) {
        devices = devices.filter(d =>
          options.platformFilter!.includes(d.operatingSystem?.toLowerCase() || '')
        );
      }

      const uncategorizedDevices = devices.map(formatCategorizedDevice);

      logger.info(`Found ${uncategorizedDevices.length} uncategorized devices`);
      return uncategorizedDevices;
    } catch (error) {
      logger.error('Error fetching uncategorized devices', error);
      throw new Error(`Failed to fetch uncategorized devices: ${error}`);
    }
  }

  /**
   * Generate comprehensive device category report
   */
  async generateDeviceCategoryReport(options?: DeviceCategoryReportOptions): Promise<DeviceCategoryReport> {
    logger.info('Generating device category report');

    try {
      // Fetch all data in parallel
      const [categories, allDevices] = await Promise.all([
        this.getAllCategories(),
        this.getAllManagedDevices()
      ]);

      // Separate categorized and uncategorized devices
      const categorizedDevices = allDevices
        .filter(d => d.deviceCategoryDisplayName)
        .map(formatCategorizedDevice);

      const uncategorizedDevices = allDevices
        .filter(d => !d.deviceCategoryDisplayName)
        .map(formatCategorizedDevice);

      // Calculate category counts
      const categoryDeviceCounts = new Map<string, number>();
      categorizedDevices.forEach(device => {
        const category = device.deviceCategoryDisplayName || 'Uncategorized';
        categoryDeviceCounts.set(category, (categoryDeviceCounts.get(category) || 0) + 1);
      });

      // Update categories with device counts
      const categoriesWithCounts = categories.map(cat => ({
        ...cat,
        deviceCount: categoryDeviceCounts.get(cat.displayName) || 0
      }));

      // Generate category breakdown
      const categoryBreakdown: CategoryBreakdown[] = categoriesWithCounts.map(cat => {
        const categoryDevices = categorizedDevices.filter(d => d.deviceCategoryDisplayName === cat.displayName);
        const totalDevices = categoryDevices.length;

        const platforms: Record<string, number> = {};
        const managementStates: Record<string, number> = {};
        let compliant = 0;
        let nonCompliant = 0;
        let unknown = 0;

        categoryDevices.forEach(device => {
          platforms[device.platform] = (platforms[device.platform] || 0) + 1;
          managementStates[device.managementState] = (managementStates[device.managementState] || 0) + 1;

          if (device.complianceState?.toLowerCase() === 'compliant') {
            compliant++;
          } else if (device.complianceState?.toLowerCase() === 'noncompliant') {
            nonCompliant++;
          } else {
            unknown++;
          }
        });

        return {
          categoryId: cat.id,
          categoryName: cat.displayName,
          deviceCount: totalDevices,
          percentage: allDevices.length > 0 ? (totalDevices / allDevices.length) * 100 : 0,
          platforms,
          complianceStats: { compliant, nonCompliant, unknown },
          managementStates
        };
      });

      // Get top categories
      const topCategories: CategoryCount[] = categoryBreakdown
        .sort((a, b) => b.deviceCount - a.deviceCount)
        .slice(0, 10)
        .map(cat => ({
          categoryName: cat.categoryName,
          count: cat.deviceCount,
          percentage: cat.percentage
        }));

      // Calculate platform distribution
      const platformDistribution: Record<string, number> = {};
      categorizedDevices.forEach(device => {
        platformDistribution[device.platform] = (platformDistribution[device.platform] || 0) + 1;
      });

      const summary: DeviceCategorySummary = {
        totalCategories: categories.length,
        totalCategorizedDevices: categorizedDevices.length,
        totalUncategorizedDevices: uncategorizedDevices.length,
        categoriesWithDevices: categoriesWithCounts.filter(cat => (cat.deviceCount || 0) > 0).length,
        emptyCategories: categoriesWithCounts.filter(cat => (cat.deviceCount || 0) === 0).length,
        categoryBreakdown,
        topCategories,
        platformDistribution,
        lastUpdated: new Date().toISOString()
      };

      logger.info('Device category report generated successfully');

      return {
        reportDate: new Date().toISOString(),
        summary,
        categories: categoriesWithCounts,
        categorizedDevices: options?.includeUncategorized === false ? categorizedDevices : categorizedDevices,
        uncategorizedDevices: options?.includeUncategorized === false ? [] : uncategorizedDevices
      };
    } catch (error) {
      logger.error('Error generating device category report', error);
      throw new Error(`Failed to generate device category report: ${error}`);
    }
  }

  // ============================================================================
  // Report 24: Mobile devices unmanaged (enrollment failed)
  // ============================================================================

  /**
   * Get all managed devices with full enrollment details
   */
  private async getAllManagedDevices(): Promise<any[]> {
    logger.info('Fetching all managed devices');

    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'deviceCategoryDisplayName',
            'userPrincipalName',
            'userDisplayName',
            'userId',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'serialNumber',
            'enrolledDateTime',
            'lastSyncDateTime',
            'managementState',
            'complianceState',
            'managedDeviceOwnerType',
            'enrollmentType',
            'isSupervised',
            'azureADDeviceId',
            'certificateExpirationDateTime',
            'ethernetMacAddress',
            'wiFiMacAddress'
          ])
          .top(999)
          .get();
      });

      const devices = await this.getAllPages(response);
      logger.info(`Found ${devices.length} managed devices`);
      return devices;
    } catch (error) {
      logger.error('Error fetching managed devices', error);
      throw new Error(`Failed to fetch managed devices: ${error}`);
    }
  }

  /**
   * Identify unmanaged devices (enrollment failures)
   */
  async getUnmanagedDevices(options?: UnmanagedDevicesReportOptions): Promise<UnmanagedDevice[]> {
    logger.info('Identifying unmanaged devices');

    try {
      const allDevices = await this.getAllManagedDevices();
      const unmanagedDevices: UnmanagedDevice[] = [];

      for (const device of allDevices) {
        // Check if device is unmanaged based on various criteria
        if (!isDeviceUnmanaged(device, options || {})) {
          continue;
        }

        // Apply date filters
        if (options?.startDate && device.enrolledDateTime) {
          const enrolledDate = new Date(device.enrolledDateTime);
          if (enrolledDate < options.startDate) continue;
        }

        if (options?.endDate && device.enrolledDateTime) {
          const enrolledDate = new Date(device.enrolledDateTime);
          if (enrolledDate > options.endDate) continue;
        }

        // Apply platform filter
        if (options?.platformFilter && options.platformFilter.length > 0) {
          if (!options.platformFilter.includes(device.operatingSystem?.toLowerCase() || '')) {
            continue;
          }
        }

        const failureType = determineFailureType(device);

        // Apply failure type filter
        if (options?.failureTypeFilter && options.failureTypeFilter.length > 0) {
          if (!options.failureTypeFilter.includes(failureType)) {
            continue;
          }
        }

        const daysSinceEnrollment = calculateDaysSinceEnrollment(device.enrolledDateTime);

        // Apply days threshold filter
        if (!options?.includeOldDevices && daysSinceEnrollment) {
          const threshold = options?.daysSinceEnrollmentThreshold || 30;
          if (daysSinceEnrollment > threshold) {
            continue;
          }
        }

        const certificateStatus = determineCertificateStatus(device);
        const troubleshootingSteps = getTroubleshootingSteps(failureType, device.errorCode);

        const unmanagedDevice: UnmanagedDevice = {
          deviceId: device.id,
          deviceName: device.deviceName,
          userPrincipalName: device.userPrincipalName || 'Unknown',
          userId: device.userId || 'Unknown',
          userDisplayName: device.userDisplayName,
          platform: device.operatingSystem || 'Unknown',
          osVersion: device.osVersion,
          model: device.model,
          manufacturer: device.manufacturer,
          serialNumber: device.serialNumber,
          enrollmentDate: device.enrolledDateTime,
          lastContactDate: device.lastSyncDateTime,
          failureType,
          failureReason: this.getFailureReason(device, failureType),
          errorCode: device.errorCode,
          managementState: device.managementState,
          enrollmentState: device.enrollmentState,
          certificateStatus,
          daysSinceEnrollment,
          troubleshootingSteps,
          additionalDetails: {
            complianceState: device.complianceState,
            ownerType: device.managedDeviceOwnerType,
            azureADDeviceId: device.azureADDeviceId
          }
        };

        unmanagedDevices.push(unmanagedDevice);
      }

      // Apply limit
      const limitedDevices = options?.limit
        ? unmanagedDevices.slice(0, options.limit)
        : unmanagedDevices;

      logger.info(`Identified ${limitedDevices.length} unmanaged devices`);
      return limitedDevices;
    } catch (error) {
      logger.error('Error identifying unmanaged devices', error);
      throw new Error(`Failed to identify unmanaged devices: ${error}`);
    }
  }

  /**
   * Get failure reason description
   */
  private getFailureReason(device: any, failureType: EnrollmentFailureType): string {
    const reasons: Record<EnrollmentFailureType, string> = {
      [EnrollmentFailureType.EnrollmentFailed]:
        'Device enrollment process failed. The device was unable to complete enrollment with Intune.',
      [EnrollmentFailureType.SiteAssignmentFailed]:
        'Device completed enrollment but failed to receive site assignment. No successful communication with management server.',
      [EnrollmentFailureType.ManagementStateFailed]:
        'Device is in a failed management state. Management agent is not functioning correctly.',
      [EnrollmentFailureType.CertificateIssueFailed]:
        'Certificate was issued but device management has failed. Certificate may be expired or invalid.',
      [EnrollmentFailureType.NoSuccessfulSync]:
        'Device enrolled successfully but has not synced with Intune service for an extended period.',
      [EnrollmentFailureType.RegistrationOnly]:
        'Device is registered in Azure AD but not fully managed by Intune. Only registration-level management active.',
      [EnrollmentFailureType.PolicyApplyFailed]:
        'Device enrolled but failed to apply policies correctly. Policy application errors detected.',
      [EnrollmentFailureType.Unknown]:
        'Unknown enrollment or management issue detected. Review device details for more information.'
    };

    let reason = reasons[failureType];

    // Add specific details if available
    if (device.managementState) {
      reason += ` (Management State: ${device.managementState})`;
    }

    return reason;
  }

  /**
   * Generate comprehensive unmanaged devices report
   */
  async generateUnmanagedDevicesReport(options?: UnmanagedDevicesReportOptions): Promise<UnmanagedDevicesReport> {
    logger.info('Generating unmanaged devices report');

    try {
      const unmanagedDevices = await this.getUnmanagedDevices(options);

      // Calculate summary statistics
      const failureTypeBreakdown: Record<EnrollmentFailureType, number> = {
        [EnrollmentFailureType.EnrollmentFailed]: 0,
        [EnrollmentFailureType.SiteAssignmentFailed]: 0,
        [EnrollmentFailureType.ManagementStateFailed]: 0,
        [EnrollmentFailureType.CertificateIssueFailed]: 0,
        [EnrollmentFailureType.NoSuccessfulSync]: 0,
        [EnrollmentFailureType.RegistrationOnly]: 0,
        [EnrollmentFailureType.PolicyApplyFailed]: 0,
        [EnrollmentFailureType.Unknown]: 0
      };

      const platformBreakdown: Record<string, number> = {};
      const uniqueUsers = new Set<string>();
      let certificateIssues = 0;
      let managementStateIssues = 0;
      let syncIssues = 0;
      const enrollmentDates: string[] = [];

      unmanagedDevices.forEach(device => {
        failureTypeBreakdown[device.failureType]++;
        platformBreakdown[device.platform] = (platformBreakdown[device.platform] || 0) + 1;
        uniqueUsers.add(device.userId);

        if (device.certificateStatus === CertificateStatus.Expired ||
            device.certificateStatus === CertificateStatus.NotIssued) {
          certificateIssues++;
        }

        if (device.failureType === EnrollmentFailureType.ManagementStateFailed) {
          managementStateIssues++;
        }

        if (device.failureType === EnrollmentFailureType.NoSuccessfulSync) {
          syncIssues++;
        }

        if (device.enrollmentDate) {
          enrollmentDates.push(device.enrollmentDate);
        }
      });

      const sortedDates = enrollmentDates.sort();
      const oldestFailure = sortedDates[0];
      const newestFailure = sortedDates[sortedDates.length - 1];

      const averageDays = unmanagedDevices
        .filter(d => d.daysSinceEnrollment !== undefined)
        .reduce((sum, d) => sum + (d.daysSinceEnrollment || 0), 0) /
        (unmanagedDevices.filter(d => d.daysSinceEnrollment !== undefined).length || 1);

      const summary: UnmanagedDevicesSummary = {
        totalUnmanagedDevices: unmanagedDevices.length,
        failureTypeBreakdown,
        platformBreakdown,
        affectedUsers: uniqueUsers.size,
        oldestFailure,
        newestFailure,
        averageDaysSinceEnrollment: Math.round(averageDays),
        certificateIssues,
        managementStateIssues,
        syncIssues,
        lastUpdated: new Date().toISOString()
      };

      // Generate recommendations based on findings
      const recommendations = this.generateRecommendations(summary, unmanagedDevices);

      logger.info('Unmanaged devices report generated successfully');

      return {
        reportDate: new Date().toISOString(),
        summary,
        unmanagedDevices,
        recommendations
      };
    } catch (error) {
      logger.error('Error generating unmanaged devices report', error);
      throw new Error(`Failed to generate unmanaged devices report: ${error}`);
    }
  }

  /**
   * Generate recommendations based on report findings
   */
  private generateRecommendations(summary: UnmanagedDevicesSummary, devices: UnmanagedDevice[]): string[] {
    const recommendations: string[] = [];

    if (summary.totalUnmanagedDevices === 0) {
      recommendations.push('No unmanaged devices detected. All devices are successfully enrolled and managed.');
      return recommendations;
    }

    recommendations.push(
      `${summary.totalUnmanagedDevices} devices require attention. Review and remediate enrollment failures.`
    );

    if (summary.certificateIssues > 0) {
      recommendations.push(
        `${summary.certificateIssues} devices have certificate issues. Review NDES/SCEP configuration and certificate validity.`
      );
    }

    if (summary.managementStateIssues > 0) {
      recommendations.push(
        `${summary.managementStateIssues} devices in failed management state. Check management agent status and connectivity.`
      );
    }

    if (summary.syncIssues > 0) {
      recommendations.push(
        `${summary.syncIssues} devices haven't synced recently. Verify network connectivity and Intune service endpoints.`
      );
    }

    if (summary.affectedUsers > 10) {
      recommendations.push(
        `${summary.affectedUsers} users affected by enrollment issues. Consider tenant-wide enrollment settings review.`
      );
    }

    // Platform-specific recommendations
    const topPlatform = Object.entries(summary.platformBreakdown)
      .sort((a, b) => b[1] - a[1])[0];

    if (topPlatform && topPlatform[1] > summary.totalUnmanagedDevices * 0.5) {
      recommendations.push(
        `${topPlatform[0]} platform accounts for ${topPlatform[1]} failures. Review platform-specific enrollment restrictions and requirements.`
      );
    }

    // Failure type recommendations
    const topFailureType = Object.entries(summary.failureTypeBreakdown)
      .sort((a, b) => b[1] - a[1])[0];

    if (topFailureType && topFailureType[1] > 0) {
      recommendations.push(
        `Most common failure type: ${topFailureType[0]} (${topFailureType[1]} devices). Focus remediation efforts here first.`
      );
    }

    if (summary.averageDaysSinceEnrollment > 30) {
      recommendations.push(
        `Average ${summary.averageDaysSinceEnrollment} days since enrollment attempt. Consider retiring and re-enrolling old failed devices.`
      );
    }

    recommendations.push(
      'Review the troubleshooting steps for each device and follow the recommended remediation actions.'
    );

    return recommendations;
  }

  // ============================================================================
  // Export Functions
  // ============================================================================

  /**
   * Export device category report to JSON
   */
  async exportCategoryReportToJSON(
    report: DeviceCategoryReport,
    outputDir: string = './reports'
  ): Promise<string> {
    const reportData: ReportData = {
      metadata: this.createMetadata(
        'Device Category Report',
        report.categorizedDevices.length + report.uncategorizedDevices.length
      ),
      data: [...report.categorizedDevices, ...report.uncategorizedDevices],
      summary: report.summary
    };

    return await OutputFormatter.format(reportData, 'json', outputDir);
  }

  /**
   * Export device category report to CSV
   */
  async exportCategoryReportToCSV(
    report: DeviceCategoryReport,
    outputDir: string = './reports'
  ): Promise<string> {
    const reportData: ReportData = {
      metadata: this.createMetadata(
        'Device Category Report',
        report.categorizedDevices.length + report.uncategorizedDevices.length
      ),
      data: [...report.categorizedDevices, ...report.uncategorizedDevices],
      summary: report.summary
    };

    return await OutputFormatter.format(reportData, 'csv', outputDir);
  }

  /**
   * Export device category report to HTML
   */
  async exportCategoryReportToHTML(
    report: DeviceCategoryReport,
    outputDir: string = './reports'
  ): Promise<string> {
    const reportData: ReportData = {
      metadata: this.createMetadata(
        'Device Category Report',
        report.categorizedDevices.length + report.uncategorizedDevices.length
      ),
      data: [...report.categorizedDevices, ...report.uncategorizedDevices],
      summary: report.summary
    };

    return await OutputFormatter.format(reportData, 'html', outputDir);
  }

  /**
   * Export unmanaged devices report to JSON
   */
  async exportUnmanagedReportToJSON(
    report: UnmanagedDevicesReport,
    outputDir: string = './reports'
  ): Promise<string> {
    const reportData: ReportData = {
      metadata: this.createMetadata('Unmanaged Devices Report', report.unmanagedDevices.length),
      data: report.unmanagedDevices.map(device => ({
        ...device,
        troubleshootingSteps: device.troubleshootingSteps.join(' | ')
      })),
      summary: {
        ...report.summary,
        recommendations: report.recommendations.join(' | ')
      }
    };

    return await OutputFormatter.format(reportData, 'json', outputDir);
  }

  /**
   * Export unmanaged devices report to CSV
   */
  async exportUnmanagedReportToCSV(
    report: UnmanagedDevicesReport,
    outputDir: string = './reports'
  ): Promise<string> {
    const reportData: ReportData = {
      metadata: this.createMetadata('Unmanaged Devices Report', report.unmanagedDevices.length),
      data: report.unmanagedDevices.map(device => ({
        deviceId: device.deviceId,
        deviceName: device.deviceName,
        userPrincipalName: device.userPrincipalName,
        platform: device.platform,
        failureType: device.failureType,
        failureReason: device.failureReason,
        managementState: device.managementState,
        certificateStatus: device.certificateStatus,
        daysSinceEnrollment: device.daysSinceEnrollment,
        enrollmentDate: device.enrollmentDate,
        lastContactDate: device.lastContactDate
      })),
      summary: report.summary
    };

    return await OutputFormatter.format(reportData, 'csv', outputDir);
  }

  /**
   * Export unmanaged devices report to HTML
   */
  async exportUnmanagedReportToHTML(
    report: UnmanagedDevicesReport,
    outputDir: string = './reports'
  ): Promise<string> {
    const reportData: ReportData = {
      metadata: this.createMetadata('Unmanaged Devices Report', report.unmanagedDevices.length),
      data: report.unmanagedDevices.map(device => ({
        deviceId: device.deviceId,
        deviceName: device.deviceName,
        userPrincipalName: device.userPrincipalName,
        platform: device.platform,
        failureType: device.failureType,
        failureReason: device.failureReason,
        managementState: device.managementState,
        certificateStatus: device.certificateStatus,
        daysSinceEnrollment: device.daysSinceEnrollment
      })),
      summary: {
        ...report.summary,
        recommendations: report.recommendations
      }
    };

    return await OutputFormatter.format(reportData, 'html', outputDir);
  }
}

// ============================================================================
// Usage Examples
// ============================================================================

/**
 * Example 1: List all device categories
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceCategoryReports } from './device-category-reports';
 *
 * async function listAllCategories() {
 *   const graphClient = Client.init({
 *     authProvider: (done) => {
 *       done(null, accessToken);
 *     }
 *   });
 *
 *   const config = { ... }; // Your app config
 *   const report = new DeviceCategoryReports(graphClient, config);
 *
 *   const categories = await report.getAllCategories();
 *   console.log(`Found ${categories.length} categories`);
 *   categories.forEach(cat => {
 *     console.log(`- ${cat.displayName}: ${cat.deviceCount || 0} devices`);
 *   });
 * }
 * ```
 */

/**
 * Example 2: Get devices in a specific category
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceCategoryReports } from './device-category-reports';
 *
 * async function getDevicesByCategory() {
 *   const graphClient = Client.init({
 *     authProvider: (done) => {
 *       done(null, accessToken);
 *     }
 *   });
 *
 *   const config = { ... }; // Your app config
 *   const report = new DeviceCategoryReports(graphClient, config);
 *
 *   // Get devices by category name
 *   const devices = await report.getDevicesByCategoryName('Corporate', {
 *     platformFilter: ['windows'],
 *     complianceFilter: ['compliant'],
 *     includeDetails: true
 *   });
 *
 *   console.log(`Found ${devices.length} devices in Corporate category`);
 *   devices.forEach(device => {
 *     console.log(`${device.deviceName} - ${device.platform} - ${device.complianceState}`);
 *   });
 * }
 * ```
 */

/**
 * Example 3: Generate comprehensive device category report and export
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceCategoryReports } from './device-category-reports';
 *
 * async function generateAndExportCategoryReport() {
 *   const graphClient = Client.init({
 *     authProvider: (done) => {
 *       done(null, accessToken);
 *     }
 *   });
 *
 *   const config = { ... }; // Your app config
 *   const report = new DeviceCategoryReports(graphClient, config);
 *
 *   // Generate comprehensive report
 *   const categoryReport = await report.generateDeviceCategoryReport({
 *     includeUncategorized: true,
 *     includeDetails: true
 *   });
 *
 *   console.log('Device Category Summary:');
 *   console.log(`Total Categories: ${categoryReport.summary.totalCategories}`);
 *   console.log(`Categorized Devices: ${categoryReport.summary.totalCategorizedDevices}`);
 *   console.log(`Uncategorized Devices: ${categoryReport.summary.totalUncategorizedDevices}`);
 *
 *   // Export to multiple formats
 *   const jsonPath = await report.exportCategoryReportToJSON(categoryReport);
 *   const csvPath = await report.exportCategoryReportToCSV(categoryReport);
 *   const htmlPath = await report.exportCategoryReportToHTML(categoryReport);
 *
 *   console.log(`Reports exported:`);
 *   console.log(`JSON: ${jsonPath}`);
 *   console.log(`CSV: ${csvPath}`);
 *   console.log(`HTML: ${htmlPath}`);
 * }
 * ```
 */

/**
 * Example 4: Identify and report unmanaged devices (enrollment failures)
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceCategoryReports, EnrollmentFailureType } from './device-category-reports';
 *
 * async function reportUnmanagedDevices() {
 *   const graphClient = Client.init({
 *     authProvider: (done) => {
 *       done(null, accessToken);
 *     }
 *   });
 *
 *   const config = { ... }; // Your app config
 *   const report = new DeviceCategoryReports(graphClient, config);
 *
 *   // Get unmanaged devices with filters
 *   const startDate = new Date();
 *   startDate.setDate(startDate.getDate() - 30); // Last 30 days
 *
 *   const unmanagedReport = await report.generateUnmanagedDevicesReport({
 *     startDate,
 *     platformFilter: ['windows', 'ios', 'android'],
 *     daysSinceEnrollmentThreshold: 7,
 *     includeOldDevices: false
 *   });
 *
 *   console.log('Unmanaged Devices Summary:');
 *   console.log(`Total Unmanaged: ${unmanagedReport.summary.totalUnmanagedDevices}`);
 *   console.log(`Affected Users: ${unmanagedReport.summary.affectedUsers}`);
 *   console.log(`Certificate Issues: ${unmanagedReport.summary.certificateIssues}`);
 *   console.log(`Management State Issues: ${unmanagedReport.summary.managementStateIssues}`);
 *
 *   console.log('\nRecommendations:');
 *   unmanagedReport.recommendations.forEach((rec, index) => {
 *     console.log(`${index + 1}. ${rec}`);
 *   });
 *
 *   // Export report
 *   const htmlPath = await report.exportUnmanagedReportToHTML(unmanagedReport);
 *   console.log(`\nReport exported to: ${htmlPath}`);
 * }
 * ```
 */

/**
 * Example 5: Monitor specific failure types and export detailed troubleshooting
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceCategoryReports, EnrollmentFailureType } from './device-category-reports';
 *
 * async function monitorCertificateFailures() {
 *   const graphClient = Client.init({
 *     authProvider: (done) => {
 *       done(null, accessToken);
 *     }
 *   });
 *
 *   const config = { ... }; // Your app config
 *   const report = new DeviceCategoryReports(graphClient, config);
 *
 *   // Get devices with certificate issues
 *   const devices = await report.getUnmanagedDevices({
 *     failureTypeFilter: [
 *       EnrollmentFailureType.CertificateIssueFailed,
 *       EnrollmentFailureType.ManagementStateFailed
 *     ],
 *     limit: 100
 *   });
 *
 *   console.log(`Found ${devices.length} devices with certificate or management issues\n`);
 *
 *   devices.forEach((device, index) => {
 *     console.log(`\nDevice ${index + 1}: ${device.deviceName || 'Unknown'}`);
 *     console.log(`User: ${device.userPrincipalName}`);
 *     console.log(`Platform: ${device.platform}`);
 *     console.log(`Failure Type: ${device.failureType}`);
 *     console.log(`Certificate Status: ${device.certificateStatus}`);
 *     console.log(`Days Since Enrollment: ${device.daysSinceEnrollment || 'Unknown'}`);
 *     console.log(`\nTroubleshooting Steps:`);
 *     device.troubleshootingSteps.forEach((step, stepIndex) => {
 *       console.log(`  ${stepIndex + 1}. ${step}`);
 *     });
 *   });
 * }
 * ```
 */

/**
 * Example 6: Automated monitoring and alerting for enrollment failures
 *
 * ```typescript
 * import { Client } from '@microsoft/microsoft-graph-client';
 * import { DeviceCategoryReports, EnrollmentFailureType } from './device-category-reports';
 *
 * async function monitorEnrollmentHealth() {
 *   const graphClient = Client.init({
 *     authProvider: (done) => {
 *       done(null, accessToken);
 *     }
 *   });
 *
 *   const config = { ... }; // Your app config
 *   const report = new DeviceCategoryReports(graphClient, config);
 *
 *   // Monitor recent enrollment attempts (last 24 hours)
 *   const yesterday = new Date();
 *   yesterday.setDate(yesterday.getDate() - 1);
 *
 *   const unmanagedReport = await report.generateUnmanagedDevicesReport({
 *     startDate: yesterday,
 *     includeOldDevices: false,
 *     daysSinceEnrollmentThreshold: 1
 *   });
 *
 *   // Alert if failure rate is high
 *   if (unmanagedReport.summary.totalUnmanagedDevices > 10) {
 *     console.log('⚠️ HIGH ENROLLMENT FAILURE RATE DETECTED!');
 *     console.log(`${unmanagedReport.summary.totalUnmanagedDevices} devices failed in last 24 hours`);
 *
 *     // Group failures by type for quick triage
 *     const failuresByType = unmanagedReport.summary.failureTypeBreakdown;
 *     console.log('\nFailure Breakdown:');
 *     Object.entries(failuresByType).forEach(([type, count]) => {
 *       if (count > 0) {
 *         console.log(`  ${type}: ${count} devices`);
 *       }
 *     });
 *
 *     // Export detailed report for investigation
 *     await report.exportUnmanagedReportToHTML(unmanagedReport, './alerts');
 *
 *     // Send alert notification (implement your notification logic here)
 *     // await sendAlertNotification(unmanagedReport);
 *   } else {
 *     console.log('✓ Enrollment health is normal');
 *     console.log(`${unmanagedReport.summary.totalUnmanagedDevices} failures in last 24 hours`);
 *   }
 * }
 * ```
 */

export default DeviceCategoryReports;
