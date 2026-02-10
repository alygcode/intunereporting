/**
 * Device Category and Enrollment Reports Module for Configuration Manager
 *
 * This module provides comprehensive device categorization and enrollment failure
 * tracking reports that replicate Configuration Manager functionality in Intune.
 *
 * Configuration Manager Report Equivalents:
 * - Report 18: List of devices in a specific device category
 * - Report 24: Mobile devices unmanaged (enrollment failed)
 *
 * Features:
 * - Device category management and analysis
 * - Enrollment failure detection and categorization
 * - Certificate and management state monitoring
 * - Comprehensive troubleshooting recommendations
 * - Multiple export formats (JSON, CSV, HTML)
 * - Advanced filtering and analytics
 *
 * @module DeviceCategoryEnrollmentReports
 * @category ConfigMgr Reports
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
 * Device with category information
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
  imei?: string;
  enrolledDateTime: string;
  lastSyncDateTime: string;
  managementState: string;
  enrollmentState?: string;
  complianceState: string;
  ownerType: string;
  enrollmentType?: string;
  isSupervised?: boolean;
  azureADDeviceId?: string;
  certificateExpirationDateTime?: string;
}

/**
 * Category summary statistics
 */
export interface CategorySummary {
  categoryId: string;
  categoryName: string;
  totalDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  platformBreakdown: Record<string, number>;
  ownerTypeBreakdown: Record<string, number>;
  averageEnrollmentAge: number;
  lastEnrollment?: string;
  firstEnrollment?: string;
}

/**
 * Enrollment state information
 */
export enum EnrollmentState {
  Enrolled = 'Enrolled',
  PendingReset = 'PendingReset',
  Failed = 'Failed',
  NotContacted = 'NotContacted',
  Blocked = 'Blocked',
  Unknown = 'Unknown'
}

/**
 * Management state information
 */
export enum ManagementState {
  Managed = 'Managed',
  RetirePending = 'RetirePending',
  RetireFailed = 'RetireFailed',
  WipePending = 'WipePending',
  WipeFailed = 'WipeFailed',
  Unhealthy = 'Unhealthy',
  DeletePending = 'DeletePending',
  RetireIssued = 'RetireIssued',
  WipeIssued = 'WipeIssued',
  WipeCanceled = 'WipeCanceled',
  RetireCanceled = 'RetireCanceled',
  Discovered = 'Discovered',
  ManagedWithErrors = 'ManagedWithErrors',
  Unknown = 'Unknown'
}

/**
 * Enrollment failure types
 */
export enum EnrollmentFailureType {
  CertificateFailure = 'Certificate Failure',
  NetworkFailure = 'Network Failure',
  PolicyFailure = 'Policy Failure',
  AuthenticationFailure = 'Authentication Failure',
  AssignmentFailure = 'Assignment Failure',
  SyncFailure = 'Sync Failure',
  ManagementStateError = 'Management State Error',
  EnrollmentBlocked = 'Enrollment Blocked',
  Unknown = 'Unknown'
}

/**
 * Certificate status
 */
export enum CertificateStatus {
  Valid = 'Valid',
  Expired = 'Expired',
  Expiring = 'Expiring Soon',
  NotIssued = 'Not Issued',
  Invalid = 'Invalid',
  Unknown = 'Unknown'
}

/**
 * Unmanaged device with enrollment failure
 */
export interface UnmanagedDevice {
  deviceId?: string;
  deviceName?: string;
  userPrincipalName: string;
  userId?: string;
  userDisplayName?: string;
  platform: string;
  osVersion?: string;
  model?: string;
  manufacturer?: string;
  serialNumber?: string;
  imei?: string;
  enrollmentDate?: string;
  lastContactDate?: string;
  managementState: string;
  enrollmentState?: string;
  failureType: EnrollmentFailureType;
  failureCategory: string;
  failureReason: string;
  errorCode?: string;
  certificateStatus: CertificateStatus;
  certificateExpirationDate?: string;
  daysSinceEnrollment?: number;
  daysSinceLastContact?: number;
  troubleshootingSteps: string[];
  recommendation: string;
  severity: 'Critical' | 'High' | 'Medium' | 'Low';
  additionalDetails: Record<string, any>;
}

/**
 * Enrollment failure analysis
 */
export interface EnrollmentFailureAnalysis {
  totalFailures: number;
  failuresByType: Record<EnrollmentFailureType, number>;
  failuresByCategory: Record<string, number>;
  failuresBySeverity: Record<string, number>;
  failuresByPlatform: Record<string, number>;
  certificateIssues: number;
  networkIssues: number;
  policyIssues: number;
  authenticationIssues: number;
  affectedUsers: number;
  affectedDepartments: string[];
  oldestFailure?: string;
  newestFailure?: string;
  averageDaysSinceFailure: number;
  recommendedActions: string[];
  criticalDevices: UnmanagedDevice[];
}

/**
 * Device assignment failure
 */
export interface DeviceAssignmentFailure {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  platform: string;
  assignmentType: string;
  failureReason: string;
  expectedAssignment?: string;
  currentAssignment?: string;
  lastAttemptDate?: string;
  retryCount: number;
  troubleshootingSteps: string[];
}

/**
 * Filter options for category reports
 */
export interface CategoryReportFilter {
  categoryId?: string;
  categoryName?: string;
  platform?: string[];
  ownerType?: 'Corporate' | 'Personal' | 'All';
  complianceState?: 'Compliant' | 'NonCompliant' | 'All';
  includeUncategorized?: boolean;
  enrolledAfter?: Date;
  enrolledBefore?: Date;
}

/**
 * Filter options for enrollment failure reports
 */
export interface EnrollmentFailureFilter {
  failureType?: EnrollmentFailureType[];
  platform?: string[];
  severity?: Array<'Critical' | 'High' | 'Medium' | 'Low'>;
  certificateStatus?: CertificateStatus[];
  minDaysSinceEnrollment?: number;
  maxDaysSinceEnrollment?: number;
  includeResolved?: boolean;
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Determine enrollment failure type from device state
 */
function determineEnrollmentFailureType(device: any): EnrollmentFailureType {
  const managementState = device.managementState?.toLowerCase() || '';
  const enrollmentState = device.enrollmentState?.toLowerCase() || '';
  const lastSync = device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : null;
  const enrolled = device.enrolledDateTime ? new Date(device.enrolledDateTime) : null;
  const certExpiration = device.certificateExpirationDateTime
    ? new Date(device.certificateExpirationDateTime)
    : null;

  // Check certificate issues first
  if (certExpiration && certExpiration < new Date()) {
    return EnrollmentFailureType.CertificateFailure;
  }

  // Check for blocked enrollment
  if (enrollmentState === 'blocked' || managementState === 'blocked') {
    return EnrollmentFailureType.EnrollmentBlocked;
  }

  // Check authentication issues
  if (device.errorCode?.includes('8018') || device.errorCode?.includes('0x80180')) {
    return EnrollmentFailureType.AuthenticationFailure;
  }

  // Check policy failures
  if (managementState === 'policyapplyfailed' || device.errorCode?.includes('policy')) {
    return EnrollmentFailureType.PolicyFailure;
  }

  // Check management state errors
  if (managementState === 'unhealthy' || managementState === 'managedwitherrors') {
    return EnrollmentFailureType.ManagementStateError;
  }

  // Check sync failures
  if (enrolled && lastSync) {
    const daysSinceSync = Math.floor(
      (new Date().getTime() - lastSync.getTime()) / (1000 * 60 * 60 * 24)
    );
    if (daysSinceSync > 7) {
      return EnrollmentFailureType.SyncFailure;
    }
  }

  // Check assignment failures
  if (enrolled && !lastSync) {
    return EnrollmentFailureType.AssignmentFailure;
  }

  // Check network failures
  if (device.errorCode?.includes('network') || device.errorCode?.includes('0x80072')) {
    return EnrollmentFailureType.NetworkFailure;
  }

  return EnrollmentFailureType.Unknown;
}

/**
 * Determine failure category
 */
function determineFailureCategory(failureType: EnrollmentFailureType): string {
  const categories: Record<EnrollmentFailureType, string> = {
    [EnrollmentFailureType.CertificateFailure]: 'Infrastructure',
    [EnrollmentFailureType.NetworkFailure]: 'Connectivity',
    [EnrollmentFailureType.PolicyFailure]: 'Configuration',
    [EnrollmentFailureType.AuthenticationFailure]: 'Identity',
    [EnrollmentFailureType.AssignmentFailure]: 'Configuration',
    [EnrollmentFailureType.SyncFailure]: 'Connectivity',
    [EnrollmentFailureType.ManagementStateError]: 'Infrastructure',
    [EnrollmentFailureType.EnrollmentBlocked]: 'Policy',
    [EnrollmentFailureType.Unknown]: 'Unknown',
  };

  return categories[failureType];
}

/**
 * Determine certificate status
 */
function determineCertificateStatus(device: any): CertificateStatus {
  if (!device.certificateExpirationDateTime) {
    return CertificateStatus.NotIssued;
  }

  const expiration = new Date(device.certificateExpirationDateTime);
  const now = new Date();
  const daysUntilExpiration = Math.floor(
    (expiration.getTime() - now.getTime()) / (1000 * 60 * 60 * 24)
  );

  if (expiration < now) {
    return CertificateStatus.Expired;
  } else if (daysUntilExpiration < 30) {
    return CertificateStatus.Expiring;
  } else if (daysUntilExpiration < 0) {
    return CertificateStatus.Invalid;
  }

  return CertificateStatus.Valid;
}

/**
 * Determine failure severity
 */
function determineFailureSeverity(
  failureType: EnrollmentFailureType,
  daysSinceEnrollment?: number
): 'Critical' | 'High' | 'Medium' | 'Low' {
  // Critical: Certificate or auth failures
  if (
    failureType === EnrollmentFailureType.CertificateFailure ||
    failureType === EnrollmentFailureType.AuthenticationFailure
  ) {
    return 'Critical';
  }

  // High: Assignment failures or recent failures
  if (
    failureType === EnrollmentFailureType.AssignmentFailure ||
    (daysSinceEnrollment && daysSinceEnrollment < 7)
  ) {
    return 'High';
  }

  // Medium: Policy, network, or sync failures
  if (
    failureType === EnrollmentFailureType.PolicyFailure ||
    failureType === EnrollmentFailureType.NetworkFailure ||
    failureType === EnrollmentFailureType.SyncFailure
  ) {
    return 'Medium';
  }

  // Low: Everything else
  return 'Low';
}

/**
 * Get troubleshooting steps for failure type
 */
function getTroubleshootingSteps(
  failureType: EnrollmentFailureType,
  certificateStatus: CertificateStatus
): string[] {
  const baseSteps = [
    'Review device enrollment status in Microsoft Intune admin center',
    'Check user license assignments in Azure AD',
    'Verify enrollment restrictions are not blocking the device',
  ];

  const typeSpecificSteps: Record<EnrollmentFailureType, string[]> = {
    [EnrollmentFailureType.CertificateFailure]: [
      'Verify NDES/SCEP connector is running and accessible',
      'Check certificate template configuration in ADCS',
      'Validate certificate profile assignments in Intune',
      'Review certificate connector logs for errors',
      'Ensure device can reach certificate authority endpoints',
      'Check for expired or revoked certificates',
      'Verify certificate renewal settings',
    ],
    [EnrollmentFailureType.NetworkFailure]: [
      'Test connectivity to Intune service endpoints',
      'Verify firewall rules allow Intune URLs',
      'Check proxy configuration on device',
      'Validate DNS resolution for Intune services',
      'Review network logs for connection failures',
      'Ensure device has stable internet connection',
    ],
    [EnrollmentFailureType.PolicyFailure]: [
      'Review assigned compliance policies for conflicts',
      'Check configuration profile compatibility',
      'Validate policy assignments and scope tags',
      'Review policy application logs on device',
      'Check for policy syntax errors',
      'Verify device meets policy requirements',
    ],
    [EnrollmentFailureType.AuthenticationFailure]: [
      'Verify user credentials are valid',
      'Check Azure AD device registration status',
      'Review Conditional Access policies',
      'Validate multi-factor authentication configuration',
      'Check for expired passwords or locked accounts',
      'Review Azure AD sign-in logs',
    ],
    [EnrollmentFailureType.AssignmentFailure]: [
      'Verify enrollment profile assignments',
      'Check device group memberships',
      'Review assignment filters configuration',
      'Validate scope tags are correctly applied',
      'Check for assignment timing issues',
      'Review enrollment status page settings',
    ],
    [EnrollmentFailureType.SyncFailure]: [
      'Trigger manual sync from Company Portal',
      'Check device management agent service status',
      'Verify Intune service health',
      'Review device event logs for sync errors',
      'Check device time/date settings',
      'Ensure device is powered on and connected',
    ],
    [EnrollmentFailureType.ManagementStateError]: [
      'Check Intune Management Extension status',
      'Review device management agent logs',
      'Verify sufficient disk space on device',
      'Check for corrupt policy cache',
      'Consider retiring and re-enrolling device',
      'Review device hardware health',
    ],
    [EnrollmentFailureType.EnrollmentBlocked]: [
      'Review device enrollment restrictions',
      'Check device type restrictions',
      'Verify platform restrictions',
      'Review device limit restrictions',
      'Check for blocked manufacturers or models',
      'Validate enrollment date restrictions',
    ],
    [EnrollmentFailureType.Unknown]: [
      'Review all available error logs',
      'Check Microsoft Intune troubleshooting blade',
      'Review audit logs for related events',
      'Contact Microsoft support with error details',
    ],
  };

  const steps = [...baseSteps, ...typeSpecificSteps[failureType]];

  // Add certificate-specific steps if needed
  if (certificateStatus === CertificateStatus.Expired) {
    steps.push('Renew expired certificate immediately');
  } else if (certificateStatus === CertificateStatus.Expiring) {
    steps.push('Schedule certificate renewal before expiration');
  }

  return steps;
}

/**
 * Get recommendation based on failure type
 */
function getRecommendation(
  failureType: EnrollmentFailureType,
  severity: string
): string {
  const recommendations: Record<EnrollmentFailureType, string> = {
    [EnrollmentFailureType.CertificateFailure]:
      'Immediate action required: Review and fix certificate infrastructure. Consider implementing automated certificate renewal.',
    [EnrollmentFailureType.NetworkFailure]:
      'Verify network connectivity and firewall rules. Ensure all required Intune endpoints are accessible.',
    [EnrollmentFailureType.PolicyFailure]:
      'Review policy configuration for errors or conflicts. Simplify policies where possible.',
    [EnrollmentFailureType.AuthenticationFailure]:
      'Verify user accounts and authentication settings. Check Conditional Access policies.',
    [EnrollmentFailureType.AssignmentFailure]:
      'Review group assignments and enrollment profiles. Verify scope tags are correctly configured.',
    [EnrollmentFailureType.SyncFailure]:
      'Check device connectivity and Intune service health. Consider manual sync or device restart.',
    [EnrollmentFailureType.ManagementStateError]:
      'Investigate management agent issues. May require device re-enrollment.',
    [EnrollmentFailureType.EnrollmentBlocked]:
      'Review and adjust enrollment restrictions. Verify device meets enrollment criteria.',
    [EnrollmentFailureType.Unknown]:
      'Gather additional diagnostic information and review all available logs.',
  };

  let recommendation = recommendations[failureType];

  if (severity === 'Critical') {
    recommendation = `CRITICAL: ${recommendation}`;
  }

  return recommendation;
}

/**
 * Calculate days since enrollment
 */
function calculateDaysSince(date?: string): number | undefined {
  if (!date) return undefined;

  const dateObj = new Date(date);
  const now = new Date();
  return Math.floor((now.getTime() - dateObj.getTime()) / (1000 * 60 * 60 * 24));
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Device Category and Enrollment Reports
 *
 * Provides comprehensive device categorization and enrollment failure analysis
 * for Configuration Manager-style reporting in Intune environments.
 */
export class DeviceCategoryEnrollmentReports extends BaseReport {
  name = 'Device Category and Enrollment Reports';
  description = 'Configuration Manager-style device category and enrollment failure reports';
  category = 'ConfigMgr Compatibility';
  enabled = true;

  constructor(graphClient: Client, config: AppConfig) {
    super(graphClient, config);
  }

  /**
   * Execute the default report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Device Category and Enrollment Reports');

    const [categories, devices] = await Promise.all([
      this.getAllDeviceCategories(),
      this.getAllManagedDevices(),
    ]);

    return {
      metadata: this.createMetadata('Device Category Overview', devices.length),
      data: devices,
      summary: {
        totalCategories: categories.length,
        totalDevices: devices.length,
      },
    };
  }

  // ============================================================================
  // Report 18: Device Category Management
  // ============================================================================

  /**
   * Get all device categories
   */
  async getAllDeviceCategories(): Promise<DeviceCategory[]> {
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

      logger.info(`Retrieved ${categories.length} device categories`);
      return categories;
    } catch (error) {
      logger.error('Failed to fetch device categories', error);
      throw new Error(`Failed to fetch device categories: ${error}`);
    }
  }

  /**
   * Get devices by category name
   */
  async getDevicesByCategory(categoryName: string): Promise<CategorizedDevice[]> {
    logger.info(`Fetching devices for category: ${categoryName}`);

    try {
      const filter = `deviceCategoryDisplayName eq '${categoryName}'`;

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
            'imei',
            'enrolledDateTime',
            'lastSyncDateTime',
            'managementState',
            'complianceState',
            'managedDeviceOwnerType',
            'enrollmentType',
            'isSupervised',
            'azureADDeviceId',
            'certificateExpirationDateTime',
          ])
          .filter(filter)
          .top(999)
          .get();
      });

      const devices: any[] = await this.getAllPages(response);
      const categorizedDevices = devices.map((d) => this.formatCategorizedDevice(d));

      logger.info(`Found ${categorizedDevices.length} devices in category '${categoryName}'`);
      return categorizedDevices;
    } catch (error) {
      logger.error('Failed to fetch devices by category', error);
      throw new Error(`Failed to fetch devices by category: ${error}`);
    }
  }

  /**
   * Get devices by category ID
   */
  async getDevicesByCategoryId(categoryId: string): Promise<CategorizedDevice[]> {
    logger.info(`Fetching devices for category ID: ${categoryId}`);

    try {
      // First get the category to get its display name
      const category = await this.getCategoryById(categoryId);
      if (!category) {
        throw new Error(`Category with ID ${categoryId} not found`);
      }

      return await this.getDevicesByCategory(category.displayName);
    } catch (error) {
      logger.error('Failed to fetch devices by category ID', error);
      throw new Error(`Failed to fetch devices by category ID: ${error}`);
    }
  }

  /**
   * Get category by ID
   */
  private async getCategoryById(categoryId: string): Promise<DeviceCategory | null> {
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
        roleScopeTagIds: response.roleScopeTagIds,
      };
    } catch (error) {
      logger.error(`Failed to fetch category ${categoryId}`, error);
      return null;
    }
  }

  /**
   * Get category summary with device counts and statistics
   */
  async getCategorySummary(): Promise<CategorySummary[]> {
    logger.info('Generating category summary');

    try {
      const categories = await this.getAllDeviceCategories();
      const allDevices = await this.getAllManagedDevices();

      const summaries: CategorySummary[] = [];

      for (const category of categories) {
        const categoryDevices = allDevices.filter(
          (d) => d.deviceCategoryDisplayName === category.displayName
        );

        const compliantDevices = categoryDevices.filter(
          (d) => d.complianceState?.toLowerCase() === 'compliant'
        ).length;

        const platformBreakdown: Record<string, number> = {};
        const ownerTypeBreakdown: Record<string, number> = {};
        const enrollmentDates: Date[] = [];

        categoryDevices.forEach((device) => {
          const platform = device.operatingSystem || 'Unknown';
          const ownerType = device.managedDeviceOwnerType || 'Unknown';

          platformBreakdown[platform] = (platformBreakdown[platform] || 0) + 1;
          ownerTypeBreakdown[ownerType] = (ownerTypeBreakdown[ownerType] || 0) + 1;

          if (device.enrolledDateTime) {
            enrollmentDates.push(new Date(device.enrolledDateTime));
          }
        });

        const sortedDates = enrollmentDates.sort((a, b) => a.getTime() - b.getTime());
        const averageAge =
          enrollmentDates.length > 0
            ? enrollmentDates.reduce(
                (sum, date) =>
                  sum + (new Date().getTime() - date.getTime()) / (1000 * 60 * 60 * 24),
                0
              ) / enrollmentDates.length
            : 0;

        summaries.push({
          categoryId: category.id,
          categoryName: category.displayName,
          totalDevices: categoryDevices.length,
          compliantDevices,
          nonCompliantDevices: categoryDevices.length - compliantDevices,
          platformBreakdown,
          ownerTypeBreakdown,
          averageEnrollmentAge: Math.round(averageAge),
          lastEnrollment: sortedDates[sortedDates.length - 1]?.toISOString(),
          firstEnrollment: sortedDates[0]?.toISOString(),
        });
      }

      logger.info(`Generated summaries for ${summaries.length} categories`);
      return summaries;
    } catch (error) {
      logger.error('Failed to generate category summary', error);
      throw new Error(`Failed to generate category summary: ${error}`);
    }
  }

  // ============================================================================
  // Report 24: Enrollment Failures
  // ============================================================================

  /**
   * Get all managed devices with enrollment details
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
            'imei',
            'enrolledDateTime',
            'lastSyncDateTime',
            'managementState',
            'enrollmentState',
            'complianceState',
            'managedDeviceOwnerType',
            'enrollmentType',
            'isSupervised',
            'azureADDeviceId',
            'certificateExpirationDateTime',
          ])
          .top(999)
          .get();
      });

      const devices = await this.getAllPages(response);
      logger.info(`Retrieved ${devices.length} managed devices`);
      return devices;
    } catch (error) {
      logger.error('Failed to fetch managed devices', error);
      throw new Error(`Failed to fetch managed devices: ${error}`);
    }
  }

  /**
   * Get unmanaged devices (enrollment failed)
   */
  async getUnmanagedDevicesEnrollmentFailed(
    filter?: EnrollmentFailureFilter
  ): Promise<UnmanagedDevice[]> {
    logger.info('Identifying unmanaged devices with enrollment failures');

    try {
      const allDevices = await this.getAllManagedDevices();
      const unmanagedDevices: UnmanagedDevice[] = [];

      for (const device of allDevices) {
        if (!this.isDeviceUnmanaged(device)) {
          continue;
        }

        const failureType = determineEnrollmentFailureType(device);
        const certificateStatus = determineCertificateStatus(device);
        const daysSinceEnrollment = calculateDaysSince(device.enrolledDateTime);
        const daysSinceLastContact = calculateDaysSince(device.lastSyncDateTime);
        const severity = determineFailureSeverity(failureType, daysSinceEnrollment);

        // Apply filters
        if (filter) {
          if (filter.failureType && !filter.failureType.includes(failureType)) continue;
          if (
            filter.platform &&
            !filter.platform.includes(device.operatingSystem?.toLowerCase())
          )
            continue;
          if (filter.severity && !filter.severity.includes(severity)) continue;
          if (
            filter.certificateStatus &&
            !filter.certificateStatus.includes(certificateStatus)
          )
            continue;
          if (
            filter.minDaysSinceEnrollment &&
            daysSinceEnrollment &&
            daysSinceEnrollment < filter.minDaysSinceEnrollment
          )
            continue;
          if (
            filter.maxDaysSinceEnrollment &&
            daysSinceEnrollment &&
            daysSinceEnrollment > filter.maxDaysSinceEnrollment
          )
            continue;
        }

        const troubleshootingSteps = getTroubleshootingSteps(failureType, certificateStatus);
        const recommendation = getRecommendation(failureType, severity);

        const unmanagedDevice: UnmanagedDevice = {
          deviceId: device.id,
          deviceName: device.deviceName || 'Unknown',
          userPrincipalName: device.userPrincipalName || 'Unknown',
          userId: device.userId,
          userDisplayName: device.userDisplayName,
          platform: device.operatingSystem || 'Unknown',
          osVersion: device.osVersion,
          model: device.model,
          manufacturer: device.manufacturer,
          serialNumber: device.serialNumber,
          imei: device.imei,
          enrollmentDate: device.enrolledDateTime,
          lastContactDate: device.lastSyncDateTime,
          managementState: device.managementState || 'Unknown',
          enrollmentState: device.enrollmentState,
          failureType,
          failureCategory: determineFailureCategory(failureType),
          failureReason: this.getFailureReason(device, failureType),
          errorCode: device.errorCode,
          certificateStatus,
          certificateExpirationDate: device.certificateExpirationDateTime,
          daysSinceEnrollment,
          daysSinceLastContact,
          troubleshootingSteps,
          recommendation,
          severity,
          additionalDetails: {
            complianceState: device.complianceState,
            ownerType: device.managedDeviceOwnerType,
            azureADDeviceId: device.azureADDeviceId,
            enrollmentType: device.enrollmentType,
          },
        };

        unmanagedDevices.push(unmanagedDevice);
      }

      logger.info(`Identified ${unmanagedDevices.length} unmanaged devices`);
      return unmanagedDevices;
    } catch (error) {
      logger.error('Failed to get unmanaged devices', error);
      throw new Error(`Failed to get unmanaged devices: ${error}`);
    }
  }

  /**
   * Analyze enrollment failures
   */
  async analyzeEnrollmentFailures(
    filter?: EnrollmentFailureFilter
  ): Promise<EnrollmentFailureAnalysis> {
    logger.info('Analyzing enrollment failures');

    try {
      const unmanagedDevices = await this.getUnmanagedDevicesEnrollmentFailed(filter);

      const failuresByType: Record<EnrollmentFailureType, number> = {
        [EnrollmentFailureType.CertificateFailure]: 0,
        [EnrollmentFailureType.NetworkFailure]: 0,
        [EnrollmentFailureType.PolicyFailure]: 0,
        [EnrollmentFailureType.AuthenticationFailure]: 0,
        [EnrollmentFailureType.AssignmentFailure]: 0,
        [EnrollmentFailureType.SyncFailure]: 0,
        [EnrollmentFailureType.ManagementStateError]: 0,
        [EnrollmentFailureType.EnrollmentBlocked]: 0,
        [EnrollmentFailureType.Unknown]: 0,
      };

      const failuresByCategory: Record<string, number> = {};
      const failuresBySeverity: Record<string, number> = {
        Critical: 0,
        High: 0,
        Medium: 0,
        Low: 0,
      };
      const failuresByPlatform: Record<string, number> = {};
      const affectedUsersSet = new Set<string>();
      const affectedDepartments: string[] = [];
      const enrollmentDates: string[] = [];

      let certificateIssues = 0;
      let networkIssues = 0;
      let policyIssues = 0;
      let authenticationIssues = 0;

      unmanagedDevices.forEach((device) => {
        failuresByType[device.failureType]++;
        failuresByCategory[device.failureCategory] =
          (failuresByCategory[device.failureCategory] || 0) + 1;
        failuresBySeverity[device.severity]++;
        failuresByPlatform[device.platform] = (failuresByPlatform[device.platform] || 0) + 1;
        affectedUsersSet.add(device.userId || device.userPrincipalName);

        if (device.enrollmentDate) {
          enrollmentDates.push(device.enrollmentDate);
        }

        if (device.failureType === EnrollmentFailureType.CertificateFailure)
          certificateIssues++;
        if (device.failureType === EnrollmentFailureType.NetworkFailure) networkIssues++;
        if (device.failureType === EnrollmentFailureType.PolicyFailure) policyIssues++;
        if (device.failureType === EnrollmentFailureType.AuthenticationFailure)
          authenticationIssues++;
      });

      const sortedDates = enrollmentDates.sort();
      const averageDays =
        unmanagedDevices.filter((d) => d.daysSinceEnrollment !== undefined).length > 0
          ? unmanagedDevices
              .filter((d) => d.daysSinceEnrollment !== undefined)
              .reduce((sum, d) => sum + (d.daysSinceEnrollment || 0), 0) /
            unmanagedDevices.filter((d) => d.daysSinceEnrollment !== undefined).length
          : 0;

      const criticalDevices = unmanagedDevices.filter((d) => d.severity === 'Critical');

      const recommendedActions = this.generateRecommendedActions(
        failuresByType,
        certificateIssues,
        networkIssues,
        policyIssues,
        authenticationIssues
      );

      const analysis: EnrollmentFailureAnalysis = {
        totalFailures: unmanagedDevices.length,
        failuresByType,
        failuresByCategory,
        failuresBySeverity,
        failuresByPlatform,
        certificateIssues,
        networkIssues,
        policyIssues,
        authenticationIssues,
        affectedUsers: affectedUsersSet.size,
        affectedDepartments,
        oldestFailure: sortedDates[0],
        newestFailure: sortedDates[sortedDates.length - 1],
        averageDaysSinceFailure: Math.round(averageDays),
        recommendedActions,
        criticalDevices,
      };

      logger.info('Enrollment failure analysis complete');
      return analysis;
    } catch (error) {
      logger.error('Failed to analyze enrollment failures', error);
      throw new Error(`Failed to analyze enrollment failures: ${error}`);
    }
  }

  /**
   * Get devices with failed assignment
   */
  async getDevicesWithFailedAssignment(): Promise<DeviceAssignmentFailure[]> {
    logger.info('Identifying devices with failed assignments');

    try {
      const allDevices = await this.getAllManagedDevices();
      const assignmentFailures: DeviceAssignmentFailure[] = [];

      for (const device of allDevices) {
        const enrolled = device.enrolledDateTime ? new Date(device.enrolledDateTime) : null;
        const lastSync = device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : null;

        // Device enrolled but never synced = assignment failure
        if (enrolled && !lastSync) {
          const daysSinceEnrollment = calculateDaysSince(device.enrolledDateTime);

          if (daysSinceEnrollment && daysSinceEnrollment > 1) {
            assignmentFailures.push({
              deviceId: device.id,
              deviceName: device.deviceName || 'Unknown',
              userPrincipalName: device.userPrincipalName || 'Unknown',
              platform: device.operatingSystem || 'Unknown',
              assignmentType: 'Initial Assignment',
              failureReason:
                'Device enrolled but failed to receive initial assignment from management server',
              expectedAssignment: 'Management policies and configurations',
              currentAssignment: 'None',
              lastAttemptDate: device.enrolledDateTime,
              retryCount: daysSinceEnrollment || 0,
              troubleshootingSteps: getTroubleshootingSteps(
                EnrollmentFailureType.AssignmentFailure,
                CertificateStatus.Unknown
              ),
            });
          }
        }
      }

      logger.info(`Identified ${assignmentFailures.length} assignment failures`);
      return assignmentFailures;
    } catch (error) {
      logger.error('Failed to get devices with assignment failures', error);
      throw new Error(`Failed to get devices with assignment failures: ${error}`);
    }
  }

  // ============================================================================
  // Export Methods
  // ============================================================================

  /**
   * Export category report to JSON
   */
  async exportCategoryReportToJSON(
    categoryName: string,
    outputDir: string = './reports'
  ): Promise<string> {
    const devices = await this.getDevicesByCategory(categoryName);

    const reportData: ReportData = {
      metadata: this.createMetadata(`Devices in Category: ${categoryName}`, devices.length),
      data: devices,
      summary: {
        categoryName,
        totalDevices: devices.length,
        platforms: this.getPlatformBreakdown(devices),
      },
    };

    return await OutputFormatter.format(reportData, 'json', outputDir);
  }

  /**
   * Export category report to CSV
   */
  async exportCategoryReportToCSV(
    categoryName: string,
    outputDir: string = './reports'
  ): Promise<string> {
    const devices = await this.getDevicesByCategory(categoryName);

    const reportData: ReportData = {
      metadata: this.createMetadata(`Devices in Category: ${categoryName}`, devices.length),
      data: devices,
      summary: {
        categoryName,
        totalDevices: devices.length,
      },
    };

    return await OutputFormatter.format(reportData, 'csv', outputDir);
  }

  /**
   * Export category report to HTML
   */
  async exportCategoryReportToHTML(
    categoryName: string,
    outputDir: string = './reports'
  ): Promise<string> {
    const devices = await this.getDevicesByCategory(categoryName);

    const reportData: ReportData = {
      metadata: this.createMetadata(`Devices in Category: ${categoryName}`, devices.length),
      data: devices,
      summary: {
        categoryName,
        totalDevices: devices.length,
        platforms: this.getPlatformBreakdown(devices),
      },
    };

    return await OutputFormatter.format(reportData, 'html', outputDir);
  }

  /**
   * Export enrollment failure report to JSON
   */
  async exportEnrollmentFailureReportToJSON(
    filter?: EnrollmentFailureFilter,
    outputDir: string = './reports'
  ): Promise<string> {
    const analysis = await this.analyzeEnrollmentFailures(filter);

    const reportData: ReportData = {
      metadata: this.createMetadata('Enrollment Failure Analysis', analysis.totalFailures),
      data: [analysis],
      summary: {
        totalFailures: analysis.totalFailures,
        criticalCount: analysis.failuresBySeverity.Critical,
        affectedUsers: analysis.affectedUsers,
      },
    };

    return await OutputFormatter.format(reportData, 'json', outputDir);
  }

  /**
   * Export enrollment failure report to CSV
   */
  async exportEnrollmentFailureReportToCSV(
    filter?: EnrollmentFailureFilter,
    outputDir: string = './reports'
  ): Promise<string> {
    const devices = await this.getUnmanagedDevicesEnrollmentFailed(filter);

    const reportData: ReportData = {
      metadata: this.createMetadata('Enrollment Failures', devices.length),
      data: devices,
      summary: {
        totalFailures: devices.length,
      },
    };

    return await OutputFormatter.format(reportData, 'csv', outputDir);
  }

  /**
   * Export enrollment failure report to HTML
   */
  async exportEnrollmentFailureReportToHTML(
    filter?: EnrollmentFailureFilter,
    outputDir: string = './reports'
  ): Promise<string> {
    const analysis = await this.analyzeEnrollmentFailures(filter);

    const reportData: ReportData = {
      metadata: this.createMetadata('Enrollment Failure Analysis', analysis.totalFailures),
      data: [analysis],
      summary: {
        totalFailures: analysis.totalFailures,
        criticalCount: analysis.failuresBySeverity.Critical,
        affectedUsers: analysis.affectedUsers,
      },
    };

    return await OutputFormatter.format(reportData, 'html', outputDir);
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Format device to categorized device
   */
  private formatCategorizedDevice(device: any): CategorizedDevice {
    return {
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      deviceCategoryId: device.deviceCategoryId,
      deviceCategoryDisplayName: device.deviceCategoryDisplayName || 'Uncategorized',
      userPrincipalName: device.userPrincipalName,
      userDisplayName: device.userDisplayName,
      userId: device.userId,
      platform: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      model: device.model,
      manufacturer: device.manufacturer,
      serialNumber: device.serialNumber,
      imei: device.imei,
      enrolledDateTime: device.enrolledDateTime || 'Unknown',
      lastSyncDateTime: device.lastSyncDateTime || 'Never',
      managementState: device.managementState || 'Unknown',
      enrollmentState: device.enrollmentState,
      complianceState: device.complianceState || 'Unknown',
      ownerType: device.managedDeviceOwnerType || 'Unknown',
      enrollmentType: device.enrollmentType,
      isSupervised: device.isSupervised,
      azureADDeviceId: device.azureADDeviceId,
      certificateExpirationDateTime: device.certificateExpirationDateTime,
    };
  }

  /**
   * Check if device is unmanaged
   */
  private isDeviceUnmanaged(device: any): boolean {
    const managementState = device.managementState?.toLowerCase() || '';
    const lastSync = device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : null;
    const enrolled = device.enrolledDateTime ? new Date(device.enrolledDateTime) : null;

    // Check for explicit failure states
    if (
      managementState === 'unhealthy' ||
      managementState === 'managedwitherrors' ||
      managementState === 'wiped' ||
      managementState === 'retired'
    ) {
      return true;
    }

    // Check for certificate expiration
    if (device.certificateExpirationDateTime) {
      const certExpiry = new Date(device.certificateExpirationDateTime);
      if (certExpiry < new Date()) {
        return true;
      }
    }

    // Check for no sync after enrollment
    if (enrolled && !lastSync) {
      const daysSinceEnrollment = Math.floor(
        (new Date().getTime() - enrolled.getTime()) / (1000 * 60 * 60 * 24)
      );
      if (daysSinceEnrollment > 1) {
        return true;
      }
    }

    // Check for stale sync
    if (enrolled && lastSync) {
      const daysSinceSync = Math.floor(
        (new Date().getTime() - lastSync.getTime()) / (1000 * 60 * 60 * 24)
      );
      if (daysSinceSync > 7) {
        return true;
      }
    }

    return false;
  }

  /**
   * Get failure reason description
   */
  private getFailureReason(device: any, failureType: EnrollmentFailureType): string {
    const reasons: Record<EnrollmentFailureType, string> = {
      [EnrollmentFailureType.CertificateFailure]:
        'Device certificate has expired or is invalid, preventing secure communication with management server',
      [EnrollmentFailureType.NetworkFailure]:
        'Device unable to establish or maintain network connection to Intune service endpoints',
      [EnrollmentFailureType.PolicyFailure]:
        'Device failed to receive or apply assigned policies and configurations',
      [EnrollmentFailureType.AuthenticationFailure]:
        'Device authentication failed due to invalid credentials or expired tokens',
      [EnrollmentFailureType.AssignmentFailure]:
        'Device enrolled but failed to receive initial assignment from management server',
      [EnrollmentFailureType.SyncFailure]:
        'Device has not synchronized with Intune service for extended period',
      [EnrollmentFailureType.ManagementStateError]:
        'Device management agent is in unhealthy or error state',
      [EnrollmentFailureType.EnrollmentBlocked]:
        'Device enrollment is blocked by restrictions or compliance requirements',
      [EnrollmentFailureType.Unknown]:
        'Unknown enrollment or management issue detected',
    };

    let reason = reasons[failureType];

    if (device.managementState) {
      reason += ` (Management State: ${device.managementState})`;
    }

    if (device.errorCode) {
      reason += ` (Error Code: ${device.errorCode})`;
    }

    return reason;
  }

  /**
   * Get platform breakdown
   */
  private getPlatformBreakdown(devices: CategorizedDevice[]): Record<string, number> {
    const breakdown: Record<string, number> = {};

    devices.forEach((device) => {
      const platform = device.platform || 'Unknown';
      breakdown[platform] = (breakdown[platform] || 0) + 1;
    });

    return breakdown;
  }

  /**
   * Generate recommended actions
   */
  private generateRecommendedActions(
    failuresByType: Record<EnrollmentFailureType, number>,
    certificateIssues: number,
    networkIssues: number,
    policyIssues: number,
    authenticationIssues: number
  ): string[] {
    const actions: string[] = [];

    if (certificateIssues > 0) {
      actions.push(
        `URGENT: ${certificateIssues} devices have certificate issues. Review NDES/SCEP infrastructure immediately.`
      );
    }

    if (authenticationIssues > 0) {
      actions.push(
        `${authenticationIssues} devices have authentication failures. Review Azure AD and Conditional Access policies.`
      );
    }

    if (networkIssues > 0) {
      actions.push(
        `${networkIssues} devices experiencing network connectivity issues. Verify firewall rules and proxy settings.`
      );
    }

    if (policyIssues > 0) {
      actions.push(
        `${policyIssues} devices failing policy application. Review policy assignments and configurations.`
      );
    }

    // Get top failure type
    const topFailure = Object.entries(failuresByType)
      .sort(([, a], [, b]) => b - a)
      .filter(([, count]) => count > 0)[0];

    if (topFailure) {
      actions.push(
        `Primary issue: ${topFailure[0]} (${topFailure[1]} devices). Focus remediation efforts here.`
      );
    }

    actions.push('Review troubleshooting steps for each device and implement recommended fixes.');
    actions.push('Monitor enrollment health regularly to catch issues early.');

    return actions;
  }
}

// ============================================================================
// Usage Examples
// ============================================================================

/**
 * Example 1: List all device categories
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const categories = await report.getAllDeviceCategories();
 *
 * console.log(`Found ${categories.length} categories:`);
 * categories.forEach(cat => {
 *   console.log(`- ${cat.displayName} (${cat.deviceCount || 0} devices)`);
 * });
 * ```
 */

/**
 * Example 2: Get devices in a specific category
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const devices = await report.getDevicesByCategory('Corporate');
 *
 * console.log(`Found ${devices.length} devices in Corporate category`);
 * devices.forEach(device => {
 *   console.log(`${device.deviceName} - ${device.platform} - ${device.complianceState}`);
 * });
 * ```
 */

/**
 * Example 3: Get category summary with statistics
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const summaries = await report.getCategorySummary();
 *
 * summaries.forEach(summary => {
 *   console.log(`\nCategory: ${summary.categoryName}`);
 *   console.log(`Total Devices: ${summary.totalDevices}`);
 *   console.log(`Compliant: ${summary.compliantDevices}`);
 *   console.log(`Non-Compliant: ${summary.nonCompliantDevices}`);
 *   console.log(`Platform Breakdown:`, summary.platformBreakdown);
 * });
 * ```
 */

/**
 * Example 4: Identify unmanaged devices with enrollment failures
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();
 *
 * console.log(`Found ${unmanagedDevices.length} unmanaged devices`);
 *
 * unmanagedDevices.forEach(device => {
 *   console.log(`\nDevice: ${device.deviceName}`);
 *   console.log(`User: ${device.userPrincipalName}`);
 *   console.log(`Failure Type: ${device.failureType}`);
 *   console.log(`Severity: ${device.severity}`);
 *   console.log(`Certificate Status: ${device.certificateStatus}`);
 *   console.log(`Recommendation: ${device.recommendation}`);
 * });
 * ```
 */

/**
 * Example 5: Analyze enrollment failures with comprehensive report
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const analysis = await report.analyzeEnrollmentFailures({
 *   severity: ['Critical', 'High'],
 *   platform: ['ios', 'android']
 * });
 *
 * console.log('Enrollment Failure Analysis:');
 * console.log(`Total Failures: ${analysis.totalFailures}`);
 * console.log(`Certificate Issues: ${analysis.certificateIssues}`);
 * console.log(`Network Issues: ${analysis.networkIssues}`);
 * console.log(`Affected Users: ${analysis.affectedUsers}`);
 *
 * console.log('\nFailures by Type:');
 * Object.entries(analysis.failuresByType).forEach(([type, count]) => {
 *   if (count > 0) console.log(`  ${type}: ${count}`);
 * });
 *
 * console.log('\nRecommended Actions:');
 * analysis.recommendedActions.forEach((action, i) => {
 *   console.log(`${i + 1}. ${action}`);
 * });
 * ```
 */

/**
 * Example 6: Find devices with assignment failures
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const assignmentFailures = await report.getDevicesWithFailedAssignment();
 *
 * console.log(`Found ${assignmentFailures.length} devices with assignment failures`);
 *
 * assignmentFailures.forEach(device => {
 *   console.log(`\nDevice: ${device.deviceName}`);
 *   console.log(`User: ${device.userPrincipalName}`);
 *   console.log(`Failure Reason: ${device.failureReason}`);
 *   console.log(`Retry Count: ${device.retryCount}`);
 *   console.log('\nTroubleshooting Steps:');
 *   device.troubleshootingSteps.forEach((step, i) => {
 *     console.log(`  ${i + 1}. ${step}`);
 *   });
 * });
 * ```
 */

/**
 * Example 7: Export category report in multiple formats
 *
 * ```typescript
 * const report = new DeviceCategoryEnrollmentReports(graphClient, config);
 * const categoryName = 'Corporate';
 *
 * // Export to JSON
 * const jsonPath = await report.exportCategoryReportToJSON(categoryName);
 * console.log(`JSON report saved to: ${jsonPath}`);
 *
 * // Export to CSV
 * const csvPath = await report.exportCategoryReportToCSV(categoryName);
 * console.log(`CSV report saved to: ${csvPath}`);
 *
 * // Export to HTML
 * const htmlPath = await report.exportCategoryReportToHTML(categoryName);
 * console.log(`HTML report saved to: ${htmlPath}`);
 * ```
 */

export default DeviceCategoryEnrollmentReports;
