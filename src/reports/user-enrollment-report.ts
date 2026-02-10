/**
 * User Enrollment Report Module for Microsoft Intune
 *
 * This module provides comprehensive user enrollment reporting functionality using the Microsoft Graph API.
 * It tracks device enrollment statistics, user enrollment status, enrollment failures, trends over time,
 * and platform-specific enrollment metrics.
 *
 * Features:
 * - Enrollment statistics and metrics
 * - User-specific enrollment status tracking
 * - Enrollment failure analysis with categorization
 * - Enrollment trends over configurable time periods
 * - Platform-specific enrollment breakdown
 * - Multiple export formats (JSON, CSV, HTML)
 * - Comprehensive error handling with retry logic
 * - Pagination support for large datasets
 *
 * @module user-enrollment-report
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { BaseReport } from './base-report';
import { ReportData, AppConfig } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Overall enrollment statistics
 */
export interface EnrollmentStatistics {
  totalDevices: number;
  totalUsers: number;
  enrolledLast24Hours: number;
  enrolledLast7Days: number;
  enrolledLast30Days: number;
  enrolledLast90Days: number;
  averageDevicesPerUser: number;
  enrollmentSuccessRate: number;
  totalEnrollmentAttempts: number;
  failedEnrollments: number;
  pendingEnrollments: number;
  lastUpdated: string;
}

/**
 * User enrollment status details
 */
export interface UserEnrollmentStatus {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  enrolledDeviceCount: number;
  devices: EnrolledDevice[];
  firstEnrollmentDate?: string;
  lastEnrollmentDate?: string;
  hasEnrollmentFailures: boolean;
  failureCount: number;
  enrollmentRestrictions?: string[];
}

/**
 * Enrolled device information
 */
export interface EnrolledDevice {
  deviceId: string;
  deviceName: string;
  platform: string;
  osVersion: string;
  model?: string;
  manufacturer?: string;
  enrollmentDate: string;
  enrollmentType: string;
  managementState: string;
  complianceState: string;
  lastSyncDateTime: string;
  isSupervised?: boolean;
  serialNumber?: string;
}

/**
 * Enrollment failure details
 */
export interface EnrollmentFailure {
  failureId: string;
  userId: string;
  userPrincipalName: string;
  userDisplayName?: string;
  deviceName?: string;
  deviceId?: string;
  platform: string;
  osVersion?: string;
  failureDateTime: string;
  errorCode: string;
  errorMessage: string;
  failureCategory: EnrollmentFailureCategory;
  failureReason: string;
  troubleshootingSteps: string[];
  correlationId?: string;
  additionalDetails?: Record<string, any>;
}

/**
 * Enrollment failure categories
 */
export enum EnrollmentFailureCategory {
  Authentication = 'Authentication',
  Authorization = 'Authorization',
  Licensing = 'Licensing',
  DeviceLimit = 'Device Limit',
  PlatformRestriction = 'Platform Restriction',
  NetworkConnectivity = 'Network Connectivity',
  CertificateIssue = 'Certificate Issue',
  PolicyConflict = 'Policy Conflict',
  DeviceConfiguration = 'Device Configuration',
  Unknown = 'Unknown'
}

/**
 * Enrollment trend data point
 */
export interface EnrollmentTrendDataPoint {
  date: string;
  totalEnrollments: number;
  successfulEnrollments: number;
  failedEnrollments: number;
  uniqueUsers: number;
  platformBreakdown: PlatformBreakdown;
}

/**
 * Platform breakdown statistics
 */
export interface PlatformBreakdown {
  iOS: number;
  Android: number;
  Windows: number;
  macOS: number;
  Linux: number;
  Other: number;
}

/**
 * Platform enrollment statistics
 */
export interface PlatformEnrollmentStats {
  platform: string;
  totalDevices: number;
  enrolledLast24Hours: number;
  enrolledLast7Days: number;
  enrolledLast30Days: number;
  enrolledLast90Days: number;
  failureCount: number;
  failureRate: number;
  averageEnrollmentTime?: number;
  topModels: DeviceModelCount[];
  topManufacturers: ManufacturerCount[];
  osVersionDistribution: OsVersionCount[];
}

/**
 * Device model count
 */
export interface DeviceModelCount {
  model: string;
  count: number;
  percentage: number;
}

/**
 * Manufacturer count
 */
export interface ManufacturerCount {
  manufacturer: string;
  count: number;
  percentage: number;
}

/**
 * OS version count
 */
export interface OsVersionCount {
  osVersion: string;
  count: number;
  percentage: number;
}

/**
 * Options for the user enrollment report
 */
export interface UserEnrollmentReportOptions {
  startDate?: Date;
  endDate?: Date;
  includeTrends?: boolean;
  includeFailures?: boolean;
  includePlatformStats?: boolean;
  includeUserDetails?: boolean;
  trendDays?: number;
  maxFailures?: number;
  platformFilter?: string[];
  userFilter?: string[];
}

/**
 * Complete enrollment report data
 */
export interface EnrollmentReportData {
  statistics: EnrollmentStatistics;
  trends?: EnrollmentTrendDataPoint[];
  failures?: EnrollmentFailure[];
  platformStats?: PlatformEnrollmentStats[];
  userEnrollments?: UserEnrollmentStatus[];
}

// ============================================================================
// Error Code Mapping and Troubleshooting
// ============================================================================

/**
 * Error code to failure category mapping with troubleshooting steps
 */
const ERROR_CODE_MAPPING: Record<string, { category: EnrollmentFailureCategory; reason: string; steps: string[] }> = {
  '80180002': {
    category: EnrollmentFailureCategory.Licensing,
    reason: 'User does not have an Intune license assigned',
    steps: [
      'Verify the user has an active Intune license in Microsoft 365 admin center',
      'Check Azure AD Premium licensing if using Conditional Access',
      'Ensure the license is properly assigned and not expired',
      'Wait up to 24 hours for license propagation across services',
      'Try re-assigning the license if issue persists'
    ]
  },
  '80180014': {
    category: EnrollmentFailureCategory.DeviceLimit,
    reason: 'User has reached the maximum device enrollment limit',
    steps: [
      'Check device enrollment restrictions in Intune portal',
      'Review the device limit settings for the user or group',
      'Identify and remove old or unused devices from the user account',
      'Verify device type restrictions are not blocking enrollment',
      'Consider increasing the device limit if appropriate'
    ]
  },
  '80180018': {
    category: EnrollmentFailureCategory.PlatformRestriction,
    reason: 'Device platform is not allowed by enrollment restrictions',
    steps: [
      'Verify the device OS version meets minimum requirements',
      'Check platform enrollment restrictions in Intune portal',
      'Review device type restrictions and allowed platforms',
      'Ensure the platform is enabled for enrollment',
      'Check if device is blocked by custom enrollment restrictions'
    ]
  },
  '0x80180026': {
    category: EnrollmentFailureCategory.Authentication,
    reason: 'Authentication failed during enrollment',
    steps: [
      'Verify user credentials are correct and account is active',
      'Check if Multi-Factor Authentication (MFA) is properly configured',
      'Ensure the user account is not locked, disabled, or expired',
      'Review Conditional Access policies that may block enrollment',
      'Clear browser/app cache and cookies, then try again',
      'Check if password has expired and needs to be reset'
    ]
  },
  '0xcaa9001f': {
    category: EnrollmentFailureCategory.CertificateIssue,
    reason: 'Certificate validation or issuance failed',
    steps: [
      'Verify NDES (Network Device Enrollment Service) configuration',
      'Check that certificate connector service is running',
      'Review certificate profiles in Intune portal',
      'Ensure device can reach certificate endpoints',
      'Verify certificate template permissions in AD CS',
      'Check firewall rules for certificate authority access'
    ]
  },
  '0x80072ee7': {
    category: EnrollmentFailureCategory.NetworkConnectivity,
    reason: 'Network connectivity issues during enrollment',
    steps: [
      'Verify internet connectivity on the device',
      'Check firewall rules allow required Intune endpoints',
      'Ensure proxy settings are configured correctly',
      'Verify required URLs are accessible and not blocked',
      'Review network security appliances for traffic blocking',
      'Test connectivity to *.manage.microsoft.com'
    ]
  },
  '0x80180012': {
    category: EnrollmentFailureCategory.Authorization,
    reason: 'User not authorized to enroll devices',
    steps: [
      'Verify user has appropriate Intune role assignments',
      'Check enrollment restrictions and group memberships',
      'Ensure user is not in an excluded group for enrollment',
      'Review Conditional Access policies affecting enrollment',
      'Verify tenant-level enrollment settings'
    ]
  },
  '0x80180013': {
    category: EnrollmentFailureCategory.PolicyConflict,
    reason: 'Conflicting policies preventing enrollment',
    steps: [
      'Review all policies assigned to the user or device',
      'Check for conflicting enrollment profiles',
      'Verify no duplicate or contradictory settings',
      'Review policy assignment priorities',
      'Check Intune policy monitoring and troubleshooting reports'
    ]
  },
  '0x87d13ba2': {
    category: EnrollmentFailureCategory.DeviceConfiguration,
    reason: 'Device configuration or management profile installation failed',
    steps: [
      'Verify device meets all prerequisites for enrollment',
      'Check device management capabilities and settings',
      'Ensure device is not already enrolled in another MDM',
      'Remove any previous MDM profiles before enrolling',
      'Verify device OS version is supported',
      'Check device storage and available memory'
    ]
  }
};

/**
 * Categorize enrollment failure by error code
 */
function categorizeEnrollmentFailure(errorCode: string): { category: EnrollmentFailureCategory; reason: string; steps: string[] } {
  const mapping = ERROR_CODE_MAPPING[errorCode];
  if (mapping) {
    return mapping;
  }

  // Default categorization for unknown errors
  return {
    category: EnrollmentFailureCategory.Unknown,
    reason: 'Unknown enrollment failure',
    steps: [
      'Review complete error message in Intune portal',
      'Check device enrollment troubleshooting events',
      'Verify all enrollment prerequisites are met',
      'Ensure device and user meet all policy requirements',
      'Try enrolling again after a few minutes',
      'Contact Microsoft support with error code and logs'
    ]
  };
}

/**
 * Normalize platform names for consistency
 */
function normalizePlatform(platform: string): string {
  const normalized = platform.toLowerCase();

  if (normalized.includes('ios') || normalized.includes('iphone') || normalized.includes('ipad')) {
    return 'iOS';
  } else if (normalized.includes('android')) {
    return 'Android';
  } else if (normalized.includes('windows')) {
    return 'Windows';
  } else if (normalized.includes('mac')) {
    return 'macOS';
  } else if (normalized.includes('linux')) {
    return 'Linux';
  }

  return 'Other';
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * User Enrollment Report Class
 *
 * Generates comprehensive enrollment reports for Microsoft Intune using Graph API.
 * Extends BaseReport to leverage common functionality like pagination and retry logic.
 *
 * @example
 * ```typescript
 * const report = new UserEnrollmentReport(graphClient, config);
 * const data = await report.execute();
 * const jsonExport = report.exportAsJSON(data);
 * const csvExport = report.exportAsCSV(data);
 * ```
 */
export class UserEnrollmentReport extends BaseReport {
  name = 'User Enrollment Report';
  description = 'Comprehensive user enrollment statistics, trends, and failure analysis';
  category = 'Enrollment';
  enabled = true;

  /**
   * Execute the enrollment report
   *
   * @param options - Report configuration options
   * @returns Complete enrollment report data
   */
  async execute(options: UserEnrollmentReportOptions = {}): Promise<ReportData> {
    logger.info('Starting User Enrollment Report generation');

    try {
      const {
        includeTrends = true,
        includeFailures = true,
        includePlatformStats = true,
        includeUserDetails = false,
        trendDays = 30,
        maxFailures = 100
      } = options;

      // Fetch all required data in parallel for better performance
      const [
        managedDevices,
        troubleshootingEvents
      ] = await Promise.all([
        this.fetchAllManagedDevices(),
        includeFailures ? this.fetchEnrollmentFailures(trendDays) : Promise.resolve([])
      ]);

      logger.info(`Fetched ${managedDevices.length} managed devices`);
      logger.info(`Fetched ${troubleshootingEvents.length} troubleshooting events`);

      // Generate statistics
      const statistics = this.calculateEnrollmentStatistics(managedDevices, troubleshootingEvents);

      // Generate report components based on options
      const reportData: EnrollmentReportData = {
        statistics
      };

      if (includeTrends) {
        reportData.trends = this.generateEnrollmentTrends(managedDevices, troubleshootingEvents, trendDays);
      }

      if (includeFailures) {
        reportData.failures = this.processEnrollmentFailures(troubleshootingEvents).slice(0, maxFailures);
      }

      if (includePlatformStats) {
        reportData.platformStats = this.generatePlatformStatistics(managedDevices);
      }

      if (includeUserDetails) {
        reportData.userEnrollments = this.generateUserEnrollmentStatus(managedDevices, troubleshootingEvents);
      }

      logger.info('User Enrollment Report generation completed successfully');

      return {
        metadata: this.createMetadata('User Enrollment Report', managedDevices.length, options),
        data: [reportData],
        summary: {
          totalDevices: statistics.totalDevices,
          totalUsers: statistics.totalUsers,
          successRate: statistics.enrollmentSuccessRate,
          failedEnrollments: statistics.failedEnrollments
        }
      };
    } catch (error) {
      logger.error('Error generating User Enrollment Report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Data Fetching Methods
  // ==========================================================================

  /**
   * Fetch all managed devices with pagination
   */
  private async fetchAllManagedDevices(): Promise<any[]> {
    logger.info('Fetching all managed devices');

    return this.retryGraphCall(async () => {
      const response = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select([
          'id',
          'deviceName',
          'managedDeviceOwnerType',
          'enrolledDateTime',
          'lastSyncDateTime',
          'operatingSystem',
          'osVersion',
          'model',
          'manufacturer',
          'serialNumber',
          'complianceState',
          'managementState',
          'managementAgent',
          'userId',
          'userPrincipalName',
          'userDisplayName',
          'enrollmentType',
          'azureADDeviceId',
          'deviceEnrollmentType',
          'isSupervised'
        ].join(','))
        .top(999)
        .get();

      return this.getAllPages<any>(response);
    });
  }

  /**
   * Fetch all enrolled users from Azure AD
   * Returns users who have Intune licenses assigned
   */
  async getAllEnrolledUsers(): Promise<any[]> {
    logger.info('Fetching all enrolled users');

    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/users')
          .select([
            'id',
            'userPrincipalName',
            'displayName',
            'department',
            'jobTitle',
            'assignedLicenses',
            'accountEnabled',
            'createdDateTime',
            'userType'
          ])
          .filter('assignedLicenses/$count ne 0')
          .header('ConsistencyLevel', 'eventual')
          .count(true)
          .top(999)
          .get();
      });

      return this.getAllPages<any>(response);
    } catch (error) {
      logger.error('Error fetching enrolled users', error);
      return [];
    }
  }

  /**
   * Fetch user information for a specific user
   */
  private async fetchUserInfo(userId: string): Promise<any> {
    return this.retryGraphCall(async () => {
      return await this.graphClient
        .api(`/users/${userId}`)
        .select(['id', 'userPrincipalName', 'displayName', 'department', 'jobTitle'])
        .get();
    });
  }

  /**
   * Fetch owned devices for a specific user
   * This includes both registered and managed devices owned by the user
   */
  async getUserOwnedDevices(userId: string): Promise<any[]> {
    logger.info(`Fetching owned devices for user: ${userId}`);

    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${userId}/ownedDevices`)
          .select([
            'id',
            'displayName',
            'deviceId',
            'operatingSystem',
            'operatingSystemVersion',
            'trustType',
            'approximateLastSignInDateTime',
            'isCompliant',
            'isManaged',
            'profileType',
            'registrationDateTime'
          ])
          .get();
      });

      return this.getAllPages<any>(response);
    } catch (error) {
      logger.warn(`Unable to fetch owned devices for user ${userId}`, error);
      return [];
    }
  }

  /**
   * Fetch devices for a specific user
   */
  async getUserDevices(userId: string): Promise<EnrolledDevice[]> {
    logger.info(`Fetching devices for user: ${userId}`);

    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${userId}/managedDevices`)
          .select([
            'id',
            'deviceName',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'enrolledDateTime',
            'enrollmentType',
            'managementState',
            'complianceState',
            'lastSyncDateTime',
            'isSupervised',
            'serialNumber'
          ].join(','))
          .get();
      });

      const devices = await this.getAllPages<any>(response);

      return devices.map(device => ({
        deviceId: device.id || '',
        deviceName: device.deviceName || 'Unknown',
        platform: normalizePlatform(device.operatingSystem || 'Unknown'),
        osVersion: device.osVersion || 'Unknown',
        model: device.model,
        manufacturer: device.manufacturer,
        enrollmentDate: device.enrolledDateTime || '',
        enrollmentType: device.enrollmentType || 'Unknown',
        managementState: device.managementState || 'Unknown',
        complianceState: device.complianceState || 'Unknown',
        lastSyncDateTime: device.lastSyncDateTime || '',
        isSupervised: device.isSupervised,
        serialNumber: device.serialNumber
      }));
    } catch (error) {
      logger.error(`Error fetching devices for user ${userId}`, error);
      return [];
    }
  }

  /**
   * Fetch managed devices by user principal name using filter
   */
  async getUserDevicesByUPN(userPrincipalName: string): Promise<EnrolledDevice[]> {
    logger.info(`Fetching managed devices for UPN: ${userPrincipalName}`);

    try {
      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/managedDevices')
          .filter(`userPrincipalName eq '${userPrincipalName}'`)
          .select([
            'id',
            'deviceName',
            'operatingSystem',
            'osVersion',
            'model',
            'manufacturer',
            'enrolledDateTime',
            'enrollmentType',
            'managementState',
            'complianceState',
            'lastSyncDateTime',
            'isSupervised',
            'serialNumber',
            'managedDeviceOwnerType'
          ].join(','))
          .top(999)
          .get();
      });

      const devices = await this.getAllPages<any>(response);

      return devices.map(device => ({
        deviceId: device.id || '',
        deviceName: device.deviceName || 'Unknown',
        platform: normalizePlatform(device.operatingSystem || 'Unknown'),
        osVersion: device.osVersion || 'Unknown',
        model: device.model,
        manufacturer: device.manufacturer,
        enrollmentDate: device.enrolledDateTime || '',
        enrollmentType: device.enrollmentType || 'Unknown',
        managementState: device.managementState || 'Unknown',
        complianceState: device.complianceState || 'Unknown',
        lastSyncDateTime: device.lastSyncDateTime || '',
        isSupervised: device.isSupervised,
        serialNumber: device.serialNumber
      }));
    } catch (error) {
      logger.error(`Error fetching devices for UPN ${userPrincipalName}`, error);
      return [];
    }
  }

  /**
   * Fetch enrollment failures from troubleshooting events
   */
  private async fetchEnrollmentFailures(days: number = 30): Promise<any[]> {
    logger.info(`Fetching enrollment failures for last ${days} days`);

    try {
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/troubleshootingEvents')
          .filter(`eventDateTime ge ${startDate.toISOString()}`)
          .top(999)
          .get();
      });

      const events = await this.getAllPages<any>(response);

      // Filter for enrollment-related events
      return events.filter(event =>
        event.eventName?.toLowerCase().includes('enrollment') ||
        event.eventName?.toLowerCase().includes('enroll')
      );
    } catch (error) {
      logger.warn('Unable to fetch troubleshooting events, may not be available', error);
      return [];
    }
  }

  /**
   * Get detailed troubleshooting events for a specific user
   */
  async getUserTroubleshootingEvents(userId: string, days: number = 30): Promise<any[]> {
    logger.info(`Fetching troubleshooting events for user: ${userId}`);

    try {
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - days);

      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/troubleshootingEvents')
          .filter(`userId eq '${userId}' and eventDateTime ge ${startDate.toISOString()}`)
          .select([
            'id',
            'eventDateTime',
            'eventName',
            'userId',
            'deviceId',
            'correlationId',
            'troubleshootingErrorDetails',
            'additionalInformation'
          ].join(','))
          .orderby('eventDateTime desc')
          .top(999)
          .get();
      });

      return this.getAllPages<any>(response);
    } catch (error) {
      logger.warn(`Unable to fetch troubleshooting events for user ${userId}`, error);
      return [];
    }
  }

  /**
   * Get enrollment troubleshooting events with enhanced details
   */
  async getEnrollmentTroubleshootingEvents(
    startDate?: Date,
    endDate?: Date,
    userId?: string
  ): Promise<any[]> {
    logger.info('Fetching enrollment troubleshooting events');

    try {
      const start = startDate || new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
      const end = endDate || new Date();

      let filterQuery = `eventDateTime ge ${start.toISOString()} and eventDateTime le ${end.toISOString()}`;

      if (userId) {
        filterQuery += ` and userId eq '${userId}'`;
      }

      const response = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/troubleshootingEvents')
          .filter(filterQuery)
          .select([
            'id',
            'eventDateTime',
            'eventName',
            'userId',
            'deviceId',
            'correlationId',
            'troubleshootingErrorDetails',
            'additionalInformation'
          ].join(','))
          .orderby('eventDateTime desc')
          .top(999)
          .get();
      });

      const events = await this.getAllPages<any>(response);

      // Filter and enhance enrollment-related events
      return events
        .filter(event =>
          event.eventName?.toLowerCase().includes('enrollment') ||
          event.eventName?.toLowerCase().includes('enroll') ||
          event.eventName?.toLowerCase().includes('mdm')
        )
        .map(event => {
          const errorCode = event.troubleshootingErrorDetails?.errorCode ||
                           event.correlationId ||
                           'Unknown';
          const { category, reason, steps } = categorizeEnrollmentFailure(errorCode);

          return {
            ...event,
            categorizedFailure: {
              category,
              reason,
              troubleshootingSteps: steps
            }
          };
        });
    } catch (error) {
      logger.warn('Unable to fetch enrollment troubleshooting events', error);
      return [];
    }
  }

  /**
   * Fetch enrollment summary report data
   */
  private async fetchEnrollmentSummary(): Promise<any> {
    try {
      return await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/reports/getDeviceEnrollmentSummary')
          .post({});
      });
    } catch (error) {
      logger.warn('Enrollment summary report not available', error);
      return null;
    }
  }

  // ==========================================================================
  // Statistics Calculation Methods
  // ==========================================================================

  /**
   * Calculate overall enrollment statistics
   */
  private calculateEnrollmentStatistics(
    devices: any[],
    failures: any[]
  ): EnrollmentStatistics {
    const now = new Date();
    const last24Hours = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    const last7Days = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const last30Days = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const last90Days = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

    const uniqueUsers = new Set(devices.map(d => d.userId).filter(Boolean));

    const enrolledLast24Hours = devices.filter(
      d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last24Hours
    ).length;

    const enrolledLast7Days = devices.filter(
      d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last7Days
    ).length;

    const enrolledLast30Days = devices.filter(
      d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last30Days
    ).length;

    const enrolledLast90Days = devices.filter(
      d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last90Days
    ).length;

    const totalAttempts = devices.length + failures.length;
    const successRate = totalAttempts > 0 ? (devices.length / totalAttempts) * 100 : 100;

    return {
      totalDevices: devices.length,
      totalUsers: uniqueUsers.size,
      enrolledLast24Hours,
      enrolledLast7Days,
      enrolledLast30Days,
      enrolledLast90Days,
      averageDevicesPerUser: uniqueUsers.size > 0 ? devices.length / uniqueUsers.size : 0,
      enrollmentSuccessRate: Math.round(successRate * 100) / 100,
      totalEnrollmentAttempts: totalAttempts,
      failedEnrollments: failures.length,
      pendingEnrollments: 0, // Would need additional API to determine
      lastUpdated: new Date().toISOString()
    };
  }

  /**
   * Get enrollment statistics for a specific date range
   */
  async getEnrollmentStatistics(startDate: Date, endDate: Date): Promise<EnrollmentStatistics> {
    logger.info(`Getting enrollment statistics from ${startDate.toISOString()} to ${endDate.toISOString()}`);

    const devices = await this.fetchAllManagedDevices();
    const failures = await this.fetchEnrollmentFailures(
      Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24))
    );

    // Filter devices by date range
    const devicesInRange = devices.filter(d => {
      if (!d.enrolledDateTime) return false;
      const enrollDate = new Date(d.enrolledDateTime);
      return enrollDate >= startDate && enrollDate <= endDate;
    });

    return this.calculateEnrollmentStatistics(devicesInRange, failures);
  }

  // ==========================================================================
  // User Enrollment Status Methods
  // ==========================================================================

  /**
   * Get enrollment status for a specific user
   */
  async getUserEnrollmentStatus(userId: string): Promise<UserEnrollmentStatus> {
    logger.info(`Getting enrollment status for user: ${userId}`);

    try {
      const [userInfo, devices, failures] = await Promise.all([
        this.fetchUserInfo(userId),
        this.getUserDevices(userId),
        this.fetchEnrollmentFailures(90)
      ]);

      const userFailures = failures.filter(f => f.userId === userId);
      const enrollmentDates = devices
        .map(d => d.enrollmentDate)
        .filter(d => d)
        .sort();

      return {
        userId: userInfo.id,
        userPrincipalName: userInfo.userPrincipalName,
        displayName: userInfo.displayName || 'Unknown',
        department: userInfo.department,
        jobTitle: userInfo.jobTitle,
        enrolledDeviceCount: devices.length,
        devices,
        firstEnrollmentDate: enrollmentDates[0],
        lastEnrollmentDate: enrollmentDates[enrollmentDates.length - 1],
        hasEnrollmentFailures: userFailures.length > 0,
        failureCount: userFailures.length
      };
    } catch (error) {
      logger.error(`Error getting enrollment status for user ${userId}`, error);
      throw error;
    }
  }

  /**
   * Get comprehensive enrollment status including owned devices
   */
  async getComprehensiveUserEnrollmentStatus(userId: string): Promise<UserEnrollmentStatus & { ownedDevices?: any[] }> {
    logger.info(`Getting comprehensive enrollment status for user: ${userId}`);

    try {
      const [userInfo, managedDevices, ownedDevices, failures] = await Promise.all([
        this.fetchUserInfo(userId),
        this.getUserDevices(userId),
        this.getUserOwnedDevices(userId),
        this.fetchEnrollmentFailures(90)
      ]);

      const userFailures = failures.filter(f => f.userId === userId);
      const enrollmentDates = managedDevices
        .map(d => d.enrollmentDate)
        .filter(d => d)
        .sort();

      return {
        userId: userInfo.id,
        userPrincipalName: userInfo.userPrincipalName,
        displayName: userInfo.displayName || 'Unknown',
        department: userInfo.department,
        jobTitle: userInfo.jobTitle,
        enrolledDeviceCount: managedDevices.length,
        devices: managedDevices,
        ownedDevices: ownedDevices,
        firstEnrollmentDate: enrollmentDates[0],
        lastEnrollmentDate: enrollmentDates[enrollmentDates.length - 1],
        hasEnrollmentFailures: userFailures.length > 0,
        failureCount: userFailures.length
      };
    } catch (error) {
      logger.error(`Error getting comprehensive enrollment status for user ${userId}`, error);
      throw error;
    }
  }

  /**
   * Generate user enrollment status for all users
   */
  private generateUserEnrollmentStatus(
    devices: any[],
    failures: any[]
  ): UserEnrollmentStatus[] {
    const userMap = new Map<string, any[]>();

    // Group devices by user
    devices.forEach(device => {
      if (device.userId) {
        if (!userMap.has(device.userId)) {
          userMap.set(device.userId, []);
        }
        userMap.get(device.userId)!.push(device);
      }
    });

    const userStatuses: UserEnrollmentStatus[] = [];

    for (const [userId, userDevices] of userMap.entries()) {
      const firstDevice = userDevices[0];
      const userFailures = failures.filter(f => f.userId === userId);

      const enrolledDevices: EnrolledDevice[] = userDevices.map(device => ({
        deviceId: device.id || '',
        deviceName: device.deviceName || 'Unknown',
        platform: normalizePlatform(device.operatingSystem || 'Unknown'),
        osVersion: device.osVersion || 'Unknown',
        model: device.model,
        manufacturer: device.manufacturer,
        enrollmentDate: device.enrolledDateTime || '',
        enrollmentType: device.enrollmentType || 'Unknown',
        managementState: device.managementState || 'Unknown',
        complianceState: device.complianceState || 'Unknown',
        lastSyncDateTime: device.lastSyncDateTime || '',
        isSupervised: device.isSupervised,
        serialNumber: device.serialNumber
      }));

      const enrollmentDates = enrolledDevices
        .map(d => d.enrollmentDate)
        .filter(d => d)
        .sort();

      userStatuses.push({
        userId,
        userPrincipalName: firstDevice.userPrincipalName || 'Unknown',
        displayName: firstDevice.userDisplayName || 'Unknown',
        enrolledDeviceCount: userDevices.length,
        devices: enrolledDevices,
        firstEnrollmentDate: enrollmentDates[0],
        lastEnrollmentDate: enrollmentDates[enrollmentDates.length - 1],
        hasEnrollmentFailures: userFailures.length > 0,
        failureCount: userFailures.length
      });
    }

    return userStatuses.sort((a, b) => b.enrolledDeviceCount - a.enrolledDeviceCount);
  }

  // ==========================================================================
  // Enrollment Methods and Status Analysis
  // ==========================================================================

  /**
   * Get breakdown of enrollment methods used across all devices
   */
  async getEnrollmentMethodsBreakdown(): Promise<Record<string, { count: number; percentage: number; devices: string[] }>> {
    logger.info('Getting enrollment methods breakdown');

    const devices = await this.fetchAllManagedDevices();
    const methodsMap = new Map<string, string[]>();

    devices.forEach(device => {
      const method = device.enrollmentType || device.deviceEnrollmentType || 'Unknown';
      if (!methodsMap.has(method)) {
        methodsMap.set(method, []);
      }
      methodsMap.get(method)!.push(device.deviceName || device.id);
    });

    const totalDevices = devices.length;
    const breakdown: Record<string, { count: number; percentage: number; devices: string[] }> = {};

    for (const [method, deviceList] of methodsMap.entries()) {
      breakdown[method] = {
        count: deviceList.length,
        percentage: Math.round((deviceList.length / totalDevices) * 10000) / 100,
        devices: deviceList.slice(0, 10) // Include first 10 devices as examples
      };
    }

    return breakdown;
  }

  /**
   * Get enrollment status summary across all managed devices
   */
  async getEnrollmentStatusSummary(): Promise<{
    totalDevices: number;
    byManagementState: Record<string, number>;
    byEnrollmentType: Record<string, number>;
    byOwnerType: Record<string, number>;
    byPlatform: Record<string, number>;
  }> {
    logger.info('Getting enrollment status summary');

    const devices = await this.fetchAllManagedDevices();

    const summary = {
      totalDevices: devices.length,
      byManagementState: {} as Record<string, number>,
      byEnrollmentType: {} as Record<string, number>,
      byOwnerType: {} as Record<string, number>,
      byPlatform: {} as Record<string, number>
    };

    devices.forEach(device => {
      // Management state
      const state = device.managementState || 'Unknown';
      summary.byManagementState[state] = (summary.byManagementState[state] || 0) + 1;

      // Enrollment type
      const enrollType = device.enrollmentType || device.deviceEnrollmentType || 'Unknown';
      summary.byEnrollmentType[enrollType] = (summary.byEnrollmentType[enrollType] || 0) + 1;

      // Owner type
      const ownerType = device.managedDeviceOwnerType || 'Unknown';
      summary.byOwnerType[ownerType] = (summary.byOwnerType[ownerType] || 0) + 1;

      // Platform
      const platform = normalizePlatform(device.operatingSystem || 'Unknown');
      summary.byPlatform[platform] = (summary.byPlatform[platform] || 0) + 1;
    });

    return summary;
  }

  // ==========================================================================
  // Enrollment Failure Methods
  // ==========================================================================

  /**
   * Get enrollment failures with categorization
   */
  async getEnrollmentFailures(days: number = 30, maxResults: number = 100): Promise<EnrollmentFailure[]> {
    logger.info(`Getting enrollment failures for last ${days} days`);

    const failures = await this.fetchEnrollmentFailures(days);
    const processed = this.processEnrollmentFailures(failures);

    return processed.slice(0, maxResults);
  }

  /**
   * Process raw troubleshooting events into structured enrollment failures
   */
  private processEnrollmentFailures(events: any[]): EnrollmentFailure[] {
    return events.map(event => {
      const errorCode = event.correlationId || event.failureDetails || 'Unknown';
      const { category, reason, steps } = categorizeEnrollmentFailure(errorCode);

      return {
        failureId: event.id || '',
        userId: event.userId || '',
        userPrincipalName: event.userPrincipalName || 'Unknown',
        userDisplayName: event.userDisplayName,
        deviceName: event.managedDeviceName,
        deviceId: event.managedDeviceIdentifier,
        platform: normalizePlatform(event.platform || 'Unknown'),
        osVersion: event.osVersion,
        failureDateTime: event.eventDateTime || new Date().toISOString(),
        errorCode,
        errorMessage: event.failureDetails || event.additionalInformation || 'No error message available',
        failureCategory: category,
        failureReason: reason,
        troubleshootingSteps: steps,
        correlationId: event.correlationId,
        additionalDetails: {
          eventName: event.eventName,
          troubleshootingErrorDetails: event.troubleshootingErrorDetails
        }
      };
    }).sort((a, b) =>
      new Date(b.failureDateTime).getTime() - new Date(a.failureDateTime).getTime()
    );
  }

  // ==========================================================================
  // Trend Analysis Methods
  // ==========================================================================

  /**
   * Generate enrollment trends over time
   */
  async getEnrollmentTrends(days: number = 30): Promise<EnrollmentTrendDataPoint[]> {
    logger.info(`Generating enrollment trends for ${days} days`);

    const devices = await this.fetchAllManagedDevices();
    const failures = await this.fetchEnrollmentFailures(days);

    return this.generateEnrollmentTrends(devices, failures, days);
  }

  /**
   * Generate enrollment trend data points
   */
  private generateEnrollmentTrends(
    devices: any[],
    failures: any[],
    days: number
  ): EnrollmentTrendDataPoint[] {
    const endDate = new Date();
    const startDate = new Date(endDate.getTime() - days * 24 * 60 * 60 * 1000);

    const trendMap = new Map<string, EnrollmentTrendDataPoint>();

    // Initialize all dates
    for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
      const dateStr = d.toISOString().split('T')[0];
      trendMap.set(dateStr, {
        date: dateStr,
        totalEnrollments: 0,
        successfulEnrollments: 0,
        failedEnrollments: 0,
        uniqueUsers: 0,
        platformBreakdown: {
          iOS: 0,
          Android: 0,
          Windows: 0,
          macOS: 0,
          Linux: 0,
          Other: 0
        }
      });
    }

    // Count successful enrollments
    const usersByDate = new Map<string, Set<string>>();

    devices.forEach(device => {
      if (device.enrolledDateTime) {
        const dateStr = new Date(device.enrolledDateTime).toISOString().split('T')[0];
        const trend = trendMap.get(dateStr);

        if (trend) {
          trend.totalEnrollments++;
          trend.successfulEnrollments++;

          const platform = normalizePlatform(device.operatingSystem || 'Other');
          if (platform in trend.platformBreakdown) {
            (trend.platformBreakdown as any)[platform]++;
          }

          // Track unique users
          if (device.userId) {
            if (!usersByDate.has(dateStr)) {
              usersByDate.set(dateStr, new Set());
            }
            usersByDate.get(dateStr)!.add(device.userId);
          }
        }
      }
    });

    // Count failed enrollments
    failures.forEach(failure => {
      if (failure.eventDateTime) {
        const dateStr = new Date(failure.eventDateTime).toISOString().split('T')[0];
        const trend = trendMap.get(dateStr);

        if (trend) {
          trend.totalEnrollments++;
          trend.failedEnrollments++;
        }
      }
    });

    // Set unique user counts
    usersByDate.forEach((users, dateStr) => {
      const trend = trendMap.get(dateStr);
      if (trend) {
        trend.uniqueUsers = users.size;
      }
    });

    return Array.from(trendMap.values()).sort((a, b) => a.date.localeCompare(b.date));
  }

  // ==========================================================================
  // Platform Statistics Methods
  // ==========================================================================

  /**
   * Get device enrollment statistics by platform
   */
  async getDeviceEnrollmentByPlatform(): Promise<PlatformEnrollmentStats[]> {
    logger.info('Getting device enrollment by platform');

    const devices = await this.fetchAllManagedDevices();
    return this.generatePlatformStatistics(devices);
  }

  /**
   * Generate platform-specific enrollment statistics
   */
  private generatePlatformStatistics(devices: any[]): PlatformEnrollmentStats[] {
    const now = new Date();
    const last24Hours = new Date(now.getTime() - 24 * 60 * 60 * 1000);
    const last7Days = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const last30Days = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const last90Days = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);

    // Group devices by platform
    const platformMap = new Map<string, any[]>();

    devices.forEach(device => {
      const platform = normalizePlatform(device.operatingSystem || 'Other');
      if (!platformMap.has(platform)) {
        platformMap.set(platform, []);
      }
      platformMap.get(platform)!.push(device);
    });

    const stats: PlatformEnrollmentStats[] = [];

    for (const [platform, platformDevices] of platformMap.entries()) {
      // Count enrollments by time period
      const enrolledLast24Hours = platformDevices.filter(
        d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last24Hours
      ).length;

      const enrolledLast7Days = platformDevices.filter(
        d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last7Days
      ).length;

      const enrolledLast30Days = platformDevices.filter(
        d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last30Days
      ).length;

      const enrolledLast90Days = platformDevices.filter(
        d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= last90Days
      ).length;

      // Count models
      const modelCounts = new Map<string, number>();
      platformDevices.forEach(device => {
        const model = device.model || 'Unknown';
        modelCounts.set(model, (modelCounts.get(model) || 0) + 1);
      });

      const topModels: DeviceModelCount[] = Array.from(modelCounts.entries())
        .map(([model, count]) => ({
          model,
          count,
          percentage: Math.round((count / platformDevices.length) * 10000) / 100
        }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 5);

      // Count manufacturers
      const manufacturerCounts = new Map<string, number>();
      platformDevices.forEach(device => {
        const manufacturer = device.manufacturer || 'Unknown';
        manufacturerCounts.set(manufacturer, (manufacturerCounts.get(manufacturer) || 0) + 1);
      });

      const topManufacturers: ManufacturerCount[] = Array.from(manufacturerCounts.entries())
        .map(([manufacturer, count]) => ({
          manufacturer,
          count,
          percentage: Math.round((count / platformDevices.length) * 10000) / 100
        }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 5);

      // Count OS versions
      const osVersionCounts = new Map<string, number>();
      platformDevices.forEach(device => {
        const osVersion = device.osVersion || 'Unknown';
        osVersionCounts.set(osVersion, (osVersionCounts.get(osVersion) || 0) + 1);
      });

      const osVersionDistribution: OsVersionCount[] = Array.from(osVersionCounts.entries())
        .map(([osVersion, count]) => ({
          osVersion,
          count,
          percentage: Math.round((count / platformDevices.length) * 10000) / 100
        }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 10);

      stats.push({
        platform,
        totalDevices: platformDevices.length,
        enrolledLast24Hours,
        enrolledLast7Days,
        enrolledLast30Days,
        enrolledLast90Days,
        failureCount: 0, // Would need failure data
        failureRate: 0, // Would need failure data
        topModels,
        topManufacturers,
        osVersionDistribution
      });
    }

    return stats.sort((a, b) => b.totalDevices - a.totalDevices);
  }

  // ==========================================================================
  // Device Ownership Tracking Methods
  // ==========================================================================

  /**
   * Get device ownership breakdown (Corporate vs Personal)
   */
  async getDeviceOwnershipBreakdown(): Promise<{
    totalDevices: number;
    corporate: { count: number; percentage: number; devices: any[] };
    personal: { count: number; percentage: number; devices: any[] };
    unknown: { count: number; percentage: number; devices: any[] };
    byPlatform: Record<string, { corporate: number; personal: number; unknown: number }>;
  }> {
    logger.info('Getting device ownership breakdown');

    const devices = await this.fetchAllManagedDevices();

    const corporate: any[] = [];
    const personal: any[] = [];
    const unknown: any[] = [];
    const platformBreakdown = new Map<string, { corporate: number; personal: number; unknown: number }>();

    devices.forEach(device => {
      const ownerType = device.managedDeviceOwnerType?.toLowerCase() || 'unknown';
      const platform = normalizePlatform(device.operatingSystem || 'Unknown');

      // Initialize platform if not exists
      if (!platformBreakdown.has(platform)) {
        platformBreakdown.set(platform, { corporate: 0, personal: 0, unknown: 0 });
      }

      const deviceInfo = {
        id: device.id,
        name: device.deviceName,
        platform,
        user: device.userPrincipalName,
        enrolledDate: device.enrolledDateTime
      };

      if (ownerType.includes('company') || ownerType.includes('corporate')) {
        corporate.push(deviceInfo);
        platformBreakdown.get(platform)!.corporate++;
      } else if (ownerType.includes('personal')) {
        personal.push(deviceInfo);
        platformBreakdown.get(platform)!.personal++;
      } else {
        unknown.push(deviceInfo);
        platformBreakdown.get(platform)!.unknown++;
      }
    });

    const totalDevices = devices.length;

    return {
      totalDevices,
      corporate: {
        count: corporate.length,
        percentage: Math.round((corporate.length / totalDevices) * 10000) / 100,
        devices: corporate
      },
      personal: {
        count: personal.length,
        percentage: Math.round((personal.length / totalDevices) * 10000) / 100,
        devices: personal
      },
      unknown: {
        count: unknown.length,
        percentage: Math.round((unknown.length / totalDevices) * 10000) / 100,
        devices: unknown
      },
      byPlatform: Object.fromEntries(platformBreakdown)
    };
  }

  /**
   * Get user device ownership summary for a specific user
   */
  async getUserDeviceOwnership(userId: string): Promise<{
    userId: string;
    totalDevices: number;
    corporateDevices: EnrolledDevice[];
    personalDevices: EnrolledDevice[];
    ownedDevices?: any[];
  }> {
    logger.info(`Getting device ownership for user: ${userId}`);

    try {
      const [managedDevices, ownedDevices] = await Promise.all([
        this.getUserDevices(userId),
        this.getUserOwnedDevices(userId)
      ]);

      const corporateDevices = managedDevices.filter(d =>
        d.enrollmentType?.toLowerCase().includes('company') ||
        d.enrollmentType?.toLowerCase().includes('corporate')
      );

      const personalDevices = managedDevices.filter(d =>
        d.enrollmentType?.toLowerCase().includes('personal') ||
        d.enrollmentType?.toLowerCase().includes('byod')
      );

      return {
        userId,
        totalDevices: managedDevices.length,
        corporateDevices,
        personalDevices,
        ownedDevices
      };
    } catch (error) {
      logger.error(`Error getting device ownership for user ${userId}`, error);
      throw error;
    }
  }

  // ==========================================================================
  // Export Methods
  // ==========================================================================

  /**
   * Export report data as JSON
   */
  exportAsJSON(reportData: ReportData): string {
    return JSON.stringify(reportData, null, 2);
  }

  /**
   * Export report data as CSV
   */
  exportAsCSV(reportData: ReportData): string {
    const lines: string[] = [];
    const data = reportData.data[0] as EnrollmentReportData;

    // Header
    lines.push('User Enrollment Report');
    lines.push(`Generated: ${reportData.metadata.generatedAt}`);
    lines.push('');

    // Statistics
    lines.push('Enrollment Statistics');
    lines.push('Metric,Value');
    lines.push(`Total Devices,${data.statistics.totalDevices}`);
    lines.push(`Total Users,${data.statistics.totalUsers}`);
    lines.push(`Enrolled Last 24 Hours,${data.statistics.enrolledLast24Hours}`);
    lines.push(`Enrolled Last 7 Days,${data.statistics.enrolledLast7Days}`);
    lines.push(`Enrolled Last 30 Days,${data.statistics.enrolledLast30Days}`);
    lines.push(`Enrolled Last 90 Days,${data.statistics.enrolledLast90Days}`);
    lines.push(`Average Devices Per User,${data.statistics.averageDevicesPerUser.toFixed(2)}`);
    lines.push(`Success Rate,${data.statistics.enrollmentSuccessRate}%`);
    lines.push(`Total Attempts,${data.statistics.totalEnrollmentAttempts}`);
    lines.push(`Failed Enrollments,${data.statistics.failedEnrollments}`);
    lines.push('');

    // Platform Statistics
    if (data.platformStats) {
      lines.push('Platform Statistics');
      lines.push('Platform,Total,Last 24h,Last 7d,Last 30d,Last 90d,Top Model');
      data.platformStats.forEach(stat => {
        const topModel = stat.topModels[0]?.model || 'N/A';
        lines.push(
          `${stat.platform},${stat.totalDevices},${stat.enrolledLast24Hours},` +
          `${stat.enrolledLast7Days},${stat.enrolledLast30Days},${stat.enrolledLast90Days},${topModel}`
        );
      });
      lines.push('');
    }

    // Failures
    if (data.failures && data.failures.length > 0) {
      lines.push('Recent Enrollment Failures');
      lines.push('Date,User,Platform,Error Code,Category,Reason');
      data.failures.forEach(failure => {
        lines.push(
          `${failure.failureDateTime},${failure.userPrincipalName},${failure.platform},` +
          `${failure.errorCode},${failure.failureCategory},"${failure.failureReason}"`
        );
      });
    }

    return lines.join('\n');
  }

  /**
   * Export report data as HTML
   */
  exportAsHTML(reportData: ReportData): string {
    const data = reportData.data[0] as EnrollmentReportData;

    let html = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>User Enrollment Report</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 20px;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 1200px;
      margin: 0 auto;
      background: white;
      padding: 30px;
      border-radius: 8px;
      box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    h1 {
      color: #0078d4;
      border-bottom: 3px solid #0078d4;
      padding-bottom: 10px;
    }
    h2 {
      color: #333;
      margin-top: 30px;
      border-bottom: 2px solid #e0e0e0;
      padding-bottom: 8px;
    }
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
      gap: 20px;
      margin: 20px 0;
    }
    .stat-card {
      background: #f9f9f9;
      padding: 20px;
      border-radius: 6px;
      border-left: 4px solid #0078d4;
    }
    .stat-label {
      font-size: 14px;
      color: #666;
      margin-bottom: 5px;
    }
    .stat-value {
      font-size: 28px;
      font-weight: bold;
      color: #333;
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
      font-weight: 600;
    }
    td {
      padding: 10px 12px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background-color: #f5f5f5;
    }
    .success { color: #107c10; font-weight: bold; }
    .failure { color: #d13438; font-weight: bold; }
    .timestamp {
      color: #666;
      font-size: 14px;
      margin-top: 20px;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>User Enrollment Report</h1>
    <p class="timestamp">Generated: ${reportData.metadata.generatedAt}</p>

    <h2>Overview Statistics</h2>
    <div class="stats-grid">
      <div class="stat-card">
        <div class="stat-label">Total Devices</div>
        <div class="stat-value">${data.statistics.totalDevices}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Total Users</div>
        <div class="stat-value">${data.statistics.totalUsers}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Success Rate</div>
        <div class="stat-value success">${data.statistics.enrollmentSuccessRate}%</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Failed Enrollments</div>
        <div class="stat-value failure">${data.statistics.failedEnrollments}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Enrolled Last 7 Days</div>
        <div class="stat-value">${data.statistics.enrolledLast7Days}</div>
      </div>
      <div class="stat-card">
        <div class="stat-label">Avg Devices Per User</div>
        <div class="stat-value">${data.statistics.averageDevicesPerUser.toFixed(2)}</div>
      </div>
    </div>
`;

    // Platform Statistics
    if (data.platformStats && data.platformStats.length > 0) {
      html += `
    <h2>Platform Statistics</h2>
    <table>
      <thead>
        <tr>
          <th>Platform</th>
          <th>Total Devices</th>
          <th>Last 24h</th>
          <th>Last 7d</th>
          <th>Last 30d</th>
          <th>Top Model</th>
        </tr>
      </thead>
      <tbody>
`;
      data.platformStats.forEach(stat => {
        const topModel = stat.topModels[0]?.model || 'N/A';
        html += `
        <tr>
          <td><strong>${stat.platform}</strong></td>
          <td>${stat.totalDevices}</td>
          <td>${stat.enrolledLast24Hours}</td>
          <td>${stat.enrolledLast7Days}</td>
          <td>${stat.enrolledLast30Days}</td>
          <td>${topModel}</td>
        </tr>
`;
      });
      html += `
      </tbody>
    </table>
`;
    }

    // Recent Failures
    if (data.failures && data.failures.length > 0) {
      html += `
    <h2>Recent Enrollment Failures (Top ${Math.min(10, data.failures.length)})</h2>
    <table>
      <thead>
        <tr>
          <th>Date</th>
          <th>User</th>
          <th>Platform</th>
          <th>Category</th>
          <th>Reason</th>
        </tr>
      </thead>
      <tbody>
`;
      data.failures.slice(0, 10).forEach(failure => {
        const date = new Date(failure.failureDateTime).toLocaleString();
        html += `
        <tr>
          <td>${date}</td>
          <td>${failure.userPrincipalName}</td>
          <td>${failure.platform}</td>
          <td><span class="failure">${failure.failureCategory}</span></td>
          <td>${failure.failureReason}</td>
        </tr>
`;
      });
      html += `
      </tbody>
    </table>
`;
    }

    html += `
  </div>
</body>
</html>
`;

    return html;
  }
}

// ============================================================================
// Export
// ============================================================================

export default UserEnrollmentReport;
