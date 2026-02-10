/**
 * Intune Enrollment Tracking Reports for Configuration Manager
 *
 * This module replicates Configuration Manager enrollment tracking reports for Microsoft Intune.
 * It provides two main reports:
 * - Report 17: List of Devices enrolled per user in Microsoft Intune (detailed)
 * - Report 30: Number of devices enrolled per user in Microsoft Intune (summary)
 *
 * Features:
 * - Comprehensive device enrollment tracking per user
 * - Enrollment count aggregation per user
 * - Device type breakdown per user
 * - Enrollment date tracking and filtering
 * - Multiple export formats (JSON, CSV, HTML)
 * - Advanced filtering capabilities
 *
 * @module intune-enrollment-tracking-reports
 */

import { Client } from '@microsoft/microsoft-graph-client';
import { BaseReport } from '../base-report';
import { ReportData, AppConfig } from '../../types';
import { Logger } from '../../core/logger';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Enrolled device details for a user
 */
export interface UserEnrolledDevice {
  deviceId: string;
  deviceName: string;
  platform: string;
  osVersion: string;
  model?: string;
  manufacturer?: string;
  enrollmentDate: string;
  enrollmentType: string;
  enrollmentMethod: string;
  managementState: string;
  complianceState: string;
  lastSyncDateTime: string;
  ownerType: string;
  serialNumber?: string;
  imei?: string;
  isSupervised?: boolean;
  deviceCategory?: string;
  enrollmentProfileName?: string;
}

/**
 * Detailed device list per user (Report 17)
 */
export interface UserDeviceList {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  email?: string;
  totalDevices: number;
  devices: UserEnrolledDevice[];
  devicesByPlatform: PlatformBreakdown;
  devicesByEnrollmentMethod: EnrollmentMethodBreakdown;
  devicesByOwnerType: OwnerTypeBreakdown;
  firstEnrollmentDate?: string;
  lastEnrollmentDate?: string;
  averageDeviceAge: number;
}

/**
 * Enrollment count summary per user (Report 30)
 */
export interface UserEnrollmentCount {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  email?: string;
  totalDeviceCount: number;
  iosDeviceCount: number;
  androidDeviceCount: number;
  windowsDeviceCount: number;
  macOsDeviceCount: number;
  linuxDeviceCount: number;
  otherDeviceCount: number;
  corporateDeviceCount: number;
  personalDeviceCount: number;
  compliantDeviceCount: number;
  nonCompliantDeviceCount: number;
  activeDeviceCount: number;
  inactiveDeviceCount: number;
  enrollmentTrend: EnrollmentTrend;
  lastEnrollmentDate?: string;
  firstEnrollmentDate?: string;
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
 * Enrollment method breakdown
 */
export interface EnrollmentMethodBreakdown {
  [method: string]: number;
}

/**
 * Owner type breakdown
 */
export interface OwnerTypeBreakdown {
  Corporate: number;
  Personal: number;
  Unknown: number;
}

/**
 * Enrollment trend data
 */
export interface EnrollmentTrend {
  last7Days: number;
  last30Days: number;
  last90Days: number;
  last365Days: number;
}

/**
 * Filter options for enrollment reports
 */
export interface EnrollmentTrackingFilters {
  userPrincipalName?: string;
  userId?: string;
  department?: string;
  startDate?: Date;
  endDate?: Date;
  platform?: string[];
  enrollmentMethod?: string[];
  ownerType?: 'Corporate' | 'Personal' | 'All';
  complianceState?: 'Compliant' | 'NonCompliant' | 'All';
  minDeviceCount?: number;
  maxDeviceCount?: number;
  includeInactiveUsers?: boolean;
  sortBy?: 'deviceCount' | 'userName' | 'lastEnrollment' | 'department';
  sortOrder?: 'asc' | 'desc';
  top?: number;
}

/**
 * Report options
 */
export interface EnrollmentTrackingReportOptions {
  filters?: EnrollmentTrackingFilters;
  includeDeviceDetails?: boolean;
  includeTrends?: boolean;
  includeBreakdowns?: boolean;
}

/**
 * Overall enrollment tracking statistics
 */
export interface EnrollmentTrackingStatistics {
  totalUsers: number;
  totalDevices: number;
  averageDevicesPerUser: number;
  medianDevicesPerUser: number;
  maxDevicesPerUser: number;
  minDevicesPerUser: number;
  usersWithMultipleDevices: number;
  usersWithSingleDevice: number;
  platformDistribution: PlatformBreakdown;
  enrollmentMethodDistribution: EnrollmentMethodBreakdown;
  ownerTypeDistribution: OwnerTypeBreakdown;
  complianceDistribution: {
    compliant: number;
    nonCompliant: number;
    unknown: number;
  };
  topDepartments: DepartmentSummary[];
  enrollmentTrends: {
    last7Days: number;
    last30Days: number;
    last90Days: number;
    last365Days: number;
  };
  generatedAt: string;
}

/**
 * Department summary
 */
export interface DepartmentSummary {
  department: string;
  userCount: number;
  deviceCount: number;
  averageDevicesPerUser: number;
}

/**
 * User licensing information
 */
export interface UserLicensingInfo {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  isLicensed: boolean;
  assignedLicenses: string[];
  intuneLicense: boolean;
  ems: boolean;
  microsoft365: boolean;
  accountEnabled: boolean;
}

/**
 * Device limit violation
 */
export interface DeviceLimitViolation {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  deviceCount: number;
  deviceLimit: number;
  excessDevices: number;
  violationSeverity: 'Low' | 'Medium' | 'High' | 'Critical';
  isLicensed: boolean;
  devices: UserEnrolledDevice[];
}

/**
 * Unusual enrollment pattern
 */
export interface UnusualEnrollmentPattern {
  patternType: 'RapidEnrollment' | 'BulkEnrollment' | 'UnusualLocation' | 'AfterHoursEnrollment' | 'SuspiciousDevice';
  severity: 'Low' | 'Medium' | 'High' | 'Critical';
  userId?: string;
  userPrincipalName?: string;
  displayName?: string;
  deviceCount: number;
  timeframe: string;
  description: string;
  affectedDevices: string[];
  detectedAt: string;
  recommendation: string;
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Intune Enrollment Tracking Reports Class
 *
 * Provides Configuration Manager-style enrollment tracking reports for Intune.
 * Implements two main reports:
 * - List of devices enrolled per user (detailed view)
 * - Number of devices enrolled per user (summary view)
 *
 * @example
 * ```typescript
 * const report = new IntuneEnrollmentTrackingReports(graphClient, config);
 *
 * // Get detailed device list per user
 * const deviceList = await report.getDevicesPerUserReport({
 *   filters: { department: 'IT' }
 * });
 *
 * // Get enrollment count summary
 * const enrollmentCounts = await report.getEnrollmentCountPerUserReport({
 *   filters: { minDeviceCount: 2 }
 * });
 * ```
 */
export class IntuneEnrollmentTrackingReports extends BaseReport {
  name = 'Intune Enrollment Tracking Reports';
  description = 'Configuration Manager-style enrollment tracking reports for Microsoft Intune';
  category = 'Configuration Manager Reports';
  enabled = true;

  /**
   * Execute the default enrollment tracking report
   */
  async execute(options: EnrollmentTrackingReportOptions = {}): Promise<ReportData> {
    logger.info('Starting Intune Enrollment Tracking Reports generation');

    try {
      const statistics = await this.generateEnrollmentStatistics(options.filters);

      return {
        metadata: this.createMetadata(
          'Intune Enrollment Tracking Overview',
          statistics.totalUsers,
          options
        ),
        data: [statistics],
        summary: {
          totalUsers: statistics.totalUsers,
          totalDevices: statistics.totalDevices,
          averageDevicesPerUser: statistics.averageDevicesPerUser
        }
      };
    } catch (error) {
      logger.error('Error generating Intune Enrollment Tracking Reports', error);
      throw error;
    }
  }

  // ==========================================================================
  // Report 17: List of Devices Enrolled Per User (Detailed)
  // ==========================================================================

  /**
   * Report 17: Get detailed list of devices enrolled per user
   *
   * Returns a comprehensive list of all devices each user has enrolled,
   * including device details, enrollment information, and breakdowns.
   *
   * @param options - Report configuration options
   * @returns Detailed device list per user
   *
   * @example
   * ```typescript
   * // Get all devices for all users
   * const report = await reports.getDevicesPerUserReport();
   *
   * // Filter by department
   * const itDevices = await reports.getDevicesPerUserReport({
   *   filters: { department: 'IT' }
   * });
   *
   * // Filter by platform
   * const iosDevices = await reports.getDevicesPerUserReport({
   *   filters: { platform: ['iOS'] }
   * });
   *
   * // Filter by date range
   * const recentDevices = await reports.getDevicesPerUserReport({
   *   filters: {
   *     startDate: new Date('2024-01-01'),
   *     endDate: new Date('2024-12-31')
   *   }
   * });
   * ```
   */
  async getDevicesPerUserReport(
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<ReportData> {
    logger.info('Generating Report 17: List of Devices Enrolled Per User');

    try {
      const filters = options.filters || {};
      const includeDeviceDetails = options.includeDeviceDetails !== false;
      const includeBreakdowns = options.includeBreakdowns !== false;

      // Fetch all managed devices
      const allDevices = await this.fetchAllManagedDevices();

      // Apply filters
      const filteredDevices = this.applyDeviceFilters(allDevices, filters);

      // Group devices by user
      const userDeviceMap = this.groupDevicesByUser(filteredDevices);

      // Fetch user details
      const userDeviceLists = await this.buildUserDeviceLists(
        userDeviceMap,
        includeDeviceDetails,
        includeBreakdowns,
        filters
      );

      // Sort results
      const sortedResults = this.sortUserDeviceLists(userDeviceLists, filters);

      // Apply top limit if specified
      const finalResults = filters.top
        ? sortedResults.slice(0, filters.top)
        : sortedResults;

      logger.info(`Report 17 complete: ${finalResults.length} users with devices`);

      return {
        metadata: this.createMetadata(
          'List of Devices Enrolled Per User in Microsoft Intune',
          finalResults.length,
          { filters, includeDeviceDetails, includeBreakdowns }
        ),
        data: finalResults,
        summary: {
          totalUsers: finalResults.length,
          totalDevices: finalResults.reduce((sum, u) => sum + u.totalDevices, 0),
          averageDevicesPerUser: finalResults.length > 0
            ? finalResults.reduce((sum, u) => sum + u.totalDevices, 0) / finalResults.length
            : 0
        }
      };
    } catch (error) {
      logger.error('Error generating devices per user report', error);
      throw error;
    }
  }

  /**
   * Get devices for a specific user
   *
   * @param userPrincipalName - User's UPN
   * @param options - Report options
   * @returns User's device list
   */
  async getDevicesForUser(
    userPrincipalName: string,
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<UserDeviceList | null> {
    logger.info(`Getting devices for user: ${userPrincipalName}`);

    const report = await this.getDevicesPerUserReport({
      ...options,
      filters: {
        ...options.filters,
        userPrincipalName
      }
    });

    return report.data[0] as UserDeviceList || null;
  }

  // ==========================================================================
  // Report 30: Number of Devices Enrolled Per User (Summary)
  // ==========================================================================

  /**
   * Report 30: Get enrollment count per user
   *
   * Returns a summary count of devices enrolled per user,
   * broken down by platform, owner type, and compliance state.
   *
   * @param options - Report configuration options
   * @returns Enrollment count summary per user
   *
   * @example
   * ```typescript
   * // Get enrollment counts for all users
   * const report = await reports.getEnrollmentCountPerUserReport();
   *
   * // Filter users with 2+ devices
   * const multiDevice = await reports.getEnrollmentCountPerUserReport({
   *   filters: { minDeviceCount: 2 }
   * });
   *
   * // Filter by department and sort by device count
   * const itCounts = await reports.getEnrollmentCountPerUserReport({
   *   filters: {
   *     department: 'IT',
   *     sortBy: 'deviceCount',
   *     sortOrder: 'desc'
   *   }
   * });
   *
   * // Get top 10 users by device count
   * const topUsers = await reports.getEnrollmentCountPerUserReport({
   *   filters: {
   *     sortBy: 'deviceCount',
   *     sortOrder: 'desc',
   *     top: 10
   *   }
   * });
   * ```
   */
  async getEnrollmentCountPerUserReport(
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<ReportData> {
    logger.info('Generating Report 30: Number of Devices Enrolled Per User');

    try {
      const filters = options.filters || {};
      const includeTrends = options.includeTrends !== false;

      // Fetch all managed devices
      const allDevices = await this.fetchAllManagedDevices();

      // Apply filters
      const filteredDevices = this.applyDeviceFilters(allDevices, filters);

      // Group devices by user
      const userDeviceMap = this.groupDevicesByUser(filteredDevices);

      // Build enrollment count summaries
      const enrollmentCounts = await this.buildEnrollmentCounts(
        userDeviceMap,
        includeTrends,
        filters
      );

      // Sort results
      const sortedResults = this.sortEnrollmentCounts(enrollmentCounts, filters);

      // Apply top limit if specified
      const finalResults = filters.top
        ? sortedResults.slice(0, filters.top)
        : sortedResults;

      logger.info(`Report 30 complete: ${finalResults.length} users with enrollment counts`);

      return {
        metadata: this.createMetadata(
          'Number of Devices Enrolled Per User in Microsoft Intune',
          finalResults.length,
          { filters, includeTrends }
        ),
        data: finalResults,
        summary: {
          totalUsers: finalResults.length,
          totalDevices: finalResults.reduce((sum, u) => sum + u.totalDeviceCount, 0),
          averageDevicesPerUser: finalResults.length > 0
            ? finalResults.reduce((sum, u) => sum + u.totalDeviceCount, 0) / finalResults.length
            : 0,
          usersWithMultipleDevices: finalResults.filter(u => u.totalDeviceCount > 1).length
        }
      };
    } catch (error) {
      logger.error('Error generating enrollment count per user report', error);
      throw error;
    }
  }

  /**
   * Get enrollment count for a specific user
   *
   * @param userPrincipalName - User's UPN
   * @param options - Report options
   * @returns User's enrollment count
   */
  async getEnrollmentCountForUser(
    userPrincipalName: string,
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<UserEnrollmentCount | null> {
    logger.info(`Getting enrollment count for user: ${userPrincipalName}`);

    const report = await this.getEnrollmentCountPerUserReport({
      ...options,
      filters: {
        ...options.filters,
        userPrincipalName
      }
    });

    return report.data[0] as UserEnrollmentCount || null;
  }

  // ==========================================================================
  // Statistics and Analytics
  // ==========================================================================

  /**
   * Generate overall enrollment tracking statistics
   *
   * @param filters - Optional filters
   * @returns Comprehensive enrollment statistics
   */
  async generateEnrollmentStatistics(
    filters?: EnrollmentTrackingFilters
  ): Promise<EnrollmentTrackingStatistics> {
    logger.info('Generating enrollment tracking statistics');

    const allDevices = await this.fetchAllManagedDevices();
    const filteredDevices = filters
      ? this.applyDeviceFilters(allDevices, filters)
      : allDevices;

    const userDeviceMap = this.groupDevicesByUser(filteredDevices);
    const deviceCounts = Array.from(userDeviceMap.values()).map(devices => devices.length);

    // Calculate statistics
    const totalUsers = userDeviceMap.size;
    const totalDevices = filteredDevices.length;
    const averageDevicesPerUser = totalUsers > 0 ? totalDevices / totalUsers : 0;

    // Calculate median
    const sortedCounts = [...deviceCounts].sort((a, b) => a - b);
    const medianDevicesPerUser = totalUsers > 0
      ? sortedCounts[Math.floor(totalUsers / 2)]
      : 0;

    const maxDevicesPerUser = deviceCounts.length > 0 ? Math.max(...deviceCounts) : 0;
    const minDevicesPerUser = deviceCounts.length > 0 ? Math.min(...deviceCounts) : 0;
    const usersWithMultipleDevices = deviceCounts.filter(count => count > 1).length;
    const usersWithSingleDevice = deviceCounts.filter(count => count === 1).length;

    // Platform distribution
    const platformDistribution = this.calculatePlatformDistribution(filteredDevices);

    // Enrollment method distribution
    const enrollmentMethodDistribution = this.calculateEnrollmentMethodDistribution(filteredDevices);

    // Owner type distribution
    const ownerTypeDistribution = this.calculateOwnerTypeDistribution(filteredDevices);

    // Compliance distribution
    const complianceDistribution = this.calculateComplianceDistribution(filteredDevices);

    // Department analysis
    const topDepartments = await this.calculateTopDepartments(userDeviceMap);

    // Enrollment trends
    const enrollmentTrends = this.calculateEnrollmentTrends(filteredDevices);

    return {
      totalUsers,
      totalDevices,
      averageDevicesPerUser: Math.round(averageDevicesPerUser * 100) / 100,
      medianDevicesPerUser,
      maxDevicesPerUser,
      minDevicesPerUser,
      usersWithMultipleDevices,
      usersWithSingleDevice,
      platformDistribution,
      enrollmentMethodDistribution,
      ownerTypeDistribution,
      complianceDistribution,
      topDepartments,
      enrollmentTrends,
      generatedAt: new Date().toISOString()
    };
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
          'imei',
          'complianceState',
          'managementState',
          'managementAgent',
          'userId',
          'userPrincipalName',
          'userDisplayName',
          'enrollmentType',
          'deviceEnrollmentType',
          'azureADDeviceId',
          'isSupervised',
          'deviceCategoryDisplayName',
          'enrollmentProfileName'
        ].join(','))
        .top(999)
        .get();

      return this.getAllPages<any>(response);
    });
  }

  /**
   * Fetch user details by ID
   */
  private async fetchUserDetails(userId: string): Promise<any> {
    try {
      return await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${userId}`)
          .select([
            'id',
            'userPrincipalName',
            'displayName',
            'department',
            'jobTitle',
            'mail',
            'accountEnabled'
          ].join(','))
          .get();
      });
    } catch (error) {
      logger.warn(`Unable to fetch user details for ${userId}`, error);
      return null;
    }
  }

  /**
   * Fetch user details by UPN
   */
  private async fetchUserDetailsByUPN(userPrincipalName: string): Promise<any> {
    try {
      return await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/users')
          .filter(`userPrincipalName eq '${userPrincipalName}'`)
          .select([
            'id',
            'userPrincipalName',
            'displayName',
            'department',
            'jobTitle',
            'mail',
            'accountEnabled'
          ].join(','))
          .top(1)
          .get()
          .then(response => response.value[0] || null);
      });
    } catch (error) {
      logger.warn(`Unable to fetch user details for ${userPrincipalName}`, error);
      return null;
    }
  }

  // ==========================================================================
  // Data Processing Methods
  // ==========================================================================

  /**
   * Apply filters to device list
   */
  private applyDeviceFilters(devices: any[], filters: EnrollmentTrackingFilters): any[] {
    let filtered = [...devices];

    // Filter by user
    if (filters.userPrincipalName) {
      filtered = filtered.filter(d =>
        d.userPrincipalName?.toLowerCase() === filters.userPrincipalName?.toLowerCase()
      );
    }

    if (filters.userId) {
      filtered = filtered.filter(d => d.userId === filters.userId);
    }

    // Filter by department (will be applied later when we have user details)

    // Filter by date range
    if (filters.startDate || filters.endDate) {
      filtered = filtered.filter(d => {
        if (!d.enrolledDateTime) return false;

        const enrollDate = new Date(d.enrolledDateTime);

        if (filters.startDate && enrollDate < filters.startDate) {
          return false;
        }

        if (filters.endDate && enrollDate > filters.endDate) {
          return false;
        }

        return true;
      });
    }

    // Filter by platform
    if (filters.platform && filters.platform.length > 0) {
      filtered = filtered.filter(d => {
        const platform = this.normalizePlatform(d.operatingSystem || '');
        return filters.platform!.includes(platform);
      });
    }

    // Filter by enrollment method
    if (filters.enrollmentMethod && filters.enrollmentMethod.length > 0) {
      filtered = filtered.filter(d => {
        const method = d.enrollmentType || d.deviceEnrollmentType || 'Unknown';
        return filters.enrollmentMethod!.includes(method);
      });
    }

    // Filter by owner type
    if (filters.ownerType && filters.ownerType !== 'All') {
      filtered = filtered.filter(d => {
        const ownerType = this.normalizeOwnerType(d.managedDeviceOwnerType);
        return ownerType === filters.ownerType;
      });
    }

    // Filter by compliance state
    if (filters.complianceState && filters.complianceState !== 'All') {
      filtered = filtered.filter(d => {
        if (filters.complianceState === 'Compliant') {
          return d.complianceState?.toLowerCase() === 'compliant';
        } else if (filters.complianceState === 'NonCompliant') {
          return d.complianceState?.toLowerCase() !== 'compliant';
        }
        return true;
      });
    }

    return filtered;
  }

  /**
   * Group devices by user
   */
  private groupDevicesByUser(devices: any[]): Map<string, any[]> {
    const userDeviceMap = new Map<string, any[]>();

    devices.forEach(device => {
      const userId = device.userId || device.userPrincipalName || 'Unknown';

      if (!userDeviceMap.has(userId)) {
        userDeviceMap.set(userId, []);
      }

      userDeviceMap.get(userId)!.push(device);
    });

    return userDeviceMap;
  }

  /**
   * Build user device lists with details
   */
  private async buildUserDeviceLists(
    userDeviceMap: Map<string, any[]>,
    includeDeviceDetails: boolean,
    includeBreakdowns: boolean,
    filters: EnrollmentTrackingFilters
  ): Promise<UserDeviceList[]> {
    const userDeviceLists: UserDeviceList[] = [];

    for (const [userId, devices] of userDeviceMap.entries()) {
      // Skip if doesn't meet device count criteria
      if (filters.minDeviceCount && devices.length < filters.minDeviceCount) {
        continue;
      }

      if (filters.maxDeviceCount && devices.length > filters.maxDeviceCount) {
        continue;
      }

      const firstDevice = devices[0];

      // Try to get detailed user info
      let userDetails: any = null;
      if (userId && userId !== 'Unknown') {
        userDetails = await this.fetchUserDetails(userId);
      }

      // Apply department filter if specified
      if (filters.department) {
        const userDept = userDetails?.department || firstDevice.department || '';
        if (userDept.toLowerCase() !== filters.department.toLowerCase()) {
          continue;
        }
      }

      // Build enrolled device list
      const enrolledDevices: UserEnrolledDevice[] = includeDeviceDetails
        ? devices.map(d => this.mapToUserEnrolledDevice(d))
        : [];

      // Calculate breakdowns
      const devicesByPlatform = includeBreakdowns
        ? this.calculatePlatformBreakdown(devices)
        : this.createEmptyPlatformBreakdown();

      const devicesByEnrollmentMethod = includeBreakdowns
        ? this.calculateEnrollmentMethodBreakdown(devices)
        : {};

      const devicesByOwnerType = includeBreakdowns
        ? this.calculateOwnerTypeBreakdown(devices)
        : { Corporate: 0, Personal: 0, Unknown: 0 };

      // Get enrollment dates
      const enrollmentDates = devices
        .map(d => d.enrolledDateTime)
        .filter(d => d)
        .sort();

      const firstEnrollmentDate = enrollmentDates[0];
      const lastEnrollmentDate = enrollmentDates[enrollmentDates.length - 1];

      // Calculate average device age in days
      const now = new Date();
      const deviceAges = devices
        .map(d => {
          if (!d.enrolledDateTime) return 0;
          return (now.getTime() - new Date(d.enrolledDateTime).getTime()) / (1000 * 60 * 60 * 24);
        })
        .filter(age => age > 0);

      const averageDeviceAge = deviceAges.length > 0
        ? Math.round(deviceAges.reduce((a, b) => a + b, 0) / deviceAges.length)
        : 0;

      userDeviceLists.push({
        userId: userDetails?.id || userId,
        userPrincipalName: userDetails?.userPrincipalName || firstDevice.userPrincipalName || 'Unknown',
        displayName: userDetails?.displayName || firstDevice.userDisplayName || 'Unknown',
        department: userDetails?.department || firstDevice.department,
        jobTitle: userDetails?.jobTitle || firstDevice.jobTitle,
        email: userDetails?.mail,
        totalDevices: devices.length,
        devices: enrolledDevices,
        devicesByPlatform,
        devicesByEnrollmentMethod,
        devicesByOwnerType,
        firstEnrollmentDate,
        lastEnrollmentDate,
        averageDeviceAge
      });
    }

    return userDeviceLists;
  }

  /**
   * Build enrollment count summaries
   */
  private async buildEnrollmentCounts(
    userDeviceMap: Map<string, any[]>,
    includeTrends: boolean,
    filters: EnrollmentTrackingFilters
  ): Promise<UserEnrollmentCount[]> {
    const enrollmentCounts: UserEnrollmentCount[] = [];

    for (const [userId, devices] of userDeviceMap.entries()) {
      // Skip if doesn't meet device count criteria
      if (filters.minDeviceCount && devices.length < filters.minDeviceCount) {
        continue;
      }

      if (filters.maxDeviceCount && devices.length > filters.maxDeviceCount) {
        continue;
      }

      const firstDevice = devices[0];

      // Try to get detailed user info
      let userDetails: any = null;
      if (userId && userId !== 'Unknown') {
        userDetails = await this.fetchUserDetails(userId);
      }

      // Apply department filter if specified
      if (filters.department) {
        const userDept = userDetails?.department || firstDevice.department || '';
        if (userDept.toLowerCase() !== filters.department.toLowerCase()) {
          continue;
        }
      }

      // Count devices by platform
      const platformCounts = this.calculatePlatformBreakdown(devices);

      // Count devices by owner type
      const corporateDeviceCount = devices.filter(d =>
        this.normalizeOwnerType(d.managedDeviceOwnerType) === 'Corporate'
      ).length;

      const personalDeviceCount = devices.filter(d =>
        this.normalizeOwnerType(d.managedDeviceOwnerType) === 'Personal'
      ).length;

      // Count compliant devices
      const compliantDeviceCount = devices.filter(d =>
        d.complianceState?.toLowerCase() === 'compliant'
      ).length;

      const nonCompliantDeviceCount = devices.length - compliantDeviceCount;

      // Count active devices (synced in last 30 days)
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

      const activeDeviceCount = devices.filter(d =>
        d.lastSyncDateTime && new Date(d.lastSyncDateTime) >= thirtyDaysAgo
      ).length;

      const inactiveDeviceCount = devices.length - activeDeviceCount;

      // Calculate enrollment trend
      const enrollmentTrend = includeTrends
        ? this.calculateUserEnrollmentTrend(devices)
        : this.createEmptyEnrollmentTrend();

      // Get enrollment dates
      const enrollmentDates = devices
        .map(d => d.enrolledDateTime)
        .filter(d => d)
        .sort();

      enrollmentCounts.push({
        userId: userDetails?.id || userId,
        userPrincipalName: userDetails?.userPrincipalName || firstDevice.userPrincipalName || 'Unknown',
        displayName: userDetails?.displayName || firstDevice.userDisplayName || 'Unknown',
        department: userDetails?.department || firstDevice.department,
        jobTitle: userDetails?.jobTitle || firstDevice.jobTitle,
        email: userDetails?.mail,
        totalDeviceCount: devices.length,
        iosDeviceCount: platformCounts.iOS,
        androidDeviceCount: platformCounts.Android,
        windowsDeviceCount: platformCounts.Windows,
        macOsDeviceCount: platformCounts.macOS,
        linuxDeviceCount: platformCounts.Linux,
        otherDeviceCount: platformCounts.Other,
        corporateDeviceCount,
        personalDeviceCount,
        compliantDeviceCount,
        nonCompliantDeviceCount,
        activeDeviceCount,
        inactiveDeviceCount,
        enrollmentTrend,
        lastEnrollmentDate: enrollmentDates[enrollmentDates.length - 1],
        firstEnrollmentDate: enrollmentDates[0]
      });
    }

    return enrollmentCounts;
  }

  /**
   * Map device to UserEnrolledDevice format
   */
  private mapToUserEnrolledDevice(device: any): UserEnrolledDevice {
    return {
      deviceId: device.id || '',
      deviceName: device.deviceName || 'Unknown',
      platform: this.normalizePlatform(device.operatingSystem || ''),
      osVersion: device.osVersion || 'Unknown',
      model: device.model,
      manufacturer: device.manufacturer,
      enrollmentDate: device.enrolledDateTime || '',
      enrollmentType: device.enrollmentType || 'Unknown',
      enrollmentMethod: device.deviceEnrollmentType || device.enrollmentType || 'Unknown',
      managementState: device.managementState || 'Unknown',
      complianceState: device.complianceState || 'Unknown',
      lastSyncDateTime: device.lastSyncDateTime || '',
      ownerType: this.normalizeOwnerType(device.managedDeviceOwnerType),
      serialNumber: device.serialNumber,
      imei: device.imei,
      isSupervised: device.isSupervised,
      deviceCategory: device.deviceCategoryDisplayName,
      enrollmentProfileName: device.enrollmentProfileName
    };
  }

  // ==========================================================================
  // Calculation Methods
  // ==========================================================================

  /**
   * Calculate platform breakdown for devices
   */
  private calculatePlatformBreakdown(devices: any[]): PlatformBreakdown {
    const breakdown: PlatformBreakdown = {
      iOS: 0,
      Android: 0,
      Windows: 0,
      macOS: 0,
      Linux: 0,
      Other: 0
    };

    devices.forEach(device => {
      const platform = this.normalizePlatform(device.operatingSystem || '');
      if (platform in breakdown) {
        (breakdown as any)[platform]++;
      }
    });

    return breakdown;
  }

  /**
   * Calculate platform distribution
   */
  private calculatePlatformDistribution(devices: any[]): PlatformBreakdown {
    return this.calculatePlatformBreakdown(devices);
  }

  /**
   * Calculate enrollment method breakdown
   */
  private calculateEnrollmentMethodBreakdown(devices: any[]): EnrollmentMethodBreakdown {
    const breakdown: EnrollmentMethodBreakdown = {};

    devices.forEach(device => {
      const method = device.enrollmentType || device.deviceEnrollmentType || 'Unknown';
      breakdown[method] = (breakdown[method] || 0) + 1;
    });

    return breakdown;
  }

  /**
   * Calculate enrollment method distribution
   */
  private calculateEnrollmentMethodDistribution(devices: any[]): EnrollmentMethodBreakdown {
    return this.calculateEnrollmentMethodBreakdown(devices);
  }

  /**
   * Calculate owner type breakdown
   */
  private calculateOwnerTypeBreakdown(devices: any[]): OwnerTypeBreakdown {
    const breakdown: OwnerTypeBreakdown = {
      Corporate: 0,
      Personal: 0,
      Unknown: 0
    };

    devices.forEach(device => {
      const ownerType = this.normalizeOwnerType(device.managedDeviceOwnerType);
      if (ownerType in breakdown) {
        (breakdown as any)[ownerType]++;
      }
    });

    return breakdown;
  }

  /**
   * Calculate owner type distribution
   */
  private calculateOwnerTypeDistribution(devices: any[]): OwnerTypeBreakdown {
    return this.calculateOwnerTypeBreakdown(devices);
  }

  /**
   * Calculate compliance distribution
   */
  private calculateComplianceDistribution(devices: any[]): {
    compliant: number;
    nonCompliant: number;
    unknown: number;
  } {
    const distribution = {
      compliant: 0,
      nonCompliant: 0,
      unknown: 0
    };

    devices.forEach(device => {
      const state = device.complianceState?.toLowerCase() || '';

      if (state === 'compliant') {
        distribution.compliant++;
      } else if (state === 'noncompliant') {
        distribution.nonCompliant++;
      } else {
        distribution.unknown++;
      }
    });

    return distribution;
  }

  /**
   * Calculate enrollment trends
   */
  private calculateEnrollmentTrends(devices: any[]): EnrollmentTrend {
    const now = new Date();
    const last7Days = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
    const last30Days = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000);
    const last90Days = new Date(now.getTime() - 90 * 24 * 60 * 60 * 1000);
    const last365Days = new Date(now.getTime() - 365 * 24 * 60 * 60 * 1000);

    return {
      last7Days: devices.filter(d =>
        d.enrolledDateTime && new Date(d.enrolledDateTime) >= last7Days
      ).length,
      last30Days: devices.filter(d =>
        d.enrolledDateTime && new Date(d.enrolledDateTime) >= last30Days
      ).length,
      last90Days: devices.filter(d =>
        d.enrolledDateTime && new Date(d.enrolledDateTime) >= last90Days
      ).length,
      last365Days: devices.filter(d =>
        d.enrolledDateTime && new Date(d.enrolledDateTime) >= last365Days
      ).length
    };
  }

  /**
   * Calculate user enrollment trend
   */
  private calculateUserEnrollmentTrend(devices: any[]): EnrollmentTrend {
    return this.calculateEnrollmentTrends(devices);
  }

  /**
   * Calculate top departments
   */
  private async calculateTopDepartments(
    userDeviceMap: Map<string, any[]>
  ): Promise<DepartmentSummary[]> {
    const departmentMap = new Map<string, { userIds: Set<string>; deviceCount: number }>();

    for (const [userId, devices] of userDeviceMap.entries()) {
      const firstDevice = devices[0];
      const department = firstDevice.department || 'Unknown';

      if (!departmentMap.has(department)) {
        departmentMap.set(department, { userIds: new Set(), deviceCount: 0 });
      }

      const deptData = departmentMap.get(department)!;
      deptData.userIds.add(userId);
      deptData.deviceCount += devices.length;
    }

    const departments: DepartmentSummary[] = [];

    for (const [department, data] of departmentMap.entries()) {
      departments.push({
        department,
        userCount: data.userIds.size,
        deviceCount: data.deviceCount,
        averageDevicesPerUser: Math.round((data.deviceCount / data.userIds.size) * 100) / 100
      });
    }

    return departments
      .sort((a, b) => b.deviceCount - a.deviceCount)
      .slice(0, 10);
  }

  // ==========================================================================
  // Additional User Enrollment Methods
  // ==========================================================================

  /**
   * Get devices enrolled per user (by userId or UPN)
   *
   * This is a flexible method that accepts either userId or userPrincipalName
   * and returns all enrolled devices for that user.
   *
   * @param userIdOrUpn - User ID or User Principal Name
   * @returns List of devices enrolled by the user
   *
   * @example
   * ```typescript
   * // By UPN
   * const devices = await reports.getDevicesEnrolledPerUser('john.doe@contoso.com');
   *
   * // By User ID
   * const devices = await reports.getDevicesEnrolledPerUser('a1b2c3d4-e5f6-7890-abcd-ef1234567890');
   * ```
   */
  async getDevicesEnrolledPerUser(userIdOrUpn: string): Promise<UserDeviceList | null> {
    logger.info(`Getting devices enrolled per user: ${userIdOrUpn}`);

    // Determine if input is UPN or userId (UPN contains @)
    const isUpn = userIdOrUpn.includes('@');

    const report = await this.getDevicesPerUserReport({
      filters: isUpn
        ? { userPrincipalName: userIdOrUpn }
        : { userId: userIdOrUpn }
    });

    return report.data[0] as UserDeviceList || null;
  }

  /**
   * Get all users with device counts
   *
   * Returns a comprehensive list of all users who have enrolled devices
   * along with their device counts and statistics.
   *
   * @param options - Optional report options
   * @returns Array of user enrollment counts
   *
   * @example
   * ```typescript
   * const allUsers = await reports.getAllUsersWithDeviceCounts();
   *
   * // Filter by department
   * const itUsers = await reports.getAllUsersWithDeviceCounts({
   *   filters: { department: 'IT' }
   * });
   * ```
   */
  async getAllUsersWithDeviceCounts(
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<UserEnrollmentCount[]> {
    logger.info('Getting all users with device counts');

    const report = await this.getEnrollmentCountPerUserReport(options);
    return report.data as UserEnrollmentCount[];
  }

  /**
   * Get users with multiple devices (exceeding threshold)
   *
   * Identifies users who have enrolled more than the specified number of devices.
   * Useful for detecting policy violations or unusual enrollment patterns.
   *
   * @param threshold - Minimum number of devices (default: 1)
   * @param options - Optional report options
   * @returns Array of users exceeding the threshold
   *
   * @example
   * ```typescript
   * // Find users with 3+ devices
   * const multiDeviceUsers = await reports.getUsersWithMultipleDevices(3);
   *
   * // Find users with 5+ devices in IT department
   * const itMultiDevice = await reports.getUsersWithMultipleDevices(5, {
   *   filters: { department: 'IT' }
   * });
   * ```
   */
  async getUsersWithMultipleDevices(
    threshold: number = 1,
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<UserEnrollmentCount[]> {
    logger.info(`Finding users with more than ${threshold} devices`);

    const report = await this.getEnrollmentCountPerUserReport({
      ...options,
      filters: {
        ...options.filters,
        minDeviceCount: threshold + 1
      }
    });

    const users = report.data as UserEnrollmentCount[];

    logger.info(`Found ${users.length} users with more than ${threshold} devices`);
    return users;
  }

  /**
   * Get detailed enrollment information for a specific user
   *
   * Returns comprehensive enrollment details including devices, licensing,
   * enrollment history, and compliance status.
   *
   * @param userIdOrUpn - User ID or User Principal Name
   * @returns Detailed user enrollment information
   *
   * @example
   * ```typescript
   * const details = await reports.getUserEnrollmentDetails('john.doe@contoso.com');
   * console.log(`Total devices: ${details.deviceCount}`);
   * console.log(`Licensed: ${details.licensing.isLicensed}`);
   * ```
   */
  async getUserEnrollmentDetails(userIdOrUpn: string): Promise<{
    user: UserDeviceList | null;
    enrollmentCount: UserEnrollmentCount | null;
    licensing: UserLicensingInfo | null;
    violatesLimit: boolean;
    limitViolation?: DeviceLimitViolation;
  }> {
    logger.info(`Getting enrollment details for user: ${userIdOrUpn}`);

    // Get device list
    const userDeviceList = await this.getDevicesEnrolledPerUser(userIdOrUpn);

    // Get enrollment count
    const isUpn = userIdOrUpn.includes('@');
    const countReport = await this.getEnrollmentCountPerUserReport({
      filters: isUpn
        ? { userPrincipalName: userIdOrUpn }
        : { userId: userIdOrUpn }
    });
    const enrollmentCount = countReport.data[0] as UserEnrollmentCount || null;

    // Get licensing information
    let licensing: UserLicensingInfo | null = null;
    try {
      licensing = await this.getUserLicensingInfo(userIdOrUpn);
    } catch (error) {
      logger.warn(`Unable to fetch licensing info for ${userIdOrUpn}`, error);
    }

    // Check device limit violation
    const deviceLimit = this.config.deviceEnrollmentLimit || 15; // Default limit
    const deviceCount = enrollmentCount?.totalDeviceCount || 0;
    const violatesLimit = deviceCount > deviceLimit;

    let limitViolation: DeviceLimitViolation | undefined;
    if (violatesLimit && userDeviceList && enrollmentCount) {
      limitViolation = {
        userId: enrollmentCount.userId,
        userPrincipalName: enrollmentCount.userPrincipalName,
        displayName: enrollmentCount.displayName,
        deviceCount,
        deviceLimit,
        excessDevices: deviceCount - deviceLimit,
        violationSeverity: this.calculateViolationSeverity(deviceCount, deviceLimit),
        isLicensed: licensing?.isLicensed || false,
        devices: userDeviceList.devices
      };
    }

    return {
      user: userDeviceList,
      enrollmentCount,
      licensing,
      violatesLimit,
      limitViolation
    };
  }

  /**
   * Get enrollment statistics aggregated by user
   *
   * Returns comprehensive statistics about device enrollments
   * grouped and analyzed by user.
   *
   * @param options - Optional report options
   * @returns Enrollment statistics by user
   *
   * @example
   * ```typescript
   * const stats = await reports.getEnrollmentStatisticsByUser();
   * console.log(`Average devices per user: ${stats.averageDevicesPerUser}`);
   * console.log(`Users with violations: ${stats.usersViolatingLimit}`);
   * ```
   */
  async getEnrollmentStatisticsByUser(
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<{
    totalUsers: number;
    totalDevices: number;
    averageDevicesPerUser: number;
    medianDevicesPerUser: number;
    usersWithSingleDevice: number;
    usersWithMultipleDevices: number;
    usersViolatingLimit: number;
    topUsersByDeviceCount: UserEnrollmentCount[];
    departmentBreakdown: DepartmentSummary[];
    platformDistribution: PlatformBreakdown;
    complianceOverview: {
      totalCompliant: number;
      totalNonCompliant: number;
      complianceRate: number;
    };
    enrollmentTrends: EnrollmentTrend;
    generatedAt: string;
  }> {
    logger.info('Generating enrollment statistics by user');

    const allUsers = await this.getAllUsersWithDeviceCounts(options);
    const deviceLimit = this.config.deviceEnrollmentLimit || 15;

    // Calculate statistics
    const totalUsers = allUsers.length;
    const totalDevices = allUsers.reduce((sum, u) => sum + u.totalDeviceCount, 0);
    const averageDevicesPerUser = totalUsers > 0 ? totalDevices / totalUsers : 0;

    // Calculate median
    const sortedCounts = allUsers.map(u => u.totalDeviceCount).sort((a, b) => a - b);
    const medianDevicesPerUser = totalUsers > 0
      ? sortedCounts[Math.floor(totalUsers / 2)]
      : 0;

    const usersWithSingleDevice = allUsers.filter(u => u.totalDeviceCount === 1).length;
    const usersWithMultipleDevices = allUsers.filter(u => u.totalDeviceCount > 1).length;
    const usersViolatingLimit = allUsers.filter(u => u.totalDeviceCount > deviceLimit).length;

    // Top users by device count
    const topUsersByDeviceCount = [...allUsers]
      .sort((a, b) => b.totalDeviceCount - a.totalDeviceCount)
      .slice(0, 10);

    // Department breakdown
    const deptMap = new Map<string, UserEnrollmentCount[]>();
    allUsers.forEach(user => {
      const dept = user.department || 'Unknown';
      if (!deptMap.has(dept)) {
        deptMap.set(dept, []);
      }
      deptMap.get(dept)!.push(user);
    });

    const departmentBreakdown: DepartmentSummary[] = Array.from(deptMap.entries())
      .map(([dept, users]) => ({
        department: dept,
        userCount: users.length,
        deviceCount: users.reduce((sum, u) => sum + u.totalDeviceCount, 0),
        averageDevicesPerUser: users.reduce((sum, u) => sum + u.totalDeviceCount, 0) / users.length
      }))
      .sort((a, b) => b.deviceCount - a.deviceCount);

    // Platform distribution
    const platformDistribution: PlatformBreakdown = {
      iOS: allUsers.reduce((sum, u) => sum + u.iosDeviceCount, 0),
      Android: allUsers.reduce((sum, u) => sum + u.androidDeviceCount, 0),
      Windows: allUsers.reduce((sum, u) => sum + u.windowsDeviceCount, 0),
      macOS: allUsers.reduce((sum, u) => sum + u.macOsDeviceCount, 0),
      Linux: allUsers.reduce((sum, u) => sum + u.linuxDeviceCount, 0),
      Other: allUsers.reduce((sum, u) => sum + u.otherDeviceCount, 0)
    };

    // Compliance overview
    const totalCompliant = allUsers.reduce((sum, u) => sum + u.compliantDeviceCount, 0);
    const totalNonCompliant = allUsers.reduce((sum, u) => sum + u.nonCompliantDeviceCount, 0);
    const complianceRate = totalDevices > 0 ? (totalCompliant / totalDevices) * 100 : 0;

    // Enrollment trends (aggregate from all users)
    const enrollmentTrends: EnrollmentTrend = {
      last7Days: allUsers.reduce((sum, u) => sum + u.enrollmentTrend.last7Days, 0),
      last30Days: allUsers.reduce((sum, u) => sum + u.enrollmentTrend.last30Days, 0),
      last90Days: allUsers.reduce((sum, u) => sum + u.enrollmentTrend.last90Days, 0),
      last365Days: allUsers.reduce((sum, u) => sum + u.enrollmentTrend.last365Days, 0)
    };

    return {
      totalUsers,
      totalDevices,
      averageDevicesPerUser: Math.round(averageDevicesPerUser * 100) / 100,
      medianDevicesPerUser,
      usersWithSingleDevice,
      usersWithMultipleDevices,
      usersViolatingLimit,
      topUsersByDeviceCount,
      departmentBreakdown,
      platformDistribution,
      complianceOverview: {
        totalCompliant,
        totalNonCompliant,
        complianceRate: Math.round(complianceRate * 100) / 100
      },
      enrollmentTrends,
      generatedAt: new Date().toISOString()
    };
  }

  /**
   * Detect unusual enrollment patterns
   *
   * Analyzes enrollment data to identify suspicious or unusual patterns
   * that may indicate security issues, policy violations, or other anomalies.
   *
   * @param options - Optional report options
   * @returns Array of detected unusual patterns
   *
   * @example
   * ```typescript
   * const patterns = await reports.detectUnusualEnrollmentPatterns();
   *
   * patterns.forEach(pattern => {
   *   console.log(`${pattern.severity}: ${pattern.description}`);
   *   console.log(`Recommendation: ${pattern.recommendation}`);
   * });
   * ```
   */
  async detectUnusualEnrollmentPatterns(
    options: EnrollmentTrackingReportOptions = {}
  ): Promise<UnusualEnrollmentPattern[]> {
    logger.info('Detecting unusual enrollment patterns');

    const patterns: UnusualEnrollmentPattern[] = [];
    const allUsers = await this.getAllUsersWithDeviceCounts(options);
    const allDevices = await this.fetchAllManagedDevices();

    // 1. Detect rapid enrollment (multiple devices enrolled in short time)
    const rapidEnrollmentPatterns = this.detectRapidEnrollment(allUsers, allDevices);
    patterns.push(...rapidEnrollmentPatterns);

    // 2. Detect bulk enrollment (many devices from same user in single day)
    const bulkEnrollmentPatterns = this.detectBulkEnrollment(allUsers, allDevices);
    patterns.push(...bulkEnrollmentPatterns);

    // 3. Detect after-hours enrollment
    const afterHoursPatterns = this.detectAfterHoursEnrollment(allDevices);
    patterns.push(...afterHoursPatterns);

    // 4. Detect users with excessive devices
    const excessiveDevicePatterns = this.detectExcessiveDevices(allUsers);
    patterns.push(...excessiveDevicePatterns);

    // 5. Detect suspicious device characteristics
    const suspiciousDevicePatterns = this.detectSuspiciousDevices(allDevices);
    patterns.push(...suspiciousDevicePatterns);

    // Sort by severity
    patterns.sort((a, b) => {
      const severityOrder = { Critical: 4, High: 3, Medium: 2, Low: 1 };
      return severityOrder[b.severity] - severityOrder[a.severity];
    });

    logger.info(`Detected ${patterns.length} unusual enrollment patterns`);
    return patterns;
  }

  // ==========================================================================
  // Licensing Methods
  // ==========================================================================

  /**
   * Get user licensing information
   *
   * @param userIdOrUpn - User ID or User Principal Name
   * @returns User licensing details
   */
  private async getUserLicensingInfo(userIdOrUpn: string): Promise<UserLicensingInfo | null> {
    logger.info(`Fetching licensing info for user: ${userIdOrUpn}`);

    try {
      const isUpn = userIdOrUpn.includes('@');

      let user: any;
      if (isUpn) {
        user = await this.fetchUserDetailsByUPN(userIdOrUpn);
      } else {
        user = await this.fetchUserDetails(userIdOrUpn);
      }

      if (!user) {
        return null;
      }

      // Fetch license details
      const licensesResponse = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${user.id}`)
          .select('assignedLicenses')
          .get();
      });

      const assignedLicenses = licensesResponse.assignedLicenses || [];
      const licenseSkuIds = assignedLicenses.map((l: any) => l.skuId);

      // Known Intune/EMS SKU IDs (partial list - expand as needed)
      const intuneSkuIds = [
        'c1ec4a95-1f05-45b3-a911-aa3fa01094f5', // EMS E3
        '9e3f2f20-cf3e-4f69-8e93-b96e1e3e3c5e', // EMS E5
        '184efa21-98c3-4e5d-95ab-d07053a96e67', // Microsoft 365 E3
        '44ac31e7-2999-4304-ad94-c948886741d4', // Microsoft 365 E5
        'e43b5b99-8dfb-405f-9987-dc307f34bcbd', // Intune Plan 1
      ];

      const hasIntuneLicense = licenseSkuIds.some((sku: string) => intuneSkuIds.includes(sku));

      return {
        userId: user.id,
        userPrincipalName: user.userPrincipalName,
        displayName: user.displayName,
        isLicensed: assignedLicenses.length > 0,
        assignedLicenses: licenseSkuIds,
        intuneLicense: hasIntuneLicense,
        ems: licenseSkuIds.some((sku: string) => sku.includes('c1ec4a95') || sku.includes('9e3f2f20')),
        microsoft365: licenseSkuIds.some((sku: string) => sku.includes('184efa21') || sku.includes('44ac31e7')),
        accountEnabled: user.accountEnabled !== false
      };
    } catch (error) {
      logger.error(`Failed to fetch licensing info for ${userIdOrUpn}`, error);
      return null;
    }
  }

  // ==========================================================================
  // Pattern Detection Methods
  // ==========================================================================

  /**
   * Detect rapid enrollment patterns
   */
  private detectRapidEnrollment(
    users: UserEnrollmentCount[],
    devices: any[]
  ): UnusualEnrollmentPattern[] {
    const patterns: UnusualEnrollmentPattern[] = [];
    const threshold = 5; // 5+ devices in 24 hours
    const timeWindowHours = 24;

    users.forEach(user => {
      const userDevices = devices.filter(d =>
        d.userPrincipalName?.toLowerCase() === user.userPrincipalName.toLowerCase()
      );

      // Check for rapid enrollment
      for (let i = 0; i < userDevices.length; i++) {
        const device = userDevices[i];
        if (!device.enrolledDateTime) continue;

        const enrollDate = new Date(device.enrolledDateTime);
        const rapidDevices = userDevices.filter(d => {
          if (!d.enrolledDateTime) return false;
          const dDate = new Date(d.enrolledDateTime);
          const hoursDiff = (dDate.getTime() - enrollDate.getTime()) / (1000 * 60 * 60);
          return hoursDiff >= 0 && hoursDiff <= timeWindowHours;
        });

        if (rapidDevices.length >= threshold) {
          patterns.push({
            patternType: 'RapidEnrollment',
            severity: rapidDevices.length >= 10 ? 'Critical' : 'High',
            userId: user.userId,
            userPrincipalName: user.userPrincipalName,
            displayName: user.displayName,
            deviceCount: rapidDevices.length,
            timeframe: `${timeWindowHours} hours`,
            description: `User enrolled ${rapidDevices.length} devices within ${timeWindowHours} hours`,
            affectedDevices: rapidDevices.map(d => d.deviceName || d.id),
            detectedAt: new Date().toISOString(),
            recommendation: 'Investigate potential unauthorized bulk enrollment or automated enrollment script'
          });
          break; // Only report once per user
        }
      }
    });

    return patterns;
  }

  /**
   * Detect bulk enrollment patterns
   */
  private detectBulkEnrollment(
    users: UserEnrollmentCount[],
    devices: any[]
  ): UnusualEnrollmentPattern[] {
    const patterns: UnusualEnrollmentPattern[] = [];

    // Check for many enrollments on same day
    const enrollmentsByDay = new Map<string, any[]>();

    devices.forEach(device => {
      if (!device.enrolledDateTime) return;
      const date = new Date(device.enrolledDateTime).toISOString().split('T')[0];
      if (!enrollmentsByDay.has(date)) {
        enrollmentsByDay.set(date, []);
      }
      enrollmentsByDay.get(date)!.push(device);
    });

    enrollmentsByDay.forEach((dayDevices, date) => {
      if (dayDevices.length >= 20) { // 20+ devices in one day
        patterns.push({
          patternType: 'BulkEnrollment',
          severity: dayDevices.length >= 50 ? 'Critical' : 'Medium',
          deviceCount: dayDevices.length,
          timeframe: date,
          description: `${dayDevices.length} devices enrolled on ${date}`,
          affectedDevices: dayDevices.map(d => d.deviceName || d.id),
          detectedAt: new Date().toISOString(),
          recommendation: 'Verify this is expected bulk enrollment (e.g., new employee onboarding, device refresh)'
        });
      }
    });

    return patterns;
  }

  /**
   * Detect after-hours enrollment
   */
  private detectAfterHoursEnrollment(devices: any[]): UnusualEnrollmentPattern[] {
    const patterns: UnusualEnrollmentPattern[] = [];
    const afterHoursDevices: any[] = [];

    devices.forEach(device => {
      if (!device.enrolledDateTime) return;

      const enrollDate = new Date(device.enrolledDateTime);
      const hour = enrollDate.getUTCHours();

      // After hours: 10 PM - 6 AM UTC (adjust based on organization timezone)
      if (hour >= 22 || hour < 6) {
        afterHoursDevices.push(device);
      }
    });

    if (afterHoursDevices.length >= 5) {
      patterns.push({
        patternType: 'AfterHoursEnrollment',
        severity: afterHoursDevices.length >= 20 ? 'High' : 'Low',
        deviceCount: afterHoursDevices.length,
        timeframe: 'After business hours (10 PM - 6 AM UTC)',
        description: `${afterHoursDevices.length} devices enrolled outside business hours`,
        affectedDevices: afterHoursDevices.map(d => d.deviceName || d.id).slice(0, 10),
        detectedAt: new Date().toISOString(),
        recommendation: 'Review after-hours enrollments for potential unauthorized access or global deployment'
      });
    }

    return patterns;
  }

  /**
   * Detect users with excessive devices
   */
  private detectExcessiveDevices(users: UserEnrollmentCount[]): UnusualEnrollmentPattern[] {
    const patterns: UnusualEnrollmentPattern[] = [];
    const limit = this.config.deviceEnrollmentLimit || 15;

    users.forEach(user => {
      if (user.totalDeviceCount > limit * 2) { // 2x the limit
        patterns.push({
          patternType: 'BulkEnrollment',
          severity: user.totalDeviceCount > limit * 3 ? 'Critical' : 'High',
          userId: user.userId,
          userPrincipalName: user.userPrincipalName,
          displayName: user.displayName,
          deviceCount: user.totalDeviceCount,
          timeframe: 'Current',
          description: `User has ${user.totalDeviceCount} enrolled devices (limit: ${limit})`,
          affectedDevices: [],
          detectedAt: new Date().toISOString(),
          recommendation: `Review and clean up devices for this user. Device limit violation: ${user.totalDeviceCount - limit} excess devices`
        });
      }
    });

    return patterns;
  }

  /**
   * Detect suspicious devices
   */
  private detectSuspiciousDevices(devices: any[]): UnusualEnrollmentPattern[] {
    const patterns: UnusualEnrollmentPattern[] = [];
    const suspiciousDevices: any[] = [];

    devices.forEach(device => {
      // Check for suspicious characteristics
      let isSuspicious = false;

      // 1. Device with no name or generic name
      if (!device.deviceName || device.deviceName === 'Unknown' || device.deviceName.length < 3) {
        isSuspicious = true;
      }

      // 2. Device never synced
      if (!device.lastSyncDateTime) {
        const enrollDate = new Date(device.enrolledDateTime);
        const daysSinceEnroll = (Date.now() - enrollDate.getTime()) / (1000 * 60 * 60 * 24);
        if (daysSinceEnroll > 7) {
          isSuspicious = true;
        }
      }

      // 3. Non-compliant and old enrollment
      if (device.complianceState === 'noncompliant') {
        const enrollDate = new Date(device.enrolledDateTime);
        const daysSinceEnroll = (Date.now() - enrollDate.getTime()) / (1000 * 60 * 60 * 24);
        if (daysSinceEnroll > 30) {
          isSuspicious = true;
        }
      }

      if (isSuspicious) {
        suspiciousDevices.push(device);
      }
    });

    if (suspiciousDevices.length > 0) {
      patterns.push({
        patternType: 'SuspiciousDevice',
        severity: 'Medium',
        deviceCount: suspiciousDevices.length,
        timeframe: 'Current',
        description: `${suspiciousDevices.length} devices with suspicious characteristics (no name, never synced, or long-term non-compliant)`,
        affectedDevices: suspiciousDevices.map(d => d.deviceName || d.id).slice(0, 20),
        detectedAt: new Date().toISOString(),
        recommendation: 'Review and potentially retire these devices from Intune management'
      });
    }

    return patterns;
  }

  /**
   * Calculate violation severity based on excess devices
   */
  private calculateViolationSeverity(
    deviceCount: number,
    limit: number
  ): 'Low' | 'Medium' | 'High' | 'Critical' {
    const excessPercent = ((deviceCount - limit) / limit) * 100;

    if (excessPercent >= 200) return 'Critical'; // 3x over limit
    if (excessPercent >= 100) return 'High';     // 2x over limit
    if (excessPercent >= 50) return 'Medium';    // 1.5x over limit
    return 'Low';
  }

  // ==========================================================================
  // Sorting Methods
  // ==========================================================================

  /**
   * Sort user device lists
   */
  private sortUserDeviceLists(
    lists: UserDeviceList[],
    filters: EnrollmentTrackingFilters
  ): UserDeviceList[] {
    const sortBy = filters.sortBy || 'deviceCount';
    const sortOrder = filters.sortOrder || 'desc';

    const sorted = [...lists].sort((a, b) => {
      let comparison = 0;

      switch (sortBy) {
        case 'deviceCount':
          comparison = a.totalDevices - b.totalDevices;
          break;
        case 'userName':
          comparison = a.displayName.localeCompare(b.displayName);
          break;
        case 'lastEnrollment':
          const aDate = a.lastEnrollmentDate ? new Date(a.lastEnrollmentDate).getTime() : 0;
          const bDate = b.lastEnrollmentDate ? new Date(b.lastEnrollmentDate).getTime() : 0;
          comparison = aDate - bDate;
          break;
        case 'department':
          comparison = (a.department || '').localeCompare(b.department || '');
          break;
        default:
          comparison = 0;
      }

      return sortOrder === 'asc' ? comparison : -comparison;
    });

    return sorted;
  }

  /**
   * Sort enrollment counts
   */
  private sortEnrollmentCounts(
    counts: UserEnrollmentCount[],
    filters: EnrollmentTrackingFilters
  ): UserEnrollmentCount[] {
    const sortBy = filters.sortBy || 'deviceCount';
    const sortOrder = filters.sortOrder || 'desc';

    const sorted = [...counts].sort((a, b) => {
      let comparison = 0;

      switch (sortBy) {
        case 'deviceCount':
          comparison = a.totalDeviceCount - b.totalDeviceCount;
          break;
        case 'userName':
          comparison = a.displayName.localeCompare(b.displayName);
          break;
        case 'lastEnrollment':
          const aDate = a.lastEnrollmentDate ? new Date(a.lastEnrollmentDate).getTime() : 0;
          const bDate = b.lastEnrollmentDate ? new Date(b.lastEnrollmentDate).getTime() : 0;
          comparison = aDate - bDate;
          break;
        case 'department':
          comparison = (a.department || '').localeCompare(b.department || '');
          break;
        default:
          comparison = 0;
      }

      return sortOrder === 'asc' ? comparison : -comparison;
    });

    return sorted;
  }

  // ==========================================================================
  // Utility Methods
  // ==========================================================================

  /**
   * Normalize platform name
   */
  private normalizePlatform(platform: string): string {
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

  /**
   * Normalize owner type
   */
  private normalizeOwnerType(ownerType: string): 'Corporate' | 'Personal' | 'Unknown' {
    if (!ownerType) return 'Unknown';

    const normalized = ownerType.toLowerCase();

    if (normalized.includes('company') || normalized.includes('corporate')) {
      return 'Corporate';
    } else if (normalized.includes('personal')) {
      return 'Personal';
    }

    return 'Unknown';
  }

  /**
   * Create empty platform breakdown
   */
  private createEmptyPlatformBreakdown(): PlatformBreakdown {
    return {
      iOS: 0,
      Android: 0,
      Windows: 0,
      macOS: 0,
      Linux: 0,
      Other: 0
    };
  }

  /**
   * Create empty enrollment trend
   */
  private createEmptyEnrollmentTrend(): EnrollmentTrend {
    return {
      last7Days: 0,
      last30Days: 0,
      last90Days: 0,
      last365Days: 0
    };
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

    // Header
    lines.push(reportData.metadata.reportName);
    lines.push(`Generated: ${reportData.metadata.generatedAt}`);
    lines.push(`Total Records: ${reportData.metadata.recordCount}`);
    lines.push('');

    // Determine report type
    const firstRecord = reportData.data[0];

    if (firstRecord && 'devices' in firstRecord) {
      // Report 17: Device List
      lines.push('User Principal Name,Display Name,Department,Job Title,Total Devices,iOS,Android,Windows,macOS,Linux,Other,Corporate,Personal,First Enrollment,Last Enrollment,Avg Device Age (days)');

      (reportData.data as UserDeviceList[]).forEach(user => {
        const platform = user.devicesByPlatform;
        const owner = user.devicesByOwnerType;

        lines.push([
          user.userPrincipalName,
          user.displayName,
          user.department || '',
          user.jobTitle || '',
          user.totalDevices,
          platform.iOS,
          platform.Android,
          platform.Windows,
          platform.macOS,
          platform.Linux,
          platform.Other,
          owner.Corporate,
          owner.Personal,
          user.firstEnrollmentDate || '',
          user.lastEnrollmentDate || '',
          user.averageDeviceAge
        ].join(','));
      });
    } else if (firstRecord && 'totalDeviceCount' in firstRecord) {
      // Report 30: Enrollment Count
      lines.push('User Principal Name,Display Name,Department,Job Title,Total Devices,iOS,Android,Windows,macOS,Linux,Other,Corporate,Personal,Compliant,Non-Compliant,Active,Inactive,Last Enrollment,First Enrollment');

      (reportData.data as UserEnrollmentCount[]).forEach(user => {
        lines.push([
          user.userPrincipalName,
          user.displayName,
          user.department || '',
          user.jobTitle || '',
          user.totalDeviceCount,
          user.iosDeviceCount,
          user.androidDeviceCount,
          user.windowsDeviceCount,
          user.macOsDeviceCount,
          user.linuxDeviceCount,
          user.otherDeviceCount,
          user.corporateDeviceCount,
          user.personalDeviceCount,
          user.compliantDeviceCount,
          user.nonCompliantDeviceCount,
          user.activeDeviceCount,
          user.inactiveDeviceCount,
          user.lastEnrollmentDate || '',
          user.firstEnrollmentDate || ''
        ].join(','));
      });
    }

    return lines.join('\n');
  }

  /**
   * Export report data as HTML
   */
  exportAsHTML(reportData: ReportData): string {
    const metadata = reportData.metadata;
    const firstRecord = reportData.data[0];

    let html = `
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${metadata.reportName}</title>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 20px;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 1400px;
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
      margin-bottom: 20px;
    }
    .metadata {
      background: #f9f9f9;
      padding: 15px;
      border-radius: 6px;
      margin-bottom: 20px;
      border-left: 4px solid #0078d4;
    }
    .metadata p {
      margin: 5px 0;
      color: #666;
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
      border-radius: 6px;
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
      font-size: 14px;
    }
    th {
      background: #0078d4;
      color: white;
      padding: 12px 8px;
      text-align: left;
      font-weight: 600;
      position: sticky;
      top: 0;
    }
    td {
      padding: 10px 8px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background-color: #f5f5f5;
    }
    .platform-badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: bold;
      margin-right: 5px;
    }
    .ios { background: #007aff; color: white; }
    .android { background: #3ddc84; color: black; }
    .windows { background: #0078d4; color: white; }
    .macos { background: #000; color: white; }
    .footer {
      margin-top: 30px;
      padding-top: 15px;
      border-top: 1px solid #ddd;
      color: #666;
      font-size: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>${metadata.reportName}</h1>

    <div class="metadata">
      <p><strong>Generated:</strong> ${new Date(metadata.generatedAt).toLocaleString()}</p>
      <p><strong>Total Records:</strong> ${metadata.recordCount}</p>
    </div>
`;

    // Add summary if available
    if (reportData.summary) {
      html += '<div class="summary">';
      Object.entries(reportData.summary).forEach(([key, value]) => {
        const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
        html += `
      <div class="summary-card">
        <h3>${label}</h3>
        <p>${typeof value === 'number' ? Math.round(value * 100) / 100 : value}</p>
      </div>`;
      });
      html += '</div>';
    }

    // Add table
    if (firstRecord && 'devices' in firstRecord) {
      // Report 17: Device List
      html += `
    <h2>Device Enrollment Details by User</h2>
    <table>
      <thead>
        <tr>
          <th>User</th>
          <th>Department</th>
          <th>Total Devices</th>
          <th>Platform Breakdown</th>
          <th>Owner Type</th>
          <th>Last Enrollment</th>
        </tr>
      </thead>
      <tbody>`;

      (reportData.data as UserDeviceList[]).forEach(user => {
        const platforms = Object.entries(user.devicesByPlatform)
          .filter(([_, count]) => count > 0)
          .map(([platform, count]) => `<span class="platform-badge ${platform.toLowerCase()}">${platform}: ${count}</span>`)
          .join(' ');

        html += `
        <tr>
          <td>
            <strong>${user.displayName}</strong><br>
            <small>${user.userPrincipalName}</small>
          </td>
          <td>${user.department || 'N/A'}</td>
          <td><strong>${user.totalDevices}</strong></td>
          <td>${platforms}</td>
          <td>Corp: ${user.devicesByOwnerType.Corporate}, Pers: ${user.devicesByOwnerType.Personal}</td>
          <td>${user.lastEnrollmentDate ? new Date(user.lastEnrollmentDate).toLocaleDateString() : 'N/A'}</td>
        </tr>`;
      });

      html += `
      </tbody>
    </table>`;
    } else if (firstRecord && 'totalDeviceCount' in firstRecord) {
      // Report 30: Enrollment Count
      html += `
    <h2>Device Enrollment Counts by User</h2>
    <table>
      <thead>
        <tr>
          <th>User</th>
          <th>Department</th>
          <th>Total</th>
          <th>iOS</th>
          <th>Android</th>
          <th>Windows</th>
          <th>macOS</th>
          <th>Corp</th>
          <th>Pers</th>
          <th>Compliant</th>
          <th>Active</th>
        </tr>
      </thead>
      <tbody>`;

      (reportData.data as UserEnrollmentCount[]).forEach(user => {
        html += `
        <tr>
          <td>
            <strong>${user.displayName}</strong><br>
            <small>${user.userPrincipalName}</small>
          </td>
          <td>${user.department || 'N/A'}</td>
          <td><strong>${user.totalDeviceCount}</strong></td>
          <td>${user.iosDeviceCount}</td>
          <td>${user.androidDeviceCount}</td>
          <td>${user.windowsDeviceCount}</td>
          <td>${user.macOsDeviceCount}</td>
          <td>${user.corporateDeviceCount}</td>
          <td>${user.personalDeviceCount}</td>
          <td>${user.compliantDeviceCount}</td>
          <td>${user.activeDeviceCount}</td>
        </tr>`;
      });

      html += `
      </tbody>
    </table>`;
    }

    html += `
    <div class="footer">
      <p>Intune Enrollment Tracking Report | Configuration Manager Style | Generated by Intune Reporting Dashboard</p>
    </div>
  </div>
</body>
</html>`;

    return html;
  }
}

// ============================================================================
// Export
// ============================================================================

export default IntuneEnrollmentTrackingReports;
