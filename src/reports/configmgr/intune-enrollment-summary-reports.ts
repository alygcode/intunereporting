/**
 * Intune Enrollment Summary Reports for Configuration Manager
 *
 * This module replicates Configuration Manager enrollment summary reports for Microsoft Intune.
 * It provides four main reports:
 * - Report 17: List of Devices enrolled per user in Microsoft Intune (All devices per user)
 * - Report 30: Number of devices enrolled per user in Microsoft Intune (User enrollment counts)
 * - Report 18: List of devices in a specific device category (Devices by category)
 * - Report 24: Mobile devices unmanaged (enrollment failed/assignment issues)
 *
 * Features:
 * - User-device relationship mapping
 * - Device category filtering and grouping
 * - Enrollment success rate tracking
 * - Failed enrollment and assignment detection
 * - Management state monitoring
 * - Certificate and enrollment validation
 * - Multiple export formats (JSON, CSV, HTML)
 * - Advanced filtering and sorting capabilities
 *
 * @module intune-enrollment-summary-reports
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
 * Device enrolled by a user with full details
 */
export interface UserDeviceEnrollment {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userId: string;
  displayName: string;
  platform: string;
  osVersion: string;
  model?: string;
  manufacturer?: string;
  enrollmentDate: string;
  enrollmentType: string;
  managementAgent: string;
  managementState: string;
  complianceState: string;
  ownerType: string;
  lastSyncDateTime: string;
  deviceCategory?: string;
  serialNumber?: string;
  imei?: string;
  isSupervised?: boolean;
  enrollmentProfileName?: string;
  azureADDeviceId?: string;
  deviceActionResults?: DeviceActionResult[];
  enrollmentStatus: 'Enrolled' | 'Pending' | 'Failed' | 'Unmanaged';
  assignmentStatus?: 'Assigned' | 'NotAssigned' | 'Failed';
  certificateIssued?: boolean;
  daysSinceEnrollment: number;
  daysSinceLastSync: number;
}

/**
 * Device action result
 */
export interface DeviceActionResult {
  actionName: string;
  actionState: string;
  startDateTime: string;
  lastUpdatedDateTime: string;
}

/**
 * User enrollment statistics and counts
 */
export interface UserEnrollmentStatistics {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  department?: string;
  jobTitle?: string;
  email?: string;
  totalDeviceCount: number;
  enrolledDeviceCount: number;
  pendingDeviceCount: number;
  failedDeviceCount: number;
  unmanagedDeviceCount: number;

  // Platform breakdown
  iosCount: number;
  androidCount: number;
  windowsCount: number;
  macOsCount: number;
  linuxCount: number;
  otherCount: number;

  // Ownership breakdown
  corporateCount: number;
  personalCount: number;

  // Compliance breakdown
  compliantCount: number;
  nonCompliantCount: number;
  unknownComplianceCount: number;

  // Activity breakdown
  activeDeviceCount: number; // Synced in last 30 days
  inactiveDeviceCount: number;

  // Enrollment metrics
  enrollmentSuccessRate: number; // Percentage
  averageEnrollmentAge: number; // Days
  oldestEnrollmentDate?: string;
  newestEnrollmentDate?: string;

  // Subscription status
  hasIntuneSubscription: boolean;
  subscriptionEnabled: boolean;

  // User status
  accountEnabled: boolean;
  createdDateTime?: string;
  lastSignInDateTime?: string;
}

/**
 * Device category information
 */
export interface DeviceCategory {
  id: string;
  displayName: string;
  description?: string;
  deviceCount: number;
  devices?: CategoryDeviceInfo[];
  platforms: PlatformBreakdown;
  ownerTypes: OwnerTypeBreakdown;
  complianceStates: ComplianceBreakdown;
  enrollmentTrend: EnrollmentTrendData;
}

/**
 * Device information in a category
 */
export interface CategoryDeviceInfo {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  userDisplayName: string;
  platform: string;
  osVersion: string;
  model?: string;
  enrollmentDate: string;
  managementState: string;
  complianceState: string;
  lastSyncDateTime: string;
  ownerType: string;
}

/**
 * Unmanaged device information (enrollment/assignment failures)
 */
export interface UnmanagedDevice {
  deviceId?: string;
  deviceName?: string;
  userId?: string;
  userPrincipalName: string;
  userDisplayName?: string;
  platform?: string;
  osVersion?: string;
  serialNumber?: string;

  // Failure information
  failureType: 'EnrollmentFailed' | 'AssignmentFailed' | 'CertificateFailed' | 'ManagementLost' | 'Unknown';
  failureReason: string;
  failureCategory: string;
  errorCode?: string;
  errorMessage?: string;

  // Status information
  managementState: string;
  enrollmentState?: string;
  certificateIssued: boolean;
  certificateExpired?: boolean;

  // Timing information
  enrollmentAttemptDate?: string;
  lastContactDate?: string;
  daysSinceFailure?: number;

  // Troubleshooting
  troubleshootingSteps: string[];
  troubleshootingEventId?: string;
  relatedEventCount?: number;

  // Recovery
  isRecoverable: boolean;
  recoveryActions: string[];
  priorityLevel: 'High' | 'Medium' | 'Low';
}

/**
 * Platform breakdown
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
 * Owner type breakdown
 */
export interface OwnerTypeBreakdown {
  Corporate: number;
  Personal: number;
  Unknown: number;
}

/**
 * Compliance state breakdown
 */
export interface ComplianceBreakdown {
  Compliant: number;
  NonCompliant: number;
  Unknown: number;
  InGracePeriod: number;
  Error: number;
}

/**
 * Enrollment trend data
 */
export interface EnrollmentTrendData {
  last7Days: number;
  last30Days: number;
  last90Days: number;
  last365Days: number;
  total: number;
}

/**
 * Enrollment success rate metrics
 */
export interface EnrollmentSuccessMetrics {
  totalAttempts: number;
  successfulEnrollments: number;
  failedEnrollments: number;
  pendingEnrollments: number;
  successRate: number; // Percentage
  failureRate: number; // Percentage

  // Breakdown by platform
  byPlatform: {
    [platform: string]: {
      attempts: number;
      successful: number;
      failed: number;
      successRate: number;
    };
  };

  // Breakdown by enrollment method
  byEnrollmentMethod: {
    [method: string]: {
      attempts: number;
      successful: number;
      failed: number;
      successRate: number;
    };
  };

  // Failure categories
  failureCategories: {
    category: string;
    count: number;
    percentage: number;
  }[];

  // Trends
  trendsLast30Days: {
    date: string;
    attempts: number;
    successful: number;
    failed: number;
    successRate: number;
  }[];

  generatedAt: string;
}

/**
 * Filter options for enrollment summary reports
 */
export interface EnrollmentSummaryFilters {
  userId?: string;
  userPrincipalName?: string;
  deviceCategory?: string;
  platform?: string[];
  ownerType?: 'Corporate' | 'Personal' | 'All';
  complianceState?: 'Compliant' | 'NonCompliant' | 'All';
  managementState?: string[];
  enrollmentStatus?: ('Enrolled' | 'Pending' | 'Failed' | 'Unmanaged')[];
  startDate?: Date;
  endDate?: Date;
  minDeviceCount?: number;
  maxDeviceCount?: number;
  includeInactiveDevices?: boolean;
  sortBy?: 'deviceCount' | 'userName' | 'enrollmentDate' | 'platform' | 'category';
  sortOrder?: 'asc' | 'desc';
  top?: number;
}

/**
 * Report options
 */
export interface EnrollmentSummaryReportOptions {
  filters?: EnrollmentSummaryFilters;
  includeDeviceDetails?: boolean;
  includeUserDetails?: boolean;
  includeTroubleshooting?: boolean;
  includeSuccessMetrics?: boolean;
}

// ============================================================================
// Main Report Class
// ============================================================================

/**
 * Intune Enrollment Summary Reports Class
 *
 * Provides Configuration Manager-style enrollment summary reports for Intune.
 * Implements four main reports:
 * - Report 17: List of devices enrolled per user
 * - Report 30: Number of devices enrolled per user
 * - Report 18: List of devices in a specific category
 * - Report 24: Mobile devices unmanaged (enrollment/assignment failures)
 *
 * @example
 * ```typescript
 * const report = new IntuneEnrollmentSummaryReports(graphClient, config);
 *
 * // Get devices enrolled per user
 * const userDevices = await report.getDevicesEnrolledPerUser('user-id');
 *
 * // Get enrollment counts for all users
 * const enrollmentCounts = await report.getUserEnrollmentCounts();
 *
 * // Get devices by category
 * const categoryDevices = await report.getDevicesByCategory('Laptops');
 *
 * // Get unmanaged devices
 * const unmanagedDevices = await report.getUnmanagedDevices();
 * ```
 */
export class IntuneEnrollmentSummaryReports extends BaseReport {
  name = 'Intune Enrollment Summary Reports';
  description = 'Configuration Manager-style enrollment summary reports for Microsoft Intune';
  category = 'Configuration Manager Reports';
  enabled = true;

  /**
   * Execute the default enrollment summary report
   */
  async execute(options: EnrollmentSummaryReportOptions = {}): Promise<ReportData> {
    logger.info('Starting Intune Enrollment Summary Reports generation');

    try {
      const [enrollmentCounts, successMetrics] = await Promise.all([
        this.getUserEnrollmentCounts(options),
        options.includeSuccessMetrics !== false
          ? this.calculateEnrollmentSuccessRates()
          : Promise.resolve(null)
      ]);

      return {
        metadata: this.createMetadata(
          'Intune Enrollment Summary Overview',
          enrollmentCounts.data.length,
          { options, includesSuccessMetrics: options.includeSuccessMetrics !== false }
        ),
        data: enrollmentCounts.data,
        summary: {
          ...enrollmentCounts.summary,
          successMetrics
        }
      };
    } catch (error) {
      logger.error('Error generating Intune Enrollment Summary Reports', error);
      throw error;
    }
  }

  // ==========================================================================
  // Report 17: List of Devices Enrolled Per User
  // ==========================================================================

  /**
   * Report 17: Get devices enrolled per user
   *
   * Returns a detailed list of all devices enrolled by a specific user or all users.
   * Includes device details, enrollment information, and management status.
   *
   * @param userId - Optional user ID to filter by specific user
   * @param options - Report configuration options
   * @returns Detailed device enrollment list
   *
   * @example
   * ```typescript
   * // Get all devices for a specific user
   * const userDevices = await report.getDevicesEnrolledPerUser('user-id-123');
   *
   * // Get all devices for all users with filtering
   * const allDevices = await report.getDevicesEnrolledPerUser(undefined, {
   *   filters: { platform: ['iOS', 'Android'], complianceState: 'Compliant' }
   * });
   *
   * // Get devices enrolled in last 30 days
   * const recentDevices = await report.getDevicesEnrolledPerUser(undefined, {
   *   filters: {
   *     startDate: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000),
   *     endDate: new Date()
   *   }
   * });
   * ```
   */
  async getDevicesEnrolledPerUser(
    userId?: string,
    options: EnrollmentSummaryReportOptions = {}
  ): Promise<ReportData> {
    logger.info(`Generating Report 17: Devices Enrolled Per User${userId ? ` (User: ${userId})` : ''}`);

    try {
      const filters = options.filters || {};
      if (userId) {
        filters.userId = userId;
      }

      // Fetch devices with enrollment information
      const devices = await this.fetchDevicesWithEnrollmentInfo(filters);

      // Enrich with user details if requested
      const enrichedDevices = options.includeUserDetails !== false
        ? await this.enrichDevicesWithUserDetails(devices)
        : devices;

      // Sort and limit results
      const sortedDevices = this.sortDevices(enrichedDevices, filters);
      const finalDevices = filters.top ? sortedDevices.slice(0, filters.top) : sortedDevices;

      // Calculate summary statistics
      const summary = this.calculateDeviceListSummary(finalDevices);

      logger.info(`Report 17 complete: ${finalDevices.length} devices found`);

      return {
        metadata: this.createMetadata(
          'List of Devices Enrolled Per User in Microsoft Intune',
          finalDevices.length,
          { userId, filters, includeUserDetails: options.includeUserDetails }
        ),
        data: finalDevices,
        summary
      };
    } catch (error) {
      logger.error('Error generating devices enrolled per user report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Report 30: Number of Devices Enrolled Per User
  // ==========================================================================

  /**
   * Report 30: Get enrollment count per user
   *
   * Returns enrollment statistics and device counts per user,
   * broken down by platform, owner type, compliance state, and activity.
   *
   * @param options - Report configuration options
   * @returns User enrollment statistics
   *
   * @example
   * ```typescript
   * // Get enrollment counts for all users
   * const counts = await report.getUserEnrollmentCounts();
   *
   * // Filter users with multiple devices
   * const multiDeviceUsers = await report.getUserEnrollmentCounts({
   *   filters: { minDeviceCount: 2 }
   * });
   *
   * // Get counts for specific department
   * const deptCounts = await report.getUserEnrollmentCounts({
   *   filters: { userPrincipalName: '*@contoso.com' }
   * });
   *
   * // Include subscription status
   * const countsWithSubscription = await report.getUserEnrollmentCounts({
   *   includeUserDetails: true
   * });
   * ```
   */
  async getUserEnrollmentCounts(
    options: EnrollmentSummaryReportOptions = {}
  ): Promise<ReportData> {
    logger.info('Generating Report 30: Number of Devices Enrolled Per User');

    try {
      const filters = options.filters || {};

      // Fetch all managed devices
      const allDevices = await this.fetchAllManagedDevices();

      // Apply filters
      const filteredDevices = this.applyDeviceFilters(allDevices, filters);

      // Group by user
      const userDeviceMap = this.groupDevicesByUser(filteredDevices);

      // Build enrollment statistics per user
      const userStatistics = await this.buildUserEnrollmentStatistics(
        userDeviceMap,
        options.includeUserDetails !== false
      );

      // Apply user-level filters
      const filteredStatistics = this.applyUserFilters(userStatistics, filters);

      // Sort results
      const sortedStatistics = this.sortUserStatistics(filteredStatistics, filters);

      // Apply top limit
      const finalStatistics = filters.top
        ? sortedStatistics.slice(0, filters.top)
        : sortedStatistics;

      // Calculate overall summary
      const summary = this.calculateEnrollmentCountSummary(finalStatistics);

      logger.info(`Report 30 complete: ${finalStatistics.length} users with device counts`);

      return {
        metadata: this.createMetadata(
          'Number of Devices Enrolled Per User in Microsoft Intune',
          finalStatistics.length,
          { filters, includeUserDetails: options.includeUserDetails }
        ),
        data: finalStatistics,
        summary
      };
    } catch (error) {
      logger.error('Error generating enrollment count per user report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Report 18: List of Devices in a Specific Device Category
  // ==========================================================================

  /**
   * Report 18: Get devices by category
   *
   * Returns all devices in a specific device category or all categories.
   * Includes category statistics and device breakdowns.
   *
   * @param categoryName - Optional category name to filter
   * @param options - Report configuration options
   * @returns Devices grouped by category
   *
   * @example
   * ```typescript
   * // Get all devices in "Laptops" category
   * const laptops = await report.getDevicesByCategory('Laptops');
   *
   * // Get all categories with devices
   * const allCategories = await report.getDevicesByCategory();
   *
   * // Get category with device details
   * const categoryDetails = await report.getDevicesByCategory('Tablets', {
   *   includeDeviceDetails: true
   * });
   *
   * // Filter by platform within category
   * const iosLaptops = await report.getDevicesByCategory('Laptops', {
   *   filters: { platform: ['iOS'] }
   * });
   * ```
   */
  async getDevicesByCategory(
    categoryName?: string,
    options: EnrollmentSummaryReportOptions = {}
  ): Promise<ReportData> {
    logger.info(`Generating Report 18: Devices by Category${categoryName ? ` (Category: ${categoryName})` : ''}`);

    try {
      // Fetch device categories from Intune
      const categories = await this.fetchDeviceCategories();

      // Fetch all managed devices
      const allDevices = await this.fetchAllManagedDevices();

      // Build category report
      const categoryReports: DeviceCategory[] = [];

      for (const category of categories) {
        // Skip if filtering by specific category
        if (categoryName && category.displayName !== categoryName) {
          continue;
        }

        // Get devices in this category
        const categoryDevices = allDevices.filter(d =>
          d.deviceCategoryDisplayName === category.displayName
        );

        // Apply additional filters
        const filteredDevices = options.filters
          ? this.applyDeviceFilters(categoryDevices, options.filters)
          : categoryDevices;

        // Build device info list if requested
        const deviceInfoList = options.includeDeviceDetails !== false
          ? filteredDevices.map(d => this.mapToCategoryDeviceInfo(d))
          : [];

        // Calculate category statistics
        const platformBreakdown = this.calculatePlatformBreakdown(filteredDevices);
        const ownerTypeBreakdown = this.calculateOwnerTypeBreakdown(filteredDevices);
        const complianceBreakdown = this.calculateComplianceBreakdown(filteredDevices);
        const enrollmentTrend = this.calculateEnrollmentTrend(filteredDevices);

        categoryReports.push({
          id: category.id,
          displayName: category.displayName,
          description: category.description,
          deviceCount: filteredDevices.length,
          devices: deviceInfoList,
          platforms: platformBreakdown,
          ownerTypes: ownerTypeBreakdown,
          complianceStates: complianceBreakdown,
          enrollmentTrend
        });
      }

      // Handle uncategorized devices
      const uncategorizedDevices = allDevices.filter(d =>
        !d.deviceCategoryDisplayName || d.deviceCategoryDisplayName === ''
      );

      if (uncategorizedDevices.length > 0 && (!categoryName || categoryName === 'Uncategorized')) {
        const filteredUncategorized = options.filters
          ? this.applyDeviceFilters(uncategorizedDevices, options.filters)
          : uncategorizedDevices;

        categoryReports.push({
          id: 'uncategorized',
          displayName: 'Uncategorized',
          description: 'Devices without an assigned category',
          deviceCount: filteredUncategorized.length,
          devices: options.includeDeviceDetails !== false
            ? filteredUncategorized.map(d => this.mapToCategoryDeviceInfo(d))
            : [],
          platforms: this.calculatePlatformBreakdown(filteredUncategorized),
          ownerTypes: this.calculateOwnerTypeBreakdown(filteredUncategorized),
          complianceStates: this.calculateComplianceBreakdown(filteredUncategorized),
          enrollmentTrend: this.calculateEnrollmentTrend(filteredUncategorized)
        });
      }

      // Sort by device count
      categoryReports.sort((a, b) => b.deviceCount - a.deviceCount);

      // Calculate summary
      const summary = {
        totalCategories: categoryReports.length,
        totalDevices: categoryReports.reduce((sum, c) => sum + c.deviceCount, 0),
        averageDevicesPerCategory: categoryReports.length > 0
          ? Math.round((categoryReports.reduce((sum, c) => sum + c.deviceCount, 0) / categoryReports.length) * 100) / 100
          : 0,
        largestCategory: categoryReports[0]?.displayName || 'N/A',
        largestCategoryCount: categoryReports[0]?.deviceCount || 0
      };

      logger.info(`Report 18 complete: ${categoryReports.length} categories with ${summary.totalDevices} devices`);

      return {
        metadata: this.createMetadata(
          'List of Devices in Specific Device Category',
          categoryReports.length,
          { categoryName, includeDeviceDetails: options.includeDeviceDetails }
        ),
        data: categoryReports,
        summary
      };
    } catch (error) {
      logger.error('Error generating devices by category report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Report 24: Mobile Devices Unmanaged (Enrollment/Assignment Failures)
  // ==========================================================================

  /**
   * Report 24: Get unmanaged devices
   *
   * Returns devices with enrollment failures, assignment failures,
   * management state issues, or certificate problems.
   *
   * @param options - Report configuration options
   * @returns Unmanaged devices with troubleshooting information
   *
   * @example
   * ```typescript
   * // Get all unmanaged devices
   * const unmanagedDevices = await report.getUnmanagedDevices();
   *
   * // Get devices with enrollment failures
   * const enrollmentFailures = await report.getUnmanagedDevices({
   *   filters: { enrollmentStatus: ['Failed'] }
   * });
   *
   * // Get devices with certificate issues
   * const certIssues = await report.getUnmanagedDevices({
   *   includeTroubleshooting: true
   * });
   *
   * // Get high-priority failures
   * const highPriorityIssues = await report.getUnmanagedDevices({
   *   filters: { platform: ['Windows', 'iOS'] }
   * });
   * ```
   */
  async getUnmanagedDevices(
    options: EnrollmentSummaryReportOptions = {}
  ): Promise<ReportData> {
    logger.info('Generating Report 24: Mobile Devices Unmanaged (Enrollment/Assignment Failures)');

    try {
      // Fetch unmanaged devices from multiple sources
      const [
        enrollmentFailures,
        assignmentFailures,
        managementStateIssues,
        certificateIssues
      ] = await Promise.all([
        this.fetchEnrollmentFailures(options.filters),
        this.fetchAssignmentFailures(options.filters),
        this.fetchManagementStateIssues(options.filters),
        this.fetchCertificateIssues(options.filters)
      ]);

      // Combine and deduplicate
      const allUnmanagedDevices = this.deduplicateUnmanagedDevices([
        ...enrollmentFailures,
        ...assignmentFailures,
        ...managementStateIssues,
        ...certificateIssues
      ]);

      // Add troubleshooting information if requested
      const enrichedDevices = options.includeTroubleshooting !== false
        ? this.enrichWithTroubleshooting(allUnmanagedDevices)
        : allUnmanagedDevices;

      // Sort by priority and date
      const sortedDevices = this.sortUnmanagedDevices(enrichedDevices);

      // Apply top limit if specified
      const finalDevices = options.filters?.top
        ? sortedDevices.slice(0, options.filters.top)
        : sortedDevices;

      // Calculate summary
      const summary = this.calculateUnmanagedDevicesSummary(finalDevices);

      logger.info(`Report 24 complete: ${finalDevices.length} unmanaged devices found`);

      return {
        metadata: this.createMetadata(
          'Mobile Devices Unmanaged (Enrollment/Assignment Failures)',
          finalDevices.length,
          { includeTroubleshooting: options.includeTroubleshooting }
        ),
        data: finalDevices,
        summary
      };
    } catch (error) {
      logger.error('Error generating unmanaged devices report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Enrollment Success Rate Tracking
  // ==========================================================================

  /**
   * Calculate enrollment success rates
   *
   * Analyzes enrollment attempts and calculates success rates
   * across platforms, methods, and time periods.
   *
   * @returns Comprehensive enrollment success metrics
   *
   * @example
   * ```typescript
   * const successMetrics = await report.calculateEnrollmentSuccessRates();
   * console.log(`Overall success rate: ${successMetrics.successRate}%`);
   * console.log(`Failed enrollments: ${successMetrics.failedEnrollments}`);
   * ```
   */
  async calculateEnrollmentSuccessRates(): Promise<EnrollmentSuccessMetrics> {
    logger.info('Calculating enrollment success rates');

    try {
      // Fetch all enrollment data
      const [managedDevices, troubleshootingEvents] = await Promise.all([
        this.fetchAllManagedDevices(),
        this.fetchTroubleshootingEvents()
      ]);

      // Count successful enrollments
      const successfulEnrollments = managedDevices.length;

      // Count failed enrollments from troubleshooting events
      const enrollmentFailureEvents = troubleshootingEvents.filter(e =>
        e.eventName?.toLowerCase().includes('enrollment') ||
        e.eventName?.toLowerCase().includes('enroll')
      );

      const failedEnrollments = enrollmentFailureEvents.length;
      const totalAttempts = successfulEnrollments + failedEnrollments;
      const successRate = totalAttempts > 0
        ? Math.round((successfulEnrollments / totalAttempts) * 10000) / 100
        : 0;
      const failureRate = 100 - successRate;

      // Calculate by platform
      const byPlatform: any = {};
      const platformCounts = new Map<string, { successful: number; failed: number }>();

      managedDevices.forEach(device => {
        const platform = this.normalizePlatform(device.operatingSystem || 'Other');
        const current = platformCounts.get(platform) || { successful: 0, failed: 0 };
        current.successful++;
        platformCounts.set(platform, current);
      });

      enrollmentFailureEvents.forEach(event => {
        const platform = this.normalizePlatform(event.platform || 'Other');
        const current = platformCounts.get(platform) || { successful: 0, failed: 0 };
        current.failed++;
        platformCounts.set(platform, current);
      });

      platformCounts.forEach((counts, platform) => {
        const attempts = counts.successful + counts.failed;
        byPlatform[platform] = {
          attempts,
          successful: counts.successful,
          failed: counts.failed,
          successRate: attempts > 0 ? Math.round((counts.successful / attempts) * 10000) / 100 : 0
        };
      });

      // Calculate by enrollment method
      const byEnrollmentMethod: any = {};
      const methodCounts = new Map<string, { successful: number; failed: number }>();

      managedDevices.forEach(device => {
        const method = device.enrollmentType || device.deviceEnrollmentType || 'Unknown';
        const current = methodCounts.get(method) || { successful: 0, failed: 0 };
        current.successful++;
        methodCounts.set(method, current);
      });

      methodCounts.forEach((counts, method) => {
        const attempts = counts.successful + counts.failed;
        byEnrollmentMethod[method] = {
          attempts,
          successful: counts.successful,
          failed: counts.failed,
          successRate: attempts > 0 ? Math.round((counts.successful / attempts) * 10000) / 100 : 0
        };
      });

      // Categorize failures
      const failureCategoryMap = new Map<string, number>();
      enrollmentFailureEvents.forEach(event => {
        const category = this.categorizeFailure(event.correlationId || event.failureReason || 'Unknown');
        failureCategoryMap.set(category, (failureCategoryMap.get(category) || 0) + 1);
      });

      const failureCategories = Array.from(failureCategoryMap.entries())
        .map(([category, count]) => ({
          category,
          count,
          percentage: failedEnrollments > 0
            ? Math.round((count / failedEnrollments) * 10000) / 100
            : 0
        }))
        .sort((a, b) => b.count - a.count);

      // Calculate trends for last 30 days
      const trendsLast30Days = this.calculateEnrollmentTrends(
        managedDevices,
        enrollmentFailureEvents,
        30
      );

      return {
        totalAttempts,
        successfulEnrollments,
        failedEnrollments,
        pendingEnrollments: 0, // Would need additional API data
        successRate,
        failureRate,
        byPlatform,
        byEnrollmentMethod,
        failureCategories,
        trendsLast30Days,
        generatedAt: new Date().toISOString()
      };
    } catch (error) {
      logger.error('Error calculating enrollment success rates', error);
      throw error;
    }
  }

  // ==========================================================================
  // Data Fetching Methods
  // ==========================================================================

  /**
   * Fetch all managed devices with full details
   */
  private async fetchAllManagedDevices(): Promise<any[]> {
    logger.info('Fetching all managed devices');

    return this.retryGraphCall(async () => {
      const response = await this.graphClient
        .api('/deviceManagement/managedDevices')
        .select([
          'id',
          'deviceName',
          'userId',
          'userPrincipalName',
          'userDisplayName',
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
          'enrollmentType',
          'deviceEnrollmentType',
          'azureADDeviceId',
          'isSupervised',
          'deviceCategoryDisplayName',
          'enrollmentProfileName',
          'emailAddress',
          'deviceActionResults'
        ].join(','))
        .top(999)
        .get();

      return this.getAllPages<any>(response);
    });
  }

  /**
   * Fetch devices with enrollment information for specific user
   */
  private async fetchUserDevices(userId: string): Promise<any[]> {
    logger.info(`Fetching devices for user: ${userId}`);

    try {
      // Try using ownedDevices endpoint
      const ownedResponse = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${userId}/ownedDevices`)
          .get();
      });

      // Try using registeredDevices endpoint
      const registeredResponse = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${userId}/registeredDevices`)
          .get();
      });

      // Try using managedDevices filter
      const managedResponse = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api('/deviceManagement/managedDevices')
          .filter(`userId eq '${userId}'`)
          .select([
            'id',
            'deviceName',
            'userId',
            'userPrincipalName',
            'operatingSystem',
            'enrolledDateTime',
            'managementState',
            'complianceState'
          ].join(','))
          .get();
      });

      // Combine and deduplicate
      const allDevices = [
        ...(ownedResponse.value || []),
        ...(registeredResponse.value || []),
        ...(managedResponse.value || [])
      ];

      const uniqueDevices = this.deduplicateDevices(allDevices);
      return uniqueDevices;
    } catch (error) {
      logger.warn(`Unable to fetch devices for user ${userId}`, error);
      return [];
    }
  }

  /**
   * Fetch devices with enrollment info applying filters
   */
  private async fetchDevicesWithEnrollmentInfo(
    filters: EnrollmentSummaryFilters
  ): Promise<UserDeviceEnrollment[]> {
    logger.info('Fetching devices with enrollment information');

    let devices: any[];

    if (filters.userId) {
      // Fetch specific user's devices
      devices = await this.fetchUserDevices(filters.userId);
    } else if (filters.userPrincipalName) {
      // Filter by UPN
      const allDevices = await this.fetchAllManagedDevices();
      devices = allDevices.filter(d =>
        d.userPrincipalName?.toLowerCase().includes(filters.userPrincipalName!.toLowerCase())
      );
    } else {
      // Fetch all devices
      devices = await this.fetchAllManagedDevices();
    }

    // Apply additional filters
    devices = this.applyDeviceFilters(devices, filters);

    // Map to UserDeviceEnrollment format
    return devices.map(d => this.mapToUserDeviceEnrollment(d));
  }

  /**
   * Fetch device categories
   */
  private async fetchDeviceCategories(): Promise<any[]> {
    logger.info('Fetching device categories');

    try {
      return await this.retryGraphCall(async () => {
        const response = await this.graphClient
          .api('/deviceManagement/deviceCategories')
          .select(['id', 'displayName', 'description'])
          .get();

        return response.value || [];
      });
    } catch (error) {
      logger.warn('Unable to fetch device categories', error);
      return [];
    }
  }

  /**
   * Fetch enrollment failures from troubleshooting events
   */
  private async fetchEnrollmentFailures(
    filters?: EnrollmentSummaryFilters
  ): Promise<UnmanagedDevice[]> {
    logger.info('Fetching enrollment failures');

    try {
      const events = await this.fetchTroubleshootingEvents();

      const enrollmentFailures = events.filter(e =>
        e.eventName?.toLowerCase().includes('enrollment') &&
        (e.failureReason || e.additionalInformation)
      );

      return enrollmentFailures.map(event => this.mapToUnmanagedDevice(event, 'EnrollmentFailed'));
    } catch (error) {
      logger.warn('Unable to fetch enrollment failures', error);
      return [];
    }
  }

  /**
   * Fetch troubleshooting events
   */
  private async fetchTroubleshootingEvents(): Promise<any[]> {
    logger.info('Fetching troubleshooting events');

    try {
      return await this.retryGraphCall(async () => {
        const response = await this.graphClient
          .api('/deviceManagement/troubleshootingEvents')
          .top(999)
          .get();

        return this.getAllPages<any>(response);
      });
    } catch (error) {
      logger.warn('Unable to fetch troubleshooting events', error);
      return [];
    }
  }

  /**
   * Fetch assignment failures
   */
  private async fetchAssignmentFailures(
    filters?: EnrollmentSummaryFilters
  ): Promise<UnmanagedDevice[]> {
    logger.info('Fetching assignment failures');

    try {
      // Get devices with completed enrollment but failed assignments
      const allDevices = await this.fetchAllManagedDevices();

      const assignmentFailures = allDevices.filter(device => {
        // Check if device has management state issues
        const mgmtState = device.managementState?.toLowerCase() || '';
        return mgmtState.includes('failed') ||
               mgmtState.includes('error') ||
               (device.enrolledDateTime && !device.lastSyncDateTime);
      });

      return assignmentFailures.map(device =>
        this.mapDeviceToUnmanagedDevice(device, 'AssignmentFailed')
      );
    } catch (error) {
      logger.warn('Unable to fetch assignment failures', error);
      return [];
    }
  }

  /**
   * Fetch management state issues
   */
  private async fetchManagementStateIssues(
    filters?: EnrollmentSummaryFilters
  ): Promise<UnmanagedDevice[]> {
    logger.info('Fetching management state issues');

    try {
      const allDevices = await this.fetchAllManagedDevices();

      const managementIssues = allDevices.filter(device => {
        const mgmtState = device.managementState?.toLowerCase() || '';
        return mgmtState === 'unmanaged' ||
               mgmtState === 'retired' ||
               mgmtState === 'deleted' ||
               mgmtState === 'orphaned';
      });

      return managementIssues.map(device =>
        this.mapDeviceToUnmanagedDevice(device, 'ManagementLost')
      );
    } catch (error) {
      logger.warn('Unable to fetch management state issues', error);
      return [];
    }
  }

  /**
   * Fetch certificate issues
   */
  private async fetchCertificateIssues(
    filters?: EnrollmentSummaryFilters
  ): Promise<UnmanagedDevice[]> {
    logger.info('Fetching certificate issues');

    try {
      const events = await this.fetchTroubleshootingEvents();

      const certificateFailures = events.filter(e =>
        e.eventName?.toLowerCase().includes('certificate') ||
        e.failureReason?.toLowerCase().includes('certificate') ||
        e.correlationId?.toLowerCase().includes('0xcaa')
      );

      return certificateFailures.map(event =>
        this.mapToUnmanagedDevice(event, 'CertificateFailed')
      );
    } catch (error) {
      logger.warn('Unable to fetch certificate issues', error);
      return [];
    }
  }

  /**
   * Fetch user details
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
            'accountEnabled',
            'createdDateTime',
            'signInActivity'
          ].join(','))
          .get();
      });
    } catch (error) {
      logger.warn(`Unable to fetch user details for ${userId}`, error);
      return null;
    }
  }

  /**
   * Check if user has Intune subscription
   */
  private async checkUserIntuneSubscription(userId: string): Promise<boolean> {
    try {
      const licenses = await this.retryGraphCall(async () => {
        return await this.graphClient
          .api(`/users/${userId}/licenseDetails`)
          .get();
      });

      // Check for Intune or EMS licenses
      const hasIntuneLicense = (licenses.value || []).some((license: any) =>
        license.skuPartNumber?.includes('INTUNE') ||
        license.skuPartNumber?.includes('EMS') ||
        license.skuPartNumber?.includes('AAD_PREMIUM')
      );

      return hasIntuneLicense;
    } catch (error) {
      logger.warn(`Unable to check Intune subscription for user ${userId}`, error);
      return false;
    }
  }

  // ==========================================================================
  // Data Processing Methods
  // ==========================================================================

  /**
   * Apply device filters
   */
  private applyDeviceFilters(
    devices: any[],
    filters: EnrollmentSummaryFilters
  ): any[] {
    let filtered = [...devices];

    // Filter by platform
    if (filters.platform && filters.platform.length > 0) {
      filtered = filtered.filter(d => {
        const platform = this.normalizePlatform(d.operatingSystem || '');
        return filters.platform!.includes(platform);
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
        } else {
          return d.complianceState?.toLowerCase() !== 'compliant';
        }
      });
    }

    // Filter by management state
    if (filters.managementState && filters.managementState.length > 0) {
      filtered = filtered.filter(d =>
        filters.managementState!.includes(d.managementState)
      );
    }

    // Filter by device category
    if (filters.deviceCategory) {
      filtered = filtered.filter(d =>
        d.deviceCategoryDisplayName === filters.deviceCategory
      );
    }

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

    // Filter inactive devices
    if (!filters.includeInactiveDevices) {
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

      filtered = filtered.filter(d =>
        d.lastSyncDateTime && new Date(d.lastSyncDateTime) >= thirtyDaysAgo
      );
    }

    return filtered;
  }

  /**
   * Apply user-level filters
   */
  private applyUserFilters(
    users: UserEnrollmentStatistics[],
    filters: EnrollmentSummaryFilters
  ): UserEnrollmentStatistics[] {
    let filtered = [...users];

    // Filter by device count
    if (filters.minDeviceCount) {
      filtered = filtered.filter(u => u.totalDeviceCount >= filters.minDeviceCount!);
    }

    if (filters.maxDeviceCount) {
      filtered = filtered.filter(u => u.totalDeviceCount <= filters.maxDeviceCount!);
    }

    return filtered;
  }

  /**
   * Group devices by user
   */
  private groupDevicesByUser(devices: any[]): Map<string, any[]> {
    const userDeviceMap = new Map<string, any[]>();

    devices.forEach(device => {
      const userId = device.userId || 'Unknown';

      if (!userDeviceMap.has(userId)) {
        userDeviceMap.set(userId, []);
      }

      userDeviceMap.get(userId)!.push(device);
    });

    return userDeviceMap;
  }

  /**
   * Build user enrollment statistics
   */
  private async buildUserEnrollmentStatistics(
    userDeviceMap: Map<string, any[]>,
    includeUserDetails: boolean
  ): Promise<UserEnrollmentStatistics[]> {
    const statistics: UserEnrollmentStatistics[] = [];

    for (const [userId, devices] of userDeviceMap.entries()) {
      if (userId === 'Unknown') continue;

      const firstDevice = devices[0];

      // Fetch user details if requested
      let userDetails: any = null;
      let hasIntuneSubscription = false;

      if (includeUserDetails) {
        userDetails = await this.fetchUserDetails(userId);
        hasIntuneSubscription = await this.checkUserIntuneSubscription(userId);
      }

      // Calculate platform counts
      const platformCounts = this.calculatePlatformBreakdown(devices);

      // Calculate ownership counts
      const ownershipCounts = this.calculateOwnerTypeBreakdown(devices);

      // Calculate compliance counts
      const complianceCounts = this.calculateComplianceBreakdown(devices);

      // Calculate active/inactive counts
      const thirtyDaysAgo = new Date();
      thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);

      const activeCount = devices.filter(d =>
        d.lastSyncDateTime && new Date(d.lastSyncDateTime) >= thirtyDaysAgo
      ).length;

      // Calculate enrollment status counts
      const enrolledCount = devices.filter(d =>
        d.managementState?.toLowerCase() === 'managed'
      ).length;

      const unmanagedCount = devices.filter(d =>
        d.managementState?.toLowerCase() === 'unmanaged'
      ).length;

      // Calculate enrollment ages
      const now = new Date();
      const enrollmentAges = devices
        .map(d => {
          if (!d.enrolledDateTime) return 0;
          return (now.getTime() - new Date(d.enrolledDateTime).getTime()) / (1000 * 60 * 60 * 24);
        })
        .filter(age => age > 0);

      const averageAge = enrollmentAges.length > 0
        ? Math.round(enrollmentAges.reduce((a, b) => a + b, 0) / enrollmentAges.length)
        : 0;

      // Get enrollment dates
      const enrollmentDates = devices
        .map(d => d.enrolledDateTime)
        .filter(d => d)
        .sort();

      // Calculate success rate
      const totalAttempts = devices.length; // Simplified - would need failure data for accurate calculation
      const successRate = enrolledCount > 0
        ? Math.round((enrolledCount / totalAttempts) * 10000) / 100
        : 0;

      statistics.push({
        userId: userDetails?.id || userId,
        userPrincipalName: userDetails?.userPrincipalName || firstDevice.userPrincipalName || 'Unknown',
        displayName: userDetails?.displayName || firstDevice.userDisplayName || 'Unknown',
        department: userDetails?.department,
        jobTitle: userDetails?.jobTitle,
        email: userDetails?.mail,
        totalDeviceCount: devices.length,
        enrolledDeviceCount: enrolledCount,
        pendingDeviceCount: 0, // Would need additional API data
        failedDeviceCount: 0, // Would need troubleshooting events
        unmanagedDeviceCount: unmanagedCount,
        iosCount: platformCounts.iOS,
        androidCount: platformCounts.Android,
        windowsCount: platformCounts.Windows,
        macOsCount: platformCounts.macOS,
        linuxCount: platformCounts.Linux,
        otherCount: platformCounts.Other,
        corporateCount: ownershipCounts.Corporate,
        personalCount: ownershipCounts.Personal,
        compliantCount: complianceCounts.Compliant,
        nonCompliantCount: complianceCounts.NonCompliant,
        unknownComplianceCount: complianceCounts.Unknown,
        activeDeviceCount: activeCount,
        inactiveDeviceCount: devices.length - activeCount,
        enrollmentSuccessRate: successRate,
        averageEnrollmentAge: averageAge,
        oldestEnrollmentDate: enrollmentDates[0],
        newestEnrollmentDate: enrollmentDates[enrollmentDates.length - 1],
        hasIntuneSubscription,
        subscriptionEnabled: hasIntuneSubscription,
        accountEnabled: userDetails?.accountEnabled !== false,
        createdDateTime: userDetails?.createdDateTime,
        lastSignInDateTime: userDetails?.signInActivity?.lastSignInDateTime
      });
    }

    return statistics;
  }

  /**
   * Enrich devices with user details
   */
  private async enrichDevicesWithUserDetails(
    devices: UserDeviceEnrollment[]
  ): Promise<UserDeviceEnrollment[]> {
    // Group by user to minimize API calls
    const userIds = [...new Set(devices.map(d => d.userId))];
    const userDetailsMap = new Map<string, any>();

    for (const userId of userIds) {
      if (userId && userId !== 'Unknown') {
        const details = await this.fetchUserDetails(userId);
        if (details) {
          userDetailsMap.set(userId, details);
        }
      }
    }

    // Enrich devices with user details
    return devices.map(device => {
      const userDetails = userDetailsMap.get(device.userId);
      if (userDetails) {
        return {
          ...device,
          displayName: userDetails.displayName || device.displayName,
          userPrincipalName: userDetails.userPrincipalName || device.userPrincipalName
        };
      }
      return device;
    });
  }

  // ==========================================================================
  // Mapping Methods
  // ==========================================================================

  /**
   * Map device to UserDeviceEnrollment format
   */
  private mapToUserDeviceEnrollment(device: any): UserDeviceEnrollment {
    const now = new Date();
    const enrolledDate = device.enrolledDateTime ? new Date(device.enrolledDateTime) : now;
    const lastSyncDate = device.lastSyncDateTime ? new Date(device.lastSyncDateTime) : now;

    const daysSinceEnrollment = Math.floor((now.getTime() - enrolledDate.getTime()) / (1000 * 60 * 60 * 24));
    const daysSinceLastSync = Math.floor((now.getTime() - lastSyncDate.getTime()) / (1000 * 60 * 60 * 24));

    // Determine enrollment status
    let enrollmentStatus: 'Enrolled' | 'Pending' | 'Failed' | 'Unmanaged' = 'Enrolled';
    const mgmtState = device.managementState?.toLowerCase() || '';

    if (mgmtState === 'unmanaged' || mgmtState === 'retired') {
      enrollmentStatus = 'Unmanaged';
    } else if (mgmtState.includes('pending')) {
      enrollmentStatus = 'Pending';
    } else if (mgmtState.includes('failed') || mgmtState.includes('error')) {
      enrollmentStatus = 'Failed';
    }

    return {
      deviceId: device.id || '',
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      userId: device.userId || '',
      displayName: device.userDisplayName || 'Unknown',
      platform: this.normalizePlatform(device.operatingSystem || ''),
      osVersion: device.osVersion || 'Unknown',
      model: device.model,
      manufacturer: device.manufacturer,
      enrollmentDate: device.enrolledDateTime || '',
      enrollmentType: device.enrollmentType || 'Unknown',
      managementAgent: device.managementAgent || 'Unknown',
      managementState: device.managementState || 'Unknown',
      complianceState: device.complianceState || 'Unknown',
      ownerType: this.normalizeOwnerType(device.managedDeviceOwnerType),
      lastSyncDateTime: device.lastSyncDateTime || '',
      deviceCategory: device.deviceCategoryDisplayName,
      serialNumber: device.serialNumber,
      imei: device.imei,
      isSupervised: device.isSupervised,
      enrollmentProfileName: device.enrollmentProfileName,
      azureADDeviceId: device.azureADDeviceId,
      deviceActionResults: device.deviceActionResults,
      enrollmentStatus,
      certificateIssued: true, // Assume true if enrolled, would need additional API data
      daysSinceEnrollment,
      daysSinceLastSync
    };
  }

  /**
   * Map device to category device info
   */
  private mapToCategoryDeviceInfo(device: any): CategoryDeviceInfo {
    return {
      deviceId: device.id || '',
      deviceName: device.deviceName || 'Unknown',
      userPrincipalName: device.userPrincipalName || 'Unknown',
      userDisplayName: device.userDisplayName || 'Unknown',
      platform: this.normalizePlatform(device.operatingSystem || ''),
      osVersion: device.osVersion || 'Unknown',
      model: device.model,
      enrollmentDate: device.enrolledDateTime || '',
      managementState: device.managementState || 'Unknown',
      complianceState: device.complianceState || 'Unknown',
      lastSyncDateTime: device.lastSyncDateTime || '',
      ownerType: this.normalizeOwnerType(device.managedDeviceOwnerType)
    };
  }

  /**
   * Map troubleshooting event to unmanaged device
   */
  private mapToUnmanagedDevice(
    event: any,
    failureType: UnmanagedDevice['failureType']
  ): UnmanagedDevice {
    const errorCode = event.correlationId || event.failureReason || 'Unknown';
    const { category, steps } = this.categorizeFailureDetails(errorCode);

    const now = new Date();
    const eventDate = event.eventDateTime ? new Date(event.eventDateTime) : now;
    const daysSinceFailure = Math.floor((now.getTime() - eventDate.getTime()) / (1000 * 60 * 60 * 24));

    return {
      deviceId: event.managedDeviceIdentifier,
      deviceName: event.deviceName,
      userId: event.userId,
      userPrincipalName: event.userPrincipalName || 'Unknown',
      userDisplayName: event.userDisplayName,
      platform: event.platform,
      failureType,
      failureReason: event.failureDetails || event.additionalInformation || 'Unknown',
      failureCategory: category,
      errorCode,
      errorMessage: event.failureDetails,
      managementState: 'Failed',
      enrollmentState: event.enrollmentState,
      certificateIssued: false,
      enrollmentAttemptDate: event.eventDateTime,
      lastContactDate: event.eventDateTime,
      daysSinceFailure,
      troubleshootingSteps: steps,
      troubleshootingEventId: event.id,
      relatedEventCount: 1,
      isRecoverable: this.isRecoverableFailure(errorCode),
      recoveryActions: this.getRecoveryActions(errorCode),
      priorityLevel: this.getPriorityLevel(errorCode, daysSinceFailure)
    };
  }

  /**
   * Map device to unmanaged device
   */
  private mapDeviceToUnmanagedDevice(
    device: any,
    failureType: UnmanagedDevice['failureType']
  ): UnmanagedDevice {
    const mgmtState = device.managementState || 'Unknown';
    const errorCode = mgmtState;
    const { category, steps } = this.categorizeFailureDetails(errorCode);

    const now = new Date();
    const enrollDate = device.enrolledDateTime ? new Date(device.enrolledDateTime) : now;
    const daysSinceFailure = Math.floor((now.getTime() - enrollDate.getTime()) / (1000 * 60 * 60 * 24));

    return {
      deviceId: device.id,
      deviceName: device.deviceName,
      userId: device.userId,
      userPrincipalName: device.userPrincipalName || 'Unknown',
      userDisplayName: device.userDisplayName,
      platform: this.normalizePlatform(device.operatingSystem || ''),
      osVersion: device.osVersion,
      serialNumber: device.serialNumber,
      failureType,
      failureReason: `Device management state: ${mgmtState}`,
      failureCategory: category,
      errorCode,
      managementState: mgmtState,
      certificateIssued: device.enrolledDateTime ? true : false,
      enrollmentAttemptDate: device.enrolledDateTime,
      lastContactDate: device.lastSyncDateTime,
      daysSinceFailure,
      troubleshootingSteps: steps,
      isRecoverable: this.isRecoverableFailure(errorCode),
      recoveryActions: this.getRecoveryActions(errorCode),
      priorityLevel: this.getPriorityLevel(errorCode, daysSinceFailure)
    };
  }

  // ==========================================================================
  // Calculation Methods
  // ==========================================================================

  /**
   * Calculate platform breakdown
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
   * Calculate compliance breakdown
   */
  private calculateComplianceBreakdown(devices: any[]): ComplianceBreakdown {
    const breakdown: ComplianceBreakdown = {
      Compliant: 0,
      NonCompliant: 0,
      Unknown: 0,
      InGracePeriod: 0,
      Error: 0
    };

    devices.forEach(device => {
      const state = device.complianceState?.toLowerCase() || 'unknown';

      if (state === 'compliant') {
        breakdown.Compliant++;
      } else if (state === 'noncompliant') {
        breakdown.NonCompliant++;
      } else if (state === 'ingraceperiod') {
        breakdown.InGracePeriod++;
      } else if (state === 'error') {
        breakdown.Error++;
      } else {
        breakdown.Unknown++;
      }
    });

    return breakdown;
  }

  /**
   * Calculate enrollment trend
   */
  private calculateEnrollmentTrend(devices: any[]): EnrollmentTrendData {
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
      ).length,
      total: devices.length
    };
  }

  /**
   * Calculate enrollment trends over time
   */
  private calculateEnrollmentTrends(
    successfulDevices: any[],
    failedEvents: any[],
    days: number
  ): any[] {
    const trends: any[] = [];
    const now = new Date();

    for (let i = days - 1; i >= 0; i--) {
      const date = new Date(now);
      date.setDate(date.getDate() - i);
      const dateStr = date.toISOString().split('T')[0];

      const daySuccessful = successfulDevices.filter(d => {
        if (!d.enrolledDateTime) return false;
        const enrollDate = new Date(d.enrolledDateTime).toISOString().split('T')[0];
        return enrollDate === dateStr;
      }).length;

      const dayFailed = failedEvents.filter(e => {
        if (!e.eventDateTime) return false;
        const eventDate = new Date(e.eventDateTime).toISOString().split('T')[0];
        return eventDate === dateStr;
      }).length;

      const attempts = daySuccessful + dayFailed;
      const successRate = attempts > 0
        ? Math.round((daySuccessful / attempts) * 10000) / 100
        : 0;

      trends.push({
        date: dateStr,
        attempts,
        successful: daySuccessful,
        failed: dayFailed,
        successRate
      });
    }

    return trends;
  }

  /**
   * Calculate device list summary
   */
  private calculateDeviceListSummary(devices: UserDeviceEnrollment[]): any {
    const platformCounts = this.calculatePlatformBreakdown(devices);
    const complianceCounts = this.calculateComplianceBreakdown(devices);
    const ownerTypeCounts = this.calculateOwnerTypeBreakdown(devices);

    const uniqueUsers = new Set(devices.map(d => d.userId));

    return {
      totalDevices: devices.length,
      totalUsers: uniqueUsers.size,
      platformBreakdown: platformCounts,
      complianceBreakdown: complianceCounts,
      ownerTypeBreakdown: ownerTypeCounts,
      enrolledDevices: devices.filter(d => d.enrollmentStatus === 'Enrolled').length,
      unmanagedDevices: devices.filter(d => d.enrollmentStatus === 'Unmanaged').length,
      failedDevices: devices.filter(d => d.enrollmentStatus === 'Failed').length
    };
  }

  /**
   * Calculate enrollment count summary
   */
  private calculateEnrollmentCountSummary(statistics: UserEnrollmentStatistics[]): any {
    const totalDevices = statistics.reduce((sum, s) => sum + s.totalDeviceCount, 0);
    const totalEnrolled = statistics.reduce((sum, s) => sum + s.enrolledDeviceCount, 0);
    const totalUnmanaged = statistics.reduce((sum, s) => sum + s.unmanagedDeviceCount, 0);
    const totalCompliant = statistics.reduce((sum, s) => sum + s.compliantCount, 0);

    const averageDevicesPerUser = statistics.length > 0
      ? Math.round((totalDevices / statistics.length) * 100) / 100
      : 0;

    const averageSuccessRate = statistics.length > 0
      ? Math.round((statistics.reduce((sum, s) => sum + s.enrollmentSuccessRate, 0) / statistics.length) * 100) / 100
      : 0;

    const usersWithSubscription = statistics.filter(s => s.hasIntuneSubscription).length;

    return {
      totalUsers: statistics.length,
      totalDevices,
      averageDevicesPerUser,
      totalEnrolledDevices: totalEnrolled,
      totalUnmanagedDevices: totalUnmanaged,
      totalCompliantDevices: totalCompliant,
      averageSuccessRate,
      usersWithIntuneSubscription: usersWithSubscription,
      usersWithMultipleDevices: statistics.filter(s => s.totalDeviceCount > 1).length,
      usersWithInactiveDevices: statistics.filter(s => s.inactiveDeviceCount > 0).length
    };
  }

  /**
   * Calculate unmanaged devices summary
   */
  private calculateUnmanagedDevicesSummary(devices: UnmanagedDevice[]): any {
    const failureTypeCounts = new Map<string, number>();
    const priorityCounts = new Map<string, number>();

    devices.forEach(device => {
      failureTypeCounts.set(
        device.failureType,
        (failureTypeCounts.get(device.failureType) || 0) + 1
      );

      priorityCounts.set(
        device.priorityLevel,
        (priorityCounts.get(device.priorityLevel) || 0) + 1
      );
    });

    return {
      totalUnmanagedDevices: devices.length,
      enrollmentFailures: failureTypeCounts.get('EnrollmentFailed') || 0,
      assignmentFailures: failureTypeCounts.get('AssignmentFailed') || 0,
      certificateFailures: failureTypeCounts.get('CertificateFailed') || 0,
      managementLostDevices: failureTypeCounts.get('ManagementLost') || 0,
      highPriorityIssues: priorityCounts.get('High') || 0,
      mediumPriorityIssues: priorityCounts.get('Medium') || 0,
      lowPriorityIssues: priorityCounts.get('Low') || 0,
      recoverableDevices: devices.filter(d => d.isRecoverable).length,
      averageDaysSinceFailure: devices.length > 0
        ? Math.round(devices.reduce((sum, d) => sum + (d.daysSinceFailure || 0), 0) / devices.length)
        : 0
    };
  }

  // ==========================================================================
  // Sorting Methods
  // ==========================================================================

  /**
   * Sort devices
   */
  private sortDevices(
    devices: UserDeviceEnrollment[],
    filters: EnrollmentSummaryFilters
  ): UserDeviceEnrollment[] {
    const sortBy = filters.sortBy || 'enrollmentDate';
    const sortOrder = filters.sortOrder || 'desc';

    return [...devices].sort((a, b) => {
      let comparison = 0;

      switch (sortBy) {
        case 'userName':
          comparison = a.displayName.localeCompare(b.displayName);
          break;
        case 'enrollmentDate':
          const aDate = a.enrollmentDate ? new Date(a.enrollmentDate).getTime() : 0;
          const bDate = b.enrollmentDate ? new Date(b.enrollmentDate).getTime() : 0;
          comparison = aDate - bDate;
          break;
        case 'platform':
          comparison = a.platform.localeCompare(b.platform);
          break;
        case 'category':
          comparison = (a.deviceCategory || '').localeCompare(b.deviceCategory || '');
          break;
        default:
          comparison = 0;
      }

      return sortOrder === 'asc' ? comparison : -comparison;
    });
  }

  /**
   * Sort user statistics
   */
  private sortUserStatistics(
    statistics: UserEnrollmentStatistics[],
    filters: EnrollmentSummaryFilters
  ): UserEnrollmentStatistics[] {
    const sortBy = filters.sortBy || 'deviceCount';
    const sortOrder = filters.sortOrder || 'desc';

    return [...statistics].sort((a, b) => {
      let comparison = 0;

      switch (sortBy) {
        case 'deviceCount':
          comparison = a.totalDeviceCount - b.totalDeviceCount;
          break;
        case 'userName':
          comparison = a.displayName.localeCompare(b.displayName);
          break;
        case 'enrollmentDate':
          const aDate = a.newestEnrollmentDate ? new Date(a.newestEnrollmentDate).getTime() : 0;
          const bDate = b.newestEnrollmentDate ? new Date(b.newestEnrollmentDate).getTime() : 0;
          comparison = aDate - bDate;
          break;
        default:
          comparison = 0;
      }

      return sortOrder === 'asc' ? comparison : -comparison;
    });
  }

  /**
   * Sort unmanaged devices
   */
  private sortUnmanagedDevices(devices: UnmanagedDevice[]): UnmanagedDevice[] {
    // Sort by priority (High > Medium > Low) and then by days since failure
    const priorityOrder = { High: 0, Medium: 1, Low: 2 };

    return [...devices].sort((a, b) => {
      const priorityDiff = priorityOrder[a.priorityLevel] - priorityOrder[b.priorityLevel];
      if (priorityDiff !== 0) return priorityDiff;

      return (b.daysSinceFailure || 0) - (a.daysSinceFailure || 0);
    });
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
   * Categorize failure
   */
  private categorizeFailure(errorCode: string): string {
    const code = errorCode.toLowerCase();

    if (code.includes('80180002')) return 'License Issue';
    if (code.includes('80180014')) return 'Device Limit Exceeded';
    if (code.includes('0x80180026')) return 'Authentication Failed';
    if (code.includes('0xcaa')) return 'Certificate Issue';
    if (code.includes('80180018')) return 'Platform Not Supported';
    if (code.includes('0x80072ee7')) return 'Network Connectivity';
    if (code.includes('unmanaged')) return 'Management Lost';
    if (code.includes('retired')) return 'Device Retired';

    return 'General Enrollment Failure';
  }

  /**
   * Categorize failure with details
   */
  private categorizeFailureDetails(errorCode: string): { category: string; steps: string[] } {
    const category = this.categorizeFailure(errorCode);
    const steps = this.getTroubleshootingSteps(category);

    return { category, steps };
  }

  /**
   * Get troubleshooting steps
   */
  private getTroubleshootingSteps(category: string): string[] {
    const troubleshootingMap: Record<string, string[]> = {
      'License Issue': [
        'Verify user has active Intune license',
        'Check Azure AD Premium licensing',
        'Review license assignment in M365 admin center',
        'Wait up to 24 hours for license propagation'
      ],
      'Device Limit Exceeded': [
        'Check device enrollment restrictions',
        'Review device limit settings',
        'Remove old or unused devices',
        'Verify device type restrictions'
      ],
      'Authentication Failed': [
        'Verify user credentials',
        'Check MFA configuration',
        'Review Conditional Access policies',
        'Ensure account is not locked'
      ],
      'Certificate Issue': [
        'Verify NDES/SCEP configuration',
        'Check certificate connector status',
        'Review certificate profiles',
        'Validate certificate template permissions'
      ],
      'Platform Not Supported': [
        'Verify device OS version',
        'Check platform enrollment restrictions',
        'Review supported OS versions',
        'Update device firmware if needed'
      ],
      'Network Connectivity': [
        'Check internet connectivity',
        'Verify firewall rules',
        'Review proxy settings',
        'Ensure required URLs are accessible'
      ],
      'Management Lost': [
        'Check device sync status',
        'Verify device is powered on',
        'Review management certificate',
        'Consider re-enrollment if needed'
      ],
      'Device Retired': [
        'Device was intentionally retired',
        'Review retirement reason',
        'Re-enroll device if needed',
        'Clean up Azure AD registration if required'
      ]
    };

    return troubleshootingMap[category] || [
      'Review error details in Intune portal',
      'Check device enrollment restrictions',
      'Verify user licenses and permissions',
      'Review troubleshooting events',
      'Contact Microsoft support if issue persists'
    ];
  }

  /**
   * Check if failure is recoverable
   */
  private isRecoverableFailure(errorCode: string): boolean {
    const code = errorCode.toLowerCase();

    // Non-recoverable scenarios
    if (code.includes('retired') || code.includes('deleted')) {
      return false;
    }

    // Most other failures are potentially recoverable
    return true;
  }

  /**
   * Get recovery actions
   */
  private getRecoveryActions(errorCode: string): string[] {
    const category = this.categorizeFailure(errorCode);

    const recoveryMap: Record<string, string[]> = {
      'License Issue': [
        'Assign Intune license to user',
        'Wait for license sync (up to 24 hours)',
        'Retry enrollment'
      ],
      'Device Limit Exceeded': [
        'Remove unused devices',
        'Increase device limit if appropriate',
        'Retry enrollment'
      ],
      'Authentication Failed': [
        'Reset user password',
        'Clear cached credentials',
        'Re-authenticate',
        'Retry enrollment'
      ],
      'Certificate Issue': [
        'Verify certificate infrastructure',
        'Regenerate certificate',
        'Update certificate profile',
        'Retry enrollment'
      ],
      'Network Connectivity': [
        'Verify network connection',
        'Connect to corporate network',
        'Retry enrollment'
      ],
      'Management Lost': [
        'Sync device manually',
        'Restart device',
        'Re-enroll if sync fails'
      ]
    };

    return recoveryMap[category] || [
      'Review and resolve underlying issue',
      'Retry enrollment process',
      'Contact IT support if issue persists'
    ];
  }

  /**
   * Get priority level for failure
   */
  private getPriorityLevel(errorCode: string, daysSinceFailure: number): 'High' | 'Medium' | 'Low' {
    // High priority: Recent failures or critical issues
    if (daysSinceFailure <= 1) return 'High';

    const category = this.categorizeFailure(errorCode);

    if (category === 'Certificate Issue' || category === 'Authentication Failed') {
      return 'High';
    }

    // Medium priority: Failures within a week
    if (daysSinceFailure <= 7) return 'Medium';

    // Low priority: Older failures
    return 'Low';
  }

  /**
   * Deduplicate devices
   */
  private deduplicateDevices(devices: any[]): any[] {
    const deviceMap = new Map<string, any>();

    devices.forEach(device => {
      const id = device.id || device.deviceId;
      if (id && !deviceMap.has(id)) {
        deviceMap.set(id, device);
      }
    });

    return Array.from(deviceMap.values());
  }

  /**
   * Deduplicate unmanaged devices
   */
  private deduplicateUnmanagedDevices(devices: UnmanagedDevice[]): UnmanagedDevice[] {
    const deviceMap = new Map<string, UnmanagedDevice>();

    devices.forEach(device => {
      const key = device.deviceId || device.userPrincipalName + device.serialNumber;
      const existing = deviceMap.get(key);

      if (!existing || this.getFailurePriority(device) > this.getFailurePriority(existing)) {
        deviceMap.set(key, device);
      }
    });

    return Array.from(deviceMap.values());
  }

  /**
   * Get failure priority for deduplication
   */
  private getFailurePriority(device: UnmanagedDevice): number {
    const typePriority: Record<string, number> = {
      'CertificateFailed': 4,
      'EnrollmentFailed': 3,
      'AssignmentFailed': 2,
      'ManagementLost': 1,
      'Unknown': 0
    };

    return typePriority[device.failureType] || 0;
  }

  /**
   * Enrich with troubleshooting information
   */
  private enrichWithTroubleshooting(devices: UnmanagedDevice[]): UnmanagedDevice[] {
    return devices.map(device => {
      // Already has troubleshooting steps from mapping
      return device;
    });
  }

  // ==========================================================================
  // Export Methods
  // ==========================================================================

  /**
   * Export report as JSON
   */
  exportAsJSON(reportData: ReportData): string {
    return JSON.stringify(reportData, null, 2);
  }

  /**
   * Export report as CSV
   */
  exportAsCSV(reportData: ReportData): string {
    const lines: string[] = [];

    // Header
    lines.push(reportData.metadata.reportName);
    lines.push(`Generated: ${reportData.metadata.generatedAt}`);
    lines.push(`Total Records: ${reportData.metadata.recordCount}`);
    lines.push('');

    // Determine report type and format accordingly
    const firstRecord = reportData.data[0];

    if (!firstRecord) {
      lines.push('No data available');
      return lines.join('\n');
    }

    // Get headers from first record
    const headers = Object.keys(firstRecord);
    lines.push(headers.join(','));

    // Add data rows
    reportData.data.forEach(record => {
      const values = headers.map(header => {
        const value = (record as any)[header];
        if (value === null || value === undefined) return '';
        if (typeof value === 'object') return JSON.stringify(value).replace(/"/g, '""');
        return `"${String(value).replace(/"/g, '""')}"`;
      });
      lines.push(values.join(','));
    });

    return lines.join('\n');
  }

  /**
   * Export report as HTML
   */
  exportAsHTML(reportData: ReportData): string {
    const metadata = reportData.metadata;

    let html = `<!DOCTYPE html>
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
      max-width: 1600px;
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
    .metadata {
      background: #f9f9f9;
      padding: 15px;
      border-radius: 6px;
      margin: 20px 0;
      border-left: 4px solid #0078d4;
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
      font-size: 13px;
    }
    th {
      background: #0078d4;
      color: white;
      padding: 12px 8px;
      text-align: left;
      font-weight: 600;
    }
    td {
      padding: 10px 8px;
      border-bottom: 1px solid #e0e0e0;
    }
    tr:hover {
      background-color: #f5f5f5;
    }
    .badge {
      display: inline-block;
      padding: 3px 8px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: bold;
    }
    .badge-success { background: #4caf50; color: white; }
    .badge-warning { background: #ff9800; color: white; }
    .badge-danger { background: #f44336; color: white; }
    .badge-info { background: #2196f3; color: white; }
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

    // Add summary section
    if (reportData.summary) {
      html += '<div class="summary">';
      Object.entries(reportData.summary).forEach(([key, value]) => {
        if (typeof value !== 'object') {
          const label = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
          html += `
        <div class="summary-card">
          <h3>${label}</h3>
          <p>${typeof value === 'number' ? Math.round(value * 100) / 100 : value}</p>
        </div>`;
        }
      });
      html += '</div>';
    }

    // Add data table (simplified - full implementation would be more complex)
    if (reportData.data && reportData.data.length > 0) {
      html += '<h2>Report Data</h2>';
      html += `<p>Total records: ${reportData.data.length}</p>`;
      html += '<p><em>View full details in JSON or CSV export format</em></p>';
    }

    html += `
    <div class="footer">
      <p>Intune Enrollment Summary Report | Configuration Manager Style</p>
      <p>Generated by Intune Reporting Dashboard</p>
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

export default IntuneEnrollmentSummaryReports;
