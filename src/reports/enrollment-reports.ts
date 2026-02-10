/**
 * Intune Enrollment Reporting Module
 *
 * This module provides comprehensive enrollment reporting functionality for Microsoft Intune,
 * including device enrollment trends, failures, user status, platform breakdown, and Autopilot deployment.
 */

import { Client } from '@microsoft/microsoft-graph-client';

// ============================================================================
// Type Definitions
// ============================================================================

export interface EnrollmentTrend {
  date: string;
  totalEnrollments: number;
  successfulEnrollments: number;
  failedEnrollments: number;
  platformBreakdown: {
    iOS: number;
    android: number;
    windows: number;
    macOS: number;
  };
}

export interface EnrollmentFailure {
  deviceId?: string;
  userId: string;
  userPrincipalName: string;
  deviceName?: string;
  platform: string;
  failureReason: string;
  failureCategory: string;
  errorCode: string;
  timestamp: string;
  troubleshootingSteps: string[];
}

export interface UserEnrollmentStatus {
  userId: string;
  userPrincipalName: string;
  displayName: string;
  enrolledDeviceCount: number;
  devices: Array<{
    deviceId: string;
    deviceName: string;
    platform: string;
    enrollmentDate: string;
    complianceState: string;
    managementState: string;
  }>;
  lastEnrollmentDate?: string;
  hasEnrollmentFailures: boolean;
}

export interface PlatformEnrollmentStats {
  platform: string;
  totalDevices: number;
  enrolledLast30Days: number;
  enrolledLast7Days: number;
  enrolledToday: number;
  failureRate: number;
  averageEnrollmentTime?: string;
  topModels: Array<{
    model: string;
    count: number;
  }>;
}

export interface AutopilotDeploymentStatus {
  deviceSerialNumber: string;
  deviceName?: string;
  assignedUser?: string;
  deploymentProfile: string;
  deploymentState: string;
  lastContactDateTime: string;
  enrollmentState: string;
  errorCode?: string;
  errorDescription?: string;
  deploymentStartTime?: string;
  deploymentEndTime?: string;
  deploymentDuration?: number;
  currentPhase?: string;
}

export interface EnrollmentReport {
  reportDate: string;
  reportPeriod: {
    startDate: string;
    endDate: string;
  };
  summary: {
    totalDevices: number;
    totalEnrollments: number;
    successfulEnrollments: number;
    failedEnrollments: number;
    successRate: number;
    activeUsers: number;
  };
  trends: EnrollmentTrend[];
  platformStats: PlatformEnrollmentStats[];
  recentFailures: EnrollmentFailure[];
  autopilotStatus: AutopilotDeploymentStatus[];
}

export interface EnrollmentReportOptions {
  startDate?: Date;
  endDate?: Date;
  includeFailures?: boolean;
  includeTrends?: boolean;
  includeAutopilot?: boolean;
  platformFilter?: string[];
  userFilter?: string[];
  top?: number;
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Categorize enrollment failure based on error code
 */
function categorizeEnrollmentFailure(errorCode: string): { category: string; steps: string[] } {
  const errorCategories: Record<string, { category: string; steps: string[] }> = {
    '80180002': {
      category: 'License Issue',
      steps: [
        'Verify the user has an active Intune license',
        'Check Azure AD Premium licensing if using Conditional Access',
        'Ensure the license is properly assigned in Microsoft 365 admin center',
        'Wait up to 24 hours for license propagation'
      ]
    },
    '80180014': {
      category: 'Device Cap Reached',
      steps: [
        'Check device enrollment restrictions in Intune portal',
        'Review the device limit settings for the user',
        'Consider removing old or unused devices',
        'Verify device type restrictions are not blocking enrollment'
      ]
    },
    '0x80180026': {
      category: 'Authentication Failed',
      steps: [
        'Verify user credentials are correct',
        'Check if MFA is properly configured',
        'Ensure the user account is not locked or disabled',
        'Review Conditional Access policies that may block enrollment',
        'Clear browser cache and try again'
      ]
    },
    '0xcaa9001f': {
      category: 'Certificate Issue',
      steps: [
        'Verify NDES/SCEP configuration',
        'Check certificate connector is running',
        'Review certificate profiles in Intune',
        'Ensure device can reach certificate endpoints',
        'Check certificate template permissions'
      ]
    },
    '80180018': {
      category: 'Platform Not Supported',
      steps: [
        'Verify the device OS version is supported',
        'Check platform enrollment restrictions',
        'Review device type restrictions in Intune portal',
        'Ensure the platform is enabled for enrollment'
      ]
    },
    '0x80072ee7': {
      category: 'Network Connectivity',
      steps: [
        'Check internet connectivity on the device',
        'Verify firewall rules allow Intune endpoints',
        'Ensure proxy settings are correct',
        'Check if required URLs are accessible',
        'Review network security appliances for blocking'
      ]
    },
    'default': {
      category: 'General Enrollment Failure',
      steps: [
        'Review error message details in Intune portal',
        'Check device enrollment restrictions',
        'Verify user has proper licenses',
        'Ensure device meets minimum OS requirements',
        'Try enrolling again after a few minutes',
        'Contact Microsoft support with error code'
      ]
    }
  };

  return errorCategories[errorCode] || errorCategories['default'];
}

/**
 * Calculate date range for reporting period
 */
function calculateDateRange(days: number): { startDate: Date; endDate: Date } {
  const endDate = new Date();
  const startDate = new Date();
  startDate.setDate(startDate.getDate() - days);
  return { startDate, endDate };
}

/**
 * Format duration in minutes
 */
function formatDuration(startTime: string, endTime: string): number {
  const start = new Date(startTime).getTime();
  const end = new Date(endTime).getTime();
  return Math.round((end - start) / 1000 / 60); // Duration in minutes
}

// ============================================================================
// Graph API Query Functions
// ============================================================================

/**
 * Get all managed devices with enrollment information
 */
async function getManagedDevices(
  graphClient: Client,
  options?: { filter?: string; top?: number }
): Promise<any[]> {
  try {
    let request = graphClient.api('/deviceManagement/managedDevices').select([
      'id',
      'deviceName',
      'managedDeviceOwnerType',
      'enrolledDateTime',
      'lastSyncDateTime',
      'operatingSystem',
      'osVersion',
      'model',
      'manufacturer',
      'complianceState',
      'managementState',
      'userId',
      'userPrincipalName',
      'userDisplayName',
      'enrollmentType',
      'managementAgent'
    ]);

    if (options?.filter) {
      request = request.filter(options.filter);
    }

    if (options?.top) {
      request = request.top(options.top);
    } else {
      request = request.top(999);
    }

    const response = await request.get();
    return response.value || [];
  } catch (error) {
    console.error('Error fetching managed devices:', error);
    throw new Error(`Failed to fetch managed devices: ${error}`);
  }
}

/**
 * Get device enrollment failures from troubleshooting events
 */
async function getEnrollmentFailures(
  graphClient: Client,
  startDate?: Date,
  endDate?: Date
): Promise<any[]> {
  try {
    let filter = "eventName eq 'Enrollment'";

    if (startDate) {
      const startDateStr = startDate.toISOString();
      filter += ` and eventDateTime ge ${startDateStr}`;
    }

    if (endDate) {
      const endDateStr = endDate.toISOString();
      filter += ` and eventDateTime le ${endDateStr}`;
    }

    const response = await graphClient
      .api('/deviceManagement/troubleshootingEvents')
      .filter(filter)
      .top(999)
      .get();

    return response.value || [];
  } catch (error) {
    console.error('Error fetching enrollment failures:', error);
    // Return empty array if troubleshooting events API is not available
    return [];
  }
}

/**
 * Get Autopilot devices and deployment status
 */
async function getAutopilotDevices(graphClient: Client): Promise<any[]> {
  try {
    const response = await graphClient
      .api('/deviceManagement/windowsAutopilotDeviceIdentities')
      .select([
        'id',
        'serialNumber',
        'model',
        'manufacturer',
        'productKey',
        'enrollmentState',
        'lastContactedDateTime',
        'addressableUserName',
        'userPrincipalName',
        'deploymentProfileAssignmentStatus',
        'deploymentProfileAssignedDateTime',
        'groupTag'
      ])
      .top(999)
      .get();

    return response.value || [];
  } catch (error) {
    console.error('Error fetching Autopilot devices:', error);
    return [];
  }
}

/**
 * Get Autopilot deployment profiles
 */
async function getAutopilotProfiles(graphClient: Client): Promise<any[]> {
  try {
    const response = await graphClient
      .api('/deviceManagement/windowsAutopilotDeploymentProfiles')
      .select(['id', 'displayName', 'description', 'createdDateTime', 'lastModifiedDateTime'])
      .get();

    return response.value || [];
  } catch (error) {
    console.error('Error fetching Autopilot profiles:', error);
    return [];
  }
}

/**
 * Get user enrollment information
 */
async function getUserEnrollmentInfo(
  graphClient: Client,
  userId: string
): Promise<any[]> {
  try {
    const response = await graphClient
      .api(`/users/${userId}/managedDevices`)
      .select([
        'id',
        'deviceName',
        'operatingSystem',
        'enrolledDateTime',
        'complianceState',
        'managementState'
      ])
      .get();

    return response.value || [];
  } catch (error) {
    console.error(`Error fetching enrollment info for user ${userId}:`, error);
    return [];
  }
}

// ============================================================================
// Report Generation Functions
// ============================================================================

/**
 * Generate enrollment trends report
 */
export async function generateEnrollmentTrends(
  graphClient: Client,
  days: number = 30
): Promise<EnrollmentTrend[]> {
  const { startDate, endDate } = calculateDateRange(days);

  // Get all devices enrolled in the period
  const devices = await getManagedDevices(graphClient, {
    filter: `enrolledDateTime ge ${startDate.toISOString()}`
  });

  // Group by date
  const trendMap = new Map<string, EnrollmentTrend>();

  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const dateStr = d.toISOString().split('T')[0];
    trendMap.set(dateStr, {
      date: dateStr,
      totalEnrollments: 0,
      successfulEnrollments: 0,
      failedEnrollments: 0,
      platformBreakdown: {
        iOS: 0,
        android: 0,
        windows: 0,
        macOS: 0
      }
    });
  }

  // Count enrollments by date
  devices.forEach(device => {
    if (device.enrolledDateTime) {
      const dateStr = new Date(device.enrolledDateTime).toISOString().split('T')[0];
      const trend = trendMap.get(dateStr);

      if (trend) {
        trend.totalEnrollments++;
        trend.successfulEnrollments++;

        const platform = device.operatingSystem?.toLowerCase() || 'unknown';
        if (platform.includes('ios')) {
          trend.platformBreakdown.iOS++;
        } else if (platform.includes('android')) {
          trend.platformBreakdown.android++;
        } else if (platform.includes('windows')) {
          trend.platformBreakdown.windows++;
        } else if (platform.includes('mac')) {
          trend.platformBreakdown.macOS++;
        }
      }
    }
  });

  // Get enrollment failures
  const failures = await getEnrollmentFailures(graphClient, startDate, endDate);
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

  return Array.from(trendMap.values()).sort((a, b) => a.date.localeCompare(b.date));
}

/**
 * Generate enrollment failures report with troubleshooting information
 */
export async function generateEnrollmentFailuresReport(
  graphClient: Client,
  days: number = 7
): Promise<EnrollmentFailure[]> {
  const { startDate } = calculateDateRange(days);
  const failures = await getEnrollmentFailures(graphClient, startDate);

  return failures.map(failure => {
    const errorCode = failure.correlationId || failure.failureReason || 'Unknown';
    const { category, steps } = categorizeEnrollmentFailure(errorCode);

    return {
      deviceId: failure.managedDeviceIdentifier,
      userId: failure.userId,
      userPrincipalName: failure.userPrincipalName || 'Unknown',
      deviceName: failure.deviceName,
      platform: failure.platform || 'Unknown',
      failureReason: failure.failureDetails || failure.additionalInformation || 'Unknown reason',
      failureCategory: category,
      errorCode: errorCode,
      timestamp: failure.eventDateTime,
      troubleshootingSteps: steps
    };
  });
}

/**
 * Get user enrollment status
 */
export async function getUserEnrollmentStatus(
  graphClient: Client,
  userId: string
): Promise<UserEnrollmentStatus> {
  try {
    // Get user info
    const user = await graphClient.api(`/users/${userId}`).select(['id', 'userPrincipalName', 'displayName']).get();

    // Get user's devices
    const devices = await getUserEnrollmentInfo(graphClient, userId);

    // Check for enrollment failures
    const failures = await getEnrollmentFailures(graphClient);
    const userFailures = failures.filter(f => f.userId === userId);

    const deviceList = devices.map(device => ({
      deviceId: device.id,
      deviceName: device.deviceName || 'Unknown',
      platform: device.operatingSystem || 'Unknown',
      enrollmentDate: device.enrolledDateTime || 'Unknown',
      complianceState: device.complianceState || 'Unknown',
      managementState: device.managementState || 'Unknown'
    }));

    const enrollmentDates = deviceList
      .map(d => d.enrollmentDate)
      .filter(d => d !== 'Unknown')
      .sort()
      .reverse();

    return {
      userId: user.id,
      userPrincipalName: user.userPrincipalName,
      displayName: user.displayName || 'Unknown',
      enrolledDeviceCount: devices.length,
      devices: deviceList,
      lastEnrollmentDate: enrollmentDates[0],
      hasEnrollmentFailures: userFailures.length > 0
    };
  } catch (error) {
    console.error(`Error getting enrollment status for user ${userId}:`, error);
    throw new Error(`Failed to get user enrollment status: ${error}`);
  }
}

/**
 * Generate enrollment statistics by platform
 */
export async function generatePlatformEnrollmentStats(
  graphClient: Client
): Promise<PlatformEnrollmentStats[]> {
  const allDevices = await getManagedDevices(graphClient);
  const { startDate: date30DaysAgo } = calculateDateRange(30);
  const { startDate: date7DaysAgo } = calculateDateRange(7);
  const { startDate: dateToday } = calculateDateRange(0);

  // Group devices by platform
  const platformMap = new Map<string, any[]>();

  allDevices.forEach(device => {
    const platform = device.operatingSystem || 'Unknown';
    if (!platformMap.has(platform)) {
      platformMap.set(platform, []);
    }
    platformMap.get(platform)?.push(device);
  });

  const stats: PlatformEnrollmentStats[] = [];

  for (const [platform, devices] of platformMap.entries()) {
    const enrolledLast30Days = devices.filter(
      d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= date30DaysAgo
    ).length;

    const enrolledLast7Days = devices.filter(
      d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= date7DaysAgo
    ).length;

    const enrolledToday = devices.filter(d => {
      if (!d.enrolledDateTime) return false;
      const enrollDate = new Date(d.enrolledDateTime);
      const today = new Date();
      return enrollDate.toDateString() === today.toDateString();
    }).length;

    // Count device models
    const modelCount = new Map<string, number>();
    devices.forEach(device => {
      const model = device.model || 'Unknown';
      modelCount.set(model, (modelCount.get(model) || 0) + 1);
    });

    const topModels = Array.from(modelCount.entries())
      .map(([model, count]) => ({ model, count }))
      .sort((a, b) => b.count - a.count)
      .slice(0, 5);

    stats.push({
      platform,
      totalDevices: devices.length,
      enrolledLast30Days,
      enrolledLast7Days,
      enrolledToday,
      failureRate: 0, // Would need enrollment failure data to calculate
      topModels
    });
  }

  return stats.sort((a, b) => b.totalDevices - a.totalDevices);
}

/**
 * Generate Autopilot deployment status report
 */
export async function generateAutopilotDeploymentReport(
  graphClient: Client
): Promise<AutopilotDeploymentStatus[]> {
  const autopilotDevices = await getAutopilotDevices(graphClient);
  const profiles = await getAutopilotProfiles(graphClient);

  const profileMap = new Map(profiles.map(p => [p.id, p.displayName]));

  return autopilotDevices.map(device => {
    const deploymentStartTime = device.deploymentProfileAssignedDateTime;
    const deploymentEndTime = device.lastContactedDateTime;

    let deploymentDuration: number | undefined;
    if (deploymentStartTime && deploymentEndTime) {
      deploymentDuration = formatDuration(deploymentStartTime, deploymentEndTime);
    }

    return {
      deviceSerialNumber: device.serialNumber,
      deviceName: device.model || 'Unknown',
      assignedUser: device.userPrincipalName || device.addressableUserName,
      deploymentProfile: device.deploymentProfileAssignmentStatus || 'Not Assigned',
      deploymentState: device.enrollmentState || 'Unknown',
      lastContactDateTime: device.lastContactedDateTime || 'Never',
      enrollmentState: device.enrollmentState || 'Pending',
      deploymentStartTime: deploymentStartTime,
      deploymentEndTime: deploymentEndTime,
      deploymentDuration: deploymentDuration,
      currentPhase: determineAutopilotPhase(device)
    };
  });
}

/**
 * Determine the current Autopilot deployment phase
 */
function determineAutopilotPhase(device: any): string {
  const state = device.enrollmentState?.toLowerCase() || '';

  if (state === 'enrolled') {
    return 'Completed';
  } else if (state === 'enrolling' || state === 'awaitingfinalconfig') {
    return 'Device Setup';
  } else if (state === 'notcontacted') {
    return 'Awaiting Deployment';
  } else if (state === 'unknown') {
    return 'Pre-provisioning';
  } else if (state.includes('error') || state.includes('failed')) {
    return 'Failed';
  }

  return 'In Progress';
}

/**
 * Generate comprehensive enrollment report
 */
export async function generateEnrollmentReport(
  graphClient: Client,
  options: EnrollmentReportOptions = {}
): Promise<EnrollmentReport> {
  const {
    startDate,
    endDate,
    includeFailures = true,
    includeTrends = true,
    includeAutopilot = true,
    top = 50
  } = options;

  const reportDate = new Date().toISOString();
  const days = startDate && endDate
    ? Math.ceil((endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24))
    : 30;

  // Fetch all data in parallel
  const [allDevices, trends, platformStats, failures, autopilotStatus] = await Promise.all([
    getManagedDevices(graphClient),
    includeTrends ? generateEnrollmentTrends(graphClient, days) : Promise.resolve([]),
    generatePlatformEnrollmentStats(graphClient),
    includeFailures ? generateEnrollmentFailuresReport(graphClient, days) : Promise.resolve([]),
    includeAutopilot ? generateAutopilotDeploymentReport(graphClient) : Promise.resolve([])
  ]);

  // Calculate summary statistics
  const { startDate: periodStart, endDate: periodEnd } = startDate && endDate
    ? { startDate, endDate }
    : calculateDateRange(days);

  const recentDevices = allDevices.filter(
    d => d.enrolledDateTime && new Date(d.enrolledDateTime) >= periodStart
  );

  const totalEnrollments = recentDevices.length + failures.length;
  const successfulEnrollments = recentDevices.length;
  const failedEnrollments = failures.length;
  const successRate = totalEnrollments > 0
    ? (successfulEnrollments / totalEnrollments) * 100
    : 0;

  const uniqueUsers = new Set(allDevices.map(d => d.userId).filter(Boolean));

  return {
    reportDate,
    reportPeriod: {
      startDate: periodStart.toISOString(),
      endDate: periodEnd.toISOString()
    },
    summary: {
      totalDevices: allDevices.length,
      totalEnrollments,
      successfulEnrollments,
      failedEnrollments,
      successRate: Math.round(successRate * 100) / 100,
      activeUsers: uniqueUsers.size
    },
    trends,
    platformStats,
    recentFailures: failures.slice(0, top),
    autopilotStatus
  };
}

/**
 * Get enrollment statistics for a specific date range
 */
export async function getEnrollmentStatistics(
  graphClient: Client,
  startDate: Date,
  endDate: Date
): Promise<{
  totalEnrollments: number;
  successfulEnrollments: number;
  failedEnrollments: number;
  successRate: number;
  byPlatform: Record<string, number>;
}> {
  const devices = await getManagedDevices(graphClient, {
    filter: `enrolledDateTime ge ${startDate.toISOString()} and enrolledDateTime le ${endDate.toISOString()}`
  });

  const failures = await getEnrollmentFailures(graphClient, startDate, endDate);

  const byPlatform: Record<string, number> = {};
  devices.forEach(device => {
    const platform = device.operatingSystem || 'Unknown';
    byPlatform[platform] = (byPlatform[platform] || 0) + 1;
  });

  const totalEnrollments = devices.length + failures.length;
  const successfulEnrollments = devices.length;
  const failedEnrollments = failures.length;
  const successRate = totalEnrollments > 0
    ? (successfulEnrollments / totalEnrollments) * 100
    : 0;

  return {
    totalEnrollments,
    successfulEnrollments,
    failedEnrollments,
    successRate: Math.round(successRate * 100) / 100,
    byPlatform
  };
}

/**
 * Export report to JSON
 */
export function exportReportToJSON(report: EnrollmentReport): string {
  return JSON.stringify(report, null, 2);
}

/**
 * Export report to CSV format
 */
export function exportReportToCSV(report: EnrollmentReport): string {
  const lines: string[] = [];

  // Summary section
  lines.push('Enrollment Report Summary');
  lines.push(`Report Date,${report.reportDate}`);
  lines.push(`Period Start,${report.reportPeriod.startDate}`);
  lines.push(`Period End,${report.reportPeriod.endDate}`);
  lines.push('');
  lines.push('Metric,Value');
  lines.push(`Total Devices,${report.summary.totalDevices}`);
  lines.push(`Total Enrollments,${report.summary.totalEnrollments}`);
  lines.push(`Successful Enrollments,${report.summary.successfulEnrollments}`);
  lines.push(`Failed Enrollments,${report.summary.failedEnrollments}`);
  lines.push(`Success Rate,${report.summary.successRate}%`);
  lines.push(`Active Users,${report.summary.activeUsers}`);
  lines.push('');

  // Platform statistics
  lines.push('Platform Statistics');
  lines.push('Platform,Total Devices,Last 30 Days,Last 7 Days,Today,Top Model');
  report.platformStats.forEach(stat => {
    const topModel = stat.topModels[0]?.model || 'N/A';
    lines.push(
      `${stat.platform},${stat.totalDevices},${stat.enrolledLast30Days},` +
      `${stat.enrolledLast7Days},${stat.enrolledToday},${topModel}`
    );
  });
  lines.push('');

  // Recent failures
  if (report.recentFailures.length > 0) {
    lines.push('Recent Enrollment Failures');
    lines.push('Timestamp,User,Platform,Error Code,Category,Reason');
    report.recentFailures.forEach(failure => {
      lines.push(
        `${failure.timestamp},${failure.userPrincipalName},${failure.platform},` +
        `${failure.errorCode},${failure.failureCategory},"${failure.failureReason}"`
      );
    });
  }

  return lines.join('\n');
}

// ============================================================================
// Main Export
// ============================================================================

export default {
  generateEnrollmentTrends,
  generateEnrollmentFailuresReport,
  getUserEnrollmentStatus,
  generatePlatformEnrollmentStats,
  generateAutopilotDeploymentReport,
  generateEnrollmentReport,
  getEnrollmentStatistics,
  exportReportToJSON,
  exportReportToCSV
};
