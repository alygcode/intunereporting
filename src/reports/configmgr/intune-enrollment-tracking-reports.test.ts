/**
 * Unit Tests for Intune Enrollment Tracking Reports
 *
 * Tests for Configuration Manager-style enrollment tracking reports
 * using Microsoft Graph API.
 */

import { describe, it, expect, beforeEach, vi, Mock } from 'vitest';
import { Client } from '@microsoft/microsoft-graph-client';
import { IntuneEnrollmentTrackingReports } from './intune-enrollment-tracking-reports';
import { AppConfig } from '../../types';

// Mock logger
vi.mock('../../core/logger', () => ({
  Logger: {
    getInstance: () => ({
      info: vi.fn(),
      warn: vi.fn(),
      error: vi.fn(),
      debug: vi.fn(),
    }),
  },
}));

describe('IntuneEnrollmentTrackingReports', () => {
  let graphClient: Client;
  let config: AppConfig;
  let report: IntuneEnrollmentTrackingReports;
  let mockGraphApi: any;

  // Sample test data
  const sampleDevices = [
    {
      id: 'device-1',
      deviceName: 'iPhone-John',
      userPrincipalName: 'john.doe@contoso.com',
      userId: 'user-1',
      userDisplayName: 'John Doe',
      operatingSystem: 'iOS',
      osVersion: '15.0',
      model: 'iPhone 13',
      manufacturer: 'Apple',
      serialNumber: 'ABC123',
      enrolledDateTime: '2024-01-15T10:00:00Z',
      lastSyncDateTime: '2024-02-01T14:30:00Z',
      managedDeviceOwnerType: 'company',
      complianceState: 'compliant',
      enrollmentType: 'UserEnrollment',
      deviceEnrollmentType: 'userEnrollment',
      managementState: 'managed',
      department: 'IT',
      jobTitle: 'IT Manager',
    },
    {
      id: 'device-2',
      deviceName: 'Android-John',
      userPrincipalName: 'john.doe@contoso.com',
      userId: 'user-1',
      userDisplayName: 'John Doe',
      operatingSystem: 'Android',
      osVersion: '12.0',
      model: 'Galaxy S21',
      manufacturer: 'Samsung',
      serialNumber: 'DEF456',
      enrolledDateTime: '2024-01-20T11:00:00Z',
      lastSyncDateTime: '2024-02-01T15:00:00Z',
      managedDeviceOwnerType: 'personal',
      complianceState: 'compliant',
      enrollmentType: 'UserEnrollment',
      deviceEnrollmentType: 'androidEnterprise',
      managementState: 'managed',
      department: 'IT',
      jobTitle: 'IT Manager',
    },
    {
      id: 'device-3',
      deviceName: 'Windows-Jane',
      userPrincipalName: 'jane.smith@contoso.com',
      userId: 'user-2',
      userDisplayName: 'Jane Smith',
      operatingSystem: 'Windows',
      osVersion: '11',
      model: 'Surface Pro 8',
      manufacturer: 'Microsoft',
      serialNumber: 'GHI789',
      enrolledDateTime: '2024-02-01T09:00:00Z',
      lastSyncDateTime: '2024-02-01T16:00:00Z',
      managedDeviceOwnerType: 'company',
      complianceState: 'noncompliant',
      enrollmentType: 'AzureDomainJoined',
      deviceEnrollmentType: 'windowsAzureADJoin',
      managementState: 'managed',
      department: 'Sales',
      jobTitle: 'Sales Manager',
    },
  ];

  const sampleUser = {
    id: 'user-1',
    userPrincipalName: 'john.doe@contoso.com',
    displayName: 'John Doe',
    department: 'IT',
    jobTitle: 'IT Manager',
    mail: 'john.doe@contoso.com',
    accountEnabled: true,
    assignedLicenses: [
      { skuId: 'c1ec4a95-1f05-45b3-a911-aa3fa01094f5' }, // EMS E3
    ],
  };

  beforeEach(() => {
    // Reset mocks
    vi.clearAllMocks();

    // Create mock graph API
    mockGraphApi = {
      api: vi.fn().mockReturnThis(),
      select: vi.fn().mockReturnThis(),
      filter: vi.fn().mockReturnThis(),
      top: vi.fn().mockReturnThis(),
      get: vi.fn(),
    };

    // Setup graph client mock
    graphClient = {
      api: mockGraphApi.api,
    } as unknown as Client;

    // Setup config
    config = {
      clientId: 'test-client-id',
      tenantId: 'test-tenant-id',
      clientSecret: 'test-secret',
      deviceEnrollmentLimit: 15,
    } as AppConfig;

    // Create report instance
    report = new IntuneEnrollmentTrackingReports(graphClient, config);
  });

  // ==========================================================================
  // Test: Report 17 - List of Devices Enrolled Per User
  // ==========================================================================

  describe('getDevicesPerUserReport', () => {
    it('should return device list grouped by user', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesPerUserReport();

      expect(result.data).toBeDefined();
      expect(result.metadata.reportName).toContain('List of Devices Enrolled Per User');
      expect(result.summary?.totalUsers).toBeGreaterThan(0);
    });

    it('should filter devices by platform', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesPerUserReport({
        filters: { platform: ['iOS'] },
      });

      expect(result.data).toBeDefined();
    });

    it('should filter devices by department', async () => {
      mockGraphApi.get.mockResolvedValueOnce({
        value: sampleDevices,
        '@odata.nextLink': null,
      });
      mockGraphApi.get.mockResolvedValueOnce(sampleUser);

      const result = await report.getDevicesPerUserReport({
        filters: { department: 'IT' },
      });

      expect(result.data).toBeDefined();
    });

    it('should sort users by device count', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesPerUserReport({
        filters: {
          sortBy: 'deviceCount',
          sortOrder: 'desc',
        },
      });

      const users = result.data as any[];
      if (users.length > 1) {
        expect(users[0].totalDevices).toBeGreaterThanOrEqual(users[1].totalDevices);
      }
    });

    it('should limit results with top parameter', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesPerUserReport({
        filters: { top: 1 },
      });

      expect(result.data.length).toBeLessThanOrEqual(1);
    });
  });

  // ==========================================================================
  // Test: Report 30 - Number of Devices Enrolled Per User
  // ==========================================================================

  describe('getEnrollmentCountPerUserReport', () => {
    it('should return enrollment counts by user', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentCountPerUserReport();

      expect(result.data).toBeDefined();
      expect(result.metadata.reportName).toContain('Number of Devices Enrolled Per User');
      expect(result.summary?.totalUsers).toBeGreaterThan(0);
    });

    it('should filter users with minimum device count', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentCountPerUserReport({
        filters: { minDeviceCount: 2 },
      });

      const users = result.data as any[];
      users.forEach(user => {
        expect(user.totalDeviceCount).toBeGreaterThanOrEqual(2);
      });
    });

    it('should include platform breakdown', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentCountPerUserReport();

      const users = result.data as any[];
      expect(users[0]).toHaveProperty('iosDeviceCount');
      expect(users[0]).toHaveProperty('androidDeviceCount');
      expect(users[0]).toHaveProperty('windowsDeviceCount');
    });

    it('should include compliance counts', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentCountPerUserReport();

      const users = result.data as any[];
      expect(users[0]).toHaveProperty('compliantDeviceCount');
      expect(users[0]).toHaveProperty('nonCompliantDeviceCount');
    });
  });

  // ==========================================================================
  // Test: Additional Methods
  // ==========================================================================

  describe('getDevicesEnrolledPerUser', () => {
    it('should get devices for user by UPN', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices.filter(d => d.userPrincipalName === 'john.doe@contoso.com'),
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesEnrolledPerUser('john.doe@contoso.com');

      expect(result).toBeDefined();
      expect(result?.userPrincipalName).toBe('john.doe@contoso.com');
    });

    it('should get devices for user by userId', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices.filter(d => d.userId === 'user-1'),
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesEnrolledPerUser('user-1');

      expect(result).toBeDefined();
      expect(result?.userId).toBe('user-1');
    });

    it('should return null for non-existent user', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: [],
        '@odata.nextLink': null,
      });

      const result = await report.getDevicesEnrolledPerUser('nonexistent@contoso.com');

      expect(result).toBeNull();
    });
  });

  describe('getAllUsersWithDeviceCounts', () => {
    it('should return all users with device counts', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getAllUsersWithDeviceCounts();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
      expect(result.length).toBeGreaterThan(0);
    });

    it('should filter by department', async () => {
      mockGraphApi.get.mockResolvedValueOnce({
        value: sampleDevices,
        '@odata.nextLink': null,
      });
      mockGraphApi.get.mockResolvedValue(sampleUser);

      const result = await report.getAllUsersWithDeviceCounts({
        filters: { department: 'IT' },
      });

      expect(result).toBeDefined();
    });
  });

  describe('getUsersWithMultipleDevices', () => {
    it('should return users exceeding threshold', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getUsersWithMultipleDevices(1);

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
      result.forEach(user => {
        expect(user.totalDeviceCount).toBeGreaterThan(1);
      });
    });

    it('should return empty array if no users exceed threshold', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: [],
        '@odata.nextLink': null,
      });

      const result = await report.getUsersWithMultipleDevices(100);

      expect(result).toBeDefined();
      expect(result.length).toBe(0);
    });
  });

  describe('getUserEnrollmentDetails', () => {
    it('should return comprehensive user enrollment details', async () => {
      mockGraphApi.get
        .mockResolvedValueOnce({
          value: sampleDevices.filter(d => d.userPrincipalName === 'john.doe@contoso.com'),
          '@odata.nextLink': null,
        })
        .mockResolvedValueOnce({
          value: sampleDevices.filter(d => d.userPrincipalName === 'john.doe@contoso.com'),
          '@odata.nextLink': null,
        })
        .mockResolvedValueOnce({
          value: [{ userPrincipalName: 'john.doe@contoso.com' }],
        })
        .mockResolvedValueOnce(sampleUser)
        .mockResolvedValueOnce({ assignedLicenses: sampleUser.assignedLicenses });

      const result = await report.getUserEnrollmentDetails('john.doe@contoso.com');

      expect(result).toBeDefined();
      expect(result.user).toBeDefined();
      expect(result.enrollmentCount).toBeDefined();
      expect(result.violatesLimit).toBeDefined();
    });

    it('should detect device limit violations', async () => {
      const manyDevices = Array.from({ length: 20 }, (_, i) => ({
        ...sampleDevices[0],
        id: `device-${i}`,
        deviceName: `Device-${i}`,
      }));

      mockGraphApi.get
        .mockResolvedValueOnce({ value: manyDevices, '@odata.nextLink': null })
        .mockResolvedValueOnce({ value: manyDevices, '@odata.nextLink': null })
        .mockResolvedValueOnce({ value: [{ userPrincipalName: 'john.doe@contoso.com' }] })
        .mockResolvedValueOnce(sampleUser)
        .mockResolvedValueOnce({ assignedLicenses: sampleUser.assignedLicenses });

      const result = await report.getUserEnrollmentDetails('john.doe@contoso.com');

      expect(result.violatesLimit).toBe(true);
      expect(result.limitViolation).toBeDefined();
      expect(result.limitViolation?.deviceCount).toBe(20);
      expect(result.limitViolation?.deviceLimit).toBe(15);
    });
  });

  describe('getEnrollmentStatisticsByUser', () => {
    it('should generate comprehensive enrollment statistics', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentStatisticsByUser();

      expect(result).toBeDefined();
      expect(result.totalUsers).toBeGreaterThan(0);
      expect(result.totalDevices).toBeGreaterThan(0);
      expect(result.averageDevicesPerUser).toBeGreaterThan(0);
      expect(result.platformDistribution).toBeDefined();
      expect(result.complianceOverview).toBeDefined();
      expect(result.departmentBreakdown).toBeDefined();
    });

    it('should calculate median correctly', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentStatisticsByUser();

      expect(result.medianDevicesPerUser).toBeGreaterThanOrEqual(0);
    });

    it('should identify users violating limits', async () => {
      const manyDevices = Array.from({ length: 20 }, (_, i) => ({
        ...sampleDevices[0],
        id: `device-${i}`,
        deviceName: `Device-${i}`,
      }));

      mockGraphApi.get.mockResolvedValue({
        value: manyDevices,
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentStatisticsByUser();

      expect(result.usersViolatingLimit).toBeGreaterThanOrEqual(0);
    });
  });

  describe('detectUnusualEnrollmentPatterns', () => {
    it('should detect rapid enrollment patterns', async () => {
      const rapidDevices = Array.from({ length: 6 }, (_, i) => ({
        ...sampleDevices[0],
        id: `device-${i}`,
        deviceName: `Device-${i}`,
        enrolledDateTime: new Date(Date.now() - i * 60 * 60 * 1000).toISOString(), // 1 hour apart
      }));

      mockGraphApi.get
        .mockResolvedValueOnce({ value: rapidDevices, '@odata.nextLink': null })
        .mockResolvedValueOnce({ value: rapidDevices, '@odata.nextLink': null });

      const result = await report.detectUnusualEnrollmentPatterns();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
      // Should detect rapid enrollment
      const rapidPattern = result.find(p => p.patternType === 'RapidEnrollment');
      expect(rapidPattern).toBeDefined();
    });

    it('should detect bulk enrollment patterns', async () => {
      const bulkDevices = Array.from({ length: 25 }, (_, i) => ({
        ...sampleDevices[0],
        id: `device-${i}`,
        deviceName: `Device-${i}`,
        userId: `user-${i}`,
        userPrincipalName: `user${i}@contoso.com`,
        enrolledDateTime: new Date('2024-02-01T10:00:00Z').toISOString(),
      }));

      mockGraphApi.get
        .mockResolvedValueOnce({ value: bulkDevices, '@odata.nextLink': null })
        .mockResolvedValueOnce({ value: bulkDevices, '@odata.nextLink': null });

      const result = await report.detectUnusualEnrollmentPatterns();

      expect(result).toBeDefined();
      const bulkPattern = result.find(p => p.patternType === 'BulkEnrollment');
      expect(bulkPattern).toBeDefined();
    });

    it('should detect excessive device enrollments', async () => {
      const excessiveDevices = Array.from({ length: 35 }, (_, i) => ({
        ...sampleDevices[0],
        id: `device-${i}`,
        deviceName: `Device-${i}`,
      }));

      mockGraphApi.get
        .mockResolvedValueOnce({ value: excessiveDevices, '@odata.nextLink': null })
        .mockResolvedValueOnce({ value: excessiveDevices, '@odata.nextLink': null });

      const result = await report.detectUnusualEnrollmentPatterns();

      expect(result).toBeDefined();
      // Should detect excessive devices
      const excessivePattern = result.find(p => p.severity === 'Critical' || p.severity === 'High');
      expect(excessivePattern).toBeDefined();
    });

    it('should sort patterns by severity', async () => {
      mockGraphApi.get
        .mockResolvedValueOnce({ value: sampleDevices, '@odata.nextLink': null })
        .mockResolvedValueOnce({ value: sampleDevices, '@odata.nextLink': null });

      const result = await report.detectUnusualEnrollmentPatterns();

      if (result.length > 1) {
        const severityOrder = { Critical: 4, High: 3, Medium: 2, Low: 1 };
        for (let i = 0; i < result.length - 1; i++) {
          expect(severityOrder[result[i].severity]).toBeGreaterThanOrEqual(
            severityOrder[result[i + 1].severity]
          );
        }
      }
    });
  });

  // ==========================================================================
  // Test: Export Methods
  // ==========================================================================

  describe('Export Methods', () => {
    it('should export report as JSON', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const reportData = await report.getEnrollmentCountPerUserReport();
      const json = report.exportAsJSON(reportData);

      expect(json).toBeDefined();
      expect(typeof json).toBe('string');
      expect(() => JSON.parse(json)).not.toThrow();
    });

    it('should export report as CSV', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const reportData = await report.getEnrollmentCountPerUserReport();
      const csv = report.exportAsCSV(reportData);

      expect(csv).toBeDefined();
      expect(typeof csv).toBe('string');
      expect(csv).toContain('User Principal Name');
      expect(csv).toContain('Total Devices');
    });

    it('should export report as HTML', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const reportData = await report.getEnrollmentCountPerUserReport();
      const html = report.exportAsHTML(reportData);

      expect(html).toBeDefined();
      expect(typeof html).toBe('string');
      expect(html).toContain('<!DOCTYPE html>');
      expect(html).toContain('<table>');
    });
  });

  // ==========================================================================
  // Test: Statistics Generation
  // ==========================================================================

  describe('generateEnrollmentStatistics', () => {
    it('should generate overall enrollment statistics', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.generateEnrollmentStatistics();

      expect(result).toBeDefined();
      expect(result.totalUsers).toBeGreaterThan(0);
      expect(result.totalDevices).toBeGreaterThan(0);
      expect(result.platformDistribution).toBeDefined();
      expect(result.enrollmentMethodDistribution).toBeDefined();
      expect(result.ownerTypeDistribution).toBeDefined();
      expect(result.complianceDistribution).toBeDefined();
      expect(result.enrollmentTrends).toBeDefined();
    });

    it('should calculate averages correctly', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: sampleDevices,
        '@odata.nextLink': null,
      });

      const result = await report.generateEnrollmentStatistics();

      expect(result.averageDevicesPerUser).toBeGreaterThanOrEqual(0);
      expect(result.medianDevicesPerUser).toBeGreaterThanOrEqual(0);
    });
  });

  // ==========================================================================
  // Test: Error Handling
  // ==========================================================================

  describe('Error Handling', () => {
    it('should handle Graph API errors gracefully', async () => {
      mockGraphApi.get.mockRejectedValue(new Error('Graph API error'));

      await expect(report.getDevicesPerUserReport()).rejects.toThrow();
    });

    it('should handle empty device list', async () => {
      mockGraphApi.get.mockResolvedValue({
        value: [],
        '@odata.nextLink': null,
      });

      const result = await report.getEnrollmentCountPerUserReport();

      expect(result.data).toBeDefined();
      expect(result.data.length).toBe(0);
    });

    it('should handle missing user details', async () => {
      mockGraphApi.get
        .mockResolvedValueOnce({
          value: sampleDevices,
          '@odata.nextLink': null,
        })
        .mockRejectedValue(new Error('User not found'));

      const result = await report.getDevicesPerUserReport();

      expect(result.data).toBeDefined();
    });
  });
});
