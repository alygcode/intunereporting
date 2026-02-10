/**
 * Unit Tests for User Enrollment Report Module
 *
 * This file contains comprehensive unit tests for the UserEnrollmentReport class.
 * It uses Jest for testing and mocks the Microsoft Graph API client.
 */

import { UserEnrollmentReport, EnrollmentFailureCategory } from './user-enrollment-report';
import { AppConfig } from '../types';
import { Client } from '@microsoft/microsoft-graph-client';

// Mock the Logger
jest.mock('../core/logger', () => ({
  Logger: {
    getInstance: jest.fn().mockReturnValue({
      info: jest.fn(),
      error: jest.fn(),
      warn: jest.fn(),
      debug: jest.fn()
    })
  }
}));

describe('UserEnrollmentReport', () => {
  let mockGraphClient: jest.Mocked<Client>;
  let report: UserEnrollmentReport;
  let config: AppConfig;

  beforeEach(() => {
    // Create mock Graph client
    mockGraphClient = {
      api: jest.fn().mockReturnThis(),
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      orderby: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      header: jest.fn().mockReturnThis(),
      count: jest.fn().mockReturnThis(),
      get: jest.fn(),
      post: jest.fn()
    } as any;

    config = {
      tenantId: 'test-tenant-id',
      clientId: 'test-client-id',
      clientSecret: 'test-client-secret',
      outputDir: './test-reports'
    };

    report = new UserEnrollmentReport(mockGraphClient as any, config);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  describe('Report Metadata', () => {
    test('should have correct report metadata', () => {
      expect(report.name).toBe('User Enrollment Report');
      expect(report.description).toContain('enrollment');
      expect(report.category).toBe('Enrollment');
      expect(report.enabled).toBe(true);
    });
  });

  describe('execute()', () => {
    test('should generate basic enrollment report', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          deviceName: 'Test Device 1',
          operatingSystem: 'iOS',
          osVersion: '15.0',
          enrolledDateTime: new Date().toISOString(),
          userId: 'user-1',
          userPrincipalName: 'test@example.com',
          complianceState: 'compliant',
          managementState: 'managed'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const result = await report.execute({
        includeTrends: false,
        includeFailures: false,
        includePlatformStats: false,
        includeUserDetails: false
      });

      expect(result.data).toBeDefined();
      expect(result.metadata).toBeDefined();
      expect(result.summary).toBeDefined();
      expect(mockGraphClient.api).toHaveBeenCalled();
    });

    test('should include trends when requested', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          enrolledDateTime: new Date().toISOString(),
          operatingSystem: 'iOS',
          userId: 'user-1'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const result = await report.execute({
        includeTrends: true,
        includeFailures: false,
        includePlatformStats: false,
        trendDays: 7
      });

      const reportData = result.data[0] as any;
      expect(reportData.trends).toBeDefined();
      expect(Array.isArray(reportData.trends)).toBe(true);
    });

    test('should handle errors gracefully', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('API Error'));

      await expect(report.execute()).rejects.toThrow('API Error');
    });
  });

  describe('getAllEnrolledUsers()', () => {
    test('should fetch all enrolled users', async () => {
      const mockUsers = [
        {
          id: 'user-1',
          userPrincipalName: 'user1@example.com',
          displayName: 'User 1',
          assignedLicenses: [{ skuId: 'license-1' }]
        },
        {
          id: 'user-2',
          userPrincipalName: 'user2@example.com',
          displayName: 'User 2',
          assignedLicenses: [{ skuId: 'license-2' }]
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockUsers,
        '@odata.nextLink': null
      });

      const users = await report.getAllEnrolledUsers();

      expect(users).toHaveLength(2);
      expect(mockGraphClient.api).toHaveBeenCalledWith('/users');
      expect(mockGraphClient.filter).toHaveBeenCalled();
    });

    test('should handle pagination', async () => {
      const mockUsers1 = [{ id: 'user-1', userPrincipalName: 'user1@example.com' }];
      const mockUsers2 = [{ id: 'user-2', userPrincipalName: 'user2@example.com' }];

      mockGraphClient.get
        .mockResolvedValueOnce({
          value: mockUsers1,
          '@odata.nextLink': '/users?$skip=1'
        })
        .mockResolvedValueOnce({
          value: mockUsers2,
          '@odata.nextLink': null
        });

      const users = await report.getAllEnrolledUsers();

      expect(users).toHaveLength(2);
    });
  });

  describe('getUserDevices()', () => {
    test('should fetch devices for a specific user', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          deviceName: 'iPhone',
          operatingSystem: 'iOS',
          osVersion: '15.0',
          enrolledDateTime: new Date().toISOString(),
          managementState: 'managed',
          complianceState: 'compliant'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const devices = await report.getUserDevices('user-123');

      expect(devices).toHaveLength(1);
      expect(devices[0].deviceName).toBe('iPhone');
      expect(devices[0].platform).toBe('iOS');
      expect(mockGraphClient.api).toHaveBeenCalledWith('/users/user-123/managedDevices');
    });

    test('should handle errors and return empty array', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('User not found'));

      const devices = await report.getUserDevices('invalid-user');

      expect(devices).toEqual([]);
    });
  });

  describe('getUserDevicesByUPN()', () => {
    test('should fetch devices by user principal name', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          deviceName: 'Test Device',
          operatingSystem: 'Windows',
          enrolledDateTime: new Date().toISOString()
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const devices = await report.getUserDevicesByUPN('test@example.com');

      expect(devices).toHaveLength(1);
      expect(mockGraphClient.filter).toHaveBeenCalledWith(
        "userPrincipalName eq 'test@example.com'"
      );
    });
  });

  describe('getUserEnrollmentStatus()', () => {
    test('should get comprehensive user enrollment status', async () => {
      const mockUserInfo = {
        id: 'user-1',
        userPrincipalName: 'test@example.com',
        displayName: 'Test User',
        department: 'IT',
        jobTitle: 'Engineer'
      };

      const mockDevices = [
        {
          id: 'device-1',
          deviceName: 'Test Device',
          operatingSystem: 'iOS',
          enrolledDateTime: new Date().toISOString()
        }
      ];

      mockGraphClient.get
        .mockResolvedValueOnce(mockUserInfo) // User info
        .mockResolvedValueOnce({ value: mockDevices, '@odata.nextLink': null }) // Devices
        .mockResolvedValueOnce({ value: [], '@odata.nextLink': null }); // Failures

      const status = await report.getUserEnrollmentStatus('user-1');

      expect(status.userId).toBe('user-1');
      expect(status.displayName).toBe('Test User');
      expect(status.enrolledDeviceCount).toBe(1);
      expect(status.hasEnrollmentFailures).toBe(false);
    });
  });

  describe('getEnrollmentFailures()', () => {
    test('should fetch and categorize enrollment failures', async () => {
      const mockFailures = [
        {
          id: 'failure-1',
          userId: 'user-1',
          userPrincipalName: 'test@example.com',
          eventDateTime: new Date().toISOString(),
          eventName: 'Enrollment Failed',
          correlationId: '80180002',
          platform: 'iOS'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockFailures,
        '@odata.nextLink': null
      });

      const failures = await report.getEnrollmentFailures(30, 50);

      expect(failures).toHaveLength(1);
      expect(failures[0].failureCategory).toBe(EnrollmentFailureCategory.Licensing);
      expect(failures[0].troubleshootingSteps).toBeDefined();
      expect(failures[0].troubleshootingSteps.length).toBeGreaterThan(0);
    });

    test('should limit results to maxResults', async () => {
      const mockFailures = Array.from({ length: 150 }, (_, i) => ({
        id: `failure-${i}`,
        eventDateTime: new Date().toISOString(),
        eventName: 'Enrollment Failed'
      }));

      mockGraphClient.get.mockResolvedValue({
        value: mockFailures,
        '@odata.nextLink': null
      });

      const failures = await report.getEnrollmentFailures(30, 50);

      expect(failures.length).toBeLessThanOrEqual(50);
    });
  });

  describe('getEnrollmentTrends()', () => {
    test('should generate enrollment trends', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          enrolledDateTime: new Date().toISOString(),
          operatingSystem: 'iOS',
          userId: 'user-1'
        },
        {
          id: 'device-2',
          enrolledDateTime: new Date(Date.now() - 24 * 60 * 60 * 1000).toISOString(),
          operatingSystem: 'Android',
          userId: 'user-2'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const trends = await report.getEnrollmentTrends(7);

      expect(trends).toHaveLength(8); // 7 days + today
      expect(trends[0]).toHaveProperty('date');
      expect(trends[0]).toHaveProperty('totalEnrollments');
      expect(trends[0]).toHaveProperty('successfulEnrollments');
      expect(trends[0]).toHaveProperty('platformBreakdown');
    });
  });

  describe('getDeviceEnrollmentByPlatform()', () => {
    test('should generate platform statistics', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          operatingSystem: 'iOS',
          model: 'iPhone 12',
          manufacturer: 'Apple',
          osVersion: '15.0',
          enrolledDateTime: new Date().toISOString()
        },
        {
          id: 'device-2',
          operatingSystem: 'Android',
          model: 'Galaxy S21',
          manufacturer: 'Samsung',
          osVersion: '11.0',
          enrolledDateTime: new Date().toISOString()
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const platformStats = await report.getDeviceEnrollmentByPlatform();

      expect(platformStats.length).toBeGreaterThan(0);
      expect(platformStats[0]).toHaveProperty('platform');
      expect(platformStats[0]).toHaveProperty('totalDevices');
      expect(platformStats[0]).toHaveProperty('topModels');
      expect(platformStats[0]).toHaveProperty('osVersionDistribution');
    });
  });

  describe('getEnrollmentMethodsBreakdown()', () => {
    test('should analyze enrollment methods', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          deviceName: 'Device 1',
          enrollmentType: 'UserEnrollment'
        },
        {
          id: 'device-2',
          deviceName: 'Device 2',
          enrollmentType: 'UserEnrollment'
        },
        {
          id: 'device-3',
          deviceName: 'Device 3',
          enrollmentType: 'DeviceEnrollmentProgram'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const methods = await report.getEnrollmentMethodsBreakdown();

      expect(methods['UserEnrollment']).toBeDefined();
      expect(methods['UserEnrollment'].count).toBe(2);
      expect(methods['DeviceEnrollmentProgram'].count).toBe(1);
    });
  });

  describe('getDeviceOwnershipBreakdown()', () => {
    test('should analyze device ownership', async () => {
      const mockDevices = [
        {
          id: 'device-1',
          managedDeviceOwnerType: 'company',
          operatingSystem: 'iOS'
        },
        {
          id: 'device-2',
          managedDeviceOwnerType: 'personal',
          operatingSystem: 'Android'
        },
        {
          id: 'device-3',
          managedDeviceOwnerType: 'company',
          operatingSystem: 'Windows'
        }
      ];

      mockGraphClient.get.mockResolvedValue({
        value: mockDevices,
        '@odata.nextLink': null
      });

      const ownership = await report.getDeviceOwnershipBreakdown();

      expect(ownership.totalDevices).toBe(3);
      expect(ownership.corporate.count).toBe(2);
      expect(ownership.personal.count).toBe(1);
      expect(ownership.byPlatform).toBeDefined();
    });
  });

  describe('Export Methods', () => {
    const mockReportData = {
      metadata: {
        reportName: 'Test Report',
        generatedAt: new Date().toISOString(),
        generatedBy: 'Test',
        recordCount: 1
      },
      data: [
        {
          statistics: {
            totalDevices: 10,
            totalUsers: 5,
            enrolledLast24Hours: 2,
            enrolledLast7Days: 5,
            enrolledLast30Days: 8,
            enrolledLast90Days: 10,
            averageDevicesPerUser: 2,
            enrollmentSuccessRate: 95,
            totalEnrollmentAttempts: 11,
            failedEnrollments: 1,
            pendingEnrollments: 0,
            lastUpdated: new Date().toISOString()
          },
          platformStats: [
            {
              platform: 'iOS',
              totalDevices: 5,
              enrolledLast24Hours: 1,
              enrolledLast7Days: 2,
              enrolledLast30Days: 4,
              enrolledLast90Days: 5,
              failureCount: 0,
              failureRate: 0,
              topModels: [{ model: 'iPhone 12', count: 3, percentage: 60 }],
              topManufacturers: [{ manufacturer: 'Apple', count: 5, percentage: 100 }],
              osVersionDistribution: [{ osVersion: '15.0', count: 5, percentage: 100 }]
            }
          ],
          failures: []
        }
      ],
      summary: {
        totalDevices: 10,
        totalUsers: 5,
        successRate: 95,
        failedEnrollments: 1
      }
    };

    test('should export as JSON', () => {
      const json = report.exportAsJSON(mockReportData);

      expect(json).toBeDefined();
      expect(() => JSON.parse(json)).not.toThrow();

      const parsed = JSON.parse(json);
      expect(parsed.data).toBeDefined();
      expect(parsed.metadata).toBeDefined();
    });

    test('should export as CSV', () => {
      const csv = report.exportAsCSV(mockReportData);

      expect(csv).toBeDefined();
      expect(csv).toContain('User Enrollment Report');
      expect(csv).toContain('Total Devices');
      expect(csv).toContain('Platform Statistics');
    });

    test('should export as HTML', () => {
      const html = report.exportAsHTML(mockReportData);

      expect(html).toBeDefined();
      expect(html).toContain('<!DOCTYPE html>');
      expect(html).toContain('User Enrollment Report');
      expect(html).toContain('Overview Statistics');
      expect(html).toContain('Platform Statistics');
    });
  });

  describe('Error Handling', () => {
    test('should retry on rate limit (429)', async () => {
      const error = new Error('Rate limited') as any;
      error.statusCode = 429;

      mockGraphClient.get
        .mockRejectedValueOnce(error)
        .mockRejectedValueOnce(error)
        .mockResolvedValueOnce({ value: [], '@odata.nextLink': null });

      const devices = await report.getAllEnrolledUsers();

      expect(devices).toEqual([]);
      expect(mockGraphClient.get).toHaveBeenCalledTimes(3);
    });

    test('should retry on server error (500)', async () => {
      const error = new Error('Server error') as any;
      error.statusCode = 500;

      mockGraphClient.get
        .mockRejectedValueOnce(error)
        .mockResolvedValueOnce({ value: [], '@odata.nextLink': null });

      const devices = await report.getAllEnrolledUsers();

      expect(devices).toEqual([]);
      expect(mockGraphClient.get).toHaveBeenCalledTimes(2);
    });

    test('should not retry on client error (404)', async () => {
      const error = new Error('Not found') as any;
      error.statusCode = 404;

      mockGraphClient.get.mockRejectedValue(error);

      await expect(report.getAllEnrolledUsers()).rejects.toThrow('Not found');
      expect(mockGraphClient.get).toHaveBeenCalledTimes(1);
    });
  });
});
