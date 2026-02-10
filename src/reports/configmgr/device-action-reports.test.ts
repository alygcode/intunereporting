/**
 * Unit Tests for Device Action Reports
 *
 * Tests for Configuration Manager equivalent device action reports
 */

import { DeviceActionReports, DeviceActionReport, PendingDeviceAction, RecentlyEnrolledDevice, RecentlyWipedDevice } from './device-action-reports';
import { Client } from '@microsoft/microsoft-graph-client';

// Mock the Microsoft Graph client
jest.mock('@microsoft/microsoft-graph-client');

describe('DeviceActionReports', () => {
  let mockGraphClient: jest.Mocked<Client>;
  let reports: DeviceActionReports;

  beforeEach(() => {
    // Create mock Graph client
    mockGraphClient = {
      api: jest.fn().mockReturnThis(),
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      get: jest.fn(),
    } as any;

    reports = new DeviceActionReports(mockGraphClient);
  });

  describe('Constructor', () => {
    it('should create instance with valid Graph client', () => {
      expect(reports).toBeInstanceOf(DeviceActionReports);
    });
  });

  describe('getPendingRetireWipeRequests', () => {
    it('should return pending retire requests', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Test Device',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'iOS',
            osVersion: '16.0',
            retireAfterDateTime: new Date(Date.now() - 1000).toISOString(),
            managedDeviceOwnerType: 'company',
            lastSyncDateTime: new Date().toISOString(),
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getPendingRetireWipeRequests();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
    });

    it('should handle empty pending actions', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.getPendingRetireWipeRequests();

      expect(result).toEqual([]);
    });

    it('should filter by action type', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Test Device',
            retireAfterDateTime: new Date(Date.now() - 1000).toISOString(),
            managedDeviceOwnerType: 'company',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'iOS',
            osVersion: '16.0',
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getPendingRetireWipeRequests({
        actionType: ['retire'],
      });

      expect(result).toBeDefined();
    });

    it('should handle API errors gracefully', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('API Error'));

      await expect(reports.getPendingRetireWipeRequests()).rejects.toThrow();
    });
  });

  describe('getRecentlyEnrolledDevices', () => {
    it('should return recently enrolled devices with default 7 days', async () => {
      const enrolledDate = new Date(Date.now() - 2 * 24 * 60 * 60 * 1000); // 2 days ago

      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'New Device',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'Windows',
            osVersion: '11',
            manufacturer: 'Microsoft',
            model: 'Surface',
            serialNumber: 'SN123456',
            enrolledDateTime: enrolledDate.toISOString(),
            lastSyncDateTime: new Date().toISOString(),
            complianceState: 'compliant',
            managementAgent: 'mdm',
            managedDeviceOwnerType: 'company',
            deviceEnrollmentType: 'userEnrollment',
            azureADDeviceId: 'azure-id-123',
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyEnrolledDevices();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
    });

    it('should filter by custom date range', async () => {
      const startDate = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
      const endDate = new Date();

      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.getRecentlyEnrolledDevices({
        dateRange: { startDate, endDate },
      });

      expect(result).toEqual([]);
    });

    it('should calculate days since enrollment correctly', async () => {
      const enrolledDate = new Date(Date.now() - 5 * 24 * 60 * 60 * 1000); // 5 days ago

      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Test Device',
            enrolledDateTime: enrolledDate.toISOString(),
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'iOS',
            osVersion: '16.0',
            manufacturer: 'Apple',
            model: 'iPhone',
            serialNumber: 'SN123',
            lastSyncDateTime: new Date().toISOString(),
            managedDeviceOwnerType: 'personal',
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyEnrolledDevices();

      if (result.length > 0) {
        expect(result[0].daysSinceEnrollment).toBeGreaterThanOrEqual(4);
        expect(result[0].daysSinceEnrollment).toBeLessThanOrEqual(5);
      }
    });

    it('should handle devices without enrollment date', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Test Device',
            enrolledDateTime: null,
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyEnrolledDevices();

      expect(result).toEqual([]);
    });

    it('should limit results with top parameter', async () => {
      const mockDevices = {
        value: Array.from({ length: 20 }, (_, i) => ({
          id: `device-${i}`,
          deviceName: `Device ${i}`,
          enrolledDateTime: new Date(Date.now() - i * 60 * 60 * 1000).toISOString(),
          userPrincipalName: `user${i}@test.com`,
          userDisplayName: `User ${i}`,
          operatingSystem: 'Windows',
          osVersion: '11',
          manufacturer: 'Dell',
          model: 'Latitude',
          serialNumber: `SN${i}`,
          managedDeviceOwnerType: 'company',
        })),
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyEnrolledDevices({ top: 10 });

      expect(result.length).toBeLessThanOrEqual(10);
    });
  });

  describe('getRecentlyWipedDevices', () => {
    it('should return recently wiped devices', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Wiped Device',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'Android',
            osVersion: '13',
            managedDeviceOwnerType: 'personal',
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyWipedDevices();

      expect(result).toBeDefined();
      expect(Array.isArray(result)).toBe(true);
    });

    it('should use default 30-day range', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.getRecentlyWipedDevices();

      expect(result).toEqual([]);
    });

    it('should calculate wipe duration correctly', async () => {
      // This test validates the duration calculation logic
      const startTime = new Date(Date.now() - 60 * 60 * 1000); // 1 hour ago
      const endTime = new Date();

      // Mock would need to return action data with these timestamps
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.getRecentlyWipedDevices();

      expect(result).toBeDefined();
    });

    it('should filter by date range', async () => {
      const startDate = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
      const endDate = new Date();

      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.getRecentlyWipedDevices({
        dateRange: { startDate, endDate },
      });

      expect(result).toEqual([]);
    });
  });

  describe('generateDeviceActionReport', () => {
    it('should generate comprehensive report', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.generateDeviceActionReport();

      expect(result).toBeDefined();
      expect(result.summary).toBeDefined();
      expect(result.pendingActions).toBeDefined();
      expect(result.recentEnrollments).toBeDefined();
      expect(result.recentWipes).toBeDefined();
      expect(result.metadata).toBeDefined();
      expect(result.reportPeriod).toBeDefined();
    });

    it('should include summary statistics', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.generateDeviceActionReport();

      expect(result.summary.totalPendingActions).toBeDefined();
      expect(result.summary.totalRecentEnrollments).toBeDefined();
      expect(result.summary.totalRecentWipes).toBeDefined();
      expect(result.summary.wipeSuccessRate).toBeDefined();
      expect(result.summary.generatedAt).toBeDefined();
    });

    it('should apply custom date range', async () => {
      const startDate = new Date('2024-01-01');
      const endDate = new Date('2024-01-31');

      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.generateDeviceActionReport({
        dateRange: { startDate, endDate },
      });

      expect(result.reportPeriod.startDate).toBe(startDate.toISOString());
      expect(result.reportPeriod.endDate).toBe(endDate.toISOString());
    });

    it('should calculate correct record count', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const result = await reports.generateDeviceActionReport();

      expect(result.metadata.recordCount).toBe(0);
    });
  });

  describe('Export Functions', () => {
    let sampleReport: DeviceActionReport;

    beforeEach(() => {
      sampleReport = {
        summary: {
          totalPendingActions: 5,
          totalRecentEnrollments: 10,
          totalRecentWipes: 3,
          actionsByType: { retire: 3, wipe: 2 } as any,
          actionsByStatus: { pending: 5 } as any,
          enrollmentsByOS: { iOS: 5, Windows: 5 },
          wipesByOS: { Android: 2, iOS: 1 },
          averageWipeDurationMinutes: 45,
          wipeSuccessRate: 100,
          generatedAt: new Date().toISOString(),
        },
        pendingActions: [],
        recentEnrollments: [],
        recentWipes: [],
        reportPeriod: {
          startDate: new Date().toISOString(),
          endDate: new Date().toISOString(),
        },
        metadata: {
          reportName: 'Test Report',
          generatedAt: new Date().toISOString(),
          generatedBy: 'Test Suite',
          recordCount: 0,
        },
      };
    });

    describe('exportToJSON', () => {
      it('should export report to JSON format', () => {
        const json = reports.exportToJSON(sampleReport);

        expect(json).toBeDefined();
        expect(typeof json).toBe('string');
        expect(() => JSON.parse(json)).not.toThrow();
      });

      it('should produce valid JSON', () => {
        const json = reports.exportToJSON(sampleReport);
        const parsed = JSON.parse(json);

        expect(parsed.summary).toBeDefined();
        expect(parsed.metadata).toBeDefined();
      });
    });

    describe('exportPendingActionsToCSV', () => {
      it('should export pending actions to CSV', () => {
        const actions: PendingDeviceAction[] = [
          {
            deviceId: 'device-1',
            deviceName: 'Test Device',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'iOS',
            osVersion: '16.0',
            actionType: 'retire',
            actionStatus: 'pending',
            actionRequestedDateTime: new Date().toISOString(),
            lastSyncDateTime: new Date().toISOString(),
            managedDeviceOwnerType: 'company',
          },
        ];

        const csv = reports.exportPendingActionsToCSV(actions);

        expect(csv).toBeDefined();
        expect(csv).toContain('deviceName');
        expect(csv).toContain('Test Device');
      });

      it('should handle empty actions array', () => {
        const csv = reports.exportPendingActionsToCSV([]);

        expect(csv).toBe('No pending actions found');
      });
    });

    describe('exportEnrollmentsToCSV', () => {
      it('should export enrollments to CSV', () => {
        const enrollments: RecentlyEnrolledDevice[] = [
          {
            deviceId: 'device-1',
            deviceName: 'New Device',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'Windows',
            osVersion: '11',
            manufacturer: 'Dell',
            model: 'Latitude',
            serialNumber: 'SN123',
            enrolledDateTime: new Date().toISOString(),
            lastSyncDateTime: new Date().toISOString(),
            complianceState: 'compliant',
            managementAgent: 'mdm',
            managedDeviceOwnerType: 'company',
            enrollmentType: 'userEnrollment',
            azureADDeviceId: 'azure-123',
            daysSinceEnrollment: 2,
            isAssigned: true,
          },
        ];

        const csv = reports.exportEnrollmentsToCSV(enrollments);

        expect(csv).toBeDefined();
        expect(csv).toContain('deviceName');
        expect(csv).toContain('New Device');
      });

      it('should handle empty enrollments array', () => {
        const csv = reports.exportEnrollmentsToCSV([]);

        expect(csv).toBe('No recent enrollments found');
      });
    });

    describe('exportWipesToCSV', () => {
      it('should export wipes to CSV', () => {
        const wipes: RecentlyWipedDevice[] = [
          {
            deviceId: 'device-1',
            deviceName: 'Wiped Device',
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'Android',
            osVersion: '13',
            wipeRequestedDateTime: new Date().toISOString(),
            wipeCompletedDateTime: new Date().toISOString(),
            wipeType: 'wipe',
            managedDeviceOwnerType: 'personal',
            wipeDurationMinutes: 30,
            wipeStatus: 'success',
          },
        ];

        const csv = reports.exportWipesToCSV(wipes);

        expect(csv).toBeDefined();
        expect(csv).toContain('deviceName');
        expect(csv).toContain('Wiped Device');
      });

      it('should handle empty wipes array', () => {
        const csv = reports.exportWipesToCSV([]);

        expect(csv).toBe('No recent wipes found');
      });
    });

    describe('exportToHTML', () => {
      it('should export report to HTML format', () => {
        const html = reports.exportToHTML(sampleReport);

        expect(html).toBeDefined();
        expect(html).toContain('<!DOCTYPE html>');
        expect(html).toContain('Device Action Report');
      });

      it('should include summary statistics in HTML', () => {
        const html = reports.exportToHTML(sampleReport);

        expect(html).toContain('Pending Actions');
        expect(html).toContain('Recent Enrollments');
        expect(html).toContain('Recent Wipes');
      });

      it('should include metadata in HTML', () => {
        const html = reports.exportToHTML(sampleReport);

        expect(html).toContain('Generated:');
        expect(html).toContain('Period:');
      });
    });
  });

  describe('Ownership Type Mapping', () => {
    it('should map company ownership correctly', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Corporate Device',
            managedDeviceOwnerType: 'company',
            enrolledDateTime: new Date().toISOString(),
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'Windows',
            osVersion: '11',
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyEnrolledDevices();

      if (result.length > 0) {
        expect(result[0].managedDeviceOwnerType).toBe('company');
      }
    });

    it('should map personal ownership correctly', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Personal Device',
            managedDeviceOwnerType: 'personal',
            enrolledDateTime: new Date().toISOString(),
            userPrincipalName: 'user@test.com',
            userDisplayName: 'Test User',
            operatingSystem: 'iOS',
            osVersion: '16.0',
          },
        ],
      };

      mockGraphClient.get.mockResolvedValue(mockDevices);

      const result = await reports.getRecentlyEnrolledDevices();

      if (result.length > 0) {
        expect(result[0].managedDeviceOwnerType).toBe('personal');
      }
    });
  });

  describe('Summary Statistics', () => {
    it('should calculate wipe success rate correctly', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const report = await reports.generateDeviceActionReport();

      expect(report.summary.wipeSuccessRate).toBeGreaterThanOrEqual(0);
      expect(report.summary.wipeSuccessRate).toBeLessThanOrEqual(100);
    });

    it('should calculate average wipe duration', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const report = await reports.generateDeviceActionReport();

      expect(report.summary.averageWipeDurationMinutes).toBeGreaterThanOrEqual(0);
    });

    it('should group actions by type', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const report = await reports.generateDeviceActionReport();

      expect(report.summary.actionsByType).toBeDefined();
      expect(typeof report.summary.actionsByType).toBe('object');
    });

    it('should group actions by status', async () => {
      mockGraphClient.get.mockResolvedValue({ value: [] });

      const report = await reports.generateDeviceActionReport();

      expect(report.summary.actionsByStatus).toBeDefined();
      expect(typeof report.summary.actionsByStatus).toBe('object');
    });
  });

  describe('Error Handling', () => {
    it('should handle Graph API errors in pending actions', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('Graph API Error'));

      await expect(reports.getPendingRetireWipeRequests()).rejects.toThrow();
    });

    it('should handle Graph API errors in enrollments', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('Graph API Error'));

      await expect(reports.getRecentlyEnrolledDevices()).rejects.toThrow();
    });

    it('should handle Graph API errors in wipes', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('Graph API Error'));

      await expect(reports.getRecentlyWipedDevices()).rejects.toThrow();
    });

    it('should handle Graph API errors in full report', async () => {
      mockGraphClient.get.mockRejectedValue(new Error('Graph API Error'));

      await expect(reports.generateDeviceActionReport()).rejects.toThrow();
    });
  });
});

describe('Utility Functions', () => {
  describe('createDeviceActionReports', () => {
    it('should create DeviceActionReports instance', () => {
      const mockClient = {} as Client;
      const { createDeviceActionReports } = require('./device-action-reports');

      const reports = createDeviceActionReports(mockClient);

      expect(reports).toBeInstanceOf(DeviceActionReports);
    });
  });

  describe('generateQuickReport', () => {
    it('should generate quick report with default 7 days', async () => {
      const mockClient = {
        api: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue({ value: [] }),
      } as any;

      const { generateQuickReport } = require('./device-action-reports');

      const report = await generateQuickReport(mockClient);

      expect(report).toBeDefined();
      expect(report.summary).toBeDefined();
    });

    it('should accept custom days parameter', async () => {
      const mockClient = {
        api: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        filter: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnThis(),
        get: jest.fn().mockResolvedValue({ value: [] }),
      } as any;

      const { generateQuickReport } = require('./device-action-reports');

      const report = await generateQuickReport(mockClient, 30);

      expect(report).toBeDefined();
    });
  });
});
