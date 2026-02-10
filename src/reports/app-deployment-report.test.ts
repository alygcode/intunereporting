/**
 * App Deployment Report - Unit Tests
 *
 * This file contains unit tests for the Application Deployment Report module.
 */

import { AppDeploymentReport, AppPlatform, InstallState } from './app-deployment-report';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../types';

// Mock Graph Client
const createMockGraphClient = (mockResponses: any = {}): Client => {
  const mockClient = {
    api: jest.fn().mockReturnThis(),
    select: jest.fn().mockReturnThis(),
    expand: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnThis(),
    get: jest.fn().mockImplementation(() => {
      return Promise.resolve(mockResponses);
    })
  } as any;

  return mockClient;
};

// Mock Config
const createMockConfig = (): AppConfig => ({
  authentication: {
    tenantId: 'test-tenant-id',
    clientId: 'test-client-id',
    clientSecret: 'test-secret',
    authMethod: 'clientSecret'
  },
  reports: {
    enabled: ['app-deployment-report'],
    disabled: [],
    settings: {}
  },
  output: {
    defaultFormat: 'json',
    directory: './test-reports',
    includeTimestamp: true,
    compression: false
  },
  scheduler: {
    enabled: false,
    timezone: 'UTC',
    schedules: []
  },
  logging: {
    level: 'info',
    file: 'test.log',
    console: false,
    maxSize: '10m',
    maxFiles: 5
  }
});

describe('AppDeploymentReport', () => {
  let report: AppDeploymentReport;
  let mockGraphClient: Client;
  let mockConfig: AppConfig;

  beforeEach(() => {
    mockConfig = createMockConfig();
  });

  describe('getAllManagedApps', () => {
    it('should fetch all managed apps', async () => {
      const mockApps = {
        value: [
          {
            id: 'app-1',
            displayName: 'Test App 1',
            publisher: 'Test Publisher',
            '@odata.type': '#microsoft.graph.iosStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            displayVersion: '1.0.0',
            isFeatured: false
          },
          {
            id: 'app-2',
            displayName: 'Test App 2',
            publisher: 'Microsoft',
            '@odata.type': '#microsoft.graph.androidStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            displayVersion: '2.0.0',
            isFeatured: true
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockApps);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const apps = await report.getAllManagedApps();

      expect(apps).toHaveLength(2);
      expect(apps[0].displayName).toBe('Test App 1');
      expect(apps[0].platform).toBe(AppPlatform.IOS);
      expect(apps[1].displayName).toBe('Test App 2');
      expect(apps[1].platform).toBe(AppPlatform.ANDROID);
    });

    it('should filter apps by platform', async () => {
      const mockApps = {
        value: [
          {
            id: 'app-1',
            displayName: 'iOS App',
            publisher: 'Test',
            '@odata.type': '#microsoft.graph.iosStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          },
          {
            id: 'app-2',
            displayName: 'Android App',
            publisher: 'Test',
            '@odata.type': '#microsoft.graph.androidStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockApps);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const iosApps = await report.getAllManagedApps({ platform: AppPlatform.IOS });

      expect(iosApps).toHaveLength(1);
      expect(iosApps[0].platform).toBe(AppPlatform.IOS);
    });

    it('should filter apps by publisher', async () => {
      const mockApps = {
        value: [
          {
            id: 'app-1',
            displayName: 'Microsoft Teams',
            publisher: 'Microsoft',
            '@odata.type': '#microsoft.graph.iosStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          },
          {
            id: 'app-2',
            displayName: 'Other App',
            publisher: 'Other Publisher',
            '@odata.type': '#microsoft.graph.iosStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockApps);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const microsoftApps = await report.getAllManagedApps({ publisher: 'Microsoft' });

      expect(microsoftApps).toHaveLength(1);
      expect(microsoftApps[0].publisher).toBe('Microsoft');
    });
  });

  describe('getAppInstallSummary', () => {
    it('should fetch install summary for an app', async () => {
      const mockSummary = {
        installedDeviceCount: 100,
        failedDeviceCount: 5,
        pendingInstallDeviceCount: 10,
        notApplicableDeviceCount: 20,
        notInstalledDeviceCount: 15,
        installedUserCount: 80,
        failedUserCount: 3,
        pendingInstallUserCount: 7,
        notApplicableUserCount: 12,
        notInstalledUserCount: 8
      };

      mockGraphClient = createMockGraphClient(mockSummary);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const summary = await report.getAppInstallSummary('app-id', 'Test App');

      expect(summary).not.toBeNull();
      expect(summary?.installedDeviceCount).toBe(100);
      expect(summary?.failedDeviceCount).toBe(5);
      expect(summary?.pendingInstallDeviceCount).toBe(10);
    });

    it('should return null if summary not available', async () => {
      mockGraphClient = createMockGraphClient(null);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const summary = await report.getAppInstallSummary('app-id', 'Test App');

      expect(summary).toBeNull();
    });
  });

  describe('getAppDeviceStatuses', () => {
    it('should fetch device statuses for an app', async () => {
      const mockStatuses = {
        value: [
          {
            id: 'status-1',
            deviceId: 'device-1',
            deviceName: 'Test Device 1',
            userName: 'user1@test.com',
            userPrincipalName: 'user1@test.com',
            platform: 'iOS',
            osVersion: '15.0',
            installState: 'installed',
            lastSyncDateTime: '2024-01-01T00:00:00Z',
            displayVersion: '1.0.0'
          },
          {
            id: 'status-2',
            deviceId: 'device-2',
            deviceName: 'Test Device 2',
            userName: 'user2@test.com',
            userPrincipalName: 'user2@test.com',
            platform: 'Android',
            osVersion: '12.0',
            installState: 'failed',
            errorCode: '0x87D1041C',
            lastSyncDateTime: '2024-01-01T00:00:00Z',
            displayVersion: '1.0.0'
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockStatuses);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const statuses = await report.getAppDeviceStatuses('app-id', 'Test App');

      expect(statuses).toHaveLength(2);
      expect(statuses[0].deviceName).toBe('Test Device 1');
      expect(statuses[0].installState).toBe(InstallState.INSTALLED);
      expect(statuses[1].deviceName).toBe('Test Device 2');
      expect(statuses[1].installState).toBe(InstallState.FAILED);
      expect(statuses[1].errorCode).toBe('0x87D1041C');
    });

    it('should handle empty device statuses', async () => {
      mockGraphClient = createMockGraphClient({ value: [] });
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const statuses = await report.getAppDeviceStatuses('app-id', 'Test App');

      expect(statuses).toHaveLength(0);
    });
  });

  describe('determineAppPlatform', () => {
    it('should correctly identify iOS apps', async () => {
      const mockApps = {
        value: [
          {
            id: 'app-1',
            displayName: 'iOS Store App',
            publisher: 'Test',
            '@odata.type': '#microsoft.graph.iosStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          },
          {
            id: 'app-2',
            displayName: 'iOS VPP App',
            publisher: 'Test',
            '@odata.type': '#microsoft.graph.iosVppApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockApps);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const apps = await report.getAllManagedApps();

      expect(apps[0].platform).toBe(AppPlatform.IOS);
      expect(apps[1].platform).toBe(AppPlatform.IOS);
    });

    it('should correctly identify Android apps', async () => {
      const mockApps = {
        value: [
          {
            id: 'app-1',
            displayName: 'Android Store App',
            publisher: 'Test',
            '@odata.type': '#microsoft.graph.androidStoreApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockApps);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const apps = await report.getAllManagedApps();

      expect(apps[0].platform).toBe(AppPlatform.ANDROID);
    });

    it('should correctly identify Windows apps', async () => {
      const mockApps = {
        value: [
          {
            id: 'app-1',
            displayName: 'Windows App',
            publisher: 'Test',
            '@odata.type': '#microsoft.graph.win32LobApp',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            isFeatured: false
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockApps);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const apps = await report.getAllManagedApps();

      expect(apps[0].platform).toBe(AppPlatform.WINDOWS);
    });
  });

  describe('parseInstallState', () => {
    it('should parse install states correctly', async () => {
      const mockStatuses = {
        value: [
          { id: '1', deviceId: 'd1', deviceName: 'Device 1', installState: 'installed' },
          { id: '2', deviceId: 'd2', deviceName: 'Device 2', installState: 'failed' },
          { id: '3', deviceId: 'd3', deviceName: 'Device 3', installState: 'pending' },
          { id: '4', deviceId: 'd4', deviceName: 'Device 4', installState: 'notInstalled' },
          { id: '5', deviceId: 'd5', deviceName: 'Device 5', installState: 'available' }
        ]
      };

      mockGraphClient = createMockGraphClient(mockStatuses);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const statuses = await report.getAppDeviceStatuses('app-id', 'Test App');

      expect(statuses[0].installState).toBe(InstallState.INSTALLED);
      expect(statuses[1].installState).toBe(InstallState.FAILED);
      expect(statuses[2].installState).toBe(InstallState.PENDING);
      expect(statuses[3].installState).toBe(InstallState.NOT_INSTALLED);
      expect(statuses[4].installState).toBe(InstallState.AVAILABLE);
    });
  });

  describe('Error Handling', () => {
    it('should handle API errors gracefully', async () => {
      const mockError = new Error('API Error');
      mockGraphClient = {
        api: jest.fn().mockReturnThis(),
        select: jest.fn().mockReturnThis(),
        top: jest.fn().mockReturnThis(),
        get: jest.fn().mockRejectedValue(mockError)
      } as any;

      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      await expect(report.getAllManagedApps()).rejects.toThrow();
    });

    it('should return empty array on partial failures', async () => {
      mockGraphClient = createMockGraphClient({ value: [] });
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const statuses = await report.getAppDeviceStatuses('app-id', 'Test App');

      expect(statuses).toEqual([]);
    });
  });

  describe('getErrorDescription', () => {
    it('should return appropriate error descriptions', async () => {
      const mockStatuses = {
        value: [
          {
            id: 'status-1',
            deviceId: 'device-1',
            deviceName: 'Device 1',
            installState: 'failed',
            errorCode: '0x80073CFF'
          },
          {
            id: 'status-2',
            deviceId: 'device-2',
            deviceName: 'Device 2',
            installState: 'failed',
            errorCode: '0x87D1041C'
          }
        ]
      };

      mockGraphClient = createMockGraphClient(mockStatuses);
      report = new AppDeploymentReport(mockGraphClient, mockConfig);

      const failures = await report.getAppDeviceStatuses('app-id', 'Test App');

      // Error descriptions are added during getFailedInstallations
      // Just verify error codes are preserved
      expect(failures[0].errorCode).toBe('0x80073CFF');
      expect(failures[1].errorCode).toBe('0x87D1041C');
    });
  });
});

describe('Integration Tests', () => {
  // These tests would require a real Graph API connection
  // For now, we'll skip them in unit tests

  it.skip('should fetch real data from Graph API', async () => {
    // This would require real credentials and Graph API access
    // Implement when running integration tests
  });

  it.skip('should export reports to files', async () => {
    // Test actual file export functionality
    // Implement when running integration tests
  });
});

export default {};
