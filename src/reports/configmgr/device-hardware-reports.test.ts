/**
 * Device Hardware Reports - Unit Tests
 *
 * Comprehensive test suite for the DeviceHardwareReports module
 */

import { DeviceHardwareReports, DeviceHardwareInfo } from './device-hardware-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';

// ============================================================================
// Mock Data
// ============================================================================

const mockDevice1 = {
  id: 'device-001',
  deviceName: 'iPhone 14 Pro',
  manufacturer: 'Apple Inc.',
  model: 'iPhone 14 Pro',
  serialNumber: 'SN001',
  userPrincipalName: 'user1@contoso.com',
  enrolledDateTime: '2024-01-15T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:00:00Z',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  totalStorageSpaceInBytes: 256 * 1024 * 1024 * 1024, // 256 GB
  freeStorageSpaceInBytes: 128 * 1024 * 1024 * 1024, // 128 GB free
  physicalMemoryInBytes: 6 * 1024 * 1024 * 1024, // 6 GB RAM
  managementAgent: 'mdm',
  complianceState: 'compliant',
};

const mockDevice2 = {
  id: 'device-002',
  deviceName: 'Samsung Galaxy S23',
  manufacturer: 'Samsung',
  model: 'Galaxy S23',
  serialNumber: 'SN002',
  userPrincipalName: 'user2@contoso.com',
  enrolledDateTime: '2024-01-20T10:00:00Z',
  lastSyncDateTime: '2024-02-06T07:30:00Z',
  operatingSystem: 'Android',
  osVersion: '14.0',
  totalStorageSpaceInBytes: 128 * 1024 * 1024 * 1024, // 128 GB
  freeStorageSpaceInBytes: 5 * 1024 * 1024 * 1024, // 5 GB free (low storage)
  physicalMemoryInBytes: 8 * 1024 * 1024 * 1024, // 8 GB RAM
  managementAgent: 'mdm',
  complianceState: 'compliant',
};

const mockDevice3 = {
  id: 'device-003',
  deviceName: 'iPad Pro',
  manufacturer: 'Apple Inc.',
  model: 'iPad Pro 12.9-inch',
  serialNumber: 'SN003',
  userPrincipalName: 'user3@contoso.com',
  enrolledDateTime: '2024-02-01T10:00:00Z',
  lastSyncDateTime: '2024-02-06T09:00:00Z',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  totalStorageSpaceInBytes: 512 * 1024 * 1024 * 1024, // 512 GB
  freeStorageSpaceInBytes: 300 * 1024 * 1024 * 1024, // 300 GB free
  physicalMemoryInBytes: 16 * 1024 * 1024 * 1024, // 16 GB RAM
  managementAgent: 'mdm',
  complianceState: 'compliant',
};

const mockDevice4 = {
  id: 'device-004',
  deviceName: 'Pixel 7',
  manufacturer: 'Google',
  model: 'Pixel 7',
  serialNumber: 'SN004',
  userPrincipalName: 'user4@contoso.com',
  enrolledDateTime: '2024-01-10T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:45:00Z',
  operatingSystem: 'Android',
  osVersion: '14.0',
  totalStorageSpaceInBytes: 128 * 1024 * 1024 * 1024, // 128 GB
  freeStorageSpaceInBytes: 80 * 1024 * 1024 * 1024, // 80 GB free
  physicalMemoryInBytes: 8 * 1024 * 1024 * 1024, // 8 GB RAM
  managementAgent: 'mdm',
  complianceState: 'noncompliant',
};

const mockDevice5 = {
  id: 'device-005',
  deviceName: 'Surface Pro 9',
  manufacturer: 'Microsoft Corporation',
  model: 'Surface Pro 9',
  serialNumber: 'SN005',
  userPrincipalName: 'user5@contoso.com',
  enrolledDateTime: '2024-01-25T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:15:00Z',
  operatingSystem: 'Windows',
  osVersion: '11.0.22621',
  totalStorageSpaceInBytes: 256 * 1024 * 1024 * 1024, // 256 GB
  freeStorageSpaceInBytes: 100 * 1024 * 1024 * 1024, // 100 GB free
  physicalMemoryInBytes: 16 * 1024 * 1024 * 1024, // 16 GB RAM
  managementAgent: 'mdm',
  complianceState: 'compliant',
};

// ============================================================================
// Mock Graph Client
// ============================================================================

class MockGraphClient {
  private responses: Map<string, any> = new Map();

  constructor() {
    this.setupDefaultResponses();
  }

  setupDefaultResponses() {
    // Devices endpoint
    this.responses.set('/deviceManagement/managedDevices', {
      value: [mockDevice1, mockDevice2, mockDevice3, mockDevice4, mockDevice5],
    });
  }

  api(endpoint: string) {
    const mockApi = {
      endpoint,
      selectFields: [] as string[],
      topValue: 999,

      select(fields: string[] | string) {
        mockApi.selectFields = Array.isArray(fields) ? fields : [fields];
        return mockApi;
      },

      top(value: number) {
        mockApi.topValue = value;
        return mockApi;
      },

      async get() {
        const response = this.responses.get(mockApi.endpoint);
        if (response) {
          return response;
        }

        // Default response
        return { value: [] };
      },
    };

    return mockApi;
  }
}

// ============================================================================
// Test Configuration
// ============================================================================

function getMockConfig(): AppConfig {
  return {
    authentication: {
      tenantId: 'test-tenant',
      clientId: 'test-client',
      clientSecret: 'test-secret',
      authMethod: 'clientSecret',
    },
    reports: {
      enabled: ['device-hardware-reports'],
      disabled: [],
      settings: {},
    },
    output: {
      defaultFormat: 'json',
      directory: './test-reports',
      includeTimestamp: false,
      compression: false,
    },
    scheduler: {
      enabled: false,
      timezone: 'UTC',
      schedules: [],
    },
    logging: {
      level: 'error',
      file: './test-logs/app.log',
      console: false,
      maxSize: '10m',
      maxFiles: 5,
    },
  };
}

// ============================================================================
// Test Suite
// ============================================================================

describe('DeviceHardwareReports', () => {
  let report: DeviceHardwareReports;
  let mockClient: any;
  let config: AppConfig;

  beforeEach(() => {
    mockClient = new MockGraphClient();
    config = getMockConfig();
    report = new DeviceHardwareReports(mockClient as unknown as Client, config);
  });

  // ==========================================================================
  // Test: getAllDeviceHardwareInfo
  // ==========================================================================

  describe('getAllDeviceHardwareInfo', () => {
    it('should retrieve all device hardware information', async () => {
      const devices = await report.getAllDeviceHardwareInfo();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBe(5);
    });

    it('should parse hardware data correctly', async () => {
      const devices = await report.getAllDeviceHardwareInfo();
      const device = devices[0];

      expect(device.deviceId).toBe(mockDevice1.id);
      expect(device.deviceName).toBe(mockDevice1.deviceName);
      expect(device.manufacturer).toBe(mockDevice1.manufacturer);
      expect(device.model).toBe(mockDevice1.model);
      expect(device.serialNumber).toBe(mockDevice1.serialNumber);
    });

    it('should correctly calculate memory in different units', async () => {
      const devices = await report.getAllDeviceHardwareInfo();
      const device = devices[0]; // iPhone 14 Pro with 6GB RAM

      expect(device.memory.totalMemoryInBytes).toBe(6 * 1024 * 1024 * 1024);
      expect(device.memory.totalMemoryInMB).toBe(6144);
      expect(device.memory.totalMemoryInGB).toBe(6);
    });

    it('should correctly calculate storage in different units', async () => {
      const devices = await report.getAllDeviceHardwareInfo();
      const device = devices[0]; // iPhone 14 Pro with 256GB storage

      expect(device.storage.totalStorageInBytes).toBe(256 * 1024 * 1024 * 1024);
      expect(device.storage.totalStorageInMB).toBe(262144);
      expect(device.storage.totalStorageInGB).toBe(256);
      expect(device.storage.freeStorageInGB).toBe(128);
      expect(device.storage.freeStoragePercentage).toBe(50);
      expect(device.storage.usedStoragePercentage).toBe(50);
    });

    it('should filter devices by operating system', async () => {
      const devices = await report.getAllDeviceHardwareInfo({
        operatingSystem: 'iOS',
      });

      expect(devices.length).toBe(2); // iPhone and iPad
      expect(devices.every((d) => d.operatingSystem.operatingSystem === 'iOS')).toBe(true);
    });

    it('should filter devices by manufacturer', async () => {
      const devices = await report.getAllDeviceHardwareInfo({
        manufacturer: 'Apple',
      });

      expect(devices.length).toBe(2);
      expect(devices.every((d) => d.manufacturer.includes('Apple'))).toBe(true);
    });

    it('should filter devices by compliance state', async () => {
      const devices = await report.getAllDeviceHardwareInfo({
        complianceState: 'compliant',
      });

      expect(devices.length).toBe(4);
      expect(devices.every((d) => d.complianceState === 'compliant')).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getDisplayConfigurationSummary (Report 9)
  // ==========================================================================

  describe('getDisplayConfigurationSummary', () => {
    it('should generate display configuration summary', async () => {
      const summary = await report.getDisplayConfigurationSummary();

      expect(summary).toBeDefined();
      expect(Array.isArray(summary)).toBe(true);
      expect(summary.length).toBeGreaterThan(0);
    });

    it('should categorize devices by display configuration', async () => {
      const summary = await report.getDisplayConfigurationSummary();

      expect(summary.every((s) => s.displayResolution)).toBeTruthy();
      expect(summary.every((s) => s.deviceCount > 0)).toBe(true);
    });

    it('should calculate percentages correctly', async () => {
      const summary = await report.getDisplayConfigurationSummary();
      const totalPercentage = summary.reduce((sum, s) => sum + s.percentage, 0);

      expect(totalPercentage).toBeCloseTo(100, 1);
    });

    it('should include device names in results', async () => {
      const summary = await report.getDisplayConfigurationSummary();

      expect(summary.every((s) => Array.isArray(s.devices))).toBe(true);
      expect(summary.every((s) => s.devices.length === s.deviceCount)).toBe(true);
    });

    it('should sort by device count descending', async () => {
      const summary = await report.getDisplayConfigurationSummary();

      for (let i = 1; i < summary.length; i++) {
        expect(summary[i - 1].deviceCount).toBeGreaterThanOrEqual(summary[i].deviceCount);
      }
    });
  });

  // ==========================================================================
  // Test: getOSDistributionSummary (Report 10)
  // ==========================================================================

  describe('getOSDistributionSummary', () => {
    it('should generate OS distribution summary', async () => {
      const summary = await report.getOSDistributionSummary();

      expect(summary).toBeDefined();
      expect(Array.isArray(summary)).toBe(true);
      expect(summary.length).toBeGreaterThan(0);
    });

    it('should group devices by OS and version', async () => {
      const summary = await report.getOSDistributionSummary();

      expect(summary.every((s) => s.operatingSystem)).toBeTruthy();
      expect(summary.every((s) => s.osVersion)).toBeTruthy();
    });

    it('should calculate device counts correctly', async () => {
      const summary = await report.getOSDistributionSummary();
      const totalDevices = summary.reduce((sum, s) => sum + s.deviceCount, 0);

      expect(totalDevices).toBe(5); // Total mock devices
    });

    it('should calculate percentages correctly', async () => {
      const summary = await report.getOSDistributionSummary();
      const totalPercentage = summary.reduce((sum, s) => sum + s.percentage, 0);

      expect(totalPercentage).toBeCloseTo(100, 1);
    });

    it('should identify different OS platforms', async () => {
      const summary = await report.getOSDistributionSummary();
      const osList = summary.map((s) => s.operatingSystem);

      expect(osList).toContain('iOS');
      expect(osList).toContain('Android');
      expect(osList).toContain('Windows');
    });
  });

  // ==========================================================================
  // Test: getMemoryRangeSummary (Report 11)
  // ==========================================================================

  describe('getMemoryRangeSummary', () => {
    it('should generate memory range summary', async () => {
      const summary = await report.getMemoryRangeSummary();

      expect(summary).toBeDefined();
      expect(Array.isArray(summary)).toBe(true);
    });

    it('should categorize devices into memory ranges', async () => {
      const summary = await report.getMemoryRangeSummary();

      expect(summary.every((s) => s.range)).toBeTruthy();
      expect(summary.every((s) => s.rangeInMB)).toBeTruthy();
    });

    it('should have correct device count per range', async () => {
      const summary = await report.getMemoryRangeSummary();
      const totalDevices = summary.reduce((sum, s) => sum + s.deviceCount, 0);

      expect(totalDevices).toBe(5); // Total mock devices
    });

    it('should include all standard memory ranges', async () => {
      const summary = await report.getMemoryRangeSummary();
      const ranges = summary.map((s) => s.range);

      expect(ranges).toContain('6-8 GB');
      expect(ranges).toContain('8-16 GB');
    });

    it('should calculate percentages correctly', async () => {
      const summary = await report.getMemoryRangeSummary();

      summary.forEach((s) => {
        expect(s.percentage).toBeGreaterThanOrEqual(0);
        expect(s.percentage).toBeLessThanOrEqual(100);
      });
    });
  });

  // ==========================================================================
  // Test: getStorageRangeSummary (Report 12)
  // ==========================================================================

  describe('getStorageRangeSummary', () => {
    it('should generate storage range summary', async () => {
      const summary = await report.getStorageRangeSummary();

      expect(summary).toBeDefined();
      expect(Array.isArray(summary)).toBe(true);
    });

    it('should categorize devices into storage ranges', async () => {
      const summary = await report.getStorageRangeSummary();

      expect(summary.every((s) => s.range)).toBeTruthy();
      expect(summary.every((s) => s.rangeInGB)).toBeTruthy();
    });

    it('should have correct device count per range', async () => {
      const summary = await report.getStorageRangeSummary();
      const totalDevices = summary.reduce((sum, s) => sum + s.deviceCount, 0);

      expect(totalDevices).toBe(5); // Total mock devices
    });

    it('should include standard storage ranges', async () => {
      const summary = await report.getStorageRangeSummary();
      const ranges = summary.map((s) => s.range);

      expect(ranges).toContain('128-256 GB');
      expect(ranges).toContain('256-512 GB');
      expect(ranges).toContain('512 GB - 1 TB');
    });

    it('should include device names in ranges', async () => {
      const summary = await report.getStorageRangeSummary();

      expect(summary.every((s) => Array.isArray(s.devices))).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getDevicesWithSpecificFreeMemory (Report 25)
  // ==========================================================================

  describe('getDevicesWithSpecificFreeMemory', () => {
    it('should find devices with memory above minimum threshold', async () => {
      const devices = await report.getDevicesWithSpecificFreeMemory(1000);

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBeGreaterThan(0);
    });

    it('should find devices within memory range', async () => {
      const devices = await report.getDevicesWithSpecificFreeMemory(2000, 3000);

      expect(devices).toBeDefined();
      devices.forEach((d) => {
        expect(d.freeMemoryInMB).toBeGreaterThanOrEqual(2000);
        expect(d.freeMemoryInMB).toBeLessThanOrEqual(3000);
      });
    });

    it('should include all required device information', async () => {
      const devices = await report.getDevicesWithSpecificFreeMemory(1000);

      if (devices.length > 0) {
        const device = devices[0];
        expect(device.deviceId).toBeDefined();
        expect(device.deviceName).toBeDefined();
        expect(device.manufacturer).toBeDefined();
        expect(device.model).toBeDefined();
        expect(device.totalMemoryInMB).toBeDefined();
        expect(device.freeMemoryInMB).toBeDefined();
        expect(device.freeMemoryPercentage).toBeDefined();
      }
    });
  });

  // ==========================================================================
  // Test: getDevicesWithSpecificFreeStorage (Report 26)
  // ==========================================================================

  describe('getDevicesWithSpecificFreeStorage', () => {
    it('should find devices with storage above minimum threshold', async () => {
      const devices = await report.getDevicesWithSpecificFreeStorage(50);

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBeGreaterThan(0);
    });

    it('should find devices within storage range', async () => {
      const devices = await report.getDevicesWithSpecificFreeStorage(50, 200);

      devices.forEach((d) => {
        expect(d.freeStorageInGB).toBeGreaterThanOrEqual(50);
        expect(d.freeStorageInGB).toBeLessThanOrEqual(200);
      });
    });

    it('should include storage calculations', async () => {
      const devices = await report.getDevicesWithSpecificFreeStorage(50);

      if (devices.length > 0) {
        const device = devices[0];
        expect(device.totalStorageInGB).toBeDefined();
        expect(device.freeStorageInGB).toBeDefined();
        expect(device.usedStorageInGB).toBeDefined();
        expect(device.freeStoragePercentage).toBeDefined();
      }
    });

    it('should calculate used storage correctly', async () => {
      const devices = await report.getDevicesWithSpecificFreeStorage(0);

      devices.forEach((d) => {
        expect(d.usedStorageInGB).toBe(d.totalStorageInGB - d.freeStorageInGB);
      });
    });
  });

  // ==========================================================================
  // Test: getDevicesWithLowMemory (Report 28)
  // ==========================================================================

  describe('getDevicesWithLowMemory', () => {
    it('should find devices with low free memory', async () => {
      const devices = await report.getDevicesWithLowMemory(512);

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
    });

    it('should only include devices below threshold', async () => {
      const threshold = 512;
      const devices = await report.getDevicesWithLowMemory(threshold);

      devices.forEach((d) => {
        expect(d.freeMemoryInMB).toBeLessThanOrEqual(threshold);
      });
    });

    it('should use default threshold if not specified', async () => {
      const devices = await report.getDevicesWithLowMemory();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getDevicesWithLowStorage (Report 29)
  // ==========================================================================

  describe('getDevicesWithLowStorage', () => {
    it('should find devices with low free storage', async () => {
      const devices = await report.getDevicesWithLowStorage(10);

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBeGreaterThan(0); // mockDevice2 has 5GB free
    });

    it('should only include devices below threshold', async () => {
      const threshold = 10;
      const devices = await report.getDevicesWithLowStorage(threshold);

      devices.forEach((d) => {
        expect(d.freeStorageInGB).toBeLessThan(threshold);
      });
    });

    it('should sort devices by free storage ascending', async () => {
      const devices = await report.getDevicesWithLowStorage(100);

      for (let i = 1; i < devices.length; i++) {
        expect(devices[i].freeStorageInGB).toBeGreaterThanOrEqual(
          devices[i - 1].freeStorageInGB
        );
      }
    });

    it('should identify the device with lowest storage', async () => {
      const devices = await report.getDevicesWithLowStorage(100);

      if (devices.length > 0) {
        const lowestDevice = devices[0];
        expect(lowestDevice.deviceName).toBe('Samsung Galaxy S23'); // Has 5GB free
        expect(lowestDevice.freeStorageInGB).toBe(5);
      }
    });

    it('should use default threshold if not specified', async () => {
      const devices = await report.getDevicesWithLowStorage();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
    });
  });

  // ==========================================================================
  // Test: generateComprehensiveReport
  // ==========================================================================

  describe('generateComprehensiveReport', () => {
    it('should generate complete hardware report', async () => {
      const reportData = await report.generateComprehensiveReport();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include all summary sections', async () => {
      const reportData = await report.generateComprehensiveReport();
      const summary = reportData.summary;

      expect(summary.totalDevices).toBeDefined();
      expect(summary.displayConfigurations).toBeDefined();
      expect(summary.osDistribution).toBeDefined();
      expect(summary.memoryRanges).toBeDefined();
      expect(summary.storageRanges).toBeDefined();
      expect(summary.lowMemoryDeviceCount).toBeDefined();
      expect(summary.lowStorageDeviceCount).toBeDefined();
    });

    it('should calculate average statistics', async () => {
      const reportData = await report.generateComprehensiveReport();
      const summary = reportData.summary;

      expect(summary.averageMemoryGB).toBeDefined();
      expect(summary.averageStorageGB).toBeDefined();
      expect(summary.averageFreeStoragePercentage).toBeDefined();

      expect(summary.averageMemoryGB).toBeGreaterThan(0);
      expect(summary.averageStorageGB).toBeGreaterThan(0);
      expect(summary.averageFreeStoragePercentage).toBeGreaterThanOrEqual(0);
      expect(summary.averageFreeStoragePercentage).toBeLessThanOrEqual(100);
    });

    it('should include top models by memory and storage', async () => {
      const reportData = await report.generateComprehensiveReport();
      const summary = reportData.summary;

      expect(summary.topMemoryModels).toBeDefined();
      expect(Array.isArray(summary.topMemoryModels)).toBe(true);
      expect(summary.topStorageModels).toBeDefined();
      expect(Array.isArray(summary.topStorageModels)).toBe(true);
    });

    it('should respect filter options', async () => {
      const reportData = await report.generateComprehensiveReport({
        operatingSystem: 'iOS',
      });

      expect(reportData.data.length).toBe(2); // Only iOS devices
    });
  });

  // ==========================================================================
  // Test: execute (Full Report)
  // ==========================================================================

  describe('execute', () => {
    it('should generate complete hardware report', async () => {
      const reportData = await report.execute();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include metadata in report', async () => {
      const reportData = await report.execute();
      const metadata = reportData.metadata;

      expect(metadata.reportName).toBe('device-hardware-reports');
      expect(metadata.generatedAt).toBeDefined();
      expect(metadata.generatedBy).toBe('Intune Reporting Dashboard');
      expect(metadata.recordCount).toBe(5);
    });

    it('should include summary statistics', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.totalDevices).toBe(5);
      expect(summary.displayConfigurations).toBeDefined();
      expect(summary.osDistribution).toBeDefined();
      expect(summary.memoryRanges).toBeDefined();
      expect(summary.storageRanges).toBeDefined();
    });
  });

  // ==========================================================================
  // Test: Error Handling
  // ==========================================================================

  describe('Error Handling', () => {
    it('should handle API errors gracefully', async () => {
      const errorClient = {
        api: () => ({
          select: () => ({
            top: () => ({
              get: async () => {
                throw new Error('API Error');
              },
            }),
          }),
        }),
      };

      const errorReport = new DeviceHardwareReports(
        errorClient as unknown as Client,
        config
      );

      await expect(errorReport.getAllDeviceHardwareInfo()).rejects.toThrow();
    });

    it('should handle empty device list', async () => {
      mockClient.responses.set('/deviceManagement/managedDevices', { value: [] });

      const devices = await report.getAllDeviceHardwareInfo();

      expect(devices.length).toBe(0);
    });

    it('should handle devices with missing optional fields', async () => {
      const minimalDevice = {
        id: 'device-999',
        deviceName: 'Minimal Device',
        operatingSystem: 'iOS',
        totalStorageSpaceInBytes: 64 * 1024 * 1024 * 1024,
        freeStorageSpaceInBytes: 32 * 1024 * 1024 * 1024,
        physicalMemoryInBytes: 4 * 1024 * 1024 * 1024,
        managementAgent: 'mdm',
        complianceState: 'compliant',
      };

      mockClient.responses.set('/deviceManagement/managedDevices', {
        value: [minimalDevice],
      });

      const devices = await report.getAllDeviceHardwareInfo();

      expect(devices.length).toBe(1);
      expect(devices[0].manufacturer).toBe('Unknown');
      expect(devices[0].serialNumber).toBe('N/A');
    });
  });
});

// ============================================================================
// Run Tests
// ============================================================================

// Export for test runner
export default describe;
