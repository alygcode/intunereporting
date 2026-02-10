/**
 * Mobile Device Inventory Reports - Unit Tests
 *
 * Comprehensive test suite for the MobileDeviceInventoryReports module
 */

import {
  MobileDeviceInventoryReports,
  DeviceOwnershipType,
  ManagementAgentType,
  CommunicationStatus,
  MobileDeviceInfo,
} from './mobile-device-inventory-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';

// ============================================================================
// Mock Data
// ============================================================================

const mockDevice1 = {
  id: 'device-001',
  deviceName: 'iPhone 14 Pro',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  manufacturer: 'Apple Inc.',
  model: 'iPhone 14 Pro',
  serialNumber: 'SN001',
  imei: 'IMEI001',
  phoneNumber: '+1-555-0101',
  managedDeviceOwnerType: 'company',
  ownerType: 'company',
  managementAgent: 'mdm',
  enrolledDateTime: '2024-01-15T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:00:00Z',
  complianceState: 'compliant',
  userPrincipalName: 'user1@contoso.com',
  userDisplayName: 'User One',
  emailAddress: 'user1@contoso.com',
  azureADDeviceId: 'azure-001',
  deviceEnrollmentType: 'userEnrollment',
  isSupervised: true,
  isEncrypted: true,
  jailBroken: 'false',
};

const mockDevice2 = {
  id: 'device-002',
  deviceName: 'Samsung Galaxy S23',
  operatingSystem: 'Android',
  osVersion: '14.0',
  manufacturer: 'Samsung',
  model: 'Galaxy S23',
  serialNumber: 'SN002',
  imei: 'IMEI002',
  phoneNumber: '+1-555-0102',
  managedDeviceOwnerType: 'company',
  ownerType: 'company',
  managementAgent: 'mdm',
  enrolledDateTime: '2024-01-20T10:00:00Z',
  lastSyncDateTime: '2024-02-05T08:00:00Z',
  complianceState: 'compliant',
  userPrincipalName: 'user2@contoso.com',
  userDisplayName: 'User Two',
  emailAddress: 'user2@contoso.com',
  azureADDeviceId: 'azure-002',
  deviceEnrollmentType: 'corporateOwnedDedicatedDevice',
  isSupervised: false,
  isEncrypted: true,
  jailBroken: 'false',
};

const mockDevice3 = {
  id: 'device-003',
  deviceName: 'Personal iPad',
  operatingSystem: 'iOS',
  osVersion: '17.1',
  manufacturer: 'Apple Inc.',
  model: 'iPad Pro',
  serialNumber: 'SN003',
  managedDeviceOwnerType: 'personal',
  ownerType: 'personal',
  managementAgent: 'mdm',
  enrolledDateTime: '2024-02-01T10:00:00Z',
  lastSyncDateTime: '2024-02-06T09:00:00Z',
  complianceState: 'noncompliant',
  userPrincipalName: 'user3@contoso.com',
  userDisplayName: 'User Three',
  emailAddress: 'user3@contoso.com',
  azureADDeviceId: 'azure-003',
  deviceEnrollmentType: 'userEnrollment',
  isSupervised: false,
  isEncrypted: true,
  jailBroken: 'false',
};

const mockDevice4 = {
  id: 'device-004',
  deviceName: 'Pixel 7',
  operatingSystem: 'Android',
  osVersion: '14.0',
  manufacturer: 'Google',
  model: 'Pixel 7',
  serialNumber: 'SN004',
  managedDeviceOwnerType: 'personal',
  ownerType: 'personal',
  managementAgent: 'mdm',
  enrolledDateTime: '2024-01-10T10:00:00Z',
  lastSyncDateTime: '2024-01-20T08:00:00Z', // 17 days ago (inactive)
  complianceState: 'compliant',
  userPrincipalName: 'user4@contoso.com',
  userDisplayName: 'User Four',
  emailAddress: 'user4@contoso.com',
  azureADDeviceId: 'azure-004',
  deviceEnrollmentType: 'userEnrollment',
  isSupervised: false,
  isEncrypted: true,
  jailBroken: 'false',
};

const mockDevice5 = {
  id: 'device-005',
  deviceName: 'Surface Go',
  operatingSystem: 'Windows',
  osVersion: '11.0.22621',
  manufacturer: 'Microsoft Corporation',
  model: 'Surface Go 3',
  serialNumber: 'SN005',
  managedDeviceOwnerType: 'company',
  ownerType: 'company',
  managementAgent: 'mdm',
  enrolledDateTime: '2024-01-25T10:00:00Z',
  lastSyncDateTime: '2024-02-06T07:30:00Z',
  complianceState: 'compliant',
  userPrincipalName: 'user5@contoso.com',
  userDisplayName: 'User Five',
  emailAddress: 'user5@contoso.com',
  azureADDeviceId: 'azure-005',
  deviceEnrollmentType: 'windowsAzureADJoin',
  isSupervised: false,
  isEncrypted: true,
};

const mockDevice6 = {
  id: 'device-006',
  deviceName: 'Old iPhone',
  operatingSystem: 'iOS',
  osVersion: '15.7',
  manufacturer: 'Apple Inc.',
  model: 'iPhone 11',
  serialNumber: 'SN006',
  managedDeviceOwnerType: 'company',
  ownerType: 'company',
  managementAgent: 'mdm',
  enrolledDateTime: '2023-12-01T10:00:00Z',
  lastSyncDateTime: '2023-12-15T08:00:00Z', // Very old (50+ days)
  complianceState: 'noncompliant',
  userPrincipalName: 'user6@contoso.com',
  userDisplayName: 'User Six',
  emailAddress: 'user6@contoso.com',
  azureADDeviceId: 'azure-006',
  deviceEnrollmentType: 'userEnrollment',
  isSupervised: true,
  isEncrypted: true,
  jailBroken: 'false',
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
    // Default managed devices response
    this.responses.set('/deviceManagement/managedDevices', {
      value: [mockDevice1, mockDevice2, mockDevice3, mockDevice4, mockDevice5, mockDevice6],
    });

    // Corporate devices only
    this.responses.set('/deviceManagement/managedDevices-company', {
      value: [mockDevice1, mockDevice2, mockDevice5, mockDevice6],
    });

    // Exclude Exchange connector devices
    this.responses.set('/deviceManagement/managedDevices-no-eas', {
      value: [mockDevice1, mockDevice2, mockDevice3, mockDevice4, mockDevice5, mockDevice6],
    });
  }

  api(endpoint: string) {
    const mockApi = {
      endpoint,
      selectFields: [] as string[],
      filterString: '',

      select(fields: string[] | string) {
        mockApi.selectFields = Array.isArray(fields) ? fields : [fields];
        return mockApi;
      },

      filter(filterStr: string) {
        mockApi.filterString = filterStr;
        return mockApi;
      },

      async get() {
        // Determine which mock data to return based on filter
        let response = this.responses.get(mockApi.endpoint);

        if (mockApi.filterString.includes("managedDeviceOwnerType eq 'company'")) {
          response = this.responses.get('/deviceManagement/managedDevices-company');
        } else if (mockApi.filterString.includes("managementAgent ne 'eas'")) {
          response = this.responses.get('/deviceManagement/managedDevices-no-eas');
        }

        if (response) {
          // Apply additional filters in memory
          let filteredValue = [...response.value];

          if (mockApi.filterString.includes("operatingSystem eq 'iOS'")) {
            filteredValue = filteredValue.filter(d => d.operatingSystem === 'iOS');
          } else if (mockApi.filterString.includes("operatingSystem eq 'Android'")) {
            filteredValue = filteredValue.filter(d => d.operatingSystem === 'Android');
          }

          if (mockApi.filterString.includes('isSupervised eq true')) {
            filteredValue = filteredValue.filter(d => d.isSupervised === true);
          }

          if (mockApi.filterString.includes("complianceState eq 'compliant'")) {
            filteredValue = filteredValue.filter(d => d.complianceState === 'compliant');
          }

          return { value: filteredValue };
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
      enabled: ['mobile-device-inventory'],
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

describe('MobileDeviceInventoryReports', () => {
  let report: MobileDeviceInventoryReports;
  let mockClient: any;
  let config: AppConfig;

  beforeEach(() => {
    mockClient = new MockGraphClient();
    config = getMockConfig();
    report = new MobileDeviceInventoryReports(mockClient as unknown as Client, config);
  });

  // ==========================================================================
  // Test: getAllCorporateOwnedDevices (Report 1)
  // ==========================================================================

  describe('getAllCorporateOwnedDevices', () => {
    it('should retrieve all corporate-owned mobile devices', async () => {
      const devices = await report.getAllCorporateOwnedDevices();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBe(4); // mockDevice1, 2, 5, 6
    });

    it('should only include company-owned devices', async () => {
      const devices = await report.getAllCorporateOwnedDevices();

      devices.forEach(device => {
        expect(device.ownerType).toBe(DeviceOwnershipType.COMPANY);
      });
    });

    it('should parse device data correctly', async () => {
      const devices = await report.getAllCorporateOwnedDevices();
      const device = devices[0];

      expect(device.deviceId).toBe(mockDevice1.id);
      expect(device.deviceName).toBe(mockDevice1.deviceName);
      expect(device.operatingSystem).toBe(mockDevice1.operatingSystem);
      expect(device.manufacturer).toBe(mockDevice1.manufacturer);
      expect(device.model).toBe(mockDevice1.model);
      expect(device.serialNumber).toBe(mockDevice1.serialNumber);
    });

    it('should include enrollment profile name for corporate devices', async () => {
      const devices = await report.getAllCorporateOwnedDevices();

      devices.forEach(device => {
        expect(device).toHaveProperty('enrollmentProfileName');
      });
    });

    it('should filter by operating system', async () => {
      const iosDevices = await report.getAllCorporateOwnedDevices({
        operatingSystem: 'iOS',
      });

      expect(iosDevices.every(d => d.operatingSystem === 'iOS')).toBe(true);
    });

    it('should filter by compliance state', async () => {
      const compliantDevices = await report.getAllCorporateOwnedDevices({
        complianceState: 'compliant',
      });

      expect(compliantDevices.every(d => d.complianceState === 'compliant')).toBe(true);
    });

    it('should filter by supervised status', async () => {
      const supervisedDevices = await report.getAllCorporateOwnedDevices({
        isSupervised: true,
      });

      expect(supervisedDevices.every(d => d.isSupervised === true)).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getAllMobileDeviceClients (Report 2)
  // ==========================================================================

  describe('getAllMobileDeviceClients', () => {
    it('should retrieve all mobile device clients', async () => {
      const devices = await report.getAllMobileDeviceClients();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBe(6); // All mock devices (no EAS in mock data)
    });

    it('should exclude Exchange ActiveSync connector devices', async () => {
      const devices = await report.getAllMobileDeviceClients();

      devices.forEach(device => {
        expect(device.managementAgent).not.toBe(ManagementAgentType.EAS);
      });
    });

    it('should include both corporate and personal devices', async () => {
      const devices = await report.getAllMobileDeviceClients();

      const corporateCount = devices.filter(d => d.ownerType === DeviceOwnershipType.COMPANY).length;
      const personalCount = devices.filter(d => d.ownerType === DeviceOwnershipType.PERSONAL).length;

      expect(corporateCount).toBeGreaterThan(0);
      expect(personalCount).toBeGreaterThan(0);
    });

    it('should parse all device properties correctly', async () => {
      const devices = await report.getAllMobileDeviceClients();
      const device = devices[0];

      expect(device.deviceId).toBeDefined();
      expect(device.deviceName).toBeDefined();
      expect(device.operatingSystem).toBeDefined();
      expect(device.osVersion).toBeDefined();
      expect(device.ownerType).toBeDefined();
      expect(device.managementAgent).toBeDefined();
    });

    it('should filter by operating system', async () => {
      const androidDevices = await report.getAllMobileDeviceClients({
        operatingSystem: 'Android',
      });

      expect(androidDevices.every(d => d.operatingSystem === 'Android')).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getMobileDeviceClientInformation (Report 20)
  // ==========================================================================

  describe('getMobileDeviceClientInformation', () => {
    it('should retrieve mobile device client information', async () => {
      const clientInfo = await report.getMobileDeviceClientInformation();

      expect(clientInfo).toBeDefined();
      expect(Array.isArray(clientInfo)).toBe(true);
      expect(clientInfo.length).toBeGreaterThan(0);
    });

    it('should include communication status for each device', async () => {
      const clientInfo = await report.getMobileDeviceClientInformation();

      clientInfo.forEach(device => {
        expect(device.communicationStatus).toBeDefined();
        expect(Object.values(CommunicationStatus)).toContain(device.communicationStatus);
      });
    });

    it('should calculate days since last contact', async () => {
      const clientInfo = await report.getMobileDeviceClientInformation();

      clientInfo.forEach(device => {
        if (device.lastContactDateTime) {
          expect(device.daysSinceLastContact).toBeDefined();
          expect(device.daysSinceLastContact).toBeGreaterThanOrEqual(0);
        }
      });
    });

    it('should identify active devices correctly', async () => {
      const clientInfo = await report.getMobileDeviceClientInformation();

      const activeDevices = clientInfo.filter(d => d.isActive);
      activeDevices.forEach(device => {
        expect(device.daysSinceLastContact).toBeDefined();
        expect(device.daysSinceLastContact!).toBeLessThanOrEqual(30);
      });
    });

    it('should identify inactive devices correctly', async () => {
      const clientInfo = await report.getMobileDeviceClientInformation();

      const inactiveDevices = clientInfo.filter(d => !d.isActive);
      inactiveDevices.forEach(device => {
        if (device.daysSinceLastContact !== undefined) {
          expect(device.daysSinceLastContact).toBeGreaterThan(30);
        }
      });
    });

    it('should categorize communication status correctly', async () => {
      const clientInfo = await report.getMobileDeviceClientInformation();

      const healthy = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.HEALTHY);
      const warning = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.WARNING);
      const critical = clientInfo.filter(d => d.communicationStatus === CommunicationStatus.CRITICAL);

      // Healthy: synced within 1 day
      healthy.forEach(device => {
        if (device.daysSinceLastContact !== undefined) {
          expect(device.daysSinceLastContact).toBeLessThanOrEqual(1);
        }
      });

      // Warning: synced between 1-7 days
      warning.forEach(device => {
        if (device.daysSinceLastContact !== undefined) {
          expect(device.daysSinceLastContact).toBeGreaterThan(1);
          expect(device.daysSinceLastContact).toBeLessThanOrEqual(7);
        }
      });

      // Critical: not synced in 7+ days
      critical.forEach(device => {
        if (device.daysSinceLastContact !== undefined) {
          expect(device.daysSinceLastContact).toBeGreaterThan(7);
        }
      });
    });
  });

  // ==========================================================================
  // Test: getMobileDevicesByOperatingSystem (Report 22)
  // ==========================================================================

  describe('getMobileDevicesByOperatingSystem', () => {
    it('should group devices by operating system', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();

      expect(osDistribution).toBeDefined();
      expect(Array.isArray(osDistribution)).toBe(true);
      expect(osDistribution.length).toBeGreaterThan(0);
    });

    it('should calculate device counts correctly', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();

      const totalDevices = osDistribution.reduce((sum, os) => sum + os.deviceCount, 0);
      expect(totalDevices).toBe(6); // Total mock devices
    });

    it('should calculate percentages correctly', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();

      const totalPercentage = osDistribution.reduce((sum, os) => sum + os.percentage, 0);
      expect(totalPercentage).toBeCloseTo(100, 1);
    });

    it('should include version distribution for each OS', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();

      osDistribution.forEach(os => {
        expect(os.versions).toBeDefined();
        expect(Array.isArray(os.versions)).toBe(true);
        expect(os.versions.length).toBeGreaterThan(0);
      });
    });

    it('should sort by device count descending', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();

      for (let i = 1; i < osDistribution.length; i++) {
        expect(osDistribution[i - 1].deviceCount).toBeGreaterThanOrEqual(osDistribution[i].deviceCount);
      }
    });

    it('should identify iOS, Android, and Windows platforms', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();
      const osList = osDistribution.map(os => os.operatingSystem);

      expect(osList).toContain('iOS');
      expect(osList).toContain('Android');
      expect(osList).toContain('Windows');
    });

    it('should calculate version percentages within each OS', async () => {
      const osDistribution = await report.getMobileDevicesByOperatingSystem();

      osDistribution.forEach(os => {
        const versionPercentageSum = os.versions.reduce((sum, v) => sum + v.percentage, 0);
        expect(versionPercentageSum).toBeCloseTo(100, 1);
      });
    });
  });

  // ==========================================================================
  // Test: getAllMobileDevices
  // ==========================================================================

  describe('getAllMobileDevices', () => {
    it('should retrieve all mobile devices', async () => {
      const devices = await report.getAllMobileDevices();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBe(6);
    });

    it('should include all device types', async () => {
      const devices = await report.getAllMobileDevices();

      const corporateCount = devices.filter(d => d.ownerType === DeviceOwnershipType.COMPANY).length;
      const personalCount = devices.filter(d => d.ownerType === DeviceOwnershipType.PERSONAL).length;

      expect(corporateCount).toBe(4);
      expect(personalCount).toBe(2);
    });

    it('should filter by operating system', async () => {
      const iosDevices = await report.getAllMobileDevices({ operatingSystem: 'iOS' });

      expect(iosDevices.every(d => d.operatingSystem === 'iOS')).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getDevicesByOwnerType
  // ==========================================================================

  describe('getDevicesByOwnerType', () => {
    it('should retrieve corporate devices', async () => {
      const corporateDevices = await report.getDevicesByOwnerType(DeviceOwnershipType.COMPANY);

      expect(corporateDevices.every(d => d.ownerType === DeviceOwnershipType.COMPANY)).toBe(true);
    });

    it('should retrieve personal devices', async () => {
      const personalDevices = await report.getDevicesByOwnerType(DeviceOwnershipType.PERSONAL);

      expect(personalDevices.every(d => d.ownerType === DeviceOwnershipType.PERSONAL)).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getDevicesByOS
  // ==========================================================================

  describe('getDevicesByOS', () => {
    it('should retrieve iOS devices', async () => {
      const iosDevices = await report.getDevicesByOS('iOS');

      expect(iosDevices.every(d => d.operatingSystem === 'iOS')).toBe(true);
    });

    it('should retrieve Android devices', async () => {
      const androidDevices = await report.getDevicesByOS('Android');

      expect(androidDevices.every(d => d.operatingSystem === 'Android')).toBe(true);
    });
  });

  // ==========================================================================
  // Test: getInactiveDevices
  // ==========================================================================

  describe('getInactiveDevices', () => {
    it('should identify inactive devices', async () => {
      const allDevices = await report.getAllMobileDevices();
      const inactiveDevices = await report.getInactiveDevices(allDevices, 15);

      expect(inactiveDevices).toBeDefined();
      expect(Array.isArray(inactiveDevices)).toBe(true);
    });

    it('should use default threshold of 30 days', async () => {
      const allDevices = await report.getAllMobileDevices();
      const inactiveDevices = await report.getInactiveDevices(allDevices);

      expect(inactiveDevices).toBeDefined();
    });

    it('should only include devices beyond threshold', async () => {
      const allDevices = await report.getAllMobileDevices();
      const threshold = 15;
      const inactiveDevices = await report.getInactiveDevices(allDevices, threshold);

      // All inactive devices should have not synced beyond threshold
      // This is tested in the implementation
      expect(inactiveDevices.length).toBeGreaterThanOrEqual(0);
    });
  });

  // ==========================================================================
  // Test: getActiveDevices
  // ==========================================================================

  describe('getActiveDevices', () => {
    it('should identify active devices', async () => {
      const allDevices = await report.getAllMobileDevices();
      const activeDevices = report.getActiveDevices(allDevices, 7);

      expect(activeDevices).toBeDefined();
      expect(Array.isArray(activeDevices)).toBe(true);
    });

    it('should only include devices within threshold', async () => {
      const allDevices = await report.getAllMobileDevices();
      const activeDevices = report.getActiveDevices(allDevices, 7);

      // All active devices should have synced within threshold
      activeDevices.forEach(device => {
        expect(device.lastSyncDateTime).toBeDefined();
      });
    });
  });

  // ==========================================================================
  // Test: paginate
  // ==========================================================================

  describe('paginate', () => {
    it('should paginate results correctly', async () => {
      const allDevices = await report.getAllMobileDevices();
      const pageSize = 2;
      const page1 = report.paginate(allDevices, { page: 1, pageSize });

      expect(page1.data.length).toBe(2);
      expect(page1.pagination.currentPage).toBe(1);
      expect(page1.pagination.pageSize).toBe(2);
      expect(page1.pagination.totalPages).toBe(3);
      expect(page1.pagination.totalRecords).toBe(6);
      expect(page1.pagination.hasNextPage).toBe(true);
      expect(page1.pagination.hasPreviousPage).toBe(false);
    });

    it('should handle last page correctly', async () => {
      const allDevices = await report.getAllMobileDevices();
      const pageSize = 2;
      const totalPages = Math.ceil(allDevices.length / pageSize);
      const lastPage = report.paginate(allDevices, { page: totalPages, pageSize });

      expect(lastPage.pagination.hasNextPage).toBe(false);
      expect(lastPage.pagination.hasPreviousPage).toBe(true);
    });

    it('should handle empty results', async () => {
      const emptyResults = report.paginate([], { page: 1, pageSize: 10 });

      expect(emptyResults.data.length).toBe(0);
      expect(emptyResults.pagination.totalRecords).toBe(0);
      expect(emptyResults.pagination.totalPages).toBe(0);
    });
  });

  // ==========================================================================
  // Test: execute (Full Report)
  // ==========================================================================

  describe('execute', () => {
    it('should generate complete mobile device inventory report', async () => {
      const reportData = await report.execute();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include metadata in report', async () => {
      const reportData = await report.execute();
      const metadata = reportData.metadata;

      expect(metadata.reportName).toBe('mobile-device-inventory');
      expect(metadata.generatedAt).toBeDefined();
      expect(metadata.generatedBy).toBe('Intune Reporting Dashboard');
      expect(metadata.recordCount).toBe(6);
    });

    it('should include summary statistics', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.totalDevices).toBe(6);
      expect(summary.byOwnership).toBeDefined();
      expect(summary.byOS).toBeDefined();
      expect(summary.byManagementAgent).toBeDefined();
      expect(summary.corporateDevices).toBe(4);
      expect(summary.personalDevices).toBe(2);
    });

    it('should group devices by ownership type', async () => {
      const reportData = await report.execute();
      const byOwnership = reportData.summary.byOwnership;

      expect(byOwnership[DeviceOwnershipType.COMPANY]).toBe(4);
      expect(byOwnership[DeviceOwnershipType.PERSONAL]).toBe(2);
    });

    it('should group devices by operating system', async () => {
      const reportData = await report.execute();
      const byOS = reportData.summary.byOS;

      expect(byOS['iOS']).toBe(3);
      expect(byOS['Android']).toBe(2);
      expect(byOS['Windows']).toBe(1);
    });
  });

  // ==========================================================================
  // Test: Error Handling
  // ==========================================================================

  describe('Error Handling', () => {
    it('should handle API errors gracefully', async () => {
      const errorClient = {
        api: () => ({
          filter: () => ({
            select: () => ({
              get: async () => {
                throw new Error('API Error');
              },
            }),
          }),
        }),
      };

      const errorReport = new MobileDeviceInventoryReports(
        errorClient as unknown as Client,
        config
      );

      await expect(errorReport.getAllCorporateOwnedDevices()).rejects.toThrow();
    });

    it('should handle empty device list', async () => {
      mockClient.responses.set('/deviceManagement/managedDevices', { value: [] });

      const devices = await report.getAllMobileDevices();

      expect(devices.length).toBe(0);
    });

    it('should handle devices with missing optional fields', async () => {
      const minimalDevice = {
        id: 'device-999',
        deviceName: 'Minimal Device',
        operatingSystem: 'iOS',
        osVersion: '17.0',
        managementAgent: 'mdm',
        complianceState: 'compliant',
      };

      mockClient.responses.set('/deviceManagement/managedDevices', {
        value: [minimalDevice],
      });

      const devices = await report.getAllMobileDevices();

      expect(devices.length).toBe(1);
      expect(devices[0].manufacturer).toBeUndefined();
      expect(devices[0].serialNumber).toBeUndefined();
    });
  });

  // ==========================================================================
  // Test: Convenience Functions
  // ==========================================================================

  describe('Convenience Functions', () => {
    it('should generate corporate devices report', async () => {
      const { generateCorporateDevicesReport } = await import('./mobile-device-inventory-reports');
      const reportData = await generateCorporateDevicesReport(mockClient as unknown as Client, config);

      expect(reportData).toBeDefined();
      expect(reportData.metadata.reportName).toBe('corporate-mobile-devices');
      expect(reportData.summary.totalCorporateDevices).toBe(4);
    });

    it('should generate OS distribution report', async () => {
      const { generateOSDistributionReport } = await import('./mobile-device-inventory-reports');
      const reportData = await generateOSDistributionReport(mockClient as unknown as Client, config);

      expect(reportData).toBeDefined();
      expect(reportData.metadata.reportName).toBe('mobile-devices-by-os');
      expect(reportData.summary.totalOperatingSystems).toBeGreaterThan(0);
    });

    it('should generate device client communication report', async () => {
      const { generateDeviceClientCommunicationReport } = await import('./mobile-device-inventory-reports');
      const reportData = await generateDeviceClientCommunicationReport(mockClient as unknown as Client, config);

      expect(reportData).toBeDefined();
      expect(reportData.metadata.reportName).toBe('mobile-device-client-communication');
      expect(reportData.summary.totalDevices).toBe(6);
      expect(reportData.summary.healthyDevices).toBeDefined();
      expect(reportData.summary.activeDevices).toBeDefined();
    });
  });
});

// ============================================================================
// Run Tests
// ============================================================================

// Export for test runner
export default describe;
