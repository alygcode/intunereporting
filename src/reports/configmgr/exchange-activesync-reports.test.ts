/**
 * Exchange ActiveSync Reports - Unit Tests
 *
 * Comprehensive test suite for the ExchangeActiveSyncReports module
 */

import {
  ExchangeActiveSyncReports,
  ExchangeAccessState,
  ExchangeAccessStateReason,
  ActiveSyncPolicyComplianceSummary,
  InactiveMobileDevice,
  MobileDeviceComplianceDetails,
  MobileDeviceSettingsSummary
} from './exchange-activesync-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';

// ============================================================================
// Mock Data
// ============================================================================

const mockDevice1 = {
  id: 'device-001',
  deviceName: 'iPhone 14 Pro',
  userPrincipalName: 'user1@contoso.com',
  emailAddress: 'user1@contoso.com',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  model: 'iPhone 14 Pro',
  manufacturer: 'Apple Inc.',
  lastSyncDateTime: new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString(), // 2 days ago
  exchangeAccessState: 'allowed',
  exchangeAccessStateReason: 'compliant',
  exchangeLastSuccessfulSyncDateTime: new Date(Date.now() - 2 * 24 * 60 * 60 * 1000).toISOString(),
  enrolledDateTime: new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString(),
  complianceState: 'compliant',
  managementAgent: 'easMdm',
  isSupervised: true,
  serialNumber: 'SN001',
  imei: 'IMEI001',
  easActivated: true,
  easActivationDateTime: new Date(Date.now() - 90 * 24 * 60 * 60 * 1000).toISOString(),
  activeSyncId: 'AS001',
  easDeviceId: 'EAS001',
  jailBroken: 'false'
};

const mockDevice2 = {
  id: 'device-002',
  deviceName: 'Samsung Galaxy S23',
  userPrincipalName: 'user2@contoso.com',
  emailAddress: 'user2@contoso.com',
  operatingSystem: 'Android',
  osVersion: '14.0',
  model: 'Galaxy S23',
  manufacturer: 'Samsung',
  lastSyncDateTime: new Date(Date.now() - 45 * 24 * 60 * 60 * 1000).toISOString(), // 45 days ago (inactive)
  exchangeAccessState: 'blocked',
  exchangeAccessStateReason: 'notCompliant',
  exchangeLastSuccessfulSyncDateTime: new Date(Date.now() - 45 * 24 * 60 * 60 * 1000).toISOString(),
  enrolledDateTime: new Date(Date.now() - 120 * 24 * 60 * 60 * 1000).toISOString(),
  complianceState: 'noncompliant',
  managementAgent: 'easMdm',
  isSupervised: false,
  serialNumber: 'SN002',
  imei: 'IMEI002',
  easActivated: true,
  easActivationDateTime: new Date(Date.now() - 120 * 24 * 60 * 60 * 1000).toISOString(),
  activeSyncId: 'AS002',
  easDeviceId: 'EAS002',
  jailBroken: 'false'
};

const mockDevice3 = {
  id: 'device-003',
  deviceName: 'iPad Air',
  userPrincipalName: 'user3@contoso.com',
  emailAddress: 'user3@contoso.com',
  operatingSystem: 'iOS',
  osVersion: '17.2',
  model: 'iPad Air',
  manufacturer: 'Apple Inc.',
  lastSyncDateTime: new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString(), // 5 days ago
  exchangeAccessState: 'quarantined',
  exchangeAccessStateReason: 'exchangeDeviceRule',
  exchangeLastSuccessfulSyncDateTime: new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString(),
  enrolledDateTime: new Date(Date.now() - 60 * 24 * 60 * 60 * 1000).toISOString(),
  complianceState: 'compliant',
  managementAgent: 'mdm',
  isSupervised: true,
  serialNumber: 'SN003',
  imei: 'IMEI003',
  easActivated: true,
  easActivationDateTime: new Date(Date.now() - 60 * 24 * 60 * 60 * 1000).toISOString(),
  activeSyncId: 'AS003',
  easDeviceId: 'EAS003',
  jailBroken: 'false'
};

const mockDevice4 = {
  id: 'device-004',
  deviceName: 'Pixel 7',
  userPrincipalName: 'user4@contoso.com',
  emailAddress: 'user4@contoso.com',
  operatingSystem: 'Android',
  osVersion: '14.0',
  model: 'Pixel 7',
  manufacturer: 'Google',
  lastSyncDateTime: new Date(Date.now() - 35 * 24 * 60 * 60 * 1000).toISOString(), // 35 days ago (inactive)
  exchangeAccessState: 'allowed',
  exchangeAccessStateReason: 'compliant',
  exchangeLastSuccessfulSyncDateTime: new Date(Date.now() - 35 * 24 * 60 * 60 * 1000).toISOString(),
  enrolledDateTime: new Date(Date.now() - 100 * 24 * 60 * 60 * 1000).toISOString(),
  complianceState: 'compliant',
  managementAgent: 'eas',
  isSupervised: false,
  serialNumber: 'SN004',
  imei: 'IMEI004',
  easActivated: true,
  easActivationDateTime: new Date(Date.now() - 100 * 24 * 60 * 60 * 1000).toISOString(),
  activeSyncId: 'AS004',
  easDeviceId: 'EAS004',
  jailBroken: 'false'
};

const mockCompliancePolicy1 = {
  id: 'policy-001',
  displayName: 'iOS Compliance Policy',
  '@odata.type': '#microsoft.graph.iosCompliancePolicy'
};

const mockCompliancePolicy2 = {
  id: 'policy-002',
  displayName: 'Android Compliance Policy',
  '@odata.type': '#microsoft.graph.androidCompliancePolicy'
};

const mockPolicyStatus1 = {
  id: 'device-001',
  status: 'compliant',
  platform: 'iOS'
};

const mockPolicyStatus2 = {
  id: 'device-002',
  status: 'noncompliant',
  platform: 'Android'
};

const mockCompliancePolicyState = {
  id: 'policy-001',
  displayName: 'iOS Compliance Policy',
  state: 'compliant',
  lastReportedDateTime: new Date().toISOString(),
  settingStates: [
    {
      setting: 'passwordRequired',
      settingName: 'Password Required',
      state: 'compliant',
      currentValue: 'true'
    },
    {
      setting: 'osMinimumVersion',
      settingName: 'Minimum OS Version',
      state: 'compliant',
      currentValue: '17.0'
    }
  ]
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
    // Managed devices endpoint
    this.responses.set('/deviceManagement/managedDevices', {
      value: [mockDevice1, mockDevice2, mockDevice3, mockDevice4]
    });

    // Compliance policies endpoint
    this.responses.set('/deviceManagement/deviceCompliancePolicies', {
      value: [mockCompliancePolicy1, mockCompliancePolicy2]
    });

    // Policy device statuses
    this.responses.set('/deviceManagement/deviceCompliancePolicies/policy-001/deviceStatuses', {
      value: [mockPolicyStatus1]
    });

    this.responses.set('/deviceManagement/deviceCompliancePolicies/policy-002/deviceStatuses', {
      value: [mockPolicyStatus2]
    });

    // Device compliance policy states
    this.responses.set('/deviceManagement/managedDevices/device-001/deviceCompliancePolicyStates', {
      value: [mockCompliancePolicyState]
    });

    this.responses.set('/deviceManagement/managedDevices/device-002/deviceCompliancePolicyStates', {
      value: []
    });

    this.responses.set('/deviceManagement/managedDevices/device-003/deviceCompliancePolicyStates', {
      value: [mockCompliancePolicyState]
    });

    this.responses.set('/deviceManagement/managedDevices/device-004/deviceCompliancePolicyStates', {
      value: [mockCompliancePolicyState]
    });

    // Device details
    this.responses.set('/deviceManagement/managedDevices/device-001', {
      deviceActionResults: []
    });
  }

  api(endpoint: string) {
    const mockApi = {
      endpoint,
      selectFields: [] as string[],
      topValue: 999,
      filterValue: '',
      expandValue: '',

      select(fields: string[] | string) {
        mockApi.selectFields = Array.isArray(fields) ? fields : [fields];
        return mockApi;
      },

      top(value: number) {
        mockApi.topValue = value;
        return mockApi;
      },

      filter(value: string) {
        mockApi.filterValue = value;
        return mockApi;
      },

      expand(value: string) {
        mockApi.expandValue = value;
        return mockApi;
      },

      async get() {
        const response = this.responses.get(mockApi.endpoint);
        if (response) {
          return response;
        }

        // Default response
        return { value: [] };
      }
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
      authMethod: 'clientSecret'
    },
    reports: {
      enabled: ['exchange-activesync-reports'],
      disabled: [],
      settings: {}
    },
    output: {
      defaultFormat: 'json',
      directory: './test-reports',
      includeTimestamp: false,
      compression: false
    },
    scheduler: {
      enabled: false,
      timezone: 'UTC',
      schedules: []
    },
    logging: {
      level: 'error',
      file: './test-logs/app.log',
      console: false,
      maxSize: '10m',
      maxFiles: 5
    }
  };
}

// ============================================================================
// Test Suite
// ============================================================================

describe('ExchangeActiveSyncReports', () => {
  let report: ExchangeActiveSyncReports;
  let mockClient: any;
  let config: AppConfig;

  beforeEach(() => {
    mockClient = new MockGraphClient();
    config = getMockConfig();
    report = new ExchangeActiveSyncReports(mockClient as unknown as Client, config);
  });

  // ==========================================================================
  // Test: getActiveSyncPolicyCompliance (Report 8)
  // ==========================================================================

  describe('getActiveSyncPolicyCompliance', () => {
    it('should retrieve ActiveSync policy compliance status', async () => {
      const summaries = await report.getActiveSyncPolicyCompliance();

      expect(summaries).toBeDefined();
      expect(Array.isArray(summaries)).toBe(true);
      expect(summaries.length).toBeGreaterThan(0);
    });

    it('should calculate compliance statistics correctly', async () => {
      const summaries = await report.getActiveSyncPolicyCompliance();

      summaries.forEach(summary => {
        expect(summary.totalDevices).toBeGreaterThanOrEqual(0);
        expect(summary.allowedDevices).toBeGreaterThanOrEqual(0);
        expect(summary.blockedDevices).toBeGreaterThanOrEqual(0);
        expect(summary.quarantinedDevices).toBeGreaterThanOrEqual(0);
        expect(summary.compliancePercentage).toBeGreaterThanOrEqual(0);
        expect(summary.compliancePercentage).toBeLessThanOrEqual(100);
      });
    });

    it('should include policy information', async () => {
      const summaries = await report.getActiveSyncPolicyCompliance();

      summaries.forEach(summary => {
        expect(summary.policyId).toBeDefined();
        expect(summary.policyName).toBeDefined();
        expect(summary.lastUpdated).toBeDefined();
      });
    });

    it('should aggregate devices by platform', async () => {
      const summaries = await report.getActiveSyncPolicyCompliance();

      summaries.forEach(summary => {
        expect(summary.devicesByPlatform).toBeDefined();
        expect(typeof summary.devicesByPlatform).toBe('object');
      });
    });

    it('should aggregate devices by access state', async () => {
      const summaries = await report.getActiveSyncPolicyCompliance();

      summaries.forEach(summary => {
        expect(summary.devicesByAccessState).toBeDefined();
        const states = Object.keys(summary.devicesByAccessState);
        expect(states.length).toBeGreaterThan(0);
      });
    });
  });

  // ==========================================================================
  // Test: getInactiveMobileDevices (Report 15)
  // ==========================================================================

  describe('getInactiveMobileDevices', () => {
    it('should find inactive devices based on threshold', async () => {
      const inactiveDevices = await report.getInactiveMobileDevices(30);

      expect(inactiveDevices).toBeDefined();
      expect(Array.isArray(inactiveDevices)).toBe(true);
    });

    it('should only include devices exceeding inactivity threshold', async () => {
      const threshold = 30;
      const inactiveDevices = await report.getInactiveMobileDevices(threshold);

      inactiveDevices.forEach(device => {
        expect(device.daysSinceLastSync).toBeGreaterThanOrEqual(threshold);
      });
    });

    it('should calculate days since last sync correctly', async () => {
      const inactiveDevices = await report.getInactiveMobileDevices(30);

      inactiveDevices.forEach(device => {
        expect(device.daysSinceLastSync).toBeGreaterThan(0);
        expect(Number.isInteger(device.daysSinceLastSync)).toBe(true);
      });
    });

    it('should include device information', async () => {
      const inactiveDevices = await report.getInactiveMobileDevices(30);

      if (inactiveDevices.length > 0) {
        const device = inactiveDevices[0];
        expect(device.deviceId).toBeDefined();
        expect(device.deviceName).toBeDefined();
        expect(device.userPrincipalName).toBeDefined();
        expect(device.operatingSystem).toBeDefined();
        expect(device.exchangeAccessState).toBeDefined();
      }
    });

    it('should sort devices by days since last sync descending', async () => {
      const inactiveDevices = await report.getInactiveMobileDevices(30);

      for (let i = 1; i < inactiveDevices.length; i++) {
        expect(inactiveDevices[i - 1].daysSinceLastSync).toBeGreaterThanOrEqual(
          inactiveDevices[i].daysSinceLastSync
        );
      }
    });

    it('should support custom inactivity thresholds', async () => {
      const devices7Days = await report.getInactiveMobileDevices(7);
      const devices30Days = await report.getInactiveMobileDevices(30);
      const devices60Days = await report.getInactiveMobileDevices(60);

      expect(devices7Days.length).toBeGreaterThanOrEqual(devices30Days.length);
      expect(devices30Days.length).toBeGreaterThanOrEqual(devices60Days.length);
    });

    it('should filter by platform when specified', async () => {
      const inactiveDevices = await report.getInactiveMobileDevices(30, {
        platform: 'Android'
      });

      inactiveDevices.forEach(device => {
        expect(device.operatingSystem).toBe('Android');
      });
    });
  });

  // ==========================================================================
  // Test: getAllMobileDeviceComplianceDetails (Report 21)
  // ==========================================================================

  describe('getAllMobileDeviceComplianceDetails', () => {
    it('should retrieve detailed compliance information', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();

      expect(deviceDetails).toBeDefined();
      expect(Array.isArray(deviceDetails)).toBe(true);
    });

    it('should include comprehensive device information', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();

      if (deviceDetails.length > 0) {
        const device = deviceDetails[0];
        expect(device.deviceId).toBeDefined();
        expect(device.deviceName).toBeDefined();
        expect(device.userPrincipalName).toBeDefined();
        expect(device.operatingSystem).toBeDefined();
        expect(device.exchangeAccessState).toBeDefined();
        expect(device.exchangeAccessStateReason).toBeDefined();
        expect(device.complianceState).toBeDefined();
        expect(device.easActivated).toBeDefined();
      }
    });

    it('should include compliance policy information', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();

      deviceDetails.forEach(device => {
        expect(Array.isArray(device.compliancePolicies)).toBe(true);
      });
    });

    it('should include Exchange ActiveSync details', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();

      if (deviceDetails.length > 0) {
        const device = deviceDetails[0];
        expect(device.easActivated).toBeDefined();
        if (device.easActivated) {
          expect(device.easActivationDateTime).toBeDefined();
        }
      }
    });

    it('should filter by exchange access state when specified', async () => {
      const allowedDevices = await report.getAllMobileDeviceComplianceDetails({
        exchangeAccessState: ExchangeAccessState.ALLOWED
      });

      allowedDevices.forEach(device => {
        expect(device.exchangeAccessState).toBe(ExchangeAccessState.ALLOWED);
      });
    });

    it('should filter by compliance state when specified', async () => {
      const compliantDevices = await report.getAllMobileDeviceComplianceDetails({
        complianceState: 'compliant'
      });

      compliantDevices.forEach(device => {
        expect(device.complianceState).toBe('compliant');
      });
    });
  });

  // ==========================================================================
  // Test: getMobileDeviceSettingsSummary (Report 34)
  // ==========================================================================

  describe('getMobileDeviceSettingsSummary', () => {
    it('should generate settings summary', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      expect(summaries).toBeDefined();
      expect(Array.isArray(summaries)).toBe(true);
    });

    it('should categorize settings correctly', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      summaries.forEach(summary => {
        expect(summary.settingCategory).toBeDefined();
        expect(['exchange', 'compliance', 'configuration']).toContain(
          summary.settingCategory
        );
      });
    });

    it('should calculate device counts per setting', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      summaries.forEach(summary => {
        expect(summary.deviceCount).toBeGreaterThanOrEqual(0);
        expect(summary.enabledCount).toBeGreaterThanOrEqual(0);
        expect(summary.disabledCount).toBeGreaterThanOrEqual(0);
        expect(summary.notConfiguredCount).toBeGreaterThanOrEqual(0);

        const total = summary.enabledCount + summary.disabledCount + summary.notConfiguredCount;
        expect(total).toBeLessThanOrEqual(summary.deviceCount);
      });
    });

    it('should calculate percentages correctly', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      summaries.forEach(summary => {
        expect(summary.percentage).toBeGreaterThanOrEqual(0);
        expect(summary.percentage).toBeLessThanOrEqual(100);
      });
    });

    it('should include affected policies', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      summaries.forEach(summary => {
        expect(Array.isArray(summary.affectedPolicies)).toBe(true);
      });
    });

    it('should sort settings by device count descending', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      for (let i = 1; i < summaries.length; i++) {
        expect(summaries[i - 1].deviceCount).toBeGreaterThanOrEqual(
          summaries[i].deviceCount
        );
      }
    });

    it('should include platform distribution', async () => {
      const summaries = await report.getMobileDeviceSettingsSummary();

      summaries.forEach(summary => {
        expect(summary.platforms).toBeDefined();
        expect(typeof summary.platforms).toBe('object');
      });
    });
  });

  // ==========================================================================
  // Test: execute (Full Report)
  // ==========================================================================

  describe('execute', () => {
    it('should generate complete Exchange ActiveSync report', async () => {
      const reportData = await report.execute();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include metadata in report', async () => {
      const reportData = await report.execute();
      const metadata = reportData.metadata;

      expect(metadata.reportName).toBe('exchange-activesync-reports');
      expect(metadata.generatedAt).toBeDefined();
      expect(metadata.generatedBy).toBe('Intune Reporting Dashboard');
    });

    it('should include all report types in data', async () => {
      const reportData = await report.execute();
      const data = reportData.data as any[];

      const types = new Set(data.map(item => item.type));
      expect(types.has('policy-compliance')).toBe(true);
      expect(types.has('inactive-device')).toBe(true);
      expect(types.has('device-details')).toBe(true);
      expect(types.has('settings-summary')).toBe(true);
    });

    it('should include comprehensive summary', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary as any;

      expect(summary.totalMobileDevices).toBeDefined();
      expect(summary.activeDevices).toBeDefined();
      expect(summary.inactiveDevices).toBeDefined();
      expect(summary.allowedDevices).toBeDefined();
      expect(summary.blockedDevices).toBeDefined();
      expect(summary.easActivatedDevices).toBeDefined();
      expect(summary.compliancePercentage).toBeDefined();
    });
  });

  // ==========================================================================
  // Test: Export Methods
  // ==========================================================================

  describe('Export Methods', () => {
    it('should have exportToJson method', () => {
      expect(typeof report.exportToJson).toBe('function');
    });

    it('should have exportToCsv method', () => {
      expect(typeof report.exportToCsv).toBe('function');
    });

    it('should have exportToHtml method', () => {
      expect(typeof report.exportToHtml).toBe('function');
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
            filter: () => ({
              top: () => ({
                get: async () => {
                  throw new Error('API Error');
                }
              })
            })
          })
        })
      };

      const errorReport = new ExchangeActiveSyncReports(
        errorClient as unknown as Client,
        config
      );

      await expect(errorReport.getActiveSyncPolicyCompliance()).resolves.toEqual([]);
    });

    it('should handle empty device lists', async () => {
      mockClient.responses.set('/deviceManagement/managedDevices', { value: [] });

      const inactiveDevices = await report.getInactiveMobileDevices(30);
      expect(inactiveDevices.length).toBe(0);
    });

    it('should handle missing optional device fields', async () => {
      const minimalDevice = {
        id: 'device-999',
        deviceName: 'Minimal Device',
        operatingSystem: 'iOS',
        lastSyncDateTime: new Date(Date.now() - 45 * 24 * 60 * 60 * 1000).toISOString(),
        complianceState: 'compliant',
        managementAgent: 'mdm'
      };

      mockClient.responses.set('/deviceManagement/managedDevices', {
        value: [minimalDevice]
      });

      const inactiveDevices = await report.getInactiveMobileDevices(30);
      expect(inactiveDevices.length).toBeGreaterThan(0);
    });
  });

  // ==========================================================================
  // Test: Exchange Access State Mapping
  // ==========================================================================

  describe('Exchange Access State Mapping', () => {
    it('should map allowed state correctly', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();
      const allowedDevices = deviceDetails.filter(
        d => d.exchangeAccessState === ExchangeAccessState.ALLOWED
      );
      expect(allowedDevices.length).toBeGreaterThan(0);
    });

    it('should map blocked state correctly', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();
      const blockedDevices = deviceDetails.filter(
        d => d.exchangeAccessState === ExchangeAccessState.BLOCKED
      );
      expect(blockedDevices.length).toBeGreaterThan(0);
    });

    it('should map quarantined state correctly', async () => {
      const deviceDetails = await report.getAllMobileDeviceComplianceDetails();
      const quarantinedDevices = deviceDetails.filter(
        d => d.exchangeAccessState === ExchangeAccessState.QUARANTINED
      );
      expect(quarantinedDevices.length).toBeGreaterThan(0);
    });
  });

  // ==========================================================================
  // Test: Compliance Trends
  // ==========================================================================

  describe('Compliance Trends', () => {
    it('should track compliance over time', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary as any;

      expect(summary.compliancePercentage).toBeDefined();
      expect(summary.compliancePercentage).toBeGreaterThanOrEqual(0);
      expect(summary.compliancePercentage).toBeLessThanOrEqual(100);
    });

    it('should calculate average days since last sync', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary as any;

      expect(summary.averageDaysSinceLastSync).toBeDefined();
      expect(summary.averageDaysSinceLastSync).toBeGreaterThanOrEqual(0);
    });
  });
});

// ============================================================================
// Run Tests
// ============================================================================

export default describe;
