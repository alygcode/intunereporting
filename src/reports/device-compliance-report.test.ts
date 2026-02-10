/**
 * Device Compliance Report - Unit Tests
 *
 * Comprehensive test suite for the DeviceComplianceReport module
 */

import { DeviceComplianceReport, ComplianceState, CompliancePlatform } from './device-compliance-report';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../types';

// ============================================================================
// Mock Data
// ============================================================================

const mockPolicy = {
  id: 'policy-123',
  displayName: 'Windows 10 Compliance Policy',
  description: 'Test policy for Windows 10 devices',
  '@odata.type': '#microsoft.graph.windows10CompliancePolicy',
  createdDateTime: '2024-01-01T00:00:00Z',
  lastModifiedDateTime: '2024-01-15T00:00:00Z',
  version: 1,
  assignments: [{ id: 'assignment-1' }],
  scheduledActionsForRule: [{ id: 'action-1' }]
};

const mockDevice = {
  id: 'device-123',
  deviceName: 'TEST-DEVICE-01',
  userPrincipalName: 'user@contoso.com',
  operatingSystem: 'Windows',
  osVersion: '10.0.19045',
  complianceState: 'compliant',
  lastSyncDateTime: '2024-02-01T12:00:00Z',
  managementAgent: 'mdm',
  manufacturer: 'Microsoft Corporation',
  model: 'Surface Pro 9',
  serialNumber: 'SN123456',
  isSupervised: false
};

const mockNonCompliantDevice = {
  ...mockDevice,
  id: 'device-456',
  deviceName: 'TEST-DEVICE-02',
  complianceState: 'noncompliant'
};

const mockPolicyState = {
  id: 'policy-123',
  displayName: 'Windows 10 Compliance Policy',
  platform: 'Windows',
  state: 'compliant',
  version: 1,
  settingCount: 5,
  settingStates: [
    {
      setting: 'passwordRequired',
      state: 'compliant',
      userId: 'user-123',
      userName: 'Test User'
    }
  ],
  lastReportedDateTime: '2024-02-01T12:00:00Z'
};

const mockNonCompliantPolicyState = {
  ...mockPolicyState,
  state: 'noncompliant',
  settingStates: [
    {
      setting: 'passwordRequired',
      state: 'noncompliant',
      errorCode: '0x80004005',
      errorDescription: 'Password does not meet complexity requirements',
      currentValue: 'false'
    },
    {
      setting: 'requireEncryption',
      state: 'noncompliant',
      errorCode: '0x80004006',
      errorDescription: 'Device encryption is not enabled'
    }
  ]
};

const mockDeviceStateSummary = {
  compliantDeviceCount: 85,
  nonCompliantDeviceCount: 10,
  errorDeviceCount: 3,
  conflictDeviceCount: 1,
  notApplicableDeviceCount: 5,
  inGracePeriodCount: 2,
  unknownDeviceCount: 4,
  remediatedDeviceCount: 0,
  notAssignedDeviceCount: 0
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
    // Policies endpoint
    this.responses.set('/deviceManagement/deviceCompliancePolicies', {
      value: [mockPolicy]
    });

    // Devices endpoint
    this.responses.set('/deviceManagement/managedDevices', {
      value: [mockDevice, mockNonCompliantDevice]
    });

    // Policy state summary
    this.responses.set(
      `/deviceManagement/deviceCompliancePolicies/${mockPolicy.id}/deviceStateSummary`,
      mockDeviceStateSummary
    );

    // Device policy states
    this.responses.set(
      `/deviceManagement/managedDevices/${mockDevice.id}/deviceCompliancePolicyStates`,
      { value: [mockPolicyState] }
    );

    this.responses.set(
      `/deviceManagement/managedDevices/${mockNonCompliantDevice.id}/deviceCompliancePolicyStates`,
      { value: [mockNonCompliantPolicyState] }
    );
  }

  api(endpoint: string) {
    const mockApi = {
      endpoint,
      selectFields: [] as string[],
      expandFields: [] as string[],
      filterString: '',
      topValue: 999,

      select(fields: string[] | string) {
        mockApi.selectFields = Array.isArray(fields) ? fields : [fields];
        return mockApi;
      },

      expand(fields: string) {
        mockApi.expandFields = [fields];
        return mockApi;
      },

      filter(filter: string) {
        mockApi.filterString = filter;
        return mockApi;
      },

      top(value: number) {
        mockApi.topValue = value;
        return mockApi;
      },

      async get() {
        // Handle filtering
        if (mockApi.filterString.includes('noncompliant')) {
          return { value: [mockNonCompliantDevice] };
        }

        if (mockApi.filterString.includes('mdm')) {
          return { value: [mockDevice, mockNonCompliantDevice] };
        }

        // Check for exact endpoint match
        const response = this.responses.get(mockApi.endpoint);
        if (response) {
          return response;
        }

        // Check for pattern matches
        for (const [pattern, data] of this.responses.entries()) {
          if (mockApi.endpoint.includes(pattern.split('/').pop() || '')) {
            return data;
          }
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
      enabled: ['device-compliance-report'],
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

describe('DeviceComplianceReport', () => {
  let report: DeviceComplianceReport;
  let mockClient: any;
  let config: AppConfig;

  beforeEach(() => {
    mockClient = new MockGraphClient();
    config = getMockConfig();
    report = new DeviceComplianceReport(mockClient as unknown as Client, config);
  });

  // ==========================================================================
  // Test: getAllCompliancePolicies
  // ==========================================================================

  describe('getAllCompliancePolicies', () => {
    it('should retrieve all compliance policies', async () => {
      const policies = await report.getAllCompliancePolicies();

      expect(policies).toBeDefined();
      expect(Array.isArray(policies)).toBe(true);
      expect(policies.length).toBeGreaterThan(0);
    });

    it('should parse policy data correctly', async () => {
      const policies = await report.getAllCompliancePolicies();
      const policy = policies[0];

      expect(policy.id).toBe(mockPolicy.id);
      expect(policy.displayName).toBe(mockPolicy.displayName);
      expect(policy.description).toBe(mockPolicy.description);
      expect(policy.platform).toBe(CompliancePlatform.WINDOWS_10);
      expect(policy.assignmentCount).toBe(1);
      expect(policy.isAssigned).toBe(true);
    });

    it('should filter policies by platform', async () => {
      const policies = await report.getAllCompliancePolicies({
        platform: CompliancePlatform.WINDOWS_10
      });

      expect(policies.length).toBeGreaterThan(0);
      expect(policies.every(p => p.platform === CompliancePlatform.WINDOWS_10)).toBe(true);
    });

    it('should handle empty policy list', async () => {
      mockClient.responses.set('/deviceManagement/deviceCompliancePolicies', { value: [] });
      const policies = await report.getAllCompliancePolicies();

      expect(policies).toBeDefined();
      expect(policies.length).toBe(0);
    });
  });

  // ==========================================================================
  // Test: getAllPolicyDeviceStateSummaries
  // ==========================================================================

  describe('getAllPolicyDeviceStateSummaries', () => {
    it('should retrieve policy device state summaries', async () => {
      const summaries = await report.getAllPolicyDeviceStateSummaries();

      expect(summaries).toBeDefined();
      expect(Array.isArray(summaries)).toBe(true);
    });

    it('should calculate compliance percentage correctly', async () => {
      const summaries = await report.getAllPolicyDeviceStateSummaries();
      const summary = summaries[0];

      const total = summary.compliantDeviceCount +
                   summary.nonCompliantDeviceCount +
                   summary.errorDeviceCount;

      const expectedPercentage = Math.round((summary.compliantDeviceCount / total) * 100 * 100) / 100;

      expect(summary.compliancePercentage).toBeCloseTo(expectedPercentage, 2);
    });

    it('should include all device state counts', async () => {
      const summaries = await report.getAllPolicyDeviceStateSummaries();
      const summary = summaries[0];

      expect(summary.compliantDeviceCount).toBeDefined();
      expect(summary.nonCompliantDeviceCount).toBeDefined();
      expect(summary.errorDeviceCount).toBeDefined();
      expect(summary.inGracePeriodCount).toBeDefined();
      expect(summary.totalDeviceCount).toBeGreaterThan(0);
    });
  });

  // ==========================================================================
  // Test: getAllDeviceComplianceDetails
  // ==========================================================================

  describe('getAllDeviceComplianceDetails', () => {
    it('should retrieve device compliance details', async () => {
      const devices = await report.getAllDeviceComplianceDetails();

      expect(devices).toBeDefined();
      expect(Array.isArray(devices)).toBe(true);
      expect(devices.length).toBeGreaterThan(0);
    });

    it('should parse device data correctly', async () => {
      const devices = await report.getAllDeviceComplianceDetails();
      const device = devices[0];

      expect(device.deviceId).toBeDefined();
      expect(device.deviceName).toBeDefined();
      expect(device.userPrincipalName).toBeDefined();
      expect(device.platform).toBeDefined();
      expect(device.overallComplianceState).toBeDefined();
      expect(device.policies).toBeDefined();
      expect(Array.isArray(device.policies)).toBe(true);
    });

    it('should filter devices by compliance state', async () => {
      const devices = await report.getAllDeviceComplianceDetails({
        complianceState: ComplianceState.COMPLIANT
      });

      expect(devices.every(d => d.overallComplianceState === ComplianceState.COMPLIANT)).toBe(true);
    });

    it('should include policy states for each device', async () => {
      const devices = await report.getAllDeviceComplianceDetails();
      const device = devices[0];

      expect(device.compliancePoliciesCount).toBeGreaterThanOrEqual(0);
      expect(device.compliantPoliciesCount).toBeGreaterThanOrEqual(0);
      expect(device.nonCompliantPoliciesCount).toBeGreaterThanOrEqual(0);
    });
  });

  // ==========================================================================
  // Test: getDevicePolicyStates
  // ==========================================================================

  describe('getDevicePolicyStates', () => {
    it('should retrieve policy states for a device', async () => {
      const states = await report.getDevicePolicyStates(mockDevice.id);

      expect(states).toBeDefined();
      expect(Array.isArray(states)).toBe(true);
    });

    it('should parse policy state correctly', async () => {
      const states = await report.getDevicePolicyStates(mockDevice.id);
      const state = states[0];

      expect(state.policyId).toBe(mockPolicyState.id);
      expect(state.policyName).toBe(mockPolicyState.displayName);
      expect(state.state).toBe(ComplianceState.COMPLIANT);
      expect(state.settingStates).toBeDefined();
      expect(Array.isArray(state.settingStates)).toBe(true);
    });

    it('should include setting states', async () => {
      const states = await report.getDevicePolicyStates(mockDevice.id);
      const state = states[0];

      expect(state.settingStates.length).toBeGreaterThan(0);
      const setting = state.settingStates[0];

      expect(setting.settingName).toBeDefined();
      expect(setting.state).toBeDefined();
    });
  });

  // ==========================================================================
  // Test: getNonCompliantDevicesWithReasons
  // ==========================================================================

  describe('getNonCompliantDevicesWithReasons', () => {
    it('should retrieve non-compliant devices', async () => {
      const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

      expect(nonCompliantDevices).toBeDefined();
      expect(Array.isArray(nonCompliantDevices)).toBe(true);
    });

    it('should include non-compliant policy information', async () => {
      const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

      if (nonCompliantDevices.length > 0) {
        const device = nonCompliantDevices[0];

        expect(device.nonCompliantPolicies).toBeDefined();
        expect(Array.isArray(device.nonCompliantPolicies)).toBe(true);
        expect(device.totalNonCompliantSettings).toBeGreaterThanOrEqual(0);
      }
    });

    it('should include remediation actions for non-compliant settings', async () => {
      const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

      if (nonCompliantDevices.length > 0) {
        const device = nonCompliantDevices[0];

        if (device.nonCompliantPolicies.length > 0) {
          const policy = device.nonCompliantPolicies[0];

          if (policy.nonCompliantSettings.length > 0) {
            const setting = policy.nonCompliantSettings[0];

            expect(setting.remediationActions).toBeDefined();
            expect(Array.isArray(setting.remediationActions)).toBe(true);
          }
        }
      }
    });

    it('should calculate total non-compliant settings correctly', async () => {
      const nonCompliantDevices = await report.getNonCompliantDevicesWithReasons();

      if (nonCompliantDevices.length > 0) {
        const device = nonCompliantDevices[0];

        const calculatedTotal = device.nonCompliantPolicies.reduce(
          (sum, policy) => sum + policy.settingCount,
          0
        );

        expect(device.totalNonCompliantSettings).toBe(calculatedTotal);
      }
    });
  });

  // ==========================================================================
  // Test: getComplianceTrends
  // ==========================================================================

  describe('getComplianceTrends', () => {
    it('should generate compliance trends', async () => {
      const trends = await report.getComplianceTrends(30);

      expect(trends).toBeDefined();
      expect(trends?.dataPoints).toBeDefined();
      expect(Array.isArray(trends?.dataPoints)).toBe(true);
    });

    it('should calculate trend direction correctly', async () => {
      const trends = await report.getComplianceTrends(30);

      expect(trends?.trend).toBeDefined();
      expect(['improving', 'declining', 'stable']).toContain(trends?.trend);
    });

    it('should identify peak and lowest compliance', async () => {
      const trends = await report.getComplianceTrends(30);

      expect(trends?.peakCompliance).toBeDefined();
      expect(trends?.lowestCompliance).toBeDefined();
      expect(trends?.peakCompliance.compliancePercentage).toBeGreaterThanOrEqual(
        trends?.lowestCompliance.compliancePercentage || 0
      );
    });

    it('should calculate average compliance percentage', async () => {
      const trends = await report.getComplianceTrends(30);

      expect(trends?.averageCompliancePercentage).toBeDefined();
      expect(trends?.averageCompliancePercentage).toBeGreaterThanOrEqual(0);
      expect(trends?.averageCompliancePercentage).toBeLessThanOrEqual(100);
    });
  });

  // ==========================================================================
  // Test: execute (Full Report)
  // ==========================================================================

  describe('execute', () => {
    it('should generate complete compliance report', async () => {
      const reportData = await report.execute();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include metadata in report', async () => {
      const reportData = await report.execute();
      const metadata = reportData.metadata;

      expect(metadata.reportName).toBe('device-compliance-report');
      expect(metadata.generatedAt).toBeDefined();
      expect(metadata.generatedBy).toBe('Intune Reporting Dashboard');
      expect(metadata.recordCount).toBeGreaterThanOrEqual(0);
    });

    it('should include summary statistics', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.totalDevices).toBeGreaterThanOrEqual(0);
      expect(summary.compliantDevices).toBeGreaterThanOrEqual(0);
      expect(summary.nonCompliantDevices).toBeGreaterThanOrEqual(0);
      expect(summary.compliancePercentage).toBeGreaterThanOrEqual(0);
      expect(summary.compliancePercentage).toBeLessThanOrEqual(100);
      expect(summary.totalPolicies).toBeGreaterThanOrEqual(0);
    });

    it('should include platform breakdown', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.byPlatform).toBeDefined();
      expect(typeof summary.byPlatform).toBe('object');
    });

    it('should identify critical issues', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.criticalIssues).toBeDefined();
      expect(Array.isArray(summary.criticalIssues)).toBe(true);
    });

    it('should include top non-compliant policies', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.topNonCompliantPolicies).toBeDefined();
      expect(Array.isArray(summary.topNonCompliantPolicies)).toBe(true);
    });
  });

  // ==========================================================================
  // Test: Error Handling
  // ==========================================================================

  describe('Error Handling', () => {
    it('should handle API errors gracefully', async () => {
      const errorClient = {
        api: () => ({
          select: () => ({ expand: () => ({ top: () => ({ get: async () => {
            throw new Error('API Error');
          }})})}})
        })
      };

      const errorReport = new DeviceComplianceReport(
        errorClient as unknown as Client,
        config
      );

      await expect(errorReport.getAllCompliancePolicies()).rejects.toThrow();
    });

    it('should handle empty responses', async () => {
      mockClient.responses.set('/deviceManagement/deviceCompliancePolicies', { value: [] });
      mockClient.responses.set('/deviceManagement/managedDevices', { value: [] });

      const policies = await report.getAllCompliancePolicies();
      const devices = await report.getAllDeviceComplianceDetails();

      expect(policies.length).toBe(0);
      expect(devices.length).toBe(0);
    });

    it('should handle missing optional fields', async () => {
      const minimalDevice = {
        id: 'device-999',
        deviceName: 'Minimal Device',
        userPrincipalName: 'user@contoso.com',
        operatingSystem: 'Windows',
        complianceState: 'compliant',
        managementAgent: 'mdm'
      };

      mockClient.responses.set('/deviceManagement/managedDevices', {
        value: [minimalDevice]
      });

      const devices = await report.getAllDeviceComplianceDetails();

      expect(devices.length).toBeGreaterThan(0);
      expect(devices[0].deviceName).toBe('Minimal Device');
    });
  });

  // ==========================================================================
  // Test: Data Validation
  // ==========================================================================

  describe('Data Validation', () => {
    it('should correctly map compliance states', async () => {
      const testStates = [
        { input: 'compliant', expected: ComplianceState.COMPLIANT },
        { input: 'noncompliant', expected: ComplianceState.NON_COMPLIANT },
        { input: 'inGracePeriod', expected: ComplianceState.IN_GRACE_PERIOD },
        { input: 'error', expected: ComplianceState.ERROR },
        { input: 'unknown', expected: ComplianceState.UNKNOWN }
      ];

      for (const testCase of testStates) {
        const testDevice = { ...mockDevice, complianceState: testCase.input };
        mockClient.responses.set('/deviceManagement/managedDevices', {
          value: [testDevice]
        });

        const devices = await report.getAllDeviceComplianceDetails();
        expect(devices[0].overallComplianceState).toBe(testCase.expected);
      }
    });

    it('should correctly determine platform from OData type', async () => {
      const testCases = [
        { odataType: '#microsoft.graph.windows10CompliancePolicy', expected: CompliancePlatform.WINDOWS_10 },
        { odataType: '#microsoft.graph.iosCompliancePolicy', expected: CompliancePlatform.IOS },
        { odataType: '#microsoft.graph.androidCompliancePolicy', expected: CompliancePlatform.ANDROID },
        { odataType: '#microsoft.graph.macOSCompliancePolicy', expected: CompliancePlatform.MAC_OS }
      ];

      for (const testCase of testCases) {
        const testPolicy = { ...mockPolicy, '@odata.type': testCase.odataType };
        mockClient.responses.set('/deviceManagement/deviceCompliancePolicies', {
          value: [testPolicy]
        });

        const policies = await report.getAllCompliancePolicies();
        expect(policies[0].platform).toBe(testCase.expected);
      }
    });
  });
});

// ============================================================================
// Run Tests
// ============================================================================

// Export for test runner
export default describe;
