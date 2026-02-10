/**
 * Windows RT Sideloading Reports - Unit Tests
 *
 * Comprehensive test suite for the WindowsRTSideloadingReports module
 */

import {
  WindowsRTSideloadingReports,
  SideloadingKey,
  SideloadingKeyDetailedStatus,
  SideloadingKeysSummary,
} from './windows-rt-sideloading-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';

// ============================================================================
// Mock Data
// ============================================================================

const mockKey1: SideloadingKey = {
  id: 'key-001',
  displayName: 'Production Sideloading Key',
  value: 'ABCD-EFGH-IJKL-MNOP-QRST-UVWX-YZ12-3456',
  totalActivation: 1200,
  lastUpdatedDateTime: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString(), // 30 days ago
  description: 'Primary production sideloading key',
};

const mockKey2: SideloadingKey = {
  id: 'key-002',
  displayName: 'Development Sideloading Key',
  value: '1234-5678-9ABC-DEFG-HIJK-LMNO-PQRS-TUVW',
  totalActivation: 4500,
  lastUpdatedDateTime: new Date(Date.now() - 60 * 24 * 60 * 60 * 1000).toISOString(), // 60 days ago
  description: 'Development and testing key',
};

const mockKey3: SideloadingKey = {
  id: 'key-003',
  displayName: 'Expired Sideloading Key',
  value: 'XXXX-YYYY-ZZZZ-1111-2222-3333-4444-5555',
  totalActivation: 5000, // Max activations reached
  lastUpdatedDateTime: new Date(Date.now() - 400 * 24 * 60 * 60 * 1000).toISOString(), // 400 days ago
  description: 'Expired key - reached max activations',
};

const mockKey4: SideloadingKey = {
  id: 'key-004',
  displayName: 'Revoked Sideloading Key',
  value: '', // Empty value indicates revoked
  totalActivation: 100,
  lastUpdatedDateTime: new Date(Date.now() - 100 * 24 * 60 * 60 * 1000).toISOString(), // 100 days ago
  description: 'Revoked for security reasons',
};

const mockKey5: SideloadingKey = {
  id: 'key-005',
  displayName: 'Unused Sideloading Key',
  value: 'AAAA-BBBB-CCCC-DDDD-EEEE-FFFF-0000-1111',
  totalActivation: 50, // Very low usage
  lastUpdatedDateTime: new Date(Date.now() - 10 * 24 * 60 * 60 * 1000).toISOString(), // 10 days ago
  description: 'Recently created but unused',
};

const mockKey6: SideloadingKey = {
  id: 'key-006',
  displayName: 'Near Expiration Key',
  value: 'NEAR-EXPR-TION-KEY1-2345-6789-ABCD-EFGH',
  totalActivation: 3000,
  lastUpdatedDateTime: new Date(Date.now() - 340 * 24 * 60 * 60 * 1000).toISOString(), // 340 days ago
  description: 'Key approaching expiration',
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
    // All sideloading keys endpoint
    this.responses.set('/deviceAppManagement/sideLoadingKeys', {
      value: [mockKey1, mockKey2, mockKey3, mockKey4, mockKey5, mockKey6],
    });

    // Individual key endpoints
    this.responses.set(`/deviceAppManagement/sideLoadingKeys/${mockKey1.id}`, mockKey1);
    this.responses.set(`/deviceAppManagement/sideLoadingKeys/${mockKey2.id}`, mockKey2);
    this.responses.set(`/deviceAppManagement/sideLoadingKeys/${mockKey3.id}`, mockKey3);
    this.responses.set(`/deviceAppManagement/sideLoadingKeys/${mockKey4.id}`, mockKey4);
    this.responses.set(`/deviceAppManagement/sideLoadingKeys/${mockKey5.id}`, mockKey5);
    this.responses.set(`/deviceAppManagement/sideLoadingKeys/${mockKey6.id}`, mockKey6);
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
      enabled: ['windows-rt-sideloading-reports'],
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

describe('WindowsRTSideloadingReports', () => {
  let report: WindowsRTSideloadingReports;
  let mockClient: any;
  let config: AppConfig;

  beforeEach(() => {
    mockClient = new MockGraphClient();
    config = getMockConfig();
    report = new WindowsRTSideloadingReports(mockClient as unknown as Client, config);
  });

  // ==========================================================================
  // Test: getAllSideloadingKeys
  // ==========================================================================

  describe('getAllSideloadingKeys', () => {
    it('should retrieve all sideloading keys', async () => {
      const keys = await report.getAllSideloadingKeys();

      expect(keys).toBeDefined();
      expect(Array.isArray(keys)).toBe(true);
      expect(keys.length).toBe(6);
    });

    it('should parse key data correctly', async () => {
      const keys = await report.getAllSideloadingKeys();
      const key = keys[0];

      expect(key.id).toBe(mockKey1.id);
      expect(key.displayName).toBe(mockKey1.displayName);
      expect(key.value).toBe(mockKey1.value);
      expect(key.totalActivation).toBe(mockKey1.totalActivation);
      expect(key.lastUpdatedDateTime).toBe(mockKey1.lastUpdatedDateTime);
    });

    it('should include all required fields', async () => {
      const keys = await report.getAllSideloadingKeys();

      keys.forEach((key) => {
        expect(key.id).toBeDefined();
        expect(key.totalActivation).toBeDefined();
      });
    });
  });

  // ==========================================================================
  // Test: getSideloadingKeyById
  // ==========================================================================

  describe('getSideloadingKeyById', () => {
    it('should retrieve a specific sideloading key', async () => {
      const key = await report.getSideloadingKeyById(mockKey1.id);

      expect(key).toBeDefined();
      expect(key.id).toBe(mockKey1.id);
      expect(key.displayName).toBe(mockKey1.displayName);
    });

    it('should include all key details', async () => {
      const key = await report.getSideloadingKeyById(mockKey2.id);

      expect(key.id).toBe(mockKey2.id);
      expect(key.displayName).toBe(mockKey2.displayName);
      expect(key.value).toBe(mockKey2.value);
      expect(key.totalActivation).toBe(mockKey2.totalActivation);
      expect(key.description).toBe(mockKey2.description);
    });
  });

  // ==========================================================================
  // Test: getSideloadingKeyDetailedStatus (Report 35)
  // ==========================================================================

  describe('getSideloadingKeyDetailedStatus', () => {
    it('should generate detailed status for a key', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey1.id);

      expect(status).toBeDefined();
      expect(status.keyId).toBe(mockKey1.id);
      expect(status.displayName).toBe(mockKey1.displayName);
    });

    it('should calculate activation metrics correctly', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey1.id);

      expect(status.totalActivations).toBe(5000); // Max activations
      expect(status.activationsUsed).toBe(1200);
      expect(status.activationsRemaining).toBe(3800);
      expect(status.usagePercentage).toBe(24); // 1200/5000 = 24%
    });

    it('should mask key value for security', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey1.id);

      expect(status.keyValue).not.toBe(mockKey1.value);
      expect(status.keyValue).toContain('****');
      expect(status.keyValue.length).toBeLessThan(mockKey1.value!.length);
    });

    it('should identify active keys correctly', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey1.id);

      expect(status.status).toBe('active');
      expect(status.statusReason).toContain('active');
    });

    it('should identify expired keys (max activations)', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey3.id);

      expect(status.status).toBe('expired');
      expect(status.statusReason).toContain('maximum activation limit');
    });

    it('should identify revoked keys', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey4.id);

      expect(status.status).toBe('revoked');
      expect(status.statusReason).toContain('revoked');
    });

    it('should calculate days since last update', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey1.id);

      expect(status.lastUpdatedDaysAgo).toBeGreaterThanOrEqual(29);
      expect(status.lastUpdatedDaysAgo).toBeLessThanOrEqual(31);
    });

    it('should identify high usage keys', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey2.id);

      expect(status.usagePercentage).toBe(90); // 4500/5000 = 90%
      expect(status.isHighUsage).toBe(true);
    });

    it('should identify keys near expiration', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey6.id);

      expect(status.isNearExpiration).toBe(true);
      expect(status.expirationWarning).toBe(true);
    });

    it('should respect custom alert thresholds', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey2.id, {
        highUsagePercentage: 95, // Raise threshold
        expirationWarningDays: 60,
      });

      expect(status.isHighUsage).toBe(false); // 90% < 95%
    });
  });

  // ==========================================================================
  // Test: getSideloadingKeysSummary (Report 36)
  // ==========================================================================

  describe('getSideloadingKeysSummary', () => {
    it('should generate summary for all keys', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(summary).toBeDefined();
      expect(summary.totalKeys).toBe(6);
    });

    it('should count keys by status correctly', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(summary.activeKeysCount).toBeGreaterThan(0);
      expect(summary.expiredKeysCount).toBeGreaterThan(0);
      expect(summary.revokedKeysCount).toBeGreaterThan(0);

      // Total should match
      const total =
        summary.activeKeysCount +
        summary.expiredKeysCount +
        summary.revokedKeysCount +
        summary.unknownKeysCount;
      expect(total).toBe(6);
    });

    it('should provide status breakdown with percentages', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(Array.isArray(summary.keysByStatus)).toBe(true);
      expect(summary.keysByStatus.length).toBeGreaterThan(0);

      summary.keysByStatus.forEach((statusGroup) => {
        expect(statusGroup.status).toBeDefined();
        expect(statusGroup.count).toBeGreaterThan(0);
        expect(statusGroup.percentage).toBeGreaterThanOrEqual(0);
        expect(statusGroup.percentage).toBeLessThanOrEqual(100);
        expect(Array.isArray(statusGroup.keys)).toBe(true);
      });
    });

    it('should calculate usage statistics correctly', async () => {
      const summary = await report.getSideloadingKeysSummary();
      const stats = summary.usageStatistics;

      expect(stats.totalActivationsAcrossAllKeys).toBe(
        1200 + 4500 + 5000 + 100 + 50 + 3000
      );
      expect(stats.averageActivationsPerKey).toBeGreaterThan(0);
      expect(stats.averageUsagePercentage).toBeGreaterThan(0);
      expect(stats.averageUsagePercentage).toBeLessThanOrEqual(100);
    });

    it('should identify most used key', async () => {
      const summary = await report.getSideloadingKeysSummary();
      const mostUsed = summary.usageStatistics.mostUsedKey;

      expect(mostUsed).not.toBeNull();
      expect(mostUsed!.keyId).toBe(mockKey3.id); // 5000 activations
      expect(mostUsed!.activations).toBe(5000);
    });

    it('should identify least used key', async () => {
      const summary = await report.getSideloadingKeysSummary();
      const leastUsed = summary.usageStatistics.leastUsedKey;

      expect(leastUsed).not.toBeNull();
      expect(leastUsed!.keyId).toBe(mockKey5.id); // 50 activations
      expect(leastUsed!.activations).toBe(50);
    });

    it('should generate alerts for expiring keys', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(summary.alerts.expiringKeysCount).toBeGreaterThan(0);
      expect(Array.isArray(summary.alerts.expiringKeys)).toBe(true);
      expect(summary.alerts.expiringKeys.length).toBe(summary.alerts.expiringKeysCount);
    });

    it('should generate alerts for high usage keys', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(summary.alerts.highUsageKeysCount).toBeGreaterThan(0);
      expect(Array.isArray(summary.alerts.highUsageKeys)).toBe(true);
      expect(summary.alerts.highUsageKeys.length).toBe(summary.alerts.highUsageKeysCount);
    });

    it('should generate alerts for unused keys', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(summary.alerts.unusedKeysCount).toBeGreaterThan(0);
      expect(Array.isArray(summary.alerts.unusedKeys)).toBe(true);
      expect(summary.alerts.unusedKeys.length).toBe(summary.alerts.unusedKeysCount);
    });

    it('should include timestamp', async () => {
      const summary = await report.getSideloadingKeysSummary();

      expect(summary.lastUpdatedDateTime).toBeDefined();
      const timestamp = new Date(summary.lastUpdatedDateTime);
      expect(timestamp.getTime()).toBeLessThanOrEqual(Date.now());
    });

    it('should filter by status', async () => {
      const summary = await report.getSideloadingKeysSummary({
        status: 'active',
      });

      expect(summary.totalKeys).toBeLessThan(6);
      expect(summary.expiredKeysCount).toBe(0);
      expect(summary.revokedKeysCount).toBe(0);
    });

    it('should filter by minimum activations', async () => {
      const summary = await report.getSideloadingKeysSummary({
        minActivations: 1000,
      });

      expect(summary.totalKeys).toBeLessThan(6);
      // All keys should have >= 1000 activations
    });

    it('should filter by maximum activations', async () => {
      const summary = await report.getSideloadingKeysSummary({
        maxActivations: 1000,
      });

      expect(summary.totalKeys).toBeLessThan(6);
      // All keys should have <= 1000 activations
    });

    it('should filter by usage percentage', async () => {
      const summary = await report.getSideloadingKeysSummary({
        minUsagePercentage: 50,
      });

      expect(summary.totalKeys).toBeLessThan(6);
    });

    it('should exclude expired keys when filtered', async () => {
      const summary = await report.getSideloadingKeysSummary({
        includeExpired: false,
      });

      expect(summary.expiredKeysCount).toBe(0);
    });

    it('should exclude revoked keys when filtered', async () => {
      const summary = await report.getSideloadingKeysSummary({
        includeRevoked: false,
      });

      expect(summary.revokedKeysCount).toBe(0);
    });

    it('should respect custom alert thresholds', async () => {
      const summary = await report.getSideloadingKeysSummary(undefined, {
        highUsagePercentage: 95,
        lowUsagePercentage: 5,
        expirationWarningDays: 10,
      });

      expect(summary.alerts).toBeDefined();
      // Alert counts may differ with custom thresholds
    });
  });

  // ==========================================================================
  // Test: generateComprehensiveReport
  // ==========================================================================

  describe('generateComprehensiveReport', () => {
    it('should generate complete sideloading key report', async () => {
      const reportData = await report.generateComprehensiveReport();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include metadata in report', async () => {
      const reportData = await report.generateComprehensiveReport();
      const metadata = reportData.metadata;

      expect(metadata.reportName).toBe('windows-rt-sideloading-reports');
      expect(metadata.generatedAt).toBeDefined();
      expect(metadata.generatedBy).toBe('Intune Reporting Dashboard');
      expect(metadata.recordCount).toBe(6);
    });

    it('should include summary and detailed statuses', async () => {
      const reportData = await report.generateComprehensiveReport();

      expect(reportData.data.summary).toBeDefined();
      expect(reportData.data.detailedStatuses).toBeDefined();
      expect(reportData.data.keys).toBeDefined();

      expect(Array.isArray(reportData.data.detailedStatuses)).toBe(true);
      expect(reportData.data.detailedStatuses.length).toBe(6);
    });

    it('should include all summary statistics', async () => {
      const reportData = await report.generateComprehensiveReport();
      const summary = reportData.summary;

      expect(summary.totalKeys).toBe(6);
      expect(summary.activeKeys).toBeDefined();
      expect(summary.expiredKeys).toBeDefined();
      expect(summary.alerts).toBeDefined();
      expect(summary.generatedAt).toBeDefined();
    });

    it('should respect filter options', async () => {
      const reportData = await report.generateComprehensiveReport({
        status: 'active',
      });

      expect(reportData.data.summary.expiredKeysCount).toBe(0);
    });

    it('should respect custom thresholds', async () => {
      const reportData = await report.generateComprehensiveReport(undefined, {
        highUsagePercentage: 95,
      });

      expect(reportData).toBeDefined();
      // Alert counts may differ with custom thresholds
    });
  });

  // ==========================================================================
  // Test: execute (Full Report)
  // ==========================================================================

  describe('execute', () => {
    it('should generate complete sideloading key report', async () => {
      const reportData = await report.execute();

      expect(reportData).toBeDefined();
      expect(reportData.metadata).toBeDefined();
      expect(reportData.data).toBeDefined();
      expect(reportData.summary).toBeDefined();
    });

    it('should include metadata in report', async () => {
      const reportData = await report.execute();
      const metadata = reportData.metadata;

      expect(metadata.reportName).toBe('windows-rt-sideloading-reports');
      expect(metadata.generatedAt).toBeDefined();
      expect(metadata.generatedBy).toBe('Intune Reporting Dashboard');
    });

    it('should include summary statistics', async () => {
      const reportData = await report.execute();
      const summary = reportData.summary;

      expect(summary.totalKeys).toBe(6);
      expect(summary.activeKeys).toBeDefined();
      expect(summary.alerts).toBeDefined();
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

      const errorReport = new WindowsRTSideloadingReports(
        errorClient as unknown as Client,
        config
      );

      await expect(errorReport.getAllSideloadingKeys()).rejects.toThrow();
    });

    it('should handle empty key list', async () => {
      mockClient.responses.set('/deviceAppManagement/sideLoadingKeys', { value: [] });

      const keys = await report.getAllSideloadingKeys();

      expect(keys.length).toBe(0);
    });

    it('should handle keys with missing optional fields', async () => {
      const minimalKey: SideloadingKey = {
        id: 'key-999',
        totalActivation: 100,
      };

      mockClient.responses.set('/deviceAppManagement/sideLoadingKeys', {
        value: [minimalKey],
      });

      const summary = await report.getSideloadingKeysSummary();

      expect(summary.totalKeys).toBe(1);
    });

    it('should handle invalid key ID gracefully', async () => {
      const errorClient = {
        api: (endpoint: string) => ({
          select: () => ({
            get: async () => {
              throw new Error('Key not found');
            },
          }),
        }),
      };

      const errorReport = new WindowsRTSideloadingReports(
        errorClient as unknown as Client,
        config
      );

      await expect(errorReport.getSideloadingKeyById('invalid-id')).rejects.toThrow();
    });
  });

  // ==========================================================================
  // Test: Usage Percentage Calculations
  // ==========================================================================

  describe('Usage Percentage Calculations', () => {
    it('should calculate usage percentage correctly for low usage', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey5.id);

      expect(status.usagePercentage).toBe(1); // 50/5000 = 1%
      expect(status.activationsRemaining).toBe(4950);
    });

    it('should calculate usage percentage correctly for high usage', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey2.id);

      expect(status.usagePercentage).toBe(90); // 4500/5000 = 90%
      expect(status.activationsRemaining).toBe(500);
    });

    it('should calculate usage percentage correctly for maxed out key', async () => {
      const status = await report.getSideloadingKeyDetailedStatus(mockKey3.id);

      expect(status.usagePercentage).toBe(100); // 5000/5000 = 100%
      expect(status.activationsRemaining).toBe(0);
    });
  });
});

// ============================================================================
// Run Tests
// ============================================================================

// Export for test runner
export default describe;
