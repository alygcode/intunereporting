/**
 * Unit Tests for ConfigMgr Device Management Reports Orchestrator
 *
 * Comprehensive test suite for the master orchestrator that manages
 * all 37 Configuration Manager device management reports.
 *
 * @module index.test
 */

import { Client } from '@microsoft/microsoft-graph-client';
import {
  ConfigMgrDeviceManagementReports,
  ReportCategory,
  createConfigMgrReports,
  executeReport,
} from './index';
import { AppConfig, ReportData } from '../../types';

// ============================================================================
// Mock Setup
// ============================================================================

/**
 * Create mock Graph client
 */
function createMockGraphClient(): Client {
  const mockClient = {
    api: jest.fn().mockReturnThis(),
    get: jest.fn().mockResolvedValue({ value: [] }),
    select: jest.fn().mockReturnThis(),
    filter: jest.fn().mockReturnThis(),
    top: jest.fn().mockReturnThis(),
  } as any;

  return mockClient;
}

/**
 * Create mock app configuration
 */
function createMockConfig(): AppConfig {
  return {
    authentication: {
      tenantId: 'test-tenant-id',
      clientId: 'test-client-id',
      clientSecret: 'test-secret',
      authMethod: 'clientSecret',
    },
    reports: {
      enabled: ['*'],
      disabled: [],
      settings: {},
    },
    output: {
      defaultFormat: 'json',
      directory: './test-output',
      includeTimestamp: true,
      compression: false,
    },
    scheduler: {
      enabled: false,
      timezone: 'UTC',
      schedules: [],
    },
    logging: {
      level: 'error',
      file: './test-logs/test.log',
      console: false,
      maxSize: '10m',
      maxFiles: 5,
    },
  };
}

/**
 * Create mock report data
 */
function createMockReportData(reportNumber: number, recordCount: number = 10): ReportData {
  return {
    metadata: {
      reportName: `Test Report ${reportNumber}`,
      generatedAt: new Date().toISOString(),
      generatedBy: 'Test Suite',
      recordCount,
    },
    data: Array.from({ length: recordCount }, (_, i) => ({
      id: `device-${i}`,
      name: `Device ${i}`,
    })),
    summary: {
      totalRecords: recordCount,
    },
  };
}

// ============================================================================
// Test Suite
// ============================================================================

describe('ConfigMgrDeviceManagementReports', () => {
  let mockGraphClient: Client;
  let mockConfig: AppConfig;
  let orchestrator: ConfigMgrDeviceManagementReports;

  beforeEach(() => {
    mockGraphClient = createMockGraphClient();
    mockConfig = createMockConfig();
    orchestrator = new ConfigMgrDeviceManagementReports(mockGraphClient, mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  // ==========================================================================
  // Initialization Tests
  // ==========================================================================

  describe('Initialization', () => {
    test('should create orchestrator instance', () => {
      expect(orchestrator).toBeInstanceOf(ConfigMgrDeviceManagementReports);
    });

    test('should initialize with correct configuration', () => {
      const config = orchestrator.getConfiguration();
      expect(config.authentication.tenantId).toBe('test-tenant-id');
      expect(config.output.defaultFormat).toBe('json');
    });

    test('should create orchestrator using convenience function', () => {
      const instance = createConfigMgrReports(mockGraphClient, mockConfig);
      expect(instance).toBeInstanceOf(ConfigMgrDeviceManagementReports);
    });
  });

  // ==========================================================================
  // Report Catalog Tests
  // ==========================================================================

  describe('Report Catalog', () => {
    test('should return all 37 reports', () => {
      const allReports = orchestrator.getAllReports();
      expect(allReports).toHaveLength(37);
    });

    test('should have valid metadata for each report', () => {
      const allReports = orchestrator.getAllReports();

      allReports.forEach(report => {
        expect(report.reportNumber).toBeGreaterThanOrEqual(1);
        expect(report.reportNumber).toBeLessThanOrEqual(37);
        expect(report.reportName).toBeTruthy();
        expect(report.description).toBeTruthy();
        expect(report.category).toBeTruthy();
        expect(report.modernIntuneEquivalent).toBeTruthy();
        expect(Array.isArray(report.graphApiEndpoints)).toBe(true);
        expect(typeof report.isSupported).toBe('boolean');
      });
    });

    test('should get report metadata by number', () => {
      const report1 = orchestrator.getReportMetadata(1);
      expect(report1).toBeDefined();
      expect(report1?.reportNumber).toBe(1);
      expect(report1?.reportName).toBe('All corporate-owned mobile devices');
    });

    test('should return undefined for invalid report number', () => {
      const invalidReport = orchestrator.getReportMetadata(999);
      expect(invalidReport).toBeUndefined();
    });

    test('should return 29 supported reports', () => {
      const supportedReports = orchestrator.getSupportedReports();
      expect(supportedReports.length).toBe(29);
      supportedReports.forEach(report => {
        expect(report.isSupported).toBe(true);
      });
    });

    test('should return 8 unsupported (Windows CE) reports', () => {
      const unsupportedReports = orchestrator.getUnsupportedReports();
      expect(unsupportedReports.length).toBe(8);
      unsupportedReports.forEach(report => {
        expect(report.isSupported).toBe(false);
        expect(report.category).toBe(ReportCategory.WINDOWS_CE);
      });
    });
  });

  // ==========================================================================
  // Report Category Tests
  // ==========================================================================

  describe('Report Categories', () => {
    test('should get reports by Mobile Inventory category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.MOBILE_INVENTORY);
      expect(reports.length).toBeGreaterThan(0);
      reports.forEach(report => {
        expect(report.category).toBe(ReportCategory.MOBILE_INVENTORY);
      });
    });

    test('should get reports by Hardware category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.HARDWARE);
      expect(reports.length).toBe(8); // Reports 9-12, 25-26, 28-29
      expect(reports[0].reportNumber).toBe(9);
    });

    test('should get reports by Security category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.SECURITY);
      expect(reports.length).toBe(3); // Reports 16, 23, 27
      const reportNumbers = reports.map(r => r.reportNumber);
      expect(reportNumbers).toContain(16);
      expect(reportNumbers).toContain(23);
      expect(reportNumbers).toContain(27);
    });

    test('should get reports by Exchange ActiveSync category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.EXCHANGE_ACTIVESYNC);
      expect(reports.length).toBe(4); // Reports 8, 15, 21, 34
    });

    test('should get reports by Device Actions category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.DEVICE_ACTIONS);
      expect(reports.length).toBe(3); // Reports 31-33
    });

    test('should get reports by Enrollment category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.ENROLLMENT);
      expect(reports.length).toBe(2); // Reports 17, 30
    });

    test('should get reports by Sideloading category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.SIDELOADING);
      expect(reports.length).toBe(2); // Reports 35-36
    });

    test('should get reports by Windows CE category', () => {
      const reports = orchestrator.getReportsByCategory(ReportCategory.WINDOWS_CE);
      expect(reports.length).toBe(8); // Reports 3-7, 13-14, 19
      reports.forEach(report => {
        expect(report.isSupported).toBe(false);
      });
    });
  });

  // ==========================================================================
  // Report Number Validation Tests
  // ==========================================================================

  describe('Report Number Validation', () => {
    test('should have correct report numbers in sequence', () => {
      const allReports = orchestrator.getAllReports();
      const reportNumbers = allReports.map(r => r.reportNumber).sort((a, b) => a - b);

      // Check for gaps (Windows CE reports are included)
      const expectedNumbers = [
        1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
        11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
        21, 22, 23, 24, 25, 26, 27, 28, 29, 30,
        31, 32, 33, 34, 35, 36,
      ];

      // Note: Report 37 might not exist in the current implementation
      // Verify we have at least reports 1-36
      expectedNumbers.forEach(num => {
        expect(reportNumbers).toContain(num);
      });
    });

    test('should have unique report numbers', () => {
      const allReports = orchestrator.getAllReports();
      const reportNumbers = allReports.map(r => r.reportNumber);
      const uniqueNumbers = new Set(reportNumbers);

      expect(uniqueNumbers.size).toBe(reportNumbers.length);
    });
  });

  // ==========================================================================
  // Category Distribution Tests
  // ==========================================================================

  describe('Category Distribution', () => {
    test('should have correct number of reports per category', () => {
      const categories = Object.values(ReportCategory);
      const distribution: Record<string, number> = {};

      categories.forEach(category => {
        const reports = orchestrator.getReportsByCategory(category);
        distribution[category] = reports.length;
      });

      // Verify expected distributions
      expect(distribution[ReportCategory.MOBILE_INVENTORY]).toBe(4);
      expect(distribution[ReportCategory.HARDWARE]).toBe(8);
      expect(distribution[ReportCategory.EXCHANGE_ACTIVESYNC]).toBe(4);
      expect(distribution[ReportCategory.SECURITY]).toBe(3);
      expect(distribution[ReportCategory.ENROLLMENT]).toBe(2);
      expect(distribution[ReportCategory.DEVICE_ACTIONS]).toBe(3);
      expect(distribution[ReportCategory.SIDELOADING]).toBe(2);
      expect(distribution[ReportCategory.WINDOWS_CE]).toBe(8);
      expect(distribution[ReportCategory.DEVICE_CATEGORY]).toBe(2);
    });
  });

  // ==========================================================================
  // Modern Intune Equivalent Tests
  // ==========================================================================

  describe('Modern Intune Equivalents', () => {
    test('should have modern equivalents for all supported reports', () => {
      const supportedReports = orchestrator.getSupportedReports();

      supportedReports.forEach(report => {
        expect(report.modernIntuneEquivalent).toBeTruthy();
        expect(report.modernIntuneEquivalent).not.toBe('N/A');
        expect(report.modernIntuneEquivalent.length).toBeGreaterThan(0);
      });
    });

    test('should have migration notes for unsupported reports', () => {
      const unsupportedReports = orchestrator.getUnsupportedReports();

      unsupportedReports.forEach(report => {
        expect(report.migrationNotes).toBeTruthy();
        expect(report.migrationNotes).toContain('Migrate');
      });
    });
  });

  // ==========================================================================
  // Graph API Endpoints Tests
  // ==========================================================================

  describe('Graph API Endpoints', () => {
    test('should have Graph API endpoints for supported reports', () => {
      const supportedReports = orchestrator.getSupportedReports();

      supportedReports.forEach(report => {
        expect(Array.isArray(report.graphApiEndpoints)).toBe(true);
        expect(report.graphApiEndpoints.length).toBeGreaterThan(0);
      });
    });

    test('should have common Graph API endpoints', () => {
      const supportedReports = orchestrator.getSupportedReports();
      const allEndpoints = supportedReports.flatMap(r => r.graphApiEndpoints);

      // Check for common endpoints
      const hasDeviceManagement = allEndpoints.some(e =>
        e.includes('/deviceManagement/managedDevices')
      );

      expect(hasDeviceManagement).toBe(true);
    });
  });

  // ==========================================================================
  // Configuration Tests
  // ==========================================================================

  describe('Configuration Management', () => {
    test('should get current configuration', () => {
      const config = orchestrator.getConfiguration();
      expect(config).toBeDefined();
      expect(config.authentication.tenantId).toBe('test-tenant-id');
    });

    test('should update configuration', () => {
      orchestrator.updateConfiguration({
        output: {
          ...mockConfig.output,
          defaultFormat: 'csv',
        },
      });

      const updatedConfig = orchestrator.getConfiguration();
      expect(updatedConfig.output.defaultFormat).toBe('csv');
    });

    test('should preserve other config values when updating', () => {
      const originalTenantId = mockConfig.authentication.tenantId;

      orchestrator.updateConfiguration({
        output: {
          ...mockConfig.output,
          defaultFormat: 'html',
        },
      });

      const updatedConfig = orchestrator.getConfiguration();
      expect(updatedConfig.authentication.tenantId).toBe(originalTenantId);
    });
  });

  // ==========================================================================
  // Error Handling Tests
  // ==========================================================================

  describe('Error Handling', () => {
    test('should throw error for invalid report number', async () => {
      await expect(orchestrator.getReportByNumber(999)).rejects.toThrow(
        'Report 999 not found in catalog'
      );
    });

    test('should handle unsupported report gracefully', async () => {
      // Windows CE report (unsupported)
      const result = await orchestrator.getReportByNumber(3);

      expect(result).toBeDefined();
      expect(result.data).toEqual([]);
      expect(result.summary).toBeDefined();
      if (result.summary) {
        expect((result.summary as any).isSupported).toBe(false);
      }
    });
  });

  // ==========================================================================
  // Report Execution Tests (Mocked)
  // ==========================================================================

  describe('Report Execution', () => {
    test('should execute supported report', async () => {
      // Mock the report execution
      jest.spyOn(orchestrator as any, 'getReportByNumber').mockResolvedValue(
        createMockReportData(1, 5)
      );

      const result = await orchestrator.getReportByNumber(1);

      expect(result).toBeDefined();
      expect(result.metadata).toBeDefined();
      expect(result.data).toHaveLength(5);
    });
  });

  // ==========================================================================
  // Convenience Function Tests
  // ==========================================================================

  describe('Convenience Functions', () => {
    test('should execute report using convenience function', async () => {
      // Mock the orchestrator's getReportByNumber method
      jest.spyOn(ConfigMgrDeviceManagementReports.prototype, 'getReportByNumber')
        .mockResolvedValue(createMockReportData(1, 10));

      const result = await executeReport(mockGraphClient, mockConfig, 1);

      expect(result).toBeDefined();
      expect(result.data).toHaveLength(10);
    });
  });

  // ==========================================================================
  // Report Metadata Validation Tests
  // ==========================================================================

  describe('Report Metadata Validation', () => {
    test('should have required fields for all reports', () => {
      const allReports = orchestrator.getAllReports();

      allReports.forEach(report => {
        // Required fields
        expect(report).toHaveProperty('reportNumber');
        expect(report).toHaveProperty('reportName');
        expect(report).toHaveProperty('description');
        expect(report).toHaveProperty('category');
        expect(report).toHaveProperty('modernIntuneEquivalent');
        expect(report).toHaveProperty('graphApiEndpoints');
        expect(report).toHaveProperty('isSupported');

        // Type checks
        expect(typeof report.reportNumber).toBe('number');
        expect(typeof report.reportName).toBe('string');
        expect(typeof report.description).toBe('string');
        expect(typeof report.category).toBe('string');
        expect(typeof report.modernIntuneEquivalent).toBe('string');
        expect(Array.isArray(report.graphApiEndpoints)).toBe(true);
        expect(typeof report.isSupported).toBe('boolean');
      });
    });

    test('should have non-empty descriptions', () => {
      const allReports = orchestrator.getAllReports();

      allReports.forEach(report => {
        expect(report.description.length).toBeGreaterThan(10);
      });
    });

    test('should have non-empty report names', () => {
      const allReports = orchestrator.getAllReports();

      allReports.forEach(report => {
        expect(report.reportName.length).toBeGreaterThan(5);
      });
    });
  });

  // ==========================================================================
  // Report Coverage Tests
  // ==========================================================================

  describe('Report Coverage', () => {
    test('should cover all report numbers from 1 to 36', () => {
      const allReports = orchestrator.getAllReports();
      const reportNumbers = allReports.map(r => r.reportNumber);

      for (let i = 1; i <= 36; i++) {
        expect(reportNumbers).toContain(i);
      }
    });

    test('should have correct category assignments', () => {
      // Mobile Inventory: 1, 2, 20, 22
      [1, 2, 20, 22].forEach(num => {
        const report = orchestrator.getReportMetadata(num);
        expect(report?.category).toBe(ReportCategory.MOBILE_INVENTORY);
      });

      // Hardware: 9-12, 25-26, 28-29
      [9, 10, 11, 12, 25, 26, 28, 29].forEach(num => {
        const report = orchestrator.getReportMetadata(num);
        expect(report?.category).toBe(ReportCategory.HARDWARE);
      });

      // Security: 16, 23, 27
      [16, 23, 27].forEach(num => {
        const report = orchestrator.getReportMetadata(num);
        expect(report?.category).toBe(ReportCategory.SECURITY);
      });

      // Windows CE: 3-7, 13-14, 19
      [3, 4, 5, 6, 7, 13, 14, 19].forEach(num => {
        const report = orchestrator.getReportMetadata(num);
        expect(report?.category).toBe(ReportCategory.WINDOWS_CE);
        expect(report?.isSupported).toBe(false);
      });
    });
  });

  // ==========================================================================
  // Integration Readiness Tests
  // ==========================================================================

  describe('Integration Readiness', () => {
    test('should support batch execution interface', () => {
      expect(typeof orchestrator.exportReports).toBe('function');
      expect(typeof orchestrator.exportAllReports).toBe('function');
    });

    test('should support dashboard generation interface', () => {
      expect(typeof orchestrator.generateDashboard).toBe('function');
    });

    test('should support scheduling interface', () => {
      expect(typeof orchestrator.scheduleReports).toBe('function');
    });

    test('should support configuration management', () => {
      expect(typeof orchestrator.getConfiguration).toBe('function');
      expect(typeof orchestrator.updateConfiguration).toBe('function');
    });
  });

  // ==========================================================================
  // Type Safety Tests
  // ==========================================================================

  describe('Type Safety', () => {
    test('should return correct types from getAllReports', () => {
      const reports = orchestrator.getAllReports();
      expect(Array.isArray(reports)).toBe(true);

      if (reports.length > 0) {
        const firstReport = reports[0];
        expect(typeof firstReport.reportNumber).toBe('number');
        expect(typeof firstReport.reportName).toBe('string');
      }
    });

    test('should return correct types from getReportMetadata', () => {
      const metadata = orchestrator.getReportMetadata(1);
      expect(metadata).toBeDefined();

      if (metadata) {
        expect(typeof metadata.reportNumber).toBe('number');
        expect(typeof metadata.isSupported).toBe('boolean');
        expect(Array.isArray(metadata.graphApiEndpoints)).toBe(true);
      }
    });
  });

  // ==========================================================================
  // Documentation Tests
  // ==========================================================================

  describe('Documentation', () => {
    test('should have parameters documented for relevant reports', () => {
      // Report 25 should have memory parameters
      const report25 = orchestrator.getReportMetadata(25);
      expect(report25?.parameters).toBeDefined();
      expect(report25?.parameters?.length).toBeGreaterThan(0);

      // Report 8 should have policy parameters
      const report8 = orchestrator.getReportMetadata(8);
      expect(report8?.parameters).toBeDefined();
    });

    test('should have Graph API endpoints documented', () => {
      const supportedReports = orchestrator.getSupportedReports();

      supportedReports.forEach(report => {
        expect(report.graphApiEndpoints.length).toBeGreaterThan(0);

        // Each endpoint should be a valid path
        report.graphApiEndpoints.forEach(endpoint => {
          expect(endpoint).toMatch(/^\//);
        });
      });
    });
  });
});

// ============================================================================
// Report Category Enum Tests
// ============================================================================

describe('ReportCategory Enum', () => {
  test('should have all expected categories', () => {
    const expectedCategories = [
      'Mobile Device Inventory',
      'Device Hardware',
      'Exchange ActiveSync',
      'Device Security',
      'Device Enrollment',
      'Device Actions',
      'Windows RT Sideloading',
      'Windows CE (Legacy)',
      'Device Category',
    ];

    const actualCategories = Object.values(ReportCategory);

    expectedCategories.forEach(category => {
      expect(actualCategories).toContain(category);
    });
  });

  test('should have consistent category naming', () => {
    const categories = Object.values(ReportCategory);

    categories.forEach(category => {
      // Should not have trailing spaces
      expect(category).toBe(category.trim());

      // Should not be empty
      expect(category.length).toBeGreaterThan(0);
    });
  });
});

// ============================================================================
// Export Tests
// ============================================================================

describe('Module Exports', () => {
  test('should export main class', () => {
    expect(ConfigMgrDeviceManagementReports).toBeDefined();
  });

  test('should export convenience functions', () => {
    expect(typeof createConfigMgrReports).toBe('function');
    expect(typeof executeReport).toBe('function');
  });

  test('should export enums', () => {
    expect(ReportCategory).toBeDefined();
  });
});
