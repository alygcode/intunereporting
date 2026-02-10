/**
 * Device Category and Enrollment Reports - Unit Tests
 *
 * Comprehensive test suite for DeviceCategoryEnrollmentReports module
 */

import {
  DeviceCategoryEnrollmentReports,
  EnrollmentFailureType,
  CertificateStatus,
} from './device-category-enrollment-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';

// ============================================================================
// Mock Data
// ============================================================================

const mockCategories = [
  {
    id: 'cat-001',
    displayName: 'Corporate',
    description: 'Corporate-owned devices',
    lastModifiedDateTime: '2024-01-15T10:00:00Z',
    roleScopeTagIds: ['0'],
  },
  {
    id: 'cat-002',
    displayName: 'Personal',
    description: 'Personal devices',
    lastModifiedDateTime: '2024-01-15T10:00:00Z',
    roleScopeTagIds: ['0'],
  },
  {
    id: 'cat-003',
    displayName: 'Kiosk',
    description: 'Kiosk devices',
    lastModifiedDateTime: '2024-01-15T10:00:00Z',
    roleScopeTagIds: ['0'],
  },
];

const mockHealthyDevice = {
  id: 'device-001',
  deviceName: 'iPhone 14 Pro',
  deviceCategoryDisplayName: 'Corporate',
  userPrincipalName: 'user1@contoso.com',
  userDisplayName: 'User One',
  userId: 'user-001',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  model: 'iPhone 14 Pro',
  manufacturer: 'Apple Inc.',
  serialNumber: 'SN001',
  imei: '123456789012345',
  enrolledDateTime: '2024-01-15T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:00:00Z',
  managementState: 'managed',
  enrollmentState: 'enrolled',
  complianceState: 'compliant',
  managedDeviceOwnerType: 'company',
  enrollmentType: 'userEnrollment',
  isSupervised: false,
  azureADDeviceId: 'aad-device-001',
  certificateExpirationDateTime: '2025-01-15T10:00:00Z',
};

const mockCertificateFailureDevice = {
  id: 'device-002',
  deviceName: 'Samsung Galaxy S23',
  deviceCategoryDisplayName: 'Corporate',
  userPrincipalName: 'user2@contoso.com',
  userDisplayName: 'User Two',
  userId: 'user-002',
  operatingSystem: 'Android',
  osVersion: '14.0',
  model: 'Galaxy S23',
  manufacturer: 'Samsung',
  serialNumber: 'SN002',
  imei: '123456789012346',
  enrolledDateTime: '2024-01-20T10:00:00Z',
  lastSyncDateTime: '2024-01-25T10:00:00Z',
  managementState: 'unhealthy',
  enrollmentState: 'enrolled',
  complianceState: 'noncompliant',
  managedDeviceOwnerType: 'company',
  enrollmentType: 'deviceEnrollmentManager',
  certificateExpirationDateTime: '2023-12-01T10:00:00Z', // Expired certificate
};

const mockAssignmentFailureDevice = {
  id: 'device-003',
  deviceName: 'iPad Pro',
  deviceCategoryDisplayName: 'Personal',
  userPrincipalName: 'user3@contoso.com',
  userDisplayName: 'User Three',
  userId: 'user-003',
  operatingSystem: 'iOS',
  osVersion: '17.2.1',
  model: 'iPad Pro',
  manufacturer: 'Apple Inc.',
  serialNumber: 'SN003',
  enrolledDateTime: '2024-02-01T10:00:00Z',
  lastSyncDateTime: null, // Never synced
  managementState: 'discovered',
  enrollmentState: 'enrolled',
  complianceState: 'unknown',
  managedDeviceOwnerType: 'personal',
  enrollmentType: 'userEnrollment',
};

const mockSyncFailureDevice = {
  id: 'device-004',
  deviceName: 'Pixel 7',
  deviceCategoryDisplayName: 'Corporate',
  userPrincipalName: 'user4@contoso.com',
  userDisplayName: 'User Four',
  userId: 'user-004',
  operatingSystem: 'Android',
  osVersion: '14.0',
  model: 'Pixel 7',
  manufacturer: 'Google',
  serialNumber: 'SN004',
  enrolledDateTime: '2024-01-01T10:00:00Z',
  lastSyncDateTime: '2024-01-05T10:00:00Z', // More than 7 days ago
  managementState: 'managed',
  enrollmentState: 'enrolled',
  complianceState: 'compliant',
  managedDeviceOwnerType: 'company',
  certificateExpirationDateTime: '2025-01-15T10:00:00Z',
};

const mockPolicyFailureDevice = {
  id: 'device-005',
  deviceName: 'Surface Pro',
  deviceCategoryDisplayName: 'Kiosk',
  userPrincipalName: 'user5@contoso.com',
  userDisplayName: 'User Five',
  userId: 'user-005',
  operatingSystem: 'Windows',
  osVersion: '11.0',
  model: 'Surface Pro 9',
  manufacturer: 'Microsoft',
  serialNumber: 'SN005',
  enrolledDateTime: '2024-01-25T10:00:00Z',
  lastSyncDateTime: '2024-02-06T08:00:00Z',
  managementState: 'policyapplyfailed',
  enrollmentState: 'enrolled',
  complianceState: 'noncompliant',
  managedDeviceOwnerType: 'company',
  certificateExpirationDateTime: '2025-01-15T10:00:00Z',
};

const mockAuthFailureDevice = {
  id: 'device-006',
  deviceName: 'MacBook Pro',
  userPrincipalName: 'user6@contoso.com',
  userDisplayName: 'User Six',
  userId: 'user-006',
  operatingSystem: 'macOS',
  osVersion: '14.2',
  model: 'MacBook Pro',
  manufacturer: 'Apple Inc.',
  serialNumber: 'SN006',
  enrolledDateTime: '2024-01-28T10:00:00Z',
  lastSyncDateTime: '2024-01-29T10:00:00Z',
  managementState: 'unhealthy',
  enrollmentState: 'failed',
  complianceState: 'unknown',
  managedDeviceOwnerType: 'company',
  errorCode: '80180014', // Authentication error code
};

const mockConfig: AppConfig = {
  auth: {
    clientId: 'test-client-id',
    clientSecret: 'test-secret',
    tenantId: 'test-tenant',
  },
  output: {
    directory: './test-reports',
    format: 'json',
  },
};

// ============================================================================
// Mock Graph Client
// ============================================================================

function createMockGraphClient(mockResponses: Record<string, any>): Client {
  const mockClient = {
    api: jest.fn((path: string) => ({
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      get: jest.fn(async () => {
        if (mockResponses[path]) {
          return mockResponses[path];
        }
        return { value: [] };
      }),
    })),
  } as unknown as Client;

  return mockClient;
}

// ============================================================================
// Test Suite
// ============================================================================

describe('DeviceCategoryEnrollmentReports', () => {
  let report: DeviceCategoryEnrollmentReports;
  let mockGraphClient: Client;

  beforeEach(() => {
    const mockResponses = {
      '/deviceManagement/deviceCategories': {
        value: mockCategories,
      },
      '/deviceManagement/managedDevices': {
        value: [
          mockHealthyDevice,
          mockCertificateFailureDevice,
          mockAssignmentFailureDevice,
          mockSyncFailureDevice,
          mockPolicyFailureDevice,
          mockAuthFailureDevice,
        ],
      },
      '/deviceManagement/deviceCategories/cat-001': mockCategories[0],
    };

    mockGraphClient = createMockGraphClient(mockResponses);
    report = new DeviceCategoryEnrollmentReports(mockGraphClient, mockConfig);
  });

  // ==========================================================================
  // Category Management Tests
  // ==========================================================================

  describe('getAllDeviceCategories', () => {
    it('should fetch all device categories', async () => {
      const categories = await report.getAllDeviceCategories();

      expect(categories).toHaveLength(3);
      expect(categories[0].displayName).toBe('Corporate');
      expect(categories[1].displayName).toBe('Personal');
      expect(categories[2].displayName).toBe('Kiosk');
    });

    it('should include category metadata', async () => {
      const categories = await report.getAllDeviceCategories();

      categories.forEach((category) => {
        expect(category).toHaveProperty('id');
        expect(category).toHaveProperty('displayName');
        expect(category).toHaveProperty('description');
        expect(category).toHaveProperty('lastModifiedDateTime');
      });
    });
  });

  describe('getDevicesByCategory', () => {
    it('should fetch devices for a specific category', async () => {
      const devices = await report.getDevicesByCategory('Corporate');

      expect(devices.length).toBeGreaterThan(0);
      devices.forEach((device) => {
        expect(device.deviceCategoryDisplayName).toBe('Corporate');
      });
    });

    it('should format device information correctly', async () => {
      const devices = await report.getDevicesByCategory('Corporate');

      devices.forEach((device) => {
        expect(device).toHaveProperty('deviceId');
        expect(device).toHaveProperty('deviceName');
        expect(device).toHaveProperty('platform');
        expect(device).toHaveProperty('osVersion');
        expect(device).toHaveProperty('managementState');
        expect(device).toHaveProperty('complianceState');
      });
    });
  });

  describe('getCategorySummary', () => {
    it('should generate summary for all categories', async () => {
      const summaries = await report.getCategorySummary();

      expect(summaries.length).toBeGreaterThan(0);
      summaries.forEach((summary) => {
        expect(summary).toHaveProperty('categoryId');
        expect(summary).toHaveProperty('categoryName');
        expect(summary).toHaveProperty('totalDevices');
        expect(summary).toHaveProperty('compliantDevices');
        expect(summary).toHaveProperty('platformBreakdown');
      });
    });

    it('should calculate device counts correctly', async () => {
      const summaries = await report.getCategorySummary();

      summaries.forEach((summary) => {
        expect(summary.totalDevices).toBeGreaterThanOrEqual(0);
        expect(summary.compliantDevices).toBeGreaterThanOrEqual(0);
        expect(summary.compliantDevices).toBeLessThanOrEqual(summary.totalDevices);
      });
    });

    it('should include platform breakdown', async () => {
      const summaries = await report.getCategorySummary();

      summaries.forEach((summary) => {
        expect(summary.platformBreakdown).toBeDefined();
        expect(typeof summary.platformBreakdown).toBe('object');
      });
    });
  });

  // ==========================================================================
  // Enrollment Failure Tests
  // ==========================================================================

  describe('getUnmanagedDevicesEnrollmentFailed', () => {
    it('should identify unmanaged devices', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      expect(unmanagedDevices.length).toBeGreaterThan(0);
      unmanagedDevices.forEach((device) => {
        expect(device).toHaveProperty('failureType');
        expect(device).toHaveProperty('failureCategory');
        expect(device).toHaveProperty('certificateStatus');
        expect(device).toHaveProperty('severity');
      });
    });

    it('should detect certificate failures', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      const certificateFailures = unmanagedDevices.filter(
        (d) => d.failureType === EnrollmentFailureType.CertificateFailure
      );

      expect(certificateFailures.length).toBeGreaterThan(0);
      certificateFailures.forEach((device) => {
        expect(device.certificateStatus).toBe(CertificateStatus.Expired);
      });
    });

    it('should detect assignment failures', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      const assignmentFailures = unmanagedDevices.filter(
        (d) => d.failureType === EnrollmentFailureType.AssignmentFailure
      );

      expect(assignmentFailures.length).toBeGreaterThan(0);
    });

    it('should detect sync failures', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      const syncFailures = unmanagedDevices.filter(
        (d) => d.failureType === EnrollmentFailureType.SyncFailure
      );

      expect(syncFailures.length).toBeGreaterThan(0);
    });

    it('should include troubleshooting steps', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      unmanagedDevices.forEach((device) => {
        expect(device.troubleshootingSteps).toBeDefined();
        expect(Array.isArray(device.troubleshootingSteps)).toBe(true);
        expect(device.troubleshootingSteps.length).toBeGreaterThan(0);
      });
    });

    it('should include recommendations', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      unmanagedDevices.forEach((device) => {
        expect(device.recommendation).toBeDefined();
        expect(typeof device.recommendation).toBe('string');
        expect(device.recommendation.length).toBeGreaterThan(0);
      });
    });

    it('should assign severity levels correctly', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      unmanagedDevices.forEach((device) => {
        expect(['Critical', 'High', 'Medium', 'Low']).toContain(device.severity);
      });

      // Certificate and auth failures should be Critical
      const criticalDevices = unmanagedDevices.filter((d) => d.severity === 'Critical');
      criticalDevices.forEach((device) => {
        expect([
          EnrollmentFailureType.CertificateFailure,
          EnrollmentFailureType.AuthenticationFailure,
        ]).toContain(device.failureType);
      });
    });
  });

  describe('analyzeEnrollmentFailures', () => {
    it('should provide comprehensive failure analysis', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      expect(analysis).toHaveProperty('totalFailures');
      expect(analysis).toHaveProperty('failuresByType');
      expect(analysis).toHaveProperty('failuresByCategory');
      expect(analysis).toHaveProperty('failuresBySeverity');
      expect(analysis).toHaveProperty('certificateIssues');
      expect(analysis).toHaveProperty('affectedUsers');
    });

    it('should count failures by type correctly', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      const totalByType = Object.values(analysis.failuresByType).reduce(
        (sum, count) => sum + count,
        0
      );
      expect(totalByType).toBe(analysis.totalFailures);
    });

    it('should count failures by severity correctly', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      const totalBySeverity = Object.values(analysis.failuresBySeverity).reduce(
        (sum, count) => sum + count,
        0
      );
      expect(totalBySeverity).toBe(analysis.totalFailures);
    });

    it('should identify certificate issues', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      expect(analysis.certificateIssues).toBeGreaterThanOrEqual(0);
      expect(analysis.certificateIssues).toBeLessThanOrEqual(analysis.totalFailures);
    });

    it('should track affected users', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      expect(analysis.affectedUsers).toBeGreaterThan(0);
      expect(analysis.affectedUsers).toBeLessThanOrEqual(analysis.totalFailures);
    });

    it('should provide recommended actions', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      expect(analysis.recommendedActions).toBeDefined();
      expect(Array.isArray(analysis.recommendedActions)).toBe(true);
      expect(analysis.recommendedActions.length).toBeGreaterThan(0);
    });

    it('should identify critical devices', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      expect(analysis.criticalDevices).toBeDefined();
      expect(Array.isArray(analysis.criticalDevices)).toBe(true);
      analysis.criticalDevices.forEach((device) => {
        expect(device.severity).toBe('Critical');
      });
    });

    it('should filter by failure type', async () => {
      const analysis = await report.analyzeEnrollmentFailures({
        failureType: [EnrollmentFailureType.CertificateFailure],
      });

      expect(analysis.totalFailures).toBe(analysis.certificateIssues);
    });

    it('should filter by severity', async () => {
      const analysis = await report.analyzeEnrollmentFailures({
        severity: ['Critical'],
      });

      analysis.criticalDevices.forEach((device) => {
        expect(device.severity).toBe('Critical');
      });
    });
  });

  describe('getDevicesWithFailedAssignment', () => {
    it('should identify devices with assignment failures', async () => {
      const assignmentFailures = await report.getDevicesWithFailedAssignment();

      expect(assignmentFailures.length).toBeGreaterThan(0);
      assignmentFailures.forEach((device) => {
        expect(device).toHaveProperty('deviceId');
        expect(device).toHaveProperty('deviceName');
        expect(device).toHaveProperty('assignmentType');
        expect(device).toHaveProperty('failureReason');
      });
    });

    it('should include troubleshooting steps for assignments', async () => {
      const assignmentFailures = await report.getDevicesWithFailedAssignment();

      assignmentFailures.forEach((device) => {
        expect(device.troubleshootingSteps).toBeDefined();
        expect(Array.isArray(device.troubleshootingSteps)).toBe(true);
        expect(device.troubleshootingSteps.length).toBeGreaterThan(0);
      });
    });

    it('should track retry count', async () => {
      const assignmentFailures = await report.getDevicesWithFailedAssignment();

      assignmentFailures.forEach((device) => {
        expect(device.retryCount).toBeGreaterThanOrEqual(0);
      });
    });
  });

  // ==========================================================================
  // Export Tests
  // ==========================================================================

  describe('Export Methods', () => {
    it('should export category report to JSON', async () => {
      const path = await report.exportCategoryReportToJSON('Corporate', './test-reports');

      expect(path).toBeDefined();
      expect(typeof path).toBe('string');
    });

    it('should export category report to CSV', async () => {
      const path = await report.exportCategoryReportToCSV('Corporate', './test-reports');

      expect(path).toBeDefined();
      expect(typeof path).toBe('string');
    });

    it('should export category report to HTML', async () => {
      const path = await report.exportCategoryReportToHTML('Corporate', './test-reports');

      expect(path).toBeDefined();
      expect(typeof path).toBe('string');
    });

    it('should export enrollment failure report to JSON', async () => {
      const path = await report.exportEnrollmentFailureReportToJSON(
        undefined,
        './test-reports'
      );

      expect(path).toBeDefined();
      expect(typeof path).toBe('string');
    });

    it('should export enrollment failure report to CSV', async () => {
      const path = await report.exportEnrollmentFailureReportToCSV(undefined, './test-reports');

      expect(path).toBeDefined();
      expect(typeof path).toBe('string');
    });

    it('should export enrollment failure report to HTML', async () => {
      const path = await report.exportEnrollmentFailureReportToHTML(undefined, './test-reports');

      expect(path).toBeDefined();
      expect(typeof path).toBe('string');
    });
  });

  // ==========================================================================
  // Integration Tests
  // ==========================================================================

  describe('Integration Tests', () => {
    it('should handle empty device list gracefully', async () => {
      const emptyClient = createMockGraphClient({
        '/deviceManagement/managedDevices': { value: [] },
        '/deviceManagement/deviceCategories': { value: [] },
      });

      const emptyReport = new DeviceCategoryEnrollmentReports(emptyClient, mockConfig);
      const unmanagedDevices = await emptyReport.getUnmanagedDevicesEnrollmentFailed();

      expect(unmanagedDevices).toHaveLength(0);
    });

    it('should handle devices without categories', async () => {
      const deviceWithoutCategory = {
        ...mockHealthyDevice,
        deviceCategoryDisplayName: null,
      };

      const clientWithUncategorized = createMockGraphClient({
        '/deviceManagement/managedDevices': { value: [deviceWithoutCategory] },
        '/deviceManagement/deviceCategories': { value: mockCategories },
      });

      const testReport = new DeviceCategoryEnrollmentReports(
        clientWithUncategorized,
        mockConfig
      );
      const summaries = await testReport.getCategorySummary();

      expect(summaries).toBeDefined();
    });

    it('should handle missing certificate data', async () => {
      const deviceWithoutCert = {
        ...mockHealthyDevice,
        certificateExpirationDateTime: null,
      };

      const clientWithoutCert = createMockGraphClient({
        '/deviceManagement/managedDevices': { value: [deviceWithoutCert] },
        '/deviceManagement/deviceCategories': { value: mockCategories },
      });

      const testReport = new DeviceCategoryEnrollmentReports(clientWithoutCert, mockConfig);
      const unmanagedDevices = await testReport.getUnmanagedDevicesEnrollmentFailed();

      expect(unmanagedDevices).toBeDefined();
    });

    it('should handle API errors gracefully', async () => {
      const errorClient = {
        api: jest.fn(() => ({
          select: jest.fn().mockReturnThis(),
          filter: jest.fn().mockReturnThis(),
          top: jest.fn().mockReturnThis(),
          get: jest.fn(async () => {
            throw new Error('API Error');
          }),
        })),
      } as unknown as Client;

      const errorReport = new DeviceCategoryEnrollmentReports(errorClient, mockConfig);

      await expect(errorReport.getAllDeviceCategories()).rejects.toThrow();
    });
  });

  // ==========================================================================
  // Data Validation Tests
  // ==========================================================================

  describe('Data Validation', () => {
    it('should validate device category structure', async () => {
      const categories = await report.getAllDeviceCategories();

      categories.forEach((category) => {
        expect(category.id).toBeDefined();
        expect(category.displayName).toBeDefined();
        expect(typeof category.id).toBe('string');
        expect(typeof category.displayName).toBe('string');
      });
    });

    it('should validate unmanaged device structure', async () => {
      const unmanagedDevices = await report.getUnmanagedDevicesEnrollmentFailed();

      unmanagedDevices.forEach((device) => {
        expect(device.userPrincipalName).toBeDefined();
        expect(device.platform).toBeDefined();
        expect(device.failureType).toBeDefined();
        expect(device.certificateStatus).toBeDefined();
        expect(device.severity).toBeDefined();
        expect(device.troubleshootingSteps).toBeDefined();
        expect(device.recommendation).toBeDefined();
      });
    });

    it('should validate failure analysis structure', async () => {
      const analysis = await report.analyzeEnrollmentFailures();

      expect(typeof analysis.totalFailures).toBe('number');
      expect(typeof analysis.affectedUsers).toBe('number');
      expect(typeof analysis.certificateIssues).toBe('number');
      expect(Array.isArray(analysis.recommendedActions)).toBe(true);
      expect(Array.isArray(analysis.criticalDevices)).toBe(true);
    });
  });
});
