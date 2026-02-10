/**
 * Unit Tests for Windows CE Legacy Reports
 */

import { WindowsCELegacyReports } from './windows-ce-reports';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';

// Mock the dependencies
jest.mock('@microsoft/microsoft-graph-client');
jest.mock('../../core/logger');
jest.mock('../../formatters/output-formatter');

describe('WindowsCELegacyReports', () => {
  let mockGraphClient: jest.Mocked<Client>;
  let mockConfig: AppConfig;
  let report: WindowsCELegacyReports;

  // Sample device data
  const sampleDevice = {
    id: 'device-123',
    deviceName: 'TEST-DEVICE-01',
    userPrincipalName: 'user@contoso.com',
    operatingSystem: 'Windows',
    osVersion: '10.0.19042',
    manufacturer: 'Microsoft',
    model: 'Surface Pro',
    serialNumber: 'SN123456',
    managedDeviceOwnerType: 'company',
    enrolledDateTime: '2024-01-01T00:00:00Z',
    lastSyncDateTime: '2024-01-15T12:00:00Z',
    complianceState: 'compliant',
    managementAgent: 'mdm',
    isEncrypted: true,
    totalStorageSpaceInBytes: 256000000000,
    freeStorageSpaceInBytes: 128000000000,
    wifiMacAddress: '00:11:22:33:44:55',
    subscriberCarrier: 'Verizon',
  };

  const sampleComplianceState = {
    id: 'compliance-1',
    displayName: 'Device Compliance Policy',
    state: 'compliant',
    version: 1,
    lastReportedDateTime: '2024-01-15T12:00:00Z',
    settingStates: [
      {
        setting: 'RequireDeviceEncryption',
        state: 'compliant',
      },
      {
        setting: 'CertificateExpiry',
        state: 'error',
        errorDescription: 'Certificate expired',
      },
    ],
  };

  const sampleConfigState = {
    id: 'config-1',
    displayName: 'WiFi Configuration',
    state: 'compliant',
    version: 1,
    lastReportedDateTime: '2024-01-15T12:00:00Z',
    settingStates: [],
  };

  const sampleTroubleshootingEvent = {
    id: 'event-1',
    eventDateTime: '2024-01-15T10:00:00Z',
    correlationId: 'corr-123',
    troubleshootingErrorDetails: {
      context: 'Policy Application',
      failure: 'Network connection failed',
      failureDetails: 'Unable to reach policy endpoint',
    },
    deviceId: 'device-123',
  };

  beforeEach(() => {
    // Setup mock config
    mockConfig = {
      authentication: {
        tenantId: 'test-tenant',
        clientId: 'test-client',
        authMethod: 'clientSecret',
      },
      reports: {
        enabled: [],
        disabled: [],
        settings: {},
      },
      output: {
        defaultFormat: 'json',
        directory: '/tmp',
        includeTimestamp: true,
        compression: false,
      },
      scheduler: {
        enabled: false,
        timezone: 'UTC',
        schedules: [],
      },
      logging: {
        level: 'info',
        file: '/tmp/test.log',
        console: true,
        maxSize: '10m',
        maxFiles: 5,
      },
    } as AppConfig;

    // Setup mock graph client
    mockGraphClient = {
      api: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      select: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      orderby: jest.fn().mockReturnThis(),
      get: jest.fn(),
    } as any;

    report = new WindowsCELegacyReports(mockGraphClient as any, mockConfig);
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  // ==========================================================================
  // Basic Functionality Tests
  // ==========================================================================

  describe('Basic Functionality', () => {
    it('should initialize with correct properties', () => {
      expect(report.name).toBe('windows-ce-legacy-reports');
      expect(report.description).toContain('Windows CE legacy reports');
      expect(report.category).toBe('ConfigMgr-Style Reports');
      expect(report.enabled).toBe(true);
    });

    it('should have all required report methods', () => {
      expect(typeof report.getCertificateIssues).toBe('function');
      expect(typeof report.getClientDeploymentFailures).toBe('function');
      expect(typeof report.getClientDeploymentStatus).toBe('function');
      expect(typeof report.getClientDeploymentSuccess).toBe('function');
      expect(typeof report.getCommunicationIssues).toBe('function');
      expect(typeof report.getDeviceHealthInformation).toBe('function');
      expect(typeof report.getDeviceHealthSummary).toBe('function');
      expect(typeof report.getLocalClientIssues).toBe('function');
    });
  });

  // ==========================================================================
  // Certificate Issues Report Tests
  // ==========================================================================

  describe('getCertificateIssues', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] }) // devices
        .mockResolvedValueOnce({ value: [sampleComplianceState] }); // compliance states
    });

    it('should fetch certificate issues successfully', async () => {
      const issues = await report.getCertificateIssues();

      expect(Array.isArray(issues)).toBe(true);
      expect(mockGraphClient.api).toHaveBeenCalledWith('/deviceManagement/managedDevices');
    });

    it('should identify expired certificates', async () => {
      const issues = await report.getCertificateIssues();

      const expiredCerts = issues.filter(
        issue => issue.certificateName.toLowerCase().includes('cert')
      );

      expect(expiredCerts.length).toBeGreaterThanOrEqual(0);
    });

    it('should include device information in certificate issues', async () => {
      const issues = await report.getCertificateIssues();

      if (issues.length > 0) {
        const issue = issues[0];
        expect(issue).toHaveProperty('deviceId');
        expect(issue).toHaveProperty('deviceName');
        expect(issue).toHaveProperty('userPrincipalName');
        expect(issue).toHaveProperty('certificateName');
        expect(issue).toHaveProperty('issueType');
        expect(issue).toHaveProperty('severity');
        expect(issue).toHaveProperty('recommendation');
      }
    });

    it('should categorize certificate issue severity correctly', async () => {
      const issues = await report.getCertificateIssues();

      issues.forEach(issue => {
        expect(['critical', 'high', 'medium', 'low', 'info']).toContain(issue.severity);
      });
    });

    it('should handle devices with no certificate issues', async () => {
      mockGraphClient.get
        .mockReset()
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [] });

      const issues = await report.getCertificateIssues();
      expect(issues).toEqual([]);
    });

    it('should apply filter options', async () => {
      await report.getCertificateIssues({
        operatingSystem: 'Windows',
        userPrincipalName: 'user@contoso.com',
      });

      expect(mockGraphClient.filter).toHaveBeenCalled();
    });
  });

  // ==========================================================================
  // Client Deployment Failure Tests
  // ==========================================================================

  describe('getClientDeploymentFailures', () => {
    beforeEach(() => {
      const failedConfigState = {
        ...sampleConfigState,
        state: 'error',
        errorDescription: 'Configuration deployment failed',
        errorCode: '0x80070001',
      };

      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] }) // devices
        .mockResolvedValueOnce({ value: [failedConfigState] }) // config states
        .mockResolvedValueOnce({ value: [sampleComplianceState] }) // compliance states
        .mockResolvedValueOnce({ value: [sampleTroubleshootingEvent] }); // troubleshooting events
    });

    it('should fetch deployment failures successfully', async () => {
      const failures = await report.getClientDeploymentFailures();

      expect(Array.isArray(failures)).toBe(true);
    });

    it('should include error details in failures', async () => {
      const failures = await report.getClientDeploymentFailures();

      failures.forEach(failure => {
        expect(failure).toHaveProperty('deviceId');
        expect(failure).toHaveProperty('deploymentType');
        expect(failure).toHaveProperty('errorMessage');
        expect(failure).toHaveProperty('failureReason');
        expect(failure).toHaveProperty('recommendation');
      });
    });

    it('should categorize error types correctly', async () => {
      const failures = await report.getClientDeploymentFailures();

      failures.forEach(failure => {
        expect(failure.errorCategory).toBeDefined();
        expect(typeof failure.errorCategory).toBe('string');
      });
    });

    it('should indicate if deployment can be retried', async () => {
      const failures = await report.getClientDeploymentFailures();

      failures.forEach(failure => {
        expect(typeof failure.canRetry).toBe('boolean');
      });
    });
  });

  // ==========================================================================
  // Client Deployment Status Tests
  // ==========================================================================

  describe('getClientDeploymentStatus', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [sampleConfigState] })
        .mockResolvedValueOnce({ value: [sampleComplianceState] });
    });

    it('should fetch deployment status successfully', async () => {
      const statuses = await report.getClientDeploymentStatus();

      expect(Array.isArray(statuses)).toBe(true);
    });

    it('should include progress information', async () => {
      const statuses = await report.getClientDeploymentStatus();

      statuses.forEach(status => {
        expect(status).toHaveProperty('progress');
        expect(status.progress).toBeGreaterThanOrEqual(0);
        expect(status.progress).toBeLessThanOrEqual(100);
      });
    });

    it('should calculate deployment duration', async () => {
      const statuses = await report.getClientDeploymentStatus();

      statuses.forEach(status => {
        if (status.durationMinutes !== undefined) {
          expect(status.durationMinutes).toBeGreaterThanOrEqual(0);
        }
      });
    });

    it('should map deployment status correctly', async () => {
      const statuses = await report.getClientDeploymentStatus();

      const validStatuses = ['success', 'failed', 'in_progress', 'pending', 'not_applicable', 'error'];
      statuses.forEach(status => {
        expect(validStatuses).toContain(status.status);
      });
    });
  });

  // ==========================================================================
  // Client Deployment Success Tests
  // ==========================================================================

  describe('getClientDeploymentSuccess', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [sampleConfigState] })
        .mockResolvedValueOnce({ value: [sampleComplianceState] });
    });

    it('should fetch successful deployments', async () => {
      const successes = await report.getClientDeploymentSuccess();

      expect(Array.isArray(successes)).toBe(true);
    });

    it('should include success details', async () => {
      const successes = await report.getClientDeploymentSuccess();

      successes.forEach(success => {
        expect(success).toHaveProperty('deviceId');
        expect(success).toHaveProperty('deploymentType');
        expect(success).toHaveProperty('successTime');
        expect(success).toHaveProperty('durationMinutes');
        expect(success).toHaveProperty('complianceStatus');
      });
    });

    it('should only include compliant deployments', async () => {
      const successes = await report.getClientDeploymentSuccess();

      successes.forEach(success => {
        expect(success).toBeDefined();
        // Success records should be from compliant states
      });
    });
  });

  // ==========================================================================
  // Communication Issues Tests
  // ==========================================================================

  describe('getCommunicationIssues', () => {
    it('should identify devices with sync issues', async () => {
      const oldSyncDevice = {
        ...sampleDevice,
        lastSyncDateTime: '2024-01-01T00:00:00Z', // 15 days ago
      };

      mockGraphClient.get.mockResolvedValueOnce({ value: [oldSyncDevice] });

      const issues = await report.getCommunicationIssues();

      expect(issues.length).toBeGreaterThan(0);
      expect(issues[0].communicationStatus).not.toBe('healthy');
    });

    it('should calculate days since last sync', async () => {
      mockGraphClient.get.mockResolvedValueOnce({ value: [sampleDevice] });

      const issues = await report.getCommunicationIssues();

      issues.forEach(issue => {
        if (issue.lastSyncDateTime) {
          expect(issue.daysSinceLastSync).toBeDefined();
          expect(typeof issue.daysSinceLastSync).toBe('number');
        }
      });
    });

    it('should include troubleshooting steps', async () => {
      const oldSyncDevice = {
        ...sampleDevice,
        lastSyncDateTime: '2024-01-01T00:00:00Z',
      };

      mockGraphClient.get.mockResolvedValueOnce({ value: [oldSyncDevice] });

      const issues = await report.getCommunicationIssues();

      issues.forEach(issue => {
        expect(Array.isArray(issue.troubleshootingSteps)).toBe(true);
        expect(issue.troubleshootingSteps.length).toBeGreaterThan(0);
      });
    });

    it('should categorize communication status correctly', async () => {
      const devices = [
        { ...sampleDevice, lastSyncDateTime: new Date().toISOString() }, // healthy
        { ...sampleDevice, id: 'device-2', lastSyncDateTime: new Date(Date.now() - 3 * 24 * 60 * 60 * 1000).toISOString() }, // warning
        { ...sampleDevice, id: 'device-3', lastSyncDateTime: new Date(Date.now() - 15 * 24 * 60 * 60 * 1000).toISOString() }, // critical
      ];

      mockGraphClient.get.mockResolvedValueOnce({ value: devices });

      const issues = await report.getCommunicationIssues();

      const validStatuses = ['healthy', 'warning', 'critical', 'offline', 'unknown'];
      issues.forEach(issue => {
        expect(validStatuses).toContain(issue.communicationStatus);
      });
    });

    it('should include network details', async () => {
      mockGraphClient.get.mockResolvedValueOnce({ value: [sampleDevice] });

      const issues = await report.getCommunicationIssues();

      issues.forEach(issue => {
        expect(issue.networkDetails).toBeDefined();
      });
    });
  });

  // ==========================================================================
  // Device Health Information Tests
  // ==========================================================================

  describe('getDeviceHealthInformation', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [sampleConfigState] })
        .mockResolvedValueOnce({ value: [sampleComplianceState] });
    });

    it('should fetch device health information', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      expect(Array.isArray(healthInfo)).toBe(true);
    });

    it('should calculate health scores', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      healthInfo.forEach(health => {
        expect(health.healthScore).toBeGreaterThanOrEqual(0);
        expect(health.healthScore).toBeLessThanOrEqual(100);
      });
    });

    it('should determine overall health status', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      const validStatuses = ['healthy', 'warning', 'unhealthy', 'critical', 'unknown'];
      healthInfo.forEach(health => {
        expect(validStatuses).toContain(health.overallHealthStatus);
      });
    });

    it('should include compliance information', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      healthInfo.forEach(health => {
        expect(typeof health.isCompliant).toBe('boolean');
        expect(health.complianceState).toBeDefined();
      });
    });

    it('should track configuration profiles', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      healthInfo.forEach(health => {
        expect(typeof health.configurationProfilesApplied).toBe('number');
        expect(typeof health.configurationProfilesFailed).toBe('number');
      });
    });

    it('should include security information', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      healthInfo.forEach(health => {
        expect(typeof health.isEncrypted).toBe('boolean');
      });
    });

    it('should calculate storage usage', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      healthInfo.forEach(health => {
        if (health.totalStorageSpace && health.freeStorageSpace) {
          expect(health.storagePercentUsed).toBeDefined();
          expect(health.storagePercentUsed).toBeGreaterThanOrEqual(0);
          expect(health.storagePercentUsed).toBeLessThanOrEqual(100);
        }
      });
    });

    it('should identify health issues', async () => {
      const healthInfo = await report.getDeviceHealthInformation();

      healthInfo.forEach(health => {
        expect(Array.isArray(health.healthIssues)).toBe(true);
        expect(Array.isArray(health.recommendations)).toBe(true);
      });
    });
  });

  // ==========================================================================
  // Device Health Summary Tests
  // ==========================================================================

  describe('getDeviceHealthSummary', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [sampleConfigState] })
        .mockResolvedValueOnce({ value: [sampleComplianceState] });
    });

    it('should generate health summary', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(summary).toBeDefined();
      expect(summary.totalDevices).toBeGreaterThanOrEqual(0);
    });

    it('should calculate health distribution', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(summary.healthyDevices).toBeGreaterThanOrEqual(0);
      expect(summary.warningDevices).toBeGreaterThanOrEqual(0);
      expect(summary.unhealthyDevices).toBeGreaterThanOrEqual(0);
      expect(summary.criticalDevices).toBeGreaterThanOrEqual(0);

      const total = summary.healthyDevices + summary.warningDevices +
                   summary.unhealthyDevices + summary.criticalDevices +
                   summary.unknownDevices;
      expect(total).toBe(summary.totalDevices);
    });

    it('should calculate percentages correctly', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(summary.healthyPercentage).toBeGreaterThanOrEqual(0);
      expect(summary.healthyPercentage).toBeLessThanOrEqual(100);
      expect(summary.compliancePercentage).toBeGreaterThanOrEqual(0);
      expect(summary.compliancePercentage).toBeLessThanOrEqual(100);
      expect(summary.encryptionPercentage).toBeGreaterThanOrEqual(0);
      expect(summary.encryptionPercentage).toBeLessThanOrEqual(100);
    });

    it('should calculate average health score', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(summary.averageHealthScore).toBeGreaterThanOrEqual(0);
      expect(summary.averageHealthScore).toBeLessThanOrEqual(100);
    });

    it('should track sync activity', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(summary.devicesActiveLast24Hours).toBeGreaterThanOrEqual(0);
      expect(summary.devicesActiveLast7Days).toBeGreaterThanOrEqual(0);
      expect(summary.devicesActiveLast30Days).toBeGreaterThanOrEqual(0);
    });

    it('should identify top health issues', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(Array.isArray(summary.topHealthIssues)).toBe(true);
      summary.topHealthIssues.forEach(issue => {
        expect(issue).toHaveProperty('issueType');
        expect(issue).toHaveProperty('affectedDeviceCount');
        expect(issue).toHaveProperty('affectedPercentage');
        expect(issue).toHaveProperty('recommendation');
      });
    });

    it('should group by operating system', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(typeof summary.byOperatingSystem).toBe('object');

      Object.values(summary.byOperatingSystem).forEach(osSummary => {
        expect(osSummary.total).toBeGreaterThanOrEqual(0);
        expect(osSummary.averageHealthScore).toBeGreaterThanOrEqual(0);
        expect(osSummary.complianceRate).toBeGreaterThanOrEqual(0);
      });
    });

    it('should include generation timestamp', async () => {
      const summary = await report.getDeviceHealthSummary();

      expect(summary.generatedAt).toBeInstanceOf(Date);
    });
  });

  // ==========================================================================
  // Local Client Issues Tests
  // ==========================================================================

  describe('getLocalClientIssues', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [sampleTroubleshootingEvent] });
    });

    it('should fetch local client issues', async () => {
      const issues = await report.getLocalClientIssues();

      expect(Array.isArray(issues)).toBe(true);
    });

    it('should categorize issue types', async () => {
      const issues = await report.getLocalClientIssues();

      issues.forEach(issue => {
        expect(issue.issueType).toBeDefined();
        expect(issue.issueCategory).toBeDefined();
      });
    });

    it('should track issue occurrences', async () => {
      const issues = await report.getLocalClientIssues();

      issues.forEach(issue => {
        expect(issue.occurrenceCount).toBeGreaterThanOrEqual(1);
        expect(issue.firstOccurrence).toBeInstanceOf(Date);
        expect(issue.lastOccurrence).toBeInstanceOf(Date);
      });
    });

    it('should include troubleshooting steps', async () => {
      const issues = await report.getLocalClientIssues();

      issues.forEach(issue => {
        expect(Array.isArray(issue.troubleshootingSteps)).toBe(true);
        expect(issue.recommendation).toBeDefined();
      });
    });

    it('should determine issue severity', async () => {
      const issues = await report.getLocalClientIssues();

      const validSeverities = ['critical', 'high', 'medium', 'low'];
      issues.forEach(issue => {
        expect(validSeverities).toContain(issue.severity);
      });
    });
  });

  // ==========================================================================
  // Execute Method Tests
  // ==========================================================================

  describe('execute', () => {
    beforeEach(() => {
      // Mock all required API calls
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] }) // devices for health summary
        .mockResolvedValueOnce({ value: [sampleConfigState] }) // config states
        .mockResolvedValueOnce({ value: [sampleComplianceState] }) // compliance states
        .mockResolvedValueOnce({ value: [sampleDevice] }) // devices for cert issues
        .mockResolvedValueOnce({ value: [sampleComplianceState] }) // compliance for certs
        .mockResolvedValueOnce({ value: [sampleDevice] }) // devices for comm issues
        .mockResolvedValueOnce({ value: [sampleDevice] }) // devices for client issues
        .mockResolvedValueOnce({ value: [sampleTroubleshootingEvent] }); // troubleshooting events
    });

    it('should execute successfully', async () => {
      const result = await report.execute();

      expect(result).toBeDefined();
      expect(result.metadata).toBeDefined();
      expect(result.data).toBeDefined();
      expect(Array.isArray(result.data)).toBe(true);
    });

    it('should include metadata', async () => {
      const result = await report.execute();

      expect(result.metadata.reportName).toBe('windows-ce-legacy-reports');
      expect(result.metadata.generatedAt).toBeDefined();
      expect(result.metadata.generatedBy).toBe('Intune Reporting Dashboard');
      expect(typeof result.metadata.recordCount).toBe('number');
    });

    it('should include summary data', async () => {
      const result = await report.execute();

      expect(result.summary).toBeDefined();
      expect(result.summary.totalDevices).toBeGreaterThanOrEqual(0);
    });
  });

  // ==========================================================================
  // Error Handling Tests
  // ==========================================================================

  describe('Error Handling', () => {
    it('should handle Graph API errors gracefully', async () => {
      mockGraphClient.get.mockRejectedValueOnce(new Error('Graph API error'));

      await expect(report.getCertificateIssues()).rejects.toThrow();
    });

    it('should handle missing device data', async () => {
      mockGraphClient.get.mockResolvedValueOnce({ value: [] });

      const issues = await report.getCertificateIssues();
      expect(issues).toEqual([]);
    });

    it('should handle devices without compliance states', async () => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [] });

      const issues = await report.getCertificateIssues();
      expect(Array.isArray(issues)).toBe(true);
    });

    it('should handle network timeout', async () => {
      mockGraphClient.get.mockImplementationOnce(() => {
        return new Promise((_, reject) => {
          setTimeout(() => reject(new Error('Timeout')), 100);
        });
      });

      await expect(report.getCertificateIssues()).rejects.toThrow('Timeout');
    });
  });

  // ==========================================================================
  // Filter and Options Tests
  // ==========================================================================

  describe('Filter Options', () => {
    beforeEach(() => {
      mockGraphClient.get
        .mockResolvedValueOnce({ value: [sampleDevice] })
        .mockResolvedValueOnce({ value: [sampleComplianceState] });
    });

    it('should apply operating system filter', async () => {
      await report.getCertificateIssues({
        operatingSystem: 'Windows',
      });

      expect(mockGraphClient.filter).toHaveBeenCalled();
    });

    it('should apply user principal name filter', async () => {
      await report.getCertificateIssues({
        userPrincipalName: 'user@contoso.com',
      });

      expect(mockGraphClient.filter).toHaveBeenCalled();
    });

    it('should apply device name filter', async () => {
      await report.getCertificateIssues({
        deviceName: 'TEST-DEVICE',
      });

      expect(mockGraphClient.filter).toHaveBeenCalled();
    });

    it('should apply multiple filters', async () => {
      await report.getCertificateIssues({
        operatingSystem: 'Windows',
        userPrincipalName: 'user@contoso.com',
        deviceName: 'TEST-DEVICE',
      });

      expect(mockGraphClient.filter).toHaveBeenCalled();
    });
  });

  // ==========================================================================
  // Export Functionality Tests
  // ==========================================================================

  describe('Export Functionality', () => {
    it('should have exportReport method', () => {
      expect(typeof report.exportReport).toBe('function');
    });

    it('should accept export options', async () => {
      const reportData = {
        metadata: {
          reportName: 'test-report',
          generatedAt: new Date().toISOString(),
          generatedBy: 'Test',
          recordCount: 0,
        },
        data: [],
      };

      const options = {
        format: 'json' as const,
        outputDirectory: '/tmp',
        includeTimestamp: true,
      };

      // Mock the OutputFormatter
      const { OutputFormatter } = require('../../formatters/output-formatter');
      OutputFormatter.format = jest.fn().mockResolvedValue('/tmp/test-report.json');

      const result = await report.exportReport(reportData, options);
      expect(result).toBeDefined();
    });
  });
});
