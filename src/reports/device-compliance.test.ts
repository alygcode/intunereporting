/**
 * Unit Tests for Device Compliance Reporter
 *
 * This file contains unit tests for the DeviceComplianceReporter class.
 * Note: These tests use mocks and don't make real API calls.
 */

import { DeviceComplianceReporter, AuthConfig, DeviceCompliance } from './device-compliance';

// Mock the Microsoft Graph client
jest.mock('@microsoft/microsoft-graph-client');
jest.mock('@azure/identity');

describe('DeviceComplianceReporter', () => {
  const validConfig: AuthConfig = {
    tenantId: 'test-tenant-id',
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret'
  };

  describe('Constructor and Validation', () => {
    it('should create instance with valid config', () => {
      expect(() => new DeviceComplianceReporter(validConfig)).not.toThrow();
    });

    it('should throw error with empty tenant ID', () => {
      const invalidConfig = { ...validConfig, tenantId: '' };
      expect(() => new DeviceComplianceReporter(invalidConfig)).toThrow('Tenant ID is required');
    });

    it('should throw error with empty client ID', () => {
      const invalidConfig = { ...validConfig, clientId: '' };
      expect(() => new DeviceComplianceReporter(invalidConfig)).toThrow('Client ID is required');
    });

    it('should throw error with empty client secret', () => {
      const invalidConfig = { ...validConfig, clientSecret: '' };
      expect(() => new DeviceComplianceReporter(invalidConfig)).toThrow('Client secret is required');
    });

    it('should throw error with whitespace-only tenant ID', () => {
      const invalidConfig = { ...validConfig, tenantId: '   ' };
      expect(() => new DeviceComplianceReporter(invalidConfig)).toThrow('Tenant ID is required');
    });
  });

  describe('Summary Generation', () => {
    let reporter: DeviceComplianceReporter;

    beforeEach(() => {
      reporter = new DeviceComplianceReporter(validConfig);
    });

    it('should generate correct summary for empty device list', () => {
      const devices: DeviceCompliance[] = [];
      const summary = (reporter as any).generateSummary(devices);

      expect(summary.totalDevices).toBe(0);
      expect(summary.compliantDevices).toBe(0);
      expect(summary.nonCompliantDevices).toBe(0);
      expect(summary.compliancePercentage).toBe(0);
    });

    it('should generate correct summary with all compliant devices', () => {
      const devices: DeviceCompliance[] = [
        {
          id: '1',
          deviceName: 'Device 1',
          userPrincipalName: 'user1@test.com',
          operatingSystem: 'Windows',
          osVersion: '10',
          complianceState: 'compliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        },
        {
          id: '2',
          deviceName: 'Device 2',
          userPrincipalName: 'user2@test.com',
          operatingSystem: 'Windows',
          osVersion: '11',
          complianceState: 'compliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        }
      ];

      const summary = (reporter as any).generateSummary(devices);

      expect(summary.totalDevices).toBe(2);
      expect(summary.compliantDevices).toBe(2);
      expect(summary.nonCompliantDevices).toBe(0);
      expect(summary.compliancePercentage).toBe(100);
    });

    it('should generate correct summary with mixed compliance states', () => {
      const devices: DeviceCompliance[] = [
        {
          id: '1',
          deviceName: 'Device 1',
          userPrincipalName: 'user1@test.com',
          operatingSystem: 'Windows',
          osVersion: '10',
          complianceState: 'compliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        },
        {
          id: '2',
          deviceName: 'Device 2',
          userPrincipalName: 'user2@test.com',
          operatingSystem: 'iOS',
          osVersion: '16',
          complianceState: 'noncompliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        },
        {
          id: '3',
          deviceName: 'Device 3',
          userPrincipalName: 'user3@test.com',
          operatingSystem: 'Android',
          osVersion: '13',
          complianceState: 'ingraceperiod',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        },
        {
          id: '4',
          deviceName: 'Device 4',
          userPrincipalName: 'user4@test.com',
          operatingSystem: 'Windows',
          osVersion: '11',
          complianceState: 'compliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        }
      ];

      const summary = (reporter as any).generateSummary(devices);

      expect(summary.totalDevices).toBe(4);
      expect(summary.compliantDevices).toBe(2);
      expect(summary.nonCompliantDevices).toBe(1);
      expect(summary.inGracePeriodDevices).toBe(1);
      expect(summary.compliancePercentage).toBe(50);
    });

    it('should group devices by operating system correctly', () => {
      const devices: DeviceCompliance[] = [
        {
          id: '1',
          deviceName: 'Device 1',
          userPrincipalName: 'user1@test.com',
          operatingSystem: 'Windows',
          osVersion: '10',
          complianceState: 'compliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        },
        {
          id: '2',
          deviceName: 'Device 2',
          userPrincipalName: 'user2@test.com',
          operatingSystem: 'Windows',
          osVersion: '11',
          complianceState: 'noncompliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        },
        {
          id: '3',
          deviceName: 'Device 3',
          userPrincipalName: 'user3@test.com',
          operatingSystem: 'iOS',
          osVersion: '16',
          complianceState: 'compliant',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        }
      ];

      const summary = (reporter as any).generateSummary(devices);

      expect(summary.byOperatingSystem['Windows']).toBeDefined();
      expect(summary.byOperatingSystem['Windows'].total).toBe(2);
      expect(summary.byOperatingSystem['Windows'].compliant).toBe(1);
      expect(summary.byOperatingSystem['Windows'].nonCompliant).toBe(1);
      expect(summary.byOperatingSystem['Windows'].compliancePercentage).toBe(50);

      expect(summary.byOperatingSystem['iOS']).toBeDefined();
      expect(summary.byOperatingSystem['iOS'].total).toBe(1);
      expect(summary.byOperatingSystem['iOS'].compliant).toBe(1);
      expect(summary.byOperatingSystem['iOS'].compliancePercentage).toBe(100);
    });

    it('should handle unknown compliance states', () => {
      const devices: DeviceCompliance[] = [
        {
          id: '1',
          deviceName: 'Device 1',
          userPrincipalName: 'user1@test.com',
          operatingSystem: 'Windows',
          osVersion: '10',
          complianceState: 'unknown',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        }
      ];

      const summary = (reporter as any).generateSummary(devices);

      expect(summary.unknownDevices).toBe(1);
      expect(summary.compliantDevices).toBe(0);
    });

    it('should handle error compliance states', () => {
      const devices: DeviceCompliance[] = [
        {
          id: '1',
          deviceName: 'Device 1',
          userPrincipalName: 'user1@test.com',
          operatingSystem: 'Windows',
          osVersion: '10',
          complianceState: 'error',
          lastReportedDateTime: new Date().toISOString(),
          isManaged: true
        }
      ];

      const summary = (reporter as any).generateSummary(devices);

      expect(summary.errorDevices).toBe(1);
    });
  });

  describe('CSV Escaping', () => {
    let reporter: DeviceComplianceReporter;

    beforeEach(() => {
      reporter = new DeviceComplianceReporter(validConfig);
    });

    it('should escape values with commas', () => {
      const result = (reporter as any).escapeCsvValue('Device, Name');
      expect(result).toBe('"Device, Name"');
    });

    it('should escape values with quotes', () => {
      const result = (reporter as any).escapeCsvValue('Device "Special"');
      expect(result).toBe('"Device ""Special"""');
    });

    it('should escape values with newlines', () => {
      const result = (reporter as any).escapeCsvValue('Device\nName');
      expect(result).toBe('"Device\nName"');
    });

    it('should not escape simple values', () => {
      const result = (reporter as any).escapeCsvValue('SimpleDevice');
      expect(result).toBe('SimpleDevice');
    });
  });

  describe('Error Message Extraction', () => {
    let reporter: DeviceComplianceReporter;

    beforeEach(() => {
      reporter = new DeviceComplianceReporter(validConfig);
    });

    it('should extract message from Error object', () => {
      const error = new Error('Test error');
      const message = (reporter as any).getErrorMessage(error);
      expect(message).toBe('Test error');
    });

    it('should handle string errors', () => {
      const message = (reporter as any).getErrorMessage('String error');
      expect(message).toBe('String error');
    });

    it('should handle objects with message property', () => {
      const error = { message: 'Object error' };
      const message = (reporter as any).getErrorMessage(error);
      expect(message).toBe('Object error');
    });

    it('should handle unknown error types', () => {
      const message = (reporter as any).getErrorMessage(null);
      expect(message).toBe('Unknown error occurred');
    });
  });

  describe('Static Methods', () => {
    it('should print summary without errors', () => {
      const summary = {
        totalDevices: 100,
        compliantDevices: 90,
        nonCompliantDevices: 8,
        unknownDevices: 1,
        inGracePeriodDevices: 1,
        errorDevices: 0,
        compliancePercentage: 90,
        byOperatingSystem: {
          Windows: {
            total: 60,
            compliant: 55,
            nonCompliant: 5,
            compliancePercentage: 91.67
          },
          iOS: {
            total: 40,
            compliant: 35,
            nonCompliant: 3,
            compliancePercentage: 87.5
          }
        },
        generatedAt: new Date().toISOString()
      };

      // Should not throw
      expect(() => DeviceComplianceReporter.printSummary(summary)).not.toThrow();
    });
  });

  describe('Utility Functions', () => {
    let reporter: DeviceComplianceReporter;

    beforeEach(() => {
      reporter = new DeviceComplianceReporter(validConfig);
    });

    it('should sleep for specified duration', async () => {
      const start = Date.now();
      await (reporter as any).sleep(100);
      const duration = Date.now() - start;

      expect(duration).toBeGreaterThanOrEqual(90); // Allow some variance
      expect(duration).toBeLessThan(200);
    });
  });
});

describe('Helper Functions', () => {
  describe('createReporterFromEnv', () => {
    const originalEnv = process.env;

    beforeEach(() => {
      jest.resetModules();
      process.env = { ...originalEnv };
    });

    afterAll(() => {
      process.env = originalEnv;
    });

    it('should create reporter with environment variables', () => {
      process.env.AZURE_TENANT_ID = 'test-tenant';
      process.env.AZURE_CLIENT_ID = 'test-client';
      process.env.AZURE_CLIENT_SECRET = 'test-secret';

      const { createReporterFromEnv } = require('./device-compliance');
      const reporter = createReporterFromEnv();

      expect(reporter).toBeInstanceOf(DeviceComplianceReporter);
    });

    it('should throw error with missing environment variables', () => {
      delete process.env.AZURE_TENANT_ID;
      delete process.env.AZURE_CLIENT_ID;
      delete process.env.AZURE_CLIENT_SECRET;

      const { createReporterFromEnv } = require('./device-compliance');

      expect(() => createReporterFromEnv()).toThrow();
    });
  });
});
