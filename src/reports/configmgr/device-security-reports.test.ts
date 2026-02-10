/**
 * Unit tests for Device Security Reports
 * Tests for jailbreak detection, certificate renewal, and health attestation reports
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { DeviceSecurityReports } from './device-security-reports';
import { Client } from '@microsoft/microsoft-graph-client';

// Mock Microsoft Graph Client
const mockGraphClient = {
  api: vi.fn().mockReturnThis(),
  filter: vi.fn().mockReturnThis(),
  select: vi.fn().mockReturnThis(),
  top: vi.fn().mockReturnThis(),
  get: vi.fn(),
} as unknown as Client;

describe('DeviceSecurityReports', () => {
  let securityReports: DeviceSecurityReports;

  beforeEach(() => {
    securityReports = new DeviceSecurityReports(mockGraphClient);
    vi.clearAllMocks();
  });

  // ============================================================================
  // Jailbreak Detection Tests (Report 23)
  // ============================================================================

  describe('Jailbreak Detection Report', () => {
    it('should fetch jailbroken devices successfully', async () => {
      const mockJailbrokenDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'iPhone-Jailbroken',
            userPrincipalName: 'user1@contoso.com',
            emailAddress: 'user1@contoso.com',
            operatingSystem: 'iOS',
            osVersion: '16.0',
            model: 'iPhone 13',
            manufacturer: 'Apple',
            jailBroken: true,
            complianceState: 'noncompliant',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            enrolledDateTime: '2023-01-01T10:00:00Z',
            managementState: 'managed'
          },
          {
            id: 'device-2',
            deviceName: 'Android-Rooted',
            userPrincipalName: 'user2@contoso.com',
            emailAddress: 'user2@contoso.com',
            operatingSystem: 'Android',
            osVersion: '13.0',
            model: 'Galaxy S21',
            manufacturer: 'Samsung',
            jailBroken: true,
            complianceState: 'noncompliant',
            lastSyncDateTime: '2024-01-15T11:00:00Z',
            enrolledDateTime: '2023-06-01T10:00:00Z',
            managementState: 'managed'
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockJailbrokenDevices);

      const devices = await securityReports.getJailbrokenDevices();

      expect(devices).toHaveLength(2);
      expect(devices[0].deviceName).toBe('iPhone-Jailbroken');
      expect(devices[0].isJailBroken).toBe(true);
      expect(devices[0].riskLevel).toBeDefined();
      expect(devices[0].riskScore).toBeGreaterThan(0);
      expect(devices[1].deviceName).toBe('Android-Rooted');
    });

    it('should generate jailbreak report with summary', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'iPhone-1',
            userPrincipalName: 'user1@contoso.com',
            emailAddress: 'user1@contoso.com',
            operatingSystem: 'iOS',
            osVersion: '16.0',
            model: 'iPhone 13',
            manufacturer: 'Apple',
            jailBroken: true,
            complianceState: 'noncompliant',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            enrolledDateTime: '2023-01-01T10:00:00Z',
            managementState: 'managed'
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const report = await securityReports.generateJailbreakReport();

      expect(report.devices).toBeDefined();
      expect(report.summary).toBeDefined();
      expect(report.summary.jailbrokenDevices).toBe(1);
      expect(report.summary.byPlatform.iOS).toBeDefined();
      expect(report.summary.riskAssessment).toBeDefined();
      expect(report.summary.riskAssessment.recommendations).toBeInstanceOf(Array);
      expect(report.summary.complianceImpact).toBeDefined();
    });

    it('should calculate jailbreak risk levels correctly', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-critical',
            deviceName: 'Critical-Device',
            userPrincipalName: 'user@contoso.com',
            operatingSystem: 'iOS',
            jailBroken: true,
            complianceState: 'noncompliant',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            enrolledDateTime: '2021-01-01T10:00:00Z', // Old device
            managementState: 'managed'
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const devices = await securityReports.getJailbrokenDevices();

      expect(devices[0].riskScore).toBeGreaterThanOrEqual(80);
      expect(devices[0].riskLevel).toBe('Critical');
      expect(devices[0].securityNotes).toBeInstanceOf(Array);
      expect(devices[0].securityNotes.length).toBeGreaterThan(0);
    });

    it('should filter jailbroken devices by operating system', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'iPhone-1',
            operatingSystem: 'iOS',
            jailBroken: true,
            complianceState: 'compliant',
            userPrincipalName: 'user1@contoso.com',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            enrolledDateTime: '2023-01-01T10:00:00Z'
          },
          {
            id: 'device-2',
            deviceName: 'Android-1',
            operatingSystem: 'Android',
            jailBroken: true,
            complianceState: 'compliant',
            userPrincipalName: 'user2@contoso.com',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            enrolledDateTime: '2023-01-01T10:00:00Z'
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const iosDevices = await securityReports.getJailbrokenDevices({
        operatingSystem: ['iOS']
      });

      expect(iosDevices).toHaveLength(1);
      expect(iosDevices[0].operatingSystem).toBe('iOS');
    });
  });

  // ============================================================================
  // Certificate Renewal Tests (Report 27)
  // ============================================================================

  describe('Certificate Renewal Report', () => {
    it('should fetch devices with certificate issues', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Device-1',
            userPrincipalName: 'user1@contoso.com',
            emailAddress: 'user1@contoso.com',
            operatingSystem: 'Windows',
            complianceState: 'noncompliant'
          }
        ],
        '@odata.nextLink': null
      };

      const mockComplianceStates = {
        value: [
          {
            displayName: 'Certificate Compliance',
            settingName: 'certificate.required',
            state: 'error'
          }
        ]
      };

      const mockConfigStates = {
        value: [
          {
            displayName: 'SCEP Certificate Profile',
            state: 'error'
          }
        ]
      };

      (mockGraphClient.get as any)
        .mockResolvedValueOnce(mockDevices)
        .mockResolvedValueOnce(mockComplianceStates)
        .mockResolvedValueOnce(mockConfigStates);

      const issues = await securityReports.getCertificateRenewalIssues();

      expect(issues).toBeInstanceOf(Array);
    });

    it('should generate certificate renewal report with summary', async () => {
      const mockDevices = {
        value: [],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const report = await securityReports.generateCertificateRenewalReport();

      expect(report.issues).toBeDefined();
      expect(report.summary).toBeDefined();
      expect(report.summary.totalDevices).toBeDefined();
      expect(report.summary.bySeverity).toBeDefined();
      expect(report.summary.byCertificateType).toBeDefined();
    });

    it('should filter certificate issues by severity', async () => {
      const mockDevices = {
        value: [],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const issues = await securityReports.getCertificateRenewalIssues({
        severity: ['Critical', 'High']
      });

      expect(issues).toBeInstanceOf(Array);
    });
  });

  // ============================================================================
  // Health Attestation Tests (Report 16)
  // ============================================================================

  describe('Health Attestation Report', () => {
    it('should fetch devices with health attestation state', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Windows-Device-1',
            userPrincipalName: 'user1@contoso.com',
            emailAddress: 'user1@contoso.com',
            operatingSystem: 'Windows',
            osVersion: '10.0.19045',
            model: 'Surface Pro 9',
            manufacturer: 'Microsoft',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'compliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'enabled',
              secureBoot: 'enabled',
              codeIntegrity: 'enabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'enabled',
              virtualSecureMode: 'enabled',
              tpmVersion: '2.0',
              pcrHashAlgorithm: 'sha256',
              bootManagerVersion: '10.0.19041.1',
              codeIntegrityCheckVersion: '10.0.19041.1'
            }
          },
          {
            id: 'device-2',
            deviceName: 'Windows-Device-2',
            userPrincipalName: 'user2@contoso.com',
            emailAddress: 'user2@contoso.com',
            operatingSystem: 'Windows',
            osVersion: '10.0.19045',
            model: 'Dell Latitude',
            manufacturer: 'Dell',
            lastSyncDateTime: '2024-01-15T11:00:00Z',
            complianceState: 'noncompliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'disabled',
              secureBoot: 'disabled',
              codeIntegrity: 'disabled',
              bootDebugging: 'enabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'enabled',
              earlyLaunchAntiMalwareDriverProtection: 'disabled',
              virtualSecureMode: 'disabled',
              tpmVersion: '1.2'
            }
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const devices = await securityReports.getDevicesByHealthAttestationState();

      expect(devices).toHaveLength(2);
      expect(devices[0].deviceName).toBe('Windows-Device-1');
      expect(devices[0].healthAttestationSupported).toBe(true);
      expect(devices[0].bitLockerStatus).toBe('Enabled');
      expect(devices[0].secureBootEnabled).toBe(true);
      expect(devices[0].healthScore).toBeGreaterThan(0);
      expect(devices[0].healthScoreRange).toBeDefined();

      expect(devices[1].deviceName).toBe('Windows-Device-2');
      expect(devices[1].bitLockerStatus).toBe('Disabled');
      expect(devices[1].testSigningEnabled).toBe(true);
      expect(devices[1].securityIssues).toBeInstanceOf(Array);
      expect(devices[1].securityIssues.length).toBeGreaterThan(0);
    });

    it('should calculate health scores correctly', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-excellent',
            deviceName: 'Excellent-Device',
            userPrincipalName: 'user@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'compliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'enabled',
              secureBoot: 'enabled',
              codeIntegrity: 'enabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'enabled',
              virtualSecureMode: 'enabled',
              tpmVersion: '2.0'
            }
          },
          {
            id: 'device-poor',
            deviceName: 'Poor-Device',
            userPrincipalName: 'user@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'noncompliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'disabled',
              secureBoot: 'disabled',
              codeIntegrity: 'disabled',
              bootDebugging: 'enabled',
              operatingSystemKernelDebugging: 'enabled',
              testSigning: 'enabled',
              earlyLaunchAntiMalwareDriverProtection: 'disabled',
              virtualSecureMode: 'disabled'
            }
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const devices = await securityReports.getDevicesByHealthAttestationState();

      expect(devices[0].healthScore).toBeGreaterThanOrEqual(90);
      expect(devices[0].healthScoreRange).toBe('Excellent');
      expect(devices[0].securityIssues.length).toBeLessThanOrEqual(1);

      expect(devices[1].healthScore).toBeLessThan(50);
      expect(devices[1].healthScoreRange).toMatch(/Poor|Critical/);
      expect(devices[1].securityIssues.length).toBeGreaterThan(3);
    });

    it('should generate health attestation report with summary', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Device-1',
            userPrincipalName: 'user1@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'compliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'enabled',
              secureBoot: 'enabled',
              codeIntegrity: 'enabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'enabled',
              virtualSecureMode: 'enabled',
              tpmVersion: '2.0'
            }
          },
          {
            id: 'device-2',
            deviceName: 'Device-2',
            userPrincipalName: 'user2@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'noncompliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'disabled',
              secureBoot: 'disabled',
              codeIntegrity: 'disabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'disabled',
              virtualSecureMode: 'disabled'
            }
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const report = await securityReports.generateHealthAttestationReport();

      expect(report.devices).toHaveLength(2);
      expect(report.summary).toBeDefined();
      expect(report.summary.totalDevices).toBe(2);
      expect(report.summary.supportedDevices).toBe(2);
      expect(report.summary.bitLockerEnabled).toBe(1);
      expect(report.summary.bitLockerDisabled).toBe(1);
      expect(report.summary.bitLockerPercentage).toBe(50);
      expect(report.summary.secureBootEnabled).toBe(1);
      expect(report.summary.secureBootPercentage).toBe(50);
      expect(report.summary.averageHealthScore).toBeGreaterThan(0);
      expect(report.summary.criticalIssues).toBeInstanceOf(Array);
      expect(report.summary.byPlatform).toBeDefined();
      expect(report.summary.tpmVersions).toBeDefined();
    });

    it('should identify critical health issues', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Device-1',
            userPrincipalName: 'user@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'compliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'disabled',
              secureBoot: 'disabled',
              codeIntegrity: 'disabled',
              testSigning: 'enabled',
              bootDebugging: 'enabled',
              operatingSystemKernelDebugging: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'disabled',
              virtualSecureMode: 'disabled'
            }
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const report = await securityReports.generateHealthAttestationReport();

      expect(report.summary.criticalIssues.length).toBeGreaterThan(0);

      const testSigningIssue = report.summary.criticalIssues.find(
        i => i.issueType === 'test_signing_enabled'
      );
      expect(testSigningIssue).toBeDefined();
      expect(testSigningIssue?.severity).toBe('Critical');
    });

    it('should filter devices by health score', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Device-1',
            userPrincipalName: 'user1@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'compliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'enabled',
              secureBoot: 'enabled',
              codeIntegrity: 'enabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'enabled',
              virtualSecureMode: 'enabled'
            }
          },
          {
            id: 'device-2',
            deviceName: 'Device-2',
            userPrincipalName: 'user2@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'noncompliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'disabled',
              secureBoot: 'disabled',
              codeIntegrity: 'disabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'disabled',
              virtualSecureMode: 'disabled'
            }
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const allDevices = await securityReports.getDevicesByHealthAttestationState();
      const poorDevices = allDevices.filter(d => d.healthScore < 50);

      expect(poorDevices.length).toBeGreaterThan(0);
    });

    it('should generate appropriate recommendations', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Device-1',
            userPrincipalName: 'user@contoso.com',
            operatingSystem: 'Windows',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            complianceState: 'noncompliant',
            deviceHealthAttestationState: {
              healthAttestationSupportedStatus: 'supported',
              bitLockerStatus: 'disabled',
              secureBoot: 'disabled',
              codeIntegrity: 'disabled',
              bootDebugging: 'disabled',
              operatingSystemKernelDebugging: 'disabled',
              testSigning: 'disabled',
              earlyLaunchAntiMalwareDriverProtection: 'disabled',
              virtualSecureMode: 'disabled'
            }
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const devices = await securityReports.getDevicesByHealthAttestationState();

      expect(devices[0].recommendations).toBeInstanceOf(Array);
      expect(devices[0].recommendations.length).toBeGreaterThan(0);
      expect(devices[0].recommendations.some(r => r.includes('BitLocker'))).toBe(true);
      expect(devices[0].recommendations.some(r => r.includes('Secure Boot'))).toBe(true);
    });
  });

  // ============================================================================
  // Export Functions Tests
  // ============================================================================

  describe('Export Functionality', () => {
    it('should handle pagination correctly', async () => {
      const mockPage1 = {
        value: [{ id: '1', jailBroken: true, deviceName: 'Device1' }],
        '@odata.nextLink': '/devices?$skip=1'
      };

      const mockPage2 = {
        value: [{ id: '2', jailBroken: true, deviceName: 'Device2' }],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any)
        .mockResolvedValueOnce(mockPage1)
        .mockResolvedValueOnce(mockPage2);

      const devices = await securityReports.getJailbrokenDevices();

      expect(devices).toHaveLength(2);
    });

    it('should handle API errors gracefully', async () => {
      (mockGraphClient.get as any).mockRejectedValue(
        new Error('Graph API error')
      );

      await expect(
        securityReports.getJailbrokenDevices()
      ).rejects.toThrow();
    });

    it('should handle empty results', async () => {
      (mockGraphClient.get as any).mockResolvedValue({
        value: [],
        '@odata.nextLink': null
      });

      const devices = await securityReports.getJailbrokenDevices();

      expect(devices).toHaveLength(0);
    });
  });

  // ============================================================================
  // Integration Tests
  // ============================================================================

  describe('Integration Tests', () => {
    it('should generate all three reports successfully', async () => {
      const mockDevices = {
        value: [
          {
            id: 'device-1',
            deviceName: 'Test-Device',
            userPrincipalName: 'user@contoso.com',
            operatingSystem: 'iOS',
            jailBroken: true,
            complianceState: 'noncompliant',
            lastSyncDateTime: '2024-01-15T10:00:00Z',
            enrolledDateTime: '2023-01-01T10:00:00Z',
            deviceHealthAttestationState: {}
          }
        ],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const [jailbreakReport, certReport, healthReport] = await Promise.all([
        securityReports.generateJailbreakReport(),
        securityReports.generateCertificateRenewalReport(),
        securityReports.generateHealthAttestationReport()
      ]);

      expect(jailbreakReport).toBeDefined();
      expect(jailbreakReport.devices).toBeDefined();
      expect(jailbreakReport.summary).toBeDefined();

      expect(certReport).toBeDefined();
      expect(certReport.issues).toBeDefined();
      expect(certReport.summary).toBeDefined();

      expect(healthReport).toBeDefined();
      expect(healthReport.devices).toBeDefined();
      expect(healthReport.summary).toBeDefined();
    });

    it('should apply filters across all report types', async () => {
      const mockDevices = {
        value: [],
        '@odata.nextLink': null
      };

      (mockGraphClient.get as any).mockResolvedValue(mockDevices);

      const filters = {
        complianceState: ['noncompliant'],
        operatingSystem: ['Windows']
      };

      const [jailbroken, certs, health] = await Promise.all([
        securityReports.getJailbrokenDevices(filters),
        securityReports.getCertificateRenewalIssues(filters),
        securityReports.getDevicesByHealthAttestationState(filters)
      ]);

      expect(jailbroken).toBeInstanceOf(Array);
      expect(certs).toBeInstanceOf(Array);
      expect(health).toBeInstanceOf(Array);
    });
  });
});
