import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

/**
 * Compliance Summary Report
 * Provides overview of device compliance status
 */
export class ComplianceSummaryReport extends BaseReport {
  name = 'compliance-summary';
  description = 'Overview of device compliance status across the organization';
  category = 'Compliance';
  enabled = true;

  async execute(): Promise<ReportData> {
    logger.info('Executing Compliance Summary Report');

    try {
      // Fetch device compliance policies
      const policiesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/deviceCompliancePolicies')
          .select(['id', 'displayName', 'description'])
          .get()
      );

      const policies = await this.getAllPages(policiesResponse);

      // Fetch devices with compliance information
      const devicesResponse = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'operatingSystem',
            'complianceState',
            'userPrincipalName',
            'lastSyncDateTime',
          ])
          .get()
      );

      const devices = await this.getAllPages(devicesResponse);

      // Build compliance data
      const data = devices.map((device: any) => ({
        deviceName: device.deviceName || 'Unknown',
        operatingSystem: device.operatingSystem || 'Unknown',
        complianceState: device.complianceState || 'Unknown',
        userPrincipalName: device.userPrincipalName || 'N/A',
        lastSync: device.lastSyncDateTime
          ? new Date(device.lastSyncDateTime).toLocaleDateString()
          : 'Never',
      }));

      // Calculate compliance statistics
      const complianceStats = {
        compliant: devices.filter((d: any) => d.complianceState === 'compliant').length,
        nonCompliant: devices.filter((d: any) => d.complianceState === 'noncompliant').length,
        inGracePeriod: devices.filter((d: any) => d.complianceState === 'inGracePeriod').length,
        unknown: devices.filter(
          (d: any) => !d.complianceState || d.complianceState === 'unknown'
        ).length,
      };

      const summary = {
        totalDevices: devices.length,
        totalPolicies: policies.length,
        compliant: complianceStats.compliant,
        nonCompliant: complianceStats.nonCompliant,
        inGracePeriod: complianceStats.inGracePeriod,
        unknown: complianceStats.unknown,
        complianceRate: devices.length > 0
          ? `${((complianceStats.compliant / devices.length) * 100).toFixed(2)}%`
          : '0%',
      };

      logger.info('Compliance Summary Report completed', { deviceCount: devices.length });

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute Compliance Summary Report', error);
      throw error;
    }
  }
}
