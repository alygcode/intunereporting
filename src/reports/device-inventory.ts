import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

/**
 * Device Inventory Report
 * Provides comprehensive inventory of all managed devices
 */
export class DeviceInventoryReport extends BaseReport {
  name = 'device-inventory';
  description = 'Comprehensive inventory of all managed devices';
  category = 'Devices';
  enabled = true;

  async execute(): Promise<ReportData> {
    logger.info('Executing Device Inventory Report');

    try {
      // Fetch all managed devices
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'operatingSystem',
            'osVersion',
            'manufacturer',
            'model',
            'serialNumber',
            'enrolledDateTime',
            'lastSyncDateTime',
            'complianceState',
            'managementAgent',
            'userPrincipalName',
            'userDisplayName',
            'deviceEnrollmentType',
            'azureADDeviceId',
          ])
          .get()
      );

      const devices = await this.getAllPages(response);

      // Transform data for reporting
      const data = devices.map((device: any) => ({
        deviceName: device.deviceName || 'Unknown',
        operatingSystem: device.operatingSystem || 'Unknown',
        osVersion: device.osVersion || 'Unknown',
        manufacturer: device.manufacturer || 'Unknown',
        model: device.model || 'Unknown',
        serialNumber: device.serialNumber || 'N/A',
        enrolledDate: device.enrolledDateTime
          ? new Date(device.enrolledDateTime).toLocaleDateString()
          : 'N/A',
        lastSync: device.lastSyncDateTime
          ? new Date(device.lastSyncDateTime).toLocaleDateString()
          : 'Never',
        complianceState: device.complianceState || 'Unknown',
        managementAgent: device.managementAgent || 'Unknown',
        userPrincipalName: device.userPrincipalName || 'N/A',
        userDisplayName: device.userDisplayName || 'N/A',
        enrollmentType: device.deviceEnrollmentType || 'Unknown',
      }));

      // Calculate summary statistics
      const summary = {
        totalDevices: data.length,
        byOS: this.groupBy(data, 'operatingSystem'),
        byCompliance: this.groupBy(data, 'complianceState'),
        byManagementAgent: this.groupBy(data, 'managementAgent'),
      };

      logger.info('Device Inventory Report completed', { deviceCount: data.length });

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute Device Inventory Report', error);
      throw error;
    }
  }

  private groupBy(data: any[], key: string): Record<string, number> {
    return data.reduce((acc, item) => {
      const value = item[key] || 'Unknown';
      acc[value] = (acc[value] || 0) + 1;
      return acc;
    }, {});
  }
}
