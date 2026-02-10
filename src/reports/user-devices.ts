import { BaseReport } from './base-report';
import { ReportData } from '../types';
import { Logger } from '../core/logger';

const logger = Logger.getInstance();

/**
 * User Devices Report
 * Shows device assignments per user
 */
export class UserDevicesReport extends BaseReport {
  name = 'user-devices';
  description = 'Device assignments and ownership by user';
  category = 'Users';
  enabled = true;

  async execute(): Promise<ReportData> {
    logger.info('Executing User Devices Report');

    try {
      // Fetch all managed devices with user information
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'operatingSystem',
            'userPrincipalName',
            'userDisplayName',
            'emailAddress',
            'complianceState',
            'lastSyncDateTime',
            'enrolledDateTime',
          ])
          .get()
      );

      const devices = await this.getAllPages(response);

      // Group devices by user
      const userDeviceMap = new Map<string, any[]>();

      devices.forEach((device: any) => {
        const userKey = device.userPrincipalName || 'Unassigned';
        if (!userDeviceMap.has(userKey)) {
          userDeviceMap.set(userKey, []);
        }
        userDeviceMap.get(userKey)!.push(device);
      });

      // Transform data for reporting
      const data: any[] = [];
      userDeviceMap.forEach((userDevices, userPrincipalName) => {
        userDevices.forEach((device: any) => {
          data.push({
            userPrincipalName,
            userDisplayName: device.userDisplayName || 'N/A',
            emailAddress: device.emailAddress || 'N/A',
            deviceName: device.deviceName || 'Unknown',
            operatingSystem: device.operatingSystem || 'Unknown',
            complianceState: device.complianceState || 'Unknown',
            enrolledDate: device.enrolledDateTime
              ? new Date(device.enrolledDateTime).toLocaleDateString()
              : 'N/A',
            lastSync: device.lastSyncDateTime
              ? new Date(device.lastSyncDateTime).toLocaleDateString()
              : 'Never',
          });
        });
      });

      // Calculate summary statistics
      const summary = {
        totalUsers: userDeviceMap.size,
        totalDevices: devices.length,
        averageDevicesPerUser: (devices.length / userDeviceMap.size).toFixed(2),
        usersWithMultipleDevices: Array.from(userDeviceMap.values()).filter(
          (devices) => devices.length > 1
        ).length,
      };

      logger.info('User Devices Report completed', {
        userCount: userDeviceMap.size,
        deviceCount: devices.length,
      });

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute User Devices Report', error);
      throw error;
    }
  }
}
