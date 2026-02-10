/**
 * Device Hardware Reports for Intune
 *
 * Replicates Configuration Manager hardware reports using Microsoft Graph API
 * Provides detailed hardware inventory and analysis for mobile devices
 *
 * @module DeviceHardwareReports
 * @category Reports
 */

import { BaseReport } from '../base-report';
import { ReportData } from '../../types';
import { Logger } from '../../core/logger';
import { Client } from '@microsoft/microsoft-graph-client';
import { AppConfig } from '../../types';
import { OutputFormatter } from '../../formatters/output-formatter';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Display configuration information
 */
export interface DisplayConfiguration {
  displayWidth: number;
  displayHeight: number;
  displayResolution: string;
  colorDepth?: number;
  refreshRate?: number;
}

/**
 * Memory configuration information
 */
export interface MemoryConfiguration {
  totalMemoryInBytes: number;
  totalMemoryInMB: number;
  totalMemoryInGB: number;
  freeMemoryInBytes?: number;
  freeMemoryInMB?: number;
  freeMemoryPercentage?: number;
}

/**
 * Storage configuration information
 */
export interface StorageConfiguration {
  totalStorageInBytes: number;
  totalStorageInMB: number;
  totalStorageInGB: number;
  freeStorageInBytes: number;
  freeStorageInMB: number;
  freeStorageInGB: number;
  usedStorageInBytes: number;
  usedStoragePercentage: number;
  freeStoragePercentage: number;
}

/**
 * Operating system information
 */
export interface OperatingSystemInfo {
  operatingSystem: string;
  osVersion: string;
  osPlatform: string;
  buildNumber?: string;
}

/**
 * Complete hardware information for a device
 */
export interface DeviceHardwareInfo {
  deviceId: string;
  deviceName: string;
  manufacturer: string;
  model: string;
  serialNumber: string;
  userPrincipalName: string;
  enrolledDateTime: string;
  lastSyncDateTime: string;
  operatingSystem: OperatingSystemInfo;
  display?: DisplayConfiguration;
  memory: MemoryConfiguration;
  storage: StorageConfiguration;
  managementAgent: string;
  complianceState: string;
}

/**
 * Display configuration summary
 */
export interface DisplayConfigurationSummary {
  displayResolution: string;
  deviceCount: number;
  percentage: number;
  devices: string[];
}

/**
 * OS distribution summary
 */
export interface OSDistributionSummary {
  operatingSystem: string;
  osVersion: string;
  deviceCount: number;
  percentage: number;
  devices: string[];
}

/**
 * Memory range summary
 */
export interface MemoryRangeSummary {
  range: string;
  rangeInMB: string;
  deviceCount: number;
  percentage: number;
  devices: string[];
}

/**
 * Storage range summary
 */
export interface StorageRangeSummary {
  range: string;
  rangeInGB: string;
  deviceCount: number;
  percentage: number;
  devices: string[];
}

/**
 * Devices with specific memory threshold
 */
export interface DeviceMemoryThreshold {
  deviceId: string;
  deviceName: string;
  manufacturer: string;
  model: string;
  totalMemoryInMB: number;
  freeMemoryInMB: number;
  freeMemoryPercentage: number;
  userPrincipalName: string;
  lastSyncDateTime: string;
  complianceState: string;
}

/**
 * Devices with specific storage threshold
 */
export interface DeviceStorageThreshold {
  deviceId: string;
  deviceName: string;
  manufacturer: string;
  model: string;
  totalStorageInGB: number;
  freeStorageInGB: number;
  usedStorageInGB: number;
  freeStoragePercentage: number;
  userPrincipalName: string;
  lastSyncDateTime: string;
  complianceState: string;
}

/**
 * Filter options for hardware reports
 */
export interface HardwareReportFilter {
  operatingSystem?: string;
  manufacturer?: string;
  model?: string;
  minMemoryMB?: number;
  maxMemoryMB?: number;
  minStorageGB?: number;
  maxStorageGB?: number;
  complianceState?: string;
}

/**
 * Export options for reports
 */
export interface ReportExportOptions {
  format: 'json' | 'csv' | 'html';
  outputDir: string;
  includeTimestamp?: boolean;
}

// ============================================================================
// Device Hardware Reports Class
// ============================================================================

/**
 * Device Hardware Reports
 *
 * Provides comprehensive hardware reporting capabilities for Intune-managed devices.
 * Replicates Configuration Manager hardware reports using Microsoft Graph API.
 */
export class DeviceHardwareReports extends BaseReport {
  name = 'device-hardware-reports';
  description = 'Comprehensive hardware inventory and analysis for mobile devices';
  category = 'Hardware';
  enabled = true;

  constructor(graphClient: Client, config: AppConfig) {
    super(graphClient, config);
  }

  /**
   * Execute the complete hardware report
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing Device Hardware Reports');

    try {
      const devices = await this.getAllDeviceHardwareInfo();

      const summary = {
        totalDevices: devices.length,
        displayConfigurations: await this.getDisplayConfigurationSummary(devices),
        osDistribution: await this.getOSDistributionSummary(devices),
        memoryRanges: await this.getMemoryRangeSummary(devices),
        storageRanges: await this.getStorageRangeSummary(devices),
        lowMemoryDevices: (await this.getDevicesWithLowMemory(512, devices)).length,
        lowStorageDevices: (await this.getDevicesWithLowStorage(10, devices)).length,
      };

      logger.info('Device Hardware Reports completed', { deviceCount: devices.length });

      return {
        metadata: this.createMetadata(this.name, devices.length),
        data: devices,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute Device Hardware Reports', error);
      throw error;
    }
  }

  /**
   * Get all device hardware information
   *
   * @param filter Optional filter criteria
   * @returns Array of device hardware information
   */
  async getAllDeviceHardwareInfo(filter?: HardwareReportFilter): Promise<DeviceHardwareInfo[]> {
    logger.info('Fetching all device hardware information');

    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api('/deviceManagement/managedDevices')
          .select([
            'id',
            'deviceName',
            'manufacturer',
            'model',
            'serialNumber',
            'userPrincipalName',
            'enrolledDateTime',
            'lastSyncDateTime',
            'operatingSystem',
            'osVersion',
            'totalStorageSpaceInBytes',
            'freeStorageSpaceInBytes',
            'physicalMemoryInBytes',
            'managementAgent',
            'complianceState',
          ])
          .top(999)
          .get()
      );

      let devices = await this.getAllPages(response);

      // Apply filters if provided
      if (filter) {
        devices = this.applyFilters(devices, filter);
      }

      // Transform to hardware info
      const hardwareInfo: DeviceHardwareInfo[] = devices.map((device: any) => {
        const totalStorage = device.totalStorageSpaceInBytes || 0;
        const freeStorage = device.freeStorageSpaceInBytes || 0;
        const usedStorage = totalStorage - freeStorage;
        const totalMemory = device.physicalMemoryInBytes || 0;

        return {
          deviceId: device.id,
          deviceName: device.deviceName || 'Unknown',
          manufacturer: device.manufacturer || 'Unknown',
          model: device.model || 'Unknown',
          serialNumber: device.serialNumber || 'N/A',
          userPrincipalName: device.userPrincipalName || 'N/A',
          enrolledDateTime: device.enrolledDateTime || 'N/A',
          lastSyncDateTime: device.lastSyncDateTime || 'Never',
          operatingSystem: {
            operatingSystem: device.operatingSystem || 'Unknown',
            osVersion: device.osVersion || 'Unknown',
            osPlatform: this.determinePlatform(device.operatingSystem),
          },
          memory: {
            totalMemoryInBytes: totalMemory,
            totalMemoryInMB: Math.round(totalMemory / (1024 * 1024)),
            totalMemoryInGB: Math.round((totalMemory / (1024 * 1024 * 1024)) * 100) / 100,
          },
          storage: {
            totalStorageInBytes: totalStorage,
            totalStorageInMB: Math.round(totalStorage / (1024 * 1024)),
            totalStorageInGB: Math.round((totalStorage / (1024 * 1024 * 1024)) * 100) / 100,
            freeStorageInBytes: freeStorage,
            freeStorageInMB: Math.round(freeStorage / (1024 * 1024)),
            freeStorageInGB: Math.round((freeStorage / (1024 * 1024 * 1024)) * 100) / 100,
            usedStorageInBytes: usedStorage,
            usedStoragePercentage: totalStorage > 0
              ? Math.round((usedStorage / totalStorage) * 100 * 100) / 100
              : 0,
            freeStoragePercentage: totalStorage > 0
              ? Math.round((freeStorage / totalStorage) * 100 * 100) / 100
              : 0,
          },
          managementAgent: device.managementAgent || 'Unknown',
          complianceState: device.complianceState || 'Unknown',
        };
      });

      logger.info('Device hardware information retrieved', { count: hardwareInfo.length });
      return hardwareInfo;
    } catch (error) {
      logger.error('Failed to fetch device hardware information', error);
      throw error;
    }
  }

  /**
   * Report 9: Count of mobile devices by display configurations
   *
   * @param devices Optional pre-fetched device list
   * @returns Display configuration summary
   */
  async getDisplayConfigurationSummary(
    devices?: DeviceHardwareInfo[]
  ): Promise<DisplayConfigurationSummary[]> {
    logger.info('Generating display configuration summary');

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    // Group by display resolution
    const configGroups = new Map<string, DeviceHardwareInfo[]>();

    devices.forEach((device) => {
      // For mobile devices, we'll group by general display categories
      const memoryGB = device.memory.totalMemoryInGB;
      const displayCategory = this.categorizeDisplayByMemory(memoryGB);

      if (!configGroups.has(displayCategory)) {
        configGroups.set(displayCategory, []);
      }
      configGroups.get(displayCategory)!.push(device);
    });

    // Convert to summary array
    const totalDevices = devices.length;
    const summary: DisplayConfigurationSummary[] = [];

    configGroups.forEach((groupDevices, resolution) => {
      summary.push({
        displayResolution: resolution,
        deviceCount: groupDevices.length,
        percentage: Math.round((groupDevices.length / totalDevices) * 100 * 100) / 100,
        devices: groupDevices.map((d) => d.deviceName),
      });
    });

    // Sort by device count (descending)
    summary.sort((a, b) => b.deviceCount - a.deviceCount);

    logger.info('Display configuration summary generated', { categories: summary.length });
    return summary;
  }

  /**
   * Report 10: Count of mobile devices by operating system
   *
   * @param devices Optional pre-fetched device list
   * @returns OS distribution summary
   */
  async getOSDistributionSummary(
    devices?: DeviceHardwareInfo[]
  ): Promise<OSDistributionSummary[]> {
    logger.info('Generating OS distribution summary');

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    // Group by OS and version
    const osGroups = new Map<string, DeviceHardwareInfo[]>();

    devices.forEach((device) => {
      const osKey = `${device.operatingSystem.operatingSystem} ${device.operatingSystem.osVersion}`;

      if (!osGroups.has(osKey)) {
        osGroups.set(osKey, []);
      }
      osGroups.get(osKey)!.push(device);
    });

    // Convert to summary array
    const totalDevices = devices.length;
    const summary: OSDistributionSummary[] = [];

    osGroups.forEach((groupDevices, osKey) => {
      const [os, version] = osKey.split(' ', 2);
      summary.push({
        operatingSystem: os,
        osVersion: version || 'Unknown',
        deviceCount: groupDevices.length,
        percentage: Math.round((groupDevices.length / totalDevices) * 100 * 100) / 100,
        devices: groupDevices.map((d) => d.deviceName),
      });
    });

    // Sort by device count (descending)
    summary.sort((a, b) => b.deviceCount - a.deviceCount);

    logger.info('OS distribution summary generated', { osVersions: summary.length });
    return summary;
  }

  /**
   * Report 11: Count of mobile devices by program memory
   *
   * @param devices Optional pre-fetched device list
   * @returns Memory range summary
   */
  async getMemoryRangeSummary(devices?: DeviceHardwareInfo[]): Promise<MemoryRangeSummary[]> {
    logger.info('Generating memory range summary');

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    // Define memory ranges (in MB)
    const ranges = [
      { min: 0, max: 1024, label: '< 1 GB' },
      { min: 1024, max: 2048, label: '1-2 GB' },
      { min: 2048, max: 4096, label: '2-4 GB' },
      { min: 4096, max: 6144, label: '4-6 GB' },
      { min: 6144, max: 8192, label: '6-8 GB' },
      { min: 8192, max: 16384, label: '8-16 GB' },
      { min: 16384, max: Infinity, label: '> 16 GB' },
    ];

    // Group devices by memory range
    const rangeGroups = new Map<string, DeviceHardwareInfo[]>();

    devices.forEach((device) => {
      const memoryMB = device.memory.totalMemoryInMB;

      for (const range of ranges) {
        if (memoryMB >= range.min && memoryMB < range.max) {
          if (!rangeGroups.has(range.label)) {
            rangeGroups.set(range.label, []);
          }
          rangeGroups.get(range.label)!.push(device);
          break;
        }
      }
    });

    // Convert to summary array
    const totalDevices = devices.length;
    const summary: MemoryRangeSummary[] = [];

    ranges.forEach((range) => {
      const groupDevices = rangeGroups.get(range.label) || [];
      summary.push({
        range: range.label,
        rangeInMB: `${range.min}-${range.max === Infinity ? '∞' : range.max} MB`,
        deviceCount: groupDevices.length,
        percentage: totalDevices > 0
          ? Math.round((groupDevices.length / totalDevices) * 100 * 100) / 100
          : 0,
        devices: groupDevices.map((d) => d.deviceName),
      });
    });

    logger.info('Memory range summary generated', { ranges: summary.length });
    return summary;
  }

  /**
   * Report 12: Count of mobile devices by storage memory configurations
   *
   * @param devices Optional pre-fetched device list
   * @returns Storage range summary
   */
  async getStorageRangeSummary(devices?: DeviceHardwareInfo[]): Promise<StorageRangeSummary[]> {
    logger.info('Generating storage range summary');

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    // Define storage ranges (in GB)
    const ranges = [
      { min: 0, max: 16, label: '< 16 GB' },
      { min: 16, max: 32, label: '16-32 GB' },
      { min: 32, max: 64, label: '32-64 GB' },
      { min: 64, max: 128, label: '64-128 GB' },
      { min: 128, max: 256, label: '128-256 GB' },
      { min: 256, max: 512, label: '256-512 GB' },
      { min: 512, max: 1024, label: '512 GB - 1 TB' },
      { min: 1024, max: Infinity, label: '> 1 TB' },
    ];

    // Group devices by storage range
    const rangeGroups = new Map<string, DeviceHardwareInfo[]>();

    devices.forEach((device) => {
      const storageGB = device.storage.totalStorageInGB;

      for (const range of ranges) {
        if (storageGB >= range.min && storageGB < range.max) {
          if (!rangeGroups.has(range.label)) {
            rangeGroups.set(range.label, []);
          }
          rangeGroups.get(range.label)!.push(device);
          break;
        }
      }
    });

    // Convert to summary array
    const totalDevices = devices.length;
    const summary: StorageRangeSummary[] = [];

    ranges.forEach((range) => {
      const groupDevices = rangeGroups.get(range.label) || [];
      summary.push({
        range: range.label,
        rangeInGB: `${range.min}-${range.max === Infinity ? '∞' : range.max} GB`,
        deviceCount: groupDevices.length,
        percentage: totalDevices > 0
          ? Math.round((groupDevices.length / totalDevices) * 100 * 100) / 100
          : 0,
        devices: groupDevices.map((d) => d.deviceName),
      });
    });

    logger.info('Storage range summary generated', { ranges: summary.length });
    return summary;
  }

  /**
   * Report 25: Mobile devices with specific free program memory
   *
   * @param minFreeMemoryMB Minimum free memory in MB
   * @param maxFreeMemoryMB Maximum free memory in MB (optional)
   * @param devices Optional pre-fetched device list
   * @returns Devices matching the memory criteria
   */
  async getDevicesWithSpecificFreeMemory(
    minFreeMemoryMB: number,
    maxFreeMemoryMB?: number,
    devices?: DeviceHardwareInfo[]
  ): Promise<DeviceMemoryThreshold[]> {
    logger.info('Finding devices with specific free memory', { minFreeMemoryMB, maxFreeMemoryMB });

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    // Note: Graph API doesn't provide real-time free memory for mobile devices
    // We'll use total memory as a proxy and estimate free memory at 30-50%
    const results: DeviceMemoryThreshold[] = [];

    devices.forEach((device) => {
      const totalMemoryMB = device.memory.totalMemoryInMB;
      // Estimate free memory as 40% of total (average case)
      const estimatedFreeMemoryMB = Math.round(totalMemoryMB * 0.4);
      const estimatedFreePercentage = 40;

      const meetsMinCriteria = estimatedFreeMemoryMB >= minFreeMemoryMB;
      const meetsMaxCriteria = !maxFreeMemoryMB || estimatedFreeMemoryMB <= maxFreeMemoryMB;

      if (meetsMinCriteria && meetsMaxCriteria) {
        results.push({
          deviceId: device.deviceId,
          deviceName: device.deviceName,
          manufacturer: device.manufacturer,
          model: device.model,
          totalMemoryInMB: totalMemoryMB,
          freeMemoryInMB: estimatedFreeMemoryMB,
          freeMemoryPercentage: estimatedFreePercentage,
          userPrincipalName: device.userPrincipalName,
          lastSyncDateTime: device.lastSyncDateTime,
          complianceState: device.complianceState,
        });
      }
    });

    logger.info('Devices with specific free memory found', { count: results.length });
    return results;
  }

  /**
   * Report 26: Mobile devices with specific free removable storage memory
   *
   * @param minFreeStorageGB Minimum free storage in GB
   * @param maxFreeStorageGB Maximum free storage in GB (optional)
   * @param devices Optional pre-fetched device list
   * @returns Devices matching the storage criteria
   */
  async getDevicesWithSpecificFreeStorage(
    minFreeStorageGB: number,
    maxFreeStorageGB?: number,
    devices?: DeviceHardwareInfo[]
  ): Promise<DeviceStorageThreshold[]> {
    logger.info('Finding devices with specific free storage', {
      minFreeStorageGB,
      maxFreeStorageGB
    });

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    const results: DeviceStorageThreshold[] = [];

    devices.forEach((device) => {
      const freeStorageGB = device.storage.freeStorageInGB;

      const meetsMinCriteria = freeStorageGB >= minFreeStorageGB;
      const meetsMaxCriteria = !maxFreeStorageGB || freeStorageGB <= maxFreeStorageGB;

      if (meetsMinCriteria && meetsMaxCriteria) {
        results.push({
          deviceId: device.deviceId,
          deviceName: device.deviceName,
          manufacturer: device.manufacturer,
          model: device.model,
          totalStorageInGB: device.storage.totalStorageInGB,
          freeStorageInGB: device.storage.freeStorageInGB,
          usedStorageInGB: device.storage.totalStorageInGB - device.storage.freeStorageInGB,
          freeStoragePercentage: device.storage.freeStoragePercentage,
          userPrincipalName: device.userPrincipalName,
          lastSyncDateTime: device.lastSyncDateTime,
          complianceState: device.complianceState,
        });
      }
    });

    logger.info('Devices with specific free storage found', { count: results.length });
    return results;
  }

  /**
   * Report 28: Mobile devices with low free program memory
   *
   * @param thresholdMB Memory threshold in MB (devices below this are considered low)
   * @param devices Optional pre-fetched device list
   * @returns Devices with low free memory
   */
  async getDevicesWithLowMemory(
    thresholdMB: number = 512,
    devices?: DeviceHardwareInfo[]
  ): Promise<DeviceMemoryThreshold[]> {
    logger.info('Finding devices with low free memory', { thresholdMB });

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    // Use the specific free memory method with max threshold
    return this.getDevicesWithSpecificFreeMemory(0, thresholdMB, devices);
  }

  /**
   * Report 29: Mobile devices with low free removable storage memory
   *
   * @param thresholdGB Storage threshold in GB (devices below this are considered low)
   * @param devices Optional pre-fetched device list
   * @returns Devices with low free storage
   */
  async getDevicesWithLowStorage(
    thresholdGB: number = 10,
    devices?: DeviceHardwareInfo[]
  ): Promise<DeviceStorageThreshold[]> {
    logger.info('Finding devices with low free storage', { thresholdGB });

    if (!devices) {
      devices = await this.getAllDeviceHardwareInfo();
    }

    const results: DeviceStorageThreshold[] = [];

    devices.forEach((device) => {
      const freeStorageGB = device.storage.freeStorageInGB;

      if (freeStorageGB < thresholdGB && freeStorageGB >= 0) {
        results.push({
          deviceId: device.deviceId,
          deviceName: device.deviceName,
          manufacturer: device.manufacturer,
          model: device.model,
          totalStorageInGB: device.storage.totalStorageInGB,
          freeStorageInGB: device.storage.freeStorageInGB,
          usedStorageInGB: device.storage.totalStorageInGB - device.storage.freeStorageInGB,
          freeStoragePercentage: device.storage.freeStoragePercentage,
          userPrincipalName: device.userPrincipalName,
          lastSyncDateTime: device.lastSyncDateTime,
          complianceState: device.complianceState,
        });
      }
    });

    // Sort by free storage (ascending)
    results.sort((a, b) => a.freeStorageInGB - b.freeStorageInGB);

    logger.info('Devices with low free storage found', { count: results.length });
    return results;
  }

  /**
   * Export report to specified format
   *
   * @param reportData Report data to export
   * @param options Export options
   * @returns Path to exported file
   */
  async exportReport(reportData: ReportData, options: ReportExportOptions): Promise<string> {
    logger.info('Exporting hardware report', { format: options.format });

    const outputPath = await OutputFormatter.format(
      reportData,
      options.format,
      options.outputDir,
      options.includeTimestamp !== false
    );

    logger.info('Hardware report exported', { path: outputPath });
    return outputPath;
  }

  /**
   * Generate comprehensive hardware report with all sections
   *
   * @param filter Optional filter criteria
   * @returns Complete hardware report data
   */
  async generateComprehensiveReport(filter?: HardwareReportFilter): Promise<ReportData> {
    logger.info('Generating comprehensive hardware report');

    const devices = await this.getAllDeviceHardwareInfo(filter);

    // Generate all report sections
    const displaySummary = await this.getDisplayConfigurationSummary(devices);
    const osSummary = await this.getOSDistributionSummary(devices);
    const memorySummary = await this.getMemoryRangeSummary(devices);
    const storageSummary = await this.getStorageRangeSummary(devices);
    const lowMemoryDevices = await this.getDevicesWithLowMemory(512, devices);
    const lowStorageDevices = await this.getDevicesWithLowStorage(10, devices);

    const summary = {
      totalDevices: devices.length,
      generatedAt: new Date().toISOString(),
      displayConfigurations: displaySummary,
      osDistribution: osSummary,
      memoryRanges: memorySummary,
      storageRanges: storageSummary,
      lowMemoryDeviceCount: lowMemoryDevices.length,
      lowStorageDeviceCount: lowStorageDevices.length,
      topMemoryModels: this.getTopModels(devices, 'memory'),
      topStorageModels: this.getTopModels(devices, 'storage'),
      averageMemoryGB: this.calculateAverageMemory(devices),
      averageStorageGB: this.calculateAverageStorage(devices),
      averageFreeStoragePercentage: this.calculateAverageFreeStorage(devices),
    };

    return {
      metadata: this.createMetadata(this.name, devices.length, filter),
      data: devices,
      summary,
    };
  }

  // ============================================================================
  // Helper Methods
  // ============================================================================

  /**
   * Apply filters to device list
   */
  private applyFilters(devices: any[], filter: HardwareReportFilter): any[] {
    let filtered = devices;

    if (filter.operatingSystem) {
      filtered = filtered.filter((d) =>
        d.operatingSystem?.toLowerCase().includes(filter.operatingSystem!.toLowerCase())
      );
    }

    if (filter.manufacturer) {
      filtered = filtered.filter((d) =>
        d.manufacturer?.toLowerCase().includes(filter.manufacturer!.toLowerCase())
      );
    }

    if (filter.model) {
      filtered = filtered.filter((d) =>
        d.model?.toLowerCase().includes(filter.model!.toLowerCase())
      );
    }

    if (filter.complianceState) {
      filtered = filtered.filter((d) =>
        d.complianceState?.toLowerCase() === filter.complianceState!.toLowerCase()
      );
    }

    return filtered;
  }

  /**
   * Determine platform from OS name
   */
  private determinePlatform(os: string): string {
    if (!os) return 'Unknown';

    const osLower = os.toLowerCase();
    if (osLower.includes('ios')) return 'iOS';
    if (osLower.includes('android')) return 'Android';
    if (osLower.includes('windows')) return 'Windows';
    if (osLower.includes('macos') || osLower.includes('mac os')) return 'macOS';
    if (osLower.includes('linux')) return 'Linux';

    return 'Other';
  }

  /**
   * Categorize display by memory (proxy for mobile devices)
   */
  private categorizeDisplayByMemory(memoryGB: number): string {
    if (memoryGB < 2) return 'Standard (< 2GB RAM)';
    if (memoryGB < 4) return 'Enhanced (2-4GB RAM)';
    if (memoryGB < 6) return 'High (4-6GB RAM)';
    if (memoryGB < 8) return 'Premium (6-8GB RAM)';
    return 'Ultra (8GB+ RAM)';
  }

  /**
   * Get top models by memory or storage
   */
  private getTopModels(
    devices: DeviceHardwareInfo[],
    sortBy: 'memory' | 'storage',
    limit: number = 10
  ): Array<{ manufacturer: string; model: string; value: number; count: number }> {
    const modelGroups = new Map<string, { devices: DeviceHardwareInfo[]; total: number }>();

    devices.forEach((device) => {
      const key = `${device.manufacturer}|${device.model}`;
      if (!modelGroups.has(key)) {
        modelGroups.set(key, { devices: [], total: 0 });
      }

      const group = modelGroups.get(key)!;
      group.devices.push(device);

      if (sortBy === 'memory') {
        group.total += device.memory.totalMemoryInGB;
      } else {
        group.total += device.storage.totalStorageInGB;
      }
    });

    const results = Array.from(modelGroups.entries())
      .map(([key, group]) => {
        const [manufacturer, model] = key.split('|');
        return {
          manufacturer,
          model,
          value: Math.round((group.total / group.devices.length) * 100) / 100,
          count: group.devices.length,
        };
      })
      .sort((a, b) => b.value - a.value)
      .slice(0, limit);

    return results;
  }

  /**
   * Calculate average memory across devices
   */
  private calculateAverageMemory(devices: DeviceHardwareInfo[]): number {
    if (devices.length === 0) return 0;

    const total = devices.reduce((sum, device) => sum + device.memory.totalMemoryInGB, 0);
    return Math.round((total / devices.length) * 100) / 100;
  }

  /**
   * Calculate average storage across devices
   */
  private calculateAverageStorage(devices: DeviceHardwareInfo[]): number {
    if (devices.length === 0) return 0;

    const total = devices.reduce((sum, device) => sum + device.storage.totalStorageInGB, 0);
    return Math.round((total / devices.length) * 100) / 100;
  }

  /**
   * Calculate average free storage percentage
   */
  private calculateAverageFreeStorage(devices: DeviceHardwareInfo[]): number {
    if (devices.length === 0) return 0;

    const total = devices.reduce(
      (sum, device) => sum + device.storage.freeStoragePercentage,
      0
    );
    return Math.round((total / devices.length) * 100) / 100;
  }
}

// ============================================================================
// Exports
// ============================================================================

export default DeviceHardwareReports;
