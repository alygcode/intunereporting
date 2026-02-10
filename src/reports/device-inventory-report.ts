/**
 * Comprehensive Device Inventory Report Module for Microsoft Intune
 *
 * This module provides detailed device inventory reporting capabilities using the Microsoft Graph API.
 * It includes hardware inventory, software inventory, ownership types, OS platforms, and health attestation data.
 *
 * Features:
 * - Complete device inventory with detailed metadata
 * - Hardware specifications and inventory
 * - Installed applications and software inventory
 * - Device ownership types and management details
 * - OS platform-specific filtering and grouping
 * - Device health attestation data
 * - Comprehensive error handling with retry logic
 * - Pagination support for large datasets
 * - Multiple export formats (JSON, CSV, HTML)
 *
 * @module device-inventory-report
 */

import { BaseReport } from './base-report';
import { ReportData, AppConfig } from '../types';
import { Client } from '@microsoft/microsoft-graph-client';
import { Logger } from '../core/logger';
import { OutputFormatter } from '../formatters/output-formatter';

const logger = Logger.getInstance();

// ============================================================================
// Type Definitions
// ============================================================================

/**
 * Complete device information from Intune
 */
export interface ManagedDevice {
  id: string;
  deviceName: string;
  managedDeviceName?: string;
  operatingSystem: string;
  osVersion: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  imei?: string;
  meid?: string;
  wiFiMacAddress?: string;
  ethernetMacAddress?: string;

  // User information
  userPrincipalName?: string;
  userDisplayName?: string;
  userId?: string;
  emailAddress?: string;

  // Enrollment and management
  enrolledDateTime?: string;
  lastSyncDateTime?: string;
  complianceState: string;
  managementAgent: string;
  managementState?: string;
  deviceEnrollmentType?: string;
  deviceRegistrationState?: string;

  // Device identifiers
  azureADDeviceId?: string;
  azureADRegistered?: boolean;
  deviceCategoryDisplayName?: string;

  // Hardware details
  totalStorageSpaceInBytes?: number;
  freeStorageSpaceInBytes?: number;
  physicalMemoryInBytes?: number;

  // Status and health
  isEncrypted?: boolean;
  isSupervised?: boolean;
  exchangeAccessState?: string;
  exchangeAccessStateReason?: string;
  remoteAssistanceSessionUrl?: string;
  activationLockBypassCode?: string;

  // Additional metadata
  partnerReportedThreatState?: string;
  jailBroken?: string;
  easActivated?: boolean;
  easActivationDateTime?: string;

  // Ownership
  managedDeviceOwnerType?: string;
  deviceActionResults?: any[];
}

/**
 * Device hardware inventory details
 */
export interface DeviceHardwareInventory {
  deviceId: string;
  deviceName: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;

  // Storage
  totalStorageGB?: number;
  freeStorageGB?: number;
  usedStoragePercentage?: number;

  // Memory
  totalMemoryGB?: number;

  // Network
  wiFiMacAddress?: string;
  ethernetMacAddress?: string;
  imei?: string;
  meid?: string;

  // Operating System
  operatingSystem: string;
  osVersion: string;
  osBuildNumber?: string;

  // Security
  isEncrypted: boolean;
  isSupervised: boolean;
  tpmSpecificationVersion?: string;

  // Battery (for mobile devices)
  batteryHealthPercentage?: number;
  batterySerialNumber?: string;

  // Additional hardware info
  processorArchitecture?: string;
  systemManagementBIOSVersion?: string;
  phoneNumber?: string;
  subscriberCarrier?: string;
}

/**
 * Detected application on a device
 */
export interface DetectedApp {
  id: string;
  displayName: string;
  version: string;
  sizeInByte?: number;
  deviceCount?: number;
  publisher?: string;
  platform?: string;
}

/**
 * Device software inventory
 */
export interface DeviceSoftwareInventory {
  deviceId: string;
  deviceName: string;
  operatingSystem: string;
  osVersion: string;
  totalApps: number;
  applications: DetectedApp[];
  lastUpdated: string;
}

/**
 * Device ownership information
 */
export interface DeviceOwnership {
  deviceId: string;
  deviceName: string;
  ownerType: string; // 'company', 'personal', 'unknown'
  userPrincipalName?: string;
  userDisplayName?: string;
  enrollmentType?: string;
  managementAgent: string;
  isManaged: boolean;
  azureADRegistered?: boolean;
}

/**
 * Device health attestation state
 */
export interface DeviceHealthAttestation {
  deviceId: string;
  deviceName: string;
  lastUpdateDateTime?: string;

  // Security features
  bitLockerStatus?: string;
  bootManagerVersion?: string;
  codeIntegrityCheckVersion?: string;
  secureBoot?: string;

  // Windows-specific
  bootDebugging?: string;
  operatingSystemKernelDebugging?: string;
  codeIntegrity?: string;
  testSigning?: string;
  safeMode?: string;
  windowsPE?: string;
  earlyLaunchAntiMalwareDriverProtection?: string;
  virtualSecureMode?: string;
  pcrHashAlgorithm?: string;

  // Boot integrity
  bootAppSecurityVersion?: string;
  bootManagerSecurityVersion?: string;
  tpmVersion?: string;
  pcr0?: string;

  // Device Guard
  secureBootConfigurationPolicyFingerPrint?: string;
  codeIntegrityPolicy?: string;
  bootRevisionListInfo?: string;
  operatingSystemRevListInfo?: string;

  // Health status
  healthStatusMismatchInfo?: string;
  healthAttestationSupportedStatus?: string;

  // Attestation
  issuedDateTime?: string;
  attestationIdentityKey?: string;
  resetCount?: number;
  restartCount?: number;
  dataExcutionPolicy?: string;
  bitLockerRecoveryKey?: string;
  contentNamespaceUrl?: string;
  contentVersion?: string;
}

/**
 * OS Platform summary
 */
export interface OSPlatformSummary {
  platform: string;
  totalDevices: number;
  compliantDevices: number;
  nonCompliantDevices: number;
  compliancePercentage: number;
  versions: Record<string, number>;
  manufacturers: Record<string, number>;
  averageLastSync?: string;
  enrollmentTypes: Record<string, number>;
}

/**
 * Device inventory report options
 */
export interface DeviceInventoryOptions {
  includeHardware?: boolean;
  includeSoftware?: boolean;
  includeHealthAttestation?: boolean;
  platformFilter?: string; // 'Windows', 'iOS', 'Android', etc.
  ownershipFilter?: string; // 'company', 'personal'
  complianceFilter?: string; // 'compliant', 'noncompliant'
  maxDevicesForSoftware?: number; // Limit software inventory to N devices (can be slow)
  outputFormats?: Array<'json' | 'csv' | 'html'>;
  outputDirectory?: string;
}

/**
 * Complete device inventory report data
 */
export interface DeviceInventoryReportData {
  metadata: {
    reportName: string;
    generatedAt: string;
    generatedBy: string;
    totalDevices: number;
    options: DeviceInventoryOptions;
  };

  summary: {
    totalDevices: number;
    byOperatingSystem: Record<string, OSPlatformSummary>;
    byOwnership: Record<string, number>;
    byCompliance: Record<string, number>;
    byManagementAgent: Record<string, number>;
    byEnrollmentType: Record<string, number>;
    totalStorageGB?: number;
    averageStorageUsedPercentage?: number;
  };

  devices: ManagedDevice[];
  hardwareInventory?: DeviceHardwareInventory[];
  softwareInventory?: DeviceSoftwareInventory[];
  ownershipDetails?: DeviceOwnership[];
  healthAttestation?: DeviceHealthAttestation[];
}

// ============================================================================
// Device Inventory Report Class
// ============================================================================

/**
 * Comprehensive Device Inventory Report Generator
 *
 * Provides detailed inventory reporting for all managed devices in Microsoft Intune,
 * including hardware specifications, software inventory, ownership details, and health attestation.
 *
 * @example
 * ```typescript
 * const report = new DeviceInventoryReport(graphClient, config);
 * const data = await report.executeWithOptions({
 *   includeHardware: true,
 *   includeSoftware: true,
 *   platformFilter: 'Windows',
 *   outputFormats: ['json', 'html'],
 *   outputDirectory: './reports'
 * });
 * ```
 */
export class DeviceInventoryReport extends BaseReport {
  name = 'device-inventory-comprehensive';
  description = 'Comprehensive device inventory with hardware, software, and health data';
  category = 'Devices';
  enabled = true;

  /**
   * Execute the basic device inventory report (implements BaseReport)
   */
  async execute(): Promise<ReportData> {
    logger.info('Executing basic Device Inventory Report');

    try {
      const devices = await this.getAllManagedDevices();
      const summary = this.generateBasicSummary(devices);

      const data = devices.map(device => this.transformDeviceForBasicReport(device));

      return {
        metadata: this.createMetadata(this.name, data.length),
        data,
        summary,
      };
    } catch (error) {
      logger.error('Failed to execute basic Device Inventory Report', error);
      throw error;
    }
  }

  /**
   * Execute comprehensive device inventory report with options
   *
   * @param options - Report options for customizing data collection
   * @returns Complete device inventory report data
   */
  async executeWithOptions(options: DeviceInventoryOptions = {}): Promise<DeviceInventoryReportData> {
    logger.info('Executing comprehensive Device Inventory Report', { options });

    try {
      // Fetch all managed devices
      const devices = await this.getAllManagedDevices(options);
      logger.info(`Retrieved ${devices.length} devices`);

      // Build comprehensive report data
      const reportData: DeviceInventoryReportData = {
        metadata: {
          reportName: this.name,
          generatedAt: new Date().toISOString(),
          generatedBy: 'Intune Reporting Dashboard',
          totalDevices: devices.length,
          options,
        },
        summary: this.generateComprehensiveSummary(devices),
        devices,
      };

      // Add hardware inventory if requested
      if (options.includeHardware) {
        logger.info('Collecting hardware inventory...');
        reportData.hardwareInventory = this.extractHardwareInventory(devices);
        logger.info(`Hardware inventory collected for ${reportData.hardwareInventory.length} devices`);
      }

      // Add software inventory if requested
      if (options.includeSoftware) {
        logger.info('Collecting software inventory...');
        const maxDevices = options.maxDevicesForSoftware || 50;
        reportData.softwareInventory = await this.getDevicesSoftwareInventory(
          devices.slice(0, maxDevices)
        );
        logger.info(`Software inventory collected for ${reportData.softwareInventory.length} devices`);
      }

      // Add ownership details
      reportData.ownershipDetails = this.extractOwnershipDetails(devices);

      // Add health attestation if requested
      if (options.includeHealthAttestation) {
        logger.info('Collecting health attestation data...');
        reportData.healthAttestation = await this.getDevicesHealthAttestation(devices);
        logger.info(`Health attestation collected for ${reportData.healthAttestation.length} devices`);
      }

      // Export to requested formats
      if (options.outputFormats && options.outputDirectory) {
        await this.exportReport(reportData, options);
      }

      logger.info('Device Inventory Report completed successfully');
      return reportData;

    } catch (error) {
      logger.error('Failed to execute comprehensive Device Inventory Report', error);
      throw error;
    }
  }

  // ==========================================================================
  // Device Data Collection Methods
  // ==========================================================================

  /**
   * Get all managed devices from Intune with pagination
   *
   * @param options - Filter options
   * @returns Array of managed devices
   */
  async getAllManagedDevices(options: DeviceInventoryOptions = {}): Promise<ManagedDevice[]> {
    try {
      let query = this.graphClient
        .api('/deviceManagement/managedDevices')
        .select([
          'id',
          'deviceName',
          'managedDeviceName',
          'operatingSystem',
          'osVersion',
          'manufacturer',
          'model',
          'serialNumber',
          'imei',
          'meid',
          'wiFiMacAddress',
          'ethernetMacAddress',
          'userPrincipalName',
          'userDisplayName',
          'userId',
          'emailAddress',
          'enrolledDateTime',
          'lastSyncDateTime',
          'complianceState',
          'managementAgent',
          'managementState',
          'deviceEnrollmentType',
          'deviceRegistrationState',
          'azureADDeviceId',
          'azureADRegistered',
          'deviceCategoryDisplayName',
          'totalStorageSpaceInBytes',
          'freeStorageSpaceInBytes',
          'physicalMemoryInBytes',
          'isEncrypted',
          'isSupervised',
          'exchangeAccessState',
          'exchangeAccessStateReason',
          'managedDeviceOwnerType',
          'partnerReportedThreatState',
          'jailBroken',
          'easActivated',
          'easActivationDateTime',
        ]);

      // Apply filters
      const filters: string[] = [];

      if (options.platformFilter) {
        filters.push(`operatingSystem eq '${options.platformFilter}'`);
      }

      if (options.ownershipFilter) {
        filters.push(`managedDeviceOwnerType eq '${options.ownershipFilter}'`);
      }

      if (options.complianceFilter) {
        filters.push(`complianceState eq '${options.complianceFilter}'`);
      }

      if (filters.length > 0) {
        query = query.filter(filters.join(' and '));
      }

      const response = await this.retryGraphCall(() => query.get());
      const devices = await this.getAllPages<ManagedDevice>(response);

      return devices;

    } catch (error) {
      logger.error('Failed to fetch managed devices', error);
      throw new Error(`Failed to fetch managed devices: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Get detailed device information by ID
   *
   * @param deviceId - Device ID
   * @returns Detailed device information
   */
  async getDeviceById(deviceId: string): Promise<ManagedDevice> {
    try {
      const device = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}`)
          .get()
      );

      return device as ManagedDevice;

    } catch (error) {
      logger.error(`Failed to fetch device ${deviceId}`, error);
      throw new Error(`Failed to fetch device: ${this.getErrorMessage(error)}`);
    }
  }

  /**
   * Get detected applications for a device
   *
   * @param deviceId - Device ID
   * @returns Array of detected applications
   */
  async getDeviceDetectedApps(deviceId: string): Promise<DetectedApp[]> {
    try {
      const response = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}/detectedApps`)
          .get()
      );

      const apps = await this.getAllPages<DetectedApp>(response);
      return apps;

    } catch (error) {
      logger.warn(`Failed to fetch apps for device ${deviceId}`, error);
      // Return empty array on error to allow report to continue
      return [];
    }
  }

  /**
   * Get device health attestation state
   *
   * @param deviceId - Device ID
   * @returns Device health attestation data
   */
  async getDeviceHealthAttestationState(deviceId: string): Promise<any> {
    try {
      const healthState = await this.retryGraphCall(() =>
        this.graphClient
          .api(`/deviceManagement/managedDevices/${deviceId}/deviceHealthAttestationState`)
          .get()
      );

      return healthState;

    } catch (error) {
      logger.warn(`Failed to fetch health attestation for device ${deviceId}`, error);
      // Return null on error to allow report to continue
      return null;
    }
  }

  // ==========================================================================
  // Hardware Inventory Methods
  // ==========================================================================

  /**
   * Extract hardware inventory from device data
   *
   * @param devices - Array of managed devices
   * @returns Array of hardware inventory records
   */
  extractHardwareInventory(devices: ManagedDevice[]): DeviceHardwareInventory[] {
    return devices.map(device => {
      const totalStorageGB = device.totalStorageSpaceInBytes
        ? Math.round(device.totalStorageSpaceInBytes / (1024 ** 3) * 100) / 100
        : undefined;

      const freeStorageGB = device.freeStorageSpaceInBytes
        ? Math.round(device.freeStorageSpaceInBytes / (1024 ** 3) * 100) / 100
        : undefined;

      const usedStoragePercentage = (totalStorageGB && freeStorageGB)
        ? Math.round(((totalStorageGB - freeStorageGB) / totalStorageGB) * 100 * 100) / 100
        : undefined;

      const totalMemoryGB = device.physicalMemoryInBytes
        ? Math.round(device.physicalMemoryInBytes / (1024 ** 3) * 100) / 100
        : undefined;

      return {
        deviceId: device.id,
        deviceName: device.deviceName,
        manufacturer: device.manufacturer,
        model: device.model,
        serialNumber: device.serialNumber,
        totalStorageGB,
        freeStorageGB,
        usedStoragePercentage,
        totalMemoryGB,
        wiFiMacAddress: device.wiFiMacAddress,
        ethernetMacAddress: device.ethernetMacAddress,
        imei: device.imei,
        meid: device.meid,
        operatingSystem: device.operatingSystem,
        osVersion: device.osVersion,
        isEncrypted: device.isEncrypted || false,
        isSupervised: device.isSupervised || false,
      };
    });
  }

  /**
   * Get hardware inventory for specific devices
   *
   * @param deviceIds - Array of device IDs
   * @returns Array of hardware inventory records
   */
  async getHardwareInventoryByDeviceIds(deviceIds: string[]): Promise<DeviceHardwareInventory[]> {
    const devices = await Promise.all(
      deviceIds.map(id => this.getDeviceById(id).catch(() => null))
    );

    const validDevices = devices.filter((d): d is ManagedDevice => d !== null);
    return this.extractHardwareInventory(validDevices);
  }

  // ==========================================================================
  // Software Inventory Methods
  // ==========================================================================

  /**
   * Get software inventory for multiple devices
   *
   * @param devices - Array of managed devices
   * @returns Array of software inventory records
   */
  async getDevicesSoftwareInventory(devices: ManagedDevice[]): Promise<DeviceSoftwareInventory[]> {
    const inventoryPromises = devices.map(async device => {
      try {
        const apps = await this.getDeviceDetectedApps(device.id);

        return {
          deviceId: device.id,
          deviceName: device.deviceName,
          operatingSystem: device.operatingSystem,
          osVersion: device.osVersion,
          totalApps: apps.length,
          applications: apps,
          lastUpdated: new Date().toISOString(),
        };
      } catch (error) {
        logger.warn(`Failed to get software inventory for device ${device.id}`, error);
        return null;
      }
    });

    const results = await Promise.all(inventoryPromises);
    return results.filter((r): r is DeviceSoftwareInventory => r !== null);
  }

  /**
   * Get software inventory for a single device
   *
   * @param deviceId - Device ID
   * @returns Software inventory for the device
   */
  async getDeviceSoftwareInventory(deviceId: string): Promise<DeviceSoftwareInventory | null> {
    try {
      const device = await this.getDeviceById(deviceId);
      const apps = await this.getDeviceDetectedApps(deviceId);

      return {
        deviceId: device.id,
        deviceName: device.deviceName,
        operatingSystem: device.operatingSystem,
        osVersion: device.osVersion,
        totalApps: apps.length,
        applications: apps,
        lastUpdated: new Date().toISOString(),
      };
    } catch (error) {
      logger.error(`Failed to get software inventory for device ${deviceId}`, error);
      return null;
    }
  }

  // ==========================================================================
  // Ownership and Management Methods
  // ==========================================================================

  /**
   * Extract ownership details from device data
   *
   * @param devices - Array of managed devices
   * @returns Array of ownership records
   */
  extractOwnershipDetails(devices: ManagedDevice[]): DeviceOwnership[] {
    return devices.map(device => ({
      deviceId: device.id,
      deviceName: device.deviceName,
      ownerType: device.managedDeviceOwnerType || 'unknown',
      userPrincipalName: device.userPrincipalName,
      userDisplayName: device.userDisplayName,
      enrollmentType: device.deviceEnrollmentType,
      managementAgent: device.managementAgent,
      isManaged: device.managementAgent !== 'unknown',
      azureADRegistered: device.azureADRegistered,
    }));
  }

  /**
   * Get devices by ownership type
   *
   * @param ownershipType - Ownership type ('company', 'personal', 'unknown')
   * @returns Array of devices with specified ownership
   */
  async getDevicesByOwnershipType(ownershipType: string): Promise<ManagedDevice[]> {
    return this.getAllManagedDevices({ ownershipFilter: ownershipType });
  }

  /**
   * Get devices by OS platform
   *
   * @param platform - OS platform ('Windows', 'iOS', 'Android', etc.)
   * @returns Array of devices on specified platform
   */
  async getDevicesByPlatform(platform: string): Promise<ManagedDevice[]> {
    return this.getAllManagedDevices({ platformFilter: platform });
  }

  /**
   * Get devices by compliance state
   *
   * @param complianceState - Compliance state ('compliant', 'noncompliant', etc.)
   * @returns Array of devices with specified compliance state
   */
  async getDevicesByComplianceState(complianceState: string): Promise<ManagedDevice[]> {
    return this.getAllManagedDevices({ complianceFilter: complianceState });
  }

  // ==========================================================================
  // Health Attestation Methods
  // ==========================================================================

  /**
   * Get health attestation data for multiple devices
   *
   * @param devices - Array of managed devices
   * @returns Array of health attestation records
   */
  async getDevicesHealthAttestation(devices: ManagedDevice[]): Promise<DeviceHealthAttestation[]> {
    // Filter to Windows devices (health attestation is primarily for Windows)
    const windowsDevices = devices.filter(d =>
      d.operatingSystem?.toLowerCase().includes('windows')
    );

    const attestationPromises = windowsDevices.map(async device => {
      try {
        const healthState = await this.getDeviceHealthAttestationState(device.id);

        if (!healthState) {
          return null;
        }

        return {
          deviceId: device.id,
          deviceName: device.deviceName,
          lastUpdateDateTime: healthState.lastUpdateDateTime,
          bitLockerStatus: healthState.bitLockerStatus,
          bootManagerVersion: healthState.bootManagerVersion,
          codeIntegrityCheckVersion: healthState.codeIntegrityCheckVersion,
          secureBoot: healthState.secureBoot,
          bootDebugging: healthState.bootDebugging,
          operatingSystemKernelDebugging: healthState.operatingSystemKernelDebugging,
          codeIntegrity: healthState.codeIntegrity,
          testSigning: healthState.testSigning,
          safeMode: healthState.safeMode,
          windowsPE: healthState.windowsPE,
          earlyLaunchAntiMalwareDriverProtection: healthState.earlyLaunchAntiMalwareDriverProtection,
          virtualSecureMode: healthState.virtualSecureMode,
          pcrHashAlgorithm: healthState.pcrHashAlgorithm,
          bootAppSecurityVersion: healthState.bootAppSecurityVersion,
          bootManagerSecurityVersion: healthState.bootManagerSecurityVersion,
          tpmVersion: healthState.tpmVersion,
          pcr0: healthState.pcr0,
          secureBootConfigurationPolicyFingerPrint: healthState.secureBootConfigurationPolicyFingerPrint,
          codeIntegrityPolicy: healthState.codeIntegrityPolicy,
          bootRevisionListInfo: healthState.bootRevisionListInfo,
          operatingSystemRevListInfo: healthState.operatingSystemRevListInfo,
          healthStatusMismatchInfo: healthState.healthStatusMismatchInfo,
          healthAttestationSupportedStatus: healthState.healthAttestationSupportedStatus,
          issuedDateTime: healthState.issuedDateTime,
          attestationIdentityKey: healthState.attestationIdentityKey,
          resetCount: healthState.resetCount,
          restartCount: healthState.restartCount,
          dataExcutionPolicy: healthState.dataExcutionPolicy,
          contentNamespaceUrl: healthState.contentNamespaceUrl,
          contentVersion: healthState.contentVersion,
        } as DeviceHealthAttestation;
      } catch (error) {
        logger.warn(`Failed to get health attestation for device ${device.id}`, error);
        return null;
      }
    });

    const results = await Promise.all(attestationPromises);
    return results.filter((r): r is DeviceHealthAttestation => r !== null);
  }

  // ==========================================================================
  // Summary and Analysis Methods
  // ==========================================================================

  /**
   * Generate basic summary for standard report
   *
   * @param devices - Array of managed devices
   * @returns Basic summary object
   */
  private generateBasicSummary(devices: ManagedDevice[]): Record<string, any> {
    return {
      totalDevices: devices.length,
      byOS: this.groupBy(devices, 'operatingSystem'),
      byCompliance: this.groupBy(devices, 'complianceState'),
      byManagementAgent: this.groupBy(devices, 'managementAgent'),
      byOwnership: this.groupBy(devices, 'managedDeviceOwnerType'),
    };
  }

  /**
   * Generate comprehensive summary with detailed statistics
   *
   * @param devices - Array of managed devices
   * @returns Comprehensive summary object
   */
  private generateComprehensiveSummary(devices: ManagedDevice[]): DeviceInventoryReportData['summary'] {
    const summary: DeviceInventoryReportData['summary'] = {
      totalDevices: devices.length,
      byOperatingSystem: this.generateOSPlatformSummaries(devices),
      byOwnership: this.groupBy(devices, 'managedDeviceOwnerType'),
      byCompliance: this.groupBy(devices, 'complianceState'),
      byManagementAgent: this.groupBy(devices, 'managementAgent'),
      byEnrollmentType: this.groupBy(devices, 'deviceEnrollmentType'),
    };

    // Calculate storage statistics
    const devicesWithStorage = devices.filter(d => d.totalStorageSpaceInBytes);
    if (devicesWithStorage.length > 0) {
      summary.totalStorageGB = Math.round(
        devicesWithStorage.reduce((sum, d) => sum + (d.totalStorageSpaceInBytes || 0), 0) / (1024 ** 3)
      );

      const usedStorages = devicesWithStorage.map(d => {
        const total = d.totalStorageSpaceInBytes || 0;
        const free = d.freeStorageSpaceInBytes || 0;
        return ((total - free) / total) * 100;
      });

      summary.averageStorageUsedPercentage = Math.round(
        usedStorages.reduce((sum, val) => sum + val, 0) / usedStorages.length * 100
      ) / 100;
    }

    return summary;
  }

  /**
   * Generate OS platform summaries with detailed statistics
   *
   * @param devices - Array of managed devices
   * @returns OS platform summaries by platform name
   */
  private generateOSPlatformSummaries(devices: ManagedDevice[]): Record<string, OSPlatformSummary> {
    const summaries: Record<string, OSPlatformSummary> = {};

    // Group devices by OS
    const devicesByOS = devices.reduce((acc, device) => {
      const os = device.operatingSystem || 'Unknown';
      if (!acc[os]) {
        acc[os] = [];
      }
      acc[os].push(device);
      return acc;
    }, {} as Record<string, ManagedDevice[]>);

    // Generate summary for each OS
    Object.entries(devicesByOS).forEach(([os, osDevices]) => {
      const compliantCount = osDevices.filter(d =>
        d.complianceState?.toLowerCase() === 'compliant'
      ).length;

      const nonCompliantCount = osDevices.filter(d =>
        d.complianceState?.toLowerCase() === 'noncompliant'
      ).length;

      summaries[os] = {
        platform: os,
        totalDevices: osDevices.length,
        compliantDevices: compliantCount,
        nonCompliantDevices: nonCompliantCount,
        compliancePercentage: osDevices.length > 0
          ? Math.round((compliantCount / osDevices.length) * 100 * 100) / 100
          : 0,
        versions: this.groupBy(osDevices, 'osVersion'),
        manufacturers: this.groupBy(osDevices, 'manufacturer'),
        enrollmentTypes: this.groupBy(osDevices, 'deviceEnrollmentType'),
      };
    });

    return summaries;
  }

  /**
   * Group devices by a property
   *
   * @param devices - Array of managed devices
   * @param key - Property key to group by
   * @returns Grouped counts
   */
  private groupBy(devices: any[], key: string): Record<string, number> {
    return devices.reduce((acc, device) => {
      const value = device[key] || 'Unknown';
      acc[value] = (acc[value] || 0) + 1;
      return acc;
    }, {});
  }

  /**
   * Transform device data for basic report output
   *
   * @param device - Managed device
   * @returns Transformed device data
   */
  private transformDeviceForBasicReport(device: ManagedDevice): any {
    return {
      deviceName: device.deviceName || 'Unknown',
      operatingSystem: device.operatingSystem || 'Unknown',
      osVersion: device.osVersion || 'Unknown',
      manufacturer: device.manufacturer || 'Unknown',
      model: device.model || 'Unknown',
      serialNumber: device.serialNumber || 'N/A',
      userPrincipalName: device.userPrincipalName || 'N/A',
      userDisplayName: device.userDisplayName || 'N/A',
      enrolledDate: device.enrolledDateTime
        ? new Date(device.enrolledDateTime).toLocaleDateString()
        : 'N/A',
      lastSync: device.lastSyncDateTime
        ? new Date(device.lastSyncDateTime).toLocaleDateString()
        : 'Never',
      complianceState: device.complianceState || 'Unknown',
      managementAgent: device.managementAgent || 'Unknown',
      ownerType: device.managedDeviceOwnerType || 'Unknown',
      enrollmentType: device.deviceEnrollmentType || 'Unknown',
    };
  }

  // ==========================================================================
  // Export Methods
  // ==========================================================================

  /**
   * Export report to multiple formats
   *
   * @param reportData - Complete report data
   * @param options - Export options
   */
  private async exportReport(
    reportData: DeviceInventoryReportData,
    options: DeviceInventoryOptions
  ): Promise<void> {
    if (!options.outputDirectory || !options.outputFormats) {
      return;
    }

    logger.info('Exporting device inventory report', {
      formats: options.outputFormats,
      directory: options.outputDirectory
    });

    for (const format of options.outputFormats) {
      try {
        const reportForExport: ReportData = {
          metadata: {
            reportName: reportData.metadata.reportName,
            generatedAt: reportData.metadata.generatedAt,
            generatedBy: reportData.metadata.generatedBy,
            recordCount: reportData.metadata.totalDevices,
            parameters: reportData.metadata.options,
          },
          data: this.prepareDataForExport(reportData),
          summary: reportData.summary,
        };

        const outputPath = await OutputFormatter.format(
          reportForExport,
          format,
          options.outputDirectory,
          true
        );

        logger.info(`Report exported to ${format.toUpperCase()}`, { path: outputPath });
      } catch (error) {
        logger.error(`Failed to export report to ${format}`, error);
      }
    }
  }

  /**
   * Prepare report data for export
   *
   * @param reportData - Complete report data
   * @returns Flattened data array for export
   */
  private prepareDataForExport(reportData: DeviceInventoryReportData): any[] {
    return reportData.devices.map(device => {
      const baseData: any = {
        deviceId: device.id,
        deviceName: device.deviceName,
        operatingSystem: device.operatingSystem,
        osVersion: device.osVersion,
        manufacturer: device.manufacturer || 'N/A',
        model: device.model || 'N/A',
        serialNumber: device.serialNumber || 'N/A',
        userPrincipalName: device.userPrincipalName || 'N/A',
        userDisplayName: device.userDisplayName || 'N/A',
        complianceState: device.complianceState,
        managementAgent: device.managementAgent,
        ownerType: device.managedDeviceOwnerType || 'Unknown',
        enrollmentType: device.deviceEnrollmentType || 'Unknown',
        enrolledDate: device.enrolledDateTime
          ? new Date(device.enrolledDateTime).toLocaleDateString()
          : 'N/A',
        lastSyncDate: device.lastSyncDateTime
          ? new Date(device.lastSyncDateTime).toLocaleDateString()
          : 'Never',
      };

      // Add hardware info if available
      if (reportData.hardwareInventory) {
        const hardware = reportData.hardwareInventory.find(h => h.deviceId === device.id);
        if (hardware) {
          baseData.totalStorageGB = hardware.totalStorageGB || 'N/A';
          baseData.freeStorageGB = hardware.freeStorageGB || 'N/A';
          baseData.storageUsed = hardware.usedStoragePercentage
            ? `${hardware.usedStoragePercentage}%`
            : 'N/A';
          baseData.totalMemoryGB = hardware.totalMemoryGB || 'N/A';
          baseData.isEncrypted = hardware.isEncrypted ? 'Yes' : 'No';
        }
      }

      // Add software info if available
      if (reportData.softwareInventory) {
        const software = reportData.softwareInventory.find(s => s.deviceId === device.id);
        if (software) {
          baseData.totalApps = software.totalApps;
        }
      }

      return baseData;
    });
  }

  // ==========================================================================
  // Utility Methods
  // ==========================================================================

  /**
   * Extract error message from various error types
   *
   * @param error - Error object
   * @returns Error message string
   */
  private getErrorMessage(error: unknown): string {
    if (error instanceof Error) {
      return error.message;
    }
    if (typeof error === 'string') {
      return error;
    }
    if (error && typeof error === 'object' && 'message' in error) {
      return String(error.message);
    }
    return 'Unknown error occurred';
  }

  /**
   * Print device inventory summary to console
   *
   * @param reportData - Report data to summarize
   */
  static printSummary(reportData: DeviceInventoryReportData): void {
    console.log('\n' + '='.repeat(70));
    console.log('DEVICE INVENTORY REPORT SUMMARY');
    console.log('='.repeat(70));
    console.log(`\nGenerated: ${new Date(reportData.metadata.generatedAt).toLocaleString()}`);
    console.log(`Total Devices: ${reportData.metadata.totalDevices}`);

    console.log('\n--- By Operating System ---');
    Object.entries(reportData.summary.byOperatingSystem).forEach(([os, summary]) => {
      console.log(`\n${os}:`);
      console.log(`  Total: ${summary.totalDevices}`);
      console.log(`  Compliant: ${summary.compliantDevices} (${summary.compliancePercentage}%)`);
      console.log(`  Non-Compliant: ${summary.nonCompliantDevices}`);
      console.log(`  Top Versions: ${Object.entries(summary.versions).slice(0, 3).map(([v, c]) => `${v} (${c})`).join(', ')}`);
    });

    console.log('\n--- By Ownership ---');
    Object.entries(reportData.summary.byOwnership).forEach(([type, count]) => {
      console.log(`  ${type}: ${count}`);
    });

    console.log('\n--- By Compliance ---');
    Object.entries(reportData.summary.byCompliance).forEach(([state, count]) => {
      console.log(`  ${state}: ${count}`);
    });

    if (reportData.hardwareInventory) {
      console.log(`\n--- Hardware Inventory ---`);
      console.log(`  Devices with hardware data: ${reportData.hardwareInventory.length}`);
      if (reportData.summary.totalStorageGB) {
        console.log(`  Total storage: ${reportData.summary.totalStorageGB} GB`);
      }
      if (reportData.summary.averageStorageUsedPercentage) {
        console.log(`  Average storage used: ${reportData.summary.averageStorageUsedPercentage}%`);
      }
    }

    if (reportData.softwareInventory) {
      console.log(`\n--- Software Inventory ---`);
      console.log(`  Devices with software data: ${reportData.softwareInventory.length}`);
      const totalApps = reportData.softwareInventory.reduce((sum, si) => sum + si.totalApps, 0);
      console.log(`  Total applications tracked: ${totalApps}`);
    }

    if (reportData.healthAttestation) {
      console.log(`\n--- Health Attestation ---`);
      console.log(`  Devices with health data: ${reportData.healthAttestation.length}`);
    }

    console.log('\n' + '='.repeat(70) + '\n');
  }
}

// ============================================================================
// Convenience Functions
// ============================================================================

/**
 * Quick function to generate a basic device inventory report
 *
 * @param graphClient - Microsoft Graph client
 * @param config - App configuration
 * @returns Basic report data
 */
export async function generateBasicDeviceInventory(
  graphClient: Client,
  config: AppConfig
): Promise<ReportData> {
  const report = new DeviceInventoryReport(graphClient, config);
  return report.execute();
}

/**
 * Quick function to generate a comprehensive device inventory report
 *
 * @param graphClient - Microsoft Graph client
 * @param config - App configuration
 * @param options - Report options
 * @returns Comprehensive report data
 */
export async function generateComprehensiveDeviceInventory(
  graphClient: Client,
  config: AppConfig,
  options: DeviceInventoryOptions
): Promise<DeviceInventoryReportData> {
  const report = new DeviceInventoryReport(graphClient, config);
  return report.executeWithOptions(options);
}

/**
 * Export all available report types and interfaces
 */
export {
  ManagedDevice,
  DeviceHardwareInventory,
  DetectedApp,
  DeviceSoftwareInventory,
  DeviceOwnership,
  DeviceHealthAttestation,
  OSPlatformSummary,
  DeviceInventoryOptions,
  DeviceInventoryReportData,
};
