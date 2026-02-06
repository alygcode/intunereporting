# Device Inventory Report Module

Comprehensive device inventory reporting module for Microsoft Intune using the Microsoft Graph API.

## Overview

The Device Inventory Report module provides detailed insights into all managed devices in your Microsoft Intune environment. It goes beyond basic device information to include hardware specifications, software inventory, ownership details, and health attestation data.

## Features

- **Complete Device Inventory**: Retrieve all managed devices with comprehensive metadata
- **Hardware Inventory**: Detailed hardware specifications including storage, memory, and network information
- **Software Inventory**: Detect and catalog installed applications on devices
- **Ownership Management**: Track device ownership types (corporate vs. personal)
- **Platform Filtering**: Filter and group devices by operating system
- **Health Attestation**: Windows device health and security status (BitLocker, Secure Boot, TPM, etc.)
- **Compliance Tracking**: Monitor device compliance states
- **Advanced Filtering**: Filter by OS platform, ownership type, or compliance state
- **Multiple Export Formats**: Export to JSON, CSV, and HTML
- **Error Handling**: Comprehensive error handling with automatic retry logic
- **Pagination Support**: Handle large device inventories efficiently

## Prerequisites

### Azure AD App Registration

You need an Azure AD app registration with the following Microsoft Graph API permissions:

**Required Permissions:**
- `DeviceManagementManagedDevices.Read.All` - Read managed device information
- `DeviceManagementApps.Read.All` - Read application inventory
- `DeviceManagementConfiguration.Read.All` - Read device configuration and health data

**Permission Type:** Application permissions (for service/daemon scenarios)

### Environment Setup

Create a `.env` file with your Azure AD credentials:

```env
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
```

## Installation

The module is part of the Intune Reporting Dashboard project:

```bash
npm install
```

## Quick Start

### Basic Device Inventory

```typescript
import { DeviceInventoryReport } from './reports/device-inventory-report';
import { createGraphClient, createConfig } from './utils';

const graphClient = await createGraphClient();
const config = createConfig();

const report = new DeviceInventoryReport(graphClient, config);
const data = await report.execute();

console.log(`Total Devices: ${data.metadata.recordCount}`);
```

### Comprehensive Inventory with All Features

```typescript
const options = {
  includeHardware: true,
  includeSoftware: true,
  includeHealthAttestation: true,
  maxDevicesForSoftware: 50,
  outputFormats: ['json', 'html', 'csv'],
  outputDirectory: './reports'
};

const reportData = await report.executeWithOptions(options);
DeviceInventoryReport.printSummary(reportData);
```

## Usage Examples

### Example 1: Windows Devices Only

```typescript
const options = {
  platformFilter: 'Windows',
  includeHardware: true,
  includeHealthAttestation: true,
  outputFormats: ['json'],
  outputDirectory: './reports/windows'
};

const reportData = await report.executeWithOptions(options);
```

### Example 2: Corporate-Owned Devices

```typescript
const options = {
  ownershipFilter: 'company',
  includeHardware: true,
  outputFormats: ['csv', 'html'],
  outputDirectory: './reports/corporate'
};

const reportData = await report.executeWithOptions(options);
```

### Example 3: Non-Compliant Devices

```typescript
const options = {
  complianceFilter: 'noncompliant',
  outputFormats: ['json', 'html'],
  outputDirectory: './reports/non-compliant'
};

const reportData = await report.executeWithOptions(options);
```

### Example 4: Software Inventory

```typescript
const options = {
  includeSoftware: true,
  maxDevicesForSoftware: 25,
  outputFormats: ['json'],
  outputDirectory: './reports/software'
};

const reportData = await report.executeWithOptions(options);

// Access software inventory
reportData.softwareInventory?.forEach(device => {
  console.log(`${device.deviceName}: ${device.totalApps} apps`);
  device.applications.forEach(app => {
    console.log(`  - ${app.displayName} (${app.version})`);
  });
});
```

### Example 5: Using Individual Methods

```typescript
const report = new DeviceInventoryReport(graphClient, config);

// Get all devices
const allDevices = await report.getAllManagedDevices();

// Get specific device
const device = await report.getDeviceById(deviceId);

// Get devices by platform
const windowsDevices = await report.getDevicesByPlatform('Windows');
const iosDevices = await report.getDevicesByPlatform('iOS');

// Get devices by ownership
const corporateDevices = await report.getDevicesByOwnershipType('company');

// Get devices by compliance
const nonCompliantDevices = await report.getDevicesByComplianceState('noncompliant');

// Get software for a device
const software = await report.getDeviceSoftwareInventory(deviceId);

// Get health attestation for devices
const healthData = await report.getDevicesHealthAttestation(windowsDevices);
```

## API Reference

### DeviceInventoryReport Class

#### Methods

##### `execute(): Promise<ReportData>`
Execute basic device inventory report (implements BaseReport interface).

**Returns:** Basic report data with device list and summary

##### `executeWithOptions(options: DeviceInventoryOptions): Promise<DeviceInventoryReportData>`
Execute comprehensive device inventory with custom options.

**Parameters:**
- `options` - Configuration options for the report

**Returns:** Comprehensive report data with all requested components

##### `getAllManagedDevices(options?: DeviceInventoryOptions): Promise<ManagedDevice[]>`
Get all managed devices with optional filtering.

**Parameters:**
- `options.platformFilter` - Filter by OS platform
- `options.ownershipFilter` - Filter by ownership type
- `options.complianceFilter` - Filter by compliance state

**Returns:** Array of managed devices

##### `getDeviceById(deviceId: string): Promise<ManagedDevice>`
Get detailed information for a specific device.

**Parameters:**
- `deviceId` - Device ID

**Returns:** Device details

##### `getDeviceDetectedApps(deviceId: string): Promise<DetectedApp[]>`
Get detected applications for a device.

**Parameters:**
- `deviceId` - Device ID

**Returns:** Array of detected applications

##### `getDeviceHealthAttestationState(deviceId: string): Promise<any>`
Get health attestation state for a device (Windows only).

**Parameters:**
- `deviceId` - Device ID

**Returns:** Health attestation data

##### `getDevicesByPlatform(platform: string): Promise<ManagedDevice[]>`
Get devices filtered by OS platform.

**Parameters:**
- `platform` - OS platform name ('Windows', 'iOS', 'Android', etc.)

**Returns:** Array of devices

##### `getDevicesByOwnershipType(ownershipType: string): Promise<ManagedDevice[]>`
Get devices filtered by ownership type.

**Parameters:**
- `ownershipType` - Ownership type ('company', 'personal', 'unknown')

**Returns:** Array of devices

##### `getDevicesByComplianceState(complianceState: string): Promise<ManagedDevice[]>`
Get devices filtered by compliance state.

**Parameters:**
- `complianceState` - Compliance state ('compliant', 'noncompliant', etc.)

**Returns:** Array of devices

##### `getDeviceSoftwareInventory(deviceId: string): Promise<DeviceSoftwareInventory | null>`
Get software inventory for a single device.

**Parameters:**
- `deviceId` - Device ID

**Returns:** Software inventory or null

##### `extractHardwareInventory(devices: ManagedDevice[]): DeviceHardwareInventory[]`
Extract hardware inventory from device data.

**Parameters:**
- `devices` - Array of managed devices

**Returns:** Array of hardware inventory records

##### `static printSummary(reportData: DeviceInventoryReportData): void`
Print formatted summary to console.

**Parameters:**
- `reportData` - Report data to summarize

### Types

#### DeviceInventoryOptions

```typescript
interface DeviceInventoryOptions {
  includeHardware?: boolean;              // Include hardware inventory
  includeSoftware?: boolean;              // Include software inventory
  includeHealthAttestation?: boolean;     // Include health attestation (Windows)
  platformFilter?: string;                // Filter by OS platform
  ownershipFilter?: string;               // Filter by ownership ('company', 'personal')
  complianceFilter?: string;              // Filter by compliance state
  maxDevicesForSoftware?: number;         // Limit software inventory (performance)
  outputFormats?: Array<'json' | 'csv' | 'html'>;  // Export formats
  outputDirectory?: string;               // Output directory for exports
}
```

#### ManagedDevice

```typescript
interface ManagedDevice {
  id: string;
  deviceName: string;
  operatingSystem: string;
  osVersion: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  userPrincipalName?: string;
  userDisplayName?: string;
  enrolledDateTime?: string;
  lastSyncDateTime?: string;
  complianceState: string;
  managementAgent: string;
  managedDeviceOwnerType?: string;
  deviceEnrollmentType?: string;
  totalStorageSpaceInBytes?: number;
  freeStorageSpaceInBytes?: number;
  physicalMemoryInBytes?: number;
  isEncrypted?: boolean;
  // ... and many more properties
}
```

#### DeviceHardwareInventory

```typescript
interface DeviceHardwareInventory {
  deviceId: string;
  deviceName: string;
  manufacturer?: string;
  model?: string;
  serialNumber?: string;
  totalStorageGB?: number;
  freeStorageGB?: number;
  usedStoragePercentage?: number;
  totalMemoryGB?: number;
  wiFiMacAddress?: string;
  ethernetMacAddress?: string;
  operatingSystem: string;
  osVersion: string;
  isEncrypted: boolean;
  isSupervised: boolean;
}
```

#### DeviceSoftwareInventory

```typescript
interface DeviceSoftwareInventory {
  deviceId: string;
  deviceName: string;
  operatingSystem: string;
  osVersion: string;
  totalApps: number;
  applications: DetectedApp[];
  lastUpdated: string;
}
```

#### DeviceHealthAttestation

```typescript
interface DeviceHealthAttestation {
  deviceId: string;
  deviceName: string;
  lastUpdateDateTime?: string;
  bitLockerStatus?: string;
  secureBoot?: string;
  codeIntegrity?: string;
  tpmVersion?: string;
  virtualSecureMode?: string;
  bootDebugging?: string;
  testSigning?: string;
  // ... and many more security properties
}
```

## Graph API Endpoints Used

The module uses the following Microsoft Graph API endpoints:

- `GET /deviceManagement/managedDevices` - List all managed devices
- `GET /deviceManagement/managedDevices/{id}` - Get device details
- `GET /deviceManagement/managedDevices/{id}/detectedApps` - Get detected applications
- `GET /deviceManagement/managedDevices/{id}/deviceHealthAttestationState` - Get health attestation

## Performance Considerations

### Software Inventory

Software inventory can be time-consuming for large device counts because it requires individual API calls per device. Use the `maxDevicesForSoftware` option to limit the number of devices:

```typescript
const options = {
  includeSoftware: true,
  maxDevicesForSoftware: 50  // Only collect software for first 50 devices
};
```

### Health Attestation

Health attestation is only applicable to Windows devices and requires additional API calls. Consider filtering to Windows devices when using this feature:

```typescript
const options = {
  platformFilter: 'Windows',
  includeHealthAttestation: true
};
```

### Pagination

The module automatically handles pagination for large device inventories. No additional configuration is needed.

### Retry Logic

Built-in retry logic with exponential backoff handles rate limiting and transient errors automatically.

## Error Handling

The module includes comprehensive error handling:

```typescript
try {
  const reportData = await report.executeWithOptions(options);
  console.log('Report generated successfully');
} catch (error) {
  console.error('Report generation failed:', error);
  // Handle error appropriately
}
```

Failed individual device operations (e.g., software inventory for one device) won't fail the entire report. They're logged as warnings and the report continues.

## Export Formats

### JSON
Complete structured data including all nested objects and arrays.

```typescript
const options = { outputFormats: ['json'] };
```

### CSV
Flattened tabular data suitable for spreadsheet applications.

```typescript
const options = { outputFormats: ['csv'] };
```

### HTML
Formatted report with styling, suitable for viewing in browsers.

```typescript
const options = { outputFormats: ['html'] };
```

### Multiple Formats
Export to multiple formats simultaneously:

```typescript
const options = { outputFormats: ['json', 'csv', 'html'] };
```

## Running Examples

The module includes comprehensive examples demonstrating various use cases:

```bash
# Run all examples
npm run example

# Run specific example (1-8)
npm run example 1

# Examples:
# 1. Basic Device Inventory
# 2. Comprehensive with Hardware Details
# 3. Software Inventory
# 4. Windows Devices Only
# 5. Corporate-Owned Devices
# 6. Non-Compliant Devices
# 7. Complete Comprehensive Report
# 8. Using Individual Methods
```

Or directly with ts-node:

```bash
ts-node src/reports/examples/device-inventory-report-example.ts
ts-node src/reports/examples/device-inventory-report-example.ts 3
```

## Best Practices

1. **Filter Early**: Use filters to reduce the dataset before collecting additional data
   ```typescript
   const options = {
     platformFilter: 'Windows',  // Filter first
     includeHealthAttestation: true  // Then collect extra data
   };
   ```

2. **Limit Software Inventory**: Set reasonable limits for software inventory
   ```typescript
   maxDevicesForSoftware: 50  // Don't process all devices if you have thousands
   ```

3. **Use Appropriate Export Formats**: Choose formats based on your needs
   - JSON: Full data, programmatic access
   - CSV: Data analysis in Excel
   - HTML: Human-readable reports

4. **Handle Errors Gracefully**: Always wrap report execution in try-catch blocks

5. **Schedule Wisely**: For automated reporting, schedule during off-peak hours

## Troubleshooting

### Authentication Errors

**Error:** "Authentication failed"

**Solution:** Verify your Azure AD credentials and permissions:
- Check AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
- Verify app has required Graph API permissions
- Ensure admin has granted consent for permissions

### Rate Limiting

**Error:** "429 Too Many Requests"

**Solution:** The module includes automatic retry logic with exponential backoff. If you still encounter issues:
- Reduce `maxDevicesForSoftware`
- Add delays between report runs
- Consider filtering to reduce API calls

### Large Device Inventories

**Issue:** Reports are slow with thousands of devices

**Solution:**
- Use filters to reduce device count
- Disable software inventory or set low `maxDevicesForSoftware`
- Run reports during off-peak hours
- Consider generating separate reports for different device groups

### Missing Health Attestation Data

**Issue:** No health attestation data returned

**Solution:**
- Health attestation only works for Windows 10/11 devices
- Devices must have health attestation enabled
- Some older Windows versions may not support all features
- Check device is properly enrolled and syncing

## Contributing

When contributing to this module:

1. Maintain TypeScript type safety
2. Include comprehensive error handling
3. Add JSDoc comments for new methods
4. Update documentation for new features
5. Add examples for new functionality
6. Follow existing code patterns
7. Test with various device types and scenarios

## Support

For issues, questions, or contributions:

1. Check existing documentation
2. Review example code
3. Search for similar issues
4. Create detailed bug reports with:
   - Environment details
   - Error messages
   - Steps to reproduce
   - Expected vs. actual behavior

## License

MIT License - See LICENSE file for details

## Related Documentation

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
- [Intune Device Management](https://docs.microsoft.com/en-us/mem/intune/fundamentals/what-is-device-management)
- [Azure AD App Registration](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [Graph API Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference)

## Changelog

### Version 1.0.0
- Initial release
- Complete device inventory
- Hardware inventory support
- Software inventory support
- Health attestation support
- Multiple export formats
- Comprehensive filtering options
- Error handling and retry logic
- Pagination support
