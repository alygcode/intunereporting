# Device Inventory Report - Quick Start Guide

## 5-Minute Setup

### 1. Configure Environment Variables

Create a `.env` file in the project root:

```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-client-secret-here
```

### 2. Install Dependencies

```bash
npm install
```

### 3. Run Your First Report

```typescript
// basic-report.ts
import { DeviceInventoryReport } from './src/reports/device-inventory-report';
import { createGraphClient, createConfig } from './utils';

async function run() {
  const graphClient = await createGraphClient();
  const config = createConfig();

  const report = new DeviceInventoryReport(graphClient, config);
  const data = await report.execute();

  console.log(`Total Devices: ${data.metadata.recordCount}`);
}

run();
```

## Common Use Cases

### Get All Windows Devices

```typescript
const report = new DeviceInventoryReport(graphClient, config);
const windowsDevices = await report.getDevicesByPlatform('Windows');
console.log(`Windows Devices: ${windowsDevices.length}`);
```

### Get Non-Compliant Devices

```typescript
const nonCompliant = await report.getDevicesByComplianceState('noncompliant');
console.log(`Non-Compliant Devices: ${nonCompliant.length}`);
```

### Generate Report with Hardware Details

```typescript
const reportData = await report.executeWithOptions({
  includeHardware: true,
  outputFormats: ['json', 'html'],
  outputDirectory: './reports'
});

DeviceInventoryReport.printSummary(reportData);
```

### Get Software Inventory

```typescript
const reportData = await report.executeWithOptions({
  includeSoftware: true,
  maxDevicesForSoftware: 25,
  outputFormats: ['json'],
  outputDirectory: './reports'
});

reportData.softwareInventory?.forEach(device => {
  console.log(`${device.deviceName}: ${device.totalApps} apps`);
});
```

## Run Example Scripts

```bash
# Run all examples
npm run example

# Run specific example (1-8)
npm run example 1

# Or directly with ts-node
ts-node src/reports/examples/device-inventory-report-example.ts
```

## Available Examples

1. **Basic Inventory** - Simple device list
2. **Comprehensive with Hardware** - Includes hardware specs
3. **Software Inventory** - Applications on devices
4. **Windows Devices Only** - Platform-specific report
5. **Corporate Devices** - Filter by ownership
6. **Non-Compliant Devices** - Compliance issues
7. **Complete Comprehensive** - All features enabled
8. **Individual Methods** - Use specific API methods

## Output Formats

### JSON (Structured Data)
```typescript
const options = {
  outputFormats: ['json'],
  outputDirectory: './reports'
};
```

### CSV (Spreadsheet-Ready)
```typescript
const options = {
  outputFormats: ['csv'],
  outputDirectory: './reports'
};
```

### HTML (Human-Readable)
```typescript
const options = {
  outputFormats: ['html'],
  outputDirectory: './reports'
};
```

### All Formats
```typescript
const options = {
  outputFormats: ['json', 'csv', 'html'],
  outputDirectory: './reports'
};
```

## Filtering Options

### By Operating System
```typescript
const options = {
  platformFilter: 'Windows'  // or 'iOS', 'Android', 'macOS'
};
```

### By Ownership
```typescript
const options = {
  ownershipFilter: 'company'  // or 'personal'
};
```

### By Compliance
```typescript
const options = {
  complianceFilter: 'noncompliant'  // or 'compliant', 'ingraceperiod'
};
```

### Combine Filters
```typescript
const options = {
  platformFilter: 'Windows',
  complianceFilter: 'noncompliant',
  includeHardware: true,
  outputFormats: ['html'],
  outputDirectory: './reports/windows-noncompliant'
};
```

## API Quick Reference

### Main Methods

```typescript
// Basic report
await report.execute()

// Comprehensive report with options
await report.executeWithOptions(options)

// Get all devices
await report.getAllManagedDevices()

// Get specific device
await report.getDeviceById(deviceId)

// Get devices by platform
await report.getDevicesByPlatform('Windows')

// Get devices by ownership
await report.getDevicesByOwnershipType('company')

// Get devices by compliance
await report.getDevicesByComplianceState('compliant')

// Get software for device
await report.getDeviceSoftwareInventory(deviceId)

// Get health attestation
await report.getDevicesHealthAttestation(devices)

// Extract hardware inventory
report.extractHardwareInventory(devices)

// Print summary to console
DeviceInventoryReport.printSummary(reportData)
```

## Troubleshooting

### "Authentication failed"
- Verify environment variables are set correctly
- Check Azure AD app has required permissions
- Ensure admin consent is granted

### "Rate limited"
- Module has automatic retry logic
- Reduce `maxDevicesForSoftware` if collecting software inventory
- Add delays between report runs

### "No health attestation data"
- Health attestation only works for Windows 10/11 devices
- Ensure devices have health attestation enabled

## Performance Tips

1. **Use filters** to reduce dataset size
2. **Limit software inventory** with `maxDevicesForSoftware`
3. **Run during off-peak hours** for large inventories
4. **Export only needed formats** to save time

## Next Steps

- Review the full documentation: `docs/DEVICE_INVENTORY_REPORT.md`
- Explore example scripts: `src/reports/examples/device-inventory-report-example.ts`
- Check the module source: `src/reports/device-inventory-report.ts`

## Support

For detailed documentation, see:
- [Full Documentation](./DEVICE_INVENTORY_REPORT.md)
- [Microsoft Graph API Docs](https://docs.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice)
