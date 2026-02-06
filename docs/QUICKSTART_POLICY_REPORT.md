# Policy Assignment Report - Quick Start Guide

## What Was Created

Three new files have been added to the Intune Reporting project:

1. **Main Module** (`src/reports/policy-assignment-report.ts`) - 38KB
   - Complete PolicyAssignmentReport class
   - All data structures and interfaces
   - Export methods (JSON, CSV, HTML)

2. **Examples** (`src/reports/examples/policy-assignment-report-example.ts`) - 18KB
   - 8 comprehensive usage examples
   - Real-world scenarios
   - Copy-paste ready code

3. **Documentation** (`docs/POLICY_ASSIGNMENT_REPORT.md`) - 16KB
   - Complete API reference
   - Usage guide
   - Troubleshooting tips

## 5-Minute Quick Start

### 1. Set Up Environment Variables

```bash
export AZURE_TENANT_ID="your-tenant-id"
export AZURE_CLIENT_ID="your-client-id"
export AZURE_CLIENT_SECRET="your-client-secret"
```

### 2. Run a Basic Report

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { PolicyAssignmentReport } from './reports/policy-assignment-report';

// Create authenticated Graph client (see examples for full code)
const graphClient = await createGraphClient();

// Create report instance
const report = new PolicyAssignmentReport(graphClient, config);

// Generate report
const result = await report.execute({
  includeDeviceStatus: true,
  includeUserStatus: true,
  includeConflicts: true,
  includeSuccessRates: true,
  outputDir: './reports'
});

// View summary
console.log(`Total Policies: ${result.summary.totalPolicies}`);
console.log(`Success Rate: ${result.summary.overallSuccessRate}%`);
console.log(`Conflicts: ${result.summary.totalConflicts}`);
```

### 3. Run the Examples

```bash
# Navigate to the project
cd /home/user/intunereporting

# Install dependencies (if needed)
npm install

# Run the comprehensive example
npx ts-node src/reports/examples/policy-assignment-report-example.ts
```

## Key Features

### What This Module Does

✅ **Fetches All Configuration Policies**
- Windows, iOS, Android, macOS policies
- Policy metadata and versions
- Creation and modification dates

✅ **Gets Policy Assignments**
- Group assignments
- User assignments
- Device assignments
- All Users/All Devices assignments

✅ **Tracks Deployment Status**
- Device-level status (success, error, conflict)
- User-level status
- Real-time compliance states

✅ **Detects Policy Conflicts**
- Identifies conflicting policies
- Shows affected devices
- Provides conflict details

✅ **Calculates Success/Failure Rates**
- Per-policy success rates
- Overall tenant success rate
- Detailed breakdown by status

✅ **Exports to Multiple Formats**
- JSON: Full data export
- CSV: Spreadsheet analysis
- HTML: Interactive web report

## Graph API Endpoints Used

```
GET /deviceManagement/deviceConfigurations
GET /deviceManagement/deviceConfigurations/{id}/assignments
GET /deviceManagement/deviceConfigurations/{id}/deviceStatuses
GET /deviceManagement/deviceConfigurations/{id}/userStatuses
GET /deviceManagement/deviceConfigurationConflictSummary
GET /groups/{id}
```

## Required Permissions

Add these to your Azure AD App Registration:

- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

## Common Use Cases

### 1. Weekly Policy Review

```typescript
// Generate comprehensive report every Monday
const result = await report.execute({
  outputDir: './reports/weekly',
  includeSuccessRates: true,
  includeConflicts: true
});
```

### 2. Troubleshoot Failed Deployments

```typescript
// Focus on deployment status
const result = await report.execute({
  includeDeviceStatus: true,
  includeSuccessRates: true
});

// Check for low success rates
const data = result.data[0];
const problematic = data.successRates.filter(r => r.successRate < 80);
```

### 3. Audit Policy Assignments

```typescript
// See which policies are assigned to which groups
const result = await report.execute({
  includeDeviceStatus: false,
  includeUserStatus: false
});

// Review assignments
const data = result.data[0];
console.log(`Total Assignments: ${data.assignments.length}`);
```

### 4. Identify Conflicts

```typescript
// Run conflict analysis
const result = await report.execute({
  includeConflicts: true,
  includeDeviceStatus: false
});

if (result.summary.totalConflicts > 0) {
  console.log('⚠️ Conflicts detected! Review the report.');
}
```

### 5. Platform-Specific Reports

```typescript
// Report only Windows policies
const result = await report.execute({
  platformFilter: 'Windows'
});
```

## Report Output Structure

### Summary Object

```typescript
{
  totalPolicies: 42,
  policiesWithAssignments: 38,
  policiesWithoutAssignments: 4,
  totalAssignments: 156,
  totalDeviceDeployments: 1234,
  totalUserDeployments: 567,
  totalConflicts: 2,
  overallSuccessRate: 94.5,
  overallFailureRate: 3.2,
  policiesByPlatform: {
    "Windows": 25,
    "iOS": 10,
    "Android": 7
  },
  generatedAt: "2024-02-04T12:00:00Z"
}
```

### Main Data Arrays

- `policies[]` - All configuration policies
- `assignments[]` - All policy assignments
- `deviceStatuses[]` - Device deployment statuses
- `userStatuses[]` - User deployment statuses
- `conflicts[]` - Policy conflicts
- `successRates[]` - Success/failure metrics

## Performance Tips

### For Large Tenants (100+ policies)

```typescript
// Start with minimal data
const result = await report.execute({
  includeDeviceStatus: false,  // Skip for faster execution
  includeUserStatus: false,
  includeConflicts: true,
  includeSuccessRates: true
});
```

### For Unreliable Networks

```typescript
// Increase retry attempts
const result = await report.execute({
  maxRetries: 10,  // Default is 3
  // Other options...
});
```

## Integration with Existing Reports

This module extends the existing Intune Reporting framework and works seamlessly with other reports:

- Uses the same `BaseReport` class
- Follows the same patterns as `device-compliance.ts`
- Compatible with the existing orchestrator
- Uses centralized logging
- Supports all output formats

## Next Steps

1. **Review the full documentation**: `docs/POLICY_ASSIGNMENT_REPORT.md`
2. **Try the examples**: `src/reports/examples/policy-assignment-report-example.ts`
3. **Integrate into your workflow**: Add to scheduler or run on-demand
4. **Customize as needed**: Extend the class for specific requirements

## Troubleshooting

### "No policies found"
- Verify Azure AD permissions
- Check that policies exist in your Intune tenant
- Review authentication credentials

### "Conflict data not available"
- Conflict API may require specific licenses
- Set `includeConflicts: false` to skip

### "Timeout errors"
- Increase `maxRetries` option
- Reduce data fetched (disable device/user statuses)
- Check network connection

## Getting Help

- Review the full documentation: `docs/POLICY_ASSIGNMENT_REPORT.md`
- Check the examples: `src/reports/examples/policy-assignment-report-example.ts`
- See the main README: `README.md`
- Review Microsoft Graph API docs

## File Locations

```
intunereporting/
├── src/
│   └── reports/
│       ├── policy-assignment-report.ts           # Main module
│       └── examples/
│           └── policy-assignment-report-example.ts  # Examples
└── docs/
    ├── POLICY_ASSIGNMENT_REPORT.md               # Full docs
    └── QUICKSTART_POLICY_REPORT.md               # This file
```

---

**Created**: 2024-02-04
**Version**: 1.0.0
**Module Size**: 38KB (main module)
