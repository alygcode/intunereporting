# Policy Assignment Report Module

Comprehensive reporting for Microsoft Intune configuration policies, including assignments, deployment status, conflicts, and success/failure metrics.

## Overview

The `PolicyAssignmentReport` class provides detailed insights into your Intune configuration policies, helping you understand:

- Which policies are deployed and to whom
- Policy deployment success and failure rates
- Device and user deployment statuses
- Policy conflicts
- Assignment details (groups, users, devices)

## Features

- **Complete Policy Inventory**: Fetch all configuration policies with metadata
- **Assignment Tracking**: Detailed assignment information including target groups
- **Deployment Status**: Device and user-level deployment status tracking
- **Conflict Detection**: Identify and report policy conflicts
- **Success Metrics**: Calculate success and failure rates for each policy
- **Multiple Export Formats**: JSON, CSV, and HTML output formats
- **Filtering Options**: Filter by policy ID, platform, or other criteria
- **Pagination Support**: Handles large datasets automatically
- **Retry Logic**: Built-in retry mechanism for API failures
- **Comprehensive Logging**: Detailed logging for debugging and monitoring

## Graph API Endpoints Used

This module utilizes the following Microsoft Graph API endpoints:

| Endpoint | Purpose |
|----------|---------|
| `GET /deviceManagement/deviceConfigurations` | Fetch all configuration policies |
| `GET /deviceManagement/deviceConfigurations/{id}/assignments` | Get policy assignments |
| `GET /deviceManagement/deviceConfigurations/{id}/deviceStatuses` | Get device deployment status |
| `GET /deviceManagement/deviceConfigurations/{id}/userStatuses` | Get user deployment status |
| `GET /deviceManagement/deviceConfigurationConflictSummary` | Get policy conflicts |
| `GET /groups/{id}` | Get group information for assignments |

## Required Permissions

Your Azure AD application requires the following Microsoft Graph API permissions:

### Application Permissions (Recommended for automation)
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

### Delegated Permissions (For user context)
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

## Installation

The module is part of the Intune Reporting package. No additional installation required.

```typescript
import { PolicyAssignmentReport } from './reports/policy-assignment-report';
```

## Quick Start

### Basic Usage

```typescript
import { Client } from '@microsoft/microsoft-graph-client';
import { PolicyAssignmentReport } from './reports/policy-assignment-report';

// Assuming you have an authenticated Graph client
const report = new PolicyAssignmentReport(graphClient, config);

// Generate report
const result = await report.execute({
  includeDeviceStatus: true,
  includeUserStatus: true,
  includeConflicts: true,
  includeSuccessRates: true
});

// Access summary
console.log(`Total Policies: ${result.summary.totalPolicies}`);
console.log(`Success Rate: ${result.summary.overallSuccessRate}%`);
```

### Export to Files

```typescript
const result = await report.execute({
  includeDeviceStatus: true,
  includeUserStatus: true,
  includeConflicts: true,
  includeSuccessRates: true,
  outputDir: './reports/policies'
});

// Files will be automatically generated:
// - policy-assignment-{timestamp}.json
// - policy-assignment-{timestamp}.csv
// - policy-assignment-{timestamp}.html
```

## Report Options

The `PolicyReportOptions` interface provides various configuration options:

```typescript
interface PolicyReportOptions {
  /** Include device deployment status (default: true) */
  includeDeviceStatus?: boolean;

  /** Include user deployment status (default: true) */
  includeUserStatus?: boolean;

  /** Include conflict analysis (default: true) */
  includeConflicts?: boolean;

  /** Include success/failure rates (default: true) */
  includeSuccessRates?: boolean;

  /** Filter by specific policy ID */
  policyId?: string;

  /** Filter by platform (Windows, iOS, Android, etc.) */
  platformFilter?: string;

  /** Maximum number of retry attempts (default: 3) */
  maxRetries?: number;

  /** Output directory for exported files */
  outputDir?: string;
}
```

## Data Structures

### PolicyConfiguration

Information about a configuration policy:

```typescript
interface PolicyConfiguration {
  id: string;
  displayName: string;
  description: string;
  platformType: string;          // Windows, iOS, Android, macOS
  technologies: string;           // MDM, MAM, Configuration
  createdDateTime: string;
  lastModifiedDateTime: string;
  assignmentCount: number;
  deviceStatusCount: number;
  userStatusCount: number;
  version?: number;
}
```

### PolicyAssignment

Assignment details for a policy:

```typescript
interface PolicyAssignment {
  id: string;
  policyId: string;
  policyName: string;
  targetGroupId?: string;
  targetGroupName?: string;
  intent?: string;               // apply, remove
  source?: string;               // direct, policySets
  targetType: string;            // Group, All Users, All Devices
  filterType?: string;
}
```

### DeviceDeploymentStatus

Device-level deployment status:

```typescript
interface DeviceDeploymentStatus {
  deviceId: string;
  deviceName: string;
  userPrincipalName: string;
  policyId: string;
  policyName: string;
  status: string;                // success, error, conflict, notApplicable
  complianceState?: string;
  lastReportedDateTime: string;
  platform?: string;
  osVersion?: string;
}
```

### UserDeploymentStatus

User-level deployment status:

```typescript
interface UserDeploymentStatus {
  userId: string;
  userPrincipalName: string;
  policyId: string;
  policyName: string;
  status: string;
  lastReportedDateTime: string;
  devicesCount?: number;
}
```

### PolicyConflict

Policy conflict information:

```typescript
interface PolicyConflict {
  id: string;
  conflictingPolicyIds: string[];
  conflictingPolicyNames: string[];
  affectedDevicesCount: number;
  conflictingSettingName?: string;
  description?: string;
}
```

### PolicySuccessRate

Success/failure rate statistics:

```typescript
interface PolicySuccessRate {
  policyId: string;
  policyName: string;
  totalTargets: number;
  successCount: number;
  failureCount: number;
  conflictCount: number;
  pendingCount: number;
  notApplicableCount: number;
  successRate: number;           // Percentage
  failureRate: number;           // Percentage
}
```

## Usage Examples

### 1. Minimal Report (Summary Only)

```typescript
const result = await report.execute({
  includeDeviceStatus: false,
  includeUserStatus: false,
  includeConflicts: false,
  includeSuccessRates: false
});

console.log(`Policies: ${result.summary.totalPolicies}`);
console.log(`Assigned: ${result.summary.policiesWithAssignments}`);
console.log(`Unassigned: ${result.summary.policiesWithoutAssignments}`);
```

### 2. Specific Policy Report

```typescript
const result = await report.execute({
  policyId: 'your-policy-id-here',
  includeDeviceStatus: true,
  includeUserStatus: true
});
```

### 3. Platform-Filtered Report

```typescript
// Report only Windows policies
const result = await report.execute({
  platformFilter: 'Windows',
  includeDeviceStatus: true,
  includeSuccessRates: true
});
```

### 4. Conflict Analysis

```typescript
const result = await report.execute({
  includeConflicts: true,
  includeDeviceStatus: false,
  includeUserStatus: false
});

if (result.data && result.data.length > 0) {
  const conflicts = result.data[0].conflicts;
  console.log(`Found ${conflicts.length} conflicts`);
}
```

### 5. Success Rate Analysis

```typescript
const result = await report.execute({
  includeSuccessRates: true,
  includeDeviceStatus: true,
  includeUserStatus: true
});

if (result.data && result.data.length > 0) {
  const successRates = result.data[0].successRates;

  // Find policies with low success rates
  const problematic = successRates.filter(r => r.successRate < 80);

  console.log('Policies needing attention:');
  problematic.forEach(r => {
    console.log(`- ${r.policyName}: ${r.successRate}%`);
  });
}
```

### 6. Export to Multiple Formats

```typescript
const result = await report.execute({
  includeDeviceStatus: true,
  includeUserStatus: true,
  includeConflicts: true,
  includeSuccessRates: true,
  outputDir: './reports/policy-assignments'
});

// Files automatically generated:
// - policy-assignment-{timestamp}.json
// - policy-assignment-{timestamp}.csv
// - policy-assignment-{timestamp}.html
```

## Output Formats

### JSON Format

Complete data export including all policies, assignments, statuses, conflicts, and success rates.

```json
{
  "policies": [...],
  "assignments": [...],
  "deviceStatuses": [...],
  "userStatuses": [...],
  "conflicts": [...],
  "successRates": [...],
  "summary": {
    "totalPolicies": 42,
    "policiesWithAssignments": 38,
    "overallSuccessRate": 94.5,
    "totalConflicts": 2
  }
}
```

### CSV Format

Multiple sections for different data types:

```csv
=== CONFIGURATION POLICIES ===
Policy ID,Policy Name,Description,Platform,...

=== ASSIGNMENTS ===
Policy Name,Target Type,Target Group,...

=== SUCCESS RATES ===
Policy Name,Total Targets,Success,Failure,...
```

### HTML Format

Interactive web report with:
- Summary cards with color-coded metrics
- Sortable tables
- Progress bars for success rates
- Responsive design
- Easy sharing and viewing

## Performance Considerations

### Large Datasets

For tenants with many policies (100+):

```typescript
// Use minimal options for faster execution
const result = await report.execute({
  includeDeviceStatus: false,  // Skip if not needed
  includeUserStatus: false,     // Skip if not needed
  includeConflicts: true,
  includeSuccessRates: true,
  maxRetries: 5                 // Increase for reliability
});
```

### Retry Configuration

```typescript
const result = await report.execute({
  maxRetries: 5,  // Increase for unreliable networks
  // Other options...
});
```

### Pagination

The module automatically handles pagination for all Graph API calls. No manual configuration required.

## Error Handling

The module includes comprehensive error handling:

```typescript
try {
  const result = await report.execute(options);
  // Process result
} catch (error) {
  console.error('Report generation failed:', error);
  // Handle error appropriately
}
```

Common errors:
- **Authentication failures**: Check credentials and permissions
- **Rate limiting**: Automatic retry with exponential backoff
- **Network errors**: Automatic retry with configurable attempts
- **Permission errors**: Verify Graph API permissions

## Logging

The module uses the centralized logging service:

```typescript
// Logs are automatically generated at various levels:
// - INFO: Report execution progress
// - DEBUG: Detailed API calls
// - WARN: Non-critical issues (e.g., missing data)
// - ERROR: Critical failures
```

## Best Practices

### 1. Use Filters for Large Tenants

```typescript
// Instead of fetching all policies
const result = await report.execute({
  platformFilter: 'Windows',  // Focus on specific platform
  includeDeviceStatus: true
});
```

### 2. Export to Files for Analysis

```typescript
// Always specify outputDir for record-keeping
const result = await report.execute({
  outputDir: './reports/monthly',
  // Other options...
});
```

### 3. Schedule Regular Reports

```typescript
// Use the scheduler to run reports automatically
// See main README for scheduler configuration
```

### 4. Monitor Success Rates

```typescript
// Regularly check for declining success rates
const result = await report.execute({
  includeSuccessRates: true
});

// Alert if overall success rate drops
if (result.summary.overallSuccessRate < 90) {
  // Send alert
}
```

### 5. Review Conflicts Regularly

```typescript
// Run conflict analysis weekly
const result = await report.execute({
  includeConflicts: true,
  includeDeviceStatus: false,
  includeUserStatus: false
});

if (result.summary.totalConflicts > 0) {
  // Investigate and resolve
}
```

## Integration Examples

### With Email Notifications

```typescript
async function sendPolicyReport() {
  const report = new PolicyAssignmentReport(graphClient, config);

  const result = await report.execute({
    outputDir: './reports/temp',
    includeSuccessRates: true
  });

  // Send HTML report via email
  await sendEmail({
    to: 'admin@company.com',
    subject: 'Weekly Policy Assignment Report',
    htmlBody: result.summary,
    attachments: ['./reports/temp/policy-assignment-*.html']
  });
}
```

### With Azure Storage

```typescript
async function archiveReport() {
  const report = new PolicyAssignmentReport(graphClient, config);

  const result = await report.execute({
    outputDir: './reports/temp'
  });

  // Upload to Azure Blob Storage
  await uploadToBlob(
    'policy-reports',
    './reports/temp/policy-assignment-*.json'
  );
}
```

### With Power BI

```typescript
// Export to CSV for Power BI consumption
const result = await report.execute({
  outputDir: './powerbi/data',
  includeDeviceStatus: true,
  includeSuccessRates: true
});

// Power BI can import the CSV files directly
```

## Troubleshooting

### Issue: No data returned

**Solution**: Check permissions and ensure policies exist

```typescript
// Verify authentication and permissions
const result = await report.execute({
  includeDeviceStatus: false,  // Start minimal
  includeUserStatus: false
});

console.log(`Found ${result.summary.totalPolicies} policies`);
```

### Issue: Timeout errors

**Solution**: Increase retry attempts and reduce data fetched

```typescript
const result = await report.execute({
  maxRetries: 10,
  includeDeviceStatus: false,  // Reduce data volume
  includeUserStatus: false
});
```

### Issue: Missing conflict data

**Solution**: Conflict API may not be available in all licenses

```typescript
// Conflicts are optional and may not be available
const result = await report.execute({
  includeConflicts: false  // Skip if not available
});
```

## Changelog

### Version 1.0.0 (2024-02-04)

- Initial release
- Support for all policy types
- Device and user deployment status
- Conflict detection
- Success rate calculations
- Multiple export formats (JSON, CSV, HTML)
- Comprehensive error handling
- Automatic pagination
- Retry logic with exponential backoff

## Contributing

Contributions are welcome! Please see the main [CONTRIBUTING.md](../CONTRIBUTING.md) for guidelines.

## License

MIT License - See [LICENSE](../LICENSE) for details

## Support

For issues or questions:
- Create an issue in the repository
- Check existing documentation
- Review [Microsoft Graph API documentation](https://docs.microsoft.com/en-us/graph/api/resources/intune-graph-overview)

## Related Modules

- [Device Compliance Report](./device-compliance.ts)
- [Application Inventory Report](./application-inventory.ts)
- [Configuration Profiles Report](./configuration-profiles.ts)
- [Security Compliance Report](./security-compliance.ts)

## References

- [Microsoft Graph API - Intune](https://docs.microsoft.com/en-us/graph/api/resources/intune-graph-overview)
- [Device Configuration](https://docs.microsoft.com/en-us/graph/api/resources/intune-deviceconfig-deviceconfiguration)
- [Assignment Resources](https://docs.microsoft.com/en-us/graph/api/resources/intune-shared-deviceconfigurationassignment)
- [Device Status](https://docs.microsoft.com/en-us/graph/api/resources/intune-deviceconfig-deviceconfigurationdevicestatus)
