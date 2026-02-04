# Intune Reporting Dashboard

A comprehensive reporting solution for Microsoft Intune that provides automated report generation, scheduling, and multiple output formats.

## Features

- **Comprehensive Reporting**: Generate detailed reports on devices, compliance, applications, users, and policies
- **Multiple Output Formats**: Export reports in JSON, CSV, or HTML formats
- **Scheduled Reporting**: Automated report generation using cron scheduling
- **CLI Interface**: Easy-to-use command-line interface for all operations
- **Configuration Management**: YAML-based configuration with environment variable support
- **Logging & Monitoring**: Built-in logging with Winston for tracking and debugging
- **Graph API Integration**: Direct integration with Microsoft Graph API for real-time data

## Prerequisites

- Node.js 16.0.0 or higher
- Azure AD App Registration with appropriate Microsoft Graph permissions
- TypeScript 5.x

## Required Microsoft Graph Permissions

Your Azure AD app registration needs the following delegated or application permissions:

- `DeviceManagementManagedDevices.Read.All`
- `DeviceManagementConfiguration.Read.All`
- `DeviceManagementApps.Read.All`
- `Organization.Read.All`

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd intunereporting
```

2. Install dependencies:
```bash
npm install
```

3. Build the project:
```bash
npm run build
```

4. Initialize configuration:
```bash
npm run dev -- init
```

5. Update the generated `intune-reports.config.yaml` with your Azure AD credentials.

6. Create a `.env` file (optional):
```bash
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
LOG_LEVEL=info
```

## Configuration

The configuration file (`intune-reports.config.yaml`) contains all settings for the dashboard:

### Authentication
```yaml
authentication:
  tenantId: YOUR_TENANT_ID
  clientId: YOUR_CLIENT_ID
  clientSecret: YOUR_CLIENT_SECRET
  authMethod: clientSecret  # or 'certificate' or 'interactive'
```

### Reports
```yaml
reports:
  enabled:
    - device-inventory
    - compliance-summary
    - application-inventory
    - user-devices
    - policy-assignments
  disabled: []
  settings:
    includeInactiveDevices: false
    daysInactive: 30
```

### Output
```yaml
output:
  defaultFormat: json  # json, csv, or html
  directory: ./reports
  includeTimestamp: true
  compression: false
```

### Scheduler
```yaml
scheduler:
  enabled: true
  timezone: UTC
  schedules:
    - name: Daily Reports
      cron: '0 6 * * *'  # 6 AM daily
      reports:
        - device-inventory
        - compliance-summary
      formats:
        - json
        - csv
      enabled: true
```

## Available Commands

### Initialize Configuration
```bash
npm run dev -- init [options]

Options:
  -o, --output <path>  Output directory for configuration file (default: ".")
```

### Run Reports
```bash
npm run dev -- run [options]

Options:
  -c, --config <path>      Path to configuration file (default: "./intune-reports.config.yaml")
  -r, --reports <reports>  Specific reports to run (comma-separated)
  -f, --format <format>    Output format: json, csv, html, all (default: "json")
  -o, --output <path>      Output directory for reports (default: "./reports")
  -a, --all                Run all available reports

Examples:
  npm run dev -- run --all --format json
  npm run dev -- run -r device-inventory,compliance-summary -f csv
  npm run dev -- run -a -f all -o ./my-reports
```

### Schedule Reports
```bash
npm run dev -- schedule [options]

Options:
  -c, --config <path>  Path to configuration file (default: "./intune-reports.config.yaml")
  -d, --daemon         Run as daemon in background

Example:
  npm run dev -- schedule
```

### List Available Reports
```bash
npm run dev -- list [options]

Options:
  -c, --config <path>  Path to configuration file (default: "./intune-reports.config.yaml")
```

### Check Status
```bash
npm run dev -- status [options]

Options:
  -c, --config <path>  Path to configuration file (default: "./intune-reports.config.yaml")
```

### Validate Configuration
```bash
npm run dev -- validate [options]

Options:
  -c, --config <path>  Path to configuration file (default: "./intune-reports.config.yaml")
```

## Available Reports

### 1. Device Inventory Report
**Name**: `device-inventory`

Provides comprehensive inventory of all managed devices including:
- Device name, OS, version
- Manufacturer and model
- Serial number
- Enrollment and last sync dates
- Compliance state
- User assignments

### 2. Compliance Summary Report
**Name**: `compliance-summary`

Overview of device compliance status:
- Total devices by compliance state
- Compliance policies count
- Compliance rate percentage
- Devices in grace period
- Non-compliant devices

### 3. Application Inventory Report
**Name**: `application-inventory`

Lists all managed applications:
- Application names and publishers
- Application types (iOS, Android, Win32, etc.)
- Creation and modification dates
- Deployment status

### 4. User Devices Report
**Name**: `user-devices`

Shows device assignments per user:
- User information (UPN, display name)
- Devices assigned to each user
- Compliance status per device
- Average devices per user

### 5. Policy Assignments Report
**Name**: `policy-assignments`

Configuration and compliance policies:
- Policy names and types
- Group assignments
- Creation and modification dates
- Policies without assignments

## Output Formats

### JSON Format
Structured JSON output with metadata and data arrays:
```json
{
  "metadata": {
    "reportName": "device-inventory",
    "generatedAt": "2024-01-15T10:30:00Z",
    "recordCount": 150
  },
  "data": [...],
  "summary": {...}
}
```

### CSV Format
Comma-separated values with metadata headers:
```csv
# Report: device-inventory
# Generated: 2024-01-15T10:30:00Z
# Records: 150

deviceName,operatingSystem,complianceState,...
Device1,Windows,compliant,...
```

### HTML Format
Styled HTML report with:
- Metadata section
- Summary statistics cards
- Sortable data table
- Responsive design

## Scheduling

The scheduler uses cron expressions for flexible scheduling:

```yaml
schedules:
  - name: Daily Morning Report
    cron: '0 6 * * *'        # Every day at 6 AM
    reports: [device-inventory]
    formats: [json, csv]
    enabled: true

  - name: Weekly Full Report
    cron: '0 8 * * 1'        # Every Monday at 8 AM
    reports: [all]
    formats: [json, csv, html]
    enabled: true

  - name: Hourly Compliance Check
    cron: '0 * * * *'        # Every hour
    reports: [compliance-summary]
    formats: [json]
    enabled: true
```

### Cron Expression Format
```
┌───────────── minute (0 - 59)
│ ┌───────────── hour (0 - 23)
│ │ ┌───────────── day of month (1 - 31)
│ │ │ ┌───────────── month (1 - 12)
│ │ │ │ ┌───────────── day of week (0 - 6) (Sunday to Saturday)
│ │ │ │ │
* * * * *
```

## Logging

Logs are stored in the `./logs` directory:

- `combined.log`: All log messages
- `error.log`: Error messages only

Configure logging in the configuration file:
```yaml
logging:
  level: info  # error, warn, info, debug
  file: ./logs/intune-reports.log
  console: true
  maxSize: 10m
  maxFiles: 5
```

## Development

### Running in Development Mode
```bash
npm run dev -- <command>
```

### Building for Production
```bash
npm run build
npm start -- <command>
```

### Running Tests
```bash
npm test
```

### Linting
```bash
npm run lint
```

## Project Structure

```
intunereporting/
├── src/
│   ├── index.ts                 # Main CLI application
│   ├── core/
│   │   ├── orchestrator.ts      # Report orchestrator
│   │   ├── config.ts            # Configuration manager
│   │   ├── logger.ts            # Logging service
│   │   └── scheduler.ts         # Scheduler manager
│   ├── reports/
│   │   ├── base-report.ts       # Base report class
│   │   ├── device-inventory.ts
│   │   ├── compliance-summary.ts
│   │   ├── application-inventory.ts
│   │   ├── user-devices.ts
│   │   └── policy-assignments.ts
│   ├── formatters/
│   │   └── output-formatter.ts  # Output format handlers
│   └── types/
│       └── index.ts             # TypeScript type definitions
├── dist/                        # Compiled JavaScript (generated)
├── logs/                        # Log files (generated)
├── reports/                     # Generated reports (generated)
├── package.json
├── tsconfig.json
├── intune-reports.config.yaml   # Configuration file (generated)
└── README.md
```

## Security Best Practices

1. **Never commit credentials**: Use environment variables or Azure Key Vault
2. **Use certificate authentication**: More secure than client secrets for production
3. **Principle of least privilege**: Only grant necessary Graph API permissions
4. **Rotate secrets regularly**: Update client secrets on a regular schedule
5. **Secure log files**: Ensure logs directory has appropriate permissions
6. **Review generated reports**: Some reports may contain sensitive information

## Troubleshooting

### Authentication Errors
- Verify tenant ID, client ID, and client secret are correct
- Ensure app registration has required Graph API permissions
- Check that admin consent has been granted for application permissions

### Empty Reports
- Verify the app has access to Intune data
- Check that devices/apps/policies exist in your tenant
- Review logs for API errors

### Scheduler Not Running
- Verify scheduler is enabled in configuration
- Check cron expressions are valid
- Review logs for scheduler errors

### Performance Issues
- For large datasets, reports may take several minutes
- Consider running reports during off-peak hours
- Enable pagination handling for Graph API calls

## API Rate Limiting

The dashboard includes automatic retry logic with exponential backoff to handle Microsoft Graph API rate limits. If you encounter rate limiting:

- Space out scheduled reports
- Reduce frequency of report generation
- Consider using application permissions instead of delegated

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests for new functionality
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For issues, questions, or contributions:
- Create an issue in the repository
- Check existing documentation
- Review Microsoft Graph API documentation

## Changelog

### Version 1.0.0
- Initial release
- Five core reports
- JSON, CSV, and HTML output formats
- Cron-based scheduling
- CLI interface
- Configuration file support
- Logging and monitoring

## Roadmap

Future enhancements:
- [ ] Additional reports (security, update compliance, etc.)
- [ ] Email delivery of reports
- [ ] Web dashboard interface
- [ ] Report templates and customization
- [ ] Azure Storage integration for report archival
- [ ] PowerBI integration
- [ ] Multi-tenant support
- [ ] Report comparison and trending
