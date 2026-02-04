# Quick Start Guide - Intune Reporting Dashboard

Get started with the Intune Reporting Dashboard in 5 minutes!

## Prerequisites

- Node.js 16+ installed
- Azure AD tenant with Intune
- App registration with required permissions

## Step 1: Azure AD Setup (5 minutes)

### Create App Registration

1. Go to [Azure Portal](https://portal.azure.com) > Azure Active Directory > App registrations
2. Click "New registration"
3. Name: `Intune Reporting Dashboard`
4. Click "Register"

### Add Permissions

1. Go to "API permissions" > "Add a permission"
2. Select "Microsoft Graph" > "Application permissions"
3. Add these permissions:
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `DeviceManagementApps.Read.All`
   - `Organization.Read.All`
4. Click "Grant admin consent"

### Create Secret

1. Go to "Certificates & secrets" > "New client secret"
2. Description: `Intune Dashboard`
3. Expiration: 12 months (or 24 months)
4. Click "Add"
5. **COPY THE SECRET VALUE** (you won't see it again!)

### Note Your IDs

From the "Overview" page, copy:
- Application (client) ID
- Directory (tenant) ID

## Step 2: Install & Build (2 minutes)

```bash
# Navigate to project directory
cd /home/user/intunereporting

# Install dependencies
npm install

# Build the project
npm run build
```

## Step 3: Initialize Configuration (1 minute)

```bash
# Create default configuration file
npm run dev -- init
```

This creates `intune-reports.config.yaml` in the current directory.

Edit the file and add your credentials:

```yaml
authentication:
  tenantId: YOUR_TENANT_ID        # From Azure Portal
  clientId: YOUR_CLIENT_ID        # From Azure Portal
  clientSecret: YOUR_CLIENT_SECRET # From Azure Portal
  authMethod: clientSecret
```

**Alternative:** Use environment variables by creating a `.env` file:

```bash
cp .env.example .env
nano .env
```

Add your credentials:
```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-secret-here
```

## Step 4: Test Connection (1 minute)

```bash
# Verify authentication and check status
npm run dev -- status
```

You should see:
```
Intune Reporting Dashboard - Status

Authentication:
  Status: ✓ Connected
  Tenant: your-tenant-id

Configuration:
  Reports Available: 5
  Reports Enabled: 5
  Scheduler: Disabled
```

## Step 5: Run Your First Report (1 minute)

### List available reports

```bash
npm run dev -- list
```

### Generate a single report

```bash
# Device inventory in JSON format
npm run dev -- run -r device-inventory -f json
```

### Generate all reports

```bash
# All reports in CSV format
npm run dev -- run --all --format csv
```

### Generate reports in multiple formats

```bash
# Device and compliance reports in JSON, CSV, and HTML
npm run dev -- run -r device-inventory,compliance-summary -f all
```

## Step 6: View Your Reports

Reports are saved to `./reports/` directory by default:

```
./reports/
  ├── device-inventory-2024-02-04_10-30-00.json
  ├── device-inventory-2024-02-04_10-30-00.csv
  ├── device-inventory-2024-02-04_10-30-00.html
  ├── compliance-summary-2024-02-04_10-30-00.json
  └── compliance-summary-2024-02-04_10-30-00.csv
```

- **JSON**: Complete data with metadata and summaries
- **CSV**: Spreadsheet-ready format for Excel/analysis
- **HTML**: Styled web page with interactive tables

## Step 7: Schedule Reports (Optional)

Edit `intune-reports.config.yaml` to configure scheduling:

```yaml
scheduler:
  enabled: true
  timezone: UTC
  schedules:
    - name: Daily Reports
      cron: '0 6 * * *'  # Run at 6 AM daily
      reports:
        - device-inventory
        - compliance-summary
      formats:
        - json
        - csv
      enabled: true
```

Start the scheduler:

```bash
npm run dev -- schedule
```

Press Ctrl+C to stop when needed.

## Available Commands

| Command | Description | Example |
|---------|-------------|---------|
| `init` | Create default configuration | `npm run dev -- init` |
| `run` | Execute one or more reports | `npm run dev -- run --all` |
| `schedule` | Start automated scheduling | `npm run dev -- schedule` |
| `list` | Show available reports | `npm run dev -- list` |
| `status` | Check system status | `npm run dev -- status` |
| `validate` | Validate configuration | `npm run dev -- validate` |

## Available Reports

1. **device-inventory** - Complete inventory of all managed devices
2. **compliance-summary** - Device compliance status overview
3. **application-inventory** - Managed and discovered applications
4. **user-devices** - Device assignments per user
5. **policy-assignments** - Configuration and compliance policies

## Common Use Cases

### Daily Compliance Report

```bash
npm run dev -- run -r compliance-summary -f html -o ./daily-reports
```

### Weekly Full Inventory

```bash
npm run dev -- run --all -f all -o ./weekly-reports
```

### Specific Device Report

```bash
npm run dev -- run -r device-inventory -f csv
```

### Custom Output Location

```bash
npm run dev -- run --all -f json -o /path/to/output
```

## Troubleshooting

### Authentication Error

```
Error: Authentication failed. Please check your credentials.
```

**Solution:**
1. Verify tenant ID, client ID, and secret in config file
2. Check app has required Graph API permissions
3. Ensure admin consent has been granted
4. Wait a few minutes after granting consent

### Empty Reports

```
✓ device-inventory: 0 records
```

**Solution:**
1. Verify devices exist in your Intune tenant
2. Check app permissions include all required scopes
3. Review logs in `./logs/combined.log`

### Permission Errors

```
Error: Insufficient privileges to complete the operation
```

**Solution:**
1. Grant required Microsoft Graph API permissions
2. Click "Grant admin consent" in Azure Portal
3. Wait 5-10 minutes for permissions to propagate
4. Restart the application

### Configuration Validation Errors

```bash
# Validate your configuration file
npm run dev -- validate
```

Fix any errors shown in the output.

## Next Steps

1. **Customize Reports**: Edit `intune-reports.config.yaml` to enable/disable specific reports
2. **Schedule Automation**: Set up cron schedules for regular reporting
3. **Review Logs**: Check `./logs/` directory for detailed execution logs
4. **Explore Examples**: Review sample reports in `./reports/` directory
5. **Read Full Documentation**: See [README.md](./README.md) for advanced features

## Quick Reference

```bash
# One-time setup
npm install && npm run build
npm run dev -- init
# Edit intune-reports.config.yaml with your credentials

# Daily usage
npm run dev -- run --all -f csv           # All reports as CSV
npm run dev -- run -r device-inventory    # Single report
npm run dev -- list                        # Show available reports
npm run dev -- status                      # Check connection

# Automation
npm run dev -- schedule                    # Start scheduler
npm run dev -- validate                    # Validate config
```

## Understanding Report Output

### JSON Format
```json
{
  "metadata": {
    "reportName": "device-inventory",
    "generatedAt": "2024-02-04T10:30:00Z",
    "recordCount": 150
  },
  "data": [...],
  "summary": {
    "totalDevices": 150,
    "byOS": { "Windows": 100, "iOS": 50 }
  }
}
```

### CSV Format
- Header row with column names
- One device per row
- Metadata as comments at top
- Compatible with Excel/Google Sheets

### HTML Format
- Styled web page
- Summary statistics cards
- Interactive data table
- Responsive design

## Getting Help

- **Logs**: Check `./logs/combined.log` and `./logs/error.log`
- **Validation**: Run `npm run dev -- validate` to check config
- **Status**: Run `npm run dev -- status` to verify connection
- **Documentation**: Read [README.md](./README.md) for details

## Production Deployment

For production use:

1. Use certificate authentication instead of client secret
2. Store credentials in Azure Key Vault
3. Run scheduler as a system service
4. Set up log rotation
5. Monitor with Azure Application Insights

---

**Ready to generate reports!** 🚀

For complete documentation, see [README.md](./README.md)
