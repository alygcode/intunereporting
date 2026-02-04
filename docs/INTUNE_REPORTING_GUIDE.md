# Intune Reporting Guide

A comprehensive guide to building reporting solutions with Microsoft Intune using the Microsoft Graph API.

## Table of Contents

- [Overview](#overview)
- [Microsoft Graph API Fundamentals](#microsoft-graph-api-fundamentals)
- [Authentication and Authorization](#authentication-and-authorization)
- [Common Reporting Scenarios](#common-reporting-scenarios)
- [Best Practices](#best-practices)
- [Rate Limiting and Optimization](#rate-limiting-and-optimization)
- [Troubleshooting](#troubleshooting)
- [Additional Resources](#additional-resources)

---

## Overview

### What is Intune Reporting?

Microsoft Intune provides comprehensive device and application management capabilities for organizations. Intune reporting allows administrators and developers to:

- **Monitor device compliance** across the organization
- **Track application deployment** status and usage
- **Analyze security posture** and policy effectiveness
- **Generate custom reports** for stakeholders
- **Automate compliance reporting** for auditing purposes
- **Identify trends** in device enrollment and usage

### Available Reporting Capabilities

Intune offers several reporting mechanisms:

1. **Built-in Reports** - Available through the Microsoft Intune admin center
2. **Microsoft Graph API** - Programmatic access to Intune data
3. **Data Warehouse** - Historical data export for long-term analysis
4. **Power BI Integration** - Visual analytics and dashboards
5. **Log Analytics** - Integration with Azure Monitor for advanced querying

This guide focuses primarily on using the Microsoft Graph API for custom reporting solutions.

### When to Use Graph API for Reporting

Consider using the Graph API when you need to:

- Automate report generation on a schedule
- Integrate Intune data with other systems
- Create custom dashboards or visualizations
- Perform bulk data analysis
- Build applications that require real-time Intune data
- Export data to external databases or data warehouses

---

## Microsoft Graph API Fundamentals

### What is Microsoft Graph?

Microsoft Graph is a unified API endpoint that provides access to data across Microsoft 365 services, including Intune. It uses RESTful HTTP requests and returns data in JSON format.

**Base URL**: `https://graph.microsoft.com`

**API Versions**:
- `v1.0` - Production-ready, stable endpoints (recommended for production)
- `beta` - Preview endpoints with latest features (may change without notice)

### Key Intune Endpoints

#### Device Management

```
GET /deviceManagement/managedDevices
GET /deviceManagement/managedDevices/{id}
GET /deviceManagement/managedDevices/{id}/deviceCompliancePolicyStates
```

#### Compliance Policies

```
GET /deviceManagement/deviceCompliancePolicies
GET /deviceManagement/deviceCompliancePolicies/{id}
GET /deviceManagement/deviceCompliancePolicies/{id}/deviceStatuses
```

#### Configuration Profiles

```
GET /deviceManagement/deviceConfigurations
GET /deviceManagement/deviceConfigurations/{id}
GET /deviceManagement/deviceConfigurations/{id}/deviceStatuses
```

#### Applications

```
GET /deviceAppManagement/mobileApps
GET /deviceAppManagement/mobileApps/{id}
GET /deviceAppManagement/mobileApps/{id}/installSummary
```

#### Users and Devices

```
GET /users/{id}/managedDevices
GET /users/{id}/deviceManagementTroubleshootingEvents
```

### Understanding Graph Query Parameters

Microsoft Graph supports OData query parameters for filtering, sorting, and pagination:

#### $filter - Filter Results

```http
GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'Windows'
GET /deviceManagement/managedDevices?$filter=complianceState eq 'noncompliant'
GET /deviceManagement/managedDevices?$filter=startswith(deviceName,'LAPTOP')
```

#### $select - Choose Specific Properties

```http
GET /deviceManagement/managedDevices?$select=deviceName,operatingSystem,complianceState
```

#### $orderby - Sort Results

```http
GET /deviceManagement/managedDevices?$orderby=lastSyncDateTime desc
```

#### $top - Limit Results

```http
GET /deviceManagement/managedDevices?$top=100
```

#### $expand - Include Related Resources

```http
GET /deviceManagement/managedDevices?$expand=deviceCompliancePolicyStates
```

#### Combining Parameters

```http
GET /deviceManagement/managedDevices?$filter=operatingSystem eq 'iOS'&$select=deviceName,userPrincipalName&$orderby=lastSyncDateTime desc&$top=50
```

---

## Authentication and Authorization

### Authentication Methods

Microsoft Graph supports several authentication methods:

1. **Application Permissions** (App-only) - For background services and automation
2. **Delegated Permissions** (On behalf of user) - For interactive applications
3. **Certificate-based Authentication** - Enhanced security for production applications

For reporting scenarios, **Application Permissions** are typically recommended.

### Creating an App Registration

#### Step 1: Register Application in Azure AD

1. Navigate to [Azure Portal](https://portal.azure.com)
2. Go to **Azure Active Directory** > **App registrations**
3. Click **New registration**
4. Provide:
   - **Name**: e.g., "Intune Reporting App"
   - **Supported account types**: Select appropriate option
   - **Redirect URI**: Leave blank for service apps
5. Click **Register**

#### Step 2: Note Application Details

After registration, save these values:
- **Application (client) ID**
- **Directory (tenant) ID**

#### Step 3: Create Client Secret

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Provide a description and expiration period
4. Click **Add**
5. **Copy the secret value immediately** (it won't be shown again)

**Security Note**: For production environments, use certificates instead of client secrets for enhanced security.

### Required API Permissions

#### For Comprehensive Device Reporting

Navigate to **API permissions** and add these Microsoft Graph permissions:

**Application Permissions** (recommended for automated reporting):

- `DeviceManagementManagedDevices.Read.All` - Read managed device information
- `DeviceManagementConfiguration.Read.All` - Read device configuration policies
- `DeviceManagementApps.Read.All` - Read application information
- `DeviceManagementServiceConfig.Read.All` - Read service configuration
- `User.Read.All` - Read user information (if correlating with user data)
- `Group.Read.All` - Read group information (if needed)
- `Directory.Read.All` - Read directory data (optional, for extended user info)

**Delegated Permissions** (for user-context applications):

- `DeviceManagementManagedDevices.Read` - Read user's managed devices
- `DeviceManagementConfiguration.Read` - Read device configurations
- `DeviceManagementApps.Read` - Read application information

#### Permission Assignment Process

1. Click **Add a permission**
2. Select **Microsoft Graph**
3. Choose **Application permissions** or **Delegated permissions**
4. Search for and select the required permissions
5. Click **Add permissions**
6. Click **Grant admin consent for [Your Organization]** (requires admin privileges)

### Authentication Code Examples

#### Python - Using MSAL

```python
from msal import ConfidentialClientApplication
import requests

# Configuration
tenant_id = "YOUR_TENANT_ID"
client_id = "YOUR_CLIENT_ID"
client_secret = "YOUR_CLIENT_SECRET"
authority = f"https://login.microsoftonline.com/{tenant_id}"
scopes = ["https://graph.microsoft.com/.default"]

# Initialize MSAL client
app = ConfidentialClientApplication(
    client_id,
    authority=authority,
    client_credential=client_secret
)

# Acquire token
result = app.acquire_token_silent(scopes, account=None)
if not result:
    result = app.acquire_token_for_client(scopes=scopes)

if "access_token" in result:
    access_token = result["access_token"]
    print("Authentication successful")
else:
    print(f"Authentication failed: {result.get('error_description')}")
    exit(1)

# Make API request
headers = {
    "Authorization": f"Bearer {access_token}",
    "Content-Type": "application/json"
}

response = requests.get(
    "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices",
    headers=headers
)

if response.status_code == 200:
    devices = response.json()
    print(f"Retrieved {len(devices.get('value', []))} devices")
else:
    print(f"Error: {response.status_code} - {response.text}")
```

#### PowerShell - Using Microsoft.Graph Module

```powershell
# Install Microsoft.Graph module if not already installed
# Install-Module Microsoft.Graph -Scope CurrentUser

# Connect using app credentials
$tenantId = "YOUR_TENANT_ID"
$clientId = "YOUR_CLIENT_ID"
$clientSecret = "YOUR_CLIENT_SECRET"

$body = @{
    Grant_Type    = "client_credentials"
    Scope         = "https://graph.microsoft.com/.default"
    Client_Id     = $clientId
    Client_Secret = $clientSecret
}

$connection = Invoke-RestMethod `
    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
    -Method POST `
    -Body $body

$token = $connection.access_token

# Make API request
$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type"  = "application/json"
}

$uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
$response = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get

Write-Host "Retrieved $($response.value.Count) devices"
```

#### C# - Using Microsoft.Identity.Client

```csharp
using Microsoft.Identity.Client;
using System.Net.Http;
using System.Net.Http.Headers;

public class IntuneReporting
{
    private readonly string tenantId = "YOUR_TENANT_ID";
    private readonly string clientId = "YOUR_CLIENT_ID";
    private readonly string clientSecret = "YOUR_CLIENT_SECRET";
    private readonly string[] scopes = new[] { "https://graph.microsoft.com/.default" };

    public async Task<string> GetAccessToken()
    {
        var app = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
            .Build();

        var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
        return result.AccessToken;
    }

    public async Task<string> GetManagedDevices()
    {
        var token = await GetAccessToken();

        using var httpClient = new HttpClient();
        httpClient.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", token);

        var response = await httpClient.GetAsync(
            "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        );

        response.EnsureSuccessStatusCode();
        return await response.Content.ReadAsStringAsync();
    }
}
```

---

## Common Reporting Scenarios

### 1. Device Compliance Report

Generate a report showing all devices and their compliance status.

```python
import requests
from datetime import datetime

def get_device_compliance_report(access_token):
    """
    Retrieve all managed devices with compliance information
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Get all managed devices with compliance state
    url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
    params = {
        "$select": "id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime,managementAgent",
        "$filter": "managementAgent eq 'mdm' or managementAgent eq 'easMdm'",
        "$orderby": "complianceState,deviceName"
    }

    devices = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        devices.extend(data.get("value", []))
        url = data.get("@odata.nextLink")  # Handle pagination
        params = None  # Params are included in nextLink

    # Generate summary statistics
    total_devices = len(devices)
    compliant = sum(1 for d in devices if d.get("complianceState") == "compliant")
    non_compliant = sum(1 for d in devices if d.get("complianceState") == "noncompliant")
    unknown = sum(1 for d in devices if d.get("complianceState") in ["unknown", None])

    report = {
        "generated_at": datetime.utcnow().isoformat(),
        "summary": {
            "total_devices": total_devices,
            "compliant": compliant,
            "non_compliant": non_compliant,
            "unknown": unknown,
            "compliance_rate": f"{(compliant/total_devices*100):.2f}%" if total_devices > 0 else "0%"
        },
        "devices": devices
    }

    return report

# Usage
report = get_device_compliance_report(access_token)
print(f"Compliance Rate: {report['summary']['compliance_rate']}")
print(f"Non-Compliant Devices: {report['summary']['non_compliant']}")
```

### 2. Application Installation Status Report

Track the installation status of applications across devices.

```python
def get_app_installation_report(access_token, app_id=None):
    """
    Get application installation status across all devices
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Get all mobile apps or specific app
    if app_id:
        apps_url = f"https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/{app_id}"
        response = requests.get(apps_url, headers=headers)
        response.raise_for_status()
        apps = [response.json()]
    else:
        apps_url = "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps"
        response = requests.get(apps_url, headers=headers)
        response.raise_for_status()
        apps = response.json().get("value", [])

    report = []

    for app in apps:
        app_id = app["id"]
        app_name = app.get("displayName", "Unknown")

        # Get installation summary
        summary_url = f"https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps/{app_id}/installSummary"
        try:
            response = requests.get(summary_url, headers=headers)
            response.raise_for_status()
            summary = response.json()

            report.append({
                "app_id": app_id,
                "app_name": app_name,
                "installed_device_count": summary.get("installedDeviceCount", 0),
                "failed_device_count": summary.get("failedDeviceCount", 0),
                "pending_install_device_count": summary.get("pendingInstallDeviceCount", 0),
                "not_applicable_device_count": summary.get("notApplicableDeviceCount", 0),
                "not_installed_device_count": summary.get("notInstalledDeviceCount", 0)
            })
        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 404:
                # No installation summary available
                continue
            raise

    return report

# Usage
app_report = get_app_installation_report(access_token)
for app in app_report:
    print(f"{app['app_name']}: {app['installed_device_count']} installed, {app['failed_device_count']} failed")
```

### 3. Device Inventory Report

Create a comprehensive inventory of all managed devices.

```python
def get_device_inventory_report(access_token):
    """
    Generate comprehensive device inventory with hardware details
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
    params = {
        "$select": "id,deviceName,userPrincipalName,emailAddress,operatingSystem,osVersion,manufacturer,model,serialNumber,imei,freeStorageSpaceInBytes,totalStorageSpaceInBytes,phoneNumber,enrolledDateTime,lastSyncDateTime,managementAgent,complianceState,azureADRegistered,isEncrypted"
    }

    devices = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        devices.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None

    # Organize by OS
    inventory = {
        "total_devices": len(devices),
        "by_os": {},
        "by_manufacturer": {},
        "devices": devices
    }

    for device in devices:
        os = device.get("operatingSystem", "Unknown")
        manufacturer = device.get("manufacturer", "Unknown")

        inventory["by_os"][os] = inventory["by_os"].get(os, 0) + 1
        inventory["by_manufacturer"][manufacturer] = inventory["by_manufacturer"].get(manufacturer, 0) + 1

    return inventory

# Usage
inventory = get_device_inventory_report(access_token)
print(f"Total Devices: {inventory['total_devices']}")
print("\nDevices by OS:")
for os, count in inventory['by_os'].items():
    print(f"  {os}: {count}")
```

### 4. Non-Compliant Devices with Policy Details

Identify non-compliant devices and the policies they're failing.

```python
def get_noncompliant_devices_with_policies(access_token):
    """
    Get all non-compliant devices with details about which policies they're failing
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Get non-compliant devices
    devices_url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
    params = {
        "$filter": "complianceState eq 'noncompliant'",
        "$select": "id,deviceName,userPrincipalName,operatingSystem,lastSyncDateTime"
    }

    response = requests.get(devices_url, headers=headers, params=params)
    response.raise_for_status()
    devices = response.json().get("value", [])

    report = []

    for device in devices:
        device_id = device["id"]

        # Get compliance policy states for this device
        policy_states_url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{device_id}/deviceCompliancePolicyStates"

        try:
            response = requests.get(policy_states_url, headers=headers)
            response.raise_for_status()
            policy_states = response.json().get("value", [])

            failed_policies = []
            for policy_state in policy_states:
                if policy_state.get("state") == "nonCompliant":
                    failed_policies.append({
                        "policy_name": policy_state.get("displayName", "Unknown"),
                        "policy_id": policy_state.get("id"),
                        "user_principal_name": policy_state.get("userPrincipalName")
                    })

            if failed_policies:
                report.append({
                    "device_name": device.get("deviceName"),
                    "user_principal_name": device.get("userPrincipalName"),
                    "operating_system": device.get("operatingSystem"),
                    "last_sync": device.get("lastSyncDateTime"),
                    "failed_policies": failed_policies
                })
        except requests.exceptions.HTTPError:
            continue

    return report

# Usage
noncompliant_report = get_noncompliant_devices_with_policies(access_token)
for device in noncompliant_report:
    print(f"\nDevice: {device['device_name']} (User: {device['user_principal_name']})")
    print("Failed Policies:")
    for policy in device['failed_policies']:
        print(f"  - {policy['policy_name']}")
```

### 5. Stale Device Report

Identify devices that haven't synced recently.

```python
from datetime import datetime, timedelta

def get_stale_devices_report(access_token, days_threshold=30):
    """
    Get devices that haven't synced within the specified number of days
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    cutoff_date = (datetime.utcnow() - timedelta(days=days_threshold)).isoformat() + "Z"

    url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
    params = {
        "$filter": f"lastSyncDateTime lt {cutoff_date}",
        "$select": "deviceName,userPrincipalName,operatingSystem,lastSyncDateTime,enrolledDateTime,complianceState",
        "$orderby": "lastSyncDateTime asc"
    }

    devices = []
    while url:
        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        devices.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
        params = None

    # Calculate days since last sync
    for device in devices:
        last_sync = datetime.fromisoformat(device["lastSyncDateTime"].replace("Z", "+00:00"))
        days_stale = (datetime.now(last_sync.tzinfo) - last_sync).days
        device["days_since_sync"] = days_stale

    report = {
        "threshold_days": days_threshold,
        "stale_device_count": len(devices),
        "devices": devices
    }

    return report

# Usage
stale_report = get_stale_devices_report(access_token, days_threshold=30)
print(f"Found {stale_report['stale_device_count']} devices not synced in 30+ days")
```

### 6. Device Configuration Profile Status

Report on configuration profile deployment status.

```python
def get_configuration_profile_status(access_token):
    """
    Get status of all device configuration profiles
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Get all device configurations
    configs_url = "https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations"
    response = requests.get(configs_url, headers=headers)
    response.raise_for_status()
    configurations = response.json().get("value", [])

    report = []

    for config in configurations:
        config_id = config["id"]
        config_name = config.get("displayName", "Unknown")

        # Get device statuses for this configuration
        statuses_url = f"https://graph.microsoft.com/v1.0/deviceManagement/deviceConfigurations/{config_id}/deviceStatuses"

        try:
            response = requests.get(statuses_url, headers=headers)
            response.raise_for_status()
            statuses = response.json().get("value", [])

            # Count status types
            status_summary = {
                "compliant": 0,
                "notApplicable": 0,
                "error": 0,
                "conflict": 0,
                "notAssigned": 0,
                "unknown": 0
            }

            for status in statuses:
                state = status.get("status", "unknown")
                status_summary[state] = status_summary.get(state, 0) + 1

            report.append({
                "configuration_id": config_id,
                "configuration_name": config_name,
                "total_devices": len(statuses),
                "status_summary": status_summary
            })
        except requests.exceptions.HTTPError:
            continue

    return report

# Usage
config_report = get_configuration_profile_status(access_token)
for config in config_report:
    print(f"\n{config['configuration_name']}")
    print(f"  Total Devices: {config['total_devices']}")
    print(f"  Compliant: {config['status_summary']['compliant']}")
    print(f"  Errors: {config['status_summary']['error']}")
```

### 7. User Device Assignment Report

Show which users have which devices assigned.

```python
def get_user_device_assignment_report(access_token, user_upn=None):
    """
    Get device assignments for users
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    if user_upn:
        # Get specific user's devices
        user_url = f"https://graph.microsoft.com/v1.0/users/{user_upn}/managedDevices"
        response = requests.get(user_url, headers=headers)
        response.raise_for_status()
        devices = response.json().get("value", [])

        report = {
            user_upn: {
                "device_count": len(devices),
                "devices": devices
            }
        }
    else:
        # Get all devices grouped by user
        devices_url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        params = {
            "$select": "deviceName,userPrincipalName,operatingSystem,complianceState,lastSyncDateTime"
        }

        all_devices = []
        url = devices_url
        while url:
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()
            data = response.json()
            all_devices.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None

        # Group by user
        report = {}
        for device in all_devices:
            upn = device.get("userPrincipalName", "Unassigned")
            if upn not in report:
                report[upn] = {
                    "device_count": 0,
                    "devices": []
                }
            report[upn]["device_count"] += 1
            report[upn]["devices"].append(device)

    return report

# Usage
user_report = get_user_device_assignment_report(access_token)
for user, data in sorted(user_report.items(), key=lambda x: x[1]['device_count'], reverse=True)[:10]:
    print(f"{user}: {data['device_count']} device(s)")
```

---

## Best Practices

### 1. Token Management

**Use Token Caching**

Avoid acquiring a new token for every request. Tokens are valid for 60-90 minutes.

```python
import time

class TokenManager:
    def __init__(self, msal_app, scopes):
        self.msal_app = msal_app
        self.scopes = scopes
        self.token = None
        self.token_expiry = 0

    def get_token(self):
        """Get valid token, refresh if needed"""
        current_time = time.time()

        # Check if token is still valid (with 5-minute buffer)
        if self.token and current_time < (self.token_expiry - 300):
            return self.token

        # Acquire new token
        result = self.msal_app.acquire_token_silent(self.scopes, account=None)
        if not result:
            result = self.msal_app.acquire_token_for_client(scopes=self.scopes)

        if "access_token" in result:
            self.token = result["access_token"]
            # Calculate expiry time
            self.token_expiry = current_time + result.get("expires_in", 3600)
            return self.token
        else:
            raise Exception(f"Token acquisition failed: {result.get('error_description')}")

# Usage
token_manager = TokenManager(msal_app, scopes)
access_token = token_manager.get_token()
```

### 2. Error Handling

Implement robust error handling for API requests.

```python
import time
import requests
from requests.exceptions import RequestException

def make_graph_request(url, headers, max_retries=3, retry_delay=5):
    """
    Make Graph API request with retry logic
    """
    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers, timeout=30)

            # Handle rate limiting (429)
            if response.status_code == 429:
                retry_after = int(response.headers.get("Retry-After", retry_delay))
                print(f"Rate limited. Waiting {retry_after} seconds...")
                time.sleep(retry_after)
                continue

            # Handle server errors (500, 502, 503, 504)
            if response.status_code >= 500:
                if attempt < max_retries - 1:
                    print(f"Server error {response.status_code}. Retrying in {retry_delay} seconds...")
                    time.sleep(retry_delay)
                    continue
                else:
                    response.raise_for_status()

            # Handle other errors
            response.raise_for_status()
            return response.json()

        except RequestException as e:
            if attempt < max_retries - 1:
                print(f"Request failed: {e}. Retrying in {retry_delay} seconds...")
                time.sleep(retry_delay)
            else:
                raise

    raise Exception(f"Failed after {max_retries} attempts")
```

### 3. Pagination Handling

Always handle pagination for large result sets.

```python
def get_all_pages(url, headers, params=None):
    """
    Fetch all pages from a Graph API endpoint
    """
    all_results = []

    while url:
        response = make_graph_request(url, headers)

        # Add results from this page
        all_results.extend(response.get("value", []))

        # Get next page URL
        url = response.get("@odata.nextLink")

        # Don't include params in subsequent requests (they're in nextLink)
        params = None

    return all_results

# Usage
devices = get_all_pages(
    "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices",
    headers,
    params={"$top": 100}
)
```

### 4. Efficient Querying

Use `$select` to retrieve only needed properties and reduce payload size.

```python
# BAD - Returns all properties (large payload)
url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"

# GOOD - Returns only needed properties
url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$select=deviceName,complianceState,operatingSystem"
```

### 5. Batch Requests

Use batch requests to combine multiple operations into a single HTTP call.

```python
def batch_get_device_details(access_token, device_ids):
    """
    Get details for multiple devices using batch request
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Construct batch request
    batch_request = {
        "requests": []
    }

    for i, device_id in enumerate(device_ids[:20]):  # Max 20 requests per batch
        batch_request["requests"].append({
            "id": str(i),
            "method": "GET",
            "url": f"/deviceManagement/managedDevices/{device_id}"
        })

    # Send batch request
    response = requests.post(
        "https://graph.microsoft.com/v1.0/$batch",
        headers=headers,
        json=batch_request
    )
    response.raise_for_status()

    # Process responses
    batch_response = response.json()
    results = {}

    for response_item in batch_response.get("responses", []):
        request_id = response_item["id"]
        if response_item["status"] == 200:
            results[device_ids[int(request_id)]] = response_item["body"]
        else:
            results[device_ids[int(request_id)]] = {"error": response_item.get("body")}

    return results

# Usage
device_ids = ["device-id-1", "device-id-2", "device-id-3"]
batch_results = batch_get_device_details(access_token, device_ids)
```

### 6. Logging and Monitoring

Implement comprehensive logging for troubleshooting and auditing.

```python
import logging
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f'intune_reporting_{datetime.now().strftime("%Y%m%d")}.log'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("IntuneReporting")

def get_devices_with_logging(access_token):
    """
    Get devices with comprehensive logging
    """
    logger.info("Starting device retrieval")

    try:
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        logger.debug(f"Requesting URL: {url}")

        response = requests.get(url, headers=headers)
        response.raise_for_status()

        devices = response.json().get("value", [])
        logger.info(f"Successfully retrieved {len(devices)} devices")

        return devices

    except requests.exceptions.HTTPError as e:
        logger.error(f"HTTP error occurred: {e.response.status_code} - {e.response.text}")
        raise
    except Exception as e:
        logger.error(f"Unexpected error occurred: {str(e)}", exc_info=True)
        raise
```

### 7. Secure Credential Management

Never hardcode credentials. Use secure methods to store and retrieve them.

```python
# BAD - Hardcoded credentials
client_secret = "abc123..."

# GOOD - Environment variables
import os
client_secret = os.environ.get("AZURE_CLIENT_SECRET")

# BETTER - Azure Key Vault
from azure.identity import DefaultAzureCredential
from azure.keyvault.secrets import SecretClient

credential = DefaultAzureCredential()
secret_client = SecretClient(
    vault_url="https://your-vault.vault.azure.net/",
    credential=credential
)
client_secret = secret_client.get_secret("ClientSecret").value
```

### 8. Data Privacy and Compliance

Be mindful of data privacy regulations when storing and processing Intune data.

**Best Practices:**
- Minimize data retention periods
- Encrypt sensitive data at rest and in transit
- Implement access controls on reports
- Log access to sensitive data
- Anonymize or pseudonymize data when possible
- Document data processing activities for compliance (GDPR, HIPAA, etc.)

### 9. Performance Optimization

**Use Parallel Processing for Multiple API Calls**

```python
import concurrent.futures

def get_multiple_device_details(access_token, device_ids):
    """
    Fetch details for multiple devices in parallel
    """
    def fetch_device(device_id):
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{device_id}"
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()

    # Use ThreadPoolExecutor for parallel requests
    with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
        future_to_device = {executor.submit(fetch_device, device_id): device_id
                           for device_id in device_ids}

        results = {}
        for future in concurrent.futures.as_completed(future_to_device):
            device_id = future_to_device[future]
            try:
                results[device_id] = future.result()
            except Exception as e:
                logger.error(f"Error fetching device {device_id}: {e}")
                results[device_id] = {"error": str(e)}

        return results
```

### 10. Testing Best Practices

**Use Mock Data for Development**

```python
import unittest
from unittest.mock import Mock, patch

class TestIntuneReporting(unittest.TestCase):

    @patch('requests.get')
    def test_get_devices(self, mock_get):
        # Mock API response
        mock_response = Mock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "value": [
                {
                    "id": "device-1",
                    "deviceName": "TEST-DEVICE",
                    "complianceState": "compliant"
                }
            ]
        }
        mock_get.return_value = mock_response

        # Test function
        devices = get_devices_with_logging("fake_token")

        # Assertions
        self.assertEqual(len(devices), 1)
        self.assertEqual(devices[0]["deviceName"], "TEST-DEVICE")
```

---

## Rate Limiting and Optimization

### Understanding Microsoft Graph Throttling

Microsoft Graph implements throttling to ensure service stability and fair resource allocation.

**Throttling Limits:**
- **Service-specific limits**: Each service (Intune, SharePoint, etc.) has its own limits
- **Per-app per-tenant**: Limits apply per application per tenant
- **Concurrent requests**: Typically ~20-30 concurrent requests
- **Requests per second**: Varies by endpoint, generally 100-1000 req/sec

**When throttling occurs:**
- HTTP Status Code: `429 Too Many Requests`
- Response Header: `Retry-After` (seconds to wait)

### Throttling Response Headers

```
HTTP/1.1 429 Too Many Requests
Content-Type: application/json
Retry-After: 60

{
  "error": {
    "code": "TooManyRequests",
    "message": "Rate limit exceeded"
  }
}
```

### Handling Throttling

#### Basic Retry Logic with Exponential Backoff

```python
import time
import random

def make_request_with_backoff(url, headers, max_retries=5):
    """
    Make request with exponential backoff retry logic
    """
    base_delay = 1

    for attempt in range(max_retries):
        try:
            response = requests.get(url, headers=headers)

            if response.status_code == 429:
                # Get retry-after header or calculate backoff
                retry_after = int(response.headers.get("Retry-After", 0))

                if retry_after > 0:
                    wait_time = retry_after
                else:
                    # Exponential backoff with jitter
                    wait_time = min(base_delay * (2 ** attempt) + random.uniform(0, 1), 60)

                logger.warning(f"Rate limited. Waiting {wait_time:.2f} seconds before retry {attempt + 1}/{max_retries}")
                time.sleep(wait_time)
                continue

            response.raise_for_status()
            return response.json()

        except requests.exceptions.HTTPError as e:
            if e.response.status_code >= 500 and attempt < max_retries - 1:
                wait_time = base_delay * (2 ** attempt)
                logger.warning(f"Server error. Retrying in {wait_time} seconds...")
                time.sleep(wait_time)
                continue
            raise

    raise Exception(f"Failed after {max_retries} attempts")
```

### Rate Limiting Optimization Strategies

#### 1. Request Batching

Combine multiple requests into a single batch operation.

```python
def batch_requests(requests_list, access_token):
    """
    Execute multiple requests as a batch
    Max 20 requests per batch
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    # Split into chunks of 20 (max batch size)
    chunk_size = 20
    all_results = []

    for i in range(0, len(requests_list), chunk_size):
        chunk = requests_list[i:i + chunk_size]

        batch_payload = {
            "requests": [
                {
                    "id": str(idx),
                    "method": req.get("method", "GET"),
                    "url": req["url"]
                }
                for idx, req in enumerate(chunk)
            ]
        }

        response = requests.post(
            "https://graph.microsoft.com/v1.0/$batch",
            headers=headers,
            json=batch_payload
        )
        response.raise_for_status()

        batch_response = response.json()
        all_results.extend(batch_response.get("responses", []))

    return all_results
```

#### 2. Caching

Implement caching to reduce redundant API calls.

```python
from functools import lru_cache
from datetime import datetime, timedelta
import hashlib
import json

class CachedGraphClient:
    def __init__(self, access_token, cache_duration_minutes=15):
        self.access_token = access_token
        self.cache = {}
        self.cache_duration = timedelta(minutes=cache_duration_minutes)

    def _get_cache_key(self, url, params):
        """Generate cache key from URL and parameters"""
        key_data = f"{url}:{json.dumps(params, sort_keys=True)}"
        return hashlib.md5(key_data.encode()).hexdigest()

    def get(self, url, params=None):
        """Get with caching"""
        cache_key = self._get_cache_key(url, params)

        # Check cache
        if cache_key in self.cache:
            cached_data, cached_time = self.cache[cache_key]
            if datetime.now() - cached_time < self.cache_duration:
                logger.debug(f"Cache hit for {url}")
                return cached_data

        # Make request
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        response = requests.get(url, headers=headers, params=params)
        response.raise_for_status()
        data = response.json()

        # Store in cache
        self.cache[cache_key] = (data, datetime.now())
        logger.debug(f"Cached response for {url}")

        return data

    def clear_cache(self):
        """Clear all cached data"""
        self.cache.clear()

# Usage
client = CachedGraphClient(access_token, cache_duration_minutes=15)
devices = client.get("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices")
```

#### 3. Delta Queries for Changed Data

Use delta queries to fetch only changed data since last query.

```python
def get_devices_delta(access_token, delta_link=None):
    """
    Get devices using delta query
    Returns: (devices, next_delta_link)
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }

    if delta_link:
        # Use existing delta link
        url = delta_link
    else:
        # Initial delta query
        url = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/delta"

    all_devices = []

    while url:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()

        all_devices.extend(data.get("value", []))

        # Check for delta link or next link
        if "@odata.deltaLink" in data:
            # Save this for next delta query
            next_delta_link = data["@odata.deltaLink"]
            url = None
        elif "@odata.nextLink" in data:
            url = data["@odata.nextLink"]
        else:
            next_delta_link = None
            url = None

    return all_devices, next_delta_link

# Usage
# First call - get all devices and delta link
devices, delta_link = get_devices_delta(access_token)
print(f"Initial fetch: {len(devices)} devices")

# Save delta_link for next run...

# Subsequent calls - get only changed devices
changed_devices, new_delta_link = get_devices_delta(access_token, delta_link)
print(f"Changed devices: {len(changed_devices)}")
```

#### 4. Request Timing and Scheduling

Distribute requests over time to avoid hitting rate limits.

```python
import time
from datetime import datetime

class RateLimitedClient:
    def __init__(self, access_token, requests_per_second=10):
        self.access_token = access_token
        self.requests_per_second = requests_per_second
        self.min_interval = 1.0 / requests_per_second
        self.last_request_time = 0

    def get(self, url, params=None):
        """Make rate-limited GET request"""
        # Wait if necessary to respect rate limit
        current_time = time.time()
        time_since_last = current_time - self.last_request_time

        if time_since_last < self.min_interval:
            sleep_time = self.min_interval - time_since_last
            time.sleep(sleep_time)

        # Make request
        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

        response = requests.get(url, headers=headers, params=params)
        self.last_request_time = time.time()

        response.raise_for_status()
        return response.json()

# Usage
client = RateLimitedClient(access_token, requests_per_second=5)
for i in range(100):
    devices = client.get("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices")
    # Automatically rate-limited to 5 requests per second
```

#### 5. Parallel Processing with Rate Limiting

```python
from concurrent.futures import ThreadPoolExecutor
import threading
import time

class ThrottledExecutor:
    def __init__(self, max_workers=5, rate_limit=10):
        self.executor = ThreadPoolExecutor(max_workers=max_workers)
        self.rate_limit = rate_limit  # requests per second
        self.min_interval = 1.0 / rate_limit
        self.lock = threading.Lock()
        self.last_execution = 0

    def submit(self, fn, *args, **kwargs):
        """Submit function with rate limiting"""
        def rate_limited_fn(*args, **kwargs):
            with self.lock:
                current_time = time.time()
                time_since_last = current_time - self.last_execution

                if time_since_last < self.min_interval:
                    time.sleep(self.min_interval - time_since_last)

                self.last_execution = time.time()

            return fn(*args, **kwargs)

        return self.executor.submit(rate_limited_fn, *args, **kwargs)

    def shutdown(self):
        self.executor.shutdown(wait=True)

# Usage
def fetch_device(device_id, access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    url = f"https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{device_id}"
    response = requests.get(url, headers=headers)
    response.raise_for_status()
    return response.json()

executor = ThrottledExecutor(max_workers=5, rate_limit=10)
device_ids = ["id1", "id2", "id3", "..."]

futures = [executor.submit(fetch_device, device_id, access_token)
           for device_id in device_ids]

results = [future.result() for future in futures]
executor.shutdown()
```

### Monitoring and Metrics

Track API usage to identify optimization opportunities.

```python
import time
from collections import defaultdict

class GraphAPIMonitor:
    def __init__(self):
        self.stats = {
            "total_requests": 0,
            "successful_requests": 0,
            "failed_requests": 0,
            "throttled_requests": 0,
            "total_response_time": 0,
            "endpoints": defaultdict(int)
        }

    def record_request(self, url, status_code, response_time):
        """Record API request metrics"""
        self.stats["total_requests"] += 1

        if status_code == 200:
            self.stats["successful_requests"] += 1
        elif status_code == 429:
            self.stats["throttled_requests"] += 1
        else:
            self.stats["failed_requests"] += 1

        self.stats["total_response_time"] += response_time

        # Extract endpoint from URL
        endpoint = url.split("?")[0].replace("https://graph.microsoft.com/v1.0/", "")
        self.stats["endpoints"][endpoint] += 1

    def get_report(self):
        """Generate performance report"""
        avg_response_time = (
            self.stats["total_response_time"] / self.stats["total_requests"]
            if self.stats["total_requests"] > 0 else 0
        )

        return {
            "total_requests": self.stats["total_requests"],
            "successful_requests": self.stats["successful_requests"],
            "failed_requests": self.stats["failed_requests"],
            "throttled_requests": self.stats["throttled_requests"],
            "success_rate": f"{(self.stats['successful_requests'] / self.stats['total_requests'] * 100):.2f}%"
                           if self.stats['total_requests'] > 0 else "0%",
            "throttle_rate": f"{(self.stats['throttled_requests'] / self.stats['total_requests'] * 100):.2f}%"
                            if self.stats['total_requests'] > 0 else "0%",
            "average_response_time": f"{avg_response_time:.3f}s",
            "top_endpoints": dict(sorted(self.stats["endpoints"].items(),
                                        key=lambda x: x[1], reverse=True)[:5])
        }

# Usage
monitor = GraphAPIMonitor()

def monitored_request(url, headers):
    start_time = time.time()
    response = requests.get(url, headers=headers)
    response_time = time.time() - start_time

    monitor.record_request(url, response.status_code, response_time)

    response.raise_for_status()
    return response.json()

# ... make requests ...

# Get report
report = monitor.get_report()
print(json.dumps(report, indent=2))
```

---

## Troubleshooting

### Common Issues and Solutions

#### 1. Authentication Failures

**Issue**: `401 Unauthorized` errors

**Solutions**:
- Verify client ID, tenant ID, and secret are correct
- Check token expiration and refresh if needed
- Ensure proper authority URL format
- Verify the application has been granted admin consent

```python
# Debug authentication
def debug_auth(tenant_id, client_id, client_secret):
    print(f"Tenant ID: {tenant_id[:8]}...")
    print(f"Client ID: {client_id[:8]}...")
    print(f"Secret length: {len(client_secret)}")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    print(f"Authority: {authority}")

    try:
        app = ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

        if "access_token" in result:
            print("✓ Authentication successful")
            print(f"Token expires in: {result.get('expires_in')} seconds")
        else:
            print("✗ Authentication failed")
            print(f"Error: {result.get('error')}")
            print(f"Description: {result.get('error_description')}")
    except Exception as e:
        print(f"✗ Exception: {e}")
```

#### 2. Permission Errors

**Issue**: `403 Forbidden` errors

**Solutions**:
- Verify required permissions are added in app registration
- Ensure admin consent has been granted
- Check if using correct permission type (Application vs Delegated)
- Wait 5-10 minutes after granting consent for changes to propagate

```python
# Check required permissions
def verify_permissions(access_token):
    """Decode JWT token to check permissions"""
    import jwt
    import base64

    # Decode without verification (for debugging only)
    decoded = jwt.decode(access_token, options={"verify_signature": False})

    print("Token Details:")
    print(f"  App ID: {decoded.get('appid')}")
    print(f"  Tenant: {decoded.get('tid')}")
    print(f"  Roles: {decoded.get('roles', [])}")
    print(f"  Scopes: {decoded.get('scp', 'None')}")
```

#### 3. Rate Limiting

**Issue**: `429 Too Many Requests` errors

**Solutions**:
- Implement exponential backoff
- Use batch requests where possible
- Implement request caching
- Reduce concurrent requests
- Use delta queries for incremental updates

#### 4. Timeout Errors

**Issue**: Requests timing out

**Solutions**:
- Increase request timeout
- Use pagination with smaller page sizes
- Implement async/parallel processing
- Use `$select` to reduce payload size

```python
# Configure timeout
response = requests.get(url, headers=headers, timeout=60)  # 60 second timeout
```

#### 5. Pagination Issues

**Issue**: Not retrieving all results

**Solutions**:
- Always check for `@odata.nextLink`
- Continue until no nextLink is present
- Handle pagination in a loop

```python
def get_all_results(url, headers):
    """Properly handle pagination"""
    all_results = []
    page_count = 0

    while url:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        data = response.json()

        results = data.get("value", [])
        all_results.extend(results)

        page_count += 1
        print(f"Page {page_count}: Retrieved {len(results)} items (Total: {len(all_results)})")

        url = data.get("@odata.nextLink")

    return all_results
```

#### 6. Data Inconsistencies

**Issue**: Data appears outdated or inconsistent

**Solutions**:
- Check device last sync time
- Verify data is being properly updated in Intune
- Use delta queries to get latest changes
- Consider eventual consistency in distributed systems

#### 7. Large Dataset Performance

**Issue**: Slow performance with large datasets

**Solutions**:
- Use `$filter` to narrow results
- Implement pagination with appropriate page size
- Use parallel processing
- Cache frequently accessed data
- Use `$select` to retrieve only needed fields

### Debugging Tools

#### cURL Command Generation

```python
def generate_curl_command(url, headers, method="GET"):
    """Generate cURL command for debugging"""
    curl = f"curl -X {method} '{url}' \\\n"
    for key, value in headers.items():
        if key.lower() != "authorization":  # Don't print full token
            curl += f"  -H '{key}: {value}' \\\n"
        else:
            curl += f"  -H 'Authorization: Bearer [REDACTED]' \\\n"
    return curl

# Usage
print(generate_curl_command(
    "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices",
    headers
))
```

#### Request/Response Logging

```python
import logging
import json

# Enable debug logging
logging.basicConfig(level=logging.DEBUG)

# Log requests and responses
def log_request_response(response):
    """Log full request and response details"""
    request = response.request

    logger.debug("=" * 80)
    logger.debug(f"REQUEST: {request.method} {request.url}")
    logger.debug(f"Headers: {dict(request.headers)}")
    if request.body:
        logger.debug(f"Body: {request.body}")

    logger.debug("-" * 80)
    logger.debug(f"RESPONSE: {response.status_code}")
    logger.debug(f"Headers: {dict(response.headers)}")
    try:
        logger.debug(f"Body: {json.dumps(response.json(), indent=2)}")
    except:
        logger.debug(f"Body: {response.text}")
    logger.debug("=" * 80)

# Usage
response = requests.get(url, headers=headers)
log_request_response(response)
```

---

## Additional Resources

### Official Documentation

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/)
- [Intune Graph API Reference](https://docs.microsoft.com/en-us/graph/api/resources/intune-graph-overview)
- [Microsoft Graph SDKs](https://docs.microsoft.com/en-us/graph/sdks/sdks-overview)
- [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) - Interactive API testing tool
- [Microsoft Graph Best Practices](https://docs.microsoft.com/en-us/graph/best-practices-concept)

### Authentication and Security

- [Microsoft Identity Platform Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/)
- [MSAL Python Documentation](https://msal-python.readthedocs.io/)
- [OAuth 2.0 Client Credentials Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-client-creds-grant-flow)

### PowerShell Resources

- [Microsoft.Graph PowerShell Module](https://docs.microsoft.com/en-us/powershell/microsoftgraph/)
- [Intune PowerShell Samples](https://github.com/microsoftgraph/powershell-intune-samples)

### Community Resources

- [Microsoft Graph Community Calls](https://aka.ms/microsoftgraphcall)
- [Microsoft Tech Community - Intune](https://techcommunity.microsoft.com/t5/microsoft-intune/bd-p/Microsoft-Intune)
- [Stack Overflow - microsoft-graph](https://stackoverflow.com/questions/tagged/microsoft-graph)

### Sample Code Repositories

- [Microsoft Graph Python Samples](https://github.com/microsoftgraph/msgraph-sdk-python)
- [Intune PowerShell Samples](https://github.com/microsoftgraph/powershell-intune-samples)
- [Microsoft Graph Training](https://github.com/microsoftgraph/msgraph-training)

### Tools

- **Graph Explorer**: Interactive tool to test Graph API calls
  - URL: https://developer.microsoft.com/graph/graph-explorer

- **Postman Collection**: Pre-built API requests for testing
  - [Microsoft Graph Postman Collection](https://www.postman.com/microsoftgraph)

- **VSCode Extension**: Microsoft Graph extension for Visual Studio Code
  - Search for "Microsoft Graph" in VSCode extensions

### Rate Limiting Documentation

- [Microsoft Graph Throttling Guidance](https://docs.microsoft.com/en-us/graph/throttling)
- [Best Practices for Avoiding Throttling](https://docs.microsoft.com/en-us/graph/throttling#best-practices-to-avoid-throttling)

### Intune Reporting Specifics

- [Intune Reports Documentation](https://docs.microsoft.com/en-us/mem/intune/fundamentals/reports)
- [Intune Data Warehouse](https://docs.microsoft.com/en-us/mem/intune/developer/reports-nav-create-intune-reports)
- [Intune PowerBI Integration](https://docs.microsoft.com/en-us/mem/intune/fundamentals/reports-powerbi)

---

## Appendix: Quick Reference

### Common HTTP Status Codes

- `200 OK` - Request succeeded
- `201 Created` - Resource created successfully
- `204 No Content` - Request succeeded, no content to return
- `400 Bad Request` - Invalid request syntax
- `401 Unauthorized` - Authentication required or failed
- `403 Forbidden` - Insufficient permissions
- `404 Not Found` - Resource not found
- `429 Too Many Requests` - Rate limit exceeded
- `500 Internal Server Error` - Server error
- `503 Service Unavailable` - Service temporarily unavailable

### Useful OData Operators

- `eq` - Equals: `$filter=complianceState eq 'compliant'`
- `ne` - Not equals: `$filter=operatingSystem ne 'Android'`
- `gt` - Greater than: `$filter=totalStorageSpaceInBytes gt 100000000`
- `lt` - Less than: `$filter=lastSyncDateTime lt 2024-01-01T00:00:00Z`
- `and` - Logical AND: `$filter=operatingSystem eq 'Windows' and complianceState eq 'compliant'`
- `or` - Logical OR: `$filter=operatingSystem eq 'iOS' or operatingSystem eq 'Android'`
- `startswith` - Starts with: `$filter=startswith(deviceName,'LAPTOP')`
- `endswith` - Ends with: `$filter=endswith(userPrincipalName,'@contoso.com')`
- `contains` - Contains: `$filter=contains(deviceName,'TEST')`

### Environment Variables Template

```bash
# .env file template
AZURE_TENANT_ID=your-tenant-id
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
GRAPH_API_VERSION=v1.0
```

### Python Requirements

```txt
# requirements.txt
msal>=1.20.0
requests>=2.28.0
python-dotenv>=0.20.0
```

---

## License

This guide is provided as-is for educational and reference purposes.

## Contributions

For questions, issues, or contributions to this guide, please refer to the repository documentation.

---

**Last Updated**: 2026-02-04
**Version**: 1.0
