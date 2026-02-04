# Microsoft Graph API Authentication Module

A comprehensive, production-ready TypeScript authentication module for Microsoft Graph API with support for multiple authentication flows, token caching, and multi-tenant scenarios.

## Features

- **Multiple Authentication Flows**
  - Client Credentials Flow (service/daemon apps)
  - Device Code Flow (headless/CLI scenarios)
  - Interactive Authentication (browser-based user login)

- **Token Management**
  - Automatic token caching with file-based persistence
  - Automatic token refresh
  - Configurable cache location
  - Token expiration handling with 5-minute buffer

- **Multi-Tenant Support**
  - Manage multiple Azure AD tenants
  - Per-tenant token caching
  - Default tenant configuration

- **Production Features**
  - Comprehensive error handling with custom error types
  - TypeScript type safety
  - Configurable logging levels
  - Certificate-based authentication support
  - Secure cache file permissions (0600)

## Installation

```bash
npm install @azure/msal-node
```

## Quick Start

### 1. Client Credentials Flow (Service/Daemon)

Best for: Background services, automated scripts, server-to-server communication

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './graph-auth';

const authManager = createClientCredentialsAuth({
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  tenantId: 'your-tenant-id',
  scopes: ['https://graph.microsoft.com/.default']
});

// Get access token
const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

// Use token to call Graph API
const response = await fetch('https://graph.microsoft.com/v1.0/users', {
  headers: { 'Authorization': `Bearer ${token}` }
});
```

### 2. Device Code Flow (Headless/CLI)

Best for: CLI tools, headless servers, devices without browser

```typescript
import { createDeviceCodeAuth, AuthFlowType } from './graph-auth';

const authManager = createDeviceCodeAuth({
  clientId: 'your-client-id',
  tenantId: 'your-tenant-id',
  scopes: ['https://graph.microsoft.com/User.Read'],
  deviceCodeCallback: (response) => {
    console.log(`Go to ${response.verificationUri} and enter code: ${response.userCode}`);
  }
});

const token = await authManager.getAccessToken(AuthFlowType.DeviceCode);
```

### 3. Interactive Authentication (Browser)

Best for: Desktop applications, user-facing apps

```typescript
import { createInteractiveAuth, AuthFlowType } from './graph-auth';

const authManager = createInteractiveAuth({
  clientId: 'your-client-id',
  tenantId: 'your-tenant-id',
  scopes: ['https://graph.microsoft.com/User.Read'],
  redirectUri: 'http://localhost:3000'
});

const tokenResponse = await authManager.authenticateInteractively();
console.log('Authenticated as:', tokenResponse.account?.username);
```

## Configuration

### Base Configuration

All authentication flows support these base options:

```typescript
interface BaseAuthConfig {
  clientId: string;              // Required: Azure AD application (client) ID
  tenantId: string;              // Required: Azure AD tenant ID
  authority?: string;            // Optional: Custom authority URL
  cloudInstance?: string;        // Optional: Cloud instance (default: login.microsoftonline.com)
  scopes?: string[];            // Optional: Permission scopes (default: ['.default'])
  cacheEnabled?: boolean;       // Optional: Enable token caching (default: true)
  cacheLocation?: string;       // Optional: Cache file path (default: ~/.graph-auth-cache.json)
  logLevel?: LogLevel;          // Optional: Logging level (default: Warning)
}
```

### Client Credentials Configuration

```typescript
interface ClientCredentialsConfig extends BaseAuthConfig {
  clientSecret?: string;              // Client secret
  certificateThumbprint?: string;     // Certificate thumbprint
  certificatePrivateKey?: string;     // Certificate private key
  certificatePath?: string;           // Path to certificate file
}
```

### Device Code Configuration

```typescript
interface DeviceCodeConfig extends BaseAuthConfig {
  deviceCodeCallback?: (response: DeviceCodeResponse) => void;  // Custom code display
  timeout?: number;                                              // Authentication timeout (ms)
}
```

### Interactive Authentication Configuration

```typescript
interface InteractiveAuthConfig extends BaseAuthConfig {
  redirectUri?: string;       // Redirect URI (default: http://localhost)
  successTemplate?: string;   // Custom success page HTML
  errorTemplate?: string;     // Custom error page HTML
  port?: number;             // Local server port
}
```

## Usage Examples

### Example 1: Using Environment Variables

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './graph-auth';

const authManager = createClientCredentialsAuth({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!,
  scopes: ['https://graph.microsoft.com/.default']
});

const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);
```

### Example 2: Certificate-Based Authentication

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './graph-auth';

const authManager = createClientCredentialsAuth({
  clientId: 'your-client-id',
  tenantId: 'your-tenant-id',
  certificatePath: '/path/to/certificate.pem',
  scopes: ['https://graph.microsoft.com/.default']
});

const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);
```

### Example 3: Custom Device Code Display

```typescript
import { createDeviceCodeAuth, AuthFlowType } from './graph-auth';

const authManager = createDeviceCodeAuth({
  clientId: 'your-client-id',
  tenantId: 'your-tenant-id',
  scopes: ['https://graph.microsoft.com/User.Read'],
  deviceCodeCallback: (response) => {
    // Send code via email, SMS, or display in UI
    console.log('='.repeat(60));
    console.log('AUTHENTICATION REQUIRED');
    console.log('='.repeat(60));
    console.log(`\nVisit: ${response.verificationUri}`);
    console.log(`Code: ${response.userCode}`);
    console.log(`\nExpires in ${response.expiresIn} seconds`);
  },
  timeout: 300000 // 5 minutes
});

const token = await authManager.getAccessToken(AuthFlowType.DeviceCode);
```

### Example 4: Multi-Tenant Application

```typescript
import { MultiTenantAuthManager, AuthFlowType } from './graph-auth';

// Create multi-tenant manager
const manager = new MultiTenantAuthManager('default-tenant-id');

// Add tenants
manager.addTenant('tenant1', {
  clientId: 'client-id-1',
  clientSecret: 'secret-1',
  tenantId: 'tenant1'
});

manager.addTenant('tenant2', {
  clientId: 'client-id-2',
  clientSecret: 'secret-2',
  tenantId: 'tenant2'
});

// Get tokens for different tenants
const token1 = await manager.getAccessToken('tenant1', AuthFlowType.ClientCredentials);
const token2 = await manager.getAccessToken('tenant2', AuthFlowType.ClientCredentials);

// Use default tenant
const defaultToken = await manager.getDefaultAccessToken(AuthFlowType.ClientCredentials);
```

### Example 5: Token Refresh and Cache Management

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './graph-auth';

const authManager = createClientCredentialsAuth({
  clientId: 'your-client-id',
  clientSecret: 'your-client-secret',
  tenantId: 'your-tenant-id',
  cacheEnabled: true,
  cacheLocation: './custom-cache/.tokens.json'
});

// First call - authenticates and caches
const token1 = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

// Second call - uses cached token
const token2 = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

// Check if token is cached and valid
const hasValid = await authManager.hasValidToken();

// Force token refresh
const refreshed = await authManager.refreshToken(AuthFlowType.ClientCredentials);

// Clear cache
await authManager.clearCache();
```

### Example 6: Making Graph API Calls

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './graph-auth';

const authManager = createClientCredentialsAuth({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!
});

async function getUsers() {
  const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

  const response = await fetch('https://graph.microsoft.com/v1.0/users?$top=10', {
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
  }

  const data = await response.json();
  return data.value;
}

async function createUser(user: any) {
  const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

  const response = await fetch('https://graph.microsoft.com/v1.0/users', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(user)
  });

  if (!response.ok) {
    throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
  }

  return response.json();
}

// Usage
const users = await getUsers();
console.log(`Found ${users.length} users`);
```

## Error Handling

The module provides custom error types for different scenarios:

```typescript
import {
  GraphAuthError,
  ConfigurationError,
  TokenAcquisitionError,
  TokenRefreshError,
  CacheError
} from './graph-auth';

try {
  const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);
} catch (error) {
  if (error instanceof ConfigurationError) {
    console.error('Invalid configuration:', error.message);
  } else if (error instanceof TokenAcquisitionError) {
    console.error('Failed to get token:', error.message);
    console.error('Error code:', error.code);
    console.error('Details:', error.details);
  } else if (error instanceof TokenRefreshError) {
    console.error('Failed to refresh token:', error.message);
  } else if (error instanceof CacheError) {
    console.error('Cache error:', error.message);
  } else if (error instanceof GraphAuthError) {
    console.error('Authentication error:', error.message);
  } else {
    console.error('Unexpected error:', error);
  }
}
```

## Azure AD App Registration

### Prerequisites

1. **Register an Azure AD Application**
   - Go to [Azure Portal](https://portal.azure.com)
   - Navigate to Azure Active Directory > App registrations
   - Click "New registration"
   - Enter name and click "Register"

2. **Configure Authentication**

   For **Client Credentials Flow**:
   - Go to "Certificates & secrets"
   - Create a new client secret
   - Copy the secret value (shown only once)

   For **Device Code Flow** or **Interactive**:
   - Go to "Authentication"
   - Click "Add a platform"
   - Select "Mobile and desktop applications"
   - Add redirect URI (e.g., `http://localhost`)
   - Enable "Allow public client flows"

3. **Set API Permissions**
   - Go to "API permissions"
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Application permissions" (for Client Credentials) or "Delegated permissions" (for Device Code/Interactive)
   - Add required permissions (e.g., User.Read.All, Directory.Read.All)
   - Click "Grant admin consent"

4. **Get Application IDs**
   - Copy "Application (client) ID"
   - Copy "Directory (tenant) ID"

### Required Permissions

Common permission scopes:

**Delegated Permissions** (user context):
- `User.Read` - Read signed-in user's profile
- `User.ReadBasic.All` - Read all users' basic profiles
- `Mail.Read` - Read user's mail
- `Calendars.Read` - Read user's calendars

**Application Permissions** (app-only context):
- `User.Read.All` - Read all users' full profiles
- `Directory.Read.All` - Read directory data
- `Device.Read.All` - Read all devices
- `DeviceManagementManagedDevices.Read.All` - Read Intune devices

## Environment Variables

Create a `.env` file:

```bash
# Azure AD Configuration
AZURE_CLIENT_ID=your-client-id
AZURE_CLIENT_SECRET=your-client-secret
AZURE_TENANT_ID=your-tenant-id

# Optional: Custom cache location
TOKEN_CACHE_LOCATION=./cache/.tokens.json

# Optional: Scopes
GRAPH_SCOPES=https://graph.microsoft.com/.default
```

Load in your application:

```typescript
import * as dotenv from 'dotenv';
dotenv.config();

const authManager = createClientCredentialsAuth({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!,
  cacheLocation: process.env.TOKEN_CACHE_LOCATION
});
```

## Security Best Practices

1. **Never commit secrets to source control**
   - Use environment variables
   - Add `.env` to `.gitignore`
   - Use Azure Key Vault in production

2. **Use certificate authentication for production**
   ```typescript
   const authManager = createClientCredentialsAuth({
     clientId: 'client-id',
     tenantId: 'tenant-id',
     certificatePath: process.env.CERT_PATH,
     scopes: ['https://graph.microsoft.com/.default']
   });
   ```

3. **Secure cache files**
   - Cache files are created with 0600 permissions (owner read/write only)
   - Store cache in secure location
   - Clear cache when no longer needed

4. **Use least privilege**
   - Request only required permissions
   - Use delegated permissions when possible
   - Regularly audit permissions

5. **Handle token expiration**
   - Tokens are automatically refreshed from cache
   - 5-minute buffer before expiration
   - Implement retry logic for API calls

## Troubleshooting

### Common Issues

**1. "AADSTS7000215: Invalid client secret"**
- Solution: Verify client secret is correct and not expired
- Create new client secret in Azure Portal

**2. "AADSTS650052: The app needs access to a service"**
- Solution: Grant admin consent for required permissions
- Go to API permissions > Grant admin consent

**3. "AADSTS70011: Invalid scope"**
- Solution: Verify scope format and permissions
- Use `https://graph.microsoft.com/.default` for client credentials

**4. Cache permission errors**
- Solution: Ensure write permissions for cache directory
- Check cache file permissions (should be 0600)

**5. Token expired errors**
- Solution: Clear cache and re-authenticate
```typescript
await authManager.clearCache();
const token = await authManager.getAccessToken(flowType);
```

### Debug Logging

Enable verbose logging:

```typescript
import { LogLevel } from '@azure/msal-node';

const authManager = createClientCredentialsAuth({
  clientId: 'client-id',
  clientSecret: 'client-secret',
  tenantId: 'tenant-id',
  logLevel: LogLevel.Verbose
});
```

## API Reference

### Classes

- `GraphAuthManager` - Main authentication manager
- `MultiTenantAuthManager` - Multi-tenant authentication manager
- `TokenCacheManager` - Token cache management

### Functions

- `createClientCredentialsAuth(config)` - Create client credentials auth manager
- `createDeviceCodeAuth(config)` - Create device code auth manager
- `createInteractiveAuth(config)` - Create interactive auth manager
- `getClientCredentialsToken(config)` - Quick helper for client credentials token
- `getDeviceCodeToken(config)` - Quick helper for device code token
- `getInteractiveToken(config)` - Quick helper for interactive token

### Error Classes

- `GraphAuthError` - Base authentication error
- `ConfigurationError` - Configuration validation errors
- `TokenAcquisitionError` - Token acquisition failures
- `TokenRefreshError` - Token refresh failures
- `CacheError` - Cache operation errors

## Testing

Example test setup:

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './graph-auth';

describe('Graph Authentication', () => {
  it('should authenticate with client credentials', async () => {
    const authManager = createClientCredentialsAuth({
      clientId: process.env.TEST_CLIENT_ID!,
      clientSecret: process.env.TEST_CLIENT_SECRET!,
      tenantId: process.env.TEST_TENANT_ID!,
      cacheEnabled: false // Disable cache for tests
    });

    const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);
    expect(token).toBeTruthy();
    expect(typeof token).toBe('string');
  });

  it('should cache tokens', async () => {
    const authManager = createClientCredentialsAuth({
      clientId: process.env.TEST_CLIENT_ID!,
      clientSecret: process.env.TEST_CLIENT_SECRET!,
      tenantId: process.env.TEST_TENANT_ID!,
      cacheEnabled: true,
      cacheLocation: './test-cache/.tokens.json'
    });

    const token1 = await authManager.getAccessToken(AuthFlowType.ClientCredentials);
    const token2 = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

    expect(token1).toBe(token2);

    // Cleanup
    await authManager.clearCache();
  });
});
```

## Performance Considerations

- **Token caching** reduces authentication overhead
- **5-minute expiration buffer** prevents expired token errors
- **Parallel tenant authentication** supported
- **Minimal dependencies** - only @azure/msal-node required

## License

MIT

## Contributing

Contributions welcome! Please follow TypeScript best practices and include tests.

## Support

For issues and questions:
- Microsoft Graph API: https://docs.microsoft.com/graph
- MSAL Node: https://github.com/AzureAD/microsoft-authentication-library-for-js
- Azure AD: https://docs.microsoft.com/azure/active-directory

## Changelog

### v1.0.0
- Initial release
- Client credentials flow
- Device code flow
- Interactive authentication
- Token caching
- Multi-tenant support
- Comprehensive error handling
