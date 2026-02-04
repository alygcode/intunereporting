# Authentication Module - Quick Start Guide

Get started with the Microsoft Graph API authentication module in just a few minutes.

## Prerequisites

- Node.js 18.0 or higher
- An Azure AD (Microsoft Entra ID) tenant
- Administrator access to create app registrations

## Step 1: Install Dependencies

```bash
npm install
```

This installs `@azure/msal-node` and all other required dependencies.

## Step 2: Azure AD App Registration

### Create the App

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Azure Active Directory** → **App registrations**
3. Click **+ New registration**
4. Name: `Intune Reporting App`
5. Click **Register**

### Copy IDs

From the app overview page:
- Copy **Application (client) ID** → Use for `AZURE_CLIENT_ID`
- Copy **Directory (tenant) ID** → Use for `AZURE_TENANT_ID`

### Create Client Secret

1. Go to **Certificates & secrets**
2. Click **+ New client secret**
3. Description: `Reporting Secret`
4. Expiration: 12 months
5. Click **Add**
6. **COPY THE VALUE NOW** (shown only once) → Use for `AZURE_CLIENT_SECRET`

### Add API Permissions

1. Go to **API permissions** → **+ Add a permission**
2. Select **Microsoft Graph** → **Application permissions**
3. Add these permissions:
   - `DeviceManagementManagedDevices.Read.All`
   - `DeviceManagementConfiguration.Read.All`
   - `Directory.Read.All`
   - `User.Read.All`
4. Click **Grant admin consent for [Your Organization]**
5. Click **Yes**

## Step 3: Configure Environment

```bash
# Copy the example environment file
cp .env.example .env

# Edit with your credentials
nano .env
```

Add your values to `.env`:

```env
AZURE_TENANT_ID=your-tenant-id-here
AZURE_CLIENT_ID=your-client-id-here
AZURE_CLIENT_SECRET=your-secret-here
```

## Step 4: Test Authentication

### Run Basic Example

```bash
npm run build
ts-node src/examples/auth-examples.ts 1
```

Expected output:
```
=== CLIENT CREDENTIALS FLOW EXAMPLE ===

Authenticating with client credentials...
✓ Successfully authenticated!
Token preview: eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6Ij...
```

## Usage Examples

### Client Credentials (Recommended for Services)

```typescript
import { createClientCredentialsAuth, AuthFlowType } from './src/auth/graph-auth';
import * as dotenv from 'dotenv';

dotenv.config();

const authManager = createClientCredentialsAuth({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!
});

const token = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

// Call Graph API
const response = await fetch('https://graph.microsoft.com/v1.0/users', {
  headers: { 'Authorization': `Bearer ${token}` }
});
```

### Device Code Flow (For CLI Tools)

```typescript
import { createDeviceCodeAuth, AuthFlowType } from './src/auth/graph-auth';

const authManager = createDeviceCodeAuth({
  clientId: process.env.AZURE_CLIENT_ID!,
  tenantId: process.env.AZURE_TENANT_ID!,
  scopes: ['https://graph.microsoft.com/User.Read']
});

const token = await authManager.getAccessToken(AuthFlowType.DeviceCode);
// User will see instructions to authenticate via browser
```

### With Token Caching

```typescript
const authManager = createClientCredentialsAuth({
  clientId: process.env.AZURE_CLIENT_ID!,
  clientSecret: process.env.AZURE_CLIENT_SECRET!,
  tenantId: process.env.AZURE_TENANT_ID!,
  cacheEnabled: true,  // Tokens are cached automatically
  cacheLocation: './cache/.tokens.json'
});

// First call - authenticates
const token1 = await authManager.getAccessToken(AuthFlowType.ClientCredentials);

// Second call - uses cache (much faster!)
const token2 = await authManager.getAccessToken(AuthFlowType.ClientCredentials);
```

## Run All Examples

```bash
# Example 1: Basic client credentials
ts-node src/examples/auth-examples.ts 1

# Example 2: Client credentials with caching
ts-node src/examples/auth-examples.ts 2

# Example 3: Device code flow
ts-node src/examples/auth-examples.ts 3

# Example 4: Interactive authentication
ts-node src/examples/auth-examples.ts 4

# Example 5: Multi-tenant support
ts-node src/examples/auth-examples.ts 5

# Example 6: Token refresh
ts-node src/examples/auth-examples.ts 6

# Example 7: Practical Intune reporting
ts-node src/examples/auth-examples.ts 7

# Example 8: Error handling
ts-node src/examples/auth-examples.ts 8
```

## Common Issues

### Invalid client secret
- Verify secret in `.env` matches Azure Portal
- Check if secret expired
- Create new secret if needed

### Missing permissions
- Go to API permissions in Azure Portal
- Click "Grant admin consent"

### Cache errors
- Ensure write permissions on cache directory
- Try: `rm ~/.graph-auth-cache.json`

## Security Tips

1. **Never commit secrets**
   - Add `.env` to `.gitignore`
   - Use environment variables

2. **Use certificates in production**
   ```typescript
   const authManager = createClientCredentialsAuth({
     clientId: 'client-id',
     tenantId: 'tenant-id',
     certificatePath: '/secure/path/cert.pem'
   });
   ```

3. **Rotate secrets regularly**
   - Set expiration dates
   - Rotate before they expire

## Next Steps

- 📖 Read [Full Documentation](/src/auth/README.md)
- 🔍 Explore [Complete Examples](/src/examples/auth-examples.ts)
- 🚀 Build your integration
- 📊 Try Intune reporting (Example 7)

## Resources

- [Microsoft Graph Docs](https://docs.microsoft.com/graph)
- [MSAL Node Docs](https://github.com/AzureAD/microsoft-authentication-library-for-js)
- [Graph Explorer](https://developer.microsoft.com/graph/graph-explorer)

---

**Ready to authenticate!** 🚀

For detailed documentation, see `/src/auth/README.md`
