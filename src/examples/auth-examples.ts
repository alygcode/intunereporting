/**
 * Comprehensive Examples for Microsoft Graph API Authentication Module
 *
 * This file demonstrates all authentication flows and features.
 * Run examples with: ts-node src/examples/auth-examples.ts
 */

import * as dotenv from 'dotenv';
import {
  GraphAuthManager,
  MultiTenantAuthManager,
  createClientCredentialsAuth,
  createDeviceCodeAuth,
  createInteractiveAuth,
  getClientCredentialsToken,
  AuthFlowType,
  ClientCredentialsConfig,
  DeviceCodeConfig,
  InteractiveAuthConfig,
  GraphAuthError,
  ConfigurationError,
  TokenAcquisitionError
} from '../auth/graph-auth';

// Load environment variables
dotenv.config();

// ============================================================================
// Example 1: Client Credentials Flow - Basic Usage
// ============================================================================

async function example1_BasicClientCredentials() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 1: Client Credentials Flow - Basic Usage');
  console.log('='.repeat(70) + '\n');

  try {
    const config: ClientCredentialsConfig = {
      clientId: process.env.AZURE_CLIENT_ID || 'your-client-id',
      clientSecret: process.env.AZURE_CLIENT_SECRET || 'your-client-secret',
      tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id',
      scopes: ['https://graph.microsoft.com/.default']
    };

    console.log('Authenticating with client credentials...');
    const authManager = createClientCredentialsAuth(config);

    const token = await authManager.getAccessToken(
      AuthFlowType.ClientCredentials,
      config
    );

    console.log('✓ Successfully authenticated!');
    console.log(`Token preview: ${token.substring(0, 50)}...`);

    // Make a Graph API call
    console.log('\nFetching organization details...');
    const response = await fetch('https://graph.microsoft.com/v1.0/organization', {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (response.ok) {
      const data = await response.json();
      const org = data.value[0];
      console.log('✓ Organization retrieved:');
      console.log(`  Display Name: ${org.displayName}`);
      console.log(`  Tenant ID: ${org.id}`);
      console.log(`  Verified Domains: ${org.verifiedDomains?.length || 0}`);
    }

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 2: Client Credentials with Token Caching
// ============================================================================

async function example2_ClientCredentialsWithCaching() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 2: Client Credentials with Token Caching');
  console.log('='.repeat(70) + '\n');

  try {
    const config: ClientCredentialsConfig = {
      clientId: process.env.AZURE_CLIENT_ID || 'your-client-id',
      clientSecret: process.env.AZURE_CLIENT_SECRET || 'your-client-secret',
      tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id',
      scopes: ['https://graph.microsoft.com/.default'],
      cacheEnabled: true,
      cacheLocation: './cache/.graph-tokens.json'
    };

    const authManager = createClientCredentialsAuth(config);

    // First call - will authenticate
    console.log('1. First authentication (will be cached)...');
    const startTime1 = Date.now();
    const token1 = await authManager.getAccessToken(
      AuthFlowType.ClientCredentials,
      config
    );
    const duration1 = Date.now() - startTime1;
    console.log(`   ✓ Token obtained in ${duration1}ms`);

    // Second call - will use cache
    console.log('\n2. Second call (should use cache)...');
    const startTime2 = Date.now();
    const token2 = await authManager.getAccessToken(
      AuthFlowType.ClientCredentials,
      config
    );
    const duration2 = Date.now() - startTime2;
    console.log(`   ✓ Token obtained in ${duration2}ms`);
    console.log(`   Cache speedup: ${Math.round((duration1 / duration2))}x faster`);
    console.log(`   Same token: ${token1 === token2}`);

    // Check cache status
    const hasValid = await authManager.hasValidToken();
    console.log(`\n3. Has valid cached token: ${hasValid}`);

    // Clear cache
    console.log('\n4. Clearing cache...');
    await authManager.clearCache();
    console.log('   ✓ Cache cleared');

    const hasValidAfter = await authManager.hasValidToken();
    console.log(`   Has valid cached token after clear: ${hasValidAfter}`);

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 3: Device Code Flow
// ============================================================================

async function example3_DeviceCodeFlow() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 3: Device Code Flow');
  console.log('='.repeat(70) + '\n');

  try {
    const config: DeviceCodeConfig = {
      clientId: process.env.AZURE_CLIENT_ID || 'your-client-id',
      tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id',
      scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/User.ReadBasic.All'
      ],
      cacheEnabled: true,
      deviceCodeCallback: (response) => {
        console.log('\n┌─────────────────────────────────────────────────────┐');
        console.log('│          AUTHENTICATION REQUIRED                    │');
        console.log('└─────────────────────────────────────────────────────┘');
        console.log('\n📱 To sign in, open a web browser and navigate to:');
        console.log(`\n   ${response.verificationUri}`);
        console.log(`\n🔑 Enter the code: ${response.userCode}`);
        console.log(`\n⏰ This code expires in ${Math.floor(response.expiresIn / 60)} minutes`);
        console.log('\n⌛ Waiting for you to complete sign-in...\n');
      },
      timeout: 300000 // 5 minutes
    };

    console.log('Initiating device code authentication...');
    const authManager = createDeviceCodeAuth(config);

    const tokenResponse = await authManager.authenticateWithDeviceCode(config);

    console.log('✓ Authentication successful!');
    console.log(`\n📧 Authenticated as: ${tokenResponse.account?.username}`);
    console.log(`👤 Name: ${tokenResponse.account?.name}`);
    console.log(`🏢 Tenant: ${tokenResponse.account?.tenantId}`);
    console.log(`⏰ Token expires: ${tokenResponse.expiresOn.toLocaleString()}`);

    // Get user profile
    console.log('\nFetching user profile...');
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${tokenResponse.accessToken}`
      }
    });

    if (response.ok) {
      const user = await response.json();
      console.log('✓ User profile retrieved:');
      console.log(`  Display Name: ${user.displayName}`);
      console.log(`  Email: ${user.mail || user.userPrincipalName}`);
      console.log(`  Job Title: ${user.jobTitle || 'N/A'}`);
      console.log(`  Office: ${user.officeLocation || 'N/A'}`);
    }

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 4: Interactive Authentication
// ============================================================================

async function example4_InteractiveAuth() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 4: Interactive Authentication');
  console.log('='.repeat(70) + '\n');

  try {
    const config: InteractiveAuthConfig = {
      clientId: process.env.AZURE_CLIENT_ID || 'your-client-id',
      tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id',
      scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Mail.Read',
        'https://graph.microsoft.com/Calendars.Read'
      ],
      redirectUri: 'http://localhost:3000',
      cacheEnabled: true
    };

    console.log('Opening browser for authentication...');
    const authManager = createInteractiveAuth(config);

    const tokenResponse = await authManager.authenticateInteractively(config);

    console.log('✓ Authentication successful!');
    console.log(`\nAuthenticated as: ${tokenResponse.account?.username}`);
    console.log(`Token expires: ${tokenResponse.expiresOn.toLocaleString()}`);

    // Get user's recent emails
    console.log('\nFetching recent emails...');
    const response = await fetch(
      'https://graph.microsoft.com/v1.0/me/messages?$top=5&$select=subject,from,receivedDateTime',
      {
        headers: {
          'Authorization': `Bearer ${tokenResponse.accessToken}`
        }
      }
    );

    if (response.ok) {
      const data = await response.json();
      console.log(`✓ Retrieved ${data.value.length} recent emails:`);
      data.value.forEach((email: any, index: number) => {
        console.log(`\n  ${index + 1}. ${email.subject}`);
        console.log(`     From: ${email.from.emailAddress.name}`);
        console.log(`     Received: ${new Date(email.receivedDateTime).toLocaleString()}`);
      });
    }

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 5: Multi-Tenant Management
// ============================================================================

async function example5_MultiTenant() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 5: Multi-Tenant Management');
  console.log('='.repeat(70) + '\n');

  try {
    // Create multi-tenant manager
    const manager = new MultiTenantAuthManager('production-tenant');

    // Add production tenant
    manager.addTenant('production-tenant', {
      clientId: process.env.PROD_CLIENT_ID || 'prod-client-id',
      clientSecret: process.env.PROD_CLIENT_SECRET || 'prod-secret',
      tenantId: 'production-tenant',
      scopes: ['https://graph.microsoft.com/.default']
    } as ClientCredentialsConfig);

    // Add staging tenant
    manager.addTenant('staging-tenant', {
      clientId: process.env.STAGING_CLIENT_ID || 'staging-client-id',
      clientSecret: process.env.STAGING_CLIENT_SECRET || 'staging-secret',
      tenantId: 'staging-tenant',
      scopes: ['https://graph.microsoft.com/.default']
    } as ClientCredentialsConfig);

    // Add development tenant
    manager.addTenant('dev-tenant', {
      clientId: process.env.DEV_CLIENT_ID || 'dev-client-id',
      clientSecret: process.env.DEV_CLIENT_SECRET || 'dev-secret',
      tenantId: 'dev-tenant',
      scopes: ['https://graph.microsoft.com/.default']
    } as ClientCredentialsConfig);

    console.log('Configured Tenants:');
    manager.getTenants().forEach((tenant, index) => {
      console.log(`  ${index + 1}. ${tenant}`);
    });

    // Get tokens for different tenants
    console.log('\nAuthenticating with multiple tenants...\n');

    for (const tenant of manager.getTenants()) {
      try {
        console.log(`Authenticating with ${tenant}...`);
        const token = await manager.getAccessToken(
          tenant,
          AuthFlowType.ClientCredentials
        );
        console.log(`✓ ${tenant}: Token obtained (${token.substring(0, 30)}...)`);

        // Get user count for each tenant
        const response = await fetch(
          'https://graph.microsoft.com/v1.0/users/$count',
          {
            headers: {
              'Authorization': `Bearer ${token}`,
              'ConsistencyLevel': 'eventual'
            }
          }
        );

        if (response.ok) {
          const count = await response.text();
          console.log(`  └─ Users in tenant: ${count}`);
        }
      } catch (error) {
        console.log(`✗ ${tenant}: Authentication failed`);
      }
    }

    // Use default tenant
    console.log('\nUsing default tenant (production)...');
    const defaultToken = await manager.getDefaultAccessToken(
      AuthFlowType.ClientCredentials
    );
    console.log(`✓ Default tenant token: ${defaultToken.substring(0, 30)}...`);

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 6: Token Refresh
// ============================================================================

async function example6_TokenRefresh() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 6: Token Refresh');
  console.log('='.repeat(70) + '\n');

  try {
    const config: ClientCredentialsConfig = {
      clientId: process.env.AZURE_CLIENT_ID || 'your-client-id',
      clientSecret: process.env.AZURE_CLIENT_SECRET || 'your-client-secret',
      tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id',
      scopes: ['https://graph.microsoft.com/.default'],
      cacheEnabled: true
    };

    const authManager = createClientCredentialsAuth(config);

    // Get initial token
    console.log('1. Getting initial token...');
    const initialResponse = await authManager.authenticateWithClientCredentials(config);
    console.log(`   ✓ Token obtained`);
    console.log(`   Expires: ${initialResponse.expiresOn.toLocaleString()}`);

    // Check cache
    const hasValid1 = await authManager.hasValidToken();
    console.log(`\n2. Has valid cached token: ${hasValid1}`);

    // Simulate token expiration by clearing cache
    console.log('\n3. Simulating token expiration (clearing cache)...');
    await authManager.clearCache();

    const hasValid2 = await authManager.hasValidToken();
    console.log(`   Has valid cached token: ${hasValid2}`);

    // Refresh token
    console.log('\n4. Refreshing token...');
    const refreshedResponse = await authManager.refreshToken(
      AuthFlowType.ClientCredentials,
      config
    );
    console.log(`   ✓ Token refreshed`);
    console.log(`   New expiration: ${refreshedResponse.expiresOn.toLocaleString()}`);

    // Verify new token works
    console.log('\n5. Verifying new token...');
    const response = await fetch('https://graph.microsoft.com/v1.0/organization', {
      headers: {
        'Authorization': `Bearer ${refreshedResponse.accessToken}`
      }
    });

    if (response.ok) {
      console.log('   ✓ New token is valid');
    } else {
      console.log('   ✗ New token failed validation');
    }

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 7: Practical Intune Device Reporting
// ============================================================================

async function example7_IntuneDeviceReporting() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 7: Practical Intune Device Reporting');
  console.log('='.repeat(70) + '\n');

  try {
    const config: ClientCredentialsConfig = {
      clientId: process.env.AZURE_CLIENT_ID || 'your-client-id',
      clientSecret: process.env.AZURE_CLIENT_SECRET || 'your-client-secret',
      tenantId: process.env.AZURE_TENANT_ID || 'your-tenant-id',
      scopes: ['https://graph.microsoft.com/.default'],
      cacheEnabled: true
    };

    console.log('Authenticating...');
    const token = await getClientCredentialsToken(config);
    console.log('✓ Authenticated\n');

    // Fetch managed devices
    console.log('Fetching Intune managed devices...');
    const devicesResponse = await fetch(
      'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=10&$select=deviceName,operatingSystem,complianceState,enrolledDateTime,lastSyncDateTime,managedDeviceOwnerType',
      {
        headers: {
          'Authorization': `Bearer ${token}`,
          'Content-Type': 'application/json'
        }
      }
    );

    if (!devicesResponse.ok) {
      throw new Error(`Graph API error: ${devicesResponse.status} ${devicesResponse.statusText}`);
    }

    const devicesData = await devicesResponse.json();
    const devices = devicesData.value;

    console.log(`✓ Retrieved ${devices.length} managed devices\n`);

    // Display device summary
    const summary = {
      total: devices.length,
      compliant: devices.filter((d: any) => d.complianceState === 'compliant').length,
      nonCompliant: devices.filter((d: any) => d.complianceState === 'noncompliant').length,
      byOS: {} as Record<string, number>,
      byOwnership: {} as Record<string, number>
    };

    devices.forEach((device: any) => {
      summary.byOS[device.operatingSystem] = (summary.byOS[device.operatingSystem] || 0) + 1;
      summary.byOwnership[device.managedDeviceOwnerType] = (summary.byOwnership[device.managedDeviceOwnerType] || 0) + 1;
    });

    console.log('📊 Device Summary:');
    console.log(`   Total Devices: ${summary.total}`);
    console.log(`   Compliant: ${summary.compliant}`);
    console.log(`   Non-Compliant: ${summary.nonCompliant}`);

    console.log('\n📱 Devices by Operating System:');
    Object.entries(summary.byOS).forEach(([os, count]) => {
      console.log(`   ${os}: ${count}`);
    });

    console.log('\n👥 Devices by Ownership:');
    Object.entries(summary.byOwnership).forEach(([type, count]) => {
      console.log(`   ${type}: ${count}`);
    });

    console.log('\n📝 Device Details:');
    devices.slice(0, 5).forEach((device: any, index: number) => {
      console.log(`\n   ${index + 1}. ${device.deviceName}`);
      console.log(`      OS: ${device.operatingSystem}`);
      console.log(`      Compliance: ${device.complianceState}`);
      console.log(`      Enrolled: ${new Date(device.enrolledDateTime).toLocaleDateString()}`);
      console.log(`      Last Sync: ${new Date(device.lastSyncDateTime).toLocaleString()}`);
      console.log(`      Owner: ${device.managedDeviceOwnerType}`);
    });

  } catch (error) {
    handleError(error);
  }
}

// ============================================================================
// Example 8: Error Handling Patterns
// ============================================================================

async function example8_ErrorHandling() {
  console.log('\n' + '='.repeat(70));
  console.log('EXAMPLE 8: Error Handling Patterns');
  console.log('='.repeat(70) + '\n');

  // Pattern 1: Configuration validation
  console.log('1. Testing configuration validation...');
  try {
    const badManager = new GraphAuthManager({
      clientId: '',  // Invalid
      tenantId: ''   // Invalid
    });
  } catch (error) {
    if (error instanceof ConfigurationError) {
      console.log(`   ✓ Configuration error caught: ${error.message}`);
    }
  }

  // Pattern 2: Token acquisition error
  console.log('\n2. Testing token acquisition error...');
  try {
    const config: ClientCredentialsConfig = {
      clientId: 'invalid-client-id',
      clientSecret: 'invalid-secret',
      tenantId: 'invalid-tenant',
      cacheEnabled: false
    };

    const authManager = createClientCredentialsAuth(config);
    await authManager.getAccessToken(AuthFlowType.ClientCredentials, config);
  } catch (error) {
    if (error instanceof TokenAcquisitionError) {
      console.log(`   ✓ Token acquisition error caught: ${error.message}`);
      if (error.code) {
        console.log(`   Error code: ${error.code}`);
      }
    }
  }

  // Pattern 3: Graceful error handling wrapper
  console.log('\n3. Graceful error handling wrapper...');

  const authenticate = async () => {
    try {
      const config: ClientCredentialsConfig = {
        clientId: process.env.AZURE_CLIENT_ID || '',
        clientSecret: process.env.AZURE_CLIENT_SECRET || '',
        tenantId: process.env.AZURE_TENANT_ID || ''
      };

      const token = await getClientCredentialsToken(config);
      return { success: true, token };
    } catch (error) {
      if (error instanceof ConfigurationError) {
        return { success: false, error: 'Configuration error', message: error.message };
      } else if (error instanceof TokenAcquisitionError) {
        return { success: false, error: 'Authentication failed', message: error.message, code: error.code };
      } else if (error instanceof GraphAuthError) {
        return { success: false, error: 'Graph auth error', message: error.message };
      } else {
        return { success: false, error: 'Unknown error', message: String(error) };
      }
    }
  };

  const result = await authenticate();
  console.log(`   Authentication result: ${result.success ? 'Success' : 'Failed'}`);
  if (!result.success) {
    console.log(`   Error: ${result.error}`);
    console.log(`   Message: ${result.message}`);
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

function handleError(error: unknown) {
  console.error('\n❌ Error occurred:\n');

  if (error instanceof ConfigurationError) {
    console.error(`Configuration Error: ${error.message}`);
    console.error(`Code: ${error.code}`);
  } else if (error instanceof TokenAcquisitionError) {
    console.error(`Token Acquisition Error: ${error.message}`);
    console.error(`Code: ${error.code}`);
    if (error.details) {
      console.error(`Details: ${JSON.stringify(error.details, null, 2)}`);
    }
  } else if (error instanceof GraphAuthError) {
    console.error(`Authentication Error: ${error.message}`);
    console.error(`Code: ${error.code}`);
  } else if (error instanceof Error) {
    console.error(`Error: ${error.message}`);
    console.error(error.stack);
  } else {
    console.error(`Unknown error: ${String(error)}`);
  }
}

// ============================================================================
// Main Menu
// ============================================================================

async function main() {
  console.log('\n' + '='.repeat(70));
  console.log('   Microsoft Graph API Authentication - Interactive Examples');
  console.log('='.repeat(70));

  const args = process.argv.slice(2);

  if (args.length === 0) {
    console.log('\nAvailable examples:');
    console.log('  1. Basic Client Credentials');
    console.log('  2. Client Credentials with Caching');
    console.log('  3. Device Code Flow');
    console.log('  4. Interactive Authentication');
    console.log('  5. Multi-Tenant Management');
    console.log('  6. Token Refresh');
    console.log('  7. Practical Intune Device Reporting');
    console.log('  8. Error Handling Patterns');
    console.log('\nUsage: ts-node auth-examples.ts [example-number]');
    console.log('Example: ts-node auth-examples.ts 1\n');
    return;
  }

  const exampleNum = parseInt(args[0]);

  switch (exampleNum) {
    case 1:
      await example1_BasicClientCredentials();
      break;
    case 2:
      await example2_ClientCredentialsWithCaching();
      break;
    case 3:
      await example3_DeviceCodeFlow();
      break;
    case 4:
      await example4_InteractiveAuth();
      break;
    case 5:
      await example5_MultiTenant();
      break;
    case 6:
      await example6_TokenRefresh();
      break;
    case 7:
      await example7_IntuneDeviceReporting();
      break;
    case 8:
      await example8_ErrorHandling();
      break;
    default:
      console.log(`\nInvalid example number: ${exampleNum}`);
      console.log('Please choose a number between 1 and 8\n');
  }
}

// Run examples
if (require.main === module) {
  main().catch(console.error);
}

export {
  example1_BasicClientCredentials,
  example2_ClientCredentialsWithCaching,
  example3_DeviceCodeFlow,
  example4_InteractiveAuth,
  example5_MultiTenant,
  example6_TokenRefresh,
  example7_IntuneDeviceReporting,
  example8_ErrorHandling
};
