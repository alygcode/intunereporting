/**
 * Microsoft Graph API Authentication Module
 *
 * Supports multiple authentication flows:
 * - Client Credentials Flow (service/daemon apps)
 * - Device Code Flow (headless/CLI scenarios)
 * - Interactive Authentication (user login)
 *
 * Features:
 * - Token caching and automatic refresh
 * - Multi-tenant support
 * - Comprehensive error handling
 * - Type-safe TypeScript implementation
 */

import {
  ConfidentialClientApplication,
  PublicClientApplication,
  ClientCredentialRequest,
  DeviceCodeRequest,
  InteractiveRequest,
  AuthenticationResult,
  AccountInfo,
  Configuration,
  LogLevel,
  CachePlugin,
  TokenCacheContext,
  INetworkModule,
  NetworkRequestOptions,
  NetworkResponse
} from '@azure/msal-node';
import * as fs from 'fs/promises';
import * as path from 'path';
import * as crypto from 'crypto';

// ============================================================================
// Types and Interfaces
// ============================================================================

/**
 * Authentication flow types
 */
export enum AuthFlowType {
  ClientCredentials = 'client_credentials',
  DeviceCode = 'device_code',
  Interactive = 'interactive'
}

/**
 * Base configuration for all authentication flows
 */
export interface BaseAuthConfig {
  clientId: string;
  tenantId: string;
  authority?: string;
  cloudInstance?: 'https://login.microsoftonline.com' | 'https://login.microsoftonline.us' | 'https://login.partner.microsoftonline.cn';
  scopes?: string[];
  cacheEnabled?: boolean;
  cacheLocation?: string;
  logLevel?: LogLevel;
}

/**
 * Configuration for client credentials flow
 */
export interface ClientCredentialsConfig extends BaseAuthConfig {
  clientSecret?: string;
  certificateThumbprint?: string;
  certificatePrivateKey?: string;
  certificatePath?: string;
}

/**
 * Configuration for device code flow
 */
export interface DeviceCodeConfig extends BaseAuthConfig {
  deviceCodeCallback?: (deviceCodeResponse: DeviceCodeResponse) => void;
  timeout?: number;
}

/**
 * Configuration for interactive authentication
 */
export interface InteractiveAuthConfig extends BaseAuthConfig {
  redirectUri?: string;
  successTemplate?: string;
  errorTemplate?: string;
  port?: number;
}

/**
 * Device code response structure
 */
export interface DeviceCodeResponse {
  userCode: string;
  deviceCode: string;
  verificationUri: string;
  expiresIn: number;
  interval: number;
  message: string;
}

/**
 * Token response structure
 */
export interface TokenResponse {
  accessToken: string;
  tokenType: string;
  expiresOn: Date;
  scopes: string[];
  account?: AccountInfo;
  idToken?: string;
}

/**
 * Token cache entry
 */
interface CacheEntry {
  accessToken: string;
  expiresOn: number;
  scopes: string[];
  tenantId: string;
  account?: AccountInfo;
}

/**
 * Cache storage structure
 */
interface CacheStorage {
  [key: string]: CacheEntry;
}

// ============================================================================
// Custom Error Classes
// ============================================================================

/**
 * Base authentication error
 */
export class GraphAuthError extends Error {
  constructor(message: string, public code?: string, public details?: any) {
    super(message);
    this.name = 'GraphAuthError';
    Object.setPrototypeOf(this, GraphAuthError.prototype);
  }
}

/**
 * Configuration error
 */
export class ConfigurationError extends GraphAuthError {
  constructor(message: string, details?: any) {
    super(message, 'CONFIGURATION_ERROR', details);
    this.name = 'ConfigurationError';
    Object.setPrototypeOf(this, ConfigurationError.prototype);
  }
}

/**
 * Token acquisition error
 */
export class TokenAcquisitionError extends GraphAuthError {
  constructor(message: string, code?: string, details?: any) {
    super(message, code, details);
    this.name = 'TokenAcquisitionError';
    Object.setPrototypeOf(this, TokenAcquisitionError.prototype);
  }
}

/**
 * Token refresh error
 */
export class TokenRefreshError extends GraphAuthError {
  constructor(message: string, details?: any) {
    super(message, 'TOKEN_REFRESH_ERROR', details);
    this.name = 'TokenRefreshError';
    Object.setPrototypeOf(this, TokenRefreshError.prototype);
  }
}

/**
 * Cache error
 */
export class CacheError extends GraphAuthError {
  constructor(message: string, details?: any) {
    super(message, 'CACHE_ERROR', details);
    this.name = 'CacheError';
    Object.setPrototypeOf(this, CacheError.prototype);
  }
}

// ============================================================================
// Token Cache Manager
// ============================================================================

/**
 * Manages token caching with file-based persistence
 */
export class TokenCacheManager {
  private cacheLocation: string;
  private cache: CacheStorage = {};
  private cacheLoaded: boolean = false;

  constructor(cacheLocation?: string) {
    this.cacheLocation = cacheLocation || path.join(
      process.env.HOME || process.env.USERPROFILE || '.',
      '.graph-auth-cache.json'
    );
  }

  /**
   * Load cache from disk
   */
  async loadCache(): Promise<void> {
    try {
      const data = await fs.readFile(this.cacheLocation, 'utf-8');
      this.cache = JSON.parse(data);
      this.cacheLoaded = true;
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        throw new CacheError('Failed to load cache', error);
      }
      // File doesn't exist yet, start with empty cache
      this.cache = {};
      this.cacheLoaded = true;
    }
  }

  /**
   * Save cache to disk
   */
  async saveCache(): Promise<void> {
    try {
      const dir = path.dirname(this.cacheLocation);
      await fs.mkdir(dir, { recursive: true });
      await fs.writeFile(this.cacheLocation, JSON.stringify(this.cache, null, 2), {
        mode: 0o600 // Restrict permissions
      });
    } catch (error) {
      throw new CacheError('Failed to save cache', error);
    }
  }

  /**
   * Generate cache key
   */
  private getCacheKey(tenantId: string, scopes: string[], clientId: string): string {
    const scopeString = scopes.sort().join(',');
    return crypto
      .createHash('sha256')
      .update(`${tenantId}:${clientId}:${scopeString}`)
      .digest('hex');
  }

  /**
   * Get token from cache
   */
  async getToken(
    tenantId: string,
    scopes: string[],
    clientId: string
  ): Promise<CacheEntry | null> {
    if (!this.cacheLoaded) {
      await this.loadCache();
    }

    const key = this.getCacheKey(tenantId, scopes, clientId);
    const entry = this.cache[key];

    if (!entry) {
      return null;
    }

    // Check if token is expired (with 5 minute buffer)
    const now = Date.now();
    const expiresOn = entry.expiresOn - (5 * 60 * 1000);

    if (now >= expiresOn) {
      // Token expired, remove from cache
      delete this.cache[key];
      await this.saveCache();
      return null;
    }

    return entry;
  }

  /**
   * Store token in cache
   */
  async setToken(
    tenantId: string,
    scopes: string[],
    clientId: string,
    token: TokenResponse
  ): Promise<void> {
    if (!this.cacheLoaded) {
      await this.loadCache();
    }

    const key = this.getCacheKey(tenantId, scopes, clientId);
    this.cache[key] = {
      accessToken: token.accessToken,
      expiresOn: token.expiresOn.getTime(),
      scopes: token.scopes,
      tenantId,
      account: token.account
    };

    await this.saveCache();
  }

  /**
   * Clear all cached tokens
   */
  async clearCache(): Promise<void> {
    this.cache = {};
    try {
      await fs.unlink(this.cacheLocation);
    } catch (error: any) {
      if (error.code !== 'ENOENT') {
        throw new CacheError('Failed to clear cache', error);
      }
    }
  }

  /**
   * Clear tokens for specific tenant
   */
  async clearTenantCache(tenantId: string): Promise<void> {
    if (!this.cacheLoaded) {
      await this.loadCache();
    }

    const keys = Object.keys(this.cache);
    for (const key of keys) {
      if (this.cache[key].tenantId === tenantId) {
        delete this.cache[key];
      }
    }

    await this.saveCache();
  }
}

// ============================================================================
// Graph Authentication Manager
// ============================================================================

/**
 * Main authentication manager for Microsoft Graph API
 */
export class GraphAuthManager {
  private config: BaseAuthConfig;
  private cacheManager: TokenCacheManager | null = null;
  private msalClient: ConfidentialClientApplication | PublicClientApplication | null = null;

  constructor(config: BaseAuthConfig) {
    this.validateConfig(config);
    this.config = {
      ...config,
      scopes: config.scopes || ['https://graph.microsoft.com/.default'],
      cacheEnabled: config.cacheEnabled !== false,
      cloudInstance: config.cloudInstance || 'https://login.microsoftonline.com'
    };

    if (this.config.cacheEnabled) {
      this.cacheManager = new TokenCacheManager(config.cacheLocation);
    }
  }

  /**
   * Validate base configuration
   */
  private validateConfig(config: BaseAuthConfig): void {
    if (!config.clientId) {
      throw new ConfigurationError('clientId is required');
    }
    if (!config.tenantId) {
      throw new ConfigurationError('tenantId is required');
    }
  }

  /**
   * Get authority URL
   */
  private getAuthority(): string {
    if (this.config.authority) {
      return this.config.authority;
    }
    return `${this.config.cloudInstance}/${this.config.tenantId}`;
  }

  /**
   * Build MSAL configuration
   */
  private buildMsalConfig(isPublic: boolean = false): Configuration {
    const authority = this.getAuthority();

    return {
      auth: {
        clientId: this.config.clientId,
        authority: authority
      },
      system: {
        loggerOptions: {
          loggerCallback: (level, message, containsPii) => {
            if (containsPii) return;
            if (this.config.logLevel && level <= this.config.logLevel) {
              console.log(`[MSAL] ${message}`);
            }
          },
          piiLoggingEnabled: false,
          logLevel: this.config.logLevel || LogLevel.Warning
        }
      }
    };
  }

  /**
   * Convert MSAL result to TokenResponse
   */
  private toTokenResponse(result: AuthenticationResult): TokenResponse {
    return {
      accessToken: result.accessToken,
      tokenType: result.tokenType || 'Bearer',
      expiresOn: result.expiresOn || new Date(Date.now() + 3600000),
      scopes: result.scopes || [],
      account: result.account || undefined,
      idToken: result.idToken || undefined
    };
  }

  // ============================================================================
  // Client Credentials Flow
  // ============================================================================

  /**
   * Authenticate using client credentials flow (service/daemon apps)
   */
  async authenticateWithClientCredentials(
    config: ClientCredentialsConfig
  ): Promise<TokenResponse> {
    try {
      // Check cache first
      if (this.cacheManager && this.config.cacheEnabled) {
        const cached = await this.cacheManager.getToken(
          config.tenantId,
          config.scopes || this.config.scopes!,
          config.clientId
        );
        if (cached) {
          return {
            accessToken: cached.accessToken,
            tokenType: 'Bearer',
            expiresOn: new Date(cached.expiresOn),
            scopes: cached.scopes,
            account: cached.account
          };
        }
      }

      // Validate credentials
      if (!config.clientSecret && !config.certificateThumbprint && !config.certificatePath) {
        throw new ConfigurationError(
          'Either clientSecret, certificateThumbprint, or certificatePath is required for client credentials flow'
        );
      }

      // Build MSAL configuration
      const msalConfig = this.buildMsalConfig();

      // Add client credentials
      if (config.clientSecret) {
        msalConfig.auth.clientSecret = config.clientSecret;
      } else if (config.certificateThumbprint && config.certificatePrivateKey) {
        msalConfig.auth.clientCertificate = {
          thumbprint: config.certificateThumbprint,
          privateKey: config.certificatePrivateKey
        };
      } else if (config.certificatePath) {
        const certData = await fs.readFile(config.certificatePath, 'utf-8');
        // Parse certificate (simplified - in production, use proper certificate parsing)
        msalConfig.auth.clientCertificate = {
          thumbprint: '', // Extract from certificate
          privateKey: certData
        };
      }

      // Create confidential client
      const client = new ConfidentialClientApplication(msalConfig);

      // Request token
      const request: ClientCredentialRequest = {
        scopes: config.scopes || this.config.scopes!,
        skipCache: false
      };

      const result = await client.acquireTokenByClientCredential(request);

      if (!result) {
        throw new TokenAcquisitionError('Failed to acquire token - no result returned');
      }

      const tokenResponse = this.toTokenResponse(result);

      // Cache token
      if (this.cacheManager && this.config.cacheEnabled) {
        await this.cacheManager.setToken(
          config.tenantId,
          tokenResponse.scopes,
          config.clientId,
          tokenResponse
        );
      }

      return tokenResponse;
    } catch (error: any) {
      if (error instanceof GraphAuthError) {
        throw error;
      }
      throw new TokenAcquisitionError(
        'Failed to authenticate with client credentials',
        error.errorCode,
        error
      );
    }
  }

  // ============================================================================
  // Device Code Flow
  // ============================================================================

  /**
   * Authenticate using device code flow (headless/CLI scenarios)
   */
  async authenticateWithDeviceCode(
    config: DeviceCodeConfig
  ): Promise<TokenResponse> {
    try {
      // Check cache first
      if (this.cacheManager && this.config.cacheEnabled) {
        const cached = await this.cacheManager.getToken(
          config.tenantId,
          config.scopes || this.config.scopes!,
          config.clientId
        );
        if (cached) {
          return {
            accessToken: cached.accessToken,
            tokenType: 'Bearer',
            expiresOn: new Date(cached.expiresOn),
            scopes: cached.scopes,
            account: cached.account
          };
        }
      }

      // Build MSAL configuration
      const msalConfig = this.buildMsalConfig(true);

      // Create public client
      const client = new PublicClientApplication(msalConfig);

      // Request token with device code
      const request: DeviceCodeRequest = {
        scopes: config.scopes || this.config.scopes!,
        deviceCodeCallback: (response) => {
          const deviceCodeResponse: DeviceCodeResponse = {
            userCode: response.userCode,
            deviceCode: response.deviceCode,
            verificationUri: response.verificationUri,
            expiresIn: response.expiresIn,
            interval: response.interval,
            message: response.message
          };

          if (config.deviceCodeCallback) {
            config.deviceCodeCallback(deviceCodeResponse);
          } else {
            // Default callback - print to console
            console.log('\n' + '='.repeat(60));
            console.log('DEVICE CODE AUTHENTICATION');
            console.log('='.repeat(60));
            console.log(response.message);
            console.log('='.repeat(60) + '\n');
          }
        },
        timeout: config.timeout
      };

      const result = await client.acquireTokenByDeviceCode(request);

      if (!result) {
        throw new TokenAcquisitionError('Failed to acquire token - no result returned');
      }

      const tokenResponse = this.toTokenResponse(result);

      // Cache token
      if (this.cacheManager && this.config.cacheEnabled) {
        await this.cacheManager.setToken(
          config.tenantId,
          tokenResponse.scopes,
          config.clientId,
          tokenResponse
        );
      }

      return tokenResponse;
    } catch (error: any) {
      if (error instanceof GraphAuthError) {
        throw error;
      }
      throw new TokenAcquisitionError(
        'Failed to authenticate with device code',
        error.errorCode,
        error
      );
    }
  }

  // ============================================================================
  // Interactive Authentication
  // ============================================================================

  /**
   * Authenticate using interactive flow (user login with browser)
   */
  async authenticateInteractively(
    config: InteractiveAuthConfig
  ): Promise<TokenResponse> {
    try {
      // Check cache first
      if (this.cacheManager && this.config.cacheEnabled) {
        const cached = await this.cacheManager.getToken(
          config.tenantId,
          config.scopes || this.config.scopes!,
          config.clientId
        );
        if (cached) {
          return {
            accessToken: cached.accessToken,
            tokenType: 'Bearer',
            expiresOn: new Date(cached.expiresOn),
            scopes: cached.scopes,
            account: cached.account
          };
        }
      }

      // Build MSAL configuration
      const msalConfig = this.buildMsalConfig(true);

      // Create public client
      const client = new PublicClientApplication(msalConfig);

      // Request token interactively
      const request: InteractiveRequest = {
        scopes: config.scopes || this.config.scopes!,
        redirectUri: config.redirectUri || 'http://localhost',
        successTemplate: config.successTemplate,
        errorTemplate: config.errorTemplate
      };

      const result = await client.acquireTokenInteractive(request);

      if (!result) {
        throw new TokenAcquisitionError('Failed to acquire token - no result returned');
      }

      const tokenResponse = this.toTokenResponse(result);

      // Cache token
      if (this.cacheManager && this.config.cacheEnabled) {
        await this.cacheManager.setToken(
          config.tenantId,
          tokenResponse.scopes,
          config.clientId,
          tokenResponse
        );
      }

      return tokenResponse;
    } catch (error: any) {
      if (error instanceof GraphAuthError) {
        throw error;
      }
      throw new TokenAcquisitionError(
        'Failed to authenticate interactively',
        error.errorCode,
        error
      );
    }
  }

  // ============================================================================
  // Token Management
  // ============================================================================

  /**
   * Get a valid access token (from cache or acquire new one)
   */
  async getAccessToken(
    flowType: AuthFlowType = AuthFlowType.ClientCredentials,
    flowConfig?: ClientCredentialsConfig | DeviceCodeConfig | InteractiveAuthConfig
  ): Promise<string> {
    const config = { ...this.config, ...flowConfig };

    let tokenResponse: TokenResponse;

    switch (flowType) {
      case AuthFlowType.ClientCredentials:
        tokenResponse = await this.authenticateWithClientCredentials(
          config as ClientCredentialsConfig
        );
        break;
      case AuthFlowType.DeviceCode:
        tokenResponse = await this.authenticateWithDeviceCode(
          config as DeviceCodeConfig
        );
        break;
      case AuthFlowType.Interactive:
        tokenResponse = await this.authenticateInteractively(
          config as InteractiveAuthConfig
        );
        break;
      default:
        throw new ConfigurationError(`Unsupported authentication flow: ${flowType}`);
    }

    return tokenResponse.accessToken;
  }

  /**
   * Refresh token (re-authenticate)
   */
  async refreshToken(
    flowType: AuthFlowType,
    flowConfig?: ClientCredentialsConfig | DeviceCodeConfig | InteractiveAuthConfig
  ): Promise<TokenResponse> {
    try {
      // Clear cache for this config
      if (this.cacheManager) {
        await this.cacheManager.clearTenantCache(this.config.tenantId);
      }

      // Re-authenticate
      const config = { ...this.config, ...flowConfig };

      switch (flowType) {
        case AuthFlowType.ClientCredentials:
          return await this.authenticateWithClientCredentials(
            config as ClientCredentialsConfig
          );
        case AuthFlowType.DeviceCode:
          return await this.authenticateWithDeviceCode(
            config as DeviceCodeConfig
          );
        case AuthFlowType.Interactive:
          return await this.authenticateInteractively(
            config as InteractiveAuthConfig
          );
        default:
          throw new ConfigurationError(`Unsupported authentication flow: ${flowType}`);
      }
    } catch (error: any) {
      throw new TokenRefreshError('Failed to refresh token', error);
    }
  }

  /**
   * Clear all cached tokens
   */
  async clearCache(): Promise<void> {
    if (this.cacheManager) {
      await this.cacheManager.clearCache();
    }
  }

  /**
   * Check if token is valid
   */
  async hasValidToken(): Promise<boolean> {
    if (!this.cacheManager || !this.config.cacheEnabled) {
      return false;
    }

    const cached = await this.cacheManager.getToken(
      this.config.tenantId,
      this.config.scopes!,
      this.config.clientId
    );

    return cached !== null;
  }
}

// ============================================================================
// Helper Functions
// ============================================================================

/**
 * Create authentication manager with client credentials
 */
export function createClientCredentialsAuth(
  config: ClientCredentialsConfig
): GraphAuthManager {
  return new GraphAuthManager(config);
}

/**
 * Create authentication manager with device code
 */
export function createDeviceCodeAuth(
  config: DeviceCodeConfig
): GraphAuthManager {
  return new GraphAuthManager(config);
}

/**
 * Create authentication manager with interactive auth
 */
export function createInteractiveAuth(
  config: InteractiveAuthConfig
): GraphAuthManager {
  return new GraphAuthManager(config);
}

/**
 * Quick helper to get access token with client credentials
 */
export async function getClientCredentialsToken(
  config: ClientCredentialsConfig
): Promise<string> {
  const manager = createClientCredentialsAuth(config);
  return manager.getAccessToken(AuthFlowType.ClientCredentials, config);
}

/**
 * Quick helper to get access token with device code
 */
export async function getDeviceCodeToken(
  config: DeviceCodeConfig
): Promise<string> {
  const manager = createDeviceCodeAuth(config);
  return manager.getAccessToken(AuthFlowType.DeviceCode, config);
}

/**
 * Quick helper to get access token with interactive auth
 */
export async function getInteractiveToken(
  config: InteractiveAuthConfig
): Promise<string> {
  const manager = createInteractiveAuth(config);
  return manager.getAccessToken(AuthFlowType.Interactive, config);
}

// ============================================================================
// Multi-Tenant Support
// ============================================================================

/**
 * Multi-tenant authentication manager
 */
export class MultiTenantAuthManager {
  private managers: Map<string, GraphAuthManager> = new Map();
  private defaultTenantId?: string;

  constructor(defaultTenantId?: string) {
    this.defaultTenantId = defaultTenantId;
  }

  /**
   * Add tenant configuration
   */
  addTenant(tenantId: string, config: BaseAuthConfig): void {
    const tenantConfig = { ...config, tenantId };
    const manager = new GraphAuthManager(tenantConfig);
    this.managers.set(tenantId, manager);

    if (!this.defaultTenantId) {
      this.defaultTenantId = tenantId;
    }
  }

  /**
   * Get authentication manager for tenant
   */
  getTenantManager(tenantId?: string): GraphAuthManager {
    const tid = tenantId || this.defaultTenantId;
    if (!tid) {
      throw new ConfigurationError('No tenant ID specified and no default tenant configured');
    }

    const manager = this.managers.get(tid);
    if (!manager) {
      throw new ConfigurationError(`No authentication manager configured for tenant: ${tid}`);
    }

    return manager;
  }

  /**
   * Get access token for tenant
   */
  async getAccessToken(
    tenantId: string,
    flowType: AuthFlowType,
    flowConfig?: ClientCredentialsConfig | DeviceCodeConfig | InteractiveAuthConfig
  ): Promise<string> {
    const manager = this.getTenantManager(tenantId);
    return manager.getAccessToken(flowType, flowConfig);
  }

  /**
   * Get access token for default tenant
   */
  async getDefaultAccessToken(
    flowType: AuthFlowType,
    flowConfig?: ClientCredentialsConfig | DeviceCodeConfig | InteractiveAuthConfig
  ): Promise<string> {
    return this.getAccessToken(this.defaultTenantId!, flowType, flowConfig);
  }

  /**
   * Clear cache for all tenants
   */
  async clearAllCaches(): Promise<void> {
    for (const manager of this.managers.values()) {
      await manager.clearCache();
    }
  }

  /**
   * Get list of configured tenants
   */
  getTenants(): string[] {
    return Array.from(this.managers.keys());
  }
}

// ============================================================================
// USAGE EXAMPLES
// ============================================================================

/**
 * Example 1: Client Credentials Flow (Service/Daemon App)
 *
 * Use case: Automated scripts, background services, server-to-server communication
 */
export async function exampleClientCredentials() {
  console.log('=== CLIENT CREDENTIALS FLOW EXAMPLE ===\n');

  try {
    // Configuration
    const config: ClientCredentialsConfig = {
      clientId: 'YOUR_CLIENT_ID',
      clientSecret: 'YOUR_CLIENT_SECRET',
      tenantId: 'YOUR_TENANT_ID',
      scopes: ['https://graph.microsoft.com/.default'],
      cacheEnabled: true,
      logLevel: LogLevel.Info
    };

    // Create authentication manager
    const authManager = createClientCredentialsAuth(config);

    // Get access token (will use cache if available)
    const token = await authManager.getAccessToken(
      AuthFlowType.ClientCredentials,
      config
    );

    console.log('Access Token:', token.substring(0, 50) + '...');
    console.log('Token obtained successfully!');

    // Use token to call Microsoft Graph API
    // Example: Get users
    const response = await fetch('https://graph.microsoft.com/v1.0/users', {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    if (response.ok) {
      const data = await response.json();
      console.log(`\nRetrieved ${data.value.length} users from Graph API`);
    }

    // Check if we have a valid cached token
    const hasValidToken = await authManager.hasValidToken();
    console.log('\nHas valid cached token:', hasValidToken);

  } catch (error) {
    if (error instanceof GraphAuthError) {
      console.error(`Authentication Error: ${error.message}`);
      console.error(`Error Code: ${error.code}`);
    } else {
      console.error('Unexpected error:', error);
    }
  }
}

/**
 * Example 2: Device Code Flow (Headless/CLI)
 *
 * Use case: CLI tools, headless servers, devices without web browser
 */
export async function exampleDeviceCode() {
  console.log('=== DEVICE CODE FLOW EXAMPLE ===\n');

  try {
    // Configuration
    const config: DeviceCodeConfig = {
      clientId: 'YOUR_CLIENT_ID',
      tenantId: 'YOUR_TENANT_ID',
      scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Directory.Read.All'
      ],
      cacheEnabled: true,
      deviceCodeCallback: (response) => {
        // Custom callback to display device code
        console.log('\n┌─────────────────────────────────────────────┐');
        console.log('│     MICROSOFT GRAPH AUTHENTICATION         │');
        console.log('└─────────────────────────────────────────────┘');
        console.log('\nTo sign in, use a web browser to open the page:');
        console.log(`\n  ${response.verificationUri}`);
        console.log(`\nAnd enter the code: ${response.userCode}`);
        console.log('\nWaiting for authentication...\n');
      },
      timeout: 120000 // 2 minutes
    };

    // Create authentication manager
    const authManager = createDeviceCodeAuth(config);

    // Get access token - will prompt user to authenticate
    const token = await authManager.getAccessToken(
      AuthFlowType.DeviceCode,
      config
    );

    console.log('✓ Authentication successful!');
    console.log('Access Token:', token.substring(0, 50) + '...\n');

    // Get full token response with account info
    const tokenResponse = await authManager.authenticateWithDeviceCode(config);
    console.log('Authenticated User:', tokenResponse.account?.username);
    console.log('Token Expires:', tokenResponse.expiresOn.toLocaleString());

  } catch (error) {
    if (error instanceof GraphAuthError) {
      console.error(`Authentication Error: ${error.message}`);
    } else {
      console.error('Unexpected error:', error);
    }
  }
}

/**
 * Example 3: Interactive Authentication (Browser-based)
 *
 * Use case: Desktop applications, user-facing applications
 */
export async function exampleInteractive() {
  console.log('=== INTERACTIVE AUTHENTICATION EXAMPLE ===\n');

  try {
    // Configuration
    const config: InteractiveAuthConfig = {
      clientId: 'YOUR_CLIENT_ID',
      tenantId: 'YOUR_TENANT_ID',
      scopes: [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Mail.Read'
      ],
      redirectUri: 'http://localhost:3000',
      cacheEnabled: true
    };

    // Create authentication manager
    const authManager = createInteractiveAuth(config);

    console.log('Opening browser for authentication...');

    // Get access token - will open browser
    const tokenResponse = await authManager.authenticateInteractively(config);

    console.log('✓ Authentication successful!');
    console.log('User:', tokenResponse.account?.username);
    console.log('Access Token:', tokenResponse.accessToken.substring(0, 50) + '...');

    // Call Graph API
    const response = await fetch('https://graph.microsoft.com/v1.0/me', {
      headers: {
        'Authorization': `Bearer ${tokenResponse.accessToken}`
      }
    });

    if (response.ok) {
      const user = await response.json();
      console.log('\nUser Profile:');
      console.log(`  Name: ${user.displayName}`);
      console.log(`  Email: ${user.mail || user.userPrincipalName}`);
      console.log(`  ID: ${user.id}`);
    }

  } catch (error) {
    if (error instanceof GraphAuthError) {
      console.error(`Authentication Error: ${error.message}`);
    } else {
      console.error('Unexpected error:', error);
    }
  }
}

/**
 * Example 4: Multi-Tenant Support
 *
 * Use case: Applications that work with multiple Azure AD tenants
 */
export async function exampleMultiTenant() {
  console.log('=== MULTI-TENANT AUTHENTICATION EXAMPLE ===\n');

  try {
    // Create multi-tenant manager
    const multiTenantManager = new MultiTenantAuthManager();

    // Add tenant configurations
    multiTenantManager.addTenant('tenant1-id', {
      clientId: 'CLIENT_ID_1',
      clientSecret: 'CLIENT_SECRET_1',
      tenantId: 'tenant1-id',
      scopes: ['https://graph.microsoft.com/.default']
    } as ClientCredentialsConfig);

    multiTenantManager.addTenant('tenant2-id', {
      clientId: 'CLIENT_ID_2',
      clientSecret: 'CLIENT_SECRET_2',
      tenantId: 'tenant2-id',
      scopes: ['https://graph.microsoft.com/.default']
    } as ClientCredentialsConfig);

    multiTenantManager.addTenant('tenant3-id', {
      clientId: 'CLIENT_ID_3',
      clientSecret: 'CLIENT_SECRET_3',
      tenantId: 'tenant3-id',
      scopes: ['https://graph.microsoft.com/.default']
    } as ClientCredentialsConfig);

    console.log('Configured Tenants:', multiTenantManager.getTenants());

    // Get tokens for different tenants
    console.log('\nAuthenticating with Tenant 1...');
    const token1 = await multiTenantManager.getAccessToken(
      'tenant1-id',
      AuthFlowType.ClientCredentials
    );
    console.log('✓ Tenant 1 token obtained');

    console.log('\nAuthenticating with Tenant 2...');
    const token2 = await multiTenantManager.getAccessToken(
      'tenant2-id',
      AuthFlowType.ClientCredentials
    );
    console.log('✓ Tenant 2 token obtained');

    // Use tokens for different tenants
    const fetchTenantUsers = async (tenantName: string, token: string) => {
      const response = await fetch('https://graph.microsoft.com/v1.0/users?$top=5', {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      if (response.ok) {
        const data = await response.json();
        console.log(`\n${tenantName}: ${data.value.length} users retrieved`);
      }
    };

    await fetchTenantUsers('Tenant 1', token1);
    await fetchTenantUsers('Tenant 2', token2);

    // Clear all caches
    console.log('\nClearing all tenant caches...');
    await multiTenantManager.clearAllCaches();
    console.log('✓ All caches cleared');

  } catch (error) {
    if (error instanceof GraphAuthError) {
      console.error(`Authentication Error: ${error.message}`);
    } else {
      console.error('Unexpected error:', error);
    }
  }
}

/**
 * Example 5: Token Refresh and Cache Management
 */
export async function exampleTokenManagement() {
  console.log('=== TOKEN MANAGEMENT EXAMPLE ===\n');

  try {
    const config: ClientCredentialsConfig = {
      clientId: 'YOUR_CLIENT_ID',
      clientSecret: 'YOUR_CLIENT_SECRET',
      tenantId: 'YOUR_TENANT_ID',
      scopes: ['https://graph.microsoft.com/.default'],
      cacheEnabled: true,
      cacheLocation: './custom-cache-location/.token-cache.json'
    };

    const authManager = createClientCredentialsAuth(config);

    // First call - will authenticate and cache
    console.log('1. First authentication (will cache token)...');
    const token1 = await authManager.getAccessToken(
      AuthFlowType.ClientCredentials,
      config
    );
    console.log('✓ Token obtained and cached');

    // Second call - will use cached token
    console.log('\n2. Second call (will use cached token)...');
    const token2 = await authManager.getAccessToken(
      AuthFlowType.ClientCredentials,
      config
    );
    console.log('✓ Token retrieved from cache');
    console.log('Same token:', token1 === token2);

    // Check if we have valid token
    const hasValid = await authManager.hasValidToken();
    console.log('\n3. Has valid cached token:', hasValid);

    // Force refresh token
    console.log('\n4. Forcing token refresh...');
    const refreshedToken = await authManager.refreshToken(
      AuthFlowType.ClientCredentials,
      config
    );
    console.log('✓ Token refreshed');
    console.log('New token different from cached:', refreshedToken.accessToken !== token1);

    // Clear cache
    console.log('\n5. Clearing cache...');
    await authManager.clearCache();
    console.log('✓ Cache cleared');

    const hasValidAfterClear = await authManager.hasValidToken();
    console.log('Has valid token after clear:', hasValidAfterClear);

  } catch (error) {
    if (error instanceof GraphAuthError) {
      console.error(`Error: ${error.message}`);
    } else {
      console.error('Unexpected error:', error);
    }
  }
}

/**
 * Example 6: Error Handling
 */
export async function exampleErrorHandling() {
  console.log('=== ERROR HANDLING EXAMPLE ===\n');

  // Example 1: Configuration Error
  try {
    const authManager = new GraphAuthManager({
      clientId: '', // Invalid - empty clientId
      tenantId: 'test-tenant'
    });
  } catch (error) {
    if (error instanceof ConfigurationError) {
      console.log('Configuration Error caught:');
      console.log(`  Message: ${error.message}`);
      console.log(`  Code: ${error.code}\n`);
    }
  }

  // Example 2: Token Acquisition Error
  try {
    const config: ClientCredentialsConfig = {
      clientId: 'invalid-client-id',
      clientSecret: 'invalid-secret',
      tenantId: 'invalid-tenant',
      cacheEnabled: false // Disable cache to force authentication
    };

    const authManager = createClientCredentialsAuth(config);
    await authManager.getAccessToken(AuthFlowType.ClientCredentials, config);
  } catch (error) {
    if (error instanceof TokenAcquisitionError) {
      console.log('Token Acquisition Error caught:');
      console.log(`  Message: ${error.message}`);
      console.log(`  Code: ${error.code}`);
      console.log(`  Details:`, error.details?.message || 'N/A');
      console.log();
    }
  }

  // Example 3: Graceful error handling
  const handleAuthentication = async () => {
    try {
      const config: ClientCredentialsConfig = {
        clientId: process.env.CLIENT_ID || '',
        clientSecret: process.env.CLIENT_SECRET || '',
        tenantId: process.env.TENANT_ID || ''
      };

      const authManager = createClientCredentialsAuth(config);
      const token = await authManager.getAccessToken(
        AuthFlowType.ClientCredentials,
        config
      );
      return { success: true, token };
    } catch (error) {
      if (error instanceof ConfigurationError) {
        return { success: false, error: 'Invalid configuration', details: error.message };
      } else if (error instanceof TokenAcquisitionError) {
        return { success: false, error: 'Authentication failed', details: error.message };
      } else if (error instanceof CacheError) {
        return { success: false, error: 'Cache error', details: error.message };
      } else {
        return { success: false, error: 'Unknown error', details: String(error) };
      }
    }
  };

  const result = await handleAuthentication();
  console.log('Authentication result:');
  console.log(JSON.stringify(result, null, 2));
}

// ============================================================================
// Export all
// ============================================================================

export default GraphAuthManager;
