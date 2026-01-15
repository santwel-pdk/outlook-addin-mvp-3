/**
 * Azure AD Token Service
 *
 * Singleton service for managing Azure AD Client Credentials OAuth 2.0 flow.
 * Acquires access tokens for Web API authentication with automatic refresh.
 *
 * @module azureTokenService
 * @see https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow
 */

import {
  AzureTokenResponse,
  AzureAdConfig,
  CachedAzureToken,
  AzureTokenState,
  AzureAdError
} from '../types/azure.types';
import { logError } from '../utils/errorHandler';

/**
 * Singleton Azure Token Manager class
 * Handles token acquisition, caching, and automatic refresh
 */
class AzureTokenManager {
  private static instance: AzureTokenManager;
  private cachedToken: CachedAzureToken | null = null;
  private refreshPromise: Promise<string> | null = null;
  private refreshTimer: NodeJS.Timeout | null = null;
  private currentConfig: AzureAdConfig | null = null;
  private readonly REFRESH_THRESHOLD_MS = 5 * 60 * 1000; // 5 minutes before expiry

  private constructor() {
    // Private constructor for singleton pattern
  }

  /**
   * Gets the singleton instance
   *
   * @returns {AzureTokenManager} AzureTokenManager instance
   */
  public static getInstance(): AzureTokenManager {
    if (!AzureTokenManager.instance) {
      AzureTokenManager.instance = new AzureTokenManager();
    }
    return AzureTokenManager.instance;
  }

  /**
   * Gets a valid access token, refreshing if necessary
   *
   * @param {AzureAdConfig} config Azure AD configuration
   * @returns {Promise<string>} Valid access token
   * @throws {Error} If token acquisition fails
   */
  public async getToken(config: AzureAdConfig): Promise<string> {
    // Validate configuration
    this.validateConfig(config);
    this.currentConfig = config;

    // Return cached token if still valid (with threshold)
    if (this.cachedToken && this.isTokenValid()) {
      console.log('AzureTokenManager: Using cached token');
      return this.cachedToken.token;
    }

    // CRITICAL: If refresh in progress, await same promise (race condition prevention)
    if (this.refreshPromise) {
      console.log('AzureTokenManager: Waiting for in-progress token refresh');
      try {
        return await this.refreshPromise;
      } catch (error) {
        // If the pending refresh failed, clear it and try again
        this.refreshPromise = null;
        throw error;
      }
    }

    // Start new token acquisition
    console.log('AzureTokenManager: Acquiring new token from Azure AD');
    this.refreshPromise = this.acquireToken(config);

    try {
      const token = await this.refreshPromise;
      this.refreshPromise = null;

      // Schedule next automatic refresh
      this.scheduleNextRefresh(config);

      return token;
    } catch (error) {
      this.refreshPromise = null;
      throw error;
    }
  }

  /**
   * Validates Azure AD configuration
   *
   * @param {AzureAdConfig} config Configuration to validate
   * @throws {Error} If configuration is invalid
   */
  private validateConfig(config: AzureAdConfig): void {
    if (!config.tenantId) {
      throw new Error('Azure AD Tenant ID is required');
    }
    if (!config.clientId) {
      throw new Error('Azure AD Client ID is required');
    }
    if (!config.clientSecret) {
      throw new Error('Azure AD Client Secret is required');
    }
    if (!config.scope) {
      throw new Error('Azure AD Scope is required');
    }
    // CRITICAL: Client Credentials scope must end with "/.default"
    if (!config.scope.endsWith('/.default')) {
      console.warn('AzureTokenManager: Scope should end with "/.default" for Client Credentials flow');
    }
  }

  /**
   * Acquires a new token from Azure AD using Client Credentials flow
   *
   * @param {AzureAdConfig} config Azure AD configuration
   * @returns {Promise<string>} Fresh access token
   * @throws {Error} If token acquisition fails
   */
  private async acquireToken(config: AzureAdConfig): Promise<string> {
    const tokenEndpoint = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`;

    // CRITICAL: Use URLSearchParams, NOT JSON body
    const body = new URLSearchParams({
      client_id: config.clientId,
      client_secret: config.clientSecret,
      scope: config.scope,
      grant_type: 'client_credentials'
    });

    try {
      const response = await fetch(tokenEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: body.toString()
      });

      if (!response.ok) {
        const errorData = await this.parseErrorResponse(response);
        throw new Error(this.formatAzureError(errorData, response.status));
      }

      const tokenResponse: AzureTokenResponse = await response.json();

      // Cache the token with expiry metadata
      const now = Date.now();
      this.cachedToken = {
        token: tokenResponse.access_token,
        expiresAt: now + (tokenResponse.expires_in * 1000),
        acquiredAt: now
      };

      console.log(`AzureTokenManager: Token acquired, expires in ${tokenResponse.expires_in} seconds`);

      return tokenResponse.access_token;
    } catch (error) {
      logError('AzureTokenManager Acquire Token', error);
      throw new Error(`Azure AD authentication failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  }

  /**
   * Parses error response from Azure AD
   *
   * @param {Response} response Fetch response object
   * @returns {Promise<AzureAdError | null>} Parsed error or null
   */
  private async parseErrorResponse(response: Response): Promise<AzureAdError | null> {
    try {
      const text = await response.text();
      return JSON.parse(text) as AzureAdError;
    } catch {
      return null;
    }
  }

  /**
   * Formats Azure AD error for user-friendly display
   *
   * @param {AzureAdError | null} errorData Azure AD error data
   * @param {number} status HTTP status code
   * @returns {string} Formatted error message
   */
  private formatAzureError(errorData: AzureAdError | null, status: number): string {
    if (errorData?.error_description) {
      // Don't expose full error description to users - sanitize it
      if (errorData.error === 'invalid_client') {
        return 'Invalid client credentials. Please verify your Azure AD configuration.';
      }
      if (errorData.error === 'invalid_scope') {
        return 'Invalid scope. Ensure scope ends with "/.default" for Client Credentials flow.';
      }
      if (errorData.error === 'unauthorized_client') {
        return 'Client is not authorized. Please check API permissions in Azure AD.';
      }
      return `Azure AD error: ${errorData.error}`;
    }
    return `Token acquisition failed with status ${status}`;
  }

  /**
   * Checks if the cached token is still valid (with refresh threshold)
   *
   * @returns {boolean} True if token exists and hasn't expired (with 5-min buffer)
   */
  private isTokenValid(): boolean {
    if (!this.cachedToken) {
      return false;
    }

    // Consider invalid if expires within threshold
    return this.cachedToken.expiresAt > Date.now() + this.REFRESH_THRESHOLD_MS;
  }

  /**
   * Schedules the next automatic token refresh
   *
   * @param {AzureAdConfig} config Azure AD configuration for refresh
   */
  private scheduleNextRefresh(config: AzureAdConfig): void {
    // Clear any existing timer
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }

    if (!this.cachedToken) {
      return;
    }

    const timeUntilExpiry = this.cachedToken.expiresAt - Date.now();
    const refreshIn = Math.max(0, timeUntilExpiry - this.REFRESH_THRESHOLD_MS);

    console.log(`AzureTokenManager: Scheduling next refresh in ${Math.round(refreshIn / 1000)} seconds`);

    this.refreshTimer = setTimeout(async () => {
      try {
        await this.getToken(config);
      } catch (error) {
        logError('AzureTokenManager Scheduled Refresh', error);
        // Retry in 1 minute if automatic refresh fails
        setTimeout(() => this.scheduleNextRefresh(config), 60000);
      }
    }, refreshIn);
  }

  /**
   * Starts automatic token refresh monitoring
   *
   * @param {AzureAdConfig} config Azure AD configuration
   */
  public startAutoRefresh(config: AzureAdConfig): void {
    this.currentConfig = config;

    if (!this.cachedToken) {
      console.warn('AzureTokenManager: Cannot start auto-refresh - no cached token');
      return;
    }

    this.scheduleNextRefresh(config);
    console.log('AzureTokenManager: Auto-refresh started');
  }

  /**
   * Stops automatic token refresh monitoring
   */
  public stopAutoRefresh(): void {
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
      console.log('AzureTokenManager: Auto-refresh stopped');
    }
  }

  /**
   * Forces an immediate token refresh
   *
   * @param {AzureAdConfig} config Azure AD configuration
   * @returns {Promise<string>} Fresh access token
   */
  public async forceRefresh(config: AzureAdConfig): Promise<string> {
    // Clear cached token to force new acquisition
    this.cachedToken = null;
    this.refreshPromise = null;

    return await this.getToken(config);
  }

  /**
   * Gets token status information for monitoring/debugging
   *
   * @returns {AzureTokenState} Token status information
   */
  public getTokenStatus(): AzureTokenState {
    const timeUntilExpiry = this.cachedToken
      ? Math.max(0, this.cachedToken.expiresAt - Date.now())
      : null;

    const timeUntilRefresh = timeUntilExpiry !== null
      ? Math.max(0, timeUntilExpiry - this.REFRESH_THRESHOLD_MS)
      : null;

    return {
      isInitialized: this.cachedToken !== null,
      hasValidToken: this.isTokenValid(),
      timeUntilExpiry,
      timeUntilRefresh,
      isRefreshing: this.refreshPromise !== null,
      error: null
    };
  }

  /**
   * Gets the remaining time until token expiry
   *
   * @returns {number | null} Time in milliseconds until expiry, or null if no token
   */
  public getTimeUntilExpiry(): number | null {
    if (!this.cachedToken) {
      return null;
    }

    return Math.max(0, this.cachedToken.expiresAt - Date.now());
  }

  /**
   * Clears all token manager state (for testing/logout)
   */
  public clear(): void {
    this.stopAutoRefresh();
    this.cachedToken = null;
    this.refreshPromise = null;
    this.currentConfig = null;
    console.log('AzureTokenManager: State cleared');
  }
}

// Export singleton instance
export const azureTokenManager = AzureTokenManager.getInstance();

/**
 * Convenience function to get a token using the singleton instance
 *
 * @param {AzureAdConfig} config Azure AD configuration
 * @returns {Promise<string>} Valid access token
 */
export async function getAzureToken(config: AzureAdConfig): Promise<string> {
  return await azureTokenManager.getToken(config);
}

/**
 * Creates Azure AD config from environment variables
 *
 * @returns {AzureAdConfig} Configuration from environment
 * @throws {Error} If required environment variables are missing
 */
export function getAzureConfigFromEnv(): AzureAdConfig {
  const tenantId = process.env.REACT_APP_AZURE_TENANT_ID;
  const clientId = process.env.REACT_APP_AZURE_CLIENT_ID;
  const clientSecret = process.env.REACT_APP_AZURE_CLIENT_SECRET;
  const scope = process.env.REACT_APP_AZURE_SCOPE;

  if (!tenantId || !clientId || !clientSecret || !scope) {
    throw new Error(
      'Azure AD configuration missing. Please set REACT_APP_AZURE_TENANT_ID, ' +
      'REACT_APP_AZURE_CLIENT_ID, REACT_APP_AZURE_CLIENT_SECRET, and REACT_APP_AZURE_SCOPE in your .env file.'
    );
  }

  return { tenantId, clientId, clientSecret, scope };
}

/**
 * Convenience function to start auto-refresh monitoring
 *
 * @param {AzureAdConfig} config Azure AD configuration
 */
export function startAzureTokenAutoRefresh(config: AzureAdConfig): void {
  azureTokenManager.startAutoRefresh(config);
}

/**
 * Convenience function to stop auto-refresh monitoring
 */
export function stopAzureTokenAutoRefresh(): void {
  azureTokenManager.stopAutoRefresh();
}

/**
 * Convenience function to get token status
 *
 * @returns {AzureTokenState} Token status information
 */
export function getAzureTokenStatus(): AzureTokenState {
  return azureTokenManager.getTokenStatus();
}
