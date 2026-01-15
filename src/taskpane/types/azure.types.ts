/**
 * Azure AD Token Type Definitions
 *
 * Types for Client Credentials OAuth 2.0 flow to acquire
 * access tokens from Azure AD for Web API authentication.
 *
 * @module azure.types
 */

/**
 * Azure AD token response from /oauth2/v2.0/token endpoint
 * @see https://learn.microsoft.com/en-us/entra/identity-platform/v2-oauth2-client-creds-grant-flow
 */
export interface AzureTokenResponse {
  access_token: string;
  token_type: string;       // "Bearer"
  expires_in: number;       // seconds until expiry
  ext_expires_in?: number;  // extended expiry for resilience
}

/**
 * Azure AD Client Credentials configuration
 * All values should come from environment variables
 */
export interface AzureAdConfig {
  tenantId: string;
  clientId: string;
  clientSecret: string;
  scope: string;  // Must end with "/.default" for client credentials
}

/**
 * Cached token with metadata for expiry tracking
 */
export interface CachedAzureToken {
  token: string;
  expiresAt: number;   // Unix timestamp in milliseconds
  acquiredAt: number;  // Unix timestamp in milliseconds
}

/**
 * Azure token service state for monitoring and UI display
 */
export interface AzureTokenState {
  isInitialized: boolean;
  hasValidToken: boolean;
  timeUntilExpiry: number | null;
  timeUntilRefresh: number | null;
  isRefreshing: boolean;
  error: string | null;
}

/**
 * Azure AD error response structure
 */
export interface AzureAdError {
  error: string;
  error_description: string;
  error_codes?: number[];
  timestamp?: string;
  trace_id?: string;
  correlation_id?: string;
}
