/**
 * SSO Authentication Service
 *
 * Handles Office.js SSO authentication with proper error handling and token management.
 *
 * @module ssoService
 */

import { handleOfficeError, logError } from '../utils/errorHandler';
import { SSOConfig, SSOServiceState, SSOToken, SSOUser, SSOError, SSOErrorCode, JWTPayload } from '../types/sso.types';

let isInitialized = false;
let ssoState: SSOServiceState = {
  isInitialized: false,
  isAuthenticated: false,
  token: null,
  error: null,
  user: null
};

/**
 * Validates SSO configuration
 *
 * @param {SSOConfig} config Configuration to validate
 * @throws {Error} If configuration is invalid
 */
function validateSSOConfig(config: SSOConfig): void {
  // Basic validation - Office.js handles most validation
  if (typeof config !== 'object') {
    throw new Error('SSO configuration must be an object');
  }
}

/**
 * Parses JWT token to extract user information and expiration
 *
 * @param {string} token JWT token string
 * @returns {SSOToken} Parsed token information
 * @throws {Error} If token parsing fails
 */
function parseJWTToken(token: string): SSOToken {
  try {
    const parts = token.split('.');
    
    if (parts.length !== 3) {
      throw new Error('Invalid JWT token format');
    }

    const payload = JSON.parse(atob(parts[1])) as JWTPayload;
    
    if (!payload.exp) {
      throw new Error('Token does not contain expiry claim');
    }

    return {
      token,
      expiresAt: payload.exp * 1000, // Convert to milliseconds
      scopes: payload.scp?.split(' ') || [],
      userId: payload.sub || '',
      audience: payload.aud,
      issuer: payload.iss
    };
  } catch (error) {
    logError('JWT Token Parsing', error);
    // Default to 1 hour if parsing fails
    return {
      token,
      expiresAt: Date.now() + 60 * 60 * 1000,
      scopes: [],
      userId: 'unknown'
    };
  }
}

/**
 * Extracts user information from JWT payload
 *
 * @param {string} token JWT token string
 * @returns {SSOUser} User information
 */
function extractUserInfo(token: string): SSOUser {
  try {
    const parts = token.split('.');
    const payload = JSON.parse(atob(parts[1])) as JWTPayload;
    
    return {
      displayName: payload.name || payload.preferred_username || 'Unknown User',
      email: payload.preferred_username || payload.email || '',
      userId: payload.sub || '',
      tenantId: payload.tid
    };
  } catch (error) {
    logError('User Info Extraction', error);
    return {
      displayName: 'Unknown User',
      email: '',
      userId: 'unknown'
    };
  }
}

/**
 * Maps Office.js SSO error codes to user-friendly messages
 *
 * @param {any} error Office.js error
 * @returns {SSOError} Mapped SSO error
 */
function mapSSOError(error: any): SSOError {
  const ssoError: SSOError = {
    code: error.code || 0,
    name: 'SSOError',
    message: error.message || 'Unknown SSO error'
  };

  switch (error.code) {
    case SSOErrorCode.USER_NOT_SIGNED_IN:
      ssoError.message = 'User is not signed in to Office. Please sign in and try again.';
      break;
    case SSOErrorCode.USER_ABORTED_CONSENT:
      ssoError.message = 'User cancelled the consent dialog. Please try again and accept the permissions.';
      break;
    case SSOErrorCode.TOKEN_TYPE_NOT_SUPPORTED:
      ssoError.message = 'The requested token type is not supported in this context.';
      break;
    case SSOErrorCode.API_NOT_AVAILABLE:
      ssoError.message = 'SSO API is not available in the current Office host or version.';
      break;
    case SSOErrorCode.ADMIN_CONSENT_REQUIRED:
      ssoError.message = 'Admin consent is required for this application. Please contact your administrator.';
      break;
    case SSOErrorCode.INTERNAL_ERROR:
      ssoError.message = 'Internal error occurred during authentication. Please try again.';
      break;
    default:
      ssoError.message = `Authentication failed: ${error.message || 'Unknown error'}`;
  }

  return ssoError;
}

/**
 * Initializes SSO authentication after Office.js is ready
 *
 * @param {SSOConfig} config SSO configuration options
 * @returns {Promise<SSOServiceState>} Promise resolving to SSO state
 * @throws {Error} If Office.js is not initialized or SSO fails
 */
export async function initializeSSO(config: SSOConfig = {}): Promise<SSOServiceState> {
  // CRITICAL: Only authenticate after Office.js is ready
  if (!Office.context) {
    throw new Error('Office.js must be initialized before SSO');
  }

  // Validate configuration
  validateSSOConfig(config);

  if (isInitialized && ssoState.isAuthenticated && ssoState.token) {
    return ssoState;
  }

  try {
    // Default SSO configuration
    const defaultConfig: SSOConfig = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: false,
      ...config
    };

    // Get access token from Office.js
    const token = await OfficeRuntime.auth.getAccessToken(defaultConfig);
    
    // Parse token and extract user info
    const parsedToken = parseJWTToken(token);
    const userInfo = extractUserInfo(token);

    // Update SSO state
    ssoState = {
      isInitialized: true,
      isAuthenticated: true,
      token: parsedToken,
      error: null,
      user: userInfo
    };

    isInitialized = true;
    console.log('SSO initialized successfully', { userId: userInfo.userId, email: userInfo.email });
    
    return ssoState;
    
  } catch (error) {
    // PATTERN: Consistent error handling
    const mappedError = mapSSOError(error);
    ssoState = {
      isInitialized: true,
      isAuthenticated: false,
      token: null,
      error: mappedError,
      user: null
    };

    logError('SSO Initialization', mappedError);
    throw mappedError;
  }
}

/**
 * Checks if SSO has been initialized
 *
 * @returns {boolean} True if SSO is initialized
 */
export function isSSOInitialized(): boolean {
  return isInitialized;
}

/**
 * Checks if user is currently authenticated via SSO
 *
 * @returns {boolean} True if authenticated
 */
export function isSSOAuthenticated(): boolean {
  return ssoState.isAuthenticated && ssoState.token !== null;
}

/**
 * Gets the current SSO state
 *
 * @returns {SSOServiceState} Current SSO state
 */
export function getSSOState(): SSOServiceState {
  return { ...ssoState };
}

/**
 * Gets the current access token if authenticated
 *
 * @returns {string | null} Access token or null if not authenticated
 */
export function getCurrentToken(): string | null {
  return ssoState.token?.token || null;
}

/**
 * Gets the current user information if authenticated
 *
 * @returns {SSOUser | null} User info or null if not authenticated
 */
export function getCurrentUser(): SSOUser | null {
  return ssoState.user;
}

/**
 * Checks if the current token is still valid (not expired)
 *
 * @returns {boolean} True if token exists and is not expired
 */
export function isTokenValid(): boolean {
  if (!ssoState.token) {
    return false;
  }
  
  // Add 5-minute buffer before expiration
  const bufferMs = 5 * 60 * 1000;
  return ssoState.token.expiresAt > Date.now() + bufferMs;
}

/**
 * Gets the time until token expiry
 *
 * @returns {number | null} Time in milliseconds until expiry, or null if no token
 */
export function getTimeUntilExpiry(): number | null {
  if (!ssoState.token) {
    return null;
  }
  
  return Math.max(0, ssoState.token.expiresAt - Date.now());
}

/**
 * Refreshes the SSO token if needed
 *
 * @param {SSOConfig} config SSO configuration options
 * @returns {Promise<string>} Promise resolving to fresh token
 * @throws {Error} If refresh fails
 */
export async function refreshToken(config: SSOConfig = {}): Promise<string> {
  if (!Office.context) {
    throw new Error('Office.js context not available');
  }

  try {
    // Always get a fresh token
    const token = await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: false, // Don't prompt during refresh
      allowConsentPrompt: false,
      forMSGraphAccess: false,
      ...config
    });

    // Update stored token
    const parsedToken = parseJWTToken(token);
    ssoState.token = parsedToken;
    ssoState.isAuthenticated = true;
    ssoState.error = null;

    console.log('SSO token refreshed successfully');
    return token;

  } catch (error) {
    const mappedError = mapSSOError(error);
    ssoState.error = mappedError;
    ssoState.isAuthenticated = false;
    
    logError('SSO Token Refresh', mappedError);
    throw mappedError;
  }
}

/**
 * Clears the current SSO state (for testing/logout scenarios)
 */
export function clearSSOState(): void {
  ssoState = {
    isInitialized: false,
    isAuthenticated: false,
    token: null,
    error: null,
    user: null
  };
  isInitialized = false;
  console.log('SSO state cleared');
}