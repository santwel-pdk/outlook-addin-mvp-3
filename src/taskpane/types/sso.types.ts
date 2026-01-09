/**
 * SSO Type Definitions
 *
 * Type definitions for Office.js SSO authentication and app-specific types.
 *
 * @module sso.types
 */

/**
 * SSO configuration options for OfficeRuntime.auth.getAccessToken()
 */
export interface SSOConfig {
  allowSignInPrompt?: boolean;
  allowConsentPrompt?: boolean;
  forMSGraphAccess?: boolean;
}

/**
 * SSO access token with parsed information
 */
export interface SSOToken {
  token: string;
  expiresAt: number;
  scopes: string[];
  userId: string;
  audience?: string;
  issuer?: string;
}

/**
 * SSO-specific error with Office.js error codes
 */
export interface SSOError extends Error {
  code: number;
  name: string;
  message: string;
}

/**
 * User information extracted from SSO token
 */
export interface SSOUser {
  displayName: string;
  email: string;
  userId: string;
  tenantId?: string;
}

/**
 * SSO service state for React components
 */
export interface SSOServiceState {
  isInitialized: boolean;
  isAuthenticated: boolean;
  token: SSOToken | null;
  error: SSOError | null;
  user: SSOUser | null;
}

/**
 * Office.js SSO error codes enum
 */
export enum SSOErrorCode {
  IDENTITY_API_NOT_SUPPORTED = 13000,
  USER_NOT_SIGNED_IN = 13001,
  USER_ABORTED_CONSENT = 13002,
  TOKEN_TYPE_NOT_SUPPORTED = 13003,
  API_NOT_AVAILABLE = 13006,
  ADMIN_CONSENT_REQUIRED = 13012,
  INTERNAL_ERROR = 13013
}

/**
 * SSO status for UI display
 */
export interface SSOStatus {
  status: 'authenticated' | 'not-authenticated' | 'authenticating' | 'error';
  message: string;
  lastAuthenticated?: Date;
  tokenExpiresAt?: Date;
}

/**
 * JWT payload structure for SSO tokens
 */
export interface JWTPayload {
  exp: number;
  iat: number;
  aud: string;
  iss: string;
  sub: string;
  preferred_username?: string;
  name?: string;
  email?: string;
  tid?: string;
  scp?: string;
  [key: string]: unknown;
}