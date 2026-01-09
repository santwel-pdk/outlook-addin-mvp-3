/**
 * Error Handler Utility
 *
 * Centralized error handling for Office.js API errors with user-friendly messages.
 * No sensitive data logging in production.
 *
 * @module errorHandler
 */

import { SSOError, SSOErrorCode } from '../types/sso.types';

/**
 * Handles Office.js API errors and returns user-friendly error messages
 *
 * @param {Error} error The error object
 * @returns {string} User-friendly error message
 */
export function handleOfficeError(context: string, error: any): string {
  // Log error for debugging (without sensitive data)
  logError(context, error);

  // Return user-friendly message
  if (error.name === 'TypeError') {
    return 'Unable to access email data. Please ensure an email is selected.';
  }

  if (error.message && error.message.includes('not initialized')) {
    return 'Office Add-in is still initializing. Please wait a moment and try again.';
  }

  if (error.message && error.message.includes('not available')) {
    return 'This feature is not available for the current email.';
  }

  return error.message || 'An unexpected error occurred. Please try again.';
}

/**
 * Logs errors for debugging without exposing sensitive data
 *
 * @param {string} context Context where the error occurred
 * @param {Error} error The error object
 */
export function logError(context: string, error: any): void {
  // Only log in development, not production
  if (process.env.NODE_ENV !== 'production') {
    console.error(`[${context}] Error:`, {
      name: error.name,
      message: error.message,
      stack: error.stack
    });
  } else {
    // In production, log only minimal information
    console.error(`[${context}] Error occurred:`, error.name || 'Unknown error');
  }
}

/**
 * Creates a user-friendly error message for display
 *
 * @param {string} operation The operation that failed
 * @param {Error} error The error object
 * @returns {string} Formatted error message
 */
export function formatErrorMessage(operation: string, error: any): string {
  const friendlyMessage = handleOfficeError(operation, error);
  return `${operation} failed: ${friendlyMessage}`;
}

/**
 * Wraps an async function with error handling
 *
 * @param {Function} fn Async function to wrap
 * @param {string} context Context for error logging
 * @returns {Function} Wrapped function with error handling
 */
export function withErrorHandling<T extends (...args: any[]) => Promise<any>>(
  fn: T,
  context: string
): T {
  return (async (...args: any[]) => {
    try {
      return await fn(...args);
    } catch (error) {
      throw new Error(handleOfficeError(context, error));
    }
  }) as T;
}

/**
 * Handles SSO-specific errors and returns user-friendly messages
 *
 * @param {any} error SSO error object from Office.js
 * @returns {string} User-friendly error message
 */
export function handleSSOError(error: any): string {
  // Log error for debugging
  logError('SSO Authentication', error);

  // Check if it's a known SSO error code
  switch (error.code) {
    case SSOErrorCode.IDENTITY_API_NOT_SUPPORTED:
      return 'Single sign-on is not properly configured for this add-in. Please check the manifest configuration.';
    
    case SSOErrorCode.USER_NOT_SIGNED_IN:
      return 'Please sign in to your Microsoft 365 account and try again.';
    
    case SSOErrorCode.USER_ABORTED_CONSENT:
      return 'Authentication was cancelled. Please try again and accept the required permissions.';
    
    case SSOErrorCode.TOKEN_TYPE_NOT_SUPPORTED:
      return 'The authentication method is not supported in this environment.';
    
    case SSOErrorCode.API_NOT_AVAILABLE:
      return 'Single sign-on is not available in your current Office version. Please update Office or contact your administrator.';
    
    case SSOErrorCode.ADMIN_CONSENT_REQUIRED:
      return 'Administrator approval is required for this application. Please contact your IT administrator.';
    
    case SSOErrorCode.INTERNAL_ERROR:
      return 'An internal authentication error occurred. Please try again later.';
    
    default:
      return error.message || 'Authentication failed. Please sign in and try again.';
  }
}

/**
 * Maps SSO error to proper SSOError object with user-friendly message
 *
 * @param {any} error Raw SSO error from Office.js
 * @returns {SSOError} Mapped SSO error object
 */
export function mapSSOError(error: any): SSOError {
  const friendlyMessage = handleSSOError(error);
  
  return {
    code: error.code || 0,
    name: 'SSOError',
    message: friendlyMessage
  };
}

/**
 * Formats SSO error message for display in UI components
 *
 * @param {SSOError} error SSO error object
 * @returns {string} Formatted error message
 */
export function formatSSOErrorMessage(error: SSOError): string {
  return `Authentication Error (${error.code}): ${error.message}`;
}

/**
 * Checks if an error is an SSO-specific error
 *
 * @param {any} error Error object to check
 * @returns {boolean} True if error is SSO-related
 */
export function isSSOError(error: any): boolean {
  return error.code && Object.values(SSOErrorCode).includes(error.code);
}

/**
 * Gets user guidance based on SSO error code
 *
 * @param {number} errorCode SSO error code
 * @returns {string} Detailed user guidance
 */
export function getSSOErrorGuidance(errorCode: number): string {
  switch (errorCode) {
    case SSOErrorCode.IDENTITY_API_NOT_SUPPORTED:
      return 'This add-in requires proper SSO configuration in the manifest. Contact your administrator.';
    
    case SSOErrorCode.USER_NOT_SIGNED_IN:
      return 'Sign in to your Microsoft 365 account in Office, then refresh this add-in.';
    
    case SSOErrorCode.USER_ABORTED_CONSENT:
      return 'Click "Try Again" and accept the permission request to use this add-in.';
    
    case SSOErrorCode.TOKEN_TYPE_NOT_SUPPORTED:
      return 'Try using a different Office application or update to the latest version.';
    
    case SSOErrorCode.API_NOT_AVAILABLE:
      return 'Update Office to the latest version or contact your administrator for support.';
    
    case SSOErrorCode.ADMIN_CONSENT_REQUIRED:
      return 'Contact your IT administrator to approve this add-in for your organization.';
    
    case SSOErrorCode.INTERNAL_ERROR:
      return 'Wait a few moments and try again. If the problem persists, restart Office.';
    
    default:
      return 'Try signing out and back in to Office, then refresh this add-in.';
  }
}
