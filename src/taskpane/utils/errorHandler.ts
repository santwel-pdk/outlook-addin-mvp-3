/**
 * Error Handler Utility
 *
 * Centralized error handling for Office.js API errors with user-friendly messages.
 * No sensitive data logging in production.
 *
 * @module errorHandler
 */

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
