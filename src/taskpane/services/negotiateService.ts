/**
 * SignalR Negotiate Service
 *
 * Handles SignalR negotiation with Azure SignalR Service or custom negotiate endpoints.
 * Includes retry logic with exponential backoff for transient failures.
 *
 * @module negotiateService
 * @see https://learn.microsoft.com/en-us/azure/azure-signalr/signalr-concept-internals#server-to-client-negotiate
 */

import {
  NegotiateResponse,
  SignalRConnectionInfo,
  NegotiateConfig
} from '../types/signalr.types';
import { logError } from '../utils/errorHandler';

const DEFAULT_MAX_RETRIES = 3;
const DEFAULT_RETRY_DELAY_MS = 1000;

/**
 * Helper function to pause execution
 *
 * @param {number} ms Milliseconds to sleep
 * @returns {Promise<void>}
 */
function sleep(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * Checks if an error is an authentication error that should not be retried
 *
 * @param {Error} error Error to check
 * @returns {boolean} True if error is auth-related
 */
function isAuthError(error: Error): boolean {
  const message = error.message.toLowerCase();
  return message.includes('401') ||
         message.includes('403') ||
         message.includes('authentication failed') ||
         message.includes('unauthorized');
}

/**
 * Validates the negotiate response structure
 *
 * @param {any} data Response data to validate
 * @returns {data is NegotiateResponse} Type guard for NegotiateResponse
 */
function isValidNegotiateResponse(data: any): data is NegotiateResponse {
  return data &&
         typeof data.url === 'string' &&
         data.url.length > 0 &&
         typeof data.accessToken === 'string' &&
         data.accessToken.length > 0;
}

/**
 * Negotiates SignalR connection using Bearer token authentication
 *
 * @param {NegotiateConfig} config Negotiate configuration
 * @param {string} bearerToken Bearer token for authentication
 * @returns {Promise<SignalRConnectionInfo>} SignalR connection information
 * @throws {Error} If negotiation fails after retries or on auth error
 */
export async function negotiate(
  config: NegotiateConfig,
  bearerToken: string
): Promise<SignalRConnectionInfo> {
  const maxRetries = config.maxRetries ?? DEFAULT_MAX_RETRIES;
  const retryDelayMs = config.retryDelayMs ?? DEFAULT_RETRY_DELAY_MS;

  if (!config.negotiateUrl) {
    throw new Error('Negotiate URL is required');
  }

  if (!bearerToken) {
    throw new Error('Bearer token is required for negotiate');
  }

  let lastError: Error | null = null;

  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      console.log(`NegotiateService: Attempting negotiation (${attempt}/${maxRetries})`);

      const response = await fetch(config.negotiateUrl, {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${bearerToken}`,
          'Content-Type': 'application/json'
        }
      });

      if (!response.ok) {
        const errorText = await response.text().catch(() => 'Unknown error');

        // CRITICAL: Don't retry on authentication errors
        if (response.status === 401 || response.status === 403) {
          const authError = new Error(
            `Authentication failed during SignalR negotiation (${response.status}): ${errorText}`
          );
          logError('NegotiateService Auth Error', authError);
          throw authError;
        }

        throw new Error(`Negotiate request failed: ${response.status} - ${errorText}`);
      }

      const data = await response.json();

      // CRITICAL: Validate response structure
      if (!isValidNegotiateResponse(data)) {
        throw new Error('Invalid negotiate response: missing or empty url or accessToken');
      }

      console.log('NegotiateService: Negotiation successful');

      return {
        url: data.url,
        accessToken: data.accessToken
      };

    } catch (error) {
      lastError = error instanceof Error ? error : new Error(String(error));

      // CRITICAL: Don't retry on authentication errors
      if (isAuthError(lastError)) {
        throw lastError;
      }

      logError(`NegotiateService Attempt ${attempt}`, lastError);

      // If this was the last attempt, throw the error
      if (attempt === maxRetries) {
        break;
      }

      // Exponential backoff: 1s, 2s, 4s, ...
      const delay = retryDelayMs * Math.pow(2, attempt - 1);
      console.log(`NegotiateService: Retrying in ${delay}ms...`);
      await sleep(delay);
    }
  }

  // All retries exhausted
  const finalError = new Error(
    `SignalR negotiation failed after ${maxRetries} attempts: ${lastError?.message || 'Unknown error'}`
  );
  logError('NegotiateService Final Failure', finalError);
  throw finalError;
}

/**
 * Creates negotiate config from environment variables
 *
 * @returns {NegotiateConfig | null} Configuration from environment or null if not configured
 */
export function getNegotiateConfigFromEnv(): NegotiateConfig | null {
  const negotiateUrl = process.env.REACT_APP_SIGNALR_NEGOTIATE_URL;

  if (!negotiateUrl) {
    return null;
  }

  return {
    negotiateUrl,
    maxRetries: DEFAULT_MAX_RETRIES,
    retryDelayMs: DEFAULT_RETRY_DELAY_MS
  };
}

/**
 * Checks if negotiate endpoint is configured
 *
 * @returns {boolean} True if negotiate URL is configured in environment
 */
export function isNegotiateConfigured(): boolean {
  return !!process.env.REACT_APP_SIGNALR_NEGOTIATE_URL;
}
