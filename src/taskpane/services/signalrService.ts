/**
 * SignalR Connection Service
 *
 * Handles SignalR connection management with proper error handling and typing.
 *
 * @module signalrService
 */

import { HubConnectionBuilder, HubConnection, HubConnectionState, LogLevel } from '@microsoft/signalr';
import { handleOfficeError, logError } from '../utils/errorHandler';
import { SignalRConfig, SignalRMessage } from '../types/signalr.types';
import { getValidToken } from './tokenManagerService';

let isInitialized = false;
let connection: HubConnection | null = null;
let currentConfig: SignalRConfig | null = null;

/**
 * Validates SignalR configuration
 *
 * @param {SignalRConfig} config Configuration to validate
 * @throws {Error} If configuration is invalid
 */
function validateSignalRConfig(config: SignalRConfig): void {
  if (!config.hubUrl) {
    throw new Error('SignalR Hub URL is required. Please set REACT_APP_SIGNALR_HUB_URL in your .env file.');
  }

  if (!config.hubUrl.startsWith('https://')) {
    throw new Error('SignalR Hub URL must use HTTPS for security. Current URL: ' + config.hubUrl);
  }

  // Validate authentication configuration
  if (!config.ssoTokenProvider && !config.accessToken) {
    console.warn('No authentication configured. Either provide ssoTokenProvider for SSO or accessToken for static auth.');
  }

  // Optional: Validate token format if static token provided
  if (config.accessToken && config.accessToken.trim().length < 10) {
    console.warn('SignalR access token appears to be too short. Please verify your REACT_APP_SIGNALR_ACCESS_TOKEN.');
  }
}

/**
 * Initializes SignalR connection after Office.js is ready
 *
 * @param {SignalRConfig} config SignalR configuration options
 * @returns {Promise<HubConnection>} Promise resolving to SignalR connection
 * @throws {Error} If Office.js is not initialized or SignalR connection fails
 */
export async function initializeSignalR(config: SignalRConfig): Promise<HubConnection> {
  // CRITICAL: Only connect after Office.js is ready
  if (!Office.context) {
    throw new Error('Office.js must be initialized before SignalR');
  }

  // Validate configuration
  validateSignalRConfig(config);

  if (isInitialized && connection) {
    return connection;
  }

  try {
    // PATTERN: Build connection with automatic reconnection and SSO token provider
    connection = new HubConnectionBuilder()
      .withUrl(config.hubUrl, {
        accessTokenFactory: async () => {
          try {
            // Prefer SSO token provider for fresh tokens
            if (config.ssoTokenProvider) {
              return await config.ssoTokenProvider();
            }
            // Use TokenManager as fallback for SSO
            if (!config.accessToken) {
              return await getValidToken();
            }
            // Fallback to static token for backward compatibility
            return config.accessToken || '';
          } catch (error) {
            logError('SignalR Token Factory', error);
            // Return static token if SSO fails
            return config.accessToken || '';
          }
        }
      })
      .withAutomaticReconnect(config.reconnectPolicy || [0, 2000, 10000, 30000]) // exponential backoff
      .configureLogging(LogLevel.Information)
      .build();

    // PATTERN: Event handlers for connection lifecycle
    connection.onclose(async (error) => {
      // Use existing error handler pattern
      const message = handleOfficeError('SignalR Connection', error);
      logError('SignalR Connection Closed', error);
      isInitialized = false;
    });

    connection.onreconnecting((error) => {
      logError('SignalR Reconnecting', error);
    });

    connection.onreconnected((connectionId) => {
      console.log('SignalR reconnected with connection ID:', connectionId);
    });

    await connection.start();
    isInitialized = true;
    currentConfig = config;
    console.log('SignalR connected successfully');
    
    return connection;
    
  } catch (error) {
    // PATTERN: Consistent error handling
    const message = handleOfficeError('SignalR Initialization', error);
    throw new Error(message);
  }
}

/**
 * Checks if SignalR has been initialized
 *
 * @returns {boolean} True if SignalR is initialized
 */
export function isSignalRInitialized(): boolean {
  return isInitialized;
}

/**
 * Gets the SignalR connection (throws if not initialized)
 *
 * @returns {HubConnection} SignalR connection
 * @throws {Error} If SignalR is not initialized
 */
export function getSignalRConnection(): HubConnection {
  if (!connection) {
    throw new Error('SignalR is not initialized. Call initializeSignalR() first.');
  }
  return connection;
}

/**
 * Gets the current connection state
 *
 * @returns {HubConnectionState} Current connection state
 */
export function getConnectionState(): HubConnectionState {
  return connection?.state || HubConnectionState.Disconnected;
}

/**
 * Checks if SignalR connection is currently connected
 *
 * @returns {boolean} True if connected
 */
export function isConnected(): boolean {
  return connection?.state === HubConnectionState.Connected;
}

/**
 * Sends a message through SignalR connection
 *
 * @param {string} methodName Hub method name
 * @param {any} data Data to send
 * @returns {Promise<void>} Promise resolving when message is sent
 * @throws {Error} If not connected or send fails
 */
export async function sendMessage(methodName: string, data: any): Promise<void> {
  if (!connection) {
    throw new Error('SignalR is not initialized');
  }

  if (connection.state !== HubConnectionState.Connected) {
    throw new Error('SignalR is not connected');
  }

  try {
    await connection.invoke(methodName, data);
  } catch (error) {
    const message = handleOfficeError('SignalR Send Message', error);
    throw new Error(message);
  }
}

/**
 * Registers a handler for incoming SignalR messages
 *
 * @param {string} methodName Hub method name to listen for
 * @param {Function} handler Function to handle incoming messages
 * @throws {Error} If connection is not available
 */
export function onMessage(methodName: string, handler: (...args: any[]) => void): void {
  if (!connection) {
    throw new Error('SignalR is not initialized');
  }

  connection.on(methodName, handler);
}

/**
 * Removes a message handler
 *
 * @param {string} methodName Hub method name
 * @param {Function} handler Handler function to remove
 */
export function offMessage(methodName: string, handler?: (...args: any[]) => void): void {
  if (!connection) {
    return;
  }

  if (handler) {
    connection.off(methodName, handler);
  } else {
    connection.off(methodName);
  }
}

/**
 * Gracefully stops the SignalR connection
 *
 * @returns {Promise<void>} Promise resolving when connection is stopped
 */
export async function stopSignalR(): Promise<void> {
  if (!connection) {
    return;
  }

  try {
    await connection.stop();
    console.log('SignalR connection stopped');
  } catch (error) {
    logError('SignalR Stop', error);
  } finally {
    connection = null;
    isInitialized = false;
    currentConfig = null;
  }
}

/**
 * Gets the current SignalR configuration
 *
 * @returns {SignalRConfig | null} Current configuration or null if not initialized
 */
export function getCurrentConfig(): SignalRConfig | null {
  return currentConfig;
}