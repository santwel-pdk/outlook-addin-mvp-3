/**
 * SignalR Connection Service
 *
 * Handles SignalR connection management with proper error handling and typing.
 *
 * @module signalrService
 */

import { HubConnectionBuilder, HubConnection, HubConnectionState, LogLevel, HttpTransportType } from '@microsoft/signalr';
import { handleOfficeError, logError } from '../utils/errorHandler';
import { SignalRConfig, SignalRMessage, SignalRConnectionInfo, SignalRHandlerConfig } from '../types/signalr.types';
import { getValidToken } from './tokenManagerService';
import { negotiate, isNegotiateConfigured } from './negotiateService';

let isInitialized = false;
let connection: HubConnection | null = null;
let currentConfig: SignalRConfig | null = null;
let currentConnectionInfo: SignalRConnectionInfo | null = null;
let isUsingNegotiateFlow = false;

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

  // When using negotiate flow, require Azure token provider
  if (config.negotiateUrl) {
    if (!config.negotiateUrl.startsWith('https://')) {
      throw new Error('SignalR Negotiate URL must use HTTPS for security.');
    }
    if (!config.azureTokenProvider && !config.ssoTokenProvider && !config.accessToken) {
      console.warn('Negotiate URL configured but no token provider available for authentication.');
    }
    return; // Skip other auth validation for negotiate flow
  }

  // Validate authentication configuration for direct connection
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
 * Supports both direct connection and negotiate flow for Azure SignalR Service
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
    let connectionUrl = config.hubUrl;
    let tokenFactory: () => Promise<string>;
    let skipNegotiation = false;
    let transportType: HttpTransportType | undefined;

    // NEW: Check if using negotiate flow (Azure SignalR Service)
    if (config.negotiateUrl) {
      console.log('[SignalR] Using negotiate flow for Azure SignalR Service');
      isUsingNegotiateFlow = true;

      // Get bearer token for negotiate endpoint
      let bearerToken: string;
      try {
        if (config.azureTokenProvider) {
          bearerToken = await config.azureTokenProvider();
        } else if (config.ssoTokenProvider) {
          bearerToken = await config.ssoTokenProvider();
        } else if (config.accessToken) {
          bearerToken = config.accessToken;
        } else {
          bearerToken = await getValidToken();
        }
      } catch (tokenError) {
        logError('SignalR Token Acquisition for Negotiate', tokenError);
        throw new Error('Failed to acquire token for SignalR negotiation');
      }

      // Call negotiate endpoint
      const connectionInfo = await negotiate(
        { negotiateUrl: config.negotiateUrl },
        bearerToken
      );

      currentConnectionInfo = connectionInfo;
      connectionUrl = connectionInfo.url;
      tokenFactory = () => Promise.resolve(connectionInfo.accessToken);

      // CRITICAL: Must set skipNegotiation when using pre-negotiated URL
      skipNegotiation = true;
      transportType = HttpTransportType.WebSockets;

      console.log('[SignalR] Negotiation successful, connecting to:', connectionUrl);
    } else {
      // Existing direct connection flow
      isUsingNegotiateFlow = false;
      tokenFactory = async () => {
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
      };
    }

    // Build connection with appropriate configuration
    const builder = new HubConnectionBuilder()
      .withUrl(connectionUrl, {
        accessTokenFactory: tokenFactory,
        skipNegotiation: skipNegotiation,
        transport: transportType
      })
      .withAutomaticReconnect(config.reconnectPolicy || [0, 2000, 10000, 30000])
      .configureLogging(LogLevel.Information);

    connection = builder.build();

    // PATTERN: Event handlers for connection lifecycle
    connection.onclose(async (error) => {
      console.log('[SignalR] Connection closed', error ? `- Error: ${error.message}` : '');
      handleOfficeError('SignalR Connection', error);
      logError('SignalR Connection Closed', error);
      isInitialized = false;
      currentConnectionInfo = null;
    });

    connection.onreconnecting((error) => {
      console.log('[SignalR] Reconnecting...', error ? `- Error: ${error.message}` : '');
      logError('SignalR Reconnecting', error);
    });

    connection.onreconnected(async (connectionId) => {
      console.log('[SignalR] Reconnected with connection ID:', connectionId);

      // CRITICAL: Re-negotiate if using negotiate flow (token may have expired)
      if (isUsingNegotiateFlow && config.negotiateUrl) {
        console.log('[SignalR] Connection restored - negotiate token may need refresh on next reconnect');
        // Note: The current connection should still work, but if it drops again,
        // we may need to re-negotiate. For now, just log the reconnection.
      }
    });

    // CRITICAL: Register all handlers BEFORE starting connection
    // This ensures no messages are missed during or immediately after connection
    // @see https://learn.microsoft.com/en-us/aspnet/core/signalr/javascript-client
    if (config.handlers && config.handlers.length > 0) {
      console.log(`[SignalR] Registering ${config.handlers.length} handler(s) BEFORE start()`);
      for (const handlerConfig of config.handlers) {
        console.log(`[SignalR] Registering handler for method: "${handlerConfig.methodName}"`);
        connection.on(handlerConfig.methodName, (...args: any[]) => {
          console.log(`[SignalR] Handler "${handlerConfig.methodName}" invoked with:`, args);
          handlerConfig.handler(...args);
        });
      }
    } else {
      console.log('[SignalR] No handlers configured - messages may be missed');
    }

    // NOW start the connection - handlers are already registered
    await connection.start();
    isInitialized = true;
    currentConfig = config;
    console.log('[SignalR] Connection started successfully - handlers were registered first');

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
    console.log('[SignalR] Connection stopped');
  } catch (error) {
    logError('SignalR Stop', error);
  } finally {
    connection = null;
    isInitialized = false;
    currentConfig = null;
    currentConnectionInfo = null;
    isUsingNegotiateFlow = false;
  }
}

/**
 * Checks if currently using negotiate flow
 *
 * @returns {boolean} True if using negotiate flow
 */
export function isUsingNegotiate(): boolean {
  return isUsingNegotiateFlow;
}

/**
 * Gets the current connection info from negotiate (if using negotiate flow)
 *
 * @returns {SignalRConnectionInfo | null} Connection info or null
 */
export function getNegotiateConnectionInfo(): SignalRConnectionInfo | null {
  return currentConnectionInfo;
}

/**
 * Gets the current SignalR configuration
 *
 * @returns {SignalRConfig | null} Current configuration or null if not initialized
 */
export function getCurrentConfig(): SignalRConfig | null {
  return currentConfig;
}