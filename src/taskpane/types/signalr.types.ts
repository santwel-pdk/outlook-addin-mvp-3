/**
 * SignalR Type Definitions
 *
 * Type definitions for SignalR real-time communication and app-specific types.
 *
 * @module signalr.types
 */

import { HubConnection, HubConnectionState } from '@microsoft/signalr';

/**
 * SignalR connection state for React components
 */
export interface SignalRConnectionState {
  connection: HubConnection | null;
  connectionState: HubConnectionState;
  isConnected: boolean;
  error: string | null;
  lastMessage: SignalRMessage | null;
}

/**
 * SignalR message structure for real-time notifications
 */
export interface SignalRMessage {
  type: string;
  payload: any;
  timestamp: Date;
  id: string;
}

/**
 * Handler registration configuration for SignalR initialization
 * Handlers configured here will be registered BEFORE connection.start()
 * @see https://learn.microsoft.com/en-us/aspnet/core/signalr/javascript-client
 */
export interface SignalRHandlerConfig {
  /** Hub method name to listen for (case-sensitive!) */
  methodName: string;
  /** Handler function to invoke when message received */
  handler: (...args: any[]) => void;
}

/**
 * SignalR configuration options
 */
export interface SignalRConfig {
  hubUrl: string;
  negotiateUrl?: string; // NEW: If provided, use negotiate flow for Azure SignalR Service
  accessToken?: string; // Keep for backward compatibility
  ssoTokenProvider?: () => Promise<string>; // SSO token factory
  azureTokenProvider?: () => Promise<string>; // NEW: Azure AD token provider for negotiate
  reconnectPolicy?: number[];
  /** Handlers to register BEFORE connection starts - CRITICAL for receiving initial messages */
  handlers?: SignalRHandlerConfig[];
}

/**
 * SignalR connection status for UI display
 */
export interface SignalRStatus {
  status: 'connected' | 'disconnected' | 'reconnecting' | 'error';
  message: string;
  lastConnected?: Date;
  reconnectAttempts?: number;
}

/**
 * SignalR service state
 */
export interface SignalRServiceState {
  isInitialized: boolean;
  connection: HubConnection | null;
  config: SignalRConfig | null;
  error: Error | null;
}

/**
 * SignalR negotiate endpoint response
 * From Azure SignalR Service or custom negotiate endpoint
 * @see https://learn.microsoft.com/en-us/azure/azure-signalr/signalr-concept-internals#server-to-client-negotiate
 */
export interface NegotiateResponse {
  url: string;           // SignalR hub URL to connect to
  accessToken: string;   // Bearer token for this connection
  availableTransports?: Array<{
    transport: string;
    transferFormats: string[];
  }>;
}

/**
 * SignalR connection info extracted from negotiate response
 */
export interface SignalRConnectionInfo {
  url: string;
  accessToken: string;
}

/**
 * Negotiate service configuration
 */
export interface NegotiateConfig {
  negotiateUrl: string;
  maxRetries?: number;
  retryDelayMs?: number;
}