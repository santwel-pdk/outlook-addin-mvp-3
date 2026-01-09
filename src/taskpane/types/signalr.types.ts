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
 * SignalR configuration options
 */
export interface SignalRConfig {
  hubUrl: string;
  accessToken?: string; // Keep for backward compatibility
  ssoTokenProvider?: () => Promise<string>; // NEW: SSO token factory
  reconnectPolicy?: number[];
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