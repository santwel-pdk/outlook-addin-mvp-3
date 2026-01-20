/**
 * useSignalR Hook
 *
 * React hook to manage SignalR connection in components with loading and error states.
 * CRITICAL: Handlers are now passed during initialization to ensure they are registered
 * BEFORE connection.start() - this prevents missing initial messages.
 *
 * @module useSignalR
 * @see https://learn.microsoft.com/en-us/aspnet/core/signalr/javascript-client
 */

import { useState, useEffect } from 'react';
import { HubConnectionState } from '@microsoft/signalr';
import { initializeSignalR, isSignalRInitialized, offMessage, getConnectionState } from '../services/signalrService';
import { SignalRConnectionState, SignalRMessage, SignalRConfig } from '../types/signalr.types';
import { useOfficeContext } from './useOfficeContext';

/**
 * Hook to initialize and manage SignalR connection
 *
 * @returns {SignalRConnectionState} SignalR connection with loading and error states
 */
export function useSignalR(): SignalRConnectionState {
  const [state, setState] = useState<SignalRConnectionState>({
    connection: null,
    connectionState: HubConnectionState.Disconnected,
    isConnected: false,
    error: null,
    lastMessage: null
  });

  const { isInitialized: isOfficeReady } = useOfficeContext();

  useEffect(() => {
    let isMounted = true;
    
    // CRITICAL: Wait for Office.js before connecting SignalR
    if (!isOfficeReady) {
      return undefined;
    }

    const connectSignalR = async () => {
      // Skip if already initialized
      if (isSignalRInitialized()) {
        if (isMounted) {
          setState({
            connection: null, // Don't expose raw connection
            connectionState: getConnectionState(),
            isConnected: getConnectionState() === HubConnectionState.Connected,
            error: null,
            lastMessage: null
          });
        }
        return;
      }

      try {
        // Get configuration from environment variables
        const hubUrl = process.env.REACT_APP_SIGNALR_HUB_URL;

        // Provide helpful error message if environment variables are not configured
        if (!hubUrl) {
          throw new Error(
            'SignalR configuration missing. Please create a .env file with REACT_APP_SIGNALR_HUB_URL. ' +
            'See .env.example for a template.'
          );
        }

        // CRITICAL: Define message handler BEFORE initializing connection
        // This handler will be registered BEFORE connection.start() is called
        // to ensure no messages are missed during or immediately after connection
        //
        // NOTE: SignalR can send messages in various formats depending on server implementation:
        // 1. Single object: SendAsync("Method", { type, payload, ... })
        // 2. Multiple args: SendAsync("Method", type, payload, data)
        // 3. Raw data: SendAsync("Method", "some string or data")
        // This handler normalizes any format into a SignalRMessage structure.
        const messageHandler = (...args: any[]) => {
          console.log('[useSignalR] Raw message received:', args);

          if (!isMounted) {
            return;
          }

          // Normalize the incoming message into SignalRMessage format
          let normalizedMessage: SignalRMessage;

          if (args.length === 0) {
            // No arguments - create empty notification
            normalizedMessage = {
              type: 'notification',
              payload: null,
              timestamp: new Date(),
              id: crypto.randomUUID ? crypto.randomUUID() : `msg-${Date.now()}`
            };
          } else if (args.length === 1 && typeof args[0] === 'object' && args[0] !== null) {
            // Single object argument - check if it's already a SignalRMessage-like structure
            const obj = args[0];
            normalizedMessage = {
              type: obj.type || obj.Type || 'notification',
              payload: obj.payload || obj.Payload || obj.data || obj.Data || obj,
              timestamp: obj.timestamp || obj.Timestamp ? new Date(obj.timestamp || obj.Timestamp) : new Date(),
              id: obj.id || obj.Id || (crypto.randomUUID ? crypto.randomUUID() : `msg-${Date.now()}`)
            };
          } else if (args.length === 1) {
            // Single primitive argument
            normalizedMessage = {
              type: 'notification',
              payload: args[0],
              timestamp: new Date(),
              id: crypto.randomUUID ? crypto.randomUUID() : `msg-${Date.now()}`
            };
          } else {
            // Multiple arguments - first is often the type/method, rest is payload
            normalizedMessage = {
              type: typeof args[0] === 'string' ? args[0] : 'notification',
              payload: args.length === 2 ? args[1] : args.slice(1),
              timestamp: new Date(),
              id: crypto.randomUUID ? crypto.randomUUID() : `msg-${Date.now()}`
            };
          }

          console.log('[useSignalR] Normalized message:', normalizedMessage);

          setState(prev => ({
            ...prev,
            lastMessage: normalizedMessage
          }));
        };

        // CRITICAL: Pass handlers in config so they are registered BEFORE start()
        // Method names are case-sensitive and must match server-side exactly
        const config: SignalRConfig = {
          hubUrl: hubUrl,
          accessToken: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCIsImtpZCI6Ijc3MTI5NjE2MiJ9.eyJhc3JzLnMudWlkIjoiYjQxZjYwNGMtOTZhMS00MTAxLTlmMmQtYzgwNjAxMjhkYzVhIiwibmJmIjoxNzY4ODgwMDA2LCJleHAiOjE3Njg4ODM2MDYsImlhdCI6MTc2ODg4MDAwNiwiaXNzIjoiYXp1cmUtc2lnbmFsciIsImF1ZCI6Imh0dHBzOi8vc2lnbmFsci1vdXRsb29rLWFkZGluLnNlcnZpY2Uuc2lnbmFsci5uZXQvY2xpZW50Lz9odWI9ZW1haWxub3RpZmljYXRpb25zIn0.rs-Z7Pneft59qZkg1Mm-7cgl9ylEonqYDgXTzRLN03E",//process.env.REACT_APP_SIGNALR_ACCESS_TOKEN,
          reconnectPolicy: [0, 2000, 10000, 30000],
          handlers: [
            { methodName: 'NotificationReceived', handler: messageHandler },
            { methodName: 'BroadcastMessage', handler: messageHandler }
          ]
        };

        // Handlers are now registered INSIDE initializeSignalR, BEFORE connection.start()
        const connection = await initializeSignalR(config);

        if (isMounted) {
          setState({
            connection: null, // Don't expose raw connection for security
            connectionState: connection.state,
            isConnected: true,
            error: null,
            lastMessage: null
          });
        }
      } catch (error) {
        console.error('[useSignalR] Connection error:', error);
        if (isMounted) {
          setState(prev => ({
            ...prev,
            connectionState: HubConnectionState.Disconnected,
            isConnected: false,
            error: (error as Error).message
          }));
        }
      }
    };

    connectSignalR();

    return () => {
      isMounted = false;
      // Clean up message handlers
      offMessage('NotificationReceived');
      offMessage('BroadcastMessage');
    };
  }, [isOfficeReady]);

  // Monitor connection state changes
  useEffect(() => {
    if (!isSignalRInitialized()) {
      return undefined;
    }

    const checkConnectionState = () => {
      const currentState = getConnectionState();
      setState(prev => ({
        ...prev,
        connectionState: currentState,
        isConnected: currentState === HubConnectionState.Connected
      }));
    };

    // Check connection state periodically
    const interval = setInterval(checkConnectionState, 2000);

    return () => {
      clearInterval(interval);
    };
  }, [isOfficeReady]);

  return state;
}