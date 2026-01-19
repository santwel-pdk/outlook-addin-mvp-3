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
        const messageHandler = (message: SignalRMessage) => {
          console.log('[useSignalR] Message received:', message);
          if (isMounted) {
            setState(prev => ({
              ...prev,
              lastMessage: {
                ...message,
                timestamp: new Date(message.timestamp) // Ensure Date object
              }
            }));
          }
        };

        // CRITICAL: Pass handlers in config so they are registered BEFORE start()
        // Method names are case-sensitive and must match server-side exactly
        const config: SignalRConfig = {
          hubUrl: hubUrl,
          accessToken: process.env.REACT_APP_SIGNALR_ACCESS_TOKEN,
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