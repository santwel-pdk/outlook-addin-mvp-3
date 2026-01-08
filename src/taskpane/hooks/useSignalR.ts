/**
 * useSignalR Hook
 *
 * React hook to manage SignalR connection in components with loading and error states.
 *
 * @module useSignalR
 */

import { useState, useEffect } from 'react';
import { HubConnectionState } from '@microsoft/signalr';
import { initializeSignalR, isSignalRInitialized, onMessage, offMessage, getConnectionState } from '../services/signalrService';
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

        const config: SignalRConfig = {
          hubUrl: hubUrl,
          accessToken: process.env.REACT_APP_SIGNALR_ACCESS_TOKEN,
          reconnectPolicy: [0, 2000, 10000, 30000]
        };
        
        const connection = await initializeSignalR(config);
        
        // PATTERN: Setup message handlers for real-time notifications
        const messageHandler = (message: SignalRMessage) => {
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

        // Register for notification messages
        onMessage('NotificationReceived', messageHandler);
        onMessage('BroadcastMessage', messageHandler);

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