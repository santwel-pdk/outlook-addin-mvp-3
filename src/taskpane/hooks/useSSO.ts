/**
 * useSSO Hook
 *
 * React hook to manage SSO authentication state in components with loading and error handling.
 *
 * @module useSSO
 */

import { useState, useEffect, useCallback } from 'react';
import { 
  getSSOState, 
  isSSOInitialized, 
  isSSOAuthenticated, 
  initializeSSO, 
  refreshToken,
  getCurrentUser,
  getCurrentToken,
  isTokenValid
} from '../services/ssoService';
import { startTokenAutoRefresh, stopTokenAutoRefresh, getTokenStatus } from '../services/tokenManagerService';
import { SSOServiceState, SSOConfig, SSOUser, SSOStatus } from '../types/sso.types';
import { isOfficeInitialized } from '../services/officeService';
import { logError } from '../utils/errorHandler';

/**
 * Hook to manage SSO authentication state and operations
 *
 * @returns {object} SSO state and authentication functions
 */
export function useSSO() {
  const [ssoState, setSSOState] = useState<SSOServiceState>({
    isInitialized: false,
    isAuthenticated: false,
    token: null,
    error: null,
    user: null
  });

  const [isLoading, setIsLoading] = useState(false);
  const [status, setStatus] = useState<SSOStatus>({
    status: 'not-authenticated',
    message: 'Authentication not started'
  });

  /**
   * Updates SSO state from service
   */
  const updateSSOState = useCallback(() => {
    const currentState = getSSOState();
    setSSOState(currentState);

    // Update status based on current state
    if (currentState.error) {
      setStatus({
        status: 'error',
        message: currentState.error.message
      });
    } else if (currentState.isAuthenticated && currentState.token) {
      setStatus({
        status: 'authenticated',
        message: `Authenticated as ${currentState.user?.displayName || 'User'}`,
        lastAuthenticated: new Date(),
        tokenExpiresAt: new Date(currentState.token.expiresAt)
      });
    } else if (currentState.isInitialized) {
      setStatus({
        status: 'not-authenticated',
        message: 'Authentication required'
      });
    }
  }, []);

  /**
   * Initialize SSO authentication
   */
  const initializeAuthentication = useCallback(async (config: SSOConfig = {}) => {
    if (!isOfficeInitialized()) {
      logError('useSSO', new Error('Office.js not initialized'));
      return;
    }

    if (isSSOInitialized() && isSSOAuthenticated()) {
      updateSSOState();
      return;
    }

    setIsLoading(true);
    setStatus({
      status: 'authenticating',
      message: 'Authenticating with Microsoft 365...'
    });

    try {
      await initializeSSO(config);
      startTokenAutoRefresh();
      updateSSOState();
    } catch (error) {
      logError('SSO Hook Initialization', error);
      updateSSOState(); // This will capture the error state
    } finally {
      setIsLoading(false);
    }
  }, [updateSSOState]);

  /**
   * Refresh authentication token
   */
  const refreshAuthentication = useCallback(async (config: SSOConfig = {}) => {
    if (!isSSOAuthenticated()) {
      await initializeAuthentication(config);
      return;
    }

    setIsLoading(true);
    setStatus({
      status: 'authenticating',
      message: 'Refreshing authentication...'
    });

    try {
      await refreshToken(config);
      updateSSOState();
    } catch (error) {
      logError('SSO Hook Refresh', error);
      updateSSOState();
    } finally {
      setIsLoading(false);
    }
  }, [initializeAuthentication, updateSSOState]);

  /**
   * Get current authentication status
   */
  const getAuthenticationStatus = useCallback(() => {
    const tokenStatus = getTokenStatus();
    return {
      isAuthenticated: tokenStatus.isAuthenticated,
      hasValidToken: tokenStatus.hasValidToken,
      timeUntilExpiry: tokenStatus.timeUntilExpiry,
      isRefreshing: tokenStatus.isRefreshing,
      user: getCurrentUser(),
      token: getCurrentToken()
    };
  }, []);

  /**
   * Sign out (clear SSO state)
   */
  const signOut = useCallback(() => {
    stopTokenAutoRefresh();
    setSSOState({
      isInitialized: false,
      isAuthenticated: false,
      token: null,
      error: null,
      user: null
    });
    setStatus({
      status: 'not-authenticated',
      message: 'Signed out'
    });
  }, []);

  // Initialize SSO when Office.js is ready
  useEffect(() => {
    let isMounted = true;

    const initializeIfReady = async () => {
      // Wait for Office.js to be ready
      if (!isOfficeInitialized()) {
        // Check again in 100ms
        setTimeout(initializeIfReady, 100);
        return;
      }

      if (isMounted) {
        // Check if already initialized
        if (isSSOInitialized()) {
          updateSSOState();
        } else {
          // Auto-initialize with default config
          await initializeAuthentication({
            allowSignInPrompt: false, // Don't prompt automatically
            allowConsentPrompt: false
          });
        }
      }
    };

    initializeIfReady();

    return () => {
      isMounted = false;
      stopTokenAutoRefresh();
    };
  }, [initializeAuthentication, updateSSOState]);

  // Periodic state updates to handle token expiration
  useEffect(() => {
    if (!ssoState.isInitialized) {
      return undefined;
    }

    const interval = setInterval(() => {
      updateSSOState();
    }, 30000); // Update every 30 seconds

    return () => clearInterval(interval);
  }, [ssoState.isInitialized, updateSSOState]);

  return {
    // State
    ssoState,
    isLoading,
    status,
    isInitialized: ssoState.isInitialized,
    isAuthenticated: ssoState.isAuthenticated,
    user: ssoState.user,
    token: ssoState.token,
    error: ssoState.error,

    // Actions
    initialize: initializeAuthentication,
    refresh: refreshAuthentication,
    signOut,
    getStatus: getAuthenticationStatus,

    // Computed properties
    isTokenValid: ssoState.token ? isTokenValid() : false,
    hasError: !!ssoState.error
  };
}