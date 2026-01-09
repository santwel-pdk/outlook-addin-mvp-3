/**
 * Token Manager Service
 *
 * Singleton service for managing SSO token lifecycle with automatic refresh.
 * Prevents concurrent token requests and ensures fresh tokens for SignalR.
 *
 * @module tokenManagerService
 */

import { refreshToken, getCurrentToken, isTokenValid, getTimeUntilExpiry, isSSOAuthenticated } from './ssoService';
import { logError } from '../utils/errorHandler';
import { SSOConfig } from '../types/sso.types';

/**
 * Singleton token manager class
 */
class TokenManager {
  private static instance: TokenManager;
  private readonly REFRESH_THRESHOLD_MS = 5 * 60 * 1000; // 5 minutes
  private refreshPromise: Promise<string> | null = null;
  private refreshTimer: NodeJS.Timeout | null = null;

  private constructor() {
    // Private constructor for singleton pattern
  }

  /**
   * Gets the singleton instance
   *
   * @returns {TokenManager} TokenManager instance
   */
  public static getInstance(): TokenManager {
    if (!TokenManager.instance) {
      TokenManager.instance = new TokenManager();
    }
    return TokenManager.instance;
  }

  /**
   * Gets a valid access token, refreshing if necessary
   * 
   * @param {SSOConfig} config SSO configuration for refresh
   * @returns {Promise<string>} Valid access token
   * @throws {Error} If token acquisition fails
   */
  public async getToken(config: SSOConfig = {}): Promise<string> {
    // If user is not authenticated, throw error
    if (!isSSOAuthenticated()) {
      throw new Error('User is not authenticated. Please initialize SSO first.');
    }

    // Check if token is still valid (refresh 5 minutes before expiry)
    if (this.isTokenStillValid()) {
      const currentToken = getCurrentToken();
      if (currentToken) {
        return currentToken;
      }
    }

    // If already refreshing, wait for that refresh to complete
    if (this.refreshPromise) {
      try {
        return await this.refreshPromise;
      } catch (error) {
        // If the pending refresh failed, clear it and try again
        this.refreshPromise = null;
        throw error;
      }
    }

    // Start a new refresh
    this.refreshPromise = this.performTokenRefresh(config);
    
    try {
      const newToken = await this.refreshPromise;
      this.refreshPromise = null; // Clear the promise after success
      
      // Schedule the next automatic refresh
      this.scheduleNextRefresh();
      
      return newToken;
    } catch (error) {
      this.refreshPromise = null; // Clear the promise after failure
      throw error;
    }
  }

  /**
   * Checks if the current token is still valid with buffer
   *
   * @returns {boolean} True if token exists and hasn't expired (with 5-min buffer)
   */
  private isTokenStillValid(): boolean {
    if (!isTokenValid()) {
      return false;
    }
    
    const timeUntilExpiry = getTimeUntilExpiry();
    if (timeUntilExpiry === null) {
      return false;
    }
    
    return timeUntilExpiry > this.REFRESH_THRESHOLD_MS;
  }

  /**
   * Performs the actual token refresh
   *
   * @param {SSOConfig} config SSO configuration
   * @returns {Promise<string>} Fresh access token
   */
  private async performTokenRefresh(config: SSOConfig): Promise<string> {
    try {
      console.log('TokenManager: Refreshing access token');
      const newToken = await refreshToken(config);
      console.log('TokenManager: Token refreshed successfully');
      return newToken;
    } catch (error) {
      logError('TokenManager Refresh', error);
      throw new Error(`Failed to refresh token: ${error.message}`);
    }
  }

  /**
   * Schedules the next automatic token refresh
   */
  private scheduleNextRefresh(): void {
    // Clear any existing timer
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
    }

    const timeUntilExpiry = getTimeUntilExpiry();
    if (timeUntilExpiry === null) {
      return;
    }

    // Schedule refresh 5 minutes before expiry
    const refreshIn = Math.max(0, timeUntilExpiry - this.REFRESH_THRESHOLD_MS);
    
    console.log(`TokenManager: Scheduling next refresh in ${Math.round(refreshIn / 1000)} seconds`);
    
    this.refreshTimer = setTimeout(async () => {
      try {
        await this.getToken(); // This will trigger a refresh
      } catch (error) {
        logError('Scheduled Token Refresh', error);
        // Retry in 1 minute if automatic refresh fails
        setTimeout(() => this.scheduleNextRefresh(), 60000);
      }
    }, refreshIn);
  }

  /**
   * Starts automatic token refresh monitoring
   */
  public startAutoRefresh(): void {
    if (!isSSOAuthenticated()) {
      console.warn('TokenManager: Cannot start auto-refresh - user not authenticated');
      return;
    }

    this.scheduleNextRefresh();
    console.log('TokenManager: Auto-refresh started');
  }

  /**
   * Stops automatic token refresh monitoring
   */
  public stopAutoRefresh(): void {
    if (this.refreshTimer) {
      clearTimeout(this.refreshTimer);
      this.refreshTimer = null;
      console.log('TokenManager: Auto-refresh stopped');
    }
  }

  /**
   * Forces an immediate token refresh
   * 
   * @param {SSOConfig} config SSO configuration
   * @returns {Promise<string>} Fresh access token
   */
  public async forceRefresh(config: SSOConfig = {}): Promise<string> {
    // Clear any existing refresh promise to force a new one
    this.refreshPromise = null;
    
    return await this.getToken(config);
  }

  /**
   * Gets token information for debugging/monitoring
   * 
   * @returns {object} Token status information
   */
  public getTokenStatus(): {
    isAuthenticated: boolean;
    hasValidToken: boolean;
    timeUntilExpiry: number | null;
    timeUntilRefresh: number | null;
    isRefreshing: boolean;
  } {
    const timeUntilExpiry = getTimeUntilExpiry();
    const timeUntilRefresh = timeUntilExpiry !== null 
      ? Math.max(0, timeUntilExpiry - this.REFRESH_THRESHOLD_MS)
      : null;

    return {
      isAuthenticated: isSSOAuthenticated(),
      hasValidToken: isTokenValid(),
      timeUntilExpiry,
      timeUntilRefresh,
      isRefreshing: this.refreshPromise !== null
    };
  }

  /**
   * Clears all token manager state (for testing/logout)
   */
  public clear(): void {
    this.stopAutoRefresh();
    this.refreshPromise = null;
    console.log('TokenManager: State cleared');
  }
}

// Export singleton instance
export const tokenManager = TokenManager.getInstance();

/**
 * Convenience function to get a token using the singleton instance
 * 
 * @param {SSOConfig} config SSO configuration
 * @returns {Promise<string>} Valid access token
 */
export async function getValidToken(config: SSOConfig = {}): Promise<string> {
  return await tokenManager.getToken(config);
}

/**
 * Convenience function to start auto-refresh monitoring
 */
export function startTokenAutoRefresh(): void {
  tokenManager.startAutoRefresh();
}

/**
 * Convenience function to stop auto-refresh monitoring
 */
export function stopTokenAutoRefresh(): void {
  tokenManager.stopAutoRefresh();
}

/**
 * Convenience function to get token status
 * 
 * @returns {object} Token status information
 */
export function getTokenStatus(): ReturnType<TokenManager['getTokenStatus']> {
  return tokenManager.getTokenStatus();
}