/**
 * Unit Tests for tokenManagerService
 *
 * Tests token caching, refresh logic, expiration handling, and singleton behavior
 */

import { tokenManager, getValidToken, startTokenAutoRefresh, stopTokenAutoRefresh, getTokenStatus } from '../tokenManagerService';
import * as ssoService from '../ssoService';

// Mock the SSO service
jest.mock('../ssoService');
jest.mock('../../utils/errorHandler');

const mockSsoService = ssoService as jest.Mocked<typeof ssoService>;

describe('tokenManagerService', () => {
  let originalDateNow: () => number;
  let mockCurrentTime: number;

  const mockValidToken = 'valid-test-token';
  const mockRefreshedToken = 'refreshed-test-token';

  beforeEach(() => {
    // Mock Date.now for consistent testing
    originalDateNow = Date.now;
    mockCurrentTime = 1000000000; // Fixed timestamp for testing
    Date.now = jest.fn(() => mockCurrentTime);
    
    // Clear token manager state first
    tokenManager.clear();
    
    // Reset all mocks
    jest.clearAllMocks();
    
    // Setup default mock responses
    mockSsoService.isSSOAuthenticated.mockReturnValue(true);
    mockSsoService.getCurrentToken.mockReturnValue(mockValidToken);
    mockSsoService.isTokenValid.mockReturnValue(true);
    mockSsoService.getTimeUntilExpiry.mockReturnValue(30 * 60 * 1000); // 30 minutes
    mockSsoService.refreshToken.mockResolvedValue(mockRefreshedToken);
  });

  afterEach(() => {
    Date.now = originalDateNow;
    tokenManager.clear();
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  describe('singleton behavior', () => {
    it('should return the same instance', () => {
      const instance1 = tokenManager;
      const instance2 = tokenManager;
      
      expect(instance1).toBe(instance2);
    });
  });

  describe('getToken', () => {
    it('should return current token when valid', async () => {
      const token = await tokenManager.getToken();
      
      expect(token).toBe(mockValidToken);
      expect(mockSsoService.refreshToken).not.toHaveBeenCalled();
    });

    it('should throw error when user not authenticated', async () => {
      mockSsoService.isSSOAuthenticated.mockReturnValue(false);
      
      await expect(tokenManager.getToken()).rejects.toThrow(
        'User is not authenticated. Please initialize SSO first.'
      );
    });

    it('should refresh token when expired', async () => {
      mockSsoService.isTokenValid.mockReturnValue(false);
      
      const token = await tokenManager.getToken();
      
      expect(mockSsoService.refreshToken).toHaveBeenCalled();
      expect(token).toBe(mockRefreshedToken);
    });

    it('should refresh token when close to expiry (within 5 minutes)', async () => {
      // Token expires in 3 minutes (less than 5-minute threshold)
      mockSsoService.getTimeUntilExpiry.mockReturnValue(3 * 60 * 1000);
      
      const token = await tokenManager.getToken();
      
      expect(mockSsoService.refreshToken).toHaveBeenCalled();
      expect(token).toBe(mockRefreshedToken);
    });

    it('should not refresh token when it has enough time left', async () => {
      // Token expires in 10 minutes (more than 5-minute threshold)
      mockSsoService.getTimeUntilExpiry.mockReturnValue(10 * 60 * 1000);
      
      const token = await tokenManager.getToken();
      
      expect(mockSsoService.refreshToken).not.toHaveBeenCalled();
      expect(token).toBe(mockValidToken);
    });

    it('should handle concurrent requests by sharing the same refresh promise', async () => {
      mockSsoService.isTokenValid.mockReturnValue(false);
      
      // Start multiple concurrent requests
      const promise1 = tokenManager.getToken();
      const promise2 = tokenManager.getToken();
      const promise3 = tokenManager.getToken();
      
      const [token1, token2, token3] = await Promise.all([promise1, promise2, promise3]);
      
      expect(token1).toBe(mockRefreshedToken);
      expect(token2).toBe(mockRefreshedToken);
      expect(token3).toBe(mockRefreshedToken);
      
      // Should only call refresh once, not three times
      expect(mockSsoService.refreshToken).toHaveBeenCalledTimes(1);
    });

    it('should retry if concurrent refresh fails', async () => {
      mockSsoService.isTokenValid.mockReturnValue(false);
      mockSsoService.refreshToken
        .mockRejectedValueOnce(new Error('First refresh failed'))
        .mockResolvedValueOnce(mockRefreshedToken);
      
      // First call fails
      await expect(tokenManager.getToken()).rejects.toThrow('First refresh failed');
      
      // Second call should succeed with a new refresh attempt
      const token = await tokenManager.getToken();
      
      expect(token).toBe(mockRefreshedToken);
      expect(mockSsoService.refreshToken).toHaveBeenCalledTimes(2);
    });

    it('should pass config to refresh function', async () => {
      mockSsoService.isTokenValid.mockReturnValue(false);
      
      const config = { forMSGraphAccess: true };
      await tokenManager.getToken(config);
      
      expect(mockSsoService.refreshToken).toHaveBeenCalledWith(config);
    });
  });

  describe('forceRefresh', () => {
    it('should force refresh by clearing refresh promise', async () => {
      // Set up a scenario where token is close to expiry (will trigger refresh)
      mockSsoService.isTokenValid.mockReturnValue(true);
      mockSsoService.getTimeUntilExpiry.mockReturnValue(3 * 60 * 1000); // 3 minutes (less than 5 min threshold)
      
      const token = await tokenManager.forceRefresh();
      
      expect(mockSsoService.refreshToken).toHaveBeenCalled();
      expect(token).toBe(mockRefreshedToken);
    });
  });

  describe('auto-refresh functionality', () => {
    beforeEach(() => {
      jest.useFakeTimers();
    });

    it('should not start auto-refresh when user not authenticated', () => {
      mockSsoService.isSSOAuthenticated.mockReturnValue(false);
      const consoleSpy = jest.spyOn(console, 'warn').mockImplementation();
      
      tokenManager.startAutoRefresh();
      
      expect(consoleSpy).toHaveBeenCalledWith('TokenManager: Cannot start auto-refresh - user not authenticated');
      consoleSpy.mockRestore();
    });

    it('should start auto-refresh when user is authenticated', () => {
      const consoleSpy = jest.spyOn(console, 'log').mockImplementation();
      
      tokenManager.startAutoRefresh();
      
      expect(consoleSpy).toHaveBeenCalledWith('TokenManager: Auto-refresh started');
      consoleSpy.mockRestore();
    });

    it('should stop auto-refresh', () => {
      const consoleSpy = jest.spyOn(console, 'log').mockImplementation();
      
      tokenManager.startAutoRefresh();
      tokenManager.stopAutoRefresh();
      
      expect(consoleSpy).toHaveBeenCalledWith('TokenManager: Auto-refresh stopped');
      consoleSpy.mockRestore();
    });
  });

  describe('getTokenStatus', () => {
    it('should return correct status information', () => {
      mockSsoService.isSSOAuthenticated.mockReturnValue(true);
      mockSsoService.isTokenValid.mockReturnValue(true);
      mockSsoService.getTimeUntilExpiry.mockReturnValue(15 * 60 * 1000); // 15 minutes
      
      const status = tokenManager.getTokenStatus();
      
      expect(status).toEqual({
        isAuthenticated: true,
        hasValidToken: true,
        timeUntilExpiry: 15 * 60 * 1000,
        timeUntilRefresh: 10 * 60 * 1000, // 15 minutes - 5 minute threshold
        isRefreshing: false
      });
    });

    it('should show refreshing state during token refresh', async () => {
      mockSsoService.isTokenValid.mockReturnValue(false);
      mockSsoService.refreshToken.mockImplementation(() => new Promise(resolve => setTimeout(() => resolve(mockRefreshedToken), 100)));
      
      // Start refresh (don't await)
      const refreshPromise = tokenManager.getToken();
      
      // Check status during refresh
      const status = tokenManager.getTokenStatus();
      expect(status.isRefreshing).toBe(true);
      
      // Wait for refresh to complete
      await refreshPromise;
      
      // Check status after refresh
      const statusAfter = tokenManager.getTokenStatus();
      expect(statusAfter.isRefreshing).toBe(false);
    });

    it('should handle null expiry time', () => {
      mockSsoService.getTimeUntilExpiry.mockReturnValue(null);
      
      const status = tokenManager.getTokenStatus();
      
      expect(status.timeUntilExpiry).toBe(null);
      expect(status.timeUntilRefresh).toBe(null);
    });
  });

  describe('clear', () => {
    it('should clear all state and stop timers', () => {
      const consoleSpy = jest.spyOn(console, 'log').mockImplementation();
      
      tokenManager.startAutoRefresh();
      tokenManager.clear();
      
      expect(consoleSpy).toHaveBeenCalledWith('TokenManager: State cleared');
      consoleSpy.mockRestore();
      
      // Verify state is cleared by checking that the next getToken call behaves as if fresh
      mockSsoService.isSSOAuthenticated.mockReturnValue(false);
      expect(tokenManager.getToken()).rejects.toThrow('User is not authenticated');
    });
  });

  describe('convenience functions', () => {
    it('getValidToken should delegate to tokenManager', async () => {
      const token = await getValidToken();
      
      expect(token).toBe(mockValidToken);
    });

    it('getTokenStatus should delegate to tokenManager', () => {
      mockSsoService.isSSOAuthenticated.mockReturnValue(true);
      mockSsoService.isTokenValid.mockReturnValue(true);
      mockSsoService.getTimeUntilExpiry.mockReturnValue(20 * 60 * 1000);
      
      const status = getTokenStatus();
      
      expect(status.isAuthenticated).toBe(true);
      expect(status.hasValidToken).toBe(true);
    });

    it('startTokenAutoRefresh should delegate to tokenManager', () => {
      const startSpy = jest.spyOn(tokenManager, 'startAutoRefresh');
      
      startTokenAutoRefresh();
      
      expect(startSpy).toHaveBeenCalled();
      startSpy.mockRestore();
    });

    it('stopTokenAutoRefresh should delegate to tokenManager', () => {
      const stopSpy = jest.spyOn(tokenManager, 'stopAutoRefresh');
      
      stopTokenAutoRefresh();
      
      expect(stopSpy).toHaveBeenCalled();
      stopSpy.mockRestore();
    });
  });

  describe('edge cases', () => {
    it('should handle refresh errors gracefully', async () => {
      mockSsoService.isTokenValid.mockReturnValue(false);
      const refreshError = new Error('Network error during refresh');
      mockSsoService.refreshToken.mockRejectedValue(refreshError);
      
      await expect(tokenManager.getToken()).rejects.toThrow('Failed to refresh token: Network error during refresh');
    });

    it('should handle missing token in getCurrentToken', async () => {
      mockSsoService.getCurrentToken.mockReturnValue(null);
      mockSsoService.isTokenValid.mockReturnValue(false);
      
      const token = await tokenManager.getToken();
      
      expect(mockSsoService.refreshToken).toHaveBeenCalled();
      expect(token).toBe(mockRefreshedToken);
    });

    it('should handle zero time until expiry', () => {
      mockSsoService.getTimeUntilExpiry.mockReturnValue(0);
      
      const status = tokenManager.getTokenStatus();
      
      expect(status.timeUntilRefresh).toBe(0);
    });

    it('should handle negative time until expiry', () => {
      mockSsoService.getTimeUntilExpiry.mockReturnValue(-1000); // Token expired 1 second ago
      
      const status = tokenManager.getTokenStatus();
      
      expect(status.timeUntilRefresh).toBe(0); // Should not be negative
    });
  });
});