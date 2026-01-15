/**
 * Unit Tests for azureTokenService
 *
 * Tests Azure AD Client Credentials token acquisition, caching, refresh logic,
 * expiration handling, and singleton behavior
 */

import {
  azureTokenManager,
  getAzureToken,
  getAzureTokenStatus,
  startAzureTokenAutoRefresh,
  stopAzureTokenAutoRefresh,
  getAzureConfigFromEnv
} from '../azureTokenService';
import { AzureAdConfig } from '../../types/azure.types';

// Mock the error handler
jest.mock('../../utils/errorHandler');

// Mock global fetch
const mockFetch = jest.fn();
global.fetch = mockFetch;

describe('azureTokenService', () => {
  let originalDateNow: () => number;
  let mockCurrentTime: number;

  const mockConfig: AzureAdConfig = {
    tenantId: 'test-tenant-id',
    clientId: 'test-client-id',
    clientSecret: 'test-client-secret',
    scope: 'api://test-api/.default'
  };

  const mockTokenResponse = {
    access_token: 'mock-access-token-12345',
    token_type: 'Bearer',
    expires_in: 3600 // 1 hour
  };

  const mockRefreshedTokenResponse = {
    access_token: 'refreshed-access-token-67890',
    token_type: 'Bearer',
    expires_in: 3600
  };

  beforeEach(() => {
    // Mock Date.now for consistent testing
    originalDateNow = Date.now;
    mockCurrentTime = 1000000000;
    Date.now = jest.fn(() => mockCurrentTime);

    // Clear token manager state
    azureTokenManager.clear();

    // Reset all mocks
    jest.clearAllMocks();
    mockFetch.mockReset();

    // Setup default successful response
    mockFetch.mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(mockTokenResponse)
    });
  });

  afterEach(() => {
    Date.now = originalDateNow;
    azureTokenManager.clear();
    jest.clearAllTimers();
    jest.useRealTimers();
  });

  describe('singleton behavior', () => {
    it('should return the same instance', () => {
      const instance1 = azureTokenManager;
      const instance2 = azureTokenManager;

      expect(instance1).toBe(instance2);
    });
  });

  describe('getToken', () => {
    it('should acquire token from Azure AD on first call', async () => {
      const token = await azureTokenManager.getToken(mockConfig);

      expect(token).toBe('mock-access-token-12345');
      expect(mockFetch).toHaveBeenCalledTimes(1);
      expect(mockFetch).toHaveBeenCalledWith(
        `https://login.microsoftonline.com/${mockConfig.tenantId}/oauth2/v2.0/token`,
        expect.objectContaining({
          method: 'POST',
          headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
        })
      );
    });

    it('should return cached token on subsequent calls', async () => {
      // First call - acquires token
      await azureTokenManager.getToken(mockConfig);

      // Second call - should use cache
      const token = await azureTokenManager.getToken(mockConfig);

      expect(token).toBe('mock-access-token-12345');
      expect(mockFetch).toHaveBeenCalledTimes(1); // Only called once
    });

    it('should refresh token when close to expiry (within 5 minutes)', async () => {
      // First call - acquires token
      await azureTokenManager.getToken(mockConfig);

      // Advance time to 56 minutes (4 minutes before expiry, within 5-min threshold)
      mockCurrentTime += 56 * 60 * 1000;

      // Setup new token response
      mockFetch.mockResolvedValue({
        ok: true,
        json: () => Promise.resolve(mockRefreshedTokenResponse)
      });

      // Should trigger refresh
      const token = await azureTokenManager.getToken(mockConfig);

      expect(token).toBe('refreshed-access-token-67890');
      expect(mockFetch).toHaveBeenCalledTimes(2);
    });

    it('should not refresh token when it has enough time left', async () => {
      // First call - acquires token
      await azureTokenManager.getToken(mockConfig);

      // Advance time to 30 minutes (30 minutes left, more than 5-min threshold)
      mockCurrentTime += 30 * 60 * 1000;

      const token = await azureTokenManager.getToken(mockConfig);

      expect(token).toBe('mock-access-token-12345');
      expect(mockFetch).toHaveBeenCalledTimes(1); // No refresh
    });

    it('should handle concurrent requests by sharing the same promise', async () => {
      // Start multiple concurrent requests
      const promise1 = azureTokenManager.getToken(mockConfig);
      const promise2 = azureTokenManager.getToken(mockConfig);
      const promise3 = azureTokenManager.getToken(mockConfig);

      const [token1, token2, token3] = await Promise.all([promise1, promise2, promise3]);

      expect(token1).toBe('mock-access-token-12345');
      expect(token2).toBe('mock-access-token-12345');
      expect(token3).toBe('mock-access-token-12345');

      // Should only call fetch once
      expect(mockFetch).toHaveBeenCalledTimes(1);
    });

    it('should throw error for invalid configuration - missing tenantId', async () => {
      const invalidConfig = { ...mockConfig, tenantId: '' };

      await expect(azureTokenManager.getToken(invalidConfig)).rejects.toThrow(
        'Azure AD Tenant ID is required'
      );
    });

    it('should throw error for invalid configuration - missing clientId', async () => {
      const invalidConfig = { ...mockConfig, clientId: '' };

      await expect(azureTokenManager.getToken(invalidConfig)).rejects.toThrow(
        'Azure AD Client ID is required'
      );
    });

    it('should throw error for invalid configuration - missing clientSecret', async () => {
      const invalidConfig = { ...mockConfig, clientSecret: '' };

      await expect(azureTokenManager.getToken(invalidConfig)).rejects.toThrow(
        'Azure AD Client Secret is required'
      );
    });

    it('should throw error for invalid configuration - missing scope', async () => {
      const invalidConfig = { ...mockConfig, scope: '' };

      await expect(azureTokenManager.getToken(invalidConfig)).rejects.toThrow(
        'Azure AD Scope is required'
      );
    });

    it('should handle network errors', async () => {
      mockFetch.mockRejectedValue(new Error('Network error'));

      await expect(azureTokenManager.getToken(mockConfig)).rejects.toThrow(
        'Azure AD authentication failed: Network error'
      );
    });

    it('should handle invalid_client error from Azure AD', async () => {
      mockFetch.mockResolvedValue({
        ok: false,
        status: 401,
        text: () => Promise.resolve(JSON.stringify({
          error: 'invalid_client',
          error_description: 'Client credentials are invalid'
        }))
      });

      await expect(azureTokenManager.getToken(mockConfig)).rejects.toThrow(
        'Azure AD authentication failed: Invalid client credentials'
      );
    });

    it('should handle invalid_scope error from Azure AD', async () => {
      mockFetch.mockResolvedValue({
        ok: false,
        status: 400,
        text: () => Promise.resolve(JSON.stringify({
          error: 'invalid_scope',
          error_description: 'Scope is invalid'
        }))
      });

      await expect(azureTokenManager.getToken(mockConfig)).rejects.toThrow(
        'Azure AD authentication failed: Invalid scope'
      );
    });

    it('should retry after failed concurrent refresh', async () => {
      mockFetch
        .mockRejectedValueOnce(new Error('First attempt failed'))
        .mockResolvedValueOnce({
          ok: true,
          json: () => Promise.resolve(mockTokenResponse)
        });

      // First call fails
      await expect(azureTokenManager.getToken(mockConfig)).rejects.toThrow('First attempt failed');

      // Second call should succeed
      const token = await azureTokenManager.getToken(mockConfig);

      expect(token).toBe('mock-access-token-12345');
      expect(mockFetch).toHaveBeenCalledTimes(2);
    });
  });

  describe('forceRefresh', () => {
    it('should force a new token acquisition', async () => {
      // First call - acquires token
      await azureTokenManager.getToken(mockConfig);

      // Setup refreshed token response
      mockFetch.mockResolvedValue({
        ok: true,
        json: () => Promise.resolve(mockRefreshedTokenResponse)
      });

      // Force refresh
      const token = await azureTokenManager.forceRefresh(mockConfig);

      expect(token).toBe('refreshed-access-token-67890');
      expect(mockFetch).toHaveBeenCalledTimes(2);
    });
  });

  describe('getTokenStatus', () => {
    it('should return correct status when no token', () => {
      const status = azureTokenManager.getTokenStatus();

      expect(status).toEqual({
        isInitialized: false,
        hasValidToken: false,
        timeUntilExpiry: null,
        timeUntilRefresh: null,
        isRefreshing: false,
        error: null
      });
    });

    it('should return correct status with valid token', async () => {
      await azureTokenManager.getToken(mockConfig);

      const status = azureTokenManager.getTokenStatus();

      expect(status.isInitialized).toBe(true);
      expect(status.hasValidToken).toBe(true);
      expect(status.timeUntilExpiry).toBe(3600 * 1000); // 1 hour in ms
      expect(status.timeUntilRefresh).toBe(3600 * 1000 - 5 * 60 * 1000); // 55 minutes
      expect(status.isRefreshing).toBe(false);
    });

    it('should show refreshing state during token acquisition', async () => {
      // Setup slow response
      mockFetch.mockImplementation(() => new Promise(resolve =>
        setTimeout(() => resolve({
          ok: true,
          json: () => Promise.resolve(mockTokenResponse)
        }), 100)
      ));

      // Start refresh (don't await)
      const refreshPromise = azureTokenManager.getToken(mockConfig);

      // Check status during refresh
      const status = azureTokenManager.getTokenStatus();
      expect(status.isRefreshing).toBe(true);

      // Wait for refresh to complete
      await refreshPromise;

      // Check status after refresh
      const statusAfter = azureTokenManager.getTokenStatus();
      expect(statusAfter.isRefreshing).toBe(false);
    });
  });

  describe('auto-refresh', () => {
    beforeEach(() => {
      jest.useFakeTimers();
    });

    it('should not start auto-refresh without cached token', () => {
      const consoleSpy = jest.spyOn(console, 'warn').mockImplementation();

      azureTokenManager.startAutoRefresh(mockConfig);

      expect(consoleSpy).toHaveBeenCalledWith(
        'AzureTokenManager: Cannot start auto-refresh - no cached token'
      );
      consoleSpy.mockRestore();
    });

    it('should start auto-refresh with cached token', async () => {
      await azureTokenManager.getToken(mockConfig);
      const consoleSpy = jest.spyOn(console, 'log').mockImplementation();

      azureTokenManager.startAutoRefresh(mockConfig);

      expect(consoleSpy).toHaveBeenCalledWith('AzureTokenManager: Auto-refresh started');
      consoleSpy.mockRestore();
    });

    it('should stop auto-refresh', async () => {
      await azureTokenManager.getToken(mockConfig);
      const consoleSpy = jest.spyOn(console, 'log').mockImplementation();

      azureTokenManager.startAutoRefresh(mockConfig);
      azureTokenManager.stopAutoRefresh();

      expect(consoleSpy).toHaveBeenCalledWith('AzureTokenManager: Auto-refresh stopped');
      consoleSpy.mockRestore();
    });
  });

  describe('clear', () => {
    it('should clear all state', async () => {
      await azureTokenManager.getToken(mockConfig);

      azureTokenManager.clear();

      const status = azureTokenManager.getTokenStatus();
      expect(status.isInitialized).toBe(false);
      expect(status.hasValidToken).toBe(false);
    });
  });

  describe('convenience functions', () => {
    it('getAzureToken should delegate to tokenManager', async () => {
      const token = await getAzureToken(mockConfig);

      expect(token).toBe('mock-access-token-12345');
    });

    it('getAzureTokenStatus should delegate to tokenManager', async () => {
      await getAzureToken(mockConfig);

      const status = getAzureTokenStatus();

      expect(status.isInitialized).toBe(true);
      expect(status.hasValidToken).toBe(true);
    });

    it('startAzureTokenAutoRefresh should delegate to tokenManager', async () => {
      await azureTokenManager.getToken(mockConfig);
      const startSpy = jest.spyOn(azureTokenManager, 'startAutoRefresh');

      startAzureTokenAutoRefresh(mockConfig);

      expect(startSpy).toHaveBeenCalledWith(mockConfig);
      startSpy.mockRestore();
    });

    it('stopAzureTokenAutoRefresh should delegate to tokenManager', () => {
      const stopSpy = jest.spyOn(azureTokenManager, 'stopAutoRefresh');

      stopAzureTokenAutoRefresh();

      expect(stopSpy).toHaveBeenCalled();
      stopSpy.mockRestore();
    });
  });

  describe('getAzureConfigFromEnv', () => {
    const originalEnv = process.env;

    beforeEach(() => {
      process.env = { ...originalEnv };
    });

    afterEach(() => {
      process.env = originalEnv;
    });

    it('should create config from environment variables', () => {
      process.env.REACT_APP_AZURE_TENANT_ID = 'env-tenant-id';
      process.env.REACT_APP_AZURE_CLIENT_ID = 'env-client-id';
      process.env.REACT_APP_AZURE_CLIENT_SECRET = 'env-client-secret';
      process.env.REACT_APP_AZURE_SCOPE = 'api://env-api/.default';

      const config = getAzureConfigFromEnv();

      expect(config).toEqual({
        tenantId: 'env-tenant-id',
        clientId: 'env-client-id',
        clientSecret: 'env-client-secret',
        scope: 'api://env-api/.default'
      });
    });

    it('should throw error when environment variables are missing', () => {
      delete process.env.REACT_APP_AZURE_TENANT_ID;

      expect(() => getAzureConfigFromEnv()).toThrow(
        'Azure AD configuration missing'
      );
    });
  });

  describe('request body format', () => {
    it('should use URLSearchParams format, not JSON', async () => {
      await azureTokenManager.getToken(mockConfig);

      const fetchCall = mockFetch.mock.calls[0];
      const requestBody = fetchCall[1].body;

      // Verify it's URLSearchParams format (contains key=value pairs)
      expect(requestBody).toContain('client_id=test-client-id');
      expect(requestBody).toContain('client_secret=test-client-secret');
      expect(requestBody).toContain('grant_type=client_credentials');
      expect(requestBody).toContain('scope=api%3A%2F%2Ftest-api%2F.default');

      // Verify header is correct
      expect(fetchCall[1].headers).toEqual({
        'Content-Type': 'application/x-www-form-urlencoded'
      });
    });
  });
});
