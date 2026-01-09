/**
 * Unit Tests for ssoService
 *
 * Tests SSO authentication logic and error handling
 */

import {
  initializeSSO,
  isSSOInitialized,
  isSSOAuthenticated,
  getSSOState,
  getCurrentToken,
  getCurrentUser,
  isTokenValid,
  getTimeUntilExpiry,
  refreshToken,
  clearSSOState
} from '../ssoService';
import { SSOConfig, SSOErrorCode } from '../../types/sso.types';

// Mock Office.js globals
const mockGetAccessToken = jest.fn();
const originalOffice = (global as any).Office;
const originalOfficeRuntime = (global as any).OfficeRuntime;

// Mock error handler
jest.mock('../../utils/errorHandler', () => ({
  logError: jest.fn()
}));

// Sample JWT token for testing (base64 encoded)
const createMockJWT = (exp: number, sub = 'test-user-id', name = 'Test User', email = 'test@example.com') => {
  const header = btoa(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = btoa(JSON.stringify({
    exp,
    iat: Math.floor(Date.now() / 1000),
    sub,
    name,
    preferred_username: email,
    email,
    aud: 'test-audience',
    iss: 'https://login.microsoftonline.com/test-tenant',
    tid: 'test-tenant-id',
    scp: 'User.Read Mail.Read'
  }));
  const signature = 'mock-signature';
  return `${header}.${payload}.${signature}`;
};

describe('ssoService', () => {
  let mockConfig: SSOConfig;

  beforeEach(() => {
    // Reset all mocks
    jest.clearAllMocks();
    
    // Mock Office.js context
    (global as any).Office = {
      context: {
        mailbox: {},
        ui: {}
      }
    };

    // Mock OfficeRuntime.auth
    (global as any).OfficeRuntime = {
      auth: {
        getAccessToken: mockGetAccessToken
      }
    };

    mockConfig = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: false
    };

    // Clear SSO state before each test
    clearSSOState();

    // Mock successful token response
    const futureExp = Math.floor(Date.now() / 1000) + 3600; // 1 hour from now
    mockGetAccessToken.mockResolvedValue(createMockJWT(futureExp));
  });

  afterEach(() => {
    // Restore original globals
    (global as any).Office = originalOffice;
    (global as any).OfficeRuntime = originalOfficeRuntime;
    clearSSOState();
  });

  describe('initializeSSO', () => {
    it('should throw error if Office.js not initialized', async () => {
      (global as any).Office = { context: null };
      
      await expect(initializeSSO(mockConfig)).rejects.toThrow(
        'Office.js must be initialized before SSO'
      );
    });

    it('should initialize successfully with valid token', async () => {
      const result = await initializeSSO(mockConfig);
      
      expect(result.isInitialized).toBe(true);
      expect(result.isAuthenticated).toBe(true);
      expect(result.token).toBeTruthy();
      expect(result.user).toBeTruthy();
      expect(result.user?.email).toBe('test@example.com');
      expect(mockGetAccessToken).toHaveBeenCalledWith(mockConfig);
    });

    it('should return existing state if already initialized and authenticated', async () => {
      // Initialize once
      await initializeSSO(mockConfig);
      
      // Clear mocks and try again
      jest.clearAllMocks();
      
      const result = await initializeSSO(mockConfig);
      
      expect(result.isAuthenticated).toBe(true);
      expect(mockGetAccessToken).not.toHaveBeenCalled(); // Should not call again
    });

    it('should handle user not signed in error (13001)', async () => {
      const error = { code: SSOErrorCode.USER_NOT_SIGNED_IN, message: 'User not signed in' };
      mockGetAccessToken.mockRejectedValue(error);
      
      await expect(initializeSSO(mockConfig)).rejects.toMatchObject({
        code: SSOErrorCode.USER_NOT_SIGNED_IN,
        message: 'User is not signed in to Office. Please sign in and try again.'
      });
    });

    it('should handle user aborted consent error (13002)', async () => {
      const error = { code: SSOErrorCode.USER_ABORTED_CONSENT, message: 'User aborted' };
      mockGetAccessToken.mockRejectedValue(error);
      
      await expect(initializeSSO(mockConfig)).rejects.toMatchObject({
        code: SSOErrorCode.USER_ABORTED_CONSENT,
        message: 'User cancelled the consent dialog. Please try again and accept the permissions.'
      });
    });

    it('should handle identity API not supported error (13000)', async () => {
      const error = { code: SSOErrorCode.IDENTITY_API_NOT_SUPPORTED, message: 'Identity API not supported' };
      mockGetAccessToken.mockRejectedValue(error);
      
      await expect(initializeSSO(mockConfig)).rejects.toMatchObject({
        code: SSOErrorCode.IDENTITY_API_NOT_SUPPORTED,
        message: 'The identity API is not supported for this add-in. Please check the manifest configuration.'
      });
    });

    it('should handle admin consent required error (13012)', async () => {
      const error = { code: SSOErrorCode.ADMIN_CONSENT_REQUIRED, message: 'Admin consent required' };
      mockGetAccessToken.mockRejectedValue(error);
      
      await expect(initializeSSO(mockConfig)).rejects.toMatchObject({
        code: SSOErrorCode.ADMIN_CONSENT_REQUIRED,
        message: 'Admin consent is required for this application. Please contact your administrator.'
      });
    });

    it('should handle API not available error (13006)', async () => {
      const error = { code: SSOErrorCode.API_NOT_AVAILABLE, message: 'API not available' };
      mockGetAccessToken.mockRejectedValue(error);
      
      await expect(initializeSSO(mockConfig)).rejects.toMatchObject({
        code: SSOErrorCode.API_NOT_AVAILABLE,
        message: 'SSO API is not available in the current Office host or version.'
      });
    });

    it('should use default config when none provided', async () => {
      await initializeSSO();
      
      expect(mockGetAccessToken).toHaveBeenCalledWith({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: false
      });
    });
  });

  describe('JWT token parsing', () => {
    it('should correctly parse valid JWT token', async () => {
      const futureExp = Math.floor(Date.now() / 1000) + 3600;
      const mockToken = createMockJWT(futureExp, 'user123', 'John Doe', 'john@contoso.com');
      mockGetAccessToken.mockResolvedValue(mockToken);
      
      const result = await initializeSSO(mockConfig);
      
      expect(result.token?.userId).toBe('user123');
      expect(result.token?.expiresAt).toBe(futureExp * 1000);
      expect(result.user?.displayName).toBe('John Doe');
      expect(result.user?.email).toBe('john@contoso.com');
    });

    it('should handle malformed JWT token gracefully', async () => {
      mockGetAccessToken.mockResolvedValue('invalid.jwt.token');
      
      const result = await initializeSSO(mockConfig);
      
      // Should still succeed but with fallback values
      expect(result.isAuthenticated).toBe(true);
      expect(result.token?.userId).toBe('unknown');
      expect(result.user?.displayName).toBe('Unknown User');
    });

    it('should handle JWT without expiry claim', async () => {
      const header = btoa(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
      const payload = btoa(JSON.stringify({ sub: 'test-user' })); // No exp claim
      const signature = 'mock-signature';
      const tokenWithoutExp = `${header}.${payload}.${signature}`;
      
      mockGetAccessToken.mockResolvedValue(tokenWithoutExp);
      
      const result = await initializeSSO(mockConfig);
      
      expect(result.isAuthenticated).toBe(true);
      expect(result.token?.expiresAt).toBeGreaterThan(Date.now());
    });
  });

  describe('state management functions', () => {
    beforeEach(async () => {
      await initializeSSO(mockConfig);
    });

    it('should report correct initialization state', () => {
      expect(isSSOInitialized()).toBe(true);
      expect(isSSOAuthenticated()).toBe(true);
    });

    it('should return current token', () => {
      const token = getCurrentToken();
      expect(token).toBeTruthy();
      expect(typeof token).toBe('string');
    });

    it('should return current user', () => {
      const user = getCurrentUser();
      expect(user).toBeTruthy();
      expect(user?.email).toBe('test@example.com');
    });

    it('should validate token correctly', () => {
      expect(isTokenValid()).toBe(true);
    });

    it('should return time until expiry', () => {
      const timeUntilExpiry = getTimeUntilExpiry();
      expect(timeUntilExpiry).toBeGreaterThan(0);
      expect(timeUntilExpiry).toBeLessThan(3600000); // Less than 1 hour
    });
  });

  describe('token refresh', () => {
    beforeEach(async () => {
      await initializeSSO(mockConfig);
    });

    it('should refresh token successfully', async () => {
      const newExp = Math.floor(Date.now() / 1000) + 7200; // 2 hours from now
      const newToken = createMockJWT(newExp);
      mockGetAccessToken.mockResolvedValue(newToken);
      
      const refreshedToken = await refreshToken();
      
      expect(refreshedToken).toBe(newToken);
      expect(isSSOAuthenticated()).toBe(true);
    });

    it('should handle refresh errors', async () => {
      const error = { code: SSOErrorCode.USER_NOT_SIGNED_IN, message: 'Session expired' };
      mockGetAccessToken.mockRejectedValue(error);
      
      await expect(refreshToken()).rejects.toMatchObject({
        code: SSOErrorCode.USER_NOT_SIGNED_IN
      });
    });

    it('should use non-prompt config during refresh', async () => {
      await refreshToken();
      
      expect(mockGetAccessToken).toHaveBeenCalledWith({
        allowSignInPrompt: false,
        allowConsentPrompt: false,
        forMSGraphAccess: false
      });
    });
  });

  describe('token expiration edge cases', () => {
    it('should consider token invalid when expired', async () => {
      const pastExp = Math.floor(Date.now() / 1000) - 3600; // 1 hour ago
      const expiredToken = createMockJWT(pastExp);
      mockGetAccessToken.mockResolvedValue(expiredToken);
      
      await initializeSSO(mockConfig);
      
      expect(isTokenValid()).toBe(false);
      expect(getTimeUntilExpiry()).toBe(0);
    });

    it('should consider token invalid when close to expiry', async () => {
      const soonExp = Math.floor(Date.now() / 1000) + 60; // 1 minute from now (less than 5 min buffer)
      const soonToExpireToken = createMockJWT(soonExp);
      mockGetAccessToken.mockResolvedValue(soonToExpireToken);
      
      await initializeSSO(mockConfig);
      
      expect(isTokenValid()).toBe(false); // Should be false due to 5-minute buffer
    });
  });

  describe('clearSSOState', () => {
    it('should clear all state', async () => {
      await initializeSSO(mockConfig);
      expect(isSSOAuthenticated()).toBe(true);
      
      clearSSOState();
      
      expect(isSSOInitialized()).toBe(false);
      expect(isSSOAuthenticated()).toBe(false);
      expect(getCurrentToken()).toBe(null);
      expect(getCurrentUser()).toBe(null);
    });
  });

  describe('error scenarios', () => {
    it('should handle network errors gracefully', async () => {
      const networkError = new Error('Network error');
      mockGetAccessToken.mockRejectedValue(networkError);
      
      await expect(initializeSSO(mockConfig)).rejects.toMatchObject({
        message: 'Authentication failed: Network error'
      });
    });

    it('should handle missing Office.js context during refresh', async () => {
      await initializeSSO(mockConfig);
      (global as any).Office = { context: null };
      
      await expect(refreshToken()).rejects.toThrow('Office.js context not available');
    });
  });
});