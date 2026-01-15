/**
 * Unit Tests for negotiateService
 *
 * Tests SignalR negotiate endpoint handling, retry logic,
 * response validation, and error handling
 */

import {
  negotiate,
  getNegotiateConfigFromEnv,
  isNegotiateConfigured
} from '../negotiateService';
import { NegotiateConfig } from '../../types/signalr.types';

// Mock the error handler
jest.mock('../../utils/errorHandler');

// Mock global fetch
const mockFetch = jest.fn();
global.fetch = mockFetch;

describe('negotiateService', () => {
  const mockConfig: NegotiateConfig = {
    negotiateUrl: 'https://api.example.com/signalr/negotiate',
    maxRetries: 3,
    retryDelayMs: 100 // Short delay for tests
  };

  const mockBearerToken = 'mock-bearer-token-12345';

  const mockNegotiateResponse = {
    url: 'wss://signalr-hub.azure.com/client/?hub=notifications',
    accessToken: 'negotiate-access-token-67890',
    availableTransports: [
      { transport: 'WebSockets', transferFormats: ['Text', 'Binary'] }
    ]
  };

  beforeEach(() => {
    jest.clearAllMocks();
    mockFetch.mockReset();

    // Setup default successful response
    mockFetch.mockResolvedValue({
      ok: true,
      json: () => Promise.resolve(mockNegotiateResponse)
    });
  });

  describe('negotiate', () => {
    it('should successfully negotiate and return connection info', async () => {
      const result = await negotiate(mockConfig, mockBearerToken);

      expect(result).toEqual({
        url: mockNegotiateResponse.url,
        accessToken: mockNegotiateResponse.accessToken
      });
    });

    it('should call negotiate endpoint with Bearer token', async () => {
      await negotiate(mockConfig, mockBearerToken);

      expect(mockFetch).toHaveBeenCalledWith(
        mockConfig.negotiateUrl,
        expect.objectContaining({
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${mockBearerToken}`,
            'Content-Type': 'application/json'
          }
        })
      );
    });

    it('should throw error when negotiateUrl is missing', async () => {
      const invalidConfig = { ...mockConfig, negotiateUrl: '' };

      await expect(negotiate(invalidConfig, mockBearerToken)).rejects.toThrow(
        'Negotiate URL is required'
      );
    });

    it('should throw error when bearerToken is missing', async () => {
      await expect(negotiate(mockConfig, '')).rejects.toThrow(
        'Bearer token is required for negotiate'
      );
    });

    describe('response validation', () => {
      it('should throw error when response is missing url', async () => {
        mockFetch.mockResolvedValue({
          ok: true,
          json: () => Promise.resolve({
            accessToken: 'some-token'
          })
        });

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Invalid negotiate response: missing or empty url or accessToken'
        );
      });

      it('should throw error when response is missing accessToken', async () => {
        mockFetch.mockResolvedValue({
          ok: true,
          json: () => Promise.resolve({
            url: 'wss://some-url.com'
          })
        });

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Invalid negotiate response: missing or empty url or accessToken'
        );
      });

      it('should throw error when url is empty string', async () => {
        mockFetch.mockResolvedValue({
          ok: true,
          json: () => Promise.resolve({
            url: '',
            accessToken: 'some-token'
          })
        });

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Invalid negotiate response: missing or empty url or accessToken'
        );
      });

      it('should throw error when accessToken is empty string', async () => {
        mockFetch.mockResolvedValue({
          ok: true,
          json: () => Promise.resolve({
            url: 'wss://some-url.com',
            accessToken: ''
          })
        });

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Invalid negotiate response: missing or empty url or accessToken'
        );
      });
    });

    describe('retry logic', () => {
      it('should retry on transient errors', async () => {
        mockFetch
          .mockRejectedValueOnce(new Error('Network error'))
          .mockRejectedValueOnce(new Error('Network error'))
          .mockResolvedValueOnce({
            ok: true,
            json: () => Promise.resolve(mockNegotiateResponse)
          });

        const result = await negotiate(mockConfig, mockBearerToken);

        expect(result).toEqual({
          url: mockNegotiateResponse.url,
          accessToken: mockNegotiateResponse.accessToken
        });
        expect(mockFetch).toHaveBeenCalledTimes(3);
      });

      it('should retry on 500 server errors', async () => {
        mockFetch
          .mockResolvedValueOnce({
            ok: false,
            status: 500,
            text: () => Promise.resolve('Internal Server Error')
          })
          .mockResolvedValueOnce({
            ok: true,
            json: () => Promise.resolve(mockNegotiateResponse)
          });

        const result = await negotiate(mockConfig, mockBearerToken);

        expect(result).toEqual({
          url: mockNegotiateResponse.url,
          accessToken: mockNegotiateResponse.accessToken
        });
        expect(mockFetch).toHaveBeenCalledTimes(2);
      });

      it('should fail after max retries exceeded', async () => {
        mockFetch.mockRejectedValue(new Error('Persistent network error'));

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'SignalR negotiation failed after 3 attempts: Persistent network error'
        );

        expect(mockFetch).toHaveBeenCalledTimes(3);
      });
    });

    describe('authentication errors (no retry)', () => {
      it('should NOT retry on 401 Unauthorized', async () => {
        mockFetch.mockResolvedValue({
          ok: false,
          status: 401,
          text: () => Promise.resolve('Unauthorized')
        });

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Authentication failed during SignalR negotiation (401)'
        );

        // Should only attempt once - no retry on auth errors
        expect(mockFetch).toHaveBeenCalledTimes(1);
      });

      it('should NOT retry on 403 Forbidden', async () => {
        mockFetch.mockResolvedValue({
          ok: false,
          status: 403,
          text: () => Promise.resolve('Forbidden')
        });

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Authentication failed during SignalR negotiation (403)'
        );

        // Should only attempt once - no retry on auth errors
        expect(mockFetch).toHaveBeenCalledTimes(1);
      });

      it('should NOT retry when error contains "authentication failed"', async () => {
        mockFetch.mockRejectedValue(new Error('Authentication failed: invalid token'));

        await expect(negotiate(mockConfig, mockBearerToken)).rejects.toThrow(
          'Authentication failed: invalid token'
        );

        expect(mockFetch).toHaveBeenCalledTimes(1);
      });
    });

    describe('exponential backoff', () => {
      it('should use exponential backoff between retries', async () => {
        jest.useFakeTimers();

        mockFetch
          .mockRejectedValueOnce(new Error('Error 1'))
          .mockRejectedValueOnce(new Error('Error 2'))
          .mockResolvedValueOnce({
            ok: true,
            json: () => Promise.resolve(mockNegotiateResponse)
          });

        const negotiatePromise = negotiate(mockConfig, mockBearerToken);

        // First attempt happens immediately
        expect(mockFetch).toHaveBeenCalledTimes(1);

        // Advance time for first retry (100ms * 2^0 = 100ms)
        await jest.advanceTimersByTimeAsync(100);
        expect(mockFetch).toHaveBeenCalledTimes(2);

        // Advance time for second retry (100ms * 2^1 = 200ms)
        await jest.advanceTimersByTimeAsync(200);
        expect(mockFetch).toHaveBeenCalledTimes(3);

        const result = await negotiatePromise;
        expect(result).toEqual({
          url: mockNegotiateResponse.url,
          accessToken: mockNegotiateResponse.accessToken
        });

        jest.useRealTimers();
      });
    });

    describe('custom retry configuration', () => {
      it('should respect custom maxRetries', async () => {
        const customConfig = { ...mockConfig, maxRetries: 1 };
        mockFetch.mockRejectedValue(new Error('Persistent error'));

        await expect(negotiate(customConfig, mockBearerToken)).rejects.toThrow(
          'SignalR negotiation failed after 1 attempts'
        );

        expect(mockFetch).toHaveBeenCalledTimes(1);
      });

      it('should use default maxRetries when not specified', async () => {
        const minimalConfig: NegotiateConfig = {
          negotiateUrl: mockConfig.negotiateUrl
        };
        mockFetch.mockRejectedValue(new Error('Persistent error'));

        await expect(negotiate(minimalConfig, mockBearerToken)).rejects.toThrow(
          'SignalR negotiation failed after 3 attempts'
        );

        expect(mockFetch).toHaveBeenCalledTimes(3);
      });
    });
  });

  describe('getNegotiateConfigFromEnv', () => {
    const originalEnv = process.env;

    beforeEach(() => {
      process.env = { ...originalEnv };
    });

    afterEach(() => {
      process.env = originalEnv;
    });

    it('should return config when negotiate URL is set', () => {
      process.env.REACT_APP_SIGNALR_NEGOTIATE_URL = 'https://api.example.com/negotiate';

      const config = getNegotiateConfigFromEnv();

      expect(config).toEqual({
        negotiateUrl: 'https://api.example.com/negotiate',
        maxRetries: 3,
        retryDelayMs: 1000
      });
    });

    it('should return null when negotiate URL is not set', () => {
      delete process.env.REACT_APP_SIGNALR_NEGOTIATE_URL;

      const config = getNegotiateConfigFromEnv();

      expect(config).toBeNull();
    });

    it('should return null when negotiate URL is empty string', () => {
      process.env.REACT_APP_SIGNALR_NEGOTIATE_URL = '';

      const config = getNegotiateConfigFromEnv();

      expect(config).toBeNull();
    });
  });

  describe('isNegotiateConfigured', () => {
    const originalEnv = process.env;

    beforeEach(() => {
      process.env = { ...originalEnv };
    });

    afterEach(() => {
      process.env = originalEnv;
    });

    it('should return true when negotiate URL is configured', () => {
      process.env.REACT_APP_SIGNALR_NEGOTIATE_URL = 'https://api.example.com/negotiate';

      expect(isNegotiateConfigured()).toBe(true);
    });

    it('should return false when negotiate URL is not configured', () => {
      delete process.env.REACT_APP_SIGNALR_NEGOTIATE_URL;

      expect(isNegotiateConfigured()).toBe(false);
    });
  });
});
