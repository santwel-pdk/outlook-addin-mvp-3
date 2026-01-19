/**
 * Unit Tests for useSignalR Hook
 *
 * Tests SignalR hook behavior including:
 * - Connection initialization
 * - Handler registration (MUST be BEFORE connection.start())
 * - State management
 * - Error handling
 */

import { renderHook, act, waitFor } from '@testing-library/react';
import { HubConnectionState } from '@microsoft/signalr';
import { useSignalR } from '../useSignalR';

// Mock signalrService
const mockInitializeSignalR = jest.fn();
const mockIsSignalRInitialized = jest.fn();
const mockOffMessage = jest.fn();
const mockGetConnectionState = jest.fn();

jest.mock('../../services/signalrService', () => ({
  initializeSignalR: (...args: any[]) => mockInitializeSignalR(...args),
  isSignalRInitialized: () => mockIsSignalRInitialized(),
  offMessage: (...args: any[]) => mockOffMessage(...args),
  getConnectionState: () => mockGetConnectionState()
}));

// Mock useOfficeContext
const mockUseOfficeContext = jest.fn();
jest.mock('../useOfficeContext', () => ({
  useOfficeContext: () => mockUseOfficeContext()
}));

// Mock environment variables
const originalEnv = process.env;

describe('useSignalR', () => {
  beforeEach(() => {
    jest.clearAllMocks();

    // Setup default mocks
    mockIsSignalRInitialized.mockReturnValue(false);
    mockGetConnectionState.mockReturnValue(HubConnectionState.Disconnected);
    mockUseOfficeContext.mockReturnValue({ isInitialized: true });

    // Setup environment variables
    process.env = {
      ...originalEnv,
      REACT_APP_SIGNALR_HUB_URL: 'https://test-hub.com/notifications',
      REACT_APP_SIGNALR_ACCESS_TOKEN: 'test-token'
    };
  });

  afterEach(() => {
    process.env = originalEnv;
  });

  describe('initialization', () => {
    it('should wait for Office.js before connecting', () => {
      mockUseOfficeContext.mockReturnValue({ isInitialized: false });

      const { result } = renderHook(() => useSignalR());

      expect(result.current.isConnected).toBe(false);
      expect(result.current.connectionState).toBe(HubConnectionState.Disconnected);
      expect(mockInitializeSignalR).not.toHaveBeenCalled();
    });

    it('should initialize SignalR when Office.js is ready', async () => {
      const mockConnection = { state: HubConnectionState.Connected };
      mockInitializeSignalR.mockResolvedValue(mockConnection);

      const { result } = renderHook(() => useSignalR());

      await waitFor(() => {
        expect(mockInitializeSignalR).toHaveBeenCalled();
      });

      await waitFor(() => {
        expect(result.current.isConnected).toBe(true);
      });
    });

    it('should pass handlers in config for registration BEFORE start()', async () => {
      const mockConnection = { state: HubConnectionState.Connected };
      mockInitializeSignalR.mockResolvedValue(mockConnection);

      renderHook(() => useSignalR());

      await waitFor(() => {
        expect(mockInitializeSignalR).toHaveBeenCalled();
      });

      // Verify handlers were passed in config
      const configArg = mockInitializeSignalR.mock.calls[0][0];
      expect(configArg.handlers).toBeDefined();
      expect(configArg.handlers.length).toBe(2);
      expect(configArg.handlers[0].methodName).toBe('NotificationReceived');
      expect(configArg.handlers[1].methodName).toBe('BroadcastMessage');
    });

    it('should skip initialization if already initialized', async () => {
      mockIsSignalRInitialized.mockReturnValue(true);
      mockGetConnectionState.mockReturnValue(HubConnectionState.Connected);

      const { result } = renderHook(() => useSignalR());

      await waitFor(() => {
        expect(result.current.isConnected).toBe(true);
      });

      expect(mockInitializeSignalR).not.toHaveBeenCalled();
    });
  });

  describe('error handling', () => {
    it('should handle missing hub URL', async () => {
      process.env.REACT_APP_SIGNALR_HUB_URL = '';

      const { result } = renderHook(() => useSignalR());

      await waitFor(() => {
        expect(result.current.error).toContain('SignalR configuration missing');
      });

      expect(result.current.isConnected).toBe(false);
    });

    it('should handle connection errors', async () => {
      const error = new Error('Connection failed');
      mockInitializeSignalR.mockRejectedValue(error);

      const { result } = renderHook(() => useSignalR());

      await waitFor(() => {
        expect(result.current.error).toBe('Connection failed');
      });

      expect(result.current.isConnected).toBe(false);
      expect(result.current.connectionState).toBe(HubConnectionState.Disconnected);
    });
  });

  describe('message handling', () => {
    it('should update lastMessage when handler is invoked', async () => {
      const mockConnection = { state: HubConnectionState.Connected };
      mockInitializeSignalR.mockResolvedValue(mockConnection);

      const { result } = renderHook(() => useSignalR());

      await waitFor(() => {
        expect(mockInitializeSignalR).toHaveBeenCalled();
      });

      // Get the handler that was passed to initializeSignalR
      const configArg = mockInitializeSignalR.mock.calls[0][0];
      const messageHandler = configArg.handlers[0].handler;

      // Simulate receiving a message
      const testMessage = {
        type: 'notification',
        payload: { text: 'Hello' },
        timestamp: new Date().toISOString(),
        id: 'msg-1'
      };

      act(() => {
        messageHandler(testMessage);
      });

      await waitFor(() => {
        expect(result.current.lastMessage).not.toBeNull();
        expect(result.current.lastMessage?.type).toBe('notification');
        expect(result.current.lastMessage?.id).toBe('msg-1');
      });
    });
  });

  describe('cleanup', () => {
    it('should clean up handlers on unmount', async () => {
      const mockConnection = { state: HubConnectionState.Connected };
      mockInitializeSignalR.mockResolvedValue(mockConnection);

      const { unmount } = renderHook(() => useSignalR());

      await waitFor(() => {
        expect(mockInitializeSignalR).toHaveBeenCalled();
      });

      unmount();

      expect(mockOffMessage).toHaveBeenCalledWith('NotificationReceived');
      expect(mockOffMessage).toHaveBeenCalledWith('BroadcastMessage');
    });
  });

  describe('connection state monitoring', () => {
    it('should update connection state periodically', async () => {
      jest.useFakeTimers();

      mockIsSignalRInitialized.mockReturnValue(true);
      mockGetConnectionState.mockReturnValue(HubConnectionState.Connected);

      const { result } = renderHook(() => useSignalR());

      // Fast-forward time to trigger the interval
      act(() => {
        jest.advanceTimersByTime(2000);
      });

      await waitFor(() => {
        expect(result.current.connectionState).toBe(HubConnectionState.Connected);
      });

      // Simulate disconnection
      mockGetConnectionState.mockReturnValue(HubConnectionState.Disconnected);

      act(() => {
        jest.advanceTimersByTime(2000);
      });

      await waitFor(() => {
        expect(result.current.connectionState).toBe(HubConnectionState.Disconnected);
        expect(result.current.isConnected).toBe(false);
      });

      jest.useRealTimers();
    });
  });
});
