/**
 * Unit Tests for signalrService
 *
 * Tests SignalR connection management and error handling logic
 */

import { HubConnectionState } from '@microsoft/signalr';
import {
  initializeSignalR,
  isSignalRInitialized,
  getSignalRConnection,
  getConnectionState,
  isConnected,
  sendMessage,
  onMessage,
  offMessage,
  stopSignalR,
  getCurrentConfig
} from '../signalrService';
import { SignalRConfig } from '../../types/signalr.types';

// Mock @microsoft/signalr
const mockStart = jest.fn();
const mockStop = jest.fn();
const mockInvoke = jest.fn();
const mockOn = jest.fn();
const mockOff = jest.fn();
const mockOnclose = jest.fn();
const mockOnreconnecting = jest.fn();
const mockOnreconnected = jest.fn();

const mockConnection = {
  start: mockStart,
  stop: mockStop,
  invoke: mockInvoke,
  on: mockOn,
  off: mockOff,
  onclose: mockOnclose,
  onreconnecting: mockOnreconnecting,
  onreconnected: mockOnreconnected,
  state: HubConnectionState.Disconnected
};

const mockHubConnectionBuilder = {
  withUrl: jest.fn().mockReturnThis(),
  withAutomaticReconnect: jest.fn().mockReturnThis(),
  configureLogging: jest.fn().mockReturnThis(),
  build: jest.fn().mockReturnValue(mockConnection)
};

jest.mock('@microsoft/signalr', () => ({
  HubConnectionBuilder: jest.fn(() => mockHubConnectionBuilder),
  HubConnectionState: {
    Disconnected: 'Disconnected',
    Connecting: 'Connecting',
    Connected: 'Connected',
    Disconnecting: 'Disconnecting',
    Reconnecting: 'Reconnecting'
  },
  LogLevel: {
    Information: 'Information'
  }
}));

// Mock error handler
jest.mock('../../utils/errorHandler', () => ({
  handleOfficeError: jest.fn((_context, error) => error?.message || 'Mocked error'),
  logError: jest.fn()
}));

describe('signalrService', () => {
  let mockConfig: SignalRConfig;
  let originalOffice: any;

  beforeEach(() => {
    // Reset all mocks
    jest.clearAllMocks();
    
    // Mock Office.js context
    originalOffice = (global as any).Office;
    (global as any).Office = {
      context: {
        mailbox: {},
        ui: {}
      }
    };

    mockConfig = {
      hubUrl: 'https://test-hub.com/notifications',
      accessToken: 'test-token',
      reconnectPolicy: [0, 1000, 5000]
    };

    // Reset connection state
    mockConnection.state = HubConnectionState.Disconnected;
    mockStart.mockResolvedValue(undefined);
    mockStop.mockResolvedValue(undefined);
    mockInvoke.mockResolvedValue(undefined);
  });

  afterEach(() => {
    // Restore original Office context
    (global as any).Office = originalOffice;
    
    // Stop any connections
    stopSignalR();
  });

  describe('initializeSignalR', () => {
    it('should throw error if Office.js not initialized', async () => {
      // Mock Office.context as null
      (global as any).Office = { context: null };
      
      await expect(initializeSignalR(mockConfig)).rejects.toThrow(
        'Office.js must be initialized before SignalR'
      );
    });

    it('should establish connection successfully', async () => {
      mockConnection.state = HubConnectionState.Connected;
      
      const result = await initializeSignalR(mockConfig);
      
      expect(result).toBe(mockConnection);
      expect(mockHubConnectionBuilder.withUrl).toHaveBeenCalledWith(
        'https://test-hub.com/notifications',
        { accessTokenFactory: expect.any(Function) }
      );
      expect(mockHubConnectionBuilder.withAutomaticReconnect).toHaveBeenCalledWith([0, 1000, 5000]);
      expect(mockStart).toHaveBeenCalled();
    });

    it('should return existing connection if already initialized', async () => {
      // Initialize once
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      // Clear mocks and try again
      jest.clearAllMocks();
      
      const result = await initializeSignalR(mockConfig);
      
      expect(result).toBe(mockConnection);
      expect(mockStart).not.toHaveBeenCalled(); // Should not start again
    });

    it('should handle connection errors gracefully', async () => {
      const error = new Error('Network error');
      mockStart.mockRejectedValue(error);
      
      await expect(initializeSignalR(mockConfig)).rejects.toThrow('Network error');
    });

    it('should register connection lifecycle handlers', async () => {
      mockConnection.state = HubConnectionState.Connected;
      
      await initializeSignalR(mockConfig);
      
      expect(mockOnclose).toHaveBeenCalled();
      expect(mockOnreconnecting).toHaveBeenCalled();
      expect(mockOnreconnected).toHaveBeenCalled();
    });
  });

  describe('isSignalRInitialized', () => {
    it('should return false when not initialized', () => {
      expect(isSignalRInitialized()).toBe(false);
    });

    it('should return true when initialized', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      expect(isSignalRInitialized()).toBe(true);
    });
  });

  describe('getSignalRConnection', () => {
    it('should throw error when not initialized', () => {
      expect(() => getSignalRConnection()).toThrow(
        'SignalR is not initialized. Call initializeSignalR() first.'
      );
    });

    it('should return connection when initialized', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      expect(getSignalRConnection()).toBe(mockConnection);
    });
  });

  describe('getConnectionState', () => {
    it('should return Disconnected when not initialized', () => {
      expect(getConnectionState()).toBe(HubConnectionState.Disconnected);
    });

    it('should return current connection state when initialized', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      expect(getConnectionState()).toBe(HubConnectionState.Connected);
    });
  });

  describe('isConnected', () => {
    it('should return false when not connected', () => {
      expect(isConnected()).toBe(false);
    });

    it('should return true when connected', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      expect(isConnected()).toBe(true);
    });
  });

  describe('sendMessage', () => {
    beforeEach(async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
    });

    it('should send message successfully when connected', async () => {
      await sendMessage('TestMethod', { data: 'test' });
      
      expect(mockInvoke).toHaveBeenCalledWith('TestMethod', { data: 'test' });
    });

    it('should throw error when not connected', async () => {
      mockConnection.state = HubConnectionState.Disconnected;
      
      await expect(sendMessage('TestMethod', {})).rejects.toThrow(
        'SignalR is not connected'
      );
    });

    it('should throw error when not initialized', async () => {
      await stopSignalR();
      
      await expect(sendMessage('TestMethod', {})).rejects.toThrow(
        'SignalR is not initialized'
      );
    });

    it('should handle send errors gracefully', async () => {
      const error = new Error('Send failed');
      mockInvoke.mockRejectedValue(error);
      
      await expect(sendMessage('TestMethod', {})).rejects.toThrow('Send failed');
    });
  });

  describe('onMessage', () => {
    beforeEach(async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
    });

    it('should register message handler', () => {
      const handler = jest.fn();
      
      onMessage('TestEvent', handler);
      
      expect(mockOn).toHaveBeenCalledWith('TestEvent', handler);
    });

    it('should throw error when not initialized', async () => {
      await stopSignalR();
      
      expect(() => onMessage('TestEvent', jest.fn())).toThrow(
        'SignalR is not initialized'
      );
    });
  });

  describe('offMessage', () => {
    beforeEach(async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
    });

    it('should remove specific message handler', () => {
      const handler = jest.fn();
      
      offMessage('TestEvent', handler);
      
      expect(mockOff).toHaveBeenCalledWith('TestEvent', handler);
    });

    it('should remove all handlers for event when no handler specified', () => {
      offMessage('TestEvent');
      
      expect(mockOff).toHaveBeenCalledWith('TestEvent');
    });

    it('should handle gracefully when not initialized', () => {
      stopSignalR();
      
      expect(() => offMessage('TestEvent')).not.toThrow();
    });
  });

  describe('stopSignalR', () => {
    it('should stop connection gracefully', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      await stopSignalR();
      
      expect(mockStop).toHaveBeenCalled();
      expect(isSignalRInitialized()).toBe(false);
    });

    it('should handle stop errors gracefully', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      const error = new Error('Stop failed');
      mockStop.mockRejectedValue(error);
      
      await expect(stopSignalR()).resolves.not.toThrow();
    });

    it('should handle gracefully when not initialized', async () => {
      await expect(stopSignalR()).resolves.not.toThrow();
    });
  });

  describe('getCurrentConfig', () => {
    it('should return null when not initialized', () => {
      expect(getCurrentConfig()).toBeNull();
    });

    it('should return current config when initialized', async () => {
      mockConnection.state = HubConnectionState.Connected;
      await initializeSignalR(mockConfig);
      
      expect(getCurrentConfig()).toEqual(mockConfig);
    });
  });
});