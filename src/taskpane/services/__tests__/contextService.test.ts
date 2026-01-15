/**
 * Unit Tests for contextService
 *
 * Tests context detection and ItemChanged event handling
 */

import {
  getContextType,
  getFullContextState,
  subscribeToContextChanges,
  getCurrentItemId
} from '../contextService';
import { OutlookContextType } from '../../types/context.types';

// Mock Office.js global
const mockOffice = {
  context: {
    mailbox: {
      item: null as any,
      addHandlerAsync: jest.fn()
    }
  },
  MailboxEnums: {
    ItemType: {
      Message: 'message',
      Appointment: 'appointment'
    }
  },
  EventType: {
    ItemChanged: 'itemChanged'
  },
  AsyncResultStatus: {
    Succeeded: 'succeeded',
    Failed: 'failed'
  }
};

// Setup global Office mock
(global as any).Office = mockOffice;

describe('contextService', () => {
  beforeEach(() => {
    // Reset mock before each test
    mockOffice.context.mailbox.item = null;
    jest.clearAllMocks();
  });

  describe('getContextType', () => {
    it('should return NoItem when no item is selected', () => {
      mockOffice.context.mailbox.item = null;

      expect(getContextType()).toBe(OutlookContextType.NoItem);
    });

    it('should return NoItem when mailbox is undefined', () => {
      const originalMailbox = mockOffice.context.mailbox;
      (mockOffice.context as any).mailbox = undefined;

      expect(getContextType()).toBe(OutlookContextType.NoItem);

      mockOffice.context.mailbox = originalMailbox;
    });

    it('should return MessageRead for read message item', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'message',
        from: {
          emailAddress: 'test@example.com',
          displayName: 'Test User'
        }
      };

      expect(getContextType()).toBe(OutlookContextType.MessageRead);
    });

    it('should return MessageCompose for compose message item', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'message',
        // Compose mode doesn't have .from as an object, it would be undefined
        // or have async methods
        from: undefined
      };

      expect(getContextType()).toBe(OutlookContextType.MessageCompose);
    });

    it('should return AppointmentAttendee for appointment read item', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'appointment',
        organizer: {
          emailAddress: 'organizer@example.com',
          displayName: 'Organizer'
        }
        // No getAsync method means it's read mode (attendee)
      };

      expect(getContextType()).toBe(OutlookContextType.AppointmentAttendee);
    });

    it('should return AppointmentOrganizer for appointment compose item', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'appointment',
        organizer: {
          getAsync: jest.fn() // Has async getter means compose mode (organizer)
        }
      };

      expect(getContextType()).toBe(OutlookContextType.AppointmentOrganizer);
    });

    it('should return Unknown for unknown item type', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'unknown'
      };

      expect(getContextType()).toBe(OutlookContextType.Unknown);
    });
  });

  describe('getFullContextState', () => {
    it('should return complete state for message read', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'message',
        itemId: 'msg-123',
        from: { emailAddress: 'test@example.com' }
      };

      const state = getFullContextState();

      expect(state.contextType).toBe(OutlookContextType.MessageRead);
      expect(state.itemId).toBe('msg-123');
      expect(state.isMessage).toBe(true);
      expect(state.isAppointment).toBe(false);
      expect(state.isReadMode).toBe(true);
      expect(state.isComposeMode).toBe(false);
      expect(state.supportsPin).toBe(true);
    });

    it('should return complete state for message compose', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'message',
        itemId: 'draft-456',
        from: undefined
      };

      const state = getFullContextState();

      expect(state.contextType).toBe(OutlookContextType.MessageCompose);
      expect(state.isMessage).toBe(true);
      expect(state.isAppointment).toBe(false);
      expect(state.isReadMode).toBe(false);
      expect(state.isComposeMode).toBe(true);
      expect(state.supportsPin).toBe(true);
    });

    it('should return complete state for appointment attendee', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'appointment',
        itemId: 'appt-789',
        organizer: { emailAddress: 'organizer@example.com' }
      };

      const state = getFullContextState();

      expect(state.contextType).toBe(OutlookContextType.AppointmentAttendee);
      expect(state.isMessage).toBe(false);
      expect(state.isAppointment).toBe(true);
      expect(state.isReadMode).toBe(true);
      expect(state.isComposeMode).toBe(false);
      expect(state.supportsPin).toBe(false); // Pinning not supported for appointments
    });

    it('should return complete state for appointment organizer', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'appointment',
        itemId: 'appt-new',
        organizer: { getAsync: jest.fn() }
      };

      const state = getFullContextState();

      expect(state.contextType).toBe(OutlookContextType.AppointmentOrganizer);
      expect(state.isMessage).toBe(false);
      expect(state.isAppointment).toBe(true);
      expect(state.isReadMode).toBe(false);
      expect(state.isComposeMode).toBe(true);
      expect(state.supportsPin).toBe(false); // Pinning not supported for appointments
    });

    it('should return null itemId when no item is selected', () => {
      mockOffice.context.mailbox.item = null;

      const state = getFullContextState();

      expect(state.itemId).toBeNull();
    });
  });

  describe('getCurrentItemId', () => {
    it('should return item ID when item is selected', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'message',
        itemId: 'item-123'
      };

      expect(getCurrentItemId()).toBe('item-123');
    });

    it('should return null when no item is selected', () => {
      mockOffice.context.mailbox.item = null;

      expect(getCurrentItemId()).toBeNull();
    });

    it('should return null when item has no itemId', () => {
      mockOffice.context.mailbox.item = {
        itemType: 'message'
        // No itemId property
      };

      expect(getCurrentItemId()).toBeNull();
    });
  });

  describe('subscribeToContextChanges', () => {
    it('should return unsubscribe function', () => {
      const callback = jest.fn();
      const unsubscribe = subscribeToContextChanges(callback);

      expect(typeof unsubscribe).toBe('function');
    });

    it('should allow unsubscribing', () => {
      const callback = jest.fn();
      const unsubscribe = subscribeToContextChanges(callback);

      // Unsubscribe should not throw
      expect(() => unsubscribe()).not.toThrow();
    });
  });
});
