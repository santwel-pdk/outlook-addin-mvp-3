/**
 * Unit Tests for appointmentService
 *
 * Tests appointment data retrieval for both Read and Compose modes
 */

import {
  getAppointmentSubject,
  getAppointmentOrganizer,
  getAppointmentRequiredAttendees,
  getAppointmentOptionalAttendees,
  getAppointmentStart,
  getAppointmentEnd,
  getAppointmentLocation
} from '../appointmentService';

// Mock the contextService
jest.mock('../contextService', () => ({
  getContextType: jest.fn()
}));

// Mock the officeService
jest.mock('../officeService', () => ({
  getMailboxContext: jest.fn()
}));

import { getContextType } from '../contextService';
import { getMailboxContext } from '../officeService';
import { OutlookContextType } from '../../types/context.types';

const mockGetContextType = getContextType as jest.Mock;
const mockGetMailboxContext = getMailboxContext as jest.Mock;

// Mock Office.js global
const mockOffice = {
  MailboxEnums: {
    ItemType: {
      Message: 'message',
      Appointment: 'appointment'
    }
  },
  AsyncResultStatus: {
    Succeeded: 'succeeded',
    Failed: 'failed'
  }
};

(global as any).Office = mockOffice;

describe('appointmentService', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  describe('getAppointmentSubject', () => {
    it('should return subject for read appointment (attendee)', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          subject: 'Team Meeting'
        }
      });

      const result = await getAppointmentSubject();
      expect(result).toBe('Team Meeting');
    });

    it('should return "(No Subject)" for appointment without subject', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          subject: ''
        }
      });

      const result = await getAppointmentSubject();
      expect(result).toBe('(No Subject)');
    });

    it('should throw error when no item is selected', async () => {
      mockGetMailboxContext.mockReturnValue({
        item: null
      });

      await expect(getAppointmentSubject()).rejects.toThrow('No appointment is currently selected');
    });

    it('should throw error when item is not an appointment', async () => {
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'message'
        }
      });

      await expect(getAppointmentSubject()).rejects.toThrow('Current item is not an appointment');
    });

    it('should get subject via async for compose appointment (organizer)', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentOrganizer);

      const mockGetAsync = jest.fn((callback) => {
        callback({
          status: 'succeeded',
          value: 'New Meeting'
        });
      });

      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          subject: { getAsync: mockGetAsync }
        }
      });

      const result = await getAppointmentSubject();
      expect(result).toBe('New Meeting');
      expect(mockGetAsync).toHaveBeenCalled();
    });
  });

  describe('getAppointmentOrganizer', () => {
    it('should return organizer for read appointment', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          organizer: {
            emailAddress: 'organizer@example.com',
            displayName: 'John Organizer'
          }
        }
      });

      const result = await getAppointmentOrganizer();
      expect(result).toBe('organizer@example.com');
    });

    it('should return displayName when emailAddress is not available', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          organizer: {
            displayName: 'John Organizer'
          }
        }
      });

      const result = await getAppointmentOrganizer();
      expect(result).toBe('John Organizer');
    });

    it('should return "Unknown Organizer" when no organizer info', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          organizer: null
        }
      });

      const result = await getAppointmentOrganizer();
      expect(result).toBe('Unknown Organizer');
    });
  });

  describe('getAppointmentRequiredAttendees', () => {
    it('should return required attendees for read appointment', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          requiredAttendees: [
            { emailAddress: 'user1@example.com', displayName: 'User 1' },
            { emailAddress: 'user2@example.com', displayName: 'User 2' }
          ]
        }
      });

      const result = await getAppointmentRequiredAttendees();
      expect(result).toEqual(['user1@example.com', 'user2@example.com']);
    });

    it('should return empty array when no required attendees', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          requiredAttendees: []
        }
      });

      const result = await getAppointmentRequiredAttendees();
      expect(result).toEqual([]);
    });

    it('should get attendees via async for compose appointment', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentOrganizer);

      const mockGetAsync = jest.fn((callback) => {
        callback({
          status: 'succeeded',
          value: [
            { emailAddress: 'attendee@example.com' }
          ]
        });
      });

      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          requiredAttendees: { getAsync: mockGetAsync }
        }
      });

      const result = await getAppointmentRequiredAttendees();
      expect(result).toEqual(['attendee@example.com']);
    });
  });

  describe('getAppointmentOptionalAttendees', () => {
    it('should return optional attendees for read appointment', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          optionalAttendees: [
            { emailAddress: 'optional@example.com', displayName: 'Optional User' }
          ]
        }
      });

      const result = await getAppointmentOptionalAttendees();
      expect(result).toEqual(['optional@example.com']);
    });

    it('should return empty array when no optional attendees', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          optionalAttendees: null
        }
      });

      const result = await getAppointmentOptionalAttendees();
      expect(result).toEqual([]);
    });
  });

  describe('getAppointmentStart', () => {
    it('should return start time for read appointment', async () => {
      const startDate = new Date('2024-01-15T10:00:00');
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          start: startDate
        }
      });

      const result = await getAppointmentStart();
      expect(result).toEqual(startDate);
    });

    it('should get start time via async for compose appointment', async () => {
      const startDate = new Date('2024-01-15T14:00:00');
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentOrganizer);

      const mockGetAsync = jest.fn((callback) => {
        callback({
          status: 'succeeded',
          value: startDate
        });
      });

      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          start: { getAsync: mockGetAsync }
        }
      });

      const result = await getAppointmentStart();
      expect(result).toEqual(startDate);
    });
  });

  describe('getAppointmentEnd', () => {
    it('should return end time for read appointment', async () => {
      const endDate = new Date('2024-01-15T11:00:00');
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          end: endDate
        }
      });

      const result = await getAppointmentEnd();
      expect(result).toEqual(endDate);
    });

    it('should get end time via async for compose appointment', async () => {
      const endDate = new Date('2024-01-15T15:00:00');
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentOrganizer);

      const mockGetAsync = jest.fn((callback) => {
        callback({
          status: 'succeeded',
          value: endDate
        });
      });

      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          end: { getAsync: mockGetAsync }
        }
      });

      const result = await getAppointmentEnd();
      expect(result).toEqual(endDate);
    });
  });

  describe('getAppointmentLocation', () => {
    it('should return location for read appointment', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          location: 'Conference Room A'
        }
      });

      const result = await getAppointmentLocation();
      expect(result).toBe('Conference Room A');
    });

    it('should return empty string when no location', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentAttendee);
      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          location: ''
        }
      });

      const result = await getAppointmentLocation();
      expect(result).toBe('');
    });

    it('should get location via async for compose appointment', async () => {
      mockGetContextType.mockReturnValue(OutlookContextType.AppointmentOrganizer);

      const mockGetAsync = jest.fn((callback) => {
        callback({
          status: 'succeeded',
          value: 'Virtual Meeting'
        });
      });

      mockGetMailboxContext.mockReturnValue({
        item: {
          itemType: 'appointment',
          location: { getAsync: mockGetAsync }
        }
      });

      const result = await getAppointmentLocation();
      expect(result).toBe('Virtual Meeting');
    });
  });
});
