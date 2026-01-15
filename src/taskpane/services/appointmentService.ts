/**
 * Appointment Service for Calendar Operations
 *
 * Service layer for accessing current appointment/calendar data using Office.js APIs.
 * All functions use async/await pattern with proper error handling.
 * Handles both Read (attendee) and Compose (organizer) modes.
 *
 * @module appointmentService
 */

import { getMailboxContext } from './officeService';
import { handleOfficeError } from '../utils/errorHandler';
import { getContextType } from './contextService';
import { OutlookContextType } from '../types/context.types';

/**
 * Gets the current appointment subject.
 * Handles both Read (direct access) and Compose (async getter) modes.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-subject-member
 *
 * @returns {Promise<string>} Promise resolving to appointment subject
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentSubject(): Promise<string> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - direct access
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      return apptRead.subject || '(No Subject)';
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || '(No Subject)');
        } else {
          reject(new Error(result.error?.message || 'Failed to get subject'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentSubject', error));
  }
}

/**
 * Gets the appointment organizer.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-organizer-member
 *
 * @returns {Promise<string>} Promise resolving to organizer email or display name
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentOrganizer(): Promise<string> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - organizer is directly available
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      if (apptRead.organizer) {
        return apptRead.organizer.emailAddress || apptRead.organizer.displayName || 'Unknown Organizer';
      }
      return 'Unknown Organizer';
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.organizer.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const organizer = result.value;
          resolve(organizer?.emailAddress || organizer?.displayName || 'Unknown Organizer');
        } else {
          reject(new Error(result.error?.message || 'Failed to get organizer'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentOrganizer', error));
  }
}

/**
 * Gets the appointment required attendees.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-requiredattendees-member
 *
 * @returns {Promise<string[]>} Promise resolving to array of required attendee emails
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentRequiredAttendees(): Promise<string[]> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - direct access
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      if (apptRead.requiredAttendees && Array.isArray(apptRead.requiredAttendees)) {
        return apptRead.requiredAttendees.map(
          (attendee) => attendee.emailAddress || attendee.displayName || 'Unknown'
        );
      }
      return [];
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.requiredAttendees.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const attendees = result.value || [];
          resolve(
            attendees.map(
              (attendee: Office.EmailAddressDetails) =>
                attendee.emailAddress || attendee.displayName || 'Unknown'
            )
          );
        } else {
          reject(new Error(result.error?.message || 'Failed to get required attendees'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentRequiredAttendees', error));
  }
}

/**
 * Gets the appointment optional attendees.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-optionalattendees-member
 *
 * @returns {Promise<string[]>} Promise resolving to array of optional attendee emails
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentOptionalAttendees(): Promise<string[]> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - direct access
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      if (apptRead.optionalAttendees && Array.isArray(apptRead.optionalAttendees)) {
        return apptRead.optionalAttendees.map(
          (attendee) => attendee.emailAddress || attendee.displayName || 'Unknown'
        );
      }
      return [];
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.optionalAttendees.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const attendees = result.value || [];
          resolve(
            attendees.map(
              (attendee: Office.EmailAddressDetails) =>
                attendee.emailAddress || attendee.displayName || 'Unknown'
            )
          );
        } else {
          reject(new Error(result.error?.message || 'Failed to get optional attendees'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentOptionalAttendees', error));
  }
}

/**
 * Gets the appointment start time.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-start-member
 *
 * @returns {Promise<Date>} Promise resolving to start date/time
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentStart(): Promise<Date> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - direct access
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      return apptRead.start || new Date();
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.start.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || new Date());
        } else {
          reject(new Error(result.error?.message || 'Failed to get start time'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentStart', error));
  }
}

/**
 * Gets the appointment end time.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-end-member
 *
 * @returns {Promise<Date>} Promise resolving to end date/time
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentEnd(): Promise<Date> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - direct access
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      return apptRead.end || new Date();
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.end.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || new Date());
        } else {
          reject(new Error(result.error?.message || 'Failed to get end time'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentEnd', error));
  }
}

/**
 * Gets the appointment location.
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member
 *
 * @returns {Promise<string>} Promise resolving to appointment location
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getAppointmentLocation(): Promise<string> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No appointment is currently selected');
    }

    if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
      throw new Error('Current item is not an appointment');
    }

    const contextType = getContextType();

    // AppointmentRead (attendee mode) - direct access
    if (contextType === OutlookContextType.AppointmentAttendee) {
      const apptRead = item as Office.AppointmentRead;
      return apptRead.location || '';
    }

    // AppointmentCompose (organizer mode) - async getter
    const apptCompose = item as Office.AppointmentCompose;
    return new Promise((resolve, reject) => {
      apptCompose.location.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || '');
        } else {
          reject(new Error(result.error?.message || 'Failed to get location'));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getAppointmentLocation', error));
  }
}
