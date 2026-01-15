/**
 * Context Detection Service
 *
 * Service for detecting the current Outlook context (item type and mode).
 * Handles ItemChanged events for pinned task pane support.
 *
 * @module contextService
 */

import {
  OutlookContextType,
  OutlookContextState,
  ContextChangeCallback
} from '../types/context.types';
import { logError } from '../utils/errorHandler';

// Subscribers for context change events
let contextChangeCallbacks: ContextChangeCallback[] = [];

// Track if ItemChanged handler is registered
let isHandlerRegistered = false;

/**
 * Determines the current Outlook context type from the mailbox item.
 * CRITICAL: Must handle null/undefined item.
 *
 * @returns {OutlookContextType} The detected context type
 */
export function getContextType(): OutlookContextType {
  try {
    // CRITICAL: Always check for null/undefined
    const mailbox = Office.context?.mailbox;
    if (!mailbox) {
      return OutlookContextType.NoItem;
    }

    const item = mailbox.item;
    if (!item) {
      return OutlookContextType.NoItem;
    }

    const itemType = item.itemType;

    // Check if Message
    if (itemType === Office.MailboxEnums.ItemType.Message) {
      // Detect Read vs Compose mode
      // Read mode: .from is an EmailAddressDetails object
      // Compose mode: .from is undefined or has getAsync method
      const messageItem = item as Office.MessageRead | Office.MessageCompose;

      // Check if it's a read item by verifying .from exists and is not a function
      if ('from' in messageItem && messageItem.from && typeof (messageItem as any).from !== 'function') {
        // MessageRead has .from as EmailAddressDetails
        return OutlookContextType.MessageRead;
      }

      // MessageCompose mode
      return OutlookContextType.MessageCompose;
    }

    // Check if Appointment
    if (itemType === Office.MailboxEnums.ItemType.Appointment) {
      // Detect Attendee (read) vs Organizer (compose) mode
      // AppointmentRead (attendee): .organizer is EmailAddressDetails
      // AppointmentCompose (organizer): .organizer is undefined, has .organizer.getAsync
      const appointmentItem = item as Office.AppointmentRead | Office.AppointmentCompose;

      // Check if it's a read item (attendee view)
      // In read mode, organizer is directly accessible as EmailAddressDetails
      if ('organizer' in appointmentItem && appointmentItem.organizer) {
        // Check if organizer is an object (AppointmentRead) vs has getAsync (AppointmentCompose)
        if (typeof (appointmentItem as any).organizer.getAsync !== 'function') {
          return OutlookContextType.AppointmentAttendee;
        }
      }

      // AppointmentCompose mode (organizer)
      return OutlookContextType.AppointmentOrganizer;
    }

    return OutlookContextType.Unknown;
  } catch (error) {
    logError('getContextType', error);
    return OutlookContextType.Unknown;
  }
}

/**
 * Gets the complete context state with all derived properties.
 *
 * @returns {OutlookContextState} Full context state object
 */
export function getFullContextState(): OutlookContextState {
  const contextType = getContextType();

  const isMessage =
    contextType === OutlookContextType.MessageRead ||
    contextType === OutlookContextType.MessageCompose;

  const isAppointment =
    contextType === OutlookContextType.AppointmentOrganizer ||
    contextType === OutlookContextType.AppointmentAttendee;

  const isReadMode =
    contextType === OutlookContextType.MessageRead ||
    contextType === OutlookContextType.AppointmentAttendee;

  const isComposeMode =
    contextType === OutlookContextType.MessageCompose ||
    contextType === OutlookContextType.AppointmentOrganizer;

  // Get item ID if available
  let itemId: string | null = null;
  try {
    const item = Office.context?.mailbox?.item;
    if (item && 'itemId' in item) {
      itemId = (item as any).itemId || null;
    }
  } catch {
    // Ignore errors getting item ID
  }

  return {
    contextType,
    itemId,
    isMessage,
    isAppointment,
    isReadMode,
    isComposeMode,
    // CRITICAL: Pinning only supported for messages, not appointments
    supportsPin: isMessage
  };
}

/**
 * Registers the ItemChanged event handler for pinned task pane support.
 * CRITICAL: Handler must always check for null item before accessing properties.
 *
 * @returns {Promise<void>}
 */
export async function registerItemChangedHandler(): Promise<void> {
  if (isHandlerRegistered) {
    console.log('ItemChanged handler already registered');
    return;
  }

  try {
    const mailbox = Office.context?.mailbox;
    if (!mailbox) {
      console.warn('Mailbox not available, cannot register ItemChanged handler');
      return;
    }

    await new Promise<void>((resolve, reject) => {
      mailbox.addHandlerAsync(
        Office.EventType.ItemChanged,
        onItemChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            isHandlerRegistered = true;
            console.log('ItemChanged handler registered successfully');
            resolve();
          } else {
            const errorMessage = result.error?.message || 'Unknown error';
            console.error('Failed to register ItemChanged handler:', errorMessage);
            reject(new Error(errorMessage));
          }
        }
      );
    });
  } catch (error) {
    logError('registerItemChangedHandler', error);
    throw error;
  }
}

/**
 * Removes the ItemChanged event handler.
 *
 * @returns {Promise<void>}
 */
export async function unregisterItemChangedHandler(): Promise<void> {
  if (!isHandlerRegistered) {
    return;
  }

  try {
    const mailbox = Office.context?.mailbox;
    if (!mailbox) {
      return;
    }

    await new Promise<void>((resolve, reject) => {
      mailbox.removeHandlerAsync(
        Office.EventType.ItemChanged,
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            isHandlerRegistered = false;
            console.log('ItemChanged handler unregistered');
            resolve();
          } else {
            reject(new Error(result.error?.message || 'Failed to unregister'));
          }
        }
      );
    });
  } catch (error) {
    logError('unregisterItemChangedHandler', error);
  }
}

/**
 * Internal handler for ItemChanged events.
 * Notifies all subscribers of the context change.
 *
 * @param {any} eventArgs Event arguments (not used but required for handler signature)
 */
function onItemChanged(_eventArgs: any): void {
  // CRITICAL: Always verify item is not null before proceeding
  const context = getFullContextState();

  console.log('ItemChanged event fired, new context:', context.contextType);

  // Notify all subscribers
  notifySubscribers(context);
}

/**
 * Subscribes to context change events.
 *
 * @param {ContextChangeCallback} callback Function to call when context changes
 * @returns {() => void} Unsubscribe function
 */
export function subscribeToContextChanges(callback: ContextChangeCallback): () => void {
  contextChangeCallbacks.push(callback);

  // Return unsubscribe function
  return () => {
    contextChangeCallbacks = contextChangeCallbacks.filter((cb) => cb !== callback);
  };
}

/**
 * Notifies all subscribers of a context change.
 *
 * @param {OutlookContextState} context The new context state
 */
function notifySubscribers(context: OutlookContextState): void {
  contextChangeCallbacks.forEach((callback) => {
    try {
      callback(context);
    } catch (error) {
      logError('contextChangeCallback', error);
    }
  });
}

/**
 * Gets the current item ID if available.
 *
 * @returns {string | null} The item ID or null if not available
 */
export function getCurrentItemId(): string | null {
  try {
    const item = Office.context?.mailbox?.item;
    if (item && 'itemId' in item) {
      return (item as any).itemId || null;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Checks if ItemChanged handler is currently registered.
 *
 * @returns {boolean} True if handler is registered
 */
export function isItemChangedHandlerRegistered(): boolean {
  return isHandlerRegistered;
}
