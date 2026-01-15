/**
 * Context Type Definitions
 *
 * Type definitions for Outlook context detection and state management.
 * Used to determine the current item type and mode (read/compose).
 *
 * @module context.types
 */

/**
 * Outlook item context enumeration
 * Represents the different contexts in which the add-in can operate
 */
export enum OutlookContextType {
  /** Reading an email message */
  MessageRead = 'MessageRead',
  /** Composing an email message */
  MessageCompose = 'MessageCompose',
  /** Organizing/creating an appointment (meeting organizer) */
  AppointmentOrganizer = 'AppointmentOrganizer',
  /** Attending/viewing an appointment (meeting attendee) */
  AppointmentAttendee = 'AppointmentAttendee',
  /** No item selected */
  NoItem = 'NoItem',
  /** Unknown or unsupported context */
  Unknown = 'Unknown'
}

/**
 * Current context state
 * Provides comprehensive information about the current Outlook context
 */
export interface OutlookContextState {
  /** The detected context type */
  contextType: OutlookContextType;
  /** The ID of the currently selected item, or null if no item */
  itemId: string | null;
  /** True if the current item is a message (email) */
  isMessage: boolean;
  /** True if the current item is an appointment (calendar item) */
  isAppointment: boolean;
  /** True if in read mode (viewing existing item) */
  isReadMode: boolean;
  /** True if in compose mode (creating/editing item) */
  isComposeMode: boolean;
  /** True if pinning is supported in this context (messages only) */
  supportsPin: boolean;
}

/**
 * Callback type for context change subscribers
 */
export type ContextChangeCallback = (context: OutlookContextState) => void;

/**
 * Context change event arguments
 */
export interface ContextChangeEventArgs {
  /** The previous context state */
  previousContext: OutlookContextState | null;
  /** The new context state */
  currentContext: OutlookContextState;
  /** Timestamp of the change */
  timestamp: Date;
}
