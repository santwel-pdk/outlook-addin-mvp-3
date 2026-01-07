/**
 * Office.js Type Definitions
 *
 * Type extensions and definitions for Office.js APIs and app-specific types.
 *
 * @module office.types
 */

/**
 * Email data structure for displaying email information
 */
export interface EmailData {
  subject: string;
  from: string;
  receivedDate: Date;
  recipients: string[];
  bodyPreview?: string;
}

/**
 * Outlook item type (Message in Read or Compose mode)
 */
export type OutlookItem = Office.MessageRead | Office.MessageCompose;

/**
 * Office context wrapper with initialization state
 */
export interface OfficeContextState {
  context: Office.Context | null;
  mailbox: Office.Mailbox | null;
  isInitialized: boolean;
  error: Error | null;
}

/**
 * Mailbox item state for React components
 */
export interface MailboxItemState {
  item: Office.Item | null;
  itemType: Office.MailboxEnums.ItemType | null;
  isLoading: boolean;
  error: Error | null;
}

/**
 * WebView2 detection result
 */
export interface WebView2Status {
  isWebView2: boolean;
  browserEngine: string;
  userAgent: string;
}
