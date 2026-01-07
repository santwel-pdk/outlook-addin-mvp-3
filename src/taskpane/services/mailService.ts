/**
 * Mail Service for Email Operations
 *
 * Service layer for accessing current email/mailbox data using Office.js APIs.
 * All functions use async/await pattern with proper error handling.
 *
 * @module mailService
 */

import { getMailboxContext } from './officeService';
import { handleOfficeError } from '../utils/errorHandler';

/**
 * Gets the current email subject
 * https://learn.microsoft.com/en-us/javascript/api/outlook/office.messageread#outlook-office-messageread-subject-member
 *
 * @returns {Promise<string>} Promise resolving to email subject
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getCurrentEmailSubject(): Promise<string> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No email item is currently selected');
    }

    // For read mode, subject is directly available
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      return item.subject || '(No Subject)';
    }

    return '(No Subject)';
  } catch (error) {
    throw new Error(handleOfficeError('getCurrentEmailSubject', error));
  }
}

/**
 * Gets the current email sender address
 *
 * @returns {Promise<string>} Promise resolving to sender email address
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getCurrentEmailFrom(): Promise<string> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item as Office.MessageRead;

    if (!item) {
      throw new Error('No email item is currently selected');
    }

    if (item.itemType === Office.MailboxEnums.ItemType.Message && item.from) {
      return item.from.emailAddress || item.from.displayName || 'Unknown Sender';
    }

    return 'Unknown Sender';
  } catch (error) {
    throw new Error(handleOfficeError('getCurrentEmailFrom', error));
  }
}

/**
 * Gets the email received date
 *
 * @returns {Promise<Date>} Promise resolving to received date
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getEmailReceivedDate(): Promise<Date> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item as Office.MessageRead;

    if (!item) {
      throw new Error('No email item is currently selected');
    }

    if (item.dateTimeCreated) {
      return item.dateTimeCreated;
    }

    return new Date();
  } catch (error) {
    throw new Error(handleOfficeError('getEmailReceivedDate', error));
  }
}

/**
 * Gets the current email's recipient addresses (To field)
 *
 * @returns {Promise<string[]>} Promise resolving to array of recipient emails
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getCurrentEmailRecipients(): Promise<string[]> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item as Office.MessageRead;

    if (!item) {
      throw new Error('No email item is currently selected');
    }

    if (item.to && Array.isArray(item.to)) {
      return item.to.map((recipient) => recipient.emailAddress || recipient.displayName);
    }

    return [];
  } catch (error) {
    throw new Error(handleOfficeError('getCurrentEmailRecipients', error));
  }
}

/**
 * Gets the current email body preview
 *
 * @param {number} maxLength Maximum characters to return (default: 200)
 * @returns {Promise<string>} Promise resolving to email body preview
 * @throws {Error} If mailbox context is not available or API call fails
 */
export async function getEmailBodyPreview(maxLength: number = 200): Promise<string> {
  try {
    const mailbox = getMailboxContext();
    const item = mailbox.item;

    if (!item) {
      throw new Error('No email item is currently selected');
    }

    return new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          const preview = result.value.substring(0, maxLength);
          resolve(preview + (result.value.length > maxLength ? '...' : ''));
        } else {
          reject(new Error(result.error.message));
        }
      });
    });
  } catch (error) {
    throw new Error(handleOfficeError('getEmailBodyPreview', error));
  }
}
