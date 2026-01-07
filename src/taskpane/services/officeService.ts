/**
 * Office.js Initialization Service
 *
 * Handles Office.js initialization with proper error handling and typing.
 *
 * @module officeService
 */

import { enforceWebView2 } from './webview2Service';

let isInitialized = false;
let officeContext: Office.Context | null = null;

/**
 * Initializes Office.js and verifies WebView2 environment
 *
 * @returns {Promise<Office.Context>} Promise resolving to Office context
 * @throws {Error} If Office.js fails to initialize or WebView2 check fails
 */
export async function initializeOffice(): Promise<Office.Context> {
  if (isInitialized && officeContext) {
    return officeContext;
  }

  return new Promise((resolve, reject) => {
    try {
      Office.onReady((info) => {
        try {
          // Enforce WebView2 before proceeding
          enforceWebView2();

          if (info.host === Office.HostType.Outlook) {
            officeContext = Office.context;
            isInitialized = true;
            console.log('Office.js initialized successfully for Outlook');
            resolve(Office.context);
          } else {
            reject(new Error(`Unsupported Office host: ${info.host}`));
          }
        } catch (error) {
          reject(error);
        }
      });
    } catch (error) {
      reject(new Error(`Office.js initialization failed: ${error.message}`));
    }
  });
}

/**
 * Checks if Office.js has been initialized
 *
 * @returns {boolean} True if Office is initialized
 */
export function isOfficeInitialized(): boolean {
  return isInitialized;
}

/**
 * Gets the Office context (throws if not initialized)
 *
 * @returns {Office.Context} Office context
 * @throws {Error} If Office is not initialized
 */
export function getOfficeContext(): Office.Context {
  if (!officeContext) {
    throw new Error('Office.js is not initialized. Call initializeOffice() first.');
  }
  return officeContext;
}

/**
 * Gets the current mailbox context
 *
 * @returns {Office.Mailbox} Mailbox context
 * @throws {Error} If Office is not initialized
 */
export function getMailboxContext(): Office.Mailbox {
  const context = getOfficeContext();
  if (!context.mailbox) {
    throw new Error('Mailbox context is not available');
  }
  return context.mailbox;
}
