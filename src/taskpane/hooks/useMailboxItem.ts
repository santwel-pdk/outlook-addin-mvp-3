/**
 * useMailboxItem Hook
 *
 * React hook to access current mailbox item (email) data with loading and error states.
 *
 * @module useMailboxItem
 */

import { useState, useEffect } from 'react';
import { useOfficeContext } from './useOfficeContext';
import {
  getCurrentEmailSubject,
  getCurrentEmailFrom,
  getEmailReceivedDate,
  getCurrentEmailRecipients
} from '../services/mailService';
import { EmailData } from '../types/office.types';
import { handleOfficeError } from '../utils/errorHandler';

interface MailboxItemState {
  emailData: EmailData | null;
  isLoading: boolean;
  error: string | null;
}

/**
 * Hook to get current email item data
 *
 * @returns {MailboxItemState} Email data with loading and error states
 */
export function useMailboxItem(): MailboxItemState {
  const { isInitialized, error: officeError } = useOfficeContext();
  const [state, setState] = useState<MailboxItemState>({
    emailData: null,
    isLoading: true,
    error: null
  });

  useEffect(() => {
    let isMounted = true;

    const fetchEmailData = async () => {
      if (!isInitialized) {
        return;
      }

      if (officeError) {
        if (isMounted) {
          setState({
            emailData: null,
            isLoading: false,
            error: officeError.message
          });
        }
        return;
      }

      try {
        const [subject, from, receivedDate, recipients] = await Promise.all([
          getCurrentEmailSubject(),
          getCurrentEmailFrom(),
          getEmailReceivedDate(),
          getCurrentEmailRecipients()
        ]);

        if (isMounted) {
          setState({
            emailData: {
              subject,
              from,
              receivedDate,
              recipients
            },
            isLoading: false,
            error: null
          });
        }
      } catch (error) {
        if (isMounted) {
          setState({
            emailData: null,
            isLoading: false,
            error: handleOfficeError('useMailboxItem', error)
          });
        }
      }
    };

    fetchEmailData();

    return () => {
      isMounted = false;
    };
  }, [isInitialized, officeError]);

  return state;
}
