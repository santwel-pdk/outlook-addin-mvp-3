/**
 * useOfficeContext Hook
 *
 * React hook to access Office.js context in components with loading and error states.
 *
 * @module useOfficeContext
 */

import { useState, useEffect } from 'react';
import { initializeOffice, isOfficeInitialized } from '../services/officeService';
import { OfficeContextState } from '../types/office.types';

/**
 * Hook to initialize and access Office.js context
 *
 * @returns {OfficeContextState} Office context with loading and error states
 */
export function useOfficeContext(): OfficeContextState {
  const [state, setState] = useState<OfficeContextState>({
    context: null,
    mailbox: null,
    isInitialized: false,
    error: null
  });

  useEffect(() => {
    let isMounted = true;

    const initOffice = async () => {
      // Skip if already initialized
      if (isOfficeInitialized()) {
        if (isMounted) {
          setState({
            context: Office.context,
            mailbox: Office.context.mailbox,
            isInitialized: true,
            error: null
          });
        }
        return;
      }

      try {
        const context = await initializeOffice();

        if (isMounted) {
          setState({
            context,
            mailbox: context.mailbox,
            isInitialized: true,
            error: null
          });
        }
      } catch (error) {
        if (isMounted) {
          setState({
            context: null,
            mailbox: null,
            isInitialized: false,
            error: error as Error
          });
        }
      }
    };

    initOffice();

    return () => {
      isMounted = false;
    };
  }, []);

  return state;
}
