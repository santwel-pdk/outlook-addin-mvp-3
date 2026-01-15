/**
 * useOutlookContext Hook
 *
 * React hook to access current Outlook context (item type and mode).
 * Automatically updates when the selected item changes (via ItemChanged event).
 *
 * @module useOutlookContext
 */

import { useState, useEffect, useCallback } from 'react';
import { useOfficeContext } from './useOfficeContext';
import {
  getFullContextState,
  subscribeToContextChanges,
  registerItemChangedHandler
} from '../services/contextService';
import { OutlookContextState, OutlookContextType } from '../types/context.types';

/**
 * Default context state when Office is not initialized or no item is selected
 */
const defaultContextState: OutlookContextState = {
  contextType: OutlookContextType.NoItem,
  itemId: null,
  isMessage: false,
  isAppointment: false,
  isReadMode: false,
  isComposeMode: false,
  supportsPin: false
};

/**
 * Hook to access and track current Outlook context
 *
 * @returns {OutlookContextState} Current Outlook context state
 */
export function useOutlookContext(): OutlookContextState {
  const { isInitialized, error: officeError } = useOfficeContext();
  const [contextState, setContextState] = useState<OutlookContextState>(defaultContextState);

  // Handle context changes from ItemChanged events
  const handleContextChange = useCallback((newContext: OutlookContextState) => {
    setContextState(newContext);
  }, []);

  useEffect(() => {
    let isMounted = true;
    let unsubscribe: (() => void) | null = null;

    const initContext = async () => {
      // Wait for Office.js to be initialized
      if (!isInitialized) {
        return;
      }

      // If Office initialization failed, keep default state
      if (officeError) {
        return;
      }

      try {
        // Get initial context state
        const initialContext = getFullContextState();
        if (isMounted) {
          setContextState(initialContext);
        }

        // Register for ItemChanged events (for pinned task pane)
        try {
          await registerItemChangedHandler();
        } catch (error) {
          // ItemChanged registration may fail in some contexts, but that's OK
          console.warn('Could not register ItemChanged handler:', error);
        }

        // Subscribe to context changes
        unsubscribe = subscribeToContextChanges((newContext) => {
          if (isMounted) {
            handleContextChange(newContext);
          }
        });
      } catch (error) {
        console.error('Error initializing context:', error);
      }
    };

    initContext();

    return () => {
      isMounted = false;
      if (unsubscribe) {
        unsubscribe();
      }
    };
  }, [isInitialized, officeError, handleContextChange]);

  return contextState;
}

/**
 * Hook to get just the context type (lighter weight than full state)
 *
 * @returns {OutlookContextType} Current context type
 */
export function useContextType(): OutlookContextType {
  const { contextType } = useOutlookContext();
  return contextType;
}

/**
 * Hook to check if current context is a message (read or compose)
 *
 * @returns {boolean} True if current item is a message
 */
export function useIsMessage(): boolean {
  const { isMessage } = useOutlookContext();
  return isMessage;
}

/**
 * Hook to check if current context is an appointment (organizer or attendee)
 *
 * @returns {boolean} True if current item is an appointment
 */
export function useIsAppointment(): boolean {
  const { isAppointment } = useOutlookContext();
  return isAppointment;
}
