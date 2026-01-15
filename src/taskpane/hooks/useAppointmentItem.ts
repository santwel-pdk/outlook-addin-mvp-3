/**
 * useAppointmentItem Hook
 *
 * React hook to access current appointment/calendar item data with loading and error states.
 *
 * @module useAppointmentItem
 */

import { useState, useEffect } from 'react';
import { useOfficeContext } from './useOfficeContext';
import { useOutlookContext } from './useOutlookContext';
import {
  getAppointmentSubject,
  getAppointmentOrganizer,
  getAppointmentRequiredAttendees,
  getAppointmentOptionalAttendees,
  getAppointmentStart,
  getAppointmentEnd,
  getAppointmentLocation
} from '../services/appointmentService';
import { AppointmentData } from '../types/office.types';
import { handleOfficeError } from '../utils/errorHandler';

/**
 * State interface for appointment item data
 */
interface AppointmentItemState {
  appointmentData: AppointmentData | null;
  isLoading: boolean;
  error: string | null;
}

/**
 * Hook to get current appointment item data
 *
 * @returns {AppointmentItemState} Appointment data with loading and error states
 */
export function useAppointmentItem(): AppointmentItemState {
  const { isInitialized, error: officeError } = useOfficeContext();
  const { isAppointment, itemId } = useOutlookContext();
  const [state, setState] = useState<AppointmentItemState>({
    appointmentData: null,
    isLoading: true,
    error: null
  });

  useEffect(() => {
    let isMounted = true;

    const fetchAppointmentData = async () => {
      // Wait for Office to be initialized
      if (!isInitialized) {
        return;
      }

      // If Office initialization failed, set error state
      if (officeError) {
        if (isMounted) {
          setState({
            appointmentData: null,
            isLoading: false,
            error: officeError.message
          });
        }
        return;
      }

      // Only fetch if current context is an appointment
      if (!isAppointment) {
        if (isMounted) {
          setState({
            appointmentData: null,
            isLoading: false,
            error: null
          });
        }
        return;
      }

      try {
        // Fetch all appointment data in parallel
        const [
          subject,
          organizer,
          requiredAttendees,
          optionalAttendees,
          start,
          end,
          location
        ] = await Promise.all([
          getAppointmentSubject(),
          getAppointmentOrganizer(),
          getAppointmentRequiredAttendees(),
          getAppointmentOptionalAttendees(),
          getAppointmentStart(),
          getAppointmentEnd(),
          getAppointmentLocation()
        ]);

        if (isMounted) {
          setState({
            appointmentData: {
              subject,
              organizer,
              requiredAttendees,
              optionalAttendees,
              start,
              end,
              location: location || undefined
            },
            isLoading: false,
            error: null
          });
        }
      } catch (error) {
        if (isMounted) {
          setState({
            appointmentData: null,
            isLoading: false,
            error: handleOfficeError('useAppointmentItem', error)
          });
        }
      }
    };

    // Reset loading state when item changes
    setState((prev) => ({ ...prev, isLoading: true }));
    fetchAppointmentData();

    return () => {
      isMounted = false;
    };
  }, [isInitialized, officeError, isAppointment, itemId]);

  return state;
}
