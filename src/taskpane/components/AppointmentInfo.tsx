/**
 * AppointmentInfo Component
 *
 * Displays current appointment/calendar item information including subject,
 * organizer, attendees, and time. Uses Fluent UI components for consistent
 * Office look-and-feel.
 *
 * @module AppointmentInfo
 */

import * as React from 'react';
import {
  Card,
  CardHeader,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  makeStyles,
  tokens,
  Badge
} from '@fluentui/react-components';
import {
  CalendarLtr24Regular,
  Person24Regular,
  People24Regular,
  Clock24Regular,
  Location24Regular
} from '@fluentui/react-icons';
import { AppointmentData } from '../types/office.types';

/**
 * Props for AppointmentInfo component
 */
interface AppointmentInfoProps {
  appointmentData: AppointmentData | null;
  isLoading: boolean;
  error: string | null;
  isOrganizer?: boolean;
}

const useStyles = makeStyles({
  card: {
    margin: tokens.spacingVerticalM,
    maxWidth: '100%'
  },
  infoRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS
  },
  icon: {
    color: tokens.colorBrandForeground1,
    marginTop: '2px'
  },
  loadingContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    padding: tokens.spacingVerticalXXL
  },
  attendeeList: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS
  },
  attendeeBadge: {
    marginRight: tokens.spacingHorizontalXS,
    marginBottom: tokens.spacingVerticalXS
  },
  headerBadge: {
    marginLeft: tokens.spacingHorizontalS
  },
  timeRange: {
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalXS
  }
});

/**
 * Formats a date/time for display
 */
function formatDateTime(date: Date): string {
  return date.toLocaleString(undefined, {
    weekday: 'short',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  });
}

/**
 * Formats duration between two dates
 */
function formatDuration(start: Date, end: Date): string {
  const durationMs = end.getTime() - start.getTime();
  const durationMins = Math.round(durationMs / (1000 * 60));

  if (durationMins < 60) {
    return `${durationMins} min`;
  }

  const hours = Math.floor(durationMins / 60);
  const mins = durationMins % 60;

  if (mins === 0) {
    return `${hours} hr`;
  }

  return `${hours} hr ${mins} min`;
}

/**
 * AppointmentInfo component displaying appointment data
 *
 * @param {AppointmentInfoProps} props Component props
 * @returns {JSX.Element} Appointment info component
 */
const AppointmentInfo: React.FC<AppointmentInfoProps> = ({
  appointmentData,
  isLoading,
  error,
  isOrganizer = false
}) => {
  const styles = useStyles();

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label="Loading appointment data..." />
      </div>
    );
  }

  if (error) {
    return (
      <MessageBar intent="error">
        <MessageBarBody>{error}</MessageBarBody>
      </MessageBar>
    );
  }

  if (!appointmentData) {
    return (
      <MessageBar intent="warning">
        <MessageBarBody>No appointment data available. Please select an appointment.</MessageBarBody>
      </MessageBar>
    );
  }

  const allAttendees = [
    ...appointmentData.requiredAttendees,
    ...appointmentData.optionalAttendees
  ];

  return (
    <Card className={styles.card}>
      <CardHeader
        header={
          <div style={{ display: 'flex', alignItems: 'center' }}>
            <Text weight="semibold" size={400}>
              Appointment Information
            </Text>
            <Badge
              className={styles.headerBadge}
              appearance="filled"
              color={isOrganizer ? 'brand' : 'informative'}
            >
              {isOrganizer ? 'Organizer' : 'Attendee'}
            </Badge>
          </div>
        }
      />

      <div style={{ padding: '12px' }}>
        {/* Subject */}
        <div className={styles.infoRow}>
          <CalendarLtr24Regular className={styles.icon} />
          <div>
            <Text size={200} weight="semibold">
              Subject:
            </Text>
            <br />
            <Text size={300}>{appointmentData.subject}</Text>
          </div>
        </div>

        {/* Organizer */}
        <div className={styles.infoRow}>
          <Person24Regular className={styles.icon} />
          <div>
            <Text size={200} weight="semibold">
              Organizer:
            </Text>
            <br />
            <Text size={300}>{appointmentData.organizer}</Text>
          </div>
        </div>

        {/* Time */}
        <div className={styles.infoRow}>
          <Clock24Regular className={styles.icon} />
          <div className={styles.timeRange}>
            <Text size={200} weight="semibold">
              Time:
            </Text>
            <Text size={300}>
              {formatDateTime(appointmentData.start)}
            </Text>
            <Text size={300}>
              to {formatDateTime(appointmentData.end)}
            </Text>
            <Text size={200} style={{ color: tokens.colorNeutralForeground3 }}>
              ({formatDuration(appointmentData.start, appointmentData.end)})
            </Text>
          </div>
        </div>

        {/* Location (if available) */}
        {appointmentData.location && (
          <div className={styles.infoRow}>
            <Location24Regular className={styles.icon} />
            <div>
              <Text size={200} weight="semibold">
                Location:
              </Text>
              <br />
              <Text size={300}>{appointmentData.location}</Text>
            </div>
          </div>
        )}

        {/* Attendees */}
        {allAttendees.length > 0 && (
          <div className={styles.infoRow}>
            <People24Regular className={styles.icon} />
            <div>
              <Text size={200} weight="semibold">
                Attendees ({allAttendees.length}):
              </Text>
              <div className={styles.attendeeList} style={{ marginTop: '4px' }}>
                {appointmentData.requiredAttendees.map((attendee, index) => (
                  <Badge
                    key={`req-${index}`}
                    className={styles.attendeeBadge}
                    appearance="outline"
                    color="important"
                  >
                    {attendee}
                  </Badge>
                ))}
                {appointmentData.optionalAttendees.map((attendee, index) => (
                  <Badge
                    key={`opt-${index}`}
                    className={styles.attendeeBadge}
                    appearance="outline"
                    color="subtle"
                  >
                    {attendee} (optional)
                  </Badge>
                ))}
              </div>
            </div>
          </div>
        )}
      </div>
    </Card>
  );
};

export default AppointmentInfo;
