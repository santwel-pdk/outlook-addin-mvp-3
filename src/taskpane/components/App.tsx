import * as React from "react";
import EmailInfo from "./EmailInfo";
import AppointmentInfo from "./AppointmentInfo";
import SignalRStatus from "./SignalRStatus";
import { makeStyles, MessageBar, MessageBarBody, tokens } from "@fluentui/react-components";
import { useMailboxItem } from "../hooks/useMailboxItem";
import { useAppointmentItem } from "../hooks/useAppointmentItem";
import { useOutlookContext } from "../hooks/useOutlookContext";
import { useSignalR } from "../hooks/useSignalR";
import { OutlookContextType } from "../types/context.types";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
  },
  noItemContainer: {
    padding: tokens.spacingVerticalL,
    textAlign: "center",
  },
});

/**
 * Main App component with context-aware rendering.
 * Displays EmailInfo for messages, AppointmentInfo for calendar items.
 */
const App: React.FC<AppProps> = (_props: AppProps) => {
  const styles = useStyles();
  const signalrState = useSignalR();

  // Get current Outlook context
  const { contextType, isMessage, isAppointment } = useOutlookContext();

  // Get email data (only fetches when in message context)
  const mailboxState = useMailboxItem();

  // Get appointment data (only fetches when in appointment context)
  const appointmentState = useAppointmentItem();

  /**
   * Renders the appropriate content based on current Outlook context
   */
  const renderContent = () => {
    switch (contextType) {
      case OutlookContextType.MessageRead:
      case OutlookContextType.MessageCompose:
        return (
          <EmailInfo
            emailData={mailboxState.emailData}
            isLoading={mailboxState.isLoading}
            error={mailboxState.error}
          />
        );

      case OutlookContextType.AppointmentOrganizer:
        return (
          <AppointmentInfo
            appointmentData={appointmentState.appointmentData}
            isLoading={appointmentState.isLoading}
            error={appointmentState.error}
            isOrganizer={true}
          />
        );

      case OutlookContextType.AppointmentAttendee:
        return (
          <AppointmentInfo
            appointmentData={appointmentState.appointmentData}
            isLoading={appointmentState.isLoading}
            error={appointmentState.error}
            isOrganizer={false}
          />
        );

      case OutlookContextType.NoItem:
        return (
          <div className={styles.noItemContainer}>
            <MessageBar intent="info">
              <MessageBarBody>
                Select an email or calendar item to view details.
              </MessageBarBody>
            </MessageBar>
          </div>
        );

      case OutlookContextType.Unknown:
      default:
        return (
          <div className={styles.noItemContainer}>
            <MessageBar intent="warning">
              <MessageBarBody>
                Unable to determine the current context. Please select an item.
              </MessageBarBody>
            </MessageBar>
          </div>
        );
    }
  };

  return (
    <div className={styles.root}>
      {renderContent()}
      <SignalRStatus signalrState={signalrState} />
    </div>
  );
};

export default App;
