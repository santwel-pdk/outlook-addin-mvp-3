/**
 * SignalRStatus Component
 *
 * Displays SignalR connection status and last received message.
 * Uses Fluent UI components for consistent Office look-and-feel.
 *
 * @module SignalRStatus
 */

import * as React from 'react';
import {
  Card,
  CardHeader,
  Text,
  Spinner,
  MessageBar,
  MessageBarBody,
  Badge,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import { 
  Connected24Regular, 
  WifiOff24Regular, 
  ArrowClockwise24Regular,
  Warning24Regular,
  ChatBubblesQuestion24Regular 
} from '@fluentui/react-icons';
import { HubConnectionState } from '@microsoft/signalr';
import { SignalRConnectionState } from '../types/signalr.types';

interface SignalRStatusProps {
  signalrState: SignalRConnectionState;
}

const useStyles = makeStyles({
  card: {
    margin: tokens.spacingVerticalM,
    maxWidth: '100%'
  },
  statusRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS
  },
  messageRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium
  },
  connectedIcon: {
    color: tokens.colorStatusSuccessBackground2
  },
  disconnectedIcon: {
    color: tokens.colorStatusDangerBackground2
  },
  reconnectingIcon: {
    color: tokens.colorStatusWarningBackground2
  },
  errorIcon: {
    color: tokens.colorStatusDangerBackground2
  },
  loadingContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    padding: tokens.spacingVerticalXXL
  },
  badgeContainer: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalXS
  }
});

/**
 * Gets the appropriate icon and style for connection state
 */
const getConnectionIcon = (state: HubConnectionState, styles: any) => {
  switch (state) {
    case HubConnectionState.Connected:
      return <Connected24Regular className={styles.connectedIcon} />;
    case HubConnectionState.Disconnected:
      return <WifiOff24Regular className={styles.disconnectedIcon} />;
    case HubConnectionState.Connecting:
    case HubConnectionState.Reconnecting:
      return <ArrowClockwise24Regular className={styles.reconnectingIcon} />;
    default:
      return <Warning24Regular className={styles.errorIcon} />;
  }
};

/**
 * Gets the appropriate badge appearance for connection state
 */
const getBadgeAppearance = (state: HubConnectionState): "filled" | "outline" | "tint" => {
  switch (state) {
    case HubConnectionState.Connected:
      return 'filled';
    case HubConnectionState.Connecting:
    case HubConnectionState.Reconnecting:
      return 'tint';
    default:
      return 'outline';
  }
};

/**
 * Gets the badge color for connection state
 */
const getBadgeColor = (state: HubConnectionState) => {
  switch (state) {
    case HubConnectionState.Connected:
      return 'success';
    case HubConnectionState.Connecting:
    case HubConnectionState.Reconnecting:
      return 'warning';
    default:
      return 'danger';
  }
};

/**
 * Gets user-friendly connection status text
 */
const getStatusText = (state: HubConnectionState): string => {
  switch (state) {
    case HubConnectionState.Connected:
      return 'Connected';
    case HubConnectionState.Connecting:
      return 'Connecting...';
    case HubConnectionState.Reconnecting:
      return 'Reconnecting...';
    case HubConnectionState.Disconnected:
      return 'Disconnected';
    default:
      return 'Unknown';
  }
};

/**
 * Safely formats a timestamp, handling edge cases like invalid dates
 */
const formatTimestamp = (timestamp: Date | string | number | undefined): string => {
  if (!timestamp) {
    return 'Unknown';
  }

  const date = timestamp instanceof Date ? timestamp : new Date(timestamp);

  // Check for invalid date
  if (isNaN(date.getTime())) {
    return 'Invalid time';
  }

  return date.toLocaleTimeString();
};

/**
 * SignalRStatus component displaying connection status and messages
 *
 * @param {SignalRStatusProps} props Component props
 * @returns {JSX.Element} SignalR status component
 */
const SignalRStatus: React.FC<SignalRStatusProps> = ({ signalrState }) => {
  const styles = useStyles();
  const { connectionState, isConnected, error, lastMessage } = signalrState;

  // Show error state
  if (error) {
    return (
      <MessageBar intent="error">
        <MessageBarBody>
          SignalR Error: {error}
        </MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <Card className={styles.card}>
      <CardHeader 
        header={
          <div className={styles.badgeContainer}>
            <Text weight="semibold" size={400}>Real-time Connection</Text>
            <Badge 
              appearance={getBadgeAppearance(connectionState)}
              color={getBadgeColor(connectionState)}
              size="small"
            >
              {getStatusText(connectionState)}
            </Badge>
          </div>
        } 
      />

      <div style={{ padding: '12px' }}>
        <div className={styles.statusRow}>
          {getConnectionIcon(connectionState, styles)}
          <div>
            <Text size={200} weight="semibold">Status:</Text>
            <br />
            <Text size={300}>
              {getStatusText(connectionState)}
              {connectionState === HubConnectionState.Connecting || 
               connectionState === HubConnectionState.Reconnecting ? (
                <Spinner size="tiny" style={{ marginLeft: '8px' }} />
              ) : null}
            </Text>
          </div>
        </div>

        {lastMessage && (
          <div className={styles.messageRow}>
            <ChatBubblesQuestion24Regular />
            <div style={{ flex: 1 }}>
              <Text size={200} weight="semibold">Latest Message:</Text>
              <br />
              <Text size={300} weight="semibold">Type:</Text> <Text size={300}>{lastMessage.type}</Text>
              <br />
              <Text size={300} weight="semibold">Time:</Text> <Text size={300}>{formatTimestamp(lastMessage.timestamp)}</Text>
              {lastMessage.payload && (
                <>
                  <br />
                  <Text size={300} weight="semibold">Content:</Text>
                  <br />
                  <Text size={200} style={{ fontFamily: 'monospace' }}>
                    {typeof lastMessage.payload === 'string' 
                      ? lastMessage.payload 
                      : JSON.stringify(lastMessage.payload, null, 2)
                    }
                  </Text>
                </>
              )}
            </div>
          </div>
        )}

        {isConnected && !lastMessage && (
          <MessageBar intent="info">
            <MessageBarBody>
              Connection established. Waiting for real-time messages...
            </MessageBarBody>
          </MessageBar>
        )}
      </div>
    </Card>
  );
};

export default SignalRStatus;