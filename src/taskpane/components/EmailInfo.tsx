/**
 * EmailInfo Component
 *
 * Displays current email information including subject, from, and received date.
 * Uses Fluent UI components for consistent Office look-and-feel.
 *
 * @module EmailInfo
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
  tokens
} from '@fluentui/react-components';
import { Mail24Regular, Person24Regular, Calendar24Regular } from '@fluentui/react-icons';
import { EmailInfoProps } from '../types/app.types';

const useStyles = makeStyles({
  card: {
    margin: tokens.spacingVerticalM,
    maxWidth: '100%'
  },
  infoRow: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS
  },
  icon: {
    color: tokens.colorBrandForeground1
  },
  loadingContainer: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    padding: tokens.spacingVerticalXXL
  }
});

/**
 * EmailInfo component displaying email data
 *
 * @param {EmailInfoProps} props Component props
 * @returns {JSX.Element} Email info component
 */
const EmailInfo: React.FC<EmailInfoProps> = ({ emailData, isLoading, error }) => {
  const styles = useStyles();

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner label="Loading email data..." />
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

  if (!emailData) {
    return (
      <MessageBar intent="warning">
        <MessageBarBody>No email data available. Please select an email.</MessageBarBody>
      </MessageBar>
    );
  }

  return (
    <Card className={styles.card}>
      <CardHeader header={<Text weight="semibold" size={400}>Email Information</Text>} />

      <div style={{ padding: '12px' }}>
        <div className={styles.infoRow}>
          <Mail24Regular className={styles.icon} />
          <div>
            <Text size={200} weight="semibold">Subject:</Text>
            <br />
            <Text size={300}>{emailData.subject}</Text>
          </div>
        </div>

        <div className={styles.infoRow}>
          <Person24Regular className={styles.icon} />
          <div>
            <Text size={200} weight="semibold">From:</Text>
            <br />
            <Text size={300}>{emailData.from}</Text>
          </div>
        </div>

        <div className={styles.infoRow}>
          <Calendar24Regular className={styles.icon} />
          <div>
            <Text size={200} weight="semibold">Received:</Text>
            <br />
            <Text size={300}>{emailData.receivedDate.toLocaleString()}</Text>
          </div>
        </div>
      </div>
    </Card>
  );
};

export default EmailInfo;
