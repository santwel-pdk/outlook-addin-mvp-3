/**
 * SSOStatus Component
 *
 * Displays SSO authentication status, user information, and error handling.
 * Uses Fluent UI components for consistent Office look-and-feel.
 *
 * @module SSOStatus
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
  Button,
  makeStyles,
  tokens
} from '@fluentui/react-components';
import { 
  PersonAccounts24Regular,
  Person24Regular,
  ShieldKeyhole24Regular,
  Warning24Regular,
  Key24Regular,
  ArrowClockwise24Regular,
  SignOut24Regular
} from '@fluentui/react-icons';
import { useSSO } from '../hooks/useSSO';
import { formatSSOErrorMessage, getSSOErrorGuidance } from '../utils/errorHandler';

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
  userInfoRow: {
    display: 'flex',
    alignItems: 'flex-start',
    gap: tokens.spacingHorizontalS,
    marginBottom: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: tokens.borderRadiusMedium
  },
  errorContainer: {
    marginTop: tokens.spacingVerticalS,
    padding: tokens.spacingVerticalS,
    backgroundColor: tokens.colorStatusDangerBackground1,
    borderRadius: tokens.borderRadiusMedium,
    border: `1px solid ${tokens.colorStatusDangerBorder1}`
  },
  authenticatedIcon: {
    color: tokens.colorStatusSuccessBackground2
  },
  notAuthenticatedIcon: {
    color: tokens.colorStatusDangerBackground2
  },
  authenticatingIcon: {
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
  },
  actionButtons: {
    display: 'flex',
    gap: tokens.spacingHorizontalS,
    marginTop: tokens.spacingVerticalS,
    flexWrap: 'wrap'
  }
});

/**
 * Gets the appropriate icon for authentication state
 */
const getAuthIcon = (status: string, isLoading: boolean, styles: any) => {
  if (isLoading) {
    return <Spinner size="extra-small" />;
  }

  switch (status) {
    case 'authenticated':
      return <PersonAccounts24Regular className={styles.authenticatedIcon} />;
    case 'authenticating':
      return <ShieldKeyhole24Regular className={styles.authenticatingIcon} />;
    case 'error':
      return <Warning24Regular className={styles.errorIcon} />;
    default:
      return <Person24Regular className={styles.notAuthenticatedIcon} />;
  }
};

/**
 * Gets the appropriate badge appearance for authentication state
 */
const getBadgeAppearance = (status: string): "filled" | "outline" | "tint" => {
  switch (status) {
    case 'authenticated':
      return 'filled';
    case 'authenticating':
      return 'tint';
    default:
      return 'outline';
  }
};

/**
 * Gets the badge color for authentication state
 */
const getBadgeColor = (status: string) => {
  switch (status) {
    case 'authenticated':
      return 'success';
    case 'authenticating':
      return 'warning';
    case 'error':
      return 'danger';
    default:
      return 'subtle';
  }
};

/**
 * Gets user-friendly status text
 */
const getStatusText = (status: string, isLoading: boolean): string => {
  if (isLoading) return 'Authenticating...';

  switch (status) {
    case 'authenticated':
      return 'Authenticated';
    case 'authenticating':
      return 'Authenticating...';
    case 'error':
      return 'Error';
    default:
      return 'Not Authenticated';
  }
};

/**
 * SSOStatus component displaying authentication status and user info
 *
 * @returns {JSX.Element} SSO status component
 */
const SSOStatus: React.FC = () => {
  const styles = useStyles();
  const { 
    ssoState, 
    isLoading, 
    status, 
    user, 
    token, 
    error, 
    hasError,
    initialize, 
    refresh, 
    signOut,
    isTokenValid
  } = useSSO();

  // Handle initialization
  const handleInitialize = async () => {
    await initialize({
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      forMSGraphAccess: false
    });
  };

  // Handle token refresh
  const handleRefresh = async () => {
    await refresh();
  };

  // Show loading state during authentication
  if (isLoading && !ssoState.isInitialized) {
    return (
      <Card className={styles.card}>
        <div className={styles.loadingContainer}>
          <Spinner size="small" label="Initializing authentication..." />
        </div>
      </Card>
    );
  }

  return (
    <Card className={styles.card}>
      <CardHeader 
        header={
          <div className={styles.badgeContainer}>
            <Text weight="semibold" size={400}>Authentication Status</Text>
            <Badge 
              appearance={getBadgeAppearance(status.status)}
              color={getBadgeColor(status.status)}
              size="small"
            >
              {getStatusText(status.status, isLoading)}
            </Badge>
          </div>
        } 
      />

      <div style={{ padding: '12px' }}>
        <div className={styles.statusRow}>
          {getAuthIcon(status.status, isLoading, styles)}
          <div>
            <Text size={200} weight="semibold">Status:</Text>
            <br />
            <Text size={300}>
              {status.message}
              {isLoading && (
                <Spinner size="tiny" style={{ marginLeft: '8px' }} />
              )}
            </Text>
          </div>
        </div>

        {/* User Information */}
        {user && status.status === 'authenticated' && (
          <div className={styles.userInfoRow}>
            <PersonAccounts24Regular />
            <div style={{ flex: 1 }}>
              <Text size={200} weight="semibold">Signed in as:</Text>
              <br />
              <Text size={300} weight="semibold">Name:</Text> <Text size={300}>{user.displayName}</Text>
              <br />
              <Text size={300} weight="semibold">Email:</Text> <Text size={300}>{user.email}</Text>
              {token && (
                <>
                  <br />
                  <Text size={300} weight="semibold">Token expires:</Text>{' '}
                  <Text size={300}>
                    {new Date(token.expiresAt).toLocaleString()}
                    {!isTokenValid && (
                      <Text size={200} style={{ color: tokens.colorStatusDangerBackground2, marginLeft: '4px' }}>
                        (Expired)
                      </Text>
                    )}
                  </Text>
                </>
              )}
            </div>
          </div>
        )}

        {/* Error Display */}
        {hasError && error && (
          <div className={styles.errorContainer}>
            <div className={styles.statusRow}>
              <Warning24Regular className={styles.errorIcon} />
              <div>
                <Text size={300} weight="semibold" style={{ color: tokens.colorStatusDangerForeground1 }}>
                  {formatSSOErrorMessage(error)}
                </Text>
                <br />
                <Text size={200} style={{ color: tokens.colorStatusDangerForeground2 }}>
                  {getSSOErrorGuidance(error.code)}
                </Text>
              </div>
            </div>
          </div>
        )}

        {/* Action Buttons */}
        <div className={styles.actionButtons}>
          {!ssoState.isAuthenticated && (
            <Button
              appearance="primary"
              icon={<Key24Regular />}
              onClick={handleInitialize}
              disabled={isLoading}
            >
              {isLoading ? 'Authenticating...' : 'Sign In'}
            </Button>
          )}

          {ssoState.isAuthenticated && (
            <>
              <Button
                appearance="secondary"
                icon={<ArrowClockwise24Regular />}
                onClick={handleRefresh}
                disabled={isLoading}
              >
                {isLoading ? 'Refreshing...' : 'Refresh Token'}
              </Button>
              
              <Button
                appearance="outline"
                icon={<SignOut24Regular />}
                onClick={signOut}
                disabled={isLoading}
              >
                Sign Out
              </Button>
            </>
          )}
        </div>

        {/* Status Messages */}
        {ssoState.isAuthenticated && !hasError && (
          <MessageBar intent="success" style={{ marginTop: tokens.spacingVerticalS }}>
            <MessageBarBody>
              Authentication successful. You can now use real-time features.
            </MessageBarBody>
          </MessageBar>
        )}

        {!ssoState.isAuthenticated && !hasError && ssoState.isInitialized && (
          <MessageBar intent="info" style={{ marginTop: tokens.spacingVerticalS }}>
            <MessageBarBody>
              Sign in with your Microsoft 365 account to enable real-time features and personalized experience.
            </MessageBarBody>
          </MessageBar>
        )}
      </div>
    </Card>
  );
};

export default SSOStatus;