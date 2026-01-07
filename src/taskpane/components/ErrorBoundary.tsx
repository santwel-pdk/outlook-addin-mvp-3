/**
 * ErrorBoundary Component
 *
 * React error boundary for graceful error handling and recovery.
 * Catches JavaScript errors anywhere in the child component tree.
 *
 * @module ErrorBoundary
 */

import * as React from 'react';
import { MessageBar, MessageBarBody, Button, makeStyles, tokens } from '@fluentui/react-components';
import { ErrorBoundaryProps, ErrorBoundaryState } from '../types/app.types';
import { logError } from '../utils/errorHandler';

const useStyles = makeStyles({
  container: {
    padding: tokens.spacingVerticalL,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM
  },
  errorDetails: {
    marginTop: tokens.spacingVerticalS,
    padding: tokens.spacingHorizontalM,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    fontFamily: 'monospace',
    fontSize: '12px',
    maxHeight: '200px',
    overflow: 'auto'
  }
});

/**
 * Error Boundary component for catching React errors
 */
class ErrorBoundary extends React.Component<ErrorBoundaryProps, ErrorBoundaryState> {
  constructor(props: ErrorBoundaryProps) {
    super(props);
    this.state = {
      hasError: false,
      error: null
    };
  }

  static getDerivedStateFromError(error: Error): ErrorBoundaryState {
    return {
      hasError: true,
      error
    };
  }

  componentDidCatch(error: Error, errorInfo: React.ErrorInfo): void {
    logError('ErrorBoundary', {
      message: error.message,
      stack: error.stack,
      componentStack: errorInfo.componentStack
    });
  }

  handleReset = (): void => {
    this.setState({
      hasError: false,
      error: null
    });
  };

  render(): React.ReactNode {
    if (this.state.hasError) {
      if (this.props.fallback) {
        return this.props.fallback;
      }

      return <ErrorFallback error={this.state.error} onReset={this.handleReset} />;
    }

    return this.props.children;
  }
}

/**
 * Default error fallback UI
 */
interface ErrorFallbackProps {
  error: Error | null;
  onReset: () => void;
}

const ErrorFallback: React.FC<ErrorFallbackProps> = ({ error, onReset }) => {
  const styles = useStyles();
  const isDevelopment = process.env.NODE_ENV !== 'production';

  return (
    <div className={styles.container}>
      <MessageBar intent="error">
        <MessageBarBody>
          <strong>Something went wrong</strong>
          <p>The add-in encountered an unexpected error. Please try reloading.</p>
        </MessageBarBody>
      </MessageBar>

      <Button appearance="primary" onClick={onReset}>
        Try Again
      </Button>

      {isDevelopment && error && (
        <div className={styles.errorDetails}>
          <strong>Error Details (Development Only):</strong>
          <pre>{error.message}</pre>
          {error.stack && (
            <>
              <strong>Stack Trace:</strong>
              <pre>{error.stack}</pre>
            </>
          )}
        </div>
      )}
    </div>
  );
};

export default ErrorBoundary;
