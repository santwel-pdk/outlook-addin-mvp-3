/**
 * Application Type Definitions
 *
 * App-specific TypeScript interfaces and types.
 *
 * @module app.types
 */

import { EmailData } from './office.types';

/**
 * Application state
 */
export interface AppState {
  isOfficeInitialized: boolean;
  isLoading: boolean;
  error: string | null;
  platform: string;
}

/**
 * Email Info Component Props
 */
export interface EmailInfoProps {
  emailData: EmailData | null;
  isLoading: boolean;
  error: string | null;
}

/**
 * Header Component Props
 */
export interface HeaderProps {
  title: string;
  platform?: string;
  showPlatform?: boolean;
}

/**
 * Error Boundary Props
 */
export interface ErrorBoundaryProps {
  children: React.ReactNode;
  fallback?: React.ReactNode;
}

/**
 * Error Boundary State
 */
export interface ErrorBoundaryState {
  hasError: boolean;
  error: Error | null;
}
