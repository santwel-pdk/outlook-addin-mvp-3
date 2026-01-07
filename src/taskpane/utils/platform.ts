/**
 * Platform Detection Utility
 *
 * Helper functions to detect Windows vs macOS for platform-specific code.
 *
 * @module platform
 */

/**
 * Checks if the current platform is macOS
 *
 * @returns {boolean} True if running on macOS
 */
export function isMacOS(): boolean {
  return navigator.platform.toUpperCase().indexOf('MAC') >= 0;
}

/**
 * Checks if the current platform is Windows
 *
 * @returns {boolean} True if running on Windows
 */
export function isWindows(): boolean {
  return navigator.platform.toUpperCase().indexOf('WIN') >= 0;
}

/**
 * Gets the platform name as a string
 *
 * @returns {string} Platform name ('macOS', 'Windows', or 'Unknown')
 */
export function getPlatformName(): string {
  if (isMacOS()) {
    return 'macOS';
  }

  if (isWindows()) {
    return 'Windows';
  }

  return 'Unknown';
}

/**
 * Checks if the current platform is a desktop platform
 *
 * @returns {boolean} True if running on desktop (Windows or macOS)
 */
export function isDesktop(): boolean {
  return isWindows() || isMacOS();
}
