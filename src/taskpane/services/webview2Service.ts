/**
 * WebView2 Detection and Enforcement Service
 *
 * Verifies that the add-in is running in WebView2 (Edge Chromium).
 * Throws error if running in legacy IE11/Trident.
 *
 * @module webview2Service
 */

/**
 * Verifies that the add-in is running in WebView2 (Edge Chromium).
 * Throws error if running in legacy IE11/Trident.
 *
 * @throws {Error} If running in IE11/Trident browser
 * @returns {boolean} True if WebView2 is detected
 */
export function enforceWebView2(): boolean {
  const userAgent = window.navigator.userAgent;

  // Check if running in IE11/Trident (NOT allowed)
  if (userAgent.indexOf('Trident') !== -1) {
    throw new Error(
      'This add-in requires WebView2 (Edge Chromium). ' +
      'Please update your Office installation to use WebView2.'
    );
  }

  // Verify Edge Chromium
  const isEdgeChromium = userAgent.indexOf('Edg/') !== -1;
  const isChrome = userAgent.includes('Chrome');

  if (!isEdgeChromium && !isChrome) {
    console.warn('WebView2 not detected. Some features may not work.');
    return false;
  }

  console.log('WebView2 (Edge Chromium) detected:', userAgent);
  return true;
}

/**
 * Gets the current browser engine name
 *
 * @returns {string} Browser engine name
 */
export function getBrowserEngine(): string {
  const userAgent = window.navigator.userAgent;

  if (userAgent.indexOf('Trident') !== -1) {
    return 'IE11/Trident';
  }

  if (userAgent.indexOf('Edg/') !== -1) {
    return 'WebView2 (Edge Chromium)';
  }

  if (userAgent.includes('Chrome')) {
    return 'Chrome/Chromium';
  }

  if (userAgent.includes('Safari')) {
    return 'Safari/WebKit';
  }

  return 'Unknown';
}

/**
 * Checks if the environment is running WebView2
 *
 * @returns {boolean} True if WebView2 is active
 */
export function isWebView2(): boolean {
  const userAgent = window.navigator.userAgent;
  return userAgent.indexOf('Edg/') !== -1;
}
