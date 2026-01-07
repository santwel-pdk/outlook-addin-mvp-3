/**
 * Unit Tests for webview2Service
 *
 * Tests WebView2 detection and enforcement logic
 */

import { enforceWebView2, getBrowserEngine, isWebView2 } from '../webview2Service';

describe('webview2Service', () => {
  let originalUserAgent: string;

  beforeEach(() => {
    originalUserAgent = navigator.userAgent;
  });

  afterEach(() => {
    // Restore original user agent
    Object.defineProperty(navigator, 'userAgent', {
      value: originalUserAgent,
      writable: true
    });
  });

  describe('enforceWebView2', () => {
    it('should throw error when running in IE11/Trident', () => {
      // Mock IE11 user agent
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Windows NT 10.0; Trident/7.0; rv:11.0) like Gecko',
        writable: true
      });

      expect(() => enforceWebView2()).toThrow('This add-in requires WebView2');
    });

    it('should return true when WebView2 is detected', () => {
      // Mock Edge Chromium user agent
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36 Edg/120.0.0.0',
        writable: true
      });

      expect(enforceWebView2()).toBe(true);
    });

    it('should warn and return false when browser is not WebView2 or Chrome', () => {
      const consoleWarnSpy = jest.spyOn(console, 'warn').mockImplementation();

      // Mock Safari user agent
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15',
        writable: true
      });

      const result = enforceWebView2();

      expect(result).toBe(false);
      expect(consoleWarnSpy).toHaveBeenCalledWith(
        expect.stringContaining('WebView2 not detected')
      );

      consoleWarnSpy.mockRestore();
    });
  });

  describe('getBrowserEngine', () => {
    it('should return "IE11/Trident" for IE11', () => {
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Windows NT 10.0; Trident/7.0; rv:11.0) like Gecko',
        writable: true
      });

      expect(getBrowserEngine()).toBe('IE11/Trident');
    });

    it('should return "WebView2 (Edge Chromium)" for Edge', () => {
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 Chrome/120.0.0.0 Edg/120.0.0.0',
        writable: true
      });

      expect(getBrowserEngine()).toBe('WebView2 (Edge Chromium)');
    });

    it('should return "Safari/WebKit" for Safari', () => {
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 Safari/605.1.15',
        writable: true
      });

      expect(getBrowserEngine()).toBe('Safari/WebKit');
    });
  });

  describe('isWebView2', () => {
    it('should return true for WebView2 user agent', () => {
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Windows NT 10.0) AppleWebKit/537.36 Edg/120.0.0.0',
        writable: true
      });

      expect(isWebView2()).toBe(true);
    });

    it('should return false for non-WebView2 user agent', () => {
      Object.defineProperty(navigator, 'userAgent', {
        value: 'Mozilla/5.0 (Macintosh) Safari/605.1.15',
        writable: true
      });

      expect(isWebView2()).toBe(false);
    });
  });
});
