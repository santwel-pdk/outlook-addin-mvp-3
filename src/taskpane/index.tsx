import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import ErrorBoundary from "./components/ErrorBoundary";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { enforceWebView2 } from "./services/webview2Service";

/* global document, Office, module, require, HTMLElement */

const title = "outlook-addin-mvp-3";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  try {
    // Enforce WebView2 before rendering
    enforceWebView2();

    root?.render(
      <FluentProvider theme={webLightTheme}>
        <ErrorBoundary>
          <App title={title} />
        </ErrorBoundary>
      </FluentProvider>
    );
  } catch (error) {
    console.error('Failed to initialize add-in:', error);
    root?.render(
      <FluentProvider theme={webLightTheme}>
        <div style={{ padding: '20px', color: 'red' }}>
          <h2>Initialization Error</h2>
          <p>{(error as Error).message}</p>
        </div>
      </FluentProvider>
    );
  }
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}
