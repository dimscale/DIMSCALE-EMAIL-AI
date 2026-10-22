import * as React from "react";
import { createRoot } from "react-dom/client";
import MessagePanel from "./MessagePanel";
import BackendStatus from "./BackendStatus";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

/* global document, Office, HTMLElement */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <MessagePanel />
      <BackendStatus />
    </FluentProvider>
  );
});
