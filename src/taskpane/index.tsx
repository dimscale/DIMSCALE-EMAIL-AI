import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./App";

/* global document, Office, HTMLElement, location */

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  // resumo.html is the same bundle: the path alone tells the "Resumo" ribbon button apart (F-012 AC-50).
  root?.render(<App summaryEntry={/\/resumo\.html$/.test(location.pathname)} />);
});
