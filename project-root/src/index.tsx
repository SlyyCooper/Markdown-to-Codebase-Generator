import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./taskpane/components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import { setTaskpaneDimensions } from "./taskpane/taskpane";

/* global document, Office */
const title = "";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

Office.onReady(() => {
  setTaskpaneDimensions();
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});
