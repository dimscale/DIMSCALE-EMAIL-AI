import * as React from "react";
import { strings } from "./strings";
import { WORDMARK_PATHS } from "./wordmarkPaths";

/* Brandbook p. 7: the logo is only black or white, never coloured. */
export const Wordmark = ({ color, width = 108 }: { color: "#000000" | "#ffffff" | "currentColor"; width?: number }) => (
  <svg
    role="img"
    aria-label={strings.wordmarkName}
    xmlns="http://www.w3.org/2000/svg"
    width={width}
    height={(width * 42.831117) / 509.4936}
    viewBox="0 0 509.4936 42.831117"
    style={{ display: "block", flex: "none" }}
  >
    {WORDMARK_PATHS.map((p, i) => (
      <path key={i} transform={p.t} d={p.d} fill={color} />
    ))}
  </svg>
);
