/* THROWAWAY prototype. Inline wordmark; colour is only black or white (brandbook p. 7). */
import * as React from "react";
import { WORDMARK_PATHS } from "./wordmarkPaths";

export const Wordmark = ({ color, width }: { color: "#231f20" | "#ffffff"; width: number }) => (
  <svg
    role="img"
    aria-label="DIMSCALE"
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

/* Add-in icon preview: white tile with the "D" glyph (first wordmark path) in a green. */
export const DIcon = ({ size, green }: { size: number; green: string }) => (
  <svg
    role="img"
    aria-label="Ícone D"
    width={size}
    height={size}
    viewBox="-9.75 -6.6 56 56"
    style={{ display: "block" }}
  >
    <rect x="-9.75" y="-6.6" width="56" height="56" fill="#ffffff" />
    <path transform={WORDMARK_PATHS[0].t} d={WORDMARK_PATHS[0].d} fill={green} />
  </svg>
);
