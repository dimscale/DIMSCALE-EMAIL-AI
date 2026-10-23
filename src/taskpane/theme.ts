import {
  Theme,
  teamsHighContrastTheme,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import * as React from "react";

/* global Office, window */

export type ThemeKind = "light" | "dark" | "highContrast";
export interface OfficeThemeColors {
  bodyBackgroundColor?: string;
  bodyForegroundColor?: string;
}

const HEX = /^#[0-9a-f]{6}$/i;

/* WCAG relative luminance of "#rrggbb"; null when the value is not a hex triplet. */
export function hexLuminance(hex: string | undefined): number | null {
  if (!hex || !HEX.test(hex)) return null;
  const [r, g, b] = [1, 3, 5].map((i) => {
    const v = parseInt(hex.slice(i, i + 2), 16) / 255;
    return v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4);
  });
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}

/* Dark below 0.179: the point where white and black text have equal contrast (spec F-012). */
export function resolveTheme(o: {
  office?: OfficeThemeColors; // only passed when Mailbox 1.14 is supported
  prefersDark?: boolean;
  forcedColors?: boolean;
}): { kind: ThemeKind; theme: Theme } {
  if (o.forcedColors) return { kind: "highContrast", theme: teamsHighContrastTheme };
  const lum = hexLuminance(o.office?.bodyBackgroundColor);
  const dark = lum === null ? !!o.prefersDark : lum < 0.179;
  const base = dark ? webDarkTheme : webLightTheme;
  if (lum === null) return { kind: dark ? "dark" : "light", theme: base };
  const fg = HEX.test(o.office?.bodyForegroundColor ?? "")
    ? o.office?.bodyForegroundColor
    : undefined;
  return {
    kind: dark ? "dark" : "light",
    theme: {
      ...base,
      colorNeutralBackground1: o.office?.bodyBackgroundColor as string,
      ...(fg ? { colorNeutralForeground1: fg } : {}),
    },
  };
}

function read(office?: OfficeThemeColors) {
  const mq = (q: string) =>
    typeof window !== "undefined" && window.matchMedia ? window.matchMedia(q).matches : false;
  let supported = false;
  try {
    supported = !!Office.context?.requirements?.isSetSupported?.("Mailbox", "1.14");
  } catch {
    supported = false;
  }
  return resolveTheme({
    office: supported ? (office ?? Office.context?.officeTheme) : undefined,
    prefersDark: mq("(prefers-color-scheme: dark)"),
    forcedColors: mq("(forced-colors: active)"),
  });
}

/* Follows the Office theme (OfficeThemeChanged, Mailbox 1.14) with prefers-color-scheme as fallback. */
export function useOfficeTheme() {
  const [t, setT] = React.useState(() => read());
  React.useEffect(() => {
    const update = () => setT(read());
    const mqs =
      typeof window !== "undefined" && window.matchMedia
        ? ["(prefers-color-scheme: dark)", "(forced-colors: active)"].map((q) =>
            window.matchMedia(q)
          )
        : [];
    mqs.forEach((m) => m.addEventListener?.("change", update));
    const mailbox = Office.context?.mailbox;
    const canListen =
      !!mailbox?.addHandlerAsync &&
      !!Office.context?.requirements?.isSetSupported?.("Mailbox", "1.14");
    // The event args carry officeTheme, but the typings declare the generic AsyncResult callback.
    const handler = ((args: { officeTheme?: OfficeThemeColors }) =>
      setT(read(args?.officeTheme))) as unknown as (r: Office.AsyncResult<void>) => void;
    if (canListen) mailbox.addHandlerAsync(Office.EventType.OfficeThemeChanged, handler);
    return () => {
      mqs.forEach((m) => m.removeEventListener?.("change", update));
      if (canListen) mailbox.removeHandlerAsync?.(Office.EventType.OfficeThemeChanged, handler);
    };
  }, []);
  return t;
}
