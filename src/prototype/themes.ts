/* THROWAWAY prototype (F-012 step 1). Brand themes per the spec token mapping. */
import {
  BrandVariants,
  Theme,
  createDarkTheme,
  createLightTheme,
} from "@fluentui/react-components";

const KEYS = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150, 160] as const;

const rgb = (h: string) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
const hex = (c: number[]) =>
  "#" + c.map((v) => Math.round(v).toString(16).padStart(2, "0")).join("");

/* Fills the steps missing from `anchors` by linear RGB interpolation between the nearest anchors. */
function ramp(anchors: Record<number, string>): BrandVariants {
  const known = Object.keys(anchors).map(Number);
  const out: Record<number, string> = {};
  for (const k of KEYS) {
    if (anchors[k]) {
      out[k] = anchors[k];
      continue;
    }
    const lo = Math.max(...known.filter((x) => x < k));
    const hi = Math.min(...known.filter((x) => x > k));
    const t = (k - lo) / (hi - lo);
    out[k] = hex(rgb(anchors[lo]).map((v, i) => v + (rgb(anchors[hi])[i] - v) * t));
  }
  return out as unknown as BrandVariants;
}

export const C = {
  navy: "#102d3f",
  navyDeep: "#081b24",
  ink: "#231f20",
  paper: "#fbfcfc",
  mist: "#f1f4f5",
  line: "#d9e0e3",
  teal: "#7fd1c7",
  muted: "#617484",
  secondaryLight: "#405969",
  secondaryDark: "#9fb2bf",
  orangeDark: "#C2410C", // accessible primary on light (white label 5.18:1)
  orangeBright: "#F26B21", // mockup-like orange, only as accent on dark surfaces
};

const lightNavy = ramp({
  10: "#020608",
  40: "#050f15",
  60: "#081b24",
  70: "#0c2433",
  80: "#102d3f",
  90: "#173e54",
  100: "#245064",
  160: C.mist,
});
const darkTeal = ramp({
  10: "#031214",
  20: "#0d3a3f",
  40: "#5fb9ae",
  70: C.teal,
  80: "#96dad1",
  100: "#a6e0d9",
  110: "#bbe8e2",
  160: "#f0fbf9",
});
const lightOrange = ramp({
  10: "#1a0803",
  40: "#6b2406",
  60: "#8f3008",
  70: "#a8380a",
  80: C.orangeDark,
  90: "#d9531a",
  100: "#e8672e",
  160: "#fff4ee",
});
const darkOrange = ramp({
  10: "#1a0803",
  20: "#3a1608",
  40: "#d85a14",
  70: C.orangeBright,
  80: "#f58245",
  100: "#fa9d6d",
  110: "#fcb18a",
  160: "#fff4ee",
});

const font = '"Helvetica Neue", Helvetica, Arial, sans-serif'; // brandbook: Helvetica for online/e-mail

const light = (brand: BrandVariants): Theme => ({
  ...createLightTheme(brand),
  fontFamilyBase: font,
  colorNeutralBackground1: "#ffffff",
  colorNeutralBackground2: C.mist,
  colorNeutralForeground1: C.ink,
  colorNeutralForeground2: C.secondaryLight,
  colorNeutralForeground3: C.secondaryLight,
  colorNeutralStroke1: C.muted,
  colorNeutralStrokeAccessible: C.muted,
  colorNeutralStroke2: C.line,
});

const dark = (brand: BrandVariants): Theme => ({
  ...createDarkTheme(brand),
  fontFamilyBase: font,
  colorNeutralBackground1: C.navyDeep,
  colorNeutralBackground2: C.navy,
  colorNeutralForeground1: C.paper,
  colorNeutralForeground2: C.secondaryDark,
  colorNeutralForeground3: C.secondaryDark,
  colorNeutralForegroundOnBrand: C.navyDeep, // white on teal/orange fails
  colorNeutralStroke1: "#7a8f9d",
  colorNeutralStrokeAccessible: "#7a8f9d",
  colorNeutralStroke2: "#173e54",
});

export type Accent = "navy" | "orange";
export const themes = (accent: Accent) => ({
  light: light(accent === "navy" ? lightNavy : lightOrange),
  dark: dark(accent === "navy" ? darkTeal : darkOrange),
});
