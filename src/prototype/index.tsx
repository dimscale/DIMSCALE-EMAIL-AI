/* THROWAWAY prototype (F-012 step 1). Dev-server only (webpack entry exists in development mode). */
import * as React from "react";
import { createRoot } from "react-dom/client";
import {
  FluentProvider,
  Radio,
  RadioGroup,
  Text,
  webDarkTheme,
  webLightTheme,
} from "@fluentui/react-components";
import { Direction, Panel } from "./Panel";
import { Accent, themes } from "./themes";
import { DIcon } from "./Wordmark";

/* global document */

const VARIANTS: { label: string; dir: Direction; accent: Accent }[] = [
  { label: "A-navy (referência, primário navy)", dir: "A", accent: "navy" },
  { label: "A-orange (laranja acessível #C2410C)", dir: "A", accent: "orange" },
  { label: "B “Navy header”", dir: "B", accent: "navy" },
  { label: "C “Cards”", dir: "C", accent: "navy" },
];

const GREENS = ["#2BD94A", "#1F9D4C", "#14743A"];
const SIZES = [16, 32, 64, 80, 128];

const lum = (h: string) => {
  const c = [1, 3, 5]
    .map((i) => parseInt(h.slice(i, i + 2), 16) / 255)
    .map((v) => (v <= 0.03928 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4)));
  return 0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2];
};
const ratio = (a: string, b: string) => {
  const [hi, lo] = [lum(a), lum(b)].sort((x, y) => y - x);
  return ((hi + 0.05) / (lo + 0.05)).toFixed(2);
};

const IconStrip = () => (
  <div>
    <h2 style={{ fontSize: 18 }}>
      Ícone do suplemento: &quot;D&quot; verde em branco (16 / 32 / 64 / 80 / 128 px)
    </h2>
    {GREENS.map((g) => (
      <div key={g} style={{ marginBottom: 16 }}>
        <Text weight="semibold">
          {g} (contraste com branco {ratio(g, "#ffffff")}:1, mínimo 3:1 para elementos gráficos)
        </Text>
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginTop: 8 }}>
          {[
            ["Faixa clara", "#eef1f3"],
            ["Faixa escura", "#1f1f1f"],
          ].map(([name, bg]) => (
            <div
              key={name}
              style={{
                background: bg,
                padding: 12,
                display: "flex",
                alignItems: "flex-end",
                gap: 12,
                borderRadius: 6,
              }}
            >
              <span
                style={{
                  color: bg === "#1f1f1f" ? "#fff" : "#000",
                  fontSize: 12,
                  alignSelf: "center",
                  width: 70,
                }}
              >
                {name}
              </span>
              {SIZES.map((sz) => (
                <div
                  key={sz}
                  style={{
                    textAlign: "center",
                    color: bg === "#1f1f1f" ? "#fff" : "#000",
                    fontSize: 12,
                  }}
                >
                  <DIcon size={sz} green={g} />
                  {sz}
                </div>
              ))}
            </div>
          ))}
        </div>
      </div>
    ))}
  </div>
);

const App = () => {
  const [width, setWidth] = React.useState("320");
  const [mode, setMode] = React.useState("light");
  const dark = mode === "dark";
  const w = Number(width);
  return (
    <FluentProvider
      theme={dark ? webDarkTheme : webLightTheme}
      style={{ minHeight: "100vh", padding: 16, background: dark ? "#111" : "#d8dde0" }}
    >
      <h1 style={{ fontSize: 20, margin: "0 0 4px" }}>
        PROTOTYPE F-012 (descartável, dados fictícios, sem Office.js)
      </h1>
      <div style={{ display: "flex", gap: 32, flexWrap: "wrap", marginBottom: 16 }}>
        <RadioGroup
          layout="horizontal"
          value={width}
          onChange={(_, d) => setWidth(d.value)}
          aria-label="Largura"
        >
          <Radio value="320" label="320 px" />
          <Radio value="450" label="450 px" />
        </RadioGroup>
        <RadioGroup
          layout="horizontal"
          value={mode}
          onChange={(_, d) => setMode(d.value)}
          aria-label="Tema"
        >
          <Radio value="light" label="Claro" />
          <Radio value="dark" label="Escuro" />
        </RadioGroup>
      </div>
      <div
        style={{
          display: "flex",
          gap: 24,
          alignItems: "flex-start",
          overflowX: "auto",
          paddingBottom: 16,
        }}
      >
        {VARIANTS.map((v) => (
          <div key={v.label} style={{ flex: "none" }}>
            <p style={{ margin: "0 0 4px", fontWeight: 600, color: dark ? "#fff" : "#000" }}>
              {v.label} &middot; {w} px &middot; {dark ? "escuro" : "claro"}
            </p>
            <FluentProvider
              theme={themes(v.accent)[mode as "light" | "dark"]}
              style={{
                width: w,
                boxSizing: "border-box",
                border: "1px solid #888",
                overflow: "hidden",
              }}
            >
              <Panel dir={v.dir} dark={dark} wide={w >= 400} />
            </FluentProvider>
          </div>
        ))}
      </div>
      <div style={{ color: dark ? "#fff" : "#000" }}>
        <IconStrip />
      </div>
    </FluentProvider>
  );
};

const el = document.getElementById("container");
if (el) createRoot(el).render(<App />);
