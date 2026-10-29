import * as React from "react";
import { FluentProvider, Tab, TabList, makeStyles, tokens } from "@fluentui/react-components";
import { ShieldCheckmarkRegular } from "@fluentui/react-icons";
import BackendStatus from "./BackendStatus";
import AiTools, { DEFAULT_OPTIONS, ToolOptions, ToolTab } from "./AiTools";
import { useAutoSummary } from "./autoSummary";
import MessagePanel from "./MessagePanel";
import ResponderPanel from "./ResponderPanel";
import { strings } from "./strings";
import { useOfficeTheme } from "./theme";
import { Wordmark } from "./Wordmark";

type View = ToolTab | "auto";

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    width: "100%",
    maxWidth: "100%",
    boxSizing: "border-box",
    overflowX: "hidden", // safety net only: nothing inside should exceed the pane
    overflowWrap: "anywhere",
    display: "flex",
    flexDirection: "column",
    fontSize: tokens.fontSizeBase300,
  },
  header: {
    display: "flex",
    alignItems: "center",
    height: "40px",
    padding: "0 12px",
    boxSizing: "border-box",
  },
  h1: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    margin: 0,
    fontSize: tokens.fontSizeBase400,
    fontWeight: tokens.fontWeightSemibold,
  },
  main: {
    flex: 1,
    minWidth: 0,
    width: "100%",
    boxSizing: "border-box",
    padding: "0 12px",
    maxWidth: "640px",
  },
  panel: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    minWidth: 0,
    paddingTop: "8px",
    paddingBottom: "12px",
    "& > *": { minWidth: 0, maxWidth: "100%" },
    "&[hidden]": { display: "none" },
  },
  row: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px", minWidth: 0 },
  icon: { display: "flex", fontSize: "20px", flex: "none" },
  secondary: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  disclaimer: {
    margin: "0 0 8px",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  footer: {
    padding: "8px 12px 12px",
    minWidth: 0,
    boxSizing: "border-box",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
});

// summaryEntry: the pane was opened from the "Resumo" ribbon button (resumo.html): one summary, as if clicked.
const App: React.FC<{ summaryEntry?: boolean }> = ({ summaryEntry = false }) => {
  const s = useStyles();
  const auto = useAutoSummary();
  const [options, setOptions] = React.useState<ToolOptions>(DEFAULT_OPTIONS);
  const { kind, theme } = useOfficeTheme();
  const [noticeHost, setNoticeHost] = React.useState<HTMLElement | null>(null);
  const [view, setView] = React.useState<View>("summary");
  const logo = kind === "dark" ? "#ffffff" : kind === "light" ? "#000000" : "currentColor";
  return (
    <FluentProvider theme={theme} className={s.root}>
      <header className={s.header}>
        <h1 className={s.h1}>
          <Wordmark color={logo} />
          <span>{strings.brandSuffix}</span>
        </h1>
      </header>
      <main className={s.main}>
        <TabList
          selectedValue={view}
          onTabSelect={(_, d) => setView(d.value as View)}
          aria-label={strings.tabsLabel}
        >
          <Tab value="summary" id="tab-summary" aria-controls="panel-tools">
            {strings.tabSummary}
          </Tab>
          <Tab value="reply" id="tab-reply" aria-controls="panel-tools">
            {strings.tabReply}
          </Tab>
          <Tab value="translate" id="tab-translate" aria-controls="panel-tools">
            {strings.tabTranslate}
          </Tab>
          <Tab value="auto" id="tab-auto" aria-controls="panel-auto">
            {strings.tabAuto}
          </Tab>
        </TabList>
        <div ref={setNoticeHost} />
        {/* Message context, the three cards and the Automático panel stay mounted so state survives tab switches. */}
        <div
          role="tabpanel"
          id="panel-tools"
          aria-labelledby={`tab-${view === "auto" ? "summary" : view}`}
          className={s.panel}
          hidden={view === "auto"}
          style={view === "auto" ? { display: "none" } : undefined}
        >
          <MessagePanel
            renderBelow={(ctx) =>
              ctx && (
                <AiTools
                  key={ctx.itemKey}
                  view={ctx.view}
                  itemKey={ctx.itemKey}
                  tab={view}
                  options={options}
                  onOptions={setOptions}
                  auto={auto}
                  entryRun={summaryEntry && ctx.itemKey === 0}
                />
              )
            }
          />
          <div className={s.row}>
            <span className={s.icon} aria-hidden="true">
              <ShieldCheckmarkRegular />
            </span>
            <span className={s.secondary}>{strings.assurance}</span>
          </div>
        </div>
        <div
          role="tabpanel"
          id="panel-auto"
          aria-labelledby="tab-auto"
          className={s.panel}
          hidden={view !== "auto"}
          style={view !== "auto" ? { display: "none" } : undefined}
        >
          <ResponderPanel active={view === "auto"} />
        </div>
      </main>
      <footer className={s.footer}>
        <p className={s.disclaimer}>{strings.aiFooter}</p>
        <BackendStatus noticeHost={noticeHost} />
      </footer>
    </FluentProvider>
  );
};

export default App;
