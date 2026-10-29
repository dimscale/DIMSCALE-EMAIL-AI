import * as React from "react";
import {
  Badge,
  Card,
  FluentProvider,
  Tab,
  TabList,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ShieldCheckmarkRegular } from "@fluentui/react-icons";
import BackendStatus from "./BackendStatus";
import AiTools from "./AiTools";
import MessagePanel from "./MessagePanel";
import ResponderPanel from "./ResponderPanel";
import { strings } from "./strings";
import { useOfficeTheme } from "./theme";
import { Wordmark } from "./Wordmark";

type View = "analyze" | "reply";

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
  assistantIntro: {
    margin: "4px 0 0",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
  },
  tools: { display: "flex", flexDirection: "column", gap: "8px" },
  toolCard: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    padding: "12px",
    minWidth: 0,
    maxWidth: "100%",
    boxSizing: "border-box",
  },
  toolTitle: { margin: 0, fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold },
  toolDescription: {
    margin: 0,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: "18px",
  },
  footer: { padding: "8px 12px 12px", minWidth: 0, boxSizing: "border-box", borderTop: `1px solid ${tokens.colorNeutralStroke2}` },
});

const AiToolCard = ({ title, description }: { title: string; description: string }) => {
  const s = useStyles();
  return (
    <Card className={s.toolCard}>
      <div className={s.row}>
        <p className={s.toolTitle}>{title}</p>
        <Badge appearance="outline">{strings.aiComingSoon}</Badge>
      </div>
      <p className={s.toolDescription}>{description}</p>
    </Card>
  );
};

const App: React.FC = () => {
  const s = useStyles();
  const { kind, theme } = useOfficeTheme();
  const [noticeHost, setNoticeHost] = React.useState<HTMLElement | null>(null);
  const [view, setView] = React.useState<View>("analyze");
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
          <Tab value="analyze" id="tab-analyze" aria-controls="panel-analyze">
            {strings.tabAnalyze}
          </Tab>
          <Tab value="reply" id="tab-reply" aria-controls="panel-reply">
            {strings.tabReply}
          </Tab>
        </TabList>
        <div ref={setNoticeHost} />
        {/* Both panels stay mounted so the message state and ItemChanged survive tab switches. */}
        <div
          role="tabpanel"
          id="panel-analyze"
          aria-labelledby="tab-analyze"
          className={s.panel}
          hidden={view !== "analyze"}
          style={view !== "analyze" ? { display: "none" } : undefined}
        >
          <MessagePanel
            renderBelow={(ctx) => (
              <section aria-labelledby="ai-tools-heading">
                <h2 id="ai-tools-heading" className={s.toolTitle}>
                  {strings.aiAssistantTitle}
                </h2>
                <p className={s.assistantIntro}>{strings.aiAssistantIntro}</p>
                <div className={s.tools}>
                  {ctx && <AiTools key={ctx.itemKey} view={ctx.view} />}
                  <AiToolCard
                    title={strings.aiImproveTitle}
                    description={strings.aiImproveDescription}
                  />
                </div>
              </section>
            )}
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
          id="panel-reply"
          aria-labelledby="tab-reply"
          className={s.panel}
          hidden={view !== "reply"}
          style={view !== "reply" ? { display: "none" } : undefined}
        >
          <ResponderPanel active={view === "reply"} />
        </div>
      </main>
      <footer className={s.footer}>
        <BackendStatus noticeHost={noticeHost} />
      </footer>
    </FluentProvider>
  );
};

export default App;
