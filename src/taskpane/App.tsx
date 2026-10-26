import * as React from "react";
import { Card, FluentProvider, Tab, TabList, makeStyles, tokens } from "@fluentui/react-components";
import { ShieldCheckmarkRegular, SparkleRegular } from "@fluentui/react-icons";
import BackendStatus from "./BackendStatus";
import MessagePanel from "./MessagePanel";
import ResponderPanel from "./ResponderPanel";
import { strings } from "./strings";
import { useOfficeTheme } from "./theme";
import { Wordmark } from "./Wordmark";

type View = "analyze" | "reply";

const useStyles = makeStyles({
  root: { minHeight: "100vh", display: "flex", flexDirection: "column", fontSize: tokens.fontSizeBase300 },
  header: { display: "flex", alignItems: "center", height: "40px", padding: "0 12px", boxSizing: "border-box" },
  h1: { display: "flex", alignItems: "center", gap: "8px", margin: 0, fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold },
  main: { flex: 1, padding: "0 12px", maxWidth: "640px" },
  panel: { display: "flex", flexDirection: "column", gap: "12px", paddingTop: "8px", paddingBottom: "12px" },
  row: { display: "flex", alignItems: "center", gap: "8px" },
  icon: { display: "flex", fontSize: "20px", flex: "none" },
  secondary: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  footer: { padding: "8px 12px 12px", borderTop: `1px solid ${tokens.colorNeutralStroke2}` },
});

const ComingSoon = ({ text, icon }: { text: string; icon?: React.ReactNode }) => {
  const s = useStyles();
  return (
    <Card>
      <div className={s.row}>
        {icon && (
          <span className={s.icon} aria-hidden="true">
            {icon}
          </span>
        )}
        <span>{text}</span>
      </div>
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
        <TabList selectedValue={view} onTabSelect={(_, d) => setView(d.value as View)} aria-label={strings.tabsLabel}>
          <Tab value="analyze" id="tab-analyze" aria-controls="panel-analyze">
            {strings.tabAnalyze}
          </Tab>
          <Tab value="reply" id="tab-reply" aria-controls="panel-reply">
            {strings.tabReply}
          </Tab>
        </TabList>
        <div ref={setNoticeHost} />
        {/* Both panels stay mounted so the message state and ItemChanged survive tab switches. */}
        <div role="tabpanel" id="panel-analyze" aria-labelledby="tab-analyze" className={s.panel} hidden={view !== "analyze"}>
          <MessagePanel />
          <ComingSoon text={strings.comingSoonAnalyze} icon={<SparkleRegular />} />
          <div className={s.row}>
            <span className={s.icon} aria-hidden="true">
              <ShieldCheckmarkRegular />
            </span>
            <span className={s.secondary}>{strings.assurance}</span>
          </div>
        </div>
        <div role="tabpanel" id="panel-reply" aria-labelledby="tab-reply" className={s.panel} hidden={view !== "reply"}>
          <ResponderPanel />
        </div>
      </main>
      <footer className={s.footer}>
        <BackendStatus noticeHost={noticeHost} />
      </footer>
    </FluentProvider>
  );
};

export default App;
