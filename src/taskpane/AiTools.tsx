import * as React from "react";
import {
  Badge,
  Button,
  Card,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { analyze, AnalyzeAction, AnalyzeErrorCode, RETRYABLE, Summary } from "./analyzeClient";
import { getTokenSilent } from "./auth";
import { MessageView, toReplyHtml } from "./messageView";
import { openReplyForm } from "./replyForm";
import { strings } from "./strings";

/* global Office, navigator */

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "8px" },
  row: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px" },
  note: {
    margin: 0,
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: "18px",
  },
  card: { display: "flex", flexDirection: "column", gap: "6px", padding: "12px" },
  title: { margin: 0, fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold },
  list: { margin: 0, paddingLeft: "20px", overflowWrap: "anywhere" },
  text: { margin: 0, overflowWrap: "anywhere" },
  reply: { margin: 0, whiteSpace: "pre-wrap", overflowWrap: "anywhere", userSelect: "text" },
});

const errorText: Record<AnalyzeErrorCode, string> = {
  unauthorized: strings.aiErrUnauthorized,
  payload: strings.aiErrPayload,
  cap: strings.aiErrCap,
  disabled: strings.aiErrDisabled,
  not_configured: strings.aiErrNotConfigured,
  unavailable: strings.aiErrUnavailable,
  invalid: strings.aiErrInvalid,
  other: strings.aiErrOther,
};

type Err = { text: string; retry?: AnalyzeAction };
type CopyState = { which: "summary" | "reply"; ok: boolean } | null;

// Own-property lookup: an unknown enum value (or "__proto__") yields no Badge.
const label = (map: Record<string, string>, k: string) =>
  Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;

const AiTools: React.FC<{ view: MessageView }> = ({ view }) => {
  const s = useStyles();
  const [pending, setPending] = React.useState<AnalyzeAction | null>(null);
  const [summary, setSummary] = React.useState<Summary | null>(null);
  const [reply, setReply] = React.useState<string | null>(null);
  const [error, setError] = React.useState<Err | null>(null);
  const [copy, setCopy] = React.useState<CopyState>(null);
  const [openFailed, setOpenFailed] = React.useState(false);
  const busy = React.useRef(false);
  const mounted = React.useRef(true);
  // The pane remounts this component (new key) on ItemChanged, so results and late responses never cross emails.
  React.useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
    };
  }, []);

  const empty = view.aiText.trim() === "";

  const run = async (action: AnalyzeAction) => {
    if (busy.current || empty) return;
    busy.current = true;
    setPending(action);
    setError(null);
    setCopy(null);
    setOpenFailed(false);
    try {
      const auth = await getTokenSilent();
      if (!mounted.current) return;
      if (auth.kind !== "token") {
        setError({ text: strings.aiNeedSignIn });
        return;
      }
      const res = await analyze(auth.token, {
        action,
        subject: view.aiSubject,
        fromName: view.aiFromName,
        fromDomain: view.fromDomain,
        body: view.aiText,
      });
      if (!mounted.current) return;
      if (res.kind === "summary") setSummary(res.summary);
      else if (res.kind === "reply") setReply(res.reply);
      else
        setError({
          text: errorText[res.code],
          retry: RETRYABLE.includes(res.code) ? action : undefined,
        });
    } finally {
      busy.current = false;
      if (mounted.current) setPending(null);
    }
  };

  const copyText = async (which: "summary" | "reply", text: string) => {
    let ok = false;
    try {
      await navigator.clipboard.writeText(text);
      ok = true;
    } catch {
      ok = false; // clipboard missing or blocked in this webview
    }
    if (mounted.current) setCopy({ which, ok });
  };

  const openReply = async () => {
    if (reply === null) return;
    setOpenFailed(false);
    const item = Office.context.mailbox?.item as Office.MessageRead | undefined | null;
    const ok = await openReplyForm(item, toReplyHtml(reply));
    if (mounted.current && !ok) setOpenFailed(true);
  };

  const copyStatus = (which: "summary" | "reply") =>
    copy?.which === which ? (
      <span className={s.note} role="status">
        {copy.ok ? strings.aiCopied : strings.aiCopyFailed}
      </span>
    ) : null;

  const typeLabel = summary && label(strings.aiTypeLabels, summary.type);
  const priorityLabel = summary && label(strings.aiPriorityLabels, summary.priority);
  const toneLabel = summary && label(strings.aiToneLabels, summary.tone);

  return (
    <div className={s.root}>
      <div className={s.row}>
        <Button
          appearance="primary"
          disabled={empty || pending !== null}
          aria-busy={pending === "summarize"}
          icon={pending === "summarize" ? <Spinner size="tiny" /> : undefined}
          onClick={() => void run("summarize")}
        >
          {pending === "summarize" ? strings.aiSummarizing : strings.aiSummarize}
        </Button>
        <Button
          disabled={empty || pending !== null}
          aria-busy={pending === "reply"}
          icon={pending === "reply" ? <Spinner size="tiny" /> : undefined}
          onClick={() => void run("reply")}
        >
          {pending === "reply" ? strings.aiReplying : strings.aiSuggestReply}
        </Button>
      </div>
      <p className={s.note}>{strings.aiPrivacyNotice}</p>
      {empty && <p className={s.note}>{strings.aiNoText}</p>}
      {error && (
        <MessageBar intent="error" role="alert">
          <MessageBarBody>{error.text}</MessageBarBody>
          {error.retry && (
            <MessageBarActions>
              <Button
                onClick={() => void run(error.retry as AnalyzeAction)}
                disabled={pending !== null}
              >
                {strings.retry}
              </Button>
            </MessageBarActions>
          )}
        </MessageBar>
      )}
      {summary && (
        <Card className={s.card}>
          <h3 className={s.title}>{strings.aiSummaryTitle}</h3>
          <ul className={s.list}>
            {summary.points.map((p, i) => (
              <li key={i} dir="auto">
                {p}
              </li>
            ))}
          </ul>
          <div className={s.row}>
            {typeLabel && <Badge appearance="outline">{typeLabel}</Badge>}
            {priorityLabel && <Badge appearance="outline">{priorityLabel}</Badge>}
            {toneLabel && <Badge appearance="outline">{toneLabel}</Badge>}
          </div>
          {summary.deadline && (
            <p className={s.text} dir="auto">
              {strings.aiDeadline(summary.deadline)}
            </p>
          )}
          {summary.requestedAction && (
            <p className={s.text} dir="auto">
              {strings.aiRequestedAction(summary.requestedAction)}
            </p>
          )}
          <p className={s.note}>{strings.aiAdvisory}</p>
          <div className={s.row}>
            <Button
              onClick={() =>
                void copyText("summary", summary.points.map((p) => `• ${p}`).join("\n"))
              }
            >
              {strings.aiCopy}
            </Button>
            {copyStatus("summary")}
          </div>
        </Card>
      )}
      {reply !== null && (
        <Card className={s.card}>
          <h3 className={s.title}>{strings.aiReplyTitle}</h3>
          <p className={s.reply} dir="auto">
            {reply}
          </p>
          <p className={s.note}>{strings.aiAdvisory}</p>
          <div className={s.row}>
            <Button appearance="primary" onClick={() => void openReply()}>
              {strings.aiOpenReply}
            </Button>
            <Button onClick={() => void copyText("reply", reply)}>{strings.aiCopy}</Button>
          </div>
          {copyStatus("reply")}
          {openFailed && (
            <p className={s.note} role="status">
              {strings.aiOpenReplyFailed}
            </p>
          )}
        </Card>
      )}
    </div>
  );
};

export default AiTools;
