import * as React from "react";
import {
  Badge,
  Button,
  Field,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Select,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import {
  analyze,
  AnalyzeAction,
  AnalyzeErrorCode,
  AnalyzePayload,
  RETRYABLE,
  ReplyLanguage,
  ReplyTone,
  Summary,
  TargetLanguage,
  Translation,
} from "./analyzeClient";
import { AutoSummary, AUTO_DEBOUNCE_MS } from "./autoSummary";
import AutoSummarySwitch from "./AutoSummarySwitch";
import { getTokenSilent } from "./auth";
import { MessageView, toReplyHtml, translatePayload } from "./messageView";
import { openReplyForm } from "./replyForm";
import { strings } from "./strings";
import ToolCard from "./ToolCard";

/* global Office, navigator */

export type ToolTab = "summary" | "reply" | "translate";
export type ToolOptions = { language: ReplyLanguage; tone: ReplyTone; target: TargetLanguage };
export const DEFAULT_OPTIONS: ToolOptions = { language: "match", tone: "none", target: "auto" };

const useStyles = makeStyles({
  root: { display: "flex", flexDirection: "column", gap: "8px", minWidth: 0, maxWidth: "100%" },
  row: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px", minWidth: 0 },
  field: { flex: "1 1 120px", minWidth: 0 },
  note: {
    margin: 0,
    overflowWrap: "anywhere",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: "18px",
  },
  list: { margin: 0, paddingLeft: "20px", overflowWrap: "anywhere" },
  text: { margin: 0, overflowWrap: "anywhere" },
  pre: { margin: 0, whiteSpace: "pre-wrap", overflowWrap: "anywhere", userSelect: "text" },
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
// Auto mode pauses for the pane session on these (F-003 AC-44).
const PAUSE_CODES: AnalyzeErrorCode[] = ["cap", "unauthorized", "disabled", "not_configured"];

type Err = { action: AnalyzeAction; text: string; retry?: boolean };
type CopyState = { which: AnalyzeAction; ok: boolean } | null;

const LANGUAGES: [ReplyLanguage, string][] = [
  ["match", strings.replyLanguageMatch],
  ["pt", strings.replyLanguagePt],
  ["en", strings.replyLanguageEn],
];
const TONES: [ReplyTone, string][] = [
  ["none", strings.replyToneNone],
  ["concise", strings.replyToneConcise],
  ["detailed", strings.replyToneDetailed],
  ["formal", strings.replyToneFormal],
  ["friendly", strings.replyToneFriendly],
];
const TARGETS: [TargetLanguage, string][] = [
  ["auto", strings.translateAuto],
  ["pt", strings.translatePt],
  ["en", strings.translateEn],
];

// Own-property lookup: an unknown enum value (or "__proto__") yields no Badge.
const label = (map: Record<string, string>, k: string) =>
  Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;

type Props = {
  view: MessageView;
  itemKey: number;
  tab: ToolTab | "auto";
  options: ToolOptions;
  onOptions: (o: ToolOptions) => void;
  auto: AutoSummary;
  entryRun?: boolean; // resumo.html opened on this item: one summary, counted as a click (F-012 AC-50)
};

const AiTools: React.FC<Props> = ({ view, itemKey, tab, options, onOptions, auto, entryRun }) => {
  const s = useStyles();
  const [pending, setPending] = React.useState<AnalyzeAction | null>(null);
  const [summary, setSummary] = React.useState<Summary | null>(null);
  const [reply, setReply] = React.useState<string | null>(null);
  const [translation, setTranslation] = React.useState<Translation | null>(null);
  const [error, setError] = React.useState<Err | null>(null);
  const [copy, setCopy] = React.useState<CopyState>(null);
  const [openFailed, setOpenFailed] = React.useState(false);
  const busy = React.useRef(false);
  const mounted = React.useRef(true);
  const autoCtl = React.useRef<AbortController | null>(null);
  const id = view.itemId || `k${itemKey}`;
  const empty = view.aiText.trim() === "";

  // The pane remounts this component (new key) on ItemChanged, so results and late responses never cross emails.
  React.useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      autoCtl.current?.abort(); // item changed: abort the in-flight auto request
    };
  }, []);

  const run = async (action: AnalyzeAction, isAuto = false) => {
    if (busy.current || empty) return;
    busy.current = true;
    if (action === "summarize") auto.attempted.add(id);
    const ctl = isAuto ? new AbortController() : null;
    autoCtl.current = ctl;
    setPending(action);
    setError(null);
    setCopy(null);
    setOpenFailed(false);
    try {
      const token = await getTokenSilent();
      if (!mounted.current) return;
      if (token.kind !== "token") {
        setError({ action, text: strings.aiNeedSignIn });
        return;
      }
      const base = {
        subject: view.aiSubject,
        fromName: view.aiFromName,
        fromDomain: view.fromDomain,
        body: view.aiText,
      };
      const payload: AnalyzePayload =
        action === "summarize"
          ? { action, ...base }
          : action === "reply"
            ? { action, ...base, language: options.language, tone: options.tone }
            : translatePayload(view, options.target);
      const res = await analyze(token.token, payload, undefined, ctl?.signal);
      if (!mounted.current || res.kind === "aborted") return;
      if (res.kind === "summary") setSummary(res.summary);
      else if (res.kind === "reply") setReply(res.reply);
      else if (res.kind === "translation") setTranslation(res.translation);
      else if (res.kind === "error") {
        let text = errorText[res.code];
        if (isAuto && PAUSE_CODES.includes(res.code)) {
          text = res.code === "cap" ? strings.autoPausedCap : `${text} ${strings.autoPausedSuffix}`;
          auto.pause();
        }
        setError({ action, text, retry: RETRYABLE.includes(res.code) });
      }
    } finally {
      busy.current = false;
      if (autoCtl.current === ctl) autoCtl.current = null;
      if (mounted.current) setPending(null);
    }
  };

  // Entry page resumo.html: one summary for the item loaded at start-up, as if clicked.
  React.useEffect(() => {
    if (entryRun) void run("summarize");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Switching auto mode off aborts an in-flight auto request at once.
  React.useEffect(() => {
    if (!auto.on) autoCtl.current?.abort();
  }, [auto.on]);

  // Auto mode: one summary per item per pane session, 800 ms after the item (or the Resumo tab) settles.
  React.useEffect(() => {
    if (!auto.active || tab !== "summary" || empty || auto.attempted.has(id)) return undefined;
    const t = setTimeout(() => {
      auto.attempted.add(id);
      void run("summarize", true);
    }, AUTO_DEBOUNCE_MS);
    return () => clearTimeout(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [auto.active, tab, empty, id]);

  const copyText = async (which: AnalyzeAction, text: string) => {
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

  const after = (which: AnalyzeAction, extra?: React.ReactNode) => (
    <>
      <p className={s.note}>{strings.aiAdvisory}</p>
      {extra}
      {copy?.which === which && (
        <span className={s.note} role="status">
          {copy.ok ? strings.aiCopied : strings.aiCopyFailed}
        </span>
      )}
    </>
  );

  const button = (action: AnalyzeAction, text: string, busyText: string, primary: boolean) => (
    <Button
      appearance={primary ? "primary" : "secondary"}
      disabled={empty || pending !== null}
      aria-busy={pending === action}
      icon={pending === action ? <Spinner size="tiny" /> : undefined}
      onClick={() => void run(action)}
    >
      {pending === action ? busyText : text}
    </Button>
  );

  const errorBar = (action: AnalyzeAction) => (
    <>
      {empty && <p className={s.note}>{strings.aiNoText}</p>}
      {error?.action === action && (
        <MessageBar intent="error" role="alert">
          <MessageBarBody>{error.text}</MessageBarBody>
          {error.retry && (
            <MessageBarActions>
              <Button onClick={() => void run(action)} disabled={pending !== null}>
                {strings.retry}
              </Button>
            </MessageBarActions>
          )}
        </MessageBar>
      )}
    </>
  );

  const typeLabel = summary && label(strings.aiTypeLabels, summary.type);
  const priorityLabel = summary && label(strings.aiPriorityLabels, summary.priority);
  const toneLabel = summary && label(strings.aiToneLabels, summary.tone);

  const select = <T extends string>(
    name: string,
    value: T,
    items: [T, string][],
    onChange: (v: T) => void
  ) => (
    <Field label={name} className={s.field}>
      <Select value={value} onChange={(_, d) => onChange(d.value as T)}>
        {items.map(([v, text]) => (
          <option key={v} value={v}>
            {text}
          </option>
        ))}
      </Select>
    </Field>
  );

  // All three cards stay mounted (results, options and pending state survive tab switches); only one is shown.
  const show = (t: ToolTab) => (tab === t ? undefined : { display: "none" });
  return (
    <div className={s.root}>
      <div hidden={tab !== "summary"} style={show("summary")}>
        <ToolCard
          title={strings.aiSummaryTitle}
          options={auto.available ? <AutoSummarySwitch auto={auto} /> : undefined}
          action={button("summarize", strings.aiSummarize, strings.aiSummarizing, true)}
          error={errorBar("summarize")}
          result={
            summary && (
              <>
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
              </>
            )
          }
          onCopy={() =>
            summary && void copyText("summarize", summary.points.map((p) => `• ${p}`).join("\n"))
          }
          after={after("summarize")}
        />
      </div>
      <div hidden={tab !== "reply"} style={show("reply")}>
        <ToolCard
          title={strings.aiReplyTitle}
          options={
            <div className={s.row}>
              {select(strings.replyLanguageLabel, options.language, LANGUAGES, (language) =>
                onOptions({ ...options, language })
              )}
              {select(strings.replyToneLabel, options.tone, TONES, (tone) =>
                onOptions({ ...options, tone })
              )}
            </div>
          }
          action={button("reply", strings.aiSuggestReply, strings.aiReplying, reply === null)}
          error={errorBar("reply")}
          result={
            reply !== null ? (
              <p className={s.pre} dir="auto">
                {reply}
              </p>
            ) : null
          }
          onCopy={() => reply !== null && void copyText("reply", reply)}
          after={after(
            "reply",
            <>
              <Button appearance="primary" onClick={() => void openReply()}>
                {strings.aiOpenReply}
              </Button>
              {openFailed && (
                <p className={s.note} role="status">
                  {strings.aiOpenReplyFailed}
                </p>
              )}
            </>
          )}
        />
      </div>
      <div hidden={tab !== "translate"} style={show("translate")}>
        <ToolCard
          title={strings.aiTranslationTitle}
          options={select(strings.translateTargetLabel, options.target, TARGETS, (target) =>
            onOptions({ ...options, target })
          )}
          action={button("translate", strings.aiTranslate, strings.aiTranslating, true)}
          error={errorBar("translate")}
          result={
            translation && (
              <>
                {translation.subject.trim() !== "" && (
                  <p className={s.text} dir="auto">
                    {strings.aiTranslatedSubject(translation.subject)}
                  </p>
                )}
                <p className={s.pre} dir="auto">
                  {translation.body}
                </p>
              </>
            )
          }
          onCopy={() =>
            translation &&
            void copyText(
              "translate",
              translation.subject.trim() !== ""
                ? `${translation.subject}\n\n${translation.body}`
                : translation.body
            )
          }
          after={after("translate")}
        />
      </div>
      <p className={s.note}>{strings.aiPrivacyNotice}</p>
    </div>
  );
};

export default AiTools;
