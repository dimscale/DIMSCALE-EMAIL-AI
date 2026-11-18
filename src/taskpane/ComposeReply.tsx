import * as React from "react";
import {
  Button,
  Field,
  Input,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Select,
  Spinner,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { analyze, RETRYABLE, ReplyLanguage, ReplyTone } from "./analyzeClient";
import {
  errorText,
  isBlank,
  LANGUAGES,
  MAX_DRAFT,
  MAX_HISTORY,
  MAX_INSTRUCTION,
  TONES,
} from "./AiTools";
import { getTokenSilent } from "./auth";
import { composePayload, Draft, isCompose, readDraft } from "./composeItem";
import { insertIntoDraft } from "./composeInsert";
import { strings } from "./strings";
import ToolCard from "./ToolCard";

/* global Office, navigator */

const useStyles = makeStyles({
  row: { display: "flex", flexWrap: "wrap", alignItems: "center", gap: "8px", minWidth: 0 },
  field: { flex: "1 1 120px", minWidth: 0 },
  note: {
    margin: 0,
    overflowWrap: "anywhere",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    lineHeight: "18px",
  },
  code: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  refine: { display: "flex", flexWrap: "wrap", gap: "4px", alignItems: "flex-end", minWidth: 0 },
  pre: { margin: 0, whiteSpace: "pre-wrap", overflowWrap: "anywhere", userSelect: "text" },
});

type Err = { text: string; code?: string; retry?: boolean };
type Insert = "idle" | "busy" | "done" | "failed";

/* F-010: the only tool in compose is the reply card. Every request and every write needs a click. The parent
   remounts this component (new key) on ItemChanged, so state, the snapshot and late responses never cross drafts. */
const ComposeReply: React.FC = () => {
  const s = useStyles();
  const [language, setLanguage] = React.useState<ReplyLanguage>("match");
  const [tone, setTone] = React.useState<ReplyTone>("none");
  const [pending, setPending] = React.useState(false);
  const [reply, setReply] = React.useState<string | null>(null);
  const [error, setError] = React.useState<Err | null>(null);
  const [copy, setCopy] = React.useState<boolean | null>(null);
  const [insert, setInsert] = React.useState<Insert>("idle");
  const [instruction, setInstruction] = React.useState("");
  const [history, setHistory] = React.useState<string[]>([]);
  const [revised, setRevised] = React.useState(false);
  const inputRef = React.useRef<HTMLInputElement>(null);
  const busy = React.useRef(false);
  const inserting = React.useRef(false);
  const mounted = React.useRef(true);
  const ctl = React.useRef<AbortController | null>(null);
  const snapshot = React.useRef<Draft | null>(null); // draft read at the first Sugerir click of the chain (refine uses it)

  React.useEffect(() => {
    mounted.current = true;
    return () => {
      mounted.current = false;
      ctl.current?.abort(); // item changed: abort the in-flight request
    };
  }, []);

  const run = async (refine?: string) => {
    if (busy.current) return;
    if (
      refine !== undefined &&
      (reply === null ||
        reply.length > MAX_DRAFT ||
        isBlank(refine) ||
        refine.length > MAX_INSTRUCTION)
    )
      return;
    busy.current = true;
    const c = new AbortController();
    ctl.current = c;
    setPending(true);
    setError(null);
    setCopy(null);
    setRevised(false);
    try {
      let draft = refine !== undefined ? snapshot.current : null;
      if (!draft) {
        const item = Office.context.mailbox?.item;
        draft = isCompose(item) ? await readDraft(item) : null;
        if (!mounted.current) return;
        if (!draft) {
          setError({ text: strings.composeReadFailed });
          return;
        }
        if (isBlank(draft.body)) {
          setError({ text: strings.composeEmpty });
          return;
        }
        snapshot.current = draft;
      }
      const token = await getTokenSilent();
      if (!mounted.current) return;
      if (token.kind !== "token") {
        setError({ text: strings.aiNeedSignIn });
        return;
      }
      const res = await analyze(
        token.token,
        composePayload(
          draft,
          language,
          tone,
          refine !== undefined ? { instruction: refine, draft: reply as string } : undefined
        ),
        undefined,
        c.signal
      );
      if (!mounted.current || res.kind === "aborted") return;
      if (res.kind === "reply") {
        if (refine !== undefined && reply !== null) {
          setHistory((h) => [...h, reply].slice(-MAX_HISTORY));
          setInstruction("");
          setRevised(true);
        } else setHistory([]);
        setReply(res.reply);
        setInsert("idle"); // a new text can be inserted
      } else if (res.kind === "error") {
        setError({
          text:
            refine !== undefined && res.code === "invalid"
              ? strings.aiRefineInvalid
              : errorText[res.code],
          code: res.why,
          retry: refine === undefined && RETRYABLE.includes(res.code),
        });
      } else setError({ text: errorText.other });
    } finally {
      busy.current = false;
      if (ctl.current === c) ctl.current = null;
      if (mounted.current) {
        setPending(false);
        if (refine !== undefined) inputRef.current?.focus();
      }
    }
  };

  const copyText = async (text: string) => {
    let ok = false;
    try {
      await navigator.clipboard.writeText(text);
      ok = true;
    } catch {
      ok = false; // clipboard missing or blocked in this webview
    }
    if (mounted.current) setCopy(ok);
  };

  const insertReply = async () => {
    if (reply === null || inserting.current || insert === "done") return;
    inserting.current = true;
    setInsert("busy");
    const item = Office.context.mailbox?.item;
    const ok = isCompose(item) && (await insertIntoDraft(item, reply));
    inserting.current = false;
    if (mounted.current) setInsert(ok ? "done" : "failed");
  };

  const blank = isBlank(instruction) || instruction.trim().length > MAX_INSTRUCTION;
  const tooLong = reply !== null && reply.length > MAX_DRAFT;
  const sendRefine = () => {
    if (!blank && !pending && !tooLong) void run(instruction.trim());
  };
  const prevVersion = () => {
    if (history.length === 0 || pending) return;
    if (history.length === 1) inputRef.current?.focus(); // the button disappears with the last step
    setReply(history[history.length - 1]);
    setHistory(history.slice(0, -1));
    setCopy(null);
    setInsert("idle");
  };

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

  return (
    <ToolCard
      title={strings.composeTitle}
      options={
        <>
          <p className={s.note}>{strings.composePrivacyNotice}</p>
          <div className={s.row}>
            {select(strings.replyLanguageLabel, language, LANGUAGES, setLanguage)}
            {select(strings.replyToneLabel, tone, TONES, setTone)}
          </div>
        </>
      }
      action={
        <Button
          appearance={reply === null ? "primary" : "secondary"}
          disabled={pending}
          aria-busy={pending}
          icon={pending ? <Spinner size="tiny" /> : undefined}
          onClick={() => void run()}
        >
          {pending ? strings.aiReplying : strings.aiSuggestReply}
        </Button>
      }
      error={
        error && (
          <MessageBar intent="error" role="alert">
            <MessageBarBody>
              {error.text}
              {error.code && <span className={s.code}> {strings.aiErrCode(error.code)}</span>}
            </MessageBarBody>
            {error.retry && (
              <MessageBarActions>
                <Button onClick={() => void run()} disabled={pending}>
                  {strings.retry}
                </Button>
              </MessageBarActions>
            )}
          </MessageBar>
        )
      }
      result={
        reply !== null ? (
          <p className={s.pre} dir="auto">
            {reply}
          </p>
        ) : null
      }
      onCopy={() => reply !== null && void copyText(reply)}
      after={
        <>
          <p className={s.note}>{strings.aiAdvisory}</p>
          <Button
            appearance="primary"
            disabled={insert === "busy" || insert === "done"}
            onClick={() => void insertReply()}
          >
            {insert === "busy" ? strings.composeInserting : strings.composeInsert}
          </Button>
          <p className={s.note}>{strings.composeInsertHelp}</p>
          {insert === "done" && (
            <span className={s.note} role="status">
              {strings.composeInserted}
            </span>
          )}
          {insert === "failed" && (
            <MessageBar intent="error" role="alert">
              <MessageBarBody>{strings.composeInsertFailed}</MessageBarBody>
            </MessageBar>
          )}
          {copy !== null && (
            <span className={s.note} role="status">
              {copy ? strings.aiCopied : strings.aiCopyFailed}
            </span>
          )}
          {tooLong ? (
            <p className={s.note} role="status">
              {strings.aiRefineTooLong}
            </p>
          ) : (
            <Field label={strings.aiRefineLabel}>
              <div className={s.refine}>
                <Input
                  ref={inputRef}
                  style={{ flex: "1 1 120px", minWidth: "120px" }}
                  value={instruction}
                  maxLength={MAX_INSTRUCTION}
                  placeholder={strings.aiRefinePlaceholder}
                  readOnly={pending}
                  aria-disabled={pending}
                  onChange={(_, d) => setInstruction(d.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter" && !e.nativeEvent.isComposing) {
                      e.preventDefault();
                      sendRefine();
                    }
                  }}
                />
                <Button appearance="secondary" disabled={blank || pending} onClick={sendRefine}>
                  {strings.aiRefineSend}
                </Button>
              </div>
            </Field>
          )}
          <span className={s.note} role="status">
            {revised ? strings.aiRefineDone : ""}
          </span>
          {history.length > 0 && (
            <Button disabled={pending} onClick={prevVersion}>
              {strings.aiPrevVersion}
            </Button>
          )}
        </>
      }
    />
  );
};

export default ComposeReply;
