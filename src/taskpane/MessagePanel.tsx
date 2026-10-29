import * as React from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Button,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Skeleton,
  SkeletonItem,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { MailRegular } from "@fluentui/react-icons";
import { toMessageView, MessageView } from "./messageView";
import { strings } from "./strings";

/* global Office */

type State =
  | { kind: "loading" }
  | { kind: "none" }
  | { kind: "error"; view: MessageView }
  | { kind: "ok"; view: MessageView };

const useStyles = makeStyles({
  fields: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    paddingTop: "8px",
    paddingBottom: "8px",
  },
  label: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  value: { overflowWrap: "anywhere" },
  preview: {
    margin: 0,
    fontFamily: "inherit",
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
    maxHeight: "120px",
    overflowY: "auto",
    padding: "8px",
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "1px" },
  },
  secondary: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    display: "block",
    paddingTop: "4px",
  },
  empty: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "8px",
    padding: "24px 0",
    textAlign: "center",
  },
  hidden: {
    position: "absolute",
    width: "1px",
    height: "1px",
    overflow: "hidden",
    clip: "rect(0 0 0 0)",
    whiteSpace: "nowrap",
  },
});

const sender = (v: MessageView) =>
  v.senderName && v.senderEmail
    ? `${v.senderName} <${v.senderEmail}>`
    : v.senderName || v.senderEmail || "—";

export type BelowCtx = { view: MessageView; itemKey: number } | null;

// renderBelow: the App renders the AI section after the message; ctx is null unless the message loaded (ok state).
const MessagePanel: React.FC<{ renderBelow?: (ctx: BelowCtx) => React.ReactNode }> = ({
  renderBelow,
}) => {
  const s = useStyles();
  const [state, setState] = React.useState<State>({ kind: "loading" });
  const [attempt, setAttempt] = React.useState(0);

  // Re-read when the user selects another message (pinned pane). Event args are ignored: the item is read fresh.
  React.useEffect(() => {
    const mailbox = Office.context?.mailbox;
    if (!mailbox?.addHandlerAsync) return undefined;
    const onChange = () => setAttempt((n) => n + 1);
    mailbox.addHandlerAsync(Office.EventType.ItemChanged, onChange);
    return () => mailbox.removeHandlerAsync?.(Office.EventType.ItemChanged, onChange);
  }, []);

  React.useEffect(() => {
    const item = Office.context.mailbox?.item as Office.MessageRead | undefined | null; // undefined outside Outlook, null when none selected
    if (!item) {
      setState({ kind: "none" });
      return undefined;
    }
    let active = true;
    setState({ kind: "loading" });
    try {
      item.body.getAsync(Office.CoercionType.Text, (result: Office.AsyncResult<string>) => {
        if (!active) return;
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setState({ kind: "ok", view: toMessageView(item, result.value) });
        } else {
          setState({ kind: "error", view: toMessageView(item, "") }); // never log result.error: it may carry content
        }
      });
    } catch {
      setState({ kind: "error", view: toMessageView(item, "") }); // never log the exception: it may carry content
    }
    return () => {
      active = false;
    };
  }, [attempt]);

  const below = renderBelow?.(state.kind === "ok" ? { view: state.view, itemKey: attempt } : null);
  if (state.kind === "loading") {
    return (
      <>
        <Skeleton aria-busy="true" role="status">
          <span className={s.hidden}>{strings.loading}</span>
          <div className={s.fields} aria-hidden="true">
            <SkeletonItem size={16} />
            <SkeletonItem size={16} />
            <SkeletonItem size={16} />
          </div>
        </Skeleton>
        {below}
      </>
    );
  }
  if (state.kind === "none") {
    return (
      <>
        <div className={s.empty}>
          <MailRegular fontSize={32} aria-hidden="true" />
          <p>{strings.noItem}</p>
        </div>
        {below}
      </>
    );
  }
  const v = state.view;
  const field = (label: string, value: string) => (
    <div>
      <div className={s.label}>{label}</div>
      <div className={s.value} dir="auto">
        {value}
      </div>
    </div>
  );
  return (
    <>
      <div>
        <div className={s.fields}>
          {field(strings.subject, v.subject)}
          {field(strings.from, sender(v))}
          {field(strings.date, v.date)}
        </div>
        {state.kind === "error" ? (
          <MessageBar intent="error" role="alert">
            <MessageBarBody>{strings.error}</MessageBarBody>
            <MessageBarActions>
              <Button onClick={() => setAttempt((n) => n + 1)}>{strings.retry}</Button>
            </MessageBarActions>
          </MessageBar>
        ) : (
          <Accordion collapsible>
            <AccordionItem value="msg">
              <AccordionHeader icon={<MailRegular />} expandIconPosition="end">
                {strings.messageSection}
              </AccordionHeader>
              <AccordionPanel>
                <pre
                  className={s.preview}
                  role="region"
                  aria-label={strings.bodyPreview}
                  tabIndex={0}
                  dir="auto"
                >
                  {v.body.trim() ? v.body : strings.noBody}
                </pre>
                {v.truncated && <span className={s.secondary}>{strings.truncated}</span>}
                <span className={s.secondary}>{strings.bodyLength(v.bodyLength)}</span>
              </AccordionPanel>
            </AccordionItem>
          </Accordion>
        )}
      </div>
      {below}
    </>
  );
};

export default MessagePanel;
