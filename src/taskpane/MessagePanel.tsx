import * as React from "react";
import { toMessageView, MessageView } from "./messageView";
import { strings } from "./strings";

/* global Office */

type State = { kind: "loading" } | { kind: "none" } | { kind: "error" } | { kind: "ok"; view: MessageView };

const MessagePanel: React.FC = () => {
  const [state, setState] = React.useState<State>({ kind: "loading" });
  const [attempt, setAttempt] = React.useState(0);

  React.useEffect(() => {
    const item = Office.context.mailbox.item as Office.MessageRead | undefined;
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
          setState({ kind: "error" }); // never log result.error: it may carry content
        }
      });
    } catch {
      setState({ kind: "error" }); // never log the exception: it may carry content
    }
    return () => {
      active = false;
    };
  }, [attempt]);

  if (state.kind === "loading") return <p>{strings.loading}</p>;
  if (state.kind === "none") return <p>{strings.noItem}</p>;
  if (state.kind === "error") {
    return (
      <div>
        <p role="alert">{strings.error}</p>
        <button onClick={() => setAttempt((n) => n + 1)}>{strings.retry}</button>
      </div>
    );
  }
  const v = state.view;
  return (
    <div style={{ padding: 12 }}>
      <h2>{strings.title}</h2>
      <p>
        <strong>{strings.subject}:</strong> {v.subject}
      </p>
      <p>
        <strong>{strings.from}:</strong> {v.senderName ? `${v.senderName} <${v.senderEmail}>` : v.senderEmail}
      </p>
      <p>
        <strong>{strings.date}:</strong> {v.date}
      </p>
      <p>
        <strong>{strings.body}:</strong>
      </p>
      <pre style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>{v.body.trim() ? v.body : strings.noBody}</pre>
      {v.truncated && <p>{strings.truncated}</p>}
      <p>{strings.bodyLength(v.bodyLength)}</p>
    </div>
  );
};

export default MessagePanel;
