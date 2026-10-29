import * as React from "react";
import { Button, Card, Tooltip, makeStyles, tokens } from "@fluentui/react-components";
import { ArrowMaximizeRegular, ArrowMinimizeRegular, CopyRegular } from "@fluentui/react-icons";
import { strings } from "./strings";

const useStyles = makeStyles({
  card: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    padding: "12px",
    minWidth: 0,
    maxWidth: "100%",
    boxSizing: "border-box",
  },
  head: { display: "flex", alignItems: "center", gap: "4px", minWidth: 0 },
  title: {
    margin: 0,
    flex: "1 1 0",
    minWidth: 0,
    overflowWrap: "anywhere",
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
  },
  result: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    minWidth: 0,
    maxWidth: "100%",
    boxSizing: "border-box",
    overflowWrap: "anywhere",
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "1px" },
  },
  collapsed: { maxHeight: "240px", overflowY: "auto" },
});

type Props = {
  title: string;
  options?: React.ReactNode; // Select / Switch row
  action: React.ReactNode; // primary button
  error?: React.ReactNode;
  result: React.ReactNode | null; // null: no result yet, so no Copiar / Expandir
  onCopy?: () => void;
  after?: React.ReactNode; // advisory, copy status, extra buttons (below the result region)
};

/* F-012 card anatomy: h2 title, optional options, primary button, result region (text nodes only), copy and expand. */
const ToolCard: React.FC<Props> = ({ title, options, action, error, result, onCopy, after }) => {
  const s = useStyles();
  const [expanded, setExpanded] = React.useState(false);
  const titleId = React.useId();
  const hasResult = result !== null;
  const expandLabel = expanded ? strings.aiCollapse : strings.aiExpand;
  return (
    <Card className={s.card}>
      <div className={s.head}>
        <h2 id={titleId} className={s.title}>
          {title}
        </h2>
        {hasResult && onCopy && (
          <Tooltip content={strings.aiCopy} relationship="label">
            <Button appearance="subtle" size="small" icon={<CopyRegular />} onClick={onCopy} />
          </Tooltip>
        )}
        {hasResult && (
          <Tooltip content={expandLabel} relationship="label">
            <Button
              appearance="subtle"
              size="small"
              aria-expanded={expanded}
              icon={expanded ? <ArrowMinimizeRegular /> : <ArrowMaximizeRegular />}
              onClick={() => setExpanded((v) => !v)}
            />
          </Tooltip>
        )}
      </div>
      {options}
      {action}
      {error}
      {hasResult && (
        <div
          role="region"
          aria-labelledby={titleId}
          tabIndex={0}
          className={`${s.result} ${expanded ? "" : s.collapsed}`}
        >
          {result}
        </div>
      )}
      {hasResult && after}
    </Card>
  );
};

export default ToolCard;
