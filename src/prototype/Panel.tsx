/* THROWAWAY prototype (F-012 step 1). Fake data, no Office.js, no network. */
import * as React from "react";
import {
  Accordion,
  AccordionHeader,
  AccordionItem,
  AccordionPanel,
  Badge,
  Button,
  Card,
  Divider,
  Tab,
  TabList,
  Text,
  Textarea,
  Tooltip,
  makeStyles,
  mergeClasses,
  tokens,
} from "@fluentui/react-components";
import {
  ArrowClockwiseRegular,
  CalendarClockRegular,
  ChatMultipleRegular,
  ChatRegular,
  CheckmarkCircleFilled,
  DocumentBulletListRegular,
  EmojiRegular,
  MailRegular,
  ShieldCheckmarkRegular,
  SparkleRegular,
  TableSimpleRegular,
  TextDescriptionRegular,
  TextGrammarWandRegular,
  TranslateRegular,
  WarningRegular,
  WrenchRegular,
} from "@fluentui/react-icons";
import { C } from "./themes";
import { Wordmark } from "./Wordmark";

export type Direction = "A" | "B" | "C";

const useStyles = makeStyles({
  root: {
    minHeight: "100%",
    fontSize: tokens.fontSizeBase300,
    lineHeight: tokens.lineHeightBase300,
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    height: "44px",
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
  tabs: { paddingTop: "4px" },
  section: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    paddingTop: "12px",
    paddingBottom: "12px",
  },
  heading: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    margin: 0,
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground1,
  },
  icon: { fontSize: "20px", flex: "none" },
  secondary: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  list: { margin: 0, paddingLeft: "20px", display: "flex", flexDirection: "column", gap: "4px" },
  chips: { display: "flex", flexWrap: "wrap", gap: "6px" },
  field: { display: "flex", flexDirection: "column", gap: "0" },
  fieldWide: { display: "grid", gridTemplateColumns: "120px 1fr", columnGap: "8px" },
  label: { color: tokens.colorNeutralForeground3, fontSize: tokens.fontSizeBase200 },
  value: { overflowWrap: "anywhere" },
  preview: {
    whiteSpace: "pre-wrap",
    overflowWrap: "anywhere",
    maxHeight: "96px",
    overflowY: "auto",
    padding: "8px",
    borderRadius: tokens.borderRadiusLarge,
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    ":focus-visible": { outline: `2px solid ${tokens.colorStrokeFocus2}`, outlineOffset: "1px" },
  },
  textarea: {
    width: "100%",
    borderRadius: tokens.borderRadiusLarge,
    "> textarea": { minHeight: "112px" },
  },
  primary: { borderRadius: tokens.borderRadiusLarge, minHeight: "36px" },
  buttons: { display: "flex", flexWrap: "wrap", gap: "8px" },
  full: { width: "100%" },
  btn: { borderRadius: tokens.borderRadiusLarge },
  tiles: { display: "grid", gap: "8px" },
  tile: {
    borderRadius: tokens.borderRadiusLarge,
    height: "auto",
    minHeight: "64px",
    minWidth: 0,
    padding: "8px 4px",
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "normal",
    textAlign: "center",
  },
  row: { display: "flex", alignItems: "center", gap: "8px" },
  ok: { color: tokens.colorStatusSuccessForeground1, fontSize: "16px", flex: "none" },
  footer: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    paddingTop: "12px",
    paddingBottom: "12px",
  },
  dot: {
    width: "8px",
    height: "8px",
    borderRadius: "50%",
    backgroundColor: tokens.colorStatusSuccessForeground1,
    flex: "none",
  },
});

type Pal = { page: string; card: string };
const palette = (d: Direction, dark: boolean): Pal =>
  dark
    ? { page: C.navyDeep, card: d === "C" ? C.navy : C.navyDeep }
    : { page: d === "C" ? C.mist : d === "B" ? C.paper : "#ffffff", card: "#ffffff" };

const Section = ({
  dir,
  pal,
  id,
  icon,
  title,
  first,
  children,
}: {
  dir: Direction;
  pal: Pal;
  id: string;
  icon: React.ReactNode;
  title: string;
  first?: boolean;
  children: React.ReactNode;
}) => {
  const s = useStyles();
  const body = (
    <>
      <h2 id={id} className={s.heading}>
        <span className={s.icon} aria-hidden="true">
          {icon}
        </span>
        {title}
      </h2>
      {children}
    </>
  );
  if (dir === "C") {
    return (
      <Card
        role="group"
        aria-labelledby={id}
        style={{
          backgroundColor: pal.card,
          gap: "8px",
          padding: "16px",
          marginBottom: "12px",
          boxShadow: "none",
          border: `1px solid ${tokens.colorNeutralStroke2}`,
        }}
      >
        {body}
      </Card>
    );
  }
  return (
    <>
      {!first && <Divider />}
      <section aria-labelledby={id} className={s.section}>
        {body}
      </section>
    </>
  );
};

const Field = ({ label, value, wide }: { label: string; value: string; wide: boolean }) => {
  const s = useStyles();
  return (
    <div className={mergeClasses(s.field, wide && s.fieldWide)}>
      <span className={s.label}>{label}</span>
      <span className={s.value} dir="auto">
        {value}
      </span>
    </div>
  );
};

const TOOLS: [string, React.ReactElement][] = [
  ["Resumir", <TextDescriptionRegular key="a" />],
  ["Resumo da conversa", <ChatMultipleRegular key="b" />],
  ["Traduzir", <TranslateRegular key="c" />],
  ["Melhorar texto", <TextGrammarWandRegular key="d" />],
  ["Extrair dados", <TableSimpleRegular key="e" />],
  ["Verificar tom", <EmojiRegular key="f" />],
  ["Seguimento", <CalendarClockRegular key="g" />],
];

const REPLY =
  "Bom dia Marta,\n\nObrigado pelo contacto. Confirmamos a receção do mapa de quantidades da Obra 203 e enviaremos a proposta até 15 de novembro. Propomos a visita técnica na próxima semana; indique-nos duas datas que lhe convenham.\n\nCom os melhores cumprimentos,\nThiago Santos";

const EXCERPT =
  "Bom dia,\nSegue em anexo o mapa de quantidades para a reabilitação do Edifício Aurora (Obra 203). Solicitamos proposta de orçamento até 15 de novembro e uma visita técnica ao local antes da entrega.\nCom os melhores cumprimentos,\nMarta Exemplo";

export const Panel = ({ dir, dark, wide }: { dir: Direction; dark: boolean; wide: boolean }) => {
  const s = useStyles();
  const pal = palette(dir, dark);
  const [tab, setTab] = React.useState<string>("analisar");
  const pad = wide ? "16px" : "12px";
  const band = dir === "B";
  const logo = band || dark ? "#ffffff" : "#231f20";

  const message = (
    <Accordion collapsible defaultOpenItems="msg">
      <AccordionItem value="msg">
        <AccordionHeader icon={<MailRegular />} expandIconPosition="end" size="large">
          Mensagem
        </AccordionHeader>
        <AccordionPanel>
          <div className={s.section} style={{ paddingTop: 0 }}>
            <Field
              wide={wide}
              label="Assunto"
              value="Pedido de orçamento – Reabilitação Edifício Aurora (Obra 203)"
            />
            <Field wide={wide} label="De" value="Marta Exemplo <marta.exemplo@cliente.example>" />
            <Field wide={wide} label="Data" value="22/10/2026 09:14" />
            <span className={s.label}>Excerto</span>
            <div
              className={s.preview}
              role="region"
              aria-label="Pré-visualização do corpo"
              tabIndex={0}
              dir="auto"
            >
              {EXCERPT}
            </div>
            <span className={s.secondary}>(truncado) Comprimento do corpo: 1840 caracteres</span>
          </div>
        </AccordionPanel>
      </AccordionItem>
    </Accordion>
  );

  const reply = (
    <Section
      dir={dir}
      pal={pal}
      id={`${dir}-reply`}
      icon={<ChatMultipleRegular />}
      title="Resposta sugerida"
      first={tab === "responder"}
    >
      <Textarea
        className={s.textarea}
        readOnly
        value={REPLY}
        resize="none"
        aria-label="Resposta sugerida"
      />
      <Button appearance="primary" className={mergeClasses(s.primary, !wide && s.full)}>
        Abrir resposta para revisão
      </Button>
      <div className={s.buttons}>
        <Button appearance="outline" className={s.btn} icon={<ArrowClockwiseRegular />}>
          Regenerar
        </Button>
        <Button appearance="outline" className={s.btn} icon={<ChatRegular />}>
          Ajustar tom
        </Button>
      </div>
    </Section>
  );

  return (
    <div
      className={s.root}
      style={{ backgroundColor: pal.page, color: tokens.colorNeutralForeground1 }}
    >
      <header
        className={s.header}
        style={{
          padding: `0 ${pad}`,
          ...(band
            ? { backgroundColor: C.navy, borderBottom: dark ? `2px solid ${C.teal}` : undefined }
            : {}),
        }}
      >
        <h1 className={s.h1} style={{ color: band ? "#ffffff" : undefined }}>
          <Wordmark color={logo} width={wide ? 120 : 108} />
          <span style={{ color: band ? C.teal : undefined }}>Mail AI</span>
        </h1>
      </header>
      {!band && !dark && dir === "A" && <Divider />}
      {!band && dark && dir === "A" && (
        <div style={{ height: 1, backgroundColor: C.teal, opacity: 0.6 }} />
      )}
      <main style={{ padding: `0 ${pad}` }}>
        <div className={s.tabs}>
          <TabList
            selectedValue={tab}
            onTabSelect={(_, d) => setTab(String(d.value))}
            aria-label="Vistas"
          >
            <Tab value="analisar">Analisar</Tab>
            <Tab value="responder">Responder</Tab>
          </TabList>
        </div>
        <div style={{ paddingTop: dir === "C" ? "12px" : 0 }}>
          {dir === "C" ? (
            <Card
              style={{
                backgroundColor: pal.card,
                padding: "4px 8px",
                marginBottom: "12px",
                boxShadow: "none",
                border: `1px solid ${tokens.colorNeutralStroke2}`,
              }}
            >
              {message}
            </Card>
          ) : (
            message
          )}
          {tab === "analisar" && (
            <>
              <Section
                dir={dir}
                pal={pal}
                id={`${dir}-sum`}
                icon={<SparkleRegular />}
                title="Resumo inteligente"
              >
                <ul className={s.list}>
                  <li>
                    O cliente pede um orçamento para a reabilitação do Edifício Aurora (Obra 203).
                  </li>
                  <li>A proposta deve ser entregue até 15 de novembro.</li>
                  <li>Pede também uma visita técnica ao local antes da entrega.</li>
                </ul>
                <div className={s.chips}>
                  <Badge
                    size="large"
                    appearance="tint"
                    color="brand"
                    shape="rounded"
                    icon={<MailRegular />}
                  >
                    Pedido de orçamento
                  </Badge>
                  <Badge
                    size="large"
                    appearance="tint"
                    color="danger"
                    shape="rounded"
                    icon={<WarningRegular />}
                  >
                    Prioridade alta
                  </Badge>
                  <Badge
                    size="large"
                    appearance="tint"
                    color="informative"
                    shape="rounded"
                    icon={<EmojiRegular />}
                  >
                    Tom neutro
                  </Badge>
                </div>
                <span className={s.secondary}>Sugestões da IA, podem estar erradas.</span>
              </Section>
              <Section
                dir={dir}
                pal={pal}
                id={`${dir}-det`}
                icon={<DocumentBulletListRegular />}
                title="Detalhes extraídos"
              >
                <Field wide={wide} label="Prazo" value="15 de novembro de 2026" />
                <Field wide={wide} label="Projeto" value="Obra 203 – Edifício Aurora" />
                <Field
                  wide={wide}
                  label="Ação necessária"
                  value="Enviar proposta e agendar visita técnica"
                />
              </Section>
            </>
          )}
          {reply}
          {tab === "analisar" && (
            <>
              <Section
                dir={dir}
                pal={pal}
                id={`${dir}-tools`}
                icon={<WrenchRegular />}
                title="Ferramentas rápidas"
              >
                <div
                  className={s.tiles}
                  style={{ gridTemplateColumns: `repeat(${wide ? 4 : 3}, minmax(0, 1fr))` }}
                >
                  {TOOLS.map(([name, icon]) => (
                    <Tooltip key={name} content={name} relationship="description" withArrow>
                      <Button appearance="outline" className={s.tile}>
                        <span style={{ fontSize: 20, display: "flex" }} aria-hidden="true">
                          {icon}
                        </span>
                        <span>{name}</span>
                      </Button>
                    </Tooltip>
                  ))}
                </div>
              </Section>
              <Section
                dir={dir}
                pal={pal}
                id={`${dir}-act`}
                icon={<ShieldCheckmarkRegular />}
                title="Atividade IA"
              >
                <div className={s.row}>
                  <CheckmarkCircleFilled className={s.ok} aria-hidden="true" />
                  <span>Email analisado</span>
                </div>
                <div className={s.row}>
                  <CheckmarkCircleFilled className={s.ok} aria-hidden="true" />
                  <span>Resposta aberta para revisão</span>
                </div>
                <Tooltip
                  content="O envio é sempre feito por si."
                  relationship="description"
                  withArrow
                >
                  <div className={s.row} tabIndex={0} style={{ width: "fit-content" }}>
                    <ShieldCheckmarkRegular className={s.icon} aria-hidden="true" />
                    <Text weight="semibold">Sem envio automático</Text>
                  </div>
                </Tooltip>
              </Section>
            </>
          )}
        </div>
      </main>
      <Divider />
      <footer className={s.footer} style={{ padding: `12px ${pad}` }}>
        <div className={s.row} role="status">
          <span className={s.dot} aria-hidden="true" />
          <span className={s.secondary}>Ligado como Thiago Santos</span>
        </div>
        <span className={s.secondary}>Fixe o painel 📌 para o manter aberto entre emails</span>
      </footer>
    </div>
  );
};
