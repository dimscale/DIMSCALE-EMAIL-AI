import * as React from "react";
import {
  Button,
  Checkbox,
  Input,
  MessageBar,
  MessageBarBody,
  Spinner,
  Switch,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { ShieldCheckmarkRegular, AddRegular } from "@fluentui/react-icons";
import { getTokenSilent } from "./auth";
import {
  CandidateContact,
  fetchRecentSenders,
  getGraphTokenSilent,
  requestGraphConsent,
} from "./graph";
import {
  deleteSettings,
  fetchSettings,
  saveSettings,
  ServerStatus,
} from "./settingsClient";
import { strings } from "./strings";

/* global Office */

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    paddingTop: "8px",
    paddingBottom: "12px",
  },
  row: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
  },
  statusText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  sectionTitle: {
    fontSize: tokens.fontSizeBase300,
    fontWeight: tokens.fontWeightSemibold,
    margin: 0,
  },
  listContainer: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    maxHeight: "260px",
    overflowY: "auto",
    paddingRight: "4px",
  },
  contactItem: {
    display: "flex",
    alignItems: "flex-start",
  },
  contactText: {
    overflowWrap: "anywhere",
    fontSize: tokens.fontSizeBase200,
  },
  manualForm: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  manualInput: {
    flex: 1,
  },
  buttonRow: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    marginTop: "4px",
  },
  assuranceRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: "8px",
    marginTop: "4px",
  },
  assuranceIcon: {
    fontSize: "20px",
    color: tokens.colorNeutralForeground3,
    flex: "none",
  },
  assuranceText: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    lineHeight: "16px",
  },
  confirmBox: {
    padding: "8px",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  confirmButtons: {
    display: "flex",
    gap: "8px",
  },
});

const EMAIL_REGEX = /^[^@\s]+@[^@\s]+\.[^@\s]+$/;

export const ResponderPanel: React.FC = () => {
  const s = useStyles();

  const [enabled, setEnabled] = React.useState<boolean>(false);
  const [serverStatus, setServerStatus] = React.useState<ServerStatus>("off");
  const [selectedContacts, setSelectedContacts] = React.useState<string[]>([]);
  const [candidates, setCandidates] = React.useState<CandidateContact[]>([]);
  const [manualAddress, setManualAddress] = React.useState<string>("");

  const [graphConsentNeeded, setGraphConsentNeeded] = React.useState<boolean>(false);
  const [graphDenied, setGraphDenied] = React.useState<boolean>(false);

  const [loadingInitial, setLoadingInitial] = React.useState<boolean>(true);
  const [loadingSenders, setLoadingSenders] = React.useState<boolean>(false);
  const [saving, setSaving] = React.useState<boolean>(false);
  const [confirmDelete, setConfirmDelete] = React.useState<boolean>(false);

  const [notice, setNotice] = React.useState<{ intent: "info" | "success" | "error"; text: string } | null>(null);

  const getOwnAddress = (): string => {
    try {
      return Office.context?.mailbox?.userProfile?.emailAddress?.trim().toLowerCase() || "";
    } catch {
      return "";
    }
  };

  // 1. Initial load: read settings from backend and check Graph token
  React.useEffect(() => {
    let active = true;

    async function init() {
      setLoadingInitial(true);
      const authRes = await getTokenSilent();
      if (!active) return;

      if (authRes?.kind === "token") {
        const settingsRes = await fetchSettings(authRes.token);
        if (active && settingsRes.kind === "ok") {
          setEnabled(settingsRes.data.enabled);
          setServerStatus(settingsRes.data.status);
          setSelectedContacts(settingsRes.data.contacts);
        }
      }

      // Check Graph permissions for candidates
      const graphRes = await getGraphTokenSilent();
      if (!active) return;

      if (graphRes?.kind === "token") {
        setGraphConsentNeeded(false);
        setLoadingSenders(true);
        try {
          const list = await fetchRecentSenders(graphRes.token, getOwnAddress());
          if (active) setCandidates(list);
        } catch {
          // Graph call failed gracefully
        } finally {
          if (active) setLoadingSenders(false);
        }
      } else if (graphRes?.kind === "needConsent") {
        setGraphConsentNeeded(true);
      } else {
        setGraphDenied(true);
      }

      if (active) setLoadingInitial(false);
    }

    init();

    return () => {
      active = false;
    };
  }, []);

  const handleRequestConsent = async () => {
    setNotice(null);
    const res = await requestGraphConsent();
    if (res.kind === "token") {
      setGraphConsentNeeded(false);
      setGraphDenied(false);
      setLoadingSenders(true);
      try {
        const list = await fetchRecentSenders(res.token, getOwnAddress());
        setCandidates(list);
      } catch {
        // graceful fallback
      } finally {
        setLoadingSenders(false);
      }
    } else {
      setGraphConsentNeeded(false);
      setGraphDenied(true);
    }
  };

  const handleToggleSwitch = (_: unknown, data: { checked: boolean }) => {
    setNotice(null);
    if (data.checked && selectedContacts.length === 0) {
      setNotice({ intent: "error", text: strings.needAtLeastOneContact });
      return;
    }
    setEnabled(data.checked);
  };

  const handleToggleContact = (address: string) => {
    setNotice(null);
    const lower = address.toLowerCase();
    const isSelected = selectedContacts.includes(lower);

    if (isSelected) {
      setSelectedContacts(selectedContacts.filter((c) => c !== lower));
    } else {
      if (selectedContacts.length >= 50) {
        setNotice({ intent: "error", text: strings.contactsCapReached });
        return;
      }
      setSelectedContacts([...selectedContacts, lower]);
    }
  };

  const handleAddManual = (e: React.FormEvent) => {
    e.preventDefault();
    setNotice(null);
    const raw = manualAddress.trim().toLowerCase();
    if (!raw || !EMAIL_REGEX.test(raw) || raw.length > 254) {
      setNotice({ intent: "error", text: strings.invalidEmail });
      return;
    }
    const own = getOwnAddress();
    if (own && raw === own) {
      setNotice({ intent: "error", text: strings.ownAddressNotAllowed });
      return;
    }
    if (selectedContacts.includes(raw)) {
      setNotice({ intent: "error", text: strings.contactAlreadyAdded });
      return;
    }
    if (selectedContacts.length >= 50) {
      setNotice({ intent: "error", text: strings.contactsCapReached });
      return;
    }

    setSelectedContacts([raw, ...selectedContacts]);
    setManualAddress("");
  };

  const handleSave = async () => {
    setNotice(null);
    if (enabled && selectedContacts.length === 0) {
      setNotice({ intent: "error", text: strings.needAtLeastOneContact });
      return;
    }

    setSaving(true);
    const authRes = await getTokenSilent();
    if (authRes.kind !== "token") {
      setNotice({ intent: "error", text: strings.signInError });
      setSaving(false);
      return;
    }

    const payload = {
      enabled,
      ownAddress: getOwnAddress(),
      contacts: selectedContacts,
    };

    const res = await saveSettings(authRes.token, payload);
    setSaving(false);

    if (res.kind === "ok") {
      setEnabled(res.data.enabled);
      setServerStatus(res.data.status);
      setSelectedContacts(res.data.contacts);
      setNotice({ intent: "success", text: strings.saveSuccess });
    } else {
      setNotice({ intent: "error", text: strings.shortError });
    }
  };

  const handleDeleteAll = async () => {
    setNotice(null);
    const authRes = await getTokenSilent();
    if (authRes.kind !== "token") {
      setNotice({ intent: "error", text: strings.signInError });
      return;
    }

    setSaving(true);
    const res = await deleteSettings(authRes.token);
    setSaving(false);
    setConfirmDelete(false);

    if (res.ok) {
      setEnabled(false);
      setServerStatus("off");
      setSelectedContacts([]);
      setNotice({ intent: "success", text: strings.saveSuccess });
    } else {
      setNotice({ intent: "error", text: strings.shortError });
    }
  };

  const renderStatus = () => {
    switch (serverStatus) {
      case "active":
        return strings.statusActive;
      case "pending":
        return strings.statusPending;
      case "not_in_scope":
        return strings.statusNotInScope;
      case "error":
        return strings.statusError;
      default:
        return strings.statusOff;
    }
  };

  // Combine selected contacts and candidates, ensuring selected are on top and deduplicated
  const displayedContacts = React.useMemo(() => {
    const list: Array<{ address: string; name: string; selected: boolean }> = [];
    const seen = new Set<string>();

    for (const addr of selectedContacts) {
      seen.add(addr);
      const cand = candidates.find((c) => c.address === addr);
      list.push({
        address: addr,
        name: cand?.name || "",
        selected: true,
      });
    }

    for (const c of candidates) {
      if (!seen.has(c.address)) {
        seen.add(c.address);
        list.push({
          address: c.address,
          name: c.name,
          selected: false,
        });
      }
    }

    return list;
  }, [selectedContacts, candidates]);

  const atCap = selectedContacts.length >= 50;

  if (loadingInitial) {
    return (
      <div className={s.container}>
        <Spinner size="medium" label={strings.loadingSettings} />
      </div>
    );
  }

  return (
    <div className={s.container}>
      <div className={s.row}>
        <Switch
          checked={enabled}
          onChange={handleToggleSwitch}
          label={strings.autoRepliesSwitch}
        />
        <span className={s.statusText}>
          {strings.statusLabel} {renderStatus()}
        </span>
      </div>

      {notice && (
        <MessageBar intent={notice.intent}>
          <MessageBarBody>{notice.text}</MessageBarBody>
        </MessageBar>
      )}

      {graphConsentNeeded && (
        <MessageBar intent="info">
          <MessageBarBody>
            <div className={s.row}>
              <span>{strings.permissionNeeded}</span>
              <Button size="small" appearance="primary" onClick={handleRequestConsent}>
                {strings.allowAccess}
              </Button>
            </div>
          </MessageBarBody>
        </MessageBar>
      )}

      {graphDenied && (
        <MessageBar intent="warning">
          <MessageBarBody>{strings.permissionDenied}</MessageBarBody>
        </MessageBar>
      )}

      <div>
        <h2 className={s.sectionTitle}>
          {strings.contactsHeader(selectedContacts.length)}
        </h2>
        {atCap && <span className={s.statusText}>{strings.contactsCapReached}</span>}
      </div>

      {loadingSenders && <Spinner size="tiny" label={strings.loadingSenders} />}

      <div className={s.listContainer} role="group" aria-label={strings.contactsHeader(selectedContacts.length)}>
        {displayedContacts.map((c) => {
          const isSelected = c.selected;
          const labelText = c.name && c.name !== c.address ? `${c.name} <${c.address}>` : c.address;
          return (
            <div key={c.address} className={s.contactItem} dir="auto">
              <Checkbox
                checked={isSelected}
                disabled={!isSelected && atCap}
                onChange={() => handleToggleContact(c.address)}
                label={{
                  children: <span className={s.contactText}>{labelText}</span>,
                }}
                aria-label={`${c.name} ${c.address}`}
              />
            </div>
          );
        })}
      </div>

      <form onSubmit={handleAddManual} className={s.manualForm}>
        <Input
          className={s.manualInput}
          value={manualAddress}
          onChange={(_, d) => setManualAddress(d.value)}
          placeholder={strings.addAddressPlaceholder}
          aria-label={strings.addAddressPlaceholder}
          type="email"
        />
        <Button
          appearance="secondary"
          icon={<AddRegular />}
          onClick={handleAddManual}
          disabled={!manualAddress.trim() || atCap}
        >
          {strings.addAddressButton}
        </Button>
      </form>

      {confirmDelete ? (
        <div className={s.confirmBox}>
          <span>{strings.deleteAllConfirm}</span>
          <div className={s.confirmButtons}>
            <Button appearance="primary" onClick={handleDeleteAll} disabled={saving}>
              {strings.confirmYes}
            </Button>
            <Button appearance="secondary" onClick={() => setConfirmDelete(false)}>
              {strings.confirmCancel}
            </Button>
          </div>
        </div>
      ) : (
        <div className={s.buttonRow}>
          <Button
            appearance="primary"
            onClick={handleSave}
            disabled={saving}
          >
            {saving ? strings.saving : strings.saveButton}
          </Button>
          <Button
            appearance="subtle"
            onClick={() => setConfirmDelete(true)}
            disabled={saving || selectedContacts.length === 0}
          >
            {strings.deleteAllButton}
          </Button>
        </div>
      )}

      <div className={s.assuranceRow}>
        <ShieldCheckmarkRegular className={s.assuranceIcon} aria-hidden="true" />
        <span className={s.assuranceText}>{strings.responderAssurance}</span>
      </div>
    </div>
  );
};

export default ResponderPanel;
