import { useState, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const STATUS = {
  idle: { label: "Pending", color: "#94a3b8" },
  sending: { label: "Sending…", color: "#f59e0b" },
  sent: { label: "Sent ✓", color: "#22c55e" },
  failed: { label: "Failed ✗", color: "#ef4444" },
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function formatPhone(raw) {
  const digits = String(raw).replace(/\D/g, "");
  return digits.startsWith("0") ? "91" + digits.slice(1) : digits;
}

export default function App() {
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [phoneCol, setPhoneCol] = useState("");
  const [msgCol, setMsgCol] = useState("");
  const [customMsg, setCustomMsg] = useState("");
  const [useCustomMsg, setUseCustomMsg] = useState(false);
  const [statuses, setStatuses] = useState({});
  const [sending, setSending] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const [fileName, setFileName] = useState("");
  const [delay, setDelay] = useState(3);
  const fileRef = useRef();
  const abortRef = useRef(false);

  const parseFile = (file) => {
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
      if (!data.length) return;
      const cols = Object.keys(data[0]);
      setHeaders(cols);
      setRows(data.map((r, i) => ({ ...r, __id: i })));
      setStatuses({});
      // Auto-detect phone column
      const ph = cols.find((c) => /phone|mobile|number|whatsapp/i.test(c));
      if (ph) setPhoneCol(ph);
      const msg = cols.find((c) => /message|msg|text/i.test(c));
      if (msg) setMsgCol(msg);
    };
    reader.readAsArrayBuffer(file);
  };

  const onDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    const file = e.dataTransfer.files[0];
    if (file) parseFile(file);
  }, []);

  const onFileChange = (e) => {
    if (e.target.files[0]) parseFile(e.target.files[0]);
  };

  const buildMessage = (row) => {
    if (useCustomMsg) {
      return customMsg.replace(/\{(\w+)\}/g, (_, k) => row[k] ?? "");
    }
    return msgCol ? String(row[msgCol]) : "";
  };

  const openWhatsApp = (phone, message) => {
    const url = `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
    window.open(url, "_blank");
  };

  const sendAll = async () => {
    if (!phoneCol) return alert("Please select the Phone column.");
    if (!useCustomMsg && !msgCol) return alert("Please select a Message column or use a custom message.");
    setSending(true);
    abortRef.current = false;

    for (let i = 0; i < rows.length; i++) {
      if (abortRef.current) break;
      const row = rows[i];
      const id = row.__id;
      setStatuses((s) => ({ ...s, [id]: "sending" }));
      const phone = formatPhone(row[phoneCol]);
      const msg = buildMessage(row);
      try {
        openWhatsApp(phone, msg);
        await sleep(delay * 1000);
        setStatuses((s) => ({ ...s, [id]: "sent" }));
      } catch {
        setStatuses((s) => ({ ...s, [id]: "failed" }));
      }
    }
    setSending(false);
  };

  const sendSingle = (row) => {
    if (!phoneCol) return alert("Please select the Phone column.");
    const phone = formatPhone(row[phoneCol]);
    const msg = buildMessage(row);
    openWhatsApp(phone, msg);
    setStatuses((s) => ({ ...s, [row.__id]: "sent" }));
  };

  const stopSending = () => { abortRef.current = true; };
  const reset = () => { setRows([]); setHeaders([]); setStatuses({}); setFileName(""); setPhoneCol(""); setMsgCol(""); };

  const sentCount = Object.values(statuses).filter((s) => s === "sent").length;
  const failedCount = Object.values(statuses).filter((s) => s === "failed").length;

  return (
    <div style={styles.root}>
      {/* Header */}
      <div style={styles.header}>
        <div style={styles.logo}>
          <svg width="32" height="32" viewBox="0 0 32 32" fill="none">
            <circle cx="16" cy="16" r="16" fill="#25D366"/>
            <path d="M16 7C11.03 7 7 11.03 7 16c0 1.59.42 3.08 1.14 4.37L7 25l4.77-1.12A8.96 8.96 0 0016 25c4.97 0 9-4.03 9-9s-4.03-9-9-9zm4.42 12.42c-.18.5-.88.91-1.45.97-.4.04-.9.07-2.68-.56-2.25-.8-3.7-3.1-3.82-3.24-.12-.14-.97-1.28-.97-2.44 0-1.16.61-1.73.82-1.96.22-.24.47-.3.63-.3h.45c.15 0 .35-.06.54.41.2.48.7 1.7.76 1.82.06.12.1.27.02.43-.08.16-.12.26-.24.4-.12.14-.25.31-.36.42-.12.12-.24.25-.1.48.14.24.62.97 1.33 1.57.91.8 1.67 1.04 1.91 1.16.24.12.38.1.52-.06.14-.16.6-.7.76-.94.16-.24.32-.2.54-.12.22.08 1.42.67 1.66.79.24.12.4.18.46.28.06.1.06.56-.12 1.06z" fill="white"/>
          </svg>
          <span style={styles.logoText}>BulkWA</span>
        </div>
        <p style={styles.tagline}>Send WhatsApp messages to your entire contact list in minutes</p>
      </div>

      <div style={styles.container}>
        {/* Upload Zone */}
        {!rows.length ? (
          <div
            style={{ ...styles.dropZone, ...(dragOver ? styles.dropZoneActive : {}) }}
            onDrop={onDrop}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onClick={() => fileRef.current.click()}
          >
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={onFileChange} />
            <div style={styles.uploadIcon}>📊</div>
            <p style={styles.uploadTitle}>Drop your Excel or CSV file here</p>
            <p style={styles.uploadSub}>or click to browse &nbsp;·&nbsp; .xlsx, .xls, .csv supported</p>
          </div>
        ) : (
          <>
            {/* File info bar */}
            <div style={styles.fileBar}>
              <span style={styles.fileChip}>📄 {fileName}</span>
              <span style={styles.rowCount}>{rows.length} contacts loaded</span>
              <button style={styles.resetBtn} onClick={reset}>✕ Remove</button>
            </div>

            {/* Config panel */}
            <div style={styles.configGrid}>
              <div style={styles.configCard}>
                <label style={styles.label}>📞 Phone Number Column *</label>
                <select style={styles.select} value={phoneCol} onChange={(e) => setPhoneCol(e.target.value)}>
                  <option value="">— select column —</option>
                  {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>

              <div style={styles.configCard}>
                <label style={styles.label}>⏱ Delay Between Messages</label>
                <div style={styles.delayRow}>
                  <input type="range" min="1" max="15" value={delay} onChange={(e) => setDelay(+e.target.value)} style={styles.slider} />
                  <span style={styles.delayVal}>{delay}s</span>
                </div>
              </div>

              <div style={{ ...styles.configCard, gridColumn: "1 / -1" }}>
                <div style={styles.msgToggle}>
                  <label style={styles.label}>💬 Message</label>
                  <div style={styles.toggle}>
                    <button
                      style={{ ...styles.toggleBtn, ...(useCustomMsg ? {} : styles.toggleActive) }}
                      onClick={() => setUseCustomMsg(false)}
                    >From column</button>
                    <button
                      style={{ ...styles.toggleBtn, ...(useCustomMsg ? styles.toggleActive : {}) }}
                      onClick={() => setUseCustomMsg(true)}
                    >Custom template</button>
                  </div>
                </div>

                {!useCustomMsg ? (
                  <select style={styles.select} value={msgCol} onChange={(e) => setMsgCol(e.target.value)}>
                    <option value="">— select message column —</option>
                    {headers.map((h) => <option key={h} value={h}>{h}</option>)}
                  </select>
                ) : (
                  <>
                    <textarea
                      style={styles.textarea}
                      placeholder="Type your message… Use {ColumnName} for dynamic values e.g. Hello {Name}, your order {OrderID} is ready!"
                      value={customMsg}
                      onChange={(e) => setCustomMsg(e.target.value)}
                      rows={3}
                    />
                    <div style={styles.tagHints}>
                      Available tags: {headers.map((h) => (
                        <span key={h} style={styles.tag} onClick={() => setCustomMsg((m) => m + `{${h}}`)}>{`{${h}}`}</span>
                      ))}
                    </div>
                  </>
                )}
              </div>
            </div>

            {/* Stats */}
            {Object.keys(statuses).length > 0 && (
              <div style={styles.statsBar}>
                <span style={{ color: "#22c55e" }}>✓ {sentCount} sent</span>
                <span style={{ color: "#ef4444" }}>✗ {failedCount} failed</span>
                <span style={{ color: "#94a3b8" }}>◷ {rows.length - sentCount - failedCount} pending</span>
              </div>
            )}

            {/* Send All Button */}
            <div style={styles.actionRow}>
              {!sending ? (
                <button style={styles.sendAllBtn} onClick={sendAll}>
                  <span>🚀</span> Send to All {rows.length} Contacts
                </button>
              ) : (
                <button style={{ ...styles.sendAllBtn, background: "#ef4444" }} onClick={stopSending}>
                  ⏹ Stop Sending
                </button>
              )}
            </div>

            {/* Table */}
            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>#</th>
                    {headers.map((h) => <th key={h} style={styles.th}>{h}</th>)}
                    <th style={styles.th}>Status</th>
                    <th style={styles.th}>Action</th>
                  </tr>
                </thead>
                <tbody>
                  {rows.map((row, i) => {
                    const st = statuses[row.__id] || "idle";
                    return (
                      <tr key={row.__id} style={{ background: i % 2 === 0 ? "#0f172a" : "#111827" }}>
                        <td style={styles.td}>{i + 1}</td>
                        {headers.map((h) => <td key={h} style={styles.td}>{String(row[h])}</td>)}
                        <td style={styles.td}>
                          <span style={{ ...styles.statusBadge, background: STATUS[st].color + "22", color: STATUS[st].color }}>
                            {STATUS[st].label}
                          </span>
                        </td>
                        <td style={styles.td}>
                          <button style={styles.sendBtn} onClick={() => sendSingle(row)}>
                            Send
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>

      <div style={styles.footer}>
        ⚠️ This tool opens WhatsApp Web for each contact. Keep your browser tab active while bulk sending. Use delays responsibly to avoid account restrictions.
      </div>
    </div>
  );
}

const styles = {
  root: {
    minHeight: "100vh",
    background: "#020617",
    color: "#e2e8f0",
    fontFamily: "'DM Sans', 'Segoe UI', sans-serif",
    display: "flex",
    flexDirection: "column",
  },
  header: {
    background: "linear-gradient(135deg, #064e3b 0%, #065f46 50%, #022c22 100%)",
    padding: "32px 24px 28px",
    textAlign: "center",
    borderBottom: "1px solid #134e4a",
  },
  logo: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "10px",
    marginBottom: "8px",
  },
  logoText: {
    fontSize: "28px",
    fontWeight: "800",
    color: "#fff",
    letterSpacing: "-0.5px",
  },
  tagline: {
    color: "#6ee7b7",
    margin: 0,
    fontSize: "14px",
  },
  container: {
    maxWidth: "1100px",
    width: "100%",
    margin: "0 auto",
    padding: "28px 20px",
    flex: 1,
  },
  dropZone: {
    border: "2px dashed #1e3a2f",
    borderRadius: "16px",
    padding: "64px 32px",
    textAlign: "center",
    cursor: "pointer",
    transition: "all 0.2s",
    background: "#0a1f14",
  },
  dropZoneActive: {
    borderColor: "#25D366",
    background: "#0d2b1e",
  },
  uploadIcon: { fontSize: "48px", marginBottom: "12px" },
  uploadTitle: { fontSize: "20px", fontWeight: "700", color: "#e2e8f0", margin: "0 0 6px" },
  uploadSub: { color: "#64748b", fontSize: "13px", margin: 0 },
  fileBar: {
    display: "flex",
    alignItems: "center",
    gap: "12px",
    background: "#0d2b1e",
    border: "1px solid #134e4a",
    borderRadius: "10px",
    padding: "10px 16px",
    marginBottom: "20px",
    flexWrap: "wrap",
  },
  fileChip: { fontSize: "14px", color: "#6ee7b7", fontWeight: "600" },
  rowCount: { fontSize: "13px", color: "#94a3b8", marginLeft: "auto" },
  resetBtn: {
    background: "transparent",
    border: "1px solid #334155",
    color: "#94a3b8",
    borderRadius: "6px",
    padding: "4px 10px",
    cursor: "pointer",
    fontSize: "12px",
  },
  configGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "16px",
    marginBottom: "20px",
  },
  configCard: {
    background: "#0f172a",
    border: "1px solid #1e293b",
    borderRadius: "12px",
    padding: "16px",
  },
  label: { display: "block", fontSize: "12px", fontWeight: "600", color: "#94a3b8", marginBottom: "8px", textTransform: "uppercase", letterSpacing: "0.05em" },
  select: {
    width: "100%",
    background: "#1e293b",
    border: "1px solid #334155",
    borderRadius: "8px",
    color: "#e2e8f0",
    padding: "8px 12px",
    fontSize: "14px",
    outline: "none",
  },
  delayRow: { display: "flex", alignItems: "center", gap: "12px" },
  slider: { flex: 1, accentColor: "#25D366" },
  delayVal: { fontSize: "14px", fontWeight: "700", color: "#25D366", minWidth: "28px" },
  msgToggle: { display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "10px" },
  toggle: { display: "flex", gap: "4px", background: "#1e293b", borderRadius: "8px", padding: "3px" },
  toggleBtn: { background: "transparent", border: "none", color: "#64748b", padding: "5px 12px", borderRadius: "6px", cursor: "pointer", fontSize: "12px", fontWeight: "600" },
  toggleActive: { background: "#25D366", color: "#000" },
  textarea: {
    width: "100%",
    background: "#1e293b",
    border: "1px solid #334155",
    borderRadius: "8px",
    color: "#e2e8f0",
    padding: "10px 12px",
    fontSize: "14px",
    resize: "vertical",
    outline: "none",
    boxSizing: "border-box",
    fontFamily: "inherit",
  },
  tagHints: { marginTop: "8px", fontSize: "12px", color: "#64748b", display: "flex", flexWrap: "wrap", gap: "6px", alignItems: "center" },
  tag: { background: "#134e4a", color: "#6ee7b7", padding: "2px 8px", borderRadius: "4px", cursor: "pointer", fontFamily: "monospace" },
  statsBar: {
    display: "flex",
    gap: "20px",
    background: "#0f172a",
    border: "1px solid #1e293b",
    borderRadius: "10px",
    padding: "10px 18px",
    fontSize: "13px",
    fontWeight: "700",
    marginBottom: "16px",
  },
  actionRow: { marginBottom: "20px", display: "flex", justifyContent: "center" },
  sendAllBtn: {
    display: "flex",
    alignItems: "center",
    gap: "10px",
    background: "linear-gradient(135deg, #25D366, #128c5e)",
    border: "none",
    borderRadius: "12px",
    color: "#fff",
    fontWeight: "800",
    fontSize: "16px",
    padding: "14px 36px",
    cursor: "pointer",
    boxShadow: "0 4px 20px #25d36640",
    transition: "all 0.2s",
  },
  tableWrap: {
    overflowX: "auto",
    border: "1px solid #1e293b",
    borderRadius: "12px",
  },
  table: { width: "100%", borderCollapse: "collapse" },
  th: {
    background: "#0d2b1e",
    color: "#6ee7b7",
    padding: "10px 14px",
    fontSize: "11px",
    fontWeight: "700",
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    textAlign: "left",
    borderBottom: "1px solid #134e4a",
    whiteSpace: "nowrap",
  },
  td: {
    padding: "10px 14px",
    fontSize: "13px",
    color: "#cbd5e1",
    borderBottom: "1px solid #1e293b",
    maxWidth: "200px",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  statusBadge: {
    display: "inline-block",
    padding: "3px 10px",
    borderRadius: "20px",
    fontSize: "11px",
    fontWeight: "700",
  },
  sendBtn: {
    background: "#134e4a",
    border: "1px solid #25D366",
    color: "#25D366",
    borderRadius: "6px",
    padding: "5px 12px",
    cursor: "pointer",
    fontSize: "12px",
    fontWeight: "600",
    transition: "all 0.15s",
  },
  footer: {
    background: "#0a0f1a",
    borderTop: "1px solid #1e293b",
    color: "#475569",
    fontSize: "12px",
    padding: "14px 24px",
    textAlign: "center",
  },
};
