import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";

// ── THEME ──────────────────────────────────────────────────────────────────
const T = {
  bg: "#0d0d0f",
  surface: "#141418",
  surface2: "#1c1c22",
  border: "#2a2a35",
  text: "#e8e6f0",
  muted: "#6b6880",
  accent: "#f0a500",
  red: "#e05252",
  green: "#3ecf8e",
  blue: "#4d9fff",
  purple: "#a78bfa",
};

const css = (obj) => Object.entries(obj).map(([k, v]) => `${k.replace(/([A-Z])/g, "-$1").toLowerCase()}:${v}`).join(";");

// ── HELPERS ────────────────────────────────────────────────────────────────
function parseFile(file) {
  return new Promise((resolve, reject) => {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      Papa.parse(file, {
        header: true, skipEmptyLines: true,
        complete: (r) => resolve(r.data),
        error: reject,
      });
    } else {
      const reader = new FileReader();
      reader.onload = (e) => {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        resolve(XLSX.utils.sheet_to_json(ws, { defval: "" }));
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    }
  });
}

function normalise(s) {
  return String(s || "").toLowerCase().trim();
}

function pct(n) {
  if (!isFinite(n) || n <= 0) return "—";
  return (n * 100).toFixed(1) + "%";
}

function money(n) {
  return "$" + Number(n || 0).toFixed(2);
}

// ── ANALYSIS ENGINE ────────────────────────────────────────────────────────
function analyseReports({ stRows, tgRows }) {
  // Detect column names flexibly
  const stSample = stRows[0] || {};
  const tgSample = tgRows[0] || {};

  const stTermCol = Object.keys(stSample).find(k => /search.?term|customer.?search/i.test(k)) || "Customer Search Term";
  const stSpendCol = Object.keys(stSample).find(k => /^spend$|total.?cost/i.test(k)) || "Spend";
  const stOrderCol = Object.keys(stSample).find(k => /orders|purchases/i.test(k) && !/halo|new.to|promo/i.test(k)) || "7 Day Total Orders (#)";
  const stSalesCol = Object.keys(stSample).find(k => /total.?sales|^sales$/i.test(k) && !/halo|new.to|promo/i.test(k)) || "7 Day Total Sales ($)";
  const stClickCol = Object.keys(stSample).find(k => /^clicks$/i.test(k)) || "Clicks";
  const stCampCol  = Object.keys(stSample).find(k => /campaign.?name/i.test(k)) || "Campaign Name";
  const stMatchCol = Object.keys(stSample).find(k => /match.?type/i.test(k)) || "Match Type";

  const tgKwCol    = Object.keys(tgSample).find(k => /^targeting$/i.test(k)) || "Targeting";
  const tgSpendCol = Object.keys(tgSample).find(k => /^spend$|total.?cost/i.test(k)) || "Spend";
  const tgOrderCol = Object.keys(tgSample).find(k => /orders|purchases/i.test(k) && !/halo|new.to|promo/i.test(k)) || "7 Day Total Orders (#)";
  const tgSalesCol = Object.keys(tgSample).find(k => /total.?sales|^sales$/i.test(k) && !/halo|new.to|promo/i.test(k)) || "7 Day Total Sales ($)";
  const tgClickCol = Object.keys(tgSample).find(k => /^clicks$/i.test(k)) || "Clicks";
  const tgMatchCol = Object.keys(tgSample).find(k => /match.?type/i.test(k)) || "Match Type";
  const tgCampCol  = Object.keys(tgSample).find(k => /campaign.?name/i.test(k)) || "Campaign Name";
  const tgCpcCol   = Object.keys(tgSample).find(k => /cpc|cost.per.click/i.test(k)) || "Cost per click (CPC)";

  // ── Build existing keyword set (all match types)
  const existingKws = new Set(tgRows.map(r => normalise(r[tgKwCol])));
  const existingExact = new Set(tgRows.filter(r => normalise(r[tgMatchCol]) === "exact").map(r => normalise(r[tgKwCol])));

  // ── Aggregate search terms
  const stMap = {};
  for (const r of stRows) {
    const term = normalise(r[stTermCol]);
    if (!term || term.length < 4) continue;
    if (/^b0[a-z0-9]{8}$/.test(term)) continue; // skip ASIN targets
    if (!stMap[term]) stMap[term] = { term, spend: 0, orders: 0, sales: 0, clicks: 0, campaigns: new Set(), matchTypes: new Set() };
    stMap[term].spend   += parseFloat(r[stSpendCol]) || 0;
    stMap[term].orders  += parseFloat(r[stOrderCol]) || 0;
    stMap[term].sales   += parseFloat(r[stSalesCol]) || 0;
    stMap[term].clicks  += parseFloat(r[stClickCol]) || 0;
    if (r[stCampCol]) stMap[term].campaigns.add(normalise(r[stCampCol]));
    if (r[stMatchCol]) stMap[term].matchTypes.add(normalise(r[stMatchCol]));
  }

  const allTerms = Object.values(stMap).map(t => ({
    ...t,
    campaigns: [...t.campaigns],
    matchTypes: [...t.matchTypes],
    acos: t.sales > 0 ? t.spend / t.sales : null,
    cvr: t.clicks > 0 ? t.orders / t.clicks : null,
    inExisting: existingKws.has(t.term),
    inExact: existingExact.has(t.term),
  }));

  // ── SECTION 1: Wasted spend (zero orders, significant spend)
  const wastedTerms = allTerms
    .filter(t => t.orders === 0 && t.spend >= 15)
    .sort((a, b) => b.spend - a.spend);

  // ── SECTION 2: Duplicate converting terms (same term, 2+ campaigns)
  const duplicates = allTerms
    .filter(t => t.orders > 0 && t.campaigns.length >= 2)
    .sort((a, b) => b.spend - a.spend);

  // ── SECTION 3: Harvest candidates (converting in non-exact, not in exact)
  const harvestCandidates = allTerms
    .filter(t => t.orders >= 2 && !t.inExact && t.acos !== null && t.acos <= 0.55)
    .filter(t => !t.matchTypes.every(m => m === "exact"))
    .sort((a, b) => a.acos - b.acos);

  // ── SECTION 4: New terms (converting, not targeted at all)
  const newTerms = allTerms
    .filter(t => t.orders >= 2 && !t.inExisting && t.acos !== null && t.acos <= 0.50)
    .sort((a, b) => b.orders - a.orders);

  // ── SECTION 5: Keywords to pause from targeting report
  const tgMap = {};
  for (const r of tgRows) {
    const kw = normalise(r[tgKwCol]);
    if (!kw || kw.length < 3) continue;
    if (!tgMap[kw]) tgMap[kw] = { kw, spend: 0, orders: 0, sales: 0, clicks: 0, matchType: normalise(r[tgMatchCol]), campaigns: new Set(), cpc: 0 };
    tgMap[kw].spend   += parseFloat(r[tgSpendCol]) || 0;
    tgMap[kw].orders  += parseFloat(r[tgOrderCol]) || 0;
    tgMap[kw].sales   += parseFloat(r[tgSalesCol]) || 0;
    tgMap[kw].clicks  += parseFloat(r[tgClickCol]) || 0;
    tgMap[kw].cpc      = parseFloat(r[tgCpcCol]) || tgMap[kw].cpc;
    if (r[tgCampCol]) tgMap[kw].campaigns.add(r[tgCampCol]);
  }

  const pauseList = Object.values(tgMap)
    .filter(t => t.orders === 0 && t.spend >= 20)
    .map(t => ({ ...t, campaigns: [...t.campaigns] }))
    .sort((a, b) => b.spend - a.spend);

  // ── SECTION 6: Scale candidates (low ACoS, converting, already in exact)
  const scaleCandidates = Object.values(tgMap)
    .filter(t => {
      const acos = t.sales > 0 ? t.spend / t.sales : null;
      return t.orders >= 5 && acos !== null && acos <= 0.30 && t.spend > 20;
    })
    .map(t => ({ ...t, acos: t.spend / t.sales, campaigns: [...t.campaigns] }))
    .sort((a, b) => a.acos - b.acos);

  // ── SUMMARY
  const totalSpend = allTerms.reduce((s, t) => s + t.spend, 0);
  const totalSales = allTerms.reduce((s, t) => s + t.sales, 0);
  const totalOrders = allTerms.reduce((s, t) => s + t.orders, 0);
  const wastedSpend = wastedTerms.reduce((s, t) => s + t.spend, 0);
  const overallAcos = totalSales > 0 ? totalSpend / totalSales : null;

  return {
    summary: { totalSpend, totalSales, totalOrders, wastedSpend, overallAcos, termCount: allTerms.length },
    wastedTerms,
    duplicates,
    harvestCandidates,
    newTerms,
    pauseList,
    scaleCandidates,
  };
}

// ── AI ANALYSIS ────────────────────────────────────────────────────────────
async function askClaude(prompt) {
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      model: "claude-sonnet-4-20250514",
      max_tokens: 1000,
      messages: [{ role: "user", content: prompt }],
    }),
  });
  const data = await res.json();
  return data.content?.find(b => b.type === "text")?.text || "No response";
}

// ── COMPONENTS ─────────────────────────────────────────────────────────────
function Tag({ color, children }) {
  const colors = {
    red:    { bg: "rgba(224,82,82,0.12)",   color: T.red,    border: "rgba(224,82,82,0.3)" },
    green:  { bg: "rgba(62,207,142,0.12)",  color: T.green,  border: "rgba(62,207,142,0.3)" },
    orange: { bg: "rgba(240,165,0,0.12)",   color: T.accent, border: "rgba(240,165,0,0.3)" },
    blue:   { bg: "rgba(77,159,255,0.12)",  color: T.blue,   border: "rgba(77,159,255,0.3)" },
    purple: { bg: "rgba(167,139,250,0.12)", color: T.purple, border: "rgba(167,139,250,0.3)" },
  };
  const c = colors[color] || colors.blue;
  return (
    <span style={{ display:"inline-block", fontSize:9, fontFamily:"monospace", letterSpacing:"0.12em",
      textTransform:"uppercase", padding:"2px 7px", borderRadius:2, fontWeight:600,
      background:c.bg, color:c.color, border:`1px solid ${c.border}` }}>
      {children}
    </span>
  );
}

function AcosCell({ v }) {
  if (v === null || !isFinite(v)) return <span style={{ color: T.muted }}>—</span>;
  const pv = v * 100;
  const color = pv <= 30 ? T.green : pv <= 50 ? T.accent : T.red;
  return <span style={{ color, fontWeight: 600 }}>{pv.toFixed(1)}%</span>;
}

function SectionHeader({ num, title, count, color = T.accent }) {
  return (
    <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16,
      paddingBottom:10, borderBottom:`1px solid ${T.border}` }}>
      <span style={{ fontFamily:"monospace", fontSize:10, background:color,
        color:"#000", padding:"2px 8px", fontWeight:700, letterSpacing:"0.1em" }}>
        {num}
      </span>
      <span style={{ fontSize:16, fontWeight:700, color:T.text, letterSpacing:"0.02em" }}>{title}</span>
      {count !== undefined && (
        <span style={{ marginLeft:"auto", fontFamily:"monospace", fontSize:10,
          color:color, border:`1px solid ${color}`, padding:"2px 8px", borderRadius:2,
          background:`${color}18` }}>
          {count} items
        </span>
      )}
    </div>
  );
}

function DataTable({ cols, rows, maxRows = 50 }) {
  const [show, setShow] = useState(maxRows);
  if (!rows.length) return <p style={{ color:T.muted, fontSize:12, fontStyle:"italic" }}>No items found.</p>;
  return (
    <>
      <div style={{ overflowX:"auto" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
          <thead>
            <tr style={{ borderBottom:`1px solid ${T.border}` }}>
              {cols.map(c => (
                <th key={c.key} style={{ textAlign:"left", padding:"8px 10px", fontFamily:"monospace",
                  fontSize:9, letterSpacing:"0.12em", textTransform:"uppercase", color:T.muted,
                  background:T.surface2, whiteSpace:"nowrap" }}>
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.slice(0, show).map((row, i) => (
              <tr key={i} style={{ borderBottom:`1px solid ${T.border}22` }}
                onMouseEnter={e => e.currentTarget.style.background = T.surface2}
                onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                {cols.map(c => (
                  <td key={c.key} style={{ padding:"7px 10px", color:T.text, verticalAlign:"top" }}>
                    {c.render ? c.render(row) : row[c.key]}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {rows.length > show && (
        <button onClick={() => setShow(s => s + 50)}
          style={{ marginTop:10, padding:"6px 16px", background:"transparent",
            border:`1px solid ${T.border}`, color:T.muted, cursor:"pointer",
            fontSize:11, fontFamily:"monospace", borderRadius:2 }}>
          Show more ({rows.length - show} remaining)
        </button>
      )}
    </>
  );
}

function AiPanel({ results }) {
  const [question, setQuestion] = useState("");
  const [answer, setAnswer] = useState("");
  const [loading, setLoading] = useState(false);

  const ask = async () => {
    if (!question.trim()) return;
    setLoading(true);
    setAnswer("");
    const context = `
You are an Amazon Advertising expert. Here is the account data summary:
- Total spend: $${results.summary.totalSpend.toFixed(2)}
- Total sales: $${results.summary.totalSales.toFixed(2)}
- Overall ACoS: ${results.summary.overallAcos ? (results.summary.overallAcos * 100).toFixed(1) + "%" : "N/A"}
- Wasted spend (zero-order terms): $${results.summary.wastedSpend.toFixed(2)}
- Keywords to pause: ${results.pauseList.length}
- Duplicated converting terms: ${results.duplicates.length}
- Harvest candidates (not in exact): ${results.harvestCandidates.length}
- New untargeted converting terms: ${results.newTerms.length}
- Scale candidates (sub-30% ACoS): ${results.scaleCandidates.length}

Top wasted terms: ${results.wastedTerms.slice(0, 5).map(t => `"${t.term}" ($${t.spend.toFixed(2)})`).join(", ")}
Top harvest candidates: ${results.harvestCandidates.slice(0, 5).map(t => `"${t.term}" (${(t.acos * 100).toFixed(1)}% ACoS, ${t.orders} orders)`).join(", ")}
Top new terms: ${results.newTerms.slice(0, 5).map(t => `"${t.term}" (${t.orders} orders, ${(t.acos * 100).toFixed(1)}% ACoS)`).join(", ")}
Top duplicated terms: ${results.duplicates.slice(0, 5).map(t => `"${t.term}" (${t.campaigns.length} campaigns, $${t.spend.toFixed(2)} spend)`).join(", ")}

User question: ${question}

Give a direct, specific, actionable answer. No fluff. Use concrete numbers from the data above.
    `.trim();
    try {
      const ans = await askClaude(context);
      setAnswer(ans);
    } catch (e) {
      setAnswer("Error: " + e.message);
    }
    setLoading(false);
  };

  return (
    <div style={{ background:T.surface, border:`1px solid ${T.border}`, borderRadius:4, padding:20, marginTop:24 }}>
      <div style={{ fontSize:13, fontWeight:700, color:T.accent, marginBottom:12,
        fontFamily:"monospace", letterSpacing:"0.1em", textTransform:"uppercase" }}>
        ⚡ Ask About This Data
      </div>
      <div style={{ display:"flex", gap:8 }}>
        <input
          value={question}
          onChange={e => setQuestion(e.target.value)}
          onKeyDown={e => e.key === "Enter" && ask()}
          placeholder="e.g. Which keywords should I pause first? What new campaigns should I build?"
          style={{ flex:1, background:T.surface2, border:`1px solid ${T.border}`, borderRadius:3,
            padding:"9px 14px", color:T.text, fontSize:12, fontFamily:"monospace", outline:"none" }}
        />
        <button onClick={ask} disabled={loading}
          style={{ padding:"9px 20px", background:T.accent, color:"#000", border:"none",
            borderRadius:3, cursor:"pointer", fontWeight:700, fontSize:12, fontFamily:"monospace",
            opacity: loading ? 0.6 : 1 }}>
          {loading ? "..." : "ASK"}
        </button>
      </div>
      {answer && (
        <div style={{ marginTop:14, padding:"14px 16px", background:T.surface2,
          borderRadius:3, fontSize:12, color:T.text, lineHeight:1.7,
          borderLeft:`3px solid ${T.accent}`, whiteSpace:"pre-wrap" }}>
          {answer}
        </div>
      )}
    </div>
  );
}

// ── UPLOAD ZONE ────────────────────────────────────────────────────────────
function UploadZone({ label, hint, onFile, uploaded, color = T.blue }) {
  const ref = useRef();
  const [drag, setDrag] = useState(false);

  const handle = (file) => { if (file) onFile(file); };

  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true); }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]); }}
      style={{ border:`2px dashed ${drag ? color : uploaded ? T.green : T.border}`,
        borderRadius:6, padding:"20px 16px", cursor:"pointer", textAlign:"center",
        background: uploaded ? "rgba(62,207,142,0.05)" : drag ? `${color}10` : T.surface,
        transition:"all 0.2s" }}>
      <input ref={ref} type="file" accept=".xlsx,.xls,.csv"
        style={{ display:"none" }} onChange={e => handle(e.target.files[0])} />
      <div style={{ fontSize:22, marginBottom:6 }}>{uploaded ? "✅" : "📂"}</div>
      <div style={{ fontSize:13, fontWeight:600, color: uploaded ? T.green : T.text, marginBottom:4 }}>
        {uploaded ? uploaded : label}
      </div>
      <div style={{ fontSize:11, color:T.muted }}>{hint}</div>
    </div>
  );
}

// ── MAIN APP ───────────────────────────────────────────────────────────────
export default function App() {
  const [stFile, setStFile] = useState(null);
  const [tgFile, setTgFile] = useState(null);
  const [results, setResults] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("waste");
  const [acosTarget, setAcosTarget] = useState(30);

  const analyse = useCallback(async () => {
    if (!stFile || !tgFile) return;
    setLoading(true);
    setError("");
    try {
      const [stRows, tgRows] = await Promise.all([parseFile(stFile), parseFile(tgFile)]);
      const r = analyseReports({ stRows, tgRows });
      setResults(r);
      setActiveTab("waste");
    } catch (e) {
      setError("Parse error: " + e.message);
    }
    setLoading(false);
  }, [stFile, tgFile]);

  const tabs = [
    { id:"waste",    label:"Wasted Spend",    color:T.red },
    { id:"pause",    label:"Pause Keywords",  color:T.red },
    { id:"dupe",     label:"Duplicates",      color:T.accent },
    { id:"harvest",  label:"Harvest",         color:T.green },
    { id:"new",      label:"New Campaigns",   color:T.blue },
    { id:"scale",    label:"Scale Bids",      color:T.purple },
  ];

  // ── RENDER ────────────────────────────────────────────────────────────────
  return (
    <div style={{ minHeight:"100vh", background:T.bg, color:T.text,
      fontFamily:"'DM Sans', system-ui, sans-serif" }}>

      {/* HEADER */}
      <div style={{ background:T.surface, borderBottom:`1px solid ${T.border}`,
        padding:"20px 32px", display:"flex", alignItems:"center", gap:16 }}>
        <div style={{ width:8, height:8, background:T.accent, borderRadius:"50%" }} />
        <span style={{ fontFamily:"monospace", fontSize:13, fontWeight:700,
          letterSpacing:"0.15em", textTransform:"uppercase", color:T.text }}>
          Amazon Ads Audit Tool
        </span>
        <span style={{ marginLeft:"auto", fontFamily:"monospace", fontSize:10,
          color:T.muted, letterSpacing:"0.1em" }}>
          ACTIVE GREEN PRO · CAD
        </span>
      </div>

      {/* UPLOAD PANEL */}
      {!results && (
        <div style={{ maxWidth:760, margin:"60px auto", padding:"0 24px" }}>
          <div style={{ textAlign:"center", marginBottom:40 }}>
            <div style={{ fontFamily:"monospace", fontSize:11, color:T.accent,
              letterSpacing:"0.2em", textTransform:"uppercase", marginBottom:12 }}>
              Upload Reports to Analyse
            </div>
            <div style={{ fontSize:28, fontWeight:800, color:T.text, lineHeight:1.1, marginBottom:8 }}>
              Drop your Amazon reports.<br />Get your action list.
            </div>
            <div style={{ fontSize:13, color:T.muted }}>
              Supports .xlsx and .csv · All analysis runs in-browser
            </div>
          </div>

          <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:20 }}>
            <UploadZone
              label="Search Term Report"
              hint="65-day or 2-year STR (.xlsx or .csv)"
              onFile={f => setStFile(f)}
              uploaded={stFile?.name}
              color={T.blue}
            />
            <UploadZone
              label="Targeting Report"
              hint="SP Targeting report (.xlsx)"
              onFile={f => setTgFile(f)}
              uploaded={tgFile?.name}
              color={T.green}
            />
          </div>

          <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:20,
            background:T.surface, border:`1px solid ${T.border}`, borderRadius:4,
            padding:"12px 16px" }}>
            <span style={{ fontSize:12, color:T.muted, fontFamily:"monospace", letterSpacing:"0.08em" }}>
              TARGET ACoS %
            </span>
            <input type="number" value={acosTarget} min={10} max={80}
              onChange={e => setAcosTarget(Number(e.target.value))}
              style={{ width:64, background:T.surface2, border:`1px solid ${T.border}`,
                borderRadius:3, padding:"5px 8px", color:T.text, fontSize:13,
                fontFamily:"monospace", textAlign:"center", outline:"none" }}
            />
            <span style={{ fontSize:11, color:T.muted }}>Used to colour-code ACoS values throughout</span>
          </div>

          {error && (
            <div style={{ background:"rgba(224,82,82,0.1)", border:`1px solid ${T.red}`,
              borderRadius:4, padding:"10px 14px", fontSize:12, color:T.red, marginBottom:16 }}>
              {error}
            </div>
          )}

          <button onClick={analyse} disabled={!stFile || !tgFile || loading}
            style={{ width:"100%", padding:"14px", background: (!stFile || !tgFile) ? T.surface2 : T.accent,
              color: (!stFile || !tgFile) ? T.muted : "#000", border:"none", borderRadius:4,
              cursor: (!stFile || !tgFile) ? "not-allowed" : "pointer", fontWeight:800,
              fontSize:14, fontFamily:"monospace", letterSpacing:"0.1em",
              textTransform:"uppercase", transition:"all 0.2s" }}>
            {loading ? "Analysing..." : "Run Audit →"}
          </button>

          <div style={{ marginTop:24, padding:"14px 16px", background:T.surface,
            border:`1px solid ${T.border}`, borderRadius:4, fontSize:11, color:T.muted,
            lineHeight:1.7 }}>
            <strong style={{ color:T.text }}>What this tool checks:</strong><br />
            Zero-order wasted spend · Keywords to pause · Duplicate terms across campaigns ·
            Converting terms not in exact match · New untargeted converters · Under-bid high performers
          </div>
        </div>
      )}

      {/* RESULTS */}
      {results && (
        <div style={{ maxWidth:1200, margin:"0 auto", padding:"24px 32px" }}>

          {/* SUMMARY BAR */}
          <div style={{ display:"grid", gridTemplateColumns:"repeat(5,1fr)",
            gap:1, background:T.border, borderRadius:4, overflow:"hidden",
            marginBottom:28, border:`1px solid ${T.border}` }}>
            {[
              { label:"Total Spend",    val:money(results.summary.totalSpend),   color:T.accent },
              { label:"Total Sales",    val:money(results.summary.totalSales),   color:T.green },
              { label:"Overall ACoS",  val:pct(results.summary.overallAcos),    color:T.red },
              { label:"Wasted Spend",  val:money(results.summary.wastedSpend),  color:T.red },
              { label:"Action Items",  val: results.pauseList.length + results.duplicates.length + results.harvestCandidates.length, color:T.blue },
            ].map(k => (
              <div key={k.label} style={{ background:T.surface, padding:"16px 20px" }}>
                <div style={{ fontFamily:"monospace", fontSize:9, letterSpacing:"0.14em",
                  textTransform:"uppercase", color:T.muted, marginBottom:4 }}>{k.label}</div>
                <div style={{ fontFamily:"monospace", fontSize:22, fontWeight:800,
                  color:k.color, lineHeight:1 }}>{k.val}</div>
              </div>
            ))}
          </div>

          {/* TABS */}
          <div style={{ display:"flex", gap:4, marginBottom:20, flexWrap:"wrap" }}>
            {tabs.map(tab => {
              const counts = {
                waste:   results.wastedTerms.length,
                pause:   results.pauseList.length,
                dupe:    results.duplicates.length,
                harvest: results.harvestCandidates.length,
                new:     results.newTerms.length,
                scale:   results.scaleCandidates.length,
              };
              const active = activeTab === tab.id;
              return (
                <button key={tab.id} onClick={() => setActiveTab(tab.id)}
                  style={{ padding:"7px 14px", border:`1px solid ${active ? tab.color : T.border}`,
                    background: active ? `${tab.color}18` : T.surface,
                    color: active ? tab.color : T.muted, cursor:"pointer",
                    fontSize:11, fontFamily:"monospace", borderRadius:3,
                    fontWeight: active ? 700 : 400, letterSpacing:"0.08em",
                    transition:"all 0.15s" }}>
                  {tab.label}
                  <span style={{ marginLeft:6, background: active ? tab.color : T.surface2,
                    color: active ? "#000" : T.muted, padding:"1px 6px",
                    borderRadius:10, fontSize:9, fontWeight:700 }}>
                    {counts[tab.id]}
                  </span>
                </button>
              );
            })}
            <button onClick={() => { setResults(null); setStFile(null); setTgFile(null); }}
              style={{ marginLeft:"auto", padding:"7px 14px", border:`1px solid ${T.border}`,
                background:"transparent", color:T.muted, cursor:"pointer",
                fontSize:11, fontFamily:"monospace", borderRadius:3 }}>
              ← New Upload
            </button>
          </div>

          {/* TAB CONTENT */}
          <div style={{ background:T.surface, border:`1px solid ${T.border}`,
            borderRadius:4, padding:24 }}>

            {activeTab === "waste" && (
              <>
                <SectionHeader num="01" title="Wasted Spend — Zero-Order Search Terms"
                  count={results.wastedTerms.length} color={T.red} />
                <p style={{ fontSize:12, color:T.muted, marginBottom:14 }}>
                  Search terms with $15+ spend and zero orders. Add as negatives immediately.
                  Terms currently in Exact match also need the keyword paused.
                </p>
                <DataTable
                  rows={results.wastedTerms}
                  cols={[
                    { key:"term",     label:"Search Term",  render:r => <span style={{ color:T.text }}>{r.term}</span> },
                    { key:"spend",    label:"Spend",        render:r => <span style={{ color:T.red, fontWeight:600 }}>{money(r.spend)}</span> },
                    { key:"clicks",   label:"Clicks",       render:r => r.clicks },
                    { key:"matchTypes",label:"Match Types", render:r => r.matchTypes.map((m,i) => <Tag key={i} color={m==="exact"?"red":m==="broad"?"orange":"blue"}>{m||"auto"}</Tag>) },
                    { key:"campaigns",label:"Campaigns",    render:r => <span style={{ fontSize:11, color:T.muted }}>{r.campaigns.length} campaign{r.campaigns.length>1?"s":""}</span> },
                    { key:"action",   label:"Action",       render:r => r.matchTypes.includes("exact") ? <Tag color="red">PAUSE + NEGATE</Tag> : <Tag color="orange">NEGATE</Tag> },
                  ]}
                />
              </>
            )}

            {activeTab === "pause" && (
              <>
                <SectionHeader num="02" title="Keywords to Pause — From Targeting Report"
                  count={results.pauseList.length} color={T.red} />
                <p style={{ fontSize:12, color:T.muted, marginBottom:14 }}>
                  Active keyword targets with $20+ spend and zero orders.
                  These are deliberate bids that haven't converted — pause them.
                </p>
                <DataTable
                  rows={results.pauseList}
                  cols={[
                    { key:"kw",       label:"Keyword",     render:r => <span style={{ color:T.text }}>{r.kw}</span> },
                    { key:"spend",    label:"Spend",       render:r => <span style={{ color:T.red, fontWeight:600 }}>{money(r.spend)}</span> },
                    { key:"clicks",   label:"Clicks",      render:r => r.clicks },
                    { key:"matchType",label:"Match",       render:r => <Tag color={r.matchType==="exact"?"red":r.matchType==="broad"?"orange":"blue"}>{r.matchType||"—"}</Tag> },
                    { key:"cpc",      label:"CPC",         render:r => money(r.cpc) },
                    { key:"camp",     label:"Campaigns",   render:r => <span style={{ fontSize:11, color:T.muted }}>{r.campaigns.length} campaign{r.campaigns.length>1?"s":""}</span> },
                    { key:"action",   label:"Action",      render:() => <Tag color="red">PAUSE</Tag> },
                  ]}
                />
              </>
            )}

            {activeTab === "dupe" && (
              <>
                <SectionHeader num="03" title="Duplicate Converting Terms — Self-Competing Campaigns"
                  count={results.duplicates.length} color={T.accent} />
                <p style={{ fontSize:12, color:T.muted, marginBottom:14 }}>
                  Converting search terms triggering in 2+ campaigns simultaneously.
                  You're bidding against yourself, inflating CPCs.
                  Keep each term in the best-performing campaign only — negate from all others.
                </p>
                <DataTable
                  rows={results.duplicates}
                  cols={[
                    { key:"term",     label:"Search Term",   render:r => <span style={{ color:T.text }}>{r.term}</span> },
                    { key:"camps",    label:"# Campaigns",   render:r => <span style={{ color:T.red, fontWeight:600 }}>{r.campaigns.length}</span> },
                    { key:"spend",    label:"Total Spend",   render:r => <span style={{ color:T.accent, fontWeight:600 }}>{money(r.spend)}</span> },
                    { key:"orders",   label:"Orders",        render:r => <span style={{ color:T.green }}>{r.orders}</span> },
                    { key:"acos",     label:"Blended ACoS",  render:r => <AcosCell v={r.acos} /> },
                    { key:"action",   label:"Fix",           render:() => <Tag color="orange">CONSOLIDATE</Tag> },
                  ]}
                />
              </>
            )}

            {activeTab === "harvest" && (
              <>
                <SectionHeader num="04" title="Harvest Candidates — Move to Exact Match"
                  count={results.harvestCandidates.length} color={T.green} />
                <p style={{ fontSize:12, color:T.muted, marginBottom:14 }}>
                  Converting terms currently in Auto/Broad/Phrase that are NOT in exact match.
                  Add as Exact targets in a controlled campaign, then negate from the originating campaign.
                </p>
                <DataTable
                  rows={results.harvestCandidates}
                  cols={[
                    { key:"term",     label:"Search Term",   render:r => <span style={{ color:T.text }}>{r.term}</span> },
                    { key:"orders",   label:"Orders",        render:r => <span style={{ color:T.green, fontWeight:600 }}>{r.orders}</span> },
                    { key:"sales",    label:"Sales",         render:r => money(r.sales) },
                    { key:"spend",    label:"Spend",         render:r => money(r.spend) },
                    { key:"acos",     label:"ACoS",          render:r => <AcosCell v={r.acos} /> },
                    { key:"cvr",      label:"CVR",           render:r => r.cvr ? (r.cvr*100).toFixed(1)+"%" : "—" },
                    { key:"matchTypes",label:"Current Match",render:r => r.matchTypes.map((m,i)=><Tag key={i} color="orange">{m||"auto"}</Tag>) },
                    { key:"action",   label:"Action",        render:() => <Tag color="green">ADD EXACT</Tag> },
                  ]}
                />
              </>
            )}

            {activeTab === "new" && (
              <>
                <SectionHeader num="05" title="New Campaign Keywords — Not Targeted Anywhere"
                  count={results.newTerms.length} color={T.blue} />
                <p style={{ fontSize:12, color:T.muted, marginBottom:14 }}>
                  Converting search terms with 2+ orders that are NOT in any current campaign target.
                  These are proven buyers you're getting by accident. Build exact match campaigns around them.
                </p>
                <DataTable
                  rows={results.newTerms}
                  cols={[
                    { key:"term",   label:"Search Term",  render:r => <span style={{ color:T.text }}>{r.term}</span> },
                    { key:"orders", label:"2-Yr Orders",  render:r => <span style={{ color:T.green, fontWeight:600 }}>{r.orders}</span> },
                    { key:"sales",  label:"Sales",        render:r => money(r.sales) },
                    { key:"spend",  label:"Spend",        render:r => money(r.spend) },
                    { key:"acos",   label:"ACoS",         render:r => <AcosCell v={r.acos} /> },
                    { key:"cvr",    label:"CVR",          render:r => r.cvr ? (r.cvr*100).toFixed(1)+"%" : "—" },
                    { key:"action", label:"Action",       render:() => <Tag color="blue">NEW EXACT CAMPAIGN</Tag> },
                  ]}
                />
              </>
            )}

            {activeTab === "scale" && (
              <>
                <SectionHeader num="06" title="Scale Bids — High Performers Under-Bidding"
                  count={results.scaleCandidates.length} color={T.purple} />
                <p style={{ fontSize:12, color:T.muted, marginBottom:14 }}>
                  Keywords with 5+ orders, ACoS under 30%, and meaningful spend.
                  These are converting efficiently — raise bids 20–40% to capture more impression share.
                </p>
                <DataTable
                  rows={results.scaleCandidates}
                  cols={[
                    { key:"kw",     label:"Keyword",      render:r => <span style={{ color:T.text }}>{r.kw}</span> },
                    { key:"orders", label:"Orders",       render:r => <span style={{ color:T.green, fontWeight:600 }}>{r.orders}</span> },
                    { key:"acos",   label:"ACoS",         render:r => <AcosCell v={r.acos} /> },
                    { key:"spend",  label:"Spend",        render:r => money(r.spend) },
                    { key:"cpc",    label:"Current CPC",  render:r => money(r.cpc) },
                    { key:"sugBid", label:"Suggested Bid",render:r => <span style={{ color:T.purple, fontWeight:600 }}>{money(r.cpc * 1.3)}</span> },
                    { key:"action", label:"Action",       render:() => <Tag color="purple">RAISE BID +30%</Tag> },
                  ]}
                />
              </>
            )}

          </div>

          {/* AI PANEL */}
          <AiPanel results={results} />

        </div>
      )}
    </div>
  );
}
