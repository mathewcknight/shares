import React, { useState, useMemo } from "react";
import { Plus, X, ScanLine, Check, ChevronDown, ChevronRight } from "lucide-react";

// ---- scoring model -------------------------------------------------------
const CYCLE = ["", "−", "NA", "+", "++"];
const WEIGHT = { "−": -2, NA: 0, "+": 1, "++": 2 };
const GLYPH = { "": "·", "−": "−", NA: "n/a", "+": "+", "++": "++" };

let _id = 1;
const uid = () => `x${_id++}`;

// ---- inline markdown renderer (bold / code / links / italic) -------------
function Inline({ text, C }) {
  const nodes = [];
  const re = /(\*\*[^*]+\*\*)|(`[^`]+`)|(\[[^\]]+\]\([^)]+\))|(\*[^*\s][^*]*\*)/g;
  let last = 0, m, k = 0;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) nodes.push(text.slice(last, m.index));
    const tok = m[0];
    if (tok.startsWith("**")) {
      nodes.push(<strong key={k++} style={{ color: C.paper, fontWeight: 700 }}>{tok.slice(2, -2)}</strong>);
    } else if (tok.startsWith("`")) {
      nodes.push(<code key={k++} style={{ background: "#0a0b09", border: `1px solid ${C.line}`, borderRadius: 4, padding: "1px 5px", fontSize: "0.9em", color: C.amber }}>{tok.slice(1, -1)}</code>);
    } else if (tok.startsWith("[")) {
      const mm = tok.match(/\[([^\]]+)\]\(([^)]+)\)/);
      nodes.push(<a key={k++} href={mm[2]} target="_blank" rel="noreferrer" style={{ color: C.amber, textDecoration: "underline" }}>{mm[1]}</a>);
    } else {
      nodes.push(<em key={k++} style={{ color: C.paper }}>{tok.slice(1, -1)}</em>);
    }
    last = m.index + tok.length;
  }
  if (last < text.length) nodes.push(text.slice(last));
  return <>{nodes}</>;
}

// ---- parser: keep raw markdown + section heading -------------------------
function stripMd(s) {
  return s.replace(/\*\*(.*?)\*\*/g, "$1").replace(/`(.*?)`/g, "$1").replace(/\*(.*?)\*/g, "$1").replace(/#+\s*/g, "").trim();
}
function extractItems(raw) {
  if (!raw) return [];
  const out = [];
  let section = "";
  raw.split(/\r?\n/).forEach((line) => {
    const h = line.match(/^\s{0,3}#{1,6}\s+(.+?)\s*$/);
    const boldHead = line.match(/^\s*\*\*(.+?)\*\*:?\s*$/);
    if (h) { section = stripMd(h[1]); return; }
    if (boldHead) { section = stripMd(boldHead[1]); return; }
    const b = line.match(/^\s*(?:[-•*]|\d+[.)])\s+(.*\S)\s*$/);
    if (!b) return;
    const rawText = b[1].replace(/⚠️/g, "").trim();
    if (stripMd(rawText).length < 4) return;
    out.push({ id: uid(), raw: rawText, section, checked: false });
  });
  return out;
}

export default function ACHWorkbench() {
  const [raw, setRaw] = useState("");
  const [items, setItems] = useState([]);
  const [hyps, setHyps] = useState([]);
  const [scores, setScores] = useState({});
  const [newHyp, setNewHyp] = useState("");
  const [newItem, setNewItem] = useState("");
  const [collapsed, setCollapsed] = useState({});
  const [showText, setShowText] = useState(true);

  const parse = () => {
    const found = extractItems(raw);
    setItems((prev) => {
      const have = new Set(prev.map((i) => stripMd(i.raw).toLowerCase()));
      const merged = [...prev];
      found.forEach((f) => {
        const key = stripMd(f.raw).toLowerCase();
        if (!have.has(key)) { merged.push(f); have.add(key); }
      });
      return merged;
    });
  };

  const toggle = (id) => setItems((p) => p.map((i) => (i.id === id ? { ...i, checked: !i.checked } : i)));
  const removeItem = (id) => {
    setItems((p) => p.filter((i) => i.id !== id));
    setScores((s) => { const n = { ...s }; Object.keys(n).forEach((k) => k.startsWith(id + "|") && delete n[k]); return n; });
  };
  const addItem = () => {
    const t = newItem.trim(); if (!t) return;
    setItems((p) => [...p, { id: uid(), raw: t, section: "Manual", checked: true }]);
    setNewItem("");
  };

  const addHyp = () => { const t = newHyp.trim(); if (!t) return; setHyps((p) => [...p, { id: uid(), label: t }]); setNewHyp(""); };
  const removeHyp = (id) => {
    setHyps((p) => p.filter((h) => h.id !== id));
    setScores((s) => { const n = { ...s }; Object.keys(n).forEach((k) => k.endsWith("|" + id) && delete n[k]); return n; });
  };
  const renameHyp = (id, label) => setHyps((p) => p.map((h) => (h.id === id ? { ...h, label } : h)));

  const cycle = (itemId, hypId) => {
    const k = `${itemId}|${hypId}`;
    setScores((s) => {
      const cur = s[k] || "";
      const next = CYCLE[(CYCLE.indexOf(cur) + 1) % CYCLE.length];
      const n = { ...s };
      if (next === "") delete n[k]; else n[k] = next;
      return n;
    });
  };

  const checked = items.filter((i) => i.checked);

  const groups = useMemo(() => {
    const order = []; const map = {};
    items.forEach((it) => {
      const s = it.section || "—";
      if (!map[s]) { map[s] = []; order.push(s); }
      map[s].push(it);
    });
    return order.map((s) => ({ section: s, items: map[s] }));
  }, [items]);

  const analysis = useMemo(() => {
    const tally = {}, refuted = {};
    hyps.forEach((h) => {
      let sum = 0, ref = false;
      checked.forEach((it) => {
        const v = scores[`${it.id}|${h.id}`];
        if (v) sum += WEIGHT[v];
        if (v === "−") ref = true;
      });
      tally[h.id] = sum; refuted[h.id] = ref;
    });
    let best = -Infinity;
    hyps.forEach((h) => { if (!refuted[h.id] && tally[h.id] > best) best = tally[h.id]; });
    const strongest = {};
    hyps.forEach((h) => { strongest[h.id] = !refuted[h.id] && best > 0 && tally[h.id] === best; });
    const nonDiag = {};
    checked.forEach((it) => {
      const vals = hyps.map((h) => scores[`${it.id}|${h.id}`]).filter((v) => v);
      nonDiag[it.id] = vals.length >= 2 && vals.every((v) => v === vals[0]);
    });
    return { tally, refuted, strongest, nonDiag };
  }, [checked, hyps, scores]);

  const C = { ink: "#0d0e0c", line: "#2b2c27", paper: "#e9e4d6", dim: "#8f8c7d", amber: "#d9a441", refuteText: "#c98a8a", green: "#9fd8a8", greenBg: "#16271c" };
  const mono = "'JetBrains Mono', ui-monospace, monospace";
  const serif = "'Instrument Serif', Georgia, serif";
  const cellColor = (v) => v === "−" ? C.refuteText : v === "++" || v === "+" ? C.green : v === "NA" ? C.dim : "#56554c";

  return (
    <div style={{ background: C.ink, color: C.paper, fontFamily: mono, minHeight: "100%", backgroundImage: "radial-gradient(circle at 15% -10%, #1b1c17 0%, transparent 45%), radial-gradient(circle at 110% 10%, #1a1813 0%, transparent 40%)" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Instrument+Serif:ital@0;1&family=JetBrains+Mono:wght@400;500;700&display=swap');
        * { box-sizing: border-box; }
        .stagger { animation: rise .5s cubic-bezier(.2,.7,.2,1) both; }
        @keyframes rise { from { opacity:0; transform: translateY(8px);} to {opacity:1; transform:none;} }
        textarea, input { font-family: ${mono}; }
        textarea:focus, input:focus { outline: 1px solid ${C.amber}; }
        .cellbtn:hover { background:#20211c !important; }
        .ev:hover { background:#141511; }
        ::-webkit-scrollbar { height:10px; width:10px;}
        ::-webkit-scrollbar-thumb { background:${C.line}; border-radius:6px;}
      `}</style>

      <div style={{ maxWidth: 1060, margin: "0 auto", padding: "34px 26px 90px" }}>
        <div className="stagger" style={{ marginBottom: 26 }}>
          <div style={{ fontSize: 11, letterSpacing: 3, color: C.amber, textTransform: "uppercase" }}>Analysis of Competing Hypotheses</div>
          <h1 style={{ fontFamily: serif, fontSize: 52, lineHeight: 0.95, margin: "6px 0 8px", fontWeight: 400 }}>ACH Workbench</h1>
          <p style={{ color: C.dim, fontSize: 13, maxWidth: 640, margin: 0 }}>
            Paste a MegaStonk dossier — it's kept readable, grouped by section, with its formatting intact. Tick the items you judge diagnostic, frame your hypotheses, and score each cell. The matrix tallies live, retires disproved columns, and dims evidence that doesn't discriminate.
          </p>
        </div>

        <Section n="01" title="Ingest dossier" C={C} serif={serif}>
          <textarea value={raw} onChange={(e) => setRaw(e.target.value)} placeholder="Paste the full MegaStonk output here…" spellCheck={false}
            style={{ width: "100%", minHeight: 110, background: "#0a0b09", color: C.paper, border: `1px solid ${C.line}`, borderRadius: 8, padding: 12, fontSize: 12.5, resize: "vertical" }} />
          <div style={{ display: "flex", gap: 10, marginTop: 10, alignItems: "center", flexWrap: "wrap" }}>
            <Btn onClick={parse} C={C} primary><ScanLine size={14} /> Extract evidence</Btn>
            <span style={{ color: C.dim, fontSize: 12 }}>{items.length} item{items.length === 1 ? "" : "s"} · {checked.length} in matrix</span>
            <label style={{ marginLeft: "auto", color: C.dim, fontSize: 11, display: "flex", gap: 6, alignItems: "center", cursor: "pointer" }}>
              <input type="checkbox" checked={showText} onChange={() => setShowText((v) => !v)} /> full text in matrix
            </label>
          </div>
        </Section>

        <Section n="02" title="Evidence — tick to add to matrix" C={C} serif={serif}>
          {items.length === 0 && <Empty C={C}>No items yet. Extract from a dossier above, or add one manually below.</Empty>}
          {groups.map((g) => {
            const open = collapsed[g.section] !== true;
            const checkedInGroup = g.items.filter((i) => i.checked).length;
            return (
              <div key={g.section} style={{ marginBottom: 10 }}>
                <div onClick={() => setCollapsed((c) => ({ ...c, [g.section]: open }))}
                  style={{ display: "flex", alignItems: "center", gap: 7, cursor: "pointer", padding: "6px 4px", borderBottom: `1px solid ${C.line}`, marginBottom: 4 }}>
                  {open ? <ChevronDown size={14} color={C.amber} /> : <ChevronRight size={14} color={C.amber} />}
                  <span style={{ fontFamily: serif, fontSize: 18, color: C.paper }}>{g.section}</span>
                  <span style={{ color: C.dim, fontSize: 11, marginLeft: "auto" }}>{checkedInGroup}/{g.items.length}</span>
                </div>
                {open && g.items.map((it) => (
                  <div key={it.id} className="ev" style={{ display: "flex", alignItems: "flex-start", gap: 11, padding: "8px 8px", borderRadius: 6, fontSize: 13, lineHeight: 1.5 }}>
                    <span onClick={() => toggle(it.id)} style={{ width: 17, height: 17, marginTop: 2, flexShrink: 0, border: `1px solid ${it.checked ? C.amber : C.line}`, background: it.checked ? C.amber : "transparent", borderRadius: 4, display: "grid", placeItems: "center", cursor: "pointer" }}>
                      {it.checked && <Check size={12} color={C.ink} strokeWidth={3} />}
                    </span>
                    <span style={{ flex: 1, color: it.checked ? C.paper : C.dim }}>
                      <Inline text={it.raw} C={it.checked ? C : { ...C, paper: C.dim }} />
                    </span>
                    <X size={14} color={C.dim} style={{ flexShrink: 0, marginTop: 3, cursor: "pointer" }} onClick={() => removeItem(it.id)} />
                  </div>
                ))}
              </div>
            );
          })}
          <Adder value={newItem} set={setNewItem} onAdd={addItem} placeholder="Add evidence manually…" C={C} />
        </Section>

        <Section n="03" title="Hypotheses — the competing explanations" C={C} serif={serif}>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 4 }}>
            {hyps.map((h, idx) => (
              <div key={h.id} style={{ display: "flex", alignItems: "center", gap: 6, background: "#0a0b09", border: `1px solid ${C.line}`, borderRadius: 6, padding: "5px 8px" }}>
                <span style={{ color: C.amber, fontSize: 11 }}>H{idx + 1}</span>
                <input value={h.label} onChange={(e) => renameHyp(h.id, e.target.value)} style={{ background: "transparent", border: "none", color: C.paper, fontSize: 12.5, width: Math.max(80, h.label.length * 7.5) }} />
                <X size={13} color={C.dim} onClick={() => removeHyp(h.id)} style={{ cursor: "pointer" }} />
              </div>
            ))}
          </div>
          <Adder value={newHyp} set={setNewHyp} onAdd={addHyp} placeholder="Add a hypothesis…" C={C} />
        </Section>

        <Section n="04" title="Matrix" C={C} serif={serif}>
          {checked.length === 0 || hyps.length === 0 ? (
            <Empty C={C}>Tick evidence and add at least one hypothesis to build the matrix.</Empty>
          ) : (
            <div style={{ overflowX: "auto" }}>
              <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12.5 }}>
                <thead>
                  <tr>
                    <th style={{ textAlign: "left", padding: "8px 10px", borderBottom: `1px solid ${C.line}`, color: C.dim, fontWeight: 500, minWidth: 260 }}>Evidence</th>
                    {hyps.map((h, idx) => {
                      const ref = analysis.refuted[h.id], best = analysis.strongest[h.id];
                      return (
                        <th key={h.id} style={{ padding: "8px 6px", borderBottom: `1px solid ${C.line}`, minWidth: 86, textAlign: "center", color: ref ? "#5c5b53" : best ? C.green : C.paper, background: best ? C.greenBg : "transparent", textDecoration: ref ? "line-through" : "none", verticalAlign: "bottom" }} title={h.label}>
                          <div style={{ fontSize: 10, color: C.amber, opacity: ref ? 0.4 : 1 }}>H{idx + 1}</div>
                          <div style={{ maxWidth: 110, margin: "0 auto", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{h.label}</div>
                        </th>
                      );
                    })}
                  </tr>
                </thead>
                <tbody>
                  {checked.map((it) => {
                    const nd = analysis.nonDiag[it.id];
                    return (
                      <tr key={it.id}>
                        <td style={{ padding: "9px 10px", borderBottom: `1px solid #1d1e1a`, color: nd ? "#6b6a60" : C.paper, verticalAlign: "top", lineHeight: 1.5, maxWidth: showText ? 520 : 300 }} title={nd ? "Non-diagnostic: scores the same across all hypotheses" : ""}>
                          <span style={{ color: C.dim, fontSize: 10, marginRight: 7 }}>{it.section}</span>
                          {showText ? <Inline text={it.raw} C={nd ? { ...C, paper: "#6b6a60" } : C} /> : <span>{stripMd(it.raw).slice(0, 80)}{stripMd(it.raw).length > 80 ? "…" : ""}</span>}
                          {nd && <span style={{ color: "#46453e", fontSize: 10 }}> · non-diagnostic</span>}
                        </td>
                        {hyps.map((h) => {
                          const v = scores[`${it.id}|${h.id}`] || "";
                          const ref = analysis.refuted[h.id], best = analysis.strongest[h.id];
                          return (
                            <td key={h.id} style={{ borderBottom: `1px solid #1d1e1a`, background: ref ? "#101010" : best ? "#10180f" : "transparent", opacity: nd ? 0.45 : 1, textAlign: "center", padding: 0 }}>
                              <button className="cellbtn" onClick={() => cycle(it.id, h.id)} style={{ width: "100%", height: 36, border: "none", background: "transparent", color: cellColor(v), fontFamily: mono, fontSize: v === "NA" ? 11 : 16, fontWeight: 700, cursor: "pointer" }}>{GLYPH[v]}</button>
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                  <tr>
                    <td style={{ padding: "10px", color: C.dim, textTransform: "uppercase", letterSpacing: 1, fontSize: 11 }}>Tally</td>
                    {hyps.map((h) => {
                      const ref = analysis.refuted[h.id], best = analysis.strongest[h.id];
                      return (
                        <td key={h.id} style={{ textAlign: "center", padding: "10px 6px", fontWeight: 700, fontSize: 15, color: ref ? "#5c5b53" : best ? C.green : C.paper, background: best ? C.greenBg : ref ? "#101010" : "transparent" }}>
                          {ref ? "✕" : analysis.tally[h.id] > 0 ? "+" + analysis.tally[h.id] : analysis.tally[h.id]}
                        </td>
                      );
                    })}
                  </tr>
                </tbody>
              </table>

              <div style={{ display: "flex", flexWrap: "wrap", gap: 16, marginTop: 16, fontSize: 11, color: C.dim }}>
                <Leg sw={C.refuteText} t="− disproves (retires column)" />
                <Leg sw={C.dim} t="n/a — not applicable" />
                <Leg sw={C.green} t="+ / ++ consistent" />
                <span style={{ display: "flex", gap: 6, alignItems: "center" }}><span style={{ width: 12, height: 12, background: C.greenBg, border: `1px solid ${C.green}` }} /> strongest surviving hypothesis</span>
                <span style={{ display: "flex", gap: 6, alignItems: "center" }}><span style={{ width: 12, height: 12, background: "#101010", border: `1px solid ${C.line}` }} /> retired / non-diagnostic (dimmed)</span>
                <span style={{ color: "#56554c" }}>click a cell to cycle · empty → − → n/a → + → ++</span>
              </div>
            </div>
          )}
        </Section>
      </div>
    </div>
  );
}

function Section({ n, title, C, serif, children }) {
  return (
    <div className="stagger" style={{ marginTop: 30 }}>
      <div style={{ display: "flex", alignItems: "baseline", gap: 12, marginBottom: 12 }}>
        <span style={{ color: C.amber, fontSize: 12, letterSpacing: 2 }}>{n}</span>
        <h2 style={{ fontFamily: serif, fontSize: 24, fontWeight: 400, margin: 0 }}>{title}</h2>
        <div style={{ flex: 1, height: 1, background: C.line }} />
      </div>
      {children}
    </div>
  );
}
function Btn({ children, onClick, C, primary }) {
  return (
    <button onClick={onClick} style={{ display: "inline-flex", alignItems: "center", gap: 7, background: primary ? C.amber : "transparent", color: primary ? C.ink : C.paper, border: `1px solid ${primary ? C.amber : C.line}`, borderRadius: 6, padding: "7px 13px", fontSize: 12.5, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>{children}</button>
  );
}
function Adder({ value, set, onAdd, placeholder, C }) {
  return (
    <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
      <input value={value} onChange={(e) => set(e.target.value)} onKeyDown={(e) => e.key === "Enter" && onAdd()} placeholder={placeholder}
        style={{ flex: 1, background: "#0a0b09", border: `1px solid ${C.line}`, color: C.paper, borderRadius: 6, padding: "8px 10px", fontSize: 12.5 }} />
      <Btn onClick={onAdd} C={C}><Plus size={14} /> Add</Btn>
    </div>
  );
}
function Empty({ children, C }) {
  return (<div style={{ border: `1px dashed ${C.line}`, borderRadius: 8, padding: "18px 16px", color: C.dim, fontSize: 12.5, textAlign: "center" }}>{children}</div>);
}
function Leg({ sw, t }) {
  return (<span style={{ display: "flex", gap: 6, alignItems: "center" }}><span style={{ color: sw, fontWeight: 700 }}>●</span>{t}</span>);
}
