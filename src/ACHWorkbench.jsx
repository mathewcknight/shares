import React, { useState, useMemo } from "react";
import { Plus, X, Check, ChevronDown, ChevronRight } from "lucide-react";

const CYCLE = ["", "−", "NA", "+", "++"];
const WEIGHT = { "−": -2, NA: 0, "+": 1, "++": 2 };
const GLYPH  = { "": "·", "−": "−", NA: "n/a", "+": "+", "++": "++" };

let _id = 1;
const uid = () => `x${_id++}`;

function stripMd(s) {
  return s.replace(/\*\*(.*?)\*\*/g, "$1").replace(/`(.*?)`/g, "$1")
          .replace(/\*(.*?)\*/g, "$1").replace(/#+\s*/g, "").trim();
}

// ---- inline markdown renderer --------------------------------------------
function Inline({ text, C }) {
  const nodes = [];
  const re = /(\*\*[^*]+\*\*)|(`[^`]+`)|(\[[^\]]+\]\([^)]+\))|(\*[^*\s][^*]*\*)/g;
  let last = 0, m, k = 0;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) nodes.push(text.slice(last, m.index));
    const tok = m[0];
    if (tok.startsWith("**"))
      nodes.push(<strong key={k++} style={{ color: C.paper, fontWeight: 700 }}>{tok.slice(2, -2)}</strong>);
    else if (tok.startsWith("`"))
      nodes.push(<code key={k++} style={{ background: "#0a0b09", border: `1px solid ${C.line}`, borderRadius: 4, padding: "1px 5px", fontSize: "0.9em", color: C.amber }}>{tok.slice(1, -1)}</code>);
    else if (tok.startsWith("[")) {
      const mm = tok.match(/\[([^\]]+)\]\(([^)]+)\)/);
      nodes.push(<a key={k++} href={mm[2]} target="_blank" rel="noreferrer" style={{ color: C.amber, textDecoration: "underline" }}>{mm[1]}</a>);
    } else
      nodes.push(<em key={k++} style={{ color: C.paper }}>{tok.slice(1, -1)}</em>);
    last = m.index + tok.length;
  }
  if (last < text.length) nodes.push(text.slice(last));
  return <>{nodes}</>;
}

// ---- markdown → structured blocks ---------------------------------------
function parseToBlocks(raw) {
  if (!raw) return [];
  const blocks = [];
  const lines  = raw.split(/\r?\n/);
  let section  = "";
  let i        = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Heading
    const h = line.match(/^\s{0,3}(#{1,6})\s+(.+?)\s*$/);
    if (h) {
      section = stripMd(h[2]);
      blocks.push({ type: "heading", level: h[1].length, text: h[2], section });
      i++; continue;
    }
    const boldHead = line.match(/^\s*\*\*(.+?)\*\*:?\s*$/);
    if (boldHead) {
      section = stripMd(boldHead[1]);
      blocks.push({ type: "heading", level: 3, text: boldHead[1], section });
      i++; continue;
    }

    // Horizontal rule / separator line — skip
    if (/^\s*---+\s*$/.test(line)) { i++; continue; }

    // Table — collect all consecutive pipe lines
    if (/^\s*\|/.test(line)) {
      const tableLines = [];
      while (i < lines.length && /^\s*\|/.test(lines[i])) {
        tableLines.push(lines[i]);
        i++;
      }
      if (tableLines.length < 2) continue;
      const headerCells = tableLines[0].split("|").slice(1, -1).map(c => c.trim());
      // Line 1 is separator if all cells are only dashes/colons/spaces
      const isSep = r => r.split("|").slice(1, -1).every(c => /^[\s:|-]+$/.test(c));
      const dataStart = isSep(tableLines[1]) ? 2 : 1;
      blocks.push({ type: "table-header", cells: headerCells, section });
      for (let r = dataStart; r < tableLines.length; r++) {
        const cells = tableLines[r].split("|").slice(1, -1).map(c => c.trim());
        if (!cells.length || cells.every(c => /^[\s:|-]+$/.test(c))) continue;
        const rawText = cells.filter(Boolean).join(" · ");
        if (stripMd(rawText).length >= 4)
          blocks.push({ type: "table-row", id: uid(), cells, headers: headerCells, section, checked: false, raw: rawText });
      }
      continue;
    }

    // Bullet / numbered list
    const b = line.match(/^\s*(?:[-•*]|\d+[.)])\s+(.*\S)\s*$/);
    if (b) {
      const rawText = b[1].replace(/⚠️/g, "").trim();
      if (stripMd(rawText).length >= 4)
        blocks.push({ type: "bullet", id: uid(), raw: rawText, section, checked: false });
      i++; continue;
    }

    // Paragraph (non-selectable, rendered as context)
    const trimmed = line.trim();
    if (trimmed) blocks.push({ type: "para", text: trimmed, section });
    i++;
  }
  return blocks;
}

// ---- rendered dossier component -----------------------------------------
function RenderedDossier({ blocks, onToggle, C, mono, serif, collapsed, onCollapseToggle }) {
  const els = [];
  let i = 0;
  let sectionCollapsed = false;

  while (i < blocks.length) {
    const blk = blocks[i];

    // ── heading ──────────────────────────────────────────────────────────
    if (blk.type === "heading") {
      sectionCollapsed = collapsed[blk.section] === true;
      els.push(
        <div key={`h${i}`}
          onClick={() => onCollapseToggle(blk.section)}
          style={{
            display: "flex", alignItems: "center", gap: 8, cursor: "pointer",
            padding: "7px 2px",
            marginTop: blk.level <= 2 ? 20 : 12,
            borderBottom: blk.level <= 2 ? `1px solid ${C.line}` : "none",
          }}
        >
          {sectionCollapsed
            ? <ChevronRight size={13} color={C.amber} />
            : <ChevronDown  size={13} color={C.amber} />}
          <span style={{
            fontFamily: blk.level <= 2 ? serif : mono,
            fontSize:   blk.level === 1 ? 22 : blk.level === 2 ? 17 : 13,
            fontWeight: blk.level <= 2 ? 400 : 500,
            color:      blk.level <= 2 ? C.paper : C.amber,
            letterSpacing: blk.level >= 3 ? 1.5 : 0,
            textTransform: blk.level >= 3 ? "uppercase" : "none",
          }}>
            <Inline text={blk.text} C={C} />
          </span>
        </div>
      );
      i++; continue;
    }

    if (sectionCollapsed) { i++; continue; }

    // ── table ─────────────────────────────────────────────────────────────
    if (blk.type === "table-header") {
      const headers = blk.cells;
      const rows = [];
      let j = i + 1;
      while (j < blocks.length && blocks[j].type === "table-row") { rows.push(blocks[j]); j++; }

      els.push(
        <div key={`t${i}`} style={{ overflowX: "auto", margin: "8px 0 4px" }}>
          <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12 }}>
            <thead>
              <tr>
                <th style={{ width: 26, borderBottom: `1px solid ${C.line}` }} />
                {headers.map((h, k) => (
                  <th key={k} style={{
                    textAlign: "left", padding: "4px 10px 4px 4px",
                    borderBottom: `1px solid ${C.line}`,
                    color: C.dim, fontWeight: 500, fontSize: 10.5,
                    letterSpacing: 0.5, whiteSpace: "nowrap",
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map(row => (
                <tr key={row.id} onClick={() => onToggle(row.id)} className="ev"
                  style={{ cursor: "pointer", background: row.checked ? "#151a10" : "transparent" }}>
                  <td style={{ padding: "5px 6px", verticalAlign: "middle" }}>
                    <span style={{
                      width: 14, height: 14, display: "grid", placeItems: "center",
                      border: `1px solid ${row.checked ? C.amber : C.line}`,
                      background: row.checked ? C.amber : "transparent",
                      borderRadius: 3, flexShrink: 0,
                    }}>
                      {row.checked && <Check size={9} color={C.ink} strokeWidth={3} />}
                    </span>
                  </td>
                  {row.cells.map((cell, k) => (
                    <td key={k} style={{
                      padding: "5px 10px 5px 4px", verticalAlign: "top",
                      borderBottom: `1px solid #161710`,
                      color: row.checked ? C.paper : "#6b6960",
                      lineHeight: 1.5,
                    }}>
                      <Inline text={cell} C={row.checked ? C : { ...C, paper: "#6b6960", amber: "#6b6960" }} />
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
      i = j; continue;
    }

    // ── bullet ───────────────────────────────────────────────────────────
    if (blk.type === "bullet") {
      els.push(
        <div key={`b${i}`} className="ev" onClick={() => onToggle(blk.id)}
          style={{
            display: "flex", alignItems: "flex-start", gap: 10,
            padding: "5px 8px", borderRadius: 5, cursor: "pointer",
            fontSize: 13, lineHeight: 1.5,
            background: blk.checked ? "#151a10" : "transparent",
          }}>
          <span style={{
            width: 14, height: 14, marginTop: 3, flexShrink: 0,
            display: "grid", placeItems: "center",
            border: `1px solid ${blk.checked ? C.amber : C.line}`,
            background: blk.checked ? C.amber : "transparent",
            borderRadius: 3,
          }}>
            {blk.checked && <Check size={9} color={C.ink} strokeWidth={3} />}
          </span>
          <span style={{ color: blk.checked ? C.paper : "#6b6960" }}>
            <Inline text={blk.raw} C={blk.checked ? C : { ...C, paper: "#6b6960", amber: "#6b6960" }} />
          </span>
        </div>
      );
      i++; continue;
    }

    // ── paragraph ─────────────────────────────────────────────────────────
    if (blk.type === "para") {
      els.push(
        <p key={`p${i}`} style={{ color: "#5a5950", fontSize: 12, lineHeight: 1.6, margin: "4px 2px" }}>
          <Inline text={blk.text} C={{ ...C, paper: "#5a5950", amber: "#8a7a4a" }} />
        </p>
      );
      i++; continue;
    }

    i++;
  }

  return <div>{els}</div>;
}

// ---- main component -------------------------------------------------------
export default function ACHWorkbench() {
  const [raw,        setRaw]        = useState("");
  const [blocks,     setBlocks]     = useState([]);
  const [manualItems, setManualItems] = useState([]);
  const [hyps,       setHyps]       = useState([]);
  const [scores,     setScores]     = useState({});
  const [newHyp,     setNewHyp]     = useState("");
  const [newItem,    setNewItem]    = useState("");
  const [collapsed,  setCollapsed]  = useState({});
  const [showText,   setShowText]   = useState(true);

  const render = () => {
    const found = parseToBlocks(raw);
    setBlocks(prev => {
      const wasChecked = new Set(prev.filter(b => b.checked && b.raw).map(b => b.raw));
      return found.map(b => b.id ? { ...b, checked: wasChecked.has(b.raw) } : b);
    });
  };

  const toggleBlock = (id) =>
    setBlocks(prev => prev.map(b => b.id === id ? { ...b, checked: !b.checked } : b));

  const onCollapseToggle = (section) =>
    setCollapsed(c => ({ ...c, [section]: !c[section] }));

  const selectAll   = () => setBlocks(prev => prev.map(b => b.id ? { ...b, checked: true }  : b));
  const deselectAll = () => setBlocks(prev => prev.map(b => b.id ? { ...b, checked: false } : b));

  const addItem = () => {
    const t = newItem.trim(); if (!t) return;
    setManualItems(p => [...p, { id: uid(), raw: t, section: "Manual" }]);
    setNewItem("");
  };
  const removeManual = (id) => {
    setManualItems(p => p.filter(i => i.id !== id));
    setScores(s => { const n = { ...s }; Object.keys(n).forEach(k => k.startsWith(id + "|") && delete n[k]); return n; });
  };

  const addHyp    = () => { const t = newHyp.trim(); if (!t) return; setHyps(p => [...p, { id: uid(), label: t }]); setNewHyp(""); };
  const removeHyp = (id) => {
    setHyps(p => p.filter(h => h.id !== id));
    setScores(s => { const n = { ...s }; Object.keys(n).forEach(k => k.endsWith("|" + id) && delete n[k]); return n; });
  };
  const renameHyp = (id, label) => setHyps(p => p.map(h => h.id === id ? { ...h, label } : h));

  const cycle = (itemId, hypId) => {
    const k = `${itemId}|${hypId}`;
    setScores(s => {
      const cur  = s[k] || "";
      const next = CYCLE[(CYCLE.indexOf(cur) + 1) % CYCLE.length];
      const n    = { ...s };
      if (next === "") delete n[k]; else n[k] = next;
      return n;
    });
  };

  const checkedBlocks = blocks.filter(b => b.id && b.checked);
  const checked       = [...checkedBlocks, ...manualItems];
  const totalItems    = blocks.filter(b => b.id).length;

  const analysis = useMemo(() => {
    const tally = {}, refuted = {};
    hyps.forEach(h => {
      let sum = 0, ref = false;
      checked.forEach(it => {
        const v = scores[`${it.id}|${h.id}`];
        if (v) sum += WEIGHT[v];
        if (v === "−") ref = true;
      });
      tally[h.id] = sum; refuted[h.id] = ref;
    });
    let best = -Infinity;
    hyps.forEach(h => { if (!refuted[h.id] && tally[h.id] > best) best = tally[h.id]; });
    const strongest = {};
    hyps.forEach(h => { strongest[h.id] = !refuted[h.id] && best > 0 && tally[h.id] === best; });
    const nonDiag = {};
    checked.forEach(it => {
      const vals = hyps.map(h => scores[`${it.id}|${h.id}`]).filter(v => v);
      nonDiag[it.id] = vals.length >= 2 && vals.every(v => v === vals[0]);
    });
    return { tally, refuted, strongest, nonDiag };
  }, [checked, hyps, scores]);

  const C     = { ink: "#0d0e0c", line: "#2b2c27", paper: "#e9e4d6", dim: "#8f8c7d", amber: "#d9a441", refuteText: "#c98a8a", green: "#9fd8a8", greenBg: "#16271c" };
  const mono  = "'JetBrains Mono', ui-monospace, monospace";
  const serif = "'Instrument Serif', Georgia, serif";
  const cellColor = v => v === "−" ? C.refuteText : v === "++" || v === "+" ? C.green : v === "NA" ? C.dim : "#56554c";

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
        .ev:hover { background:#111309 !important; }
        ::-webkit-scrollbar { height:10px; width:10px;}
        ::-webkit-scrollbar-thumb { background:${C.line}; border-radius:6px;}
      `}</style>

      <div style={{ maxWidth: 1060, margin: "0 auto", padding: "34px 26px 90px" }}>
        <div className="stagger" style={{ marginBottom: 26 }}>
          <div style={{ fontSize: 11, letterSpacing: 3, color: C.amber, textTransform: "uppercase" }}>Analysis of Competing Hypotheses</div>
          <h1 style={{ fontFamily: serif, fontSize: 52, lineHeight: 0.95, margin: "6px 0 8px", fontWeight: 400 }}>ACH Workbench</h1>
          <p style={{ color: C.dim, fontSize: 13, maxWidth: 640, margin: 0 }}>
            Paste a MegaStonk dossier and click <em>Render</em>. The dossier appears below with its formatting intact — click any row or bullet to select it as evidence. Then frame your hypotheses and score the matrix.
          </p>
        </div>

        {/* ── 01 Ingest ───────────────────────────────────────────────────── */}
        <Section n="01" title="Ingest dossier" C={C} serif={serif}>
          <textarea value={raw} onChange={e => setRaw(e.target.value)}
            placeholder="Paste the full MegaStonk output here…" spellCheck={false}
            style={{ width: "100%", minHeight: 100, background: "#0a0b09", color: C.paper, border: `1px solid ${C.line}`, borderRadius: 8, padding: 12, fontSize: 12.5, resize: "vertical" }} />
          <div style={{ display: "flex", gap: 10, marginTop: 10, alignItems: "center", flexWrap: "wrap" }}>
            <Btn onClick={render} C={C} primary>Render dossier</Btn>
            {totalItems > 0 && <>
              <span style={{ color: C.dim, fontSize: 12 }}>{checked.length} / {totalItems + manualItems.length} selected</span>
              <button onClick={selectAll}   style={ghostBtn(C)}>select all</button>
              <button onClick={deselectAll} style={ghostBtn(C)}>deselect all</button>
            </>}
            {checked.length > 0 && (
              <label style={{ marginLeft: "auto", color: C.dim, fontSize: 11, display: "flex", gap: 6, alignItems: "center", cursor: "pointer" }}>
                <input type="checkbox" checked={showText} onChange={() => setShowText(v => !v)} /> full text in matrix
              </label>
            )}
          </div>

          {blocks.length > 0 && (
            <div style={{ marginTop: 18, borderTop: `1px solid ${C.line}`, paddingTop: 12 }}>
              <RenderedDossier
                blocks={blocks} onToggle={toggleBlock}
                C={C} mono={mono} serif={serif}
                collapsed={collapsed} onCollapseToggle={onCollapseToggle}
              />
            </div>
          )}
        </Section>

        {/* ── 02 Manual evidence ──────────────────────────────────────────── */}
        <Section n="02" title="Add evidence manually" C={C} serif={serif}>
          {manualItems.length === 0 && blocks.length === 0 &&
            <Empty C={C}>Render a dossier above, or type evidence here directly.</Empty>}
          {manualItems.map(it => (
            <div key={it.id} className="ev" style={{ display: "flex", alignItems: "flex-start", gap: 11, padding: "7px 8px", borderRadius: 6, fontSize: 13, lineHeight: 1.5 }}>
              <span style={{ flex: 1 }}><Inline text={it.raw} C={C} /></span>
              <X size={14} color={C.dim} style={{ flexShrink: 0, marginTop: 3, cursor: "pointer" }} onClick={() => removeManual(it.id)} />
            </div>
          ))}
          <Adder value={newItem} set={setNewItem} onAdd={addItem} placeholder="Type evidence and press Enter or Add…" C={C} />
        </Section>

        {/* ── 03 Hypotheses ───────────────────────────────────────────────── */}
        <Section n="03" title="Hypotheses — the competing explanations" C={C} serif={serif}>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 8, marginBottom: 4 }}>
            {hyps.map((h, idx) => (
              <div key={h.id} style={{ display: "flex", alignItems: "center", gap: 6, background: "#0a0b09", border: `1px solid ${C.line}`, borderRadius: 6, padding: "5px 8px" }}>
                <span style={{ color: C.amber, fontSize: 11 }}>H{idx + 1}</span>
                <input value={h.label} onChange={e => renameHyp(h.id, e.target.value)}
                  style={{ background: "transparent", border: "none", color: C.paper, fontSize: 12.5, width: Math.max(80, h.label.length * 7.5) }} />
                <X size={13} color={C.dim} onClick={() => removeHyp(h.id)} style={{ cursor: "pointer" }} />
              </div>
            ))}
          </div>
          <Adder value={newHyp} set={setNewHyp} onAdd={addHyp} placeholder="Add a hypothesis…" C={C} />
        </Section>

        {/* ── 04 Matrix ───────────────────────────────────────────────────── */}
        <Section n="04" title="Matrix" C={C} serif={serif}>
          {checked.length === 0 || hyps.length === 0 ? (
            <Empty C={C}>Select evidence and add at least one hypothesis to build the matrix.</Empty>
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
                  {checked.map(it => {
                    const nd = analysis.nonDiag[it.id];
                    return (
                      <tr key={it.id}>
                        <td style={{ padding: "9px 10px", borderBottom: `1px solid #1d1e1a`, color: nd ? "#6b6a60" : C.paper, verticalAlign: "top", lineHeight: 1.5, maxWidth: showText ? 520 : 300 }}
                          title={nd ? "Non-diagnostic: same score across all hypotheses" : ""}>
                          <span style={{ color: C.dim, fontSize: 10, marginRight: 7 }}>{it.section}</span>
                          {showText
                            ? <Inline text={it.raw} C={nd ? { ...C, paper: "#6b6a60" } : C} />
                            : <span>{stripMd(it.raw).slice(0, 80)}{stripMd(it.raw).length > 80 ? "…" : ""}</span>}
                          {nd && <span style={{ color: "#46453e", fontSize: 10 }}> · non-diagnostic</span>}
                        </td>
                        {hyps.map(h => {
                          const v = scores[`${it.id}|${h.id}`] || "";
                          const ref = analysis.refuted[h.id], best = analysis.strongest[h.id];
                          return (
                            <td key={h.id} style={{ borderBottom: `1px solid #1d1e1a`, background: ref ? "#101010" : best ? "#10180f" : "transparent", opacity: nd ? 0.45 : 1, textAlign: "center", padding: 0 }}>
                              <button className="cellbtn" onClick={() => cycle(it.id, h.id)}
                                style={{ width: "100%", height: 36, border: "none", background: "transparent", color: cellColor(v), fontFamily: mono, fontSize: v === "NA" ? 11 : 16, fontWeight: 700, cursor: "pointer" }}>
                                {GLYPH[v]}
                              </button>
                            </td>
                          );
                        })}
                      </tr>
                    );
                  })}
                  <tr>
                    <td style={{ padding: "10px", color: C.dim, textTransform: "uppercase", letterSpacing: 1, fontSize: 11 }}>Tally</td>
                    {hyps.map(h => {
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
                <Leg sw={C.dim}        t="n/a — not applicable" />
                <Leg sw={C.green}      t="+ / ++ consistent" />
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

// ---- shared helpers -------------------------------------------------------
const ghostBtn = C => ({
  background: "transparent", border: `1px solid ${C.line}`, color: C.dim,
  borderRadius: 5, padding: "3px 9px", fontSize: 11, cursor: "pointer", fontFamily: "inherit",
});

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
      <input value={value} onChange={e => set(e.target.value)} onKeyDown={e => e.key === "Enter" && onAdd()} placeholder={placeholder}
        style={{ flex: 1, background: "#0a0b09", border: `1px solid ${C.line}`, color: C.paper, borderRadius: 6, padding: "8px 10px", fontSize: 12.5 }} />
      <Btn onClick={onAdd} C={C}><Plus size={14} /> Add</Btn>
    </div>
  );
}
function Empty({ children, C }) {
  return <div style={{ border: `1px dashed ${C.line}`, borderRadius: 8, padding: "18px 16px", color: C.dim, fontSize: 12.5, textAlign: "center" }}>{children}</div>;
}
function Leg({ sw, t }) {
  return <span style={{ display: "flex", gap: 6, alignItems: "center" }}><span style={{ color: sw, fontWeight: 700 }}>●</span>{t}</span>;
}
