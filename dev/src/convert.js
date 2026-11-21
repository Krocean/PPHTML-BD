// Build-time converter that runs inside Edge (headless): turns modern CSS / JavaScript into what the Internet
// Explorer 11 engine of the PowerPoint host understands. The viewport of a page is fixed (the canvas), so media
// queries, viewport units and clamp() can be settled here once and for all.
// Bundled by lab\tools\build.js into assets\tools\convert.js (global IHConvert).
import postcss from "postcss";
import nesting from "postcss-nesting";
import cssVariables from "postcss-css-variables";
import valueParser from "postcss-value-parser";
import selectorParser from "postcss-selector-parser";
import autoprefixer from "autoprefixer";
import flexbugs from "postcss-flexbugs-fixes";
import * as acorn from "acorn";

// ------------------------------------------------------------------ small helpers
const round = (v, d = 3) => { const k = Math.pow(10, d); return Math.round(v * k) / k; };
const px = v => round(v, 2) + "px";

function splitTop(str, sepChar) {       // split at separators that are not inside brackets or strings
  const out = []; let depth = 0, cur = "", q = null;
  for (let i = 0; i < str.length; i++) {
    const c = str[i];
    if (q) { cur += c; if (c === q) { q = null; } continue; }
    if (c === '"' || c === "'") { q = c; cur += c; continue; }
    if (c === "(" || c === "[") { depth++; } else if (c === ")" || c === "]") { depth--; }
    if (depth === 0 && (sepChar === " " ? /\s/.test(c) : c === sepChar)) { if (cur.trim() || sepChar !== " ") { out.push(cur.trim()); } cur = ""; continue; }
    cur += c;
  }
  if (cur.trim() || (sepChar !== " " && out.length)) { out.push(cur.trim()); }
  return out;
}

// ------------------------------------------------------------------ lengths that can be settled now
function makeLengths(ctx) {
  const unit = { px: 1, vw: ctx.W / 100, vh: ctx.H / 100, vmin: Math.min(ctx.W, ctx.H) / 100, vmax: Math.max(ctx.W, ctx.H) / 100, rem: ctx.rem,
    dvw: ctx.W / 100, svw: ctx.W / 100, lvw: ctx.W / 100, dvh: ctx.H / 100, svh: ctx.H / 100, lvh: ctx.H / 100, vi: ctx.W / 100, vb: ctx.H / 100,
    pt: 4 / 3, pc: 16, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 };
  // expression evaluator: numbers, lengths, + - * /, brackets, min / max / clamp / calc
  function evaluate(text) {
    const toks = []; const rx = /\s*(?:([0-9]*\.?[0-9]+(?:e[-+]?\d+)?)([a-z%]*)|([a-z-]+)\(|([-+*\/(),]))/gi; let m, last = 0;
    while ((m = rx.exec(text))) {
      if (m.index !== last) { return null; }
      last = rx.lastIndex;
      if (m[1] !== undefined) { toks.push({ t: "n", v: parseFloat(m[1]), u: (m[2] || "").toLowerCase() }); }
      else if (m[3]) { toks.push({ t: "f", v: m[3].toLowerCase() }); }
      else { toks.push({ t: m[4] }); }
    }
    if (text.slice(last).trim()) { return null; }
    let i = 0;
    const peek = () => toks[i], take = () => toks[i++];
    function factor() {
      const k = take(); if (!k) { throw 0; }
      if (k.t === "n") {
        if (!k.u) { return { v: k.v, d: 0 }; }
        if (unit[k.u] === undefined) { throw 0; }
        return { v: k.v * unit[k.u], d: 1 };
      }
      if (k.t === "-") { const f = factor(); return { v: -f.v, d: f.d }; }
      if (k.t === "+") { return factor(); }
      if (k.t === "(") { const e = expr(); if (!take() || toks[i - 1].t !== ")") { throw 0; } return e; }
      if (k.t === "f") {
        const args = [expr()];
        while (peek() && peek().t === ",") { take(); args.push(expr()); }
        if (!take() || toks[i - 1].t !== ")") { throw 0; }
        const d = args[0].d; if (args.some(a => a.d !== d)) { throw 0; }
        if (k.v === "calc") { if (args.length !== 1) { throw 0; } return args[0]; }
        if (k.v === "min") { return { v: Math.min.apply(null, args.map(a => a.v)), d }; }
        if (k.v === "max") { return { v: Math.max.apply(null, args.map(a => a.v)), d }; }
        if (k.v === "clamp") { if (args.length !== 3) { throw 0; } return { v: Math.max(args[0].v, Math.min(args[1].v, args[2].v)), d }; }
        throw 0;
      }
      throw 0;
    }
    function term() {
      let a = factor();
      while (peek() && (peek().t === "*" || peek().t === "/")) {
        const op = take().t, b = factor();
        if (op === "*") { if (a.d && b.d) { throw 0; } a = { v: a.v * b.v, d: a.d || b.d }; }
        else { if (b.d || b.v === 0) { throw 0; } a = { v: a.v / b.v, d: a.d }; }
      }
      return a;
    }
    function expr() {
      let a = term();
      while (peek() && (peek().t === "+" || peek().t === "-")) {
        const op = take().t, b = term(); if (a.d !== b.d) { throw 0; }
        a = { v: op === "+" ? a.v + b.v : a.v - b.v, d: a.d };
      }
      return a;
    }
    try { const r = expr(); if (i !== toks.length) { return null; } return r; } catch (e) { return null; }
  }
  return {
    unit,
    evaluate,
    // rewrites a declaration value: min()/max()/clamp() and units IE11 lacks become px when that is possible
    lower(value) {
      if (!/(clamp|min|max)\(|\d(vmin|vmax|dv[wh]|sv[wh]|lv[wh]|vi|vb)\b/i.test(value)) { return value; }
      const tree = valueParser(value);
      tree.walk(node => {
        if (node.type === "function" && /^(clamp|min|max|calc)$/i.test(node.value)) {
          const text = valueParser.stringify(node);
          if (!/(clamp|min|max)\(|\d(vmin|vmax|dv[wh]|sv[wh]|lv[wh]|vi|vb)\b/i.test(text)) { return; }
          const r = evaluate(text);
          if (r) { node.type = "word"; node.value = r.d ? px(r.v) : String(round(r.v, 4)); return false; }
        } else if (node.type === "word") {
          const m = /^(-?[0-9]*\.?[0-9]+)(vmin|vmax|dv[wh]|sv[wh]|lv[wh]|vi|vb)$/i.exec(node.value);
          if (m) { node.value = px(parseFloat(m[1]) * unit[m[2].toLowerCase()]); }
        }
      });
      return tree.toString();
    }
  };
}

// ------------------------------------------------------------------ colours
let colorCanvas = null;
function resolveColor(text) {            // any colour Edge understands -> rgb() / rgba(); null when it cannot
  try {
    if (typeof document === "undefined") { return null; }
    if (!colorCanvas) { colorCanvas = document.createElement("canvas"); colorCanvas.width = colorCanvas.height = 1; }
    const c = colorCanvas.getContext("2d", { willReadFrequently: true });
    c.clearRect(0, 0, 1, 1); c.fillStyle = "rgba(1,2,3,0.5)"; const before = c.fillStyle;
    c.fillStyle = text; if (c.fillStyle === before && !/^rgba\(1,\s*2,\s*3/.test(text)) { return null; }
    c.fillRect(0, 0, 1, 1);
    const d = c.getImageData(0, 0, 1, 1).data;
    if (d[3] === 255) { return "rgb(" + d[0] + ", " + d[1] + ", " + d[2] + ")"; }
    if (d[3] === 0) { return "rgba(0, 0, 0, 0)"; }
    // read the colour itself (not premultiplied) by painting it opaque as well
    const m = /^#([0-9a-f]{6})([0-9a-f]{2})$/i.exec(c.fillStyle) || null;
    const a = round(d[3] / 255, 3);
    if (m) { const n = parseInt(m[1], 16); return "rgba(" + (n >> 16) + ", " + ((n >> 8) & 255) + ", " + (n & 255) + ", " + a + ")"; }
    const mm = /^rgba\(\s*([\d.]+),\s*([\d.]+),\s*([\d.]+),\s*([\d.]+)/.exec(c.fillStyle);
    if (mm) { return "rgba(" + Math.round(mm[1]) + ", " + Math.round(mm[2]) + ", " + Math.round(mm[3]) + ", " + round(parseFloat(mm[4]), 3) + ")"; }
    const ms = /^color\(srgb\s+([\d.e-]+)\s+([\d.e-]+)\s+([\d.e-]+)(?:\s*\/\s*([\d.e-]+))?\)/.exec(c.fillStyle);
    if (ms) { const k = v => Math.max(0, Math.min(255, Math.round(parseFloat(v) * 255))); return "rgba(" + k(ms[1]) + ", " + k(ms[2]) + ", " + k(ms[3]) + ", " + round(ms[4] === undefined ? 1 : parseFloat(ms[4]), 3) + ")"; }
    return "rgba(" + d[0] + ", " + d[1] + ", " + d[2] + ", " + a + ")";       // image data is not premultiplied
  } catch (e) { return null; }
}
function lowerColors(value) {
  if (!/#[0-9a-f]{4}\b|#[0-9a-f]{8}\b|\b(rgba?|hsla?)\([^(),]+\)|\b(oklch|oklab|lch|lab|hwb|color|color-mix|light-dark)\(/i.test(value)) { return value; }
  const tree = valueParser(value);
  tree.walk(node => {
    if (node.type === "word" && /^#([0-9a-f]{4}|[0-9a-f]{8})$/i.test(node.value)) {
      let h = node.value.slice(1); if (h.length === 4) { h = h.replace(/./g, "$&$&"); }
      const n = parseInt(h.slice(0, 6), 16);
      node.value = "rgba(" + (n >> 16) + ", " + ((n >> 8) & 255) + ", " + (n & 255) + ", " + round(parseInt(h.slice(6), 16) / 255, 3) + ")";
    } else if (node.type === "function" && /^(rgba?|hsla?|oklch|oklab|lch|lab|hwb|color|color-mix|light-dark)$/i.test(node.value)) {
      const text = valueParser.stringify(node);
      if (/^(rgba?|hsla?)$/i.test(node.value) && text.indexOf(",") >= 0 && !(/^(rgb|hsl)$/i.test(node.value) && splitTop(text.slice(text.indexOf("(") + 1, -1), ",").length === 4)) { return false; }
      let t = text;
      if (/^light-dark$/i.test(node.value)) { t = splitTop(text.slice(text.indexOf("(") + 1, -1), ",")[0] || text; }
      const c = /var\(/.test(t) ? null : resolveColor(t);
      if (c) { node.type = "word"; node.value = c; }
      return false;
    }
  });
  return tree.toString();
}

// ------------------------------------------------------------------ media queries against the fixed canvas
function mediaAnswer(params, ctx, lengths) {     // true / false / null (cannot tell: leave the rule alone)
  const one = q => {
    q = q.trim().toLowerCase(); if (!q) { return true; }
    let neg = false, res = true;
    const m0 = /^(not|only)\s+/.exec(q); if (m0) { neg = m0[1] === "not"; q = q.slice(m0[0].length); }
    const parts = q.split(/\s+and\s+/);
    for (let part of parts) {
      part = part.trim();
      let r;
      if (!part.startsWith("(")) { r = part === "all" || part === "screen" ? true : (part === "print" || part === "speech" ? false : null); }
      else {
        const inner = part.replace(/^\(|\)$/g, "").trim();
        const len = s => { const e = lengths.evaluate(s.replace(/\b([\d.]+)em\b/g, (a, n) => (n * 16) + "px")); return e ? e.v : NaN; };
        let m;
        if ((m = /^(min-|max-)?(width|height)\s*:\s*(.+)$/.exec(inner))) {
          const have = m[2] === "width" ? ctx.W : ctx.H, want = len(m[3]);
          r = isNaN(want) ? null : (m[1] === "min-" ? have >= want : m[1] === "max-" ? have <= want : have === want);
        } else if ((m = /^(width|height)\s*(<=|>=|<|>|=)\s*(.+)$/.exec(inner))) {
          const have = m[1] === "width" ? ctx.W : ctx.H, want = len(m[3]);
          r = isNaN(want) ? null : (m[2] === "<=" ? have <= want : m[2] === ">=" ? have >= want : m[2] === "<" ? have < want : m[2] === ">" ? have > want : have === want);
        } else if ((m = /^(.+?)\s*(<=|<)\s*(width|height)\s*(<=|<)\s*(.+)$/.exec(inner))) {
          const have = m[3] === "width" ? ctx.W : ctx.H, lo = len(m[1]), hi = len(m[5]);
          r = isNaN(lo) || isNaN(hi) ? null : ((m[2] === "<=" ? lo <= have : lo < have) && (m[4] === "<=" ? have <= hi : have < hi));
        } else if ((m = /^orientation\s*:\s*(\w+)$/.exec(inner))) { r = (ctx.W >= ctx.H ? "landscape" : "portrait") === m[1]; }
        else if ((m = /^(min-|max-)?aspect-ratio\s*:\s*([\d.]+)\s*(?:\/\s*([\d.]+))?$/.exec(inner))) {
          const have = ctx.W / ctx.H, want = parseFloat(m[2]) / (m[3] ? parseFloat(m[3]) : 1);
          r = m[1] === "min-" ? have >= want : m[1] === "max-" ? have <= want : Math.abs(have - want) < 0.001;
        } else if ((m = /^prefers-color-scheme\s*:\s*(\w+)$/.exec(inner))) { r = m[1] === "light"; }
        else if ((m = /^prefers-reduced-motion\s*:\s*([\w-]+)$/.exec(inner))) { r = m[1] === "no-preference"; }
        else if (/^prefers-reduced-motion$/.test(inner)) { r = false; }
        else if ((m = /^prefers-contrast\s*:\s*([\w-]+)$/.exec(inner))) { r = m[1] === "no-preference"; }
        else if ((m = /^(any-)?hover\s*:\s*(\w+)$/.exec(inner))) { r = m[2] === "hover"; }
        else if (/^(any-)?hover$/.test(inner)) { r = true; }
        else if ((m = /^(any-)?pointer\s*:\s*(\w+)$/.exec(inner))) { r = m[2] === "fine"; }
        else { r = null; }
      }
      if (r === null) { return null; }
      if (!r) { res = false; }
    }
    return neg ? !res : res;
  };
  let any = false, unknown = false;
  for (const q of splitTop(params, ",")) { const r = one(q); if (r === null) { unknown = true; } else if (r) { any = true; } }
  if (any) { return true; }
  return unknown ? null : false;
}

// ------------------------------------------------------------------ grid templates -> data for the page-side adapter
function parseTracks(text, lengths, warn) {
  // returns { t: [track, ...] } where a track is a string for -ms-grid-columns, or { auto: "fit"|"fill", min: px, t: [tracks] }
  text = text.replace(/\[[^\]]*\]/g, " ").trim();
  if (!text || text === "none") { return null; }
  if (/\b(subgrid|masonry)\b/.test(text)) { warn("grid: subgrid / masonry tracks are not supported"); return null; }
  const out = [];
  const one = tok => {
    tok = tok.trim(); if (!tok) { return; }
    let m = /^repeat\(\s*([^,]+),([\s\S]+)\)$/i.exec(tok);
    if (m) {
      const inner = []; for (const x of splitTop(m[2].trim(), " ")) { const k = single(x); if (k) { inner.push(k); } }
      const n = m[1].trim().toLowerCase();
      if (n === "auto-fit" || n === "auto-fill") {
        let min = 0;
        for (const x of splitTop(m[2].trim(), " ")) {
          const mm = /^minmax\(\s*([^,]+),/i.exec(x); const e = lengths.evaluate(mm ? mm[1].trim() : x);
          min += e && e.d ? e.v : 0;
        }
        out.push({ auto: n === "auto-fit" ? "fit" : "fill", min: round(min, 2), t: inner });
      } else { const cnt = Math.max(1, parseInt(n, 10) || 1); for (let i = 0; i < cnt; i++) { for (const k of inner) { out.push(k); } } }
      return;
    }
    const k = single(tok); if (k) { out.push(k); }
  };
  const single = tok => {
    tok = lengths.lower(tok.trim());
    if (/^fit-content\(/i.test(tok)) { return "auto"; }
    let m = /^minmax\(\s*([^,]+),\s*([\s\S]+)\)$/i.exec(tok);
    if (m) {
      let a = lenOrKeep(m[1].trim()), b = lenOrKeep(m[2].trim());
      if (a === "0") { a = "0px"; }
      if (a === "auto") { a = "min-content"; }
      if (b === "auto") { b = "max-content"; }
      return "minmax(" + a + ", " + b + ")";
    }
    return lenOrKeep(tok);
  };
  const lenOrKeep = v => {
    if (/^(auto|min-content|max-content)$/i.test(v)) { return v.toLowerCase(); }
    if (/^[\d.]+fr$/i.test(v)) { return v.toLowerCase(); }
    const m = /^(-?[\d.]+)rem$/i.exec(v); if (m) { return px(parseFloat(m[1]) * lengths.unit.rem); }
    return v;
  };
  for (const tok of splitTop(text, " ")) { one(tok); }
  return out.length ? { t: out } : null;
}
function parseAreas(text) {
  const rows = []; const rx = /"([^"]*)"|'([^']*)'/g; let m;
  while ((m = rx.exec(text))) { rows.push((m[1] !== undefined ? m[1] : m[2]).trim().split(/\s+/)); }
  if (!rows.length) { return null; }
  const map = {};
  rows.forEach((cells, r) => cells.forEach((name, c) => {
    if (/^\.+$/.test(name)) { return; }
    const a = map[name] || (map[name] = [r + 1, c + 1, r + 2, c + 2]);   // row start, col start, row end, col end (lines)
    a[0] = Math.min(a[0], r + 1); a[1] = Math.min(a[1], c + 1); a[2] = Math.max(a[2], r + 2); a[3] = Math.max(a[3], c + 2);
  }));
  return { map, rows: rows.length, cols: Math.max.apply(null, rows.map(r => r.length)) };
}
function parseLine(text) {            // "2", "span 3", "auto", "-1", "name" -> number | {span:n} | null | {name}
  text = text.trim().toLowerCase();
  if (!text || text === "auto") { return null; }
  let m = /^span\s+(\d+)$/.exec(text) || /^(\d+)\s+span$/.exec(text); if (m) { return { span: parseInt(m[1], 10) }; }
  if (/^-?\d+$/.test(text)) { return parseInt(text, 10); }
  return { name: text };
}
function parsePlacement(text) { const p = splitTop(text, "/"); return [parseLine(p[0] || ""), p.length > 1 ? parseLine(p[1]) : undefined]; }

const VIRTUAL = { "gap": 1, "row-gap": 1, "column-gap": 1, "grid-gap": 1, "grid-row-gap": 1, "grid-column-gap": 1, "grid-template-columns": 1, "grid-template-rows": 1,
  "grid-template-areas": 1, "grid-template": 1, "grid": 1, "grid-auto-rows": 1, "grid-auto-columns": 1, "grid-auto-flow": 1, "grid-column": 1, "grid-row": 1, "grid-area": 1,
  "grid-column-start": 1, "grid-column-end": 1, "grid-row-start": 1, "grid-row-end": 1, "justify-items": 1, "justify-self": 1, "aspect-ratio": 1, "object-fit": 1, "object-position": 1, "min-width": 1, "min-height": 1 };

function virtualProps(decls, lengths, warn) {      // decls: [{prop, value}] of one rule -> object for the adapter (or null)
  const v = {}; let any = false;
  const set = (k, val) => { if (val !== null && val !== undefined) { v[k] = val; any = true; } };
  const gapLen = s => { s = lengths.lower(s.trim()); const m = /^(-?[\d.]+)rem$/i.exec(s); return m ? px(parseFloat(m[1]) * lengths.unit.rem) : (s === "normal" ? "0px" : s); };
  for (const d of decls) {
    const val = d.value.replace(/\s*!important\s*$/i, "").trim();
    switch (d.prop) {
      case "gap": case "grid-gap": { const p = splitTop(val, " "); set("rg", gapLen(p[0])); set("cg", gapLen(p[1] !== undefined ? p[1] : p[0])); break; }
      case "row-gap": case "grid-row-gap": set("rg", gapLen(val)); break;
      case "column-gap": case "grid-column-gap": set("cg", gapLen(val)); break;
      case "grid-template-columns": set("cols", parseTracks(val, lengths, warn) || 0); break;
      case "grid-template-rows": set("rows", parseTracks(val, lengths, warn) || 0); break;
      case "grid-template-areas": set("areas", parseAreas(val) || 0); break;
      case "grid-template": case "grid": {
        if (/auto-flow/.test(val)) { warn("grid: the 'grid' shorthand with auto-flow is not converted"); break; }
        const p = splitTop(val, "/");
        if (/["']/.test(p[0])) {
          set("areas", parseAreas(p[0]) || 0);
          const sizes = p[0].replace(/"[^"]*"|'[^']*'/g, " | ").split("|").map(s => s.trim());
          const rows = []; const n = (p[0].match(/"[^"]*"|'[^']*'/g) || []).length;
          for (let i = 1; i <= n; i++) { rows.push(sizes[i] || "auto"); }
          set("rows", parseTracks(rows.join(" "), lengths, warn) || 0);
        } else { set("rows", parseTracks(p[0] || "", lengths, warn) || 0); }
        if (p.length > 1) { set("cols", parseTracks(p[1], lengths, warn) || 0); }
        break;
      }
      case "grid-auto-rows": set("arows", (parseTracks(val, lengths, warn) || { t: ["auto"] }).t.filter(x => typeof x === "string")); break;
      case "grid-auto-columns": set("acols", (parseTracks(val, lengths, warn) || { t: ["auto"] }).t.filter(x => typeof x === "string")); break;
      case "grid-auto-flow": set("flow", /column/.test(val) ? "column" : "row"); break;
      case "grid-column": { const p = parsePlacement(val); set("cs", p[0] === null ? 0 : p[0]); set("ce", p[1] === undefined || p[1] === null ? 0 : p[1]); break; }
      case "grid-row": { const p = parsePlacement(val); set("rs", p[0] === null ? 0 : p[0]); set("re", p[1] === undefined || p[1] === null ? 0 : p[1]); break; }
      case "grid-column-start": set("cs", parseLine(val) || 0); break;
      case "grid-column-end": set("ce", parseLine(val) || 0); break;
      case "grid-row-start": set("rs", parseLine(val) || 0); break;
      case "grid-row-end": set("re", parseLine(val) || 0); break;
      case "grid-area": {
        const p = splitTop(val, "/");
        if (p.length === 1 && !/^(auto|-?\d+|span\s)/i.test(p[0])) { set("area", p[0].trim()); }
        else { set("rs", parseLine(p[0] || "") || 0); set("cs", parseLine(p[1] || "") || 0); set("re", parseLine(p[2] || "") || 0); set("ce", parseLine(p[3] || "") || 0); }
        break;
      }
      case "justify-items": set("ji", val.replace(/^(safe|unsafe|legacy)\s+/, "")); break;
      case "justify-self": set("js", val.replace(/^(safe|unsafe)\s+/, "")); break;
      case "aspect-ratio": {
        const m = /([\d.]+)\s*(?:\/\s*([\d.]+))?/.exec(val.replace(/\bauto\b/g, ""));
        set("ratio", m ? round(parseFloat(m[1]) / (m[2] ? parseFloat(m[2]) : 1), 5) : 0);
        break;
      }
      case "min-width": set("mw0", 1); break;
      case "min-height": set("mh0", 1); break;
      case "object-fit": set("fit", val); break;
      case "object-position": set("fpos", val); break;
    }
  }
  return any ? v : null;
}

// ------------------------------------------------------------------ selectors
const IE_PSEUDO = { hover: 1, active: 1, focus: 1, visited: 1, link: 1, "first-child": 1, "last-child": 1, "nth-child": 1, "nth-last-child": 1, "nth-of-type": 1, "nth-last-of-type": 1,
  "first-of-type": 1, "last-of-type": 1, "only-child": 1, "only-of-type": 1, not: 1, root: 1, empty: 1, target: 1, checked: 1, disabled: 1, enabled: 1, indeterminate: 1, before: 1, after: 1,
  "first-line": 1, "first-letter": 1, selection: 1, lang: 1 };
const DYNAMIC = /:(hover|active|focus|visited|target|checked|disabled|enabled)\b|::?(before|after|first-line|first-letter|selection|placeholder|marker)\b/i;

function expandSelector(sel) {           // :is() / :where() -> plain lists, :focus-visible -> :focus, :not(a, b) -> :not(a):not(b)
  if (!/:(is|where|matches|-webkit-any|focus-visible|not)\b/i.test(sel)) { return [sel]; }
  let list = [sel], guard = 0;
  while (guard++ < 8) {
    let changed = false; const next = [];
    for (const s of list) {
      const m = /:(is|where|matches|-webkit-any)\(/i.exec(s);
      if (!m) { next.push(s); continue; }
      let depth = 0, end = -1;
      for (let i = m.index + m[0].length - 1; i < s.length; i++) { if (s[i] === "(") { depth++; } else if (s[i] === ")") { depth--; if (depth === 0) { end = i; break; } } }
      if (end < 0) { next.push(s); continue; }
      const head = s.slice(0, m.index), tail = s.slice(end + 1), alts = splitTop(s.slice(m.index + m[0].length, end), ",");
      for (const a of alts) {
        // a type selector must come first inside a compound: "p:is(.a)" stays "p.a", ".x:is(p)" becomes "p.x"
        const compoundStart = Math.max(head.lastIndexOf(" "), head.lastIndexOf(">"), head.lastIndexOf("+"), head.lastIndexOf("~")) + 1;
        const comp = head.slice(compoundStart);
        if (/^[a-z*]/i.test(a) && comp && !/[\s>+~]/.test(a)) {
          const mt = /^([a-z][\w-]*|\*)/i.exec(a);
          next.push(head.slice(0, compoundStart) + mt[0] + comp + a.slice(mt[0].length) + tail);
        } else { next.push(head + a + tail); }
      }
      changed = true;
    }
    list = next; if (!changed) { break; }
  }
  return list.map(s => s.replace(/:focus-visible\b/gi, ":focus").replace(/:not\(([^()]*,[^()]*)\)/gi, (all, inner) => splitTop(inner, ",").map(x => ":not(" + x + ")").join("")));
}
function unknownToIe(sel) {
  let bad = false;
  sel.replace(/::?([a-z-]+)/gi, (all, name) => { if (!IE_PSEUDO[name.toLowerCase()] && !/^-ms-/i.test(name)) { bad = true; } return all; });
  return bad;
}
function specificity(sel) {
  let a = 0, b = 0, c = 0;
  try {
    selectorParser(root => {
      root.walk(n => {
        if (n.type === "id") { a++; } else if (n.type === "class" || n.type === "attribute") { b++; }
        else if (n.type === "pseudo") { if (/^::|^:(before|after|first-line|first-letter)$/i.test(n.value)) { c++; } else if (!/^:(not|is|where|matches)$/i.test(n.value)) { b++; } }
        else if (n.type === "tag") { c++; }
      });
    }).processSync(sel);
  } catch (e) { }
  return a * 10000 + b * 100 + c;
}

// ------------------------------------------------------------------ the PostCSS plugins
const LOGICAL = { "margin-inline": ["margin-left", "margin-right"], "margin-block": ["margin-top", "margin-bottom"], "padding-inline": ["padding-left", "padding-right"], "padding-block": ["padding-top", "padding-bottom"],
  "inset-inline": ["left", "right"], "inset-block": ["top", "bottom"] };
const LOGICAL1 = { "margin-inline-start": "margin-left", "margin-inline-end": "margin-right", "margin-block-start": "margin-top", "margin-block-end": "margin-bottom",
  "padding-inline-start": "padding-left", "padding-inline-end": "padding-right", "padding-block-start": "padding-top", "padding-block-end": "padding-bottom",
  "inset-inline-start": "left", "inset-inline-end": "right", "inset-block-start": "top", "inset-block-end": "bottom", "inline-size": "width", "block-size": "height",
  "min-inline-size": "min-width", "max-inline-size": "max-width", "min-block-size": "min-height", "max-block-size": "max-height",
  "border-inline-start": "border-left", "border-inline-end": "border-right", "border-block-start": "border-top", "border-block-end": "border-bottom" };

function fixedViewport(ctx, lengths, notes) {
  return {
    postcssPlugin: "ih-fixed-viewport",
    Once(root) {
      let again = true, guard = 0;
      while (again && guard++ < 6) {
        again = false;
        root.walkAtRules(at => {
          const name = at.name.toLowerCase();
          if (name === "media") {
            const r = mediaAnswer(at.params, ctx, lengths);
            if (r === true) { at.replaceWith(at.nodes || []); again = true; }
            else if (r === false) { at.remove(); }
          } else if (name === "supports") {
            if (/^\s*not\b/i.test(at.params)) { at.remove(); } else { at.replaceWith(at.nodes || []); again = true; }
          } else if (name === "layer") {
            if (at.nodes) { at.replaceWith(at.nodes); again = true; } else { at.remove(); }
          } else if (name === "property") { at.remove(); }
        });
      }
    }
  };
}

function lowerDeclarations(ctx, lengths, state) {
  const warn = m => { if (state.warnings.indexOf(m) < 0) { state.warnings.push(m); } };
  const note = m => { state.counts[m] = (state.counts[m] || 0) + 1; };
  return {
    postcssPlugin: "ih-lower",
    Once(root) {
      // 1. selectors
      root.walkRules(rule => {
        if (rule.parent && rule.parent.type === "atrule" && /keyframes$/i.test(rule.parent.name)) { return; }
        let sels = [];
        for (const s of rule.selectors) { const ex = expandSelector(s); if (ex.length !== 1 || ex[0] !== s) { note("selectors rewritten (:is / :where / :focus-visible)"); } sels = sels.concat(ex); }
        const good = sels.filter(s => !unknownToIe(s)), bad = sels.filter(s => unknownToIe(s));
        if (good.length && bad.length) { rule.cloneAfter({ selectors: bad }); rule.selectors = good; }
        else { rule.selectors = sels; }
      });
      // 2. declarations
      root.walkDecls(d => {
        const prop = d.prop.toLowerCase();
        if (prop.indexOf("--") === 0) { return; }
        let value = d.value;
        if (prop === "inset") {
          const p = splitTop(value, " "); const t = p[0], r = p[1] !== undefined ? p[1] : t, b = p[2] !== undefined ? p[2] : t, l = p[3] !== undefined ? p[3] : r;
          d.cloneBefore({ prop: "top", value: t }); d.cloneBefore({ prop: "right", value: r }); d.cloneBefore({ prop: "bottom", value: b }); d.cloneBefore({ prop: "left", value: l });
          d.remove(); note("inset written out"); return;
        }
        if (LOGICAL[prop]) { const p = splitTop(value, " "); d.cloneBefore({ prop: LOGICAL[prop][0], value: p[0] }); d.cloneBefore({ prop: LOGICAL[prop][1], value: p[1] !== undefined ? p[1] : p[0] }); d.remove(); note("logical properties written out"); return; }
        if (LOGICAL1[prop]) { d.prop = LOGICAL1[prop]; note("logical properties written out"); }
        if (prop === "place-items" || prop === "place-content" || prop === "place-self") {
          const p = splitTop(value, " "), kind = prop.slice(6);
          d.cloneBefore({ prop: "align-" + kind, value: p[0] }); d.cloneBefore({ prop: "justify-" + kind, value: p[1] !== undefined ? p[1] : p[0] }); d.remove(); note("place-* written out"); return;
        }
        if (/^(align|justify)-(content|items|self)$/.test(prop) && /^(start|end)$/i.test(value.trim())) { d.cloneBefore({ prop: prop, value: "flex-" + value.trim().toLowerCase() }); note("start / end alignment given its older name as well"); }
        if (prop === "display" && /^(inline-)?grid$/i.test(value.trim())) { d.cloneBefore({ prop: "display", value: /^inline/i.test(value.trim()) ? "-ms-inline-grid" : "-ms-grid" }); state.usesGrid = true; return; }
        if (prop === "position" && /^sticky$/i.test(value.trim())) { d.cloneBefore({ prop: "position", value: "relative" }); warn("position: sticky acts as position: relative in the slide show (a slide does not scroll)"); return; }
        if (prop === "backdrop-filter" || prop === "-webkit-backdrop-filter") { warn("backdrop-filter has no effect in the slide show (IE11 has no equivalent): the blur behind translucent panels is missing there"); return; }
        if (prop === "mix-blend-mode" || prop === "background-blend-mode") { warn(prop + " has no effect in the slide show (IE11 has no equivalent)"); return; }
        if (prop === "filter" && /\b(blur|brightness|contrast|drop-shadow|grayscale|hue-rotate|invert|opacity|saturate|sepia)\(/i.test(value)) { warn("CSS filter (" + value.trim().slice(0, 40) + ") has no effect in the slide show (IE11 has no equivalent)"); return; }
        if (prop === "user-select" && !d.parent.some(x => x.prop === "-ms-user-select")) { d.cloneBefore({ prop: "-ms-user-select", value: value }); }
        const v1 = lengths.lower(value); if (v1 !== value) { note("clamp / min / max / viewport units settled for the canvas"); value = v1; }
        const v2 = lowerColors(value); if (v2 !== value) { note("colour notations rewritten"); value = v2; }
        if (value !== d.value) { d.value = value; }
      });
      // 3. what IE cannot express at all goes to the page-side adapter
      root.walkRules(rule => {
        if (rule.parent && rule.parent.type === "atrule" && /keyframes$/i.test(rule.parent.name)) { return; }
        const picked = []; const multicol = rule.some(x => x.type === "decl" && /^(columns|column-count|column-width)$/i.test(x.prop));
        rule.each(d => {
          if (d.type !== "decl") { return; }
          const prop = d.prop.toLowerCase();
          if (!VIRTUAL[prop]) { return; }
          if (prop === "column-gap" && multicol) { return; }
          if ((prop === "min-width" || prop === "min-height") && !/^0(px|%)?$/.test(d.value.trim())) { return; }
          picked.push({ prop, value: d.value, important: d.important, node: d });
        });
        if (!picked.length) { return; }
        const v = virtualProps(picked, lengths, warn);
        if (!v) { return; }
        const imp = picked.some(p => p.important) ? 1 : 0;
        for (const sel of rule.selectors) {
          if (DYNAMIC.test(sel)) { warn("layout properties in a rule with :hover / ::before etc. are not converted: " + sel); continue; }
          state.table.push({ s: sel, p: v, w: imp * 1000000 + specificity(sel), o: state.order++ });
        }
        for (const k in v) { if (k === "mw0" || k === "mh0") { continue; } if (k === "rg" || k === "cg") { state.usesGap = true; } else if (k === "ratio") { state.usesRatio = true; } else if (k === "fit" || k === "fpos") { state.usesFit = true; } }
      });
    }
  };
}

// ------------------------------------------------------------------ scoped CSS variables -> override rules
function splitCompounds(sel) {           // "a > b c" -> [{ comb: "", text: "a" }, { comb: " > ", text: "b" }, { comb: " ", text: "c" }]
  const parts = []; let depth = 0, cur = "", comb = "", q = null;
  const push = () => { if (cur) { parts.push({ comb, text: cur }); cur = ""; comb = ""; } };
  for (let i = 0; i < sel.length; i++) {
    const c = sel[i];
    if (q) { cur += c; if (c === q) { q = null; } continue; }
    if (c === '"' || c === "'") { q = c; cur += c; continue; }
    if (c === "(" || c === "[") { depth++; } else if (c === ")" || c === "]") { depth--; }
    if (depth === 0 && (/\s/.test(c) || c === ">" || c === "+" || c === "~")) { push(); comb += c; continue; }
    cur += c;
  }
  push();
  for (const p of parts) { p.comb = p.comb.trim() ? " " + p.comb.trim() + " " : (p.comb ? " " : ""); }
  return parts;
}
function mergeOne(comp, scope) {         // the element matched by comp is also the scope element
  const pe = /(::[\w-]+(\([^)]*\))?|:(before|after|first-line|first-letter))$/i.exec(comp);
  let base = comp, tail = "";
  if (pe) { base = comp.slice(0, pe.index); tail = pe[0]; }
  const st = /^([a-z][\w-]*|\*)/i.exec(scope), bt = /^([a-z][\w-]*|\*)/i.exec(base);
  let rest = scope;
  if (st) {
    if (bt && bt[0] !== "*" && st[0] !== "*" && bt[0].toLowerCase() !== st[0].toLowerCase()) { return null; }
    rest = scope.slice(st[0].length);
    if (!bt) { base = st[0] + base; }
  }
  return base + rest + tail;
}
function combineScope(scope, sel) {      // selectors under which a rule for "sel" sees the variables of "scope"
  scope = scope.trim(); sel = sel.trim();
  const single = splitCompounds(scope).length === 1;
  const parts = splitCompounds(sel), out = [];
  if (!parts.length) { return out; }
  const merged = where => {
    if (!single) { return null; }
    const copy = parts.map(p => ({ comb: p.comb, text: p.text })), i = where === "first" ? 0 : copy.length - 1;
    const t = mergeOne(copy[i].text, scope); if (!t) { return null; }
    copy[i].text = t;
    return copy.map(p => p.comb + p.text).join("");
  };
  if (!/^(:root|html)(?![\w-])/i.test(parts[0].text)) { out.push(scope + " " + sel); }
  const a = merged("first"); if (a && out.indexOf(a) < 0) { out.push(a); }
  if (parts.length > 1) { const b = merged("last"); if (b && out.indexOf(b) < 0) { out.push(b); } }
  return out;
}
const VAR_RX = /var\(\s*(--[\w-]+)\s*(?:,\s*((?:[^()]|\([^()]*\))*))?\)/g;
function resolveVars(value, map, depth) {
  if (depth > 8 || value.indexOf("var(") < 0) { return value; }
  return value.replace(VAR_RX, (all, name, fb) => {
    const v = map[name];
    if (v !== undefined) { return resolveVars(String(v).trim(), map, depth + 1); }
    if (fb !== undefined) { return resolveVars(fb.trim(), map, depth + 1); }
    return all;
  });
}
function scopeVars(state) {
  const inKeyframes = rule => rule.parent && rule.parent.type === "atrule" && /keyframes$/i.test(rule.parent.name);
  const usesOf = v => { const out = []; String(v).replace(/var\(\s*(--[\w-]+)/g, (a, n) => { out.push(n); return a; }); return out; };
  return {
    postcssPlugin: "ih-scope-vars",
    Once(root) {
      const rootMap = {}, scoped = [];
      root.walkRules(rule => {
        if (inKeyframes(rule)) { return; }
        const decls = rule.nodes.filter(n => n.type === "decl" && n.prop.indexOf("--") === 0);
        if (!decls.length) { return; }
        for (const sel of rule.selectors) {
          if (/^(:root|html|body)$/i.test(sel.trim())) { for (const d of decls) { rootMap[d.prop] = d.value; } }
          else {
            let e = scoped.find(x => x.sel === sel.trim());
            if (!e) { e = { sel: sel.trim(), map: {} }; scoped.push(e); }
            for (const d of decls) { e.map[d.prop] = d.value; }
          }
        }
      });
      state.rootVars = rootMap;
      if (!scoped.length) { return; }
      const targets = [];
      root.walkRules(rule => { if (!inKeyframes(rule) && rule.some(n => n.type === "decl" && n.prop.indexOf("--") !== 0 && /var\(/.test(n.value))) { targets.push(rule); } });
      let made = 0;
      for (const sc of scoped) {
        const names = new Set(Object.keys(sc.map)); let grew = true, guard = 0;
        while (grew && guard++ < 10) { grew = false; for (const k in rootMap) { if (!names.has(k) && usesOf(rootMap[k]).some(n => names.has(n))) { names.add(k); grew = true; } } }
        const map = Object.assign({}, rootMap, sc.map);
        for (const rule of targets) {
          const decls = rule.nodes.filter(n => n.type === "decl" && n.prop.indexOf("--") !== 0 && usesOf(n.value).some(x => names.has(x)));
          if (!decls.length) { continue; }
          const sels = [];
          for (const sel of rule.selectors) { for (const v of combineScope(sc.sel, sel)) { if (sels.indexOf(v) < 0) { sels.push(v); } } }
          if (!sels.length) { continue; }
          const nr = postcss.rule({ selectors: sels });
          for (const d of decls) { nr.append(postcss.decl({ prop: d.prop, value: resolveVars(d.value, map, 0), important: d.important })); }
          (rule.ihLast || rule).after(nr); rule.ihLast = nr; made++;
        }
      }
      if (made) { state.counts["CSS variables that a class or attribute overrides (themes, variants) written out as " + made + " override rules"] = 1; }
    }
  };
}

function rootFontSize(css) {
  let size = 16;
  const m = /(?:^|[}\s,])(?:html|:root)\s*\{[^}]*?font-size\s*:\s*([\d.]+)(px|%)/i.exec(css);
  if (m) { size = m[2] === "%" ? 16 * parseFloat(m[1]) / 100 : parseFloat(m[1]); }
  return size;
}

function buildState() { return { table: [], order: 0, warnings: [], counts: {}, usesGrid: false, usesGap: false, usesRatio: false, usesFit: false }; }

export function convertCss(css, opts) {
  const ctx = { W: opts.canvasW || 1280, H: opts.canvasH || 720, rem: rootFontSize(css) };
  const lengths = makeLengths(ctx), state = opts.state || buildState();
  const notes = [];
  // one pass per step: plugins of one pass do not run strictly one after the other
  const step = (plugins, text) => postcss(plugins).process(text, { from: undefined });
  let text = step([nesting()], css).css;
  text = step([fixedViewport(ctx, lengths, notes)], text).css;
  text = step([scopeVars(state)], text).css;
  const vars = step([cssVariables({ preserve: false, preserveAtRulesOrder: true })], text);
  for (const w of vars.warnings()) { if (state.warnings.indexOf(w.text) < 0) { state.warnings.push(w.text); } }
  text = vars.css;
  text = step([lowerDeclarations(ctx, lengths, state)], text).css;
  const result = step([flexbugs(), autoprefixer({ overrideBrowserslist: ["ie 11"], grid: false, flexbox: "no-2009", remove: false })], text);
  const out = result.css;
  for (const w of result.warnings()) { const t = w.text; if (!/^variable '--/.test(t) || true) { if (state.warnings.indexOf(t) < 0) { state.warnings.push(t); } } }
  return { css: out, state };
}

export function convertInlineStyle(style, opts, state) {     // one style="" attribute -> { style, v }
  const ctx = { W: opts.canvasW || 1280, H: opts.canvasH || 720, rem: opts.rem || 16 };
  const lengths = makeLengths(ctx); const local = buildState();
  let root;
  if (state.rootVars && style.indexOf("var(") >= 0) { style = resolveVars(style, state.rootVars, 0); }
  try { root = postcss.parse("a{" + style + "}"); } catch (e) { return { style, v: null }; }
  const plugin = lowerDeclarations(ctx, lengths, local);
  plugin.Once(root);
  for (const w of local.warnings) { if (state.warnings.indexOf(w) < 0) { state.warnings.push(w); } }
  for (const k in local.counts) { state.counts[k] = (state.counts[k] || 0) + local.counts[k]; }
  if (local.usesGrid) { state.usesGrid = true; } if (local.usesGap) { state.usesGap = true; } if (local.usesRatio) { state.usesRatio = true; } if (local.usesFit) { state.usesFit = true; }
  let text = "";
  root.walkDecls(d => { text += (text ? "; " : "") + d.prop + ": " + d.value + (d.important ? " !important" : ""); });
  let v = null;
  for (const e of local.table) { v = Object.assign(v || {}, e.p); }
  return { style: text, v };
}

// ------------------------------------------------------------------ JavaScript
export function isEs5(code) {
  try { acorn.parse(code, { ecmaVersion: 5, sourceType: "script", allowReturnOutsideFunction: true, allowHashBang: true }); return true; } catch (e) { return false; }
}
export function convertJs(code, Babel) {
  if (isEs5(code)) { return { code: null }; }
  try {
    const out = Babel.transform(code, {
      sourceType: "script", compact: false, comments: true, babelrc: false, configFile: false,
      presets: [["env", { targets: { ie: "11" }, modules: false }]],
      assumptions: { arrayLikeIsIterable: true, noDocumentAll: true, setPublicClassFields: true, constantSuper: true, noClassCalls: true, setClassMethods: true, objectRestNoSymbols: true,
        setSpreadProperties: true, pureGetters: true, ignoreFunctionLength: true, noNewArrows: true, privateFieldsAsProperties: true, superIsCallableConstructor: true,
        skipForOfIteratorClosing: true, mutableTemplateObject: true, ignoreToPrimitiveHint: true }
    });
    return { code: out.code };
  } catch (e) {
    return { code: null, error: String(e && e.message ? e.message : e).split("\n")[0].slice(0, 300) };
  }
}

// ------------------------------------------------------------------ one page
// job: { canvasW, canvasH, scripts: [{ code, module }], styles: [css, ...], inline: [style attribute, ...] }
export function convertPage(job, Babel) {
  const state = buildState();
  const res = { scripts: [], css: null, inline: [], table: [], notes: [], warnings: [], uses: {} };
  const opts = { canvasW: job.canvasW, canvasH: job.canvasH };
  const all = (job.styles || []).join("\n");
  if (all.trim()) {
    try { res.css = convertCss(all, { canvasW: opts.canvasW, canvasH: opts.canvasH, state }).css; }
    catch (e) { res.css = null; res.warnings.push("The style sheets could not be parsed, so they were left as they are: " + String(e && e.message ? e.message : e).split("\n")[0].slice(0, 200)); }
  }
  opts.rem = rootFontSize(all);
  for (const s of (job.inline || [])) {
    if (!/(var\(|gap|grid|inset|aspect-ratio|object-(fit|position)|place-|clamp\(|min\(|max\(|vmin|vmax|[dsl]v[wh]|#[0-9a-f]{8}|#[0-9a-f]{4}\b|-inline|-block|backdrop|blend|oklch|color-mix|sticky|\b(rgb|hsl)a?\([^(),]+\))/i.test(s)) { res.inline.push(null); continue; }
    const r = convertInlineStyle(s, opts, state);
    res.inline.push(r.style === s && !r.v ? null : r);
  }
  let converted = 0;
  for (const s of (job.scripts || [])) {
    const r = convertJs(s.code, Babel);
    if (r.error) { res.warnings.push("A script could not be converted (" + r.error + "); it was left as it is."); }
    if (r.code !== null && r.code !== undefined) { converted++; }
    res.scripts.push(r.code === undefined ? null : r.code);
  }
  if (converted) { res.notes.push(converted + " script block(s) rewritten to the JavaScript of IE11 (arrow functions, let / const, classes, template strings, async ...)"); }
  const kinds = Object.keys(state.counts);
  if (kinds.length) { res.notes.push("style sheets: " + kinds.join("; ")); }
  if (state.usesGrid) { res.notes.push("CSS grid is laid out by the page-side adapter (IE11's own grid, items placed by script)"); }
  if (state.usesGap) { res.notes.push("gap is applied by the page-side adapter (as grid tracks / margins)"); }
  state.table.sort((a, b) => a.w - b.w || a.o - b.o);
  res.allow = res.css !== null || (job.inline || []).length ? ["css-grid", "css-gap", "css-aspect-ratio", "css-object-fit", "css-sticky", "css-backdrop", "css-filter", "css-blend", "css-place", "css-inset"] : [];
  res.table = state.table.map(t => ({ s: t.s, p: t.p })); res.warnings = res.warnings.concat(state.warnings);
  res.uses = { grid: state.usesGrid, gap: state.usesGap, ratio: state.usesRatio, fit: state.usesFit };
  return res;
}
