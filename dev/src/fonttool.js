// Font subsetting for the build (runs in Edge headless, or in Node for the dev-time font pack):
// HarfBuzz's subsetter (WebAssembly) cuts a font down to the characters a page uses and pins a variable font
// to one weight; the result is packed as WOFF (the format the IE11 engine and every other browser reads).
// Bundled by lab\tools\build.js into assets\tools\fonttool.js (global IHFont).
import pako from "pako";
import wasmB64 from "harfbuzzjs/dist/harfbuzz-subset.wasm";

let hb = null;
const fonts = {};     // key -> Uint8Array (sfnt)

function b64ToBytes(b64) {
  const bin = atob(b64), out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) { out[i] = bin.charCodeAt(i); }
  return out;
}
function bytesToB64(bytes) {
  let s = ""; const CH = 0x8000;
  for (let i = 0; i < bytes.length; i += CH) { s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH)); }
  return btoa(s);
}
const TAG = t => t.split("").reduce((a, ch) => (a << 8) + ch.charCodeAt(0), 0);

export async function init() {
  if (hb) { return "ready"; }
  const { instance } = await WebAssembly.instantiate(b64ToBytes(wasmB64));
  hb = instance.exports; hb._initialize();
  return "ready";
}
export function add(key, b64) { fonts[key] = b64ToBytes(b64); return fonts[key].length; }
export function addBytes(key, bytes) { fonts[key] = bytes; return bytes.length; }
export function has(key) { return !!fonts[key]; }
export function drop(key) { delete fonts[key]; return "ok"; }

// -> Uint8Array (sfnt). opts: { wght: number (pin a variable font), index: face index, keepAxes: true (do not instance), hinting: true }
export function subsetBytes(key, text, opts) {
  opts = opts || {};
  const font = fonts[key];
  if (!hb) { throw new Error("font tool not initialised"); }
  if (!font) { throw new Error("font not loaded: " + key); }
  const heap = () => new Uint8Array(hb.memory.buffer);
  const input = hb.hb_subset_input_create_or_fail();
  if (!input) { throw new Error("the subsetter could not start"); }
  const buf = hb.malloc(font.byteLength);
  heap().set(font, buf);
  const blob = hb.hb_blob_create(buf, font.byteLength, 2, 0, 0);
  const face = hb.hb_face_create(blob, opts.index || 0);
  hb.hb_blob_destroy(blob);
  let subset = 0, result = 0;
  try {
    const feats = hb.hb_subset_input_set(input, 6);
    hb.hb_set_clear(feats);
    for (const f of ["kern", "liga", "clig", "calt", "ccmp", "locl", "mark", "mkmk", "rlig"]) { hb.hb_set_add(feats, TAG(f)); }
    const names = hb.hb_subset_input_set(input, 4);          // name records: also the typographic family / style
    for (const id of [16, 17, 21, 22]) { hb.hb_set_add(names, id); }
    if (!opts.hinting) { hb.hb_subset_input_set_flags(input, hb.hb_subset_input_get_flags(input) | 1); }
    const set = hb.hb_subset_input_unicode_set(input);
    const seen = {};
    for (const ch of text) { const cp = ch.codePointAt(0); if (cp >= 32 && !seen[cp]) { seen[cp] = 1; hb.hb_set_add(set, cp); } }
    if (!opts.keepAxes) {
      if (hb.hb_subset_input_pin_all_axes_to_default) { hb.hb_subset_input_pin_all_axes_to_default(input, face); }
      if (typeof opts.wght === "number") { hb.hb_subset_input_pin_axis_location(input, face, TAG("wght"), opts.wght); }
    }
    subset = hb.hb_subset_or_fail(face, input);
    if (!subset) { throw new Error("the subsetter refused the font"); }
    result = hb.hb_face_reference_blob(subset);
    const off = hb.hb_blob_get_data(result, 0), len = hb.hb_blob_get_length(result);
    if (!len) { throw new Error("the subset came out empty"); }
    return heap().slice(off, off + len);
  } finally {
    hb.hb_subset_input_destroy(input);
    if (result) { hb.hb_blob_destroy(result); }
    if (subset) { hb.hb_face_destroy(subset); }
    hb.hb_face_destroy(face);
    hb.free(buf);
  }
}

// sfnt -> WOFF 1.0
export function toWoff(sfnt) {
  const dv = new DataView(sfnt.buffer, sfnt.byteOffset, sfnt.byteLength);
  const n = dv.getUint16(4);
  const tables = [];
  for (let i = 0; i < n; i++) {
    const p = 12 + i * 16;
    tables.push({ tag: dv.getUint32(p), sum: dv.getUint32(p + 4), off: dv.getUint32(p + 8), len: dv.getUint32(p + 12) });
  }
  tables.sort((a, b) => (a.tag >>> 0) - (b.tag >>> 0));
  let total = 44 + 20 * n, sfntSize = 12 + 16 * n;
  for (const t of tables) {
    const raw = sfnt.subarray(t.off, t.off + t.len);
    const z = pako.deflate(raw, { level: 9 });
    t.data = z.length < raw.length ? z : raw;
    total += (t.data.length + 3) & ~3; sfntSize += (t.len + 3) & ~3;
  }
  const out = new Uint8Array(total), o = new DataView(out.buffer);
  o.setUint32(0, 0x774F4646); o.setUint32(4, dv.getUint32(0)); o.setUint32(8, total); o.setUint16(12, n); o.setUint16(14, 0); o.setUint32(16, sfntSize);
  let pos = 44 + 20 * n;
  tables.forEach((t, i) => {
    const p = 44 + i * 20;
    o.setUint32(p, t.tag); o.setUint32(p + 4, pos); o.setUint32(p + 8, t.data.length); o.setUint32(p + 12, t.len); o.setUint32(p + 16, t.sum);
    out.set(t.data, pos); pos += (t.data.length + 3) & ~3;
  });
  return out;
}

// -> base64 WOFF
export function subset(key, text, wght, index) {
  const opts = { index: index || 0 }; if (typeof wght === "number" && wght > 0) { opts.wght = wght; }
  return bytesToB64(toWoff(subsetBytes(key, text, opts)));
}
export function exportsList() { return Object.keys(hb || {}).filter(k => /^hb_subset/.test(k)).join(" "); }
