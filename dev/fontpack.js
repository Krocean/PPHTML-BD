// Dev-time: builds the font pack that ships with the skill (assets\fonts), for PCs that do not have these fonts.
// Every font here is openly licensed (SIL OFL / Apache). Each file is the original cut down to the characters
// of everyday Korean and English text: KS X 1001 Hangul (2,350 syllables), Latin, punctuation, the usual symbols.
// Variable fonts keep their weight axis, so the build can still pin any weight.
//   node fontpack.js        (after node build.js; needs the packages of package.json and, for the two Noto
//                            fonts, Windows 11 with the Noto Korean fonts in C:\Windows\Fonts)
const fs = require("fs");
const path = require("path");
const pako = require("pako");
const wawoff2 = require("wawoff2");
const f = require("./out/fonttool.cjs");
const skill = path.join(__dirname, "..");
const outDir = path.join(skill, "assets", "fonts");
fs.mkdirSync(outDir, { recursive: true });
const nodePaths = [path.join(__dirname, "node_modules")].concat((process.env.NODE_PATH || "").split(path.delimiter).filter(Boolean));
const mod = rel => { for (const p of nodePaths) { if (fs.existsSync(path.join(p, rel))) { return path.join(p, rel); } } throw new Error("not found in node_modules: " + rel); };

function charset() {
  const set = new Set();
  const add = (a, b) => { for (let c = a; c <= b; c++) { set.add(c); } };
  add(0x20, 0x7E); add(0xA0, 0xFF);                       // Latin
  add(0x2010, 0x2027); add(0x2030, 0x205E);               // punctuation
  add(0x20A0, 0x20BF);                                    // currency
  add(0x2100, 0x214F); add(0x2150, 0x218B);               // letterlike, number forms
  add(0x2190, 0x21FF); add(0x2200, 0x22FF);               // arrows, math
  add(0x2300, 0x23FF);                                    // technical
  add(0x2460, 0x24FF);                                    // enclosed
  add(0x2500, 0x257F); add(0x2580, 0x259F); add(0x25A0, 0x25FF); add(0x2600, 0x26FF); add(0x2700, 0x27BF);   // boxes, shapes, symbols
  add(0x3000, 0x303F);                                    // CJK punctuation
  add(0x3131, 0x318E);                                    // Hangul compatibility jamo
  add(0x3200, 0x321E); add(0x3260, 0x327F);               // enclosed Hangul
  add(0xFF01, 0xFF5E); add(0xFFE0, 0xFFE6);               // full width
  const dec = new TextDecoder("euc-kr");
  for (let lead = 0xB0; lead <= 0xC8; lead++) {
    for (let trail = 0xA1; trail <= 0xFE; trail++) {
      const ch = dec.decode(new Uint8Array([lead, trail]));
      if (ch.length === 1 && ch.charCodeAt(0) >= 0xAC00 && ch.charCodeAt(0) <= 0xD7A3) { set.add(ch.charCodeAt(0)); }
    }
  }
  return Array.from(set).sort((a, b) => a - b);
}
function nameOf(sfnt, id) {             // name table entry (Windows, English)
  const dv = new DataView(sfnt.buffer, sfnt.byteOffset, sfnt.byteLength), n = dv.getUint16(4);
  for (let i = 0; i < n; i++) {
    const p = 12 + i * 16;
    if (dv.getUint32(p) !== 0x6E616D65) { continue; }
    const off = dv.getUint32(p + 8), count = dv.getUint16(off + 2), strOff = off + dv.getUint16(off + 4);
    for (let k = 0; k < count; k++) {
      const r = off + 6 + k * 12;
      if (dv.getUint16(r) === 3 && dv.getUint16(r + 4) === 0x409 && dv.getUint16(r + 6) === id) {
        const len = dv.getUint16(r + 8), so = strOff + dv.getUint16(r + 10); let s = "";
        for (let j = 0; j < len; j += 2) { s += String.fromCharCode(dv.getUint16(so + j)); }
        return s;
      }
    }
  }
  return "";
}
function fromWoff(woff) {               // WOFF 1.0 -> sfnt
  const dv = new DataView(woff.buffer, woff.byteOffset, woff.byteLength);
  if (dv.getUint32(0) !== 0x774F4646) { throw new Error("not a WOFF file"); }
  const n = dv.getUint16(12), tables = [];
  for (let i = 0; i < n; i++) {
    const p = 44 + i * 20, off = dv.getUint32(p + 4), comp = dv.getUint32(p + 8), orig = dv.getUint32(p + 12);
    let data = woff.subarray(off, off + comp);
    if (comp !== orig) { data = pako.inflate(data); }
    tables.push({ tag: dv.getUint32(p), sum: dv.getUint32(p + 16), data });
  }
  let size = 12 + 16 * n; for (const t of tables) { size += (t.data.length + 3) & ~3; }
  const out = new Uint8Array(size), o = new DataView(out.buffer);
  let e = 0; while ((1 << (e + 1)) <= n) { e++; }
  o.setUint32(0, dv.getUint32(4)); o.setUint16(4, n); o.setUint16(6, (1 << e) * 16); o.setUint16(8, e); o.setUint16(10, n * 16 - (1 << e) * 16);
  let pos = 12 + 16 * n;
  tables.forEach((t, i) => { const p = 12 + i * 16; o.setUint32(p, t.tag); o.setUint32(p + 4, t.sum); o.setUint32(p + 8, pos); o.setUint32(p + 12, t.data.length); out.set(t.data, pos); pos += (t.data.length + 3) & ~3; });
  return out;
}

(async () => {
  await f.init();
  const cps = charset();
  const text = String.fromCodePoint.apply(null, cps);
  console.log("characters:", cps.length, "(Hangul " + cps.filter(c => c >= 0xAC00).length + ")");
  const win = "C:/Windows/Fonts";
  const jobs = [
    { out: "NotoSansKR-VF.subset.ttf", read: () => new Uint8Array(fs.readFileSync(path.join(win, "NotoSansKR-VF.ttf"))), axes: true },
    { out: "NotoSerifKR-VF.subset.ttf", read: () => new Uint8Array(fs.readFileSync(path.join(win, "NotoSerifKR-VF.ttf"))), axes: true },
    { out: "Pretendard-VF.subset.ttf", read: () => new Uint8Array(fs.readFileSync(mod("pretendard/dist/public/variable/PretendardVariable.ttf"))), axes: true },
    { out: "Inter-VF.subset.ttf", read: async () => new Uint8Array(await wawoff2.decompress(fs.readFileSync(mod("@fontsource-variable/inter/files/inter-latin-wght-normal.woff2")))), axes: true },
    { out: "Roboto-VF.subset.ttf", read: async () => new Uint8Array(await wawoff2.decompress(fs.readFileSync(mod("@fontsource-variable/roboto/files/roboto-latin-wght-normal.woff2")))), axes: true },
    { out: "OpenSans-VF.subset.ttf", read: async () => new Uint8Array(await wawoff2.decompress(fs.readFileSync(mod("@fontsource-variable/open-sans/files/open-sans-latin-wght-normal.woff2")))), axes: true },
    { out: "Geist-VF.subset.ttf", read: async () => new Uint8Array(await wawoff2.decompress(fs.readFileSync(mod("@fontsource-variable/geist/files/geist-latin-wght-normal.woff2")))), axes: true },
    { out: "IBMPlexSansKR-Light.subset.ttf", read: () => fromWoff(new Uint8Array(fs.readFileSync(mod("@ibm/plex-sans-kr/fonts/complete/woff/hinted/IBMPlexSansKR-Light.woff")))) },
    { out: "IBMPlexSansKR-Regular.subset.ttf", read: () => fromWoff(new Uint8Array(fs.readFileSync(mod("@ibm/plex-sans-kr/fonts/complete/woff/hinted/IBMPlexSansKR-Regular.woff")))) },
    { out: "IBMPlexSansKR-SemiBold.subset.ttf", read: () => fromWoff(new Uint8Array(fs.readFileSync(mod("@ibm/plex-sans-kr/fonts/complete/woff/hinted/IBMPlexSansKR-SemiBold.woff")))) }
  ];
  const licence = []; let total = 0;
  for (const j of jobs) {
    let src;
    try { src = await j.read(); } catch (e) { console.log("SKIPPED " + j.out + ": " + e.message); continue; }
    f.addBytes(j.out, src);
    const sub = f.subsetBytes(j.out, text, { keepAxes: !!j.axes });
    fs.writeFileSync(path.join(outDir, j.out), Buffer.from(sub));
    total += sub.length;
    console.log(j.out.padEnd(36), (sub.length / 1048576).toFixed(2) + " MB", "| family:", nameOf(sub, 16) || nameOf(sub, 1) || nameOf(src, 1));
    licence.push(j.out + "  (" + (nameOf(src, 16) || nameOf(src, 1)) + ")\n  " + [nameOf(src, 0), nameOf(src, 13), nameOf(src, 14)].filter(Boolean).join("\n  "));
    f.drop(j.out);
  }
  fs.writeFileSync(path.join(outDir, "LICENSE-fonts.txt"),
    "The font files in this folder are subsets (fewer characters, otherwise unchanged) of openly licensed fonts.\n" +
    "Copyright and licence statements carried by the original fonts:\n\n" + licence.join("\n\n") + "\n\n" +
    "They are licensed under the SIL Open Font License, Version 1.1 (https://openfontlicense.org) unless a statement above says\n" +
    "otherwise. The licence permits use, modification, embedding and redistribution, provided the fonts are not sold by\n" +
    "themselves and derivative fonts keep the licence.\n");
  console.log("total " + (total / 1048576).toFixed(2) + " MB");
})().catch(e => { console.error(e); process.exit(1); });
