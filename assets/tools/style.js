/* ppt-html: reads the look of the page it is evaluated in (get-style.ps1 runs it in the tool browser).
   Returns a JSON text: colours by how much they are used, fonts, sizes, corner shapes, shadows, buttons, cards,
   headings, CSS variables of the root. Nothing is sent anywhere; the page is only measured. */
(function () {
  var cv = document.createElement("canvas"); cv.width = cv.height = 1;
  var cx = cv.getContext("2d"), cache = {};
  function col(s) {
    var m, a, d, r;
    if (!s) { return null; }
    if (cache[s] !== undefined) { return cache[s]; }
    m = /^rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:[,\s\/]+([\d.]+%?))?\s*\)$/.exec(s);
    if (m) {
      a = m[4] === undefined ? 1 : (/%$/.test(m[4]) ? parseFloat(m[4]) / 100 : parseFloat(m[4]));
      r = [Math.round(+m[1]), Math.round(+m[2]), Math.round(+m[3]), a];
    } else {
      try {
        cx.clearRect(0, 0, 1, 1); cx.fillStyle = "#000000"; cx.fillStyle = s; cx.fillRect(0, 0, 1, 1);
        d = cx.getImageData(0, 0, 1, 1).data;
        r = d[3] ? [Math.round(d[0] * 255 / d[3]), Math.round(d[1] * 255 / d[3]), Math.round(d[2] * 255 / d[3]), d[3] / 255] : [0, 0, 0, 0];
      } catch (e) { r = null; }
    }
    cache[s] = r;
    return r;
  }
  function hex(c) { return "#" + [c[0], c[1], c[2]].map(function (v) { v = Math.max(0, Math.min(255, v)); return (v < 16 ? "0" : "") + v.toString(16); }).join("").toUpperCase(); }
  function add(map, key, w, extra) { var o; if (!key) { return; } o = map[key] || (map[key] = { key: key, w: 0 }); o.w += w; if (extra && !o.x) { o.x = extra; } }
  function top(map, n) {
    return Object.keys(map).map(function (k) { return map[k]; }).sort(function (a, b) { return b.w - a.w; }).slice(0, n)
      .map(function (o) { var r = { v: o.key, w: Math.round(o.w * 10) / 10 }; if (o.x) { r.x = o.x; } return r; });
  }
  function first(f) { return String(f || "").split(",")[0].replace(/["']/g, "").trim(); }
  var VW = window.innerWidth, VH = window.innerHeight, screenArea = VW * VH;
  var els = Array.prototype.slice.call(document.querySelectorAll("body, body *")).slice(0, 6000);
  var bg = {}, fg = {}, fam = {}, size = {}, rad = {}, shadow = {}, border = {}, vivid = {}, texts = [], buttons = [], cards = [], grads = 0, chars = 0, upper = 0;
  els.forEach(function (el) {
    var tag = el.tagName, cs, r, area, b, hasBg, own = 0, n, c, f, fs, bw, bc, hasBorder, sh, rr, cls, isBtn;
    if (/^(SCRIPT|STYLE|NOSCRIPT|TEMPLATE|BR|WBR|META|LINK)$/.test(tag)) { return; }
    cs = window.getComputedStyle(el);
    if (!cs || cs.display === "none" || cs.visibility === "hidden" || Number(cs.opacity) === 0) { return; }
    r = el.getBoundingClientRect();
    if (r.width < 2 || r.height < 2) { return; }
    area = r.width * r.height;
    b = col(cs.backgroundColor); hasBg = !!b && b[3] > 0.5;
    if (hasBg) { add(bg, hex(b), Math.min(area, screenArea * 3)); }
    if (/gradient/.test(cs.backgroundImage) && area > screenArea * 0.1) { grads++; }
    for (n = el.firstChild; n; n = n.nextSibling) { if (n.nodeType === 3) { own += n.nodeValue.replace(/\s+/g, " ").trim().length; } }
    c = col(cs.color);
    if (own) {
      fs = parseFloat(cs.fontSize) || 0; f = first(cs.fontFamily);
      if (c) { add(fg, hex(c), own); }
      add(fam, f, own, cs.fontFamily);
      add(size, Math.round(fs) + "px/" + cs.fontWeight, own);
      chars += own;
      if (cs.textTransform === "uppercase" && fs <= 15) { upper += own; }
      texts.push({ t: tag.toLowerCase(), px: Math.round(fs * 10) / 10, w: cs.fontWeight, f: f, ls: cs.letterSpacing, lh: cs.lineHeight, c: c ? hex(c) : "", n: own });
    }
    bw = parseFloat(cs.borderTopWidth) || 0; bc = col(cs.borderTopColor);
    hasBorder = bw > 0 && cs.borderTopStyle !== "none" && !!bc && bc[3] > 0.2;
    sh = !!cs.boxShadow && cs.boxShadow !== "none";
    rr = parseFloat(cs.borderTopLeftRadius) || 0;
    if ((hasBg || hasBorder || sh) && r.width > 24 && r.height > 16 && tag !== "BODY" && tag !== "HTML") {
      add(rad, (rr > 12 && rr >= Math.min(r.width, r.height) / 2 - 1) ? "pill" : Math.round(rr) + "px", 1);
      if (sh) { add(shadow, cs.boxShadow, 1); }
      if (hasBorder) { add(border, hex(bc) + " " + bw + "px", 1); }
    }
    cls = typeof el.className === "string" ? el.className : "";
    isBtn = tag === "BUTTON" || (tag === "INPUT" && /submit|button/.test(el.type || "")) || el.getAttribute("role") === "button" || /(^|[\s_-])(btn|button|cta)([\s_-]|$)/i.test(cls);
    if (isBtn && r.width > 30 && r.height > 18 && buttons.length < 40) {
      buttons.push({ bg: hasBg ? hex(b) : "", c: c ? hex(c) : "", r: Math.round(rr), h: Math.round(r.height), w: cs.fontWeight, px: Math.round(parseFloat(cs.fontSize) || 0), bd: hasBorder ? hex(bc) : "", pill: rr > 8 && rr >= r.height / 2 - 1 });
    }
    if ((hasBg || hasBorder || sh) && area > screenArea * 0.015 && area < screenArea * 0.45 && el.children.length > 0 && cards.length < 80 && !isBtn) {
      cards.push({ bg: hasBg ? hex(b) : "", r: Math.round(rr), bd: hasBorder ? hex(bc) + " " + bw + "px" : "", sh: sh ? cs.boxShadow : "", p: Math.round(parseFloat(cs.paddingTop) || 0) + "/" + Math.round(parseFloat(cs.paddingLeft) || 0), w: Math.round(r.width), h: Math.round(r.height) });
    }
    /* colours with a hue: how saturated they are counts as much as how often they are used */
    [[hasBg ? b : null, Math.min(area, 60000) / 600 * (isBtn ? 3 : 1)], [own ? c : null, own], [tag === "A" ? c : null, 20]].forEach(function (p) {
      var k = p[0], mx, mn, sat;
      if (!k) { return; }
      mx = Math.max(k[0], k[1], k[2]); mn = Math.min(k[0], k[1], k[2]); sat = mx ? (mx - mn) / mx : 0;
      if (sat < 0.25 || mx < 48) { return; }
      add(vivid, hex(k), p[1] * sat);
    });
  });
  function pageBg() {
    var list = [document.body, document.documentElement], i, c, t;
    for (i = 0; i < list.length; i++) { c = col(window.getComputedStyle(list[i]).backgroundColor); if (c && c[3] > 0.5) { return hex(c); } }
    t = top(bg, 1);
    return t.length && t[0].w > screenArea * 0.5 ? t[0].v : "#FFFFFF";
  }
  /* headings: the declared ones, and the largest texts whatever their tag */
  function pick(tag) { var best = null; texts.forEach(function (t) { if (t.t === tag && (!best || t.n > best.n)) { best = t; } }); return best; }
  var bySize = texts.slice().sort(function (a, b) { return b.px - a.px || b.n - a.n; }), largest = [], seen = {};
  bySize.forEach(function (t) { var k = t.px + "/" + t.w + "/" + t.f; if (!seen[k] && largest.length < 6 && t.n >= 2) { seen[k] = 1; largest.push(t); } });
  var bodyMap = {}, bodyBest = null;
  texts.forEach(function (t) { var k; if (t.px < 12 || t.px > 22) { return; } k = t.px + "|" + t.w + "|" + t.f + "|" + t.lh + "|" + t.c; bodyMap[k] = (bodyMap[k] || 0) + t.n; });
  Object.keys(bodyMap).forEach(function (k) { if (!bodyBest || bodyMap[k] > bodyMap[bodyBest]) { bodyBest = k; } });
  var body = null, bp;
  if (bodyBest) { bp = bodyBest.split("|"); body = { px: Number(bp[0]), w: bp[1], f: bp[2], lh: bp[3], c: bp[4] }; }
  /* design tokens the page declares itself */
  var vars = {}, count = 0, i, j, k, rules, ru, pn;
  for (i = 0; i < document.styleSheets.length; i++) {
    try { rules = document.styleSheets[i].cssRules; } catch (e) { continue; }
    if (!rules) { continue; }
    for (j = 0; j < rules.length && j < 4000; j++) {
      ru = rules[j];
      if (ru.type !== 1 || !/(^|,)\s*(:root|html|body)\s*(,|$)/.test(ru.selectorText || "")) { continue; }
      for (k = 0; k < ru.style.length; k++) {
        pn = ru.style[k];
        if (pn.indexOf("--") === 0 && count < 120 && vars[pn] === undefined) { vars[pn] = ru.style.getPropertyValue(pn).trim().slice(0, 140); count++; }
      }
    }
  }
  var loaded = [];
  try { document.fonts.forEach(function (ff) { var nm = String(ff.family).replace(/["']/g, ""); if (ff.status === "loaded" && loaded.indexOf(nm) < 0 && loaded.length < 20) { loaded.push(nm); } }); } catch (e) { }
  return JSON.stringify({
    url: String(location.href).slice(0, 300), title: String(document.title).slice(0, 120), viewport: [VW, VH], docHeight: document.documentElement.scrollHeight,
    pageBg: pageBg(), bg: top(bg, 10), text: top(fg, 10), vivid: top(vivid, 10), fonts: top(fam, 8), sizes: top(size, 14),
    radius: top(rad, 8), shadows: top(shadow, 5), borders: top(border, 6),
    h1: pick("h1"), h2: pick("h2"), h3: pick("h3"), largest: largest, body: body,
    buttons: buttons.slice(0, 14), cards: cards.slice(0, 20), gradients: grads, upperShare: chars ? Math.round(upper / chars * 1000) / 1000 : 0,
    vars: vars, loadedFonts: loaded, chars: chars
  });
}())
