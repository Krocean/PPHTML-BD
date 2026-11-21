/* ppt-html design kit - deck tools: the things a slide can do because it is a web page.
   Load after kit.js, then call  PH.deck()  once in the page script (after PH.nav).

     slide picker   press and HOLD any "next" / "previous" control (or click the page number in the dock):
                    every slide appears as a small picture; release on one, or click one, to jump there
     editing        the pencil in the dock. "글자": click any text and type. "옮기기": drag a text, a card, a chart
                    or a picture to another place (double-click puts it back). "가- / 가+": smaller / larger type.
                    What was changed is kept inside the presentation (PowerPoint asks to save when it closes) or,
                    in the web version, in the browser.
     pen            the pen in the dock: draw on the slide with the mouse; "지우기" clears
     zoom           data-zoom on an element: click to enlarge it, click again to put it back
     fragments      class ph-frag: hidden at first, revealed one by one by the "next" controls of the page
     hints          data-tip="설명": the explanation appears while the pointer is on the element
     timer          PH.timer(el, seconds): a countdown that starts and pauses on click
     auto advance   PH.deck({ autoNext: 12 }): go to the next slide after 12 seconds
     kept values    PH.keep("name", value) / PH.kept("name", fallback): small values of the page's own (votes,
                    ticked boxes) that stay in the presentation the same way the edits do

   ES5 only. Nothing here runs per frame; overlays appear at once (no fades over the whole page). */
(function (window, document) {
  "use strict";
  var PH = window.PH;
  if (!PH) { return; }
  var deck = null, opts = {}, dock = null, held = false, holdTimer = 0, picker = null, shade = null;
  var st = { v: 2, t: {}, m: {}, s: {}, d: {} }, loaded = false, orig = {};
  var editing = false, mode = "text", bar = null, barTip = null, modeBtns = [], active = null, lastSel = null, drag = null, editBtn = null;
  var pen = null, penOn = false, stroke = null, pts = "", lastX = 0, lastY = 0, penBtn = null, clearBtn = null;
  var zoomed = null, zoomShade = null, zoomSave = null, frags = [], hint = null;
  var near = false, nearTimer = 0;
  var docEl = document.documentElement;
  /* The tool bar stays out of sight (also in the still picture of the slide). It shows while the pointer is in
     the bottom right corner of the page, and while one of its tools is in use. */
  function syncDock() { if (dock) { PH.toggle(dock, "is-open", near || editing || penOn || !!picker); } }
  function setNear(on) {
    if (on) { clearTimeout(nearTimer); nearTimer = 0; if (!near) { near = true; syncDock(); } return; }
    if (near && !nearTimer) { nearTimer = setTimeout(function () { nearTimer = 0; near = false; syncDock(); }, 500); }
  }
  PH.on(document, "mousemove", function (e) { if (dock) { setNear(e.clientX > docEl.clientWidth - 360 && e.clientY > docEl.clientHeight - 76); } });
  PH.on(document, "mouseout", function (e) { if (dock && !e.relatedTarget) { setNear(false); } });

  function icon(d, parent) {
    var s = PH.svg("svg", { width: 16, height: 16, viewBox: "0 0 16 16" }, parent);
    PH.svg("path", { d: d, fill: "none", stroke: "currentColor", "stroke-width": 1.7, "stroke-linecap": "round", "stroke-linejoin": "round" }, s);
    return s;
  }
  function empty(o) { var k; for (k in o) { if (o.hasOwnProperty(k)) { return false; } } return true; }

  /* ---------------------------------------------------------------- what the page keeps */
  function load() {
    var text, o;
    if (loaded) { return; }
    loaded = true;
    try {
      text = window.pptHtml && window.pptHtml.loadState ? window.pptHtml.loadState() : "";
      o = text ? JSON.parse(text) : null;
      if (o && o.v === 2) { st.t = o.t || {}; st.m = o.m || {}; st.s = o.s || {}; st.d = o.d || {}; }
    } catch (e) { }
  }
  function save() {
    var none = empty(st.t) && empty(st.m) && empty(st.s) && empty(st.d);
    try { if (window.pptHtml && window.pptHtml.saveState) { window.pptHtml.saveState(none ? "" : JSON.stringify(st)); } } catch (e) { }
  }
  PH.keep = function (name, value) { load(); if (value === undefined || value === null) { delete st.d[name]; } else { st.d[name] = value; } save(); };
  PH.kept = function (name, fallback) { load(); return st.d.hasOwnProperty(name) ? st.d[name] : fallback; };

  /* ---------------------------------------------------------------- slide picker */
  function closePicker() {
    if (!picker) { return; }
    document.body.removeChild(picker.el); document.body.removeChild(shade);
    picker = null; shade = null;
    syncDock();
  }
  function openPicker(byHold) {
    var box, h;
    if (picker || !deck || !deck.slides || !deck.slides.length) { return; }
    shade = PH.make("div", "ph-shade", null, document.body);
    box = PH.make("div", "ph-picker ph-noedit" + (deck.slides.length > 15 ? " ph-small" : ""), null, document.body);
    h = PH.make("h3", "", "슬라이드 고르기", box);
    PH.make("small", "", byHold ? "가려는 장 위에서 손을 떼세요" : "가려는 장을 누르세요", h);
    deck.slides.forEach(function (s) {
      var p = PH.make("div", "ph-pick" + (s.n === deck.n ? " is-cur" : ""), null, box), img = PH.make("div", "ph-pick-img", null, p);
      if (s.thumb) { img.style.backgroundImage = "url(" + s.thumb + ")"; } else { PH.make("b", "", String(s.n), img); }
      PH.make("span", "", s.n + ". " + (s.title || ""), p);
      p.setAttribute("data-pick", s.n);
    });
    (function () {
      var small = deck.slides.length > 15, cell = small ? 160 : 210, cols = Math.min(deck.slides.length, small ? 7 : 5), w = Math.max(cols * cell + 38, 420);
      box.style.width = w + "px"; box.style.marginLeft = (-w / 2) + "px";
    }());
    picker = { el: box, byHold: !!byHold, hot: null };
    syncDock();
  }
  function pick(el) {
    var n = Number(el.getAttribute("data-pick"));
    closePicker();
    if (deck && n !== deck.n) { PH.goto(n); }
  }
  PH.picker = { open: function () { openPicker(false); }, close: closePicker };

  function zone(t, names) { var i; for (i = 0; i < names.length; i++) { if (PH.up(t, names[i])) { return true; } } return false; }
  PH.on(document, "mousedown", function (e) {
    var el = PH.upAttr(e.target, "data-next") || PH.upAttr(e.target, "data-prev") || PH.upAttr(e.target, "data-hold");
    clearTimeout(holdTimer);
    if (!el || !deck || editing || penOn) { return; }
    holdTimer = setTimeout(function () { held = true; openPicker(true); }, 450);
  }, true);
  /* while the button is still down after a long press, the picture under the pointer is marked */
  PH.on(document, "mousemove", function (e) {
    var hit, p;
    if (!picker || !picker.byHold) { return; }
    hit = document.elementFromPoint(e.clientX, e.clientY); p = hit ? PH.up(hit, "ph-pick") : null;
    if (p !== picker.hot) { if (picker.hot) { PH.remove(picker.hot, "is-hot"); } if (p) { PH.add(p, "is-hot"); } picker.hot = p; }
  }, true);
  PH.on(document, "mouseup", function (e) {
    var hit, p;
    clearTimeout(holdTimer);
    if (picker && picker.byHold) {
      picker.byHold = false;
      hit = document.elementFromPoint(e.clientX, e.clientY); p = hit ? PH.up(hit, "ph-pick") : null;
      if (p) { pick(p); }
    }
    if (held) { setTimeout(function () { held = false; }, 60); }
  }, true);
  /* the host talks to the page by "clicking" hidden elements of its own: those are not the audience's clicks */
  function hostSignal(t) { return !!t && (t.id === "ihCmd" || t.id === "ihShow" || t.id === "ihHide"); }
  PH.on(document, "click", function (e) {
    var t = e.target, el;
    if (hostSignal(t)) { return; }
    if (held) { held = false; e.stopPropagation(); e.preventDefault(); return; }       /* the click that ends a long press */
    if (picker) {
      e.stopPropagation(); e.preventDefault();
      if ((el = PH.up(t, "ph-pick"))) { pick(el); } else if (!PH.up(t, "ph-picker")) { closePicker(); }
      return;
    }
    if (zoomed) { e.stopPropagation(); e.preventDefault(); unzoom(); return; }
    if (editing) {
      if (zone(t, ["ph-editbar", "ph-dock"])) { return; }
      e.stopPropagation(); e.preventDefault();
      if (mode !== "text") { return; }
      el = PH.upAttr(t, "data-ph-t");
      if (el && el !== active) { begin(el); } else if (!el) { finish(); }
      return;
    }
    if (penOn) { if (!PH.up(t, "ph-dock")) { e.stopPropagation(); e.preventDefault(); } return; }
    if ((el = PH.upAttr(t, "data-zoom"))) { e.stopPropagation(); zoom(el); }
  }, true);

  /* ---------------------------------------------------------------- editing */
  /* An element is found again by its place in the page ("#list/div:2/b:0"), so a change survives as long as the
     page is built the same way. Each changed text remembers what it replaced: when the page's own text is later
     rewritten in the HTML, the old change is not laid over the new text. */
  function hash(s) { var h = 5381, i; for (i = 0; i < s.length; i++) { h = ((h << 5) + h + s.charCodeAt(i)) | 0; } return (h >>> 0).toString(36); }
  function keyOf(el) {
    var parts = [], n = 0, i, sib;
    while (el && el !== document.body && n < 30) {
      if (el.id) { parts.unshift("#" + el.id); break; }
      i = 0;
      for (sib = el.previousSibling; sib; sib = sib.previousSibling) { if (sib.nodeType === 1 && sib.tagName === el.tagName) { i++; } }
      parts.unshift(el.tagName.toLowerCase() + ":" + i);
      el = el.parentNode; n++;
    }
    return parts.join("/");
  }
  function find(key) {
    var parts = key.split("/"), el = document.body, i, p, tag, want, c, n;
    for (i = 0; i < parts.length && el; i++) {
      if (parts[i].charAt(0) === "#") { el = document.getElementById(parts[i].substring(1)); continue; }
      p = parts[i].split(":"); tag = p[0]; want = Number(p[1]); n = 0;
      for (c = el.firstChild; c; c = c.nextSibling) { if (c.nodeType === 1 && c.tagName.toLowerCase() === tag) { if (n === want) { break; } n++; } }
      el = c;
    }
    return el || null;
  }
  function isHtml(el) { return el.nodeType === 1 && el.namespaceURI !== "http://www.w3.org/2000/svg" && !/^(script|style)$/i.test(el.tagName); }
  function hasText(el) { var node; for (node = el.firstChild; node; node = node.nextSibling) { if (node.nodeType === 3 && /\S/.test(node.nodeValue)) { return true; } } return false; }
  function stamp() {
    var all = document.body.getElementsByTagName("*"), i, el;
    for (i = 0; i < all.length; i++) {
      el = all[i];
      if (!isHtml(el) || !hasText(el) || el.hasAttribute("data-count")) { continue; }
      if (PH.upAttr(el.parentNode, "data-ph-t") || zone(el, ["ph-nav", "ph-dock", "ph-chart", "ph-table", "ph-noedit"])) { continue; }
      el.setAttribute("data-ph-t", "1");
    }
  }
  function unstamp() { PH.$$("[data-ph-t]").forEach(function (el) { el.removeAttribute("data-ph-t"); }); }
  function place(el, mv) { el.style.transition = mv ? "none" : ""; el.style.transform = mv ? "translate(" + mv[0] + "px," + mv[1] + "px)" : ""; }
  /* lays the kept changes over the page: runs after every "show", because a page may draw its text again then */
  function applyEdits() {
    var k, el, cur;
    load();
    for (k in st.t) {
      if (st.t.hasOwnProperty(k) && (el = find(k)) && el !== active) {
        cur = el.innerHTML;
        if (cur !== st.t[k][1] && hash(cur) === st.t[k][0]) { orig[k] = cur; el.innerHTML = st.t[k][1]; }
      }
    }
    for (k in st.s) { if (st.s.hasOwnProperty(k) && (el = find(k))) { el.style.fontSize = st.s[k] + "px"; } }
    for (k in st.m) { if (st.m.hasOwnProperty(k) && (el = find(k)) && el !== zoomed) { place(el, st.m[k]); } }
  }
  PH.afterShow = applyEdits;
  function finish() {
    var key, html;
    if (!active) { return; }
    active.contentEditable = "false"; PH.remove(active, "is-edit");
    key = keyOf(active); html = active.innerHTML;
    if (html !== orig[key]) { st.t[key] = [hash(orig[key]), html]; } else { delete st.t[key]; }
    active = null;
    save();
    if (window.__ihSend) { window.__ihSend("focus"); }          /* the keyboard goes back to the slide show */
  }
  function begin(el) {
    var key;
    finish();
    key = keyOf(el);
    if (orig[key] === undefined) { orig[key] = el.innerHTML; }
    active = el; lastSel = el; el.contentEditable = "true"; PH.add(el, "is-edit");
    try { el.focus(); } catch (e) { }
  }
  function resize(up) {
    var el = active || lastSel, px, key;
    if (!el || !el.parentNode) { return; }
    px = parseFloat(window.getComputedStyle(el).fontSize) || 16;
    px = PH.clamp(Math.round(px * (up ? 1.1 : 1 / 1.1) * 10) / 10, 9, 320);
    key = keyOf(el);
    el.style.fontSize = px + "px"; st.s[key] = px; save();
  }
  function resetAll() {
    var k, el;
    finish();
    for (k in st.t) { if (st.t.hasOwnProperty(k) && orig[k] !== undefined && (el = find(k))) { el.innerHTML = orig[k]; } }
    for (k in st.s) { if (st.s.hasOwnProperty(k) && (el = find(k))) { el.style.fontSize = ""; } }
    for (k in st.m) { if (st.m.hasOwnProperty(k) && (el = find(k))) { place(el, null); } }
    st.t = {}; st.s = {}; st.m = {}; lastSel = null;
    save();
    if (editing) { unstamp(); stamp(); }
  }
  function setMode(m) {
    finish(); mode = m;
    modeBtns.forEach(function (b) { PH.toggle(b[1], "is-on", b[0] === m); });
    PH.toggle(document.body, "ph-moving", editing && m === "move");
    if (barTip) { barTip.textContent = m === "move" ? "끌어서 옮기기 · 두 번 누르면 제자리" : "글자를 누르고 입력"; }
  }
  function setEdit(on) {
    var b;
    if (on === editing) { return; }
    if (on && penOn) { setPen(false); }
    editing = on; PH.toggle(document.body, "ph-editing", on);
    if (editBtn) { PH.toggle(editBtn, "is-on", on); }
    syncDock();
    if (on) {
      stamp();
      bar = PH.make("div", "ph-editbar ph-noedit", null, document.body); modeBtns = [];
      b = PH.make("span", "", "글자", bar); modeBtns.push(["text", b]); PH.on(b, "click", function () { setMode("text"); });
      b = PH.make("span", "", "옮기기", bar); modeBtns.push(["move", b]); PH.on(b, "click", function () { setMode("move"); });
      PH.make("i", "", null, bar);
      b = PH.make("span", "", "가−", bar); b.setAttribute("title", "글자 작게"); PH.on(b, "click", function () { resize(false); });
      b = PH.make("span", "", "가+", bar); b.setAttribute("title", "글자 크게"); PH.on(b, "click", function () { resize(true); });
      PH.make("i", "", null, bar);
      b = PH.make("span", "", "원래대로", bar); PH.on(b, "click", resetAll);
      b = PH.make("span", "ph-eb-done", "완료", bar); PH.on(b, "click", function () { setEdit(false); });
      barTip = PH.make("em", "", "", bar);
      PH.on(bar, "mousedown", function (e) { e.preventDefault(); });       /* the text being typed keeps its caret */
      setMode("text");
    } else {
      finish(); drag = null;
      PH.remove(document.body, "ph-moving");
      if (bar) { document.body.removeChild(bar); bar = null; barTip = null; }
      unstamp();
    }
  }
  function movable(t) {
    var n = 0, el = t;
    while (el && el !== document.body && n < 14) {
      if (el.nodeType === 1 && el.hasAttribute && (el.hasAttribute("data-ph-t") || el.hasAttribute("data-move") || PH.has(el, "ph-card") || PH.has(el, "ph-band") || PH.has(el, "ph-chart") || /^img$/i.test(el.tagName))) { return el; }
      el = el.parentNode; n++;
    }
    return null;
  }
  PH.on(document, "mousedown", function (e) {
    var el, key, mv;
    if (!editing || zone(e.target, ["ph-editbar", "ph-dock"])) { return; }
    e.stopPropagation();                                          /* the page's own controls rest while editing */
    if (mode !== "move") { return; }
    e.preventDefault();
    el = movable(e.target);
    if (!el || zone(el, ["ph-nav"])) { return; }
    key = keyOf(el); mv = st.m[key] || [0, 0];
    drag = { el: el, key: key, x: e.clientX, y: e.clientY, dx: mv[0], dy: mv[1], now: [mv[0], mv[1]] };
    PH.add(el, "is-move");
  }, true);
  PH.on(document, "mousemove", function (e) {
    if (!drag) { return; }
    drag.now = [Math.round(drag.dx + e.clientX - drag.x), Math.round(drag.dy + e.clientY - drag.y)];
    place(drag.el, drag.now);
  }, true);
  PH.on(document, "mouseup", function () {
    if (!drag) { return; }
    PH.remove(drag.el, "is-move");
    if (drag.now[0] || drag.now[1]) { st.m[drag.key] = drag.now; } else { delete st.m[drag.key]; place(drag.el, null); }
    lastSel = drag.el; drag = null;
    save();
  }, true);
  PH.on(document, "dblclick", function (e) {
    var el, key;
    if (!editing || mode !== "move" || zone(e.target, ["ph-editbar", "ph-dock"])) { return; }
    el = movable(e.target); if (!el) { return; }
    key = keyOf(el); delete st.m[key]; place(el, null); save();
  }, true);
  PH.on(document, "keydown", function (e) {
    if (!active) { return; }
    if (e.keyCode === 13 && !e.shiftKey) { e.preventDefault(); e.stopPropagation(); finish(); try { document.body.focus(); } catch (x) { } }
  }, true);
  PH.on(document, "focusout", function (e) { if (active && e.target === active) { setTimeout(function () { if (active && document.activeElement !== active) { finish(); } }, 0); } });
  PH.edit = { on: function () { setEdit(true); }, off: function () { setEdit(false); }, active: function () { return editing; }, mode: setMode, reset: resetAll };

  /* ---------------------------------------------------------------- pen */
  function setPen(on) {
    if (on && editing) { setEdit(false); }
    penOn = on;
    if (on && !pen) {
      pen = PH.svg("svg", { "class": "ph-pen", width: docEl.clientWidth, height: docEl.clientHeight }, document.body);
      PH.svg("rect", { x: 0, y: 0, width: "100%", height: "100%", fill: "#000000", "fill-opacity": 0 }, pen);
      PH.on(pen, "mousedown", function (e) {
        lastX = e.clientX; lastY = e.clientY; pts = lastX + "," + lastY + " " + (lastX + 0.1) + "," + lastY;
        stroke = PH.svg("polyline", { fill: "none", stroke: PH.color("accent"), "stroke-width": 3.5, "stroke-linecap": "round", "stroke-linejoin": "round", points: pts }, pen);
        e.preventDefault();
      });
      PH.on(document, "mousemove", function (e) {
        if (!stroke) { return; }
        if (Math.abs(e.clientX - lastX) + Math.abs(e.clientY - lastY) < 3) { return; }
        lastX = e.clientX; lastY = e.clientY; pts += " " + lastX + "," + lastY;
        stroke.setAttribute("points", pts);
      });
      PH.on(document, "mouseup", function () { stroke = null; });
    }
    if (pen) { pen.style.display = on ? "block" : "none"; pen.setAttribute("width", docEl.clientWidth); pen.setAttribute("height", docEl.clientHeight); }
    if (penBtn) { PH.toggle(penBtn, "is-on", on); }
    if (clearBtn) { clearBtn.style.display = on ? "inline-block" : "none"; }
    syncDock();
  }
  function clearPen() { var n; if (!pen) { return; } while ((n = pen.lastChild) && n.tagName !== "rect") { pen.removeChild(n); } }
  PH.pen = { on: function () { setPen(true); }, off: function () { setPen(false); }, clear: clearPen, strokes: function () { return pen ? pen.getElementsByTagName("polyline").length : 0; } };

  /* ---------------------------------------------------------------- zoom */
  function zoom(el) {
    var r = el.getBoundingClientRect(), W = docEl.clientWidth, H = docEl.clientHeight, k = Math.min(W * 0.9 / r.width, H * 0.86 / r.height, 2.6), mv = st.m[keyOf(el)] || [0, 0];
    if (k < 1.08) { return; }
    zoomShade = PH.make("div", "ph-shade", null, document.body);
    zoomSave = { t: el.style.transform, z: el.style.zIndex, p: el.style.position };
    if (window.getComputedStyle(el).position === "static") { el.style.position = "relative"; }
    el.style.zIndex = 79;
    el.style.transform = "translate(" + Math.round(W / 2 - (r.left + r.width / 2) + mv[0]) + "px," + Math.round(H / 2 - (r.top + r.height / 2) + mv[1]) + "px) scale(" + (Math.round(k * 100) / 100) + ")";
    zoomed = el;
  }
  function unzoom() {
    if (!zoomed) { return; }
    zoomed.style.transform = zoomSave.t; zoomed.style.zIndex = zoomSave.z; zoomed.style.position = zoomSave.p;
    document.body.removeChild(zoomShade);
    zoomed = null; zoomShade = null;
  }
  PH.zoom = { on: zoom, off: unzoom, active: function () { return !!zoomed; } };

  /* ---------------------------------------------------------------- fragments, hints, timer */
  function resetFrags() {
    frags = PH.$$(".ph-frag");
    frags.forEach(function (f) { PH.toggle(f, "is-on", PH.preloading()); });      /* the still picture shows everything */
  }
  PH.beforeNext = function () {
    var i;
    for (i = 0; i < frags.length; i++) { if (!PH.has(frags[i], "is-on")) { PH.add(frags[i], "is-on"); return true; } }
    return false;
  };
  PH.on(document, "mouseover", function (e) {
    var el = PH.upAttr(e.target, "data-tip"), r;
    if (!el) { return; }
    if (!hint) { hint = PH.make("div", "ph-hint ph-noedit", null, document.body); }
    hint.textContent = el.getAttribute("data-tip"); PH.add(hint, "is-on");
    r = el.getBoundingClientRect();
    hint.style.left = PH.clamp(r.left, 8, docEl.clientWidth - hint.offsetWidth - 8) + "px";
    hint.style.top = (r.bottom + 8 + hint.offsetHeight > docEl.clientHeight ? r.top - hint.offsetHeight - 8 : r.bottom + 8) + "px";
  });
  PH.on(document, "mouseout", function (e) { if (hint && PH.upAttr(e.target, "data-tip")) { PH.remove(hint, "is-on"); } });
  /* PH.timer(el, 300): shows 5:00; click starts, click again pauses; turns red at 0:00 */
  PH.timer = function (el, seconds) {
    var left = seconds, id = 0;
    function paint() { var m = Math.floor(left / 60), s = left % 60; el.textContent = m + ":" + (s < 10 ? "0" : "") + s; PH.toggle(el, "is-over", left <= 0); }
    function stop() { if (id) { PH.stop(id); id = 0; } }
    PH.add(el, "ph-timer"); PH.add(el, "ph-noedit"); paint();
    PH.on(el, "click", function () {
      if (id) { stop(); return; }
      if (left <= 0) { left = seconds; paint(); return; }
      id = PH.every(1000, function () { left--; paint(); if (left <= 0) { stop(); } });
    });
    PH.onShow(function () { stop(); left = seconds; paint(); });
    return { reset: function () { stop(); left = seconds; paint(); } };
  };

  /* ---------------------------------------------------------------- dock */
  function buildDock() {
    var prev, next, num;
    dock = PH.make("div", "ph-dock ph-noedit", null, document.body);
    if (deck) {
      prev = PH.make("span", "", null, dock); icon("M10 3L5 8l5 5", prev); prev.setAttribute("data-prev", ""); prev.setAttribute("title", "이전 · 길게 누르면 슬라이드 고르기");
      num = PH.make("span", "ph-dock-n", deck.n + " / " + deck.total, dock); num.setAttribute("data-hold", ""); num.setAttribute("title", "슬라이드 고르기");
      PH.on(num, "click", function () { openPicker(false); });
      next = PH.make("span", "", null, dock); icon("M6 3l5 5-5 5", next); next.setAttribute("data-next", ""); next.setAttribute("title", "다음 · 길게 누르면 슬라이드 고르기");
    }
    if (opts.edit !== false || opts.pen !== false) { if (deck) { PH.make("i", "", null, dock); } }
    if (opts.edit !== false) {
      editBtn = PH.make("span", "", null, dock); icon("M3 13l1-3.2L10.8 3l2.2 2.2L6.2 12z", editBtn); editBtn.setAttribute("title", "고치기: 글자, 자리, 크기");
      PH.on(editBtn, "click", function () { setEdit(!editing); });
    }
    if (opts.pen !== false) {
      penBtn = PH.make("span", "", null, dock); icon("M2 11c2-6 4 3 6-2s3-3 6-5", penBtn); penBtn.setAttribute("title", "펜으로 그리기");
      PH.on(penBtn, "click", function () { setPen(!penOn); });
      clearBtn = PH.make("span", "", "지우기", dock); clearBtn.style.display = "none";
      PH.on(clearBtn, "click", clearPen);
    }
  }

  /* PH.deck({ dock: true, progress: true, edit: true, pen: true, autoNext: 0 }) */
  PH.deck = function (o) {
    var p;
    opts = o || {};
    deck = window.__ihDeck || (window.pptHtml && window.pptHtml.deck) || null;
    load();
    if (opts.dock !== false) { buildDock(); }
    if (opts.progress !== false && deck && deck.total > 1) { p = PH.make("div", "ph-progress", null, document.body); p.style.width = (deck.n / deck.total * 100) + "%"; }
    PH.onShow(function () {
      resetFrags();
      if (opts.autoNext > 0 && !PH.preloading()) { PH.wait(opts.autoNext * 1000, function () { PH.next(); }); }
    });
    PH.onHide(function () { closePicker(); unzoom(); if (penOn) { setPen(false); } if (editing) { setEdit(false); } });
    return deck;
  };
}(window, document));
