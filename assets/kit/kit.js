/* ppt-html design kit - behaviour. ES5 only: it runs as written in the slide show engine (IE11) and in every
   current browser. Everything is drawn with HTML and SVG, so it stays sharp when PowerPoint enlarges the page.
   Load it at the END of <body> (after the markup), then call PH.* from your own script below it.

   Markup that works without any script of your own:
     data-goto="3" / data-next / data-prev      jump between slides
     data-tab="panelId" (inside .ph-seg / .ph-chips) + <div class="ph-panel" id="panelId">   tabs
     .ph-acc > .ph-acc-q + .ph-acc-a            accordion (container with data-single: one open at a time)
     .ph-switch                                  on/off switch (fires "ph-change")
     data-count="1284" [data-digits="1"]        number that counts up when the slide appears
     .ph-meter > i[data-w="72"]                  bar that fills to 72 %
     .ph-rise [.ph-d1 ... .ph-d6]               entrance, replayed every time the slide appears

   Smooth motion: the engine repaints every pixel that moves, so the kit only moves small things. Anything
   large marked .ph-rise simply appears; if the PC still cannot keep up, all motion is switched off (ph-lite). */
(function (window, document) {
  "use strict";
  var PH = window.PH = {};
  var SVGNS = "http://www.w3.org/2000/svg";
  var showFns = [], hideFns = [], timers = [], anims = [], colors = {}, shown = false, quiet = false, lite = false, measured = false;

  /* ---------------------------------------------------------------- helpers */
  PH.$ = function (sel, root) { return (root || document).querySelector(sel); };
  PH.$$ = function (sel, root) { var l = (root || document).querySelectorAll(sel), out = [], i; for (i = 0; i < l.length; i++) { out.push(l[i]); } return out; };
  PH.has = function (el, c) { return (" " + (el.getAttribute("class") || "") + " ").indexOf(" " + c + " ") >= 0; };
  PH.add = function (el, c) { if (!PH.has(el, c)) { var k = el.getAttribute("class") || ""; el.setAttribute("class", k ? k + " " + c : c); } };
  PH.remove = function (el, c) { var k = (" " + (el.getAttribute("class") || "") + " ").replace(" " + c + " ", " ").replace(/^\s+|\s+$/g, ""); el.setAttribute("class", k); };
  PH.toggle = function (el, c, on) { if (on === undefined) { on = !PH.has(el, c); } if (on) { PH.add(el, c); } else { PH.remove(el, c); } return on; };
  PH.make = function (tag, cls, text, parent) { var el = document.createElement(tag); if (cls) { el.className = cls; } if (text !== undefined && text !== null) { el.textContent = text; } if (parent) { parent.appendChild(el); } return el; };
  PH.svg = function (tag, attrs, parent) { var el = document.createElementNS(SVGNS, tag), k; for (k in attrs) { if (attrs.hasOwnProperty(k)) { el.setAttribute(k, attrs[k]); } } if (parent) { parent.appendChild(el); } return el; };
  PH.on = function (el, type, fn, capture) { el.addEventListener(type, fn, !!capture); };
  PH.fire = function (el, type, detail) { var e = document.createEvent("CustomEvent"); e.initCustomEvent(type, true, false, detail); el.dispatchEvent(e); };
  PH.clamp = function (v, a, b) { return v < a ? a : (v > b ? b : v); };
  PH.lerp = function (a, b, t) { return a + (b - a) * t; };
  PH.ease = function (t) { return 1 - Math.pow(1 - t, 3); };
  PH.fmt = function (n, digits) { var s = Number(n).toFixed(digits || 0), p = s.split("."); p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ","); return p.join("."); };
  function r1(v) { return Math.round(v * 10) / 10; }
  PH.up = function (el, cls) { var n = 0; while (el && el !== document && n < 14) { if (el.getAttribute && PH.has(el, cls)) { return el; } el = el.parentNode; n++; } return null; };
  PH.upAttr = function (el, attr) { var n = 0; while (el && el !== document && n < 14) { if (el.hasAttribute && el.hasAttribute(attr)) { return el; } el = el.parentNode; n++; } return null; };

  /* theme colours for drawings: PH.color("c1") ... "c6", "line", "ink", "ink3", "surface", "bg", "accent", or any CSS colour */
  PH.color = function (name) {
    if (!name) { name = "c1"; }
    if (/^(#|rgb|hsl)/.test(name)) { return name; }
    if (!colors[name]) {
      var probe = PH.make("i", "ph-" + (/^c\d$/.test(name) ? name : "c" + name), null, document.body);
      probe.style.display = "none";
      colors[name] = window.getComputedStyle(probe).color || "#888888";
      document.body.removeChild(probe);
    }
    return colors[name];
  };
  PH.palette = function (n) { var out = [], i; for (i = 0; i < n; i++) { out.push(PH.color("c" + (i % 6 + 1))); } return out; };

  /* ---------------------------------------------------------------- time: everything stops when the slide is left */
  PH.wait = function (ms, fn) { if (quiet) { fn(); return 0; } var id = setTimeout(fn, ms); timers.push([0, id]); return id; };
  PH.every = function (ms, fn) { if (quiet) { return 0; } var id = setInterval(fn, ms); timers.push([1, id]); return id; };
  PH.stop = function (id) { clearTimeout(id); clearInterval(id); };
  function clearTimers() { var i; for (i = 0; i < timers.length; i++) { if (timers[i][0]) { clearInterval(timers[i][1]); } else { clearTimeout(timers[i][1]); } } timers = []; }
  /* PH.animate(600, function (t) { ... t goes 0..1, already eased ... }, done)   (jumps to the end when motion is off) */
  PH.animate = function (ms, step, done, delay) {
    var a = { live: true }, start = 0;
    if (quiet || lite) { step(1); if (done) { done(); } return a; }
    function frame(now) {
      if (!a.live) { return; }
      if (!start) { start = now + (delay || 0); }
      var t = (now - start) / ms;
      if (t < 0) { window.requestAnimationFrame(frame); return; }
      if (t >= 1) { a.live = false; step(1); if (done) { done(); } return; }
      step(PH.ease(t));
      window.requestAnimationFrame(frame);
    }
    anims.push(a);
    step(0);
    window.requestAnimationFrame(frame);
    return a;
  };
  function stopAnims() { var i; for (i = 0; i < anims.length; i++) { anims[i].live = false; } anims = []; }
  PH.still = function () { return quiet || lite; };
  PH.preloading = function () { return quiet; };              /* true while the host loads the page ahead of its slide */          /* true while motion is off: draw the final state at once */

  /* sizes: PowerPoint sets the zoom of the page after it has loaded, so anything measured in script is measured
     again when the viewport changes. PH.onResize(fn) is for drawings of your own. */
  var sized = [], lastVW = 0;
  PH.onResize = function (fn) { sized.push(fn); };
  function checkSize() {
    var w = document.documentElement.clientWidth, i;
    if (w === lastVW) { return; }
    lastVW = w;
    for (i = 0; i < sized.length; i++) { try { sized[i](); } catch (e) { } }
  }
  PH.on(window, "resize", function () { setTimeout(checkSize, 40); });

  /* ---------------------------------------------------------------- slide life cycle */
  PH.onShow = function (fn) { showFns.push(fn); };
  PH.onHide = function (fn) { hideFns.push(fn); };
  PH.goto = function (n) { if (window.pptHtml) { window.pptHtml.gotoSlide(Number(n)); } };
  PH.beforeNext = null;                                       /* a function that returns true when it used the "next" itself (fragments) */
  PH.next = function () { if (PH.beforeNext && PH.beforeNext()) { return; } if (window.pptHtml) { window.pptHtml.next(); } };
  PH.prev = function () { if (window.pptHtml) { window.pptHtml.prev(); } };
  PH.hosted = function () { return !!(window.pptHtml && window.pptHtml.host); };
  function setLite(on) { lite = on; PH.toggle(document.documentElement, "ph-lite", on || quiet); }
  /* large things do not animate: moving or fading a big card costs more frames than anything else on the page */
  function markStill() {
    PH.$$(".ph-rise").forEach(function (el) {
      var big = PH.has(el, "ph-card") || el.offsetWidth * el.offsetHeight > 70000 || el.querySelector(".ph-card");
      PH.toggle(el, "ph-still", !!big);
    });
  }
  /* is the page keeping up? counted once, while it appears for the first time in front of an audience */
  function measure() {
    var n = 0, t0 = 0;
    if (measured || quiet) { return; }
    measured = true;
    function f(t) {
      if (!t0) { t0 = t; }
      n++;
      if (t - t0 < 700) { window.requestAnimationFrame(f); return; }
      if ((n - 1) * 1000 / (t - t0) < 22) { setLite(true); }
    }
    window.requestAnimationFrame(f);
  }
  function runShow() {
    var i, force;
    clearTimers(); stopAnims();
    markStill();
    PH.remove(document.body, "ph-in");
    force = document.body.offsetWidth;
    if (quiet || lite) { PH.add(document.body, "ph-in"); } else { setTimeout(function () { PH.add(document.body, "ph-in"); }, 40); }
    for (i = 0; i < showFns.length; i++) { try { showFns[i](); } catch (e) { if (window.console) { window.console.error(e); } } }
    if (PH.afterShow) { try { PH.afterShow(); } catch (e2) { } }            /* kit-deck.js lays the kept edits over the page here */
    measure();
    return force;
  }
  window.pptHtmlOnShow = function () { shown = true; quiet = false; PH.toggle(document.documentElement, "ph-lite", lite); runShow(); };
  window.pptHtmlOnHide = function () { var i; shown = false; clearTimers(); stopAnims(); for (i = 0; i < hideFns.length; i++) { try { hideFns[i](); } catch (e) { } } };

  /* ---------------------------------------------------------------- site header */
  /* PH.nav({ brand: "서비스 이름", links: [["개요", 1], ["실적", 2]], page: 2, next: true })   (numbers are slide numbers) */
  PH.nav = function (cfg) {
    var nav = PH.make("div", "ph-nav"), brand = PH.make("div", "ph-brand", null, nav), links = PH.make("div", "ph-links", null, nav), i, s, cta, last;
    PH.make("i", "", null, brand); brand.appendChild(document.createTextNode(cfg.brand || ""));
    cfg.links = cfg.links || [];
    if (cfg.links.length > 7) { PH.add(links, "ph-many"); }
    if (cfg.links.length > 12) { PH.add(links, "ph-lots"); }
    if (cfg.links.length) { brand.setAttribute("data-goto", cfg.links[0][1]); }
    for (i = 0; i < cfg.links.length; i++) {
      s = PH.make("span", cfg.links[i][1] === cfg.page ? "is-on" : "", cfg.links[i][0], links);
      if (cfg.links[i][1] !== cfg.page) { s.setAttribute("data-goto", cfg.links[i][1]); }
    }
    if (cfg.next !== false && cfg.links.length) {
      last = cfg.links[cfg.links.length - 1][1];
      cta = PH.make("div", "ph-cta", null, nav);
      if (cfg.page !== last) { cta.appendChild(document.createTextNode(cfg.nextText || "다음")); PH.make("i", "", "→", cta); cta.setAttribute("data-next", ""); }
      else { cta.appendChild(document.createTextNode(cfg.firstText || "처음으로")); PH.make("i", "", "↺", cta); cta.setAttribute("data-goto", cfg.links[0][1]); }
    }
    document.body.appendChild(nav);
    return nav;
  };
  /* a ring that follows the pointer and grows over anything clickable */
  PH.ring = function () {
    var c = PH.make("div", "ph-ring", null, document.body), big = false;
    PH.on(document, "mousemove", function (e) { c.style.left = e.clientX + "px"; c.style.top = e.clientY + "px"; });
    PH.on(document, "mouseover", function (e) {
      var el = e.target, hot = false, n = 0, k;
      while (el && el !== document.body && n < 8) {
        k = el.getAttribute ? (el.getAttribute("class") || "") : "";
        if (/ph-(hot|btn|chip|cta|step|acc-q|switch|tl-pt|tr|pick)\b/.test(k) || (el.hasAttribute && (el.hasAttribute("data-goto") || el.hasAttribute("data-next") || el.hasAttribute("data-tab") || el.hasAttribute("data-zoom")))) { hot = true; break; }
        el = el.parentNode; n++;
      }
      if (hot !== big) { big = hot; PH.toggle(c, "is-big", hot); }
    });
    return c;
  };
  /* text that appears as if it were being typed */
  PH.type = function (el, text, done, speed) {
    var i = 0, node = document.createTextNode(""), caret = PH.make("span", "ph-caret"), id;
    if (quiet || lite) { el.textContent = text; if (done) { done(); } return 0; }
    el.textContent = ""; el.appendChild(node); el.appendChild(caret);
    id = PH.every(speed || 30, function () {
      i += (text.charAt(i) === " " ? 2 : 1);
      if (i >= text.length) { PH.stop(id); if (caret.parentNode) { caret.parentNode.removeChild(caret); } node.nodeValue = text; if (done) { done(); } return; }
      node.nodeValue = text.substring(0, i);
    });
    return id;
  };
  /* a number that counts up in about a dozen steps (not every frame: each change repaints its surroundings) */
  PH.countUp = function (el, to, digits, ms) {
    var n = 0, steps = 11, id;
    if (quiet || lite) { el.textContent = PH.fmt(to, digits || 0); return; }
    el.textContent = PH.fmt(0, digits || 0);
    id = PH.every(Math.round((ms || 560) / steps), function () {
      n++;
      el.textContent = PH.fmt(to * PH.ease(n / steps), digits || 0);
      if (n >= steps) { PH.stop(id); el.textContent = PH.fmt(to, digits || 0); }
    });
  };

  /* ---------------------------------------------------------------- markup behaviours */
  function selectTab(tab) {
    var id = tab.getAttribute("data-tab"), sibs = tab.parentNode.children, i, panel = document.getElementById(id), ps;
    for (i = 0; i < sibs.length; i++) { if (sibs[i].hasAttribute("data-tab")) { PH.toggle(sibs[i], "is-on", sibs[i] === tab); } }
    if (panel) {
      ps = panel.parentNode.children;
      for (i = 0; i < ps.length; i++) { if (PH.has(ps[i], "ph-panel")) { PH.toggle(ps[i], "is-on", ps[i] === panel); } }
    }
    PH.fire(tab, "ph-tab", id);
  }
  PH.selectTab = function (id) { var t = PH.$('[data-tab="' + id + '"]'); if (t) { selectTab(t); } };
  PH.on(document, "click", function (e) {
    var t = e.target, el;
    if ((el = PH.upAttr(t, "data-goto"))) { PH.goto(el.getAttribute("data-goto")); return; }
    if ((el = PH.upAttr(t, "data-next"))) { PH.next(); return; }
    if ((el = PH.upAttr(t, "data-prev"))) { PH.prev(); return; }
    if ((el = PH.upAttr(t, "data-tab"))) { selectTab(el); return; }
    if ((el = PH.up(t, "ph-acc-q"))) {
      var acc = el.parentNode, wrap = acc.parentNode, open = !PH.has(acc, "is-on"), k;
      if (open && wrap.hasAttribute("data-single")) { for (k = 0; k < wrap.children.length; k++) { PH.remove(wrap.children[k], "is-on"); } }
      PH.toggle(acc, "is-on", open); PH.fire(acc, "ph-change", open); return;
    }
    if ((el = PH.up(t, "ph-switch"))) { PH.fire(el, "ph-change", PH.toggle(el, "is-on")); return; }
  });
  function fillStatics() {
    PH.$$("[data-count]").forEach(function (el) { PH.countUp(el, Number(el.getAttribute("data-count")), Number(el.getAttribute("data-digits") || 0)); });
    PH.$$(".ph-meter i[data-w]").forEach(function (el) {
      if (quiet || lite) { el.style.width = el.getAttribute("data-w") + "%"; return; }
      el.style.width = "0"; PH.wait(90, function () { el.style.width = el.getAttribute("data-w") + "%"; });
    });
  }
  PH.onShow(fillStatics);

  /* steps: PH.steps(container, function (index, stepElement) { ... }, autoAdvanceMs)   children: .ph-step */
  PH.steps = function (box, onChange, autoMs) {
    var list = PH.$$(".ph-step", box), cur = -1, timer = 0;
    function set(i, byUser) {
      var k;
      cur = (i + list.length) % list.length;
      for (k = 0; k < list.length; k++) { PH.toggle(list[k], "is-on", k === cur); PH.toggle(list[k], "is-done", k < cur); }
      if (onChange) { onChange(cur, list[cur]); }
      if (byUser && timer) { PH.stop(timer); timer = 0; }
    }
    list.forEach(function (el, i) { PH.on(el, "click", function () { set(i, true); }); });
    PH.onShow(function () { set(0); if (autoMs) { timer = PH.every(autoMs, function () { set(cur + 1); }); } });
    return { set: set, get: function () { return cur; } };
  };
  /* timeline: PH.timeline(el, [{ when: "3월", what: "시범 운영" }, ...], function (index, item) { ... }, startIndex) */
  PH.timeline = function (el, items, onPick, start) {
    var pts = [], fill, cur = -1;
    PH.add(el, "ph-tl"); el.textContent = "";
    PH.make("div", "ph-tl-line", null, el); fill = PH.make("div", "ph-tl-fill", null, el);
    function set(i) {
      var k; cur = i;
      for (k = 0; k < pts.length; k++) { PH.toggle(pts[k], "is-on", k === i); PH.toggle(pts[k], "is-done", k < i); }
      fill.style.width = ((i + 0.5) / items.length * 100) + "%";
      if (onPick) { onPick(i, items[i]); }
    }
    items.forEach(function (it, i) {
      var p = PH.make("div", "ph-tl-pt", null, el);
      p.style.left = ((i + 0.5) / items.length * 100) + "%";
      PH.make("b", "", it.when, p); PH.make("i", "", null, p); PH.make("span", "", it.what, p);
      PH.on(p, "click", function () { set(i); });
      pts.push(p);
    });
    PH.onShow(function () { fill.style.width = "0"; PH.wait(120, function () { set(start || 0); }); });
    return { set: set, get: function () { return cur; } };
  };
  /* slider: PH.slider(el, { min: 0, max: 100, step: 5, value: 40 }, function (value) { ... }) */
  PH.slider = function (el, cfg, onChange) {
    var fill, knob, val = cfg.value, drag = false;
    PH.add(el, "ph-slider"); el.textContent = "";
    PH.make("div", "ph-slider-track", null, el); fill = PH.make("div", "ph-slider-fill", null, el); knob = PH.make("div", "ph-slider-knob", null, el);
    function paint() { var p = (val - cfg.min) / (cfg.max - cfg.min) * 100; fill.style.width = p + "%"; knob.style.left = p + "%"; }
    function set(v, quietly) {
      var st = cfg.step || 1;
      v = Math.round((v - cfg.min) / st) * st + cfg.min;
      v = PH.clamp(Number(v.toFixed(6)), cfg.min, cfg.max);
      if (v !== val || quietly) { val = v; paint(); if (onChange && !quietly) { onChange(val); } }
    }
    function fromEvent(e) { var r = el.getBoundingClientRect(); set(cfg.min + PH.clamp((e.clientX - r.left) / r.width, 0, 1) * (cfg.max - cfg.min)); }
    PH.on(el, "mousedown", function (e) { drag = true; fromEvent(e); e.preventDefault(); });
    PH.on(document, "mousemove", function (e) { if (drag) { fromEvent(e); } });
    PH.on(document, "mouseup", function () { drag = false; });
    paint();
    return { set: function (v) { set(v, true); if (onChange) { onChange(val); } }, get: function () { return val; } };
  };

  /* table: PH.table(el, { cols: [{ key: "name", label: "매장" }, { key: "sales", label: "매출", num: true, bar: true, digits: 0 },
                                  { key: "state", label: "상태", badge: { "정상": "ph-good", "주의": "ph-warn" } }],
                            rows: [{ name: "...", sales: 1200, state: "정상" }], sort: "sales", desc: true }, function (row) { ... }) */
  PH.table = function (el, cfg, onPick) {
    var rows = cfg.rows.slice(), sortKey = cfg.sort || null, desc = cfg.desc !== false, picked = null, table = PH.make("table", "ph-table", null, el), thead = PH.make("thead", "", null, table), tbody = PH.make("tbody", "", null, table);
    function colOf(key) { var i; for (i = 0; i < cfg.cols.length; i++) { if (cfg.cols[i].key === key) { return cfg.cols[i]; } } return null; }
    function draw() {
      var tr = PH.make("tr"), list = rows.slice(), max = {};
      thead.textContent = ""; tbody.textContent = "";
      cfg.cols.forEach(function (c) {
        var th = PH.make("th", (c.num ? "ph-r " : "") + (cfg.sortable === false ? "" : "is-sortable"), c.label, tr);
        PH.make("i", "", sortKey === c.key ? (desc ? "↓" : "↑") : "", th);
        if (cfg.sortable !== false) { PH.on(th, "click", function () { if (sortKey === c.key) { desc = !desc; } else { sortKey = c.key; desc = !!c.num; } draw(); }); }
        if (c.bar) { max[c.key] = 0; rows.forEach(function (r) { if (r[c.key] > max[c.key]) { max[c.key] = r[c.key]; } }); }
      });
      thead.appendChild(tr);
      if (sortKey) {
        list.sort(function (a, b) { var x = a[sortKey], y = b[sortKey], c = colOf(sortKey), d = (c && c.num) ? (x - y) : String(x).localeCompare(String(y)); return desc ? -d : d; });
      }
      list.forEach(function (r) {
        var row = PH.make("tr", "ph-tr" + (r === picked ? " is-on" : ""), null, tbody);
        cfg.cols.forEach(function (c) {
          var td = PH.make("td", c.num ? "ph-r" : "", null, row), v = r[c.key], bar;
          if (c.bar && max[c.key] > 0) { bar = PH.make("span", "ph-bar", null, td); bar.style.width = Math.max(2, Math.round(v / max[c.key] * (c.barWidth || 70))) + "px"; }
          if (c.badge) { PH.make("span", "ph-badge " + (c.badge[v] || "ph-plain"), v, td); }
          else { td.appendChild(document.createTextNode(c.num ? PH.fmt(v, c.digits || 0) + (c.unit || "") : v)); }
        });
        PH.on(row, "click", function () { picked = r; draw(); if (onPick) { onPick(r); } });
      });
    }
    draw();
    return { update: function (list) { rows = list.slice(); if (rows.indexOf(picked) < 0) { picked = null; } draw(); }, pick: function (r) { picked = r; draw(); if (onPick && r) { onPick(r); } } };
  };

  /* ---------------------------------------------------------------- charts */
  function niceMax(v) { var p, m; if (!(v > 0)) { return 1; } p = Math.pow(10, Math.floor(Math.log(v) / Math.LN10)); m = v / p; return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p; }
  function chartBox(el) {
    var W, H;
    PH.add(el, "ph-chart"); PH.remove(el, "is-in"); el.textContent = "";
    if (window.getComputedStyle(el).position === "static") { el.style.position = "relative"; }
    W = el.clientWidth; H = el.clientHeight;
    return { W: W, H: H, svg: PH.svg("svg", { width: W, height: H, viewBox: "0 0 " + W + " " + H }, el), tip: PH.make("div", "ph-tip", null, el) };
  }
  function showTip(box, html, x, y) {
    var tip = box.tip, w;
    tip.innerHTML = html; PH.add(tip, "is-on");
    w = tip.offsetWidth;
    tip.style.left = PH.clamp(x - w / 2, 0, box.W - w) + "px"; tip.style.top = Math.max(0, y - tip.offsetHeight - 10) + "px";
  }
  function hideTip(box) { PH.remove(box.tip, "is-on"); }
  function seriesOf(cfg) { var s = cfg.series || [{ name: "", values: cfg.values || [] }], pal = PH.palette(s.length), i, out = []; for (i = 0; i < s.length; i++) { out.push({ name: s[i].name, values: s[i].values, color: s[i].color ? PH.color(s[i].color) : pal[i], off: false }); } return out; }
  function legend(el, series, redraw) {
    var box = PH.make("div", "ph-legend", null, el);
    box.style.cssText = "position:absolute;right:0;top:-4px;";
    series.forEach(function (s) {
      var item = PH.make("span", s.off ? "is-off" : "", null, box), dot = PH.make("i", "", null, item);
      dot.style.background = s.color; item.appendChild(document.createTextNode(s.name));
      PH.on(item, "click", function () { var live = series.filter(function (x) { return !x.off; }).length; if (!s.off && live < 2) { return; } s.off = !s.off; redraw(); });
    });
  }
  function axes(box, m, max, labels, unit, band) {
    var g = PH.svg("g", {}, box.svg), i, y, n = labels.length, x, pw = box.W - m.l - m.r, ph = box.H - m.t - m.b;
    for (i = 0; i <= 4; i++) {
      y = Math.round(m.t + ph - ph * i / 4) + 0.5;
      PH.svg("line", { x1: m.l, x2: box.W - m.r, y1: y, y2: y, "class": "ph-axis", "stroke-dasharray": i ? "2 4" : "" }, g);
      PH.svg("text", { x: m.l - 8, y: y + 4, "text-anchor": "end" }, g).textContent = PH.fmt(max * i / 4, max < 10 ? 1 : 0) + (i === 4 && unit ? unit : "");
    }
    for (i = 0; i < n; i++) {
      x = band ? m.l + pw * (i + 0.5) / n : m.l + (n > 1 ? pw * i / (n - 1) : pw / 2);
      PH.svg("text", { x: r1(x), y: box.H - m.b + 20, "text-anchor": "middle" }, g).textContent = labels[i];
    }
  }
  function tipHtml(label, series, i, unit, digits) {
    var s = "<b>" + label + "</b>", k;
    for (k = 0; k < series.length; k++) { if (!series[k].off) { s += "<br>" + (series.length > 1 ? series[k].name + " " : "") + PH.fmt(series[k].values[i], digits || 0) + (unit || ""); } }
    return s;
  }
  function reveal(el) { if (quiet || lite) { PH.add(el, "is-in"); } else { PH.wait(40, function () { PH.add(el, "is-in"); }); } }

  /* PH.bars(el, { labels: ["1월", ...], series: [{ name: "매출", values: [..] }, ...] | values: [..], unit: "억", max, digits }, onPick)
     The bars are plain boxes that grow by a CSS transition: no script runs per frame. */
  PH.bars = function (el, cfg, onPick) {
    var series = seriesOf(cfg), box, picked = -1, cols = [];
    function dim() { cols.forEach(function (c, k) { PH.toggle(c, "is-dim", picked >= 0 && k !== picked); }); }
    function draw(animate) {
      var m = { l: 46, r: 8, t: series.length > 1 ? 30 : 14, b: 30 }, live = series.filter(function (s) { return !s.off; }), n = cfg.labels.length, max = 0, pw, ph, gw, bw;
      box = chartBox(el); cols = [];
      if (series.length > 1) { legend(el, series, function () { draw(true); }); }
      live.forEach(function (s) { s.values.forEach(function (v) { if (v > max) { max = v; } }); });
      max = cfg.max || niceMax(max);
      pw = box.W - m.l - m.r; ph = box.H - m.t - m.b; gw = pw / n; bw = Math.min(54, gw * 0.72 / live.length);
      axes(box, m, max, cfg.labels, cfg.unit, true);
      cfg.labels.forEach(function (lab, i) {
        var col = PH.make("div", "ph-col", null, el), x0 = (gw - bw * live.length) / 2, top = 0;
        col.style.cssText = "left:" + r1(m.l + gw * i) + "px;top:" + m.t + "px;width:" + r1(gw) + "px;height:" + ph + "px;" + (onPick ? "cursor:pointer;" : "");
        live.forEach(function (s, k) {
          var b = PH.make("div", "ph-colbar", null, col), h = Math.max(1, ph * s.values[i] / max), gapPx = live.length > 1 ? 1.5 : 0;
          b.style.cssText = "left:" + r1(x0 + bw * k + gapPx) + "px;width:" + r1(bw - gapPx * 2) + "px;height:" + r1(h) + "px;background:" + s.color + ";transition-delay:" + (i * 30) + "ms;";
          if (h > top) { top = h; }
        });
        PH.on(col, "mouseover", function () { cols.forEach(function (c, k) { PH.toggle(c, "is-dim", k !== i); }); showTip(box, tipHtml(lab, series, i, cfg.unit, cfg.digits), m.l + gw * (i + 0.5), m.t + ph - top); });
        PH.on(col, "mouseout", function () { dim(); hideTip(box); });
        if (onPick) { PH.on(col, "click", function () { picked = i; dim(); onPick(i, lab); }); }
        cols.push(col);
      });
      dim();
      if (animate) { reveal(el); } else { PH.add(el, "is-in"); }
    }
    PH.onShow(function () { draw(true); }); PH.onResize(function () { draw(false); });
    return { update: function (c) { var k; for (k in c) { if (c.hasOwnProperty(k)) { cfg[k] = c[k]; } } series = seriesOf(cfg); picked = -1; draw(true); }, pick: function (i) { picked = i; dim(); } };
  };

  /* PH.lines(el, { labels: [...], series: [{ name, values }], unit, max, digits, area: true }) */
  PH.lines = function (el, cfg) {
    var series = seriesOf(cfg), box;
    function draw(animate) {
      var m = { l: 46, r: 14, t: series.length > 1 ? 30 : 14, b: 30 }, live = series.filter(function (s) { return !s.off; }), n = cfg.labels.length, max = 0, pw, ph, lines = [], guide, dots = [], hit;
      box = chartBox(el);
      if (series.length > 1) { legend(el, series, function () { draw(true); }); }
      live.forEach(function (s) { s.values.forEach(function (v) { if (v > max) { max = v; } }); });
      max = cfg.max || niceMax(max);
      pw = box.W - m.l - m.r; ph = box.H - m.t - m.b;
      axes(box, m, max, cfg.labels, cfg.unit, false);
      function px(i) { return m.l + (n > 1 ? pw * i / (n - 1) : pw / 2); }
      function py(v) { return m.t + ph - ph * v / max; }
      live.forEach(function (s, k) {
        var d = "", i, len = 0, path, area;
        for (i = 0; i < n; i++) { d += (i ? "L" : "M") + r1(px(i)) + " " + r1(py(s.values[i])); if (i) { len += Math.sqrt(Math.pow(px(i) - px(i - 1), 2) + Math.pow(py(s.values[i]) - py(s.values[i - 1]), 2)); } }
        if (cfg.area !== false && k === 0) { area = PH.svg("path", { d: d + "L" + r1(px(n - 1)) + " " + (m.t + ph) + "L" + r1(px(0)) + " " + (m.t + ph) + "Z", fill: s.color, "fill-opacity": 0 }, box.svg); }
        path = PH.svg("path", { d: d, fill: "none", stroke: s.color, "stroke-width": 3, "stroke-linecap": "round", "stroke-linejoin": "round", "stroke-dasharray": r1(len + 2), "stroke-dashoffset": r1(len + 2) }, box.svg);
        lines.push({ path: path, len: len + 2, area: area });
        dots.push(PH.svg("circle", { r: 5, fill: PH.color("surface"), stroke: s.color, "stroke-width": 3, opacity: 0 }, box.svg));
      });
      guide = PH.svg("line", { y1: m.t, y2: m.t + ph, "class": "ph-axis", opacity: 0 }, box.svg);
      hit = PH.svg("rect", { x: m.l, y: m.t, width: pw, height: ph, fill: "#000000", "fill-opacity": 0 }, box.svg);
      PH.on(hit, "mousemove", function (e) {
        var r = box.svg.getBoundingClientRect(), x = (e.clientX - r.left) * box.W / r.width, i = PH.clamp(Math.round((x - m.l) / pw * (n - 1)), 0, n - 1), top = 1e9;
        guide.setAttribute("x1", r1(px(i))); guide.setAttribute("x2", r1(px(i))); guide.setAttribute("opacity", 1);
        live.forEach(function (s, k) { var y = py(s.values[i]); dots[k].setAttribute("cx", r1(px(i))); dots[k].setAttribute("cy", r1(y)); dots[k].setAttribute("opacity", 1); if (y < top) { top = y; } });
        showTip(box, tipHtml(cfg.labels[i], series, i, cfg.unit, cfg.digits), px(i), top);
      });
      PH.on(hit, "mouseout", function () { guide.setAttribute("opacity", 0); dots.forEach(function (d) { d.setAttribute("opacity", 0); }); hideTip(box); });
      function shape(t) { lines.forEach(function (o) { o.path.setAttribute("stroke-dashoffset", r1(o.len * (1 - t))); if (o.area && (t === 1 || t === 0)) { o.area.setAttribute("fill-opacity", t ? 0.13 : 0); } }); }
      if (animate) { PH.animate(700, shape); } else { shape(1); }
    }
    PH.onShow(function () { draw(true); }); PH.onResize(function () { draw(false); });
    return { update: function (c) { var k; for (k in c) { if (c.hasOwnProperty(k)) { cfg[k] = c[k]; } } series = seriesOf(cfg); draw(true); } };
  };

  /* PH.donut(el, { items: [{ label: "매장", value: 62 }, ...], unit: "%", center: "합계", digits: 0 }, onPick)
     Draws the ring on the left and a clickable list on the right. */
  PH.donut = function (el, cfg, onPick) {
    var box, arcs = [], rowsEl = [], pal, total = 0, big, small, size, R, C;
    function label(i) {
      var it = i < 0 ? null : cfg.items[i];
      big.textContent = it ? PH.fmt(it.value, cfg.digits || 0) + (cfg.unit || "") : PH.fmt(total, cfg.digits || 0) + (cfg.unit || "");
      small.textContent = it ? it.label + " · " + Math.round(it.value / total * 100) + "%" : (cfg.center || "합계");
    }
    function focus(i) {
      arcs.forEach(function (a, k) { a.setAttribute("opacity", i < 0 || k === i ? 1 : 0.28); });
      rowsEl.forEach(function (r, k) { r.style.opacity = i < 0 || k === i ? 1 : 0.45; });
      label(i);
    }
    function draw(animate) {
      var acc = 0, list, g;
      PH.add(el, "ph-chart"); el.textContent = ""; arcs = []; rowsEl = []; total = 0;
      if (window.getComputedStyle(el).position === "static") { el.style.position = "relative"; }
      pal = PH.palette(cfg.items.length);
      cfg.items.forEach(function (it) { total += it.value; });
      size = Math.min(el.clientHeight, el.clientWidth * 0.5); R = size * 0.38; C = 2 * Math.PI * R;
      box = { W: el.clientWidth, H: el.clientHeight, svg: PH.svg("svg", { width: size, height: size, viewBox: "0 0 " + size + " " + size }, el) };
      box.svg.style.top = Math.round((el.clientHeight - size) / 2) + "px";
      g = PH.svg("g", { transform: "rotate(-90 " + size / 2 + " " + size / 2 + ")" }, box.svg);
      PH.svg("circle", { cx: size / 2, cy: size / 2, r: R, fill: "none", stroke: PH.color("line"), "stroke-width": size * 0.13, opacity: 0.6 }, g);
      cfg.items.forEach(function (it, i) {
        var frac = it.value / total, a = PH.svg("circle", { cx: size / 2, cy: size / 2, r: R, fill: "none", stroke: it.color ? PH.color(it.color) : pal[i], "stroke-width": size * 0.13, "stroke-dasharray": "0 " + r1(C), "stroke-dashoffset": r1(-acc * C), style: "cursor:pointer" }, g);
        a.ihFrac = frac; acc += frac; arcs.push(a);
        PH.on(a, "mouseover", function () { focus(i); }); PH.on(a, "mouseout", function () { focus(-1); });
        if (onPick) { PH.on(a, "click", function () { onPick(i, it); }); }
      });
      big = PH.svg("text", { x: size / 2, y: size / 2 + size * 0.03, "text-anchor": "middle", "class": "ph-val", style: "font-size:" + Math.round(size * 0.15) + "px" }, box.svg);
      small = PH.svg("text", { x: size / 2, y: size / 2 + size * 0.13, "text-anchor": "middle", style: "font-size:" + Math.max(11, Math.round(size * 0.055)) + "px" }, box.svg);
      list = PH.make("div", "", null, el);
      list.style.cssText = "position:absolute;left:" + Math.round(size + 28) + "px;right:0;top:50%;";
      cfg.items.forEach(function (it, i) {
        var row = PH.make("div", "", null, list), dot = PH.make("i", "", null, row), val;
        row.style.cssText = "position:relative;height:34px;line-height:34px;font-size:14.5px;cursor:pointer;white-space:nowrap;";
        dot.style.cssText = "display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:10px;background:" + (it.color ? PH.color(it.color) : pal[i]);
        row.appendChild(document.createTextNode(it.label));
        val = PH.make("b", "", PH.fmt(it.value, cfg.digits || 0) + (cfg.unit || ""), row);
        val.style.cssText = "position:absolute;right:0;top:0;";
        PH.on(row, "mouseover", function () { focus(i); }); PH.on(row, "mouseout", function () { focus(-1); });
        if (onPick) { PH.on(row, "click", function () { onPick(i, it); }); }
        rowsEl.push(row);
      });
      list.style.marginTop = (-cfg.items.length * 17) + "px";
      label(-1);
      function sweep(t) { arcs.forEach(function (a) { a.setAttribute("stroke-dasharray", r1(Math.max(0, a.ihFrac * C * t - 2)) + " " + r1(C)); }); }
      if (animate === false) { sweep(1); } else { PH.animate(650, sweep); }
    }
    PH.onShow(function () { draw(true); }); PH.onResize(function () { draw(false); });
    return { update: function (c) { var k; for (k in c) { if (c.hasOwnProperty(k)) { cfg[k] = c[k]; } } draw(true); } };
  };

  /* PH.spark(el, [3, 5, 4, 8, 7], "c1")   a small trend line that fills its element */
  PH.spark = function (el, values, color) {
    function draw(animate) {
      var box = chartBox(el), max = Math.max.apply(null, values), min = Math.min.apply(null, values), d = "", i, x, y, len = 0, px, py, path;
      for (i = 0; i < values.length; i++) {
        x = 3 + (box.W - 6) * i / (values.length - 1); y = 3 + (box.H - 6) * (1 - (values[i] - min) / ((max - min) || 1));
        d += (i ? "L" : "M") + r1(x) + " " + r1(y);
        if (i) { len += Math.sqrt(Math.pow(x - px, 2) + Math.pow(y - py, 2)); }
        px = x; py = y;
      }
      path = PH.svg("path", { d: d, fill: "none", stroke: PH.color(color || "c1"), "stroke-width": 2.5, "stroke-linecap": "round", "stroke-linejoin": "round", "stroke-dasharray": r1(len + 2), "stroke-dashoffset": r1(len + 2) }, box.svg);
      PH.svg("circle", { cx: r1(px), cy: r1(py), r: 3.5, fill: PH.color(color || "c1") }, box.svg);
      if (animate === false) { path.setAttribute("stroke-dashoffset", 0); } else { PH.animate(600, function (t) { path.setAttribute("stroke-dashoffset", r1((len + 2) * (1 - t))); }); }
    }
    PH.onShow(function () { draw(true); }); PH.onResize(function () { draw(false); });
  };

  /* ---------------------------------------------------------------- start
     The page's own script runs after this file and registers its PH.onShow callbacks; the first "show" therefore
     waits for the document to finish loading. Inside PowerPoint (and in the web version) every page is loaded
     before its slide is on screen: it is then drawn in its final state without motion, and the motion plays
     when the host says the slide has appeared. */
  function first() {
    var n = 0, id;
    lastVW = document.documentElement.clientWidth;
    if (!shown) { quiet = PH.hosted(); PH.toggle(document.documentElement, "ph-lite", quiet || lite); runShow(); }
    id = setInterval(function () { checkSize(); if (++n > 20) { clearInterval(id); } }, 200);
  }
  if (document.readyState === "complete") { setTimeout(first, 0); } else { PH.on(window, "load", first); }
}(window, document));
