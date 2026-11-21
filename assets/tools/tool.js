/* ppt-html build tool page (runs in Edge headless, never inside a presentation).
   IHTool.convert  : modern CSS / JavaScript of one page -> IE11 (see convert.js)
   IHTool.show     : load a page in a frame of exactly the canvas size (reference picture, font survey)
   IHTool.fontUse  : which font stacks / weights the shown page uses
   IHTool.subset   : cut a font down to the characters a page needs (see fonttool.js) */
(function () {
  var frame = null;
  function doc() { return frame && frame.contentDocument; }
  window.IHTool = {
    convert: function (jobText) {
      return JSON.stringify(IHConvert.convertPage(JSON.parse(jobText), window.Babel));
    },
    show: function (html, w, h, settleMs) {
      return new Promise(function (resolve) {
        if (frame) { frame.remove(); }
        document.documentElement.style.cssText = "margin:0;padding:0;overflow:hidden;background:#fff";
        document.body.style.cssText = "margin:0;padding:0;overflow:hidden";
        frame = document.createElement("iframe");
        frame.style.cssText = "position:fixed;left:0;top:0;border:0;margin:0;padding:0;width:" + w + "px;height:" + h + "px;background:#fff";
        var done = false;
        function finish(how) { if (done) { return; } done = true; setTimeout(function () { resolve(how); }, settleMs || 0); }
        frame.onload = function () { finish("loaded"); };
        setTimeout(function () { finish("timeout"); }, 15000);
        frame.setAttribute("sandbox", "allow-scripts allow-same-origin");
        frame.srcdoc = html;
        document.body.appendChild(frame);
      });
    },
    hide: function () { if (frame) { frame.remove(); frame = null; } return "ok"; },
    measure: function () {
      var d = doc(); if (!d || !d.documentElement) { return "{}"; }
      var e = d.documentElement, b = d.body;
      return JSON.stringify({ scrollW: Math.max(e.scrollWidth, b ? b.scrollWidth : 0), scrollH: Math.max(e.scrollHeight, b ? b.scrollHeight : 0), clientW: e.clientWidth, clientH: e.clientHeight });
    },
    /* the shown page: its background colour, the script errors its harness collected, whether its script ran to the end, its size */
    info: function () {
      var d = doc(); if (!d || !d.documentElement) { return "{}"; }
      var w = frame.contentWindow, e = d.documentElement, b = d.body;
      function col(el) {
        if (!el) { return ""; }
        var m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(w.getComputedStyle(el).backgroundColor);
        if (!m || (m[4] !== undefined && parseFloat(m[4]) < 0.5)) { return ""; }
        return "#" + [m[1], m[2], m[3]].map(function (x) { return ("0" + parseInt(x, 10).toString(16)).slice(-2); }).join("").toUpperCase();
      }
      return JSON.stringify({
        bg: col(b) || col(e), errors: (w.__ihErr || []).join(" | "), ready: !!(b && b.getAttribute("data-ih-ready") === "1"), title: d.title || "",
        scrollW: Math.max(e.scrollWidth, b ? b.scrollWidth : 0), scrollH: Math.max(e.scrollHeight, b ? b.scrollHeight : 0), clientW: e.clientWidth, clientH: e.clientHeight
      });
    },
    /* every font stack + weight that text of the page is set in (hidden elements included: they may be shown later) */
    fontUse: function () {
      var d = doc(); if (!d || !d.body) { return "[]"; }
      var seen = {}, out = [], win = frame.contentWindow;
      function add(el, pseudo) {
        var cs = win.getComputedStyle(el, pseudo || null);
        if (pseudo) { var c = cs.content; if (!c || c === "none" || c === "normal" || c === '""') { return; } }
        var w = parseInt(cs.fontWeight, 10) || (cs.fontWeight === "bold" ? 700 : 400);
        var key = cs.fontFamily + "|" + w;
        if (!seen[key]) { seen[key] = 1; out.push({ stack: cs.fontFamily, weight: w }); }
      }
      var all = d.body.getElementsByTagName("*");
      for (var i = 0; i < all.length; i++) {
        var el = all[i], tag = el.tagName.toLowerCase();
        if (tag === "script" || tag === "style") { continue; }
        var has = tag === "input" || tag === "textarea" || tag === "button" || tag === "select" || tag === "text" || tag === "tspan";
        for (var n = el.firstChild; n && !has; n = n.nextSibling) { if (n.nodeType === 3 && /\S/.test(n.nodeValue)) { has = true; } }
        if (has) { add(el); }
        add(el, "::before"); add(el, "::after");
      }
      return JSON.stringify(out);
    }
  };
}());
