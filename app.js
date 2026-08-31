// app.js — merged runtime for the IW28/IW38 dashboard.
// This file combines three originally-separate scripts into one, so the
// project ships as a single HTML + single JS file:
//   1) support.js   — GENERATED dc-runtime (do not hand-edit this section;
//                      rebuild with `cd dc-runtime && bun run build` and
//                      re-paste it in if the runtime itself ever changes)
//   2) iw-core.js   — hand-written pure logic (xlsx parsing + metrics),
//                      originally an ES module, now exposed as window.IWCore
//   3) firebase-init.js — hand-written shared-storage wiring, originally
//                      loaded as <script type="module">, now using
//                      dynamic import() so it can run as a classic script

// GENERATED from dc-runtime/src/*.ts — do not edit. Rebuild with `cd dc-runtime && bun run build`.
"use strict";
(() => {
  var __defProp = Object.defineProperty;
  var __defNormalProp = (obj, key, value) => key in obj ? __defProp(obj, key, { enumerable: true, configurable: true, writable: true, value }) : obj[key] = value;
  var __publicField = (obj, key, value) => __defNormalProp(obj, typeof key !== "symbol" ? key + "" : key, value);

  // src/react.ts
  function getReact() {
    const R = window.React;
    if (!R) throw new Error("dc-runtime: window.React is not available yet");
    return R;
  }
  function getReactDOM() {
    const RD = window.ReactDOM;
    if (!RD) throw new Error("dc-runtime: window.ReactDOM is not available yet");
    return RD;
  }
  var h = ((...args) => getReact().createElement(
    ...args
  ));

  // src/parse.ts
  function parseDcDocument(doc) {
    const dc = doc.querySelector("x-dc");
    if (!dc) return null;
    const scriptEl = doc.querySelector("script[data-dc-script]");
    const { props, preview } = parseDataProps(
      scriptEl?.getAttribute("data-props") ?? null
    );
    return {
      template: dc.innerHTML,
      js: scriptEl ? scriptEl.textContent || "" : "",
      props,
      preview
    };
  }
  function parseDcText(src) {
    const openMatch = /<x-dc(?:\s[^>]*)?>/.exec(src);
    if (!openMatch) return null;
    const close = src.lastIndexOf("</x-dc>");
    if (close === -1 || close < openMatch.index) return null;
    const template = src.slice(openMatch.index + openMatch[0].length, close);
    const doc = new DOMParser().parseFromString(src, "text/html");
    const scriptEl = doc.querySelector("script[data-dc-script]");
    const { props, preview } = parseDataProps(
      scriptEl?.getAttribute("data-props") ?? null
    );
    return {
      template,
      js: scriptEl ? scriptEl.textContent || "" : "",
      props,
      preview
    };
  }
  function parseDataProps(raw) {
    if (!raw) return { props: null, preview: null };
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      return { props: null, preview: null };
    }
    if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
      return { props: null, preview: null };
    }
    const obj = parsed;
    const preview = obj.$preview && typeof obj.$preview === "object" ? obj.$preview : null;
    const rest = {};
    for (const k of Object.keys(obj)) {
      if (k[0] !== "$") rest[k] = obj[k];
    }
    return { props: Object.keys(rest).length ? rest : null, preview };
  }
  function dcNameFromPath(pathname) {
    let p = pathname || "";
    try {
      p = decodeURIComponent(p);
    } catch {
    }
    const base = p.split("/").pop() || "Root";
    return base.replace(/\.dc\.html$/, "").replace(/\.html?$/, "") || "Root";
  }

  // src/boot.ts
  var BASE_CSS = `
    .sc-placeholder{background:color-mix(in srgb,currentColor 8%,transparent);
      border:1px solid color-mix(in srgb,currentColor 50%,transparent);
      border-radius:2px;box-sizing:border-box;overflow:hidden}
    @keyframes sc-shine{0%{background-position:100% 50%}100%{background-position:0% 50%}}
    html.sc-dc-streaming .sc-placeholder,
    html.sc-dc-streaming .sc-interp.sc-missing{position:relative;
      background:color-mix(in srgb,currentColor 5%,transparent);
      border-color:transparent}
    html.sc-dc-streaming .sc-placeholder::before,
    html.sc-dc-streaming .sc-interp.sc-missing::before{content:'';
      position:absolute;inset:0;pointer-events:none;
      background:linear-gradient(90deg,rgba(217,119,87,0) 25%,rgba(247,225,211,.95) 37%,rgba(217,119,87,0) 63%);
      background-size:400% 100%;animation:sc-shine 1.4s ease infinite}
    html.sc-dc-streaming .sc-placeholder:nth-child(n+9 of .sc-placeholder)::before,
    html.sc-dc-streaming .sc-interp.sc-missing:nth-child(n+9 of .sc-interp.sc-missing)::before{animation:none;
      background:color-mix(in srgb,currentColor 8%,transparent)}
    .sc-placeholder-error{padding:4px 8px;font:11px/1.4 ui-monospace,monospace;
      color:color-mix(in srgb,currentColor 70%,transparent);word-break:break-word}
    .sc-interp.sc-missing{display:inline-block;width:2em;height:1em;overflow:hidden;
      vertical-align:text-bottom;background:rgba(255,255,255,.3);border:1px solid rgba(0,0,0,.5);
      border-radius:2px;box-sizing:border-box;color:transparent;
      user-select:none}
    .sc-interp.sc-unresolved{font-family:ui-monospace,monospace;font-size:.85em;
      color:color-mix(in srgb,currentColor 50%,transparent);
      background:color-mix(in srgb,currentColor 10%,transparent);border-radius:3px;
      padding:0 3px}
    .sc-host.sc-has-error{position:relative}
    .sc-logic-error{position:absolute;top:8px;left:8px;z-index:2147483647;max-width:60ch;
      padding:6px 10px;background:#b00020;color:#fff;font:12px/1.4 ui-monospace,monospace;
      border-radius:4px;white-space:pre-wrap;pointer-events:none}
    /* Mirrors PRINT_BASELINE_CSS in apps/web deck-stage-export.ts \u2014 keep both
       in sync until dc-runtime regains a build step. */
    @media print {
      @page { margin: 0.5cm; }
      figure, table { break-inside: avoid; }
      #dc-root, #dc-root > .sc-host { height: auto; }
      *, *::before, *::after {
        print-color-adjust: exact; -webkit-print-color-adjust: exact;
        backdrop-filter: none !important; -webkit-backdrop-filter: none !important;
        animation-delay: -99s !important; animation-duration: .001s !important;
        animation-iteration-count: 1 !important; animation-fill-mode: both !important;
        animation-play-state: running !important; transition-duration: 0s !important;
      }
    }
  `;
  var FULL_PAGE_CSS = "html,body{height:100%;margin:0}#dc-root,#dc-root>.sc-host{height:100%}";
  function rootNameForDocument(doc, loc) {
    let bootPath = loc.pathname || "";
    if (!/\.dc\.html?$/i.test(safeDecode(bootPath))) {
      try {
        bootPath = new URL(doc.baseURI || "/").pathname;
      } catch {
      }
    }
    return dcNameFromPath(bootPath);
  }
  function safeDecode(s) {
    try {
      return decodeURIComponent(s);
    } catch {
      return s;
    }
  }
  function boot(runtime, doc = document) {
    const parsed = parseDcDocument(doc);
    if (!parsed) return null;
    const React = getReact();
    const rootName = rootNameForDocument(doc, location);
    runtime.markFetched(rootName);
    runtime.setRootName(rootName);
    runtime.adoptParsed(rootName, parsed);
    if (!window.__resources) {
      fetch(location.href).then((res) => res.ok ? res.text() : "").then((t) => {
        const raw = t ? parseDcText(t) : null;
        if (raw?.template) runtime.updateHtml(rootName, raw.template);
      }).catch(() => {
      });
    }
    const dc = doc.querySelector("x-dc");
    const hostEl = doc.createElement("div");
    hostEl.id = "dc-root";
    dc.replaceWith(hostEl);
    if (!parsed.preview) {
      const s = doc.createElement("style");
      s.textContent = FULL_PAGE_CSS;
      doc.head.appendChild(s);
    }
    const Root = runtime.getDC(rootName);
    const entry = runtime.registry.get(rootName);
    function StandaloneRoot() {
      const [, setTick] = React.useState(0);
      React.useEffect(() => {
        const sub = () => setTick((n) => n + 1);
        entry.subs.add(sub);
        return () => {
          entry.subs.delete(sub);
        };
      }, []);
      const defaults = React.useMemo(() => {
        const d = {};
        for (const k in entry.propsMeta || {}) {
          const v = entry.propsMeta?.[k]?.default;
          if (v !== void 0) d[k] = v;
        }
        return d;
      }, [entry.propsMeta]);
      return h(Root, { ...defaults, ...entry.propOverrides || {} });
    }
    const ReactDOM = getReactDOM();
    if (ReactDOM.createRoot)
      ReactDOM.createRoot(hostEl).render(h(StandaloneRoot));
    else ReactDOM.render(h(StandaloneRoot), hostEl);
    return rootName;
  }

  // src/expr.ts
  var IDENT_RE = /^[A-Za-z_$][A-Za-z0-9_$]*/;
  var NUMBER_RE = /^-?\d+(\.\d+)?$/;
  function resolve(vals, src) {
    const expr = String(src).trim();
    if (!expr) return void 0;
    if (expr[0] === "(" && expr[expr.length - 1] === ")" && parensWrapWhole(expr)) {
      return resolve(vals, expr.slice(1, -1));
    }
    const eq = findTopLevelEquality(expr);
    if (eq) {
      const lv = resolve(vals, expr.slice(0, eq.index));
      const rv = resolve(vals, expr.slice(eq.index + eq.op.length));
      switch (eq.op) {
        case "===":
          return lv === rv;
        case "!==":
          return lv !== rv;
        case "==":
          return lv == rv;
        default:
          return lv != rv;
      }
    }
    if (expr[0] === "!") return !resolve(vals, expr.slice(1));
    if (expr === "true") return true;
    if (expr === "false") return false;
    if (expr === "null") return null;
    if (expr === "undefined") return void 0;
    if (NUMBER_RE.test(expr)) return Number(expr);
    if (expr.length >= 2 && (expr[0] === '"' || expr[0] === "'") && expr[expr.length - 1] === expr[0]) {
      return expr.slice(1, -1);
    }
    return resolvePath(vals, expr);
  }
  function parensWrapWhole(expr) {
    let depth = 0;
    for (let i = 0; i < expr.length - 1; i++) {
      if (expr[i] === "(") depth++;
      else if (expr[i] === ")") {
        depth--;
        if (depth === 0) return false;
      }
    }
    return true;
  }
  function findTopLevelEquality(expr) {
    let depth = 0;
    for (let i = 0; i < expr.length; i++) {
      const c = expr[i];
      if (c === "[" || c === "(") depth++;
      else if (c === "]" || c === ")") depth--;
      else if (depth === 0 && (c === "=" || c === "!") && expr[i + 1] === "=") {
        if (i > 0 && (expr[i - 1] === "=" || expr[i - 1] === "!")) continue;
        if (!expr.slice(0, i).trim()) continue;
        const op = expr[i + 2] === "=" ? c + "==" : c + "=";
        return { index: i, op };
      }
    }
    return null;
  }
  function resolvePath(vals, expr) {
    const head = expr.match(IDENT_RE);
    if (!head) return void 0;
    let cur = vals == null ? void 0 : vals[head[0]];
    let i = head[0].length;
    while (i < expr.length) {
      if (expr[i] === ".") {
        const m = expr.slice(i + 1).match(IDENT_RE) || expr.slice(i + 1).match(/^\d+/);
        if (!m) return void 0;
        cur = cur == null ? void 0 : cur[m[0]];
        i += 1 + m[0].length;
      } else if (expr[i] === "[") {
        let depth = 1;
        let j = i + 1;
        while (j < expr.length && depth > 0) {
          if (expr[j] === "[") depth++;
          else if (expr[j] === "]") {
            depth--;
            if (depth === 0) break;
          }
          j++;
        }
        if (depth !== 0) return void 0;
        const key = resolve(vals, expr.slice(i + 1, j));
        cur = cur == null ? void 0 : cur[key];
        i = j + 1;
      } else {
        return void 0;
      }
    }
    return cur;
  }

  // src/encode.ts
  var CAMEL_ATTR = "sc-camel-";
  var INLINE_TEXT_TAGS = new Set(
    "a abbr b bdi bdo br cite code del dfn em i ins kbd mark q s samp small span strike strong sub sup u var wbr".split(
      " "
    )
  );
  var RAW_WRAP = {
    select: "sc-raw-select",
    table: "sc-raw-table",
    tbody: "sc-raw-tbody",
    thead: "sc-raw-thead",
    tfoot: "sc-raw-tfoot",
    tr: "sc-raw-tr",
    td: "sc-raw-td",
    th: "sc-raw-th",
    caption: "sc-raw-caption"
  };
  var RAW_UNWRAP = Object.fromEntries(
    Object.entries(RAW_WRAP).map(([k, v]) => [v, k])
  );
  var EVENT_MAP = {
    onclick: "onClick",
    onchange: "onChange",
    oninput: "onInput",
    onsubmit: "onSubmit",
    onkeydown: "onKeyDown",
    onkeyup: "onKeyUp",
    onkeypress: "onKeyPress",
    onmousedown: "onMouseDown",
    onmouseup: "onMouseUp",
    onmouseenter: "onMouseEnter",
    onmouseleave: "onMouseLeave",
    onfocus: "onFocus",
    onblur: "onBlur",
    ondoubleclick: "onDoubleClick",
    oncontextmenu: "onContextMenu",
    onmousemove: "onMouseMove",
    onmouseover: "onMouseOver",
    onmouseout: "onMouseOut",
    onpointerdown: "onPointerDown",
    onpointerup: "onPointerUp",
    onpointermove: "onPointerMove",
    onpointerenter: "onPointerEnter",
    onpointerleave: "onPointerLeave",
    onpointercancel: "onPointerCancel",
    onpointerover: "onPointerOver",
    onpointerout: "onPointerOut",
    ongotpointercapture: "onGotPointerCapture",
    onlostpointercapture: "onLostPointerCapture",
    ontouchstart: "onTouchStart",
    ontouchend: "onTouchEnd",
    ontouchmove: "onTouchMove",
    ontouchcancel: "onTouchCancel",
    ondragstart: "onDragStart",
    ondragend: "onDragEnd",
    ondragenter: "onDragEnter",
    ondragleave: "onDragLeave",
    ondragover: "onDragOver",
    onanimationstart: "onAnimationStart",
    onanimationend: "onAnimationEnd",
    onanimationiteration: "onAnimationIteration",
    ontransitionend: "onTransitionEnd"
  };
  var ATTRS = `(?:[^>"']|"[^"]*"|'[^']*')*`;
  var IMPORT_SELF_CLOSE_RE = new RegExp(
    "<(x-import|dc-import)(" + ATTRS + ")/>",
    "gi"
  );
  var CAMEL_ATTR_RE = /(\s)([a-z]+[A-Z][A-Za-z0-9]*)(\s*=)/g;
  function encodeCamelAttrs(html) {
    return html.replace(
      CAMEL_ATTR_RE,
      (_, sp, name, eq) => sp + CAMEL_ATTR + name.replace(/[A-Z]/g, (c) => "-" + c.toLowerCase()) + eq
    );
  }
  function encodeCase(html) {
    html = html.replace(
      IMPORT_SELF_CLOSE_RE,
      (_, t, a) => "<" + t + a + "></" + t + ">"
    );
    html = html.replace(/<helmet(\s|>)/gi, "<sc-helmet$1");
    html = html.replace(/<\/helmet\s*>/gi, "</sc-helmet>");
    html = encodeCamelAttrs(html);
    for (const [real, alias] of Object.entries(RAW_WRAP)) {
      html = html.replace(
        new RegExp("(</?)" + real + "(?=[\\s>])", "gi"),
        "$1" + alias
      );
    }
    return html;
  }
  function kebabToCamel(s) {
    return s.replace(/-([a-z])/g, (_, c) => c.toUpperCase());
  }
  function cssToObj(css) {
    const o = {};
    for (const decl of css.split(";")) {
      const i = decl.indexOf(":");
      if (i < 0) continue;
      const prop = decl.slice(0, i).trim();
      o[prop.startsWith("--") ? prop : kebabToCamel(prop)] = decl.slice(i + 1).trim();
    }
    return o;
  }
  function compileAttr(raw) {
    const whole = raw.match(/^\s*\{\{([\s\S]+?)\}\}\s*$/);
    if (whole) {
      const path = whole[1];
      return (vals) => resolve(vals, path);
    }
    if (raw.includes("{{")) {
      const parts = raw.split(/\{\{([\s\S]+?)\}\}/g);
      return (vals) => parts.map((s, i) => i & 1 ? resolve(vals, s) ?? "" : s).join("");
    }
    return () => raw;
  }

  // src/compile.ts
  function collectProps(node, kind, host) {
    const propGetters = [];
    const pseudoClasses = [];
    let hintSize = null;
    for (const { name, value } of [...node.attributes]) {
      if (name === "sc-name" || name === "data-dc-tpl") continue;
      let key = name;
      if (key.startsWith(CAMEL_ATTR))
        key = kebabToCamel(key.slice(CAMEL_ATTR.length));
      if (key === "hint-size") {
        hintSize = value;
        continue;
      }
      if (key.startsWith("style-")) {
        pseudoClasses.push(host.pseudoClass(key.slice(6), value));
        continue;
      }
      if (kind !== "dom") {
        if (key.includes("-") && !(kind === "x-import" && (key.startsWith("aria-") || key.startsWith("data-"))))
          key = kebabToCamel(key);
      } else {
        if (key === "class") key = "className";
        else if (key === "for") key = "htmlFor";
        else if (key.startsWith("on"))
          key = EVENT_MAP[key] || "on" + key[2].toUpperCase() + key.slice(3);
      }
      propGetters.push([key, compileAttr(value)]);
    }
    return { propGetters, pseudoClasses, hintSize };
  }
  var HOST_STYLE_PROPS = /* @__PURE__ */ new Set([
    "position",
    "left",
    "right",
    "top",
    "bottom",
    "inset",
    "width",
    "height",
    "z-index",
    "transform"
  ]);
  function hostPositionStyle(style) {
    const all = typeof style === "string" ? cssToObj(style) : style != null && typeof style === "object" ? style : null;
    if (!all) return void 0;
    const out = {};
    for (const [k, v] of Object.entries(all)) {
      const kebab = k.replace(/[A-Z]/g, (c) => "-" + c.toLowerCase());
      if (HOST_STYLE_PROPS.has(kebab)) out[k] = v;
    }
    return Object.keys(out).length ? out : void 0;
  }
  function compileTemplate(html, host) {
    const tpl = document.createElement("template");
    //! nosemgrep: direct-inner-html-assignment
    tpl.innerHTML = encodeCase(html);
    let tplN = 0;
    (function stamp(node) {
      if (node.nodeType === Node.ELEMENT_NODE) {
        node.setAttribute("data-dc-tpl", String(tplN++));
      }
      for (const c of node.childNodes) stamp(c);
    })(tpl.content);
    const builders = walkChildren(tpl.content, host);
    const render = ((vals, ctx) => builders.map((b, i) => b(vals || {}, ctx, i)));
    render.__annotated = tpl.innerHTML;
    return render;
  }
  function walkChildren(node, host) {
    return [...node.childNodes].map((c) => walk(c, host)).filter((b) => b != null);
  }
  var SLIDE_ID_VALUE_RE = /^[0-9a-f]{8}$/;
  var DECK_CONTROL_FLOW_RE = /^(sc-if|sc-for|sc-else|dc-import|x-import)$/;
  var DECK_AUX_RE = /^(template|script|style|sc-helmet|helmet)$/;
  function isDeckMountTag(el) {
    if (el.localName === "deck-stage") return true;
    return el.localName === "x-import" && (el.getAttribute("component-from-global-scope") || "") === "deck-stage";
  }
  function walkDeckChildren(el, host) {
    const pairs = [...el.childNodes].map((c) => ({ c, b: walk(c, host) })).filter((p) => p.b !== null);
    const kids = pairs.map((p) => p.b);
    const seen = /* @__PURE__ */ new Set();
    const wsSeen = /* @__PURE__ */ new Map();
    const keys = [];
    const nextSlideId = new Array(pairs.length);
    {
      let upcoming = null;
      for (let j = pairs.length - 1; j >= 0; j--) {
        const n = pairs[j].c;
        if (n.nodeType === Node.ELEMENT_NODE) {
          const t = n.localName;
          upcoming = !DECK_AUX_RE.test(t) && !DECK_CONTROL_FLOW_RE.test(t) ? n.getAttribute("data-om-slide-id") : null;
        }
        nextSlideId[j] = upcoming;
      }
    }
    for (let j = 0; j < pairs.length; j++) {
      const { c } = pairs[j];
      if (c.nodeType === Node.TEXT_NODE) {
        if ((c.nodeValue ?? "").trim() === "") {
          const base = nextSlideId[j] ? "omid-ws:" + nextSlideId[j] : "omid-ws:aux";
          const n = wsSeen.get(base) ?? 0;
          wsSeen.set(base, n + 1);
          keys.push(n === 0 ? base : base + ":" + n);
          continue;
        }
        return { kids, keys: null };
      }
      if (c.nodeType !== Node.ELEMENT_NODE) {
        keys.push(j);
        continue;
      }
      const child = c;
      const tag = child.localName;
      if (DECK_AUX_RE.test(tag)) {
        keys.push(j);
        continue;
      }
      if (DECK_CONTROL_FLOW_RE.test(tag)) return { kids, keys: null };
      const v = child.getAttribute("data-om-slide-id");
      if (!v || !SLIDE_ID_VALUE_RE.test(v) || seen.has(v)) {
        return { kids, keys: null };
      }
      seen.add(v);
      keys.push("omid:" + v);
    }
    return { kids, keys };
  }
  function renderDeckKids(kids, kidKeys, vals, ctx) {
    return kids.map((b, j) => {
      const k = kidKeys ? kidKeys[j] : j;
      const out = b(vals, ctx, k);
      return kidKeys != null && typeof out === "string" ? h(getReact().Fragment, { key: k }, out) : out;
    });
  }
  function walk(node, host) {
    if (node.nodeType === Node.TEXT_NODE) return walkText(node);
    if (node.nodeType !== Node.ELEMENT_NODE) return null;
    const el = node;
    const tag = el.tagName.toLowerCase();
    if (tag === "sc-for") return walkFor(el, host);
    if (tag === "sc-if") return walkIf(el, host);
    if (tag === "x-import") return walkXImport(el, host);
    if (tag === "sc-helmet") return host.helmet(el);
    if (tag === "dc-import") return walkComponent(el, host);
    return walkElement(el, host);
  }
  var warnedHoles = /* @__PURE__ */ new Set();
  function warnUnresolved(ctx, what) {
    const key = (ctx?.__name || "?") + "\0" + what;
    if (warnedHoles.has(key)) return;
    warnedHoles.add(key);
    console.warn("[dc-runtime] " + (ctx?.__name || "template") + ": " + what);
  }
  function walkText(node) {
    const txt = node.nodeValue ?? "";
    if (!txt.includes("{{")) {
      if (!txt.trim() && !txt.includes(" ")) return null;
      return () => txt;
    }
    const parts = txt.split(/\{\{([\s\S]+?)\}\}/g);
    return (vals, ctx, key) => h(
      getReact().Fragment,
      { key },
      ...parts.map((p, i) => {
        if (!(i & 1)) return p;
        const v = resolve(vals, p);
        if (v === void 0) {
          if (!ctx?.__streamingNow) {
            if (document.body?.hasAttribute("data-dc-editor-on")) {
              return h(
                "span",
                { key: i, className: "sc-interp sc-unresolved" },
                "{{ " + p.trim() + " }}"
              );
            }
            warnUnresolved(
              ctx,
              "{{ " + p.trim() + " }} never resolved \u2014 rendered as empty"
            );
            return null;
          }
          return h(
            "span",
            { key: i, className: "sc-interp sc-missing" },
            p.trim()
          );
        }
        if (getReact().isValidElement(v) || Array.isArray(v)) {
          return h(getReact().Fragment, { key: i }, v);
        }
        if (v === null || typeof v === "boolean") return null;
        return h("span", { key: i, className: "sc-interp" }, String(v));
      })
    );
  }
  function walkFor(el, host) {
    const listGet = compileAttr(el.getAttribute("list") || "");
    const asName = el.getAttribute("as") || "item";
    const hintN = parseInt(el.getAttribute("hint-placeholder-count") || "0", 10);
    const kids = walkChildren(el, host);
    const listSrc = el.getAttribute("list") || "";
    return (vals, ctx, key) => {
      let list = listGet(vals);
      if (!Array.isArray(list)) {
        if (!ctx?.__streamingNow) {
          if (list !== void 0 && list !== null) {
            warnUnresolved(
              ctx,
              'sc-for list="' + listSrc + '" is not an array (' + typeof list + ")"
            );
          }
          list = [];
        } else {
          list = hintN > 0 ? Array(hintN).fill(void 0) : [];
        }
      }
      return h(
        getReact().Fragment,
        { key },
        list.map((item, i) => {
          const sub = { ...vals, [asName]: item, $index: i };
          return h(
            getReact().Fragment,
            { key: i },
            kids.map((b, j) => b(sub, ctx, j))
          );
        })
      );
    };
  }
  function walkIf(el, host) {
    const valGet = compileAttr(el.getAttribute("value") || "");
    const hintRaw = el.getAttribute("hint-placeholder-val");
    const hintGet = hintRaw != null ? compileAttr(hintRaw) : null;
    const kids = walkChildren(el, host);
    return (vals, ctx, key) => {
      let v = valGet(vals);
      if (v === void 0 && hintGet && ctx?.__streamingNow) v = hintGet(vals);
      return v ? h(
        getReact().Fragment,
        { key },
        kids.map((b, j) => b(vals, ctx, j))
      ) : null;
    };
  }
  function walkComponent(el, host) {
    const name = el.getAttribute("name") || el.getAttribute("component") || "";
    el.removeAttribute("name");
    el.removeAttribute("component");
    const tplId = el.getAttribute("data-dc-tpl");
    const styleRaw = el.getAttribute("style");
    el.removeAttribute("style");
    const styleGet = styleRaw != null ? compileAttr(styleRaw) : null;
    const { propGetters, hintSize } = collectProps(el, "dc-import", host);
    const kids = walkChildren(el, host);
    return (vals, ctx, key) => {
      const props = {
        key,
        __hintSize: hintSize,
        __tplId: tplId,
        __hostStyle: styleGet ? hostPositionStyle(styleGet(vals)) : void 0
      };
      for (const [k, g] of propGetters) {
        const v = g(vals);
        if (k === "dcProps") {
          if (v && typeof v === "object") Object.assign(props, v);
          continue;
        }
        props[k] = v;
      }
      if (kids.length) props.children = kids.map((b, j) => b(vals, ctx, j));
      return h(host.component(name), props);
    };
  }
  function walkXImport(el, host) {
    const globalNameGet = compileAttr(
      el.getAttribute("component-from-global-scope") || ""
    );
    const exportNameGet = compileAttr(
      el.getAttribute("component") || el.getAttribute("name") || ""
    );
    const fromRaw = el.getAttribute("from") || (el.getAttribute("component-from-global-scope") ? "" : el.getAttribute("src") || el.getAttribute("import") || "");
    const urls = fromRaw.trim() ? fromRaw.trim().split(/\s+/) : [];
    const url = urls.length ? urls[urls.length - 1] : "";
    const kindOf = (u) => /\.(jsx|tsx)(\?|#|$)/i.test(u) ? "jsx" : "js";
    const tplId = el.getAttribute("data-dc-tpl");
    const styleRaw = el.getAttribute("style");
    el.removeAttribute("style");
    const styleGet = styleRaw != null ? compileAttr(styleRaw) : null;
    const wrap = tplId != null || styleGet != null;
    const { propGetters, hintSize } = collectProps(el, "x-import", host);
    const hasContent = el.children.length > 0 || !!(el.textContent || "").trim();
    const deckKeyed = hasContent && isDeckMountTag(el) ? walkDeckChildren(el, host) : null;
    const kids = deckKeyed ? deckKeyed.kids : hasContent ? walkChildren(el, host) : [];
    const kidKeys = deckKeyed?.keys ?? null;
    const urlBindable = fromRaw.includes("{{");
    if (urls.length && !urlBindable) {
      let prev;
      for (const u of urls) prev = host.loadExternal(kindOf(u), u, prev);
    }
    const evalName = (g, vals) => {
      const v = g(vals);
      const s = v == null ? "" : String(v);
      return s.includes("{{") ? "" : s;
    };
    return (vals, ctx, key) => {
      const globalName = evalName(globalNameGet, vals);
      const name = globalName || evalName(exportNameGet, vals);
      const C = !name || urlBindable ? null : globalName ? host.resolveExternalGlobal(url, globalName) : host.resolveExternal(url, name);
      const hostStyle = styleGet ? hostPositionStyle(styleGet(vals)) : void 0;
      const wrapper = wrap ? {
        key,
        className: "sc-host-x",
        "data-dc-tpl": tplId,
        style: hostStyle || { display: "contents" }
      } : null;
      if (!C) {
        const error = urlBindable ? "x-import `from` cannot contain {{ \u2026 }} \u2014 module URLs are resolved at parse time; use a literal URL" : host.resolveExternalError(url, name);
        const ph = host.placeholder({
          key: wrapper ? void 0 : key,
          name,
          hintSize,
          error
        });
        return wrapper ? h("div", wrapper, ph) : ph;
      }
      const props = wrapper ? {} : { key };
      let unresolvedHole = false;
      for (const [k, g] of propGetters) {
        if (k === "component" || k === "componentFromGlobalScope" || k === "from") {
          continue;
        }
        const v = g(vals);
        if (v === void 0) unresolvedHole = true;
        if (k === "dcProps") {
          if (v && typeof v === "object") Object.assign(props, v);
          continue;
        }
        props[k] = v;
      }
      if (unresolvedHole && ctx?.__htmlStreamingNow) {
        const ph = host.placeholder({
          key: wrapper ? void 0 : key,
          name,
          hintSize,
          error: null
        });
        return wrapper ? h("div", wrapper, ph) : ph;
      }
      if (kids.length) {
        props.children = renderDeckKids(kids, kidKeys, vals, ctx);
      }
      return wrapper ? h("div", wrapper, h(C, props)) : h(C, props);
    };
  }
  function contentKey(el) {
    const clone = el.cloneNode(true);
    for (const d of clone.querySelectorAll("*")) {
      while (d.attributes.length) d.removeAttribute(d.attributes[0].name);
    }
    const s = clone.innerHTML;
    let h2 = 5381;
    for (let i = 0; i < s.length; i++) h2 = (h2 << 5) + h2 + s.charCodeAt(i) | 0;
    return s.length + "." + (h2 >>> 0).toString(36);
  }
  var NEVER_CONTENT_KEYED = new Set(
    "script style textarea option title select canvas iframe video audio".split(
      " "
    )
  );
  var NOT_INLINE_SELECTOR = ":not(" + [...INLINE_TEXT_TAGS].join(",") + ")";
  function walkElement(el, host) {
    const realTag = RAW_UNWRAP[el.localName] || el.localName;
    const tplId = el.getAttribute("data-dc-tpl");
    const inlineOnly = el.childNodes.length > 0 && !NEVER_CONTENT_KEYED.has(realTag) && el.querySelector(NOT_INLINE_SELECTOR) === null;
    const keySuffix = inlineOnly ? "|" + contentKey(el) : "";
    const { propGetters, pseudoClasses } = collectProps(el, "dom", host);
    const deckKeyed = isDeckMountTag(el) ? walkDeckChildren(el, host) : null;
    const kids = deckKeyed ? deckKeyed.kids : walkChildren(el, host);
    const kidKeys = deckKeyed?.keys ?? null;
    return (vals, ctx, key) => {
      const props = {
        key: key + keySuffix,
        "data-dc-tpl": tplId
      };
      for (const [k, g] of propGetters) {
        let v = g(vals);
        if (k === "style" && typeof v === "string") v = cssToObj(v);
        if ((k === "value" || k === "checked") && v === void 0) {
          v = k === "checked" ? false : "";
        }
        props[k] = v;
      }
      if (pseudoClasses.length) {
        props.className = [props.className, ...pseudoClasses].filter(Boolean).join(" ");
      }
      return h(realTag, props, ...renderDeckKids(kids, kidKeys, vals, ctx));
    };
  }

  // src/logic.ts
  var StreamableLogic = class {
    constructor(props) {
      __publicField(this, "props");
      __publicField(this, "state", {});
      /** Back-pointer to the wrapper component, installed after construction. */
      __publicField(this, "__host");
      this.props = props || {};
    }
    setState(update, cb) {
      this.__host && this.__host.__setLogicState(update, cb);
    }
    forceUpdate() {
      this.__host && this.__host.forceUpdate();
    }
    componentDidMount() {
    }
    componentDidUpdate(_prevProps) {
    }
    componentWillUnmount() {
    }
    /** The flat object the template renders against (merged over props). */
    renderVals() {
      return {};
    }
  };
  function evalDcLogic(src) {
    //! nosemgrep: eval-and-function-constructor
    const fn = new Function(
      "DCLogic",
      "StreamableLogic",
      "React",
      src + '\n;return (typeof Component!=="undefined"&&Component)||undefined;'
    );
    return fn(StreamableLogic, StreamableLogic, getReact());
  }

  // src/component.ts
  function shallowEqual(a, b) {
    if (!b) return false;
    const ak = Object.keys(a).filter((k) => k !== "children");
    const bk = Object.keys(b).filter((k) => k !== "children");
    if (ak.length !== bk.length) return false;
    for (const k of ak) if (a[k] !== b[k]) return false;
    return true;
  }
  function Placeholder({
    name,
    hintSize,
    streaming,
    error
  }) {
    const [w, hgt] = (hintSize || "100%,60px").split(",");
    return h(
      "div",
      {
        className: "sc-placeholder" + (streaming ? " sc-streaming" : ""),
        style: { width: w.trim(), height: hgt && hgt.trim() },
        title: name
      },
      error ? h(
        "div",
        { className: "sc-placeholder-error" },
        (name ? name + ": " : "") + error
      ) : null
    );
  }
  function hintToMin(hint) {
    if (!hint) return void 0;
    const [w, hgt] = hint.split(",");
    return { minWidth: w.trim(), minHeight: hgt && hgt.trim() };
  }
  function createComponentFactory(registry, ensureFetched) {
    const React = getReact();
    const AncestorContext = React.createContext([]);
    class StreamableComponent extends React.Component {
      constructor(props) {
        super(props);
        __publicField(this, "__name");
        __publicField(this, "__sub");
        __publicField(this, "__needsDidMount", false);
        /** Snapshot of the registry's streaming flags taken at render time —
         *  builders read it off the RenderCtx (this) to pick placeholder vs
         *  render-nothing for unresolved values. */
        __publicField(this, "__streamingNow", false);
        __publicField(this, "__htmlStreamingNow", false);
        /** When a construct throws, remember the (class, registry.ver, props)
         *  triple so render-time reconcile doesn't re-attempt it on every parent
         *  re-render. A registry bump (new class, template, external module
         *  resolving via bumpAll) changes `ver` and breaks the memo so an
         *  env-dependent constructor can self-heal. */
        __publicField(this, "__failedLogic", null);
        __publicField(this, "__failedUserProps", null);
        __publicField(this, "__failedVer", -1);
        /** Per-instance constructor error — kept here (not on the registry entry)
         *  so one instance's successful construct can't hide a sibling's failure,
         *  and a construct can never wipe an eval error `updateJs` recorded on
         *  `r.logicError`. */
        __publicField(this, "__ctorError", null);
        __publicField(this, "logic");
        this.__name = props.__name;
        this.state = { __v: 0, __err: null };
        this.__sub = () => {
          if (this.state.__err) this.setState({ __err: null });
          this.forceUpdate();
        };
        this.__makeLogic(registry.get(this.__name).Logic, null);
        ensureFetched(this.__name);
      }
      /** Error-boundary hook: a render crash anywhere in this DC's subtree
       *  (its own template, an x-import'd component, a child DC without its
       *  own deeper boundary) lands here instead of unmounting the page. */
      static getDerivedStateFromError(e) {
        return { __err: e instanceof Error && e.message ? e.message : String(e) };
      }
      componentDidCatch(e, info) {
        console.error(
          "[dc-runtime] render error in <" + this.__name + ">:",
          e,
          info?.componentStack || ""
        );
      }
      /** Instantiate the logic class (or the no-op base) and adopt `prevState`
       *  over its initial state — used both at mount and on hot-swap. */
      __makeLogic(Logic, prevState) {
        const L = Logic || StreamableLogic;
        try {
          this.logic = new L(this.__userProps());
          this.__failedLogic = null;
          this.__failedUserProps = null;
          this.__ctorError = null;
        } catch (e) {
          console.error(e);
          this.__failedLogic = Logic;
          this.__failedUserProps = this.__userProps();
          this.__failedVer = registry.get(this.__name).ver;
          this.__ctorError = this.__name + ": " + (e instanceof Error && e.message ? e.message : String(e));
          this.logic = new StreamableLogic(
            this.__userProps()
          );
        }
        this.logic.__host = this;
        if (prevState)
          this.logic.state = { ...this.logic.state || {}, ...prevState };
      }
      /** The props the author's logic + template see — internal __-prefixed
       *  wiring stripped. */
      __userProps() {
        const { __name, __hintSize, __tplId, __hostStyle, ...rest } = this.props;
        return rest;
      }
      __setLogicState(update, cb) {
        const prev = this.logic.state;
        const patch = typeof update === "function" ? update(prev) : update;
        this.logic.state = { ...prev, ...patch };
        this.setState((s) => ({ __v: s.__v + 1 }), cb);
      }
      /** Swap the logic instance when the registry's Logic class changed
       *  (streaming completion, hot reload). State carries over; didMount
       *  re-fires after the swap commits so refs exist. */
      __reconcileLogic() {
        const r = registry.get(this.__name);
        const Next = r.Logic;
        const Cur = this.logic.constructor;
        if (Next === Cur || !Next && Cur === StreamableLogic || Next === this.__failedLogic && r.ver === this.__failedVer && shallowEqual(this.__userProps(), this.__failedUserProps)) {
          return;
        }
        if (!this.__needsDidMount) {
          try {
            this.logic.componentWillUnmount();
          } catch (e) {
            console.error(e);
          }
        }
        this.__makeLogic(Next, this.logic.state);
        this.__needsDidMount = true;
      }
      componentDidMount() {
        registry.get(this.__name).subs.add(this.__sub);
        try {
          this.logic.componentDidMount();
        } catch (e) {
          console.error(e);
        }
      }
      componentDidUpdate(prevProps) {
        this.logic.props = this.__userProps();
        if (this.__needsDidMount) {
          if (this.state.__err || !registry.get(this.__name).tpl) return;
          this.__needsDidMount = false;
          try {
            this.logic.componentDidMount();
          } catch (e) {
            console.error(e);
          }
        } else {
          try {
            this.logic.componentDidUpdate(prevProps);
          } catch (e) {
            console.error(e);
          }
        }
      }
      componentWillUnmount() {
        registry.get(this.__name).subs.delete(this.__sub);
        if (!this.__needsDidMount) {
          try {
            this.logic.componentWillUnmount();
          } catch (e) {
            console.error(e);
          }
        }
      }
      render() {
        const r = registry.get(this.__name);
        const cls = "sc-host" + (r.htmlStreaming ? " sc-streaming-html" : "") + (r.jsStreaming ? " sc-streaming-js" : "");
        const hintStyle = r.htmlStreaming ? hintToMin(this.props.__hintSize) : void 0;
        const hostStyle = this.props.__hostStyle || hintStyle ? { ...hintStyle || {}, ...this.props.__hostStyle || {} } : void 0;
        const hostBase = {
          className: cls,
          style: hostStyle,
          "data-sc-name": this.__name,
          "data-dc-tpl": this.props.__tplId
        };
        const chain = Array.isArray(this.context) ? this.context : [];
        if (chain.includes(this.__name)) {
          const cycle = [
            ...chain.slice(chain.indexOf(this.__name)),
            this.__name
          ].join(" \u2192 ");
          return h(
            "div",
            { ...hostBase, className: cls + " sc-has-error" },
            h(Placeholder, {
              name: this.__name,
              hintSize: this.props.__hintSize,
              error: "circular import: " + cycle
            })
          );
        }
        if (this.state.__err) {
          return h(
            "div",
            { ...hostBase, className: cls + " sc-has-error" },
            h(
              "div",
              { className: "sc-logic-error", "data-omelette-chrome": "" },
              this.__name + ": " + this.state.__err
            ),
            h(Placeholder, {
              name: this.__name,
              hintSize: this.props.__hintSize,
              error: this.state.__err
            })
          );
        }
        this.__reconcileLogic();
        if (!r.tpl) {
          return h(
            "div",
            hostBase,
            h(Placeholder, { name: this.__name, hintSize: this.props.__hintSize })
          );
        }
        const userProps = this.__userProps();
        this.logic.props = userProps;
        let vals = userProps;
        let renderErr = r.logicError || this.__ctorError;
        try {
          vals = { ...userProps, ...this.logic.renderVals() || {} };
        } catch (e) {
          console.error(e);
          renderErr = this.__name + ".renderVals(): " + (e instanceof Error && e.message ? e.message : String(e));
        }
        this.__streamingNow = !!(r.htmlStreaming || r.jsStreaming);
        this.__htmlStreamingNow = !!r.htmlStreaming;
        return h(
          "div",
          { ...hostBase, className: cls + (renderErr ? " sc-has-error" : "") },
          renderErr && h(
            "div",
            { className: "sc-logic-error", "data-omelette-chrome": "" },
            renderErr
          ),
          h(
            AncestorContext.Provider,
            { value: [...chain, this.__name] },
            r.tpl(vals, this)
          )
        );
      }
    }
    __publicField(StreamableComponent, "contextType", AncestorContext);
    const named = /* @__PURE__ */ new Map();
    function getDC(name) {
      const hit = named.get(name);
      if (hit) return hit;
      function Dispatcher(p) {
        const [, setTick] = React.useState(0);
        React.useEffect(() => {
          const sub = () => setTick((n) => n + 1);
          registry.get(name).subs.add(sub);
          return () => {
            registry.get(name).subs.delete(sub);
          };
        }, []);
        ensureFetched(name);
        return h(StreamableComponent, { ...p, __name: name });
      }
      Dispatcher.displayName = name;
      named.set(name, Dispatcher);
      return Dispatcher;
    }
    return {
      getDC,
      StreamableComponent
    };
  }

  // src/bundled.ts
  function bundledBlob(url) {
    const blobs = window.__resourceBlobs;
    const b = blobs ? blobs[url.split("#")[0]] : void 0;
    return b instanceof Blob ? b : null;
  }

  // src/cdn.ts
  var REACT_URL = "https://unpkg.com/react@18.3.1/umd/react.production.min.js";
  var REACT_SRI = "sha384-DGyLxAyjq0f9SPpVevD6IgztCFlnMF6oW/XQGmfe+IsZ8TqEiDrcHkMLKI6fiB/Z";
  var REACT_DOM_URL = "https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js";
  var REACT_DOM_SRI = "sha384-gTGxhz21lVGYNMcdJOyq01Edg0jhn/c22nsx0kyqP0TxaV5WVdsSH1fSDUf5YJj1";
  var BABEL_URL = "https://unpkg.com/@babel/standalone@7.29.0/babel.min.js";
  var BABEL_SRI = "sha384-m08KidiNqLdpJqLq95G/LEi8Qvjl/xUYll3QILypMoQ65QorJ9Lvtp2RXYGBFj1y";
  function cdnScriptFor(url, sri) {
    const res = window.__resources;
    const v = res ? res[url] : void 0;
    return typeof v === "string" && v ? { src: v } : { src: url, integrity: sri };
  }

  // src/external.ts
  var isCustomElementName = (n) => !n.includes(".") && n.includes("-");
  function isRenderableType(g) {
    if (typeof g === "function") return !isElementClass(g);
    return typeof g === "object" && g !== null && typeof g.$$typeof === "symbol";
  }
  function resolveDottedPath(root, name) {
    let cur = root;
    for (const seg of name.split(".")) {
      if (cur == null) return void 0;
      cur = cur[seg];
    }
    return cur;
  }
  var GLOBAL_POLL_INTERVAL_MS = 50;
  var GLOBAL_POLL_TIMEOUT_MS = 3e4;
  function createExternalModules(onResolved) {
    const cache = /* @__PURE__ */ new Map();
    let babelLoading = null;
    const reportedMissing = /* @__PURE__ */ new Map();
    const polling = /* @__PURE__ */ new Set();
    function ensureBabel() {
      if (window.Babel) return Promise.resolve();
      if (babelLoading) return babelLoading;
      const babel = cdnScriptFor(BABEL_URL, BABEL_SRI);
      babelLoading = new Promise((res, rej) => {
        const s = document.createElement("script");
        s.src = babel.src;
        if (babel.integrity) {
          s.integrity = babel.integrity;
          s.crossOrigin = "anonymous";
        }
        s.onload = () => res();
        s.onerror = rej;
        document.head.appendChild(s);
      });
      return babelLoading;
    }
    const pending = /* @__PURE__ */ new Map();
    function load(kind, url, after) {
      const existing = pending.get(url);
      if (existing) return existing;
      cache.set(url, null);
      console.info("[dc-runtime] x-import: loading", url, "(" + kind + ")");
      const ready = Promise.all([
        kind === "jsx" ? ensureBabel() : Promise.resolve(),
        after ?? Promise.resolve()
      ]);
      const p = ready.then(() => {
        const pre = bundledBlob(url);
        if (pre) return pre.text();
        return fetch(url).then((r) => {
          if (!r.ok) throw new Error("HTTP " + r.status);
          return r.text();
        });
      }).then((src) => {
        const code = kind === "jsx" ? window.Babel.transform(src, {
          filename: url,
          presets: ["react", "typescript"]
        }).code : src;
        const module = { exports: {} };
        const before = new Set(Object.keys(window));
        //! nosemgrep: eval-and-function-constructor
        new Function("React", "module", "exports", "require", code)(
          getReact(),
          module,
          module.exports,
          () => ({})
        );
        const globals = {};
        for (const k of Object.keys(window)) {
          if (!before.has(k) && typeof window[k] === "function") {
            globals[k] = window[k];
          }
        }
        cache.set(url, { mod: module.exports, globals });
        console.info(
          "[dc-runtime] x-import: loaded",
          url,
          "\u2014 exports:",
          Object.keys(module.exports),
          "window globals:",
          Object.keys(globals)
        );
        onResolved();
      }).catch((e) => {
        cache.set(url, {
          mod: {},
          globals: {},
          error: "failed to load: " + (e instanceof Error && e.message ? e.message : String(e))
        });
        console.error(
          "[dc-runtime] x-import: FAILED to load",
          url,
          "(" + kind + ")",
          e
        );
        onResolved();
      });
      pending.set(url, p);
      return p;
    }
    function resolve2(url, name) {
      const entry = cache.get(url);
      if (!entry) return null;
      const { mod, globals } = entry;
      const C = mod && mod[name] || globals && globals[name] || typeof window !== "undefined" && window[name] || mod && mod.default;
      if (typeof C === "function") return C;
      const key = url + "\0" + name;
      if (!reportedMissing.has(key)) {
        reportedMissing.set(
          key,
          entry.error || 'no export named "' + name + '" (has: ' + Object.keys(mod).join(", ") + ")"
        );
        console.error(
          "[dc-runtime] x-import: module",
          url,
          "loaded but has no component named",
          JSON.stringify(name),
          "\u2014 available exports:",
          Object.keys(mod),
          "window globals:",
          Object.keys(globals),
          ". The module must `module.exports = {" + name + "}` or set `window." + name + "`."
        );
      }
      return null;
    }
    function waitForGlobal(name) {
      if (polling.has(name)) return;
      polling.add(name);
      const started = Date.now();
      const isCE = isCustomElementName(name);
      const tick = () => {
        const found = isCE ? customElements.get(name) : isRenderableType(resolveDottedPath(window, name));
        if (found) {
          polling.delete(name);
          onResolved();
          return;
        }
        if (Date.now() - started >= GLOBAL_POLL_TIMEOUT_MS) {
          console.warn(
            "[dc-runtime] x-import: global",
            JSON.stringify(name),
            "never appeared on window after " + GLOBAL_POLL_TIMEOUT_MS + "ms"
          );
          return;
        }
        setTimeout(tick, GLOBAL_POLL_INTERVAL_MS);
      };
      setTimeout(tick, GLOBAL_POLL_INTERVAL_MS);
    }
    function resolveGlobal(url, name) {
      const isCE = isCustomElementName(name);
      if (!url) {
        if (isCE) {
          if (customElements.get(name)) return name;
          waitForGlobal(name);
          return null;
        }
        const g2 = resolveDottedPath(window, name);
        if (isRenderableType(g2)) return g2;
        waitForGlobal(name);
        return null;
      }
      const entry = cache.get(url);
      if (!entry) return null;
      if (isCE && customElements.get(name)) return name;
      const g = entry.globals[name] ?? resolveDottedPath(window, name);
      if (isRenderableType(g)) return g;
      if (name.includes(".")) return null;
      const key = url + "\0global\0" + name;
      if (!reportedMissing.has(key)) {
        reportedMissing.set(key, null);
        if (isCE && !customElements.get(name)) {
          console.warn(
            "[dc-runtime] x-import:",
            url,
            "loaded but no custom element",
            JSON.stringify(name),
            "is registered and window." + name + " is not a function \u2014 rendering <" + name + "> as an unknown element."
          );
        }
      }
      return name;
    }
    function getError(url, name) {
      const entry = cache.get(url);
      if (entry?.error) return entry.error;
      return reportedMissing.get(url + "\0" + name) || null;
    }
    return { load, resolve: resolve2, resolveGlobal, getError };
  }
  function isElementClass(g) {
    try {
      return typeof g === "function" && typeof HTMLElement !== "undefined" && g.prototype instanceof HTMLElement;
    } catch {
      return false;
    }
  }

  // src/atomics.ts
  var ATOMIC_CSS = (
    // layout
    ".fx{display:flex}.col{display:flex;flex-direction:column}.grid{display:grid}.ac{align-items:center}.jc{justify-content:center}.jb{justify-content:space-between}.f1{flex:1}.noshrink{flex-shrink:0}.wrap{flex-wrap:wrap}.fw5{font-weight:500}.fw6{font-weight:600}.fw7{font-weight:700}.fw8{font-weight:800}.fs11{font-size:11px}.fs12{font-size:12px}.fs13{font-size:13px}.fs14{font-size:14px}.fs15{font-size:15px}.fs16{font-size:16px}.fs20{font-size:20px}.fs22{font-size:22px}.upper{text-transform:uppercase}.tc{text-align:center}.nowrap{white-space:nowrap}.gap8{gap:8px}.gap10{gap:10px}.gap12{gap:12px}.gap16{gap:16px}.gap24{gap:24px}.m0{margin:0}.mt8{margin-top:8px}.mt12{margin-top:12px}.mt16{margin-top:16px}.mb8{margin-bottom:8px}.mb12{margin-bottom:12px}.mb16{margin-bottom:16px}.posrel{position:relative}.posabs{position:absolute}.round{border-radius:50%}.ohide{overflow:hidden}.bbox{box-sizing:border-box}.pointer{cursor:pointer}.w100{width:100%}.b0{border:none}"
  );

  // src/helmet.ts
  var DESIGN_DOC_MODE_RE = /<meta\b[^>]*\bname\s*=\s*["']design_doc_mode["'][^>]*\b(?:content|value)\s*=\s*["'](\w+)["']/i;
  var CANVAS_BG_LIGHT = "#f0eee6";
  var CANVAS_BG_DARK = "#2e2c26";
  function createHelmetManager(doc, isStreaming) {
    const mounted = /* @__PURE__ */ new Set();
    const live = /* @__PURE__ */ new Map();
    let designDocMode = null;
    let canvasStyleEl = null;
    let appTheme = "light";
    try {
      const ds = doc.documentElement.dataset.theme;
      appTheme = ds === "dark" || ds === "light" ? ds : new URLSearchParams(doc.defaultView?.location.search ?? "").get(
        "theme"
      ) === "dark" ? "dark" : "light";
    } catch {
    }
    function applyCanvasBg() {
      if (!canvasStyleEl) return;
      const bg = appTheme === "dark" ? CANVAS_BG_DARK : CANVAS_BG_LIGHT;
      canvasStyleEl.textContent = `html,body{background:${bg}}#dc-root>.sc-host{position:relative}`;
    }
    function postDesignMode(mode) {
      if (window.parent === window) return;
      try {
        window.parent.postMessage({ type: "__dc_design_mode", mode }, "*");
      } catch {
      }
    }
    function setDesignDocMode(mode) {
      if (mode === designDocMode) return;
      designDocMode = mode;
      postDesignMode(mode);
      if (mode === "canvas") {
        doc.documentElement.setAttribute("data-dc-canvas", "");
        canvasStyleEl = doc.createElement("style");
        canvasStyleEl.setAttribute("data-dc-canvas", "");
        applyCanvasBg();
        doc.head.appendChild(canvasStyleEl);
      } else {
        doc.documentElement.removeAttribute("data-dc-canvas");
        canvasStyleEl?.remove();
        canvasStyleEl = null;
      }
    }
    window.addEventListener("message", (e) => {
      const type = e.data && e.data.type;
      if (type === "__dc_theme") {
        const t = e.data.theme;
        if (t === "light" || t === "dark") {
          appTheme = t;
          applyCanvasBg();
        }
        return;
      }
      if (!designDocMode || type !== "__dc_probe") return;
      postDesignMode(designDocMode);
    });
    function compile(node) {
      const raw = [...node.children];
      const helmetClosed = node.nextSibling != null || node.parentNode?.nextSibling != null;
      if (node.hasAttribute("data-dc-atomics") && !mounted.has("__dc-atomics")) {
        mounted.add("__dc-atomics");
        const el = doc.createElement("style");
        el.id = "__dc-atomics";
        el.textContent = ATOMIC_CSS;
        doc.head.appendChild(el);
      }
      return (_vals, ctx) => {
        const name = ctx && ctx.__name || "";
        const streaming = !!(name && isStreaming(name));
        for (let i = 0; i < raw.length; i++) {
          const child = raw[i];
          const tag = child.tagName;
          const mayBePartial = streaming && !helmetClosed && i === raw.length - 1;
          if (tag === "SCRIPT") {
            if (mayBePartial) continue;
            const key = "SCRIPT|" + (child.getAttribute("src") || child.textContent || "");
            if (mounted.has(key)) continue;
            mounted.add(key);
            const el = doc.createElement("script");
            for (const { name: an, value } of [...child.attributes])
              el.setAttribute(an, value);
            if (child.textContent) el.textContent = child.textContent;
            doc.head.appendChild(el);
          } else if (tag === "LINK" || tag === "META") {
            if (mayBePartial) continue;
            const key = tag + "|" + (child.getAttribute("href") || child.getAttribute("src") || child.outerHTML);
            if (mounted.has(key)) continue;
            mounted.add(key);
            if (tag === "LINK") {
              const rel = (child.getAttribute("rel") || "").toLowerCase().split(/\s+/);
              const href = (child.getAttribute("href") || "").trim();
              const res = window.__resources;
              const pre = res && rel.includes("stylesheet") && !rel.includes("alternate") ? res[href] : void 0;
              const blob = typeof pre === "string" && pre ? bundledBlob(pre) : null;
              if (blob) {
                const el = doc.createElement("style");
                if (child.hasAttribute("disabled")) {
                  el.setAttribute("media", "not all");
                } else if (child.getAttribute("media")) {
                  el.setAttribute("media", child.getAttribute("media"));
                }
                if (child.getAttribute("title"))
                  el.setAttribute("title", child.getAttribute("title"));
                void blob.text().then((css) => {
                  el.textContent = css;
                });
                doc.head.appendChild(el);
                continue;
              }
            }
            doc.head.appendChild(child.cloneNode(true));
          } else {
            const key = name + "|" + i;
            let el = live.get(key);
            if (!el || el.tagName !== tag) {
              if (el) el.remove();
              el = doc.createElement(tag.toLowerCase());
              live.set(key, el);
              doc.head.appendChild(el);
            }
            for (const { name: an, value } of [...child.attributes]) {
              if (el.getAttribute(an) !== value) el.setAttribute(an, value);
            }
            if (el.textContent !== child.textContent)
              el.textContent = child.textContent;
          }
        }
        return null;
      };
    }
    return { compile, setDesignDocMode };
  }

  // src/pseudo.ts
  function scanUnquotedUrl(css, i) {
    if (css[i] !== "u" && css[i] !== "U" || css.slice(i, i + 4).toLowerCase() !== "url(" || /[a-z0-9_-]/i.test(css[i - 1] ?? "")) {
      return -1;
    }
    let j = i + 4;
    while (j < css.length && /\s/.test(css[j])) j++;
    if (css[j] === '"' || css[j] === "'") return -1;
    while (j < css.length && css[j] !== ")") {
      if (css[j] === "\\") j++;
      j++;
    }
    return j < css.length ? j + 1 : css.length;
  }
  function stripComments(css) {
    let out = "";
    let quote = "";
    for (let i = 0; i < css.length; i++) {
      const c = css[i];
      if (quote) {
        if (c === "\\") {
          out += c + (css[i + 1] ?? "");
          i++;
          continue;
        }
        if (c === quote) quote = "";
        out += c;
      } else if (c === "'" || c === '"') {
        quote = c;
        out += c;
      } else if (c === "/" && css[i + 1] === "*") {
        const end = css.indexOf("*/", i + 2);
        i = end === -1 ? css.length : end + 1;
        out += " ";
      } else {
        const end = scanUnquotedUrl(css, i);
        if (end === -1) out += c;
        else {
          out += css.slice(i, end);
          i = end - 1;
        }
      }
    }
    return out;
  }
  function importantify(css) {
    css = stripComments(css);
    const decls = [];
    let start = 0;
    let depth = 0;
    let quote = "";
    for (let i = 0; i < css.length; i++) {
      const c = css[i];
      if (quote) {
        if (c === "\\") i++;
        else if (c === quote) quote = "";
      } else if (c === "'" || c === '"') quote = c;
      else if (c === "(") depth++;
      else if (c === ")") depth = Math.max(0, depth - 1);
      else if (c === ";" && depth === 0) {
        decls.push(css.slice(start, i));
        start = i + 1;
      } else {
        const end = scanUnquotedUrl(css, i);
        if (end !== -1) i = end - 1;
      }
    }
    decls.push(css.slice(start));
    return decls.map((d) => d.trim()).filter(Boolean).map((d) => /!\s*important$/i.test(d) ? d : d + " !important").join(";");
  }
  function createPseudoSheet(doc) {
    let el = null;
    const cache = /* @__PURE__ */ new Map();
    let n = 0;
    return (pseudo, css) => {
      const k = pseudo + "|" + css;
      const hit = cache.get(k);
      if (hit) return hit;
      if (!el) {
        el = doc.createElement("style");
        doc.head.appendChild(el);
      }
      const cls = "scp" + (n++).toString(36);
      const isPseudoElement = pseudo === "before" || pseudo === "after";
      const sel = isPseudoElement ? "." + cls + "::" + pseudo : "." + cls + ":" + pseudo;
      el.sheet.insertRule(
        sel + "{" + (isPseudoElement ? css : importantify(css)) + "}",
        el.sheet.cssRules.length
      );
      cache.set(k, cls);
      return cls;
    };
  }

  // src/registry.ts
  function createRegistry() {
    const entries = /* @__PURE__ */ Object.create(null);
    function get(name) {
      return entries[name] || (entries[name] = {
        html: "",
        tpl: null,
        Logic: null,
        jsStreaming: false,
        htmlStreaming: false,
        ver: 0,
        subs: /* @__PURE__ */ new Set(),
        fetched: false
      });
    }
    function bump(name) {
      const r = get(name);
      r.ver++;
      for (const fn of r.subs) fn();
    }
    return {
      entries,
      get,
      bump,
      bumpAll() {
        for (const n in entries) bump(n);
      }
    };
  }

  // src/runtime.ts
  var COMPONENT_DIR = ".";
  function createRuntime(doc = document) {
    const registry = createRegistry();
    const pseudoClass = createPseudoSheet(doc);
    const helmet = createHelmetManager(
      doc,
      (name) => registry.get(name).htmlStreaming
    );
    const external = createExternalModules(() => registry.bumpAll());
    const factory = createComponentFactory(registry, ensureFetched);
    const host = {
      component: (name) => factory.getDC(name),
      placeholder: (props) => h(Placeholder, props),
      helmet: (node) => helmet.compile(node),
      loadExternal: (kind, url, after) => external.load(kind, url, after),
      resolveExternal: (url, name) => external.resolve(url, name),
      resolveExternalGlobal: (url, name) => external.resolveGlobal(url, name),
      resolveExternalError: (url, name) => external.getError(url, name),
      pseudoClass
    };
    function ensureFetched(name) {
      const r = registry.get(name);
      if (r.fetched) return;
      r.fetched = true;
      const url = COMPONENT_DIR + "/" + encodeURIComponent(name) + ".dc.html";
      const res = window.__resources;
      const pre = res ? res[url] : void 0;
      const target = typeof pre === "string" && pre ? pre : url;
      const blob = bundledBlob(target);
      (blob ? blob.text() : fetch(target).then((res2) => {
        if (!res2.ok) {
          console.error(
            '[dc-runtime] sibling fetch for "' + name + '" failed:',
            url,
            "returned",
            res2.status,
            "\u2014 the reference renders as an empty placeholder."
          );
          return "";
        }
        return res2.text();
      })).then((t) => {
        if (!t) return;
        const parsed = parseDcText(t);
        if (!parsed) {
          console.error(
            '[dc-runtime] sibling fetch for "' + name + '":',
            url,
            "has no <x-dc> block \u2014 not a Design Component."
          );
          return;
        }
        if (parsed.props) r.propsMeta = parsed.props;
        if (parsed.preview) r.preview = parsed.preview;
        if (parsed.template && !r.html) updateHtml(name, parsed.template);
        if (parsed.js && !r.Logic) updateJs(name, parsed.js);
      }).catch(
        (e) => console.error(
          '[dc-runtime] sibling fetch for "' + name + '" threw:',
          url,
          e
        )
      );
    }
    let rootName = null;
    function updateHtml(name, html) {
      const r = registry.get(name);
      r.html = html;
      if (name === rootName) {
        const mode = DESIGN_DOC_MODE_RE.exec(html)?.[1] ?? null;
        if (mode || !r.htmlStreaming) helmet.setDesignDocMode(mode);
      }
      try {
        r.tpl = compileTemplate(html, host);
      } catch (e) {
        console.error("[dc-runtime] template compile FAILED for", name, e);
      }
      registry.bump(name);
    }
    function updateJs(name, src) {
      const r = registry.get(name);
      const seq = r.jsSeq = (r.jsSeq || 0) + 1;
      try {
        const Cls = evalDcLogic(src);
        if (r.jsSeq !== seq) return;
        if (typeof Cls !== "function") {
          r.logicError = name + ".dc.html: <script data-dc-script> must define `class Component extends DCLogic`";
        } else {
          r.logicError = null;
          r.Logic = Cls;
        }
      } catch (e) {
        if (r.jsSeq !== seq) return;
        console.error(
          "[dc-runtime] logic class eval FAILED for",
          name,
          "\u2014 the template renders with props only.",
          e
        );
        r.logicError = name + ": " + (e instanceof Error && e.message ? e.message : String(e));
      }
      registry.bump(name);
    }
    function setStreaming(name, kind, on) {
      const r = registry.get(name);
      if (kind === "html") r.htmlStreaming = !!on;
      else r.jsStreaming = !!on;
      let any = false;
      for (const n in registry.entries) {
        const e = registry.entries[n];
        if (e && (e.htmlStreaming || e.jsStreaming)) {
          any = true;
          break;
        }
      }
      doc.documentElement.classList.toggle("sc-dc-streaming", any);
      registry.bump(name);
    }
    function dcUpdate(name, kind, content, streaming) {
      if (streaming) registry.get(name).fetched = true;
      if (kind === "html") {
        setStreaming(name, "html", !!streaming);
        updateHtml(name, content);
      } else if (kind === "js") {
        setStreaming(name, "js", !!streaming);
        if (!streaming) updateJs(name, content);
      } else if (kind === "props") {
        const { props, preview } = parseDataProps(content);
        const r = registry.get(name);
        r.propsMeta = props ?? void 0;
        r.preview = preview;
        registry.bump(name);
      }
    }
    function setProps(name, overrides) {
      registry.get(name).propOverrides = overrides && typeof overrides === "object" ? { ...overrides } : null;
      registry.bump(name);
    }
    function adoptParsed(name, parsed) {
      if (!parsed) return;
      const r = registry.get(name);
      if (parsed.props) r.propsMeta = parsed.props;
      if (parsed.preview) r.preview = parsed.preview;
      if (parsed.template) updateHtml(name, parsed.template);
      if (parsed.js) updateJs(name, parsed.js);
    }
    return {
      registry,
      getDC: factory.getDC,
      updateHtml,
      updateJs,
      dcUpdate,
      setProps,
      adoptParsed,
      setRootName: (name) => {
        rootName = name;
      },
      markFetched: (name) => {
        registry.get(name).fetched = true;
      },
      annotatedTemplate: (name) => {
        const r = registry.get(name);
        return r.tpl && r.tpl.__annotated || null;
      },
      templateSource: (name) => registry.get(name).html || null,
      StreamableLogic
    };
  }

  // src/stream-state.ts
  function createStreamTracker(staleMs = 6e4, now = Date.now) {
    const since = /* @__PURE__ */ new Map();
    const liveOne = (n) => {
      const t = since.get(n);
      if (t === void 0) return false;
      if (now() - t > staleMs) {
        since.delete(n);
        return false;
      }
      return true;
    };
    return {
      push(name, streaming, viewportKey) {
        if (viewportKey === "dc-model") return;
        if (streaming) since.set(name, now());
        else since.delete(name);
      },
      live(name) {
        if (name !== void 0) return liveOne(name);
        for (const n of [...since.keys()]) if (liveOne(n)) return true;
        return false;
      }
    };
  }

  // src/index.ts
  function hideRawTemplate() {
    const s = document.createElement("style");
    s.textContent = "x-dc{display:none!important}";
    document.head.appendChild(s);
  }
  function loadScript(src, integrity) {
    return new Promise((resolve2, reject) => {
      //! nosemgrep: create-script-element
      const s = document.createElement("script");
      s.src = src;
      if (integrity) {
        s.integrity = integrity;
        s.crossOrigin = "anonymous";
      }
      s.async = false;
      s.onload = () => resolve2();
      s.onerror = () => reject(new Error(`failed to load ${src}`));
      document.head.appendChild(s);
    });
  }
  function loadReactUmd() {
    const w = window;
    if (w.React && w.ReactDOM) return Promise.resolve();
    const react = cdnScriptFor(REACT_URL, REACT_SRI);
    const reactDom = cdnScriptFor(REACT_DOM_URL, REACT_DOM_SRI);
    return Promise.all([
      loadScript(react.src, react.integrity),
      loadScript(reactDom.src, reactDom.integrity)
    ]).then(() => void 0);
  }
  function init() {
    const runtime = createRuntime(document);
    let rootName = "Root";
    const baseCss = document.createElement("style");
    baseCss.textContent = BASE_CSS;
    document.head.prepend(baseCss);
    const notifyHost = () => {
      if (window.parent === window) return;
      const r = runtime.registry.entries[rootName];
      try {
        window.parent.postMessage(
          {
            type: "__dc_booted",
            rootName,
            propsMeta: r && r.propsMeta || null,
            preview: r && r.preview || null
          },
          "*"
        );
      } catch {
      }
    };
    const streams = createStreamTracker();
    const api = {
      __dcUpdate: (name, kind, content, streaming, viewportKey) => {
        streams.push(name, streaming, viewportKey);
        runtime.dcUpdate(name, kind, content, streaming);
        if (name === rootName && !streaming && kind === "props") notifyHost();
      },
      __dcStreaming: (name) => streams.live(name),
      __dcSetProps: (name, overrides) => runtime.setProps(name, overrides),
      /** Name of the component currently mounted as the page root — DC tools
       *  push their template-stream here when targeting "the open page". */
      __dcRootName: () => rootName,
      /** Editor bridge — the encoded, `data-dc-tpl`-annotated template source.
       *  The host editor parses this into its own template DOM so it can map a
       *  rendered node (carrying the same `data-dc-tpl`) back to the source
       *  node that emitted it. Returns the encoded form (`sc-camel-*` attrs,
       *  `<sc-raw-*>`/`<sc-helmet>` tags); the editor decodes on serialize. */
      __dcAnnotatedTemplate: (name) => runtime.annotatedTemplate(name),
      /** Editor bridge — the *original* (decoded) template source. */
      __dcTemplateSource: (name) => runtime.templateSource(name),
      __dcBoot: () => {
        rootName = boot(runtime, document) ?? rootName;
        notifyHost();
      },
      __dcRegistry: runtime.registry.entries,
      getDC: (name) => runtime.getDC(name),
      // `DCLogic` is the documented base class name; `StreamableLogic` is the
      // implementation alias kept for any project that already references it.
      DCLogic: runtime.StreamableLogic,
      StreamableLogic: runtime.StreamableLogic
    };
    Object.assign(window, api);
    window.__dcContentKeyed = true;
    if (document.readyState !== "loading") api.__dcBoot();
    else document.addEventListener("DOMContentLoaded", () => api.__dcBoot());
  }
  hideRawTemplate();
  loadReactUmd().then(init).catch((err) => {
    console.error("[dc] failed to load React or boot:", err);
    throw err;
  });
})();

// ===================== iw-core.js (merged) =====================
  // ---- iw-core.js (merged; originally a separate ES module) ----
  window.IWCore = (function(){
  /* iw-core.js — pure logic for the IW28/IW38 dashboard.
     xlsx parsing + all derived metrics. No DOM, no rendering. */
  
  const INK = "#201e1d";
  const ACC = "#ec3013";
  const ACC7 = "#ae1800";
  const ACC1 = "#fff2ef";
  const N200 = "#eae7e7";
  const N300 = "#d7d3d3";
  const N400 = "#bab6b6";
  const N500 = "#9b9797";
  const N600 = "#7d7979";
  const N700 = "#605d5d";
  const N800 = "#444141";
  const OK = "#1f7a4d";
  const OK_BG = "#e6f4ea";
  
  const DEFAULT_SETTINGS = { overdueDays: 30, monthsBack: 18, monthlyTarget: 0, monthStartCount: null, monthStartKey: "" };
  
  const PRIORITIES = [
    { value: "critical", label: "Critical", color: ACC, order: 0 },
    { value: "high", label: "High", color: N800, order: 1 },
    { value: "normal", label: "Normal", color: N500, order: 2 },
  ];
  const PRIO_MAP = Object.fromEntries(PRIORITIES.map((p) => [p.value, p]));
  
  const REASONS = [
    { value: "", label: "— ยังไม่ระบุ" },
    { value: "spare", label: "รออะไหล่" },
    { value: "budget", label: "รองบประมาณ" },
    { value: "shutdown", label: "รอ Shutdown" },
    { value: "production", label: "รอสายการผลิต" },
    { value: "contractor", label: "รอผู้รับเหมา" },
    { value: "engineering", label: "รอ Engineering" },
    { value: "owner", label: "Owner Pending" },
    { value: "cancel", label: "ขอยกเลิก" },
  ];
  const REASON_MAP = Object.fromEntries(REASONS.map((r) => [r.value, r.label]));
  
  const STATUS_META = {
    NTCR: { label: "NTCR · เปิดงานเริ่มต้น", color: N400 },
    NTAP: { label: "NTAP · เปิดโดย User", color: ACC },
    NTMC: { label: "NTMC · รอ User กดรับ", color: N600 },
    NTAC: { label: "NTAC · เจ้าของพื้นที่รับงานแล้ว", color: N700 },
    NTSA: { label: "NTSA · รอ SHE Approve", color: N700 },
    NSTM: { label: "NSTM · Self-Maintenance (User ทำ)", color: N600 },
    NTSP: { label: "NTSP · Self-Maintenance เสร็จ", color: OK },
    NOCO: { label: "NOCO · ปิดงาน", color: OK },
    NTCN: { label: "NTCN · ยกเลิก", color: OK },
    NTRM: { label: "NTRM · Reject by Maintenance", color: OK },
    "(none)": { label: "ไม่มีสถานะ", color: N400 },
  };
  // สถานะที่ถือว่า "จบงาน / ไม่นับเป็นงานค้างอีกต่อไป":
  //   NOCO = ปิดงานสำเร็จ, NTSP = Self-Maintenance เสร็จ,
  //   NTCN = ยกเลิก, NTRM = ถูก Reject โดยฝ่ายซ่อมบำรุง
  const DONE_STATUSES = new Set(["NOCO", "NTSP", "NTCN", "NTRM"]);
  const KNOWN_STATUS_CODES = ["NTCR", "NTAP", "NTMC", "NTAC", "NTSA", "NSTM", "NTSP", "NOCO", "NTCN", "NTRM"];
  // ช่อง User status ของ SAP มักมีหลายรหัสปนกันในเซลล์เดียว (เช่น "REL NOCO"
  // หรือ "NTAC NOCO") — ต้องแยกเป็น token แล้วหารหัสที่รู้จัก ไม่ใช่เทียบทั้งสตริง
  // ตรง ๆ ไม่งั้นงานที่เพิ่งเปลี่ยนเป็นสถานะจบงานจะหลุดไปนับเป็นงานค้างต่อ
  function classifyStatus(raw) {
    const s = (raw || "").toString().trim().toUpperCase();
    if (!s) return "(none)";
    const tokens = s.split(/[^A-Z0-9]+/).filter(Boolean);
    const found = tokens.filter((t) => KNOWN_STATUS_CODES.includes(t));
    if (found.length) {
      const done = found.find((t) => DONE_STATUSES.has(t));
      return done || found[found.length - 1];
    }
    for (const code of KNOWN_STATUS_CODES) { if (s.includes(code)) return code; }
    return raw.toString().trim() || "(none)";
  }
  function meta(g) { return STATUS_META[g] || { label: g || "—", color: N600 }; }
  
  /* ── formatters ─────────────────────────────────────────────── */
  function fmtInt(n) { return Math.round(n || 0).toLocaleString("en-US"); }
  function fmtM(n) { return "฿" + ((n || 0) / 1e6).toFixed(2) + "M"; }
  function serialMs(s) { return Date.UTC(1899, 11, 30) + s * 86400000; }
  function fmtDate(s) {
    if (!s) return "—";
    const d = new Date(serialMs(s));
    return String(d.getUTCDate()).padStart(2, "0") + "/" + String(d.getUTCMonth() + 1).padStart(2, "0") + "/" + String(d.getUTCFullYear()).slice(2);
  }
  function fmtDateTime(ms) {
    if (!ms) return "";
    const d = new Date(ms);
    return String(d.getDate()).padStart(2, "0") + "/" + String(d.getMonth() + 1).padStart(2, "0") + "/" + d.getFullYear() + " " + String(d.getHours()).padStart(2, "0") + ":" + String(d.getMinutes()).padStart(2, "0");
  }
  function todayISO() { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); }
  function isoInDays(s) { if (!s) return null; return Math.ceil((new Date(s) - new Date()) / 86400000); }
  function ym(ms) { const d = new Date(ms); return d.getUTCFullYear() + "-" + String(d.getUTCMonth() + 1).padStart(2, "0"); }
  const MONTH_TH = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
  function mLabel(k) { const [y, m] = k.split("-"); return MONTH_TH[+m - 1] + " " + String(y).slice(2); }
  const TH_MONTH_FULL = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];
  
  /* ── xlsx parsing ───────────────────────────────────────────── */
  function unzip(buf) {
    const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
    let eocd = -1;
    for (let i = buf.length - 22; i >= 0; i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
    if (eocd < 0) throw new Error("ไม่ใช่ไฟล์ xlsx ที่ถูกต้อง");
    const cdOffset = dv.getUint32(eocd + 16, true), cdCount = dv.getUint16(eocd + 10, true);
    const files = {}; let p = cdOffset;
    for (let n = 0; n < cdCount; n++) {
      if (dv.getUint32(p, true) !== 0x02014b50) break;
      const method = dv.getUint16(p + 10, true), compSize = dv.getUint32(p + 20, true);
      const nameLen = dv.getUint16(p + 28, true), extraLen = dv.getUint16(p + 30, true), commentLen = dv.getUint16(p + 32, true);
      const lho = dv.getUint32(p + 42, true);
      const name = new TextDecoder().decode(buf.subarray(p + 46, p + 46 + nameLen));
      const lnameLen = dv.getUint16(lho + 26, true), lextraLen = dv.getUint16(lho + 28, true);
      const ds = lho + 30 + lnameLen + lextraLen;
      files[name] = { method, comp: buf.subarray(ds, ds + compSize) };
      p += 46 + nameLen + extraLen + commentLen;
    }
    return files;
  }
  async function inflateEntry(e) {
    if (e.method === 0) return e.comp;
    const ds = new DecompressionStream("deflate-raw");
    return new Uint8Array(await new Response(new Blob([e.comp]).stream().pipeThrough(ds)).arrayBuffer());
  }
  async function readEntry(files, name) { return new TextDecoder().decode(await inflateEntry(files[name])); }
  function unesc(s) { return s.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#10;/g, "\n"); }
  function parseSharedStrings(ss) {
    const out = [], re = /<si>(.*?)<\/si>/gs; let m;
    while ((m = re.exec(ss))) { const t = [...m[1].matchAll(/<t[^>]*>(.*?)<\/t>/gs)].map((x) => x[1]).join(""); out.push(unesc(t)); }
    return out;
  }
  function colToNum(c) { let n = 0; for (const x of c) n = n * 26 + (x.charCodeAt(0) - 64); return n; }
  function parseSheet(sheet, strings) {
    const rows = [], rowRe = /<row[^>]*r="(\d+)"[^>]*>(.*?)<\/row>/gs; let rm;
    while ((rm = rowRe.exec(sheet))) {
      const cells = {}, cellRe = /<c r="([A-Z]+)\d+"(?:[^>]*t="([^"]*)")?[^>]*>(?:<v>(.*?)<\/v>|<is><t[^>]*>(.*?)<\/t><\/is>)?<\/c>/gs; let cm;
      while ((cm = cellRe.exec(rm[2]))) {
        const col = colToNum(cm[1]), t = cm[2];
        let val = cm[3] !== undefined ? cm[3] : cm[4];
        if (val === undefined) continue;
        if (t === "s") val = strings[parseInt(val, 10)]; else val = unesc(val);
        cells[col] = val;
      }
      rows.push(cells);
    }
    return rows;
  }
  function colOf(header, name) {
    for (const k in header) { if ((header[k] || "").toString().trim().toLowerCase() === name.toLowerCase()) return +k; }
    for (const k in header) { if ((header[k] || "").toString().trim().toLowerCase().includes(name.toLowerCase())) return +k; }
    return null;
  }
  function numVal(v) { const n = parseFloat(v); return isFinite(n) ? n : 0; }
  
  async function parseWorkbook(bytes, name) {
    const files = unzip(bytes);
    const strings = files["xl/sharedStrings.xml"] ? parseSharedStrings(await readEntry(files, "xl/sharedStrings.xml")) : [];
    const sheetKey = Object.keys(files).find((k) => /worksheets\/sheet1\.xml$/.test(k)) || Object.keys(files).find((k) => /worksheets\/.*\.xml$/.test(k));
    if (!sheetKey) throw new Error("ไม่พบชีตข้อมูลในไฟล์");
    const rows = parseSheet(await readEntry(files, sheetKey), strings);
    const header = rows[0] || {};
    const hvals = Object.values(header).map((v) => (v || "").toString());
    const has = (s) => hvals.some((h) => h.toLowerCase().includes(s.toLowerCase()));
    const co = (n) => colOf(header, n);
    let type = "unknown";
    if (has("TotalPlnndCosts") || has("act.costs")) type = "iw38";
    else if (has("Notifictn") || has("Notif. Date") || has("Reported By")) type = "iw28";
    const records = [];
    if (type === "iw38") {
      const c = { pg: co("Planner Group"), otype: co("Order Type"), order: co("Order"), desc: co("Description"), plan: co("TotalPlnndCosts"), act: co("act.costs"), start: co("start date"), fin: co("fin. date"), us: co("User status") };
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || Object.keys(r).length === 0) continue;
        const ord = (r[c.order] || "").toString().trim(); if (!ord) continue;
        records.push({ pg: (r[c.pg] || "").toString().trim() || "(blank)", otype: (r[c.otype] || "").toString(), order: ord, desc: (r[c.desc] || "").toString(), plan: numVal(r[c.plan]), act: numVal(r[c.act]), start: numVal(r[c.start]), fin: numVal(r[c.fin]), us: (r[c.us] || "").toString().trim() });
      }
    } else if (type === "iw28") {
      const c = { notif: co("Notification"), order: co("Order"), date: co("Notif. Date"), desc: co("Description"), us: co("User status"), reporter: co("Reported By"), pg: co("Planner Group"), prio: co("Priority") };
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || Object.keys(r).length === 0) continue;
        records.push({ notif: (r[c.notif] || "").toString().trim(), order: (r[c.order] || "").toString().trim(), date: numVal(r[c.date]), desc: (r[c.desc] || "").toString(), us: (r[c.us] || "").toString().trim(), reporter: (r[c.reporter] || "").toString(), pg: (r[c.pg] || "").toString().trim() || "(blank)", prio: (r[c.prio] || "").toString() });
      }
    }
    return { type, name, records };
  }
  
  /* ── compute ────────────────────────────────────────────────── */
  function compute(st) {
    const settings = Object.assign({}, DEFAULT_SETTINGS, st.settings || {});
    const overdue = settings.overdueDays, monthsBack = settings.monthsBack;
    const actions = st.actions || {};
    const o28 = st.iw28 || [], o38 = st.iw38 || [], now = Date.now();
  
    const obo = {};
    for (const o of o38) {
      if (!o.order) continue;
      if (!obo[o.order]) obo[o.order] = { plan: 0, act: 0, us: o.us, fin: o.fin };
      obo[o.order].plan += o.plan; obo[o.order].act += o.act;
      if (o.fin) obo[o.order].fin = o.fin;
    }
  
    const recs = o28.map((n) => {
      const ord = n.order, o = ord ? obo[ord] : null, grp = classifyStatus(n.us), openS = n.date || 0;
      const openMs = openS ? serialMs(openS) : null, finMs = o && o.fin ? serialMs(o.fin) : null, isBack = !DONE_STATUSES.has(grp);
      let age = null;
      if (isBack && openMs && openMs <= now) age = Math.floor((now - openMs) / 86400000);
      else if (!isBack && openMs && finMs && finMs >= openMs) age = Math.floor((finMs - openMs) / 86400000);
      const act = actions[n.notif] || {};
      return { notif: n.notif, pg: n.pg, order: ord, matched: !!o, desc: n.desc || "", plan: o ? o.plan : 0, act: o ? o.act : 0, grp, isDone: !isBack, us28: n.us || "", openS, openMs, finMs, isBack, age,
        priority: act.priority || "", reason: act.reason || "", owner: act.owner || "", nextAction: act.nextAction || "", dueDate: act.dueDate || "", followUpDate: act.followUpDate || "" };
    });
  
    const f = st.pgFilter || "ALL", sf = st.statusFilter || "ALL", pf = st.priorityFilter || "ALL", qf = st.quickFilter || "ALL", of = st.ownerFilter || "ALL";
  
    const weekAgo = now - 7 * 86400000;
    const openedThisWeek = recs.filter((r) => r.openMs && r.openMs >= weekAgo).length;
    const closedThisWeek = recs.filter((r) => !r.isBack && r.finMs && r.finMs >= weekAgo).length;
    const velocityNet = openedThisWeek - closedThisWeek;
  
    const STALE_DAYS = 14;
    function isStale(r) {
      if (!r.isBack) return false;
      if (!r.owner && !r.dueDate) return false;
      if (!r.followUpDate) return r.age != null && r.age > STALE_DAYS;
      return (now - new Date(r.followUpDate + "T00:00:00").getTime()) > STALE_DAYS * 86400000;
    }
  
    let view = f === "ALL" ? recs.slice() : recs.filter((r) => r.pg === f);
    if (sf !== "ALL") view = view.filter((r) => r.grp === sf);
    if (pf !== "ALL") view = view.filter((r) => r.priority === pf);
    if (of !== "ALL") view = view.filter((r) => r.owner === of);
    if (qf === "critical") view = view.filter((r) => r.isBack && r.priority === "critical");
    else if (qf === "over180") view = view.filter((r) => r.isBack && r.age != null && r.age > 180);
    else if (qf === "dueWeek") view = view.filter((r) => { if (!r.isBack || !r.dueDate) return false; const d = isoInDays(r.dueDate); return d != null && d >= 0 && d <= 7; });
    else if (qf === "stale") view = view.filter((r) => r.isBack && isStale(r));
    else if (qf === "NTAP") view = view.filter((r) => r.grp === "NTAP");
  
    const usCount = {};
    for (const r of recs) { const k = r.grp || "(none)"; usCount[k] = (usCount[k] || 0) + 1; }
    const statusOptions = [{ value: "ALL", label: "ทุก Status" }].concat(
      Object.keys(usCount).sort((a, b) => usCount[b] - usCount[a]).map((k) => ({ value: k, label: (k === "(none)" ? "ไม่มีใบแจ้ง" : meta(k).label) + " (" + usCount[k] + ")" }))
    );
  
    const pgMap = {};
    for (const r of recs) { const k = r.pg; (pgMap[k] = pgMap[k] || { pg: k, total: 0, backlog: 0, plan: 0, act: 0 }); pgMap[k].total++; if (r.isBack) pgMap[k].backlog++; pgMap[k].plan += r.plan; pgMap[k].act += r.act; }
    const pgArr = Object.values(pgMap).sort((a, b) => b.backlog - a.backlog);
    const maxBack = Math.max(1, ...pgArr.map((p) => p.backlog));
    const pgBars = pgArr.map((p) => {
      const active = f === p.pg;
      return { pg: p.pg, backStr: fmtInt(p.backlog), totalStr: fmtInt(p.total), width: (p.backlog ? Math.max(1.5, (p.backlog / maxBack) * 100) : 0) + "%",
        barColor: p.backlog ? ACC : "transparent", labelColor: p.backlog ? INK : N600, numColor: p.backlog ? INK : N600, active };
    });
  
    const totalOrders = view.length;
    const backlog = view.filter((r) => r.isBack);
    const noco = view.filter((r) => !r.isBack);
    const seenO = {}; let totalPlan = 0, totalAct = 0, matchedN = 0;
    for (const r of view) { if (r.matched) matchedN++; if (r.order && !seenO[r.order]) { seenO[r.order] = 1; totalPlan += r.plan; totalAct += r.act; } }
    const seenB = {}; let backPlan = 0;
    for (const r of backlog) { if (r.order && !seenB[r.order]) { seenB[r.order] = 1; backPlan += r.plan; } }
    const ages = backlog.map((r) => r.age).filter((a) => a != null);
    const avgAge = ages.length ? Math.round(ages.reduce((s, a) => s + a, 0) / ages.length) : 0;
    const maxAge = ages.length ? Math.max(...ages) : 0;
    const compRate = totalOrders ? Math.round((noco.length / totalOrders) * 100) : 0;
    const variance = totalAct - totalPlan;
  
    const over180 = backlog.filter((r) => r.age != null && r.age > 180).length;
    const ntapCount = backlog.filter((r) => r.grp === "NTAP").length;
    const waitSpare = backlog.filter((r) => r.reason === "spare").length;
    const waitBudget = backlog.filter((r) => r.reason === "budget").length;
    const waitEng = backlog.filter((r) => r.reason === "engineering").length;
    const dueThisWeek = backlog.filter((r) => { if (!r.dueDate) return false; const d = isoInDays(r.dueDate); return d != null && d >= 0 && d <= 7; }).length;
    const overdueDue = backlog.filter((r) => { if (!r.dueDate) return false; const d = isoInDays(r.dueDate); return d != null && d < 0; }).length;
    const staleCount = recs.filter((r) => r.isBack && isStale(r)).length;
    const criticalCount = backlog.filter((r) => r.priority === "critical").length;
    const highCount = backlog.filter((r) => r.priority === "high").length;
    const noActionCount = backlog.filter((r) => !r.priority && !r.reason && !r.owner && !r.dueDate).length;
  
    const allOwners = [...new Set(Object.values(actions).map((a) => a.owner).filter(Boolean))].sort();
  
    const mt = settings.monthlyTarget || 0;
    const msc = settings.monthStartCount;
    const targetGap = mt > 0 ? backlog.length - mt : null;
    const targetPct = mt > 0 && msc > mt ? Math.max(0, Math.min(100, Math.round(((msc - backlog.length) / (msc - mt)) * 100))) : 0;
  
    /* KPI cards */
    const kpis = [
      { key: "backlog", label: "งานค้าง (ยังไม่ปิดงาน)", value: fmtInt(backlog.length), sub: totalOrders ? Math.round((backlog.length / totalOrders) * 100) + "% ของใบแจ้งในมุมมองนี้" : "—", accent: true },
      { key: "ntap", label: "NTAP เทียบงานค้างอื่น", value: fmtInt(ntapCount), sub: backlog.length ? Math.round((ntapCount / backlog.length) * 100) + "% ของงานค้างทั้งหมด · อื่น " + fmtInt(backlog.length - ntapCount) + " งาน" : "—", accent: true },
      { key: "total", label: "ใบแจ้งทั้งหมด (IW28)", value: fmtInt(totalOrders), sub: "จับคู่ Order ได้ " + fmtInt(matchedN) + (totalOrders ? " (" + Math.round((matchedN / totalOrders) * 100) + "%)" : ""), accent: false },
      { key: "noco", label: "ปิดงานแล้ว (NOCO/ยกเลิก/Reject)", value: fmtInt(noco.length), sub: compRate + "% completion rate", accent: false },
      { key: "age", label: "อายุงานค้างเฉลี่ย", value: fmtInt(avgAge), sub: "วัน · ค้างนานสุด " + fmtInt(maxAge) + " วัน", accent: false },
      { key: "plan", label: "มูลค่า Plan รวม", value: fmtM(totalPlan), sub: fmtInt(totalPlan) + " ฿", accent: false },
      { key: "act", label: "มูลค่า Actual รวม", value: fmtM(totalAct), sub: fmtInt(totalAct) + " ฿", accent: false },
      { key: "var", label: "ส่วนต่าง Act − Plan", value: (variance >= 0 ? "+" : "−") + fmtM(Math.abs(variance)), sub: totalPlan ? Math.round((variance / totalPlan) * 100) + "% เทียบแผน" : "—", accent: variance > 0 },
      { key: "backplan", label: "มูลค่างานค้าง (Plan)", value: fmtM(backPlan), sub: fmtInt(backPlan) + " ฿ รอดำเนินการ", accent: true },
    ];
  
    const quickTabs = [
      { key: "ALL", label: "ทั้งหมด", count: fmtInt(recs.filter((r) => (f === "ALL" || r.pg === f)).length), dot: INK, active: qf === "ALL" },
      { key: "NTAP", label: "NTAP (เปิดโดย User)", count: fmtInt(ntapCount), dot: ACC, active: qf === "NTAP" },
      { key: "critical", label: "Critical", count: fmtInt(criticalCount), dot: ACC, active: qf === "critical" },
      { key: "over180", label: "ค้าง > 180 วัน", count: fmtInt(over180), dot: N800, active: qf === "over180" },
      { key: "dueWeek", label: "ครบกำหนด 7 วัน", count: fmtInt(dueThisWeek), dot: N600, active: qf === "dueWeek" },
      { key: "stale", label: "Action ค้าง", count: fmtInt(staleCount), dot: N400, active: qf === "stale" },
    ];
  
    const blockers = [
      { label: "ค้างเกิน 180 วัน", n: over180, alert: true },
      { label: "รออะไหล่", n: waitSpare, alert: false },
      { label: "รองบประมาณ", n: waitBudget, alert: false },
      { label: "รอ Engineering", n: waitEng, alert: false },
      { label: "เลย Due Date", n: overdueDue, alert: true },
      { label: "ยังไม่ระบุ Action", n: noActionCount, alert: false },
      { label: "Action ค้าง (ไม่ Follow-up)", n: staleCount, alert: false },
    ];
    const maxBlocker = Math.max(1, ...blockers.map((b) => b.n));
    const blockerRows = blockers.map((b) => ({ label: b.label, count: fmtInt(b.n), width: (b.n ? Math.max(3, (b.n / maxBlocker) * 100) : 0) + "%",
      color: b.alert && b.n ? ACC : b.n ? N600 : "transparent", textColor: b.n ? INK : N600 }));
  
    /* Status split */
    const stOrder = ["NTCR", "NTAP", "NTMC", "NTAC", "NTSA", "NSTM", "NTSP", "NOCO", "NTCN", "NTRM", "(none)"];
    const stMap = {}; for (const r of view) stMap[r.grp] = (stMap[r.grp] || 0) + 1;
    const stTotal = view.length || 1;
    const seenSt = stOrder.reduce((m, k) => ((m[k] = 1), m), {});
    for (const k in stMap) if (!seenSt[k]) stOrder.push(k);
    const statusRows = stOrder.filter((k) => stMap[k]).map((k) => ({ label: meta(k).label, count: fmtInt(stMap[k]), pct: Math.round((stMap[k] / stTotal) * 100), width: Math.max(1, (stMap[k] / stTotal) * 100) + "%", color: meta(k).color }));
  
    /* Aging columns */
    const buckets = [{ label: "0–30", max: 30, color: N400 }, { label: "31–60", max: 60, color: N500 }, { label: "61–90", max: 90, color: N700 }, { label: "91–180", max: 180, color: N800 }, { label: "180+", max: 1e9, color: ACC }];
    const bc = buckets.map(() => 0);
    for (const r of backlog) { if (r.age == null) continue; for (let i = 0; i < buckets.length; i++) { if (r.age <= buckets[i].max) { bc[i]++; break; } } }
    const maxBkt = Math.max(1, ...bc);
    const agingRows = buckets.map((b, i) => ({ label: b.label, count: fmtInt(bc[i]), height: (bc[i] ? Math.max(2, (bc[i] / maxBkt) * 100) : 0) + "%", color: b.color, numColor: i === 4 && bc[i] ? ACC : INK }));
  
    /* Monthly trend */
    const mOpen = {}, mClose = {};
    for (const r of view) {
      if (r.openMs) { const k = ym(r.openMs); mOpen[k] = (mOpen[k] || 0) + 1; }
      if (!r.isBack && r.finMs) { const k = ym(r.finMs); mClose[k] = (mClose[k] || 0) + 1; }
    }
    let mKeys = Object.keys(Object.assign({}, mOpen, mClose)).filter((k) => k >= "2024-01" && k <= "2027-12").sort();
    mKeys = mKeys.slice(-monthsBack);
    const chartW = 900, chartH = 150, padY = 10;
    const maxM = Math.max(1, ...mKeys.map((k) => Math.max(mOpen[k] || 0, mClose[k] || 0)));
    const colW = mKeys.length ? chartW / mKeys.length : chartW;
    const pts = mKeys.map((k, i) => {
      const x = i * colW + colW / 2;
      return { x, yo: padY + (1 - (mOpen[k] || 0) / maxM) * (chartH - padY * 2), yc: padY + (1 - (mClose[k] || 0) / maxM) * (chartH - padY * 2) };
    });
    const openedPath = pts.map((p) => p.x.toFixed(1) + "," + p.yo.toFixed(1)).join(" ");
    const closedPath = pts.map((p) => p.x.toFixed(1) + "," + p.yc.toFixed(1)).join(" ");
    const monthChart = { width: chartW, height: chartH, openedPath, closedPath, maxLabel: fmtInt(maxM),
      labels: mKeys.map((k) => ({ text: mLabel(k) })) };
  
    /* Yearly */
    const yMap = {};
    for (const r of view) {
      if (r.openMs) { const y = new Date(r.openMs).getUTCFullYear(); (yMap[y] = yMap[y] || { year: y, opened: 0, closed: 0, backlog: 0, plan: 0, act: 0 }); yMap[y].opened++; if (r.isBack) yMap[y].backlog++; yMap[y].plan += r.plan; yMap[y].act += r.act; }
      if (!r.isBack && r.finMs) { const y = new Date(r.finMs).getUTCFullYear(); (yMap[y] = yMap[y] || { year: y, opened: 0, closed: 0, backlog: 0, plan: 0, act: 0 }); yMap[y].closed++; }
    }
    const yearRows = Object.values(yMap).sort((a, b) => a.year - b.year).map((y) => ({ year: String(y.year), opened: fmtInt(y.opened), closed: fmtInt(y.closed), backlog: fmtInt(y.backlog), plan: fmtM(y.plan), act: fmtM(y.act) }));
  
    /* Plan vs Actual per PG */
    const maxPA = Math.max(1, ...pgArr.map((p) => Math.max(p.plan, p.act)));
    const paRows = pgArr.slice().filter((p) => p.plan || p.act).sort((a, b) => b.plan - a.plan).map((p) => ({ pg: p.pg,
      planWidth: (p.plan ? Math.max(1, (p.plan / maxPA) * 100) : 0) + "%", actWidth: (p.act ? Math.max(1, (p.act / maxPA) * 100) : 0) + "%",
      planStr: fmtM(p.plan), actStr: fmtM(p.act) }));
  
    /* Row decoration */
    function decorate(r) {
      const p = r.priority ? PRIO_MAP[r.priority] : null;
      const dueDays = isoInDays(r.dueDate);
      let dueLabel = "—", dueColor = N600, dueBg = "transparent", dueWeight = "400";
      if (r.dueDate) {
        dueLabel = r.dueDate;
        if (dueDays < 0) { dueColor = ACC; dueBg = ACC1; dueWeight = "800"; }
        else if (dueDays === 0) { dueColor = ACC; dueBg = ACC1; dueWeight = "800"; }
        else if (dueDays <= 7) { dueColor = INK; dueBg = N200; dueWeight = "600"; }
        else { dueColor = N700; dueBg = "transparent"; dueWeight = "400"; }
      }
      let dueWord = "ยังไม่กำหนด";
      if (r.dueDate) dueWord = dueDays < 0 ? "เลยกำหนด " + Math.abs(dueDays) + " วัน" : dueDays === 0 ? "วันนี้" : "อีก " + dueDays + " วัน";
      return { notif: r.notif, order: r.order || "—", desc: r.desc || "(ไม่มีรายละเอียด)", pg: r.pg,
        statusLabel: meta(r.grp).label, statusColor: meta(r.grp).color, isDone: !!r.isDone,
        openStr: fmtDate(r.openS), age: r.age == null ? "—" : fmtInt(r.age),
        ageColor: r.age != null && r.age > overdue ? ACC : INK,
        planStr: r.matched && r.plan ? fmtInt(r.plan) : "—",
        prioColor: p ? p.color : N300, prioLabel: p ? p.label : "—", priority: r.priority,
        reason: r.reason, reasonLabel: REASON_MAP[r.reason] || "ยังไม่ระบุ",
        owner: r.owner, ownerLabel: r.owner || "ยังไม่มี", ownerColor: r.owner ? INK : N600,
        nextAction: r.nextAction, followUpDate: r.followUpDate,
        dueDate: r.dueDate, dueLabel, dueColor, dueBg, dueWeight, dueWord,
        stale: isStale(r), dueDays,
        rowBg: r.isDone ? OK_BG : (r.priority === "critical" || (r.dueDate && dueDays != null && dueDays <= 0) ? ACC1 : "transparent") };
    }
  
    /* Today queue — what needs doing now */
    const todaySrc = backlog.slice().map((r) => {
      const d = isoInDays(r.dueDate);
      let score = 0;
      if (r.dueDate && d != null && d <= 0) score += 1000 - d;
      else if (r.priority === "critical") score += 900;
      else if (r.dueDate && d != null && d <= 7) score += 800 - d;
      else if (r.priority === "high") score += 700;
      else if (r.age != null && r.age > 180) score += 600;
      else if (isStale(r)) score += 500;
      return { r, score };
    }).filter((x) => x.score > 0).sort((a, b) => b.score - a.score || (b.r.age || 0) - (a.r.age || 0));
    const todayRows = todaySrc.slice(0, 8).map((x) => decorate(x.r));
    const todayTotal = todaySrc.length;
  
    /* Full table */
    const trowsAll = (st.scope === "all" ? view.slice() : backlog.slice());
    const prioOrder = { critical: 0, high: 1, normal: 2, "": 3 };
    trowsAll.sort((a, b) => { const pd = (prioOrder[a.priority] || 3) - (prioOrder[b.priority] || 3); return pd !== 0 ? pd : (b.age || 0) - (a.age || 0); });
    const cap = 200, tableTotal = trowsAll.length;
    const tableRows = trowsAll.slice(0, cap).map(decorate);

    /* Due Date table/calendar — any record (open or already-done) that has
       a Due Date value set. Done ones stay visible but get highlighted
       green in decorate()/isDone so it's obvious they no longer need action. */
    const dueSrc = view.filter((r) => r.dueDate);
    dueSrc.sort((a, b) => { const da = isoInDays(a.dueDate), db = isoInDays(b.dueDate); return (da == null ? 1e9 : da) - (db == null ? 1e9 : db); });
    const dueRows = dueSrc.map(decorate);
    const dueTotal = dueRows.length;

    /* Group by exact date (YYYY-MM-DD) for calendar view */
    const dueByDate = {};
    for (const r of dueRows) { (dueByDate[r.dueDate] = dueByDate[r.dueDate] || []).push(r); }
  
    const asOf = new Date(now);
    const asOfStr = String(asOf.getDate()).padStart(2, "0") + "/" + String(asOf.getMonth() + 1).padStart(2, "0") + "/" + asOf.getFullYear();
    const asOfLong = asOf.getDate() + " " + TH_MONTH_FULL[asOf.getMonth()] + " " + asOf.getFullYear();
  
    const topPg = pgArr[0];
    const reasonGroups = {};
    for (const r of backlog) if (r.reason) reasonGroups[r.reason] = (reasonGroups[r.reason] || 0) + 1;
    const topReason = Object.entries(reasonGroups).sort((a, b) => b[1] - a[1])[0];
  
    const headline = "งานค้าง " + fmtInt(backlog.length) + " งาน อายุเฉลี่ย " + fmtInt(avgAge) + " วัน";
    const headSubParts = [];
    if (topPg && topPg.backlog) headSubParts.push(topPg.pg + " ค้างมากที่สุด " + fmtInt(topPg.backlog) + " งาน");
    if (maxAge) headSubParts.push("ค้างนานสุด " + fmtInt(maxAge) + " วัน");
    headSubParts.push("สัปดาห์นี้เปิดใหม่ " + fmtInt(openedThisWeek) + " ปิดได้ " + fmtInt(closedThisWeek) + " สุทธิ " + (velocityNet >= 0 ? "+" : "−") + fmtInt(Math.abs(velocityNet)) + " งาน");
    const headSub = headSubParts.join(" · ");
  
    const velVerdict = velocityNet > 0 ? "กองงานค้างเพิ่มขึ้น" : velocityNet < 0 ? "กองงานค้างลดลง" : "คงที่";
  
    const meetingLines = [
      "รายงานสถานะงานซ่อมบำรุง ณ วันที่ " + asOfStr,
      "",
      "ภาพรวม",
      "• ใบแจ้งทั้งหมด (IW28): " + fmtInt(totalOrders) + " ใบ | เสร็จ (NOCO): " + fmtInt(noco.length) + " งาน (" + compRate + "%)",
      "• งานค้าง: " + fmtInt(backlog.length) + " งาน | อายุเฉลี่ย: " + fmtInt(avgAge) + " วัน | ค้างนานสุด: " + fmtInt(maxAge) + " วัน",
      "• Velocity 7 วัน: เปิด " + fmtInt(openedThisWeek) + " / ปิด " + fmtInt(closedThisWeek) + " → " + velVerdict,
      ...(criticalCount > 0 ? ["• งานวิกฤต (Critical): " + fmtInt(criticalCount) + " งาน"] : []),
      "",
      "PG สถานะ",
      ...pgArr.filter((p) => p.backlog > 0).map((p) => "• " + p.pg + ": ค้าง " + fmtInt(p.backlog) + " / ทั้งหมด " + fmtInt(p.total) + " งาน"),
      "",
      "จุดที่ต้องติดตาม",
      over180 > 0 ? "• ค้างเกิน 180 วัน: " + fmtInt(over180) + " งาน (ควรประเมินยกเลิกหรือกำหนด Due Date ใหม่)" : "• ไม่มีงานค้างเกิน 180 วัน",
      dueThisWeek > 0 ? "• ครบกำหนดสัปดาห์นี้: " + fmtInt(dueThisWeek) + " งาน" : "• ไม่มีงานครบกำหนดสัปดาห์นี้",
      ...(overdueDue > 0 ? ["• เลย Due Date แล้ว: " + fmtInt(overdueDue) + " งาน"] : []),
      ...(topReason && topReason[1] >= 2 ? ["• " + (REASON_MAP[topReason[0]] || topReason[0]) + ": " + fmtInt(topReason[1]) + " งาน"] : []),
      "",
      "มูลค่า",
      "• Plan: " + fmtM(totalPlan) + " | Actual: " + fmtM(totalAct) + " | ส่วนต่าง: " + (variance >= 0 ? "+" : "−") + fmtM(Math.abs(variance)) + " (" + (totalPlan ? Math.round((variance / totalPlan) * 100) : 0) + "%)",
      "• มูลค่างานค้าง (Plan): " + fmtM(backPlan),
    ].join("\n");
  
    return { kpis, quickTabs, pgBars, statusRows, agingRows, blockerRows, monthChart, yearRows, paRows,
      todayRows, todayTotal, tableRows, tableTotal, tableShown: Math.min(cap, tableTotal),
      dueRows, dueTotal, dueByDate,
      asOfStr, asOfLong, c28: fmtInt(o28.length), c38: fmtInt(o38.length), statusOptions, allOwners,
      headline, headSub, criticalCount: fmtInt(criticalCount), highCount: fmtInt(highCount),
      over180Str: fmtInt(over180), dueWeekStr: fmtInt(dueThisWeek), overdueDueStr: fmtInt(overdueDue),
      staleStr: fmtInt(staleCount), openedStr: fmtInt(openedThisWeek), closedStr: fmtInt(closedThisWeek),
      velNetStr: (velocityNet >= 0 ? "+" : "−") + fmtInt(Math.abs(velocityNet)), velVerdict, velUp: velocityNet > 0,
      backlogCount: backlog.length, backlogStr: fmtInt(backlog.length), avgAgeStr: fmtInt(avgAge), maxAgeStr: fmtInt(maxAge),
      matchedStr: fmtInt(matchedN), nocoStr: fmtInt(noco.length), compRate, totalStr: fmtInt(totalOrders),
      planStr: fmtM(totalPlan), actStr: fmtM(totalAct), varStr: (variance >= 0 ? "+" : "−") + fmtM(Math.abs(variance)),
      varPct: totalPlan ? Math.round((variance / totalPlan) * 100) + "%" : "—", backPlanStr: fmtM(backPlan), backPlanBaht: fmtInt(backPlan) + " ฿",
      mt, targetGap, targetPct, targetPctStr: targetPct + "%", meetingLines };
  }

    return { INK, ACC, ACC7, ACC1, N200, N300, N400, N500, N600, N700, N800, DEFAULT_SETTINGS, PRIORITIES, PRIO_MAP, REASONS, REASON_MAP, meta, fmtInt, fmtM, serialMs, fmtDate, fmtDateTime, todayISO, isoInDays, parseWorkbook, compute };
  })();

// ===================== firebase-init.js (merged) =====================
  // ---- firebase-init.js (merged; converted static imports -> dynamic import()
  //      so this can run as a classic <script>, not type="module") ----
  (async () => {
    const {initializeApp} = await import("https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js");
    const {getFirestore,
  doc,
  setDoc,
  onSnapshot,} = await import("https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js");

  const firebaseConfig = {
    apiKey: "AIzaSyBtY_7oiNaQGgtJR9V_vpbBOxOsO62qnHs",
    authDomain: "cm-maintenance-cost.firebaseapp.com",
    projectId: "cm-maintenance-cost",
    storageBucket: "cm-maintenance-cost.firebasestorage.app",
    messagingSenderId: "1030249075400",
    appId: "1:1030249075400:web:dc3274cf059128bb078bfd",
    measurementId: "G-CYNHSD8Q22",
  };
  const app = initializeApp(firebaseConfig);
  const db = getFirestore(app);
  // Each xlsx source gets its own document so a single update never has to
  // rewrite the other source's (larger) payload, and so neither document gets
  // anywhere close to Firestore's 1MB-per-document ceiling.
  const COLLECTION = "dashboard";
  window.__fb = {
    /** Overwrite the shared document for one source ("iw28" | "iw38"). */
    async saveDoc(id, data) {
      await setDoc(doc(db, COLLECTION, id), data);
    },
    /** Subscribe to live updates for one source. Returns an unsubscribe fn.
     *  `cb` fires immediately with the current value (or null) and again
     *  every time any visitor overwrites that document. */
    subscribe(id, cb) {
      return onSnapshot(
        doc(db, COLLECTION, id),
        (snap) => cb(snap.exists() ? snap.data() : null),
        (err) => {
          console.error("[firebase] subscribe error for", id, err);
          cb(null, err);
        }
      );
    },
  };
  window.dispatchEvent(new Event("fb-ready"));
  })();
