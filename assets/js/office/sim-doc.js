/* Office Academy — simulateur Microsoft Word (kind "doc")
   JavaScript vanilla, aucune dépendance (file:// et GitHub Pages).
   Contrat : assets/js/office/SPEC.md, section « kind "doc" ».
   Tout est scoppé à `host` (classes .wd-*), plusieurs instances peuvent coexister. */
(function () {
  "use strict";
  if (!window.OFFICE || typeof window.OFFICE.sim !== "function") return;

  /* ------------------------------------------------------------------ constantes */
  const IS_MAC = /Mac|iPhone|iPad|iPod/i.test(navigator.platform || navigator.userAgent || "");
  const MOD = IS_MAC ? "⌘" : "Ctrl";
  const CM = 37.7953; // px par cm à 96 dpi
  const BLOCK_RE = /^(P|H[1-6]|BLOCKQUOTE|LI|PRE|DIV|TD|TH)$/;
  const BLOCK_SEL = "p,h1,h2,h3,h4,h5,h6,blockquote,li,pre,div,td,th";
  const BLOCKY_RE = /^(P|H[1-6]|BLOCKQUOTE|PRE|DIV|UL|OL|TABLE|HR|SECTION|FIGURE|ARTICLE)$/;
  const FONTS = [
    ["Aptos", 'Aptos, Calibri, Carlito, "Segoe UI", Arial, sans-serif'],
    ["Aptos Display", '"Aptos Display", Aptos, Calibri, Carlito, "Segoe UI", Arial, sans-serif'],
    ["Calibri", 'Calibri, Carlito, "Segoe UI", Arial, sans-serif'],
    ["Arial", "Arial, Helvetica, sans-serif"],
    ["Times New Roman", '"Times New Roman", Times, serif'],
    ["Georgia", "Georgia, serif"],
    ["Garamond", 'Garamond, "EB Garamond", Georgia, serif'],
    ["Verdana", "Verdana, Geneva, sans-serif"],
    ["Courier New", '"Courier New", Courier, monospace'],
    ["Comic Sans MS", '"Comic Sans MS", "Comic Neue", cursive']
  ];
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 36, 48, 72];
  const STYLES = [["p", "Normal"], ["h1", "Titre 1"], ["h2", "Titre 2"], ["h3", "Titre 3"], ["blockquote", "Citation"]];
  const STYLE_KEYS = { p: "Ctrl+Maj+N", h1: "Ctrl+Alt+1", h2: "Ctrl+Alt+2", h3: "Ctrl+Alt+3" };
  const THEME = ["#000000", "#FFFFFF", "#0E2841", "#E8E8E8", "#156082", "#E97132", "#196B24", "#0F9ED5", "#A02B93", "#4EA72E"];
  const STD = ["#C00000", "#FF0000", "#FFC000", "#FFFF00", "#92D050", "#00B050", "#00B0F0", "#0070C0", "#002060", "#7030A0"];
  const HILITE = [["#FFFF00", "Jaune"], ["#00FF00", "Vert vif"], ["#00FFFF", "Turquoise"], ["#FF00FF", "Rose"], ["#0000FF", "Bleu"],
    ["#FF0000", "Rouge"], ["#000080", "Bleu foncé"], ["#008080", "Bleu-vert"], ["#008000", "Vert"], ["#800080", "Violet"],
    ["#800000", "Rouge foncé"], ["#808000", "Jaune foncé"], ["#808080", "Gris 50 %"], ["#C0C0C0", "Gris 25 %"], ["#000000", "Noir"]];
  const MARGINS = {
    normal: { l: "Normales", y: 2.5, x: 2.5 },
    narrow: { l: "Étroites", y: 1.27, x: 1.27 },
    moderate: { l: "Modérées", y: 2.54, x: 1.91 },
    wide: { l: "Larges", y: 2.54, x: 5.08 }
  };
  let UID = 0;

  /* ------------------------------------------------------------------ icônes (SVG inline, viewBox 24) */
  const T = (x, y, s, txt, w) => `<text x="${x}" y="${y}" style="font-size:${s}px${w ? ";font-weight:" + w : ""}">${txt}</text>`;
  const ICONS = {
    word: '<rect x="2" y="2" width="20" height="20" rx="5" class="f0"/><path d="M6.2 7.5 8.4 16.5 11 9.6h2l2.6 6.9 2.2-9" class="w"/>',
    paste: '<rect x="5" y="4.5" width="14" height="16.5" rx="2"/><rect x="8.8" y="2.6" width="6.4" height="4" rx="1.2" class="fa"/><path d="M8.5 11h7M8.5 14.3h7M8.5 17.6h4.5"/>',
    cut: '<circle cx="6.5" cy="17.5" r="2.7"/><circle cx="17.5" cy="17.5" r="2.7"/><path d="M8.4 15.4 17 3.8M15.6 15.4 7 3.8"/>',
    copy: '<rect x="8.5" y="8.5" width="11.5" height="12.5" rx="2"/><path d="M15.5 8.5V5a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v9.5a2 2 0 0 0 2 2h2.5"/>',
    undo: '<path d="M9 6.5 4.5 11 9 15.5"/><path d="M5 11h9.5a5 5 0 0 1 0 10H12"/>',
    redo: '<path d="m15 6.5 4.5 4.5-4.5 4.5"/><path d="M19 11H9.5a5 5 0 0 0 0 10H12"/>',
    save: '<path d="M5 3.5h11l3.5 3.5v12a1.5 1.5 0 0 1-1.5 1.5H6A1.5 1.5 0 0 1 4.5 19V4a.5.5 0 0 1 .5-.5z"/><path d="M8 3.5v4.5h7V3.5M7.5 20.5v-6h9v6"/>',
    ul: '<circle cx="5" cy="6.5" r="1.5" class="fa"/><circle cx="5" cy="12" r="1.5" class="fa"/><circle cx="5" cy="17.5" r="1.5" class="fa"/><path d="M9.5 6.5H20M9.5 12H20M9.5 17.5H20"/>',
    ol: T(2.2, 9, 7, "1", 700) + T(2.2, 14.6, 7, "2", 700) + T(2.2, 20.2, 7, "3", 700) + '<path d="M9.5 6.5H20M9.5 12H20M9.5 17.5H20"/>',
    outdent: '<path d="M4 4.5h16M11 9.5h9M11 14.5h9M4 19.5h16"/><path d="M7.5 9 4.5 12l3 3" class="a"/>',
    indent: '<path d="M4 4.5h16M11 9.5h9M11 14.5h9M4 19.5h16"/><path d="m4.5 9 3 3-3 3" class="a"/>',
    alL: '<path d="M4 5h16M4 9.5h10M4 14h16M4 18.5h10"/>',
    alC: '<path d="M4 5h16M7 9.5h10M4 14h16M7 18.5h10"/>',
    alR: '<path d="M4 5h16M10 9.5h10M4 14h16M10 18.5h10"/>',
    alJ: '<path d="M4 5h16M4 9.5h16M4 14h16M4 18.5h16"/>',
    pilcrow: '<path d="M11 4h8M13.5 4v16M17 4v16M11 4a4 4 0 0 0 0 8h2.5"/>',
    fcolor: '<path d="m7 15.5 5-12 5 12M8.8 11.2h6.4"/><rect x="3.5" y="18" width="17" height="3.5" rx=".8" class="bar"/>',
    hilite: '<path d="m13.5 3.5 5 5-7 7H6.5v-5z"/><path d="m9 12.5 2.5 2.5M6.5 15.5l-2 2"/><rect x="3.5" y="18" width="17" height="3.5" rx=".8" class="bar"/>',
    clear: '<path d="m3.5 16 4.5-11 4.5 11M5.2 12h5.6"/><path d="m14.5 13.5 6 6M20.5 13.5l-6 6" class="r"/>',
    grow: T(2, 18, 15, "A", 600) + '<path d="m15.5 9 3-3.5 3 3.5" class="a"/>',
    shrink: T(3, 18, 12, "A", 600) + '<path d="m15.5 6 3 3.5 3-3.5" class="a"/>',
    sub: T(3, 15, 13, "x", 600) + T(13, 21, 8, "2", 700),
    sup: T(3, 18, 13, "x", 600) + T(13, 10, 8, "2", 700),
    find: '<circle cx="10.5" cy="10.5" r="6"/><path d="m15 15 5.5 5.5"/>',
    replace: T(2, 10, 8, "ab", 700) + '<path d="M13 6.5h5.5v4.5M16.5 9l2 2 2-2" class="a"/>' + T(10, 21, 8, "ac", 700),
    selall: '<rect x="4" y="4" width="16" height="16" rx="1.5" stroke-dasharray="2.6 2"/><path d="M8 9h8M8 12h8M8 15h5"/>',
    table: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/><path d="M3.5 9.5h17M3.5 14.5h17M9.2 4.5v15M14.8 4.5v15"/><path d="M3.5 6a1.5 1.5 0 0 1 1.5-1.5h14A1.5 1.5 0 0 1 20.5 6v3.5h-17z" class="fa o"/>',
    pagebreak: '<path d="M6 2.5v6h12v-6M6 21.5v-6h12v6"/><path d="M2.5 12h3M8 12h2.5M13.5 12H16M18.5 12h3" class="a"/>',
    link: '<path d="M10 14a4.5 4.5 0 0 0 6.4 0l3-3a4.5 4.5 0 0 0-6.4-6.4l-1.2 1.2"/><path d="M14 10a4.5 4.5 0 0 0-6.4 0l-3 3a4.5 4.5 0 0 0 6.4 6.4l1.2-1.2"/>',
    comment: '<path d="M4 4.5h16v11.5H10.5L6 20v-4H4z"/><path d="M8 8.8h8M8 12h5" class="a"/>',
    cdel: '<path d="M4 4.5h16v11.5H10.5L6 20v-4H4z"/><path d="m9.5 7.5 5 5M14.5 7.5l-5 5" class="r"/>',
    header: '<rect x="5" y="3" width="14" height="18" rx="1.5"/><rect x="7.5" y="5.5" width="9" height="3.2" rx=".6" class="fa"/><path d="M8 12.5h8M8 15.5h8M8 18.5h5"/>',
    footer: '<rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M8 6h8M8 9h8M8 12h5"/><rect x="7.5" y="15.3" width="9" height="3.2" rx=".6" class="fa"/>',
    pagenum: '<rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M10.6 9 9.6 17M14.4 9l-1 8M8.4 11.5h7.4M8 14.5h7.4" class="a"/>',
    date: '<rect x="3.5" y="5" width="17" height="15" rx="2"/><path d="M3.5 9.8h17M8 3v4M16 3v4"/><rect x="7" y="12.5" width="3" height="3" rx=".5" class="fa"/>',
    columns: '<path d="M3.5 5h7M3.5 8.5h7M3.5 12h7M3.5 15.5h7M3.5 19h7M13.5 5h7M13.5 8.5h7M13.5 12h7M13.5 15.5h7M13.5 19h4"/>',
    col1: '<rect x="5" y="3" width="14" height="18" rx="1.5"/><path d="M8 7h8M8 10h8M8 13h8M8 16h8"/>',
    col2: '<rect x="4" y="3" width="16" height="18" rx="1.5"/><path d="M6.5 7h4.5M6.5 10h4.5M6.5 13h4.5M6.5 16h4.5M13 7h4.5M13 10h4.5M13 13h4.5M13 16h3"/>',
    col3: '<rect x="3" y="3" width="18" height="18" rx="1.5"/><path d="M5.3 7h3M5.3 10h3M5.3 13h3M5.3 16h3M10.5 7h3M10.5 10h3M10.5 13h3M10.5 16h3M15.7 7h3M15.7 10h3M15.7 13h2"/>',
    margins: '<rect x="5" y="3" width="14" height="18" rx="1"/><rect x="8" y="6" width="8" height="12" stroke-dasharray="2 1.6" class="a"/>',
    orient: '<rect x="3" y="8.5" width="11" height="13" rx="1"/><rect x="9" y="2.5" width="12.5" height="9" rx="1" class="a"/>',
    portrait: '<rect x="6" y="3" width="12" height="18" rx="1.2"/><path d="M9 7h6M9 10h6M9 13h6"/>',
    landscape: '<rect x="3" y="6" width="18" height="12" rx="1.2"/><path d="M6.5 9.5h11M6.5 12.5h11M6.5 15h7"/>',
    toc: '<path d="M4 5h11M4 9.5h11M7 14h8M7 18.5h8"/><path d="M18 5h2M18 9.5h2M18 14h2M18 18.5h2" class="a"/>',
    refresh: '<path d="M19.5 10.5A7.7 7.7 0 0 0 5.6 7M4.5 13.5A7.7 7.7 0 0 0 18.4 17"/><path d="M5 3v4.4h4.4M19 21v-4.4h-4.4" class="a"/>',
    footnote: T(1.5, 18, 12, "ab", 700) + T(15.5, 10, 9, "1", 800) + '<path d="M2.5 21.5h11" class="a"/>',
    stats: T(1.5, 10.5, 8.5, "ABC", 800) + T(2.5, 21, 8.5, "123", 800),
    nav: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M9.5 4v16"/><path d="M5.2 8h2.4M5.2 11h2.4M5.2 14h2.4" class="a"/>',
    ruler: '<rect x="2.5" y="8" width="19" height="8" rx="1.2"/><path d="M6 8v3M9.5 8v4.5M13 8v3M16.5 8v4.5M20 8v3"/>',
    zoom: '<circle cx="10.5" cy="10.5" r="6"/><path d="m15 15 5.5 5.5M8 10.5h5M10.5 8v5"/>',
    onepage: '<rect x="6" y="3" width="12" height="18" rx="1.2"/><path d="M3 12h1.5M19.5 12H21" class="a"/>',
    pagewidth: '<rect x="6" y="3" width="12" height="18" rx="1.2"/><path d="M2.5 12h5M16.5 12h5M4.5 10l-2 2 2 2M19.5 10l2 2-2 2" class="a"/>',
    z100: T(2.5, 16, 9, "100", 800),
    trash: '<path d="M4 7h16M9.5 7V4.5h5V7M6.3 7l1 13h9.4l1-13"/><path d="M10 11v5.5M14 11v5.5"/>',
    close: '<path d="M6.5 6.5l11 11M17.5 6.5l-11 11"/>',
    check: '<path d="m5 12.5 4.5 4.5L19 7.5"/>',
    cloud: '<path d="M7 18.5a4.5 4.5 0 0 1-.6-9A6 6 0 0 1 18 10a4 4 0 0 1-.5 8.5z"/>',
    prev: '<path d="m14.5 6-6 6 6 6"/>',
    next: '<path d="m9.5 6 6 6-6 6"/>',
    up: '<path d="m6 14.5 6-6 6 6"/>',
    down: '<path d="m6 9.5 6 6 6-6"/>',
    chev: '<path d="m7.5 10 4.5 4.5 4.5-4.5"/>',
    download: '<path d="M12 4v11M7 10.5l5 5 5-5M5 20h14"/>',
    back: '<path d="M19 12H5.5M11 6l-6 6 6 6"/>',
    share: '<path d="M12 15V3.5M7.5 8 12 3.5 16.5 8M5 12.5V20h14v-7.5"/>',
    info: '<circle cx="12" cy="12" r="8.5"/><path d="M12 11v5.5M12 7.6v.2"/>',
    rowA: '<rect x="3" y="11" width="18" height="9.5" rx="1.2"/><path d="M3 15.8h18M9 11v9.5M15 11v9.5"/><path d="M12 2.5v6M9 5.5h6" class="a"/>',
    rowB: '<rect x="3" y="3.5" width="18" height="9.5" rx="1.2"/><path d="M3 8.2h18M9 3.5V13M15 3.5V13"/><path d="M12 15.5v6M9 18.5h6" class="a"/>',
    colL: '<rect x="11" y="3" width="9.5" height="18" rx="1.2"/><path d="M15.8 3v18M11 9h9.5M11 15h9.5"/><path d="M2.5 12h6M5.5 9v6" class="a"/>',
    colR: '<rect x="3.5" y="3" width="9.5" height="18" rx="1.2"/><path d="M8.2 3v18M3.5 9H13M3.5 15H13"/><path d="M15.5 12h6M18.5 9v6" class="a"/>',
    delRow: '<rect x="3" y="3.5" width="18" height="17" rx="1.2"/><path d="M3 9.2h18M3 14.8h18"/><path d="m9.5 10.2 5 3.6M14.5 10.2l-5 3.6" class="r"/>',
    delCol: '<rect x="3.5" y="3" width="17" height="18" rx="1.2"/><path d="M9.2 3v18M14.8 3v18"/><path d="m10.2 9.5 3.6 5M13.8 9.5l-3.6 5" class="r"/>',
    delTbl: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/><path d="M3.5 9.5h17M9.2 4.5v15M14.8 4.5v15"/><path d="m14 13 6 6M20 13l-6 6" class="r"/>',
    hdrRow: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/><path d="M3.5 6a1.5 1.5 0 0 1 1.5-1.5h14A1.5 1.5 0 0 1 20.5 6v3.5h-17z" class="fa"/><path d="M3.5 14.5h17M9.2 9.5v10M14.8 9.5v10"/>',
    tstyle: '<rect x="3.5" y="4.5" width="17" height="15" rx="1.5"/><path d="M3.5 9.5h17M3.5 14.5h17"/><path d="M3.5 9.5h17v5h-17z" class="fa o"/>',
    quote: '<path d="M6 17.5c-1.5-4 .5-8.5 4-10M14 17.5c-1.5-4 .5-8.5 4-10"/><circle cx="7.5" cy="15.5" r="2.2" class="fa"/><circle cx="15.5" cy="15.5" r="2.2" class="fa"/>',
    exit: '<rect x="4" y="4" width="16" height="16" rx="2"/><path d="m9 9 6 6M15 9l-6 6" class="r"/>',
    goto: '<path d="M12 4v13M7 12l5 5 5-5M5 20.5h14" class="a"/>',
    spell: T(1, 12, 9.5, "ABC", 800) + '<path d="m6 16.5 3 3 7-7" class="ok"/>',
    eye: '<path d="M2.5 12s3.5-6.5 9.5-6.5 9.5 6.5 9.5 6.5-3.5 6.5-9.5 6.5S2.5 12 2.5 12z"/><circle cx="12" cy="12" r="2.8"/>',
    file: '<path d="M6 2.5h8l5 5V20a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 5 20V4a1.5 1.5 0 0 1 1-1.5z"/><path d="M14 2.5V8h5"/>',
    reset: '<path d="M4.5 12a7.5 7.5 0 1 0 2.2-5.3"/><path d="M4 3.5v4.5h4.5" class="a"/>',
    print: '<path d="M7 9V3.5h10V9M7 17H4.5V9h15v8H17"/><rect x="7" y="14" width="10" height="7" rx=".8"/>'
  };
  const ic = (n, c) => `<svg class="wd-ic${c ? " " + c : ""}" viewBox="0 0 24 24" aria-hidden="true" focusable="false">${ICONS[n] || ""}</svg>`;

  /* ------------------------------------------------------------------ utilitaires */
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const normS = s => String(s == null ? "" : s).replace(/[’‘]/g, "'").replace(/[\s ]+/g, " ").trim();
  const tipTxt = (t, k) => esc(t + (k ? " (" + k.replace(/Ctrl/g, MOD) + ")" : ""));
  const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
  const reduced = () => { try { return matchMedia("(prefers-reduced-motion: reduce)").matches; } catch (e) { return false; } };
  function rgbToHex(c) {
    if (!c) return null;
    const m = String(c).match(/rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:\s*[,/]\s*([\d.]+%?))?\s*\)/i);
    if (!m) { return /^#[0-9a-f]{6}$/i.test(c) ? c.toLowerCase() : null; }
    let a = m[4] == null ? 1 : (m[4].endsWith("%") ? parseFloat(m[4]) / 100 : parseFloat(m[4]));
    if (a === 0) return null;
    return "#" + [m[1], m[2], m[3]].map(v => clamp(Math.round(+v), 0, 255).toString(16).padStart(2, "0")).join("");
  }
  function mix(hex, t) { // t>0 : vers blanc, t<0 : vers noir
    const n = parseInt(hex.slice(1), 16), c = [n >> 16, (n >> 8) & 255, n & 255];
    const r = c.map(v => Math.round(t > 0 ? v + (255 - v) * t : v * (1 + t)));
    return "#" + r.map(v => v.toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  function unwrap(el) { const p = el.parentNode; if (!p) return; while (el.firstChild) p.insertBefore(el.firstChild, el); p.removeChild(el); }
  function isBlank(el) { return !el.textContent.replace(/​/g, "").trim() && !el.querySelector("img,table,hr,sup.o-fn,.o-toc"); }
  function newP(br) { const p = document.createElement("p"); if (br !== false) p.innerHTML = "<br>"; return p; }
  function closestIn(node, sel, stop) {
    let n = node && node.nodeType === 3 ? node.parentNode : node;
    while (n && n !== stop && n.nodeType === 1) { if (n.matches(sel)) return n; n = n.parentNode; }
    return null;
  }
  function cleanHTML(html) {
    const t = document.createElement("template");
    t.innerHTML = String(html || "");
    t.content.querySelectorAll("script,style,iframe,object,embed,link,meta").forEach(e => e.remove());
    t.content.querySelectorAll("*").forEach(e => {
      [...e.attributes].forEach(a => { if (/^on/i.test(a.name)) e.removeAttribute(a.name); });
      if (e.hasAttribute("href") && /^\s*javascript:/i.test(e.getAttribute("href"))) e.setAttribute("href", "#");
    });
    return t.innerHTML;
  }
  function cleanPaste(html) {
    const t = document.createElement("template");
    t.innerHTML = html;
    const OK = /^(P|H1|H2|H3|UL|OL|LI|B|STRONG|I|EM|U|S|STRIKE|A|TABLE|THEAD|TBODY|TR|TD|TH|BR|BLOCKQUOTE|SUP|SUB)$/;
    const walk = el => {
      [...el.childNodes].forEach(c => {
        if (c.nodeType === 8) { c.remove(); return; }
        if (c.nodeType !== 1) return;
        if (/^(SCRIPT|STYLE|META|LINK|TITLE|IFRAME|OBJECT|SVG|IMG|VIDEO|AUDIO|BUTTON|INPUT|SELECT|TEXTAREA)$/.test(c.tagName)) { c.remove(); return; }
        walk(c);
        if (/^H[4-6]$/.test(c.tagName)) { const h = document.createElement("h3"); while (c.firstChild) h.appendChild(c.firstChild); c.replaceWith(h); return; }
        if (!OK.test(c.tagName)) { unwrap(c); return; }
        const href = c.tagName === "A" ? c.getAttribute("href") : null;
        [...c.attributes].forEach(a => c.removeAttribute(a.name));
        if (href && !/^\s*javascript:/i.test(href)) c.setAttribute("href", href);
      });
    };
    walk(t.content);
    return t.innerHTML;
  }

  /* ================================================================== montage */
  window.OFFICE.sim("doc", function mount(host, setup, ctx) {
    setup = setup || {};
    ctx = ctx || {};
    const N = ++UID;
    const HL = "wd-find-" + N, HLC = "wd-cur-" + N;
    const toast = m => { try { ctx.toast && ctx.toast(m); } catch (e) { /* */ } };
    const st = {
      zoom: 1, userZoom: false, columns: 1, margins: "normal", orient: "portrait", nav: true, ruler: true,
      showCmts: true, marks: false, hf: null, lastRange: null, tab: "home", textColor: "#C00000", hiColor: "#FFFF00",
      hist: [], hidx: -1, histLock: false, findQ: "", findCase: false, matches: [], cur: -1, dead: false,
      pop: null, popBtn: null, modal: null, navSig: "", cmtSig: "", fnSig: "", activeCmt: -1, docName: setup.name || setup.title || "Rapport", saveT: 0
    };
    const listeners = [];
    const on = (t, ev, fn, opt) => { t.addEventListener(ev, fn, opt); listeners.push([t, ev, fn, opt]); };
    const timers = new Set();
    const later = (fn, ms) => { const id = setTimeout(() => { timers.delete(id); if (!st.dead) fn(); }, ms); timers.add(id); return id; };

    /* ---------------------------------------------------------------- gabarit */
    const big = (cmd, n, label, t, k, extra) => `<button type="button" class="wd-btn wd-big" ${cmd} title="${tipTxt(t || label.replace(/<[^>]+>/g, " "), k)}" ${extra || ""}>${ic(n)}<span class="wd-bl">${label}</span></button>`;
    const sm = (cmd, n, label, t, k, extra) => `<button type="button" class="wd-btn wd-sm" ${cmd} title="${tipTxt(t || label, k)}" ${extra || ""}>${ic(n)}${label ? `<span>${label}</span>` : ""}</button>`;
    const ico = (cmd, n, t, k, extra) => `<button type="button" class="wd-btn wd-ico" ${cmd} title="${tipTxt(t, k)}" ${extra || ""}>${ic(n)}</button>`;
    const let_ = (cmd, html, t, k) => `<button type="button" class="wd-btn wd-ico wd-let" data-cmd="${cmd}" title="${tipTxt(t, k)}">${html}</button>`;
    const C_ = c => `data-cmd="${c}"`;
    const A_ = (c, a) => `data-cmd="${c}" data-arg="${a}"`;
    const M_ = m => `data-menu="${m}"`;
    const grp = (label, inner, cls) => `<div class="wd-grp ${cls || ""}"><div class="wd-grp-in">${inner}</div><div class="wd-grp-l">${label}</div></div>`;
    const split = (cmd, n, menu, t, k) => `<span class="wd-split" data-split="${menu}">${ico(C_(cmd), n, t, k)}<button type="button" class="wd-btn wd-arr" ${M_(menu)} title="${esc(t)} : choisir">${ic("chev")}</button></span>`;
    const chevL = l => `${l}<i class="wd-cv">${ic("chev")}</i>`;

    const fontSel = `<select class="wd-sel wd-font" data-cmd="font" title="Police" aria-label="Police">${FONTS.map(f => `<option value="${esc(f[0])}">${esc(f[0])}</option>`).join("")}</select>`;
    const sizeSel = `<select class="wd-sel wd-size" data-cmd="size" title="Taille de police" aria-label="Taille de police">${SIZES.map(s => `<option value="${s}">${String(s).replace(".", ",")}</option>`).join("")}</select>`;
    const gallery = `<div class="wd-gal" role="listbox" aria-label="Styles">${STYLES.map(s => `<button type="button" class="wd-sty" data-style="${s[0]}" title="${tipTxt(s[1], STYLE_KEYS[s[0]])}"><span class="wd-sty-pv wd-sty-${s[0]}">AaBbCcDd</span><span class="wd-sty-n">${s[1]}</span></button>`).join("")}</div>`;

    const PANELS = {
      home:
        grp("Presse-papiers", big(C_("paste"), "paste", "Coller", "Coller", "Ctrl+V") + `<div class="wd-col">${sm(C_("cut"), "cut", "Couper", "Couper", "Ctrl+X")}${sm(C_("copy"), "copy", "Copier", "Copier", "Ctrl+C")}</div>`) +
        grp("Police", `<div class="wd-rows"><div class="wd-row">${fontSel}${sizeSel}${ico(C_("grow"), "grow", "Augmenter la taille de police", "Ctrl+Maj+>")}${ico(C_("shrink"), "shrink", "Réduire la taille de police", "Ctrl+<")}<span class="wd-sep"></span>${ico(C_("clear"), "clear", "Effacer toute la mise en forme", "Ctrl+Espace")}</div>
          <div class="wd-row">${let_("bold", "<b>G</b>", "Gras", "Ctrl+G")}${let_("italic", "<i>I</i>", "Italique", "Ctrl+I")}${let_("underline", "<u>S</u>", "Souligné", "Ctrl+U")}${let_("strike", "<s>abc</s>", "Barré")}${ico(C_("sub"), "sub", "Indice", "Ctrl+=")}${ico(C_("sup"), "sup", "Exposant", "Ctrl+Maj++")}<span class="wd-sep"></span>${split("hilite", "hilite", "hilite", "Couleur de surbrillance du texte")}${split("color", "fcolor", "color", "Couleur de police")}</div></div>`) +
        grp("Paragraphe", `<div class="wd-rows"><div class="wd-row">${ico(C_("ul"), "ul", "Puces", "Ctrl+Maj+L")}${ico(C_("ol"), "ol", "Numérotation")}<span class="wd-sep"></span>${ico(C_("outdent"), "outdent", "Diminuer le retrait")}${ico(C_("indent"), "indent", "Augmenter le retrait")}<span class="wd-sep"></span>${ico(C_("marks"), "pilcrow", "Afficher tout (marques de paragraphe)", "Ctrl+*")}</div>
          <div class="wd-row">${ico(A_("al", "left"), "alL", "Aligner à gauche", "Ctrl+L")}${ico(A_("al", "center"), "alC", "Centrer", "Ctrl+E")}${ico(A_("al", "right"), "alR", "Aligner à droite", "Ctrl+R")}${ico(A_("al", "justify"), "alJ", "Justifier", "Ctrl+J")}</div></div>`) +
        grp("Styles", gallery, "wd-grp-sty") +
        grp("Modification", `<div class="wd-col">${sm(C_("find"), "find", "Rechercher", "Rechercher", "Ctrl+F")}${sm(C_("replace"), "replace", "Remplacer", "Remplacer", "Ctrl+H")}${sm(C_("selall"), "selall", "Sélectionner tout", "Sélectionner tout", "Ctrl+A")}</div>`),
      ins:
        grp("Pages", big(C_("pagebreak"), "pagebreak", "Saut de<br>page", "Saut de page", "Ctrl+Entrée")) +
        grp("Tableaux", big(M_("table"), "table", chevL("Tableau"), "Insérer un tableau")) +
        grp("Liens", big(C_("link"), "link", "Lien", "Insérer un lien hypertexte", "Ctrl+K")) +
        grp("Commentaires", big(C_("comment"), "comment", "Commentaire", "Nouveau commentaire", "Ctrl+Alt+M")) +
        grp("En-tête et pied de page", big(C_("header"), "header", "En-tête", "Modifier l'en-tête") + big(C_("footer"), "footer", "Pied de<br>page", "Modifier le pied de page") + big(C_("pagenum"), "pagenum", "Numéro<br>de page", "Insérer le numéro de page (pied de page)")) +
        grp("Texte", `<div class="wd-col">${sm(C_("date"), "date", "Date et heure", "Insérer la date du jour")}${sm(A_("style", "blockquote"), "quote", "Citation", "Appliquer le style Citation")}</div>`),
      lay:
        grp("Mise en page", big(M_("margins"), "margins", chevL("Marges"), "Marges") + big(M_("orient"), "orient", chevL("Orientation"), "Orientation") + big(M_("cols"), "columns", chevL("Colonnes"), "Colonnes") + big(C_("pagebreak"), "pagebreak", "Sauts", "Insérer un saut de page", "Ctrl+Entrée")) +
        grp("Paragraphe", `<div class="wd-col">${sm(C_("indent"), "indent", "Retrait +", "Augmenter le retrait")}${sm(C_("outdent"), "outdent", "Retrait −", "Diminuer le retrait")}</div>`),
      ref:
        grp("Table des matières", big(C_("toc"), "toc", "Table des<br>matières", "Insérer une table des matières automatique (Titres 1 à 3)") + `<div class="wd-col">${sm(C_("tocUpdate"), "refresh", "Mettre à jour la table", "Mettre à jour la table des matières", "F9")}</div>`) +
        grp("Notes de bas de page", big(C_("footnote"), "footnote", "Insérer une note<br>de bas de page", "Insérer une note de bas de page", "Ctrl+Alt+F")),
      rev:
        grp("Vérification", big(C_("stats"), "stats", "Statistiques", "Statistiques : nombre de mots, caractères…")) +
        grp("Commentaires", big(C_("comment"), "comment", "Nouveau<br>commentaire", "Nouveau commentaire", "Ctrl+Alt+M") + big(C_("cmtDel"), "cdel", "Supprimer", "Supprimer le commentaire actif") + `<div class="wd-col">${sm(C_("cmtPrev"), "prev", "Précédent", "Commentaire précédent")}${sm(C_("cmtNext"), "next", "Suivant", "Commentaire suivant")}${sm(C_("cmtShow"), "eye", "Afficher les commentaires", "Afficher / masquer les commentaires", "", 'data-tog="cmts"')}</div>`),
      view:
        grp("Afficher", `<div class="wd-col">${sm(C_("rulerToggle"), "ruler", "Règle", "Afficher la règle", "", 'data-tog="ruler"')}${sm(C_("navToggle"), "nav", "Volet de navigation", "Volet de navigation", "", 'data-tog="nav"')}${sm(C_("marks"), "pilcrow", "Marques de paragraphe", "Afficher les marques de paragraphe", "", 'data-tog="marks"')}</div>`) +
        grp("Zoom", big(C_("zoomDlg"), "zoom", "Zoom", "Zoom") + big(C_("zoom100"), "z100", "100 %", "Zoom à 100 %") + `<div class="wd-col">${sm(C_("zoomPage"), "onepage", "Une page", "Afficher une page entière")}${sm(C_("zoomWidth"), "pagewidth", "Largeur de page", "Ajuster à la largeur de la page")}</div>`),
      tbl:
        grp("Lignes et colonnes", big(C_("rowA"), "rowA", "Insérer<br>au-dessus", "Insérer une ligne au-dessus") + big(C_("rowB"), "rowB", "Insérer<br>en dessous", "Insérer une ligne en dessous") + `<div class="wd-col">${sm(C_("colL"), "colL", "Insérer à gauche", "Insérer une colonne à gauche")}${sm(C_("colR"), "colR", "Insérer à droite", "Insérer une colonne à droite")}</div>`) +
        grp("Supprimer", `<div class="wd-col">${sm(C_("delRow"), "delRow", "Supprimer la ligne")}${sm(C_("delCol"), "delCol", "Supprimer la colonne")}${sm(C_("delTbl"), "delTbl", "Supprimer le tableau")}</div>`) +
        grp("Options de style", big(C_("hdrRow"), "hdrRow", "Ligne<br>d'en-tête", "Ligne d'en-tête en gras", "", 'data-tog="hdrRow"') + big(M_("tstyle"), "tstyle", chevL("Styles"), "Styles de tableau")),
      hf:
        grp("En-tête et pied de page", big(C_("pagenum"), "pagenum", "Numéro<br>de page", "Insérer le numéro de page") + big(C_("date"), "date", "Date et<br>heure", "Insérer la date du jour")) +
        grp("Navigation", `<div class="wd-col">${sm(C_("gotoHdr"), "header", "Atteindre l'en-tête")}${sm(C_("gotoFtr"), "footer", "Atteindre le pied de page")}</div>`) +
        grp("Fermer", big(C_("hfClose"), "exit", "Fermer l'en-tête<br>et le pied de page", "Fermer l'en-tête et le pied de page", "Échap"))
    };
    const TABS = [["file", "Fichier"], ["home", "Accueil"], ["ins", "Insertion"], ["lay", "Mise en page"], ["ref", "Références"], ["rev", "Révision"], ["view", "Affichage"], ["tbl", "Disposition du tableau", 1], ["hf", "En-tête et pied de page", 1]];

    const app = document.createElement("div");
    app.className = "wd-app";
    app.setAttribute("data-wd", String(N));
    app.innerHTML = `
<style>::highlight(${HL}){background-color:#ffe36e;color:inherit}::highlight(${HLC}){background-color:#ff9f1c;color:#000}</style>
<div class="wd-title">
  <div class="wd-tl">
    <span class="wd-logo" title="Word">${ic("word")}</span>
    <button type="button" class="wd-auto on" data-cmd="autosave" title="Enregistrement automatique (simulé)"><span>Enregistrement automatique</span><i></i></button>
    ${ico(C_("save"), "save", "Enregistrer", "Ctrl+S", 'data-tb="1"')}${ico(C_("undo"), "undo", "Annuler", "Ctrl+Z", 'data-tb="1"')}${ico(C_("redo"), "redo", "Rétablir", "Ctrl+Y", 'data-tb="1"')}
  </div>
  <div class="wd-tc"><span class="wd-docname"></span><span class="wd-saved" aria-live="polite">${ic("check")}<span>Enregistré</span></span></div>
  <div class="wd-tr">
    <button type="button" class="wd-tsearch" data-cmd="find" title="${tipTxt("Rechercher", "Ctrl+F")}">${ic("find")}<span>Rechercher</span></button>
    <button type="button" class="wd-tbtn" data-cmd="cmtShow" title="Afficher / masquer les commentaires">${ic("comment")}<span>Commentaires</span></button>
    <button type="button" class="wd-tbtn wd-share" data-cmd="share" title="Partager">${ic("share")}<span>Partager</span></button>
    <span class="wd-av" title="Vous">V</span>
  </div>
</div>
<div class="wd-tabs" role="tablist">${TABS.map(t => `<button type="button" role="tab" class="wd-tab${t[2] ? " wd-ctx" : ""}" data-tab="${t[0]}"${t[2] ? " hidden" : ""}>${t[1]}</button>`).join("")}<span class="wd-tabbar"></span></div>
<div class="wd-ribbon">${Object.keys(PANELS).map(k => `<div class="wd-panel" data-panel="${k}" role="tabpanel"${k === "home" ? "" : " hidden"}>${PANELS[k]}</div>`).join("")}</div>
<div class="wd-body">
  <aside class="wd-nav" aria-label="Navigation">
    <div class="wd-nav-h"><b>Navigation</b>${ico(C_("navToggle"), "close", "Fermer le volet")}</div>
    <label class="wd-nav-s">${ic("find")}<input type="search" class="wd-nav-q" placeholder="Rechercher dans le document" aria-label="Rechercher dans le document"></label>
    <div class="wd-nav-tabs"><button type="button" class="on" data-ntab="h">Titres</button><button type="button" data-ntab="r">Résultats</button></div>
    <div class="wd-nav-list" data-nlist="h"></div>
    <div class="wd-nav-list" data-nlist="r" hidden></div>
  </aside>
  <div class="wd-canvas">
    <div class="wd-scrollin">
      <div class="wd-rulerbar"><div class="wd-ruler"></div></div>
      <div class="wd-sizer">
        <div class="wd-stage">
          <div class="wd-page">
            <div class="wd-hf wd-hdr" contenteditable="false" spellcheck="false" data-ph="En-tête" title="Double-cliquez pour modifier l'en-tête"></div>
            <div class="wd-root wd-doc" contenteditable="true" spellcheck="true" lang="fr" role="textbox" aria-multiline="true" aria-label="Document"></div>
            <div class="wd-fnotes"></div>
            <div class="wd-hf wd-ftr" contenteditable="false" spellcheck="false" data-ph="Pied de page" title="Double-cliquez pour modifier le pied de page"></div>
          </div>
          <div class="wd-cmts" aria-label="Commentaires"></div>
          <div class="wd-caretfx" aria-hidden="true"></div>
          <div class="wd-flashbox" aria-hidden="true"></div>
        </div>
      </div>
    </div>
  </div>
  <div class="wd-findp" hidden role="dialog" aria-label="Rechercher et remplacer">
    <div class="wd-findp-h"><b>Rechercher et remplacer</b>${ico(C_("findClose"), "close", "Fermer", "Échap")}</div>
    <label class="wd-fld"><span>Rechercher</span><input type="text" class="wd-fq" autocomplete="off"></label>
    <label class="wd-fld"><span>Remplacer par</span><input type="text" class="wd-fr" autocomplete="off"></label>
    <label class="wd-chk"><input type="checkbox" class="wd-fc"> Respecter la casse</label>
    <div class="wd-findp-n" aria-live="polite">Saisissez un texte à rechercher.</div>
    <div class="wd-findp-b">
      <button type="button" class="wd-sbtn" data-cmd="findPrev" title="Précédent">${ic("up")}</button>
      <button type="button" class="wd-sbtn" data-cmd="findNext" title="Suivant (Entrée)">${ic("down")}</button>
      <span class="wd-grow"></span>
      <button type="button" class="wd-sbtn" data-cmd="replOne">Remplacer</button>
      <button type="button" class="wd-pbtn" data-cmd="replAll">Remplacer tout</button>
    </div>
  </div>
</div>
<div class="wd-status">
  <button type="button" class="wd-st-i wd-st-page" data-cmd="stats" title="Statistiques">Page 1 sur 1</button>
  <button type="button" class="wd-st-i wd-st-words" data-cmd="stats" title="Statistiques">0 mot</button>
  <span class="wd-st-i wd-hide-s">${ic("spell")} Français (France)</span>
  <span class="wd-st-i wd-st-mode wd-hide-s" hidden></span>
  <span class="wd-grow"></span>
  <span class="wd-st-i wd-hide-s">${ic("eye")} Focus</span>
  <div class="wd-zoom">
    <button type="button" class="wd-zb" data-cmd="zoomOut" title="Zoom arrière">−</button>
    <input type="range" class="wd-zr" min="50" max="200" step="5" value="100" aria-label="Zoom">
    <button type="button" class="wd-zb" data-cmd="zoomIn" title="Zoom avant">+</button>
    <button type="button" class="wd-zl" data-cmd="zoomDlg" title="Niveau de zoom">100 %</button>
  </div>
</div>
<div class="wd-bs" hidden>
  <div class="wd-bs-side">
    <button type="button" class="wd-bs-back" data-cmd="bsClose" title="Retour">${ic("back")}</button>
    <button type="button" class="wd-bs-i on" data-bs="info">${ic("info")} Informations</button>
    <button type="button" class="wd-bs-i" data-bs="save">${ic("download")} Enregistrer une copie</button>
    <button type="button" class="wd-bs-i" data-bs="print">${ic("print")} Imprimer</button>
    <button type="button" class="wd-bs-i" data-bs="reset">${ic("reset")} Réinitialiser</button>
  </div>
  <div class="wd-bs-main"></div>
</div>`;
    host.appendChild(app);

    const q = s => app.querySelector(s);
    const qa = s => [...app.querySelectorAll(s)];
    const root = q(".wd-root"), hdr = q(".wd-hdr"), ftr = q(".wd-ftr"), page = q(".wd-page"), stage = q(".wd-stage"),
      sizer = q(".wd-sizer"), canvas = q(".wd-canvas"), cmts = q(".wd-cmts"), fnotes = q(".wd-fnotes"), ruler = q(".wd-ruler"),
      nav = q(".wd-nav"), findp = q(".wd-findp"), caretfx = q(".wd-caretfx"), flashbox = q(".wd-flashbox"), bs = q(".wd-bs");
    q(".wd-docname").textContent = st.docName;
    try { document.execCommand("defaultParagraphSeparator", false, "p"); } catch (e) { /* */ }

    /* ---------------------------------------------------------------- éditeur actif & sélection */
    const edActive = () => st.hf ? (st.hf === "ftr" ? ftr : hdr) : root;
    const inEd = n => !!n && (root.contains(n) || hdr.contains(n) || ftr.contains(n));
    function curRange() {
      const s = window.getSelection();
      if (s && s.rangeCount) { const r = s.getRangeAt(0); if (edActive().contains(r.commonAncestorContainer)) return r; }
      return null;
    }
    function restore() {
      const ed = edActive(), s = window.getSelection();
      if (s && s.rangeCount) { const r0 = s.getRangeAt(0); if (ed.contains(r0.commonAncestorContainer) && document.activeElement === ed) return r0; }
      let r = st.lastRange;
      if (!r || !ed.contains(r.startContainer) || !r.startContainer.isConnected) {
        r = document.createRange();
        const fb = ed.querySelector("p,h1,h2,h3,li,td,blockquote") || ed;
        r.selectNodeContents(fb); r.collapse(true);
      }
      try { ed.focus({ preventScroll: true }); } catch (e) { ed.focus(); }
      s.removeAllRanges(); s.addRange(r.cloneRange());
      return s.getRangeAt(0);
    }
    const savePts = () => { const r = curRange(); return r ? { sc: r.startContainer, so: r.startOffset, ec: r.endContainer, eo: r.endOffset } : null; };
    const nlen = n => n.nodeType === 3 ? n.length : n.childNodes.length;
    function restorePts(p, map) {
      if (!p) return;
      const m = n => (map && map.get(n)) || n;
      try {
        const r = document.createRange(), sc = m(p.sc), ec = m(p.ec);
        if (!sc.isConnected || !ec.isConnected) return;
        r.setStart(sc, Math.min(p.so, nlen(sc))); r.setEnd(ec, Math.min(p.eo, nlen(ec)));
        const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      } catch (e) { /* */ }
    }
    function caretIn(el, end) {
      if (!el) return;
      const r = document.createRange();
      r.selectNodeContents(el); r.collapse(!end);
      const s = window.getSelection(); s.removeAllRanges(); s.addRange(r);
      st.lastRange = r.cloneRange();
    }
    function caretAfter(node) {
      const r = document.createRange(); r.setStartAfter(node); r.collapse(true);
      const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); st.lastRange = r.cloneRange();
    }
    function blockOf(n, ed) {
      ed = ed || edActive();
      while (n && n !== ed) {
        if (n.nodeType === 1 && BLOCK_RE.test(n.tagName) && !(n.classList.contains("o-toc"))) return n;
        n = n.parentNode;
      }
      return null;
    }
    function topOf(n, ed) { ed = ed || root; while (n && n.parentNode !== ed) { n = n.parentNode; if (!n) return null; } return n; }
    const isStray = k => (k.nodeType === 3 && /\S/.test(k.data)) || (k.nodeType === 1 && !BLOCKY_RE.test(k.tagName) && !k.hasAttribute("data-wd-ui"));
    function ensureBlocks(ed) {
      const kids = [...ed.childNodes];
      if (!kids.some(isStray)) return;
      const pts = savePts(); let p = null;
      kids.forEach(k => {
        if (isStray(k)) { if (!p) { p = newP(false); k.parentNode.insertBefore(p, k); } p.appendChild(k); }
        else if (k.nodeType === 1 || /\S/.test(k.data)) p = null;
      });
      restorePts(pts);
    }
    function selBlocks() {
      const ed = edActive();
      ensureBlocks(ed);
      const r = curRange() || restore();
      if (!r) return [];
      const out = [], add = b => { if (b && b !== ed && !out.includes(b) && !closestIn(b, ".o-toc", ed)) out.push(b); };
      const startB = blockOf(r.startContainer, ed);
      const skipEnd = (!r.collapsed && r.endOffset === 0) ? blockOf(r.endContainer, ed) : null;
      add(startB);
      if (!r.collapsed) {
        const w = document.createTreeWalker(r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_TEXT);
        let n;
        while ((n = w.nextNode())) {
          if (!r.intersectsNode(n)) continue;
          let b = null;
          if (n.nodeType === 3) { if (/\S/.test(n.data)) b = blockOf(n, ed); }
          else if (BLOCK_RE.test(n.tagName) && !n.querySelector(BLOCK_SEL)) b = n;
          if (b && b !== skipEnd) add(b);
        }
      }
      if (!out.length && r.collapsed && r.startContainer === ed) {
        const k = ed.childNodes[r.startOffset] || ed.childNodes[r.startOffset - 1];
        if (k && k.nodeType === 1 && BLOCK_RE.test(k.tagName)) add(k);
      }
      return out;
    }

    /* ---------------------------------------------------------------- styles de paragraphe */
    function setTag(b, tag, map) {
      tag = tag.toUpperCase();
      if (b.tagName === tag) return b;
      if (/^(LI|TD|TH)$/.test(b.tagName)) {
        if (tag === "P") return b;
        const n = document.createElement(tag);
        [...b.childNodes].forEach(c => { if (!/^(UL|OL)$/.test(c.nodeName)) n.appendChild(c); });
        b.insertBefore(n, b.firstChild);
        map.set(b, n);
        return n;
      }
      let p = b.parentElement;
      if (p && /^(LI|TD|TH)$/.test(p.tagName) && tag === "P") { const first = b.firstChild; unwrap(b); map.set(b, first && first.parentNode ? first.parentNode : p); return p; }
      if (p && p.tagName === "BLOCKQUOTE" && p !== edActive() && edActive().contains(p)) {
        if (tag === "BLOCKQUOTE") return b;
        if (p.children.length === 1) unwrap(p); else p.after(b);
      }
      const n = document.createElement(tag);
      const sty = b.getAttribute("style");
      if (sty) n.setAttribute("style", sty);
      while (b.firstChild) n.appendChild(b.firstChild);
      if (!n.firstChild) n.innerHTML = "<br>";
      b.replaceWith(n);
      map.set(b, n);
      return n;
    }
    function applyStyle(tag) {
      clearPreview();
      restore();
      const pts = savePts(), map = new Map();
      const blocks = selBlocks();
      if (!blocks.length) return;
      blocks.forEach(b => setTag(b, tag, map));
      restorePts(pts, map);
      after();
    }
    let pvBlocks = [];
    function preview(tag) {
      clearPreview();
      const r = st.lastRange;
      if (!r || !edActive().contains(r.startContainer)) return;
      const s = window.getSelection();
      const keep = s.rangeCount ? s.getRangeAt(0).cloneRange() : null;
      s.removeAllRanges(); s.addRange(r.cloneRange());
      pvBlocks = selBlocks().map(b => (/^(LI|TD|TH)$/.test(b.tagName) && b.firstElementChild && BLOCK_RE.test(b.firstElementChild.tagName)) ? b.firstElementChild : b);
      if (keep) { s.removeAllRanges(); s.addRange(keep); }
      pvBlocks.forEach(b => b.classList.add("wd-pv", "wd-pv-" + tag));
    }
    function clearPreview() {
      pvBlocks.forEach(b => { b.classList.remove("wd-pv", "wd-pv-p", "wd-pv-h1", "wd-pv-h2", "wd-pv-h3", "wd-pv-blockquote"); if (!b.getAttribute("class")) b.removeAttribute("class"); });
      pvBlocks = [];
    }

    /* ---------------------------------------------------------------- commandes de caractère */
    function exec(cmd, val, css) {
      restore();
      try { document.execCommand("styleWithCSS", false, !!css); } catch (e) { /* */ }
      let ok = false;
      try { ok = document.execCommand(cmd, false, val == null ? null : val); } catch (e) { ok = false; }
      after();
      return ok;
    }
    function setColor(c) { st.textColor = c; app.style.setProperty("--wd-fc", c); exec("foreColor", c, true); }
    function setHilite(c) {
      if (c) { st.hiColor = c; app.style.setProperty("--wd-hc", c); }
      restore();
      try { document.execCommand("styleWithCSS", false, true); } catch (e) { /* */ }
      const v = c || "transparent";
      let ok = false;
      try { ok = document.execCommand("hiliteColor", false, v); } catch (e) { ok = false; }
      if (!ok) { try { document.execCommand("backColor", false, v); } catch (e) { /* */ } }
      after();
    }
    function fontSizePt(pt) {
      restore();
      const r = curRange();
      if (!r) return;
      try { document.execCommand("styleWithCSS", false, false); document.execCommand("fontSize", false, "7"); } catch (e) { return; }
      edActive().querySelectorAll('font[size="7"]').forEach(f => {
        const sp = document.createElement("span"); sp.style.fontSize = pt + "pt";
        if (f.getAttribute("face")) sp.style.fontFamily = f.getAttribute("face");
        if (f.getAttribute("color")) sp.style.color = f.getAttribute("color");
        while (f.firstChild) sp.appendChild(f.firstChild);
        sp.querySelectorAll("span").forEach(x => { if (x.style.fontSize) x.style.fontSize = ""; });
        f.replaceWith(sp);
        const rr = document.createRange(); rr.selectNodeContents(sp);
        const s = window.getSelection(); s.removeAllRanges(); s.addRange(rr);
      });
      after();
    }
    function curPt() {
      const r = curRange() || st.lastRange;
      if (!r) return 12;
      let n = r.startContainer; if (n.nodeType === 3) n = n.parentNode;
      if (n.nodeType !== 1) return 12;
      return Math.round(parseFloat(getComputedStyle(n).fontSize) * 0.75 * 2) / 2;
    }
    function growFont(dir) {
      const pt = curPt();
      let nx = dir > 0 ? SIZES.find(s => s > pt) : [...SIZES].reverse().find(s => s < pt);
      if (!nx) nx = dir > 0 ? 72 : 8;
      fontSizePt(nx);
    }
    function setFont(name) {
      const f = FONTS.find(x => x[0] === name);
      exec("fontName", f ? f[1] : name, true);
    }
    function clearFormatting(charsOnly) {
      const ed = edActive();
      let r = restore();
      if (r.collapsed) {
        const b = blockOf(r.startContainer, ed);
        if (b) { const rr = document.createRange(); rr.selectNodeContents(b); const s = window.getSelection(); s.removeAllRanges(); s.addRange(rr); r = rr; }
      }
      // protège notes de bas de page et numéros de page (removeFormat supprime les <sup>)
      const prot = [];
      ed.querySelectorAll("sup.o-fn, span.wd-pgnum").forEach(x => { const a = document.createElement("a"); a.setAttribute("data-wd-prot", ""); x.replaceWith(a); a.appendChild(x); prot.push(a); });
      try { document.execCommand("styleWithCSS", false, false); document.execCommand("removeFormat", false, null); } catch (e) { /* */ }
      prot.forEach(a => { if (a.isConnected) unwrap(a); });
      const r2 = curRange();
      if (r2) {
        const nr = document.createRange();
        ed.querySelectorAll("span,font,b,strong,i,em,u,s,strike,mark:not(.o-comment)").forEach(x => {
          nr.selectNodeContents(x);
          if (r2.compareBoundaryPoints(Range.START_TO_START, nr) <= 0 && r2.compareBoundaryPoints(Range.END_TO_END, nr) >= 0 && !x.closest(".o-toc")) {
            if (x.matches("span.wd-pgnum")) return;
            unwrap(x);
          }
        });
      }
      if (!charsOnly) {
        const pts = savePts(), map = new Map();
        selBlocks().forEach(b => {
          const n = setTag(b, "p", map);
          if (n && n.style) { n.style.textAlign = ""; n.style.marginLeft = ""; if (!n.getAttribute("style")) n.removeAttribute("style"); }
        });
        restorePts(pts, map);
      }
      after();
    }
    function align(a) {
      restore();
      selBlocks().forEach(b => {
        b.style.textAlign = (a === "left" && b.tagName !== "BLOCKQUOTE") ? "" : a;
        if (!b.getAttribute("style")) b.removeAttribute("style");
      });
      after();
    }
    function indent(dir) {
      const r = restore();
      if (closestIn(r.startContainer, "li", edActive())) { exec(dir > 0 ? "indent" : "outdent"); return; }
      selBlocks().forEach(b => {
        if (/^(TD|TH)$/.test(b.tagName)) return;
        const cur = parseFloat(b.style.marginLeft) || 0; // cm
        const nv = Math.max(0, Math.round((cur + dir * 1.25) * 100) / 100);
        b.style.marginLeft = nv ? nv + "cm" : "";
        if (!b.getAttribute("style")) b.removeAttribute("style");
      });
      after();
    }
    function list(kind) {
      const r = restore();
      const li = closestIn(r.startContainer, "li", edActive());
      // un titre ne doit pas finir dans la puce : on garde le style si le bloc est un titre
      exec(kind === "ol" ? "insertOrderedList" : "insertUnorderedList");
      if (!li) { // Chrome peut imbriquer la liste dans un <p>
        root.querySelectorAll("p > ul, p > ol").forEach(l => { const p = l.parentElement; if (p.children.length === 1 && isBlank(p) === false && p.textContent === l.textContent) { p.replaceWith(l); } });
      }
    }

    /* ---------------------------------------------------------------- insertions de blocs */
    function insertBlock(node, ensureAfter) {
      const r = restore();
      const top = topOf(r.startContainer, root) || (r.startContainer === root ? root.childNodes[r.startOffset] : null);
      if (!top || top.nodeType !== 1) root.appendChild(node);
      else if (/^(P|H[1-6]|BLOCKQUOTE|PRE|DIV)$/.test(top.tagName) && !top.classList.contains("o-toc")) {
        if (!r.collapsed) r.deleteContents();
        const pre = document.createRange(); pre.setStart(top, 0); pre.setEnd(r.startContainer, r.startOffset);
        const post = document.createRange(); post.setStart(r.startContainer, r.startOffset); post.setEnd(top, top.childNodes.length);
        if (isBlank(top)) top.before(node);
        else if (!pre.toString().trim()) top.before(node);
        else if (!post.toString().trim()) top.after(node);
        else {
          const frag = post.extractContents();
          const nb = document.createElement(top.tagName);
          if (top.getAttribute("style")) nb.setAttribute("style", top.getAttribute("style"));
          nb.appendChild(frag);
          top.after(nb); top.after(node);
        }
      } else top.after(node);
      if (ensureAfter) {
        const nx = node.nextElementSibling;
        if (!nx || !/^(P|H[1-6]|BLOCKQUOTE|UL|OL)$/.test(nx.tagName)) node.after(newP());
      }
      return node;
    }
    function insertAtCaret(node) {
      const r = restore();
      r.deleteContents();
      r.insertNode(node);
      caretAfter(node);
      return node;
    }
    function insertPageBreak() {
      if (st.hf) return toast("Fermez l'en-tête pour insérer un saut de page.");
      const hr = document.createElement("hr"); hr.className = "o-pagebreak";
      insertBlock(hr, true);
      caretIn(hr.nextElementSibling);
      after();
    }
    function insertTable(rows, cols) {
      if (st.hf) return toast("Fermez l'en-tête pour insérer un tableau.");
      rows = clamp(rows | 0, 1, 50); cols = clamp(cols | 0, 1, 12);
      const t = document.createElement("table"); t.className = "wd-tbl";
      const tb = t.createTBody();
      for (let i = 0; i < rows; i++) { const tr = tb.insertRow(); for (let j = 0; j < cols; j++) tr.insertCell().innerHTML = "<br>"; }
      insertBlock(t, true);
      caretIn(t.rows[0].cells[0]);
      after();
      if (!reduced()) { t.classList.add("wd-born"); later(() => { t.classList.remove("wd-born"); if (!t.getAttribute("class")) t.removeAttribute("class"); }, 700); }
    }

    /* ---------------------------------------------------------------- tableaux */
    function cellCtx() {
      const r = curRange() || st.lastRange;
      if (!r) return null;
      const td = closestIn(r.startContainer, "td,th", root);
      if (!td) return null;
      return { td, tr: td.parentElement, table: td.closest("table") };
    }
    function cellBold(c) {
      const w = document.createTreeWalker(c, NodeFilter.SHOW_TEXT);
      let n, any = false;
      while ((n = w.nextNode())) {
        if (!/\S/.test(n.data)) continue;
        any = true;
        if ((parseInt(getComputedStyle(n.parentElement).fontWeight, 10) || 400) < 600) return false;
      }
      return any ? true : null;
    }
    function rowBold(tr) {
      const v = [...tr.cells].map(cellBold).filter(x => x !== null);
      return v.length > 0 && v.every(Boolean);
    }
    function delTable(t) {
      const nx = t.nextElementSibling;
      if (nx && nx.tagName === "P") { t.remove(); caretIn(nx); }
      else { const p = newP(); t.replaceWith(p); caretIn(p); }
    }
    function tblOp(op) {
      const c = cellCtx();
      if (!c) return toast("Placez d'abord le curseur dans un tableau.");
      const { td, tr, table } = c;
      const ci = td.cellIndex;
      const mk = tag => { const x = document.createElement(tag || "td"); x.innerHTML = "<br>"; return x; };
      if (op === "rowA" || op === "rowB") {
        const n = document.createElement("tr");
        for (let i = 0; i < tr.cells.length; i++) n.appendChild(mk());
        if (op === "rowA") tr.before(n); else tr.after(n);
        caretIn(n.cells[Math.min(ci, n.cells.length - 1)]);
      } else if (op === "colL" || op === "colR") {
        [...table.rows].forEach(row => {
          const ref = row.cells[Math.min(ci, row.cells.length - 1)];
          const n = mk(ref && ref.tagName === "TH" ? "th" : "td");
          if (!ref) row.appendChild(n); else if (op === "colL") ref.before(n); else ref.after(n);
        });
      } else if (op === "delRow") {
        const nx = tr.nextElementSibling || tr.previousElementSibling;
        tr.remove();
        if (!table.rows.length) delTable(table); else if (nx) caretIn(nx.cells[0]);
      } else if (op === "delCol") {
        [...table.rows].forEach(row => { if (row.cells[ci]) row.cells[ci].remove(); if (!row.cells.length) row.remove(); });
        if (!table.rows.length) delTable(table); else caretIn(table.rows[0].cells[Math.max(0, ci - 1)]);
      } else if (op === "delTbl") delTable(table);
      else if (op === "hdrRow") {
        const r0 = table.rows[0], was = rowBold(r0);
        [...r0.cells].forEach(cell => {
          if (was) { cell.querySelectorAll("b,strong").forEach(unwrap); cell.style.fontWeight = "normal"; }
          else cell.style.fontWeight = "bold";
        });
      }
      after();
    }
    function tableStyle(k) {
      const c = cellCtx();
      if (!c) return toast("Placez d'abord le curseur dans un tableau.");
      c.table.classList.remove("wd-ts-grid", "wd-ts-blue", "wd-ts-zebra", "wd-ts-orange");
      if (k !== "grid") c.table.classList.add("wd-ts-" + k);
      after();
    }

    /* ---------------------------------------------------------------- liens, commentaires, notes */
    function clampToBlock(r) {
      const b = blockOf(r.startContainer);
      if (b && !b.contains(r.endContainer)) r.setEnd(b, b.childNodes.length);
      return r;
    }
    function wordRange(r) {
      const n = r.startContainer;
      if (n.nodeType !== 3) return r;
      const d = n.data; let a = r.startOffset, b = r.startOffset;
      while (a > 0 && /[\wÀ-ÿ'’-]/.test(d[a - 1])) a--;
      while (b < d.length && /[\wÀ-ÿ'’-]/.test(d[b])) b++;
      const rr = document.createRange(); rr.setStart(n, a); rr.setEnd(n, b);
      return rr;
    }
    async function addLink() {
      const r0 = restore().cloneRange();
      const ex = closestIn(r0.startContainer, "a", edActive());
      const txt = ex ? ex.textContent : r0.toString();
      const v = await ask("Insérer un lien hypertexte", [
        { name: "text", label: "Texte à afficher", value: txt, ph: "Texte du lien" },
        { name: "url", label: "Adresse", value: ex ? ex.getAttribute("href") : "https://", ph: "https://www.exemple.fr" }
      ], "OK");
      if (!v) return refocus(r0);
      let u = (v.url || "").trim();
      if (!u || u === "https://") return refocus(r0);
      if (/^\s*javascript:/i.test(u)) u = "#";
      else if (!/^(https?:|mailto:|tel:|#|\/|\.)/i.test(u)) u = (/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/i.test(u) ? "mailto:" : "https://") + u;
      const label = (v.text || "").trim();
      if (ex) { ex.setAttribute("href", u); if (label && label !== ex.textContent) ex.textContent = label; refocus(); caretAfter(ex); }
      else {
        const a = document.createElement("a"); a.setAttribute("href", u); a.title = "Ctrl+clic pour suivre le lien";
        const r = clampToBlock(r0);
        if (!r.collapsed) { a.appendChild(r.extractContents()); if (label && label !== txt) a.textContent = label; }
        else a.textContent = label || u;
        r.insertNode(a);
        refocus(); caretAfter(a);
      }
      after();
    }
    async function addComment() {
      if (st.hf) return toast("Les commentaires s'ajoutent dans le corps du document.");
      let r = restore().cloneRange();
      if (r.collapsed) r = wordRange(r);
      if (r.collapsed || !r.toString().trim()) { toast("Sélectionnez d'abord le texte à commenter."); return; }
      clampToBlock(r);
      const v = await ask("Nouveau commentaire", [{ name: "note", label: "Commentaire", area: true, ph: "Votre remarque…" }], "Publier");
      if (!v || !v.note.trim()) return refocus(r);
      const m = document.createElement("mark");
      m.className = "o-comment";
      m.setAttribute("data-note", v.note.trim());
      m.setAttribute("data-author", "Vous");
      m.setAttribute("data-date", new Date().toISOString());
      m.appendChild(r.extractContents());
      r.insertNode(m);
      refocus(); caretAfter(m);
      st.showCmts = true;
      after();
      later(() => { const marks = commentMarks(); activateCmt(marks.indexOf(m)); }, 200);
    }
    async function addFootnote() {
      if (st.hf) return toast("Fermez l'en-tête pour insérer une note.");
      const r0 = restore().cloneRange(); r0.collapse(false);
      const v = await ask("Note de bas de page", [{ name: "note", label: "Texte de la note", area: true, ph: "Source, précision…" }], "Insérer");
      if (!v || !v.note.trim()) return refocus(r0);
      const s = document.createElement("sup");
      s.className = "o-fn"; s.contentEditable = "false"; s.setAttribute("data-note", v.note.trim()); s.textContent = "#";
      refocus(r0);
      r0.insertNode(s); caretAfter(s);
      renumberFn();
      after();
    }
    function refocus(r) {
      const ed = edActive();
      try { ed.focus({ preventScroll: true }); } catch (e) { ed.focus(); }
      if (r) { const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); st.lastRange = r.cloneRange(); }
    }
    const commentMarks = () => [...root.querySelectorAll("mark.o-comment")];
    function delComment(i) {
      const ms = commentMarks();
      const m = ms[i != null ? i : st.activeCmt];
      if (!m) return toast("Aucun commentaire sélectionné.");
      unwrap(m); st.activeCmt = -1; st.cmtSig = ""; after();
    }
    function activateCmt(i) {
      const ms = commentMarks();
      ms.forEach((m, k) => m.classList.toggle("wd-cmt-on", k === i));
      qa(".wd-card").forEach(c => c.classList.toggle("on", +c.dataset.i === i));
      st.activeCmt = i;
      positionCards();
    }
    function cmtStep(d) {
      const ms = commentMarks();
      if (!ms.length) return toast("Ce document ne contient aucun commentaire.");
      const i = ((st.activeCmt < 0 ? (d > 0 ? -1 : 0) : st.activeCmt) + d + ms.length) % ms.length;
      activateCmt(i); scrollToEl(ms[i]);
    }

    /* ---------------------------------------------------------------- en-tête / pied de page */
    function hfEdit(which) {
      st.hf = which || null;
      page.classList.toggle("wd-hfmode", !!which);
      root.contentEditable = which ? "false" : "true";
      hdr.contentEditable = ftr.contentEditable = which ? "true" : "false";
      const tab = q('.wd-tab[data-tab="hf"]');
      tab.hidden = !which;
      if (which) {
        const z = which === "ftr" ? ftr : hdr;
        if (!z.textContent.trim() && !z.querySelector(".wd-pgnum")) z.innerHTML = "";
        try { z.focus({ preventScroll: true }); } catch (e) { z.focus(); }
        caretIn(z, true);
        setTab("hf");
        scrollToEl(z, true);
      } else {
        if (st.tab === "hf") setTab("home");
        try { root.focus({ preventScroll: true }); } catch (e) { /* */ }
        st.lastRange = null;
        restore();
      }
      updateTabs(); changed();
    }
    function insertPageNum() {
      if (!st.hf) hfEdit("ftr");
      const s = document.createElement("span");
      s.className = "wd-pgnum"; s.contentEditable = "false"; s.setAttribute("data-field", "PAGE"); s.textContent = "1";
      const z = st.hf === "hdr" ? hdr : ftr;
      if (!z.textContent.trim()) { z.innerHTML = '<p style="text-align:center"></p>'; z.firstChild.appendChild(s); caretAfter(s); }
      else insertAtCaret(s);
      after();
    }
    function insertDate() {
      const d = new Date().toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" });
      insertAtCaret(document.createTextNode(d));
      after();
    }

    /* ---------------------------------------------------------------- table des matières */
    const headingEls = () => [...root.querySelectorAll("h1,h2,h3")].filter(h => !h.closest(".o-toc"));
    const tocSig = () => headingEls().map(h => h.tagName[1] + "|" + normS(h.textContent)).join("\n");
    function fillToc(toc) {
      const hs = headingEls();
      toc.contentEditable = "false";
      toc.setAttribute("data-sig", tocSig());
      toc.innerHTML = `<div class="wd-toc-bar" data-wd-ui="1"><button type="button" class="wd-toc-up" data-cmd="tocUpdate" title="Mettre à jour la table (F9)">${ic("refresh")}<span>Mettre à jour la table…</span></button></div>` +
        `<p class="o-toc-title">Table des matières</p>` +
        (hs.length ? hs.map((h, i) => `<p class="o-toc-e o-toc-${h.tagName[1]}"><a href="#" data-h="${i}" title="Clic pour atteindre le titre"><span class="o-toc-t">${esc(normS(h.textContent) || "(titre vide)")}</span><span class="o-toc-d"></span><span class="o-toc-p">${pageOf(h)}</span></a></p>`).join("")
          : `<p class="o-toc-e">Aucune entrée de table des matières n'a été trouvée.</p>`);
    }
    function refreshTocPages(toc) {
      const hs = headingEls();
      toc.querySelectorAll("a[data-h]").forEach(a => { const h = hs[+a.dataset.h]; const p = a.querySelector(".o-toc-p"); if (h && p) p.textContent = pageOf(h); });
    }
    function insertToc() {
      if (st.hf) return toast("Fermez l'en-tête pour insérer une table des matières.");
      let toc = root.querySelector(".o-toc");
      if (toc) { updateToc(); toast("La table des matières existait déjà : elle a été mise à jour."); return; }
      toc = document.createElement("div"); toc.className = "o-toc";
      fillToc(toc);
      insertBlock(toc, true);
      refreshTocPages(toc);
      caretIn(toc.nextElementSibling);
      if (!headingEls().length) toast("Aucun titre trouvé : appliquez les styles Titre 1/2/3 puis mettez la table à jour.");
      after();
    }
    function updateToc() {
      const toc = root.querySelector(".o-toc");
      if (!toc) return toast("Insérez d'abord une table des matières (Références › Table des matières).");
      fillToc(toc); refreshTocPages(toc);
      toc.classList.remove("wd-stale");
      if (!reduced()) { toc.classList.add("wd-flash"); later(() => toc.classList.remove("wd-flash"), 900); }
      toast("Table des matières mise à jour.");
      after();
    }

    /* ---------------------------------------------------------------- pages (estimation) */
    const PAGE = () => {
      const pw = st.orient === "landscape" ? 1123 : 794, ph = st.orient === "landscape" ? 794 : 1123;
      const m = MARGINS[st.margins] || MARGINS.normal;
      return { pw, ph, mx: m.x * CM, my: m.y * CM, contentH: ph - 2 * m.y * CM };
    };
    const scale = () => { const w = stage.offsetWidth; return w ? (stage.getBoundingClientRect().width / w) || 1 : 1; };
    function pageAtY(y) {
      const g = PAGE(), s = scale(), rt = root.getBoundingClientRect().top;
      let p = 1, seg = 0;
      for (const hr of root.querySelectorAll("hr.o-pagebreak")) {
        const b = hr.getBoundingClientRect();
        const top = (b.top - rt) / s;
        if (top > y) break;
        p += 1 + Math.floor(Math.max(0, top - seg - 1) / g.contentH);
        seg = (b.bottom - rt) / s;
      }
      return p + Math.floor(Math.max(0, y - seg - 1) / g.contentH);
    }
    function pageOf(el) {
      if (!el || !el.getBoundingClientRect || !root.contains(el)) return 1;
      return pageAtY((el.getBoundingClientRect().top - root.getBoundingClientRect().top) / scale());
    }
    const pageCount = () => pageAtY(root.scrollHeight - 2);

    /* ---------------------------------------------------------------- recherche */
    function buildIndex(ed, normalize, skipToc) {
      let str = "", prevSpace = true, lastB;
      const map = [];
      const w = document.createTreeWalker(ed, NodeFilter.SHOW_TEXT, {
        acceptNode: n => (n.parentElement && (n.parentElement.closest("[data-wd-ui]") || (skipToc && n.parentElement.closest(".o-toc")))) ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT
      });
      let n;
      while ((n = w.nextNode())) {
        const b = blockOf(n, ed);
        if (str.length && b !== lastB) {
          if (normalize) { if (!prevSpace) { str += " "; map.push(null); prevSpace = true; } }
          else { str += "\n"; map.push(null); }
        }
        lastB = b;
        const d = n.data;
        for (let i = 0; i < d.length; i++) {
          let c = d[i];
          if (c === " ") c = " ";
          else if (c === "’" || c === "‘") c = "'";
          if (normalize && /\s/.test(c)) { if (prevSpace) continue; c = " "; prevSpace = true; }
          else prevSpace = false;
          str += c; map.push([n, i]);
        }
      }
      return { str, map };
    }
    function findRanges(qs, cs) {
      const out = [];
      qs = String(qs || "").replace(/ /g, " ").replace(/[’‘]/g, "'");
      if (!qs) return out;
      const idx = buildIndex(root, false, true);
      const hay = cs ? idx.str : idx.str.toLowerCase(), nee = cs ? qs : qs.toLowerCase();
      let i = hay.indexOf(nee);
      while (i >= 0 && out.length < 2000) {
        const a = idx.map[i], b = idx.map[i + nee.length - 1];
        let ok = a && b;
        for (let k = i; ok && k < i + nee.length; k++) if (!idx.map[k]) ok = false;
        if (ok) { const r = document.createRange(); r.setStart(a[0], a[1]); r.setEnd(b[0], b[1] + 1); out.push(r); }
        i = hay.indexOf(nee, i + Math.max(1, nee.length));
      }
      return out;
    }
    const HAS_HL = !!(window.CSS && CSS.highlights && window.Highlight);
    function paintMatches() {
      if (!HAS_HL) return;
      try {
        if (st.matches.length) CSS.highlights.set(HL, new Highlight(...st.matches)); else CSS.highlights.delete(HL);
        const c = st.matches[st.cur];
        if (c) CSS.highlights.set(HLC, new Highlight(c)); else CSS.highlights.delete(HLC);
      } catch (e) { /* */ }
    }
    function clearMatches() { st.matches = []; st.cur = -1; if (HAS_HL) { try { CSS.highlights.delete(HL); CSS.highlights.delete(HLC); } catch (e) { /* */ } } }
    function runFind(qs, keepCur) {
      st.findQ = qs;
      st.matches = findRanges(qs, st.findCase);
      if (!keepCur || st.cur >= st.matches.length) st.cur = st.matches.length ? 0 : -1;
      paintMatches();
      const n = st.matches.length;
      q(".wd-findp-n").textContent = !qs ? "Saisissez un texte à rechercher." : n ? `${st.cur + 1} sur ${n} résultat${n > 1 ? "s" : ""}` : "Aucun résultat.";
      q(".wd-findp-n").classList.toggle("wd-bad", !!qs && !n);
      return n;
    }
    function gotoMatch(d) {
      if (!st.matches.length) runFind(q(".wd-fq").value);
      if (!st.matches.length) return;
      st.cur = (st.cur + d + st.matches.length) % st.matches.length;
      paintMatches(); runFind(st.findQ, true);
      const r = st.matches[st.cur];
      st.lastRange = r.cloneRange();
      scrollToRange(r);
    }
    function replaceRange(r, rep) {
      if (r.startContainer === r.endContainer && r.startContainer.nodeType === 3) r.startContainer.replaceData(r.startOffset, r.endOffset - r.startOffset, rep);
      else { r.deleteContents(); if (rep) r.insertNode(document.createTextNode(rep)); }
    }
    function replOne() {
      if (!st.matches.length) runFind(q(".wd-fq").value);
      const r = st.matches[st.cur];
      if (!r) return;
      replaceRange(r, q(".wd-fr").value);
      root.normalize();
      runFind(st.findQ, true);
      if (st.matches.length) gotoMatch(0);
      after();
    }
    function replAll() {
      const qs = q(".wd-fq").value;
      const rs = findRanges(qs, st.findCase);
      if (!qs) return;
      const rep = q(".wd-fr").value;
      for (let i = rs.length - 1; i >= 0; i--) replaceRange(rs[i], rep);
      root.normalize();
      runFind(qs);
      toast(rs.length ? `Nous avons terminé. Nous avons effectué ${rs.length} remplacement${rs.length > 1 ? "s" : ""}.` : "Aucune occurrence trouvée.");
      after();
    }
    function openFind(replace) {
      findp.hidden = false;
      requestAnimationFrame(() => findp.classList.add("wd-in"));
      const fq = q(".wd-fq");
      const r = curRange();
      if (r && !r.collapsed && r.toString().length < 80) fq.value = r.toString();
      (replace && fq.value ? q(".wd-fr") : fq).focus();
      fq.select && !replace && fq.select();
      runFind(fq.value);
    }
    function closeFind() { findp.classList.remove("wd-in"); findp.hidden = true; clearMatches(); restore(); }
    function navSearch(qs) {
      const box = q('[data-nlist="r"]');
      setNavTab(qs ? "r" : "h");
      st.matches = findRanges(qs, false); st.cur = st.matches.length ? 0 : -1;
      paintMatches();
      if (!qs) { box.innerHTML = ""; return; }
      if (!st.matches.length) { box.innerHTML = `<div class="wd-nav-empty">Aucun résultat pour « ${esc(qs)} ».</div>`; return; }
      box.innerHTML = `<div class="wd-nav-cnt">${st.matches.length} résultat${st.matches.length > 1 ? "s" : ""}</div>` + st.matches.slice(0, 60).map((r, i) => {
        const t = r.startContainer.data || "";
        const a = Math.max(0, r.startOffset - 28), b = r.startContainer === r.endContainer ? r.endOffset : t.length;
        return `<button type="button" class="wd-nav-r" data-r="${i}" style="--d:${Math.min(i, 12)}">${a ? "…" : ""}${esc(t.slice(a, r.startOffset))}<mark>${esc(t.slice(r.startOffset, b))}</mark>${esc(t.slice(b, b + 40))}…</button>`;
      }).join("");
    }

    /* ---------------------------------------------------------------- défilement / effets */
    function scrollToEl(el, instant) {
      if (!el) return;
      const cr = canvas.getBoundingClientRect(), er = el.getBoundingClientRect();
      const top = er.top - cr.top + canvas.scrollTop - 90;
      canvas.scrollTo({ top: Math.max(0, top), behavior: (instant || reduced()) ? "auto" : "smooth" });
      flash(er);
    }
    function scrollToRange(r) {
      const rects = r.getClientRects(); const er = rects[0] || r.getBoundingClientRect();
      const cr = canvas.getBoundingClientRect();
      if (er.top < cr.top + 40 || er.bottom > cr.bottom - 40) canvas.scrollTo({ top: Math.max(0, er.top - cr.top + canvas.scrollTop - cr.height / 3), behavior: reduced() ? "auto" : "smooth" });
    }
    function flash(er) {
      if (reduced()) return;
      const sr = stage.getBoundingClientRect(), s = scale();
      flashbox.style.cssText = `top:${(er.top - sr.top) / s - 4}px;left:${(er.left - sr.left) / s - 6}px;width:${er.width / s + 12}px;height:${er.height / s + 8}px`;
      flashbox.classList.remove("on"); void flashbox.offsetWidth; flashbox.classList.add("on");
    }
    function caretPulse() {
      if (reduced()) return;
      const r = curRange();
      if (!r || !r.collapsed) return;
      let rect = r.getClientRects()[0];
      if (!rect) { const b = blockOf(r.startContainer); if (!b) return; rect = b.getBoundingClientRect(); }
      if (!rect || (!rect.height && !rect.top)) return;
      const sr = stage.getBoundingClientRect(), s = scale();
      caretfx.style.cssText = `top:${(rect.top - sr.top) / s}px;left:${(rect.left - sr.left) / s}px;height:${rect.height / s}px`;
      caretfx.classList.remove("on"); void caretfx.offsetWidth; caretfx.classList.add("on");
    }
    function ripple(btn, e) {
      if (reduced() || !btn) return;
      const r = btn.getBoundingClientRect(), sz = Math.max(r.width, r.height) * 1.6;
      const sp = document.createElement("span");
      sp.className = "wd-ripple";
      sp.style.cssText = `width:${sz}px;height:${sz}px;left:${(e.clientX || r.left + r.width / 2) - r.left - sz / 2}px;top:${(e.clientY || r.top + r.height / 2) - r.top - sz / 2}px`;
      btn.appendChild(sp);
      later(() => sp.remove(), 600);
    }

    /* ---------------------------------------------------------------- page, zoom, règle, colonnes */
    function applyPage() {
      const g = PAGE();
      page.style.setProperty("--wd-pw", g.pw + "px");
      page.style.setProperty("--wd-ph", g.ph + "px");
      page.style.setProperty("--wd-mx", g.mx.toFixed(1) + "px");
      page.style.setProperty("--wd-my", g.my.toFixed(1) + "px");
      root.style.columnCount = st.columns > 1 ? String(st.columns) : "";
      root.style.columnGap = st.columns > 1 ? "1.25cm" : "";
      if (!root.getAttribute("style")) root.removeAttribute("style");
      page.classList.toggle("wd-cols", st.columns > 1);
      drawRuler(); sizeStage();
    }
    function sizeStage() {
      const z = st.zoom;
      stage.style.transform = z === 1 ? "" : `scale(${z})`;
      sizer.style.width = Math.ceil(stage.offsetWidth * z) + "px";
      sizer.style.height = Math.ceil(stage.offsetHeight * z) + "px";
      const off = (stage.offsetWidth - page.offsetWidth) * z / 2;
      ruler.style.transform = off > 0 ? `translateX(${(-off).toFixed(1)}px)` : "";
    }
    function setZoom(z, user) {
      z = clamp(Math.round(z * 100) / 100, 0.5, 2);
      st.zoom = z; if (user) st.userZoom = true;
      q(".wd-zr").value = Math.round(z * 100);
      q(".wd-zl").textContent = Math.round(z * 100) + " %";
      sizeStage(); drawRuler();
      later(positionCards, 30);
    }
    function fitWidth(force) {
      const w = canvas.clientWidth;
      if (!w || (st.userZoom && !force)) return;
      const extra = cmts.classList.contains("on") ? 270 : 0;
      setZoom(clamp(Math.floor((w - 40 - extra) / PAGE().pw * 20) / 20, 0.5, force ? 2 : 1), !!force);
    }
    function fitPage() {
      const h = canvas.clientHeight - 50, w = canvas.clientWidth - 40;
      if (!h || !w) return;
      setZoom(clamp(Math.floor(Math.min(h / PAGE().ph, w / PAGE().pw) * 20) / 20, 0.5, 2), true);
    }
    function drawRuler() {
      const g = PAGE(), z = st.zoom, W = g.pw * z, mx = g.mx * z, c = CM * z;
      let s = `<svg width="${W.toFixed(0)}" height="22" viewBox="0 0 ${W.toFixed(1)} 22" aria-hidden="true">`;
      s += `<rect x="0" y="4" width="${W.toFixed(1)}" height="14" class="wd-rm"/><rect x="${mx.toFixed(1)}" y="4" width="${(W - 2 * mx).toFixed(1)}" height="14" class="wd-rw"/>`;
      for (let k = -Math.floor(mx / c * 4); mx + k * c / 4 <= W; k++) {
        const x = mx + k * c / 4;
        if (x < 0) continue;
        if (k % 4 === 0) { if (k) s += `<text x="${x.toFixed(1)}" y="15" class="wd-rn">${Math.abs(k / 4)}</text>`; }
        else s += `<line x1="${x.toFixed(1)}" x2="${x.toFixed(1)}" y1="${k % 2 ? 10 : 8.5}" y2="${k % 2 ? 12 : 13.5}" class="wd-rt"/>`;
      }
      s += `<path d="M${mx - 5} 1h10l-5 5z" class="wd-rk"/><path d="M${mx - 5} 21h10l-5-5z" class="wd-rk"/><path d="M${W - mx - 5} 21h10l-5-5z" class="wd-rk"/></svg>`;
      ruler.innerHTML = s;
      ruler.style.width = W.toFixed(0) + "px";
    }
    function setColumns(n) {
      st.columns = clamp(n | 0, 1, 3);
      applyPage(); after();
      toast(st.columns === 1 ? "Mise en page sur une colonne." : `Mise en page sur ${st.columns} colonnes.`);
    }

    /* ---------------------------------------------------------------- texte / mots */
    function textOf(el) {
      let out = "";
      const walk = n => {
        for (const c of n.childNodes) {
          if (c.nodeType === 3) out += c.data;
          else if (c.nodeType === 1) {
            if (c.hasAttribute("data-wd-ui")) continue;
            if (c.tagName === "BR") { out += "\n"; continue; }
            const blk = BLOCK_RE.test(c.tagName) || /^(TR|TABLE|UL|OL|HR)$/.test(c.tagName);
            if (blk) out += "\n";
            walk(c);
            if (blk) out += /^(TD|TH)$/.test(c.tagName) ? "\t" : "\n";
          }
        }
      };
      walk(el);
      return out.replace(/ /g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{2,}/g, "\n").replace(/[ ]{2,}/g, " ").trim();
    }
    const countWords = t => (t.match(/[^\s ]+/g) || []).filter(w => /[\p{L}\p{N}]/u.test(w)).length;
    function stats() {
      const t = textOf(root);
      const paras = [...root.querySelectorAll("p,h1,h2,h3,h4,h5,h6,li,blockquote")].filter(p => !p.closest(".o-toc") && p.textContent.trim()).length;
      return { pages: pageCount(), words: countWords(t), chars: t.replace(/\s/g, "").length, charsSp: t.replace(/\n/g, "").length, paras, lines: Math.max(1, Math.round(root.scrollHeight / 21)) };
    }

    /* ---------------------------------------------------------------- api.fmt */
    function fmtOf(frag) {
      const res = { found: false, bold: false, italic: false, underline: false, strike: false, color: null, highlight: null, tag: null, align: null, list: null };
      const f = normS(frag).toLowerCase();
      if (!f) return res;
      const idx = buildIndex(root, true, true);
      const i = idx.str.toLowerCase().indexOf(f);
      if (i < 0) return res;
      let k = i;
      while (k < i + f.length && (idx.str[k] === " " || !idx.map[k])) k++;
      if (!idx.map[k]) return res;
      const node = idx.map[k][0], el = node.parentElement;
      const cs = getComputedStyle(el);
      const fw = cs.fontWeight === "bold" ? 700 : (parseInt(cs.fontWeight, 10) || 400);
      const block = blockOf(node, root);
      res.found = true;
      res.bold = fw >= 600;
      res.italic = /italic|oblique/.test(cs.fontStyle);
      for (let a = el; a && a !== root.parentNode; a = a.parentElement) {
        const s = getComputedStyle(a), d = (s.textDecorationLine || "") + " " + (s.textDecoration || "");
        if (/underline/.test(d)) res.underline = true;
        if (/line-through/.test(d)) res.strike = true;
        if (a === block || a === root) break;
      }
      res.color = rgbToHex(cs.color);
      for (let a = el; a && a !== block && a !== root; a = a.parentElement) {
        if (a.classList && (a.classList.contains("o-comment") || a.classList.contains("wd-pv"))) continue;
        const h = rgbToHex(getComputedStyle(a).backgroundColor);
        if (h) { res.highlight = h; break; }
      }
      let tag = block ? block.tagName.toLowerCase() : "p";
      if (tag === "div") tag = "p";
      if (block && tag === "p" && block.parentElement && block.parentElement.tagName === "BLOCKQUOTE" && block.parentElement !== root) tag = "blockquote";
      res.tag = tag;
      const al = getComputedStyle(block || root).textAlign || "left";
      res.align = /center/.test(al) ? "center" : /right|end/.test(al) ? "right" : /justify/.test(al) ? "justify" : "left";
      const li = closestIn(node, "li", root);
      res.list = li && li.parentElement ? (li.parentElement.tagName === "OL" ? "ol" : li.parentElement.tagName === "UL" ? "ul" : null) : null;
      return res;
    }
    function tablesInfo() {
      return [...root.querySelectorAll("table")].map(t => {
        const rows = [...t.rows];
        return { rows: rows.length, cols: rows.reduce((m, r) => Math.max(m, [...r.cells].reduce((s, c) => s + (c.colSpan || 1), 0)), 0), headerBold: rows.length ? rowBold(rows[0]) : false };
      });
    }

    /* ---------------------------------------------------------------- volets : navigation, commentaires, notes */
    function setNavTab(t) {
      qa("[data-ntab]").forEach(b => b.classList.toggle("on", b.dataset.ntab === t));
      qa("[data-nlist]").forEach(l => { l.hidden = l.dataset.nlist !== t; });
    }
    function renderNav(force) {
      const hs = headingEls(), sig = tocSig();
      const box = q('[data-nlist="h"]');
      if (force || sig !== st.navSig) {
        st.navSig = sig;
        box.innerHTML = hs.length ? hs.map((h, i) => `<button type="button" class="wd-nav-i wd-l${h.tagName[1]}" data-h="${i}" style="--d:${Math.min(i, 14)}" title="${esc(normS(h.textContent))}">${esc(normS(h.textContent) || "(titre vide)")}</button>`).join("")
          : `<div class="wd-nav-empty">Créez un plan interactif de votre document.<br><br>Pour commencer, accédez à l'onglet <b>Accueil</b> et appliquez les styles <b>Titre 1</b>, <b>Titre 2</b>… à vos titres.</div>`;
      }
      markNavCurrent();
    }
    function markNavCurrent() {
      const r = st.lastRange;
      const hs = headingEls();
      let cur = -1;
      if (r && root.contains(r.startContainer)) hs.forEach((h, i) => { if (h.contains(r.startContainer) || (h.compareDocumentPosition(r.startContainer) & Node.DOCUMENT_POSITION_FOLLOWING)) cur = i; });
      qa(".wd-nav-i").forEach(b => b.classList.toggle("on", +b.dataset.h === cur));
    }
    function renderCmts() {
      const ms = commentMarks();
      const show = st.showCmts && ms.length > 0;
      const was = cmts.classList.contains("on");
      cmts.classList.toggle("on", show);
      const sig = ms.map(m => (m.getAttribute("data-note") || "") + "¦" + m.textContent).join("\n");
      if (sig !== st.cmtSig && !cmts.contains(document.activeElement)) {
        st.cmtSig = sig;
        cmts.innerHTML = ms.map((m, i) => {
          const d = new Date(m.getAttribute("data-date") || Date.now());
          const when = isNaN(d) ? "" : d.toLocaleTimeString("fr-FR", { hour: "2-digit", minute: "2-digit" });
          const au = m.getAttribute("data-author") || "Relecteur";
          return `<div class="wd-card${i === st.activeCmt ? " on" : ""}" data-i="${i}"><div class="wd-card-h"><span class="wd-cav">${esc(au.charAt(0).toUpperCase())}</span><b>${esc(au)}</b><time>${esc(when)}</time><button type="button" class="wd-btn wd-ico wd-card-x" data-cmd="cmtDelI" data-arg="${i}" title="Supprimer le commentaire">${ic("trash")}</button></div><div class="wd-card-q">« ${esc(normS(m.textContent).slice(0, 60))} »</div><div class="wd-card-t" contenteditable="true" spellcheck="true" data-ci="${i}">${esc(m.getAttribute("data-note") || "")}</div></div>`;
        }).join("");
      }
      if (show !== was) { sizeStage(); if (!st.userZoom) fitWidth(); }
      positionCards();
    }
    function positionCards() {
      if (!cmts.classList.contains("on")) return;
      const ms = commentMarks(), sr = cmts.getBoundingClientRect(), s = scale();
      let y = 0;
      qa(".wd-card").forEach(c => {
        const m = ms[+c.dataset.i];
        if (!m) return;
        const top = (m.getBoundingClientRect().top - sr.top) / s - 8;
        const t = Math.max(top, y);
        c.style.transform = `translateY(${t.toFixed(1)}px)`;
        y = t + c.offsetHeight + 10;
      });
    }
    function renumberFn() {
      root.querySelectorAll("sup.o-fn").forEach((s, i) => { const v = String(i + 1); if (s.textContent !== v) s.textContent = v; if (s.getAttribute("contenteditable") !== "false") s.contentEditable = "false"; });
    }
    function renderFn() {
      const sups = [...root.querySelectorAll("sup.o-fn")];
      const sig = sups.map(s => s.getAttribute("data-note") || "").join("\n");
      if (sig === st.fnSig || fnotes.contains(document.activeElement)) return;
      st.fnSig = sig;
      fnotes.innerHTML = sups.length ? sups.map((s, i) => `<div class="wd-fn"><sup>${i + 1}</sup><span class="wd-fn-t" contenteditable="true" data-fi="${i}" spellcheck="true">${esc(s.getAttribute("data-note") || "")}</span></div>`).join("") : "";
    }

    /* ---------------------------------------------------------------- état du ruban */
    let stateRaf = 0;
    function schedState() { if (!stateRaf) stateRaf = requestAnimationFrame(() => { stateRaf = 0; if (!st.dead) updateState(); }); }
    function qcs(c) { try { return document.queryCommandState(c); } catch (e) { return false; } }
    function updateState() {
      const r = curRange();
      const inTbl = !!(r && closestIn(r.startContainer, "td,th", root));
      const tt = q('.wd-tab[data-tab="tbl"]');
      if (tt.hidden === inTbl) { tt.hidden = !inTbl; if (!inTbl && st.tab === "tbl") setTab("home"); updateTabs(); }
      if (inTbl) { const c = cellCtx(); q('[data-tog="hdrRow"]').classList.toggle("on", !!(c && rowBold(c.table.rows[0]))); }
      if (!r) return;
      const on = (sel, v) => qa(sel).forEach(b => b.classList.toggle("on", !!v));
      on('[data-cmd="bold"]', qcs("bold")); on('[data-cmd="italic"]', qcs("italic")); on('[data-cmd="underline"]', qcs("underline"));
      on('[data-cmd="strike"]', qcs("strikeThrough")); on('[data-cmd="sub"]', qcs("subscript")); on('[data-cmd="sup"]', qcs("superscript"));
      let n = r.startContainer; if (n.nodeType === 3) n = n.parentNode;
      const b = blockOf(r.startContainer);
      const li = closestIn(r.startContainer, "li", edActive());
      on('[data-cmd="ul"]', li && li.parentElement.tagName === "UL"); on('[data-cmd="ol"]', li && li.parentElement.tagName === "OL");
      if (n && n.nodeType === 1) {
        const cs = getComputedStyle(n);
        const fam = (cs.fontFamily.split(",")[0] || "").replace(/["']/g, "").trim();
        const fsel = q(".wd-font"); const hit = FONTS.find(f => f[0].toLowerCase() === fam.toLowerCase());
        if (hit) fsel.value = hit[0]; else if (/calibri|carlito/i.test(fam)) fsel.value = "Calibri";
        const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
        const ssel = q(".wd-size");
        if (!SIZES.includes(pt)) { let o = ssel.querySelector("option[data-x]"); if (!o) { o = document.createElement("option"); o.setAttribute("data-x", ""); ssel.prepend(o); } o.value = pt; o.textContent = String(pt).replace(".", ","); }
        ssel.value = String(pt);
        const al = getComputedStyle(b || edActive()).textAlign;
        const a = /center/.test(al) ? "center" : /right|end/.test(al) ? "right" : /justify/.test(al) ? "justify" : "left";
        qa('[data-cmd="al"]').forEach(x => x.classList.toggle("on", x.dataset.arg === a));
      }
      let tag = b ? b.tagName.toLowerCase() : "p";
      if (b && b.parentElement && b.parentElement.tagName === "BLOCKQUOTE" && b.parentElement !== edActive()) tag = "blockquote";
      qa(".wd-sty").forEach(x => x.classList.toggle("on", x.dataset.style === tag || (tag !== "h1" && tag !== "h2" && tag !== "h3" && tag !== "blockquote" && x.dataset.style === "p")));
      // commentaire actif sous le curseur
      const m = closestIn(r.startContainer, "mark.o-comment", root);
      if (m) { const i = commentMarks().indexOf(m); if (i !== st.activeCmt) activateCmt(i); }
      markNavCurrent();
      const pg = b && root.contains(b) ? pageOf(b) : 1;
      q(".wd-st-page").textContent = `Page ${pg} sur ${Math.max(pg, pageCount())}`;
    }
    function updateTabs() {
      const act = q(`.wd-tab[data-tab="${st.tab}"]`);
      qa(".wd-tab").forEach(t => { t.classList.toggle("on", t === act); t.setAttribute("aria-selected", t === act ? "true" : "false"); });
      const bar = q(".wd-tabbar");
      if (act && !act.hidden) { bar.style.width = (act.offsetWidth - 16) + "px"; bar.style.transform = `translateX(${act.offsetLeft + 8}px)`; bar.style.opacity = "1"; }
      qa("[data-tog]").forEach(b => {
        const k = b.dataset.tog;
        if (k === "hdrRow") return;
        b.classList.toggle("on", !!(k === "nav" ? st.nav : k === "ruler" ? st.ruler : k === "marks" ? st.marks : k === "cmts" ? st.showCmts : false));
      });
      q('[data-cmd="marks"]').classList.toggle("on", st.marks);
    }
    function setTab(t) {
      if (t === "file") { openBackstage(); return; }
      st.tab = t;
      qa(".wd-panel").forEach(p => {
        const show = p.dataset.panel === t;
        if (show && p.hidden) { p.hidden = false; if (!reduced()) { p.classList.remove("wd-enter"); void p.offsetWidth; p.classList.add("wd-enter"); } }
        else if (!show) p.hidden = true;
      });
      updateTabs();
    }

    /* ---------------------------------------------------------------- menus déroulants */
    const MENUS = {
      color: () => `<div class="wd-pop-t">Couleurs du thème</div><div class="wd-sw wd-sw10">${[0, .8, .6, .4, -.25, -.5].map(t => THEME.map(c => {
        const v = t === 0 ? c : (c === "#FFFFFF" ? mix("#FFFFFF", -Math.abs(t) * 0.6) : c === "#000000" ? mix("#000000", Math.abs(t)) : mix(c, t));
        return `<button type="button" data-cmd="color" data-arg="${v}" style="--c:${v}" title="${v}"></button>`;
      }).join("")).join("")}</div><div class="wd-pop-t">Couleurs standard</div><div class="wd-sw wd-sw10">${STD.map(c => `<button type="button" data-cmd="color" data-arg="${c}" style="--c:${c}" title="${c}"></button>`).join("")}</div><button type="button" class="wd-mi" data-cmd="color" data-arg="#000000"><i class="wd-dot" style="--c:#000"></i>Automatique</button>`,
      hilite: () => `<div class="wd-sw wd-sw5">${HILITE.map(h => `<button type="button" data-cmd="hilite" data-arg="${h[0]}" style="--c:${h[0]}" title="${h[1]}"></button>`).join("")}</div><button type="button" class="wd-mi" data-cmd="hilite" data-arg="none">${ic("close")}Aucune couleur</button>`,
      table: () => `<div class="wd-pop-t wd-tg-l">Insérer un tableau</div><div class="wd-tgrid">${Array.from({ length: 80 }, (_, i) => `<i data-tr="${Math.floor(i / 10) + 1}" data-tc="${i % 10 + 1}"></i>`).join("")}</div><button type="button" class="wd-mi" data-cmd="tableDlg">${ic("table")}Insérer un tableau…</button>`,
      cols: () => [[1, "Une", "col1"], [2, "Deux", "col2"], [3, "Trois", "col3"]].map(c => `<button type="button" class="wd-mi wd-mi-l${st.columns === c[0] ? " on" : ""}" data-cmd="cols" data-arg="${c[0]}">${ic(c[2])}<span>${c[1]}</span></button>`).join(""),
      margins: () => Object.keys(MARGINS).map(k => { const m = MARGINS[k]; const f = v => String(v).replace(".", ","); return `<button type="button" class="wd-mi wd-mi-l${st.margins === k ? " on" : ""}" data-cmd="margins" data-arg="${k}">${ic("margins")}<span><b>${m.l}</b><small>Haut : ${f(m.y)} cm · Bas : ${f(m.y)} cm<br>Gauche : ${f(m.x)} cm · Droite : ${f(m.x)} cm</small></span></button>`; }).join(""),
      orient: () => [["portrait", "Portrait"], ["landscape", "Paysage"]].map(o => `<button type="button" class="wd-mi wd-mi-l${st.orient === o[0] ? " on" : ""}" data-cmd="orient" data-arg="${o[0]}">${ic(o[0])}<span>${o[1]}</span></button>`).join(""),
      tstyle: () => `<div class="wd-pop-t">Styles de tableau</div><div class="wd-tsg">${[["grid", "Grille du tableau"], ["blue", "Tableau Grille 4 – Accentuation 1"], ["zebra", "Tableau Liste 2 (lignes à bandes)"], ["orange", "Tableau Grille 4 – Accentuation 2"]].map(s => `<button type="button" class="wd-tsp wd-tsp-${s[0]}" data-cmd="tstyle" data-arg="${s[0]}" title="${s[1]}"><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i><i></i></button>`).join("")}</div>`
    };
    function openMenu(btn, name) {
      const same = st.popBtn === btn;
      closeMenu();
      if (same || !MENUS[name]) return;
      const pop = document.createElement("div");
      pop.className = "wd-pop wd-pop-" + name;
      pop.innerHTML = MENUS[name]();
      app.appendChild(pop);
      const ar = app.getBoundingClientRect(), br = btn.getBoundingClientRect();
      const left = clamp(br.left - ar.left, 8, Math.max(8, ar.width - pop.offsetWidth - 8));
      pop.style.left = left + "px";
      pop.style.top = (br.bottom - ar.top + 4) + "px";
      btn.classList.add("wd-open");
      st.pop = pop; st.popBtn = btn;
      if (name === "table") {
        const lab = pop.querySelector(".wd-tg-l");
        pop.addEventListener("mouseover", e => {
          const c = e.target.closest("i[data-tr]");
          if (!c) return;
          const R = +c.dataset.tr, Cc = +c.dataset.tc;
          pop.querySelectorAll("i[data-tr]").forEach(x => x.classList.toggle("on", +x.dataset.tr <= R && +x.dataset.tc <= Cc));
          lab.textContent = `Tableau ${Cc} × ${R}`;
        });
        pop.addEventListener("mouseleave", () => { pop.querySelectorAll("i.on").forEach(x => x.classList.remove("on")); lab.textContent = "Insérer un tableau"; });
        pop.addEventListener("click", e => { const c = e.target.closest("i[data-tr]"); if (c) { closeMenu(); insertTable(+c.dataset.tr, +c.dataset.tc); } });
      }
    }
    function closeMenu() {
      if (st.pop) { const p = st.pop; p.classList.add("wd-out"); later(() => p.remove(), 140); }
      if (st.popBtn) st.popBtn.classList.remove("wd-open");
      st.pop = null; st.popBtn = null;
    }

    /* ---------------------------------------------------------------- boîtes de dialogue */
    function ask(title, fields, okLabel) {
      return new Promise(res => {
        const m = document.createElement("div");
        m.className = "wd-modal";
        m.innerHTML = `<form class="wd-dlg" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="wd-dlg-h"><b>${esc(title)}</b><button type="button" class="wd-btn wd-ico" data-x title="Fermer">${ic("close")}</button></div><div class="wd-dlg-b">${fields.map(f => `<label class="wd-fld"><span>${esc(f.label)}</span>${f.area ? `<textarea name="${f.name}" rows="3" placeholder="${esc(f.ph || "")}">${esc(f.value || "")}</textarea>` : `<input name="${f.name}" type="${f.type || "text"}" value="${esc(f.value == null ? "" : f.value)}"${f.min != null ? ` min="${f.min}"` : ""}${f.max != null ? ` max="${f.max}"` : ""} placeholder="${esc(f.ph || "")}" autocomplete="off">`}</label>`).join("")}</div><div class="wd-dlg-f"><span class="wd-grow"></span><button type="submit" class="wd-pbtn">${esc(okLabel || "OK")}</button><button type="button" class="wd-sbtn" data-x>Annuler</button></div></form>`;
        app.appendChild(m);
        const form = m.querySelector("form");
        let closed = false;
        const done = v => { if (closed) return; closed = true; st.modal = null; m.classList.add("wd-out"); later(() => m.remove(), 180); res(v); };
        st.modal = () => done(null);
        form.addEventListener("submit", e => { e.preventDefault(); const v = {}; fields.forEach(f => { v[f.name] = form.elements[f.name].value; }); done(v); });
        m.addEventListener("mousedown", e => { if (e.target === m) done(null); });
        m.addEventListener("click", e => { if (e.target.closest("[data-x]")) done(null); });
        form.addEventListener("keydown", e => {
          if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); done(null); }
          else if (e.key === "Enter" && e.target.tagName === "TEXTAREA" && !e.shiftKey) { e.preventDefault(); form.requestSubmit ? form.requestSubmit() : form.dispatchEvent(new Event("submit", { cancelable: true })); }
        });
        later(() => { const f = form.querySelector("input,textarea"); if (f) { f.focus(); if (f.select && f.type !== "number") f.select(); } }, 40);
      });
    }
    function info(title, html) {
      const m = document.createElement("div");
      m.className = "wd-modal";
      m.innerHTML = `<div class="wd-dlg" role="dialog" aria-modal="true" aria-label="${esc(title)}"><div class="wd-dlg-h"><b>${esc(title)}</b><button type="button" class="wd-btn wd-ico" data-x title="Fermer">${ic("close")}</button></div><div class="wd-dlg-b">${html}</div><div class="wd-dlg-f"><span class="wd-grow"></span><button type="button" class="wd-pbtn" data-x>Fermer</button></div></div>`;
      app.appendChild(m);
      const done = () => { st.modal = null; m.classList.add("wd-out"); later(() => m.remove(), 180); restore(); };
      st.modal = done;
      m.addEventListener("click", e => { if (e.target === m || e.target.closest("[data-x]")) done(); });
      later(() => { const b = m.querySelector(".wd-pbtn"); b && b.focus(); }, 40);
    }
    function showStats() {
      const s = stats();
      const row = (k, v) => `<tr><td>${k}</td><td><b>${v.toLocaleString("fr-FR")}</b></td></tr>`;
      info("Statistiques", `<table class="wd-stat">${row("Pages", s.pages)}${row("Mots", s.words)}${row("Caractères (espaces non compris)", s.chars)}${row("Caractères (espaces compris)", s.charsSp)}${row("Paragraphes", s.paras)}${row("Lignes (estimation)", s.lines)}</table>`);
    }

    /* ---------------------------------------------------------------- Fichier (backstage) */
    function openBackstage(sec) {
      bs.hidden = false;
      requestAnimationFrame(() => bs.classList.add("wd-in"));
      bsSection(sec || "info");
    }
    function closeBackstage() { bs.classList.remove("wd-in"); later(() => { bs.hidden = true; }, reduced() ? 0 : 220); restore(); }
    function bsSection(k) {
      qa(".wd-bs-i").forEach(b => b.classList.toggle("on", b.dataset.bs === k));
      const main = q(".wd-bs-main");
      const s = stats();
      if (k === "info") main.innerHTML = `<h2>Informations</h2><div class="wd-bs-card"><div class="wd-bs-doc">${ic("word")}<div><b>${esc(st.docName)}</b><small>Document Word · simulation Office Academy</small></div></div><table class="wd-stat">${[["Pages", s.pages], ["Mots", s.words], ["Caractères", s.charsSp], ["Paragraphes", s.paras], ["Commentaires", commentMarks().length], ["Colonnes", st.columns], ["Orientation", st.orient === "landscape" ? "Paysage" : "Portrait"]].map(r => `<tr><td>${r[0]}</td><td><b>${r[1]}</b></td></tr>`).join("")}</table></div>`;
      else if (k === "save") main.innerHTML = `<h2>Enregistrer une copie</h2><div class="wd-bs-card"><p>Téléchargez le document au format <b>Word 97-2003 / HTML (.doc)</b> : il s'ouvre dans Microsoft Word avec ses styles (Titres, Citation), tableaux, en-tête, pied de page, notes et commentaires.</p><button type="button" class="wd-pbtn wd-bs-go" data-cmd="download">${ic("download")} Télécharger document.doc</button></div>`;
      else if (k === "print") main.innerHTML = `<h2>Imprimer</h2><div class="wd-bs-card"><p>L'impression est simulée dans cette formation. Pour imprimer réellement, téléchargez la copie (.doc) puis utilisez <b>Fichier › Imprimer</b> (<kbd>${MOD}</kbd>+<kbd>P</kbd>) dans Word.</p><button type="button" class="wd-sbtn" data-bs="save">${ic("download")} Enregistrer une copie</button></div>`;
      else if (k === "reset") main.innerHTML = `<h2>Réinitialiser</h2><div class="wd-bs-card"><p>Revenir au document de départ du lab. <b>Toutes vos modifications seront perdues.</b></p><button type="button" class="wd-pbtn wd-danger" data-cmd="resetDoc">${ic("reset")} Réinitialiser le document</button></div>`;
    }
    function download() {
      try {
        const f = exportFile();
        const blob = new Blob(["﻿", f.content], { type: f.type });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob); a.download = f.name; a.style.display = "none";
        document.body.appendChild(a); a.click();
        later(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
        toast("Copie téléchargée : " + f.name);
      } catch (e) { toast("Téléchargement impossible dans ce navigateur."); }
    }

    /* ---------------------------------------------------------------- historique (annuler / rétablir) */
    function caretOff() {
      const r = curRange();
      if (!r || !root.contains(r.startContainer)) return -1;
      const pre = document.createRange(); pre.selectNodeContents(root); pre.setEnd(r.startContainer, r.startOffset);
      return pre.toString().length;
    }
    function setCaretOff(off) {
      if (off < 0) return;
      const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let n, acc = 0, last = null;
      while ((n = w.nextNode())) {
        last = n;
        if (acc + n.length >= off) { const r = document.createRange(); r.setStart(n, off - acc); r.collapse(true); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); st.lastRange = r.cloneRange(); return; }
        acc += n.length;
      }
      if (last) caretAfter(last);
    }
    const snapHTML = () => root.innerHTML.replace(/ ?\bwd-(pv(-\w+)?|cmt-on|flash|born)\b/g, "").replace(/ class=""/g, "");
    function pushHist() {
      if (st.histLock) return;
      const s = { html: snapHTML(), hdr: hdr.innerHTML, ftr: ftr.innerHTML, cols: st.columns, margins: st.margins, orient: st.orient, off: caretOff() };
      const cur = st.hist[st.hidx];
      if (cur && cur.html === s.html && cur.hdr === s.hdr && cur.ftr === s.ftr && cur.cols === s.cols && cur.margins === s.margins && cur.orient === s.orient) { cur.off = s.off; return; }
      st.hist = st.hist.slice(0, st.hidx + 1);
      st.hist.push(s);
      if (st.hist.length > 120) st.hist.shift();
      st.hidx = st.hist.length - 1;
    }
    function applySnap(s) {
      st.histLock = true;
      root.innerHTML = s.html; hdr.innerHTML = s.hdr; ftr.innerHTML = s.ftr;
      st.columns = s.cols; st.margins = s.margins; st.orient = s.orient;
      applyPage();
      try { root.focus({ preventScroll: true }); } catch (e) { /* */ }
      setCaretOff(s.off);
      st.histLock = false;
    }
    function undo() {
      clearTimeout(tRefresh); refresh(true);
      if (st.hidx <= 0) return toast("Rien à annuler.");
      st.hidx--; applySnap(st.hist[st.hidx]); refresh(true, true);
    }
    function redo() {
      if (st.hidx >= st.hist.length - 1) return toast("Rien à rétablir.");
      st.hidx++; applySnap(st.hist[st.hidx]); refresh(true, true);
    }

    /* ---------------------------------------------------------------- rafraîchissement global */
    let tRefresh = 0;
    function changed() {
      const sv = q(".wd-saved");
      sv.classList.add("busy"); sv.lastChild.textContent = "Enregistrement…";
      clearTimeout(tRefresh);
      tRefresh = later(refresh, 150);
    }
    const after = () => { changed(); schedState(); };
    function normalizeDoc() {
      ensureBlocks(root);
      [...root.children].forEach(d => {
        if (d.tagName === "DIV" && !d.classList.contains("o-toc") && !d.hasAttribute("data-wd-ui")) {
          if (d.querySelector(BLOCK_SEL + ",ul,ol,table")) unwrap(d);
          else { const pts = savePts(), map = new Map(); setTag(d, "p", map); restorePts(pts, map); }
        }
      });
      if (!root.querySelector("p,h1,h2,h3,h4,h5,h6,li,td,blockquote,hr,.o-toc,table")) { root.innerHTML = "<p><br></p>"; if (!st.hf) caretIn(root.firstChild); }
      const last = root.lastElementChild;
      if (last && (last.tagName === "TABLE" || last.tagName === "HR" || last.classList.contains("o-toc"))) last.after(newP());
      root.querySelectorAll(".o-toc").forEach(t => { if (t.getAttribute("contenteditable") !== "false") t.contentEditable = "false"; });
    }
    function refresh(silent, noHist) {
      if (st.dead) return;
      clearTimeout(tRefresh);
      if (!st.hf) normalizeDoc();
      renumberFn(); renderFn();
      const toc = root.querySelector(".o-toc");
      if (toc) { toc.classList.toggle("wd-stale", toc.getAttribute("data-sig") !== tocSig()); }
      renderNav(); renderCmts(); sizeStage();
      const s = stats();
      q(".wd-st-words").textContent = `${s.words.toLocaleString("fr-FR")} mot${s.words > 1 ? "s" : ""}`;
      const pc = s.pages;
      const b = st.lastRange && blockOf(st.lastRange.startContainer, root);
      const pg = b ? pageOf(b) : 1;
      q(".wd-st-page").textContent = `Page ${Math.min(pg, pc)} sur ${pc}`;
      q(".wd-st-mode").hidden = !st.hf; q(".wd-st-mode").textContent = st.hf ? "Modification de l'en-tête et du pied de page" : "";
      if (!noHist) pushHist();
      clearTimeout(st.saveT);
      st.saveT = later(() => { const sv = q(".wd-saved"); sv.classList.remove("busy"); sv.lastChild.textContent = "Enregistré"; }, 600);
      if (!silent || noHist) { try { ctx.changed && ctx.changed(); } catch (e) { /* */ } }
    }

    /* ---------------------------------------------------------------- export Word (.doc HTML) */
    function exportFile() {
      const g = PAGE(), m = MARGINS[st.margins] || MARGINS.normal, land = st.orient === "landscape";
      const clone = root.cloneNode(true);
      clone.querySelectorAll("[data-wd-ui]").forEach(e => e.remove());
      clone.querySelectorAll("[contenteditable]").forEach(e => e.removeAttribute("contenteditable"));
      clone.querySelectorAll("[class]").forEach(e => { e.classList.remove("wd-pv", "wd-pv-p", "wd-pv-h1", "wd-pv-h2", "wd-pv-h3", "wd-pv-blockquote", "wd-cmt-on", "wd-flash", "wd-stale", "wd-born"); if (!e.getAttribute("class")) e.removeAttribute("class"); });
      // titres : signets pour la TM
      const hs = [...clone.querySelectorAll("h1,h2,h3")].filter(h => !h.closest(".o-toc"));
      hs.forEach((h, i) => { const a = document.createElement("a"); a.setAttribute("name", "_Toc" + (100 + i)); h.insertBefore(a, h.firstChild); });
      // table des matières → champ TOC
      clone.querySelectorAll(".o-toc").forEach(t => {
        const es = [...t.querySelectorAll("a[data-h]")];
        const tab = (g.pw - 2 * g.mx) * 0.75;
        let h = `<p class="MsoTocHeading">Table des matières</p>`;
        if (!es.length) h += `<p class="MsoToc1"><!--[if supportFields]><span style='mso-element:field-begin'></span> TOC \\o "1-3" \\h \\z \\u <span style='mso-element:field-separator'></span><![endif]-->Aucune entrée de table des matières n'a été trouvée.<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]--></p>`;
        es.forEach((a, i) => {
          const lvl = (a.parentElement.className.match(/o-toc-(\d)/) || [0, 1])[1];
          const txt = esc(a.querySelector(".o-toc-t") ? a.querySelector(".o-toc-t").textContent : a.textContent);
          const pg = esc(a.querySelector(".o-toc-p") ? a.querySelector(".o-toc-p").textContent : "");
          h += `<p class="MsoToc${lvl}" style="tab-stops:right dotted ${tab.toFixed(1)}pt">${i === 0 ? `<!--[if supportFields]><span style='mso-element:field-begin'></span> TOC \\o "1-3" \\h \\z \\u <span style='mso-element:field-separator'></span><![endif]-->` : ""}<a href="#_Toc${100 + +a.dataset.h}">${txt}<span style='mso-tab-count:1 dotted'>. . . . . </span>${pg}</a>${i === es.length - 1 ? `<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->` : ""}</p>`;
        });
        const d = document.createElement("div"); d.innerHTML = h;
        t.replaceWith(...d.childNodes);
      });
      // sauts de page
      clone.querySelectorAll("hr.o-pagebreak").forEach(hr => { const s = document.createElement("span"); s.setAttribute("data-wd-br", ""); hr.replaceWith(s); });
      // citation
      clone.querySelectorAll("blockquote").forEach(b => { const p = document.createElement("p"); p.className = "MsoQuote"; if (b.getAttribute("style")) p.setAttribute("style", b.getAttribute("style")); while (b.firstChild) p.appendChild(b.firstChild); b.replaceWith(p); });
      // tableaux : bordures en ligne
      clone.querySelectorAll("table").forEach(t => {
        t.setAttribute("border", "1"); t.setAttribute("cellspacing", "0"); t.setAttribute("cellpadding", "0");
        t.setAttribute("style", "border-collapse:collapse;border:none;mso-border-alt:solid windowtext .5pt;mso-yfti-tbllook:1184;width:100%");
        const blue = t.classList.contains("wd-ts-blue"), orange = t.classList.contains("wd-ts-orange"), zebra = t.classList.contains("wd-ts-zebra");
        [...t.rows].forEach((r, ri) => [...r.cells].forEach(c => {
          let s = "border:solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;padding:0cm 5.4pt;vertical-align:top;";
          if ((blue || orange) && ri === 0) s += `background:${blue ? "#156082" : "#E97132"};color:white;font-weight:bold;`;
          else if ((blue || orange || zebra) && ri % 2 === 1) s += `background:${blue ? "#C1E4F5" : orange ? "#FBE2D5" : "#F2F2F2"};`;
          if (c.getAttribute("style")) s += c.getAttribute("style");
          c.setAttribute("style", s);
        }));
      });
      // commentaires
      const cl = [];
      clone.querySelectorAll("mark.o-comment").forEach((mk, i) => {
        const n = i + 1, s = document.createElement("span");
        s.setAttribute("style", "mso-comment-continuation:" + n);
        while (mk.firstChild) s.appendChild(mk.firstChild);
        const ref = document.createElement("span"); ref.setAttribute("data-wd-com", String(n));
        mk.replaceWith(s, ref);
        cl.push({ n, note: mk.getAttribute("data-note") || "", au: mk.getAttribute("data-author") || "Vous" });
      });
      // notes de bas de page
      const fl = [];
      clone.querySelectorAll("sup.o-fn").forEach((sp, i) => {
        const n = i + 1, ph = document.createElement("span"); ph.setAttribute("data-wd-ftn", String(n));
        sp.replaceWith(ph); fl.push({ n, note: sp.getAttribute("data-note") || "" });
      });
      let body = clone.innerHTML
        .replace(/<span data-wd-br=""><\/span>/g, "<br clear=all style='mso-special-character:line-break;page-break-before:always'>")
        .replace(/<span data-wd-com="(\d+)"><\/span>/g, (_, n) => `<span class=MsoCommentReference><span style='mso-special-character:comment'>&nbsp;<![if !supportAnnotations]><a class=msocomanchor id="_anchor_${n}" href="#_msocom_${n}" name="_msoanchor_${n}">[V${n}]</a><![endif]></span></span>`)
        .replace(/<span data-wd-ftn="(\d+)"><\/span>/g, (_, n) => `<a style='mso-footnote-id:ftn${n}' href="#_ftn${n}" name="_ftnref${n}" title=""><span class=MsoFootnoteReference><span style='mso-special-character:footnote'><![if !supportFootnotes]>[${n}]<![endif]></span></span></a>`);
      const hf = (el, cls) => {
        const c = el.cloneNode(true);
        c.querySelectorAll("[contenteditable]").forEach(e => e.removeAttribute("contenteditable"));
        c.querySelectorAll("span.wd-pgnum").forEach(s => { s.outerHTML = `<!--[if supportFields]><span style='mso-element:field-begin'></span> PAGE <span style='mso-element:field-separator'></span><![endif]-->1<!--[if supportFields]><span style='mso-element:field-end'></span><![endif]-->`; });
        let h = c.innerHTML.trim();
        if (!/^<(p|h\d|div|table|ul|ol)/i.test(h)) h = `<p class=${cls}>${h || "&nbsp;"}</p>`;
        else h = h.replace(/<p(?=[\s>])/g, `<p class=${cls}`);
        return h;
      };
      const comments = cl.length ? `<div style='mso-element:comment-list'><![if !supportAnnotations]><hr class=msocomoff align=left size=1 width="33%"><![endif]>${cl.map(c => `<div style='mso-element:comment'><![if !supportAnnotations]><div id="_com_${c.n}" class=msocomtxt><![endif]><span style='mso-comment-author:"${esc(c.au)}"'><a name="_msocom_${c.n}"></a></span><p class=MsoCommentText><span class=MsoCommentReference><span style='mso-special-character:comment'>&nbsp;<![if !supportAnnotations]><a href="#_msoanchor_${c.n}" class=msocomoff>[V${c.n}]</a><![endif]></span></span>${esc(c.note)}</p><![if !supportAnnotations]></div><![endif]></div>`).join("")}</div>` : "";
      const footnotes = fl.length ? `<div style='mso-element:footnote-list'><![if !supportFootnotes]><br clear=all><hr align=left size=1 width="33%"><![endif]>${fl.map(f => `<div style='mso-element:footnote' id=ftn${f.n}><p class=MsoFootnoteText><a style='mso-footnote-id:ftn${f.n}' href="#_ftnref${f.n}" name="_ftn${f.n}" title=""><span class=MsoFootnoteReference><span style='mso-special-character:footnote'><![if !supportFootnotes]>[${f.n}]<![endif]></span></span></a> ${esc(f.note)}</p></div>`).join("")}</div>` : "";
      const cm = v => String(v).replace(",", ".") + "cm";
      const content = `<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<title>${esc(st.docName)}</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom><w:TrackMoves>false</w:TrackMoves><w:TrackFormatting/><w:DoNotOptimizeForBrowser/><w:LidThemeOther>FR</w:LidThemeOther></w:WordDocument></xml><![endif]-->
<style>
<!--
@page Section1 { size:${land ? "29.7cm 21.0cm" : "21.0cm 29.7cm"}; ${land ? "mso-page-orientation:landscape; " : ""}margin:${cm(m.y)} ${cm(m.x)} ${cm(m.y)} ${cm(m.x)}; mso-header-margin:1.25cm; mso-footer-margin:1.25cm; mso-header:h1; mso-footer:f1; mso-paper-source:0;${st.columns > 1 ? ` mso-columns:${st.columns} even 35.4pt;` : ""} }
div.Section1 { page:Section1; }
p.MsoNormal, li.MsoNormal, div.MsoNormal, p, li { margin:0cm 0cm 8.0pt 0cm; line-height:115%; font-size:12.0pt; font-family:"Aptos",sans-serif; mso-fareast-language:FR; }
li { margin-bottom:0; }
h1 { mso-style-name:"Titre 1"; mso-outline-level:1; margin:18.0pt 0cm 4.0pt 0cm; font-size:20.0pt; font-family:"Aptos Display",sans-serif; color:#0F4761; font-weight:normal; page-break-after:avoid; }
h2 { mso-style-name:"Titre 2"; mso-outline-level:2; margin:8.0pt 0cm 4.0pt 0cm; font-size:16.0pt; font-family:"Aptos Display",sans-serif; color:#0F4761; font-weight:normal; page-break-after:avoid; }
h3 { mso-style-name:"Titre 3"; mso-outline-level:3; margin:8.0pt 0cm 4.0pt 0cm; font-size:14.0pt; font-family:"Aptos",sans-serif; color:#0F4761; font-weight:normal; page-break-after:avoid; }
p.MsoQuote { mso-style-name:Citation; margin:10.0pt 43.2pt 8.0pt 43.2pt; text-align:center; font-style:italic; color:#404040; }
p.MsoTocHeading { margin:12.0pt 0cm 6.0pt 0cm; font-size:20.0pt; font-family:"Aptos Display",sans-serif; color:#0F4761; }
p.MsoToc1 { margin:0cm 0cm 5.0pt 0cm; } p.MsoToc2 { margin:0cm 0cm 5.0pt 12.0pt; } p.MsoToc3 { margin:0cm 0cm 5.0pt 24.0pt; }
p.MsoHeader, p.MsoFooter { margin:0cm; font-size:10.0pt; tab-stops:center 8.0cm right 16.0cm; }
p.MsoFootnoteText { margin:0cm; font-size:10.0pt; } span.MsoFootnoteReference { vertical-align:super; }
p.MsoCommentText { margin:0cm; font-size:10.0pt; } span.MsoCommentReference { font-size:8.0pt; }
a:link { color:#467886; text-decoration:underline; }
table { border-collapse:collapse; } td, th { font-size:12.0pt; }
-->
</style>
</head>
<body lang=FR style='tab-interval:35.4pt'>
<div class=Section1>
${body}
</div>
${comments}
${footnotes}
<div style='mso-element:header' id=h1>${hf(hdr, "MsoHeader")}</div>
<div style='mso-element:footer' id=f1>${hf(ftr, "MsoFooter")}</div>
</body>
</html>`;
      return { name: "document.doc", type: "application/msword", content };
    }

    /* ---------------------------------------------------------------- commandes */
    const CMD = {
      paste: () => {
        restore();
        if (navigator.clipboard && navigator.clipboard.readText) {
          navigator.clipboard.readText().then(t => { if (t) { restore(); document.execCommand("insertText", false, t); after(); } }).catch(() => toast(`Utilisez ${MOD}+V pour coller.`));
        } else toast(`Utilisez ${MOD}+V pour coller.`);
      },
      cut: () => { if (!exec("cut")) toast(`Utilisez ${MOD}+X pour couper.`); },
      copy: () => { if (!exec("copy")) toast(`Utilisez ${MOD}+C pour copier.`); else toast("Copié dans le Presse-papiers."); },
      undo, redo,
      save: () => { refresh(true); toast("Document enregistré (simulation). Fichier › Enregistrer une copie pour obtenir le .doc."); },
      autosave: b => { b.classList.toggle("on"); },
      share: () => toast("Partage simulé : dans Word, Partager › Copier le lien."),
      bold: () => exec("bold"), italic: () => exec("italic"), underline: () => exec("underline"), strike: () => exec("strikeThrough"),
      sub: () => exec("subscript"), sup: () => exec("superscript"),
      color: (b, a) => setColor(a || st.textColor),
      hilite: (b, a) => setHilite(a === "none" ? null : (a || st.hiColor)),
      grow: () => growFont(1), shrink: () => growFont(-1),
      clear: () => clearFormatting(false),
      ul: () => list("ul"), ol: () => list("ol"), indent: () => indent(1), outdent: () => indent(-1),
      marks: () => { st.marks = !st.marks; page.classList.toggle("wd-marks", st.marks); updateTabs(); },
      al: (b, a) => align(a),
      style: (b, a) => applyStyle(a),
      find: () => { if (!st.nav) toggleNav(true); const i = q(".wd-nav-q"); i.focus(); i.select(); const r = curRange(); if (r && !r.collapsed && r.toString().length < 60) { i.value = r.toString(); navSearch(i.value); } },
      replace: () => openFind(true),
      findClose: () => closeFind(), findNext: () => gotoMatch(1), findPrev: () => gotoMatch(-1), replOne, replAll,
      selall: () => { const ed = edActive(); ed.focus(); const r = document.createRange(); r.selectNodeContents(ed); const s = window.getSelection(); s.removeAllRanges(); s.addRange(r); },
      pagebreak: insertPageBreak,
      tableDlg: async () => {
        const r0 = restore().cloneRange();
        const v = await ask("Insérer un tableau", [{ name: "c", label: "Nombre de colonnes", type: "number", value: 3, min: 1, max: 12 }, { name: "r", label: "Nombre de lignes", type: "number", value: 3, min: 1, max: 50 }], "OK");
        refocus(r0);
        if (v) insertTable(parseInt(v.r, 10) || 1, parseInt(v.c, 10) || 1);
      },
      link: addLink, comment: addComment, footnote: addFootnote,
      header: () => hfEdit("hdr"), footer: () => hfEdit("ftr"), hfClose: () => hfEdit(null),
      gotoHdr: () => { hfEdit("hdr"); }, gotoFtr: () => { hfEdit("ftr"); },
      pagenum: insertPageNum, date: insertDate,
      cols: (b, a) => setColumns(+a),
      margins: (b, a) => { st.margins = a; applyPage(); if (!st.userZoom) fitWidth(); after(); },
      orient: (b, a) => { st.orient = a; applyPage(); if (!st.userZoom) fitWidth(); after(); },
      toc: insertToc, tocUpdate: updateToc,
      stats: showStats,
      cmtDel: () => delComment(), cmtDelI: (b, a) => delComment(+a), cmtPrev: () => cmtStep(-1), cmtNext: () => cmtStep(1),
      cmtShow: () => { st.showCmts = !st.showCmts; st.cmtSig = ""; renderCmts(); updateTabs(); if (st.showCmts && !commentMarks().length) toast("Aucun commentaire dans ce document."); },
      navToggle: () => toggleNav(), rulerToggle: () => { st.ruler = !st.ruler; q(".wd-rulerbar").hidden = !st.ruler; updateTabs(); },
      zoom100: () => setZoom(1, true), zoomPage: fitPage, zoomWidth: () => fitWidth(true),
      zoomIn: () => setZoom(st.zoom + 0.1, true), zoomOut: () => setZoom(st.zoom - 0.1, true),
      zoomDlg: async () => { const v = await ask("Zoom", [{ name: "z", label: "Pourcentage (50 à 200)", type: "number", value: Math.round(st.zoom * 100), min: 50, max: 200 }], "OK"); if (v) setZoom((parseInt(v.z, 10) || 100) / 100, true); restore(); },
      rowA: () => tblOp("rowA"), rowB: () => tblOp("rowB"), colL: () => tblOp("colL"), colR: () => tblOp("colR"),
      delRow: () => tblOp("delRow"), delCol: () => tblOp("delCol"), delTbl: () => tblOp("delTbl"), hdrRow: () => tblOp("hdrRow"),
      tstyle: (b, a) => tableStyle(a),
      bsClose: closeBackstage, download, resetDoc: () => { closeBackstage(); reset(); toast("Document réinitialisé."); }
    };
    function toggleNav(force) {
      st.nav = force == null ? !st.nav : !!force;
      nav.classList.toggle("off", !st.nav);
      updateTabs();
      later(() => { if (!st.userZoom) fitWidth(); positionCards(); }, 320);
    }

    /* ---------------------------------------------------------------- événements */
    // le ruban ne doit pas voler le focus / la sélection du document
    on(app, "mousedown", e => {
      const b = e.target.closest(".wd-btn,.wd-sty,.wd-pop button,.wd-pop i,.wd-tab,.wd-tsearch,.wd-tbtn,.wd-zb,.wd-zl,.wd-st-i,.wd-auto,.wd-toc-up");
      if (b && !b.closest(".wd-modal,.wd-findp,.wd-bs,.wd-cmts .wd-card-t")) e.preventDefault();
      if (b) ripple(b.classList.contains("wd-sty") || b.classList.contains("wd-btn") || b.classList.contains("wd-tab") ? b : null, e);
    });
    on(app, "click", e => {
      const t = e.target;
      const tab = t.closest(".wd-tab");
      if (tab) { setTab(tab.dataset.tab); return; }
      const mb = t.closest("[data-menu]");
      if (mb) { openMenu(mb, mb.dataset.menu); return; }
      const sty = t.closest(".wd-sty");
      if (sty) { applyStyle(sty.dataset.style); return; }
      const bsi = t.closest("[data-bs]");
      if (bsi) { bsSection(bsi.dataset.bs); return; }
      const nh = t.closest(".wd-nav-i,[data-h]");
      if (nh && (nh.closest(".wd-nav") || nh.closest(".o-toc"))) {
        e.preventDefault();
        const h = headingEls()[+nh.dataset.h];
        if (h) { if (st.hf) hfEdit(null); scrollToEl(h); try { root.focus({ preventScroll: true }); } catch (x) { /* */ } caretIn(h); schedState(); }
        return;
      }
      const nr = t.closest(".wd-nav-r");
      if (nr) { st.cur = +nr.dataset.r; paintMatches(); const r = st.matches[st.cur]; if (r) { scrollToRange(r); try { root.focus({ preventScroll: true }); } catch (x) { /* */ } const s = window.getSelection(); s.removeAllRanges(); s.addRange(r.cloneRange()); st.lastRange = r.cloneRange(); } qa(".wd-nav-r").forEach(x => x.classList.toggle("on", x === nr)); return; }
      const nt = t.closest("[data-ntab]");
      if (nt) { setNavTab(nt.dataset.ntab); return; }
      const card = t.closest(".wd-card");
      if (card && !t.closest("[data-cmd]")) { const i = +card.dataset.i; activateCmt(i); return; }
      const cb = t.closest("[data-cmd]");
      if (cb && app.contains(cb) && !cb.matches("select")) {
        const fn = CMD[cb.dataset.cmd];
        if (fn) {
          if (cb.closest(".wd-pop")) closeMenu();
          if (!bs.hidden && !cb.closest(".wd-bs")) closeBackstage();
          try { fn(cb, cb.dataset.arg); } catch (err) { if (window.console) console.warn("[sim-doc]", err); }
        }
        return;
      }
      if (t.closest("a[href]") && inEd(t) && (e.ctrlKey || e.metaKey)) {
        const href = t.closest("a").getAttribute("href");
        if (href && href !== "#" && !/^javascript:/i.test(href)) window.open(href, "_blank", "noopener");
      }
    });
    on(app, "change", e => {
      const t = e.target;
      if (t.matches(".wd-font")) setFont(t.value);
      else if (t.matches(".wd-size")) fontSizePt(parseFloat(t.value));
      else if (t.matches(".wd-fc")) { st.findCase = t.checked; runFind(q(".wd-fq").value); }
    });
    on(app, "input", e => {
      const t = e.target;
      if (t === root || root.contains(t)) { changed(); if (e.inputType && /^insert/.test(e.inputType)) caretPulse(); return; }
      if (t === hdr || t === ftr || hdr.contains(t) || ftr.contains(t)) { changed(); caretPulse(); return; }
      if (t.matches(".wd-zr")) { setZoom(+t.value / 100, true); return; }
      if (t.matches(".wd-fq")) { runFind(t.value); return; }
      if (t.matches(".wd-nav-q")) { navSearch(t.value); return; }
      if (t.matches(".wd-card-t")) { const m = commentMarks()[+t.dataset.ci]; if (m) { m.setAttribute("data-note", t.textContent); st.cmtSig = ""; changed(); } return; }
      if (t.matches(".wd-fn-t")) { const s = root.querySelectorAll("sup.o-fn")[+t.dataset.fi]; if (s) { s.setAttribute("data-note", t.textContent); st.fnSig = ""; changed(); } }
    });
    on(app, "focusout", e => { if (e.target.matches && e.target.matches(".wd-card-t,.wd-fn-t")) later(() => { st.cmtSig = ""; st.fnSig = ""; renderCmts(); renderFn(); }, 50); });
    on(app, "mouseover", e => {
      const s = e.target.closest(".wd-sty");
      if (s && !s.contains(e.relatedTarget)) preview(s.dataset.style);
    });
    on(app, "mouseout", e => {
      const s = e.target.closest(".wd-sty");
      if (s && !s.contains(e.relatedTarget)) clearPreview();
    });
    on(app, "keydown", e => {
      const t = e.target;
      if (t.matches && t.matches(".wd-fq,.wd-fr") && e.key === "Enter") { e.preventDefault(); if (t.matches(".wd-fr")) replOne(); else gotoMatch(e.shiftKey ? -1 : 1); return; }
      if (t.matches && t.matches(".wd-nav-q") && e.key === "Enter") { e.preventDefault(); if (st.matches.length) { st.cur = (st.cur + 1) % st.matches.length; paintMatches(); scrollToRange(st.matches[st.cur]); } return; }
      if (t.matches && t.matches(".wd-card-t,.wd-fn-t") && e.key === "Enter") { e.preventDefault(); t.blur(); return; }
      if (e.key === "Escape") {
        if (st.modal) return;
        if (st.pop) { closeMenu(); e.preventDefault(); return; }
        if (!bs.hidden) { closeBackstage(); return; }
        if (!findp.hidden) { closeFind(); return; }
        if (st.hf) { hfEdit(null); return; }
        return;
      }
      const ed = edActive();
      if (!(t === ed || ed.contains(t))) {
        if ((e.ctrlKey || e.metaKey) && !e.altKey && (e.key === "z" || e.key === "y") && !t.matches("input,textarea,[contenteditable=true]")) { e.preventDefault(); e.key === "z" && !e.shiftKey ? undo() : redo(); }
        return;
      }
      onEditorKey(e, ed);
    });
    function onEditorKey(e, ed) {
      const mod = e.ctrlKey || e.metaKey;
      const k = (e.key || "").toLowerCase();
      const alt = e.altKey && !(e.getModifierState && e.getModifierState("AltGraph"));
      const stop = () => { e.preventDefault(); e.stopPropagation(); };
      if (mod && alt) {
        if (e.code === "Digit1" || k === "1" || k === "&") { stop(); return applyStyle("h1"); }
        if (e.code === "Digit2" || k === "2" || k === "é") { stop(); return applyStyle("h2"); }
        if (e.code === "Digit3" || k === "3" || k === '"') { stop(); return applyStyle("h3"); }
        if (e.code === "KeyM") { stop(); return addComment(); }
        if (e.code === "KeyF") { stop(); return addFootnote(); }
        return;
      }
      if (mod && !e.altKey) {
        if (e.shiftKey && (k === "n")) { stop(); return applyStyle("p"); }
        if (e.shiftKey && (k === "l")) { stop(); return list("ul"); }
        if (e.shiftKey && (k === "z")) { stop(); return redo(); }
        if (e.shiftKey && (k === ">" || k === "." || e.code === "Period")) { stop(); return growFont(1); }
        if (e.shiftKey && (k === "+" || k === "=")) { stop(); return exec("superscript"); }
        if (e.shiftKey && (k === "*" || k === "8")) { stop(); return CMD.marks(); }
        if (e.shiftKey) return;
        switch (k) {
          case "b": case "g": stop(); return exec("bold");
          case "i": stop(); return exec("italic");
          case "u": stop(); return exec("underline");
          case "e": stop(); return align("center");
          case "l": stop(); return align("left");
          case "r": stop(); return align("right");
          case "j": stop(); return align("justify");
          case "z": stop(); return undo();
          case "y": stop(); return redo();
          case "f": stop(); return CMD.find();
          case "h": stop(); return openFind(true);
          case "k": stop(); return addLink();
          case "s": stop(); return CMD.save();
          case "<": stop(); return growFont(-1);
          case "=": stop(); return exec("subscript");
          case " ": stop(); return clearFormatting(true);
          case "enter": stop(); return insertPageBreak();
          case "p": stop(); return openBackstage("print");
        }
        return;
      }
      if (e.key === "F9") { e.preventDefault(); return updateToc(); }
      if (e.key === "Tab") {
        const r = curRange(); if (!r) return;
        const td = closestIn(r.startContainer, "td,th", ed);
        if (td) {
          e.preventDefault();
          const table = td.closest("table"), cells = [...table.querySelectorAll("td,th")];
          let nx = cells[cells.indexOf(td) + (e.shiftKey ? -1 : 1)];
          if (!nx && !e.shiftKey) { const tr = document.createElement("tr"); [...td.parentElement.cells].forEach(() => { const c = document.createElement("td"); c.innerHTML = "<br>"; tr.appendChild(c); }); td.parentElement.after(tr); nx = tr.cells[0]; }
          if (nx) { const rr = document.createRange(); rr.selectNodeContents(nx); const s = window.getSelection(); s.removeAllRanges(); s.addRange(rr); }
          after(); return;
        }
        if (closestIn(r.startContainer, "li", ed)) { e.preventDefault(); return exec(e.shiftKey ? "outdent" : "indent"); }
        e.preventDefault();
        if (!e.shiftKey) { document.execCommand("insertText", false, "      "); after(); }
        return;
      }
      if (e.key === "Enter" && !e.shiftKey && ed === root) {
        const r = curRange(); if (!r) return;
        const b = blockOf(r.startContainer, ed);
        if (b && /^(H[1-6]|BLOCKQUOTE)$/.test(b.tagName) && !closestIn(b, "li,td,th", ed)) {
          e.preventDefault();
          if (!r.collapsed) r.deleteContents();
          const tail = document.createRange(); tail.setStart(r.startContainer, r.startOffset); tail.setEnd(b, b.childNodes.length);
          const atEnd = !tail.toString().trim();
          const frag = tail.extractContents();
          let nb;
          if (atEnd) { nb = newP(); if (b.style.textAlign && b.tagName !== "BLOCKQUOTE") nb.style.textAlign = b.style.textAlign; }
          else { nb = document.createElement(b.tagName); if (b.getAttribute("style")) nb.setAttribute("style", b.getAttribute("style")); nb.appendChild(frag); }
          b.after(nb);
          if (isBlank(b) && !b.querySelector("br")) b.innerHTML = "<br>";
          if (isBlank(nb) && !nb.querySelector("br")) nb.innerHTML = "<br>";
          caretIn(nb);
          after();
        }
      }
    }
    on(app, "dblclick", e => {
      const t = e.target;
      if (hdr.contains(t) && st.hf !== "hdr") { e.preventDefault(); hfEdit("hdr"); }
      else if (ftr.contains(t) && st.hf !== "ftr") { e.preventDefault(); hfEdit("ftr"); }
      else if (st.hf && root.contains(t)) { e.preventDefault(); hfEdit(null); }
    });
    on(app, "focusin", e => { if (st.hf) { if (hdr.contains(e.target)) st.hf = "hdr"; else if (ftr.contains(e.target)) st.hf = "ftr"; } });
    on(root, "paste", e => {
      const cd = e.clipboardData; if (!cd) return;
      const html = cd.getData("text/html"), txt = cd.getData("text/plain");
      e.preventDefault();
      if (html) document.execCommand("insertHTML", false, cleanPaste(html));
      else if (txt) document.execCommand("insertText", false, txt);
      after();
    });
    [hdr, ftr].forEach(z => on(z, "paste", e => { const cd = e.clipboardData; if (!cd) return; e.preventDefault(); document.execCommand("insertText", false, cd.getData("text/plain")); after(); }));
    on(canvas, "wheel", e => { if (e.ctrlKey) { e.preventDefault(); setZoom(st.zoom + (e.deltaY < 0 ? 0.1 : -0.1), true); } }, { passive: false });
    on(canvas, "scroll", () => { if (st.pop) closeMenu(); });
    on(canvas, "mousedown", e => { if (e.target === canvas || e.target.classList.contains("wd-scrollin") || e.target === sizer) { e.preventDefault(); restore(); } });
    on(document, "selectionchange", () => {
      if (st.dead) return;
      const s = window.getSelection();
      if (!s || !s.rangeCount) return;
      const r = s.getRangeAt(0);
      if (edActive().contains(r.commonAncestorContainer)) { st.lastRange = r.cloneRange(); schedState(); }
    });
    on(document, "pointerdown", e => { if (st.pop && !st.pop.contains(e.target) && !(st.popBtn && st.popBtn.contains(e.target))) closeMenu(); }, true);
    let ro = null;
    if (window.ResizeObserver) {
      ro = new ResizeObserver(() => { if (st.dead) return; sizeStage(); if (!st.userZoom) fitWidth(); updateTabs(); positionCards(); });
      ro.observe(canvas); ro.observe(page);
    } else on(window, "resize", () => { sizeStage(); fitWidth(); updateTabs(); });

    /* ---------------------------------------------------------------- reset / init */
    function reset() {
      st.hf = null; page.classList.remove("wd-hfmode"); root.contentEditable = "true"; hdr.contentEditable = ftr.contentEditable = "false";
      q('.wd-tab[data-tab="hf"]').hidden = true; q('.wd-tab[data-tab="tbl"]').hidden = true;
      clearPreview(); closeMenu(); if (!findp.hidden) { findp.hidden = true; findp.classList.remove("wd-in"); } clearMatches();
      root.innerHTML = cleanHTML(setup.html || "<p><br></p>");
      hdr.innerHTML = cleanHTML(setup.header || "");
      ftr.innerHTML = cleanHTML(setup.footer || "");
      root.querySelectorAll(".o-toc").forEach(t => { if (!t.getAttribute("data-sig")) fillToc(t); });
      root.querySelectorAll("sup.o-fn").forEach(s => { s.contentEditable = "false"; });
      st.columns = clamp(parseInt(setup.columns, 10) || 1, 1, 3); st.margins = "normal"; st.orient = "portrait";
      st.activeCmt = -1; st.cmtSig = st.fnSig = st.navSig = "";
      applyPage();
      st.hist = []; st.hidx = -1;
      normalizeDoc();
      st.lastRange = null;
      const first = root.querySelector("p,h1,h2,h3,li,blockquote,td");
      if (first) { const r = document.createRange(); r.selectNodeContents(first); r.collapse(true); st.lastRange = r; }
      canvas.scrollTop = 0;
      if (st.tab === "hf" || st.tab === "tbl") setTab("home");
      refresh(true);
      try { ctx.changed && ctx.changed(); } catch (e) { /* */ }
    }
    function destroy() {
      st.dead = true;
      listeners.forEach(([t, ev, fn, opt]) => t.removeEventListener(ev, fn, opt));
      listeners.length = 0;
      timers.forEach(clearTimeout); timers.clear();
      if (stateRaf) cancelAnimationFrame(stateRaf);
      if (ro) ro.disconnect();
      clearMatches();
      if (app.parentNode) app.parentNode.removeChild(app);
    }

    reset();
    st.nav = host.clientWidth === 0 || host.clientWidth >= 1000;
    nav.classList.toggle("off", !st.nav);
    setTab("home");
    requestAnimationFrame(() => { if (!st.dead) { fitWidth(); updateTabs(); } });

    /* ---------------------------------------------------------------- API des check */
    const api = {
      get root() { return root; },
      $: sel => { try { return root.querySelector(sel); } catch (e) { return null; } },
      $$: sel => { try { return [...root.querySelectorAll(sel)]; } catch (e) { return []; } },
      count: sel => { try { return root.querySelectorAll(sel).length; } catch (e) { return 0; } },
      text: () => textOf(root),
      words: () => countWords(textOf(root)),
      has: t => { const n = normS(t).toLowerCase(); return !!n && normS(textOf(root)).toLowerCase().includes(n); },
      fmt: frag => fmtOf(frag),
      headings: () => headingEls().map(h => ({ level: +h.tagName[1], text: normS(h.textContent) })),
      get header() { return normS(textOf(hdr)); },
      get footer() { return normS(textOf(ftr)); },
      get columns() { return st.columns; },
      get toc() { return !!root.querySelector(".o-toc"); },
      get tocFresh() { const t = root.querySelector(".o-toc"); return !!t && t.getAttribute("data-sig") === tocSig(); },
      tables: () => tablesInfo()
    };

    return { api, reset, destroy, exportFile };
  });
})();
