/* Office Academy — simulateur PowerPoint (kind "slides")
   Vanilla JS, aucune dépendance. S'enregistre via OFFICE.sim("slides", mount).
   Valeurs exposées à l'API des check :
     layout     : "title" | "content" | "two" | "section" | "blank"
     transition : "none" | "fade" (Fondu) | "push" (Poussée) | "zoom" (Zoom) | "morph" (Morphose)
     anim       : "none" | "appear" (Apparition) | "fly" (Entrée) | "zoom" (Zoom) | "bounce" (Rebond)
     shape      : "rect" | "round" | "ellipse" | "arrow"  ("" si l'élément n'est pas une forme)
     color/fill : "#rrggbb" ou "" (= couleur du thème)          size : taille de police (pt, diapo 960 px)
*/
(function () {
  "use strict";
  if (!window.OFFICE || typeof window.OFFICE.sim !== "function") return;

  /* ---------- constantes ---------- */
  const THEMES = { blanc: "Blanc", nuit: "Nuit", ocean: "Océan", sunset: "Sunset", foret: "Forêt", neon: "Néon" };
  const LAYOUTS = { title: "Diapositive de titre", content: "Titre et contenu", two: "Deux contenus", section: "En-tête de section", blank: "Vide" };
  const TRANS = [["none", "Aucune"], ["fade", "Fondu"], ["push", "Poussée"], ["zoom", "Zoom"], ["morph", "Morphose"]];
  const ANIMS = [["none", "Aucune"], ["appear", "Apparition"], ["fly", "Entrée"], ["zoom", "Zoom"], ["bounce", "Rebond"]];
  const TRANSK = TRANS.map(t => t[0]), ANIMK = ANIMS.map(a => a[0]);
  const TYPES = ["title", "subtitle", "body", "text", "shape", "icon"];
  const SHAPES = { rect: "Rectangle", round: "Rectangle arrondi", ellipse: "Ellipse", arrow: "Flèche" };
  const SIZES = [10, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 96, 120];
  const COLORS = ["#000000", "#ffffff", "#1f2937", "#64748b", "#c43e1c", "#e11d48", "#f97316", "#f59e0b",
    "#16a34a", "#0d9488", "#0ea5e9", "#2563eb", "#7c3aed", "#db2777"];
  const ICONS = ["🚀", "💡", "📈", "📊", "✅", "⭐", "🎯", "📅", "👥", "💬", "🔒", "⚙️", "🏆", "📦", "💰", "🌍", "❤️", "⚠️", "🛒", "🎉"];
  const DEF = {
    title: { x: 6, y: 5, w: 88, h: 16, size: 40, bold: true, align: "left" },
    subtitle: { x: 15, y: 54, w: 70, h: 14, size: 26, bold: false, align: "center" },
    body: { x: 6, y: 24, w: 88, h: 68, size: 24, bold: false, align: "left" },
    text: { x: 30, y: 43, w: 40, h: 14, size: 24, bold: false, align: "left" },
    shape: { x: 38, y: 33, w: 24, h: 34, size: 22, bold: true, align: "center" },
    icon: { x: 44, y: 38, w: 12, h: 24, size: 80, bold: false, align: "center" }
  };
  const PH = {
    title: [["title", { x: 8, y: 26, w: 84, h: 24, size: 54, bold: true, align: "center" }],
      ["subtitle", { x: 15, y: 54, w: 70, h: 14, size: 26, align: "center" }]],
    content: [["title", { x: 6, y: 5, w: 88, h: 16, size: 40, bold: true, align: "left" }],
      ["body", { x: 6, y: 25, w: 88, h: 67, size: 24, align: "left" }]],
    two: [["title", { x: 6, y: 5, w: 88, h: 16, size: 40, bold: true, align: "left" }],
      ["body", { x: 6, y: 25, w: 42.5, h: 67, size: 22, align: "left" }],
      ["body", { x: 51.5, y: 25, w: 42.5, h: 67, size: 22, align: "left" }]],
    section: [["title", { x: 8, y: 36, w: 84, h: 22, size: 50, bold: true, align: "center" }]],
    blank: []
  };
  const PHTXT = { title: "Cliquez pour ajouter un titre", subtitle: "Cliquez pour ajouter un sous-titre",
    body: "Cliquez pour ajouter du texte", text: "Saisissez du texte", shape: "", icon: "" };

  /* ---------- icônes (SVG 20×20) ---------- */
  const sv = (p, fill) => `<svg viewBox="0 0 20 20" aria-hidden="true" fill="${fill ? "currentColor" : "none"}" stroke="${fill ? "none" : "currentColor"}" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round">${p}</svg>`;
  const I = {
    newslide: sv('<rect x="2.5" y="4" width="15" height="11.5" rx="1.6"/><path d="M10 7.3v5M7.5 9.8h5"/>'),
    layout: sv('<rect x="2.5" y="3.5" width="15" height="13" rx="1.6"/><path d="M2.5 7.5h15M10 7.5v9"/>'),
    dup: sv('<rect x="6.5" y="6.5" width="11" height="9" rx="1.5"/><path d="M3.5 12.5V5A1.5 1.5 0 0 1 5 3.5h8.5"/>'),
    trash: sv('<path d="M4 6h12M8 6V4h4v2M5.5 6l.8 10.5h7.4l.8-10.5"/>'),
    up: sv('<path d="M10 15.5v-11M5.5 9 10 4.5 14.5 9"/>'),
    down: sv('<path d="M10 4.5v11M5.5 11 10 15.5 14.5 11"/>'),
    chev: sv('<path d="M6 8l4 4 4-4"/>'),
    tl: sv('<path d="M3 4.5h14M3 8.5h9M3 12.5h14M3 16.5h9"/>'),
    tc: sv('<path d="M3 4.5h14M5.5 8.5h9M3 12.5h14M5.5 16.5h9"/>'),
    tr: sv('<path d="M3 4.5h14M8 8.5h9M3 12.5h14M8 16.5h9"/>'),
    oL: sv('<path d="M3 2.5v15"/><rect x="5.5" y="5" width="10" height="4" rx="1"/><rect x="5.5" y="11" width="6" height="4" rx="1"/>'),
    oC: sv('<path d="M10 2.5v15"/><rect x="4" y="5" width="12" height="4" rx="1"/><rect x="6.5" y="11" width="7" height="4" rx="1"/>'),
    oR: sv('<path d="M17 2.5v15"/><rect x="4.5" y="5" width="10" height="4" rx="1"/><rect x="8.5" y="11" width="6" height="4" rx="1"/>'),
    oT: sv('<path d="M2.5 3h15"/><rect x="5" y="5.5" width="4" height="10" rx="1"/><rect x="11" y="5.5" width="4" height="6" rx="1"/>'),
    oM: sv('<path d="M2.5 10h15"/><rect x="5" y="4" width="4" height="12" rx="1"/><rect x="11" y="6.5" width="4" height="7" rx="1"/>'),
    oB: sv('<path d="M2.5 17h15"/><rect x="5" y="4.5" width="4" height="10" rx="1"/><rect x="11" y="8.5" width="4" height="6" rx="1"/>'),
    front: sv('<rect x="3" y="3" width="9" height="9" rx="1.2"/><rect x="8" y="8" width="9" height="9" rx="1.2" fill="currentColor" fill-opacity=".45"/>'),
    back: sv('<rect x="8" y="8" width="9" height="9" rx="1.2"/><rect x="3" y="3" width="9" height="9" rx="1.2" fill="currentColor" fill-opacity=".45"/>'),
    textbox: sv('<rect x="2.5" y="3.5" width="15" height="13" rx="1" stroke-dasharray="2.2 2"/><path d="M6.5 7.5h7M10 7.5v6"/>'),
    rect: sv('<rect x="3" y="5" width="14" height="10" rx=".5"/>'),
    round: sv('<rect x="3" y="5" width="14" height="10" rx="3.5"/>'),
    ellipse: sv('<ellipse cx="10" cy="10" rx="7.5" ry="5.5"/>'),
    arrow: sv('<path d="M2.5 7.8h8.7V4.5L17.5 10l-6.3 5.5v-3.3H2.5z"/>'),
    smile: sv('<circle cx="10" cy="10" r="7.5"/><path d="M7 11.8c1.6 1.9 4.4 1.9 6 0M7.6 8h.01M12.4 8h.01"/>'),
    play: sv('<path d="M6.5 4.2v11.6L16 10z"/>', 1),
    playc: sv('<rect x="2.5" y="4" width="15" height="11" rx="1.6"/><path d="M8.5 7.3v5.4L13 10z" fill="currentColor"/>'),
    undo: sv('<path d="M7.5 4.5 3.5 8.5l4 4"/><path d="M3.5 8.5H12a4.5 4.5 0 0 1 0 9H9"/>'),
    redo: sv('<path d="M12.5 4.5l4 4-4 4"/><path d="M16.5 8.5H8a4.5 4.5 0 0 0 0 9h3"/>'),
    notes: sv('<rect x="3" y="3" width="14" height="14" rx="1.6"/><path d="M6 7h8M6 10h8M6 13h5"/>'),
    fill: sv('<path d="M4 10.5 9.5 5l5 5L9 15.5z"/><path d="M15.8 12.8s1.6 2 1.6 3.1a1.6 1.6 0 0 1-3.2 0c0-1.1 1.6-3.1 1.6-3.1z"/>'),
    star: sv('<path d="M10 2.3l2.4 4.9 5.3.8-3.9 3.7.9 5.3L10 14.5 5.3 17l.9-5.3L2.3 8l5.3-.8z"/>', 1),
    eye: sv('<path d="M1.5 10S4.5 4.5 10 4.5 18.5 10 18.5 10 15.5 15.5 10 15.5 1.5 10 1.5 10z"/><circle cx="10" cy="10" r="2.5"/>'),
    all: sv('<rect x="2.5" y="2.5" width="6.5" height="6.5" rx="1"/><rect x="11" y="2.5" width="6.5" height="6.5" rx="1"/><rect x="2.5" y="11" width="6.5" height="6.5" rx="1"/><rect x="11" y="11" width="6.5" height="6.5" rx="1"/>'),
    sun: sv('<circle cx="10" cy="10" r="3.4"/><path d="M10 2v1.8M10 16.2V18M2 10h1.8M16.2 10H18M4.3 4.3l1.3 1.3M14.4 14.4l1.3 1.3M4.3 15.7l1.3-1.3M14.4 5.6l1.3-1.3"/>'),
    search: sv('<circle cx="9" cy="9" r="5.5"/><path d="M13 13l4 4"/>'),
    sel: sv('<path d="M4 3l11 6.5-5 1.2-2.4 4.8z"/>')
  };

  /* ---------- utilitaires ---------- */
  const clamp = (v, a, b) => Math.min(b, Math.max(a, v));
  const r2 = v => Math.round(v * 100) / 100;
  const r1 = v => Math.round(v * 10) / 10;
  const norm = s => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase();
  const fold = s => String(s || "").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").trim();
  const reduced = () => !!(window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches);
  const canAnim = typeof Element !== "undefined" && typeof Element.prototype.animate === "function";
  function h(tag, cls, html) { const d = document.createElement(tag); if (cls) d.className = cls; if (html != null) d.innerHTML = html; return d; }
  const isTextType = t => t === "title" || t === "subtitle" || t === "body" || t === "text";
  const hiddenInShow = e => (isTextType(e.type) || e.type === "icon") && !String(e.text || "").trim();

  /* ---------- ruban : petits constructeurs HTML ---------- */
  function B(cmd, icon, label, title, o) {
    o = o || {};
    const t = (title || label || "").replace(/<br>/g, " ").replace(/"/g, "&quot;");
    return `<button type="button" class="sl-b${o.big ? " sl-big" : ""}${o.cls ? " " + o.cls : ""}" data-cmd="${cmd}"` +
      (o.arg != null ? ` data-arg="${o.arg}"` : "") + (o.need ? ` data-need="${o.need}"` : "") +
      (o.prev ? ` data-prev="${o.prev}"` : "") + (o.pop ? ` data-pop="1" aria-haspopup="true"` : "") +
      ` title="${t}" aria-label="${t}">${icon || ""}${label ? `<span>${label}</span>` : ""}</button>`;
  }
  const G = (label, inner, cls) => `<div class="sl-grp${cls ? " " + cls : ""}"><div class="sl-gin">${inner}</div><div class="sl-glab">${label}</div></div>`;
  const COL = s => `<div class="sl-col">${s}</div>`;
  const ROW = s => `<div class="sl-row">${s}</div>`;
  const SEP = '<span class="sl-vsep"></span>';

  function ribbonHTML() {
    const newSl = `<div class="sl-split">${B("newSlide", I.newslide, "Nouvelle<br>diapositive", "Nouvelle diapositive (Ctrl+M)", { big: 1 })}${B("newSlideMenu", I.chev, "", "Choisir la disposition de la nouvelle diapositive", { cls: "sl-splitdd", pop: 1 })}</div>`;
    const home =
      G("Diapositives", newSl + COL(B("layoutMenu", I.layout, "Disposition", "Disposition de la diapositive", { pop: 1 }) +
        B("dupSlide", I.dup, "Dupliquer", "Dupliquer la diapositive (Ctrl+D sans sélection)") +
        B("delSlide", I.trash, "Supprimer", "Supprimer la diapositive"))) +
      G("Police", ROW(`<select class="sl-size" data-need="el" title="Taille de police" aria-label="Taille de police">${SIZES.map(s => `<option value="${s}">${s}</option>`).join("")}</select>` +
        B("grow", '<b class="sl-gl">A<sup>+</sup></b>', "", "Augmenter la taille de police", { need: "el" }) +
        B("shrink", '<b class="sl-gl">A<sup>−</sup></b>', "", "Réduire la taille de police", { need: "el" })) +
        ROW(B("bold", '<b class="sl-gl sl-gb">G</b>', "", "Gras (Ctrl+G)", { need: "el" }) +
          B("colorMenu", '<span class="sl-acol"><b>A</b><i></i></span>', "", "Couleur de police", { need: "el", pop: 1 }) + SEP +
          B("alignL", I.tl, "", "Aligner le texte à gauche", { need: "el" }) +
          B("alignC", I.tc, "", "Centrer le texte", { need: "el" }) +
          B("alignR", I.tr, "", "Aligner le texte à droite", { need: "el" }))) +
      G("Dessin", COL(B("fillMenu", I.fill, "Remplissage", "Remplissage de la forme", { need: "el", pop: 1 }) +
        B("dupEl", I.dup, "Dupliquer", "Dupliquer l'élément (Ctrl+D)", { need: "el" }) +
        B("delEl", I.trash, "Supprimer", "Supprimer l'élément (Suppr)", { need: "el" }))) +
      G("Aligner sur la diapositive", `<div class="sl-grid3">` +
        B("objL", I.oL, "", "Aligner à gauche", { need: "el" }) + B("objC", I.oC, "", "Centrer horizontalement", { need: "el" }) +
        B("objR", I.oR, "", "Aligner à droite", { need: "el" }) + B("objT", I.oT, "", "Aligner en haut", { need: "el" }) +
        B("objM", I.oM, "", "Aligner au milieu", { need: "el" }) + B("objB", I.oB, "", "Aligner en bas", { need: "el" }) +
        `</div>` + COL(B("front", I.front, "Premier plan", "Mettre au premier plan", { need: "el" }) +
          B("back", I.back, "Arrière-plan", "Mettre à l'arrière-plan", { need: "el" }))) +
      G("Édition", COL(B("undo", I.undo, "Annuler", "Annuler (Ctrl+Z)") + B("redo", I.redo, "Rétablir", "Rétablir (Ctrl+Y)")));

    const ins =
      G("Diapositives", newSl) +
      G("Texte", B("textbox", I.textbox, "Zone de<br>texte", "Insérer une zone de texte", { big: 1 })) +
      G("Formes", `<div class="sl-grid2">` + Object.keys(SHAPES).map(k => B("shape", I[k], SHAPES[k].replace("Rectangle arrondi", "Arrondi"), "Insérer : " + SHAPES[k], { arg: k })).join("") + `</div>`) +
      G("Illustrations", B("iconMenu", I.smile, "Icônes", "Insérer une icône", { big: 1, pop: 1 }));

    const themeCards = Object.keys(THEMES).map(k =>
      `<button type="button" class="sl-thc" data-cmd="theme" data-arg="${k}" title="Thème ${THEMES[k]}" aria-label="Thème ${THEMES[k]}">` +
      `<span class="sl-slide sl-th-${k} sl-l-title"><span class="sl-deco"></span><span class="sl-el sl-t-title" style="left:6%;top:18%;width:88%;height:44%;--fs:170;text-align:center"><span class="sl-txt">Aa</span></span>` +
      `<span class="sl-thbar"></span></span><span class="sl-thn">${THEMES[k]}</span></button>`).join("");
    const design = G("Thèmes", `<div class="sl-gal">${themeCards}</div>`) +
      G("Personnaliser", B("layoutMenu", I.layout, "Disposition", "Disposition de la diapositive", { big: 1, pop: 1 }));

    const transGal = TRANS.map(([k, l]) =>
      `<button type="button" class="sl-b sl-gi" data-cmd="trans" data-arg="${k}" data-prev="trans:${k}" title="Transition : ${l}" aria-label="Transition : ${l}">` +
      `<span class="sl-gic sl-gt-${k}"><i class="a"></i><i class="b"></i></span><span>${l}</span></button>`).join("");
    const trans = G("Transition vers cette diapositive", `<div class="sl-gal">${transGal}</div>`) +
      G("Aperçu", B("transPreview", I.eye, "Aperçu", "Aperçu de la transition", { big: 1 })) +
      G("Minutage", B("transAll", I.all, "Appliquer<br>partout", "Appliquer cette transition à toutes les diapositives", { big: 1 }));

    const animGal = ANIMS.map(([k, l]) =>
      `<button type="button" class="sl-b sl-gi" data-cmd="anim" data-arg="${k}" data-prev="anim:${k}" title="Animation : ${l}" aria-label="Animation : ${l}">` +
      `<span class="sl-gic sl-ga-${k}"><i class="st">${k === "none" ? "" : I.star}</i></span><span>${l}</span></button>`).join("");
    const anim = G("Animation", `<div class="sl-gal">${animGal}</div>`) +
      G("Aperçu", B("animPreview", I.eye, "Aperçu", "Aperçu des animations de la diapositive", { big: 1 })) +
      G("Minutage", ROW('<span class="sl-ordlab">Ordre : –</span>') +
        ROW(B("animEarlier", I.up, "Avancer", "Lire plus tôt", { need: "anim" }) + B("animLater", I.down, "Reculer", "Lire plus tard", { need: "anim" })));

    const show = G("Démarrage du diaporama",
      B("showStart", I.play, "À partir<br>du début", "À partir du début (F5)", { big: 1 }) +
      B("showCur", I.playc, "À partir de la<br>diapositive actuelle", "À partir de la diapositive actuelle (Maj+F5)", { big: 1 })) +
      G("Moniteurs", B("toggleNotes", I.notes, "Notes du<br>présentateur", "Afficher / masquer les notes", { big: 1 }));

    return `<div class="sl-panel" data-tab="home">${home}</div><div class="sl-panel" data-tab="insert">${ins}</div>` +
      `<div class="sl-panel" data-tab="design">${design}</div><div class="sl-panel" data-tab="trans">${trans}</div>` +
      `<div class="sl-panel" data-tab="anim">${anim}</div><div class="sl-panel" data-tab="show">${show}</div>`;
  }

  const TABS = [["file", "Fichier"], ["home", "Accueil"], ["insert", "Insertion"], ["design", "Création"],
    ["trans", "Transitions"], ["anim", "Animations"], ["show", "Diaporama"]];

  /* =====================================================================
     MONTAGE
     ===================================================================== */
  window.OFFICE.sim("slides", function mount(host, setup, ctx) {
    setup = setup || {};
    const C = Object.assign({ changed() {}, toast() {}, esc: s => String(s), confetti() {} }, ctx || {});
    let _id = 0;
    const uid = p => (p || "e") + (++_id);

    /* ---------- état ---------- */
    const S = { theme: "blanc", slides: [], cur: 0, sel: null, played: false };
    let past = [], future = [], editing = null, drag = null, active = false, show = null, preview = null;
    let tab = "home", lastDown = {}, lastNudge = {}, notesPushed = false, focusArea = "canvas";
    let raf = 0, flashId = null, enterFx = false, hoverB = null, hoverT = 0;
    const pendingThumbs = new Set();

    function normEl(e) {
      e = e || {};
      const type = TYPES.indexOf(e.type) >= 0 ? e.type : "text";
      const d = DEF[type];
      const n = k => (typeof e[k] === "number" && isFinite(e[k])) ? e[k] : (e[k] != null && e[k] !== "" && isFinite(+e[k]) ? +e[k] : d[k]);
      return {
        id: uid("e"), type, text: e.text == null ? "" : String(e.text),
        x: n("x"), y: n("y"), w: Math.max(1, n("w")), h: Math.max(1, n("h")), size: Math.max(6, n("size")),
        bold: e.bold == null ? !!d.bold : !!e.bold, color: e.color ? String(e.color) : "", fill: e.fill ? String(e.fill) : "",
        shape: type === "shape" ? (SHAPES[e.shape] ? e.shape : "rect") : "",
        anim: ANIMK.indexOf(e.anim) >= 0 ? e.anim : "none", order: e.order | 0,
        align: ["left", "center", "right"].indexOf(e.align) >= 0 ? e.align : d.align
      };
    }
    function renumber(s) {
      const list = s.els.filter(e => e.anim !== "none").sort((a, b) => (a.order || 1e9) - (b.order || 1e9));
      list.forEach((e, i) => { e.order = i + 1; });
      s.els.forEach(e => { if (e.anim === "none") e.order = 0; });
    }
    const maxOrder = s => s.els.reduce((m, e) => Math.max(m, e.anim !== "none" ? e.order : 0), 0);

    function mkSlide(src) {
      src = src || {};
      const layout = LAYOUTS[src.layout] ? src.layout : "content";
      const s = { id: uid("s"), layout, transition: TRANSK.indexOf(src.transition) >= 0 ? src.transition : "none", notes: src.notes ? String(src.notes) : "", els: [] };
      let bodies = [];
      if (layout === "two") bodies = [src.left != null ? src.left : (src.body != null ? src.body : ""), src.right != null ? src.right : ""];
      else if (src.body != null) bodies = [src.body];
      else if (src.left != null || src.right != null) bodies = [[src.left, src.right].filter(v => v != null && v !== "").join("\n")];
      let bi = 0;
      PH[layout].forEach(([type, g]) => {
        let text = type === "title" ? src.title : type === "subtitle" ? (src.subtitle != null ? src.subtitle : src.body) : bodies[bi++];
        s.els.push(normEl(Object.assign({ type }, g, { text: text == null ? "" : text })));
      });
      if (layout === "section" && src.subtitle) s.els.push(normEl({ type: "subtitle", text: src.subtitle, x: 15, y: 60, w: 70, h: 12, size: 24, align: "center" }));
      if (Array.isArray(src.els)) src.els.forEach(e => s.els.push(normEl(e)));
      renumber(s);
      return s;
    }
    function init() {
      const t = fold(setup.theme);
      S.theme = THEMES[t] ? t : "blanc";
      const src = Array.isArray(setup.slides) && setup.slides.length ? setup.slides : [{ layout: "title" }];
      S.slides = src.map(mkSlide);
      S.cur = 0; S.sel = null; S.played = false; past = []; future = [];
    }

    const cur = () => S.slides[S.cur];
    const elById = id => { const s = cur(); return s ? s.els.find(e => e.id === id) || null : null; };
    const selEl = () => (S.sel ? elById(S.sel) : null);

    /* ---------- historique ---------- */
    const snap = () => JSON.stringify({ t: S.theme, s: S.slides, c: S.cur });
    function pushHist(str) { past.push(str || snap()); if (past.length > 80) past.shift(); future = []; }
    function restore(str) {
      const o = JSON.parse(str);
      S.theme = o.t; S.slides = o.s; S.cur = clamp(o.c, 0, S.slides.length - 1);
      if (S.sel && !selEl()) S.sel = null;
    }
    function undo() {
      stopEdit();
      if (!past.length) { C.toast("Rien à annuler"); return; }
      future.push(snap()); restore(past.pop()); renderAll(); C.changed();
    }
    function redo() {
      stopEdit();
      if (!future.length) { C.toast("Rien à rétablir"); return; }
      past.push(snap()); restore(future.pop()); renderAll(); C.changed();
    }
    function mutate(fn) { stopEdit(); pushHist(); fn(); renderAll(); C.changed(); }

    /* ---------- DOM de l'application ---------- */
    host.classList.add("sl-host");
    host.textContent = "";
    const root = h("div", "sl-app sl-tab-home");
    root.tabIndex = -1;
    root.setAttribute("role", "application");
    root.setAttribute("aria-label", "Simulateur PowerPoint");
    root.innerHTML =
      `<div class="sl-titlebar"><span class="sl-logo" aria-hidden="true">P</span>` +
      `<div class="sl-qa">${B("undo", I.undo, "", "Annuler (Ctrl+Z)")}${B("redo", I.redo, "", "Rétablir (Ctrl+Y)")}${B("showStart", I.play, "", "Diaporama à partir du début (F5)")}</div>` +
      `<div class="sl-fname"><b>Présentation1</b><span> — PowerPoint</span></div>` +
      `<div class="sl-search">${I.search}<span>Rechercher</span></div>` +
      `${B("toggleLight", I.sun, "", "Ruban clair / sombre", { cls: "sl-lt" })}<div class="sl-win" aria-hidden="true"><i></i><i></i><i></i></div></div>` +
      `<div class="sl-tabs" role="tablist">${TABS.map(([k, l]) => `<button type="button" role="tab" class="sl-tab${k === "file" ? " sl-tfile" : ""}" data-tab="${k}">${l}</button>`).join("")}<span class="sl-tabind"></span></div>` +
      `<div class="sl-ribbon">${ribbonHTML()}</div>` +
      `<div class="sl-main"><div class="sl-thumbs" role="navigation" aria-label="Miniatures des diapositives"><div class="sl-tlist"></div>` +
      `<button type="button" class="sl-addslide" data-cmd="newSlide" title="Nouvelle diapositive (Ctrl+M)">+ Nouvelle diapositive</button></div>` +
      `<div class="sl-center"><div class="sl-stagewrap"><div class="sl-canvas"></div></div>` +
      `<div class="sl-notes"><textarea placeholder="Cliquez pour ajouter des notes" aria-label="Notes du présentateur" spellcheck="false"></textarea></div></div></div>` +
      `<div class="sl-status"><span class="sl-st-slide"></span><span class="sl-st-theme"></span><span class="sl-st-sp"></span>` +
      `${B("toggleNotes", I.notes, "Notes", "Afficher / masquer les notes")}${B("showCur", I.play, "", "Diaporama à partir de la diapositive actuelle (Maj+F5)")}<span class="sl-st-zoom">Ajusté</span></div>` +
      `<div class="sl-pop" hidden></div>`;
    host.appendChild(root);
    const q = s => root.querySelector(s);
    const qa = s => Array.prototype.slice.call(root.querySelectorAll(s));
    const tl = q(".sl-tlist"), wrap = q(".sl-stagewrap"), canvas = q(".sl-canvas"), notesTa = q(".sl-notes textarea"), pop = q(".sl-pop");
    let slideEl = null, guidesEl = null;

    /* ---------- rendu des diapositives ---------- */
    function fillText(t, e, edit) {
      t.textContent = "";
      const txt = e.text || "";
      if (!txt) { if (edit && PHTXT[e.type]) t.setAttribute("data-ph", PHTXT[e.type]); return; }
      t.removeAttribute("data-ph");
      const lines = txt.split("\n");
      if (e.type === "body") {
        lines.forEach(l => { const li = h("div", "sl-li"); li.textContent = l; t.appendChild(li); });
      } else {
        lines.forEach((l, i) => { if (i) t.appendChild(document.createElement("br")); t.appendChild(document.createTextNode(l)); });
      }
    }
    function place(d, e) {
      const st = d.style;
      st.left = e.x + "%"; st.top = e.y + "%"; st.width = e.w + "%"; st.height = e.h + "%";
      st.setProperty("--fs", e.size);
      st.fontWeight = e.bold ? "700" : "400";
      st.color = e.color || "";
      st.textAlign = e.align || "";
      const f = d.firstChild;
      if (f) f.style.background = e.fill || "";
    }
    function buildEl(e, mode) {
      const d = h("div", "sl-el sl-t-" + e.type + (e.type === "shape" ? " sl-sh-" + e.shape : "") +
        (mode === "edit" && !String(e.text).trim() && isTextType(e.type) ? " sl-empty" : ""));
      d.dataset.id = e.id;
      if (e.anim !== "none" && e.order) { d.dataset.order = e.order; d.dataset.anim = e.anim; }
      d.appendChild(h("div", "sl-fill"));
      const t = h("div", "sl-txt");
      fillText(t, e, mode === "edit");
      d.appendChild(t);
      place(d, e);
      if (mode === "edit" && e.anim !== "none") { const b = h("span", "sl-badge"); b.textContent = e.order; d.appendChild(b); }
      return d;
    }
    function buildSlide(s, mode) {
      const d = h("div", "sl-slide sl-th-" + S.theme + " sl-l-" + s.layout + (mode === "edit" ? " sl-edit" : ""));
      d.appendChild(h("div", "sl-deco"));
      s.els.forEach(e => { if (mode !== "edit" && hiddenInShow(e)) return; d.appendChild(buildEl(e, mode)); });
      return d;
    }
    const blackLayer = () => h("div", "sl-slide sl-black");

    function thumbNode(i) {
      const s = S.slides[i], n = S.slides.length;
      const fx = s.transition !== "none" || s.els.some(e => e.anim !== "none");
      const t = h("div", "sl-thumb" + (i === S.cur ? " on" : "") + (s.id === flashId ? " sl-new" : ""));
      t.dataset.i = i; t.tabIndex = 0; t.setAttribute("role", "button");
      t.setAttribute("aria-label", "Diapositive " + (i + 1) + (i === S.cur ? " (sélectionnée)" : ""));
      const mini = (cmd, ic, ti, dis) => `<button type="button" class="sl-b sl-mini" data-cmd="${cmd}" data-i="${i}" title="${ti}" aria-label="${ti}"${dis ? " disabled" : ""}>${ic}</button>`;
      t.innerHTML = `<span class="sl-tn">${i + 1}${fx ? '<i class="sl-tfx" title="Transition ou animation">★</i>' : ""}</span><div class="sl-tbox"></div>` +
        `<div class="sl-tact">${mini("slideUp", I.up, "Monter", i === 0)}${mini("slideDown", I.down, "Descendre", i === n - 1)}${mini("dupSlide", I.dup, "Dupliquer")}${mini("delSlide", I.trash, "Supprimer", n < 2)}</div>`;
      t.querySelector(".sl-tbox").appendChild(buildSlide(s, "thumb"));
      return t;
    }
    function renderThumbs() {
      tl.textContent = "";
      S.slides.forEach((s, i) => tl.appendChild(thumbNode(i)));
      flashId = null;
      const on = tl.children[S.cur];
      if (on) {
        const horiz = tl.scrollWidth > tl.clientWidth + 4 && getComputedStyle(tl).flexDirection === "row";
        const box = tl.parentNode;
        if (horiz) {
          if (on.offsetLeft < box.scrollLeft || on.offsetLeft + on.offsetWidth > box.scrollLeft + box.clientWidth) box.scrollLeft = on.offsetLeft - 12;
        } else if (on.offsetTop < box.scrollTop || on.offsetTop + on.offsetHeight > box.scrollTop + box.clientHeight) {
          box.scrollTop = on.offsetTop - 12;
        }
      }
    }
    function scheduleThumb(i) {
      pendingThumbs.add(i);
      if (raf) return;
      raf = requestAnimationFrame(() => {
        raf = 0;
        pendingThumbs.forEach(k => {
          const t = tl.children[k], s = S.slides[k];
          if (t && s) { const b = t.querySelector(".sl-tbox"); b.textContent = ""; b.appendChild(buildSlide(s, "thumb")); }
        });
        pendingThumbs.clear();
      });
    }
    function renderCanvas() {
      stopPreview();
      canvas.textContent = "";
      const s = cur();
      slideEl = buildSlide(s, "edit");
      if (enterFx) { slideEl.classList.add("sl-enter"); enterFx = false; }
      guidesEl = h("div", "sl-guides");
      slideEl.appendChild(guidesEl);
      canvas.appendChild(slideEl);
      updateSel();
    }
    const domOf = id => (slideEl ? slideEl.querySelector('.sl-el[data-id="' + id + '"]') : null);
    function updateSel() {
      if (!slideEl) return;
      slideEl.querySelectorAll(".sl-el.sl-sel").forEach(d => {
        d.classList.remove("sl-sel");
        d.querySelectorAll(".sl-hd").forEach(x => x.remove());
      });
      const d = S.sel && domOf(S.sel);
      if (d) {
        d.classList.add("sl-sel");
        ["nw", "n", "ne", "e", "se", "s", "sw", "w"].forEach(k => { const x = h("i", "sl-hd sl-hd-" + k); x.dataset.h = k; d.appendChild(x); });
      } else S.sel = null;
      updateUI();
    }
    function renderAll() { renderThumbs(); renderCanvas(); }

    function updateUI() {
      const e = selEl(), s = cur();
      TABS.forEach(([k]) => root.classList.toggle("sl-tab-" + k, k === tab));
      qa(".sl-tab").forEach(b => { const on = b.dataset.tab === tab; b.classList.toggle("on", on); b.setAttribute("aria-selected", on ? "true" : "false"); });
      qa(".sl-panel").forEach(p => p.classList.toggle("on", p.dataset.tab === tab));
      const ind = q(".sl-tabind"), at = q('.sl-tab[data-tab="' + tab + '"]');
      if (ind && at) { ind.style.left = at.offsetLeft + 10 + "px"; ind.style.width = Math.max(0, at.offsetWidth - 20) + "px"; }
      qa("[data-need]").forEach(b => { const n = b.dataset.need; b.disabled = n === "el" ? !e : n === "anim" ? !(e && e.anim !== "none") : false; });
      const on = (sel, v) => qa(sel).forEach(b => b.classList.toggle("on", !!v));
      on('[data-cmd="bold"]', e && e.bold);
      on('[data-cmd="alignL"]', e && e.align === "left");
      on('[data-cmd="alignC"]', e && e.align === "center");
      on('[data-cmd="alignR"]', e && e.align === "right");
      qa('[data-cmd="theme"]').forEach(b => b.classList.toggle("on", b.dataset.arg === S.theme));
      qa('[data-cmd="trans"]').forEach(b => b.classList.toggle("on", s && b.dataset.arg === s.transition));
      qa('[data-cmd="anim"]').forEach(b => b.classList.toggle("on", !!e && b.dataset.arg === e.anim));
      const sz = q(".sl-size");
      if (sz) {
        if (e) {
          const v = String(e.size);
          if (!Array.prototype.some.call(sz.options, o => o.value === v)) { const o = document.createElement("option"); o.value = v; o.textContent = v; sz.appendChild(o); }
          sz.value = v;
        } else sz.selectedIndex = -1;
      }
      const ac = q(".sl-acol i"); if (ac) ac.style.background = (e && e.color) || "";
      const ol = q(".sl-ordlab");
      if (ol) ol.textContent = e && e.anim !== "none" ? "Ordre : " + e.order + " / " + maxOrder(s) : (e ? "Ordre : aucune animation" : "Ordre : –");
      qa('[data-cmd="undo"]').forEach(b => { b.disabled = !past.length; });
      qa('[data-cmd="redo"]').forEach(b => { b.disabled = !future.length; });
      on('[data-cmd="toggleNotes"]', !root.classList.contains("sl-nonotes"));
      q(".sl-st-slide").textContent = "Diapositive " + (S.cur + 1) + " sur " + S.slides.length;
      q(".sl-st-theme").textContent = "Thème « " + THEMES[S.theme] + " »";
      if (document.activeElement !== notesTa) notesTa.value = (s && s.notes) || "";
    }

    /* ---------- commandes ---------- */
    function goSlide(i) {
      i = clamp(i, 0, S.slides.length - 1);
      stopEdit();
      if (i === S.cur && slideEl) return;
      S.cur = i; S.sel = null; enterFx = true;
      renderAll();
    }
    function newSlide(layout) {
      mutate(() => {
        const c = cur();
        const l = LAYOUTS[layout] ? layout : (c && c.layout !== "title" ? c.layout : "content");
        const s = mkSlide({ layout: l });
        S.slides.splice(S.cur + 1, 0, s);
        S.cur++; S.sel = null; flashId = s.id; enterFx = true;
      });
    }
    function cloneSlide(s) {
      const c = JSON.parse(JSON.stringify(s));
      c.id = uid("s"); c.els.forEach(e => { e.id = uid("e"); });
      return c;
    }
    function dupSlide(i) {
      mutate(() => {
        const c = cloneSlide(S.slides[i]);
        S.slides.splice(i + 1, 0, c);
        S.cur = i + 1; S.sel = null; flashId = c.id; enterFx = true;
      });
      C.toast("Diapositive dupliquée");
    }
    function delSlide(i) {
      if (S.slides.length < 2) { C.toast("Une présentation doit contenir au moins une diapositive."); return; }
      mutate(() => {
        S.slides.splice(i, 1);
        if (S.cur > i || S.cur >= S.slides.length) S.cur = Math.max(0, S.cur - 1);
        S.sel = null;
      });
    }
    function moveSlide(i, dir) {
      const j = i + dir;
      if (j < 0 || j >= S.slides.length) return;
      mutate(() => {
        const a = S.slides; const t = a[i]; a[i] = a[j]; a[j] = t;
        if (S.cur === i) S.cur = j; else if (S.cur === j) S.cur = i;
        flashId = t.id;
      });
    }
    function applyLayout(s, layout) {
      if (!LAYOUTS[layout]) return;
      const pool = { title: [], subtitle: [], body: [] }, others = [];
      s.els.forEach(e => { if (pool[e.type]) pool[e.type].push(e); else others.push(e); });
      const take = type => pool[type].shift() || (type === "subtitle" ? pool.body.shift() : type === "body" ? pool.subtitle.shift() : null);
      const fresh = PH[layout].map(([type, g]) => {
        const n = normEl(Object.assign({ type }, g));
        const o = take(type);
        if (o) { n.text = o.text; n.anim = o.anim; n.order = o.order; n.color = o.color; if (type !== "title") n.bold = o.bold; }
        return n;
      });
      const left = [].concat(pool.title, pool.subtitle, pool.body).filter(e => String(e.text).trim());
      s.layout = layout;
      s.els = fresh.concat(left, others);
      renumber(s);
    }
    function elSet(patch) {
      const e = selEl();
      if (!e) { C.toast("Sélectionnez d'abord un élément de la diapositive."); return; }
      mutate(() => Object.assign(e, typeof patch === "function" ? patch(e) : patch));
    }
    function growFont(dir) {
      elSet(e => {
        let i = SIZES.findIndex(v => v >= e.size);
        if (i < 0) i = SIZES.length - 1;
        if (dir > 0) i = SIZES[i] > e.size ? i : i + 1; else i = i - 1;
        return { size: SIZES[clamp(i, 0, SIZES.length - 1)] };
      });
    }
    function objAlign(k) {
      elSet(e => ({
        objL: { x: 0 }, objC: { x: r2(50 - e.w / 2) }, objR: { x: r2(100 - e.w) },
        objT: { y: 0 }, objM: { y: r2(50 - e.h / 2) }, objB: { y: r2(100 - e.h) }
      })[k]);
    }
    function zOrder(front) {
      const e = selEl(); if (!e) return;
      mutate(() => { const s = cur(); s.els = s.els.filter(x => x !== e); if (front) s.els.push(e); else s.els.unshift(e); });
    }
    function delEl() {
      const e = selEl(); if (!e) return;
      mutate(() => { const s = cur(); s.els = s.els.filter(x => x !== e); renumber(s); S.sel = null; });
    }
    function dupEl() {
      const e = selEl(); if (!e) return;
      mutate(() => {
        const s = cur(); const c = JSON.parse(JSON.stringify(e));
        c.id = uid("e"); c.x = r2(clamp(e.x + 3, -e.w + 2, 98)); c.y = r2(clamp(e.y + 3, -e.h + 2, 98));
        if (c.anim !== "none") c.order = maxOrder(s) + 1;
        s.els.splice(s.els.indexOf(e) + 1, 0, c);
        S.sel = c.id;
      });
    }
    function addEl(type, extra) {
      let id = null;
      mutate(() => {
        const s = cur(); const d = DEF[type];
        const e = normEl(Object.assign({ type }, d, extra || {}));
        e.x = r2(50 - e.w / 2); e.y = r2(50 - e.h / 2);
        let guard = 0;
        while (guard++ < 12 && s.els.some(o => Math.abs(o.x - e.x) < .5 && Math.abs(o.y - e.y) < .5)) { e.x += 3; e.y += 3; }
        s.els.push(e); S.sel = e.id; id = e.id;
      });
      if (type === "text" && id) startEdit(id);
    }
    function setTheme(t) {
      if (!THEMES[t] || t === S.theme) { if (THEMES[t]) pulseCanvas(); return; }
      mutate(() => { S.theme = t; });
      pulseCanvas();
    }
    function pulseCanvas() { if (slideEl && !reduced()) { slideEl.classList.remove("sl-flash"); void slideEl.offsetWidth; slideEl.classList.add("sl-flash"); } }
    function setTrans(t) {
      mutate(() => { cur().transition = t; });
      if (t !== "none") previewTrans(t);
    }
    function setAnim(t) {
      const e = selEl();
      if (!e) { C.toast("Sélectionnez d'abord un élément de la diapositive (clic sur un titre, un texte ou une forme)."); return; }
      const id = e.id;
      mutate(() => {
        const s = cur();
        if (t === "none") { e.anim = "none"; e.order = 0; }
        else { if (e.anim === "none" || !e.order) e.order = maxOrder(s) + 1; e.anim = t; }
        renumber(s);
      });
      if (t !== "none") { const d = domOf(id); if (d) animEl(d, t, slideEl); }
    }
    function shiftOrder(dir) {
      const e = selEl(); if (!e || e.anim === "none") return;
      const s = cur(); const o = s.els.find(x => x.anim !== "none" && x.order === e.order + dir);
      if (!o) return;
      mutate(() => { const t = o.order; o.order = e.order; e.order = t; });
    }

    function run(cmd, arg, btn) {
      const bi = btn && btn.dataset.i != null ? +btn.dataset.i : S.cur;
      switch (cmd) {
        case "newSlide": return newSlide();
        case "newSlideAs": return newSlide(arg);
        case "newSlideMenu": return layoutPop(btn, "newSlideAs");
        case "layoutMenu": return layoutPop(btn, "setLayout");
        case "setLayout": return mutate(() => { applyLayout(cur(), arg); S.sel = null; });
        case "dupSlide": return dupSlide(bi);
        case "delSlide": return delSlide(bi);
        case "slideUp": return moveSlide(bi, -1);
        case "slideDown": return moveSlide(bi, 1);
        case "grow": return growFont(1);
        case "shrink": return growFont(-1);
        case "bold": return elSet(e => ({ bold: !e.bold }));
        case "colorMenu": return swatchPop(btn, "setColor", "Couleur de police", "Automatique");
        case "setColor": return elSet({ color: arg || "" });
        case "fillMenu": return swatchPop(btn, "setFill", "Remplissage", "Par défaut (thème)");
        case "setFill": return elSet({ fill: arg || "" });
        case "alignL": return elSet({ align: "left" });
        case "alignC": return elSet({ align: "center" });
        case "alignR": return elSet({ align: "right" });
        case "objL": case "objC": case "objR": case "objT": case "objM": case "objB": return objAlign(cmd);
        case "front": return zOrder(true);
        case "back": return zOrder(false);
        case "delEl": return delEl();
        case "dupEl": return dupEl();
        case "textbox": return addEl("text", { text: "Votre texte" });
        case "shape": return addEl("shape", { shape: arg, text: "" });
        case "iconMenu": return iconPop(btn);
        case "addIcon": return addEl("icon", { text: arg });
        case "theme": return setTheme(arg);
        case "trans": return setTrans(arg);
        case "transAll": {
          const t = cur().transition;
          mutate(() => S.slides.forEach(s => { s.transition = t; }));
          return C.toast("Transition « " + TRANS.find(x => x[0] === t)[1] + " » appliquée à toutes les diapositives");
        }
        case "transPreview": {
          const t = cur().transition;
          if (t === "none") return C.toast("Aucune transition sur cette diapositive : choisissez-en une dans la galerie.");
          return previewTrans(t);
        }
        case "anim": return setAnim(arg);
        case "animPreview": return previewSlideAnims();
        case "animEarlier": return shiftOrder(-1);
        case "animLater": return shiftOrder(1);
        case "showStart": return startShow(0);
        case "showCur": return startShow(S.cur);
        case "undo": return undo();
        case "redo": return redo();
        case "toggleNotes": root.classList.toggle("sl-nonotes"); return updateUI();
        case "toggleLight": root.classList.toggle("sl-light"); return;
      }
    }

    /* ---------- menus contextuels (pop) ---------- */
    let popFor = null;
    function openPop(btn, html) {
      if (popFor === btn && !pop.hidden) { closePop(); return; }
      pop.innerHTML = html; pop.hidden = false; popFor = btn;
      const rr = root.getBoundingClientRect(), br = btn.getBoundingClientRect();
      pop.style.left = "0px"; pop.style.top = (br.bottom - rr.top + 4) + "px";
      const pw = pop.offsetWidth;
      pop.style.left = clamp(br.left - rr.left, 8, Math.max(8, rr.width - pw - 8)) + "px";
    }
    function closePop() { if (!pop.hidden) { pop.hidden = true; pop.innerHTML = ""; } popFor = null; }
    function layoutPop(btn, cmd) {
      const cards = Object.keys(LAYOUTS).map(k => {
        const bars = PH[k].map(([t, g]) => `<i class="m-${t}" style="left:${g.x}%;top:${g.y}%;width:${g.w}%;height:${g.h}%"></i>`).join("");
        return `<button type="button" class="sl-lay${cmd === "setLayout" && cur().layout === k ? " on" : ""}" data-cmd="${cmd}" data-arg="${k}"><span class="sl-lmini">${bars}</span><span>${LAYOUTS[k]}</span></button>`;
      }).join("");
      openPop(btn, `<div class="sl-pt">${cmd === "setLayout" ? "Disposition" : "Nouvelle diapositive"}</div><div class="sl-lays">${cards}</div>`);
    }
    function swatchPop(btn, cmd, title, autoLabel) {
      const e = selEl(); const curv = e ? (cmd === "setColor" ? e.color : e.fill) : "";
      const sw = COLORS.map(c => `<button type="button" class="sl-sw${curv === c ? " on" : ""}" data-cmd="${cmd}" data-arg="${c}" style="background:${c}" title="${c}" aria-label="${c}"></button>`).join("");
      openPop(btn, `<div class="sl-pt">${title}</div><button type="button" class="sl-swauto${!curv ? " on" : ""}" data-cmd="${cmd}" data-arg="">${autoLabel}</button><div class="sl-sws">${sw}</div>`);
    }
    function iconPop(btn) {
      openPop(btn, `<div class="sl-pt">Icônes</div><div class="sl-icons">${ICONS.map(c => `<button type="button" data-cmd="addIcon" data-arg="${c}" aria-label="Insérer ${c}">${c}</button>`).join("")}</div>`);
    }

    /* ---------- édition de texte ---------- */
    function readText(t, type) {
      let s = (t.innerText || t.textContent || "").replace(/ /g, " ").replace(/\r/g, "");
      if (type === "body") return s.split("\n").map(l => l.replace(/\s+$/, "")).filter(l => l.trim() !== "").join("\n");
      return s.replace(/\s+$/, "");
    }
    function startEdit(id) {
      const e = elById(id); if (!e) return;
      stopEdit();
      S.sel = id; updateSel();
      const d = domOf(id); if (!d) return;
      const t = d.querySelector(".sl-txt");
      editing = { id, before: e.text, snap: snap(), pushed: false, t, d, e };
      d.classList.add("sl-editing");
      d.classList.remove("sl-empty");
      if (e.type === "body" && !e.text) t.innerHTML = '<div class="sl-li"><br></div>';
      t.contentEditable = "true";
      t.spellcheck = false;
      t.addEventListener("input", onEditInput);
      t.addEventListener("blur", onEditBlur);
      t.addEventListener("paste", onPaste);
      try { t.focus({ preventScroll: true }); } catch (_) { t.focus(); }
      const r = document.createRange(); r.selectNodeContents(t);
      const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(r);
    }
    function onEditInput() {
      if (!editing) return;
      if (!editing.pushed) { pushHist(editing.snap); editing.pushed = true; }
      editing.e.text = readText(editing.t, editing.e.type);
      scheduleThumb(S.cur);
      C.changed();
    }
    function onEditBlur() { stopEdit(); }
    function onPaste(ev) {
      ev.preventDefault();
      const txt = (ev.clipboardData && ev.clipboardData.getData("text/plain")) || "";
      if (document.queryCommandSupported && document.queryCommandSupported("insertText")) document.execCommand("insertText", false, txt);
    }
    function stopEdit() {
      if (!editing) return;
      const ed = editing; editing = null;
      const { t, d, e } = ed;
      t.removeEventListener("input", onEditInput);
      t.removeEventListener("blur", onEditBlur);
      t.removeEventListener("paste", onPaste);
      if (d.isConnected) {
        e.text = readText(t, e.type);
        t.contentEditable = "false";
        t.removeAttribute("contenteditable");
        d.classList.remove("sl-editing");
        fillText(t, e, true);
        d.classList.toggle("sl-empty", !String(e.text).trim() && isTextType(e.type));
      }
      if (e.text !== ed.before && !ed.pushed) pushHist(ed.snap);
      try { const s = window.getSelection(); if (s && t.contains(s.anchorNode)) s.removeAllRanges(); } catch (_) { /* rien */ }
      scheduleThumb(S.cur);
      updateUI();
      if (e.text !== ed.before) C.changed();
    }

    /* ---------- souris : sélection, déplacement, redimensionnement, guides ---------- */
    function onWrapDown(ev) {
      if (ev.button !== 0 || show) return;
      closePop();
      focusArea = "canvas";
      if (editing && editing.t.contains(ev.target)) return;
      stopEdit();
      const elD = ev.target.closest(".sl-el"), hd = ev.target.closest(".sl-hd");
      try { root.focus({ preventScroll: true }); } catch (_) { /* rien */ }
      if (!elD || !slideEl.contains(elD)) { if (S.sel) { S.sel = null; updateSel(); } lastDown = {}; return; }
      const id = elD.dataset.id, now = Date.now();
      ev.preventDefault();
      if (!hd && lastDown.id === id && now - lastDown.t < 450) { lastDown = {}; startEdit(id); return; }
      lastDown = { id, t: now };
      if (S.sel !== id) { S.sel = id; updateSel(); }
      const e = elById(id); if (!e) return;
      drag = { mode: hd ? "resize" : "move", h: hd ? hd.dataset.h : "", id, sx: ev.clientX, sy: ev.clientY,
        o: { x: e.x, y: e.y, w: e.w, h: e.h }, r: slideEl.getBoundingClientRect(), moved: false, pid: ev.pointerId };
      try { wrap.setPointerCapture(ev.pointerId); } catch (_) { /* rien */ }
    }
    function candidates(id, axis) {
      const c = [{ v: 0, k: "edge" }, { v: 50, k: "center" }, { v: 100, k: "edge" }];
      cur().els.forEach(o => {
        if (o.id === id) return;
        const p = axis === "x" ? o.x : o.y, s = axis === "x" ? o.w : o.h;
        c.push({ v: p, k: "obj" }, { v: p + s / 2, k: "obj" }, { v: p + s, k: "obj" });
      });
      return c;
    }
    function snapAxis(pos, size, cands, thr) {
      const an = [pos, pos + size / 2, pos + size];
      let best = null;
      cands.forEach(c => an.forEach(a => { const d = c.v - a; if (Math.abs(d) <= thr && (!best || Math.abs(d) < Math.abs(best))) best = d; }));
      if (best == null) return { pos, lines: [] };
      pos += best;
      const an2 = [pos, pos + size / 2, pos + size];
      return { pos, lines: cands.filter(c => an2.some(a => Math.abs(a - c.v) < .02)) };
    }
    function snapEdge(v, cands, thr) {
      let best = null;
      cands.forEach(c => { const d = c.v - v; if (Math.abs(d) <= thr && (best == null || Math.abs(d) < Math.abs(best))) best = d; });
      if (best == null) return { v, lines: [] };
      v += best;
      return { v, lines: cands.filter(c => Math.abs(c.v - v) < .02) };
    }
    function drawGuides(vl, hl) {
      if (!guidesEl) return;
      const seen = {}; let html = "";
      vl.forEach(c => { const k = "v" + r1(c.v); if (seen[k]) return; seen[k] = 1; html += `<i class="sl-guide v ${c.k}" style="left:${c.v}%"></i>`; });
      hl.forEach(c => { const k = "h" + r1(c.v); if (seen[k]) return; seen[k] = 1; html += `<i class="sl-guide h ${c.k}" style="top:${c.v}%"></i>`; });
      guidesEl.innerHTML = html;
    }
    function onWrapMove(ev) {
      if (!drag || ev.pointerId !== drag.pid) return;
      let dx = (ev.clientX - drag.sx) / drag.r.width * 100, dy = (ev.clientY - drag.sy) / drag.r.height * 100;
      if (!drag.moved) {
        if (Math.hypot(ev.clientX - drag.sx, ev.clientY - drag.sy) < 3) return;
        drag.moved = true; pushHist(); root.classList.add("sl-dragging");
      }
      const e = elById(drag.id); if (!e) return;
      const o = drag.o, tx = 7 / drag.r.width * 100, ty = 7 / drag.r.height * 100, snapOn = !ev.altKey;
      let vl = [], hl = [];
      if (drag.mode === "move") {
        if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
        let x = o.x + dx, y = o.y + dy;
        if (snapOn) {
          const sx = snapAxis(x, o.w, candidates(e.id, "x"), tx); x = sx.pos; vl = sx.lines;
          const sy = snapAxis(y, o.h, candidates(e.id, "y"), ty); y = sy.pos; hl = sy.lines;
        }
        e.x = r2(clamp(x, -o.w + 2, 98)); e.y = r2(clamp(y, -o.h + 2, 98));
      } else {
        const k = drag.h; let x = o.x, y = o.y, w = o.w, hh = o.h;
        if (k.indexOf("e") >= 0) { let r = o.x + o.w + dx; if (snapOn) { const s = snapEdge(r, candidates(e.id, "x"), tx); r = s.v; vl = s.lines; } w = r - o.x; }
        if (k.indexOf("w") >= 0) { let l = o.x + dx; if (snapOn) { const s = snapEdge(l, candidates(e.id, "x"), tx); l = s.v; vl = s.lines; } x = l; w = o.x + o.w - l; }
        if (k.indexOf("s") >= 0) { let b = o.y + o.h + dy; if (snapOn) { const s = snapEdge(b, candidates(e.id, "y"), ty); b = s.v; hl = s.lines; } hh = b - o.y; }
        if (k.indexOf("n") >= 0) { let t = o.y + dy; if (snapOn) { const s = snapEdge(t, candidates(e.id, "y"), ty); t = s.v; hl = s.lines; } y = t; hh = o.y + o.h - t; }
        if (ev.shiftKey && k.length === 2) { hh = w * o.h / o.w; if (k.indexOf("n") >= 0) y = o.y + o.h - hh; vl = []; hl = []; }
        if (w < 2) { if (k.indexOf("w") >= 0) x = o.x + o.w - 2; w = 2; }
        if (hh < 2) { if (k.indexOf("n") >= 0) y = o.y + o.h - 2; hh = 2; }
        e.x = r2(x); e.y = r2(y); e.w = r2(w); e.h = r2(hh);
      }
      const d = domOf(e.id); if (d) place(d, e);
      drawGuides(vl, hl);
      scheduleThumb(S.cur);
    }
    function onWrapUp(ev) {
      if (!drag || ev.pointerId !== drag.pid) return;
      const moved = drag.moved; drag = null;
      try { wrap.releasePointerCapture(ev.pointerId); } catch (_) { /* rien */ }
      root.classList.remove("sl-dragging");
      drawGuides([], []);
      if (moved) { lastDown = {}; updateUI(); C.changed(); }
    }

    /* ---------- clavier ---------- */
    function nudge(dx, dy) {
      const e = selEl(); if (!e) return;
      const now = Date.now();
      if (!(lastNudge.id === e.id && now - lastNudge.t < 900)) pushHist();
      lastNudge = { id: e.id, t: now };
      e.x = r2(clamp(e.x + dx, -e.w + 2, 98)); e.y = r2(clamp(e.y + dy, -e.h + 2, 98));
      const d = domOf(e.id); if (d) place(d, e);
      scheduleThumb(S.cur); updateUI(); C.changed();
    }
    function onKey(ev) {
      if (show) { showKey(ev); return; }
      if (!active || !host.isConnected) return;
      const k = ev.key, mod = ev.ctrlKey || ev.metaKey;
      if (editing) {
        if (k === "Escape") { ev.preventDefault(); stopEdit(); try { root.focus({ preventScroll: true }); } catch (_) { /* rien */ } }
        return;
      }
      if (!pop.hidden && k === "Escape") { ev.preventDefault(); closePop(); return; }
      const tg = ev.target;
      if (tg && (tg.tagName === "INPUT" || tg.tagName === "TEXTAREA" || tg.tagName === "SELECT" || tg.isContentEditable)) return;
      if (mod && !ev.altKey) {
        const kl = k.toLowerCase();
        if (kl === "z" && !ev.shiftKey) { ev.preventDefault(); undo(); }
        else if (kl === "y" || (kl === "z" && ev.shiftKey)) { ev.preventDefault(); redo(); }
        else if (kl === "d") { ev.preventDefault(); if (selEl()) dupEl(); else dupSlide(S.cur); }
        else if (kl === "m") { ev.preventDefault(); newSlide(); }
        else if ((kl === "g" || kl === "b") && selEl()) { ev.preventDefault(); run("bold"); }
        return;
      }
      if (k === "F5") { ev.preventDefault(); startShow(ev.shiftKey ? S.cur : 0); return; }
      const e = selEl();
      if (e) {
        const st = ev.shiftKey ? 5 : 1;
        if (k === "Delete" || k === "Backspace") { ev.preventDefault(); delEl(); }
        else if (k === "ArrowLeft") { ev.preventDefault(); nudge(-st, 0); }
        else if (k === "ArrowRight") { ev.preventDefault(); nudge(st, 0); }
        else if (k === "ArrowUp") { ev.preventDefault(); nudge(0, -st); }
        else if (k === "ArrowDown") { ev.preventDefault(); nudge(0, st); }
        else if (k === "Enter" || k === "F2") { ev.preventDefault(); startEdit(e.id); }
        else if (k === "Escape") { ev.preventDefault(); S.sel = null; updateSel(); }
        else if (k === "Tab") { ev.preventDefault(); cycleSel(ev.shiftKey ? -1 : 1); }
      } else {
        if (k === "ArrowDown" || k === "PageDown" || (k === "ArrowRight" && focusArea === "thumbs")) { ev.preventDefault(); goSlide(S.cur + 1); }
        else if (k === "ArrowUp" || k === "PageUp" || (k === "ArrowLeft" && focusArea === "thumbs")) { ev.preventDefault(); goSlide(S.cur - 1); }
        else if ((k === "Delete" || k === "Backspace") && focusArea === "thumbs") { ev.preventDefault(); delSlide(S.cur); }
        else if (k === "Tab" && focusArea === "canvas" && cur().els.length) { ev.preventDefault(); cycleSel(1); }
        else if ((k === "Enter" || k === " ") && tg && tg.classList && tg.classList.contains("sl-thumb")) { ev.preventDefault(); goSlide(+tg.dataset.i); }
      }
    }
    function cycleSel(dir) {
      const els = cur().els; if (!els.length) return;
      let i = els.findIndex(x => x.id === S.sel);
      i = i < 0 ? (dir > 0 ? 0 : els.length - 1) : (i + dir + els.length) % els.length;
      S.sel = els[i].id; updateSel();
    }

    /* ---------- animations d'éléments (WAAPI) ---------- */
    function animEl(d, type, stageEl) {
      d.classList.remove("sl-pend");
      if (!canAnim) return null;
      d.getAnimations && d.getAnimations().forEach(a => a.cancel());
      if (reduced()) return d.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 160 });
      const sr = stageEl.getBoundingClientRect(), er = d.getBoundingClientRect();
      if (type === "appear") return d.animate([{ opacity: 0 }, { opacity: 0, offset: .5 }, { opacity: 1, offset: .51 }, { opacity: 1 }], { duration: 260 });
      if (type === "fly") {
        const dy = Math.max(40, sr.bottom - er.top);
        return d.animate([{ transform: `translateY(${dy}px)`, opacity: 0 }, { transform: `translateY(${dy * .45}px)`, opacity: 1, offset: .35 }, { transform: "translateY(0)", opacity: 1 }],
          { duration: 700, easing: "cubic-bezier(.2,.8,.2,1)" });
      }
      if (type === "zoom") return d.animate([{ transform: "scale(.15)", opacity: 0 }, { transform: "scale(1)", opacity: 1 }], { duration: 560, easing: "cubic-bezier(.34,1.56,.64,1)" });
      if (type === "bounce") {
        const up = Math.max(40, er.bottom - sr.top);
        return d.animate([
          { transform: `translateY(${-up}px)`, opacity: 0, offset: 0, easing: "cubic-bezier(.55,0,1,.45)" },
          { transform: "translateY(0)", opacity: 1, offset: .45, easing: "cubic-bezier(0,.55,.45,1)" },
          { transform: "translateY(-16%)", opacity: 1, offset: .62, easing: "cubic-bezier(.55,0,1,.45)" },
          { transform: "translateY(0)", opacity: 1, offset: .78, easing: "cubic-bezier(0,.55,.45,1)" },
          { transform: "translateY(-5%)", opacity: 1, offset: .88, easing: "cubic-bezier(.55,0,1,.45)" },
          { transform: "translateY(0)", opacity: 1, offset: 1 }], { duration: 1000 });
      }
      return null;
    }
    const stepsOf = s => Array.from(new Set(s.els.filter(e => e.anim !== "none" && e.order && !hiddenInShow(e)).map(e => e.order))).sort((a, b) => a - b);
    function previewSlideAnims() {
      if (!slideEl) return;
      const s = cur(), st = stepsOf(s);
      if (!st.length) { C.toast("Aucune animation sur cette diapositive : sélectionnez un élément puis choisissez un effet."); return; }
      const ids = st.map(o => s.els.filter(e => e.anim !== "none" && e.order === o));
      ids.forEach(g => g.forEach(e => { const d = domOf(e.id); if (d) d.classList.add("sl-pend"); }));
      ids.forEach((g, i) => setTimeout(() => {
        g.forEach(e => { const d = domOf(e.id); if (d) animEl(d, e.anim, slideEl); });
      }, 250 + i * (reduced() ? 200 : 750)));
    }

    /* ---------- transitions de diapositives ---------- */
    function runTransition(stage, oldL, newL, oldS, newS, type, onDone) {
      const R = reduced();
      let D = ({ fade: 700, push: 700, zoom: 650, morph: 1000 })[type] || 600;
      if (R) { type = "fade"; D = 180; }
      const anims = [];
      const ease = "cubic-bezier(.65,0,.35,1)";
      const o = { duration: D, easing: ease, fill: "both" };
      oldL.style.zIndex = "1"; newL.style.zIndex = "2";
      if (canAnim) {
        if (type === "fade") anims.push(newL.animate([{ opacity: 0 }, { opacity: 1 }], o));
        else if (type === "push") {
          anims.push(oldL.animate([{ transform: "translateX(0)" }, { transform: "translateX(-100%)" }], o));
          anims.push(newL.animate([{ transform: "translateX(100%)" }, { transform: "translateX(0)" }], o));
        } else if (type === "zoom") {
          anims.push(oldL.animate([{ opacity: 1, transform: "scale(1)" }, { opacity: 0, transform: "scale(1.35)" }], o));
          anims.push(newL.animate([{ opacity: 0, transform: "scale(.35)" }, { opacity: 1, transform: "scale(1)" }], { duration: D, easing: "cubic-bezier(.2,.8,.2,1)", fill: "both" }));
        } else if (type === "morph") morph(stage, oldL, newL, oldS, newS, D, anims);
      }
      let done = false;
      const fin = () => {
        if (done) return; done = true;
        anims.forEach(a => { try { a.cancel(); } catch (_) { /* rien */ } });
        if (oldL.parentNode) oldL.remove();
        newL.style.zIndex = "";
        newL.querySelectorAll(".sl-el").forEach(d => { d.style.visibility = ""; });
        if (onDone) onDone();
      };
      if (!anims.length) fin();
      else Promise.all(anims.map(a => a.finished)).then(fin, fin);
      return { finish: fin };
    }
    function morph(stage, oldL, newL, oldS, newS, D, anims) {
      const k = stage.getBoundingClientRect().width / 960;
      oldL.style.zIndex = "3"; newL.style.zIndex = "1";
      oldL.classList.add("sl-morph-old");
      const om = new Map((oldS.els || []).map(e => [e.id, e])), nm = new Map((newS.els || []).map(e => [e.id, e]));
      const olds = Array.prototype.map.call(oldL.querySelectorAll(".sl-el"), d => ({ d, e: om.get(d.dataset.id) })).filter(x => x.e && !x.d.classList.contains("sl-pend"));
      const news = Array.prototype.map.call(newL.querySelectorAll(".sl-el"), d => ({ d, e: nm.get(d.dataset.id) })).filter(x => x.e && !x.d.classList.contains("sl-pend"));
      const used = new Set(), pairs = new Map();
      const pick = fn => olds.find(x => !used.has(x) && fn(x.e));
      news.forEach(n => {
        const t = norm(n.e.text); if (!t) return;
        const m = pick(e => e.type === n.e.type && norm(e.text) === t) || pick(e => norm(e.text) === t);
        if (m) { used.add(m); pairs.set(n, m); }
      });
      news.forEach(n => {
        if (pairs.has(n)) return;
        const m = pick(e => e.type === n.e.type && (n.e.type !== "shape" || e.shape === n.e.shape));
        if (m) { used.add(m); pairs.set(n, m); }
      });
      const ease = "cubic-bezier(.65,0,.35,1)";
      news.forEach(n => {
        const m = pairs.get(n);
        if (!m) { anims.push(n.d.animate([{ opacity: 0 }, { opacity: 1 }], { duration: D * .55, delay: D * .45, fill: "both", easing: ease })); return; }
        const a = m.e, b = n.e;
        const ca = getComputedStyle(m.d).color, cb = getComputedStyle(n.d).color;
        m.d.style.visibility = "hidden";
        anims.push(n.d.animate([
          { left: a.x + "%", top: a.y + "%", width: a.w + "%", height: a.h + "%", fontSize: (a.size * k) + "px", color: ca },
          { left: b.x + "%", top: b.y + "%", width: b.w + "%", height: b.h + "%", fontSize: (b.size * k) + "px", color: cb }],
          { duration: D, easing: ease, fill: "both" }));
        const fa = getComputedStyle(m.d.firstChild).backgroundColor, fb = getComputedStyle(n.d.firstChild).backgroundColor;
        if (fa !== fb) anims.push(n.d.firstChild.animate([{ backgroundColor: fa }, { backgroundColor: fb }], { duration: D, easing: ease, fill: "both" }));
      });
      anims.push(oldL.animate([{ opacity: 1 }, { opacity: 0 }], { duration: D * .45, easing: ease, fill: "both" }));
    }
    function stopPreview() { if (preview) { const p = preview; preview = null; p.finish(); } }
    function previewTrans(type) {
      if (type === "none" || show || !slideEl) return;
      stopPreview(); stopEdit();
      const s = cur(), prev = S.slides[S.cur - 1];
      const tp = h("div", "sl-tprev");
      const oldL = prev ? buildSlide(prev, "show") : blackLayer(), newL = buildSlide(s, "show");
      tp.appendChild(oldL); tp.appendChild(newL);
      canvas.appendChild(tp);
      let p = null;
      p = runTransition(tp, oldL, newL, prev || { els: [] }, s, type, () => { tp.remove(); if (preview === p) preview = null; });
      if (tp.isConnected) preview = p;
    }

    /* ---------- diaporama ---------- */
    function buildShowLayer(s, revealAll) {
      const L = buildSlide(s, "show");
      if (!revealAll) L.querySelectorAll(".sl-el[data-order]").forEach(d => d.classList.add("sl-pend"));
      return L;
    }
    function startShow(from) {
      if (show) return;
      stopEdit(); closePop(); stopPreview();
      S.played = true; C.changed();
      const ov = h("div", "sl-show");
      ov.setAttribute("role", "dialog"); ov.setAttribute("aria-label", "Diaporama");
      ov.innerHTML = `<div class="sl-sstage"></div><div class="sl-sprog"><i></i></div>` +
        `<div class="sl-sbar"><button type="button" data-s="prev" aria-label="Précédent">‹</button><span class="sl-scount"></span>` +
        `<button type="button" data-s="next" aria-label="Suivant">›</button><button type="button" data-s="exit" class="sl-sexit">Échap · Quitter</button></div>` +
        `<div class="sl-shint">Clic, → ou Espace : suivant · ← : précédent · Échap : quitter</div>`;
      document.body.appendChild(ov);
      const stage = ov.querySelector(".sl-sstage");
      show = { ov, stage, i: from, step: 0, layer: null, tr: null, els: [], fs: false, end: true, hideT: 0, ov0: document.documentElement.style.overflow };
      document.documentElement.style.overflow = "hidden";
      const bl = blackLayer(); stage.appendChild(bl); show.layer = bl;
      ov.addEventListener("click", ev => {
        const b = ev.target.closest("[data-s]");
        if (b) { ev.stopPropagation(); const a = b.dataset.s; if (a === "prev") showPrev(); else if (a === "next") showNext(); else exitShow(); return; }
        showNext();
      });
      ov.addEventListener("contextmenu", ev => { ev.preventDefault(); showPrev(); });
      ov.addEventListener("pointermove", wake);
      wake();
      const rq = ov.requestFullscreen || ov.webkitRequestFullscreen;
      if (rq) {
        try {
          const p = rq.call(ov);
          if (p && p.then) p.then(() => { if (show && show.ov === ov) show.fs = true; }, () => { /* repli overlay */ });
          else if (show) show.fs = true;
        } catch (_) { /* repli overlay */ }
      }
      requestAnimationFrame(() => { if (show && show.ov === ov) goShowSlide(from, true, false); });
      try { ov.focus(); } catch (_) { /* rien */ }
    }
    function wake() {
      if (!show) return;
      show.ov.classList.remove("sl-idle");
      clearTimeout(show.hideT);
      const ov = show.ov;
      show.hideT = setTimeout(() => ov.classList.add("sl-idle"), 2200);
    }
    function finishShowAnims() {
      if (!show) return;
      if (show.tr) { const t = show.tr; show.tr = null; t.finish(); }
      show.els.forEach(a => { try { a.finish(); } catch (_) { /* rien */ } });
      show.els = [];
    }
    function goShowSlide(i, animate, revealAll) {
      if (!show) return;
      const to = S.slides[i]; if (!to) return;
      finishShowAnims();
      const old = show.layer, oldS = show.end ? { els: [] } : S.slides[show.i] || { els: [] };
      const nl = buildShowLayer(to, revealAll);
      show.stage.appendChild(nl);
      show.end = false; show.i = i; show.step = revealAll ? stepsOf(to).length : 0; show.layer = nl;
      if (old) {
        if (animate && to.transition !== "none") show.tr = runTransition(show.stage, old, nl, oldS, to, to.transition);
        else old.remove();
      }
      updateShowUI();
    }
    function showEnd() {
      finishShowAnims();
      const L = h("div", "sl-slide sl-endl");
      L.innerHTML = "<div><b>Fin du diaporama</b><span>Cliquez ou appuyez sur Échap pour quitter.</span></div>";
      show.stage.appendChild(L);
      const old = show.layer;
      show.tr = runTransition(show.stage, old, L, { els: [] }, { els: [] }, "fade");
      show.layer = L; show.end = true;
      updateShowUI();
    }
    function showNext() {
      if (!show) return;
      finishShowAnims();
      if (show.end && show.layer && show.layer.classList.contains("sl-endl")) { exitShow(); return; }
      const s = S.slides[show.i], st = stepsOf(s);
      if (show.step < st.length) {
        const o = st[show.step++];
        show.layer.querySelectorAll('.sl-el[data-order="' + o + '"]').forEach(d => {
          const a = animEl(d, d.dataset.anim, show.stage); if (a) show.els.push(a);
        });
        updateShowUI();
        return;
      }
      if (show.i < S.slides.length - 1) goShowSlide(show.i + 1, true, false);
      else showEnd();
    }
    function showPrev() {
      if (!show) return;
      finishShowAnims();
      if (show.end) { show.end = false; const L = show.layer; show.layer = null; goShowSlide(show.i, false, true); if (L && L.parentNode) L.remove(); return; }
      const st = stepsOf(S.slides[show.i]);
      if (show.step > 0) {
        const o = st[--show.step];
        show.layer.querySelectorAll('.sl-el[data-order="' + o + '"]').forEach(d => { d.getAnimations && d.getAnimations().forEach(a => a.cancel()); d.classList.add("sl-pend"); });
        updateShowUI();
        return;
      }
      if (show.i > 0) goShowSlide(show.i - 1, false, true);
    }
    function updateShowUI() {
      if (!show) return;
      const n = S.slides.length, st = stepsOf(S.slides[show.i] || { els: [] });
      const frac = show.end ? n : show.i + (st.length ? show.step / st.length : 1);
      show.ov.querySelector(".sl-sprog i").style.width = clamp(frac / n * 100, 0, 100) + "%";
      show.ov.querySelector(".sl-scount").textContent = show.end ? "Fin" : "Diapositive " + (show.i + 1) + " / " + n;
    }
    function showKey(ev) {
      const k = ev.key;
      if ([" ", "ArrowRight", "ArrowDown", "PageDown", "Enter", "n", "N"].indexOf(k) >= 0) { ev.preventDefault(); showNext(); }
      else if (["ArrowLeft", "ArrowUp", "PageUp", "Backspace", "p", "P"].indexOf(k) >= 0) { ev.preventDefault(); showPrev(); }
      else if (k === "Escape") { ev.preventDefault(); exitShow(); }
      else if (k === "Home") { ev.preventDefault(); goShowSlide(0, false, false); }
      else if (k === "End") { ev.preventDefault(); goShowSlide(S.slides.length - 1, false, true); }
      else if (k === "F5") ev.preventDefault();
      wake();
    }
    function exitShow() {
      if (!show) return;
      const sh = show;
      finishShowAnims();
      show = null;
      clearTimeout(sh.hideT);
      const fe = document.fullscreenElement || document.webkitFullscreenElement;
      if (fe === sh.ov) { try { const p = (document.exitFullscreen || document.webkitExitFullscreen).call(document); if (p && p.catch) p.catch(() => {}); } catch (_) { /* rien */ } }
      sh.ov.remove();
      document.documentElement.style.overflow = sh.ov0 || "";
      if (!sh.end && sh.i !== S.cur && S.slides[sh.i]) { S.cur = sh.i; S.sel = null; }
      if (host.isConnected) { renderAll(); try { root.focus({ preventScroll: true }); } catch (_) { /* rien */ } }
      active = true;
    }
    function onFs() {
      if (!show) return;
      const fe = document.fullscreenElement || document.webkitFullscreenElement;
      if (fe === show.ov) show.fs = true;
      else if (show.fs) exitShow();
    }

    /* ---------- écouteurs ---------- */
    function onDocDown(ev) {
      active = host.contains(ev.target);
      if (!pop.hidden && !pop.contains(ev.target) && !(popFor && popFor.contains(ev.target))) closePop();
    }
    root.addEventListener("click", ev => {
      const t = ev.target.closest(".sl-tab");
      if (t && root.contains(t)) {
        if (t.dataset.tab === "file") { C.toast("Simulation : votre présentation est conservée tant que la page reste ouverte."); return; }
        tab = t.dataset.tab; stopEdit(); closePop(); updateUI(); return;
      }
      const b = ev.target.closest("[data-cmd]");
      if (b && root.contains(b)) {
        if (b.disabled) return;
        if (!b.dataset.pop) closePop();
        run(b.dataset.cmd, b.dataset.arg, b);
        return;
      }
      const th = ev.target.closest(".sl-thumb");
      if (th) { focusArea = "thumbs"; goSlide(+th.dataset.i); }
    });
    root.addEventListener("change", ev => {
      if (ev.target.classList && ev.target.classList.contains("sl-size")) { const v = +ev.target.value; if (v) elSet({ size: v }); }
    });
    root.addEventListener("pointerdown", ev => {
      if (ev.target.closest(".sl-thumbs")) focusArea = "thumbs";
    });
    root.addEventListener("mousedown", ev => {
      // éviter que le clic dans le ruban fasse perdre la sélection de texte en cours d'édition
      if (editing && ev.target.closest(".sl-ribbon,.sl-titlebar") && !ev.target.closest("select")) ev.preventDefault();
    });
    root.addEventListener("pointerover", ev => {
      const b = ev.target.closest("[data-prev]");
      if (b === hoverB) return;
      hoverB = b; clearTimeout(hoverT);
      if (!b) return;
      hoverT = setTimeout(() => {
        if (hoverB !== b) return;
        const [kind, val] = b.dataset.prev.split(":");
        if (val === "none") return;
        if (kind === "trans") previewTrans(val);
        else if (kind === "anim") { const e = selEl(), d = e && domOf(e.id); if (d && !editing) animEl(d, val, slideEl); }
      }, 220);
    });
    root.addEventListener("pointerleave", () => { hoverB = null; clearTimeout(hoverT); });
    root.addEventListener("focusin", () => { active = true; });
    wrap.addEventListener("pointerdown", onWrapDown);
    wrap.addEventListener("pointermove", onWrapMove);
    wrap.addEventListener("pointerup", onWrapUp);
    wrap.addEventListener("pointercancel", onWrapUp);
    notesTa.addEventListener("input", () => {
      if (!notesPushed) { pushHist(); notesPushed = true; }
      cur().notes = notesTa.value; updateUI(); C.changed();
    });
    notesTa.addEventListener("blur", () => { notesPushed = false; });
    document.addEventListener("keydown", onKey);
    document.addEventListener("pointerdown", onDocDown, true);
    document.addEventListener("fullscreenchange", onFs);
    document.addEventListener("webkitfullscreenchange", onFs);

    /* ---------- API des check ---------- */
    const exportEl = e => ({
      type: e.type, text: e.text, x: r1(e.x), y: r1(e.y), w: r1(e.w), h: r1(e.h), size: e.size, bold: !!e.bold,
      color: e.color || "", fill: e.fill || "", shape: e.shape || "", anim: e.anim || "none", order: e.order | 0, align: e.align || "left"
    });
    const api = {
      get slides() { return S.slides.map(s => ({ layout: s.layout, transition: s.transition || "none", notes: s.notes || "", els: s.els.map(exportEl) })); },
      get theme() { return S.theme; },
      count() { return S.slides.length; },
      allText(i) { const s = S.slides[i]; return s ? s.els.map(e => e.text).filter(Boolean).join("\n") : ""; },
      find(txt) {
        const q2 = String(txt == null ? "" : txt).toLowerCase(), out = [];
        S.slides.forEach((s, i) => s.els.forEach(e => { if (String(e.text).toLowerCase().indexOf(q2) >= 0 && (q2 || e.text)) out.push({ slide: i, el: exportEl(e) }); }));
        return out;
      },
      centered(el) { return !!el && isFinite(+el.x) && isFinite(+el.w) && Math.abs((+el.x) + (+el.w) / 2 - 50) <= 3; },
      get played() { return S.played; },
      get current() { return S.cur; }
    };

    init();
    renderAll();
    requestAnimationFrame(() => { if (host.isConnected) updateUI(); });

    return {
      api,
      reset() {
        stopEdit(); exitShow(); stopPreview(); closePop();
        _id = 0; init(); enterFx = true; renderAll(); C.changed();
      },
      destroy() {
        stopEdit(); exitShow(); stopPreview(); closePop();
        clearTimeout(hoverT);
        if (raf) cancelAnimationFrame(raf);
        document.removeEventListener("keydown", onKey);
        document.removeEventListener("pointerdown", onDocDown, true);
        document.removeEventListener("fullscreenchange", onFs);
        document.removeEventListener("webkitfullscreenchange", onFs);
        host.textContent = "";
        host.classList.remove("sl-host");
      }
    };
  });
})();
