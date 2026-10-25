/* Office Academy — simulateur Excel (kind "sheet")
   Grille éditable + moteur de formules FR (alias EN), poignée de recopie, F4, Ctrl+D/R, Somme auto,
   mise en forme conditionnelle, tri, filtre, tableau, volets figés, graphiques SVG animés, TCD.
   API des check : voir assets/js/office/SPEC.md (section « kind "sheet" »). */
(function () {
  "use strict";
  /* ================= utilitaires ================= */
  class XErr { constructor(c) { this.c = c; } toString() { return this.c; } }
  const E = c => new XErr(c);
  const isErr = v => v instanceof XErr;
  const colIdx = s => s.toUpperCase().split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const colName = i => { let s = ""; i++; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };
  const refName = (c, r) => colName(c) + (r + 1);
  const parseRef = r => { const m = /^\$?([A-Z]{1,2})\$?(\d+)$/i.exec(String(r).trim()); return m ? { c: colIdx(m[1]), r: +m[2] - 1 } : null; };
  const parseRange = s => {
    const [a, b] = String(s).split(":"); const p = parseRef(a), q = b ? parseRef(b) : p;
    if (!p || !q) return null;
    return { c1: Math.min(p.c, q.c), r1: Math.min(p.r, q.r), c2: Math.max(p.c, q.c), r2: Math.max(p.r, q.r) };
  };
  const rangeName = g => g.c1 === g.c2 && g.r1 === g.r2 ? refName(g.c1, g.r1) : refName(g.c1, g.r1) + ":" + refName(g.c2, g.r2);
  const norm = s => String(s == null ? "" : s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
  const DAY0 = Date.UTC(1899, 11, 30);
  const toSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - DAY0) / 864e5;
  const fromSerial = n => new Date(DAY0 + Math.floor(n) * 864e5);
  const MOIS = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const JOURS = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];

  /* ================= moteur de formules ================= */
  const ALIAS = {
    SUM: "SOMME", AVERAGE: "MOYENNE", COUNT: "NB", COUNTA: "NBVAL", COUNTBLANK: "NB.VIDE", IF: "SI", IFS: "SI.CONDITIONS",
    AND: "ET", OR: "OU", NOT: "NON", IFERROR: "SIERREUR", ROUND: "ARRONDI", ROUNDUP: "ARRONDI.SUP", ROUNDDOWN: "ARRONDI.INF",
    INT: "ENT", POWER: "PUISSANCE", SQRT: "RACINE", SUMIF: "SOMME.SI", COUNTIF: "NB.SI", AVERAGEIF: "MOYENNE.SI",
    SUMIFS: "SOMME.SI.ENS", COUNTIFS: "NB.SI.ENS", SUMPRODUCT: "SOMMEPROD", VLOOKUP: "RECHERCHEV", HLOOKUP: "RECHERCHEH",
    XLOOKUP: "RECHERCHEX", MATCH: "EQUIV", CONCATENATE: "CONCATENER", LEFT: "GAUCHE", RIGHT: "DROITE", MID: "STXT", LEN: "NBCAR",
    UPPER: "MAJUSCULE", LOWER: "MINUSCULE", PROPER: "NOMPROPRE", TRIM: "SUPPRESPACE", TEXT: "TEXTE", TODAY: "AUJOURDHUI",
    NOW: "MAINTENANT", YEAR: "ANNEE", MONTH: "MOIS", DAY: "JOUR", WEEKDAY: "JOURSEM", ISBLANK: "ESTVIDE", ISNUMBER: "ESTNUM",
    ISTEXT: "ESTTEXTE", MAXIFS: "MAX.SI.ENS", MINIFS: "MIN.SI.ENS", MEDIAN: "MEDIANE", LARGE: "GRANDE.VALEUR",
    SMALL: "PETITE.VALEUR", RANK: "RANG", SUBSTITUTE: "SUBSTITUE", FIND: "TROUVE", SEARCH: "CHERCHE",
    "ANNÉE": "ANNEE", "MÉDIANE": "MEDIANE", "CONCATÉNER": "CONCATENER",
  };
  const canon = n => { n = n.toUpperCase(); return ALIAS[n] || n; };

  function tokenize(src) {
    const t = []; let i = 0;
    while (i < src.length) {
      const ch = src[i];
      if (/\s/.test(ch)) { i++; continue; }
      if (ch === '"') {
        let j = i + 1, s = "";
        for (;;) {
          if (j >= src.length) throw E("#VALEUR!");
          if (src[j] === '"') { if (src[j + 1] === '"') { s += '"'; j += 2; continue; } break; }
          s += src[j++];
        }
        t.push({ t: "str", v: s }); i = j + 1; continue;
      }
      const rest = src.slice(i); let m;
      if ((m = /^([A-ZÀ-Ü_][A-ZÀ-Ü0-9_.]*)\s*\(/i.exec(rest))) { t.push({ t: "fn", v: canon(m[1]) }); i += m[0].length; continue; }
      if ((m = /^(\$?[A-Z]{1,2}\$?\d+)(?::(\$?[A-Z]{1,2}\$?\d+))?(?![A-Z0-9_(])/i.exec(rest))) {
        if (m[2]) t.push({ t: "rng", a: m[1], b: m[2] }); else t.push({ t: "ref", v: m[1] });
        i += m[0].length; continue;
      }
      if ((m = /^\$?([A-Z]{1,2}):\$?([A-Z]{1,2})(?![A-Z0-9_(])/i.exec(rest))) { t.push({ t: "col", a: m[1], b: m[2] }); i += m[0].length; continue; }
      if ((m = /^(VRAI|FAUX|TRUE|FALSE)(?![A-Z0-9_.(])/i.exec(rest))) { t.push({ t: "bool", v: /^(vrai|true)$/i.test(m[1]) }); i += m[0].length; continue; }
      if ((m = /^(\d+(?:[.,]\d+)?|[.,]\d+)(?:[eE][+-]?\d+)?/.exec(rest))) { t.push({ t: "num", v: parseFloat(m[0].replace(",", ".")) }); i += m[0].length; continue; }
      if ((m = /^(<>|<=|>=)/.exec(rest))) { t.push({ t: "op", v: m[1] }); i += 2; continue; }
      if ("+-*/^&=<>%".includes(ch)) { t.push({ t: "op", v: ch }); i++; continue; }
      if (ch === "(" || ch === ")") { t.push({ t: ch }); i++; continue; }
      if (ch === ";" || ch === ",") { t.push({ t: ";" }); i++; continue; }
      if ((m = /^#(N\/A|NOM\?|VALEUR!|DIV\/0!|REF!|NOMBRE!|NULL!)/i.exec(rest))) { t.push({ t: "err", v: "#" + m[1].toUpperCase() }); i += m[0].length; continue; }
      throw E("#NOM?");
    }
    return t;
  }
  function parse(src) {
    const t = tokenize(src); let p = 0;
    const peek = () => t[p], next = () => t[p++];
    const isOp = (...ops) => peek() && peek().t === "op" && ops.includes(peek().v);
    function expr() { return cmp(); }
    function cmp() { let a = cat(); while (isOp("=", "<>", "<", ">", "<=", ">=")) { const op = next().v; a = { k: "bin", op, a, b: cat() }; } return a; }
    function cat() { let a = add(); while (isOp("&")) { next(); a = { k: "bin", op: "&", a, b: add() }; } return a; }
    function add() { let a = mul(); while (isOp("+", "-")) { const op = next().v; a = { k: "bin", op, a, b: mul() }; } return a; }
    function mul() { let a = pow(); while (isOp("*", "/")) { const op = next().v; a = { k: "bin", op, a, b: pow() }; } return a; }
    function pow() { let a = un(); while (isOp("^")) { next(); a = { k: "bin", op: "^", a, b: un() }; } return a; }
    function un() { if (isOp("-")) { next(); return { k: "neg", a: un() }; } if (isOp("+")) { next(); return un(); } return post(); }
    function post() { let a = prim(); while (isOp("%")) { next(); a = { k: "pct", a }; } return a; }
    function prim() {
      const k = next();
      if (!k) throw E("#VALEUR!");
      switch (k.t) {
        case "num": return { k: "num", v: k.v };
        case "str": return { k: "str", v: k.v };
        case "bool": return { k: "bool", v: k.v };
        case "err": return { k: "err", v: k.v };
        case "ref": { const r = parseRef(k.v); return { k: "ref", c: r.c, r: r.r }; }
        case "rng": { const g = parseRange(k.a + ":" + k.b); return { k: "rng", ...g }; }
        case "col": { const a = colIdx(k.a), b = colIdx(k.b); return { k: "rng", c1: Math.min(a, b), c2: Math.max(a, b), r1: 0, r2: 9999 }; }
        case "(": { const e = expr(); if (!peek() || peek().t !== ")") throw E("#VALEUR!"); next(); return e; }
        case "fn": {
          const args = [];
          if (peek() && peek().t === ")") { next(); return { k: "fn", n: k.v, args }; }
          for (;;) {
            if (peek() && (peek().t === ";" || peek().t === ")")) args.push({ k: "empty" }); else args.push(expr());
            const s = next();
            if (!s) throw E("#VALEUR!");
            if (s.t === ")") break;
            if (s.t !== ";") throw E("#VALEUR!");
          }
          return { k: "fn", n: k.v, args };
        }
      }
      throw E("#VALEUR!");
    }
    const ast = expr();
    if (p < t.length) throw E("#VALEUR!");
    return ast;
  }

  /* tableaux : { arr:true, rows:[[v]] } (une plage évaluée en est un, avec c1/r1) */
  const isArr = v => v && v.arr === true;
  const flat = a => a.rows.reduce((x, r) => x.concat(r), []);
  const scal = v => { if (isArr(v)) { if (!v.rows.length || !v.rows[0].length) throw E("#VALEUR!"); v = v.rows[0][0]; } if (isErr(v)) throw v; return v; };
  function toNum(v) {
    v = scal(v);
    if (v === null || v === "") return 0;
    if (typeof v === "number") return v;
    if (typeof v === "boolean") return v ? 1 : 0;
    const n = parseNumText(v);
    if (n === null) throw E("#VALEUR!");
    return n;
  }
  function parseNumText(s) {
    const t = String(s).replace(/[\s\u00a0\u202f]/g, "").replace("€", "");
    if (/^-?\d+([.,]\d+)?%$/.test(t)) return parseFloat(t.replace(",", ".")) / 100;
    if (/^-?(\d+([.,]\d+)?|[.,]\d+)$/.test(t)) return parseFloat(t.replace(",", "."));
    const d = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(t);
    if (d) return toSerial(+d[3] < 100 ? 2000 + +d[3] : +d[3], +d[2], +d[1]);
    return null;
  }
  const numFr = n => {
    if (!isFinite(n)) return "#NOMBRE!";
    if (Number.isInteger(n)) return String(n);
    return String(+n.toPrecision(12)).replace(".", ",");
  };
  function toText(v) {
    v = scal(v);
    if (v === null) return "";
    if (typeof v === "boolean") return v ? "VRAI" : "FAUX";
    if (typeof v === "number") return numFr(v);
    return String(v);
  }
  const toBool = v => { v = scal(v); if (typeof v === "boolean") return v; if (typeof v === "number") return v !== 0; if (v === null || v === "") return false;
    const s = String(v).toUpperCase(); if (s === "VRAI" || s === "TRUE") return true; if (s === "FAUX" || s === "FALSE") return false; throw E("#VALEUR!"); };
  function compare(a, b) {
    const rank = v => typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2;
    if (a === null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
    if (b === null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (typeof a === "string") { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    return a < b ? -1 : a > b ? 1 : 0;
  }
  function binop(op, a, b) {
    if (isErr(a)) throw a; if (isErr(b)) throw b;
    switch (op) {
      case "+": return toNum(a) + toNum(b);
      case "-": return toNum(a) - toNum(b);
      case "*": return toNum(a) * toNum(b);
      case "/": { const d = toNum(b); if (d === 0) throw E("#DIV/0!"); return toNum(a) / d; }
      case "^": return Math.pow(toNum(a), toNum(b));
      case "&": return toText(a) + toText(b);
      case "=": return compare(a, b) === 0;
      case "<>": return compare(a, b) !== 0;
      case "<": return compare(a, b) < 0;
      case ">": return compare(a, b) > 0;
      case "<=": return compare(a, b) <= 0;
      case ">=": return compare(a, b) >= 0;
    }
  }
  const safe = f => { try { return f(); } catch (e) { if (isErr(e)) return e; throw e; } };
  function elementwise(op, a, b) {
    if (!isArr(a) && !isArr(b)) return binop(op, scal(a), scal(b));
    const A = isArr(a) ? a.rows : [[a]], B = isArr(b) ? b.rows : [[b]];
    const R = Math.max(A.length, B.length), C = Math.max(A[0].length, B[0].length);
    const pick = (M, r, c) => M[M.length === 1 ? 0 : r] ? M[M.length === 1 ? 0 : r][M[0].length === 1 ? 0 : c] : E("#N/A");
    const rows = [];
    for (let r = 0; r < R; r++) { const row = []; for (let c = 0; c < C; c++) row.push(safe(() => binop(op, pick(A, r, c), pick(B, r, c)))); rows.push(row); }
    return { arr: true, rows };
  }
  const mapArr = (v, f) => isArr(v) ? { arr: true, rows: v.rows.map(r => r.map(x => safe(() => f(x)))) } : f(scal(v));

  /* critères SOMME.SI / NB.SI */
  function matcher(crit) {
    crit = scal(crit);
    if (typeof crit === "number" || typeof crit === "boolean") return v => v !== null && compare(typeof v === "string" && parseNumText(v) !== null ? parseNumText(v) : v, crit) === 0;
    let s = String(crit == null ? "" : crit), op = "=";
    const m = /^(<>|<=|>=|<|>|=)/.exec(s); if (m) { op = m[1]; s = s.slice(m[1].length); }
    const n = s.trim() === "" ? null : parseNumText(s);
    if (n !== null && op !== "=" && op !== "<>" || n !== null) {
      return v => {
        let x = v; if (typeof x === "string") { const p = parseNumText(x); if (p === null) return op === "<>"; x = p; }
        if (typeof x !== "number") return op === "<>";
        return op === "=" ? x === n : op === "<>" ? x !== n : op === "<" ? x < n : op === ">" ? x > n : op === "<=" ? x <= n : x >= n;
      };
    }
    if (s === "") return op === "<>" ? v => v !== null && v !== "" : v => v === null || v === "";
    const re = new RegExp("^" + s.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".") + "$", "i");
    if (op === "=" || op === "<>") return v => { const ok = v !== null && re.test(toText(v)); return op === "=" ? ok : !ok; };
    return v => { if (v === null || typeof v !== "string") return false; const c = compare(v, s); return op === "<" ? c < 0 : op === ">" ? c > 0 : op === "<=" ? c <= 0 : c >= 0; };
  }
  const nums = args => { const o = []; for (const a of args) { if (isArr(a)) { for (const v of flat(a)) { if (isErr(v)) throw v; if (typeof v === "number") o.push(v); } } else o.push(toNum(a)); } return o; };
  const round = (x, n, mode) => {
    const f = Math.pow(10, n), s = x < 0 ? -1 : 1, ax = Math.abs(x) * f, e = +ax.toPrecision(15);
    const r = mode === "up" ? Math.ceil(e) : mode === "down" ? Math.floor(e) : Math.round(e);
    return s * r / f;
  };
  const argArr = v => isArr(v) ? v : { arr: true, rows: [[scal(v)]] };
  const vec = v => { const a = argArr(v); return a.rows.length === 1 ? a.rows[0] : a.rows.map(r => r[0]); };
  function lookupIdx(list, val, mode) { // mode 0 exact (jokers), 1 ≤ trié croissant, -1 ≥ trié décroissant
    if (mode === 0) {
      const m = typeof val === "string" && /[*?]/.test(val) ? matcher(val) : null;
      for (let i = 0; i < list.length; i++) if (m ? m(list[i]) : list[i] !== null && compare(list[i], val) === 0) return i;
      return -1;
    }
    let best = -1;
    for (let i = 0; i < list.length; i++) {
      if (list[i] === null) continue;
      const c = compare(list[i], val);
      if (mode === 1) { if (c <= 0) best = i; else break; } else { if (c >= 0) best = i; else break; }
    }
    return best;
  }
  function ifsPairs(args, start) {
    const pairs = [];
    for (let i = start; i + 1 < args.length; i += 2) pairs.push({ list: flat(argArr(args[i])), m: matcher(args[i + 1]) });
    if (!pairs.length) throw E("#VALEUR!");
    return idx => pairs.every(p => p.m(p.list[idx] === undefined ? null : p.list[idx]));
  }
  function textFmt(v, f) {
    const n = typeof v === "number" ? v : parseNumText(v);
    const F = String(f).toLowerCase();
    if (/[jdma y]/.test(F) && /(jj|dd|mm|aa|yy)/.test(F) && n !== null) {
      const d = fromSerial(n), dd = d.getUTCDate(), mm = d.getUTCMonth() + 1, yy = d.getUTCFullYear();
      return F.replace(/jjjj|dddd/g, JOURS[d.getUTCDay()]).replace(/mmmm/g, MOIS[mm - 1]).replace(/aaaa|yyyy/g, yy).replace(/aa|yy/g, String(yy).slice(2))
        .replace(/jj|dd/g, String(dd).padStart(2, "0")).replace(/mm/g, String(mm).padStart(2, "0"));
    }
    if (n === null) return toText(v);
    const pct = F.includes("%"), dec = (/[.,](0+)/.exec(F) || ["", ""])[1].length, grp = /#\s|0\s0|# ##/.test(F);
    const x = pct ? n * 100 : n;
    let s = x.toLocaleString("fr-FR", { minimumFractionDigits: dec, maximumFractionDigits: dec, useGrouping: grp });
    s = s.replace(/[\u202f\u00a0]/g, " ");
    return s + (pct ? "%" : "") + (F.includes("€") ? " €" : "");
  }
  const FN = {
    SOMME: a => nums(a).reduce((s, x) => s + x, 0),
    MOYENNE: a => { const n = nums(a); if (!n.length) throw E("#DIV/0!"); return n.reduce((s, x) => s + x, 0) / n.length; },
    MIN: a => { const n = nums(a); return n.length ? Math.min(...n) : 0; },
    MAX: a => { const n = nums(a); return n.length ? Math.max(...n) : 0; },
    MEDIANE: a => { const n = nums(a).sort((x, y) => x - y); if (!n.length) throw E("#NOMBRE!"); const h = n.length >> 1; return n.length % 2 ? n[h] : (n[h - 1] + n[h]) / 2; },
    "GRANDE.VALEUR": a => { const n = nums([a[0]]).sort((x, y) => y - x), k = toNum(a[1]); if (k < 1 || k > n.length) throw E("#NOMBRE!"); return n[k - 1]; },
    "PETITE.VALEUR": a => { const n = nums([a[0]]).sort((x, y) => x - y), k = toNum(a[1]); if (k < 1 || k > n.length) throw E("#NOMBRE!"); return n[k - 1]; },
    RANG: a => { const v = toNum(a[0]), n = nums([a[1]]), asc = a[2] !== undefined && toNum(a[2]) !== 0; if (!n.includes(v)) throw E("#N/A"); return 1 + n.filter(x => asc ? x < v : x > v).length; },
    NB: a => a.reduce((s, x) => s + (isArr(x) ? flat(x).filter(v => typeof v === "number").length : (safe(() => toNum(x)) instanceof XErr ? 0 : 1)), 0),
    NBVAL: a => a.reduce((s, x) => s + (isArr(x) ? flat(x).filter(v => v !== null).length : 1), 0),
    "NB.VIDE": a => flat(argArr(a[0])).filter(v => v === null || v === "").length,
    ET: a => { const v = a.flatMap(x => isArr(x) ? flat(x).filter(y => typeof y === "boolean" || typeof y === "number") : [x]); if (!v.length) throw E("#VALEUR!"); return v.every(toBool); },
    OU: a => { const v = a.flatMap(x => isArr(x) ? flat(x).filter(y => typeof y === "boolean" || typeof y === "number") : [x]); if (!v.length) throw E("#VALEUR!"); return v.some(toBool); },
    NON: a => !toBool(a[0]),
    ARRONDI: a => round(toNum(a[0]), toNum(a[1] ?? 0)),
    "ARRONDI.SUP": a => round(toNum(a[0]), toNum(a[1] ?? 0), "up"),
    "ARRONDI.INF": a => round(toNum(a[0]), toNum(a[1] ?? 0), "down"),
    ABS: a => mapArr(a[0], x => Math.abs(toNum(x))),
    ENT: a => Math.floor(toNum(a[0])),
    MOD: a => { const x = toNum(a[0]), d = toNum(a[1]); if (d === 0) throw E("#DIV/0!"); return x - d * Math.floor(x / d); },
    PUISSANCE: a => Math.pow(toNum(a[0]), toNum(a[1])),
    RACINE: a => { const x = toNum(a[0]); if (x < 0) throw E("#NOMBRE!"); return Math.sqrt(x); },
    "SOMME.SI": a => { const r = flat(argArr(a[0])), m = matcher(a[1]), s = a[2] !== undefined ? flat(argArr(a[2])) : r; let t = 0;
      r.forEach((v, i) => { if (m(v) && typeof s[i] === "number") t += s[i]; }); return t; },
    "NB.SI": a => { const m = matcher(a[1]); return flat(argArr(a[0])).filter(v => m(v)).length; },
    "MOYENNE.SI": a => { const r = flat(argArr(a[0])), m = matcher(a[1]), s = a[2] !== undefined ? flat(argArr(a[2])) : r; const o = [];
      r.forEach((v, i) => { if (m(v) && typeof s[i] === "number") o.push(s[i]); }); if (!o.length) throw E("#DIV/0!"); return o.reduce((x, y) => x + y, 0) / o.length; },
    "SOMME.SI.ENS": a => { const s = flat(argArr(a[0])), ok = ifsPairs(a, 1); let t = 0; s.forEach((v, i) => { if (typeof v === "number" && ok(i)) t += v; }); return t; },
    "NB.SI.ENS": a => { const n = flat(argArr(a[0])).length, ok = ifsPairs(a, 0); let t = 0; for (let i = 0; i < n; i++) if (ok(i)) t++; return t; },
    "MAX.SI.ENS": a => { const s = flat(argArr(a[0])), ok = ifsPairs(a, 1), o = s.filter((v, i) => typeof v === "number" && ok(i)); return o.length ? Math.max(...o) : 0; },
    "MIN.SI.ENS": a => { const s = flat(argArr(a[0])), ok = ifsPairs(a, 1), o = s.filter((v, i) => typeof v === "number" && ok(i)); return o.length ? Math.min(...o) : 0; },
    SOMMEPROD: a => {
      const ms = a.map(argArr), R = ms[0].rows.length, C = ms[0].rows[0].length;
      if (ms.some(m => m.rows.length !== R || m.rows[0].length !== C)) throw E("#VALEUR!");
      let t = 0;
      for (let r = 0; r < R; r++) for (let c = 0; c < C; c++) {
        let p = 1;
        for (const m of ms) { const v = m.rows[r][c]; if (isErr(v)) throw v; p *= typeof v === "number" ? v : typeof v === "boolean" ? (ms.length === 1 ? +v : +v) : 0; }
        t += p;
      }
      return t;
    },
    RECHERCHEV: a => {
      const val = scal(a[0]), tb = argArr(a[1]), col = toNum(a[2]), approx = a[3] === undefined || a[3] === null ? true : toBool(a[3]);
      if (col < 1) throw E("#VALEUR!"); if (col > tb.rows[0].length) throw E("#REF!");
      const i = lookupIdx(tb.rows.map(r => r[0]), val, approx ? 1 : 0);
      if (i < 0) throw E("#N/A"); const v = tb.rows[i][col - 1]; if (isErr(v)) throw v; return v === null ? 0 : v;
    },
    RECHERCHEH: a => {
      const val = scal(a[0]), tb = argArr(a[1]), row = toNum(a[2]), approx = a[3] === undefined || a[3] === null ? true : toBool(a[3]);
      if (row < 1) throw E("#VALEUR!"); if (row > tb.rows.length) throw E("#REF!");
      const i = lookupIdx(tb.rows[0], val, approx ? 1 : 0);
      if (i < 0) throw E("#N/A"); const v = tb.rows[row - 1][i]; if (isErr(v)) throw v; return v === null ? 0 : v;
    },
    RECHERCHEX: a => {
      const val = scal(a[0]), look = vec(a[1]), ret = argArr(a[2]), mode = a[4] != null ? toNum(a[4]) : 0;
      let i;
      if (mode === 0 || mode === 2) i = lookupIdx(look, val, 0);
      else {
        i = -1; let bestV = null;
        look.forEach((v, k) => { if (v === null) return; const c = compare(v, val);
          if (c === 0 && i < 0) { i = k; bestV = v; return; }
          if (mode === -1 && c < 0 && (bestV === null || compare(v, bestV) > 0) && !(i >= 0 && compare(bestV, val) === 0)) { i = k; bestV = v; }
          if (mode === 1 && c > 0 && (bestV === null || compare(v, bestV) < 0) && !(i >= 0 && compare(bestV, val) === 0)) { i = k; bestV = v; } });
      }
      if (i < 0) { if (a[3] !== undefined && a[3] !== null) return scal(a[3]); throw E("#N/A"); }
      const v = ret.rows.length === 1 && look.length > 1 ? ret.rows[0][i] : (ret.rows[i] || [])[0];
      if (v === undefined) throw E("#VALEUR!"); if (isErr(v)) throw v; return v === null ? 0 : v;
    },
    INDEX: a => {
      const m = argArr(a[0]); let r = toNum(a[1] ?? 0), c = a[2] !== undefined && a[2] !== null ? toNum(a[2]) : 0;
      if (m.rows.length === 1 && c === 0) { c = r; r = 1; } if (c === 0) c = 1;
      const row = m.rows[r - 1]; if (!row || row[c - 1] === undefined) throw E("#REF!");
      const v = row[c - 1]; if (isErr(v)) throw v; return v === null ? 0 : v;
    },
    EQUIV: a => { const val = scal(a[0]), list = vec(a[1]), t = a[2] !== undefined && a[2] !== null ? toNum(a[2]) : 1;
      const i = lookupIdx(list, val, t === 0 ? 0 : t > 0 ? 1 : -1); if (i < 0) throw E("#N/A"); return i + 1; },
    CONCAT: a => a.map(x => isArr(x) ? flat(x).map(toText).join("") : toText(x)).join(""),
    CONCATENER: a => a.map(toText).join(""),
    GAUCHE: a => toText(a[0]).slice(0, a[1] === undefined ? 1 : toNum(a[1])),
    DROITE: a => { const s = toText(a[0]), n = a[1] === undefined ? 1 : toNum(a[1]); return n <= 0 ? "" : s.slice(-n); },
    STXT: a => { const s = toText(a[0]), st = toNum(a[1]); if (st < 1) throw E("#VALEUR!"); return s.substr(st - 1, toNum(a[2])); },
    NBCAR: a => toText(a[0]).length,
    MAJUSCULE: a => toText(a[0]).toUpperCase(),
    MINUSCULE: a => toText(a[0]).toLowerCase(),
    NOMPROPRE: a => toText(a[0]).toLowerCase().replace(/(^|[^a-zà-ÿ])([a-zà-ÿ])/g, (m, p, c) => p + c.toUpperCase()),
    SUPPRESPACE: a => toText(a[0]).replace(/\s+/g, " ").trim(),
    SUBSTITUE: a => { const s = toText(a[0]), o = toText(a[1]), n = toText(a[2]); if (!o) return s;
      if (a[3] === undefined) return s.split(o).join(n); const k = toNum(a[3]); let i = -1; for (let j = 0; j < k; j++) { i = s.indexOf(o, i + 1); if (i < 0) return s; } return s.slice(0, i) + n + s.slice(i + o.length); },
    TROUVE: a => { const i = toText(a[1]).indexOf(toText(a[0]), (a[2] ? toNum(a[2]) : 1) - 1); if (i < 0) throw E("#VALEUR!"); return i + 1; },
    CHERCHE: a => { const i = toText(a[1]).toLowerCase().indexOf(toText(a[0]).toLowerCase(), (a[2] ? toNum(a[2]) : 1) - 1); if (i < 0) throw E("#VALEUR!"); return i + 1; },
    TEXTE: a => textFmt(scal(a[0]), toText(a[1])),
    AUJOURDHUI: () => { const d = new Date(); return toSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()); },
    MAINTENANT: () => { const d = new Date(); return toSerial(d.getFullYear(), d.getMonth() + 1, d.getDate()) + (d.getHours() * 60 + d.getMinutes()) / 1440; },
    DATE: a => toSerial(toNum(a[0]), toNum(a[1]), toNum(a[2])),
    ANNEE: a => fromSerial(toNum(a[0])).getUTCFullYear(),
    MOIS: a => fromSerial(toNum(a[0])).getUTCMonth() + 1,
    JOUR: a => fromSerial(toNum(a[0])).getUTCDate(),
    JOURSEM: a => { const d = fromSerial(toNum(a[0])).getUTCDay(), t = a[1] ? toNum(a[1]) : 1; return t === 2 ? (d + 6) % 7 + 1 : t === 3 ? (d + 6) % 7 : d + 1; },
    ESTVIDE: a => { const v = isArr(a[0]) ? a[0].rows[0][0] : a[0]; return v === null; },
    ESTNUM: a => { const v = isArr(a[0]) ? a[0].rows[0][0] : a[0]; return typeof v === "number"; },
    ESTTEXTE: a => { const v = isArr(a[0]) ? a[0].rows[0][0] : a[0]; return typeof v === "string"; },
  };
  const LAZY = {
    SI(args, ev) { const c = toBool(ev(args[0])); const n = c ? args[1] : args[2]; if (!n) return c; if (n.k === "empty") return c ? 0 : false; return ev(n); },
    "SI.CONDITIONS"(args, ev) { for (let i = 0; i + 1 < args.length; i += 2) if (toBool(ev(args[i]))) return ev(args[i + 1]); throw E("#N/A"); },
    SIERREUR(args, ev) { try { const v = ev(args[0]); if (isErr(v)) return ev(args[1]); if (isArr(v)) return scal(v); return v; } catch (e) { if (isErr(e)) return args[1] && args[1].k !== "empty" ? ev(args[1]) : 0; throw e; } },
  };
  const FN_NAMES = [...Object.keys(FN), ...Object.keys(LAZY)].sort();
  const FN_HELP = {
    SOMME: "SOMME(nombre1; [nombre2]; …)", MOYENNE: "MOYENNE(nombre1; [nombre2]; …)", MIN: "MIN(nombre1; …)", MAX: "MAX(nombre1; …)",
    NB: "NB(valeur1; …) — compte les nombres", NBVAL: "NBVAL(valeur1; …) — compte les cellules non vides", "NB.VIDE": "NB.VIDE(plage)",
    SI: "SI(test_logique; [valeur_si_vrai]; [valeur_si_faux])", "SI.CONDITIONS": "SI.CONDITIONS(test1; valeur1; [test2; valeur2]; …)",
    ET: "ET(logique1; [logique2]; …)", OU: "OU(logique1; [logique2]; …)", NON: "NON(logique)", SIERREUR: "SIERREUR(valeur; valeur_si_erreur)",
    ARRONDI: "ARRONDI(nombre; no_chiffres)", "ARRONDI.SUP": "ARRONDI.SUP(nombre; no_chiffres)", "ARRONDI.INF": "ARRONDI.INF(nombre; no_chiffres)",
    "SOMME.SI": "SOMME.SI(plage; critère; [somme_plage])", "NB.SI": "NB.SI(plage; critère)", "MOYENNE.SI": "MOYENNE.SI(plage; critère; [moyenne_plage])",
    "SOMME.SI.ENS": "SOMME.SI.ENS(somme_plage; plage_critères1; critère1; …)", "NB.SI.ENS": "NB.SI.ENS(plage_critères1; critère1; …)",
    SOMMEPROD: "SOMMEPROD(matrice1; [matrice2]; …)", RECHERCHEV: "RECHERCHEV(valeur_cherchée; table_matrice; no_index_col; [valeur_proche])",
    RECHERCHEH: "RECHERCHEH(valeur_cherchée; table_matrice; no_index_lig; [valeur_proche])",
    RECHERCHEX: "RECHERCHEX(valeur_cherchée; tableau_recherche; tableau_renvoyé; [si_non_trouvé]; [mode_correspondance])",
    INDEX: "INDEX(matrice; no_ligne; [no_colonne])", EQUIV: "EQUIV(valeur_cherchée; tableau_recherche; [type]) — 0 = exact",
    CONCAT: "CONCAT(texte1; [texte2]; …)", GAUCHE: "GAUCHE(texte; [no_car])", DROITE: "DROITE(texte; [no_car])", STXT: "STXT(texte; no_départ; no_car)",
    NBCAR: "NBCAR(texte)", MAJUSCULE: "MAJUSCULE(texte)", NOMPROPRE: "NOMPROPRE(texte)", SUPPRESPACE: "SUPPRESPACE(texte)",
    TEXTE: "TEXTE(valeur; format_texte) — ex. \"jj/mm/aaaa\"", DATE: "DATE(année; mois; jour)", AUJOURDHUI: "AUJOURDHUI()",
    ANNEE: "ANNEE(numéro_de_série)", MOIS: "MOIS(numéro_de_série)", JOUR: "JOUR(numéro_de_série)", RANG: "RANG(nombre; référence; [ordre])",
    "MAX.SI.ENS": "MAX.SI.ENS(plage_max; plage_critères1; critère1; …)", "MIN.SI.ENS": "MIN.SI.ENS(plage_min; plage_critères1; critère1; …)",
    SUBSTITUE: "SUBSTITUE(texte; ancien_texte; nouveau_texte; [no_position])", MOD: "MOD(nombre; diviseur)",
  };

  /* décalage des références relatives (recopie, copier/coller, tri) */
  function shiftFormula(f, dc, dr) {
    return f.replace(/("(?:[^"]|"")*")|(^|[^A-Z0-9_.$À-Ü])(\$?)([A-Z]{1,2})(\$?)(\d+)(?![A-Z0-9_(])/gi, (m, str, pre, d1, col, d2, row) => {
      if (str) return str;
      const c = d1 ? colIdx(col) : colIdx(col) + dc, r = d2 ? +row - 1 : +row - 1 + dr;
      if (c < 0 || r < 0) return pre + "#REF!";
      return pre + d1 + colName(c) + d2 + (r + 1);
    });
  }
  function cycleAbs(ref) {
    const m = /^(\$?)([A-Z]{1,2})(\$?)(\d+)$/i.exec(ref); if (!m) return ref;
    const st = (m[1] ? 2 : 0) + (m[3] ? 1 : 0); // 0 A1, 3 $A$1, 1 A$1, 2 $A1
    const nx = { 0: 3, 3: 1, 1: 2, 2: 0 }[st];
    return (nx & 2 ? "$" : "") + m[2].toUpperCase() + (nx & 1 ? "$" : "") + m[4];
  }

  /* format d'affichage */
  function display(v, fmt) {
    if (isErr(v)) return v.c;
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return v ? "VRAI" : "FAUX";
    if (typeof v === "number") {
      if (!isFinite(v)) return "#NOMBRE!";
      const L = (n, d) => n.toLocaleString("fr-FR", { minimumFractionDigits: d, maximumFractionDigits: d }).replace(/[\u202f\u00a0]/g, " ");
      switch (fmt) {
        case "num": return L(v, 2);
        case "int": return L(v, 0);
        case "eur": return L(v, 2) + " €";
        case "pct": return L(v * 100, Math.abs(v * 100 - Math.round(v * 100)) > 1e-9 ? 1 : 0) + " %";
        case "date": { const d = fromSerial(v); return String(d.getUTCDate()).padStart(2, "0") + "/" + String(d.getUTCMonth() + 1).padStart(2, "0") + "/" + d.getUTCFullYear(); }
      }
      if (Number.isInteger(v)) return String(v);
      const s = String(+v.toPrecision(10));
      return s.includes("e") ? s : s.replace(".", ",");
    }
    return String(v);
  }
  /* saisie → valeur brute (+ format déduit) */
  function parseInput(s) {
    if (typeof s === "number" || typeof s === "boolean") return { raw: s };
    s = String(s == null ? "" : s);
    const t = s.trim();
    if (t === "") return { raw: null };
    if (t[0] === "=" && t.length > 1) return { raw: t };
    if (/^(vrai|faux|true|false)$/i.test(t)) return { raw: /^(vrai|true)$/i.test(t) };
    const c = t.replace(/[\s\u00a0\u202f]/g, "");
    if (/^-?\d+([.,]\d+)?%$/.test(c)) return { raw: parseFloat(c.replace(",", ".")) / 100, fmt: "pct" };
    if (/^-?\d+([.,]\d+)?€$|^€-?\d+([.,]\d+)?$/.test(c)) return { raw: parseFloat(c.replace("€", "").replace(",", ".")), fmt: "eur" };
    const d = /^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/.exec(c);
    if (d && +d[2] >= 1 && +d[2] <= 12) return { raw: toSerial(+d[3] < 100 ? 2000 + +d[3] : +d[3], +d[2], +d[1]), fmt: "date" };
    if (/^-?(\d+([.,]\d+)?|[.,]\d+)$/.test(c)) return { raw: parseFloat(c.replace(",", ".")) };
    return { raw: s };
  }

  const PALETTE = ["", "#FFF2CC", "#E2EFDA", "#DDEBF7", "#FCE4D6", "#F8CBAD", "#C6E0B4", "#BDD7EE", "#FFE699", "#D9D9D9", "#FFC7CE", "#C6EFCE"];
  const CF_STYLES = [
    { id: "red", label: "Remplissage rouge clair, texte rouge foncé", fill: "#FFC7CE", color: "#9C0006" },
    { id: "green", label: "Remplissage vert, texte vert foncé", fill: "#C6EFCE", color: "#006100" },
    { id: "yellow", label: "Remplissage jaune, texte jaune foncé", fill: "#FFEB9C", color: "#9C5700" },
  ];
  const SERIES = ["#21A366", "#2B7CD3", "#D35230", "#F2B705", "#7E57C2", "#00A3A3"];
  const ICON = {
    undo: '<path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>', redo: '<path d="m15 14 5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>',
    fill: '<path d="m19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2a2 2 0 0 0 2.8 0z"/><path d="M5 2l5 5"/><path d="M2 13h15"/><path d="M22 20a2 2 0 1 1-4 0c0-1.6 1.7-2.4 2-4 .3 1.6 2 2.4 2 4z"/>',
    al: '<path d="M3 6h18M3 12h12M3 18h16"/>', ac: '<path d="M3 6h18M6 12h12M4 18h16"/>', ar: '<path d="M3 6h18M9 12h12M5 18h16"/>',
    sum: '<path d="M18 7V4H6l6 8-6 8h12v-3"/>', down: '<path d="M12 5v14"/><path d="m19 12-7 7-7-7"/>', right: '<path d="M5 12h14"/><path d="m12 5 7 7-7 7"/>',
    az: '<path d="M3 16l4 4 4-4M7 20V4"/><path d="M14 4h6l-6 7h6M14 15h6M17 15v5"/>', za: '<path d="M3 8l4-4 4 4M7 4v16"/><path d="M14 4h6l-6 7h6M14 15h6M17 15v5"/>',
    filter: '<path d="M3 4h18l-7 9v6l-4 2v-8z"/>', cf: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 21V9"/><path d="M13 13h5M13 17h3"/>',
    table: '<rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 10h18M3 15h18M9 4v16"/>', freeze: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 8h18" stroke-width="3"/>',
    bar: '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/>', line: '<path d="M3 17l6-6 4 4 8-8"/><path d="M3 21h18"/>', pie: '<path d="M21 12A9 9 0 1 1 12 3v9z"/><path d="M15 3.5A9 9 0 0 1 20.5 9H15z"/>',
    pivot: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18M9 3v18"/><path d="m14 14 3 3-3 3" transform="translate(-1 -3)"/>',
    fx: '<path d="M8 21c2 0 2.5-1.5 3-4l2-10c.5-2.5 1-4 3-4"/><path d="M6 11h8"/><path d="m15 13 5 6M20 13l-5 6"/>',
    show: '<path d="M4 7h16M4 12h10M4 17h7"/><path d="m16 15 4 4M20 15l-4 4"/>',
  };
  const svg = (k, s = 18) => `<svg viewBox="0 0 24 24" width="${s}" height="${s}" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${ICON[k]}</svg>`;

  /* ================= simulateur ================= */
  function mount(host, setup, ctx) {
    const esc = ctx.esc;
    const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;
    const NC = Math.min(12, Math.max(3, setup.cols || 8)), NR = Math.min(30, Math.max(3, setup.rows || 15));
    let S; // état
    const fresh = () => {
      const cells = {};
      Object.entries(setup.data || {}).forEach(([k, v]) => { const p = parseInput(v); if (p.raw !== null) cells[k.toUpperCase()] = { raw: p.raw, fmt: p.fmt || "" }; });
      Object.entries(setup.fmt || {}).forEach(([k, f]) => { (cells[k.toUpperCase()] = cells[k.toUpperCase()] || { raw: null }).fmt = f; });
      (setup.bold || []).forEach(k => { (cells[k.toUpperCase()] = cells[k.toUpperCase()] || { raw: null }).b = true; });
      return { cells, cf: [], charts: [], pivots: [], filter: null, table: null, frozen: false, showF: false };
    };
    S = fresh();
    let sel = { ac: 0, ar: 0, c1: 0, r1: 0, c2: 0, r2: 0 };
    let hist = [], fut = [], memo = {}, astCache = {}, edit = null, clip = null;

    /* ---------- calcul ---------- */
    const key = (c, r) => refName(c, r);
    const cell = (c, r) => S.cells[key(c, r)];
    const rawOf = (c, r) => { const x = cell(c, r); return x ? x.raw : null; };
    function getVal(c, r) {
      const k = key(c, r);
      if (k in memo) { if (memo[k] === CYC) return E("#REF!"); return memo[k]; }
      const raw = rawOf(c, r);
      if (typeof raw !== "string" || raw[0] !== "=") return (memo[k] = raw === undefined ? null : raw);
      memo[k] = CYC;
      let v;
      try {
        const ast = astCache[raw] || (astCache[raw] = parse(raw.slice(1)));
        v = ev(ast);
        if (isArr(v)) v = v.rows[0][0];
        if (v === null) v = 0;
        if (typeof v === "number" && !isFinite(v)) v = E("#NOMBRE!");
      } catch (e) { v = isErr(e) ? e : E("#VALEUR!"); }
      return (memo[k] = v);
    }
    const CYC = {};
    function ev(n) {
      switch (n.k) {
        case "num": case "str": case "bool": return n.v;
        case "err": throw E(n.v);
        case "empty": return null;
        case "ref": { if (n.c >= NC || n.r >= NR) return null; const v = getVal(n.c, n.r); if (isErr(v)) throw v; return v; }
        case "rng": {
          const r2 = Math.min(n.r2, NR - 1), c2 = Math.min(n.c2, NC - 1), rows = [];
          for (let r = n.r1; r <= r2; r++) { const row = []; for (let c = n.c1; c <= c2; c++) row.push(getVal(c, r)); rows.push(row); }
          if (!rows.length) rows.push([null]);
          return { arr: true, rows, c1: n.c1, r1: n.r1 };
        }
        case "neg": return mapArr(ev(n.a), x => -toNum(x));
        case "pct": return mapArr(ev(n.a), x => toNum(x) / 100);
        case "bin": return elementwise(n.op, ev(n.a), ev(n.b));
        case "fn": {
          if (LAZY[n.n]) return LAZY[n.n](n.args, ev);
          const f = FN[n.n]; if (!f) throw E("#NOM?");
          return f(n.args.map(a => a.k === "empty" ? null : ev(a)).map(v => v === null ? null : v).filter((v, i) => !(v === null && n.args[i].k === "empty" && i === n.args.length - 1 && false)));
        }
      }
      throw E("#VALEUR!");
    }
    const recalc = () => { memo = {}; };
    const valAt = ref => { const p = parseRef(ref); return p ? getVal(p.c, p.r) : E("#REF!"); };

    /* ---------- historique ---------- */
    const snap = () => JSON.stringify(S);
    function commit(fn, label) {
      hist.push(snap()); if (hist.length > 80) hist.shift(); fut = [];
      fn(); recalc(); render(); ctx.changed();
      if (label) ctx.toast(label);
    }
    function undo() { if (!hist.length) return; fut.push(snap()); S = JSON.parse(hist.pop()); recalc(); render(); ctx.changed(); }
    function redo() { if (!fut.length) return; hist.push(snap()); S = JSON.parse(fut.pop()); recalc(); render(); ctx.changed(); }

    function setRaw(c, r, text) {
      const k = key(c, r), p = parseInput(text);
      const x = S.cells[k] || {};
      x.raw = p.raw;
      if (p.fmt && !x.fmt) x.fmt = p.fmt;
      if (x.raw === null && !x.b && !x.fill && !x.fmt && !x.align) delete S.cells[k]; else S.cells[k] = x;
    }
    const each = (g, f) => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) f(c, r); };
    const selG = () => ({ c1: sel.c1, r1: sel.r1, c2: sel.c2, r2: sel.r2 });
    const styleSel = (prop, val) => commit(() => each(selG(), (c, r) => { const k = key(c, r); const x = S.cells[k] || (S.cells[k] = { raw: null }); x[prop] = val; }));

    /* zone de données contiguë (CurrentRegion) */
    const filled = (c, r) => c >= 0 && r >= 0 && c < NC && r < NR && rawOf(c, r) !== null && rawOf(c, r) !== undefined && rawOf(c, r) !== "";
    function region(c, r) {
      let g = { c1: c, r1: r, c2: c, r2: r }, ch = true;
      if (!filled(c, r)) {
        const nb = [[c, r - 1], [c, r + 1], [c - 1, r], [c + 1, r]].find(([x, y]) => filled(x, y));
        if (!nb) return null;
      }
      while (ch) {
        ch = false;
        const any = (cs, ce, rs, re) => { for (let y = rs; y <= re; y++) for (let x = cs; x <= ce; x++) if (filled(x, y)) return true; return false; };
        if (g.r1 > 0 && any(g.c1 - 1, g.c2 + 1, g.r1 - 1, g.r1 - 1)) { g.r1--; ch = true; }
        if (g.r2 < NR - 1 && any(g.c1 - 1, g.c2 + 1, g.r2 + 1, g.r2 + 1)) { g.r2++; ch = true; }
        if (g.c1 > 0 && any(g.c1 - 1, g.c1 - 1, g.r1 - 1, g.r2 + 1)) { g.c1--; ch = true; }
        if (g.c2 < NC - 1 && any(g.c2 + 1, g.c2 + 1, g.r1 - 1, g.r2 + 1)) { g.c2++; ch = true; }
      }
      return g;
    }

    /* ---------- DOM ---------- */
    host.classList.add("xs");
    host.innerHTML = `
      <div class="xs-win">
        <div class="xs-title"><span class="xs-logo">X</span><b>${esc(setup.sheetName || "Classeur1")}.xlsx</b><span class="xs-saved">Enregistré ✓</span><span class="xs-spacer"></span>
          <span class="xs-dots"><i></i><i></i><i></i></span></div>
        <div class="xs-tabs" role="tablist">
          ${["Accueil", "Insertion", "Formules", "Données"].map((t, i) => `<button type="button" role="tab" data-tab="${i}" class="${i ? "" : "on"}">${t}</button>`).join("")}
        </div>
        <div class="xs-ribbon">
          <div class="xs-rp on" data-rp="0">
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="undo" title="Annuler (Ctrl+Z)">${svg("undo")}</button><button type="button" data-a="redo" title="Rétablir (Ctrl+Y)">${svg("redo")}</button>
            </div><small>Annuler</small></div>
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="bold" title="Gras (Ctrl+G / Ctrl+B)"><b style="font:900 15px Georgia,serif">G</b></button>
              <span class="xs-dd"><button type="button" data-a="fillmenu" title="Couleur de remplissage">${svg("fill")}<i class="xs-fillbar"></i></button>
                <div class="xs-pop xs-swatches">${PALETTE.map(c => `<button type="button" data-fill="${c}" style="background:${c || "#fff"}" title="${c || "Aucun remplissage"}">${c ? "" : "∅"}</button>`).join("")}</div></span>
            </div><small>Police</small></div>
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="al" title="Aligner à gauche">${svg("al")}</button><button type="button" data-a="ac" title="Centrer">${svg("ac")}</button><button type="button" data-a="ar" title="Aligner à droite">${svg("ar")}</button>
            </div><small>Alignement</small></div>
            <div class="xs-g"><div class="xs-gb">
              <select data-a="fmt" title="Format de nombre"><option value="">Standard</option><option value="num">Nombre</option><option value="int">Nombre (0 déc.)</option><option value="eur">Monétaire €</option><option value="pct">Pourcentage</option><option value="date">Date courte</option></select>
              <button type="button" data-a="pct" title="Style de pourcentage (Ctrl+Maj+%)">%</button><button type="button" data-a="eur" title="Format Monétaire">€</button>
            </div><small>Nombre</small></div>
            <div class="xs-g"><div class="xs-gb">
              <span class="xs-dd"><button type="button" data-a="cfmenu" class="wide" title="Mise en forme conditionnelle">${svg("cf")}<span>MFC</span></button>
                <div class="xs-pop xs-form xs-cf"></div></span>
              <button type="button" data-a="table" class="wide" title="Mettre sous forme de tableau">${svg("table")}<span>Tableau</span></button>
            </div><small>Styles</small></div>
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="sum" class="wide" title="Somme automatique (Alt+=)">${svg("sum")}<span>Somme</span></button>
              <button type="button" data-a="filldown" title="Recopier vers le bas (Ctrl+D)">${svg("down")}</button><button type="button" data-a="fillright" title="Recopier à droite (Ctrl+R)">${svg("right")}</button>
            </div><small>Édition</small></div>
          </div>
          <div class="xs-rp" data-rp="1">
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-chart="bar" class="wide" title="Histogramme groupé">${svg("bar")}<span>Histogramme</span></button>
              <button type="button" data-chart="line" class="wide" title="Courbes">${svg("line")}<span>Courbes</span></button>
              <button type="button" data-chart="pie" class="wide" title="Secteurs">${svg("pie")}<span>Secteurs</span></button>
            </div><small>Graphiques (depuis la sélection)</small></div>
            <div class="xs-g"><div class="xs-gb">
              <span class="xs-dd"><button type="button" data-a="pivotmenu" class="wide" title="Tableau croisé dynamique">${svg("pivot")}<span>Tableau croisé dynamique</span></button>
                <div class="xs-pop xs-form xs-pv"></div></span>
            </div><small>Tableaux</small></div>
          </div>
          <div class="xs-rp" data-rp="2">
            <div class="xs-g"><div class="xs-gb">
              <span class="xs-dd"><button type="button" data-a="fnmenu" class="wide" title="Insérer une fonction (Maj+F3)">${svg("fx")}<span>Insérer une fonction</span></button>
                <div class="xs-pop xs-fnlist">${FN_NAMES.map(n => `<button type="button" data-fn="${n}"><b>${n}</b><small>${esc(FN_HELP[n] || "")}</small></button>`).join("")}</div></span>
              <button type="button" data-a="sum" class="wide">${svg("sum")}<span>Somme auto</span></button>
            </div><small>Bibliothèque de fonctions</small></div>
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="showf" class="wide" title="Afficher les formules (Ctrl+$)">${svg("show")}<span>Afficher les formules</span></button>
            </div><small>Vérification des formules</small></div>
          </div>
          <div class="xs-rp" data-rp="3">
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="az" class="wide" title="Trier de A à Z">${svg("az")}<span>Trier A→Z</span></button>
              <button type="button" data-a="za" class="wide" title="Trier de Z à A">${svg("za")}<span>Trier Z→A</span></button>
              <button type="button" data-a="filter" class="wide" title="Filtrer (Ctrl+Maj+L)">${svg("filter")}<span>Filtrer</span></button>
            </div><small>Trier et filtrer</small></div>
            <div class="xs-g"><div class="xs-gb">
              <button type="button" data-a="freeze" class="wide" title="Figer la ligne supérieure">${svg("freeze")}<span>Figer la 1re ligne</span></button>
            </div><small>Affichage</small></div>
          </div>
        </div>
        <div class="xs-fxbar">
          <input class="xs-name" aria-label="Zone Nom" spellcheck="false">
          <span class="xs-fxi">${svg("fx", 16)}</span>
          <input class="xs-fx" aria-label="Barre de formule" spellcheck="false" autocomplete="off">
        </div>
        <div class="xs-gridwrap" tabindex="0" aria-label="Feuille de calcul">
          <table class="xs-grid"><thead><tr><th class="xs-corner"></th>${Array.from({ length: NC }, (_, c) => `<th data-ch="${c}" style="${setup.widths && setup.widths[colName(c)] ? `width:${setup.widths[colName(c)]}px;min-width:${setup.widths[colName(c)]}px` : ""}">${colName(c)}</th>`).join("")}</tr></thead>
          <tbody>${Array.from({ length: NR }, (_, r) => `<tr data-r="${r}"><th data-rh="${r}">${r + 1}</th>${Array.from({ length: NC }, (_, c) => `<td data-c="${c}" data-r="${r}"></td>`).join("")}</tr>`).join("")}</tbody></table>
          <div class="xs-selbox" aria-hidden="true"><i class="xs-handle" title="Poignée de recopie : fais glisser"></i></div>
          <input class="xs-edit" spellcheck="false" autocomplete="off" aria-label="Édition de cellule">
          <div class="xs-ac" hidden></div>
          <div class="xs-tip" hidden></div>
        </div>
        <div class="xs-status"><span class="xs-sheet">${esc(setup.sheetName || "Feuil1")}</span><span class="xs-st-msg">Prêt</span><span class="xs-spacer"></span><span class="xs-agg"></span></div>
      </div>
      <div class="xs-objects"></div>`;

    const $ = s => host.querySelector(s), $$ = s => [...host.querySelectorAll(s)];
    const wrap = $(".xs-gridwrap"), grid = $(".xs-grid"), selbox = $(".xs-selbox"), editor = $(".xs-edit"), fx = $(".xs-fx"), nameBox = $(".xs-name");
    const acBox = $(".xs-ac"), tipBox = $(".xs-tip"), objects = $(".xs-objects"), statusMsg = $(".xs-st-msg");
    const td = (c, r) => grid.querySelector(`td[data-c="${c}"][data-r="${r}"]`);

    /* ---------- rendu ---------- */
    function cfMatch(rule, v) {
      if (rule.op === "bars" || rule.op === "scale") return false;
      if (v === null || isErr(v)) return false;
      const n = typeof v === "number" ? v : parseNumText(v);
      switch (rule.op) {
        case ">": return n !== null && n > +rule.v1;
        case "<": return n !== null && n < +rule.v1;
        case "=": return n !== null && !isNaN(parseFloat(rule.v1)) ? n === +rule.v1 : norm(toText(v)) === norm(rule.v1);
        case "between": return n !== null && n >= Math.min(+rule.v1, +rule.v2) && n <= Math.max(+rule.v1, +rule.v2);
        case "contains": return norm(toText(v)).includes(norm(rule.v1));
      }
      return false;
    }
    function hiddenRows() {
      const h = new Set(); const F = S.filter; if (!F) return h;
      for (let r = F.g.r1 + 1; r <= F.g.r2; r++) {
        for (const [c, excl] of Object.entries(F.excl || {})) if (excl.length && excl.includes(display(getVal(+c, r), (cell(+c, r) || {}).fmt))) { h.add(r); break; }
      }
      return h;
    }
    function render() {
      const hid = hiddenRows();
      const tbl = S.table, flt = S.filter;
      const cfInfo = S.cf.map(rule => {
        const g = parseRange(rule.range); if (!g) return null;
        if (rule.op === "bars" || rule.op === "scale") {
          const vs = []; each(g, (c, r) => { const v = getVal(c, r); if (typeof v === "number") vs.push(v); });
          return { rule, g, min: Math.min(0, ...vs), max: Math.max(...vs, 0), lo: Math.min(...vs), hi: Math.max(...vs) };
        }
        return { rule, g };
      }).filter(Boolean);
      for (let r = 0; r < NR; r++) {
        const tr = grid.tBodies[0].rows[r];
        tr.hidden = hid.has(r);
        for (let c = 0; c < NC; c++) {
          const el = tr.cells[c + 1], x = cell(c, r) || {}, v = getVal(c, r);
          const isF = typeof x.raw === "string" && x.raw[0] === "=";
          let txt = S.showF && isF ? x.raw : display(v, x.fmt);
          let st = "", cls = "";
          const num = typeof v === "number" && !(S.showF && isF);
          const al = x.align || (isErr(v) || typeof v === "boolean" ? "center" : num ? "right" : "left");
          st += `text-align:${al};`;
          if (x.b) st += "font-weight:700;";
          let fill = x.fill || "", color = "";
          if (tbl && c >= tbl.c1 && c <= tbl.c2 && r >= tbl.r1 && r <= tbl.r2) {
            if (r === tbl.r1) { cls += " xs-th"; } else if ((r - tbl.r1) % 2 === 0) cls += " xs-band";
            cls += " xs-intable";
          }
          let bar = "";
          for (const ci of cfInfo) {
            const g = ci.g; if (c < g.c1 || c > g.c2 || r < g.r1 || r > g.r2) continue;
            if (ci.rule.op === "bars" && typeof v === "number" && ci.max > ci.min) bar = `<i class="xs-databar" style="width:${Math.max(2, (v - ci.min) / (ci.max - ci.min) * 100)}%;background:linear-gradient(90deg,${ci.rule.color || "#63BE7B"},${ci.rule.color || "#63BE7B"}55)"></i>`;
            else if (ci.rule.op === "scale" && typeof v === "number") { const t = ci.hi > ci.lo ? (v - ci.lo) / (ci.hi - ci.lo) : .5; fill = t < .5 ? mix("#F8696B", "#FFEB84", t * 2) : mix("#FFEB84", "#63BE7B", (t - .5) * 2); }
            else if (cfMatch(ci.rule, v)) { const s = CF_STYLES.find(z => z.id === ci.rule.style) || CF_STYLES[0]; fill = ci.rule.color || s.fill; color = s.color; }
          }
          if (fill) st += `background:${fill};`;
          if (color) st += `color:${color};`;
          if (isErr(v)) cls += " xs-err";
          if (S.frozen && r === 0) cls += " xs-frozen";
          el.className = cls.trim();
          el.setAttribute("style", st);
          const fbtn = flt && r === flt.g.r1 && c >= flt.g.c1 && c <= flt.g.c2 ? `<button type="button" class="xs-fbtn${(flt.excl && flt.excl[c] && flt.excl[c].length) ? " on" : ""}" data-fcol="${c}" title="Filtrer">▾</button>` : "";
          el.innerHTML = bar + `<span>${esc(txt)}</span>` + fbtn;
          if (isErr(v)) el.title = errHelp(v.c); else el.removeAttribute("title");
        }
      }
      grid.querySelectorAll("thead th[data-ch]").forEach((th, c) => th.classList.toggle("on", c >= sel.c1 && c <= sel.c2));
      grid.querySelectorAll("tbody th[data-rh]").forEach((th, r) => th.classList.toggle("on", r >= sel.r1 && r <= sel.r2));
      grid.classList.toggle("xs-showf", S.showF);
      $('[data-a="showf"]').classList.toggle("on", S.showF);
      $('[data-a="filter"]').classList.toggle("on", !!S.filter);
      $('[data-a="table"]').classList.toggle("on", !!S.table);
      $('[data-a="freeze"]').classList.toggle("on", S.frozen);
      placeSel(); syncBars(); renderObjects();
    }
    function mix(a, b, t) {
      const p = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16)), A = p(a), B = p(b);
      return "#" + A.map((x, i) => Math.round(x + (B[i] - x) * t).toString(16).padStart(2, "0")).join("");
    }
    function errHelp(c) {
      return { "#DIV/0!": "Division par zéro : le diviseur est vide ou nul.", "#N/A": "Valeur introuvable (recherche sans résultat).",
        "#NOM?": "Nom de fonction inconnu ou texte sans guillemets.", "#VALEUR!": "Type de valeur incorrect (texte dans un calcul ?).",
        "#REF!": "Référence invalide ou circulaire.", "#NOMBRE!": "Nombre invalide." }[c] || c;
    }
    function placeSel() {
      const a = td(sel.c1, sel.r1), b = td(sel.c2, sel.r2);
      if (!a || !b) return;
      const wr = wrap.getBoundingClientRect(), ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
      const x = ra.left - wr.left + wrap.scrollLeft, y = ra.top - wr.top + wrap.scrollTop;
      selbox.style.cssText = `left:${x - 1}px;top:${y - 1}px;width:${rb.right - ra.left + 1}px;height:${rb.bottom - ra.top + 1}px`;
      selbox.classList.toggle("multi", sel.c1 !== sel.c2 || sel.r1 !== sel.r2);
      grid.querySelectorAll("td.xs-act").forEach(e => e.classList.remove("xs-act"));
      const act = td(sel.ac, sel.ar); if (act) act.classList.add("xs-act");
      grid.querySelectorAll("td.xs-insel").forEach(e => e.classList.remove("xs-insel"));
      if (sel.c1 !== sel.c2 || sel.r1 !== sel.r2) each(selG(), (c, r) => { if (c !== sel.ac || r !== sel.ar) td(c, r).classList.add("xs-insel"); });
    }
    function syncBars() {
      nameBox.value = sel.c1 === sel.c2 && sel.r1 === sel.r2 ? refName(sel.ac, sel.ar) : rangeName(selG());
      if (!edit) { const raw = rawOf(sel.ac, sel.ar); const x = cell(sel.ac, sel.ar) || {}; fx.value = raw === null || raw === undefined ? "" : typeof raw === "number" ? (x.fmt === "pct" ? numFr(+(raw * 100).toPrecision(12)) + " %" : x.fmt === "date" ? display(raw, "date") : numFr(raw)) : typeof raw === "boolean" ? (raw ? "VRAI" : "FAUX") : raw; }
      const x = cell(sel.ac, sel.ar) || {};
      $('[data-a="fmt"]').value = x.fmt || "";
      $('[data-a="bold"]').classList.toggle("on", !!x.b);
      $(".xs-fillbar").style.background = x.fill || "#FFE699";
      // agrégats de la barre d'état
      const vs = []; let cnt = 0; each(selG(), (c, r) => { const v = getVal(c, r); if (v !== null && v !== "") cnt++; if (typeof v === "number") vs.push(v); });
      const agg = $(".xs-agg");
      if (cnt > 1) agg.innerHTML = `${vs.length ? `<span>Moyenne : <b>${display(vs.reduce((s, x) => s + x, 0) / vs.length, x.fmt)}</b></span>` : ""}<span>Nb : <b>${cnt}</b></span>${vs.length ? `<span>Somme : <b>${display(vs.reduce((s, x) => s + x, 0), (cell(sel.c1, sel.r1) || {}).fmt)}</b></span>` : ""}`;
      else agg.innerHTML = "";
    }

    /* ---------- graphiques & TCD ---------- */
    function chartData(rng) {
      const g = parseRange(rng); if (!g) return null;
      const V = (c, r) => getVal(c, r);
      const headRow = (() => { for (let c = g.c1 + (g.c2 > g.c1 ? 1 : 0); c <= g.c2; c++) if (typeof V(c, g.r1) === "string") return true; return false; })();
      const labCol = g.c2 > g.c1 && (() => { for (let r = g.r1 + (headRow ? 1 : 0); r <= g.r2; r++) if (typeof V(g.c1, r) === "string") return true; return false; })();
      const r0 = g.r1 + (headRow ? 1 : 0), c0 = g.c1 + (labCol ? 1 : 0);
      const cats = []; for (let r = r0; r <= g.r2; r++) cats.push(labCol ? display(V(g.c1, r), (cell(g.c1, r) || {}).fmt) : String(r - r0 + 1));
      const series = [];
      for (let c = c0; c <= g.c2; c++) {
        const vals = []; for (let r = r0; r <= g.r2; r++) { const v = V(c, r); vals.push(typeof v === "number" ? v : 0); }
        series.push({ name: headRow ? toText(V(c, g.r1)) : "Série " + (c - c0 + 1), vals, fmt: (cell(c, r0) || {}).fmt });
      }
      return { cats, series };
    }
    function chartSVG(ch, i) {
      const d = chartData(ch.range); if (!d || !d.series.length || !d.cats.length) return "<p class='xs-empty'>Plage vide</p>";
      const W = 460, H = 250, L = 52, B = 36, T = 16, R = 12, pw = W - L - R, ph = H - T - B;
      const fmtN = (v, f) => display(v, f === "pct" ? "pct" : f === "eur" ? "eur" : "").replace(",00 €", " €");
      if (ch.type === "pie") {
        const s = d.series[0], tot = s.vals.reduce((a, b) => a + Math.max(0, b), 0) || 1; let a0 = -Math.PI / 2;
        const cx = 120, cy = H / 2, rr = 92;
        const slices = s.vals.map((v, k) => {
          const a1 = a0 + Math.max(0, v) / tot * Math.PI * 2, large = a1 - a0 > Math.PI ? 1 : 0;
          const p = `M${cx} ${cy}L${cx + rr * Math.cos(a0)} ${cy + rr * Math.sin(a0)}A${rr} ${rr} 0 ${large} 1 ${cx + rr * Math.cos(a1)} ${cy + rr * Math.sin(a1)}Z`;
          const mid = (a0 + a1) / 2; a0 = a1;
          return `<path d="${p}" fill="${SERIES[k % SERIES.length]}" class="xs-slice" style="--d:${k * 90}ms;--tx:${Math.cos(mid) * 6}px;--ty:${Math.sin(mid) * 6}px"><title>${esc(d.cats[k])} : ${fmtN(v, s.fmt)} (${Math.round(v / tot * 100)} %)</title></path>`;
        }).join("");
        const leg = d.cats.map((c, k) => `<g transform="translate(250 ${30 + k * 20})"><rect width="11" height="11" rx="3" fill="${SERIES[k % SERIES.length]}"/><text x="18" y="10">${esc(c)} — ${Math.round(Math.max(0, s.vals[k]) / tot * 100)} %</text></g>`).join("");
        return `<svg viewBox="0 0 ${W} ${H}" class="xs-chart">${slices}${leg}</svg>`;
      }
      const all = d.series.flatMap(s => s.vals), mx = Math.max(0, ...all), mn = Math.min(0, ...all);
      const step = niceStep((mx - mn) / 4 || 1), top = Math.ceil(mx / step) * step || step, bot = Math.floor(mn / step) * step;
      const y = v => T + ph - (v - bot) / (top - bot) * ph;
      let gridL = "";
      for (let v = bot; v <= top + 1e-9; v += step) gridL += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="xs-gl"/><text x="${L - 6}" y="${y(v) + 4}" text-anchor="end">${fmtN(+v.toPrecision(10), d.series[0].fmt)}</text>`;
      const n = d.cats.length, bw = pw / n;
      const cats = d.cats.map((c, k) => `<text x="${L + bw * k + bw / 2}" y="${H - B + 16}" text-anchor="middle">${esc(c.length > 11 ? c.slice(0, 10) + "…" : c)}</text>`).join("");
      let marks = "";
      if (ch.type === "bar") {
        const gw = bw * .72, sw = gw / d.series.length;
        d.series.forEach((s, si) => s.vals.forEach((v, k) => {
          const x = L + bw * k + (bw - gw) / 2 + sw * si, y0 = y(Math.max(0, v)), h = Math.abs(y(v) - y(0));
          marks += `<rect x="${x + 1}" y="${y0}" width="${Math.max(2, sw - 2)}" height="${Math.max(0, h)}" rx="3" fill="${SERIES[si % SERIES.length]}" class="xs-barm" style="--d:${(k * d.series.length + si) * 45}ms;transform-origin:0 ${y(0)}px"><title>${esc(s.name)} · ${esc(d.cats[k])} : ${fmtN(v, s.fmt)}</title></rect>`;
        }));
      } else {
        d.series.forEach((s, si) => {
          const pts = s.vals.map((v, k) => [L + bw * k + bw / 2, y(v)]);
          marks += `<path d="M${pts.map(p => p.join(" ")).join("L")}" fill="none" stroke="${SERIES[si % SERIES.length]}" stroke-width="3" stroke-linejoin="round" stroke-linecap="round" class="xs-linem" pathLength="1"/>`;
          marks += pts.map((p, k) => `<circle cx="${p[0]}" cy="${p[1]}" r="4.5" fill="#fff" stroke="${SERIES[si % SERIES.length]}" stroke-width="2.5" class="xs-dot" style="--d:${300 + k * 60}ms"><title>${esc(s.name)} · ${esc(d.cats[k])} : ${fmtN(s.vals[k], s.fmt)}</title></circle>`).join("");
        });
      }
      const leg = d.series.length > 1 ? d.series.map((s, si) => `<g transform="translate(${L + si * 110} 4)"><rect width="10" height="10" rx="3" fill="${SERIES[si % SERIES.length]}"/><text x="15" y="9">${esc(s.name)}</text></g>`).join("") : "";
      return `<svg viewBox="0 0 ${W} ${H}" class="xs-chart">${gridL}<line x1="${L}" x2="${W - R}" y1="${y(0)}" y2="${y(0)}" class="xs-axis"/>${marks}${cats}${leg}</svg>`;
    }
    const niceStep = x => { const p = Math.pow(10, Math.floor(Math.log10(x))), f = x / p; return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * p; };
    function pivotHTML(pv) {
      const g = parseRange(pv.src); if (!g) return "";
      const heads = []; for (let c = g.c1; c <= g.c2; c++) heads.push(toText(getVal(c, g.r1)));
      const rc = heads.indexOf(pv.rows) + g.c1, vc = heads.indexOf(pv.values) + g.c1;
      if (rc < g.c1 || vc < g.c1) return "<p class='xs-empty'>Champs introuvables</p>";
      const m = new Map();
      for (let r = g.r1 + 1; r <= g.r2; r++) {
        const k = toText(getVal(rc, r)); if (k === "") continue; const v = getVal(vc, r);
        const o = m.get(k) || { s: 0, n: 0, cnt: 0 }; o.cnt++; if (typeof v === "number") { o.s += v; o.n++; } m.set(k, o);
      }
      const val = o => pv.agg === "count" ? o.cnt : pv.agg === "avg" ? (o.n ? o.s / o.n : 0) : o.s;
      const keys = [...m.keys()].sort((a, b) => compare(a, b));
      const tot = { s: 0, n: 0, cnt: 0 }; m.forEach(o => { tot.s += o.s; tot.n += o.n; tot.cnt += o.cnt; });
      const f = pv.agg === "count" ? "" : (cell(vc, g.r1 + 1) || {}).fmt;
      const lab = { sum: "Somme de ", count: "Nombre de ", avg: "Moyenne de " }[pv.agg] + pv.values;
      const mx = Math.max(...keys.map(k => val(m.get(k))), 1);
      return `<table class="xs-pt"><thead><tr><th>Étiquettes de lignes</th><th>${esc(lab)}</th></tr></thead><tbody>${keys.map((k, i) => `<tr style="--d:${i * 50}ms"><td>${esc(k)}</td><td><i style="width:${val(m.get(k)) / mx * 100}%"></i><span>${display(val(m.get(k)), f || (pv.agg === "avg" ? "num" : ""))}</span></td></tr>`).join("")}</tbody>
        <tfoot><tr><td>Total général</td><td><span>${display(val(tot), f || (pv.agg === "avg" ? "num" : ""))}</span></td></tr></tfoot></table>`;
    }
    let objSig = "";
    function renderObjects() {
      const sig = JSON.stringify([S.charts, S.pivots]);
      const fresh = sig !== objSig; objSig = sig;
      if (!S.charts.length && !S.pivots.length) { objects.innerHTML = ""; return; }
      objects.innerHTML = S.charts.map((ch, i) => `<figure class="xs-obj${fresh && i === S.charts.length - 1 ? " xs-new" : ""}"><figcaption><b>${{ bar: "Histogramme", line: "Courbes", pie: "Secteurs" }[ch.type]}</b> · ${esc(ch.range)}<button type="button" data-delchart="${i}" title="Supprimer">×</button></figcaption>${chartSVG(ch, i)}</figure>`).join("")
        + S.pivots.map((pv, i) => `<figure class="xs-obj"><figcaption><b>Tableau croisé dynamique</b> · ${esc(pv.rows)} × ${esc(pv.values)}<button type="button" data-delpivot="${i}" title="Supprimer">×</button></figcaption>${pivotHTML(pv)}</figure>`).join("");
      if (!fresh) objects.querySelectorAll(".xs-barm,.xs-linem,.xs-dot,.xs-slice,.xs-pt tr").forEach(e => e.style.animation = "none");
    }

    /* ---------- sélection & navigation ---------- */
    function select(c, r, extend) {
      c = Math.max(0, Math.min(NC - 1, c)); r = Math.max(0, Math.min(NR - 1, r));
      if (extend) { sel.c1 = Math.min(sel.ac, c); sel.c2 = Math.max(sel.ac, c); sel.r1 = Math.min(sel.ar, r); sel.r2 = Math.max(sel.ar, r); sel.fc = c; sel.fr = r; }
      else sel = { ac: c, ar: r, c1: c, r1: r, c2: c, r2: r, fc: c, fr: r };
      grid.querySelectorAll("thead th[data-ch]").forEach((th, i) => th.classList.toggle("on", i >= sel.c1 && i <= sel.c2));
      grid.querySelectorAll("tbody th[data-rh]").forEach((th, i) => th.classList.toggle("on", i >= sel.r1 && i <= sel.r2));
      placeSel(); syncBars();
      const t = td(extend ? c : sel.ac, extend ? r : sel.ar); if (t && t.scrollIntoView) { const wr = wrap.getBoundingClientRect(), tr = t.getBoundingClientRect(); if (tr.bottom > wr.bottom || tr.top < wr.top + 26 || tr.right > wr.right || tr.left < wr.left + 40) t.scrollIntoView({ block: "nearest", inline: "nearest" }); }
    }
    function move(dc, dr, extend) {
      if (extend) select((sel.fc ?? sel.ac) + dc, (sel.fr ?? sel.ar) + dr, true);
      else select(sel.ac + dc, sel.ar + dr);
    }

    /* ---------- édition ---------- */
    let pointIns = null; // { s, e } texte de référence inséré par clic pendant la saisie d'une formule
    function startEdit(initial, mode) {
      edit = { c: sel.ac, r: sel.ar, mode: mode || "enter", via: "cell" };
      const t = td(sel.ac, sel.ar), wr = wrap.getBoundingClientRect(), rr = t.getBoundingClientRect();
      editor.style.cssText = `display:block;left:${rr.left - wr.left + wrap.scrollLeft}px;top:${rr.top - wr.top + wrap.scrollTop}px;min-width:${rr.width}px;height:${rr.height}px`;
      editor.value = initial;
      fx.value = initial;
      editor.focus();
      const L = editor.value.length; editor.setSelectionRange(L, L);
      pointIns = null; onEditInput();
      statusMsg.textContent = mode === "edit" ? "Modifier" : "Entrer";
    }
    function endEdit(save, dc, dr) {
      if (!edit) return;
      const { c, r } = edit, v = (edit.via === "fx" ? fx : editor).value;
      edit = null; editor.style.display = "none"; acBox.hidden = true; tipBox.hidden = true; clearRefHi();
      statusMsg.textContent = "Prêt";
      if (save) {
        let val = v;
        if (/^=/.test(val)) { const open = (val.match(/\(/g) || []).length - (val.match(/\)/g) || []).length; if (open > 0) val += ")".repeat(open); }
        const before = rawOf(c, r);
        const pi = parseInput(val);
        if (pi.raw !== before || (pi.fmt && !(cell(c, r) || {}).fmt)) commit(() => setRaw(c, r, val));
        else render();
        const v2 = getVal(c, r);
        if (isErr(v2)) { const t = td(c, r); t.classList.remove("xs-shake"); void t.offsetWidth; t.classList.add("xs-shake"); }
        else if (typeof val === "string" && val[0] === "=" && !reduce) { const t = td(c, r); t.classList.remove("xs-flash"); void t.offsetWidth; t.classList.add("xs-flash"); }
      } else syncBars();
      if (dc || dr) move(dc || 0, dr || 0);
      wrap.focus({ preventScroll: true });
    }
    const activeInput = () => (edit && edit.via === "fx" ? fx : editor);
    function onEditInput() {
      const inp = activeInput(), v = inp.value;
      (inp === fx ? editor : fx).value = v;
      if (v[0] === "=") { highlightRefs(v); autocomplete(inp); } else { clearRefHi(); acBox.hidden = true; tipBox.hidden = true; }
    }
    const REFCOLORS = ["#2B7CD3", "#D35230", "#8E44AD", "#21A366", "#C27C0E", "#0097A7"];
    function clearRefHi() { host.querySelectorAll(".xs-refhi").forEach(e => e.remove()); }
    function highlightRefs(f) {
      clearRefHi();
      const re = /("(?:[^"]|"")*")|(\$?[A-Z]{1,2}\$?\d+)(?::(\$?[A-Z]{1,2}\$?\d+))?(?![A-Z0-9_(])/gi; let m, k = 0;
      const wr = wrap.getBoundingClientRect();
      while ((m = re.exec(f))) {
        if (m[1]) continue;
        const g = parseRange(m[2] + (m[3] ? ":" + m[3] : "")); if (!g || g.c2 >= NC || g.r2 >= NR) continue;
        const a = td(g.c1, g.r1).getBoundingClientRect(), b = td(g.c2, g.r2).getBoundingClientRect(), col = REFCOLORS[k++ % REFCOLORS.length];
        const d = document.createElement("div"); d.className = "xs-refhi";
        d.style.cssText = `left:${a.left - wr.left + wrap.scrollLeft}px;top:${a.top - wr.top + wrap.scrollTop}px;width:${b.right - a.left}px;height:${b.bottom - a.top}px;--c:${col}`;
        wrap.append(d);
      }
    }
    function autocomplete(inp) {
      const pos = inp.selectionStart, before = inp.value.slice(0, pos);
      const m = /([A-ZÀ-Ü][A-ZÀ-Ü.]*)$/i.exec(before);
      const host2 = inp === fx ? fx : editor, r = host2.getBoundingClientRect(), wr = wrap.getBoundingClientRect();
      const placeBox = box => { box.style.left = Math.max(0, r.left - wr.left + wrap.scrollLeft) + "px"; box.style.top = (r.bottom - wr.top + wrap.scrollTop + 2) + "px"; };
      if (m && /[=(;:+\-*/^&<>,]$/.test(before.slice(0, pos - m[1].length)) && !/\d/.test(m[1])) {
        const q = m[1].toUpperCase(), list = FN_NAMES.filter(n => n.startsWith(q) || (Object.keys(ALIAS).find(a => a.startsWith(q) && ALIAS[a] === n))).slice(0, 7);
        if (list.length) {
          acBox.innerHTML = list.map((n, i) => `<button type="button" data-ac="${n}" class="${i ? "" : "on"}"><b>${n}</b><small>${esc((FN_HELP[n] || "").replace(/^[^(]+/, ""))}</small></button>`).join("");
          acBox.hidden = false; tipBox.hidden = true; placeBox(acBox); acBox.dataset.len = m[1].length; return;
        }
      }
      acBox.hidden = true;
      // info-bulle de syntaxe : fonction ouverte la plus proche
      let depth = 0, fn = null;
      for (let i = pos - 1; i >= 0; i--) {
        const ch = before[i];
        if (ch === ")") depth++;
        else if (ch === "(") { if (depth === 0) { const mm = /([A-ZÀ-Ü][A-ZÀ-Ü0-9.]*)$/i.exec(before.slice(0, i)); fn = mm && canon(mm[1]); break; } depth--; }
      }
      if (fn && FN_HELP[fn]) {
        const argN = (() => { let d = 0, n = 0; const s = before.slice(before.toUpperCase().lastIndexOf(fn.length ? "(" : "(")); for (const ch of before.slice(before.lastIndexOf("(", pos))) { if (ch === "(") d++; else if (ch === ")") d--; else if (ch === ";" && d === 1) n++; } return n; })();
        const parts = FN_HELP[fn].split(/(\(|;|\))/);
        let ai = -1; const html = parts.map(p => { if (p === "(") { ai = 0; return p; } if (p === ";") { ai++; return p; } if (p === ")") { ai = -9; return p; } return ai === argN ? `<b>${esc(p)}</b>` : esc(p); }).join("");
        tipBox.innerHTML = html; tipBox.hidden = false; placeBox(tipBox);
      } else tipBox.hidden = true;
    }
    function acceptAC(name) {
      const inp = activeInput(), pos = inp.selectionStart, n = +acBox.dataset.len || 0;
      inp.value = inp.value.slice(0, pos - n) + name + "(" + inp.value.slice(pos);
      const p = pos - n + name.length + 1; inp.setSelectionRange(p, p); acBox.hidden = true; inp.focus(); onEditInput();
    }
    const canPoint = inp => {
      if (!inp.value.startsWith("=")) return false;
      if (pointIns && inp.selectionStart === pointIns.e) return true;
      return /[=(;:+\-*/^&<>,]\s*$/.test(inp.value.slice(0, inp.selectionStart));
    };
    function insertRef(txt) {
      const inp = activeInput();
      let s = inp.selectionStart, e = inp.selectionEnd;
      if (pointIns && s === pointIns.e) s = pointIns.s;
      inp.value = inp.value.slice(0, s) + txt + inp.value.slice(e);
      pointIns = { s, e: s + txt.length };
      inp.setSelectionRange(pointIns.e, pointIns.e); onEditInput();
    }
    function f4() {
      const inp = activeInput(), v = inp.value, pos = inp.selectionStart;
      const re = /\$?[A-Z]{1,2}\$?\d+/gi; let m;
      while ((m = re.exec(v))) {
        if (pos >= m.index && pos <= m.index + m[0].length) {
          const nr = cycleAbs(m[0]); inp.value = v.slice(0, m.index) + nr + v.slice(m.index + m[0].length);
          const p = m.index + nr.length; inp.setSelectionRange(p, p); if (pointIns) pointIns = { s: m.index, e: p };
          onEditInput(); return;
        }
      }
    }

    /* ---------- recopie ---------- */
    function copyCell(sc, sr, dc, dr) {
      const s = cell(sc, sr), k = key(dc, dr);
      if (!s) { delete S.cells[k]; return; }
      const n = JSON.parse(JSON.stringify(s));
      if (typeof n.raw === "string" && n.raw[0] === "=") n.raw = shiftFormula(n.raw, dc - sc, dr - sr);
      S.cells[k] = n;
    }
    function fillTo(src, dst) { // src = sélection d'origine, dst = sélection étendue (vers le bas ou la droite)
      commit(() => {
        if (dst.r2 > src.r2 || dst.r1 < src.r1) {
          for (let c = src.c1; c <= src.c2; c++) {
            const seed = []; for (let r = src.r1; r <= src.r2; r++) seed.push(rawOf(c, r));
            const series = seed.length >= 2 && seed.every(v => typeof v === "number");
            const st = series ? seed[1] - seed[0] : 0;
            for (let r = dst.r1; r <= dst.r2; r++) {
              if (r >= src.r1 && r <= src.r2) continue;
              if (series) { copyCell(c, src.r1, c, r); S.cells[key(c, r)].raw = seed[0] + st * (r - src.r1); }
              else { const off = ((r - src.r1) % seed.length + seed.length) % seed.length; copyCell(c, src.r1 + off, c, r); }
            }
          }
        } else {
          for (let r = src.r1; r <= src.r2; r++) {
            const n = src.c2 - src.c1 + 1;
            for (let c = dst.c1; c <= dst.c2; c++) { if (c >= src.c1 && c <= src.c2) continue; copyCell(src.c1 + (((c - src.c1) % n) + n) % n, r, c, r); }
          }
        }
      }, "Recopié ✓");
      sel.c1 = dst.c1; sel.c2 = dst.c2; sel.r1 = dst.r1; sel.r2 = dst.r2; placeSel(); syncBars();
    }
    function fillDown() { const g = selG(); if (g.r1 === g.r2) { if (g.r1 === 0) return; g.r1--; } commit(() => { for (let c = g.c1; c <= g.c2; c++) for (let r = g.r1 + 1; r <= g.r2; r++) copyCell(c, g.r1, c, r); }, "Recopié vers le bas ✓"); }
    function fillRight() { const g = selG(); if (g.c1 === g.c2) { if (g.c1 === 0) return; g.c1--; } commit(() => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1 + 1; c <= g.c2; c++) copyCell(g.c1, r, c, r); }, "Recopié vers la droite ✓"); }

    /* ---------- commandes ---------- */
    function autoSum() {
      const g = selG();
      if (g.c1 === g.c2 && g.r1 === g.r2) {
        const c = sel.ac, r = sel.ar; let r0 = r - 1;
        while (r0 >= 0 && typeof getVal(c, r0) === "number") r0--;
        if (r0 < r - 1) return startEdit(`=SOMME(${refName(c, r0 + 1)}:${refName(c, r - 1)})`, "enter");
        let c0 = c - 1; while (c0 >= 0 && typeof getVal(c0, r) === "number") c0--;
        if (c0 < c - 1) return startEdit(`=SOMME(${refName(c0 + 1, r)}:${refName(c - 1, r)})`, "enter");
        return startEdit("=SOMME()", "enter");
      }
      if (g.r2 + 1 >= NR) return;
      commit(() => { for (let c = g.c1; c <= g.c2; c++) setRaw(c, g.r2 + 1, `=SOMME(${refName(c, g.r1)}:${refName(c, g.r2)})`); }, "Somme automatique ✓");
    }
    function sortBy(dir) {
      const g = region(sel.ac, sel.ar);
      if (!g || g.r2 <= g.r1) return ctx.toast("Place-toi dans une colonne du tableau à trier");
      const col = sel.ac, rows = [];
      for (let r = g.r1 + 1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(cell(c, r) ? JSON.parse(JSON.stringify(cell(c, r))) : null); rows.push({ r, row, v: getVal(col, r) }); }
      rows.sort((a, b) => { if (a.v === null) return 1; if (b.v === null) return -1; const c = compare(isErr(a.v) ? "" : a.v, isErr(b.v) ? "" : b.v); return dir === "asc" ? c : -c; });
      commit(() => rows.forEach((o, i) => {
        const nr = g.r1 + 1 + i;
        o.row.forEach((x, j) => { const k = key(g.c1 + j, nr); if (!x) { delete S.cells[k]; return; } if (typeof x.raw === "string" && x.raw[0] === "=") x.raw = shiftFormula(x.raw, 0, nr - o.r); S.cells[k] = x; });
      }), dir === "asc" ? "Trié de A à Z ✓" : "Trié de Z à A ✓");
    }
    function toggleFilter() {
      if (S.filter) return commit(() => { S.filter = null; }, "Filtre désactivé");
      const g = region(sel.ac, sel.ar) || region(0, 0);
      if (!g) return ctx.toast("Sélectionne une cellule du tableau");
      commit(() => { S.filter = { g, excl: {} }; }, "Filtre activé : clique sur ▾ dans les en-têtes");
    }
    function toggleTable() {
      if (S.table) return commit(() => { S.table = null; }, "Tableau converti en plage");
      const g = region(sel.ac, sel.ar);
      if (!g) return ctx.toast("Sélectionne une cellule du tableau");
      commit(() => { S.table = g; if (!S.filter) S.filter = { g, excl: {} }; }, "Mis sous forme de tableau ✓");
    }
    function addChart(type) {
      let g = selG();
      if (g.c1 === g.c2 && g.r1 === g.r2) { const rg = region(sel.ac, sel.ar); if (!rg) return ctx.toast("Sélectionne d'abord les données du graphique"); g = rg; }
      commit(() => { S.charts.push({ type, range: rangeName(g) }); }, "Graphique inséré ✓");
      ctx.confetti && setTimeout(() => { const f = objects.lastElementChild; f && f.scrollIntoView({ block: "nearest", behavior: reduce ? "auto" : "smooth" }); }, 50);
    }

    /* ---------- popovers ---------- */
    function closePops(except) { $$(".xs-dd.open").forEach(d => { if (d !== except) d.classList.remove("open"); }); }
    function openCF(dd) {
      const box = dd.querySelector(".xs-cf"), g = selG();
      box.innerHTML = `<h5>Mise en forme conditionnelle</h5><label>Plage <input data-k="range" value="${rangeName(g)}"></label>
        <label>Règle <select data-k="op"><option value=">">Supérieur à…</option><option value="<">Inférieur à…</option><option value="between">Entre…</option><option value="=">Égal à…</option><option value="contains">Texte qui contient…</option><option value="bars">Barres de données</option><option value="scale">Nuances de couleurs (rouge→vert)</option></select></label>
        <div class="xs-row"><input data-k="v1" placeholder="Valeur"><input data-k="v2" placeholder="et" hidden></div>
        <label>Format <select data-k="style">${CF_STYLES.map(s => `<option value="${s.id}">${s.label}</option>`).join("")}</select></label>
        <div class="xs-row"><button type="button" class="xs-ok" data-cfok>Appliquer</button><button type="button" data-cfclear>Effacer les règles</button></div>
        ${S.cf.length ? `<ul class="xs-rules">${S.cf.map((r, i) => `<li><span>${esc(r.range)} : ${esc({ ">": "> ", "<": "< ", "=": "= ", between: "entre ", contains: "contient ", bars: "barres", scale: "nuances" }[r.op] + (r.v1 ?? "") + (r.op === "between" ? " et " + r.v2 : ""))}</span><button type="button" data-cfdel="${i}">×</button></li>`).join("")}</ul>` : ""}`;
      const op = box.querySelector('[data-k="op"]');
      op.onchange = () => { box.querySelector('[data-k="v2"]').hidden = op.value !== "between"; const nv = op.value === "bars" || op.value === "scale"; box.querySelector('[data-k="v1"]').hidden = nv; box.querySelector('[data-k="style"]').parentElement.hidden = nv; };
    }
    function openPivot(dd) {
      const box = dd.querySelector(".xs-pv");
      const g = region(sel.ac, sel.ar) || region(0, 0);
      if (!g) { box.innerHTML = "<p>Sélectionne une cellule de tes données (avec une ligne d'en-têtes).</p>"; return; }
      const heads = []; for (let c = g.c1; c <= g.c2; c++) heads.push(toText(getVal(c, g.r1)));
      const opts = heads.map(h => `<option>${esc(h)}</option>`).join("");
      box.innerHTML = `<h5>Créer un tableau croisé dynamique</h5><p class="xs-mut">Source : <b>${rangeName(g)}</b></p>
        <label>Lignes <select data-k="rows">${opts}</select></label>
        <label>Valeurs <select data-k="values">${opts}</select></label>
        <label>Synthèse <select data-k="agg"><option value="sum">Somme</option><option value="count">Nombre</option><option value="avg">Moyenne</option></select></label>
        <div class="xs-row"><button type="button" class="xs-ok" data-pvok data-src="${rangeName(g)}">Créer</button></div>`;
      const sv = box.querySelector('[data-k="values"]');
      const firstNum = heads.findIndex((h, i) => typeof getVal(g.c1 + i, g.r1 + 1) === "number"); if (firstNum >= 0) sv.selectedIndex = firstNum;
    }
    function openFilter(c, btn) {
      host.querySelectorAll(".xs-fpop").forEach(e => e.remove());
      const F = S.filter; const vals = new Set();
      for (let r = F.g.r1 + 1; r <= F.g.r2; r++) vals.add(display(getVal(c, r), (cell(c, r) || {}).fmt));
      const excl = (F.excl && F.excl[c]) || [];
      const pop = document.createElement("div"); pop.className = "xs-pop xs-fpop xs-form";
      const list = [...vals].sort((a, b) => compare(parseNumText(a) ?? a, parseNumText(b) ?? b));
      pop.innerHTML = `<h5>Filtrer ${esc(toText(getVal(c, F.g.r1)))}</h5><div class="xs-fl">${list.map(v => `<label><input type="checkbox" value="${esc(v)}" ${excl.includes(v) ? "" : "checked"}> ${esc(v || "(Vides)")}</label>`).join("")}</div>
        <div class="xs-row"><button type="button" class="xs-ok" data-fok="${c}">OK</button><button type="button" data-fall>Tout</button></div>`;
      wrap.append(pop);
      const wr = wrap.getBoundingClientRect(), br = btn.getBoundingClientRect();
      pop.style.left = Math.min(br.left - wr.left + wrap.scrollLeft - 150, wrap.scrollWidth - 230) + "px"; pop.style.top = (br.bottom - wr.top + wrap.scrollTop + 4) + "px";
      pop.querySelector("[data-fall]").onclick = () => pop.querySelectorAll("input").forEach(i => { i.checked = true; });
      pop.querySelector("[data-fok]").onclick = () => {
        const ex = [...pop.querySelectorAll("input")].filter(i => !i.checked).map(i => i.value);
        pop.remove(); commit(() => { S.filter.excl = Object.assign({}, S.filter.excl, { [c]: ex }); }, ex.length ? "Filtre appliqué" : "Filtre effacé");
      };
    }

    /* ---------- événements ---------- */
    host.addEventListener("click", e => {
      const t = e.target.closest("button, [data-tab]");
      if (!e.target.closest(".xs-dd")) closePops();
      if (!t || !host.contains(t)) return;
      if (t.dataset.tab !== undefined) {
        $$(".xs-tabs button").forEach(b => b.classList.toggle("on", b === t));
        $$(".xs-rp").forEach(p => p.classList.toggle("on", p.dataset.rp === t.dataset.tab));
        return;
      }
      const a = t.dataset.a;
      if (t.dataset.fill !== undefined) { closePops(); return styleSel("fill", t.dataset.fill); }
      if (t.dataset.chart) return addChart(t.dataset.chart);
      if (t.dataset.fn) { closePops(); if (!edit) startEdit("=" + t.dataset.fn + "(", "enter"); else { const inp = activeInput(); const p = inp.selectionStart; inp.value = inp.value.slice(0, p) + t.dataset.fn + "(" + inp.value.slice(p); inp.focus(); onEditInput(); } return; }
      if (t.dataset.ac) return acceptAC(t.dataset.ac);
      if (t.dataset.delchart !== undefined) return commit(() => S.charts.splice(+t.dataset.delchart, 1));
      if (t.dataset.delpivot !== undefined) return commit(() => S.pivots.splice(+t.dataset.delpivot, 1));
      if (t.dataset.fcol !== undefined) { e.stopPropagation(); return openFilter(+t.dataset.fcol, t); }
      if (t.hasAttribute("data-cfok")) {
        const box = t.closest(".xs-cf"), v = k => box.querySelector(`[data-k="${k}"]`).value;
        const op = v("op"), range = v("range").toUpperCase().replace(/\$/g, "");
        if (!parseRange(range)) return ctx.toast("Plage invalide");
        if ((op !== "bars" && op !== "scale") && v("v1").trim() === "") return ctx.toast("Indique une valeur");
        const st = CF_STYLES.find(s => s.id === v("style"));
        closePops();
        return commit(() => S.cf.push({ range, op, v1: op === "contains" || op === "=" ? v("v1") : parseFloat(String(v("v1")).replace(",", ".")), v2: op === "between" ? parseFloat(String(v("v2")).replace(",", ".")) : undefined, style: st.id, color: op === "bars" ? "#63BE7B" : op === "scale" ? "" : st.fill }), "Règle appliquée ✓");
      }
      if (t.hasAttribute("data-cfclear")) { closePops(); return commit(() => { S.cf = []; }, "Règles effacées"); }
      if (t.dataset.cfdel !== undefined) { commit(() => S.cf.splice(+t.dataset.cfdel, 1)); return openCF(t.closest(".xs-dd")); }
      if (t.hasAttribute("data-pvok")) {
        const box = t.closest(".xs-pv"), v = k => box.querySelector(`[data-k="${k}"]`).value;
        closePops();
        return commit(() => S.pivots.push({ src: t.dataset.src, rows: v("rows"), values: v("values"), agg: v("agg") }), "Tableau croisé dynamique créé ✓");
      }
      switch (a) {
        case "undo": return undo();
        case "redo": return redo();
        case "bold": return styleSel("b", !(cell(sel.ac, sel.ar) || {}).b);
        case "al": return styleSel("align", "left");
        case "ac": return styleSel("align", "center");
        case "ar": return styleSel("align", "right");
        case "pct": return styleSel("fmt", "pct");
        case "eur": return styleSel("fmt", "eur");
        case "sum": return autoSum();
        case "filldown": return fillDown();
        case "fillright": return fillRight();
        case "az": return sortBy("asc");
        case "za": return sortBy("desc");
        case "filter": return toggleFilter();
        case "table": return toggleTable();
        case "freeze": return commit(() => { S.frozen = !S.frozen; }, S.frozen ? "Volets libérés" : "Ligne 1 figée ✓");
        case "showf": return commit(() => { S.showF = !S.showF; });
        case "fillmenu": case "cfmenu": case "pivotmenu": case "fnmenu": {
          const dd = t.closest(".xs-dd"), o = !dd.classList.contains("open"); closePops(dd); dd.classList.toggle("open", o);
          if (o && a === "cfmenu") openCF(dd); if (o && a === "pivotmenu") openPivot(dd);
          return;
        }
      }
    });
    host.addEventListener("change", e => { if (e.target.dataset.a === "fmt") styleSel("fmt", e.target.value); });
    document.addEventListener("click", docClick);
    function docClick(e) { if (!host.contains(e.target)) closePops(); if (!e.target.closest(".xs-fpop") && !e.target.closest(".xs-fbtn")) host.querySelectorAll(".xs-fpop").forEach(p => p.remove()); }

    // souris sur la grille
    let drag = null;
    grid.addEventListener("pointerdown", e => {
      const t = e.target.closest("td"), th = e.target.closest("th");
      if (e.target.closest(".xs-fbtn")) return;
      if (th && th.dataset.ch !== undefined) { if (edit) endEdit(true); const c = +th.dataset.ch; sel = { ac: c, ar: 0, c1: c, r1: 0, c2: c, r2: NR - 1, fc: c, fr: NR - 1 }; placeSel(); syncBars(); render(); wrap.focus({ preventScroll: true }); return; }
      if (th && th.dataset.rh !== undefined) { if (edit) endEdit(true); const r = +th.dataset.rh; sel = { ac: 0, ar: r, c1: 0, r1: r, c2: NC - 1, r2: r, fc: NC - 1, fr: r }; placeSel(); syncBars(); render(); wrap.focus({ preventScroll: true }); return; }
      if (!t) return;
      const c = +t.dataset.c, r = +t.dataset.r;
      if (edit) {
        const inp = activeInput();
        if (canPoint(inp)) { e.preventDefault(); drag = { mode: "point", c, r }; insertRef(refName(c, r)); return; }
        endEdit(true);
      }
      e.preventDefault();
      if (e.shiftKey) select(c, r, true); else select(c, r);
      drag = { mode: "sel" };
      wrap.focus({ preventScroll: true });
    });
    grid.addEventListener("pointerover", e => {
      if (!drag) return; const t = e.target.closest("td"); if (!t) return;
      const c = +t.dataset.c, r = +t.dataset.r;
      if (drag.mode === "sel") select(c, r, true);
      if (drag.mode === "point") insertRef(rangeName({ c1: Math.min(drag.c, c), r1: Math.min(drag.r, r), c2: Math.max(drag.c, c), r2: Math.max(drag.r, r) }));
    });
    grid.addEventListener("dblclick", e => { const t = e.target.closest("td"); if (!t || e.target.closest(".xs-fbtn")) return; const raw = fx.value; startEdit(raw, "edit"); });
    const onUp = () => { if (drag && drag.mode === "fill") finishFill(); drag = null; };
    addEventListener("pointerup", onUp);

    // poignée de recopie
    const handle = $(".xs-handle");
    handle.addEventListener("pointerdown", e => { e.preventDefault(); e.stopPropagation(); if (edit) endEdit(true); drag = { mode: "fill", src: selG(), dst: selG() }; host.classList.add("xs-filling"); });
    addEventListener("pointermove", onFillMove, { passive: true });
    function onFillMove(e) {
      if (!drag || drag.mode !== "fill") return;
      const el = document.elementFromPoint(e.clientX, e.clientY), t = el && el.closest && el.closest("td");
      if (!t || !grid.contains(t)) return;
      const c = +t.dataset.c, r = +t.dataset.r, s = drag.src;
      const dv = r > s.r2 ? r - s.r2 : r < s.r1 ? r - s.r1 : 0, dh = c > s.c2 ? c - s.c2 : c < s.c1 ? c - s.c1 : 0;
      const d = Math.abs(dv) >= Math.abs(dh) ? { c1: s.c1, c2: s.c2, r1: Math.min(s.r1, r), r2: Math.max(s.r2, r) } : { r1: s.r1, r2: s.r2, c1: Math.min(s.c1, c), c2: Math.max(s.c2, c) };
      drag.dst = d; sel.c1 = d.c1; sel.c2 = d.c2; sel.r1 = d.r1; sel.r2 = d.r2; placeSel();
    }
    function finishFill() {
      host.classList.remove("xs-filling");
      const { src, dst } = drag;
      if (dst.r1 === src.r1 && dst.r2 === src.r2 && dst.c1 === src.c1 && dst.c2 === src.c2) return;
      fillTo(src, dst);
    }

    // clavier sur la grille
    wrap.addEventListener("keydown", e => {
      if (e.target !== wrap) return;
      const mod = e.ctrlKey || e.metaKey, k = e.key;
      if (mod && (k === "z" || k === "Z")) { e.preventDefault(); return e.shiftKey ? redo() : undo(); }
      if (mod && (k === "y" || k === "Y")) { e.preventDefault(); return redo(); }
      if (mod && (k === "d" || k === "D")) { e.preventDefault(); return fillDown(); }
      if (mod && (k === "r" || k === "R")) { e.preventDefault(); return fillRight(); }
      if (mod && (k === "b" || k === "B" || k === "g" || k === "G")) { e.preventDefault(); return styleSel("b", !(cell(sel.ac, sel.ar) || {}).b); }
      if (mod && e.shiftKey && (k === "l" || k === "L")) { e.preventDefault(); return toggleFilter(); }
      if (mod && (k === "`" || k === "$" || k === "\"")) { e.preventDefault(); return commit(() => { S.showF = !S.showF; }); }
      if (mod && (k === "c" || k === "C" || k === "x" || k === "X")) {
        const g = selG(); clip = { g, cut: /x/i.test(k), cells: {} }; each(g, (c, r) => { clip.cells[key(c, r)] = cell(c, r) ? JSON.parse(JSON.stringify(cell(c, r))) : null; });
        const tsv = []; for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(display(getVal(c, r), (cell(c, r) || {}).fmt)); tsv.push(row.join("\t")); }
        try { navigator.clipboard && navigator.clipboard.writeText(tsv.join("\n")).catch(() => {}); } catch { /* */ }
        selbox.classList.add("xs-ants"); ctx.toast(clip.cut ? "Couper : sélectionne la destination puis Ctrl+V" : "Copié : sélectionne la destination puis Ctrl+V"); e.preventDefault(); return;
      }
      if (mod && (k === "v" || k === "V")) {
        if (!clip) return; e.preventDefault();
        const g = clip.g, dc = sel.ac - g.c1, dr = sel.ar - g.r1;
        commit(() => {
          const src = clip.cells;
          if (clip.cut) each(g, (c, r) => delete S.cells[key(c, r)]);
          each(g, (c, r) => { const x = src[key(c, r)], tc = c + dc, tr = r + dr; if (tc >= NC || tr >= NR) return; const k2 = key(tc, tr);
            if (!x) { delete S.cells[k2]; return; } const n = JSON.parse(JSON.stringify(x)); if (!clip.cut && typeof n.raw === "string" && n.raw[0] === "=") n.raw = shiftFormula(n.raw, dc, dr); S.cells[k2] = n; });
          if (clip.cut) clip = null;
        }, "Collé ✓");
        selbox.classList.remove("xs-ants"); return;
      }
      if (e.altKey && (k === "=" || k === "+")) { e.preventDefault(); return autoSum(); }
      if (e.shiftKey && k === "F3") { e.preventDefault(); const dd = $('[data-a="fnmenu"]').closest(".xs-dd"); $$(".xs-tabs button")[2].click(); dd.classList.add("open"); return; }
      const arrows = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] };
      if (arrows[k]) {
        e.preventDefault(); let [dc, dr] = arrows[k];
        if (mod) { // Ctrl+flèche : bord de la zone
          let c = e.shiftKey ? (sel.fc ?? sel.ac) : sel.ac, r = e.shiftKey ? (sel.fr ?? sel.ar) : sel.ar;
          const f0 = filled(c + dc, r + dr);
          while (c + dc >= 0 && c + dc < NC && r + dr >= 0 && r + dr < NR && (f0 ? filled(c + dc, r + dr) : !filled(c + dc, r + dr))) { c += dc; r += dr; }
          if (!f0 && c + dc >= 0 && c + dc < NC && r + dr >= 0 && r + dr < NR) { c += dc; r += dr; }
          return e.shiftKey ? select(c, r, true) : select(c, r);
        }
        return move(dc, dr, e.shiftKey);
      }
      if (k === "Enter") { e.preventDefault(); return move(0, e.shiftKey ? -1 : 1); }
      if (k === "Tab") { e.preventDefault(); return move(e.shiftKey ? -1 : 1, 0); }
      if (k === "Escape") { selbox.classList.remove("xs-ants"); clip = clip && clip.cut ? null : clip; return; }
      if (k === "F2") { e.preventDefault(); return startEdit(fx.value, "edit"); }
      if (k === "Delete" || k === "Backspace") { e.preventDefault(); return commit(() => each(selG(), (c, r) => { const x = S.cells[key(c, r)]; if (x) { x.raw = null; if (!x.b && !x.fill && !x.fmt && !x.align) delete S.cells[key(c, r)]; } })); }
      if (k === "Home") { e.preventDefault(); return select(0, mod ? 0 : sel.ar); }
      if (mod && k === "a") { e.preventDefault(); sel = { ac: sel.ac, ar: sel.ar, c1: 0, r1: 0, c2: NC - 1, r2: NR - 1 }; placeSel(); syncBars(); render(); return; }
      if (k.length === 1 && !mod && !e.altKey) { e.preventDefault(); startEdit(k, "enter"); }
    });
    function editKeys(e) {
      const inp = e.target, k = e.key;
      if (!acBox.hidden) {
        const items = [...acBox.querySelectorAll("[data-ac]")], i = items.findIndex(b => b.classList.contains("on"));
        if (k === "ArrowDown" || k === "ArrowUp") { e.preventDefault(); items[i].classList.remove("on"); items[(i + (k === "ArrowDown" ? 1 : items.length - 1)) % items.length].classList.add("on"); return; }
        if (k === "Tab") { e.preventDefault(); return acceptAC(items[i].dataset.ac); }
        if (k === "Escape") { e.preventDefault(); acBox.hidden = true; return; }
      }
      if (k === "F4") { e.preventDefault(); return f4(); }
      if (k === "Enter") { e.preventDefault(); return endEdit(true, 0, e.shiftKey ? -1 : 1); }
      if (k === "Tab") { e.preventDefault(); return endEdit(true, e.shiftKey ? -1 : 1, 0); }
      if (k === "Escape") { e.preventDefault(); return endEdit(false); }
      if (inp === editor && edit && edit.mode === "enter" && k.startsWith("Arrow") && !e.shiftKey) {
        if (canPoint(inp)) { // mode Pointer : les flèches déplacent la référence insérée
          e.preventDefault();
          const base = pointIns ? parseRef(inp.value.slice(pointIns.s, pointIns.e).split(":")[0]) : { c: edit.c, r: edit.r };
          const d = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[k];
          const c = Math.max(0, Math.min(NC - 1, base.c + d[0])), r = Math.max(0, Math.min(NR - 1, base.r + d[1]));
          return insertRef(refName(c, r));
        }
        e.preventDefault(); const d = { ArrowUp: [0, -1], ArrowDown: [0, 1], ArrowLeft: [-1, 0], ArrowRight: [1, 0] }[k];
        return endEdit(true, d[0], d[1]);
      }
      if (/^[a-z0-9]$/i.test(k) || k === "Backspace") pointIns = null;
    }
    editor.addEventListener("keydown", editKeys);
    fx.addEventListener("keydown", editKeys);
    editor.addEventListener("input", () => { onEditInput(); });
    fx.addEventListener("focus", () => { if (!edit) { edit = { c: sel.ac, r: sel.ar, mode: "edit", via: "fx" }; statusMsg.textContent = "Modifier"; onEditInput(); } else edit.via = "fx"; });
    fx.addEventListener("input", () => { if (!edit) edit = { c: sel.ac, r: sel.ar, mode: "edit", via: "fx" }; edit.via = "fx"; if (editor.style.display !== "block") { /* édition dans la barre */ } onEditInput(); });
    editor.addEventListener("focus", () => { if (edit) edit.via = "cell"; });
    editor.addEventListener("blur", () => setTimeout(() => { if (edit && edit.via === "cell" && document.activeElement !== editor && document.activeElement !== fx && !host.contains(document.activeElement)) endEdit(true); }, 120));
    fx.addEventListener("blur", () => setTimeout(() => { if (edit && edit.via === "fx" && document.activeElement !== fx && !host.contains(document.activeElement)) endEdit(true); }, 120));
    nameBox.addEventListener("keydown", e => {
      if (e.key !== "Enter") return; e.preventDefault();
      const g = parseRange(nameBox.value.replace(/\$/g, ""));
      if (!g || g.c2 >= NC || g.r2 >= NR) { ctx.toast("Référence invalide"); return; }
      sel = { ac: g.c1, ar: g.r1, ...g, fc: g.c2, fr: g.r2 }; render(); wrap.focus();
    });
    const onResize = () => placeSel();
    addEventListener("resize", onResize);
    wrap.addEventListener("scroll", () => { if (edit) highlightRefs(activeInput().value); }, { passive: true });

    recalc(); render();
    requestAnimationFrame(placeSel);

    /* ---------- API des check ---------- */
    const up = r => String(r).toUpperCase().replace(/\$/g, "");
    const api = {
      val: ref => { const v = valAt(up(ref)); return isErr(v) ? v.c : v === null || v === undefined ? "" : v; },
      num: ref => { const v = api.val(ref); if (typeof v === "number") return v; if (typeof v === "boolean") return NaN; const n = v === "" ? null : parseNumText(v); return n === null ? NaN : n; },
      f: ref => { const p = parseRef(up(ref)); const raw = p ? rawOf(p.c, p.r) : null; if (typeof raw !== "string" || raw[0] !== "=") return "";
        return raw.replace(/("(?:[^"]|"")*")|\s+/g, (m, s) => s || "").replace(/("(?:[^"]|"")*")|[^"]+/g, (m, s) => s || m.toUpperCase()); },
      isFormula: ref => api.f(ref) !== "",
      hasFn: (ref, name) => { const f = api.f(ref); if (!f) return false; const n = canon(String(name)); const names = [n, ...Object.keys(ALIAS).filter(a => ALIAS[a] === n)];
        return names.some(x => new RegExp("(^|[^A-Z0-9_.])" + x.replace(/\./g, "\\.") + "\\(").test(f.replace(/"(?:[^"]|"")*"/g, '""'))); },
      hasAbs: ref => /\$[A-Z]{0,2}\$?\d|\$[A-Z]{1,2}\d/.test(api.f(ref).replace(/"(?:[^"]|"")*"/g, "")),
      range: rng => { const g = parseRange(up(rng)); if (!g) return []; const o = []; each(g, (c, r) => o.push(api.val(refName(c, r)))); return o; },
      style: ref => { const x = S.cells[up(ref)] || {}; return { bold: !!x.b, fill: x.fill || "", fmt: x.fmt || "", align: x.align || "" }; },
      near: (a, b, eps = 0.01) => typeof a === "number" && typeof b === "number" && Math.abs(a - b) <= eps,
      sortedBy: (col, dir = "asc") => {
        const c = colIdx(String(col).replace(/\d+/g, "")); const g = region(c, 0) || region(c, 1); if (!g) return false;
        const vs = []; for (let r = g.r1 + 1; r <= g.r2; r++) { const v = getVal(c, r); if (v !== null && v !== "" && !isErr(v)) vs.push(v); }
        if (vs.length < 2) return false;
        for (let i = 1; i < vs.length; i++) { const d = compare(vs[i - 1], vs[i]); if (dir === "desc" ? d < 0 : d > 0) return false; }
        return true;
      },
      get cf() { return S.cf.map(r => ({ range: r.range, op: r.op, v1: r.v1, v2: r.v2, color: r.color })); },
      get charts() { return S.charts.map(c => ({ type: c.type, range: c.range })); },
      get pivots() { return S.pivots.map(p => ({ rows: p.rows, values: p.values, agg: p.agg })); },
      get filter() { return !!S.filter; },
      get filtered() { return S.filter ? Object.values(S.filter.excl || {}).some(x => x.length) : false; },
      get table() { return !!S.table; },
      get frozen() { return S.frozen; },
      cell: ref => { const x = S.cells[up(ref)] || {}; return { raw: x.raw ?? "", val: api.val(ref), style: api.style(ref) }; },
    };
    return {
      api,
      reset() { hist = []; fut = []; S = fresh(); sel = { ac: 0, ar: 0, c1: 0, r1: 0, c2: 0, r2: 0 }; recalc(); objSig = ""; render(); ctx.changed(); },
      destroy() { removeEventListener("pointerup", onUp); removeEventListener("pointermove", onFillMove); removeEventListener("resize", onResize); document.removeEventListener("click", docClick); },
      exportFile(initial) {
        const st = initial ? fresh() : S, keep = S; S = st; recalc();
        let maxC = 0, maxR = 0; Object.keys(S.cells).forEach(k => { const p = parseRef(k); if (S.cells[k].raw !== null) { maxC = Math.max(maxC, p.c); maxR = Math.max(maxR, p.r); } });
        const lines = [];
        for (let r = 0; r <= maxR; r++) { const row = []; for (let c = 0; c <= maxC; c++) { const v = getVal(c, r); let s = typeof v === "number" ? numFr(v) : display(v, ""); if (/[;"\n]/.test(s)) s = '"' + s.replace(/"/g, '""') + '"'; row.push(s); } lines.push(row.join(";")); }
        S = keep; recalc();
        return { name: (setup.sheetName || "exercice").replace(/[^\w-]+/g, "_") + ".csv", type: "text/csv;charset=utf-8", content: "\ufeff" + lines.join("\r\n") };
      },
      focus() { wrap.focus({ preventScroll: true }); },
      /* accès de test (QA automatisée) : saisie comme au clavier + sélection */
      test: {
        set(ref, text) { const p = parseRef(up(ref)); commit(() => setRaw(p.c, p.r, text)); },
        select(rng) { const g = parseRange(up(rng)); sel = { ac: g.c1, ar: g.r1, ...g, fc: g.c2, fr: g.r2 }; render(); },
      },
    };
  }
  /* exposé pour les tests / autres modules */
  const engine = { parse, shiftFormula, cycleAbs, display, parseInput, FN_NAMES };
  if (typeof window !== "undefined" && window.OFFICE) { window.OFFICE.sim("sheet", mount); window.OFFICE.sheetEngine = engine; }
  if (typeof module !== "undefined") module.exports = { mount, engine };
})();
