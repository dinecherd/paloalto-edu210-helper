/* Lab Double barrière — moteur : état, validation commit/push, console Panorama simulée, missions,
   packet tracer, terminal CLI et quiz. Données : lab-data.js (window.LAB_DATA). Aucune dépendance. */
(function () {
  "use strict";
  const { LEVELS, EX, FLOWS, CLI, QUIZ } = window.LAB_DATA;
  const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const clone = o => JSON.parse(JSON.stringify(o));
  const J = o => JSON.stringify(o);
  const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const sleep = ms => new Promise(r => setTimeout(r, reduce ? Math.min(ms, 60) : ms));
  const toast = m => window.PAN && PAN.toast ? PAN.toast(m) : console.log(m);
  const confetti = el => window.PAN && PAN.confetti && PAN.confetti(el);
  const KEY = "panw-lab-double-barriere-v1";

  /* =========================================================
     RÉFÉRENTIEL DU LAB
     ========================================================= */
  const DG = {
    "DG-Perimetre": { fw: "FW-PERIM", node: "perim", model: "PA-3410", serial: "021201004321", ip: "192.168.100.11", stack: "TS-Perimetre",
      zones: ["Untrust", "DMZ", "Transit"],
      ifaces: { "ethernet1/1": ["Untrust", "203.0.113.2/24"], "ethernet1/2": ["DMZ", "172.16.10.1/24"], "ethernet1/3": ["Transit", "10.255.0.1/30"] } },
    "DG-Interne": { fw: "FW-INT", node: "int", model: "PA-1410", serial: "013201007654", ip: "192.168.100.12", stack: "TS-Interne",
      zones: ["Transit", "Users", "Servers", "Admin"],
      ifaces: { "ethernet1/1": ["Transit", "10.255.0.2/30"], "ethernet1/2": ["Users", "10.10.10.1/24"], "ethernet1/3": ["Servers", "10.20.20.1/24"], "ethernet1/4": ["Admin", "10.30.30.1/24"] } },
  };
  const DGS = Object.keys(DG);
  const LOCS = ["Shared", ...DGS];
  const ALL_ZONES = ["Untrust", "DMZ", "Transit", "Users", "Servers", "Admin"];
  const ALL_IFACES = ["ethernet1/1", "ethernet1/2", "ethernet1/3", "ethernet1/4"];
  const APPS = { "web-browsing": ["tcp", 80], ssl: ["tcp", 443], dns: ["udp", 53], ntp: ["udp", 123], "ms-ds-smb": ["tcp", 445], ldap: ["tcp", 389], kerberos: ["tcp", 88], "ms-rdp": ["tcp", 3389], ssh: ["tcp", 22], smtp: ["tcp", 25], snmp: ["udp", 161], ftp: ["tcp", 21], ping: ["icmp", 0] };
  const PRE_SVC = { "service-http": ["tcp", [80, 8080]], "service-https": ["tcp", [443]] };
  const EDL = { "panw-known-ip-list": ["198.51.100.66"], "panw-highrisk-ip-list": ["198.51.100.77"], "panw-bulletproof-ip-list": ["198.51.100.88"] };
  const PROFILES = { "PG-BP-Strict": "AV default · AS strict · Vuln strict · URL BP · File Blocking BP · WildFire default" };
  const LOGFWD = { "LF-Panorama": "traffic, threat, url, wildfire → Panorama + syslog SIEM" };
  const COLORS = { color1: ["Red", "#EF4444"], color2: ["Green", "#22C55E"], color3: ["Blue", "#3B82F6"], color4: ["Yellow", "#EAB308"], color5: ["Copper", "#B87333"], color6: ["Orange", "#F97316"], color7: ["Purple", "#A855F7"], color8: ["Gray", "#94A3B8"] };
  const KL = { tag: "Tag", addr: "Adresse", grp: "Groupe", svc: "Service" };
  const ZCLS = { Untrust: "z-un", DMZ: "z-dmz", Transit: "z-tr", Users: "z-us", Servers: "z-srv", Admin: "z-adm" };
  const NAME_RE = /^[A-Za-z0-9][\w .-]{0,62}$/;

  /* ---------- IP ---------- */
  const validIp = s => /^\d{1,3}(\.\d{1,3}){3}$/.test(s) && s.split(".").every(n => +n <= 255);
  const ip2i = s => s.split(".").reduce((a, b) => a * 256 + +b, 0);
  const validNet = v => { const [a, m] = String(v).split("/"); return validIp(a) && (m === undefined || (/^\d+$/.test(m) && +m <= 32)); };
  const normNet = v => (String(v).includes("/") ? String(v) : v + "/32");
  const netRange = v => { const [a, m = "32"] = String(v).split("/"); const size = 2 ** (32 - +m), lo = Math.floor(ip2i(a) / size) * size; return { lo, hi: lo + size - 1 }; };
  const inNet = (ip, net) => { const r = netRange(net), x = ip2i(ip); return x >= r.lo && x <= r.hi; };
  const isPrivate = ip => ["10.0.0.0/8", "172.16.0.0/12", "192.168.0.0/16"].some(n => inNet(ip, n));
  const validPorts = p => /^\d+(-\d+)?(,\d+(-\d+)?)*$/.test(p) && p.split(/[,-]/).every(n => +n >= 1 && +n <= 65535);
  const portIn = (port, spec) => String(spec).split(",").some(s => { const [a, b] = s.split("-").map(Number); return b ? port >= a && port <= b : port === a; });

  /* =========================================================
     ÉTAT (persisté dans localStorage)
     ========================================================= */
  const emptyLoc = dg => Object.assign({ addr: [], grp: [], svc: [], tag: [] }, dg ? { pre: [], post: [], nat: [] } : {});
  const emptyCfg = () => ({ Shared: emptyLoc(false), "DG-Perimetre": emptyLoc(true), "DG-Interne": emptyLoc(true) });
  const slice = (cfg, dg) => clone({ Shared: cfg.Shared, [dg]: cfg[dg] });
  function fresh() {
    const cfg = emptyCfg();
    return { v: 1, cfg, committed: clone(cfg), pushed: { "DG-Perimetre": slice(cfg, "DG-Perimetre"), "DG-Interne": slice(cfg, "DG-Interne") },
      lastPush: {}, jobs: [], job: 1240, done: {}, assisted: {}, setups: {}, traces: {}, cli: {}, quiz: { best: 0 }, cur: EX[0].id,
      ui: { tab: "objects", node: "tag", ctx: "Shared", sel: null } };
  }
  let st;
  try { st = JSON.parse(localStorage.getItem(KEY)); } catch { st = null; }
  if (!st || st.v !== 1) st = fresh();
  const save = () => { try { localStorage.setItem(KEY, J(st)); } catch { /* quota / mode privé */ } };
  const pending = () => J(st.cfg) !== J(st.committed);
  const inSync = dg => J(slice(st.committed, dg)) === J(st.pushed[dg]);

  /* =========================================================
     RÉSOLUTION D'OBJETS & ÉVALUATION DE POLITIQUE
     ========================================================= */
  const scope = ctx => (ctx === "Shared" ? ["Shared"] : [ctx, "Shared"]);
  const lookup = (cfg, ctx, k, name) => { for (const l of scope(ctx)) { const o = cfg[l] && cfg[l][k].find(x => x.name === name); if (o) return o; } return null; };
  const visible = (cfg, ctx, k) => (ctx === "Shared" ? ["Shared"] : ["Shared", ctx]).flatMap(l => cfg[l][k].map(o => ({ o, loc: l })));
  const addrExists = (cfg, ctx, n) => n === "any" || !!EDL[n] || !!lookup(cfg, ctx, "addr", n) || !!lookup(cfg, ctx, "grp", n);
  const svcExists = (cfg, ctx, n) => n === "any" || n === "application-default" || !!PRE_SVC[n] || !!lookup(cfg, ctx, "svc", n);

  function ranges(cfg, ctx, name, depth = 0) {
    if (EDL[name]) return EDL[name].map(i => ({ lo: ip2i(i), hi: ip2i(i) }));
    const a = lookup(cfg, ctx, "addr", name);
    if (a) {
      if (a.type === "ip-netmask") return [netRange(a.value)];
      if (a.type === "ip-range") { const [x, y] = a.value.split("-"); return [{ lo: ip2i(x), hi: ip2i(y) }]; }
      return [];
    }
    const g = lookup(cfg, ctx, "grp", name);
    return g && depth < 6 ? g.members.flatMap(m => ranges(cfg, ctx, m, depth + 1)) : [];
  }
  const addrMatch = (cfg, ctx, list, ip) => list.includes("any") || list.some(n => ranges(cfg, ctx, n).some(r => ip2i(ip) >= r.lo && ip2i(ip) <= r.hi));
  const zoneMatch = (list, z) => list.includes("any") || list.includes(z);
  function svcMatch(cfg, ctx, list, f) {
    if (list.includes("any")) return true;
    if (list.includes("application-default")) { const d = APPS[f.app]; if (d && d[0] === f.proto && (d[0] === "icmp" || d[1] === f.port)) return true; }
    return list.some(n => {
      if (PRE_SVC[n]) return PRE_SVC[n][0] === f.proto && PRE_SVC[n][1].includes(f.port);
      const s = lookup(cfg, ctx, "svc", n); return s && s.proto === f.proto && portIn(f.port, s.port);
    });
  }
  const ruleMatch = (cfg, dg, r, f) => !r.disabled && zoneMatch(r.from, f.from) && zoneMatch(r.to, f.to) && addrMatch(cfg, dg, r.src, f.src) && addrMatch(cfg, dg, r.dst, f.dst)
    && (r.app.includes("any") || (f.app && r.app.includes(f.app))) && svcMatch(cfg, dg, r.svc, f);

  function evalPolicy(cfg, dg, f) {
    const list = [...cfg[dg].pre.map(r => ({ r, base: "pre" })), ...cfg[dg].post.map(r => ({ r, base: "post" }))];
    for (let i = 0; i < list.length; i++) if (ruleMatch(cfg, dg, list[i].r, f)) return { name: list[i].r.name, action: list[i].r.action, base: list[i].base, r: list[i].r, index: i + 1 };
    return f.from === f.to ? { name: "intrazone-default", action: "allow", base: "default", index: list.length + 1 } : { name: "interzone-default", action: "deny", base: "default", index: list.length + 2 };
  }
  function evalNat(cfg, dg, f) {
    return cfg[dg].nat.find(n => !n.disabled && zoneMatch(n.from, f.from) && (n.to === "any" || n.to === f.to) && addrMatch(cfg, dg, n.src, f.src) && addrMatch(cfg, dg, n.dst, f.dst)
      && (n.svc === "any" || svcMatch(cfg, dg, [n.svc], f))) || null;
  }
  function zoneOf(dg, ip) {
    for (const [z, cidr] of Object.values(DG[dg].ifaces)) if (inNet(ip, cidr)) return z;
    if (dg === "DG-Interne") return "Transit";
    return inNet(ip, "10.0.0.0/8") ? "Transit" : "Untrust";
  }

  /* ---------- shadowing ---------- */
  const setCovers = (a, b) => a.includes("any") || (!b.includes("any") && b.every(x => a.includes(x)));
  function addrCovers(cfg, dg, a, b) {
    if (a.includes("any")) return true;
    if (b.includes("any")) return false;
    const ra = a.flatMap(n => ranges(cfg, dg, n)), rb = b.flatMap(n => ranges(cfg, dg, n));
    return rb.length > 0 && rb.every(x => ra.some(y => x.lo >= y.lo && x.hi <= y.hi));
  }
  function shadows(cfg, dg) {
    const list = [...cfg[dg].pre, ...cfg[dg].post].filter(r => !r.disabled), out = [];
    for (let j = 1; j < list.length; j++) for (let i = 0; i < j; i++) {
      const A = list[i], B = list[j];
      if (setCovers(A.from, B.from) && setCovers(A.to, B.to) && addrCovers(cfg, dg, A.src, B.src) && addrCovers(cfg, dg, A.dst, B.dst)
        && setCovers(A.app, B.app) && (A.svc.includes("any") || J([...A.svc].sort()) === J([...B.svc].sort()))) { out.push([A.name, B.name]); break; }
    }
    return out;
  }

  /* =========================================================
     VALIDATION (commit Panorama / push pare-feu)
     ========================================================= */
  const pathOf = loc => (loc === "Shared" ? "shared" : `device-group -> ${loc}`);
  function validatePano(cfg) {
    const errors = [], warnings = [];
    for (const loc of LOCS) {
      cfg[loc].grp.forEach(g => g.members.forEach(m => { if (!addrExists(cfg, loc, m) || m === "any") errors.push(`${pathOf(loc)} -> address-group -> ${g.name} -> static '${m}' is not a valid reference`); }));
      cfg[loc].addr.forEach(a => (a.tags || []).forEach(t => { if (!lookup(cfg, loc, "tag", t)) errors.push(`${pathOf(loc)} -> address -> ${a.name} -> tag '${t}' is not a valid reference`); }));
    }
    for (const dg of DGS) {
      const seen = new Set();
      for (const base of ["pre", "post"]) for (const r of cfg[dg][base]) {
        const p = `device-group -> ${dg} -> ${base}-rulebase -> security -> rules -> ${r.name}`;
        if (seen.has(r.name)) errors.push(`${p} : rule name is not unique within the device group`);
        seen.add(r.name);
        r.src.forEach(n => addrExists(cfg, dg, n) || errors.push(`${p} -> source '${n}' is not a valid reference`));
        r.dst.forEach(n => addrExists(cfg, dg, n) || errors.push(`${p} -> destination '${n}' is not a valid reference`));
        r.svc.forEach(n => svcExists(cfg, dg, n) || errors.push(`${p} -> service '${n}' is not a valid reference`));
        r.tags.forEach(t => lookup(cfg, dg, "tag", t) || errors.push(`${p} -> tag '${t}' is not a valid reference`));
        if (!r.disabled && r.action === "allow" && r.app.includes("any") && r.svc.includes("any")) warnings.push(`${dg} : rule '${r.name}' allows any application on any service`);
      }
      for (const r of cfg[dg].nat) {
        const p = `device-group -> ${dg} -> pre-rulebase -> nat -> rules -> ${r.name}`;
        [...r.src, ...r.dst].forEach(n => addrExists(cfg, dg, n) || errors.push(`${p} -> '${n}' is not a valid reference`));
        if (r.dnat && !lookup(cfg, dg, "addr", r.dnat)) errors.push(`${p} -> destination-translation -> translated-address '${r.dnat}' is not a valid reference`);
        if (!svcExists(cfg, dg, r.svc) || r.svc === "application-default") errors.push(`${p} -> service '${r.svc}' is not a valid reference`);
      }
      shadows(cfg, dg).forEach(([a, b]) => warnings.push(`${dg} : Rule '${a}' shadows rule '${b}'`));
    }
    return { errors, warnings };
  }
  function validatePush(cfg, dg) {
    const errors = [], Z = DG[dg].zones;
    for (const base of ["pre", "post"]) for (const r of cfg[dg][base]) {
      r.from.forEach(z => z === "any" || Z.includes(z) || errors.push(`rulebase -> security -> rules -> ${r.name} -> from '${z}' is not a valid reference`));
      r.to.forEach(z => z === "any" || Z.includes(z) || errors.push(`rulebase -> security -> rules -> ${r.name} -> to '${z}' is not a valid reference`));
    }
    for (const r of cfg[dg].nat) {
      r.from.forEach(z => z === "any" || Z.includes(z) || errors.push(`rulebase -> nat -> rules -> ${r.name} -> from '${z}' is not a valid reference`));
      if (r.to !== "any" && !Z.includes(r.to)) errors.push(`rulebase -> nat -> rules -> ${r.name} -> to '${r.to}' is not a valid reference`);
      if (r.snat.type === "dipp" && !DG[dg].ifaces[r.snat.iface]) errors.push(`rulebase -> nat -> rules -> ${r.name} -> source-translation -> interface '${r.snat.iface}' is not a valid reference`);
    }
    return errors;
  }

  /* ---------- commit / push (logique pure, réutilisée par la GUI et la CLI) ---------- */
  const now = () => new Date().toLocaleTimeString("fr-FR");
  function addJob(j) { st.jobs.unshift(Object.assign({ t: now() }, j)); st.jobs = st.jobs.slice(0, 40); }
  function doCommit() {
    const id = ++st.job, v = validatePano(st.cfg);
    if (v.errors.length) { addJob({ id, type: "Commit", dev: "Panorama", ok: false, msg: v.errors[0] }); return { id, ok: false, ...v }; }
    st.committed = clone(st.cfg);
    addJob({ id, type: "Commit", dev: "Panorama", ok: true, msg: v.warnings.length ? `${v.warnings.length} warning(s)` : "" });
    return { id, ok: true, ...v };
  }
  function doPush(dg) {
    const id = ++st.job, errors = validatePush(st.committed, dg), warnings = shadows(st.committed, dg).map(([a, b]) => `Rule '${a}' shadows rule '${b}'`);
    const ok = !errors.length;
    if (ok) st.pushed[dg] = slice(st.committed, dg);
    st.lastPush[dg] = { ok, id, t: Date.now(), at: now() };
    addJob({ id, type: "CommitAll", dev: `${DG[dg].fw} (${dg})`, ok, msg: errors[0] || "" });
    return { id, ok, errors, warnings };
  }

  /* =========================================================
     SOLUTIONS → CHECKLIST, CLI, ÉTAPES GUI
     ========================================================= */
  const eqSet = (a, b) => J([...a].sort()) === J([...b].sort());
  const fmt = v => (Array.isArray(v) ? v.join(", ") : String(v || "none"));
  const baseLabel = b => ({ pre: "Pre Rules", post: "Post Rules", nat: "NAT" }[b]);
  const q = s => (/[\s"]/.test(s) || s === "" ? `"${String(s).replace(/"/g, "'")}"` : s);
  const LL = v => (Array.isArray(v) ? (v.length === 1 ? q(v[0]) : `[ ${v.map(q).join(" ")} ]`) : q(v));
  const locCli = loc => (loc === "Shared" ? "shared" : `device-group ${loc}`);

  function itemCli(it) {
    const o = it.o;
    if (it.k === "tag") return [`set ${locCli(it.loc)} tag ${q(o.name)} color ${o.color}`];
    if (it.k === "addr") return [`set ${locCli(it.loc)} address ${q(o.name)} ${o.type} ${o.value}${o.desc ? ` description ${q(o.desc)}` : ""}${o.tags && o.tags.length ? ` tag ${LL(o.tags)}` : ""}`];
    if (it.k === "grp") return [`set ${locCli(it.loc)} address-group ${q(o.name)} static ${LL(o.members)}${o.desc ? ` description ${q(o.desc)}` : ""}`];
    if (it.k === "svc") return [`set ${locCli(it.loc)} service ${q(o.name)} protocol ${o.proto} port ${o.port}${o.desc ? ` description ${q(o.desc)}` : ""}`];
    if (it.k === "rule") {
      let s = `set device-group ${it.dg} ${it.base}-rulebase security rules ${q(o.name)} from ${LL(o.from)} to ${LL(o.to)} source ${LL(o.src)} destination ${LL(o.dst)} application ${LL(o.app)} service ${LL(o.svc)} action ${o.action}`;
      if (o.profile) s += ` profile-setting group ${q(o.profile)}`;
      if (o.lf) s += ` log-setting ${q(o.lf)}`;
      if (o.tags.length) s += ` tag ${LL(o.tags)}`;
      if (o.desc) s += ` description ${q(o.desc)}`;
      if (o.disabled) s += " disabled yes";
      return it.pos ? [s, `move device-group ${it.dg} ${it.base}-rulebase security rules ${q(o.name)} ${it.pos}`] : [s];
    }
    let s = `set device-group ${it.dg} pre-rulebase nat rules ${q(o.name)} from ${LL(o.from)} to ${o.to} source ${LL(o.src)} destination ${LL(o.dst)} service ${o.svc}`;
    if (o.snat.type === "dipp") s += ` source-translation dynamic-ip-and-port interface-address interface ${o.snat.iface}`;
    if (o.dnat) s += ` destination-translation translated-address ${q(o.dnat)}`;
    if (o.desc) s += ` description ${q(o.desc)}`;
    return [s];
  }
  function cfgToCli(cfg) {
    const out = [];
    for (const loc of LOCS) {
      for (const k of ["tag", "addr", "grp", "svc"]) cfg[loc][k].forEach(o => out.push(...itemCli({ k, loc, o })));
      if (loc !== "Shared") {
        for (const base of ["pre", "post"]) cfg[loc][base].forEach(o => out.push(...itemCli({ k: "rule", dg: loc, base, o })));
        cfg[loc].nat.forEach(o => out.push(...itemCli({ k: "nat", dg: loc, o })));
      }
    }
    return out;
  }
  const c = s => `<code>${esc(s)}</code>`;
  function itemGui(it) {
    const o = it.o;
    if (it.k === "tag") return `<b>Objects › Tags › Add</b> · Device Group <i>${it.loc}</i> · Name ${c(o.name)} · Color ${COLORS[o.color][0]}`;
    if (it.k === "addr") return `<b>Objects › Addresses › Add</b> · Device Group <i>${it.loc}</i> · Name ${c(o.name)} · Type IP Netmask · ${c(o.value)}${o.tags.length ? ` · Tags ${c(o.tags.join(", "))}` : ""}`;
    if (it.k === "grp") return `<b>Objects › Address Groups › Add</b> · Device Group <i>${it.loc}</i> · Name ${c(o.name)} · Type Static · membres ${c(o.members.join(", "))}`;
    if (it.k === "svc") return `<b>Objects › Services › Add</b> · Device Group <i>${it.loc}</i> · Name ${c(o.name)} · Protocol ${o.proto.toUpperCase()} · Destination Port ${c(o.port)}`;
    if (it.k === "rule") return `<b>Policies › Security › ${baseLabel(it.base)} › Add</b> (Device Group <i>${it.dg}</i>)<ul>
      <li><b>General</b> : Name ${c(o.name)}${o.tags.length ? ` · Tags ${c(o.tags.join(", "))}` : ""}</li>
      <li><b>Source</b> : Zone ${c(fmt(o.from))} · Address ${c(fmt(o.src))}</li>
      <li><b>Destination</b> : Zone ${c(fmt(o.to))} · Address ${c(fmt(o.dst))}</li>
      <li><b>Application</b> : ${c(fmt(o.app))}</li><li><b>Service/URL Category</b> : ${c(fmt(o.svc))}</li>
      <li><b>Actions</b> : Action ${c(o.action)}${o.profile ? ` · Profile Type Group ${c(o.profile)}` : ""} · Log at Session End${o.lf ? ` · Log Forwarding ${c(o.lf)}` : ""}</li>
      ${it.pos ? `<li><b>Move › Move ${it.pos === "top" ? "Top" : "Bottom"}</b></li>` : ""}</ul>`;
    return `<b>Policies › NAT › Pre Rules › Add</b> (Device Group <i>${it.dg}</i>)<ul>
      <li><b>General</b> : Name ${c(o.name)}</li>
      <li><b>Original Packet</b> : Source Zone ${c(fmt(o.from))} · Destination Zone ${c(o.to)} · Source ${c(fmt(o.src))} · Destination ${c(fmt(o.dst))} · Service ${c(o.svc)}</li>
      <li><b>Translated Packet</b> : ${o.snat.type === "dipp" ? `Source Translation Dynamic IP And Port › Interface Address ${c(o.snat.iface)}` : "Source Translation None"}${o.dnat ? ` · Destination Translation Static IP › ${c(o.dnat)}` : ""}</li></ul>`;
  }

  function objDiff(k, cur, o) {
    if (k === "tag" && cur.color !== o.color) return `Couleur attendue ${COLORS[o.color][0]} (actuelle ${COLORS[cur.color] ? COLORS[cur.color][0] : cur.color}).`;
    if (k === "addr") {
      if (cur.type !== o.type) return `Type attendu IP Netmask (actuel ${cur.type}).`;
      if (normNet(cur.value) !== normNet(o.value)) return `Valeur attendue ${o.value} (actuelle ${cur.value}).`;
      const miss = (o.tags || []).filter(t => !(cur.tags || []).includes(t));
      if (miss.length) return `Tag manquant : ${miss.join(", ")}.`;
    }
    if (k === "grp" && !eqSet(cur.members, o.members)) return `Membres attendus : ${o.members.join(", ")}.`;
    if (k === "svc" && (cur.proto !== o.proto || cur.port !== o.port)) return `Attendu ${o.proto.toUpperCase()}/${o.port} (actuel ${cur.proto.toUpperCase()}/${cur.port}).`;
    return "";
  }
  function findRuleAnywhere(name) {
    for (const dg of DGS) for (const b of ["pre", "post", "nat"]) if (st.cfg[dg][b].some(r => r.name === name)) return `${dg} › ${baseLabel(b)}`;
    return "";
  }
  function critItem(it) {
    const o = it.o;
    if (KL[it.k]) {
      const cur = st.cfg[it.loc][it.k].find(x => x.name === o.name);
      const other = !cur && LOCS.find(l => st.cfg[l][it.k].some(x => x.name === o.name));
      const d = cur ? objDiff(it.k, cur, o) : "";
      return [{ ok: !!cur && !d, label: `${KL[it.k]} ${o.name} · ${it.loc}`, hint: cur ? d : other ? `Existe dans ${other} : mauvais emplacement (attendu ${it.loc}).` : "Pas encore créé." }];
    }
    const list = st.cfg[it.dg][it.k === "nat" ? "nat" : it.base], r = list.find(x => x.name === o.name);
    const title = `${it.k === "nat" ? "Règle NAT" : "Règle"} ${o.name} · ${it.dg} › ${baseLabel(it.k === "nat" ? "nat" : it.base)}`;
    if (!r) { const w = findRuleAnywhere(o.name); return [{ ok: false, label: title, hint: w ? `Trouvée dans ${w} : mauvais emplacement.` : "Pas encore créée (nom exact, sensible à la casse)." }]; }
    const out = [{ ok: true, label: title }];
    const F = it.k === "nat"
      ? [["from", "Source Zone"], ["to", "Destination Zone"], ["src", "Source Address"], ["dst", "Destination Address"], ["svc", "Service"]]
      : [["from", "Source Zone"], ["src", "Source Address"], ["to", "Destination Zone"], ["dst", "Destination Address"], ["app", "Application"], ["svc", "Service"], ["action", "Action"]];
    if (it.k === "rule") { if (o.profile) F.push(["profile", "Profile Group"]); if (o.lf) F.push(["lf", "Log Forwarding"]); }
    for (const [f, lab] of F) {
      const ok = Array.isArray(o[f]) ? eqSet(r[f], o[f]) : r[f] === o[f];
      let hint = ok ? "" : `Actuellement : ${fmt(r[f])}.`;
      if (!ok && f === "dst" && r.dst.includes("H-WEB-DMZ") && o.dst.includes("H-WEB-PUB")) hint += " Rappel : adresse de destination pré-NAT (publique).";
      out.push({ ok, label: `${lab} = ${fmt(o[f])}`, hint, sub: true });
    }
    if (it.k === "rule" && o.tags.length) { const miss = o.tags.filter(t => !r.tags.includes(t)); out.push({ ok: !miss.length, label: `Tag ${o.tags.join(", ")}`, hint: miss.length ? "Tag absent de la règle." : "", sub: true }); }
    if (it.k === "nat") {
      out.push({ ok: J(r.snat) === J(o.snat), label: o.snat.type === "dipp" ? `SNAT Dynamic IP and Port · ${o.snat.iface}` : "Pas de NAT source", hint: `Actuellement : ${r.snat.type === "dipp" ? "DIPP " + r.snat.iface : "none"}.`, sub: true });
      out.push({ ok: r.dnat === o.dnat, label: o.dnat ? `DNAT vers ${o.dnat}` : "Pas de NAT destination", hint: `Actuellement : ${r.dnat || "none"}.`, sub: true });
    }
    out.push({ ok: !r.disabled, label: "Règle activée", hint: "La règle est désactivée.", sub: true });
    if (it.pos === "top") out.push({ ok: list[0] === r, label: "En première position", hint: `Position actuelle : ${list.indexOf(r) + 1}. Utilise Move › Top.`, sub: true });
    if (it.pos === "bottom") out.push({ ok: list[list.length - 1] === r, label: "En dernière position", hint: "Utilise Move › Bottom.", sub: true });
    return out;
  }
  const hasRule = (cfg, dg, n) => !!cfg && [...cfg[dg].pre, ...cfg[dg].post].some(r => r.name === n && !r.disabled);
  const CUSTOM = {
    commitPush: () => [
      { ok: !pending(), label: "Commit to Panorama fait (aucune modification en attente)" },
      ...DGS.map(dg => ({ ok: !!(st.lastPush[dg] && st.lastPush[dg].ok) && inSync(dg), label: `${DG[dg].fw} : dernier push réussi et In sync`, hint: st.lastPush[dg] && !st.lastPush[dg].ok ? "Le dernier push a échoué : lis le job." : "" })),
      { ok: hasRule(st.pushed["DG-Perimetre"], "DG-Perimetre", "Allow-Web-DMZ"), label: "FW-PERIM exécute Allow-Web-DMZ", hint: "Termine l'exercice « Publier le serveur web » puis pousse." },
      { ok: hasRule(st.pushed["DG-Interne"], "DG-Interne", "Users-to-Servers"), label: "FW-INT exécute Users-to-Servers", hint: "Termine l'exercice « Utilisateurs → serveurs » puis pousse." },
    ],
    debug: () => {
      const t = st.setups.debug;
      if (!t) return [{ ok: false, label: "Scénario chargé", hint: "Clique « Charger le scénario »." }];
      const pe = validatePano(st.cfg).errors.filter(e => e.includes("DG-Interne")), de = validatePush(st.cfg, "DG-Interne"), lp = st.lastPush["DG-Interne"];
      return [{ ok: true, label: "Scénario chargé" },
        { ok: !pe.length, label: "Commit Panorama sans erreur dans DG-Interne", hint: pe[0] || "" },
        { ok: !de.length, label: "Toutes les zones des règles existent sur FW-INT", hint: de[0] || "" },
        { ok: !!(lp && lp.ok && lp.t > t && inSync("DG-Interne")), label: "DG-Interne poussé avec succès après correction", hint: "Commit and Push." }];
    },
    shadow: () => {
      const t = st.setups.shadow;
      if (!t) return [{ ok: false, label: "Scénario chargé", hint: "Clique « Charger le scénario »." }];
      const tmp = st.cfg["DG-Perimetre"].pre.find(r => r.name === "Temp-Allow-All-Out"), sh = shadows(st.cfg, "DG-Perimetre"), lp = st.lastPush["DG-Perimetre"];
      return [{ ok: true, label: "Scénario chargé" },
        { ok: !tmp || tmp.disabled, label: "Temp-Allow-All-Out supprimée ou désactivée" },
        { ok: !sh.length, label: "Aucun warning de shadowing dans DG-Perimetre", hint: sh.length ? `Rule '${sh[0][0]}' shadows rule '${sh[0][1]}'` : "" },
        { ok: !!(lp && lp.ok && lp.t > t && inSync("DG-Perimetre")), label: "DG-Perimetre poussé après correction" }];
    },
    trace: () => FLOWS.filter(f => f.ex).map(f => ({ ok: !!st.traces[f.id], label: `${f.label} → ${f.expect === "allow" ? "autorisé" : "bloqué"}`, hint: "Lance ce flux dans le packet tracer." })),
  };
  const critOf = ex => (ex.check ? CUSTOM[ex.check]() : ex.sol.flatMap(critItem));

  const SETUP = {
    debug() {
      const pre = st.cfg["DG-Interne"].pre.filter(r => r.name !== "Legacy-DNS");
      pre.unshift({ name: "Legacy-DNS", desc: "Importée de l'ancien pare-feu (2017)", tags: [], from: ["DMZ"], to: ["Servers"], src: ["N-USERS"], dst: ["H-DNS-EXT"], app: ["dns"], svc: ["application-default"], action: "allow", profile: "", lf: "", logEnd: true, disabled: false });
      st.cfg["DG-Interne"].pre = pre;
    },
    shadow() {
      const pre = st.cfg["DG-Perimetre"].pre.filter(r => r.name !== "Temp-Allow-All-Out");
      let i = 0; while (i < pre.length && /^Block-Malicious/.test(pre[i].name)) i++;
      pre.splice(i, 0, { name: "Temp-Allow-All-Out", desc: "TEMPORAIRE en attendant la recette (2019)", tags: [], from: ["Transit"], to: ["Untrust"], src: ["any"], dst: ["any"], app: ["any"], svc: ["any"], action: "allow", profile: "", lf: "", logEnd: true, disabled: false });
      st.cfg["DG-Perimetre"].pre = pre;
      /* la règle masquée doit exister pour que le warning apparaisse : on complète les prérequis manquants */
      if (!st.cfg["DG-Perimetre"].pre.some(r => r.name === "Internal-to-Internet"))
        ["tags", "addr-int", "grp", "perim-out"].forEach(id => { const ex = EX.find(e => e.id === id); if (!st.done[id]) st.assisted[id] = true; applySol(ex); });
    },
  };
  function applySol(ex) {
    for (const it of ex.sol) {
      const o = clone(it.o);
      if (KL[it.k]) { const L = st.cfg[it.loc][it.k], i = L.findIndex(x => x.name === o.name); i >= 0 ? (L[i] = o) : L.push(o); continue; }
      const bases = it.k === "nat" ? ["nat"] : ["pre", "post"];
      bases.forEach(b => { st.cfg[it.dg][b] = st.cfg[it.dg][b].filter(r => r.name !== o.name); });
      const L = st.cfg[it.dg][it.k === "nat" ? "nat" : it.base];
      if (it.pos === "top") L.unshift(o); else L.push(o);
    }
  }

  /* =========================================================
     PROGRESSION
     ========================================================= */
  const RANKS = [[0, "Stagiaire"], [400, "Admin junior"], [1000, "Ingénieur sécurité"], [1800, "Architecte NGFW"], [2600, "Expert PCNSE"]];
  const xp = () => EX.reduce((s, e) => s + (st.done[e.id] ? (st.assisted[e.id] ? 30 : 100) : 0), 0) + CLI.filter(x => st.cli[x.id]).length * 50 + (st.quiz.best || 0) * 10;
  const rank = x => RANKS.filter(r => x >= r[0]).pop()[1];
  function checkAll(silent) {
    let fresh = null;
    for (const ex of EX) {
      if (st.done[ex.id]) continue;
      const cr = critOf(ex);
      if (cr.length && cr.every(x => x.ok)) { st.done[ex.id] = Date.now(); fresh = fresh || ex; if (!silent) toast(`🎉 Exercice réussi : ${ex.title} (+${st.assisted[ex.id] ? 30 : 100} XP)`); }
    }
    if (fresh && !silent) { st.celebrate = fresh.id; }
  }

  /* =========================================================
     COMPOSANTS UI : modale, sélecteur multiple
     ========================================================= */
  function modal({ title, html, onOk, okLabel = "OK", wide, onOpen, cls = "" }) {
    const m = document.createElement("div");
    m.className = "lb-modal " + cls;
    m.innerHTML = `<div class="lb-dlg${wide ? " wide" : ""}" role="dialog" aria-modal="true" aria-label="${esc(title)}">
      <header><span>${esc(title)}</span><button type="button" class="x" data-close aria-label="Fermer">×</button></header>
      <div class="lb-dlg-b">${html}</div><div class="lb-dlg-err" role="alert"></div>
      <footer>${onOk ? `<button type="button" class="btn sm primary" data-ok>${okLabel}</button>` : ""}<button type="button" class="btn sm ghost" data-close>${onOk ? "Cancel" : "Fermer"}</button></footer></div>`;
    document.body.append(m);
    const prev = document.activeElement;
    requestAnimationFrame(() => m.classList.add("in"));
    const close = () => { m.classList.remove("in"); document.removeEventListener("keydown", onKey); setTimeout(() => m.remove(), 250); prev && prev.focus && prev.focus(); };
    const onKey = e => { if (e.key === "Escape") close(); };
    document.addEventListener("keydown", onKey);
    m.addEventListener("click", e => {
      if (e.target === m || e.target.closest("[data-close]")) return close();
      if (e.target.closest("[data-ok]")) {
        const err = onOk(m);
        if (err) { const box = $(".lb-dlg-err", m); box.textContent = err; const d = $(".lb-dlg", m); d.classList.remove("shake"); void d.offsetWidth; d.classList.add("shake"); }
        else close();
      }
    });
    m.close = close;
    if (onOpen) onOpen(m);
    setTimeout(() => { const f = $("input,select,button", $(".lb-dlg-b", m)) || $("[data-close]", m); f && f.focus(); }, 60);
    return m;
  }
  /* groups : [[label, [values]]] */
  function picker(host, { groups, val, any = true }) {
    let v = [...val];
    const draw = () => {
      const isAny = any && v.includes("any");
      const rest = groups.map(([g, vals]) => [g, vals.filter(x => !v.includes(x))]).filter(([, vals]) => vals.length);
      host.innerHTML = `${any ? `<label class="pk-any"><input type="checkbox"${isAny ? " checked" : ""}> any</label>` : ""}
        <div class="pk-box${isAny ? " dis" : ""}">${isAny ? `<span class="mut small">Tout est accepté</span>` : v.map(x => `<span class="pk-chip">${esc(x)}<button type="button" data-rm="${esc(x)}" aria-label="Retirer ${esc(x)}">×</button></span>`).join("")}
        ${isAny ? "" : `<select class="pk-add" aria-label="Ajouter"><option value="">+ Add…</option>${rest.map(([g, vals]) => `<optgroup label="${esc(g)}">${vals.map(x => `<option>${esc(x)}</option>`).join("")}</optgroup>`).join("")}</select>`}</div>`;
    };
    host.addEventListener("change", e => {
      if (e.target.matches(".pk-any input")) { v = e.target.checked ? ["any"] : []; draw(); }
      else if (e.target.matches(".pk-add") && e.target.value) { v.push(e.target.value); draw(); const s = $(".pk-add", host); s && s.focus(); }
    });
    host.addEventListener("click", e => { const b = e.target.closest("[data-rm]"); if (b) { v = v.filter(x => x !== b.dataset.rm); draw(); } });
    draw();
    return { get: () => v.slice() };
  }
  const opt = (vals, cur) => vals.map(x => { const [v, l] = Array.isArray(x) ? x : [x, x]; return `<option value="${esc(v)}"${v === cur ? " selected" : ""}>${esc(l)}</option>`; }).join("");

  /* =========================================================
     MUTATIONS
     ========================================================= */
  function changed(flash) {
    if (flash) st.ui.flash = flash;
    checkAll();
    save(); renderPano(); renderMissions(); renderProgress(); renderTraceSync();
  }
  function refsTo(name, loc) {
    const refs = [];
    for (const l of loc === "Shared" ? LOCS : [loc]) {
      st.cfg[l].grp.forEach(g => g.members.includes(name) && refs.push(`address-group ${g.name} (${l})`));
      st.cfg[l].addr.forEach(a => (a.tags || []).includes(name) && refs.push(`address ${a.name} (${l})`));
      if (l === "Shared") continue;
      for (const b of ["pre", "post"]) st.cfg[l][b].forEach(r => ["src", "dst", "svc", "tags"].some(f => r[f].includes(name)) && refs.push(`security rule ${r.name} (${l})`));
      st.cfg[l].nat.forEach(r => [...r.src, ...r.dst, r.dnat, r.svc].includes(name) && refs.push(`nat rule ${r.name} (${l})`));
    }
    return refs;
  }
  function renameRefs(a, b, loc) {
    const sw = arr => arr.map(x => (x === a ? b : x));
    for (const l of loc === "Shared" ? LOCS : [loc]) {
      st.cfg[l].grp.forEach(g => (g.members = sw(g.members)));
      st.cfg[l].addr.forEach(x => (x.tags = sw(x.tags || [])));
      if (l === "Shared") continue;
      for (const bs of ["pre", "post"]) st.cfg[l][bs].forEach(r => ["src", "dst", "svc", "tags"].forEach(f => (r[f] = sw(r[f]))));
      st.cfg[l].nat.forEach(r => { r.src = sw(r.src); r.dst = sw(r.dst); if (r.dnat === a) r.dnat = b; if (r.svc === a) r.svc = b; });
    }
  }
  const nameErr = n => (!n ? "Name : champ obligatoire." : !NAME_RE.test(n) ? "Name : 63 caractères max, doit commencer par une lettre ou un chiffre (lettres, chiffres, espace, _ . - autorisés)." : "");

  /* ---------- éditeurs d'objets ---------- */
  function editObject(k, loc, name) {
    const ctx = st.ui.ctx, isNew = !name;
    const cur = isNew ? null : st.cfg[loc][k].find(x => x.name === name);
    const L0 = isNew ? ctx : loc;
    const locSel = `<div class="field"><label>Location</label><select id="o-loc"${isNew ? "" : " disabled"}>${opt(LOCS, L0)}</select></div>`;
    const tagG = [["Tags", visible(st.cfg, ctx, "tag").map(x => x.o.name)]];
    let html = `<div class="f2"><div class="field"><label for="o-name">Name</label><input type="text" id="o-name" maxlength="63" value="${esc(cur ? cur.name : "")}"></div>${locSel}</div>`;
    if (k === "addr") html += `<div class="f2"><div class="field"><label for="o-type">Type</label><select id="o-type">${opt([["ip-netmask", "IP Netmask"], ["ip-range", "IP Range"], ["fqdn", "FQDN"]], cur ? cur.type : "ip-netmask")}</select></div>
      <div class="field"><label for="o-val">Valeur</label><input type="text" id="o-val" value="${esc(cur ? cur.value : "")}" placeholder="10.10.10.0/24"></div></div>`;
    if (k === "grp") html += `<div class="field"><span class="lbl">Type : Static · Members</span><div id="o-mem" class="pk"></div></div>`;
    if (k === "svc") html += `<div class="f2"><div class="field"><label for="o-proto">Protocol</label><select id="o-proto">${opt([["tcp", "TCP"], ["udp", "UDP"]], cur ? cur.proto : "tcp")}</select></div>
      <div class="field"><label for="o-port">Destination Port</label><input type="text" id="o-port" value="${esc(cur ? cur.port : "")}" placeholder="8443"></div></div>`;
    if (k === "tag") html += `<div class="field"><span class="lbl">Color</span><div class="sw-row" role="radiogroup">${Object.entries(COLORS).map(([id, [n, hx]]) => `<label class="sw"><input type="radio" name="o-col" value="${id}"${(cur ? cur.color : "color1") === id ? " checked" : ""}><i style="background:${hx}"></i>${n}</label>`).join("")}</div></div>`;
    if (k !== "tag") html += `<div class="field"><label for="o-desc">Description</label><input type="text" id="o-desc" value="${esc(cur ? cur.desc || "" : "")}"></div>`;
    if (k === "addr") html += `<div class="field"><span class="lbl">Tags</span><div id="o-tags" class="pk"></div></div>`;
    let pk = {};
    modal({ title: `${{ addr: "Address", grp: "Address Group", svc: "Service", tag: "Tag" }[k]}${isNew ? "" : " · " + name}`, html, okLabel: "OK",
      onOpen: m => {
        if (k === "addr") pk.tags = picker($("#o-tags", m), { groups: tagG, val: cur ? cur.tags || [] : [], any: false });
        if (k === "grp") pk.mem = picker($("#o-mem", m), { groups: [["Addresses", visible(st.cfg, ctx, "addr").map(x => x.o.name)], ["Address Groups", visible(st.cfg, ctx, "grp").map(x => x.o.name).filter(n => n !== name)]], val: cur ? cur.members : [], any: false });
      },
      onOk: m => {
        const nm = $("#o-name", m).value.trim(), L = $("#o-loc", m).value;
        const e = nameErr(nm); if (e) return e;
        const ns = k === "addr" || k === "grp" ? ["addr", "grp"] : [k];
        if ((nm !== name || L !== loc) && ns.some(kk => st.cfg[L][kk].some(x => x.name === nm))) return `'${nm}' existe déjà dans ${L}.`;
        const o = { name: nm };
        if (k === "addr") {
          o.type = $("#o-type", m).value; o.value = $("#o-val", m).value.trim(); o.tags = pk.tags.get(); o.desc = $("#o-desc", m).value.trim();
          if (o.type === "ip-netmask" && !validNet(o.value)) return "Valeur invalide : adresse IP ou CIDR attendu (ex. 10.10.10.0/24).";
          if (o.type === "ip-range") { const [a, b] = o.value.split("-"); if (!validIp(a || "") || !validIp(b || "") || ip2i(a) > ip2i(b)) return "Plage invalide (ex. 10.0.0.10-10.0.0.20)."; }
          if (o.type === "fqdn" && !/^(?=.{1,253}$)([a-z0-9-]+\.)+[a-z]{2,}$/i.test(o.value)) return "FQDN invalide.";
          if (L === "Shared" && o.tags.some(t => !lookup(st.cfg, "Shared", "tag", t))) return "Un objet Shared ne peut utiliser que des tags Shared.";
        }
        if (k === "grp") {
          o.members = pk.mem.get(); o.desc = $("#o-desc", m).value.trim();
          if (!o.members.length) return "Un groupe statique doit contenir au moins un membre.";
          const bad = o.members.find(x => !addrExists(st.cfg, L, x)); if (bad) return `'${bad}' n'est pas visible depuis ${L} : un groupe Shared ne peut contenir que des objets Shared.`;
        }
        if (k === "svc") { o.proto = $("#o-proto", m).value; o.port = $("#o-port", m).value.replace(/\s/g, ""); o.desc = $("#o-desc", m).value.trim(); if (!validPorts(o.port)) return "Port invalide (ex. 8443, 8000-8100 ou 80,8080)."; }
        if (k === "tag") o.color = ($("input[name=o-col]:checked", m) || {}).value || "color1";
        const list = st.cfg[L][k];
        if (isNew) list.push(o); else { list[list.indexOf(cur)] = o; if (nm !== name) renameRefs(name, nm, L); }
        st.ui.sel = `${L}|${nm}`;
        changed(`${L}|${nm}`);
        toast(isNew ? `${KL[k]} ${nm} créé dans ${L} (candidate)` : `${KL[k]} ${nm} modifié`);
      } });
  }

  /* ---------- éditeur de règle de sécurité ---------- */
  function addrGroups(ctx) {
    return [["Addresses", visible(st.cfg, ctx, "addr").map(x => x.o.name)], ["Address Groups", visible(st.cfg, ctx, "grp").map(x => x.o.name)], ["External Dynamic Lists", Object.keys(EDL)]];
  }
  const blankRule = () => ({ name: "", desc: "", tags: [], from: ["any"], to: ["any"], src: ["any"], dst: ["any"], app: ["any"], svc: ["application-default"], action: "allow", profile: "", lf: "", logEnd: true, disabled: false });
  function editRule(dg, base, name, cloneOf) {
    const isNew = !name, src = isNew ? (cloneOf ? Object.assign(clone(cloneOf), { name: cloneOf.name.slice(0, 57) + "-copie" }) : blankRule()) : st.cfg[dg][base].find(r => r.name === name);
    const r = clone(src);
    const svcMode = r.svc.includes("any") ? "any" : r.svc.includes("application-default") ? "application-default" : "select";
    const tabs = [["general", "General"], ["source", "Source"], ["destination", "Destination"], ["application", "Application"], ["service", "Service/URL Category"], ["actions", "Actions"]];
    const html = `<div class="dt-tabs" role="tablist">${tabs.map(([id, l], i) => `<button type="button" role="tab" data-dt="${id}" class="${i ? "" : "on"}">${l}</button>`).join("")}</div>
      <div class="dt-pane on" data-pane="general">
        <div class="field"><label for="r-name">Name</label><input type="text" id="r-name" maxlength="63" value="${esc(r.name)}" placeholder="ex. Users-to-Servers"></div>
        <div class="field"><label for="r-desc">Description</label><input type="text" id="r-desc" value="${esc(r.desc)}"></div>
        <div class="field"><span class="lbl">Tags</span><div id="r-tags" class="pk"></div></div>
        <div class="f2"><div class="field"><label for="r-base">Rulebase</label><select id="r-base">${opt([["pre", "Pre Rules"], ["post", "Post Rules"]], base)}</select></div>
        <div class="field"><label>Device Group</label><input type="text" value="${dg}" disabled></div></div>
      </div>
      <div class="dt-pane" data-pane="source"><div class="field"><span class="lbl">Source Zone</span><div id="r-from" class="pk"></div></div><div class="field"><span class="lbl">Source Address</span><div id="r-src" class="pk"></div></div></div>
      <div class="dt-pane" data-pane="destination"><div class="field"><span class="lbl">Destination Zone</span><div id="r-to" class="pk"></div></div><div class="field"><span class="lbl">Destination Address</span><div id="r-dst" class="pk"></div></div>
        <p class="small mut">Zones proposées : toutes celles des template stacks. Une zone absente du pare-feu cible fera échouer le push.</p></div>
      <div class="dt-pane" data-pane="application"><div class="field"><span class="lbl">Applications</span><div id="r-app" class="pk"></div></div></div>
      <div class="dt-pane" data-pane="service"><div class="field"><label for="r-svcm">Service</label><select id="r-svcm">${opt([["application-default", "application-default"], ["any", "any"], ["select", "select"]], svcMode)}</select></div><div id="r-svc" class="pk"${svcMode === "select" ? "" : " hidden"}></div></div>
      <div class="dt-pane" data-pane="actions">
        <div class="f2"><div class="field"><label for="r-act">Action</label><select id="r-act">${opt(["allow", "deny", "drop", "reset-client", "reset-server", "reset-both"], r.action)}</select></div>
        <div class="field"><label for="r-prof">Profile Setting · Group</label><select id="r-prof">${opt([["", "None"], ...Object.keys(PROFILES)], r.profile)}</select></div></div>
        <div class="f2"><div class="field"><label for="r-lf">Log Forwarding</label><select id="r-lf">${opt([["", "None"], ...Object.keys(LOGFWD)], r.lf)}</select></div>
        <div class="field"><label class="ck"><input type="checkbox" id="r-log"${r.logEnd ? " checked" : ""}> Log at Session End</label></div></div>
      </div>`;
    const pk = {};
    modal({ title: `Security Policy Rule${isNew ? "" : " · " + name}`, html, wide: true,
      onOpen: m => {
        const zones = [["Zones", ALL_ZONES]];
        pk.tags = picker($("#r-tags", m), { groups: [["Tags", visible(st.cfg, dg, "tag").map(x => x.o.name)]], val: r.tags, any: false });
        pk.from = picker($("#r-from", m), { groups: zones, val: r.from });
        pk.to = picker($("#r-to", m), { groups: zones, val: r.to });
        pk.src = picker($("#r-src", m), { groups: addrGroups(dg), val: r.src });
        pk.dst = picker($("#r-dst", m), { groups: addrGroups(dg), val: r.dst });
        pk.app = picker($("#r-app", m), { groups: [["Applications", Object.keys(APPS)]], val: r.app });
        pk.svc = picker($("#r-svc", m), { groups: [["Services", [...Object.keys(PRE_SVC), ...visible(st.cfg, dg, "svc").map(x => x.o.name)]]], val: svcMode === "select" ? r.svc : [], any: false });
        $(".dt-tabs", m).addEventListener("click", e => { const b = e.target.closest("[data-dt]"); if (!b) return; $$("[data-dt]", m).forEach(x => x.classList.toggle("on", x === b)); $$("[data-pane]", m).forEach(p => p.classList.toggle("on", p.dataset.pane === b.dataset.dt)); });
        $("#r-svcm", m).addEventListener("change", e => { $("#r-svc", m).hidden = e.target.value !== "select"; });
      },
      onOk: m => {
        const goto = id => { $(`[data-dt="${id}"]`, m).click(); };
        const n = $("#r-name", m).value.trim(), nb = $("#r-base", m).value;
        const e = nameErr(n); if (e) { goto("general"); return e; }
        if (n !== name && [...st.cfg[dg].pre, ...st.cfg[dg].post].some(x => x.name === n)) { goto("general"); return `Une règle '${n}' existe déjà dans ${dg} (les noms sont uniques par device group).`; }
        const o = { name: n, desc: $("#r-desc", m).value.trim(), tags: pk.tags.get(), from: pk.from.get(), to: pk.to.get(), src: pk.src.get(), dst: pk.dst.get(), app: pk.app.get(),
          svc: $("#r-svcm", m).value === "select" ? pk.svc.get() : [$("#r-svcm", m).value], action: $("#r-act", m).value, profile: $("#r-prof", m).value, lf: $("#r-lf", m).value, logEnd: $("#r-log", m).checked, disabled: isNew ? false : r.disabled };
        for (const [f, tab, lab] of [["from", "source", "Source Zone"], ["src", "source", "Source Address"], ["to", "destination", "Destination Zone"], ["dst", "destination", "Destination Address"], ["app", "application", "Application"], ["svc", "service", "Service"]])
          if (!o[f].length) { goto(tab); return `${lab} : choisis au moins une valeur ou « any ».`; }
        if (o.profile && o.action !== "allow") o.profile = "";
        if (isNew) st.cfg[dg][nb].push(o);
        else {
          const L = st.cfg[dg][base], i = L.findIndex(x => x.name === name);
          if (nb === base) L[i] = o; else { L.splice(i, 1); st.cfg[dg][nb].push(o); }
        }
        st.ui.node = nb; st.ui.sel = n;
        changed(n);
        toast(isNew ? `Règle ${n} ajoutée en bas des ${baseLabel(nb)} de ${dg}` : `Règle ${n} modifiée`);
      } });
  }

  /* ---------- éditeur NAT ---------- */
  const blankNat = () => ({ name: "", desc: "", from: ["any"], to: "any", src: ["any"], dst: ["any"], svc: "any", snat: { type: "none", iface: "" }, dnat: "", disabled: false });
  function editNat(dg, name, cloneOf) {
    const isNew = !name, r = clone(isNew ? (cloneOf ? Object.assign(clone(cloneOf), { name: cloneOf.name.slice(0, 57) + "-copie" }) : blankNat()) : st.cfg[dg].nat.find(x => x.name === name));
    const svcs = ["any", ...Object.keys(PRE_SVC), ...visible(st.cfg, dg, "svc").map(x => x.o.name)];
    const addrs = visible(st.cfg, dg, "addr").map(x => x.o.name);
    const tabs = [["general", "General"], ["orig", "Original Packet"], ["trans", "Translated Packet"]];
    const html = `<div class="dt-tabs" role="tablist">${tabs.map(([id, l], i) => `<button type="button" role="tab" data-dt="${id}" class="${i ? "" : "on"}">${l}</button>`).join("")}</div>
      <div class="dt-pane on" data-pane="general"><div class="field"><label for="n-name">Name</label><input type="text" id="n-name" maxlength="63" value="${esc(r.name)}"></div>
        <div class="field"><label for="n-desc">Description</label><input type="text" id="n-desc" value="${esc(r.desc)}"></div></div>
      <div class="dt-pane" data-pane="orig">
        <div class="field"><span class="lbl">Source Zone</span><div id="n-from" class="pk"></div></div>
        <div class="f2"><div class="field"><label for="n-to">Destination Zone</label><select id="n-to">${opt(["any", ...ALL_ZONES], r.to)}</select></div>
        <div class="field"><label for="n-svc">Service</label><select id="n-svc">${opt(svcs, r.svc)}</select></div></div>
        <div class="field"><span class="lbl">Source Address</span><div id="n-src" class="pk"></div></div>
        <div class="field"><span class="lbl">Destination Address</span><div id="n-dst" class="pk"></div></div>
        <p class="small mut">La zone de destination d'une règle NAT est celle de la route vers l'adresse <b>d'origine</b> (avant traduction).</p></div>
      <div class="dt-pane" data-pane="trans">
        <div class="f2"><div class="field"><label for="n-st">Source Address Translation</label><select id="n-st">${opt([["none", "None"], ["dipp", "Dynamic IP And Port · Interface Address"]], r.snat.type)}</select></div>
        <div class="field"><label for="n-if">Interface</label><select id="n-if">${opt([["", "—"], ...ALL_IFACES], r.snat.iface)}</select></div></div>
        <div class="field"><label for="n-dn">Destination Address Translation · Static IP · Translated Address</label><select id="n-dn">${opt([["", "None"], ...addrs], r.dnat)}</select></div></div>`;
    const pk = {};
    modal({ title: `NAT Policy Rule${isNew ? "" : " · " + name}`, html, wide: true,
      onOpen: m => {
        pk.from = picker($("#n-from", m), { groups: [["Zones", ALL_ZONES]], val: r.from });
        pk.src = picker($("#n-src", m), { groups: addrGroups(dg), val: r.src });
        pk.dst = picker($("#n-dst", m), { groups: addrGroups(dg), val: r.dst });
        $(".dt-tabs", m).addEventListener("click", e => { const b = e.target.closest("[data-dt]"); if (!b) return; $$("[data-dt]", m).forEach(x => x.classList.toggle("on", x === b)); $$("[data-pane]", m).forEach(p => p.classList.toggle("on", p.dataset.pane === b.dataset.dt)); });
      },
      onOk: m => {
        const n = $("#n-name", m).value.trim(), e = nameErr(n); if (e) return e;
        if (n !== name && st.cfg[dg].nat.some(x => x.name === n)) return `Une règle NAT '${n}' existe déjà dans ${dg}.`;
        const o = { name: n, desc: $("#n-desc", m).value.trim(), from: pk.from.get(), to: $("#n-to", m).value, src: pk.src.get(), dst: pk.dst.get(), svc: $("#n-svc", m).value,
          snat: { type: $("#n-st", m).value, iface: $("#n-st", m).value === "dipp" ? $("#n-if", m).value : "" }, dnat: $("#n-dn", m).value, disabled: r.disabled };
        if (!o.from.length || !o.src.length || !o.dst.length) return "Zones et adresses : choisis au moins une valeur ou « any ».";
        if (o.snat.type === "dipp" && !o.snat.iface) { $('[data-dt="trans"]', m).click(); return "Dynamic IP And Port : choisis l'interface."; }
        if (o.snat.type === "none" && !o.dnat) { $('[data-dt="trans"]', m).click(); return "Une règle NAT doit traduire la source, la destination, ou les deux."; }
        const L = st.cfg[dg].nat;
        if (isNew) L.push(o); else L[L.findIndex(x => x.name === name)] = o;
        st.ui.sel = n; changed(n);
        toast(isNew ? `Règle NAT ${n} ajoutée` : `Règle NAT ${n} modifiée`);
      } });
  }

  /* =========================================================
     CONSOLE PANORAMA
     ========================================================= */
  const TABS = [["panorama", "Panorama"], ["policies", "Policies"], ["objects", "Objects"]];
  const TREE = {
    objects: [["addr", "Addresses", "📍"], ["grp", "Address Groups", "🗂️"], ["svc", "Services", "🔌"], ["tag", "Tags", "🏷️"], ["edl", "External Dynamic Lists", "📜"], ["prof", "Security Profile Groups", "🛡️"], ["lf", "Log Forwarding", "📤"]],
    policies: [["pre", "Security · Pre Rules", "⤒"], ["post", "Security · Post Rules", "⤓"], ["nat", "NAT · Pre Rules", "🔀"], ["def", "Security · Default Rules", "⚙️"]],
    panorama: [["devices", "Managed Devices", "🖥️"], ["dgs", "Device Groups", "📁"], ["tpl", "Templates", "🧩"], ["jobs", "Jobs", "🧾"]],
  };
  const pz = $("#pano");
  const cells = v => (v.length === 1 && v[0] === "any" ? `<span class="any">any</span>` : v.map(x => `<span class="cv">${esc(x)}</span>`).join(""));
  const zcells = v => (v.length === 1 && v[0] === "any" ? `<span class="any">any</span>` : v.map(z => `<span class="zt ${ZCLS[z] || ""}">${esc(z)}</span>`).join(""));
  const tagChip = (t, ctx) => { const o = lookup(st.cfg, ctx, "tag", t); const col = o && COLORS[o.color] ? COLORS[o.color][1] : "#64748B"; return `<span class="tg" style="--tc:${col}">${esc(t)}</span>`; };

  function renderPano() {
    const ui = st.ui;
    if (ui.tab === "policies" && ui.ctx === "Shared") ui.ctx = "DG-Perimetre";
    const pend = pending(), outs = DGS.filter(d => !inSync(d));
    const crumbs = { objects: "Objects", policies: "Policies", panorama: "Panorama" };
    const node = TREE[ui.tab].find(n => n[0] === ui.node) || TREE[ui.tab][0];
    ui.node = node[0];
    pz.innerHTML = `
      <div class="pz-top">
        <div class="pz-logo" aria-hidden="true"><b>▰</b><span>PANORAMA</span></div>
        <nav class="pz-tabs" aria-label="Onglets">${TABS.map(([id, l]) => `<button type="button" data-tab="${id}" class="${ui.tab === id ? "on" : ""}">${l}</button>`).join("")}</nav>
        ${ui.tab !== "panorama" ? `<label class="pz-ctx">Device Group <select id="pzCtx">${opt(ui.tab === "policies" ? DGS : LOCS, ui.ctx)}</select></label>` : `<span class="pz-ctx"></span>`}
        <div class="pz-cm">
          <button type="button" class="pz-commit${pend ? " pend" : ""}" data-act="menu" aria-haspopup="true" aria-expanded="false">Commit${pend ? `<i class="pz-dot" title="Modifications en attente"></i>` : ""} ▾</button>
          <div class="pz-menu" role="menu" hidden>
            <button role="menuitem" type="button" data-act="commit">Commit to Panorama</button>
            <button role="menuitem" type="button" data-act="push">Push to Devices</button>
            <button role="menuitem" type="button" data-act="commitpush">Commit and Push</button>
            <hr><button role="menuitem" type="button" data-act="preview">Preview Changes</button>
            <button role="menuitem" type="button" data-act="revert">Revert Changes</button>
          </div>
        </div>
      </div>
      <div class="pz-body">
        <aside class="pz-tree" aria-label="Arborescence">${TREE[ui.tab].map(([id, l, ic]) => `<button type="button" data-node="${id}" class="${ui.node === id ? "on" : ""}"><span aria-hidden="true">${ic}</span>${l}</button>`).join("")}</aside>
        <div class="pz-main">
          <div class="pz-crumb">${crumbs[ui.tab]} › <b>${node[1]}</b>${ui.tab !== "panorama" ? ` <span class="mut">· ${ui.ctx}</span>` : ""}</div>
          <div class="pz-tool">${toolbar()}</div>
          <div class="pz-table">${table()}</div>
        </div>
      </div>
      <div class="pz-status"><span>👤 admin</span><span>Panorama 11.1.4</span>
        <span class="${pend ? "w" : "g"}">${pend ? "● Modifications non commitées" : "● Candidate = running"}</span>
        <span class="${outs.length ? "w" : "g"}">${outs.length ? "⟳ Out of sync : " + outs.map(d => DG[d].fw).join(", ") : "✓ Pare-feux In sync"}</span>
        <button type="button" class="pz-jobs" data-act="jobs">🧾 Tasks (${st.jobs.length})</button></div>`;
    if (ui.flash) {
      const row = pz.querySelector(`tr[data-key="${CSS.escape(ui.flash)}"]`);
      if (row) { row.classList.add("flash"); row.scrollIntoView({ block: "nearest" }); }
      ui.flash = null;
    }
  }

  function toolbar() {
    const n = st.ui.node, b = (a, l, extra = "") => `<button type="button" data-act="${a}" ${extra}>${l}</button>`;
    if (["addr", "grp", "svc", "tag"].includes(n)) return b("add", "＋ Add") + b("del", "🗑 Delete") + (n === "addr" ? b("clone", "⧉ Clone") : "") + `<span class="pz-hint">Clique un nom pour l'éditer</span>`;
    if (["pre", "post", "nat"].includes(n)) return b("add", "＋ Add") + b("del", "🗑 Delete") + b("clone", "⧉ Clone") + b("enable", "✓ Enable") + b("disable", "⊘ Disable") +
      `<span class="pz-sep"></span>` + b("top", "⤒ Top") + b("up", "↑ Up") + b("down", "↓ Down") + b("bottom", "⤓ Bottom");
    if (n === "devices") return b("commit", "Commit to Panorama") + b("push", "Push to Devices");
    return `<span class="pz-hint">Lecture seule dans ce lab</span>`;
  }

  function table() {
    const ui = st.ui, ctx = ui.ctx, n = ui.node, sel = k => (ui.sel === k ? " sel" : "");
    const locCell = l => `<span class="loc ${l === "Shared" ? "sh" : ""}">${l}</span>`;
    const empty = (cols, msg) => `<tr><td colspan="${cols}" class="empty">${msg}</td></tr>`;
    if (n === "addr" || n === "grp" || n === "svc" || n === "tag") {
      const rows = visible(st.cfg, ctx, n);
      const head = { addr: ["Name", "Location", "Type", "Address", "Tags"], grp: ["Name", "Location", "Members", "Tags"], svc: ["Name", "Location", "Protocol", "Destination Port"], tag: ["Name", "Location", "Color"] }[n];
      const body = rows.map(({ o, loc }) => {
        const k = `${loc}|${o.name}`, nm = `<button type="button" class="lnk" data-act="edit" data-key="${esc(k)}">${esc(o.name)}</button>${o.desc ? `<small>${esc(o.desc)}</small>` : ""}`;
        const cols = {
          addr: () => [nm, locCell(loc), { "ip-netmask": "IP Netmask", "ip-range": "IP Range", fqdn: "FQDN" }[o.type], `<span class="mono">${esc(o.value)}</span>`, (o.tags || []).map(t => tagChip(t, loc)).join("")],
          grp: () => [nm, locCell(loc), `<span class="mono small">${o.members.map(esc).join(", ")}</span>`, ""],
          svc: () => [nm, locCell(loc), o.proto.toUpperCase(), `<span class="mono">${esc(o.port)}</span>`],
          tag: () => [nm, locCell(loc), `<span class="tg" style="--tc:${COLORS[o.color][1]}">${COLORS[o.color][0]}</span>`] }[n]();
        return `<tr data-key="${esc(k)}" class="${sel(k)}">${cols.map(x => `<td>${x}</td>`).join("")}</tr>`;
      }).join("");
      return `<table class="pz-t"><thead><tr>${head.map(h => `<th>${h}</th>`).join("")}</tr></thead><tbody>${body || empty(head.length, `Aucun objet visible depuis ${ctx}. Clique ＋ Add.`)}</tbody></table>`;
    }
    if (n === "edl") return `<table class="pz-t"><thead><tr><th>Name</th><th>Location</th><th>Type</th><th>Source</th><th>Exemple d'entrée (lab)</th></tr></thead><tbody>${Object.entries(EDL).map(([k, v]) => `<tr><td><b>${k}</b></td><td><span class="loc sh">Predefined</span></td><td>Predefined IP List</td><td>Palo Alto Networks</td><td class="mono">${v.join(", ")}</td></tr>`).join("")}</tbody></table>`;
    if (n === "prof") return `<table class="pz-t"><thead><tr><th>Name</th><th>Location</th><th>Profils</th></tr></thead><tbody>${Object.entries(PROFILES).map(([k, v]) => `<tr><td><b>${k}</b></td><td><span class="loc sh">Shared</span></td><td class="small">${v}</td></tr>`).join("")}</tbody></table>`;
    if (n === "lf") return `<table class="pz-t"><thead><tr><th>Name</th><th>Location</th><th>Match list</th></tr></thead><tbody>${Object.entries(LOGFWD).map(([k, v]) => `<tr><td><b>${k}</b></td><td><span class="loc sh">Shared</span></td><td class="small">${v}</td></tr>`).join("")}</tbody></table>`;
    if (n === "pre" || n === "post") {
      const L = st.cfg[ctx][n], sh = new Map(shadows(st.cfg, ctx).map(([a, b]) => [b, a]));
      const off = n === "post" ? st.cfg[ctx].pre.length : 0;
      const body = L.map((r, i) => `<tr data-key="${esc(r.name)}" class="${sel(r.name)}${r.disabled ? " dis" : ""}">
        <td class="num">${i + 1 + off}</td>
        <td><button type="button" class="lnk" data-act="edit" data-key="${esc(r.name)}">${esc(r.name)}</button>${sh.has(r.name) ? `<span class="shw" title="Masquée par ${esc(sh.get(r.name))}">⚠ shadowed</span>` : ""}${r.tags.map(t => tagChip(t, ctx)).join("")}</td>
        <td>${zcells(r.from)}</td><td>${cells(r.src)}</td><td>${zcells(r.to)}</td><td>${cells(r.dst)}</td><td>${cells(r.app)}</td><td>${cells(r.svc)}</td>
        <td><span class="act ${r.action === "allow" ? "al" : "dn"}">${r.action === "allow" ? "✓" : "⊘"} ${r.action}</span></td>
        <td class="opt">${r.profile ? `<span title="${esc(r.profile)}">🛡️</span>` : ""}${r.lf ? `<span title="${esc(r.lf)}">📤</span>` : ""}${r.logEnd ? "" : `<span title="Pas de log">🚫</span>`}</td></tr>`).join("");
      return `<table class="pz-t rules"><thead><tr><th>#</th><th>Name</th><th>Source Zone</th><th>Source Address</th><th>Destination Zone</th><th>Destination Address</th><th>Application</th><th>Service</th><th>Action</th><th>Options</th></tr></thead>
        <tbody>${body || empty(10, `Aucune règle dans ${ctx} › ${baseLabel(n)}. Clique ＋ Add.`)}</tbody></table>`;
    }
    if (n === "nat") {
      const body = st.cfg[ctx].nat.map((r, i) => `<tr data-key="${esc(r.name)}" class="${sel(r.name)}${r.disabled ? " dis" : ""}"><td class="num">${i + 1}</td>
        <td><button type="button" class="lnk" data-act="edit" data-key="${esc(r.name)}">${esc(r.name)}</button></td><td>${zcells(r.from)}</td><td>${zcells([r.to])}</td><td>${cells(r.src)}</td><td>${cells(r.dst)}</td><td>${cells([r.svc])}</td>
        <td class="small">${r.snat.type === "dipp" ? `DIPP · ${r.snat.iface}` : `<span class="any">none</span>`}</td><td class="small">${r.dnat ? `→ ${esc(r.dnat)}` : `<span class="any">none</span>`}</td></tr>`).join("");
      return `<table class="pz-t rules"><thead><tr><th>#</th><th>Name</th><th>Source Zone</th><th>Dest. Zone</th><th>Source Address</th><th>Destination Address</th><th>Service</th><th>Source Translation</th><th>Destination Translation</th></tr></thead>
        <tbody>${body || empty(9, `Aucune règle NAT dans ${ctx}.`)}</tbody></table>`;
    }
    if (n === "def") return `<table class="pz-t rules"><thead><tr><th>Name</th><th>Type</th><th>Zones</th><th>Action</th><th>Log</th></tr></thead><tbody>
      <tr class="dis"><td>intrazone-default</td><td>intrazone</td><td>même zone</td><td><span class="act al">✓ allow</span></td><td>non (par défaut)</td></tr>
      <tr class="dis"><td>interzone-default</td><td>interzone</td><td>zones différentes</td><td><span class="act dn">⊘ deny</span></td><td>non (par défaut)</td></tr></tbody></table>
      <p class="small mut pz-note">Évaluées après toutes les Pre Rules, règles locales et Post Rules. Elles ne journalisent pas par défaut : un Deny-All-Log en Post Rule rend les refus visibles.</p>`;
    if (n === "devices") return `<table class="pz-t"><thead><tr><th>Device Name</th><th>Model</th><th>Serial</th><th>IP Address</th><th>Device Group</th><th>Template</th><th>Connected</th><th>Shared Policy</th><th>Last Push</th></tr></thead><tbody>${DGS.map(dg => {
      const d = DG[dg], lp = st.lastPush[dg];
      return `<tr><td><b>${d.fw}</b></td><td>${d.model}</td><td class="mono">${d.serial}</td><td class="mono">${d.ip}</td><td>${dg}</td><td>${d.stack}</td><td><span class="dotc g"></span>Connected</td>
        <td>${inSync(dg) ? `<span class="act al">In sync</span>` : `<span class="act wn">Out of sync</span>`}</td>
        <td class="small">${lp ? `${lp.ok ? "✓ Succeeded" : "✗ Failed"} · job ${lp.id} · ${lp.at}` : "—"}</td></tr>`;
    }).join("")}</tbody></table>`;
    if (n === "dgs") return `<table class="pz-t"><thead><tr><th>Device Group</th><th>Parent</th><th>Pare-feu</th><th>Objets</th><th>Règles sécurité</th><th>Règles NAT</th></tr></thead><tbody>${DGS.map(dg => `<tr><td><b>${dg}</b></td><td>Shared</td><td>${DG[dg].fw}</td><td>${["addr", "grp", "svc", "tag"].reduce((s, k) => s + st.cfg[dg][k].length, 0)}</td><td>${st.cfg[dg].pre.length} pre · ${st.cfg[dg].post.length} post</td><td>${st.cfg[dg].nat.length}</td></tr>`).join("")}</tbody></table>`;
    if (n === "tpl") return DGS.map(dg => `<h4 class="pz-h">${DG[dg].stack} → ${DG[dg].fw}</h4><table class="pz-t"><thead><tr><th>Interface</th><th>Zone</th><th>Adresse</th><th>Virtual Router</th></tr></thead><tbody>${Object.entries(DG[dg].ifaces).map(([i, [z, a]]) => `<tr><td class="mono">${i}</td><td>${zcells([z])}</td><td class="mono">${a}</td><td>default</td></tr>`).join("")}</tbody></table>`).join("");
    if (n === "jobs") return jobsTable();
    return "";
  }
  const jobsTable = () => `<table class="pz-t"><thead><tr><th>ID</th><th>Type</th><th>Équipement</th><th>Statut</th><th>Heure</th><th>Détail</th></tr></thead><tbody>${st.jobs.map(j => `<tr><td class="mono">${j.id}</td><td>${j.type}</td><td>${esc(j.dev)}</td><td>${j.ok ? `<span class="act al">OK</span>` : `<span class="act dn">FAIL</span>`}</td><td class="mono small">${j.t}</td><td class="small mono">${esc(j.msg)}</td></tr>`).join("") || `<tr><td colspan="6" class="empty">Aucun job pour l'instant.</td></tr>`}</tbody></table>`;

  /* ---------- actions de la console ---------- */
  function moveRule(L, i, where) {
    if (i < 0) return;
    const [r] = L.splice(i, 1);
    const j = where === "top" ? 0 : where === "bottom" ? L.length : where === "up" ? Math.max(0, i - 1) : Math.min(L.length, i + 1);
    L.splice(j, 0, r);
  }
  function act(a, key) {
    const ui = st.ui, n = ui.node, ctx = ui.ctx;
    const menu = $(".pz-menu", pz);
    if (a === "menu") { const open = menu.hidden; menu.hidden = !open; $(".pz-commit", pz).setAttribute("aria-expanded", String(open)); if (open) $("button", menu).focus(); return; }
    if (["commit", "push", "commitpush", "preview", "revert"].includes(a)) return commitAction(a);
    if (a === "jobs") return modal({ title: "Tasks · Jobs", html: jobsTable(), wide: true });
    const isObj = ["addr", "grp", "svc", "tag"].includes(n), isRule = ["pre", "post", "nat"].includes(n);
    if (a === "add") return isObj ? editObject(n, ctx, null) : n === "nat" ? editNat(ctx, null) : isRule ? editRule(ctx, n, null) : 0;
    if (a === "edit") {
      if (isObj) { const [loc, ...rest] = key.split("|"); return editObject(n, loc, rest.join("|")); }
      return n === "nat" ? editNat(ctx, key) : editRule(ctx, n, key);
    }
    const sel = ui.sel;
    if (!sel) return toast("Sélectionne d'abord une ligne.");
    if (isObj) {
      const [loc, ...rest] = sel.split("|"), name = rest.join("|"), L = st.cfg[loc] && st.cfg[loc][n], o = L && L.find(x => x.name === name);
      if (!o) return toast("Sélectionne d'abord une ligne.");
      if (a === "del") {
        const refs = refsTo(name, loc);
        if (refs.length) return modal({ title: "Delete", html: `<div class="callout bad"><b>Impossible de supprimer '${esc(name)}'</b>L'objet est encore référencé par :<ul>${refs.map(r => `<li>${esc(r)}</li>`).join("")}</ul>Retire d'abord ces références.</div>` });
        L.splice(L.indexOf(o), 1); ui.sel = null; changed(); return toast(`${name} supprimé`);
      }
      if (a === "clone") { const cp = Object.assign(clone(o), { name: (o.name + "-copie").slice(0, 63) }); if (L.some(x => x.name === cp.name)) return toast("Copie déjà existante"); L.push(cp); ui.sel = `${loc}|${cp.name}`; return changed(ui.sel); }
      return;
    }
    if (!isRule) return;
    const L = st.cfg[ctx][n], i = L.findIndex(r => r.name === sel);
    if (i < 0) return toast("Sélectionne d'abord une règle de cette liste.");
    if (a === "del") { L.splice(i, 1); ui.sel = null; changed(); return toast(`Règle ${sel} supprimée`); }
    if (a === "clone") return n === "nat" ? editNat(ctx, null, L[i]) : editRule(ctx, n, null, L[i]);
    if (a === "enable" || a === "disable") { L[i].disabled = a === "disable"; return changed(sel); }
    if (["top", "up", "down", "bottom"].includes(a)) { moveRule(L, i, a); return changed(sel); }
  }
  pz.addEventListener("click", e => {
    const t = e.target.closest("[data-tab],[data-node],[data-act],tr[data-key]");
    if (!t) return;
    if (!t.closest(".pz-cm")) { const m = $(".pz-menu", pz); if (m) m.hidden = true; }
    if (t.dataset.tab) { st.ui.tab = t.dataset.tab; st.ui.node = TREE[t.dataset.tab][0][0]; st.ui.sel = null; save(); return renderPano(); }
    if (t.dataset.node) { st.ui.node = t.dataset.node; st.ui.sel = null; save(); return renderPano(); }
    if (t.dataset.act) return act(t.dataset.act, t.dataset.key);
    st.ui.sel = t.dataset.key; $$("tr.sel", pz).forEach(r => r.classList.remove("sel")); t.classList.add("sel");
  });
  pz.addEventListener("dblclick", e => { const tr = e.target.closest("tr[data-key]"); if (tr) act("edit", tr.dataset.key); });
  pz.addEventListener("change", e => { if (e.target.id === "pzCtx") { st.ui.ctx = e.target.value; st.ui.sel = null; save(); renderPano(); } });
  document.addEventListener("click", e => { if (!e.target.closest(".pz-cm")) { const m = $(".pz-menu", pz); if (m && !m.hidden) m.hidden = true; } });

  /* ---------- commit / push : dialogues ---------- */
  function diffList() {
    const items = cfg => { const m = new Map(); for (const loc of LOCS) for (const k of Object.keys(cfg[loc])) cfg[loc][k].forEach((o, i) => m.set(`${loc} › ${{ addr: "Address", grp: "Address Group", svc: "Service", tag: "Tag", pre: "Security Pre Rule", post: "Security Post Rule", nat: "NAT Rule" }[k]} › ${o.name}`, J(o) + (["pre", "post", "nat"].includes(k) ? "#" + i : ""))); return m; };
    const a = items(st.committed), b = items(st.cfg), out = [];
    b.forEach((v, k) => { if (!a.has(k)) out.push(["add", k]); else if (a.get(k) !== v) out.push(["mod", k]); });
    a.forEach((v, k) => { if (!b.has(k)) out.push(["del", k]); });
    return out;
  }
  const diffHtml = () => { const d = diffList(); return d.length ? `<ul class="df">${d.map(([t, k]) => `<li class="${t}"><b>${{ add: "＋ ajouté", mod: "✎ modifié", del: "－ supprimé" }[t]}</b> ${esc(k)}</li>`).join("")}</ul>` : `<p class="mut">Aucune modification : la configuration candidate est identique à la running config de Panorama.</p>`; };
  function commitAction(a) {
    if (a === "preview") return modal({ title: "Preview Changes · candidate vs running (Panorama)", html: diffHtml(), wide: true });
    if (a === "revert") {
      if (!pending()) return toast("Rien à annuler.");
      return modal({ title: "Revert Changes", html: `<p>Annuler toutes les modifications non commitées et revenir à la dernière configuration commitée sur Panorama ?</p>${diffHtml()}`, okLabel: "Revert", onOk: () => { st.cfg = clone(st.committed); st.ui.sel = null; changed(); toast("Modifications annulées"); } });
    }
    const withCommit = a !== "push", withPush = a !== "commit";
    const dgBoxes = withPush ? `<fieldset class="lb-fs"><legend>Push Scope</legend>${DGS.map(dg => `<label class="ck"><input type="checkbox" name="pdg" value="${dg}"${!inSync(dg) || a === "commitpush" ? " checked" : ""}> <b>${dg}</b> → ${DG[dg].fw} · ${DG[dg].stack} <span class="act ${inSync(dg) ? "al" : "wn"}">${inSync(dg) ? "In sync" : "Out of sync"}</span></label>`).join("")}</fieldset>` : "";
    const warn = a === "push" && pending() ? `<div class="callout warn"><b>Modifications non commitées</b>Push to Devices envoie la configuration <b>commitée</b> sur Panorama. Les changements en attente ne seront pas inclus.</div>` : "";
    modal({ title: { commit: "Commit to Panorama", push: "Push to Devices", commitpush: "Commit and Push" }[a], wide: true, okLabel: { commit: "Commit", push: "Push", commitpush: "Commit and Push" }[a],
      html: `${warn}${withCommit ? `<h4 class="pz-h">Changements à commiter</h4>${diffHtml()}` : ""}${dgBoxes}`,
      onOk: m => {
        const dgs = $$("input[name=pdg]:checked", m).map(x => x.value);
        if (withPush && !dgs.length) return "Choisis au moins un device group.";
        setTimeout(() => runJob({ commit: withCommit, dgs: withPush ? dgs : [] }), 260);
      } });
  }

  async function runJob({ commit, dgs }) {
    const m = modal({ title: "Jobs en cours", cls: "job", wide: true, html: `
      <div class="jm-topo" aria-hidden="true">
        <div class="jm-n" data-n="pano"><b>🛰️</b>Panorama</div>
        <div class="jm-w" data-w="DG-Perimetre"></div><div class="jm-n" data-n="DG-Perimetre"><b>🧱</b>FW-PERIM</div>
        <div class="jm-w" data-w="DG-Interne"></div><div class="jm-n" data-n="DG-Interne"><b>🧱</b>FW-INT</div>
      </div>
      <div class="jm-bar"><i></i></div><ol class="jm-log" aria-live="polite"></ol><div class="jm-res"></div>` });
    const log = $(".jm-log", m), bar = $(".jm-bar i", m), res = $(".jm-res", m);
    const line = async (html, cls = "", d = 220) => { const li = document.createElement("li"); li.className = cls; li.innerHTML = html; log.append(li); log.scrollTop = log.scrollHeight; await sleep(d); };
    const prog = p => { bar.style.width = p + "%"; };
    const steps = (commit ? 1 : 0) + dgs.length; let done = 0, ok = true;
    const node = n => $(`[data-n="${n}"]`, m), wire = n => $(`[data-w="${n}"]`, m);
    if (commit) {
      node("pano").classList.add("busy");
      await line(`<b>Job ${st.job + 1}</b> · Commit · Panorama`, "h");
      await line("Validation de la configuration candidate…", "mut", 500); prog(20);
      const r = doCommit();
      node("pano").classList.remove("busy");
      if (!r.ok) {
        for (const e of r.errors) await line(`Validation Error: ${esc(e)}`, "err", 120);
        await line("Commit failed", "err b"); node("pano").classList.add("bad"); ok = false;
      } else {
        for (const w of r.warnings) await line(`Warning: ${esc(w)}`, "warn", 120);
        await line("Configuration committed successfully", "ok b"); node("pano").classList.add("ok");
      }
      prog(100 * ++done / steps);
    }
    if (ok) for (const dg of dgs) {
      const fw = DG[dg].fw;
      wire(dg).classList.add("flow"); node(dg).classList.add("busy");
      await line(`<b>Job ${st.job + 1}</b> · Push to Devices · ${dg} + ${DG[dg].stack} → ${fw} (${DG[dg].serial})`, "h");
      await line(`${fw} : réception de la configuration, validation…`, "mut", 650);
      const r = doPush(dg);
      wire(dg).classList.remove("flow"); node(dg).classList.remove("busy");
      if (!r.ok) {
        for (const e of r.errors) await line(`${fw} : Validation Error: ${esc(e)}`, "err", 120);
        await line(`${fw} : Commit failed`, "err b"); node(dg).classList.add("bad"); ok = false;
      } else {
        for (const w of r.warnings) await line(`${fw} : Warning: ${esc(w)}`, "warn", 120);
        await line(`${fw} : Configuration committed successfully`, "ok b"); node(dg).classList.add("ok");
      }
      prog(100 * ++done / steps);
    }
    res.innerHTML = ok ? `<div class="jm-ok">✓ ${dgs.length ? "Push terminé" : "Commit terminé"} avec succès</div>` : `<div class="jm-bad">✗ Échec : lis les erreurs ci-dessus, corrige, puis relance.</div>`;
    const dlg = $(".lb-dlg", m);
    dlg.classList.add(ok ? "glow-ok" : "shake");
    if (ok) confetti(res);
    changed();
    return ok;
  }

  /* =========================================================
     MISSIONS & PROGRESSION
     ========================================================= */
  const missions = $("#missions");
  const show = { hint: false, sol: false, solTab: "gui" };
  function renderProgress() {
    const nDone = EX.filter(e => st.done[e.id]).length, pct = Math.round(100 * nDone / EX.length), x = xp(), C = 2 * Math.PI * 26;
    $("#lbProgress").innerHTML = `
      <div class="pg-ring" aria-label="${pct}% des exercices"><svg viewBox="0 0 64 64"><circle cx="32" cy="32" r="26" class="bg"/><circle cx="32" cy="32" r="26" class="fg" style="stroke-dasharray:${C};stroke-dashoffset:${C * (1 - pct / 100)}"/></svg><b>${pct}%</b></div>
      <div><b class="pg-n">${nDone}/${EX.length}</b><span class="mut small">exercices</span></div>
      <div><b class="pg-n">${CLI.filter(x2 => st.cli[x2.id]).length}/${CLI.length}</b><span class="mut small">défis CLI</span></div>
      <div><b class="pg-n">${st.quiz.best || 0}/${QUIZ.length}</b><span class="mut small">meilleur quiz</span></div>
      <div><b class="pg-n grad">${x} XP</b><span class="mut small">rang : <b>${rank(x)}</b></span></div>
      <div class="pg-act"><button type="button" class="btn sm ghost" id="lbReset">↺ Réinitialiser le lab</button></div>`;
  }
  $("#lbProgress").addEventListener("click", e => {
    if (e.target.id !== "lbReset") return;
    modal({ title: "Réinitialiser le lab", html: "<p>Effacer toute la configuration, l'historique des jobs et la progression ?</p>", okLabel: "Tout effacer", onOk: () => { st = fresh(); save(); renderAll(); toast("Lab réinitialisé"); } });
  });
  function renderMissions() {
    const cur = EX.find(e => e.id === st.cur) || EX[0];
    const cr = critOf(cur), done = !!st.done[cur.id], idx = EX.indexOf(cur);
    const where = cur.where.section ? "Packet tracer" : `${{ objects: "Objects", policies: "Policies", panorama: "Panorama" }[cur.where.tab]}${cur.where.ctx ? " · " + cur.where.ctx : ""}`;
    let sol = "";
    if (show.sol && cur.sol) {
      const cli = cur.sol.flatMap(itemCli).join("\n");
      sol = `<div class="ms-sol"><div class="seg2"><button type="button" data-st="gui" class="${show.solTab === "gui" ? "on" : ""}">🖱️ GUI</button><button type="button" data-st="cli" class="${show.solTab === "cli" ? "on" : ""}">⌨️ CLI</button></div>
        ${show.solTab === "gui" ? `<ol class="ms-gui">${cur.sol.map(it => `<li>${itemGui(it)}</li>`).join("")}</ol>` : `<pre class="code ms-cli">${esc(cli)}</pre><button type="button" class="btn sm ghost" data-m="copy">Copier la CLI</button> <span class="small mut">à coller dans le terminal Panorama (mode configure)</span>`}</div>`;
    } else if (show.sol) sol = `<div class="ms-sol"><p class="small">${cur.hint}</p></div>`;
    missions.innerHTML = `
      <div class="ms-levels">${LEVELS.map(L => { const xs = EX.filter(e => e.lvl === L.n), d = xs.filter(e => st.done[e.id]).length;
        return `<details${xs.includes(cur) ? " open" : ""}><summary><span>${L.icon}</span> Niveau ${L.n} · ${L.title}<em class="${d === xs.length ? "full" : ""}">${d}/${xs.length}</em></summary>
          <ul>${xs.map(e => `<li><button type="button" data-ex="${e.id}" class="${e === cur ? "cur" : ""}${st.done[e.id] ? " done" : ""}"><i>${st.done[e.id] ? "✓" : EX.indexOf(e) + 1}</i>${e.title}</button></li>`).join("")}</ul></details>`; }).join("")}</div>
      <article class="ms-card${done ? " is-done" : ""}${st.celebrate === cur.id ? " celebrate" : ""}">
        <div class="row"><span class="badge">Exercice ${idx + 1}/${EX.length}</span><span class="badge or">${esc(where)}</span>${done ? `<span class="badge ok">✓ Réussi${st.assisted[cur.id] ? " (assisté)" : ""}</span>` : ""}</div>
        <h3>${cur.title}</h3>
        <p class="ms-brief">${cur.brief}</p>
        ${cur.setup ? `<button type="button" class="btn sm ${st.setups[cur.setup] ? "ghost" : "primary"}" data-m="setup">⚙️ ${st.setups[cur.setup] ? "Recharger" : "Charger"} le scénario</button>` : ""}
        <ul class="ms-crit">${cr.map(x => `<li class="${x.ok ? "ok" : "ko"}${x.sub ? " sub" : ""}"><i aria-hidden="true">${x.ok ? "✓" : "○"}</i><span>${esc(x.label)}${!x.ok && x.hint ? `<small>${esc(x.hint)}</small>` : ""}</span></li>`).join("")}</ul>
        ${done ? `<div class="callout ok ms-learn"><b>📘 À retenir</b>${cur.learn}</div>` : ""}
        ${show.hint && !show.sol ? `<div class="callout ms-hint"><b>💡 Indice</b>${cur.hint}</div>` : ""}
        <div class="ms-btns">
          <button type="button" class="btn sm" data-m="go">🎯 Aller à l'écran</button>
          <button type="button" class="btn sm ghost" data-m="hint">💡 Indice</button>
          <button type="button" class="btn sm ghost" data-m="sol">🧩 Solution</button>
          ${cur.sol && !done ? `<button type="button" class="btn sm ghost" data-m="apply" title="Remplit la configuration à ta place (30 XP au lieu de 100)">⚡ Appliquer</button>` : ""}
        </div>
        ${sol}
        <div class="ms-nav"><button type="button" class="btn sm ghost" data-m="prev"${idx ? "" : " disabled"}>← Précédent</button><button type="button" class="btn sm ${done ? "primary" : "ghost"}" data-m="next"${idx < EX.length - 1 ? "" : " disabled"}>Suivant →</button></div>
      </article>`;
    if (st.celebrate === cur.id) { confetti($(".ms-card", missions)); st.celebrate = null; save(); }
  }
  function selectEx(id) { st.cur = id; show.hint = show.sol = false; st.celebrate = null; save(); renderMissions(); }
  function goWhere(w) {
    if (w.section) return document.getElementById(w.section).scrollIntoView({ behavior: reduce ? "auto" : "smooth" });
    st.ui.tab = w.tab; st.ui.node = w.node; if (w.ctx) st.ui.ctx = w.ctx; st.ui.sel = null; save(); renderPano();
    pz.scrollIntoView({ behavior: reduce ? "auto" : "smooth", block: "start" });
    pz.classList.remove("pulse-hl"); void pz.offsetWidth; pz.classList.add("pulse-hl");
  }
  missions.addEventListener("click", e => {
    const ex = e.target.closest("[data-ex]"); if (ex) return selectEx(ex.dataset.ex);
    const b = e.target.closest("[data-m],[data-st]"); if (!b) return;
    const cur = EX.find(x => x.id === st.cur), i = EX.indexOf(cur);
    if (b.dataset.st) { show.solTab = b.dataset.st; return renderMissions(); }
    switch (b.dataset.m) {
      case "go": return goWhere(cur.where);
      case "hint": show.hint = !show.hint; show.sol = false; return renderMissions();
      case "sol": show.sol = !show.sol; return renderMissions();
      case "copy": return window.PAN && PAN.copy ? PAN.copy(cur.sol.flatMap(itemCli).join("\n")) : 0;
      case "apply": st.assisted[cur.id] = true; applySol(cur); changed(); goWhere(cur.where); return toast("Solution appliquée dans la config candidate (pense au commit !)");
      case "setup": SETUP[cur.setup](); st.setups[cur.setup] = Date.now(); changed(); goWhere(cur.where); return toast("Scénario chargé : lance un Commit pour voir le problème");
      case "prev": return selectEx(EX[Math.max(0, i - 1)].id);
      case "next": return selectEx(EX[Math.min(EX.length - 1, i + 1)].id);
    }
  });

  /* =========================================================
     TOPOLOGIE SVG (héros + packet tracer)
     ========================================================= */
  const NODES = {
    inet: { x: 80, y: 190, ic: "🌐", t: "Internet", s: "Untrust", c: "#EF4444" },
    perim: { x: 330, y: 190, ic: "🧱", t: "FW-PERIM", s: "PA-3410", fw: 1 },
    dmz: { x: 330, y: 330, ic: "🖥️", t: "DMZ", s: "172.16.10.0/24", c: "#F59E0B" },
    int: { x: 600, y: 190, ic: "🧱", t: "FW-INT", s: "PA-1410", fw: 1 },
    users: { x: 880, y: 70, ic: "💻", t: "Users", s: "10.10.10.0/24", c: "#38BDF8" },
    srv: { x: 880, y: 190, ic: "🗄️", t: "Servers", s: "10.20.20.0/24", c: "#22C55E" },
    adm: { x: 880, y: 310, ic: "🔑", t: "Admin", s: "10.30.30.0/24", c: "#A78BFA" },
    pano: { x: 465, y: 48, ic: "🛰️", t: "Panorama", s: "M-200", pano: 1 },
  };
  const LINKS = [["inet", "perim"], ["perim", "dmz"], ["perim", "int"], ["int", "users"], ["int", "srv"], ["int", "adm"]];
  const pts = keys => keys.map(k => `${NODES[k].x},${NODES[k].y}`).join(" L");
  function topoSvg(id, ambient) {
    const nodes = Object.entries(NODES).map(([k, n]) => { const w = n.fw ? 164 : 152, x0 = -w / 2;
      return `<g class="tn${n.fw ? " fw" : ""}${n.pano ? " pano" : ""}" id="${id}-${k}" transform="translate(${n.x},${n.y})" style="--nc:${n.c || "var(--or)"}">
      <rect x="${x0}" y="-30" width="${w}" height="60" rx="14"/><text class="ic" x="${x0 + 10}" y="9">${n.ic}</text>
      <text class="t" x="${x0 + 44}" y="-3">${n.t}</text><text class="s" x="${x0 + 44}" y="15">${n.s}</text></g>`; }).join("");
    const links = LINKS.map(([a, b]) => `<path class="tl" d="M${NODES[a].x},${NODES[a].y} L${NODES[b].x},${NODES[b].y}"/>`).join("");
    const mg = ["perim", "int"].map(k => `<path class="tm" d="M${NODES.pano.x},${NODES.pano.y} L${NODES[k].x},${NODES[k].y}"/>`).join("");
    const amb = ambient ? [
      ["inet", "perim", "dmz"], ["users", "int", "perim", "inet"], ["inet", "perim", "int", "srv"], ["adm", "int", "srv"],
    ].map((p, i) => {
      const blocked = i === 2, d = `M${pts(blocked ? ["inet", "perim"] : p)}`;
      return `<circle r="6" class="pk ${blocked ? "bad" : ["ok", "ok2", "", "ok3"][i]}"><animateMotion dur="${blocked ? 2.4 : 3.2 + i * .4}s" begin="${i * .8}s" repeatCount="indefinite" path="${d}"/></circle>`;
    }).join("") : "";
    return `<svg viewBox="0 0 1000 380" class="topo" role="img" aria-label="Topologie double barrière : Internet, FW-PERIM, DMZ, Transit, FW-INT, Users, Servers, Admin, Panorama">
      <defs><filter id="${id}-gl"><feGaussianBlur stdDeviation="3.5" result="b"/><feMerge><feMergeNode in="b"/><feMergeNode in="SourceGraphic"/></feMerge></filter></defs>
      ${mg}${links}<text class="tlab" x="465" y="182">Transit</text><text class="tlab" x="465" y="242">10.255.0.0/30</text><text class="tlab" x="205" y="182">Untrust</text><text class="tlab" x="205" y="242">203.0.113.0/24</text>
      <g filter="url(#${id}-gl)">${amb}</g>${nodes}<g class="fx"></g></svg>`;
  }
  $("#heroTopo").innerHTML = topoSvg("ht", !reduce);

  /* =========================================================
     PACKET TRACER
     ========================================================= */
  const trTopo = $("#trTopo"), trLog = $("#trLog");
  trTopo.innerHTML = topoSvg("tt", false);
  $("#trApp").innerHTML = opt(Object.keys(APPS), "ssl");
  $("#trFlows").innerHTML = FLOWS.map(f => `<button type="button" class="tr-f" data-f="${f.id}"><b>${f.label}</b><span class="mono">${f.src} → ${f.dst} · ${f.app}</span>
    <em class="${f.expect === "allow" ? "al" : "dn"}">attendu : ${f.expect === "allow" ? "autorisé" : "bloqué"}${f.bypass ? " · FW-PERIM compromis" : ""}</em><i class="st"></i></button>`).join("");
  function renderTraceSync() {
    const out = DGS.filter(d => !inSync(d) || pending());
    $("#trSync").innerHTML = out.length ? `<div class="callout warn small"><b>Config non poussée</b>Le traceur utilise la dernière config <b>poussée</b>. ${pending() ? "Tu as des modifications non commitées. " : ""}Fais un Commit and Push pour tester tes derniers changements.</div>` : `<div class="callout ok small"><b>Pare-feux In sync</b>Le traceur utilise exactement ce que tu as poussé.</div>`;
    $$(".tr-f", $("#trFlows")).forEach(b => b.classList.toggle("passed", !!st.traces[b.dataset.f]));
  }
  const whereIp = ip => (inNet(ip, "10.10.10.0/24") ? "users" : inNet(ip, "10.20.20.0/24") ? "srv" : inNet(ip, "10.30.30.0/24") ? "adm" : inNet(ip, "172.16.10.0/24") ? "dmz" : "inet");
  const inside = n => ["users", "srv", "adm"].includes(n);
  let tracing = false;
  function movePacket(svg, a, b, cls) {
    return new Promise(res => {
      const A = NODES[a], B = NODES[b], fx = $(".fx", svg);
      const p = document.createElementNS("http://www.w3.org/2000/svg", "circle");
      p.setAttribute("r", 9); p.setAttribute("class", "pk " + cls); fx.append(p);
      const T = reduce ? 1 : 700, t0 = performance.now();
      const step = t => { const k = Math.min(1, (t - t0) / T), e = k < .5 ? 2 * k * k : 1 - (-2 * k + 2) ** 2 / 2;
        p.setAttribute("cx", A.x + (B.x - A.x) * e); p.setAttribute("cy", A.y + (B.y - A.y) * e);
        if (k < 1) requestAnimationFrame(step); else { p.remove(); res(); } };
      requestAnimationFrame(step);
    });
  }
  function nodeFx(svg, k, cls, label) {
    const g = $(`#tt-${k}`, svg); g.classList.remove("ok", "bad", "warn"); void g.getBBox(); g.classList.add(cls);
    if (label) {
      const n = NODES[k], t = document.createElementNS("http://www.w3.org/2000/svg", "text");
      t.setAttribute("x", n.x); t.setAttribute("y", n.y + (k === "perim" ? -44 : 50)); t.setAttribute("class", "bubble " + cls); t.setAttribute("text-anchor", "middle"); t.textContent = label;
      $(".fx", svg).append(t);
    }
  }
  const logLi = (html, cls = "") => { const li = document.createElement("li"); li.className = cls; li.innerHTML = html; trLog.append(li); };
  async function trace(flow, bypass) {
    if (tracing) return; tracing = true;
    const svg = $("svg", trTopo);
    $(".fx", svg).innerHTML = ""; $$(".tn", svg).forEach(g => g.classList.remove("ok", "bad", "warn"));
    trLog.innerHTML = "";
    let src = flow.src, dst = flow.dst;
    const app = flow.app, [proto, port] = APPS[app];
    const sN = whereIp(src), dN = whereIp(dst);
    logLi(`<b>${esc(src)} → ${esc(dst)}</b> · ${app} (${proto}${proto === "icmp" ? "" : "/" + port})${bypass ? ` · <span class="w">FW-PERIM compromis</span>` : ""}`, "h");
    if (sN === "inet" && dN === "inet" && dst !== "203.0.113.20") { logLi("Flux Internet → Internet : il ne traverse pas le lab.", "warn"); tracing = false; return null; }
    if (sN === dN && inside(sN)) { logLi("Même zone, même réseau : le flux ne traverse aucun pare-feu.", "warn"); tracing = false; return null; }
    const hops = inside(sN) ? (inside(dN) ? ["DG-Interne"] : ["DG-Interne", "DG-Perimetre"]) : (inside(dN) ? ["DG-Perimetre", "DG-Interne"] : ["DG-Perimetre"]);
    let pos = sN, verdict = "allow", stopAt = null;
    for (const dg of hops) {
      const fwN = DG[dg].node;
      await movePacket(svg, pos, fwN, "ok"); pos = fwN;
      if (bypass && dg === "DG-Perimetre") { nodeFx(svg, fwN, "warn", "bypass"); logLi(`<b>FW-PERIM</b> · compromis : le paquet passe sans contrôle`, "warn"); await sleep(350); continue; }
      const cfg = st.pushed[dg];
      const from = zoneOf(dg, src), toPre = zoneOf(dg, dst);
      const nat = evalNat(cfg, dg, { from, to: toPre, src, dst, app, proto, port });
      let dstPost = dst;
      if (nat && nat.dnat) { const a = lookup(cfg, dg, "addr", nat.dnat); if (a) dstPost = a.value.split(/[/-]/)[0]; }
      if (dg === "DG-Perimetre" && dst === "203.0.113.20" && dstPost === dst) {
        nodeFx(svg, fwN, "bad", "pas de DNAT"); logLi(`<b>FW-PERIM</b> · aucune règle NAT ne traduit 203.0.113.20 : paquet perdu`, "err");
        verdict = "block"; stopAt = dg; break;
      }
      const to = zoneOf(dg, dstPost);
      if (nat) logLi(`<b>${DG[dg].fw}</b> · NAT <code>${esc(nat.name)}</code>${nat.dnat ? ` · destination ${dst} → ${dstPost}` : ""}${nat.snat.type === "dipp" ? ` · source ${src} → ${DG[dg].ifaces[nat.snat.iface] ? DG[dg].ifaces[nat.snat.iface][1].split("/")[0] : "?"}` : ""}`, "mut");
      const m = evalPolicy(cfg, dg, { from, to, src, dst, app, proto, port });
      const allowed = m.action === "allow";
      logLi(`<b>${DG[dg].fw}</b> · <span class="zt ${ZCLS[from]}">${from}</span> → <span class="zt ${ZCLS[to]}">${to}</span> · règle <code>${esc(m.name)}</code> ${m.base !== "default" ? `<small>(${m.base === "pre" ? "pre" : "post"} rule, index ${m.index})</small>` : "<small>(règle par défaut)</small>"} · <b class="${allowed ? "g" : "r"}">${m.action}</b>`, allowed ? "ok" : "err");
      nodeFx(svg, fwN, allowed ? "ok" : "bad", (allowed ? "✓ " : "✕ ") + m.name);
      await sleep(450);
      if (!allowed) { verdict = "block"; stopAt = dg; break; }
      if (nat && nat.snat.type === "dipp" && DG[dg].ifaces[nat.snat.iface]) src = DG[dg].ifaces[nat.snat.iface][1].split("/")[0];
      dst = dstPost;
    }
    if (verdict === "allow") {
      const fin = whereIp(dst);
      await movePacket(svg, pos, fin, "ok");
      if (fin === "inet" && isPrivate(src)) { verdict = "noreturn"; nodeFx(svg, fin, "warn", "pas de retour"); logLi(`Le paquet sort avec l'IP privée <b>${src}</b> : sans SNAT, la réponse d'Internet ne revient jamais.`, "warn"); }
      else { nodeFx(svg, fin, "ok", "✓ reçu"); logLi(`Livré à <b>${esc(dst)}</b> (${NODES[fin].t})`, "ok"); }
    } else {
      const fx = $(".fx", svg), n = NODES[pos];
      fx.insertAdjacentHTML("beforeend", `<circle class="burst" cx="${n.x}" cy="${n.y}" r="12"/>`);
    }
    const allowedEnd = verdict === "allow";
    const okExp = flow.expect ? (flow.expect === "allow" ? allowedEnd : verdict === "block" && (!flow.stop || stopAt === flow.stop)) && (!flow.bypass || bypass) : null;
    if (flow.expect) {
      logLi(okExp ? `🎯 Résultat conforme à la recette (${flow.expect === "allow" ? "autorisé" : "bloqué"}).` : `✗ Résultat non conforme : attendu ${flow.expect === "allow" ? "autorisé" : "bloqué"}${flow.stop ? ` par ${DG[flow.stop].fw}` : ""}.`, okExp ? "ok b" : "err b");
      if (flow.id) { st.traces[flow.id] = !!okExp; changed(); }
    }
    tracing = false;
    return verdict;
  }
  $("#trFlows").addEventListener("click", e => {
    const b = e.target.closest("[data-f]"); if (!b) return;
    const f = FLOWS.find(x => x.id === b.dataset.f);
    $("#trSrc").value = f.src; $("#trDst").value = f.dst; $("#trApp").value = f.app;
    $("#trBypass").checked = !!f.bypass;
    trace(f, $("#trBypass").checked);
  });
  $("#trForm").addEventListener("submit", e => {
    e.preventDefault();
    const s = $("#trSrc").value.trim(), d = $("#trDst").value.trim();
    if (!validIp(s) || !validIp(d)) return toast("Adresse IP invalide");
    const preset = FLOWS.find(f => f.src === s && f.dst === d && f.app === $("#trApp").value);
    trace(preset || { src: s, dst: d, app: $("#trApp").value }, $("#trBypass").checked);
  });

  /* =========================================================
     TERMINAL CLI
     ========================================================= */
  const term = { target: "pano", conf: false, hist: [], hi: 0 };
  const tOut = $("#termOut"), tIn = $("#termIn");
  const host = () => (term.target === "pano" ? "Panorama" : DG[term.target].fw + "(active)");
  const prompt = () => `admin@${host()}${term.conf ? "#" : ">"} `;
  const out = (txt, cls = "") => { const d = document.createElement("div"); d.className = "tl " + cls; d.textContent = txt; tOut.append(d); tOut.scrollTop = tOut.scrollHeight; };
  const updPrompt = () => { $("#termPrompt").textContent = prompt(); };
  function tokenize(line) {
    const toks = [], stack = [toks], re = /"([^"]*)"|\[|\]|[^\s[\]"]+/g; let m;
    while ((m = re.exec(line))) {
      const cur = stack[stack.length - 1];
      if (m[0] === "[") { const a = []; cur.push(a); stack.push(a); }
      else if (m[0] === "]") { if (stack.length > 1) stack.pop(); }
      else cur.push(m[1] !== undefined ? m[1] : m[0]);
    }
    return toks;
  }
  const arr = v => (Array.isArray(v) ? v : [v]);
  function cliSet(t) {
    let loc, i;
    if (t[1] === "shared") { loc = "Shared"; i = 2; } else if (t[1] === "device-group") { loc = t[2]; i = 3; if (!DG[loc]) return `Invalid syntax : device group '${loc}' inconnu (${DGS.join(", ")}).`; } else return "Invalid syntax : set shared … | set device-group <nom> …";
    const what = t[i], name = t[i + 1];
    const kmap = { address: "addr", "address-group": "grp", service: "svc", tag: "tag" };
    if (kmap[what]) {
      if (!name || nameErr(name)) return "Invalid syntax : nom d'objet invalide.";
      const k = kmap[what], L = st.cfg[loc][k];
      if ((k === "addr" && st.cfg[loc].grp.some(x => x.name === name)) || (k === "grp" && st.cfg[loc].addr.some(x => x.name === name))) return `'${name}' existe déjà dans ${loc} avec un autre type.`;
      let o = L.find(x => x.name === name); const isNew = !o;
      o = o ? clone(o) : { addr: { name, type: "ip-netmask", value: "", tags: [], desc: "" }, grp: { name, members: [], desc: "" }, svc: { name, proto: "tcp", port: "", desc: "" }, tag: { name, color: "color1" } }[k];
      for (let j = i + 2; j < t.length; j++) {
        const a = t[j], v = t[j + 1];
        if (["ip-netmask", "ip-range", "fqdn"].includes(a) && k === "addr") { o.type = a; o.value = String(v); j++; }
        else if (a === "description") { o.desc = String(v); j++; }
        else if (a === "tag" && k === "addr") { o.tags = arr(v); j++; }
        else if (a === "static" && k === "grp") { o.members = arr(v); j++; }
        else if (a === "protocol" && k === "svc") { o.proto = String(v); j++; if (t[j + 1] === "port") { o.port = String(t[j + 2]); j += 2; } }
        else if (a === "port" && k === "svc") { o.port = String(v); j++; }
        else if (a === "color" && k === "tag") { if (!COLORS[v]) return `Invalid syntax : couleur ${v} (color1…color8).`; o.color = v; j++; }
        else return `Invalid syntax : '${a}' inattendu.`;
      }
      if (k === "addr" && o.type === "ip-netmask" && !validNet(o.value)) return "Invalid syntax : ip-netmask invalide.";
      if (k === "svc" && o.port && !validPorts(o.port)) return "Invalid syntax : port invalide.";
      if (isNew) L.push(o); else L[L.findIndex(x => x.name === name)] = o;
      return null;
    }
    if (what === "pre-rulebase" || what === "post-rulebase") {
      if (loc === "Shared") return "Ce lab ne gère que les rulebases des device groups.";
      const base = what === "pre-rulebase" ? "pre" : "post", kind = t[i + 1], rn = t[i + 3];
      if (t[i + 2] !== "rules" || !rn || nameErr(rn)) return "Invalid syntax : … security rules <nom> …";
      if (kind === "security") {
        const L = st.cfg[loc][base]; let r = L.find(x => x.name === rn); const isNew = !r;
        if (isNew && [...st.cfg[loc].pre, ...st.cfg[loc].post].some(x => x.name === rn)) return `Une règle '${rn}' existe déjà dans l'autre rulebase de ${loc}.`;
        r = r ? clone(r) : blankRule(); r.name = rn;
        const F = { from: "from", to: "to", source: "src", destination: "dst", application: "app", service: "svc", tag: "tags" };
        for (let j = i + 4; j < t.length; j++) {
          const a = t[j], v = t[j + 1];
          if (F[a]) { r[F[a]] = arr(v); j++; }
          else if (a === "action") { r.action = String(v); j++; }
          else if (a === "description") { r.desc = String(v); j++; }
          else if (a === "profile-setting" && v === "group") { r.profile = String(t[j + 2]); j += 2; }
          else if (a === "log-setting") { r.lf = String(v); j++; }
          else if (a === "log-end") { r.logEnd = v === "yes"; j++; }
          else if (a === "disabled") { r.disabled = v === "yes"; j++; }
          else return `Invalid syntax : '${a}' inattendu.`;
        }
        if (isNew) L.push(r); else L[L.findIndex(x => x.name === rn)] = r;
        return null;
      }
      if (kind === "nat") {
        if (base !== "pre") return "Ce lab utilise les NAT Pre Rules.";
        const L = st.cfg[loc].nat; let r = L.find(x => x.name === rn); const isNew = !r;
        r = r ? clone(r) : blankNat(); r.name = rn;
        for (let j = i + 4; j < t.length; j++) {
          const a = t[j], v = t[j + 1];
          if (a === "from") { r.from = arr(v); j++; }
          else if (a === "to") { r.to = String(arr(v)[0]); j++; }
          else if (a === "source") { r.src = arr(v); j++; }
          else if (a === "destination") { r.dst = arr(v); j++; }
          else if (a === "service") { r.svc = String(v); j++; }
          else if (a === "description") { r.desc = String(v); j++; }
          else if (a === "source-translation" && v === "dynamic-ip-and-port") { const k2 = t.indexOf("interface", j); if (k2 < 0) return "Invalid syntax : interface-address interface <ethernetX/Y>"; r.snat = { type: "dipp", iface: String(t[k2 + 1]) }; j = k2 + 1; }
          else if (a === "destination-translation" && v === "translated-address") { r.dnat = String(t[j + 2]); j += 2; }
          else return `Invalid syntax : '${a}' inattendu.`;
        }
        if (isNew) L.push(r); else L[L.findIndex(x => x.name === rn)] = r;
        return null;
      }
    }
    return "Invalid syntax : address | address-group | service | tag | pre-rulebase | post-rulebase.";
  }
  function cliDelete(t) {
    let loc, i;
    if (t[1] === "shared") { loc = "Shared"; i = 2; } else if (t[1] === "device-group" && DG[t[2]]) { loc = t[2]; i = 3; } else return "Invalid syntax.";
    const kmap = { address: "addr", "address-group": "grp", service: "svc", tag: "tag" };
    if (kmap[t[i]]) {
      const L = st.cfg[loc][kmap[t[i]]], n = t[i + 1], j = L.findIndex(x => x.name === n);
      if (j < 0) return `'${n}' introuvable dans ${loc}.`;
      const refs = refsTo(n, loc); if (refs.length) return `Impossible de supprimer '${n}' : référencé par ${refs.join(", ")}.`;
      L.splice(j, 1); return null;
    }
    if (/^(pre|post)-rulebase$/.test(t[i]) && t[i + 2] === "rules") {
      const L = st.cfg[loc][t[i + 1] === "nat" ? "nat" : t[i].slice(0, -9)], j = L ? L.findIndex(x => x.name === t[i + 3]) : -1;
      if (j < 0) return `Règle '${t[i + 3]}' introuvable.`;
      L.splice(j, 1); return null;
    }
    return "Invalid syntax.";
  }
  function cliMove(t) {
    if (t[1] !== "device-group" || !DG[t[2]] || !/^(pre|post)-rulebase$/.test(t[3]) || t[5] !== "rules") return "Invalid syntax : move device-group <dg> pre-rulebase security rules <nom> top|bottom|before <r>|after <r>";
    const L = st.cfg[t[2]][t[4] === "nat" ? "nat" : t[3].slice(0, -9)], i = L.findIndex(x => x.name === t[6]);
    if (i < 0) return `Règle '${t[6]}' introuvable.`;
    const w = t[7];
    if (w === "top" || w === "bottom") { moveRule(L, i, w); return null; }
    if (w === "before" || w === "after") {
      const [r] = L.splice(i, 1), j = L.findIndex(x => x.name === t[8]);
      if (j < 0) { L.splice(i, 0, r); return `Règle '${t[8]}' introuvable.`; }
      L.splice(w === "before" ? j : j + 1, 0, r); return null;
    }
    return "Invalid syntax : top | bottom | before <règle> | after <règle>";
  }
  const pad = (s, n) => String(s).padEnd(n);
  function showRules(cfg, dg) {
    const L = [...cfg[dg].pre.map(r => [r, "pre"]), ...cfg[dg].post.map(r => [r, "post"])];
    if (!L.length) return ["(aucune règle Panorama : seules les règles par défaut s'appliquent)"];
    return L.flatMap(([r, b], i) => [`"${r.name}; index: ${i + 1}" {  # ${b}-rulebase (Panorama)${r.disabled ? " DISABLED" : ""}`, `        from ${fmt(r.from)};`, `        source ${fmt(r.src)};`, `        to ${fmt(r.to)};`, `        destination ${fmt(r.dst)};`, `        application/service ${fmt(r.app)} / ${fmt(r.svc)};`, `        action ${r.action};`, `        terminal yes;`, `}`]);
  }
  function markCli(id) { if (!st.cli[id]) { st.cli[id] = Date.now(); toast(`⌨️ Défi CLI réussi : ${CLI.find(c => c.id === id).title} (+50 XP)`); save(); renderCli(); renderProgress(); } }
  async function exec(line) {
    out(prompt() + line, "cmd");
    const raw = line.trim(); if (!raw) return;
    const t = tokenize(raw), w = t.slice(0, 3).join(" ");
    const T = term.target, isPano = T === "pano";
    if (raw === "clear") { tOut.innerHTML = ""; return; }
    if (raw === "help" || raw === "?") {
      if (term.conf) return ["Mode configuration :", "  set shared|device-group <dg> address <n> ip-netmask <cidr> [tag [ … ]] [description \"…\"]", "  set … address-group <n> static [ a b ]", "  set … service <n> protocol tcp|udp port <p>", "  set … tag <n> color color1…color8", "  set device-group <dg> pre-rulebase|post-rulebase security rules <n> from … to … source … destination … application … service … action allow|deny|drop [profile-setting group <g>] [log-setting <lf>]", "  set device-group <dg> pre-rulebase nat rules <n> … [source-translation dynamic-ip-and-port interface-address interface <if>] [destination-translation translated-address <obj>]", "  move device-group <dg> pre-rulebase security rules <n> top|bottom|before <r>|after <r>", "  delete …   show   commit   exit"].forEach(l => out(l, "mut"));
      const L = isPano ? ["show system info", "show devices connected", "show devicegroups", "show jobs all", "commit", "commit-all shared-policy device-group <dg>", "configure", "clear"] : ["show system info", "show interface all", "show running security-policy", "show running nat-policy", "show jobs all", "test security-policy-match from <z> to <z> source <ip> destination <ip> protocol <6|17|1> destination-port <p> [application <app>]", "clear"];
      return L.forEach(l => out("  " + l, "mut"));
    }
    if (term.conf) {
      if (raw === "exit" || raw === "quit") { term.conf = false; updPrompt(); return; }
      if (raw === "show") { const l = cfgToCli(st.cfg); (l.length ? l : ["# configuration candidate vide"]).forEach(x => out(x, "mut")); return; }
      if (raw === "commit") return cliCommit();
      const err = t[0] === "set" ? cliSet(t) : t[0] === "delete" ? cliDelete(t) : t[0] === "move" ? cliMove(t) : "Unknown command : " + t[0];
      if (err) return out(err, "err");
      const a = st.cfg.Shared.addr.find(x => x.name === "H-NTP");
      if (a && a.type === "ip-netmask" && normNet(a.value) === "10.20.20.123/32") markCli("c-addr");
      changed();
      return;
    }
    if (raw === "configure") { if (!isPano) return out("Ce pare-feu est géré par Panorama : configure-le depuis Panorama dans ce lab.", "warn"); term.conf = true; updPrompt(); return out("Entering configuration mode", "mut"); }
    if (raw === "show system info") {
      if (isPano) ["hostname: Panorama", "ip-address: 192.168.100.10", "model: M-200", "serial: 0007PANO4821", "sw-version: 11.1.4", "app-version: 8845-8956", "system-mode: panorama", "operational-mode: normal"].forEach(l => out(l));
      else { const d = DG[T]; [`hostname: ${d.fw}`, `ip-address: ${d.ip}`, `model: ${d.model}`, `serial: ${d.serial}`, "sw-version: 11.0.4-h1", "app-version: 8845-8956", "threat-version: 8845-8956", "device-group: " + T, "template-stack: " + d.stack, "operational-mode: normal"].forEach(l => out(l)); if (T === "DG-Interne") markCli("c-info"); }
      return;
    }
    if (isPano && raw === "show devices connected") {
      out(`${pad("Name", 11)}${pad("Serial", 15)}${pad("Address", 17)}${pad("Model", 9)}${pad("Device Group", 15)}${pad("Connected", 11)}Policy`);
      out("-".repeat(86), "mut");
      DGS.forEach(dg => out(`${pad(DG[dg].fw, 11)}${pad(DG[dg].serial, 15)}${pad(DG[dg].ip, 17)}${pad(DG[dg].model, 9)}${pad(dg, 15)}${pad("yes", 11)}${inSync(dg) ? "In sync" : "Out of sync"}`));
      return markCli("c-dev");
    }
    if (isPano && raw === "show devicegroups") return DGS.forEach(dg => { out(`Group: ${dg}   Shared policy md5sum: ${inSync(dg) ? "in sync" : "out of sync"}`); out(`  ${DG[dg].fw} (${DG[dg].serial})  Connected: yes`, "mut"); });
    if (raw === "show jobs all") {
      const J2 = isPano ? st.jobs : st.jobs.filter(j => j.type === "CommitAll" && j.dev.includes(DG[T].fw));
      out(`${pad("ID", 7)}${pad("Type", 11)}${pad("Status", 8)}${pad("Result", 8)}Completed`); out("-".repeat(50), "mut");
      if (!J2.length) return out("(aucun job)", "mut");
      return J2.forEach(j => out(`${pad(j.id, 7)}${pad(j.type, 11)}${pad("FIN", 8)}${pad(j.ok ? "OK" : "FAIL", 8)}${j.t}`, j.ok ? "" : "err"));
    }
    if (isPano && raw === "commit") return cliCommit();
    if (isPano && t[0] === "commit-all") {
      const dg = t[t.indexOf("device-group") + 1];
      if (t[1] !== "shared-policy" || !DG[dg]) return out("Invalid syntax : commit-all shared-policy device-group <DG-Perimetre|DG-Interne>", "err");
      if (pending()) out("Attention : des modifications ne sont pas commitées sur Panorama et ne seront pas poussées.", "warn");
      out(`Job enqueued with jobid ${st.job + 1}`, "mut"); await sleep(600);
      const r = doPush(dg);
      r.errors.forEach(e => out(`${DG[dg].fw}: Validation Error: ${e}`, "err"));
      r.warnings.forEach(wn => out(`${DG[dg].fw}: Warning: ${wn}`, "warn"));
      out(r.ok ? `${DG[dg].fw}: Configuration committed successfully` : `${DG[dg].fw}: Commit failed`, r.ok ? "ok" : "err");
      if (r.ok && dg === "DG-Interne") markCli("c-push");
      changed(); return;
    }
    if (!isPano && raw === "show interface all") {
      out(`${pad("name", 14)}${pad("zone", 10)}${pad("vsys", 6)}ip`); out("-".repeat(50), "mut");
      Object.entries(DG[T].ifaces).forEach(([i, [z, a]]) => out(`${pad(i, 14)}${pad(z, 10)}${pad("1", 6)}${a}`)); out(`${pad("management", 14)}${pad("-", 10)}${pad("-", 6)}${DG[T].ip}/24`); return;
    }
    if (!isPano && w === "show running security-policy") { showRules(st.pushed[T], T).forEach(l => out(l)); if (T === "DG-Perimetre") markCli("c-run"); return; }
    if (!isPano && w === "show running nat-policy") { const L = st.pushed[T][T].nat; if (!L.length) return out("(aucune règle NAT)", "mut"); return L.forEach(r => out(`"${r.name}" { from ${fmt(r.from)}; to ${r.to}; source ${fmt(r.src)}; destination ${fmt(r.dst)}; service ${r.svc};${r.snat.type === "dipp" ? ` translate-to "src: ${r.snat.iface} (dynamic-ip-and-port)";` : ""}${r.dnat ? ` translate-to "dst: ${r.dnat}";` : ""} }`)); }
    if (!isPano && t[0] === "test" && t[1] === "security-policy-match") {
      const g = k => { const i = t.indexOf(k); return i > 0 ? String(t[i + 1]) : ""; };
      const f = { from: g("from"), to: g("to"), src: g("source"), dst: g("destination"), app: g("application") || undefined, proto: { 6: "tcp", 17: "udp", 1: "icmp" }[g("protocol")], port: +g("destination-port") };
      if (!f.from || !f.to || !validIp(f.src) || !validIp(f.dst) || !f.proto) return out("Invalid syntax : from <zone> to <zone> source <ip> destination <ip> protocol <6|17|1> destination-port <port> [application <app>]", "err");
      if (!DG[T].zones.includes(f.from) || !DG[T].zones.includes(f.to)) return out(`Zone inconnue sur ${DG[T].fw} (zones : ${DG[T].zones.join(", ")}).`, "err");
      const m = evalPolicy(st.pushed[T], T, f);
      out(`"${m.name}; index: ${m.index}" {`); if (m.r) { out(`        from ${fmt(m.r.from)};`); out(`        source ${fmt(m.r.src)};`); out(`        to ${fmt(m.r.to)};`); out(`        destination ${fmt(m.r.dst)};`); out(`        application/service ${fmt(m.r.app)} / ${fmt(m.r.svc)};`); }
      out(`        action ${m.action};`, m.action === "allow" ? "ok" : "err"); out("        terminal yes;"); out("}");
      if (T === "DG-Interne" && m.r && m.action === "allow" && f.dst === "10.20.20.53" && f.port === 53) markCli("c-tspm");
      return;
    }
    out(`Unknown command: ${t[0] || raw}. Tape « help ».`, "err");
  }
  async function cliCommit() {
    out(`Job enqueued with jobid ${st.job + 1}`, "mut"); await sleep(500);
    const r = doCommit();
    r.errors.forEach(e => out("Validation Error: " + e, "err"));
    r.warnings.forEach(wn => out("Warning: " + wn, "warn"));
    out(r.ok ? "Configuration committed successfully" : "Commit failed", r.ok ? "ok" : "err");
    if (r.ok) markCli("c-commit");
    changed();
  }
  const COMPL = ["show system info", "show devices connected", "show devicegroups", "show jobs all", "show interface all", "show running security-policy", "show running nat-policy", "test security-policy-match from ", "commit", "commit-all shared-policy device-group ", "configure", "clear", "help", "exit",
    "set shared address ", "set shared address-group ", "set shared service ", "set shared tag ", "set device-group DG-Perimetre ", "set device-group DG-Interne ", "move device-group ", "delete "];
  $("#termForm").addEventListener("submit", e => {
    e.preventDefault();
    const v = tIn.value;
    if (v.trim()) { term.hist.push(v); term.hi = term.hist.length; }
    tIn.value = "";
    v.split("\n").reduce((p, l) => p.then(() => exec(l)), Promise.resolve());
  });
  tIn.addEventListener("keydown", e => {
    if (e.key === "ArrowUp" && term.hi > 0) { tIn.value = term.hist[--term.hi]; e.preventDefault(); }
    else if (e.key === "ArrowDown") { term.hi = Math.min(term.hist.length, term.hi + 1); tIn.value = term.hist[term.hi] || ""; e.preventDefault(); }
    else if (e.key === "Tab") {
      e.preventDefault();
      const m = COMPL.filter(x => x.startsWith(tIn.value) && x !== tIn.value);
      if (m.length === 1) tIn.value = m[0];
      else if (m.length > 1) { let p = m[0]; m.forEach(x => { while (!x.startsWith(p)) p = p.slice(0, -1); }); if (p.length > tIn.value.length) tIn.value = p; else out(m.join("   "), "mut"); }
    }
  });
  tIn.addEventListener("paste", e => {
    const txt = (e.clipboardData || window.clipboardData).getData("text");
    if (!txt.includes("\n")) return;
    e.preventDefault();
    const lines = txt.split(/\r?\n/).filter(l => l.trim());
    if (term.target === "pano" && !term.conf && lines[0].startsWith("set ")) lines.unshift("configure");
    lines.reduce((p, l) => p.then(() => exec(l)), Promise.resolve());
  });
  $("#termTarget").addEventListener("change", e => { term.target = e.target.value; term.conf = false; updPrompt(); out(`— connexion SSH à ${host()} —`, "mut"); tIn.focus(); });
  $("#term").addEventListener("click", e => { if (!e.target.closest("select") && !getSelection().toString()) tIn.focus(); });
  function renderCli() {
    $("#cliCh").innerHTML = `<h3>Défis CLI</h3>` + CLI.map(x => `<details class="cc${st.cli[x.id] ? " done" : ""}"><summary><i>${st.cli[x.id] ? "✓" : "⌨"}</i><span><b>${x.title}</b><small>${x.target === "pano" ? "Panorama" : DG[x.target].fw} · ${x.goal}</small></span></summary>
      <pre class="code">${esc(x.answer)}</pre><button type="button" class="btn sm ghost" data-try="${x.id}">Préparer dans le terminal</button></details>`).join("");
  }
  $("#cliCh").addEventListener("click", e => {
    const b = e.target.closest("[data-try]"); if (!b) return;
    const x = CLI.find(y => y.id === b.dataset.try);
    if (term.target !== x.target) { $("#termTarget").value = x.target; $("#termTarget").dispatchEvent(new Event("change")); }
    const lines = x.answer.split("\n");
    if (lines.length > 1 && lines[0] === "configure" && term.conf) lines.shift();
    if (lines.length > 1) { exec(lines[0]); tIn.value = lines[1]; } else tIn.value = lines[0];
    tIn.focus();
  });

  /* =========================================================
     QUIZ
     ========================================================= */
  const qz = { i: 0, score: 0, answered: false, order: [] };
  const qBox = $("#quizBox");
  function quizStart() { qz.i = 0; qz.score = 0; qz.answered = false; qz.order = QUIZ.map((_, i) => i).sort(() => Math.random() - .5); renderQuiz(); }
  function renderQuiz() {
    if (qz.i >= QUIZ.length) {
      const pct = Math.round(100 * qz.score / QUIZ.length);
      if (qz.score > (st.quiz.best || 0)) { st.quiz.best = qz.score; save(); renderProgress(); }
      qBox.innerHTML = `<div class="qz-card qz-end glass"><div class="qz-big">${qz.score}<small>/${QUIZ.length}</small></div><p>${pct >= 80 ? "Excellent : tu maîtrises la double barrière et Panorama." : pct >= 50 ? "Pas mal ! Relis les explications et retente." : "Refais les exercices du lab puis retente le quiz."}</p><button type="button" class="btn primary" data-q="restart">Recommencer</button></div>`;
      if (pct >= 80) confetti($(".qz-big", qBox));
      return;
    }
    const Q = QUIZ[qz.order[qz.i]];
    qBox.innerHTML = `<div class="qz-card glass">
      <div class="qz-dots">${QUIZ.map((_, i) => `<i class="${i < qz.i ? "d" : i === qz.i ? "c" : ""}"></i>`).join("")}</div>
      <span class="kicker">Question ${qz.i + 1}/${QUIZ.length} · score ${qz.score}</span>
      <h3>${esc(Q.q)}</h3>
      <div class="qz-opts">${Q.o.map((o, i) => `<button type="button" data-o="${i}">${String.fromCharCode(65 + i)}. ${esc(o)}</button>`).join("")}</div>
      <div class="qz-exp" aria-live="polite"></div></div>`;
  }
  qBox.addEventListener("click", e => {
    const b = e.target.closest("[data-o],[data-q]"); if (!b) return;
    if (b.dataset.q === "restart") return quizStart();
    if (b.dataset.q === "next") { qz.i++; qz.answered = false; return renderQuiz(); }
    if (qz.answered) return;
    qz.answered = true;
    const Q = QUIZ[qz.order[qz.i]], ok = +b.dataset.o === Q.a;
    if (ok) qz.score++;
    $$("[data-o]", qBox).forEach(x => { x.disabled = true; if (+x.dataset.o === Q.a) x.classList.add("right"); });
    if (!ok) b.classList.add("wrong");
    $(".qz-exp", qBox).innerHTML = `<div class="callout ${ok ? "ok" : "bad"}"><b>${ok ? "✓ Bonne réponse" : "✗ Raté"}</b>${esc(Q.e)}</div><button type="button" class="btn sm primary" data-q="next">${qz.i < QUIZ.length - 1 ? "Question suivante →" : "Voir le score"}</button>`;
  });

  /* =========================================================
     DÉMARRAGE
     ========================================================= */
  function renderAll() { checkAll(true); renderPano(); renderMissions(); renderProgress(); renderTraceSync(); renderCli(); updPrompt(); }
  renderAll();
  quizStart();
  out("PAN-OS / Panorama CLI simulée · tape « help » pour la liste des commandes.", "mut");
  if (window.PAN && PAN.palette) PAN.palette.add([
    { icon: "🧪", label: "Lab : console Panorama", sub: "Simulation", href: "#lab", group: "Lab" },
    { icon: "📦", label: "Lab : packet tracer", sub: "Tracer un flux", href: "#trace", group: "Lab" },
    { icon: "⌨️", label: "Lab : terminal CLI", sub: "Défis CLI", href: "#cli", group: "Lab" },
    { icon: "❓", label: "Lab : quiz", sub: "14 questions", href: "#quiz", group: "Lab" },
  ]);
})();
