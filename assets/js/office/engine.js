/* Office Academy — moteur de la formation Word · Excel · PowerPoint
   Registre (OFFICE.add / OFFICE.sim), routage par hash (#/excel, #/excel/x03/labs, #/export),
   progression locale (XP, rangs), cours, labs vérifiés en direct, défis raccourcis, quiz,
   mode présentation formateur, exports hors ligne (supports HTML, fichiers d'exercice, kit .zip, attestation).
   Contenu : assets/js/office/{word,excel,powerpoint}.js — contrat : assets/js/office/SPEC.md */
(function () {
  "use strict";
  const APPS = {
    word: { id: "word", name: "Word", letter: "W", c1: "#2B7CD3", c2: "#6CB8FF", rgb: "43,124,211", tag: "Traitement de texte",
      desc: "Rédiger, structurer et automatiser des documents professionnels impeccables." },
    excel: { id: "excel", name: "Excel", letter: "X", c1: "#21A366", c2: "#5FD99A", rgb: "33,163,102", tag: "Tableur",
      desc: "Calculer, analyser et visualiser des données de façon fiable et rapide." },
    powerpoint: { id: "powerpoint", name: "PowerPoint", letter: "P", c1: "#D35230", c2: "#FF9A73", rgb: "211,82,48", tag: "Présentation",
      desc: "Concevoir des présentations claires, animées et vraiment percutantes." },
  };
  const LEVELS = {
    1: { name: "Débutant", icon: "🌱", color: "#22C55E", sub: "Prendre en main l'interface et les gestes essentiels" },
    2: { name: "Intermédiaire", icon: "🚀", color: "#38BDF8", sub: "Gagner en méthode, en structure et en vitesse" },
    3: { name: "Avancé", icon: "⚡", color: "#A78BFA", sub: "Maîtriser les fonctions puissantes et les documents complexes" },
    4: { name: "Expert", icon: "👑", color: "#F59E0B", sub: "Automatiser, industrialiser, former les autres" },
  };
  const KIND = { sheet: "Tableur simulé", doc: "Document simulé", slides: "Présentation simulée", keys: "Défi raccourcis", steps: "Dans la vraie appli" };
  const DATA = { word: [], excel: [], powerpoint: [] }, SIMS = {};
  const OFFICE = window.OFFICE = {
    add(app, mods) { DATA[app] = (DATA[app] || []).concat(mods || []).sort((a, b) => a.id.localeCompare(b.id)); },
    sim(kind, fn) { SIMS[kind] = fn; },
    data: DATA, apps: APPS, levels: LEVELS,
  };

  const PAN = window.PAN || {};
  const esc = PAN.esc || (s => String(s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c])));
  const toast = PAN.toast || (m => console.log(m));
  const confetti = PAN.confetti || (() => {});
  const reduce = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const clone = o => JSON.parse(JSON.stringify(o || {}));
  const strip = h => String(h || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  const slug = s => String(s).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^\w]+/g, "-").replace(/^-|-$/g, "").slice(0, 60);
  const debounce = (f, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => f(...a), ms); }; };

  /* ================= progression ================= */
  const KEY = "office-academy:v1";
  const blank = () => ({ seen: {}, labs: {}, quiz: {}, keys: {}, name: "", trainer: false });
  let P = blank();
  try { P = Object.assign(blank(), JSON.parse(localStorage.getItem(KEY) || "{}")); } catch { /* stockage indisponible */ }
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(P)); } catch { /* */ } refreshXP(); };
  const allMods = () => Object.keys(DATA).flatMap(a => DATA[a].map(m => ({ app: a, m })));
  const findMod = id => { for (const a of Object.keys(DATA)) { const i = DATA[a].findIndex(m => m.id === id); if (i >= 0) return { app: a, m: DATA[a][i], i }; } return null; };
  function modStats(m) {
    const L = m.lessons.length, seen = (P.seen[m.id] || []).filter(i => i < L).length;
    const labs = m.labs.filter(l => P.labs[l.id] && P.labs[l.id].done).length, q = P.quiz[m.id];
    const qs = q ? Math.min(1, q.best / q.n / .8) : 0;
    const pct = Math.round(((L ? seen / L : 1) * .3 + (m.labs.length ? labs / m.labs.length : 1) * .5 + qs * .2) * 100);
    return { seen, L, labs, nLabs: m.labs.length, quiz: q, pct, done: pct >= 100 };
  }
  const xpOf = m => { const s = modStats(m); return s.seen * 10 + s.labs * 50 + (s.quiz ? s.quiz.best * 8 : 0); };
  const totalXP = () => allMods().reduce((t, { m }) => t + xpOf(m), 0);
  const RANKS = [[0, "Curieux", "🐣"], [150, "Apprenti", "🌱"], [500, "Opérationnel", "🛠️"], [1100, "Confirmé", "🚀"], [1900, "Expert", "⚡"], [2900, "Maître Office", "👑"]];
  const rankOf = xp => { let r = RANKS[0], nx = null; for (let i = 0; i < RANKS.length; i++) if (xp >= RANKS[i][0]) { r = RANKS[i]; nx = RANKS[i + 1] || null; } return { name: r[1], icon: r[2], min: r[0], next: nx }; };
  const appStats = a => { const ms = DATA[a]; const p = ms.length ? Math.round(ms.reduce((t, m) => t + modStats(m).pct, 0) / ms.length) : 0; return { pct: p, done: ms.filter(m => modStats(m).done).length, n: ms.length }; };
  function refreshXP() {
    const xp = totalXP(), r = rankOf(xp);
    document.querySelectorAll("[data-o-xp]").forEach(el => { el.textContent = xp.toLocaleString("fr-FR"); });
    document.querySelectorAll("[data-o-rank]").forEach(el => { el.textContent = r.icon + " " + r.name; });
    document.querySelectorAll("[data-o-xpbar]").forEach(el => { el.style.width = (r.next ? Math.min(100, (xp - r.min) / (r.next[0] - r.min) * 100) : 100) + "%"; });
  }
  function award(xp, msg, el) {
    const before = rankOf(totalXP() - xp).name, after = rankOf(totalXP());
    toast(`${msg} · +${xp} XP`);
    confetti(el);
    const f = document.createElement("div"); f.className = "o-xpfly"; f.textContent = "+" + xp + " XP"; document.body.append(f); setTimeout(() => f.remove(), 1500);
    if (after.name !== before) setTimeout(() => { toast(`Nouveau rang : ${after.icon} ${after.name} !`); confetti(); }, 900);
  }

  /* ================= petits composants ================= */
  const ring = (pct, size = 54, stroke = 5, label) => {
    const r = (size - stroke) / 2, c = 2 * Math.PI * r;
    return `<svg class="o-ring" width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" aria-label="${pct} %"><circle cx="${size / 2}" cy="${size / 2}" r="${r}" stroke-width="${stroke}" class="o-ring-bg"/>
      <circle cx="${size / 2}" cy="${size / 2}" r="${r}" stroke-width="${stroke}" class="o-ring-fg" stroke-dasharray="${c}" stroke-dashoffset="${c * (1 - pct / 100)}" style="--c:${c}"/>
      <text x="50%" y="50%" dy=".35em" text-anchor="middle">${label != null ? label : pct + "%"}</text></svg>`;
  };
  const appLogo = (a, size = 44) => `<span class="o-logo" style="--s:${size}px;--a1:${APPS[a].c1};--a2:${APPS[a].c2}">${APPS[a].letter}</span>`;
  const levelChip = l => `<span class="o-lvl" style="--lc:${LEVELS[l].color}">${LEVELS[l].icon} ${LEVELS[l].name}</span>`;
  function keysHTML(s) {
    const parts = splitKeys(s);
    return `<span class="o-keys">${parts.map(p => `<kbd>${esc(p)}</kbd>`).join("<i>+</i>")}</span>`;
  }
  function splitKeys(s) {
    s = String(s).trim();
    const out = []; let cur = "";
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === "+" && cur !== "") { out.push(cur.trim()); cur = ""; continue; }
      cur += ch;
    }
    if (cur !== "") out.push(cur.trim());
    return out.filter(Boolean);
  }
  function download(name, content, type) {
    const blob = content instanceof Blob ? content : new Blob([content], { type: type || "text/plain;charset=utf-8" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }

  /* ================= routage ================= */
  const root = document.getElementById("o-root");
  let current = null; // { destroy() } de la vue active (labs)
  function route() {
    const h = location.hash.replace(/^#\/?/, "").split("/").filter(Boolean);
    if (current && current.destroy) { try { current.destroy(); } catch { /* */ } }
    current = null;
    let html, after, app = "";
    if (!h.length) { html = viewHub(); after = afterHub; }
    else if (h[0] === "export") { html = viewExport(); after = afterExport; }
    else if (APPS[h[0]] && !h[1]) { app = h[0]; html = viewApp(h[0]); after = () => afterApp(h[0]); }
    else if (APPS[h[0]] && h[1]) {
      const f = findMod(h[1]);
      if (!f) { location.hash = "#/" + h[0]; return; }
      app = f.app; html = viewModule(f, h[2] || "cours", h[3]); after = () => afterModule(f, h[2] || "cours", h[3]);
    } else { location.hash = ""; return; }
    const prevApp = document.body.dataset.app;
    const swap = () => {
      document.body.dataset.app = app; root.dataset.app = app;
      root.innerHTML = html; after && after();
      PAN.reveal && PAN.reveal();
      refreshXP();
    };
    const sameModule = root.dataset.mod && h[1] && root.dataset.mod === h[1];
    root.dataset.mod = h[1] || "";
    if (PAN.transition && prevApp !== undefined) PAN.transition(swap); else swap();
    if (!sameModule) scrollTo({ top: 0, behavior: "instant" in window ? "instant" : "auto" });
  }
  addEventListener("hashchange", route);

  /* ================= vue : accueil ================= */
  function counts() {
    const ms = allMods().map(x => x.m);
    return { mods: ms.length, labs: ms.reduce((t, m) => t + m.labs.length, 0), q: ms.reduce((t, m) => t + m.quiz.length, 0),
      lessons: ms.reduce((t, m) => t + m.lessons.length, 0), hours: Math.round(ms.reduce((t, m) => t + (m.duration || 45), 0) / 60),
      sim: ms.reduce((t, m) => t + m.labs.filter(l => ["sheet", "doc", "slides"].includes(l.kind)).length, 0) };
  }
  function viewHub() {
    const n = counts(), xp = totalXP(), r = rankOf(xp);
    return `
    <header class="o-hero">
      <div class="o-hero-bg" aria-hidden="true"><i class="b1"></i><i class="b2"></i><i class="b3"></i></div>
      <div class="wrap o-hero-in">
        <div class="o-hero-txt">
          <span class="kicker fade-up">Office Academy · Microsoft 365 en français</span>
          <h1 class="o-h1"><span class="l"><span>Word, Excel, PowerPoint :</span></span><span class="l"><span class="o-grad">du premier clic</span></span><span class="l"><span class="o-grad">au niveau expert.</span></span></h1>
          <p class="lead fade-up" style="animation-delay:.35s">${n.mods} modules progressifs, <b>${n.labs} labs</b> dont ${n.sim} dans des simulateurs qui <b>vérifient ton travail en direct</b>, des défis raccourcis chronométrés, des quiz, un mode formateur et l'export complet hors ligne.</p>
          <div class="o-cta fade-up" style="animation-delay:.45s">
            <a class="btn primary" href="#/excel">Commencer la formation →</a>
            <a class="btn ghost" href="#/export">🎓 Espace formateur & export</a>
          </div>
          <div class="stats o-stats fade-up" style="animation-delay:.55s">
            <div class="stat"><b data-count="${n.mods}">0</b><span>modules</span></div>
            <div class="stat"><b data-count="${n.labs}">0</b><span>labs pratiques</span></div>
            <div class="stat"><b data-count="${n.q}">0</b><span>questions de quiz</span></div>
            <div class="stat"><b data-count="${n.hours}" data-suffix=" h">0</b><span>de contenu</span></div>
          </div>
        </div>
        <div class="o-hero-art" aria-hidden="true">
          <div class="o-stage">
            <div class="o-tile t-w" data-tilt="14"><span>W</span><small>Word</small></div>
            <div class="o-tile t-x" data-tilt="14"><span>X</span><small>Excel</small></div>
            <div class="o-tile t-p" data-tilt="14"><span>P</span><small>PowerPoint</small></div>
            <div class="o-mini glass">
              <div class="o-mini-bar"><i></i><i></i><i></i><b>Lab · Budget.xlsx</b></div>
              <div class="o-mini-fx"><em>fx</em><span class="o-type"></span><i class="o-caret"></i></div>
              <div class="o-mini-grid">${["Région", "CA", "Objectif", "Nord", "48 200", "✓", "Sud", "39 750", "✓", "Est", "51 900", "✓"].map((x, i) => `<span class="${i < 3 ? "h" : ""}" style="--d:${i * 70}ms">${x}</span>`).join("")}</div>
              <div class="o-mini-chart">${[62, 48, 80, 70, 92].map((v, i) => `<i style="--h:${v}%;--d:${i * 120}ms"></i>`).join("")}</div>
              <div class="o-mini-ok">✓ Tâche validée · +50 XP</div>
            </div>
            <div class="o-orbit"><i></i><i></i><i></i></div>
          </div>
        </div>
      </div>
    </header>

    <section class="o-sec" id="apps">
      <div class="wrap">
        <span class="kicker" data-reveal>Trois parcours</span>
        <h2 data-reveal>Choisis ton application.</h2>
        <div class="o-apps" data-stagger="120" data-stagger-type="zoom">
          ${Object.values(APPS).map(A => { const s = appStats(A.id), ms = DATA[A.id];
            return `<a class="o-appcard" href="#/${A.id}" style="--a1:${A.c1};--a2:${A.c2};--arg:${A.rgb}" data-tilt="6">
              <div class="o-appcard-top">${appLogo(A.id, 58)}<div><small>${A.tag}</small><h3>${A.name}</h3></div>${ring(s.pct, 58, 5)}</div>
              <p>${A.desc}</p>
              <div class="o-ladder">${[1, 2, 3, 4].map(l => { const lm = ms.filter(m => m.level === l); return `<span style="--lc:${LEVELS[l].color}" title="${LEVELS[l].name}"><b>${LEVELS[l].icon}</b>${lm.length} mod.</span>`; }).join("")}</div>
              <ul>${ms.slice(0, 4).map(m => `<li>${m.icon} ${esc(m.title)}</li>`).join("")}${ms.length > 4 ? `<li class="mut">+ ${ms.length - 4} autres modules…</li>` : ""}</ul>
              <div class="o-appcard-foot"><span>${ms.reduce((t, m) => t + m.labs.length, 0)} labs · ${ms.reduce((t, m) => t + m.quiz.length, 0)} questions</span><b>Ouvrir le parcours →</b></div>
            </a>`; }).join("")}
        </div>
      </div>
    </section>

    <section class="o-sec">
      <div class="wrap">
        <span class="kicker" data-reveal>Progression</span>
        <h2 data-reveal>Quatre paliers, une vraie montée en compétence.</h2>
        <div class="o-levels" data-stagger="110">
          ${[1, 2, 3, 4].map(l => `<div class="o-level" style="--lc:${LEVELS[l].color}"><div class="o-level-i">${LEVELS[l].icon}</div><b>${LEVELS[l].name}</b><p>${LEVELS[l].sub}</p><span class="o-level-n">Niveau ${l}</span></div>`).join("")}
        </div>
      </div>
    </section>

    <section class="o-sec">
      <div class="wrap">
        <span class="kicker" data-reveal>Pédagogie active</span>
        <h2 data-reveal>Des labs comme dans la vraie application.</h2>
        <div class="o-feats" data-stagger="90">
          <div class="card o-feat"><div class="o-feat-i">🧪</div><h3>Labs vérifiés en direct</h3><p class="mut">Un tableur, un traitement de texte et un éditeur de diapos simulés : chaque tâche se coche toute seule dès qu'elle est réussie.</p></div>
          <div class="card o-feat"><div class="o-feat-i">⌨️</div><h3>Défis raccourcis</h3><p class="mut">Tape les raccourcis au clavier, chronomètre et précision à l'appui : la mémoire musculaire s'installe vite.</p></div>
          <div class="card o-feat"><div class="o-feat-i">✅</div><h3>Quiz expliqués</h3><p class="mut">Chaque réponse est justifiée. Ton meilleur score débloque la validation du module.</p></div>
          <div class="card o-feat"><div class="o-feat-i">🎤</div><h3>Mode présentation</h3><p class="mut">Projette le cours en plein écran : diapos générées, minuteur, pointeur laser, écran noir.</p></div>
          <div class="card o-feat"><div class="o-feat-i">📦</div><h3>Export hors ligne</h3><p class="mut">Supports HTML imprimables, fichiers d'exercice (.csv, .doc, .rtf), kit formateur .zip, site complet.</p></div>
          <div class="card o-feat"><div class="o-feat-i">🏅</div><h3>XP, rangs & attestation</h3><p class="mut">Progression sauvegardée dans le navigateur, exportable, et attestation de formation imprimable.</p></div>
        </div>
      </div>
    </section>

    <section class="o-sec">
      <div class="wrap">
        <div class="o-me glass" data-reveal="zoom">
          <div class="o-me-l"><span class="kicker">Ma progression</span><h2><span data-o-rank>${r.icon} ${r.name}</span></h2>
            <p class="mut"><b data-o-xp>${xp}</b> XP${r.next ? ` · prochain rang « ${r.next[1]} » à ${r.next[0]} XP` : " · rang maximal atteint 👑"}</p>
            <div class="o-xpbar"><i data-o-xpbar></i></div></div>
          <div class="o-me-r">${Object.values(APPS).map(A => { const s = appStats(A.id); return `<a href="#/${A.id}" class="o-me-app" style="--a1:${A.c1};--a2:${A.c2}">${ring(s.pct, 76, 6)}<b>${A.name}</b><small>${s.done}/${s.n} modules validés</small></a>`; }).join("")}</div>
        </div>
      </div>
    </section>`;
  }
  let typeTimer = null;
  function afterHub() {
    const el = root.querySelector(".o-type"); if (!el) return;
    const F = ["=SOMME(B2:B13)", "=RECHERCHEX(A2;Clients[ID];Clients[Ville])", "=SI(C2>=Objectif;\"✓\";\"À relancer\")", "=SOMME.SI.ENS(CA;Région;\"Nord\")", "=SIERREUR(RECHERCHEV(A2;Tarifs;2;FAUX);0)"];
    let i = 0, j = 0, del = false;
    clearInterval(typeTimer);
    if (reduce) { el.textContent = F[0]; return; }
    typeTimer = setInterval(() => {
      if (!document.body.contains(el)) return clearInterval(typeTimer);
      const s = F[i];
      if (!del) { j++; el.textContent = s.slice(0, j); if (j >= s.length) { del = true; j = s.length + 18; } }
      else { j--; if (j <= s.length) el.textContent = s.slice(0, Math.max(0, j)); if (j <= 0) { del = false; i = (i + 1) % F.length; } }
    }, 65);
  }

  /* ================= vue : application ================= */
  function viewApp(a) {
    const A = APPS[a], ms = DATA[a], s = appStats(a);
    return `
    <section class="o-apphero" style="--a1:${A.c1};--a2:${A.c2}">
      <div class="wrap o-apphero-in">
        <div>${crumbs([[A.name]])}
          <div class="o-apphero-t">${appLogo(a, 72)}<div><span class="kicker">${A.tag} · Microsoft 365</span><h1>Formation ${A.name}</h1></div></div>
          <p class="lead">${A.desc} ${ms.length} modules, du niveau débutant au niveau expert.</p>
          <div class="o-cta">${ms[0] ? `<a class="btn primary" href="#/${a}/${(ms.find(m => !modStats(m).done) || ms[0]).id}">${s.pct ? "Reprendre" : "Commencer"} →</a>` : ""}
            <button class="btn ghost" type="button" data-exp-app="${a}">⬇ Support complet (HTML)</button>
            <a class="btn ghost" href="#/export">📦 Export & kit formateur</a></div>
        </div>
        <div class="o-apphero-r">${ring(s.pct, 132, 10)}<small>${s.done}/${s.n} modules validés</small></div>
      </div>
    </section>
    <section class="o-sec o-path-sec">
      <div class="wrap">
        <div class="o-filter" role="tablist" aria-label="Filtrer par niveau">
          <button type="button" class="on" data-lv="0">Tous</button>${[1, 2, 3, 4].map(l => `<button type="button" data-lv="${l}" style="--lc:${LEVELS[l].color}">${LEVELS[l].icon} ${LEVELS[l].name}</button>`).join("")}
        </div>
        <div class="o-timeline">
          ${[1, 2, 3, 4].map(l => { const lm = ms.filter(m => m.level === l); if (!lm.length) return "";
            return `<div class="o-tl-level" data-level="${l}" style="--lc:${LEVELS[l].color}">
              <div class="o-tl-head" data-reveal="left"><span class="o-tl-dot">${LEVELS[l].icon}</span><div><b>Niveau ${l} · ${LEVELS[l].name}</b><small>${LEVELS[l].sub}</small></div></div>
              <div class="o-tl-mods" data-stagger="90">${lm.map(m => modCard(a, m)).join("")}</div></div>`; }).join("")}
        </div>
      </div>
    </section>`;
  }
  function modCard(a, m) {
    const s = modStats(m);
    return `<a class="card hover o-modcard${s.done ? " done" : ""}" href="#/${a}/${m.id}" data-tilt="5">
      <div class="o-modcard-top"><span class="o-modicon">${m.icon}</span>${ring(s.pct, 48, 4.5, s.done ? "✓" : null)}</div>
      <small class="o-modid">${m.id.toUpperCase()} · ${m.duration || 45} min</small>
      <h3>${esc(m.title)}</h3><p class="mut">${esc(m.summary || "")}</p>
      <div class="o-modmeta"><span>📖 ${s.seen}/${s.L}</span><span>🧪 ${s.labs}/${s.nLabs}</span><span>✅ ${s.quiz ? s.quiz.best + "/" + s.quiz.n : "—"}</span></div>
    </a>`;
  }
  function afterApp(a) {
    root.querySelectorAll("[data-lv]").forEach(b => b.addEventListener("click", () => {
      root.querySelectorAll("[data-lv]").forEach(x => x.classList.toggle("on", x === b));
      root.querySelectorAll(".o-tl-level").forEach(l => { l.hidden = b.dataset.lv !== "0" && l.dataset.level !== b.dataset.lv; });
    }));
  }
  const crumbs = items => `<nav class="o-crumbs" aria-label="Fil d'Ariane"><a href="#/">Office Academy</a>${items.map(([t, h]) => `<span>›</span>${h ? `<a href="${h}">${esc(t)}</a>` : `<b>${esc(t)}</b>`}`).join("")}</nav>`;

  /* ================= vue : module ================= */
  const TABS = [["cours", "📖 Cours"], ["labs", "🧪 Labs"], ["raccourcis", "⌨️ Raccourcis"], ["quiz", "✅ Quiz"]];
  function viewModule(f, tab, sub) {
    const { app: a, m, i } = f, A = APPS[a], ms = DATA[a], prev = ms[i - 1], next = ms[i + 1], s = modStats(m);
    return `
    <section class="o-modhero" style="--a1:${A.c1};--a2:${A.c2}">
      <div class="wrap">
        ${crumbs([[A.name, "#/" + a], [m.title]])}
        <div class="o-modhero-in">
          <span class="o-modhero-icon">${m.icon}</span>
          <div class="o-modhero-t"><div class="o-modhero-tags">${levelChip(m.level)}<span class="chip">⏱ ${m.duration || 45} min</span><span class="chip">${m.lessons.length} leçons</span><span class="chip">${m.labs.length} lab${m.labs.length > 1 ? "s" : ""}</span><span class="chip">${m.quiz.length} questions</span></div>
            <h1>${esc(m.title)}</h1><p class="lead">${esc(m.summary || "")}</p></div>
          <div class="o-modhero-r">${ring(s.pct, 92, 8)}
            <div class="o-modhero-btns"><button class="btn sm primary" type="button" data-present>🎤 Présenter</button><button class="btn sm ghost" type="button" data-exp-mod="${m.id}">⬇ Support HTML</button></div></div>
        </div>
        <div class="o-tabs" role="tablist">${TABS.map(([k, t]) => `<a role="tab" href="#/${a}/${m.id}/${k}" class="${k === tab ? "on" : ""}" aria-selected="${k === tab}">${t}${k === "labs" ? ` <small>${s.labs}/${s.nLabs}</small>` : k === "cours" ? ` <small>${s.seen}/${s.L}</small>` : k === "quiz" && s.quiz ? ` <small>${s.quiz.best}/${s.quiz.n}</small>` : ""}</a>`).join("")}<span class="o-tabs-ink"></span></div>
      </div>
    </section>
    <section class="o-modbody"><div class="wrap" id="o-tab"></div>
      <div class="wrap o-modnav">${prev ? `<a class="card hover" href="#/${a}/${prev.id}"><small>← Module précédent</small><b>${prev.icon} ${esc(prev.title)}</b></a>` : "<span></span>"}
      ${next ? `<a class="card hover o-next" href="#/${a}/${next.id}"><small>Module suivant →</small><b>${next.icon} ${esc(next.title)}</b></a>` : `<a class="card hover o-next" href="#/${a}"><small>Fin du parcours 🎉</small><b>Retour au parcours ${A.name}</b></a>`}</div>
    </section>`;
  }
  function afterModule(f, tab, sub) {
    const host = root.querySelector("#o-tab");
    const ink = root.querySelector(".o-tabs-ink"), on = root.querySelector(".o-tabs .on");
    if (ink && on) requestAnimationFrame(() => { ink.style.cssText = `left:${on.offsetLeft}px;width:${on.offsetWidth}px`; });
    root.querySelector("[data-present]").onclick = () => present(f);
    if (tab === "labs") current = tabLabs(host, f, +sub || 0);
    else if (tab === "raccourcis") tabKeys(host, f);
    else if (tab === "quiz") tabQuiz(host, f);
    else tabCours(host, f, +sub || 0);
  }

  /* ----- cours ----- */
  function tabCours(host, f, idx) {
    const { app: a, m } = f;
    idx = Math.max(0, Math.min(m.lessons.length - 1, idx));
    const seen = new Set(P.seen[m.id] || []);
    host.innerHTML = `
      <div class="o-cours">
        <aside class="o-cours-side">
          <div class="card o-obj"><h4>🎯 Objectifs</h4><ul>${(m.objectives || []).map(o => `<li>${esc(o)}</li>`).join("")}</ul></div>
          <ol class="o-lessons">${m.lessons.map((l, i) => `<li><button type="button" data-l="${i}" class="${i === idx ? "on" : ""}${seen.has(i) ? " seen" : ""}"><span>${seen.has(i) ? "✓" : i + 1}</span>${esc(l.title)}</button></li>`).join("")}</ol>
        </aside>
        <article class="o-lesson card" aria-live="polite"></article>
      </div>`;
    const art = host.querySelector(".o-lesson");
    function show(i, anim) {
      idx = i; const l = m.lessons[i];
      art.innerHTML = `<div class="o-lesson-h"><small>Leçon ${i + 1} / ${m.lessons.length}</small><h2>${esc(l.title)}</h2></div><div class="o-content">${l.html}</div>
        <div class="o-lesson-nav">${i > 0 ? `<button class="btn ghost" type="button" data-go="${i - 1}">← Précédente</button>` : "<span></span>"}
        ${i < m.lessons.length - 1 ? `<button class="btn primary" type="button" data-go="${i + 1}">J'ai compris, leçon suivante →</button>` : `<a class="btn primary" href="#/${a}/${m.id}/labs" data-golab>Passer au lab 🧪</a>`}</div>`;
      if (anim && !reduce) { art.classList.remove("o-in"); void art.offsetWidth; art.classList.add("o-in"); }
      host.querySelectorAll("[data-l]").forEach(b => b.classList.toggle("on", +b.dataset.l === i));
      enhance(art);
    }
    function markSeen(i) {
      const arr = new Set(P.seen[m.id] || []); if (arr.has(i)) return;
      arr.add(i); P.seen[m.id] = [...arr]; save();
      const b = host.querySelector(`[data-l="${i}"]`); if (b) { b.classList.add("seen"); b.querySelector("span").textContent = "✓"; }
      toast("Leçon validée · +10 XP");
    }
    host.addEventListener("click", e => {
      const l = e.target.closest("[data-l]"), g = e.target.closest("[data-go]"), gl = e.target.closest("[data-golab]");
      if (l) show(+l.dataset.l, true);
      if (g) { if (+g.dataset.go > idx) markSeen(idx); show(+g.dataset.go, true); art.scrollIntoView({ block: "start", behavior: reduce ? "auto" : "smooth" }); }
      if (gl) markSeen(idx);
    });
    show(idx);
  }
  // enrichit le contenu : chemins de ruban animés, boutons copier sur les formules
  function enhance(el) {
    el.querySelectorAll(".o-path").forEach(p => {
      if (p.dataset.done) return; p.dataset.done = 1;
      const parts = p.textContent.split("›").map(x => x.trim()).filter(Boolean);
      p.innerHTML = parts.map((x, i) => `<span style="--d:${i * 140}ms">${esc(x)}</span>`).join("<i>›</i>");
    });
    el.querySelectorAll("pre.code").forEach(pre => {
      if (pre.dataset.done) return; pre.dataset.done = 1;
      const b = document.createElement("button"); b.type = "button"; b.className = "o-copy"; b.textContent = "Copier";
      b.onclick = () => { PAN.copy ? PAN.copy(pre.innerText.replace(/Copier$/, "").trim()) : navigator.clipboard.writeText(pre.innerText); b.textContent = "✓"; setTimeout(() => { b.textContent = "Copier"; }, 1200); };
      pre.append(b);
    });
  }

  /* ----- raccourcis ----- */
  function tabKeys(host, f) {
    const { m, app: a } = f;
    const keyLab = m.labs.findIndex(l => l.kind === "keys");
    host.innerHTML = `<div class="o-keysgrid" data-stagger="50">${(m.shortcuts || []).map(s => `<div class="card o-keycard">${keysHTML(s.keys)}<span>${esc(s.desc)}</span></div>`).join("") || "<p class='mut'>Pas de raccourci pour ce module.</p>"}</div>
      <div class="callout" style="margin-top:22px">Raccourcis Windows. Sur Mac, <kbd>Ctrl</kbd> devient généralement <kbd>⌘</kbd> et <kbd>Alt</kbd> devient <kbd>⌥</kbd>.
      ${keyLab >= 0 ? ` <a href="#/${a}/${m.id}/labs/${keyLab}"><b>→ Lance le défi raccourcis de ce module</b></a>` : ""}</div>`;
    PAN.reveal && PAN.reveal();
  }

  /* ----- quiz ----- */
  function tabQuiz(host, f) {
    const { m } = f;
    let order, qi, score, answers;
    const shuffle = arr => { const a = arr.slice(); for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; } return a; };
    function start() { order = m.quiz.map((q, i) => ({ q, i, ch: shuffle(q.choices.map((c, k) => ({ c, k }))) })); qi = 0; score = 0; answers = []; show(); }
    function show() {
      const o = order[qi];
      host.innerHTML = `<div class="o-quiz card">
        <div class="o-quiz-top"><span>Question ${qi + 1} / ${order.length}</span><div class="o-quiz-prog">${order.map((_, k) => `<i class="${k < qi ? (answers[k] ? "ok" : "ko") : k === qi ? "cur" : ""}"></i>`).join("")}</div><span>Score : <b>${score}</b></span></div>
        <h2 class="o-quiz-q">${esc(o.q.q)}</h2>
        <div class="o-choices">${o.ch.map((c, k) => `<button type="button" data-k="${c.k}" style="--d:${k * 60}ms"><span>${"ABCD"[k]}</span>${esc(c.c)}</button>`).join("")}</div>
        <div class="o-explain" hidden></div>
      </div>`;
      if (P.trainer) host.querySelector(`[data-k="${o.q.answer}"]`).classList.add("o-hint");
      host.querySelectorAll(".o-choices button").forEach(b => b.onclick = () => pick(+b.dataset.k, b));
    }
    function pick(k, btn) {
      const o = order[qi], ok = k === o.q.answer;
      answers[qi] = ok; if (ok) score++;
      host.querySelectorAll(".o-choices button").forEach(b => { b.disabled = true; if (+b.dataset.k === o.q.answer) b.classList.add("ok"); });
      if (!ok) btn.classList.add("ko"); else if (!reduce) confetti(btn);
      const ex = host.querySelector(".o-explain"); ex.hidden = false;
      ex.innerHTML = `<b>${ok ? "✓ Bonne réponse !" : "✗ Pas tout à fait."}</b> ${esc(o.q.explain || "")}<div><button class="btn primary sm" type="button">${qi < order.length - 1 ? "Question suivante →" : "Voir mon score"}</button></div>`;
      ex.querySelector("button").onclick = () => { qi++; qi < order.length ? show() : end(); };
      ex.querySelector("button").focus();
    }
    function end() {
      const n = order.length, pct = Math.round(score / n * 100), prev = P.quiz[m.id];
      if (!prev || score > prev.best) {
        const gain = (score - (prev ? prev.best : 0)) * 8;
        P.quiz[m.id] = { best: score, n, at: Date.now() }; save();
        if (gain > 0) award(gain, "Nouveau record au quiz", host);
      }
      host.innerHTML = `<div class="o-quiz card o-quiz-end">
        ${ring(pct, 150, 12)}<h2>${pct >= 80 ? "Excellent ! 🎉" : pct >= 60 ? "Bien joué 👍" : "Encore un effort 💪"}</h2>
        <p class="lead">${score} bonne${score > 1 ? "s" : ""} réponse${score > 1 ? "s" : ""} sur ${n}. ${pct >= 80 ? "Le quiz du module est validé." : "Il faut 80 % pour valider le quiz : relis les leçons et retente !"}</p>
        <div class="o-cta"><button class="btn primary" type="button" data-again>↻ Recommencer</button><button class="btn ghost" type="button" data-review>Voir le corrigé</button></div>
        <div class="o-review" hidden>${m.quiz.map((q, i) => { const k = order.findIndex(o => o.i === i); return `<div class="o-rev ${answers[k] ? "ok" : "ko"}"><b>${answers[k] ? "✓" : "✗"} ${esc(q.q)}</b><p>Réponse : <b>${esc(q.choices[q.answer])}</b> — ${esc(q.explain || "")}</p></div>`; }).join("")}</div>
      </div>`;
      if (pct >= 80) confetti(host.querySelector(".o-ring"));
      host.querySelector("[data-again]").onclick = start;
      host.querySelector("[data-review]").onclick = () => { host.querySelector(".o-review").hidden = false; };
    }
    start();
  }

  /* ----- labs ----- */
  function tabLabs(host, f, li) {
    const { m } = f;
    if (!m.labs.length) { host.innerHTML = "<p class='mut'>Pas de lab pour ce module.</p>"; return null; }
    li = Math.max(0, Math.min(m.labs.length - 1, li));
    host.innerHTML = `${m.labs.length > 1 ? `<div class="o-labpick">${m.labs.map((l, i) => `<a href="#/${f.app}/${m.id}/labs/${i}" class="${i === li ? "on" : ""}${P.labs[l.id] && P.labs[l.id].done ? " done" : ""}"><small>${KIND[l.kind] || l.kind}</small>${esc(l.title)}</a>`).join("")}</div>` : ""}<div class="o-labhost"></div>`;
    return mountLab(host.querySelector(".o-labhost"), m.labs[li], f);
  }
  function mountLab(box, lab, f) {
    const st = P.labs[lab.id] || {};
    const isSim = !!SIMS[lab.kind] && ["sheet", "doc", "slides"].includes(lab.kind);
    const tasks = lab.kind === "keys" ? (lab.setup.sequence || []).map(s => ({ text: `${esc(s.desc)}`, keys: s.keys })) : (lab.tasks || []);
    box.innerHTML = `
      <div class="o-lab o-lab-${lab.kind}">
        <div class="o-lab-brief">
          <div class="o-lab-h"><span class="chip">${KIND[lab.kind] || lab.kind}</span>${st.done ? '<span class="chip o-donechip">✓ Réussi</span>' : ""}<h2>${esc(lab.title)}</h2></div>
          <div class="o-content o-scn">${lab.scenario || ""}</div>
          <div class="o-taskhead"><b>Tâches</b><span class="o-taskcount"></span></div>
          <div class="o-taskbar"><i></i></div>
          <ol class="o-tasks">${tasks.map((t, i) => `<li data-t="${i}"><span class="o-tick" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="m5 12 5 5 9-10"/></svg></span>
            <div><div class="o-task-t">${t.text}</div>${t.hint ? `<button type="button" class="o-hintb" data-hint="${i}">💡 Indice</button><div class="o-hint-t" hidden>${t.hint}</div>` : ""}</div>
            ${lab.kind === "steps" ? `<input type="checkbox" aria-label="Fait" data-manual="${i}">` : ""}</li>`).join("")}</ol>
          <div class="o-lab-actions">
            ${lab.solution ? `<button type="button" class="btn sm ghost" data-sol>🔑 Corrigé</button>` : ""}
            ${isSim ? `<button type="button" class="btn sm ghost" data-reset>↺ Recommencer</button>` : ""}
            ${isSim ? `<button type="button" class="btn sm ghost" data-full>⛶ Plein écran</button>` : ""}
            ${starterOf(lab, f) ? `<button type="button" class="btn sm ghost" data-starter>⬇ Fichier d'exercice</button>` : ""}
            ${isSim && lab.kind !== "slides" ? `<button type="button" class="btn sm ghost" data-mine>💾 Mon travail</button>` : ""}
          </div>
          <div class="o-sol o-content" hidden>${lab.solution || ""}</div>
        </div>
        <div class="o-lab-stage"><div class="o-sim"></div></div>
      </div>`;
    const simEl = box.querySelector(".o-sim"), labEl = box.querySelector(".o-lab");
    enhance(box);
    let sim = null, keysCtl = null, doneOnce = !!st.done, results = tasks.map(() => false);
    const manual = new Set(lab.kind === "steps" ? (st.t || []).map((v, i) => v ? i : -1).filter(i => i >= 0) : []);
    function paint() {
      const n = results.filter(Boolean).length;
      box.querySelectorAll(".o-tasks li").forEach((li, i) => {
        const was = li.classList.contains("ok");
        li.classList.toggle("ok", results[i]);
        if (results[i] && !was && !reduce) { li.classList.remove("pop"); void li.offsetWidth; li.classList.add("pop"); }
      });
      box.querySelector(".o-taskcount").textContent = `${n} / ${tasks.length}`;
      box.querySelector(".o-taskbar i").style.width = (tasks.length ? n / tasks.length * 100 : 0) + "%";
      if (n === tasks.length && tasks.length && !doneOnce) {
        doneOnce = true;
        P.labs[lab.id] = { done: true, t: results.map(Number), at: Date.now() }; save();
        celebrate(labEl, lab, f);
      } else if (!doneOnce) { P.labs[lab.id] = { done: false, t: results.map(Number) }; save(); }
    }
    const evaluate = () => {
      if (!sim) return;
      results = tasks.map(t => { try { return !!t.check(sim.api); } catch { return false; } });
      paint();
    };
    if (isSim) {
      try {
        sim = SIMS[lab.kind](simEl, clone(lab.setup), { changed: debounce(evaluate, 120), toast, esc, confetti });
        labEl.__sim = sim; labEl.__eval = evaluate; // accès QA automatisée
        evaluate();
      } catch (e) { console.error(e); simEl.innerHTML = `<div class="callout bad">Le simulateur n'a pas pu démarrer : ${esc(e.message)}</div>`; }
    } else if (lab.kind === "keys") {
      keysCtl = keysLab(simEl, lab, (i, ok) => { results[i] = ok; paint(); });
    } else if (lab.kind === "steps") {
      results = tasks.map((_, i) => manual.has(i));
      box.querySelectorAll("[data-manual]").forEach(cb => { cb.checked = manual.has(+cb.dataset.manual); cb.onchange = () => { results[+cb.dataset.manual] = cb.checked; paint(); }; });
      simEl.innerHTML = stepsPanel(lab, f);
      simEl.addEventListener("click", e => { if (e.target.closest("[data-starter]")) downloadStarter(lab, f); });
      paint();
    } else simEl.innerHTML = `<div class="callout warn">Type de lab inconnu : ${esc(lab.kind)}</div>`;
    if (lab.kind === "steps") box.querySelectorAll(".o-tasks li").forEach(li => li.addEventListener("click", e => { if (e.target.closest("button, input, a")) return; const cb = li.querySelector("input"); cb.checked = !cb.checked; cb.onchange(); }));

    box.addEventListener("click", e => {
      const h = e.target.closest("[data-hint]");
      if (h) { const t = h.nextElementSibling; t.hidden = !t.hidden; h.classList.toggle("on", !t.hidden); return; }
      if (e.target.closest("[data-sol]")) { const s = box.querySelector(".o-sol"); s.hidden = !s.hidden; if (!s.hidden) s.scrollIntoView({ block: "nearest", behavior: reduce ? "auto" : "smooth" }); return; }
      if (e.target.closest("[data-reset]") && sim) { sim.reset(); evaluate(); toast("Lab réinitialisé"); return; }
      if (e.target.closest("[data-full]")) { const el = labEl; if (document.fullscreenElement) document.exitFullscreen(); else if (el.requestFullscreen) el.requestFullscreen().catch(() => labEl.classList.toggle("o-fullfake")); else labEl.classList.toggle("o-fullfake"); return; }
      if (e.target.closest("[data-starter]") && e.target.closest(".o-lab-actions")) { downloadStarter(lab, f); return; }
      if (e.target.closest("[data-mine]") && sim && sim.exportFile) { const x = sim.exportFile(); download(slug(lab.id + "-" + lab.title) + "." + x.name.split(".").pop(), x.content, x.type); toast("Ton travail est téléchargé ⬇"); }
    });
    return { destroy() { if (sim && sim.destroy) sim.destroy(); if (keysCtl) keysCtl.destroy(); } };
  }
  function celebrate(el, lab, f) {
    award(50, "Lab réussi 🎉", el);
    const o = document.createElement("div"); o.className = "o-win";
    const next = f.m.labs[f.m.labs.indexOf(lab) + 1];
    o.innerHTML = `<div class="o-win-box glass"><div class="o-win-badge">🏆</div><h3>Lab réussi !</h3><p>${esc(lab.title)} — toutes les tâches sont validées.</p>
      <div class="o-cta">${next ? `<a class="btn primary" href="#/${f.app}/${f.m.id}/labs/${f.m.labs.indexOf(next)}">Lab suivant →</a>` : `<a class="btn primary" href="#/${f.app}/${f.m.id}/quiz">Passer au quiz ✅</a>`}<button class="btn ghost" type="button">Rester ici</button></div></div>`;
    el.append(o);
    requestAnimationFrame(() => o.classList.add("show"));
    o.querySelector("button").onclick = () => { o.classList.remove("show"); setTimeout(() => o.remove(), 300); };
  }
  function stepsPanel(lab, f) {
    const A = APPS[f.app], file = lab.setup && lab.setup.file;
    return `<div class="o-steps-panel" style="--a1:${A.c1};--a2:${A.c2}">
      <div class="o-steps-ill">${appLogo(f.app, 84)}<div class="o-steps-glow"></div></div>
      <h3>Lab à réaliser dans ${esc((lab.setup && lab.setup.app) || A.name)}</h3>
      <p class="mut">Ce lab utilise des fonctionnalités qu'un navigateur ne peut pas simuler fidèlement. Ouvre ${esc(A.name)} sur ton poste, suis les tâches, puis coche-les au fur et à mesure.</p>
      ${file ? `<button class="btn primary" type="button" data-starter>⬇ Télécharger « ${esc(file.name)} »</button>` : ""}
      <ol class="o-steps-how"><li>${file ? "Télécharge le fichier de départ" : `Ouvre un nouveau document ${esc(A.name)}`}</li><li>Réalise chaque tâche (indices disponibles)</li><li>Coche les tâches réussies — le corrigé détaille chaque étape</li></ol>
    </div>`;
  }

  /* ----- défi raccourcis ----- */
  const KEYNAME = { "entrée": "Enter", "entree": "Enter", "enter": "Enter", "espace": " ", "space": " ", "suppr": "Delete", "supprimer": "Delete", "échap": "Escape", "echap": "Escape", "esc": "Escape",
    "tab": "Tab", "retour arrière": "Backspace", "retour": "Backspace", "début": "Home", "debut": "Home", "fin": "End", "pg.suiv": "PageDown", "page suivante": "PageDown", "pgsuiv": "PageDown", "pg.préc": "PageUp", "page précédente": "PageUp", "pgprec": "PageUp",
    "flèche bas": "ArrowDown", "fleche bas": "ArrowDown", "↓": "ArrowDown", "flèche haut": "ArrowUp", "fleche haut": "ArrowUp", "↑": "ArrowUp", "flèche gauche": "ArrowLeft", "fleche gauche": "ArrowLeft", "←": "ArrowLeft", "flèche droite": "ArrowRight", "fleche droite": "ArrowRight", "→": "ArrowRight" };
  function parseCombo(s) {
    const p = splitKeys(s), o = { ctrl: false, shift: false, alt: false, key: "" };
    p.forEach(x => { const l = x.toLowerCase();
      if (l === "ctrl" || l === "cmd" || l === "⌘") o.ctrl = true; else if (l === "maj" || l === "shift" || l === "⇧") o.shift = true; else if (l === "alt" || l === "option" || l === "⌥") o.alt = true;
      else o.key = KEYNAME[l] || (x.length === 1 ? x.toUpperCase() : /^f\d{1,2}$/i.test(x) ? x.toUpperCase() : x); });
    return o;
  }
  const CODEKEY = c => { let m; if ((m = /^Key([A-Z])$/.exec(c))) return m[1]; if ((m = /^(?:Digit|Numpad)(\d)$/.exec(c))) return m[1];
    return { NumpadAdd: "+", Equal: "=", Semicolon: ";", Minus: "-", NumpadSubtract: "-", Comma: ",", Period: ".", Slash: "/", Quote: "'", Backquote: "`", BracketLeft: "[", BracketRight: "]" }[c] || ""; };
  function comboMatch(e, o) {
    const ctrl = e.ctrlKey || e.metaKey;
    const k = e.key.length === 1 ? e.key.toUpperCase() : e.key;
    const named = o.key.length > 1;
    const symbol = !named && !/^[A-Z0-9]$/.test(o.key);
    if (ctrl !== o.ctrl || e.altKey !== o.alt) return false;
    if (!symbol && e.shiftKey !== o.shift) return false;
    if (named) return e.key === o.key || (o.key.startsWith("F") && e.key.toUpperCase() === o.key);
    return k === o.key || CODEKEY(e.code) === o.key || (o.key === "+" && (e.key === "+" || e.code === "NumpadAdd" || e.code === "Equal"));
  }
  const isMod = k => ["Control", "Shift", "Alt", "Meta", "AltGraph", "CapsLock"].includes(k);
  function keysLab(el, lab, onTask) {
    const seq = lab.setup.sequence || [];
    let i = 0, t0 = 0, errs = 0, miss = 0, timer = null, finished = false;
    el.innerHTML = `<div class="o-kl" tabindex="0" aria-label="Zone de capture du clavier">
      <div class="o-kl-top"><span>⏱ <b class="o-kl-time">0,0 s</b></span><span>🎯 <b class="o-kl-acc">100 %</b></span><span>🏁 <b class="o-kl-n">0/${seq.length}</b></span>${P.keys[lab.id] ? `<span>🥇 record <b>${(P.keys[lab.id] / 1000).toFixed(1).replace(".", ",")} s</b></span>` : ""}</div>
      <div class="o-kl-card"><small class="o-kl-step"></small><h3 class="o-kl-desc"></h3><div class="o-kl-keys"></div><p class="o-kl-msg">Clique ici puis tape le raccourci au clavier</p></div>
      <div class="o-kl-foot"><button type="button" class="btn sm ghost" data-kshow>👁 Montrer</button><button type="button" class="btn sm ghost" data-kskip>Passer →</button><button type="button" class="btn sm ghost" data-krestart>↺ Recommencer</button></div>
      <div class="o-kl-track">${seq.map((s, k) => `<i data-k="${k}" title="${esc(s.desc)}"></i>`).join("")}</div>
    </div>`;
    const z = el.querySelector(".o-kl"), $ = s => el.querySelector(s);
    function show(reveal) {
      const s = seq[i];
      $(".o-kl-step").textContent = `Raccourci ${i + 1} sur ${seq.length}`;
      $(".o-kl-desc").textContent = s.desc;
      $(".o-kl-keys").innerHTML = splitKeys(s.keys).map(k => `<kbd class="${reveal ? "rev" : ""}">${reveal ? esc(k) : "?"}</kbd>`).join("<i>+</i>");
      el.querySelectorAll(".o-kl-track i").forEach((d, k) => d.classList.toggle("cur", k === i));
      miss = 0;
    }
    function tick() { if (!t0 || finished) return; $(".o-kl-time").textContent = ((Date.now() - t0) / 1000).toFixed(1).replace(".", ",") + " s"; }
    function stats() { const n = i + errs; $(".o-kl-acc").textContent = (n ? Math.round(i / n * 100) : 100) + " %"; $(".o-kl-n").textContent = `${i}/${seq.length}`; }
    function next(ok) {
      el.querySelector(`.o-kl-track i[data-k="${i}"]`).classList.add(ok ? "ok" : "skip");
      onTask(i, ok);
      $(".o-kl-keys").querySelectorAll("kbd").forEach((k, n) => { k.textContent = splitKeys(seq[i].keys)[n]; k.classList.add(ok ? "hit" : "rev"); });
      i++; stats();
      if (i >= seq.length) return setTimeout(finish, 450);
      setTimeout(() => show(false), ok ? 420 : 700);
    }
    function finish() {
      finished = true; clearInterval(timer);
      const ms = Date.now() - t0, acc = Math.round(seq.length / (seq.length + errs) * 100);
      const rec = !P.keys[lab.id] || ms < P.keys[lab.id];
      if (rec) { P.keys[lab.id] = ms; save(); }
      $(".o-kl-card").innerHTML = `<div class="o-kl-end"><div class="o-win-badge">⌨️</div><h3>${(ms / 1000).toFixed(1).replace(".", ",")} s · ${acc} % de précision</h3><p>${rec ? "🥇 Nouveau record personnel !" : "Ton record : " + (P.keys[lab.id] / 1000).toFixed(1).replace(".", ",") + " s"}</p></div>`;
      confetti(z);
    }
    function onKey(e) {
      if (isMod(e.key)) return;
      e.preventDefault(); e.stopPropagation();
      if (finished || i >= seq.length) return;
      if (!t0) { t0 = Date.now(); timer = setInterval(tick, 100); }
      if (comboMatch(e, parseCombo(seq[i].keys))) { z.classList.remove("bad"); void z.offsetWidth; z.classList.add("good"); setTimeout(() => z.classList.remove("good"), 400); next(true); }
      else {
        errs++; miss++; stats();
        z.classList.remove("bad"); void z.offsetWidth; z.classList.add("bad");
        const got = [(e.ctrlKey || e.metaKey) && "Ctrl", e.shiftKey && "Maj", e.altKey && "Alt", e.key.length === 1 ? e.key.toUpperCase() : e.key].filter(Boolean).join("+");
        $(".o-kl-msg").textContent = `Tu as tapé ${got}${miss >= 2 ? " — voici la réponse 👇" : " — réessaie !"}`;
        if (miss >= 2) show(true);
      }
    }
    z.addEventListener("keydown", onKey);
    z.addEventListener("focus", () => { z.classList.add("on"); if (!finished) $(".o-kl-msg").textContent = "À toi : tape le raccourci (le navigateur est neutralisé ici)"; });
    z.addEventListener("blur", () => { z.classList.remove("on"); if (!finished) $(".o-kl-msg").textContent = "Clique ici puis tape le raccourci au clavier"; });
    el.addEventListener("click", e => {
      if (e.target.closest("[data-kshow]")) { show(true); z.focus(); }
      else if (e.target.closest("[data-kskip]")) { if (!t0) { t0 = Date.now(); timer = setInterval(tick, 100); } if (!finished) { errs++; next(false); } z.focus(); }
      else if (e.target.closest("[data-krestart]")) { const host = el; destroy(); keysLab(host, lab, onTask); host.querySelector(".o-kl").focus(); }
      else z.focus();
    });
    function destroy() { clearInterval(timer); }
    show(false);
    return { destroy };
  }

  /* ================= mode présentation ================= */
  function present(f) {
    const { app: a, m } = f, A = APPS[a];
    const slides = [
      `<div class="o-ps-cover">${appLogo(a, 110)}<span class="kicker">${A.name} · ${LEVELS[m.level].icon} Niveau ${m.level} ${LEVELS[m.level].name}</span><h1>${m.icon} ${esc(m.title)}</h1><p>${esc(m.summary || "")}</p><small>⏱ ${m.duration || 45} min · ${m.lessons.length} leçons · ${m.labs.length} lab(s)</small></div>`,
      `<h2>🎯 Objectifs</h2><ul class="o-ps-list">${(m.objectives || []).map((o, i) => `<li style="--d:${i * 120}ms">${esc(o)}</li>`).join("")}</ul>`,
      ...m.lessons.map((l, i) => `<small class="o-ps-k">Leçon ${i + 1} / ${m.lessons.length}</small><h2>${esc(l.title)}</h2><div class="o-content">${l.html}</div>`),
      (m.shortcuts || []).length ? `<h2>⌨️ Raccourcis à retenir</h2><div class="o-ps-keys">${m.shortcuts.map((s, i) => `<div style="--d:${i * 80}ms">${keysHTML(s.keys)}<span>${esc(s.desc)}</span></div>`).join("")}</div>` : "",
      `<h2>🧪 Place à la pratique</h2><div class="o-ps-labs">${m.labs.map((l, i) => `<div style="--d:${i * 140}ms"><small>${KIND[l.kind]}</small><b>${esc(l.title)}</b><div class="o-content">${l.scenario || ""}</div></div>`).join("")}</div>`,
      `<div class="o-ps-cover"><h1>✅ Quiz</h1><p>${m.quiz.length} questions pour valider le module.</p><small>Ouvre l'onglet Quiz sur chaque poste</small></div>`,
    ].filter(Boolean);
    let i = 0, t0 = Date.now(), laser = false, black = false;
    const o = document.createElement("div"); o.className = "o-present"; o.dataset.app = a; o.style.cssText = `--a1:${A.c1};--a2:${A.c2}`;
    o.innerHTML = `<div class="o-ps-slide"></div><div class="o-ps-bar"><i></i></div>
      <div class="o-ps-ui"><span class="o-ps-n"></span><span class="o-ps-timer">00:00</span><span class="o-ps-help"><kbd>←</kbd><kbd>→</kbd> naviguer · <kbd>L</kbd> laser · <kbd>B</kbd> écran noir · <kbd>F</kbd> plein écran · <kbd>Échap</kbd> quitter</span><button type="button" data-x>✕</button></div>
      <div class="o-ps-laser" hidden></div><div class="o-ps-black" hidden></div>`;
    document.body.append(o); document.body.classList.add("o-presenting");
    const sl = o.querySelector(".o-ps-slide"), laserEl = o.querySelector(".o-ps-laser");
    function show(n, dir) {
      i = Math.max(0, Math.min(slides.length - 1, n));
      sl.className = "o-ps-slide" + (reduce ? "" : dir < 0 ? " in-l" : " in-r");
      sl.innerHTML = `<div class="o-ps-in">${slides[i]}</div>`;
      enhance(sl);
      o.querySelector(".o-ps-n").textContent = `${i + 1} / ${slides.length}`;
      o.querySelector(".o-ps-bar i").style.width = ((i + 1) / slides.length * 100) + "%";
    }
    const timer = setInterval(() => { const s = Math.floor((Date.now() - t0) / 1000); o.querySelector(".o-ps-timer").textContent = String(Math.floor(s / 60)).padStart(2, "0") + ":" + String(s % 60).padStart(2, "0"); }, 500);
    function key(e) {
      const k = e.key;
      if (["ArrowRight", " ", "PageDown", "Enter"].includes(k)) { e.preventDefault(); show(i + 1, 1); }
      else if (["ArrowLeft", "PageUp", "Backspace"].includes(k)) { e.preventDefault(); show(i - 1, -1); }
      else if (k === "Home") show(0, -1); else if (k === "End") show(slides.length - 1, 1);
      else if (k === "Escape") close();
      else if (k === "l" || k === "L") { laser = !laser; laserEl.hidden = !laser; o.classList.toggle("laser", laser); }
      else if (k === "b" || k === "B" || k === ".") { black = !black; o.querySelector(".o-ps-black").hidden = !black; }
      else if (k === "f" || k === "F") { document.fullscreenElement ? document.exitFullscreen() : o.requestFullscreen && o.requestFullscreen().catch(() => {}); }
      else if (k === "t" || k === "T") t0 = Date.now();
      e.stopPropagation();
    }
    const move = e => { if (laser) laserEl.style.transform = `translate(${e.clientX}px,${e.clientY}px)`; };
    function close() { clearInterval(timer); removeEventListener("keydown", key, true); removeEventListener("pointermove", move); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); o.remove(); document.body.classList.remove("o-presenting"); }
    addEventListener("keydown", key, true); addEventListener("pointermove", move, { passive: true });
    o.addEventListener("click", e => { if (e.target.closest("[data-x]")) return close(); if (e.target.closest("a, button, pre")) return; show(i + (e.clientX > innerWidth / 3 ? 1 : -1), e.clientX > innerWidth / 3 ? 1 : -1); });
    show(0, 1);
    if (o.requestFullscreen) o.requestFullscreen().catch(() => {});
  }

  /* ================= exports ================= */
  const DOC_CSS = `
    *{box-sizing:border-box}body{font:15px/1.6 "Segoe UI",Inter,Arial,sans-serif;color:#1f2937;margin:0;background:#f4f6fb}
    .page{max-width:900px;margin:0 auto;padding:40px 44px;background:#fff}
    header.cover{padding:60px 44px;color:#fff;background:linear-gradient(135deg,var(--a1),var(--a2));max-width:900px;margin:0 auto}
    header.cover h1{font-size:40px;line-height:1.1;margin:10px 0}header.cover .k{letter-spacing:.14em;text-transform:uppercase;font-size:12px;font-weight:800;opacity:.9}
    .logo{display:inline-grid;place-items:center;width:58px;height:58px;border-radius:16px;background:rgba(255,255,255,.18);font:900 30px Georgia,serif;border:2px solid rgba(255,255,255,.5)}
    h2{font-size:26px;margin:46px 0 6px;color:var(--a1);border-bottom:3px solid var(--a1);padding-bottom:6px;page-break-before:always}
    h2.first{page-break-before:avoid}h3{font-size:19px;margin:28px 0 8px}h4{margin:18px 0 6px}
    .meta{color:#64748b;font-size:13px}.lvl{display:inline-block;padding:2px 10px;border-radius:99px;background:#eef2ff;font-weight:700;font-size:12px}
    .box{border:1px solid #e5e7eb;border-radius:12px;padding:14px 18px;margin:14px 0;background:#fafbff}
    .callout{border-left:4px solid #3b82f6;background:#eff6ff;padding:10px 14px;border-radius:8px;margin:12px 0}.callout.ok{border-color:#16a34a;background:#f0fdf4}
    .callout.warn{border-color:#d97706;background:#fffbeb}.callout.bad{border-color:#dc2626;background:#fef2f2}
    .o-tip{background:#fefce8;border:1px dashed #eab308;padding:8px 12px;border-radius:8px;margin:10px 0}
    .o-path{display:inline-block;background:#f1f5f9;border-radius:8px;padding:4px 10px;font-weight:600;font-size:13px;margin:6px 0}
    kbd{display:inline-block;border:1px solid #cbd5e1;border-bottom-width:3px;border-radius:6px;padding:0 6px;font:600 12px/1.6 Consolas,monospace;background:#fff}
    pre.code,code{font-family:Consolas,"JetBrains Mono",monospace}pre.code{background:#0f172a;color:#e2e8f0;padding:12px 14px;border-radius:10px;white-space:pre-wrap}
    code{background:#f1f5f9;padding:1px 5px;border-radius:4px}
    table{border-collapse:collapse;width:100%;margin:10px 0}th,td{border:1px solid #e2e8f0;padding:6px 10px;text-align:left;vertical-align:top}th{background:#f8fafc}
    .tasks li{margin:6px 0}.sol{background:#f0fdf4;border:1px solid #bbf7d0;border-radius:12px;padding:10px 18px}
    .q{margin:12px 0}.q ol{margin:6px 0}.q .ans{color:#15803d;font-weight:700}.toc a{color:inherit;text-decoration:none}.toc li{margin:4px 0}
    .tick{display:inline-block;width:14px;height:14px;border:2px solid #94a3b8;border-radius:4px;margin-right:8px;vertical-align:-2px}
    footer{color:#94a3b8;font-size:12px;text-align:center;padding:30px}
    @media print{body{background:#fff}.page{padding:0 8mm}header.cover{-webkit-print-color-adjust:exact;print-color-adjust:exact}h2{page-break-before:always}}`;
  function moduleHTML(a, m, sol) {
    const A = APPS[a];
    return `<h2 id="${m.id}">${m.icon} ${esc(m.title)}</h2>
      <p class="meta"><span class="lvl">${LEVELS[m.level].icon} Niveau ${m.level} · ${LEVELS[m.level].name}</span> · ${m.id.toUpperCase()} · ${m.duration || 45} min</p>
      <p><i>${esc(m.summary || "")}</i></p>
      <div class="box"><b>Objectifs</b><ul>${(m.objectives || []).map(o => `<li>${esc(o)}</li>`).join("")}</ul></div>
      ${m.lessons.map((l, i) => `<h3>${i + 1}. ${esc(l.title)}</h3>${l.html}`).join("")}
      ${(m.shortcuts || []).length ? `<h3>Raccourcis clavier</h3><table><tr><th>Raccourci</th><th>Action</th></tr>${m.shortcuts.map(s => `<tr><td>${splitKeys(s.keys).map(k => `<kbd>${esc(k)}</kbd>`).join(" + ")}</td><td>${esc(s.desc)}</td></tr>`).join("")}</table>` : ""}
      ${m.labs.map(l => `<h3>🧪 ${esc(l.title)} <span class="meta">(${KIND[l.kind] || l.kind})</span></h3>${l.scenario || ""}
        <ol class="tasks">${l.kind === "keys" ? (l.setup.sequence || []).map(s => `<li><span class="tick"></span>${esc(s.desc)}${sol ? ` → ${splitKeys(s.keys).map(k => `<kbd>${esc(k)}</kbd>`).join(" + ")}` : ""}</li>`).join("") : (l.tasks || []).map(t => `<li><span class="tick"></span>${t.text}${sol && t.hint ? `<div class="meta">💡 ${t.hint}</div>` : ""}</li>`).join("")}</ol>
        ${l.kind === "sheet" ? sheetTable(l.setup) : ""}${l.kind === "doc" && l.setup.html ? `<div class="box"><b>Document de départ</b>${l.setup.html}</div>` : ""}
        ${sol && l.solution ? `<div class="sol"><b>Corrigé</b>${l.solution}</div>` : ""}`).join("")}
      <h3>✅ Quiz</h3>${m.quiz.map((q, i) => `<div class="q"><b>${i + 1}. ${esc(q.q)}</b><ol type="A">${q.choices.map((c, k) => `<li class="${sol && k === q.answer ? "ans" : ""}">${esc(c)}${sol && k === q.answer ? " ✓" : ""}</li>`).join("")}</ol>${sol ? `<div class="meta">${esc(q.explain || "")}</div>` : ""}</div>`).join("")}`;
  }
  function sheetTable(s) {
    const d = s.data || {}; let mc = 0, mr = 0;
    Object.keys(d).forEach(k => { const m = /^([A-Z]+)(\d+)$/i.exec(k); if (!m) return; mc = Math.max(mc, m[1].toUpperCase().charCodeAt(0) - 64); mr = Math.max(mr, +m[2]); });
    if (!mc) return "";
    let h = `<table style="font-size:12px"><tr><th></th>${Array.from({ length: mc }, (_, c) => `<th>${String.fromCharCode(65 + c)}</th>`).join("")}</tr>`;
    for (let r = 1; r <= mr; r++) h += `<tr><th>${r}</th>${Array.from({ length: mc }, (_, c) => { const v = d[String.fromCharCode(65 + c) + r]; return `<td>${v == null ? "" : esc(typeof v === "number" ? String(v).replace(".", ",") : v)}</td>`; }).join("")}</tr>`;
    return `<div class="box"><b>Données de départ${s.sheetName ? " — " + esc(s.sheetName) : ""}</b>${h}</table></div>`;
  }
  function handbook(apps, opts = {}) {
    const sol = !!opts.solutions, one = opts.module;
    const list = one ? [one] : apps.flatMap(a => DATA[a].map(m => ({ app: a, m })));
    const A = APPS[list[0] ? list[0].app : "word"];
    const title = one ? `${APPS[one.app].name} — ${one.m.title}` : apps.length === 1 ? `Formation ${APPS[apps[0]].name}` : "Formation Microsoft Office";
    return `<!doctype html><html lang="fr"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}${sol ? " (formateur)" : ""}</title>
      <style>:root{--a1:${A.c1};--a2:${A.c2}}${DOC_CSS}</style></head><body>
      <header class="cover"><span class="logo">${apps.length === 1 || one ? A.letter : "O"}</span><div class="k">Office Academy · ${sol ? "Support formateur avec corrigés" : "Support stagiaire"}</div><h1>${esc(title)}</h1>
      <div>${list.length} module${list.length > 1 ? "s" : ""} · ${list.reduce((t, x) => t + x.m.labs.length, 0)} labs · ${list.reduce((t, x) => t + x.m.quiz.length, 0)} questions · généré le ${new Date().toLocaleDateString("fr-FR")}</div></header>
      <div class="page">
      ${list.length > 1 ? `<div class="box toc"><b>Sommaire</b><ol>${list.map(x => `<li><a href="#${x.m.id}">${x.m.icon} ${APPS[x.app].name} · ${esc(x.m.title)}</a> <span class="meta">— ${LEVELS[x.m.level].name}</span></li>`).join("")}</ol></div>` : ""}
      ${list.map((x, i) => moduleHTML(x.app, x.m, sol).replace(/^<h2 /, i === 0 && list.length === 1 ? '<h2 class="first" ' : "<h2 ")).join("")}
      </div><footer>Office Academy — ${esc(PAN.SITE_URL || "")}pages/office.html · Microsoft, Word, Excel et PowerPoint sont des marques de Microsoft Corporation.</footer></body></html>`;
  }
  const wordDoc = (title, body) => `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta charset="utf-8"><title>${esc(title)}</title>
    <!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->
    <style>body{font-family:Calibri,Arial,sans-serif;font-size:11pt}h1{font-size:20pt;color:#2B579A}h2{font-size:16pt;color:#2B579A}h3{font-size:13pt;color:#1F3864}table{border-collapse:collapse}td,th{border:1px solid #999;padding:4pt}</style></head><body>${body}</body></html>`;
  function slidesRTF(setup) {
    const u = s => String(s || "").replace(/[\\{}]/g, "\\$&").replace(/[^\x20-\x7e]/g, ch => "\\u" + (ch.charCodeAt(0) > 32767 ? ch.charCodeAt(0) - 65536 : ch.charCodeAt(0)) + "?");
    const p = (lvl, t) => `\\pard\\s${lvl + 1}\\outlinelevel${lvl} ${u(t)}\\par\n`;
    let body = "";
    (setup.slides || []).forEach(s => {
      body += p(0, s.title || "(Diapositive)");
      [s.subtitle, s.body, s.left, s.right].filter(Boolean).forEach(b => String(b).split("\n").filter(Boolean).forEach(line => { body += p(1, line); }));
    });
    return `{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Calibri;}}{\\stylesheet{\\s1\\outlinelevel0 heading 1;}{\\s2\\outlinelevel1 heading 2;}}\n${body}}`;
  }
  function starterOf(lab, f) {
    if (lab.kind === "steps") return lab.setup && lab.setup.file ? true : null;
    if (lab.kind === "doc") return !!(lab.setup && lab.setup.html);
    if (lab.kind === "slides") return !!(lab.setup && lab.setup.slides);
    if (lab.kind === "sheet") return !!SIMS.sheet;
    return null;
  }
  function starterFile(lab, f) {
    const base = slug(lab.id + "-" + lab.title);
    if (lab.kind === "steps" && lab.setup.file) {
      const fl = lab.setup.file, t = fl.type || fl.name.split(".").pop();
      if (t === "doc") return { name: fl.name.replace(/\.\w+$/, "") + ".doc", content: wordDoc(fl.name, fl.content), type: "application/msword" };
      return { name: fl.name, content: (t === "csv" ? "\ufeff" : "") + fl.content, type: t === "csv" ? "text/csv;charset=utf-8" : "text/plain;charset=utf-8" };
    }
    if (lab.kind === "doc") return { name: base + ".doc", content: wordDoc(lab.title, lab.setup.html), type: "application/msword" };
    if (lab.kind === "slides") return { name: base + "-plan.rtf", content: slidesRTF(lab.setup), type: "application/rtf" };
    if (lab.kind === "sheet" && SIMS.sheet) {
      const tmp = document.createElement("div"); tmp.style.cssText = "position:fixed;left:-9999px;top:0;width:900px;visibility:hidden"; document.body.append(tmp);
      let out = null;
      try { const s = SIMS.sheet(tmp, clone(lab.setup), { changed() {}, toast() {}, esc, confetti() {} }); out = s.exportFile(true); s.destroy && s.destroy(); } catch (e) { console.error(e); }
      tmp.remove();
      if (out) out.name = base + ".csv";
      return out;
    }
    return null;
  }
  function downloadStarter(lab, f) {
    const x = starterFile(lab, f); if (!x) return toast("Pas de fichier de départ pour ce lab");
    download(x.name, x.content, x.type);
    toast(lab.kind === "slides" ? "Plan .rtf : PowerPoint › Nouvelle diapositive › Diapositives à partir d'un plan" : `« ${x.name} » téléchargé ⬇`);
  }
  /* ZIP minimal (stockage sans compression), noms UTF-8 */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = ~0; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return ~c >>> 0; };
  function zip(files) {
    const enc = new TextEncoder(), parts = [], cd = []; let off = 0;
    const d = new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const f of files) {
      const name = enc.encode(f.name), data = typeof f.content === "string" ? enc.encode(f.content) : f.content, crc = crc32(data);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true); h.setUint16(10, time, true); h.setUint16(12, date, true);
      h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, name.length, true); h.setUint16(28, 0, true);
      parts.push(h.buffer, name, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 0, true); c.setUint16(12, time, true); c.setUint16(14, date, true);
      c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, name.length, true); c.setUint32(42, off, true);
      cd.push(c.buffer, name);
      off += 30 + name.length + data.length;
    }
    const cdSize = cd.reduce((t, x) => t + (x.byteLength || x.length), 0);
    const e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
    return new Blob([...parts, ...cd, e.buffer], { type: "application/zip" });
  }
  function trainerKit(apps) {
    const files = [];
    const readme = `OFFICE ACADEMY — KIT FORMATEUR\r\nGénéré le ${new Date().toLocaleString("fr-FR")}\r\n\r\nContenu :\r\n- Supports-stagiaires/ : cours, énoncés des labs et quiz (sans réponses) — à distribuer ou imprimer\r\n- Supports-formateur/ : mêmes supports AVEC corrigés des labs et réponses des quiz\r\n- Exercices/ : fichiers de départ des labs (.csv pour Excel, .doc pour Word, .rtf = plan importable dans PowerPoint)\r\n\r\nOuvre les fichiers .html dans n'importe quel navigateur (hors ligne). Ctrl+P pour imprimer ou enregistrer en PDF.\r\nVersion interactive (labs simulés, quiz, mode présentation) : ${(PAN.SITE_URL || "") + "pages/office.html"}\r\nSite complet hors ligne : ${PAN.ZIP_URL || ""}\r\n`;
    files.push({ name: "LISEZMOI.txt", content: readme });
    apps.forEach(a => {
      const N = APPS[a].name;
      files.push({ name: `Supports-stagiaires/Formation-${N}.html`, content: handbook([a]) });
      files.push({ name: `Supports-formateur/Formation-${N}-corriges.html`, content: handbook([a], { solutions: true }) });
      DATA[a].forEach(m => m.labs.forEach(l => { const x = starterOf(l) && starterFile(l, { app: a, m }); if (x) files.push({ name: `Exercices/${N}/${x.name}`, content: x.content }); }));
    });
    return zip(files);
  }

  /* ================= vue : export & formateur ================= */
  function viewExport() {
    const n = counts(), xp = totalXP(), r = rankOf(xp);
    const done = allMods().filter(x => modStats(x.m).done);
    return `
    <section class="o-apphero o-exphero">
      <div class="wrap">${crumbs([["Espace formateur & export"]])}
        <div class="o-apphero-t"><span class="o-exp-ic">📦</span><div><span class="kicker">Hors ligne · Impression · Animation de session</span><h1>Espace formateur & export</h1></div></div>
        <p class="lead">Tout ce qu'il faut pour animer une session, même sans réseau : supports à distribuer, corrigés, fichiers d'exercice, site complet et attestations.</p>
      </div>
    </section>
    <section class="o-sec"><div class="wrap">
      <div class="o-exp-grid" data-stagger="80">
        <div class="card o-exp o-exp-hero glow-border">
          <div class="o-exp-i">🧳</div><h3>Kit formateur complet (.zip)</h3>
          <p class="mut">Supports stagiaires + supports formateur avec corrigés + tous les fichiers d'exercice des ${n.labs} labs, rangés par application.</p>
          <div class="o-exp-apps">${Object.values(APPS).map(A => `<label><input type="checkbox" value="${A.id}" checked data-kit-app> ${A.name}</label>`).join("")}</div>
          <button class="btn primary" type="button" data-kit>⬇ Générer le kit .zip</button>
        </div>
        <div class="card o-exp"><div class="o-exp-i">🌐</div><h3>Site complet hors ligne</h3><p class="mut">Toute la plateforme interactive (labs simulés, quiz, mode présentation). Dézippe puis ouvre <b>pages/office.html</b> — aucune installation, aucun serveur.</p>
          <a class="btn primary" href="${esc(PAN.ZIP_URL || "#")}">⬇ Télécharger le site (.zip)</a></div>
        <div class="card o-exp"><div class="o-exp-i">📄</div><h3>Supports de cours HTML</h3><p class="mut">Un fichier autonome par application, imprimable (Ctrl+P → PDF).</p>
          <div class="o-exp-rows">${Object.values(APPS).map(A => `<div><b style="color:${A.c2}">${A.name}</b><button class="btn sm ghost" type="button" data-exp-app="${A.id}">Stagiaire</button><button class="btn sm ghost" type="button" data-exp-app="${A.id}" data-sol="1">Formateur 🔑</button></div>`).join("")}
          <div><b>Les 3</b><button class="btn sm ghost" type="button" data-exp-all>Stagiaire</button><button class="btn sm ghost" type="button" data-exp-all data-sol="1">Formateur 🔑</button></div></div></div>
        <div class="card o-exp"><div class="o-exp-i">🗂️</div><h3>Fichiers d'exercice</h3><p class="mut">Fichiers de départ lab par lab.</p>
          <select data-starter-sel class="o-select">${allMods().map(({ app: a, m }) => m.labs.filter(l => starterOf(l)).map(l => `<option value="${l.id}">${APPS[a].name} · ${m.id.toUpperCase()} · ${esc(l.title)}</option>`).join("")).join("")}</select>
          <button class="btn sm primary" type="button" data-starter-go>⬇ Télécharger</button></div>
        <div class="card o-exp"><div class="o-exp-i">🎓</div><h3>Mode formateur</h3><p class="mut">Met en évidence la bonne réponse des quiz pendant l'animation et rappelle les corrigés. À désactiver avant de laisser la main aux stagiaires.</p>
          <label class="o-switch"><input type="checkbox" data-trainer ${P.trainer ? "checked" : ""}><span></span> Mode formateur ${P.trainer ? "activé" : "désactivé"}</label>
          <p class="small mut">Astuce : dans chaque module, le bouton 🎤 <b>Présenter</b> projette le cours en plein écran (laser <kbd>L</kbd>, écran noir <kbd>B</kbd>).</p></div>
        <div class="card o-exp"><div class="o-exp-i">💾</div><h3>Progression</h3><p class="mut">${r.icon} ${r.name} · <b>${xp}</b> XP · ${done.length}/${n.mods} modules validés. Sauvegardée dans ce navigateur : exporte-la pour changer de poste.</p>
          <div class="o-cta"><button class="btn sm ghost" type="button" data-prog-exp>⬇ Exporter</button><label class="btn sm ghost">⬆ Importer<input type="file" accept=".json" data-prog-imp hidden></label><button class="btn sm ghost" type="button" data-prog-reset>🗑 Réinitialiser</button></div></div>
      </div>

      <div class="o-cert-wrap card glass" data-reveal="zoom">
        <div class="o-cert-form">
          <span class="kicker">Attestation</span><h2>Attestation de formation</h2>
          <p class="mut">Génère une attestation imprimable listant les modules validés sur ce poste.</p>
          <label>Nom du stagiaire <input type="text" data-cert-name value="${esc(P.name || "")}" placeholder="Prénom Nom"></label>
          <label>Formateur / organisme <input type="text" data-cert-org placeholder="Optionnel"></label>
          <label>Parcours <select data-cert-app><option value="all">Microsoft Office (les 3)</option>${Object.values(APPS).map(A => `<option value="${A.id}">${A.name}</option>`).join("")}</select></label>
          <div class="o-cta"><button class="btn primary" type="button" data-cert>⬇ Télécharger l'attestation</button></div>
        </div>
        <div class="o-cert-prev" aria-hidden="true"><div class="o-cert-paper"><small>ATTESTATION DE FORMATION</small><b data-cert-prev>${esc(P.name || "Prénom Nom")}</b><span>a suivi la formation Microsoft Office</span><i>${done.length} module(s) validé(s) · ${xp} XP</i><div class="o-cert-seal">★</div></div></div>
      </div>
    </div></section>`;
  }
  function certificate(name, org, app) {
    const list = allMods().filter(x => (app === "all" || x.app === app) && modStats(x.m).done);
    const parts = allMods().filter(x => (app === "all" || x.app === app));
    const A = app === "all" ? { name: "Microsoft Office", c1: "#7C3AED", c2: "#F59E0B" } : APPS[app];
    const hours = Math.round(list.reduce((t, x) => t + (x.m.duration || 45), 0) / 60 * 10) / 10;
    return `<!doctype html><html lang="fr"><head><meta charset="utf-8"><title>Attestation — ${esc(name)}</title><style>
      @page{size:A4 landscape;margin:0}*{box-sizing:border-box}body{margin:0;font-family:Georgia,"Times New Roman",serif;background:#eee}
      .c{width:297mm;height:210mm;margin:0 auto;background:#fff;position:relative;padding:22mm 26mm;overflow:hidden;-webkit-print-color-adjust:exact;print-color-adjust:exact}
      .c:before{content:"";position:absolute;inset:9mm;border:2px solid ${A.c1};border-radius:6mm}.c:after{content:"";position:absolute;inset:12mm;border:1px solid ${A.c2};border-radius:4mm}
      .band{position:absolute;right:-40mm;top:-40mm;width:120mm;height:120mm;border-radius:50%;background:radial-gradient(circle,${A.c2}55,transparent 70%)}
      .k{font:700 12px/1 Arial,sans-serif;letter-spacing:.3em;color:${A.c1};text-transform:uppercase}h1{font-size:44px;margin:8mm 0 4mm;color:#111}
      .n{font-size:38px;font-style:italic;color:${A.c1};border-bottom:1px solid #ccc;display:inline-block;padding:0 10mm 2mm;margin:4mm 0}
      p{font-size:16px;color:#333;max-width:220mm;line-height:1.5}ul{columns:2;font:13px/1.5 Arial,sans-serif;color:#444;margin:4mm 0;padding-left:5mm;max-height:52mm;overflow:hidden}
      .f{position:absolute;bottom:22mm;left:26mm;right:26mm;display:flex;justify-content:space-between;font:13px Arial,sans-serif;color:#555}
      .seal{position:absolute;right:30mm;bottom:34mm;width:34mm;height:34mm;border-radius:50%;background:conic-gradient(${A.c1},${A.c2},${A.c1});display:grid;place-items:center;color:#fff;font:900 30px Arial;box-shadow:0 0 0 3mm #fff,0 0 0 3.6mm ${A.c1}}
      @media print{body{background:#fff}}</style></head><body><div class="c"><div class="band"></div>
      <div class="k">Office Academy · Attestation de formation</div><h1>Attestation</h1>
      <p>Nous attestons que</p><div class="n">${esc(name || "—")}</div>
      <p>a suivi la formation <b>${esc(A.name)}</b> et validé <b>${list.length} module(s) sur ${parts.length}</b> (environ ${String(hours).replace(".", ",")} h de cours, labs pratiques et quiz)${list.length ? " :" : "."}</p>
      <ul>${list.map(x => `<li>${APPS[x.app].name} · ${LEVELS[x.m.level].name} — ${esc(x.m.title)}</li>`).join("")}</ul>
      <div class="seal">★</div>
      <div class="f"><span>Fait le ${new Date().toLocaleDateString("fr-FR", { day: "numeric", month: "long", year: "numeric" })}</span><span>${org ? "Formateur / organisme : " + esc(org) : "Signature du formateur : ____________________"}</span></div>
      </div><script>setTimeout(function(){try{print()}catch(e){}},400)<\/script></body></html>`;
  }
  function afterExport() {
    const $ = s => root.querySelector(s);
    root.addEventListener("click", expClick);
    $("[data-trainer]").onchange = e => { P.trainer = e.target.checked; save(); e.target.parentElement.lastChild.textContent = " Mode formateur " + (P.trainer ? "activé" : "désactivé"); toast(P.trainer ? "Mode formateur activé 🎓" : "Mode formateur désactivé"); };
    $("[data-prog-imp]").onchange = e => {
      const fl = e.target.files[0]; if (!fl) return;
      fl.text().then(t => { const d = JSON.parse(t); if (!d || typeof d !== "object" || !d.labs) throw new Error(); P = Object.assign(blank(), d); save(); toast("Progression importée ✓"); route(); }).catch(() => toast("Fichier de progression invalide"));
    };
    $("[data-cert-name]").oninput = e => { $("[data-cert-prev]").textContent = e.target.value || "Prénom Nom"; };
  }
  function expClick(e) {
    const t = e.target.closest("button"); if (!t || !root.contains(t)) return;
    if (t.hasAttribute("data-kit")) {
      const apps = [...root.querySelectorAll("[data-kit-app]:checked")].map(i => i.value);
      if (!apps.length) return toast("Choisis au moins une application");
      t.disabled = true; t.textContent = "⏳ Génération…";
      setTimeout(() => { try { download(`Office-Academy-kit-formateur-${new Date().toISOString().slice(0, 10)}.zip`, trainerKit(apps)); toast("Kit formateur généré ⬇"); confetti(t); } catch (err) { console.error(err); toast("Erreur pendant la génération du kit"); } t.disabled = false; t.textContent = "⬇ Générer le kit .zip"; }, 30);
    }
    if (t.hasAttribute("data-exp-all")) { const sol = !!t.dataset.sol; download(`Formation-Office${sol ? "-formateur" : ""}.html`, handbook(Object.keys(APPS), { solutions: sol }), "text/html;charset=utf-8"); toast("Support téléchargé ⬇"); }
    if (t.hasAttribute("data-starter-go")) { const id = root.querySelector("[data-starter-sel]").value; for (const { app: a, m } of allMods()) { const l = m.labs.find(x => x.id === id); if (l) downloadStarter(l, { app: a, m }); } }
    if (t.hasAttribute("data-prog-exp")) download(`office-academy-progression-${new Date().toISOString().slice(0, 10)}.json`, JSON.stringify(P, null, 2), "application/json");
    if (t.hasAttribute("data-prog-reset")) {
      if (t.dataset.armed) { P = blank(); save(); toast("Progression réinitialisée"); route(); return; }
      t.dataset.armed = 1; t.textContent = "⚠ Confirmer ?"; setTimeout(() => { if (t.isConnected) { delete t.dataset.armed; t.textContent = "🗑 Réinitialiser"; } }, 3000);
    }
    if (t.hasAttribute("data-cert")) {
      const name = root.querySelector("[data-cert-name]").value.trim(); if (!name) { toast("Indique le nom du stagiaire"); root.querySelector("[data-cert-name]").focus(); return; }
      P.name = name; save();
      download(`Attestation-${slug(name)}.html`, certificate(name, root.querySelector("[data-cert-org]").value.trim(), root.querySelector("[data-cert-app]").value), "text/html;charset=utf-8");
      toast("Attestation téléchargée : ouvre-la pour l'imprimer ou l'enregistrer en PDF"); confetti(t);
    }
  }
  // boutons d'export présents dans plusieurs vues
  document.addEventListener("click", e => {
    const b = e.target.closest("[data-exp-app], [data-exp-mod]"); if (!b) return;
    if (b.dataset.expMod) { const f = findMod(b.dataset.expMod); download(`${f.m.id}-${slug(f.m.title)}.html`, handbook([f.app], { module: f, solutions: P.trainer }), "text/html;charset=utf-8"); toast(P.trainer ? "Support formateur (avec corrigés) téléchargé ⬇" : "Support du module téléchargé ⬇"); return; }
    const a = b.dataset.expApp, sol = !!b.dataset.sol;
    download(`Formation-${APPS[a].name}${sol ? "-formateur" : ""}.html`, handbook([a], { solutions: sol }), "text/html;charset=utf-8"); toast("Support téléchargé ⬇");
  });

  /* ================= démarrage ================= */
  function start() {
    if (!root) return;
    if (PAN.palette) {
      const items = [{ icon: "📦", label: "Office · Espace formateur & export", sub: "Kit .zip, supports, attestation", href: "#/export", group: "Office Academy", kw: "export zip formateur attestation" }];
      Object.values(APPS).forEach(A => {
        items.push({ icon: A.letter, label: `Formation ${A.name}`, sub: A.desc, href: "#/" + A.id, group: "Office Academy", kw: "office formation" });
        DATA[A.id].forEach(m => items.push({ icon: m.icon, label: `${A.name} · ${m.title}`, sub: `${LEVELS[m.level].name} · ${m.lessons.map(l => l.title).join(" · ")}`, href: `#/${A.id}/${m.id}`, group: "Office Academy", kw: strip(m.summary) + " " + (m.shortcuts || []).map(s => s.keys + " " + s.desc).join(" ") }));
      });
      PAN.palette.add(items);
    }
    route();
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start); else setTimeout(start);
})();
