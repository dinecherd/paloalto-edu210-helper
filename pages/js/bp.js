/* Bonnes pratiques — démos animées (pages/bonnes-pratiques.html)
   héros first-match · onglets · simulateur first-match · constructeur de règle · NAT · chaîne d'inspection · logs · checklist */
(function () {
  const $ = (s, r = document) => r.querySelector(s), $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const RM = matchMedia("(prefers-reduced-motion: reduce)").matches;
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const esc = s => PAN.esc(String(s));
  const hooks = {};   // par onglet : { start(), stop() }

  /* ================= héros : la première règle qui correspond ================= */
  (function hero() {
    const rows = $$("#hr-rows .hr-row"), scan = $("#hr-scan"), cap = $("#hr-cap");
    const SEQ = [
      { r: 0, why: "source ∉ EDL", cap: "Règle 1 : la source n'est pas une IP malveillante → règle suivante." },
      { r: 1, why: "ssl ≠ quic", cap: "Règle 2 : l'application est <b>ssl</b>, pas quic → règle suivante." },
      { r: 2, why: "zone LAN ≠ WAN", cap: "Règle 3 : la zone source ne correspond pas → règle suivante." },
      { r: 3, why: "✓ tout correspond", hit: "allow-r", cap: "Règle 4 : <b>tout correspond → allow</b> + profils. La règle 5 n'est <b>jamais lue</b>." },
    ];
    const setCap = h => { cap.classList.remove("pop"); void cap.offsetWidth; cap.innerHTML = h; cap.classList.add("pop"); };
    const place = row => { scan.style.height = row.offsetHeight + "px"; scan.style.transform = `translateY(${row.offsetTop}px)`; };
    function final() {
      SEQ.forEach(s => { rows[s.r].classList.add(s.hit ? "hit" : "miss"); if (s.hit) rows[s.r].classList.add(s.hit); rows[s.r].querySelector(".why").textContent = s.why; });
      rows[4].classList.add("skip");
      cap.innerHTML = SEQ[3].cap;
    }
    if (RM) return final();
    (async function loop() {
      for (;;) {
        rows.forEach(r => { r.className = "hr-row"; r.querySelector(".why").textContent = ""; });
        scan.style.opacity = 0;
        setCap("📦 Un paquet arrive : <b>LAN → WAN</b>, application <b>ssl</b>. Le pare-feu lit les règles de haut en bas…");
        await sleep(2200);
        scan.style.opacity = 1;
        for (const s of SEQ) {
          const row = rows[s.r];
          place(row); await sleep(600);
          row.querySelector(".why").textContent = s.why;
          row.classList.add(s.hit ? "hit" : "miss"); if (s.hit) row.classList.add(s.hit);
          setCap(s.cap);
          if (s.hit) { await sleep(400); rows[4].classList.add("skip"); scan.style.opacity = 0; await sleep(3800); }
          else await sleep(1300);
        }
      }
    })();
  })();

  /* ================= onglets (lien direct : bonnes-pratiques.html#nat) ================= */
  const tabs = $$("#bp-tabs [data-tab]"), ink = $("#bp-ink");
  let cur = null;
  const moveInk = t => { ink.style.width = t.offsetWidth + "px"; ink.style.transform = `translateX(${t.offsetLeft}px)`; };
  function show(id, push) {
    if (!tabs.some(t => t.dataset.tab === id)) id = tabs[0].dataset.tab;
    if (id === cur) return;
    if (cur && hooks[cur] && hooks[cur].stop) hooks[cur].stop();
    cur = id;
    tabs.forEach(t => {
      const on = t.dataset.tab === id;
      t.setAttribute("aria-selected", on); t.tabIndex = on ? 0 : -1;
      $("#tab-" + t.dataset.tab).hidden = !on;
      if (on) { moveInk(t); t.scrollIntoView({ block: "nearest", inline: "nearest" }); }
    });
    if (hooks[id] && hooks[id].start) requestAnimationFrame(() => hooks[id].start());
    PAN.reveal();
    if (push) history.replaceState(null, "", "#" + id);
  }
  tabs.forEach((t, i) => {
    t.addEventListener("click", () => show(t.dataset.tab, true));
    t.addEventListener("keydown", e => {
      if (e.key !== "ArrowRight" && e.key !== "ArrowLeft") return;
      const n = tabs[(i + (e.key === "ArrowRight" ? 1 : -1) + tabs.length) % tabs.length];
      n.focus(); show(n.dataset.tab, true);
    });
  });
  $$("[data-go]").forEach(a => a.addEventListener("click", e => {
    e.preventDefault(); show(a.dataset.go, true);
    $("#bp-tabs").scrollIntoView({ behavior: RM ? "auto" : "smooth", block: "start" });
  }));
  addEventListener("hashchange", () => show(location.hash.slice(1)));
  addEventListener("resize", () => { const t = tabs.find(x => x.dataset.tab === cur); if (t) moveInk(t); });

  /* ================= simulateur first-match ================= */
  (function firstMatch() {
    const RB = [
      { n: "Block-Bad-IP", from: "any", to: "any", src: "EDL-Malveillantes", dst: "any", user: "any", app: "any", act: "deny" },
      { n: "Block-QUIC", from: "LAN", to: "WAN", src: "any", dst: "any", user: "any", app: "quic", act: "deny" },
      { n: "WAN-to-DMZ-Web", from: "WAN", to: "DMZ", src: "any", dst: "PUB-WEB-01", user: "any", app: "web-browsing, ssl", act: "allow" },
      { n: "LAN-Admins-to-DMZ", from: "LAN", to: "DMZ", src: "any", dst: "any", user: "IT-Admins", app: "ssh, ms-rdp", act: "allow" },
      { n: "LAN-to-Internet-Web", from: "LAN", to: "WAN", src: "NET-LAN-Users", dst: "any", user: "any", app: "web-browsing, ssl", act: "allow" },
      { n: "Deny-All-Log", from: "any", to: "any", src: "any", dst: "any", user: "any", app: "any", act: "deny" },
    ];
    const FLOWS = [
      { icon: "🌐", label: "Poste LAN → google.com", from: "LAN", to: "WAN", srcIp: "192.168.1.50", srcIn: ["NET-LAN-Users"], dstIp: "142.250.74.110", dstIn: [], user: "j.martin", groups: ["Utilisateurs"], app: "ssl",
        end: "Autorisé par une règle précise : App-ID (ssl), service application-default et groupe de profils. Le trafic passe, mais il est inspecté." },
      { icon: "⚡", label: "Poste LAN → YouTube en QUIC", from: "LAN", to: "WAN", srcIp: "192.168.1.50", srcIn: ["NET-LAN-Users"], dstIp: "142.250.74.14", dstIn: [], user: "j.martin", groups: ["Utilisateurs"], app: "quic",
        end: "Bloqué volontairement : le navigateur bascule aussitôt en HTTPS/TCP, que le pare-feu sait déchiffrer et filtrer. L'utilisateur ne voit aucune différence." },
      { icon: "🌍", label: "Internet → site web publié", from: "WAN", to: "DMZ", srcIp: "198.51.100.7", srcIn: [], dstIp: "203.0.113.20", dstIn: ["PUB-WEB-01"], user: "unknown", groups: [], app: "ssl",
        end: "La règle utilise l'IP publique (pré-NAT) et la zone DMZ (post-NAT). L'onglet NAT montre pourquoi." },
      { icon: "👩‍💻", label: "Admin IT → SSH serveur DMZ", from: "LAN", to: "DMZ", srcIp: "192.168.1.20", srcIn: ["NET-LAN-Users"], dstIp: "10.10.20.10", dstIn: [], user: "a.durand", groups: ["IT-Admins"], app: "ssh",
        end: "User-ID : la règle autorise un groupe AD, pas une IP. Si l'admin change de poste, l'accès suit." },
      { icon: "🧑‍🎓", label: "Stagiaire → RDP serveur DMZ", from: "LAN", to: "DMZ", srcIp: "192.168.1.77", srcIn: ["NET-LAN-Users"], dstIp: "10.10.20.10", dstIn: [], user: "stagiaire", groups: ["Utilisateurs"], app: "ms-rdp",
        end: "Le stagiaire n'est pas dans IT-Admins : aucune règle ne correspond, il tombe sur Deny-All-Log. Le refus est journalisé." },
      { icon: "🏴‍☠️", label: "Poste LAN → BitTorrent", from: "LAN", to: "WAN", srcIp: "192.168.1.50", srcIn: ["NET-LAN-Users"], dstIp: "91.189.95.21", dstIn: [], user: "j.martin", groups: ["Utilisateurs"], app: "bittorrent",
        end: "App-ID reconnaît BitTorrent même déguisé sur le port 443. L'application n'est autorisée nulle part : refus en fin de rulebase." },
      { icon: "☠️", label: "IP malveillante → serveur", from: "WAN", to: "DMZ", srcIp: "185.220.101.4", srcIn: ["EDL-Malveillantes"], dstIp: "203.0.113.20", dstIn: ["PUB-WEB-01"], user: "unknown", groups: [], app: "ssl",
        end: "Arrêté dès la règle 1, avant même la règle du serveur web. L'EDL est mise à jour automatiquement : rien à maintenir à la main." },
    ];
    function test(r, f) {
      if (r.from !== "any" && r.from !== f.from) return `zone src ${f.from} ≠ ${r.from}`;
      if (r.to !== "any" && r.to !== f.to) return `zone dst ${f.to} ≠ ${r.to}`;
      if (r.src !== "any" && !f.srcIn.includes(r.src)) return `source ∉ ${r.src}`;
      if (r.dst !== "any" && !f.dstIn.includes(r.dst)) return `dst ∉ ${r.dst}`;
      if (r.user !== "any" && !f.groups.includes(r.user)) return `${f.user} ∉ ${r.user}`;
      if (r.app !== "any" && !r.app.split(", ").includes(f.app)) return `appli ${f.app} ∉ règle`;
      return "";
    }
    const board = $("#fm-board"), flowsEl = $("#fm-flows"), pkt = $("#fm-pkt"), out = $("#fm-out");
    board.innerHTML = `<div class="fm-row th"><span>#</span><span>Nom</span><span>Zones</span><span>Source</span><span>Destination</span><span>Utilisateur</span><span>Application</span><span>Action</span></div>`
      + RB.map((r, i) => `<div class="fm-row"><span>${i + 1}</span><span><b>${esc(r.n)}</b></span><span>${r.from} → ${r.to}</span><span>${esc(r.src)}</span><span>${esc(r.dst)}</span><span>${esc(r.user)}</span><span>${esc(r.app)}</span><span><em class="${r.act}">${r.act}</em></span><i class="why"></i></div>`).join("")
      + `<span class="scan" id="fm-scan"></span>`;
    const rows = $$(".fm-row:not(.th)", board), scan = $("#fm-scan");
    flowsEl.innerHTML = FLOWS.map((f, i) => `<button type="button" class="chipb" data-i="${i}">${f.icon} ${esc(f.label)}</button>`).join("");
    let token = 0, started = false;
    async function run(i) {
      const my = ++token, f = FLOWS[i];
      $$(".chipb", flowsEl).forEach(b => b.classList.toggle("on", +b.dataset.i === i));
      rows.forEach(r => { r.className = "fm-row"; r.querySelector(".why").textContent = ""; });
      pkt.innerHTML = `<b>📦</b><span><i>zones</i>${f.from} → ${f.to}</span><span><i>src</i>${f.srcIp}</span><span><i>dst</i>${f.dstIp}</span><span><i>user</i>${esc(f.user)}</span><span><i>app</i>${esc(f.app)}</span>`;
      pkt.classList.remove("pop"); void pkt.offsetWidth; pkt.classList.add("pop");
      out.innerHTML = `<div class="verdict wait"><span class="big">🔎</span><div><b>Évaluation en cours…</b>Le pare-feu compare le paquet à chaque règle, dans l'ordre.</div></div>`;
      scan.style.opacity = 1;
      for (let k = 0; k < RB.length; k++) {
        const row = rows[k];
        scan.style.height = row.offsetHeight + "px"; scan.style.transform = `translateY(${row.offsetTop}px)`;
        await sleep(RM ? 60 : 520); if (my !== token) return;
        const why = test(RB[k], f);
        if (why) { row.querySelector(".why").textContent = "✕ " + why; row.classList.add("miss"); await sleep(RM ? 40 : 480); if (my !== token) return; continue; }
        const r = RB[k], allow = r.act === "allow";
        row.querySelector(".why").textContent = allow ? "✓ correspond → allow" : "✓ correspond → deny";
        row.classList.add("hit", allow ? "allow-r" : "deny-r");
        rows.slice(k + 1).forEach(x => x.classList.add("skip"));
        scan.style.opacity = 0;
        out.innerHTML = `<div class="verdict ${allow ? "allow" : "deny"}"><span class="big">${allow ? "✅" : "⛔"}</span><div><b>${allow ? "Autorisé" : "Bloqué"} par la règle ${k + 1} · ${esc(r.n)}</b>${esc(f.end)}${k + 1 < RB.length ? ` <span class="mut">${RB.length - k - 1 > 1 ? `Les ${RB.length - k - 1} règles suivantes ne sont` : "La règle suivante n'est"} pas lue${RB.length - k - 1 > 1 ? "s" : ""}.</span>` : ""}</div></div>`;
        return;
      }
    }
    flowsEl.addEventListener("click", e => { const b = e.target.closest(".chipb"); if (b) run(+b.dataset.i); });
    hooks.regles = { start() { if (!started) { started = true; run(0); } }, stop() { token++; scan.style.opacity = 0; } };
  })();

  /* ================= constructeur de règle pas à pas ================= */
  (function builder() {
    const R = "set rulebase security rules LAN-to-Internet-Web";
    const WTABS = ["General", "Source", "Destination", "Application", "Service/URL Category", "Actions"];
    const STEPS = [
      { t: "Créer l'objet adresse", path: "Objects › Addresses › Add",
        f: [["Name", "NET-LAN-Users"], ["Type", "IP Netmask"], ["Adresse", "192.168.1.0/24"], ["Tags", "LAN"]],
        obj: "🏷️ NET-LAN-Users",
        why: "Un <b>objet nommé</b> plutôt qu'une IP en dur : lisible dans les logs, réutilisable, modifiable à un seul endroit. Le <b>tag</b> colore la règle dans la rulebase.",
        cli: ["set tag LAN color color3", "set address NET-LAN-Users ip-netmask 192.168.1.0/24", "set address NET-LAN-Users tag LAN"] },
      { t: "Préparer le groupe de profils", path: "Objects › Security Profile Groups › Add",
        f: [["Name", "PG-Default"], ["Antivirus", "default"], ["Anti-Spyware", "strict"], ["Vulnerability", "strict"], ["URL Filtering", "default"], ["File Blocking", "basic file blocking"], ["WildFire Analysis", "default"]],
        obj: "🧬 PG-Default",
        why: "Le groupe réunit les <b>6 moteurs d'inspection</b>. Sans lui, une règle allow laisse passer virus et exploits. Astuce : nommé <code>default</code>, il s'applique tout seul aux nouvelles règles.",
        cli: ['set profile-group PG-Default virus default spyware strict vulnerability strict url-filtering default file-blocking "basic file blocking" wildfire-analysis default'] },
      { t: "Onglet General : nommer et documenter", path: "Policies › Security › Add", tab: 0,
        f: [["Name", "LAN-to-Internet-Web"], ["Rule Type", "universal (default)"], ["Description", "Navigation web des postes LAN – ticket #1234"], ["Tags", "LAN"]],
        rule: { name: "LAN-to-Internet-Web", tag: "LAN" },
        why: "Un nom explicite <b>source-to-destination-usage</b> et une description avec le <b>n° de ticket</b> : dans deux ans, on sait encore pourquoi la règle existe.",
        cli: [`${R} rule-type universal tag LAN description "Navigation web des postes LAN - ticket #1234"`] },
      { t: "Onglet Source : d'où vient le trafic", path: "Policies › Security › LAN-to-Internet-Web", tab: 1,
        f: [["Source Zone", "LAN"], ["Source Address", "NET-LAN-Users"], ["Source User", "any"]],
        rule: { zones: "LAN → …", src: "NET-LAN-Users", user: "any" },
        why: "La <b>zone source</b> est celle de l'interface d'entrée du paquet. On se limite à l'objet réseau du LAN plutôt que <code>any</code>.",
        cli: [`${R} from LAN source NET-LAN-Users source-user any`] },
      { t: "Onglet Destination : où il va", path: "Policies › Security › LAN-to-Internet-Web", tab: 2,
        f: [["Destination Zone", "WAN"], ["Destination Address", "any"]],
        rule: { zones: "LAN → WAN", dst: "any" },
        why: "La <b>zone de destination</b> est déduite de la table de routage : Internet → route par défaut → ethernet1/1 → zone <b>WAN</b>.",
        cli: [`${R} to WAN destination any`] },
      { t: "Onglet Application : quoi exactement", path: "Policies › Security › LAN-to-Internet-Web", tab: 3,
        f: [["Applications", "web-browsing, ssl"]],
        rule: { app: "web-browsing, ssl" },
        why: "<b>App-ID</b> identifie l'application quel que soit le port. On autorise web-browsing et ssl, <b>rien d'autre</b> : BitTorrent ou un tunnel sur le port 443 seront refusés.",
        cli: [`${R} application [ web-browsing ssl ]`] },
      { t: "Onglet Service : les bons ports", path: "Policies › Security › LAN-to-Internet-Web", tab: 4,
        f: [["Service", "application-default"], ["URL Category", "any"]],
        rule: { svc: "application-default" },
        why: "<b>application-default</b> : ssl seulement sur 443, web-browsing seulement sur 80. Une application sur un port inhabituel est refusée. Le filtrage URL se fait via le profil, pas ici.",
        cli: [`${R} service application-default category any`] },
      { t: "Onglet Actions : autoriser ET inspecter", path: "Policies › Security › LAN-to-Internet-Web", tab: 5,
        f: [["Action", "Allow"], ["Profile Type", "Group"], ["Group Profile", "PG-Default"], ["Log at Session Start", "☐"], ["Log at Session End", "☑"], ["Log Forwarding", "LF-Default"]],
        rule: { act: "allow", prof: "PG-Default", log: "fin de session → LF-Default" },
        why: "<b>Allow + groupe de profils</b> = trafic autorisé <b>et</b> inspecté. Log en fin de session uniquement : le log au début charge le CPU pour rien.",
        cli: [`${R} action allow profile-setting group PG-Default`, `${R} log-start no log-end yes log-setting LF-Default`] },
      { t: "Placer la règle et commiter", path: "Policies › Security › glisser-déposer", mini: true,
        rule: { pos: "au-dessus de Deny-All-Log" },
        why: "Première règle qui correspond = appliquée : la règle doit être <b>au-dessus du refus final</b>. Puis <b>Preview Changes → Validate → Commit</b>.",
        cli: ["move rulebase security rules LAN-to-Internet-Web before Deny-All-Log", "commit"] },
    ];
    const FIELDS = [["name", "Nom"], ["tag", "Tag"], ["zones", "Zones"], ["src", "Source"], ["user", "Utilisateur"], ["dst", "Destination"], ["app", "Application"], ["svc", "Service"], ["act", "Action"], ["prof", "Profils"], ["log", "Log"], ["pos", "Position"]];
    const dotsEl = $("#bd-steps"), form = $("#bd-form"), wtabs = $("#bd-tabs"), cli = $("#bd-cli"), card = $("#rcard");
    dotsEl.innerHTML = STEPS.map((s, i) => `<button type="button" class="bd-dot" role="tab" aria-label="Étape ${i + 1} : ${esc(s.t)}" data-i="${i}"></button>`).join("");
    $("#rc-fields").innerHTML = FIELDS.map(([k, l]) => `<div class="rf" data-k="${k}"><span>${l}</span><b>—</b></div>`).join("");
    let idx = 0, token = 0, auto = false, autoT = 0, started = false;

    async function type(el, text, speed) {
      if (RM) { el.textContent = text; return; }
      el.classList.add("typing");
      for (let i = 1; i <= text.length; i++) { el.textContent = text.slice(0, i); await sleep(speed); }
      el.classList.remove("typing");
    }
    async function go(i) {
      idx = Math.max(0, Math.min(STEPS.length - 1, i));
      const my = ++token, s = STEPS[idx];
      clearTimeout(autoT);
      $$(".bd-dot", dotsEl).forEach((d, k) => { d.classList.toggle("on", k === idx); d.classList.toggle("done", k < idx); d.setAttribute("aria-selected", k === idx); });
      $("#bd-num").textContent = `ÉTAPE ${idx + 1} / ${STEPS.length}`;
      $("#bd-title").textContent = s.t;
      $("#bd-path").textContent = s.path;
      wtabs.innerHTML = s.tab == null ? "" : WTABS.map((w, k) => `<span class="${k === s.tab ? "on" : k < s.tab ? "done" : ""}">${w}</span>`).join("");
      // fiche règle : cumul des étapes 0..idx, les champs de l'étape courante clignotent
      const st = {}; STEPS.slice(0, idx + 1).forEach(x => Object.assign(st, x.rule || {}));
      $$(".rf", card).forEach(rf => {
        const k = rf.dataset.k, v = st[k];
        rf.classList.toggle("set", v != null); rf.querySelector("b").textContent = v != null ? v : "—";
        rf.classList.remove("flash");
        if (s.rule && k in s.rule) { void rf.offsetWidth; rf.classList.add("flash"); }
      });
      $("#rc-objs").innerHTML = STEPS.slice(0, idx + 1).filter(x => x.obj).map(x => `<span class="badge or">${x.obj}</span>`).join("");
      const done = idx === STEPS.length - 1;
      card.classList.toggle("ready", done);
      $("#rc-state").textContent = done ? "✓ prête" : "en construction";
      $("#rc-state").className = "badge" + (done ? " ok" : "");
      const why = $("#bd-why"); why.classList.remove("pop"); why.innerHTML = s.why;
      why.style.animation = "none"; void why.offsetWidth; why.style.animation = "";
      $("#bd-prev").disabled = idx === 0;
      $("#bd-next").textContent = done ? "Terminé ✓" : "Suivant ▶";
      $("#bd-next").disabled = done;
      // écran PAN-OS
      if (s.mini) {
        form.innerHTML = `<div class="mini-rb"><div><span>1 · Block-Bad-IP</span><em class="deny">deny</em></div><div><span>2 · Block-QUIC</span><em class="deny">deny</em></div>
          <div class="new"><span>3 · LAN-to-Internet-Web</span><em class="allow">allow</em></div><div><span>4 · Deny-All-Log</span><em class="deny">deny</em></div></div>
          <span class="commit-ok">✓ Commit réussi · 0 erreur · 0 avertissement</span>`;
        if (!RM) setTimeout(() => { if (my === token && PAN.confetti) PAN.confetti(card); }, 1500);
      } else {
        form.innerHTML = s.f.map(([l], k) => `<div class="ff" style="animation-delay:${k * 60}ms"><label>${esc(l)}</label><div class="in"></div></div>`).join("");
      }
      cli.textContent = "";
      const ins = $$(".in", form);
      for (let k = 0; !s.mini && k < s.f.length; k++) { await type(ins[k], s.f[k][1], 22); if (my !== token) return; }
      cli.classList.add("typing");
      for (const line of s.cli) {
        const start = cli.textContent;
        if (RM) cli.textContent = start + line + "\n";
        else for (let c = 1; c <= line.length; c += 3) { cli.textContent = start + line.slice(0, c); await sleep(10); if (my !== token) return; }
        cli.textContent = start + line + "\n";
      }
      cli.classList.remove("typing");
      if (auto && !done) autoT = setTimeout(() => { if (my === token) go(idx + 1); }, 2600);
      if (done) setAuto(false);
    }
    function setAuto(on) { auto = on; $("#bd-play").textContent = on ? "⏸ Pause" : "⏵ Lecture auto"; if (!on) clearTimeout(autoT); }
    dotsEl.addEventListener("click", e => { const d = e.target.closest(".bd-dot"); if (d) { setAuto(false); go(+d.dataset.i); } });
    $("#bd-next").addEventListener("click", () => { setAuto(false); go(idx + 1); });
    $("#bd-prev").addEventListener("click", () => { setAuto(false); go(idx - 1); });
    $("#bd-reset").addEventListener("click", () => { setAuto(false); go(0); });
    $("#bd-play").addEventListener("click", () => {
      if (auto) return setAuto(false);
      setAuto(true);
      if (idx === STEPS.length - 1) go(0); else go(idx + 1);
    });
    hooks.exemple = { start() { if (!started) { started = true; go(0); } }, stop() { setAuto(false); } };
  })();

  /* ================= NAT animé ================= */
  (function nat() {
    const X = { L: 110, IN: 300, C: 380, OUT: 460, R: 650 };
    const MODES = {
      dnat: {
        l: { zone: "ZONE WAN", ico: "🌍", name: "Client Internet", ip: "198.51.100.7", fill: "rgba(239,68,68,.06)", stroke: "rgba(239,68,68,.5)" },
        r: { zone: "ZONE DMZ", ico: "🖥️", name: "SRV-WEB-01", ip: "10.10.20.10", fill: "rgba(56,189,248,.06)", stroke: "rgba(56,189,248,.5)" },
        ifl: "ethernet1/1", ifr: "ethernet1/3", src: "198.51.100.7", dst: "203.0.113.20:443",
        st: [
          { x: "L", t: "Le client Internet vise l'<b>IP publique</b> 203.0.113.20 sur le port 443." },
          { x: "IN", hl: "ifl", t: "Le paquet entre par ethernet1/1 → <b>zone source = WAN</b>." },
          { x: "C", hl: "nat", t: "<b>Politique NAT</b> : la route vers 203.0.113.20 sort par ethernet1/1, donc la règle NAT est <b>From WAN To WAN</b>." },
          { x: "C", hl: "sec", t: "<b>Politique de sécurité</b> : destination = <b>203.0.113.20 (pré-NAT)</b>, mais zone de destination = <b>DMZ</b> (post-NAT, route vers 10.10.20.10)." },
          { x: "OUT", hl: "ifr", dst: "10.10.20.10:443", t: "<b>Traduction</b> à la sortie : destination 203.0.113.20 → <b>10.10.20.10</b>." },
          { x: "R", t: "Le serveur reçoit le paquet. La réponse est re-traduite automatiquement grâce à la session." },
        ] },
      snat: {
        l: { zone: "ZONE LAN", ico: "💻", name: "Poste LAN", ip: "192.168.1.50", fill: "rgba(34,197,94,.06)", stroke: "rgba(34,197,94,.5)" },
        r: { zone: "ZONE WAN", ico: "🌍", name: "Internet", ip: "8.8.8.8", fill: "rgba(239,68,68,.06)", stroke: "rgba(239,68,68,.5)" },
        ifl: "ethernet1/2", ifr: "ethernet1/1", src: "192.168.1.50:51515", dst: "8.8.8.8:443",
        st: [
          { x: "L", t: "Le poste envoie vers Internet avec son <b>IP privée</b>, non routable sur Internet." },
          { x: "IN", hl: "ifl", t: "Le paquet entre par ethernet1/2 → <b>zone source = LAN</b>." },
          { x: "C", hl: "nat", t: "<b>Politique NAT</b> : <b>From LAN To WAN</b>, Dynamic IP and Port sur l'interface ethernet1/1." },
          { x: "C", hl: "sec", t: "<b>Politique de sécurité</b> : source = <b>192.168.1.50 (IP privée, pré-NAT)</b>, From LAN To WAN." },
          { x: "OUT", hl: "ifr", src: "203.0.113.10:20431", t: "<b>Traduction</b> : source → <b>203.0.113.10</b> (IP de ethernet1/1), port source changé (DIPP)." },
          { x: "R", t: "Internet répond à 203.0.113.10 ; le pare-feu retrouve la session et renvoie la réponse au poste." },
        ] },
    };
    const g = $("#nat-pkt"), list = $("#nat-steps");
    let mode = "dnat", i = -1, playing = !RM, timer = 0, started = false;
    const T = id => document.getElementById(id);
    function setup() {
      const m = MODES[mode];
      T("zl").setAttribute("fill", m.l.fill); T("zl").setAttribute("stroke", m.l.stroke);
      T("zr").setAttribute("fill", m.r.fill); T("zr").setAttribute("stroke", m.r.stroke);
      T("zl-t").textContent = m.l.zone; T("zr-t").textContent = m.r.zone;
      T("nl-i").textContent = m.l.ico; T("nl-n").textContent = m.l.name; T("nl-ip").textContent = m.l.ip;
      T("nr-i").textContent = m.r.ico; T("nr-n").textContent = m.r.name; T("nr-ip").textContent = m.r.ip;
      T("if-l").textContent = m.ifl; T("if-r").textContent = m.ifr;
      list.innerHTML = m.st.map((s, k) => `<li data-k="${k}"><span>${s.t}</span></li>`).join("");
      i = -1; stage(0);
    }
    function stage(k) {
      const m = MODES[mode]; i = k; const s = m.st[k];
      g.style.transform = `translate(${X[s.x]}px,200px)`;
      ["nat-box", "sec-box", "fw-box"].forEach(id => T(id).classList.remove("hl"));
      ["if-l", "if-r"].forEach(id => T(id).classList.remove("ifhl"));
      if (s.hl === "nat") T("nat-box").classList.add("hl");
      if (s.hl === "sec") T("sec-box").classList.add("hl");
      if (s.hl === "ifl") T("if-l").classList.add("ifhl");
      if (s.hl === "ifr") T("if-r").classList.add("ifhl");
      if (s.x === "C") T("fw-box").classList.add("hl");
      // en-tête : valeurs traduites à partir de l'étape qui les change
      const tr = m.st.slice(0, k + 1).reduce((a, x) => Object.assign(a, x.src ? { src: x.src } : {}, x.dst ? { dst: x.dst } : {}), {});
      [["h-src", "src"], ["h-dst", "dst"]].forEach(([id, key]) => {
        const el = T(id); el.textContent = tr[key] || m[key];
        el.classList.toggle("chg", !!tr[key]); el.classList.remove("flash");
        if (s[key]) { void el.getBBox(); el.classList.add("flash"); }
      });
      $$("li", list).forEach((li, n) => { li.classList.toggle("on", n === k); li.classList.toggle("done", n < k); });
      clearTimeout(timer);
      if (playing) timer = setTimeout(() => stage(k + 1 < m.st.length ? k + 1 : 0), k + 1 < m.st.length ? 2600 : 4200);
    }
    function setPlay(on) { playing = on; $("#nat-play").textContent = on ? "⏸ Pause" : "▶ Lecture"; clearTimeout(timer); if (on) stage(i < 0 ? 0 : i); }
    $("#nat-mode").addEventListener("click", e => {
      const b = e.target.closest("button"); if (!b || b.dataset.m === mode) return;
      mode = b.dataset.m; $$("button", $("#nat-mode")).forEach(x => x.classList.toggle("on", x === b)); setup();
    });
    $("#nat-play").addEventListener("click", () => setPlay(!playing));
    $("#nat-step").addEventListener("click", () => { setPlay(false); stage((i + 1) % MODES[mode].st.length); });
    $("#nat-reset").addEventListener("click", () => stage(0));
    list.addEventListener("click", e => { const li = e.target.closest("li"); if (li) { setPlay(false); stage(+li.dataset.k); } });
    $("#nat-play").textContent = playing ? "⏸ Pause" : "▶ Lecture";
    hooks.nat = { start() { if (!started) { started = true; setup(); } else if (playing) stage(i); }, stop() { clearTimeout(timer); } };
  })();

  /* ================= chaîne d'inspection des profils ================= */
  (function pipe() {
    const PT = [
      { ico: "✅", w: 6 },
      { ico: "🦠", gate: "av", log: "virus · Win32.Trojan.Agent · reset-both", name: "Virus" },
      { ico: "📡", gate: "as", log: "spyware · requête DNS C2 → sinkhole", name: "Callback C2" },
      { ico: "💥", gate: "vp", log: "vulnerability · Apache Log4j RCE · reset-both", name: "Exploit Log4j" },
      { ico: "🎣", gate: "url", log: "url · catégorie phishing · block-url", name: "Lien de phishing" },
      { ico: "📦", gate: "fb", log: "file · setup.exe téléchargé · block", name: "Exécutable" },
      { ico: "❓", gate: "wf", log: "wildfire · fichier inconnu → verdict malicious", name: "Malware zero-day" },
    ];
    const BAG = PT.flatMap(p => Array(p.w || 1).fill(p));
    const track = $("#pipe-track"), log = $("#pipe-log"), gates = {};
    $$(".gate", track).forEach(gt => gates[gt.dataset.g] = gt);
    const cnt = { ok: 0, bl: 0, lk: 0 };
    let on = true, running = false, timer = 0, started = false;
    const now = () => new Date().toLocaleTimeString("fr-FR");
    function bump(k) { cnt[k]++; const el = $("#ps-" + k); el.querySelector("b").textContent = cnt[k]; el.classList.remove("bump"); void el.offsetWidth; el.classList.add("bump"); }
    function addLog(cls, txt) {
      const d = document.createElement("div"); d.innerHTML = `<span class="t">${now()}</span> <span class="${cls}">${txt}</span>`;
      log.prepend(d); while (log.children.length > 8) log.lastChild.remove();
    }
    const cx = el => { const a = el.getBoundingClientRect(), b = track.getBoundingClientRect(); return a.left - b.left + a.width / 2; };
    function spawn() {
      const p = BAG[Math.floor(Math.random() * BAG.length)];
      const el = document.createElement("span"); el.className = "pk " + (p.gate ? "bad" : "ok"); el.textContent = p.ico;
      track.append(el);
      const x0 = cx($(".pipe-src", track)), xEnd = cx($(".pipe-dst", track));
      const stopAt = p.gate && on ? cx(gates[p.gate]) : xEnd;
      const dy = (Math.random() * 16 - 8).toFixed(0);
      const dur = (stopAt - x0) * 4;
      el.animate([{ transform: `translate(${x0}px,${dy}px)` }, { transform: `translate(${stopAt}px,${dy}px)` }], { duration: dur, easing: "linear", fill: "forwards" })
        .finished.then(() => {
          if (p.gate && on) {
            const gt = gates[p.gate]; gt.classList.remove("hit"); void gt.offsetWidth; gt.classList.add("hit");
            const b = document.createElement("span"); b.className = "boom"; b.style.left = stopAt + "px"; track.append(b);
            setTimeout(() => b.remove(), 650);
            bump("bl"); addLog("b", `⛔ threat · ${p.log}`);
          } else if (p.gate) { bump("lk"); addLog("l", `⚠ ${p.name} livré à l'utilisateur · aucune inspection`); }
          else bump("ok");
          el.animate([{ opacity: 1 }, { opacity: 0, transform: `translate(${stopAt}px,${dy}px) scale(${p.gate && on ? 1.8 : .6})` }], { duration: 300, fill: "forwards" }).finished.then(() => el.remove());
        }).catch(() => el.remove());
    }
    function loop() { if (!running) return; spawn(); timer = setTimeout(loop, 650 + Math.random() * 500); }
    function setRun(r) { running = r; $("#pipe-play").textContent = r ? "⏸ Pause" : "▶ Lecture"; clearTimeout(timer); if (r) loop(); }
    $("#pipe-mode").addEventListener("click", e => {
      const b = e.target.closest("button"); if (!b) return;
      on = b.dataset.m === "on"; $$("button", $("#pipe-mode")).forEach(x => x.classList.toggle("on", x === b));
      Object.values(gates).forEach(gt => gt.classList.toggle("off", !on));
      addLog(on ? "o" : "l", on ? "🛡️ groupe de profils PG-Default appliqué à la règle" : "⚠ profile-setting : none · le trafic autorisé n'est plus inspecté");
    });
    let userPaused = false;
    $("#pipe-play").addEventListener("click", () => { userPaused = running; setRun(!running); });
    $("#pipe-reset").addEventListener("click", () => { ["ok", "bl", "lk"].forEach(k => { cnt[k] = 0; $("#ps-" + k + " b").textContent = 0; }); log.innerHTML = ""; });
    hooks.profils = {
      start() {
        if (!started) { started = true; addLog("o", "🛡️ groupe de profils PG-Default appliqué à la règle"); }
        if (!userPaused && !RM) setRun(true);
      },
      stop() { setRun(false); $$(".pk", track).forEach(x => x.remove()); },
    };
  })();

  /* ================= objets : cartes à retourner (tactile) ================= */
  $$(".flip").forEach(f => f.addEventListener("click", () => f.classList.toggle("on")));

  /* ================= logs : graduations de la frise 24 h ================= */
  $$(".tl-bar[data-every]").forEach(bar => {
    const every = +bar.dataset.every, n = 1440 / every;
    bar.innerHTML = Array.from({ length: n }, (_, k) => { const p = (k + .5) / n; return `<i style="left:${(p * 100).toFixed(2)}%;animation-delay:${(p * 12).toFixed(2)}s"></i>`; }).join("");
  });

  /* ================= checklist ================= */
  (function checklist() {
    const ITEMS = [
      "Aucune règle any → any / allow",
      "Toutes les règles allow utilisent App-ID + service application-default",
      "Un groupe de profils de sécurité sur chaque règle allow",
      "Log at Session End + Log Forwarding sur toutes les règles",
      "Règle Deny-All-Log en fin de rulebase (ou interzone-default loggée)",
      "QUIC bloqué si le décryptage SSL est en place",
      "Règles NAT : IP pré-NAT et bonnes zones vérifiées",
      "Objets nommés selon la convention, aucun objet inutilisé",
      "Management : permitted-ip, HTTP/Telnet désactivés, comptes nominatifs",
      "Mises à jour dynamiques planifiées (App/Threat, AV, WildFire)",
      "NTP, DNS et syslog/Panorama configurés",
      "Config exportée et sauvegardée avant mise en prod",
      "Flux testés avec test security-policy-match",
    ];
    const KEY = "panw-bp-checklist";
    let st = {};
    try { st = JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { st = {}; }
    const list = $("#chk"), ring = $(".ring"), fg = $("#ring-fg");
    list.innerHTML = ITEMS.map((t, i) => `<li><label><input type="checkbox" data-i="${i}"${st[i] ? " checked" : ""}><span>${esc(t)}</span></label></li>`).join("");
    function upd() {
      const n = ITEMS.filter((_, i) => st[i]).length, p = n / ITEMS.length;
      fg.style.strokeDashoffset = (301.6 * (1 - p)).toFixed(1);
      $("#ring-pct").textContent = Math.round(p * 100) + "%";
      ring.classList.toggle("full", n === ITEMS.length);
      $("#chk-count").textContent = `${n} / ${ITEMS.length} points validés` + (n === ITEMS.length ? " · prêt pour la prod 🚀" : "");
      try { localStorage.setItem(KEY, JSON.stringify(st)); } catch (e) { /* stockage indisponible */ }
    }
    list.addEventListener("change", e => {
      st[e.target.dataset.i] = e.target.checked; upd();
      if (e.target.checked && ITEMS.every((_, i) => st[i]) && PAN.confetti) PAN.confetti(ring);
    });
    $("#chk-reset").addEventListener("click", () => { st = {}; $$("input", list).forEach(c => c.checked = false); upd(); });
    upd();
  })();

  show(location.hash.slice(1));
})();
