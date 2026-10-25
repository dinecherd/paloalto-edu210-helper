/* Office Academy — PowerPoint : 8 modules du débutant à l'expert */
(function () {
  "use strict";

  /* ---------- Petits utilitaires pour des check robustes (jamais d'exception) ---------- */
  // normalise : minuscules, sans accents, espaces réduits
  const N = s => String(s == null ? "" : s).toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, " ").trim();
  const S = a => (a && Array.isArray(a.slides)) ? a.slides : [];
  const E = s => (s && Array.isArray(s.els)) ? s.els : [];
  const byType = (s, t) => E(s).filter(e => e && e.type === t);
  const titleOf = s => { const t = byType(s, "title")[0]; return t ? N(t.text) : ""; };
  const slideText = s => N(E(s).map(e => (e && e.text) || "").join(" "));
  const lines = t => String(t == null ? "" : t).split(/\n/).map(x => x.trim()).filter(Boolean);
  const on = v => !!v && !["none", "aucune", "aucun", "non", "off", "0", "false"].includes(N(v));
  const lay = s => N(s && s.layout);
  // index de la 1re diapo dont le titre contient txt (-1 sinon)
  const idxTitle = (a, txt) => S(a).findIndex(s => titleOf(s).includes(N(txt)));
  // tous les éléments (de toutes les diapos) dont le texte contient txt
  const elsWith = (a, txt) => {
    const out = [];
    S(a).forEach((s, i) => E(s).forEach(e => { if (e && N(e.text).includes(N(txt))) out.push({ slide: i, el: e }); }));
    return out;
  };
  const num = v => { const n = Number(v); return isFinite(n) ? n : 0; };

  OFFICE.add("powerpoint", [

  /* ======================================================================
     p01 — Niveau 1 : Premiers pas
     ====================================================================== */
  {
    id: "p01",
    level: 1,
    title: "Premiers pas : interface, diapositives et dispositions",
    icon: "🧭",
    duration: 50,
    summary: "Prendre ses repères dans PowerPoint 365 et construire un premier deck propre, diapositive par diapositive.",
    objectives: [
      "Être capable de se repérer dans le ruban, le volet des miniatures et la zone de notes",
      "Être capable d'ajouter, dupliquer, déplacer et supprimer des diapositives",
      "Être capable de choisir la bonne disposition plutôt que d'empiler des zones de texte",
      "Être capable de passer d'un mode d'affichage à l'autre et de lancer le diaporama"
    ],
    lessons: [
      {
        title: "L'interface de PowerPoint 365",
        html: `<p>À l'ouverture, PowerPoint propose l'écran <b>Accueil</b> : <i>Nouvelle présentation</i>, modèles, fichiers récents. Une fois la présentation ouverte, l'écran se découpe en quatre zones :</p>
<ul>
<li><b>Le ruban</b> : onglets <i>Fichier, Accueil, Insertion, Dessin, Création, Transitions, Animations, Diaporama, Enregistrer, Révision, Affichage</i>. Chaque onglet regroupe des commandes par groupes (ex. <i>Accueil › Police</i>).</li>
<li><b>Le volet des miniatures</b> (à gauche) : la liste des diapositives, où l'on ajoute, réorganise et supprime.</li>
<li><b>La zone de diapositive</b> (au centre) : l'espace de travail.</li>
<li><b>Le volet Notes</b> (sous la diapo) : ce que vous direz à l'oral, invisible pour le public. Affichez-le avec le bouton <i>Notes</i> de la barre d'état.</li>
</ul>
<p>La <b>barre d'état</b>, tout en bas, affiche le numéro de diapo, les boutons des modes d'affichage et le zoom.</p>
<div class="callout">La <b>zone Rechercher</b> (en haut, loupe) permet de taper le nom d'une commande (« transition », « masque ») : pratique quand on ne sait pas dans quel onglet elle se trouve.</div>
<div class="o-tip">💡 Double-cliquez sur un onglet pour réduire le ruban et gagner de la place ; <kbd>Ctrl</kbd>+<kbd>F1</kbd> fait la même chose.</div>`
      },
      {
        title: "Ajouter et organiser les diapositives",
        html: `<p>Une présentation est une suite de diapositives. Les gestes de base :</p>
<table class="o-table"><thead><tr><th>Action</th><th>Comment</th></tr></thead><tbody>
<tr><td>Nouvelle diapositive</td><td><kbd>Ctrl</kbd>+<kbd>M</kbd> ou <i>Accueil › Nouvelle diapositive</i></td></tr>
<tr><td>Dupliquer</td><td>Clic droit sur la miniature › <i>Dupliquer la diapositive</i></td></tr>
<tr><td>Déplacer</td><td>Glisser la miniature, ou <kbd>Ctrl</kbd>+<kbd>Flèche haut</kbd> / <kbd>Flèche bas</kbd></td></tr>
<tr><td>Supprimer</td><td>Sélectionner la miniature puis <kbd>Suppr</kbd></td></tr>
<tr><td>Masquer</td><td>Clic droit › <i>Masquer la diapositive</i> (elle reste dans le fichier mais n'est pas projetée)</td></tr>
</tbody></table>
<div class="o-path">Accueil › Diapositives › Nouvelle diapositive</div>
<p>La petite flèche sous <i>Nouvelle diapositive</i> permet de choisir directement la disposition de la diapo créée.</p>
<div class="callout warn">Piège : <kbd>Suppr</kbd> supprime ce qui est <b>sélectionné</b>. Si le curseur est dans une zone de texte, c'est du texte qui disparaît, pas la diapo. Cliquez d'abord sur la miniature.</div>`
      },
      {
        title: "Les dispositions : la structure avant la décoration",
        html: `<p>Une <b>disposition</b> est un gabarit de diapo qui contient des <b>espaces réservés</b> (« Cliquez pour ajouter un titre »). Les principales :</p>
<ul>
<li><b>Diapositive de titre</b> : titre + sous-titre, pour l'ouverture.</li>
<li><b>Titre et contenu</b> : la disposition la plus courante (titre + liste à puces, tableau, graphique…).</li>
<li><b>Titre de section</b> : pour annoncer un changement de partie.</li>
<li><b>Deux contenus</b> / <b>Comparaison</b> : avant/après, pour/contre.</li>
<li><b>Titre seul</b> et <b>Vide</b> : pour les visuels pleine page.</li>
</ul>
<div class="o-path">Accueil › Diapositives › Disposition</div>
<div class="callout ok">Bonne pratique : utilisez toujours les espaces réservés de titre. Le titre sert au mode Plan, aux zooms, à la navigation et au <b>vérificateur d'accessibilité</b>. Une zone de texte ajoutée « à la main » n'est pas un titre.</div>
<div class="callout bad">Erreur fréquente : partir de la disposition <i>Vide</i> et tout dessiner avec des zones de texte. Résultat : polices incohérentes, rien ne suit quand on change de thème.</div>
<p><b>Rétablir</b> (<i>Accueil › Diapositives › Rétablir</i>) remet les espaces réservés à leur position et leur mise en forme d'origine : idéal pour réparer une diapo « bricolée ».</p>`
      },
      {
        title: "Saisir du texte et gérer les niveaux de puces",
        html: `<p>Cliquez dans un espace réservé et tapez. Dans une liste à puces :</p>
<ul>
<li><kbd>Entrée</kbd> crée une nouvelle puce ; <kbd>Maj</kbd>+<kbd>Entrée</kbd> va à la ligne <b>sans</b> nouvelle puce.</li>
<li><kbd>Tab</kbd> descend d'un niveau (sous-puce) ; <kbd>Maj</kbd>+<kbd>Tab</kbd> remonte.</li>
<li><kbd>Ctrl</kbd>+<kbd>Entrée</kbd> passe à l'espace réservé suivant (et crée une diapo si vous êtes sur le dernier).</li>
</ul>
<p>Le <b>mode Plan</b> (<i>Affichage › Mode Plan</i>) montre tout le texte des titres et des puces : idéal pour écrire le fond d'abord, la forme ensuite.</p>
<div class="callout warn">Piège : quand le texte déborde, PowerPoint réduit automatiquement la police (« Ajustement automatique »). C'est un signal d'alarme, pas une solution : coupez le texte ou scindez la diapo (bouton d'options d'ajustement › <i>Fractionner le texte entre deux diapositives</i>).</div>
<div class="o-tip">💡 Visez au maximum 5-6 puces de 6-8 mots. Si vous écrivez des phrases complètes, elles vont dans les <b>notes</b>.</div>`
      },
      {
        title: "Les modes d'affichage et le diaporama",
        html: `<table class="o-table"><thead><tr><th>Mode</th><th>Usage</th></tr></thead><tbody>
<tr><td><b>Normal</b></td><td>Édition diapo par diapo (mode par défaut)</td></tr>
<tr><td><b>Mode Plan</b></td><td>Écrire/relire le texte de toute la présentation</td></tr>
<tr><td><b>Trieuse de diapositives</b></td><td>Vue d'ensemble, réorganiser, repérer les sections</td></tr>
<tr><td><b>Page de notes</b></td><td>Relire et imprimer les notes du présentateur</td></tr>
<tr><td><b>Lecture</b></td><td>Diaporama dans la fenêtre, pour se relire</td></tr>
<tr><td><b>Diaporama</b></td><td>Plein écran, devant le public</td></tr>
</tbody></table>
<div class="o-path">Affichage › Affichages des présentations › Trieuse de diapositives</div>
<p>Lancer le diaporama : <kbd>F5</kbd> depuis le début, <kbd>Maj</kbd>+<kbd>F5</kbd> depuis la diapo en cours. <kbd>Échap</kbd> pour sortir.</p>
<div class="o-tip">💡 Relisez toujours en mode Diaporama avant d'envoyer : les fautes, débordements et animations ratées sautent aux yeux en plein écran.</div>`
      }
    ],
    shortcuts: [
      { keys: "Ctrl+M", desc: "Nouvelle diapositive" },
      { keys: "Ctrl+D", desc: "Dupliquer l'objet ou la diapositive sélectionnée" },
      { keys: "Ctrl+Entrée", desc: "Aller à l'espace réservé suivant" },
      { keys: "Tab", desc: "Abaisser le niveau de puce" },
      { keys: "Maj+Tab", desc: "Remonter le niveau de puce" },
      { keys: "F5", desc: "Diaporama depuis le début" },
      { keys: "Maj+F5", desc: "Diaporama depuis la diapositive active" },
      { keys: "Ctrl+F1", desc: "Réduire / afficher le ruban" }
    ],
    labs: [
      {
        id: "p01-lab1",
        title: "Lab 1 — Le bilan de la boulangerie Martin",
        kind: "slides",
        scenario: `<p>Claire, gérante de la <b>Boulangerie Martin</b> (3 boutiques), présente son bilan annuel à sa banque la semaine prochaine. Son stagiaire a commencé un brouillon… très brouillon. Vous reprenez la main : titre clair, diapo de chiffres clés, bonne disposition pour la comparaison, puis relecture en diaporama.</p>`,
        setup: {
          theme: "blanc",
          slides: [
            { layout: "title", title: "presentation", subtitle: "" },
            { layout: "content", title: "Qui sommes-nous", body: "3 boutiques à Lyon\n18 salariés\nFondée en 1998" },
            { layout: "content", title: "Avant / Après la rénovation", body: "Avant : 220 clients par jour, file d'attente dehors. Après : 340 clients par jour, deux caisses, coin snacking" },
            { layout: "content", title: "Projets 2026", body: "Ouvrir une 4e boutique\nLancer la vente en ligne\nRecruter 2 apprentis" }
          ]
        },
        tasks: [
          {
            text: "Sur la diapo 1, remplace le titre « presentation » par un vrai titre contenant <b>Bilan 2025</b>, et ajoute un sous-titre (ex. « Boulangerie Martin — Rendez-vous banque »).",
            hint: "Double-clique dans l'espace réservé de titre de la première diapo pour le modifier.",
            check: a => { const s = S(a)[0]; if (!s) return false; const sub = byType(s, "subtitle")[0]; return titleOf(s).includes("bilan 2025") && !!sub && N(sub.text).length >= 3; }
          },
          {
            text: "Ajoute une diapositive <b>Titre et contenu</b> intitulée <b>Chiffres clés</b> avec au moins <b>3 puces</b> (ex. CA 1,2 M€ / +8 % / 410 000 baguettes).",
            hint: "Ctrl+M ou bouton Nouvelle diapositive, puis une puce par ligne (Entrée).",
            check: a => S(a).some(s => titleOf(s).includes("chiffres cles") && byType(s, "body").some(b => lines(b.text).length >= 3))
          },
          {
            text: "La diapo « Avant / Après » mélange tout dans une seule puce : passe-la en disposition <b>Deux contenus</b> et répartis « Avant » à gauche, « Après » à droite.",
            hint: "Sélectionne la diapo › Disposition › Deux contenus, puis colle chaque partie dans sa colonne.",
            check: a => { const i = idxTitle(a, "avant"); if (i < 0) return false; const s = S(a)[i], b = byType(s, "body").map(e => N(e.text)); return lay(s) === "two" && b.some(t => t.includes("avant") && !t.includes("apres")) && b.some(t => t.includes("apres") && !t.includes("avant")); }
          },
          {
            text: "Place la diapo « Projets 2026 » <b>juste après</b> « Chiffres clés » et termine par une diapo de conclusion : le deck doit compter au moins <b>5 diapositives</b>.",
            hint: "Glisse les miniatures ou utilise les boutons monter/descendre du volet.",
            check: a => { const c = idxTitle(a, "chiffres cles"), p = idxTitle(a, "projets 2026"); return a.count() >= 5 && c >= 0 && p === c + 1; }
          },
          {
            text: "Relis ton travail en lançant le <b>diaporama</b>.",
            hint: "Bouton Diaporama (ou F5), puis Échap pour revenir.",
            check: a => a.played === true
          }
        ],
        solution: `<ol>
<li>Diapo 1 : titre « Bilan 2025 », sous-titre « Boulangerie Martin — Rendez-vous banque ».</li>
<li><kbd>Ctrl</kbd>+<kbd>M</kbd> (disposition Titre et contenu), titre « Chiffres clés », puces : « CA : 1,2 M€ », « +8 % vs 2024 », « 410 000 baguettes ».</li>
<li>Diapo « Avant / Après » : <i>Disposition › Deux contenus</i>. Gauche : « Avant : 220 clients/jour, file dehors ». Droite : « Après : 340 clients/jour, 2 caisses, snacking ».</li>
<li>Réordonner : Bilan 2025 › Qui sommes-nous › Avant/Après › Chiffres clés › Projets 2026 › « Merci — questions ? ».</li>
<li><kbd>F5</kbd>, parcourir, <kbd>Échap</kbd>.</li>
</ol>`
      },
      {
        id: "p01-lab2",
        title: "Lab 2 — Réflexes clavier du débutant",
        kind: "keys",
        scenario: `<p>Les présentateurs efficaces ne cherchent pas les boutons. Enchaînez ces raccourcis d'édition et de lancement le plus vite possible.</p>`,
        setup: {
          sequence: [
            { keys: "Ctrl+M", desc: "Nouvelle diapositive" },
            { keys: "Ctrl+Entrée", desc: "Espace réservé suivant" },
            { keys: "Tab", desc: "Abaisser la puce d'un niveau" },
            { keys: "Maj+Tab", desc: "Remonter la puce d'un niveau" },
            { keys: "Ctrl+Z", desc: "Annuler" },
            { keys: "Ctrl+Y", desc: "Rétablir" },
            { keys: "F5", desc: "Diaporama depuis le début" },
            { keys: "Maj+F5", desc: "Diaporama depuis la diapo active" }
          ]
        },
        solution: `<ol><li>Ctrl+M</li><li>Ctrl+Entrée</li><li>Tab</li><li>Maj+Tab</li><li>Ctrl+Z</li><li>Ctrl+Y</li><li>F5</li><li>Maj+F5</li></ol>`
      }
    ],
    quiz: [
      { q: "Quel raccourci insère une nouvelle diapositive ?", choices: ["Ctrl+N", "Ctrl+M", "Ctrl+Maj+D", "Alt+N"], answer: 1, explain: "Ctrl+M crée une nouvelle diapositive. Ctrl+N crée une nouvelle présentation." },
      { q: "Pourquoi préférer l'espace réservé de titre à une simple zone de texte ?", choices: ["Il est plus grand", "Il est reconnu comme titre (mode Plan, navigation, accessibilité, thème)", "Il ne peut pas être déplacé", "Il s'imprime en couleur"], answer: 1, explain: "Le titre structure la présentation : il alimente le mode Plan, les zooms, les lecteurs d'écran et suit le thème." },
      { q: "Dans une liste à puces, comment créer une sous-puce ?", choices: ["Maj+Entrée", "Ctrl+Entrée", "Tab en début de puce", "Alt+Flèche bas"], answer: 2, explain: "Tab abaisse le niveau, Maj+Tab le remonte. Maj+Entrée passe à la ligne sans nouvelle puce." },
      { q: "Quel mode d'affichage est le plus adapté pour réorganiser 40 diapositives ?", choices: ["Page de notes", "Lecture", "Mode Plan", "Trieuse de diapositives"], answer: 3, explain: "La Trieuse affiche toutes les miniatures en grille : on glisse-dépose facilement et on voit les sections." },
      { q: "Le texte déborde et PowerPoint réduit la police tout seul. Que faire ?", choices: ["Rien, c'est prévu pour", "Réduire encore la police à 10 pt", "Couper le texte ou le répartir sur deux diapositives", "Passer en disposition Vide"], answer: 2, explain: "L'ajustement automatique est un signal : trop de texte. On synthétise, on déplace le détail dans les notes ou on scinde." }
    ]
  },

  /* ======================================================================
     p02 — Niveau 1 : Thèmes, texte, formes, alignement
     ====================================================================== */
  {
    id: "p02",
    level: 1,
    title: "Mettre en forme : thèmes, texte, formes et alignement",
    icon: "🎨",
    duration: 55,
    summary: "Donner une allure professionnelle et cohérente à un deck en quelques clics, sans « bricolage » diapo par diapo.",
    objectives: [
      "Être capable d'appliquer un thème et une variante à toute la présentation",
      "Être capable de mettre en forme le texte (taille, gras, couleur) avec mesure",
      "Être capable d'insérer formes et icônes et de les aligner proprement",
      "Être capable de réutiliser une mise en forme avec Reproduire la mise en forme"
    ],
    lessons: [
      {
        title: "Thèmes et variantes",
        html: `<p>Un <b>thème</b> est un ensemble cohérent : <b>couleurs</b> (palette de 12), <b>polices</b> (titres + corps), <b>effets</b> et <b>styles d'arrière-plan</b>. Il s'applique à toutes les diapositives d'un coup.</p>
<div class="o-path">Création › Thèmes</div>
<p>Le groupe <b>Variantes</b> propose des déclinaisons du thème ; la flèche « Autres » donne accès séparément à <i>Couleurs</i>, <i>Polices</i>, <i>Effets</i> et <i>Styles d'arrière-plan</i>.</p>
<div class="callout ok">Bonne pratique : choisissez le thème <b>tôt</b>. Toutes les couleurs que vous utiliserez ensuite seront prises dans les <b>couleurs du thème</b> (première ligne du nuancier) et suivront si vous changez de charte.</div>
<div class="callout warn">Piège : une couleur choisie dans « Couleurs standard » ou « Autres couleurs » est figée. Si l'entreprise change de charte, elle ne bougera pas.</div>`
      },
      {
        title: "Mettre en forme le texte",
        html: `<p>Le groupe <i>Accueil › Police</i> regroupe police, taille, <b>gras</b> (<kbd>Ctrl</kbd>+<kbd>G</kbd>), <i>italique</i> (<kbd>Ctrl</kbd>+<kbd>I</kbd>), souligné (<kbd>Ctrl</kbd>+<kbd>U</kbd>) et couleur. Le groupe <i>Paragraphe</i> gère puces, alignement (centrer : <kbd>Ctrl</kbd>+<kbd>E</kbd>), interligne et colonnes.</p>
<ul>
<li>Tailles repères en projection : titres 32-44 pt, corps 20-28 pt, jamais moins de 18 pt.</li>
<li>Deux polices maximum (celles du thème, de préférence).</li>
<li>Le <b>gras</b> sert à faire ressortir 1 ou 2 mots clés, pas des paragraphes entiers.</li>
</ul>
<div class="o-path">Accueil › Police › Couleur de police</div>
<div class="o-tip">💡 <kbd>Ctrl</kbd>+<kbd>Espace</kbd> efface la mise en forme manuelle des caractères sélectionnés et revient au style du thème.</div>
<div class="callout bad">Erreur fréquente : souligner pour mettre en valeur. À l'écran, le souligné évoque un lien hypertexte. Préférez le gras ou la couleur d'accent.</div>`
      },
      {
        title: "Formes, icônes et zones de texte",
        html: `<p>Pour illustrer un processus ou mettre un chiffre en valeur, utilisez des <b>formes</b> plutôt que des cliparts :</p>
<div class="o-path">Insertion › Illustrations › Formes</div>
<ul>
<li>Maintenir <kbd>Maj</kbd> en traçant : cercle parfait, carré parfait, trait droit.</li>
<li>Une forme sélectionnée accepte du texte : il suffit de taper.</li>
<li>L'onglet contextuel <b>Format de forme</b> apparaît : remplissage, contour, effets, taille exacte.</li>
<li><i>Insertion › Icônes</i> propose des milliers de pictogrammes vectoriels recolorables.</li>
</ul>
<div class="callout">Une <b>zone de texte</b> (<i>Insertion › Zone de texte</i>) est utile pour une légende ou une annotation, mais ne remplace pas les espaces réservés de la disposition.</div>
<div class="o-tip">💡 Mise en forme réussie ? <i>Clic droit › Définir comme forme par défaut</i> : toutes les formes suivantes la reprendront.</div>`
      },
      {
        title: "Aligner, répartir, reproduire",
        html: `<p>Rien ne trahit plus un deck amateur que des éléments « presque » alignés. Utilisez les outils plutôt que l'œil :</p>
<div class="o-path">Format de forme › Organiser › Aligner › Aligner au centre</div>
<ul>
<li><b>Aligner sur la diapositive</b> (une seule forme) ou <b>Aligner les objets sélectionnés</b> (plusieurs).</li>
<li><b>Distribuer horizontalement / verticalement</b> : espacements identiques.</li>
<li>Les <b>repères actifs</b> (pointillés rouges) apparaissent quand vous déplacez un objet : ils signalent centres et alignements.</li>
<li><kbd>Maj</kbd>+<kbd>F9</kbd> affiche la grille, <kbd>Alt</kbd>+<kbd>F9</kbd> les repères.</li>
</ul>
<p><b>Reproduire la mise en forme</b> (pinceau, <i>Accueil › Presse-papiers</i>) : un clic = une application, double-clic = mode verrouillé jusqu'à <kbd>Échap</kbd>. Clavier : <kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>C</kbd> puis <kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>V</kbd>.</p>
<div class="callout warn">Piège : « Aligner au centre » avec <i>Aligner les objets sélectionnés</i> coché aligne les objets entre eux, pas sur la diapo. Vérifiez l'option en bas du menu Aligner.</div>`
      }
    ],
    shortcuts: [
      { keys: "Ctrl+G", desc: "Gras" },
      { keys: "Ctrl+I", desc: "Italique" },
      { keys: "Ctrl+E", desc: "Centrer le paragraphe" },
      { keys: "Ctrl+Espace", desc: "Effacer la mise en forme des caractères" },
      { keys: "Ctrl+Maj+C", desc: "Copier la mise en forme" },
      { keys: "Ctrl+Maj+V", desc: "Coller la mise en forme" },
      { keys: "Maj+F9", desc: "Afficher / masquer la grille" },
      { keys: "Alt+F9", desc: "Afficher / masquer les repères" }
    ],
    labs: [
      {
        id: "p02-lab1",
        title: "Lab 1 — Le flyer du Salon Talents RH",
        kind: "slides",
        scenario: `<p>Le service RH de <b>Néolia</b> (ETI de 900 salariés) organise le <b>Salon Talents 2026</b>. Le deck d'annonce, qui tournera sur les écrans d'accueil, est fade et mal aligné. Donnez-lui une identité : thème, titre centré, mot clé en gras, appel à l'action dans une forme, et une icône.</p>`,
        setup: {
          theme: "blanc",
          slides: [
            { layout: "title", title: "Salon Talents 2026", subtitle: "Rencontrez nos équipes et nos métiers" },
            { layout: "content", title: "Infos pratiques", body: "Date : 12 mars 2026\nLieu : siège de Néolia, Nantes\nEntrée libre de 9 h à 17 h" },
            { layout: "content", title: "Au programme", body: "Job dating\nAteliers CV\nVisite des ateliers" }
          ]
        },
        tasks: [
          {
            text: "Applique un <b>thème</b> autre que « blanc » à toute la présentation.",
            hint: "Onglet Création › Thèmes.",
            check: a => !!a.theme && N(a.theme) !== "blanc"
          },
          {
            text: "Sur « Infos pratiques », mets en <b>gras</b> l'élément qui contient la date (12 mars).",
            hint: "Sélectionne la zone de texte, puis bouton Gras.",
            check: a => elsWith(a, "12 mars").some(r => r.el.bold === true)
          },
          {
            text: "Ajoute une <b>forme</b> (rectangle arrondi par exemple) contenant le texte <b>Inscrivez-vous</b>.",
            hint: "Insère une forme puis double-clique dessus pour taper le texte.",
            check: a => elsWith(a, "inscri").some(r => r.el.type === "shape")
          },
          {
            text: "Descends la forme « Inscrivez-vous » dans la <b>moitié basse</b> de la diapositive, puis <b>centre-la horizontalement</b> sur la diapo.",
            hint: "Glisse la forme vers le bas (ou Maj+Flèche bas), puis Accueil › Aligner sur la diapositive › Centrer horizontalement.",
            check: a => elsWith(a, "inscri").some(r => r.el.type === "shape" && num(r.el.y) >= 50 && a.centered(r.el) === true)
          },
          {
            text: "Illustre la diapo « Au programme » avec une <b>icône</b>.",
            hint: "Bouton Icône du simulateur, puis choisis un pictogramme.",
            check: a => { const i = idxTitle(a, "au programme"); return i >= 0 && byType(S(a)[i], "icon").length >= 1; }
          }
        ],
        solution: `<ol>
<li><i>Création › Thèmes</i> : choisir par exemple « ocean ».</li>
<li>Diapo 2 : sélectionner le corps contenant « Date : 12 mars 2026 » › <kbd>Ctrl</kbd>+<kbd>G</kbd>.</li>
<li>Insérer un rectangle arrondi, double-cliquer et taper « Inscrivez-vous sur talents.neolia.fr ».</li>
<li>Faire glisser la forme vers le bas de la diapo, puis <i>Accueil › Aligner sur la diapositive › Centrer horizontalement</i>.</li>
<li>Diapo 3 : insérer une icône (🤝 ou 🎓) à droite de la liste.</li>
</ol>`
      }
    ],
    quiz: [
      { q: "Que contient un thème PowerPoint ?", choices: ["Uniquement une image de fond", "Couleurs, polices, effets et styles d'arrière-plan", "Les animations et transitions", "Le texte des diapositives"], answer: 1, explain: "Le thème définit l'identité visuelle : palette, polices titre/corps, effets et arrière-plans." },
      { q: "Pourquoi choisir ses couleurs dans la ligne « Couleurs du thème » ?", choices: ["Elles sont plus vives", "Elles suivent automatiquement si on change de thème ou de variante", "Elles s'impriment mieux", "Elles sont obligatoires"], answer: 1, explain: "Les couleurs du thème sont des références : elles se mettent à jour avec la charte. Les couleurs standard sont figées." },
      { q: "Quel est le raccourci du gras dans PowerPoint en français ?", choices: ["Ctrl+B", "Ctrl+G", "Ctrl+Maj+G", "Alt+G"], answer: 1, explain: "Dans Office en français, Ctrl+G = Gras, Ctrl+I = Italique, Ctrl+U = Souligné." },
      { q: "Comment espacer régulièrement 4 formes alignées horizontalement ?", choices: ["À l'œil avec les flèches", "Format de forme › Aligner › Distribuer horizontalement", "Les grouper", "Les mettre dans un tableau"], answer: 1, explain: "Distribuer horizontalement calcule des écarts identiques entre les objets sélectionnés." },
      { q: "Comment appliquer la même mise en forme à plusieurs objets d'affilée avec le pinceau ?", choices: ["Clic simple sur le pinceau", "Double-clic sur Reproduire la mise en forme, puis Échap pour arrêter", "Clic droit › Coller", "Ctrl+Maj+F"], answer: 1, explain: "Le double-clic verrouille l'outil : on peut l'appliquer autant de fois que nécessaire, Échap pour sortir." }
    ]
  },

  /* ======================================================================
     p03 — Niveau 2 : Structurer un message
     ====================================================================== */
  {
    id: "p03",
    level: 2,
    title: "Structurer un message : storytelling, 1 idée par diapo, sections",
    icon: "🧱",
    duration: 60,
    summary: "Passer d'un catalogue de puces à une histoire claire qui mène le public jusqu'à une décision.",
    objectives: [
      "Être capable de construire un fil narratif (situation, complication, résolution)",
      "Être capable d'écrire des titres-messages et d'appliquer la règle « une idée par diapo »",
      "Être capable d'organiser un deck en sections et de le relire en mode Plan",
      "Être capable de conclure par un appel à l'action explicite"
    ],
    lessons: [
      {
        title: "Partir du message, pas des diapositives",
        html: `<p>Avant d'ouvrir PowerPoint, répondez à trois questions :</p>
<ol>
<li><b>Qui</b> est le public et que sait-il déjà ?</li>
<li><b>Que</b> doit-il retenir en une phrase (le « message clé ») ?</li>
<li><b>Quoi faire</b> après : valider, financer, changer une pratique ?</li>
</ol>
<p>Le schéma classique <b>Situation › Complication › Résolution</b> (SCR) fonctionne pour presque tout : « Nos délais de livraison sont de 48 h (situation). Depuis janvier, 1 commande sur 5 arrive en retard (complication). Un second quai de chargement règle le problème pour 80 k€ (résolution). »</p>
<div class="callout ok">Bonne pratique : écrivez votre histoire en <b>mode Plan</b> (<i>Affichage › Mode Plan</i>), uniquement avec les titres. Si la suite des titres raconte l'histoire à elle seule, la structure est bonne.</div>`
      },
      {
        title: "Titres-messages et règle « une idée par diapo »",
        html: `<p>Un titre-thème (« Résultats T3 ») oblige le public à deviner. Un <b>titre-message</b> donne la conclusion : « Le T3 dépasse l'objectif de 12 % grâce au e-commerce ».</p>
<table class="o-table"><thead><tr><th>Titre-thème</th><th>Titre-message</th></tr></thead><tbody>
<tr><td>Budget</td><td>Le projet tient dans l'enveloppe de 80 k€</td></tr>
<tr><td>Planning</td><td>Mise en service possible dès septembre</td></tr>
<tr><td>Risques</td><td>Deux risques à surveiller, aucun bloquant</td></tr>
</tbody></table>
<p><b>Une idée par diapo</b> : si une diapo défend deux idées, faites-en deux. Une diapo de plus ne coûte rien ; une diapo illisible coûte l'attention du public.</p>
<div class="callout warn">Piège : la diapo « mur de texte » qui sert à la fois de support oral et de document à envoyer. Faites deux livrables, ou mettez le détail dans les notes et exportez-les.</div>
<div class="o-tip">💡 Test des 3 secondes : un spectateur doit comprendre l'idée d'une diapo en 3 secondes. Sinon, simplifiez.</div>`
      },
      {
        title: "Organiser avec les sections",
        html: `<p>Les <b>sections</b> regroupent les diapositives en chapitres dans le volet des miniatures : on peut les réduire, les déplacer d'un bloc, les renommer, et même ne projeter qu'une section.</p>
<div class="o-path">Accueil › Diapositives › Section › Ajouter une section</div>
<ul>
<li>Clic droit sur une section : <i>Renommer, Supprimer la section, Monter/Descendre la section</i>.</li>
<li>En <b>Trieuse de diapositives</b>, les sections apparaissent comme des intertitres : vue d'ensemble immédiate.</li>
<li>Complétez par des <b>diapos de section</b> (disposition <i>Titre de section</i>) : elles annoncent au public le changement de partie.</li>
</ul>
<div class="callout">Les sections sont invisibles en diaporama ; les diapos « Titre de section », elles, sont projetées. Les deux se complètent.</div>`
      },
      {
        title: "Ouvrir et conclure",
        html: `<ul>
<li><b>Ouverture</b> : un fait marquant, une question, un chiffre. Puis le sommaire en 3-4 points maximum, en 2e position.</li>
<li><b>Corps</b> : chaque partie = une diapo de section + 2 à 5 diapos d'idées.</li>
<li><b>Conclusion</b> : rappel du message clé + <b>appel à l'action</b> (« Nous demandons votre décision sur… »). Évitez de finir sur « Merci / Questions ? » seul : c'est la diapo qui reste affichée le plus longtemps.</li>
</ul>
<div class="callout bad">Erreur fréquente : terminer sur une diapo de remerciements générique. Laissez plutôt affichée la diapo « Décision attendue » pendant les questions.</div>
<div class="o-tip">💡 Pendant le diaporama, tapez le numéro d'une diapo puis <kbd>Entrée</kbd> pour y sauter directement — utile pour revenir à la décision lors des questions.</div>`
      }
    ],
    shortcuts: [
      { keys: "Ctrl+Flèche haut", desc: "Monter la diapositive sélectionnée" },
      { keys: "Ctrl+Flèche bas", desc: "Descendre la diapositive sélectionnée" },
      { keys: "Alt+Maj+Flèche gauche", desc: "Remonter le niveau (mode Plan)" },
      { keys: "Alt+Maj+Flèche droite", desc: "Abaisser le niveau (mode Plan)" },
      { keys: "Ctrl+Entrée", desc: "Espace réservé suivant / nouvelle diapo" },
      { keys: "Maj+Entrée", desc: "Saut de ligne sans nouvelle puce" }
    ],
    labs: [
      {
        id: "p03-lab1",
        title: "Lab 1 — Le second quai de chargement",
        kind: "slides",
        scenario: `<p>Karim, responsable logistique chez <b>TransAlp</b>, doit obtenir du comité de direction un budget de <b>80 k€</b> pour un second quai de chargement. Son brouillon tient en une diapo « fourre-tout » de 9 puces, le sommaire est à la fin et rien ne conclut. Restructurez : sommaire en 2e position, sections, une idée par diapo et une demande de décision explicite.</p>`,
        setup: {
          theme: "ocean",
          slides: [
            { layout: "title", title: "Projet second quai", subtitle: "Comité de direction — avril 2026" },
            { layout: "content", title: "Tout sur le projet", body: "Délai actuel 48 h\n1 commande sur 5 en retard depuis janvier\nCamions en attente 40 min en moyenne\nPénalités clients : 6 k€ par mois\nSolution : second quai\nCoût : 80 k€\nTravaux : 6 semaines en été\nRetour sur investissement : 14 mois\nRisque : permis de construire" },
            { layout: "content", title: "Sommaire", body: "Le constat\nLa solution\nLa décision attendue" }
          ]
        },
        tasks: [
          {
            text: "Déplace la diapo <b>Sommaire</b> en <b>2e position</b> (juste après le titre).",
            hint: "Glisse la miniature ou utilise le bouton « monter ».",
            check: a => titleOf(S(a)[1]).includes("sommaire")
          },
          {
            text: "Ajoute au moins <b>2 diapos de section</b> (disposition <i>Titre de section</i>), par exemple « Le constat » et « La solution ».",
            hint: "Nouvelle diapo › Disposition › Section.",
            check: a => S(a).filter(s => lay(s) === "section" && titleOf(s).length > 0).length >= 2
          },
          {
            text: "Fini le fourre-tout : <b>aucune</b> zone de texte ne doit dépasser <b>5 puces</b>, et le deck doit compter au moins <b>7 diapositives</b>.",
            hint: "Duplique la diapo de 9 puces, puis garde 3-4 puces sur chaque copie avec un titre adapté.",
            check: a => a.count() >= 7 && S(a).every(s => E(s).every(e => !(e.type === "body" || e.type === "text" || e.type === "shape") || lines(e.text).length <= 5))
          },
          {
            text: "Écris au moins un <b>titre-message</b> chiffré : un titre contenant <b>80 k€</b> (ex. « Un second quai pour 80 k€, rentabilisé en 14 mois »).",
            hint: "Le titre doit contenir « 80 k€ » (ou 80 k).",
            check: a => S(a).some(s => /80 ?k/.test(titleOf(s)))
          },
          {
            text: "Termine par un <b>appel à l'action</b> : le <b>titre</b> de la <b>dernière</b> diapo doit contenir le mot <b>décision</b>.",
            hint: "Ex. titre « Décision attendue aujourd'hui », puces : budget 80 k€, lancement des travaux en juillet.",
            check: a => a.count() > 0 && titleOf(S(a)[a.count() - 1]).includes("decision")
          }
        ],
        solution: `<ol>
<li>Ordre final : Titre › Sommaire › [Section] Le constat › « 1 commande sur 5 arrive en retard » › « Les retards nous coûtent 6 k€ par mois » › [Section] La solution › « Un second quai pour 80 k€, rentabilisé en 14 mois » › « 6 semaines de travaux cet été, un seul risque » › « Décision attendue aujourd'hui ».</li>
<li>Chaque diapo d'idée garde 2 à 4 puces issues de la diapo d'origine, qui est supprimée ou réduite.</li>
<li>La dernière diapo : « Décision attendue : valider 80 k€ pour un démarrage en juillet ».</li>
</ol>`
      }
    ],
    quiz: [
      { q: "Quel titre est un titre-message ?", choices: ["Résultats", "Ventes T3", "Le T3 dépasse l'objectif de 12 %", "Point sur les ventes"], answer: 2, explain: "Un titre-message énonce la conclusion de la diapo, le public n'a pas à la deviner." },
      { q: "Dans le schéma SCR, que représente le « C » ?", choices: ["Conclusion", "Complication", "Contexte", "Chiffres"], answer: 1, explain: "Situation › Complication (le problème, la tension) › Résolution (votre proposition)." },
      { q: "Où créer une section dans PowerPoint 365 ?", choices: ["Insertion › Section", "Accueil › Diapositives › Section › Ajouter une section", "Création › Section", "Affichage › Section"], answer: 1, explain: "La commande Section se trouve dans le groupe Diapositives de l'onglet Accueil (ou clic droit entre deux miniatures)." },
      { q: "Les sections sont-elles visibles par le public en diaporama ?", choices: ["Oui, en bandeau", "Non, ce sont des repères d'organisation ; il faut une diapo Titre de section pour l'annoncer", "Oui, à chaque transition", "Seulement en mode présentateur"], answer: 1, explain: "Les sections organisent le fichier. Pour le public, on ajoute une diapo de disposition Titre de section." },
      { q: "Quelle diapo vaut-il mieux laisser affichée pendant les questions ?", choices: ["« Merci ! »", "Le sommaire", "La diapo de décision / appel à l'action", "Une diapo noire"], answer: 2, explain: "La diapo finale reste longtemps à l'écran : elle doit rappeler ce que vous attendez du public." }
    ]
  },

  /* ======================================================================
     p04 — Niveau 2 : Animations et transitions
     ====================================================================== */
  {
    id: "p04",
    level: 2,
    title: "Animations et transitions maîtrisées",
    icon: "✨",
    duration: 55,
    summary: "Utiliser le mouvement pour guider le regard, pas pour distraire : sobriété, ordre et déclencheurs.",
    objectives: [
      "Être capable de choisir une transition sobre et de l'appliquer à tout le deck",
      "Être capable d'animer des éléments dans un ordre qui sert le discours",
      "Être capable de régler le démarrage (au clic, avec, après) et la durée",
      "Être capable d'utiliser le volet Animation et les déclencheurs"
    ],
    lessons: [
      {
        title: "Transitions : une seule, discrète",
        html: `<p>La <b>transition</b> est l'effet entre deux diapositives. Elle se règle dans l'onglet dédié :</p>
<div class="o-path">Transitions › Accès à cette diapositive › Fondu</div>
<ul>
<li>Choisissez <b>une</b> transition sobre (<i>Fondu</i>, <i>Poussée</i>, ou <i>Aucune</i>) puis cliquez <b>Appliquer partout</b> (groupe <i>Minutage</i>).</li>
<li><b>Durée</b> : 0,5 à 1 s. Au-delà, le public attend.</li>
<li>Réservez une transition différente (<i>Morphose</i>, <i>Zoom</i>) à un moment fort, par exemple l'entrée dans une nouvelle section.</li>
</ul>
<div class="callout bad">Erreur fréquente : une transition différente par diapo (Vortex, Rideaux, Origami…). L'effet « catalogue » décrédibilise immédiatement le propos.</div>`
      },
      {
        title: "Les 4 familles d'animation",
        html: `<table class="o-table"><thead><tr><th>Famille</th><th>Couleur</th><th>Usage</th></tr></thead><tbody>
<tr><td>Entrée</td><td>verte</td><td>Faire apparaître un élément (Apparition, Fondu, Zoom, Balayer)</td></tr>
<tr><td>Emphase</td><td>jaune</td><td>Attirer l'attention sur un élément déjà visible (Impulsion, Couleur)</td></tr>
<tr><td>Sortie</td><td>rouge</td><td>Faire disparaître un élément</td></tr>
<tr><td>Trajectoire</td><td>—</td><td>Déplacer un élément le long d'un chemin</td></tr>
</tbody></table>
<div class="o-path">Animations › Animation › Ajouter une animation</div>
<div class="callout warn">Piège : choisir un effet dans la galerie <b>remplace</b> l'animation existante. Pour cumuler deux effets sur un même objet (entrée puis emphase), utilisez <b>Ajouter une animation</b>.</div>
<div class="callout ok">Règle d'or : une animation doit avoir une <b>raison</b> — révéler pas à pas, souligner, montrer une étape. Si vous ne pouvez pas dire pourquoi, supprimez-la.</div>`
      },
      {
        title: "Ordre, démarrage et volet Animation",
        html: `<p>Le <b>volet Animation</b> (<i>Animations › Animation avancée › Volet Animation</i>) liste tous les effets de la diapo dans leur ordre d'exécution. Vous pouvez les glisser pour les réordonner.</p>
<ul>
<li><b>Au clic</b> : attend votre clic (ou <kbd>Espace</kbd>/<kbd>N</kbd>).</li>
<li><b>Avec la précédente</b> : démarre en même temps que l'effet précédent.</li>
<li><b>Après la précédente</b> : démarre quand le précédent se termine (enchaînement automatique).</li>
</ul>
<div class="o-path">Animations › Minutage › Démarrer › Après la précédente</div>
<p>Pour une liste à puces : <i>Options d'effet › Par paragraphe</i> fait apparaître une puce par clic.</p>
<div class="o-tip">💡 <b>Reproduire l'animation</b> (<i>Animations › Animation avancée</i>) copie un effet et tous ses réglages sur un autre objet, comme le pinceau pour la mise en forme.</div>`
      },
      {
        title: "Déclencheurs : l'interactivité simple",
        html: `<p>Un <b>déclencheur</b> lance une animation quand on clique sur un objet précis (et non n'importe où). Idéal pour un quiz ou une FAQ : un clic sur « Réponse A » révèle l'explication.</p>
<div class="o-path">Animations › Animation avancée › Déclencheur › Sur clic de</div>
<ul>
<li>Nommez vos objets dans le <b>volet Sélection</b> (<kbd>Alt</kbd>+<kbd>F10</kbd>) : « Bouton A », « Réponse A »… sinon la liste « Rectangle 12, Rectangle 13 » est illisible.</li>
<li><i>Sur signet</i> déclenche un effet à un moment précis d'une vidéo.</li>
</ul>
<div class="callout warn">Piège : avec un déclencheur, cliquer ailleurs sur la diapo fait passer à la suivante. Prévenez-vous dans les notes et testez en diaporama.</div>`
      }
    ],
    shortcuts: [
      { keys: "Alt+F10", desc: "Volet Sélection (nommer, masquer, réordonner les objets)" },
      { keys: "Maj+F5", desc: "Tester la diapo active en diaporama" },
      { keys: "Espace", desc: "Diaporama : animation ou diapo suivante" },
      { keys: "P", desc: "Diaporama : étape précédente" },
      { keys: "Ctrl+Z", desc: "Annuler la dernière animation ajoutée" },
      { keys: "F6", desc: "Passer au volet suivant (ruban, miniatures, notes)" }
    ],
    labs: [
      {
        id: "p04-lab1",
        title: "Lab 1 — Le processus de l'entrepôt",
        kind: "slides",
        scenario: `<p>L'entrepôt <b>LogiNord</b> forme ses intérimaires avec un deck projeté en salle de pause. Le responsable veut que le processus apparaisse <b>étape par étape</b>, que le chiffre clé ressorte et que l'enchaînement soit sobre et uniforme.</p>`,
        setup: {
          theme: "foret",
          slides: [
            { layout: "title", title: "Bienvenue chez LogiNord", subtitle: "Formation intérimaires" },
            { layout: "title", title: "Processus d'un colis", subtitle: "3 étapes, 1 objectif : zéro erreur" },
            { layout: "content", title: "Notre engagement", body: "Taux de service : 98 %\nPréparation en moins de 2 h\nZéro colis perdu" },
            { layout: "content", title: "Sécurité", body: "Chaussures obligatoires\nGilet haute visibilité\nAllées piétonnes" }
          ]
        },
        tasks: [
          {
            text: "Applique la <b>même transition sobre</b> (fondu) à <b>toutes</b> les diapositives.",
            hint: "Choisis Fondu puis applique-le partout (ou diapo par diapo).",
            check: a => { const t = S(a).map(s => N(s.transition)); return t.length > 0 && t.every(x => on(x) && x === t[0] && (x.includes("fond") || x.includes("fade"))); }
          },
          {
            text: "Sur la diapo « Processus d'un colis », ajoute <b>3 formes</b> : <b>Réception</b>, <b>Stockage</b>, <b>Expédition</b>.",
            hint: "Insère 3 formes (ou flèches) et tape le texte dedans.",
            check: a => { const i = idxTitle(a, "processus"); if (i < 0) return false; const t = slideText(S(a)[i]); return ["reception", "stockage", "expedition"].every(k => t.includes(k)); }
          },
          {
            text: "Anime ces 3 étapes pour qu'elles apparaissent dans l'<b>ordre</b> Réception › Stockage › Expédition.",
            hint: "Donne une animation d'entrée à chaque forme et règle l'ordre 1, 2, 3.",
            check: a => {
              const i = idxTitle(a, "processus"); if (i < 0) return false;
              const els = E(S(a)[i]);
              const f = k => els.find(e => N(e.text).includes(k));
              const r = f("reception"), s = f("stockage"), x = f("expedition");
              if (!r || !s || !x || ![r, s, x].every(e => on(e.anim))) return false;
              return num(r.order) < num(s.order) && num(s.order) < num(x.order);
            }
          },
          {
            text: "Sur « Notre engagement », fais ressortir le chiffre <b>98 %</b> avec une animation <b>Zoom</b>.",
            hint: "Sélectionne l'élément contenant 98 % › Animation › Zoom.",
            check: a => elsWith(a, "98").some(r => N(r.el.anim).includes("zoom"))
          },
          {
            text: "La diapo de <b>titre</b> ne doit porter <b>aucune</b> animation d'élément, et teste le résultat en <b>diaporama</b>.",
            hint: "Vérifie que le titre et le sous-titre de la diapo 1 sont sur « aucune », puis lance le diaporama.",
            check: a => a.played === true && E(S(a)[0]).every(e => !on(e.anim))
          }
        ],
        solution: `<ol>
<li>Transitions : sélectionner chaque diapo (ou toutes) › Fondu.</li>
<li>Diapo 2 : trois rectangles arrondis (ou flèches) alignés horizontalement, textes « Réception », « Stockage », « Expédition ».</li>
<li>Animer chaque forme (Entrée ou Apparition) ; ordre 1 = Réception, 2 = Stockage, 3 = Expédition.</li>
<li>Diapo 3 : animation Zoom sur le corps contenant « Taux de service : 98 % ».</li>
<li>Diapo 1 sans animation ; lancer le diaporama et vérifier l'enchaînement.</li>
</ol>`
      }
    ],
    quiz: [
      { q: "Quelle est la meilleure pratique pour les transitions ?", choices: ["Une transition différente par diapo", "Une transition sobre appliquée partout, une exception pour un moment fort", "Toujours Vortex pour dynamiser", "Durée de 3 secondes minimum"], answer: 1, explain: "La cohérence rassure ; l'exception ne fonctionne que si elle est rare." },
      { q: "Comment cumuler une animation d'entrée ET une emphase sur la même forme ?", choices: ["Cliquer deux effets dans la galerie", "Utiliser Ajouter une animation pour le second effet", "Dupliquer la forme", "C'est impossible"], answer: 1, explain: "La galerie remplace l'effet existant ; Ajouter une animation l'empile." },
      { q: "Quelle option enchaîne automatiquement un effet à la fin du précédent ?", choices: ["Au clic", "Avec la précédente", "Après la précédente", "Sur signet"], answer: 2, explain: "Après la précédente démarre quand l'effet précédent est terminé ; Avec la précédente le lance en même temps." },
      { q: "À quoi sert un déclencheur « Sur clic de » ?", choices: ["À lancer le diaporama", "À lancer une animation quand on clique sur un objet précis", "À changer de thème", "À enregistrer le minutage"], answer: 1, explain: "C'est la base de l'interactivité : quiz, FAQ, révélation à la demande." },
      { q: "Pourquoi renommer les objets dans le volet Sélection avant de créer des déclencheurs ?", choices: ["Pour l'accessibilité uniquement", "Parce que la liste des déclencheurs affiche les noms d'objets (Rectangle 12…)", "C'est obligatoire", "Pour réduire la taille du fichier"], answer: 1, explain: "Des noms explicites (« Bouton A ») évitent de se tromper d'objet dans le menu Déclencheur et le volet Animation." }
    ]
  },

  /* ======================================================================
     p05 — Niveau 3 : Masque et modèles
     ====================================================================== */
  {
    id: "p05",
    level: 3,
    title: "Masque des diapositives et modèles d'entreprise",
    icon: "🏛️",
    duration: 65,
    summary: "Industrialiser la charte graphique : un masque bien construit, des dispositions sur mesure et un modèle .potx partagé.",
    objectives: [
      "Être capable de modifier le masque et les dispositions sans toucher aux diapositives",
      "Être capable de créer une disposition personnalisée avec ses espaces réservés",
      "Être capable de gérer pied de page, numéro et logo de façon centralisée",
      "Être capable d'enregistrer et de diffuser un modèle d'entreprise (.potx)"
    ],
    lessons: [
      {
        title: "Comprendre le masque",
        html: `<p>Le <b>masque des diapositives</b> est la diapo « mère » : tout ce qui y est placé (polices, couleurs, logo, positions) se répercute sur les dispositions qui en dépendent, puis sur les diapositives.</p>
<div class="o-path">Affichage › Modes masque › Masque des diapositives</div>
<ul>
<li>La <b>grande miniature</b> en haut = le masque (règles communes).</li>
<li>Les <b>petites miniatures</b> dessous = les dispositions (Titre, Titre et contenu, Section…). Survolez-les : une info-bulle indique quelles diapositives les utilisent.</li>
<li><i>Fermer le mode Masque</i> pour revenir.</li>
</ul>
<div class="callout ok">Bonne pratique : une modification qui doit concerner <b>toutes</b> les diapos (police des titres, logo) se fait sur le masque ; une modification propre à un type de diapo se fait sur sa disposition.</div>
<div class="callout bad">Erreur fréquente : copier-coller le logo sur chacune des 60 diapos. Au changement de logo, il faut tout refaire — et l'un d'eux sera toujours décalé.</div>`
      },
      {
        title: "Créer des dispositions sur mesure",
        html: `<p>En mode masque : <i>Insérer une disposition</i>, puis <i>Insérer un espace réservé</i> (Contenu, Texte, Image, Graphique, Tableau, SmartArt, Média…).</p>
<div class="o-path">Masque des diapositives › Mise en page du masque › Insérer un espace réservé</div>
<ul>
<li>Renommez la disposition (clic droit › <i>Renommer la disposition</i>) : « Chiffre clé », « Témoignage », « 3 colonnes »…</li>
<li>Un espace réservé <b>Image</b> recadre automatiquement la photo insérée : fini les images déformées.</li>
<li>Tapez un texte d'invite parlant dans l'espace réservé (« Saisir le chiffre clé, ex. +12 % »).</li>
</ul>
<div class="callout warn">Piège : supprimer une disposition utilisée par des diapos n'est pas possible ; et un masque surchargé de 40 dispositions inutiles perd tout le monde. Faites le ménage (clic droit › <i>Supprimer la disposition</i>) avant de diffuser.</div>`
      },
      {
        title: "Pied de page, numéros et logo",
        html: `<p>Les zones de pied de page existent déjà sur le masque ; il faut les <b>activer</b> :</p>
<div class="o-path">Insertion › Texte › En-tête/Pied de page</div>
<ul>
<li>Cochez <b>Numéro de diapositive</b>, <b>Pied de page</b> (« Confidentiel — Néolia ») et éventuellement <b>Date</b>.</li>
<li><b>Ne pas afficher sur la diapositive titre</b> : l'option qu'on oublie.</li>
<li><i>Appliquer partout</i> plutôt que <i>Appliquer</i> (diapo seule).</li>
</ul>
<p>Le logo se colle sur le masque (en haut ou en bas, à taille raisonnable). Pour une disposition sans logo (diapo plein écran photo), cochez <i>Masquer les graphiques d'arrière-plan</i> dans <i>Mettre en forme l'arrière-plan</i>.</p>
<div class="o-tip">💡 Le numéro de diapo n'est pas un luxe : en réunion, « revenons à la 14 » est bien plus efficace que « la diapo bleue avec le camembert ».</div>`
      },
      {
        title: "Enregistrer et diffuser un modèle .potx",
        html: `<div class="o-path">Fichier › Enregistrer sous › Modèle PowerPoint (*.potx)</div>
<ul>
<li>Un <b>.potx</b> s'ouvre comme une <b>nouvelle présentation</b> sans titre : impossible d'écraser le modèle par erreur.</li>
<li>Dans Microsoft 365, l'administrateur peut publier les modèles dans une bibliothèque SharePoint ou un dossier de <b>modèles d'organisation</b> : ils apparaissent dans <i>Fichier › Nouveau</i>, onglet de l'organisation.</li>
<li>Pour un usage personnel, définissez un dossier de modèles par défaut : <i>Fichier › Options › Enregistrement › Dossier par défaut des modèles personnels</i>.</li>
</ul>
<p>Appliquer un modèle à un deck existant : <i>Création › Thèmes › Autres › Rechercher des thèmes</i>, puis choisir le .potx : le masque est remplacé, et les diapos qui utilisent des dispositions de même nom suivent.</p>
<div class="callout warn">Piège : réutiliser des diapos d'un ancien deck par copier-coller peut importer un <b>second masque</b>. Collez avec <i>Utiliser le thème de destination</i> (icône du collage) et surveillez le mode masque.</div>`
      }
    ],
    shortcuts: [
      { keys: "Maj+F9", desc: "Grille (placer précisément les espaces réservés)" },
      { keys: "Alt+F9", desc: "Repères" },
      { keys: "Alt+F10", desc: "Volet Sélection" },
      { keys: "Ctrl+Maj+C", desc: "Copier la mise en forme" },
      { keys: "Ctrl+Maj+V", desc: "Coller la mise en forme" },
      { keys: "Ctrl+Y", desc: "Répéter la dernière action" }
    ],
    labs: [
      {
        id: "p05-lab1",
        title: "Lab 1 — Remettre un deck dans la charte",
        kind: "slides",
        scenario: `<p>Chez <b>Assurances Rivage</b>, la charte impose le thème bleu « ocean », des dispositions standard et un sous-titre daté. Un commercial a présenté son deck « à sa façon » : disposition inadaptée, comparatif écrit en une colonne. Remettez-le aux normes <b>en utilisant les dispositions</b>, pas en redessinant.</p>`,
        setup: {
          theme: "sunset",
          slides: [
            { layout: "title", title: "Offre Habitation Sérénité", subtitle: "Réunion commerciale" },
            { layout: "content", title: "Partie 2 — Déploiement", body: "" },
            { layout: "content", title: "Comparatif des formules", body: "Essentielle : 18 €/mois, dégâts des eaux, incendie. Confort : 29 €/mois, + vol, + bris de glace" },
            { layout: "content", title: "Calendrier", body: "Formation des agents : mai\nLancement : 1er juin\nBilan : septembre" }
          ]
        },
        tasks: [
          {
            text: "Applique le thème de la charte : <b>ocean</b>.",
            hint: "Création › Thèmes › ocean.",
            check: a => N(a.theme) === "ocean"
          },
          {
            text: "La diapo « Partie 2 — Déploiement » est un intertitre : passe-la en disposition <b>Titre de section</b>.",
            hint: "Disposition › Section.",
            check: a => { const i = idxTitle(a, "deploiement"); return i >= 0 && lay(S(a)[i]) === "section"; }
          },
          {
            text: "Passe le « Comparatif des formules » en <b>Deux contenus</b> : Essentielle à gauche, Confort à droite.",
            hint: "Disposition › Deux contenus, puis répartis le texte dans les deux colonnes.",
            check: a => { const i = idxTitle(a, "comparatif"); if (i < 0) return false; const s = S(a)[i], b = byType(s, "body").map(e => N(e.text)); return lay(s) === "two" && b.some(t => t.includes("essentielle") && !t.includes("confort")) && b.some(t => t.includes("confort") && !t.includes("essentielle")); }
          },
          {
            text: "Complète le sous-titre de la diapo 1 avec l'année <b>2026</b> (ex. « Réunion commerciale — mai 2026 »).",
            hint: "Double-clique dans le sous-titre.",
            check: a => byType(S(a)[0], "subtitle").some(e => N(e.text).includes("2026"))
          }
        ],
        solution: `<ol>
<li><i>Création › Thèmes</i> › ocean.</li>
<li>Diapo 2 : <i>Disposition › Titre de section</i>.</li>
<li>Diapo 3 : <i>Disposition › Deux contenus</i> ; gauche « Essentielle — 18 €/mois : dégâts des eaux, incendie », droite « Confort — 29 €/mois : + vol, + bris de glace ».</li>
<li>Sous-titre diapo 1 : « Réunion commerciale — mai 2026 ».</li>
</ol>
<p>Dans le vrai PowerPoint, le logo et le pied de page « Confidentiel » se poseraient sur le masque, pas sur chaque diapo (voir Lab 2).</p>`
      },
      {
        id: "p05-lab2",
        title: "Lab 2 — Construire le modèle .potx de l'entreprise",
        kind: "steps",
        scenario: `<p>Dans le <b>vrai PowerPoint</b>, créez le modèle officiel d'Assurances Rivage : masque, disposition « Chiffre clé », pied de page et numéros, enregistrement en .potx.</p>`,
        setup: { app: "PowerPoint" },
        tasks: [
          { text: "Nouvelle présentation vide › <i>Affichage › Masque des diapositives</i>. Sur le <b>masque</b> (grande miniature), choisissez des polices de titre et de corps dans <i>Polices</i>, et une palette bleue dans <i>Couleurs</i>.", hint: "Les menus Couleurs et Polices sont dans le groupe Arrière-plan de l'onglet Masque des diapositives." },
          { text: "Insérez un logo (une forme avec le texte « RIVAGE » suffit) en haut à droite du <b>masque</b> et vérifiez qu'il apparaît sur toutes les dispositions.", hint: "Insérez-le sur la grande miniature, pas sur une disposition." },
          { text: "<i>Insérer une disposition</i>, renommez-la <b>Chiffre clé</b>, ajoutez un espace réservé Texte centré en grand (72 pt) avec l'invite « Saisir le chiffre ».", hint: "Clic droit sur la disposition › Renommer la disposition." },
          { text: "Supprimez au moins 3 dispositions inutiles, puis fermez le mode masque.", hint: "Clic droit › Supprimer la disposition (impossible si elle est utilisée)." },
          { text: "<i>Insertion › En-tête/Pied de page</i> : numéro de diapositive + pied de page « Confidentiel — Assurances Rivage », cochez <b>Ne pas afficher sur la diapositive titre</b>, <i>Appliquer partout</i>.", hint: "Les zones de pied de page doivent exister sur le masque." },
          { text: "<i>Fichier › Enregistrer sous › Modèle PowerPoint (*.potx)</i>, puis rouvrez-le : vérifiez qu'il s'ouvre en « Présentation1 » non titrée.", hint: "PowerPoint propose automatiquement le dossier Modèles Office personnalisés." }
        ],
        solution: `<ol>
<li>Masque : polices et couleurs définies une seule fois, logo sur la grande miniature.</li>
<li>Disposition « Chiffre clé » avec espace réservé texte centré, invite explicite.</li>
<li>Dispositions inutilisées supprimées pour un masque lisible.</li>
<li>Pied de page et numéro activés partout sauf sur la diapo titre.</li>
<li>Enregistrement en .potx : chaque ouverture crée une nouvelle présentation basée sur le modèle.</li>
</ol>`
      }
    ],
    quiz: [
      { q: "Où modifier la police des titres pour toutes les diapositives d'un coup ?", choices: ["Sur chaque diapo", "Sur la grande miniature du masque des diapositives", "Dans Fichier › Options", "Dans le volet Notes"], answer: 1, explain: "La grande miniature est le masque : ses réglages se propagent à toutes les dispositions et diapositives." },
      { q: "Quel est l'intérêt d'un espace réservé Image dans une disposition ?", choices: ["Il compresse l'image", "Il recadre automatiquement l'image à la bonne taille et position", "Il ajoute un texte de remplacement", "Il anime l'image"], answer: 1, explain: "L'image insérée remplit l'espace réservé, recadrée : toutes les photos sont homogènes." },
      { q: "Quelle option oublie-t-on souvent dans En-tête/Pied de page ?", choices: ["Appliquer", "Ne pas afficher sur la diapositive titre", "Date fixe", "Aperçu"], answer: 1, explain: "Sans elle, le numéro « 1 » et le pied de page s'affichent sur la couverture." },
      { q: "Que se passe-t-il quand on ouvre un fichier .potx ?", choices: ["On modifie le modèle", "Une nouvelle présentation sans titre basée sur le modèle est créée", "Il s'ouvre en lecture seule", "Il lance le diaporama"], answer: 1, explain: "C'est le comportement d'un modèle : le fichier d'origine reste intact." },
      { q: "Vous collez des diapos d'un ancien deck et le volet masque montre deux masques. Comment l'éviter ?", choices: ["Coller avec Conserver la mise en forme source", "Coller avec Utiliser le thème de destination", "Coller en image", "Renommer le fichier"], answer: 1, explain: "Utiliser le thème de destination adapte les diapos au masque courant au lieu d'importer le masque d'origine." }
    ]
  },

  /* ======================================================================
     p06 — Niveau 3 : Présenter
     ====================================================================== */
  {
    id: "p06",
    level: 3,
    title: "Présenter : mode présentateur, notes, minutage et raccourcis",
    icon: "🎤",
    duration: 55,
    summary: "Être à l'aise le jour J : notes utiles, timing maîtrisé, navigation au clavier et réponses aux imprévus.",
    objectives: [
      "Être capable de rédiger des notes du présentateur exploitables",
      "Être capable d'utiliser le mode présentateur et ses outils",
      "Être capable de répéter et minuter sa présentation (Vérifier le minutage, Coach)",
      "Être capable de piloter le diaporama au clavier, y compris face aux imprévus"
    ],
    lessons: [
      {
        title: "Des notes qui aident vraiment",
        html: `<p>Les <b>notes du présentateur</b> se saisissent sous la diapo (bouton <i>Notes</i> de la barre d'état). Elles ne sont visibles que de vous, dans le mode présentateur.</p>
<ul>
<li>Écrivez des <b>repères</b>, pas un discours : la phrase d'accroche, les 2-3 chiffres à citer, la transition vers la diapo suivante.</li>
<li>Ajoutez des consignes de régie : « PAUSE », « Demander qui utilise déjà l'outil », « Montrer la démo ».</li>
<li>Relisez-les en <i>Affichage › Page de notes</i> ; imprimez-les si besoin (<i>Fichier › Imprimer › Pages de notes</i>).</li>
</ul>
<div class="callout warn">Piège : les notes partent avec le fichier. Si vous envoyez le .pptx à un client, il lira « Ne pas évoquer le retard du lot 3 ». Vérifiez-les ou utilisez <i>Fichier › Informations › Vérifier la présence de problèmes › Inspecter le document</i>.</div>`
      },
      {
        title: "Le mode présentateur",
        html: `<p>Avec deux écrans, PowerPoint affiche le diaporama au public et le <b>mode présentateur</b> sur votre écran : diapo en cours, diapo suivante, notes, chronomètre et heure.</p>
<div class="o-path">Diaporama › Moniteurs › Utiliser le mode Présentateur</div>
<ul>
<li><kbd>Alt</kbd>+<kbd>F5</kbd> lance directement le mode présentateur (pratique pour répéter avec un seul écran).</li>
<li>Outils : stylo et surligneur, <b>zoom</b> sur une partie de la diapo, vue de toutes les diapos, écran noir, sous-titres.</li>
<li><i>Paramètres d'affichage › Permuter le mode présentateur et le diaporama</i> si les écrans sont inversés.</li>
</ul>
<div class="o-tip">💡 Agrandissez la police des notes avec les boutons A+ / A- du mode présentateur : lisible à un mètre, debout.</div>`
      },
      {
        title: "Répéter et minuter",
        html: `<ul>
<li><b>Vérifier le minutage</b> (<i>Diaporama › Configuration › Vérifier le minutage</i>) : chronomètre chaque diapo pendant votre répétition et propose de conserver les durées.</li>
<li><b>Répéter avec Coach</b> (<i>Diaporama › Répéter avec Coach</i>) : l'assistant analyse votre débit, vos mots de remplissage (« euh », « en fait »), la lecture des diapos, et produit un rapport.</li>
<li><b>Configurer le diaporama</b> : <i>Visionné par une personne</i>, <i>Borne interactive</i> (boucle automatique), choix des diapositives à projeter, ignorer ou non le minutage.</li>
</ul>
<div class="o-path">Diaporama › Configuration › Configurer le diaporama</div>
<div class="callout warn">Piège : après une répétition minutée, les diapos avancent toutes seules en réunion. Décochez <i>Utiliser le minutage</i> si vous voulez garder la main.</div>
<div class="callout ok">Repère : 1 à 2 minutes par diapo d'idée. 20 minutes de créneau = 10 à 15 diapos, questions comprises.</div>`
      },
      {
        title: "Piloter le diaporama au clavier",
        html: `<table class="o-table"><thead><tr><th>Touche</th><th>Effet</th></tr></thead><tbody>
<tr><td><kbd>N</kbd>, <kbd>Espace</kbd>, <kbd>Entrée</kbd>, <kbd>Flèche droite</kbd></td><td>Étape / diapo suivante</td></tr>
<tr><td><kbd>P</kbd>, <kbd>Flèche gauche</kbd>, Retour arrière</td><td>Étape / diapo précédente</td></tr>
<tr><td>numéro + <kbd>Entrée</kbd></td><td>Aller à la diapo n°</td></tr>
<tr><td><kbd>B</kbd> ou <kbd>.</kbd></td><td>Écran noir (et retour)</td></tr>
<tr><td><kbd>W</kbd> ou <kbd>,</kbd></td><td>Écran blanc (et retour)</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>P</kbd> / <kbd>E</kbd></td><td>Stylo / effacer les annotations</td></tr>
<tr><td><kbd>Ctrl</kbd>+<kbd>S</kbd></td><td>Liste de toutes les diapositives</td></tr>
<tr><td><kbd>Échap</kbd></td><td>Quitter le diaporama</td></tr>
</tbody></table>
<div class="o-tip">💡 L'écran noir (<kbd>B</kbd>) ramène l'attention sur vous pendant une discussion. Re-<kbd>B</kbd> pour reprendre exactement où vous étiez.</div>
<div class="callout">En diaporama, <kbd>F1</kbd> affiche la liste complète des commandes disponibles.</div>`
      }
    ],
    shortcuts: [
      { keys: "F5", desc: "Diaporama depuis le début" },
      { keys: "Maj+F5", desc: "Diaporama depuis la diapo active" },
      { keys: "Alt+F5", desc: "Mode présentateur" },
      { keys: "B", desc: "Écran noir / retour" },
      { keys: "W", desc: "Écran blanc / retour" },
      { keys: "Ctrl+P", desc: "Stylo (en diaporama)" },
      { keys: "E", desc: "Effacer les annotations" },
      { keys: "Échap", desc: "Quitter le diaporama" }
    ],
    labs: [
      {
        id: "p06-lab1",
        title: "Lab 1 — Préparer l'intervention du séminaire",
        kind: "slides",
        scenario: `<p>Sophie, directrice commerciale de <b>VinoSud</b>, intervient 15 minutes au séminaire annuel des 120 commerciaux. Son deck est prêt, mais ses notes sont vides et rien n'est prévu pour la séance de questions. Préparez son « filet de sécurité ».</p>`,
        setup: {
          theme: "nuit",
          slides: [
            { layout: "title", title: "2026 : l'année de l'export", subtitle: "Séminaire commercial VinoSud" },
            { layout: "content", title: "L'export a porté notre croissance", body: "+18 % de CA à l'export\n3 nouveaux pays\n42 % du CA total" },
            { layout: "content", title: "Résultats par zone", body: "Europe du Nord : +25 %\nAsie : +12 %\nAmérique : stable" },
            { layout: "content", title: "Nos 3 priorités", body: "Recruter 2 export managers\nSalon ProWein en mars\nCatalogue en 4 langues" }
          ]
        },
        tasks: [
          {
            text: "Rédige des <b>notes du présentateur</b> sur <b>chaque</b> diapositive (au moins une phrase de 15 caractères).",
            hint: "Sélectionne chaque diapo et tape dans la zone Notes.",
            check: a => a.count() > 0 && S(a).every(s => String(s.notes || "").trim().length >= 15)
          },
          {
            text: "Dans les notes de « Résultats par zone », prévois une consigne de régie contenant le mot <b>pause</b>.",
            hint: "Ex. « PAUSE 3 s après le chiffre Europe du Nord, puis demander qui travaille cette zone. »",
            check: a => { const i = idxTitle(a, "resultats par zone"); return i >= 0 && N(S(a)[i].notes).includes("pause"); }
          },
          {
            text: "Ajoute une <b>dernière</b> diapo pour la séance de questions, dont le texte contient <b>questions</b>.",
            hint: "Disposition Titre de section ou Titre, ex. « Vos questions — et vos premiers contacts export ? ».",
            check: a => a.count() >= 5 && N(a.allText(a.count() - 1)).includes("question")
          },
          {
            text: "Répète : lance le <b>diaporama</b> au moins une fois.",
            hint: "Bouton Diaporama, puis navigue avec Espace / P, Échap pour sortir.",
            check: a => a.played === true
          }
        ],
        solution: `<ol>
<li>Notes diapo 1 : « Accroche : il y a 3 ans, l'export pesait 20 % de notre CA. » ; diapo 2 : « Citer les 3 pays : Suède, Japon, Canada. » ; etc.</li>
<li>Notes « Résultats par zone » : « PAUSE après Europe du Nord, demander qui couvre la zone. »</li>
<li>Nouvelle diapo finale « Vos questions » avec un rappel des 3 priorités.</li>
<li>Lancer le diaporama, utiliser B pour l'écran noir pendant les échanges.</li>
</ol>`
      },
      {
        id: "p06-lab2",
        title: "Lab 2 — Pilote de diaporama",
        kind: "keys",
        scenario: `<p>Le vidéoprojecteur de la salle est loin et la télécommande a rendu l'âme. Vous pilotez au clavier : démarrage, navigation, écran noir, annotation, sortie.</p>`,
        setup: {
          sequence: [
            { keys: "F5", desc: "Lancer le diaporama depuis le début" },
            { keys: "N", desc: "Diapo suivante" },
            { keys: "P", desc: "Diapo précédente" },
            { keys: "B", desc: "Écran noir pendant la discussion" },
            { keys: "W", desc: "Écran blanc" },
            { keys: "Ctrl+P", desc: "Prendre le stylo" },
            { keys: "E", desc: "Effacer les annotations" },
            { keys: "Échap", desc: "Quitter le diaporama" },
            { keys: "Alt+F5", desc: "Relancer en mode présentateur" },
            { keys: "Maj+F5", desc: "Reprendre depuis la diapo active" }
          ]
        },
        solution: `<ol><li>F5</li><li>N</li><li>P</li><li>B</li><li>W</li><li>Ctrl+P</li><li>E</li><li>Échap</li><li>Alt+F5</li><li>Maj+F5</li></ol>`
      }
    ],
    quiz: [
      { q: "Quel raccourci lance directement le mode présentateur ?", choices: ["F5", "Maj+F5", "Alt+F5", "Ctrl+F5"], answer: 2, explain: "Alt+F5 ouvre le mode présentateur, même avec un seul écran : idéal pour répéter." },
      { q: "En diaporama, quelle touche affiche un écran noir ?", choices: ["N", "B", "W", "Échap"], answer: 1, explain: "B (ou le point) bascule sur un écran noir ; W (ou la virgule) sur un écran blanc." },
      { q: "Comment aller directement à la diapo 14 pendant le diaporama ?", choices: ["Ctrl+14", "Taper 14 puis Entrée", "F14", "Alt+Flèche droite 14 fois"], answer: 1, explain: "Saisir le numéro puis Entrée saute directement à la diapositive." },
      { q: "Après « Vérifier le minutage », les diapos défilent seules en réunion. Pourquoi ?", choices: ["Un bug", "Le minutage enregistré est utilisé ; il faut décocher Utiliser le minutage", "Le thème est animé", "Le mode Lecture est actif"], answer: 1, explain: "Les durées conservées font avancer les diapos. Décochez Diaporama › Utiliser le minutage pour garder la main." },
      { q: "Quel risque présentent les notes du présentateur quand on envoie le .pptx ?", choices: ["Aucun, elles sont supprimées", "Elles sont lisibles par le destinataire", "Elles alourdissent le fichier de 50 %", "Elles cassent les animations"], answer: 1, explain: "Les notes restent dans le fichier : relisez-les ou inspectez le document avant envoi, ou envoyez un PDF sans notes." }
    ]
  },

  /* ======================================================================
     p07 — Niveau 4 : Design avancé
     ====================================================================== */
  {
    id: "p07",
    level: 4,
    title: "Design avancé : Morphose, Zoom, grilles et data-viz",
    icon: "💎",
    duration: 70,
    summary: "Les techniques des designers de présentation : mouvement fluide avec Morphose, navigation non linéaire, mise en page sur grille et graphiques qui parlent.",
    objectives: [
      "Être capable de créer une animation fluide avec la transition Morphose",
      "Être capable de construire une navigation par Zoom de résumé, de section ou de diapositive",
      "Être capable de structurer une diapo sur une grille et une hiérarchie visuelle claire",
      "Être capable de transformer un graphique par défaut en graphique qui délivre un message"
    ],
    lessons: [
      {
        title: "La transition Morphose",
        html: `<p><b>Morphose</b> anime automatiquement les différences entre deux diapos consécutives : un objet présent sur les deux qui change de position, de taille, de couleur ou de rotation glisse de l'un à l'autre.</p>
<div class="o-path">Transitions › Accès à cette diapositive › Morphose</div>
<ol>
<li>Construisez la diapo A.</li>
<li><b>Dupliquez-la</b> (<kbd>Ctrl</kbd>+<kbd>D</kbd> sur la miniature) : c'est ce qui garantit que les objets sont « les mêmes ».</li>
<li>Sur la copie B, déplacez/agrandissez/recolorez.</li>
<li>Appliquez <b>Morphose</b> à la diapo <b>B</b> (la transition s'applique à l'arrivée).</li>
</ol>
<p><i>Options d'effet</i> : Objets (défaut), Mots, Caractères — ce dernier anime les lettres d'un titre qui se réorganisent.</p>
<div class="callout warn">Piège : si Morphose fait un fondu au lieu d'un déplacement, PowerPoint n'a pas apparié les objets. Forcez l'appariement en nommant les deux objets à l'identique avec le préfixe <code>!!</code> dans le volet Sélection (<code>!!Cercle</code>).</div>`
      },
      {
        title: "Zoom : la navigation non linéaire",
        html: `<div class="o-path">Insertion › Liens › Zoom › Zoom de résumé</div>
<ul>
<li><b>Zoom de résumé</b> : crée une diapo sommaire cliquable avec une vignette par section ; en diaporama, on zoome dans la section puis on revient au résumé.</li>
<li><b>Zoom de section</b> : vignette vers une section existante, à placer où vous voulez.</li>
<li><b>Zoom de diapositive</b> : lien vers une diapo précise (annexe, démo, chiffre détaillé).</li>
</ul>
<p>Dans <i>Format du zoom</i> : <b>Revenir au zoom</b> (retour automatique au sommaire), <b>Transition de zoom</b>, image de la vignette personnalisée.</p>
<div class="callout ok">Cas d'usage idéal : une présentation « à la carte » devant un client — vous choisissez la section selon ses questions, sans faire défiler 30 diapos.</div>
<div class="o-tip">💡 Mettez vos annexes dans une section masquée accessible par Zoom de diapositive : elles ne gênent pas le fil, mais sont à un clic si on vous pose la question.</div>`
      },
      {
        title: "Grilles et hiérarchie visuelle",
        html: `<p>Les designers placent tout sur une <b>grille</b> : marges identiques (ex. 1,27 cm), colonnes régulières, alignements stricts. Dans PowerPoint :</p>
<ul>
<li><b>Repères</b> personnalisés (<kbd>Alt</kbd>+<kbd>F9</kbd>, clic droit sur un repère › <i>Ajouter un repère vertical</i>) ; placés sur le masque, ils sont verrouillés pour toutes les diapos.</li>
<li><b>Grille</b> (<kbd>Maj</kbd>+<kbd>F9</kbd>) et <i>Aligner les objets sur la grille</i>.</li>
<li>Taille et position exactes dans <i>Format de forme › Taille</i>.</li>
</ul>
<p><b>Hiérarchie</b> : le regard doit savoir où aller en premier. Leviers : <b>taille</b> (le chiffre clé 3× plus gros que le reste), <b>contraste</b> (couleur d'accent unique), <b>position</b> (en haut à gauche lu en premier), <b>espace blanc</b>.</p>
<div class="callout bad">Erreur fréquente : tout mettre en gras, en couleur et en grand. Quand tout est important, rien ne l'est.</div>`
      },
      {
        title: "Data-viz : des graphiques qui délivrent un message",
        html: `<p>Un graphique par défaut montre des données ; un bon graphique montre <b>une conclusion</b>.</p>
<ol>
<li><b>Titre-message</b> : « Le e-commerce double en 2 ans », pas « CA par canal ».</li>
<li><b>Le bon type</b> : barres pour comparer, courbes pour une évolution, secteurs seulement pour 2-3 parts d'un tout.</li>
<li><b>Désencombrer</b> : supprimez quadrillage, légende inutile, axe redondant (étiquettes de données à la place). <i>Création de graphique › Ajouter un élément de graphique</i>.</li>
<li><b>Une couleur d'accent</b> sur la série qui porte le message, le reste en gris.</li>
</ol>
<div class="o-path">Insertion › Illustrations › Graphique</div>
<div class="callout">Un graphique copié depuis Excel avec <i>Collage spécial › Coller avec liaison</i> se met à jour quand le classeur change (<i>Création de graphique › Actualiser les données</i>). Attention aux chemins si vous déplacez les fichiers.</div>
<div class="o-tip">💡 Animez un histogramme <i>Par série</i> ou <i>Par catégorie</i> (Options d'effet) pour révéler la comparaison au moment où vous en parlez.</div>`
      }
    ],
    shortcuts: [
      { keys: "Ctrl+D", desc: "Dupliquer (diapo ou objet) — base de la Morphose" },
      { keys: "Alt+F10", desc: "Volet Sélection (renommer pour Morphose)" },
      { keys: "Maj+F9", desc: "Grille" },
      { keys: "Alt+F9", desc: "Repères" },
      { keys: "Ctrl+Maj+C", desc: "Copier la mise en forme" },
      { keys: "Ctrl+Maj+V", desc: "Coller la mise en forme" },
      { keys: "Maj+F5", desc: "Tester la transition sur la diapo active" }
    ],
    labs: [
      {
        id: "p07-lab1",
        title: "Lab 1 — Le chiffre qui grandit",
        kind: "slides",
        scenario: `<p>La start-up <b>GreenBox</b> lève des fonds. Pour la diapo de traction, l'investisseur doit <b>voir</b> la croissance : une forme qui s'étire d'une diapo à l'autre grâce à Morphose, et un chiffre clé qui domine visuellement la diapo.</p>`,
        setup: {
          theme: "neon",
          slides: [
            { layout: "title", title: "GreenBox", subtitle: "Levée de fonds — série A" },
            { layout: "blank" },
            { layout: "content", title: "Une croissance de +42 % par trimestre", body: "Clients actifs : 3 200\nPanier moyen : 38 €\nChurn : 2,1 %" }
          ]
        },
        tasks: [
          {
            text: "Sur la diapo vide (2), ajoute une <b>forme</b> (la « barre » des ventes 2024), puis <b>duplique</b> la diapo.",
            hint: "Insère un rectangle, puis bouton Dupliquer du volet miniatures.",
            check: a => S(a).filter(s => byType(s, "shape").length >= 1).length >= 2
          },
          {
            text: "Sur la copie, <b>agrandis ou déplace</b> nettement la forme et applique la transition <b>Morphose</b> à cette copie.",
            hint: "La transition se met sur la diapo d'arrivée (la seconde).",
            check: a => {
              const sl = S(a);
              for (let i = 1; i < sl.length; i++) {
                if (!N(sl[i].transition).includes("morph")) continue;
                const A = byType(sl[i - 1], "shape"), B = byType(sl[i], "shape");
                if (A.some(p => B.some(q => Math.abs(num(p.x) - num(q.x)) > 5 || Math.abs(num(p.y) - num(q.y)) > 5 || Math.abs(num(p.w) - num(q.w)) > 5 || Math.abs(num(p.h) - num(q.h)) > 5))) return true;
              }
              return false;
            }
          },
          {
            text: "Sur la diapo « +42 % », ajoute une zone de texte <b>+42 %</b> et rends-la <b>plus grande</b> (taille de police) que tous les autres éléments de la diapo.",
            hint: "Augmente fortement la taille du texte +42 % (le titre contient aussi 42 : c'est la zone de texte ou la forme qui compte).",
            check: a => {
              const i = idxTitle(a, "42"); if (i < 0) return false;
              const els = E(S(a)[i]);
              const k = els.find(e => e.type !== "title" && N(e.text).replace(/\s/g, "").includes("+42"));
              return !!k && num(k.size) > 0 && els.every(e => e === k || num(e.size) < num(k.size));
            }
          },
          {
            text: "Élargis la zone du chiffre clé à au moins <b>60 %</b> de la largeur de la diapo, <b>centre-la horizontalement</b> sur la diapositive et vérifie l'effet en <b>diaporama</b>.",
            hint: "Tire la poignée droite de la zone de texte, puis Accueil › Aligner sur la diapositive › Centrer horizontalement, puis Diaporama.",
            check: a => { const i = idxTitle(a, "42"); if (i < 0 || !a.played) return false; return E(S(a)[i]).some(e => e.type !== "title" && N(e.text).replace(/\s/g, "").includes("+42") && num(e.w) >= 60 && a.centered(e) === true); }
          }
        ],
        solution: `<ol>
<li>Diapo 2 : insérer un rectangle étroit en bas à gauche (barre 2024). Dupliquer la diapo.</li>
<li>Sur la copie (diapo 3) : étirer le rectangle en hauteur (barre 2026), éventuellement changer sa couleur. <i>Transitions › Morphose</i> sur cette diapo.</li>
<li>Diapo « +42 % » : zone de texte « +42 % », taille 96 (plus grand que le titre), couleur d'accent.</li>
<li>Élargir la zone (≥ 60 % de la largeur), Centrer horizontalement sur la diapo, puis lancer le diaporama : la barre « pousse » entre les deux diapos.</li>
</ol>`
      },
      {
        id: "p07-lab2",
        title: "Lab 2 — Sommaire Zoom et graphique lié",
        kind: "steps",
        scenario: `<p>Dans le <b>vrai PowerPoint</b>, transformez une présentation commerciale en deck « à la carte » avec un Zoom de résumé, et ajoutez un graphique Excel lié qui délivre un message.</p>`,
        setup: {
          app: "PowerPoint",
          file: { name: "ventes-canal.csv", type: "csv", content: "Année;Magasin;E-commerce;Grossistes\n2024;410000;95000;120000\n2025;405000;150000;118000\n2026;398000;205000;121000" }
        },
        tasks: [
          { text: "Créez 3 sections (« Offre », « Références », « Tarifs ») contenant chacune 2 diapos.", hint: "Accueil › Section › Ajouter une section." },
          { text: "<i>Insertion › Zoom › Zoom de résumé</i> : cochez la 1re diapo de chaque section. Vérifiez que l'option <b>Revenir au zoom</b> est active.", hint: "Onglet Format du zoom, groupe Options du zoom." },
          { text: "Ouvrez le CSV dans Excel, créez un histogramme groupé, copiez-le et collez-le dans PowerPoint avec <b>Utiliser le thème de destination et lier les données</b>.", hint: "Icône Options de collage après Ctrl+V." },
          { text: "Désencombrez : supprimez le quadrillage, mettez <b>E-commerce</b> en couleur d'accent et les autres séries en gris, ajoutez un titre-message (« Le e-commerce double en 2 ans »).", hint: "Création de graphique › Ajouter un élément de graphique ; clic sur une série › Remplissage." },
          { text: "Modifiez une valeur dans Excel puis <i>Création de graphique › Actualiser les données</i> dans PowerPoint. Testez la navigation par zoom en diaporama.", hint: "Le classeur source doit rester au même emplacement." }
        ],
        solution: `<ol>
<li>Sections créées, Zoom de résumé inséré dans une nouvelle section « Résumé » en tête.</li>
<li>Graphique lié : la modification dans Excel apparaît après Actualiser les données.</li>
<li>Message visuel : une seule série en couleur, titre-message, étiquettes de données plutôt qu'axe vertical.</li>
</ol>`
      }
    ],
    quiz: [
      { q: "Sur quelle diapo applique-t-on la transition Morphose ?", choices: ["La diapo de départ", "La diapo d'arrivée (la seconde)", "Les deux", "La diapo de titre"], answer: 1, explain: "Une transition décrit l'arrivée sur une diapo : Morphose se met sur la seconde diapo de la paire." },
      { q: "Pourquoi dupliquer la diapo avant de modifier pour Morphose ?", choices: ["Pour gagner du temps uniquement", "Pour que PowerPoint reconnaisse les mêmes objets sur les deux diapos", "C'est imposé par le masque", "Pour garder une sauvegarde"], answer: 1, explain: "Les objets dupliqués sont appariés automatiquement ; sinon on peut forcer avec le préfixe !! dans le volet Sélection." },
      { q: "Quel outil crée un sommaire cliquable avec retour automatique ?", choices: ["Lien hypertexte", "Bouton d'action", "Zoom de résumé", "SmartArt"], answer: 2, explain: "Insertion › Zoom › Zoom de résumé crée la diapo sommaire et les retours (option Revenir au zoom)." },
      { q: "Quel type de graphique pour montrer l'évolution du CA sur 12 mois ?", choices: ["Secteurs", "Courbes", "Anneau", "Radar"], answer: 1, explain: "Les courbes montrent une tendance dans le temps ; les secteurs sont réservés à quelques parts d'un tout." },
      { q: "Quel principe améliore le plus la lisibilité d'un graphique dans une présentation ?", choices: ["Ajouter un effet 3D", "Une couleur d'accent sur la série qui porte le message, le reste en gris", "Afficher toutes les légendes et quadrillages", "Utiliser une couleur par barre"], answer: 1, explain: "Le contraste guide le regard vers la conclusion ; le reste devient contexte." }
    ]
  },

  /* ======================================================================
     p08 — Niveau 4 : Accessibilité, export, collaboration, IA
     ====================================================================== */
  {
    id: "p08",
    level: 4,
    title: "Accessibilité, export, co-édition, Designer et Copilot",
    icon: "🚀",
    duration: 70,
    summary: "Livrer une présentation accessible à tous, la diffuser dans le bon format, la construire à plusieurs et accélérer avec l'IA.",
    objectives: [
      "Être capable de rendre un deck accessible (titres, textes de remplacement, ordre de lecture, contraste, sous-titres)",
      "Être capable d'exporter en PDF, en vidéo et d'inspecter le document avant diffusion",
      "Être capable de co-éditer avec commentaires, @mentions et historique des versions",
      "Être capable d'utiliser Designer et Copilot de façon critique"
    ],
    lessons: [
      {
        title: "Accessibilité : les règles qui changent tout",
        html: `<div class="o-path">Révision › Accessibilité › Vérifier l'accessibilité</div>
<p>Le vérificateur classe les problèmes en <b>Erreurs</b>, <b>Avertissements</b> et <b>Conseils</b>. Les plus fréquents :</p>
<ul>
<li><b>Titre de diapositive manquant</b> : chaque diapo doit avoir un titre unique (même masqué hors de la diapo si besoin).</li>
<li><b>Texte de remplacement manquant</b> : clic droit sur l'image › <i>Afficher le texte de remplacement</i>. Décrivez l'information, pas l'image (« Graphique : le e-commerce passe de 95 à 205 k€ »). Image décorative : cochez <i>Marquer comme décoratif</i>.</li>
<li><b>Ordre de lecture</b> : <i>Révision › Vérifier l'accessibilité › Volet Ordre de lecture</i>.</li>
<li><b>Contraste</b> insuffisant entre texte et fond.</li>
</ul>
<div class="callout">En diaporama, les <b>sous-titres en direct</b> (<i>Diaporama › Paramètres des sous-titres</i>) transcrivent votre voix à l'écran, voire la traduisent.</div>
<div class="callout warn">Piège : un titre placé dans une simple zone de texte n'est <b>pas</b> un titre pour le lecteur d'écran. Utilisez l'espace réservé de titre de la disposition.</div>`
      },
      {
        title: "Exporter et diffuser",
        html: `<table class="o-table"><thead><tr><th>Besoin</th><th>Format / commande</th></tr></thead><tbody>
<tr><td>Envoyer un support à lire</td><td><i>Fichier › Exporter › Créer un document PDF/XPS</i> (Options : diapos, documents, pages de notes)</td></tr>
<tr><td>Diffuser sans présentateur</td><td><i>Fichier › Exporter › Créer une vidéo</i> (MP4, Full HD ou 4K, minutage et narrations)</td></tr>
<tr><td>Imprimer un support</td><td><i>Fichier › Exporter › Créer des documents</i> (vers Word)</td></tr>
<tr><td>Ouverture directe en diaporama</td><td><i>Enregistrer sous › Diaporama PowerPoint (*.ppsx)</i></td></tr>
</tbody></table>
<p>Pour une vidéo narrée : <i>Enregistrer › Enregistrer › Depuis le début</i> (caméra, micro, annotations), puis export.</p>
<div class="o-path">Fichier › Informations › Vérifier la présence de problèmes › Inspecter le document</div>
<div class="callout warn">Piège : avant d'envoyer le .pptx à l'extérieur, l'<b>Inspecteur de document</b> retire commentaires, notes, métadonnées et contenu hors diapo. Faites-le sur une <b>copie</b> : la suppression est définitive.</div>
<div class="o-tip">💡 <i>Fichier › Informations › Compresser les médias</i> et <i>Format de l'image › Compresser les images</i> divisent souvent la taille d'un deck par 3 avant envoi par e-mail.</div>`
      },
      {
        title: "Co-édition dans Microsoft 365",
        html: `<p>Enregistrez le fichier sur <b>OneDrive</b> ou <b>SharePoint</b> pour travailler à plusieurs en même temps : les avatars des co-auteurs apparaissent sur les diapos qu'ils modifient.</p>
<div class="o-path">Partager › Partager › Copier le lien</div>
<ul>
<li><b>Commentaires</b> : <i>Révision › Nouveau commentaire</i> ; tapez <code>@</code> suivi d'un nom pour mentionner et notifier un collègue, ou lui <b>attribuer</b> une tâche.</li>
<li><b>Historique des versions</b> : <i>Fichier › Informations › Historique des versions</i> pour comparer ou restaurer.</li>
<li>Répartissez le travail par <b>sections</b> (une section par contributeur) et utilisez la <b>diapo masquée</b> « Travail en cours » pour les brouillons.</li>
</ul>
<div class="callout bad">Erreur fréquente : s'envoyer des versions « deck_v3_final_VRAIMENT-final.pptx » par e-mail. Un seul fichier partagé, l'historique fait le reste.</div>`
      },
      {
        title: "Designer et Copilot : l'IA comme assistant",
        html: `<p><b>Designer</b> (<i>Accueil › Designer</i>) propose des mises en page à partir du contenu de la diapo : photo, liste, chronologie, chiffres. Il fonctionne mieux sur des diapos simples (titre + 1 image ou 3-6 puces) et respecte le thème.</p>
<p><b>Copilot</b> (<i>Accueil › Copilot</i>, licence Microsoft 365 Copilot) peut :</p>
<ul>
<li>créer une présentation à partir d'une invite ou d'un document Word/PDF (« Crée une présentation de 8 diapos à partir de /rapport-annuel.docx ») ;</li>
<li>ajouter une diapo, résumer le deck, organiser en sections, proposer des notes du présentateur ;</li>
<li>utiliser le <b>modèle de votre entreprise</b> si vous partez d'un fichier basé sur ce modèle.</li>
</ul>
<div class="callout warn">Piège : Copilot produit un <b>premier jet</b>. Vérifiez chaque chiffre, retravaillez les titres en titres-messages, supprimez les diapos génériques et relancez le vérificateur d'accessibilité (textes de remplacement générés à contrôler).</div>
<div class="callout ok">Une bonne invite précise : public, objectif, nombre de diapos, ton, source. « Pour le comité de direction, 6 diapos, ton factuel, à partir de /bilan-T3.docx, terminer par une demande de décision. »</div>`
      }
    ],
    shortcuts: [
      { keys: "Alt+F10", desc: "Volet Sélection (noms et ordre des objets)" },
      { keys: "F7", desc: "Vérifier l'orthographe" },
      { keys: "Ctrl+K", desc: "Insérer un lien hypertexte" },
      { keys: "Ctrl+Alt+M", desc: "Nouveau commentaire" },
      { keys: "Ctrl+S", desc: "Enregistrer (enregistrement automatique sur OneDrive)" },
      { keys: "Alt+F5", desc: "Mode présentateur (sous-titres, outils)" }
    ],
    labs: [
      {
        id: "p08-lab1",
        title: "Lab 1 — Deck accessible et prêt à diffuser",
        kind: "steps",
        scenario: `<p>La mairie de <b>Saint-Aubin</b> publie sur son site la présentation du budget participatif. Elle doit être <b>accessible</b> (obligation légale RGAA) et diffusée en PDF et en vidéo. Travaillez dans le <b>vrai PowerPoint</b> sur une présentation de 5-6 diapos (la vôtre ou une générée par Designer).</p>`,
        setup: { app: "PowerPoint" },
        tasks: [
          { text: "Lancez <i>Révision › Vérifier l'accessibilité</i> et corrigez toutes les <b>erreurs</b> (titres manquants, textes de remplacement).", hint: "Cliquez sur chaque problème : le volet propose une action recommandée." },
          { text: "Rédigez un texte de remplacement <b>informatif</b> pour chaque image ou graphique, et marquez les éléments purement décoratifs comme tels.", hint: "Clic droit › Afficher le texte de remplacement." },
          { text: "Ouvrez le <b>volet Ordre de lecture</b> et remettez sur chaque diapo l'ordre logique : titre, puis contenu.", hint: "Révision › Vérifier l'accessibilité › Volet Ordre de lecture." },
          { text: "Sur une <b>copie</b> du fichier, lancez l'<b>Inspecteur de document</b> et supprimez commentaires, notes et métadonnées personnelles.", hint: "Fichier › Informations › Vérifier la présence de problèmes › Inspecter le document." },
          { text: "<i>Fichier › Exporter › Créer un document PDF/XPS</i> en conservant les <b>balises de structure du document pour l'accessibilité</b> (bouton Options).", hint: "L'option est cochée par défaut, vérifiez-la." },
          { text: "<i>Fichier › Exporter › Créer une vidéo</i> en Full HD (1080p), 5 secondes par diapo.", hint: "Sans minutage enregistré, PowerPoint utilise la durée par défaut." }
        ],
        solution: `<ol>
<li>Vérificateur : 0 erreur ; chaque diapo a un titre unique dans l'espace réservé de titre.</li>
<li>Textes de remplacement décrivant l'information ; décoratifs marqués.</li>
<li>Ordre de lecture : titre en premier sur chaque diapo.</li>
<li>Copie inspectée, puis PDF balisé et MP4 1080p exportés.</li>
</ol>`
      },
      {
        id: "p08-lab2",
        title: "Lab 2 — Co-écrire avec l'équipe et l'IA",
        kind: "steps",
        scenario: `<p>Votre équipe marketing prépare la <b>présentation de lancement</b> d'un nouveau produit pour le comité de direction. Vous travaillez à trois sur le même fichier OneDrive, avec l'aide de Designer et (si vous avez la licence) de Copilot.</p>`,
        setup: {
          app: "PowerPoint",
          file: { name: "brief-lancement.txt", type: "txt", content: "Produit : GreenBox Mini, boîte de compost d'appartement.\nCible : urbains 25-40 ans.\nPrix : 49 €.\nLancement : septembre 2026, 120 magasins partenaires + e-commerce.\nObjectif : 15 000 unités la première année.\nBudget marketing demandé : 180 k€.\nDécision attendue du comité : valider le budget et la date." }
        },
        tasks: [
          { text: "Enregistrez la présentation sur OneDrive, <i>Partager › Copier le lien</i> avec droits de modification, et créez une <b>section par contributeur</b>.", hint: "Accueil › Section › Ajouter une section." },
          { text: "Ajoutez un commentaire avec une <b>@mention</b> qui attribue une tâche à un collègue (« @Lina peux-tu vérifier le prix ? »).", hint: "Révision › Nouveau commentaire, puis tapez @." },
          { text: "Avec <b>Copilot</b> (ou à la main si vous n'avez pas la licence) : générez 6 diapos à partir du brief, avec une invite précisant public, ton, nombre de diapos et la demande de décision finale.", hint: "Accueil › Copilot › Créer une présentation à partir d'un fichier." },
          { text: "Relisez le résultat : corrigez au moins un chiffre ou un titre générique et transformez les titres en <b>titres-messages</b>.", hint: "Comparez chaque chiffre avec le brief." },
          { text: "Utilisez <b>Designer</b> sur au moins 2 diapos et gardez une proposition cohérente avec le thème.", hint: "Accueil › Designer : sélectionnez une diapo simple." },
          { text: "Ouvrez l'<b>historique des versions</b> et repérez la version d'avant la génération IA.", hint: "Fichier › Informations › Historique des versions." }
        ],
        solution: `<ol>
<li>Fichier partagé sur OneDrive, une section par personne, commentaires avec @mentions et tâches attribuées.</li>
<li>Invite type : « Pour le comité de direction, crée 6 diapos au ton factuel à partir de brief-lancement.txt ; termine par la décision attendue : valider 180 k€ et un lancement en septembre 2026. »</li>
<li>Relecture humaine : chiffres vérifiés (49 €, 15 000 unités, 180 k€), titres-messages, Designer sur les diapos visuelles.</li>
<li>Historique des versions : restauration possible à tout moment.</li>
</ol>`
      }
    ],
    quiz: [
      { q: "Quelle erreur d'accessibilité est la plus fréquente dans PowerPoint ?", choices: ["Police trop grande", "Titre de diapositive manquant ou texte de remplacement absent", "Trop de diapositives", "Thème sombre"], answer: 1, explain: "Le lecteur d'écran s'appuie sur les titres et les textes de remplacement : ce sont les premières erreurs signalées." },
      { q: "Que faire d'une image purement décorative ?", choices: ["La supprimer", "La marquer comme décorative", "Lui donner le texte « image »", "La mettre en arrière-plan"], answer: 1, explain: "Marquer comme décoratif indique au lecteur d'écran de l'ignorer, sans bruit inutile." },
      { q: "Pourquoi lancer l'Inspecteur de document sur une copie ?", choices: ["Il ne fonctionne pas sur l'original", "Les suppressions (notes, commentaires) sont définitives", "Il change le thème", "Il compresse les images"], answer: 1, explain: "Les éléments retirés ne peuvent pas être récupérés par Annuler : gardez l'original intact." },
      { q: "Comment notifier un collègue dans un commentaire PowerPoint 365 ?", choices: ["L'écrire en majuscules", "Taper @ suivi de son nom", "Le mettre en gras", "Envoyer le fichier par e-mail"], answer: 1, explain: "La @mention envoie une notification et peut attribuer le commentaire comme tâche." },
      { q: "Quelle est la bonne posture face à une présentation générée par Copilot ?", choices: ["La projeter telle quelle", "La considérer comme un premier jet à vérifier et retravailler", "La supprimer systématiquement", "La convertir en PDF immédiatement"], answer: 1, explain: "L'IA accélère le premier jet ; la vérification des faits, les titres-messages et l'accessibilité restent votre responsabilité." }
    ]
  }

  ]);
})();
