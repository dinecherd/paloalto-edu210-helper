/* Office Academy — Excel : 8 modules du débutant à l'expert */
(function () {
  "use strict";

  /* Petits utilitaires pour les check (ne lèvent jamais d'exception) */
  const strip = (s) => String(s || "").replace(/\$/g, "");
  const rangeStarts = (r, ref) => strip(r).toUpperCase().indexOf(ref) === 0;
  // texte attendu, sans tenir compte de la casse ni des espaces autour (comme les comparaisons d'Excel)
  const same = (v, s) => typeof v === "string" && v.trim().toLowerCase() === s.toLowerCase();
  // la plage r (ex. "D1:D13", "$D$2:$D$13") est dans la colonne col et couvre les lignes r1..r2
  const covers = (r, col, r1, r2) => {
    const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(strip(r).toUpperCase().trim());
    return !!m && m[1] === col && (m[3] || m[1]) === col && +m[2] <= r1 && +(m[4] || m[2]) >= r2;
  };
  // somme / max / comptage conditionnels recalculés depuis la grille (pour le tableau de bord x07)
  const agg = (a, region, gamme) => {
    let sum = 0, sumG = 0, n3000 = 0, max = 0;
    for (let i = 2; i <= 13; i++) {
      if (!same(a.val("B" + i), String(region))) continue;
      const ca = a.num("D" + i) * a.num("E" + i);
      sum += ca;
      if (same(a.val("C" + i), gamme)) sumG += ca;
      if (ca >= 3000) n3000++;
      if (ca > max) max = ca;
    }
    return { sum, sumG, n3000, max };
  };

  OFFICE.add("excel", [
    /* ================================================================== */
    /* NIVEAU 1 — x01                                                       */
    /* ================================================================== */
    {
      id: "x01",
      level: 1,
      title: "Premiers pas : l'interface, la saisie et les premières formules",
      icon: "🧭",
      duration: 50,
      summary: "Apprivoiser le classeur, saisir proprement des données et écrire ses premiers calculs avec SOMME et MOYENNE.",
      objectives: [
        "Être capable de se repérer dans l'interface d'Excel (ruban, zone Nom, barre de formule, onglets de feuilles)",
        "Être capable de saisir et corriger des textes, des nombres et des dates",
        "Être capable d'appliquer un format de nombre (monétaire, pourcentage, date)",
        "Être capable d'écrire une formule avec + - * / et les fonctions SOMME et MOYENNE"
      ],
      lessons: [
        {
          title: "L'interface d'Excel 365",
          html: `<p>Un fichier Excel est un <b>classeur</b> composé de <b>feuilles</b> (onglets en bas de l'écran). Chaque feuille est une grille de <b>colonnes</b> (A, B, C… jusqu'à XFD) et de <b>lignes</b> (1 à 1 048 576). L'intersection d'une colonne et d'une ligne est une <b>cellule</b>, identifiée par sa référence : <code>B3</code> = colonne B, ligne 3.</p>
<table class="o-table"><thead><tr><th>Zone</th><th>Rôle</th></tr></thead><tbody>
<tr><td><b>Ruban</b></td><td>Onglets Accueil, Insertion, Mise en page, Formules, Données, Révision, Affichage : toutes les commandes</td></tr>
<tr><td><b>Zone Nom</b></td><td>À gauche de la barre de formule : affiche la cellule active, permet d'y taper <code>D25</code> pour s'y rendre</td></tr>
<tr><td><b>Barre de formule</b></td><td>Montre le <i>contenu réel</i> de la cellule (la formule), alors que la cellule montre le <i>résultat</i></td></tr>
<tr><td><b>Barre d'état</b></td><td>En bas : Moyenne, Nombre et Somme de la sélection, sans aucune formule</td></tr>
</tbody></table>
<div class="o-path">Fichier › Options › Personnaliser le ruban</div>
<div class="o-tip">💡 Sélectionnez quelques nombres et regardez la barre d'état : Excel affiche déjà leur somme. Clic droit sur la barre d'état pour ajouter Min et Max.</div>`
        },
        {
          title: "Saisir et corriger des données",
          html: `<p>Cliquez sur une cellule et tapez : la saisie se valide avec <kbd>Entrée</kbd> (descend), <kbd>Tab</kbd> (va à droite) ou une flèche. <kbd>Échap</kbd> annule la saisie en cours.</p>
<ul>
<li><b>Texte</b> : aligné à gauche par défaut.</li>
<li><b>Nombre</b> : aligné à droite. En français, la virgule est le séparateur décimal : tapez <code>12,5</code>.</li>
<li><b>Date</b> : tapez <code>15/03/2026</code>. Excel la stocke comme un <i>numéro de série</i> (nombre de jours depuis le 01/01/1900), ce qui permet de calculer des écarts de dates.</li>
</ul>
<p>Pour corriger une cellule sans tout retaper : <kbd>F2</kbd> (ou double-clic) passe en mode Modification. <kbd>Alt</kbd>+<kbd>Entrée</kbd> insère un retour à la ligne <i>dans</i> la cellule.</p>
<div class="callout warn"><b>Piège :</b> un nombre aligné à gauche est souvent un <b>texte</b> (import, espace parasite, point au lieu de virgule). Il sera ignoré par SOMME. Un petit triangle vert dans l'angle le signale.</div>
<div class="callout ok"><b>Bonne pratique :</b> une information par cellule, une ligne d'en-têtes, pas de lignes vides au milieu des données. Tous les outils (tri, filtre, TCD) en dépendent.</div>`
        },
        {
          title: "Formats de nombre",
          html: `<p>Le format change l'<b>affichage</b>, jamais la valeur stockée. <code>0,2</code> affiché en pourcentage donne <code>20 %</code>, mais vaut toujours 0,2 dans les calculs.</p>
<div class="o-path">Accueil › Nombre › Format de nombre</div>
<table class="o-table"><thead><tr><th>Format</th><th>Exemple</th><th>Usage</th></tr></thead><tbody>
<tr><td>Standard</td><td>1234,5</td><td>par défaut</td></tr>
<tr><td>Nombre (2 déc.)</td><td>1 234,50</td><td>quantités, mesures</td></tr>
<tr><td>Monétaire / Comptabilité</td><td>1 234,50 €</td><td>montants</td></tr>
<tr><td>Pourcentage</td><td>12,50 %</td><td>taux</td></tr>
<tr><td>Date courte</td><td>15/03/2026</td><td>dates</td></tr>
</tbody></table>
<p>Pour tous les réglages : <kbd>Ctrl</kbd>+<kbd>1</kbd> ouvre la boîte <b>Format de cellule</b> (onglets Nombre, Alignement, Police, Bordure, Remplissage, Protection).</p>
<div class="callout warn"><b>Piège :</b> si vous voyez <code>########</code>, ce n'est pas une erreur : la colonne est trop étroite. Double-cliquez sur la bordure droite de l'en-tête de colonne pour l'ajuster.</div>`
        },
        {
          title: "Vos premières formules",
          html: `<p>Une formule commence <b>toujours</b> par <code>=</code>. Elle peut combiner des nombres, des références de cellules et des opérateurs :</p>
<pre class="code">=B2*C2        multiplication
=D10-F2       soustraction
=(B2+B3)/2    parenthèses prioritaires
=B2^2         puissance</pre>
<p>Plutôt que de taper <code>B2</code>, <b>cliquez</b> sur la cellule pendant la saisie : Excel insère la référence. Si B2 change, le résultat se recalcule automatiquement — c'est tout l'intérêt de calculer avec des références plutôt qu'avec des valeurs.</p>
<h4>Les deux fonctions de base</h4>
<pre class="code">=SOMME(D2:D9)      additionne la plage D2 à D9
=MOYENNE(D2:D9)    moyenne arithmétique</pre>
<p>Le <code>:</code> signifie « jusqu'à ». Les arguments d'une fonction sont séparés par un <b>point-virgule</b> en version française : <code>=SOMME(B2:B5;D2:D5)</code>.</p>
<div class="o-path">Accueil › Édition › Somme automatique</div>
<div class="callout bad"><b>Erreur fréquente :</b> écrire <code>=SOMME(1200+1260+405)</code> en recopiant les valeurs. La formule ne se mettra jamais à jour : référencez toujours les cellules.</div>
<div class="o-tip">💡 <kbd>Alt</kbd>+<kbd>=</kbd> insère directement <code>=SOMME(…)</code> avec la plage au-dessus déjà proposée.</div>`
        }
      ],
      shortcuts: [
        { keys: "F2", desc: "Modifier la cellule active" },
        { keys: "Alt+Entrée", desc: "Retour à la ligne dans une cellule" },
        { keys: "Ctrl+Entrée", desc: "Valider la saisie dans toutes les cellules sélectionnées" },
        { keys: "Ctrl+;", desc: "Insérer la date du jour" },
        { keys: "Ctrl+1", desc: "Ouvrir Format de cellule" },
        { keys: "Ctrl+Flèche bas", desc: "Aller à la dernière cellule remplie de la colonne" },
        { keys: "Ctrl+Z", desc: "Annuler" },
        { keys: "Ctrl+Y", desc: "Rétablir" }
      ],
      labs: [
        {
          id: "x01-lab1",
          title: "Lab 1 — Le budget du séminaire commercial",
          kind: "sheet",
          scenario: `<p>Vous êtes assistant(e) de direction chez <b>Novalis</b>, PME de 40 salariés. Le séminaire commercial annuel approche et la directrice veut le budget prévisionnel ce soir. Les devis ont été saisis : il reste à calculer les totaux et à vérifier que l'enveloppe de <b>7 500 €</b> est respectée.</p>`,
          setup: {
            cols: 6, rows: 13, sheetName: "Séminaire",
            data: {
              A1: "Poste", B1: "Quantité", C1: "Prix unitaire", D1: "Total",
              A2: "Location salle", B2: 1, C2: 1200,
              A3: "Traiteur déjeuner", B3: 45, C3: 28,
              A4: "Pause café", B4: 90, C4: 4.5,
              A5: "Vidéoprojecteur", B5: 2, C5: 85,
              A6: "Hébergement (nuitées)", B6: 20, C6: 96,
              A7: "Navette autocar", B7: 3, C7: 350,
              A8: "Badges imprimés", B8: 50, C8: 1.2,
              A9: "Animateur", B9: 1, C9: 900,
              A10: "Total", A11: "Moyenne par poste", A12: "Écart au budget",
              E2: "Budget alloué", F2: 7500
            },
            fmt: { C2: "eur", C3: "eur", C4: "eur", C5: "eur", C6: "eur", C7: "eur", C8: "eur", C9: "eur", F2: "eur" },
            widths: { A: 170, E: 120 }
          },
          tasks: [
            {
              text: "En <b>D2</b>, calcule le total de la ligne (<code>=B2*C2</code>) puis recopie la formule jusqu'en <b>D9</b>.",
              hint: "Tape = puis clique sur B2, tape *, clique sur C2, Entrée. Sélectionne D2:D9 puis Ctrl+D.",
              check: a => a.isFormula("D2") && a.isFormula("D9") && a.near(a.num("D2"), 1200) && a.near(a.num("D6"), 1920) && a.near(a.num("D9"), 900)
            },
            {
              text: "En <b>D10</b>, calcule le coût total du séminaire avec la fonction <b>SOMME</b>.",
              hint: "=SOMME(D2:D9) — ou Alt+= depuis D10.",
              check: a => a.hasFn("D10", "SOMME") && a.near(a.num("D10"), 6965)
            },
            {
              text: "En <b>D11</b>, calcule le coût moyen par poste avec <b>MOYENNE</b>.",
              hint: "=MOYENNE(D2:D9) — attention à ne pas inclure le total D10 !",
              check: a => a.hasFn("D11", "MOYENNE") && a.near(a.num("D11"), 870.625)
            },
            {
              text: "En <b>D12</b>, calcule l'écart entre le budget alloué (<b>F2</b>) et le total (<b>D10</b>). Le résultat doit être positif si l'on reste sous l'enveloppe.",
              hint: "=F2-D10",
              check: a => a.isFormula("D12") && a.near(a.num("D12"), 535)
            },
            {
              text: "Mets les en-têtes <b>A1:D1</b> en <b>gras</b> et applique le format <b>euro</b> à la plage <b>D2:D12</b>.",
              hint: "Sélectionne A1:D1 → Gras. Sélectionne D2:D12 → Format de nombre › Monétaire (€).",
              check: a => a.style("A1").bold && a.style("D1").bold && a.style("D2").fmt === "eur" && a.style("D10").fmt === "eur" && a.style("D12").fmt === "eur"
            }
          ],
          solution: `<ol>
<li>D2 : <code>=B2*C2</code>, puis sélection D2:D9 et <kbd>Ctrl</kbd>+<kbd>D</kbd>. Résultats : 1 200 · 1 260 · 405 · 170 · 1 920 · 1 050 · 60 · 900.</li>
<li>D10 : <code>=SOMME(D2:D9)</code> → <b>6 965 €</b>.</li>
<li>D11 : <code>=MOYENNE(D2:D9)</code> → <b>870,63 €</b>.</li>
<li>D12 : <code>=F2-D10</code> → <b>535 €</b> de marge : le budget est respecté.</li>
<li>A1:D1 en gras ; D2:D12 au format Monétaire.</li>
</ol>`
        },
        {
          id: "x01-lab2",
          title: "Lab 2 — Défi clavier : saisir sans la souris",
          kind: "keys",
          scenario: `<p>Les utilisateurs rapides d'Excel touchent rarement la souris pendant la saisie. Enchaînez ces raccourcis essentiels le plus vite possible.</p>`,
          setup: {
            sequence: [
              { keys: "F2", desc: "Modifier la cellule active" },
              { keys: "Échap", desc: "Annuler la saisie en cours" },
              { keys: "Alt+Entrée", desc: "Retour à la ligne dans la cellule" },
              { keys: "Ctrl+Entrée", desc: "Valider dans toute la sélection" },
              { keys: "Ctrl+;", desc: "Insérer la date du jour" },
              { keys: "Ctrl+Flèche bas", desc: "Aller en bas des données" },
              { keys: "Alt+=", desc: "Insérer une somme automatique" },
              { keys: "Maj+Espace", desc: "Sélectionner la ligne entière" },
              { keys: "Ctrl+Z", desc: "Annuler la dernière action" },
              { keys: "Ctrl+Y", desc: "Rétablir" }
            ]
          },
          solution: `<p>À retenir : <kbd>F2</kbd> pour corriger, <kbd>Ctrl</kbd>+<kbd>Entrée</kbd> pour remplir une sélection d'un coup, <kbd>Ctrl</kbd>+flèches pour naviguer dans de grands tableaux, <kbd>Alt</kbd>+<kbd>=</kbd> pour la somme automatique, <kbd>Maj</kbd>+<kbd>Espace</kbd> pour sélectionner la ligne (<kbd>Ctrl</kbd>+<kbd>Espace</kbd> pour la colonne, souvent capté par le système sur Mac).</p>`
        }
      ],
      quiz: [
        {
          q: "Que montre la barre de formule lorsque la cellule D10 contient =SOMME(D2:D9) ?",
          choices: ["Le résultat 6965", "La formule =SOMME(D2:D9)", "Le format de la cellule", "Rien tant qu'on n'appuie pas sur F2"],
          answer: 1,
          explain: "La cellule affiche le résultat, la barre de formule affiche le contenu réel, c'est-à-dire la formule."
        },
        {
          q: "Quel est le séparateur d'arguments dans une fonction Excel en français ?",
          choices: ["La virgule ,", "Le deux-points :", "Le point-virgule ;", "L'espace"],
          answer: 2,
          explain: "En français, la virgule sert aux décimales ; les arguments sont donc séparés par « ; ». Le « : » définit une plage."
        },
        {
          q: "Une cellule affiche ########. Que faut-il faire ?",
          choices: ["Corriger la formule", "Élargir la colonne", "Supprimer le format", "Redémarrer Excel"],
          answer: 1,
          explain: "Les dièses indiquent que le nombre ou la date ne tient pas dans la largeur de colonne. Un double-clic sur la bordure de l'en-tête ajuste la largeur."
        },
        {
          q: "Vous tapez 0,2 puis appliquez le format Pourcentage. Quelle valeur est utilisée dans les calculs ?",
          choices: ["20", "0,2", "20 %, sous forme de texte", "2"],
          answer: 1,
          explain: "Le format ne change que l'affichage (20 %). La valeur stockée reste 0,2."
        },
        {
          q: "Pourquoi écrire =B2*C2 plutôt que =1*1200 ?",
          choices: ["C'est plus court", "Le résultat se met à jour si B2 ou C2 change", "Excel refuse les nombres dans les formules", "Pour que la cellule soit en gras"],
          answer: 1,
          explain: "Avec des références, la formule se recalcule automatiquement à chaque modification des données. Avec des constantes, elle reste figée."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 1 — x02                                                       */
    /* ================================================================== */
    {
      id: "x02",
      level: 1,
      title: "Recopie, références relatives et absolues, mise en forme",
      icon: "📐",
      duration: 55,
      summary: "Le secret des tableaux qui se calculent tout seuls : recopier une formule et figer ce qui doit l'être avec le $.",
      objectives: [
        "Être capable de recopier une formule avec la poignée de recopie, Ctrl+D et Ctrl+R",
        "Être capable de distinguer références relatives, absolues et mixtes",
        "Être capable de figer une référence avec F4",
        "Être capable de mettre en forme un tableau lisible (gras, remplissage, bordures, alignement)"
      ],
      lessons: [
        {
          title: "La recopie de formules",
          html: `<p>Une fois la première formule écrite, inutile de la retaper : on la <b>recopie</b>.</p>
<ul>
<li><b>Poignée de recopie</b> : le petit carré en bas à droite de la cellule active. Faites-la glisser, ou <b>double-cliquez</b> dessus pour recopier jusqu'à la fin des données adjacentes.</li>
<li><kbd>Ctrl</kbd>+<kbd>D</kbd> : recopie vers le bas la première cellule de la sélection.</li>
<li><kbd>Ctrl</kbd>+<kbd>R</kbd> : recopie vers la droite.</li>
</ul>
<div class="o-path">Accueil › Édition › Remplissage › En bas</div>
<p>La poignée sait aussi prolonger des séries : tapez <code>Janvier</code> puis tirez, Excel écrit Février, Mars… Même chose avec <code>Lun</code>, <code>T1</code> ou une date.</p>
<div class="o-tip">💡 Après une recopie, le bouton <b>Options de recopie</b> permet de choisir « Recopier les valeurs sans la mise en forme » pour ne pas écraser vos bordures.</div>`
        },
        {
          title: "Références relatives : le comportement par défaut",
          html: `<p>Par défaut, une référence est <b>relative</b> : Excel la mémorise comme un <i>déplacement</i> depuis la cellule de la formule. En C2, <code>=B2*1,2</code> signifie « la cellule juste à gauche × 1,2 ». Recopiée en C3, elle devient <code>=B3*1,2</code>.</p>
<table class="o-table"><thead><tr><th>Formule en C2</th><th>Recopiée en C5</th><th>Recopiée en D2</th></tr></thead><tbody>
<tr><td><code>=B2*2</code></td><td><code>=B5*2</code></td><td><code>=C2*2</code></td></tr>
</tbody></table>
<p>C'est exactement ce que l'on veut pour un calcul ligne par ligne. Mais si la formule pointe vers une cellule <b>unique</b> (un taux, un total), la référence va « glisser » et viser des cellules vides.</p>
<div class="callout bad"><b>Erreur fréquente :</b> <code>=B2/B11</code> recopié en E3 devient <code>=B3/B12</code> : B12 est vide, résultat <code>#DIV/0!</code>.</div>`
        },
        {
          title: "Références absolues et mixtes : le $ et la touche F4",
          html: `<p>Le symbole <code>$</code> <b>verrouille</b> la colonne et/ou la ligne qui le suit :</p>
<table class="o-table"><thead><tr><th>Écriture</th><th>Type</th><th>Lors de la recopie</th></tr></thead><tbody>
<tr><td><code>B2</code></td><td>Relative</td><td>colonne et ligne bougent</td></tr>
<tr><td><code>$B$2</code></td><td>Absolue</td><td>rien ne bouge</td></tr>
<tr><td><code>B$2</code></td><td>Mixte</td><td>la ligne 2 reste fixe</td></tr>
<tr><td><code>$B2</code></td><td>Mixte</td><td>la colonne B reste fixe</td></tr>
</tbody></table>
<p>Pendant la saisie, placez le curseur sur la référence et appuyez sur <kbd>F4</kbd> : elle passe successivement par <code>$B$2</code> → <code>B$2</code> → <code>$B2</code> → <code>B2</code>.</p>
<pre class="code">=B2*(1+$H$1)     prix TTC avec un taux de TVA unique en H1
=B2/$B$11        part de chaque ligne dans le total</pre>
<div class="callout ok"><b>Bonne pratique :</b> ne jamais écrire un taux en dur (<code>*1,2</code>) dans des dizaines de formules. Placez-le dans une cellule « paramètre » et référencez-la en absolu : un seul changement met à jour tout le tableau.</div>
<div class="o-tip">💡 Pour une table de multiplication, la formule <code>=$A2*B$1</code> est le cas d'école de la référence mixte.</div>`
        },
        {
          title: "Mettre en forme un tableau",
          html: `<p>Un tableau lisible se repère en un coup d'œil : en-têtes distincts, montants alignés, peu de couleurs.</p>
<div class="o-path">Accueil › Police › Couleur de remplissage</div>
<ul>
<li><b>Gras</b> (<kbd>Ctrl</kbd>+<kbd>G</kbd> dans Excel en français) et <b>remplissage</b> pour la ligne d'en-têtes.</li>
<li><b>Bordures</b> : Accueil › Police › Bordures › Toutes les bordures.</li>
<li><b>Alignement</b> : texte à gauche, nombres à droite, en-têtes centrés. « Renvoyer à la ligne automatiquement » pour les en-têtes longs.</li>
<li><b>Styles de cellules</b> : Accueil › Styles › Styles de cellules (Titre, Total, Entrée…).</li>
<li><b>Reproduire la mise en forme</b> (le pinceau) : copie le format d'une cellule vers d'autres.</li>
</ul>
<div class="callout warn"><b>Piège :</b> « Fusionner et centrer » casse le tri, le filtre et la sélection de colonnes. Préférez <b>Centrer sur plusieurs colonnes</b> (Ctrl+1 › Alignement › Horizontal).</div>`
        }
      ],
      shortcuts: [
        { keys: "F4", desc: "Basculer la référence relative / absolue / mixte" },
        { keys: "Ctrl+D", desc: "Recopier vers le bas" },
        { keys: "Ctrl+R", desc: "Recopier vers la droite" },
        { keys: "Ctrl+G", desc: "Gras" },
        { keys: "Ctrl+Maj+%", desc: "Format Pourcentage" },
        { keys: "Ctrl+Maj+$", desc: "Format Monétaire" },
        { keys: "Ctrl+1", desc: "Format de cellule" }
      ],
      labs: [
        {
          id: "x02-lab1",
          title: "Lab 1 — La grille tarifaire des fournitures",
          kind: "sheet",
          scenario: `<p>Chez <b>BureauPlus</b>, distributeur de fournitures, vous préparez la grille tarifaire client. Les prix HT sont saisis ; la TVA (H1) et la remise fidélité (H2) sont des <b>paramètres</b> qui changeront l'an prochain. Construisez des formules qui s'adaptent automatiquement.</p>`,
          setup: {
            cols: 8, rows: 12, sheetName: "Tarifs",
            data: {
              A1: "Article", B1: "Prix HT", C1: "Prix TTC", D1: "TTC remisé", E1: "Part du HT",
              A2: "Ramette papier A4", B2: 4.9,
              A3: "Stylos bille (boîte de 50)", B3: 12.5,
              A4: "Agrafeuse", B4: 18,
              A5: "Classeur à levier", B5: 3.2,
              A6: "Notes repositionnables (lot 12)", B6: 9.8,
              A7: "Calculatrice de bureau", B7: 24.9,
              A8: "Lampe de bureau LED", B8: 39,
              A9: "Corbeille à papier", B9: 7.5,
              A10: "Tapis de souris", B10: 6,
              A11: "Total",
              G1: "TVA", H1: 0.2, G2: "Remise", H2: 0.05
            },
            fmt: { B2: "eur", B3: "eur", B4: "eur", B5: "eur", B6: "eur", B7: "eur", B8: "eur", B9: "eur", B10: "eur", H1: "pct", H2: "pct" },
            widths: { A: 210 }
          },
          tasks: [
            {
              text: "En <b>C2</b>, calcule le prix TTC avec <code>=B2*(1+$H$1)</code> puis recopie jusqu'en <b>C10</b>.",
              hint: "Tape =B2*(1+H1 puis F4 pour obtenir $H$1. Ensuite sélectionne C2:C10 et Ctrl+D.",
              check: a => a.hasAbs("C2") && a.isFormula("C10") && a.near(a.num("C10"), a.num("B10") * (1 + a.num("H1"))) && a.near(a.num("C2"), a.num("B2") * (1 + a.num("H1")))
            },
            {
              text: "En <b>D2</b>, applique la remise : <code>=C2*(1-$H$2)</code>, recopiée jusqu'en <b>D10</b>.",
              hint: "C2 reste relatif (il doit glisser), $H$2 est absolu (il ne doit pas bouger).",
              check: a => a.hasAbs("D2") && a.isFormula("D10") && a.near(a.num("D10"), a.num("C10") * (1 - a.num("H2"))) && a.near(a.num("D5"), a.num("C5") * (1 - a.num("H2")))
            },
            {
              text: "En <b>B11</b>, calcule le total des prix HT avec <b>SOMME</b>.",
              hint: "=SOMME(B2:B10)",
              check: a => a.hasFn("B11", "SOMME") && a.near(a.num("B11"), 125.8)
            },
            {
              text: "En <b>E2</b>, calcule la part de chaque article dans le total HT (<code>=B2/$B$11</code>), recopie jusqu'en <b>E10</b> et applique le format <b>pourcentage</b>.",
              hint: "Sans $ sur B11, la recopie donne #DIV/0!. Format : Accueil › Nombre › Pourcentage.",
              check: a => a.hasAbs("E2") && a.near(a.num("E10"), 6 / 125.8, 0.0005) && a.near(a.num("E2"), 4.9 / 125.8, 0.0005) && a.style("E2").fmt === "pct" && a.style("E10").fmt === "pct"
            },
            {
              text: "Le gouvernement passe ces articles à <b>5,5 %</b> de TVA : modifie uniquement <b>H1</b> et constate que toute la grille se met à jour.",
              hint: "Tape 0,055 (ou 5,5 %) dans H1.",
              check: a => a.near(a.num("H1"), 0.055, 0.0001) && a.isFormula("C2") && a.near(a.num("C2"), 5.1695)
            },
            {
              text: "Mise en forme : en-têtes <b>A1:E1</b> en gras avec une <b>couleur de remplissage</b>.",
              hint: "Sélectionne A1:E1 → Gras puis Couleur de remplissage.",
              check: a => a.style("A1").bold && a.style("E1").bold && !!a.style("A1").fill && !!a.style("E1").fill
            }
          ],
          solution: `<ol>
<li>C2 : <code>=B2*(1+$H$1)</code> → 5,88 € ; recopie C2:C10 (<kbd>Ctrl</kbd>+<kbd>D</kbd>).</li>
<li>D2 : <code>=C2*(1-$H$2)</code> → 5,59 € ; recopie D2:D10.</li>
<li>B11 : <code>=SOMME(B2:B10)</code> → 125,80 €.</li>
<li>E2 : <code>=B2/$B$11</code> → 3,90 % ; recopie jusqu'en E10 (4,77 %), format Pourcentage.</li>
<li>H1 = 0,055 : C2 passe à 5,17 €, toutes les lignes suivent car elles pointent vers $H$1.</li>
<li>A1:E1 : Gras + Couleur de remplissage.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "La formule =B2*$H$1 en C2 est recopiée en C7. Que devient-elle ?",
          choices: ["=B7*$H$6", "=B7*$H$1", "=B2*$H$1", "=C7*$H$1"],
          answer: 1,
          explain: "B2 est relative et devient B7 ; $H$1 est absolue et ne bouge pas."
        },
        {
          q: "Quelle touche fait défiler les types de référence ($A$1, A$1, $A1, A1) pendant la saisie ?",
          choices: ["F2", "F4", "F9", "Ctrl+$"],
          answer: 1,
          explain: "F4 bascule entre les quatre types de référence pour la référence sous le curseur."
        },
        {
          q: "Dans la référence mixte B$2, qu'est-ce qui reste fixe lors de la recopie ?",
          choices: ["La colonne B", "La ligne 2", "Les deux", "Aucun"],
          answer: 1,
          explain: "Le $ verrouille ce qui le suit immédiatement : ici la ligne 2. La colonne B peut changer."
        },
        {
          q: "Vous recopiez =B2/B11 vers le bas et obtenez #DIV/0! dès la 2e ligne. Pourquoi ?",
          choices: ["B11 est au format texte", "La référence au total a glissé vers B12, vide", "Il faut utiliser SOMME", "La colonne est trop étroite"],
          answer: 1,
          explain: "Sans $, B11 devient B12, B13… cellules vides (= 0). Il faut écrire $B$11."
        },
        {
          q: "Pourquoi éviter « Fusionner et centrer » sur un tableau de données ?",
          choices: ["Cela efface les formules", "Cela gêne le tri, le filtre et la sélection", "Ce n'est pas disponible en français", "Cela augmente la taille du fichier"],
          answer: 1,
          explain: "Les cellules fusionnées perturbent tri, filtre et sélection. « Centrer sur plusieurs colonnes » donne le même rendu sans ces inconvénients."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 2 — x03                                                       */
    /* ================================================================== */
    {
      id: "x03",
      level: 2,
      title: "Fonctions logiques et statistiques conditionnelles",
      icon: "🧠",
      duration: 60,
      summary: "Faire décider Excel : SI, ET, OU pour les règles métier, NB.SI et SOMME.SI pour compter et totaliser, SIERREUR pour des tableaux propres.",
      objectives: [
        "Être capable d'écrire un test avec SI et des opérateurs de comparaison",
        "Être capable de combiner plusieurs conditions avec ET et OU",
        "Être capable de compter et totaliser selon un critère avec NB.SI, SOMME.SI et MOYENNE.SI",
        "Être capable d'intercepter une erreur avec SIERREUR"
      ],
      lessons: [
        {
          title: "La fonction SI",
          html: `<p><code>SI</code> renvoie une valeur si une condition est vraie, une autre sinon :</p>
<pre class="code">=SI(test_logique; valeur_si_vrai; valeur_si_faux)
=SI(D2>=E2;"Oui";"Non")
=SI(C2>1000;C2*5%;0)</pre>
<p>Opérateurs de comparaison : <code>=</code> <code>&lt;&gt;</code> (différent) <code>&gt;</code> <code>&lt;</code> <code>&gt;=</code> <code>&lt;=</code>. Les textes s'écrivent entre guillemets droits ; les nombres sans guillemets.</p>
<div class="o-path">Formules › Bibliothèque de fonctions › Logique › SI</div>
<div class="callout warn"><b>Piège :</b> <code>=SI(C2&gt;"1000";…)</code> compare à un <i>texte</i> : le résultat est faux sans message d'erreur. Jamais de guillemets autour d'un nombre.</div>
<div class="o-tip">💡 Pour un résultat vide, utilisez deux guillemets : <code>=SI(B2="";"";B2*2)</code>.</div>`
        },
        {
          title: "Combiner les conditions : ET, OU, NON et SI imbriqués",
          html: `<p><code>ET</code> renvoie VRAI si <b>toutes</b> les conditions sont vraies, <code>OU</code> si <b>au moins une</b> l'est, <code>NON</code> inverse.</p>
<pre class="code">=SI(ET(D2>=E2;C2>0);D2*3%;0)        objectif atteint ET déjà présent l'an dernier
=SI(OU(B2="Nord";B2="Est");"Zone 1";"Zone 2")</pre>
<p>Pour plus de deux issues, on peut imbriquer des SI… ou, plus lisible dans Excel 365, utiliser <code>SI.CONDITIONS</code> :</p>
<pre class="code">=SI(C2>=200000;"A";SI(C2>=150000;"B";"C"))
=SI.CONDITIONS(C2>=200000;"A";C2>=150000;"B";VRAI;"C")</pre>
<div class="callout ok"><b>Bonne pratique :</b> dans un SI imbriqué, testez les seuils dans l'ordre (du plus grand au plus petit). Le premier test vrai l'emporte.</div>`
        },
        {
          title: "Compter et additionner sous condition",
          html: `<table class="o-table"><thead><tr><th>Fonction</th><th>Syntaxe</th><th>Exemple</th></tr></thead><tbody>
<tr><td>NB</td><td>compte les nombres</td><td><code>=NB(C2:C11)</code></td></tr>
<tr><td>NBVAL</td><td>compte les cellules non vides</td><td><code>=NBVAL(A2:A11)</code></td></tr>
<tr><td>NB.SI</td><td>(plage ; critère)</td><td><code>=NB.SI(F2:F11;"Oui")</code></td></tr>
<tr><td>SOMME.SI</td><td>(plage_critère ; critère ; plage_somme)</td><td><code>=SOMME.SI(B2:B11;"Nord";D2:D11)</code></td></tr>
<tr><td>MOYENNE.SI</td><td>(plage_critère ; critère ; plage_moyenne)</td><td><code>=MOYENNE.SI(B2:B11;"Sud";D2:D11)</code></td></tr>
</tbody></table>
<p>Le critère peut être un texte (<code>"Nord"</code>), une comparaison entre guillemets (<code>"&gt;150000"</code>, <code>"&lt;&gt;Sud"</code>), un joker (<code>"*bois"</code>) ou une cellule (<code>J1</code>). Pour combiner cellule et opérateur : <code>"&gt;"&amp;J1</code>.</p>
<div class="callout warn"><b>Piège :</b> l'ordre des arguments de SOMME.SI place la plage de <b>critère</b> en premier et la plage à <b>additionner</b> en dernier. Les deux plages doivent avoir la même taille.</div>`
        },
        {
          title: "Gérer les erreurs avec SIERREUR",
          html: `<table class="o-table"><thead><tr><th>Erreur</th><th>Cause typique</th></tr></thead><tbody>
<tr><td><code>#DIV/0!</code></td><td>division par zéro ou par une cellule vide</td></tr>
<tr><td><code>#N/A</code></td><td>valeur introuvable (recherche)</td></tr>
<tr><td><code>#NOM?</code></td><td>fonction mal orthographiée, guillemets oubliés</td></tr>
<tr><td><code>#VALEUR!</code></td><td>calcul sur du texte</td></tr>
<tr><td><code>#REF!</code></td><td>référence vers une cellule supprimée</td></tr>
</tbody></table>
<pre class="code">=SIERREUR(D2/C2-1;"Nouveau")</pre>
<p>SIERREUR renvoie le calcul s'il fonctionne, sinon la valeur de remplacement.</p>
<div class="callout bad"><b>Erreur fréquente :</b> envelopper toutes les formules dans SIERREUR « pour être tranquille ». Vous masquez aussi les vraies erreurs (#NOM?, #REF!). Ne l'utilisez que pour un cas prévu et compris.</div>
<div class="o-path">Formules › Vérification des formules › Évaluation de formule</div>`
        }
      ],
      shortcuts: [
        { keys: "Maj+F3", desc: "Insérer une fonction (assistant)" },
        { keys: "Ctrl+A", desc: "Après le nom d'une fonction : ouvrir ses arguments" },
        { keys: "F9", desc: "Recalculer / évaluer une partie de formule sélectionnée" },
        { keys: "Ctrl+Maj+Entrée", desc: "Valider une formule matricielle (anciennes versions)" },
        { keys: "Ctrl+Maj+U", desc: "Développer / réduire la barre de formule" },
        { keys: "F4", desc: "Figer la référence d'une plage de critères" }
      ],
      labs: [
        {
          id: "x03-lab1",
          title: "Lab 1 — Les primes de l'équipe commerciale",
          kind: "sheet",
          scenario: `<p>Vous êtes contrôleur de gestion chez <b>Hydra Équipements</b>. La DRH attend le calcul des primes annuelles. Règle : une prime de <b>3 % du CA 2026</b> est versée si l'objectif est atteint <b>et</b> si le commercial était déjà présent en 2025 (les nouveaux arrivants ont un CA N-1 à 0 et relèvent d'un autre dispositif).</p>`,
          setup: {
            cols: 11, rows: 12, sheetName: "Primes",
            data: {
              A1: "Commercial", B1: "Région", C1: "CA 2025", D1: "CA 2026", E1: "Objectif", F1: "Atteint", G1: "Évolution", H1: "Prime",
              A2: "Martin", B2: "Nord", C2: 182000, D2: 195000, E2: 190000,
              A3: "Bernard", B3: "Sud", C3: 154000, D3: 149000, E3: 160000,
              A4: "Petit", B4: "Est", C4: 0, D4: 88000, E4: 80000,
              A5: "Robert", B5: "Nord", C5: 210000, D5: 232000, E5: 220000,
              A6: "Richard", B6: "Ouest", C6: 98000, D6: 121000, E6: 110000,
              A7: "Dubois", B7: "Sud", C7: 175000, D7: 168000, E7: 170000,
              A8: "Moreau", B8: "Est", C8: 133000, D8: 141000, E8: 140000,
              A9: "Laurent", B9: "Ouest", C9: 0, D9: 64000, E9: 70000,
              A10: "Simon", B10: "Nord", C10: 119000, D10: 117000, E10: 125000,
              A11: "Michel", B11: "Sud", C11: 142000, D11: 158000, E11: 150000,
              J2: "Objectifs atteints", J3: "CA 2026 Nord", J4: "Total des primes"
            },
            bold: ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"],
            fmt: { G2: "pct", G3: "pct", G4: "pct", G5: "pct", G6: "pct", G7: "pct", G8: "pct", G9: "pct", G10: "pct", G11: "pct" },
            widths: { J: 150 }
          },
          tasks: [
            {
              text: "En <b>F2</b>, affiche <code>\"Oui\"</code> si le CA 2026 atteint l'objectif, <code>\"Non\"</code> sinon, puis recopie jusqu'en <b>F11</b>.",
              hint: "=SI(D2>=E2;\"Oui\";\"Non\")",
              check: a => a.hasFn("F2", "SI") && same(a.val("F2"), "Oui") && same(a.val("F3"), "Non") && same(a.val("F7"), "Non") && same(a.val("F8"), "Oui") && same(a.val("F11"), "Oui")
            },
            {
              text: "En <b>G2</b>, calcule l'évolution <code>D2/C2-1</code> en affichant <code>\"Nouveau\"</code> à la place de l'erreur pour les nouveaux arrivants. Recopie jusqu'en <b>G11</b>.",
              hint: "=SIERREUR(D2/C2-1;\"Nouveau\") — Petit et Laurent ont un CA 2025 à 0 (#DIV/0!).",
              check: a => a.hasFn("G2", "SIERREUR") && a.near(a.num("G2"), 0.0714, 0.001) && same(a.val("G4"), "Nouveau") && same(a.val("G9"), "Nouveau") && a.near(a.num("G11"), 0.1127, 0.001)
            },
            {
              text: "En <b>H2</b>, calcule la prime : 3 % du CA 2026 si l'objectif est atteint <b>ET</b> que le CA 2025 est supérieur à 0, sinon 0. Recopie jusqu'en <b>H11</b>.",
              hint: "=SI(ET(D2>=E2;C2>0);D2*3%;0)",
              check: a => a.hasFn("H2", "ET") && a.hasFn("H2", "SI") && a.near(a.num("H2"), 5850) && a.num("H4") === 0 && a.near(a.num("H5"), 6960) && a.num("H10") === 0
            },
            {
              text: "En <b>K2</b>, compte le nombre d'objectifs atteints avec <b>NB.SI</b>.",
              hint: "=NB.SI(F2:F11;\"Oui\")",
              check: a => a.hasFn("K2", "NB.SI") && a.num("K2") === 6
            },
            {
              text: "En <b>K3</b>, calcule le CA 2026 de la région Nord avec <b>SOMME.SI</b>.",
              hint: "=SOMME.SI(B2:B11;\"Nord\";D2:D11)",
              check: a => a.hasFn("K3", "SOMME.SI") && a.near(a.num("K3"), 544000)
            },
            {
              text: "En <b>K4</b>, calcule le montant total des primes à verser.",
              hint: "=SOMME(H2:H11)",
              check: a => a.isFormula("K4") && a.near(a.num("K4"), 25410)
            }
          ],
          solution: `<ol>
<li>F2 : <code>=SI(D2>=E2;"Oui";"Non")</code> → 6 « Oui » (Martin, Petit, Robert, Richard, Moreau, Michel).</li>
<li>G2 : <code>=SIERREUR(D2/C2-1;"Nouveau")</code> → 7,14 % pour Martin, « Nouveau » pour Petit et Laurent.</li>
<li>H2 : <code>=SI(ET(D2>=E2;C2>0);D2*3%;0)</code> → Martin 5 850, Robert 6 960, Richard 3 630, Moreau 4 230, Michel 4 740 ; Petit 0 (nouvel arrivant).</li>
<li>K2 : <code>=NB.SI(F2:F11;"Oui")</code> → 6.</li>
<li>K3 : <code>=SOMME.SI(B2:B11;"Nord";D2:D11)</code> → 544 000.</li>
<li>K4 : <code>=SOMME(H2:H11)</code> → 25 410.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Que renvoie =SI(ET(5>3;2>4);\"A\";\"B\") ?",
          choices: ["A", "B", "VRAI", "#VALEUR!"],
          answer: 1,
          explain: "ET exige que toutes les conditions soient vraies ; 2>4 est faux, donc SI renvoie la valeur_si_faux « B »."
        },
        {
          q: "Quelle formule additionne les montants de D2:D11 pour la région « Sud » (colonne B) ?",
          choices: ["=SOMME.SI(D2:D11;\"Sud\";B2:B11)", "=SOMME.SI(B2:B11;\"Sud\";D2:D11)", "=NB.SI(B2:B11;\"Sud\")", "=SOMME(SI(B2:B11=Sud))"],
          answer: 1,
          explain: "SOMME.SI(plage_critère ; critère ; plage_somme) : on cherche « Sud » dans B et on additionne D."
        },
        {
          q: "Comment compter les CA supérieurs à la valeur de la cellule J1 ?",
          choices: ["=NB.SI(D2:D11;\">J1\")", "=NB.SI(D2:D11;\">\"&J1)", "=NB.SI(D2:D11;>J1)", "=NB(D2:D11>J1)"],
          answer: 1,
          explain: "L'opérateur est un texte qu'on concatène avec la cellule : \">\"&J1. Écrit \">J1\", Excel chercherait le texte « J1 »."
        },
        {
          q: "Une formule affiche #NOM?. Quelle est la cause la plus probable ?",
          choices: ["Une division par zéro", "Un nom de fonction mal orthographié ou des guillemets oubliés", "Une valeur introuvable", "Une colonne trop étroite"],
          answer: 1,
          explain: "#NOM? signale un nom inconnu : faute de frappe (SOMEM), fonction anglaise non reconnue, ou texte sans guillemets."
        },
        {
          q: "Pourquoi ne faut-il pas mettre SIERREUR sur toutes les formules par principe ?",
          choices: ["Elle ralentit fortement Excel", "Elle masque aussi les erreurs imprévues qui signalent un vrai problème", "Elle ne fonctionne qu'avec RECHERCHEV", "Elle convertit les nombres en texte"],
          answer: 1,
          explain: "SIERREUR attrape toutes les erreurs, y compris #REF! ou #NOM? dues à un bug. On la réserve aux cas attendus (division par zéro connue, recherche sans résultat)."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 2 — x04                                                       */
    /* ================================================================== */
    {
      id: "x04",
      level: 2,
      title: "Trier, filtrer, mettre en évidence : exploiter une liste de données",
      icon: "🔎",
      duration: 55,
      summary: "Transformer une liste brute en outil de pilotage : tri, filtres, mise en forme conditionnelle, tableaux structurés et volets figés.",
      objectives: [
        "Être capable de trier une liste sur une ou plusieurs colonnes",
        "Être capable de filtrer des données (textes, nombres, couleurs)",
        "Être capable de créer des règles de mise en forme conditionnelle",
        "Être capable de mettre une plage sous forme de tableau et d'en tirer parti",
        "Être capable de figer les volets pour garder les en-têtes visibles"
      ],
      lessons: [
        {
          title: "Trier les données",
          html: `<p>Cliquez dans <b>une seule cellule</b> de la colonne à trier (pas toute la colonne) : Excel détecte la zone contiguë et la ligne d'en-têtes.</p>
<div class="o-path">Données › Trier et filtrer › Trier de A à Z</div>
<ul>
<li><b>A → Z / Z → A</b> : tri rapide sur la colonne active.</li>
<li><b>Trier</b> (bouton avec boîte de dialogue) : plusieurs niveaux, ex. Catégorie puis Stock croissant ; tri par couleur de cellule ; liste personnalisée (Lun, Mar…).</li>
</ul>
<div class="callout bad"><b>Erreur fréquente :</b> sélectionner une seule colonne puis trier en choisissant « Continuer avec la sélection actuelle ». Les lignes sont désolidarisées : chaque référence se retrouve avec le stock d'un autre produit. Choisissez toujours « Étendre la sélection ».</div>
<div class="o-tip">💡 Avant un tri risqué, ajoutez une colonne « N° » numérotée 1, 2, 3… : un tri croissant sur elle restaure l'ordre d'origine.</div>`
        },
        {
          title: "Filtrer",
          html: `<p>Le filtre ajoute des flèches déroulantes dans la ligne d'en-têtes et <b>masque</b> les lignes qui ne correspondent pas (rien n'est supprimé).</p>
<div class="o-path">Données › Trier et filtrer › Filtrer</div>
<ul>
<li>Cases à cocher pour choisir des valeurs (Connectique, Audio…).</li>
<li><b>Filtres textuels</b> : contient, commence par… ; <b>Filtres numériques</b> : supérieur à, 10 premiers, au-dessus de la moyenne ; <b>Filtres chronologiques</b> : ce mois-ci, trimestre précédent…</li>
<li><b>Filtrer par couleur</b> : très pratique combiné à la mise en forme conditionnelle.</li>
</ul>
<p>Les numéros de ligne passent en bleu et la barre d'état indique « 5 enregistrements trouvés sur 12 ». Pour tout réafficher : Données › Effacer.</p>
<div class="callout warn"><b>Piège :</b> <code>SOMME</code> additionne aussi les lignes masquées par le filtre. Pour ne totaliser que le visible, utilisez <code>=SOUS.TOTAL(109;D2:D13)</code> ou <code>=AGREGAT(9;5;D2:D13)</code>.</div>`
        },
        {
          title: "La mise en forme conditionnelle",
          html: `<p>La mise en forme conditionnelle (MFC) colore automatiquement les cellules selon leur valeur : l'œil voit l'alerte avant de lire le chiffre.</p>
<div class="o-path">Accueil › Styles › Mise en forme conditionnelle › Règles de mise en surbrillance des cellules › Inférieur à…</div>
<ul>
<li><b>Règles de surbrillance</b> : supérieur à, inférieur à, entre, égal à, texte qui contient, date, doublons.</li>
<li><b>Règles des valeurs plus/moins élevées</b> : 10 premiers, au-dessus de la moyenne.</li>
<li><b>Barres de données, nuances de couleurs, jeux d'icônes</b> : mini-graphiques dans les cellules.</li>
<li><b>Nouvelle règle › Utiliser une formule</b> : ex. <code>=$D2&lt;$E2</code> pour colorer toute la ligne quand le stock est sous le seuil.</li>
</ul>
<div class="o-path">Accueil › Styles › Mise en forme conditionnelle › Gérer les règles</div>
<div class="callout ok"><b>Bonne pratique :</b> une ou deux couleurs maximum, avec un sens clair (rouge = action requise). Un tableau arc-en-ciel n'alerte plus personne.</div>`
        },
        {
          title: "Mettre sous forme de tableau",
          html: `<p>Un <b>tableau structuré</b> transforme une plage en objet intelligent :</p>
<div class="o-path">Accueil › Styles › Mettre sous forme de tableau</div>
<ul>
<li>Lignes à bandes, filtre automatique, en-têtes qui restent visibles en défilant.</li>
<li>Il <b>s'étend tout seul</b> quand on ajoute une ligne ou une colonne : graphiques, TCD et formules suivent.</li>
<li>Formules lisibles avec les <b>références structurées</b> : <code>=[@Stock]*[@Prix]</code> ou <code>=SOMME(Stock[Prix])</code>. Une formule saisie dans une colonne se propage à toute la colonne.</li>
<li>Onglet contextuel <b>Création de tableau</b> : nom du tableau, ligne Total, segments (filtres visuels).</li>
</ul>
<div class="o-tip">💡 Renommez vos tableaux (Création de tableau › Propriétés › Nom du tableau) : <code>tblStock</code> se lit mieux que <code>Tableau3</code> dans les formules.</div>`
        },
        {
          title: "Figer les volets",
          html: `<p>Sur une liste de 500 lignes, les en-têtes disparaissent dès qu'on descend. <b>Figer</b> les garde à l'écran.</p>
<div class="o-path">Affichage › Fenêtre › Figer les volets › Figer la ligne supérieure</div>
<ul>
<li><b>Figer la ligne supérieure</b> / <b>la première colonne</b>.</li>
<li><b>Figer les volets</b> : fige tout ce qui est au-dessus et à gauche de la cellule active. Cliquez en B2 pour figer la ligne 1 <i>et</i> la colonne A.</li>
</ul>
<div class="callout warn"><b>Piège :</b> figer ne concerne que l'affichage à l'écran. Pour répéter les en-têtes à l'<b>impression</b> : Mise en page › Imprimer les titres › Lignes à répéter en haut.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+Maj+L", desc: "Activer / désactiver le filtre" },
        { keys: "Alt+Flèche bas", desc: "Ouvrir la liste du filtre de la colonne" },
        { keys: "Ctrl+L", desc: "Mettre sous forme de tableau (Ctrl+T fonctionne aussi)" },
        { keys: "Ctrl+Maj+Flèche bas", desc: "Étendre la sélection jusqu'en bas des données" },
        { keys: "Ctrl+A", desc: "Sélectionner la zone de données" },
        { keys: "Alt+F1", desc: "Créer un graphique sur la feuille" },
        { keys: "Ctrl+Q", desc: "Analyse rapide de la sélection (dans Excel)" }
      ],
      labs: [
        {
          id: "x04-lab1",
          title: "Lab 1 — L'inventaire de l'entrepôt informatique",
          kind: "sheet",
          scenario: `<p>Responsable logistique chez <b>TechDépôt</b>, vous préparez la réunion de réapprovisionnement de lundi. La liste du stock doit faire ressortir les articles critiques et rester confortable à consulter.</p>`,
          setup: {
            cols: 6, rows: 14, sheetName: "Stock",
            data: {
              A1: "Référence", B1: "Désignation", C1: "Catégorie", D1: "Stock", E1: "Seuil", F1: "Prix",
              A2: "REF-101", B2: "Câble HDMI 2 m", C2: "Connectique", D2: 45, E2: 20, F2: 8.9,
              A3: "REF-102", B3: "Souris sans fil", C3: "Périphériques", D3: 12, E3: 15, F3: 19.9,
              A4: "REF-103", B4: "Clavier AZERTY", C4: "Périphériques", D4: 8, E4: 10, F4: 29.9,
              A5: "REF-104", B5: "Écran 24 pouces", C5: "Écrans", D5: 6, E5: 5, F5: 149,
              A6: "REF-105", B6: "Hub USB-C", C6: "Connectique", D6: 31, E6: 15, F6: 24.5,
              A7: "REF-106", B7: "Webcam HD", C7: "Périphériques", D7: 17, E7: 10, F7: 39,
              A8: "REF-107", B8: "Casque audio", C8: "Audio", D8: 4, E8: 8, F8: 59,
              A9: "REF-108", B9: "Rallonge 5 prises", C9: "Électricité", D9: 52, E9: 25, F9: 12.4,
              A10: "REF-109", B10: "Écran 27 pouces", C10: "Écrans", D10: 3, E10: 4, F10: 219,
              A11: "REF-110", B11: "Adaptateur DP-HDMI", C11: "Connectique", D11: 26, E11: 20, F11: 11.5,
              A12: "REF-111", B12: "Enceinte Bluetooth", C12: "Audio", D12: 9, E12: 6, F12: 45,
              A13: "REF-112", B13: "Station d'accueil", C13: "Connectique", D13: 14, E13: 8, F13: 129
            },
            bold: ["A1", "B1", "C1", "D1", "E1", "F1"],
            fmt: { F2: "eur", F3: "eur", F4: "eur", F5: "eur", F6: "eur", F7: "eur", F8: "eur", F9: "eur", F10: "eur", F11: "eur", F12: "eur", F13: "eur" },
            widths: { B: 170, C: 120 }
          },
          tasks: [
            {
              text: "Trie la liste par <b>Stock croissant</b> pour voir d'abord les articles les plus bas.",
              hint: "Clique dans une cellule de la colonne D puis Trier A→Z.",
              check: a => a.sortedBy("D", "asc")
            },
            {
              text: "Applique une <b>mise en forme conditionnelle</b> sur <b>D2:D13</b> : couleur pour les stocks <b>inférieurs à 20</b>.",
              hint: "Sélectionne D2:D13 › Mise en forme conditionnelle › Inférieur à › 20.",
              check: a => a.cf.some(r => covers(r.range, "D", 2, 13) && r.op === "<" && Number(r.v1) === 20)
            },
            {
              text: "Active le <b>filtre</b> sur la ligne d'en-têtes.",
              hint: "Données › Filtrer (Ctrl+Maj+L).",
              check: a => a.filter === true
            },
            {
              text: "<b>Fige la ligne 1</b> pour garder les en-têtes visibles.",
              hint: "Affichage › Figer les volets › Figer la ligne supérieure.",
              check: a => a.frozen === true
            },
            {
              text: "<b>Mets la liste sous forme de tableau</b> pour qu'elle s'étende automatiquement aux futurs articles.",
              hint: "Accueil › Mettre sous forme de tableau.",
              check: a => a.table === true
            }
          ],
          solution: `<ol>
<li>Clic en D2 › Trier A→Z : l'Écran 27 pouces (3) arrive en tête, puis le Casque audio (4), l'Écran 24 pouces (6)…</li>
<li>D2:D13 › MFC « Inférieur à » 20 : 8 articles ressortent (3, 4, 6, 8, 9, 12, 14, 17).</li>
<li>Données › Filtrer : les flèches apparaissent sur la ligne 1.</li>
<li>Affichage › Figer les volets › Figer la ligne supérieure.</li>
<li>Accueil › Mettre sous forme de tableau (avec en-têtes).</li>
</ol>
<p>Pour aller plus loin dans le vrai Excel : une règle par formule <code>=$D2&lt;$E2</code> colore les lignes réellement sous leur seuil (souris, clavier, casque, écran 27 pouces).</p>`
        },
        {
          id: "x04-lab2",
          title: "Lab 2 — Défi clavier : l'analyste pressé",
          kind: "keys",
          scenario: `<p>Votre manager vous donne 30 secondes pour filtrer, sélectionner et visualiser une liste. Prouvez que le clavier va plus vite que la souris.</p>`,
          setup: {
            sequence: [
              { keys: "Ctrl+Maj+L", desc: "Activer le filtre" },
              { keys: "Alt+Flèche bas", desc: "Ouvrir la liste du filtre" },
              { keys: "Ctrl+A", desc: "Sélectionner toute la zone de données" },
              { keys: "Ctrl+Maj+Flèche bas", desc: "Étendre la sélection jusqu'en bas" },
              { keys: "Ctrl+Maj+%", desc: "Appliquer le format Pourcentage" },
              { keys: "Alt+F1", desc: "Insérer un graphique sur la feuille" },
              { keys: "F4", desc: "Répéter la dernière action / figer une référence" },
              { keys: "F9", desc: "Recalculer le classeur" }
            ]
          },
          solution: `<p><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>L</kbd> est le raccourci le plus rentable de ce module : un filtre en une seconde. <kbd>Alt</kbd>+<kbd>F1</kbd> crée instantanément un graphique à partir de la sélection.</p>`
        }
      ],
      quiz: [
        {
          q: "Que fait un filtre sur les lignes qui ne correspondent pas au critère ?",
          choices: ["Il les supprime", "Il les masque temporairement", "Il les déplace en bas", "Il les colore en gris"],
          answer: 1,
          explain: "Le filtre masque les lignes ; elles réapparaissent quand on efface le filtre. Rien n'est supprimé."
        },
        {
          q: "Une liste est filtrée. Quelle fonction totalise uniquement les lignes visibles ?",
          choices: ["SOMME", "SOMME.SI", "SOUS.TOTAL(109;…)", "NBVAL"],
          answer: 2,
          explain: "SOUS.TOTAL avec le code 109 (ou 9) ignore les lignes masquées par un filtre ; SOMME les inclut."
        },
        {
          q: "Quel avantage principal apporte « Mettre sous forme de tableau » ?",
          choices: ["Le fichier est protégé", "La plage s'étend automatiquement avec les nouvelles lignes", "Les formules deviennent des valeurs", "Les cellules sont fusionnées"],
          answer: 1,
          explain: "Un tableau structuré s'agrandit automatiquement ; formules, graphiques et TCD qui s'y réfèrent suivent sans retouche."
        },
        {
          q: "Pour figer à la fois la ligne 1 et la colonne A, où placer la cellule active avant Figer les volets ?",
          choices: ["A1", "A2", "B1", "B2"],
          answer: 3,
          explain: "Figer les volets fige tout ce qui est au-dessus et à gauche de la cellule active : en B2, la ligne 1 et la colonne A."
        },
        {
          q: "Quelle règle de MFC par formule colore toute la ligne quand le stock (D) est sous le seuil (E) ?",
          choices: ["=D2<E2", "=$D2<$E2", "=$D$2<$E$2", "=D$2<E$2"],
          answer: 1,
          explain: "La colonne est figée ($D, $E) pour que chaque cellule de la ligne teste D et E ; la ligne reste relative pour s'adapter à chaque ligne."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 3 — x05                                                       */
    /* ================================================================== */
    {
      id: "x05",
      level: 3,
      title: "Fonctions de recherche, de texte et de date",
      icon: "🧩",
      duration: 70,
      summary: "Croiser des tableaux avec RECHERCHEV, RECHERCHEX et INDEX/EQUIV, puis nettoyer et assembler textes et dates comme un pro.",
      objectives: [
        "Être capable d'aller chercher une information dans une table avec RECHERCHEV",
        "Être capable d'utiliser RECHERCHEX et le couple INDEX/EQUIV",
        "Être capable de nettoyer et assembler du texte (SUPPRESPACE, NOMPROPRE, GAUCHE, &)",
        "Être capable d'extraire des informations d'une date (ANNEE, MOIS, JOUR, TEXTE)"
      ],
      lessons: [
        {
          title: "RECHERCHEV",
          html: `<p>RECHERCHEV cherche une valeur dans la <b>première colonne</b> d'une table et renvoie la valeur d'une autre colonne de la même ligne.</p>
<pre class="code">=RECHERCHEV(valeur_cherchée; table; n°_colonne; [valeur_proche])
=RECHERCHEV(B2;$I$2:$L$7;2;FAUX)</pre>
<ul>
<li><b>table</b> : en absolu (<kbd>F4</kbd>) pour la recopie.</li>
<li><b>n°_colonne</b> : 1 = première colonne de la table (pas de la feuille !).</li>
<li><b>FAUX</b> (ou 0) : correspondance exacte. <b>VRAI</b> : correspondance approchée, réservée aux barèmes triés (tranches d'imposition, remises par palier).</li>
</ul>
<div class="callout bad"><b>Erreur fréquente :</b> oublier le 4e argument. La valeur par défaut est VRAI : Excel renvoie un résultat <i>faux mais plausible</i> au lieu de #N/A.</div>
<div class="callout warn"><b>Limites :</b> la clé doit être à gauche ; le n° de colonne est figé (insérer une colonne dans la table casse le résultat).</div>`
        },
        {
          title: "RECHERCHEX et INDEX/EQUIV",
          html: `<p><b>RECHERCHEX</b> (Excel 365 / 2021) remplace avantageusement RECHERCHEV :</p>
<pre class="code">=RECHERCHEX(B2;$I$2:$I$7;$K$2:$K$7)
=RECHERCHEX(B2;$I$2:$I$7;$K$2:$K$7;"Inconnu")</pre>
<ul>
<li>Exacte par défaut, cherche à gauche comme à droite, 4e argument « si non trouvé » intégré.</li>
<li>Peut renvoyer plusieurs colonnes d'un coup et chercher depuis la fin (dernière occurrence).</li>
</ul>
<p><b>INDEX/EQUIV</b> : le duo universel, compatible avec toutes les versions.</p>
<pre class="code">=EQUIV(B2;$I$2:$I$7;0)                        position de B2 dans la liste
=INDEX($L$2:$L$7;EQUIV(B2;$I$2:$I$7;0))       valeur à cette position</pre>
<table class="o-table"><thead><tr><th></th><th>RECHERCHEV</th><th>RECHERCHEX</th><th>INDEX/EQUIV</th></tr></thead><tbody>
<tr><td>Recherche à gauche</td><td>non</td><td>oui</td><td>oui</td></tr>
<tr><td>Exacte par défaut</td><td>non</td><td>oui</td><td>avec 0</td></tr>
<tr><td>Résiste à l'insertion de colonnes</td><td>non</td><td>oui</td><td>oui</td></tr>
<tr><td>Compatible Excel 2016</td><td>oui</td><td>non</td><td>oui</td></tr>
</tbody></table>
<div class="o-tip">💡 Si le fichier circule chez des partenaires équipés d'anciennes versions, RECHERCHEX y affichera #NOM? : préférez INDEX/EQUIV.</div>`
        },
        {
          title: "Nettoyer et assembler du texte",
          html: `<table class="o-table"><thead><tr><th>Fonction</th><th>Exemple</th><th>Résultat</th></tr></thead><tbody>
<tr><td>SUPPRESPACE</td><td><code>=SUPPRESPACE("  LEROY ")</code></td><td>LEROY</td></tr>
<tr><td>NOMPROPRE / MAJUSCULE / MINUSCULE</td><td><code>=NOMPROPRE("LEROY")</code></td><td>Leroy</td></tr>
<tr><td>GAUCHE / DROITE / STXT</td><td><code>=STXT("FR-75-0042";4;2)</code></td><td>75</td></tr>
<tr><td>NBCAR</td><td><code>=NBCAR("Excel")</code></td><td>5</td></tr>
<tr><td>&amp; / CONCAT</td><td><code>=B2&amp;" "&amp;A2</code></td><td>Camille Leroy</td></tr>
<tr><td>TEXTE</td><td><code>=TEXTE(C2;"mmmm aaaa")</code></td><td>mars 2019</td></tr>
</tbody></table>
<pre class="code">=MINUSCULE(GAUCHE(B2;1)&amp;SUPPRESPACE(A2))&amp;"@lumen-conseil.fr"</pre>
<div class="callout warn"><b>Piège :</b> les données importées contiennent souvent des espaces invisibles : <code>"Lyon "</code> ≠ <code>"Lyon"</code>, et RECHERCHEV renvoie #N/A. Un SUPPRESPACE résout la moitié des « bugs » de recherche.</div>
<div class="o-path">Données › Outils de données › Remplissage instantané</div>
<div class="o-tip">💡 Le <b>Remplissage instantané</b> (<kbd>Ctrl</kbd>+<kbd>E</kbd>) devine un motif à partir de deux exemples tapés à la main. Pratique, mais le résultat n'est pas une formule : il ne se met pas à jour.</div>`
        },
        {
          title: "Calculer avec les dates",
          html: `<p>Une date est un nombre (jours depuis le 1/1/1900) : on peut l'additionner, la soustraire, la comparer.</p>
<pre class="code">=AUJOURDHUI()               date du jour (recalculée chaque jour)
=ANNEE(C2)  =MOIS(C2)  =JOUR(C2)
=DATE(2026;12;31)           construit une date
=C2+30                      échéance à 30 jours
=AUJOURDHUI()-C2            nombre de jours écoulés
=DATEDIF(C2;AUJOURDHUI();"y")   ancienneté en années révolues</pre>
<div class="callout warn"><b>Piège :</b> une date saisie <code>2019.03.15</code> ou importée en texte n'est pas reconnue : ANNEE renvoie #VALEUR!. Vérifiez qu'elle est alignée à droite, sinon : Données › Convertir.</div>
<div class="callout ok"><b>Bonne pratique :</b> pour un rapport figé à une date donnée, mettez la date de référence dans une cellule plutôt que AUJOURDHUI() : le résultat restera reproductible.</div>`
        }
      ],
      shortcuts: [
        { keys: "F4", desc: "Figer la table de recherche ($)" },
        { keys: "Ctrl+E", desc: "Remplissage instantané" },
        { keys: "Ctrl+;", desc: "Insérer la date du jour" },
        { keys: "Ctrl+Maj+:", desc: "Insérer l'heure actuelle" },
        { keys: "Ctrl+[", desc: "Sélectionner les antécédents de la formule" },
        { keys: "F5", desc: "Atteindre une cellule ou une plage nommée" }
      ],
      labs: [
        {
          id: "x05-lab1",
          title: "Lab 1 — Enrichir le carnet de commandes",
          kind: "sheet",
          scenario: `<p>Chez <b>Atelier Mobilier Pro</b>, l'export des commandes ne contient que le code client. L'assistante commerciale doit compléter chaque ligne avec la ville, le commercial et la remise contractuelle, à partir du référentiel clients (colonnes I à L).</p>`,
          setup: {
            cols: 12, rows: 12, sheetName: "Commandes",
            data: {
              A1: "Commande", B1: "Client", C1: "Montant HT", D1: "Ville", E1: "Commercial", F1: "Remise", G1: "Net HT",
              A2: "CMD-1001", B2: "C003", C2: 1250,
              A3: "CMD-1002", B3: "C001", C3: 840,
              A4: "CMD-1003", B4: "C005", C4: 2310,
              A5: "CMD-1004", B5: "C002", C5: 460,
              A6: "CMD-1005", B6: "C004", C6: 1780,
              A7: "CMD-1006", B7: "C001", C7: 395,
              A8: "CMD-1007", B8: "C006", C8: 1120,
              A9: "CMD-1008", B9: "C099", C9: 650,
              A10: "CMD-1009", B10: "C005", C10: 980,
              A11: "CMD-1010", B11: "C003", C11: 1515,
              I1: "Code", J1: "Ville", K1: "Commercial", L1: "Remise",
              I2: "C001", J2: "Lyon", K2: "Sophie", L2: 0.05,
              I3: "C002", J3: "Lille", K3: "Karim", L3: 0.03,
              I4: "C003", J4: "Nantes", K4: "Sophie", L4: 0.08,
              I5: "C004", J5: "Bordeaux", K5: "Julie", L5: 0,
              I6: "C005", J6: "Marseille", K6: "Karim", L6: 0.1,
              I7: "C006", J7: "Strasbourg", K7: "Julie", L7: 0.04
            },
            bold: ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "I1", "J1", "K1", "L1"],
            fmt: { L2: "pct", L3: "pct", L4: "pct", L5: "pct", L6: "pct", L7: "pct", C2: "eur", C3: "eur", C4: "eur", C5: "eur", C6: "eur", C7: "eur", C8: "eur", C9: "eur", C10: "eur", C11: "eur" },
            widths: { J: 110 }
          },
          tasks: [
            {
              text: "En <b>D2</b>, récupère la ville avec <b>RECHERCHEV</b> (table <b>$I$2:$L$7</b> en absolu, correspondance exacte), puis recopie jusqu'en <b>D11</b>.",
              hint: "=RECHERCHEV(B2;$I$2:$L$7;2;FAUX) — FAUX peut s'écrire 0.",
              check: a => a.hasFn("D2", "RECHERCHEV") && a.hasAbs("D2") && a.val("D2") === "Nantes" && a.val("D7") === "Lyon" && a.val("D8") === "Strasbourg"
            },
            {
              text: "La commande CMD-1008 affiche <code>#N/A</code> (client C099 absent du référentiel). Modifie la formule de <b>D9</b> pour afficher <code>\"Client inconnu\"</code> à la place.",
              hint: "=SIERREUR(RECHERCHEV(B9;$I$2:$L$7;2;FAUX);\"Client inconnu\") — tu peux l'appliquer à toute la colonne.",
              check: a => a.hasFn("D9", "SIERREUR") && a.hasFn("D9", "RECHERCHEV") && same(a.val("D9"), "Client inconnu")
            },
            {
              text: "En <b>E2</b>, récupère le commercial avec <b>RECHERCHEX</b>, puis recopie jusqu'en <b>E11</b>.",
              hint: "=RECHERCHEX(B2;$I$2:$I$7;$K$2:$K$7)",
              check: a => a.hasFn("E2", "RECHERCHEX") && a.val("E2") === "Sophie" && a.val("E4") === "Karim" && a.val("E6") === "Julie" && a.val("E11") === "Sophie"
            },
            {
              text: "En <b>F2</b>, récupère la remise avec <b>INDEX</b> et <b>EQUIV</b>, puis recopie jusqu'en <b>F11</b>.",
              hint: "=INDEX($L$2:$L$7;EQUIV(B2;$I$2:$I$7;0))",
              check: a => a.hasFn("F2", "INDEX") && a.hasFn("F2", "EQUIV") && a.near(a.num("F2"), 0.08, 0.0001) && a.near(a.num("F4"), 0.1, 0.0001) && a.near(a.num("F8"), 0.04, 0.0001)
            },
            {
              text: "En <b>G2</b>, calcule le montant net <code>=C2*(1-F2)</code> et recopie jusqu'en <b>G11</b>.",
              hint: "La ligne 9 restera en erreur tant que le client n'est pas créé : c'est normal.",
              check: a => a.isFormula("G2") && a.near(a.num("G2"), 1150) && a.near(a.num("G4"), 2079) && a.near(a.num("G11"), 1393.8)
            }
          ],
          solution: `<ol>
<li>D2 : <code>=RECHERCHEV(B2;$I$2:$L$7;2;FAUX)</code> → Nantes, Lyon, Marseille, Lille, Bordeaux, Lyon, Strasbourg, #N/A, Marseille, Nantes.</li>
<li>D9 (ou toute la colonne) : <code>=SIERREUR(RECHERCHEV(B9;$I$2:$L$7;2;FAUX);"Client inconnu")</code>.</li>
<li>E2 : <code>=RECHERCHEX(B2;$I$2:$I$7;$K$2:$K$7)</code> → Sophie, Sophie, Karim…</li>
<li>F2 : <code>=INDEX($L$2:$L$7;EQUIV(B2;$I$2:$I$7;0))</code> → 8 %, 5 %, 10 %…</li>
<li>G2 : <code>=C2*(1-F2)</code> → 1 150 ; G4 = 2 079 ; G11 = 1 393,80.</li>
</ol>`
        },
        {
          id: "x05-lab2",
          title: "Lab 2 — Créer les comptes des collaborateurs",
          kind: "sheet",
          scenario: `<p>Le service informatique de <b>Lumen Conseil</b> doit créer les comptes de messagerie à partir d'un export RH mal formaté : noms en majuscules, espaces parasites. Construisez des formules qui produisent identifiants et adresses propres.</p>`,
          setup: {
            cols: 11, rows: 10, sheetName: "Comptes",
            data: {
              A1: "Nom (export)", B1: "Prénom", C1: "Entrée", D1: "Nom", E1: "Identifiant", F1: "E-mail", G1: "Année",
              A2: "  LEROY", B2: "Camille", C2: 43539,
              A3: "MARTIN", B3: "Hugo", C3: 44440,
              A4: "GARNIER ", B4: "Inès", C4: 42380,
              A5: "FOURNIER", B5: "Lucas", C5: 45096,
              A6: " MERCIER", B6: "Chloé", C6: 43409,
              A7: "BONNET", B7: "Théo", C7: 43885,
              A8: "LAMBERT ", B8: "Emma", C8: 45390,
              A9: "FAURE", B9: "Nathan", C9: 43038,
              J2: "Entrés avant 2020"
            },
            bold: ["A1", "B1", "C1", "D1", "E1", "F1", "G1"],
            fmt: { C2: "date", C3: "date", C4: "date", C5: "date", C6: "date", C7: "date", C8: "date", C9: "date" },
            widths: { A: 120, F: 220, J: 140 }
          },
          tasks: [
            {
              text: "En <b>D2</b>, obtiens le nom propre sans espaces (<code>Leroy</code>) avec <b>NOMPROPRE</b> et <b>SUPPRESPACE</b>, recopie jusqu'en <b>D9</b>.",
              hint: "=NOMPROPRE(SUPPRESPACE(A2))",
              check: a => a.hasFn("D2", "NOMPROPRE") && a.hasFn("D2", "SUPPRESPACE") && a.val("D2") === "Leroy" && a.val("D4") === "Garnier" && a.val("D6") === "Mercier"
            },
            {
              text: "En <b>E2</b>, construis l'identifiant : initiale du prénom + nom, en minuscules (<code>cleroy</code>). Recopie jusqu'en <b>E9</b>.",
              hint: "=MINUSCULE(GAUCHE(B2;1)&SUPPRESPACE(A2))",
              check: a => a.hasFn("E2", "GAUCHE") && a.val("E2") === "cleroy" && a.val("E4") === "igarnier" && a.val("E9") === "nfaure"
            },
            {
              text: "En <b>F2</b>, construis l'adresse e-mail <code>identifiant@lumen-conseil.fr</code> avec l'opérateur <b>&amp;</b>. Recopie jusqu'en <b>F9</b>.",
              hint: "=E2&\"@lumen-conseil.fr\"",
              check: a => a.f("F2").indexOf("&") >= 0 && a.val("F2") === "cleroy@lumen-conseil.fr" && a.val("F7") === "tbonnet@lumen-conseil.fr"
            },
            {
              text: "En <b>G2</b>, extrais l'année d'entrée avec <b>ANNEE</b>, recopie jusqu'en <b>G9</b>.",
              hint: "=ANNEE(C2)",
              check: a => a.hasFn("G2", "ANNEE") && a.num("G2") === 2019 && a.num("G4") === 2016 && a.num("G9") === 2017
            },
            {
              text: "En <b>K2</b>, compte les collaborateurs entrés <b>avant 2020</b> avec <b>NB.SI</b> sur la colonne G.",
              hint: "=NB.SI(G2:G9;\"<2020\")",
              check: a => a.hasFn("K2", "NB.SI") && a.num("K2") === 4
            }
          ],
          solution: `<ol>
<li>D2 : <code>=NOMPROPRE(SUPPRESPACE(A2))</code> → Leroy, Martin, Garnier…</li>
<li>E2 : <code>=MINUSCULE(GAUCHE(B2;1)&amp;SUPPRESPACE(A2))</code> → cleroy, hmartin, igarnier, lfournier, cmercier, tbonnet, elambert, nfaure.</li>
<li>F2 : <code>=E2&amp;"@lumen-conseil.fr"</code>.</li>
<li>G2 : <code>=ANNEE(C2)</code> → 2019, 2021, 2016, 2023, 2018, 2020, 2024, 2017.</li>
<li>K2 : <code>=NB.SI(G2:G9;"&lt;2020")</code> → 4 (Leroy, Garnier, Mercier, Faure).</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Dans =RECHERCHEV(B2;$I$2:$L$7;3;FAUX), quelle colonne de la feuille est renvoyée ?",
          choices: ["C", "I", "K", "L"],
          answer: 2,
          explain: "Le n° de colonne se compte à partir du début de la table : I = 1, J = 2, K = 3."
        },
        {
          q: "Que se passe-t-il si l'on omet le 4e argument de RECHERCHEV sur une liste de codes non triée ?",
          choices: ["Excel affiche #NOM?", "La recherche est approchée et peut renvoyer un résultat faux sans erreur", "Excel demande confirmation", "La recherche est exacte par défaut"],
          answer: 1,
          explain: "Par défaut, RECHERCHEV utilise la correspondance approchée (VRAI). Sur une liste non triée, elle renvoie un résultat plausible mais faux."
        },
        {
          q: "Quel est l'avantage de RECHERCHEX sur RECHERCHEV ?",
          choices: ["Elle fonctionne sous Excel 2010", "Elle peut chercher une colonne à gauche de la clé et est exacte par défaut", "Elle est plus rapide à taper car sans argument", "Elle trie la table automatiquement"],
          answer: 1,
          explain: "RECHERCHEX travaille avec deux plages indépendantes (recherche et retour), donc dans n'importe quel sens, et sa correspondance est exacte par défaut."
        },
        {
          q: "Que renvoie =EQUIV(\"C003\";I2:I7;0) si I2:I7 contient C001…C006 ?",
          choices: ["C003", "3", "4", "Nantes"],
          answer: 1,
          explain: "EQUIV renvoie la position dans la plage : C003 est la 3e valeur."
        },
        {
          q: "RECHERCHEV renvoie #N/A alors que le code « Lyon » semble bien présent dans la table. Cause probable ?",
          choices: ["La table n'est pas triée", "Un espace invisible en fin de texte (« Lyon  »)", "La colonne est trop étroite", "Il faut écrire LYON en majuscules"],
          answer: 1,
          explain: "Les espaces parasites issus des imports sont la cause n°1. SUPPRESPACE nettoie la valeur. La recherche n'est pas sensible à la casse."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 3 — x06                                                       */
    /* ================================================================== */
    {
      id: "x06",
      level: 3,
      title: "Graphiques et tableaux croisés dynamiques",
      icon: "📊",
      duration: 70,
      summary: "Faire parler les chiffres : choisir le bon graphique et synthétiser des milliers de lignes en trois clics avec un tableau croisé dynamique.",
      objectives: [
        "Être capable de choisir le type de graphique adapté au message",
        "Être capable de créer et personnaliser un graphique (titre, axes, étiquettes)",
        "Être capable de construire un tableau croisé dynamique (Lignes, Colonnes, Valeurs, Filtres)",
        "Être capable de changer la synthèse d'un TCD (somme, nombre, moyenne) et de l'actualiser"
      ],
      lessons: [
        {
          title: "Choisir le bon graphique",
          html: `<table class="o-table"><thead><tr><th>Message</th><th>Graphique</th></tr></thead><tbody>
<tr><td>Évolution dans le temps</td><td>Courbes (ou histogramme pour peu de périodes)</td></tr>
<tr><td>Comparaison entre catégories</td><td>Barres / histogramme groupé, trié</td></tr>
<tr><td>Part d'un tout (≤ 5 parts)</td><td>Secteurs ou anneau</td></tr>
<tr><td>Écart réel / objectif</td><td>Histogramme groupé ou graphique combiné</td></tr>
<tr><td>Contribution à une variation</td><td>Cascade</td></tr>
<tr><td>Relation entre deux mesures</td><td>Nuage de points</td></tr>
</tbody></table>
<div class="o-path">Insertion › Graphiques › Graphiques recommandés</div>
<div class="callout bad"><b>Erreur fréquente :</b> un secteur à 12 parts ou un histogramme 3D. Les angles et la perspective trompent l'œil : restez en 2D, peu de couleurs, une idée par graphique.</div>`
        },
        {
          title: "Créer et soigner un graphique",
          html: `<ol>
<li>Sélectionnez les données <b>avec leurs en-têtes</b> (ex. A1:B13).</li>
<li>Insertion › Graphiques › choisissez le type (ou <kbd>Alt</kbd>+<kbd>F1</kbd> pour le type par défaut).</li>
<li>Onglet <b>Création de graphique</b> : Ajouter un élément graphique (titre, étiquettes, axes), Intervertir les lignes/colonnes, Modifier le type.</li>
<li>Bouton <b>+</b> à droite du graphique pour cocher rapidement les éléments.</li>
</ol>
<div class="o-path">Création de graphique › Ajouter un élément graphique › Étiquettes de données</div>
<div class="callout ok"><b>Bonne pratique :</b> un titre qui dit la conclusion (« Le CA dépasse l'objectif 10 mois sur 12 ») plutôt que la description (« CA 2026 »).</div>
<div class="o-tip">💡 Basé sur un <b>tableau structuré</b>, le graphique intègre automatiquement chaque nouveau mois ajouté. Pour un graphique tenant dans une cellule : Insertion › Graphiques sparkline.</div>`
        },
        {
          title: "Le tableau croisé dynamique (TCD)",
          html: `<p>Un TCD résume une liste (une ligne par transaction) en tableau de synthèse, sans aucune formule.</p>
<div class="o-path">Insertion › Tableaux › Tableau croisé dynamique › À partir d'un tableau ou d'une plage</div>
<p>Dans le volet <b>Champs de tableau croisé dynamique</b>, faites glisser les champs dans 4 zones :</p>
<table class="o-table"><thead><tr><th>Zone</th><th>Exemple</th></tr></thead><tbody>
<tr><td><b>Lignes</b></td><td>Région</td></tr>
<tr><td><b>Colonnes</b></td><td>Produit</td></tr>
<tr><td><b>Valeurs</b></td><td>Somme de CA</td></tr>
<tr><td><b>Filtres</b></td><td>Vendeur</td></tr>
</tbody></table>
<div class="callout ok"><b>Bonne pratique :</b> source propre = une ligne d'en-têtes uniques, aucune ligne ou colonne vide, aucun sous-total dans les données. Mettez-la sous forme de tableau pour que le TCD suive les nouvelles lignes.</div>`
        },
        {
          title: "Paramétrer, actualiser et illustrer un TCD",
          html: `<ul>
<li><b>Synthèse</b> : clic droit sur une valeur › Synthétiser les valeurs par › Somme, Nombre, Moyenne, Max…</li>
<li><b>Afficher les valeurs</b> : % du total général, différence par rapport au mois précédent, rang…</li>
<li><b>Grouper</b> : clic droit sur des dates › Grouper › Mois / Trimestres / Années.</li>
<li><b>Segments</b> et <b>chronologie</b> : Analyse du tableau croisé dynamique › Insérer un segment.</li>
<li><b>Graphique croisé dynamique</b> : Analyse du tableau croisé dynamique › Graphique croisé dynamique.</li>
</ul>
<div class="o-path">Analyse du tableau croisé dynamique › Données › Actualiser</div>
<div class="callout warn"><b>Piège :</b> un TCD ne se met <b>pas</b> à jour automatiquement quand les données changent. Pensez à <b>Actualiser</b> (<kbd>Alt</kbd>+<kbd>F5</kbd>) ou <b>Actualiser tout</b> (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>F5</kbd>).</div>
<div class="o-tip">💡 Double-cliquez sur un total du TCD : Excel crée une nouvelle feuille avec les lignes de détail qui le composent. Idéal pour justifier un chiffre en réunion.</div>`
        }
      ],
      shortcuts: [
        { keys: "Alt+F1", desc: "Graphique par défaut sur la feuille" },
        { keys: "F11", desc: "Graphique sur une nouvelle feuille (dans Excel)" },
        { keys: "Alt+F5", desc: "Actualiser le TCD" },
        { keys: "Ctrl+Alt+F5", desc: "Actualiser toutes les connexions et TCD" },
        { keys: "Ctrl+Maj+L", desc: "Filtre sur la source avant analyse" },
        { keys: "Ctrl+A", desc: "Sélectionner la source du graphique" }
      ],
      labs: [
        {
          id: "x06-lab1",
          title: "Lab 1 — Le reporting mensuel des ventes",
          kind: "sheet",
          scenario: `<p>Directrice commerciale d'<b>Écoverde</b>, fabricant de mobilier de jardin, vous présentez demain le bilan annuel au comité. Il faut le total, le meilleur mois et deux graphiques lisibles.</p>`,
          setup: {
            cols: 6, rows: 16, sheetName: "Mensuel",
            data: {
              A1: "Mois", B1: "CA", C1: "Objectif",
              A2: "Janvier", B2: 42500, C2: 40000,
              A3: "Février", B3: 38900, C3: 40000,
              A4: "Mars", B4: 45200, C4: 42000,
              A5: "Avril", B5: 47800, C5: 42000,
              A6: "Mai", B6: 51300, C6: 45000,
              A7: "Juin", B7: 49700, C7: 45000,
              A8: "Juillet", B8: 36400, C8: 40000,
              A9: "Août", B9: 29800, C9: 35000,
              A10: "Septembre", B10: 46100, C10: 45000,
              A11: "Octobre", B11: 52600, C11: 48000,
              A12: "Novembre", B12: 58900, C12: 50000,
              A13: "Décembre", B13: 63400, C13: 55000,
              A14: "Total", A15: "Meilleur CA", A16: "Meilleur mois"
            },
            bold: ["A1", "B1", "C1", "A14", "A15", "A16"],
            fmt: { B14: "eur", B15: "eur" },
            widths: { A: 120 }
          },
          tasks: [
            {
              text: "En <b>B14</b>, calcule le CA annuel.",
              hint: "=SOMME(B2:B13)",
              check: a => a.hasFn("B14", "SOMME") && a.near(a.num("B14"), 562600)
            },
            {
              text: "En <b>B15</b>, affiche le meilleur CA mensuel avec <b>MAX</b>.",
              hint: "=MAX(B2:B13)",
              check: a => a.hasFn("B15", "MAX") && a.num("B15") === 63400
            },
            {
              text: "En <b>B16</b>, retrouve le nom du meilleur mois avec <b>INDEX</b>/<b>EQUIV</b>.",
              hint: "=INDEX(A2:A13;EQUIV(B15;B2:B13;0))",
              check: a => a.hasFn("B16", "INDEX") && a.hasFn("B16", "EQUIV") && a.val("B16") === "Décembre"
            },
            {
              text: "Sélectionne <b>A1:B13</b> et insère un graphique en <b>courbes</b> pour montrer l'évolution du CA.",
              hint: "Sélection A1:B13 › Graphique › Courbes.",
              check: a => a.charts.some(c => c.type === "line" && rangeStarts(c.range, "A1"))
            },
            {
              text: "Sélectionne <b>A1:C13</b> et insère un graphique en <b>barres</b> comparant CA et objectif.",
              hint: "Sélection A1:C13 › Graphique › Barres.",
              check: a => a.charts.some(c => c.type === "bar" && rangeStarts(c.range, "A1") && /C1[0-9]$/.test(strip(c.range).toUpperCase()))
            }
          ],
          solution: `<ol>
<li>B14 : <code>=SOMME(B2:B13)</code> → 562 600 €.</li>
<li>B15 : <code>=MAX(B2:B13)</code> → 63 400 €.</li>
<li>B16 : <code>=INDEX(A2:A13;EQUIV(B15;B2:B13;0))</code> → Décembre.</li>
<li>A1:B13 › Courbes : on lit la saisonnalité (creux de juillet-août, forte fin d'année).</li>
<li>A1:C13 › Barres : 9 mois sur 12 au-dessus de l'objectif (sauf février, juillet, août).</li>
</ol>`
        },
        {
          id: "x06-lab2",
          title: "Lab 2 — Analyser les ventes avec un TCD",
          kind: "sheet",
          scenario: `<p>Chez <b>Rivière & Glisse</b>, loueur et vendeur de kayaks et paddles, le gérant veut savoir quelle région rapporte le plus et quels produits se vendent le plus souvent. Les 14 dernières ventes sont listées.</p>`,
          setup: {
            cols: 8, rows: 16, sheetName: "Ventes",
            data: {
              A1: "Vendeur", B1: "Région", C1: "Produit", D1: "CA",
              A2: "Alice", B2: "Ouest", C2: "Kayak", D2: 1250,
              A3: "Bruno", B3: "Est", C3: "Paddle", D3: 890,
              A4: "Alice", B4: "Ouest", C4: "Paddle", D4: 640,
              A5: "Chloé", B5: "Sud", C5: "Kayak", D5: 1480,
              A6: "David", B6: "Est", C6: "Kayak", D6: 1320,
              A7: "Bruno", B7: "Est", C7: "Gilet", D7: 210,
              A8: "Chloé", B8: "Sud", C8: "Paddle", D8: 760,
              A9: "Alice", B9: "Ouest", C9: "Gilet", D9: 185,
              A10: "Emma", B10: "Sud", C10: "Kayak", D10: 1390,
              A11: "David", B11: "Est", C11: "Paddle", D11: 920,
              A12: "Emma", B12: "Sud", C12: "Gilet", D12: 240,
              A13: "Alice", B13: "Ouest", C13: "Kayak", D13: 1170,
              A14: "Chloé", B14: "Sud", C14: "Gilet", D14: 160,
              A15: "Bruno", B15: "Est", C15: "Kayak", D15: 1410,
              F1: "Contrôle Est"
            },
            bold: ["A1", "B1", "C1", "D1", "F1"],
            fmt: { D2: "eur", D3: "eur", D4: "eur", D5: "eur", D6: "eur", D7: "eur", D8: "eur", D9: "eur", D10: "eur", D11: "eur", D12: "eur", D13: "eur", D14: "eur", D15: "eur", G1: "eur" }
          },
          tasks: [
            {
              text: "Crée un <b>tableau croisé dynamique</b> avec <b>Région</b> en Lignes et la <b>Somme de CA</b> en Valeurs.",
              hint: "Sélectionne A1:D15 › Tableau croisé dynamique › Lignes : Région, Valeurs : CA, Somme.",
              check: a => a.pivots.some(p => p.rows === "Région" && p.values === "CA" && p.agg === "sum")
            },
            {
              text: "Crée un second TCD qui <b>compte</b> le nombre de ventes par <b>Produit</b>.",
              hint: "Lignes : Produit, Valeurs : CA, synthèse Nombre.",
              check: a => a.pivots.some(p => p.rows === "Produit" && p.agg === "count")
            },
            {
              text: "Crée un troisième TCD donnant le <b>panier moyen</b> (moyenne du CA) par <b>Vendeur</b>.",
              hint: "Lignes : Vendeur, Valeurs : CA, synthèse Moyenne.",
              check: a => a.pivots.some(p => p.rows === "Vendeur" && p.values === "CA" && p.agg === "avg")
            },
            {
              text: "Pour contrôler ton TCD, calcule en <b>G1</b> le CA de la région <b>Est</b> avec <b>SOMME.SI</b>.",
              hint: "=SOMME.SI(B2:B15;\"Est\";D2:D15)",
              check: a => a.hasFn("G1", "SOMME.SI") && a.near(a.num("G1"), 4750)
            }
          ],
          solution: `<ol>
<li>TCD Région / Somme de CA : Est 4 750 €, Sud 4 030 €, Ouest 3 245 € (total 12 025 €).</li>
<li>TCD Produit / Nombre : Kayak 6, Paddle 4, Gilet 4.</li>
<li>TCD Vendeur / Moyenne : Bruno 836,67 €, David 1 120 €, Emma 815 €, Chloé 800 €, Alice 811,25 €.</li>
<li>G1 : <code>=SOMME.SI(B2:B15;"Est";D2:D15)</code> → 4 750 €, identique au TCD.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Quel graphique choisir pour montrer l'évolution du CA sur 12 mois ?",
          choices: ["Secteurs", "Courbes", "Nuage de points", "Anneau"],
          answer: 1,
          explain: "La courbe est le graphique de l'évolution dans le temps : la pente se lit immédiatement."
        },
        {
          q: "Vous avez modifié des montants dans la source d'un TCD. Que faut-il faire ?",
          choices: ["Rien, il se met à jour seul", "Actualiser le TCD", "Recréer le TCD", "Appuyer sur F2"],
          answer: 1,
          explain: "Un TCD travaille sur un cache : il faut l'actualiser (Analyse du tableau croisé dynamique › Actualiser ou Alt+F5)."
        },
        {
          q: "Dans quelle zone du TCD placer le champ CA pour obtenir des totaux ?",
          choices: ["Lignes", "Colonnes", "Valeurs", "Filtres"],
          answer: 2,
          explain: "Les champs numériques à agréger vont dans Valeurs ; Lignes et Colonnes servent aux axes d'analyse."
        },
        {
          q: "Que se passe-t-il si vous double-cliquez sur un total dans un TCD ?",
          choices: ["Le TCD est supprimé", "Une feuille avec le détail des lignes concernées est créée", "La cellule passe en mode modification", "Le graphique se met à jour"],
          answer: 1,
          explain: "C'est l'« extraction » : Excel génère une nouvelle feuille contenant les enregistrements source de ce total."
        },
        {
          q: "Pourquoi éviter un graphique en secteurs à 12 parts ?",
          choices: ["Excel ne le permet pas", "Les angles proches sont impossibles à comparer à l'œil", "Il ne s'imprime pas en couleur", "Il ne gère pas les nombres décimaux"],
          answer: 1,
          explain: "L'œil compare mal les angles : au-delà de 4-5 parts, un histogramme trié est bien plus lisible."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 4 — x07                                                       */
    /* ================================================================== */
    {
      id: "x07",
      level: 4,
      title: "Fonctions multi-critères, SOMMEPROD et tableau de bord",
      icon: "🎛️",
      duration: 80,
      summary: "Construire un tableau de bord piloté par un sélecteur : SOMME.SI.ENS, NB.SI.ENS, MAX.SI.ENS, SI.CONDITIONS et SOMMEPROD.",
      objectives: [
        "Être capable d'agréger selon plusieurs critères avec les fonctions .SI.ENS",
        "Être capable d'utiliser SOMMEPROD pour des calculs pondérés sans colonne intermédiaire",
        "Être capable de classer des valeurs par paliers avec SI.CONDITIONS",
        "Être capable de concevoir un tableau de bord dynamique piloté par une cellule de sélection"
      ],
      lessons: [
        {
          title: "Les fonctions .SI.ENS",
          html: `<p>Les fonctions .SI.ENS acceptent jusqu'à 127 couples (plage ; critère). Attention : la plage de <b>calcul</b> vient en <b>premier</b>, contrairement à SOMME.SI.</p>
<pre class="code">=SOMME.SI.ENS(F2:F13;B2:B13;"Nord";C2:C13;"Logiciel")
=NB.SI.ENS(B2:B13;J1;F2:F13;">=3000")
=MOYENNE.SI.ENS(F2:F13;B2:B13;J1)
=MAX.SI.ENS(F2:F13;B2:B13;J1)
=MIN.SI.ENS(F2:F13;B2:B13;J1)</pre>
<p>Tous les critères sont combinés en <b>ET</b>. Pour un OU (Nord <i>ou</i> Est), additionnez deux SOMME.SI.ENS.</p>
<div class="callout warn"><b>Piège :</b> SOMME.SI(plage_critère ; critère ; plage_somme) mais SOMME.SI.ENS(plage_somme ; plage_critère ; critère…). Inverser provoque #VALEUR! ou un résultat absurde. Astuce : utilisez toujours .SI.ENS, même pour un seul critère.</div>
<div class="o-tip">💡 Dates entre deux bornes : <code>=SOMME.SI.ENS(F:F;A:A;"&gt;="&amp;J2;A:A;"&lt;="&amp;J3)</code>.</div>`
        },
        {
          title: "SOMMEPROD : le couteau suisse",
          html: `<p>SOMMEPROD multiplie les plages terme à terme puis additionne :</p>
<pre class="code">=SOMMEPROD(D2:D13;E2:E13)       = D2*E2 + D3*E3 + … + D13*E13</pre>
<p>C'est le CA total sans colonne intermédiaire. Avec des tests logiques, elle devient un SOMME.SI.ENS ultra-souple (conditions OU, calculs sur les colonnes, fonctions sur les dates) :</p>
<pre class="code">=SOMMEPROD((B2:B13="Nord")*(C2:C13="Logiciel")*D2:D13*E2:E13)
=SOMMEPROD(((B2:B13="Nord")+(B2:B13="Est")>0)*F2:F13)     Nord OU Est
=SOMMEPROD((MOIS(A2:A13)=3)*F2:F13)                          ventes de mars</pre>
<p>Chaque test renvoie VRAI/FAUX, que la multiplication convertit en 1/0.</p>
<div class="callout warn"><b>Piège :</b> toutes les plages doivent avoir exactement la même taille, sinon #VALEUR!. Évitez les colonnes entières (A:A) : SOMMEPROD calculerait un million de lignes.</div>`
        },
        {
          title: "Classer par paliers : SI.CONDITIONS et recherche approchée",
          html: `<pre class="code">=SI.CONDITIONS(F2>=5000;"A";F2>=3000;"B";VRAI;"C")</pre>
<p>Les conditions sont testées dans l'ordre ; <code>VRAI</code> en dernier joue le rôle de « sinon ». Sans lui, une valeur qui ne remplit aucune condition renvoie #N/A.</p>
<p>Quand les paliers changent souvent, sortez-les dans une <b>table de barème</b> triée et utilisez une recherche approchée :</p>
<pre class="code">=RECHERCHEX(F2;$M$2:$M$4;$N$2:$N$4;;-1)     -1 = valeur exacte ou inférieure
=RECHERCHEV(F2;$M$2:$N$4;2;VRAI)            barème trié par ordre croissant</pre>
<div class="callout ok"><b>Bonne pratique :</b> les règles métier (seuils, taux) dans des cellules paramètres, jamais en dur dans les formules. Le métier modifie la règle sans toucher aux formules.</div>`
        },
        {
          title: "Concevoir un tableau de bord",
          html: `<ol>
<li><b>Séparer</b> : une feuille Données (tableau structuré), une feuille Calculs, une feuille Tableau de bord.</li>
<li><b>Un sélecteur</b> : une cellule avec une liste déroulante qui pilote tous les indicateurs.</li>
<li><b>Des KPI</b> calculés avec .SI.ENS / SOMMEPROD qui pointent vers le sélecteur.</li>
<li><b>Visualiser</b> : graphiques, MFC, sparklines ; segments si l'on passe par des TCD.</li>
<li><b>Protéger</b> : Révision › Protéger la feuille, en déverrouillant seulement le sélecteur.</li>
</ol>
<div class="o-path">Données › Outils de données › Validation des données › Autoriser : Liste</div>
<div class="o-path">Formules › Noms définis › Gestionnaire de noms</div>
<div class="o-tip">💡 Nommez la cellule sélecteur <code>RegionChoisie</code> : <code>=SOMME.SI.ENS(Ventes[CA];Ventes[Région];RegionChoisie)</code> se lit comme une phrase.</div>
<div class="callout bad"><b>Erreur fréquente :</b> copier-coller les chiffres chaque mois dans le tableau de bord. Un vrai tableau de bord ne contient <b>que des formules</b> : on remplace les données, tout se recalcule.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+F3", desc: "Gestionnaire de noms" },
        { keys: "Ctrl+Maj+F3", desc: "Créer des noms à partir de la sélection" },
        { keys: "F3", desc: "Coller un nom dans la formule" },
        { keys: "Alt+Flèche bas", desc: "Ouvrir une liste déroulante de validation" },
        { keys: "Ctrl+Maj+U", desc: "Agrandir la barre de formule" },
        { keys: "Ctrl+`", desc: "Afficher les formules au lieu des résultats" },
        { keys: "F9", desc: "Évaluer une portion de formule sélectionnée" }
      ],
      labs: [
        {
          id: "x07-lab1",
          title: "Lab 1 — Le tableau de bord commercial par région",
          kind: "sheet",
          scenario: `<p>Vous êtes analyste chez <b>Datacomp</b>, éditeur de logiciels qui vend aussi du matériel et des formations. Le directeur commercial veut un mini tableau de bord : il tape une région en <b>J1</b> et tous les indicateurs s'actualisent.</p>`,
          setup: {
            cols: 10, rows: 14, sheetName: "Tableau de bord",
            data: {
              A1: "Commercial", B1: "Région", C1: "Gamme", D1: "Qté", E1: "Prix unit.", F1: "CA", G1: "Niveau",
              A2: "Lemoine", B2: "Nord", C2: "Logiciel", D2: 3, E2: 1200,
              A3: "Garcia", B3: "Sud", C3: "Matériel", D3: 10, E3: 340,
              A4: "Lemoine", B4: "Nord", C4: "Formation", D4: 2, E4: 950,
              A5: "Roux", B5: "Est", C5: "Logiciel", D5: 5, E5: 1200,
              A6: "Garcia", B6: "Sud", C6: "Logiciel", D6: 2, E6: 1200,
              A7: "Fontaine", B7: "Nord", C7: "Matériel", D7: 14, E7: 340,
              A8: "Roux", B8: "Est", C8: "Formation", D8: 4, E8: 950,
              A9: "Chevalier", B9: "Sud", C9: "Formation", D9: 6, E9: 950,
              A10: "Fontaine", B10: "Nord", C10: "Logiciel", D10: 4, E10: 1200,
              A11: "Chevalier", B11: "Sud", C11: "Matériel", D11: 8, E11: 340,
              A12: "Roux", B12: "Est", C12: "Matériel", D12: 12, E12: 340,
              A13: "Lemoine", B13: "Nord", C13: "Formation", D13: 1, E13: 950,
              I1: "Région choisie", J1: "Nord",
              I2: "CA total (toutes régions)", I3: "CA de la région", I4: "CA région · Logiciel",
              I5: "Ventes ≥ 3 000 € (région)", I6: "Meilleure vente (région)"
            },
            bold: ["A1", "B1", "C1", "D1", "E1", "F1", "G1", "I1", "J1"],
            fmt: { E2: "eur", J2: "eur", J3: "eur", J4: "eur", J6: "eur" },
            widths: { A: 100, I: 190 }
          },
          tasks: [
            {
              text: "En <b>F2</b>, calcule le CA de la ligne (<code>=D2*E2</code>) et recopie jusqu'en <b>F13</b>.",
              hint: "=D2*E2 puis Ctrl+D sur F2:F13.",
              check: a => a.isFormula("F2") && a.isFormula("F13") && a.near(a.num("F2"), 3600) && a.near(a.num("F7"), 4760) && a.near(a.num("F13"), 950)
            },
            {
              text: "En <b>G2</b>, classe chaque vente avec <b>SI.CONDITIONS</b> : <code>\"A\"</code> si CA ≥ 5 000, <code>\"B\"</code> si CA ≥ 3 000, sinon <code>\"C\"</code>. Recopie jusqu'en <b>G13</b>.",
              hint: "=SI.CONDITIONS(F2>=5000;\"A\";F2>=3000;\"B\";VRAI;\"C\") — si VRAI n'est pas accepté, termine par F2>=0;\"C\".",
              check: a => a.hasFn("G2", "SI.CONDITIONS") && a.val("G2") === "B" && a.val("G4") === "C" && a.val("G5") === "A" && a.val("G13") === "C"
            },
            {
              text: "En <b>J2</b>, calcule le CA total <b>sans utiliser la colonne F</b>, avec <b>SOMMEPROD</b> sur les quantités et les prix.",
              hint: "=SOMMEPROD(D2:D13;E2:E13)",
              check: a => a.hasFn("J2", "SOMMEPROD") && a.f("J2").indexOf("F") < 0 && a.near(a.num("J2"), 44110)
            },
            {
              text: "En <b>J3</b>, calcule le CA de la région choisie (critère = cellule <b>J1</b>) avec <b>SOMME.SI.ENS</b> ; en <b>J4</b>, le CA de la région choisie pour la gamme <code>\"Logiciel\"</code> (2 critères).",
              hint: "J3 : =SOMME.SI.ENS(F2:F13;B2:B13;J1) · J4 : =SOMME.SI.ENS(F2:F13;B2:B13;J1;C2:C13;\"Logiciel\")",
              check: a => {
                const e = agg(a, a.val("J1"), "Logiciel");
                return a.hasFn("J3", "SOMME.SI.ENS") && strip(a.f("J3")).indexOf("J1") >= 0 && a.near(a.num("J3"), e.sum) &&
                  a.hasFn("J4", "SOMME.SI.ENS") && a.near(a.num("J4"), e.sumG) && e.sum > 0;
              }
            },
            {
              text: "En <b>J5</b>, compte les ventes de la région choisie d'au moins 3 000 € avec <b>NB.SI.ENS</b> ; en <b>J6</b>, affiche sa meilleure vente avec <b>MAX.SI.ENS</b>.",
              hint: "J5 : =NB.SI.ENS(B2:B13;J1;F2:F13;\">=3000\") · J6 : =MAX.SI.ENS(F2:F13;B2:B13;J1)",
              check: a => {
                const e = agg(a, a.val("J1"), "Logiciel");
                return a.hasFn("J5", "NB.SI.ENS") && a.num("J5") === e.n3000 && a.hasFn("J6", "MAX.SI.ENS") && a.near(a.num("J6"), e.max) && e.sum > 0;
              }
            },
            {
              text: "Teste ton tableau de bord : remplace <b>Nord</b> par <b>Sud</b> en J1. Tous les indicateurs doivent s'actualiser.",
              hint: "Sud : CA 14 220 €, Logiciel 2 400 €, 2 ventes ≥ 3 000 €, meilleure vente 5 700 €.",
              check: a => same(a.val("J1"), "Sud") && a.near(a.num("J3"), 14220) && a.near(a.num("J4"), 2400) && a.num("J5") === 2 && a.near(a.num("J6"), 5700)
            }
          ],
          solution: `<ol>
<li>F2 : <code>=D2*E2</code> → 3 600, 3 400, 1 900, 6 000, 2 400, 4 760, 3 800, 5 700, 4 800, 2 720, 4 080, 950.</li>
<li>G2 : <code>=SI.CONDITIONS(F2>=5000;"A";F2>=3000;"B";VRAI;"C")</code> → deux « A » (6 000 et 5 700).</li>
<li>J2 : <code>=SOMMEPROD(D2:D13;E2:E13)</code> → 44 110 €.</li>
<li>J3 : <code>=SOMME.SI.ENS(F2:F13;B2:B13;J1)</code> → Nord 16 010 € ; J4 : <code>=SOMME.SI.ENS(F2:F13;B2:B13;J1;C2:C13;"Logiciel")</code> → 8 400 €.</li>
<li>J5 : <code>=NB.SI.ENS(B2:B13;J1;F2:F13;"&gt;=3000")</code> → 3 ; J6 : <code>=MAX.SI.ENS(F2:F13;B2:B13;J1)</code> → 4 800 €.</li>
<li>J1 = Sud : 14 220 € · 2 400 € · 2 · 5 700 €. (Est : 13 880 € · 6 000 € · 3 · 6 000 €.)</li>
</ol>
<p>Dans le vrai Excel, ajoutez en J1 une liste déroulante (Données › Validation des données › Liste : <code>Nord;Sud;Est</code>).</p>`
        }
      ],
      quiz: [
        {
          q: "Quel est l'ordre des arguments de SOMME.SI.ENS ?",
          choices: ["plage_critère ; critère ; plage_somme", "plage_somme ; plage_critère1 ; critère1 ; …", "critère ; plage_somme ; plage_critère", "plage_somme ; critère1 ; plage_critère1"],
          answer: 1,
          explain: "Avec .SI.ENS, la plage à additionner vient en premier, suivie des couples (plage ; critère)."
        },
        {
          q: "Que calcule =SOMMEPROD(D2:D4;E2:E4) si D = 2, 3, 1 et E = 10, 20, 5 ?",
          choices: ["41", "85", "6", "35"],
          answer: 1,
          explain: "2×10 + 3×20 + 1×5 = 20 + 60 + 5 = 85."
        },
        {
          q: "Dans =SI.CONDITIONS(F2>=5000;\"A\";F2>=3000;\"B\"), que renvoie Excel si F2 = 1 000 ?",
          choices: ["\"B\"", "0", "#N/A", "Une cellule vide"],
          answer: 2,
          explain: "Aucune condition n'est vraie et il n'y a pas de cas par défaut : #N/A. On ajoute VRAI;\"C\" en dernier couple."
        },
        {
          q: "Comment additionner le CA des régions Nord OU Est avec les fonctions .SI.ENS ?",
          choices: ["=SOMME.SI.ENS(F:F;B:B;\"Nord\";B:B;\"Est\")", "Additionner deux SOMME.SI.ENS, une par région", "=SOMME.SI.ENS(F:F;B:B;\"Nord ou Est\")", "C'est impossible"],
          answer: 1,
          explain: "Les critères d'une même SOMME.SI.ENS sont combinés en ET (une ligne ne peut pas être Nord ET Est). Pour un OU, on additionne deux formules, ou on passe par SOMMEPROD."
        },
        {
          q: "Quelle est la meilleure pratique pour un tableau de bord mensuel ?",
          choices: ["Coller les chiffres en valeurs chaque mois", "Des KPI en formules pointant vers les données et un sélecteur", "Un onglet par mois copié de l'onglet précédent", "Saisir les totaux à la main pour éviter les erreurs"],
          answer: 1,
          explain: "Un tableau de bord ne contient que des formules : on met à jour la source, tout se recalcule. Le sélecteur rend l'analyse interactive."
        }
      ]
    },

    /* ================================================================== */
    /* NIVEAU 4 — x08                                                       */
    /* ================================================================== */
    {
      id: "x08",
      level: 4,
      title: "Automatiser : Power Query, matrices dynamiques et macros VBA",
      icon: "🤖",
      duration: 90,
      summary: "Ne plus jamais refaire à la main : importer et nettoyer avec Power Query, calculer avec FILTRE, UNIQUE et LET, automatiser avec une macro VBA.",
      objectives: [
        "Être capable d'importer et de nettoyer un fichier CSV avec Power Query, puis de l'actualiser",
        "Être capable d'utiliser les fonctions matricielles dynamiques FILTRE, UNIQUE, TRIER et LET",
        "Être capable d'enregistrer, lire et modifier une macro VBA simple",
        "Être capable d'enregistrer un classeur prenant en charge les macros en connaissant les règles de sécurité"
      ],
      lessons: [
        {
          title: "Power Query : importer et transformer",
          html: `<p>Power Query est l'outil d'<b>ETL</b> d'Excel : il importe une source (CSV, dossier, base, web), applique des étapes de nettoyage <b>enregistrées</b>, et recharge le tout d'un clic.</p>
<div class="o-path">Données › Récupérer et transformer des données › À partir d'un fichier texte/CSV</div>
<p>Cliquez sur <b>Transformer les données</b> pour ouvrir l'éditeur Power Query. Transformations courantes :</p>
<ul>
<li><b>Accueil › Utiliser la première ligne pour les en-têtes</b></li>
<li><b>Transformer › Type de données</b> (Nombre décimal, Date…) — vérifiez les paramètres régionaux (Utilisation des paramètres régionaux… › Français)</li>
<li><b>Transformer › Format › Supprimer les espaces</b> et <b>Majuscule à chaque mot</b></li>
<li><b>Accueil › Supprimer les lignes › Supprimer les doublons / les lignes vides</b></li>
<li><b>Ajouter une colonne › Colonne personnalisée</b> : <code>[Quantité] * [Prix]</code></li>
</ul>
<p>Chaque action s'ajoute dans <b>Étapes appliquées</b> (volet de droite) : on peut la renommer, la modifier, la supprimer. Terminez par <b>Accueil › Fermer et charger</b>.</p>
<div class="o-path">Données › Requêtes et connexions › Actualiser tout</div>
<div class="callout ok"><b>Bonne pratique :</b> le mois prochain, remplacez le fichier source (même nom, même structure) et cliquez sur Actualiser tout : tout le nettoyage est rejoué en une seconde.</div>`
        },
        {
          title: "Les matrices dynamiques : FILTRE, UNIQUE, TRIER",
          html: `<p>Depuis Excel 365 / 2021, une formule peut renvoyer <b>plusieurs valeurs</b> qui « débordent » (<i>spill</i>) dans les cellules voisines.</p>
<pre class="code">=UNIQUE(B2:B50)                               liste sans doublons
=TRIER(UNIQUE(B2:B50))                        triée A→Z
=FILTRE(A2:F50;C2:C50="Nord";"Aucun résultat")
=FILTRE(A2:F50;(C2:C50="Nord")*(F2:F50>1000))     ET
=TRIER(FILTRE(A2:F50;F2:F50>1000);6;-1)           trié par CA décroissant
=NB(UNIQUE(A2:A50))                               nombre de clients distincts</pre>
<p>Référencez une plage débordante avec <code>#</code> : <code>=H2#</code> désigne tout le résultat qui part de H2, quelle que soit sa taille.</p>
<div class="callout warn"><b>Piège :</b> <code>#PROPAGATION!</code> signifie que la zone de débordement n'est pas vide (une valeur bloque). Libérez les cellules en dessous ou à droite.</div>
<div class="o-tip">💡 Une liste déroulante alimentée par <code>=H2#</code> (Validation des données › Liste › Source : <code>=$H$2#</code>) s'enrichit automatiquement de chaque nouvelle région.</div>`
        },
        {
          title: "LET et LAMBDA : des formules lisibles",
          html: `<p><b>LET</b> nomme des calculs intermédiaires à l'intérieur d'une formule : plus lisible, et chaque calcul n'est effectué qu'une fois.</p>
<pre class="code">=LET(
  ca; D2:D13*E2:E13;
  seuil; 3000;
  SOMME(FILTRE(ca; ca>=seuil; 0))
)</pre>
<p><b>LAMBDA</b> va plus loin : créer sa propre fonction, enregistrée dans le Gestionnaire de noms.</p>
<pre class="code">Nom : TTC
Fait référence à : =LAMBDA(ht;taux; ht*(1+taux))
Utilisation : =TTC(B2;$H$1)</pre>
<div class="o-path">Formules › Noms définis › Définir un nom</div>
<div class="callout warn"><b>Compatibilité :</b> FILTRE, UNIQUE, LET et LAMBDA n'existent pas dans Excel 2019 et antérieur : vos destinataires verront <code>#NOM?</code>. Vérifiez la version de vos lecteurs avant de diffuser.</div>`
        },
        {
          title: "Macros VBA : enregistrer, lire, modifier",
          html: `<p>Commencez par afficher l'onglet Développeur :</p>
<div class="o-path">Fichier › Options › Personnaliser le ruban › cocher Développeur</div>
<ol>
<li><b>Développeur › Code › Enregistrer une macro</b> : nommez-la (sans espace), choisissez « Ce classeur ».</li>
<li>Effectuez les actions (mise en forme, tri…), puis <b>Arrêter l'enregistrement</b>.</li>
<li><b>Développeur › Visual Basic</b> (<kbd>Alt</kbd>+<kbd>F11</kbd>) : le code est dans <i>Modules › Module1</i>.</li>
</ol>
<pre class="code">Sub MiseEnFormeRapport()
    ' Met en forme l'en-tête et ajuste les colonnes
    With Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
    End With
    Columns("A:F").AutoFit
    Range("F2:F" &amp; Cells(Rows.Count, "A").End(xlUp).Row).NumberFormat = "#,##0.00 €"
    MsgBox "Rapport mis en forme : " &amp; Cells(Rows.Count, "A").End(xlUp).Row - 1 &amp; " lignes."
End Sub</pre>
<div class="callout warn"><b>Piège :</b> une macro ne s'annule pas avec <kbd>Ctrl</kbd>+<kbd>Z</kbd>. Testez-la sur une copie du fichier. Enregistrez au format <b>Classeur Excel (prenant en charge les macros) *.xlsm</b> : en .xlsx, le code est supprimé.</div>
<div class="callout bad"><b>Sécurité :</b> n'activez jamais les macros d'un fichier reçu par e-mail dont vous n'êtes pas sûr. Les macros sont le premier vecteur de logiciels malveillants dans Office.</div>
<div class="o-tip">💡 Le code enregistré contient beaucoup de <code>.Select</code> inutiles : supprimez-les et travaillez directement sur les objets (<code>Range("A1").Font.Bold = True</code>). Le code devient plus rapide et plus robuste.</div>`
        }
      ],
      shortcuts: [
        { keys: "Alt+F11", desc: "Ouvrir l'éditeur Visual Basic" },
        { keys: "Alt+F8", desc: "Afficher la liste des macros" },
        { keys: "F5", desc: "Exécuter la macro (dans l'éditeur VBA)" },
        { keys: "F8", desc: "Exécuter pas à pas (dans l'éditeur VBA)" },
        { keys: "Ctrl+Alt+F5", desc: "Actualiser toutes les requêtes Power Query" },
        { keys: "Ctrl+Maj+U", desc: "Agrandir la barre de formule pour une formule LET" }
      ],
      labs: [
        {
          id: "x08-lab1",
          title: "Lab 1 — Industrialiser l'import des ventes avec Power Query et les matrices dynamiques",
          kind: "steps",
          scenario: `<p>Chaque mois, le logiciel de caisse de <b>Maison Fournil</b> (5 boulangeries) exporte un CSV « sale » : espaces parasites, casse irrégulière, doublons, lignes vides. Vous allez construire une requête Power Query réutilisable, puis une analyse en formules dynamiques. <b>À réaliser dans Excel 365.</b></p>`,
          setup: {
            app: "Excel",
            file: {
              name: "ventes_fournil.csv",
              type: "csv",
              content: "Date;Boutique;Produit;Quantité;Prix unitaire\n02/09/2026;  Lyon Centre;baguette tradition;320;1,20\n02/09/2026;Villeurbanne;CROISSANT;180;1,10\n02/09/2026;Lyon Centre;pain au chocolat;210;1,25\n03/09/2026;Caluire ;Baguette Tradition;150;1,20\n03/09/2026;Bron;tarte aux pommes;12;18,50\n02/09/2026;Villeurbanne;CROISSANT;180;1,10\n;;;;\n04/09/2026;Écully;pain de campagne;45;3,90\n04/09/2026;Lyon Centre;tarte aux pommes;9;18,50\n04/09/2026;bron;croissant;140;1,10\n05/09/2026;Caluire;pain au chocolat;95;1,25\n05/09/2026;Écully;baguette tradition;175;1,20\n05/09/2026;Villeurbanne;Pain de campagne;38;3,90\n"
            }
          },
          tasks: [
            { text: "Télécharge <b>ventes_fournil.csv</b> puis importe-le : <b>Données › À partir d'un fichier texte/CSV</b> › <b>Transformer les données</b>." },
            { text: "Dans l'éditeur Power Query : sur Boutique et Produit, <b>Transformer › Format › Supprimer les espaces</b> puis <b>Majuscule à chaque mot</b>." },
            { text: "Supprime les <b>lignes vides</b> et les <b>doublons</b> (Accueil › Supprimer les lignes). Vérifie que les types sont corrects (Date, Nombre entier, Nombre décimal avec paramètres régionaux Français)." },
            { text: "Ajoute une <b>colonne personnalisée</b> CA = <code>[Quantité] * [Prix unitaire]</code>, puis <b>Fermer et charger</b> dans un tableau nommé <code>tblVentes</code>." },
            { text: "À côté, écris <code>=TRIER(UNIQUE(tblVentes[Boutique]))</code> puis, à droite de ce résultat, <code>=SOMME.SI.ENS(tblVentes[CA];tblVentes[Boutique];H2#)</code>." },
            { text: "Écris une formule <b>LET</b> + <b>FILTRE</b> qui liste les ventes de plus de 150 €, triées par CA décroissant. Modifie une ligne du CSV, enregistre-le et clique sur <b>Actualiser tout</b>." }
          ],
          solution: `<ol>
<li>L'aperçu détecte le séparateur « ; ». Si les décimales sont mal lues, Type de données › Utilisation des paramètres régionaux… › Décimal / Français (France).</li>
<li>Après nettoyage : boutiques « Lyon Centre », « Villeurbanne », « Caluire », « Bron », « Écully » ; produits homogènes (« Croissant », « Baguette Tradition »…).</li>
<li>Supprimer les lignes vides retire la ligne « ;;;; » ; Supprimer les doublons (toutes les colonnes sélectionnées) retire la 2e ligne « 02/09 · Villeurbanne · Croissant · 180 ». Résultat : <b>11 lignes</b>.</li>
<li>Colonne CA : <code>[Quantité] * [Prix unitaire]</code>, type Nombre décimal. Fermer et charger, puis Création de tableau › Nom du tableau : tblVentes.</li>
<li>Exemple en H2 : <code>=TRIER(UNIQUE(tblVentes[Boutique]))</code> ; en I2 : <code>=SOMME.SI.ENS(tblVentes[CA];tblVentes[Boutique];H2#)</code> → Lyon Centre 384 + 262,50 + 166,50 = 813 €.</li>
<li><pre class="code">=LET(t; tblVentes; ca; tblVentes[CA];
  TRIER(FILTRE(t; ca>150; "Aucune"); 6; -1))</pre> Après modification du CSV, Actualiser tout rejoue toutes les étapes : les formules dynamiques suivent automatiquement.</li>
</ol>`
        },
        {
          id: "x08-lab2",
          title: "Lab 2 — Votre première macro VBA : le rapport en un clic",
          kind: "steps",
          scenario: `<p>Le contrôleur de gestion de <b>Transports Delmas</b> reçoit chaque lundi la même extraction de tournées et passe 15 minutes à la mettre en forme. Vous allez automatiser cette routine. <b>À réaliser dans Excel pour Windows ou Mac (onglet Développeur).</b></p>`,
          setup: {
            app: "Excel",
            file: {
              name: "tournees_semaine.csv",
              type: "csv",
              content: "Tournée;Chauffeur;Départ;Km;Colis;Carburant (L)\nT01;Nadia Benali;Lyon;182;64;21,4\nT02;Marc Olivier;Lyon;145;52;17,9\nT03;Sami Haddad;Vienne;210;71;25,8\nT04;Julie Morel;Villefranche;168;58;19,6\nT05;Paul Girard;Lyon;96;40;11,8\nT06;Léa Fontaine;Bourgoin;233;77;27,9\nT07;Karim Saïdi;Vienne;158;49;18,7\nT08;Anne Roche;Lyon;121;45;14,2\nT09;Hugo Perrin;Villefranche;187;62;22,5\nT10;Inès Vidal;Bourgoin;204;69;24,1\n"
            }
          },
          tasks: [
            { text: "Ouvre <b>tournees_semaine.csv</b> dans Excel et affiche l'onglet <b>Développeur</b> (Fichier › Options › Personnaliser le ruban)." },
            { text: "<b>Développeur › Enregistrer une macro</b> nommée <code>MiseEnFormeTournees</code> : mets l'en-tête en gras sur fond bleu, ajuste les colonnes, applique le format nombre à 1 décimale à la colonne F, puis <b>Arrêter l'enregistrement</b>." },
            { text: "Ouvre l'éditeur (<kbd>Alt</kbd>+<kbd>F11</kbd>), lis le code généré dans Module1 et supprime les lignes <code>.Select</code> / <code>Selection</code> inutiles." },
            { text: "Ajoute à la fin de la macro le calcul de la consommation aux 100 km en colonne G pour chaque ligne (boucle <code>For</code>), puis un <code>MsgBox</code> qui affiche le total des km." },
            { text: "Teste la macro sur un fichier fraîchement ouvert (<kbd>Alt</kbd>+<kbd>F8</kbd> › Exécuter), puis enregistre le classeur au format <b>.xlsm</b>." }
          ],
          solution: `<pre class="code">Sub MiseEnFormeTournees()
    Dim derniere As Long, i As Long, totalKm As Double
    derniere = Cells(Rows.Count, "A").End(xlUp).Row

    With Range("A1:G1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(31, 78, 121)
    End With
    Range("G1").Value = "L/100 km"

    For i = 2 To derniere
        Cells(i, "G").Value = Cells(i, "F").Value / Cells(i, "D").Value * 100
        totalKm = totalKm + Cells(i, "D").Value
    Next i

    Range("F2:G" &amp; derniere).NumberFormat = "0.0"
    Columns("A:G").AutoFit
    MsgBox "Total de la semaine : " &amp; totalKm &amp; " km"
End Sub</pre>
<p>Résultat attendu : total <b>1 704 km</b> ; consommation de T01 = 21,4 / 182 × 100 ≈ 11,8 L/100 km. Le format <code>"0.0"</code> s'écrit à l'anglaise en VBA (point décimal), Excel l'affiche avec une virgule.</p>`
        }
      ],
      quiz: [
        {
          q: "Le mois suivant, le CSV source a été remplacé par un nouveau (même structure). Que faire pour obtenir les données nettoyées ?",
          choices: ["Refaire toutes les étapes Power Query", "Données › Actualiser tout", "Copier-coller le nouveau CSV dans la feuille", "Recréer le tableau structuré"],
          answer: 1,
          explain: "Les étapes appliquées sont mémorisées : Actualiser tout les rejoue sur la nouvelle source."
        },
        {
          q: "Une formule =UNIQUE(B2:B50) affiche #PROPAGATION!. Pourquoi ?",
          choices: ["La plage contient des doublons", "Des cellules non vides bloquent la zone de débordement", "UNIQUE n'existe pas en français", "La plage n'est pas triée"],
          answer: 1,
          explain: "Le résultat a besoin de cellules libres pour se déverser ; une valeur présente dans la zone provoque #PROPAGATION!."
        },
        {
          q: "Quel est l'intérêt principal de LET ?",
          choices: ["Créer une macro sans VBA", "Nommer des calculs intermédiaires pour une formule plus lisible et calculée une seule fois", "Protéger une cellule", "Rendre une formule compatible avec Excel 2010"],
          answer: 1,
          explain: "LET déclare des variables locales à la formule : lisibilité accrue et pas de calcul répété."
        },
        {
          q: "Dans quel format enregistrer un classeur contenant une macro ?",
          choices: [".xlsx", ".xlsm", ".csv", ".xltx"],
          answer: 1,
          explain: "Le format .xlsm (Classeur Excel prenant en charge les macros) conserve le code VBA ; un .xlsx le supprime."
        },
        {
          q: "Vous lancez une macro qui efface une colonne par erreur. Ctrl+Z…",
          choices: ["Annule la macro entière", "Annule uniquement la dernière ligne de code", "Ne fonctionne pas : les actions d'une macro ne s'annulent pas", "Supprime la macro"],
          answer: 2,
          explain: "L'exécution d'une macro vide la pile d'annulation. D'où la règle : toujours tester sur une copie."
        }
      ]
    }
  ]);
})();
