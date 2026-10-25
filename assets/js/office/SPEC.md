# Office Academy — contrat de contenu (modules, labs, quiz)

Site statique (GitHub Pages + file://), FR, aucune dépendance. Page : `pages/office.html`.
Moteur : `assets/js/office/engine.js` (expose `window.OFFICE`). Contenu : un fichier par application :

- `assets/js/office/word.js`        → `OFFICE.add("word", [ ...modules ])`
- `assets/js/office/excel.js`       → `OFFICE.add("excel", [ ...modules ])`
- `assets/js/office/powerpoint.js`  → `OFFICE.add("powerpoint", [ ...modules ])`

JavaScript plain (ES2020, pas de module/import), fichier auto-exécuté :
```js
/* Office Academy — Excel : 8 modules du débutant à l'expert */
OFFICE.add("excel", [ { ...module }, ... ]);
```

## Module
```js
{
  id: "x01",                 // préfixe w (Word) / x (Excel) / p (PowerPoint) + 2 chiffres, unique
  level: 1,                  // 1 Débutant · 2 Intermédiaire · 3 Avancé · 4 Expert  (2 modules par niveau, 8 au total)
  title: "Premiers pas : l'interface et la saisie",
  icon: "🧭",                // un emoji
  duration: 45,              // minutes estimées (lecture + lab + quiz)
  summary: "Une phrase d'accroche.",
  objectives: ["…", "…", "…"],               // 3 à 5 objectifs « Être capable de … »
  lessons: [ { title: "…", html: "…" } ],    // 3 à 5 leçons (chaque leçon = 1 « slide » en mode formateur)
  shortcuts: [ { keys: "Ctrl+Maj+L", desc: "Activer le filtre" } ],  // 4 à 8 raccourcis (Windows ; le moteur affiche ⌘ sur Mac)
  labs: [ Lab, ... ],                         // 1 à 2 labs
  quiz: [ { q: "…", choices: ["a","b","c","d"], answer: 2, explain: "…" } ]  // 5 questions, answer = index 0-based
}
```

### HTML autorisé dans `lessons[].html`, `scenario`, `tasks[].text`, `hint`, `solution`
Balises simples : `p ul ol li b i code kbd table thead tbody tr th td h4 br`.
Classes disponibles :
- `<div class="callout">…</div>` (info), `callout ok` (bonne pratique), `callout warn` (piège), `callout bad` (erreur fréquente)
- `<kbd>Ctrl</kbd>+<kbd>C</kbd>` touches
- `<pre class="code">=SOMME(B2:B10)</pre>` formule / code VBA
- `<table class="o-table">` tableau comparatif
- `<div class="o-path">Accueil › Styles › Titre 1</div>` chemin dans le ruban (rendu en fil d'Ariane animé)
- `<div class="o-tip">💡 …</div>` astuce de formateur
Pas de `<script>`, pas d'images externes, pas de style inline sauf si indispensable.
Le contenu doit être **vrai pour Microsoft 365 / Office 2021 en français** (noms de menus FR, fonctions Excel FR).

## Lab
```js
{
  id: "x01-lab1",
  title: "Lab 1 — Le budget du séminaire",
  kind: "sheet" | "doc" | "slides" | "keys" | "steps",
  scenario: "<p>Contexte métier réaliste…</p>",
  setup: { … },              // dépend de kind (voir plus bas)
  tasks: [ { text: "…", hint: "…", check: api => booléen } ],   // 3 à 6 tâches, vérifiées en direct
  solution: "<ol><li>…</li></ol>"   // corrigé détaillé (affiché sur demande)
}
```
`check` est appelé à chaque modification (doit être rapide, ne jamais lever d'exception — le moteur
l'entoure d'un try/catch, une exception = tâche non validée). Les tâches sont validées dans n'importe quel ordre.

### kind "sheet" (tableur Excel simulé)
```js
setup: {
  cols: 8, rows: 15,                         // taille de la grille (max 12 × 30)
  data: { A1: "Produit", B1: "Prix HT", A2: "Stylo", B2: 1.5, C2: "=B2*1,2" },  // nombres = number, formules = "=…"
  bold: ["A1","B1"],                         // optionnel
  fmt: { B2: "eur" },                        // optionnel : "num" (2 déc.) | "eur" | "pct" | "int" | "date"
  widths: { A: 160 },                        // optionnel, px
  sheetName: "Budget"                        // optionnel
}
```
Moteur de formules : séparateur d'arguments **`;`** (toujours), décimales `.` ou `,` (`=B2*1,2` ok),
références relatives/absolues `A1 $A$1 A$1 $A1`, plages `A1:B10`, opérateurs `+ - * / ^ & = <> < > <= >= %`,
chaînes `"…"`. Fonctions (noms FR, alias EN acceptés) :
SOMME MOYENNE MIN MAX NB NBVAL NB.VIDE SI SI.CONDITIONS ET OU NON SIERREUR ARRONDI ARRONDI.SUP ARRONDI.INF ABS ENT MOD
PUISSANCE RACINE SOMME.SI NB.SI MOYENNE.SI SOMME.SI.ENS NB.SI.ENS SOMMEPROD RECHERCHEV RECHERCHEH RECHERCHEX INDEX EQUIV
CONCAT CONCATENER GAUCHE DROITE STXT NBCAR MAJUSCULE MINUSCULE NOMPROPRE SUPPRESPACE TEXTE AUJOURDHUI DATE ANNEE MOIS JOUR
ESTVIDE ESTNUM ESTTEXTE MAX.SI.ENS MIN.SI.ENS.
Critères SOMME.SI/NB.SI : `">100"`, `"Paris"`, `"<>Paris"`, `"*ris"` (joker *), ou référence de cellule.
Pas de fonctions matricielles dynamiques (FILTRE, UNIQUE, TRIER) ni de LET/LAMBDA : si un module Expert en parle,
le lab correspondant doit être de kind "steps" (à faire dans le vrai Excel).

Outils du simulateur (le stagiaire les utilise) : barre de formule, clic sur une cellule pendant la saisie d'une
formule = insère la référence, F4 = bascule $ , Ctrl+D / Ctrl+R recopie vers le bas / la droite (avec ajustement des
références relatives), Gras, couleur de remplissage, format de nombre, alignement, Trier A→Z / Z→A (colonne active,
zone de données contiguë, ligne 1 = en-têtes), Filtre (ligne 1), Mise en forme conditionnelle (règles > < = entre
contient, sur la sélection), Mettre sous forme de tableau, Figer la ligne 1, Graphique (barres, courbes, secteurs à
partir de la sélection), Tableau croisé dynamique (champ Lignes + champ Valeurs + Somme/Nombre/Moyenne).

API des `check` (sheet) :
```js
api.val("C2")              // valeur calculée : number | string | boolean | "" | code d'erreur "#DIV/0!" "#N/A" "#NOM?" "#VALEUR!" "#REF!"
api.num("C2")              // Number(val) ou NaN
api.f("C2")                // formule normalisée : MAJUSCULES, sans espaces, ";" — ex. "=SOMME(B2:B10)" ; "" si pas formule
api.isFormula("C2")        // booléen
api.hasFn("C2","SOMME")    // la formule utilise cette fonction (nom FR ou alias EN)
api.hasAbs("C2")           // la formule contient au moins une référence absolue ($)
api.range("B2:B6")         // tableau des valeurs (ligne par ligne)
api.style("A1")            // { bold, fill, fmt, align }  fill = couleur hex ou ""
api.near(a, b, eps=0.01)   // |a-b| <= eps
api.sortedBy("B", "asc"|"desc")   // la zone de données (lignes 2..n non vides) est triée selon cette colonne
api.cf                     // [{ range:"C2:C10", op:">"|"<"|"="|"between"|"contains", v1, v2, color }]
api.charts                 // [{ type:"bar"|"line"|"pie", range:"A1:B6" }]
api.pivots                 // [{ rows:"Région", values:"CA", agg:"sum"|"count"|"avg" }]   (noms d'en-têtes)
api.filter                 // true si filtre actif
api.table                  // true si « mis sous forme de tableau »
api.frozen                 // true si ligne 1 figée
api.cell(ref)              // { raw, val, style }
```
Exemple : `{ text: "En <b>D2</b>, calcule le TTC avec <code>=B2*(1+$F$1)</code> puis recopie jusqu'à D8.", hint: "F4 pour figer $F$1, puis Ctrl+D.", check: a => a.hasAbs("D2") && a.near(a.num("D8"), 54) }`

### kind "doc" (traitement de texte Word simulé)
```js
setup: { html: "<h1>Rapport</h1><p>Texte…</p>" }   // contenu initial de la page (p h1 h2 h3 ul ol li table blockquote b i u)
```
Outils : Styles (Normal = p, Titre 1/2/3 = h1/h2/h3, Citation = blockquote), Gras/Italique/Souligné/Barré,
couleur de police, surlignage, alignement (gauche, centre, droite, justifié), listes à puces / numérotées, retrait,
Insérer tableau (lignes × colonnes), saut de page (`hr.o-pagebreak`), table des matières (`div.o-toc`, générée
depuis h1-h3, bouton « Mettre à jour »), lien hypertexte, commentaire (`mark.o-comment[data-note]`), note de bas de
page (`sup.o-fn`), en-tête / pied de page (zones éditables), colonnes (1/2/3), Rechercher/Remplacer, Effacer la
mise en forme, Statistiques (mots).
API des `check` (doc) :
```js
api.root                   // élément de la page (contenteditable)
api.$(sel) / api.$$(sel)   // querySelector / querySelectorAll (tableau) dans la page
api.count(sel)             // nombre d'éléments
api.text()                 // texte complet de la page
api.words()                // nombre de mots
api.has("texte")           // le texte est présent (insensible à la casse)
api.fmt("fragment")        // mise en forme du 1er passage contenant ce fragment :
                           // { found, bold, italic, underline, strike, color, highlight, tag:"p"|"h1"…, align:"left"|"center"|"right"|"justify", list:"ul"|"ol"|null }
api.headings()             // [{ level:1, text:"…" }] dans l'ordre
api.header / api.footer    // texte de l'en-tête / du pied de page
api.columns                // 1 | 2 | 3
api.toc                    // true si une table des matières est insérée ; api.tocFresh : à jour par rapport aux titres
api.tables()               // [{ rows, cols, headerBold }]
```

### kind "slides" (PowerPoint simulé)
```js
setup: {
  theme: "blanc",            // "blanc" | "nuit" | "ocean" | "sunset" | "foret" | "neon"
  slides: [
    { layout: "title", title: "Lancement produit", subtitle: "Comité 2026" },
    { layout: "content", title: "Ordre du jour", body: "Contexte\nObjectifs\nPlanning" },   // \n = puces
    { layout: "two", title: "Avant / Après", left: "…", right: "…" },
    { layout: "section", title: "Partie 2" },
    { layout: "blank" }
  ]
}
```
Outils : miniatures (ajouter, dupliquer, supprimer, monter/descendre), dispositions, thèmes, zone de texte, formes
(rectangle, arrondi, ellipse, flèche), icône (emoji), déplacement/redimensionnement à la souris, édition du texte
(double-clic), taille / gras / couleur, aligner sur la diapo (gauche, centre, droite, haut, milieu, bas),
Animations par élément (aucune, apparition, entrée, zoom, rebond, + ordre), Transitions par diapo (aucune, fondu,
poussée, zoom, morphose), notes du présentateur, mode Diaporama plein écran.
API des `check` (slides) :
```js
api.slides                 // [{ layout, transition, notes, els:[{ type:"title"|"subtitle"|"body"|"text"|"shape"|"icon", text, x,y,w,h (en % de la diapo), size, bold, color, fill, shape, anim, order }] }]
api.theme                  // nom du thème
api.count()                // nombre de diapos
api.allText(i)             // texte concaténé de la diapo i (0-based)
api.find(txt)              // [{ slide, el }] éléments dont le texte contient txt (insensible à la casse)
api.centered(el)           // élément centré horizontalement (±3 %)
api.played                 // true si le diaporama a été lancé au moins une fois
```

### kind "keys" (défi raccourcis clavier, chronométré)
```js
setup: { sequence: [ { keys: "Ctrl+B", desc: "Mettre en gras" }, { keys: "Ctrl+Maj+L", desc: "Filtre" } ] }   // 5 à 10
```
`tasks` facultatif (générées automatiquement). N'utilise PAS de raccourcis réservés au navigateur :
Ctrl+W, Ctrl+T, Ctrl+N, Ctrl+Q, Ctrl+Tab, Ctrl+Maj+T/N/Q/W, Alt+F4, F11, F12, Ctrl+L, Ctrl+D (favoris), Ctrl+J, Ctrl+H.
Notation : `Ctrl`, `Maj`, `Alt`, `F1`…`F12`, lettres majuscules, `Entrée`, `Espace`, `Suppr`, `Échap`, `Flèche bas`, `+`, `;`.

### kind "steps" (lab dans la vraie application)
```js
setup: { app: "Excel", file: { name: "ventes.csv", type: "csv"|"doc"|"txt", content: "…" } }  // file optionnel : fichier de départ téléchargeable
```
Pour ce qui ne se simule pas (publipostage, macros VBA, Power Query, TCD avancés, FILTRE/LET/LAMBDA, Designer,
enregistrement vidéo…). `tasks` sans `check` : le stagiaire coche lui-même. CSV : séparateur `;`, décimales `,`.

## Progression pédagogique attendue (2 modules par niveau)
Niveau 1 Débutant → 4 Expert, chaque module construit sur le précédent. Scénarios métier variés et crédibles
(PME, RH, commercial, logistique, événementiel…). Labs = le cœur : au moins un lab simulé (sheet/doc/slides)
par module des niveaux 1-3 ; au moins un lab « keys » dans l'application ; niveau 4 peut mixer simulé + steps.
Vérifie bien tes `check` mentalement : valeurs attendues exactes (calcule-les !), noms de cellules corrects.
