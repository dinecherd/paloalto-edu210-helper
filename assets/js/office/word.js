/* Office Academy — Word : 8 modules du débutant à l'expert */
(function () {
  "use strict";

  /* Petits utilitaires pour des vérifications robustes (casse, espaces). */
  const norm = (s) => String(s == null ? "" : s).replace(/\s+/g, " ").trim().toLowerCase();
  const hasH = (a, lvl, frag) =>
    a.headings().some((h) => h.level === lvl && norm(h.text).includes(norm(frag)));
  const inList = (a, frag, type) => {
    const f = a.fmt(frag);
    return !!(f && f.found && f.list === type);
  };

  OFFICE.add("word", [
    /* ===================================================================
       NIVEAU 1 — w01 : Premiers pas
       =================================================================== */
    {
      id: "w01",
      level: 1,
      title: "Premiers pas : l'interface, la saisie et l'enregistrement",
      icon: "🧭",
      duration: 45,
      summary: "Apprivoisez le ruban, saisissez et corrigez un texte proprement, puis enregistrez-le au bon format et au bon endroit.",
      objectives: [
        "Être capable de se repérer dans le ruban, les onglets, les groupes et la barre d'outils Accès rapide",
        "Être capable de saisir un texte sans « bricolage » (pas d'espaces ni de retours à la ligne pour aligner)",
        "Être capable de sélectionner rapidement un mot, une ligne, un paragraphe ou tout le document",
        "Être capable de rechercher et remplacer un terme dans tout le document",
        "Être capable d'enregistrer en .docx, d'exporter en PDF et de retrouver l'historique des versions"
      ],
      lessons: [
        {
          title: "L'interface de Word 365",
          html: `<p>Word s'organise autour du <b>ruban</b> : une rangée d'<b>onglets</b> (Fichier, Accueil, Insertion, Dessin, Création, Mise en page, Références, Publipostage, Révision, Affichage, Aide), chacun découpé en <b>groupes</b> de commandes.</p>
<ul>
<li><b>Fichier</b> ouvre le mode « Backstage » : Nouveau, Ouvrir, Informations, Enregistrer sous, Imprimer, Exporter, Options.</li>
<li><b>Accueil</b> regroupe 80 % des besoins quotidiens : Presse-papiers, Police, Paragraphe, Styles, Modification.</li>
<li>La <b>barre d'outils Accès rapide</b> (au-dessus ou sous le ruban) accueille vos commandes favorites : clic droit sur n'importe quel bouton › <i>Ajouter à la barre d'outils Accès rapide</i>.</li>
<li>La <b>barre d'état</b> (en bas) affiche le numéro de page, le nombre de mots, la langue et le zoom. Clic droit dessus pour choisir les informations affichées.</li>
<li>Le champ <b>Rechercher</b> (loupe, en haut) trouve une commande dont vous ignorez l'emplacement : tapez « filigrane » et Word vous la propose.</li>
</ul>
<div class="o-path">Affichage › Afficher › Règle</div>
<p>Affichez la règle : elle montre les retraits et les tabulations du paragraphe courant, indispensable pour comprendre la mise en page.</p>
<div class="o-tip">💡 Double-cliquez sur un onglet pour replier/déplier le ruban et gagner de la place à l'écran.</div>
<div class="callout">Les <b>onglets contextuels</b> (Création de tableau, Disposition, Format de l'image, En-tête et pied de page…) n'apparaissent que lorsqu'un objet est sélectionné. Si vous ne trouvez pas une commande de tableau, cliquez d'abord <i>dans</i> le tableau.</div>`
        },
        {
          title: "Saisir proprement : les marques de mise en forme",
          html: `<p>Word gère lui-même le retour à la ligne : on n'appuie sur <kbd>Entrée</kbd> qu'à la <b>fin d'un paragraphe</b>.</p>
<div class="o-path">Accueil › Paragraphe › Afficher tout (¶)</div>
<p>Ce bouton révèle les caractères invisibles :</p>
<table class="o-table"><thead><tr><th>Symbole</th><th>Signification</th><th>Touche</th></tr></thead>
<tbody>
<tr><td>¶</td><td>Fin de paragraphe</td><td><kbd>Entrée</kbd></td></tr>
<tr><td>↵</td><td>Saut de ligne (même paragraphe)</td><td><kbd>Maj</kbd>+<kbd>Entrée</kbd></td></tr>
<tr><td>·</td><td>Espace</td><td><kbd>Espace</kbd></td></tr>
<tr><td>°</td><td>Espace insécable</td><td><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>Espace</kbd></td></tr>
<tr><td>→</td><td>Tabulation</td><td><kbd>Tab</kbd></td></tr>
</tbody></table>
<div class="callout bad">Erreur fréquente : aligner du texte avec une série d'espaces ou ajouter des lignes vides pour « passer à la page suivante ». Au moindre changement, tout se décale. Utilisez l'alignement, les tabulations et le saut de page (<kbd>Ctrl</kbd>+<kbd>Entrée</kbd>).</div>
<div class="callout warn">En français, on met une espace avant <b>: ; ! ?</b> et à l'intérieur des « guillemets ». Utilisez l'espace insécable pour éviter qu'un « ? » ou « € » se retrouve seul en début de ligne. La correction automatique de Word en ajoute souvent pour vous.</div>
<p>Les soulignements ondulés : <b>rouge</b> = orthographe, <b>bleu double</b> = grammaire/style. Clic droit pour voir les suggestions, ou <kbd>F7</kbd> pour lancer le volet <i>Éditeur</i>.</p>`
        },
        {
          title: "Sélectionner vite et bien",
          html: `<p>La plupart des commandes s'appliquent à la <b>sélection</b>. Maîtriser la sélection, c'est gagner un temps énorme.</p>
<table class="o-table"><thead><tr><th>Je veux sélectionner…</th><th>Geste</th></tr></thead>
<tbody>
<tr><td>Un mot</td><td>Double-clic</td></tr>
<tr><td>Une phrase</td><td><kbd>Ctrl</kbd>+clic dans la phrase</td></tr>
<tr><td>Un paragraphe</td><td>Triple-clic</td></tr>
<tr><td>Une ligne</td><td>Clic dans la marge gauche (curseur en flèche)</td></tr>
<tr><td>Tout le document</td><td><kbd>Ctrl</kbd>+<kbd>A</kbd> ou triple-clic dans la marge</td></tr>
<tr><td>Du curseur jusqu'à un point</td><td><kbd>Maj</kbd>+clic</td></tr>
<tr><td>Mot par mot au clavier</td><td><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>Flèche droite</kbd></td></tr>
<tr><td>Un bloc vertical</td><td><kbd>Alt</kbd>+glisser</td></tr>
</tbody></table>
<div class="o-tip">💡 <kbd>Ctrl</kbd> maintenu pendant que vous sélectionnez à la souris permet d'ajouter plusieurs passages non contigus à la sélection.</div>
<p>Pour déplacer : sélection puis <kbd>Ctrl</kbd>+<kbd>X</kbd> / <kbd>Ctrl</kbd>+<kbd>V</kbd>, ou glisser-déposer. Au collage, la petite balise <b>(Ctrl)</b> propose : <i>Conserver la mise en forme source</i>, <i>Fusionner la mise en forme</i>, <i>Conserver uniquement le texte</i>. Ce dernier est idéal pour coller depuis un site web.</p>
<div class="o-path">Accueil › Modification › Remplacer</div>
<p><b>Rechercher et remplacer</b> (<kbd>Ctrl</kbd>+<kbd>H</kbd>) corrige un nom mal orthographié dans tout le document en un clic sur <i>Remplacer tout</i>. Le bouton <i>Plus &gt;&gt;</i> donne accès à <i>Respecter la casse</i>, <i>Mot entier</i> et aux caractères spéciaux (^p = marque de paragraphe).</p>`
        },
        {
          title: "Enregistrer, exporter, retrouver une version",
          html: `<div class="o-path">Fichier › Enregistrer sous › Parcourir</div>
<table class="o-table"><thead><tr><th>Format</th><th>Usage</th></tr></thead>
<tbody>
<tr><td><code>.docx</code></td><td>Format standard Word, sans macros</td></tr>
<tr><td><code>.docm</code></td><td>Document contenant des macros VBA</td></tr>
<tr><td><code>.dotx</code></td><td>Modèle (sert de base à de nouveaux documents)</td></tr>
<tr><td><code>.pdf</code></td><td>Diffusion figée, non modifiable facilement</td></tr>
<tr><td><code>.odt</code> / <code>.rtf</code></td><td>Échange avec d'autres traitements de texte</td></tr>
</tbody></table>
<div class="o-path">Fichier › Exporter › Créer un document PDF/XPS</div>
<p>Dans la fenêtre d'export, le bouton <i>Options…</i> permet de n'exporter que certaines pages et de <b>créer des signets à partir des titres</b> (navigation dans le PDF).</p>
<div class="callout ok">Enregistrez dans <b>OneDrive</b> ou <b>SharePoint</b> : l'<b>enregistrement automatique</b> (interrupteur en haut à gauche) s'active, et <i>Fichier › Informations › Historique des versions</i> vous permet de revenir à une version d'hier.</div>
<div class="callout warn">Fichier local et Word se ferme brutalement ? Rouvrez Word : le volet <i>Récupération de document</i> propose les fichiers non enregistrés. Sinon : <i>Fichier › Informations › Gérer le document › Récupérer des documents non enregistrés</i>.</div>
<div class="o-tip">💡 Nommez vos fichiers de façon triable : <code>2026-10-25_Relance_Dupont_v2.docx</code>. La date en tête (AAAA-MM-JJ) classe automatiquement par ordre chronologique.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+S", desc: "Enregistrer" },
        { keys: "F12", desc: "Enregistrer sous" },
        { keys: "Ctrl+Z", desc: "Annuler la dernière action" },
        { keys: "Ctrl+Y", desc: "Rétablir / répéter" },
        { keys: "Ctrl+A", desc: "Sélectionner tout le document" },
        { keys: "Ctrl+H", desc: "Rechercher et remplacer" },
        { keys: "Maj+Entrée", desc: "Saut de ligne dans le même paragraphe" },
        { keys: "F7", desc: "Orthographe et grammaire (Éditeur)" }
      ],
      labs: [
        {
          id: "w01-lab1",
          title: "Lab 1 — La lettre de relance",
          kind: "doc",
          scenario: `<p>Vous êtes assistant(e) comptable chez <b>Mobilier Martin SARL</b>. Votre collègue a tapé à la hâte une lettre de relance pour une facture impayée. Le nom du client est mal écrit partout, un paragraphe a été collé deux fois et la lettre n'a pas de formule de politesse. Corrigez-la avant envoi.</p>`,
          setup: {
            html: `<p>Mobilier Martin SARL<br>12 rue des Artisans<br>69007 Lyon</p>
<p>Sté Dupond<br>Service comptabilité<br>4 avenue Jean Jaurès<br>38000 Grenoble</p>
<p>Lyon, le 25 octobre 2026</p>
<p>Objet : relance facture n° 2026-118</p>
<p>Madame, Monsieur,</p>
<p>Sauf erreur de notre part, la facture n° 2026-118 du 15 septembre 2026, d'un montant de 4 860,00 € TTC, adressée à la Sté Dupond, reste à ce jour impayée alors que son échéance était fixée au 15 octobre.</p>
<p>[À SUPPRIMER] Sauf erreur de notre part, la facture reste à ce jour impayée.</p>
<p>Nous vous remercions de bien vouloir procéder à son règlement dans les meilleurs délais. Si votre paiement a été effectué entre-temps, merci de ne pas tenir compte de ce courrier. La Sté Dupond reste un client précieux et nous restons à votre disposition pour tout renseignement.</p>
<p>Claire Martin<br>Responsable comptabilité</p>`
          },
          tasks: [
            {
              text: "Remplacez <b>toutes</b> les occurrences de « Sté Dupond » par « Société Dupont ».",
              hint: "Ctrl+H (ou Accueil › Modification › Remplacer), puis « Remplacer tout ».",
              check: (a) => a.has("Société Dupont") && !a.has("Sté Dupond")
            },
            {
              text: "Mettez la ligne d'objet « Objet : relance facture n° 2026-118 » en <b>gras</b>.",
              hint: "Clic dans la marge gauche pour sélectionner la ligne, puis Ctrl+G.",
              check: (a) => { const f = a.fmt("relance facture"); return !!(f.found && f.bold); }
            },
            {
              text: "Supprimez le paragraphe en double qui commence par « [À SUPPRIMER] ».",
              hint: "Triple-clic dans le paragraphe pour le sélectionner entièrement, puis Suppr.",
              check: (a) => !a.has("[à supprimer]") && a.has("procéder à son règlement")
            },
            {
              text: "Ajoutez avant la signature une formule de politesse contenant « salutations distinguées ».",
              hint: "Par exemple : « Veuillez agréer, Madame, Monsieur, nos salutations distinguées. »",
              check: (a) => norm(a.text()).includes("salutations distinguées")
            },
            {
              text: "Mettez le nom du signataire « Claire Martin » en <i>italique</i>.",
              hint: "Double-clic sur « Claire », Maj+clic après « Martin », puis Ctrl+I.",
              check: (a) => { const f = a.fmt("Claire Martin"); return !!(f.found && f.italic); }
            }
          ],
          solution: `<ol>
<li><kbd>Ctrl</kbd>+<kbd>H</kbd> › Rechercher : <code>Sté Dupond</code> › Remplacer par : <code>Société Dupont</code> › <i>Remplacer tout</i> (3 remplacements).</li>
<li>Sélectionnez la ligne « Objet : … » (clic dans la marge) › <kbd>Ctrl</kbd>+<kbd>G</kbd>.</li>
<li>Triple-clic sur le paragraphe « [À SUPPRIMER] … » › <kbd>Suppr</kbd>.</li>
<li>Placez le curseur à la fin du dernier paragraphe › <kbd>Entrée</kbd> › tapez « Veuillez agréer, Madame, Monsieur, nos salutations distinguées. »</li>
<li>Sélectionnez « Claire Martin » › <kbd>Ctrl</kbd>+<kbd>I</kbd>.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Vous ne trouvez pas l'onglet « Disposition » des tableaux. Pourquoi ?",
          choices: ["Il faut l'activer dans les Options", "C'est un onglet contextuel : il faut cliquer dans un tableau", "Il n'existe que dans Word pour Mac", "Il est dans l'onglet Fichier"],
          answer: 1,
          explain: "Les onglets contextuels (Création de tableau, Disposition) n'apparaissent que lorsque le curseur est dans un tableau."
        },
        {
          q: "Quel est le bon moyen de commencer un nouveau chapitre en haut d'une nouvelle page ?",
          choices: ["Appuyer sur Entrée jusqu'à changer de page", "Insérer des espaces", "Insérer un saut de page (Ctrl+Entrée)", "Réduire la taille de police"],
          answer: 2,
          explain: "Le saut de page force la page suivante quel que soit le contenu au-dessus. Les lignes vides se décalent dès qu'on modifie le texte."
        },
        {
          q: "Comment sélectionner tout un paragraphe en un geste ?",
          choices: ["Triple-clic dans le paragraphe", "Double-clic dans le paragraphe", "Ctrl+clic dans le paragraphe", "Clic droit › Sélectionner"],
          answer: 0,
          explain: "Double-clic = mot, Ctrl+clic = phrase, triple-clic = paragraphe."
        },
        {
          q: "Quel format choisir pour un document contenant des macros VBA ?",
          choices: [".docx", ".dotx", ".pdf", ".docm"],
          answer: 3,
          explain: "Le .docx ne peut pas contenir de macros : Word les supprime à l'enregistrement. Il faut un .docm (ou .dotm pour un modèle)."
        },
        {
          q: "Que fait Maj+Entrée ?",
          choices: ["Un saut de page", "Un saut de ligne sans créer de nouveau paragraphe", "Un saut de section", "Valide la correction orthographique"],
          answer: 1,
          explain: "Le saut de ligne (↵) passe à la ligne en restant dans le même paragraphe : pratique pour une adresse, sans espace entre les lignes."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 1 — w02 : Mise en forme
       =================================================================== */
    {
      id: "w02",
      level: 1,
      title: "Mise en forme des caractères, des paragraphes et des listes",
      icon: "🎨",
      duration: 50,
      summary: "Donnez à vos documents une allure professionnelle : polices, alignements, espacements, retraits et listes maîtrisés.",
      objectives: [
        "Être capable de distinguer mise en forme des caractères et mise en forme des paragraphes",
        "Être capable de régler alignement, interlignes, espacements avant/après et retraits",
        "Être capable de créer des listes à puces et numérotées, et de changer de niveau",
        "Être capable de reproduire une mise en forme avec le pinceau",
        "Être capable d'utiliser les raccourcis de mise en forme les plus courants"
      ],
      lessons: [
        {
          title: "Caractères : police, taille, attributs",
          html: `<div class="o-path">Accueil › Police</div>
<p>La mise en forme <b>des caractères</b> s'applique à la sélection : police, taille, <b>gras</b> (<kbd>Ctrl</kbd>+<kbd>G</kbd>), <i>italique</i> (<kbd>Ctrl</kbd>+<kbd>I</kbd>), souligné (<kbd>Ctrl</kbd>+<kbd>U</kbd>), barré, exposant (<kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>+</kbd>), indice (<kbd>Ctrl</kbd>+<kbd>=</kbd>), couleur de police et couleur de surbrillance.</p>
<p>Le bouton <b>Modifier la casse</b> (Aa) — ou <kbd>Maj</kbd>+<kbd>F3</kbd> qui alterne minuscules / MAJUSCULES / Majuscule En Début De Mot — évite de retaper un titre écrit en majuscules.</p>
<p>Pour tout régler d'un coup (espacement entre caractères, petites majuscules, masqué) : le <b>lanceur de boîte de dialogue</b> (petite flèche en bas à droite du groupe Police) ou <kbd>Ctrl</kbd>+<kbd>D</kbd> dans Word.</p>
<div class="callout bad">Une couleur seule ne transmet pas une information : une personne daltonienne ou une impression en noir et blanc la perdra. Doublez toujours la couleur d'un mot (« Obligatoire : ») ou d'un gras.</div>
<div class="o-tip">💡 Limitez-vous à <b>deux polices</b> par document (une pour les titres, une pour le texte). La police par défaut de Microsoft 365 est <b>Aptos</b> (Calibri sur les versions antérieures).</div>`
        },
        {
          title: "Paragraphes : alignement, espacements, retraits",
          html: `<p>La mise en forme <b>de paragraphe</b> s'applique au paragraphe entier où se trouve le curseur — inutile de le sélectionner.</p>
<div class="o-path">Accueil › Paragraphe › Interligne et espacement de paragraphe</div>
<table class="o-table"><thead><tr><th>Réglage</th><th>Raccourci</th><th>Conseil</th></tr></thead>
<tbody>
<tr><td>Aligné à gauche</td><td><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>G</kbd></td><td>Par défaut, le plus lisible</td></tr>
<tr><td>Centré</td><td><kbd>Ctrl</kbd>+<kbd>E</kbd></td><td>Titres, pages de garde</td></tr>
<tr><td>Aligné à droite</td><td><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>D</kbd></td><td>Date, lieu dans un courrier</td></tr>
<tr><td>Justifié</td><td><kbd>Ctrl</kbd>+<kbd>J</kbd></td><td>Textes longs, avec la coupure de mots</td></tr>
</tbody></table>
<p><b>Espacement avant / après</b> (Mise en page › Paragraphe › Espacement) : c'est lui, et non une ligne vide, qui doit séparer les paragraphes. Exemple : 0 pt avant, 8 pt après.</p>
<p><b>Retraits</b> : retrait gauche, retrait droit, <b>première ligne</b> (alinéa) ou <b>suspendu</b> (la 1ʳᵉ ligne dépasse, utile pour une bibliographie). Réglez-les avec les curseurs de la règle ou dans la boîte <i>Paragraphe</i>.</p>
<div class="callout warn">Les options <i>Paragraphe › Enchaînements</i> sont précieuses : <b>Paragraphes solidaires</b> (garde un titre avec le paragraphe qui suit) et <b>Éviter veuves et orphelines</b> (pas de ligne isolée en haut ou en bas de page).</div>`
        },
        {
          title: "Listes à puces, numérotées et à plusieurs niveaux",
          html: `<div class="o-path">Accueil › Paragraphe › Puces / Numérotation / Liste à plusieurs niveaux</div>
<ul>
<li><b>Puces</b> : éléments sans ordre (avantages, matériel à prévoir).</li>
<li><b>Numérotation</b> : étapes, procédure, classement.</li>
<li><b>Liste à plusieurs niveaux</b> : plan 1. / 1.1 / 1.1.1.</li>
</ul>
<p>Dans une liste : <kbd>Tab</kbd> descend d'un niveau, <kbd>Maj</kbd>+<kbd>Tab</kbd> remonte. <kbd>Entrée</kbd> deux fois de suite sort de la liste.</p>
<p>Numérotation qui reprend à 4 au lieu de 1 ? Clic droit sur le numéro › <b>Recommencer à 1</b>. À l'inverse, <i>Continuer la numérotation</i> relie deux listes séparées par un paragraphe.</p>
<div class="callout bad">Taper « 1. » ou « - » à la main puis une tabulation : Word convertit souvent en liste automatique (Correction automatique), mais si vous désactivez cela et tapez les numéros vous-même, ils ne se mettront jamais à jour quand vous insérerez une étape.</div>
<div class="o-tip">💡 Pour une puce personnalisée (✔, ➜, logo) : flèche du bouton Puces › <i>Définir une puce…</i> › Symbole.</div>`
        },
        {
          title: "Reproduire et effacer une mise en forme",
          html: `<div class="o-path">Accueil › Presse-papiers › Reproduire la mise en forme</div>
<p>Le <b>pinceau</b> copie la mise en forme d'un passage pour l'appliquer ailleurs : cliquez dans le modèle, cliquez sur le pinceau, puis « peignez » le texte cible. <b>Double-cliquez</b> sur le pinceau pour l'appliquer à plusieurs endroits ; <kbd>Échap</kbd> pour arrêter.</p>
<p>Au clavier : <kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>C</kbd> copie la mise en forme, <kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>V</kbd> la colle.</p>
<div class="o-path">Accueil › Police › Effacer toute la mise en forme (A avec gomme)</div>
<p>Remet le texte sélectionné au style <i>Normal</i>. <kbd>Ctrl</kbd>+<kbd>Espace</kbd> efface seulement la mise en forme des caractères.</p>
<div class="callout ok">Bonne pratique : si vous appliquez la même mise en forme manuelle plus de deux fois (titre en bleu 14 pt gras…), c'est le signe qu'il vous faut un <b>style</b> — c'est l'objet du module suivant.</div>
<div class="o-tip">💡 Pour comprendre pourquoi un texte « ne veut pas » changer : <kbd>Maj</kbd>+<kbd>F1</kbd> ouvre le volet <i>Révéler la mise en forme</i> qui détaille police, paragraphe et style appliqués.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+G", desc: "Gras" },
        { keys: "Ctrl+I", desc: "Italique" },
        { keys: "Ctrl+U", desc: "Souligné" },
        { keys: "Ctrl+E", desc: "Centrer le paragraphe" },
        { keys: "Ctrl+J", desc: "Justifier le paragraphe" },
        { keys: "Maj+F3", desc: "Modifier la casse" },
        { keys: "Ctrl+Maj+C", desc: "Copier la mise en forme" },
        { keys: "Ctrl+Espace", desc: "Effacer la mise en forme des caractères" }
      ],
      labs: [
        {
          id: "w02-lab1",
          title: "Lab 1 — La note de service RH",
          kind: "doc",
          scenario: `<p>Le service RH de <b>Logistique Rhône-Alpes</b> diffuse demain une note sur la nouvelle charte du télétravail. Le texte est validé mais complètement brut : pas de titre visible, conditions et étapes noyées dans le texte. Mettez-le en forme pour qu'il se lise en 30 secondes.</p>`,
          setup: {
            html: `<p>Note de service : nouvelle charte du télétravail</p>
<p>À compter du 1er décembre 2026, tous les collaborateurs dont le poste est compatible peuvent demander jusqu'à trois jours de télétravail par semaine, sous réserve de respecter au moins deux jours sur site par semaine. Cette organisation vise à améliorer l'équilibre entre vie professionnelle et vie personnelle tout en préservant le collectif de travail.</p>
<p>Conditions à remplir :</p>
<p>Ordinateur portable fourni par l'entreprise</p>
<p>Connexion internet personnelle stable</p>
<p>Espace de travail calme et adapté</p>
<p>Plages de joignabilité de 9 h 30 à 16 h 30</p>
<p>Comment faire sa demande :</p>
<p>Remplir le formulaire sur l'intranet</p>
<p>Faire valider la demande par son manager</p>
<p>Signer l'avenant transmis par les RH</p>
<p>Les demandes doivent parvenir au service RH avant le 30 novembre 2026.</p>
<p>Le service des ressources humaines</p>`
          },
          tasks: [
            {
              text: "Mettez le titre « Note de service : nouvelle charte du télétravail » en <b>gras</b> et <b>centré</b>.",
              hint: "Sélectionnez la ligne, Ctrl+G puis Ctrl+E.",
              check: (a) => { const f = a.fmt("nouvelle charte du"); return !!(f.found && f.bold && f.align === "center"); }
            },
            {
              text: "<b>Justifiez</b> le paragraphe d'introduction (« À compter du 1er décembre… »).",
              hint: "Cliquez simplement dans le paragraphe : l'alignement s'applique au paragraphe entier.",
              check: (a) => { const f = a.fmt("compter du 1er"); return !!(f.found && f.align === "justify"); }
            },
            {
              text: "Transformez les 4 conditions (Ordinateur… → Plages…) en <b>liste à puces</b>.",
              hint: "Sélectionnez les 4 lignes, puis Accueil › Paragraphe › Puces.",
              check: (a) => ["Ordinateur portable", "Connexion internet", "Espace de travail", "Plages de joignabilit"].every((t) => inList(a, t, "ul"))
            },
            {
              text: "Transformez les 3 étapes de la demande en <b>liste numérotée</b>.",
              hint: "Sélectionnez de « Remplir… » à « Signer… », puis Numérotation.",
              check: (a) => ["Remplir le formulaire", "Faire valider", "Signer l"].every((t) => inList(a, t, "ol"))
            },
            {
              text: "Mettez en <b>surbrillance</b> (surligneur) la date limite « 30 novembre 2026 ».",
              hint: "Accueil › Police › Couleur de surbrillance du texte.",
              check: (a) => { const f = a.fmt("30 novembre"); return !!(f.found && f.highlight); }
            },
            {
              text: "<u>Soulignez</u> la contrainte « au moins deux jours sur site » dans l'introduction.",
              hint: "Sélectionnez exactement ce passage, puis Ctrl+U.",
              check: (a) => { const f = a.fmt("deux jours sur site"); return !!(f.found && f.underline); }
            }
          ],
          solution: `<ol>
<li>Sélectionnez la 1ʳᵉ ligne › <kbd>Ctrl</kbd>+<kbd>G</kbd> › <kbd>Ctrl</kbd>+<kbd>E</kbd>.</li>
<li>Clic dans le paragraphe « À compter du… » › <kbd>Ctrl</kbd>+<kbd>J</kbd> (ou bouton Justifier).</li>
<li>Sélectionnez les lignes « Ordinateur portable… » à « Plages de joignabilité… » › <i>Accueil › Paragraphe › Puces</i>.</li>
<li>Sélectionnez « Remplir le formulaire… » à « Signer l'avenant… » › <i>Numérotation</i> (1., 2., 3.).</li>
<li>Sélectionnez « 30 novembre 2026 » › <i>Couleur de surbrillance du texte</i> (jaune).</li>
<li>Sélectionnez « au moins deux jours sur site » › <kbd>Ctrl</kbd>+<kbd>U</kbd>.</li>
</ol>`
        },
        {
          id: "w02-lab2",
          title: "Lab 2 — Défi raccourcis de mise en forme",
          kind: "keys",
          scenario: `<p>Les pros de la bureautique ne lâchent pas le clavier. Enchaînez ces raccourcis Word le plus vite possible : le chrono tourne !</p>`,
          setup: {
            sequence: [
              { keys: "Ctrl+G", desc: "Mettre en gras" },
              { keys: "Ctrl+I", desc: "Mettre en italique" },
              { keys: "Ctrl+U", desc: "Souligner" },
              { keys: "Ctrl+E", desc: "Centrer le paragraphe" },
              { keys: "Ctrl+Alt+1", desc: "Appliquer le style Titre 1" },
              { keys: "Ctrl+Z", desc: "Annuler" },
              { keys: "Ctrl+Y", desc: "Rétablir" },
              { keys: "Maj+Entrée", desc: "Saut de ligne" },
              { keys: "Ctrl+Entrée", desc: "Saut de page" },
              { keys: "Ctrl+A", desc: "Tout sélectionner" }
            ]
          },
          solution: `<ul><li>Gras <kbd>Ctrl</kbd>+<kbd>G</kbd>, Italique <kbd>Ctrl</kbd>+<kbd>I</kbd>, Souligné <kbd>Ctrl</kbd>+<kbd>U</kbd>, Centrer <kbd>Ctrl</kbd>+<kbd>E</kbd>.</li><li>Titre 1 <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>1</kbd>, Annuler <kbd>Ctrl</kbd>+<kbd>Z</kbd>, Rétablir <kbd>Ctrl</kbd>+<kbd>Y</kbd>.</li><li>Saut de ligne <kbd>Maj</kbd>+<kbd>Entrée</kbd>, saut de page <kbd>Ctrl</kbd>+<kbd>Entrée</kbd>, tout sélectionner <kbd>Ctrl</kbd>+<kbd>A</kbd>.</li></ul>`
        }
      ],
      quiz: [
        {
          q: "Pour centrer un paragraphe, que faut-il sélectionner ?",
          choices: ["Tout le paragraphe obligatoirement", "Rien : il suffit que le curseur soit dans le paragraphe", "Tout le document", "Le premier mot uniquement"],
          answer: 1,
          explain: "L'alignement est une mise en forme de paragraphe : elle s'applique au paragraphe entier où se trouve le curseur."
        },
        {
          q: "Comment séparer proprement deux paragraphes ?",
          choices: ["Avec une ligne vide", "Avec plusieurs espaces", "Avec un saut de page", "Avec l'espacement après paragraphe"],
          answer: 3,
          explain: "L'espacement avant/après (en points) se règle une fois, s'applique partout via les styles et ne crée pas de paragraphes vides parasites."
        },
        {
          q: "Une liste numérotée reprend à 5 alors que vous voulez 1. Que faites-vous ?",
          choices: ["Clic droit sur le numéro › Recommencer à 1", "Supprimer la liste et taper 1. à la main", "Ctrl+Z", "Changer de police"],
          answer: 0,
          explain: "« Recommencer à 1 » démarre une nouvelle liste ; « Continuer la numérotation » fait l'inverse."
        },
        {
          q: "Comment appliquer la mise en forme d'un titre à 5 autres titres d'un coup ?",
          choices: ["Copier-coller les titres", "Simple clic sur le pinceau, puis les 5 titres", "Double-clic sur le pinceau, puis peindre les 5 titres, puis Échap", "Ctrl+A puis Gras"],
          answer: 2,
          explain: "Le double-clic verrouille le pinceau pour plusieurs applications. (Mieux encore : créez un style !)"
        },
        {
          q: "Dans une liste à puces, quelle touche fait descendre un élément d'un niveau ?",
          choices: ["Entrée", "Tab", "Maj+Tab", "Espace"],
          answer: 1,
          explain: "Tab descend d'un niveau, Maj+Tab remonte."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 2 — w03 : Styles, navigation, table des matières
       =================================================================== */
    {
      id: "w03",
      level: 2,
      title: "Styles, titres, navigation et table des matières",
      icon: "🗂️",
      duration: 55,
      summary: "Les styles sont le super-pouvoir de Word : une structure propre, une navigation instantanée et une table des matières qui se met à jour seule.",
      objectives: [
        "Être capable d'appliquer et de modifier les styles Titre 1, Titre 2, Normal et Citation",
        "Être capable de naviguer et réorganiser un document avec le volet de navigation",
        "Être capable d'insérer et de mettre à jour une table des matières automatique",
        "Être capable de numéroter automatiquement les titres (1, 1.1, 1.1.1)"
      ],
      lessons: [
        {
          title: "Pourquoi les styles changent tout",
          html: `<p>Un <b>style</b> est un ensemble nommé de mises en forme (police, taille, couleur, espacements, enchaînements…). Au lieu de mettre « Arial 16 gras bleu » à la main sur chaque titre, vous appliquez <b>Titre 1</b>.</p>
<div class="o-path">Accueil › Styles › Titre 1</div>
<ul>
<li><b>Cohérence</b> : tous les titres d'un même niveau sont identiques.</li>
<li><b>Modification globale</b> : changez le style, les 40 titres changent.</li>
<li><b>Structure</b> : Word sait ce qui est un titre → volet de navigation, table des matières, signets PDF, accessibilité.</li>
</ul>
<table class="o-table"><thead><tr><th>Style</th><th>Raccourci</th><th>Usage</th></tr></thead>
<tbody>
<tr><td>Normal</td><td><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>N</kbd></td><td>Corps de texte</td></tr>
<tr><td>Titre 1</td><td><kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>1</kbd></td><td>Grandes parties</td></tr>
<tr><td>Titre 2</td><td><kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>2</kbd></td><td>Sous-parties</td></tr>
<tr><td>Titre 3</td><td><kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>3</kbd></td><td>Sous-sous-parties</td></tr>
<tr><td>Citation</td><td>—</td><td>Verbatim, témoignage</td></tr>
</tbody></table>
<div class="callout bad">Erreur fréquente : un paragraphe de texte mis en Titre 2 « parce qu'il était joli ». Il apparaît alors dans la table des matières ! Le style décrit la <b>fonction</b> du texte, pas son apparence.</div>`
        },
        {
          title: "Modifier et créer des styles",
          html: `<p>Deux méthodes pour modifier un style :</p>
<ol>
<li><b>Par l'exemple</b> : mettez en forme un titre comme vous le souhaitez, puis clic droit sur <i>Titre 1</i> dans la galerie › <b>Mettre à jour Titre 1 pour correspondre à la sélection</b>.</li>
<li><b>Par la boîte de dialogue</b> : clic droit sur le style › <b>Modifier…</b> › bouton <i>Format</i> (Police, Paragraphe, Bordure, Numérotation…).</li>
</ol>
<div class="o-path">Accueil › Styles › lanceur (↘) › Nouveau style</div>
<p>Le volet <b>Styles</b> (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>Maj</kbd>+<kbd>S</kbd>) liste tous les styles et permet d'en créer. Pour un nouveau style de paragraphe, choisissez bien <i>Style basé sur</i> (héritage) et <i>Style du paragraphe suivant</i> (après un Titre 1, on revient en Normal).</p>
<div class="callout warn">Dans la boîte <i>Modifier le style</i>, l'option <b>Mettre à jour automatiquement</b> doit rester décochée : sinon la moindre mise en forme manuelle sur un titre modifie tous les titres du document.</div>
<div class="o-tip">💡 <i>Création › Mise en forme du document</i> propose des jeux de styles complets et <i>Couleurs</i> / <i>Polices</i> changent le thème de tout le document en un clic — à condition d'avoir utilisé des styles.</div>`
        },
        {
          title: "Le volet de navigation",
          html: `<div class="o-path">Affichage › Afficher › Volet de navigation</div>
<p>Le volet (aussi <kbd>Ctrl</kbd>+<kbd>F</kbd>) a trois onglets : <b>Titres</b>, <b>Pages</b> (miniatures) et <b>Résultats</b> (recherche).</p>
<ul>
<li>Cliquez sur un titre pour y sauter immédiatement.</li>
<li><b>Glissez-déposez</b> un titre pour déplacer toute la section (titre + sous-titres + texte).</li>
<li>Clic droit sur un titre : <i>Hausser</i> / <i>Abaisser</i> d'un niveau, <i>Nouveau titre avant/après</i>, <i>Supprimer</i> la section, <i>Réduire tout</i>.</li>
</ul>
<div class="callout">Le volet n'affiche que les paragraphes ayant un style de titre (ou un niveau hiérarchique). Un titre vide de la liste ? C'est qu'il a été mis en gras à la main au lieu de recevoir un style.</div>
<p>Le <b>mode Plan</b> (Affichage › Plan) offre la même puissance pour restructurer un long rapport, avec affichage par niveau.</p>
<div class="o-tip">💡 Dans un document de 80 pages, <kbd>Maj</kbd>+<kbd>F5</kbd> vous ramène à l'endroit de votre dernière modification, même après réouverture du fichier.</div>`
        },
        {
          title: "Table des matières et numérotation des titres",
          html: `<div class="o-path">Références › Table des matières › Table automatique 1</div>
<p>Word construit la table à partir des styles Titre 1 à Titre 3 (par défaut). Placez le curseur à l'endroit voulu (souvent après la page de garde) avant d'insérer.</p>
<div class="o-path">Références › Table des matières › Mettre à jour la table › Mettre à jour toute la table</div>
<div class="callout warn">La table <b>ne se met pas à jour toute seule</b>. Après avoir ajouté un titre ou changé la pagination, mettez-la à jour (ou <kbd>F9</kbd> dans la table). Choisissez « toute la table » pour récupérer les nouveaux titres, pas seulement les numéros de page.</div>
<p><i>Table des matières personnalisée…</i> permet de régler le nombre de niveaux, les points de suite et d'intégrer vos propres styles (bouton <i>Options</i>).</p>
<div class="o-path">Accueil › Paragraphe › Liste à plusieurs niveaux › 1 Titre 1 / 1.1 Titre 2</div>
<p>Choisissez dans la bibliothèque la liste qui affiche « Titre 1, Titre 2… » : les titres se numérotent automatiquement et se renumérotent si vous déplacez une section.</p>
<div class="callout bad">Ne tapez jamais « 2.3 » à la main dans un titre : au premier déplacement de section, la numérotation devient fausse.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+Alt+1", desc: "Appliquer le style Titre 1" },
        { keys: "Ctrl+Alt+2", desc: "Appliquer le style Titre 2" },
        { keys: "Ctrl+Alt+3", desc: "Appliquer le style Titre 3" },
        { keys: "Ctrl+Maj+N", desc: "Appliquer le style Normal" },
        { keys: "Ctrl+Alt+Maj+S", desc: "Afficher le volet Styles" },
        { keys: "Ctrl+F", desc: "Volet de navigation / Rechercher" },
        { keys: "F9", desc: "Mettre à jour le champ ou la table sélectionnée" },
        { keys: "Maj+F5", desc: "Revenir à la dernière modification" }
      ],
      labs: [
        {
          id: "w03-lab1",
          title: "Lab 1 — Le rapport d'activité structuré",
          kind: "doc",
          scenario: `<p>Le directeur de <b>Pixel &amp; Co</b>, agence web de 25 personnes, a rédigé son rapport d'activité annuel. Les titres sont simplement en gras, un paragraphe est en Titre 2 par erreur et il n'y a pas de sommaire. Structurez le document avec les styles et ajoutez une table des matières à jour.</p>`,
          setup: {
            html: `<p><b>Rapport d'activité 2026 — Pixel &amp; Co</b></p>
<p><b>Synthèse</b></p>
<p>L'année 2026 a été marquée par une croissance soutenue de notre activité, portée par les refontes de sites e-commerce et le lancement de notre offre d'accompagnement en accessibilité numérique.</p>
<p>« Nos clients nous renouvellent leur confiance parce que nous livrons dans les délais. » — Sophie, cheffe de projet</p>
<p><b>Résultats</b></p>
<p><b>Chiffre d'affaires</b></p>
<p>Le chiffre d'affaires atteint 2,4 M€, en hausse de 18 % par rapport à 2025. La part des contrats récurrents (maintenance, hébergement) représente désormais 35 % du total.</p>
<h2>Ce paragraphe a été mis en Titre 2 par erreur alors qu'il s'agit de texte courant décrivant la répartition par secteur.</h2>
<p><b>Ressources humaines</b></p>
<p>Nous avons accueilli 6 nouveaux collaborateurs, dont 2 alternants. Le taux de rotation est resté stable à 8 %.</p>
<p><b>Perspectives</b></p>
<p>En 2027, nous visons l'ouverture d'une antenne à Bordeaux et la certification de l'équipe sur le référentiel RGAA.</p>`
          },
          tasks: [
            {
              text: "Appliquez le style <b>Titre 1</b> aux trois parties : « Synthèse », « Résultats », « Perspectives ».",
              hint: "Clic dans la ligne › Accueil › Styles › Titre 1 (ou Ctrl+Alt+1).",
              check: (a) => hasH(a, 1, "synthèse") && hasH(a, 1, "résultats") && hasH(a, 1, "perspectives")
            },
            {
              text: "Appliquez le style <b>Titre 2</b> à « Chiffre d'affaires » et « Ressources humaines ».",
              hint: "Ces sous-parties appartiennent à « Résultats ».",
              check: (a) => hasH(a, 2, "chiffre") && hasH(a, 2, "ressources humaines")
            },
            {
              text: "Remettez en style <b>Normal</b> le paragraphe passé en Titre 2 par erreur.",
              hint: "Clic dans ce paragraphe › Styles › Normal (Ctrl+Maj+N).",
              check: (a) => { const f = a.fmt("par erreur"); return !!(f.found && f.tag === "p"); }
            },
            {
              text: "Appliquez le style <b>Citation</b> au témoignage de Sophie.",
              hint: "Galerie des styles › Citation.",
              check: (a) => { const f = a.fmt("renouvellent leur confiance"); return !!(f.found && f.tag === "blockquote"); }
            },
            {
              text: "Insérez une <b>table des matières</b> sous le titre du rapport.",
              hint: "Curseur après la 1ʳᵉ ligne › Références › Table des matières.",
              check: (a) => !!a.toc
            },
            {
              text: "Vérifiez que la table des matières est <b>à jour</b> avec les 5 titres (mettez-la à jour si besoin).",
              hint: "Bouton « Mettre à jour la table » après avoir appliqué tous les styles.",
              check: (a) => !!(a.toc && a.tocFresh) && a.headings().length >= 5 && !a.headings().some((h) => norm(h.text).includes("par erreur"))
            }
          ],
          solution: `<ol>
<li>Clic dans « Synthèse » › <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>1</kbd> ; idem pour « Résultats » et « Perspectives ».</li>
<li>Clic dans « Chiffre d'affaires » › <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>2</kbd> ; idem « Ressources humaines ».</li>
<li>Clic dans le paragraphe « Ce paragraphe a été mis en Titre 2… » › style <i>Normal</i>.</li>
<li>Clic dans le témoignage › style <i>Citation</i>.</li>
<li>Curseur à la fin de « Rapport d'activité 2026 — Pixel &amp; Co » › <kbd>Entrée</kbd> › <i>Références › Table des matières</i>.</li>
<li><i>Mettre à jour la table</i> : elle doit lister Synthèse, Résultats (Chiffre d'affaires, Ressources humaines), Perspectives.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "D'où Word tire-t-il les entrées d'une table des matières automatique ?",
          choices: ["Des paragraphes en gras", "Des paragraphes ayant un style de titre (Titre 1, 2, 3…)", "Des lignes en majuscules", "Des signets"],
          answer: 1,
          explain: "La table automatique repose sur les styles de titre (ou les niveaux hiérarchiques). Un titre mis en gras à la main n'y apparaît pas."
        },
        {
          q: "Vous avez ajouté un chapitre, mais il n'apparaît pas dans la table des matières. Que faire ?",
          choices: ["Recréer la table", "Références › Mettre à jour la table › Mettre à jour toute la table", "Mettre à jour les numéros de page uniquement", "Enregistrer le document"],
          answer: 1,
          explain: "« Numéros de page uniquement » ne récupère pas les nouveaux titres : il faut mettre à jour toute la table."
        },
        {
          q: "Comment déplacer tout un chapitre (titre, sous-titres et texte) en un geste ?",
          choices: ["Couper/coller page par page", "Glisser son titre dans le volet de navigation", "Changer son style", "Insérer un saut de section"],
          answer: 1,
          explain: "Dans le volet de navigation, glisser un titre déplace toute la section qu'il contient."
        },
        {
          q: "Pourquoi laisser décochée l'option « Mettre à jour automatiquement » d'un style ?",
          choices: ["Elle ralentit Word", "Elle empêche l'impression", "Toute mise en forme manuelle sur un paragraphe modifierait le style partout", "Elle supprime la table des matières"],
          answer: 2,
          explain: "Avec cette option, un simple gras sur un titre se propage à tous les titres du même style."
        },
        {
          q: "Quelle est la bonne façon d'obtenir des titres numérotés 1, 1.1, 1.2, 2… ?",
          choices: ["Taper les numéros à la main", "Liste à plusieurs niveaux liée aux styles Titre 1 / Titre 2", "Liste à puces", "Tabulations"],
          answer: 1,
          explain: "Une liste à plusieurs niveaux associée aux styles de titre renumérote automatiquement quand on déplace ou insère une section."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 2 — w04 : Tableaux, colonnes, sauts, en-têtes
       =================================================================== */
    {
      id: "w04",
      level: 2,
      title: "Tableaux, colonnes, sauts de page, en-têtes et pieds de page",
      icon: "📐",
      duration: 55,
      summary: "Maîtrisez la mise en page : des tableaux propres, du texte sur plusieurs colonnes, des sauts maîtrisés et des en-têtes/pieds de page cohérents.",
      objectives: [
        "Être capable d'insérer, remplir et mettre en forme un tableau (ligne d'en-tête, styles, fusion)",
        "Être capable de répéter la ligne d'en-tête d'un tableau sur plusieurs pages",
        "Être capable de disposer du texte en colonnes et de forcer un saut de colonne ou de page",
        "Être capable de créer un en-tête et un pied de page avec numéro de page"
      ],
      lessons: [
        {
          title: "Créer et remplir un tableau",
          html: `<div class="o-path">Insertion › Tableaux › Tableau</div>
<p>Survolez la grille pour choisir colonnes × lignes, ou <i>Insérer un tableau…</i> pour saisir les dimensions et le comportement d'ajustement automatique.</p>
<ul>
<li><kbd>Tab</kbd> passe à la cellule suivante ; dans la <b>dernière cellule</b>, <kbd>Tab</kbd> ajoute une ligne.</li>
<li><kbd>Maj</kbd>+<kbd>Tab</kbd> revient à la cellule précédente (pour insérer une vraie tabulation dans une cellule : <kbd>Ctrl</kbd>+<kbd>Tab</kbd> dans Word).</li>
<li>Le <b>+</b> qui apparaît sur la bordure, entre deux lignes ou colonnes, insère une ligne/colonne à cet endroit.</li>
</ul>
<div class="o-path">Disposition › Lignes et colonnes › Insérer en dessous</div>
<p>L'onglet contextuel <b>Disposition</b> gère la structure : insérer/supprimer, <b>Fusionner les cellules</b>, <b>Fractionner</b>, <i>Ajustement automatique</i> (au contenu, à la fenêtre), <i>Distribuer les colonnes</i>, alignement dans la cellule et <b>Trier</b>.</p>
<div class="o-tip">💡 Vous avez une liste séparée par des points-virgules ou des tabulations ? Sélectionnez-la › <i>Insertion › Tableau › Convertir le texte en tableau</i>.</div>
<div class="callout bad">N'utilisez pas un tableau sans bordures pour faire de la mise en page « en colonnes » d'un texte courant : les lecteurs d'écran le liront cellule par cellule. Préférez les colonnes de la mise en page.</div>`
        },
        {
          title: "Mettre en forme un tableau",
          html: `<div class="o-path">Création de tableau › Styles de tableau</div>
<p>Un style de tableau applique bordures, trames et couleurs en un clic. Les cases à cocher <b>Options de style de tableau</b> activent : <i>Ligne d'en-tête</i>, <i>Ligne Total</i>, <i>Lignes à bandes</i>, <i>Première colonne</i>…</p>
<div class="o-path">Disposition › Données › Répéter les lignes d'en-tête</div>
<p>Indispensable pour un tableau qui tient sur plusieurs pages : la ligne de titres réapparaît en haut de chaque page. Elle signale aussi aux technologies d'assistance quelle ligne contient les en-têtes.</p>
<div class="callout warn">Ne recopiez jamais la ligne d'en-tête à la main en haut de la page 2 : au moindre ajout de ligne, elle se retrouvera au milieu du tableau.</div>
<p><i>Propriétés du tableau › Ligne</i> : décochez <b>Autoriser le fractionnement des lignes sur plusieurs pages</b> pour qu'une ligne haute ne soit jamais coupée en deux.</p>
<div class="o-tip">💡 Pour calculer un total simple : <i>Disposition › Données › Formule</i> propose <code>=SUM(ABOVE)</code>. Pour de vrais calculs, insérez plutôt un tableau Excel.</div>`
        },
        {
          title: "Sauts de page, de section et colonnes",
          html: `<div class="o-path">Mise en page › Sauts de page</div>
<table class="o-table"><thead><tr><th>Saut</th><th>Effet</th></tr></thead>
<tbody>
<tr><td>Page (<kbd>Ctrl</kbd>+<kbd>Entrée</kbd>)</td><td>Le texte suivant commence sur une nouvelle page</td></tr>
<tr><td>Colonne (<kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>Entrée</kbd>)</td><td>Passe en haut de la colonne suivante</td></tr>
<tr><td>Section — Page suivante</td><td>Nouvelle section sur une nouvelle page (orientation, marges, en-tête différents)</td></tr>
<tr><td>Section — Continu</td><td>Nouvelle section sur la même page (ex. passer en 2 colonnes au milieu)</td></tr>
</tbody></table>
<div class="o-path">Mise en page › Colonnes › Deux</div>
<p><i>Autres colonnes…</i> permet de régler la largeur, l'espacement et d'ajouter une <b>ligne séparatrice</b>. Si vous sélectionnez d'abord un passage, Word insère automatiquement des sauts de section continus autour.</p>
<div class="callout">Pour qu'un titre commence toujours en haut de page sans saut manuel : modifiez le style Titre 1 › <i>Format › Paragraphe › Enchaînements › Saut de page avant</i>.</div>
<div class="o-tip">💡 Affichez les marques (¶) pour voir et supprimer un saut : il apparaît comme « ……Saut de page…… ».</div>`
        },
        {
          title: "En-têtes, pieds de page et numéros de page",
          html: `<div class="o-path">Insertion › En-tête et pied de page › En-tête › Modifier l'en-tête</div>
<p>Ou plus simplement : <b>double-clic</b> dans la marge haute ou basse. L'onglet contextuel <b>En-tête et pied de page</b> apparaît ; double-clic dans le corps du texte (ou <i>Fermer l'en-tête et le pied de page</i>) pour revenir.</p>
<div class="o-path">En-tête et pied de page › Numéro de page › Bas de page › Numéro normal 2</div>
<ul>
<li>« Page X sur Y » : galerie <i>Bas de page › Page X sur Y</i> (champs PAGE et NUMPAGES).</li>
<li><b>Première page différente</b> : pas d'en-tête sur la page de garde.</li>
<li><b>Pages paires et impaires différentes</b> : documents imprimés en recto verso.</li>
<li><i>Format des numéros de page</i> : chiffres romains, <b>À partir de</b> 1.</li>
</ul>
<div class="callout warn">Le bouton <b>Lier au précédent</b> (groupe Navigation) relie l'en-tête d'une section à celui de la précédente. Pour un en-tête différent par chapitre, insérez un saut de section, puis <b>désactivez « Lier au précédent »</b> <i>avant</i> de modifier le texte.</div>
<div class="callout bad">Taper le numéro de page à la main dans le pied de page : il affichera « 1 » sur toutes les pages. Utilisez toujours la commande Numéro de page.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+Entrée", desc: "Saut de page" },
        { keys: "Ctrl+Maj+Entrée", desc: "Saut de colonne" },
        { keys: "Tab", desc: "Cellule suivante (ajoute une ligne en fin de tableau)" },
        { keys: "Maj+Tab", desc: "Cellule précédente" },
        { keys: "Alt+Maj+Flèche haut", desc: "Monter la ligne de tableau / le paragraphe" },
        { keys: "Alt+Maj+Flèche bas", desc: "Descendre la ligne de tableau / le paragraphe" }
      ],
      labs: [
        {
          id: "w04-lab1",
          title: "Lab 1 — Le catalogue de formations",
          kind: "doc",
          scenario: `<p>L'organisme <b>Atelier Numérique</b> prépare son catalogue 2027. Il manque le tableau récapitulatif des sessions, l'en-tête et le pied de page, et les conditions générales doivent démarrer sur une nouvelle page. Le texte de présentation sera lu plus facilement sur deux colonnes.</p>`,
          setup: {
            html: `<h1>Catalogue des formations 2027</h1>
<p>Depuis 2012, Atelier Numérique accompagne les PME de la région dans la montée en compétences de leurs équipes. Nos formations courtes, en présentiel ou à distance, sont animées par des formateurs certifiés et finançables par votre OPCO. Chaque session accueille au maximum huit stagiaires afin de garantir un suivi individualisé et beaucoup de pratique.</p>
<p>Nos parcours couvrent la bureautique, la communication digitale et la gestion de projet. Ils s'adressent aussi bien aux débutants qu'aux utilisateurs confirmés qui souhaitent gagner en efficacité au quotidien.</p>
<h2>Sessions programmées</h2>
<p>Insérez le tableau ici.</p>
<h2>Conditions générales</h2>
<p>Toute inscription est ferme après signature de la convention de formation. Une annulation moins de 10 jours ouvrés avant la session donne lieu à une facturation de 50 % du montant.</p>`
          },
          tasks: [
            {
              text: "Remplacez « Insérez le tableau ici. » par un <b>tableau de 3 colonnes et au moins 4 lignes</b> : Formation / Durée / Tarif HT, puis 3 sessions (ex. Word expert — 2 jours — 790 €).",
              hint: "Insertion › Tableau, puis Tab pour passer de cellule en cellule.",
              check: (a) => a.tables().some((t) => t.cols === 3 && t.rows >= 4) && a.$$("table").some((t) => norm(t.textContent).includes("tarif"))
            },
            {
              text: "Mettez la <b>ligne d'en-tête</b> du tableau en gras.",
              hint: "Sélectionnez la 1ʳᵉ ligne du tableau › Ctrl+G.",
              check: (a) => a.tables().some((t) => t.cols === 3 && t.headerBold)
            },
            {
              text: "Insérez un <b>saut de page</b> pour que « Conditions générales » commence sur une nouvelle page.",
              hint: "Curseur juste avant « Conditions générales » › Ctrl+Entrée.",
              check: (a) => a.count("hr.o-pagebreak") >= 1
            },
            {
              text: "Ajoutez un <b>en-tête</b> contenant « Atelier Numérique — Catalogue 2027 ».",
              hint: "Double-clic dans la zone d'en-tête de la page.",
              check: (a) => norm(a.header).includes("catalogue 2027") && norm(a.header).includes("atelier")
            },
            {
              text: "Ajoutez un <b>pied de page</b> contenant « contact@atelier-numerique.fr ».",
              hint: "Double-clic dans la zone de pied de page.",
              check: (a) => norm(a.footer).replace(/\s/g, "").includes("contact@atelier-numerique.fr")
            },
            {
              text: "Disposez le texte sur <b>2 colonnes</b>.",
              hint: "Mise en page › Colonnes › Deux.",
              check: (a) => a.columns === 2
            }
          ],
          solution: `<ol>
<li>Sélectionnez « Insérez le tableau ici. » et supprimez-le › <i>Insertion › Tableau</i> 3 × 4 › saisissez Formation | Durée | Tarif HT, puis par ex. Word expert | 2 jours | 790 € ; Excel tableaux croisés | 1 jour | 420 € ; Réseaux sociaux pro | 2 jours | 690 €.</li>
<li>Sélectionnez la 1ʳᵉ ligne › <kbd>Ctrl</kbd>+<kbd>G</kbd> (dans le vrai Word : cochez aussi <i>Ligne d'en-tête</i> et <i>Répéter les lignes d'en-tête</i>).</li>
<li>Curseur au début de « Conditions générales » › <kbd>Ctrl</kbd>+<kbd>Entrée</kbd>.</li>
<li>Double-clic dans l'en-tête › « Atelier Numérique — Catalogue 2027 ».</li>
<li>Double-clic dans le pied de page › « contact@atelier-numerique.fr ».</li>
<li><i>Mise en page › Colonnes › Deux</i>.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Votre tableau tient sur 3 pages. Comment afficher la ligne de titres en haut de chaque page ?",
          choices: ["La copier à la main en haut de chaque page", "Disposition › Données › Répéter les lignes d'en-tête", "Insérer un en-tête de page", "Fusionner les cellules"],
          answer: 1,
          explain: "« Répéter les lignes d'en-tête » est dynamique et indique aussi la ligne d'en-tête aux lecteurs d'écran."
        },
        {
          q: "Que se passe-t-il quand on appuie sur Tab dans la dernière cellule d'un tableau ?",
          choices: ["Rien", "Une tabulation est insérée", "Une nouvelle ligne est ajoutée", "On sort du tableau"],
          answer: 2,
          explain: "Tab dans la dernière cellule crée une nouvelle ligne : pratique pour remplir un tableau au clavier."
        },
        {
          q: "Vous voulez une page en orientation paysage au milieu d'un document portrait. Que faut-il ?",
          choices: ["Un saut de page", "Un saut de colonne", "Des sauts de section (page suivante) avant et après", "Un tableau"],
          answer: 2,
          explain: "L'orientation est une propriété de section : il faut isoler la page dans sa propre section."
        },
        {
          q: "Le chapitre 2 doit avoir un en-tête différent du chapitre 1, mais en le modifiant, le chapitre 1 change aussi. Pourquoi ?",
          choices: ["L'en-tête est verrouillé", "« Lier au précédent » est toujours activé dans la section 2", "Il faut un saut de page", "Le document est protégé"],
          answer: 1,
          explain: "Après le saut de section, désactivez « Lier au précédent » dans l'en-tête de la section 2 avant de le modifier."
        },
        {
          q: "Quel raccourci insère un saut de colonne ?",
          choices: ["Ctrl+Entrée", "Maj+Entrée", "Ctrl+Maj+Entrée", "Alt+Entrée"],
          answer: 2,
          explain: "Ctrl+Entrée = saut de page, Maj+Entrée = saut de ligne, Ctrl+Maj+Entrée = saut de colonne."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 3 — w05 : Documents longs
       =================================================================== */
    {
      id: "w05",
      level: 3,
      title: "Documents longs : sections, notes, légendes, renvois et index",
      icon: "📚",
      duration: 70,
      summary: "Rapports, mémoires, manuels : les outils qui rendent un document de 100 pages fiable, navigable et facile à maintenir.",
      objectives: [
        "Être capable de découper un document en sections avec numérotation et en-têtes indépendants",
        "Être capable d'insérer notes de bas de page et notes de fin",
        "Être capable de légender figures et tableaux et d'y faire référence par des renvois dynamiques",
        "Être capable de générer une table des illustrations et un index",
        "Être capable d'insérer des liens hypertextes et des signets"
      ],
      lessons: [
        {
          title: "Architecture d'un document long",
          html: `<p>Un document long type se compose de sections distinctes :</p>
<table class="o-table"><thead><tr><th>Section</th><th>Contenu</th><th>Numérotation</th></tr></thead>
<tbody>
<tr><td>1</td><td>Page de garde (Insertion › Pages › Page de garde)</td><td>Aucune (Première page différente)</td></tr>
<tr><td>2</td><td>Sommaire, remerciements</td><td>Romains i, ii, iii</td></tr>
<tr><td>3</td><td>Corps du document</td><td>Arabes à partir de 1</td></tr>
<tr><td>4</td><td>Annexes (souvent en paysage)</td><td>A-1, A-2…</td></tr>
</tbody></table>
<div class="o-path">Mise en page › Sauts de page › Sauts de section › Page suivante</div>
<p>Puis dans le pied de page de la section 3 : désactivez <b>Lier au précédent</b>, <i>Numéro de page › Format des numéros de page › À partir de : 1</i>.</p>
<div class="o-tip">💡 Le champ <code>STYLEREF</code> (Insertion › QuickPart › Champ › StyleRef › Titre 1) affiche automatiquement dans l'en-tête le titre du chapitre en cours, comme dans un livre.</div>
<div class="callout ok">Travaillez toujours avec les styles et le volet de navigation activé : c'est la colonne vertébrale du document long.</div>`
        },
        {
          title: "Notes de bas de page et notes de fin",
          html: `<div class="o-path">Références › Notes de bas de page › Insérer une note de bas de page</div>
<p>Word place un appel de note numéroté (¹) au point d'insertion et ouvre la note en bas de la page. Les numéros se recalculent automatiquement si vous insérez une note avant.</p>
<ul>
<li><b>Note de bas de page</b> (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>F</kbd>) : en bas de la page concernée.</li>
<li><b>Note de fin</b> (<kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>D</kbd>) : regroupées en fin de document ou de section.</li>
<li>Double-clic sur le numéro de la note pour revenir à l'appel dans le texte.</li>
</ul>
<p>Le lanceur du groupe <i>Notes de bas de page</i> règle le format (1, 2, 3 / a, b, c / *), la numérotation continue ou recommençant à chaque section/page, et permet de <b>convertir</b> notes de bas de page ↔ notes de fin.</p>
<div class="callout bad">Ne tapez jamais un « ¹ » en exposant suivi d'un texte en bas de page : au premier ajout de paragraphe, la note se retrouvera sur la mauvaise page.</div>
<div class="callout warn">Pour supprimer une note, supprimez l'<b>appel</b> dans le texte (le petit chiffre), pas le texte de la note : sinon le numéro reste orphelin.</div>`
        },
        {
          title: "Légendes, renvois et table des illustrations",
          html: `<div class="o-path">Références › Légendes › Insérer une légende</div>
<p>Sélectionnez l'image ou le tableau, choisissez l'étiquette (<i>Figure</i>, <i>Tableau</i>, <i>Équation</i> ou <i>Nouvelle étiquette…</i>) et la position (au-dessus pour un tableau, en dessous pour une figure, par convention). Word génère « Figure 3 » avec un champ <code>SEQ</code>.</p>
<div class="o-path">Références › Légendes › Renvoi</div>
<p>Un <b>renvoi</b> insère « voir Figure 3, page 12 » sous forme de champ : si la figure devient la n° 4, le texte suit. Type de référence : Figure, Titre, Signet, Note de bas de page… ; Insérer un renvoi à : <i>Étiquette et numéro uniquement</i>, <i>Numéro de page</i>, <i>Au-dessus/Au-dessous</i>.</p>
<div class="callout warn">Les renvois et numéros de légende sont des <b>champs</b> : ils ne se mettent à jour qu'à l'impression ou sur demande. Avant de diffuser : <kbd>Ctrl</kbd>+<kbd>A</kbd> puis <kbd>F9</kbd>. Un « Erreur ! Source du renvoi introuvable » signale que la cible a été supprimée.</div>
<div class="o-path">Références › Légendes › Insérer une table des illustrations</div>
<p>Elle fonctionne comme une table des matières, mais à partir des légendes d'une étiquette donnée.</p>
<div class="o-tip">💡 <i>Insertion › Liens › Signet</i> nomme un endroit précis (ex. <code>Tarifs2027</code>) : vous pouvez ensuite y faire un renvoi ou un lien hypertexte interne.</div>`
        },
        {
          title: "Index et liens hypertextes",
          html: `<div class="o-path">Références › Index › Entrée</div>
<p>Sélectionnez un terme puis <kbd>Alt</kbd>+<kbd>Maj</kbd>+<kbd>X</kbd> : la boîte <i>Marquer les entrées d'index</i> s'ouvre. <b>Marquer tout</b> repère toutes les occurrences. Word insère des champs <code>{ XE "terme" }</code> visibles uniquement avec ¶ affiché.</p>
<div class="o-path">Références › Index › Insérer l'index</div>
<p>Choisissez le nombre de colonnes, la mise en forme et l'alignement des numéros. Après de nouvelles entrées : <i>Mettre à jour l'index</i>.</p>
<div class="callout">Pensez à masquer les marques de paragraphe (¶) avant d'insérer l'index ou la table : les champs XE visibles faussent la pagination.</div>
<div class="o-path">Insertion › Liens › Lien</div>
<p><kbd>Ctrl</kbd>+<kbd>K</kbd> crée un lien vers une page web, un fichier, une adresse e-mail ou un <b>emplacement dans ce document</b> (titres et signets). Rédigez un texte de lien explicite (« Consulter le guide RGPD de la CNIL ») plutôt que « cliquez ici ».</p>
<div class="o-tip">💡 <kbd>Ctrl</kbd>+clic suit un lien dans Word. Pour cliquer directement : <i>Fichier › Options › Options avancées</i> › décochez <i>Utiliser CTRL+clic pour suivre un lien hypertexte</i>.</div>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+Alt+F", desc: "Insérer une note de bas de page" },
        { keys: "Ctrl+Alt+D", desc: "Insérer une note de fin" },
        { keys: "Alt+Maj+X", desc: "Marquer une entrée d'index" },
        { keys: "Ctrl+K", desc: "Insérer un lien hypertexte" },
        { keys: "F9", desc: "Mettre à jour les champs sélectionnés" },
        { keys: "Ctrl+Maj+F9", desc: "Convertir un champ en texte figé" },
        { keys: "F5", desc: "Atteindre (page, section, signet…)" }
      ],
      labs: [
        {
          id: "w05-lab1",
          title: "Lab 1 — Le mémoire de conformité RGPD",
          kind: "doc",
          scenario: `<p>Vous finalisez pour la <b>Clinique vétérinaire des Tilleuls</b> un mémoire sur sa mise en conformité RGPD. Il faut le rendre « professionnel » : sommaire à jour, chaque partie sur une nouvelle page, notes de bas de page pour les références juridiques, un lien vers la CNIL, en-tête et pied de page.</p>`,
          setup: {
            html: `<p><b>Mise en conformité RGPD — Clinique vétérinaire des Tilleuls</b></p>
<h1>Contexte</h1>
<p>La clinique traite chaque jour des données personnelles de propriétaires d'animaux : identité, coordonnées, historique des paiements. Le règlement général sur la protection des données impose de tenir un registre des traitements.</p>
<h2>Périmètre</h2>
<p>Le périmètre couvre le logiciel de gestion des rendez-vous, la messagerie et les dossiers papier archivés à l'accueil.</p>
<h1>Analyse des écarts</h1>
<p>Les durées de conservation ne sont pas définies et les consentements aux rappels de vaccination par SMS ne sont pas tracés.</p>
<h2>Plan d'action</h2>
<p>Désigner un référent, rédiger le registre, mettre à jour les mentions d'information et paramétrer la purge automatique des dossiers inactifs depuis cinq ans.</p>`
          },
          tasks: [
            {
              text: "Insérez une <b>table des matières</b> sous le titre du mémoire.",
              hint: "Curseur après la 1ʳᵉ ligne › Références › Table des matières.",
              check: (a) => !!a.toc
            },
            {
              text: "Insérez des <b>sauts de page</b> pour que « Contexte » et « Analyse des écarts » commencent chacun sur une nouvelle page (au moins 2 sauts).",
              hint: "Curseur au début de chaque Titre 1 › Ctrl+Entrée.",
              check: (a) => a.count("hr.o-pagebreak") >= 2
            },
            {
              text: "Ajoutez <b>2 notes de bas de page</b> : après « données personnelles » (« Article 4 du RGPD ») et après « registre des traitements » (« Article 30 du RGPD »).",
              hint: "Références › Insérer une note de bas de page (Ctrl+Alt+F).",
              check: (a) => a.count("sup.o-fn") >= 2
            },
            {
              text: "Transformez « registre des traitements » (ou une phrase de votre choix) en <b>lien hypertexte</b> vers <code>https://www.cnil.fr</code>.",
              hint: "Sélectionnez le texte › Insertion › Lien (Ctrl+K).",
              check: (a) => a.$$("a[href]").some((l) => /cnil\.fr/i.test(l.getAttribute("href") || ""))
            },
            {
              text: "En-tête : « Mémoire RGPD » ; pied de page : « Diffusion restreinte ».",
              hint: "Double-clic dans les zones d'en-tête et de pied de page.",
              check: (a) => norm(a.header).includes("mémoire rgpd") && norm(a.footer).includes("diffusion restreinte")
            },
            {
              text: "Ajoutez à la fin un <b>Titre 1</b> « Annexes », puis <b>mettez à jour</b> la table des matières.",
              hint: "Nouveau paragraphe › Ctrl+Alt+1 › puis « Mettre à jour la table ».",
              check: (a) => hasH(a, 1, "annexes") && !!(a.toc && a.tocFresh)
            }
          ],
          solution: `<ol>
<li>Curseur à la fin de la 1ʳᵉ ligne › <kbd>Entrée</kbd> › <i>Références › Table des matières</i>.</li>
<li>Curseur au début de « Contexte » › <kbd>Ctrl</kbd>+<kbd>Entrée</kbd> ; idem au début de « Analyse des écarts ». (Dans le vrai Word, l'option <i>Saut de page avant</i> du style Titre 1 automatise cela.)</li>
<li>Curseur après « données personnelles » › <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>F</kbd> › « Article 4 du RGPD » ; idem après « registre des traitements » › « Article 30 du RGPD ».</li>
<li>Sélectionnez « registre des traitements » › <kbd>Ctrl</kbd>+<kbd>K</kbd> › <code>https://www.cnil.fr</code>.</li>
<li>Double-clic en haut › « Mémoire RGPD » ; en bas › « Diffusion restreinte ».</li>
<li>À la fin : <kbd>Entrée</kbd> › « Annexes » › <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>1</kbd> › <i>Mettre à jour la table</i>.</li>
</ol>`
        },
        {
          id: "w05-lab2",
          title: "Lab 2 — Légendes, renvois et index dans le vrai Word",
          kind: "steps",
          scenario: `<p>Les légendes, renvois et index reposent sur des champs que seul le vrai Word gère. Téléchargez le fichier de départ, collez son contenu dans un nouveau document Word et suivez les étapes.</p>`,
          setup: {
            app: "Word",
            file: {
              name: "manuel-entrepot.txt",
              type: "txt",
              content: "Manuel d'accueil — Entrepôt de Saint-Priest\n\nPlan de l'entrepôt\nL'entrepôt est divisé en quatre zones : réception, stockage, préparation, expédition. Le schéma ci-dessous présente leur disposition.\n[Insérer ici une forme ou une image représentant le plan]\n\nHoraires des équipes\nLes horaires sont résumés dans le tableau ci-dessous.\nÉquipe;Début;Fin\nMatin;5h00;13h00\nAprès-midi;13h00;21h00\nNuit;21h00;5h00\n\nSécurité\nLe port des chaussures de sécurité est obligatoire dans toutes les zones. Le chariot élévateur est réservé aux caristes titulaires du CACES. Tout incident doit être signalé au chef d'équipe.\n"
            }
          },
          tasks: [
            { text: "Mettez « Plan de l'entrepôt », « Horaires des équipes » et « Sécurité » en <b>Titre 1</b>. Convertissez les 4 lignes séparées par « ; » en tableau (<i>Insertion › Tableau › Convertir le texte en tableau</i>, séparateur : Autre « ; »).", hint: "Sélectionnez exactement les 4 lignes Équipe…Nuit avant de convertir." },
            { text: "Insérez une image ou une forme pour le plan, puis <i>Références › Insérer une légende</i> : « Figure 1 — Plan des zones » (en dessous). Légendez le tableau : « Tableau 1 — Horaires » (au-dessus).", hint: "Sélectionnez d'abord l'objet à légender." },
            { text: "Remplacez « ci-dessous » dans chaque paragraphe par un <b>renvoi</b> : <i>Références › Renvoi</i> › Type Figure / Tableau › « Étiquette et numéro uniquement ».", hint: "Le texte devient « présente leur disposition (voir Figure 1) »." },
            { text: "Marquez comme entrées d'index « CACES », « chariot élévateur » et « chaussures de sécurité » (<kbd>Alt</kbd>+<kbd>Maj</kbd>+<kbd>X</kbd>, <i>Marquer tout</i>), puis insérez l'index en fin de document.", hint: "Masquez les ¶ avant d'insérer l'index." },
            { text: "Insérez une nouvelle figure <b>avant</b> la figure 1, légendez-la, puis <kbd>Ctrl</kbd>+<kbd>A</kbd> › <kbd>F9</kbd> : vérifiez que l'ancienne figure et son renvoi sont devenus « Figure 2 ».", hint: "C'est tout l'intérêt des champs : rien à renuméroter à la main." }
          ],
          solution: `<ol>
<li>Styles Titre 1 sur les 3 titres ; sélection des lignes « Équipe;Début;Fin » à « Nuit;21h00;5h00 » › <i>Convertir le texte en tableau</i> › Séparer le texte au niveau des : Autre « ; » › 3 colonnes.</li>
<li>Sélection de l'image › <i>Références › Légendes › Insérer une légende</i> › Étiquette Figure › Position Au-dessous de la sélection › ajoutez « — Plan des zones ». Pour le tableau : Étiquette Tableau, Au-dessus.</li>
<li>Sélection de « ci-dessous » › <i>Renvoi</i> › Catégorie Figure › Insérer un renvoi à : Étiquette et numéro uniquement › Figure 1 › Insérer.</li>
<li>Sélection « CACES » › <kbd>Alt</kbd>+<kbd>Maj</kbd>+<kbd>X</kbd> › Marquer tout ; idem pour les autres termes. Fin du document › <i>Références › Index › Insérer l'index</i>.</li>
<li>Après insertion d'une nouvelle figure légendée, <kbd>Ctrl</kbd>+<kbd>A</kbd> puis <kbd>F9</kbd> : les numéros SEQ et les renvois se recalculent.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Pour numéroter le corps du document à partir de 1 après un sommaire en chiffres romains, il faut :",
          choices: ["Un saut de page et un numéro tapé à la main", "Un saut de section, désactiver « Lier au précédent », Format des numéros › À partir de 1", "Deux documents séparés", "Supprimer le sommaire"],
          answer: 1,
          explain: "Le format et le point de départ de la numérotation sont propres à chaque section."
        },
        {
          q: "Vous déplacez un tableau : sa légende devient « Tableau 2 » mais le texte dit encore « voir Tableau 3 ». Que faire ?",
          choices: ["Corriger le texte à la main", "Ctrl+A puis F9 pour mettre à jour les champs", "Supprimer la légende", "Enregistrer en PDF"],
          answer: 1,
          explain: "Les renvois sont des champs : ils se recalculent à la mise à jour (F9) ou à l'impression."
        },
        {
          q: "Quelle est la bonne méthode pour supprimer une note de bas de page ?",
          choices: ["Effacer le texte de la note en bas de page", "Supprimer l'appel de note dans le texte", "Passer la note en blanc", "Supprimer le pied de page"],
          answer: 1,
          explain: "Supprimer l'appel (le chiffre dans le texte) supprime la note et renumérote les suivantes."
        },
        {
          q: "Que signifie le message « Erreur ! Source du renvoi introuvable » ?",
          choices: ["Le document est corrompu", "La cible du renvoi (légende, titre, signet) a été supprimée", "L'imprimante est déconnectée", "Le renvoi n'a pas encore été enregistré"],
          answer: 1,
          explain: "Le champ REF pointe vers un signet qui n'existe plus : recréez le renvoi vers la bonne cible."
        },
        {
          q: "Quel raccourci ouvre la boîte « Marquer les entrées d'index » ?",
          choices: ["Ctrl+Alt+F", "Alt+Maj+X", "Ctrl+K", "Alt+F9"],
          answer: 1,
          explain: "Alt+Maj+X marque une entrée d'index ; Ctrl+Alt+F insère une note de bas de page."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 3 — w06 : Révision et collaboration
       =================================================================== */
    {
      id: "w06",
      level: 3,
      title: "Révision et collaboration : commentaires, suivi des modifications, comparer, protéger",
      icon: "🤝",
      duration: 60,
      summary: "Relisez à plusieurs sans jamais perdre la trace de qui a changé quoi, puis verrouillez la version finale.",
      objectives: [
        "Être capable d'insérer, répondre et résoudre des commentaires",
        "Être capable d'activer le suivi des modifications et d'accepter ou refuser chaque changement",
        "Être capable de comparer deux versions d'un document",
        "Être capable de restreindre la modification et de nettoyer un document avant diffusion"
      ],
      lessons: [
        {
          title: "Commentaires",
          html: `<div class="o-path">Révision › Commentaires › Nouveau commentaire</div>
<p>Sélectionnez le passage concerné puis <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>M</kbd>. Le commentaire s'affiche dans la marge, avec votre nom (défini dans <i>Fichier › Options › Général › Nom d'utilisateur</i>).</p>
<ul>
<li><b>@mention</b> : tapez <code>@Prénom</code> dans un commentaire d'un fichier OneDrive/SharePoint : la personne reçoit un e-mail.</li>
<li><b>Répondre</b> : créez un fil de discussion plutôt qu'un nouveau commentaire.</li>
<li><b>Résoudre</b> : marque le sujet comme traité sans le supprimer (on peut le rouvrir).</li>
<li><b>Supprimer › Supprimer tous les commentaires du document</b> avant diffusion externe.</li>
</ul>
<div class="o-path">Révision › Commentaires › Afficher les commentaires › Liste</div>
<div class="callout ok">Un bon commentaire est actionnable : « Remplacer 30 jours par 45 jours, conformément à la loi LME ? » plutôt que « bof ».</div>
<div class="callout warn">Un commentaire n'est pas une modification : si vous voulez changer le texte, modifiez-le avec le suivi activé, et réservez le commentaire à l'explication.</div>`
        },
        {
          title: "Le suivi des modifications",
          html: `<div class="o-path">Révision › Suivi › Suivi des modifications</div>
<p>Une fois activé (<kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>E</kbd>), chaque insertion apparaît soulignée en couleur, chaque suppression barrée, avec un trait dans la marge. Chaque relecteur a sa couleur.</p>
<table class="o-table"><thead><tr><th>Affichage</th><th>Pour…</th></tr></thead>
<tbody>
<tr><td>Marques simples</td><td>Lire le texte modifié, un trait rouge signale les zones changées</td></tr>
<tr><td>Toutes les marques</td><td>Voir chaque insertion/suppression</td></tr>
<tr><td>Aucune marque</td><td>Voir le résultat final (les marques existent toujours !)</td></tr>
<tr><td>Original</td><td>Voir le texte avant modifications</td></tr>
</tbody></table>
<div class="o-path">Révision › Modifications › Accepter › Accepter et passer à la suivante</div>
<p>Parcourez avec <i>Précédente</i> / <i>Suivante</i>, acceptez ou refusez une par une, ou <i>Accepter toutes les modifications</i>. Le <b>volet Révisions</b> liste tout. <i>Afficher les marques › Personnes</i> filtre par relecteur.</p>
<div class="callout bad">Erreur classique : passer en « Aucune marque » et envoyer le document. Le destinataire verra toutes les modifications (et les commentaires) en l'ouvrant. Il faut <b>accepter/refuser</b> pour les faire disparaître.</div>
<div class="o-tip">💡 <i>Suivi des modifications › Verrouiller le suivi</i> (avec mot de passe) empêche un relecteur de désactiver le suivi.</div>`
        },
        {
          title: "Comparer et combiner des versions",
          html: `<p>Un partenaire vous renvoie le contrat modifié… <b>sans</b> suivi des modifications ? Pas de panique.</p>
<div class="o-path">Révision › Comparer › Comparer…</div>
<p>Choisissez le <b>Document original</b> et le <b>Document révisé</b>, indiquez le nom sous lequel attribuer les écarts. Word produit un troisième document où toutes les différences apparaissent comme des modifications suivies, avec les deux versions affichées côte à côte.</p>
<p><b>Combiner</b> fusionne les révisions de plusieurs relecteurs (chacun ayant travaillé sur une copie) dans un seul document.</p>
<div class="callout">Le bouton <i>Plus &gt;&gt;</i> de la boîte Comparer permet d'ignorer certains écarts (mise en forme, espaces, en-têtes) pour se concentrer sur le fond.</div>
<div class="o-tip">💡 Pour une co-rédaction en temps réel, enregistrez le fichier sur OneDrive/SharePoint et cliquez sur <b>Partager</b> : chacun voit le curseur des autres, sans jamais avoir à comparer de versions.</div>`
        },
        {
          title: "Protéger et préparer la diffusion",
          html: `<div class="o-path">Révision › Protéger › Restreindre la modification</div>
<ol>
<li><i>Restrictions de mise en forme</i> : limiter aux styles autorisés.</li>
<li><i>Restrictions de modifications</i> : <b>Modifications suivies</b>, <b>Commentaires</b>, <b>Remplissage de formulaires</b> ou <b>Aucune modification (Lecture seule)</b>. Des exceptions peuvent être accordées à certaines zones.</li>
<li><b>Oui, activer la protection</b> (mot de passe facultatif).</li>
</ol>
<div class="o-path">Fichier › Informations › Protéger le document</div>
<p>Autres options : <i>Toujours ouvrir en lecture seule</i>, <b>Chiffrer avec mot de passe</b> (impossible à ouvrir sans le mot de passe), <i>Marquer comme final</i> (simple avertissement, pas une sécurité).</p>
<div class="callout warn">La « Restriction de la modification » n'est pas un chiffrement : elle évite les erreurs, pas un utilisateur déterminé. Pour la confidentialité, utilisez <i>Chiffrer avec mot de passe</i> — et ne perdez pas ce mot de passe, Microsoft ne peut pas le récupérer.</div>
<div class="o-path">Fichier › Informations › Vérifier l'absence de problèmes › Inspecter le document</div>
<p>L'Inspecteur détecte et supprime commentaires, révisions, propriétés (auteur), texte masqué, en-têtes… avant d'envoyer un document à l'extérieur.</p>`
        }
      ],
      shortcuts: [
        { keys: "Ctrl+Alt+M", desc: "Nouveau commentaire" },
        { keys: "Ctrl+Maj+E", desc: "Activer/désactiver le suivi des modifications" },
        { keys: "F7", desc: "Orthographe et grammaire (Éditeur)" },
        { keys: "Maj+F7", desc: "Dictionnaire des synonymes" },
        { keys: "Ctrl+Alt+F", desc: "Insérer une note de bas de page (référence de relecture)" }
      ],
      labs: [
        {
          id: "w06-lab1",
          title: "Lab 1 — Relecture du contrat de prestation",
          kind: "doc",
          scenario: `<p>Juriste chez <b>Événements Horizon</b>, vous relisez un projet de contrat de prestation avec un traiteur. Vous devez signaler les clauses à discuter par des commentaires, corriger le délai de paiement et marquer comme « à supprimer » une clause d'exclusivité déséquilibrée (dans le simulateur, le texte barré représente une suppression suivie).</p>`,
          setup: {
            html: `<h1>Contrat de prestation de services — traiteur</h1>
<p>Entre la société Événements Horizon, ci-après « le Client », et la société Saveurs du Rhône, ci-après « le Prestataire ».</p>
<h2>Article 3 — Prix et paiement</h2>
<p>Le prix forfaitaire est fixé à 18 500 € HT pour 250 couverts. Les factures sont payables à 90 jours fin de mois à compter de leur émission.</p>
<h2>Article 5 — Annulation</h2>
<p>En cas d'annulation par le Client moins de 30 jours avant l'événement, des pénalités égales à 100 % du prix resteront dues au Prestataire.</p>
<h2>Article 7 — Exclusivité</h2>
<p>Le Client s'engage à confier au Prestataire l'ensemble de ses événements pendant cinq ans. Cette clause d'exclusivité ne peut être résiliée par le Client.</p>
<h2>Article 9 — Responsabilité</h2>
<p>Le Prestataire est responsable de la conformité sanitaire des repas servis.</p>`
          },
          tasks: [
            {
              text: "Ajoutez un <b>commentaire</b> sur le mot « pénalités » (article 5) demandant de les ramener à 50 %.",
              hint: "Sélectionnez « pénalités » › Révision › Nouveau commentaire.",
              check: (a) => a.$$("mark.o-comment").some((m) => norm(m.textContent).includes("pénalit") && norm(m.getAttribute("data-note")).length > 0)
            },
            {
              text: "Ajoutez au moins <b>un second commentaire</b>, sur un autre passage de votre choix (ex. l'article 9).",
              hint: "Un commentaire = une remarque actionnable.",
              check: (a) => a.$$("mark.o-comment").filter((m) => norm(m.getAttribute("data-note")).length > 0).length >= 2
            },
            {
              text: "Corrigez le délai de paiement : remplacez « 90 jours » par « 45 jours » (délai maximal légal fin de mois).",
              hint: "Rechercher/Remplacer, ou sélection et saisie.",
              check: (a) => a.has("45 jours fin de mois") && !a.has("90 jours")
            },
            {
              text: "<s>Barrez</s> la phrase « Cette clause d'exclusivité ne peut être résiliée par le Client. » pour signaler sa suppression.",
              hint: "Sélectionnez la phrase › bouton Barré (abc) du groupe Police.",
              check: (a) => { const f = a.fmt("ne peut être résiliée"); return !!(f.found && f.strike); }
            },
            {
              text: "Ajoutez en fin de document la phrase « Version relue le 25/10/2026 » en <i>italique</i>.",
              hint: "Nouveau paragraphe à la fin, saisie, puis Ctrl+I.",
              check: (a) => { const f = a.fmt("relue le 25/10/2026"); return !!(f.found && f.italic); }
            }
          ],
          solution: `<ol>
<li>Sélectionnez « pénalités » › <i>Nouveau commentaire</i> › « Proposer 50 % du prix, 100 % est excessif. »</li>
<li>Sélectionnez par ex. « conformité sanitaire » › commentaire « Préciser l'assurance RC du Prestataire. »</li>
<li>Sélectionnez « 90 » › tapez « 45 » (dans le vrai Word, avec le suivi activé : « 90 » apparaît barré, « 45 » souligné).</li>
<li>Sélectionnez la dernière phrase de l'article 7 › <i>Accueil › Police › Barré</i>.</li>
<li>À la fin : <kbd>Entrée</kbd> › « Version relue le 25/10/2026 » › <kbd>Ctrl</kbd>+<kbd>I</kbd>.</li>
</ol>`
        },
        {
          id: "w06-lab2",
          title: "Lab 2 — Suivi, comparaison et protection dans le vrai Word",
          kind: "steps",
          scenario: `<p>Le suivi des modifications, la comparaison et la protection ne se simulent pas fidèlement : pratiquez-les dans Word avec le fichier de départ.</p>`,
          setup: {
            app: "Word",
            file: {
              name: "charte-informatique.txt",
              type: "txt",
              content: "Charte informatique — Version 1\n\n1. Mots de passe\nLe mot de passe doit comporter au moins 8 caractères et être changé tous les ans.\n\n2. Messagerie\nL'usage personnel de la messagerie professionnelle est toléré s'il reste raisonnable.\n\n3. Postes de travail\nChaque collaborateur verrouille sa session lorsqu'il quitte son poste.\n\n4. Sanctions\nTout manquement pourra faire l'objet d'une sanction disciplinaire.\n"
            }
          },
          tasks: [
            { text: "Collez le texte dans Word, enregistrez-le sous <code>charte-v1.docx</code>, puis enregistrez une copie <code>charte-v2.docx</code>.", hint: "Fichier › Enregistrer une copie." },
            { text: "Dans la v2, activez le <b>suivi des modifications</b> (<kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>E</kbd>), passez à 12 caractères, supprimez « et être changé tous les ans », ajoutez un article « 5. Télétravail ». Observez l'affichage <i>Toutes les marques</i> puis <i>Marques simples</i>.", hint: "Révision › Suivi › Afficher pour révision." },
            { text: "Désactivez le suivi, faites une modification « en douce » dans la v2, enregistrez. Puis <i>Révision › Comparer › Comparer…</i> v1 (original) et v2 (révisé) : la modification cachée apparaît.", hint: "Le document de comparaison est un nouveau fichier." },
            { text: "Dans le document de comparaison, <b>acceptez</b> tout sauf la suppression sur le changement annuel, que vous <b>refusez</b>.", hint: "Révision › Modifications › Suivante / Accepter / Refuser." },
            { text: "Protégez le résultat : <i>Révision › Protéger › Restreindre la modification</i> › Commentaires uniquement › Oui, activer la protection. Puis lancez <i>Fichier › Informations › Vérifier l'absence de problèmes › Inspecter le document</i>.", hint: "Testez : vous ne pouvez plus taper dans le texte, seulement commenter." }
          ],
          solution: `<ol>
<li>Coller, <i>Fichier › Enregistrer sous</i> charte-v1.docx, puis <i>Enregistrer une copie</i> charte-v2.docx et ouvrir la v2.</li>
<li><kbd>Ctrl</kbd>+<kbd>Maj</kbd>+<kbd>E</kbd> : les ajouts apparaissent soulignés, les suppressions barrées.</li>
<li><i>Révision › Comparer › Comparer…</i> › Document original : charte-v1, Document révisé : charte-v2 › OK.</li>
<li>Parcourir avec <i>Suivante</i>, <i>Accepter et passer à la suivante</i> / <i>Refuser et passer à la suivante</i>.</li>
<li><i>Restreindre la modification</i> › cocher « Autoriser uniquement ce type de modifications » › <i>Commentaires</i> › <i>Oui, activer la protection</i>. Inspecteur : supprimer propriétés et informations personnelles.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Vous passez l'affichage en « Aucune marque » puis envoyez le document. Que voit le destinataire ?",
          choices: ["Le texte final, sans aucune trace", "Potentiellement toutes les modifications suivies et les commentaires", "Un document vide", "Une version PDF"],
          answer: 1,
          explain: "« Aucune marque » est un simple affichage : les révisions restent dans le fichier tant qu'elles ne sont pas acceptées ou refusées."
        },
        {
          q: "Un client vous renvoie un contrat modifié sans suivi des modifications. Comment repérer les changements ?",
          choices: ["Relire mot à mot", "Révision › Comparer › Comparer…", "Révision › Combiner", "Fichier › Historique des versions"],
          answer: 1,
          explain: "Comparer produit un document où tous les écarts apparaissent comme des modifications suivies."
        },
        {
          q: "Quelle protection empêche réellement d'ouvrir un document sans autorisation ?",
          choices: ["Marquer comme final", "Restreindre la modification", "Chiffrer avec mot de passe", "Toujours ouvrir en lecture seule"],
          answer: 2,
          explain: "Seul le chiffrement empêche l'ouverture ; les autres options limitent les modifications ou avertissent."
        },
        {
          q: "Comment empêcher un relecteur de désactiver le suivi des modifications ?",
          choices: ["Lui demander gentiment", "Verrouiller le suivi (ou Restreindre la modification › Modifications suivies)", "Enregistrer en PDF", "Désactiver la correction automatique"],
          answer: 1,
          explain: "« Verrouiller le suivi » ou la restriction « Modifications suivies » imposent le suivi pour toute modification."
        },
        {
          q: "Que faire avant d'envoyer un document à l'extérieur pour retirer commentaires et métadonnées ?",
          choices: ["Fichier › Informations › Vérifier l'absence de problèmes › Inspecter le document", "Ctrl+A puis Suppr", "Changer le nom du fichier", "Masquer la règle"],
          answer: 0,
          explain: "L'Inspecteur de document recherche et supprime commentaires, révisions, propriétés et texte masqué."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 4 — w07 : Publipostage
       =================================================================== */
    {
      id: "w07",
      level: 4,
      title: "Publipostage : lettres, e-mails et étiquettes personnalisés",
      icon: "✉️",
      duration: 70,
      summary: "Générez en quelques minutes des centaines de courriers, e-mails ou étiquettes personnalisés à partir d'un fichier de données.",
      objectives: [
        "Être capable de préparer une source de données propre (Excel, CSV, contacts Outlook)",
        "Être capable d'insérer des champs de fusion, un bloc d'adresse et une ligne de salutation",
        "Être capable d'utiliser les règles Si…Alors…Sinon et de formater nombres et dates fusionnés",
        "Être capable de filtrer les destinataires et de fusionner vers un document, une imprimante ou des e-mails",
        "Être capable de produire des étiquettes d'adresse"
      ],
      lessons: [
        {
          title: "Le principe et la source de données",
          html: `<p>Un publipostage combine un <b>document principal</b> (le modèle de lettre, identique pour tous) et une <b>source de données</b> (un tableau : une ligne par destinataire, une colonne par information).</p>
<div class="o-path">Publipostage › Démarrer la fusion et le publipostage › Lettres</div>
<div class="o-path">Publipostage › Sélection des destinataires › Utiliser une liste existante…</div>
<p>Sources acceptées : classeur Excel (.xlsx), fichier texte/CSV, liste de contacts Outlook, base Access, ou tableau Word.</p>
<div class="callout ok">Une source propre : <b>une seule ligne d'en-têtes</b> en ligne 1, des noms de colonnes courts sans caractères spéciaux (<code>Prenom</code>, <code>CP</code>, <code>Montant</code>), aucune ligne vide ni cellule fusionnée, une colonne par information (Prénom et Nom séparés !).</div>
<div class="callout warn">CSV : Word utilise le séparateur de liste des paramètres régionaux (le « ; » en France). Les <b>codes postaux</b> commençant par 0 (01000…) peuvent perdre leur zéro si le fichier a été enregistré depuis Excel avec une colonne au format Nombre : passez la colonne en Texte.</div>
<div class="o-tip">💡 Débutant ? <i>Démarrer la fusion › Assistant Fusion et publipostage pas à pas</i> vous guide en 6 étapes dans un volet.</div>`
        },
        {
          title: "Insérer les champs de fusion",
          html: `<div class="o-path">Publipostage › Champs d'écriture et d'insertion › Insérer un champ de fusion</div>
<p>Chaque champ apparaît entre chevrons : <code>«Prenom»</code>. Laissez les espaces et la ponctuation <b>hors</b> des champs : <code>«Prenom» «Nom»</code>.</p>
<ul>
<li><b>Bloc d'adresse</b> : assemble automatiquement nom, société, adresse (bouton <i>Faire correspondre les champs</i> si vos colonnes ont des noms inhabituels).</li>
<li><b>Ligne de salutation</b> : « Cher Monsieur Dupont, » avec une formule de repli pour les données incomplètes.</li>
<li><b>Aperçu des résultats</b> : remplace les chevrons par les vraies données ; flèches pour faire défiler les destinataires.</li>
<li><b>Rechercher les erreurs</b> : simule la fusion et signale les champs manquants.</li>
</ul>
<div class="callout warn">Les montants et dates d'Excel arrivent souvent mal formatés (<code>1250,5</code> au lieu de <code>1 250,50 €</code>, ou une date en format américain). Corrigez avec un <b>commutateur de champ</b> : <kbd>Alt</kbd>+<kbd>F9</kbd> affiche les codes, puis :</div>
<pre class="code">{ MERGEFIELD Montant \\# "# ##0,00 €" }
{ MERGEFIELD Echeance \\@ "dd/MM/yyyy" }</pre>
<p><kbd>Alt</kbd>+<kbd>F9</kbd> à nouveau pour revenir à l'affichage normal, puis <kbd>F9</kbd> sur le champ pour le mettre à jour.</p>`
        },
        {
          title: "Règles, filtres et tri",
          html: `<div class="o-path">Publipostage › Champs d'écriture et d'insertion › Règles › Si…Alors…Sinon…</div>
<p>Exemple : si <code>Civilite</code> est égal à <code>Mme</code>, insérer « Chère », sinon « Cher ». Le champ généré :</p>
<pre class="code">{ IF { MERGEFIELD Civilite } = "Mme" "Chère" "Cher" } «Civilite» «Nom»,</pre>
<p>Autres règles : <i>Ignorer l'enregistrement si…</i> (exclure), <i>Demander</i> (saisie à la fusion), <i>Numéro d'enregistrement</i>.</p>
<div class="o-path">Publipostage › Sélection des destinataires › Modifier la liste de destinataires</div>
<p>Décochez des lignes, ou utilisez <b>Filtrer…</b> (ex. <code>Ville</code> <i>Égal à</i> <code>Lyon</code>, ou <code>Montant</code> <i>Supérieur à</i> <code>1000</code>) et <b>Trier…</b> (par code postal pour un affranchissement en nombre).</p>
<div class="callout">Les filtres dans la liste de destinataires portent sur du texte : pour des montants, assurez-vous que la colonne source est bien numérique, sinon « 900 » sera considéré supérieur à « 1000 ».</div>
<div class="o-tip">💡 Le bouton <i>Ajouter / supprimer des enregistrements en double</i> détecte les destinataires présents deux fois.</div>`
        },
        {
          title: "Terminer la fusion : documents, e-mails, étiquettes",
          html: `<div class="o-path">Publipostage › Terminer › Terminer &amp; fusionner</div>
<table class="o-table"><thead><tr><th>Option</th><th>Résultat</th></tr></thead>
<tbody>
<tr><td>Modifier des documents individuels…</td><td>Un nouveau document « Lettres1 » avec toutes les lettres (séparées par des sauts de section) — vérifiable avant impression</td></tr>
<tr><td>Imprimer des documents…</td><td>Envoi direct à l'imprimante</td></tr>
<tr><td>Envoyer des messages électroniques…</td><td>Un e-mail par destinataire via Outlook (champ À : colonne Email, ligne Objet, format HTML)</td></tr>
</tbody></table>
<div class="callout warn">L'envoi d'e-mails par publipostage nécessite <b>Outlook classique</b> pour bureau configuré comme messagerie par défaut ; il n'est pas possible de joindre une pièce jointe différente par destinataire.</div>
<div class="o-path">Publipostage › Démarrer la fusion et le publipostage › Étiquettes…</div>
<p>Choisissez le fournisseur et la référence (ex. Avery L7160 : 21 étiquettes par planche), sélectionnez la liste, insérez le bloc d'adresse dans la 1ʳᵉ étiquette, puis <b>Mettre à jour les étiquettes</b> pour recopier la mise en page sur toute la planche.</p>
<div class="callout ok">Conservez le document principal : à la prochaine campagne, il suffit de l'ouvrir, Word demande de se reconnecter à la source (mise à jour), et la fusion est prête.</div>`
        }
      ],
      shortcuts: [
        { keys: "Alt+F9", desc: "Afficher/masquer les codes de champ" },
        { keys: "F9", desc: "Mettre à jour le champ sélectionné" },
        { keys: "Ctrl+F9", desc: "Insérer des accolades de champ vides" },
        { keys: "Alt+Maj+N", desc: "Fusionner vers un nouveau document" },
        { keys: "Alt+Maj+K", desc: "Aperçu de la fusion" },
        { keys: "Alt+Maj+E", desc: "Modifier la source de données" }
      ],
      labs: [
        {
          id: "w07-lab1",
          title: "Lab 1 — Campagne de relance des cotisations",
          kind: "steps",
          scenario: `<p>Vous êtes chargé(e) de la vie associative de la <b>Fédération des Commerçants du Vieux-Lyon</b>. Huit adhérents n'ont pas réglé leur cotisation 2027. Vous devez produire une lettre personnalisée pour chacun (avec « Cher/Chère » selon la civilité et le montant bien formaté), puis les étiquettes pour les enveloppes. Téléchargez le fichier des adhérents.</p>`,
          setup: {
            app: "Word",
            file: {
              name: "adherents-relance.csv",
              type: "csv",
              content: "Civilite;Prenom;Nom;Societe;Adresse;CP;Ville;Montant;Echeance;Email\nMme;Camille;Roux;Librairie du Change;3 place du Change;69005;Lyon;250,00;15/01/2027;camille.roux@exemple.fr\nM.;Karim;Benali;Épices & Co;12 rue Saint-Jean;69005;Lyon;180,00;15/01/2027;karim.benali@exemple.fr\nMme;Inès;Moreau;Atelier Soie;8 rue du Bœuf;69005;Lyon;250,00;15/01/2027;ines.moreau@exemple.fr\nM.;Thomas;Garnier;Café des Juges;1 place Neuve Saint-Jean;69005;Lyon;320,00;31/01/2027;thomas.garnier@exemple.fr\nMme;Lucie;Faure;Bijoux Lucie;22 rue Juiverie;69005;Lyon;180,00;31/01/2027;lucie.faure@exemple.fr\nM.;Hugo;Perrin;Fromagerie Perrin;5 montée du Gourguillon;69005;Lyon;250,00;15/02/2027;hugo.perrin@exemple.fr\nMme;Sarah;Lambert;Galerie Lambert;14 quai Romain Rolland;69005;Lyon;1250,00;15/02/2027;sarah.lambert@exemple.fr\nM.;Antoine;Chevalier;Bouchon Antoine;9 rue des Trois-Maries;69005;Lyon;320,00;15/02/2027;antoine.chevalier@exemple.fr\n"
            }
          },
          tasks: [
            { text: "Nouveau document Word › <i>Publipostage › Démarrer la fusion et le publipostage › Lettres</i> › <i>Sélection des destinataires › Utiliser une liste existante</i> › choisissez <code>adherents-relance.csv</code>. Vérifiez dans <i>Modifier la liste de destinataires</i> que les 8 adhérents et les colonnes sont bien reconnus.", hint: "Si tout tient dans une seule colonne, le séparateur « ; » n'a pas été reconnu : vérifiez les paramètres régionaux, ou ouvrez le CSV dans Excel et enregistrez-le en .xlsx." },
            { text: "Rédigez la lettre : bloc d'adresse (ou champs «Prenom» «Nom», «Societe», «Adresse», «CP» «Ville»), puis une salutation avec la règle <i>Si…Alors…Sinon</i> : Civilite = Mme → « Chère », sinon « Cher », suivie de «Civilite» «Nom»,", hint: "Publipostage › Règles › Si…Alors…Sinon…" },
            { text: "Insérez dans le corps : « Votre cotisation de «Montant» est à régler avant le «Echeance». » Puis <kbd>Alt</kbd>+<kbd>F9</kbd> et ajoutez le commutateur <code>\\# \"# ##0,00 €\"</code> au champ Montant. Vérifiez avec <i>Aperçu des résultats</i> que Sarah Lambert affiche « 1 250,00 € ».", hint: "Alt+F9 à nouveau pour masquer les codes, puis F9 pour mettre à jour." },
            { text: "Filtrez pour ne garder que les échéances au <b>15/01/2027</b> (<i>Modifier la liste › Filtrer…</i>), puis <i>Terminer &amp; fusionner › Modifier des documents individuels › Tous</i>. Vous devez obtenir 3 lettres.", hint: "Roux, Benali, Moreau." },
            { text: "Retirez le filtre, enregistrez le document principal, puis créez les <b>étiquettes</b> : nouveau document › <i>Démarrer la fusion › Étiquettes</i> (Avery L7160) › même liste › Bloc d'adresse › <i>Mettre à jour les étiquettes</i> › fusion vers un nouveau document.", hint: "Sans « Mettre à jour les étiquettes », seule la 1ʳᵉ étiquette de chaque planche est remplie." }
          ],
          solution: `<ol>
<li>Démarrer la fusion › Lettres ; Sélection des destinataires › Utiliser une liste existante › adherents-relance.csv (8 lignes, 10 colonnes).</li>
<li>Règles › Si…Alors…Sinon : Nom du champ <code>Civilite</code>, Comparaison <i>Égal à</i>, Comparer à <code>Mme</code>, Insérer ce texte « Chère », Sinon « Cher ». Puis espace, champ Civilite, espace, champ Nom, virgule.</li>
<li>Code de champ attendu : <code>{ MERGEFIELD Montant \\# "# ##0,00 €" }</code>. Aperçu : « 1 250,00 € » pour Sarah Lambert, « 250,00 € » pour Camille Roux.</li>
<li>Filtrer : Champ <code>Echeance</code> <i>Égal à</i> <code>15/01/2027</code> › 3 destinataires › Terminer &amp; fusionner › Modifier des documents individuels.</li>
<li>Étiquettes : Démarrer la fusion › Étiquettes › Avery A4/A5 › L7160 ; Utiliser une liste existante ; Bloc d'adresse ; Mettre à jour les étiquettes ; Terminer &amp; fusionner.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Quels sont les deux éléments d'un publipostage ?",
          choices: ["Un modèle .dotx et une macro", "Un document principal et une source de données", "Un PDF et un fichier Excel", "Une table des matières et un index"],
          answer: 1,
          explain: "Le document principal contient le texte commun et les champs ; la source de données fournit une ligne par destinataire."
        },
        {
          q: "Le montant fusionné s'affiche « 1250,5 ». Comment obtenir « 1 250,50 € » ?",
          choices: ["Le modifier à la main dans chaque lettre", "Ajouter le commutateur \\# \"# ##0,00 €\" au champ MERGEFIELD", "Mettre le champ en gras", "Changer la police"],
          answer: 1,
          explain: "Le commutateur numérique \\# formate la valeur fusionnée ; on l'ajoute après Alt+F9."
        },
        {
          q: "Vous voulez « Chère Madame » ou « Cher Monsieur » selon le destinataire. Quel outil ?",
          choices: ["Règles › Si…Alors…Sinon…", "Rechercher/Remplacer après fusion", "Deux documents principaux", "Ligne de salutation uniquement"],
          answer: 0,
          explain: "La règle Si…Alors…Sinon insère un texte différent selon la valeur d'un champ (ex. Civilite)."
        },
        {
          q: "Pour les étiquettes, seule la première étiquette de chaque planche est remplie. Qu'avez-vous oublié ?",
          choices: ["Aperçu des résultats", "Mettre à jour les étiquettes", "Rechercher les erreurs", "Le bloc d'adresse"],
          answer: 1,
          explain: "« Mettre à jour les étiquettes » recopie la mise en page et le champ «Enregistrement suivant» dans toutes les étiquettes."
        },
        {
          q: "Quelle option permet de relire toutes les lettres avant de les imprimer ?",
          choices: ["Imprimer des documents", "Envoyer des messages électroniques", "Modifier des documents individuels", "Aperçu avant impression"],
          answer: 2,
          explain: "« Modifier des documents individuels » crée un nouveau document contenant toutes les lettres, modifiables et vérifiables."
        }
      ]
    },

    /* ===================================================================
       NIVEAU 4 — w08 : Modèles, champs, contrôles, macros, accessibilité
       =================================================================== */
    {
      id: "w08",
      level: 4,
      title: "Modèles, champs, contrôles de contenu, macros VBA et accessibilité",
      icon: "⚙️",
      duration: 80,
      summary: "Industrialisez vos documents : modèles d'entreprise, formulaires à remplir, automatisation par macros et documents accessibles à tous.",
      objectives: [
        "Être capable de créer un modèle .dotx et de l'utiliser comme base de nouveaux documents",
        "Être capable d'insérer des champs (date, nom de fichier, propriétés) et de les mettre à jour",
        "Être capable de concevoir un formulaire avec des contrôles de contenu et de le protéger",
        "Être capable d'enregistrer, écrire et exécuter une macro VBA simple",
        "Être capable de vérifier et corriger l'accessibilité d'un document"
      ],
      lessons: [
        {
          title: "Modèles (.dotx) et Normal.dotm",
          html: `<p>Un <b>modèle</b> contient tout ce qui doit être identique d'un document à l'autre : logo dans l'en-tête, styles, marges, texte standard, contrôles de contenu. On ne le modifie pas pour chaque usage : on crée un document <i>à partir</i> de lui.</p>
<div class="o-path">Fichier › Enregistrer sous › Type : Modèle Word (*.dotx)</div>
<p>Par défaut, Word propose le dossier <code>Documents\\Modèles Office personnalisés</code>. Vos modèles apparaissent alors dans <b>Fichier › Nouveau › Personnel</b>.</p>
<table class="o-table"><thead><tr><th>Action sur un .dotx</th><th>Résultat</th></tr></thead>
<tbody>
<tr><td>Double-clic dans l'Explorateur</td><td>Crée un nouveau « Document1 » basé sur le modèle</td></tr>
<tr><td>Clic droit › Ouvrir</td><td>Ouvre le modèle lui-même pour le modifier</td></tr>
</tbody></table>
<div class="callout warn"><b>Normal.dotm</b> est le modèle de tous les nouveaux documents vierges. Le cocher « Définir par défaut » dans une boîte Police ou Paragraphe le modifie : faites-le en connaissance de cause. S'il est corrompu, Word le recrée si vous le renommez (dossier <code>%appdata%\\Microsoft\\Templates</code>).</div>
<div class="o-tip">💡 En entreprise, déposez les modèles officiels dans un dossier partagé et déclarez-le dans <i>Fichier › Options › Enregistrement › Emplacement des modèles personnels par défaut</i> (ou via les modèles organisationnels SharePoint).</div>`
        },
        {
          title: "Champs et contrôles de contenu",
          html: `<div class="o-path">Insertion › Texte › QuickPart › Champ…</div>
<table class="o-table"><thead><tr><th>Champ</th><th>Affiche</th></tr></thead>
<tbody>
<tr><td><code>{ DATE \\@ "d MMMM yyyy" }</code></td><td>La date du jour (change à chaque ouverture !)</td></tr>
<tr><td><code>{ CREATEDATE \\@ "dd/MM/yyyy" }</code></td><td>La date de création (stable — préférable pour un courrier)</td></tr>
<tr><td><code>{ PAGE }</code> / <code>{ NUMPAGES }</code></td><td>Numéro de page / nombre total</td></tr>
<tr><td><code>{ FILENAME \\p }</code></td><td>Chemin complet du fichier</td></tr>
<tr><td><code>{ STYLEREF "Titre 1" }</code></td><td>Le titre du chapitre en cours</td></tr>
</tbody></table>
<p><kbd>Alt</kbd>+<kbd>F9</kbd> bascule l'affichage des codes, <kbd>F9</kbd> met à jour, <kbd>Ctrl</kbd>+<kbd>F9</kbd> insère des accolades de champ (on ne tape jamais les accolades au clavier).</p>
<div class="o-path">Fichier › Options › Personnaliser le ruban › cocher Développeur</div>
<div class="o-path">Développeur › Contrôles › Contrôle du contenu de texte brut</div>
<p>Les <b>contrôles de contenu</b> créent des zones à remplir : texte enrichi, texte brut, image, liste déroulante, zone de liste déroulante (avec saisie libre), sélecteur de date, case à cocher. Le bouton <b>Mode Création</b> modifie le texte d'invite (« Cliquez ici pour saisir le nom du client »), <b>Propriétés</b> définit le titre, la balise et le verrouillage (<i>Le contrôle de contenu ne peut pas être supprimé</i>).</p>
<div class="callout ok">Pour un formulaire : <i>Restreindre la modification › Remplissage de formulaires</i> — l'utilisateur ne peut écrire que dans les contrôles.</div>`
        },
        {
          title: "Macros VBA : enregistrer et écrire",
          html: `<div class="o-path">Affichage › Macros › Enregistrer une macro…</div>
<p>Nommez la macro (sans espace), choisissez où la stocker (<i>Tous les documents (Normal.dotm)</i> ou le document courant), affectez-lui éventuellement un bouton ou un raccourci, puis réalisez les actions. <i>Arrêter l'enregistrement</i> quand c'est fini.</p>
<p>Pour lire ou modifier le code : <kbd>Alt</kbd>+<kbd>F11</kbd> (éditeur Visual Basic). Pour exécuter : <kbd>Alt</kbd>+<kbd>F8</kbd>.</p>
<pre class="code">Sub NettoyerEspaces()
    ' Remplace les espaces multiples par une seule espace
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "( ){2,}"
        .Replacement.Text = " "
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
    MsgBox "Espaces nettoyées !", vbInformation
End Sub</pre>
<div class="callout warn">Un document avec macros doit être enregistré en <b>.docm</b> (ou modèle .dotm). Les fichiers téléchargés depuis Internet ou reçus par e-mail voient leurs macros <b>bloquées</b> par défaut (bandeau « Risque de sécurité ») : c'est une protection contre les virus de macro. N'activez que les macros dont vous connaissez l'origine.</div>
<div class="o-tip">💡 L'enregistreur ne capture pas les clics souris dans le texte : déplacez-vous au clavier (flèches, <kbd>Ctrl</kbd>+<kbd>Fin</kbd>) pendant l'enregistrement.</div>`
        },
        {
          title: "Accessibilité : des documents pour tous",
          html: `<div class="o-path">Révision › Accessibilité › Vérifier l'accessibilité</div>
<p>Le volet liste les <b>Erreurs</b>, <b>Avertissements</b> et <b>Conseils</b>, avec la correction proposée en un clic. Les règles essentielles :</p>
<ul>
<li><b>Structure par styles</b> : Titre 1, puis Titre 2… sans sauter de niveau (pas de Titre 3 directement sous un Titre 1).</li>
<li><b>Texte de remplacement</b> sur chaque image (clic droit › <i>Afficher le texte de remplacement</i>) ; cochez <i>Élément décoratif</i> pour les images purement esthétiques.</li>
<li><b>Tableaux</b> simples, avec <i>Ligne d'en-tête</i> cochée et <i>Répéter les lignes d'en-tête</i> ; pas de cellules fusionnées ni de tableau de mise en page.</li>
<li><b>Liens explicites</b> : « Télécharger le règlement intérieur (PDF) » plutôt que « cliquez ici ».</li>
<li><b>Contraste</b> suffisant et information jamais portée par la seule couleur.</li>
<li>Pas de soulignement hors liens, pas de texte en image, langue du document correcte (<i>Révision › Langue</i>).</li>
</ul>
<div class="callout ok">À l'export PDF, cochez dans <i>Options…</i> <b>Balises de structure de document pour l'accessibilité</b> et <b>Créer des signets à l'aide des titres</b>.</div>
<div class="o-tip">💡 Activez <i>Fichier › Options › Accessibilité › Laisser le vérificateur d'accessibilité s'exécuter pendant que je travaille</i> : la barre d'état vous alerte en continu.</div>`
        }
      ],
      shortcuts: [
        { keys: "Alt+F8", desc: "Liste des macros (exécuter)" },
        { keys: "Alt+F11", desc: "Éditeur Visual Basic" },
        { keys: "Alt+F9", desc: "Afficher/masquer les codes de champ" },
        { keys: "Ctrl+F9", desc: "Insérer un champ vide { }" },
        { keys: "Ctrl+Maj+F9", desc: "Figer un champ en texte" },
        { keys: "F9", desc: "Mettre à jour les champs" }
      ],
      labs: [
        {
          id: "w08-lab1",
          title: "Lab 1 — Rendre accessible le livret d'accueil",
          kind: "doc",
          scenario: `<p>La mairie de <b>Saint-Clair-sur-Loire</b> publie son livret d'accueil des nouveaux habitants en ligne. Le vérificateur d'accessibilité signale plusieurs erreurs : hiérarchie des titres incohérente, lien « cliquez ici », tableau sans ligne d'en-tête identifiée, soulignement trompeur. Corrigez-les.</p>`,
          setup: {
            html: `<h1>Bienvenue à Saint-Clair-sur-Loire</h1>
<p>Ce livret rassemble les informations utiles pour votre installation dans la commune.</p>
<h3>Démarches administratives</h3>
<p>Pensez à vous inscrire sur les listes électorales et à signaler votre changement d'adresse. Pour la liste complète des démarches, <a href="https://www.service-public.fr">cliquez ici</a>.</p>
<h3>Horaires de la mairie</h3>
<table><tbody><tr><td>Jour</td><td>Matin</td><td>Après-midi</td></tr><tr><td>Lundi au vendredi</td><td>8 h 30 – 12 h</td><td>13 h 30 – 17 h</td></tr><tr><td>Samedi</td><td>9 h – 12 h</td><td>Fermé</td></tr></tbody></table>
<h2>Collecte des déchets</h2>
<p>Les bacs jaunes sont collectés le mardi. <u>La date de dépôt des encombrants</u> est fixée au premier lundi de chaque mois.</p>`
          },
          tasks: [
            {
              text: "Corrigez la <b>hiérarchie des titres</b> : aucun niveau ne doit être sauté (un Titre 3 ne peut pas suivre directement un Titre 1).",
              hint: "Passez « Démarches administratives » et « Horaires de la mairie » en Titre 2.",
              check: (a) => {
                const hs = a.headings();
                return hs.length >= 4 && hs[0].level === 1 && hs.every((h, i) => i === 0 || h.level <= hs[i - 1].level + 1);
              }
            },
            {
              text: "Remplacez le texte du lien « cliquez ici » par un <b>texte explicite</b> (ex. « Consulter les démarches sur Service-Public.fr »).",
              hint: "Clic dans le lien › Ctrl+K (Insertion › Lien) › modifiez « Texte à afficher ». Dans le vrai Word : clic droit › Modifier le lien.",
              check: (a) => {
                const ls = a.$$("a[href]");
                return ls.length >= 1 && !a.has("cliquez ici") && ls.every((l) => norm(l.textContent).length >= 10);
              }
            },
            {
              text: "Mettez la <b>ligne d'en-tête</b> du tableau des horaires en gras (Jour / Matin / Après-midi).",
              hint: "Sélectionnez la 1ʳᵉ ligne › Ctrl+G. Dans le vrai Word : Création de tableau › Ligne d'en-tête.",
              check: (a) => a.tables().length >= 1 && a.tables().every((t) => t.headerBold)
            },
            {
              text: "Retirez le <b>soulignement</b> de « La date de dépôt des encombrants » (un soulignement fait croire à un lien) et mettez-le plutôt en gras.",
              hint: "Sélection › Ctrl+U pour retirer, puis Ctrl+G.",
              check: (a) => { const f = a.fmt("date de dépôt"); return !!(f.found && !f.underline && f.bold); }
            }
          ],
          solution: `<ol>
<li>Clic dans « Démarches administratives » › <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>2</kbd> ; idem « Horaires de la mairie ». Ordre obtenu : 1, 2, 2, 2.</li>
<li>Clic dans le lien › <kbd>Ctrl</kbd>+<kbd>K</kbd> (ou <i>Insertion › Lien</i> ; dans le vrai Word : clic droit › <i>Modifier le lien</i>) › <i>Texte à afficher</i> : « Consulter les démarches sur Service-Public.fr » › OK. (Sélectionner tout le lien puis retaper par-dessus supprime le lien.)</li>
<li>Sélectionnez la ligne Jour / Matin / Après-midi › <kbd>Ctrl</kbd>+<kbd>G</kbd>. Dans Word : <i>Création de tableau › Ligne d'en-tête</i> et <i>Disposition › Répéter les lignes d'en-tête</i>.</li>
<li>Sélectionnez « La date de dépôt des encombrants » › <kbd>Ctrl</kbd>+<kbd>U</kbd> (retire) › <kbd>Ctrl</kbd>+<kbd>G</kbd>.</li>
</ol>`
        },
        {
          id: "w08-lab2",
          title: "Lab 2 — Modèle de devis avec formulaire et macro",
          kind: "steps",
          scenario: `<p>L'entreprise de rénovation <b>BâtiPro 38</b> veut un modèle de devis standardisé : les commerciaux ne remplissent que des zones prévues, et une macro nettoie le texte avant envoi. À réaliser dans le vrai Word (onglet Développeur requis).</p>`,
          setup: { app: "Word" },
          tasks: [
            { text: "Affichez l'onglet <b>Développeur</b> (<i>Fichier › Options › Personnaliser le ruban</i>). Dans un nouveau document, tapez « Devis n° », « Client : », « Date : », « Type de travaux : », « Acompte versé : ».", hint: "Cochez Développeur dans la liste de droite." },
            { text: "Insérez les contrôles : <b>texte brut</b> après « Devis n° » et « Client : », <b>sélecteur de date</b> après « Date : » (format jj/MM/aaaa), <b>liste déroulante</b> après « Type de travaux : » (Peinture, Plomberie, Électricité, Menuiserie), <b>case à cocher</b> après « Acompte versé : ». Personnalisez les textes d'invite en <i>Mode Création</i>.", hint: "Développeur › Contrôles ; Propriétés › Ajouter… pour les choix de la liste." },
            { text: "Ajoutez dans le pied de page le champ <code>{ FILENAME }</code> et « Page { PAGE } sur { NUMPAGES } » (Insertion › QuickPart › Champ). Protégez : <i>Restreindre la modification › Remplissage de formulaires</i> › Oui, activer la protection.", hint: "Testez : seuls les contrôles sont modifiables." },
            { text: "Ôtez la protection, puis <kbd>Alt</kbd>+<kbd>F11</kbd> › <i>Insertion › Module</i> dans le projet du document : collez la macro <code>NettoyerEspaces</code> de la leçon. Tapez du texte avec plusieurs espaces et lancez-la avec <kbd>Alt</kbd>+<kbd>F8</kbd>.", hint: "Le module doit être dans « Project (Document1) » pour voyager avec le fichier." },
            { text: "Réactivez la protection et enregistrez en <b>Modèle Word prenant en charge les macros (*.dotm)</b> sous le nom <code>Devis-BatiPro.dotm</code>. Créez un devis via <i>Fichier › Nouveau › Personnel</i>.", hint: "En .dotx la macro serait perdue." },
            { text: "Lancez <i>Révision › Vérifier l'accessibilité</i> sur le devis et corrigez les éventuels avertissements (titres, texte de remplacement du logo…).", hint: "Une image de logo sans texte de remplacement est une erreur." }
          ],
          solution: `<ol>
<li><i>Fichier › Options › Personnaliser le ruban</i> › cocher <i>Développeur</i> › OK.</li>
<li><i>Développeur › Contrôles</i> : Aa (texte brut), sélecteur de date (Propriétés › Afficher la date ainsi : dd/MM/yyyy), liste déroulante (Propriétés › Ajouter… pour chaque type), case à cocher. Mode Création pour modifier les invites.</li>
<li>Pied de page › <i>Insertion › QuickPart › Champ…</i> › FileName ; Page ; NumPages. <i>Révision › Restreindre la modification</i> › Remplissage de formulaires › Oui, activer la protection.</li>
<li><kbd>Alt</kbd>+<kbd>F11</kbd> › clic droit sur Project (Document1) › Insertion › Module › coller la macro › fermer l'éditeur › <kbd>Alt</kbd>+<kbd>F8</kbd> › NettoyerEspaces › Exécuter.</li>
<li><i>Fichier › Enregistrer sous</i> › Type : Modèle Word prenant en charge les macros (*.dotm) › dossier Modèles Office personnalisés.</li>
<li><i>Révision › Vérifier l'accessibilité</i> › appliquer les corrections suggérées.</li>
</ol>`
        }
      ],
      quiz: [
        {
          q: "Vous double-cliquez sur un fichier .dotx dans l'Explorateur. Que se passe-t-il ?",
          choices: ["Le modèle s'ouvre pour modification", "Un nouveau document basé sur le modèle est créé", "Word affiche une erreur", "Le modèle est converti en PDF"],
          answer: 1,
          explain: "Double-clic = nouveau document basé sur le modèle ; pour modifier le modèle lui-même : clic droit › Ouvrir."
        },
        {
          q: "Pour une date de courrier qui ne doit pas changer à chaque ouverture, quel champ choisir ?",
          choices: ["DATE", "CREATEDATE", "NUMPAGES", "FILENAME"],
          answer: 1,
          explain: "DATE affiche la date du jour à chaque mise à jour ; CREATEDATE conserve la date de création du document."
        },
        {
          q: "Vous enregistrez un document contenant une macro au format .docx. Que se passe-t-il ?",
          choices: ["La macro est conservée", "Word avertit que le code VBA sera supprimé", "La macro est convertie en champ", "Le fichier devient un modèle"],
          answer: 1,
          explain: "Le .docx ne peut pas contenir de macros : il faut enregistrer en .docm (ou .dotm)."
        },
        {
          q: "Quelle protection permet de n'autoriser la saisie que dans les contrôles de contenu ?",
          choices: ["Marquer comme final", "Chiffrer avec mot de passe", "Restreindre la modification › Remplissage de formulaires", "Suivi des modifications"],
          answer: 2,
          explain: "Le mode « Remplissage de formulaires » verrouille tout le document sauf les champs et contrôles."
        },
        {
          q: "Lequel de ces éléments est une erreur d'accessibilité ?",
          choices: ["Un lien intitulé « Télécharger le règlement (PDF) »", "Un Titre 3 placé directement sous un Titre 1", "Une image décorative marquée « Élément décoratif »", "Un tableau avec ligne d'en-tête répétée"],
          answer: 1,
          explain: "Sauter un niveau de titre casse la structure pour les lecteurs d'écran et la navigation."
        }
      ]
    }
  ]);
})();
