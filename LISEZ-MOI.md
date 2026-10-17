# SnipCompta — complément Excel pour vos pièces justificatives

SnipCompta ajoute un onglet **SNIPCOMPTA** au ruban d'Excel et un panneau à droite de la feuille. Vous y importez vos relevés bancaires, factures et scans (PDF, PNG, JPG), puis vous entourez à la souris ce que vous voulez récupérer. La valeur arrive dans la cellule Excel sélectionnée, avec un lien vers la preuve : cliquer sur la cellule réaffiche la pièce à l'endroit exact.

Vos documents ne quittent jamais votre ordinateur : la lecture, l'OCR et l'extraction se font dans Excel, et les pièces sont enregistrées dans le classeur lui-même.

---

## Ce que fait le complément

**Snips** (liens cellule ↔ preuve)
- **Text Snip** : entourez un texte ou un montant → il s'inscrit dans la cellule active (montants convertis en nombres).
- **Validation / Exception Snip** : ✓ vert ou ✗ rouge. Si la cellule contient déjà une valeur, elle est **conservée** (seuls la couleur et le lien sont ajoutés).
- **Sum Snip** : additionne plusieurs zones, même sur plusieurs pages.
- **Table Snip** : extrait un tableau sur une ou plusieurs pages (colonnes ajustables, lignes de suite fusionnées).
- **Ctrl + glisser** : ajoute une preuve supplémentaire au snip de la cellule active (plusieurs preuves pour une cellule).
- **Relier un snip à d'autres cellules** (onglet Snips) : une preuve pour plusieurs cellules.
- **Clic droit** sur une zone : aller à la cellule, convertir (Texte ↔ Validation ↔ Exception), supprimer.
- **Couleurs personnalisables** pour chaque type de snip.
- **Après un tri** dans Excel, les liens suivent leurs valeurs (et bouton « Recaler les liens »).

**Annotations sur les pièces**
- Bulles de commentaire (avec « résolu »), surlignage, texte, flèches, coches, croix.
- **Caviardage** : masque définitivement des informations sensibles (le texte disparaît réellement de la pièce).
- **Commentaire Excel lié** : la remarque devient un commentaire sur la cellule.

**Documents**
- Import de fichiers, **dossiers entiers**, **archives ZIP**, TXT/CSV/HTML (convertis en PDF), **PDF protégés par mot de passe**, collage Ctrl+V.
- **Organiseur** : dossiers de rangement, renommer, retirer.
- **Export des pièces annotées** (PDF avec snips, annotations et commentaires).
- **Retirer les pages sans snips / sans exceptions**, **compresser** les pièces.
- **Nouvelle version d'un document** (même nom, contenu différent) : le travail est transféré et recalé sur le texte.
- OCR des scans (français + anglais), fait sur votre PC. Avant la lecture, une copie de l'image est nettoyée (fond, ombres, contraste) et les pages scannées de travers sont redressées. La pièce garde son aspect ; seule une page penchée s'affiche redressée.
- **Lecture des scans plus précise** (Paramètres) : moteur plus fin, un peu plus lent, téléchargé la première fois (environ 16 Mo). Activable et désactivable à tout moment ; les scans ouverts sont alors relus.
- **Tableaux sans traits** : l'outil Tableau reconnaît les colonnes par leur alignement (textes alignés à gauche, montants alignés à droite), même quand un libellé déborde sur la colonne voisine ou qu'un montant y est collé, et aussi sur les scans. Les libellés sur deux lignes (même décalées) et les en-têtes sur deux lignes ne créent plus de fausse colonne. Dans l'aperçu : glissez un en-tête **sur** la colonne voisine pour les fusionner, ou **entre** deux colonnes pour la déplacer ; **×** retire une colonne ou une ligne (↺ pour la remettre). Si vous entourez un tel passage avec l'outil Texte, SnipCompta propose « Insérer en tableau » (une valeur par cellule).
- **Modèles** (menu ⋯ → Modèles) : gardez vos réglages pour les réappliquer en un clic sur un autre classeur ou un autre client — découpage d'un tableau (« Enregistrer ce découpage… » dans l'aperçu ; ensuite chaque tableau entouré reprend les mêmes colonnes), champs d'extraction (Automatisation) et réglages de rapprochement (repris par nom de colonne, même si l'ordre change). Les modèles restent sur ce PC et sont aussi rangés dans le classeur ; « Tout exporter » / « Importer un fichier… » pour les passer à un collègue ou un autre PC.
- **Pivoter une pièce** : boutons ↺ ↻ à côté du zoom (ou menu ⋯). La page est vraiment tournée et enregistrée ainsi ; les snips déjà posés suivent. Maj + clic sur ↻ (ou « Pivoter tout le document » dans le menu) tourne toutes les pages. Un scan couché est relu automatiquement une fois redressé.
- **Ajuster un snip** : cliquez un snip pour le sélectionner, puis tirez ses poignées (bords et coins) pour agrandir ou réduire la zone. La valeur de la cellule se met à jour. Seul le texte réellement dans la zone est pris, même au milieu d'un mot.

**Automatisation**
- **Document Matching** : cellule par cellule, ou **ligne par ligne** (une pièce par ligne d'échantillon) avec options par colonne : obligatoire, approximatif (fautes, abréviations), partiel, tolérance (montant ou %), modèles réutilisables.
- **Find all sums** : dans le document affiché (chaque total et ses composants → feuille « Sommes trouvées ») ou à partir de totaux sélectionnés dans Excel.
- **Document Extraction** : formulaire par modèle (champs recalés sur leur libellé, une ligne par document ou par page, modèles), **factures** (fournisseur, n°, dates, HT, TVA, TTC, SIRET, TVA intracom., IBAN, contrôle HT + TVA = TTC) et **reçus**.
- **Version Compare** : différences entre deux versions d'un document, exportables.
- **Recherche** dans toutes les pièces (dates dans tous les formats) avec bouton **Snip**, et recherche automatique de la cellule sélectionnée (option).
- **Piste d'audit** : feuille listant tous les snips avec liens cliquables.

Raccourcis : `Ctrl+=` Texte · `Ctrl+,` Validation · `Ctrl+.` Exception · `Ctrl+[` Somme · `Ctrl+M` Tableau · `Ctrl+Z` annuler · `1`…`9` outils · `←` `→` pages · `Entrée` insérer · `Échap` annuler.

**Non inclus** (demandent l'IA ou des services en ligne de DataSnipper) : agents Excel, extraction par IA, WebSnip, intégrations SharePoint/Xero, fenêtre de visionneuse détachable, documents Word/Excel en source (enregistrez-les en PDF).

---

## Installation (environ 20 minutes, une seule fois)

Un complément Excel est une petite page web : il faut donc **(1)** mettre les fichiers en ligne, **(2)** indiquer cette adresse dans `manifest.xml`, **(3)** déclarer le complément dans Excel. La méthode ci-dessous est gratuite.

Il faut **Excel Microsoft 365 sur Windows** (l'onglet du ruban utilise une technologie Microsoft, le « runtime partagé », absente des vieilles versions d'Excel vendues sans abonnement).

### Étape 1 — Mettre les fichiers en ligne avec GitHub Pages (gratuit)

1. Créez un compte sur **https://github.com** (bouton *Sign up*). Retenez votre nom d'utilisateur, par exemple `zakaria-l`.
2. Une fois connecté, cliquez en haut à droite sur **+** → **New repository**.
   - *Repository name* : `snipcompta`
   - Cochez **Public** (obligatoire pour GitHub Pages gratuit ; seul le code est public, jamais vos documents).
   - Cliquez **Create repository**.
3. Sur la page qui s'affiche, cliquez le lien **uploading an existing file**.
4. Décompressez le zip sur votre PC, ouvrez le dossier `snipcompta`, **sélectionnez tout son contenu** (fichiers *et* dossiers `assets` et `lib`) et glissez-le dans la page GitHub. Attendez la fin du chargement, puis cliquez **Commit changes**.
5. Allez dans **Settings** (onglet du dépôt) → **Pages** (menu de gauche).
   - *Source* : **Deploy from a branch**
   - *Branch* : **main** et **/(root)** → **Save**.
6. Patientez 1 à 2 minutes et rechargez la page : l'adresse s'affiche, du type
   `https://zakaria-l.github.io/snipcompta/`
7. Vérification : ouvrez `https://zakaria-l.github.io/snipcompta/taskpane.html` dans votre navigateur. Vous devez voir SnipCompta en « mode démo ». Vous pouvez déjà l'essayer avec un PDF.

### Étape 2 — Mettre votre adresse dans le manifeste

1. Sur votre PC, faites un clic droit sur `manifest.xml` → **Ouvrir avec** → **Bloc-notes**.
2. Appuyez sur **Ctrl + H** (Remplacer).
   - *Rechercher* : `https://VOTRE-ADRESSE`
   - *Remplacer par* : votre adresse **sans la barre oblique finale**, ex. `https://zakaria-l.github.io/snipcompta`
   - Cliquez **Remplacer tout**, puis enregistrez (**Ctrl + S**).

### Étape 3 — Déclarer le complément dans Excel (double-clic)

1. Vérifiez que `manifest.xml` (avec votre adresse) est bien en ligne sur GitHub.
2. Double-cliquez **`Installer-SnipCompta.bat`**. Si Windows affiche « Windows a protégé votre ordinateur », cliquez **Informations complémentaires** → **Exécuter quand même** (le fichier est un simple script lisible avec le Bloc-notes).
3. Le script vous demande de fermer Excel si besoin, télécharge le manifeste, le déclare dans Excel et vide l'ancien cache. Appuyez sur une touche pour fermer la fenêtre.
4. Ouvrez Excel : l'onglet **SNIPCOMPTA** est dans le ruban.

Pour retirer le complément : double-cliquez **`Desinstaller-SnipCompta.bat`** (vos classeurs ne sont pas modifiés).

---

## Utilisation au quotidien

1. Onglet **SNIPCOMPTA** → **Importer des documents** (ou glissez vos PDF dans le panneau). Toutes les pages défilent à la suite ; en grand, les vignettes des pages s'affichent à gauche.
2. Cliquez la cellule de destination dans Excel.
3. Choisissez un outil dans le ruban, puis sur la pièce :
   - **un clic** sur un montant ou un mot le capture (il est surligné au survol) ;
   - **un double-clic** capture toute la ligne ;
   - **glisser un cadre** capture une zone ; pendant le glisser, la valeur lue s'affiche à côté du curseur.
4. Une bulle confirme : « 1 484,86 → B3 ✓ ». Le bouton **Annuler** de la bulle (ou Ctrl+Z) retire le snip.

**Fenêtre séparée agrandie** : la barre d'outils reprend les icônes du ruban, les vignettes des pages sont à gauche et la liste des snips de la pièce à droite (cellule, valeur, page). Un clic sur une ligne montre la zone et sélectionne la cellule Excel ; les petits boutons T, ✓, ✗, Σ, ▦ filtrent la liste. Le bouton ▥ de la barre du haut l'affiche ou la masque.

**Retrouver une valeur** : cliquez une cellule Excel (par exemple `550`) : toutes les occurrences de cette valeur dans la pièce ouverte sont surlignées en orange (550,00 · -550,00 · 550,00 € — mais pas 1 550,00 ni 550,75). Une barre en haut indique « 1 / 3 » : flèches ou **F3** pour passer de l'une à l'autre, **Échap** pour effacer. Si la valeur n'est pas dans cette pièce, la barre propose les autres pièces qui la contiennent. (Désactivable dans **Paramètres**.)

Zoom : **Ctrl + molette**, boutons **↔** (largeur du texte : les marges blanches sont ignorées pour afficher le texte plus gros) et **▯** (page entière). Le zoom est retenu pour chaque document. Les longues opérations (OCR, import de dossier, recherche…) affichent une barre de progression avec **Annuler** (ou Échap). L'apparence suit le thème d'Office (sombre ou clair) ; vous pouvez la forcer dans **Paramètres**.

**Exemple — saisir un relevé bancaire :**
1. Cellule `A2` sélectionnée, bouton **Tableau** du ruban.
2. **Cliquez n'importe où dans le tableau** : il est détecté tout seul (cadre violet). Vous pouvez aussi l'entourer à la main.
3. Vérifiez l'aperçu : **×** retire une ligne (titre, sous-total…), **⇹** fusionne deux colonnes. Pour couper une colonne, cliquez dans le cadre violet ; glissez un trait pour le déplacer, double-cliquez pour le supprimer.
4. Indiquez « Pages 1 à N » si toutes les pages ont la même mise en page, puis **Insérer dans Excel** (ou Entrée).

**Exemple — pointer un grand livre :** sélectionnez la colonne des montants, puis **Rapprochement** dans le ruban.

**Exemple — justifier des totaux :** sélectionnez les totaux (ex. total des frais du mois), puis **Trouver les sommes** : les montants qui les composent sont surlignés sur la pièce.

---

## Mettre à jour le complément

1. Sur GitHub, dans votre dépôt : **Add file** → **Upload files**, glissez les fichiers du paquet de mise à jour, puis **Commit changes**. Les fichiers existants sont remplacés.
2. C'est tout : au prochain lancement, SnipCompta voit la nouvelle version (fichier `version.json`) et se recharge tout seul. Si vous étiez en train de travailler, un message propose **Mettre à jour** au moment qui vous arrange.
3. Seulement si la mise à jour change le ruban (c'est indiqué avec le paquet) : fermez Excel et relancez `Installer-SnipCompta.bat`.

## Bon à savoir

- **Taille du classeur** : chaque pièce enregistrée alourdit le fichier (environ sa propre taille × 1,35). Pour les gros volumes, décochez « Enregistrer les documents dans le classeur » dans l'onglet Documents : les liens restent, il suffira de réimporter le même fichier pour revoir la preuve.
- **OCR** : la première lecture d'un scan prend 10 à 20 secondes (chargement du moteur), puis quelques secondes par page. La qualité dépend du scan ; vérifiez les montants critiques.
- **Feuilles techniques** : SnipCompta crée deux feuilles très masquées, `_SnipData` (les liens) et `_SnipDocs` (les pièces). Ne les supprimez pas si vous voulez garder les liens.
- **Excel occupé** : si une cellule est en cours de saisie, SnipCompta attend et réessaie tout seul ; appuyez simplement sur Entrée ou Échap dans Excel. L'enregistrement des snips est lui aussi réessayé automatiquement.
- **Si l'onglet n'apparaît pas** : vérifiez que l'adresse du manifeste s'ouvre bien dans le navigateur, puis relancez `Installer-SnipCompta.bat`.

## Ce qui n'est pas (encore) inclus par rapport à DataSnipper

- Extraction par IA de formulaires (« Form Snip », relevés reconnus automatiquement), agents Excel.
- Documents sources Word ou Excel : convertissez-les en PDF (Fichier → Enregistrer sous → PDF).
- Capture de pages web (WebSnip) et travail collaboratif en ligne.

SnipCompta est un outil indépendant, sans lien avec DataSnipper.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `manifest.xml` | Carte d'identité du complément pour Excel (à modifier, étape 2) |
| `version.json` | Numéro de version : permet la mise à jour automatique |
| `Installer-SnipCompta.bat`, `Desinstaller-SnipCompta.bat` | Installation / retrait en un double-clic (fournis à part) |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | Le complément lui-même |
| `assets/` | Icônes (dont `assets/rb2/` pour les boutons du ruban) |
| `lib/` | Bibliothèques open source intégrées : pdf.js (lecture PDF, Apache 2.0), Tesseract.js (OCR, Apache 2.0), pdf-lib (édition PDF, MIT), JSZip (archives, MIT/GPLv3) |
