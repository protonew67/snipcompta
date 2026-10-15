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
- OCR des scans (français + anglais), fait sur votre PC.

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

### Étape 3 — Déclarer le complément dans Excel (dossier partagé)

1. Créez un dossier, par exemple `C:\SnipCompta-manifeste`, et copiez-y le `manifest.xml` modifié.
2. Clic droit sur ce dossier → **Propriétés** → onglet **Partage** → **Partager…** → choisissez votre nom → **Partager**. Notez le **chemin réseau** affiché, du type `\\NOM-DU-PC\SnipCompta-manifeste`.
3. Dans Excel : **Fichier** → **Options** → **Centre de gestion de la confidentialité** → **Paramètres du Centre de gestion de la confidentialité…** → **Catalogues de compléments approuvés**.
4. Dans *URL du catalogue*, collez le chemin réseau → **Ajouter un catalogue** → cochez **Afficher dans le menu** → **OK** → **OK**.
5. **Fermez complètement Excel et rouvrez-le.**
6. Onglet **Accueil** → **Compléments** (ou **Insertion** → **Mes compléments**) → onglet **DOSSIER PARTAGÉ** → **SnipCompta** → **Ajouter**.
7. Un onglet **SNIPCOMPTA** apparaît dans le ruban, avec ces groupes :
   - **Importer** : Importer des documents · Organiser · Visionneuse
   - **Snips** : Texte · Validation · Exception · Somme · Tableau · Commentaire
   - **Modifier** : Annuler · Supprimer des snips (dans la sélection / tous)
   - **Automatisation** : Rapprochement · Trouver les sommes
   - **Documents** : Inclure les PDF dans Excel · Exporter · Retirer · OCR
   - **Aide** : Aide · Paramètres

> **Variante plus rapide (Microsoft 365 récent)** : **Accueil** → **Compléments** → **Plus de compléments** → **Mes compléments** → **Charger mon complément** → choisissez `manifest.xml`.

---

## Utilisation au quotidien

1. Onglet **SNIPCOMPTA** → **Importer des documents** (ou glissez vos PDF dans le panneau).
2. Cliquez la cellule de destination dans Excel.
3. Cliquez un outil dans le ruban (Texte, Validation, Exception, Somme, Tableau, Commentaire) et **entourez la zone** sur la pièce.

**Exemple — saisir un relevé bancaire :**
1. Cellule `A2` sélectionnée, bouton **Tableau** du ruban.
2. Entourez toutes les opérations de la page 1 (sans le titre ni les en-têtes de colonnes).
3. Vérifiez l'aperçu. Si une colonne est mal coupée : cliquez dans le cadre violet pour ajouter un séparateur, glissez-le pour le déplacer, double-cliquez pour le supprimer.
4. Indiquez « Pages 1 à N » si toutes les pages ont la même mise en page, puis **Insérer dans Excel**.

**Exemple — pointer un grand livre :** sélectionnez la colonne des montants, puis **Rapprochement** dans le ruban.

**Exemple — justifier des totaux :** sélectionnez les totaux (ex. total des frais du mois), puis **Trouver les sommes** : les montants qui les composent sont surlignés sur la pièce.

---

## Mettre à jour le complément

1. Sur GitHub, dans votre dépôt : **+** → **Upload files**, glissez le contenu du paquet de mise à jour (fichiers et dossiers), puis **Commit changes**. Les fichiers existants sont remplacés.
2. Remplacez `manifest.xml` dans votre dossier `SnipCompta-manifeste` par la nouvelle version.
3. Fermez Excel, videz le dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (collez ce chemin dans la barre de l'Explorateur), rouvrez Excel.

## Bon à savoir

- **Taille du classeur** : chaque pièce enregistrée alourdit le fichier (environ sa propre taille × 1,35). Pour les gros volumes, décochez « Enregistrer les documents dans le classeur » dans l'onglet Documents : les liens restent, il suffira de réimporter le même fichier pour revoir la preuve.
- **OCR** : la première lecture d'un scan prend 10 à 20 secondes (chargement du moteur), puis quelques secondes par page. La qualité dépend du scan ; vérifiez les montants critiques.
- **Feuilles techniques** : SnipCompta crée deux feuilles très masquées, `_SnipData` (les liens) et `_SnipDocs` (les pièces). Ne les supprimez pas si vous voulez garder les liens.
- **Mise à jour du complément** : remplacez les fichiers sur GitHub. Si Excel garde l'ancienne version, fermez Excel et supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.
- **Si l'onglet n'apparaît pas** : vérifiez que l'adresse du manifeste s'ouvre bien dans le navigateur, que le dossier est bien partagé et qu'Excel a été redémarré.

## Ce qui n'est pas (encore) inclus par rapport à DataSnipper

- Extraction par IA de formulaires (« Form Snip », relevés reconnus automatiquement), agents Excel.
- Documents sources Word ou Excel : convertissez-les en PDF (Fichier → Enregistrer sous → PDF).
- Capture de pages web (WebSnip) et travail collaboratif en ligne.

SnipCompta est un outil indépendant, sans lien avec DataSnipper.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `manifest.xml` | Carte d'identité du complément pour Excel (à modifier, étape 2) |
| `taskpane.html`, `taskpane.css`, `taskpane.js` | Le complément lui-même |
| `assets/` | Icônes (dont `assets/rb2/` pour les boutons du ruban) |
| `lib/` | Bibliothèques open source intégrées : pdf.js (lecture PDF, Apache 2.0), Tesseract.js (OCR, Apache 2.0), pdf-lib (édition PDF, MIT), JSZip (archives, MIT/GPLv3) |
