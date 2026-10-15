# SnipCompta — complément Excel pour vos pièces justificatives

SnipCompta ajoute un onglet **SNIPCOMPTA** au ruban d'Excel et un panneau à droite de la feuille. Vous y importez vos relevés bancaires, factures et scans (PDF, PNG, JPG), puis vous entourez à la souris ce que vous voulez récupérer. La valeur arrive dans la cellule Excel sélectionnée, avec un lien vers la preuve : cliquer sur la cellule réaffiche la pièce à l'endroit exact.

Vos documents ne quittent jamais votre ordinateur : la lecture, l'OCR et l'extraction se font dans Excel, et les pièces sont enregistrées dans le classeur lui-même.

---

## Ce que fait le complément

| Fonction | Ce qui se passe |
|---|---|
| **Texte** (Text Snip) | Entourez un texte ou un montant → il s'inscrit dans la cellule active (montants convertis en nombres, ex. « 1 484,86 » → 1484,86). Cellule en bleu. |
| **Somme** (Sum Snip) | Entourez un ou plusieurs montants, sur une ou plusieurs zones/pages → le total s'inscrit dans la cellule. Cellule en jaune. |
| **Valider** (Validation Snip) | Entourez l'élément vérifié → ✓ vert dans la cellule, coche verte sur la pièce. |
| **Exception** (Exception Snip) | Entourez l'anomalie → ✗ rouge dans la cellule, marque rouge sur la pièce. |
| **Commentaire** | Entourez une zone et tapez une remarque → commentaire Excel sur la cellule, lié à la pièce. |
| **Tableau** (Table Snip) | Entourez un tableau (ex. les opérations d'un relevé) → colonnes détectées automatiquement, ajustables, puis insertion de toutes les lignes, sur plusieurs pages d'un coup. Les lignes de suite (libellé sur 2 lignes) sont fusionnées. |
| **Recherche** | Cherche un texte ou un montant dans tous les documents importés. |
| **Trouver les sommes** | Sélectionnez des totaux dans Excel → SnipCompta retrouve dans les pièces la suite de montants qui donne chaque total. |
| **Supprimer des snips** | Dans la sélection ou tous : les liens et couleurs sont retirés, les valeurs restent. |
| **Exporter** | Crée une feuille « Piste d'audit » listant chaque snip (cellule cliquable, valeur, document, page, date). |
| **Rapprochement** | Sélectionnez une colonne Excel (montants, n° de facture, libellés) → chaque cellule retrouvée dans les pièces passe en vert avec un lien vers la preuve, les autres en rouge. |
| **Liens cellule ↔ preuve** | Cliquez une cellule « snipée » → la pièce s'ouvre à la bonne page, zone surlignée. Cliquez une zone colorée sur la pièce → Excel sélectionne la cellule. Les liens suivent les cellules si vous insérez des lignes. |
| **OCR** | Les scans et photos sont lus automatiquement (français + anglais). |

Raccourcis (quand le panneau a le focus) : `1` Naviguer · `2` Texte · `3` Somme · `4` Valider · `5` Exception · `6` Tableau · `7` Commentaire · `Entrée` insérer la somme / le tableau · `Échap` annuler · `↶` annule le dernier snip.

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
| `assets/` | Icônes (dont `assets/rb/` pour les boutons du ruban) |
| `lib/` | Bibliothèques open source intégrées : pdf.js (lecture PDF, Mozilla, licence Apache 2.0) et Tesseract.js (OCR, licence Apache 2.0) |
