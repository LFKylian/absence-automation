# Ô Talents - Gestion des absences

Application web Google Apps Script permettant d'identifier les élèves non présents à une séance de tutorat, de vérifier leurs contacts responsables, puis de préparer des brouillons Gmail à partir de modèles de messages.

L'application est pensée pour être utilisée depuis ordinateur et smartphone. Le choix des dossiers Drive se fait avec un navigateur intégré compatible mobile, sans Google Picker.

## Fonctionnalités principales

- Sélection d'un ou plusieurs dossiers Drive contenant les feuilles d'absences.
- Navigation dans `Mon Drive`, `Partagés avec moi`, `Récents` et `Suivis`.
- Lecture automatique des dates de séance disponibles dans les feuilles d'absences.
- Vérification fichier par fichier de la colonne correspondant à la date sélectionnée.
- Extraction des élèves absents, non renseignés ou avec statut inconnu.
- Association automatique avec les emails élève et parents depuis le fichier d'informations élèves.
- Prévisualisation et choix du modèle de mail par élève.
- Création de brouillons Gmail, sans envoi automatique.

## Utilisation

1. Ouvrir la webapp Apps Script déployée.
2. Dans `0 - Dossiers d'absence`, cliquer sur `Ajouter un dossier Drive`.
3. Choisir un ou plusieurs dossiers contenant les fichiers d'absences, puis confirmer.
4. Dans `1 - Date de la séance`, cliquer sur `Charger les dates`.
5. Naviguer jusqu'à la séance voulue avec les boutons précédent/suivant.
6. Cliquer sur `Vérifier la correspondance dans les fichiers`.
7. Contrôler les fichiers inclus dans l'extraction.
8. Cliquer sur `Lancer l'extraction`.
9. Vérifier les élèves détectés et choisir, si nécessaire, le modèle de mail à utiliser.
10. Cliquer sur `Créer les brouillons Gmail`.

Les brouillons sont créés dans Gmail. Aucun email n'est envoyé automatiquement.

## Données attendues

### Dossiers d'absences

Les dossiers sélectionnés doivent contenir des fichiers Google Sheets ou Excel lisibles par Drive. Chaque fichier doit contenir un onglet nommé `Feuil1`.

Par défaut, l'application lit les dates de séance sur la ligne `5`, à partir de la colonne `H`. Les élèves sont lu·e·s à partir de la ligne suivant les `6` lignes d'en-tête.

Les statuts reconnus comme présents sont notamment `présent`, `présente`, `present`, `presente` et leurs variantes en majuscules.

Les statuts reconnus comme absents sont notamment `abs`, `Absent`, `Absente`, `ABSENT`, `ABSENTE`.

Une valeur vide ou inconnue est signalée dans l'interface afin que l'utilisateur puisse la vérifier.

### Fichier d'informations élèves

Le fichier d'informations élèves sert à retrouver les emails de l'élève (P0) et des parents (P1, P2). Contrairement aux dossiers d'absences, son choix se fait de manière dynamique depuis l'interface via le navigateur de fichiers intégré. 

La configuration de la structure attendue de cet onglet se trouve sous forme de constantes au début de `Code.gs` (nom de l'onglet, nombre de lignes d'en-tête, colonnes des emails et du genre).

## Sélection des dossiers Drive

Le navigateur de dossiers intégré remplace Google Picker pour éviter les problèmes de sélection sur mobile.

Il permet de parcourir :

- `Mon Drive`
- `Partagés avec moi`
- `Récents`
- `Suivis`

Chaque ligne de dossier dispose d'une case de sélection et d'un bouton d'ouverture. Plusieurs dossiers peuvent être sélectionnés avant confirmation.

## Modèles de mails

Deux modèles sont configurés par défaut dans `Code.gs` :

- première absence ;
- absences répétées.

L'utilisateur peut aussi renseigner un modèle personnalisé depuis l'interface.

Les placeholders disponibles et interprétés par le code sont :

- `{{NOM}}`
- `{{PRENOM}}`
- `{{EMETTEUR}}`
- `{{JUSTIFICATIF_EMAIL}}`
- `{{COORDINATEUR_EMAIL}}`
- `{{e}}`, qui ajoute `e`, rien, ou `(e)` selon le genre renseigné.

## Développeur

### Structure du dépôt

- `Code.gs` : logique serveur Apps Script. Lit Drive, ouvre les fichiers Sheets/Excel, extrait les absences, associe les emails et crée les brouillons Gmail.
- `Interface.html` : interface web, CSS responsive, navigateur de dossiers Drive, appels `google.script.run` et rendu des résultats.
- `appsscript.json` : manifeste Apps Script, configuration du runtime, de la webapp, des scopes OAuth et du service Drive avancé.

### Manifeste Apps Script

Le projet utilise le runtime V8, le fuseau `Europe/Paris`, et le service avancé Drive API v2 exposé sous le symbole `Drive`.

La webapp est configurée avec :

- `executeAs: USER_ACCESSING` : les actions Drive/Gmail sont exécutées avec les droits de l'utilisateur connecté ;
- `access: ANYONE` : l'accès dépend du déploiement et de l'authentification Google.

Scopes utilisés :

- Drive : navigation et lecture des dossiers/fichiers ;
- Spreadsheets : lecture des feuilles d'absence et du fichier d'informations élèves ;
- Gmail compose : création de brouillons ;
- userinfo.email : affichage de l'utilisateur connecté ;
- script.container.ui et script.external_request : scopes Apps Script nécessaires au contexte webapp.

### Flux technique

1. `doGet()` sert `Interface.html`.
2. L'interface appelle `listFolders(parentId, source)` pour afficher les dossiers Drive.
3. `getAbsenceDates(folderIds)` lit les dates disponibles dans le premier fichier compatible trouvé.
4. `checkDateAcrossFiles(folderIds, targetCol, targetLabel)` vérifie la date sélectionnée dans chaque fichier.
5. L'interface construit `fileResolutions` avec les fichiers inclus et leur colonne résolue.
6. `getAbsentsWithEmails(fileResolutions, infoFileId)` lit les statuts, filtre les élèves non présents et ajoute les emails.
7. `getDefaultTemplates()` renvoie les modèles par défaut.
8. `createAllDrafts(draftsData)` crée les brouillons Gmail.

### Points de configuration importants

Les constantes en haut de `Code.gs` pilotent la structure des fichiers :

- `INFO_SHEET_NAME`, `INFO_HEADER_ROWS`
- `ABSENCE_SHEET_NAME`, `ABSENCE_HEADER_ROWS`
- `DATE_COLUMN`, `DATE_ROW`, `DATE_START_COL`
- `P0_EMAIL_COLUMN`, `P1_EMAIL_COLUMN`, `P2_EMAIL_COLUMN`, `GENDER_COLUMN`

En cas de changement de format des feuilles, ces constantes sont les premiers éléments à vérifier.

### Fichiers Excel

Les fichiers Excel sont convertis temporairement en Google Sheets via le service avancé Drive, puis envoyés à la corbeille après lecture. Le service Drive avancé doit donc rester activé dans Apps Script et dans `appsscript.json`.

### Déploiement

Après modification du code :

1. Pousser les fichiers dans le projet Apps Script, par exemple avec `clasp` si le dépôt est synchronisé.
2. Vérifier que le service avancé Drive API v2 est activé.
3. Créer ou mettre à jour le déploiement webapp.
4. Tester avec un compte ayant accès aux dossiers Drive et au compte Gmail cible.

## Sécurité et confidentialité

L'application lit des données élèves et prépares des emails aux responsables. Le dépôt public ne doit pas contenir de données personnelles réelles, d'exports de feuilles, de logs sensibles ou d'identifiants privés autres que les IDs de configuration explicitement assumés par l'équipe projet.