// ============================================================
// CONFIGURATION
// ============================================================
const INFO_SHEET_NAME  = "Tutoré.e.s 25-26";
const INFO_HEADER_ROWS = 3;

const ABSENCE_SHEET_NAME  = "Feuil1";
const ABSENCE_HEADER_ROWS = 6;
const DATE_COLUMN         = 12; // colonne par défaut (fallback)
const DATE_ROW            = 5;  // ligne contenant les dates de séance
const DATE_START_COL      = 8;  // colonne H (1-indexé)

const P0_EMAIL_COLUMN = 14;
const P1_EMAIL_COLUMN = 7;
const P2_EMAIL_COLUMN = 8;
const GENDER_COLUMN = 3; // Colonne C : "M" ou "F"

const NAME_COLUMN = 1;    // Colonne A
const SURNAME_COLUMN = 2; // Colonne B

// ── Valeurs reconnues (casse exacte, sans normalisation volontaire) ──
const PRESENT_VALUES = [
  "présent","présente","present","presente",
  "Présent","Présente","Present","Presente",
  "PRÉSENT","PRÉSENTE","PRESENT","PRESENTE",
  "Présent.e"
];
const ABSENT_VALUES = [
  "abs", "Abs",
  "absent","absente","Absent","Absente",
  "ABSENT","ABSENTE", "ABS",
  "Absent.e"
];

const FIRST_ABSENCE_SUBJECT  = "[Ô Talents] Absence de {{NOM}} {{PRENOM}}";
const FIRST_ABSENCE_TEMPLATE =
  "Madame, Monsieur,\n\n" +
  "{{NOM}} {{PRENOM}}, inscrit{{e}} dans la Cordée de la réussite INSA Toulouse a été identifié{{e}} " +
  "absent{{e}} à la séance de tutorat du jeudi 21 mai 2026.\n" +
  "Sauf erreur de notre part, cette absence n'avait pas été signalée préalablement, c'est pourquoi " +
  "nous vous demandons de bien vouloir adresser un mail à {{JUSTIFICATIF_EMAIL}} pour justifier son absence.\n\n" +
  "Nous vous rappelons que ce programme d'accompagnement repose sur le volontariat d'étudiants et que " +
  "toute absence nuit à la dynamique du groupe et à la motivation des participants.\n\n" +
  "Nous vous remercions de l'attention que vous porterez à notre demande.\n\n" +
  "Cordialement,\n{{EMETTEUR}} pour l'équipe Egalité des chances Ô Talents INSA Toulouse\n\n--\nLes coordinateurs\n" +
  "-------------------------------------------------\n" +
  "Programme Egalité des Chances\n" +
  "Association Ô Talents\n" +
  "INSA de Toulouse\n" +
  "{{COORDINATEUR_EMAIL}}\n" +
  "-------------------------------------------------";

const MUCH_ABSENCES_SUBJECT  = "[Ô Talents] Absences répétées — maintien de l'engagement de {{NOM}} {{PRENOM}}";
const MUCH_ABSENCES_TEMPLATE =
  "Madame, Monsieur,\n\n" +
  "{{NOM}} {{PRENOM}}, inscrit{{e}} dans la Cordée de la réussite INSA Toulouse a été identifié{{e}} absent{{e}} " +
  "à de nombreuses reprises, sans justifications.\n" +
  "Nous vous contactons donc pour nous enquérir de sa motivation à rester au sein du dispositif Ô Talents. " +
  "Si {{PRENOM}} ne participe plus aux séances de tutorat, nous devrons considérer son retrait du dispositif.\n\n" +
  "Nous restons à votre disposition pour tout renseignement.\n\n" +
  "Cordialement,\n{{EMETTEUR}} pour l'équipe Egalité des chances Ô Talents INSA Toulouse\n\n--\nLes coordinateurs\n" +
  "-------------------------------------------------\n" +
  "Programme Egalité des Chances\n" +
  "Association Ô Talents\n" +
  "INSA de Toulouse\n" +
  "{{COORDINATEUR_EMAIL}}\n" +
  "-------------------------------------------------";

function resetPermissions() {
  ScriptApp.invalidateAuth();
  Logger.log("Autorisation révoquée. Relancez le script pour accepter les nouvelles permissions.");
}

// ============================================================
// POINT D'ENTRÉE UI
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("Interface")
    .setTitle("Ô Talents — Gestion des absences");
}

// ============================================================
// CLASSIFICATION DU STATUT — pure function, testable sans Drive
// ============================================================
function classifyStatus(raw) {
  const trimmed = String(raw).trim();
  if (!trimmed)                       return "empty";
  if (PRESENT_VALUES.includes(trimmed)) return "present";
  if (ABSENT_VALUES.includes(trimmed))  return "absent";
  return "unknown";
}

// ============================================================
// CLASSIFICATION DU GENRE — pure function
// Retourne "M", "F", ou "unknown"
// ============================================================
function classifyGender(raw) {
  const val = String(raw).trim().toUpperCase();
  if (val === "M") return "M";
  if (val === "F") return "F";
  return "unknown";
}

// ============================================================
// ADRESSE MAIL DE L'UTILISATEUR CONNECTÉ
// Fonctionne uniquement avec executeAs: USER_ACCESSING
// ============================================================
function getUserEmail() {
  const email = Session.getActiveUser().getEmail();
  Logger.log(`👤 Utilisateur connecté : ${email}`);
  return email || "";
}

// ============================================================
// LECTURE DES DATES DE SÉANCE — appelée depuis l'UI
//
// Lit la ligne DATE_ROW à partir de la colonne DATE_START_COL.
// S'arrête dès que 2 cases consécutives sont vides.
// Retourne : [{ col: 8, label: "21/05/2026" }, ...]
// Lit depuis le 1er fichier trouvé (les dates sont identiques
// sur tous les fichiers d'un même dossier).
// ============================================================
function getAbsenceDates(folderIds) {
  if (!folderIds || folderIds.length === 0) {
    throw new Error("Sélectionnez au moins un dossier d'absences.");
  }
  const ids = folderIds;

  for (const folderId of ids) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const files  = folder.getFilesByType(MimeType.GOOGLE_SHEETS);

      // Inclut aussi les fichiers Excel
      const allFiles = [];
      const gSheets  = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
      const xlsx     = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
      const xlsxLeg  = folder.getFilesByType(MimeType.MICROSOFT_EXCEL_LEGACY);
      while (gSheets.hasNext()) allFiles.push(gSheets.next());
      while (xlsx.hasNext())    allFiles.push(xlsx.next());
      while (xlsxLeg.hasNext()) allFiles.push(xlsxLeg.next());

      if (allFiles.length === 0) continue;

      const { ss, tempId } = openSheetFile(allFiles[0]); // 1er fichier suffit pour les dates
      const sheet          = ss.getSheetByName(ABSENCE_SHEET_NAME);

      if (!sheet) { closeSheetFile(tempId); continue; }

      const lastCol = sheet.getLastColumn();
      if (lastCol < DATE_START_COL) { closeSheetFile(tempId); continue; }

      const numCols    = lastCol - DATE_START_COL + 1;
      const dateValues = sheet.getRange(DATE_ROW, DATE_START_COL, 1, numCols).getValues()[0];

      closeSheetFile(tempId);

      const dates = [];
      let emptyCount = 0;

      for (let i = 0; i < dateValues.length; i++) {
        const val     = dateValues[i];
        const isEmpty = (val === "" || val === null || val === undefined);

        if (isEmpty) {
          emptyCount++;
          if (emptyCount >= 2) break;
        } else {
          emptyCount = 0;
          const label = (val instanceof Date)
            ? Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy")
            : String(val).trim();
          dates.push({ col: DATE_START_COL + i, label });
        }
      }

      Logger.log(`📅 ${dates.length} date(s) : ${dates.map(d => d.label).join(", ")}`);
      return dates;

    } catch (e) {
      Logger.log(`❌ getAbsenceDates : ${e.message}`);
    }
  }
  return [];
}

// ============================================================
// TEMPLATES PAR DÉFAUT
// ============================================================
function getDefaultTemplates() {
  return {
    template1: { subject: FIRST_ABSENCE_SUBJECT,  body: FIRST_ABSENCE_TEMPLATE },
    template2: { subject: MUCH_ABSENCES_SUBJECT,   body: MUCH_ABSENCES_TEMPLATE }
  };
}

// ============================================================
// ORCHESTRATEUR — signature modifiée
//
// Reçoit fileResolutions : [{ fileId, resolvedCol }]
// construits par l'UI à partir des fichiers inclus après vérification.
// ============================================================
function getAbsentsWithEmails(fileResolutions, infoFileId) {
  if (!fileResolutions || fileResolutions.length === 0) return [];
  if (!infoFileId) throw new Error("Sélectionnez le fichier d'informations élèves.");

  const emailMap   = buildEmailMap(infoFileId);
  const allStudents = [];

  for (const { fileId, resolvedCol } of fileResolutions) {
    try {
      const file     = DriveApp.getFileById(fileId);
      const students = processAbsenceSheet(file, ABSENCE_SHEET_NAME, ABSENCE_HEADER_ROWS, resolvedCol);
      allStudents.push(...students);
      Logger.log(`  📄 ${file.getName()} — col.${resolvedCol} — ${students.length} élève(s) lu(s)`);
    } catch (e) {
      Logger.log(`❌ Fichier ${fileId} : ${e.message}`);
    }
  }

  const nonPresent = allStudents.filter(s => s.statusType !== "present");

  return nonPresent.map(({ name, surname, rawStatus, statusType, gender }) => {
    const key    = `${name.toLocaleLowerCase()}|${surname.toLocaleLowerCase()}`;
    const emails = emailMap.get(key) || { p0: "", p1: "", p2: "" };
    return { name, surname, emails, hasParentEmail: !!(emails.p1 || emails.p2), rawStatus, statusType, gender };
  });
}

// ============================================================
// LECTURE D'UNE FEUILLE D'ABSENCE
//
// Retourne TOUS les élèves (présents inclus) avec leur statut.
// Le filtrage se fait dans getAbsentsWithEmails.
// ============================================================
function processAbsenceSheet(file, sheetName, headerRows, dateColumn) {
  const { ss, tempId } = openSheetFile(file);
  const sheet          = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`⚠️ Onglet "${sheetName}" absent dans ${file.getName()}`);
    closeSheetFile(tempId);
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRows) { closeSheetFile(tempId); return []; }

  const data     = sheet.getRange(headerRows + 1, 1, lastRow - headerRows, sheet.getLastColumn()).getValues();
  const students = [];

  data.forEach((row, i) => {
    const name    = String(row[NAME_COLUMN - 1]).trim();
    const surname = String(row[SURNAME_COLUMN - 1]).trim();
    if (!name || !surname) return;

    const rawStatus  = String(row[dateColumn - 1]).trim();
    const statusType = classifyStatus(rawStatus);
    const gender     = classifyGender(row[GENDER_COLUMN - 1]);
    
    Logger.log(`    ${name} ${surname} — "${rawStatus}" (${statusType}) ligne ${headerRows + 1 + i}`);
    students.push({ name, surname, rawStatus, statusType, gender });
  });

  const counts = ["present","absent","unknown","empty"]
    .map(t => `${t}:${students.filter(s => s.statusType === t).length}`).join(" ");
  Logger.log(`  📊 ${counts}`);

  closeSheetFile(tempId);
  return students;
}

// ============================================================
// MAP NOM→EMAILS
// ============================================================
function buildEmailMap(infoFileId) {
  const emailMap = new Map();
  let   tempId   = null;

  try {
    const infoFile = DriveApp.getFileById(infoFileId);
    const opened = openSheetFile(infoFile);
    tempId       = opened.tempId;
    const sheet  = opened.ss.getSheetByName(INFO_SHEET_NAME);

    if (!sheet) {
      Logger.log(`⚠️ Onglet "${INFO_SHEET_NAME}" absent`);
      closeSheetFile(tempId);
      return emailMap;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= INFO_HEADER_ROWS) { closeSheetFile(tempId); return emailMap; }

    const data = sheet.getRange(
      INFO_HEADER_ROWS + 1, 1,
      lastRow - INFO_HEADER_ROWS,
      sheet.getLastColumn()
    ).getValues();

    data.forEach(row => {
      const key = `${String(row[NAME_COLUMN - 1]).trim().toLocaleLowerCase()}|${String(row[SURNAME_COLUMN - 1]).trim().toLocaleLowerCase()}`;
      emailMap.set(key, {
        p0: row[P0_EMAIL_COLUMN - 1] || "",
        p1: row[P1_EMAIL_COLUMN - 1] || "",
        p2: row[P2_EMAIL_COLUMN - 1] || ""
      });
    });

  } catch (e) {
    Logger.log(`❌ Erreur buildEmailMap : ${e.message}`);
  }

  if (tempId) {
  closeSheetFile(tempId);
  }
  
  return emailMap;
}

// ============================================================
// CRÉATION DES BROUILLONS
// ============================================================
function createAllDrafts(draftsData) {
  const results = [];

  draftsData.forEach(({ to, cc, subject, body }) => {
    try {
      GmailApp.createDraft(to, subject, body, { cc: cc || "" });
      results.push({ to, success: true });
      Logger.log(`✅ Brouillon créé pour : ${to}`);
    } catch (e) {
      results.push({ to, success: false, error: e.message });
      Logger.log(`❌ Échec brouillon pour ${to} : ${e.message}`);
    }
  });

  return results;
}

// ============================================================
// NAVIGATEUR DE DOSSIERS DRIVE — appelé depuis l'UI
//
// Remplace Google Picker pour une expérience mobile fiable.
// Supporte Mon Drive, Partagés avec moi, Récents et Suivis.
// ============================================================
function listFolders(parentId, source) {
  const selectedSource = source || "my-drive";
  const root = DriveApp.getRootFolder();

  if (!parentId && selectedSource !== "my-drive") {
    return listSpecialFolderView(selectedSource);
  }

  const folder = parentId ? DriveApp.getFolderById(parentId) : root;
  const folders = [];
  const it = folder.getFolders();

  while (it.hasNext()) {
    const child = it.next();
    folders.push(folderToBrowserItem(child));
  }

  sortBrowserFolders(folders);

  return {
    source: selectedSource,
    current: {
      id: parentId || "",
      name: parentId ? folder.getName() : getFolderSourceLabel(selectedSource)
    },
    parent: getFolderParentInfo(folder, root.getId(), selectedSource),
    breadcrumbs: buildFolderBreadcrumbs(folder, root.getId(), selectedSource),
    folders
  };
}

function listInfoFiles(parentId, source) {
  const selectedSource = source || "my-drive";
  const root = DriveApp.getRootFolder();

  if (!parentId && selectedSource !== "my-drive") {
    return listSpecialInfoFileView(selectedSource);
  }

  const folder = parentId ? DriveApp.getFolderById(parentId) : root;
  const folders = [];
  const fit = folder.getFolders();
  while (fit.hasNext()) folders.push(folderToBrowserItem(fit.next()));
  sortBrowserFolders(folders);

  const files = getAllSheetFiles(folder).map(file => ({
    id: file.getId(),
    name: file.getName(),
    mimeType: file.getMimeType()
  }));
  sortBrowserFolders(files);

  return {
    source: selectedSource,
    current: {
      id: parentId || "",
      name: parentId ? folder.getName() : getFolderSourceLabel(selectedSource)
    },
    parent: getFolderParentInfo(folder, root.getId(), selectedSource),
    breadcrumbs: buildFolderBreadcrumbs(folder, root.getId(), selectedSource),
    folders,
    files
  };
}

function listSpecialInfoFileView(source) {
  const config = getInfoFileSourceConfig(source);
  const items = listDriveItemsWithDriveApi(config.q, config.orderBy);
  const folders = items.filter(item => item.mimeType === MimeType.FOLDER)
    .map(item => ({ id: item.id, name: item.name, hasChildren: folderHasChildren(item.id) }));
  const files = items.filter(item => item.mimeType !== MimeType.FOLDER)
    .map(item => ({ id: item.id, name: item.name, mimeType: item.mimeType }));

  if (source !== "recent") {
    sortBrowserFolders(folders);
    sortBrowserFolders(files);
  }

  return {
    source,
    current: { id: "", name: config.label },
    parent: null,
    breadcrumbs: [{ id: "", name: config.label, source }],
    folders,
    files
  };
}

function getInfoFileSourceConfig(source) {
  const fileTypes = [
    "mimeType = 'application/vnd.google-apps.folder'",
    "mimeType = 'application/vnd.google-apps.spreadsheet'",
    "mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'",
    "mimeType = 'application/vnd.ms-excel'"
  ].join(" or ");
  const query = "(" + fileTypes + ") and trashed = false";

  if (source === "shared") return { label: "Partagés avec moi", q: query + " and sharedWithMe", orderBy: "title" };
  if (source === "recent") return { label: "Récents", q: query, orderBy: "recency desc" };
  if (source === "starred") return { label: "Suivis", q: query + " and starred = true", orderBy: "title" };
  return { label: "Mon Drive", q: query, orderBy: "title" };
}


function listSpecialFolderView(source) {
  const config = getFolderSourceConfig(source);
  const folders = listFoldersWithDriveApi(config.q, config.orderBy);
  if (source !== "recent") sortBrowserFolders(folders);

  return {
    source,
    current: { id: "", name: config.label },
    parent: null,
    breadcrumbs: [{ id: "", name: config.label, source }],
    folders
  };
}

function getFolderSourceConfig(source) {
  const folderQuery = "mimeType = 'application/vnd.google-apps.folder' and trashed = false";
  if (source === "shared") {
    return { label: "Partagés avec moi", q: folderQuery + " and sharedWithMe", orderBy: "title" };
  }
  if (source === "recent") {
    return { label: "Récents", q: folderQuery, orderBy: "recency desc" };
  }
  if (source === "starred") {
    return { label: "Suivis", q: folderQuery + " and starred = true", orderBy: "title" };
  }
  return { label: "Mon Drive", q: folderQuery, orderBy: "title" };
}

function listFoldersWithDriveApi(query, orderBy) {
  return listDriveItemsWithDriveApi(query, orderBy).map(item => ({
    id: item.id,
    name: item.name,
    hasChildren: folderHasChildren(item.id)
  }));
}

function listDriveItemsWithDriveApi(query, orderBy) {
  const itemsOut = [];
  let pageToken = null;

  do {
    const response = Drive.Files.list({
      q: query,
      orderBy: orderBy,
      maxResults: 100,
      pageToken: pageToken,
      fields: "items(id,title,mimeType),nextPageToken"
    });

    const items = response.items || [];
    items.forEach(item => {
      itemsOut.push({
        id: item.id,
        name: item.title,
        mimeType: item.mimeType
      });
    });

    pageToken = response.nextPageToken;
  } while (pageToken && itemsOut.length < 300);

  return itemsOut;
}

function folderToBrowserItem(folder) {
  return {
    id: folder.getId(),
    name: folder.getName(),
    hasChildren: folder.getFolders().hasNext()
  };
}

function folderHasChildren(folderId) {
  try {
    return DriveApp.getFolderById(folderId).getFolders().hasNext();
  } catch (e) {
    return true;
  }
}

function sortBrowserFolders(folders) {
  folders.sort((a, b) => a.name.localeCompare(b.name, "fr", { sensitivity: "base" }));
}

function getFolderSourceLabel(source) {
  return getFolderSourceConfig(source).label;
}

function getFolderParentInfo(folder, rootId, source) {
  const id = folder.getId();
  if (id === rootId) return null;

  const parents = folder.getParents();
  if (!parents.hasNext()) return { id: "", name: getFolderSourceLabel(source), source };

  const parent = parents.next();
  return {
    id: parent.getId() === rootId ? "" : parent.getId(),
    name: parent.getId() === rootId ? getFolderSourceLabel(source) : parent.getName(),
    source
  };
}

function buildFolderBreadcrumbs(folder, rootId, source) {
  const crumbs = [];
  let current = folder;
  let guard = 0;

  while (current && guard < 20) {
    const id = current.getId();
    crumbs.unshift({
      id: id === rootId ? "" : id,
      name: id === rootId ? getFolderSourceLabel(source) : current.getName(),
      source
    });

    if (id === rootId) break;
    const parents = current.getParents();
    if (!parents.hasNext()) {
      crumbs.unshift({ id: "", name: getFolderSourceLabel(source), source });
      break;
    }
    current = parents.next();
    guard++;
  }

  return crumbs;
}

// ============================================================
// UTILITAIRE
// ============================================================
function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
}

// ============================================================
// HELPER — formater une valeur de cellule en string affichable
// ============================================================
function formatDateValue(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return String(val).trim();
}

// ============================================================
// HELPER — vérifie si une valeur de cellule correspond au label
//
// En Apps Script, une cellule contenant une date est toujours
// lue comme un objet Date, quelle que soit son format d'affichage.
// On compare donc par composantes (jour/mois/année) pour être
// robuste aux différences de format.
// ============================================================
function datesMatch(val, targetLabel) {
  if (val === null || val === undefined || val === "") return false;

  if (val instanceof Date) {
    // Comparaison directe après formatage
    if (Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy") === targetLabel) return true;

    // Parsage du label dd/MM/yyyy → comparaison par composantes
    const parts = String(targetLabel).split("/");
    if (parts.length === 3) {
      const [d, m, y] = parts.map(Number);
      return val.getDate() === d && (val.getMonth() + 1) === m && val.getFullYear() === y;
    }
    return false;
  }

  // Valeur texte : comparaison directe
  return String(val).trim() === String(targetLabel).trim();
}

// ============================================================
// HELPER - OUVERTURE UNIVERSELLE — Google Sheets natif ou Excel
//
// Retourne { ss: Spreadsheet, tempId: string|null }
// tempId est non-null uniquement pour les fichiers convertis :
// il faut appeler closeSheetFile(tempId) après usage pour nettoyer.
//
// Nécessite le service Drive avancé (Drive API) activé dans le projet.
// ============================================================
function openSheetFile(file) {
  const mime = file.getMimeType();

  if (mime === MimeType.GOOGLE_SHEETS) {
    return { ss: SpreadsheetApp.open(file), tempId: null };
  }

  // Formats Excel acceptés
  if (mime === MimeType.MICROSOFT_EXCEL || mime === MimeType.MICROSOFT_EXCEL_LEGACY) {
    Logger.log(`  🔄 Conversion temporaire : ${file.getName()}`);
    const tempFile = Drive.Files.copy(
      { title: `__temp__${file.getName()}`, mimeType: MimeType.GOOGLE_SHEETS },
      file.getId()
    );
    return { ss: SpreadsheetApp.openById(tempFile.id), tempId: tempFile.id };
  }

  throw new Error(`Format non supporté (${mime}) pour le fichier : ${file.getName()}`);
}

function closeSheetFile(tempId) {
  if (tempId) {
    try {
      DriveApp.getFileById(tempId).setTrashed(true);
      Logger.log(`  🗑️ Fichier temporaire supprimé : ${tempId}`);
    } catch (e) {
      Logger.log(`  ⚠️ Impossible de supprimer le fichier temporaire ${tempId} : ${e.message}`);
    }
  }
}

// ============================================================
// HELPER - récupère tous les fichiers Sheet+Excel d'un dossier
// ============================================================
function getAllSheetFiles(folder) {
  const files   = [];
  const gSheets = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
  const xlsx    = folder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  const xlsxLeg = folder.getFilesByType(MimeType.MICROSOFT_EXCEL_LEGACY);
  while (gSheets.hasNext()) files.push(gSheets.next());
  while (xlsx.hasNext())    files.push(xlsx.next());
  while (xlsxLeg.hasNext()) files.push(xlsxLeg.next());
  return files;
}

// ============================================================
// VÉRIFICATION DE CORRESPONDANCE PAR FICHIER — appelée depuis l'UI
//
// Pour chaque fichier des dossiers, cherche la colonne correspondant
// à la date sélectionnée (targetCol = colonne de référence, targetLabel
// = label affiché dans le navigateur).
//
// Retourne un tableau JSON-sérialisable :
// [{
//   fileId, fileName,
//   status: "ok" | "partial" | "missing",
//   resolvedCol: number | null,  // colonne à utiliser pour l'extraction
//   resolvedLabel: string | null, // valeur trouvée dans la cellule
//   reason: string               // message explicatif pour l'UI
// }]
// ============================================================

function checkDateAcrossFiles(folderIds, targetCol, targetLabel) {
  if (!folderIds || folderIds.length === 0) {
    throw new Error("Sélectionnez au moins un dossier d'absences.");
  }
  const ids     = folderIds;
  const results = [];

  for (const folderId of ids) {
    try {
      const folder   = DriveApp.getFolderById(folderId);
      const allFiles = getAllSheetFiles(folder);

      for (const file of allFiles) {
        const fileId   = file.getId();
        const fileName = file.getName();
        let tempId     = null;

        try {
          const opened = openSheetFile(file);
          tempId       = opened.tempId;
          const sheet  = opened.ss.getSheetByName(ABSENCE_SHEET_NAME);

          if (!sheet) {
            results.push({ fileId, fileName, status: "missing", resolvedCol: null, resolvedLabel: null,
              reason: `Onglet "${ABSENCE_SHEET_NAME}" introuvable dans ce fichier` });
            closeSheetFile(tempId);
            continue;
          }

          const lastCol = sheet.getLastColumn();
          if (lastCol < DATE_START_COL) {
            results.push({ fileId, fileName, status: "missing", resolvedCol: null, resolvedLabel: null,
              reason: "Aucune colonne de date à partir de la colonne H" });
            closeSheetFile(tempId);
            continue;
          }

          const numCols    = lastCol - DATE_START_COL + 1;
          const dateValues = sheet.getRange(DATE_ROW, DATE_START_COL, 1, numCols).getValues()[0];

          // Étape 1 : colonne de référence
          const targetIdx = targetCol - DATE_START_COL;
          if (targetIdx >= 0 && targetIdx < dateValues.length) {
            const val = dateValues[targetIdx];
            if (val !== "" && val !== null && datesMatch(val, targetLabel)) {
              results.push({ fileId, fileName, status: "ok", resolvedCol: targetCol,
                resolvedLabel: formatDateValue(val),
                reason: `Correspondance exacte en colonne ${targetCol}` });
              closeSheetFile(tempId);
              continue;
            }
          }

          // Étape 2 : recherche dans les autres colonnes
          let found = false;
          for (let i = 0; i < dateValues.length; i++) {
            const val = dateValues[i];
            if (val === "" || val === null) continue;
            if (datesMatch(val, targetLabel)) {
              const foundCol = DATE_START_COL + i;
              results.push({ fileId, fileName, status: "partial", resolvedCol: foundCol,
                resolvedLabel: formatDateValue(val),
                reason: `Date trouvée en colonne ${foundCol} au lieu de la colonne de référence ${targetCol} — layout différent` });
              found = true;
              break;
            }
          }

          if (!found) {
            results.push({ fileId, fileName, status: "missing", resolvedCol: null, resolvedLabel: null,
              reason: "Aucune séance à cette date dans ce fichier" });
          }

          closeSheetFile(tempId);

        } catch (e) {
          closeSheetFile(tempId);
          results.push({ fileId, fileName, status: "missing", resolvedCol: null, resolvedLabel: null,
            reason: `Erreur de lecture : ${e.message}` });
        }
      }
    } catch (e) {
      Logger.log(`❌ Dossier ${folderId} : ${e.message}`);
    }
  }

  Logger.log(`📋 Vérification : ${results.length} fichier(s) analysé(s)`);
  results.forEach(r => Logger.log(`  [${r.status}] ${r.fileName} — ${r.reason}`));
  return results;
}
