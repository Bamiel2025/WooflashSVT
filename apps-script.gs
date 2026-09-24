/* ============================================================
   BASE DE RÉPONSES QUIZZ SVT — GOOGLE APPS SCRIPT
   ============================================================
   INSTALLATION (5 minutes, une seule fois) :

   1. Va sur https://script.google.com → « Nouveau projet ».
   2. Colle TOUT ce fichier dans l'éditeur, puis enregistre.
   3. En haut, clique sur le sélecteur de projet (à côté de
      « Services + ») → onglet « Feuilles ».
      (Optionnel : tu peux créer une feuille ici pour lier
      le script à un classeur existant ; sinon, la feuille
      sera créée automatiquement au premier envoi.)
   4. Clique sur « Déployer » → « Nouveau déploiement » →
      type « Application Web » :
        - Exécuter en tant que : Moi
        - Accès : « Tout le monde » (Anyone)
      → Déployer → autorise ton compte Google.
   5. Copie l'URL de l'application web
      (https://script.google.com/macros/s/.../exec) et colle-la
      dans index.html à la ligne :
        const SHEETS_URL = "...";

   OÙ VOIR LES RÉPONSES : ouvre la feuille Google Sheets liée.
   Chaque envoi d'élève ajoute une ligne : date, prénom, nom,
   classe, quiz, score, réponses. Un redépôt du même élève
   remplace sa ligne précédente (compteur de dépôts conservé).
   ============================================================ */

// Laisser "" pour utiliser la feuille liée au script (recommandé).
var SPREADSHEET_ID = "";
var SHEET_NAME = "Réponses";

/* Réception d'un envoi d'élève (POST JSON) */
function doPost(e) {
  return handle(e && e.postData ? e.postData.contents : "");
}
/* Test navigateur : GET ?d=<json> ou GET simple (ping) */
function doGet(e) {
  if (e && e.parameter && e.parameter.json === "1") {
    return getSubmissionsJson();
  }
  if (!e || !e.parameter || !e.parameter.d) {
    return ContentService.createTextOutput("OK — base quizz SVT en ligne.")
      .setMimeType(ContentService.MimeType.TEXT);
  }
  return handle(e.parameter.d);
}

/* Sheets convertit une classe comme "3E1" en nombre (3,00E+01) : on force le format
   texte pour ces cas afin de garder un affichage propre dans la feuille. */
function sheetText(v) {
  var s = String(v == null ? "" : v);
  return (/^[0-9]+[Ee][0-9]+$/.test(s)) ? "'" + s : s;
}

function handle(raw) {
  try {
var o = JSON.parse(raw || "{}");
  var sheet = getSheet();
    var q = QUIZ_INFO[o.quiz] || { n: "?", total: 10 };
  var score = computeScore(o.a, o.quiz);
  var key = ((o.p || "") + "|" + (o.n || "") + "|" + (o.c || "")).trim().toUpperCase();
  /* un redépôt remplace la ligne précédente de l'élève */
  var data = sheet.getDataRange().getValues();
    var rowIdx = -1, maxK = 0;
    /* i = 0 : la feuille n'a pas d'en-tête par défaut — ne jamais sauter la 1re ligne.
       Une ligne d'en-tête éventuelle ne matchera aucune clé élève. */
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][7]).trim().toUpperCase() === key) { rowIdx = i + 1; maxK = Number(data[i][6]) || 1; }
    }
  var k = rowIdx > 0 ? maxK + 1 : Math.max(1, Number(o.k) || 1);
  var row = [
      new Date(),
      o.p || "", o.n || "", sheetText(o.c),
      "Quiz " + q.n, o.quiz,
      k, key,
      score + "/" + q.total,
      (o.a || []).map(function (v) { return v == null ? "-" : (v + 1); }).join(","),
      raw || ""
    ];
    if (rowIdx > 0) sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
    return ContentService.createTextOutput("OK")
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput("ERREUR: " + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function getSheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty("SS_ID");
  var ss = null;
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SPREADSHEET_ID
      ? SpreadsheetApp.openById(SPREADSHEET_ID)
      : SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    ss = SpreadsheetApp.create("Base de réponses quizz SVT");
    props.setProperty("SS_ID", ss.getId());
  } else if (!ssId && SPREADSHEET_ID) {
    props.setProperty("SS_ID", ss.getId());
  }
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    var sheets = ss.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName() !== SHEET_NAME && sheets[i].getLastRow() === 0) {
        ss.deleteSheet(sheets[i]);
      }
    }
  }
  return sh;
}

/* Endpoint GET ?json=1 pour récupérer toutes les réponses (NDJSON) */
function getSubmissionsJson() {
  try {
    var sheet = getSheet();
    var data = sheet.getDataRange().getValues();
    var lines = [];
    /* i = 0 : la feuille créée par ce script n'a PAS d'en-tête — partir de 1
       laissait toujours la 1re ligne hors export. Une ligne d'en-tête
       éventuelle est écartée par isValidRaw/rebuildRaw. */
    for (var i = 0; i < data.length; i++) {
      if (!data[i]) continue;
      var raw = data[i][10]; // colonne Code NDJSON
      /* Repli : si le code brut est absent ou illisible (cellule tronquée/modifiée),
         on reconstruit un code valide depuis les colonnes : aucune ligne n'est perdue. */
      if (!isValidRaw(raw)) raw = rebuildRaw(data[i]);
      if (raw && isValidRaw(raw)) lines.push(raw);
    }
    var out = ContentService.createTextOutput(lines.join("\n"))
      .setMimeType(ContentService.MimeType.TEXT);
    return out;
  } catch (err) {
    return ContentService.createTextOutput("ERREUR: " + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/* Un code brut est valide s'il désigne un quiz connu avec le bon nombre de réponses */
function isValidRaw(raw) {
  try {
    var o = JSON.parse(raw);
    return !!(o && o.quiz && QUIZ_INFO[o.quiz] && o.a && o.a.length === QUIZ_INFO[o.quiz].total);
  } catch (e) { return false; }
}

/* Reconstruit un code NDJSON depuis les colonnes d'une ligne :
   B=prénom, C=nom, D=classe, F=id quiz, G=dépôts, J=réponses (1-4, "-"=sans réponse), A=date */
function rebuildRaw(row) {
  try {
    var quiz = row[5], info = QUIZ_INFO[quiz];
    if (!info) return "";
    var parts = String(row[9] == null ? "" : row[9]).split(",");
    if (parts.length !== info.total) return "";
    var a = [];
    for (var i = 0; i < parts.length; i++) {
      var v = parts[i].trim();
      if (v === "-" || v === "") { a.push(null); continue; }
      var n = parseInt(v, 10) - 1;
      if (isNaN(n) || n < 0 || n > 3) return "";
      a.push(n);
    }
    var d = row[0];
    var o = {
      v: 2, quiz: quiz,
      p: String(row[1] || ""), n: String(row[2] || ""),
      c: String(row[3] == null ? "" : row[3]),
      k: Number(row[6]) || 1, a: a,
      t: (d instanceof Date) ? d.toISOString() : new Date().toISOString()
    };
    return JSON.stringify(o);
  } catch (e) { return ""; }
}

/* Corrigé minimal (id quiz → n° affiché, nb questions, bonnes réponses en index 0-3) */
var QUIZ_INFO = {
  eb01: { n: 1,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb02: { n: 2,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb03: { n: 3,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb04: { n: 4,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb05: { n: 5,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb06: { n: 6,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb07: { n: 7,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb08: { n: 8,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb09: { n: 9,  total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb10: { n: 10, total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb11: { n: 11, total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb12: { n: 12, total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb13: { n: 13, total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb14: { n: 14, total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb15: { n: 15, total: 10, a: [0,1,2,3,0,1,2,3,0,1] },
  eb16: { n: 16, total: 10, a: [0,1,2,3,0,1,2,3,0,2] },
  eb17: { n: 17, total: 10, a: [0,1,2,3,0,1,2,3,1,2] },
  eb18: { n: 18, total: 10, a: [0,1,2,3,0,1,2,3,1,2] },
  eb19: { n: 19, total: 10, a: [3,0,1,2,1,2,0,1,2,1] },
  eb20: { n: 20, total: 10, a: [0,1,0,2,1,3,2,0,3,1] },
  eb21: { n: 21, total: 10, a: [0,1,2,0,3,1,2,3,0,1] },
  eb22: { n: 22, total: 10, a: [1,0,2,3,1,0,3,2,1,0] },
  eb23: { n: 23, total: 10, a: [2,0,1,3,0,2,1,3,2,0] },
  eb24: { n: 24, total: 10, a: [0,2,1,3,2,0,1,3,0,2] },
  eb25: { n: 25, total: 10, a: [1,3,0,2,1,0,3,2,0,1] },
  eb26: { n: 26, total: 10, a: [0,1,3,2,0,2,1,3,1,0] },
  eb27: { n: 27, total: 10, a: [2,0,1,3,1,2,0,3,2,0] },
  eb28: { n: 28, total: 10, a: [0,3,1,2,0,1,3,2,1,0] },
  eb29: { n: 29, total: 10, a: [1,0,2,1,3,0,2,3,1,2] },
  eb30: { n: 30, total: 10, a: [0,1,2,0,1,3,0,2,3,1] },
  eb31: { n: 31, total: 10, a: [2,0,1,3,0,2,1,0,3,2] },
  eb32: { n: 32, total: 10, a: [0,1,0,2,3,1,0,2,1,3] },
  eb33: { n: 33, total: 10, a: [1,2,0,3,2,0,1,3,0,2] },
  eb34: { n: 34, total: 10, a: [0,3,2,1,0,1,3,2,0,1] },
  eb35: { n: 35, total: 10, a: [1,0,3,2,0,2,1,3,1,0] },
  eb36: { n: 36, total: 10, a: [0,1,2,3,1,0,2,0,3,1] },
  eb37: { n: 37, total: 10, a: [2,0,1,0,3,2,1,3,0,1] },
  eb38: { n: 38, total: 10, a: [0,2,1,3,0,1,2,0,3,1] },
  eb39: { n: 39, total: 10, a: [0,1,2,3,1,2,0,1,0,2] }
};

function computeScore(a, quizId) {
  var info = QUIZ_INFO[quizId];
  if (!info || !info.a || !a) return "?";
  var s = 0;
  for (var i = 0; i < info.a.length; i++) if (a[i] === info.a[i]) s++;
  return s;
}
