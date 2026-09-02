/* ============================================================
   BASE DE RÉPONSES QUIZZ SVT — GOOGLE APPS SCRIPT
   ============================================================
   INSTALLATION (5 minutes, une seule fois) :

   1. Va sur https://script.google.com → « Nouveau projet ».
   2. Colle TOUT ce fichier dans l'éditeur, puis enregistre.
   3. En haut, clique sur le sélecteur de projet (à côté de
      « Services + ») → onglet « Feuilles » → « Créer une
      feuille de calcul ». (Alternative : crée une feuille
      manuellement et mets son ID dans SPREADSHEET_ID.)
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
const SPREADSHEET_ID = "";
const SHEET_NAME = "Réponses";

/* Réception d'un envoi d'élève (POST JSON) */
function doPost(e) {
  return handle(e && e.postData ? e.postData.contents : "");
}
/* Test navigateur : GET ?d=<json> ou GET simple (ping) */
function doGet(e) {
  if (!e || !e.parameter || !e.parameter.d) {
    return ContentService.createTextOutput("OK — base quizz SVT en ligne.")
      .setMimeType(ContentService.MimeType.TEXT);
  }
  return handle(e.parameter.d);
}

function handle(raw) {
  try {
    const o = JSON.parse(raw || "{}");
    const sheet = getSheet();
    const q = QUIZ_INFO[o.quiz] || { n: "?", total: 10 };
    const score = computeScore(o.a, o.quiz);
    const key = ((o.p || "") + "|" + (o.n || "") + "|" + (o.c || "")).trim().toUpperCase();
    /* un redépôt remplace la ligne précédente de l'élève */
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1, maxK = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][7]).trim().toUpperCase() === key) { rowIdx = i + 1; maxK = Number(data[i][6]) || 1; }
    }
    const k = rowIdx > 0 ? maxK + 1 : Math.max(1, Number(o.k) || 1);
    const row = [
      new Date(),
      o.p || "", o.n || "", o.c || "",
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
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Aucune feuille : lie une feuille au script ou remplis SPREADSHEET_ID.");
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(["Date", "Prénom", "Nom", "Classe", "Quiz", "ID quiz", "Dépôts", "Clé élève", "Score", "Réponses (numéros)", "Code NDJSON"]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 11).setFontWeight("bold");
  }
  return sh;
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
  eb15: { n: 15, total: 10, a: [0,1,2,3,0,1,2,3,0,1] }
};

function computeScore(a, quizId) {
  const info = QUIZ_INFO[quizId];
  if (!info || !info.a || !a) return "?";
  let s = 0;
  for (let i = 0; i < info.a.length; i++) if (a[i] === info.a[i]) s++;
  return s;
}
