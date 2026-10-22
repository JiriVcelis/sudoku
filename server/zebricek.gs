/**
 * Sudoku – společný žebříček v Google tabulce.
 *
 * Nasazení: Rozšíření → Apps Script → vlož tento kód → Nasadit → Nové nasazení → Webová aplikace
 *   Spustit jako: Já, Přístup: Kdokoli. Adresu končící /exec vlož do config.js (scoresUrl).
 * Mazání žebříčku z aplikace: Nastavení projektu → Vlastnosti skriptu → ADMIN_KEY = tvůj debug kód.
 * Při první instalaci jednou spusť funkci setup() (založí listy a vloží dosavadní výsledky).
 */
const SHEETS = { "": "Výsledky", "main": "Výsledky", "test": "Test" };
const LEVELS = ["novice", "beginner", "easy", "medium", "hard", "expert"];
const HEADER = ["id", "level", "name", "secs", "errors", "hints", "at", "debug", "validace"];
// Počet čísel k doplnění podle obtížnosti a minimální čas na jedno číslo; rychlejší výsledky čekají na ověření.
const EMPTY = { novice: 21, beginner: 27, easy: 33, medium: 41, hard: 49, expert: 56 };
const MIN_SECS_PER_CELL = 2;
const REVIEW = "CHECK"; // ke schválení přepiš na OK (nebo smaž), k zamítnutí smaž řádek

function sheet_(key) {
  const name = SHEETS[String(key || "")] || SHEETS[""];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(HEADER); sh.setFrozenRows(1); }
  if (sh.getRange(1, HEADER.length).getValue() === "") sh.getRange(1, HEADER.length).setValue(HEADER[HEADER.length - 1]);
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function rows_(sh) {
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, HEADER.length).getValues() : [];
}
function int_(v, max) {
  const n = Math.round(Number(v));
  return isFinite(n) && n >= 0 && n <= max ? n : null;
}

// Čtení: všechny výsledky daného listu.
function doGet(e) {
  const sh = sheet_(e && e.parameter && e.parameter.sheet);
  // výsledky čekající na ověření se nezobrazují (schválení = ve sloupci validace napsat OK nebo text smazat)
  const results = rows_(sh).filter(r => r[0] !== "" && (String(r[8]).trim() === "" || String(r[8]).trim().toUpperCase() === "OK")).map(r => ({
    id: String(r[0]), level: String(r[1]), name: String(r[2]).replace(/^'/, ""),
    secs: Number(r[3]) || 0, errors: Number(r[4]) || 0, hints: Number(r[5]) || 0,
    at: Number(r[6]) || 0, debug: r[7] === true || String(r[7]).toUpperCase() === "TRUE"
  }));
  return json_({ ok: true, results: results });
}

// Zápis: přidání výsledku, nebo smazání (jen s ADMIN_KEY).
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const sh = sheet_(d.sheet);

    if (d.action === "add") {
      const x = d.entry || {};
      const id = String(x.id || "").slice(0, 40);
      if (!/^[\w-]+$/.test(id)) return json_({ ok: false, error: "id" });
      if (LEVELS.indexOf(x.level) < 0) return json_({ ok: false, error: "level" });
      let name = String(x.name || "").replace(/\s+/g, " ").trim().slice(0, 30);
      if (!name) return json_({ ok: false, error: "name" });
      if (/^[=+\-@]/.test(name)) name = "'" + name; // aby se jméno nebralo jako vzorec
      const secs = int_(x.secs, 86400), errors = int_(x.errors, 999), hints = int_(x.hints, 81);
      if (secs === null || errors === null || hints === null) return json_({ ok: false, error: "values" });
      const at = int_(x.at, 9e15) || Date.now();
      const ids = rows_(sh).map(r => String(r[0]));
      if (ids.indexOf(id) >= 0) return json_({ ok: true, duplicate: true });
      const review = !x.debug && secs < MIN_SECS_PER_CELL * EMPTY[x.level];
      sh.appendRow([id, x.level, name, secs, errors, hints, at, !!x.debug, review ? REVIEW : ""]);
      return json_({ ok: true, review: review });
    }

    if (d.action === "reset" || d.action === "resetDebug") {
      const key = PropertiesService.getScriptProperties().getProperty("ADMIN_KEY");
      if (!key || String(d.key || "") !== key) return json_({ ok: false, error: "denied" });
      const data = rows_(sh);
      let removed = 0;
      for (let i = data.length - 1; i >= 0; i--) {
        const isDbg = data[i][7] === true || String(data[i][7]).toUpperCase() === "TRUE";
        if (d.action === "reset" || isDbg) { sh.deleteRow(i + 2); removed++; }
      }
      return json_({ ok: true, removed: removed });
    }

    return json_({ ok: false, error: "action" });
  } catch (err) {
    return json_({ ok: false, error: "server" });
  } finally {
    lock.releaseLock();
  }
}

// Jednorázově: založí listy a do ostrého žebříčku vloží dosavadní výsledky.
function setup() {
  const sh = sheet_("");
  sheet_("test");
  if (rows_(sh).length) return;
  [
    ["1792502879782-xb28", "novice", "Jirka", 80, 0, 0, 1792502879782, false],
    ["1792501375769-x5jw", "novice", "Jirka", 94, 0, 0, 1792501375769, false],
    ["1792503701448-mgnl", "novice", "Jirka", 115, 1, 0, 1792503701448, false],
    ["1792500647934-ilem", "easy", "Jirka", 337, 7, 0, 1792500647934, false]
  ].forEach(r => sh.appendRow(r));
}
