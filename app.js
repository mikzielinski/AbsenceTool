"use strict";

const MONTHS = [
  "Sep-25", "Oct-25", "Nov-25", "Dec-25",
  "Jan-26", "Feb-26", "Mar-26", "Apr-26",
  "May-26", "Jun-26", "Jul-26", "Aug-26",
];

const FLEXI_PROJECT_MAP = {
  "1402 holidays (normal vacation days)": { type: "Vacation", code: "748" },
  "1405 company days": { type: "Extra vacation days", code: "629" },
};

const BALANCE_TYPE_MAP = {
  "1402 holidays (normal vacation days)": "Holiday",
  "1403 special holidays": "Special",
};

const RED_FILL = { fgColor: { rgb: "FFFF0000" } };
const HEADER_FILL = { fgColor: { rgb: "FF4472C4" } };
const HEADER_FONT = { bold: true, color: { rgb: "FFFFFFFF" } };

const state = {
  sourceWorkbook: null,
  sourceWorkbookName: "",
  sourceSheets: [],
  p3WorkbookA: null,
  p3WorkbookAName: "",
  p3SheetsA: [],
  p3WorkbookB: null,
  p3WorkbookBName: "",
  p3SheetsB: [],
  masterWorkbook: null,
  masterWorkbookName: "",
  masterWorkbookBytes: null,
  masterSheets: [],
  masterWorkbookP2: null,
  masterWorkbookNameP2: "",
  masterSheetsP2: [],
};

// Wipe ALL workbook data from state.
// Called at the start of every process — nothing carried over between runs.
function clearWorkbookState() {
  state.sourceWorkbook      = null;
  state.masterWorkbook      = null;
  state.masterWorkbookBytes = null;
  state.p3WorkbookA         = null;
  state.p3WorkbookB         = null;
  state.masterWorkbookP2    = null;
}

const ui = {
  monthLabel: document.getElementById("monthLabel"),
  tabP1: document.getElementById("tabP1"),
  tabP2: document.getElementById("tabP2"),
  tabP3: document.getElementById("tabP3"),
  panelP1: document.getElementById("panelP1"),
  panelP2: document.getElementById("panelP2"),
  panelP3: document.getElementById("panelP3"),
  sourceFile: document.getElementById("sourceFile"),
  sourceSheet1a: document.getElementById("sourceSheet1a"),
  p3FileA: document.getElementById("p3FileA"),
  p3SheetA: document.getElementById("p3SheetA"),
  p3FileB: document.getElementById("p3FileB"),
  p3SheetB: document.getElementById("p3SheetB"),
  masterFile: document.getElementById("masterFile"),
  masterSheet: document.getElementById("masterSheet"),
  masterFileP2: document.getElementById("masterFileP2"),
  masterSheetP2: document.getElementById("masterSheetP2"),
  payslipFiles: document.getElementById("payslipFiles"),
  run1: document.getElementById("run1"),
  run2: document.getElementById("run2"),
  run3: document.getElementById("run3"),
  run4: document.getElementById("run4"),
  resultSummary: document.getElementById("resultSummary"),
  logOutput: document.getElementById("logOutput"),
  statusSource: null,
  statusMaster: null,
  statusP3FileA: null,
  statusP3FileB: null,
  statusMasterP2: null,
  statusPdfs: null,
  errorBox: document.getElementById("errorBox"),
};

init();

function init() {
  if (window.pdfjsLib && window.pdfjsLib.GlobalWorkerOptions) {
    window.pdfjsLib.GlobalWorkerOptions.workerSrc =
      "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
  }

  fillMonthSelect();
  setActiveTab("p1");
  bindEvents();
  fillSelect(ui.sourceSheet1a, []);
  fillSelect(ui.p3SheetA, []);
  fillSelect(ui.p3SheetB, []);
  fillSelect(ui.masterSheet, []);
  fillSelect(ui.masterSheetP2, []);
  refreshUiReadiness();
  log("Gotowe. Zrób procesy 1-2-3-4 na ekranie.");
}

function fillMonthSelect() {
  ui.monthLabel.innerHTML = "";
  MONTHS.forEach((month) => {
    const option = document.createElement("option");
    option.value = month;
    option.textContent = month;
    ui.monthLabel.appendChild(option);
  });
  const now = new Date().toLocaleString("en-US", { month: "short", year: "2-digit" }).replace(" ", "-");
  ui.monthLabel.value = MONTHS.includes(now) ? now : MONTHS[0];
}

function bindEvents() {
  ui.tabP1.addEventListener("click", () => setActiveTab("p1"));
  ui.tabP2.addEventListener("click", () => setActiveTab("p2"));
  ui.tabP3.addEventListener("click", () => setActiveTab("p3"));
  ui.sourceFile.addEventListener("change", onSourceFilePicked);
  ui.masterFile.addEventListener("change", onMasterFilePicked);
  ui.p3FileA.addEventListener("change", onP3FileAPicked);
  ui.p3FileB.addEventListener("change", onP3FileBPicked);
  ui.masterFileP2.addEventListener("change", onMasterFileP2Picked);
  ui.payslipFiles.addEventListener("change", refreshUiReadiness);
  ui.sourceSheet1a.addEventListener("change", refreshUiReadiness);
  ui.p3SheetA.addEventListener("change", refreshUiReadiness);
  ui.p3SheetB.addEventListener("change", refreshUiReadiness);
  ui.masterSheet.addEventListener("change", refreshUiReadiness);
  ui.masterSheetP2.addEventListener("change", refreshUiReadiness);
  ui.run1.addEventListener("click", runProcess1);
  ui.run2.addEventListener("click", runProcess2);
  ui.run3.addEventListener("click", runProcess3);
  ui.run4.addEventListener("click", runProcess4);
}

function setActiveTab(key) {
  const onP1 = key === "p1";
  const onP2 = key === "p2";
  const onP3 = key === "p3";
  ui.tabP1.classList.toggle("active", onP1);
  ui.tabP2.classList.toggle("active", onP2);
  ui.tabP3.classList.toggle("active", onP3);
  ui.panelP1.classList.toggle("hidden", !onP1);
  ui.panelP2.classList.toggle("hidden", !onP2);
  ui.panelP3.classList.toggle("hidden", !onP3);
  refreshUiReadiness();
}

async function onSourceFilePicked(event) {
  const file = event.target.files?.[0];
  fillSelect(ui.sourceSheet1a, []);
  state.sourceWorkbook = null;
  if (!file) { hideInlineError(); refreshUiReadiness(); return; }
  try {
    // Only read sheet names for the dropdown.
    // runProcess1 and runProcess2 always re-read the file fresh from the input.
    const wb = await readWorkbookFromFile(file);
    state.sourceWorkbook = wb;
    state.sourceWorkbookName = file.name;
    state.sourceSheets = getWorkbookSheetNames(wb);
    fillSelect(ui.sourceSheet1a, state.sourceSheets);
    hideInlineError();
    log(`Source: ${file.name} (${state.sourceSheets.length} arkusz(e): ${state.sourceSheets.join(", ")})`);
    refreshUiReadiness();
  } catch (error) {
    showError(`Nie udało się odczytać arkuszy z source file: ${error.message}`);
  }
}

async function onMasterFilePicked(event) {
  const file = event.target.files?.[0];
  fillSelect(ui.masterSheet, []);
  state.masterWorkbook = null;
  state.masterWorkbookBytes = null;
  if (!file) { hideInlineError(); refreshUiReadiness(); return; }
  try {
    // Only read sheet names for the dropdown.
    // runProcess2 always re-reads the file fresh from the input.
    const wb = await readWorkbookFromFile(file);
    state.masterWorkbook = wb;
    state.masterWorkbookName = file.name;
    state.masterSheets = getWorkbookSheetNames(wb);
    fillSelect(ui.masterSheet, state.masterSheets);
    hideInlineError();
    log(`Master: ${file.name} (${state.masterSheets.length} arkusz(e): ${state.masterSheets.join(", ")})`);
    refreshUiReadiness();
  } catch (error) {
    showError(`Nie udało się odczytać arkuszy z master file: ${error.message}`);
  }
}

async function onP3FileAPicked(event) {
  const file = event.target.files?.[0];
  fillSelect(ui.p3SheetA, []);
  state.p3WorkbookA = null;
  state.p3TableA    = null;
  if (!file) { refreshUiReadiness(); return; }
  try {
    // Read sheet names ONLY for the dropdown — nothing else stored
    const buf    = await file.arrayBuffer();
    const wb     = XLSX.read(buf, { type: "array", bookSheets: true });
    const sheets = getWorkbookSheetNames(wb);
    fillSelect(ui.p3SheetA, sheets);
    hideInlineError();
    log(`Plik A: ${file.name} (${sheets.length} arkusz(e): ${sheets.join(", ")})`);
    refreshUiReadiness();
  } catch (error) {
    showError(`Nie udało się odczytać arkuszy z pliku A: ${error.message}`);
  }
}

async function onP3FileBPicked(event) {
  const file = event.target.files?.[0];
  fillSelect(ui.p3SheetB, []);
  state.p3WorkbookB = null;
  state.p3TableB    = null;
  if (!file) { refreshUiReadiness(); return; }
  try {
    // Read sheet names ONLY for the dropdown — nothing else stored
    const buf    = await file.arrayBuffer();
    const wb     = XLSX.read(buf, { type: "array", bookSheets: true });
    const sheets = getWorkbookSheetNames(wb);
    fillSelect(ui.p3SheetB, sheets);
    hideInlineError();
    log(`Plik B: ${file.name} (${sheets.length} arkusz(e): ${sheets.join(", ")})`);
    refreshUiReadiness();
  } catch (error) {
    showError(`Nie udało się odczytać arkuszy z pliku B: ${error.message}`);
  }
}

async function onMasterFileP2Picked(event) {
  const file = event.target.files?.[0];
  fillSelect(ui.masterSheetP2, []);
  state.masterWorkbookP2 = null;
  if (!file) { refreshUiReadiness(); return; }
  try {
    // Only read sheet names for the dropdown.
    // runProcess4 always re-reads the file fresh from the input.
    const wb = await readWorkbookFromFile(file);
    const sheets = getWorkbookSheetNames(wb);
    state.masterWorkbookP2 = wb;
    fillSelect(ui.masterSheetP2, sheets);
    hideInlineError();
    log(`Master P4: ${file.name} (${sheets.length} arkusz(e): ${sheets.join(", ")})`);
    refreshUiReadiness();
  } catch (error) {
    showError(`Nie udało się odczytać arkuszy z master P4: ${error.message}`);
  }
}

function clearP3Selections(logMessage = "") {
  state.p3WorkbookA = null;
  state.p3WorkbookAName = "";
  state.p3SheetsA = [];
  fillSelect(ui.p3SheetA, []);
  if (ui.p3FileA) {
    ui.p3FileA.value = "";
  }
  state.p3WorkbookB = null;
  state.p3WorkbookBName = "";
  state.p3SheetsB = [];
  fillSelect(ui.p3SheetB, []);
  if (ui.p3FileB) {
    ui.p3FileB.value = "";
  }
  if (logMessage) {
    log(logMessage);
  }
}

function fillSelect(selectEl, values) {
  selectEl.innerHTML = "";
  const cleanValues = Array.isArray(values)
    ? values.map((v) => String(v || "").trim()).filter(Boolean)
    : [];
  if (cleanValues.length === 0) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "-- brak arkuszy --";
    selectEl.appendChild(option);
    return;
  }
  cleanValues.forEach((name) => {
    const option = document.createElement("option");
    option.value = name;
    option.textContent = name;
    selectEl.appendChild(option);
  });
}

async function readWorkbookBundle(file) {
  const buffer = await file.arrayBuffer();
  const workbook = readWorkbookFromBuffer(buffer);
  return { workbook, buffer };
}

async function readWorkbookFromFile(file) {
  const buffer = await file.arrayBuffer();
  return readWorkbookFromBuffer(buffer);
}

function readWorkbookFromBuffer(buffer) {
  let workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellStyles: true });
  let sheetNames = getWorkbookSheetNames(workbook);

  // Fallbacks for files with non-standard metadata/order.
  if (sheetNames.length === 0) {
    workbook = XLSX.read(buffer, { type: "array", cellDates: true, cellStyles: true, codepage: 65001 });
    sheetNames = getWorkbookSheetNames(workbook);
  }
  if (sheetNames.length === 0) {
    const probe = XLSX.read(buffer, { type: "array", bookSheets: true, bookProps: true });
    sheetNames = getWorkbookSheetNames(probe);
  }

  if (sheetNames.length === 0) {
    throw new Error("Plik nie zawiera żadnych arkuszy lub ma nieobsługiwany format.");
  }

  workbook.SheetNames = sheetNames;
  return workbook;
}

function getWorkbookSheetNames(workbook) {
  const fromSheetNames = Array.isArray(workbook?.SheetNames) ? workbook.SheetNames : [];
  const fromSheetObject = workbook?.Sheets ? Object.keys(workbook.Sheets) : [];
  const fromWorkbookMeta = Array.isArray(workbook?.Workbook?.Sheets)
    ? workbook.Workbook.Sheets.map((sheet) => sheet?.name)
    : [];
  return Array.from(new Set([...fromSheetNames, ...fromSheetObject, ...fromWorkbookMeta]
    .map((name) => String(name || "").trim())
    .filter(Boolean)));
}

function sheetToRows(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) {
    throw new Error(`Sheet '${sheetName}' not found.`);
  }
  return XLSX.utils.sheet_to_json(ws, { defval: null, raw: false });
}

function showError(message) {
  log(`ERROR: ${message}`);
  showInlineError(message);
  window.alert(`Ups, coś poszło nie tak.\n\n${message}`);
}

function log(message) {
  const ts = new Date().toLocaleTimeString("pl-PL");
  ui.logOutput.textContent += `[${ts}] ${message}\n`;
  ui.logOutput.scrollTop = ui.logOutput.scrollHeight;
}

function clearLog() {
  ui.logOutput.textContent = "";
}

function setResultSummary(lines) {
  ui.resultSummary.classList.remove("empty-state");
  ui.resultSummary.innerHTML = `<ul>${lines.map((line) => `<li>${escapeHtml(line)}</li>`).join("")}</ul>`;
}

function setBadge(el, ok, text) {
  if (!el) return;
  el.classList.remove("ok", "warn");
  el.classList.add(ok ? "ok" : "warn");
  el.textContent = text;
}

function refreshUiReadiness() {
  const hasSource  = !!(ui.sourceFile.files && ui.sourceFile.files[0]);
  const hasMaster  = !!(ui.masterFile.files && ui.masterFile.files[0]);
  const hasPdfs    = (ui.payslipFiles.files || []).length > 0;

  const readyProcess1 = hasSource && !!ui.sourceSheet1a.value;
  const readyProcess2 = hasSource && hasMaster && !!ui.sourceSheet1a.value && !!ui.masterSheet.value;
  const readyProcess3 = !!(ui.p3FileA.files && ui.p3FileA.files[0]) && !!(ui.p3FileB.files && ui.p3FileB.files[0]) && !!ui.p3SheetA.value && !!ui.p3SheetB.value;
  const p4MasterReady = !!(ui.masterFileP2.files && ui.masterFileP2.files[0]) && !!ui.masterSheetP2.value;
  const readyProcess4 = hasPdfs && !!p4MasterReady;
  ui.run1.disabled = !readyProcess1;
  ui.run2.disabled = !readyProcess2;
  ui.run3.disabled = !readyProcess3;
  ui.run4.disabled = !readyProcess4;
}

function showInlineError(message) {
  if (!ui.errorBox) return;
  ui.errorBox.textContent = message;
  ui.errorBox.classList.remove("hidden");
}

function hideInlineError() {
  if (!ui.errorBox) return;
  ui.errorBox.textContent = "";
  ui.errorBox.classList.add("hidden");
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function normalizeSap(value) {
  if (value === null || value === undefined) {
    return "";
  }
  const raw = String(value).trim();
  if (!raw) {
    return "";
  }
  const num = Number(raw);
  if (!Number.isNaN(num)) {
    return String(Math.trunc(num));
  }
  return raw.split(".")[0].trim();
}

function parseDate(value) {
  if (!value) {
    return null;
  }
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }
  const parsed = new Date(value);
  if (!Number.isNaN(parsed.getTime())) {
    return new Date(Date.UTC(parsed.getFullYear(), parsed.getMonth(), parsed.getDate()));
  }
  const isoMatch = String(value).trim().match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
  if (isoMatch) {
    const day = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const year = Number(isoMatch[3]);
    const utc = new Date(Date.UTC(year, month - 1, day));
    if (utc.getUTCFullYear() === year && utc.getUTCMonth() === month - 1 && utc.getUTCDate() === day) {
      return utc;
    }
  }
  return null;
}

function formatDateIso(date) {
  return [
    date.getUTCFullYear(),
    String(date.getUTCMonth() + 1).padStart(2, "0"),
    String(date.getUTCDate()).padStart(2, "0"),
  ].join("-");
}

function addWorkingDays(startDate, nDays) {
  const whole = nDays > 0 ? Math.max(1, Math.floor(Number(nDays))) : 0;
  if (whole <= 0) {
    return new Date(startDate.getTime());
  }
  let current = new Date(startDate.getTime());
  let counted = 1;
  while (counted < whole) {
    current = new Date(current.getTime() + 24 * 60 * 60 * 1000);
    if (current.getUTCDay() >= 1 && current.getUTCDay() <= 5) {
      counted += 1;
    }
  }
  return current;
}

function nextWorkday(date) {
  let current = new Date(date.getTime() + 24 * 60 * 60 * 1000);
  while (current.getUTCDay() === 0 || current.getUTCDay() === 6) {
    current = new Date(current.getTime() + 24 * 60 * 60 * 1000);
  }
  return current;
}

function validateProcess1Inputs() {
  if (!ui.sourceFile.files || !ui.sourceFile.files[0]) {
    throw new Error("Wybierz Current Holiday Report.");
  }
  if (!ui.sourceSheet1a.value) {
    throw new Error("Wybierz arkusz source dla Process 1.");
  }
}

function validateProcess2Inputs() {
  if (!ui.sourceFile.files || !ui.sourceFile.files[0]) {
    throw new Error("Wybierz Current Holiday Report.");
  }
  if (!ui.masterFile.files || !ui.masterFile.files[0]) {
    throw new Error("Wybierz Holiday Balance master file.");
  }
  if (!ui.sourceSheet1a.value || !ui.masterSheet.value) {
    throw new Error("Wybierz wymagane sheety dla Process 2.");
  }
}

function validateProcess3Inputs() {
  if (!ui.p3FileA.files || !ui.p3FileA.files[0])
    throw new Error("Wybierz plik A dla Process 3.");
  if (!ui.p3FileB.files || !ui.p3FileB.files[0])
    throw new Error("Wybierz plik B dla Process 3.");
  if (!ui.p3SheetA.value || !ui.p3SheetB.value)
    throw new Error("Wybierz arkusze dla Process 3.");
}

async function runProcess1() {
  clearLog();
  clearWorkbookState();  // always start fresh — no data from previous processes
  try {
    validateProcess1Inputs();
    log("=== Process 1 ===");
    const sourceFile = ui.sourceFile.files && ui.sourceFile.files[0];
    if (!sourceFile) throw new Error("Wybierz plik source.");
    log(`Czytam source: ${sourceFile.name}`);
    const sourceWb   = await readWorkbookFromFile(sourceFile);
    const sourceRows = sheetToRows(sourceWb, ui.sourceSheet1a.value);
    const flexiRows  = buildFlexiAbsenceRows(sourceRows);

    const flexiWb = buildWorkbookFromRows(
      "Flexi absence input",
      [
        "EMPLOYEE NUMBER",
        "Absence Code",
        "NAME",
        "Absence Type",
        "From",
        "To",
        "Days",
      ],
      flexiRows
    );
    writeWorkbookDownload(flexiWb, "Flexi-absence input.xlsx");
    log(`Flexi input generated (${flexiRows.length} row(s)).`);

    setResultSummary([
      "Process 1 zakończony.",
      `Flexi row(s): ${flexiRows.length}.`,
      "Wynik: input do Flexi został pobrany.",
      "Pobrany plik: Flexi-absence input.xlsx.",
    ]);
  } catch (error) {
    showError(error.message);
  }
}

async function runProcess2() {
  clearLog();
  clearWorkbookState();  // always start fresh — no data from previous processes
  try {
    validateProcess2Inputs();
    log("=== Process 2 ===");

    // Read source file fresh from disk
    const sourceFile = ui.sourceFile.files && ui.sourceFile.files[0];
    if (!sourceFile) throw new Error("Wybierz plik source.");
    log(`Czytam source: ${sourceFile.name}`);
    const sourceWb = await readWorkbookFromFile(sourceFile);
    const sourceRows = sheetToRows(sourceWb, ui.sourceSheet1a.value);
    const balanceSummaryRows = buildBalanceSummaryFromSource(sourceRows);
    log(`  Source: ${balanceSummaryRows.length} rekordów.`);

    // Read master file fresh from disk
    const masterFile = ui.masterFile.files && ui.masterFile.files[0];
    if (!masterFile) throw new Error("Wybierz plik master.");
    log(`Czytam master: ${masterFile.name}`);
    const masterBundle = await readWorkbookBundle(masterFile);
    const masterWb    = masterBundle.workbook;
    const masterBytes = masterBundle.buffer;

    const monthLabel  = ui.monthLabel.value;
    const sheetName   = ui.masterSheet.value;

    const updatePlan = buildMasterUpdatePlan(
      masterWb,
      sheetName,
      balanceSummaryRows,
      monthLabel
    );

    const outBytes = applyMasterUpdatesViaWorksheetXml(masterBytes, updatePlan);
    const outName  = buildUpdatedMasterName(masterFile.name, monthLabel);
    writeBytesDownload(outBytes, outName);

    clearP3Selections("Process 3: wybierz ponownie pliki A i B do porównania.");
    log(`Master updated (${updatePlan.updates.length} employee(s)).`);
    setResultSummary([
      "Process 2 zakończony.",
      `Master updated row(s): ${updatePlan.updates.length}.`,
      "Wynik: utworzono nowy plik master po zmianach bez ingerencji w formatowanie.",
      "Process 3: wymagany ponowny wybór plików A i B.",
      `Pobrany plik: ${outName}.`,
    ]);
    refreshUiReadiness();
  } catch (error) {
    showError(error.message);
  }
}

async function runProcess3() {
  clearLog();
  clearWorkbookState();  // always start fresh — no data from previous processes
  state.p3TableA = null;
  state.p3TableB = null;
  try {
    validateProcess3Inputs();
    log("=== Process 3 ===");

    // Read ONLY from the file inputs — no state, no cache
    const fileA = ui.p3FileA.files && ui.p3FileA.files[0];
    const fileB = ui.p3FileB.files && ui.p3FileB.files[0];
    if (!fileA) throw new Error("Wybierz plik A.");
    if (!fileB) throw new Error("Wybierz plik B.");

    const sheetA = ui.p3SheetA.value;
    const sheetB = ui.p3SheetB.value;

    // Read plik A fresh from disk — arrayBuffer() reads the actual file bytes
    log(`Czytam plik A: ${fileA.name} / ${sheetA}`);
    const bufA = await fileA.arrayBuffer();
    const tableA = await buildTempTable(bufA, sheetA, {
      idCandidates:    ["Holid", "SAP ID", "ID"],
      valueCandidates: ["Total holidays"],
      nameCandidates:  ["Name", "Employee", "Employee Name"],
      valueKey: "totalHolidays",
    });
    log(`  Plik A: ${tableA.size} rekordów.`);
    tableA.forEach((v, sap) => log(`  ID=${sap} | Total holidays=${v.totalHolidays}`));

    // Read plik B fresh from disk
    log(`Czytam plik B: ${fileB.name} / ${sheetB}`);
    const bufB = await fileB.arrayBuffer();
    const tableB = await buildTempTable(bufB, sheetB, {
      idCandidates:    ["SAP ID", "ID", "Holid"],
      valueCandidates: ["Total holiday balance"],
      nameCandidates:  ["Employee", "Name", "Employee Name"],
      valueKey: "totalHolidayBalance",
    });
    log(`  Plik B: ${tableB.size} rekordów.`);
    tableB.forEach((v, sap) => log(`  ID=${sap} | Total holiday balance=${v.totalHolidayBalance}`));

    // ── Compare ──
    const compareRows = [];
    tableA.forEach((recA, sap) => {
      const recB = tableB.get(sap);
      const src  = recA.totalHolidays;
      const mst  = recB ? recB.totalHolidayBalance : null;
      const name = recA.name || (recB ? recB.name : "") || sap;
      compareRows.push({
        "SAP ID":                         sap,
        "Employee Name":                  name,
        "Total holidays (source)":        src,
        "Total holiday balance (master)": mst,
        Match: src !== null && mst !== null ? round2(src) === round2(mst) : false,
      });
    });

    const outBytes3 = buildBalanceComparisonWorkbook(compareRows);
    const outName   = `Balance comparison ${todayIso()}.xlsx`;
    downloadStyledXlsx(outBytes3, outName);
    const mismatch = compareRows.filter((r) => r.Match !== true).length;
    setResultSummary([
      "Process 3 zakończony.",
      `Records: ${compareRows.length}.`,
      `Matched: ${compareRows.length - mismatch}.`,
      `Mismatched: ${mismatch}.`,
      `Pobrany plik: ${outName}.`,
    ]);
    log(`Gotowe. Niezgodności: ${mismatch}.`);

  } catch (error) {
    showError(error.message);
  } finally {
    // No state was stored — nothing to clear
    state.p3TableA = null;
    state.p3TableB = null;
  }
}

// Build a temporary lookup table from an xlsx file.
// Reads raw <v> values from XML — exactly what Excel shows, no formula evaluation.
// Returns Map<sapId, { [valueKey]: number|null, name: string }>
async function buildTempTable(arrayBuffer, sheetName, opts) {
  if (!window.fflate) throw new Error("fflate not available.");
  const zipped = window.fflate.unzipSync(new Uint8Array(arrayBuffer));
  const shared = parseSharedStrings(zipped);

  const wbXml  = window.fflate.strFromU8(zipped["xl/workbook.xml"]);
  const wbRels = window.fflate.strFromU8(zipped["xl/_rels/workbook.xml.rels"]);
  const wsPath = resolveWorksheetPath(wbXml, wbRels, sheetName);
  if (!zipped[wsPath]) throw new Error(`Arkusz '${sheetName}' nie znaleziony.`);
  const wsXml = window.fflate.strFromU8(zipped[wsPath]);

  const doc = new DOMParser().parseFromString(wsXml, "application/xml");
  const sd  = findDirectChildByLocalName(doc.documentElement, "sheetData");
  if (!sd) throw new Error("Brak sheetData.");
  const xmlRows = getDirectChildrenByLocalName(sd, "row");
  if (!xmlRows.length) throw new Error(`Arkusz '${sheetName}' jest pusty.`);

  // Read raw <v> value for a cell (ignore formula entirely)
  function rawVal(cell) {
    if (!cell) return null;
    const t   = cell.getAttribute("t") || "";
    const vEl = findDirectChildByLocalName(cell, "v");
    const v   = vEl ? vEl.textContent.trim() : "";
    if (!v) return null;
    if (t === "s") {
      const idx = parseInt(v, 10);
      return Number.isNaN(idx) ? null : (shared[idx] || null);
    }
    const n = Number(v);
    return Number.isFinite(n) ? n : v;
  }

  // Build cell map for the whole sheet: ref → raw value
  const cellMap = new Map();
  xmlRows.forEach((row) => {
    getDirectChildrenByLocalName(row, "c").forEach((cell) => {
      const ref = cell.getAttribute("r") || "";
      cellMap.set(ref, rawVal(cell));
    });
  });

  // Find header row — first row where ID candidate column contains a string
  // Scan rows 1 and 2 to build headerMap: colLetter → normalised header text
  const headerMap = {};
  const headerRaw = {};

  function scanHeaderRow(rowEl) {
    getDirectChildrenByLocalName(rowEl, "c").forEach((cell) => {
      const ref = cell.getAttribute("r") || "";
      const col = ref.replace(/[0-9]/g, "");
      const v   = rawVal(cell);
      if (typeof v === "string" && v.trim()) {
        const norm = v.replace(/\s+/g, " ").trim();
        headerMap[col] = norm.toLowerCase();
        headerRaw[col] = norm;
      }
    });
  }

  // Scan ONLY row 1 for headers — row 2 contains sub-labels (Holiday/Special)
  // that must NOT overwrite the real column headers from row 1.
  scanHeaderRow(xmlRows[0]);

  log(`  Headers found: ${Object.values(headerRaw).join(" | ")}`);

  // Find column by EXACT name match only (case-insensitive, whitespace-normalised).
  // No partial matching — prevents "Holiday" subheader columns interfering.
  function findCol(candidates) {
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/\s+/g, " ").trim();
      for (const [col, hdr] of Object.entries(headerMap)) {
        if (hdr === cl) return col;
      }
    }
    return null;
  }

  const idCol    = findCol(opts.idCandidates);
  const valueCol = findCol(opts.valueCandidates);
  const nameCol  = findCol(opts.nameCandidates || []);

  if (!idCol)
    throw new Error(`Brak kolumny ID w '${sheetName}'. Szukano: ${opts.idCandidates.join(", ")}. Nagłówki: ${Object.values(headerRaw).join(" | ")}`);
  if (!valueCol)
    throw new Error(`Brak kolumny wartości w '${sheetName}'. Szukano: ${opts.valueCandidates.join(", ")}. Nagłówki: ${Object.values(headerRaw).join(" | ")}`);

  log(`  ID col: ${idCol} (${headerRaw[idCol]}), Value col: ${valueCol} (${headerRaw[valueCol]})`);

  // Build result map
  const result = new Map();
  // Determine which row numbers are data rows (skip header rows)
  const rowNums = xmlRows.map((r) => parseInt(r.getAttribute("r") || "0", 10))
                         .filter((n) => n > 0)
                         .sort((a, b) => a - b);

  for (const rn of rowNums) {
    const idRaw = cellMap.get(`${idCol}${rn}`);
    const sap   = normalizeSap(idRaw);
    if (!sap || !/^[0-9]+$/.test(sap)) continue;  // skip header/non-numeric rows

    const val  = cellMap.get(`${valueCol}${rn}`);
    const num  = typeof val === "number" ? Math.round(val * 100) / 100 : null;
    const name = nameCol ? String(cellMap.get(`${nameCol}${rn}`) || "").trim() : "";

    const rec = { name };
    rec[opts.valueKey] = num;
    result.set(sap, rec);
  }

  return result;
}

function resolveMasterForP2() {
  const f = ui.masterFileP2.files && ui.masterFileP2.files[0];
  if (f && ui.masterSheetP2.value) return { file: f, sheet: ui.masterSheetP2.value };
  throw new Error("Wybierz plik master i arkusz dla Process 4.");
}

async function runProcess4() {
  clearLog();
  clearWorkbookState();  // always start fresh — no data from previous processes
  try {
    log("=== Process 4 ===");
    const pdfFiles = Array.from(ui.payslipFiles.files || []);
    if (pdfFiles.length === 0) throw new Error("Wybierz co najmniej jeden payslip PDF.");

    const masterFile  = ui.masterFileP2.files && ui.masterFileP2.files[0];
    const masterSheet = ui.masterSheetP2.value;
    if (!masterFile)  throw new Error("Wybierz plik master dla Process 4.");
    if (!masterSheet) throw new Error("Wybierz arkusz mastera dla Process 4.");
    log(`Reading master fresh from disk: ${masterFile.name} / ${masterSheet}`);

    const masterBuf = await masterFile.arrayBuffer();
    const holidaySummary = await readSheetColumnsDirect(masterBuf, masterSheet, {
      idCandidates:    ["SAP ID", "ID", "Holid"],
      nameCandidates:  ["Employee", "Name", "Employee Name"],
      valueCandidates: ["Total holiday balance"],
      valueLabel:      "Total holiday balance",
    });
    const specialSummary = await readSheetColumnsDirect(masterBuf, masterSheet, {
      idCandidates:    ["SAP ID", "ID", "Holid"],
      nameCandidates:  ["Employee", "Name", "Employee Name"],
      valueCandidates: ["Total special holiday balance"],
      valueLabel:      "Total special holiday balance",
    });

    const totals   = {};
    const specials = {};
    const names    = {};
    holidaySummary.forEach((r) => {
      totals[r["SAP ID"]] = r["Total holiday balance"];
      if (r["Employee Name"]) names[r["SAP ID"]] = r["Employee Name"];
    });
    specialSummary.forEach((r) => {
      specials[r["SAP ID"]] = r["Total special holiday balance"];
    });

    log(`  Master records read: ${Object.keys(totals).length}`);
    Object.entries(totals).forEach(([sap, val]) => {
      log(`  SAP ${sap}: Holiday=${val}, Special=${specials[sap] ?? "n/a"}`);
    });

    // Parse payslip PDFs
    const records = [];
    for (const file of pdfFiles) {
      log(`Parsing ${file.name} ...`);
      const { records: fileRecords, pagesTotal, pagesWithText } = await parsePayslipBatch(file);
      records.push(...fileRecords);
      log(`  pages: ${pagesTotal}, with text: ${pagesWithText}, employees: ${fileRecords.length}`);
      fileRecords.forEach((rec) => {
        log(
          `  p${rec.page ?? "?"}: ID=${rec.sap_id} | `
          + `Holiday=${rec.payslip_holidays_total ?? "null"} | `
          + `Special=${rec.payslip_special_total ?? "null"}`
        );
      });
    }

    if (records.length === 0) throw new Error("Nie udało się sparsować żadnych rekordów z payslipów.");

    const reconciliationRows = buildPayslipReportRows(records, totals, specials, names);
    const outBytes4 = buildPayslipReportWorkbook(reconciliationRows);
    const outName   = `Payslip reconciliation ${todayIso()}.xlsx`;
    downloadStyledXlsx(outBytes4, outName);
    const mismatches = reconciliationRows.filter((row) => !(row.Result && row["Special result"])).length;
    setResultSummary([
      "Process 4 zakończony.",
      `Records: ${reconciliationRows.length}.`,
      `Matched: ${reconciliationRows.length - mismatches}.`,
      `Mismatched: ${mismatches}.`,
      `Pobrany plik: ${outName}.`,
    ]);
    log(`Reconciliation completed. Mismatch(es): ${mismatches}.`);
  } catch (error) {
    showError(error.message);
  }
}

function buildFlexiAbsenceRows(rows) {
  const required = ["Project", "User", "SAP ID", "Start Date", "Days"];
  const missing = required.filter((column) => !rows.some((row) => Object.hasOwn(row, column)));
  if (missing.length > 0) {
    throw new Error(`Brak kolumn w source sheet (Process 1/2): ${missing.join(", ")}`);
  }

  const filtered = rows
    .map((row) => ({
      projectKey: String(row.Project || "").trim().toLowerCase(),
      user: String(row.User || "").trim(),
      sap: normalizeSap(row["SAP ID"]),
      startDate: parseDate(row["Start Date"]),
      days: Number(row.Days || 0),
    }))
    .filter((row) => row.sap && row.startDate && Object.hasOwn(FLEXI_PROJECT_MAP, row.projectKey))
    .sort((a, b) => {
      if (a.sap !== b.sap) return a.sap.localeCompare(b.sap);
      if (a.projectKey !== b.projectKey) return a.projectKey.localeCompare(b.projectKey);
      return a.startDate - b.startDate;
    });

  if (filtered.length === 0) {
    throw new Error("Brak pasujących rekordów Process 1/2 po filtrze project map.");
  }

  const grouped = [];
  let i = 0;
  while (i < filtered.length) {
    const base = filtered[i];
    const map = FLEXI_PROJECT_MAP[base.projectKey];
    let currentStart = base.startDate;
    let currentDays = Number(base.days || 0);
    let currentTo = addWorkingDays(currentStart, currentDays);
    let j = i + 1;

    while (j < filtered.length) {
      const row = filtered[j];
      if (row.sap !== base.sap || row.projectKey !== base.projectKey) {
        break;
      }
      const expectedNext = nextWorkday(currentTo);
      if (formatDateIso(row.startDate) === formatDateIso(expectedNext)) {
        currentDays += Number(row.days || 0);
        currentTo = addWorkingDays(currentStart, currentDays);
        j += 1;
      } else {
        break;
      }
    }

    grouped.push({
      "EMPLOYEE NUMBER": base.sap,
      "Absence Code": map.code,
      NAME: base.user,
      "Absence Type": map.type,
      From: formatDateIso(currentStart),
      To: formatDateIso(currentTo),
      Days: round2(currentDays),
    });
    i = j;
  }
  return grouped;
}

function buildBalanceSummaryFromSource(rows) {
  const resultMap = new Map();
  rows.forEach((row) => {
    const sap = normalizeSap(row["SAP ID"] ?? row.ID);
    const project = String(row.Project || "").trim().toLowerCase();
    if (!sap || !project) {
      return;
    }
    const type = BALANCE_TYPE_MAP[project] || "Holiday";
    const days = Number(row.Days || 0);
    if (!resultMap.has(sap)) {
      resultMap.set(sap, { "SAP ID": sap, Holiday: 0, Special: 0 });
    }
    const rec = resultMap.get(sap);
    rec[type] = round2(Number(rec[type] || 0) + days);
  });
  return Array.from(resultMap.values());
}

function buildMasterUpdatePlan(workbook, sheetName, summaryRows, monthLabel) {
  const ws = workbook.Sheets[sheetName];
  if (!ws || !ws["!ref"]) {
    throw new Error("Master sheet jest pusty lub nie istnieje.");
  }
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const monthCols = findMonthColumns(ws, range, monthLabel);
  const monthIdx = monthCols.holidayIdx;
  const specialIdx = monthCols.specialIdx;
  const totalHolidayIdx = findColumnIndexByHeaderContains(ws, range, "total holiday balance");
  const totalSpecialIdx = findColumnIndexByHeaderContains(ws, range, "total special holiday balance");

  const bySap = new Map();
  for (let r = 0; r <= range.e.r; r += 1) {
    const sap = normalizeSap(getWorksheetCellValue(ws, r, 0));
    if (sap) bySap.set(sap, r);
  }

  const updates = [];
  summaryRows.forEach((rec) => {
    const rowIdx = bySap.get(rec["SAP ID"]);
    if (rowIdx === undefined) {
      log(`(!) SAP ${rec["SAP ID"]} not found in master - skipped.`);
      return;
    }
    updates.push({
      sap: rec["SAP ID"],
      rowIdx,
      holiday: rec.Holiday ? round2(rec.Holiday) : null,
      special: rec.Special ? round2(rec.Special) : null,
    });
    log(`OK SAP ${rec["SAP ID"]}: Holiday=${round2(rec.Holiday)}, Special=${round2(rec.Special)}`);
  });

  return { sheetName, monthIdx, specialIdx, updates };
}

function applyMasterUpdatePlanToWorkbook(workbook, updatePlan) {
  const ws = workbook?.Sheets?.[updatePlan.sheetName];
  if (!ws) {
    return;
  }
  // Only write month columns — never touch Total holiday/special balance cells.
  updatePlan.updates.forEach((update) => {
    setWorksheetCellValue(ws, update.rowIdx, update.monthIdx ?? updatePlan.monthIdx, update.holiday);
    setWorksheetCellValue(ws, update.rowIdx, update.specialIdx ?? updatePlan.specialIdx, update.special);
  });
}

function findMonthColumns(ws, range, monthLabel) {
  const target = normalizeMonthLabel(monthLabel);
  const seenMonths = [];
  const matches = [];
  const maxHeaderRow = Math.min(range.e.r, 4);
  for (let r = 0; r <= maxHeaderRow; r += 1) {
    for (let c = 0; c <= range.e.c; c += 1) {
      const normalized = normalizeMonthCell(ws, r, c);
      if (normalized) {
        seenMonths.push(normalized);
      }
      if (normalized === target) {
        matches.push({ row: r, col: c });
      }
    }
  }

  for (const match of matches) {
    const pair = resolveMonthPairColumns(ws, range, match.row, match.col);
    if (pair) {
      return pair;
    }
  }

  throw new Error(
    `Month '${monthLabel}' not found in master header. Found: ${Array.from(new Set(seenMonths)).join(", ")}`
  );
}

function normalizeMonthCell(ws, rowIdx, colIdx) {
  const ref = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const cell = ws[ref];
  if (!cell) {
    return "";
  }
  return (
    normalizeMonthLabel(cell.w) ||
    normalizeMonthLabel(cell.v) ||
    normalizeMonthLabel(getWorksheetCellValue(ws, rowIdx, colIdx))
  );
}

function resolveMonthPairColumns(ws, range, monthRow, monthCol) {
  const pairRows = [monthRow + 1, monthRow + 2].filter((row) => row <= range.e.r);
  for (const row of pairRows) {
    const left = normalizeSubheaderLabel(getWorksheetCellValue(ws, row, monthCol));
    const right = normalizeSubheaderLabel(getWorksheetCellValue(ws, row, monthCol + 1));
    if (left === "holiday" && right === "special") {
      return { holidayIdx: monthCol, specialIdx: monthCol + 1 };
    }
    if (left === "special" && right === "holiday") {
      return { holidayIdx: monthCol + 1, specialIdx: monthCol };
    }
  }

  if (monthCol + 1 <= range.e.c) {
    return { holidayIdx: monthCol, specialIdx: monthCol + 1 };
  }
  if (monthCol - 1 >= 0) {
    return { holidayIdx: monthCol, specialIdx: monthCol - 1 };
  }
  return null;
}

function findColumnIndexByHeaderContains(ws, range, fragment) {
  const needle = String(fragment || "").trim().toLowerCase();
  if (!needle) {
    return null;
  }
  const maxHeaderRow = Math.min(range.e.r, 4);
  for (let r = 0; r <= maxHeaderRow; r += 1) {
    for (let c = 0; c <= range.e.c; c += 1) {
      const value = String(getWorksheetCellValue(ws, r, c) || "").trim().toLowerCase();
      if (value.includes(needle)) {
        return c;
      }
    }
  }
  return null;
}

function normalizeSubheaderLabel(value) {
  const text = String(value ?? "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
  if (!text) {
    return "";
  }
  if (text.includes("special")) {
    return "special";
  }
  if (text.includes("holiday")) {
    return "holiday";
  }
  return "";
}

function getWorksheetCellValue(ws, rowIdx, colIdx) {
  const ref = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const cell = ws[ref];
  if (!cell) return null;
  return cell.w ?? cell.v ?? null;
}

function setWorksheetCellValue(ws, rowIdx, colIdx, value) {
  const ref = XLSX.utils.encode_cell({ r: rowIdx, c: colIdx });
  const prev = ws[ref] || {};
  if (value === null || value === undefined || value === "") {
    if (prev && Object.keys(prev).length > 0) {
      delete prev.v;
      if (prev.f) {
        delete prev.f;
      }
      ws[ref] = prev;
    } else {
      ws[ref] = { t: "z" };
    }
    return;
  }
  const cell = { ...prev, t: "n", v: Number(value) };
  if (cell.f) {
    delete cell.f;
  }
  ws[ref] = cell;
}

function normalizeMonthLabel(raw) {
  if (raw === null || raw === undefined) {
    return "";
  }
  const text = String(raw).trim();
  if (!text) {
    return "";
  }
  const shortMatch = text.match(/^([A-Za-z]{3})[-\s]?(\d{2}|\d{4})$/);
  if (shortMatch) {
    const month = shortMatch[1].slice(0, 1).toUpperCase() + shortMatch[1].slice(1, 3).toLowerCase();
    const year = shortMatch[2].length === 2 ? shortMatch[2] : shortMatch[2].slice(-2);
    return `${month}-${year}`.toLowerCase();
  }
  const date = parseMonthDate(raw);
  if (!date) {
    return "";
  }
  const month = date.toLocaleString("en-US", { month: "short" });
  const year = String(date.getFullYear()).slice(-2);
  return `${month}-${year}`.toLowerCase();
}

function parseMonthDate(raw) {
  if (raw instanceof Date && !Number.isNaN(raw.getTime())) {
    return raw;
  }
  if (typeof raw === "number" && Number.isFinite(raw) && XLSX?.SSF?.parse_date_code) {
    const dc = XLSX.SSF.parse_date_code(raw);
    if (dc && dc.y && dc.m) {
      return new Date(dc.y, dc.m - 1, dc.d || 1);
    }
  }
  const text = String(raw).trim();
  const slash = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2}|\d{4})$/);
  if (slash) {
    const a = Number(slash[1]);
    const b = Number(slash[2]);
    const year = slash[3].length === 2 ? 2000 + Number(slash[3]) : Number(slash[3]);
    const dayFirst = new Date(year, b - 1, a);
    const monthFirst = new Date(year, a - 1, b);
    if (dayFirst.getDate() === 1) return dayFirst;
    if (monthFirst.getDate() === 1) return monthFirst;
    if (a > 12) return dayFirst;
    if (b > 12) return monthFirst;
    return dayFirst;
  }
  const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) {
    return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
  }
  const fallback = new Date(text);
  if (!Number.isNaN(fallback.getTime())) {
    return fallback;
  }
  return null;
}

function loadSourceBalanceSummary(rows) {
  if (rows.length === 0) {
    throw new Error("Holiday Balance tab is empty.");
  }
  const cols = Object.keys(rows[0]);
  const idCol = findColumn(cols, ["SAP ID", "Holid", "ID"]);
  const nameCol = findColumn(cols, ["Name", "Employee", "Employee Name"]);
  const totalCol = findColumn(cols, ["Total holidays"]);
  if (!idCol || !totalCol) {
    throw new Error(`File A must contain columns 'SAP ID' (or 'Holid') and 'Total holidays'. Found: ${cols.join(", ")}`);
  }
  return rows
    .filter((row) => row[idCol] !== null && row[idCol] !== "")
    .map((row) => ({
      "SAP ID": normalizeSap(row[idCol]),
      "Employee Name": nameCol ? String(row[nameCol] || "").trim() : "",
      "Total holidays": toNumberOrNull(row[totalCol]),
    }))
    .filter((row) => row["SAP ID"]);
}

function findColumn(columns, candidates) {
  // Normalise all column names (collapse whitespace/newlines)
  const normalised = columns.map((c) => ({
    original: c,
    norm: String(c || "").replace(/[\s\n\r]+/g, " ").trim().toLowerCase(),
  }));

  // 1. Exact match
  for (const candidate of candidates) {
    const cl = candidate.toLowerCase().replace(/[\s\n\r]+/g, " ").trim();
    const found = normalised.find((n) => n.norm === cl);
    if (found) return found.original;
  }
  // 2. Column header starts with candidate (e.g. "total holidays" → "Total holidays")
  for (const candidate of candidates) {
    const cl = candidate.toLowerCase().replace(/[\s\n\r]+/g, " ").trim();
    if (cl.length < 6) continue;
    const found = normalised.find((n) => n.norm.startsWith(cl));
    if (found) return found.original;
  }
  // 3. Candidate starts with column header (short header contained in longer candidate)
  for (const candidate of candidates) {
    const cl = candidate.toLowerCase().replace(/[\s\n\r]+/g, " ").trim();
    const found = normalised.find((n) => n.norm.length >= 4 && cl.startsWith(n.norm));
    if (found) return found.original;
  }
  return null;
}

function readCollectiveTotals(workbook, sheetName) {
  const rows = sheetToRows(workbook, sheetName);
  if (rows.length === 0) {
    throw new Error("Master sheet is empty.");
  }
  const header = Object.keys(rows[0]);
  const idCol = findColumn(header, ["SAP ID", "ID", "Holid"]);
  const holidayCol = findColumn(header, ["Total holiday balance"]);
  const specialCol = findColumn(header, ["Total special holiday balance"]);
  const nameCol = findColumn(header, ["Employee", "Name", "Employee Name"]);
  if (!idCol) {
    throw new Error(`File B must contain ID column ('SAP ID', 'ID' or 'Holid'). Found: ${header.join(", ")}`);
  }
  if (!holidayCol) {
    throw new Error(`File B must contain column 'Total holiday balance'. Found: ${header.join(", ")}`);
  }

  const totals = {};
  const specials = {};
  const names = {};
  const duplicateIds = new Set();
  rows.forEach((row) => {
    const sap = normalizeSap(row[idCol]);
    if (!sap) return;
    if (Object.hasOwn(totals, sap)) {
      duplicateIds.add(sap);
      return;
    }
    totals[sap] = holidayCol ? toNumberOrNull(row[holidayCol]) : null;
    specials[sap] = specialCol ? toNumberOrNull(row[specialCol]) : null;
    if (nameCol && row[nameCol]) {
      names[sap] = String(row[nameCol]).trim();
    }
  });
  if (duplicateIds.size > 0) {
    throw new Error(`File B contains duplicated ID(s): ${Array.from(duplicateIds).slice(0, 10).join(", ")}.`);
  }
  return { totals, specials, names };
}

function findColumnContaining(columns, fragment) {
  const f = fragment.toLowerCase();
  return columns.find((col) => String(col).toLowerCase().includes(f)) || null;
}

// ─────────────────────────────────────────────────────────────────────────────
// buildCellTable  +  readSheetColumnsDirect
//
// Strategy: read the entire sheet XML into a flat table of raw values.
// For every cell with a formula, RE-EVALUATE the formula using the raw values
// of all other cells in the SAME sheet.  Cross-sheet references use the
// cached <v> of the cell that contains them (Excel wrote that value last time
// the file was saved — it is stale for formula cells, but correct for
// cross-sheet link cells like D3 = 'PrevYear'!AH3 whose source never changes).
//
// Result: a Map  ref → number|string|null  with fully recalculated values.
// ─────────────────────────────────────────────────────────────────────────────

function parseSharedStrings(zipped) {
  const ssBytes = zipped["xl/sharedStrings.xml"];
  if (!ssBytes) return [];
  const xml = window.fflate.strFromU8(ssBytes);
  const doc  = new DOMParser().parseFromString(xml, "application/xml");
  const sis  = doc.getElementsByTagName("si");
  const arr  = [];
  for (let i = 0; i < sis.length; i++) {
    const ts = sis[i].getElementsByTagName("t");
    let s = "";
    for (let j = 0; j < ts.length; j++) s += ts[j].textContent || "";
    arr.push(s);
  }
  return arr;
}

function getCellTextValue(cellNode, sharedStrings) {
  if (!cellNode || !cellNode.getAttribute) return "";
  const t = cellNode.getAttribute("t") || "";
  const vNode = findDirectChildByLocalName(cellNode, "v");
  const v = vNode ? vNode.textContent : "";
  if (t === "s") {
    const idx = parseInt(v, 10);
    return Number.isNaN(idx) ? "" : (sharedStrings[idx] || "");
  }
  if (t === "inlineStr") {
    const is = findDirectChildByLocalName(cellNode, "is");
    if (is) {
      const tNodes = is.getElementsByTagName("t");
      let s = "";
      for (let i = 0; i < tNodes.length; i++) s += tNodes[i].textContent || "";
      return s;
    }
  }
  return v || "";
}

// Build cell table for ONE sheet XML.
// Returns Map<ref, {raw, formula}>
function buildSheetCellTable(wsXml, sharedStrings) {
  const doc = new DOMParser().parseFromString(wsXml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0)
    throw new Error("Cannot parse sheet XML.");
  const sd = findDirectChildByLocalName(doc.documentElement, "sheetData");
  if (!sd) return new Map();

  const table = new Map();
  getDirectChildrenByLocalName(sd, "row").forEach((row) => {
    getDirectChildrenByLocalName(row, "c").forEach((cell) => {
      const ref  = cell.getAttribute("r") || "";
      const type = cell.getAttribute("t") || "";
      const vEl  = findDirectChildByLocalName(cell, "v");
      const fEl  = findDirectChildByLocalName(cell, "f");
      const vTxt = vEl ? vEl.textContent.trim() : "";
      const fTxt = fEl ? fEl.textContent.trim() : "";

      let raw = null;
      if (type === "s") {
        const idx = parseInt(vTxt, 10);
        raw = Number.isNaN(idx) ? null : (sharedStrings[idx] || null);
      } else if (vTxt !== "") {
        const n = Number(vTxt);
        raw = Number.isFinite(n) ? n : vTxt;
      }
      table.set(ref, { raw, formula: fTxt || null });
    });
  });
  return table;
}

// Build a multi-sheet lookup: sheetName → Map<ref, {raw,formula}>
function buildAllSheetTables(zipped, sharedStrings) {
  const wbXml  = window.fflate.strFromU8(zipped["xl/workbook.xml"]);
  const wbRels = window.fflate.strFromU8(zipped["xl/_rels/workbook.xml.rels"]);

  const wbDoc  = new DOMParser().parseFromString(wbXml,  "application/xml");
  const relDoc = new DOMParser().parseFromString(wbRels, "application/xml");

  // Build rId → path map
  const ridToPath = {};
  relDoc.getElementsByTagName("Relationship").forEach
    ? Array.from(relDoc.getElementsByTagName("Relationship")).forEach((r) => {
        ridToPath[r.getAttribute("Id")] = r.getAttribute("Target");
      })
    : (() => {
        const rels = relDoc.getElementsByTagName("Relationship");
        for (let i = 0; i < rels.length; i++) {
          ridToPath[rels[i].getAttribute("Id")] = rels[i].getAttribute("Target");
        }
      })();

  const tables = {};  // sheetName → Map
  const sheetEls = wbDoc.getElementsByTagName("sheet");
  for (let i = 0; i < sheetEls.length; i++) {
    const el   = sheetEls[i];
    const name = el.getAttribute("name") || "";
    const rid  = el.getAttribute("r:id") ||
                 el.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id") || "";
    let target = ridToPath[rid] || "";
    if (!target) continue;
    target = target.replace(/^\//, "");
    if (!target.startsWith("xl/")) target = "xl/" + target.replace(/^\.?\//, "");
    const wsBytes = zipped[target];
    if (!wsBytes) continue;
    const wsXml = window.fflate.strFromU8(wsBytes);
    tables[name] = buildSheetCellTable(wsXml, sharedStrings);
  }
  return tables;
}


// Evaluate all formulas for the TARGET sheet using all sheet tables for cross-sheet lookup.
function evaluateSheetFormulas(targetTable, allTables) {

  function getVal(sheetName, ref) {
    const tbl = sheetName ? (allTables[sheetName] || targetTable) : targetTable;
    const cell = tbl ? tbl.get(ref) : null;
    if (!cell) return 0;
    return typeof cell.raw === "number" ? cell.raw : 0;
  }

  function evalFormula(formula, selfRef) {
    let f = formula.startsWith("=") ? formula.slice(1) : formula;
    f = f.replace(/'([^']+)'!([A-Z]{1,3}[0-9]+)/g, (_, sheet, ref) => String(getVal(sheet, ref)));
    f = f.replace(/([A-Za-z][A-Za-z0-9 ]*)!([A-Z]{1,3}[0-9]+)/g, (_, sheet, ref) => String(getVal(sheet.trim(), ref)));
    f = f.replace(/([A-Z]{1,3}[0-9]+)/g, (match) => {
      if (match === selfRef) return "0";
      const cell = targetTable.get(match);
      if (!cell) return "0";
      return typeof cell.raw === "number" ? String(cell.raw) : "0";
    });
    if (!/^[\d\s+\-*/.()eE]+$/i.test(f)) return null;
    try {
      const result = Function('"use strict"; return (' + f + ');')();
      return Number.isFinite(result) ? Math.round(result * 100) / 100 : null;
    } catch { return null; }
  }

  for (let pass = 0; pass < 15; pass++) {
    let changed = false;
    targetTable.forEach((cell, ref) => {
      if (!cell.formula) return;
      const result = evalFormula(cell.formula, ref);
      if (result !== null && cell.raw !== result) {
        cell.raw = result;
        changed = true;
      }
    });
    if (!changed) break;
  }
}

async function readSheetColumnsDirect(arrayBuffer, sheetName, opts) {
  if (!window.fflate) throw new Error("fflate not available.");

  const bytes  = new Uint8Array(arrayBuffer);
  const zipped = window.fflate.unzipSync(bytes);

  const sharedStrings = parseSharedStrings(zipped);

  // ── Build cell tables for ALL sheets in the workbook ──
  // This lets us resolve cross-sheet formula references (e.g. 'Sep2024'!AH3)
  const allTables = buildAllSheetTables(zipped, sharedStrings);
  const targetTable = allTables[sheetName];
  if (!targetTable) throw new Error(`Sheet '${sheetName}' not found in workbook.`);

  // ── Evaluate formulas using cross-sheet data ──
  evaluateSheetFormulas(targetTable, allTables);

  // Flatten to a simple Map<ref, value> for the rest of the function
  const values = new Map();
  targetTable.forEach((cell, ref) => values.set(ref, cell.raw));

  // ── Find headers directly from values map ──
  // Scan refs to find which row numbers exist, sorted ascending
  const rowNums = Array.from(new Set(
    Array.from(values.keys()).map((ref) => parseInt(ref.replace(/[A-Z]+/g, ""), 10))
  )).filter((n) => !Number.isNaN(n)).sort((a, b) => a - b);

  const headerMap = {};
  const headerRaw = {};

  // Read a cell value as text from the values map
  function cellText(ref) {
    const v = values.get(ref);
    return v !== null && v !== undefined ? String(v).trim() : "";
  }

  // Check if a row looks like a header row (first non-empty cell is text, not a pure number)
  function isHeaderRow(rNum) {
    const colKeys = Array.from(values.keys())
      .filter((ref) => parseInt(ref.replace(/[A-Z]+/g, ""), 10) === rNum)
      .map((ref) => ref.replace(/[0-9]/g, ""));
    for (const col of colKeys) {
      const v = values.get(`${col}${rNum}`);
      if (v === null || v === undefined) continue;
      if (typeof v === "string" && v.trim()) return true;
      if (typeof v === "number") return false;
    }
    return false;
  }

  function addHeaderRowFromValues(rNum) {
    Array.from(values.keys())
      .filter((ref) => parseInt(ref.replace(/[A-Z]+/g, ""), 10) === rNum)
      .forEach((ref) => {
        const col = ref.replace(/[0-9]/g, "");
        const val = values.get(ref);
        if (typeof val !== "string" || !val.trim()) return;
        const norm = val.replace(/\s+/g, " ").trim();
        if (norm) {
          headerMap[col] = norm.toLowerCase();
          headerRaw[col] = norm;
        }
      });
  }

  if (rowNums[0]) addHeaderRowFromValues(rowNums[0]);
  if (rowNums[1] && isHeaderRow(rowNums[1])) addHeaderRowFromValues(rowNums[1]);

  // ── Column finder ──
  function findCol(candidates) {
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/\s+/g, " ").trim();
      for (const [col, hdr] of Object.entries(headerMap)) {
        if (hdr === cl) return col;
      }
    }
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/\s+/g, " ").trim();
      if (cl.length < 6) continue;
      for (const [col, hdr] of Object.entries(headerMap)) {
        if (hdr.startsWith(cl)) return col;
      }
    }
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/\s+/g, " ").trim();
      for (const [col, hdr] of Object.entries(headerMap)) {
        if (hdr.length >= 4 && cl.startsWith(hdr)) return col;
      }
    }
    return null;
  }

  const idCol    = findCol(opts.idCandidates);
  const nameCol  = findCol(opts.nameCandidates || []);
  const valueCol = findCol(opts.valueCandidates);

  if (!idCol)
    throw new Error(
      `ID column not found in '${sheetName}'. ` +
      `Looking for: ${opts.idCandidates.join(", ")}. ` +
      `Headers found: ${Object.values(headerRaw).join(" | ")}`
    );
  if (!valueCol)
    throw new Error(
      `Value column not found in '${sheetName}'. ` +
      `Looking for: ${opts.valueCandidates.join(", ")}. ` +
      `Headers found: ${Object.values(headerRaw).join(" | ")}`
    );

  // ── Read data rows from values map ──
  const result = [];
  rowNums.forEach((rNum) => {
    const idRef = `${idCol}${rNum}`;
    const idRaw = values.get(idRef);
    const sap   = normalizeSap(idRaw);
    if (!sap || !/^[0-9]+$/.test(sap)) return;

    const valRef = `${valueCol}${rNum}`;
    const val    = values.get(valRef);
    const num    = typeof val === "number" ? val : null;

    let name = "";
    if (nameCol) {
      const nameRef = `${nameCol}${rNum}`;
      const nameVal = values.get(nameRef);
      name = typeof nameVal === "string" ? nameVal : "";
    }

    const rec = { "SAP ID": sap, "Employee Name": name };
    rec[opts.valueLabel] = num;
    result.push(rec);
  });

  return result;
}



function buildBalanceComparisonRows(srcSummary, totals, names) {
  return srcSummary.map((rec) => {
    const sap = rec["SAP ID"];
    const src = toNumberOrNull(rec["Total holidays"]);
    const mst = toNumberOrNull(totals[sap]);
    return {
      "SAP ID": sap,
      "Employee Name": rec["Employee Name"] || names[sap] || sap,
      "Total holidays (source)": src,
      "Total holiday balance (master)": mst,
      Match: src !== null && mst !== null ? round2(src) === round2(mst) : false,
    };
  });
}

function buildWorkbookFromRows(sheetName, headers, rows) {
  const wsRows = [headers, ...rows.map((row) => headers.map((h) => row[h] ?? null))];
  const ws = XLSX.utils.aoa_to_sheet(wsRows);
  applyHeaderStyles(ws, headers.length);
  autoWidth(ws, wsRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  return wb;
}


// ─────────────────────────────────────────────────────────────────────────────
// Build styled xlsx from scratch using raw XML + fflate zip.
// This bypasses xlsx.js style limitations and supports red fill properly.
// ─────────────────────────────────────────────────────────────────────────────

function buildStyledXlsx(sheetName, headers, rows, isRedRow) {
  // Escape XML special chars
  function esc(v) {
    return String(v ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // Shared strings table
  const strings = [];
  const strIdx  = new Map();
  function si(val) {
    const s = String(val ?? "");
    if (!strIdx.has(s)) { strIdx.set(s, strings.length); strings.push(s); }
    return strIdx.get(s);
  }

  // Pre-register all strings
  headers.forEach((h) => si(h));
  rows.forEach((row) => {
    headers.forEach((h) => {
      const v = row[h];
      if (v === null || v === undefined) return;
      if (typeof v !== "number") si(String(v));
    });
  });

  // Column letter helper
  function colLetter(n) {
    let s = "";
    for (let c = n; c >= 0; c = Math.floor(c / 26) - 1)
      s = String.fromCharCode(65 + (c % 26)) + s;
    return s;
  }

  // Style indices:
  // 0 = default
  // 1 = header (blue bg, white bold)
  // 2 = red fill
  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><sz val="11"/><name val="Calibri"/><b/><color rgb="FFFFFFFF"/></font>
    <font><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFill="1"/>
  </cellXfs>
</styleSheet>`;

  // Build sheet XML
  let sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>`;

  // Header row (style 1)
  sheetXml += `<row r="1">`;
  headers.forEach((h, ci) => {
    const ref = `${colLetter(ci)}1`;
    sheetXml += `<c r="${ref}" t="s" s="1"><v>${si(h)}</v></c>`;
  });
  sheetXml += `</row>`;

  // Data rows
  rows.forEach((row, ri) => {
    const rn  = ri + 2;
    const red = isRedRow(row);
    const s   = red ? ` s="2"` : "";
    sheetXml += `<row r="${rn}">`;
    headers.forEach((h, ci) => {
      const ref = `${colLetter(ci)}${rn}`;
      const v   = row[h];
      if (v === null || v === undefined) {
        sheetXml += `<c r="${ref}"${s}/>`;
      } else if (typeof v === "number") {
        sheetXml += `<c r="${ref}"${s}><v>${v}</v></c>`;
      } else if (typeof v === "boolean") {
        sheetXml += `<c r="${ref}" t="b"${s}><v>${v ? 1 : 0}</v></c>`;
      } else {
        sheetXml += `<c r="${ref}" t="s"${s}><v>${si(String(v))}</v></c>`;
      }
    });
    sheetXml += `</row>`;
  });

  sheetXml += `</sheetData></worksheet>`;

  // Shared strings XML
  const ssXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${strings.length}" uniqueCount="${strings.length}">
${strings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join("")}
</sst>`;

  const wbXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="${esc(sheetName)}" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

  const appRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

  const enc = window.fflate.strToU8;
  const zipped = {
    "[Content_Types].xml":          enc(contentTypes),
    "_rels/.rels":                   enc(appRels),
    "xl/workbook.xml":               enc(wbXml),
    "xl/_rels/workbook.xml.rels":    enc(wbRels),
    "xl/worksheets/sheet1.xml":      enc(sheetXml),
    "xl/sharedStrings.xml":          enc(ssXml),
    "xl/styles.xml":                 enc(stylesXml),
  };

  return window.fflate.zipSync(zipped, { level: 6 });
}

function downloadStyledXlsx(bytes, filename) {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement("a");
  a.href     = url;
  a.download = filename.endsWith(".xlsx") ? filename : filename + ".xlsx";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function buildBalanceComparisonWorkbook(rows) {
  const headers = [
    "SAP ID",
    "Employee Name",
    "Total holidays (source)",
    "Total holiday balance (master)",
    "Match",
  ];
  return buildStyledXlsx(
    "Balance comparison",
    headers,
    rows,
    (row) => row.Match !== true
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PDF parsing — Danish payslip format
//
// Expected page layout (from working Python version):
//   ...
//   <Employee Name>
//   Employeeno  12345
//   ...
//   Holidays total   <prev_balance>   <new_balance>
//   Special holidays <prev_balance>   <new_balance>   (optional, sometimes single value)
//
// We always take the LAST (rightmost / new balance) number from each row.
// Numbers use Danish locale: thousands sep = "." and decimal sep = ","
// e.g. "1.234,56" = 1234.56
// ─────────────────────────────────────────────────────────────────────────────

async function extractPageText(page) {
  const textContent = await page.getTextContent();
  const items = (textContent.items || [])
    .map((item) => ({
      str: String(item?.str || "").trim(),
      x: Number(item?.transform?.[4] || 0),
      y: Number(item?.transform?.[5] || 0),
    }))
    .filter((item) => item.str);
  if (items.length === 0) return "";

  // Group into lines by y-coordinate (tolerance 2pt)
  items.sort((a, b) => {
    const dy = b.y - a.y;
    return Math.abs(dy) > 2 ? dy : a.x - b.x;
  });
  const lines = [];
  let cur = [], curY = null;
  items.forEach((item) => {
    if (curY === null || Math.abs(item.y - curY) <= 2) {
      cur.push(item);
      if (curY === null) curY = item.y;
    } else {
      lines.push(cur);
      cur = [item];
      curY = item.y;
    }
  });
  if (cur.length > 0) lines.push(cur);

  return lines
    .map((l) => l.sort((a, b) => a.x - b.x).map((i) => i.str).join(" ").replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join("\n");
}

function splitTextLines(text) {
  return String(text || "").split(/\r?\n/).map((l) => l.replace(/\s+/g, " ").trim()).filter(Boolean);
}

// Parse a Danish-locale number string → float | null
// "1.234,56" → 1234.56 | "14,56" → 14.56 | "14.56" → 14.56 | "-5,00" → -5
function parseDanishNumber(raw) {
  const s = String(raw || "").replace(/\u00a0/g, "").trim();
  if (!s || !/[0-9]/.test(s)) return null;
  const clean = s.replace(/[^0-9,.\-]/g, "");
  const hasDot = clean.includes(".");
  const hasComma = clean.includes(",");
  let numeric;
  if (hasDot && hasComma) {
    // Danish: "1.234,56" — dot=thousands, comma=decimal
    if (clean.lastIndexOf(",") > clean.lastIndexOf(".")) {
      numeric = Number(clean.replaceAll(".", "").replaceAll(",", "."));
    } else {
      // US style fallback: "1,234.56"
      numeric = Number(clean.replaceAll(",", ""));
    }
  } else if (hasComma) {
    numeric = Number(clean.replaceAll(",", "."));
  } else {
    numeric = Number(clean);
  }
  return Number.isFinite(numeric) ? round2(numeric) : null;
}

// Extract all numbers from a text line, returning array of floats
function extractLineNumbers(line) {
  // Split on whitespace first so "2,08 9,56" becomes ["2,08", "9,56"] not one merged token.
  return line.split(/\s+/).map((t) => parseDanishNumber(t)).filter((n) => n !== null);
}

async function parsePayslipBatch(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const pdf = await window.pdfjsLib.getDocument({ data: bytes }).promise;
  const pagesTotal = pdf.numPages || 0;
  if (pagesTotal <= 0) throw new Error(`Payslip PDF has no pages: ${file.name}`);

  const records = [];
  let pagesWithText = 0;

  for (let i = 1; i <= pagesTotal; i += 1) {
    const page = await pdf.getPage(i);
    const text = await extractPageText(page);
    if (!text) continue;
    pagesWithText += 1;

    // Only process pages that look like an employee payslip
    if (!/employeeno/i.test(text)) continue;

    const rec = parseEmployeeBlock(text);
    rec.page = i;
    if (rec.sap_id === "UNKNOWN") {
      log(`  p${i}: could not read SAP ID — skipped`);
      continue;
    }
    records.push(rec);
  }

  if (pagesWithText === 0) throw new Error(`PDF contains no readable text: ${file.name}`);
  if (records.length === 0) throw new Error(`No employee records found in: ${file.name}`);
  return { records, pagesTotal, pagesWithText };
}

function parseEmployeeBlock(text) {
  const lines = splitTextLines(text);

  // ── SAP ID ──
  // Matches: "Employeeno 12345" or "Employeeno: 12345"
  let sap_id = "UNKNOWN";
  for (const line of lines) {
    const m = line.match(/employeeno[:\s]+([0-9]{4,})/i);
    if (m) { sap_id = normalizeSap(m[1]); break; }
  }
  // Fallback: look for "Employeeno" then number on same or next line
  if (sap_id === "UNKNOWN") {
    for (let i = 0; i < lines.length; i += 1) {
      if (!/employeeno/i.test(lines[i])) continue;
      const nums = lines[i].match(/([0-9]{4,})/);
      if (nums) { sap_id = normalizeSap(nums[1]); break; }
      const next = lines[i + 1] || "";
      const numsNext = next.match(/([0-9]{4,})/);
      if (numsNext) { sap_id = normalizeSap(numsNext[1]); break; }
    }
  }

  // ── Name ── line just before "Employeeno"
  let name = "Unknown";
  for (let i = 0; i < lines.length; i += 1) {
    if (!/employeeno/i.test(lines[i])) continue;
    const prev = lines[i - 1] || "";
    if (prev && !/[0-9]/.test(prev) && prev.trim().split(/\s+/).length >= 2) {
      name = prev.trim();
    }
    break;
  }

  // ── Holidays total → new balance (last number on the row) ──
  // Row looks like: "Holidays total  14,56  28,56"  (prev + new)
  // or single:      "Holidays total  28,56"
  let payslip_holidays_total = null;
  for (const line of lines) {
    if (!/holidays\s+total/i.test(line)) continue;
    const nums = extractLineNumbers(line);
    if (nums.length >= 1) {
      payslip_holidays_total = nums[nums.length - 1]; // always take new (rightmost)
    }
    break;
  }

  // ── Special holidays → new balance ──
  // Row: "Special holidays  5,00  10,00"  or just "Special holidays  10,00"
  let payslip_special_total = null;
  for (const line of lines) {
    if (!/special\s+holidays?/i.test(line)) continue;
    const nums = extractLineNumbers(line);
    if (nums.length >= 1) {
      payslip_special_total = nums[nums.length - 1];
    }
    break;
  }

  return { sap_id, name, payslip_holidays_total, payslip_special_total };
}

function buildPayslipReportRows(records, totals, specials, names) {
  return records.map((rec) => {
    const sap = rec.sap_id;
    const eh = totals[sap] ?? null;
    const es = specials[sap] ?? null;
    const ph = rec.payslip_holidays_total;
    const ps = rec.payslip_special_total;
    return {
      "SAP ID / Employee No": sap,
      "Employee Name": names[sap] || rec.name || "Unknown",
      "Excel report value": eh,
      "Payslip value": ph,
      Result: eh !== null && ph !== null ? round2(eh) === round2(ph) : false,
      "Excel special holiday balance": es,
      "Payslip special holiday value": ps,
      "Special result": es !== null && ps !== null ? round2(es) === round2(ps) : true,
    };
  });
}

function buildPayslipReportWorkbook(rows) {
  const headers = [
    "SAP ID / Employee No",
    "Employee Name",
    "Excel report value",
    "Payslip value",
    "Result",
    "Excel special holiday balance",
    "Payslip special holiday value",
    "Special result",
  ];
  return buildStyledXlsx(
    "Reconciliation",
    headers,
    rows,
    (row) => row.Result !== true || row["Special result"] !== true
  );
}

function markRed(ws, rowIndex, colIndex) {
  const ref = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
  if (!ws[ref]) ws[ref] = { t: "s", v: "" };
  ws[ref].s = { fill: RED_FILL };
}

function applyHeaderStyles(ws, count) {
  for (let c = 0; c < count; c += 1) {
    const ref = XLSX.utils.encode_cell({ r: 0, c });
    if (!ws[ref]) ws[ref] = { t: "s", v: "" };
    ws[ref].s = {
      fill: HEADER_FILL,
      font: HEADER_FONT,
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
    };
  }
}

function autoWidth(ws, rows) {
  const widthMap = [];
  rows.forEach((row) => {
    row.forEach((value, idx) => {
      const len = String(value ?? "").length;
      widthMap[idx] = Math.min(Math.max(widthMap[idx] || 8, len + 2), 50);
    });
  });
  ws["!cols"] = widthMap.map((wch) => ({ wch }));
}

function writeWorkbookDownload(workbook, filename) {
  XLSX.writeFile(workbook, ensureXlsxName(filename), { cellStyles: true });
}

function writeBytesDownload(bytes, filename) {
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const safeName = ensureXlsxName(filename);
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = safeName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  // Revoke asynchronously to avoid race with browser download handling.
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function buildUpdatedMasterName(baseName, monthLabel) {
  const original = String(baseName || "master.xlsx");
  const cleanMonth = String(monthLabel || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
  const suffix = cleanMonth ? `-${cleanMonth}` : "";
  if (original.toLowerCase().endsWith(".xlsx")) {
    return `${original.slice(0, -5)}-updated${suffix}.xlsx`;
  }
  return `${original}-updated${suffix}.xlsx`;
}

function applyMasterUpdatesViaWorksheetXml(workbookBytes, updatePlan) {
  if (!workbookBytes) {
    throw new Error("Brak oryginalnych danych pliku master do zapisu.");
  }
  if (!window.fflate) {
    throw new Error("Brak biblioteki do bezpiecznego zapisu XLSX (fflate).");
  }
  const bytes = new Uint8Array(workbookBytes);
  const zipped = window.fflate.unzipSync(bytes);
  const workbookXmlPath = "xl/workbook.xml";
  const workbookRelsXmlPath = "xl/_rels/workbook.xml.rels";
  const workbookXml = getZipXml(zipped, workbookXmlPath);
  const updatedWorkbookXml = ensureWorkbookFullRecalcOnLoad(workbookXml);
  const workbookRelsXml = getZipXml(zipped, workbookRelsXmlPath);
  const sheetFilePath = resolveWorksheetPath(workbookXml, workbookRelsXml, updatePlan.sheetName);
  const sheetXml = getZipXml(zipped, sheetFilePath);
  const updatedSheetXml = updateWorksheetXmlValues(sheetXml, updatePlan);
  zipped[workbookXmlPath] = window.fflate.strToU8(updatedWorkbookXml);
  zipped[sheetFilePath] = window.fflate.strToU8(updatedSheetXml);
  return window.fflate.zipSync(zipped, { level: 0 });
}

function getZipXml(zipped, path) {
  const u8 = zipped[path];
  if (!u8) {
    throw new Error(`Brak pliku '${path}' w archiwum XLSX.`);
  }
  return window.fflate.strFromU8(u8);
}

function resolveWorksheetPath(workbookXml, workbookRelsXml, sheetName) {
  const wbDoc = new DOMParser().parseFromString(workbookXml, "application/xml");
  const relDoc = new DOMParser().parseFromString(workbookRelsXml, "application/xml");
  const workbookSheets = wbDoc.getElementsByTagName("sheet");
  let relId = "";
  for (let i = 0; i < workbookSheets.length; i += 1) {
    const sheet = workbookSheets[i];
    if (sheet.getAttribute("name") === sheetName) {
      relId = sheet.getAttribute("r:id") || sheet.getAttributeNS(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "id"
      ) || "";
      break;
    }
  }
  if (!relId) {
    throw new Error(`Nie znaleziono arkusza '${sheetName}' w workbook.xml.`);
  }

  const relNodes = relDoc.getElementsByTagName("Relationship");
  let target = "";
  for (let i = 0; i < relNodes.length; i += 1) {
    const rel = relNodes[i];
    if (rel.getAttribute("Id") === relId) {
      target = rel.getAttribute("Target") || "";
      break;
    }
  }
  if (!target) {
    throw new Error(`Nie znaleziono relacji arkusza '${sheetName}' (r:id=${relId}).`);
  }
  const clean = target.replace(/^\//, "");
  if (clean.startsWith("xl/")) {
    return clean;
  }
  return `xl/${clean.replace(/^\.?\//, "")}`;
}

function ensureWorkbookFullRecalcOnLoad(workbookXml) {
  const doc = new DOMParser().parseFromString(workbookXml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw new Error("Nie udało się sparsować workbook.xml.");
  }
  const root = doc.documentElement;
  if (!root || root.localName !== "workbook") {
    throw new Error("Niepoprawna struktura workbook.xml.");
  }
  const calcPr = findDirectChildByLocalName(root, "calcPr")
    || getOrCreateDirectChildByLocalName(doc, root, "calcPr");
  calcPr.setAttribute("calcMode", "auto");
  calcPr.setAttribute("fullCalcOnLoad", "1");
  calcPr.setAttribute("forceFullCalc", "1");

  const serialized = new XMLSerializer().serializeToString(doc);
  if (serialized.startsWith("<?xml")) {
    return serialized;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${serialized}`;
}

function updateWorksheetXmlValues(sheetXml, updatePlan) {
  const doc = new DOMParser().parseFromString(sheetXml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0) {
    throw new Error("Nie udało się sparsować XML arkusza master.");
  }
  const sheetData = findDirectChildByLocalName(doc.documentElement, "sheetData");
  if (!sheetData) {
    throw new Error("Arkusz XLSX nie zawiera sekcji sheetData.");
  }
  // Only write to the month columns (holiday + special for the selected month).
  // Do NOT touch Total holiday balance / Total special holiday balance —
  // those cells contain SUM formulas. Writing a raw value there would replace
  // the formula with a single month's number, breaking Process 3 and Process 4.
  // fullCalcOnLoad="1" in workbook.xml ensures Excel recalculates on open.
  updatePlan.updates.forEach((update) => {
    setExistingWorksheetXmlCellByCoords(doc, sheetData, update.rowIdx + 1, updatePlan.monthIdx + 1, update.holiday);
    setExistingWorksheetXmlCellByCoords(doc, sheetData, update.rowIdx + 1, updatePlan.specialIdx + 1, update.special);
  });
  const serialized = new XMLSerializer().serializeToString(doc);
  // Avoid double XML declaration which can corrupt the worksheet.
  if (serialized.startsWith("<?xml")) {
    return serialized;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${serialized}`;
}

function setExistingWorksheetXmlCellByCoords(doc, sheetData, rowNumber, columnNumber, value) {
  const ref = `${columnNumberToName(columnNumber)}${rowNumber}`;
  const rowNode = findWorksheetXmlRow(sheetData, rowNumber);
  if (!rowNode) {
    throw new Error(`Brak wiersza ${rowNumber} w arkuszu master. Zatrzymano zapis, aby nie zmieniać formatowania.`);
  }
  const cellNode = findCellNodeByRef(rowNode, ref);
  if (!cellNode) {
    throw new Error(`Brak komórki ${ref} w arkuszu master. Zatrzymano zapis, aby nie zmieniać formatowania.`);
  }
  writeNumericValueToCell(doc, cellNode, value);
}

function findWorksheetXmlRow(sheetData, rowNumber) {
  const rows = getDirectChildrenByLocalName(sheetData, "row");
  for (let i = 0; i < rows.length; i += 1) {
    if (Number(rows[i].getAttribute("r") || 0) === rowNumber) {
      return rows[i];
    }
  }
  return null;
}

function writeNumericValueToCell(doc, cellNode, value) {
  const hasFormula = Boolean(findDirectChildByLocalName(cellNode, "f"));
  if (value === null || value === undefined || value === "") {
    if (hasFormula) {
      const v = getOrCreateDirectChildByLocalName(doc, cellNode, "v");
      v.textContent = "";
      return;
    }
    removeDirectChildrenByLocalName(cellNode, "v");
    removeDirectChildrenByLocalName(cellNode, "is");
    cellNode.removeAttribute("t");
    return;
  }

  const numericText = String(Number(value));
  cellNode.removeAttribute("t");
  removeDirectChildrenByLocalName(cellNode, "is");
  const valueNode = getOrCreateDirectChildByLocalName(doc, cellNode, "v");
  valueNode.textContent = numericText;
}

function getDirectChildrenByLocalName(parentNode, localName) {
  const result = [];
  const children = parentNode?.childNodes || [];
  for (let i = 0; i < children.length; i += 1) {
    const child = children[i];
    if (child.nodeType === 1 && child.localName === localName) {
      result.push(child);
    }
  }
  return result;
}

function findDirectChildByLocalName(parentNode, localName) {
  const children = parentNode?.childNodes || [];
  for (let i = 0; i < children.length; i += 1) {
    const child = children[i];
    if (child.nodeType === 1 && child.localName === localName) {
      return child;
    }
  }
  return null;
}

function findCellNodeByRef(rowNode, ref) {
  const cells = getDirectChildrenByLocalName(rowNode, "c");
  for (let i = 0; i < cells.length; i += 1) {
    if ((cells[i].getAttribute("r") || "") === ref) {
      return cells[i];
    }
  }
  return null;
}

function getOrCreateDirectChildByLocalName(doc, parentNode, localName) {
  const existing = findDirectChildByLocalName(parentNode, localName);
  if (existing) {
    return existing;
  }
  const created = doc.createElementNS(parentNode.namespaceURI, localName);
  parentNode.appendChild(created);
  return created;
}

function removeDirectChildrenByLocalName(parentNode, localName) {
  const children = [...(parentNode?.childNodes || [])];
  children.forEach((child) => {
    if (child.nodeType === 1 && child.localName === localName) {
      parentNode.removeChild(child);
    }
  });
}

function columnNumberToName(columnNumber) {
  let n = Number(columnNumber);
  let name = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name || "A";
}

function ensureXlsxName(name) {
  return name.endsWith(".xlsx") ? name : `${name}.xlsx`;
}

function round2(value) {
  const num = Number(value);
  if (Number.isNaN(num)) return 0;
  return Math.round(num * 100) / 100;
}

function toNumberOrNull(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return round2(value);
  }
  // Use parseDanishNumber which handles both Danish (1.234,56) and standard formats
  const parsed = parseDanishNumber(String(value));
  if (parsed !== null) {
    return parsed;
  }
  const num = Number(String(value).replaceAll(",", "."));
  return Number.isNaN(num) ? null : round2(num);
}

function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
