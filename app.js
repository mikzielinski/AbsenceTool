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
  if (!file) {
    state.sourceWorkbook = null;
    state.sourceWorkbookName = "";
    state.sourceSheets = [];
    fillSelect(ui.sourceSheet1a, []);
    hideInlineError();
    refreshUiReadiness();
    return;
  }
  try {
    state.sourceWorkbook = await readWorkbookFromFile(file);
    state.sourceWorkbookName = file.name;
    state.sourceSheets = getWorkbookSheetNames(state.sourceWorkbook);
    fillSelect(ui.sourceSheet1a, state.sourceSheets);
    hideInlineError();
    log(`Source loaded: ${file.name} (${state.sourceSheets.length} sheet(s)).`);
    log(`Source sheets: ${state.sourceSheets.join(" | ")}`);
    refreshUiReadiness();
  } catch (error) {
    state.sourceWorkbook = null;
    state.sourceWorkbookName = "";
    state.sourceSheets = [];
    fillSelect(ui.sourceSheet1a, []);
    showError(`Nie udało się wczytać source file: ${error.message}`);
  }
}

async function onMasterFilePicked(event) {
  const file = event.target.files?.[0];
  if (!file) {
    state.masterWorkbook = null;
    state.masterWorkbookName = "";
    state.masterWorkbookBytes = null;
    state.masterSheets = [];
    fillSelect(ui.masterSheet, []);
    hideInlineError();
    refreshUiReadiness();
    return;
  }
  try {
    const loaded = await readWorkbookBundle(file);
    state.masterWorkbook = loaded.workbook;
    state.masterWorkbookName = file.name;
    state.masterWorkbookBytes = loaded.buffer;
    state.masterSheets = getWorkbookSheetNames(state.masterWorkbook);
    fillSelect(ui.masterSheet, state.masterSheets);
    hideInlineError();
    log(`Master loaded: ${file.name} (${state.masterSheets.length} sheet(s)).`);
    log(`Master sheets: ${state.masterSheets.join(" | ")}`);
    refreshUiReadiness();
  } catch (error) {
    state.masterWorkbook = null;
    state.masterWorkbookName = "";
    state.masterWorkbookBytes = null;
    state.masterSheets = [];
    fillSelect(ui.masterSheet, []);
    refreshUiReadiness();
    showError(`Nie udało się wczytać master file: ${error.message}`);
  }
}

async function onP3FileAPicked(event) {
  const file = event.target.files?.[0];
  fillSelect(ui.p3SheetA, []);
  state.p3WorkbookA = null;
  if (!file) { refreshUiReadiness(); return; }
  try {
    // Only read sheet names (needed for the dropdown) — NOT the full workbook.
    // The actual data is always read fresh from the file input in runProcess3.
    const wb = await readWorkbookFromFile(file);
    const sheets = getWorkbookSheetNames(wb);
    state.p3WorkbookA = wb;  // kept only so validateProcess3Inputs can check sheet list
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
  if (!file) { refreshUiReadiness(); return; }
  try {
    const wb = await readWorkbookFromFile(file);
    const sheets = getWorkbookSheetNames(wb);
    state.p3WorkbookB = wb;  // kept only so validateProcess3Inputs can check sheet list
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
  const hasSource = !!state.sourceWorkbook;
  const hasMaster = !!state.masterWorkbook;
  const hasPdfs = (ui.payslipFiles.files || []).length > 0;
  const hasMasterP4 = !!state.masterWorkbookP2;

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
  if (!state.sourceWorkbook) {
    throw new Error("Wybierz Current Holiday Report.");
  }
  if (!ui.sourceSheet1a.value) {
    throw new Error("Wybierz arkusz source dla Process 1.");
  }
}

function validateProcess2Inputs() {
  if (!state.sourceWorkbook) {
    throw new Error("Wybierz Current Holiday Report.");
  }
  if (!state.masterWorkbook) {
    throw new Error("Wybierz Holiday Balance master file.");
  }
  if (!ui.sourceSheet1a.value || !ui.masterSheet.value) {
    throw new Error("Wybierz wymagane sheety dla Process 2.");
  }
}

function validateProcess3Inputs() {
  if (!ui.p3FileA.files || !ui.p3FileA.files[0]) {
    throw new Error("Wybierz plik A dla Process 3.");
  }
  if (!ui.p3FileB.files || !ui.p3FileB.files[0]) {
    throw new Error("Wybierz plik B dla Process 3.");
  }
  if (!ui.p3SheetA.value || !ui.p3SheetB.value) {
    throw new Error("Wybierz wymagane sheety dla Process 3 (plik A i plik B).");
  }
}

async function runProcess1() {
  clearLog();
  try {
    validateProcess1Inputs();
    log("=== Process 1 ===");
    const sourceRows = sheetToRows(state.sourceWorkbook, ui.sourceSheet1a.value);
    const flexiRows = buildFlexiAbsenceRows(sourceRows);

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
  try {
    validateProcess2Inputs();
    log("=== Process 2 ===");
    const monthLabel = ui.monthLabel.value;
    const sourceRows = sheetToRows(state.sourceWorkbook, ui.sourceSheet1a.value);
    const balanceSummaryRows = buildBalanceSummaryFromSource(sourceRows);
    const updatePlan = buildMasterUpdatePlan(
      state.masterWorkbook,
      ui.masterSheet.value,
      balanceSummaryRows,
      monthLabel
    );
    const outBytes = applyMasterUpdatesViaWorksheetXml(state.masterWorkbookBytes, updatePlan);
    state.masterWorkbook = readWorkbookFromBuffer(
      outBytes.buffer.slice(outBytes.byteOffset, outBytes.byteOffset + outBytes.byteLength)
    );
    state.masterWorkbookBytes = outBytes.buffer.slice(
      outBytes.byteOffset,
      outBytes.byteOffset + outBytes.byteLength
    );
    state.masterSheets = getWorkbookSheetNames(state.masterWorkbook);
    fillSelect(ui.masterSheet, state.masterSheets);
    if (state.masterSheets.includes(updatePlan.sheetName)) {
      ui.masterSheet.value = updatePlan.sheetName;
    }
    const outName = buildUpdatedMasterName(state.masterWorkbookName, monthLabel);
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
  try {
    validateProcess3Inputs();
    log("=== Process 3 ===");

    // Always read both files completely fresh from the file input elements.
    // No in-memory state is used — each run is fully independent.
    const fileA = ui.p3FileA.files && ui.p3FileA.files[0];
    const fileB = ui.p3FileB.files && ui.p3FileB.files[0];
    if (!fileA) throw new Error("Wybierz plik A dla Process 3.");
    if (!fileB) throw new Error("Wybierz plik B dla Process 3.");

    const sheetA = ui.p3SheetA.value;
    const sheetB = ui.p3SheetB.value;
    if (!sheetA) throw new Error("Wybierz arkusz z pliku A.");
    if (!sheetB) throw new Error("Wybierz arkusz z pliku B.");

    log(`Plik A: ${fileA.name} / arkusz: ${sheetA}`);
    log(`Plik B: ${fileB.name} / arkusz: ${sheetB}`);

    // Read plik A via raw XML to bypass xlsx.js formula cache.
    // We need column "Total holidays" — exact cell values, not cached formulas.
    log("Czytam plik A (raw XML)...");
    const bufA = await fileA.arrayBuffer();
    const srcSummary = await readSheetColumnsDirect(bufA, sheetA, {
      idCandidates:    ["SAP ID", "Holid", "ID"],
      nameCandidates:  ["Name", "Employee", "Employee Name"],
      valueCandidates: ["Total holidays"],
      valueLabel:      "Total holidays",
    });
    log(`  Plik A: ${srcSummary.length} rekordów.`);
    srcSummary.forEach((r) => log(`  ID=${r["SAP ID"]} | Total holidays=${r["Total holidays"]}`));

    // Read plik B via raw XML — column "Total holiday balance".
    log("Czytam plik B (raw XML)...");
    const bufB = await fileB.arrayBuffer();
    const masterSummary = await readSheetColumnsDirect(bufB, sheetB, {
      idCandidates:    ["SAP ID", "ID", "Holid"],
      nameCandidates:  ["Employee", "Name", "Employee Name"],
      valueCandidates: ["Total holiday balance"],
      valueLabel:      "Total holiday balance",
    });
    log(`  Plik B: ${masterSummary.length} rekordów.`);
    masterSummary.forEach((r) => log(`  ID=${r["SAP ID"]} | Total holiday balance=${r["Total holiday balance"]}`));

    // Build lookup from plik B
    const masterById = {};
    const nameById   = {};
    masterSummary.forEach((r) => {
      masterById[r["SAP ID"]] = r["Total holiday balance"];
      if (r["Employee Name"]) nameById[r["SAP ID"]] = r["Employee Name"];
    });

    // Compare
    const compareRows = srcSummary.map((rec) => {
      const sap = rec["SAP ID"];
      const src = rec["Total holidays"];
      const mst = masterById[sap] !== undefined ? masterById[sap] : null;
      return {
        "SAP ID":                          sap,
        "Employee Name":                   rec["Employee Name"] || nameById[sap] || sap,
        "Total holidays (source)":         src,
        "Total holiday balance (master)":  mst,
        Match: src !== null && mst !== null ? round2(src) === round2(mst) : false,
      };
    });

    const outWb = buildBalanceComparisonWorkbook(compareRows);
    const outName = `Balance comparison ${todayIso()}.xlsx`;
    writeWorkbookDownload(outWb, outName);
    const mismatch = compareRows.filter((row) => row.Match !== true).length;
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
  }
}

function resolveMasterForP2() {
  if (state.masterWorkbookP2 && ui.masterSheetP2.value) {
    return { wb: state.masterWorkbookP2, sheet: ui.masterSheetP2.value };
  }
  throw new Error("Wybierz osobny master file i sheet dla Process 4.");
}

async function runProcess4() {
  clearLog();
  try {
    log("=== Process 4 ===");
    const pdfFiles = Array.from(ui.payslipFiles.files || []);
    if (pdfFiles.length === 0) throw new Error("Wybierz co najmniej jeden payslip PDF.");

    // Always read the master file FRESH from disk — never from in-memory state.
    // The user must pick the already-updated master (output from Process 2).
    if (!state.masterWorkbookP2) throw new Error("Wybierz plik master dla Process 4.");
    const masterSheet = ui.masterSheetP2.value;
    if (!masterSheet) throw new Error("Wybierz arkusz mastera dla Process 4.");

    // Always read master completely fresh via raw XML — no in-memory state used.
    const masterFileEl = ui.masterFileP2;
    const masterFile   = masterFileEl.files && masterFileEl.files[0];
    if (!masterFile) throw new Error("Wybierz plik master dla Process 4.");
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
    const outWb = buildPayslipReportWorkbook(reconciliationRows);
    const outName = `Payslip reconciliation ${todayIso()}.xlsx`;
    writeWorkbookDownload(outWb, outName);
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

// Build a complete cell table from sheet XML.
// Returns: { values: Map<ref, number|string|null>, formulas: Map<ref, string> }
function buildRawCellTable(sheetXml, sharedStrings) {
  const doc = new DOMParser().parseFromString(sheetXml, "application/xml");
  if (doc.getElementsByTagName("parsererror").length > 0)
    throw new Error("Cannot parse sheet XML.");
  const sd = findDirectChildByLocalName(doc.documentElement, "sheetData");
  if (!sd) throw new Error("No sheetData in sheet.");

  const values   = new Map();  // ref → number | string | null
  const formulas = new Map();  // ref → formula string (same-sheet arithmetic only)

  getDirectChildrenByLocalName(sd, "row").forEach((row) => {
    getDirectChildrenByLocalName(row, "c").forEach((cell) => {
      const ref  = cell.getAttribute("r") || "";
      const type = cell.getAttribute("t") || "";
      const vEl  = findDirectChildByLocalName(cell, "v");
      const fEl  = findDirectChildByLocalName(cell, "f");
      const vTxt = vEl ? vEl.textContent.trim() : "";
      const fTxt = fEl ? fEl.textContent.trim() : "";

      // Raw value stored in the file
      let raw = null;
      if (type === "s") {
        const idx = parseInt(vTxt, 10);
        raw = Number.isNaN(idx) ? null : (sharedStrings[idx] || null);
      } else if (vTxt !== "") {
        const n = Number(vTxt);
        raw = Number.isFinite(n) ? n : vTxt;
      }
      values.set(ref, raw);

      // Store formula only if it contains NO cross-sheet references
      // (cross-sheet refs can't be resolved here — we use cached <v> for those)
      if (fTxt) {
        const isCrossSheet = fTxt.includes("!") && (fTxt.includes("'") || /[A-Za-z].*!/.test(fTxt));
        if (!isCrossSheet) {
          formulas.set(ref, fTxt);
        }
        // For cross-sheet refs: raw already holds the cached <v> — keep it as-is
      }
    });
  });

  return { values, formulas };
}

// Evaluate all formulas in-place, writing results back into values.
// Runs multiple passes until stable (handles chains like AE→AG→AJ).
function evaluateFormulas(values, formulas) {
  // Simple tokeniser: replace cell refs with their values, then eval arithmetic
  function evalFormula(formula, values, selfRef) {
    let f = formula.startsWith("=") ? formula.slice(1) : formula;

    // Replace cell refs with their current values
    f = f.replace(/\b([A-Z]{1,3}[0-9]+)\b/g, (match) => {
      if (match === selfRef) return "0";
      const v = values.get(match);
      if (typeof v === "number") return String(v);
      return "0";
    });

    // Safety: only allow arithmetic
    if (!/^[\d\s+\-*/.()e]+$/i.test(f)) return null;
    try {
      // eslint-disable-next-line no-new-func
      const result = Function('"use strict"; return (' + f + ');')();
      return Number.isFinite(result) ? Math.round(result * 100) / 100 : null;
    } catch {
      return null;
    }
  }

  // Up to 10 passes to resolve chains
  for (let pass = 0; pass < 10; pass++) {
    let changed = false;
    formulas.forEach((formula, ref) => {
      const result = evalFormula(formula, values, ref);
      if (result !== null) {
        const prev = values.get(ref);
        if (prev !== result) {
          values.set(ref, result);
          changed = true;
        }
      }
    });
    if (!changed) break;
  }

  return values;
}

async function readSheetColumnsDirect(arrayBuffer, sheetName, opts) {
  if (!window.fflate) throw new Error("fflate not available.");

  const bytes  = new Uint8Array(arrayBuffer);
  const zipped = window.fflate.unzipSync(bytes);

  const sharedStrings = parseSharedStrings(zipped);
  const wbXml  = window.fflate.strFromU8(zipped["xl/workbook.xml"]);
  const wbRels = window.fflate.strFromU8(zipped["xl/_rels/workbook.xml.rels"]);
  const wsPath = resolveWorksheetPath(wbXml, wbRels, sheetName);
  if (!zipped[wsPath]) throw new Error(`Sheet file not found: ${wsPath}`);
  const wsXml = window.fflate.strFromU8(zipped[wsPath]);

  // ── Build raw cell table ──
  const { values, formulas } = buildRawCellTable(wsXml, sharedStrings);

  // ── Evaluate all arithmetic formulas ──
  evaluateFormulas(values, formulas);

  // ── Rebuild row structure from the sheet XML for header/data reading ──
  const doc = new DOMParser().parseFromString(wsXml, "application/xml");
  const sd  = findDirectChildByLocalName(doc.documentElement, "sheetData");
  const xmlRows = getDirectChildrenByLocalName(sd, "row");

  // ── Find headers ──
  const headerMap = {};  // colLetter → normalised lowercase
  const headerRaw = {};  // colLetter → original text

  function isTextRow(row) {
    const cells = getDirectChildrenByLocalName(row, "c");
    for (const cell of cells) {
      const t  = cell.getAttribute("t") || "";
      const vEl = findDirectChildByLocalName(cell, "v");
      const v  = vEl ? vEl.textContent.trim() : "";
      if (!v) continue;
      if (t === "s") return true;
      if (!t && Number.isFinite(Number(v))) return false;
      return true;
    }
    return false;
  }

  function addHeaderRow(row) {
    getDirectChildrenByLocalName(row, "c").forEach((cell) => {
      const ref = cell.getAttribute("r") || "";
      const col = ref.replace(/[0-9]/g, "");
      const val = getCellTextValue(cell, sharedStrings);
      const norm = (val || "").replace(/[\s\n\r]+/g, " ").trim();
      if (norm) {
        headerMap[col] = norm.toLowerCase();
        headerRaw[col] = norm;
      }
    });
  }

  if (xmlRows[0]) addHeaderRow(xmlRows[0]);
  if (xmlRows[1] && isTextRow(xmlRows[1])) addHeaderRow(xmlRows[1]);

  // ── Column finder ──
  function findCol(candidates) {
    // 1. Exact
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/[\s\n\r]+/g, " ").trim();
      for (const [col, hdr] of Object.entries(headerMap)) {
        if (hdr === cl) return col;
      }
    }
    // 2. Header starts with candidate
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/[\s\n\r]+/g, " ").trim();
      if (cl.length < 6) continue;
      for (const [col, hdr] of Object.entries(headerMap)) {
        if (hdr.startsWith(cl)) return col;
      }
    }
    // 3. Candidate starts with header
    for (const cand of candidates) {
      const cl = cand.toLowerCase().replace(/[\s\n\r]+/g, " ").trim();
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

  // ── Read data rows using the evaluated cell table ──
  const result = [];
  xmlRows.forEach((row) => {
    const rNum = row.getAttribute("r");
    const idRef = `${idCol}${rNum}`;
    const idRaw = values.get(idRef);
    const sap   = normalizeSap(idRaw);
    if (!sap || !/^[0-9]+$/.test(sap)) return;  // skip header/empty rows

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

function buildBalanceComparisonWorkbook(rows) {
  const headers = [
    "SAP ID",
    "Employee Name",
    "Total holidays (source)",
    "Total holiday balance (master)",
    "Match",
  ];
  const wsRows = [headers, ...rows.map((row) => headers.map((h) => row[h] ?? null))];
  const ws = XLSX.utils.aoa_to_sheet(wsRows);
  applyHeaderStyles(ws, headers.length);
  rows.forEach((row, idx) => {
    if (row.Match !== true) {
      const ref = XLSX.utils.encode_cell({ r: idx + 1, c: 2 });
      if (!ws[ref]) ws[ref] = { t: "s", v: "" };
      ws[ref].s = { fill: RED_FILL };
    }
  });
  autoWidth(ws, wsRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Balance comparison");
  return wb;
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
  const tokens = line.match(/-?[0-9][0-9\s.,]*/g) || [];
  return tokens.map((t) => parseDanishNumber(t.replace(/\s/g, ""))).filter((n) => n !== null);
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
  const wsRows = [headers, ...rows.map((row) => headers.map((h) => row[h] ?? null))];
  const ws = XLSX.utils.aoa_to_sheet(wsRows);
  applyHeaderStyles(ws, headers.length);
  rows.forEach((row, idx) => {
    const r = idx + 1;
    if (row.Result !== true) {
      markRed(ws, r, 2);
      markRed(ws, r, 3);
      markRed(ws, r, 4);
    }
    if (row["Special result"] !== true) {
      markRed(ws, r, 5);
      markRed(ws, r, 6);
      markRed(ws, r, 7);
    }
  });
  autoWidth(ws, wsRows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Reconciliation");
  return wb;
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
  XLSX.writeFile(workbook, ensureXlsxName(filename));
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
