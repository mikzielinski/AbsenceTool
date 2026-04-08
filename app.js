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
  masterWorkbook: null,
  masterWorkbookName: "",
  masterWorkbookBytes: null,
  masterFileHandle: null,
  masterSheets: [],
  masterWorkbookP2: null,
  masterWorkbookNameP2: "",
  masterSheetsP2: [],
};

const ui = {
  monthLabel: document.getElementById("monthLabel"),
  tabP1: document.getElementById("tabP1"),
  tabP2: document.getElementById("tabP2"),
  panelP1: document.getElementById("panelP1"),
  panelP2: document.getElementById("panelP2"),
  sourceFile: document.getElementById("sourceFile"),
  sourceSheet1a: document.getElementById("sourceSheet1a"),
  sourceSheet1b: document.getElementById("sourceSheet1b"),
  masterFileFs: document.getElementById("masterFileFs"),
  masterSheet: document.getElementById("masterSheet"),
  masterFileP2: document.getElementById("masterFileP2"),
  masterSheetP2: document.getElementById("masterSheetP2"),
  payslipFiles: document.getElementById("payslipFiles"),
  run1a: document.getElementById("run1a"),
  run1b: document.getElementById("run1b"),
  run2: document.getElementById("run2"),
  resultSummary: document.getElementById("resultSummary"),
  logOutput: document.getElementById("logOutput"),
  statusSource: document.getElementById("sourceStatus"),
  statusMaster: document.getElementById("masterStatus"),
  statusMasterP2: document.getElementById("masterP2Status"),
  statusPdfs: document.getElementById("payslipStatus"),
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
  fillSelect(ui.sourceSheet1b, []);
  fillSelect(ui.masterSheet, []);
  fillSelect(ui.masterSheetP2, []);
  refreshUiReadiness();
  log("Gotowe. Zrób kroki 1-2-3 na ekranie.");
  if (ui.masterFileFs && !supportsFsAccess()) {
    ui.masterFileFs.classList.add("hidden");
  }
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
  ui.sourceFile.addEventListener("change", onSourceFilePicked);
  if (ui.masterFileFs) {
    ui.masterFileFs.addEventListener("click", onMasterFileFsPicked);
  }
  ui.masterFileP2.addEventListener("change", onMasterFileP2Picked);
  ui.payslipFiles.addEventListener("change", refreshUiReadiness);
  ui.sourceSheet1a.addEventListener("change", refreshUiReadiness);
  ui.sourceSheet1b.addEventListener("change", refreshUiReadiness);
  ui.masterSheet.addEventListener("change", refreshUiReadiness);
  ui.masterSheetP2.addEventListener("change", refreshUiReadiness);
  ui.run1a.addEventListener("click", runProcess1a);
  ui.run1b.addEventListener("click", runProcess1b);
  ui.run2.addEventListener("click", runProcess2);
}

function setActiveTab(key) {
  const onP1 = key === "p1";
  ui.tabP1.classList.toggle("active", onP1);
  ui.tabP2.classList.toggle("active", !onP1);
  ui.panelP1.classList.toggle("hidden", !onP1);
  ui.panelP2.classList.toggle("hidden", onP1);
  refreshUiReadiness();
}

async function onSourceFilePicked(event) {
  const file = event.target.files?.[0];
  if (!file) {
    state.sourceWorkbook = null;
    state.sourceWorkbookName = "";
    state.sourceSheets = [];
    fillSelect(ui.sourceSheet1a, []);
    fillSelect(ui.sourceSheet1b, []);
    hideInlineError();
    refreshUiReadiness();
    return;
  }
  try {
    state.sourceWorkbook = await readWorkbookFromFile(file);
    state.sourceWorkbookName = file.name;
    state.sourceSheets = getWorkbookSheetNames(state.sourceWorkbook);
    fillSelect(ui.sourceSheet1a, state.sourceSheets);
    fillSelect(ui.sourceSheet1b, state.sourceSheets);
    hideInlineError();
    log(`Source loaded: ${file.name} (${state.sourceSheets.length} sheet(s)).`);
    log(`Source sheets: ${state.sourceSheets.join(" | ")}`);
    refreshUiReadiness();
  } catch (error) {
    state.sourceWorkbook = null;
    state.sourceWorkbookName = "";
    state.sourceSheets = [];
    fillSelect(ui.sourceSheet1a, []);
    fillSelect(ui.sourceSheet1b, []);
    showError(`Nie udało się wczytać source file: ${error.message}`);
  }
}

async function onMasterFileFsPicked() {
  if (!supportsFsAccess()) {
    showError("Ta przeglądarka nie wspiera bezpośredniego nadpisywania pliku.");
    return;
  }
  try {
    const [handle] = await window.showOpenFilePicker({
      multiple: false,
      types: [
        {
          description: "Excel workbook",
          accept: {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"],
          },
        },
      ],
    });
    if (!handle) return;
    const file = await handle.getFile();
    const loaded = await readWorkbookBundle(file);
    state.masterWorkbook = loaded.workbook;
    state.masterWorkbookBytes = loaded.buffer;
    state.masterWorkbookName = file.name;
    state.masterFileHandle = handle;
    state.masterSheets = getWorkbookSheetNames(state.masterWorkbook);
    fillSelect(ui.masterSheet, state.masterSheets);
    hideInlineError();
    log(`Master loaded in overwrite mode: ${file.name} (${state.masterSheets.length} sheet(s)).`);
    log(`Master sheets: ${state.masterSheets.join(" | ")}`);
    refreshUiReadiness();
  } catch (error) {
    if (error && error.name === "AbortError") {
      log("Wybór pliku master anulowany.");
      return;
    }
    showError(`Nie udało się wczytać master file (tryb nadpisania): ${error.message}`);
  }
}

async function onMasterFileP2Picked(event) {
  const file = event.target.files?.[0];
  if (!file) {
    state.masterWorkbookP2 = null;
    state.masterWorkbookNameP2 = "";
    state.masterSheetsP2 = [];
    fillSelect(ui.masterSheetP2, []);
    hideInlineError();
    refreshUiReadiness();
    return;
  }
  try {
    state.masterWorkbookP2 = await readWorkbookFromFile(file);
    state.masterWorkbookNameP2 = file.name;
    state.masterSheetsP2 = getWorkbookSheetNames(state.masterWorkbookP2);
    fillSelect(ui.masterSheetP2, state.masterSheetsP2);
    hideInlineError();
    log(`Master (P2) loaded: ${file.name} (${state.masterSheetsP2.length} sheet(s)).`);
    log(`Master (P2) sheets: ${state.masterSheetsP2.join(" | ")}`);
    refreshUiReadiness();
  } catch (error) {
    state.masterWorkbookP2 = null;
    state.masterWorkbookNameP2 = "";
    state.masterSheetsP2 = [];
    fillSelect(ui.masterSheetP2, []);
    showError(`Nie udało się wczytać master file (P2): ${error.message}`);
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
  const hasMasterP2 = !!state.masterWorkbookP2;
  const hasPdfs = (ui.payslipFiles.files || []).length > 0;
  setBadge(
    ui.statusSource,
    hasSource,
    hasSource
      ? `OK: ${state.sourceWorkbookName} (${state.sourceSheets.length} ark.)`
      : "Brak pliku source"
  );
  setBadge(
    ui.statusMaster,
    hasMaster,
    hasMaster
      ? `OK: ${state.masterWorkbookName} (${state.masterSheets.length} ark.)`
      : "Brak pliku master"
  );
  setBadge(
    ui.statusMasterP2,
    hasMasterP2,
    hasMasterP2
      ? `OK: ${state.masterWorkbookNameP2} (${state.masterSheetsP2.length} ark.)`
      : "Opcjonalne"
  );
  setBadge(ui.statusPdfs, hasPdfs, hasPdfs ? `OK: ${ui.payslipFiles.files.length} PDF` : "Brak PDF");

  const readyP1 = hasSource && hasMaster && !!ui.sourceSheet1a.value && !!ui.sourceSheet1b.value && !!ui.masterSheet.value;
  const p2MasterReady = (state.masterWorkbookP2 && ui.masterSheetP2.value) || (hasMaster && ui.masterSheet.value);
  const readyP2 = hasPdfs && !!p2MasterReady;
  ui.run1a.disabled = !readyP1;
  ui.run1b.disabled = !readyP1;
  ui.run2.disabled = !readyP2;
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

function validateP1Inputs() {
  if (!state.sourceWorkbook) {
    throw new Error("Wybierz Current Holiday Report.");
  }
  if (!state.masterWorkbook) {
    throw new Error("Wybierz Holiday Balance master file.");
  }
  if (!ui.sourceSheet1a.value || !ui.sourceSheet1b.value || !ui.masterSheet.value) {
    throw new Error("Wybierz wymagane sheety dla Process 1.");
  }
}

async function runProcess1a() {
  clearLog();
  try {
    validateP1Inputs();
    log("=== Process 1a ===");
    const monthLabel = ui.monthLabel.value;
    const sourceRows = sheetToRows(state.sourceWorkbook, ui.sourceSheet1a.value);
    const flexiRows = buildFlexiAbsenceRows(sourceRows);
    const balanceSummaryRows = buildBalanceSummaryFromSource(sourceRows);

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
    log(`Flexi file generated (${flexiRows.length} row(s)).`);

    const updatePlan = buildMasterUpdatePlan(
      state.masterWorkbook,
      ui.masterSheet.value,
      balanceSummaryRows,
      monthLabel
    );
    const saveResult = await saveMasterWorkbook(updatePlan);
    applyMasterUpdatePlanToWorkbook(state.masterWorkbook, updatePlan);
    log(`Master updated (${updatePlan.updates.length} employee(s)).`);

    setResultSummary([
      `Process 1a zakończony.`,
      `Flexi row(s): ${flexiRows.length}.`,
      `Master updated row(s): ${updatePlan.updates.length}.`,
      `Master nadpisany: ${saveResult.name}.`,
      `Pobrane pliki: Flexi-absence input.xlsx.`,
    ]);
  } catch (error) {
    showError(error.message);
  }
}

async function runProcess1b() {
  clearLog();
  try {
    validateP1Inputs();
    log("=== Process 1b ===");
    const srcSummary = loadSourceBalanceSummary(
      sheetToRows(state.sourceWorkbook, ui.sourceSheet1b.value)
    );
    const { totals, names } = readCollectiveTotals(
      state.masterWorkbook,
      ui.masterSheet.value
    );
    const compareRows = buildBalanceComparisonRows(srcSummary, totals, names);
    const outWb = buildBalanceComparisonWorkbook(compareRows);
    const outName = `Balance comparison ${todayIso()}.xlsx`;
    writeWorkbookDownload(outWb, outName);
    const mismatch = compareRows.filter((row) => row.Match !== true).length;
    setResultSummary([
      "Process 1b zakończony.",
      `Records: ${compareRows.length}.`,
      `Matched: ${compareRows.length - mismatch}.`,
      `Mismatched: ${mismatch}.`,
      `Pobrany plik: ${outName}.`,
    ]);
    log(`Comparison completed. Mismatch(es): ${mismatch}.`);
  } catch (error) {
    showError(error.message);
  }
}

function resolveMasterForP2() {
  if (state.masterWorkbookP2 && ui.masterSheetP2.value) {
    return { wb: state.masterWorkbookP2, sheet: ui.masterSheetP2.value };
  }
  if (state.masterWorkbook && ui.masterSheet.value) {
    return { wb: state.masterWorkbook, sheet: ui.masterSheet.value };
  }
  throw new Error("Wybierz master file/sheet dla Process 2 (lub użyj z Process 1).");
}

async function runProcess2() {
  clearLog();
  try {
    log("=== Process 2 ===");
    const pdfFiles = Array.from(ui.payslipFiles.files || []);
    if (pdfFiles.length === 0) {
      throw new Error("Wybierz co najmniej jeden payslip PDF.");
    }
    const { wb, sheet } = resolveMasterForP2();
    const { totals, specials, names } = readCollectiveTotals(wb, sheet);
    const records = [];
    for (const file of pdfFiles) {
      log(`Parsing ${file.name} ...`);
      const fileRecords = await parsePayslipBatch(file);
      records.push(...fileRecords);
      log(`  ${fileRecords.length} employee record(s).`);
    }
    const reconciliationRows = buildPayslipReportRows(records, totals, specials, names);
    const outWb = buildPayslipReportWorkbook(reconciliationRows);
    const outName = `Payslip reconciliation ${todayIso()}.xlsx`;
    writeWorkbookDownload(outWb, outName);
    const mismatches = reconciliationRows.filter((row) => !(row.Result && row["Special result"])).length;
    setResultSummary([
      "Process 2 zakończony.",
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
    throw new Error(`Brak kolumn w source sheet (1a): ${missing.join(", ")}`);
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
    throw new Error("Brak pasujących rekordów 1a po filtrze project map.");
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
  const idCol = findColumn(cols, ["Holid", "SAP ID", "ID"]);
  const nameCol = findColumn(cols, ["Name", "Employee", "Employee Name"]);
  const totalCol = findColumn(cols, ["Total holidays", "Total holiday balance"]);
  if (!idCol || !totalCol) {
    throw new Error(`Holiday Balance tab missing required columns. Found: ${cols.join(", ")}`);
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
  const byLower = new Map(columns.map((c) => [String(c).trim().toLowerCase(), c]));
  for (const candidate of candidates) {
    const key = candidate.toLowerCase();
    if (byLower.has(key)) {
      return byLower.get(key);
    }
  }
  return null;
}

function readCollectiveTotals(workbook, sheetName) {
  const rows = sheetToRows(workbook, sheetName);
  if (rows.length === 0) {
    throw new Error("Master sheet is empty.");
  }
  const header = Object.keys(rows[0]);
  const holidayCol = findColumnContaining(header, "total holiday balance");
  const specialCol = findColumnContaining(header, "total special holiday balance");
  const nameCol = findColumn(header, ["Employee", "Name", "Employee Name"]);
  if (!holidayCol) {
    throw new Error("Master: 'Total holiday balance' column not found.");
  }

  const totals = {};
  const specials = {};
  const names = {};
  rows.forEach((row) => {
    const sap = normalizeSap(row[header[0]]);
    if (!sap) return;
    totals[sap] = round2(Number(row[holidayCol] || 0));
    specials[sap] = round2(Number((specialCol && row[specialCol]) || 0));
    if (nameCol && row[nameCol]) {
      names[sap] = String(row[nameCol]).trim();
    }
  });
  return { totals, specials, names };
}

function findColumnContaining(columns, fragment) {
  const f = fragment.toLowerCase();
  return columns.find((col) => String(col).toLowerCase().includes(f)) || null;
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

async function parsePayslipBatch(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const loadingTask = window.pdfjsLib.getDocument({ data: bytes });
  const pdf = await loadingTask.promise;
  const records = [];
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const textContent = await page.getTextContent();
    const text = textContent.items.map((item) => item.str).join("\n");
    if (!/Employeeno/i.test(text)) {
      continue;
    }
    const rec = parseEmployeeBlock(text);
    if (rec.sap_id !== "UNKNOWN") {
      records.push(rec);
    }
  }
  return records;
}

function parseEmployeeBlock(text) {
  const em = text.match(/Employeeno[:\s]+(\d+)/i);
  const sap = em ? em[1].trim() : "UNKNOWN";
  const nameMatch = text.match(/^([A-Za-z][A-Za-z\s]+?)\s*\nEmployeeno/im);
  const name = nameMatch ? nameMatch[1].trim() : "Unknown";
  const hTotal = text.match(/Holidays total\s+([\d.,\-]+)\s+([\d.,\-]+)/i);
  const pHoliday = hTotal ? parseFlexibleNum(hTotal[2]) : null;
  const specialTwo = text.match(/Special holidays\s+([\d.,\-]+)\s+([\d.,\-]+)/i);
  const specialOne = text.match(/Special holidays\s+([\d.,\-]+)/i);
  const pSpecial = specialTwo ? parseFlexibleNum(specialTwo[2]) : (specialOne ? parseFlexibleNum(specialOne[1]) : null);
  return {
    sap_id: sap,
    name,
    payslip_holidays_total: pHoliday,
    payslip_special_total: pSpecial,
  };
}

function parseFlexibleNum(raw) {
  const value = String(raw || "").trim();
  if (!value) return null;
  const hasDot = value.includes(".");
  const hasComma = value.includes(",");
  if (hasDot && hasComma) {
    if (value.lastIndexOf(".") > value.lastIndexOf(",")) {
      return round2(Number(value.replaceAll(",", "")));
    }
    return round2(Number(value.replaceAll(".", "").replaceAll(",", ".")));
  }
  if (hasComma) {
    return round2(Number(value.replaceAll(",", ".")));
  }
  return round2(Number(value));
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

async function saveMasterWorkbook(updatePlan) {
  const outputName = ensureXlsxName(state.masterWorkbookName || "master.xlsx");
  if (!state.masterWorkbookBytes) {
    throw new Error("Brak oryginalnych danych pliku master do zapisu.");
  }
  if (!window.fflate) {
    throw new Error("Brak biblioteki do zapisu punktowego XLSX (fflate).");
  }
  if (state.masterFileHandle && supportsFsAccess()) {
    const updatedBytes = applyMasterUpdatesViaWorksheetXml(state.masterWorkbookBytes, updatePlan);
    await writeArrayToHandle(state.masterFileHandle, updatedBytes);
    state.masterWorkbookBytes = updatedBytes.buffer.slice(
      updatedBytes.byteOffset,
      updatedBytes.byteOffset + updatedBytes.byteLength
    );
    return { mode: "overwrite", name: outputName };
  }

  // Hard requirement: never generate a new master download file.
  if (!supportsFsAccess()) {
    throw new Error("Ta przeglądarka nie wspiera nadpisania pliku master. Użyj Chrome/Edge i przycisku trybu NADPISANIA.");
  }

  throw new Error("Wybierz master przez przycisk 'tryb NADPISANIA', aby zapisać zmiany w tym samym pliku.");
}

function applyMasterUpdatesViaWorksheetXml(workbookBytes, updatePlan) {
  const bytes = new Uint8Array(workbookBytes);
  const zipped = window.fflate.unzipSync(bytes);
  const workbookXmlPath = "xl/workbook.xml";
  const workbookRelsXmlPath = "xl/_rels/workbook.xml.rels";
  const workbookXml = getZipXml(zipped, workbookXmlPath);
  const workbookRelsXml = getZipXml(zipped, workbookRelsXmlPath);
  const sheetFilePath = resolveWorksheetPath(workbookXml, workbookRelsXml, updatePlan.sheetName);
  const sheetXml = getZipXml(zipped, sheetFilePath);
  const updatedSheetXml = updateWorksheetXmlValues(sheetXml, updatePlan);
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

function updateWorksheetXmlValues(sheetXml, updatePlan) {
  let updatedXml = sheetXml;
  updatePlan.updates.forEach((update) => {
    updatedXml = updateWorksheetXmlCellByRef(
      updatedXml,
      `${columnNumberToName(updatePlan.monthIdx + 1)}${update.rowIdx + 1}`,
      update.holiday
    );
    updatedXml = updateWorksheetXmlCellByRef(
      updatedXml,
      `${columnNumberToName(updatePlan.specialIdx + 1)}${update.rowIdx + 1}`,
      update.special
    );
  });
  return updatedXml;
}

function updateWorksheetXmlCellByRef(sheetXml, ref, value) {
  const escapedRef = escapeRegex(ref);
  const pattern = new RegExp(`<c\\b[^>]*\\br="${escapedRef}"[^>]*>[\\s\\S]*?<\\/c>`);
  const match = sheetXml.match(pattern);
  if (!match) {
    return sheetXml;
  }
  const originalCell = match[0];
  const updatedCell = rewriteCellXmlValue(originalCell, value);
  if (updatedCell === originalCell) {
    return sheetXml;
  }
  return sheetXml.replace(originalCell, updatedCell);
}

function rewriteCellXmlValue(cellXml, value) {
  const openTagMatch = cellXml.match(/^<c\b[^>]*>/);
  const closeTag = "</c>";
  if (!openTagMatch || !cellXml.endsWith(closeTag)) {
    return cellXml;
  }
  let openTag = openTagMatch[0];
  let inner = cellXml.slice(openTag.length, -closeTag.length);
  const hasFormula = /<f[\s>][\s\S]*?<\/f>/.test(inner);

  if (value === null || value === undefined || value === "") {
    if (hasFormula) {
      if (/<v(?:\s[^>]*)?>[\s\S]*?<\/v>/.test(inner)) {
        inner = inner.replace(/<v(?:\s[^>]*)?>[\s\S]*?<\/v>/, "<v></v>");
      } else {
        inner += "<v></v>";
      }
      return `${openTag}${inner}${closeTag}`;
    }
    inner = inner.replace(/<v(?:\s[^>]*)?>[\s\S]*?<\/v>/g, "");
    inner = inner.replace(/<is[\s\S]*?<\/is>/g, "");
    return `${openTag}${inner}${closeTag}`;
  }

  const numericText = String(Number(value));
  openTag = openTag.replace(/\s+t="[^"]*"/g, "");
  inner = inner.replace(/<is[\s\S]*?<\/is>/g, "");
  if (/<v(?:\s[^>]*)?>[\s\S]*?<\/v>/.test(inner)) {
    inner = inner.replace(/<v(?:\s[^>]*)?>[\s\S]*?<\/v>/, `<v>${numericText}</v>`);
  } else if (hasFormula) {
    inner = inner.replace(/(<f[\s\S]*?<\/f>)/, `$1<v>${numericText}</v>`);
  } else {
    inner += `<v>${numericText}</v>`;
  }
  return `${openTag}${inner}${closeTag}`;
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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

function cellRefToPos(ref) {
  const match = String(ref || "").match(/^([A-Z]+)(\d+)$/);
  if (!match) return { col: 0, row: 0 };
  const letters = match[1];
  const row = Number(match[2]);
  let col = 0;
  for (let i = 0; i < letters.length; i += 1) {
    col = col * 26 + (letters.charCodeAt(i) - 64);
  }
  return { col, row };
}

function compareCellRefs(a, b) {
  const pa = cellRefToPos(a);
  const pb = cellRefToPos(b);
  if (pa.row !== pb.row) return pa.row - pb.row;
  return pa.col - pb.col;
}

async function writeArrayToHandle(handle, array) {
  const writable = await handle.createWritable();
  await writable.write(array);
  await writable.close();
}

function supportsFsAccess() {
  return Boolean(window.isSecureContext && window.showSaveFilePicker && window.showOpenFilePicker);
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
  const num = Number(String(value).replaceAll(",", "."));
  return Number.isNaN(num) ? null : round2(num);
}

function todayIso() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
