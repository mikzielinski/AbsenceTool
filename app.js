"use strict";

const compareButton = document.getElementById("compareButton");
const startYearInput = document.getElementById("startYear");
const endYearInput = document.getElementById("endYear");
const payrollDatesInput = document.getElementById("payrollDates");
const errorsContainer = document.getElementById("errors");
const summaryCards = document.getElementById("summaryCards");
const missingTableContainer = document.getElementById("missingTable");
const extraTableContainer = document.getElementById("extraTable");

const holidayNameByKey = {
  "01-01": "New Year's Day",
  "12-25": "Christmas Day",
  "12-26": "2nd Christmas Day"
};

initializeDefaults();
renderSummary(0, 0, 0);
renderTable(missingTableContainer, []);
renderTable(extraTableContainer, []);

compareButton.addEventListener("click", () => {
  errorsContainer.textContent = "";

  const startYear = Number(startYearInput.value);
  const endYear = Number(endYearInput.value);

  if (!isValidYear(startYear) || !isValidYear(endYear)) {
    errorsContainer.textContent = "Podaj poprawny rok od i do (1900-2100).";
    return;
  }

  if (startYear > endYear) {
    errorsContainer.textContent = "Rok od nie może być większy niż rok do.";
    return;
  }

  const parseResult = parsePayrollDates(payrollDatesInput.value);
  if (parseResult.invalidTokens.length > 0) {
    errorsContainer.textContent = `Niepoprawne daty: ${parseResult.invalidTokens.join(", ")}`;
    return;
  }

  const officialHolidays = buildOfficialHolidaySet(startYear, endYear);
  const payrollDates = parseResult.dates;

  const missingDates = [...officialHolidays].filter((date) => !payrollDates.has(date)).sort();
  const extraDates = [...payrollDates].filter((date) => !officialHolidays.has(date)).sort();

  renderSummary(officialHolidays.size, missingDates.length, extraDates.length);
  renderTable(missingTableContainer, missingDates);
  renderTable(extraTableContainer, extraDates);
});

function initializeDefaults() {
  const currentYear = new Date().getFullYear();
  startYearInput.value = currentYear;
  endYearInput.value = currentYear + 1;
}

function isValidYear(year) {
  return Number.isInteger(year) && year >= 1900 && year <= 2100;
}

function parsePayrollDates(rawText) {
  const invalidTokens = [];
  const dates = new Set();
  const tokens = rawText
    .split(/[\s,;]+/)
    .map((value) => value.trim())
    .filter(Boolean);

  for (const token of tokens) {
    if (!isValidIsoDate(token)) {
      invalidTokens.push(token);
      continue;
    }
    dates.add(token);
  }

  return { dates, invalidTokens };
}

function isValidIsoDate(dateText) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateText)) {
    return false;
  }

  const [year, month, day] = dateText.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year &&
    date.getUTCMonth() === month - 1 &&
    date.getUTCDate() === day
  );
}

function buildOfficialHolidaySet(startYear, endYear) {
  const holidays = new Set();

  for (let year = startYear; year <= endYear; year += 1) {
    addDate(holidays, `${year}-01-01`);
    addDate(holidays, `${year}-12-25`);
    addDate(holidays, `${year}-12-26`);

    const easterSunday = getEasterSundayUtc(year);
    addDate(holidays, formatDate(addDays(easterSunday, -3)), "Maundy Thursday");
    addDate(holidays, formatDate(addDays(easterSunday, -2)), "Good Friday");
    addDate(holidays, formatDate(addDays(easterSunday, 0)), "Easter Sunday");
    addDate(holidays, formatDate(addDays(easterSunday, 1)), "Easter Monday");
    addDate(holidays, formatDate(addDays(easterSunday, 26)), "General Prayer Day");
    addDate(holidays, formatDate(addDays(easterSunday, 39)), "Ascension Day");
    addDate(holidays, formatDate(addDays(easterSunday, 49)), "Whit Sunday");
    addDate(holidays, formatDate(addDays(easterSunday, 50)), "Whit Monday");
  }

  return holidays;
}

function addDate(set, dateString, optionalName) {
  set.add(dateString);
  if (optionalName) {
    holidayNameByKey[dateString.slice(5)] = optionalName;
  }
}

function getEasterSundayUtc(year) {
  const a = year % 19;
  const b = Math.floor(year / 100);
  const c = year % 100;
  const d = Math.floor(b / 4);
  const e = b % 4;
  const f = Math.floor((b + 8) / 25);
  const g = Math.floor((b - f + 1) / 3);
  const h = (19 * a + b - d - g + 15) % 30;
  const i = Math.floor(c / 4);
  const k = c % 4;
  const l = (32 + 2 * e + 2 * i - h - k) % 7;
  const m = Math.floor((a + 11 * h + 22 * l) / 451);
  const month = Math.floor((h + l - 7 * m + 114) / 31);
  const day = ((h + l - 7 * m + 114) % 31) + 1;

  return new Date(Date.UTC(year, month - 1, day));
}

function addDays(date, deltaDays) {
  const cloned = new Date(date.getTime());
  cloned.setUTCDate(cloned.getUTCDate() + deltaDays);
  return cloned;
}

function formatDate(date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function renderSummary(officialCount, missingCount, extraCount) {
  summaryCards.innerHTML = `
    <div class="card">
      <div class="label">Oficjalne święta DK (zakres)</div>
      <div class="value">${officialCount}</div>
    </div>
    <div class="card">
      <div class="label">Brakujące</div>
      <div class="value ${missingCount === 0 ? "ok" : "warn"}">${missingCount}</div>
    </div>
    <div class="card">
      <div class="label">Nadmiarowe</div>
      <div class="value ${extraCount === 0 ? "ok" : "danger"}">${extraCount}</div>
    </div>
  `;
}

function renderTable(container, dates) {
  if (dates.length === 0) {
    container.innerHTML = '<p class="empty-state">Brak pozycji.</p>';
    return;
  }

  const rows = dates
    .map((date) => {
      const holidayName = holidayNameByKey[date.slice(5)] || "-";
      return `<tr><td>${date}</td><td>${holidayName}</td></tr>`;
    })
    .join("");

  container.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Data</th>
          <th>Nazwa święta</th>
        </tr>
      </thead>
      <tbody>
        ${rows}
      </tbody>
    </table>
  `;
}
