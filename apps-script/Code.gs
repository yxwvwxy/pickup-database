function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Auto")
    .addItem("Parse lists", "parseDailyCarrierLists")
    .addItem("Fill today", "fillToday")
    .addItem("Confirm", "confirmToday")
    .addItem("Run manual", "runManualToday")
    .addToUi();
  rolloverDailyListIfNewDay_();
}

function onEdit(e) {
  restorePastePlaceholderOnEdit_(e);
}

function normalize(str) {
  return String(str || "").replace(/\u00A0/g, " ").trim().toLowerCase();
}

function compact(str) {
  return normalize(str).replace(/\s+/g, "");
}

function mapCarrierName(value) {
  const raw = String(value || "").trim();
  if (normalize(raw) === "80s") return "Han Express";
  return raw;
}

const CHARTER_GROUPS = [
  { id: "CHARTER_FM", locations: ["SF NJ122", "SF-FM", "SF FM", "JD NJ1570", "Yanwen-FM"] },
  { id: "CHARTER_SHIPCUBE", locations: ["ShipCube PA11200", "ShipCube PA700"] },
  { id: "CHARTER_CAPACITY_1101_1112_1000", locations: ["Capacity NJ1101", "Capacity NJ1112", "Capacity NJ1000"] },
  { id: "CHARTER_CAPACITY_1980_1600", locations: ["Capacity NJ1980", "Capacity NJ1600"] }
];

const COL = { ADDRESS: 2, CARRIER: 3, TRUCK: 4, STATE: 5, PRICE: 6, ARRIVAL: 7, PALLETS: 8, TIKTOK_LABEL: 9, TIKTOK_PRICE: 10, TIKTOK_TRUCK: 11 };

const TIKTOK_MERCHANTS = ["NJ TT1001", "NJ TT245", "NJ TT511", "Swift X NJ650"];
const TIKTOK_PRICE_ORDER = [200, 100, 360, 180];

const RULES = [
  { type: "carrier", value: "NYQZ", truckRule: "53'", useDbPrice: true },
  { type: "carrier", value: "运力卡车", truckRule: "DB", useDbPrice: true },
  { type: "location", value: "Emsom CT500", truckRule: "53'", useDbPrice: true }
];

function getSpreadsheetTz() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}

function parseDateString(str) {
  const s = String(str || "").trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[1]) - 1, Number(m[2]));
  return null;
}

function isValidDate(value) {
  const d = new Date(value);
  return value !== "" && value !== null && value !== undefined && !isNaN(d.getTime());
}

function parseSheetDate(cellOrValue, displayValue) {
  if (displayValue !== undefined) {
    const fromDisplay = parseDateString(displayValue);
    if (fromDisplay) return fromDisplay;
  }
  if (cellOrValue && typeof cellOrValue.getDisplayValue === "function") {
    const shown = parseDateString(cellOrValue.getDisplayValue());
    if (shown) return shown;
    return sheetDateToLocalDate(cellOrValue.getValue());
  }
  const fromString = parseDateString(cellOrValue);
  if (fromString) return fromString;
  return sheetDateToLocalDate(cellOrValue);
}

function sheetDateToLocalDate(value) {
  if (!isValidDate(value)) return null;
  const str = Utilities.formatDate(new Date(value), getSpreadsheetTz(), "yyyy-MM-dd");
  const parts = str.split("-");
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function sheetDateTime_(value) {
  const d = parseSheetDate(value);
  return d ? d.getTime() : 0;
}

function isEffectiveOnOrBefore(effectDate, selectedDate) {
  const effect = parseSheetDate(effectDate);
  const selected = parseSheetDate(selectedDate);
  if (!effect) return true;
  if (!selected) return false;
  return effect.getTime() <= selected.getTime();
}

function matchLocation(input, dbLocation) {
  const a = normalize(input);
  const b = normalize(dbLocation);
  if (!a || !b) return false;
  if (a === b) return true;
  if (compact(input) === compact(dbLocation)) return true;
  if (b.startsWith(a + " ") || b.startsWith(a + "\t")) return true;
  const ka = stationAliasKeys_(input);
  const kb = stationAliasKeys_(dbLocation);
  return ka.some(k => kb.indexOf(k) !== -1);
}

function loadRules_() {
  return RULES.map(r => ({
    type: normalize(r.type),
    value: normalize(r.value),
    truckRule: r.truckRule,
    useDbPrice: r.useDbPrice === true || String(r.useDbPrice).toUpperCase() === "TRUE" || String(r.useDbPrice) === "1"
  }));
}

function getCharterGroupByLocation(name) {
  for (const g of CHARTER_GROUPS) {
    for (const m of g.locations) {
      if (matchLocation(name, m)) return g.id;
    }
  }
  return "";
}

function loadDatabase(dbSheet) {
  const lastRow = dbSheet.getLastRow();
  const lastCol = Math.max(dbSheet.getLastColumn(), 11);
  const db = [];
  for (let r = 2; r <= lastRow; r++) {
    const row = [];
    for (let c = 1; c <= lastCol; c++) {
      const cell = dbSheet.getRange(r, c);
      row.push(c === 11 ? parseSheetDate(cell.getValue(), cell.getDisplayValue()) : cell.getValue());
    }
    db.push(row);
  }
  return db;
}

function getCandidates(location, db, dateValue) {
  const all = db.filter(r =>
    matchLocation(location, r[1]) &&
    String(r[0]).toLowerCase() !== "inactive" &&
    isEffectiveOnOrBefore(r[10], dateValue)
  );
  const exact = all.filter(r =>
    normalize(r[1]) === normalize(location) ||
    compact(r[1]) === compact(location)
  );
  return exact.length ? exact : all;
}

function fillMissingFromOlder(match, location, db, dateValue) {
  const filled = match.slice();
  const matchDate = parseSheetDate(match[10]);
  if (!matchDate) return filled;

  const older = db
    .filter(r =>
      matchLocation(location, r[1]) &&
      String(r[0]).toLowerCase() !== "inactive" &&
      isEffectiveOnOrBefore(r[10], dateValue) &&
      parseSheetDate(r[10]) &&
      sheetDateTime_(r[10]) < matchDate.getTime()
    )
    .sort((a, b) => sheetDateTime_(b[10]) - sheetDateTime_(a[10]));

  if (!older.length) return filled;

  const source = older[0];
  [2, 3, 4, 5, 6, 7, 8, 9].forEach(idx => {
    const empty =
      idx === 6
        ? !Number(String(filled[6]).replace(/[$,]/g, ""))
        : !String(filled[idx] || "").trim();
    if (empty && String(source[idx] || "").trim()) filled[idx] = source[idx];
  });
  return filled;
}

function enrichCharterDbRows(db) {
  CHARTER_GROUPS.forEach(group => {
    const rows = db.filter(r => group.locations.some(loc => matchLocation(r[1], loc)));
    if (!rows.length) return;

    const effectiveDate = rows.map(r => r[10]).find(d => parseSheetDate(d));
    const price = rows.map(r => Number(String(r[6]).replace(/[$,]/g, ""))).find(p => p > 0);
    const carrier = rows.map(r => mapCarrierName(r[3])).find(Boolean);
    const state = rows.map(r => String(r[5] || "").trim()).find(Boolean);

    rows.forEach(r => {
      if (!parseSheetDate(r[10]) && effectiveDate) r[10] = effectiveDate;
      const cur = Number(String(r[6]).replace(/[$,]/g, ""));
      if (!cur && price) r[6] = price;
      if (String(r[3] || "").trim()) r[3] = mapCarrierName(r[3]);
      else if (carrier) r[3] = carrier;
      if (!String(r[5] || "").trim() && state) r[5] = state;
    });
  });
}

function applyCharterPricing(data, charterMeta) {
  const groups = {};
  charterMeta.forEach((meta, i) => {
    if (!meta || !data[i][1]) return;
    const key = normalize(meta.group);
    if (!groups[key]) groups[key] = { price: meta.price, rows: [] };
    groups[key].rows.push(i);
  });
  Object.values(groups).forEach(({ price, rows }) => {
    rows.sort((a, b) => a - b);
    rows.forEach((idx, j) => {
      data[idx][COL.PRICE] = j === 0 ? price : "";
    });
  });
}

function fillPickupData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (cell.getColumn() !== 1) {
    SpreadsheetApp.getUi().alert("Select column A");
    return;
  }
  const dateValue = parseSheetDate(cell);
  if (!dateValue) {
    SpreadsheetApp.getUi().alert("Column A must be a date.");
    return;
  }
  const startRow = cell.getRow();
  fillConfirmCfg_(sheet, startRow, findDateBlockEnd_(sheet, startRow), dateValue);
}

function confirmToday() {
  const selected = requireEwrDateCell_();
  if (!selected) return;
  const sheet = selected.sheet;
  const dateValue = selected.date;
  const startRow = selected.cell.getRow();
  const endRow = findDateBlockEnd_(sheet, startRow);
  fillConfirmCfg_(sheet, startRow, endRow, dateValue);
  writeDayPriceTotal_(sheet, startRow, endRow);
  const afterTotal = findDateBlockEnd_(sheet, startRow);
  writeUnconfirmedPickupNotes_(sheet, startRow, afterTotal);
  writeTikTokSummary_(sheet, startRow, findDateBlockEnd_(sheet, startRow));
}

function runManualToday() {
  const selected = requireEwrDateCell_();
  if (!selected) return;
  const sheet = selected.sheet;
  const dateValue = selected.date;
  const startRow = selected.cell.getRow();
  const endRow = findDateBlockEnd_(sheet, startRow);
  fillConfirmCfg_(sheet, startRow, endRow, dateValue, { writeCarrier: true });
  writeDayPriceTotal_(sheet, startRow, endRow);
  const afterTotal = findDateBlockEnd_(sheet, startRow);
  writeUnconfirmedPickupNotes_(sheet, startRow, afterTotal);
  writeTikTokSummary_(sheet, startRow, findDateBlockEnd_(sheet, startRow));
}

function isBlankTimeOrPallet_(value) {
  if (value === "" || value === null || value === undefined) return true;
  if (typeof value === "number") return false;
  if (value instanceof Date) return isNaN(value.getTime());
  return !String(value).trim();
}

function is600NeedConfirm_(location, carrier, time, pallets, destMap) {
  if (rowReceivingDest_(location, carrier, destMap) !== "600") return false;
  return isBlankTimeOrPallet_(time) && isBlankTimeOrPallet_(pallets);
}

function loadLocationDestMap_(ss) {
  const map = {};
  loadLocalResultRows_(ss).forEach(r => {
    if (!r.official) return;
    const dest = resultDest_(r);
    const locKey = normalize(r.official);
    const both = locKey + "|" + normalize(mapCarrierName(r.carrier));
    map[both] = dest;
    if (!map[locKey]) map[locKey] = dest;
  });
  return map;
}

function rowReceivingDest_(location, carrier, destMap) {
  const c = normalize(mapCarrierName(carrier));
  if (c.indexOf("运力") !== -1 || c.indexOf("transportation") !== -1) return "600";
  const locKey = normalize(location);
  return destMap[locKey + "|" + c] || destMap[locKey] || "600";
}

function cellHasValue_(value) {
  if (value === "" || value === null || value === undefined) return false;
  if (typeof value === "string") return !!value.trim();
  if (value instanceof Date) return !isNaN(value.getTime());
  return true;
}

function findDateBlockEnd_(sheet, startRow) {
  const nextDateRow = findNextDateRow_(sheet, startRow);
  const last = nextDateRow ? nextDateRow - 1 : Math.max(sheet.getLastRow(), startRow);
  if (last <= startRow) return startRow;
  const data = sheet.getRange(startRow, 2, last - startRow + 1, 8).getValues();
  let end = startRow;
  for (let i = 0; i < data.length; i++) {
    if (data[i].some(cellHasValue_)) end = startRow + i;
  }
  return end;
}

function preferCarrierMatches_(rows, ewrCarrier) {
  const c = normalize(mapCarrierName(ewrCarrier));
  if (!c) return rows;
  const hits = rows.filter(r => normalize(mapCarrierName(r[3])) === c);
  return hits.length ? hits : rows;
}

function findNextDateRow_(sheet, startRow) {
  const last = Math.max(sheet.getLastRow(), startRow);
  const dates = sheet.getRange(startRow + 1, 1, Math.max(last - startRow, 1), 1).getValues();
  for (let i = 0; i < dates.length; i++) {
    if (parseSheetDate(dates[i][0])) return startRow + 1 + i;
  }
  return 0;
}

function unconfirmedNoteCols_() {
  return 5;
}

function clearUnconfirmedPickupNotes_(sheet, startRow, endRow) {
  const nextDateRow = findNextDateRow_(sheet, startRow);
  const clearStart = startRow + 1;
  const clearEnd = nextDateRow ? nextDateRow - 1 : Math.max(endRow, startRow);
  if (clearEnd < clearStart) return;
  sheet.getRange(clearStart, 10, clearEnd - clearStart + 1, unconfirmedNoteCols_()).clearContent();
}

function writeUnconfirmedPickupNotes_(sheet, startRow, endRow) {
  clearUnconfirmedPickupNotes_(sheet, startRow, endRow);
  const destMap = loadLocationDestMap_(SpreadsheetApp.getActiveSpreadsheet());
  const data = sheet.getRange(startRow, 1, endRow - startRow + 1, 9).getValues();
  const pending = [];
  for (let i = 0; i < data.length; i++) {
    const location = String(data[i][1] || "").trim();
    if (!location) continue;
    const carrier = String(data[i][COL.CARRIER] || "").trim();
    if (!is600NeedConfirm_(location, carrier, data[i][COL.ARRIVAL], data[i][COL.PALLETS], destMap)) continue;
    pending.push({
      row: startRow + i,
      location: location,
      carrier: carrier
    });
  }
  if (!pending.length) return;

  const carriers = [];
  const seen = {};
  pending.forEach(p => {
    const name = p.carrier || "未知承运商";
    const key = normalize(name);
    if (seen[key]) return;
    seen[key] = true;
    carriers.push(name);
  });

  const anchor = pending[0].row;
  const need = 2 + pending.length;
  const lastWrite = anchor + need - 1;
  const nextDateRow = findNextDateRow_(sheet, startRow);
  if (nextDateRow && lastWrite >= nextDateRow) {
    sheet.insertRowsBefore(nextDateRow, lastWrite - nextDateRow + 1);
  } else if (lastWrite > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), lastWrite - sheet.getMaxRows());
  }

  const cols = Math.max(carriers.length, 1);
  const rows = [];
  rows.push(["需确认是否有提货"].concat(Array(cols - 1).fill("")));
  const carrierRow = carriers.slice();
  while (carrierRow.length < cols) carrierRow.push("");
  rows.push(carrierRow);
  pending.forEach(p => {
    const row = [p.location];
    while (row.length < cols) row.push("");
    rows.push(row);
  });
  sheet.getRange(anchor, 10, rows.length, cols).setValues(rows);
  sheet.getRange(anchor, 10, rows.length, cols).setFontWeight("normal");
  sheet.getRange(anchor, 10).setFontWeight("bold");
  sheet.getRange(anchor + 1, 10, 1, cols).setFontWeight("bold");
}

function fillConfirmCfg_(sheet, startRow, endRow, dateValue, options) {
  const opts = options || {};
  const weekday = Utilities.formatDate(dateValue, getSpreadsheetTz(), "EEE");
  const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 12);
  const data = range.getValues();
  const charterMeta = new Array(data.length).fill(null);
  const db = loadDatabase(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database"));
  enrichCharterDbRows(db);
  const rules = loadRules_();

  for (let i = 0; i < data.length; i++) {
    const location = String(data[i][1] || "").trim();
    if (!location) continue;

    const ewrCarrier = mapCarrierName(data[i][COL.CARRIER]);
    const candidates = getCandidates(location, db, dateValue);
    if (!candidates.length) {
      data[i][COL.ADDRESS] = "";
      data[i][COL.STATE] = "";
      data[i][COL.PRICE] = "";
      continue;
    }

    const dated = candidates.filter(r => parseSheetDate(r[10]));
    const pool = preferCarrierMatches_(dated.length ? dated : candidates, ewrCarrier);
    const maxDate = Math.max(...pool.map(r => sheetDateTime_(r[10])));
    const latest = preferCarrierMatches_(
      pool.filter(r => sheetDateTime_(r[10]) === maxDate),
      ewrCarrier
    );
    const isCharterLoc = latest.some(r => getCharterGroupByLocation(r[1]));

    let match = null;
    let hasSchedule = false;

    for (const r of latest) {
      const raw = String(r[7] || "").trim();
      if (raw) hasSchedule = true;
      const days = raw.replace(/,/g, " ").replace(/\s+/g, " ").trim().split(" ");
      if (!raw || days.includes(weekday)) {
        match = r;
        break;
      }
    }

    if (!match) {
      if (!hasSchedule || isCharterLoc) {
        match = latest[0];
      } else {
        data[i][COL.ADDRESS] = "";
        data[i][COL.STATE] = "";
        data[i][COL.PRICE] = "";
        continue;
      }
    }

    match = fillMissingFromOlder(match, location, db, dateValue);

    const charterGroup = getCharterGroupByLocation(match[1]);
    const isCharter = !!charterGroup;
    const basePrice = Number(String(match[6]).replace(/[$,]/g, ""));

    let matchedRule = rules.find(r => r.type === "carrier" && normalize(mapCarrierName(r.value)) === normalize(ewrCarrier));
    if (!matchedRule) {
      matchedRule = rules.find(r => r.type === "location" && r.value === normalize(location));
    }
    data[i][COL.ADDRESS] = match[2];
    data[i][COL.STATE] = match[5];
    if (opts.writeCarrier) {
      const dbCarrier = mapCarrierName(match[3]);
      if (dbCarrier) data[i][COL.CARRIER] = dbCarrier;
    }

    if (isCharter) {
      charterMeta[i] = { group: charterGroup, price: basePrice };
      data[i][COL.PRICE] = "";
    } else {
      let price = basePrice;
      const truck = String(data[i][COL.TRUCK]).trim();
      if (!(matchedRule && matchedRule.useDbPrice) && ["53", "53'", "53'", "53'"].includes(truck)) {
        price *= 2;
      }
      data[i][COL.PRICE] = price;
    }
  }

  applyCharterPricing(data, charterMeta);
  const n = data.length;
  sheet.getRange(startRow, 3, n, 1).setValues(data.map(r => [r[COL.ADDRESS]]));
  if (opts.writeCarrier) {
    sheet.getRange(startRow, 4, n, 1).setValues(data.map(r => [r[COL.CARRIER]]));
  }
  sheet.getRange(startRow, 6, n, 1).setValues(data.map(r => [r[COL.STATE]]));
  sheet.getRange(startRow, 7, n, 1).setValues(data.map(r => [r[COL.PRICE]]));
}

function isTikTokMerchant(name) {
  return TIKTOK_MERCHANTS.some(m => matchLocation(name, m) || matchLocation(m, name));
}

function defaultTruckType(value) {
  const raw = String(value || "").replace(/[\u2018\u2019\u2032]/g, "'").trim();
  if (!raw) return "26'";
  const n = raw.replace(/'/g, "");
  if (n === "26" || n === "53") return n + "'";
  return raw.endsWith("'") ? raw : raw + "'";
}

function formatPriceTag(value) {
  const n = Number(String(value || "").replace(/[$,]/g, ""));
  if (!isFinite(n)) return "";
  return "$" + (Number.isInteger(n) ? n : n);
}

function palletNumber(value) {
  const s = String(value || "").replace(/[$,]/g, "").trim();
  if (!s) return 0;
  if (/[+＋]/.test(s)) {
    return s.split(/[+＋]/).reduce((sum, part) => {
      const n = Number(String(part || "").trim());
      return sum + (isFinite(n) ? n : 0);
    }, 0);
  }
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function priceAmount_(value) {
  const n = Number(String(value || "").replace(/[$,]/g, "").trim());
  return isFinite(n) ? n : 0;
}

function writeDayPriceTotal_(sheet, startRow, endRow) {
  const n = endRow - startRow + 1;
  const locs = sheet.getRange(startRow, 2, n, 1).getValues();
  const prices = sheet.getRange(startRow, 7, n, 1).getValues();
  let lastB = startRow;
  let total = 0;
  for (let i = 0; i < n; i++) {
    if (!String(locs[i][0] || "").trim()) continue;
    lastB = startRow + i;
    total += priceAmount_(prices[i][0]);
  }
  const nextDateRow = findNextDateRow_(sheet, startRow);
  let totalRow = lastB + 1;
  if (nextDateRow && totalRow === nextDateRow) {
    sheet.insertRowsBefore(nextDateRow, 1);
  } else if (totalRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  } else if (String(sheet.getRange(totalRow, 2).getValue() || "").trim()) {
    sheet.insertRowsAfter(lastB, 1);
    totalRow = lastB + 1;
  }
  sheet.getRange(totalRow, 7).setValue(total);
}

function writeTikTokSummary_(sheet, startRow, endRow) {
  const data = sheet.getRange(startRow, 1, endRow - startRow + 1, 12).getValues();
  applyTikTokSummary(data);
  sheet.getRange(startRow, 10, 1, 3).setValues([[
    data[0][COL.TIKTOK_LABEL],
    data[0][COL.TIKTOK_PRICE],
    data[0][COL.TIKTOK_TRUCK]
  ]]);
}

function applyTikTokSummary(data) {
  data[0][COL.TIKTOK_LABEL] = "";
  data[0][COL.TIKTOK_PRICE] = "";
  data[0][COL.TIKTOK_TRUCK] = "";

  const priceCount = {};
  const extraOrder = [];
  const truckParts = [];

  data.forEach(row => {
    if (!isTikTokMerchant(row[1])) return;

    const truck = defaultTruckType(row[COL.TRUCK]);
    truckParts.push(truck + "-" + palletNumber(row[COL.PALLETS]) + "plts");

    const tag = formatPriceTag(row[COL.PRICE]);
    if (!tag) return;
    if (!priceCount[tag]) {
      priceCount[tag] = 0;
      const amount = Number(String(tag).replace(/[$,]/g, ""));
      if (TIKTOK_PRICE_ORDER.indexOf(amount) === -1) extraOrder.push(tag);
    }
    priceCount[tag] += 1;
  });

  if (!truckParts.length) return false;

  const orderedTags = TIKTOK_PRICE_ORDER
    .map(n => "$" + n)
    .filter(tag => priceCount[tag])
    .concat(extraOrder);

  data[0][COL.TIKTOK_LABEL] = "TikTok Inc.";
  data[0][COL.TIKTOK_PRICE] = orderedTags.map(tag => tag + "-" + priceCount[tag]).join("/");
  data[0][COL.TIKTOK_TRUCK] = truckParts.join("+");
  return true;
}

const ANALYSIS_SHEET = "results";
const DAILY_PASTE_START_ROW = 3;
const DAILY_PASTE_HEADER_ROW = 2;
const DAILY_PASTE_PLACEHOLDER = "粘贴在此处";
const ANALYSIS_SUMMARY_START_ROW = 3;
const ANALYSIS_DETAIL_TITLE_ROW = 10;
const ANALYSIS_DETAIL_HEADER_ROW = 11;
const ANALYSIS_DETAIL_START_ROW = 12;
const ANALYSIS_DETAIL_COLS = 7;
const DAILY_CARRIERS = [
  { key: "han", name: "Han Express", col: 1, color: "#D0E2FF" },
  { key: "nyqz", name: "NYQZ", col: 2, color: "#FFF2CC" },
  { key: "fm", name: "FM", col: 3, color: "#D9EAD3" },
  { key: "sd", name: "SD", col: 4, color: "#FCE5CD" }
];
const HUB_COMPACT = {
  nj600: true,
  ewr600: true,
  nj936: true,
  ewr936: true,
  nj100: true,
  ewr100: true,
  jfk: true,
  ewr: true
};

function findDailyListSheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = normalize(sheets[i].getName());
    if (n === "daily list" || n === "daily sheet") return sheets[i];
  }
  return null;
}

function isDailyListSheet_(sheet) {
  const n = normalize(sheet && sheet.getName());
  return n === "daily list" || n === "daily sheet";
}

function isEwrPickupSheet_(sheet) {
  const n = normalize(sheet && sheet.getName());
  return n.indexOf("ewr") !== -1 && n.indexOf("提货") !== -1;
}

function requireDailyListDateCell_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (!isDailyListSheet_(sheet) || cell.getRow() !== 1 || cell.getColumn() > 4) {
    SpreadsheetApp.getUi().alert("先选中 daily list 的 A1:D1 日期格。");
    return null;
  }
  const date = parseSheetDate(cell);
  if (!date) {
    SpreadsheetApp.getUi().alert("A1:D1 必须是日期。");
    return null;
  }
  return { sheet: sheet, date: date };
}

function requireEwrDateCell_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (!isEwrPickupSheet_(sheet) || cell.getColumn() !== 1) {
    SpreadsheetApp.getUi().alert("先选中 EWR本地提货单量 A 列的当日日期。");
    return null;
  }
  const date = parseSheetDate(cell);
  if (!date) {
    SpreadsheetApp.getUi().alert("A 列必须是日期。");
    return null;
  }
  return { sheet: sheet, cell: cell, date: date };
}

function findAnalysisSheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = normalize(sheets[i].getName());
    if (n === "results" || n === "分析结果") return sheets[i];
  }
  return null;
}

function ensureAnalysisSheet_(ss) {
  let sheet = findAnalysisSheet_(ss);
  if (sheet) {
    if (sheet.getName() !== ANALYSIS_SHEET) sheet.setName(ANALYSIS_SHEET);
    sheet.getRange("A1").setValue(resultsTitle_(null));
    return sheet;
  }
  sheet = ss.insertSheet(ANALYSIS_SHEET);
  sheet.getRange("A1").setValue(resultsTitle_(null));
  sheet.getRange("A2:D2").setValues([["承运商", "类型", "车次", "站点"]]);
  sheet.getRange("A10").setValue("站点明细");
  sheet.getRange("A11:G11").setValues([[
    "Carrier", "Kind", "Raw station", "Official location", "Truck", "Dest", "Job #"
  ]]);
  formatAnalysisSheet_(sheet, 1);
  return sheet;
}

function resultsTitle_(date) {
  const d = date || todayLocalDate_();
  return "Results " + Utilities.formatDate(d, getSpreadsheetTz(), "yyyy-MM-dd");
}

function formatAnalysisSheet_(sheet, summaryCount) {
  const last = Math.max(sheet.getLastRow(), ANALYSIS_DETAIL_START_ROW);
  sheet.getRange(1, 1, last, ANALYSIS_DETAIL_COLS).setHorizontalAlignment("left");
  const extra = Math.max(summaryCount - 1, 1);
  sheet.getRange(4, 3, extra, 2).setBackground(null);
  sheet.getRange("C2:D3").setBackground("#FFF2CC");
  sheet.getRange("A1:D3").setFontWeight("bold");
  sheet.getRange("A10:G11").setFontWeight("bold");
  sheet.autoResizeColumn(3);
  sheet.autoResizeColumn(4);
}

function todayLocalDate_() {
  const str = Utilities.formatDate(new Date(), getSpreadsheetTz(), "yyyy-MM-dd");
  const parts = str.split("-");
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function yesterdayLocalDate_() {
  const today = todayLocalDate_();
  return new Date(today.getFullYear(), today.getMonth(), today.getDate() - 1);
}

function writeDailyListDate_(sheet, date) {
  const cell = sheet.getRange("A1");
  cell.setValue(date);
  cell.setNumberFormat("yyyy-mm-dd");
}

function dailyListDate_(sheet) {
  return parseSheetDate(sheet.getRange("A1"));
}

function dateFromListText_(text) {
  const lines = String(text || "").replace(/\r/g, "").split("\n");
  for (let i = 0; i < lines.length; i++) {
    const d = parseDateString(String(lines[i] || "").trim());
    if (d) return d;
  }
  return null;
}

function dateFromPastedLists_(sheet) {
  for (let i = 0; i < DAILY_CARRIERS.length; i++) {
    const d = dateFromListText_(readPasteColumn_(sheet, DAILY_CARRIERS[i].col, DAILY_PASTE_START_ROW));
    if (d) return d;
  }
  return null;
}

function pastePlaceholderForCol_(col) {
  let name = "";
  DAILY_CARRIERS.forEach(c => {
    if (c.col === col) name = c.name;
  });
  return "粘贴 (" + name + ") 列表在此处";
}

function isPastePlaceholder_(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (normalize(raw) === normalize(DAILY_PASTE_PLACEHOLDER)) return true;
  return /^粘贴\s*[（(].+[）)]\s*列表在此处$/.test(raw);
}

function stripPastePlaceholder_(value) {
  return String(value || "")
    .replace(/粘贴\s*[（(][^）)]+[）)]\s*列表在此处/g, "")
    .replace(/粘贴在此处/g, "")
    .replace(/^\s+|\s+$/g, "");
}

function showPastePlaceholders_(sheet) {
  for (let c = 1; c <= 4; c++) {
    setPastePlaceholderCell_(sheet.getRange(DAILY_PASTE_START_ROW, c), c);
  }
}

function sameSheetDate_(a, b) {
  const left = parseSheetDate(a);
  const right = parseSheetDate(b);
  return !!(left && right && left.getTime() === right.getTime());
}

function rolloverDailyListIfNewDay_() {
  const sheet = findDailyListSheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet) return;
  const listDate = yesterdayLocalDate_();
  if (sameSheetDate_(dailyListDate_(sheet), listDate)) {
    restoreEmptyPastePlaceholders_();
    return;
  }
  writeDailyListDate_(sheet, listDate);
  sheet.getRange(DAILY_PASTE_START_ROW, 1, 1, 4).clearContent();
  showPastePlaceholders_(sheet);
}

function restoreEmptyPastePlaceholders_() {
  const sheet = findDailyListSheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet) return;
  for (let c = 1; c <= 4; c++) {
    applyPasteCellState_(sheet.getRange(DAILY_PASTE_START_ROW, c), c);
  }
}

function restorePastePlaceholderOnEdit_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const n = normalize(sheet.getName());
  if (n !== "daily list" && n !== "daily sheet") return;
  const startRow = e.range.getRow();
  const endRow = startRow + e.range.getNumRows() - 1;
  const startCol = e.range.getColumn();
  const endCol = startCol + e.range.getNumColumns() - 1;
  if (endRow < DAILY_PASTE_START_ROW || startRow > DAILY_PASTE_START_ROW) return;
  if (endCol < 1 || startCol > 4) return;
  for (let c = Math.max(1, startCol); c <= Math.min(4, endCol); c++) {
    applyPasteCellState_(sheet.getRange(DAILY_PASTE_START_ROW, c), c);
  }
}

function applyPasteCellState_(cell, col) {
  const cleaned = stripPastePlaceholder_(cell.getValue());
  if (!cleaned) {
    setPastePlaceholderCell_(cell, col);
    return;
  }
  if (cleaned !== String(cell.getValue() || "")) cell.setValue(cleaned);
  cell.setFontColor("#000000").setFontStyle("normal");
}

function setPastePlaceholderCell_(cell, col) {
  cell.setValue(pastePlaceholderForCol_(col)).setFontColor("#B7B7B7").setFontStyle("italic");
}

function clearDailyListContent_(sheet) {
  const paste = sheet.getRange(DAILY_PASTE_START_ROW, 1, 1, 4);
  paste.clearContent();
  paste.setFontColor("#B7B7B7").setFontStyle("italic");
  showPastePlaceholders_(sheet);
  const analysis = findAnalysisSheet_(SpreadsheetApp.getActiveSpreadsheet());
  if (analysis) clearAnalysisResults_(analysis);
}

function clearAnalysisResults_(sheet) {
  const summaryRows = ANALYSIS_DETAIL_TITLE_ROW - ANALYSIS_SUMMARY_START_ROW;
  if (summaryRows > 0) {
    sheet.getRange(ANALYSIS_SUMMARY_START_ROW, 1, summaryRows, 4).clearContent();
  }
  const last = Math.max(sheet.getLastRow(), ANALYSIS_DETAIL_START_ROW);
  const range = sheet.getRange(ANALYSIS_DETAIL_START_ROW, 1, last - ANALYSIS_DETAIL_START_ROW + 1, ANALYSIS_DETAIL_COLS);
  range.clearContent();
  range.clearDataValidations();
}

function parseDailyCarrierLists() {
  const selected = requireDailyListDateCell_();
  if (!selected) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = selected.sheet;
  const dbNames = loadDatabaseLocationNames_(ss);
  const parsed = [];
  const summary = [];
  let trucks = 0;

  DAILY_CARRIERS.forEach(c => {
    const text = readPasteColumn_(sheet, c.col, DAILY_PASTE_START_ROW);
    const result = parseCarrierText_(text, c.key);
    trucks += result.trucks;
    result.stations.forEach(st => {
      const kind = st.kind === "customs" ? "清关行" : "本地";
      parsed.push([
        c.name,
        kind,
        st.raw,
        kind === "清关行" ? "" : findOfficialLocation_(st.raw, dbNames, c.key),
        st.truck,
        st.dest,
        st.job
      ]);
    });
    if (c.key === "han") {
      const local = result.stations.filter(st => st.kind !== "customs");
      const customs = result.stations.filter(st => st.kind === "customs");
      const customsJobs = {};
      result.stations.forEach(st => {
        if (st.kind !== "customs") return;
        customsJobs[st.dest + "|" + st.job] = true;
      });
      let customsTrucks = 0;
      Object.keys(customsJobs).forEach(job => {
        customsTrucks += String(job).split("+").filter(Boolean).length;
      });
      const localTrucks = result.trucks - customsTrucks;
      if (local.length) summary.push([c.name, "本地", localTrucks, local.length]);
      if (customs.length) summary.push([c.name, "清关行", customsTrucks, customs.length]);
    } else if (result.stations.length) {
      summary.push([c.name, "本地", result.trucks, result.stations.length]);
    }
  });
  summary.unshift(["合计", "", trucks, parsed.length]);

  const analysis = ensureAnalysisSheet_(ss);
  const listDate = selected.date;
  writeDailyListDate_(sheet, listDate);
  analysis.getRange("A1").setValue(resultsTitle_(listDate));
  clearAnalysisResults_(analysis);
  analysis.getRange(ANALYSIS_SUMMARY_START_ROW, 1, summary.length, 4).setValues(summary);
  if (parsed.length) {
    analysis.getRange(ANALYSIS_DETAIL_START_ROW, 1, parsed.length, ANALYSIS_DETAIL_COLS).setValues(parsed);
    const rule = locationDropdownRule_(ss);
    parsed.forEach((row, i) => {
      const cell = analysis.getRange(ANALYSIS_DETAIL_START_ROW + i, 4);
      if (row[1] === "本地" && rule) cell.setDataValidation(rule);
      else cell.clearDataValidations();
    });
  }
  formatAnalysisSheet_(analysis, summary.length);
  ss.setActiveSheet(analysis);

  SpreadsheetApp.getUi().alert("Parsed " + trucks + " trucks, " + parsed.length + " stations. Results " + Utilities.formatDate(listDate, getSpreadsheetTz(), "yyyy-MM-dd") + " is ready.");
}

function fillToday() {
  const selected = requireEwrDateCell_();
  if (!selected) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = selected.sheet;
  const date = selected.date;

  const localRows = loadLocalResultRows_(ss);
  if (!localRows.length) {
    SpreadsheetApp.getUi().alert("Parse the 4 lists first.");
    return;
  }
  const startRow = selected.cell.getRow();
  const dbNames = loadDatabaseLocationNames_(ss);
  const dropNames = loadColoredDropdownNames_(sheet, startRow);
  const resolved = localRows.map(r => {
    const official = pickAllowedName_(r.official, dropNames)
      || pickAllowedName_(r.raw, dropNames)
      || pickAllowedName_(r.official, dbNames)
      || r.official
      || r.raw;
    return {
      carrier: r.carrier,
      official: official,
      truck: r.truck,
      dest: r.dest
    };
  }).filter(r => r.official);
  const fillRows = buildFillRows_(resolved, loadTransportationLocations_(ss));
  if (!fillRows.length) {
    SpreadsheetApp.getUi().alert("No local official locations to fill.");
    return;
  }

  ensureBlockRows_(sheet, startRow, fillRows.length);
  const noon = new Date(date.getFullYear(), date.getMonth(), date.getDate(), 12, 0, 0);
  const range = sheet.getRange(startRow, 1, fillRows.length, 8);
  const bRange = sheet.getRange(startRow, 2, fillRows.length, 1);
  bRange.clearDataValidations();
  const data = range.getValues();
  for (let i = 0; i < fillRows.length; i++) {
    data[i][0] = i === 0 ? date : "";
    data[i][1] = fillRows[i].official;
    data[i][COL.CARRIER] = fillRows[i].carrier;
    data[i][COL.TRUCK] = fillRows[i].truck;
    data[i][COL.ARRIVAL] = fillRows[i].needTime ? noon : "";
  }
  range.setValues(data);
  sheet.getRange(startRow, 1).setNumberFormat("yyyy-mm-dd");
  applyColoredEwrDropdown_(sheet, startRow, fillRows.length);
  sheet.getRange(startRow, 8, fillRows.length, 1)
    .setHorizontalAlignment("center")
    .setNumberFormat("yyyy-mm-dd hh:mm:ss");
}

function loadColoredDropdownNames_(sheet, beforeRow) {
  const cell = findExistingEwrDropdownCell_(sheet, beforeRow);
  return dropdownNamesFromRule_(cell && cell.getDataValidation());
}

function pickAllowedName_(raw, names) {
  const value = String(raw || "").trim();
  const list = names || [];
  if (!value || !list.length) return "";
  for (let i = 0; i < list.length; i++) {
    if (normalize(list[i]) === normalize(value) || compact(list[i]) === compact(value)) return list[i];
  }
  const aliases = stationAliasKeys_(value);
  const hits = list.filter(n => {
    const keys = stationAliasKeys_(n);
    return aliases.some(a => keys.indexOf(a) !== -1);
  });
  return hits.length ? hits[0] : "";
}

function ewrLocationName_(official, allowed) {
  return pickAllowedName_(official, allowed) || String(official || "").trim();
}

function findEwrSheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = normalize(sheets[i].getName());
    if (n.indexOf("ewr") !== -1 && n.indexOf("提货") !== -1) return sheets[i];
  }
  return ss.getActiveSheet();
}

function findExistingEwrDropdownCell_(sheet, beforeRow) {
  const probe = Math.min(beforeRow - 1, 10175);
  if (probe < 2) return sheet.getRange(2, 2);
  const lookback = Math.min(probe - 1, 400);
  const start = probe - lookback;
  const rules = sheet.getRange(start, 2, lookback + 1, 1).getDataValidations();
  for (let i = rules.length - 1; i >= 0; i--) {
    if (rules[i][0]) return sheet.getRange(start + i, 2);
  }
  return sheet.getRange(2, 2);
}

function applyColoredEwrDropdown_(sheet, startRow, numRows) {
  const src = findExistingEwrDropdownCell_(sheet, startRow);
  if (!src || !src.getDataValidation()) return;
  src.copyTo(sheet.getRange(startRow, 2, numRows, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}

function dropdownNamesFromRule_(rule) {
  if (!rule) return [];
  const type = rule.getCriteriaType();
  const vals = rule.getCriteriaValues();
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return (vals[0] || []).map(v => String(v || "").trim()).filter(Boolean);
  }
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && vals[0]) {
    const seen = {};
    const names = [];
    vals[0].getValues().forEach(r => {
      const name = String(r[0] || "").trim();
      if (!name || seen[normalize(name)]) return;
      seen[normalize(name)] = true;
      names.push(name);
    });
    return names;
  }
  return [];
}

function loadExistingEwrLocationNames_(sheet, beforeRow) {
  const end = Math.max(2, Math.min(beforeRow - 1, sheet.getLastRow()));
  if (end < 2) return [];
  const values = sheet.getRange(2, 2, end - 1, 1).getValues();
  const names = [];
  const seen = {};
  values.forEach(r => {
    const name = String(r[0] || "").trim();
    if (!name || seen[normalize(name)]) return;
    seen[normalize(name)] = true;
    names.push(name);
  });
  return names;
}

function ewrLocationDropdownRule_(ss) {
  const range = databaseLocationRange_(ss);
  if (!range) return null;
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(true)
    .build();
}

function loadLocalResultRows_(ss) {
  const sheet = findAnalysisSheet_(ss);
  if (!sheet || sheet.getLastRow() < ANALYSIS_DETAIL_START_ROW) return [];
  const n = sheet.getLastRow() - ANALYSIS_DETAIL_START_ROW + 1;
  return sheet.getRange(ANALYSIS_DETAIL_START_ROW, 1, n, ANALYSIS_DETAIL_COLS).getValues()
    .map(r => ({
      carrier: String(r[0] || "").trim(),
      kind: String(r[1] || "").trim(),
      raw: String(r[2] || "").trim(),
      official: String(r[3] || "").trim(),
      truck: String(r[4] || "").trim(),
      dest: String(r[5] || "").trim()
    }))
    .filter(r => r.kind === "本地");
}

function resultDest_(row) {
  const d = compact(row.dest);
  if (d === "jfk") return "JFK";
  if (d === "936" || d === "100") return "936";
  return "600";
}

function resultCarrierKey_(name) {
  const n = normalize(mapCarrierName(name));
  if (n === "han express" || n === "80s") return "han";
  if (n === "sd") return "sd";
  if (n === "fm") return "fm";
  if (n === "nyqz") return "nyqz";
  return "";
}

function groupSameOfficial_(rows) {
  const order = [];
  const buckets = {};
  (rows || []).forEach(r => {
    const key = normalize(r.official);
    if (!key) return;
    if (!buckets[key]) {
      buckets[key] = [];
      order.push(key);
    }
    buckets[key].push(r);
  });
  const out = [];
  order.forEach(key => {
    buckets[key].forEach(r => out.push(r));
  });
  return out;
}

function buildFillRows_(localRows, transportNames) {
  const out = [];
  const placed = [];
  function add(official, truck, needTime, carrier) {
    if (!String(official || "").trim()) return;
    out.push({
      official: official,
      truck: truck || "",
      needTime: needTime,
      carrier: carrier || ""
    });
  }
  function take(rows, needTime) {
    groupSameOfficial_(rows).forEach(r => {
      placed.push(r);
      add(r.official, r.truck, needTime, r.carrier);
    });
  }
  take(localRows.filter(r => resultDest_(r) === "JFK"), false);
  take(localRows.filter(r => resultDest_(r) === "936"), false);
  ["han", "sd", "fm", "nyqz"].forEach(key => {
    take(localRows.filter(r => resultDest_(r) === "600" && resultCarrierKey_(r.carrier) === key), true);
  });
  take(localRows.filter(r => resultDest_(r) === "600" && placed.indexOf(r) === -1), true);
  const seen = {};
  (transportNames || []).forEach(name => {
    const key = normalize(name);
    if (!key || seen[key]) return;
    seen[key] = true;
    add(name, "", true, "运力卡车");
  });
  return out;
}

function loadTransportationLocations_(ss) {
  const dbSheet = ss.getSheetByName("database");
  if (!dbSheet || dbSheet.getLastRow() < 2) return [];
  const n = dbSheet.getLastRow() - 1;
  const status = dbSheet.getRange(2, 1, n, 1).getValues();
  const names = dbSheet.getRange(2, 2, n, 1).getValues();
  const carriers = dbSheet.getRange(2, 4, n, 1).getValues();
  const out = [];
  const seen = {};
  for (let i = 0; i < n; i++) {
    if (String(status[i][0]).toLowerCase() === "inactive") continue;
    const carrier = normalize(mapCarrierName(carriers[i][0]));
    if (carrier.indexOf("运力") === -1 && carrier.indexOf("transportation") === -1) continue;
    const name = String(names[i][0] || "").trim();
    if (!name || seen[normalize(name)]) continue;
    seen[normalize(name)] = true;
    out.push(name);
  }
  return out;
}

function ensureBlockRows_(sheet, startRow, need) {
  const last = Math.max(sheet.getLastRow(), startRow);
  let nextDateRow = 0;
  for (let r = startRow + 1; r <= last; r++) {
    if (parseSheetDate(sheet.getRange(r, 1).getValue())) {
      nextDateRow = r;
      break;
    }
  }
  const available = nextDateRow ? nextDateRow - startRow : need;
  if (available < need) sheet.insertRowsAfter(startRow, need - available);
}

function localOfficialDropdownRule_(ss, names) {
  const uniq = [];
  const seen = {};
  (names || []).forEach(n => {
    const v = String(n || "").trim();
    if (!v || seen[normalize(v)]) return;
    seen[normalize(v)] = true;
    uniq.push(v);
  });
  if (!uniq.length) return null;
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(uniq, true)
    .setAllowInvalid(true)
    .build();
}

function ensureSheet_(ss, name, rows, cols) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const needRows = rows - sheet.getMaxRows();
  const needCols = cols - sheet.getMaxColumns();
  if (needRows > 0) sheet.insertRowsAfter(sheet.getMaxRows(), needRows);
  if (needCols > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols);
  return sheet;
}

function readPasteColumn_(sheet, col, startRow) {
  const cleaned = stripPastePlaceholder_(sheet.getRange(startRow, col).getValue());
  if (!cleaned || isPastePlaceholder_(cleaned)) return "";
  return cleaned;
}

function databaseLocationRange_(ss) {
  const dbSheet = ss.getSheetByName("database");
  if (!dbSheet || dbSheet.getLastRow() < 2) return null;
  return dbSheet.getRange(2, 2, dbSheet.getLastRow() - 1, 1);
}

function loadDatabaseLocationNames_(ss) {
  const range = databaseLocationRange_(ss);
  if (!range) return [];
  const seen = {};
  const names = [];
  range.getValues().forEach(r => {
    const name = String(r[0] || "").trim();
    if (!name || seen[name]) return;
    seen[name] = true;
    names.push(name);
  });
  return names;
}

function locationDropdownRule_(ss) {
  const range = databaseLocationRange_(ss);
  if (!range) return null;
  return SpreadsheetApp.newDataValidation()
    .requireValueInRange(range, true)
    .setAllowInvalid(false)
    .build();
}

function findOfficialLocation_(raw, dbNames, carrierKey) {
  const names = dbNames || [];
  if (carrierKey === "fm") {
    if (isSfNjStation_(raw) || /^(sf-fm|sffm)$/.test(compact(raw))) {
      const fm = pickOfficialName_(names, ["sfnj122", "sf-fm", "sffm"]);
      if (fm) return fm;
    }
    if (isYanwenStation_(raw)) {
      const fm = pickOfficialName_(names, ["yanwen-fm", "yanwenfm"]);
      if (fm) return fm;
    }
  }
  if (carrierKey === "han") {
    if (isSfNjStation_(raw)) {
      const han = pickOfficialName_(names, ["sfnj120"]);
      if (han) return han;
    }
    if (isYanwenStation_(raw)) {
      const han = pickOfficialName_(names, ["yanwen-han", "yanwenhan"]);
      if (han) return han;
    }
  }
  if (carrierKey === "sd") {
    const cap = pickCapacityOfficial_(raw, names);
    if (cap) return cap;
  }

  const exactName = names.find(n => normalize(n) === normalize(raw) || compact(n) === compact(raw));
  if (exactName) return exactName;

  const aliases = stationAliasKeys_(raw);
  const exact = names.filter(n => {
    const keys = stationAliasKeys_(n);
    return aliases.some(a => keys.indexOf(a) !== -1);
  });
  if (exact.length === 1) return exact[0];
  const hits = names.filter(n => matchLocation(raw, n));
  if (hits.length === 1) return hits[0];
  if (exact.length) return exact[0];
  return "";
}

function pickCapacityOfficial_(raw, names) {
  const c = compact(raw).replace(/^capacity/, "");
  const keys = {
    nj1101: ["capacitynj1101"],
    nj1112: ["capacitynj1112"],
    nj1000: ["capacitynj1000"],
    nj1600: ["capacitynj1600"],
    nj1980: ["capacitynj1980"]
  }[c];
  if (!keys) return "";
  return pickOfficialName_(names, keys);
}

function pickOfficialName_(dbNames, keys) {
  for (let i = 0; i < dbNames.length; i++) {
    const aliases = stationAliasKeys_(dbNames[i]);
    if (keys.some(k => aliases.indexOf(k) !== -1)) return dbNames[i];
  }
  return "";
}

function isSfNjStation_(name) {
  return /^sfnj\d+$/.test(compact(name));
}

function isYanwenStation_(name) {
  return compact(name).indexOf("yanwen") !== -1;
}

function stationAliasKeys_(name) {
  const c = compact(name);
  if (!c) return [];
  const groups = [
    ["nj1001", "njtt1001"],
    ["njtt650", "swiftxnj650", "swiftx"],
    ["pattern4711", "patternpa4711"],
    ["de901c", "虫洞de901c"],
    ["pa3363", "西邮2pa3363"],
    ["thgnj300", "thgingenuity"],
    ["brandfoxpa400", "brandfox400"],
    ["staci10", "stacinj10"],
    ["gxonj25", "gxo-kendo", "gxokendo"],
    ["iherbnj1540", "iherb-nj1540", "iherb-epa", "iherbepa"],
    ["sfnj122", "sf-fm", "sffm"],
    ["nj250", "ocnj250"],
    ["nj1401", "ocnj1401", "cainiaonj1401"],
    ["shipbob", "shipbob4779"],
    ["stockxnj1", "stockx"],
    ["quickboxnj415", "quickboxfulfillment-nj", "quickboxfulfillmentnj"]
  ];
  for (let i = 0; i < groups.length; i++) {
    if (groups[i].indexOf(c) !== -1) return groups[i].slice();
  }
  return [c];
}

function isHubStation_(name) {
  return !!HUB_COMPACT[compact(name)];
}

function cleanStationToken_(token) {
  return String(token || "")
    .replace(/[（(][^）)]*[）)]/g, "")
    .replace(/[✅❌🚗❗🌟]/g, "")
    .replace(/UNITRANS\d+/ig, "")
    .replace(/\s+/g, " ")
    .trim();
}

function splitRealStations_(route) {
  const s = String(route || "");
  const parts = [];
  let buf = "";
  let depth = 0;
  for (let i = 0; i < s.length; i++) {
    const ch = s.charAt(i);
    if (ch === "(" || ch === "（") depth++;
    if ((ch === "-" || ch === "–" || ch === "—") && depth <= 0) {
      const t = cleanStationToken_(buf);
      if (t) parts.push(t);
      buf = "";
      continue;
    }
    if (ch === ")" || ch === "）") depth = Math.max(0, depth - 1);
    buf += ch;
  }
  const last = cleanStationToken_(buf);
  if (last) parts.push(last);
  return parts.filter(p => p && !isHubStation_(p));
}

function parseJobLine_(line) {
  let s = String(line || "").trim();
  if (!s) return null;
  const cancelled = /❌/.test(s);
  const transfer100 = /转\s*100/.test(s);
  s = s.replace(/[✅❌🚗🌟]/g, "")
    .replace(/转\s*100/g, "")
    .replace(/没货/g, "")
    .trim();
  const m = s.match(/^(\d+(?:\s*[+＋]\s*\d+)*)\s*[.\u3001、:：]\s*(.+)$/);
  if (!m) return null;
  const nums = m[1].split(/[+＋]/).map(x => x.replace(/\s+/g, "")).filter(Boolean);
  return { cancelled, transfer100, nums, route: m[2].trim() };
}

function classifySection_(header, carrierKey) {
  const h = compact(header);
  if (carrierKey === "han") {
    if (header.indexOf("转运") === -1 && (header.indexOf("本地提货") !== -1 || (header.indexOf("提货") !== -1 && /(600|936|100|jfk)/i.test(h)))) {
      let dest = "600";
      if (/jfk/i.test(header) && h.indexOf("600") === -1 && h.indexOf("936") === -1 && h.indexOf("100") === -1) dest = "JFK";
      else if (h.indexOf("936") !== -1 || (h.indexOf("100") !== -1 && h.indexOf("600") === -1)) dest = "936";
      return { kind: "local", dest: dest };
    }
    if (header.indexOf("转运") !== -1 && /jfk/i.test(header) && (h.indexOf("600") !== -1 || h.indexOf("100") !== -1)) {
      return { kind: "customs", dest: h.indexOf("100") !== -1 && h.indexOf("600") === -1 ? "100" : "600" };
    }
    return { ignore: true };
  }
  if (header.indexOf("本地提货") !== -1) return { kind: "local", dest: "600" };
  return { ignore: true };
}

function localTruckType_(nums) {
  return nums.length >= 2 ? "53'" : "26'";
}

function parseCarrierText_(text, carrierKey) {
  const lines = String(text || "").replace(/\r/g, "").split("\n");
  let section = { ignore: true };
  const stations = [];
  let trucks = 0;

  lines.forEach(line => {
    const trimmed = String(line || "").trim();
    if (!trimmed) return;
    if (/^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/.test(trimmed)) return;
    if (/【.+】/.test(trimmed) || /本地提货|转运/.test(trimmed)) {
      section = classifySection_(trimmed, carrierKey);
      return;
    }

    const job = parseJobLine_(trimmed);
    if (!job || job.cancelled) return;

    const kind = job.transfer100 ? "local" : section.kind;
    const dest = job.transfer100 ? "936" : section.dest;
    if (!kind || section.ignore && !job.transfer100) return;

    const real = splitRealStations_(job.route);
    if (kind === "customs") {
      if (!real.length) return;
      trucks += job.nums.length;
      real.forEach(raw => {
        stations.push({ raw: raw, truck: "", kind: "customs", dest: dest, job: job.nums.join("+") });
      });
      return;
    }

    if (kind !== "local") return;
    trucks += 1;
    const truck = carrierKey === "han" ? localTruckType_(job.nums) : "";
    real.forEach(raw => {
      stations.push({ raw: raw, truck: truck, kind: "local", dest: dest, job: job.nums.join("+") });
    });
  });

  return { trucks: trucks, stations: stations };
}

function loadParsedTruckLookup_(ss) {
  const sheet = findAnalysisSheet_(ss);
  const map = new Map();
  if (!sheet || sheet.getLastRow() < ANALYSIS_DETAIL_START_ROW) return map;
  const n = sheet.getLastRow() - ANALYSIS_DETAIL_START_ROW + 1;
  const rows = sheet.getRange(ANALYSIS_DETAIL_START_ROW, 1, n, 5).getValues();
  rows.forEach(r => {
    const truck = String(r[4] || "").trim();
    if (!truck) return;
    const names = [r[2], r[3]].map(v => String(v || "").trim()).filter(Boolean);
    names.forEach(name => {
      stationAliasKeys_(name).forEach(key => {
        if (!map.has(key)) map.set(key, []);
        if (map.get(key).indexOf(truck) === -1) map.get(key).push(truck);
      });
    });
  });
  return map;
}

function lookupTruck_(map, location) {
  const keys = stationAliasKeys_(location);
  const trucks = [];
  keys.forEach(key => {
    (map.get(key) || []).forEach(t => {
      if (trucks.indexOf(t) === -1) trucks.push(t);
    });
  });
  return trucks.length === 1 ? trucks[0] : "";
}

