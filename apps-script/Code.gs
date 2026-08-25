function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Autofill")
    .addItem("Run", "fillPickupData")
    .addToUi();
}

function normalize(str) {
  return String(str || "").replace(/\u00A0/g, " ").trim().toLowerCase();
}

function compact(str) {
  return normalize(str).replace(/\s+/g, "");
}

const CHARTER_GROUPS = [
  { id: "CHARTER_FM", locations: ["SF-FM", "JD NJ1570", "Yanwen"] },
  { id: "CHARTER_SHIPCUBE", locations: ["ShipCube PA11200", "ShipCube PA700"] },
  { id: "CHARTER_CAPACITY_1112_1000", locations: ["Capacity NJ1112", "Capacity NJ1000"] },
  { id: "CHARTER_CAPACITY_1980_1600", locations: ["Capacity NJ1980", "Capacity NJ1600"] }
];

const COL = { ADDRESS: 2, CARRIER: 3, TRUCK: 4, STATE: 5, PRICE: 6, PALLETS: 8, TIKTOK_LABEL: 9, TIKTOK_PRICE: 10, TIKTOK_TRUCK: 11 };

const TIKTOK_MERCHANTS = ["NJ TT1001", "NJ TT245", "NJ TT511", "Swift X NJ650"];
const TIKTOK_PRICE_ORDER = [200, 100, 360, 180];

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
  return false;
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
      parseSheetDate(r[10]).getTime() < matchDate.getTime()
    )
    .sort((a, b) => parseSheetDate(b[10]).getTime() - parseSheetDate(a[10]).getTime());

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
    const carrier = rows.map(r => String(r[3] || "").trim()).find(Boolean);
    const state = rows.map(r => String(r[5] || "").trim()).find(Boolean);

    rows.forEach(r => {
      if (!parseSheetDate(r[10]) && effectiveDate) r[10] = effectiveDate;
      const cur = Number(String(r[6]).replace(/[$,]/g, ""));
      if (!cur && price) r[6] = price;
      if (!String(r[3] || "").trim() && carrier) r[3] = carrier;
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const cell = sheet.getActiveCell();
  if (cell.getColumn() !== 1) {
    SpreadsheetApp.getUi().alert("Select column A");
    return;
  }

  const startRow = cell.getRow();
  const dateValue = parseSheetDate(cell);
  const weekday = Utilities.formatDate(dateValue, getSpreadsheetTz(), "EEE");

  let endRow = startRow;
  while (sheet.getRange(endRow + 1, 2).getValue()) endRow++;

  const range = sheet.getRange(startRow, 1, endRow - startRow + 1, 12);
  const data = range.getValues();
  const charterMeta = new Array(data.length).fill(null);

  const db = loadDatabase(ss.getSheetByName("database"));
  enrichCharterDbRows(db);

  const rulesRaw = ss.getSheetByName("rules").getDataRange().getValues();
  rulesRaw.shift();
  const rules = rulesRaw.map(r => ({
    type: normalize(r[0]),
    value: normalize(r[1]),
    truckRule: r[2],
    useDbPrice: String(r[3]).toUpperCase() === "TRUE"
  }));

  for (let i = 0; i < data.length; i++) {
    const location = data[i][1];
    if (!location) continue;

    const candidates = getCandidates(location, db, dateValue);
    if (!candidates.length) continue;

    const dated = candidates.filter(r => parseSheetDate(r[10]));
    const pool = dated.length ? dated : candidates;
    const maxDate = Math.max(...pool.map(r => parseSheetDate(r[10]).getTime()));
    const latest = pool.filter(r => parseSheetDate(r[10]).getTime() === maxDate);
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
        continue;
      }
    }

    match = fillMissingFromOlder(match, location, db, dateValue);

    const charterGroup = getCharterGroupByLocation(match[1]);
    const isCharter = !!charterGroup;
    const basePrice = Number(String(match[6]).replace(/[$,]/g, ""));

    let matchedRule = rules.find(r => r.type === "carrier" && r.value === normalize(match[3]));
    if (!matchedRule) {
      matchedRule = rules.find(r => r.type === "location" && r.value === normalize(location));
    }
    if (matchedRule) {
      const t = String(matchedRule.truckRule).trim();
      if (t === "53'") data[i][COL.TRUCK] = "53'";
      else if (t === "DB") data[i][COL.TRUCK] = match[4];
    }

    data[i][COL.ADDRESS] = match[2];
    data[i][COL.CARRIER] = match[3];
    data[i][COL.STATE] = match[5];

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
  applyTikTokSummary(data);
  range.setValues(data);
  sheet.getRange(startRow, 10, 1, 3).setWrapStrategy(SpreadsheetApp.WrapStrategy.OVERFLOW);
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
  const n = Number(String(value || "").replace(/[$,]/g, "").trim());
  return isFinite(n) ? n : 0;
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
