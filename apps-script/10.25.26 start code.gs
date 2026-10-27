/**
 * 10.25.26 start code. Paste this over the Google Sheet's Code.gs.
 * 10.24.26 end code stays in the repo and is not pasted into the sheet.
 *
 * Parse lists reads the text pasted on daily list and writes Records directly.
 * results is not used.
 *
 * One line is one truck. The size is written after the trip number:
 *   20. 26'|SireApps NJ26-PackageMan NJ208-NJ600(4PM)
 * 26' -> 26FT, 53' -> 53FT. Cargo Van is skipped until a rule is added.
 * A line containing ❌ is cancelled and skipped.
 * JFK, JFK175, EWR, EWR600, EWR936, NJ600, NJ936 at the end are the receiving
 * warehouse, not a pickup. Pickups are split on "-" outside parentheses.
 * Origin Name 1/2/3 are columns M/S/Y. A name found in Pickup Rate column F
 * is written with that spelling. Partner ID and Partner Name come from
 * columns K and L, and the address comes from column G. Those partner cells
 * are copied only when they are filled. A name Pickup Rate does not have is
 * still written, and that Origin Name cell is highlighted.
 * Pickup Window for each stop comes from Pickup Rate column H. The first
 * stop is Records column L. Trip_Pallet_Count stays blank for manual entry.
 * RATE and Carrier come from the one Pickup Rate row with the same stops,
 * the same 26' or 53', and destination EWR or JFK. NJ600 and EWR600 match EWR.
 * JFK175 matches JFK. No matching price is left blank and highlighted.
 * Capacity-EWR is not grouped by 5 stops. Each Capacity truck on the list is
 * one full truck at that row's price. JD NJ1570 / Yanwen / SF NJ122 is a flat
 * charter: one or more of those stops is still that row's price. Weekday text
 * on the rate row is ignored.
 * DATE is the date selected on daily list. REGION is EWR600.
 * LANE_TYPE is Scheduled. Row 2 of Records is the sample and is not touched.
 */

var DAILY_PASTE_START_ROW = 3;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Auto")
    .addItem("Parse lists", "parseDailyCarrierLists")
    .addToUi();
}

function parseDailyCarrierLists() {
  const selected = requireDailyListDateCell_();
  if (!selected) return;
  parseDailyCarrierListsFrom1025_(selected);
}

function normalize(str) {
  return String(str || "").replace(/\u00A0/g, " ").trim().toLowerCase();
}

function getSpreadsheetTz() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}

function parseDateString(str) {
  const s = String(str || "").trim();
  if (!s) return null;
  let m = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const first = Number(m[1]);
  const second = Number(m[2]);
  const year = Number(m[3]);
  if (first > 12) return new Date(year, second - 1, first);
  return new Date(year, first - 1, second);
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

function isDailyListSheet_(sheet) {
  const n = normalize(sheet && sheet.getName());
  return n === "daily list" || n === "daily sheet";
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

function isPastePlaceholder_(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (normalize(raw) === "粘贴在此处") return true;
  return /^粘贴\s*[（(].+[）)]\s*列表在此处$/.test(raw);
}

function stripPastePlaceholder_(value) {
  return String(value || "")
    .replace(/粘贴\s*[（(][^）)]+[）)]\s*列表在此处/g, "")
    .replace(/粘贴在此处/g, "")
    .replace(/^\s+|\s+$/g, "");
}

var HIGHLIGHT_1025_ = "#FFF2CC";
var RECORD_COLS_1025_ = 27;
var RECORD_FIRST_ROW_1025_ = 3;
var ORIGIN_SLOTS_1025_ = [
  { partnerId: 10, partnerName: 11, window: 12, origin: 13, address: 14, pallet: 15 },
  { partnerId: 16, partnerName: 17, window: 18, origin: 19, address: 20, pallet: 21 },
  { partnerId: 22, partnerName: 23, window: 24, origin: 25, address: 26, pallet: 27 }
];
var STATE_CODES_1025_ = {
  al: 1, ak: 1, az: 1, ar: 1, ca: 1, co: 1, ct: 1, de: 1, fl: 1, ga: 1,
  hi: 1, ia: 1, id: 1, il: 1, in: 1, ks: 1, ky: 1, la: 1, ma: 1, md: 1,
  me: 1, mi: 1, mn: 1, mo: 1, ms: 1, mt: 1, nc: 1, nd: 1, ne: 1, nh: 1,
  nj: 1, nm: 1, nv: 1, ny: 1, oh: 1, ok: 1, or: 1, pa: 1, ri: 1, sc: 1,
  sd: 1, tn: 1, tx: 1, ut: 1, va: 1, vt: 1, wa: 1, wi: 1, wv: 1, wy: 1
};

function parseDailyCarrierListsFrom1025_(selected) {
  const lines = readDailyListLines1025_(selected.sheet);
  const parsed = parseTruckLines1025_(lines);
  const merchants = loadMerchants1025_();
  const rates = loadPickupRates1025_(merchants.byCompact);
  const yunli = yunliTrucks1025_(rates);
  const trucks = parsed.trucks.concat(yunli);
  if (!trucks.length) {
    SpreadsheetApp.getUi().alert("daily list 里没有识别到车次，Pickup Rate 里也没有运力部卡车。Records 没有改动。");
    return;
  }
  const built = trucks.map(function (truck) {
    return buildRecord1025_(truck, merchants.byCompact, rates, selected.date);
  });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Records");
  if (!sheet) {
    SpreadsheetApp.getUi().alert("找不到 Records 表。");
    return;
  }
  const startRow = replaceRecordRows1025_(sheet, selected.date, built);
  const missingNames = built.reduce(function (n, row) { return n + row.missingNameCols.length; }, 0);
  const missingRates = built.filter(function (row) { return row.missingRate; }).length;
  const dateText = Utilities.formatDate(selected.date, getSpreadsheetTz(), "dd/MM/yyyy");
  let message = "已写入 Records 第 " + startRow + " 行起，共 " + built.length + " 车（" + dateText + "）。";
  if (yunli.length) message += "\n其中运力部卡车 " + yunli.length + " 条，没有这车时删掉对应行。";
  message += "\n板数留空。Pickup Window 用 Pickup Rate 的提货时间。";
  if (missingNames) message += "\n" + missingNames + " 个提货点在 Pickup Rate 里找不到，Origin Name 已标黄。";
  if (missingRates) message += "\n" + missingRates + " 车在 Pickup Rate 里没有对应价格，RATE 已留空并标黄。";
  if (parsed.cargoVans) message += "\n跳过 " + parsed.cargoVans + " 行 Cargo Van，这条规则还没定。";
  if (parsed.extraStops) message += "\n有 " + parsed.extraStops + " 车超过 3 个提货点，只写入前 3 个。";
  SpreadsheetApp.getUi().alert(message);
}

function readDailyListLines1025_(sheet) {
  const last = Math.max(sheet.getLastRow(), DAILY_PASTE_START_ROW);
  const height = last - DAILY_PASTE_START_ROW + 1;
  const values = sheet.getRange(DAILY_PASTE_START_ROW, 1, height, 4).getDisplayValues();
  const lines = [];
  values.forEach(function (row) {
    row.forEach(function (cell) {
      const cleaned = stripPastePlaceholder_(cell);
      if (!cleaned || isPastePlaceholder_(cleaned)) return;
      cleaned.split(/\r?\n/).forEach(function (line) {
        const text = String(line || "").trim();
        if (text) lines.push(text);
      });
    });
  });
  return lines;
}

function parseTruckLines1025_(lines) {
  const trucks = [];
  let cargoVans = 0;
  let extraStops = 0;
  lines.forEach(function (line) {
    const cleaned = String(line || "").replace(/✅/g, "").trim();
    if (!cleaned || cleaned.indexOf("❌") !== -1) return;
    if (/^\d+\.\s*cargo\s*van\b/i.test(cleaned)) {
      cargoVans++;
      return;
    }
    const truck = parseTruckLine1025_(cleaned) || parseFmCharterLine1025_(cleaned);
    if (!truck) return;
    if (truck.extra) extraStops++;
    trucks.push(truck);
  });
  return { trucks: trucks, cargoVans: cargoVans, extraStops: extraStops };
}

function parseTruckLine1025_(line) {
  const matched = String(line || "").match(/^\d+\.\s*(26|53)\s*['’‘′]\s*[|｜]\s*(.+)$/);
  if (!matched) return null;
  const parts = splitRouteOutsideParens_(matched[2]);
  let dest = "";
  if (parts.length) {
    const tail = splitStopName1025_(parts[parts.length - 1]);
    if (isReceivingWarehouse1025_(tail)) {
      dest = destKey1025_(tail);
      parts.pop();
    }
  }
  const stops = [];
  parts.forEach(function (part) {
    const name = splitStopName1025_(part);
    if (name) stops.push(name);
  });
  if (!stops.length) return null;
  return {
    vehicleKey: matched[1],
    vehicleType: matched[1] + "FT",
    dest: dest,
    stops: stops.slice(0, 3),
    extra: stops.length > 3
  };
}

function parseFmCharterLine1025_(line) {
  const matched = String(line || "").match(/^\d+\.\s*(.+)$/);
  if (!matched) return null;
  const body = String(matched[1] || "").trim();
  if (/^(26|53)\b/.test(body)) return null;
  const parts = splitRouteOutsideParens_(body);
  let dest = "";
  if (parts.length) {
    const tail = splitStopName1025_(parts[parts.length - 1]);
    if (isReceivingWarehouse1025_(tail)) {
      dest = destKey1025_(tail);
      parts.pop();
    }
  }
  const stops = [];
  parts.forEach(function (part) {
    const name = splitStopName1025_(part);
    if (name) stops.push(name);
  });
  if (!stops.length || !stops.every(looksLikeFmCharterStop1025_)) return null;
  return {
    vehicleKey: "26",
    vehicleType: "26FT",
    dest: dest || "ewr",
    stops: stops.slice(0, 3),
    extra: stops.length > 3
  };
}

function looksLikeFmCharterStop1025_(name) {
  const raw = compactStop_(name);
  if (raw.indexOf("yanwen") !== -1) return true;
  if (raw === "sffm" || raw === "sfnj122") return true;
  if (raw === "jdnj1570") return true;
  return false;
}

function splitStopName1025_(token) {
  const text = String(token || "").trim();
  const matched = text.match(/^(.*?)(?:\s*[\(（].*[\)）])\s*$/);
  return String(matched ? matched[1] : text).trim();
}

function isReceivingWarehouse1025_(name) {
  const key = compactStop_(name);
  return key === "jfk" || key === "jfk175" || key === "ewr" || key === "ewr600" ||
    key === "ewr936" || key === "nj600" || key === "nj936";
}

function destKey1025_(name) {
  const key = compactStop_(name);
  if (key === "jfk" || key === "jfk175") return "jfk";
  if (key === "ewr" || key === "ewr600" || key === "nj600") return "ewr";
  return key;
}

function compactStop_(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9\u4e00-\u9fff]/g, "");
}

function splitRouteOutsideParens_(text) {
  const parts = [];
  let buf = "";
  let depth = 0;
  const source = String(text || "");
  for (let i = 0; i < source.length; i++) {
    const ch = source.charAt(i);
    if (ch === "(" || ch === "（") depth++;
    if (ch === ")" || ch === "）") depth = Math.max(0, depth - 1);
    if (ch === "-" && depth === 0) {
      if (buf.trim()) parts.push(buf.trim());
      buf = "";
      continue;
    }
    buf += ch;
  }
  if (buf.trim()) parts.push(buf.trim());
  return parts;
}

function loadMerchants1025_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pickup Rate");
  const byCompact = {};
  if (!sheet || sheet.getLastRow() < 2) return { byCompact: byCompact };
  const values = sheet.getRange(2, 6, sheet.getLastRow() - 1, 7).getValues();
  values.forEach(function (row) {
    const name = String(row[0] || "").trim();
    if (!name) return;
    byCompact[compactStop_(name)] = {
      name: name,
      address: String(row[1] || "").trim(),
      window: String(row[2] || "").trim(),
      partnerId: row[5] === "" || row[5] === null ? "" : row[5],
      partnerName: String(row[6] || "").trim()
    };
  });
  return { byCompact: byCompact };
}

function lookupMerchant1025_(name, byCompact) {
  const raw = compactStop_(name);
  if (!raw) return null;
  if (byCompact[raw]) return byCompact[raw];
  if (raw === "willownj3509" && byCompact.wilownj3509) return byCompact.wilownj3509;
  if (raw === "ttnj1001" && byCompact.nj1001) return byCompact.nj1001;
  if (raw === "sffm" && byCompact.sfnj122) return byCompact.sfnj122;
  const flipped = flipStateBrand1025_(raw);
  if (flipped && byCompact[flipped]) return byCompact[flipped];
  const capacity = lookupCapacityMerchant1025_(raw, byCompact);
  if (capacity) return capacity;
  return null;
}

function lookupCapacityMerchant1025_(raw, byCompact) {
  if (String(raw || "").indexOf("capacity") !== 0) return null;
  const digits = String(raw).match(/(\d{3,5})$/);
  if (!digits) return null;
  const hits = [];
  const seen = {};
  Object.keys(byCompact).forEach(function (key) {
    const merchant = byCompact[key];
    const name = merchant && merchant.name;
    if (!name || seen[name]) return;
    const compactName = compactStop_(name);
    if (compactName.indexOf("capacity") === -1 || compactName.indexOf(digits[1]) === -1) return;
    seen[name] = true;
    hits.push(merchant);
  });
  return hits.length === 1 ? hits[0] : null;
}

function flipStateBrand1025_(raw) {
  const matched = String(raw || "").match(/^([a-z]{2})([a-z]+)(\d+)$/);
  if (!matched) return "";
  return matched[2] + matched[1] + matched[3];
}

function stopKey1025_(name, byCompact) {
  const found = lookupMerchant1025_(name, byCompact);
  if (found) return compactStop_(found.name);
  let raw = compactStop_(name);
  if (raw === "willownj3509") raw = "wilownj3509";
  const matched = raw.match(/^([a-z]{2})([a-z]+)(\d+)$/);
  if (!matched) return raw;
  const left = matched[1];
  const right = matched[2];
  const digits = matched[3];
  if (STATE_CODES_1025_[left] && !STATE_CODES_1025_[right]) return right + left + digits;
  if (STATE_CODES_1025_[right] && !STATE_CODES_1025_[left]) return left + right + digits;
  return raw;
}

function loadPickupRates1025_(byCompact) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pickup Rate");
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const rates = [];
  values.forEach(function (row) {
    const route = String(row[0] || "").trim();
    const vehicle = vehicleKey1025_(row[2]);
    if (!route || !vehicle) return;
    const special = specialRateKind1025_(route);
    if (special) {
      const price = parsePrice1025_(row[1]);
      if (price === "") return;
      rates.push({
        special: special,
        keys: [],
        dest: "ewr",
        vehicle: vehicle,
        price: price,
        carrier: String(row[3] || "").trim()
      });
      return;
    }
    const parsed = parseRateRoute1025_(route, byCompact);
    if (!parsed.dest || !parsed.keys.length) return;
    rates.push({
      keys: parsed.keys,
      names: parsed.names,
      dest: parsed.dest,
      vehicle: vehicle,
      price: parsePrice1025_(row[1]),
      carrier: String(row[3] || "").trim()
    });
  });
  return rates;
}

function parseRateRoute1025_(route, byCompact) {
  const tokens = splitRouteOutsideParens_(route);
  let dest = "";
  if (tokens.length && isReceivingWarehouse1025_(tokens[tokens.length - 1])) {
    dest = destKey1025_(tokens.pop());
  }
  const keys = [];
  const names = [];
  let i = 0;
  while (i < tokens.length) {
    let foundAt = -1;
    for (let j = tokens.length; j > i; j--) {
      if (lookupMerchant1025_(tokens.slice(i, j).join("-"), byCompact)) {
        foundAt = j;
        break;
      }
    }
    if (foundAt !== -1) {
      const joined = tokens.slice(i, foundAt).join("-");
      const merchant = lookupMerchant1025_(joined, byCompact);
      keys.push(stopKey1025_(joined, byCompact));
      names.push(merchant ? merchant.name : joined);
      i = foundAt;
    } else {
      keys.push(stopKey1025_(tokens[i], byCompact));
      names.push(String(tokens[i] || "").trim());
      i++;
    }
  }
  return { keys: keys, names: names, dest: dest };
}

function yunliTrucks1025_(rates) {
  const trucks = [];
  rates.forEach(function (rate) {
    if (String(rate.carrier || "").replace(/\s/g, "") !== "运力部卡车") return;
    if (!rate.names || !rate.names.length || !rate.dest) return;
    if (rate.vehicle !== "26" && rate.vehicle !== "53") return;
    trucks.push({
      vehicleKey: rate.vehicle,
      vehicleType: rate.vehicle + "FT",
      dest: rate.dest,
      stops: rate.names.slice(0, 3),
      extra: rate.names.length > 3
    });
  });
  return trucks;
}

function vehicleKey1025_(value) {
  const text = String(value || "").toLowerCase();
  if (text.indexOf("53") !== -1) return "53";
  if (text.indexOf("cargo") !== -1) return "van";
  if (text.indexOf("26") !== -1) return "26";
  return "";
}

function parsePrice1025_(value) {
  if (typeof value === "number" && isFinite(value)) return value;
  const text = String(value || "").replace(/[$,\s]/g, "");
  if (!text) return "";
  const number = Number(text);
  return isFinite(number) ? number : "";
}

function specialRateKind1025_(route) {
  const key = compactStop_(route);
  if (key === "capacity" || key === "capacityewr") return "capacity";
  if (key.indexOf("jdnj1570") !== -1 && key.indexOf("yanwen") !== -1 && key.indexOf("sfnj122") !== -1) return "fm";
  return "";
}

function isCapacityStop1025_(name, byCompact) {
  const merchant = lookupMerchant1025_(name, byCompact);
  return compactStop_(merchant ? merchant.name : name).indexOf("capacity") !== -1;
}

function isFmCharterStop1025_(name, byCompact) {
  const merchant = lookupMerchant1025_(name, byCompact);
  const raw = compactStop_(name);
  const label = compactStop_(merchant ? merchant.name : "");
  if (raw.indexOf("yanwen") !== -1 || label.indexOf("yanwen") !== -1) return true;
  if (raw === "sffm" || raw === "sfnj122" || label === "sfnj122") return true;
  if (raw === "jdnj1570" || label === "jdnj1570") return true;
  return false;
}

function specialPickupRate1025_(rates, kind) {
  const hits = rates.filter(function (rate) {
    return rate.special === kind && rate.vehicle === "26";
  });
  return hits.length === 1 ? hits[0] : null;
}

function findPickupRate1025_(truck, byCompact, rates) {
  const keys = truck.stops.map(function (name) { return stopKey1025_(name, byCompact); });
  const hits = rates.filter(function (rate) {
    if (rate.special) return false;
    if (rate.vehicle !== truck.vehicleKey || rate.dest !== truck.dest) return false;
    if (rate.keys.length !== keys.length) return false;
    for (let i = 0; i < keys.length; i++) {
      if (rate.keys[i] !== keys[i]) return false;
    }
    return true;
  });
  if (hits.length === 1) return hits[0];
  if (truck.vehicleKey !== "26" || (truck.dest && truck.dest !== "ewr")) return null;
  if (truck.stops.length && truck.stops.every(function (name) { return isCapacityStop1025_(name, byCompact); })) {
    return specialPickupRate1025_(rates, "capacity");
  }
  if (truck.stops.length && truck.stops.every(function (name) { return isFmCharterStop1025_(name, byCompact); })) {
    return specialPickupRate1025_(rates, "fm");
  }
  return null;
}

function buildRecord1025_(truck, byCompact, rates, date) {
  const row = [];
  for (let i = 0; i < RECORD_COLS_1025_; i++) row.push("");
  row[0] = Utilities.formatDate(date, getSpreadsheetTz(), "dd/MM/yyyy");
  row[1] = "EWR600";
  row[2] = "Scheduled";
  row[3] = truck.vehicleType;
  row[6] = truck.stops.length;
  const rate = findPickupRate1025_(truck, byCompact, rates);
  if (rate && rate.price !== "") {
    row[4] = rate.price;
    row[5] = rate.carrier;
  }
  const missingNameCols = [];
  truck.stops.forEach(function (stopName, index) {
    const slot = ORIGIN_SLOTS_1025_[index];
    if (!slot) return;
    const merchant = lookupMerchant1025_(stopName, byCompact);
    if (!merchant) {
      row[slot.origin - 1] = stopName;
      missingNameCols.push(slot.origin);
      return;
    }
    row[slot.origin - 1] = merchant.name;
    row[slot.address - 1] = merchant.address;
    if (merchant.window) row[slot.window - 1] = merchant.window;
    if (merchant.partnerId !== "") row[slot.partnerId - 1] = merchant.partnerId;
    if (merchant.partnerName) row[slot.partnerName - 1] = merchant.partnerName;
  });
  return {
    values: row,
    missingRate: !(rate && rate.price !== ""),
    missingNameCols: missingNameCols
  };
}

function recordDateTextMatches1025_(shown, date, dateText) {
  const text = String(shown || "").trim();
  if (!text) return false;
  if (text === dateText) return true;
  const parsed = parseSheetDate(text);
  return !!(parsed && parsed.getFullYear() === date.getFullYear() &&
    parsed.getMonth() === date.getMonth() && parsed.getDate() === date.getDate());
}

function replaceRecordRows1025_(sheet, date, built) {
  const dateText = Utilities.formatDate(date, getSpreadsheetTz(), "dd/MM/yyyy");
  const last = sheet.getLastRow();
  const existing = [];
  if (last >= RECORD_FIRST_ROW_1025_) {
    const values = sheet.getRange(RECORD_FIRST_ROW_1025_, 1, last - RECORD_FIRST_ROW_1025_ + 1, 1).getDisplayValues();
    values.forEach(function (row, index) {
      if (recordDateTextMatches1025_(row[0], date, dateText)) existing.push(RECORD_FIRST_ROW_1025_ + index);
    });
  }
  let start = nextRecordStart1025_(sheet);
  if (existing.length) {
    start = existing[0];
    const contiguous = existing.every(function (row, index) { return row === start + index; });
    if (contiguous) {
      const extra = built.length - existing.length;
      if (extra > 0) sheet.insertRowsAfter(start + existing.length - 1, extra);
      if (extra < 0) sheet.deleteRows(start + built.length, -extra);
    } else {
      for (let i = existing.length - 1; i >= 1; i--) sheet.deleteRow(existing[i]);
      if (built.length > 1) sheet.insertRowsAfter(start, built.length - 1);
    }
  }
  const range = sheet.getRange(start, 1, built.length, RECORD_COLS_1025_);
  range.setBackground(null);
  sheet.getRange(start, 1, built.length, 1).setNumberFormat("@");
  applyRecordAmountFormats1025_(sheet);
  range.setValues(built.map(function (item) { return item.values; }));
  built.forEach(function (item, index) {
    const row = start + index;
    if (item.missingRate) sheet.getRange(row, 5).setBackground(HIGHLIGHT_1025_);
    item.missingNameCols.forEach(function (col) {
      sheet.getRange(row, col).setBackground(HIGHLIGHT_1025_);
    });
  });
  return start;
}

function nextRecordStart1025_(sheet) {
  const last = sheet.getLastRow();
  if (last < RECORD_FIRST_ROW_1025_) return RECORD_FIRST_ROW_1025_;
  const values = sheet.getRange(RECORD_FIRST_ROW_1025_, 1, last - RECORD_FIRST_ROW_1025_ + 1, 1).getDisplayValues();
  let lastRecord = 0;
  values.forEach(function (row, index) {
    if (String(row[0] || "").trim()) lastRecord = RECORD_FIRST_ROW_1025_ + index;
  });
  if (!lastRecord) return RECORD_FIRST_ROW_1025_;
  return lastRecord + 2;
}

function applyRecordAmountFormats1025_(sheet) {
  sheet.getRange("E2:E").setNumberFormat("$#,##0");
  sheet.getRange("H2:H").setNumberFormat("0");
}
