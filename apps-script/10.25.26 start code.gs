/**
 * 10.25.26 start code. Paste this over the Google Sheet's Code.gs.
 * 10.24.26 end code stays in the repo and is not pasted into the sheet.
 *
 * Parse lists reads the text pasted on daily list and writes Records directly.
 * Select row-1 date to parse A-D (Han Express / Inspire / FM / SD).
 * In Han Express (A), only the 【NJ600提货】 section is used.
 * Select E2 (NJ100/936) to parse column E 【NJ100提货】 or 【NJ936提货】.
 * results is not used.
 *
 * One line is one truck. The size is written after the trip number:
 *   20. 26'|SireApps NJ26-PackageMan NJ208-NJ600(4PM)
 * 26' -> 26FT, 53' -> 53FT. Cargo Van is skipped until a rule is added.
 * A line containing ❌ is cancelled and skipped.
 * JFK, JFK175, EWR, EWR600, EWR936, NJ600, NJ936 at the end are the receiving
 * warehouse, not a pickup. Pickups are split on "-" outside parentheses.
 * Origin Name 1/2/3 are columns M/S/Y. A name found in Pickup Rate column F
 * is written with that spelling, except TTNJ1001 and NJ1001, which stay
 * exactly as written on the list while still mapping to Pickup Rate NJ1001.
 * Partner ID and Partner Name come from
 * columns K and L, and the address comes from column G. Those partner cells
 * are copied only when they are filled. A name Pickup Rate does not have is
 * still written, and that Origin Name cell is highlighted.
 * Pickup Window for each stop comes from Pickup Rate column H. The first
 * stop is Records column L. Origin pallet counts are typed in columns O, U,
 * and AA. Confirm pallet count: select the first Trip_Pallet_Count cell of
 * the block (e.g. H23), or drag H23:H67; it only confirms that row range,
 * fills mapping for added origins, and updates the trip rate.
 * Origin Name 1/2/3 use a dropdown from Pickup Rate column F.
 * RATE and Carrier come from the one Pickup Rate row with the same stops,
 * the same 26' or 53', and destination EWR or JFK. NJ600, EWR600, NJ100,
 * and NJ936 temporarily match EWR rates. JFK175 matches JFK.
 * No matching price is left blank and highlighted.
 * Capacity-EWR is not grouped by 5 stops. Each Capacity truck on the list is
 * one full truck at that row's price. JD NJ1570 / Yanwen / SF NJ122 is a flat
 * charter: one truck at that row's price even when the FM list splits them
 * onto separate lines (merged into Origin 1/2/3). Weekday text on the rate
 * row is ignored.
 * DATE is the date selected on daily list. REGION is EWR600.
 * LANE_TYPE is Scheduled. Row 2 of Records is the sample and is not touched.
 * Parse always appends after the last record with one blank row in between.
 * Existing rows for the same date are not overwritten.
 */

var DAILY_PASTE_START_ROW = 3;

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Auto")
    .addItem("Parse lists", "parseDailyCarrierLists")
    .addItem("Confirm pallet count", "confirmPalletCount")
    .addToUi();
  showYesterdayOnDailyList_();
  applyOriginNameDropdowns1025_();
  styleDailyPasteCells1025_();
}

function onEdit(e) {
  handleDailyListPasteEdit_(e);
}

function handleDailyListPasteEdit_(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (!isDailyListSheet_(sheet)) return;
  const startRow = e.range.getRow();
  const endRow = e.range.getLastRow();
  if (endRow < DAILY_PASTE_START_ROW || startRow > DAILY_PASTE_START_ROW) return;
  const col = e.range.getColumn();
  const lastCol = Math.min(e.range.getLastColumn(), 5);
  for (let c = col; c <= lastCol; c++) {
    styleDailyPasteCell1025_(sheet.getRange(DAILY_PASTE_START_ROW, c));
  }
}

function styleDailyPasteCells1025_() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let sheet = null;
  for (let i = 0; i < sheets.length; i++) {
    if (isDailyListSheet_(sheets[i])) {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) return;
  for (let c = 1; c <= 5; c++) {
    styleDailyPasteCell1025_(sheet.getRange(DAILY_PASTE_START_ROW, c));
  }
}

function dailyPastePlaceholder1025_(column) {
  if (column === 1) return "粘贴 (Han Express) 列表在此处";
  if (column === 2) return "粘贴 (Inspire) 列表在此处";
  if (column === 3) return "粘贴 (FM) 列表在此处";
  if (column === 4) return "粘贴 (SD) 列表在此处";
  if (column === 5) return "粘贴（NJ100/936）提货列表在此处";
  return "粘贴在此处";
}

function styleDailyPasteCell1025_(cell) {
  const raw = String(cell.getDisplayValue() || "");
  const trimmed = raw.trim();
  const placeholder = dailyPastePlaceholder1025_(cell.getColumn());
  const cleaned = stripPastePlaceholder_(raw).trim();
  if (!trimmed || isPastePlaceholder_(trimmed) || !cleaned) {
    applyDailyPastePlaceholderStyle_(cell, placeholder);
    return;
  }
  applyDailyPasteBodyStyle_(cell, cleaned);
}

function applyDailyPastePlaceholderStyle_(cell, placeholder) {
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(placeholder).build());
  cell.setFontFamily("Arial");
  cell.setFontColor("#B7B7B7");
  cell.setFontStyle("italic");
  cell.setFontWeight("normal");
}

function applyDailyPasteBodyStyle_(cell, text) {
  cell.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).build());
  cell.setFontFamily("Arial");
  cell.setFontColor("#000000");
  cell.setFontStyle("normal");
  cell.setFontWeight("normal");
}

function showYesterdayOnDailyList_() {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let sheet = null;
  for (let i = 0; i < sheets.length; i++) {
    const name = normalize(sheets[i].getName());
    if (name === "daily list" || name === "daily sheet") {
      sheet = sheets[i];
      break;
    }
  }
  if (!sheet) return;
  const todayText = Utilities.formatDate(new Date(), getSpreadsheetTz(), "yyyy-MM-dd");
  const parts = todayText.split("-");
  const yesterday = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]) - 1);
  const cell = sheet.getRange("A1");
  const current = parseSheetDate(cell);
  if (current && current.getFullYear() === yesterday.getFullYear() &&
      current.getMonth() === yesterday.getMonth() &&
      current.getDate() === yesterday.getDate()) return;
  cell.setValue(yesterday);
  cell.setNumberFormat("yyyy-mm-dd");
}

function parseDailyCarrierLists() {
  styleDailyPasteCells1025_();
  const selected = requireDailyListParseSelection_();
  if (!selected) return;
  parseDailyCarrierListsFrom1025_(selected);
}

function confirmPalletCount() {
  const block = requireConfirmPalletBlock_();
  if (!block) return;
  const sheet = block.sheet;
  const start = block.start;
  const end = block.end;
  const merchants = loadMerchants1025_();
  const rates = loadPickupRates1025_(merchants.byCompact);
  const height = end - start + 1;
  const range = sheet.getRange(start, 1, height, RECORD_COLS_1025_);
  const values = range.getValues();
  let lastDate = "";
  if (start > RECORD_FIRST_ROW_1025_) {
    lastDate = String(sheet.getRange(start - 1, 1).getDisplayValue() || "").trim();
  }
  let confirmed = 0;
  let missingNames = 0;
  let missingRates = 0;
  const highlights = [];
  values.forEach(function (row, rowIndex) {
    const names = ORIGIN_SLOTS_1025_.map(function (slot) {
      return String(row[slot.origin - 1] || "").trim();
    });
    const pallets = ORIGIN_SLOTS_1025_.map(function (slot) {
      return palletRaw1025_(row[slot.pallet - 1]);
    });
    const hasName = names.some(function (name) { return name; });
    const hasPallet = pallets.some(function (item) { return item; });
    if (!hasName && !hasPallet) return;
    if (row[0]) lastDate = row[0];
    if (hasName && !row[0] && lastDate) {
      row[0] = lastDate;
      if (!row[1]) row[1] = "EWR600";
      if (!row[2]) row[2] = "Scheduled";
    }
    ORIGIN_SLOTS_1025_.forEach(function (slot, index) {
      if (!names[index]) {
        row[slot.partnerId - 1] = "";
        row[slot.partnerName - 1] = "";
        row[slot.window - 1] = "";
        row[slot.address - 1] = "";
        highlights.push({ rowIndex: rowIndex, col: slot.origin, missing: false });
        return;
      }
      const merchant = lookupMerchant1025_(names[index], merchants.byCompact);
      if (!merchant) {
        missingNames++;
        highlights.push({ rowIndex: rowIndex, col: slot.origin, missing: true, plain: true });
        return;
      }
      row[slot.origin - 1] = keepsListSpelling1025_(names[index]) ? names[index] : merchant.name;
      row[slot.address - 1] = merchant.address;
      row[slot.window - 1] = merchant.window || "";
      row[slot.partnerId - 1] = merchant.partnerId === "" ? "" : merchant.partnerId;
      row[slot.partnerName - 1] = merchant.partnerName || "";
      highlights.push({ rowIndex: rowIndex, col: slot.origin, missing: false, plain: false });
    });
    const stops = names.filter(function (name) { return name; });
    if (hasName) row[6] = stops.length;
    if (hasPallet) {
      const tripPallet = combineTripPalletCount1025_(pallets);
      if (tripPallet) {
        row[7] = tripPallet.value;
        if (tripPallet.asText) {
          highlights.push({ rowIndex: rowIndex, col: 8, tripPalletText: true });
        }
      }
    }
    if (stops.length) {
      const matched = confirmRate1025_(stops, row[3], merchants.byCompact, rates);
      if (matched && matched.rate && matched.rate.price !== "") {
        row[4] = matched.rate.price;
        row[5] = matched.rate.carrier;
        if (!vehicleKey1025_(row[3])) row[3] = matched.vehicleType;
        highlights.push({ rowIndex: rowIndex, col: 5, missing: false });
      } else {
        row[4] = "";
        row[5] = "";
        missingRates++;
        highlights.push({ rowIndex: rowIndex, col: 5, missing: true });
      }
    }
    confirmed++;
  });
  range.setValues(values);
  const originRule = originNameDropdownRule1025_();
  highlights.forEach(function (item) {
    const cell = sheet.getRange(start + item.rowIndex, item.col);
    if (item.tripPalletText) {
      cell.setNumberFormat("@");
      return;
    }
    if (item.plain) {
      cell.clearDataValidations();
      cell.setNumberFormat("@");
      cell.setBackground(HIGHLIGHT_1025_);
      return;
    }
    if (item.col === 5) {
      cell.setBackground(item.missing ? HIGHLIGHT_1025_ : null);
      return;
    }
    cell.setDataValidation(originRule);
    cell.setBackground(item.missing ? HIGHLIGHT_1025_ : ORIGIN_NAME_FILL_1025_);
  });
  applyOriginPalletFormats1025_(sheet);
  let message = "已确认 H" + start + ":H" + end + "，共 " + confirmed + " 行。";
  message += "\nTrip Pallet Count：纯数字加总；含文本则按原文拼到 H（如 10+515pcs）。";
  message += "\n后加的提货点已按 Pickup Rate 补上地址、时间和 Partner，并重算了车次价格。";
  if (missingNames) message += "\n" + missingNames + " 个提货点在 Pickup Rate 里找不到，Origin Name 已标黄。";
  if (missingRates) message += "\n" + missingRates + " 车没有对应价格，RATE 已留空并标黄。";
  SpreadsheetApp.getUi().alert(message);
}

function requireConfirmPalletBlock_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (normalize(sheet.getName()) !== "records") {
    SpreadsheetApp.getUi().alert("先到 Records，选中该段第一条 Trip_Pallet_Count（H 列），例如 H23，或拖选 H23:H67。");
    return null;
  }
  const active = sheet.getActiveRange();
  if (!active || active.getColumn() !== 8 || active.getLastColumn() !== 8) {
    SpreadsheetApp.getUi().alert("请选中 H 列 Trip_Pallet_Count：该段第一格（如 H23），或拖选整段（如 H23:H67）。");
    return null;
  }
  let start = active.getRow();
  let end = active.getLastRow();
  if (start < RECORD_FIRST_ROW_1025_) {
    SpreadsheetApp.getUi().alert("请从 Records 第 " + RECORD_FIRST_ROW_1025_ + " 行起的 H 列选择。");
    return null;
  }
  if (start === end) end = findConfirmBlockEndRow1025_(sheet, start);
  if (end < start) {
    SpreadsheetApp.getUi().alert("选中位置没有可确认的车次。");
    return null;
  }
  return { sheet: sheet, start: start, end: end };
}

function findConfirmBlockEndRow1025_(sheet, start) {
  const last = sheet.getLastRow();
  if (last < start) return start;
  const height = last - start + 1;
  const values = sheet.getRange(start, 1, height, RECORD_COLS_1025_).getValues();
  let end = start - 1;
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const hasDate = String(row[0] || "").trim();
    const hasName = ORIGIN_SLOTS_1025_.some(function (slot) {
      return String(row[slot.origin - 1] || "").trim();
    });
    const hasPallet = ORIGIN_SLOTS_1025_.some(function (slot) {
      return !!palletRaw1025_(row[slot.pallet - 1]);
    });
    if (!hasDate && !hasName && !hasPallet) {
      if (i === 0) return start - 1;
      break;
    }
    end = start + i;
  }
  return end;
}

function palletRaw1025_(value) {
  if (typeof value === "number" && isFinite(value)) {
    return { kind: "number", number: value, text: String(value) };
  }
  const text = String(value || "").replace(/\u00A0/g, " ").trim();
  if (!text) return null;
  const normalized = text.replace(/,/g, "");
  if (/^[+-]?\d+(\.\d+)?$/.test(normalized)) {
    const number = Number(normalized);
    if (isFinite(number)) return { kind: "number", number: number, text: normalized };
  }
  return { kind: "text", text: text };
}

function combineTripPalletCount1025_(pallets) {
  const parts = [];
  pallets.forEach(function (item) { if (item) parts.push(item); });
  if (!parts.length) return null;
  const hasText = parts.some(function (item) { return item.kind === "text"; });
  if (!hasText) {
    let sum = 0;
    parts.forEach(function (item) { sum += item.number; });
    return { value: sum, asText: false };
  }
  return {
    value: parts.map(function (item) {
      return item.kind === "number" ? String(item.number) : item.text;
    }).join("+"),
    asText: true
  };
}

function palletNumber1025_(value) {
  const raw = palletRaw1025_(value);
  return raw && raw.kind === "number" ? raw.number : null;
}

function confirmRate1025_(stops, vehicleType, byCompact, rates) {
  const vehicleKey = vehicleKey1025_(vehicleType);
  const keys = stops.map(function (name) { return stopKey1025_(name, byCompact); });
  const hits = rates.filter(function (rate) {
    if (rate.special || !rate.keys) return false;
    if (vehicleKey && rate.vehicle !== vehicleKey) return false;
    if (rate.keys.length !== keys.length) return false;
    for (let i = 0; i < keys.length; i++) {
      if (rate.keys[i] !== keys[i]) return false;
    }
    return true;
  });
  const picked = pickConfirmRate1025_(hits, vehicleKey);
  if (picked) return picked;
  if (vehicleKey && vehicleKey !== "26") return null;
  const special = findPickupRate1025_({ stops: stops, vehicleKey: "26", dest: "ewr" }, byCompact, rates);
  if (!special || !special.special) return null;
  return { rate: special, vehicleType: vehicleType || "26FT" };
}

function pickConfirmRate1025_(hits, vehicleKey) {
  if (!hits.length) return null;
  const ewr = hits.filter(function (rate) { return rate.dest === "ewr"; });
  const pool = ewr.length ? ewr : hits;
  const vehicles = [];
  pool.forEach(function (rate) {
    if (vehicles.indexOf(rate.vehicle) === -1) vehicles.push(rate.vehicle);
  });
  if (!vehicleKey && vehicles.length !== 1) return null;
  const chosen = vehicleKey || vehicles[0];
  const finalHits = pool.filter(function (rate) { return rate.vehicle === chosen; });
  if (finalHits.length !== 1) return null;
  return { rate: finalHits[0], vehicleType: chosen + "FT" };
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

function requireDailyListParseSelection_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (!isDailyListSheet_(sheet)) {
    SpreadsheetApp.getUi().alert("先到 daily list，选中 A1:D1 日期格，或 E2（NJ100/936）。");
    return null;
  }
  if (cell.getRow() === 2 && cell.getColumn() === 5) {
    const date = parseSheetDate(sheet.getRange("A1"));
    if (!date) {
      SpreadsheetApp.getUi().alert("A1 必须是日期。");
      return null;
    }
    return { sheet: sheet, date: date, mode: "nj100936" };
  }
  if (cell.getRow() === 1 && cell.getColumn() <= 4) {
    const date = parseSheetDate(cell);
    if (!date) {
      SpreadsheetApp.getUi().alert("A1:D1 必须是日期。");
      return null;
    }
    return { sheet: sheet, date: date, mode: "nj600" };
  }
  SpreadsheetApp.getUi().alert("先选中 daily list 的 A1:D1 日期格，或 E2（NJ100/936）。");
  return null;
}

function isPastePlaceholder_(value) {
  const raw = String(value || "").trim();
  if (!raw) return false;
  if (normalize(raw) === "粘贴在此处") return true;
  return /^粘贴\s*[（(].+[）)]\s*(?:提货)?列表在此处$/.test(raw);
}

function stripPastePlaceholder_(value) {
  return String(value || "")
    .replace(/粘贴\s*[（(][^）)]+[）)]\s*(?:提货)?列表在此处/g, "")
    .replace(/粘贴在此处/g, "")
    .replace(/^\s+|\s+$/g, "");
}

var HIGHLIGHT_1025_ = "#FFF2CC";
var ORIGIN_NAME_FILL_1025_ = "#FCE5CD";
var ORIGIN_PALLET_FILL_1025_ = "#D9EAD3";
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
  const mode = selected.mode || "nj600";
  const lines = readDailyListLines1025_(selected.sheet, mode);
  const parsed = parseTruckLines1025_(lines);
  const merchants = loadMerchants1025_();
  const rates = loadPickupRates1025_(merchants.byCompact);
  const yunli = mode === "nj600" ? yunliTrucks1025_(rates) : [];
  const trucks = parsed.trucks.concat(yunli);
  if (!trucks.length) {
    const sectionLabel = mode === "nj100936"
      ? "E 列【NJ100提货】/【NJ936提货】"
      : "A-D 列表（Han Express 仅【NJ600提货】）";
    SpreadsheetApp.getUi().alert("daily list 的" + sectionLabel + "里没有识别到车次" +
      (mode === "nj600" ? "，Pickup Rate 里也没有运力部卡车" : "") + "。Records 没有改动。");
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
  if (mode === "nj100936") message += "\n来源：E 列【NJ100提货】/【NJ936提货】。";
  else message += "\n来源：A-D（Han Express / Inspire / FM / SD；A 列仅【NJ600提货】）。";
  if (yunli.length) message += "\n其中运力部卡车 " + yunli.length + " 条，没有这车时删掉对应行。";
  message += "\n板数留空。";
  if (missingNames) message += "\n" + missingNames + " 个提货点在 Pickup Rate 里找不到，Origin Name 已标黄。";
  if (missingRates) message += "\n" + missingRates + " 车在 Pickup Rate 里没有对应价格，RATE 已留空并标黄。";
  if (parsed.cargoVans) message += "\n跳过 " + parsed.cargoVans + " 行 Cargo Van，这条规则还没定。";
  if (parsed.extraStops) message += "\n有 " + parsed.extraStops + " 车超过 3 个提货点，暂未支持完整写入（不能直接忽略），目前只写了前 3 个。";
  if (parsed.skipped && parsed.skipped.length) {
    message += "\n有 " + parsed.skipped.length + " 行编号车次未能识别，已跳过：";
    parsed.skipped.slice(0, 5).forEach(function (line) {
      message += "\n- " + line;
    });
    if (parsed.skipped.length > 5) message += "\n- …共 " + parsed.skipped.length + " 行";
  }
  applyOriginNameDropdowns1025_();
  SpreadsheetApp.getUi().alert(message);
}

function originNameChoices1025_() {
  const rate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Pickup Rate");
  if (!rate || rate.getLastRow() < 2) return [];
  const names = rate.getRange(2, 6, rate.getLastRow() - 1, 1).getDisplayValues()
    .map(function (row) { return String(row[0] || "").trim(); })
    .filter(function (name) { return name; });
  const extras = ["TT NJ1001", "TTNJ1001", "NJTT1001", "NJ1001"];
  extras.forEach(function (name) {
    if (names.indexOf(name) === -1) names.push(name);
  });
  return names;
}

function originNameDropdownRule1025_(names) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(names || originNameChoices1025_(), true)
    .setAllowInvalid(false)
    .build();
}

function applyOriginNameDropdowns1025_() {
  const records = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Records");
  if (!records) return;
  const endRow = originFormatEndRow1025_(records);
  const height = endRow - RECORD_FIRST_ROW_1025_ + 1;
  const names = originNameChoices1025_();
  const allowed = {};
  names.forEach(function (name) { allowed[name] = true; });
  const rule = names.length ? originNameDropdownRule1025_(names) : null;
  ORIGIN_SLOTS_1025_.forEach(function (slot) {
    const range = records.getRange(RECORD_FIRST_ROW_1025_, slot.origin, height, 1);
    range.setBackground(ORIGIN_NAME_FILL_1025_);
    if (rule) range.setDataValidation(rule);
    const values = range.getDisplayValues();
    values.forEach(function (row, index) {
      const name = String(row[0] || "").trim();
      if (!name || allowed[name]) return;
      const cell = records.getRange(RECORD_FIRST_ROW_1025_ + index, slot.origin);
      cell.clearDataValidations();
      cell.setNumberFormat("@");
      cell.setBackground(HIGHLIGHT_1025_);
    });
  });
  applyOriginPalletFormats1025_(records);
}

function originFormatEndRow1025_(sheet) {
  const last = Math.max(sheet.getLastRow(), RECORD_FIRST_ROW_1025_);
  const values = sheet.getRange(RECORD_FIRST_ROW_1025_, 1, last - RECORD_FIRST_ROW_1025_ + 1, 1).getDisplayValues();
  let lastRecord = RECORD_FIRST_ROW_1025_ - 1;
  values.forEach(function (row, index) {
    if (String(row[0] || "").trim()) lastRecord = RECORD_FIRST_ROW_1025_ + index;
  });
  return Math.max(lastRecord + 30, RECORD_FIRST_ROW_1025_ + 30);
}

function applyOriginPalletFormats1025_(sheet) {
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Records");
  if (!sheet) return;
  const endRow = originFormatEndRow1025_(sheet);
  const height = endRow - RECORD_FIRST_ROW_1025_ + 1;
  ORIGIN_SLOTS_1025_.forEach(function (slot) {
    sheet.getRange(RECORD_FIRST_ROW_1025_, slot.pallet, height, 1)
      .setNumberFormat("@")
      .setHorizontalAlignment("center")
      .setBackground(ORIGIN_PALLET_FILL_1025_);
  });
}

function readDailyListLines1025_(sheet, mode) {
  const last = Math.max(sheet.getLastRow(), DAILY_PASTE_START_ROW);
  const height = last - DAILY_PASTE_START_ROW + 1;
  const lines = [];
  // defaultVehicle26: FM (col C) and SD (col D) lists never write truck size.
  if (mode === "nj100936") {
    const values = sheet.getRange(DAILY_PASTE_START_ROW, 5, height, 1).getDisplayValues();
    values.forEach(function (row) {
      const cleaned = stripPastePlaceholder_(row[0]);
      if (!cleaned || isPastePlaceholder_(cleaned)) return;
      extractPickupSectionLines1025_(cleaned, /【\s*NJ100\s*提货\s*】|【\s*NJ936\s*提货\s*】/)
        .forEach(function (line) {
          lines.push({ text: line, defaultVehicle26: false });
        });
    });
    return lines;
  }
  const values = sheet.getRange(DAILY_PASTE_START_ROW, 1, height, 4).getDisplayValues();
  values.forEach(function (row) {
    row.forEach(function (cell, colIndex) {
      const cleaned = stripPastePlaceholder_(cell);
      if (!cleaned || isPastePlaceholder_(cleaned)) return;
      const defaultVehicle26 = colIndex === 2 || colIndex === 3;
      if (colIndex === 0) {
        extractPickupSectionLines1025_(cleaned, /【\s*NJ600\s*提货\s*】/)
          .forEach(function (line) {
            lines.push({ text: line, defaultVehicle26: false });
          });
        return;
      }
      cleaned.split(/\r?\n/).forEach(function (line) {
        const text = String(line || "").trim();
        if (text) lines.push({ text: text, defaultVehicle26: defaultVehicle26 });
      });
    });
  });
  return lines;
}

function extractPickupSectionLines1025_(text, sectionPattern) {
  const lines = [];
  let inSection = false;
  String(text || "").split(/\r?\n/).forEach(function (raw) {
    const line = String(raw || "").trim();
    if (!line) return;
    if (sectionPattern.test(line)) {
      inSection = true;
      return;
    }
    if (/🌟\s*【/.test(line) || /^【[^】]+】/.test(line)) {
      inSection = false;
      return;
    }
    if (inSection) lines.push(line);
  });
  return lines;
}

function parseTruckLines1025_(lines) {
  const trucks = [];
  let cargoVans = 0;
  let extraStops = 0;
  const skipped = [];
  lines.forEach(function (item) {
    const text = typeof item === "string" ? item : (item && item.text);
    const defaultVehicle26 = !!(item && item.defaultVehicle26);
    const cleaned = String(text || "").replace(/✅/g, "").trim();
    if (!cleaned || cleaned.indexOf("❌") !== -1) return;
    if (/^\d+\.\s*cargo\s*van\b/i.test(cleaned)) {
      cargoVans++;
      return;
    }
    if (!/^\d+\./.test(cleaned)) return;
    let truck = parseTruckLine1025_(cleaned) ||
      parseFmCharterLine1025_(cleaned) ||
      parseLooseTruckLine1025_(cleaned);
    if (!truck) {
      skipped.push(cleaned);
      return;
    }
    if (defaultVehicle26 && !truck.vehicleKey) {
      truck.vehicleKey = "26";
      truck.vehicleType = "26FT";
    }
    if (mergeFmCharterTruck1025_(trucks, truck)) {
      if (trucks[trucks.length - 1].extra) extraStops++;
      return;
    }
    if (truck.extra) extraStops++;
    trucks.push(truck);
  });
  return { trucks: trucks, cargoVans: cargoVans, extraStops: extraStops, skipped: skipped };
}

// JD NJ1570 / Yanwen / SF NJ122 charter is one truck even when the FM list
 // splits them across lines like "1. JD NJ1570-EWR" and "2. SF NJ122-EWR".
function isFmCharterTruck1025_(truck) {
  return !!(truck && truck.stops && truck.stops.length &&
    truck.stops.every(looksLikeFmCharterStop1025_));
}

function mergeFmCharterTruck1025_(trucks, truck) {
  if (!isFmCharterTruck1025_(truck) || !trucks.length) return false;
  const prev = trucks[trucks.length - 1];
  if (!isFmCharterTruck1025_(prev)) return false;
  truck.stops.forEach(function (name) {
    const exists = prev.stops.some(function (stop) {
      return compactStop_(stop) === compactStop_(name);
    });
    if (exists) return;
    if (prev.stops.length < 3) prev.stops.push(name);
    else prev.extra = true;
  });
  prev.vehicleKey = "26";
  prev.vehicleType = "26FT";
  prev.dest = prev.dest || truck.dest || "ewr";
  return true;
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

// Numbered pickup lines without 26'| / 53'|, e.g. "3. Staci NJ10-EWR".
// Vehicle is left empty and filled from the matching Pickup Rate row when possible.
function parseLooseTruckLine1025_(line) {
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
  if (!stops.length || !dest) return null;
  if (stops.every(looksLikeFmCharterStop1025_)) return null;
  return {
    vehicleKey: "",
    vehicleType: "",
    dest: dest,
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
    key === "ewr936" || key === "nj100" || key === "nj600" || key === "nj936";
}

function destKey1025_(name) {
  const key = compactStop_(name);
  if (key === "jfk" || key === "jfk175") return "jfk";
  // Temporarily treat NJ100 / NJ936 the same as NJ600 for rate matching.
  if (key === "ewr" || key === "ewr600" || key === "nj600" ||
      key === "nj100" || key === "nj936" || key === "ewr936") return "ewr";
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
  // F name, G address, H 提货时间, I 关门时间(不用), J 提货公司, K Partner ID, L Partner Name
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

function keepsListSpelling1025_(name) {
  const raw = compactStop_(name);
  return raw === "nj1001" || raw === "ttnj1001" || raw === "njtt1001";
}

function lookupMerchant1025_(name, byCompact) {
  const raw = compactStop_(name);
  if (!raw) return null;
  if (byCompact[raw]) return byCompact[raw];
  if (raw === "willownj3509" && byCompact.wilownj3509) return byCompact.wilownj3509;
  if ((raw === "ttnj1001" || raw === "njtt1001") && byCompact.nj1001) return byCompact.nj1001;
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
  const vehicleKey = truck.vehicleKey;
  let hits = rates.filter(function (rate) {
    if (rate.special) return false;
    if (vehicleKey && rate.vehicle !== vehicleKey) return false;
    if (rate.dest !== truck.dest) return false;
    if (rate.keys.length !== keys.length) return false;
    for (let i = 0; i < keys.length; i++) {
      if (rate.keys[i] !== keys[i]) return false;
    }
    return true;
  });
  if (!vehicleKey && hits.length > 1) {
    const only26 = hits.filter(function (rate) { return rate.vehicle === "26"; });
    hits = only26.length ? only26 : hits;
  }
  if (hits.length === 1) return hits[0];
  if (vehicleKey && vehicleKey !== "26") return null;
  if (!vehicleKey) return null;
  if (truck.dest && truck.dest !== "ewr") return null;
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
  row[3] = truck.vehicleType || "";
  row[6] = truck.stops.length;
  const rate = findPickupRate1025_(truck, byCompact, rates);
  if (rate && rate.price !== "") {
    row[4] = rate.price;
    row[5] = rate.carrier;
    if (!row[3] && rate.vehicle) row[3] = rate.vehicle + "FT";
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
    row[slot.origin - 1] = keepsListSpelling1025_(stopName) ? stopName : merchant.name;
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

function replaceRecordRows1025_(sheet, date, built) {
  const start = nextRecordStart1025_(sheet);
  const range = sheet.getRange(start, 1, built.length, RECORD_COLS_1025_);
  range.setBackground(null);
  ORIGIN_SLOTS_1025_.forEach(function (slot) {
    sheet.getRange(start, slot.origin, built.length, 1).clearDataValidations();
  });
  sheet.getRange(start, 1, built.length, 1).setNumberFormat("@");
  applyRecordAmountFormats1025_(sheet);
  range.setValues(built.map(function (item) { return item.values; }));
  const originRule = originNameDropdownRule1025_();
  built.forEach(function (item, index) {
    const row = start + index;
    if (item.missingRate) sheet.getRange(row, 5).setBackground(HIGHLIGHT_1025_);
    const missing = {};
    item.missingNameCols.forEach(function (col) { missing[col] = true; });
    ORIGIN_SLOTS_1025_.forEach(function (slot) {
      const cell = sheet.getRange(row, slot.origin);
      const name = String(item.values[slot.origin - 1] || "").trim();
      if (!name) {
        cell.setDataValidation(originRule);
        cell.setBackground(ORIGIN_NAME_FILL_1025_);
        return;
      }
      if (missing[slot.origin]) {
        cell.clearDataValidations();
        cell.setNumberFormat("@");
        cell.setBackground(HIGHLIGHT_1025_);
        return;
      }
      cell.setDataValidation(originRule);
      cell.setBackground(ORIGIN_NAME_FILL_1025_);
    });
  });
  applyOriginPalletFormats1025_(sheet);
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
