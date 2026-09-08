/**
 * Search Lift (Wordstat) внутри меню «Автоматизация».
 * Кнопка-рисунок на листе может вызывать fillWordstat.
 */

var WORDSTAT_URL = 'https://searchapi.api.cloud.yandex.net/v2/wordstat/dynamics';
var DEFAULT_REGION = '225';
var DEFAULT_SHEET = 'Wordstat';
var REQUEST_PAUSE_MS = 300;
var MAX_RETRIES = 3;
var PROTECT_PREFIX = 'Отчётность:';
var HEADER_SCAN_ROWS = 15;
var MAX_MONTHS = 48;

var MONTH_NAMES = [
  ['январ', 1],
  ['феврал', 2],
  ['март', 3],
  ['апрел', 4],
  ['май', 5],
  ['июл', 7],
  ['июн', 6],
  ['август', 8],
  ['сентябр', 9],
  ['октябр', 10],
  ['ноябр', 11],
  ['декабр', 12]
];

var MONTH_LABELS = [
  'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
  'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'
];

var SKIP_PHRASE = /итог|соотношен|динамик|график|сравнить|шаблон|как заполн|легенд|показател|комментар|бренд[оа]вый запрос|небренд[оа]вый запрос|wordstat|спрос\s*[—\-]/i;
var SECTION_HEADER = /^(не)?бренд[оа]вый\s+запрос\b/i;
var TOTAL_ROW = /итог/i;

function showSearchLiftSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SearchLiftSidebar')
    .setTitle('Search Lift')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

function fillWordstat() {
  var lift = loadLiftConfig_();
  if (!lift) {
    showSearchLiftSidebar();
    return;
  }
  runSearchLift(lift);
}

function fillWordstatOverwrite() {
  fillWordstat();
}

function getSearchLiftState() {
  var s = getSettings_();
  var now = new Date();
  var y = now.getFullYear();
  var years = [];
  for (var i = y - 6; i <= y + 2; i++) years.push(i);
  var lift = loadLiftConfig_() || {
    fromMonth: 7,
    fromYear: y - 1,
    toMonth: now.getMonth() + 1,
    toYear: y,
    pairs: [{ aMonth: 7, aYear: y - 1, bMonth: 7, bYear: y }]
  };
  return {
    settingsOk: !!(s.apiKey && sFolderOk_(s.folderId)),
    years: years,
    lift: lift
  };
}

function showWordstatSettings() {
  var s = getSettings_();
  var html = HtmlService.createHtmlOutput(settingsHtml_(s))
    .setWidth(460)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, 'Настройки Wordstat API');
}

function saveWordstatSettings(apiKey, folderId, regionId, sheetName) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperties({
    WORDSTAT_API_KEY: String(apiKey || '').trim(),
    WORDSTAT_FOLDER_ID: String(folderId || '').trim(),
    WORDSTAT_REGION_ID: String(regionId || DEFAULT_REGION).trim() || DEFAULT_REGION,
    WORDSTAT_SHEET_NAME: String(sheetName || '').trim()
  });
  return 'Сохранено';
}

function testWordstatConnection() {
  var s = getSettings_();
  if (!s.apiKey || !s.folderId) {
    throw new Error('Сначала сохраните API-ключ и Folder ID.');
  }
  var now = new Date();
  var from = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() - 1, 1));
  var to = lastDayUtc_(from.getUTCFullYear(), from.getUTCMonth());
  var res = fetchDynamics_('яндекс', from, to, s);
  var n = (res.results || []).length;
  return 'Ок. Wordstat ответил, точек: ' + n + '. Регион ' + s.regionId + '.';
}

/**
 * Главный прогон Search Lift из сайдбара.
 * payload: { fromMonth, fromYear, toMonth, toYear, pairs: [{aMonth,aYear,bMonth,bYear}] }
 */
function runSearchLift(payload) {
  var settings = getSettings_();
  if (!settings.apiKey || !sFolderOk_(settings.folderId)) {
    throw new Error('Сначала откройте Настройки Wordstat API и укажите ключ и Folder ID.');
  }

  var spec = normalizeLiftPayload_(payload);
  saveLiftConfig_(spec);

  var sheet = getTargetSheet_(settings.sheetName);
  if (!sheet) {
    throw new Error('Не найден лист. Откройте нужный лист или укажите имя в настройках.');
  }

  var headerRow = detectHeaderRow_(sheet);
  var months = monthsInRange_(spec.fromYear, spec.fromMonth, spec.toYear, spec.toMonth);
  var layout = rebuildHeader_(sheet, headerRow, months, spec.pairs);

  var values = sheet.getDataRange().getValues();
  var jobs = collectPhraseJobs_(values, headerRow);

  if (!jobs.length) {
    throw new Error('В колонке A нет фраз. Напишите запросы внутри блока (между заголовком секции и ИТОГО).');
  }

  var ss = SpreadsheetApp.getActive();
  var filled = 0;
  var errors = [];

  for (var i = 0; i < jobs.length; i++) {
    var job = jobs[i];
    ss.toast((i + 1) + '/' + jobs.length + '  ' + job.phrase, 'Wordstat', 8);
    SpreadsheetApp.flush();
    try {
      var map = fetchMonthCounts_(job.phrase, months, settings);
      var rowVals = [];
      for (var m = 0; m < months.length; m++) {
        var key = months[m].year + '-' + pad2_(months[m].month);
        var n = map[key];
        rowVals.push(isFiniteNumber_(n) ? n : 0);
      }
      var row1 = job.row + 1;
      var col1 = layout.monthStartCol + 1;
      sheet.getRange(row1, col1, 1, months.length).setValues([rowVals]);
      filled += months.length;
    } catch (err) {
      errors.push(job.phrase + ': ' + err.message);
      var zeros = [];
      for (var z = 0; z < months.length; z++) zeros.push(0);
      sheet.getRange(job.row + 1, layout.monthStartCol + 1, 1, months.length).setValues([zeros]);
    }
    Utilities.sleep(REQUEST_PAUSE_MS);
  }

  writeTotalsAndDeltas_(sheet, headerRow, layout, months, spec.pairs);
  applyNumberFormats_(sheet, headerRow, layout, months.length, spec.pairs.length);
  protectTemplateOnSheet_(sheet, headerRow, layout);

  var msg = 'Готово. Фраз: ' + jobs.length + ', ячеек месяцев: ' + filled + '.';
  if (errors.length) {
    msg += '\nОшибки API (в эти строки записаны нули), ' + errors.length + ':\n' + errors.slice(0, 8).join('\n');
    if (errors.length > 8) msg += '\n…';
  }
  return msg;
}

function protectTemplate() {
  var settings = getSettings_();
  var sheet = getTargetSheet_(settings.sheetName);
  if (!sheet) {
    throw new Error('Не найден лист.');
  }
  var headerRow = detectHeaderRow_(sheet);
  var parsed = parseHeaderRow_(sheet.getRange(headerRow + 1, 1, 1, sheet.getLastColumn() || 2).getValues()[0]);
  if (!parsed.months.length) {
    throw new Error('Сначала постройте шапку Search Lift (период в панели).');
  }
  var layout = {
    monthStartCol: parsed.months[0].col,
    lastCol: (parsed.deltas.length ? parsed.deltas[parsed.deltas.length - 1].col : parsed.months[parsed.months.length - 1].col)
  };
  protectTemplateOnSheet_(sheet, headerRow, layout);
  return 'Шаблон защищён: шапка, столбцы Δ и строки ИТОГО. Фразы в колонке A можно править (будет предупреждение, если задеть защиту).';
}

function normalizeLiftPayload_(payload) {
  var p = payload || {};
  var fromMonth = Number(p.fromMonth);
  var fromYear = Number(p.fromYear);
  var toMonth = Number(p.toMonth);
  var toYear = Number(p.toYear);
  if (!validMonthYear_(fromMonth, fromYear) || !validMonthYear_(toMonth, toYear)) {
    throw new Error('Укажите период: месяц и год «с» и «по».');
  }
  if (fromYear > toYear || (fromYear === toYear && fromMonth > toMonth)) {
    throw new Error('Начало периода позже конца.');
  }
  var months = monthsInRange_(fromYear, fromMonth, toYear, toMonth);
  var pairsIn = p.pairs || [];
  var pairs = [];
  for (var i = 0; i < pairsIn.length; i++) {
    var aM = Number(pairsIn[i].aMonth);
    var aY = Number(pairsIn[i].aYear);
    var bM = Number(pairsIn[i].bMonth);
    var bY = Number(pairsIn[i].bYear);
    if (!validMonthYear_(aM, aY) || !validMonthYear_(bM, bY)) {
      throw new Error('В сравнении ' + (i + 1) + ' укажите оба месяца с годом.');
    }
    if (!monthInList_(months, aY, aM) || !monthInList_(months, bY, bM)) {
      throw new Error('Сравнение ' + (i + 1) + ': оба месяца должны входить в выбранный период.');
    }
    pairs.push({ aMonth: aM, aYear: aY, bMonth: bM, bYear: bY });
  }
  return {
    fromMonth: fromMonth,
    fromYear: fromYear,
    toMonth: toMonth,
    toYear: toYear,
    pairs: pairs
  };
}

function rebuildHeader_(sheet, headerRow, months, pairs) {
  var row1 = headerRow + 1;
  var startCol = 2;
  var lastCol = startCol + months.length + pairs.length - 1;
  if (lastCol < startCol) lastCol = startCol;
  var used = Math.max(sheet.getLastColumn(), lastCol + 8, 20);
  var lastDataRow = Math.max(sheet.getLastRow(), 1);

  try {
    sheet.getRange(row1, startCol, 1, used).breakApart();
  } catch (e1) {}

  if (used > lastCol) {
    sheet.getRange(1, lastCol + 1, lastDataRow, used - lastCol).clearContent();
    sheet.getRange(row1, lastCol + 1, 1, used - lastCol).clearFormat();
  }

  var monthLabels = [];
  for (var i = 0; i < months.length; i++) {
    monthLabels.push(formatMonthHeader_(months[i].year, months[i].month));
  }
  if (monthLabels.length) {
    sheet.getRange(row1, startCol, 1, monthLabels.length).setValues([monthLabels]);
    sheet.getRange(row1, startCol, 1, monthLabels.length)
      .setFontWeight('bold')
      .setBackground('#fff2cc')
      .setWrap(true);
  }

  var deltaLabels = [];
  for (var d = 0; d < pairs.length; d++) {
    deltaLabels.push(formatDeltaHeader_(pairs[d]));
  }
  if (deltaLabels.length) {
    var deltaStart = startCol + months.length;
    sheet.getRange(row1, deltaStart, 1, deltaLabels.length).setValues([deltaLabels]);
    sheet.getRange(row1, deltaStart, 1, deltaLabels.length)
      .setFontWeight('bold')
      .setBackground('#d0e2ff')
      .setWrap(true);
  }

  try {
    var title = String(sheet.getRange(1, 1).getValue() || '');
    if (title) {
      sheet.getRange(1, 1, 1, Math.max(used, lastCol)).breakApart();
      sheet.getRange(1, 1, 1, lastCol).merge().setValue(title).setFontWeight('bold');
    }
  } catch (e2) {}

  return {
    monthStartCol: startCol - 1,
    lastCol: lastCol - 1,
    deltaStartCol: startCol - 1 + months.length
  };
}

function collectPhraseJobs_(values, headerRow) {
  var jobs = [];
  for (var r = 0; r < values.length; r++) {
    if (r === headerRow) continue;
    var phrase = String(values[r][0] || '').trim();
    if (!isPhrase_(phrase)) continue;
    jobs.push({ row: r, phrase: phrase });
  }
  return jobs;
}

function writeTotalsAndDeltas_(sheet, headerRow, layout, months, pairs) {
  var values = sheet.getDataRange().getValues();
  var blocks = findBlocks_(values, headerRow);
  var lastRow = values.length;
  var monthStart = layout.monthStartCol;
  var nMonths = months.length;
  var nDeltas = pairs.length;

  for (var b = 0; b < blocks.length; b++) {
    var block = blocks[b];
    if (!block.phraseRows.length) continue;
    var first = block.phraseRows[0] + 1;
    var last = block.phraseRows[block.phraseRows.length - 1] + 1;
    if (block.totalRow != null) {
      var sums = [];
      for (var m = 0; m < nMonths; m++) {
        var L = colLetter_(monthStart + m + 1);
        sums.push('=SUM(' + L + first + ':' + L + last + ')');
      }
      sheet.getRange(block.totalRow + 1, monthStart + 1, 1, nMonths).setFormulas([sums]);
    }
  }

  if (!nDeltas) return;

  var deltaStart = layout.deltaStartCol;
  var rowsToFill = {};
  for (var r = 0; r < lastRow; r++) {
    if (r === headerRow) continue;
    var a = String(values[r][0] || '').trim();
    if (!a) continue;
    if (isPhrase_(a) || TOTAL_ROW.test(a)) rowsToFill[r] = true;
  }

  var rowIndexes = Object.keys(rowsToFill);
  for (var i = 0; i < rowIndexes.length; i++) {
    var row0 = Number(rowIndexes[i]);
    var row1 = row0 + 1;
    var formulas = [];
    for (var p = 0; p < nDeltas; p++) {
      var pair = pairs[p];
      var colA = monthStart + indexOfMonth_(months, pair.aYear, pair.aMonth);
      var colB = monthStart + indexOfMonth_(months, pair.bYear, pair.bMonth);
      var la = colLetter_(colA + 1);
      var lb = colLetter_(colB + 1);
      formulas.push('=IF(' + lb + row1 + '=0,"",(' + la + row1 + '-' + lb + row1 + ')/' + lb + row1 + ')');
    }
    sheet.getRange(row1, deltaStart + 1, 1, nDeltas).setFormulas([formulas]);
  }
}

function applyNumberFormats_(sheet, headerRow, layout, nMonths, nDeltas) {
  var lastRow = Math.max(sheet.getLastRow(), headerRow + 2);
  var startRow = headerRow + 2;
  var nRows = lastRow - startRow + 1;
  if (nRows < 1) return;
  if (nMonths) {
    sheet.getRange(startRow, layout.monthStartCol + 1, nRows, nMonths).setNumberFormat('#,##0');
  }
  if (nDeltas) {
    sheet.getRange(startRow, layout.deltaStartCol + 1, nRows, nDeltas).setNumberFormat('0.0%');
  }
}

function protectTemplateOnSheet_(sheet, headerRow, layout) {
  clearScriptProtections_(sheet);
  var lastRow = Math.max(sheet.getLastRow(), headerRow + 2);
  var lastCol = layout.lastCol + 1;

  addWarnProtect_(sheet.getRange(1, 1, headerRow + 1, lastCol), 'шапка');

  if (layout.deltaStartCol <= layout.lastCol) {
    var nDeltaCols = layout.lastCol - layout.deltaStartCol + 1;
    var nProtRows = lastRow - headerRow - 1;
    if (nDeltaCols > 0 && nProtRows > 0) {
      addWarnProtect_(sheet.getRange(headerRow + 2, layout.deltaStartCol + 1, nProtRows, nDeltaCols), 'дельты');
    }
  }

  var values = sheet.getDataRange().getValues();
  for (var r = 0; r < values.length; r++) {
    if (TOTAL_ROW.test(String(values[r][0] || ''))) {
      addWarnProtect_(sheet.getRange(r + 1, 1, 1, lastCol), 'итог');
    }
  }
}

function addWarnProtect_(range, tag) {
  var prot = range.protect().setDescription(PROTECT_PREFIX + ' ' + tag);
  prot.setWarningOnly(true);
}

function clearScriptProtections_(sheet) {
  var list = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < list.length; i++) {
    if (String(list[i].getDescription() || '').indexOf(PROTECT_PREFIX) === 0) {
      list[i].remove();
    }
  }
}

function findBlocks_(values, headerRow) {
  var blocks = [];
  var r = 0;
  while (r < values.length) {
    var label = String(values[r][0] || '').trim();
    if (SECTION_HEADER.test(label)) {
      var phraseRows = [];
      var totalRow = null;
      r++;
      while (r < values.length) {
        var t = String(values[r][0] || '').trim();
        if (SECTION_HEADER.test(t) && r !== headerRow) break;
        if (TOTAL_ROW.test(t)) {
          totalRow = r;
          r++;
          break;
        }
        if (isPhrase_(t)) phraseRows.push(r);
        r++;
      }
      blocks.push({ phraseRows: phraseRows, totalRow: totalRow });
      continue;
    }
    r++;
  }
  if (!blocks.length) {
    var phrases = [];
    var total = null;
    for (var i = 0; i < values.length; i++) {
      if (i === headerRow) continue;
      var a = String(values[i][0] || '').trim();
      if (TOTAL_ROW.test(a) && total == null) total = i;
      else if (isPhrase_(a)) phrases.push(i);
    }
    if (phrases.length) blocks.push({ phraseRows: phrases, totalRow: total });
  }
  return blocks;
}

function detectHeaderRow_(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 2);
  var n = Math.min(HEADER_SCAN_ROWS, Math.max(sheet.getLastRow(), 2));
  var values = sheet.getRange(1, 1, n, lastCol).getValues();
  var best = 1;
  var bestN = -1;
  for (var r = 0; r < values.length; r++) {
    var parsed = parseHeaderRow_(values[r]);
    if (parsed.months.length > bestN) {
      bestN = parsed.months.length;
      best = r;
    }
  }
  if (bestN < 1) return 1;
  return best;
}

function parseHeaderRow_(row) {
  var months = [];
  var deltas = [];
  for (var c = 1; c < row.length; c++) {
    var delta = parseDeltaHeader_(row[c]);
    if (delta) {
      delta.col = c;
      deltas.push(delta);
      continue;
    }
    var month = parseMonthHeader_(row[c]);
    if (month) {
      month.col = c;
      months.push(month);
    }
  }
  return { months: months, deltas: deltas };
}

function parseMonthHeader_(raw) {
  var text = String(raw || '').trim();
  if (!text) return null;
  if (/[Δδ]/.test(text) || text.indexOf('/') !== -1) return null;
  if (/[%]/.test(text)) return null;
  var low = text.toLowerCase().replace(/ё/g, 'е');
  if (/итог|соотношен|динамик|комментар/.test(low)) return null;

  var month = null;
  for (var i = 0; i < MONTH_NAMES.length; i++) {
    if (low.indexOf(MONTH_NAMES[i][0]) !== -1) {
      month = MONTH_NAMES[i][1];
      break;
    }
  }
  var year = null;
  var yyyy = low.match(/\b(20\d{2})\b/);
  var yy = low.match(/\b(\d{2})\b/);
  if (yyyy) year = Number(yyyy[1]);
  else if (yy) year = 2000 + Number(yy[1]);
  if (!month || !year) return null;
  if (year < 2018 || year > 2100) return null;
  return { year: year, month: month };
}

function parseDeltaHeader_(raw) {
  var text = String(raw || '').trim();
  if (!text) return null;
  if (!/[Δδ]/.test(text) && text.indexOf('/') === -1) return null;
  if (!/[Δδ]/.test(text) && !/\//.test(text)) return null;
  if (!/[Δδ]/.test(text)) return null;
  var parts = text.split('/');
  if (parts.length < 2) return null;
  var a = parseMonthHeader_(parts[0].replace(/[Δδ]/g, ''));
  var b = parseMonthHeader_(parts.slice(1).join('/'));
  if (!a || !b) return null;
  return { a: a, b: b };
}

function isPhrase_(phrase) {
  if (!phrase) return false;
  if (SKIP_PHRASE.test(phrase)) return false;
  if (SECTION_HEADER.test(phrase)) return false;
  if (parseMonthHeader_(phrase)) return false;
  if (phrase.length > 400) return false;
  return true;
}

function fetchMonthCounts_(phrase, months, settings) {
  var minY = months[0].year;
  var minM = months[0].month;
  var maxY = months[0].year;
  var maxM = months[0].month;
  for (var i = 1; i < months.length; i++) {
    if (months[i].year < minY || (months[i].year === minY && months[i].month < minM)) {
      minY = months[i].year;
      minM = months[i].month;
    }
    if (months[i].year > maxY || (months[i].year === maxY && months[i].month > maxM)) {
      maxY = months[i].year;
      maxM = months[i].month;
    }
  }
  var from = new Date(Date.UTC(minY, minM - 1, 1));
  var to = lastDayUtc_(maxY, maxM - 1);
  var data = fetchDynamics_(phrase, from, to, settings);
  var map = {};
  var results = data.results || data.dynamics || [];
  for (var j = 0; j < results.length; j++) {
    var item = results[j];
    var d = parseApiDate_(item.date);
    if (!d) continue;
    var key = d.getUTCFullYear() + '-' + pad2_(d.getUTCMonth() + 1);
    var count = item.count;
    if (count === undefined || count === null) count = item.value;
    var n = Number(count);
    map[key] = isFiniteNumber_(n) ? n : 0;
  }
  return map;
}

function fetchDynamics_(phrase, fromDate, toDate, settings) {
  var payload = {
    phrase: phrase,
    period: 'PERIOD_MONTHLY',
    fromDate: iso_(fromDate),
    toDate: isoEnd_(toDate),
    regions: [String(settings.regionId || DEFAULT_REGION)],
    folderId: settings.folderId
  };
  var lastErr = 'неизвестная ошибка';
  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    var resp = UrlFetchApp.fetch(WORDSTAT_URL, {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: { Authorization: 'Api-Key ' + settings.apiKey },
      payload: JSON.stringify(payload)
    });
    var code = resp.getResponseCode();
    var body = resp.getContentText();
    if (code >= 200 && code < 300) {
      return JSON.parse(body);
    }
    lastErr = 'HTTP ' + code + ' ' + clip_(body, 240);
    if (code === 429 || code >= 500) {
      Utilities.sleep(500 * attempt);
      continue;
    }
    throw new Error(lastErr);
  }
  throw new Error(lastErr);
}

function getTargetSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  if (name) {
    var named = ss.getSheetByName(name);
    if (named) return named;
  }
  var active = ss.getActiveSheet();
  if (active) return active;
  var fallback = ss.getSheetByName(DEFAULT_SHEET);
  return fallback || ss.getSheets()[0];
}

function getSettings_() {
  var p = PropertiesService.getDocumentProperties();
  return {
    apiKey: p.getProperty('WORDSTAT_API_KEY') || '',
    folderId: p.getProperty('WORDSTAT_FOLDER_ID') || '',
    regionId: p.getProperty('WORDSTAT_REGION_ID') || DEFAULT_REGION,
    sheetName: p.getProperty('WORDSTAT_SHEET_NAME') || ''
  };
}

function loadLiftConfig_() {
  var raw = PropertiesService.getDocumentProperties().getProperty('LIFT_CONFIG');
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function saveLiftConfig_(spec) {
  PropertiesService.getDocumentProperties().setProperty('LIFT_CONFIG', JSON.stringify(spec));
}

function sFolderOk_(id) {
  return !!(id && String(id).trim());
}

function isFiniteNumber_(n) {
  return typeof n === 'number' && isFinite(n);
}

function validMonthYear_(month, year) {
  return month >= 1 && month <= 12 && year >= 2018 && year <= 2100;
}

function monthsInRange_(y1, m1, y2, m2) {
  var out = [];
  var y = y1;
  var m = m1;
  while (y < y2 || (y === y2 && m <= m2)) {
    out.push({ year: y, month: m });
    m++;
    if (m > 12) {
      m = 1;
      y++;
    }
    if (out.length > MAX_MONTHS) {
      throw new Error('Слишком длинный период (больше ' + MAX_MONTHS + ' месяцев).');
    }
  }
  return out;
}

function monthInList_(months, year, month) {
  return indexOfMonth_(months, year, month) !== -1;
}

function indexOfMonth_(months, year, month) {
  for (var i = 0; i < months.length; i++) {
    if (months[i].year === year && months[i].month === month) return i;
  }
  return -1;
}

function formatMonthHeader_(year, month) {
  return MONTH_LABELS[month - 1] + ' ' + String(year).slice(-2);
}

function formatDeltaHeader_(pair) {
  var a = MONTH_LABELS[pair.aMonth - 1].toLowerCase() + ' ' + String(pair.aYear).slice(-2);
  var b = MONTH_LABELS[pair.bMonth - 1].toLowerCase() + ' ' + String(pair.bYear).slice(-2);
  return 'Δ ' + a + ' / ' + b;
}

function lastDayUtc_(year, monthIndex) {
  return new Date(Date.UTC(year, monthIndex + 1, 0));
}

function iso_(d) {
  return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'00:00:00'Z'");
}

function isoEnd_(d) {
  return Utilities.formatDate(d, 'UTC', "yyyy-MM-dd'T'23:59:59'Z'");
}

function parseApiDate_(value) {
  if (!value) return null;
  if (typeof value === 'object' && value.seconds) {
    return new Date(Number(value.seconds) * 1000);
  }
  var s = String(value);
  var m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

function colLetter_(n) {
  var s = '';
  var x = n;
  while (x > 0) {
    var m = (x - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

function clip_(s, n) {
  s = String(s || '').replace(/\s+/g, ' ');
  return s.length > n ? s.slice(0, n) + '…' : s;
}

function settingsHtml_(s) {
  var api = htmlEscape_(s.apiKey);
  var folder = htmlEscape_(s.folderId);
  var region = htmlEscape_(s.regionId);
  var sheet = htmlEscape_(s.sheetName);
  return (
    '<!DOCTYPE html><html><head><base target="_top">' +
    '<style>' +
    'body{font:13px/1.4 Arial,sans-serif;padding:8px 4px;color:#222}' +
    'label{display:block;margin:10px 0 4px;font-weight:bold}' +
    'input{width:100%;box-sizing:border-box;padding:6px 8px;border:1px solid #ccc;border-radius:4px}' +
    '.hint{color:#666;font-size:12px;margin:4px 0 0}' +
    '.row{margin-top:16px;display:flex;gap:8px}' +
    'button{padding:8px 12px;border:0;border-radius:4px;cursor:pointer}' +
    '.ok{background:#1a73e8;color:#fff}' +
    '.test{background:#e8eaed}' +
    '#msg{margin-top:10px;min-height:18px}' +
    '</style></head><body>' +
    '<div class="hint">Ключ — Yandex Cloud / AI Studio. Folder ID вида b1g… Регион 225 = Россия. Имя листа можно оставить пустым — возьмётся активный.</div>' +
    '<label>API-ключ</label><input id="api" type="password" value="' + api + '">' +
    '<label>Folder ID</label><input id="folder" value="' + folder + '">' +
    '<label>Регион Wordstat</label><input id="region" value="' + region + '">' +
    '<label>Имя листа (необязательно)</label><input id="sheet" value="' + sheet + '">' +
    '<div class="row">' +
    '<button class="ok" onclick="save()">Сохранить</button>' +
    '<button class="test" onclick="test()">Проверить связь</button>' +
    '</div><div id="msg"></div>' +
    '<script>' +
    'function val(id){return document.getElementById(id).value;}' +
    'function save(){google.script.run.withSuccessHandler(function(t){msg(t);}).withFailureHandler(err).saveWordstatSettings(val("api"),val("folder"),val("region"),val("sheet"));}' +
    'function test(){google.script.run.withSuccessHandler(function(t){msg(t);}).withFailureHandler(err).testWordstatConnection();}' +
    'function msg(t){document.getElementById("msg").textContent=t;}' +
    'function err(e){msg(e.message||String(e));}' +
    '</script></body></html>'
  );
}

function htmlEscape_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/"/g, '&quot;')
    .replace(/</g, '&lt;');
}
