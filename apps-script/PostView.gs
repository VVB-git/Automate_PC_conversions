/**
 * Чистка сырого post-view CPM отчёта на активном листе:
 * вырезает колонки от ID компании до programmatic ID,
 * на лист «Post-view» пишет уникальные post-view с шапкой, группы по типу и счётчик PV.
 */

var PV_RESULT_SHEET = 'Post-view';
var PV_HEADER_SCAN_ROWS = 20;

function runPostViewCleanup() {
  var ui = SpreadsheetApp.getUi();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) {
    ui.alert('Обработка уже идёт. Дождитесь окончания.');
    return;
  }
  try {
    var source = SpreadsheetApp.getActiveSheet();
    if (!source) {
      throw new Error('Нет активного листа.');
    }
    if (source.getName() === PV_RESULT_SHEET) {
      throw new Error('Встаньте на лист с сырым отчётом, не на «' + PV_RESULT_SHEET + '».');
    }

    var headerInfo = pvLocateHeader_(source);
    var cut = pvFindCutRange_(headerInfo.headers);
    if (cut) {
      source.deleteColumns(cut.start + 1, cut.end - cut.start + 1);
      SpreadsheetApp.flush();
      headerInfo = pvLocateHeader_(source);
    }

    var cols = pvMapResultColumns_(headerInfo.headers);
    var lastRow = source.getLastRow();
    if (lastRow <= headerInfo.row) {
      throw new Error('На листе нет строк данных после шапки.');
    }

    var width = headerInfo.headers.length;
    var range = source.getRange(headerInfo.row + 1, 1, lastRow - headerInfo.row, width);
    var data = range.getValues();
    var shown = range.getDisplayValues();
    var result = pvFilterAndDedup_(data, shown, cols);
    pvWriteResultSheet_(result.rows, result.pvByType);

    SpreadsheetApp.getActive().toast(
      'Уникальных post-view: ' + result.rows.length,
      'Post-view',
      8
    );
  } catch (e) {
    ui.alert(e.message || String(e));
  } finally {
    lock.releaseLock();
  }
}

function pvLocateHeader_(sheet) {
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  if (lastCol < 1 || lastRow < 1) {
    throw new Error('Лист пустой. Вставьте сырой post-view отчёт.');
  }
  var scanRows = Math.min(PV_HEADER_SCAN_ROWS, lastRow);
  var values = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  var bestRow = 0;
  var bestScore = -1;
  var r;
  for (r = 0; r < values.length; r++) {
    var score = pvHeaderScore_(values[r]);
    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }
  if (bestScore < 2) {
    throw new Error(
      'Не нашёл шапку отчёта. Нужны колонки вроде «тип конверсии», «дата конверсии», «Client ID YM».'
    );
  }
  return {
    row: bestRow + 1,
    headers: values[bestRow]
  };
}

function pvHeaderScore_(row) {
  var score = 0;
  if (pvFindCol_(row, pvIsCampaignId_) >= 0) score += 1;
  if (pvFindCol_(row, pvIsProgrammaticId_) >= 0) score += 1;
  if (pvFindCol_(row, pvIsConversionName_) >= 0) score += 2;
  if (pvFindCol_(row, pvIsConversionType_) >= 0) score += 2;
  if (pvFindCol_(row, pvIsConversionDate_) >= 0) score += 2;
  if (pvFindCol_(row, pvIsLastVisitDate_) >= 0) score += 2;
  if (pvFindCol_(row, pvIsImpressionDate_) >= 0) score += 2;
  if (pvFindCol_(row, pvIsClientIdYm_) >= 0) score += 2;
  return score;
}

function pvFindCutRange_(headers) {
  var start = pvFindCol_(headers, pvIsCampaignId_);
  var end = pvFindCol_(headers, pvIsProgrammaticId_);
  if (start < 0 || end < 0) {
    return null;
  }
  if (end < start) {
    throw new Error(
      'Колонка programmatic ID стоит левее ID компании — проверьте шапку, вырез не делаю.'
    );
  }
  return { start: start, end: end };
}

function pvMapResultColumns_(headers) {
  var conversion = pvFindCol_(headers, pvIsConversionName_);
  var type = pvFindCol_(headers, pvIsConversionType_);
  var conversionDate = pvFindCol_(headers, pvIsConversionDate_);
  var lastVisit = pvFindCol_(headers, pvIsLastVisitDate_);
  var impression = pvFindCol_(headers, pvIsImpressionDate_);
  var clientId = pvFindCol_(headers, pvIsClientIdYm_);
  var missing = [];
  if (conversion < 0) missing.push('ID Конверсии');
  if (type < 0) missing.push('тип конверсии');
  if (conversionDate < 0) missing.push('дата конверсии');
  if (lastVisit < 0) missing.push('дата последнего визита');
  if (impression < 0) missing.push('дата показа');
  if (clientId < 0) missing.push('Client ID YM');
  if (missing.length) {
    throw new Error('Не найдены колонки: ' + missing.join(', ') + '.');
  }
  return {
    conversion: conversion,
    type: type,
    conversionDate: conversionDate,
    lastVisit: lastVisit,
    impression: impression,
    clientId: clientId
  };
}

function pvFilterAndDedup_(data, shown, cols) {
  shown = shown || data;
  var filterPostView = pvSheetHasPostViewTypes_(shown, cols.type);
  var byClient = {};
  var order = [];
  var pvByType = {};
  var i;
  for (i = 0; i < data.length; i++) {
    var row = data[i];
    var text = shown[i] || row;
    if (filterPostView && !pvIsPostViewType_(text[cols.type])) {
      continue;
    }
    var typeName = pvCellText_(text[cols.type] || row[cols.type]);
    if (typeName && typeName !== 'N/A') {
      if (!pvByType.hasOwnProperty(typeName)) {
        pvByType[typeName] = 0;
      }
      pvByType[typeName] += 1;
    }
    var clientId = pvCellText_(text[cols.clientId] || row[cols.clientId]);
    if (!clientId || clientId === 'N/A') {
      continue;
    }
    var stamp = pvToTime_(row[cols.conversionDate] || text[cols.conversionDate]);
    var candidate = {
      conversion: text[cols.conversion] || row[cols.conversion],
      type: text[cols.type] || row[cols.type],
      conversionDate: row[cols.conversionDate],
      lastVisit: row[cols.lastVisit],
      impression: row[cols.impression],
      clientId: text[cols.clientId] || row[cols.clientId],
      stamp: stamp,
      index: i
    };
    var prev = byClient[clientId];
    if (!prev) {
      byClient[clientId] = candidate;
      order.push(clientId);
      continue;
    }
    if (candidate.stamp > prev.stamp || (candidate.stamp === prev.stamp && candidate.index > prev.index)) {
      byClient[clientId] = candidate;
    }
  }
  var out = [];
  for (i = 0; i < order.length; i++) {
    var item = byClient[order[i]];
    out.push([
      item.conversion,
      item.type,
      item.conversionDate,
      item.lastVisit,
      item.impression,
      item.clientId
    ]);
  }
  out.sort(pvCompareResultRows_);
  return { rows: out, pvByType: pvByType };
}

var PV_RESULT_HEADERS = [
  'ID Конверсии',
  'Тип конверсии',
  'Дата конверсии',
  'Дата последнего визита',
  'Дата показа',
  'clientID YM'
];

function pvCompareResultRows_(a, b) {
  var typeA = String(a[1] == null ? '' : a[1]);
  var typeB = String(b[1] == null ? '' : b[1]);
  var byType = typeA.localeCompare(typeB, 'ru', { sensitivity: 'base' });
  if (byType !== 0) {
    return byType;
  }
  return pvToTime_(b[2]) - pvToTime_(a[2]);
}

function pvWriteResultSheet_(rows, pvByType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PV_RESULT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PV_RESULT_SHEET);
  }
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), 10);
  sheet.getRange(1, 1, lastRow, lastCol).clearContent().clearFormat();

  sheet.getRange(1, 1, 1, 6).setValues([PV_RESULT_HEADERS]);
  sheet.getRange(1, 1, 1, 6)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setFrozenRows(1);

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  var summary = pvSummaryRows_(pvByType);
  sheet.getRange(1, 8, 1, 2).setValues([['Тип конверсии', 'PV']]);
  sheet.getRange(1, 8, 1, 2)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  if (summary.length) {
    sheet.getRange(2, 8, summary.length, 2).setValues(summary);
    sheet.getRange(2, 9, summary.length, 1).setHorizontalAlignment('center');
  }

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 190);
  sheet.setColumnWidth(5, 170);
  sheet.setColumnWidth(6, 180);
  sheet.setColumnWidth(7, 24);
  sheet.setColumnWidth(8, 220);
  sheet.setColumnWidth(9, 70);
  ss.setActiveSheet(sheet);
  SpreadsheetApp.flush();
}

function pvSummaryRows_(pvByType) {
  var names = [];
  var key;
  for (key in pvByType) {
    if (pvByType.hasOwnProperty(key)) {
      names.push(key);
    }
  }
  names.sort(function(a, b) {
    return a.localeCompare(b, 'ru', { sensitivity: 'base' });
  });
  var rows = [];
  var i;
  for (i = 0; i < names.length; i++) {
    rows.push([names[i], pvByType[names[i]]]);
  }
  return rows;
}

function pvFindCol_(headers, pred) {
  var i;
  for (i = 0; i < headers.length; i++) {
    if (pred(pvNorm_(headers[i]))) {
      return i;
    }
  }
  return -1;
}

function pvNorm_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/[_./\\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function pvIsCampaignId_(n) {
  return n.indexOf('id компании') >= 0 || n.indexOf('id кампании') >= 0;
}

function pvIsProgrammaticId_(n) {
  if (n.indexOf('programmatic id') >= 0 || n === 'programmaticid') {
    return true;
  }
  return n.indexOf('программатик') >= 0 && n.indexOf('id') >= 0;
}

function pvIsConversionName_(n) {
  if (n.indexOf('кол-во') >= 0 || n.indexOf('количество') >= 0) {
    return false;
  }
  if (n.indexOf('тип') >= 0 || n.indexOf('дата') >= 0) {
    return false;
  }
  if (n.indexOf('кампан') >= 0 || n.indexOf('компани') >= 0) {
    return false;
  }
  return n.indexOf('конверси') >= 0;
}

function pvIsConversionType_(n) {
  return n.indexOf('тип конверсии') >= 0;
}

function pvIsConversionDate_(n) {
  return n.indexOf('дата конверсии') >= 0;
}

function pvIsLastVisitDate_(n) {
  if (n.indexOf('дата последнего визита') >= 0 || n.indexOf('дата последнего клика') >= 0) {
    return true;
  }
  return n.indexOf('дата последн') >= 0 && (n.indexOf('визит') >= 0 || n.indexOf('клик') >= 0);
}

function pvIsImpressionDate_(n) {
  return n.indexOf('дата показа') >= 0;
}

function pvIsClientIdYm_(n) {
  var compact = n.replace(/\s+/g, '');
  if (compact.indexOf('clientidym') >= 0 || compact.indexOf('clientidykm') >= 0) {
    return true;
  }
  if (n.indexOf('client') < 0) {
    return false;
  }
  return n.indexOf('ym') >= 0 || n.indexOf('ykm') >= 0 || n.indexOf('ям') >= 0;
}

function pvSheetHasPostViewTypes_(data, typeIdx) {
  var i;
  for (i = 0; i < data.length; i++) {
    if (pvIsPostViewType_(data[i][typeIdx])) {
      return true;
    }
  }
  return false;
}

function pvIsPostViewType_(value) {
  var n = pvNorm_(value).replace(/\s+/g, '');
  if (!n) {
    return false;
  }
  return n.indexOf('postview') >= 0 || n.indexOf('постview') >= 0 || n.indexOf('постпросмотр') >= 0;
}

function pvCellText_(value) {
  if (value == null || value === '') {
    return '';
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? '' : String(value.getTime());
  }
  return String(value).replace(/\s+/g, ' ').trim();
}

function pvToTime_(value) {
  if (value == null || value === '') {
    return 0;
  }
  if (Object.prototype.toString.call(value) === '[object Date]') {
    var t = value.getTime();
    return isNaN(t) ? 0 : t;
  }
  var s = String(value).trim();
  if (!s) {
    return 0;
  }
  var iso = Date.parse(s);
  if (!isNaN(iso)) {
    return iso;
  }
  var m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    var year = parseInt(m[3], 10);
    if (year < 100) {
      year += 2000;
    }
    var dt = new Date(
      year,
      parseInt(m[2], 10) - 1,
      parseInt(m[1], 10),
      m[4] ? parseInt(m[4], 10) : 0,
      m[5] ? parseInt(m[5], 10) : 0,
      m[6] ? parseInt(m[6], 10) : 0
    );
    var ts = dt.getTime();
    return isNaN(ts) ? 0 : ts;
  }
  return 0;
}
