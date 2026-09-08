/**
 * Запись первого листа в формате примера: П.П. / тип / дата-время / цель в ЯМ.
 */

var YM_HEADERS = ['П.П.', 'Тип конверсии', 'Дата/время', 'Цель в ЯМ'];

function removeSheetCharts_(sheet) {
  var charts = sheet.getCharts();
  var i;
  for (i = 0; i < charts.length; i++) {
    sheet.removeChart(charts[i]);
  }
}

function writeConversionSheet_(rows) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  removeSheetCharts_(sheet);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), 4);
  sheet.getRange(1, 1, lastRow, lastCol).clearContent().clearFormat();

  sheet.getRange(1, 1, 1, 4).setValues([YM_HEADERS]);
  sheet.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (rows.length) {
    var values = [];
    for (var i = 0; i < rows.length; i++) {
      values.push([
        i + 1,
        rows[i].type,
        rows[i].dateTime,
        rows[i].goalName
      ]);
    }
    sheet.getRange(2, 1, values.length, 4).setValues(values);
    sheet.getRange(2, 1, values.length, 3).setHorizontalAlignment('center');
    sheet.getRange(2, 4, values.length, 1).setHorizontalAlignment('left');
    sheet.getRange(1, 1, values.length + 1, 4)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    sheet.getRange(1, 1, 1, 4)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 520);
  sheet.setFrozenRows(1);
  SpreadsheetApp.flush();
}

function interestsHeaderRow_(level) {
  var headers = [];
  var i;
  for (i = 1; i <= level; i++) {
    headers.push('Интерес (ур. ' + i + ')');
  }
  headers.push('Посетители');
  headers.push('Аффинити-индекс, %');
  return headers;
}

function writeInterestsSheet_(report) {
  report = report || {};
  var settings = report.settings || {};
  var level = Number(report.level) || 1;
  var rows = report.rows || [];
  var headers = interestsHeaderRow_(level);
  var colCount = headers.length;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  removeSheetCharts_(sheet);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var lastCol = Math.max(sheet.getLastColumn(), colCount, 2);
  sheet.getRange(1, 1, lastRow, lastCol).clearContent().clearFormat();
  sheet.setFrozenRows(0);

  var sampledText = report.sampled
    ? ('да, доля выборки ' + formatSampleShare_(report.sampleShare))
    : 'нет';

  sheet.getRange(1, 1, 5, 2).setValues([
    ['Отчёт', 'Долгосрочные интересы'],
    ['Счётчик', String(settings.counterId || '')],
    ['Период', (settings.date1 || '') + ' — ' + (settings.date2 || '')],
    ['Метка', settings.label
      ? (settings.label + ' (' + (settings.labelField || '') + ')')
      : 'все источники'],
    ['Уровень / семплирование', String(level) + ' / ' + sampledText]
  ]);
  sheet.getRange(1, 1, 5, 1).setFontWeight('bold');

  var headerRow = 7;
  sheet.getRange(headerRow, 1, 1, colCount).setValues([headers]);
  sheet.getRange(headerRow, 1, 1, colCount)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  if (rows.length) {
    var values = [];
    var r;
    for (r = 0; r < rows.length; r++) {
      var line = [];
      var names = rows[r].names || [];
      var c;
      for (c = 0; c < level; c++) {
        line.push(names[c] || '');
      }
      line.push(rows[r].users);
      line.push(rows[r].affinity);
      values.push(line);
    }
    sheet.getRange(headerRow + 1, 1, values.length, colCount).setValues(values);
    sheet.getRange(headerRow + 1, 1, values.length, level).setHorizontalAlignment('left');
    sheet.getRange(headerRow + 1, level + 1, values.length, 2).setHorizontalAlignment('center');
    sheet.getRange(headerRow + 1, level + 1, values.length, 1).setNumberFormat('#,##0');
    sheet.getRange(headerRow + 1, level + 2, values.length, 1).setNumberFormat('0.00');
    sheet.getRange(headerRow, 1, values.length + 1, colCount)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    sheet.getRange(headerRow, 1, 1, colCount)
      .setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  }

  var widths = [240, 240, 240, 120, 160];
  var w;
  for (w = 0; w < colCount; w++) {
    sheet.setColumnWidth(w + 1, w < level ? 260 : widths[w] || 140);
  }
  sheet.setFrozenRows(headerRow);
  writeInterestsAffinityChart_(sheet, rows, level, colCount, headerRow);
  SpreadsheetApp.flush();
}

function writeInterestsAffinityChart_(sheet, rows, level, colCount, headerRow) {
  if (!rows || !rows.length) {
    return;
  }

  var chartRows = [];
  var i;
  for (i = 0; i < rows.length; i++) {
    var names = rows[i].names || [];
    var affinity = Number(rows[i].affinity);
    if (!isFinite(affinity)) {
      continue;
    }
    chartRows.push({
      name: String(names[level - 1] || ''),
      affinity: affinity
    });
  }
  if (!chartRows.length) {
    return;
  }

  chartRows.sort(function(a, b) {
    return a.affinity - b.affinity;
  });

  var dataCol = colCount + 2;
  var values = [['Интерес', 'Аффинити-индекс']];
  for (i = 0; i < chartRows.length; i++) {
    values.push([chartRows[i].name, chartRows[i].affinity]);
  }

  var range = sheet.getRange(headerRow, dataCol, values.length, 2);
  range.setValues(values);
  range.setFontWeight('normal');
  sheet.getRange(headerRow, dataCol, 1, 2).setFontWeight('bold');
  sheet.getRange(headerRow + 1, dataCol + 1, values.length - 1, 1).setNumberFormat('0.00');
  sheet.setColumnWidth(dataCol, 220);
  sheet.setColumnWidth(dataCol + 1, 140);

  var chartHeight = Math.min(1200, Math.max(360, 48 + chartRows.length * 22));
  var chart = sheet.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(range)
    .setPosition(1, dataCol + 3, 0, 0)
    .setNumHeaders(1)
    .setOption('title', 'Аффинити-индекс по интересам (ур. ' + level + ')')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: 'Аффинити-индекс' })
    .setOption('vAxis', { title: 'Интерес' })
    .setOption('width', 640)
    .setOption('height', chartHeight)
    .build();
  sheet.insertChart(chart);
}
