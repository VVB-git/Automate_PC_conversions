/**
 * Разворот отчёта Метрики (время визита × целевые визиты) в строки таблицы.
 */

function formatConversionTime_(raw) {
  var value = String(raw || '').trim();
  var match = value.match(/^(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})/);
  if (match) {
    return match[1] + ' ' + match[2];
  }
  return value;
}

function dimensionTime_(dimension) {
  if (!dimension) {
    return '';
  }
  return formatConversionTime_(dimension.name || dimension.id || '');
}

function flattenGoalTable_(table, conversionType, goalMap) {
  var rows = [];
  var chunks = table && table.chunks ? table.chunks : [];
  for (var c = 0; c < chunks.length; c++) {
    var goalIds = chunks[c].goalIds || [];
    var data = chunks[c].data || [];
    for (var i = 0; i < data.length; i++) {
      var time = dimensionTime_(data[i].dimensions && data[i].dimensions[0]);
      var metrics = data[i].metrics || [];
      for (var g = 0; g < goalIds.length; g++) {
        var count = Number(metrics[g] || 0);
        if (!(count > 0)) {
          continue;
        }
        var n = Math.round(count);
        if (n < 1) {
          n = 1;
        }
        var goalName = goalMap[String(goalIds[g])] || String(goalIds[g]);
        for (var k = 0; k < n; k++) {
          rows.push({
            type: conversionType,
            dateTime: time,
            goalName: goalName
          });
        }
      }
    }
  }
  return rows;
}
