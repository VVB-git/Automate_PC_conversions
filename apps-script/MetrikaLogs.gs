/**
 * Клиент Management API и Reporting API Яндекс.Метрики (без Logs API).
 * Прямые и ассоциированные — те же сегменты, что в конструкторе отчётов.
 */

var YM_API = 'https://api-metrika.yandex.net';
var YM_MAX_RETRIES = 4;
var YM_STAT_LIMIT = 10000;
var YM_GOAL_CHUNK = 10;
var YM_ATTRIBUTION = 'cross_device_last_significant';

function parseMetrikaError_(code, body) {
  var message = '';
  try {
    var json = JSON.parse(body);
    if (json.message) {
      message = json.message;
    }
    if (json.errors && json.errors.length) {
      message = json.errors.map(function(err) {
        var part = err.message || err.error_type || '';
        if (err.location) {
          part += ' [' + err.location + ']';
        }
        return part;
      }).filter(Boolean).join('; ');
    }
  } catch (e) {
    if (body) {
      message = String(body).substring(0, 240);
    }
  }
  if (code === 401) {
    return 'Токен неверный или истёк. Вставьте новый OAuth-токен в боковое меню.';
  }
  if (code === 403) {
    return 'Нет доступа к этому счётчику. Проверьте номер счётчика и права токена (нужно metrika:read).';
  }
  if (code === 404) {
    return 'Счётчик не найден. Проверьте номер.';
  }
  if (code === 429) {
    return 'Метрика временно ограничила запросы. Подождите минуту и нажмите ещё раз.';
  }
  return message ? ('Метрика: ' + message) : ('Ошибка Метрики, код ' + code);
}

function encodeQuery_(params) {
  var keys = Object.keys(params);
  var chunks = [];
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    var value = params[key];
    if (value === '' || value === null || typeof value === 'undefined') {
      continue;
    }
    chunks.push(encodeURIComponent(key) + '=' + encodeURIComponent(value));
  }
  return chunks.join('&');
}

function metrikaFetch_(pathAndQuery, token) {
  var url = YM_API + pathAndQuery;
  var lastError = 'Не удалось связаться с Метрикой.';
  for (var attempt = 1; attempt <= YM_MAX_RETRIES; attempt++) {
    var response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'OAuth ' + token,
        Accept: 'application/json'
      },
      muteHttpExceptions: true,
      followRedirects: true
    });
    var code = response.getResponseCode();
    var body = response.getContentText();
    if (code >= 200 && code < 300) {
      try {
        return JSON.parse(body);
      } catch (e) {
        throw new Error('Метрика вернула ответ, который не получилось прочитать.');
      }
    }
    lastError = parseMetrikaError_(code, body);
    var retryable = code === 429 || code >= 500;
    if (!retryable || attempt === YM_MAX_RETRIES) {
      var error = new Error(lastError);
      error.httpCode = code;
      throw error;
    }
    Utilities.sleep(1500 * attempt);
  }
  throw new Error(lastError);
}

function metrikaStat_(token, params) {
  var json = metrikaFetch_('/stat/v1/data?' + encodeQuery_(params), token);
  Utilities.sleep(250);
  return json;
}

function fetchGoals_(settings) {
  var json = metrikaFetch_(
    '/management/v1/counter/' + settings.counterId + '/goals',
    settings.token
  );
  var goals = json.goals || [];
  return goals.map(function(g) {
    return { id: String(g.id), name: g.name || '' };
  }).filter(function(g) {
    return g.id && g.name;
  });
}

function escapeFilterValue_(value) {
  return String(value).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function apiFieldForLabel_(labelField) {
  if (labelField === 'utm_medium') {
    return 'UTMMedium';
  }
  if (labelField === 'utm_campaign') {
    return 'UTMCampaign';
  }
  if (labelField === 'lastTrafficSource') {
    return 'lastTrafficSource';
  }
  return 'UTMSource';
}

function attributedField_(fieldName) {
  return 'ym:s:' + YM_ATTRIBUTION.replace(/-/g, '_') + fieldName;
}

function visitLabelFilter_(settings, negate) {
  var field = attributedField_(apiFieldForLabel_(settings.labelField));
  var op = negate ? '!@' : '=@';
  return field + op + "'" + escapeFilterValue_(settings.label) + "'";
}

function userHadLabelFilter_(settings) {
  var field = 'ym:s:' + apiFieldForLabel_(settings.labelField);
  return 'USER(' + field + "=@'" + escapeFilterValue_(settings.label) + "')";
}

function baseVisitFilters_(settings) {
  var parts = ["ym:s:isRobot=='No'"];
  if (settings.label) {
    parts.push("ym:s:automaticUTMMedium!='test'");
  }
  return parts;
}

function joinFilters_(parts) {
  return parts.filter(Boolean).join(' AND ');
}

function chunkArray_(items, size) {
  var chunks = [];
  for (var i = 0; i < items.length; i += size) {
    chunks.push(items.slice(i, i + size));
  }
  return chunks;
}

function fetchStatPages_(settings, params, onProgress, stepLabel) {
  var offset = 0;
  var rows = [];
  var sampled = false;
  var page = 0;
  while (true) {
    page++;
    params.limit = YM_STAT_LIMIT;
    params.offset = offset + 1;
    if (onProgress) {
      onProgress(stepLabel + (page > 1 ? (', стр. ' + page) : ''), null);
    }
    var json = metrikaStat_(settings.token, params);
    if (json.sampled) {
      sampled = true;
    }
    var data = json.data || [];
    for (var i = 0; i < data.length; i++) {
      rows.push(data[i]);
    }
    offset += data.length;
    var total = Number(json.total_rows || 0);
    if (!data.length || offset >= total) {
      break;
    }
  }
  return { data: rows, sampled: sampled };
}

function fetchGoalTable_(settings, extraFilters, goalIds, onProgress, stepLabel) {
  var chunks = chunkArray_(goalIds, YM_GOAL_CHUNK);
  var sampled = false;
  var combined = [];
  for (var c = 0; c < chunks.length; c++) {
    var chunk = chunks[c];
    var metrics = chunk.map(function(id) {
      return 'ym:s:goal' + id + 'visits';
    });
    var filters = joinFilters_(
      baseVisitFilters_(settings).concat(extraFilters || [])
    );
    var params = {
      ids: settings.counterId,
      date1: settings.date1,
      date2: settings.date2,
      dimensions: 'ym:s:dateTime',
      metrics: metrics.join(','),
      accuracy: 'full',
      lang: 'ru',
      attribution: settings.label ? YM_ATTRIBUTION : 'lastsign'
    };
    if (filters) {
      params.filters = filters;
    }
    var label = stepLabel;
    if (chunks.length > 1) {
      label += ' (' + (c + 1) + '/' + chunks.length + ')';
    }
    var page = fetchStatPages_(settings, params, onProgress, label);
    if (page.sampled) {
      sampled = true;
    }
    combined.push({ goalIds: chunk, data: page.data });
  }
  return { chunks: combined, sampled: sampled };
}

function fetchTotalsSnapshot_(settings, onProgress) {
  var metrics = ['ym:s:anyGoalVisits'];
  var i;
  for (i = 0; i < settings.goalIds.length && metrics.length < 12; i++) {
    metrics.push('ym:s:goal' + settings.goalIds[i] + 'visits');
  }
  var params = {
    ids: settings.counterId,
    date1: settings.date1,
    date2: settings.date2,
    metrics: metrics.join(','),
    accuracy: 'full',
    lang: 'ru',
    attribution: settings.label ? YM_ATTRIBUTION : 'lastsign'
  };
  var filters = joinFilters_(baseVisitFilters_(settings));
  if (filters) {
    params.filters = filters;
  }
  if (onProgress) {
    onProgress('Сверяю итоги Метрики', 30);
  }
  var json = metrikaStat_(settings.token, params);
  var totals = json.totals || [];
  var selectedSum = 0;
  for (i = 1; i < totals.length; i++) {
    selectedSum += Number(totals[i] || 0);
  }
  return {
    anyGoalVisits: Number(totals[0] || 0),
    selectedGoalVisits: selectedSum,
    sampled: !!json.sampled
  };
}

function fetchAnyGoalTable_(settings, extraFilters, onProgress, stepLabel) {
  var filters = joinFilters_(
    baseVisitFilters_(settings).concat(extraFilters || [])
  );
  var params = {
    ids: settings.counterId,
    date1: settings.date1,
    date2: settings.date2,
    dimensions: 'ym:s:dateTime',
    metrics: 'ym:s:anyGoalVisits',
    accuracy: 'full',
    lang: 'ru',
    attribution: settings.label ? YM_ATTRIBUTION : 'lastsign'
  };
  if (filters) {
    params.filters = filters;
  }
  var page = fetchStatPages_(settings, params, onProgress, stepLabel);
  return {
    chunks: [{ goalIds: ['any'], data: page.data }],
    sampled: page.sampled
  };
}

function fetchConversionRows_(settings, goalMap, onProgress) {
  onProgress = onProgress || function() {};
  var sampled = false;
  var rows = [];
  var snapshot = fetchTotalsSnapshot_(settings, onProgress);
  sampled = !!snapshot.sampled;

  if (!settings.label) {
    onProgress('Считаю конверсии по всем источникам', 40);
    var allTable = fetchGoalTable_(settings, [], settings.goalIds, onProgress, 'Все источники');
    sampled = sampled || allTable.sampled;
    rows = rows.concat(flattenGoalTable_(allTable, 'Прямая', goalMap));
    if (!rows.length && snapshot.anyGoalVisits > 0) {
      onProgress('Беру целевые визиты по любой цели', 70);
      var anyTable = fetchAnyGoalTable_(settings, [], onProgress, 'Любая цель');
      sampled = sampled || anyTable.sampled;
      rows = rows.concat(flattenGoalTable_(anyTable, 'Прямая', {
        any: 'Целевые визиты по любой цели'
      }));
    }
  } else {
    onProgress('Считаю прямые конверсии', 35);
    var directTable = fetchGoalTable_(
      settings,
      [visitLabelFilter_(settings, false)],
      settings.goalIds,
      onProgress,
      'Прямые'
    );
    sampled = sampled || directTable.sampled;
    rows = rows.concat(flattenGoalTable_(directTable, 'Прямая', goalMap));

    onProgress('Считаю ассоциированные конверсии', 65);
    var assocTable = fetchGoalTable_(
      settings,
      [userHadLabelFilter_(settings), visitLabelFilter_(settings, true)],
      settings.goalIds,
      onProgress,
      'Ассоциированные'
    );
    sampled = sampled || assocTable.sampled;
    rows = rows.concat(flattenGoalTable_(assocTable, 'Ассоциированная', goalMap));
  }

  rows.sort(function(a, b) {
    if (a.dateTime < b.dateTime) {
      return -1;
    }
    if (a.dateTime > b.dateTime) {
      return 1;
    }
    if (a.goalName < b.goalName) {
      return -1;
    }
    if (a.goalName > b.goalName) {
      return 1;
    }
    return 0;
  });

    return { rows: rows, sampled: sampled, snapshot: snapshot };
}
