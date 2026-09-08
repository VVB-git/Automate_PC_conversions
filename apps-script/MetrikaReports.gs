/**
 * Reporting API: отчёт «Долгосрочные интересы» (preset interests2).
 */

var YM_INTEREST_DIMENSIONS = [
  'ym:s:interest2d1',
  'ym:s:interest2d2',
  'ym:s:interest2d3'
];

function interestDimensionIds_(level) {
  return YM_INTEREST_DIMENSIONS.slice(0, level);
}

function affinityToPercent_(value) {
  if (value == null || value === '') {
    return '';
  }
  var n = Number(value);
  if (!isFinite(n)) {
    return value;
  }
  if (Math.abs(n) <= 15) {
    n = n * 100;
  }
  return Math.round(n * 100) / 100;
}

function dimensionName_(dimension) {
  if (!dimension) {
    return '';
  }
  return String(dimension.name || dimension.id || '').trim();
}

function fetchInterestsReport_(settings, onProgress) {
  onProgress = onProgress || function() {};
  var level = settings.interestLevel;
  var dimensionIds = interestDimensionIds_(level);
  var extraFilters = [];
  if (settings.label) {
    extraFilters.push(visitLabelFilter_(settings, false));
  }
  var filters = joinFilters_(baseVisitFilters_(settings).concat(extraFilters));
  var params = {
    ids: settings.counterId,
    date1: settings.date1,
    date2: settings.date2,
    dimensions: dimensionIds.join(','),
    metrics: 'ym:s:users,ym:s:affinityIndexInterests2',
    accuracy: 'full',
    lang: 'ru',
    attribution: YM_ATTRIBUTION,
    sort: '-ym:s:users',
    filters: filters
  };

  var offset = 0;
  var rows = [];
  var sampled = false;
  var sampleShare = 1;
  var page = 0;

  while (true) {
    page++;
    params.limit = YM_STAT_LIMIT;
    params.offset = offset + 1;
    onProgress(
      'Запрашиваю отчёт по интересам' + (page > 1 ? (', стр. ' + page) : ''),
      Math.min(80, 20 + page * 12)
    );
    var json = metrikaStat_(settings.token, params);
    if (json.sampled) {
      sampled = true;
    }
    if (json.sample_share != null) {
      sampleShare = Number(json.sample_share);
    }
    var data = json.data || [];
    for (var i = 0; i < data.length; i++) {
      var dims = data[i].dimensions || [];
      var metrics = data[i].metrics || [];
      var names = [];
      var d;
      for (d = 0; d < level; d++) {
        names.push(dimensionName_(dims[d]));
      }
      if (!names[level - 1]) {
        continue;
      }
      rows.push({
        names: names,
        users: Number(metrics[0] || 0),
        affinity: affinityToPercent_(metrics[1])
      });
    }
    offset += data.length;
    var total = Number(json.total_rows || 0);
    if (!data.length || offset >= total) {
      break;
    }
  }

  return {
    settings: settings,
    level: level,
    rows: rows,
    sampled: sampled,
    sampleShare: sampleShare
  };
}
