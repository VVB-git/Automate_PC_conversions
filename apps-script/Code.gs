/**
 * Выгрузки из Яндекс.Метрики и чистка post-view CPM в Google Таблице.
 *
 * Установка: Расширения → Apps Script → скопировать файлы из папки apps-script.
 * После сохранения обновите таблицу — появится меню «Автоматизация».
 */

var YM_PROP_TOKEN = 'ym_oauth_token';
var YM_PROP_FORM = 'ym_sidebar_form';
var YM_PROP_INTERESTS_FORM = 'ym_interests_form';
var YM_PROGRESS_KEY = 'ym_pc_progress';
var YM_INTERESTS_PROGRESS_KEY = 'ym_interests_progress';

function onInstall(e) {
  onOpen(e);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Автоматизация')
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Конверсии')
        .addItem('Post-click из Метрики', 'showPostClickSidebar')
        .addItem('Post-view таблица', 'runPostViewCleanup')
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Аудитория')
        .addItem('Интересы из Метрики', 'showInterestsSidebar')
    )
    .addSubMenu(
      SpreadsheetApp.getUi()
        .createMenu('Search Lift')
        .addItem('Открыть панель', 'showSearchLiftSidebar')
        .addItem('Заполнить по последним настройкам', 'fillWordstat')
        .addItem('Настройки Wordstat API', 'showWordstatSettings')
        .addItem('Защитить шаблон', 'protectTemplate')
    )
    .addSeparator()
    .addToUi();
}

function showPostClickSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Post-click конверсии')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSidebarState() {
  var form = readSavedForm_();
  var token = getSavedToken_();
  return {
    counterId: form.counterId || '',
    date1: form.date1 || '',
    date2: form.date2 || '',
    label: form.hasOwnProperty('label') ? String(form.label || '') : 'programmatic',
    labelField: form.labelField || 'utm_source',
    selectedGoalIds: form.selectedGoalIds || [],
    tokenSaved: !!token
  };
}

function getJobProgress() {
  var raw = CacheService.getDocumentCache().get(YM_PROGRESS_KEY)
    || PropertiesService.getDocumentProperties().getProperty(YM_PROGRESS_KEY);
  if (!raw) {
    return { step: '', label: '', percent: 0 };
  }
  try {
    return JSON.parse(raw);
  } catch (e) {
    return { step: '', label: '', percent: 0 };
  }
}

function loadCounterGoals(form) {
  try {
    setProgress_('goals', 'Запрашиваю список целей', 15);
    var settings = normalizeForm_(form, { allowEmptyGoals: true });
    if (!settings.goalIds.length) {
      settings.goalIds = readSavedForm_().selectedGoalIds || [];
    }
    persistSecretsAndForm_(settings);
    var goals = fetchGoals_(settings);
    setProgress_('idle', 'Цели загружены', 0);
    return {
      ok: true,
      goals: goals,
      message: goals.length
        ? ('Найдено целей: ' + goals.length)
        : 'У счётчика нет целей.'
    };
  } catch (e) {
    setProgress_('error', e.message, 0);
    throw e;
  }
}

function runPostClickJob(form) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) {
    throw new Error('Выгрузка уже идёт. Дождитесь окончания.');
  }
  try {
    setProgress_('start', 'Проверяю настройки', 5);
    var settings = normalizeForm_(form, { allowEmptyGoals: false });
    persistSecretsAndForm_(settings);

    setProgress_('goals', 'Сверяю выбранные цели', 12);
    var allGoals = fetchGoals_(settings);
    var goalMap = {};
    var i;
    for (i = 0; i < allGoals.length; i++) {
      goalMap[String(allGoals[i].id)] = allGoals[i].name;
    }
    var missing = [];
    for (i = 0; i < settings.goalIds.length; i++) {
      if (!goalMap[settings.goalIds[i]]) {
        missing.push(settings.goalIds[i]);
      }
    }
    if (missing.length) {
      throw new Error('Не найдены цели с ID: ' + missing.join(', ') + '. Загрузите список целей заново.');
    }

    setProgress_('report', 'Запрашиваю отчёт Метрики', 20);
    var result = fetchConversionRows_(settings, goalMap, function(label, percent) {
      setProgress_('report', label, percent == null ? 45 : percent);
    });
    var rows = result.rows;

    setProgress_('write', 'Пишу таблицу', 92);
    writeConversionSheet_(rows);

    setProgress_('done', 'Готово', 100);
    SpreadsheetApp.getActive().toast(
      'Строк: ' + rows.length,
      'Post-click конверсии',
      8
    );
    var message = rows.length
      ? ('Готово. Записано конверсий: ' + rows.length + '.')
      : 'Готово. По выбранным целям строк нет.';
    if (result.snapshot) {
      message += ' В Метрике за период: целевые визиты по любой цели — '
        + result.snapshot.anyGoalVisits
        + ', по выбранным целям — '
        + result.snapshot.selectedGoalVisits + '.';
      if (!rows.length && result.snapshot.anyGoalVisits > 0) {
        message += ' Отметьте в сайдбаре те цели, по которым есть конверсии (не только автоцели).';
      }
    }
    if (result.sampled) {
      message += ' Внимание: Метрика отдала сэмплированные данные, числа могут отличаться от кабинета.';
    }
    return {
      ok: true,
      count: rows.length,
      sampled: !!result.sampled,
      message: message
    };
  } catch (e) {
    setProgress_('error', e.message, 0);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function setProgress_(step, label, percent) {
  var payload = JSON.stringify({
    step: step,
    label: label || '',
    percent: percent || 0,
    t: Date.now()
  });
  CacheService.getDocumentCache().put(YM_PROGRESS_KEY, payload, 600);
  PropertiesService.getDocumentProperties().setProperty(YM_PROGRESS_KEY, payload);
}

function getSavedToken_() {
  return String(PropertiesService.getDocumentProperties().getProperty(YM_PROP_TOKEN) || '').trim();
}

function readSavedForm_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(YM_PROP_FORM);
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw) || {};
  } catch (e) {
    return {};
  }
}

function persistSecretsAndForm_(settings) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(YM_PROP_TOKEN, settings.token);
  props.setProperty(YM_PROP_FORM, JSON.stringify({
    counterId: settings.counterId,
    date1: settings.date1,
    date2: settings.date2,
    label: settings.label,
    labelField: settings.labelField,
    selectedGoalIds: settings.goalIds || []
  }));
}

function normalizeForm_(form, options) {
  form = form || {};
  options = options || {};
  var counterId = String(form.counterId || '').replace(/\s+/g, '');
  if (!/^\d+$/.test(counterId)) {
    throw new Error('Укажите номер счётчика цифрами.');
  }

  var date1 = String(form.date1 || '').trim();
  var date2 = String(form.date2 || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date1) || !/^\d{4}-\d{2}-\d{2}$/.test(date2)) {
    throw new Error('Укажите период в формате ГГГГ-ММ-ДД.');
  }
  if (date1 > date2) {
    throw new Error('Дата начала периода не может быть позже даты окончания.');
  }
  var typedToken = String(form.token || '').trim().replace(/^OAuth\s+/i, '');
  var token = typedToken || getSavedToken_();
  if (!token) {
    throw new Error('Вставьте OAuth-токен Яндекс.Метрики.');
  }

  var label = String(form.label == null ? '' : form.label).trim();

  var labelField = String(form.labelField || 'utm_source').trim();
  var allowedFields = {
    utm_source: true,
    utm_medium: true,
    utm_campaign: true,
    lastTrafficSource: true
  };
  if (!allowedFields[labelField]) {
    throw new Error('Неизвестное поле метки. Выберите utm_source, utm_medium, utm_campaign или lastTrafficSource.');
  }

  var goalIds = [];
  var rawGoals = form.goalIds || form.selectedGoalIds || [];
  if (typeof rawGoals === 'string') {
    rawGoals = rawGoals.split(/[,;\s]+/);
  }
  var seen = {};
  for (var i = 0; i < rawGoals.length; i++) {
    var id = String(rawGoals[i] || '').trim();
    if (id && /^\d+$/.test(id) && !seen[id]) {
      seen[id] = true;
      goalIds.push(id);
    }
  }
  if (!options.allowEmptyGoals && !goalIds.length) {
    throw new Error('Выберите хотя бы одну цель.');
  }

  return {
    counterId: counterId,
    date1: date1,
    date2: date2,
    token: token,
    label: label,
    labelField: labelField,
    goalIds: goalIds
  };
}

function showInterestsSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('InterestsSidebar')
    .setTitle('Интересы аудитории')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getInterestsSidebarState() {
  var form = readSavedInterestsForm_();
  var token = getSavedToken_();
  return {
    counterId: form.counterId || '',
    date1: form.date1 || '',
    date2: form.date2 || '',
    label: form.hasOwnProperty('label') ? String(form.label || '') : 'programmatic',
    labelField: form.labelField || 'utm_source',
    interestLevel: form.interestLevel || 1,
    tokenSaved: !!token
  };
}

function getInterestsJobProgress() {
  var raw = CacheService.getDocumentCache().get(YM_INTERESTS_PROGRESS_KEY)
    || PropertiesService.getDocumentProperties().getProperty(YM_INTERESTS_PROGRESS_KEY);
  if (!raw) {
    return { step: '', label: '', percent: 0 };
  }
  try {
    return JSON.parse(raw);
  } catch (e) {
    return { step: '', label: '', percent: 0 };
  }
}

function runInterestsJob(form) {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(2000)) {
    throw new Error('Выгрузка уже идёт. Дождитесь окончания.');
  }
  try {
    setInterestsProgress_('start', 'Проверяю настройки', 5);
    var settings = normalizeInterestsForm_(form);
    persistInterestsSecretsAndForm_(settings);

    setInterestsProgress_('report', 'Запрашиваю отчёт по интересам', 25);
    var report = fetchInterestsReport_(settings, function(label, percent) {
      setInterestsProgress_('report', label, percent == null ? 45 : percent);
    });

    setInterestsProgress_('write', 'Пишу таблицу', 90);
    writeInterestsSheet_(report);

    setInterestsProgress_('done', 'Готово', 100);
    SpreadsheetApp.getActive().toast(
      'Строк: ' + report.rows.length,
      'Интересы аудитории',
      8
    );
    var sampledNote = report.sampled
      ? ' Внимание: Метрика отдала сэмплированные данные (доля выборки: '
        + formatSampleShare_(report.sampleShare) + ').'
      : '';
    return {
      ok: true,
      count: report.rows.length,
      sampled: !!report.sampled,
      message: report.rows.length
        ? ('Готово. Записано интересов: ' + report.rows.length + '.' + sampledNote)
        : ('Готово. За выбранный период нет данных по интересам.' + sampledNote)
    };
  } catch (e) {
    setInterestsProgress_('error', e.message, 0);
    throw e;
  } finally {
    lock.releaseLock();
  }
}

function setInterestsProgress_(step, label, percent) {
  var payload = JSON.stringify({
    step: step,
    label: label || '',
    percent: percent || 0,
    t: Date.now()
  });
  CacheService.getDocumentCache().put(YM_INTERESTS_PROGRESS_KEY, payload, 600);
  PropertiesService.getDocumentProperties().setProperty(YM_INTERESTS_PROGRESS_KEY, payload);
}

function readSavedInterestsForm_() {
  var raw = PropertiesService.getDocumentProperties().getProperty(YM_PROP_INTERESTS_FORM);
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw) || {};
  } catch (e) {
    return {};
  }
}

function persistInterestsSecretsAndForm_(settings) {
  var props = PropertiesService.getDocumentProperties();
  props.setProperty(YM_PROP_TOKEN, settings.token);
  props.setProperty(YM_PROP_INTERESTS_FORM, JSON.stringify({
    counterId: settings.counterId,
    date1: settings.date1,
    date2: settings.date2,
    label: settings.label,
    labelField: settings.labelField,
    interestLevel: settings.interestLevel
  }));
}

function normalizeInterestsForm_(form) {
  form = form || {};
  var counterId = String(form.counterId || '').replace(/\s+/g, '');
  if (!/^\d+$/.test(counterId)) {
    throw new Error('Укажите номер счётчика цифрами.');
  }

  var date1 = String(form.date1 || '').trim();
  var date2 = String(form.date2 || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date1) || !/^\d{4}-\d{2}-\d{2}$/.test(date2)) {
    throw new Error('Укажите период в формате ГГГГ-ММ-ДД.');
  }
  if (date1 > date2) {
    throw new Error('Дата начала периода не может быть позже даты окончания.');
  }

  var typedToken = String(form.token || '').trim().replace(/^OAuth\s+/i, '');
  var token = typedToken || getSavedToken_();
  if (!token) {
    throw new Error('Вставьте OAuth-токен Яндекс.Метрики.');
  }

  var label = String(form.label == null ? '' : form.label).trim();

  var labelField = String(form.labelField || 'utm_source').trim();
  var allowedFields = {
    utm_source: true,
    utm_medium: true,
    utm_campaign: true,
    lastTrafficSource: true
  };
  if (!allowedFields[labelField]) {
    throw new Error('Неизвестное поле метки. Выберите utm_source, utm_medium, utm_campaign или lastTrafficSource.');
  }

  var interestLevel = Number(form.interestLevel);
  if (interestLevel !== 1 && interestLevel !== 2 && interestLevel !== 3) {
    throw new Error('Выберите уровень интересов: 1, 2 или 3.');
  }

  return {
    counterId: counterId,
    date1: date1,
    date2: date2,
    token: token,
    label: label,
    labelField: labelField,
    interestLevel: interestLevel
  };
}

function formatSampleShare_(share) {
  var n = Number(share);
  if (!isFinite(n) || n <= 0) {
    return 'н/д';
  }
  if (n <= 1) {
    n = n * 100;
  }
  return (Math.round(n * 100) / 100) + '%';
}
