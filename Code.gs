const MENU_NAME = 'Отзывы';

const SHEET_BASE = 'База';
const SHEET_LOG = 'Лог отзывов';

const HEADER_ROW = 3;

const PROP_REVIEWS_N8N_WEBHOOK_URL = 'REVIEWS_N8N_WEBHOOK_URL';
const PROP_BITRIX_WEBHOOK_BASE_URL = 'BITRIX_WEBHOOK_BASE_URL';
const PROP_REVIEWS_N8N_BATCH_SIZE = 'REVIEWS_N8N_BATCH_SIZE';

const DEFAULT_N8N_BATCH_SIZE = 10;

const BITRIX_RESPONSIBLE_USER_ID = 374;
const BITRIX_TASK_TAG = 'Ответ на положительный отзыв';

const ROCKETDATA_REVIEW_URL_PREFIX = 'https://go.rocketdata.io/reviews-management/reviews/';

const STATUS_NEW = 'new';
const STATUS_SENT_TO_N8N = 'sent_to_n8n';
const STATUS_REPLY_GENERATED = 'reply_generated';
const STATUS_TASK_CREATED = 'task_created';
const STATUS_ERROR = 'error';
const STATUS_TASK_CREATE_ERROR = 'task_create_error';

const REQUIRED_HEADERS = ['Рейтинг', 'Автор', 'Идентификатор отзыва', 'Комментарий'];
const SERVICE_HEADERS = [
  'Статус обработки',
  'Ошибка обработки',
  'Дата обработки',
  'Bitrix task ID',
  'Ссылка на задачу Bitrix',
  'Дата создания задачи Bitrix'
];
const LOG_HEADERS = ['ДатаВремя', 'Этап', 'Статус', 'Количество', 'Детали', 'Ошибка'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(MENU_NAME)
    .addItem('Отзывы: обновить базу из исходных данных', 'importReviewsToBase')
    .addItem('Отзывы: отправить новые отзывы в n8n', 'sendNewReviewsToN8n')
    .addItem('Отзывы: создать задачи Bitrix', 'createBitrixTasksForGeneratedReviewReplies')
    .addSeparator()
    .addItem('Отзывы: полный цикл', 'runReviewsFullCycle')
    .addSeparator()
    .addItem('Отзывы: повторить ошибки n8n', 'retryReviewsN8nErrors')
    .addItem('Отзывы: повторить ошибки Bitrix', 'retryReviewsBitrixErrors')
    .addToUi();
}

function importReviewsToBase() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getActiveSheet();

  if (sourceSheet.getName() === SHEET_BASE || sourceSheet.getName() === SHEET_LOG) {
    throw new Error('Активным должен быть исходный лист с отзывами, а не служебный лист.');
  }

  appendReviewLog_('import', 'start', 0, 'Старт импорта из листа: ' + sourceSheet.getName(), '');

  const sourceHeaders = getHeadersFromRow_(sourceSheet, HEADER_ROW);
  assertRequiredHeaders_(sourceHeaders);

  const baseSheet = getOrCreateBaseSheet_(sourceHeaders);
  const baseHeaders = getHeadersFromRow_(baseSheet, 1);
  const sourceHeaderMap = createHeaderMap_(sourceHeaders);
  const baseHeaderMap = createHeaderMap_(baseHeaders);
  const existingIds = getExistingReviewIds_(baseSheet, baseHeaderMap);

  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceHeaders.length;
  const dataRowCount = Math.max(0, lastRow - HEADER_ROW);
  appendReviewLog_('import', 'found', dataRowCount, 'Найдено строк данных на исходном листе', '');

  if (dataRowCount === 0) {
    appendReviewLog_('import', 'success', 0, 'Нет строк для импорта', '');
    return { imported: 0, found: 0 };
  }

  const sourceValues = sourceSheet.getRange(HEADER_ROW + 1, 1, dataRowCount, lastColumn).getValues();
  const rowsToAppend = [];
  const now = new Date();
  const stats = {
    imported: 0,
    duplicates: 0,
    lowRating: 0,
    emptyComment: 0,
    emptyReviewId: 0,
    emptyRows: 0
  };

  sourceValues.forEach(function(row) {
    if (isEmptyRow_(row)) {
      stats.emptyRows += 1;
      return;
    }

    const reviewId = normalizeText_(getValueByHeader_(row, sourceHeaderMap, 'Идентификатор отзыва'));
    const comment = normalizeText_(getValueByHeader_(row, sourceHeaderMap, 'Комментарий'));
    const rating = parseRating_(getValueByHeader_(row, sourceHeaderMap, 'Рейтинг'));

    if (!reviewId) {
      stats.emptyReviewId += 1;
      return;
    }
    if (!comment) {
      stats.emptyComment += 1;
      return;
    }
    if (rating < 4) {
      stats.lowRating += 1;
      return;
    }
    if (existingIds[reviewId]) {
      stats.duplicates += 1;
      return;
    }

    const baseRow = buildBaseRowFromSource_(row, sourceHeaders, baseHeaders, baseHeaderMap);
    baseRow[baseHeaderMap['Статус обработки'][0]] = STATUS_NEW;
    baseRow[baseHeaderMap['Ошибка обработки'][0]] = '';
    baseRow[baseHeaderMap['Дата обработки'][0]] = now;
    rowsToAppend.push(baseRow);
    existingIds[reviewId] = true;
    stats.imported += 1;
  });

  if (rowsToAppend.length > 0) {
    baseSheet.getRange(baseSheet.getLastRow() + 1, 1, rowsToAppend.length, baseHeaders.length).setValues(rowsToAppend);
  }

  appendReviewLog_('import', 'imported', stats.imported, 'Перенесено новых положительных отзывов', '');
  appendReviewLog_('import', 'skipped_duplicates', stats.duplicates, 'Пропущено дублей', '');
  appendReviewLog_('import', 'skipped_low_rating', stats.lowRating, 'Пропущено из-за рейтинга ниже 4', '');
  appendReviewLog_('import', 'skipped_empty_comment', stats.emptyComment, 'Пропущено из-за пустого комментария', '');
  appendReviewLog_('import', 'skipped_empty_review_id', stats.emptyReviewId, 'Пропущено из-за пустого идентификатора отзыва', '');

  return stats;
}

function sendNewReviewsToN8n() {
  return sendReviewsToN8nByStatuses_(['', STATUS_NEW, STATUS_ERROR]);
}

function retryReviewsN8nErrors() {
  return sendReviewsToN8nByStatuses_([STATUS_ERROR]);
}

function createBitrixTasksForGeneratedReviewReplies() {
  return createBitrixTasksByStatuses_([STATUS_REPLY_GENERATED]);
}

function retryReviewsBitrixErrors() {
  return createBitrixTasksByStatuses_([STATUS_TASK_CREATE_ERROR]);
}

function runReviewsFullCycle() {
  importReviewsToBase();
  sendNewReviewsToN8n();
  createBitrixTasksForGeneratedReviewReplies();
}

function sendReviewsToN8nByStatuses_(allowedStatuses) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const baseSheet = getExistingBaseSheet_();
  const headers = getHeadersFromRow_(baseSheet, 1);
  assertRequiredHeaders_(headers);
  ensureBaseServiceColumns_(baseSheet, headers);

  const refreshedHeaders = getHeadersFromRow_(baseSheet, 1);
  const headerMap = createHeaderMap_(refreshedHeaders);
  const webhookUrl = getRequiredScriptProperty_(PROP_REVIEWS_N8N_WEBHOOK_URL);
  const batchSize = getN8nBatchSize_();
  const rows = readSheetObjects_(baseSheet, refreshedHeaders, headerMap);
  const candidates = rows.filter(function(row) {
    const status = normalizeText_(getCellValue_(row.values, headerMap, 'Статус обработки'));
    const reviewId = normalizeText_(getCellValue_(row.values, headerMap, 'Идентификатор отзыва'));
    const comment = normalizeText_(getCellValue_(row.values, headerMap, 'Комментарий'));
    const rating = parseRating_(getCellValue_(row.values, headerMap, 'Рейтинг'));
    return allowedStatuses.indexOf(status) !== -1 && reviewId && comment && rating >= 4;
  });

  if (candidates.length === 0) {
    appendReviewLog_('n8n', 'success', 0, 'Нет отзывов для отправки в n8n', '');
    return { sent: 0, results: 0, errors: 0 };
  }

  let sent = 0;
  let resultCount = 0;
  let errorCount = 0;
  for (let offset = 0; offset < candidates.length; offset += batchSize) {
    const batch = candidates.slice(offset, offset + batchSize);
    const batchResult = sendReviewsBatchToN8n_(baseSheet, headerMap, webhookUrl, batch);
    sent += batchResult.sent;
    resultCount += batchResult.results;
    errorCount += batchResult.errors;
  }

  return { sent: sent, results: resultCount, errors: errorCount };
}

function sendReviewsBatchToN8n_(baseSheet, headerMap, webhookUrl, batch) {
  appendReviewLog_('n8n', 'send_batch', batch.length, 'Отправка пачки отзывов в n8n', '');
  updateRows_(baseSheet, headerMap, batch, {
    'Статус обработки': STATUS_SENT_TO_N8N,
    'Дата обработки': new Date(),
    'Ошибка обработки': ''
  });

  const payload = {
    source: 'google_apps_script',
    action: 'review_reply_generate_batch',
    items: batch.map(function(row) {
      return {
        review_id: normalizeText_(getCellValue_(row.values, headerMap, 'Идентификатор отзыва')),
        author: normalizeText_(getCellValue_(row.values, headerMap, 'Автор')),
        comment: normalizeText_(getCellValue_(row.values, headerMap, 'Комментарий'))
      };
    })
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    const body = response.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error('n8n вернул HTTP ' + code + ': ' + truncate_(body, 1000));
    }

    const parsed = parseJsonResponse_(body, 'n8n');
    const results = normalizeN8nResults_(parsed);
    applyN8nResults_(baseSheet, headerMap, batch, results);
    appendReviewLog_('n8n', 'success', results.length, 'Успешно получены ответы от n8n', '');
    return { sent: batch.length, results: results.length, errors: 0 };
  } catch (error) {
    const errorMessage = errorToString_(error);
    updateRows_(baseSheet, headerMap, batch, {
      'Статус обработки': STATUS_ERROR,
      'Ошибка обработки': errorMessage,
      'Дата обработки': new Date()
    });
    appendReviewLog_('n8n', 'error', batch.length, 'Ошибка отправки/получения пачки n8n', errorMessage);
    return { sent: batch.length, results: 0, errors: batch.length };
  }
}

function createBitrixTasksByStatuses_(allowedStatuses) {
  const baseSheet = getExistingBaseSheet_();
  const headers = getHeadersFromRow_(baseSheet, 1);
  assertRequiredHeaders_(headers);
  ensureBaseServiceColumns_(baseSheet, headers);

  const refreshedHeaders = getHeadersFromRow_(baseSheet, 1);
  const headerMap = createHeaderMap_(refreshedHeaders);
  const webhookBaseUrl = normalizeBitrixWebhookBaseUrl_(getRequiredScriptProperty_(PROP_BITRIX_WEBHOOK_BASE_URL));
  const rows = readSheetObjects_(baseSheet, refreshedHeaders, headerMap);
  const candidates = rows.filter(function(row) {
    const status = normalizeText_(getCellValue_(row.values, headerMap, 'Статус обработки'));
    const replyText = normalizeText_(getCellValue_(row.values, headerMap, 'Ответ на отзыв'));
    const taskId = normalizeText_(getCellValue_(row.values, headerMap, 'Bitrix task ID'));
    return allowedStatuses.indexOf(status) !== -1 && replyText && !taskId;
  });

  if (candidates.length === 0) {
    appendReviewLog_('bitrix', 'success', 0, 'Нет задач Bitrix для создания', '');
    return { created: 0 };
  }

  appendReviewLog_('bitrix', 'start', candidates.length, 'Создание задач Bitrix', '');
  let created = 0;
  candidates.forEach(function(row) {
    const reviewId = normalizeText_(getCellValue_(row.values, headerMap, 'Идентификатор отзыва'));
    const title = buildBitrixTaskTitle_(row.values, headerMap);
    const description = buildBitrixTaskDescription_(row.values, headerMap);
    const payload = {
      fields: {
        TITLE: title,
        DESCRIPTION: description,
        RESPONSIBLE_ID: BITRIX_RESPONSIBLE_USER_ID,
        TAGS: [BITRIX_TASK_TAG]
      }
    };

    try {
      const response = UrlFetchApp.fetch(webhookBaseUrl + 'tasks.task.add.json', {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = response.getResponseCode();
      const body = response.getContentText();
      if (code < 200 || code >= 300) {
        throw new Error('Bitrix24 вернул HTTP ' + code + ': ' + truncate_(body, 1000));
      }

      const parsed = parseJsonResponse_(body, 'Bitrix24');
      const taskId = extractBitrixTaskId_(parsed);
      if (!taskId) {
        throw new Error('Bitrix24 не вернул ID задачи: ' + truncate_(body, 1000));
      }

      const taskUrl = buildBitrixTaskUrl_(webhookBaseUrl, taskId);
      updateSingleRow_(baseSheet, headerMap, row.rowNumber, {
        'Bitrix task ID': taskId,
        'Ссылка на задачу Bitrix': taskUrl,
        'Дата создания задачи Bitrix': new Date(),
        'Статус обработки': STATUS_TASK_CREATED,
        'Ошибка обработки': ''
      });
      created += 1;
      appendReviewLog_('bitrix', 'success', 1, 'Создана задача Bitrix для отзыва ' + reviewId + ', task ID: ' + taskId, '');
    } catch (error) {
      const errorMessage = errorToString_(error);
      updateSingleRow_(baseSheet, headerMap, row.rowNumber, {
        'Статус обработки': STATUS_TASK_CREATE_ERROR,
        'Ошибка обработки': errorMessage
      });
      appendReviewLog_('bitrix', 'error', 1, 'Ошибка создания задачи Bitrix для отзыва ' + reviewId, errorMessage);
    }
  });

  appendReviewLog_('bitrix', 'finished', created, 'Создание задач Bitrix завершено', '');
  return { created: created, total: candidates.length };
}

function getOrCreateBaseSheet_(sourceHeaders) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let baseSheet = spreadsheet.getSheetByName(SHEET_BASE);
  if (!baseSheet) {
    baseSheet = spreadsheet.insertSheet(SHEET_BASE);
    const headers = sourceHeaders.slice();
    SERVICE_HEADERS.forEach(function(header) {
      if (headers.indexOf(header) === -1) {
        headers.push(header);
      }
    });
    baseSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return baseSheet;
  }

  let baseHeaders = getHeadersFromRow_(baseSheet, 1);
  if (baseHeaders.length === 0 || baseHeaders.every(function(header) { return !header; })) {
    baseHeaders = sourceHeaders.slice();
    SERVICE_HEADERS.forEach(function(header) {
      if (baseHeaders.indexOf(header) === -1) {
        baseHeaders.push(header);
      }
    });
    baseSheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
    return baseSheet;
  }

  ensureBaseServiceColumns_(baseSheet, baseHeaders);
  return baseSheet;
}

function getExistingBaseSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BASE);
  if (!sheet) {
    throw new Error('Лист «' + SHEET_BASE + '» не найден. Сначала выполните импорт отзывов в базу.');
  }
  return sheet;
}

function ensureBaseServiceColumns_(sheet, headers) {
  const existingHeaders = headers || getHeadersFromRow_(sheet, 1);
  const toAdd = SERVICE_HEADERS.filter(function(header) {
    return existingHeaders.indexOf(header) === -1;
  });
  if (existingHeaders.indexOf('Ответ на отзыв') === -1) {
    toAdd.unshift('Ответ на отзыв');
  }
  if (toAdd.length > 0) {
    sheet.getRange(1, existingHeaders.length + 1, 1, toAdd.length).setValues([toAdd]);
  }
}

function getHeadersFromRow_(sheet, rowNumber) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return [];
  }
  return sheet.getRange(rowNumber, 1, 1, lastColumn).getValues()[0].map(function(header) {
    return normalizeText_(header);
  });
}

function createHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    if (!header) {
      return;
    }
    if (!map[header]) {
      map[header] = [];
    }
    map[header].push(index);
  });
  return map;
}

function assertRequiredHeaders_(headers) {
  const map = createHeaderMap_(headers);
  const missing = REQUIRED_HEADERS.filter(function(header) {
    return !map[header];
  });
  if (missing.length > 0) {
    throw new Error('Отсутствуют обязательные колонки: ' + missing.join(', '));
  }
}

function getExistingReviewIds_(sheet, headerMap) {
  const ids = {};
  if (!headerMap['Идентификатор отзыва'] || sheet.getLastRow() < 2) {
    return ids;
  }
  const reviewIdColumn = headerMap['Идентификатор отзыва'][0] + 1;
  const values = sheet.getRange(2, reviewIdColumn, sheet.getLastRow() - 1, 1).getValues();
  values.forEach(function(row) {
    const id = normalizeText_(row[0]);
    if (id) {
      ids[id] = true;
    }
  });
  return ids;
}

function buildBaseRowFromSource_(sourceRow, sourceHeaders, baseHeaders, baseHeaderMap) {
  const baseRow = new Array(baseHeaders.length).fill('');
  sourceHeaders.forEach(function(header, sourceIndex) {
    const targetIndexes = baseHeaderMap[header];
    if (!header || !targetIndexes) {
      return;
    }
    const duplicatePosition = countPreviousHeaderOccurrences_(sourceHeaders, header, sourceIndex);
    const targetIndex = targetIndexes[Math.min(duplicatePosition, targetIndexes.length - 1)];
    baseRow[targetIndex] = sourceRow[sourceIndex];
  });
  return baseRow;
}

function countPreviousHeaderOccurrences_(headers, header, currentIndex) {
  let count = 0;
  for (let i = 0; i < currentIndex; i += 1) {
    if (headers[i] === header) {
      count += 1;
    }
  }
  return count;
}

function readSheetObjects_(sheet, headers, headerMap) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .map(function(values, index) {
      return { rowNumber: index + 2, values: values };
    })
    .filter(function(row) {
      return !isEmptyRow_(row.values);
    });
}

function applyN8nResults_(sheet, headerMap, sentRows, results) {
  const sentById = {};
  sentRows.forEach(function(row) {
    const reviewId = normalizeText_(getCellValue_(row.values, headerMap, 'Идентификатор отзыва'));
    sentById[reviewId] = row;
  });

  const returnedIds = {};
  results.forEach(function(result) {
    const reviewId = normalizeText_(result.review_id || result.id || result.reviewId);
    const row = sentById[reviewId];
    if (!row) {
      return;
    }
    returnedIds[reviewId] = true;
    if (result.ok === false || result.error) {
      updateSingleRow_(sheet, headerMap, row.rowNumber, {
        'Статус обработки': STATUS_ERROR,
        'Ошибка обработки': normalizeText_(result.error) || 'n8n вернул ошибку по отзыву',
        'Дата обработки': new Date()
      });
      return;
    }

    const replyText = normalizeText_(result.reply_text || result.replyText || result.reply);
    if (!replyText) {
      updateSingleRow_(sheet, headerMap, row.rowNumber, {
        'Статус обработки': STATUS_ERROR,
        'Ошибка обработки': 'n8n не вернул reply_text по отзыву',
        'Дата обработки': new Date()
      });
      return;
    }

    updateSingleRow_(sheet, headerMap, row.rowNumber, {
      'Ответ на отзыв': replyText,
      'Статус обработки': STATUS_REPLY_GENERATED,
      'Ошибка обработки': '',
      'Дата обработки': new Date()
    });
  });

  sentRows.forEach(function(row) {
    const reviewId = normalizeText_(getCellValue_(row.values, headerMap, 'Идентификатор отзыва'));
    if (!returnedIds[reviewId]) {
      updateSingleRow_(sheet, headerMap, row.rowNumber, {
        'Статус обработки': STATUS_ERROR,
        'Ошибка обработки': 'n8n не вернул результат для отзыва ' + reviewId,
        'Дата обработки': new Date()
      });
    }
  });
}

function normalizeN8nResults_(parsed) {
  if (Array.isArray(parsed)) {
    return parsed;
  }
  if (parsed && Array.isArray(parsed.results)) {
    return parsed.results;
  }
  throw new Error('n8n вернул неожиданный формат ответа');
}

function buildBitrixTaskTitle_(row, headerMap) {
  const catalog = normalizeText_(getCellValue_(row, headerMap, 'Каталог'));
  const city = normalizeText_(getCellValue_(row, headerMap, 'Город'));
  const rating = normalizeText_(getCellValue_(row, headerMap, 'Рейтинг'));
  const details = [];
  if (catalog) {
    details.push(catalog);
  }
  if (city) {
    details.push(city);
  }
  if (rating) {
    details.push('рейтинг ' + rating);
  }
  return details.length > 0
    ? 'Ответить на положительный отзыв — ' + details.join(', ')
    : 'Ответить на положительный отзыв';
}

function buildBitrixTaskDescription_(row, headerMap) {
  const reviewId = normalizeText_(getCellValue_(row, headerMap, 'Идентификатор отзыва'));
  const company = normalizeText_(getCellValue_(row, headerMap, 'Компания')) || normalizeText_(getCellValue_(row, headerMap, 'Код филиала'));
  const lines = [
    '[B]Нужно ответить на положительный отзыв[/B]',
    '',
    '[B]Рейтинг:[/B] ' + escapeBitrixText_(getCellValue_(row, headerMap, 'Рейтинг')),
    '[B]Автор:[/B] ' + escapeBitrixText_(getCellValue_(row, headerMap, 'Автор')),
    '[B]Город:[/B] ' + escapeBitrixText_(getCellValue_(row, headerMap, 'Город')),
    '[B]Компания:[/B] ' + escapeBitrixText_(company),
    '[B]Каталог:[/B] ' + escapeBitrixText_(getCellValue_(row, headerMap, 'Каталог')),
    '',
    '[B]Текст отзыва:[/B]',
    escapeBitrixText_(getCellValue_(row, headerMap, 'Комментарий')),
    '',
    '[B]Подготовленный ответ:[/B]',
    escapeBitrixText_(getCellValue_(row, headerMap, 'Ответ на отзыв')),
    '',
    '[B]Ссылка на отзыв:[/B]',
    ROCKETDATA_REVIEW_URL_PREFIX + encodeURIComponent(reviewId)
  ];
  return lines.join('\n');
}

function escapeBitrixText_(value) {
  return normalizeText_(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function extractBitrixTaskId_(response) {
  if (!response) {
    return '';
  }
  if (response.result && response.result.task && response.result.task.id) {
    return String(response.result.task.id);
  }
  if (response.result && response.result.id) {
    return String(response.result.id);
  }
  if (response.task && response.task.id) {
    return String(response.task.id);
  }
  return '';
}

function buildBitrixTaskUrl_(webhookBaseUrl, taskId) {
  const match = webhookBaseUrl.match(/^(https?:\/\/[^/]+)\/rest\/([^/]+)\//);
  if (!match) {
    return '';
  }
  return match[1] + '/company/personal/user/' + match[2] + '/tasks/task/view/' + taskId + '/';
}

function updateRows_(sheet, headerMap, rows, valuesByHeader) {
  rows.forEach(function(row) {
    updateSingleRow_(sheet, headerMap, row.rowNumber, valuesByHeader);
  });
}

function updateSingleRow_(sheet, headerMap, rowNumber, valuesByHeader) {
  Object.keys(valuesByHeader).forEach(function(header) {
    if (!headerMap[header]) {
      throw new Error('Колонка «' + header + '» не найдена на листе «' + sheet.getName() + '».');
    }
    sheet.getRange(rowNumber, headerMap[header][0] + 1).setValue(valuesByHeader[header]);
  });
}

function getCellValue_(row, headerMap, header) {
  if (!headerMap[header]) {
    return '';
  }
  return row[headerMap[header][0]];
}

function getValueByHeader_(row, headerMap, header) {
  return getCellValue_(row, headerMap, header);
}

function parseRating_(value) {
  if (typeof value === 'number') {
    return value;
  }
  const text = normalizeText_(value).replace(',', '.');
  const match = text.match(/-?\d+(\.\d+)?/);
  if (!match) {
    return 0;
  }
  const parsed = Number(match[0]);
  return isNaN(parsed) ? 0 : parsed;
}

function isEmptyRow_(row) {
  return row.every(function(value) {
    return normalizeText_(value) === '';
  });
}

function normalizeText_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value).trim();
}

function getN8nBatchSize_() {
  const raw = PropertiesService.getScriptProperties().getProperty(PROP_REVIEWS_N8N_BATCH_SIZE);
  const parsed = Number(raw);
  if (!raw || !isFinite(parsed) || parsed <= 0) {
    return DEFAULT_N8N_BATCH_SIZE;
  }
  return Math.floor(parsed);
}

function getRequiredScriptProperty_(name) {
  const value = normalizeText_(PropertiesService.getScriptProperties().getProperty(name));
  if (!value) {
    throw new Error('Не задано Script Property: ' + name);
  }
  return value;
}

function normalizeBitrixWebhookBaseUrl_(url) {
  return url.charAt(url.length - 1) === '/' ? url : url + '/';
}

function parseJsonResponse_(body, sourceName) {
  try {
    return JSON.parse(body);
  } catch (error) {
    throw new Error(sourceName + ' вернул невалидный JSON: ' + truncate_(body, 1000));
  }
}

function appendReviewLog_(stage, status, count, details, error) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_LOG);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  }
  sheet.appendRow([new Date(), stage, status, count, details || '', error || '']);
}

function truncate_(value, maxLength) {
  const text = normalizeText_(value);
  return text.length > maxLength ? text.slice(0, maxLength) + '…' : text;
}

function errorToString_(error) {
  return error && error.message ? error.message : String(error);
}
