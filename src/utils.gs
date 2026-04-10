function genId(prefix) {
  return String(prefix || 'ID') + '-' + Utilities.getUuid().slice(0, 8).toUpperCase() + '-' + Date.now().toString(36).toUpperCase();
}

function nowDateTimeStr_() {
  return formatDateTime_(new Date());
}

function nowDateStr_() {
  return formatDate_(new Date());
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, APP_INFO.timezone, 'yyyy/MM/dd HH:mm:ss');
}

function formatDate_(date) {
  return Utilities.formatDate(date, APP_INFO.timezone, 'yyyy/MM/dd');
}

function parseDateValue_(value) {
  if (!value) return null;
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;

  const normalized = String(value).trim();
  if (!normalized) return null;

  const matched = normalized.match(
    /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/
  );
  if (matched) {
    const parsed = new Date(
      Number(matched[1]),
      Number(matched[2]) - 1,
      Number(matched[3]),
      Number(matched[4] || 0),
      Number(matched[5] || 0),
      Number(matched[6] || 0)
    );
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  const fallback = new Date(normalized.replace(/\//g, '-'));
  return isNaN(fallback.getTime()) ? null : fallback;
}

function normalizeDateInput_(value) {
  const parsed = parseDateValue_(value);
  return parsed ? formatDate_(parsed) : '';
}

function addMinutes_(date, minutes) {
  return new Date(date.getTime() + Number(minutes || 0) * 60000);
}

function sheetToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0];
  return values.slice(1).map(function (row, index) {
    const record = { _row: index + 2 };
    headers.forEach(function (header, columnIndex) {
      const value = row[columnIndex];
      if (value instanceof Date) {
        const hasTime = value.getHours() !== 0
          || value.getMinutes() !== 0
          || value.getSeconds() !== 0
          || value.getMilliseconds() !== 0;
        record[header] = Utilities.formatDate(
          value,
          APP_INFO.timezone,
          hasTime ? 'yyyy/MM/dd HH:mm:ss' : 'yyyy/MM/dd'
        );
        return;
      }
      record[header] = value;
    });
    return record;
  });
}

function toRowValues_(columns, record) {
  return columns.map(function (column) {
    const value = record[column];
    return value == null ? '' : value;
  });
}

function normalizeString_(value) {
  return String(value || '').trim();
}

function normalizeEmail_(value) {
  return normalizeString_(value).toLowerCase();
}

function normalizeTextArea_(value, maxLength) {
  return normalizeString_(value).slice(0, Number(maxLength || 1000));
}

function normalizeUserAgent_(value) {
  return normalizeString_(value).slice(0, 300);
}

function normalizeNonNegativeNumber_(value, fallback) {
  if (value === '' || value == null) return fallback == null ? '' : Number(fallback || 0);
  const numeric = Number(value);
  if (!isFinite(numeric) || numeric < 0) return fallback == null ? '' : Number(fallback || 0);
  return Math.round(numeric * 100) / 100;
}

function normalizePositiveInteger_(value, fallback, min, max) {
  const normalizedFallback = Number(fallback || 1);
  const numeric = Number(value);
  const resolved = isFinite(numeric) ? Math.floor(numeric) : normalizedFallback;
  const minValue = Number(min || 1);
  const maxValue = Number(max || resolved || normalizedFallback);
  return Math.min(Math.max(resolved, minValue), maxValue);
}

function buildPaginationMeta_(totalItems, page, pageSize) {
  const total = Math.max(0, Number(totalItems) || 0);
  const safePageSize = normalizePositiveInteger_(
    pageSize,
    FIXED_VALUES.lists.defaultPageSize,
    1,
    FIXED_VALUES.lists.maxPageSize
  );
  const totalPages = Math.max(1, Math.ceil(total / safePageSize));
  const currentPage = normalizePositiveInteger_(page, 1, 1, totalPages);
  const startIndex = total === 0 ? 0 : (currentPage - 1) * safePageSize;
  const endIndex = total === 0 ? 0 : Math.min(startIndex + safePageSize, total);

  return {
    page: currentPage,
    pageSize: safePageSize,
    totalItems: total,
    totalPages: totalPages,
    hasPrev: currentPage > 1,
    hasNext: currentPage < totalPages,
    start: total === 0 ? 0 : startIndex + 1,
    end: endIndex,
  };
}

function paginateItems_(items, page, pageSize) {
  const list = items || [];
  const meta = buildPaginationMeta_(list.length, page, pageSize);
  const startIndex = meta.start > 0 ? meta.start - 1 : 0;
  return {
    items: list.slice(startIndex, startIndex + meta.pageSize),
    pagination: meta,
  };
}

function copyObject_(value) {
  return Object.assign({}, value || {});
}

function throwAppError_(code, message) {
  throw new Error(String(code || 'APP_ERROR') + ': ' + String(message || 'エラーが発生しました。'));
}

function runWithScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function getRequiredProperty_(key) {
  const value = PropertiesService.getScriptProperties().getProperty(String(key || ''));
  if (!value) {
    throwAppError_('CONFIG_MISSING', 'Script Properties に ' + key + ' を設定してください。');
  }
  return value;
}

function getDashboardCacheKey_(filter) {
  var version = getDashboardCacheVersion_();
  const mode = String((filter && filter.mode) || 'month');
  if (mode === 'all') return FIXED_VALUES.cacheKeys.dashboardAll + ':' + version;
  if (mode === 'year') return FIXED_VALUES.cacheKeys.dashboardYear + ':' + version + ':' + String(filter.year || '');
  return FIXED_VALUES.cacheKeys.dashboardMonth + ':' + version + ':' + String(filter.year || '') + ':' + String(filter.month || '');
}

function invalidateDashboardCache_() {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(SCRIPT_PROPERTY_KEYS.dashboardCacheVersion, String(Date.now()));
}

function getDashboardCacheVersion_() {
  const properties = PropertiesService.getScriptProperties();
  var version = properties.getProperty(SCRIPT_PROPERTY_KEYS.dashboardCacheVersion);
  if (!version) {
    version = String(Date.now());
    properties.setProperty(SCRIPT_PROPERTY_KEYS.dashboardCacheVersion, version);
  }
  return version;
}
