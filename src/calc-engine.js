import { parseCellValue } from './config.js';
import { readSheetRange, readSingleCell, getHeaders } from './tencent-api.js';

const MODEL_CONFIG = {
  "F5631":  ["000005", 6, "J", "M1", 179],
  "F3500":  ["000005", 6, "K", "N1", 179],
  "C210":   ["000003", 4, "AC", "E1", 180],
  "C220":   ["000003", 4, "AD", "F1", 180],
  "C230":   ["000003", 4, "AE", "G1", 180],
  "C240A":  ["000003", 4, "AF", "H1", 180],
  "C3050A": ["000003", 4, "AG", "I1", 180],
  "C280":   ["000003", 4, "AH", "J1", 180],
  "330N":   ["00000a", 3, "H", "I1", 216],
  "F3600":  ["00000a", 3, "M", "O1", 216],
  "C204":   ["000006", 4, "AA", "F2", 225],
  "C307":   ["000006", 4, "AB", "G2", 225],
  "C305":   ["000006", 4, "AC", "H2", 225],
  "C310":   ["000006", 4, "AD", "I2", 225],
  "4110B":  ["000001", 4, "AB", "I2", 185],
  "5118G":  ["000001", 4, "AD", "L2", 185],
  "R4110":  ["000001", 4, "AE", "K2", 185],
  "6001C":  ["000001", 4, "AF", "M2", 185],
  "R403":   ["000001", 4, "AJ", "AK1", 185],
  "R6207":  ["000004", 3, "O", "I1", 201],
  "R6205":  ["000004", 3, "S", "J1", 201],
  "R6048":  ["000004", 3, "W", "K1", 201],
  "304铁桶": ["00000c", 3, "I", "L1", 186],
  "304吨桶": ["00000c", 3, "J", "M1", 186],
  "350T":   ["000009", 3, "N", "K1", 241],
  "8001A":  ["000009", 3, "Q", "O1", 241],
  "INOVOL R8315": ["000004", 3, "AB", "AP1", 180],
};

const FILE_ID = "DRnhDemRIS25mdnFF";
const CONFIG_FILE_ID = "DRnhDemRIS25mdnFF";
const CACHE_TTL = 300;
const LIMIT_DATE_CACHE_TTL = 300;
const MODEL_CONFIG_CACHE_TTL = 300;

const memoryCache = new Map();
const limitDateCache = new Map();
let modelConfigCache = {};
let modelConfigCacheTime = 0;

function parseDate(dateStr) {
  const s = String(dateStr || '').trim();
  if (!s) return null;

  const match1 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match1) return new Date(parseInt(match1[1]), parseInt(match1[2]) - 1, parseInt(match1[3]));

  const match2 = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
  if (match2) return new Date(parseInt(match2[1]), parseInt(match2[2]) - 1, parseInt(match2[3]));

  const match3 = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (match3) return new Date(parseInt(match3[1]), parseInt(match3[2]) - 1, parseInt(match3[3]));

  const match4 = s.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (match4) return new Date(new Date().getFullYear(), parseInt(match4[1]) - 1, parseInt(match4[2]));

  const match5 = s.match(/^(\d{1,2})\.(\d{1,2})$/);
  if (match5) return new Date(new Date().getFullYear(), parseInt(match5[1]) - 1, parseInt(match5[2]));

  return null;
}

function parseNumber(val) {
  const s = String(val || '').trim();
  if (!s) return null;
  try {
    const num = parseFloat(s);
    if (!isNaN(num)) return num;
  } catch (e) {}
  const m = s.match(/\d+(?:\.\d+)?/);
  if (m) {
    try { return parseFloat(m[0]); } catch (e) {}
  }
  return null;
}

function dateOnly(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function formatDate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}-${m}-${dd}`;
}

function addDays(d, days) {
  const r = new Date(d);
  r.setDate(r.getDate() + days);
  return r;
}

function colLetterToIndex(col) {
  let result = 0;
  for (let i = 0; i < col.length; i++) {
    result = result * 26 + (col.charCodeAt(i) - 64);
  }
  return result - 1;
}

function getFromMemory(cacheKey) {
  const entry = memoryCache.get(cacheKey);
  if (entry && (Date.now() - entry.ts) < CACHE_TTL * 1000) {
    return entry.data;
  }
  return null;
}

function setMemory(cacheKey, data) {
  memoryCache.set(cacheKey, { data, ts: Date.now() });
}

async function readLimitDate(sheetId, limitCell, kvAdminCache) {
  const cacheKey = `${sheetId}:${limitCell}`;
  const cachedEntry = limitDateCache.get(cacheKey);
  if (cachedEntry && (Date.now() - cachedEntry.ts) < LIMIT_DATE_CACHE_TTL * 1000) {
    return cachedEntry.data;
  }

  const cellValue = await readSingleCell(sheetId, limitCell, CONFIG_FILE_ID, kvAdminCache);
  const limitDate = parseDate(cellValue);

  limitDateCache.set(cacheKey, { data: limitDate, ts: Date.now() });
  return limitDate;
}

async function loadModelConfigsFromSheet(kvAdminCache) {
  const now = Date.now();
  if (Object.keys(modelConfigCache).length > 0 && (now - modelConfigCacheTime) < MODEL_CONFIG_CACHE_TTL * 1000) {
    return modelConfigCache;
  }

  const gridData = await readSheetRange('dc53jt', 'A2:F200', CONFIG_FILE_ID, kvAdminCache);
  const rows = gridData.rows || [];
  const configs = {};

  for (const row of rows) {
    const values = row.values || [];
    if (values.length < 6) continue;
    const cells = values.slice(0, 6).map(v => v.cellValue ? parseCellValue(v.cellValue) : '');

    const modelName = String(cells[0] || '').trim();
    if (!modelName) continue;

    try {
      let sheetId = String(cells[1] || '').trim();
      if (/^\d+$/.test(sheetId) && sheetId.length < 6) {
        sheetId = sheetId.padStart(6, '0');
      }
      const startRow = parseInt(cells[2]) || 4;
      const capacityCol = String(cells[3] || '').trim();
      const limitCell = String(cells[4] || '').trim();
      const rowCount = parseInt(cells[5]) || 180;
      configs[modelName] = [sheetId, startRow, capacityCol, limitCell, rowCount];
    } catch (e) {}
  }

  modelConfigCache = configs;
  modelConfigCacheTime = now;
  return configs;
}

function getModelConfig(model) {
  if (MODEL_CONFIG[model]) return MODEL_CONFIG[model];
  const sheetConfigs = modelConfigCache;
  if (sheetConfigs[model]) return sheetConfigs[model];
  return null;
}

async function getSheetData(sheetId, startRow, capacityCol, limitCell, rowCount, kvAdminCache) {
  const cacheKey = `${sheetId}:${startRow}:${capacityCol}:${limitCell}`;

  const cached = getFromMemory(cacheKey);
  if (cached && cached.dateCapacityMap) return cached;

  const endRow = startRow + rowCount - 1;

  // Read date column (A) and capacity column in parallel
  const dateRange = `A${startRow}:A${endRow}`;
  const capRange = `${capacityCol}${startRow}:${capacityCol}${endRow}`;

  const [dateGrid, capGrid] = await Promise.all([
    readSheetRange(sheetId, dateRange, FILE_ID, kvAdminCache),
    readSheetRange(sheetId, capRange, FILE_ID, kvAdminCache),
  ]);

  const dateRows = dateGrid.rows || [];
  const capRows = capGrid.rows || [];

  const dateCapacityMap = {};
  const maxRows = Math.max(dateRows.length, capRows.length);

  for (let i = 0; i < maxRows; i++) {
    let d = null;

    if (i < dateRows.length) {
      const dateValues = dateRows[i].values || [];
      if (dateValues.length > 0 && dateValues[0].cellValue) {
        const dateVal = parseCellValue(dateValues[0].cellValue);
        d = parseDate(dateVal);
      }
    }

    if (d && i < capRows.length) {
      const capValues = capRows[i].values || [];
      if (capValues.length > 0 && capValues[0].cellValue) {
        const capStr = parseCellValue(capValues[0].cellValue);
        const capVal = parseNumber(capStr);
        if (capVal !== null) {
          const dateKey = formatDate(dateOnly(d));
          dateCapacityMap[dateKey] = capVal;
        }
      }
    }
  }

  const limitDate = await readLimitDate(sheetId, limitCell, kvAdminCache);

  const result = { dateCapacityMap, limitDate };
  setMemory(cacheKey, result);
  return result;
}

async function calculateDeliveryDate(model, tonnageStr, expectedDateStr, kvAdminCache) {
  const config = getModelConfig(model);
  if (!config) {
    return { date: '请联系商务支持', error: `型号 ${model} 暂无排产数据，请检查型号是否正确` };
  }

  const tonnage = parseNumber(tonnageStr);
  if (tonnage === null || tonnage <= 0) {
    return { date: '', error: '吨位不能为空或无效' };
  }

  const expectedDate = parseDate(expectedDateStr);
  if (!expectedDate) {
    return { date: '', error: `期望发货日期格式无效: ${expectedDateStr}` };
  }

  const [sheetId, startRow, capacityCol, limitCell, rowCount] = config;

  const sheetData = await getSheetData(sheetId, startRow, capacityCol, limitCell, rowCount, kvAdminCache);
  const { dateCapacityMap, limitDate } = sheetData;

  if (!dateCapacityMap || Object.keys(dateCapacityMap).length === 0) {
    const cacheKey = `${sheetId}:${startRow}:${capacityCol}:${limitCell}`;
    memoryCache.delete(cacheKey);
    return { date: '请联系商务支持', error: '排产数据读取失败，请稍后重试或联系管理员检查Token' };
  }

  let effectiveLimitDate = limitDate;
  if (!effectiveLimitDate) {
    const allDates = Object.keys(dateCapacityMap).map(k => parseDate(k)).filter(Boolean);
    if (allDates.length > 0) {
      effectiveLimitDate = dateOnly(allDates.reduce((a, b) => a > b ? a : b));
    } else {
      return { date: '请联系商务支持', error: '上限日期未设置且无排产数据' };
    }
  }

  const expectedDateOnly = dateOnly(expectedDate);
  const limitDateOnly = dateOnly(effectiveLimitDate);

  const sortedDateKeys = Object.keys(dateCapacityMap)
    .map(k => {
      const d = parseDate(k);
      if (!d) return null;
      const dOnly = dateOnly(d);
      if (dOnly >= expectedDateOnly && dOnly <= limitDateOnly) return { key: k, d: dOnly };
      return null;
    })
    .filter(Boolean)
    .sort((a, b) => a.d.getTime() - b.d.getTime());

  if (sortedDateKeys.length === 0) {
    const maxDataDate = Object.keys(dateCapacityMap)
      .map(k => parseDate(k))
      .filter(Boolean)
      .reduce((a, b) => a > b ? a : b, null);
    const maxText = maxDataDate ? formatDate(dateOnly(maxDataDate)) : '未知';
    return { date: '请联系商务支持', error: `排产数据只到${maxText}，期望日期${expectedDateStr}超出范围` };
  }

  // Check if all dates have sufficient capacity
  const allSufficient = sortedDateKeys.every(item => {
    const cap = dateCapacityMap[item.key];
    return cap !== undefined && cap >= tonnage;
  });

  if (allSufficient) {
    return { date: expectedDateStr, error: '' };
  }

  // Find continuous intervals where capacity >= tonnage
  const intervals = [];
  let currentStart = null;

  for (let i = 0; i < sortedDateKeys.length; i++) {
    const { key, d } = sortedDateKeys[i];
    const cap = dateCapacityMap[key] || 0;

    if (cap >= tonnage) {
      if (currentStart === null) currentStart = d;
    } else {
      if (currentStart !== null) {
        intervals.push({ start: currentStart, end: i > 0 ? sortedDateKeys[i - 1].d : d });
        currentStart = null;
      }
    }
  }

  if (currentStart !== null) {
    intervals.push({ start: currentStart, end: sortedDateKeys[sortedDateKeys.length - 1].d });
  }

  if (intervals.length === 0) {
    return { date: '请联系商务支持', error: `从${expectedDateStr}到${formatDate(limitDateOnly)}均无足够产能` };
  }

  const lastInterval = intervals[intervals.length - 1];
  if (lastInterval.end < limitDateOnly) {
    return { date: '请联系商务支持', error: '请联系商务支持' };
  }

  const resultDate = lastInterval.start;
  if (resultDate.getTime() === expectedDateOnly.getTime()) {
    return { date: expectedDateStr, error: '' };
  }

  return { date: formatDate(resultDate), error: '' };
}

async function refreshCapacityData(kvAdminCache) {
  memoryCache.clear();
  limitDateCache.clear();
  modelConfigCache = {};
  modelConfigCacheTime = 0;

  try {
    await loadModelConfigsFromSheet(kvAdminCache);
  } catch (e) {}
}

// 预加载所有型号的产能数据到内存缓存，加速后续 calculate-date 调用
async function preloadAllCapacityData(kvAdminCache) {
  const configs = await loadModelConfigsFromSheet(kvAdminCache);
  const allConfigs = { ...MODEL_CONFIG, ...configs };
  const models = Object.keys(allConfigs);
  let loadedCount = 0;
  let errorCount = 0;

  for (const model of models) {
    const config = allConfigs[model];
    if (!config) continue;
    const [sheetId, startRow, capacityCol, limitCell, rowCount] = config;
    try {
      await getSheetData(sheetId, startRow, capacityCol, limitCell, rowCount, kvAdminCache);
      loadedCount++;
    } catch (e) {
      console.warn(`[preload] Failed for ${model}: ${e.message}`);
      errorCount++;
    }
  }

  return { loaded: loadedCount, errors: errorCount, total: models.length };
}

export {
  MODEL_CONFIG,
  parseDate,
  parseNumber,
  colLetterToIndex,
  formatDate,
  dateOnly,
  calculateDeliveryDate,
  refreshCapacityData,
  loadModelConfigsFromSheet,
  preloadAllCapacityData,
};
