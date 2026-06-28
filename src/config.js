// Tencent Docs API config
const FILE_ID = "DRnhDemRIS25mdnFF";
const SHEET_ID = "000007";
const MODEL_FILE_ID = "DRmxUY0RBQVJXRXpC";
const MODEL_SHEET_ID = "fkayvi";
const USER_FILE_ID = "DRmxUY0RBQVJXRXpC";
const USER_SHEET_ID = "s9osf8";
const CONFIG_FILE_ID = "DRnhDemRIS25mdnFF";

// API base URL
const BASE_URL = "https://docs.qq.com/openapi/spreadsheet/v3";

// Default credentials - will be overridden by env vars
let CLIENT_ID = 'da815d1227294457b43413bdc16e3e90';
let OPEN_ID = '9bc172e5338147d8a35c1438ea8d1577';
let ACCESS_TOKEN = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbHQiOiJkYTgxNWQxMjI3Mjk0NDU3YjQzNDEzYmRjMTZlM2U5MCIsInR5cCI6MSwiZXhwIjoxNzgyMDk0NTcyLjEwODc1MywiaWF0IjoxNzc5NTAyNTcyLjEwODc1Mywic3ViIjoiOWJjMTcyZTUzMzgxNDdkOGEzNWMxNDM4ZWE4ZDE1NzcifQ.rm3BIdD1V7FrCwdToT2arErs06xWF7hTqAh0KsCKsdw';

const ACCESS_PASSWORD = 'queue2025';
const ADMIN_EMPLOYEE_ID = "20150465";
const ADMIN_KEYS = ["TENCENT_ACCESS_TOKEN", "RENDER_API_KEY", "GITHUB_TOKEN"];

// Timeouts
const HTTP_TIMEOUT = 5000; // 5 seconds

// Cache TTLs (seconds)
const USER_CACHE_TTL = 120;
const MODEL_CACHE_TTL = 300;
const ORDERS_CACHE_TTL = 120;
const CALC_CACHE_TTL = 30;
const EMPTY_ROW_CACHE_TTL = 30;

// In-memory caches (ephemeral in Workers - will reset on cold start)
let usersCache = { data: null, timestamp: 0 };
let modelsCache = { data: null, timestamp: 0 };
let ordersCache = { data: null, timestamp: 0 };
let filteredCache = { timestamp: 0 };
let emptyRowCache = { row: 0, timestamp: 0 };
let calcResultCache = { key: null, result: null, timestamp: 0 };
let tempRowTracker = {}; // key: submitter_id -> {row_index, timestamp}
const TEMP_ROW_TIMEOUT = 300; // 5 minutes

// Limit date cache and model config cache
let limitDateCache = {};
let modelConfigCache = {};
let modelConfigCacheTime = 0;

// Function to initialize config from env
function initConfig(env) {
  if (env.CLIENT_ID) CLIENT_ID = env.CLIENT_ID;
  if (env.OPEN_ID) OPEN_ID = env.OPEN_ID;
  if (env.ACCESS_TOKEN) ACCESS_TOKEN = env.ACCESS_TOKEN;
}

// Get effective Tencent token (admin-configured or default)
function getEffectiveToken(kvAdminCache) {
  if (kvAdminCache && kvAdminCache.TENCENT_ACCESS_TOKEN) {
    return kvAdminCache.TENCENT_ACCESS_TOKEN;
  }
  return ACCESS_TOKEN;
}

// Get Beijing time string
function getBeijingTimeStr() {
  const now = new Date();
  const bj = new Date(now.getTime() + 8 * 3600000);
  return bj.toISOString().replace('T', ' ').substring(0, 19);
}

// Mask secret value
function maskSecret(value) {
  if (!value) return '';
  const s = String(value);
  if (s.length <= 8) return s[0] + '***' + s[s.length - 1];
  return s.substring(0, 4) + '***' + s.substring(s.length - 4);
}

// Decode JWT expiry
function decodeTokenExpiry(token) {
  if (!token || typeof token !== 'string') return 0;
  const parts = token.split('.');
  if (parts.length !== 3) return 0;
  try {
    const payload = parts[1].padEnd(parts[1].length + (4 - parts[1].length % 4) % 4, '=');
    const data = JSON.parse(atob(payload.replace(/-/g, '+').replace(/_/g, '/')));
    return data.exp || 0;
  } catch (e) {
    return 0;
  }
}

// Parse cell value from Tencent API response
function parseCellValue(cellValue) {
  if (!cellValue) return '';
  if (cellValue.text) {
    const text = cellValue.text.trim();
    // Parse Chinese date format: "6月26日"
    const mmddMatch = text.match(/^(\d{1,2})月(\d{1,2})日$/);
    if (mmddMatch) {
      const year = new Date().getFullYear();
      return `${year}-${String(mmddMatch[1]).padStart(2, '0')}-${String(mmddMatch[2]).padStart(2, '0')}`;
    }
    // Parse "2026年6月26日"
    const ymdMatch = text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
    if (ymdMatch) {
      return `${ymdMatch[1]}-${String(ymdMatch[2]).padStart(2, '0')}-${String(ymdMatch[3]).padStart(2, '0')}`;
    }
    return text;
  }
  if (cellValue.number !== undefined) return String(cellValue.number);
  if (cellValue.time) {
    const t = cellValue.time;
    const result = `${t.year}-${String(t.month).padStart(2, '0')}-${String(t.day).padStart(2, '0')}`;
    if (result === '1899-12-30') return '';
    return result;
  }
  if (cellValue.select) {
    const vals = (cellValue.select.value || []);
    return vals.length > 0 ? vals[0] : '';
  }
  if (cellValue.link) {
    return cellValue.link.text || cellValue.link.url || '';
  }
  return '';
}

// Build cell value for writing to Tencent API
function buildCellValue(value, isDate = false, isNumber = false, fontSize = 14) {
  const cell = {};
  if (!value || String(value).trim() === '') {
    cell.cellValue = { text: '' };
  } else if (isNumber) {
    try {
      cell.cellValue = { number: parseFloat(value) };
    } catch (e) {
      cell.cellValue = { text: String(value) };
    }
  } else if (isDate) {
    try {
      const parts = String(value).split('-');
      if (parts.length === 3 && parts[0].length === 4) {
        cell.cellValue = { time: { year: parseInt(parts[0]), month: parseInt(parts[1]), day: parseInt(parts[2]) } };
      } else {
        cell.cellValue = { text: String(value) };
      }
    } catch (e) {
      cell.cellValue = { text: String(value) };
    }
  } else {
    cell.cellValue = { text: String(value) };
  }
  if (fontSize) {
    const textFormat = { fontSize, font: 'SimSun' };
    cell.cellFormat = { textFormat };
    cell.textFormat = textFormat;
  }
  return cell;
}

// Normalize user key (handle 20150465.0 -> 20150465)
function normalizeUserKey(value) {
  const text = String(value || '').trim();
  if (text.endsWith('.0')) return text.substring(0, text.length - 2);
  return text;
}

export {
  FILE_ID, SHEET_ID, MODEL_FILE_ID, MODEL_SHEET_ID, USER_FILE_ID, USER_SHEET_ID, CONFIG_FILE_ID,
  BASE_URL, CLIENT_ID, OPEN_ID, ACCESS_TOKEN, ACCESS_PASSWORD,
  ADMIN_EMPLOYEE_ID, ADMIN_KEYS, HTTP_TIMEOUT,
  USER_CACHE_TTL, MODEL_CACHE_TTL, ORDERS_CACHE_TTL, CALC_CACHE_TTL, EMPTY_ROW_CACHE_TTL,
  usersCache, modelsCache, ordersCache, filteredCache, emptyRowCache, calcResultCache,
  tempRowTracker, TEMP_ROW_TIMEOUT, limitDateCache, modelConfigCache, modelConfigCacheTime,
  initConfig, getEffectiveToken, getBeijingTimeStr, maskSecret, decodeTokenExpiry,
  parseCellValue, buildCellValue, normalizeUserKey,
};
