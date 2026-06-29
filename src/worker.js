import {
  initConfig, ACCESS_PASSWORD, ADMIN_EMPLOYEE_ID,
  MODEL_FILE_ID, MODEL_SHEET_ID,
  MODEL_CACHE_TTL, CALC_CACHE_TTL,
  usersCache, modelsCache, ordersCache, filteredCache,
  emptyRowCache, calcResultCache, tempRowTracker, TEMP_ROW_TIMEOUT,
  modelConfigCache, modelConfigCacheTime, limitDateCache,
  getEffectiveToken, getBeijingTimeStr, parseCellValue, maskSecret,
} from './config.js';
import {
  handleAuthCheck, handleAuthLogin, handleAuthUsers, handleUpdatePassword,
} from './auth.js';
import {
  handleGetOrders, handleGetOrder, handleCreateOrder,
  handleUpdateOrder, handleDeleteOrder,
  handleClearTempRow, handleCleanupUserTempRows,
  getNextEmptyRow,
} from './orders.js';
import {
  handleAdminStatus, handleAdminValidate, handleAdminUpdate,
  handleAdminHealth, handleAdminTriggerDeploy,
  handleAdminModelConfigs, handleAdminSaveModelConfigs,
} from './admin.js';
import {
  calculateDeliveryDate, refreshCapacityData, preloadAllCapacityData, MODEL_CONFIG,
} from './calc-engine.js';
import { getHeaders, readSheetRange, readSingleCell } from './tencent-api.js';

let nextTempRowBase = 1000;
let tempRowSeq = 0;

const AUTH_PUBLIC_ROUTES = [
  '/auth/login',
  '/auth/users',
  '/api/clear-temp-row',
  '/api/admin/deploy',
];

const ADMIN_REQUIRED_ROUTES = [
  '/api/admin/status',
  '/api/admin/validate',
  '/api/admin/update',
  '/api/admin/health',
];

function isAdminRequired(path, method) {
  if (ADMIN_REQUIRED_ROUTES.includes(path)) return true;
  if (path === '/api/admin/model-configs' && method === 'POST') return true;
  return false;
}

function isAuthRequired(path) {
  return !AUTH_PUBLIC_ROUTES.includes(path);
}

function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status,
    headers: { 'Content-Type': 'application/json' },
  });
}

function checkAuth(request) {
  const password = request.headers.get('X-Access-Password');
  if (password !== ACCESS_PASSWORD) {
    return jsonResponse({ success: false, error: '未授权', need_auth: true }, 401);
  }
  return null;
}

function checkAdmin(request) {
  const employeeId = request.headers.get('X-Employee-Id');
  if (employeeId !== ADMIN_EMPLOYEE_ID) {
    return jsonResponse({ success: false, error: '需要管理员权限' }, 403);
  }
  return null;
}

async function loadKvCache(env) {
  if (!env.ADMIN_KV) return null;
  try {
    const cache = {};
    const secretKeys = ['TENCENT_ACCESS_TOKEN', 'RENDER_API_KEY', 'GITHUB_TOKEN', 'RENDER_SERVICE_ID'];
    for (const key of secretKeys) {
      const val = await env.ADMIN_KV.get(key);
      if (val !== null) cache[key] = val;
    }
    const modelConfigsStr = await env.ADMIN_KV.get('MODEL_CONFIGS');
    if (modelConfigsStr) {
      try {
        cache.MODEL_CONFIGS = JSON.parse(modelConfigsStr);
      } catch (e) {
        console.warn(`[worker] Failed to parse MODEL_CONFIGS from KV: ${e.message}`);
      }
    }
    return cache;
  } catch (e) {
    console.warn(`[worker] Failed to load KV: ${e.message}`);
    return {};
  }
}

async function handleGetModels(request, env, kvAdminCache) {
  const now = Math.floor(Date.now() / 1000);

  if (modelsCache.data !== null && (now - modelsCache.timestamp) < MODEL_CACHE_TTL) {
    return jsonResponse({ success: true, models: modelsCache.data, version: 'cf-v1' });
  }

  // Merge hardcoded MODEL_CONFIG keys with sheet data
  const modelNames = [...Object.keys(MODEL_CONFIG)];

  try {
    const gridData = await readSheetRange(MODEL_SHEET_ID, 'A1:A100', MODEL_FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];
    for (const row of rows) {
      for (const val of (row.values || [])) {
        if (val && val.cellValue) {
          const cellVal = parseCellValue(val.cellValue);
          if (cellVal && !modelNames.includes(cellVal)) {
            modelNames.push(cellVal);
          }
        }
      }
    }
  } catch (e) {
    console.warn(`[models] Sheet read optional, using cached: ${e.message}`);
  }

  // Deduplicate while preserving order
  const uniqueModels = [...new Set(modelNames)];

  modelsCache.data = uniqueModels;
  modelsCache.timestamp = now;

  return jsonResponse({ success: true, models: uniqueModels, version: 'cf-v1' });
}

async function handleCalculateDate(request, env, kvAdminCache) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return jsonResponse({ success: false, error: '请求体解析失败' }, 400);
  }

  const { model, tonnage, expected_date, submitter_id, use_cache } = body;
  const employee_id = submitter_id || '';
  if (!model || tonnage === undefined || !expected_date) {
    return jsonResponse({ success: false, error: '缺少必要参数: model, tonnage, expected_date' }, 400);
  }

  const calcCacheKey = `${model}|${tonnage}|${expected_date}`;
  const now = Math.floor(Date.now() / 1000);

  // 如果前端传了 use_cache=true，优先使用缓存结果快速返回
  if (use_cache && calcResultCache.key === calcCacheKey && (now - calcResultCache.timestamp) < CALC_CACHE_TTL) {
    try {
      const actualEmptyRow = await getNextEmptyRow(SHEET_ID, 2, 2, kvAdminCache);
      return jsonResponse({
        success: true,
        calculated_date: calcResultCache.result.date,
        row_index: actualEmptyRow,
        message: 'from_cache',
      });
    } catch (e) {
      // 缓存回退失败，继续正常计算
    }
  }

  if (calcResultCache.key === calcCacheKey && (now - calcResultCache.timestamp) < CALC_CACHE_TTL) {
    const submitterId = employee_id || 'anonymous';
    const existing = tempRowTracker[submitterId];
    if (existing && (now - existing.timestamp) < TEMP_ROW_TIMEOUT) {
      return jsonResponse({
        success: true,
        calculated_date: calcResultCache.result.date,
        row_index: existing.row_index,
        message: '',
      });
    }
  }

  let result;
  try {
    result = await calculateDeliveryDate(model, String(tonnage), expected_date, kvAdminCache);
  } catch (e) {
    return jsonResponse({ success: false, error: `计算失败: ${e.message}` });
  }

  if (result.error) {
    return jsonResponse({ success: false, error: result.error });
  }

  calcResultCache.key = calcCacheKey;
  calcResultCache.result = result;
  calcResultCache.timestamp = now;

  // 获取实际空白行号（A列为空的第一行），替代临时行号
  let targetRow;
  try {
    targetRow = await getNextEmptyRow(SHEET_ID, 2, 2, kvAdminCache);
  } catch (e) {
    // 如果获取空白行失败，回退到临时行号
    const submitterId = employee_id || 'anonymous';
    const existing = tempRowTracker[submitterId];
    if (existing && (now - existing.timestamp) < TEMP_ROW_TIMEOUT) {
      targetRow = existing.row_index;
      existing.timestamp = now;
    } else {
      tempRowSeq += 1;
      targetRow = nextTempRowBase + tempRowSeq;
      tempRowTracker[submitterId] = { row_index: targetRow, timestamp: now };
    }
  }

  return jsonResponse({
    success: true,
    calculated_date: result.date,
    row_index: targetRow,
    message: '',
  });
}

async function handleRefreshCapacityData(request, env, kvAdminCache) {
  try {
    await refreshCapacityData(kvAdminCache);
    return jsonResponse({
      success: true,
      message: '产能缓存已刷新',
      time: getBeijingTimeStr(),
    });
  } catch (e) {
    return jsonResponse({
      success: false,
      error: `刷新失败: ${e.message}`,
    }, 500);
  }
}

async function handleCapacityPreload(request, env, kvAdminCache) {
  try {
    const result = await preloadAllCapacityData(kvAdminCache);
    return jsonResponse({
      success: true,
      message: `产能数据预加载完成：${result.loaded}/${result.total} 个型号`,
      capacity_data: {}, // 前端不需要原始数据，只需要预热后端缓存
      time: getBeijingTimeStr(),
    });
  } catch (e) {
    return jsonResponse({
      success: false,
      error: `预加载失败: ${e.message}`,
    }, 500);
  }
}

function handleCacheStatus(request, env, kvAdminCache) {
  const now = Math.floor(Date.now() / 1000);

  const cacheInfo = {
    usersCache: {
      hasData: usersCache.data !== null,
      age: usersCache.timestamp ? now - usersCache.timestamp : null,
    },
    modelsCache: {
      hasData: modelsCache.data !== null,
      age: modelsCache.timestamp ? now - modelsCache.timestamp : null,
    },
    ordersCache: {
      hasData: ordersCache.data !== null,
      age: ordersCache.timestamp ? now - ordersCache.timestamp : null,
    },
    filteredCache: {
      timestamp: filteredCache.timestamp || null,
    },
    emptyRowCache: {
      row: emptyRowCache.row || 0,
      age: emptyRowCache.timestamp ? now - emptyRowCache.timestamp : null,
    },
    calcResultCache: {
      hasData: calcResultCache.result !== null,
      key: calcResultCache.key || null,
      age: calcResultCache.timestamp ? now - calcResultCache.timestamp : null,
    },
    tempRowTracker: {
      entries: Object.keys(tempRowTracker).length,
      keys: Object.keys(tempRowTracker),
    },
    modelConfigCache: {
      hasData: Object.keys(modelConfigCache).length > 0,
      age: modelConfigCacheTime ? now - modelConfigCacheTime : null,
    },
    limitDateCache: {
      entries: Object.keys(limitDateCache).length,
    },
    time: getBeijingTimeStr(),
  };

  return jsonResponse({ success: true, data: cacheInfo });
}

async function handleTestConnection(request, env, kvAdminCache) {
  try {
    const value = await readSingleCell('000007', 'A1', null, kvAdminCache);
    const token = getEffectiveToken(kvAdminCache);

    return jsonResponse({
      success: true,
      message: '腾讯文档 API 连接正常',
      token_source: (kvAdminCache && kvAdminCache.TENCENT_ACCESS_TOKEN) ? 'admin_kv' : 'default',
      token_masked: maskSecret(token),
      cell_value: value || '(空)',
      time: getBeijingTimeStr(),
    });
  } catch (e) {
    return jsonResponse({
      success: false,
      error: `连接测试失败: ${e.message}`,
      time: getBeijingTimeStr(),
    }, 500);
  }
}

function handleDiagnoseCalcEngine(request, env, kvAdminCache) {
  const configs = (kvAdminCache && kvAdminCache.MODEL_CONFIGS)
    ? kvAdminCache.MODEL_CONFIGS
    : MODEL_CONFIG;

  const modelCount = Object.keys(configs).length;
  const sampleModels = Object.keys(configs).slice(0, 5);

  return jsonResponse({
    success: true,
    engine: 'calc-engine.js',
    model_count: modelCount,
    sample_models: sampleModels,
    calc_cache: {
      key: calcResultCache.key,
      age: calcResultCache.timestamp
        ? Math.floor(Date.now() / 1000) - calcResultCache.timestamp
        : null,
    },
    temp_tracker_entries: Object.keys(tempRowTracker).length,
    time: getBeijingTimeStr(),
  });
}

async function handleOrdersRoute(method, path, request, env, kvAdminCache) {
  if (path === '/api/orders') {
    if (method === 'POST') return handleCreateOrder(request, env, kvAdminCache);
    if (method === 'GET') return handleGetOrders(request, env, kvAdminCache);
    return jsonResponse({ success: false, error: '方法不允许' }, 405);
  }

  if (path.startsWith('/api/orders/')) {
    const rowMatch = path.match(/^\/api\/orders\/(.+)$/);
    if (!rowMatch) {
      return jsonResponse({ success: false, error: '无效的请求路径' }, 400);
    }
    const row = decodeURIComponent(rowMatch[1]);

    switch (method) {
      case 'GET':
        return handleGetOrder(request, env, kvAdminCache, row);
      case 'PUT':
        return handleUpdateOrder(request, env, kvAdminCache, row);
      case 'DELETE':
        return handleDeleteOrder(request, env, kvAdminCache, row);
      default:
        return jsonResponse({ success: false, error: '方法不允许' }, 405);
    }
  }

  return null;
}

export default {
  async fetch(request, env, ctx) {
    initConfig(env);

    const url = new URL(request.url);
    const path = url.pathname;
    const method = request.method;

    const kvAdminCache = await loadKvCache(env);

    if (path.startsWith('/static/') || path.startsWith('/assets/')) {
      if (env.ASSETS) return env.ASSETS.fetch(request);
      return jsonResponse({ success: false, error: '静态资源未配置' }, 404);
    }

    if (isAuthRequired(path)) {
      const authErr = checkAuth(request);
      if (authErr) return authErr;
    }

    if (isAdminRequired(path, method)) {
      const adminErr = checkAdmin(request);
      if (adminErr) return adminErr;
    }

    if (path === '/api/models' && method === 'GET') {
      return handleGetModels(request, env, kvAdminCache);
    }

    if (path === '/api/calculate-date' && method === 'POST') {
      return handleCalculateDate(request, env, kvAdminCache);
    }

    if (path === '/api/refresh-capacity-data' && method === 'POST') {
      return handleRefreshCapacityData(request, env, kvAdminCache);
    }

    if (path === '/api/capacity-preload' && method === 'GET') {
      return handleCapacityPreload(request, env, kvAdminCache);
    }

    if (path === '/api/cache-status' && method === 'GET') {
      return handleCacheStatus(request, env, kvAdminCache);
    }

    if (path === '/api/test-connection' && method === 'GET') {
      return handleTestConnection(request, env, kvAdminCache);
    }

    if (path === '/api/diag-calc-engine' && method === 'GET') {
      return handleDiagnoseCalcEngine(request, env, kvAdminCache);
    }

    if (path === '/auth/check' && method === 'GET') {
      return handleAuthCheck(request, env, kvAdminCache);
    }

    if (path === '/auth/login' && method === 'POST') {
      return handleAuthLogin(request, env, kvAdminCache);
    }

    if (path === '/auth/users' && method === 'GET') {
      return handleAuthUsers(request, env, kvAdminCache);
    }

    if (path === '/api/users/password' && method === 'PUT') {
      return handleUpdatePassword(request, env, kvAdminCache);
    }

    if (path === '/api/clear-temp-row' && method === 'POST') {
      return handleClearTempRow(request, env, kvAdminCache);
    }

    if (path === '/api/cleanup-user-temp-rows' && method === 'POST') {
      return handleCleanupUserTempRows(request, env, kvAdminCache);
    }

    const ordersResult = await handleOrdersRoute(method, path, request, env, kvAdminCache);
    if (ordersResult) return ordersResult;

    if (path === '/api/admin/status' && method === 'GET') {
      return handleAdminStatus(request, env, kvAdminCache);
    }

    if (path === '/api/admin/validate' && method === 'POST') {
      return handleAdminValidate(request, env, kvAdminCache);
    }

    if (path === '/api/admin/update' && method === 'POST') {
      return handleAdminUpdate(request, env, kvAdminCache);
    }

    if (path === '/api/admin/health' && method === 'GET') {
      return handleAdminHealth(request, env, kvAdminCache);
    }

    if (path === '/api/admin/deploy' && method === 'POST') {
      return handleAdminTriggerDeploy(request, env, kvAdminCache);
    }

    if (path === '/api/admin/model-configs') {
      if (method === 'GET') {
        return handleAdminModelConfigs(request, env, kvAdminCache);
      }
      if (method === 'POST') {
        return handleAdminSaveModelConfigs(request, env, kvAdminCache);
      }
      return jsonResponse({ success: false, error: '方法不允许' }, 405);
    }

    if (env.ASSETS && (path === '/' || !path.startsWith('/api/') && !path.startsWith('/auth/'))) {
      return env.ASSETS.fetch(request);
    }

    return jsonResponse({ success: false, error: `路由未找到: ${method} ${path}` }, 404);
  },
};
