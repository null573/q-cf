import {
  ADMIN_KEYS, ADMIN_EMPLOYEE_ID, ACCESS_TOKEN,
  maskSecret, decodeTokenExpiry, getBeijingTimeStr,
  USER_FILE_ID, USER_SHEET_ID, HTTP_TIMEOUT,
} from './config.js';
import { readSingleCell, getHeaders } from './tencent-api.js';
import { MODEL_CONFIG } from './calc-engine.js';

function getAdminSecret(name, kvAdminCache) {
  if (kvAdminCache && kvAdminCache[name] !== undefined && kvAdminCache[name] !== '') {
    return kvAdminCache[name];
  }
  const defaults = {
    TENCENT_ACCESS_TOKEN: ACCESS_TOKEN,
    RENDER_API_KEY: '',
    GITHUB_TOKEN: '',
    RENDER_SERVICE_ID: '',
  };
  return defaults[name] || '';
}

async function setAdminSecret(name, value, env) {
  if (env.ADMIN_KV) {
    await env.ADMIN_KV.put(name, value);
  }

  const renderKey = (env.ADMIN_KV)
    ? (name === 'RENDER_API_KEY' ? value : (await env.ADMIN_KV.get('RENDER_API_KEY') || ''))
    : '';
  const serviceId = env.RENDER_SERVICE_ID || '';

  if (renderKey && serviceId && name !== 'RENDER_API_KEY') {
    try {
      await fetch(`https://api.render.com/v1/services/${serviceId}/env-vars/${name}`, {
        method: 'PUT',
        headers: {
          'Authorization': `Bearer ${renderKey}`,
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({ value: value }),
        signal: AbortSignal.timeout(HTTP_TIMEOUT),
      });
    } catch (e) {
      console.warn(`[admin] Failed to sync ${name} to Render: ${e.message}`);
    }
  }
}

function summarizeAdminKey(name, kvAdminCache) {
  const value = getAdminSecret(name, kvAdminCache);
  const hasValue = Boolean(value && value.length > 0);

  if (name === 'TENCENT_ACCESS_TOKEN') {
    const expiry = hasValue ? decodeTokenExpiry(value) : 0;
    const now = Math.floor(Date.now() / 1000);
    return {
      name,
      present: hasValue,
      masked: hasValue ? maskSecret(value) : '',
      expiry: expiry,
      expired: hasValue && expiry > 0 && expiry < now,
      expires_at: expiry > 0 ? new Date(expiry * 1000).toISOString() : null,
    };
  }

  return {
    name,
    present: hasValue,
    masked: hasValue ? maskSecret(value) : '',
  };
}

async function validateTencentToken(token) {
  try {
    const value = await readSingleCell(USER_SHEET_ID, 'A1', USER_FILE_ID, { TENCENT_ACCESS_TOKEN: token });
    if (value === '') {
      return { valid: false, message: '令牌无效：无法读取表格数据' };
    }
    return { valid: true, message: '腾讯文档令牌有效' };
  } catch (e) {
    return { valid: false, message: `令牌验证失败: ${e.message}` };
  }
}

async function validateRenderKey(token) {
  try {
    const resp = await fetch('https://api.render.com/v1/services?limit=1', {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      signal: AbortSignal.timeout(HTTP_TIMEOUT),
    });
    if (resp.ok || resp.status === 200) {
      return { valid: true, message: 'Render API 密钥有效' };
    }
    if (resp.status === 401 || resp.status === 403) {
      return { valid: false, message: 'Render API 密钥无效：认证失败' };
    }
    const text = await resp.text();
    return { valid: false, message: `Render API 返回 HTTP ${resp.status}: ${text.substring(0, 200)}` };
  } catch (e) {
    return { valid: false, message: `Render API 连接失败: ${e.message}` };
  }
}

async function validateGithubToken(token) {
  try {
    const resp = await fetch('https://api.github.com/user', {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'Cloudflare-Worker-Admin',
      },
      signal: AbortSignal.timeout(HTTP_TIMEOUT),
    });
    if (resp.ok) {
      const data = await resp.json();
      return { valid: true, message: `GitHub 令牌有效 (${data.login || 'unknown'})` };
    }
    if (resp.status === 401) {
      return { valid: false, message: 'GitHub 令牌无效：认证失败' };
    }
    return { valid: false, message: `GitHub API 返回 HTTP ${resp.status}` };
  } catch (e) {
    return { valid: false, message: `GitHub API 连接失败: ${e.message}` };
  }
}

async function validateAdminKey(key, value) {
  if (!value || String(value).trim() === '') {
    return { valid: false, key, message: '值不能为空' };
  }

  switch (key) {
    case 'TENCENT_ACCESS_TOKEN':
      return validateTencentToken(String(value));
    case 'RENDER_API_KEY':
      return validateRenderKey(String(value));
    case 'GITHUB_TOKEN':
      return validateGithubToken(String(value));
    default:
      return { valid: false, key, message: `未知的凭据类型: ${key}` };
  }
}

async function handleAdminStatus(request, env, kvAdminCache) {
  const statuses = {};
  for (const key of ADMIN_KEYS) {
    statuses[key] = summarizeAdminKey(key, kvAdminCache);
  }
  return new Response(JSON.stringify({ success: true, data: statuses, time: getBeijingTimeStr() }), {
    headers: { 'Content-Type': 'application/json' },
  });
}

async function handleAdminValidate(request, env, kvAdminCache) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: '请求体解析失败' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const { key, value } = body;
  if (!key || value === undefined || value === null) {
    return new Response(JSON.stringify({ success: false, error: '缺少 key 或 value' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const result = await validateAdminKey(key, value);
  return new Response(JSON.stringify({ success: true, key, ...result }), {
    headers: { 'Content-Type': 'application/json' },
  });
}

async function handleAdminUpdate(request, env, kvAdminCache) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: '请求体解析失败' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const { key, value } = body;
  if (!key || value === undefined || value === null) {
    return new Response(JSON.stringify({ success: false, error: '缺少 key 或 value' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  if (!ADMIN_KEYS.includes(key)) {
    return new Response(JSON.stringify({ success: false, error: `无效的凭据类型: ${key}` }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  await setAdminSecret(key, String(value), env);

  if (kvAdminCache) {
    kvAdminCache[key] = String(value);
  }

  return new Response(JSON.stringify({
    success: true,
    message: `已保存 ${key}`,
    time: getBeijingTimeStr(),
  }), {
    headers: { 'Content-Type': 'application/json' },
  });
}

async function handleAdminHealth(request, env, kvAdminCache) {
  const results = {};
  for (const key of ADMIN_KEYS) {
    const secret = getAdminSecret(key, kvAdminCache);
    if (!secret) {
      results[key] = { valid: false, message: '未配置' };
    } else {
      try {
        results[key] = await validateAdminKey(key, secret);
      } catch (e) {
        results[key] = { valid: false, message: `检查失败: ${e.message}` };
      }
    }
  }

  const allValid = Object.values(results).every(r => r.valid === true);

  return new Response(JSON.stringify({
    success: true,
    healthy: allValid,
    checks: results,
    time: getBeijingTimeStr(),
  }), {
    headers: { 'Content-Type': 'application/json' },
  });
}

async function handleAdminTriggerDeploy(request, env, kvAdminCache) {
  const renderKey = env.ADMIN_KV
    ? (await env.ADMIN_KV.get('RENDER_API_KEY') || '')
    : '';
  const serviceId = env.RENDER_SERVICE_ID
    || (kvAdminCache && kvAdminCache.RENDER_SERVICE_ID)
    || '';

  if (!renderKey) {
    return new Response(JSON.stringify({
      success: false,
      error: 'Render API 密钥未配置，请先在管理页面配置 RENDER_API_KEY',
    }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  if (!serviceId) {
    return new Response(JSON.stringify({
      success: false,
      error: 'Render 服务 ID 未配置，请设置 RENDER_SERVICE_ID 环境变量',
    }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  try {
    const resp = await fetch(`https://api.render.com/v1/services/${serviceId}/deploys`, {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${renderKey}`,
        'Content-Type': 'application/json',
      },
      signal: AbortSignal.timeout(HTTP_TIMEOUT),
    });

    const data = await resp.json().catch(() => ({}));

    return new Response(JSON.stringify({
      success: resp.ok || resp.status === 201,
      message: (resp.ok || resp.status === 201) ? '部署已触发' : `部署触发失败: HTTP ${resp.status}`,
      deploy: data,
      time: getBeijingTimeStr(),
    }), {
      headers: { 'Content-Type': 'application/json' },
    });
  } catch (e) {
    return new Response(JSON.stringify({
      success: false,
      error: `部署触发失败: ${e.message}`,
    }), {
      status: 500,
      headers: { 'Content-Type': 'application/json' },
    });
  }
}

async function handleAdminModelConfigs(request, env, kvAdminCache) {
  let configs = {};

  if (kvAdminCache && kvAdminCache.MODEL_CONFIGS) {
    configs = { ...kvAdminCache.MODEL_CONFIGS };
  } else {
    configs = { ...MODEL_CONFIG };
  }

  return new Response(JSON.stringify({ success: true, data: configs }), {
    headers: { 'Content-Type': 'application/json' },
  });
}

async function handleAdminSaveModelConfigs(request, env, kvAdminCache) {
  let body;
  try {
    body = await request.json();
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: '请求体解析失败' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  const configs = body.configs || body;

  if (!configs || typeof configs !== 'object' || Array.isArray(configs)) {
    return new Response(JSON.stringify({ success: false, error: '无效的配置数据，需要 JSON 对象' }), {
      status: 400,
      headers: { 'Content-Type': 'application/json' },
    });
  }

  if (env.ADMIN_KV) {
    await env.ADMIN_KV.put('MODEL_CONFIGS', JSON.stringify(configs));
  }

  if (kvAdminCache) {
    kvAdminCache.MODEL_CONFIGS = configs;
  }

  return new Response(JSON.stringify({
    success: true,
    message: '模型配置已保存',
    count: Object.keys(configs).length,
    time: getBeijingTimeStr(),
  }), {
    headers: { 'Content-Type': 'application/json' },
  });
}

export {
  getAdminSecret,
  setAdminSecret,
  summarizeAdminKey,
  validateTencentToken,
  validateRenderKey,
  validateGithubToken,
  validateAdminKey,
  handleAdminStatus,
  handleAdminValidate,
  handleAdminUpdate,
  handleAdminHealth,
  handleAdminTriggerDeploy,
  handleAdminModelConfigs,
  handleAdminSaveModelConfigs,
};
