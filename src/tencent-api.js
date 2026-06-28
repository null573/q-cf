import { BASE_URL, HTTP_TIMEOUT, getEffectiveToken } from './config.js';

// Get headers for Tencent API
function getHeaders(kvAdminCache) {
  const token = getEffectiveToken(kvAdminCache);
  return {
    'Content-Type': 'application/json',
    'Access-Token': token,
    'Open-Id': '9bc172e5338147d8a35c1438ea8d1577',
    'Client-Id': 'da815d1227294457b43413bdc16e3e90',
  };
}

// Read a range from a sheet
async function readSheetRange(sheetId, rangeStr, fileId = null, kvAdminCache = null) {
  const fid = fileId || 'DRnhDemRIS25mdnFF';
  const url = `${BASE_URL}/files/${fid}/${sheetId}/${rangeStr}`;
  try {
    const resp = await fetch(url, {
      headers: getHeaders(kvAdminCache),
      signal: AbortSignal.timeout(HTTP_TIMEOUT),
    });
    if (resp.status === 200) {
      const data = await resp.json();
      if (data.code && data.code !== 0) {
        console.warn(`[WARN] Tencent API error: ${data.code} ${data.message} for ${rangeStr}`);
        return {};
      }
      return data.gridData || {};
    }
    console.warn(`[WARN] readSheetRange HTTP ${resp.status}: ${rangeStr}`);
  } catch (e) {
    console.warn(`[WARN] readSheetRange error: ${e.message} for ${rangeStr}`);
  }
  return {};
}

// Read a single cell
async function readSingleCell(sheetId, cell, fileId = null, kvAdminCache = null) {
  const gridData = await readSheetRange(sheetId, `${cell}:${cell}`, fileId, kvAdminCache);
  const rows = gridData.rows || [];
  if (rows.length > 0) {
    for (const v of (rows[0].values || [])) {
      if (v.cellValue) {
        const { parseCellValue } = await import('./config.js');
        return parseCellValue(v.cellValue);
      }
    }
  }
  return '';
}

// Batch update
async function batchUpdate(requests, fileId = null, kvAdminCache = null) {
  const fid = fileId || 'DRnhDemRIS25mdnFF';
  const url = `${BASE_URL}/files/${fid}/batchUpdate`;
  try {
    const resp = await fetch(url, {
      method: 'POST',
      headers: getHeaders(kvAdminCache),
      body: JSON.stringify({ requests }),
      signal: AbortSignal.timeout(HTTP_TIMEOUT),
    });
    return resp;
  } catch (e) {
    console.warn(`[WARN] batchUpdate error: ${e.message}`);
    throw e;
  }
}

export { getHeaders, readSheetRange, readSingleCell, batchUpdate };
