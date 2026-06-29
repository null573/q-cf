import {
  FILE_ID, SHEET_ID, ORDERS_CACHE_TTL, EMPTY_ROW_CACHE_TTL, TEMP_ROW_TIMEOUT,
  ordersCache, filteredCache, emptyRowCache, tempRowTracker,
  parseCellValue, buildCellValue, normalizeUserKey, getBeijingTimeStr
} from './config.js';
import { readSheetRange, batchUpdate } from './tencent-api.js';
import { calculateDeliveryDate } from './calc-engine.js';
import {
  canOperateOrder, orderMatchesExpected, resolveSubmitterName,
  getUserById, normalizeViewMode
} from './auth.js';

async function getNextEmptyRow(sheetId, startFrom = 2, maxBatches = 4, kvAdminCache) {
  const now = Date.now();
  if (emptyRowCache.row >= startFrom && (now - emptyRowCache.timestamp) < EMPTY_ROW_CACHE_TTL * 1000) {
    return emptyRowCache.row;
  }

  const batchSize = 200;
  let batchesDone = 0;

  for (let offset = startFrom - 1; offset < 2000; offset += batchSize) {
    batchesDone++;
    if (batchesDone > maxBatches) break;

    const start = offset + 1;
    const end = offset + batchSize;
    const rangeStr = `A${start}:A${end}`;
    const gridData = await readSheetRange(sheetId, rangeStr, FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];

    if (!rows || rows.length === 0) break;

    for (let i = 0; i < rows.length; i++) {
      const actualRow = start + i;
      if (actualRow < 2) continue;

      let hasData = false;
      const values = rows[i].values || [];
      for (const v of values) {
        const cv = v.cellValue;
        if (cv) {
          const text = parseCellValue(cv);
          if (text.trim()) {
            hasData = true;
            break;
          }
        }
      }
      if (!hasData) {
        emptyRowCache.row = actualRow;
        emptyRowCache.timestamp = now;
        return actualRow;
      }
    }

    if (rows.length < batchSize) {
      const actualRow = start + rows.length;
      emptyRowCache.row = actualRow;
      emptyRowCache.timestamp = now;
      return actualRow;
    }
  }

  const defaultRow = startFrom;
  emptyRowCache.row = defaultRow;
  emptyRowCache.timestamp = now;
  return defaultRow;
}

async function ensureSheetRows(minRowCount, kvAdminCache) {
  try {
    const url = `https://docs.qq.com/openapi/spreadsheet/v3/files/${FILE_ID}`;
    const { getHeaders } = await import('./tencent-api.js');
    const headers = getHeaders(kvAdminCache);

    const resp = await fetch(url, {
      headers,
      signal: AbortSignal.timeout(5000),
    });

    if (resp.status !== 200) return false;

    const data = await resp.json();
    if (data.code && data.code !== 0) return false;

    const sheets = data.data ? (data.data.sheets || []) : [];
    let currentRowCount = 0;
    for (const s of sheets) {
      if (s.sheetID === SHEET_ID) {
        currentRowCount = s.rowCount || 0;
        if (currentRowCount <= 0) {
          const gp = s.gridProperties || {};
          currentRowCount = gp.rowCount || 0;
        }
        break;
      }
    }

    if (currentRowCount >= minRowCount) return true;

    const rowsToAdd = Math.max(500, minRowCount - currentRowCount);
    const body = {
      requests: [{
        insertDimension: {
          range: {
            sheetID: SHEET_ID,
            dimension: 'ROWS',
            startIndex: currentRowCount + 1,
            endIndex: currentRowCount + 1 + rowsToAdd
          }
        }
      }]
    };

    const updateResp = await batchUpdate(body.requests, FILE_ID, kvAdminCache);
    if (updateResp.status === 200) {
      const result = await updateResp.json();
      if (result.ret === 0 || result.responses) {
        return true;
      }
    }
    return false;
  } catch (e) {
    console.warn(`[WARN] ensureSheetRows error: ${e.message}`);
    return false;
  }
}

function isDateString(value) {
  if (!value) return false;
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return true;
  if (/^\d{4}年\d{1,2}月\d{1,2}日$/.test(text)) return true;
  if (/^\d{1,2}月\d{1,2}日$/.test(text)) return true;
  return false;
}

async function writeOrderRow(rowIndex0based, model, tonnage, customer, expectedDate,
    calculatedDate, queueDate, submitter, remark, serialNo, submitterId, submitTime, kvAdminCache) {
  const queueDateIsDate = isDateString(queueDate);

  const requests = [
    {
      updateRangeRequest: {
        sheetId: SHEET_ID,
        gridData: {
          startRow: rowIndex0based,
          startColumn: 0,
          rows: [{
            values: [
              buildCellValue(model),
              buildCellValue(tonnage, false, true),
              buildCellValue(customer),
              buildCellValue(expectedDate, true),
            ]
          }]
        }
      }
    },
    {
      updateRangeRequest: {
        sheetId: SHEET_ID,
        gridData: {
          startRow: rowIndex0based,
          startColumn: 5,
          rows: [{
            values: [
              buildCellValue(queueDate, queueDateIsDate),
              buildCellValue(submitter),
              buildCellValue(remark),
              buildCellValue(serialNo),
              buildCellValue(''),
              buildCellValue(submitterId),
              buildCellValue(submitTime),
            ]
          }]
        }
      }
    }
  ];

  return await batchUpdate(requests, FILE_ID, kvAdminCache);
}

async function deleteRow(rowIndex1based, kvAdminCache) {
  if (rowIndex1based < 2) {
    throw new Error('无效行号，不能删除表头或不存在的行');
  }

  const rowIndex0based = rowIndex1based - 1;
  const requests = [
    {
      updateRangeRequest: {
        sheetId: SHEET_ID,
        gridData: {
          startRow: rowIndex0based,
          startColumn: 1,
          rows: [{ values: [buildCellValue('')] }]
        }
      }
    },
    {
      updateRangeRequest: {
        sheetId: SHEET_ID,
        gridData: {
          startRow: rowIndex0based,
          startColumn: 7,
          rows: [{ values: [buildCellValue('')] }]
        }
      }
    },
    {
      updateRangeRequest: {
        sheetId: SHEET_ID,
        gridData: {
          startRow: rowIndex0based,
          startColumn: 9,
          rows: [{ values: [buildCellValue('DELETED')] }]
        }
      }
    }
  ];

  return await batchUpdate(requests, FILE_ID, kvAdminCache);
}

async function clearTempRow(rowIndex1based, kvAdminCache) {
  if (rowIndex1based < 2) return;

  const gridData = await readSheetRange(SHEET_ID, `K${rowIndex1based}:K${rowIndex1based}`, FILE_ID, kvAdminCache);
  const rows = gridData.rows || [];
  if (rows.length > 0) {
    const values = rows[0].values || [];
    if (values.length > 0) {
      const cv = values[0].cellValue;
      if (cv && parseCellValue(cv).trim()) {
        return;
      }
    }
  }

  const requests = [{
    updateRangeRequest: {
      sheetId: SHEET_ID,
      gridData: {
        startRow: rowIndex1based - 1,
        startColumn: 0,
        rows: [{
          values: [
            buildCellValue(''),
            buildCellValue(''),
            buildCellValue(''),
            buildCellValue(''),
          ]
        }]
      }
    }
  }];

  return await batchUpdate(requests, FILE_ID, kvAdminCache);
}

async function fetchAllOrdersRaw(kvAdminCache) {
  const now = Date.now();
  const cacheValid = ordersCache.data && ordersCache.data.length > 0
      && (now - ordersCache.timestamp) < ORDERS_CACHE_TTL * 1000;
  if (cacheValid && !ordersCache._refreshFlag) {
    return ordersCache.data;
  }

  let lastDataRow = 1;
  const batchSize = 500;

  for (let offset = 0; offset < 2000; offset += batchSize) {
    const start = offset + 1;
    const end = offset + batchSize;
    const rangeStr = `A${start}:A${end}`;
    const gridData = await readSheetRange(SHEET_ID, rangeStr, FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];

    for (let i = 0; i < rows.length; i++) {
      const actualRow = start + i;
      const values = rows[i].values || [];
      if (values.length > 0) {
        const cv = values[0].cellValue;
        if (cv) {
          const text = parseCellValue(cv);
          if (text.trim()) {
            lastDataRow = Math.max(lastDataRow, actualRow);
          }
        }
      }
    }
  }

  if (lastDataRow <= 1) {
    if (ordersCache.data && ordersCache.data.length > 0) {
      return ordersCache.data;
    }
    return [];
  }

  const allRowsByOffset = {};
  const dataBatchSize = 200;

  for (let offset = 1; offset < lastDataRow; offset += dataBatchSize) {
    const start = offset + 1;
    const end = Math.min(offset + dataBatchSize, lastDataRow);
    const rangeStr = `A${start}:L${end}`;
    const gridData = await readSheetRange(SHEET_ID, rangeStr, FILE_ID, kvAdminCache);
    allRowsByOffset[offset] = gridData.rows || [];
  }

  const orders = [];
  const offsets = Object.keys(allRowsByOffset).map(Number).sort((a, b) => a - b);

  for (const offset of offsets) {
    const rows = allRowsByOffset[offset];
    const startRow = offset + 1;

    for (let i = 0; i < rows.length; i++) {
      const actualRow = startRow + i;
      const values = rows[i].values || [];
      if (!values.length) continue;

      function getCol(idx) {
        if (idx < values.length) {
          const cv = values[idx].cellValue;
          if (cv) return parseCellValue(cv);
        }
        return '';
      }

      const rowData = [getCol(0), getCol(1), getCol(2), getCol(3), getCol(4),
                       getCol(5), getCol(6), getCol(7), getCol(8), getCol(9),
                       getCol(10), getCol(11)];

      if (!rowData[0]) continue;

      if (rowData[9] && rowData[9].trim().toUpperCase() === 'DELETED') continue;

      orders.push({
        row_index: actualRow,
        model: rowData[0],
        tonnage: rowData[1],
        customer: rowData[2],
        expected_date: rowData[3],
        calculated_date: rowData[4],
        queue_date: rowData[5],
        submitter: rowData[6],
        remark: rowData[7],
        serial_no: rowData[8],
        last_entry: rowData[9],
        submitter_id: rowData[10],
        submit_time: rowData[11]
      });
    }
  }

  if (!orders.length && ordersCache.data && ordersCache.data.length > 0) {
    return ordersCache.data;
  }

  if (orders.length) {
    ordersCache.data = orders;
    ordersCache.timestamp = now;
    filteredCache.timestamp = 0;
  }
  return orders;
}

async function getFilteredOrders(submitterId, currentUser, viewMode, submitterName, kvAdminCache) {
  const now = Date.now();
  const submitterNameResolved = await resolveSubmitterName(submitterId, submitterName, kvAdminCache);
  const accessLevel = (currentUser || {}).access_level || 'self';
  const cacheKey = `${accessLevel}:${viewMode}:${normalizeUserKey(submitterId)}:${submitterNameResolved}:${(currentUser || {}).department || ''}`;

  if (filteredCache[cacheKey] !== undefined
      && filteredCache.timestamp === ordersCache.timestamp
      && (now - ordersCache.timestamp) < ORDERS_CACHE_TTL * 1000) {
    return filteredCache[cacheKey];
  }

  const allOrders = await fetchAllOrdersRaw(kvAdminCache);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const orders = [];
  for (const order of allOrders) {
    const permitted = await canOperateOrder(order, currentUser, submitterId, submitterNameResolved, viewMode, kvAdminCache);
    if (!permitted) continue;

    const expectedDateStr = order.expected_date;
    if (expectedDateStr) {
      try {
        const expectedDate = new Date(expectedDateStr + 'T00:00:00');
        if (expectedDate < today) continue;
      } catch (e) {
        // ignore parse errors
      }
    }

    orders.push(order);
  }

  orders.sort((a, b) => {
    const qdA = a.queue_date || '';
    const qdB = b.queue_date || '';
    if (qdA && qdB && qdA.length >= 10 && qdB.length >= 10) {
      return qdA.localeCompare(qdB);
    }
    if (qdA && qdA.length >= 10) return -1;
    if (qdB && qdB.length >= 10) return 1;
    return 0;
  });

  filteredCache[cacheKey] = orders;
  filteredCache.timestamp = ordersCache.timestamp;
  return orders;
}

function clearOrderCaches() {
  ordersCache.data = null;
  ordersCache.timestamp = 0;
  const keys = Object.keys(filteredCache).filter(k => k !== 'timestamp');
  for (const key of keys) {
    delete filteredCache[key];
  }
  filteredCache.timestamp = 0;
}

function parseOrderFromRow(values, rowIndex) {
  function getCol(idx) {
    if (idx < values.length) {
      const cv = values[idx].cellValue;
      if (cv) return parseCellValue(cv);
    }
    return '';
  }

  return {
    row_index: rowIndex,
    model: getCol(0),
    tonnage: getCol(1),
    customer: getCol(2),
    expected_date: getCol(3),
    calculated_date: getCol(4),
    queue_date: getCol(5),
    submitter: getCol(6),
    remark: getCol(7),
    serial_no: getCol(8),
    last_entry: getCol(9),
    submitter_id: getCol(10),
    submit_time: getCol(11)
  };
}

async function handleGetOrders(request, env, kvAdminCache) {
  try {
    const url = new URL(request.url);
    if (url.searchParams.get('refresh') === '1') {
      ordersCache.timestamp = 0;
    }

    const submitterId = url.searchParams.get('submitter_id') || '';
    const submitterName = url.searchParams.get('submitter_name') || '';
    const currentUser = await getUserById(submitterId, kvAdminCache) || {};
    const isAdmin = currentUser.access_level === 'admin';
    const requestedViewMode = url.searchParams.get('view_mode') || 'mine';
    const viewMode = normalizeViewMode(currentUser, requestedViewMode);
    const modelFilter = (url.searchParams.get('model_filter') || '').trim();
    const customerFilter = (url.searchParams.get('customer_filter') || '').trim().toLowerCase();
    const sortType = (url.searchParams.get('sort') || '').trim();
    let page = parseInt(url.searchParams.get('page') || '1', 10);
    let perPage = parseInt(url.searchParams.get('per_page') || '20', 10);
    if (page < 1) page = 1;
    if (perPage < 1) perPage = 20;
    if (perPage > 100) perPage = 100;

    let orders = await getFilteredOrders(submitterId, currentUser, viewMode, submitterName, kvAdminCache);

    if (modelFilter) {
      orders = orders.filter(o => o.model === modelFilter);
    }
    if (customerFilter) {
      orders = orders.filter(o => o.customer.toLowerCase().includes(customerFilter));
    }
    if (sortType) {
      let keyFn;
      if (sortType === 'model') {
        keyFn = o => o.model || '';
      } else if (sortType === 'queueDate') {
        keyFn = o => o.queue_date || '9999-12-31';
      } else if (sortType === 'tonnage') {
        keyFn = o => {
          try { return parseFloat(o.tonnage || 0); } catch (e) { return 0; }
        };
      } else {
        keyFn = () => '';
      }
      orders = [...orders].sort((a, b) => {
        const va = keyFn(a);
        const vb = keyFn(b);
        if (typeof va === 'number' && typeof vb === 'number') return va - vb;
        return String(va).localeCompare(String(vb));
      });
    }

    const total = orders.length;
    const startIdx = (page - 1) * perPage;
    const endIdx = startIdx + perPage;
    const paginatedOrders = orders.slice(startIdx, endIdx);

    return new Response(JSON.stringify({
      success: true,
      orders: paginatedOrders,
      is_admin: isAdmin,
      access_level: currentUser.access_level || 'self',
      department: currentUser.department || '',
      view_mode: viewMode,
      pagination: {
        page,
        per_page: perPage,
        total,
        total_pages: Math.ceil(total / perPage)
      }
    }), { headers: { 'Content-Type': 'application/json' } });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleGetOrder(request, rowIndex, env, kvAdminCache) {
  try {
    const url = new URL(request.url);
    const submitterId = url.searchParams.get('submitter_id') || '';
    const submitterName = url.searchParams.get('submitter_name') || '';
    const currentUser = await getUserById(submitterId, kvAdminCache) || {};

    const gridData = await readSheetRange(SHEET_ID, `A${rowIndex}:L${rowIndex}`, FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];
    if (!rows.length) {
      return new Response(JSON.stringify({ success: false, error: '订单不存在' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const values = rows[0].values || [];
    if (!values.length) {
      return new Response(JSON.stringify({ success: false, error: '订单不存在' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const order = parseOrderFromRow(values, rowIndex);

    const permitted = await canOperateOrder(order, currentUser, submitterId, submitterName, 'all', kvAdminCache);
    if (!permitted) {
      return new Response(JSON.stringify({ success: false, error: '无权查看他人订单' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    return new Response(JSON.stringify({ success: true, order }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleCreateOrder(request, env, kvAdminCache) {
  try {
    const data = await request.json();
    const model = data.model || '';
    const tonnage = data.tonnage || '';
    const customer = data.customer || '';
    const expectedDate = data.expected_date || '';
    const queueDate = data.queue_date || '';
    const submitter = data.submitter || '未知用户';
    const submitterId = data.submitter_id || '';

    const remark = `${tonnage}${customer}`;
    const submitTime = getBeijingTimeStr();

    // 查找"自助排队"表A列为空的第一个空白行
    const targetRow = await getNextEmptyRow(SHEET_ID, 2, 4, kvAdminCache);

    // 确保表格有足够的行，如满则自动创建新行
    await ensureSheetRows(targetRow + 10, kvAdminCache);

    const writeIdx = targetRow - 1;
    const serialNo = String(targetRow);
    const resp = await writeOrderRow(
      writeIdx, model, tonnage, customer, expectedDate,
      '', queueDate, submitter, remark, serialNo, submitterId, submitTime, kvAdminCache
    );
    const result = await resp.json();

    if (result.responses) {
      const updated = (result.responses[0] && result.responses[0].updateRangeResponse)
          ? result.responses[0].updateRangeResponse.updatedCells : 0;
      if (updated > 0) {
        // 清除该用户和该行的临时行跟踪
        const tempKey = `${submitterId}`;
        if (tempRowTracker[tempKey]) {
          delete tempRowTracker[tempKey];
        }
        // 清除空白行缓存，因为已占用
        emptyRowCache.row = 0;
        emptyRowCache.timestamp = 0;
        clearOrderCaches();
        return new Response(JSON.stringify({ success: true, message: '订单创建成功', row_index: targetRow }), {
          headers: { 'Content-Type': 'application/json' }
        });
      }
      return new Response(JSON.stringify({ success: false, error: '写入0个单元格' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    return new Response(JSON.stringify({ success: false, error: JSON.stringify(result) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleUpdateOrder(request, rowIndex, env, kvAdminCache) {
  try {
    const data = await request.json();
    const model = data.model || '';
    const tonnage = data.tonnage || '';
    const customer = data.customer || '';
    const expectedDate = data.expected_date || '';
    const queueDate = data.queue_date || '';
    const submitter = data.submitter || '';
    const submitterId = data.submitter_id || '';

    const remark = `${tonnage}${customer}`;
    const currentUser = await getUserById(submitterId, kvAdminCache) || {};

    const gridData = await readSheetRange(SHEET_ID, `A${rowIndex}:L${rowIndex}`, FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];
    if (rows.length) {
      const origValues = rows[0].values || [];
      const origData = origValues.map(v => parseCellValue(v.cellValue));
      const originalTonnage = origData.length > 1 ? origData[1] : '0';
      const originalOrder = {
        submitter: origData.length > 6 ? origData[6] : '',
        submitter_id: origData.length > 10 ? origData[10] : ''
      };

      const permitted = await canOperateOrder(originalOrder, currentUser, submitterId, submitter, 'all', kvAdminCache);
      if (!permitted) {
        return new Response(JSON.stringify({ success: false, error: '无权修改他人订单' }), {
          headers: { 'Content-Type': 'application/json' }
        });
      }

      try {
        if (parseFloat(tonnage) > parseFloat(originalTonnage)) {
          return new Response(JSON.stringify({ success: false, error: '吨位只能改小不能改大' }), {
            headers: { 'Content-Type': 'application/json' }
          });
        }
      } catch (e) {
        // ignore parse errors
      }
    } else {
      return new Response(JSON.stringify({ success: false, error: '订单不存在' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const writeIdx = rowIndex - 1;
    let calculatedDate = '';
    try {
      const calcResult = await calculateDeliveryDate(model, tonnage, expectedDate, kvAdminCache);
      calculatedDate = calcResult.date || '';
    } catch (e) {
      calculatedDate = '';
    }

    const resp = await writeOrderRow(
      writeIdx, model, tonnage, customer, expectedDate,
      calculatedDate, queueDate, submitter, remark, String(rowIndex), submitterId,
      getBeijingTimeStr(), kvAdminCache
    );
    const result = await resp.json();

    if (result.responses) {
      clearOrderCaches();
      return new Response(JSON.stringify({ success: true, message: '订单修改成功' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }
    return new Response(JSON.stringify({ success: false, error: JSON.stringify(result) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleDeleteOrder(request, rowIndex, env, kvAdminCache) {
  try {
    const data = await request.json().catch(() => ({}));
    const expectedOrder = data.order || data;
    const url = new URL(request.url);
    const submitterId = url.searchParams.get('submitter_id') || '';
    const submitterName = url.searchParams.get('submitter_name') || '';
    const currentUser = await getUserById(submitterId, kvAdminCache) || {};

    const gridData = await readSheetRange(SHEET_ID, `A${rowIndex}:L${rowIndex}`, FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];
    if (!rows.length) {
      return new Response(JSON.stringify({ success: false, error: '订单不存在' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const origValues = rows[0].values || [];
    const origData = origValues.map(v => parseCellValue(v.cellValue));
    const originalOrder = {
      model: origData.length > 0 ? origData[0] : '',
      customer: origData.length > 2 ? origData[2] : '',
      submitter: origData.length > 6 ? origData[6] : '',
      submitter_id: origData.length > 10 ? origData[10] : '',
      submit_time: origData.length > 11 ? origData[11] : ''
    };

    if (!orderMatchesExpected(originalOrder, expectedOrder)) {
      return new Response(JSON.stringify({ success: false, error: '订单行号已变化，请刷新后重试，未执行删除' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const permitted = await canOperateOrder(originalOrder, currentUser, submitterId, submitterName, 'all', kvAdminCache);
    if (!permitted) {
      return new Response(JSON.stringify({ success: false, error: '无权删除他人订单' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const resp = await deleteRow(rowIndex, kvAdminCache);
    const result = await resp.json();

    if (result.responses) {
      clearOrderCaches();
      return new Response(JSON.stringify({ success: true, message: '订单删除成功' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }
    return new Response(JSON.stringify({ success: false, error: JSON.stringify(result) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleClearTempRow(request, env, kvAdminCache) {
  try {
    const data = await request.json();
    const rowIndex = parseInt(data.row_index, 10) || 0;

    if (rowIndex > 0) {
      await clearTempRow(rowIndex, kvAdminCache);

      const keysToRemove = [];
      for (const key of Object.keys(tempRowTracker)) {
        if (tempRowTracker[key].row_index === rowIndex) {
          keysToRemove.push(key);
        }
      }
      for (const key of keysToRemove) {
        delete tempRowTracker[key];
      }
    }

    return new Response(JSON.stringify({ success: true }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleCleanupUserTempRows(request, env, kvAdminCache) {
  try {
    const data = await request.json();
    const submitterId = data.submitter_id || '';

    if (!submitterId) {
      return new Response(JSON.stringify({ success: true, message: '无用户ID' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const now = Date.now();
    const expiredRows = [];
    const keysToRemove = [];

    for (const key of Object.keys(tempRowTracker)) {
      const info = tempRowTracker[key];
      if (info.submitter_id === submitterId) {
        if (now - info.timestamp * 1000 > TEMP_ROW_TIMEOUT * 1000) {
          expiredRows.push(info.row_index);
          keysToRemove.push(key);
        }
      }
    }

    for (const key of keysToRemove) {
      delete tempRowTracker[key];
    }

    for (const rowIndex of expiredRows) {
      try {
        await clearTempRow(rowIndex, kvAdminCache);
      } catch (e) {
        console.warn(`[cleanup] 清空临时行失败 row=${rowIndex}: ${e.message}`);
      }
    }

    return new Response(JSON.stringify({ success: true, cleared_rows: expiredRows }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

export {
  getNextEmptyRow, ensureSheetRows, writeOrderRow, deleteRow, clearTempRow,
  fetchAllOrdersRaw, getFilteredOrders, clearOrderCaches, parseOrderFromRow,
  handleGetOrders, handleGetOrder, handleCreateOrder, handleUpdateOrder,
  handleDeleteOrder, handleClearTempRow, handleCleanupUserTempRows
};
