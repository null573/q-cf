import {
  USER_FILE_ID, USER_SHEET_ID, USER_CACHE_TTL, ACCESS_PASSWORD,
  usersCache, parseCellValue, normalizeUserKey
} from './config.js';
import { readSheetRange, batchUpdate } from './tencent-api.js';

async function readUsers(kvAdminCache) {
  const now = Date.now();
  if (usersCache.data !== null && (now - usersCache.timestamp) < USER_CACHE_TTL * 1000) {
    return usersCache.data;
  }

  const gridData = await readSheetRange(USER_SHEET_ID, 'A2:F200', USER_FILE_ID, kvAdminCache);
  const rows = gridData.rows || [];
  const users = [];

  for (const row of rows) {
    const values = row.values || [];
    const rowData = values.map(v => parseCellValue(v.cellValue));
    if (rowData.length >= 3 && rowData[0] && rowData[1]) {
      const role = rowData.length >= 4 ? (rowData[3] || '').trim() : '';
      const department = rowData.length >= 5 ? (rowData[4] || '').trim() : '';
      const permissionText = rowData.length >= 6 ? (rowData[5] || '').trim() : '';
      const isAdmin = role === '管理员' || permissionText === '能操作所有数据';
      const isManager = role === '经理' || permissionText === '能操作本部门所有数据';
      const accessLevel = isAdmin ? 'admin' : (isManager ? 'department' : 'self');
      users.push({
        name: rowData[0],
        employee_id: rowData[1],
        password: rowData[2],
        is_admin: isAdmin,
        is_manager: isManager,
        role: role,
        department: department,
        access_level: accessLevel,
        permission: permissionText
      });
    }
  }

  usersCache.data = users;
  usersCache.timestamp = now;
  return users;
}

async function getUserById(employeeId, kvAdminCache) {
  const currentId = normalizeUserKey(employeeId);
  if (!currentId) return null;
  const users = await readUsers(kvAdminCache);
  for (const user of users) {
    if (normalizeUserKey(user.employee_id) === currentId) {
      return user;
    }
  }
  return null;
}

async function getUserByName(name, kvAdminCache) {
  const currentName = String(name || '').trim();
  if (!currentName) return null;
  const users = await readUsers(kvAdminCache);
  for (const user of users) {
    if (String(user.name || '').trim() === currentName) {
      return user;
    }
  }
  return null;
}

async function isUserAdmin(employeeId, kvAdminCache) {
  const user = await getUserById(employeeId, kvAdminCache);
  return !!(user && user.access_level === 'admin');
}

async function resolveSubmitterName(submitterId, submitterName, kvAdminCache) {
  const name = String(submitterName || '').trim();
  if (name && name !== '用户') return name;

  const currentId = normalizeUserKey(submitterId);
  if (!currentId) return name;

  const users = await readUsers(kvAdminCache);
  for (const user of users) {
    if (normalizeUserKey(user.employee_id) === currentId) {
      return String(user.name || '').trim();
    }
  }
  return name;
}

async function getOrderSubmitterUser(order, kvAdminCache) {
  return await getUserById(order.submitter_id, kvAdminCache)
      || await getUserByName(order.submitter, kvAdminCache);
}

function isSameSubmitter(order, submitterId, submitterName) {
  const currentId = normalizeUserKey(submitterId);
  const rowId = normalizeUserKey(order.submitter_id || '');
  const currentName = String(submitterName || '').trim();
  const rowName = String(order.submitter || '').trim();

  if (currentId && rowId && currentId === rowId) return true;
  if (currentName && rowName && currentName === rowName) return true;
  return false;
}

function orderMatchesExpected(order, expected) {
  if (!expected) return true;
  const keys = ['model', 'customer', 'submitter_id', 'submit_time'];
  for (const key of keys) {
    const expectedValue = String(expected[key] || '').trim();
    if (expectedValue && String(order[key] || '').trim() !== expectedValue) {
      return false;
    }
  }
  return true;
}

async function canOperateOrder(order, currentUser, submitterId, submitterName, viewMode, kvAdminCache) {
  const accessLevel = (currentUser || {}).access_level || 'self';
  if (viewMode === 'mine') {
    return isSameSubmitter(order, submitterId, submitterName);
  }
  if (accessLevel === 'admin') return true;
  if (accessLevel === 'department' || accessLevel === 'self') {
    const currentDept = String((currentUser || {}).department || '').trim();
    if (!currentDept) return true;
    const orderUser = await getOrderSubmitterUser(order, kvAdminCache);
    const orderDept = String((orderUser || {}).department || '').trim();
    return !!(currentDept && orderDept && currentDept === orderDept);
  }
  return isSameSubmitter(order, submitterId, submitterName);
}

function normalizeViewMode(currentUser, requestedViewMode) {
  if (requestedViewMode === 'all') return 'all';
  return 'mine';
}

async function handleAuthCheck(request, env, kvAdminCache) {
  const password = request.headers.get('X-Access-Password') || '';
  const authorized = password === ACCESS_PASSWORD;
  const status = authorized ? 200 : 401;
  return new Response(JSON.stringify({ authorized }), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}

async function handleAuthLogin(request, env, kvAdminCache) {
  try {
    const data = await request.json();
    const employeeId = data.employee_id || '';
    const password = data.password || '';
    const users = await readUsers(kvAdminCache);

    for (const user of users) {
      if (user.employee_id === employeeId) {
        if (user.password === password) {
          return new Response(JSON.stringify({
            success: true,
            user: { name: user.name, employee_id: user.employee_id },
            access_password: ACCESS_PASSWORD
          }), { headers: { 'Content-Type': 'application/json' } });
        }
        return new Response(JSON.stringify({ success: false, error: '密码错误' }), {
          headers: { 'Content-Type': 'application/json' }
        });
      }
    }
    return new Response(JSON.stringify({ success: false, error: '员工号不存在' }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleAuthUsers(request, env, kvAdminCache) {
  try {
    const users = await readUsers(kvAdminCache);
    const userList = users.map(u => ({ name: u.name, employee_id: u.employee_id }));
    return new Response(JSON.stringify({ success: true, users: userList }), {
      headers: { 'Content-Type': 'application/json' }
    });
  } catch (e) {
    return new Response(JSON.stringify({ success: false, error: String(e) }), {
      headers: { 'Content-Type': 'application/json' }
    });
  }
}

async function handleUpdatePassword(request, env, kvAdminCache) {
  try {
    const data = await request.json();
    const employeeId = data.employee_id || '';
    const oldPassword = data.old_password || '';
    const newPassword = data.new_password || '';

    if (!employeeId || !oldPassword || !newPassword) {
      return new Response(JSON.stringify({ success: false, error: '参数不完整' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }
    if (newPassword.length < 6) {
      return new Response(JSON.stringify({ success: false, error: '新密码至少6位' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }
    const hasLetter = /[a-zA-Z]/.test(newPassword);
    const hasDigit = /[0-9]/.test(newPassword);
    if (!(hasLetter && hasDigit)) {
      return new Response(JSON.stringify({ success: false, error: '密码必须同时包含字母和数字' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const gridData = await readSheetRange(USER_SHEET_ID, 'A2:C200', USER_FILE_ID, kvAdminCache);
    const rows = gridData.rows || [];
    let targetRow = null;

    for (let i = 0; i < rows.length; i++) {
      const values = rows[i].values || [];
      const rowData = values.map(v => parseCellValue(v.cellValue));
      if (rowData.length >= 2 && normalizeUserKey(rowData[1]) === normalizeUserKey(employeeId)) {
        if (rowData.length >= 3 && rowData[2] === oldPassword) {
          targetRow = i + 2;
        } else {
          return new Response(JSON.stringify({ success: false, error: '旧密码错误' }), {
            headers: { 'Content-Type': 'application/json' }
          });
        }
        break;
      }
    }

    if (targetRow === null) {
      return new Response(JSON.stringify({ success: false, error: '员工号不存在' }), {
        headers: { 'Content-Type': 'application/json' }
      });
    }

    const resp = await batchUpdate([{
      updateRangeRequest: {
        sheetId: USER_SHEET_ID,
        gridData: {
          startRow: targetRow - 1,
          startColumn: 2,
          rows: [{ values: [{ cellValue: { text: newPassword } }] }]
        }
      }
    }], USER_FILE_ID, kvAdminCache);

    const result = await resp.json();
    if (result.responses) {
      usersCache.data = null;
      usersCache.timestamp = 0;
      return new Response(JSON.stringify({ success: true, message: '密码修改成功' }), {
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

export {
  readUsers, getUserById, getUserByName, isUserAdmin,
  resolveSubmitterName, getOrderSubmitterUser, canOperateOrder,
  normalizeViewMode, isSameSubmitter, orderMatchesExpected,
  handleAuthCheck, handleAuthLogin, handleAuthUsers, handleUpdatePassword
};
