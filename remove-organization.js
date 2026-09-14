/* 通訊錄與移除名單只保留於記憶體；不使用 Seat 或服務帳號申請表快取。 */
'use strict';

function directoryEmployeeIds(workbook) {
  const ids = new Set();
  let skipped = 0;
  let found = false;
  for (const sheet of workbook.worksheets) {
    let headerRow = 0, column = 0;
    sheet.eachRow((row, number) => {
      if (headerRow) return;
      row.eachCell((cell, col) => {
        if (cellText(cell).replace(/\s/g, '') === '員工編號') {
          if (column) throw new Error(`${sheet.name} 有重複的員工編號欄。`);
          column = col;
        }
      });
      if (column) headerRow = number;
    });
    if (!headerRow) continue;
    found = true;
    sheet.eachRow((row, number) => {
      if (number <= headerRow) return;
      try {
        const cell = row.getCell(column);
        if (!cellText(cell).trim()) {
          if (row.actualCellCount) skipped++;
          return;
        }
        const value = typeof cell.value === 'object' && cell.value ? cell.value.result : cell.value;
        if (typeof value === 'number' && (!Number.isSafeInteger(value) || value < 0)) {
          skipped++;
          return;
        }
        const id = employeeIdText(cell);
        if (!/^\d{7}$/.test(id)) { skipped++; return; }
        ids.add(id);
      } catch (_) { skipped++; }
    });
  }
  if (!found || !ids.size) throw new Error('找不到有效的「員工編號」欄或通訊錄為空。');
  return { ids, skipped };
}

function organizationRemovalCandidates(logins, ids) {
  return [...new Map(logins.map(login => [login.toLowerCase(), login])).values()]
    .flatMap(login => {
      const match = /^IISI-(\d{7})$/i.exec(login);
      if (match && ids.has(match[1])) return [];
      return [{ login, employeeId: match ? match[1] : '',
        reason: match ? '不在通訊錄中' : '帳號格式不符（須為 IISI-7 碼數字）', selected: false, result: '' }];
    });
}

async function removalRequest(token, path, method = 'GET') {
  const response = await fetch(`https://api.github.com${path}`, {
    method, cache: 'no-store', headers: {
      Authorization: `Bearer ${token}`, Accept: 'application/vnd.github+json',
      'X-GitHub-Api-Version': '2026-03-10'
    }
  });
  if (response.status !== (method === 'DELETE' ? 204 : 200)) {
    let message = `HTTP ${response.status}`;
    try { message += `：${(await response.json()).message || 'GitHub API 回應異常'}`; } catch (_) {}
    throw new Error(message);
  }
  return method === 'DELETE' ? null : response.json();
}

async function freshRemovalLogins(token) {
  const logins = new Set();
  // 成員 API 回傳陣列，沒有 total_seats；逐頁讀取至不足一頁為止。
  for (let page = 1; page <= 1000; page++) {
    const data = await removalRequest(token, `/orgs/IISI-Group/members?filter=all&role=all&per_page=100&page=${page}`);
    if (!Array.isArray(data) || data.length > 100) {
      throw new Error('組織成員 API 回傳格式異常，未產生移除名單。');
    }
    for (const member of data) {
      const login = member?.login;
      if (typeof login !== 'string' || !login || logins.has(login.toLowerCase())) {
        throw new Error('組織成員帳號缺漏或分頁重複，請重新比對。');
      }
      logins.add(login.toLowerCase());
    }
    if (data.length < 100) return [...logins];
  }
  throw new Error('組織成員分頁超過上限，未產生移除名單。');
}

async function requireRemovalAdmin(token) {
  const membership = await removalRequest(token, '/user/memberships/orgs/IISI-Group');
  if (membership.state !== 'active' || membership.role !== 'admin' || !membership.user?.login) {
    throw new Error('此功能僅限 IISI-Group 組織擁有者操作。');
  }
  return membership.user.login.toLowerCase();
}

// 保護名單必須即時、完整；任何讀取錯誤都不能視為沒有 Seat。
async function freshRemovalSeatLogins(token) {
  const logins = new Set();
  let total = null;
  for (let page = 1; page <= 1000; page++) {
    const data = await removalRequest(token, `/orgs/IISI-Group/copilot/billing/seats?per_page=100&page=${page}`);
    if (!Array.isArray(data.seats) || data.seats.length > 100 ||
        !Number.isSafeInteger(data.total_seats) || data.total_seats < 0) {
      throw new Error('Seat 資料格式異常，無法確認保護名單。');
    }
    if (total !== null && total !== data.total_seats) throw new Error('Seat 名單已異動，請重新比對。');
    total = data.total_seats;
    for (const seat of data.seats) {
      const login = seat.assignee?.login;
      if (typeof login !== 'string' || !login || logins.has(login.toLowerCase())) {
        throw new Error('Seat 帳號缺漏或重複，無法確認保護名單。');
      }
      // 即使 pending_cancellation_date 有值，仍屬於 Seat 保護對象。
      logins.add(login.toLowerCase());
    }
    if (logins.size === total) return logins;
    if (logins.size > total || data.seats.length < 100) throw new Error('Seat 名單不完整，無法確認保護名單。');
  }
  throw new Error('Seat 分頁超過上限，無法確認保護名單。');
}

function protectRemovalSeatRows(rows, seatLogins) {
  for (const row of rows) {
    row.hasSeat = seatLogins.has(row.login.toLowerCase());
    if (row.hasSeat) {
      row.blocked = '有 Copilot Seat，禁止移出';
      row.selected = false;
    }
  }
}

(function setupOrganizationRemoval() {
  const el = id => document.getElementById(id);
  const directoryZone = el('removeDirectoryZone');
  let ids = null, rows = [], comparedToken = '', busy = false;
  const status = text => { el('removeStatus').textContent = text; };
  function reset() {
    rows = []; comparedToken = ''; el('removeConfirmation').value = ''; render();
  }
  function controls() {
    const available = rows.filter(row => !row.blocked && !row.result);
    const selected = available.filter(row => row.selected).length;
    const valid = comparedToken && comparedToken === localStorage.getItem(LS_TOKEN_KEY);
    el('removeDirectory').disabled = busy;
    directoryZone.setAttribute('aria-disabled', String(busy));
    directoryZone.tabIndex = busy ? -1 : 0;
    el('removeCompare').disabled = busy || !ids;
    el('removeSelectAll').disabled = busy || !valid || !available.length;
    el('removeSelectAll').checked = selected > 0 && selected === available.length;
    el('removeSelectAll').indeterminate = selected > 0 && selected < available.length;
    el('removeConfirmation').disabled = busy || !valid || !selected;
    el('removeExecute').disabled = busy || !valid || !selected || el('removeConfirmation').value !== 'REMOVE';
    el('removeExecute').textContent = `確認移出 IISI-Group（${selected} 人）`;
  }
  function render() {
    el('removeRows').replaceChildren();
    for (const row of rows) {
      const tr = document.createElement('tr');
      const td = document.createElement('td');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox'; checkbox.checked = row.selected;
      checkbox.disabled = busy || !!row.blocked || !!row.result;
      checkbox.setAttribute('aria-label', `選取 ${row.login}`);
      checkbox.addEventListener('change', () => {
        row.selected = checkbox.checked; el('removeConfirmation').value = ''; controls();
      });
      td.append(checkbox); tr.append(td);
      for (const text of [row.login, row.employeeId || '—', row.reason, row.hasSeat ? '有 Seat（禁止移出）' : '無 Seat', row.blocked || row.result || '待確認']) {
        const cell = document.createElement('td'); cell.textContent = text; tr.append(cell);
      }
      el('removeRows').append(tr);
    }
    controls();
  }
  async function loadDirectory(file) {
    if (busy || !file) return;
    ids = null; reset();
    directoryZone.classList.remove('loaded');
    el('removeDirectoryStatus').textContent = '解析中…';
    el('removeDirectoryStatus').className = 'zone-status';
    busy = true; controls();
    try {
      if (!/\.xlsx$/i.test(file.name)) throw new Error('請上傳 XLSX 通訊錄。');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
      const parsed = directoryEmployeeIds(workbook);
      ids = parsed.ids;
      el('removeDirectoryStatus').textContent = `✓ ${file.name}（${ids.size} 位不重複員工，略過 ${parsed.skipped} 筆無效資料）`;
      el('removeDirectoryStatus').className = 'zone-status ok';
      directoryZone.classList.add('loaded');
      status('通訊錄已載入，可取得 GitHub 帳號並比對。');
    } catch (error) {
      el('removeDirectoryStatus').textContent = `上傳失敗：${error.message}`;
      el('removeDirectoryStatus').className = 'zone-status error';
      status('通訊錄無效，無法比對或移除。');
    } finally { busy = false; el('removeDirectory').value = ''; controls(); }
  }
  el('removeDirectory').addEventListener('change', event => loadDirectory(event.target.files[0]));
  directoryZone.addEventListener('click', event => {
    if (!busy && event.target !== el('removeDirectory')) el('removeDirectory').click();
  });
  directoryZone.addEventListener('keydown', event => {
    if (event.key === 'Enter' || event.key === ' ') {
      event.preventDefault();
      if (!busy) el('removeDirectory').click();
    }
  });
  directoryZone.addEventListener('dragover', event => {
    event.preventDefault();
    if (!busy) directoryZone.classList.add('drag-over');
  });
  directoryZone.addEventListener('dragleave', event => {
    if (!directoryZone.contains(event.relatedTarget)) directoryZone.classList.remove('drag-over');
  });
  directoryZone.addEventListener('drop', event => {
    event.preventDefault();
    directoryZone.classList.remove('drag-over');
    return loadDirectory(event.dataTransfer.files[0]);
  });
  el('removeConfirmation').addEventListener('input', controls);
  el('removeSelectAll').addEventListener('change', event => {
    rows.forEach(row => { if (!row.blocked && !row.result) row.selected = event.target.checked; });
    el('removeConfirmation').value = ''; render();
  });
  el('removeCompare').addEventListener('click', async () => {
    if (busy || !ids) return;
    reset(); busy = true; controls(); status('正在確認管理員身分並取得所有組織成員…');
    const token = localStorage.getItem(LS_TOKEN_KEY);
    try {
      if (!token) throw new Error('請先設定 GitHub Token。');
      const self = await requireRemovalAdmin(token);
      const logins = await freshRemovalLogins(token);
      status('正在取得完整 Copilot Seat 保護名單…');
      const seatLogins = await freshRemovalSeatLogins(token);
      if (token !== localStorage.getItem(LS_TOKEN_KEY)) throw new Error('Token 已變更，請重新比對。');
      rows = organizationRemovalCandidates(logins, ids);
      rows.forEach(row => { if (row.login.toLowerCase() === self) row.blocked = '目前操作帳號，禁止移除自己'; });
      protectRemovalSeatRows(rows, seatLogins);
      comparedToken = token;
      status(`已比對 ${logins.length} 位組織成員／${ids.size} 位員工；${rows.length} 位待檢視，其中 ${rows.filter(row => row.hasSeat).length} 位有 Seat，禁止移出。請勾選要移出的人員。`);
    } catch (error) { reset(); status(`比對失敗：${error.message}`); }
    finally { busy = false; render(); }
  });
  el('removeExecute').addEventListener('click', async () => {
    const token = localStorage.getItem(LS_TOKEN_KEY);
    const selected = rows.filter(row => row.selected && !row.blocked && !row.result);
    if (busy || !ids || !selected.length ||
        !comparedToken || token !== comparedToken || el('removeConfirmation').value !== 'REMOVE') return;
    busy = true; el('removeConfirmation').value = ''; render();
    let succeeded = 0;
    let protectedCount = 0;
    try {
      const self = await requireRemovalAdmin(token);
      status('執行前重新確認 Copilot Seat 保護名單…');
      const seatLogins = await freshRemovalSeatLogins(token);
      protectRemovalSeatRows(rows, seatLogins);
      render();
      for (const row of selected) {
        if (token !== localStorage.getItem(LS_TOKEN_KEY)) throw new Error('Token 已變更，停止後續移除。');
        if (row.login.toLowerCase() === self) throw new Error('禁止移除目前操作帳號。');
        if (row.hasSeat) { protectedCount++; continue; }
        status(`正在移除 ${row.login}（${succeeded + 1}/${selected.length}）…`);
        try {
          await removalRequest(token, `/orgs/IISI-Group/members/${encodeURIComponent(row.login)}`, 'DELETE');
          row.result = '已完成移除 API（204）'; row.selected = false; succeeded++;
        } catch (error) {
          row.result = `失敗：${error.message}`;
          throw error; // 權限、限流或網路錯誤時停止，避免繼續批次操作。
        }
        render();
      }
      status(`執行完成：成功 ${succeeded} 人，略過 ${protectedCount} 位有 Seat 的帳號。請重新比對以確認最新名單。`);
    } catch (error) {
      status(`已停止：成功 ${succeeded} 人；${error.message}。其餘未執行，請重新比對後再確認。`);
    } finally {
      comparedToken = ''; busy = false;
      rows.forEach(row => { row.selected = false; });
      localStorage.removeItem(LS_SEAT_CACHE_KEY);
      render();
    }
  });
  function tokenChanged() {
    if (!busy) { reset(); status('Token 已更新，請重新比對。'); }
  }
  el('tokenOk').addEventListener('click', tokenChanged);
  el('tokenDelete').addEventListener('click', tokenChanged);
  window.addEventListener('storage', event => { if (event.key === LS_TOKEN_KEY || event.key === null) tokenChanged(); });
  controls();
})();
