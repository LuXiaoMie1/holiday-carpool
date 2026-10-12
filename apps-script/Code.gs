// 連假出遊乘車調度 — Google 試算表後端
// 綁在一張試算表上（擴充功能 → Apps Script），部署成「網頁應用程式」。
// 資料存在工作表 store：col | id | json | updatedAt
// GET  → 回傳整份狀態 {entries, ideas, meta}
// POST → body 為 JSON：{action:'set'|'delete', col, id, data}

const SHEET_NAME = 'store';

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['col', 'id', 'json', 'updatedAt']);
  }
  return sh;
}

function readAll_() {
  const rows = sheet_().getDataRange().getValues();
  const out = { entries: {}, ideas: {}, meta: {} };
  for (let i = 1; i < rows.length; i++) {
    const col = String(rows[i][0] || ''), id = String(rows[i][1] || ''), json = rows[i][2];
    if (!col || !id) continue;
    try {
      const d = JSON.parse(json);
      if (col === 'meta') out.meta = d;
      else { if (!out[col]) out[col] = {}; out[col][id] = d; }
    } catch (e) { /* 壞掉的列直接跳過 */ }
  }
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return json_(readAll_());
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const col = String(body.col || ''), id = String(body.id || '');
    if (!/^(entries|ideas|meta)$/.test(col) || !/^[A-Za-z0-9_\-.~:@+]{1,200}$/.test(id)) {
      return json_({ ok: false, error: 'bad path' });
    }
    const sh = sheet_();
    const rows = sh.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === col && String(rows[i][1]) === id) { rowIdx = i + 1; break; }
    }
    if (body.action === 'delete') {
      if (rowIdx > 0) sh.deleteRow(rowIdx);
    } else if (body.action === 'set') {
      const json = JSON.stringify(body.data || {});
      if (json.length > 20000) return json_({ ok: false, error: 'too big' });
      if (rowIdx > 0) sh.getRange(rowIdx, 3, 1, 2).setValues([[json, new Date()]]);
      else sh.appendRow([col, id, json, new Date()]);
    } else {
      return json_({ ok: false, error: 'bad action' });
    }
    return json_({ ok: true, state: readAll_() });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    lock.releaseLock();
  }
}
