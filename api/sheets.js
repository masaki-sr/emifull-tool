// ============================================================
// Google Sheets API 直結版（GASを経由しない）
// ------------------------------------------------------------
// 対応済み（このファイルだけで完結）:
//   ・GET  ?sheet=シート名           → シートの全行をJSONで取得
//   ・POST { sheet, ...fields }     → 新規行を追加（IDが空ならID列があれば自動採番）
//   ・POST { sheet, action:'updateRow', updateKey, [updateKey]:値, ...fields }
//                                    → 該当行を更新（他の列は保持）
//   ・POST { sheet, action:'deleteRow', deleteKey, [deleteKey]:値 }
//                                    → 該当行を削除
//
// 未対応のもの（completeTodo, getConsultTodos, getBusySeasonRecords など）は、
// 今まで通りGAS側（LEGACY_GAS_URL）へそのまま転送する（段階移行のための保険）。
// ============================================================

import { google } from 'googleapis';

export const config = { runtime: 'nodejs' };

const SPREADSHEET_ID   = process.env.GOOGLE_SPREADSHEET_ID;
const CLIENT_EMAIL     = process.env.GOOGLE_CLIENT_EMAIL;
const PRIVATE_KEY_RAW  = process.env.GOOGLE_PRIVATE_KEY || '';
const PRIVATE_KEY      = PRIVATE_KEY_RAW.includes('\\n') ? PRIVATE_KEY_RAW.replace(/\\n/g, '\n') : PRIVATE_KEY_RAW;

// 未対応アクションの転送先（これまで使っていたGAS ウェブアプリURL）
const LEGACY_GAS_URL = 'https://script.google.com/macros/s/AKfycbzm3eLCAQxUdZE8vpW7N2mOL3IiNT0c_KLzDs8YIBjU446SWocEoQ1SxUZVjtYV0JG5Rw/exec';

const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type',
};

let _sheetsClient = null;
function getSheetsClient() {
  if (_sheetsClient) return _sheetsClient;
  if (!CLIENT_EMAIL || !PRIVATE_KEY) {
    throw new Error('GOOGLE_CLIENT_EMAIL または GOOGLE_PRIVATE_KEY が設定されていません（Vercelの環境変数を確認してください）');
  }
  const auth = new google.auth.JWT(
    CLIENT_EMAIL,
    null,
    PRIVATE_KEY,
    ['https://www.googleapis.com/auth/spreadsheets']
  );
  _sheetsClient = google.sheets({ version: 'v4', auth });
  return _sheetsClient;
}

function columnLetter(oneBasedIndex) {
  let n = oneBasedIndex;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// シートの全データ（見出し行含む）を取得
async function getAllValues(sheets, sheetName) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: sheetName,
  });
  return res.data.values || [];
}

// ── GET: シートの全行をオブジェクト配列で返す ──
async function handleGetSheetData(sheets, sheetName) {
  const rows = await getAllValues(sheets, sheetName);
  if (rows.length === 0) return [];
  const headers = rows[0];
  return rows.slice(1)
    .filter(row => row.some(cell => String(cell || '').trim() !== '')) // 完全な空行は除外
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
}

// ── POST（action指定なし）: 新規行を追加 ──
async function handleAppendRow(sheets, sheetName, data) {
  const rows = await getAllValues(sheets, sheetName);
  const headers = rows[0] || [];
  if (headers.length === 0) {
    return { error: `シート「${sheetName}」の見出し行が見つかりません` };
  }

  // ID列があり、かつIDが指定されていなければ自動採番（既存の最大値+1）
  const idCol = headers.indexOf('ID');
  if (idCol !== -1 && !data['ID']) {
    let maxId = 0;
    for (let i = 1; i < rows.length; i++) {
      const v = parseInt(rows[i][idCol], 10);
      if (!isNaN(v) && v > maxId) maxId = v;
    }
    data['ID'] = String(maxId + 1);
  }

  const newRow = headers.map(h => (data[h] !== undefined ? data[h] : ''));

  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A1`,
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: { values: [newRow] },
  });

  return { success: true, savedData: newRow, headers: headers };
}

// 指定キー列・キー値に一致する行番号（1始まり、見出し込み）を探す
function findRowNumber(rows, headers, keyCol, keyVal) {
  const idx = headers.indexOf(keyCol);
  if (idx === -1) return { error: `${keyCol}列が見つかりません` };
  const targetId = String(keyVal || '');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idx]) === targetId) {
      return { rowNum: i + 1 };
    }
  }
  return { error: `該当ID(${targetId})が見つかりません` };
}

// ── POST action=updateRow: 既存行を更新（未指定の列は既存値を保持） ──
async function handleUpdateRow(sheets, sheetName, data) {
  const rows = await getAllValues(sheets, sheetName);
  const headers = rows[0] || [];
  const keyCol = data.updateKey || 'ID';
  const keyVal = data[keyCol] !== undefined ? data[keyCol] : data.id;

  const found = findRowNumber(rows, headers, keyCol, keyVal);
  if (found.error) return { error: found.error + `（シート: ${sheetName}）` };

  const existingRow = rows[found.rowNum - 1] || [];
  const newRow = headers.map((h, i) => (data[h] !== undefined ? data[h] : (existingRow[i] !== undefined ? existingRow[i] : '')));

  await sheets.spreadsheets.values.update({
    spreadsheetId: SPREADSHEET_ID,
    range: `${sheetName}!A${found.rowNum}:${columnLetter(headers.length)}${found.rowNum}`,
    valueInputOption: 'USER_ENTERED',
    requestBody: { values: [newRow] },
  });

  return { success: true, updatedRow: found.rowNum };
}

// ── POST action=deleteRow: 行を完全に削除 ──
async function handleDeleteRow(sheets, sheetName, data) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetMeta = (meta.data.sheets || []).find(s => s.properties.title === sheetName);
  if (!sheetMeta) return { error: `シート「${sheetName}」が見つかりません` };
  const sheetId = sheetMeta.properties.sheetId;

  const rows = await getAllValues(sheets, sheetName);
  const headers = rows[0] || [];
  const keyCol = data.deleteKey || 'client_id';
  const keyVal = data[keyCol] !== undefined ? data[keyCol] : data.id;

  const found = findRowNumber(rows, headers, keyCol, keyVal);
  if (found.error) return { error: found.error + `（シート: ${sheetName}）` };

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [{
        deleteDimension: {
          range: { sheetId, dimension: 'ROWS', startIndex: found.rowNum - 1, endIndex: found.rowNum },
        },
      }],
    },
  });

  return { success: true, deletedRow: found.rowNum };
}

// 未対応のaction・パラメータをこれまでのGASへそのまま転送する（段階移行の保険）
async function forwardToLegacyGas(req) {
  const url = new URL(req.url, 'https://dummy.local');
  const targetUrl = LEGACY_GAS_URL + (url.search || '');

  const opts = { method: req.method };
  if (req.method === 'POST') {
    opts.headers = { 'Content-Type': 'text/plain;charset=utf-8' };
    opts.body = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});
  }
  const r = await fetch(targetUrl, opts);
  const text = await r.text();
  return { status: r.status, text };
}

export default async function handler(req, res) {
  if (req.method === 'OPTIONS') {
    Object.entries(CORS_HEADERS).forEach(([k, v]) => res.setHeader(k, v));
    res.status(200).end();
    return;
  }
  Object.entries(CORS_HEADERS).forEach(([k, v]) => res.setHeader(k, v));

  try {
    const sheetParam  = req.query.sheet;
    const actionParam = req.query.action;

    // ---- GET: 汎用シート取得（?sheet=名前）のみ、ここで完結させる ----
    if (req.method === 'GET' && sheetParam && !actionParam) {
      const sheets = getSheetsClient();
      const data = await handleGetSheetData(sheets, sheetParam);
      res.status(200).json(data);
      return;
    }

    // ---- GET: その他（?action=xxx など特殊なもの）はGASへ転送 ----
    if (req.method === 'GET') {
      const { status, text } = await forwardToLegacyGas(req);
      res.status(status);
      res.setHeader('Content-Type', 'application/json; charset=utf-8');
      res.send(text);
      return;
    }

    // ---- POST ----
    if (req.method === 'POST') {
      let data = req.body;
      if (typeof data === 'string') {
        try { data = JSON.parse(data); } catch (e) {
          res.status(400).json({ error: 'POSTボディがJSONとして解釈できません' });
          return;
        }
      }
      if (!data || typeof data !== 'object') {
        res.status(400).json({ error: 'POSTボディが空です' });
        return;
      }

      const sheetName = data.sheet;
      const action = data.action;

      // 未対応のaction（completeTodo, completeConsultTodo, saveWorkAvailability 等）はGASへ転送
      const SUPPORTED_ACTIONS = ['updateRow', 'deleteRow', undefined];
      if (!sheetName || !SUPPORTED_ACTIONS.includes(action)) {
        const { status, text } = await forwardToLegacyGas(req);
        res.status(status);
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.send(text);
        return;
      }

      const sheets = getSheetsClient();
      let result;
      if (action === 'updateRow') {
        result = await handleUpdateRow(sheets, sheetName, data);
      } else if (action === 'deleteRow') {
        result = await handleDeleteRow(sheets, sheetName, data);
      } else {
        result = await handleAppendRow(sheets, sheetName, data);
      }

      const status = result && result.error ? 400 : 200;
      res.status(status).json(result);
      return;
    }

    res.status(405).json({ error: 'サポートされていないメソッドです' });
  } catch (e) {
    console.error('[api/sheets] エラー:', e);
    res.status(500).json({ error: 'Sheets API接続失敗: ' + (e.message || String(e)) });
  }
}
