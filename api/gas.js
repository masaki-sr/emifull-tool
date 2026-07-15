// Vercel Edge Runtime（タイムアウト30秒、Hobbyプランでも有効）
export const config = { runtime: 'edge' };

const GAS_BASE = 'https://script.google.com/macros/s/AKfycbzm3eLCAQxUdZE8vpW7N2mOL3IiNT0c_KLzDs8YIBjU446SWocEoQ1SxUZVjtYV0JG5Rw/exec';

/**
 * GASへのfetch（リダイレクトを手動追跡）
 *
 * GAS Web Appの動作：
 *   1. POST受信 → doPost(e)実行 → スプレッドシート書き込み完了
 *   2. GASが302を返す（script.googleusercontent.com/macros/echo?... への応答配信用リダイレクト）
 *   3. リダイレクト先はGASスクリプトを再実行しない（事前生成済みJSONを返すだけ）
 *   → doPostは302を返す前に完了しているため、302→GETへの変換は保存済みの応答取得のみ
 */
async function fetchGas(url, method, body, contentType, hops, logs) {
  const opts = { method, redirect: 'manual' };
  if (body !== null && body !== undefined) {
    opts.body = body;
    opts.headers = { 'Content-Type': contentType || 'text/plain;charset=utf-8' };
  }

  logs.push(`[${hops}] ${method} ${url.replace(GAS_BASE, 'GAS_BASE').slice(0, 80)}`);
  const r = await fetch(url, opts);
  logs.push(`[${hops}] status=${r.status}`);

  const isRedirect = r.status >= 301 && r.status <= 308;
  if (isRedirect && hops < 5) {
    const loc = r.headers.get('location');
    if (!loc) {
      logs.push(`[${hops}] redirect但しLocationなし → そのまま返す`);
      return r;
    }
    // 307/308はPOSTを維持、301/302/303はGETへ変換（RFC標準）
    // ※GASは302でdoPost応答を返す。302後GETになるのは「応答取得」であり保存処理は既に完了済み
    const keepPost = r.status === 307 || r.status === 308;
    const nextMethod = keepPost ? method : 'GET';
    const nextBody   = nextMethod === 'POST' ? body : null;
    const nextCt     = nextMethod === 'POST' ? contentType : null;
    logs.push(`[${hops}] redirect ${r.status}: ${method} → ${nextMethod} (GASはdoPost完了後に302を返すため保存は完了済み)`);
    return fetchGas(loc, nextMethod, nextBody, nextCt, hops + 1, logs);
  }
  return r;
}

export default async function handler(req) {
  const corsHeaders = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  if (req.method === 'OPTIONS') {
    return new Response(null, { status: 200, headers: corsHeaders });
  }

  const url = new URL(req.url);
  const params = url.searchParams.toString();
  const targetUrl = params ? `${GAS_BASE}?${params}` : GAS_BASE;
  const logs = [];

  logs.push(`[api/gas] 受信: ${req.method} params=${params || '(none)'}`);

  try {
    let body = null;
    let contentType = null;

    if (req.method === 'POST') {
      body = await req.text();
      contentType = req.headers.get('content-type') || 'text/plain;charset=utf-8';
      logs.push(`[api/gas] POST bodyLen=${body.length} ct=${contentType}`);
      logs.push(`[api/gas] POST body先頭: ${body.slice(0, 150)}`);

      if (body.length === 0) {
        console.error('[api/gas] ❌ POSTボディが空です。フロント側の送信を確認してください。');
        return new Response(JSON.stringify({ error: 'POSTボディが空です' }), {
          status: 400,
          headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
        });
      }
    }

    logs.push(`[api/gas] → GASへ転送: ${req.method} ${targetUrl.replace(GAS_BASE, 'GAS_BASE')}`);
    const gasRes = await fetchGas(targetUrl, req.method, body, contentType, 0, logs);

    const text = await gasRes.text();
    const gasCt = gasRes.headers.get('content-type') || '';
    const trimmed = (text || '').trim();
    logs.push(`[api/gas] 最終 status=${gasRes.status} ct=${gasCt} resLen=${text.length}`);
    logs.push(`[api/gas] 最終 response先頭: ${trimmed.slice(0, 150)}`);

    // 本文がJSONらしいか（content-type or 先頭文字で判定）
    const looksJson =
      gasCt.toLowerCase().includes('application/json') ||
      trimmed.startsWith('{') ||
      trimmed.startsWith('[');

    // GASが混雑・タイムアウトでHTMLエラーページ（"An error occurred" 等）や
    // 非2xxを返した場合、それをJSONと偽って返さない（フロントの JSON.parse を壊さない）。
    // 適切なHTTPステータス（504/502等）で返し、フロント側でリトライ・警告できるようにする。
    if (!gasRes.ok || !looksJson) {
      const status = (gasRes.status && gasRes.status >= 400) ? gasRes.status : 502;
      logs.push(`[api/gas] ⚠ 非JSON/エラー応答 → status=${status} で error JSON を返す`);
      console.error(logs.join('\n'));
      return new Response(JSON.stringify({
        error: 'GASが正しいJSONを返しませんでした（混雑またはタイムアウトの可能性）',
        gasStatus: gasRes.status,
        detail: trimmed.slice(0, 200),
      }), {
        status,
        headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
      });
    }

    // successフラグをログに出す
    try {
      const parsed = JSON.parse(trimmed);
      if (parsed && typeof parsed.success !== 'undefined') {
        logs.push(`[api/gas] success=${parsed.success}${parsed.error ? ' error=' + parsed.error : ''}${parsed.sheetName ? ' sheetName=' + parsed.sheetName : ''}${parsed.writtenRow ? ' writtenRow=' + parsed.writtenRow : ''}`);
      }
    } catch (_) { /* レスポンスがJSONでない場合は無視 */ }

    // まとめてログ出力
    console.log(logs.join('\n'));

    return new Response(trimmed || '{}', {
      status: 200,
      headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
    });
  } catch (e) {
    logs.push(`[api/gas] ❌ ERROR: ${e.name} ${e.message}`);
    console.error(logs.join('\n'));
    return new Response(JSON.stringify({ error: 'GAS接続失敗: ' + (e.message || String(e)) }), {
      status: 502,
      headers: { ...corsHeaders, 'Content-Type': 'application/json; charset=utf-8' },
    });
  }
}
