/**
 * =====================================================================
 * 【正本】standards/gas/Gemini.gs — Gemini API 呼び出しの共通化
 * =====================================================================
 *
 * 各リポジトリへは Gemini.gs としてそのままコピーする（中身は変えない）。
 * リポジトリごとに違うもの（APIキーの取り出し方・モデル名・プロンプト）は
 * 呼び出し側から引数で渡す。この中には書かない。
 *
 * ■ なぜ共通化したか
 *   2026-08 の調査で、Gemini の呼び出しは 7 リポジトリ 12 か所にあり、
 *   そのうち **再試行を持っていたのは 2 リポジトリだけ**だった。
 *   Gemini は混み合うと 429 / 503 を返す。再試行が無い側では、
 *   一括処理の途中で数件が黙って落ち、先生の画面には
 *   「何件か返ってこなかった」としか出ない。取りこぼしの原因が
 *   分からないまま「AIは当てにならない」という評価になる。
 *
 * ■ API キーは URL クエリに入れない
 *   ?key=... はアクセスログやプロキシに残る。x-goog-api-key ヘッダで渡す。
 *
 * ■ 使い方
 *   const text = GigaGemini.call({ apiKey: key, prompt: '…' });
 *   const obj  = GigaGemini.callJson({ apiKey: key, prompt: '…' });
 *   const res  = GigaGemini.callAll([{ apiKey: key, prompt: 'a' }, …]);
 *
 *   既存の関数名（callGeminiApi_ など）はそのまま残し、その中身を
 *   GigaGemini への委譲に置き換える。呼び出し側は書き換えない。
 * =====================================================================
 */
var GigaGemini = (function () {
  var DEFAULTS = {
    model: 'gemini-2.0-flash',
    maxAttempts: 3,        // 初回 + 最大2回の再試行
    baseDelayMs: 1000,     // 1s → 2s → 4s
    apiVersion: 'v1beta',
    chunkSize: 8,          // callAll が一度に投げる本数
    maxRetryAfterMs: 16000, // Retry-After が長すぎるときの頭打ち（GAS の6分上限を守るため）
  };

  // 一時的なエラーだけ再試行する。400（プロンプト不正）や 403（キー不正）は
  // 何度投げても同じなので、待たせるだけ無駄。
  var RETRIABLE = [429, 500, 502, 503, 504];

  /**
   * 次の再試行までの待ち時間。
   * Gemini が Retry-After を返してきたらそれに従う（こちらの決め打ちより正確なため）。
   * ただし長すぎる指示は頭打ちにする。GAS の実行は6分で切られるので、
   * 素直に待つと処理そのものが落ちて、待った意味が無くなる。
   */
  function waitMsFor(attempt, headers, opts) {
    var base = (opts && opts.baseDelayMs) || DEFAULTS.baseDelayMs;
    var cap = (opts && opts.maxRetryAfterMs) || DEFAULTS.maxRetryAfterMs;
    var h = headers || {};
    var raw = h['Retry-After'] || h['retry-after'];
    var sec = parseInt(raw, 10);
    if (!isNaN(sec) && sec > 0) return Math.min(sec * 1000, cap);
    return base * Math.pow(2, attempt);
  }

  function endpoint(model, apiVersion) {
    return 'https://generativelanguage.googleapis.com/' + (apiVersion || DEFAULTS.apiVersion) +
      '/models/' + (model || DEFAULTS.model) + ':generateContent';
  }

  function buildBody(req) {
    var body = { contents: [{ parts: [{ text: String(req.prompt || '') }] }] };
    if (req.systemInstruction) {
      body.systemInstruction = { parts: [{ text: String(req.systemInstruction) }] };
    }
    if (req.generationConfig) body.generationConfig = req.generationConfig;
    if (req.parts) body.contents = [{ parts: req.parts }];   // 画像・PDF を混ぜるとき
    return body;
  }

  function buildOptions(req) {
    if (!req.apiKey) throw new Error('BAD_INPUT: Gemini APIキーが設定されていません。設定画面から保存してください');
    return {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': req.apiKey },
      payload: JSON.stringify(buildBody(req)),
      muteHttpExceptions: true,
    };
  }

  /** 応答 JSON から本文テキストを取り出す。取れなければ null。 */
  function extractText(json) {
    var c = json && json.candidates && json.candidates[0];
    var parts = c && c.content && c.content.parts;
    var text = parts && parts[0] && parts[0].text;
    return typeof text === 'string' ? text : null;
  }

  /** 応答コードと本文から、先生に見せる日本語のエラーを作る。 */
  function failure(code, body) {
    var detail = '';
    try {
      var j = JSON.parse(body);
      detail = (j.error && j.error.message) ? j.error.message : '';
    } catch (e) { /* JSON でないときは素の本文を使わない（長すぎるため） */ }
    if (code === 429) return new Error('AI_BUSY: AIが混み合っています。少し時間をおいて、もう一度お試しください');
    if (code === 400) return new Error('AI_ERROR: AIへの依頼の形が正しくありませんでした' + (detail ? '（' + detail + '）' : ''));
    if (code === 403) return new Error('AI_ERROR: Gemini APIキーが使えませんでした。設定画面でキーを確かめてください');
    return new Error('AI_ERROR: AIとの通信に失敗しました（HTTP ' + code + '）。しばらくしてからお試しください');
  }

  /**
   * 1件呼び、応答 JSON をそのまま返す。組み立て済みの payload を渡したいとき用。
   * 一時エラー（429/5xx）は指数バックオフで再試行する。
   *
   * @param {{apiKey: string, payload?: Object, url?: string, model?: string,
   *          apiVersion?: string, maxAttempts?: number, baseDelayMs?: number,
   *          maxRetryAfterMs?: number, log?: function(string), logLabel?: string}} req
   *        payload を渡さない場合は prompt などから組み立てる。
   * @return {Object} Gemini の応答 JSON
   */
  function callRaw(req) {
    if (!req.apiKey) throw new Error('BAD_INPUT: Gemini APIキーが設定されていません。設定画面から保存してください');
    var url = req.url || endpoint(req.model, req.apiVersion);
    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': req.apiKey },
      payload: JSON.stringify(req.payload || buildBody(req)),
      muteHttpExceptions: true,
    };
    var maxAttempts = req.maxAttempts || DEFAULTS.maxAttempts;
    var label = req.logLabel || 'Gemini';

    var lastCode = 0;
    var lastBody = '';
    for (var attempt = 0; attempt < maxAttempts; attempt++) {
      var res = UrlFetchApp.fetch(url, options);
      lastCode = res.getResponseCode();
      lastBody = res.getContentText();
      if (lastCode === 200) return JSON.parse(lastBody);
      if (RETRIABLE.indexOf(lastCode) === -1) break;
      if (attempt === maxAttempts - 1) break;   // 最後の失敗のあとは待たない
      var wait = waitMsFor(attempt, (res.getHeaders && res.getHeaders()) || {}, req);
      if (req.log) req.log(label + ': HTTP ' + lastCode + ' のため ' + wait + 'ms 後に再試行します（' + (attempt + 1) + '/' + (maxAttempts - 1) + '）');
      Utilities.sleep(wait);
    }
    throw failure(lastCode, lastBody);
  }

  /**
   * 1件呼んで本文テキストを返す。
   * @param {{apiKey: string, prompt: string, model?: string, systemInstruction?: string,
   *          generationConfig?: Object, parts?: Array, maxAttempts?: number,
   *          baseDelayMs?: number, apiVersion?: string}} req
   * @return {string} 生成された本文
   */
  function call(req) {
    var text = extractText(callRaw(req));
    // 200 でも本文が空のことがある（安全フィルタで止められたとき等）。
    // 空文字を「成功」として返すと、画面に空欄が保存されて原因が分からなくなる。
    if (text === null || text === '') {
      throw new Error('AI_EMPTY: AIから答えが返りませんでした。書いた内容を少し変えてお試しください');
    }
    return text.trim();
  }

  /**
   * JSON で答えさせて、コードフェンスを取り除いてパースする。
   * @return {Object}
   */
  function callJson(req) {
    var merged = {};
    for (var k in req) merged[k] = req[k];
    // responseMimeType を指定できるモデルでは指定する。指定できなくても
    // 下の取り出しで救えるので、失敗しても致命的ではない。
    merged.generationConfig = merged.generationConfig || { responseMimeType: 'application/json' };
    return parseJsonText(call(merged));
  }

  /** 文字列から JSON を取り出す（```json の囲みや前後の説明文があっても拾う） */
  function parseJsonText(text) {
    var cleaned = String(text == null ? '' : text).replace(/```json|```/g, '').trim();
    var start = cleaned.indexOf('{');
    var end = cleaned.lastIndexOf('}');
    if (start === -1 || end === -1 || end < start) {
      var s2 = cleaned.indexOf('[');
      var e2 = cleaned.lastIndexOf(']');
      if (s2 === -1 || e2 === -1 || e2 < s2) {
        throw new Error('AI_ERROR: AIの答えをJSONとして読めませんでした');
      }
      return JSON.parse(cleaned.slice(s2, e2 + 1));
    }
    return JSON.parse(cleaned.slice(start, end + 1));
  }

  /**
   * まとめて呼ぶ（UrlFetchApp.fetchAll）。40人分などを直列で回すと6分の実行上限に当たる。
   * レート制限を避けるため chunkSize ずつに小分けする。
   *
   * 例外は投げない。件数と順序を保ったまま {ok, text, error} の配列で返すので、
   * 呼び出し側は「何件成功して、どれが落ちたか」を先生に見せられる。
   * 一時エラーだったものだけ、chunk 単位で再試行する。
   *
   * @param {Array<Object>} reqs call と同じ形のリクエスト配列
   * @return {Array<{ok: boolean, text: string, error: string}>}
   */
  function callAll(reqs, opts) {
    var list = reqs || [];
    var o = opts || {};
    var chunkSize = o.chunkSize || DEFAULTS.chunkSize;
    var maxAttempts = o.maxAttempts || DEFAULTS.maxAttempts;
    var baseDelayMs = o.baseDelayMs || DEFAULTS.baseDelayMs;

    var results = new Array(list.length);
    for (var head = 0; head < list.length; head += chunkSize) {
      var idx = [];
      for (var i = head; i < Math.min(head + chunkSize, list.length); i++) idx.push(i);

      for (var attempt = 0; attempt < maxAttempts && idx.length > 0; attempt++) {
        if (attempt > 0) Utilities.sleep(baseDelayMs * Math.pow(2, attempt - 1));
        var requests = idx.map(function (n) {
          var r = list[n];
          var opt = buildOptions(r);
          opt.url = endpoint(r.model, r.apiVersion);
          return opt;
        });
        var responses = UrlFetchApp.fetchAll(requests);
        var retryIdx = [];
        for (var j = 0; j < responses.length; j++) {
          var n2 = idx[j];
          var code = responses[j].getResponseCode();
          var body = responses[j].getContentText();
          if (code === 200) {
            var t = extractText(JSON.parse(body));
            results[n2] = (t === null || t === '')
              ? { ok: false, text: '', error: 'AIから答えが返りませんでした' }
              : { ok: true, text: t.trim(), error: '' };
          } else if (RETRIABLE.indexOf(code) !== -1) {
            retryIdx.push(n2);
            results[n2] = { ok: false, text: '', error: failure(code, body).message };
          } else {
            results[n2] = { ok: false, text: '', error: failure(code, body).message };
          }
        }
        idx = retryIdx;
      }
    }
    return results;
  }

  return {
    call: call,
    callRaw: callRaw,
    waitMsFor: waitMsFor,
    callJson: callJson,
    callAll: callAll,
    parseJsonText: parseJsonText,
    extractText: extractText,
    endpoint: endpoint,
    DEFAULTS: DEFAULTS,
    RETRIABLE: RETRIABLE,
  };
})();

// Node（テスト）から読めるようにする。GAS では module が無いので何もしない。
if (typeof module !== 'undefined' && module.exports) module.exports = GigaGemini;
