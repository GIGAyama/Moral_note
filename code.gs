/**
 * ココロの羅針盤 - Standalone & Auto-Recovery Edition (v2.4)
 * GIGA Standard v2 Compliant
 * 
 * 【教員向け解説】
 * このスクリプトは、Googleスプレッドシートをデータベースとして利用し、
 * 授業の進行、生徒の意見集約、AI分析を行うためのバックエンドプログラムです。
 */

// =================================================================
// 1. 定数・設定 (CONSTANTS)
// =================================================================
const APP_NAME = "こころスコープ";
const DB_FILE_NAME = "こころスコープ_Data";
const SCRIPT_PROP = PropertiesService.getScriptProperties();

// Gemini API設定 (APIキーは「先生用ダッシュボード」の設定画面から入力します)
const GEMINI_API_KEY = SCRIPT_PROP.getProperty('GEMINI_API_KEY');

// 認証設定 (デフォルトパスワード)
const DEFAULT_TEACHER_PASSWORD = "admin";

// =================================================================
// 2. コア機能 (CORE FUNCTIONS)
// =================================================================

/**
 * Webアプリとしてのアクセスポイント (GETリクエスト処理)
 * index.html を表示します。
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTMLファイル内で別のファイルを読み込むための関数
 * js.html や css.html をインクルードするのに使用します。
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =================================================================
// 3. データベース管理 (DATABASE MANAGEMENT)
// =================================================================

/**
 * データベース(スプレッドシート)を取得します。
 * IDが存在しない場合やファイルが見つからない場合は、自動的に新規作成・修復を試みます。
 * @return {Spreadsheet} スプレッドシートオブジェクト
 */
function getDB() {
  const dbId = SCRIPT_PROP.getProperty('DB_ID');
  let ss = null;

  // 1. 既存IDがあるか確認
  if (dbId) {
    try {
      ss = SpreadsheetApp.openById(dbId);
      // シート構造の健全性チェック（名簿シートがなければ修復）
      if (!ss.getSheetByName('名簿')) {
        setupSheets(ss);
      }
    } catch (e) {
      console.warn("DB access failed. ID exists but open failed.", e);
      ss = null;
    }
  }

  // 2. SSが取得できなかった場合、新規作成 (Auto-Setup)
  if (!ss) {
    try {
      ss = createDB();
    } catch (e) {
      console.error("Failed to create DB.", e);
      throw new Error("データベースの作成に失敗しました。Googleドライブの容量等を確認してください。");
    }
  }

  return ss;
}

/**
 * 新規にデータベース用スプレッドシートを作成し、初期設定を行います。
 */
function createDB() {
  const ss = SpreadsheetApp.create(DB_FILE_NAME);
  const newId = ss.getId();
  
  // プロパティにIDを保存
  SCRIPT_PROP.setProperty('DB_ID', newId);

  // シート構築
  setupSheets(ss);

  return ss;
}

/**
 * スプレッドシートのシート構造を定義・適用します。
 * 必要なシートがない場合は自動作成します。
 */
function setupSheets(ss) {
  // 1. 設定シート
  let configSheet = ss.getSheetByName('設定');
  if (!configSheet) {
    configSheet = ss.insertSheet('設定');
    configSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]).setBackground('#e8eaed').setFontWeight('bold');
    configSheet.getRange(2, 1, 2, 2).setValues([
      ['AppName', APP_NAME],
      ['GeminiApiKey', '']
    ]);
  }

  // 2. 名簿シート
  let userSheet = ss.getSheetByName('名簿');
  if (!userSheet) {
    userSheet = ss.insertSheet('名簿');
    userSheet.getRange(1, 1, 1, 4).setValues([['studentId', 'name', 'ruby', 'deletedAt']]).setBackground('#e8eaed').setFontWeight('bold');
    // デモデータ
    userSheet.getRange(2, 1, 3, 4).setValues([
      ['s001', '佐藤 健太', 'さとう けんた', ''],
      ['s002', '鈴木 愛', 'すずき あい', ''],
      ['s003', '高橋 翔', 'たかはし かける', '']
    ]);
  }

  // 3. 授業シート
  let sessionSheet = ss.getSheetByName('授業');
  if (!sessionSheet) {
    sessionSheet = ss.insertSheet('授業');
    sessionSheet.getRange(1, 1, 1, 8).setValues([['sessionId', 'date', 'title', 'inputType', 'options', 'status', 'phase', 'deletedAt']]).setBackground('#e8eaed').setFontWeight('bold');
    // デモセッション
    const demoOptions = JSON.stringify({minLabel: '正直に言う', maxLabel: '黙っている', tags: ['葛藤', '不安', '決意']});
    sessionSheet.getRange(2, 1, 1, 8).setValues([
      ['demo_01', new Date(), '正直な心（デモ）', 'SLIDER', demoOptions, 'ACTIVE', 'BEFORE', '']
    ]);
  }

  // 4. 記録シート
  let logSheet = ss.getSheetByName('記録');
  if (!logSheet) {
    logSheet = ss.insertSheet('記録');
    logSheet.getRange(1, 1, 1, 8).setValues([['logId', 'sessionId', 'studentId', 'phase', 'value', 'text', 'timestamp', 'deletedAt']]).setBackground('#e8eaed').setFontWeight('bold');
  }

  // 5. 単元シート (Unit Management)
  let unitSheet = ss.getSheetByName('単元');
  if (!unitSheet) {
    unitSheet = ss.insertSheet('単元');
    unitSheet.getRange(1, 1, 1, 7).setValues([['unitId', 'title', 'inputType', 'options', 'memo', 'createdAt', 'deletedAt']]).setBackground('#e8eaed').setFontWeight('bold');
  }

  // デフォルトシートの削除
  const sheet1 = ss.getSheetByName('シート1');
  if (sheet1) {
    try { ss.deleteSheet(sheet1); } catch(e) {}
  }
}

// =================================================================
// 4. データアクセス・API (DATA ACCESS)
// =================================================================

/**
 * アプリ起動時の初期データを取得します。
 * 名簿と現在アクティブな授業情報を返します。
 */
function getInitialData() {
  try {
    const ss = getDB(); 
    
    // 名簿取得
    const userSheet = ss.getSheetByName('名簿');
    let users = [];
    if (userSheet && userSheet.getLastRow() > 1) {
      users = userSheet.getDataRange().getValues().slice(1)
        .filter(r => !r[3]) // deletedAtがないもの
        .map(r => ({ id: r[0], name: r[1], ruby: r[2] }));
    }

    // アクティブな授業を取得
    const sessionSheet = ss.getSheetByName('授業');
    let activeSession = null;
    if (sessionSheet && sessionSheet.getLastRow() > 1) {
      const sessions = sessionSheet.getDataRange().getValues().slice(1)
        .filter(r => r[5] === 'ACTIVE' && !r[7])
        .map(r => {
           let opts = {};
           try { opts = r[4] ? JSON.parse(r[4]) : {}; } catch(e) { console.warn('JSON Parse Error', e); }

           // 日付は文字列に変換して返す（GASのDateオブジェクト対策）
           return {
             id: r[0],
             date: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
             title: r[2],
             inputType: r[3],
             options: opts,
             phase: r[6] || 'BEFORE'
           };
        });
      if (sessions.length > 0) activeSession = sessions[0];
    }

    return { success: true, users: users, activeSession: activeSession };

  } catch (e) {
    console.error(e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 生徒端末からの定期的な状態確認（ポーリング）用関数
 * CacheServiceを使用してスプレッドシートへのアクセスを減らし、高速に応答します。
 */
function getPollingData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('POLLING_DATA');
  
  if (cached) {
    return JSON.parse(cached);
  }

  // キャッシュがない場合はDBから取得
  const ss = getDB();
  const sessionSheet = ss.getSheetByName('授業');
  let data = { activeSession: null };

  if (sessionSheet && sessionSheet.getLastRow() > 1) {
    const sessions = sessionSheet.getDataRange().getValues().slice(1)
      .filter(r => r[5] === 'ACTIVE' && !r[7]);
    
    if (sessions.length > 0) {
      const r = sessions[0];
      data.activeSession = {
        id: r[0],
        // statusとphaseのみ返す（通信量削減）
        status: r[5],
        phase: r[6] || 'BEFORE'
      };
    }
  }

  // 10秒間キャッシュする
  cache.put('POLLING_DATA', JSON.stringify(data), 10);
  return data;
}

/**
 * 授業の進行フェーズを変更します（教師用）
 * @param {string} sessionId - 授業ID
 * @param {string} newPhase - BEFORE, AFTER, CLOSED
 */
function updateSessionPhase(sessionId, newPhase) {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId && data[i][5] === 'ACTIVE') {
      sheet.getRange(i + 1, 7).setValue(newPhase);
      // キャッシュを破棄して即座に反映
      CacheService.getScriptCache().remove('POLLING_DATA');
      return { success: true, phase: newPhase };
    }
  }
  return { success: false, error: 'セッションが見つかりません' };
}

/**
 * 授業を終了します（ACTIVEステータスをCLOSEDに変更）
 */
function closeSession(sessionId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId && data[i][5] === 'ACTIVE') {
      sheet.getRange(i + 1, 6).setValue('CLOSED');
      CacheService.getScriptCache().remove('POLLING_DATA');
      return { success: true };
    }
  }
  return { success: false, error: 'セッションが見つかりません' };
}

/**
 * 新規授業を作成します
 * @param {string} title - 授業名
 * @param {string} inputType - SLIDER or TAGS
 * @param {string} optionsJson - 設定オプションのJSON文字列
 */
function createSession(title, inputType, optionsJson) {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  
  // 既存のアクティブな授業があれば自動的に終了させる
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'ACTIVE') {
      sheet.getRange(i + 1, 6).setValue('CLOSED');
      CacheService.getScriptCache().remove('POLLING_DATA');
    }
  }
  
  const sessionId = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  sheet.appendRow([
    sessionId,
    new Date(),
    title,
    inputType,
    optionsJson,
    'ACTIVE',
    'BEFORE',
    ''
  ]);
  
  return { success: true };
}

// =================================================================
// 5. 生徒管理 (STUDENT MANAGEMENT)
// =================================================================

/**
 * 名簿に生徒を1名追加します
 */
function addStudent(name, ruby) {
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  const studentId = 's' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
  sheet.appendRow([studentId, name, ruby, '']);
  return { success: true, id: studentId };
}

/**
 * 名簿に生徒を一括追加します（CSV/Excel形式の貼り付け対応）
 * @param {Array} students - {name, ruby} の配列
 */
function addStudentBulk(students) {
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  if (!sheet) return { success: false, error: '名簿シートが見つかりません' };
  
  const rows = students.map((s, i) => {
    // ユニークID生成 (タイムスタンプ + インデックスで重複回避)
    const suffix = ('000' + i).slice(-3);
    const studentId = 's' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + suffix;
    return [studentId, s.name, s.ruby, ''];
  });
  
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  
  return { success: true, count: rows.length };
}

/**
 * 名簿から生徒を削除します（物理削除ではなく、削除日時を入れる論理削除）
 */
function deleteStudent(studentId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentId) {
      sheet.getRange(i + 1, 4).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false, error: '生徒が見つかりません' };
}

/**
 * 名簿リストを取得します
 */
function getStudents() {
  const ss = getDB();
  const userSheet = ss.getSheetByName('名簿');
  if (!userSheet || userSheet.getLastRow() <= 1) return { success: true, users: [] };

  const users = userSheet.getDataRange().getValues().slice(1)
    .filter(r => !r[3]) // deletedAt
    .map(r => ({ id: r[0], name: r[1], ruby: r[2] }));
    
  return { success: true, users: users };
}

// =================================================================
// 6. ログ・分析 (LOGGING & ANALYTICS)
// =================================================================

/**
 * 生徒からの回答ログを保存します
 */
function submitLog(data) {
  const ss = getDB();
  const sheet = ss.getSheetByName('記録');
  const logId = Utilities.getUuid();
  const timestamp = new Date();
  
  sheet.appendRow([
    logId,
    data.sessionId,
    data.studentId,
    data.phase,
    data.value,
    data.text,
    timestamp,
    ''
  ]);

  return { success: true, logId: logId };
}

/**
 * 教師用ダッシュボード向けの授業ログ取得（散布図表示用）
 */
function getSessionLogs(sessionId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('記録');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const rawData = sheet.getDataRange().getValues();
  
  const logs = rawData.slice(1)
    .filter(r => r[1] === sessionId && !r[7])
    .map(r => ({
      logId: r[0],
      studentId: r[2],
      phase: r[3],
      value: r[4],
      text: r[5]
    }));
  return logs;
}

/**
 * 特定生徒の過去の授業ログを含めたポートフォリオデータを取得します
 */
function getStudentPortfolio(studentId) {
  const ss = getDB();
  
  // 1. 全授業取得
  const sessionSheet = ss.getSheetByName('授業');
  if (!sessionSheet) return [];
  const sessions = sessionSheet.getDataRange().getValues().slice(1)
    .filter(r => !r[7]) // deletedAt
    .map(r => ({
      id: r[0],
      date: r[1],
      title: r[2],
      inputType: r[3],
      options: r[4] ? JSON.parse(r[4]) : {}
    }));

  // 2. 生徒の全ログ取得
  const logSheet = ss.getSheetByName('記録');
  if (!logSheet) return [];
  const logs = logSheet.getDataRange().getValues().slice(1)
    .filter(r => r[2] === studentId && !r[7])
    .map(r => ({
      sessionId: r[1],
      phase: r[3],
      value: r[4],
      text: r[5]
    }));

  // 3. データ結合
  const portfolio = sessions.map(s => {
    const sLogs = logs.filter(l => l.sessionId === s.id);
    const before = sLogs.find(l => l.phase === 'BEFORE');
    const after = sLogs.find(l => l.phase === 'AFTER');
    
    // ログがない授業はスキップ
    if (!before && !after) return null;

    return {
      title: s.title,
      date: s.date instanceof Date ? s.date.toISOString() : String(s.date),
      inputType: s.inputType,
      options: s.options,
      before: before ? { value: before.value, text: before.text } : null,
      after: after ? { value: after.value, text: after.text } : null
    };
  }).filter(p => p !== null);

  // 日付の新しい順にソート
  return portfolio.sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * 生徒間共有用の匿名意見一覧を取得
 * 名前を含まず、意見のみを返します。
 */
function getAnonymousOpinions(sessionId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('記録');
  if (!sheet) return [];
  
  // ログ取得
  const logs = sheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] === sessionId && !r[7])
    .map(r => ({
      phase: r[3],
      value: r[4],
      text: r[5]
    }));

  // 空の記述は除外
  return logs.filter(l => l.text && l.text.trim() !== '');
}

/**
 * 教師用レポート: 生徒ごとの変容サマリー（Before -> After）を取得
 */
function getStudentSummaries(sessionId) {
  const ss = getDB();
  const logSheet = ss.getSheetByName('記録');
  const userSheet = ss.getSheetByName('名簿');
  if (!logSheet || logSheet.getLastRow() <= 1) return [];

  // 名簿マップ作成
  const users = {};
  if (userSheet && userSheet.getLastRow() > 1) {
    userSheet.getDataRange().getValues().slice(1)
      .filter(r => !r[3])
      .forEach(r => { users[r[0]] = { name: r[1], ruby: r[2] }; });
  }

  // ログ集計
  const logData = logSheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] === sessionId && !r[7]);

  const grouped = {};
  logData.forEach(r => {
    const sid = r[2];
    if (!grouped[sid]) grouped[sid] = {};
    grouped[sid][r[3]] = {
      value: r[4],
      text: r[5],
      timestamp: r[6] instanceof Date ? r[6].toISOString() : String(r[6])
    };
  });

  return Object.keys(grouped).map(sid => ({
    studentId: sid,
    name: users[sid] ? users[sid].name : sid,
    ruby: users[sid] ? users[sid].ruby : '',
    before: grouped[sid]['BEFORE'] || null,
    after: grouped[sid]['AFTER'] || null
  }));
}

/**
 * 特定生徒の特定授業でのログを取得（振り返り入力画面での過去ログ表示用）
 */
function getStudentLogs(sessionId, studentId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('記録');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  const rawData = sheet.getDataRange().getValues();
  
  return rawData.slice(1)
    .filter(r => r[1] === sessionId && r[2] === studentId && !r[7])
    .map(r => ({
      phase: r[3],
      value: r[4],
      text: r[5],
      timestamp: r[6] instanceof Date ? r[6].toISOString() : String(r[6])
    }));
}

// =================================================================
// 7. 単元管理 (UNIT MANAGEMENT)
// =================================================================

/**
 * 単元の保存・新規作成
 */
function saveUnit(unitData) {
  const ss = getDB();
  const sheet = ss.getSheetByName('単元');
  // 既存データがある場合は更新
  if (unitData.unitId) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === unitData.unitId) {
        sheet.getRange(i + 1, 2).setValue(unitData.title);
        sheet.getRange(i + 1, 3).setValue(unitData.inputType);
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(unitData.options));
        sheet.getRange(i + 1, 5).setValue(unitData.memo || '');
        return { success: true, unitId: unitData.unitId };
      }
    }
  }

  // 新規作成
  const newId = 'u' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
  sheet.appendRow([
    newId,
    unitData.title,
    unitData.inputType,
    JSON.stringify(unitData.options),
    unitData.memo || '',
    new Date(),
    ''
  ]);
  return { success: true, unitId: newId };
}

/**
 * 単元リストの取得
 */
function getUnits() {
  const ss = getDB();
  const sheet = ss.getSheetByName('単元');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getDataRange().getValues().slice(1)
    .filter(r => !r[6]) // deletedAt
    .map(r => {
      let opts = {};
      try { opts = r[3] ? JSON.parse(r[3]) : {}; } catch(e) {}
      return {
        unitId: r[0],
        title: r[1],
        inputType: r[2],
        options: opts,
        memo: r[4],
        createdAt: r[5] instanceof Date ? r[5].toISOString() : String(r[5])
      };
    });
}

/**
 * 指定した単元データをもとに、新しい授業を開始します
 */
function startSessionFromUnit(unitId) {
  const ss = getDB();
  const unitSheet = ss.getSheetByName('単元');
  const sessionSheet = ss.getSheetByName('授業');
  
  const unitData = unitSheet.getDataRange().getValues().slice(1).find(r => r[0] === unitId);
  if (!unitData) return { success: false, error: '単元が見つかりません' };

  // 既存のアクティブセッションをクローズ
  const sessionData = sessionSheet.getDataRange().getValues();
  for (let i = 1; i < sessionData.length; i++) {
    if (sessionData[i][5] === 'ACTIVE') {
      sessionSheet.getRange(i + 1, 6).setValue('CLOSED');
      CacheService.getScriptCache().remove('POLLING_DATA');
    }
  }

  // 新規セッション
  const sessionId = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
  sessionSheet.appendRow([
    sessionId,
    new Date(),
    unitData[1], // title
    unitData[2], // inputType
    unitData[3], // optionsJson
    'ACTIVE',
    'BEFORE',
    ''
  ]);

  return { success: true };
}

/**
 * 単元を削除（論理削除）
 */
function deleteUnit(unitId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('単元');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === unitId) {
      sheet.getRange(i + 1, 7).setValue(new Date());
      return { success: true };
    }
  }
  return { success: false };
}

// =================================================================
// 8. AI機能 / Gemini API (Gemini INTEGRATION)
// =================================================================

/**
 * AIによるソクラテス的問いかけ生成
 */
function generateSocraticQuestion(sessionTitle, studentText, inputType, studentValue) {
  const apiKey = GEMINI_API_KEY || SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  if (!apiKey) {
    return { success: false, reason: 'NO_API_KEY' };
  }

  const prompt = `あなたは小学校の道徳の授業で、子供たちの「深い思考」を引き出すソクラテス的対話の専門家です。

【授業テーマ】${sessionTitle}
【子供の回答】${studentText || '（記述なし）'}
【入力タイプ】${inputType === 'SLIDER' ? 'スライダー（値: ' + studentValue + '/100）' : inputType}

以下のルールに従い、この子供に対して1つだけ「問いかけ」を生成してください：
- 小学3〜6年生が理解できるやさしい日本語で書く
- 40文字以内の短い問いかけにする
- 答えを誘導せず、考えを広げる問いにする
- 「なぜ？」「もし〜だったら？」「具体的には？」のいずれかのパターンを使う
- 記述が空や極端に短い場合は、まず自分の考えを言葉にするよう促す

問いかけのみを出力してください（説明や前置き不要）。`;

  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 100, temperature: 0.7 }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      const question = json.candidates[0].content.parts[0].text.trim();
      return { success: true, question: question };
    }
    return { success: false, reason: 'NO_RESPONSE' };
  } catch (e) {
    console.error('Gemini API Error:', e);
    return { success: false, reason: e.toString() };
  }
}

/**
 * PDF指導案からの授業設定抽出
 * クライアントから送信されたBase64データを受け取ります。
 */
function parseLessonPdf(base64Data) {
  const apiKey = GEMINI_API_KEY || SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { success: false, error: 'AI機能を使うにはAPIキーを設定してください' };

  try {
    // Gemini APIコール
    const prompt = `
あなたはベテランの学校教師です。提供された「学習指導案（略案）」または「年間指導計画」のPDFを読み取り、授業支援アプリ「こころスコープ」に登録するための設定データを抽出してください。

【抽出ルール】
- 文書内に複数の単元（授業）が含まれている場合は、可能な限りすべて抽出してください。
- 文書が1つの単元のみの場合は、それを1つだけ抽出してください。

【授業タイプの判定基準】
- SLIDER (スライダー): 「賛成 vs 反対」「A vs B」のように、意見が2つの対立軸に分かれる場合。葛藤場面や価値判断を問うもの。
- TAGS (感情タグ): 「うれしい、かなしい」などの感情や、「納得、疑問」などの思考状態を多面的に選択させたい場合。

【出力フォーマット（JSONのみ）】
{
  "units": [
    {
      "title": "授業のタイトル（主題名・教材名など）",
      "inputType": "SLIDER" または "TAGS",
      "options": {
        "minLabel": "SLIDERの場合の左端（例: 正直に言う）",
        "maxLabel": "SLIDERの場合の右端（例: 黙っている）",
        "tags": ["TAGSの場合の選択肢リスト（4つ程度）"]
      },
      "memo": "ねらいや留意点の要約（100文字以内）"
    }
  ]
}

※JSON以外の余計なテキストは一切含めないでください。`;

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: 'application/pdf', data: base64Data } }
        ]
      }],
      generationConfig: { response_mime_type: "application/json" }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    console.log(`Gemini API Response: ${responseCode} - ${responseText}`); // DEBUG LOG

    if (responseCode !== 200) {
      return { success: false, error: `AI API Error (${responseCode}): ${responseText}` };
    }

    const json = JSON.parse(responseText);
    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      const resultText = json.candidates[0].content.parts[0].text;
      // Clean up markdown code blocks if present (relaxed regex)
      const cleanedText = resultText.replace(/```json/g, '').replace(/```/g, '').trim();
      const result = JSON.parse(cleanedText);
      return { success: true, data: result };
    }
    return { success: false, error: 'AIからの応答がありませんでした (No candidates)' };

  } catch (e) {
    console.error(e);
    return { success: false, error: 'PDFの解析に失敗: ' + e.toString() };
  }
}

/**
 * Gemini APIキーの保存
 */
function saveGeminiApiKey(apiKey) {
  SCRIPT_PROP.setProperty('GEMINI_API_KEY', apiKey || '');
  return { success: true };
}

/**
 * 権限認証を強制するためのダミー関数
 * エディタ上でこの関数を選択して実行すると、必要な権限の承認画面が表示されます。
 */
function forceAuth() {
  SpreadsheetApp.getActiveSpreadsheet();
  Session.getActiveUser().getEmail();
  UrlFetchApp.fetch("https://www.google.com");
  console.log("Auth complete");
}

/**
 * Gemini APIキーの有無確認
 */
function hasGeminiApiKey() {
  const key = SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  return { hasKey: !!(key && key.length > 0) };
}

// =================================================================
// 9. その他ユーティリティ (UTILITIES)
// =================================================================

/**
 * 全授業セッション一覧を取得（履歴表示用）
 */
function getAllSessions() {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getDataRange().getValues().slice(1)
    .filter(r => !r[7])
    .map(r => {
      let opts = {};
      try { opts = r[4] ? JSON.parse(r[4]) : {}; } catch (e) { }
      return {
        id: r[0],
        date: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
        title: r[2],
        inputType: r[3],
        options: opts,
        status: r[5],
        phase: r[6]
      };
    })
    .sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * AIによる所見自動作成
 * 児童の全学習記録をもとに通知表用の所見文を生成します
 */
function generateObservation(studentId) {
  const apiKey = GEMINI_API_KEY || SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { success: false, error: 'Gemini APIキーを設定してください' };

  const ss = getDB();
  const userSheet = ss.getSheetByName('名簿');
  const users = userSheet.getDataRange().getValues().slice(1);
  const student = users.find(r => r[0] === studentId && !r[3]);
  if (!student) return { success: false, error: '児童が見つかりません' };

  const portfolio = getStudentPortfolio(studentId);
  if (!portfolio || portfolio.length === 0) return { success: false, error: '学習記録がありません' };

  const typeLabel = { SLIDER: 'スライダー', TAGS: '感情タグ', QUADRANT: '座標軸', RANKING: 'ランキング', CURVE: '心情曲線', MANDALA: 'マンダラ', ACTION: '宣言カード' };

  const historyText = portfolio.map(p => {
    let entry = `【${p.title}】(形式:${typeLabel[p.inputType] || p.inputType}, 日付:${new Date(p.date).toLocaleDateString('ja-JP')})`;
    if (p.before) entry += `\n  導入時: 値=${p.before.value}, 記述="${p.before.text || '（なし）'}"`;
    if (p.after) entry += `\n  振り返り: 値=${p.after.value}, 記述="${p.after.text || '（なし）'}"`;
    return entry;
  }).join('\n\n');

  const prompt = `あなたはベテランの小学校教師です。道徳の授業における児童の学習記録を分析し、通知表に記載する「所見」を作成してください。

【児童名】${student[1]}（${student[2]}）

【学習記録（${portfolio.length}回分）】
${historyText}

【所見作成のルール】
- 200〜300文字程度で簡潔にまとめる
- 児童の成長や変容を具体的に記述する
- 導入時と振り返り時の変化に注目する
- 記述内容から読み取れる思考の深まりを評価する
- ポジティブな表現を中心にしつつ、今後の課題も示唆する
- 「〜できました」「〜が見られました」などの所見文体で書く
- 具体的な授業名やエピソードを含める

所見文のみを出力してください（説明や前置き不要）。`;

  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
    const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { maxOutputTokens: 500, temperature: 0.5 }
    };

    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      const observation = json.candidates[0].content.parts[0].text.trim();
      return { success: true, observation: observation, name: student[1], ruby: student[2] };
    }
    return { success: false, error: 'AIからの応答がありませんでした' };
  } catch (e) {
    console.error('Gemini API Error:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 教師用パスワードの照合
 */
function checkTeacherPassword(password) {
  const setPass = SCRIPT_PROP.getProperty('TEACHER_PASSWORD');
  const correctPass = setPass || DEFAULT_TEACHER_PASSWORD;
  return { success: (password === correctPass) };
}


function getParentFolderId(folder) {
  try {
    const parents = folder.getParents();
    if (parents.hasNext()) return parents.next().getId();
    return null;
  } catch(e) {
    return null;
  }
}

/**
 * 初回セットアップ用関数（手動実行用）
 */
function manualSetup() {
  try {
    const ss = getDB();
    console.log("✅ セットアップ完了！ DB ID:", ss.getId());
  } catch(e) {
    console.error("セットアップ失敗:", e);
  }
}
