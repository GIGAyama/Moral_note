/**
 * ココロの羅針盤 - Standalone & Auto-Recovery Edition (v2.0)
 * GIGA Standard v2 Compliant
 */

// 定数定義
const APP_NAME = "こころスコープ";
const DB_FILE_NAME = "こころスコープ_Data";
const SCRIPT_PROP = PropertiesService.getScriptProperties();

// Gemini API設定 (オプション)
const GEMINI_API_KEY = SCRIPT_PROP.getProperty('GEMINI_API_KEY');

// 認証設定
const DEFAULT_TEACHER_PASSWORD = "admin";

/* ==============================================
   Core Functions
   ============================================== */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/* ==============================================
   Database Management (Auto-Setup & Recovery)
   ============================================== */

/**
 * データベース(SS)を取得する。
 * DriveAppを使わず、try-catchのみで存在確認を行う。
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
 * 新規スプレッドシート作成とシート構築
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
 * シート構造の定義と適用
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

/* ==============================================
   Data Access Functions (Public API)
   ============================================== */

/**
 * 初期データ取得 (クライアントから最初に呼ばれる)
 * 日付型を文字列に変換して安全に返す。
 */
function getInitialData() {
  try {
    const ss = getDB(); 
    
    // 名簿取得
    const userSheet = ss.getSheetByName('名簿');
    let users = [];
    if (userSheet && userSheet.getLastRow() > 1) {
      users = userSheet.getDataRange().getValues().slice(1)
        .filter(r => !r[3]) // deletedAt
        .map(r => ({ id: r[0], name: r[1], ruby: r[2] }));
    }

    // アクティブな授業を取得
    const sessionSheet = ss.getSheetByName('授業');
    let activeSession = null;
    if (sessionSheet && sessionSheet.getLastRow() > 1) {
      const sessions = sessionSheet.getDataRange().getValues().slice(1)
        .filter(r => r[5] === 'ACTIVE' && !r[7])
        .map(r => {
           // JSONパースの安全策
           let opts = {};
           try { opts = r[4] ? JSON.parse(r[4]) : {}; } catch(e) { console.warn('JSON Parse Error', e); }

           return {
             id: r[0],
             date: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), // 日付を文字列化(重要)
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
 * ポーリング用軽量データ取得 (CacheService使用)
 * 生徒端末からの頻繁なアクセスに耐えるよう最適化
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
        // statusとphaseのみ返す（軽量化）
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
 * 教師用パスワード確認
 */
function checkTeacherPassword(password) {
  const setPass = SCRIPT_PROP.getProperty('TEACHER_PASSWORD');
  const correctPass = setPass || DEFAULT_TEACHER_PASSWORD;
  return { success: (password === correctPass) };
}

/**
 * ログ送信
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
 * 授業ログ取得
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
 * 名簿取得 (一覧のみ)
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

/**
 * 新規授業作成
 */
function createSession(title, inputType, optionsJson) {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][5] === 'ACTIVE') {
      sheet.getRange(i + 1, 6).setValue('CLOSED');
      // キャッシュ破棄
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

/**
 * 授業を終了する（ACTIVE→CLOSED）
 */
function closeSession(sessionId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId && data[i][5] === 'ACTIVE') {
      sheet.getRange(i + 1, 6).setValue('CLOSED');
      // キャッシュ破棄
      CacheService.getScriptCache().remove('POLLING_DATA');
      return { success: true };
    }
  }
  return { success: false, error: 'セッションが見つかりません' };
}

/**
 * フェーズ変更（教師用）
 */
function updateSessionPhase(sessionId, newPhase) {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === sessionId && data[i][5] === 'ACTIVE') {
      sheet.getRange(i + 1, 7).setValue(newPhase);
      // キャッシュ破棄
      CacheService.getScriptCache().remove('POLLING_DATA');
      return { success: true, phase: newPhase };
    }
  }
  return { success: false, error: 'セッションが見つかりません' };
}

/**
 * 特定生徒の授業ログ取得（振り返り用）
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

/**
 * 名簿に生徒を追加
 */
function addStudent(name, ruby) {
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  const studentId = 's' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
  sheet.appendRow([studentId, name, ruby, '']);
  return { success: true, id: studentId };
}

/**
 * 名簿一括登録
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
 * 名簿から生徒を削除（論理削除）
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

/* ==============================================
   Gemini API Integration (AI Socrates)
   ============================================== */

/**
 * Gemini APIで問いかけを生成する
 * APIキーが未設定ならスキップ
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
 * Gemini APIキーを保存する（教師用）
 */
function saveGeminiApiKey(apiKey) {
  SCRIPT_PROP.setProperty('GEMINI_API_KEY', apiKey || '');
  return { success: true };
}

/**
 * Gemini APIキーが設定済みか確認
 */
function hasGeminiApiKey() {
  const key = SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  return { hasKey: !!(key && key.length > 0) };
}

/* ==============================================
   Unit Management (Phase 5: Pre-registry & AI Import)
   ============================================== */

/**
 * 単元を保存・更新
 */
function saveUnit(unitData) {
  const ss = getDB();
  const sheet = ss.getSheetByName('単元');
  
  // 既存があれば更新、なければ新規
  if (unitData.unitId) {
    // 更新ロジック (簡易的な全検索)
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === unitData.unitId) {
        // title, inputType, options, memo
        sheet.getRange(i + 1, 2).setValue(unitData.title);
        sheet.getRange(i + 1, 3).setValue(unitData.inputType);
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(unitData.options));
        sheet.getRange(i + 1, 5).setValue(unitData.memo || '');
        return { success: true, unitId: unitData.unitId };
      }
    }
  }

  // 新規追加
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
 * 単元リストを取得
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
 * 単元から授業を開始
 */
function startSessionFromUnit(unitId) {
  const ss = getDB();
  const unitSheet = ss.getSheetByName('単元');
  const sessionSheet = ss.getSheetByName('授業');
  
  // 単元データ取得
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

  // 新しいセッションを作成
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

/**
 * 授業の匿名意見一覧を取得（共有用）
 */
function getAnonymousOpinions(sessionId) {
  const ss = getDB();
  const sheet = ss.getSheetByName('記録');
  if (!sheet) return [];
  
  // ログ取得 (deletedAtなし)
  const logs = sheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] === sessionId && !r[7])
    .map(r => ({
      phase: r[3],
      value: r[4],
      text: r[5]
    }));

  // 空の意見は除外
  return logs.filter(l => l.text && l.text.trim() !== '');
}

/**
 * AI PDF Import
 * Drive上のPDFファイルIDを受け取り、解析結果を返す
 */
function parseLessonPdf(fileId) {
  const apiKey = GEMINI_API_KEY || SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  if (!apiKey) return { success: false, error: 'AI機能を使うにはAPIキーを設定してください' };

  try {
    // 1. PDF取得 & Base64エンコード
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());
    const mimeType = blob.getContentType(); // application/pdf

    // 2. Gemini APIコール
    const prompt = `
あなたはベテランの学校教師です。提供された「学習指導案（略案）」のPDFを読み取り、授業支援アプリ「こころスコープ」に登録するための設定データを抽出してください。

【授業タイプの判定基準】
- SLIDER (スライダー): 「賛成 vs 反対」「A vs B」のように、意見が2つの対立軸に分かれる場合。葛藤場面や価値判断を問うもの。
- TAGS (感情タグ): 「うれしい、かなしい」などの感情や、「納得、疑問」などの思考状態を多面的に選択させたい場合。

【出力フォーマット（JSONのみ）】
{
  "title": "授業のタイトル（主題名や教材名など、子供に分かりやすいもの）",
  "inputType": "SLIDER" または "TAGS",
  "options": {
    "minLabel": "SLIDERの場合の左端ラベル（例: 正直に言う）",
    "maxLabel": "SLIDERの場合の右端ラベル（例: 黙っている）",
    "tags": ["TAGSの場合の選択肢リスト", "4つ程度"]
  },
  "memo": "指導上の留意点やねらいを100文字以内で要約"
}

※JSON以外の余計なテキストは一切含めないでください。`;

    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
    const payload = {
      contents: [{
        parts: [
          { text: prompt },
          { inline_data: { mime_type: mimeType, data: base64 } }
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

    const json = JSON.parse(response.getContentText());
    if (json.candidates && json.candidates[0] && json.candidates[0].content) {
      const resultText = json.candidates[0].content.parts[0].text;
      const result = JSON.parse(resultText);
      return { success: true, data: result };
    }
    return { success: false, error: 'AIからの応答がありませんでした' };

  } catch (e) {
    console.error(e);
    return { success: false, error: 'PDFの読み取りに失敗しました。ファイルIDが正しいか、権限があるか確認してください。詳細: ' + e.toString() };
  }
}

/**
 * Custom Drive Picker Backend
 * 指定フォルダ（またはルート）内のPDFファイルとフォルダ一覧を返す
 */
function getDriveFiles(folderId) {
  try {
    const parent = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    const folders = [];
    const files = [];

    // Subfolders
    const folderIt = parent.getFolders();
    while (folderIt.hasNext()) {
      const f = folderIt.next();
      folders.push({ id: f.getId(), name: f.getName(), type: 'folder' });
    }

    // PDF Files (MIME type filter)
    const fileIt = parent.getFilesByType(MimeType.PDF);
    while (fileIt.hasNext()) {
      const f = fileIt.next();
      files.push({ id: f.getId(), name: f.getName(), type: 'file', mimeType: f.getMimeType() });
    }
    
    // Sort by name
    folders.sort((a, b) => a.name.localeCompare(b.name));
    files.sort((a, b) => a.name.localeCompare(b.name));

    return { 
      success: true, 
      currentFolderId: parent.getId(),
      currentFolderName: parent.getName(),
      parentFolderId: getParentFolderId(parent),
      items: [...folders, ...files] 
    };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
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

/* ==============================================
   Session History (Phase 4: My Log & Reports)
   ============================================== */

/**
 * 児童ポートフォリオ取得（一括取得・パフォーマンス最適化）
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
    .filter(r => r[2] === studentId && !r[7]) // studentId match & not deleted
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
    
    // ログが一つもない授業は除外するか？ -> いや、参加した授業なら表示したいが、ログがないなら参加してないかも。
    // ここでは「少なくとも1つログがある」または「授業日」でフィルタするが、
    // シンプルに「ログがあるもの」だけ返すのがポートフォリオらしい。
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

  // 日付降順
  return portfolio.sort((a, b) => new Date(b.date) - new Date(a.date));
}

/**
 * My Log (Student History)
 */
function getAllSessions() {
  const ss = getDB();
  const sheet = ss.getSheetByName('授業');
  if (!sheet || sheet.getLastRow() <= 1) return [];

  return sheet.getDataRange().getValues().slice(1)
    .filter(r => !r[7])
    .map(r => {
      let opts = {};
      try { opts = r[4] ? JSON.parse(r[4]) : {}; } catch(e) {}
      return {
        id: r[0],
        date: r[1] instanceof Date ? r[1].toISOString() : String(r[1]),
        title: r[2],
        inputType: r[3],
        options: opts,
        status: r[5],
        phase: r[6] || 'BEFORE'
      };
    });
}

/**
 * 教師用: 全生徒の変容サマリーを取得
 */
function getStudentSummaries(sessionId) {
  const ss = getDB();
  const logSheet = ss.getSheetByName('記録');
  const userSheet = ss.getSheetByName('名簿');
  if (!logSheet || logSheet.getLastRow() <= 1) return [];

  // 名簿をマップ化
  const users = {};
  if (userSheet && userSheet.getLastRow() > 1) {
    userSheet.getDataRange().getValues().slice(1)
      .filter(r => !r[3])
      .forEach(r => { users[r[0]] = { name: r[1], ruby: r[2] }; });
  }

  // ログを生徒ごとに集計
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
 * 【重要】初回承認用・手動セットアップ関数
 * 1. エディタのツールバーで「manualSetup」を選択
 * 2. 「実行」をクリック
 * 3. 権限の承認画面が出たら許可する
 */
function manualSetup() {
  try {
    const ss = getDB();
    console.log("---------------------------------------------------");
    console.log("✅ セットアップ完了！");
    console.log("データベースID:", ss.getId());
    console.log("データベース名:", ss.getName());
    console.log("WebアプリのURLを再読み込みしてください。");
    console.log("---------------------------------------------------");
  } catch(e) {
    console.error("セットアップ失敗:", e);
  }
}
