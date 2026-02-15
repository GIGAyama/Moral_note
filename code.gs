/**
 * ココロの羅針盤 - Standalone & Auto-Recovery Edition (v1.3)
 * GIGA Standard v2 Compliant
 */

// 定数定義
const APP_NAME = "ココロの羅針盤";
const DB_FILE_NAME = "ココロの羅針盤_Data";
const SCRIPT_PROP = PropertiesService.getScriptProperties();

// Gemini API設定 (オプション)
const GEMINI_API_KEY = SCRIPT_PROP.getProperty('GEMINI_API_KEY');

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
      phase: r[3],
      value: r[4],
      text: r[5]
    }));

  return logs;
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
