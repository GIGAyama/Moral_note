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

// 認証設定
// ※既定パスワード（admin）は廃止しました。パスワードが未設定のときは
//   「まず管理者パスワードを設定してください」という状態になります。
// ※パスワードは平文で保存せず、SHA-256のハッシュ値を TEACHER_PASSWORD_HASH に保存します。
const MAX_LOGIN_ATTEMPTS = 5;         // 最大試行回数
const LOCKOUT_DURATION_MIN = 10;      // ロックアウト時間（分）

// =================================================================
// 1b. バリデーション (VALIDATION)
// =================================================================
const MAX_NAME_LENGTH = 50;
const MAX_TEXT_LENGTH = 1000;
const MAX_VALUE_LENGTH = 5000;

function sanitizeText(text) {
  if (!text) return '';
  return String(text).replace(/<[^>]*>/g, '').substring(0, MAX_TEXT_LENGTH);
}

function validateStudentName(name) {
  if (!name || typeof name !== 'string') return false;
  const trimmed = name.trim();
  return trimmed.length > 0 && trimmed.length <= MAX_NAME_LENGTH;
}

function validateEmail(email) {
  if (!email) return true; // 空はOK
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email).trim());
}

// =================================================================
// 1c. 認可 (AUTHORIZATION)
// =================================================================
/**
 * 【教員向け解説】
 * ここは「アクセスしている人が先生かどうか」をサーバー側で判定する部分です。
 * 画面のパスワード入力は表示を切り替えるためのものにすぎず、
 * 本当の鍵はこの assertTeacher_() です。先生専用の関数は必ず冒頭でこれを呼びます。
 *
 * 【初回セットアップの手順】
 *  1. スクリプトエディタで manualSetup() を1回実行してください。
 *     → 実行した人のメールアドレスが OWNER_EMAIL に記録され、その人が先生になります。
 *  2. 先生が複数いる場合は、スクリプトプロパティ TEACHER_EMAILS に
 *     カンマ区切りでメールアドレスを登録してください。
 *     （例: sato@example.ed.jp,suzuki@example.ed.jp）
 *     TEACHER_EMAILS が設定されている場合は、そちらを優先して照合します。
 *  ※どちらも登録されていない場合は、安全のため全員拒否します（誰でも通す動きにはしません）。
 *
 * 【デプロイ設定について】
 * appsscript.json では webapp を executeAs: USER_DEPLOYING / access: DOMAIN にしています。
 * USER_ACCESSING にすると児童にもデータ用スプレッドシートの権限が必要になり、
 * シートを直接開けば全員分の記録が読めてしまうため、ここでの認可が意味をなさなくなるからです。
 * （同一ドメインであれば USER_DEPLOYING でも Session.getActiveUser().getEmail() は取得できます）
 */

/**
 * アクセスしている人のメールアドレスを取得します（小文字・前後の空白なし）。
 * 匿名アクセスなどで取得できない場合は空文字を返します。
 */
function getCallerEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
  } catch (e) {
    console.warn('メールアドレスを取得できませんでした:', e);
    return '';
  }
}

/**
 * スクリプトプロパティ TEACHER_EMAILS（カンマ区切り）を配列にして返します。
 */
function getTeacherEmails_() {
  const raw = SCRIPT_PROP.getProperty('TEACHER_EMAILS') || '';
  return String(raw).split(',')
    .map(v => String(v).toLowerCase().trim())
    .filter(v => v !== '');
}

/**
 * 初回セットアップを実行した人のメールアドレスを OWNER_EMAIL に記録します。
 * すでに記録されている場合は上書きしません（あとから来た人が管理者になれないようにするため）。
 * ※児童のアクセスで勝手に記録されないよう、手動実行の関数からのみ呼び出します。
 */
function rememberOwnerEmail_() {
  const existing = SCRIPT_PROP.getProperty('OWNER_EMAIL');
  if (existing) return existing;
  const email = getCallerEmail_();
  if (!email) return '';
  SCRIPT_PROP.setProperty('OWNER_EMAIL', email);
  console.log('OWNER_EMAIL を記録しました: ' + email);
  return email;
}

/**
 * このウェブアプリを公開した人（＝スプレッドシートをコピーして配った先生）の
 * メールアドレスを返します。
 *
 * appsscript.json は executeAs: USER_DEPLOYING なので、ウェブアプリの中では
 * 「実行しているユーザー」＝公開した先生です。スクリプトエディタから手で
 * 実行したときは、実行した本人になります。
 * userinfo.email スコープだけで取れるので、権限は広がりません。
 */
function getDeployerEmail_() {
  try {
    return String(Session.getEffectiveUser().getEmail() || '').toLowerCase().trim();
  } catch (e) {
    return '';
  }
}

/**
 * アクセスしている人が先生かどうかを判定します。
 *
 * 次の順で照合します。
 *  1. TEACHER_EMAILS（複数の先生で使う学級。設定してあれば、これだけを見ます）
 *  2. OWNER_EMAIL（前の配り方で manualSetup を実行済みの学級。消すと動かなくなります）
 *  3. このアプリを公開した本人（＝コピーしてデプロイした先生）
 *
 * 3 を足した理由:
 * スプレッドシートのコピーで配る形にしたので、先生の手元には毎回まっさらな
 * コピーが届きます。そこで OWNER_EMAIL を頼りにすると、
 * **デプロイした先生自身が「先生として登録されていません」と拒まれます。**
 * 公開した本人であることは Session.getEffectiveUser() で確かめられるので、
 * 初期設定の儀式を1つ減らします。
 *
 * どこにも当てはまらない、またはメールアドレスが取得できない場合は false です
 * （＝誰でも通す動きにはしません）。
 */
function isTeacher_() {
  const email = getCallerEmail_();
  if (!email) return false;

  const teachers = getTeacherEmails_();
  if (teachers.length > 0) {
    return teachers.indexOf(email) !== -1;
  }

  const owner = String(SCRIPT_PROP.getProperty('OWNER_EMAIL') || '').toLowerCase().trim();
  if (owner !== '') return owner === email;

  const deployer = getDeployerEmail_();
  return deployer !== '' && deployer === email;
}

/**
 * 先生でなければ例外を投げて処理を止めます。
 * 先生専用の関数は、いちばん最初にこれを呼んでください。
 */
function assertTeacher_() {
  if (isTeacher_()) return true;

  const email = getCallerEmail_();
  if (!email) {
    // 匿名アクセスでデプロイされている場合もここに来ます（安全のため拒否します）
    throw new Error('権限がありません。学校（同じドメイン）のGoogleアカウントでログインしてください。メールアドレスが確認できないため、先生用の機能は使えません。');
  }
  throw new Error('権限がありません。この機能は先生専用です。（' + email + ' は先生として登録されていません。スクリプトプロパティ TEACHER_EMAILS に追加するか、manualSetup を実行してください）');
}

/**
 * 先生、または「その児童本人」だけに許可します。
 * 児童が自分の記録（マイページ）を見る場合のために用意した判定です。
 * 本人かどうかは、名簿に登録されたメールアドレスとログイン中のメールアドレスで照合します。
 */
function assertTeacherOrSelf_(studentId) {
  if (isTeacher_()) return true;

  const email = getCallerEmail_();
  if (!email) {
    throw new Error('権限がありません。学校（同じドメイン）のGoogleアカウントでログインしてください。メールアドレスが確認できないため、記録は表示できません。');
  }

  const ss = getDB();
  const userSheet = ss.getSheetByName('名簿');
  if (userSheet && userSheet.getLastRow() > 1) {
    const me = userSheet.getDataRange().getValues().slice(1)
      .find(r => r[0] === studentId && !r[3]);
    if (me && me[4] && String(me[4]).toLowerCase().trim() === email) {
      return true;
    }
  }
  throw new Error('権限がありません。自分の記録だけが見られます。（見られない場合は、名簿に自分のメールアドレスが登録されているか先生に確認してください）');
}

/**
 * 児童が読んでよい授業かどうかを確かめます。
 *
 * 先生は、どの授業でも読めます。
 * 児童が読めるのは「いま行われている授業（status が ACTIVE）」だけです。
 * 終わった授業の記述は、先生の画面からしか開けません。
 */
function assertActiveSessionOrTeacher_(ss, sessionId) {
  if (isTeacher_()) return true;

  if (!sessionId) {
    throw new Error('授業が指定されていません。');
  }

  const sheet = ss.getSheetByName('授業');
  if (sheet && sheet.getLastRow() > 1) {
    const row = sheet.getDataRange().getValues().slice(1)
      .find(function (r) { return r[0] === sessionId && !r[7]; });
    if (row && row[5] === 'ACTIVE') return true;
  }

  throw new Error('いま行われている授業のぶんだけが見られます。'
    + '（終わった授業の記録は、先生の画面から見てください）');
}

/**
 * パスワードをSHA-256でハッシュ化して16進文字列にします。
 * 平文のパスワードは保存も比較もしません。
 */
function hashPassword_(password) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(password || ''),
    Utilities.Charset.UTF_8
  );
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    hex += ('0' + (bytes[i] & 0xFF).toString(16)).slice(-2);
  }
  return hex;
}

// =================================================================
// 2. コア機能 (CORE FUNCTIONS)
// =================================================================

/**
 * Webアプリとしてのアクセスポイント (GETリクエスト処理)
 * app-shell.html を表示します。
 *
 * 【なぜ index ではなく app-shell という名前なのか】
 * リポジトリには CNAME があり、GitHub Pages が moral-note.giga-school.com として
 * 中身をそのまま配っています。この外枠を index.html という名前で置くと、
 * **その住所を開いた先生に GAS 用のテンプレートがそのまま配られます。**
 * `<?!= include('css'); ?>` はブラウザには意味が無く黙って捨てられるので、
 * スタイルもスクリプトも当たらない、ほぼ白い画面になります。
 * index.html は導入案内のページに使い、GAS の外枠はこの名前で持ちます。
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('app-shell');
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
 * このアプリが使うシートの並びと、1行目に置く見出し。
 *
 * 【なぜ「見出しの表」を1か所に持つのか】
 * このアプリは、シートを**列の番号**で読み書きしています。たとえば記録シートは
 * `r[2]` が児童ID、`r[4]` が数値、`r[7]` が削除の印です（getSessionLogs など）。
 * 先生が誤字を直しに表を開き、列を1本だけ挿し込むと、
 *   ・散布図の点が全部ずれる
 *   ・消したはずの記録が戻る（deletedAt が別の列になるため）
 *   ・「すでに送信済みです」が出なくなり、同じ児童の記録が二重に入る
 * が、**画面にエラーを1つも出さないまま**起きます。
 * 表を1か所にまとめておけば、その「ずれ」を機械で見つけられます。
 *
 * ここに1行足すと、すでに配ったスプレッドシートにも自動でそろいます（ensureSheets_）。
 */
var SHEETS_ = [
  {
    name: '設定',
    header: ['Key', 'Value'],
    initialRows: [['AppName', APP_NAME], ['GeminiApiKey', '']],
  },
  {
    name: '名簿',
    header: ['studentId', 'name', 'ruby', 'deletedAt', 'email'],
    initialRows: [
      ['s001', '佐藤 健太', 'さとう けんた', '', ''],
      ['s002', '鈴木 愛', 'すずき あい', '', ''],
      ['s003', '高橋 翔', 'たかはし かける', '', ''],
    ],
  },
  {
    name: '授業',
    header: ['sessionId', 'date', 'title', 'inputType', 'options', 'status', 'phase', 'deletedAt'],
    initialRows: 'demoSession',   // 日付を毎回作り直すので、関数で組み立てます
  },
  {
    name: '記録',
    header: ['logId', 'sessionId', 'studentId', 'phase', 'value', 'text', 'timestamp', 'deletedAt'],
  },
  {
    name: '単元',
    header: ['unitId', 'title', 'inputType', 'options', 'memo', 'createdAt', 'deletedAt'],
  },
];

/** 授業シートに置く見本の1行。新しく作ったときだけ入ります。 */
function demoSessionRows_() {
  const demoOptions = JSON.stringify({
    minLabel: '正直に言う', maxLabel: '黙っている', tags: ['葛藤', '不安', '決意'],
  });
  return [['demo_01', new Date(), '正直な心（デモ）', 'SLIDER', demoOptions, 'ACTIVE', 'BEFORE', '']];
}

/** SHEETS_ の initialRows を、実際に書き込む2次元配列にして返します。 */
function initialRowsOf_(spec) {
  if (spec.initialRows === 'demoSession') return demoSessionRows_();
  return spec.initialRows || [];
}

/**
 * 見出し行を書きます。空のシートにだけ使います。
 *
 * @param {boolean} withSamples 見本の行も入れるか。
 *   **シートを新しく作ったときだけ true にします。**
 *   先生が中身を消して空にしただけのシートに見本を書き戻すと、
 *   消したはずの「佐藤 健太」やデモの授業が、次に開いたときに戻ってきます。
 *   先生から見れば「消えない」ので、何度も消すことになります。
 */
function writeHeader_(sheet, spec, withSamples) {
  sheet.getRange(1, 1, 1, spec.header.length)
    .setValues([spec.header])
    .setBackground('#e8eaed')
    .setFontWeight('bold');
  if (!withSamples) return;
  const rows = initialRowsOf_(spec);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, spec.header.length).setValues(rows);
  }
}

/**
 * 足りないシートだけを作ります。
 *
 * ふつうは1枚も足りないので、その場合は**ロックを取らずに帰ります**。
 * 40台が一斉に開く朝の会に、全員がロック待ちの行列に並ぶのを避けるためです。
 * 先生が誤って1枚消してしまったときも、次に開いた人が作り直します。
 * 消えた中身は戻りませんが、画面が真っ白になることは無くなります。
 *
 * @return {Spreadsheet} 受け取ったものをそのまま返します（呼び出し側で繋げて書けるように）
 */
function ensureSheets_(ss) {
  const missing = SHEETS_.filter(function (spec) {
    const sheet = ss.getSheetByName(spec.name);
    return !sheet || sheet.getLastRow() === 0;
  });
  if (!missing.length) return ss;

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    // ロックが取れないのは、ほかの誰かが今まさに作っている最中のときです。
    // 作りかけの表に二重に書き込むより、そのまま返して次の読み込みに任せます。
    return ss;
  }
  try {
    SHEETS_.forEach(function (spec) {
      let sheet = ss.getSheetByName(spec.name);
      if (sheet && sheet.getLastRow() > 0) return;   // ロックを待つ間に誰かが作っていた
      const isNew = !sheet;
      if (isNew) sheet = ss.insertSheet(spec.name);
      writeHeader_(sheet, spec, isNew);
    });

    // 新しいスプレッドシートに最初からある空の「シート1」を片づけます。
    // 中身のあるシートは消しません（先生が自分用に使っていることがあります）。
    ['シート1', 'Sheet1'].forEach(function (name) {
      const sheet = ss.getSheetByName(name);
      if (sheet && sheet.getLastRow() === 0 && ss.getSheets().length > 1) {
        try { ss.deleteSheet(sheet); } catch (e) { /* 消せなくても困りません */ }
      }
    });
  } finally {
    lock.releaseLock();
  }
  return ss;
}

/**
 * 見出し行を、実際の幅ぶん読み出します（点検と修整で同じ見方をするため）。
 */
function readHeaderRow_(sheet, spec) {
  const width = Math.min(
    Math.max(spec.header.length, sheet.getLastColumn()),
    sheet.getMaxColumns()
  );
  const actual = sheet.getRange(1, 1, 1, width).getValues()[0];
  return {
    width: width,
    cell: function (i) {
      const v = actual[i];
      return v === undefined || v === null ? '' : String(v).trim();
    },
  };
}

/**
 * そのシートの列が「ずれている」かどうかを見ます。
 * ずれているシートには、**1列も書き込みません。**
 *
 * ずれている証拠は2つあります。
 *
 *  (a) 想定の範囲に、別の言葉の見出しが入っている
 *  (b) 想定より右に、何か入っている
 *
 * ⚠️ (b) を見落とすと、次の形で記録を壊します。
 *    先生が「記録」シートの timestamp と deletedAt のあいだに1列挿すと、
 *    見出しは […, timestamp, （空）, deletedAt] になります。
 *    (a) だけを見ていると、空欄はずれの証拠にならないので「ずれていない」と
 *    判断し、空いた8列目に deletedAt と書き足してしまいます。
 *    見出しは […, timestamp, deletedAt, deletedAt] になり、
 *    本物の削除の印は9列目に残ったままです。
 *    アプリは削除の印を8列目（r[7]）で読むので、
 *    **消したはずの記録が全部戻り、二重送信の判定も効かなくなります。**
 *    しかも見出しは「正しく見える」ので、点検でも気づけません。
 */
function isSheetShifted_(sheet, spec) {
  const head = readHeaderRow_(sheet, spec);

  for (let i = 0; i < spec.header.length; i++) {
    const now = head.cell(i);
    if (now !== '' && now !== spec.header[i]) return true;      // (a)
  }
  for (let j = spec.header.length; j < head.width; j++) {
    if (head.cell(j) !== '') return true;                        // (b)
  }
  return false;
}

/**
 * シートの作りが SHEETS_ のとおりかを点検します。**ここでは何も書き換えません。**
 *
 * ⚠️ 見出しがずれていても、勝手に上書きしてはいけません。
 *    ずれているのは見出しではなく**中身のほう**なので、見出しだけ正しくすると
 *    「間違った列に正しいラベルが付いた」状態になり、事故が見えなくなります。
 *    どこがどうずれているかを言うだけにして、直すのは人の仕事にします。
 *
 * @return {{sheet: string, kind: string, detail: string, fixable: boolean}[]}
 *         見つかったもの。想定どおりなら空の配列。
 */
function checkSheets_(ss) {
  const found = [];

  SHEETS_.forEach(function (spec) {
    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) {
      found.push({ sheet: spec.name, kind: 'シートが無い', detail: '「' + spec.name + '」シートがありません', fixable: true });
      return;
    }
    if (sheet.getLastRow() === 0) {
      found.push({ sheet: spec.name, kind: '見出しが無い', detail: '1行目が空です', fixable: true });
      return;
    }

    const head = readHeaderRow_(sheet, spec);
    const width = head.width;
    const cell = head.cell;
    const hasRows = sheet.getLastRow() > 1;

    // このシートの列がずれているなら、**1列も直しません**（isSheetShifted_ 参照）。
    // ここで「直せます」と言っておきながら repairSheets_ が何もしないと、
    // 先生は直すつもりで OK を押したのに「直せるところはありませんでした」と
    // 言われることになります。同じ判断を両方で使います。
    const shifted = isSheetShifted_(sheet, spec);

    // (1) 見出しが違う列
    const wrong = [];
    const blank = [];
    for (let i = 0; i < spec.header.length; i++) {
      if (cell(i) === spec.header[i]) continue;
      if (cell(i) === '') blank.push(i); else wrong.push(i);
    }

    // 空欄の見出しは、その列に**データが1つも無く**、かつ
    // **シート全体がずれていないとき**だけ書き足せます。
    blank.forEach(function (i) {
      const dirty = hasRows && columnHasValue_(sheet, i + 1);
      const why = shifted ? '（このシートは列がずれているので、自動では直しません）'
        : dirty ? '（この列にデータが入っているので、自動では直しません）' : '';
      found.push({
        sheet: spec.name,
        kind: '見出しが空',
        detail: (i + 1) + '列目が「' + spec.header[i] + '」のはずが空です' + why,
        fixable: !dirty && !shifted,
      });
    });

    if (wrong.length) {
      const at = wrong[0];
      found.push({
        sheet: spec.name,
        kind: '見出しがちがう',
        detail: (at + 1) + '列目が「' + spec.header[at] + '」のはずが「' + cell(at) + '」になっています'
          + (wrong.length > 1 ? '（ほか ' + (wrong.length - 1) + ' 列もずれています）' : ''),
        fixable: false,
      });
    }

    // (2) 想定より右にある列。読み書きでは無視していますが、列を挿した跡のことがあります。
    const extra = [];
    for (let j = spec.header.length; j < width; j++) {
      if (cell(j) !== '') extra.push(cell(j));
    }
    if (extra.length) {
      found.push({
        sheet: spec.name,
        kind: '列が多い',
        detail: (spec.header.length + 1) + '列目から先に「' + extra.join('」「') + '」があります',
        fixable: false,
      });
    }
  });

  return found;
}

/** その列（1始まり）に、見出し行より下で何か入っているかを見ます。 */
function columnHasValue_(sheet, column) {
  const rows = sheet.getLastRow() - 1;
  if (rows <= 0) return false;
  if (column > sheet.getMaxColumns()) return false;
  return sheet.getRange(2, column, rows, 1).getValues()
    .some(function (r) { return String(r[0] === undefined || r[0] === null ? '' : r[0]).trim() !== ''; });
}

/**
 * 点検で見つかったもののうち、**安全に直せるものだけ**を直します。
 *
 * 直すもの:
 *   ・シートが丸ごと無い     → 見出しを付けて作る
 *   ・1行目が空             → 見出しを書く
 *   ・見出しが空欄で、その列にデータが1つも無い → 見出しを書き足す
 *     （アプリの更新で列が増えたときが、これに当たります）
 *
 * 直さないもの:
 *   ・見出しが別の言葉になっている → 列がずれた跡です。上書きすると事故が隠れます
 *   ・想定より右に列がある         → 先生が自分用に足した列かもしれません
 *
 * @return {{fixed: string[], left: object[]}} 直したものと、人に任せたもの
 */
function repairSheets_(ss) {
  const fixed = [];

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    throw new Error('ほかの人が同時に使っています。少し待ってからもう一度おためしください。');
  }
  try {
    SHEETS_.forEach(function (spec) {
      let sheet = ss.getSheetByName(spec.name);

      if (!sheet) {
        sheet = ss.insertSheet(spec.name);
        writeHeader_(sheet, spec, true);
        fixed.push('「' + spec.name + '」シートを作りました');
        return;
      }
      if (sheet.getLastRow() === 0) {
        // 見出しだけを書きます。見本の行は入れません
        // （中身を消して空にしただけかもしれないので、勝手に足しません）。
        writeHeader_(sheet, spec, false);
        fixed.push('「' + spec.name + '」の見出しを書きました');
        return;
      }

      // 見出しが空欄で、その列にデータが無いものだけ書き足します。
      if (sheet.getMaxColumns() < spec.header.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), spec.header.length - sheet.getMaxColumns());
      }
      // ⚠️ このシートの列がずれていたら、**1列も触りません**（isSheetShifted_ 参照）。
      //    ずれているシートは、人が中身を見て直すしかありません。
      if (isSheetShifted_(sheet, spec)) return;

      const readCell = readHeaderRow_(sheet, spec).cell;
      for (let i = 0; i < spec.header.length; i++) {
        const now = readCell(i);
        if (now === spec.header[i]) continue;
        if (now !== '') continue;                       // 別の言葉が入っている → 触らない
        if (columnHasValue_(sheet, i + 1)) continue;    // データがある → 触らない
        sheet.getRange(1, i + 1)
          .setValue(spec.header[i])
          .setBackground('#e8eaed')
          .setFontWeight('bold');
        fixed.push('「' + spec.name + '」の' + (i + 1) + '列目に見出し「' + spec.header[i] + '」を書き足しました');
      }
    });
  } finally {
    lock.releaseLock();
  }

  return { fixed: fixed, left: checkSheets_(ss) };
}

/**
 * このアプリがデータを読み書きする表計算ファイルを返します。
 *
 * ■ いまの配り方（コンテナバインド）
 *   スプレッドシートのコピーを配り、そのファイルにこのスクリプトが束ねられています。
 *   束ねられているファイルがそのままデータベースなので、IDの控えも自動生成も要りません。
 *   先生は「どこにできたのか」を探さなくてよく、いま開いているそのファイルが中身です。
 *
 * ■ 前の配り方（独立スクリプト）ですでに公開している学級
 *   script.new で作った独立スクリプトには束ねられたファイルが無く、
 *   getActiveSpreadsheet() は null を返します。その学級ではこれまでどおり
 *   スクリプトプロパティ DB_ID の表計算ファイルを開きます。
 *   **ここを消すと、すでに使っている学級の記録が見えなくなります。**
 *
 * 【自己修復について】
 * 以前はここで「開けなかったら新しく作り直す」ことをしていました。やめました。
 * その形は、権限が足りない人が1回開いただけで、学級全員の記録が入ったファイルから
 * 空のファイルへ静かに差し替わります。画面には何も出ず「記録が消えた」ようにしか
 * 見えません。**開けないときは、作り直さずにその理由を出して止まります。**
 */
function getDB() {
  const bound = getBoundSpreadsheet_();
  if (bound) return ensureSheets_(bound);

  const dbId = SCRIPT_PROP.getProperty('DB_ID');
  if (dbId) {
    try {
      return ensureSheets_(SpreadsheetApp.openById(dbId));
    } catch (e) {
      console.error('DB_ID のスプレッドシートを開けませんでした。', e);
      throw new Error(
        'データを入れている表計算ファイルを開けませんでした。'
        + 'ファイルが削除されたか、開く権限がありません。'
        + '（作り直すと、これまでの記録が見えなくなるため、自動では作りません。'
        + '先生の方は、スクリプトのプロパティ DB_ID をご確認ください）'
      );
    }
  }

  throw new Error(
    'データを入れる表計算ファイルが見つかりません。'
    + 'このアプリは、配布されたスプレッドシートのコピーの中で動かしてください。'
    + '（スプレッドシートを開き、「拡張機能」＞「Apps Script」からデプロイした URL をお使いください）'
  );
}

/**
 * このスクリプトが束ねられているスプレッドシートを返します。
 * 独立スクリプトとして動いている場合は null を返します。
 */
function getBoundSpreadsheet_() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet() || null;
  } catch (e) {
    return null;   // 独立スクリプトでは例外になる版があります
  }
}

// -----------------------------------------------------------------
// 3b. スプレッドシートのメニュー（コンテナバインドのときだけ出ます）
// -----------------------------------------------------------------

/**
 * スプレッドシートを開いたときに、上に「こころスコープ」メニューを作ります。
 * ウェブアプリとして動いているときは画面が無いので、何もしません。
 */
function onOpen(e) {
  try {
    SpreadsheetApp.getUi()
      .createMenu(APP_NAME)
      .addItem('シートを点検する', 'showSheetCheck')
      .addItem('直せるところを直す', 'showSheetRepair')
      .addToUi();
  } catch (err) {
    // ウェブアプリ文脈では画面が無い。ここに来て構わない。
  }
}

/**
 * メニュー「シートを点検する」から呼びます。**何も書き換えません。**
 *
 * ⚠️ google.script.run は末尾 `_` の無い関数を誰でも直接呼べます。この関数も
 *    児童から呼べてしまうので、**1行目で getUi() を取ります**。
 *    ウェブアプリ文脈では画面が無いためここで例外になり、
 *    シートを1枚も読まずに終わります。
 *    なお返す内容は見出しの並びだけで、児童の記述や氏名は含みません。
 */
function showSheetCheck() {
  const ui = SpreadsheetApp.getUi();   // 画面が無ければ、ここで止まります
  const found = checkSheets_(getDB());
  ui.alert('シートの点検', describeFindings_(found), ui.ButtonSet.OK);
}

/**
 * メニュー「直せるところを直す」から呼びます。
 * 安全に直せるものだけを直し、残りは人に任せます。
 *
 * ⚠️ showSheetCheck と同じ理由で、1行目で getUi() を取ります。
 *    書き換えを伴うので、実行前に必ず確認を取ります。
 */
function showSheetRepair() {
  const ui = SpreadsheetApp.getUi();   // 画面が無ければ、ここで止まります
  const ss = getDB();

  const before = checkSheets_(ss);
  const fixable = before.filter(function (f) { return f.fixable; });
  if (!before.length) {
    ui.alert('シートの点検', 'シートの作りは想定どおりです。直すところはありません。', ui.ButtonSet.OK);
    return;
  }
  if (!fixable.length) {
    ui.alert(
      'シートの点検',
      describeFindings_(before) + '\n\nこれらは自動では直せません。上の説明のとおりに、手で直してください。',
      ui.ButtonSet.OK
    );
    return;
  }

  const answer = ui.alert(
    'シートを直します',
    '次のところを直します。\n\n'
      + fixable.map(function (f) { return '・「' + f.sheet + '」' + f.detail; }).join('\n')
      + '\n\n直してよろしいですか。',
    ui.ButtonSet.OK_CANCEL
  );
  if (answer !== ui.Button.OK) return;

  const result = repairSheets_(ss);
  const text = (result.fixed.length ? '直しました。\n\n' + result.fixed.map(function (m) { return '・' + m; }).join('\n') : '直せるところはありませんでした。')
    + '\n\n' + describeFindings_(result.left);
  ui.alert('シートの点検', text, ui.ButtonSet.OK);
}

/** 点検の結果を、先生に読める日本語にします。 */
function describeFindings_(found) {
  if (!found.length) return 'シートの作りは想定どおりです。';
  return '次のところが、アプリの想定と違います。\n\n'
    + found.map(function (f) { return '・「' + f.sheet + '」' + f.kind + '：' + f.detail; }).join('\n')
    + '\n\n列の並びは変えないでください。'
    + '児童ID・数値・削除の印は、列の「番号」で読み書きしています。'
    + '列を挿したり並べ替えたりすると、散布図がずれたり、'
    + '消したはずの記録が戻ったりします。';
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

    // アクセス中のユーザーのメールアドレスと、先生かどうかを取得
    const callerEmail = getCallerEmail_();
    const isTeacher = isTeacher_();

    // 名簿取得
    // 【重要】名簿の全件（氏名・ふりがな・メール）は先生にだけ返します。
    // 児童には、自分の情報（自動ログイン用）と授業情報だけを返します。
    const userSheet = ss.getSheetByName('名簿');
    let users = [];
    let autoLoginStudent = null;
    if (userSheet && userSheet.getLastRow() > 1) {
      const rows = userSheet.getDataRange().getValues().slice(1)
        .filter(r => !r[3]); // deletedAtがないもの

      if (isTeacher) {
        users = rows.map(r => ({ id: r[0], name: r[1], ruby: r[2], email: r[4] || '' }));
      }

      // メールアドレスで自動照合（自分の行だけを返すので、児童に返しても問題ありません）
      if (callerEmail) {
        const matched = rows.find(r => r[4] && String(r[4]).toLowerCase().trim() === callerEmail);
        if (matched) {
          autoLoginStudent = { id: matched[0], name: matched[1], ruby: matched[2] };
        }
      }
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

    return {
      success: true,
      isTeacher: isTeacher,
      users: users,
      rosterHidden: !isTeacher, // 児童には名簿を返していないことを画面側に伝えます
      activeSession: activeSession,
      autoLoginStudent: autoLoginStudent,
      callerEmail: callerEmail
    };

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
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
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
function addStudent(name, ruby, email) {
  assertTeacher_(); // 先生専用
  if (!validateStudentName(name)) {
    return { success: false, error: '名前は1〜50文字で入力してください' };
  }
  if (email && !validateEmail(email)) {
    return { success: false, error: 'メールアドレスの形式が正しくありません' };
  }
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  const studentId = 's' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
  sheet.appendRow([studentId, sanitizeText(name).trim(), sanitizeText(ruby).trim(), '', (email || '').trim()]);
  return { success: true, id: studentId };
}

/**
 * 名簿に生徒を一括追加します（CSV/Excel形式の貼り付け対応）
 * @param {Array} students - {name, ruby} の配列
 */
function addStudentBulk(students) {
  assertTeacher_(); // 先生専用
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  if (!sheet) return { success: false, error: '名簿シートが見つかりません' };

  const rows = students.map((s, i) => {
    // ユニークID生成 (タイムスタンプ + インデックスで重複回避)
    const suffix = ('000' + i).slice(-3);
    const studentId = 's' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss') + suffix;
    return [studentId, s.name, s.ruby, '', s.email || ''];
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
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
  const ss = getDB();
  const userSheet = ss.getSheetByName('名簿');
  if (!userSheet || userSheet.getLastRow() <= 1) return { success: true, users: [] };

  const users = userSheet.getDataRange().getValues().slice(1)
    .filter(r => !r[3]) // deletedAt
    .map(r => ({ id: r[0], name: r[1], ruby: r[2], email: r[4] || '' }));

  return { success: true, users: users };
}

/**
 * 児童のメールアドレスを更新します（教師用）
 */
function updateStudentEmail(studentId, email) {
  assertTeacher_(); // 先生専用
  const ss = getDB();
  const sheet = ss.getSheetByName('名簿');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === studentId && !data[i][3]) {
      sheet.getRange(i + 1, 5).setValue(email || '');
      return { success: true };
    }
  }
  return { success: false, error: '生徒が見つかりません' };
}

// =================================================================
// 6. ログ・分析 (LOGGING & ANALYTICS)
// =================================================================

/**
 * 生徒からの回答ログを保存します
 * サーバーサイドでメールアドレス検証を行い、なりすましを防止します
 */
function submitLog(data) {
  const ss = getDB();

  // メールアドレスによる本人確認
  let callerEmail = '';
  try {
    callerEmail = Session.getActiveUser().getEmail() || '';
  } catch (e) { }

  if (callerEmail) {
    const userSheet = ss.getSheetByName('名簿');
    if (userSheet && userSheet.getLastRow() > 1) {
      const rows = userSheet.getDataRange().getValues().slice(1);
      const student = rows.find(r => r[0] === data.studentId && !r[3]);
      if (student && student[4] && String(student[4]).toLowerCase().trim() !== callerEmail.toLowerCase().trim()) {
        return { success: false, error: '認証エラー: 自分のアカウントでログインしてください' };
      }
    }
  }

  // 入力値バリデーション
  if (!data.sessionId || !data.studentId || !data.phase) {
    return { success: false, error: '必須パラメータが不足しています' };
  }
  const safeText = sanitizeText(data.text);
  const safeValue = String(data.value || '').substring(0, MAX_VALUE_LENGTH);

  const logSheet = ss.getSheetByName('記録');
  if (!logSheet) {
    return { success: false, error: '「記録」シートがありません。先生にお知らせください。' };
  }

  // 【なぜロックで囲むのか】
  // 道徳の授業では、40人が「送信」をほぼ同時に押します。
  // ロックが無いと、重複の確認をした人と、行を足す人が入れ替わります。
  //   児童A: 重複を確認 → 無い
  //   児童A（もう一度押した）: 重複を確認 → まだ無い（1回目がまだ書けていない）
  //   → 同じ児童の BEFORE が2行入る
  // こうなると散布図が二重点になり、変容の集計が壊れます。
  // 「すでに送信済みです」も出ないので、**誰も気づきません。**
  //
  // 囲む範囲は「重複の確認 → 1行足す」だけに絞ります。
  // 名簿の照合や入力の検査は上で済ませてあり、ロックの外です。
  const logId = Utilities.getUuid();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { success: false, error: 'いま混み合っています。少し待ってから、もう一度送信してください。' };
  }
  try {
    if (logSheet.getLastRow() > 1) {
      const existing = logSheet.getDataRange().getValues().slice(1)
        .find(r => r[1] === data.sessionId && r[2] === data.studentId && r[3] === data.phase && !r[7]);
      if (existing) {
        return { success: false, error: 'すでにこのフェーズで送信済みです', duplicate: true };
      }
    }

    logSheet.appendRow([
      logId,
      data.sessionId,
      data.studentId,
      data.phase,
      safeValue,
      safeText,
      new Date(),
      ''
    ]);
    SpreadsheetApp.flush();   // ロックを離す前に、確実に書き込んでおく
  } finally {
    lock.releaseLock();
  }

  return { success: true, logId: logId };
}

/**
 * 教師用ダッシュボード向けの授業ログ取得（散布図表示用）
 */
function getSessionLogs(sessionId) {
  assertTeacher_(); // 先生専用
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
  // 先生、または本人だけが見られます（児童が自分の記録を見る画面でも使うため）
  assertTeacherOrSelf_(studentId);
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

  // 【この関数について】
  // これは児童も呼ぶ機能です（入力画面の「みんなの考えを見る」）。
  // 名前を伏せた記述を並べて、教室で読み合うために作ってあります。
  //
  // ⚠️ 以前は sessionId を受け取ったまま何も確かめずに引いていました。
  //    google.script.run は誰でも直接呼べるので、児童がコンソールから
  //    **過去のどの授業の記述本文でも**引けてしまいます。
  //    道徳の記述には、家庭のことや友だちとのもめごとが書かれます。
  //    卒業した学年のぶんまで、名前は伏せられていても本文は全部読めました。
  //
  //    いまは「いま開いている授業」に限ります。先生は授業の一覧から
  //    どれでも見られます（getSessionLogs / getStudentSummaries は
  //    assertTeacher_ を通ります）。
  assertActiveSessionOrTeacher_(ss, sessionId);

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
  assertTeacher_(); // 先生専用
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
  // 先生、または本人だけが見られます
  assertTeacherOrSelf_(studentId);
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
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
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
    // 通信・再試行・応答の取り出しは正本 Gemini.gs（GigaGemini）に任せる。
    // API キーは正本側で x-goog-api-key ヘッダに載る（URL クエリには入れない）。
    const question = GigaGemini.call({
      apiKey: apiKey,
      prompt: prompt,
      generationConfig: { maxOutputTokens: 100, temperature: 0.7 }
    });
    return { success: true, question: question };
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
  assertTeacher_(); // 先生専用
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

    // 通信・再試行・コードフェンスの除去つき JSON 解釈は正本 Gemini.gs に任せる。
    // PDF は parts に inline_data として混ぜる（正本の buildBody が req.parts を尊重する）。
    const result = GigaGemini.callJson({
      apiKey: apiKey,
      parts: [
        { text: prompt },
        { inline_data: { mime_type: 'application/pdf', data: base64Data } }
      ],
      generationConfig: { response_mime_type: 'application/json' }
    });
    return { success: true, data: result };

  } catch (e) {
    console.error(e);
    return { success: false, error: 'PDFの解析に失敗: ' + e.toString() };
  }
}

/**
 * Gemini APIキーの保存
 */
function saveGeminiApiKey(apiKey) {
  assertTeacher_(); // 先生専用
  SCRIPT_PROP.setProperty('GEMINI_API_KEY', apiKey || '');
  return { success: true };
}

/**
 * 権限認証を強制するためのダミー関数
 * エディタ上でこの関数を選択して実行すると、必要な権限の承認画面が表示されます。
 * あわせて、実行した人のメールアドレスを OWNER_EMAIL に記録します。
 */
/**
 * 必要な権限の同意画面を、先生にまとめて出させるための関数です。
 *
 * ⚠️ 末尾に `_` が付いているのは、google.script.run から呼ばせないためです。
 *    以前は `forceAuth`（`_` 無し）で、しかも1行目が rememberOwnerEmail_() でした。
 *    **児童がブラウザのコンソールから `runGoogleScript('forceAuth')` と打つだけで、
 *    その児童が恒久的に先生になれました。**
 *    manualSetup と同じ穴が、こちらにだけ残っていました。
 *    OWNER_EMAIL を記録するのは manualSetup の仕事なので、ここではしません。
 */
function forceAuth_() {
  SpreadsheetApp.getActiveSpreadsheet();
  Session.getActiveUser().getEmail();
  UrlFetchApp.fetch("https://www.google.com");
  console.log("Auth complete");
}

/**
 * Gemini APIキーの有無確認
 */
function hasGeminiApiKey() {
  assertTeacher_(); // 先生専用
  const key = SCRIPT_PROP.getProperty('GEMINI_API_KEY');
  return { hasKey: !!(key && key.length > 0) };
}

// =================================================================
// 9. その他ユーティリティ (UTILITIES)
// =================================================================

/**
 * 授業データをCSV形式で取得（教師用エクスポート）
 */
function exportSessionCsv(sessionId) {
  assertTeacher_(); // 先生専用
  const ss = getDB();

  // 授業情報
  const sessionSheet = ss.getSheetByName('授業');
  const sessionRow = sessionSheet.getDataRange().getValues().slice(1).find(r => r[0] === sessionId);
  if (!sessionRow) return { success: false, error: '授業が見つかりません' };
  const sessionTitle = sessionRow[2];
  const inputType = sessionRow[3];

  // 名簿マップ
  const userSheet = ss.getSheetByName('名簿');
  const users = {};
  if (userSheet && userSheet.getLastRow() > 1) {
    userSheet.getDataRange().getValues().slice(1)
      .filter(r => !r[3])
      .forEach(r => { users[r[0]] = { name: r[1], ruby: r[2] }; });
  }

  // ログ取得
  const logSheet = ss.getSheetByName('記録');
  const logs = logSheet.getDataRange().getValues().slice(1)
    .filter(r => r[1] === sessionId && !r[7]);

  // 児童ごとにグルーピング
  const grouped = {};
  logs.forEach(r => {
    const sid = r[2];
    if (!grouped[sid]) grouped[sid] = {};
    grouped[sid][r[3]] = { value: r[4], text: r[5] };
  });

  // CSV組み立て
  const rows = [['名前', 'ふりがな', '導入_値', '導入_理由', '振り返り_値', '振り返り_理由']];
  Object.keys(grouped).forEach(sid => {
    const u = users[sid] || { name: sid, ruby: '' };
    const b = grouped[sid]['BEFORE'] || {};
    const a = grouped[sid]['AFTER'] || {};
    rows.push([
      u.name, u.ruby,
      String(b.value || ''), String(b.text || ''),
      String(a.value || ''), String(a.text || '')
    ]);
  });

  const csv = rows.map(r => r.map(c => '"' + String(c).replace(/"/g, '""') + '"').join(',')).join('\n');
  return { success: true, csv: csv, title: sessionTitle, inputType: inputType };
}

/**
 * 全授業セッション一覧を取得（履歴表示用）
 */
function getAllSessions() {
  assertTeacher_(); // 先生専用
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
  assertTeacher_(); // 先生専用
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

  // 【重要】外部AI（Gemini）には氏名・ふりがな・メールアドレスを送りません。
  // 児童は「この児童」という仮名で伝え、実名への差し戻しは画面側（js.html）で行います。
  const prompt = `あなたはベテランの小学校教師です。道徳の授業における児童の学習記録を分析し、通知表に記載する「所見」を作成してください。

【対象児童】この児童

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
- 児童を指すときは、かならず「この児童」と書く（あとで実名に置き換えるため）

所見文のみを出力してください（説明や前置き不要）。`;

  try {
    // 通信・再試行・応答の取り出しは正本 Gemini.gs（GigaGemini）に任せる。
    const observation = GigaGemini.call({
      apiKey: apiKey,
      prompt: prompt,
      generationConfig: { maxOutputTokens: 500, temperature: 0.5 }
    });
    return { success: true, observation: observation, name: student[1], ruby: student[2] };
  } catch (e) {
    console.error('Gemini API Error:', e);
    return { success: false, error: e.toString() };
  }
}

/**
 * 教師用パスワードの照合（試行回数制限付き）
 *
 * 【重要】この関数はセキュリティ境界ではありません。
 * ここで行っているのは「先生用の画面に切り替えるか」という表示上の判定だけです。
 * データを守っている本当の境界は、各関数の冒頭で呼んでいる assertTeacher_()
 * （＝ログイン中のGoogleアカウントのメールアドレス照合）です。
 * そのため、パスワードを知っていても先生として登録されていない人は
 * 名簿や記録を1件も取得できません。
 */
function checkTeacherPassword(password) {
  const cache = CacheService.getScriptCache();
  const lockKey = 'TEACHER_LOGIN_LOCK';
  const attemptKey = 'TEACHER_LOGIN_ATTEMPTS';

  // ロックアウト確認
  const locked = cache.get(lockKey);
  if (locked) {
    return { success: false, error: 'ログイン試行回数の上限に達しました。しばらく待ってから再試行してください。', locked: true };
  }

  // パスワードはハッシュ値だけを保存しています（平文比較はしません）
  const storedHash = SCRIPT_PROP.getProperty('TEACHER_PASSWORD_HASH');
  if (!storedHash) {
    // 既定パスワード（admin）は廃止したので、未設定のときは誰も通しません
    return {
      success: false,
      needsSetup: true,
      error: 'まず管理者パスワードを設定してください。'
    };
  }

  if (hashPassword_(password) === storedHash) {
    // 成功: 試行カウントをリセット
    cache.remove(attemptKey);
    return { success: true, requirePasswordChange: false };
  }

  // 失敗: 試行回数をインクリメント
  let attempts = parseInt(cache.get(attemptKey) || '0') + 1;
  if (attempts >= MAX_LOGIN_ATTEMPTS) {
    cache.put(lockKey, 'true', LOCKOUT_DURATION_MIN * 60);
    cache.remove(attemptKey);
    return { success: false, error: `${MAX_LOGIN_ATTEMPTS}回連続で失敗しました。${LOCKOUT_DURATION_MIN}分間ロックされます。`, locked: true };
  }
  cache.put(attemptKey, String(attempts), 600); // 10分間保持
  return { success: false, error: `パスワードが違います（残り${MAX_LOGIN_ATTEMPTS - attempts}回）` };
}

/**
 * 教師用パスワードの設定・変更
 * パスワードはSHA-256のハッシュ値にしてから保存します（平文では保存しません）。
 * ※パスワードそのものはセキュリティ境界ではないため、本当の確認は assertTeacher_() で行います。
 */
function changeTeacherPassword(currentPass, newPass) {
  assertTeacher_(); // 先生専用（ここが本当の入り口の鍵です）

  const storedHash = SCRIPT_PROP.getProperty('TEACHER_PASSWORD_HASH');
  // 未設定のときは「初回設定」なので、現在のパスワードは不要です
  if (storedHash && hashPassword_(currentPass) !== storedHash) {
    return { success: false, error: '現在のパスワードが違います' };
  }
  if (!newPass || String(newPass).length < 4) {
    return { success: false, error: 'パスワードは4文字以上にしてください' };
  }
  SCRIPT_PROP.setProperty('TEACHER_PASSWORD_HASH', hashPassword_(newPass));
  // 以前のバージョンで平文保存されていたパスワードが残っていれば消します
  SCRIPT_PROP.deleteProperty('TEACHER_PASSWORD');
  return { success: true };
}

/**
 * 管理者パスワードが設定済みかどうかを返します（画面の出し分け用）
 */
function hasTeacherPassword() {
  const storedHash = SCRIPT_PROP.getProperty('TEACHER_PASSWORD_HASH');
  return { hasPassword: !!storedHash };
}


/**
 * 【この関数について】
 * どこからも呼ばれていませんが、末尾に `_` を付けて非公開にしてあります。
 * google.script.run は末尾 `_` の無い関数を誰でも直接呼べるためです。
 */
function getParentFolderId_(folder) {
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
 * スクリプトエディタでこの関数を1回実行してください。
 * 実行した人のメールアドレスが OWNER_EMAIL に記録され、その人が先生として扱われます。
 * （先生が複数いる場合は、スクリプトプロパティ TEACHER_EMAILS にカンマ区切りで登録してください）
 */
function manualSetup() {
  // ⚠️ この関数は末尾に `_` が無いので、google.script.run から誰でも呼べます。
  //    以前はここが「まだ OWNER_EMAIL が無ければ、呼んだ人を先生にする」形でした。
  //    スプレッドシートのコピーで配る今の形では、先生の手元に届くコピーは
  //    毎回 OWNER_EMAIL が空です。先生がデプロイして URL を配ったあと、
  //    先生より先に児童がこれを呼ぶと、**その児童が恒久的に先生になります**
  //    （取り消すにはスクリプトのプロパティを手で書き換えるしかありません）。
  //
  //    そこで「呼んだ人が、このアプリを公開した本人であること」を先に確かめます。
  //    スクリプトエディタから先生が実行したときは、両方とも先生自身になるので通ります。
  //    児童のブラウザから呼ばれたときは、実行ユーザー（先生）と
  //    アクセスユーザー（児童）が食い違うので、ここで止まります。
  const caller = getCallerEmail_();
  const deployer = getDeployerEmail_();
  if (!caller || !deployer || caller !== deployer) {
    throw new Error('この操作は、このアプリを公開したご本人だけが実行できます。'
      + 'スクリプトエディタから manualSetup を実行してください。');
  }

  try {
    const owner = rememberOwnerEmail_();

    // 必要な権限の同意を、ここでまとめて出します。
    // 校内から外部への通信が塞がれていると最後の1つで失敗しますが、
    // 生成AIを使わなければ困らないので、止めずに知らせるだけにします。
    try {
      forceAuth_();
    } catch (authError) {
      console.warn('⚠ 外部への通信を確かめられませんでした（生成AIを使わないなら問題ありません）:', authError);
    }

    const ss = getDB();
    console.log("✅ セットアップ完了！ 使用中のスプレッドシート:", ss.getName());
    console.log("📄 このスクリプトが束ねられているファイルか:", getBoundSpreadsheet_() ? "はい（コンテナバインド）" : "いいえ（独立スクリプト）");

    const found = checkSheets_(ss);
    if (found.length) {
      console.warn("⚠ シートの作りが想定と違います:");
      found.forEach(function (f) { console.warn('  ・「' + f.sheet + '」' + f.kind + '：' + f.detail); });
      console.warn("  スプレッドシートの「" + APP_NAME + "」メニュー＞「直せるところを直す」でお試しください。");
    } else {
      console.log("✅ シートの作りは想定どおりです。");
    }

    if (owner) {
      console.log("👤 管理者(OWNER_EMAIL)として登録されました:", owner);
    } else {
      console.warn("⚠ メールアドレスを取得できませんでした。スクリプトプロパティ TEACHER_EMAILS に先生のメールアドレスを登録してください。");
    }
  } catch(e) {
    console.error("セットアップ失敗:", e);
    throw e;   // 画面にも出す（以前は握りつぶしていたので、失敗が成功に見えていました）
  }
}
