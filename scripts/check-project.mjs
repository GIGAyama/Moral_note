#!/usr/bin/env node
/**
 * このリポジトリの最低限の点検。
 *
 *   node scripts/check-project.mjs
 *
 * ── なぜ要るか ──────────────────────────────────
 *
 * このリポジトリは main に入ると deploy.yml が clasp push して、既存の
 * デプロイを差し替えます。**数分で教室に届きます。**
 * その deploy.yml は package.json の quality / ci / check を順に探し、
 * どれも無ければ「飛ばします」と表示して**緑のまま本番へ push します**。
 * これまでこのリポジトリはどれも持っていなかったので、ゲートは1度も
 * 走っていませんでした。この文書が、その穴をふさぐためのものです。
 *
 * ここで見るのは、手元で確かめられることだけです。本番への疎通や
 * 見た目は見られないので、見たふりをしません。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');
const exists = (f) => fs.existsSync(path.join(ROOT, f));

const problems = [];
const notes = [];
const pending = [];   // 人の手が要るもの（落とさずに、毎回目立たせる）
const fail = (id, msg) => problems.push(`${id}: ${msg}`);
const ok = (id, msg) => notes.push(`✅ ${id} — ${msg}`);

const CODE = read('code.gs');

// -----------------------------------------------------------------
// 1. GAS が読むファイルが、ちゃんと GAS へ送られるか
// -----------------------------------------------------------------
// .claspignore は「まず全部を除外してから、必要なものだけ戻す」書き方です。
// 戻し忘れると、GAS 側でそのファイルだけが古いまま（または無いまま）になり、
// 画面が白くなります。押してみるまで気づけないので、ここで見ます。
{
  const claspignore = read('.claspignore');
  const allowed = claspignore
    .split('\n')
    .filter((l) => l.startsWith('!'))
    .map((l) => l.slice(1).trim());

  const shipped = (f) => allowed.includes(f) || (allowed.includes('*.gs') && f.endsWith('.gs'));

  // code.gs が名前で読んでいる .html を全部集めます
  const needed = new Set();
  for (const m of CODE.matchAll(/createTemplateFromFile\(\s*['"]([^'"]+)['"]/g)) needed.add(m[1]);
  for (const m of CODE.matchAll(/createHtmlOutputFromFile\(\s*['"]([^'"]+)['"]/g)) needed.add(m[1]);

  // 外枠の .html が include している名前も集めます
  for (const shell of [...needed]) {
    if (!exists(`${shell}.html`)) continue;
    for (const m of read(`${shell}.html`).matchAll(/include\(\s*['"]([^'"]+)['"]/g)) needed.add(m[1]);
  }

  for (const name of needed) {
    if (!exists(`${name}.html`)) {
      fail('E_HTML_MISSING', `code.gs が読む ${name}.html がリポジトリにありません`);
    } else if (!shipped(`${name}.html`)) {
      fail('E_HTML_NOT_SHIPPED',
        `${name}.html が .claspignore で GAS へ送られません（"!${name}.html" の行が要ります）`);
    }
  }
  if (needed.size) ok('E_HTML_SHIPPED', `GAS が読む ${needed.size} 枚の .html はすべて送られます`);

  // 逆に、GitHub Pages に出すページを GAS へ送ってはいけません。
  // index.html を送ると GAS 側の画面と取り違えます。
  if (shipped('index.html')) {
    fail('E_LANDING_SHIPPED',
      'index.html が GAS へ送られます。あれは moral-note.giga-school.com の案内ページです');
  } else {
    ok('E_LANDING_SHIPPED', '案内ページ index.html は GAS へ送られません');
  }

  // 逆向きの取り違え: GAS の外枠を index.html という名前で置くと、
  // GitHub Pages がテンプレートをそのまま配って白い画面になります。
  if (needed.has('index')) {
    fail('E_SHELL_IS_INDEX',
      'GAS の外枠が index.html です。CNAME があるので Pages がテンプレートをそのまま配ります（app-shell.html にしてください）');
  } else {
    ok('E_SHELL_IS_INDEX', 'GAS の外枠は index.html ではありません');
  }
}

// -----------------------------------------------------------------
// 2. 公開エンドポイント（google.script.run から誰でも呼べる関数）
// -----------------------------------------------------------------
// 末尾に `_` の無いトップレベル関数は、児童のブラウザから直接呼べます。
// 増えたときに気づけるよう、一覧を固定しておきます。
{
  const KNOWN = new Set([
    // 画面
    'doGet', 'include', 'onOpen',
    // スプレッドシートのメニュー（1行目で getUi() を取るので、画面が無ければ止まります）
    'showSheetCheck', 'showSheetRepair',
    // 入力の検査（副作用なし）
    'sanitizeText', 'validateStudentName', 'validateEmail',
    // データベース
    'getDB',
    // 授業
    'getInitialData', 'getPollingData', 'updateSessionPhase', 'closeSession',
    'createSession', 'getAllSessions',
    // 名簿
    'addStudent', 'addStudentBulk', 'deleteStudent', 'getStudents', 'updateStudentEmail',
    // 記録
    'submitLog', 'getSessionLogs', 'getStudentPortfolio', 'getAnonymousOpinions',
    'getStudentSummaries', 'getStudentLogs', 'exportSessionCsv',
    // 単元
    'saveUnit', 'getUnits', 'startSessionFromUnit', 'deleteUnit',
    // AI
    'generateSocraticQuestion', 'parseLessonPdf', 'generateObservation',
    'saveGeminiApiKey', 'hasGeminiApiKey',
    // 合言葉
    'checkTeacherPassword', 'changeTeacherPassword', 'hasTeacherPassword',
    // セットアップ（公開した本人だけが通ります）
    'manualSetup',
  ]);

  const actual = [...CODE.matchAll(/^function ([A-Za-z][A-Za-z0-9]*)\s*\(/gm)].map((m) => m[1]);
  const added = actual.filter((n) => !KNOWN.has(n));
  const removed = [...KNOWN].filter((n) => !actual.includes(n));

  if (added.length) {
    fail('E_NEW_PUBLIC_ENDPOINT',
      `公開関数が増えています: ${added.join(', ')}\n`
      + '      google.script.run は末尾 `_` の無い関数を誰でも呼べます。\n'
      + '      内部で使うだけなら名前の末尾に `_` を付けてください。\n'
      + '      外から呼ぶ必要があるなら、1行目で認可（assertTeacher_ など）を通してから\n'
      + '      scripts/check-project.mjs の KNOWN に足してください。');
  }
  if (removed.length) {
    fail('E_PUBLIC_ENDPOINT_GONE',
      `一覧にある公開関数が見つかりません: ${removed.join(', ')}（画面から呼んでいれば壊れます）`);
  }
  if (!added.length && !removed.length) {
    ok('E_NEW_PUBLIC_ENDPOINT', `公開エンドポイントは ${actual.length} 件で、一覧どおりです`);
  }
}

// -----------------------------------------------------------------
// 3. 自己修復でデータベースを作り直していないか
// -----------------------------------------------------------------
// 「開けなかったら新しく作る」は、権限の足りない人が1回開くだけで
// 学級全員の記録が空のファイルへ静かに差し替わります。
{
  if (/SpreadsheetApp\.create\s*\(/.test(CODE)) {
    fail('E_DB_SELF_HEAL',
      'code.gs に SpreadsheetApp.create があります。開けないときに作り直すと、'
      + '学級の記録が入ったファイルから空のファイルへ黙って差し替わります');
  } else {
    ok('E_DB_SELF_HEAL', 'データベースを自動で作り直しません');
  }

  if (!/getActiveSpreadsheet\s*\(/.test(CODE)) {
    fail('E_NOT_CONTAINER_BOUND',
      'code.gs が getActiveSpreadsheet を使っていません。'
      + 'コンテナバインド（スプレッドシートのコピーで配る形）で動きません');
  } else {
    ok('E_NOT_CONTAINER_BOUND', '束ねられたスプレッドシートを見ています');
  }
}

// -----------------------------------------------------------------
// 4. デプロイの設定（appsscript.json）
// -----------------------------------------------------------------
// clasp push は GAS 側のマニフェストを丸ごと上書きします。webapp が欠けたまま
// 送るとウェブアプリの入口が消えます。oauthScopes が無いと GAS が保存のたびに
// スコープを推測し、同意画面が広がって既存の承認が無効になります。
{
  const manifest = JSON.parse(read('appsscript.json'));

  if (!manifest.webapp || !manifest.webapp.executeAs || !manifest.webapp.access) {
    fail('E_WEBAPP_MISSING', 'appsscript.json に webapp の executeAs / access がありません');
  } else if (manifest.webapp.executeAs !== 'USER_DEPLOYING') {
    fail('E_EXECUTE_AS',
      `executeAs が ${manifest.webapp.executeAs} です。USER_ACCESSING にすると児童にも`
      + 'スプレッドシートの権限が要り、シートを直接開けば全員分の記録が読めます');
  } else {
    ok('E_WEBAPP_MISSING', `webapp は ${manifest.webapp.executeAs} / ${manifest.webapp.access} です`);
  }

  const scopes = manifest.oauthScopes || [];
  if (!scopes.length) {
    fail('E_SCOPES_MISSING', 'appsscript.json に oauthScopes がありません');
  }
  // 使っている機能に対して、宣言したスコープが足りているか。
  // oauthScopes を明記していると GAS は勝手に足さないので、足りないものは
  // 「保存も反映も通るのに、押した瞬間だけ失敗する」形で出ます。
  const NEEDED = [
    { when: /SpreadsheetApp\.getUi\s*\(/, scope: 'https://www.googleapis.com/auth/script.container.ui',
      why: 'SpreadsheetApp.getUi()（スプレッドシートのメニュー）を使っています' },
    { when: /UrlFetchApp\.fetch/, scope: 'https://www.googleapis.com/auth/script.external_request',
      why: 'UrlFetchApp.fetch（外部への通信）を使っています' },
    { when: /SpreadsheetApp\./, scope: 'https://www.googleapis.com/auth/spreadsheets',
      why: 'SpreadsheetApp（スプレッドシート）を使っています' },
    { when: /Session\.get(Active|Effective)User/, scope: 'https://www.googleapis.com/auth/userinfo.email',
      why: 'Session でメールアドレスを見ています' },
  ];
  const allGs = fs.readdirSync(ROOT).filter((f) => f.endsWith('.gs')).map((f) => read(f)).join('\n');
  const lacking = NEEDED.filter((n) => n.when.test(allGs) && !scopes.includes(n.scope));
  if (lacking.length) {
    lacking.forEach((n) => fail('E_SCOPE_MISSING',
      `oauthScopes に ${n.scope} がありません（${n.why}）`));
  } else {
    ok('E_SCOPE_MISSING', '使っている機能に対して、宣言したスコープは足りています');
  }

  if (scopes.includes('https://www.googleapis.com/auth/drive')) {
    fail('E_SCOPE_TOO_WIDE',
      'oauthScopes にフルドライブ（auth/drive）があります。'
      + '保護者説明で「子どものドライブ全部を読めます」と言わざるを得なくなります');
  } else if (scopes.length) {
    ok('E_SCOPE_TOO_WIDE', `oauthScopes は ${scopes.length} 件で、フルドライブを含みません`);
  }
}

// -----------------------------------------------------------------
// 5. 案内ページ（GitHub Pages に出るもの）
// -----------------------------------------------------------------
{
  if (!exists('index.html')) {
    fail('E_LANDING_MISSING', 'index.html がありません。CNAME があるので、住所を開くと 404 になります');
  } else {
    const landing = read('index.html');
    if (/<\?!?=/.test(landing)) {
      fail('E_LANDING_IS_TEMPLATE',
        'index.html に GAS のテンプレート記法（<?!= ?>）があります。'
        + 'ブラウザは黙って捨てるので、ほぼ白い画面になります');
    } else {
      ok('E_LANDING_IS_TEMPLATE', 'index.html は素の HTML です');
    }
    if (!/<title>/.test(landing)) fail('E_LANDING_TITLE', 'index.html に <title> がありません');
    if (!/name=["']viewport["']/.test(landing)) fail('E_LANDING_VIEWPORT', 'index.html に viewport がありません');
    if (/user-scalable\s*=\s*no|maximum-scale\s*=\s*1/.test(landing)) {
      fail('E_LANDING_ZOOM', 'index.html が拡大を禁じています。文字を大きくできない児童・先生がいます');
    }
    if (!/<html[^>]+lang=/.test(landing)) fail('E_LANDING_LANG', 'index.html の <html> に lang がありません');
  }

  // 配っているコピーリンクが、案内ページと手順書で食い違っていないか。
  // 食い違うと、片方を見た先生だけが古いテンプレートをコピーします。
  const LINK_FILES = ['index.html', 'README.md', 'docs/copy-distribution.md'];
  const linkOf = (text) => [...text.matchAll(/spreadsheets\/d\/([A-Za-z0-9_-]{20,})\/copy/g)].map((m) => m[1]);
  const ids = new Set();
  for (const f of LINK_FILES) {
    if (exists(f)) linkOf(read(f)).forEach((id) => ids.add(id));
  }

  // テンプレートのスプレッドシートは、Google アカウントを持つ人にしか作れません。
  // リポジトリの側からは用意できないので、置き場所だけ __COPY_LINK__ で空けてあります。
  // 忘れられると案内ページのボタンが死んだリンクになるので、毎回大きく出します。
  // （まだ作られていないだけで、コードが壊れているわけではないので exit 0 のままにします）
  // docs/copy-distribution.md は「__COPY_LINK__ を差し替えてください」と
  // 説明している側なので、差し替え待ちの数には入れません。
  const waiting = LINK_FILES
    .filter((f) => f !== 'docs/copy-distribution.md')
    .filter((f) => exists(f) && read(f).includes('__COPY_LINK__'));
  if (waiting.length) {
    pending.push(
      'コピー用テンプレートのリンクが、まだ入っていません（__COPY_LINK__ のまま）:\n'
      + waiting.map((f) => `        ・${f}`).join('\n')
      + '\n      作り方と、差し替える場所は docs/copy-distribution.md にあります。'
    );
  }
  if (ids.size > 1) {
    fail('E_COPY_LINK_MISMATCH',
      `配っているコピーリンクが ${ids.size} 種類あります: ${[...ids].join(', ')}。`
      + 'テンプレートを作り直したときの差し替え漏れです');
  } else if (ids.size === 1) {
    ok('E_COPY_LINK_MISMATCH', 'コピーリンクは1種類でそろっています');
  }
}

// -----------------------------------------------------------------
// 結果
// -----------------------------------------------------------------
notes.forEach((n) => console.log(n));
if (pending.length) {
  console.log(`\n⏳ 人の手が要るものが ${pending.length} 件あります。\n`);
  pending.forEach((p) => console.log(`   ${p}`));
}
if (problems.length) {
  console.error(`\n❌ ${problems.length} 件の問題があります。\n`);
  problems.forEach((p) => console.error(`   ${p}`));
  console.error('');
  process.exit(1);
}
console.log(`\n✅ ${notes.length} 件すべて満たしました。`);
console.log('   ※ 本番への疎通・実機の見た目は、この環境では確かめられません（未確認）。');
