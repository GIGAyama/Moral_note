/**
 * code.gs のシート点検・修整のテスト。
 *
 * GAS のランタイムは手元で動かせないので、SpreadsheetApp / LockService などを
 * 偽物に差しかえ、code.gs をそのまま Node で実行して判断だけを検査します。
 * （正本 standards/gas/Gemini.test.mjs と同じ形。関数を正規表現で切り出す方式は
 *   書き方を少し変えただけで「読み取れませんでした」と落ちるので使いません。）
 *
 * ここで守りたいのは1つです。
 *
 *   **中身が入っている列の見出しを、勝手に書き換えないこと。**
 *
 * このアプリは記録シートを列の番号で読みます（r[2]=児童ID, r[4]=数値, r[7]=削除の印）。
 * 先生が列を1本挿すと見出しがずれます。そこで見出しだけを正しく直すと、
 * 「間違った列に正しいラベルが付いた」状態になり、散布図がずれていることに
 * 誰も気づけなくなります。直すのは人の仕事で、ここは見つけるだけにします。
 */
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(HERE, '..');
const SOURCE = fs.readFileSync(path.join(ROOT, 'code.gs'), 'utf8');

// ---------------------------------------------------------------
// 偽物のスプレッドシート
// ---------------------------------------------------------------

/** 2次元配列を「シート」に見せかけます。 */
function fakeSheet(name, rows) {
  const grid = rows.map((r) => [...r]);
  const width = () => grid.reduce((m, r) => Math.max(m, r.length), 0);
  const pad = (n) => grid.forEach((r) => { while (r.length < n) r.push(''); });

  const sheet = {
    _name: name,
    _grid: grid,
    getName: () => sheet._name,
    setName: (n) => { sheet._name = n; return sheet; },
    getLastRow: () => grid.length,
    getLastColumn: () => width(),
    // 本番のシートは既定で26列あります。中身の幅をそのまま返すと、
    // insertColumnsAfter の枝がテストでだけ走る（＝本番の姿を試していない）
    // ことになるので、26 を下限にします。
    getMaxColumns: () => Math.max(width(), 26),
    getDataRange: () => sheet.getRange(1, 1, Math.max(grid.length, 1), Math.max(width(), 1)),
    insertColumnsAfter: (after, howMany) => { pad(after + howMany); },
    appendRow: (row) => { grid.push([...row]); },
    getRange: (row, col, numRows = 1, numCols = 1) => ({
      getValues: () => {
        const out = [];
        for (let r = 0; r < numRows; r++) {
          const line = [];
          for (let c = 0; c < numCols; c++) {
            const src = grid[row - 1 + r];
            line.push(src === undefined ? '' : (src[col - 1 + c] ?? ''));
          }
          out.push(line);
        }
        return out;
      },
      setValues: (values) => {
        values.forEach((line, r) => {
          while (grid.length < row + r) grid.push([]);
          const target = grid[row - 1 + r];
          line.forEach((v, c) => { target[col - 1 + c] = v; });
        });
        return sheet.getRange(row, col, numRows, numCols);
      },
      setValue: (v) => sheet.getRange(row, col, 1, 1).setValues([[v]]),
      setBackground: () => sheet.getRange(row, col, numRows, numCols),
      setFontWeight: () => sheet.getRange(row, col, numRows, numCols),
    }),
  };
  return sheet;
}

/** シート名 -> 行の配列 から、偽物のスプレッドシートを作ります。 */
function fakeSpreadsheet(spec) {
  const sheets = Object.entries(spec).map(([name, rows]) => fakeSheet(name, rows));
  const ss = {
    _sheets: sheets,
    getName: () => 'テスト用',
    getId: () => 'test-id',
    getSheets: () => ss._sheets,
    getSheetByName: (n) => ss._sheets.find((s) => s.getName() === n) || null,
    insertSheet: (n) => { const s = fakeSheet(n, []); ss._sheets.push(s); return s; },
    deleteSheet: (s) => { ss._sheets = ss._sheets.filter((x) => x !== s); },
  };
  return ss;
}

/** code.gs を偽物の GAS の上で実行し、中の関数を取り出します。 */
function load({ bound = null, props = {}, activeUser = '', effectiveUser = '' } = {}) {
  const store = { ...props };
  const sandbox = {
    console: { log() {}, warn() {}, error() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => bound,
      openById: (id) => { throw new Error('openById は使わないはずです: ' + id); },
      create: () => { throw new Error('create は呼ばれてはいけません'); },
      getUi: () => { throw new Error('Exception: Cannot call SpreadsheetApp.getUi()'); },
      flush: () => {},
    },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (k) => (k in store ? store[k] : null),
        setProperty: (k, v) => { store[k] = v; },
      }),
    },
    LockService: {
      getScriptLock: () => ({ waitLock() {}, releaseLock() {} }),
    },
    Session: {
      getActiveUser: () => ({ getEmail: () => activeUser }),
      getEffectiveUser: () => ({ getEmail: () => effectiveUser }),
    },
    Utilities: {
      getUuid: () => 'uuid',
      computeDigest: () => [0],
      DigestAlgorithm: { SHA_256: 'SHA_256' },
      Charset: { UTF_8: 'UTF_8' },
    },
    HtmlService: { createTemplateFromFile: () => ({ evaluate: () => ({}) }) },
    UrlFetchApp: { fetch: () => { throw new Error('通信しません'); } },
  };
  vm.createContext(sandbox);
  vm.runInContext(SOURCE, sandbox);
  return { sandbox, store };
}

/**
 * vm の中で作られた配列やオブジェクトは別realmのプロトタイプを持つので、
 * strict な deepEqual が「構造は同じだが参照が違う」で落ちます。
 * 検査したいのは中身なので、こちら側の素の値に写してから比べます。
 */
function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

/** SHEETS_ のとおりに整った、まっとうなスプレッドシートを作ります。 */
function healthy(sandbox) {
  const spec = {};
  sandbox.SHEETS_.forEach((s) => { spec[s.name] = [[...s.header]]; });
  return fakeSpreadsheet(spec);
}

// ---------------------------------------------------------------
// 1. 点検（checkSheets_）
// ---------------------------------------------------------------

test('整っているシートでは、点検は何も言わない', () => {
  const { sandbox } = load();
  assert.deepEqual(plain(sandbox.checkSheets_(healthy(sandbox)).map((f) => f.kind)), []);
});

test('シートが1枚無ければ、見つけて「直せる」と言う', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  ss.deleteSheet(ss.getSheetByName('記録'));

  const found = sandbox.checkSheets_(ss);
  assert.equal(found.length, 1);
  assert.equal(found[0].sheet, '記録');
  assert.equal(found[0].kind, 'シートが無い');
  assert.equal(found[0].fixable, true);
});

test('見出しが別の言葉になっていたら、見つけて「直せない」と言う', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  // 先生が「記録」シートの先頭に1列挿し込んだ状態
  const sheet = ss.getSheetByName('記録');
  sheet._grid[0] = ['メモ', ...sheet._grid[0]];

  const found = sandbox.checkSheets_(ss).filter((f) => f.sheet === '記録');
  assert.ok(found.length > 0, '列のずれを見つけられていません');
  assert.ok(found.some((f) => f.kind === '見出しがちがう'));
  assert.equal(found.every((f) => f.fixable === false), true,
    '列がずれているのに「直せる」と言っています');
});

test('列がずれた見出しの説明には、何列目が何のはずかが出る', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('記録');
  sheet._grid[0] = ['メモ', ...sheet._grid[0]];

  const f = sandbox.checkSheets_(ss).find((x) => x.kind === '見出しがちがう');
  assert.match(f.detail, /1列目/);
  assert.match(f.detail, /logId/);
  assert.match(f.detail, /メモ/);
});

test('想定より右に列があれば知らせる（消しはしない）', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('単元');
  sheet._grid[0] = [...sheet._grid[0], '先生メモ'];

  const f = sandbox.checkSheets_(ss).find((x) => x.sheet === '単元' && x.kind === '列が多い');
  assert.ok(f, '増えた列に気づいていません');
  assert.equal(f.fixable, false);
  assert.match(f.detail, /先生メモ/);
});

// ---------------------------------------------------------------
// 2. 修整（repairSheets_）— ここがいちばん大事
// ---------------------------------------------------------------

test('列がずれているとき、見出しを書き換えてしまわない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('記録');
  sheet._grid[0] = ['メモ', ...sheet._grid[0]];
  sheet._grid.push(['先生のメモ', 'log-1', 'sess-1', 's001', 'BEFORE', '3', 'ほんぶん', new Date(), '']);
  const before = [...sheet._grid[0]];

  const result = sandbox.repairSheets_(ss);

  assert.deepEqual(plain(sheet._grid[0]), plain(before),
    '列がずれているのに見出しを上書きしました。事故が見えなくなります');
  assert.ok(result.left.some((f) => f.sheet === '記録'),
    '直せなかったことを報告していません');
});

test('空の列を1本挿し込まれたシートには、1列も書き足さない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('記録');

  // 先生が記録シートの先頭に空の列を1本挿した状態。
  // 挿した列は見出しも中身も空で、その右の見出しがすべて1つずつずれます。
  sheet._grid[0] = ['', ...sheet._grid[0]];
  sheet._grid.push(['', 'log-1', 'sess-1', 's001', 'BEFORE', '3', 'ほんぶん', '2026-08-23', '']);
  const before = plain(sheet._grid[0]);

  const result = sandbox.repairSheets_(ss);

  // ここで1列目に logId と書いてしまうと、「1列目は logId のはずなのに空、
  // 本物の logId は sessionId と書かれた2列目にある」状態になり、
  // ずれていることが誰にも見えなくなります。
  assert.deepEqual(plain(sheet._grid[0]), before,
    '列がずれているシートに見出しを書き足しました。ずれが見えなくなります');
  assert.ok(result.left.some((f) => f.sheet === '記録'),
    '直せなかったことを報告していません');
});

test('データが入っている列の空の見出しは、書き足さない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('名簿');
  sheet._grid[0][2] = '';                                   // ruby の見出しだけ消えている
  sheet._grid.push(['s001', '佐藤 健太', 'さとう けんた', '', '']);  // でも中身はある

  sandbox.repairSheets_(ss);

  assert.equal(sheet._grid[0][2], '',
    '中身のある列に見出しを書き足しました。列がずれている跡かもしれません');
});

test('中身の無い列の空の見出しは、書き足す（アプリの更新で列が増えた場合）', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('名簿');
  sheet._grid[0][4] = '';                                   // email の見出しだけ空
  sheet._grid.push(['s001', '佐藤 健太', 'さとう けんた', '', '']);  // email 列は空のまま

  const result = sandbox.repairSheets_(ss);

  assert.equal(sheet._grid[0][4], 'email', '空の列に見出しを書き足せていません');
  assert.ok(result.fixed.length > 0);
  assert.deepEqual(plain(result.left), [], '直したのに、まだ残っていると言っています');
});

test('消えたシートは、見出しつきで作り直す', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  ss.deleteSheet(ss.getSheetByName('記録'));

  const result = sandbox.repairSheets_(ss);

  const sheet = ss.getSheetByName('記録');
  assert.ok(sheet, 'シートを作り直していません');
  assert.deepEqual(plain(sheet._grid[0]), plain(sandbox.SHEETS_.find((s) => s.name === '記録').header));
  assert.deepEqual(plain(result.left), []);
});

test('直したあとの点検は、直っていないものだけを返す', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  ss.deleteSheet(ss.getSheetByName('単元'));                 // 直せる
  const sheet = ss.getSheetByName('記録');
  sheet._grid[0] = ['メモ', ...sheet._grid[0]];              // 直せない

  const result = sandbox.repairSheets_(ss);

  assert.ok(ss.getSheetByName('単元'), '単元シートが作り直されていません');
  assert.equal(result.left.every((f) => f.sheet === '記録'), true);
});

test('最後から2番目に列を挿されても、1列も書き足さない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sheet = ss.getSheetByName('記録');

  // timestamp（7列目）と deletedAt（8列目）のあいだに、空の列を1本挿した状態。
  // 挿した列は空なので「見出しがちがう」には引っかからず、
  // 想定の8列ぶんを見るだけでは「ずれていない」ように見えます。
  // ここで空いた8列目に deletedAt と書き足すと、見出しが
  // 「…, timestamp, deletedAt, deletedAt」になり、本物の削除の印は9列目のまま。
  // アプリは削除の印を8列目（r[7]）で読むので、
  // **消したはずの記録が全部戻り、二重送信の判定も効かなくなります。**
  sheet._grid[0].splice(7, 0, '');
  sheet._grid.push(['l1', 's1', 'st1', 'BEFORE', '3', 'ほんぶん', '2026-08-23', '', '削除済み']);
  const before = plain(sheet._grid[0]);

  const result = sandbox.repairSheets_(ss);

  assert.deepEqual(plain(sheet._grid[0]), before,
    '列がずれているのに見出しを書き足しました。削除の印が読めなくなります');
  assert.ok(result.left.some((f) => f.sheet === '記録'), '直せなかったことを報告していません');
});

test('点検が「直せる」と言ったものは、実際に直る', () => {
  // 点検はシートごと、修整は列ごと、という食い違いがあると、
  // 先生は直すつもりで OK を押したのに「直せるところはありませんでした」と
  // 言われることになります。両方が同じ判断を使っていることを確かめます。
  const { sandbox } = load();

  const cases = [
    (ss) => { ss.deleteSheet(ss.getSheetByName('単元')); },
    (ss) => { ss.getSheetByName('名簿')._grid[0][4] = ''; },
    (ss) => { const s = ss.getSheetByName('記録'); s._grid[0][0] = 'メモ'; },
    (ss) => { const s = ss.getSheetByName('記録'); s._grid[0].splice(7, 0, ''); s._grid.push(['l1','s1','st1','BEFORE','3','x','2026-08-23','','消']); },
    (ss) => { const s = ss.getSheetByName('単元'); s._grid[0].push('先生メモ'); },
  ];

  cases.forEach((breakIt, i) => {
    const ss = healthy(sandbox);
    breakIt(ss);
    const promised = sandbox.checkSheets_(ss).filter((f) => f.fixable);
    sandbox.repairSheets_(ss);
    const still = sandbox.checkSheets_(ss);

    promised.forEach((p) => {
      const stillThere = still.find((f) => f.sheet === p.sheet && f.detail === p.detail);
      assert.equal(stillThere, undefined,
        `case ${i}: 「直せる」と言った「${p.sheet}／${p.detail}」が直っていません`);
    });
  });
});

// ---------------------------------------------------------------
// 3. シートを用意する（ensureSheets_）
// ---------------------------------------------------------------

test('まっさらなスプレッドシートに、必要なシートがすべてそろう', () => {
  const { sandbox } = load();
  const ss = fakeSpreadsheet({ 'シート1': [] });

  sandbox.ensureSheets_(ss);

  sandbox.SHEETS_.forEach((spec) => {
    const sheet = ss.getSheetByName(spec.name);
    assert.ok(sheet, spec.name + ' シートが作られていません');
    assert.deepEqual(plain(sheet._grid[0]), plain(spec.header), spec.name + ' の見出しが違います');
  });
  assert.deepEqual(plain(sandbox.checkSheets_(ss)), [], '作った直後なのに点検が引っかかります');
});

test('空の「シート1」は片づけるが、中身のあるシートは消さない', () => {
  const { sandbox } = load();
  const withData = fakeSpreadsheet({ 'シート1': [['先生のメモ']] });
  sandbox.ensureSheets_(withData);
  assert.ok(withData.getSheetByName('シート1'), '中身のあるシートを消しました');

  const empty = fakeSpreadsheet({ 'シート1': [] });
  sandbox.ensureSheets_(empty);
  assert.equal(empty.getSheetByName('シート1'), null, '空のシート1が残っています');
});

test('すでに中身のあるシートを、作り直して消してしまわない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const logs = ss.getSheetByName('記録');
  logs._grid.push(['log-1', 'sess-1', 's001', 'BEFORE', '3', 'ほんぶん', new Date(), '']);

  sandbox.ensureSheets_(ss);

  assert.equal(logs._grid.length, 2, '既存の記録が消えました');
});

// ---------------------------------------------------------------
// 4. どのファイルをデータベースにするか（getDB）
// ---------------------------------------------------------------

test('束ねられたスプレッドシートがあれば、それを使う', () => {
  const { sandbox } = load({ bound: null });
  const bound = healthy(sandbox);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => bound;

  assert.equal(sandbox.getDB(), bound);
});

test('束ねられたファイルが無く、DB_ID も無ければ、作らずに止まる', () => {
  const { sandbox } = load({ bound: null });
  // 以前はここで新しいスプレッドシートを作り、DB_ID を差し替えていました。
  // 権限の足りない人が1回開くだけで学級全員の記録が空のファイルへ差し替わるので、
  // 作らずに止まることを確かめます。
  assert.throws(() => sandbox.getDB(), /見つかりません/);
});

test('DB_ID があるのに開けないとき、作り直さずに止まる', () => {
  const { sandbox } = load({ bound: null, props: { DB_ID: 'missing-id' } });
  sandbox.SpreadsheetApp.openById = () => { throw new Error('not found'); };

  assert.throws(() => sandbox.getDB(), /開けませんでした/);
});

// ---------------------------------------------------------------
// 5. 先生かどうかの判定
// ---------------------------------------------------------------

test('公開した本人は、設定なしでも先生として通る', () => {
  const { sandbox } = load({ activeUser: 'sensei@example.ed.jp', effectiveUser: 'sensei@example.ed.jp' });
  assert.equal(sandbox.isTeacher_(), true);
});

test('児童は先生として通らない', () => {
  const { sandbox } = load({ activeUser: 'child@example.ed.jp', effectiveUser: 'sensei@example.ed.jp' });
  assert.equal(sandbox.isTeacher_(), false);
});

test('メールアドレスが取れないときは、先生として通さない', () => {
  const { sandbox } = load({ activeUser: '', effectiveUser: '' });
  assert.equal(sandbox.isTeacher_(), false);
});

test('TEACHER_EMAILS があれば、そちらだけを見る', () => {
  const { sandbox } = load({
    activeUser: 'sensei@example.ed.jp',
    effectiveUser: 'sensei@example.ed.jp',
    props: { TEACHER_EMAILS: 'hoka@example.ed.jp' },
  });
  assert.equal(sandbox.isTeacher_(), false,
    'TEACHER_EMAILS に載っていない公開者が通っています');
});

test('OWNER_EMAIL がある学級（前の配り方）は、これまでどおり動く', () => {
  const { sandbox } = load({
    activeUser: 'owner@example.ed.jp',
    effectiveUser: 'sensei@example.ed.jp',
    props: { OWNER_EMAIL: 'owner@example.ed.jp' },
  });
  assert.equal(sandbox.isTeacher_(), true);
});

// ---------------------------------------------------------------
// 6. 初回セットアップの乗っ取り防止
// ---------------------------------------------------------------

test('児童のブラウザから manualSetup を呼んでも、先生になれない', () => {
  const { sandbox, store } = load({
    bound: null,
    activeUser: 'child@example.ed.jp',      // 呼んでいるのは児童
    effectiveUser: 'sensei@example.ed.jp',  // 実行しているのは先生（USER_DEPLOYING）
  });

  assert.throws(() => sandbox.manualSetup(), /公開したご本人/);
  assert.equal(store.OWNER_EMAIL, undefined,
    '児童が OWNER_EMAIL を取ってしまいました');
});

test('メニューの点検は、画面が無いところ（ウェブアプリ）では何も読まない', () => {
  const { sandbox } = load();
  const bound = healthy(sandbox);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => bound;
  // getUi() は画面が無いと例外になります。児童から google.script.run で
  // 呼ばれても、シートを1枚も読まずに終わることを確かめます。
  assert.throws(() => sandbox.showSheetCheck(), /getUi/);
  assert.throws(() => sandbox.showSheetRepair(), /getUi/);
});

// ---------------------------------------------------------------
// 7. 児童が読める範囲（getAnonymousOpinions）
// ---------------------------------------------------------------
// 「みんなの考えを見る」は児童も押す機能ですが、google.script.run は
// 誰でも直接呼べるので、sessionId を差し替えれば過去のどの授業の
// 記述本文でも引けてしまいました。道徳のノートには家庭のことが書かれます。

/** 授業と記録が入ったスプレッドシートを作ります。 */
function withLessons(sandbox) {
  const spec = {};
  sandbox.SHEETS_.forEach((s) => { spec[s.name] = [[...s.header]]; });
  spec['授業'].push(['now', new Date(), 'いまの授業', 'SLIDER', '{}', 'ACTIVE', 'BEFORE', '']);
  spec['授業'].push(['past', new Date(), '去年の授業', 'SLIDER', '{}', 'CLOSED', 'CLOSED', '']);
  spec['記録'].push(['l1', 'now', 's001', 'BEFORE', '3', 'いまの記述', new Date(), '']);
  spec['記録'].push(['l2', 'past', 's002', 'BEFORE', '4', 'おかあさんが入院していて', new Date(), '']);
  return fakeSpreadsheet(spec);
}

test('児童は、いま行われている授業の意見だけを読める', () => {
  const { sandbox } = load({ activeUser: 'child@example.ed.jp', effectiveUser: 'sensei@example.ed.jp' });
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => withLessons(sandbox);

  const opinions = plain(sandbox.getAnonymousOpinions('now'));
  assert.equal(opinions.length, 1);
  assert.equal(opinions[0].text, 'いまの記述');
});

test('児童は、終わった授業の記述本文を引けない', () => {
  const { sandbox } = load({ activeUser: 'child@example.ed.jp', effectiveUser: 'sensei@example.ed.jp' });
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => withLessons(sandbox);

  assert.throws(() => sandbox.getAnonymousOpinions('past'), /いま行われている授業/,
    '終わった授業の記述が引けてしまいます');
});

test('児童は、知らない授業IDを渡しても何も引けない', () => {
  const { sandbox } = load({ activeUser: 'child@example.ed.jp', effectiveUser: 'sensei@example.ed.jp' });
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => withLessons(sandbox);

  assert.throws(() => sandbox.getAnonymousOpinions('shiranai'), /いま行われている授業/);
  assert.throws(() => sandbox.getAnonymousOpinions(''), /授業が指定されていません|いま行われている授業/);
});

test('先生は、終わった授業も読める', () => {
  const { sandbox } = load({ activeUser: 'sensei@example.ed.jp', effectiveUser: 'sensei@example.ed.jp' });
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => withLessons(sandbox);

  const opinions = plain(sandbox.getAnonymousOpinions('past'));
  assert.equal(opinions.length, 1);
});

// ---------------------------------------------------------------
// 8. 二重送信（submitLog）
// ---------------------------------------------------------------

test('同じ児童の同じフェーズは、2回目が断られる', () => {
  const { sandbox } = load({ activeUser: '', effectiveUser: 'sensei@example.ed.jp' });
  const ss = withLessons(sandbox);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ss;

  const send = () => plain(sandbox.submitLog({
    sessionId: 'now', studentId: 's009', phase: 'BEFORE', value: '3', text: 'かんがえ',
  }));

  assert.equal(send().success, true);
  const second = send();
  assert.equal(second.success, false, '同じフェーズで2行入りました');
  assert.equal(second.duplicate, true);
});

test('重複の確認と書き込みは、ロックの中で行う', () => {
  // ロックが取れなかったときに、確認をすり抜けて書き込んでしまわないことを見ます。
  const { sandbox } = load({ activeUser: '', effectiveUser: 'sensei@example.ed.jp' });
  const ss = withLessons(sandbox);
  sandbox.SpreadsheetApp.getActiveSpreadsheet = () => ss;
  sandbox.LockService.getScriptLock = () => ({
    waitLock() { throw new Error('タイムアウト'); },
    releaseLock() {},
  });

  const before = ss.getSheetByName('記録')._grid.length;
  const result = plain(sandbox.submitLog({
    sessionId: 'now', studentId: 's009', phase: 'BEFORE', value: '3', text: 'かんがえ',
  }));

  assert.equal(result.success, false);
  assert.equal(ss.getSheetByName('記録')._grid.length, before,
    'ロックが取れていないのに書き込みました');
});

// ---------------------------------------------------------------
// 9. 見本の行が、消したのに戻ってこないこと
// ---------------------------------------------------------------

test('先生が中身を消したシートに、見本の児童名を書き戻さない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);

  // 先生がデモの3人を消し、見出しごと消してしまった状態
  const roster = ss.getSheetByName('名簿');
  roster._grid.length = 0;

  sandbox.ensureSheets_(ss);

  assert.deepEqual(plain(roster._grid[0]),
    plain(sandbox.SHEETS_.find((s) => s.name === '名簿').header),
    '見出しが戻っていません');
  assert.equal(roster._grid.length, 1,
    '消したはずの見本（佐藤 健太 など）が戻ってきました');
});

test('シートを新しく作るときは、見本の行を入れる', () => {
  const { sandbox } = load();
  const ss = fakeSpreadsheet({ 'シート1': [] });

  sandbox.ensureSheets_(ss);

  const roster = ss.getSheetByName('名簿');
  assert.ok(roster._grid.length > 1, '新規作成なのに見本が入っていません');
});

test('「直せるところを直す」は、見出しだけを書いて見本は入れない', () => {
  const { sandbox } = load();
  const ss = healthy(sandbox);
  const sessions = ss.getSheetByName('授業');
  sessions._grid.length = 0;

  sandbox.repairSheets_(ss);

  assert.equal(sessions._grid.length, 1,
    '見出しを直したついでに、デモの授業を足しました');
});
