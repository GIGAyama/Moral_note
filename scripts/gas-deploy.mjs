#!/usr/bin/env node
/* =====================================================================
 * gas-deploy.mjs — リポジトリの内容を Apps Script へ反映する（正本）
 * =====================================================================
 * 正本は GIGAyama.github.io/standards/gas/gas-deploy.mjs です。
 * 各リポジトリの scripts/gas-deploy.mjs はそのコピーで、CI の drift
 * ジョブが照合しています。直すときは正本を直してから配ってください。
 *
 * これまでは GAS エディタへ手でファイルをコピーしていました。1つ貼り
 * 忘れると起動時に「〇〇 is not defined」とだけ出ます。ここを自動にします。
 *
 * 使い方:
 *   node scripts/gas-deploy.mjs login    手元のGoogleアカウントで1度だけログインする
 *   node scripts/gas-deploy.mjs status   送るファイルを一覧する（送らない）
 *   node scripts/gas-deploy.mjs backup   いまGASにある中身を控える（送らない）
 *   node scripts/gas-deploy.mjs push     GASプロジェクトへ反映する
 *   node scripts/gas-deploy.mjs deploy   反映したうえで、既存のデプロイを新版へ更新する
 *
 * 環境変数:
 *   GAS_SCRIPT_ID          スクリプトID（GASエディタのURLに入っている）。必須
 *   GAS_DEPLOYMENT_ID      deploy のとき必須。更新するデプロイのID
 *   GAS_DEPLOYMENT_IDS     デプロイが複数あるとき（教師用と児童用など）。
 *                          カンマ区切り。GAS_DEPLOYMENT_ID の代わりに使う
 *   GAS_ROOT_DIR           GASプロジェクトの中身が置いてある場所。既定は
 *                          リポジトリ直下。gamification のように下の階層に
 *                          ある場合に指定する（例: manabi-quest）
 *   CLASPRC_JSON           任意。省略すると手元の ~/.clasprc.json を使う
 *   GAS_DEPLOY_DESCRIPTION 任意。デプロイに付ける説明（既定はコミットSHA）
 *   GAS_ALLOW_DELETIONS    後述の安全確認をあえて外すときだけ 1 にする
 * ===================================================================== */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import { spawnSync } from 'node:child_process';

const rootDir = path.resolve(process.cwd());
const BACKUP_DIR = path.join(rootDir, 'dist', 'gas-before-push');

/** GASプロジェクトの中身が置いてある場所（リポジトリからの相対）。 */
function projectRootDir() {
  const v = (process.env.GAS_ROOT_DIR || '').trim();
  return v || '.';
}

/** 使い方を示して終わります。 */
function usage(message) {
  console.error(message);
  console.error('');
  console.error('使い方: node scripts/gas-deploy.mjs <login|status|backup|push|deploy>');
  console.error('詳しくは standards/docs/gas-auto-deploy.md を参照してください。');
  process.exit(2);
}

/**
 * clasp の実体（JSファイル）を探します。見つからなければ入れ方を案内します。
 *
 * `require.resolve('@google/clasp/package.json')` は使えません。clasp の
 * `exports` が package.json を公開していないためです。node_modules を上へ辿ります。
 */
function resolveClaspEntry() {
  let dir = rootDir;
  for (;;) {
    const manifestPath = path.join(dir, 'node_modules', '@google', 'clasp', 'package.json');
    if (fs.existsSync(manifestPath)) {
      const manifest = JSON.parse(fs.readFileSync(manifestPath, 'utf8'));
      const bin = typeof manifest.bin === 'string' ? manifest.bin : manifest.bin.clasp;
      return path.join(path.dirname(manifestPath), bin);
    }
    const parent = path.dirname(dir);
    if (parent === dir) break;
    dir = parent;
  }
  console.error('clasp が入っていません。次を実行してください:');
  console.error('  npm run gas:install');
  process.exit(2);
}

/**
 * clasp を動かします。
 * シェルを挟まないので、Windowsの `.cmd` やクォートの違いに悩まされません。
 */
function clasp(args, { projectFile, authFile }) {
  const full = [];
  if (projectFile) full.push('--project', projectFile);
  if (authFile) full.push('--auth', authFile);
  full.push(...args);

  const entry = resolveClaspEntry();
  console.log('$ clasp ' + full.join(' '));
  const result = spawnSync(process.execPath, [entry, ...full], {
    stdio: 'inherit',
    cwd: rootDir
  });
  if (result.error) throw result.error;
  if (result.status !== 0) {
    console.error(`clasp ${args[0]} が失敗しました（終了コード ${result.status}）。`);
    process.exit(result.status || 1);
  }
}

/**
 * `.clasp.json` を書きます。スクリプトIDが入るためリポジトリには置かず（.gitignore）、
 * 実行のたびに環境変数から作り直します。
 *
 * fileExtension を 'gs' で固定しています。既定のままだと clasp pull が
 * サーバ側のコードを .js で書き出すので、控えとリポジトリのファイル名が
 * 食い違い、下の「消えるファイルの確認」が働かなくなります。
 *
 * @param {string} dir 置き場所
 * @param {string} scriptId
 * @param {string} [targetDir] 読み書きの対象。省略すると GAS_ROOT_DIR
 * @returns {string} 書いたファイルのパス
 */
function writeProjectFile(dir, scriptId, targetDir) {
  fs.mkdirSync(dir, { recursive: true });
  const file = path.join(dir, '.clasp.json');
  const project = {
    scriptId,
    rootDir: targetDir || projectRootDir(),
    fileExtension: 'gs'
  };
  fs.writeFileSync(file, JSON.stringify(project, null, 2) + '\n');
  return file;
}

/**
 * 認証情報を用意します。
 * CLASPRC_JSON があれば一時ファイルへ書き、無ければ手元の既定（~/.clasprc.json）に任せます。
 * @returns {?string} 認証ファイルのパス（既定に任せるときは null）
 */
function prepareAuthFile() {
  const raw = process.env.CLASPRC_JSON;
  if (!raw || !raw.trim()) return null;
  try {
    JSON.parse(raw);
  } catch {
    usage('CLASPRC_JSON がJSONとして読めません。clasp login で作られた .clasprc.json の中身を、そのまま入れてください。');
  }
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'clasp-auth-'));
  const file = path.join(dir, '.clasprc.json');
  // 他の利用者から読めない権限で置く（CIの共有ランナーでも同じ）
  fs.writeFileSync(file, raw, { mode: 0o600 });
  return file;
}

/** 環境変数を1つ読みます。空なら止めます。 */
function requireEnv(name, why) {
  const value = (process.env[name] || '').trim();
  if (!value) usage(`環境変数 ${name} が空です（${why}）。`);
  return value;
}

/**
 * 更新するデプロイのIDを読みます。
 * 教師用と児童用のように2つに分かれているアプリがあるので、複数を受けます。
 * @returns {string[]}
 */
export function deploymentIds(env) {
  const many = (env.GAS_DEPLOYMENT_IDS || '').trim();
  const one = (env.GAS_DEPLOYMENT_ID || '').trim();
  const raw = many || one;
  const ids = raw.split(',').map(s => s.trim()).filter(Boolean);
  // 同じIDを2度更新しても意味がないので、重複は落とす
  return [...new Set(ids)];
}

/** ファイル名から、GAS が同じものとみなす「見出し」を作ります。 */
export function fileStem(relPath) {
  return relPath.replace(/\.(gs|js|html|json)$/i, '');
}

/** ディレクトリの中身を再帰で並べます（相対パス）。 */
function listFiles(dir) {
  if (!fs.existsSync(dir)) return [];
  const out = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const rel = entry.name;
    if (entry.isDirectory()) {
      for (const child of listFiles(path.join(dir, rel))) out.push(path.join(rel, child));
    } else {
      out.push(rel);
    }
  }
  return out;
}

/**
 * `.claspignore` の1行をふるいの規則に直します。
 *
 * 対応するのは、この一群のリポジトリで実際に使っている書き方だけです。
 *   `**` … 区切りをまたいで何文字でも
 *   `*`  … 区切りをまたがずに何文字でも
 *   `!`  … 先頭に付けると「戻す」
 *   `#`  … 行頭のコメント
 *
 * **知らない書き方は、分かったふりをせず例外にします。**
 * ここで黙って素通りさせると「送らないのに安全と数える」という、
 * いま直しているのと同じ穴が別の形で開きます。
 */
export function claspIgnoreRule(line) {
  const negate = line.startsWith('!');
  const body = negate ? line.slice(1) : line;
  if (/[?\[\]{}()+^$]/.test(body)) {
    throw new Error(`.claspignore の書き方が読み取れません: ${line}`);
  }
  // 1文字ずつ見る。まとめて置きかえると `**/**` の扱いを取り違える
  // （区切りの前後を「必ず1つ以上」と読んでしまい、直下のファイルが外れる）。
  let source = '^';
  for (let i = 0; i < body.length; i++) {
    const ch = body[i];
    if (ch === '*') {
      if (body[i + 1] === '*') {
        if (body[i + 2] === '/') {
          source += '(?:.*\\/)?';   // `**/` … どの階層でもよい（無くてもよい）
          i += 2;
        } else {
          source += '.*';           // `**`  … 区切りをまたいで何文字でも
          i += 1;
        }
      } else {
        source += '[^/]*';          // `*`   … 区切りをまたがずに何文字でも
      }
    } else if (ch === '/') {
      source += '\\/';
    } else if (ch === '.' || ch === '\\') {
      source += '\\' + ch;
    } else {
      source += ch;
    }
  }
  return { negate, re: new RegExp(source + '$') };
}

/**
 * `clasp push` が実際に送るファイルだけに絞ります。
 *
 * ⚠️ ここを飛ばすと、**送らないファイルを「リポジトリにある」と数えて**しまいます。
 *    実例（2026-08-23・haiku-meeting）: リポジトリの `index.html` はサイトの
 *    トップで、`.claspignore` で外してあります。それでも名前が同じというだけで
 *    「安全」と数え、**本番の index.html が警告なしに消える**ところでした。
 *
 * 規則は gitignore と同じく **あとに書いたものが勝ちます**。
 * `.claspignore` が無いときは clasp の既定に任せるので、絞り込みもしません。
 *
 * @param {string[]} files      リポジトリにあるファイル（相対パス）
 * @param {?string} ignoreText  .claspignore の中身。無ければ null
 * @returns {string[]} 送るファイル
 */
export function filesToPush(files, ignoreText) {
  if (ignoreText === null || ignoreText === undefined) return files;
  const rules = ignoreText
    .split(/\r?\n/)
    .map(l => l.trim())
    .filter(l => l && !l.startsWith('#'))
    .map(claspIgnoreRule);
  if (rules.length === 0) return files;

  return files.filter(file => {
    const target = file.split(path.sep).join('/');
    let ignored = false;
    for (const rule of rules) {
      if (rule.re.test(target)) ignored = !rule.negate;
    }
    return !ignored;
  });
}

/**
 * 「送ると GAS から消えるファイル」を洗い出します。
 *
 * `clasp push --force` は GAS 側を丸ごと置き換えます。GASエディタで直接
 * 足したファイルや、リポジトリから消したつもりが本番にはまだ残っている
 * ファイルは、この push で**消えます**。学校が使っている最中に消えると
 * 戻せないので、控えと突き合わせて、消えるものがあれば止めます。
 *
 * @param {string[]} inGas   いまGASにあるファイル（控えから）
 * @param {string[]} inRepo  これから送るファイル
 * @returns {string[]} 送ると消えるファイル
 */
export function deletions(inGas, inRepo) {
  const have = new Set(inRepo.map(fileStem));
  return inGas.filter(f => !have.has(fileStem(f)));
}

/**
 * いまGASにある中身を、送る前に控えます。
 * 控えはリポジトリの外（dist/、.gitignore 済み）へ置きます。
 */
function backup(scriptId, authFile) {
  fs.rmSync(BACKUP_DIR, { recursive: true, force: true });
  fs.mkdirSync(BACKUP_DIR, { recursive: true });
  // 設定ファイルは控えの外へ置く。控えはそのままCIの成果物として持ち出すので、
  // スクリプトIDを混ぜない。
  const settingsDir = fs.mkdtempSync(path.join(os.tmpdir(), 'clasp-backup-'));
  const projectFile = writeProjectFile(settingsDir, scriptId, BACKUP_DIR);
  clasp(['pull'], { projectFile, authFile });
  console.log(`いまのGASプロジェクトを ${path.relative(rootDir, BACKUP_DIR)} に控えました。`);
}

/**
 * 控えとリポジトリを突き合わせて、送ってよいかを確かめます。
 * 危ないときは、ここで止めます（GAS には触れていない状態で終わります）。
 */
function assertSafeToPush() {
  const src = path.join(rootDir, projectRootDir());
  const inGas = listFiles(BACKUP_DIR);
  // .claspignore で外したファイルは「送らない」ので、リポジトリにあっても
  // 数に入れてはいけない。入れると、同じ名前のものが本番から黙って消える。
  //
  // ⚠️ .claspignore は **リポジトリ直下**（clasp を動かす場所）から読む。
  //    src（GAS_ROOT_DIR で下げた先）から読むと、GAS_ROOT_DIR を使っている
  //    リポジトリでファイルが見つからず、絞り込みが丸ごと効かなくなる。
  //    そこは clasp 自身の見方に合わせる。
  const ignoreFile = path.join(rootDir, '.claspignore');
  const ignoreText = fs.existsSync(ignoreFile) ? fs.readFileSync(ignoreFile, 'utf8') : null;
  const inRepo = filesToPush(listFiles(src), ignoreText)
    .filter(f => /\.(gs|html|json)$/i.test(f));

  if (inGas.length === 0) {
    console.log('控えが空でした。まだ中身の無いGASプロジェクトとみなして先へ進みます。');
    return;
  }

  // appsscript.json はGASプロジェクトの設定そのもの。これを欠いたまま送ると、
  // 権限（スコープ）やWebアプリの公開範囲が失われて、動かなくなる。
  if (!inRepo.some(f => path.basename(f) === 'appsscript.json')) {
    console.error('appsscript.json がリポジトリにありません。');
    console.error('これはGASプロジェクトの設定（権限・Webアプリの公開範囲）そのもので、');
    console.error('欠けたまま送ると本番が動かなくなります。');
    console.error('');
    console.error(`いま本番にあるものを ${path.relative(rootDir, BACKUP_DIR)}/appsscript.json に控えてあります。`);
    console.error('中身を確かめたうえでリポジトリに置き、コミットしてからやり直してください。');
    process.exit(1);
  }

  const gone = deletions(inGas, inRepo);
  if (gone.length === 0) return;

  if ((process.env.GAS_ALLOW_DELETIONS || '').trim() === '1') {
    console.log('次のファイルはGASから消えます（GAS_ALLOW_DELETIONS=1 のため続行します）:');
    gone.forEach(f => console.log('  - ' + f));
    return;
  }

  console.error('送るとGASから消えるファイルがあります。危ないので止めました。');
  gone.forEach(f => console.error('  - ' + f));
  console.error('');
  console.error('心当たりは次のどれかです:');
  console.error('  ・GASエディタで直接足したファイルがある（リポジトリに取り込んでください）');
  console.error('  ・.claspignore の書き方で、送るつもりのファイルが外れている');
  console.error('  ・本当に消したいファイルである');
  console.error('');
  console.error('本当に消してよいと分かっているときだけ、GAS_ALLOW_DELETIONS=1 を付けて実行してください。');
  console.error(`いまの本番の中身は ${path.relative(rootDir, BACKUP_DIR)} に控えてあります。`);
  process.exit(1);
}

// ── ここから実行 ──────────────────────────────────────────────
// テストから読み込むときは何もしない（純粋な関数だけを取り出せるように）
const invokedDirectly = process.argv[1]
  && path.resolve(process.argv[1]).endsWith('gas-deploy.mjs');

if (invokedDirectly) {
  const command = process.argv[2];
  if (!command || !['login', 'status', 'backup', 'push', 'deploy'].includes(command)) {
    usage(command ? `知らない指示です: ${command}` : '何をするか指定してください。');
  }

  // ログインだけは、スクリプトIDも認証ファイルも要らない。
  // 手元の既定の置き場所（~/.clasprc.json）へ書かせる。その中身を、あとで
  // GitHub の CLASPRC_JSON に登録する。
  if (command === 'login') {
    clasp(['login'], {});
    process.exit(0);
  }

  const scriptId = requireEnv('GAS_SCRIPT_ID', 'GASエディタのURLに入っているスクリプトID');
  const authFile = prepareAuthFile();
  const projectFile = writeProjectFile(rootDir, scriptId);

  if (command === 'status') {
    clasp(['status'], { projectFile, authFile });
  } else if (command === 'backup') {
    backup(scriptId, authFile);
  } else if (command === 'push') {
    backup(scriptId, authFile);
    assertSafeToPush();
    clasp(['push', '--force'], { projectFile, authFile });
  } else {
    const ids = deploymentIds(process.env);
    if (ids.length === 0) {
      usage('環境変数 GAS_DEPLOYMENT_ID（または GAS_DEPLOYMENT_IDS）が空です'
        + '（Apps Scriptの「デプロイを管理」に出ているデプロイID。既存のURLを保つために要ります）。');
    }
    backup(scriptId, authFile);
    assertSafeToPush();
    clasp(['push', '--force'], { projectFile, authFile });
    // 既存のデプロイを新しいバージョンへ差し替える。新規に作ると **URLが変わり**、
    // 先生が開いているブックマークやPWAが古いままになる。
    const label = (process.env.GAS_DEPLOY_DESCRIPTION || '').trim() || 'auto deploy';
    for (const id of ids) {
      clasp(['deploy', '--deploymentId', id, '--description', label], { projectFile, authFile });
    }
    const what = ids.length === 1 ? 'Webアプリ' : `${ids.length}つのデプロイ`;
    console.log(`${what}を新しいバージョンへ更新しました（URLは変わりません）。`);
  }
}
