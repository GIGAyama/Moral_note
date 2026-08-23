#!/usr/bin/env node
/**
 * .gs の構文検査。
 *
 * GAS は貼り付けて保存するまで誤りに気づけません。しかもこのリポジトリは
 * main に入ると数分で教室に届くので、その前に落とします。
 *
 * node は .gs という拡張子を知らないので、中身を読んで Function で組み立て、
 * 構文として通るかだけを見ます（実行はしません）。
 *
 * .html は GAS のテンプレート（<?= ?> を含みうる）なので対象外です。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const files = fs.readdirSync(ROOT).filter((f) => f.endsWith('.gs')).sort();

if (!files.length) {
  console.error('❌ .gs が1つもありません。取り違えていませんか。');
  process.exit(1);
}

let failed = 0;
for (const f of files) {
  const source = fs.readFileSync(path.join(ROOT, f), 'utf8');
  try {
    // 実行はしません。構文として読めるかどうかだけを見ます。
    new Function(source);
    console.log(`✅ ${f}`);
  } catch (e) {
    console.error(`❌ ${f}: ${e.message}`);
    failed++;
  }
}

if (failed) {
  console.error(`\n❌ ${failed} 件の .gs が構文として壊れています。`);
  process.exit(1);
}
console.log(`\n✅ ${files.length} 件の .gs は構文として読めます。`);
