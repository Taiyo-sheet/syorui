'use strict';
/*
 * 見積くん(608)の「PC版」見積作成ボタン（mitsumori-open-608.js）に、外部アプリを開くための鍵（?k=）を足す。
 *   node scripts/swap-608.js         … 確認だけ（何も変えない）
 *   node scripts/swap-608.js --yes   … 実際に入れ替える
 *
 * 鍵は、同じアプリの「スマホ版」ボタン（mitsumori-open-mobile-608.js）に入っているものを、読み取って使う（リポジトリには、鍵を入れない）。
 * 入れ替えるのは、PC版の mitsumori-open-608.js の1ファイルだけ。読み込み順・ほかのファイル・フォームは、変えない。
 * 手順：①フォームの編集中の変更がないか確認 ②バックアップ ③同じ位置で入れ替え→公開 ④順番・ファイル・フォームが同じか確認 ⑤おかしければ、バックアップのJS登録に戻す
 * 利用者の許可：2026-11-02「A」（見積くん(608)の見積ボタンに、鍵を付けてよい）
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { client } = require('./lib');

const ROOT = path.join(__dirname, '..');
const APP = 608, PC = 'mitsumori-open-608.js', MOBILE = 'mitsumori-open-mobile-608.js';
const yes = process.argv.includes('--yes');
const say = (t) => console.log(t);
const strip = (it) => (it.type === 'URL' ? it : { type: 'FILE', file: { fileKey: it.file.fileKey } });

async function snap(k) {
  const layout = (await k.api('GET', '/k/v1/app/form/layout', { app: APP })).layout;
  const fields = Object.keys((await k.api('GET', '/k/v1/app/form/fields', { app: APP })).properties).sort();
  return JSON.stringify(layout) + '|' + JSON.stringify(fields);
}
const names = (c) => ['desktop', 'mobile'].map((e) => ((c[e] && c[e].js) || []).map((x) => (x.file ? x.file.name : x.url)).join(','));

async function putCustomize(k, cur, replace) {
  const body = { app: APP, desktop: { js: [], css: (cur.desktop.css || []).map(strip) }, mobile: { js: (cur.mobile.js || []).map(strip), css: (cur.mobile.css || []).map(strip) } };
  for (const it of cur.desktop.js || []) body.desktop.js.push(it.file && it.file.name === PC && replace ? { type: 'FILE', file: { fileKey: await k.upload(replace, PC, 'text/javascript') } } : strip(it));
  if (cur.scope) body.scope = cur.scope;
  await k.api('PUT', '/k/v1/preview/app/customize', body);
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: APP }] });
  await k.waitDeploy(APP, String(APP));
}

async function main() {
  const k = client();
  const cur = await k.api('GET', '/k/v1/app/customize', { app: APP });
  const pcItem = (cur.desktop.js || []).find((x) => x.file && x.file.name === PC);
  const mbItem = (cur.mobile.js || []).find((x) => x.file && x.file.name === MOBILE);
  if (!pcItem || !mbItem) { say('PC版またはスマホ版のボタンのファイルが見つかりません。止めました'); process.exit(10); }
  const pc = Buffer.from(await k.download(pcItem.file.fileKey)).toString('utf8');
  const mb = Buffer.from(await k.download(mbItem.file.fileKey)).toString('utf8');
  const key = (mb.match(/var OPEN_KEY = '([^']+)'/) || [])[1];
  if (!key) { say('スマホ版ボタンに鍵が見つかりません。止めました'); process.exit(11); }
  if (/OPEN_KEY/.test(pc)) { say('PC版ボタンには、すでに鍵の欄があります。何もしません'); return; }
  if (!pc.includes('    var q = [];\n') || !pc.includes("var BASE = 'https://anken-kanri-1.onrender.com/mitsumori.html';\n")) { say('PC版ボタンの中身が想定と違うため、止めました（手で確認が必要です）'); process.exit(12); }
  const patched = pc
    .replace("var BASE = 'https://anken-kanri-1.onrender.com/mitsumori.html';\n", "var BASE = 'https://anken-kanri-1.onrender.com/mitsumori.html';\n  // kintoneのボタンからのみ開けるようにするための鍵（アプリ側のOPEN_KEYと一致させる）。スマホ版と同じ\n  var OPEN_KEY = " + JSON.stringify(key) + ";\n")
    .replace('    var q = [];\n', "    var q = [];\n    if (OPEN_KEY) q.push('k=' + encodeURIComponent(OPEN_KEY));\n");
  new Function(patched);   // 文法の確認
  say('見積くん(608) PC版ボタンに、鍵（スマホ版と同じ）を足します。入れ替えるのは ' + PC + ' の1ファイルだけ。');
  if (!yes) { say('（--yes がないため、確認だけです。何も変えていません）'); return; }

  const dry = spawnSync('node', ['scripts/apply-customize.js', '--manifest=customize-manifest-buttons-608.json'], { cwd: ROOT, encoding: 'utf8', env: process.env });
  if (/編集中/.test(dry.stdout + dry.stderr)) { say('フォームの編集中の変更があります。止めました（何も変えていません）'); process.exit(13); }
  const b = spawnSync('node', ['scripts/backup-app.js', String(APP)], { cwd: ROOT, encoding: 'utf8', env: process.env });
  const dir = ((b.stdout + b.stderr).match(/バックアップ先：(.+?)（/) || [])[1];
  if (b.status !== 0 || !dir) { say('バックアップに失敗しました。止めました'); process.exit(14); }
  say('バックアップ：' + dir);
  const before = await snap(k);
  await putCustomize(k, cur, Buffer.from(patched, 'utf8'));
  const after = await k.api('GET', '/k/v1/app/customize', { app: APP });
  const it = (after.desktop.js || []).find((x) => x.file && x.file.name === PC);
  const problems = [];
  if (JSON.stringify(names(after)) !== JSON.stringify(names(cur))) problems.push('読み込み順・ファイル名が、反映前と違います');
  if (!it || !Buffer.from(await k.download(it.file.fileKey)).equals(Buffer.from(patched, 'utf8'))) problems.push('登録したファイルが、想定と違います');
  if ((await snap(k)) !== before) problems.push('フォームの並び・項目が、反映前と違います');
  if (problems.length) {
    say('確認で問題がありました：\n  ・' + problems.join('\n  ・') + '\nバックアップのJS登録に戻します');
    try { await putCustomize(k, cur, null); say('戻しました（元のJSの登録）。バックアップ：' + dir); } catch (e) { say('自動で戻せませんでした：' + (e.message || e) + '\n手動で戻してください。バックアップ：' + dir); }
    process.exit(15);
  }
  say('完了：読み込み順・ファイル・フォームは、反映前と同じでした。画面を Ctrl+F5 で再読み込みして確認してください。バックアップ：' + dir);
}
main().catch((e) => { console.error(e.message || e); process.exit(1); });
