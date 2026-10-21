'use strict';
/*
 * 案件くんの「書類関係」タブの中に、「書類を作る」ボタンを置くための「スペース」項目を足す。
 * これがあると、ボタンは、そのタブを開いたときだけ見える。
 *
 *   node scripts/add-button-space.js            … 足す位置を表示するだけ（何も変えない）
 *   node scripts/add-button-space.js --yes      … 実際に足して公開する
 *
 * 足す位置：ラベル「書類関係」の行の、すぐ下（書類保管の表の、すぐ上）。既存の項目・タブは変えない。
 *
 * 守ること
 *   ・変更前のバックアップ（scripts/backup-app.js）が無いときは、--yes でも実行しない
 *   ・同じ要素IDがすでにあれば、何もせずに止まる
 *   ・足したあと、書類関係の行以外の並びが同じかを確かめる
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');
const { BUTTON_SPACE_ID } = require('../customize/doc-button.js');

const ROOT = path.join(__dirname, '..');
const hasBackup = (appId) => {
  const d = process.env.KINTONE_BACKUP_DIR || path.join(ROOT, 'backup');
  return fs.existsSync(d) && fs.readdirSync(d).some((n) => n.startsWith(appId + '_') && fs.existsSync(path.join(d, n, 'customize_live.json')));
};
const isSection = (r) => r.type === 'ROW' && (r.fields || []).some((f) => f.type === 'LABEL' && /書類関係/.test(f.label));
const isSpace = (r) => r.type === 'ROW' && (r.fields || []).some((f) => f.type === 'SPACER' && f.elementId === BUTTON_SPACE_ID);

async function main() {
  const go = process.argv.includes('--yes');
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'customize-manifest.json'), 'utf8'));
  const appId = String(process.env.KINTONE_APP_ID || manifest.app);
  const k = client();
  const cur = await k.api('GET', '/k/v1/preview/app/form/layout', { app: appId });
  const L = cur.layout;
  if (L.some(isSpace)) { console.error('要素ID「' + BUTTON_SPACE_ID + '」のスペースが、すでにあります。何もせずに止めました。'); process.exit(2); }
  const at = L.findIndex(isSection);
  if (at < 0) { console.error('ラベル「書類関係」の行が見つかりません。何もせずに止めました。'); process.exit(2); }
  const row = { type: 'ROW', fields: [{ type: 'SPACER', elementId: BUTTON_SPACE_ID, size: { width: '200', height: '50' } }] };
  const next = L.slice(0, at + 1).concat([row], L.slice(at + 1));
  console.log('対象アプリ番号 ' + appId + '：「書類関係」の行（' + at + '番目）の、すぐ下に、ボタン用のスペース（要素ID ' + BUTTON_SPACE_ID + '）を足します。');
  console.log('既存の要素 ' + L.length + '個は、すべて残ります（→ ' + next.length + '個）。');
  if (!go) { console.log('\n※ これは確認だけです。実際に足すには、最後に --yes を付けてください。'); return; }
  if (!hasBackup(appId)) { console.error('\nバックアップがありません。先に  node scripts/backup-app.js ' + appId + '  を実行してください。'); process.exit(2); }

  await k.api('PUT', '/k/v1/preview/app/form/layout', { app: appId, layout: next, revision: cur.revision });
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: appId }] });
  await k.waitDeploy(appId, '案件くん');
  const after = (await k.api('GET', '/k/v1/app/form/layout', { app: appId })).layout;
  const rest = (l) => JSON.stringify(l.filter((r) => !isSpace(r)));
  const problems = [];
  if (!after.some(isSpace)) problems.push('スペースが見つかりません');
  if (rest(after) !== rest(L)) problems.push('スペース以外の並びが、変わっています');
  if (problems.length) { problems.forEach((p) => console.error('NG ' + p)); process.exit(4); }
  console.log('OK ボタン用のスペースを足しました（位置 ' + after.findIndex(isSpace) + '番目、全体 ' + L.length + ' → ' + after.length + '個、ほかの並びは同じ）');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
