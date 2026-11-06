'use strict';
/*
 * 案件くんの「書類保管」表に、承認とPDFの列を足す（書類関係の承認ボタン・PDF化ボタンのため）。
 *
 *   node scripts/add-approval-fields.js            … 足す内容を表示するだけ（何も変えない）
 *   node scripts/add-approval-fields.js --yes      … 実際に足して公開する
 *
 * 足す列：承認状況（文字）・承認者（文字）・承認日時（日時）・PDF（添付ファイル）。既存の列・行は変えない。
 *
 * 守ること
 *   ・変更前のバックアップ（scripts/backup-app.js）が無いときは、--yes でも実行しない
 *   ・同じ項目コードがすでにあれば、足さない（すでにある列は、そのまま）
 *   ・足したあと、列が増えたこと・既存の列が残っていることを確かめる
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');
const { SUB } = require('../customize/doc-button.js');

const ROOT = path.join(__dirname, '..');
const NEW = {
  [SUB.status]: { type: 'SINGLE_LINE_TEXT', code: SUB.status, label: '承認状況' },
  [SUB.by]: { type: 'SINGLE_LINE_TEXT', code: SUB.by, label: '承認者' },
  [SUB.approvedAt]: { type: 'DATETIME', code: SUB.approvedAt, label: '承認日時' },
  [SUB.pdf]: { type: 'FILE', code: SUB.pdf, label: 'PDF' },
};
const hasBackup = (appId) => {
  const d = process.env.KINTONE_BACKUP_DIR || path.join(ROOT, 'backup');
  return fs.existsSync(d) && fs.readdirSync(d).some((n) => n.startsWith(appId + '_') && fs.existsSync(path.join(d, n, 'customize_live.json')));
};

async function main() {
  const go = process.argv.includes('--yes');
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'customize-manifest.json'), 'utf8'));
  const appId = String(process.env.KINTONE_APP_ID || manifest.app);
  const k = client();
  const props = (await k.api('GET', '/k/v1/preview/app/form/fields', { app: appId })).properties;
  const table = props[SUB.table];
  if (!table || table.type !== 'SUBTABLE') { console.error('表「' + SUB.table + '」が見つかりません。先に add-doc-storage.js を実行してください。'); process.exit(2); }
  const add = Object.keys(NEW).filter((c) => !table.fields[c]);
  console.log('対象アプリ番号 ' + appId + '：表「' + SUB.table + '」の既存の列 ' + Object.keys(table.fields).length + '個');
  if (!add.length) { console.log('承認・PDFの列は、すでにあります。何もしません。'); return; }
  console.log('足す列：' + add.map((c) => NEW[c].label + '（' + NEW[c].type + '、コード ' + c + '）').join('、'));
  if (!go) { console.log('\n※ これは確認だけです。実際に足すには、最後に --yes を付けてください。'); return; }
  if (!hasBackup(appId)) { console.error('\nバックアップがありません。先に  node scripts/backup-app.js ' + appId + '  を実行してください。'); process.exit(2); }

  // 表の中に列を足す：表のコードと、足す列だけを渡す。受け付けなければ、既存の列を含めた全体を渡して、更新する
  const only = {}; add.forEach((c) => { only[c] = NEW[c]; });
  try {
    await k.api('POST', '/k/v1/preview/app/form/fields', { app: appId, properties: { [SUB.table]: { type: 'SUBTABLE', code: SUB.table, label: table.label || SUB.table, fields: only } } });
  } catch (e) {
    console.log('（追加だけでは受け付けられませんでした：' + String(e.message).split('\n')[0] + '。全体を渡して、更新します）');
    await k.api('PUT', '/k/v1/preview/app/form/fields', { app: appId, properties: { [SUB.table]: { type: 'SUBTABLE', code: SUB.table, label: table.label || SUB.table, fields: Object.assign({}, table.fields, only) } } });
  }
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: appId }] });
  await k.waitDeploy(appId, '案件くん');
  const after = (await k.api('GET', '/k/v1/app/form/fields', { app: appId })).properties[SUB.table];
  const problems = [];
  add.forEach((c) => { if (!after.fields[c]) problems.push('列「' + c + '」がありません'); });
  Object.keys(table.fields).forEach((c) => { if (!after.fields[c]) problems.push('既存の列「' + c + '」が消えています'); });
  if (problems.length) { problems.forEach((p) => console.error('NG ' + p)); process.exit(4); }
  console.log('OK 列を足しました（' + Object.keys(table.fields).length + ' → ' + Object.keys(after.fields).length + '列、既存の列は残っています）');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
