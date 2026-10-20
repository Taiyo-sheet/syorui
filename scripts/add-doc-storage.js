'use strict';
/*
 * 案件くんに、作った書類の保管場所（「書類保管」表）を足す。
 *
 *   node scripts/add-doc-storage.js            … 足す内容を表示するだけ（何も変えない）
 *   node scripts/add-doc-storage.js --yes      … 実際に足して公開する
 *
 * 足すもの：「書類保管」表（書類名・工程・版・作成日・ファイル）。既存の項目は変えない。
 *
 * 守ること
 *   ・変更前のバックアップ（scripts/backup-app.js）が無いときは、--yes でも実行しない
 *   ・同じ項目コードがすでにあれば、何もせずに止まる（上書きしない）
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');
const { SUB } = require('../customize/doc-button.js');

const ROOT = path.join(__dirname, '..');

const TABLE = {
  type: 'SUBTABLE', code: SUB.table,
  fields: {
    [SUB.name]: { type: 'SINGLE_LINE_TEXT', code: SUB.name, label: '書類名' },
    [SUB.stage]: { type: 'SINGLE_LINE_TEXT', code: SUB.stage, label: '工程' },
    [SUB.version]: { type: 'NUMBER', code: SUB.version, label: '版' },
    [SUB.at]: { type: 'DATE', code: SUB.at, label: '作成日' },
    [SUB.file]: { type: 'FILE', code: SUB.file, label: 'ファイル' },
  },
};

function hasBackup(appId) {
  const d = path.join(ROOT, 'backup');
  return fs.existsSync(d) && fs.readdirSync(d).some((n) => n.startsWith(appId + '_') && fs.existsSync(path.join(d, n, 'customize_live.json')));
}

async function main() {
  const go = process.argv.includes('--yes');
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, 'customize-manifest.json'), 'utf8'));
  const appId = String(process.env.KINTONE_APP_ID || manifest.app);
  const k = client();

  const app = await k.api('GET', '/k/v1/app', { id: appId });
  console.log('対象アプリ：「' + app.name + '」（番号 ' + appId + '）');
  const props = (await k.api('GET', '/k/v1/app/form/fields', { app: appId })).properties;
  const codes = [];
  Object.values(props).forEach((p) => { codes.push(p.code); if (p.type === 'SUBTABLE') Object.keys(p.fields).forEach((c) => codes.push(c)); });
  const clash = [SUB.table].concat(Object.keys(TABLE.fields)).filter((c) => codes.includes(c));
  if (clash.length) {
    console.error('項目コード「' + clash.join('」「') + '」がすでにあります。何もせずに止めました。');
    console.error('（すでに足し終えているか、同じ名前の別の項目があります。画面で確認してください）');
    process.exit(2);
  }
  console.log('\n足すもの：表「' + SUB.table + '」（案件くんの項目の一番下に入ります）');
  Object.values(TABLE.fields).forEach((f) => console.log('  ・' + f.label + '（' + f.type + '、コード ' + f.code + '）'));
  console.log('既存の項目は、変わりません（現在 ' + Object.keys(props).length + '項目）。');
  if (!go) { console.log('\n※ これは確認だけです。実際に足すには、最後に --yes を付けてください。'); return; }

  if (!hasBackup(appId)) { console.error('\nバックアップがありません。先に  node scripts/backup-app.js ' + appId + '  を実行してください。'); process.exit(2); }

  await k.api('POST', '/k/v1/preview/app/form/fields', { app: appId, properties: { [SUB.table]: TABLE } });
  console.log('項目を足しました。公開しています…');
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: appId }] });
  await k.waitDeploy(appId, app.name);

  const after = (await k.api('GET', '/k/v1/app/form/fields', { app: appId })).properties;
  const t = after[SUB.table];
  const missing = t && t.type === 'SUBTABLE' ? Object.keys(TABLE.fields).filter((c) => !t.fields[c]) : Object.keys(TABLE.fields);
  if (!t || missing.length) { console.error('NG 公開後の確認：足りない項目 ' + (t ? missing.join('、') : SUB.table)); process.exit(4); }
  console.log('OK 「' + SUB.table + '」表が追加されました（既存 ' + Object.keys(props).length + '項目 → ' + Object.keys(after).length + '項目）');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
