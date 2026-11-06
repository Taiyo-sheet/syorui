'use strict';
/*
 * 書類マスタ(879)に、「書類番号」の欄を足し、書類ごとに番号（scripts/doc-numbers.json）を入れる（利用者の許可：2026-11-06「すべておすすめで進めて」＝確認② A）。
 *
 *   node scripts/add-doc-no-879.js                 … 足す内容を表示するだけ（何も変えない）
 *   node scripts/add-doc-no-879.js --yes           … 欄を足して公開する（書類名の行の、すぐ下の行に置く。既存の項目・行は変えない）
 *   node scripts/add-doc-no-879.js --numbers       … 番号を入れる内容を表示するだけ
 *   node scripts/add-doc-no-879.js --numbers --yes … 番号を入れる（使用中の書類。立場違い（施工店用など）は、同じ番号）
 *
 * 守ること：変更前のバックアップ（backup/879_...）が無いときは、--yes でも実行しない／編集中（未公開）の変更が、あるときは止める
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');

const ROOT = path.join(__dirname, '..');
const APP = 879, CODE = '書類番号';
const NUMS = JSON.parse(fs.readFileSync(path.join(__dirname, 'doc-numbers.json'), 'utf8'));
const baseOf = (n) => String(n).replace(/（(施工店用|工事部用|施工指導用)）$/, '');

function hasBackup() {
  const d = process.env.KINTONE_BACKUP_DIR || path.join(ROOT, 'backup');
  return fs.existsSync(d) && fs.readdirSync(d).some((n) => n.startsWith(APP + '_') && fs.existsSync(path.join(d, n, 'customize_live.json')));
}
const rowsOf = (lay) => lay.layout.filter((x) => x.type === 'ROW').length;

async function main() {
  const go = process.argv.includes('--yes'), nums = process.argv.includes('--numbers');
  const k = client();
  if (nums) {
    const r = await k.api('GET', '/k/v1/records', { app: APP, query: '使用状況 in ("使用中") order by 表示順 asc limit 200' });
    const plan = r.records.map((x) => ({ id: x.$id.value, name: x['書類名'].value, no: NUMS[x['書類名'].value] || NUMS[baseOf(x['書類名'].value)], cur: (x[CODE] || {}).value || '' }));
    const miss = plan.filter((p) => !p.no); const todo = plan.filter((p) => p.no && p.no !== p.cur);
    console.log('使用中の書類 ' + plan.length + '件／番号を入れる ' + todo.length + '件／番号が決まっていない ' + miss.length + '件');
    miss.forEach((p) => console.log('  番号なし：' + p.name));
    if (!go) { todo.slice(0, 8).forEach((p) => console.log('  ' + p.no + '  ' + p.name)); console.log('\n※ 確認だけです。実際に入れるには --numbers --yes'); return; }
    if (!hasBackup()) { console.error('バックアップがありません。先に  npm run backup -- 879'); process.exit(3); }
    for (let i = 0; i < todo.length; i += 100) {
      await k.api('PUT', '/k/v1/records', { app: APP, records: todo.slice(i, i + 100).map((p) => ({ id: p.id, record: { [CODE]: { value: p.no } } })) });
    }
    console.log('OK ' + todo.length + '件に番号を入れました');
    return;
  }
  const live = await k.api('GET', '/k/v1/app/form/fields', { app: APP });
  const pre = await k.api('GET', '/k/v1/preview/app/form/fields', { app: APP });
  if (Object.keys(live.properties).length !== Object.keys(pre.properties).length) { console.error('編集中（未公開）の変更があります。止めました。'); process.exit(2); }
  if (live.properties[CODE]) { console.log('「' + CODE + '」は、すでにあります。何もしません。'); return; }
  const lay = await k.api('GET', '/k/v1/app/form/layout', { app: APP });
  const idx = lay.layout.findIndex((x) => x.type === 'ROW' && x.fields.some((f) => f.code === '書類名'));
  if (idx < 0) { console.error('書類名の行が見つかりません。止めました'); process.exit(2); }
  console.log('足す欄：' + CODE + '（文字列1行）／場所：書類名の行の、すぐ下／いま ' + Object.keys(live.properties).length + '項目・' + rowsOf(lay) + '行');
  if (!go) { console.log('\n※ 確認だけです。実際に足すには --yes'); return; }
  if (!hasBackup()) { console.error('バックアップがありません。先に  npm run backup -- 879'); process.exit(3); }
  await k.api('POST', '/k/v1/preview/app/form/fields', { app: APP, properties: { [CODE]: { type: 'SINGLE_LINE_TEXT', code: CODE, label: CODE, maxLength: '10' } } });
  const lay2 = await k.api('GET', '/k/v1/preview/app/form/layout', { app: APP });
  const at = lay2.layout.findIndex((x) => x.type === 'ROW' && x.fields.some((f) => f.code === '書類名'));
  lay2.layout.splice(at + 1, 0, { type: 'ROW', fields: [{ type: 'SINGLE_LINE_TEXT', code: CODE, size: { width: '200' } }] });
  await k.api('PUT', '/k/v1/preview/app/form/layout', { app: APP, layout: lay2.layout });
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: APP }] });
  await k.waitDeploy(APP, '書類マスタ');
  const after = await k.api('GET', '/k/v1/app/form/fields', { app: APP });
  const lay3 = await k.api('GET', '/k/v1/app/form/layout', { app: APP });
  const problems = [];
  if (!after.properties[CODE]) problems.push('欄がありません');
  Object.keys(live.properties).forEach((c) => { if (!after.properties[c]) problems.push('既存の項目「' + c + '」が消えています'); });
  if (rowsOf(lay3) !== rowsOf(lay) + 1) problems.push('行の数が、想定と違います（' + rowsOf(lay) + ' → ' + rowsOf(lay3) + '）');
  if (problems.length) { problems.forEach((p) => console.error('NG ' + p)); process.exit(4); }
  console.log('OK 欄を足しました（' + Object.keys(live.properties).length + ' → ' + Object.keys(after.properties).length + '項目、' + rowsOf(lay) + ' → ' + rowsOf(lay3) + '行。既存の項目は残っています）');
}
main().catch((e) => { console.error(e.message || e); process.exit(1); });
