'use strict';
/*
 * 「作成して保存」と同じ処理を、手元（Node）から実際のキントーンに対して通しで行い、結果を確かめる。
 * 画面（ボタン）以外の全部：書類マスタの雛形の取得 → 案件くんの値の差し込み → 案件くんの「書類保管」表への保存 → 保存物の中身の確認
 * ※ 画面のボタン操作そのものは、ブラウザで案件を開いて手で確認する（手順書の「動作確認」を参照）
 *
 *   node scripts/e2e-check.js --record=案件くんのレコード番号 --doc=書類名 [--cleanup]
 *
 *   --cleanup … 確認のために足した「書類保管」表の1行を、確認後に取り除く（もとの表に戻す）
 *
 * 案件くんへの書き込みは、「書類保管」表への1行の追加だけ（--cleanup なら、もとに戻すまで）。
 * 先に scripts/add-doc-storage.js で「書類保管」表を足しておくこと。
 */
const JSZip = require('jszip');
const { client } = require('./lib');
const lib = require('../customize/doc-button.js');

const arg = (n) => ((process.argv.find((a) => a.startsWith('--' + n + '=')) || '').split('=')[1] || '');
const NAME_LABELS = ['案件名', '現場名', '工事名'];

async function one(k, name) {
  const f = await k.findApps(name);
  if (f.length !== 1) throw new Error('「' + name + '」アプリが' + (f.length ? f.length + '個あります' : '見つかりません'));
  return f[0].appId;
}
const strip = (x) => x.replace(/<[^>]*>/g, '');

async function main() {
  const recordId = arg('record'), docName = arg('doc'), cleanup = process.argv.includes('--cleanup');
  const projectApp = process.env.KINTONE_APP_ID || '575';
  if (!recordId || !docName) { console.error('使い方：node scripts/e2e-check.js --record=案件のレコード番号 --doc=書類名 [--cleanup]'); process.exit(1); }
  const k = client();
  const results = [];
  const check = (ok, msg) => { results.push(ok); console.log((ok ? 'OK  ' : 'NG  ') + msg); };

  const masterId = await one(k, '書類マスタ');
  const proj = (await k.api('GET', '/k/v1/record', { app: projectApp, id: recordId })).record;
  const S = lib.SUB;
  if (!proj[S.table]) throw new Error('案件くんに「' + S.table + '」の表がありません。先に scripts/add-doc-storage.js で足してください');
  const cell = (row, c) => (row.value[c] && row.value[c].value != null ? String(row.value[c].value) : '');
  const keep = (row) => {
    const v = {}; [S.name, S.stage, S.version, S.at].forEach((c) => { v[c] = { value: cell(row, c) }; });
    v[S.file] = { value: ((row.value[S.file] && row.value[S.file].value) || []).map((f) => ({ fileKey: f.fileKey })) };
    return { id: row.id, value: v };
  };
  const rowsBefore = proj[S.table].value;
  const props = (await k.api('GET', '/k/v1/app/form/fields', { app: projectApp })).properties;
  const map = {}; Object.keys(proj).forEach((c) => { map[c] = c; });
  Object.keys(props).forEach((c) => { if (props[c].label) map[String(props[c].label).trim()] = props[c].code || c; });
  const values = lib.buildValues(proj, { map, props });
  check(Object.keys(values).length > 1, '案件くんの値を読み取りました（' + Object.keys(values).length + '項目）');

  const q = lib.MF.name + ' = ' + lib.quote(docName) + ' and ' + lib.MF.status + ' in ("使用中") limit 1';
  const mr = (await k.api('GET', '/k/v1/records', { app: masterId, query: q })).records[0];
  if (!mr) throw new Error('書類マスタに、使用中の「' + docName + '」がありません');
  const tpl = mr[lib.MF.template].value[0];
  if (!tpl) throw new Error('「' + docName + '」の雛形が書類マスタに添付されていません');
  const ext = ((tpl.name.match(/\.([^.]+)$/) || [])[1] || '').toLowerCase();
  const tplBuf = await k.download(tpl.fileKey);
  console.log('雛形：' + tpl.name + '（' + tplBuf.length + 'バイト）');

  const version = Math.max(0, ...rowsBefore.filter((r) => cell(r, S.name) === docName).map((r) => Number(cell(r, S.version) || 0))) + 1;
  const all = Object.assign({}, values, { '版': 'v' + version, '書類名': docName });
  const out = await lib.fillTemplate(JSZip, tplBuf, ext, all);

  let projectName = ''; NAME_LABELS.some((l) => { if (values[l]) { projectName = values[l]; return true; } return false; });
  const filename = lib.buildFilename(projectName || '案件' + recordId, docName, version, ext);
  const fileKey = await k.upload(out.data, filename);
  const row = { value: {} };
  row.value[S.name] = { value: docName };
  row.value[S.stage] = { value: mr[lib.MF.stage].value };
  row.value[S.version] = { value: String(version) };
  row.value[S.at] = { value: new Date().toISOString().slice(0, 10) };
  row.value[S.file] = { value: [{ fileKey }] };
  const upd = {}; upd[S.table] = { value: rowsBefore.map(keep).concat([row]) };
  await k.api('PUT', '/k/v1/record', { app: projectApp, id: recordId, revision: proj.$revision.value, record: upd });
  console.log('案件くんの「' + S.table + '」表に保存しました（v' + version + '）');

  // 保存された書類を取り直して、中身を確かめる
  const back = (await k.api('GET', '/k/v1/record', { app: projectApp, id: recordId })).record;
  const rowsAfter = back[S.table].value;
  check(rowsAfter.length === rowsBefore.length + 1, '表の行数：' + rowsBefore.length + ' → ' + rowsAfter.length + '（既存の行は消えていません）');
  const mine = rowsAfter.find((r) => cell(r, S.name) === docName && Number(cell(r, S.version)) === version);
  check(!!mine, '書類名・版が正しく保存されています');
  const f = mine && mine.value[S.file].value[0];
  check(!!f && f.name === filename, 'ファイル名：' + (f && f.name));
  if (['docx', 'xlsx', 'pptx'].includes(ext)) {
    const zip = await JSZip.loadAsync(await k.download(f.fileKey));
    let text = '';
    for (const n of Object.keys(zip.files)) if (/\.xml$/.test(n) && /^(word|xl|ppt)\//.test(n)) text += strip(await zip.file(n).async('string')) + '\n';
    const left = (text.match(/\{\{[^}]*\}\}/g) || []);
    check(left.length === 0, '{{ }} の残り：' + (left.length ? left.join(' ') + '（名前が違う、または段落・セルをまたいでいます）' : 'なし'));
    if (values['案件名']) check(text.includes(values['案件名']), '案件名「' + values['案件名'] + '」が書類に入っています');
  } else console.log('（' + ext + ' は差し込みなしのコピーです）');

  if (cleanup) {
    const cur = (await k.api('GET', '/k/v1/record', { app: projectApp, id: recordId })).record;
    const orig = {}; orig[S.table] = { value: rowsBefore.map(keep) };
    await k.api('PUT', '/k/v1/record', { app: projectApp, id: recordId, revision: cur.$revision.value, record: orig });
    console.log('確認用に足した「' + S.table + '」の1行を取り除きました（もとの表に戻しました）');
  } else console.log('※ 確認用の1行は残してあります（案件くんのレコード ' + recordId + ' の「' + S.table + '」表で確認・削除できます）');

  console.log('\n' + (results.every(Boolean) ? '結果：すべてOK' : '結果：NGがあります'));
  process.exit(results.every(Boolean) ? 0 : 5);
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
