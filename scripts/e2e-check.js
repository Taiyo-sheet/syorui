'use strict';
/*
 * 「作成して保存」と同じ処理を、手元（Node）から実際のキントーンに対して通しで行い、結果を確かめる。
 * 画面（ボタン）以外の全部：書類マスタの雛形の取得 → 案件くんの値の差し込み → 案件書類への保存 → 保存物の中身の確認
 * ※ 画面のボタン操作そのものは、ブラウザで案件を開いて手で確認する（手順書の「動作確認」を参照）
 *
 *   node scripts/e2e-check.js --record=案件くんのレコード番号 --doc=書類名 [--cleanup]
 *
 *   --cleanup … 確認のために作った案件書類のレコードを、確認後に削除する（この実行で作った1件だけ）
 *
 * 案件くんは読むだけ。書き込むのは案件書類アプリへの1件の追加（--cleanup なら削除まで）。
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

  const masterId = await one(k, '書類マスタ'), docsId = await one(k, '案件書類');
  const proj = (await k.api('GET', '/k/v1/record', { app: projectApp, id: recordId })).record;
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

  const lq = lib.DF.projectId + ' = ' + Number(recordId) + ' and ' + lib.DF.name + ' = ' + lib.quote(docName) + ' order by ' + lib.DF.version + ' desc limit 1';
  const last = (await k.api('GET', '/k/v1/records', { app: docsId, query: lq, fields: ['$id', lib.DF.version] })).records[0];
  const version = (last ? Number(last[lib.DF.version].value) : 0) + 1;
  const all = Object.assign({}, values, { '版': 'v' + version, '書類名': docName });
  const out = await lib.fillTemplate(JSZip, tplBuf, ext, all);

  let projectName = ''; NAME_LABELS.some((l) => { if (values[l]) { projectName = values[l]; return true; } return false; });
  const filename = lib.buildFilename(projectName || '案件' + recordId, docName, version, ext);
  const fileKey = await k.upload(out.data, filename);
  const rec = {};
  rec[lib.DF.projectId] = { value: String(recordId) };
  rec[lib.DF.name] = { value: docName };
  rec[lib.DF.stage] = { value: mr[lib.MF.stage].value };
  rec[lib.DF.version] = { value: String(version) };
  rec[lib.DF.file] = { value: [{ fileKey }] };
  const saved = await k.api('POST', '/k/v1/record', { app: docsId, record: rec });
  console.log('案件書類に保存しました：レコード ' + saved.id + '（v' + version + '）');

  // 保存された書類を取り直して、中身を確かめる
  const back = (await k.api('GET', '/k/v1/record', { app: docsId, id: saved.id })).record;
  check(Number(back[lib.DF.version].value) === version && Number(back[lib.DF.projectId].value) === Number(recordId), '案件ID・版が正しく保存されています');
  const f = back[lib.DF.file].value[0];
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
    await k.api('DELETE', '/k/v1/records', { app: docsId, ids: [Number(saved.id)] });
    console.log('確認用に作った案件書類のレコード ' + saved.id + ' を削除しました');
  } else console.log('※ 確認用のレコード ' + saved.id + ' は残してあります（案件書類アプリで確認・削除できます）');

  console.log('\n' + (results.every(Boolean) ? '結果：すべてOK' : '結果：NGがあります'));
  process.exit(results.every(Boolean) ? 0 : 5);
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
