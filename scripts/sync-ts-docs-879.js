'use strict';
/*
 * BoxのTS文書管理（台帳）に合わせて、書類マスタ(879)を整える（利用者の許可：2026-11-14「AAA」＝文書番号に合わせる・zipで渡す・工程の7区分は、そのまま）。
 *
 *   node scripts/sync-ts-docs-879.js --dir=展開したフォルダ                    … 計画を表示するだけ（何も変えない）
 *   node scripts/sync-ts-docs-879.js --dir=… --form --yes                     … 「配布区分」の欄を足し、「書類番号」の字数を広げる（既存の項目は、変えない）
 *   node scripts/sync-ts-docs-879.js --dir=… --records --yes                  … 既存の書類の番号・配布区分を入れ替え、新しい書類を追加する
 *
 * 守ること：変更前のバックアップ（backup/879_...）が無いときは、--yes でも実行しない／同じ文書番号が、すでにある書類は、追加しない（二重に作らない）
 *          既存の雛形に {{ }} があるものは、差し替えない（自動入力が、消えるため）。{{ }} が無い雛形だけ、Boxの正本に差し替える（旧版は残る）
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { client } = require('./lib');

const ROOT = path.join(__dirname, '..');
const APP = 879;
const arg = (n) => ((process.argv.find((a) => a.startsWith('--' + n + '=')) || '').split('=').slice(1).join('='));
const has = (n) => process.argv.includes('--' + n);
const DIST = { '02_NDA限定': 'NDA限定', '03_社内限定': '社内限定', '01_配布可': '配布可' };
const NO = (c) => 'TS-RH-' + c + '-2611';

// 既存の書類（書類マスタの名前）→ Boxの文書番号（確かなものだけ。ほかは、番号を空にして、利用者に確認する）
const EXIST = {
  '作業手順書（施工店用）': 'WPR', '作業手順書（工事部用）': 'WPR', '作業手順書（施工指導用）': 'WPR',
  '施工計画書': 'CPL', '仕様書（施工店用）': 'SPC', '仕様書（工事部用）': 'SPC', '仕様書（施工指導用）': 'SPC', 'リスクアセスメント': 'RSK',
  '現地調査チェックリスト': 'SSC', '着工前チェックリスト': 'PCC', '切り替えタイムスケジュール': 'SWS',
  '準備品リスト（工事部用）': 'PRP', '準備品リスト（施工店用）': 'MAT', '準備品リスト（施工指導用）': 'MAT',
  '出来形品質管理表': 'AQC', '施工記録写真管理基準': 'PHO', '竣工図書作成マニュアル': 'CDM',
  '施工管理マニュアル（外部パートナー用）': 'CMP', '施工管理マニュアル（工事部用）': 'CMD', '施工指導マニュアル（施工指導担当者用）': 'SVM', '施工管理マニュアル（設備編）': 'CMR',
  '打合せ議事録（社内）': 'IMM', '打合せ議事録（社外）': 'CMM', '提案書（元請向け）': 'PRC', '提案書（発注者向け）': 'PRO',
  '原価管理表': 'CCS', '安全書類一式（施工店用）': 'SFD', '安全書類一式（工事部用）': 'SFD', '安全書類一式（施工指導用）': 'SFD',
  '建設機械作業計画書（バックホウ・移動式クレーン）': 'HEP', '竣工後ふりかえりシート（簡易版）': 'PCS', '竣工後ふりかえりシート（詳細版）': 'PCM',
};
// 新しい書類の工程（台帳の「必要書類40」のフェーズ・工程を、書類マスタの7区分に当てはめた。台帳にないマニュアル類は「受注・準備」）
const STAGE = {
  NTC: '情報取得・現状確認', IRL: '情報取得・現状確認', SSD: '情報取得・現状確認', SDM: '情報取得・現状確認', SSR: '情報取得・現状確認', QRS: '情報取得・現状確認', QRB: '情報取得・現状確認',
  IMC: '打合せ・見積・提案', IRP: '打合せ・見積・提案', CMC: '打合せ・見積・提案', HLD: '打合せ・見積・提案', CRD: '打合せ・見積・提案', CPD: '打合せ・見積・提案', QTY: '打合せ・見積・提案', SCS: '打合せ・見積・提案', RSL: '打合せ・見積・提案', SHP: '打合せ・見積・提案', EBD: '打合せ・見積・提案', EQN: '打合せ・見積・提案', PRL: '打合せ・見積・提案', PMM: '打合せ・見積・提案',
  CTR: '受注・準備',
  SCD: '施工前の書類', TRF: '施工前の書類', WRF: '施工前の書類', PMT: '施工前の書類',
  CNT: '竣工', PHA: '竣工', EI1: '竣工', EI2: '竣工', EI3: '竣工', EI4: '竣工', EI5: '竣工', ABD: '竣工', CRP: '竣工', WMF: '竣工',
};
const BASE = { '情報取得・現状確認': 160, '打合せ・見積・提案': 260, '受注・準備': 360, '施工前の書類': 460, '施工・試験': 560, '竣工': 660, '請求・原価管理': 760 };
const KIND = { docx: 'Word', xlsx: 'Excel', xlsm: 'Excel', pptx: 'PowerPoint', pdf: 'PDF' };
const AUD = { 配布可: ['TS計画推進', '施工指導', '外部パートナー', '施工店', '工事部', '営業'], NDA限定: ['TS計画推進', '外部パートナー', '施工店', '工事部'], 社内限定: ['TS計画推進', '施工指導', '工事部', '営業'] };

function hasBackup() {
  const d = process.env.KINTONE_BACKUP_DIR || path.join(ROOT, 'backup');
  return fs.existsSync(d) && fs.readdirSync(d).some((n) => n.startsWith(APP + '_') && fs.existsSync(path.join(d, n, 'customize_live.json')));
}
function scan(dir) {
  const out = [];
  for (const d of fs.readdirSync(dir)) {
    if (!DIST[d]) continue;
    for (const f of fs.readdirSync(path.join(dir, d))) {
      const m = f.match(/^(TS-RH-([A-Z0-9]+)-2611)_(.*)_v01\.([a-z]+)$/);
      if (m) out.push({ file: path.join(dir, d, f), fname: f, no: m[1], code: m[2], name: m[3], ext: m[4], dist: DIST[d] });
    }
  }
  return out;
}

async function main() {
  const dir = arg('dir'), go = has('yes');
  if (!dir) { console.error('--dir=（zipを展開したフォルダ）を付けてください'); process.exit(1); }
  const files = scan(dir), byNo = {}; files.forEach((x) => { byNo[x.no] = x; });
  const k = client();
  const all = await k.api('GET', '/k/v1/records', { app: APP, query: '使用状況 in ("使用中") order by 表示順 asc limit 200' });
  const have = new Set(all.records.map((r) => (r['書類番号'] || {}).value).filter((v) => /^TS-RH-/.test(v || '')));
  const mappedNos = new Set(Object.values(EXIST).map(NO));
  const toAdd = files.filter((x) => !mappedNos.has(x.no) && !have.has(x.no));
  const upd = all.records.filter((r) => EXIST[r['書類名'].value]).map((r) => { const n = r['書類名'].value, c = EXIST[n]; return { id: r.$id.value, name: n, no: c ? NO(c) : '', dist: c ? ((byNo[NO(c)] || {}).dist || (c === 'PRO' ? '配布可' : '')) : '', cur: (r['書類番号'] || {}).value || '' }; });
  console.log('Boxの書類 ' + files.length + '件／既存の使用中 ' + all.records.length + '件');
  console.log('既存：番号・配布区分を入れる ' + upd.length + '件');
  all.records.filter((r) => !EXIST[r['書類名'].value] && !/^TS-RH-/.test((r['書類番号'] || {}).value || '')).forEach((r) => console.log('   番号なし（台帳に、対応が見つからない）：' + r['書類名'].value));
  console.log('新しく追加 ' + toAdd.length + '件');
  if (!go) { toAdd.slice(0, 6).forEach((x) => console.log('   ' + x.no + '  ' + x.name + '  [' + (STAGE[x.code] || '受注・準備') + '／' + x.dist + ']')); console.log('\n※ 確認だけです。--form --yes ／ --records --yes で実行します'); return; }
  if (!hasBackup()) { console.error('バックアップがありません。先に  npm run backup -- 879'); process.exit(3); }

  if (has('form')) {
    const live = await k.api('GET', '/k/v1/app/form/fields', { app: APP }), pre = await k.api('GET', '/k/v1/preview/app/form/fields', { app: APP });
    const mine = Object.keys(pre.properties).filter((c) => !live.properties[c]);
    if (mine.length && !(mine.length === 1 && mine[0] === '配布区分')) { console.error('編集中（未公開）の変更があります。止めました'); process.exit(2); }
    const lay = await k.api('GET', '/k/v1/app/form/layout', { app: APP });
    if (!live.properties['配布区分'] || (live.properties['書類番号'] || {}).maxLength !== '20') {
      if ((live.properties['書類番号'] || {}).maxLength !== '20') await k.api('PUT', '/k/v1/preview/app/form/fields', { app: APP, properties: { 書類番号: { type: 'SINGLE_LINE_TEXT', code: '書類番号', label: '書類番号', maxLength: '20' } } });
      if (!pre.properties['配布区分'] && !live.properties['配布区分']) await k.api('POST', '/k/v1/preview/app/form/fields', { app: APP, properties: { 配布区分: { type: 'DROP_DOWN', code: '配布区分', label: '配布区分', required: false, options: { '配布可': { label: '配布可', index: '0' }, 'NDA限定': { label: 'NDA限定', index: '1' }, '社内限定': { label: '社内限定', index: '2' } } } } });
      const l2 = await k.api('GET', '/k/v1/preview/app/form/layout', { app: APP });
      l2.layout = l2.layout.filter((x) => !(x.type === 'ROW' && x.fields.some((f) => f.code === '配布区分')));
      const at = l2.layout.findIndex((x) => x.type === 'ROW' && x.fields.some((f) => f.code === '書類番号'));
      l2.layout.splice(at + 1, 0, { type: 'ROW', fields: [{ type: 'DROP_DOWN', code: '配布区分', size: { width: '200' } }] });
      await k.api('PUT', '/k/v1/preview/app/form/layout', { app: APP, layout: l2.layout });
      await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: APP }] });
      await k.waitDeploy(APP, '書類マスタ');
    }
    const after = await k.api('GET', '/k/v1/app/form/fields', { app: APP }), lay3 = await k.api('GET', '/k/v1/app/form/layout', { app: APP });
    const bad = [];
    if (!after.properties['配布区分']) bad.push('配布区分がありません');
    Object.keys(live.properties).forEach((c) => { if (!after.properties[c]) bad.push('既存の項目「' + c + '」が消えています'); });
    const rows = (l) => l.layout.filter((x) => x.type === 'ROW').length;
    if (rows(lay3) !== rows(lay) + (live.properties['配布区分'] ? 0 : 1)) bad.push('行の数が、想定と違います（' + rows(lay) + ' → ' + rows(lay3) + '）');
    if (bad.length) { bad.forEach((b) => console.error('NG ' + b)); process.exit(4); }
    console.log('OK 欄を整えました（' + Object.keys(live.properties).length + ' → ' + Object.keys(after.properties).length + '項目、' + rows(lay) + ' → ' + rows(lay3) + '行。書類番号は20字まで）');
  }

  if (has('records')) {
    for (let i = 0; i < upd.length; i += 100) {
      await k.api('PUT', '/k/v1/records', { app: APP, records: upd.slice(i, i + 100).map((u) => ({ id: u.id, record: { 書類番号: { value: u.no }, 配布区分: { value: u.dist || null } } })) });
    }
    console.log('OK 既存 ' + upd.length + '件の番号・配布区分を入れました');
    // 雛形に {{ }} が無い既存の書類は、Boxの正本に差し替える
    const lib = require('../customize/doc-button.js'), JSZip = require('jszip');
    let swapped = 0;
    for (const u of upd) {
      const x = byNo[u.no]; if (!x || !['docx', 'xlsx', 'pptx'].includes(x.ext)) continue;
      const r = all.records.find((q) => q.$id.value === u.id), f0 = ((r['雛形'] || {}).value || [])[0]; if (!f0) continue;
      const ext0 = (f0.name.match(/\.([^.]+)$/) || [])[1];
      if (ext0 !== x.ext) continue;
      if (f0.name === x.fname) continue;   // すでに、Boxの正本が先頭にある（二重に差し替えない）
      const txt = await lib.extractText(JSZip, Buffer.from(await k.download(f0.fileKey)), ext0);
      if (/\{\{[^{}]+\}\}/.test(txt)) continue;   // 自動入力のある雛形は、そのまま
      const pr = spawnSync('node', [path.join(__dirname, 'template-tool.js'), 'put', '--doc=' + u.name, '--file=' + x.file, '--yes'], { encoding: 'utf8' });
      console.log('   差し替え（{{ }}なし）：' + u.name + ' → ' + x.fname + (pr.status === 0 ? '' : '  ※失敗'));
      swapped++;
    }
    console.log('OK 雛形を正本に差し替えた ' + swapped + '件');
    // 新しい書類を追加
    const order = {}; let n = 0;
    for (const x of toAdd) {
      const st = STAGE[x.code] || '受注・準備', i = (order[st] = (order[st] || 0) + 1);
      const key = await k.upload(fs.readFileSync(x.file), x.fname, 'application/octet-stream');
      await k.api('POST', '/k/v1/record', { app: APP, record: {
        書類名: { value: x.name }, 工程: { value: st }, 形式: { value: KIND[x.ext] || 'Word' }, 対象: { value: '共通' }, 見られる立場: { value: AUD[x.dist] },
        表示順: { value: String(BASE[st] + 30 + i) }, 使用状況: { value: '使用中' }, 書類番号: { value: x.no }, 配布区分: { value: x.dist },
        説明: { value: 'Boxの台帳の書類（' + x.dist + '）。雛形のまま使う（案件の情報は入りません）' }, 雛形: { value: [{ fileKey: key }] } } });
      n++;
    }
    console.log('OK 新しい書類を ' + n + '件追加しました');
  }
}
main().catch((e) => { console.error(e.message || JSON.stringify(e)); process.exit(1); });
