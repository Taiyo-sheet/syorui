'use strict';
/*
 * 「書類マスタ」「案件書類」を、指定したスペースの中に作って公開し、書類マスタに初期データを入れる。
 *
 *   node scripts/create-apps.js --space=スペース番号            … 何をするかを表示するだけ（何も作らない）
 *   node scripts/create-apps.js --space=スペース番号 --yes      … 実際に作る
 *
 * 守ること
 *   ・同じ名前のアプリがすでにあれば、何も作らずに止まる（削除はしない）
 *   ・途中で失敗しても、作ったアプリは消さない（番号を表示するので、画面で確認して整理する）
 *   ・認証情報は環境変数（KINTONE_DOMAIN / KINTONE_USER / KINTONE_PASSWORD）から読む
 *
 * アプリの作成・項目追加・公開は、APIトークンでは使えません（ユーザー認証が必要）。
 */
const fs = require('fs');
const path = require('path');
const { client, parseCsv } = require('./lib');

const STAGES = ['情報取得・現状確認', '打合せ・見積・提案', '受注・準備', '施工前の書類', '施工・試験', '竣工', '請求・原価管理'];
const AUDS = ['営業', 'TS計画推進', '工事部', '外部パートナー', '施工店'];
const TARGETS = ['共通', '橋梁', '一般道路'];   // 書類の対象。仕様などが橋梁と一般道路で違う書類は、対象別に行を分ける
const opts = (list) => Object.fromEntries(list.map((n, i) => [n, { label: n, index: String(i) }]));

const MASTER_PROPS = {
  '工程': { type: 'DROP_DOWN', code: '工程', label: '工程', required: true, options: opts(STAGES) },
  '表示順': { type: 'NUMBER', code: '表示順', label: '表示順' },
  '書類名': { type: 'SINGLE_LINE_TEXT', code: '書類名', label: '書類名', required: true },
  '説明': { type: 'SINGLE_LINE_TEXT', code: '説明', label: '説明' },
  '形式': { type: 'DROP_DOWN', code: '形式', label: '形式', options: opts(['Excel', 'Word', 'PowerPoint', 'PDF']) },
  '雛形': { type: 'FILE', code: '雛形', label: '雛形' },
  '見られる立場': { type: 'CHECK_BOX', code: '見られる立場', label: '見られる立場', options: opts(AUDS) },
  '対象': { type: 'DROP_DOWN', code: '対象', label: '対象', defaultValue: '共通', options: opts(TARGETS) },
  '使用状況': { type: 'DROP_DOWN', code: '使用状況', label: '使用状況', required: true, defaultValue: '使用中', options: opts(['使用中', '廃止']) },
};
const DOCS_PROPS = {
  '案件ID': { type: 'NUMBER', code: '案件ID', label: '案件ID', required: true },
  '書類名': { type: 'SINGLE_LINE_TEXT', code: '書類名', label: '書類名' },
  '工程': { type: 'SINGLE_LINE_TEXT', code: '工程', label: '工程' },
  '版': { type: 'NUMBER', code: '版', label: '版' },
  '添付ファイル': { type: 'FILE', code: '添付ファイル', label: '添付ファイル' },
  'AIチェック結果': { type: 'MULTI_LINE_TEXT', code: 'AIチェック結果', label: 'AIチェック結果' },
};

function loadSeed() {
  const rows = parseCsv(fs.readFileSync(path.join(__dirname, '..', 'data', '書類マスタ_初期データ.csv'), 'utf8'));
  rows.forEach((r, i) => {
    const where = (i + 2) + '行目（' + r['書類名'] + '）';
    if (!STAGES.includes(r['工程'])) throw new Error('初期データの' + where + '：工程「' + r['工程'] + '」は選択肢にありません');
    if (isNaN(Number(r['表示順']))) throw new Error('初期データの' + where + '：表示順が数値ではありません');
    if (r['対象'] && !TARGETS.includes(r['対象'])) throw new Error('初期データの' + where + '：対象「' + r['対象'] + '」は選択肢にありません');
    r['見られる立場'].split('|').filter(Boolean).forEach((a) => { if (!AUDS.includes(a)) throw new Error('初期データの' + where + '：立場「' + a + '」は選択肢にありません'); });
  });
  return rows.map((r) => ({
    '工程': { value: r['工程'] }, '表示順': { value: String(r['表示順']) }, '書類名': { value: r['書類名'] },
    '説明': { value: r['説明'] }, '形式': { value: r['形式'] },
    '見られる立場': { value: r['見られる立場'].split('|').filter(Boolean) }, '対象': { value: r['対象'] || '共通' }, '使用状況': { value: r['使用状況'] || '使用中' },
  }));
}

async function main() {
  const args = process.argv.slice(2);
  const spaceId = Number((args.find((a) => a.startsWith('--space=')) || '').split('=')[1] || process.env.KINTONE_SPACE_ID);
  const go = args.includes('--yes');
  if (!spaceId) { console.error('スペース番号を --space=数字 で指定してください（スペースのURLの /space/数字 の部分）'); process.exit(1); }

  const seed = loadSeed();
  const k = client();
  const sp = await k.api('GET', '/k/v1/space', { id: spaceId });
  if (!sp.defaultThread) throw new Error('このスペースにはアプリを作れません（スレッドが見つかりません）');
  console.log('スペース：「' + sp.name + '」（番号 ' + spaceId + '）');

  for (const name of ['書類マスタ', '案件書類']) {
    const found = await k.findApps(name);
    if (found.length) {
      console.error('「' + name + '」がすでにあります（アプリ番号 ' + found.map((a) => a.appId).join('、') + '）。何も作らずに止めました。');
      console.error('削除はしません。不要なら、キントーンの画面から手で削除してから、もう一度実行してください。');
      process.exit(2);
    }
  }
  console.log('同じ名前のアプリはありません。');
  console.log('作るもの：書類マスタ（項目 ' + Object.keys(MASTER_PROPS).length + '、初期データ ' + seed.length + '件）／案件書類（項目 ' + Object.keys(DOCS_PROPS).length + '）');
  if (!go) { console.log('\n※ これは確認だけです。実際に作るには、同じコマンドの最後に --yes を付けてください。'); return; }

  const created = [];
  async function makeApp(name, props) {
    const a = await k.api('POST', '/k/v1/preview/app', { name, space: spaceId, thread: Number(sp.defaultThread) });
    created.push(name + '（アプリ番号 ' + a.app + '）');
    console.log(name + ' の下書きを作りました（アプリ番号 ' + a.app + '）');
    await k.api('POST', '/k/v1/preview/app/form/fields', { app: a.app, properties: props });
    await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: a.app }] });
    await k.waitDeploy(a.app, name);
    console.log(name + ' を公開しました');
    return a.app;
  }

  let masterId, docsId;
  try {
    masterId = await makeApp('書類マスタ', MASTER_PROPS);
    docsId = await makeApp('案件書類', DOCS_PROPS);
    await k.api('POST', '/k/v1/records', { app: masterId, records: seed });
    console.log('書類マスタに ' + seed.length + '件を登録しました');
  } catch (e) {
    console.error('\n途中で止まりました：' + e.message);
    if (created.length) console.error('ここまでに作ったもの：' + created.join('、') + '\n（自動では消しません。画面で確認し、不要なら手で削除してから、もう一度実行してください）');
    process.exit(3);
  }

  // 作成後の確認
  const problems = [];
  for (const [id, name, props] of [[masterId, '書類マスタ', MASTER_PROPS], [docsId, '案件書類', DOCS_PROPS]]) {
    const f = await k.api('GET', '/k/v1/app/form/fields', { app: id });
    Object.keys(props).forEach((c) => {
      if (!f.properties[c]) problems.push(name + '：項目「' + c + '」がありません');
      else if (f.properties[c].type !== props[c].type) problems.push(name + '：項目「' + c + '」の種類が違います（' + f.properties[c].type + '）');
    });
  }
  const recs = await k.api('GET', '/k/v1/records', { app: masterId, query: 'limit 500', totalCount: true });
  if (Number(recs.totalCount) !== seed.length) problems.push('書類マスタの件数が ' + recs.totalCount + ' 件です（' + seed.length + '件のはず）');
  const sp2 = await k.api('GET', '/k/v1/app', { id: masterId });
  if (String(sp2.spaceId) !== String(spaceId)) problems.push('書類マスタが指定のスペースに入っていません（スペース ' + sp2.spaceId + '）');

  console.log('\n--- 確認結果 ---');
  if (problems.length) { problems.forEach((p) => console.log('NG ' + p)); process.exit(4); }
  console.log('OK 項目・件数・スペースとも想定どおりです');
  console.log('書類マスタ：' + k.base + '/k/' + masterId + '/');
  console.log('案件書類　：' + k.base + '/k/' + docsId + '/');
  console.log('次：書類マスタの「雛形」欄に、Excel・Word・PowerPointの雛形を添付してください。');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
