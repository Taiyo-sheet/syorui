'use strict';
/*
 * 案件くんの画面（「書類を作る」ボタン→パネル→作成して保存）を、本物のブラウザ（Chromium）で動かして確かめる。
 * キントーン本体の代わりに、模擬キントーンと、kintone オブジェクトの簡易版を使う。
 * Chromium が見つからない環境では、このテストは自動でスキップされる。
 */
const test = require('node:test');
const fs = require('fs');
const assert = require('node:assert');
const JSZip = require('jszip');
const { createMock } = require('./mock-kintone');
const { build } = require('../scripts/make-test-templates');

const CHROMIUM = ['/opt/pw-browsers/chromium', process.env.CHROMIUM_PATH].find((p) => p && fs.existsSync(p));
let chromium; try { chromium = require('playwright-core').chromium; } catch (e) { chromium = null; }

// ページに差し込む、kintone オブジェクトの簡易版（PC・モバイルで共通に使う部分）
const FAKE_KINTONE = `
  var FIELDS = {
    c_name: { type: 'SINGLE_LINE_TEXT', code: 'c_name', label: '案件名' },
    c_addr: { type: 'SINGLE_LINE_TEXT', code: 'c_addr', label: '現場住所' },
    c_amount: { type: 'NUMBER', code: 'c_amount', label: '金額', digit: true },
    c_no: { type: 'NUMBER', code: 'c_no', label: '案件番号' },
    書類保管: { type: 'SUBTABLE', code: '書類保管', fields: { 書類保管_書類名: {}, 書類保管_工程: {}, 書類保管_版: {}, 書類保管_作成日: {}, 書類保管_ファイル: {}, 書類保管_承認状況: {}, 書類保管_承認者: {}, 書類保管_承認日時: {}, 書類保管_PDF: {} } }
  };
  function q(params) { var o = []; Object.keys(params || {}).forEach(function (k) { var v = params[k]; if (Array.isArray(v)) v.forEach(function (x, i) { o.push(encodeURIComponent(k + '[' + i + ']') + '=' + encodeURIComponent(x)); }); else if (v != null) o.push(encodeURIComponent(k) + '=' + encodeURIComponent(v)); }); return o.join('&'); }
  function api(url, method, params) {
    var get = method === 'GET';
    return fetch(url + (get ? '?' + q(params) : ''), { method: method, headers: { 'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/json' }, body: get ? undefined : JSON.stringify(params) })
      .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw j; return j; }); });
  }
  api.url = function (p) { return p + '.json'; };
  var handlers = [];
  var header = function () { return document.getElementById('hdr'); };
  var rec = function () { return { getId: function () { return 1; }, getHeaderMenuSpaceElement: header, getSpaceElement: function (id) { return id === 'doc_button_space' ? document.getElementById('tabspace') : null; }, getFieldElement: function (c) { return c === '書類保管' ? document.getElementById('tabtable') : c === '安全書類格納完了チェック' ? document.getElementById('anchor') : null; } }; };
  window.kintone = {
    api: api,
    getRequestToken: function () { return 'tok'; },
    events: { on: function (names, fn) { handlers.push(fn); } },
    app: { getId: function () { return 575; }, getFormFields: function () { return Promise.resolve(FIELDS); }, record: rec() },
    mobile: { app: { getId: function () { return 575; }, getHeaderSpaceElement: header, record: { getId: function () { return 1; } } } }
  };
  window.__fire = function (record) { handlers.forEach(function (h) { h({ record: record }); }); };
`;
const RECORD = {
  c_name: { type: 'SINGLE_LINE_TEXT', value: '東京Aビル 改修工事' }, c_addr: { type: 'SINGLE_LINE_TEXT', value: '東京都千代田区1-1' },
  c_amount: { type: 'NUMBER', value: '1234567' }, c_no: { type: 'NUMBER', value: '2026' },
};

// 案件1の「書類保管」表の行を、読みやすい形で返す
const stored = (mock) => mock.state.records['575'][0]['書類保管'].value.map((r) => ({
  name: r.value['書類保管_書類名'].value, stage: r.value['書類保管_工程'].value, version: r.value['書類保管_版'].value, file: r.value['書類保管_ファイル'].value[0],
}));

async function seed(mock) {
  const tpl = await build();
  mock.addApp('575', '案件くん');
  mock.addRecord('575', { 書類保管: { type: 'SUBTABLE', value: [] } });   // 案件くんの「書類保管」表（案件1）
  mock.addApp('10', '書類マスタ');
  const k1 = mock.addFile(tpl.docx, '議事次第.docx');
  const k2 = mock.addFile(tpl.xlsx, '現状確認.xlsx');
  // {{ }} を1つも含まない雛形（標準の手引きのような、使い回す書類）。
  // ※ ファイル名は英数字：このテスト環境の Chromium は、日本語名のダウンロードを「download」と報告するため
  const zp = await JSZip.loadAsync(tpl.docx);
  zp.file('word/document.xml', '<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>共通の手引き</w:t></w:r></w:p></w:body></w:document>');
  zp.file('word/header1.xml', '<w:hdr xmlns:w="x"><w:p><w:r><w:t>ヘッダー</w:t></w:r></w:p></w:hdr>');
  const k3 = mock.addFile(await zp.generateAsync({ type: 'nodebuffer' }), 'standard-guide.docx');
  const mk = (name, stage, order, auds, files) => mock.addRecord('10', {
    書類名: { value: name }, 工程: { value: stage }, 表示順: { value: String(order) }, 説明: { value: name + 'の説明' },
    見られる立場: { value: auds }, 使用状況: { value: '使用中' }, 雛形: { value: files },
  });
  mk('現状確認シート', '情報取得・現状確認', 110, ['営業', 'TS計画推進'], [{ fileKey: k2, name: '現状確認.xlsx' }]);
  mk('初期リスク抽出シート', '情報取得・現状確認', 130, ['営業'], []);
  mk('打合せ議事次第', '打合せ・見積・提案', 210, ['営業', '施工店'], [{ fileKey: k1, name: '議事次第.docx' }]);
  mk('標準の手引き', '打合せ・見積・提案', 215, ['営業'], [{ fileKey: k3, name: 'standard-guide.docx' }]);
  mock.addRecord('10', { 書類名: { value: '廃止した書類' }, 工程: { value: '竣工' }, 表示順: { value: '999' }, 使用状況: { value: '廃止' }, 見られる立場: { value: [] }, 雛形: { value: [] } });
}

async function openPage(browser, url, path, record) {
  const ctx = await browser.newContext();
  const page = await ctx.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('dialog', (d) => d.accept());
  await page.addInitScript(FAKE_KINTONE);
  await page.goto(url + path);
  await page.evaluate((r) => window.__fire(r), record || RECORD);
  return { page, errors, ctx };
}

const skip = !chromium || !CHROMIUM ? 'Chromium が見つからないためスキップ' : false;

test('画面：ボタン表示→工程・立場の絞り込み→作成して保存→最新vN→v2作成', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/__page');
    assert.strictEqual(await page.locator('[data-doc-hub]').count(), 1, '「書類を作る」ボタンが1つ出る');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：東京Aビル 改修工事').waitFor();

    // 工程の選択肢は書類マスタから（廃止は出ない）／立場の選択肢も書類マスタから
    const stages = await page.locator('select').nth(0).locator('option').evaluateAll((os) => os.map((o) => o.value));
    assert.deepStrictEqual(stages, ['情報取得・現状確認', '打合せ・見積・提案']);
    // 工程の選択肢に「作成済 x/y」（雛形のまま使う書類・雛形が未登録の書類は数えない）
    await page.waitForFunction(() => /打合せ・見積・提案（作成済 0\/1）/.test(document.querySelectorAll('select')[0].textContent));
    assert.match(await page.locator('select').nth(0).textContent(), /情報取得・現状確認（作成済 0\/1）/);
    const auds = await page.locator('select').nth(1).locator('option').allTextContents();
    assert.deepStrictEqual(auds, ['すべての立場', '営業', 'TS計画推進', '施工店']);

    // 雛形が未登録の書類は押せない
    assert.strictEqual(await page.getByRole('button', { name: '雛形が未登録' }).isDisabled(), true);
    // 立場で絞り込む（TS計画推進 → 現状確認シートだけ）
    await page.locator('select').nth(1).selectOption('TS計画推進');
    assert.strictEqual(await page.getByText('初期リスク抽出シート').count(), 0);
    await page.locator('select').nth(1).selectOption('');

    // 工程を切り替えて、作成して保存
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    // 作成すると、保存と同時に、その場でダウンロードされる
    const [dl1] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).click()]);
    await page.getByText(/保存してダウンロードしました：東京Aビル 改修工事_打合せ議事次第_v1\.docx/).waitFor();
    const got = []; for await (const c of await dl1.createReadStream()) got.push(c);
    const docs = stored(mock);
    assert.strictEqual(docs.length, 1);
    assert.strictEqual(docs[0].name, '打合せ議事次第');
    assert.strictEqual(docs[0].version, '1');
    assert.strictEqual(docs[0].stage, '打合せ・見積・提案');
    const f = docs[0].file;
    assert.ok(Buffer.concat(got).equals(mock.state.files[f.fileKey].buf), 'ダウンロードされた書類と、案件に保存された書類が同じ');
    const xml = await (await JSZip.loadAsync(mock.state.files[f.fileKey].buf)).file('word/document.xml').async('string');
    assert.match(xml.replace(/<[^>]*>/g, ''), /案件名：東京Aビル 改修工事/);
    assert.match(xml.replace(/<[^>]*>/g, ''), /打合せ議事次第 v1 /);
    // 置き換えできなかった {{ }} を知らせる
    assert.match(await page.locator('div', { hasText: '置き換えできずに残った項目' }).last().textContent(), /存在しない項目/);

    // 開き直すと「最新 v1・開く」、次は v2（確認ダイアログは自動で承諾）
    await page.reload();
    await page.evaluate((r) => window.__fire(r), RECORD);
    await page.locator('[data-doc-hub]').click();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.getByText('最新 v1').waitFor();
    assert.match(await page.locator('select').nth(0).textContent(), /打合せ・見積・提案（作成済 1\/1）/);
    // 「開く」で、保存した書類（v1）が取り出せる
    const [dlOpen] = await Promise.all([page.waitForEvent('download'), page.getByRole('link', { name: '開く' }).click()]);
    const again = []; for await (const c of await dlOpen.createReadStream()) again.push(c);
    assert.ok(Buffer.concat(again).equals(mock.state.files[f.fileKey].buf), '「開く」で、保存した書類が取り出せる');
    await page.getByRole('button', { name: 'v2 を作成' }).click();
    await page.getByText(/_v2\.docx/).waitFor();
    assert.deepStrictEqual(stored(mock).map((r) => r.version), ['1', '2'], 'v1 の行は残り、v2 が足される');
    assert.deepStrictEqual(errors, [], 'ブラウザのエラーなし');
  } finally { await browser.close(); await mock.close(); }
});

test('画面：まとめて作成（雛形のある書類だけ）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.getByRole('button', { name: 'この工程をまとめて作成' }).click();
    await page.getByText('1件を保存しました').waitFor();   // 雛形が未登録の1件は対象外
    assert.deepStrictEqual(stored(mock).map((r) => r.name), ['現状確認シート']);
    const f = stored(mock)[0].file;
    assert.strictEqual(f.name, '東京Aビル 改修工事_現状確認シート_v1.xlsx');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面（モバイル）：ボタンが出て、作成して保存できる', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/k/m/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.getByRole('button', { name: '作成して保存' }).click();
    await page.getByText(/ダウンロードしました/).waitFor();
    assert.strictEqual(stored(mock).length, 1);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：{{ }}のない雛形は「雛形を開く」（そのまま保存・案件には保存しない）。まとめて作成の対象外', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/__page');
    const dialogs = [];
    page.on('dialog', (d) => dialogs.push(d.message()));
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    // 判定が終わると、{{ }}ありは「作成して保存」、なしは「雛形を開く」になる
    await page.getByRole('button', { name: '作成して保存' }).waitFor();
    await page.getByRole('button', { name: '雛形を開く' }).waitFor();
    await page.getByText('雛形のまま使う書類（案件の情報は入れません）').waitFor();
    // 雛形を開く：ダウンロードされる（元のファイル名・中身そのまま）。案件には保存されない
    const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '雛形を開く' }).click()]);
    assert.strictEqual(dl.suggestedFilename(), 'standard-guide.docx');
    const saved = await dl.createReadStream(); const chunks = []; for await (const c of saved) chunks.push(c);
    assert.ok(Buffer.concat(chunks).equals(mock.state.files['fk3'].buf), '雛形がそのまま渡る');
    await page.getByText(/雛形のまま保存しました：standard-guide\.docx/).waitFor();
    assert.strictEqual(stored(mock).length, 0, '案件には保存しない');
    // まとめて作成：{{ }}のある書類だけ（雛形のままの1件は対象外）
    await page.getByRole('button', { name: 'この工程をまとめて作成' }).click();
    await page.getByText('1件を保存しました').waitFor();
    assert.ok(dialogs.some((m) => /1件の書類を作成します/.test(m) && /雛形のまま使う 1件/.test(m)), dialogs.join('|'));
    assert.deepStrictEqual(stored(mock).map((r) => r.name), ['打合せ議事次第']);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：対象（橋梁／一般道路）で、仕様書を出し分ける。「共通」はどちらにも出る', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const k = mock.addFile(tpl.docx, 'spec.docx');
    const doc = (name, target) => mock.addRecord('10', {
      書類名: { value: name }, 工程: { value: '施工前の書類' }, 表示順: { value: '400' }, 説明: { value: '' }, 見られる立場: { value: ['工事部'] },
      対象: { value: target }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k, name: 'spec.docx' }] },
    });
    doc('仕様書（橋梁）', '橋梁'); doc('仕様書（一般道路）', '一般道路'); doc('作業手順書', '共通');
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('施工前の書類');
    // 対象の選択が出る（立場の次）。案件名「東京Aビル 改修工事」には橋がないので、初期は「一般道路」
    const targets = await page.locator('select').nth(2).locator('option').evaluateAll((os) => os.map((o) => o.textContent));
    assert.deepStrictEqual(targets, ['すべての対象', '橋梁', '一般道路']);
    assert.strictEqual(await page.locator('select').nth(2).inputValue(), '一般道路');
    await page.getByText('対象は、案件名から「一般道路」と判断して選びました').waitFor();
    assert.strictEqual(await page.getByText('仕様書（橋梁）', { exact: true }).count(), 0);
    // 「すべての対象」にすると、全部見える
    await page.locator('select').nth(2).selectOption('');
    await page.getByText('仕様書（橋梁）', { exact: true }).waitFor();
    assert.strictEqual(await page.getByText('仕様書（一般道路）', { exact: true }).count(), 1);
    assert.strictEqual(await page.getByText('作業手順書').count() >= 1, true);
    // 橋梁を選ぶ → 橋梁用と共通だけ
    await page.locator('select').nth(2).selectOption('橋梁');
    assert.strictEqual(await page.getByText('仕様書（橋梁）', { exact: true }).count(), 1);
    assert.strictEqual(await page.getByText('仕様書（一般道路）', { exact: true }).count(), 0);
    assert.strictEqual(await page.getByText('作業手順書').count() >= 1, true);
    await page.waitForFunction(() => /施工前の書類（作成済 0\/2）/.test(document.querySelectorAll('select')[0].textContent));
    // 一般道路を選ぶ → 一般道路用と共通だけ
    await page.locator('select').nth(2).selectOption('一般道路');
    assert.strictEqual(await page.getByText('仕様書（橋梁）', { exact: true }).count(), 0);
    assert.strictEqual(await page.getByText('仕様書（一般道路）', { exact: true }).count(), 1);
    // 一般道路用の仕様書を作成 → 書類名ごとに版が分かれて保存される
    await page.getByRole('button', { name: '作成して保存' }).first().waitFor();
    assert.strictEqual(await page.getByText('【一般道路用】').count(), 1, '説明の行に対象を表示');
    await page.locator('div[style*="border-top"]', { hasText: '仕様書（一般道路）' }).getByRole('button').click();
    await page.getByText(/_仕様書（一般道路）_v1\.docx/).waitFor();
    assert.deepStrictEqual(stored(mock).map((r) => r.name), ['仕様書（一般道路）']);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：「共通」しかなければ、対象の選択は出ない', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    assert.strictEqual(await page.locator('select').count(), 2, '工程・立場のみ');
  } finally { await browser.close(); await mock.close(); }
});

test('画面：案件名に「橋」があれば、対象は自動で「橋梁」になる（手で変えられる）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const k = mock.addFile(tpl.docx, 'spec.docx');
    const doc = (name, target) => mock.addRecord('10', {
      書類名: { value: name }, 工程: { value: '施工前の書類' }, 表示順: { value: '400' }, 説明: { value: '' }, 見られる立場: { value: ['工事部'] },
      対象: { value: target }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k, name: 'spec.docx' }] },
    });
    doc('仕様書（橋梁）', '橋梁'); doc('仕様書（一般道路）', '一般道路');
    const bridge = Object.assign({}, RECORD, { c_name: { type: 'SINGLE_LINE_TEXT', value: '第二多摩川橋 補修工事' } });
    const { page, errors } = await openPage(browser, url, '/__page', bridge);
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('施工前の書類');
    assert.strictEqual(await page.locator('select').nth(2).inputValue(), '橋梁');
    await page.getByText('対象は、案件名から「橋梁」と判断して選びました').waitFor();
    assert.strictEqual(await page.getByText('仕様書（橋梁）', { exact: true }).count(), 1);
    assert.strictEqual(await page.getByText('仕様書（一般道路）', { exact: true }).count(), 0);
    await page.locator('select').nth(2).selectOption('一般道路');   // 手で変えられる
    assert.strictEqual(await page.getByText('仕様書（橋梁）', { exact: true }).count(), 0);
    assert.strictEqual(await page.getByText('仕様書（一般道路）', { exact: true }).count(), 1);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：全工程をまとめて作成（未作成だけ）／2回目は「すべて作成済み」', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.getByRole('button', { name: '全工程をまとめて作成' }).click();
    await page.getByText(/2件を保存しました/).waitFor();
    assert.deepStrictEqual(stored(mock).map((r) => r.name).sort(), ['現状確認シート', '打合せ議事次第'].sort(), '雛形のある書類だけ（標準の手引き・雛形なしは作らない）');
    await page.getByRole('button', { name: '全工程をまとめて作成' }).click();
    await page.getByText(/すべて作成済みか、雛形のまま使うもの/).waitFor();
    assert.strictEqual(stored(mock).length, 2, '2回目は作らない（版が増えない）');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：直したファイルを、次の版として登録できる', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).click()]);
    await page.getByText(/_v1\.docx/).waitFor();
    await page.reload(); await page.evaluate((r) => window.__fire(r), RECORD);
    await page.locator('[data-doc-hub]').click();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.getByText('最新 v1').waitFor();
    await page.locator('input[data-revision="打合せ議事次第"]').setInputFiles({ name: '直した議事次第.docx', mimeType: 'application/octet-stream', buffer: tpl.docx });
    await page.getByText(/v2 として登録しました：東京Aビル 改修工事_打合せ議事次第_v2\.docx/).waitFor();
    const rows = stored(mock);
    assert.deepStrictEqual(rows.map((r) => r.version), ['1', '2'], 'v1 は残り、v2 が足される');
    assert.ok(mock.state.files[rows[1].file.fileKey].buf.equals(tpl.docx), '直したファイルの中身が、そのまま保存される');
    await page.getByText('最新 v2').waitFor();
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：案件くんの値が空のまま入る項目を、作成後に知らせる（自動では画面を更新しない）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const blank = Object.assign({}, RECORD, { c_addr: { type: 'SINGLE_LINE_TEXT', value: '' } });
    const { page, errors } = await openPage(browser, url, '/__page', blank);
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).click()]);
    await page.getByText(/案件くんの値が空のまま入った項目.*\{\{現場住所\}\}/).waitFor();
    await page.getByText('保存はできています').waitFor();
    assert.strictEqual(stored(mock).length, 1, '保存はされている');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：議事録（{{打合せ日}} {{議事録メモ}} のある雛形）は、入力欄が出て、箇条書きが整って入る', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const z = await JSZip.loadAsync(tpl.docx);
    z.file('word/document.xml', '<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>議事録 {{案件名}}</w:t></w:r></w:p><w:p><w:r><w:t>日時：{{打合せ日}}　場所：{{場所}}</w:t></w:r></w:p><w:p><w:r><w:t>{{議事録メモ}}</w:t></w:r></w:p></w:body></w:document>');
    const k = mock.addFile(await z.generateAsync({ type: 'nodebuffer' }), 'gijiroku.docx');
    mock.addRecord('10', { 書類名: { value: '議事録' }, 工程: { value: '打合せ・見積・提案' }, 表示順: { value: '205' }, 説明: { value: '' }, 見られる立場: { value: ['営業'] }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k, name: 'gijiroku.docx' }] } });
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.locator('div[style*="border-top"]', { hasText: '議事録' }).getByRole('button', { name: '作成して保存' }).click();
    // 入力欄：{{ }} にある項目（打合せ日・場所・議事録メモ）だけ出る
    await page.locator('[data-ask="議事録メモ"]').waitFor();
    assert.strictEqual(await page.locator('[data-ask="出席者"]').count(), 0, '雛形にない項目の欄は出ない');
    await page.locator('[data-ask="打合せ日"]').fill('2026-10-21');
    await page.locator('[data-ask="場所"]').fill('現場事務所');
    await page.locator('[data-ask="議事録メモ"]').fill('- 見積の再提出\n\n● 工程表の確認\n・来週、現地調査');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).last().click()]);
    const got = []; for await (const c of await dl.createReadStream()) got.push(c);
    const xml = await (await JSZip.loadAsync(Buffer.concat(got))).file('word/document.xml').async('string');
    const text = xml.replace(/<w:br\/>/g, '\n').replace(/<[^>]*>/g, '');
    assert.match(text, /議事録 東京Aビル 改修工事/);
    assert.match(text, /日時：2026年10月21日　場所：現場事務所/);
    assert.match(text, /・見積の再提出\n・工程表の確認\n・来週、現地調査/, '記号がそろい、空行は消える');
    assert.strictEqual(stored(mock).length, 1);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：案件管理アプリを開くボタン（PC用）は、右下の丸いボタン。近づけると広がり、押すと別タブでURLを開く', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    const ctx = await browser.newContext(); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.addInitScript(FAKE_KINTONE);
    await page.addInitScript(() => { window.__opened = []; window.open = function (u) { window.__opened.push(u); return null; }; });
    await page.goto(url + '/__page');
    await page.evaluate(() => { var s = document.createElement('script'); s.src = '/__lib/taiyo-open-app.js'; document.head.appendChild(s); return new Promise((r) => { s.onload = r; }); });
    await page.evaluate((r) => window.__fire(r), RECORD);
    await page.evaluate((r) => window.__fire(r), RECORD);   // 2回開いても、ボタンは1つ
    const btn = page.locator('#taiyo-open-app-btn');
    assert.strictEqual(await btn.count(), 1);
    const box = await btn.boundingBox();
    assert.ok(box.width < 70, '最初は丸い小さなボタン：' + box.width);
    await btn.hover();
    await page.waitForFunction(() => document.getElementById('taiyo-open-app-btn').getBoundingClientRect().width > 150);
    assert.match(await btn.textContent(), /案件管理アプリを開く/);
    await btn.click();
    assert.deepStrictEqual(await page.evaluate(() => window.__opened), ['https://anken-kanri-1.onrender.com/']);
    assert.deepStrictEqual(errors, []);
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

test('画面：round-buttons.js は、元のボタンを丸くする（押したときの動き・リンク先・色は元のまま）。あとから出るボタンも丸くする', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    const ctx = await browser.newContext({ viewport: { width: 900, height: 600 } }); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.addInitScript(() => { window.__clicks = []; });
    await page.goto(url + '/__page');
    await page.evaluate(() => new Promise((r) => { const s = document.createElement('script'); s.src = '/__lib/round-buttons.js'; s.onload = r; document.head.appendChild(s); }));
    // 元のJSが、あとから出すボタン（見積くん＝ボタン、発注くん＝リンク）
    await page.evaluate(() => {
      const b = document.createElement('button'); b.id = 'open-mitsumori-btn'; b.type = 'button'; b.textContent = '🧮 見積作成システムを開く';
      b.style.cssText = 'position:fixed;right:20px;bottom:20px;padding:11px 18px;background:rgb(27,52,98);color:#fff;border:none;border-radius:8px;font-weight:700;font-size:14px;';
      b.onclick = () => window.__clicks.push('mitsumori'); document.body.appendChild(b);
      const a = document.createElement('a'); a.id = 'ts-hassou-admin-btn'; a.href = 'https://example.com/admin?key=SECRET'; a.target = '_blank'; a.textContent = '📦 発注書を業者へ送付を開く';
      a.style.cssText = 'position:fixed;right:16px;bottom:16px;padding:12px 18px;background:rgb(22,53,122);color:#fff;border-radius:24px;font-size:14px;font-weight:700;text-decoration:none;'; document.body.appendChild(a);
    });
    const m = page.locator('#open-mitsumori-btn'), h = page.locator('#ts-hassou-admin-btn');
    await page.waitForFunction(() => document.getElementById('open-mitsumori-btn').getAttribute('data-round') === '1');
    for (const l of [m, h]) { const b = await l.boundingBox(); assert.ok(b.width < 70 && Math.abs(b.width - b.height) < 20, '丸い小さなボタン：' + b.width); assert.ok(b.x + b.width > 860, '右下'); }
    await h.evaluate((e) => e.style.setProperty('bottom', '120px', 'important'));   // 実際は1アプリに1つ。このテストでは、重ならないようにずらす
    assert.strictEqual(await m.evaluate((e) => getComputedStyle(e).backgroundColor), 'rgb(27, 52, 98)', '色は元のまま');
    assert.strictEqual(await h.getAttribute('href'), 'https://example.com/admin?key=SECRET', 'リンク先・鍵は元のまま');
    await m.hover();
    await page.waitForFunction(() => document.getElementById('open-mitsumori-btn').getBoundingClientRect().width > 150);
    assert.match(await m.textContent(), /見積作成システムを開く/);
    await m.click();
    assert.deepStrictEqual(await page.evaluate(() => window.__clicks), ['mitsumori'], '押したときの動きは元のまま');
    assert.deepStrictEqual(errors, []);
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

test('画面：議事録の全項目（作成者の初期値・出席者は空なら「不明」・空の項目は「なし」）／作成後に「チェック」で指摘が出る', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const z = await JSZip.loadAsync(tpl.docx);
    const P = (t) => '<w:p><w:r><w:t>' + t + '</w:t></w:r></w:p>';
    z.file('word/document.xml', '<w:document xmlns:w="x"><w:body>' + P('■ 工事名称：{{案件名}}') + P('■ 工事場所：{{存在しない場所}}') + P('日時：{{打合せ日}}　作成者：{{作成者}}') + P('出席者：{{出席者}}') + P('議題：{{議題}}') + P('内容：{{議事録メモ}}') + P('決定：{{決定事項}}') + P('ToDo：{{ToDo}}') + P('次回：{{次回予定}}') + '</w:body></w:document>');
    const k = mock.addFile(await z.generateAsync({ type: 'nodebuffer' }), 'gijiroku2.docx');
    mock.addRecord('10', { 書類名: { value: '議事録' }, 工程: { value: '打合せ・見積・提案' }, 表示順: { value: '205' }, 説明: { value: '' }, 見られる立場: { value: ['営業'] }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k, name: 'gijiroku2.docx' }] } });
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.locator('div[style*="border-top"]', { hasText: '議事録' }).getByRole('button', { name: '作成して保存' }).click();
    await page.locator('[data-ask="ToDo"]').waitFor();
    assert.strictEqual(await page.locator('[data-ask="作成者"]').inputValue(), 'くに', '作成者の初期値（開いた人の名前が取れないときは「くに」）');
    await page.locator('[data-ask="議題"]').fill('- 工程の確認\n- 見積の再提出');
    await page.locator('[data-ask="決定事項"]').fill('● 来週着工');
    await page.locator('[data-ask="ToDo"]').fill('10/25まで 太田 見積を出す');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).last().click()]);
    const got = []; for await (const c of await dl.createReadStream()) got.push(c);
    const text = (await (await JSZip.loadAsync(Buffer.concat(got))).file('word/document.xml').async('string')).replace(/<w:br\/>/g, '\n').replace(/<[^>]*>/g, ' ');
    assert.match(text, /作成者：くに/); assert.match(text, /出席者：不明/); assert.match(text, /議題：・工程の確認\s+・見積の再提出/);
    assert.match(text, /決定：・来週着工/); assert.match(text, /ToDo：10\/25まで 太田 見積を出す/); assert.match(text, /次回：なし/);
    assert.match(text, /日時：\d{4}年\d{1,2}月\d{1,2}日/, '打合せ日の初期値は、今日');
    // 「チェック」：置き換えできなかった {{存在しない場所}} を指摘する
    await page.getByText(/置き換えできずに残った項目/).waitFor();
    await page.getByRole('button', { name: '閉じる' }).click();
    await page.locator('[data-doc-hub]').click();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.locator('div[style*="border-top"]', { hasText: '議事録' }).getByRole('button', { name: 'チェック' }).click();
    await page.getByText(/チェック結果.*置き換えられていない項目が残っています：\{\{存在しない場所\}\}/s).waitFor();
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：書類関係の承認（承認者と日時が残る）。新しい版を作ると、承認は外れる。一式承認もできる', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/__page');
    const dialogs = []; page.on('dialog', (d) => dialogs.push(d.message()));
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).click()]);
    await page.getByText(/_v1\.docx/).waitFor();
    await page.reload(); await page.evaluate((r) => window.__fire(r), RECORD);
    await page.locator('[data-doc-hub]').click();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.getByText('未承認').waitFor();
    await page.locator('div[style*="border-top"]', { hasText: '打合せ議事次第' }).getByRole('button', { name: '承認', exact: true }).click();
    await page.getByText(/1件を承認しました/).waitFor();
    assert.ok(dialogs.some((m) => /承認します/.test(m) && /置き換えられていない|チェックの指摘/.test(m)), '承認の前に、チェックの結果を見せる：' + dialogs.join('|'));
    const row = mock.state.records['575'][0]['書類保管'].value[0].value;
    assert.strictEqual(row['書類保管_承認状況'].value, '承認済み');
    assert.ok(row['書類保管_承認者'].value && /^\d{4}-\d{2}-\d{2}T/.test(row['書類保管_承認日時'].value), '承認者と日時が残る');
    assert.strictEqual(row['書類保管_ファイル'].value.length, 1, '承認しても、ファイルは消えない');
    await page.getByText(/承認済み（/).waitFor();
    // 新しい版（v2）を作ると、v2 は未承認（v1 の承認は残る）
    await page.getByRole('button', { name: 'v2 を作成' }).click();
    await page.getByText(/_v2\.docx/).waitFor();
    const rows = mock.state.records['575'][0]['書類保管'].value;
    assert.strictEqual(rows.length, 2);
    assert.strictEqual(rows[0].value['書類保管_承認状況'].value, '承認済み', 'v1 の承認は残る');
    assert.ok(!rows[1].value['書類保管_承認状況'] || rows[1].value['書類保管_承認状況'].value === '', 'v2 は未承認');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：PDF化は、承認済みの書類だけ。サーバー未設定のあいだは押せない。PDFが返ったら、同じ行に保存する', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const ctx = await browser.newContext(); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    let withServer = false;
    await page.route('**/__lib/doc-button.js', async (route) => {
      let src = fs.readFileSync(require('path').join(__dirname, '..', 'customize', 'doc-button.js'), 'utf8');
      src = src.replace(/var PDF_SERVER_URL = '[^']*';/, "var PDF_SERVER_URL = '" + (withServer ? url : '') + "';");
      await route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: src });
    });
    page.on('dialog', (d) => d.accept());
    await page.addInitScript(FAKE_KINTONE);
    const open = async () => {
      await page.goto(url + '/__page'); await page.evaluate((r) => window.__fire(r), RECORD);
      await page.locator('[data-doc-hub]').click(); await page.getByText('工程別の書類：').waitFor();
      await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    };
    await open();
    await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).click()]);
    await page.getByText(/_v1\.docx/).waitFor();
    await open();
    // 未承認の書類には、PDF化がない
    assert.strictEqual(await page.getByRole('button', { name: /PDF化/ }).count(), 0, '承認前は、PDF化を出さない');
    await page.locator('div[style*="border-top"]', { hasText: '打合せ議事次第' }).getByRole('button', { name: '承認', exact: true }).click();
    await page.getByText(/1件を承認しました/).waitFor();
    // サーバー未設定：押せない
    assert.strictEqual(await page.getByRole('button', { name: 'PDF化（サーバー未設定）' }).isDisabled(), true);
    // サーバーを設定して開き直す → PDF化 → 同じ行に保存
    withServer = true; await open();
    const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'PDF化', exact: true }).click()]);
    await page.getByText(/PDFを保存しました/).waitFor();
    const row = mock.state.records['575'][0]['書類保管'].value[0].value;
    const pdf = row['書類保管_PDF'].value[0];
    assert.match(pdf.name, /_打合せ議事次第_v1\.pdf$/);
    assert.ok(mock.state.files[pdf.fileKey].buf.slice(0, 5).toString() === '%PDF-', 'PDFが保存される');
    assert.strictEqual(row['書類保管_承認状況'].value, '承認済み', 'PDF化しても、承認は残る');
    assert.strictEqual(row['書類保管_ファイル'].value.length, 1, '元のWordも残る');
    assert.deepStrictEqual(errors, []);
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

test('画面：書類関係タブの中（目印の項目の行の、すぐ上）にボタンが出る。タブを閉じると隠れる。見出しの「閉じる」は、折り返さない', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const ctx = await browser.newContext({ viewport: { width: 600, height: 700 } }); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.addInitScript(FAKE_KINTONE);
    await page.goto(url + '/__page');
    await page.evaluate(() => {
      var tab = document.createElement('div'); tab.id = 'tabcontent';
      tab.innerHTML = '<div class="row-gaia" id="row"><div class="control-gaia"><div id="anchor">安全書類格納完了</div></div></div>';
      document.body.appendChild(tab);
    });
    await page.evaluate((r) => window.__fire(r), RECORD);
    const inline = page.locator('#tabcontent [data-doc-hub]');
    assert.strictEqual(await inline.count(), 1, 'タブの中に出る');
    assert.strictEqual(await page.locator('body > [data-doc-hub]').count(), 0, '右下の丸いボタンは、出さない');
    assert.strictEqual(await page.evaluate(() => document.getElementById('row').previousElementSibling.hasAttribute('data-doc-hub')), true, '目印の行の、すぐ上');
    await page.evaluate(() => { document.getElementById('row').style.display = 'none'; });
    await page.waitForFunction(() => document.querySelector('[data-doc-hub]').style.display === 'none');
    await page.evaluate(() => { document.getElementById('row').style.display = 'block'; });
    await page.waitForFunction(() => document.querySelector('[data-doc-hub]').style.display === 'block');
    await inline.getByRole('button').click();
    await page.getByText('工程別の書類：').waitFor();
    const close = page.getByRole('button', { name: '閉じる' });
    const box = await close.boundingBox();
    assert.ok(box.height < 50, '「閉じる」が、折り返されない（高さ ' + box.height + '）');
    assert.deepStrictEqual(errors, []);
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

test('画面：右下の丸いボタン（予備）。目印の項目がなければ、いつでも見える。近づけると文字つきに広がり、押すとパネルが開く', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const { page, errors } = await openPage(browser, url, '/__page');
    const fab = page.locator('[data-doc-hub]');
    assert.strictEqual(await fab.count(), 1); assert.strictEqual(await fab.isVisible(), true);
    const box = await fab.boundingBox(); const vp = page.viewportSize();
    assert.ok(box.width < 70 && Math.abs(box.width - box.height) < 20, '最初は丸い小さなボタン：' + box.width);
    assert.ok(box.x + box.width > vp.width - 40 && box.y + box.height < vp.height, '画面の右下');
    await page.evaluate((r) => window.__fire(r), RECORD);   // 画面を開き直しても、1つだけ
    assert.strictEqual(await page.locator('[data-doc-hub]').count(), 1);
    await fab.hover();
    await page.waitForFunction(() => document.querySelector('[data-doc-hub]').getBoundingClientRect().width > 120);
    assert.match(await fab.textContent(), /書類を作る/);
    await fab.click(); await page.getByText('工程別の書類：').waitFor();
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('画面：目印の項目の枠の呼び名が違っても（row-gaia が無くても）、目印の近くにボタンを置く。置けなければ右下', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const ctx = await browser.newContext(); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.addInitScript(FAKE_KINTONE);
    await page.goto(url + '/__page');
    await page.evaluate(() => { var t = document.createElement('div'); t.id = 'plain'; t.innerHTML = '<div><span id="anchor">安全書類格納完了</span></div>'; document.body.appendChild(t); });
    await page.evaluate((r) => window.__fire(r), RECORD);
    assert.strictEqual(await page.locator('#plain [data-doc-hub]').count(), 1, '枠の呼び名が違っても、目印の近くに置く');
    assert.strictEqual(await page.locator('#plain [data-doc-hub]').isVisible(), true);
    assert.deepStrictEqual(errors, []);
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

// タブの並びの、疑似の画面（PC：背景色で区別／スマホ：文字の色と太さで区別）を作る
const TABBAR = (kind, active) => `
  var bar = document.createElement('div'); bar.id = 'tabs';
  ['営業部', '工事部', '書類関係', '施工店', '請求入金'].forEach(function (n) {
    var d = document.createElement('div'); d.textContent = n; d.setAttribute('data-tab', n);
    d.style.cssText = ${kind === 'pc'} ? 'display:inline-block;padding:6px 10px;background:rgb(200,200,200);color:rgb(80,80,80)' : 'display:inline-block;padding:6px 10px;color:rgb(90,90,90);font-weight:400';
    bar.appendChild(d);
  });
  document.body.appendChild(bar);
  window.__select = function (name) { [].forEach.call(document.querySelectorAll('#tabs div'), function (d) {
    var on = d.getAttribute('data-tab') === name;
    if (${kind === 'pc'}) d.style.background = on ? 'rgb(52,152,219)' : 'rgb(200,200,200)'; else { d.style.color = on ? 'rgb(0,90,180)' : 'rgb(90,90,90)'; d.style.fontWeight = on ? '700' : '400'; } }); };
  window.__select(${JSON.stringify(active)});
`;

for (const kind of ['pc', 'mobile']) {
  test('画面（' + (kind === 'pc' ? 'PC：色' : 'スマホ：文字の太さと色') + '）：「書類関係」タブが選ばれている間だけ、タブの並びの、すぐ下にボタンが出る。ほかのタブ（施工店など）では隠れる', { skip }, async () => {
    const mock = createMock(); const url = await mock.listen();
    const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
    try {
      await seed(mock);
      const ctx = await browser.newContext(); const page = await ctx.newPage();
      const errors = []; page.on('pageerror', (e) => errors.push(e.message));
      await page.addInitScript(FAKE_KINTONE);
      await page.goto(url + '/__page');
      await page.evaluate(new Function(TABBAR(kind, '施工店')));   // 施工店のタブが選ばれている
      await page.evaluate((r) => window.__fire(r), RECORD);
      const btn = page.locator('[data-doc-hub]');
      await page.waitForFunction(() => { var b = document.querySelector('[data-doc-hub]'); return b && document.getElementById('tabs').nextElementSibling === b; });
      assert.strictEqual(await page.locator('body > [data-doc-hub]').count(), 1, '予備の右下のボタンは出ない');
      await page.waitForFunction(() => document.querySelector('[data-doc-hub]').style.display === 'none');   // 施工店では、隠れる
      await page.evaluate(() => window.__select('書類関係'));
      await page.waitForFunction(() => document.querySelector('[data-doc-hub]').style.display === 'block');   // 書類関係で、出る
      await btn.getByRole('button').click(); await page.getByText('工程別の書類：').waitFor();
      assert.deepStrictEqual(errors, []);
      await ctx.close();
    } finally { await browser.close(); await mock.close(); }
  });
}

test('画面：タブの見た目に差がなく、選ばれているタブが見分けられなくても、ボタンは出る（出なくなるのを避ける）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const ctx = await browser.newContext(); const page = await ctx.newPage();
    await page.addInitScript(FAKE_KINTONE);
    await page.goto(url + '/__page');
    await page.evaluate(() => { var bar = document.createElement('div'); bar.id = 'tabs'; ['営業部', '工事部', '書類関係', '施工店', '請求入金'].forEach(function (n) { var d = document.createElement('div'); d.textContent = n; bar.appendChild(d); }); document.body.appendChild(bar); });
    await page.evaluate((r) => window.__fire(r), RECORD);
    await page.waitForFunction(() => { var b = document.querySelector('[data-doc-hub]'); return b && b.style.display !== 'none'; });
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

test('画面：議事録に、音声の文字起こしを貼り付けると、かんたん整理（外部に送らない）／AIで整理（未設定のあいだは押せない）で、項目が埋まる', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const z = await JSZip.loadAsync(tpl.docx);
    const P = (t) => '<w:p><w:r><w:t>' + t + '</w:t></w:r></w:p>';
    z.file('word/document.xml', '<w:document xmlns:w="x"><w:body>' + P('出席者：{{出席者}}') + P('場所：{{場所}}') + P('議題：{{議題}}') + P('内容：{{議事録メモ}}') + P('決定：{{決定事項}}') + P('ToDo：{{ToDo}}') + P('次回：{{次回予定}}') + '</w:body></w:document>');
    const k = mock.addFile(await z.generateAsync({ type: 'nodebuffer' }), 'gijiroku3.docx');
    mock.addRecord('10', { 書類名: { value: '議事録' }, 工程: { value: '打合せ・見積・提案' }, 表示順: { value: '205' }, 説明: { value: '' }, 見られる立場: { value: ['営業'] }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k, name: 'gijiroku3.docx' }] } });
    const ctx = await browser.newContext(); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    let withAI = false; const sent = [];
    page.on('request', (r) => { if (/\/minutes$/.test(r.url()) && r.method() === 'POST') sent.push(r.url()); });
    await page.route('**/__lib/doc-button.js', async (route) => {
      let src = fs.readFileSync(require('path').join(__dirname, '..', 'customize', 'doc-button.js'), 'utf8');
      src = src.replace(/var MINUTES_SERVER_URL = '[^']*';/, "var MINUTES_SERVER_URL = '" + (withAI ? url : '') + "';");
      await route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: src });
    });
    page.on('dialog', (d) => d.accept());
    await page.addInitScript(FAKE_KINTONE);
    const open = async () => {
      await page.goto(url + '/__page'); await page.evaluate((r) => window.__fire(r), RECORD);
      await page.locator('[data-doc-hub]').click(); await page.getByText('工程別の書類：').waitFor();
      await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
      await page.locator('div[style*="border-top"]', { hasText: '議事録' }).getByRole('button', { name: '作成して保存' }).click();
      await page.locator('[data-minutes-src]').waitFor();
    };
    await open();
    assert.strictEqual(await page.getByRole('button', { name: 'AIで整理（未設定）' }).isDisabled(), true, '設定するまで、AIは押せない');
    await page.locator('[data-minutes-src]').fill('えー、出席者：太田、國井\n場所：現場事務所\n工程の確認について。\n夜間規制は9時から17時で決定しました。\n来週までに太田さんが見積書を送付します。\n次回は11月1日です。');
    // Copilot：依頼文をコピー／Copilotを開く
    await ctx.grantPermissions(['clipboard-read', 'clipboard-write'], { origin: url });
    await page.evaluate(() => { window.__opened = []; window.open = function (u) { window.__opened.push(u); return null; }; });
    await page.getByRole('button', { name: '①依頼文をコピー' }).click();
    await page.getByText(/コピーしました/).waitFor();
    const clip = await page.evaluate(() => navigator.clipboard.readText());
    assert.match(clip, /【出席者】[\s\S]*【決定事項】[\s\S]*【ToDo】[\s\S]*【文字起こし】/, 'Copilotに渡す依頼文が、コピーされる');
    await page.getByRole('button', { name: '②Copilotを開く' }).click();
    assert.deepStrictEqual(await page.evaluate(() => window.__opened), ['https://m365.cloud.microsoft/chat']);
    await page.getByRole('button', { name: 'かんたん整理（外部に送りません）' }).click();
    assert.strictEqual(await page.locator('[data-ask="出席者"]').inputValue(), '太田、國井');
    assert.strictEqual(await page.locator('[data-ask="場所"]').inputValue(), '現場事務所');
    assert.match(await page.locator('[data-ask="決定事項"]').inputValue(), /夜間規制/);
    assert.match(await page.locator('[data-ask="ToDo"]').inputValue(), /来週までに太田さん/);
    assert.match(await page.locator('[data-ask="次回予定"]').inputValue(), /11月1日/);
    assert.strictEqual(sent.length, 0, 'かんたん整理は、外部に何も送らない');
    // AI（サーバー設定あり）
    withAI = true; await open();
    await page.locator('[data-minutes-src]').fill('太田さんが03-1234-5678に連絡します。サンプル建設の件です。');
    await page.locator('[data-mask-extra]').fill('サンプル建設');
    const confirms = []; page.removeAllListeners('dialog'); page.on('dialog', (d) => { confirms.push(d.message()); d.accept(); });
    await page.getByRole('button', { name: 'AIで整理（外部のAIに送ります）' }).click();
    await page.getByText(/AIが分けました/).waitFor();
    assert.strictEqual(await page.locator('[data-ask="出席者"]').inputValue(), 'AI太田');
    assert.strictEqual(await page.locator('[data-ask="決定事項"]').inputValue(), '・AI決定');
    assert.strictEqual(sent.length, 1, 'AIを押したときだけ、送る');
    assert.ok(confirms.some((m) => /【送る内容】/.test(m) && /【人物1】さん/.test(m) && !/太田/.test(m)), '送る前に、符号に置き換えた内容を、見せる：' + confirms.join('|'));
    const body = mock.state.minutesBodies[0];
    for (const bad of ['太田', '03-1234-5678', 'サンプル建設']) assert.ok(!body.includes(bad), 'AIに、' + bad + ' を、送らない：' + body);
    assert.match(body, /【人物1】/);
    assert.strictEqual(await page.locator('[data-ask="ToDo"]').inputValue(), '・太田さんが03-1234-5678に連絡', 'AIの返事の符号を、元の言葉に戻す');
    assert.deepStrictEqual(errors, []);
    await ctx.close();
  } finally { await browser.close(); await mock.close(); }
});

test('画面：議事録（{{議事録本文}}）は、ひとつの四角に、見出し付きでまとまる。貼り付け→「かんたん整理」→四角に入る→直して作成→Wordに入る', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    await seed(mock);
    const tpl = await build();
    const z = await JSZip.loadAsync(tpl.docx);
    const P = (t) => '<w:p><w:r><w:t>' + t + '</w:t></w:r></w:p>';
    z.file('word/document.xml', '<w:document xmlns:w="x"><w:body>' + P('日時：{{打合せ日}}　作成者：{{作成者}}') + P('{{議事録本文}}') + '</w:body></w:document>');
    const k = mock.addFile(await z.generateAsync({ type: 'nodebuffer' }), 'matome.docx');
    mock.addRecord('10', { 書類名: { value: '議事録まとめ' }, 工程: { value: '打合せ・見積・提案' }, 表示順: { value: '205' }, 説明: { value: '' }, 見られる立場: { value: ['営業'] }, 使用状況: { value: '使用中' }, 雛形: { value: [{ fileKey: k, name: 'matome.docx' }] } });
    const { page, errors } = await openPage(browser, url, '/__page');
    await page.locator('[data-doc-hub]').click(); await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption('打合せ・見積・提案');
    await page.locator('div[style*="border-top"]', { hasText: '議事録まとめ' }).getByRole('button', { name: '作成して保存' }).click();
    await page.locator('[data-ask="議事録本文"]').waitFor();
    assert.strictEqual(await page.locator('[data-ask]').count(), 1, '入力欄は、ひとつの四角だけ');
    await page.locator('[data-minutes-src]').fill('【出席者】太田、國井\n【議題】\n・夜間規制\n【決定事項】\n・規制は9時から17時\n【ToDo】\n・10/25まで 太田 見積を出す');
    await page.getByRole('button', { name: 'かんたん整理（外部に送りません）' }).click();
    const box = await page.locator('[data-ask="議事録本文"]').inputValue();
    assert.match(box, /^【出席者】\n太田、國井\n\n【場所】\n/); assert.match(box, /【決定事項】\n・規制は9時から17時\n\n【ToDo】\n・10\/25まで 太田 見積を出す\n\n【次回予定】\nなし$/);
    await page.locator('[data-ask="議事録本文"]').fill(box.replace('・夜間規制', '・夜間規制の時間（直した）'));   // 四角の中で、直せる
    const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: '作成して保存' }).last().click()]);
    const got = []; for await (const c of await dl.createReadStream()) got.push(c);
    const xml = await (await JSZip.loadAsync(Buffer.concat(got))).file('word/document.xml').async('string');
    const text = xml.replace(/<w:br\/>/g, '\n').replace(/<[^>]*>/g, '');
    assert.match(text, /日時：\d{4}年\d{1,2}月\d{1,2}日　作成者：くに/, '打合せ日と作成者は、自動で入る');
    assert.match(text, /【議題】\n・夜間規制の時間（直した）/, '直した内容が、Wordに入る');
    assert.ok((xml.match(/<w:b\/>/g) || []).length >= 6, '見出しが太字');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});
