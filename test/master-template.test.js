'use strict';
/*
 * 書類マスタの「雛形を差し替える」ボタン（customize/master-template.js）を、本物のブラウザ（Chromium）で動かして確かめる。
 * キントーン本体の代わりに、模擬キントーンを使う。Chromium が無い環境では、自動でスキップされる。
 */
const test = require('node:test');
const fs = require('fs');
const assert = require('node:assert');
const { createMock } = require('./mock-kintone');
const { build } = require('../scripts/make-test-templates');

const CHROMIUM = ['/opt/pw-browsers/chromium', process.env.CHROMIUM_PATH].find((p) => p && fs.existsSync(p));
let chromium; try { chromium = require('playwright-core').chromium; } catch (e) { chromium = null; }
const skip = !chromium || !CHROMIUM ? 'Chromium が見つからないためスキップ' : false;

const FAKE = `
  function q(params) { var o = []; Object.keys(params || {}).forEach(function (k) { o.push(encodeURIComponent(k) + '=' + encodeURIComponent(params[k])); }); return o.join('&'); }
  function api(url, method, params) {
    var get = method === 'GET';
    return fetch(url + (get ? '?' + q(params) : ''), { method: method, headers: { 'X-Requested-With': 'XMLHttpRequest', 'Content-Type': 'application/json' }, body: get ? undefined : JSON.stringify(params) })
      .then(function (r) { return r.json().then(function (j) { if (!r.ok) throw j; return j; }); });
  }
  api.url = function (p) { return p + '.json'; };
  var h = {};
  window.kintone = {
    api: api, getRequestToken: function () { return 'tok'; },
    events: { on: function (n, fn) { [].concat(n).forEach(function (x) { h[x] = fn; }); } },
    app: { getId: function () { return 10; }, getHeaderMenuSpaceElement: function () { return document.getElementById('hdr'); },
      record: { getHeaderMenuSpaceElement: function () { return document.getElementById('hdr'); } } },
    mobile: { app: { getId: function () { return 10; }, getHeaderSpaceElement: function () { return document.getElementById('hdr'); } } }
  };
  if (/nohost/.test(location.search)) { window.kintone.app.getHeaderMenuSpaceElement = function () { return null; }; window.kintone.app.record.getHeaderMenuSpaceElement = function () { return null; }; }
  window.__fire = function (n, ev) { return h[n](ev); };
`;

const names = (mock, id) => mock.state.records['10'].find((r) => r.$id.value === String(id))['雛形'].value.map((f) => f.name);

async function setup(mock, url, browser, path) {
  const ctx = await browser.newContext(); const page = await ctx.newPage();
  const errors = []; page.on('pageerror', (e) => errors.push(e.message));
  const dialogs = []; page.on('dialog', (d) => { dialogs.push(d.message()); d.accept(); });
  await page.addInitScript(FAKE);
  await page.goto(url + (path || '/__page'));
  await page.addScriptTag({ content: fs.readFileSync(require.resolve('jszip/dist/jszip.min.js'), 'utf8') });
  await page.addScriptTag({ content: fs.readFileSync(__dirname + '/../customize/master-template.js', 'utf8') });
  await page.addScriptTag({ content: fs.readFileSync(__dirname + '/../customize/doc-button.js', 'utf8') });   // 書類マスタでは、master-template.js の後ろに登録する
  return { page, errors, dialogs };
}
async function choose(page, trigger, paths) {
  const [fc] = await Promise.all([page.waitForEvent('filechooser'), trigger()]);
  await fc.setFiles(paths);
}

test('書類マスタ：詳細画面で雛形を差し替える（古いものは旧版で残り、新しいものが先頭）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    const tpl = await build(); fs.writeFileSync('/tmp/new-tpl.docx', tpl.docx);
    mock.addApp('10', '書類マスタ');
    const k1 = mock.addFile(Buffer.from('OLD1'), '議事録_旧.docx');
    const id = mock.addRecord('10', { 書類名: { value: '打合せ議事録（社内）' }, 雛形: { value: [{ fileKey: k1, name: '議事録_旧.docx' }] } });
    const { page, errors, dialogs } = await setup(mock, url, browser);
    await page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    assert.strictEqual(await page.locator('[data-master-swap]').count(), 1);
    await choose(page, () => page.locator('[data-master-swap]').click(), '/tmp/new-tpl.docx');
    await page.getByText(/差し替えました/).waitFor();
    const n = names(mock, id);
    assert.strictEqual(n[0], 'new-tpl.docx', '新しい雛形が先頭');
    assert.match(n[1], /^旧版_\d{8}_議事録_旧\.docx$/, '古いものは旧版で残る');
    assert.match(dialogs[0], /差し込む項目：/);
    assert.deepStrictEqual(errors, []);
    // 旧版の中身も、そのまま残っている
    const kept = mock.state.records['10'][0]['雛形'].value[1];
    assert.strictEqual(mock.state.files[kept.fileKey].buf.toString(), 'OLD1');
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：一覧画面で、複数ファイルを、ファイル名＝書類名で、まとめて差し替える（合わないものは、差し替えず知らせる）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    // ※ このテスト環境の Chromium は、日本語のファイル名を、パスでは渡せないため、中身を直接渡す
    const pay = (name, text) => ({ name, mimeType: 'application/octet-stream', buffer: Buffer.from(text) });
    mock.addApp('10', '書類マスタ');
    const k = mock.addFile(Buffer.from('OLDSPEC'), '仕様_旧.docx');
    const a = mock.addRecord('10', { 書類名: { value: '工事仕様書' }, 雛形: { value: [{ fileKey: k, name: '仕様_旧.docx' }] } });
    const b = mock.addRecord('10', { 書類名: { value: '作業手順書' }, 雛形: { value: [] } });
    const { page, errors, dialogs } = await setup(mock, url, browser);
    await page.evaluate(() => window.__fire('app.record.index.show', {}));
    await choose(page, () => page.locator('[data-master-bulk]').click(), [pay('工事仕様書.docx', 'SPEC'), pay('作業手順書.xlsx', 'SAGYO'), pay('ぜんぜん違う名前.docx', 'X')]);
    await page.getByText(/2件を、差し替えました/).waitFor();
    assert.deepStrictEqual(names(mock, a).slice(0, 1), ['工事仕様書.docx']);
    assert.match(names(mock, a)[1], /^旧版_\d{8}_仕様_旧\.docx$/);
    assert.deepStrictEqual(names(mock, b), ['作業手順書.xlsx']);
    assert.match(dialogs[0], /ぜんぜん違う名前\.docx（同じ書類名が、ありません）/, '合わないファイルは、確認画面で知らせる');
    assert.match(await page.locator('[data-master-note]').textContent(), /対象外：1件/);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：確認で「キャンセル」なら、何も変えない', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    
    mock.addApp('10', '書類マスタ');
    const k1 = mock.addFile(Buffer.from('OLD1'), '旧.docx');
    const id = mock.addRecord('10', { 書類名: { value: 'あ' }, 雛形: { value: [{ fileKey: k1, name: '旧.docx' }] } });
    const { page } = await setup(mock, url, browser);
    page.removeAllListeners('dialog'); page.on('dialog', (d) => d.dismiss());
    await page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    await choose(page, () => page.locator('[data-master-swap]').click(), { name: 'new2.docx', mimeType: 'application/octet-stream', buffer: Buffer.from('N') });
    await page.waitForTimeout(600);
    assert.deepStrictEqual(names(mock, id), ['旧.docx']);
    assert.strictEqual(mock.state.records['10'][0]['雛形'].value.length, 1, 'キャンセルなら、何も増えない');
  } finally { await browser.close(); await mock.close(); }
});

// 案件くん（575）の見本：項目の定義と、案件1件
function addAnken(mock) {
  mock.addApp('575', '案件くん', { fields: {
    c_name: { type: 'SINGLE_LINE_TEXT', code: 'c_name', label: '工事名称' },
    c_place: { type: 'SINGLE_LINE_TEXT', code: 'c_place', label: '工事場所' },
  } });
  mock.addRecord('575', { c_name: { type: 'SINGLE_LINE_TEXT', value: '第二多摩川橋 補修工事' }, c_place: { type: 'SINGLE_LINE_TEXT', value: '東京都調布市' } });
}
const docxWith = async (text) => {
  const JSZip = require('jszip'); const z = new JSZip();
  z.file('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>');
  z.file('word/document.xml', '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' + text + '</w:t></w:r></w:p></w:body></w:document>');
  return z.generateAsync({ type: 'nodebuffer' });
};

test('書類マスタ：雛形チェッカー（差し替えの確認に、案件くんにない項目名を出す）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    mock.addRecord('10', { 書類名: { value: 'あ' }, 雛形: { value: [] } });
    const { page, dialogs } = await setup(mock, url, browser);
    await page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    const buf = await docxWith('工事名称：{{工事名称}} 場所：{{工事場所}} 打ち間違い：{{工事場所々}} 日付：{{作成日}}');
    await choose(page, () => page.locator('[data-master-swap]').click(), { name: 'new.docx', mimeType: 'application/octet-stream', buffer: buf });
    await page.getByText(/差し替えました/).waitFor();
    assert.match(dialogs[0], /【注意】案件くんにない項目名：{{工事場所々}}/);
    assert.doesNotMatch(dialogs[0], /案件くんにない項目名：[^\n]*{{工事名称}}/, '正しい名前は、注意に出さない');
    assert.doesNotMatch(dialogs[0], /案件くんにない項目名：[^\n]*{{作成日}}/, '自動で入る言葉は、注意に出さない');
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：試しに作る（見本の案件を選ぶ→記入済みを作って点検。どこにも保存しない）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    const k = mock.addFile(await docxWith('工事名称：{{工事名称}} 場所：{{工事場所}} 版：{{版}} 内容：{{決定事項}} 誤り：{{ないもの}}'), '雛形.docx');
    mock.addRecord('10', { 書類名: { value: '試験書類' }, 雛形: { value: [{ fileKey: k, name: '雛形.docx' }] } });
    const before = JSON.stringify(mock.state.records);
    const { page, errors } = await setup(mock, url, browser);
    await page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    assert.strictEqual(await page.locator('[data-master-try]').count(), 1);
    await page.locator('[data-master-try]').click();
    await page.locator('[data-try-case]').first().waitFor();
    assert.match(await page.locator('[data-try-panel]').textContent(), /第二多摩川橋 補修工事/);
    await page.locator('[data-try-case]').first().click();
    await page.getByText(/点検の結果/).waitFor();
    const t = await page.locator('[data-try-panel]').textContent();
    assert.match(t, /案件くんにない項目名（空になります）：.*ないもの/, '雛形の打ち間違いを、指摘する');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('[data-try-download]').click()]);
    const got = []; for await (const c of await dl.createReadStream()) got.push(c);
    const JSZip = require('jszip'); const xml = await (await JSZip.loadAsync(Buffer.concat(got))).file('word/document.xml').async('string');
    assert.match(xml, /第二多摩川橋 補修工事/); assert.match(xml, /東京都調布市/); assert.match(xml, /v0（試し）/); assert.match(xml, /【入力欄：決定事項】/);
    assert.strictEqual(JSON.stringify(mock.state.records), before, '案件くんにも書類マスタにも、何も書き込まない');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：スマホ（モバイル）でも、ボタンが出て、差し替えられる／場所が見つからなくても、帯でボタンが出る', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    const k1 = mock.addFile(Buffer.from('OLD'), '旧.docx');
    const id = mock.addRecord('10', { 書類名: { value: 'あ' }, 雛形: { value: [{ fileKey: k1, name: '旧.docx' }] } });
    // スマホ
    let t = await setup(mock, url, browser, '/k/m/__page');
    await t.page.evaluate((r) => window.__fire('mobile.app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    assert.strictEqual(await t.page.locator('[data-master-swap]').count(), 1, 'スマホの詳細にボタン');
    assert.strictEqual(await t.page.locator('[data-master-try]').count(), 1);
    await choose(t.page, () => t.page.locator('[data-master-swap]').click(), { name: 'm.docx', mimeType: 'application/octet-stream', buffer: await docxWith('{{工事名称}}') });
    await t.page.getByText(/差し替えました/).waitFor();
    assert.strictEqual(names(mock, id)[0], 'm.docx');
    await t.page.evaluate(() => window.__fire('mobile.app.record.index.show', {}));
    assert.strictEqual(await t.page.locator('[data-master-bulk]').count(), 1, 'スマホの一覧にボタン');
    assert.deepStrictEqual(t.errors, []);
    // ボタンの置き場所が見つからない（PC）
    t = await setup(mock, url, browser, '/__page?nohost=1');
    await t.page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    assert.strictEqual(await t.page.locator('[data-master-bar] [data-master-swap]').count(), 1, '帯の中に、ボタンが出る');
  } finally { await browser.close(); await mock.close(); }
});
