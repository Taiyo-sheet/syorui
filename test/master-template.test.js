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
  function q(params) { var o = []; Object.keys(params || {}).forEach(function (k) { var v = params[k]; if (Array.isArray(v)) v.forEach(function (x, i) { o.push(encodeURIComponent(k + '[' + i + ']') + '=' + encodeURIComponent(x)); }); else o.push(encodeURIComponent(k) + '=' + encodeURIComponent(v)); }); return o.join('&'); }
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
    app: { getId: function () { return 10; }, getHeaderSpaceElement: function () { return document.getElementById('hdr'); }, getHeaderMenuSpaceElement: function () { return document.getElementById('hdr'); },
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
    assert.doesNotMatch(t, /置き換えられていない項目が残っています/, '同じ指摘を、2回出さない');
    assert.match(await page.locator('[data-try-preview]').textContent(), /第二多摩川橋 補修工事/, '中身が、文字でも見える');
    const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('[data-try-download]').click()]);
    const got = []; for await (const c of await dl.createReadStream()) got.push(c);
    const JSZip = require('jszip'); const xml = await (await JSZip.loadAsync(Buffer.concat(got))).file('word/document.xml').async('string');
    assert.match(xml, /第二多摩川橋 補修工事/); assert.match(xml, /東京都調布市/); assert.match(xml, /v0（試し）/); assert.match(xml, /規制は9時から17時とする/, '議事録の項目は、見本の文章が入る');
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

test('書類マスタ：スマホでは、試しに作った書類の中身を、文字で見せる（ファイルを開くボタンは、出さない）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    const k = mock.addFile(await docxWith('工事名称：{{工事名称}}'), '雛形.docx');
    mock.addRecord('10', { 書類名: { value: 'い' }, 雛形: { value: [{ fileKey: k, name: '雛形.docx' }] } });
    const { page, errors } = await setup(mock, url, browser, '/k/m/__page');
    await page.evaluate((r) => window.__fire('mobile.app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    await page.locator('[data-master-try]').click();
    await page.locator('[data-try-case]').first().click();
    await page.getByText(/点検の結果：問題は/).waitFor();
    const pv = await page.locator('[data-try-preview]').textContent();
    assert.match(pv, /工事名称：第二多摩川橋 補修工事/);
    assert.strictEqual(await page.locator('[data-try-download]').count(), 0, 'スマホでは、ファイルを開くボタンを出さない');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：「雛形の状況」の一覧（雛形の有無・案件での作成数・最終作成日）と、全雛形の点検', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    const row = (n, d) => ({ id: '1', value: { 書類保管_書類名: { value: n }, 書類保管_作成日: { value: d } } });
    mock.state.records['575'][0]['書類保管'] = { type: 'SUBTABLE', value: [row('良い書類', '2026-10-01'), row('良い書類', '2026-10-20')] };
    const good = mock.addFile(await docxWith('{{工事名称}}'), 'good.docx'), typo = mock.addFile(await docxWith('{{工事名称々}}'), 'typo.docx');
    const mk = (n, st, f, aud, use) => mock.addRecord('10', { 書類名: { value: n }, 工程: { value: st }, 表示順: { value: '1' }, 使用状況: { value: use || '使用中' }, 見られる立場: { value: aud || [] }, 雛形: { value: f } });
    mk('良い書類', '打合せ', [{ fileKey: good, name: 'good.docx' }], ['施工店']); mk('打ち間違いの書類', '打合せ', [{ fileKey: typo, name: 'typo.docx' }], ['施工店', '工事部']); mk('雛形のない書類', '竣工', [], ['工事部']); mk('古い書類', '竣工', [], ['施工店'], '廃止');
    const { page, errors } = await setup(mock, url, browser);
    await page.evaluate(() => window.__fire('app.record.index.show', {}));   // 最初から、開いている。最初は「雛形が必要なもの」だけ出す
    await page.getByText(/すべての立場：書類 3 件のうち、雛形が必要（まだ入っていない）なもの 1 件　／　入っているもの 2 件/).waitFor();
    let t0 = await page.locator('[data-master-dash]').textContent();
    assert.match(t0, /雛形のない書類/); assert.doesNotMatch(t0, /良い書類/, '最初は、足りないものだけ');
    // プルダウンで、すべて（使用中）を選ぶ
    await page.locator('[data-mode-select]').selectOption('all');
    await page.getByText(/いま表示：3 件/).waitFor();
    const t = await page.locator('[data-master-dash]').textContent();
    assert.match(t, /良い書類.*good\.docx.*2.*2026-10-20/); assert.match(t, /雛形のない書類.*未登録/);
    // 立場を、プルダウンで選ぶ（工事部用：2件のうち、雛形が必要なもの1件）
    await page.locator('[data-aud-select]').selectOption('工事部');
    await page.getByText(/工事部用：書類 2 件のうち、雛形が必要（まだ入っていない）なもの 1 件　／　入っているもの 1 件　（いま表示：2 件）/).waitFor();
    await page.locator('[data-mode-select]').selectOption('need');
    await page.getByText(/いま表示：1 件/).waitFor();
    assert.doesNotMatch(await page.locator('[data-master-dash]').textContent(), /打ち間違いの書類/);
    // 廃止した書類
    await page.locator('[data-aud-select]').selectOption('施工店'); await page.locator('[data-mode-select]').selectOption('old');
    await page.getByText(/いま表示：1 件/).waitFor();
    assert.match(await page.locator('[data-master-dash]').textContent(), /古い書類（廃止）/);
    // 雛形が入っているもの（施工店用）：2件。全点検
    await page.locator('[data-mode-select]').selectOption('have');
    await page.getByText(/いま表示：2 件/).waitFor();
    await page.locator('[data-master-check-all]').click();
    await page.getByText(/2件を点検：注意 1 件/).waitFor();
    const t2 = await page.locator('[data-master-dash]').textContent();
    assert.match(t2, /良い書類.*OK/); assert.match(t2, /打ち間違いの書類.*注意：{{工事名称々}}/);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：{{ }}の候補（案件くんの項目名＋自動で入る言葉・見本の値・絞り込み・コピー）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    mock.addRecord('10', { 書類名: { value: 'あ' }, 雛形: { value: [] } });
    const { page, errors } = await setup(mock, url, browser);
    await page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    await page.locator('[data-master-cand]').click();
    await page.locator('[data-cand-item]').first().waitFor();
    const all = await page.locator('[data-cand-item]').evaluateAll((e) => e.map((x) => x.getAttribute('data-cand-item')));
    assert.ok(all.includes('工事名称') && all.includes('工事場所'), '案件くんの項目が出る');
    assert.ok(all.includes('作成日') && all.includes('議事録本文'), '自動で入る言葉も出る');
    assert.match(await page.locator('[data-cand-item="工事名称"]').textContent(), /見本：第二多摩川橋 補修工事/);
    await page.locator('[data-cand-search]').fill('場所');
    assert.ok((await page.locator('[data-cand-item]').evaluateAll((e) => e.map((x) => x.getAttribute('data-cand-item')))).every((n) => n.includes('場所')), '絞り込める');
    await page.locator('[data-cand-item="工事場所"]').click();
    await page.waitForFunction(() => /{{工事場所}}/.test(document.querySelector('[data-cand-note]').textContent));
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：雛形の履歴と、前の版に戻す（選んだ版が先頭・いまの先頭は旧版で残る）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ');
    const f = (t, n) => ({ fileKey: mock.addFile(Buffer.from(t), n), name: n });
    const id = mock.addRecord('10', { 書類名: { value: 'う' }, 雛形: { value: [f('NEW', '新.docx'), f('OLD1', '旧版_20261001_前.docx'), f('OLD0', '旧版_20260901_最初.docx')] } });
    const { page, errors } = await setup(mock, url, browser);
    await page.evaluate((r) => window.__fire('app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    await page.locator('[data-master-hist]').click();
    await page.locator('[data-hist-restore="1"]').waitFor();
    const t = await page.locator('body').textContent();
    assert.match(t, /【いま使われている】 新\.docx/); assert.match(t, /前\.docx（2026\/10\/01 に、旧版になりました）/);
    await page.locator('[data-hist-restore="1"]').click();
    await page.getByText(/戻しました/).waitFor();
    const n = names(mock, id);
    assert.strictEqual(n[0], '前.docx', '選んだ版が先頭（旧版の印は外れる）');
    assert.match(n[1], /^旧版_\d{8}_新\.docx$/, 'いまの先頭は、旧版で残る');
    assert.strictEqual(n[2], '旧版_20260901_最初.docx');
    const first = mock.state.records['10'][0]['雛形'].value[0];
    assert.strictEqual(mock.state.files[first.fileKey].buf.toString(), 'OLD1', '中身も、その版');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});

test('書類マスタ：試しに作る（議事録の{{議事録本文}}には、見本の本文が入り、見出しが付く）', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('10', '書類マスタ'); addAnken(mock);
    const k = mock.addFile(await docxWith('{{議事録本文}}'), 'gijiroku.docx');
    mock.addRecord('10', { 書類名: { value: '議事録' }, 雛形: { value: [{ fileKey: k, name: 'gijiroku.docx' }] } });
    const { page, errors } = await setup(mock, url, browser, '/k/m/__page');
    await page.evaluate((r) => window.__fire('mobile.app.record.detail.show', { record: r }), mock.state.records['10'][0]);
    await page.locator('[data-master-try]').click();
    await page.locator('[data-try-case]').first().click();
    await page.getByText(/点検の結果/).waitFor();
    const pv = await page.locator('[data-try-preview]').textContent();
    assert.match(pv, /出席者/); assert.match(pv, /見積書を提出/); assert.doesNotMatch(pv, /【入力欄/);
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});
