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
    events: { on: function (n, fn) { h[n] = fn; } },
    app: { getId: function () { return 10; }, getHeaderMenuSpaceElement: function () { return document.getElementById('hdr'); },
      record: { getHeaderMenuSpaceElement: function () { return document.getElementById('hdr'); } } }
  };
  window.__fire = function (n, ev) { return h[n](ev); };
`;

const names = (mock, id) => mock.state.records['10'].find((r) => r.$id.value === String(id))['雛形'].value.map((f) => f.name);

async function setup(mock, url, browser) {
  const ctx = await browser.newContext(); const page = await ctx.newPage();
  const errors = []; page.on('pageerror', (e) => errors.push(e.message));
  const dialogs = []; page.on('dialog', (d) => { dialogs.push(d.message()); d.accept(); });
  await page.addInitScript(FAKE);
  await page.goto(url + '/__page');
  await page.addScriptTag({ content: fs.readFileSync(require.resolve('jszip/dist/jszip.min.js'), 'utf8') });
  await page.addScriptTag({ content: fs.readFileSync(__dirname + '/../customize/master-template.js', 'utf8') });
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
