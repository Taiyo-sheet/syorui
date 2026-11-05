'use strict';
/*
 * 現場カルテ(826)の「原価管理表を作る」ボタン（customize/karte-genka.js）を、本物のブラウザ（Chromium）で動かして確かめる。模擬キントーン。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const JSZip = require('jszip');
const { createMock } = require('./mock-kintone');
const CHROMIUM = ['/opt/pw-browsers/chromium', process.env.CHROMIUM_PATH].find((p) => p && fs.existsSync(p));
let chromium; try { chromium = require('playwright-core').chromium; } catch (e) { chromium = null; }
const skip = !chromium || !CHROMIUM ? 'Chromium が見つからないためスキップ' : false;
const ui = fs.readFileSync(require('path').join(__dirname, 'ui.test.js'), 'utf8');
const st = ui.indexOf('const FAKE_KINTONE = `') + 'const FAKE_KINTONE = `'.length;
const FAKE = eval('`' + ui.slice(st, ui.indexOf('`;', st)) + '`');

async function genkaBook() {
  const z = new JSZip();
  z.file('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>');
  z.file('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="原価管理" sheetId="1" r:id="rId1"/></sheets></workbook>');
  z.file('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/></Relationships>');
  let rows = '<row r="2"><c r="AQ2" s="6"><f>SUM(AQ8:AT10)</f></c></row><row r="3"><c r="E3" t="inlineStr"><is><t>{{工事名称}}</t></is></c><c r="W3" t="inlineStr"><is><t>{{工事場所}}</t></is></c><c r="AQ3" s="6"/><c r="AG2" s="7"/><c r="AG3" s="7"/><c r="AG4" s="7"/><c r="AG5" s="8"/></row>';
  for (let r = 24; r < 40; r++) rows += `<row r="${r}"><c r="B${r}" s="3"/><c r="F${r}" s="4"/><c r="R${r}" s="4"/><c r="AJ${r}" s="5"/><c r="AV${r}" s="5"/></row>`;
  z.file('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + rows + '</sheetData></worksheet>');
  return z.generateAsync({ type: 'nodebuffer' });
}

test('現場カルテ：「原価管理表を作る」で、現場カルテの数字と請求履歴が入ったExcelがダウンロードされる。何も書き込まない', { skip }, async () => {
  const mock = createMock(); const url = await mock.listen();
  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  try {
    mock.addApp('879', '書類マスタ');
    const k = mock.addFile(await genkaBook(), 'genka.xlsx');
    mock.addRecord('879', { 書類名: { value: '原価管理表' }, 雛形: { value: [{ fileKey: k, name: 'genka.xlsx' }] } });
    mock.addApp('829', '請求くん');
    const bill = (a, d, v, kb, items) => mock.addRecord('829', { anken: { value: a }, billing_date: { value: d }, vendor_name: { value: v }, kubun: { value: kb }, details: { value: items.map(([n, m]) => ({ value: { item_name: { value: n }, bill_amount: { value: String(m) } } })) } });
    bill('0085', '2026-09-28', 'A電設', '発注連動', [['電気工事材料', 143700], ['配管', 30000]]); bill('9999', '2026-09-01', '別案件', '発注連動', [['関係ない', 777]]);
    const rec = { anken: { type: 'SINGLE_LINE_TEXT', value: '0085' }, proj_name: { type: 'SINGLE_LINE_TEXT', value: 'のと里山海道 電気工事' }, koji_code: { type: 'SINGLE_LINE_TEXT', value: 'TS52-0085' }, moto: { type: 'SINGLE_LINE_TEXT', value: '' },
      start_date: { type: 'DATE', value: '2026-09-07' }, end_date: { type: 'DATE', value: '2026-09-08' }, sales: { type: 'NUMBER', value: '16787066' }, budget: { type: 'NUMBER', value: '11457400' }, ordered: { type: 'NUMBER', value: '7900000' } };
    const ctx = await browser.newContext({ acceptDownloads: true }); const page = await ctx.newPage();
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    await page.addInitScript('window.__TS_DOC_ENGINE_ONLY = true;');
    await page.addInitScript(FAKE);
    await page.addInitScript(fs.readFileSync(require('path').join(__dirname, '..', 'customize', 'karte-genka.js'), 'utf8'));
    await page.goto(url + '/__page'); await page.evaluate((r) => window.__fire(r), rec);
    const before = JSON.stringify(mock.state.records);
    const [dl] = await Promise.all([page.waitForEvent('download'), page.locator('#ts-karte-genka-btn').click()]);
    await page.getByText(/ダウンロードしました.*請求履歴 2 行.*受注金額・予算を入れました/).waitFor();
    const path = await dl.path(); const xml = await (await JSZip.loadAsync(fs.readFileSync(path))).file('xl/worksheets/sheet1.xml').async('string');
    assert.match(xml, /<t[^>]*>のと里山海道 電気工事<\/t>/, '工事名');
    assert.match(xml, /<c r="AG2" s="7"><v>16787066<\/v><\/c>/, '受注金額'); assert.ok(!/AG4[^>]*><f>/.test(xml), '受注原価がないので、受注粗利の計算式は入れない');
    assert.match(xml, /<c r="AQ2" s="6"><v>11457400<\/v><\/c>/, '予算金額'); assert.match(xml, /<c r="AQ3" s="6"><v>7900000<\/v><\/c>/, '予算原価＝発注確定額');
    assert.match(xml, /<c r="AV24" s="5"><v>143700<\/v><\/c>/); assert.match(xml, /<c r="AV25" s="5"><v>30000<\/v><\/c>/, '請求の明細が、1行ずつ');
    assert.ok(!/別案件|関係ない/.test(xml), '別の案件の請求は入らない'); assert.ok(!/\{\{/.test(xml), '{{ }} が残らない（対応する項目がないものは、空欄）');
    assert.strictEqual(JSON.stringify(mock.state.records), before, '何も書き込まない');
    assert.deepStrictEqual(errors, []);
  } finally { await browser.close(); await mock.close(); }
});
