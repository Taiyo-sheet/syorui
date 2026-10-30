'use strict';
/*
 * 案件くんの画面イメージを、実際の customize/doc-button.js を動かして撮る（説明用の画像づくり）。
 * キントーン本体ではなく、模擬キントーンで動かす。実際の画面とは、見た目の細部（キントーンの枠・色）が違う。
 *
 *   node test/demo-screens.js 出力フォルダ
 *
 * 案件くんに「橋梁区分」（橋梁／一般道路）の項目がある前提：doc-button.js の TARGET_LABEL に '橋梁区分' を入れた状態で動かす。
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { chromium } = require('playwright-core');
const { createMock } = require('./mock-kintone');
const { build } = require('../scripts/make-test-templates');
const { parseCsv } = require('../scripts/lib');

const CHROMIUM = ['/opt/pw-browsers/chromium', process.env.CHROMIUM_PATH].find((p) => p && fs.existsSync(p));
const out = path.resolve(process.argv[2] || 'out-demo');
fs.mkdirSync(out, { recursive: true });

const FIELDS = {
  c_no: { type: 'SINGLE_LINE_TEXT', code: 'c_no', label: '案件番号' },
  c_name: { type: 'SINGLE_LINE_TEXT', code: 'c_name', label: '案件名' },
  c_place: { type: 'SINGLE_LINE_TEXT', code: 'c_place', label: '工事場所' },
  c_client: { type: 'SINGLE_LINE_TEXT', code: 'c_client', label: '発注者_会社名' },
  c_prime: { type: 'SINGLE_LINE_TEXT', code: 'c_prime', label: '元請_会社名_0' },
  c_type: { type: 'DROP_DOWN', code: 'c_type', label: '橋梁区分' },
};
const rec = (type) => ({
  c_no: { type: 'SINGLE_LINE_TEXT', value: 'TS53-012' }, c_name: { type: 'SINGLE_LINE_TEXT', value: '第二多摩川橋 ロードヒーティング設置工事' },
  c_place: { type: 'SINGLE_LINE_TEXT', value: '東京都調布市多摩川1-1' }, c_client: { type: 'SINGLE_LINE_TEXT', value: '国土交通省 関東地方整備局' },
  c_prime: { type: 'SINGLE_LINE_TEXT', value: '株式会社サンプル建設' }, c_type: { type: 'DROP_DOWN', value: type },
});

const PAGE = (type) => `
<style>
  body{margin:0;font-family:"Noto Sans CJK JP","Noto Sans JP",Meiryo,sans-serif;background:#f4f5f6;color:#333}
  .top{background:#3498db;color:#fff;padding:10px 20px;font-weight:bold}
  .bar{background:#fff;border-bottom:1px solid #ddd;padding:8px 20px;display:flex;align-items:center;gap:8px}
  .bar .t{font-size:18px;font-weight:bold;margin-right:auto}
  .rec{background:#fff;margin:16px 20px;padding:16px 20px;border:1px solid #ddd}
  .row{display:flex;border-bottom:1px solid #eee;padding:8px 0;font-size:14px}.row b{width:160px;color:#666;font-weight:normal}
  .hl{background:#fff7d6}
  .sub{border-collapse:collapse;width:100%;font-size:13px}.sub th,.sub td{border:1px solid #ddd;padding:6px 8px;text-align:left}.sub th{background:#f0f3f5;font-weight:normal;color:#666}
</style>
<div class="top">案件くん</div>
<div class="bar"><span class="t">案件詳細</span><span id="hdr"></span></div>
<div class="rec">
  <div class="row"><b>案件番号</b>TS53-012</div>
  <div class="row"><b>案件名</b>第二多摩川橋 ロードヒーティング設置工事</div>
  <div class="row"><b>工事場所</b>東京都調布市多摩川1-1</div>
  <div class="row"><b>発注者</b>国土交通省 関東地方整備局</div>
  <div class="row"><b>元請</b>株式会社サンプル建設</div>
  <div class="row hl"><b>橋梁区分</b>${type}</div>
  <div class="row" style="display:block"><b style="display:block;margin-bottom:6px">書類保管</b>
    <table class="sub"><tr><th>書類名</th><th>工程</th><th>版</th><th>作成日</th><th>ファイル</th></tr>
    <tr><td>作業手順書</td><td>施工前の書類</td><td>1</td><td>2026-10-12</td><td>📎 第二多摩川橋_作業手順書_v1.docx</td></tr>
    <tr><td>作業手順書</td><td>施工前の書類</td><td>2</td><td>2026-10-19</td><td>📎 第二多摩川橋_作業手順書_v2.docx</td></tr>
    <tr><td>打合せ議事次第</td><td>打合せ・見積・提案</td><td>1</td><td>2026-10-19</td><td>📎 第二多摩川橋_打合せ議事次第_v1.docx</td></tr></table></div>
</div>`;

(async () => {
  const mock = createMock(); const url = await mock.listen();
  const tpl = await build();
  mock.addApp('575', '案件くん'); mock.addRecord('575', { 書類保管: { type: 'SUBTABLE', value: [] } }); mock.addApp('10', '書類マスタ');
  const word = mock.addFile(tpl.docx, 'tpl.docx'), xl = mock.addFile(tpl.xlsx, 'tpl.xlsx');
  // {{ }} のない雛形（使い回す手引き）
  const zp = await JSZip.loadAsync(tpl.docx);
  zp.file('word/document.xml', '<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>共通の手引き</w:t></w:r></w:p></w:body></w:document>');
  zp.file('word/header1.xml', '<w:hdr xmlns:w="x"><w:p><w:r><w:t>共通</w:t></w:r></w:p></w:hdr>');
  const plain = mock.addFile(await zp.generateAsync({ type: 'nodebuffer' }), 'guide.docx');
  const pdf = mock.addFile(Buffer.from('%PDF-1.4'), 'manual.pdf');

  const rows = parseCsv(fs.readFileSync(path.join(__dirname, '..', 'data', '書類マスタ_初期データ.csv'), 'utf8'));
  const PLAIN = new Set(['TSマニュアル', '機器仕様書', '試験要領書']);
  const HAS = new Set(['現状確認シート', '実案件の情報取得チェックシート', '初期リスク抽出シート', 'LINEWORKS周知文', '打合せ議事次第', '見積書', '提案書', '準備品リスト', '安全書類一式', '作業手順書', '作業計画書', '試験成績書', '作業日報・KY記録', '竣工図書', '原価管理表']);
  const add = (r, name, target) => {
    const ext = r['形式'] === 'Excel' ? 'xlsx' : 'docx';
    let files = [];
    if (PLAIN.has(r['書類名'])) files = [r['形式'] === 'PDF' ? { fileKey: pdf, name: r['書類名'] + '.pdf' } : { fileKey: plain, name: r['書類名'] + '.docx' }];
    else if (HAS.has(name) || target !== '共通') files = [{ fileKey: ext === 'xlsx' ? xl : word, name: name + '.' + ext }];
    mock.addRecord('10', {
      書類名: { value: name }, 工程: { value: r['工程'] }, 表示順: { value: r['表示順'] }, 説明: { value: r['説明'] },
      見られる立場: { value: r['見られる立場'].split('|') }, 対象: { value: target }, 使用状況: { value: '使用中' }, 雛形: { value: files },
    });
  };
  rows.forEach((r) => {
    if (r['書類名'] === '仕様書') { add(r, '仕様書（橋梁）', '橋梁'); add(r, '仕様書（一般道路）', '一般道路'); }
    else add(r, r['書類名'], '共通');
  });
  // この案件では、作業手順書が v2、安全書類一式が v1 まで作成済み
  const rowsSaved = [];
  const saved = (name, stage, v) => { rowsSaved.push({ id: String(rowsSaved.length + 1), value: { 書類保管_書類名: { value: name }, 書類保管_工程: { value: stage }, 書類保管_版: { value: String(v) }, 書類保管_作成日: { value: '2026-10-19' }, 書類保管_ファイル: { value: [] } } }); mock.state.records['575'][0]['書類保管'].value = rowsSaved; };
  saved('作業手順書', '施工前の書類', 2); saved('安全書類一式', '施工前の書類', 1); saved('打合せ議事次第', '打合せ・見積・提案', 1); saved('見積書', '打合せ・見積・提案', 1);

  const browser = await chromium.launch({ executablePath: CHROMIUM, args: ['--no-sandbox'] });
  const shot = async (type, name, steps) => {
    const ctx = await browser.newContext({ viewport: { width: 980, height: 760 }, acceptDownloads: true });
    const page = await ctx.newPage();
    page.on('dialog', (d) => d.accept());
    await page.route('**/__lib/doc-button.js', async (route) => {
      const src = fs.readFileSync(path.join(__dirname, '..', 'customize', 'doc-button.js'), 'utf8').replace("var TARGET_LABEL = '';", "var TARGET_LABEL = '橋梁区分';");
      await route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: src });
    });
    await page.addInitScript(`
      var FIELDS = ${JSON.stringify(FIELDS)};
      function q(p){var o=[];Object.keys(p||{}).forEach(function(k){var v=p[k];if(Array.isArray(v))v.forEach(function(x,i){o.push(encodeURIComponent(k+'['+i+']')+'='+encodeURIComponent(x))});else if(v!=null)o.push(encodeURIComponent(k)+'='+encodeURIComponent(v))});return o.join('&')}
      function api(url,m,p){var g=m==='GET';return fetch(url+(g?'?'+q(p):''),{method:m,headers:{'X-Requested-With':'XMLHttpRequest','Content-Type':'application/json'},body:g?undefined:JSON.stringify(p)}).then(function(r){return r.json().then(function(j){if(!r.ok)throw j;return j})})}
      api.url=function(p){return p+'.json'};
      var hs=[];var hdr=function(){return document.getElementById('hdr')};
      window.kintone={api:api,getRequestToken:function(){return 't'},events:{on:function(n,f){hs.push(f)}},
        app:{getId:function(){return 575},getFormFields:function(){return Promise.resolve(FIELDS)},record:{getId:function(){return 1},getHeaderMenuSpaceElement:hdr}},
        mobile:{app:{getId:function(){return 575},getHeaderSpaceElement:hdr,record:{getId:function(){return 1}}}}};
      window.__fire=function(r){hs.forEach(function(h){h({record:r})})};`);
    await page.goto(url + '/__page');
    await page.evaluate((html) => { document.body.insertAdjacentHTML('afterbegin', html); }, PAGE(type));
    await page.evaluate((r) => window.__fire(r), rec(type));
    await steps(page);
    await page.screenshot({ path: path.join(out, name) });
    await ctx.close();
  };

  // 1) 案件の詳細画面：ヘッダーに「書類を作る」
  await shot('橋梁', '1_案件の詳細画面.png', async () => {});
  // 1b) 案件の一番下：書類保管表（作った書類が、版つきで並ぶ）
  await shot('橋梁', '1b_案件の書類保管表.png', async (page) => { await page.setViewportSize({ width: 980, height: 560 }); });
  // 1c) 右下の丸いボタン（通常）と、近づけたとき（文字つきに広がる）
  await shot('橋梁', '1c_右下の丸いボタン.png', async (page) => { await page.setViewportSize({ width: 980, height: 560 }); await page.waitForTimeout(300); });
  await shot('橋梁', '1d_近づけると広がる.png', async (page) => { await page.setViewportSize({ width: 980, height: 560 }); await page.locator('[data-doc-hub]').hover(); await page.waitForTimeout(500); });
  // 2) 橋梁の案件：施工前の書類
  const open = async (page, stage) => {
    await page.locator('[data-doc-hub]').click();
    await page.getByText('工程別の書類：').waitFor();
    await page.locator('select').nth(0).selectOption(stage);
    await page.waitForFunction(() => !/確認中/.test(document.body.innerText));
  };
  await shot('橋梁', '2_橋梁の案件_施工前の書類.png', (page) => open(page, '施工前の書類'));
  // 3) 同じ案件で、橋梁区分を「一般道路」に変えたとき
  await shot('一般道路', '3_一般道路に変えたとき.png', (page) => open(page, '施工前の書類'));
  // 4) 打合せ・見積・提案
  await shot('橋梁', '4_打合せ・見積・提案.png', (page) => open(page, '打合せ・見積・提案'));
  await browser.close(); await mock.close();
  console.log(fs.readdirSync(out).join('\n'));
})();
