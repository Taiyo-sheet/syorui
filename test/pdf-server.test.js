'use strict';
/* PDF化サーバーを、本物の LibreOffice で動かして確かめる（soffice が無い環境では、自動でスキップ） */
const test = require('node:test');
const assert = require('node:assert');
const { execFileSync } = require('node:child_process');
const { createServer } = require('../pdf-server/server');
const { build } = require('../scripts/make-test-templates');

let has = true; try { execFileSync('soffice', ['--version'], { stdio: 'ignore' }); } catch (e) { has = false; }
const skip = has ? false : 'soffice（LibreOffice）が見つからないためスキップ';
const ORIGIN = 'https://taiyo1.cybozu.com';

async function withServer(opts, fn) {
  const s = createServer(opts); await new Promise((r) => s.listen(0, '127.0.0.1', r));
  const base = 'http://127.0.0.1:' + s.address().port;
  try { await fn(base); } finally { await new Promise((r) => s.close(r)); }
}

test('pdf-server：docx・xlsx・pptx が PDF になる／許可した画面（Origin）からだけ', { skip, timeout: 170000 }, async () => {
  const tpl = await build();
  await withServer({}, async (base) => {
    for (const [ext, buf] of [['docx', tpl.docx], ['xlsx', tpl.xlsx], ['pptx', tpl.pptx]]) {
      const r = await fetch(base + '/convert?ext=' + ext, { method: 'POST', headers: { Origin: ORIGIN, 'Content-Type': 'application/octet-stream' }, body: buf });
      assert.strictEqual(r.status, 200, ext + ' → ' + r.status);
      assert.strictEqual(r.headers.get('access-control-allow-origin'), ORIGIN);
      const pdf = Buffer.from(await r.arrayBuffer());
      assert.strictEqual(pdf.slice(0, 5).toString(), '%PDF-', ext + ' が PDF になる');
      assert.ok(pdf.length > 1000);
    }
    // 事前確認（CORS）
    const pre = await fetch(base + '/convert?ext=docx', { method: 'OPTIONS', headers: { Origin: ORIGIN, 'Access-Control-Request-Method': 'POST' } });
    assert.strictEqual(pre.status, 204);
    // 許可していない画面、Originなし、不正な種類、空のファイル
    let r = await fetch(base + '/convert?ext=docx', { method: 'POST', headers: { Origin: 'https://evil.example' }, body: tpl.docx }); assert.strictEqual(r.status, 403);
    r = await fetch(base + '/convert?ext=docx', { method: 'POST', body: tpl.docx }); assert.strictEqual(r.status, 403, 'Originなしは断る');
    r = await fetch(base + '/convert?ext=exe', { method: 'POST', headers: { Origin: ORIGIN }, body: tpl.docx }); assert.strictEqual(r.status, 400);
    r = await fetch(base + '/convert?ext=docx', { method: 'POST', headers: { Origin: ORIGIN } }); assert.strictEqual(r.status, 400);
    r = await fetch(base + '/healthz'); assert.strictEqual(await r.text(), 'ok');
  });
});

test('pdf-server：鍵（PDF_SERVER_KEY）があれば、Originなしでも、鍵で使える。Wordではないファイルは、400', { skip, timeout: 170000 }, async () => {
  const tpl = await build();
  await withServer({ key: 'k-test' }, async (base) => {
    let r = await fetch(base + '/convert?ext=docx', { method: 'POST', headers: { 'X-Api-Key': 'wrong' }, body: tpl.docx }); assert.strictEqual(r.status, 403);
    r = await fetch(base + '/convert?ext=docx', { method: 'POST', headers: { 'X-Api-Key': 'k-test' }, body: tpl.docx }); assert.strictEqual(r.status, 200);
    r = await fetch(base + '/convert?ext=docx', { method: 'POST', headers: { 'X-Api-Key': 'k-test' }, body: Buffer.from('これはWordではありません') }); assert.strictEqual(r.status, 400, 'Wordではないものは、変換しない');
  });
});
