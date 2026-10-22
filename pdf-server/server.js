'use strict';
/*
 * PDF化サーバー：Word・Excel・PowerPointを、PDFに変換して返すだけ（LibreOffice を使う）。
 *   POST /convert?ext=docx|xlsx|pptx   本文＝ファイルそのもの  →  PDF（application/pdf）
 *   GET  /healthz                      → ok
 *
 * 守ること
 *   ・受け取ったファイルも、変換したPDFも、保存しない（一時ファイルは、変換のたびに消す）。ファイルの中身は、ログに出さない
 *   ・通信元（Origin）が、許可したもの（ALLOWED_ORIGINS、初期値は https://taiyo1.cybozu.com）でなければ、断る。
 *     ブラウザ以外から使う場合は、PDF_SERVER_KEY を設定して、X-Api-Key ヘッダーで渡す
 *   ・大きさの上限（25MB）、同時に変換する数の上限（2）、変換の制限時間（60秒）がある
 *   ※ 通信元の確認は、ブラウザを前提にした、簡易な守りです（ブラウザ以外は、Origin を偽装できます）。
 *     強くしたいときは、Render側で、接続元のIP制限や、鍵（PDF_SERVER_KEY）の運用を、追加してください
 */
const http = require('http');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn } = require('child_process');

const EXTS = ['docx', 'xlsx', 'pptx'];
const MAX_BYTES = 25 * 1024 * 1024;
const MAX_JOBS = 2;
const TIMEOUT_MS = 60000;

function createServer(opts) {
  const o = Object.assign({ allowedOrigins: (process.env.ALLOWED_ORIGINS || 'https://taiyo1.cybozu.com').split(',').map((s) => s.trim()).filter(Boolean), key: process.env.PDF_SERVER_KEY || '', soffice: process.env.SOFFICE || 'soffice' }, opts || {});
  let jobs = 0;

  function convert(buf, ext) {
    return new Promise((resolve, reject) => {
      const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'pdf-'));
      const done = (err, out) => { fs.rmSync(dir, { recursive: true, force: true }); err ? reject(err) : resolve(out); };
      const src = path.join(dir, 'in.' + ext);
      fs.writeFileSync(src, buf);
      const p = spawn(o.soffice, ['--headless', '--norestore', '-env:UserInstallation=file://' + path.join(dir, 'profile'), '--convert-to', 'pdf', '--outdir', dir, src], { stdio: 'ignore' });
      const timer = setTimeout(() => { p.kill('SIGKILL'); done(new Error('timeout')); }, TIMEOUT_MS);
      p.on('error', (e) => { clearTimeout(timer); done(e); });
      p.on('exit', () => {
        clearTimeout(timer);
        const out = path.join(dir, 'in.pdf');
        if (!fs.existsSync(out)) return done(new Error('no-output'));
        done(null, fs.readFileSync(out));
      });
    });
  }

  return http.createServer((req, res) => {
    const u = new URL(req.url, 'http://x');
    const origin = req.headers.origin || '';
    const originOk = origin && o.allowedOrigins.includes(origin);
    const keyOk = o.key && req.headers['x-api-key'] === o.key;
    const cors = originOk ? { 'Access-Control-Allow-Origin': origin, 'Vary': 'Origin' } : {};
    const send = (code, body, type) => { res.writeHead(code, Object.assign({ 'Content-Type': type || 'text/plain; charset=utf-8' }, cors)); res.end(body); };

    if (req.method === 'GET' && u.pathname === '/healthz') return send(200, 'ok');
    if (req.method === 'OPTIONS') {
      if (!originOk) return send(403, 'forbidden');
      res.writeHead(204, Object.assign({ 'Access-Control-Allow-Methods': 'POST, OPTIONS', 'Access-Control-Allow-Headers': 'Content-Type, X-Api-Key', 'Access-Control-Max-Age': '600' }, cors)); return res.end();
    }
    if (req.method !== 'POST' || u.pathname !== '/convert') return send(404, 'not found');
    if (!originOk && !keyOk) return send(403, 'forbidden');
    const ext = u.searchParams.get('ext');
    if (!EXTS.includes(ext)) return send(400, 'ext は docx / xlsx / pptx のいずれかです');
    if (jobs >= MAX_JOBS) return send(503, '混み合っています。少し待って、もう一度試してください');

    const chunks = []; let size = 0, aborted = false;
    req.on('data', (c) => { size += c.length; if (size > MAX_BYTES) { aborted = true; send(413, '25MBまでです'); req.destroy(); } else chunks.push(c); });
    req.on('end', async () => {
      if (aborted) return;
      if (!size) return send(400, 'ファイルが空です');
      const body = Buffer.concat(chunks);
      if (body.slice(0, 2).toString() !== 'PK') return send(400, 'Word・Excel・PowerPointのファイルではありません');   // docx・xlsx・pptx は、zip形式
      jobs++;
      try { send(200, await convert(body, ext), 'application/pdf'); }
      catch (e) { send(500, '変換に失敗しました'); console.error('convert failed:', e.message); }   // 中身は、出さない
      finally { jobs--; }
    });
  });
}

module.exports = { createServer };
if (require.main === module) {
  const port = Number(process.env.PORT || 10000);
  createServer().listen(port, () => console.log('pdf-server listening on ' + port));
}
