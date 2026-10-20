'use strict';
/* 動作確認用の簡易キントーン（メモリ上）。実際のキントーンの代わりに、スクリプトの流れを確かめるためだけのもの。 */
const http = require('http');

function createMock() {
  const s = { apps: {}, records: {}, files: {}, nextApp: 100, nextFile: 1, nextRec: {}, log: [] };
  const spaces = { 7: { id: '7', name: 'TAIYOSHEET事業部', defaultThread: '70' } };

  function addApp(id, name, extra) { s.apps[id] = Object.assign({ appId: String(id), name, spaceId: null, fields: {}, previewFields: {}, customize: { scope: 'ALL', desktop: { js: [], css: [] }, mobile: { js: [], css: [] } } }, extra || {}); s.records[id] = []; s.nextRec[id] = 1; }
  function addRecord(appId, rec) { const id = s.nextRec[appId]++; s.records[appId].push(Object.assign({ $id: { type: '__ID__', value: String(id) } }, rec)); return String(id); }
  function addFile(buf, name) { const key = 'fk' + s.nextFile++; s.files[key] = { buf, name }; return key; }

  // query: 「項目 = 値」「項目 in ("値")」を and でつないだものと order by / limit だけ解釈する
  function runQuery(appId, query) {
    let recs = s.records[appId].slice();
    query = query || '';
    const lim = query.match(/limit\s+(\d+)/); const ord = query.match(/order by\s+(\S+)\s+(asc|desc)/);
    const cond = query.replace(/order by.*$/, '').replace(/limit\s+\d+/, '').trim();
    if (cond) cond.split(/\s+and\s+/).forEach((c) => {
      let m = c.match(/^(\S+)\s*=\s*(?:"((?:\\.|[^"\\])*)"|(\d+))$/);
      if (m) { const v = m[2] !== undefined ? m[2].replace(/\\(.)/g, '$1') : m[3]; recs = recs.filter((r) => r[m[1]] && String(r[m[1]].value) === v); return; }
      m = c.match(/^(\S+)\s+in\s+\("(.*)"\)$/);
      if (m) { recs = recs.filter((r) => r[m[1]] && String(r[m[1]].value) === m[2]); return; }
      throw new Error('mock: 解釈できないクエリ ' + c);
    });
    if (ord) recs.sort((a, b) => { const x = Number(a[ord[1]].value), y = Number(b[ord[1]].value); return ord[2] === 'desc' ? y - x : x - y; });
    if (lim) recs = recs.slice(0, Number(lim[1]));
    return recs;
  }

  const server = http.createServer(async (req, res) => {
    const chunks = []; for await (const c of req) chunks.push(c);
    const raw = Buffer.concat(chunks);
    const u = new URL(req.url, 'http://x');
    const p = u.pathname.replace(/\.json$/, '');
    const q = {}; u.searchParams.forEach((v, k) => { q[k] = v; });
    const send = (code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(obj)); };
    const fs = require('fs'), path = require('path');
    if (req.method === 'GET' && (p === '/__page' || p === '/k/m/__page')) { res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' }); return res.end('<!doctype html><meta charset="utf-8"><div id="hdr"></div><script src="/__lib/jszip.min.js"></script><script src="/__lib/doc-button.js"></script>'); }
    if (req.method === 'GET' && p === '/__lib/jszip.min.js') { res.writeHead(200, { 'Content-Type': 'text/javascript' }); return res.end(fs.readFileSync(path.join(__dirname, '..', 'node_modules', 'jszip', 'dist', 'jszip.min.js'))); }
    if (req.method === 'GET' && p === '/__lib/doc-button.js') { res.writeHead(200, { 'Content-Type': 'text/javascript; charset=utf-8' }); return res.end(fs.readFileSync(path.join(__dirname, '..', 'customize', 'doc-button.js'))); }
    if (!req.headers['x-cybozu-authorization'] && !req.headers['x-requested-with'] && !req.headers['content-type']) return send(401, { code: 'CB_AU01', message: 'no auth' });
    if (!req.headers['x-cybozu-authorization'] && !req.headers['x-requested-with'] && !(req.headers['content-type'] || '').includes('json')) return send(401, { code: 'CB_AU01', message: 'no auth' });
    let body = {}; const isMulti = (req.headers['content-type'] || '').startsWith('multipart');
    if (!isMulti && raw.length) body = JSON.parse(raw.toString('utf8'));
    s.log.push(req.method + ' ' + p);
    const appOf = (id) => s.apps[String(id)];
    try {
      if (req.method === 'GET' && p === '/k/v1/space') return spaces[q.id] ? send(200, spaces[q.id]) : send(404, { message: 'space' });
      if (req.method === 'GET' && p === '/k/v1/apps') {
        const list = Object.values(s.apps).filter((a) => a.name.includes(q.name || '')).map((a) => ({ appId: a.appId, name: a.name, spaceId: a.spaceId }));
        return send(200, { apps: list.slice(Number(q.offset || 0), Number(q.offset || 0) + Number(q.limit || 100)) });
      }
      if (req.method === 'GET' && p === '/k/v1/app') { const a = appOf(q.id); return a ? send(200, { appId: a.appId, name: a.name, spaceId: a.spaceId }) : send(404, { message: 'app' }); }
      if (req.method === 'POST' && p === '/k/v1/preview/app') { const id = String(s.nextApp++); addApp(id, body.name, { spaceId: String(body.space) }); return send(200, { app: id, revision: '1' }); }
      if (req.method === 'POST' && p === '/k/v1/preview/app/form/fields') { Object.assign(appOf(body.app).previewFields, body.properties); return send(200, { revision: '2' }); }
      if (req.method === 'POST' && p === '/k/v1/preview/app/deploy') { body.apps.forEach((x) => { const a = appOf(x.app); a.fields = Object.assign({}, a.fields, a.previewFields); }); return send(200, {}); }
      if (req.method === 'GET' && p === '/k/v1/preview/app/deploy') return send(200, { apps: q['apps[0]'] ? [{ app: q['apps[0]'], status: 'SUCCESS' }] : [] });
      if (req.method === 'GET' && p === '/k/v1/app/form/fields') return send(200, { properties: appOf(q.app).fields, revision: '1' });
      if (req.method === 'GET' && p === '/k/v1/preview/app/form/fields') return send(200, { properties: appOf(q.app).previewFields, revision: '1' });
      if (req.method === 'GET' && /^\/k\/v1\/(preview\/)?app\/(form\/layout|views|settings|status)$/.test(p)) return send(200, { layout: [], views: {}, revision: '1' });
      if (req.method === 'GET' && /^\/k\/v1\/(preview\/)?app\/customize$/.test(p)) return send(200, Object.assign({ revision: '1' }, appOf(q.app).customize));
      if (req.method === 'PUT' && p === '/k/v1/preview/app/customize') {
        const a = appOf(body.app);
        const conv = (list) => list.map((it) => it.type === 'URL' ? { type: 'URL', url: it.url } : { type: 'FILE', file: { fileKey: it.file.fileKey, name: s.files[it.file.fileKey].name } });
        a.customize = { scope: body.scope || a.customize.scope, desktop: { js: conv(body.desktop.js), css: conv(body.desktop.css) }, mobile: { js: conv(body.mobile.js), css: conv(body.mobile.css) } };
        return send(200, { revision: '3' });
      }
      if (req.method === 'POST' && p === '/k/v1/file') {
        const boundary = '--' + (req.headers['content-type'].match(/boundary=(.+)$/) || [])[1];
        let from = 0, found = null;
        for (;;) {
          const b = raw.indexOf(boundary, from); if (b < 0) break;
          const hdrEnd = raw.indexOf('\r\n\r\n', b); const next = raw.indexOf('\r\n' + boundary, hdrEnd);
          if (hdrEnd < 0 || next < 0) break;
          const head = raw.subarray(b, hdrEnd).toString('utf8'); const nm = head.match(/filename="([^"]*)"/);
          if (nm) { found = { name: nm[1], buf: raw.subarray(hdrEnd + 4, next) }; break; }
          from = next + 2;
        }
        if (!found) return send(400, { message: 'no file part' });
        return send(200, { fileKey: addFile(Buffer.from(found.buf), found.name) });
      }
      if (req.method === 'GET' && p === '/k/v1/file') { const f = s.files[q.fileKey]; if (!f) return send(404, { message: 'file' }); res.writeHead(200); return res.end(f.buf); }
      if (req.method === 'POST' && p === '/k/v1/records') { body.records.forEach((r) => addRecord(body.app, r)); return send(200, { ids: [], revisions: [] }); }
      if (req.method === 'POST' && p === '/k/v1/record') {
        const rec = Object.assign({}, body.record);
        Object.keys(rec).forEach((c) => { if (Array.isArray(rec[c].value) && rec[c].value[0] && rec[c].value[0].fileKey) rec[c] = { value: rec[c].value.map((f) => ({ fileKey: f.fileKey, name: s.files[f.fileKey].name })) }; });
        return send(200, { id: addRecord(body.app, rec), revision: '1' });
      }
      if (req.method === 'GET' && p === '/k/v1/records') {
        const recs = runQuery(q.app, q.query);
        return send(200, { records: recs, totalCount: q.totalCount ? String(s.records[q.app].length) : null });
      }
      if (req.method === 'GET' && p === '/k/v1/record') { const r = s.records[q.app].find((x) => x.$id.value === String(q.id)); return r ? send(200, { record: r }) : send(404, { message: 'rec' }); }
      if (req.method === 'DELETE' && p === '/k/v1/records') { s.records[body.app] = s.records[body.app].filter((r) => !body.ids.map(String).includes(r.$id.value)); return send(200, {}); }
      return send(404, { code: 'MOCK', message: '未対応：' + req.method + ' ' + p });
    } catch (e) { return send(500, { message: e.message }); }
  });

  return {
    state: s, addApp, addRecord, addFile,
    listen: () => new Promise((r) => server.listen(0, '127.0.0.1', () => r('http://127.0.0.1:' + server.address().port))),
    close: () => new Promise((r) => server.close(r)),
  };
}

module.exports = { createMock };
