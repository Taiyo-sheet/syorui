'use strict';
/*
 * キントーンREST APIの共通部品。
 * 認証情報は環境変数だけから読む（ファイル・ログには出さない）。
 *   KINTONE_DOMAIN    例）〇〇.cybozu.com
 *   KINTONE_USER      ログイン名
 *   KINTONE_PASSWORD  パスワード
 */
const fs = require('fs');
const path = require('path');

function need(name) {
  const v = process.env[name];
  if (!v) {
    console.error('環境変数 ' + name + ' が設定されていません。（.env.example を見てください）');
    process.exit(1);
  }
  return v;
}

// .env があれば読む（すでにある環境変数は上書きしない）。.env は .gitignore 済み
function loadDotEnv() {
  const f = path.join(__dirname, '..', '.env');
  if (!fs.existsSync(f)) return;
  fs.readFileSync(f, 'utf8').split(/\r?\n/).forEach((line) => {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.*?)\s*$/);
    if (m && !(m[1] in process.env)) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  });
}

function toQuery(params) {
  const out = [];
  Object.keys(params || {}).forEach((k) => {
    const v = params[k];
    if (Array.isArray(v)) v.forEach((x, i) => out.push(encodeURIComponent(k + '[' + i + ']') + '=' + encodeURIComponent(x)));
    else if (v !== undefined && v !== null) out.push(encodeURIComponent(k) + '=' + encodeURIComponent(v));
  });
  return out.join('&');
}

function client() {
  loadDotEnv();
  const domain = need('KINTONE_DOMAIN').replace(/^https?:\/\//, '').replace(/\/.*$/, '');
  const auth = Buffer.from(need('KINTONE_USER') + ':' + need('KINTONE_PASSWORD')).toString('base64');
  const base = process.env.KINTONE_BASE_URL || 'https://' + domain;   // KINTONE_BASE_URL はテスト用（模擬サーバーへ向ける）

  async function api(method, p, params) {
    const isGet = method === 'GET';
    const url = base + p + '.json' + (isGet && params ? '?' + toQuery(params) : '');
    const res = await fetch(url, {
      method,
      headers: Object.assign({ 'X-Cybozu-Authorization': auth }, isGet ? {} : { 'Content-Type': 'application/json' }),
      body: isGet ? undefined : JSON.stringify(params || {}),
    });
    const text = await res.text();
    let json; try { json = JSON.parse(text); } catch (e) { json = { message: text.slice(0, 200) }; }
    if (!res.ok) {
      const err = new Error(method + ' ' + p + ' が失敗しました（' + res.status + '）：' + (json.message || '') + (json.code ? ' [' + json.code + ']' : ''));
      err.status = res.status; err.code = json.code;
      throw err;
    }
    return json;
  }

  async function upload(data, filename, contentType) {
    const fd = new FormData();
    fd.append('file', new Blob([data], { type: contentType || 'application/octet-stream' }), filename);
    const res = await fetch(base + '/k/v1/file.json', { method: 'POST', headers: { 'X-Cybozu-Authorization': auth }, body: fd });
    const json = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error('ファイルのアップロードに失敗しました（' + res.status + '）：' + (json.message || ''));
    return json.fileKey;
  }

  async function download(fileKey) {
    const res = await fetch(base + '/k/v1/file.json?fileKey=' + encodeURIComponent(fileKey), { headers: { 'X-Cybozu-Authorization': auth } });
    if (!res.ok) throw new Error('ファイルのダウンロードに失敗しました（' + res.status + '）');
    return Buffer.from(await res.arrayBuffer());
  }

  // アプリ名が完全一致するアプリを全件から探す（apps.json は部分一致・100件ずつ）
  async function findApps(name) {
    const found = [];
    for (let offset = 0; ; offset += 100) {
      const r = await api('GET', '/k/v1/apps', { name, limit: 100, offset });
      r.apps.forEach((a) => { if (a.name === name) found.push(a); });
      if (r.apps.length < 100) break;
    }
    return found;
  }

  async function waitDeploy(appId, label) {
    for (let i = 0; i < 90; i++) {
      await new Promise((r) => setTimeout(r, 2000));
      const s = await api('GET', '/k/v1/preview/app/deploy', { apps: [appId] });
      const st = s.apps[0].status;
      if (st === 'SUCCESS') return;
      if (st === 'FAIL' || st === 'CANCEL') throw new Error(label + ' の公開に失敗しました');
    }
    throw new Error(label + ' の公開が終わりません。キントーンの画面で状況を確認してください');
  }

  return { api, upload, download, findApps, waitDeploy, domain, base };
}

// CSV（ダブルクォート・改行対応・先頭のBOMは無視）
function parseCsv(text) {
  text = text.replace(/^﻿/, '');
  const rows = []; let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"' && text[i + 1] === '"') { cur += '"'; i++; } else if (c === '"') q = false; else cur += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(cur); cur = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && text[i + 1] === '\n') i++; row.push(cur); rows.push(row); row = []; cur = ''; }
    else cur += c;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  const head = rows.shift();
  return rows.filter((r) => r.some((x) => x !== '')).map((r) => { const o = {}; head.forEach((h, i) => { o[h] = r[i] === undefined ? '' : r[i]; }); return o; });
}

module.exports = { client, parseCsv, need, loadDotEnv };
