'use strict';
/*
 * 外部アプリ（Render）を開くボタンが、鍵つきかを、本番のキントーンから読み取って確かめる（読み取りだけ。何も変えない）。
 *   node scripts/check-open-buttons.js        （npm run check-buttons）
 *
 * 見るもの：案件くん(575)・見積くん(608)・発注くん(824)・請求くん(829)の JavaScript の中の、
 *   anken-kanri-1.onrender.com へのアドレス。鍵（?k= または OPEN_KEY）が付いているか。
 *   鍵が付いているもの同士で、鍵が同じか（食い違いの発見）。
 * 鍵そのものは、画面にも、ログにも、出さない（あり／なし、同じ／ちがう、だけ）。
 *
 * 作った理由：2026-10-27 に外部アプリが「鍵か会社のIP」だけを通すようになり、鍵のないPC版ボタンが、社外で、
 *   「このアプリは kintone から開いてください」と止まった（障害メモ_外部アプリが開けない.md）。同じことを、反映のたびに、早く見つけるため。
 */
const crypto = require('crypto');
const { client } = require('./lib');

const APPS = { 575: '案件くん', 608: '見積くん', 824: '発注くん', 829: '請求くん' };
const HOST = /anken-kanri-1\.onrender\.com/;

// JSのテキストから、外部アプリへのアドレスの有無・鍵の有無を調べる（鍵は返すが、表示はしない）
function analyze(text) {
  if (!HOST.test(text)) return null;
  const urlKey = (text.match(/anken-kanri-1\.onrender\.com[^'"\s]*?[?&]k=([A-Za-z0-9_.~%-]{4,})/) || [])[1];
  const varKey = (text.match(/var\s+OPEN_KEY\s*=\s*['"]([^'"]{4,})['"]/) || [])[1];
  const key = varKey || (urlKey ? decodeURIComponent(urlKey) : '');
  return { hasKey: !!key, key };
}
const digest = (s) => crypto.createHash('sha256').update(s).digest('hex');

async function main() {
  const k = client();
  const found = [];
  for (const app of Object.keys(APPS)) {
    let c;
    try { c = await k.api('GET', '/k/v1/app/customize', { app }); } catch (e) { console.log(APPS[app] + '(' + app + ')：読めませんでした'); continue; }
    for (const env of ['desktop', 'mobile']) {
      for (const it of (c[env] && c[env].js) || []) {
        if (!it.file) continue;
        if (it.file.name === 'taiyo-login-link.js') continue;   // ログイン用の道具（ボタンではない。鍵ではなく、押した人のログインで開く）。他のセッションの管理
        let t; try { t = Buffer.from(await k.download(it.file.fileKey)).toString('utf8'); } catch (e) { continue; }
        const r = analyze(t);
        if (r) found.push({ app, env, name: it.file.name, hasKey: r.hasKey, d: r.key ? digest(r.key) : '' });
      }
    }
  }
  let bad = 0;
  for (const f of found) {
    console.log((f.hasKey ? '○ 鍵あり　' : '✕ 鍵なし　') + APPS[f.app] + '(' + f.app + ') ' + (f.env === 'desktop' ? 'PC' : 'スマホ') + '  ' + f.name);
    if (!f.hasKey) bad++;
  }
  const keys = [...new Set(found.filter((f) => f.hasKey).map((f) => f.d))];
  if (keys.length > 1) { console.log('✕ 鍵が、ボタンによって、ちがいます（' + keys.length + '種類）'); bad++; }
  else if (keys.length === 1) console.log('○ 鍵のあるボタンは、すべて同じ鍵です');
  console.log(bad ? '\n要確認：' + bad + '件（鍵のないボタンは、社外で開けません。ただし、会社のIPからは開けます）' : '\n問題なし');
  process.exit(bad ? 2 : 0);
}
if (require.main === module) main().catch((e) => { console.error(e.message || e); process.exit(1); });
module.exports = { analyze };
