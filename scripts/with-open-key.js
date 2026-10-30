'use strict';
/*
 * 案件くん(575)のPC版ボタン（taiyo-open-app.js）に、鍵（?k=）を入れて反映する。鍵は、リポジトリに残さない。
 *   node scripts/with-open-key.js        … 本番のスマホ版ボタン（taiyo-open-app-mobile.js）から鍵を読み取り、
 *                                          手元の taiyo-open-app.js に一時的に入れて safe-deploy --app=575 を実行し、
 *                                          終わったら、必ず、鍵なしの元の内容に戻す（失敗しても戻す）
 * 鍵は、画面にも、ログにも、出さない。
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { client } = require('./lib');

const FILE = path.join(__dirname, '..', 'customize', 'taiyo-open-app.js');

(async () => {
  const k = client();
  const c = await k.api('GET', '/k/v1/app/customize', { app: 575 });
  const it = ((c.mobile && c.mobile.js) || []).find((x) => x.file && x.file.name === 'taiyo-open-app-mobile.js');
  if (!it) { console.error('本番のスマホ版ボタンが見つかりません。止めました'); process.exit(1); }
  const live = Buffer.from(await k.download(it.file.fileKey)).toString('utf8');
  const m = live.match(/[?&]k=([A-Za-z0-9_.~%-]{4,})/);
  if (!m) { console.error('スマホ版ボタンに鍵が見つかりません。止めました'); process.exit(1); }
  const original = fs.readFileSync(FILE, 'utf8');
  if (!/var OPEN_KEY = '';/.test(original)) { console.error('taiyo-open-app.js に、鍵の欄がありません'); process.exit(1); }
  let code = 1;
  try {
    fs.writeFileSync(FILE, original.replace("var OPEN_KEY = '';", "var OPEN_KEY = " + JSON.stringify(decodeURIComponent(m[1])) + ";"));
    const r = spawnSync('node', ['scripts/safe-deploy.js', '--app=575'].concat(process.argv.slice(2)), { cwd: path.join(__dirname, '..'), stdio: 'inherit', env: process.env });
    code = r.status;
  } finally { fs.writeFileSync(FILE, original); }
  console.log('（手元のファイルは、鍵なしの元の内容に戻しました）');
  process.exit(code);
})().catch((e) => { console.error(e.message || e); process.exit(1); });
