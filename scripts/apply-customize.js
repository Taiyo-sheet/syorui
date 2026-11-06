'use strict';
/*
 * 案件くんのPC用・モバイル用の両方に、JSZip と doc-button.js を登録して公開する。
 * 登録する内容は customize-manifest.json（上から順に読み込まれる。http で始まるものはURL、それ以外はこのフォルダのファイル）。
 *
 *   node scripts/apply-customize.js            … 登録後の姿を表示するだけ（何も変えない）
 *   node scripts/apply-customize.js --yes      … 実際に登録して公開する
 *   --manifest=ファイル  … 登録内容のファイルを指定する（省略時は customize-manifest.json。アプリごとに別のファイルを使える）
 *
 * なぜ @kintone/customize-uploader を使わないか
 *   customize-uploader は、マニフェストに書いたものでカスタマイズ設定を「丸ごと置き換える」ため、
 *   案件くんに元からあるJS/CSSが消えてしまう。このスクリプトは、いまの設定を読み込んで
 *   「既存を全部残したまま」追加（同名ファイルは入れ替え）するので、安全に更新できる。
 *
 * 守ること
 *   ・変更前のバックアップ（scripts/backup-app.js）が無いときは、--yes でも実行しない
 *   ・既存のJS/CSSは1つも削除しない（同名の doc-button.js だけは新しい版に入れ替える）
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');

const ROOT = path.join(__dirname, '..');

function hasBackup(appId) {
  const d = process.env.KINTONE_BACKUP_DIR || path.join(ROOT, 'backup');
  return fs.existsSync(d) && fs.readdirSync(d).some((n) => n.startsWith(appId + '_') && fs.existsSync(path.join(d, n, 'customize_live.json')));
}
const label = (it) => it.type === 'URL' ? it.url : (it.file && it.file.name) || '(ファイル)';

// 編集中（プレビュー）と、公開中（本番）の、フォームの違い
async function pendingChanges(k, appId) {
  const out = [];
  const get = async (p, q) => { try { return await k.api('GET', p, q); } catch (e) { return null; } };
  const [pl, ll, pf, lf, pv, lv] = await Promise.all([
    get('/k/v1/preview/app/form/layout', { app: appId }), get('/k/v1/app/form/layout', { app: appId }),
    get('/k/v1/preview/app/form/fields', { app: appId }), get('/k/v1/app/form/fields', { app: appId }),
    get('/k/v1/preview/app/views', { app: appId }), get('/k/v1/app/views', { app: appId })]);
  if (pl && ll && JSON.stringify(pl.layout) !== JSON.stringify(ll.layout)) out.push('フォームの並び（行数 公開中 ' + ll.layout.length + ' → 編集中 ' + pl.layout.length + '）');
  if (pf && lf) {
    const a = Object.keys(pf.properties), b = new Set(Object.keys(lf.properties));
    const added = a.filter((c) => !b.has(c)), removed = [...b].filter((c) => !pf.properties[c]);
    if (added.length || removed.length) out.push('項目：追加 ' + added.length + '件' + (removed.length ? '／削除 ' + removed.length + '件' : ''));
  }
  if (pv && lv && JSON.stringify(pv.views) !== JSON.stringify(lv.views)) out.push('一覧（ビュー）');
  return out;
}

async function main() {
  const go = process.argv.includes('--yes');
  const mArg = process.argv.find((a) => a.startsWith('--manifest='));
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, mArg ? mArg.slice(11) : 'customize-manifest.json'), 'utf8'));
  // --manifest= を指定したときは、そのファイルの app を使う（環境変数 KINTONE_APP_ID で、うっかり別のアプリに入れないため）
  const appId = String((mArg ? manifest.app : process.env.KINTONE_APP_ID) || manifest.app);
  const k = client();

  const app = await k.api('GET', '/k/v1/app', { id: appId });
  console.log('対象アプリ：「' + app.name + '」（番号 ' + appId + '）');
  const cur = await k.api('GET', '/k/v1/preview/app/customize', { app: appId });

  const plan = { app: appId, scope: cur.scope, desktop: { js: [], css: [] }, mobile: { js: [], css: [] } };
  const uploads = [];
  for (const env of ['desktop', 'mobile']) {
    for (const kind of ['js', 'css']) {
      plan[env][kind] = ((cur[env] && cur[env][kind]) || []).map((it) => it.type === 'URL' ? { type: 'URL', url: it.url } : { type: 'FILE', file: { fileKey: it.file.fileKey }, _name: it.file.name });
    }
    const add = (manifest[env] && manifest[env].js) || [];
    for (const src of add) {
      if (/^https?:\/\//.test(src)) {
        if (!plan[env].js.some((x) => x.type === 'URL' && x.url === src)) plan[env].js.push({ type: 'URL', url: src });
      } else {
        const name = path.basename(src);
        plan[env].js = plan[env].js.filter((x) => !(x.type === 'FILE' && x._name === name));   // 同名の古い版だけ入れ替える
        const entry = { type: 'FILE', file: { fileKey: null }, _name: name, _local: path.join(ROOT, src) };
        plan[env].js.push(entry); uploads.push(entry);
      }
    }
  }

  console.log('\n登録後の姿（上から順に読み込まれます）');
  for (const env of ['desktop', 'mobile']) {
    console.log('  [' + (env === 'desktop' ? 'PC用' : 'モバイル用') + ']');
    plan[env].js.forEach((it, i) => console.log('    JS ' + (i + 1) + '. ' + (it._name || label(it)) + (it._local ? '  ← 今回追加／更新' : '')));
    plan[env].css.forEach((it, i) => console.log('    CSS ' + (i + 1) + '. ' + (it._name || label(it))));
  }
  const before = ['desktop', 'mobile'].reduce((n, e) => n + ((cur[e].js || []).length + (cur[e].css || []).length), 0);
  console.log('\n既存の設定：' + before + '件（すべて残ります）');
  // 公開すると、フォームの編集中の変更（まだ公開していないもの）も、一緒に公開されてしまう。公開前に、編集中の変更が無いか確かめる
  const pend = await pendingChanges(k, appId);
  if (pend.length) {
    console.log('\n【注意】アプリ設定に、まだ公開していない変更（編集中）があります。このまま公開すると、これも、一緒に公開されます：');
    pend.forEach((t) => console.log('  ・' + t));
    if (go && !process.argv.includes('--include-pending')) { console.error('\n公開を止めました。利用者が、その変更を、公開するか、取り消してから、もう一度、実行してください（それでも、一緒に公開してよいときだけ、--include-pending）。'); process.exit(3); }
  }
  if (!go) { console.log('\n※ これは確認だけです。実際に登録するには、最後に --yes を付けてください。'); return; }

  if (!hasBackup(appId)) { console.error('\nバックアップがありません。先に  node scripts/backup-app.js ' + appId + '  を実行してください。'); process.exit(2); }

  // 同じファイルをPC用・モバイル用で別々にアップロードして登録する
  for (const e of uploads) e.file.fileKey = await k.upload(fs.readFileSync(e._local), e._name, 'text/javascript');
  const clean = (list) => list.map((it) => it.type === 'URL' ? it : { type: 'FILE', file: { fileKey: it.file.fileKey } });
  const body = { app: appId, desktop: { js: clean(plan.desktop.js), css: clean(plan.desktop.css) }, mobile: { js: clean(plan.mobile.js), css: clean(plan.mobile.css) } };
  if (cur.scope) body.scope = cur.scope;
  await k.api('PUT', '/k/v1/preview/app/customize', body);
  console.log('設定を更新しました。公開しています…');
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app: appId }] });
  await k.waitDeploy(appId, app.name);
  console.log('公開しました。案件くんの画面を再読み込み（Ctrl+F5）して、案件を開いてください。');
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
