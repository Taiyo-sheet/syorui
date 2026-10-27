'use strict';
/*
 * 書類マスタ(879)・案件くん(575)の「JavaScriptだけ」を、安全な手順で、本番に反映する（利用者が、この範囲で、Claudeが自分で反映してよいと許可：2026-10-27）
 *
 *   node scripts/safe-deploy.js --app=575 [--dry]
 *   node scripts/safe-deploy.js --app=879 [--dry]
 *
 * 手順（1つでも、失敗したら、そこで止まる。変更は、始まっていなければ、何もしない）
 *   ① 自動テストが、すべて通る
 *   ② バックアップを取る
 *   ③ フォームの編集中の変更（未公開）が、ない（あれば、止める。apply-customize の確認）
 *   ④ JavaScriptを登録して公開する（フォーム・項目・レコード・プラグインは、触らない）
 *   ⑤ 公開後に、フォームの並びと項目が、反映前と同じか／登録したファイルが、手元と同じか、確かめる
 *   ⑥ ⑤で、おかしければ、バックアップのJavaScriptの登録に、自動で戻す
 *
 * 対象は、879 と 575 の、JavaScriptの登録だけ。ほかのアプリ・フォーム・項目・レコード・プラグインの変更は、利用者の許可が、そのつど、必要。
 */
const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');
const { client } = require('./lib');

const ROOT = path.join(__dirname, '..');
const MANIFESTS = { 575: 'customize-manifest.json', 879: 'customize-manifest-master-879.json' };
const arg = (n) => ((process.argv.find((a) => a.startsWith('--' + n + '=')) || '').split('=').slice(1).join('='));
const has = (n) => process.argv.includes('--' + n);
const say = (t) => console.log(t);

function sh(cmd, args, opt) {
  const r = spawnSync(cmd, args, Object.assign({ cwd: ROOT, encoding: 'utf8', env: process.env }, opt || {}));
  return { code: r.status, out: (r.stdout || '') + (r.stderr || '') };
}

// バックアップの、JS/CSSの登録を、書き戻す（ファイルは、バックアップの files/ から、上げ直す）
async function restoreCustomize(k, app, backupDir) {
  const c = JSON.parse(fs.readFileSync(path.join(backupDir, 'customize_live.json'), 'utf8'));
  const conv = async (env, kind) => {
    const out = [];
    for (const it of (c[env] && c[env][kind]) || []) {
      if (it.type === 'URL') { out.push(it); continue; }
      const f = path.join(backupDir, 'files', 'live_' + env + '_' + kind + '_' + it.file.name.replace(/[\\/:*?"<>|]/g, '_'));
      if (!fs.existsSync(f)) throw new Error('バックアップにファイルがありません：' + it.file.name);
      out.push({ type: 'FILE', file: { fileKey: await k.upload(fs.readFileSync(f), it.file.name, 'text/javascript') } });
    }
    return out;
  };
  const body = { app, desktop: { js: await conv('desktop', 'js'), css: await conv('desktop', 'css') }, mobile: { js: await conv('mobile', 'js'), css: await conv('mobile', 'css') } };
  if (c.scope) body.scope = c.scope;
  await k.api('PUT', '/k/v1/preview/app/customize', body);
  await k.api('POST', '/k/v1/preview/app/deploy', { apps: [{ app }] });
  await k.waitDeploy(app, String(app));
}

async function snapshot(k, app) {
  const layout = (await k.api('GET', '/k/v1/app/form/layout', { app })).layout;
  const fields = Object.keys((await k.api('GET', '/k/v1/app/form/fields', { app })).properties).sort();
  return { layout: JSON.stringify(layout), fields: JSON.stringify(fields), rows: layout.length, nfields: fields.length };
}

async function main() {
  const app = Number(arg('app'));
  if (!MANIFESTS[app]) { console.error('対象は、--app=575（案件くん）か --app=879（書類マスタ）だけです。ほかは、利用者の許可が、必要です'); process.exit(1); }
  const dry = has('dry');
  say('【安全な反映】アプリ ' + app + '（' + (app === 575 ? '案件くん' : '書類マスタ') + '）のJavaScriptだけ');

  // ① 自動テスト
  if (has('tests-already-passed')) say('① 自動テスト：（このスクリプト自体のテストのため、省略）');
  else {
    const t = sh('node', ['--test'].concat(fs.readdirSync(path.join(ROOT, 'test')).filter((f) => f.endsWith('.test.js')).map((f) => 'test/' + f)));
    const sum = (t.out.match(/# tests (\d+)[\s\S]*?# pass (\d+)[\s\S]*?# fail (\d+)/) || []);
    if (t.code !== 0) { say('① 自動テスト：失敗（' + (sum[3] || '?') + '件）。反映を止めました。'); process.exit(10); }
    say('① 自動テスト：すべて通りました（' + (sum[2] || '?') + '件）');
  }

  const k = client();
  const before = await snapshot(k, app);
  say('   反映前のフォーム：' + before.rows + '行／' + before.nfields + '項目');

  // ③ を先に確認（変更が始まる前に止める）。apply-customize の、確認だけの実行で、編集中の変更を調べる
  const dryRun = sh('node', ['scripts/apply-customize.js', '--manifest=' + MANIFESTS[app]]);
  if (/編集中/.test(dryRun.out)) { say('③ フォームの編集中の変更があります。反映を止めました（何も変えていません）。\n' + (dryRun.out.match(/【注意】[\s\S]*?(?=\n\n|\n※|$)/) || [''])[0]); process.exit(11); }
  if (dryRun.code !== 0) { say('登録内容の確認で失敗しました。反映を止めました。\n' + dryRun.out); process.exit(12); }
  say('③ フォームの編集中の変更：なし');
  if (dry) { say('\n（--dry：ここまでの確認だけ。何も変えていません）'); return; }

  // ② バックアップ
  const b = sh('node', ['scripts/backup-app.js', String(app)]);
  const dir = (b.out.match(/バックアップ先：(.+?)（/) || [])[1];
  if (b.code !== 0 || !dir) { say('② バックアップに失敗しました。反映を止めました。\n' + b.out); process.exit(13); }
  say('② バックアップ：' + dir);

  // ④ 登録・公開
  const ap = sh('node', ['scripts/apply-customize.js', '--manifest=' + MANIFESTS[app], '--yes']);
  if (ap.code !== 0) { say('④ 登録に失敗しました（終了コード ' + ap.code + '）。\n' + ap.out); await rollback(k, app, dir, '登録の失敗'); process.exit(14); }
  say('④ JavaScriptを登録して公開しました');

  // ⑤ 確認
  const problems = [];
  const after = await snapshot(k, app);
  if (after.layout !== before.layout) problems.push('フォームの並びが、反映前と違います');
  if (after.fields !== before.fields) problems.push('フォームの項目が、反映前と違います');
  const c = await k.api('GET', '/k/v1/app/customize', { app });
  const manifest = JSON.parse(fs.readFileSync(path.join(ROOT, MANIFESTS[app]), 'utf8'));
  for (const env of ['desktop', 'mobile']) {
    for (const rel of ((manifest[env] && manifest[env].js) || []).filter((x) => !/^https?:/.test(x))) {
      const name = path.basename(rel);
      const it = ((c[env] && c[env].js) || []).find((x) => x.file && x.file.name === name);
      if (!it) { problems.push(env + ' に ' + name + ' が登録されていません'); continue; }
      const live = await k.download(it.file.fileKey);
      if (!Buffer.from(live).equals(fs.readFileSync(path.join(ROOT, rel)))) problems.push(env + ' の ' + name + ' が、手元と違います');
    }
  }
  if (problems.length) { say('⑤ 確認：問題があります\n  ・' + problems.join('\n  ・')); await rollback(k, app, dir, problems.join('、')); process.exit(15); }
  say('⑤ 確認：フォームの並び・項目は、反映前と同じ（' + after.rows + '行／' + after.nfields + '項目）。登録したファイルは、手元と同じ');
  say('\n完了しました。画面を再読み込み（Ctrl+F5）して確認してください。戻すときの、バックアップ：' + dir);
}

async function rollback(k, app, dir, why) {
  say('⑥ 自動で、反映前のJavaScriptの登録に、戻します（理由：' + why + '）');
  try { await restoreCustomize(k, app, dir); say('   戻しました。バックアップ：' + dir); }
  catch (e) { say('   自動で戻せませんでした：' + (e.message || e) + '\n   手動で戻してください。バックアップ：' + dir); }
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
