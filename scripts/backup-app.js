'use strict';
/*
 * アプリの設定を、変更前のバックアップとして書き出す（読むだけ。キントーン側は何も変えない）。
 *
 *   node scripts/backup-app.js [アプリ番号]     … 省略すると KINTONE_APP_ID（案件くん）
 *
 * 書き出す先：backup/<アプリ番号>_<日時>/
 *   app.json … アプリの基本情報   fields_*.json … 項目   layout_*.json … レイアウト
 *   customize_*.json … JS/CSSの設定   views / settings / status … 一覧・設定・プロセス管理
 *   files/ … アップロード済みのJS/CSSファイル本体（ファイルで登録されているもの）
 *   *_live は公開中、*_preview は下書き（未公開の変更を含む）
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');

async function main() {
  const k = client();
  const appId = process.argv[2] || process.env.KINTONE_APP_ID;
  if (!appId) { console.error('アプリ番号を指定してください（例：node scripts/backup-app.js 575）'); process.exit(1); }
  const stamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\..*/, '').replace('T', '_');
  const dir = path.join(process.env.KINTONE_BACKUP_DIR || path.join(__dirname, '..', 'backup'), appId + '_' + stamp);
  fs.mkdirSync(path.join(dir, 'files'), { recursive: true });

  const items = [
    ['app', '/k/v1/app', { id: appId }],
    ['fields_live', '/k/v1/app/form/fields', { app: appId }],
    ['fields_preview', '/k/v1/preview/app/form/fields', { app: appId }],
    ['layout_live', '/k/v1/app/form/layout', { app: appId }],
    ['layout_preview', '/k/v1/preview/app/form/layout', { app: appId }],
    ['customize_live', '/k/v1/app/customize', { app: appId }],
    ['customize_preview', '/k/v1/preview/app/customize', { app: appId }],
    ['views_live', '/k/v1/app/views', { app: appId }],
    ['views_preview', '/k/v1/preview/app/views', { app: appId }],
    ['settings_live', '/k/v1/app/settings', { app: appId }],
    ['status_live', '/k/v1/app/status', { app: appId }],
  ];
  const ng = [];
  const customize = [];
  for (const [name, p, params] of items) {
    try {
      const r = await k.api('GET', p, params);
      fs.writeFileSync(path.join(dir, name + '.json'), JSON.stringify(r, null, 2));
      console.log('OK  ' + name);
      if (name.startsWith('customize_')) customize.push([name, r]);
    } catch (e) { ng.push(name); console.log('NG  ' + name + '：' + e.message); }
  }

  // ファイルで登録されているJS/CSSの本体も保存する
  let nfiles = 0;
  for (const [name, c] of customize) {
    for (const env of ['desktop', 'mobile']) {
      for (const kind of ['js', 'css']) {
        for (const it of (c[env] && c[env][kind]) || []) {
          if (it.type !== 'FILE') continue;
          try {
            const buf = await k.download(it.file.fileKey);
            fs.writeFileSync(path.join(dir, 'files', name.replace('customize_', '') + '_' + env + '_' + kind + '_' + it.file.name.replace(/[\\/:*?"<>|]/g, '_')), buf);
            nfiles++;
          } catch (e) { ng.push(it.file.name); console.log('NG  ファイル ' + it.file.name + '：' + e.message); }
        }
      }
    }
  }
  console.log('\nバックアップ先：' + dir + '（ファイル本体 ' + nfiles + '個）');
  if (ng.length) console.log('取れなかったもの：' + ng.join('、') + '（権限不足のことがあります。アプリ管理権限のあるユーザーで実行してください）');
  const must = ['fields_live', 'layout_live', 'customize_live'];
  if (must.some((m) => ng.includes(m))) { console.error('必須の項目が取れなかったため、バックアップ失敗として扱います'); process.exit(2); }
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
