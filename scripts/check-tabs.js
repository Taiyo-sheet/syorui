'use strict';
/*
 * 案件くんのフォームの「タブの見出しの行番号」が、正しい状態（基準）と同じかを、確かめる（読み取りだけ。何も変えない）。
 *
 *   npm run check-tabs                 … いまの案件くん(575)を、基準と比べる
 *   npm run check-tabs -- --save       … いまの状態を、新しい基準として保存する（フォームとプラグインの数字を、そろえたあとに使う）
 *
 * なぜ必要か：「タブ表示」プラグインは、タブの範囲を「フォームの何行目から」で覚えている。
 *            フォームに行を足す・減らす・動かすと、後ろのタブが、ずれる。フォームを触る前と後に、これで確かめる。
 * 基準：scripts/tab-rows.json（初期は、2026-10-24 朝の、ずれていなかった状態）
 */
const fs = require('fs');
const path = require('path');
const { client } = require('./lib');

const BASE = path.join(__dirname, 'tab-rows.json');
const clean = (f) => String(f.label || '').replace(/<[^>]*>/g, '').replace(/&[^;]+;/g, '').replace(/[^　-鿿A-Za-z0-9]/g, '');
const TABS = [['営業情報', /営業情報/], ['引継ぎ担当者情報', /引継ぎ.*担当者/], ['計画', /^計画/], ['リソース管理工事部', /リソース管理工事部/], ['引継ぎ工事部', /引継ぎ.*工事部/], ['書類関係', /^書類関係/], ['施工店管理', /施工店管理/], ['請求情報', /請求情報/], ['特許関連', /特許関連/], ['手数料管理', /手数料管理/], ['施工協会管理', /施工(協)?会管理/], ['依頼管理履歴', /^依頼管理履歴$/]];

function rowsOf(layout) {
  const out = {};
  layout.forEach((r, i) => {
    if (r.type !== 'ROW') return;
    r.fields.forEach((f) => {
      if (f.type !== 'LABEL') return;
      const t = clean(f);
      TABS.forEach(([name, re]) => { if (!out[name] && i > 10 && re.test(t) && r.fields.length <= 3) out[name] = i + 1; });
    });
  });
  return { rows: out, total: layout.length };
}

(async () => {
  const k = client();
  const app = 575;   // 案件くん
  const layout = (await k.api('GET', '/k/v1/app/form/layout', { app })).layout;
  const now = rowsOf(layout);
  if (process.argv.includes('--save')) {
    fs.writeFileSync(BASE, JSON.stringify({ savedAt: new Date().toISOString().slice(0, 10), ...now }, null, 2) + '\n');
    console.log('いまの状態を、基準として、保存しました：' + BASE);
    return;
  }
  const base = JSON.parse(fs.readFileSync(BASE, 'utf8'));
  let bad = 0;
  console.log('タブの見出しの行番号（基準＝' + base.savedAt + '）');
  TABS.forEach(([name]) => {
    const b = base.rows[name], n = now.rows[name];
    const mark = b === n ? 'OK' : 'ずれ';
    if (b !== n) bad++;
    console.log('  ' + name.padEnd(14, '　') + ' 基準 ' + String(b ?? '-').padStart(4) + '   いま ' + String(n ?? '-').padStart(4) + '   ' + mark);
  });
  console.log('  総行数          基準 ' + String(base.total).padStart(4) + '   いま ' + String(now.total).padStart(4) + '   ' + (base.total === now.total ? 'OK' : 'ずれ'));
  if (base.total !== now.total) bad++;
  console.log(bad ? '\n※ ずれがあります。「タブ表示」プラグインの行番号と、合っているか確認してください。' : '\n基準と同じです。');
  process.exit(bad ? 2 : 0);
})().catch((e) => { console.error(e.message || e); process.exit(1); });
