'use strict';
/*
 * 現場カルテ(826)に、案件くん(575)の情報を、コピーする。案件番号が同じものを、つなぐ。
 *   node scripts/sync-karte.js         … 確認だけ（何も書き込まない）。変わる内容の一覧を出す
 *   node scripts/sync-karte.js --yes   … 実際に、現場カルテに書き込む（利用者の許可が要る。先に、現場カルテのバックアップを取る）
 *
 * 反映する項目（案件くんの値が、空でないときだけ。違うときだけ、書き換える）
 *   元請 ← 元請_会社名_0／元請担当 ← 元請_担当名／現調日 ← 現調予定日／現調担当 ← 現調担当／
 *   都道府県 ← 都道府県／工種・工程 ← 施工種別／施工完了 ← 施工完了日
 * 反映しないもの：案件くんに、対応する項目がないもの（希望月・仕様決定・資材発注・ヒーター製造・施工会社・施工開始・施工状況 など）は、現場カルテに、直接入力する
 * 現場カルテの、金額・実績・請求履歴などの項目は、触らない
 */
const { client } = require('./lib');
const APP_KARTE = 826, APP_ANKEN = 575;
const MAP = [['元請', '元請_会社名_0'], ['元請担当', '元請_担当名'], ['現調日', '現調予定日'], ['現調担当', '現調担当'], ['都道府県', '都道府県'], ['工種・工程', '施工種別'], ['施工完了', '施工完了日']];
const yes = process.argv.includes('--yes');
const norm = (t) => String(t == null ? '' : t).replace(/[\s　]+/g, ' ').trim();
const textOf = (f) => { if (!f) return ''; const v = f.value; if (Array.isArray(v)) return v.map((x) => (x && (x.name || x)) || '').join('、'); return norm(v); };

async function main() {
  const k = client();
  const fp = (await k.api('GET', '/k/v1/app/form/fields', { app: APP_KARTE })).properties;
  const codeOf = (label) => { const e = Object.entries(fp).find(([c, p]) => String(p.label).replace(/[\s　]/g, '') === label.replace(/[\s　]/g, '')); if (!e) throw new Error('現場カルテに「' + label + '」がありません'); return e[0]; };
  const keyCode = Object.entries(fp).find(([c, p]) => p.lookup)[0];   // 案件番号（案件くんから選択）
  const recs = (await k.api('GET', '/k/v1/records', { app: APP_KARTE, query: 'limit 500' })).records;
  const updates = []; let compared = 0, changed = 0, missing = 0;
  console.log((yes ? '【書き込みます】' : '【確認だけ・書き込みなし】') + ' 現場カルテ ' + recs.length + ' 件 ← 案件くん');
  for (const r of recs) {
    const no = norm(r[keyCode].value);
    if (!no) continue;
    const q = (await k.api('GET', '/k/v1/records', { app: APP_ANKEN, query: '案件番号_0 = "' + no.replace(/["\\]/g, '') + '" limit 1' })).records[0];
    if (!q) { missing++; console.log('[' + no + '] 案件くんに、同じ案件番号が、ありません（反映なし）'); continue; }
    const rec = {}; const lines = [];
    for (const [label, from] of MAP) {
      const code = codeOf(label), now = textOf(r[code]).replace(/^担当：$/, ''), nw = textOf(q[from]);
      compared++;
      if (nw && nw !== now) { changed++; lines.push('   ' + label.padEnd(8) + '| ' + (now || '（空）').slice(0, 14).padEnd(14) + ' → ' + nw.slice(0, 24)); rec[code] = { value: q[from].type === 'USER_SELECT' ? nw : q[from].value }; }
    }
    if (lines.length) { console.log('[' + no + '] ' + norm(r['proj_name'] && r['proj_name'].value).slice(0, 16) + '\n' + lines.join('\n')); updates.push({ id: r.$id.value, record: rec }); }
  }
  console.log('\n比べた項目 ' + compared + '／変わる項目 ' + changed + '／変わる現場カルテ ' + updates.length + ' 件／案件くんにない案件番号 ' + missing + ' 件');
  if (!yes) { console.log('（確認だけです。反映するときは、利用者の許可をもらってから --yes を付けます）'); return; }
  for (let i = 0; i < updates.length; i += 100) await k.api('PUT', '/k/v1/records', { app: APP_KARTE, records: updates.slice(i, i + 100) });
  console.log('現場カルテ ' + updates.length + ' 件に、書き込みました');
}
main().catch((e) => { console.error(e.message || e); process.exit(1); });
