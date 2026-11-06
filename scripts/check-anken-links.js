'use strict';
/*
 * 案件番号が、アプリのあいだで、そろっているかを調べる（読み取りだけ。何も変えない）。  npm run check-links
 *   ・請求くん(829)・発注くん(824)・見積くん(608)・予算02(735)・現場カルテ(826)の「案件番号」が、案件くん(575)に、あるか
 *   ・案件くんで、「決算期＋案件番号」が、TS○○-0000 の形になっているか（決算期が空だと、「-0073」のようになる）
 */
const { client } = require('./lib');
const APPS = { 829: '請求くん', 824: '発注くん', 608: '見積くん', 735: '予算02', 826: '現場カルテ' };
async function all(k, app) { const out = []; for (let off = 0; ; off += 500) { const r = await k.api('GET', '/k/v1/records', { app, query: 'limit 500 offset ' + off }); out.push(...r.records); if (r.records.length < 500) break; } return out; }
(async () => {
  const k = client();
  const anken = await all(k, 575);
  const have = new Set(anken.map((r) => r['案件番号_0'].value).filter(Boolean));
  let problems = 0;
  for (const [app, name] of Object.entries(APPS)) {
    const recs = await all(k, Number(app));
    const miss = {}; let blank = 0;
    recs.forEach((r) => { const a = r.anken && r.anken.value; if (!a) { blank++; return; } if (!have.has(a)) miss[a] = (miss[a] || 0) + 1; });
    const m = Object.entries(miss);
    console.log((m.length ? '✕ ' : '○ ') + name + '(' + app + ') ' + recs.length + '件：案件くんにない案件番号 ' + (m.length ? m.map(([a, n]) => a + '（' + n + '件）').join('、') : 'なし') + (blank ? '／案件番号が空 ' + blank + '件' : ''));
    if (m.length) problems++;
  }
  const noKi = anken.filter((r) => !/^TS\d+-\d+$/.test(r['決算期_案件番号'].value) && !/^DEMO/.test(r['案件番号_0'].value));
  console.log((noKi.length ? '✕ ' : '○ ') + '案件くん(575) ' + anken.length + '件：「決算期＋案件番号」が、TS○○-0000 の形でない案件 ' + noKi.length + ' 件' + (noKi.length ? '（例：' + noKi.slice(0, 4).map((r) => JSON.stringify(r['決算期_案件番号'].value)).join('、') + '）' : ''));
  if (noKi.length) problems++;
  console.log(problems ? '\n要確認：' + problems + ' 項目' : '\n問題なし');
  process.exit(problems ? 2 : 0);
})().catch((e) => { console.error(e.message || e); process.exit(1); });
