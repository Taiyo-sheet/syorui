'use strict';
/*
 * 書類マスタの「雛形」を、取り出す・差し替える（Claude に「この書類のここを直して」と頼むときの道具）。
 *
 *   node scripts/template-tool.js list                                   … 書類ごとの雛形の有無と、{{ }} の一覧
 *   node scripts/template-tool.js get --doc=書類名 [--out=フォルダ]        … いまの雛形を取り出す（省略時は templates-work/）
 *   node scripts/template-tool.js put --doc=書類名 --file=ファイル          … 差し替える内容を表示するだけ（何も変えない）
 *   node scripts/template-tool.js put --doc=書類名 --file=ファイル --yes    … 実際に差し替える
 *
 * 守ること
 *   ・差し替えるとき、古い雛形は消さず、「旧版_日付_」をつけて、同じ書類の「雛形」欄に残す（先頭が、使われる雛形）
 *   ・旧版は、最新の3つまで残す（それより古いものだけ、外す）
 *   ・差し替えのあと、新しい雛形が、先頭にあること・{{ }} が読めることを確かめる
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');
const { client } = require('./lib');
const lib = require('../customize/doc-button.js');

const KEEP = 3;
const arg = (n) => ((process.argv.find((a) => a.startsWith('--' + n + '=')) || '').split('=').slice(1).join('='));
const extOf = (n) => ((String(n).match(/\.([^.]+)$/) || [])[1] || '').toLowerCase();
const stamp = () => new Date().toISOString().slice(0, 10).replace(/-/g, '');

async function placeholders(buf, ext) {
  if (!['docx', 'xlsx', 'pptx'].includes(ext)) return null;
  const text = await lib.extractText(JSZip, buf, ext);
  return [...new Set((text.match(/\{\{[^{}]+\}\}/g) || []))];
}

async function main() {
  const cmd = process.argv[2];
  const k = client();
  const apps = await k.findApps('書類マスタ');
  if (apps.length !== 1) throw new Error('「書類マスタ」アプリが' + (apps.length ? apps.length + '個あります' : '見つかりません'));
  const app = apps[0].appId;
  const all = (await k.api('GET', '/k/v1/records', { app, query: 'order by 表示順 asc limit 500' })).records;

  if (cmd === 'list') {
    for (const r of all) {
      const f = (r['雛形'].value || [])[0];
      let info = '雛形なし';
      if (f) { const ph = await placeholders(await k.download(f.fileKey), extOf(f.name)); info = f.name + ' ｜' + (ph === null ? '差し込みなし' : ph.length ? ph.join(' ') : '{{ }}なし（雛形のまま使う）'); }
      console.log(r['書類名'].value + ' ｜ ' + info);
    }
    return;
  }

  const doc = arg('doc');
  const rec = all.filter((r) => r['書類名'].value === doc);
  if (!doc || rec.length !== 1) throw new Error('書類名「' + doc + '」が、書類マスタに' + rec.length + '件あります（1件のはずです）');
  const files = rec[0]['雛形'].value || [];

  if (cmd === 'get') {
    if (!files.length) throw new Error('「' + doc + '」には、雛形がありません');
    const dir = path.resolve(arg('out') || 'templates-work'); fs.mkdirSync(dir, { recursive: true });
    const to = path.join(dir, files[0].name); fs.writeFileSync(to, await k.download(files[0].fileKey));
    console.log('取り出しました：' + to);
    return;
  }

  if (cmd === 'put') {
    const file = arg('file'), go = process.argv.includes('--yes');
    if (!file || !fs.existsSync(file)) throw new Error('--file= で、差し替えるファイルを指定してください');
    const ext = extOf(file);
    const buf = fs.readFileSync(file);
    const ph = await placeholders(buf, ext);
    console.log('「' + doc + '」の雛形を、' + path.basename(file) + ' に差し替えます。');
    console.log('新しい雛形の {{ }}：' + (ph === null ? '（' + ext + '：差し込みなし）' : ph.length ? ph.join(' ') : 'なし（雛形のまま使う書類になります）'));
    console.log('いまの雛形：' + (files.length ? files.map((f) => f.name).join('、') : 'なし') + '（古いものは、「旧版_' + stamp() + '_」をつけて、残します。最新' + KEEP + '個まで）');
    if (!go) { console.log('\n※ これは確認だけです。実際に差し替えるには、最後に --yes を付けてください。'); return; }

    // 古い雛形を、旧版として残す（ダウンロード → 名前を変えて、アップロード）
    const keep = [];
    for (const f of files) {
      // すべて、取り出して、上げ直す：すでに保存されているファイルは、キントーンが、新しく上げたファイルより先に並べるため、
      // 上げ直さないと、新しい雛形が、先頭に来ないことがある
      const old = /^旧版_/.test(f.name);
      const name = old ? f.name : '旧版_' + stamp() + '_' + f.name;
      keep.push({ fileKey: await k.upload(await k.download(f.fileKey), name, 'application/octet-stream'), name });
    }
    const newKey = await k.upload(buf, path.basename(file), 'application/octet-stream');
    const next = [{ fileKey: newKey }].concat(keep.slice(0, KEEP).map((x) => ({ fileKey: x.fileKey })));
    await k.api('PUT', '/k/v1/record', { app, id: rec[0].$id.value, record: { 雛形: { value: next } } });

    const back = (await k.api('GET', '/k/v1/record', { app, id: rec[0].$id.value })).record['雛形'].value;
    const first = back[0];
    const ok = first && first.name === path.basename(file) && (ph === null || JSON.stringify(await placeholders(await k.download(first.fileKey), ext)) === JSON.stringify(ph));
    console.log((ok ? 'OK ' : 'NG ') + '差し替え後の雛形：' + back.map((f) => f.name).join('、'));
    process.exit(ok ? 0 : 4);
  }

  console.error('使い方：node scripts/template-tool.js list | get --doc=書類名 | put --doc=書類名 --file=ファイル [--yes]');
  process.exit(1);
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
