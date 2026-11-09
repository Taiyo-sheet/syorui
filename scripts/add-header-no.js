'use strict';
/*
 * Word雛形のヘッダーに、「No.{{書類番号}}」（左）と「{{版}}」（右・ページ番号の前）を足す。
 *
 *   node scripts/add-header-no.js 書類名...            … 取り出して、直した雛形を、作業フォルダに作り、見え方を確認する（書類マスタは変えない）
 *   node scripts/add-header-no.js 書類名... --put      … 確認のあと、書類マスタに差し替える（template-tool put --yes：古い雛形は残る）
 *
 * 利用者の「OK」：2026-11-09（作業手順書の見本を見て「A」）。まず、作業手順書・施工計画書・仕様書・リスクアセスメント。
 */
const fs = require('fs');
const path = require('path');
const os = require('os');
const { spawnSync } = require('child_process');
const JSZip = require('jszip');
const lib = require('../customize/doc-button.js');

const ROOT = path.join(__dirname, '..');
const RPR = '<w:rPr><w:rFonts w:ascii="游ゴシック" w:cs="游ゴシック" w:eastAsia="游ゴシック" w:hAnsi="游ゴシック"/><w:b/><w:color w:val="1F3864"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>';

async function patch(buf) {
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files).filter((n) => /^word\/header\d*\.xml$/.test(n));
  let changed = 0;
  for (const n of names) {
    let t = await zip.file(n).async('string');
    if (t.includes('{{書類番号}}')) continue;
    const p0 = t.indexOf('<w:p>') >= 0 ? t.indexOf('<w:p>') : t.indexOf('<w:p ');
    if (p0 < 0) continue;
    const pEnd = t.indexOf('</w:p>', p0);
    let para = t.slice(p0, pEnd);
    const left = '<w:r>' + RPR + '<w:t xml:space="preserve">No.{{書類番号}}　</w:t></w:r>';
    const ppr = para.indexOf('</w:pPr>');
    para = ppr >= 0 ? para.slice(0, ppr + 8) + left + para.slice(ppr + 8) : para.replace(/^<w:p[^>]*>/, (m) => m + left);
    const tabRun = '<w:t xml:space="preserve">\t</w:t><w:fldChar w:fldCharType="begin"/>';
    if (para.includes(tabRun)) para = para.replace(tabRun, '<w:t xml:space="preserve">\t{{版}}　　</w:t><w:fldChar w:fldCharType="begin"/>');
    else if (para.includes('<w:tab/></w:r>')) para = para.replace('<w:tab/></w:r>', '<w:tab/></w:r><w:r><w:rPr><w:color w:val="777777"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr><w:t xml:space="preserve">{{版}}　　</w:t></w:r>');   // タブが、別の部品のとき
    t = t.slice(0, p0) + para + t.slice(pEnd);
    zip.file(n, t); changed++;
  }
  return { data: await zip.generateAsync({ type: 'nodebuffer' }), changed };
}

(async () => {
  const args = process.argv.slice(2), put = args.includes('--put'), docs = args.filter((a) => !a.startsWith('--'));
  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'hdr-'));
  for (const doc of docs) {
    const dir = path.join(work, doc.replace(/[\/\\:*?"<>|（）]/g, '_')); fs.mkdirSync(dir);
    const g = spawnSync('node', [path.join(__dirname, 'template-tool.js'), 'get', '--doc=' + doc, '--out=' + dir], { encoding: 'utf8' });
    const f = fs.readdirSync(dir).find((x) => /\.docx$/i.test(x));
    if (!f) { console.log('NG ' + doc + '：Wordの雛形が取り出せません'); continue; }
    const r = await patch(fs.readFileSync(path.join(dir, f)));
    const outDir = path.join(dir, 'out'); fs.mkdirSync(outDir, { recursive: true }); const out = path.join(outDir, f); fs.writeFileSync(out, r.data);   // ファイル名は、もとのまま（new_ などを付けない）
    // 見え方の確認：置き換えて、文字を読む
    const filled = await lib.fillTemplate(JSZip, r.data, 'docx', { '書類番号': '09', '版': '第1版', '工事名称': 'テスト工事' });
    const z = await JSZip.loadAsync(Buffer.from(filled.data));
    const hdr = (await Promise.all(Object.keys(z.files).filter((n) => /^word\/header\d*\.xml$/.test(n)).map((n) => z.file(n).async('string')))).join('').replace(/<[^>]*>/g, ' ');
    const ok = r.changed > 0 && /No\.09/.test(hdr) && /第1版/.test(hdr);
    console.log((ok ? 'OK ' : 'NG ') + doc + '：ヘッダー ' + r.changed + ' か所／No.09 ' + (/No\.09/.test(hdr) ? 'あり' : 'なし') + '／第1版 ' + (/第1版/.test(hdr) ? 'あり' : 'なし') + '  → ' + out);
    if (ok && put) {
      const pr = spawnSync('node', [path.join(__dirname, 'template-tool.js'), 'put', '--doc=' + doc, '--file=' + out, '--yes'], { encoding: 'utf8' });
      console.log('   差し替え：' + (pr.stdout || pr.stderr).trim().split('\n').slice(-2).join(' / '));
    }
  }
})().catch((e) => { console.error(e.message || e); process.exit(1); });
