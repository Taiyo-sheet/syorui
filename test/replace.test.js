'use strict';
const test = require('node:test');
const assert = require('node:assert');
const JSZip = require('jszip');
const lib = require('../customize/doc-button.js');
const { build } = require('../scripts/make-test-templates.js');

const VALUES = {
  '案件名': '東京Aビル 改修工事',
  '現場住所': '東京都千代田区1-1',
  '発注者': '株式会社<テスト>&サンプル',
  '備考': '1行目\n2行目 & <注意>',
  '作成日': '2026年10月19日',
  '版': 'v2',
  '書類名': '打合せ議事次第',
};

async function fill(buf, ext, values = VALUES) {
  const out = await lib.fillTemplate(JSZip, buf, ext, values);
  const zip = await JSZip.loadAsync(out.data);
  const texts = {};
  for (const n of Object.keys(zip.files)) if (/\.xml$/.test(n)) texts[n] = await zip.file(n).async('string');
  return { out, texts };
}
const strip = (x) => x.replace(/<[^>]*>/g, '');

test('docx：分割された {{ }}・数値表記の項目名・特殊文字・ヘッダー', async () => {
  const { docx } = await build();
  const { out, texts } = await fill(docx, 'docx');
  const doc = texts['word/document.xml'];
  assert.match(strip(doc), /案件名：東京Aビル 改修工事/);
  assert.match(strip(doc), /現場：東京都千代田区1-1/);
  assert.match(doc, /<w:p><w:r><w:t>東京Aビル 改修工事<\/w:t><\/w:r><\/w:p>/, '数値表記の {{&#26696;件名}} が置き換わる');
  assert.match(doc, /1行目<\/w:t><w:br\/><w:t xml:space="preserve">2行目 &amp; &lt;注意&gt;/);
  assert.match(doc, /\{\{存在しない項目\}\}/, '見つからない名前は残す');
  assert.ok(out.unknown.includes('存在しない項目'));
  assert.match(strip(doc), /打合せ議事次第 v2 2026年10月19日/);
  assert.match(strip(texts['word/header1.xml']), /ヘッダー：東京Aビル/);
});

test('docx：段落をまたぐ {{ }} は触らず、書類の構造を壊さない', async () => {
  const { docx } = await build();
  const { out, texts } = await fill(docx, 'docx');
  const doc = texts['word/document.xml'];
  assert.match(doc, /<w:t>\{\{案件<\/w:t><\/w:r><\/w:p><w:p><w:r><w:t>名\}\}<\/w:t>/);
  assert.strictEqual((doc.match(/<w:p>/g) || []).length, (doc.match(/<\/w:p>/g) || []).length);
  assert.ok(out.unknown.includes('案件名'), '置き換えできなかったことを知らせる');
});

test('docx：書式（太字・斜体）の塊をまたいでも、最初の塊の書式で入る', async () => {
  const { docx } = await build();
  const { texts } = await fill(docx, 'docx');
  assert.match(texts['word/document.xml'], /<w:rPr><w:b\/><\/w:rPr><w:t>現場：東京都千代田区1-1<\/w:t>/);
});

test('xlsx：共有文字列・書式分割・ふりがな付き・直接入力', async () => {
  const { xlsx } = await build();
  const { out, texts } = await fill(xlsx, 'xlsx');
  const ss = texts['xl/sharedStrings.xml'];
  assert.match(ss, /<si><t>東京Aビル 改修工事<\/t><\/si>/);
  assert.match(strip(ss), /東京都千代田区1-1/);
  assert.match(ss, /株式会社&lt;テスト&gt;&amp;サンプル/);
  assert.ok(!/ハッチュウシャ/.test(ss), '{{ }} を含むセルのふりがなは外す');
  assert.match(ss, /コテイ/, 'ほかのセルのふりがなは触らない');
  const sheet = texts['xl/worksheets/sheet1.xml'];
  assert.match(sheet, /直接入力：1行目\n2行目 &amp; &lt;注意&gt;/);
  assert.match(sheet, /\{\{存在しない項目\}\}/);
  assert.ok(out.unknown.includes('存在しない項目'));
});

test('pptx：分割・改行（2行目以降も同じ書式）', async () => {
  const { pptx } = await build();
  const { texts } = await fill(pptx, 'pptx');
  const s = texts['ppt/slides/slide1.xml'];
  assert.match(s, /案件名：東京Aビル 改修工事/);
  assert.match(s, /1行目<\/a:t><\/a:r><a:br\/><a:r><a:rPr sz="1800"\/><a:t>2行目 &amp; &lt;注意&gt;/);
});

test('制御文字を含む値でも、書類のXMLが壊れない', async () => {
  const { docx } = await build();
  const { texts } = await fill(docx, 'docx', Object.assign({}, VALUES, { '案件名': 'A\u000bB\u0000C' }));
  assert.ok(!/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(texts['word/document.xml']));
  assert.match(strip(texts['word/document.xml']), /案件名：ABC/);
});

test('PDFなどはそのまま返す', async () => {
  const buf = Buffer.from('%PDF-1.4');
  const out = await lib.fillTemplate(JSZip, buf, 'pdf', VALUES);
  assert.strictEqual(out.data, buf);
});

test('fieldText：桁区切りは「桁区切りを表示する」項目だけ・日付・ユーザー', () => {
  assert.strictEqual(lib.fieldText({ type: 'NUMBER', value: '1234567' }, { digit: true }), '1,234,567');
  assert.strictEqual(lib.fieldText({ type: 'NUMBER', value: '2026' }, {}), '2026');
  assert.strictEqual(lib.fieldText({ type: 'DATE', value: '2026-10-05' }), '2026年10月5日');
  assert.strictEqual(lib.fieldText({ type: 'USER_SELECT', value: [{ name: '山田' }, { name: '佐藤' }] }), '山田、佐藤');
  assert.strictEqual(lib.fieldText({ type: 'CHECK_BOX', value: ['A', 'B'] }), 'A、B');
  assert.strictEqual(lib.fieldText({ type: 'SUBTABLE', value: [] }), '');
});

test('ファイル名：使えない文字・改行を置き換える', () => {
  assert.strictEqual(lib.buildFilename('A/B:C\n工事', '見積書', 3, 'xlsx'), 'A_B_C_工事_見積書_v3.xlsx');
});
test('クエリの引用符をエスケープする', () => {
  assert.strictEqual(lib.quote('a"b\\c'), '"a\\"b\\\\c"');
});

test('hasPlaceholders：{{ }}があれば true、なければ（雛形のまま）false。分割・PDFも', async () => {
  const b = await build();
  assert.strictEqual(await lib.hasPlaceholders(JSZip, b.docx, 'docx'), true);
  assert.strictEqual(await lib.hasPlaceholders(JSZip, b.xlsx, 'xlsx'), true);
  assert.strictEqual(await lib.hasPlaceholders(JSZip, b.pptx, 'pptx'), true);
  assert.strictEqual(await lib.hasPlaceholders(JSZip, Buffer.from('%PDF-1.4'), 'pdf'), false);
  // {{ }} を含まない普通の文書は false
  const z = await JSZip.loadAsync(b.docx);
  z.file('word/document.xml', '<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>共通の手引き（案件の情報は入りません）</w:t></w:r></w:p></w:body></w:document>');
  z.file('word/header1.xml', '<w:hdr xmlns:w="x"><w:p><w:r><w:t>ヘッダー</w:t></w:r></w:p></w:hdr>');
  assert.strictEqual(await lib.hasPlaceholders(JSZip, await z.generateAsync({ type: 'nodebuffer' }), 'docx'), false);
  // {{ が1つだけ・片方だけ（閉じていない）は、{{ }} とは数えない
  const z2 = await JSZip.loadAsync(b.docx);
  z2.file('word/document.xml', '<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>波括弧 { と {{ だけ</w:t></w:r></w:p></w:body></w:document>');
  z2.file('word/header1.xml', '<w:hdr xmlns:w="x"><w:p><w:r><w:t>ヘッダー</w:t></w:r></w:p></w:hdr>');
  assert.strictEqual(await lib.hasPlaceholders(JSZip, await z2.generateAsync({ type: 'nodebuffer' }), 'docx'), false);
});

test('guessTarget：「〇〇橋」「高架橋」は橋梁、地名の「橋本」などは一般道路', () => {
  for (const n of ['東京高架橋 補修工事', '第二多摩川橋 床版工事', '新大橋', '跨線橋 ロードヒーティング', '〇〇橋梁 点検', '隅田川橋 左岸']) assert.strictEqual(lib.guessTarget(n), '橋梁', n);
  for (const n of ['国道17号 ロードヒーティング', '橋本市道 舗装工事', '橋口交差点 改良', '東京Aビル 改修工事', '', undefined]) assert.strictEqual(lib.guessTarget(n), '一般道路', String(n));
});
