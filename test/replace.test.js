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

test('checkText：残った {{ }}・空の欄・案件くんの値との違い・見本の文・空の日付を指摘する／問題なければ空', () => {
  const ok = '【打合せ議事録】\n■ 工事名称：東京Aビル\n■ 工事場所：千代田区1-1\n■ 日時：2026年10月22日\n';
  assert.deepStrictEqual(lib.checkText(ok, { 工事名称: '東京Aビル', 工事場所: '千代田区1-1' }), []);
  const bad = '■ 工事名称：大阪Bビル\n■ 工事場所：\n■ 担当：{{担当者}}\n（例：工程確認／課題整理）\n日時：　　年　　月　　日\n';
  const r = lib.checkText(bad, { 工事名称: '東京Aビル', 工事場所: '千代田区1-1' });
  assert.ok(r.some((x) => /置き換えられていない.*\{\{担当者\}\}/.test(x)), r.join('|'));
  assert.ok(r.some((x) => /値が空の欄.*工事場所/.test(x)), r.join('|'));
  assert.ok(r.some((x) => /工事名称.*違います.*大阪Bビル.*東京Aビル/.test(x)), r.join('|'));
  assert.ok(r.some((x) => /見本の文/.test(x)), r.join('|'));
  assert.ok(r.some((x) => /日付が空欄/.test(x)), r.join('|'));
});

test('tidyMemo：記号をそろえ、空行を取り除く', () => {
  assert.strictEqual(lib.tidyMemo('- a\n\n● b\n・c\n* d'), '・a\n・b\n・c\n・d');
});

test('checkText：案件くんの全項目を突き合わせる（違えば指摘／複数行・作成日などは対象外／空白の違いは無視）', () => {
  const v = { 発注者: '国土交通省 岐阜国道事務所', 備考: '1行目\n2行目', 作成日: '2026年10月22日', 元請: '株式会社サンプル' };
  assert.deepStrictEqual(lib.checkText('■ 発注者：国土交通省　岐阜国道事務所\n■ 備考：別の内容\n作成日：2020年1月1日\n■ 元請：株式会社サンプル', v), [], '空白の違い・複数行・作成日は、指摘しない');
  const r = lib.checkText('■ 発注者：国土交通省 高山国道事務所', v);
  assert.ok(r.some((x) => /「発注者」が、案件くんの値と違います（書類：国土交通省 高山国道事務所／案件くん：国土交通省 岐阜国道事務所）/.test(x)), r.join('|'));
});

test('summarizeLocal：文字起こしを、決定事項・ToDo・次回予定・出席者・場所・議題・内容に分ける（簡易）', () => {
  const t = 'えー、出席者：太田、國井、藤原\n場所：現場事務所\n工程の確認について。\n夜間規制の時間は、9時から17時で決定しました。\n来週までに太田さんが見積書を送付します。\nあの、現場は雪が多いそうです。\n次回は11月1日にやります。';
  const r = lib.summarizeLocal(t);
  assert.strictEqual(r['出席者'], '太田、國井、藤原'); assert.strictEqual(r['場所'], '現場事務所');
  assert.match(r['議題'], /工程の確認について/); assert.match(r['決定事項'], /夜間規制.*決定/);
  assert.match(r['ToDo'], /来週までに太田さんが見積書を送付/); assert.match(r['議事録メモ'], /雪が多い/); assert.doesNotMatch(r['議事録メモ'], /^あの/);
  assert.match(r['次回予定'], /11月1日/);
  assert.deepStrictEqual(lib.summarizeLocal(''), { '出席者': '', '場所': '', '議題': '', '議事録メモ': '', '決定事項': '', 'ToDo': '', '次回予定': '' });
});

test('summarizeLocal：見出し付き（Copilot などの、まとめ）は、見出しごとに振り分ける', () => {
  const t = '【出席者】太田、國井\n【場所】現場事務所\n【議題】\n・工程の確認\n【打合せ内容】\n・来週、現地調査\n・資料は金曜までに送付\n【決定事項】\n- 着工は11月上旬\n【ToDo】\n・10/25まで 太田 見積を出す\n【次回予定】次回は11月1日';
  const r = lib.summarizeLocal(t);
  assert.strictEqual(r['出席者'], '太田、國井'); assert.strictEqual(r['場所'], '現場事務所');
  assert.strictEqual(r['議題'], '・工程の確認'); assert.strictEqual(r['議事録メモ'], '・来週、現地調査\n・資料は金曜までに送付');
  assert.strictEqual(r['決定事項'], '・着工は11月上旬'); assert.strictEqual(r['ToDo'], '・10/25まで 太田 見積を出す'); assert.strictEqual(r['次回予定'], '次回は11月1日');
  const r2 = lib.summarizeLocal('## 決定事項\n1. 着工は11月上旬\n## アクションアイテム\n- 太田：見積を送る（10/25）');
  assert.strictEqual(r2['決定事項'], '・1. 着工は11月上旬'); assert.strictEqual(r2['ToDo'], '・太田：見積を送る（10/25）');
});

test('maskForAI：電話・メール・郵便番号・金額・「〇〇さん」の名前・指定した言葉を符号にし、返事で元に戻す', () => {
  const t = '太田さんが03-1234-5678に連絡。メールは ota@example.co.jp。〒950-0000 新潟。金額は1,200,000円。担当の藤原様と、サンプル建設の件。皆さんお願いします。';
  const m = lib.maskForAI(t, ['サンプル建設']);
  for (const s of ['太田', '03-1234-5678', 'ota@example.co.jp', '950-0000', '1,200,000円', '藤原', 'サンプル建設']) assert.ok(!m.masked.includes(s), s + ' が、残っている：' + m.masked);
  assert.match(m.masked, /【人物1】さん/); assert.match(m.masked, /皆さん/, '「皆さん」などは、人名として、扱わない');
  assert.strictEqual(m.unmask('・' + m.masked.match(/【人物1】/)[0] + 'さんが、' + m.masked.match(/【電話1】/)[0] + 'に連絡'), '・太田さんが、03-1234-5678に連絡');
  assert.strictEqual(m.unmask('【人物99】は、そのまま'), '【人物99】は、そのまま', '知らない符号は、そのまま');
  assert.strictEqual(lib.maskForAI('', []).count, 0);
});

test('formatMinutes：見出し付きの、ひとかたまりの文章にする（空は「なし」、出席者は「不明」）', () => {
  const t = lib.formatMinutes({ 出席者: '', 場所: '現場事務所', 議題: '・工程', 議事録メモ: '・現調', 決定事項: '・11月着工', ToDo: '', 次回予定: '' });
  assert.strictEqual(t, '【出席者】\n不明\n\n【場所】\n現場事務所\n\n【議題】\n・工程\n\n【打合せ内容】\n・現調\n\n【決定事項】\n・11月着工\n\n【ToDo】\nなし\n\n【次回予定】\nなし');
});

test('{{議事録本文}}（Word）：見出しは帯の段落、箇条書きは段落、ToDo（期限｜担当｜内容）は表になる。XMLが壊れない', async () => {
  const tpl = await build(); const z = await JSZip.loadAsync(tpl.docx);
  z.file('word/document.xml', '<w:document xmlns:w="x"><w:body><w:p><w:r><w:t>上の文</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>{{議事録</w:t></w:r><w:r><w:t>本文}}</w:t></w:r></w:p><w:p><w:r><w:t>下の文</w:t></w:r></w:p></w:body></w:document>');
  const buf = await z.generateAsync({ type: 'nodebuffer' });
  const body = '【決定事項】\n・規制は9時から17時\n\n【ToDo】\n・10/25｜太田｜見積を出す\n・10/27|國井|工程表を共有\n・メモだけの行\n\n【次回予定】';
  const out = await lib.fillTemplate(JSZip, buf, 'docx', { 議事録本文: body });
  const xml = await (await JSZip.loadAsync(out.data)).file('word/document.xml').async('string');
  const count = (re) => (xml.match(re) || []).length;
  assert.strictEqual(count(/<w:p[ >]/g), count(/<\/w:p>/g), '<w:p> の数が合う'); assert.strictEqual(count(/<w:tbl>/g), 1, 'ToDoの表は1つ');
  assert.strictEqual(count(/<w:tr>/g), 3, '見出し行＋2行（「期限｜担当｜内容」の形のものだけ）'); assert.strictEqual(count(/<w:r[ >]/g), count(/<\/w:r>/g));
  assert.strictEqual(count(/w:fill="1F3864"/g), 3, '見出し3つが帯');
  const flat = xml.replace(/<\/w:p>|<\/w:tc>/g, '\n').replace(/<[^>]*>/g, '');
  assert.match(flat, /上の文/); assert.match(flat, /下の文/); assert.match(flat, /■ 決定事項/); assert.match(flat, /・規制は9時から17時/);
  assert.match(flat, /見積を出す\s+太田\s+10\/25/, '表：内容・担当・期限'); assert.match(flat, /・メモだけの行/); assert.match(flat, /■ 次回予定\s+なし/, '中身が空の見出しは「なし」');
  assert.ok(!out.unknown.includes('議事録本文'));
});
