'use strict';
/*
 * テスト用の雛形（.docx / .xlsx / .pptx）を作る。
 * わざと {{ }} を複数の書式の塊に分けたり、数値表記の項目名・共有文字列・直接入力文字列を入れてある。
 *
 *   node scripts/make-test-templates.js [出力フォルダ]     … ファイルを書き出す（既定：./out）
 *   require('./make-test-templates').build()               … テストから Buffer を受け取る
 */
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

const CT = (extra) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' + extra + '</Types>';
const RELS = (target, type) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/' + type + '" Target="' + target + '"/></Relationships>';
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

async function docx() {
  const z = new JSZip();
  z.file('[Content_Types].xml', CT('<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'));
  z.file('_rels/.rels', RELS('word/document.xml', 'officeDocument'));
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  z.file('word/document.xml', XML + '<w:document ' + W + '><w:body>' +
    // 1) 1つの塊
    '<w:p><w:r><w:t>案件名：{{案件名}}</w:t></w:r></w:p>' +
    // 2) {{ }} が3つの塊に分かれている（書式つき）
    '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>現場：{{現</w:t></w:r><w:r><w:t>場住</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>所}}</w:t></w:r></w:p>' +
    // 3) 項目名が数値表記（案 = &#26696;）
    '<w:p><w:r><w:t>{{&#26696;件名}}</w:t></w:r></w:p>' +
    // 4) 複数行・「<」「&」を含む値用
    '<w:p><w:r><w:t>備考：{{備考}}</w:t></w:r></w:p>' +
    // 5) 見つからない名前
    '<w:p><w:r><w:t>{{存在しない項目}}</w:t></w:r></w:p>' +
    // 6) 段落をまたぐ {{ ... }}（置き換えてはいけない・壊してはいけない）
    '<w:p><w:r><w:t>{{案件</w:t></w:r></w:p><w:p><w:r><w:t>名}}</w:t></w:r></w:p>' +
    // 7) 版・作成日
    '<w:p><w:r><w:t>{{書類名}} {{版}} {{作成日}}</w:t></w:r></w:p>' +
    '</w:body></w:document>');
  z.file('word/header1.xml', XML + '<w:hdr ' + W + '><w:p><w:r><w:t>ヘッダー：{{案件名}}</w:t></w:r></w:p></w:hdr>');
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function xlsx() {
  const z = new JSZip();
  z.file('[Content_Types].xml', CT('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'));
  z.file('_rels/.rels', RELS('xl/workbook.xml', 'officeDocument'));
  z.file('xl/workbook.xml', XML + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>');
  z.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>');
  // 共有文字列：0=普通 1=書式分割 2=ふりがな付き 3=ほかのセル（触らない）
  z.file('xl/sharedStrings.xml', XML + '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">' +
    '<si><t>{{案件名}}</t></si>' +
    '<si><r><rPr><b/></rPr><t>{{現場</t></r><r><t>住所}}</t></r></si>' +
    '<si><t>{{発注者}}</t><rPh sb="0" eb="3"><t>ハッチュウシャ</t></rPh></si>' +
    '<si><t>固定の文字</t><rPh sb="0" eb="2"><t>コテイ</t></rPh></si>' +
    '</sst>');
  z.file('xl/worksheets/sheet1.xml', XML + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
    '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
    '<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" t="s"><v>3</v></c></row>' +
    '<row r="3"><c r="A3" t="inlineStr"><is><t>直接入力：{{備考}}</t></is></c></row>' +
    '<row r="4"><c r="A4" t="inlineStr"><is><t>{{存在しない項目}}</t></is></c></row>' +
    '</sheetData></worksheet>');
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function pptx() {
  const z = new JSZip();
  z.file('[Content_Types].xml', CT('<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'));
  z.file('_rels/.rels', RELS('ppt/presentation.xml', 'officeDocument'));
  z.file('ppt/presentation.xml', XML + '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>');
  const NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
  z.file('ppt/slides/slide1.xml', XML + '<p:sld ' + NS + '><p:cSld><p:spTree>' +
    '<p:sp><p:txBody><a:p><a:r><a:rPr lang="ja-JP" sz="3200" b="1"/><a:t>案件名：{{案件名}}</a:t></a:r></a:p></p:txBody></p:sp>' +
    '<p:sp><p:txBody><a:p><a:r><a:rPr sz="1800"/><a:t>{{</a:t></a:r><a:r><a:rPr sz="1800"/><a:t>備考</a:t></a:r><a:r><a:rPr sz="1800"/><a:t>}}</a:t></a:r></a:p></p:txBody></p:sp>' +
    '</p:spTree></p:cSld></p:sld>');
  return z.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
}

async function build() { return { docx: await docx(), xlsx: await xlsx(), pptx: await pptx() }; }

module.exports = { build };

if (require.main === module) {
  (async () => {
    const dir = path.resolve(process.argv[2] || 'out');
    fs.mkdirSync(dir, { recursive: true });
    const b = await build();
    for (const ext of Object.keys(b)) {
      const f = path.join(dir, '【テスト】雛形.' + ext);
      fs.writeFileSync(f, b[ext]);
      console.log('作成：' + f);
    }
  })();
}
