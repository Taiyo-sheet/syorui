/*
 * 案件くん：「書類を作る」ボタン（設定なしで動く版）
 *
 * 案件くんの詳細画面の上部に「書類を作る」ボタンが出ます。押すと、
 *   ・書類マスタの書類が工程別に並ぶ（工程・立場の選択肢も書類マスタから自動で読み取る）
 *   ・「作成して保存」で、雛形の {{案件くんの項目名}} が案件の値に置き換わり、
 *     案件くんのこの案件の「書類保管」表に、版つきで保存される（別のアプリは使いません）
 *
 * 設定は要りません。
 *   ・書類マスタは名前で自動で見つけます。保管用の「書類保管」表は、案件くんに scripts/add-doc-storage.js で足します
 *   ・雛形の {{ }} には、案件くんの「項目名」をそのまま書けば入ります
 *     例）{{案件名}} {{現場住所}} {{発注者}}   ※項目コードでも可
 *   ・このほか {{作成日}} {{版}} {{書類名}} も使えます
 *   ・案件くんにない名前の {{ }} は、置き換えずに残します（保存後のメッセージで知らせます）
 *
 * 登録：案件くんの「JavaScript / CSSでカスタマイズ」に、PC用・モバイル用の両方へ
 *   1) https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js
 *   2) このファイル
 *  の順で追加する（scripts/apply-customize.js が自動でやります）
 */
(function () {
  'use strict';

  var APP_NAMES = { master: '書類マスタ' };
  // 書類マスタの項目コード（scripts/create-apps.js で作ると、この名前になる）
  var MF = { stage: '工程', order: '表示順', name: '書類名', desc: '説明', template: '雛形', audience: '見られる立場', target: '対象', status: '使用状況' };
  // 案件くんの「書類保管」表（scripts/add-doc-storage.js で足す）の項目コード
  var SUB = { table: '書類保管', name: '書類保管_書類名', stage: '書類保管_工程', version: '書類保管_版', file: '書類保管_ファイル', at: '書類保管_作成日',
    // 承認とPDF（scripts/add-approval-fields.js で、書類保管の表に足す列。足すまでは、承認・PDF化のボタンは出ない）
    status: '書類保管_承認状況', by: '書類保管_承認者', approvedAt: '書類保管_承認日時', pdf: '書類保管_PDF' };
  var APPROVED = '承認済み';
  // PDF化サーバーのURL（Render に置いた pdf-server/。空にすると、PDF化のボタンは押せなくなる）
  var PDF_SERVER_URL = 'https://ts-pdf-server.onrender.com';
  var ACTIVE = '使用中';
  var NAME_LABELS = ['工事名称', '案件名', '現場名', '工事名'];   // ファイル名に使う項目（先に見つかったもの）
  // 案件くんの項目のうち、「橋梁」「一般道路」など、書類マスタの「対象」と同じ言葉が入っているもの。
  // その項目名を入れると、案件を開いたときに、対象が自動で選ばれる（空のままなら、画面で手で選ぶ）
  var TARGET_LABEL = '';
  var STAGE_LABEL = '現在の工程';                      // 案件くんにこの項目があれば、その工程を最初に表示

  // ---------- 値の整形 ----------
  function pad(n) { return n < 10 ? '0' + n : String(n); }
  function fieldText(f, prop) {
    if (!f || f.type === 'SUBTABLE' || f.type === 'FILE') return '';
    var v = f.value;
    if (v === null || v === undefined) return '';
    if (Array.isArray(v)) return v.map(function (x) { return (x && typeof x === 'object') ? (x.name || x.code || '') : x; }).join('、');
    if (typeof v === 'object') return v.name || v.code || '';
    v = String(v);
    // 桁区切りは、案件くんの項目に「桁区切りを表示する」が付いているときだけ（付いていない案件番号などは、そのまま）
    if (f.type === 'NUMBER' && prop && prop.digit && v !== '' && !isNaN(v)) return Number(v).toLocaleString('ja-JP', { maximumFractionDigits: 10 });
    var m = v.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日';
    if (f.type === 'DATETIME' && /^\d{4}-\d{2}-\d{2}T/.test(v)) {
      var d = new Date(v);
      if (!isNaN(d)) return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日 ' + d.getHours() + ':' + pad(d.getMinutes());
    }
    return v;
  }
  function today() { var d = new Date(); return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日'; }
  function escXml(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  // XMLで使えない制御文字（縦タブなど）が値に混ざっていると、書類が壊れて開けなくなるので取り除く
  function cleanValue(v) { return String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, ''); }
  function toXmlText(v, kind, rpr) {
    var lines = cleanValue(v).split(/\r?\n/).map(escXml);
    if (lines.length === 1) return lines[0];
    if (kind === 'docx') return lines.join('</w:t><w:br/><w:t xml:space="preserve">');
    // PowerPoint：2行目以降も、置き換え元と同じ書式（文字の大きさ・色など）で続ける
    if (kind === 'pptx') return lines.join('</a:t></a:r><a:br/><a:r>' + (rpr || '') + '<a:t>');
    return lines.join('\n');
  }
  // 議事録本文を、きちんとした議事録の形（見出しの帯、箇条書き、ToDoの表）の「段落のかたまり」にする（Word用）
  //   ToDo の行が「期限｜担当｜内容」の形（｜で3つに区切る）なら、表にする。そうでなければ、箇条書きのまま
  function parseMinutes(text) {
    var sections = [], cur = null;
    String(text || '').replace(/\r/g, '').split('\n').forEach(function (raw) {
      var line = raw.replace(/\s+$/, '');
      var h = line.trim().match(/^【([^】]+)】\s*(.*)$/);
      if (h) { cur = { head: h[1], lines: [] }; sections.push(cur); if (h[2]) cur.lines.push(h[2]); return; }
      if (!line.trim()) return;
      if (!cur) { cur = { head: '', lines: [] }; sections.push(cur); }
      cur.lines.push(line.trim());
    });
    return sections;
  }
  function wp(inner, ppr) { return '<w:p>' + (ppr ? '<w:pPr>' + ppr + '</w:pPr>' : '') + inner + '</w:p>'; }
  function wr(text, rpr) { return '<w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + escXml(text) + '</w:t></w:r>'; }
  function wcell(text, w, bold, fill) {
    return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' + (fill ? '<w:shd w:val="clear" w:color="auto" w:fill="' + fill + '"/>' : '') + '</w:tcPr>' +
      wp(wr(text, bold ? '<w:b/><w:sz w:val="20"/>' : '<w:sz w:val="20"/>'), '<w:spacing w:before="40" w:after="40"/>') + '</w:tc>';
  }
  function todoTable(rows) {
    var W = [700, 5338, 1800, 1800], B = '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(function (e) { return '<w:' + e + ' w:val="single" w:sz="4" w:space="0" w:color="888888"/>'; }).join('') + '</w:tblBorders>';
    var x = '<w:tbl><w:tblPr><w:tblW w:w="9638" w:type="dxa"/>' + B + '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>' + W.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join('') + '</w:tblGrid>';
    x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + wcell('No', W[0], true, 'E8EEF7') + wcell('内容', W[1], true, 'E8EEF7') + wcell('担当', W[2], true, 'E8EEF7') + wcell('期限', W[3], true, 'E8EEF7') + '</w:tr>';
    rows.forEach(function (r, i) { x += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' + wcell(String(i + 1), W[0]) + wcell(r[2], W[1]) + wcell(r[1], W[2]) + wcell(r[0], W[3]) + '</w:tr>'; });
    return x + '</w:tbl>' + wp('', '<w:spacing w:after="60"/>');
  }
  function minutesParagraphs(text) {
    var out = '';
    parseMinutes(cleanValue(text)).forEach(function (sec) {
      if (sec.head) out += wp(wr('■ ' + sec.head, '<w:b/><w:color w:val="FFFFFF"/><w:sz w:val="22"/>'), '<w:keepNext/><w:shd w:val="clear" w:color="auto" w:fill="1F3864"/><w:spacing w:before="200" w:after="80"/><w:ind w:left="80"/>');
      var rows = [];
      var todo = /^ToDo/i.test(sec.head);
      sec.lines.forEach(function (l) {
        var body = l.replace(/^[・\-*●•‣]\s*/, '');
        var parts = body.split(/[｜|]/).map(function (x) { return x.trim(); });
        if (todo && parts.length === 3 && parts[2]) { rows.push(parts); return; }
        if (rows.length) { out += todoTable(rows); rows = []; }
        out += wp(wr(l, '<w:sz w:val="21"/>'), '<w:spacing w:after="40"/><w:ind w:left="300" w:hanging="220"/>');
      });
      if (rows.length) out += todoTable(rows);
      if (!sec.lines.length) out += wp(wr('なし', '<w:sz w:val="21"/>'), '<w:ind w:left="300"/>');
    });
    return out;
  }
  // 段落が、{{議事録本文}} だけなら、その段落ごと、上の形に入れ替える
  function expandMinutes(xml, text) {
    if (!String(text || '').trim()) return xml;
    return xml.replace(/<w:p\b(?:(?!<\/w:p>)[\s\S])*<\/w:p>/g, function (p) {
      return decodeLabel(p.replace(/<[^>]*>/g, '')) === '{{議事録本文}}' ? minutesParagraphs(text) : p;
    });
  }

  // 議事録本文：【見出し】だけの行を太字にし、ほかは、改行でつなぐ（Word用）
  function toXmlMinutes(v) {
    var lines = cleanValue(v).split(/\r?\n/);
    var parts = lines.map(function (l) {
      var t = escXml(l);
      return /^【[^】]+】$/.test(l.trim()) ? '</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">' + t + '</w:t></w:r><w:r><w:t xml:space="preserve">' : t;
    });
    return parts.join('</w:t><w:br/><w:t xml:space="preserve">');
  }
  function decodeLabel(s) {
    return s.replace(/&#x([0-9a-f]+);/gi, function (m, h) { return String.fromCodePoint(parseInt(h, 16)); })
      .replace(/&#(\d+);/g, function (m, d) { return String.fromCodePoint(parseInt(d, 10)); })
      .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')
      .replace(/ |　/g, ' ').trim();
  }
  // {{ }} が複数の書式の塊（XMLの要素）に分かれていても見つける。
  // ただし段落・セルをまたぐタグ（</w:p> など）は飛び越えない（飛び越えると書類が壊れるため）
  var TAG = '<(?!\\/?(?:w:p|a:p|si|is|c|row|w:tc|w:tr)[\\s>\\/])[^>]*>';
  var PH = new RegExp('\\{(?:' + TAG + ')*\\{((?:' + TAG + '|[^{}<])*)\\}(?:' + TAG + ')*\\}', 'g');
  function lastRunProps(xml, offset) {
    var s = xml.lastIndexOf('<a:r>', offset);
    if (s < 0) return '';
    var m = xml.slice(s, offset).match(/<a:rPr[^>]*\/>|<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/);
    return m ? m[0] : '';
  }
  // 戻り値：置き換えた XML。unknown に「値が見つからなかった {{名前}}」を集める
  function replacePlaceholders(xml, values, kind, unknown, empty) {
    return xml.replace(PH, function (whole, inner, offset, str) {
      var label = decodeLabel(inner.replace(/<[^>]*>/g, ''));
      if (Object.prototype.hasOwnProperty.call(values, label)) {
        if (empty && String(values[label]).trim() === '' && empty.indexOf(label) < 0) empty.push(label);
        if (label === '議事録本文' && kind === 'docx') return toXmlMinutes(values[label]);
        return toXmlText(values[label], kind, kind === 'pptx' ? lastRunProps(str, offset) : '');
      }
      if (unknown && unknown.indexOf(label) < 0) unknown.push(label);
      return whole;
    });
  }

  // 差し込み対象のXML（本文・ヘッダー・フッター・スライド・ノート・図形の文字など）
  var TARGET = /^(word\/(document|header\d*|footer\d*|footnotes|endnotes)\.xml|xl\/sharedStrings\.xml|xl\/worksheets\/[^\/]+\.xml|xl\/drawings\/[^\/]+\.xml|ppt\/(slides|notesSlides|slideLayouts|slideMasters)\/[^\/]+\.xml)$/;
  var OFFICE = { docx: 1, xlsx: 1, pptx: 1 };
  // Excelのふりがな（<rPh>）が {{ }} の中に混ざると名前が合わなくなるので、{{ }} を含むセルのふりがなだけ取り除く
  function dropPhonetic(xml) {
    return xml.replace(/<si>[\s\S]*?<\/si>/g, function (si) {
      return /\{[^]*\{/.test(si.replace(/<[^>]*>/g, '')) ? si.replace(/<rPh[\s\S]*?<\/rPh>/g, '') : si;
    });
  }
  function fillTemplate(JSZipLib, buf, ext, values) {
    var unknown = [], empty = [];
    if (!OFFICE[ext]) return Promise.resolve({ data: buf, unknown: unknown, empty: empty });
    return JSZipLib.loadAsync(buf).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return TARGET.test(n); });
      return Promise.all(names.map(function (n) {
        return zip.file(n).async('string').then(function (x) {
          if (n === 'xl/sharedStrings.xml') x = dropPhonetic(x);
          if (ext === 'docx' && n === 'word/document.xml' && values['議事録本文']) x = expandMinutes(x, values['議事録本文']);   // 議事録本文は、きちんとした形に
          var y = replacePlaceholders(x, values, ext, unknown, empty);
          // 名前が違うものに加えて、{{ }} が段落やセルをまたいでいて置き換えられなかったものも拾う
          (y.replace(/<[^>]*>/g, '').match(/\{\{[^{}]*\}\}/g) || []).forEach(function (t) {
            var l = decodeLabel(t.slice(2, -2));
            if (unknown.indexOf(l) < 0) unknown.push(l);
          });
          zip.file(n, y);
        });
      })).then(function () { return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }); })
        .then(function (data) { return { data: data, unknown: unknown, empty: empty }; });
    });
  }
  // 雛形の中の {{ }} の名前を、全部集める（なければ空の配列）
  function listPlaceholders(JSZipLib, buf, ext) {
    if (!OFFICE[ext]) return Promise.resolve([]);
    return JSZipLib.loadAsync(buf).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return TARGET.test(n); });
      return Promise.all(names.map(function (n) {
        return zip.file(n).async('string').then(function (x) {
          if (n === 'xl/sharedStrings.xml') x = dropPhonetic(x);
          return (x.replace(/<[^>]*>/g, '').match(/\{\{[^{}]+\}\}/g) || []).map(function (t) { return decodeLabel(t.slice(2, -2)); });
        });
      })).then(function (r) { return [].concat.apply([], r); });
    });
  }
  // 雛形に {{ }} が1つでもあるか。なければ「案件の情報は入れず、雛形のまま使う書類」として扱う
  function hasPlaceholders(JSZipLib, buf, ext) {
    return listPlaceholders(JSZipLib, buf, ext).then(function (l) { return l.length > 0; });
  }
  // 議事録など、作成のときに入力してもらう項目（雛形に {{打合せ日}} {{議事録メモ}} などと書くと、その入力欄が出る）
  //   blank：入力されなかったときに入れる言葉（なければ、空のまま入れて、作成後に知らせる）
  //   init ：入力欄の初めの値（関数なら、開いたときに決める）
  function loginName() { try { return kintone.getLoginUser().name || ''; } catch (e) { return ''; } }
  var ASK = [
    { key: '議事録本文', label: '③ まとまった議事録（この内容が、書類に入ります。ここで、直せます）', type: 'doc' },
    { key: '打合せ日', label: '打合せ日（初期値は今日）', type: 'date', blank: function () { return today(); } },
    { key: '場所', label: '場所', type: 'text' },
    { key: '出席者', label: '出席者（分からなければ空のまま→「不明」）', type: 'text', blank: '不明' },
    { key: '作成者', label: '作成者（初期値は、開いた人）', type: 'text', init: function () { return loginName() || 'くに'; }, blank: function () { return loginName() || 'くに'; } },
    { key: '議題', label: '議題（箇条書き。1行に1項目）', type: 'memo', blank: 'なし' },
    { key: '議事録メモ', label: '打合せ内容（箇条書き。1行に1項目）', type: 'memo', blank: 'なし' },
    { key: '決定事項', label: '決定事項（箇条書き。1行に1項目）', type: 'memo', blank: 'なし' },
    { key: 'ToDo', label: 'ToDo（いつまでに・誰が・何を。1行に1つ）', type: 'memo', blank: 'なし' },
    { key: '次回予定', label: '次回予定（あれば。「次回は〇月〇日」など）', type: 'text', blank: 'なし' }
  ];
  function askDefault(a) { return typeof a.blank === 'function' ? a.blank() : (a.blank || ''); }
  // 音声の文字起こしを、議事録の項目に分ける（簡易版。外部には送りません。精度は高くないので、必ず、確認・修正してから作成する）
  //   戻り値：{ 出席者, 場所, 議題, 議事録メモ, 決定事項, ToDo, 次回予定 }（見つからなかった項目は、空）
  // 見出し（【決定事項】■ToDo など）が付いた文章（Copilot や Teams の「まとめ」を、見出し付きで出したもの）は、見出しごとに、そのまま振り分ける
  var HEAD = [
    ['出席者', /^(出席者|参加者|出席)$/], ['場所', /^(場所|会場|開催場所)$/], ['議題', /^(議題|アジェンダ|テーマ|目的)$/],
    ['議事録メモ', /^(打合せ内容|打ち合わせ内容|内容|議事|議事録|要点|概要|討議内容|協議内容|ディスカッション|主な内容)$/],
    ['決定事項', /^(決定事項|決定|合意事項|合意)$/], ['ToDo', /^(todo|to-do|やること|宿題|宿題事項|アクション|アクションアイテム|アクション項目|対応事項|今後の対応)$/i],
    ['次回予定', /^(次回予定|次回|次回打合せ|次回の打合せ|次のステップ)$/]
  ];
  function headOf(line) {
    var t = line.replace(/^[\s#＃■□◆◇●○▼▽【\[（(「*・\-]+/, '').replace(/[\s】\]）)」:：]+$/, '').replace(/^[0-9０-９]+[.．、)）]\s*/, '').trim();
    for (var i = 0; i < HEAD.length; i++) { if (HEAD[i][1].test(t)) return HEAD[i][0]; }
    return '';
  }
  function summarizeByHeadings(text) {
    var lines = String(text || '').replace(/\r/g, '').split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
    var onlyHeads = lines.filter(function (l) { return headOf(l); }).length;
    if (onlyHeads < 2) return null;   // 見出しだけの行が、2つ以上あるときだけ、見出しで振り分ける
    var out = { '出席者': [], '場所': [], '議題': [], '議事録メモ': [], '決定事項': [], 'ToDo': [], '次回予定': [] }, cur = '';
    lines.forEach(function (line) {
      var hk = headOf(line);
      if (hk) { cur = hk; return; }
      var inl = line.match(/^[\s■□◆◇●○▼▽【\[#＃*・\-]*([^:：】\]\s]{1,10})[】\]]?\s*[:：]?\s*(.+)$/);   // 「【次回予定】次回は…」「決定事項：…」のように、同じ行に中身がある
      if (inl && headOf(inl[1])) { cur = headOf(inl[1]); out[cur].push(inl[2].trim()); return; }
      if (cur) out[cur].push(line.replace(/^[\s・\-*●•‣]+/, '').trim());
    });
    var res = {};
    Object.keys(out).forEach(function (k) {
      var list = out[k].filter(Boolean);
      res[k] = (k === '出席者' || k === '場所' || k === '次回予定') ? list.join('、') : list.map(function (l) { return '・' + l; }).join('\n');
    });
    return res;
  }
  function summarizeLocal(text) {
    var byHead = summarizeByHeadings(text); if (byHead) return byHead;
    var out = { '出席者': '', '場所': '', '議題': [], '議事録メモ': [], '決定事項': [], 'ToDo': [], '次回予定': '' };
    var FILLER = /^(えー+|あー+|えっと|あの+|その+|まあ|はい|ええ|うん|そうですね)[、。,\s]*/;
    String(text || '').replace(/\r/g, '').split(/\n|(?<=[。！？!?])/).forEach(function (raw) {
      var l = raw.replace(/^[\s・\-*●•‣]+/, '').replace(FILLER, '').trim();
      if (l.length < 2) return;
      var m;
      if ((m = l.match(/^(?:出席者|参加者)[:：]\s*(.+)$/))) { out['出席者'] = m[1].trim(); return; }
      if ((m = l.match(/^(?:場所|会場)[:：]\s*(.+)$/))) { out['場所'] = m[1].trim(); return; }
      if (/次回/.test(l) && !out['次回予定']) { out['次回予定'] = l; return; }
      if (/(決定|決まり|決めまし|確定|で行く|で進め|で合意|ということで|とします|することにな)/.test(l)) { out['決定事項'].push(l); return; }
      if (/(まで(に)?|お願い|対応|送付|提出|手配|連絡|確認し|調整|共有)/.test(l) && /(\d{1,2}[\/月]\d{1,2}|来週|今週|明日|明後日|今月|来月|月末|至急|早め)/.test(l)) { out['ToDo'].push(l); return; }
      if (/(について|の件|議題|テーマ)/.test(l) && l.length <= 40) { out['議題'].push(l); return; }
      out['議事録メモ'].push(l);
    });
    ['議題', '議事録メモ', '決定事項', 'ToDo'].forEach(function (k) { out[k] = out[k].join('\n'); });
    return out;
  }

  // 外部のAIに送る前に、個人情報などを、符号（【人物1】【電話1】…）に置き換える。AIの返事が戻ったら、元の言葉に、戻す。
  // 置き換えるもの：電話番号・メールアドレス・郵便番号・金額（〇円）・「〇〇さん／様／氏／君／くん」の名前・利用者が、指定した言葉
  // ※ 敬称のない名前や、文脈で分かる情報は、置き換えられません（完全ではありません）。送る内容は、送る前に、画面で確認できます
  function maskForAI(text, extraWords) {
    var map = {}, counts = {}, src = String(text || '');
    var put = function (kind, orig) {
      for (var k in map) { if (map[k] === orig) return k; }
      counts[kind] = (counts[kind] || 0) + 1; var tok = '【' + kind + counts[kind] + '】'; map[tok] = orig; return tok;
    };
    (extraWords || []).filter(Boolean).forEach(function (w) { src = src.split(w).join(put('伏せ', w)); });
    src = src.replace(/0\d{1,4}[-ー－]\d{1,4}[-ー－]\d{3,4}/g, function (m) { return put('電話', m); });
    src = src.replace(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g, function (m) { return put('メール', m); });
    src = src.replace(/〒?\s?\d{3}[-ー－]\d{4}/g, function (m) { return put('郵便番号', m); });
    src = src.replace(/[¥￥]\s?[\d,，]+|[\d,，.]+\s?(?:万円|千円|円)/g, function (m) { return put('金額', m); });
    src = src.replace(/([一-龥々ァ-ヶー]{2,5})(さん|様|氏|君|くん|殿)/g, function (m, name, h) {
      if (/^(お客|担当|先方|当方|各位|関係|責任|営業|現場|職長|元請|発注|下請|施主|所長|課長|部長)/.test(name)) return m;
      return put('人物', name) + h;
    });
    return {
      masked: src,
      unmask: function (t) { return String(t || '').replace(/【(?:人物|電話|メール|郵便番号|金額|伏せ)\d+】/g, function (tok) { return map[tok] != null ? map[tok] : tok; }); },
      count: Object.keys(map).length
    };
  }

  // 項目に分けた結果を、見出し付きの、ひとかたまりの文章にする
  var MINUTES_ORDER = [['出席者', '出席者', '不明'], ['場所', '場所', ''], ['議題', '議題', 'なし'], ['議事録メモ', '打合せ内容', 'なし'], ['決定事項', '決定事項', 'なし'], ['ToDo', 'ToDo', 'なし'], ['次回予定', '次回予定', 'なし']];
  function formatMinutes(res) {
    return MINUTES_ORDER.map(function (o) {
      var v = String((res && res[o[0]]) || '').trim() || o[2];
      return '【' + o[1] + '】\n' + v;
    }).join('\n\n');
  }
  // 箇条書きの記号（- * ● • ・）をそろえる。空行は取り除く
  function tidyMemo(t) {
    return String(t || '').split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean)
      .map(function (l) { return l.replace(/^[-*●•・‣]\s*/, '・'); }).join('\n');
  }
  // 書類の文字を取り出す（docx・xlsx・pptx）。段落・セルの区切りは、改行にする
  function extractText(JSZipLib, buf, ext) {
    if (!OFFICE[ext]) return Promise.resolve('');
    return JSZipLib.loadAsync(buf).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return TARGET.test(n); });
      return Promise.all(names.map(function (n) {
        return zip.file(n).async('string').then(function (x) {
          return decodeLabel(x.replace(/<\/(w:p|a:p|si|row|w:tr)>/g, '\n').replace(/<[^>]*>/g, ''));
        });
      })).then(function (r) { return r.join('\n'); });
    });
  }
  // 書類のチェック（外部には送らず、この画面の中だけで行う）。指摘のリストを返す（なければ空）
  //   values：案件くんの値（項目名 → 値）。書類の「工事名称：〇〇」が、案件くんの値と違えば、指摘する
  // 案件くんの値と、書類の値を、突き合わせる項目。書類に「項目名：値」の行があり、値が違えば、指摘する（案件くんの全項目が対象）
  var NOT_CHECKED = { '作成日': 1, '版': 1, '書類名': 1 };
  function checkText(text, values) {
    var out = [], t = String(text || '');
    var left = (t.match(/\{\{[^{}]+\}\}/g) || []).filter(function (x, i, a) { return a.indexOf(x) === i; });
    if (left.length) out.push('置き換えられていない項目が残っています：' + left.join(' '));
    // 「■ 工事場所：」のあとが、空のまま
    var blanks = [];
    t.split('\n').forEach(function (l) {
      var m = l.match(/^[\s■●◆□・]*([^\s:：（）()]{1,14})[:：]\s*$/);
      if (m && blanks.indexOf(m[1]) < 0) blanks.push(m[1]);
    });
    if (blanks.length) out.push('値が空の欄があります：' + blanks.join('、'));
    // 案件くんの値と、書類の値が違う（案件くんの項目名ごとに、書類の「項目名：値」の行と突き合わせる）
    var norm = function (x) { return String(x).replace(/\s+/g, ''); };
    Object.keys(values || {}).forEach(function (lb) {
      var want = values[lb] == null ? '' : String(values[lb]).trim();
      if (NOT_CHECKED[lb] || !want || /\n/.test(want) || want.length > 80 || lb.length > 30) return;
      var m = t.match(new RegExp('(?:^|\\n)[\\s■●◆□・]*' + lb.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '[:：][ \\t　]*([^\\n]+)'));
      if (m && m[1].trim() && norm(m[1]) !== norm(want)) out.push('「' + lb + '」が、案件くんの値と違います（書類：' + m[1].trim() + '／案件くん：' + want + '）');
    });
    if (/（例[:：]|\(例[:：]/.test(t)) out.push('「（例：…）」の見本の文が、そのまま残っています');
    if (/(^|\n)[^\n]*[:：]?\s*(　|\s)+年(　|\s)+月(　|\s)+日/.test(t)) out.push('日付が空欄のようです（「　年　月　日」）');
    return out;
  }

  // 案件名から、橋梁か一般道路かを推定する。「〇〇橋」「高架橋」「跨線橋」「橋梁」が入っていれば橋梁。
  // 「橋本」「橋口」のような地名の「橋」は、橋として数えない
  var BRIDGE = /高架|跨線|橋梁|[^\s\u3000橋]橋(?![本口田爪立場詰間])/;
  function guessTarget(name) { return BRIDGE.test(String(name || '')) ? '橋梁' : '一般道路'; }
  function safeName(s) { return String(s).replace(/[\u0000-\u001F\\\/:*?"<>|]/g, '_').slice(0, 120); }
  function buildFilename(projectName, docName, version, ext) {
    return safeName(projectName + '_' + docName + '_v' + version) + (ext ? '.' + ext : '');
  }
  function quote(s) { return '"' + String(s).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"'; }

  function buildValues(record, lm) {
    var values = {};
    Object.keys(lm.map).forEach(function (label) {
      var code = lm.map[label], f = record[code];
      if (f && f.type !== 'SUBTABLE' && f.type !== 'FILE') values[label] = fieldText(f, lm.props[code]);
    });
    values['作成日'] = today();
    return values;
  }


  // 書類マスタ（master-template.js）が、「試しに作る」で、この部品だけを借りるときの出口（この画面では、ボタンなどは、一切出さない）
  if (typeof kintone !== 'undefined' && typeof window !== 'undefined' && window.__TS_DOC_ENGINE_ONLY) {
    window.TSDocEngine = { fillTemplate: fillTemplate, buildValues: buildValues, checkText: checkText, extractText: extractText, fieldText: fieldText, ASK: ASK, NAME_LABELS: NAME_LABELS };
    return;
  }
  // Node（テスト）から使うときは、ここまでの部品だけを渡して終わる
  if (typeof kintone === 'undefined') {
    if (typeof module !== 'undefined' && module.exports) {
      module.exports = { formatMinutes: formatMinutes, maskForAI: maskForAI, summarizeLocal: summarizeLocal, checkText: checkText, extractText: extractText, tidyMemo: tidyMemo, fieldText: fieldText, replacePlaceholders: replacePlaceholders, fillTemplate: fillTemplate, hasPlaceholders: hasPlaceholders, guessTarget: guessTarget, buildFilename: buildFilename, buildValues: buildValues, quote: quote, SUB: SUB, MF: MF, today: today };
    }
    return;
  }

  var isMobile = /^\/k\/(guest\/\d+\/)?m\//.test(location.pathname);
  var recApi = isMobile ? kintone.mobile.app.record : kintone.app.record;
  var appApi = isMobile ? kintone.mobile.app : kintone.app;
  var STORE_KEY = 'doc_hub_audience';
  var appIds = {};

  function api(method, path, params) { return kintone.api(kintone.api.url(path, true), method, params); }

  function findAppId(name) {
    if (appIds[name]) return Promise.resolve(appIds[name]);
    return api('GET', '/k/v1/apps', { name: name, limit: 100 }).then(function (r) {
      var m = r.apps.filter(function (a) { return a.name === name; });
      if (!m.length) throw new Error('「' + name + '」アプリが見つかりません。アプリ名が合っているか、閲覧権限があるか確認してください');
      if (m.length > 1) throw new Error('「' + name + '」という名前のアプリが' + m.length + '個あります。どれを使うか決められないため止めました。名前を変えるか、不要なアプリを整理してください');
      appIds[name] = m[0].appId;
      return appIds[name];
    });
  }

  // 案件くんの「項目名」→ 項目コード の対応と、項目の設定（桁区切りなど）を読み取る。
  // 画面の getFormFields → 項目取得API → 項目コードのみ、の順に試す
  function loadLabelMap(record) {
    var props = {}, map = {};
    Object.keys(record).forEach(function (c) { map[c] = c; });
    function apply(p) {
      Object.keys(p || {}).forEach(function (c) {
        var pr = p[c]; if (!pr) return;
        props[pr.code || c] = pr;
        if (pr.label) map[String(pr.label).trim()] = pr.code || c;
      });
    }
    var first;
    try { first = Promise.resolve(kintone.app.getFormFields && kintone.app.getFormFields()); } catch (e) { first = Promise.resolve(null); }
    return first.catch(function () { return null; }).then(function (p) {
      if (p && Object.keys(p).length) { apply(p); return; }
      return api('GET', '/k/v1/app/form/fields', { app: appApi.getId() }).then(function (r) { apply(r.properties); }, function () {});
    }).then(function () { return { map: map, props: props }; });
  }
  // ---------- ファイル ----------
  function download(fileKey) {
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open('GET', kintone.api.url('/k/v1/file', true) + '?fileKey=' + encodeURIComponent(fileKey));
      xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
      xhr.responseType = 'arraybuffer';
      xhr.onload = function () { xhr.status === 200 ? resolve(xhr.response) : reject(new Error('雛形をダウンロードできませんでした（' + xhr.status + '）')); };
      xhr.onerror = function () { reject(new Error('雛形をダウンロードできませんでした')); };
      xhr.send();
    });
  }
  function upload(data, filename) {
    return new Promise(function (resolve, reject) {
      var fd = new FormData();
      fd.append('__REQUEST_TOKEN__', kintone.getRequestToken());
      fd.append('file', new Blob([data]), filename);
      var xhr = new XMLHttpRequest();
      xhr.open('POST', kintone.api.url('/k/v1/file', true));
      xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
      xhr.onload = function () { xhr.status === 200 ? resolve(JSON.parse(xhr.responseText).fileKey) : reject(new Error('書類をアップロードできませんでした（' + xhr.status + '）')); };
      xhr.onerror = function () { reject(new Error('書類をアップロードできませんでした')); };
      xhr.send(fd);
    });
  }

  // 雛形をそのまま保存（ダウンロード）する。案件には保存しない
  function saveAs(data, filename) {
    var url = URL.createObjectURL(new Blob([data]));
    var a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
  }
  var kinds = {};   // 雛形のファイルキー → 'merge'（{{ }}あり）／'plain'（雛形のまま）
  var phs = {};     // 雛形のファイルキー → 雛形の中の {{ }} の名前
  function templateOf(rec) { var f = rec[MF.template] && rec[MF.template].value; return f && f[0]; }
  function extOf(name) { return ((String(name).match(/\.([^.]+)$/) || [])[1] || '').toLowerCase(); }
  var pending = {};
  function inspect(rec) {
    var tpl = templateOf(rec);
    if (!tpl) return Promise.resolve('merge');
    if (kinds[tpl.fileKey]) return Promise.resolve(kinds[tpl.fileKey]);
    if (pending[tpl.fileKey]) return pending[tpl.fileKey];
    pending[tpl.fileKey] = download(tpl.fileKey).then(function (buf) { return listPlaceholders(JSZip, buf, extOf(tpl.name)); })
      .then(function (list) { phs[tpl.fileKey] = list; kinds[tpl.fileKey] = list.length ? 'merge' : 'plain'; delete pending[tpl.fileKey]; return kinds[tpl.fileKey]; },
        function () { delete pending[tpl.fileKey]; return 'merge'; });   // 調べられないときは、従来どおり「作成して保存」にする
    return pending[tpl.fileKey];
  }
  function asksOf(rec) { var t = templateOf(rec); var l = (t && phs[t.fileKey]) || []; return ASK.filter(function (a) { return l.indexOf(a.key) >= 0; }); }
  function kindOf(rec) { var t = templateOf(rec); return t ? kinds[t.fileKey] : undefined; }
  function openAsIs(rec) {
    var tpl = templateOf(rec);
    return download(tpl.fileKey).then(function (buf) { saveAs(buf, tpl.name); return tpl.name; });
  }

  // ---------- 書類の保管（案件くんの「書類保管」表） ----------
  function projectRecord(projectId) {
    return api('GET', '/k/v1/record', { app: appApi.getId(), id: projectId }).then(function (r) {
      if (!r.record[SUB.table]) throw new Error('案件くんに「' + SUB.table + '」の表がありません。先に scripts/add-doc-storage.js で追加してください');
      return r.record;
    });
  }
  function rowsOf(record) { return (record[SUB.table] && record[SUB.table].value) || []; }
  function cellText(row, code) { var f = row.value[code]; return f && f.value != null ? String(f.value) : ''; }
  // 書類名ごとの最新版（同じ版が並んだときは、後ろの行を採る）
  function existingOf(record) {
    var map = {};
    rowsOf(record).forEach(function (row) {
      var n = cellText(row, SUB.name), v = Number(cellText(row, SUB.version) || 0), files = row.value[SUB.file] && row.value[SUB.file].value;
      var pdfs = row.value[SUB.pdf] && row.value[SUB.pdf].value;
      if (n && (!map[n] || v >= map[n].version)) map[n] = {
        version: v, file: files && files[0] ? files[0] : null, rowId: row.id,
        approved: cellText(row, SUB.status) === APPROVED, approver: cellText(row, SUB.by), approvedAt: cellText(row, SUB.approvedAt),
        pdf: pdfs && pdfs[0] ? pdfs[0] : null
      };
    });
    return map;
  }
  // 書類保管の既存の行を、そのまま送り直せる形にする（送らなかった行は消えてしまうため）
  function keepRow(row) {
    var v = {};
    [SUB.name, SUB.stage, SUB.version, SUB.at].forEach(function (c) { v[c] = { value: cellText(row, c) }; });
    [SUB.status, SUB.by, SUB.approvedAt].forEach(function (c) { if (row.value[c]) v[c] = { value: cellText(row, c) }; });   // 列がある表だけ
    v[SUB.file] = { value: ((row.value[SUB.file] && row.value[SUB.file].value) || []).map(function (f) { return { fileKey: f.fileKey }; }) };
    if (row.value[SUB.pdf]) v[SUB.pdf] = { value: (row.value[SUB.pdf].value || []).map(function (f) { return { fileKey: f.fileKey }; }) };
    return { id: row.id, value: v };
  }
  function isoToday() { var d = new Date(); return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()); }
  function openStored(ex) {
    if (ex.data) { saveAs(ex.data, ex.filename); return Promise.resolve(ex.filename); }
    if (!ex.file) return Promise.reject(new Error('保存された書類のファイルが見つかりません'));
    return download(ex.file.fileKey).then(function (buf) { saveAs(buf, ex.file.name); return ex.file.name; });
  }

  // 書類1通を、案件の「書類保管」表に、1行足して保存する（作成も、修正版の登録も、ここを通る）
  function storeRow(record, ctx, docName, stage, version, filename, data) {
    return upload(data, filename).then(function (fileKey) {
      var row = { value: {} };
      row.value[SUB.name] = { value: docName };
      row.value[SUB.stage] = { value: stage };
      row.value[SUB.version] = { value: String(version) };
      row.value[SUB.at] = { value: isoToday() };
      row.value[SUB.file] = { value: [{ fileKey: fileKey }] };
      var update = {};
      update[SUB.table] = { value: rowsOf(record).map(keepRow).concat([row]) };
      // revision を付けて更新する：読んだあとに別の人が案件を更新していたら、上書きせずに失敗する
      return api('PUT', '/k/v1/record', { app: appApi.getId(), id: ctx.projectId, revision: record.$revision.value, record: update }).then(function () {
        return { filename: filename, version: version, data: data };
      }, function (e) {
        if (e && e.code === 'GAIA_CO02') throw new Error('別の人が同時に、この案件を更新しました。もう一度押してください');
        throw e;
      });
    });
  }

  function createDoc(masterRec, ctx, extra) {
    var docName = fieldText(masterRec[MF.name]);
    var files = masterRec[MF.template].value;
    if (!files || !files.length) return Promise.reject(new Error('「' + docName + '」の雛形が書類マスタに登録されていません'));
    var tpl = files[0];
    var ext = ((tpl.name.match(/\.([^.]+)$/) || [])[1] || '').toLowerCase();

    // 保存の直前に、案件を読み直して最新の版を確かめる（同時に別の人が作ったときに、同じ版が2つできないように）
    return Promise.all([projectRecord(ctx.projectId), download(tpl.fileKey)]).then(function (res) {
      var record = res[0], cur = existingOf(record)[docName];
      var version = (cur ? cur.version : 0) + 1;
      var blank = {}; ASK.forEach(function (a) { blank[a.key] = askDefault(a); });   // 入力されなかった項目の言葉（なければ空のまま入れて、あとで知らせる）
      var values = Object.assign(blank, ctx.values, extra || {}, { '版': 'v' + version, '書類名': docName });
      return fillTemplate(JSZip, res[1], ext, values).then(function (out) {
        var filename = buildFilename(ctx.projectName, docName, version, ext);
        return storeRow(record, ctx, docName, fieldText(masterRec[MF.stage]), version, filename, out.data).then(function (r) {
          r.unknown = out.unknown; r.empty = out.empty;
          return r;
        });
      });
    });
  }

  // 修正したファイルを、次の版（v2, v3…）として登録する（担当者が書類を開いて直し、その直したファイルを戻す）
  function registerRevision(masterRec, file, ctx) {
    var docName = fieldText(masterRec[MF.name]);
    var ext = extOf(file.name);
    if (!ext) return Promise.reject(new Error('ファイルの種類（.docx など）が分かりません'));
    return Promise.all([projectRecord(ctx.projectId), readFile(file)]).then(function (res) {
      var record = res[0], cur = existingOf(record)[docName];
      var version = (cur ? cur.version : 0) + 1;
      return storeRow(record, ctx, docName, fieldText(masterRec[MF.stage]), version, buildFilename(ctx.projectName, docName, version, ext), res[1]);
    });
  }
  // 書類保管の、指定した書類の最新版の行を、書き換えて保存する（change は、行の値を直接書き換える関数）
  function updateRows(ctx, names, change) {
    return projectRecord(ctx.projectId).then(function (record) {
      var ex = existingOf(record), ids = {};
      names.forEach(function (n) { if (!ex[n]) throw new Error('「' + n + '」は、まだ保存されていません'); ids[ex[n].rowId] = n; });
      var rows = rowsOf(record).map(keepRow);
      rows.forEach(function (r) { if (ids[r.id]) change(r, ids[r.id]); });
      var update = {}; update[SUB.table] = { value: rows };
      return api('PUT', '/k/v1/record', { app: appApi.getId(), id: ctx.projectId, revision: record.$revision.value, record: update }).catch(function (e) {
        if (e && e.code === 'GAIA_CO02') throw new Error('別の人が同時に、この案件を更新しました。もう一度押してください');
        throw e;
      });
    });
  }
  // 書類関係の承認：承認者（開いた人）と日時を、書類保管の行に残す。新しい版を作ると、承認は、外れる（新しい行には、承認がないため）
  function approveDocs(ctx, names) {
    var who = loginName() || '（不明）', now = new Date().toISOString();
    return updateRows(ctx, names, function (r) {
      r.value[SUB.status] = { value: APPROVED }; r.value[SUB.by] = { value: who }; r.value[SUB.approvedAt] = { value: now };
    }).then(function () { return { who: who, at: now }; });
  }
  // PDF化：承認済みの書類を、PDF化サーバー（pdf-server/）で変換して、同じ行の「PDF」に保存する
  function convertToPdf(ctx, name, ex) {
    if (!PDF_SERVER_URL) return Promise.reject(new Error('PDF化サーバーが、まだ設定されていません（doc-button.js の PDF_SERVER_URL）'));
    if (!ex.approved) return Promise.reject(new Error('承認されていない書類は、PDF化しません'));
    var ext = extOf((ex.file && ex.file.name) || ex.filename || '');
    return (ex.data ? Promise.resolve(ex.data) : download(ex.file.fileKey)).then(function (buf) {
      return fetch(PDF_SERVER_URL.replace(/\/$/, '') + '/convert?ext=' + encodeURIComponent(ext), { method: 'POST', headers: { 'Content-Type': 'application/octet-stream' }, body: buf });
    }).then(function (r) {
      if (!r.ok) throw new Error('PDF化に失敗しました（サーバーの返事：' + r.status + '）');
      return r.arrayBuffer();
    }).then(function (pdf) {
      var head = String.fromCharCode.apply(null, new Uint8Array(pdf.slice(0, 5)));
      if (head !== '%PDF-') throw new Error('PDFではないものが返ってきました。保存しませんでした');
      var base = ((ex.file && ex.file.name) || ex.filename || name).replace(/\.[^.]+$/, '');
      return upload(pdf, base + '.pdf').then(function (fileKey) {
        return updateRows(ctx, [name], function (r) { r.value[SUB.pdf] = { value: [{ fileKey: fileKey }] }; }).then(function () { return { name: base + '.pdf', data: pdf }; });
      });
    });
  }
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var fr = new FileReader();
      fr.onload = function () { resolve(fr.result); };
      fr.onerror = function () { reject(new Error('ファイルを読み込めませんでした')); };
      fr.readAsArrayBuffer(file);
    });
  }

  function loadMaster() {
    return findAppId(APP_NAMES.master).then(function (appId) {
      var q = MF.status + ' in (' + quote(ACTIVE) + ') order by ' + MF.order + ' asc limit 500';
      return api('GET', '/k/v1/records', { app: appId, query: q }).then(function (r) { return r.records; });
    });
  }

  // ---------- 画面 ----------
  function el(tag, style, text) {
    var e = document.createElement(tag);
    if (style) e.style.cssText = style;
    if (text) e.textContent = text;
    return e;
  }
  function makeSelect(options, labelAll) {
    var s = el('select', 'font-size:14px;padding:6px;margin:0 8px 8px 0;max-width:100%');
    if (labelAll) { var o0 = document.createElement('option'); o0.value = ''; o0.textContent = labelAll; s.appendChild(o0); }
    options.forEach(function (v) { var o = document.createElement('option'); o.value = v; o.textContent = v; s.appendChild(o); });
    return s;
  }
  function uniq(arr) { return arr.filter(function (v, i) { return v && arr.indexOf(v) === i; }); }
  // 書類の対象（共通／橋梁／一般道路…）。「共通」と未入力は、どの対象でも出す
  function targetOf(r) { var t = r[MF.target] ? fieldText(r[MF.target]) : ''; return t === '共通' ? '' : t; }
  function matchTarget(r, sel) { var t = targetOf(r); return !t || !sel || t === sel; }
  function hasTemplate(r) { return !!(r[MF.template] && r[MF.template].value && r[MF.template].value.length); }
  function unknownNote(list, empty) {
    var t = '';
    if (list && list.length) t += '\n置き換えできずに残った項目（名前が違う、または {{ }} が段落・セルをまたいでいます）：' + list.map(function (n) { return '{{' + n + '}}'; }).join(' ');
    // 書類のチェック（外部には送らず、この画面の中だけで行う）：案件くんの値が空のまま入った項目
    if (empty && empty.length) t += '\n案件くんの値が空のまま入った項目（案件くんに入力してから、作り直してください）：' + empty.map(function (n) { return '{{' + n + '}}'; }).join(' ');
    return t;
  }

  function openPanel(ctx) {
    var overlay = el('div', 'position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:10000;overflow:auto;-webkit-overflow-scrolling:touch');
    var box = el('div', 'background:#fff;max-width:720px;margin:24px auto;border-radius:8px;padding:16px;font-size:14px;box-sizing:border-box;width:calc(100% - 24px)');
    var head = el('div', 'display:flex;justify-content:space-between;align-items:center;margin-bottom:10px');
    head.appendChild(el('div', 'font-weight:bold;font-size:17px', '工程別の書類：' + ctx.projectName));
    var close = el('button', 'padding:6px 12px;font-size:14px;cursor:pointer;white-space:nowrap;flex:none;margin-left:10px', '閉じる'); close.type = 'button';
    close.addEventListener('click', function () { document.body.removeChild(overlay); });
    head.appendChild(close);
    box.appendChild(head);
    var body = el('div', '', '読み込み中…');
    box.appendChild(body);
    overlay.appendChild(box);
    document.body.appendChild(overlay);

    var masters = [], existing = {};
    Promise.all([loadMaster(), projectRecord(ctx.projectId)]).then(function (res) {
      masters = res[0]; existing = existingOf(res[1]);
      build();
    }).catch(function (e) { body.textContent = ''; body.appendChild(el('div', 'color:#c0392b', '読み込めませんでした：' + ((e && e.message) || ''))); });

    function build() {
      body.innerHTML = '';
      if (!masters.length) { body.textContent = '書類マスタに、使用中の書類が登録されていません。'; return; }
      var stages = uniq(masters.map(function (r) { return fieldText(r[MF.stage]); }));
      var auds = uniq([].concat.apply([], masters.map(function (r) { return (r[MF.audience] && r[MF.audience].value) || []; })));

      var stageSel = makeSelect(stages);
      if (ctx.currentStage && stages.indexOf(ctx.currentStage) >= 0) stageSel.value = ctx.currentStage;
      var audSel = makeSelect(auds, 'すべての立場');
      try { var sv = localStorage.getItem(STORE_KEY); if (sv && auds.indexOf(sv) >= 0) audSel.value = sv; } catch (e) {}
      // 対象（橋梁／一般道路）：書類マスタに「共通」以外の書類があるときだけ出す
      var targets = uniq(masters.map(targetOf));
      var tgtSel = makeSelect(targets, 'すべての対象');
      var autoNote = '';
      if (ctx.target && targets.indexOf(ctx.target) >= 0) tgtSel.value = ctx.target;
      else if (targets.indexOf(ctx.guessedTarget) >= 0) { tgtSel.value = ctx.guessedTarget; autoNote = '対象は、案件名から「' + ctx.guessedTarget + '」と判断して選びました。違うときは選び直してください。'; }
      var bar = el('div'); bar.appendChild(stageSel); if (auds.length) bar.appendChild(audSel); if (targets.length) bar.appendChild(tgtSel);
      var batch = el('button', 'padding:6px 12px;font-size:13px;cursor:pointer;margin-bottom:8px', 'この工程をまとめて作成'); batch.type = 'button';
      var batchAll = el('button', 'padding:6px 12px;font-size:13px;cursor:pointer;margin-bottom:8px;margin-left:8px', '全工程をまとめて作成'); batchAll.type = 'button';
      bar.appendChild(batch); bar.appendChild(batchAll);
      if (ctx.caps && ctx.caps.approval) {
        var approveAll = el('button', 'padding:6px 12px;font-size:13px;cursor:pointer;margin-bottom:8px;margin-left:8px;font-weight:bold', '書類一式を承認'); approveAll.type = 'button';
        approveAll.addEventListener('click', function () {
          var names = Object.keys(existing).filter(function (n) { return !existing[n].approved; });
          approveNames(names, '未承認の書類');
        });
        bar.appendChild(approveAll);
      }
      var list = el('div'), msg = el('div', 'margin-top:8px;font-size:13px;white-space:pre-line');
      if (autoNote && targets.length) bar.appendChild(el('div', 'font-size:12px;color:#666;margin-bottom:6px', autoNote));
      body.appendChild(bar); body.appendChild(list); body.appendChild(msg);

      // 工程の選択肢に「作成済 x/y」を出す（y＝案件の情報を入れる書類。雛形のまま使う書類・雛形が未登録の書類は数えない）
      function updateStageLabels() {
        Array.prototype.forEach.call(stageSel.options, function (o) {
          var rows = masters.filter(function (r) {
            if (fieldText(r[MF.stage]) !== o.value) return false;
            return (!audSel.value || ((r[MF.audience] && r[MF.audience].value) || []).indexOf(audSel.value) >= 0) && matchTarget(r, tgtSel.value);
          }).filter(function (r) { return hasTemplate(r) && kindOf(r) !== 'plain'; });
          var done = rows.filter(function (r) { return existing[fieldText(r[MF.name])]; }).length;
          o.textContent = o.value + (rows.length ? '（作成済 ' + done + '/' + rows.length + '）' : '');
        });
      }
      function setMsg(t, err) { msg.textContent = t; msg.style.color = err ? '#c0392b' : '#1f7a3d'; }
      // 注意（置き換えできなかった項目・空の項目）があるときは、読めるように、画面を自動では更新しない
      function reloadSoon(t, stay) {
        if (stay) { setMsg(t + '\n（保存はできています。内容を確認してください。案件の「書類保管」表は、画面を開き直すと更新されます）', true); batch.disabled = false; batchAll.disabled = false; updateStageLabels(); show(); return; }
        setMsg(t + '（画面を更新します）', false); setTimeout(function () { location.reload(); }, 2500);
      }
      function shown() {
        return masters.filter(function (r) {
          if (fieldText(r[MF.stage]) !== stageSel.value) return false;
          return (!audSel.value || ((r[MF.audience] && r[MF.audience].value) || []).indexOf(audSel.value) >= 0) && matchTarget(r, tgtSel.value);
        });
      }
      // 作成に成功したら、画面上の「最新 vN」もその場で更新する（まとめ作成が途中で失敗しても、再実行で二重に作らない）
      function makeOne(rec, extra) {
        return createDoc(rec, ctx, extra).then(function (r) {
          existing[fieldText(rec[MF.name])] = { version: r.version, filename: r.filename, data: r.data };
          return r;
        });
      }
      // 議事録などの入力欄（雛形に {{打合せ日}} {{議事録メモ}} などがあるときだけ出す）
      // Microsoft Copilot に渡す依頼文（コピーして、Copilot のチャットに貼り、続けて、文字起こしを貼る）。結果は、見出し付きになるので、そのまま貼り付けて「かんたん整理」で、入力欄に入る
  var COPILOT_PROMPT = '以下は、打合せの音声の文字起こしです。次の見出しで、箇条書きにまとめてください。分からない項目は、空にしてください。書かれていないことは、作らないでください。「えー」「あの」などは、取り除いてください。\n【出席者】\n【場所】\n【議題】\n【打合せ内容】\n【決定事項】\n【ToDo】（「期限｜担当｜内容」の形で、1行に1つ。例：10/25｜太田｜見積を出す）\n【次回予定】\n\n【文字起こし】\n';
  var COPILOT_URL = 'https://m365.cloud.microsoft/chat';   // Microsoft 365 Copilot のチャット（会社の契約。開けなければ、ここを直す）
  var MINUTES_SERVER_URL = '';   // AIで議事録を整理するサーバーのURL（pdf-server の /minutes）。空のあいだは、「AIで整理」は押せない
  function askValues(asks, name) {
        return new Promise(function (resolve) {
          var cover = el('div', 'position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:10001;overflow:auto');
          var card = el('div', 'background:#fff;max-width:640px;margin:24px auto;border-radius:8px;padding:16px;font-size:14px;box-sizing:border-box;width:calc(100% - 24px)');
          card.appendChild(el('div', 'font-weight:bold;font-size:16px;margin-bottom:10px', '「' + name + '」の内容'));
          var inputs = {};
          var single = asks.some(function (a) { return a.key === '議事録本文'; });   // 議事録本文：ひとつの四角に、見出し付きで、まとめて入れる
          var isMinutes = single || asks.some(function (a) { return a.key === '議事録メモ'; });
          var shown = single ? asks.filter(function (a) { return a.key === '議事録本文'; }) : asks;   // 打合せ日・作成者などは、入力欄を出さず、自動で入れる
          var fillFrom = function (res) {
            if (inputs['議事録本文']) { inputs['議事録本文'].value = formatMinutes(res); return; }
            Object.keys(res).forEach(function (key) { if (inputs[key] && res[key]) inputs[key].value = res[key]; });
          };
          if (isMinutes) {
            card.appendChild(el('div', 'font-size:14px;font-weight:bold;color:#1f3a6e;margin-top:4px', '① 音声の文字起こしを、ここに貼り付けます'));
            var src = el('textarea', 'width:100%;height:300px;box-sizing:border-box;font-size:14px;padding:8px;margin-top:4px;line-height:1.5;background:#f7fafc'); src.setAttribute('data-minutes-src', '1');
            card.appendChild(src);
            // Copilot を使う場合：①依頼文をコピー → ②Copilot を開く → 依頼文と、文字起こしを貼る → できた文章を、上の欄に貼る → 「かんたん整理」
            var cop = el('div', 'margin:6px 0 2px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;font-size:12px');
            cop.appendChild(el('span', 'color:#666', 'Copilotを使う場合：'));
            var copy = el('button', 'padding:5px 9px;font-size:12px;cursor:pointer', '①依頼文をコピー'); copy.type = 'button';
            copy.addEventListener('click', function () {
              var done = function () { note.textContent = 'コピーしました。Copilotのチャットに貼り、続けて、文字起こしを貼ってください。できた文章を、上の欄に貼って、「かんたん整理」を押します。'; };
              var fallback = function () { var t = document.createElement('textarea'); t.value = COPILOT_PROMPT; document.body.appendChild(t); t.select(); try { document.execCommand('copy'); done(); } catch (e) { note.textContent = 'コピーできませんでした。'; } document.body.removeChild(t); };
              if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(COPILOT_PROMPT).then(done, fallback); else fallback();
            });
            var openCop = el('button', 'padding:5px 9px;font-size:12px;cursor:pointer', '②Copilotを開く'); openCop.type = 'button';
            openCop.addEventListener('click', function () { window.open(COPILOT_URL, '_blank', 'noopener'); });
            cop.appendChild(copy); cop.appendChild(openCop); card.appendChild(cop);
            var tools = el('div', 'margin:4px 0 8px;display:flex;gap:8px;flex-wrap:wrap;align-items:center');
            var easy = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer', 'かんたん整理（外部に送りません）'); easy.type = 'button';
            easy.addEventListener('click', function () { fillFrom(summarizeLocal(src.value)); note.textContent = '簡易に分けました。内容を、確認・修正してください。'; });
            var ai = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer;font-weight:bold', MINUTES_SERVER_URL ? 'AIで整理（外部のAIに送ります）' : 'AIで整理（未設定）'); ai.type = 'button'; ai.disabled = !MINUTES_SERVER_URL;
            var note = el('span', 'font-size:12px;color:#1f7a3d');
            var extra = el('input', 'width:100%;box-sizing:border-box;font-size:12px;padding:4px;margin-bottom:4px'); extra.type = 'text'; extra.placeholder = 'AIに送らず、伏せたい言葉（コンマ区切り。例：会社名、物件名）'; extra.setAttribute('data-mask-extra', '1');
            card.appendChild(el('div', 'font-size:14px;font-weight:bold;color:#1f3a6e;margin-top:8px', '② 下のボタンで、整理します（整理した結果は、③に入ります）'));
            card.appendChild(extra);
            ai.addEventListener('click', function () {
              if (!src.value.trim()) { note.textContent = '文字起こしを、貼り付けてください'; return; }
              var mk = maskForAI(src.value, extra.value.split(/[,、，]/).map(function (x) { return x.trim(); }));
              var preview = mk.masked.length > 400 ? mk.masked.slice(0, 400) + '…（以下、省略）' : mk.masked;
              if (!confirm('次の内容を、外部のAIに送って、整理します（電話番号・メール・金額・「〇〇さん」の名前などは、符号に置き換えて、送ります。結果は、元の言葉に戻します）。\n\n【送る内容】\n' + preview + '\n\nよろしいですか？')) return;
              ai.disabled = true; note.textContent = 'AIが、整理しています…';
              fetch(MINUTES_SERVER_URL.replace(/\/$/, '') + '/minutes', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ text: mk.masked }) })
                .then(function (r) { if (!r.ok) throw new Error('AIの整理に失敗しました（' + r.status + '）'); return r.json(); })
                .then(function (res) { Object.keys(res).forEach(function (k) { res[k] = mk.unmask(res[k]); }); fillFrom(res); ai.disabled = false; note.textContent = 'AIが分けました（置き換えた言葉 ' + mk.count + ' 個は、元に戻しました）。内容を、確認・修正してください（AIの案です。間違いが、含まれる可能性があります）。'; })
                .catch(function (e) { ai.disabled = false; note.style.color = '#c0392b'; note.textContent = (e && e.message) || 'AIの整理に失敗しました'; });
            });
            tools.appendChild(easy); tools.appendChild(ai); tools.appendChild(note); card.appendChild(tools);
          }
          shown.forEach(function (a) {
            card.appendChild(el('div', a.key === '議事録本文' ? 'font-size:14px;font-weight:bold;color:#1f3a6e;margin-top:8px' : 'font-size:12px;color:#666;margin-top:8px', a.label));
            var inp = a.type === 'doc' ? el('textarea', 'width:100%;height:150px;box-sizing:border-box;font-size:14px;padding:8px;line-height:1.5') : a.type === 'memo' ? el('textarea', 'width:100%;height:160px;box-sizing:border-box;font-size:14px;padding:6px') : el('input', 'width:100%;box-sizing:border-box;font-size:14px;padding:6px');
            if (a.type === 'date') { inp.type = 'date'; inp.value = isoToday(); } else if (a.type === 'text') { inp.type = 'text'; if (a.init) inp.value = a.init(); }
            inp.setAttribute('data-ask', a.key); inputs[a.key] = inp; card.appendChild(inp);
          });
          var bar = el('div', 'margin-top:12px;display:flex;gap:8px;justify-content:flex-end');
          var cancel = el('button', 'padding:8px 14px;cursor:pointer', 'やめる'); cancel.type = 'button';
          var ok = el('button', 'padding:8px 14px;cursor:pointer;font-weight:bold', '作成して保存'); ok.type = 'button';
          function close(v) { document.body.removeChild(cover); resolve(v); }
          cancel.addEventListener('click', function () { close(null); });
          ok.addEventListener('click', function () {
            var v = {};
            asks.forEach(function (a) {
              if (!inputs[a.key]) { v[a.key] = askDefault(a); return; }   // 入力欄を出さなかった項目は、初期値（打合せ日は今日、作成者は開いた人）
              var t = inputs[a.key].value;
              if (a.type === 'date') { var m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/); t = m ? m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日' : t; }
              v[a.key] = a.type === 'memo' ? tidyMemo(t) : a.type === 'doc' ? t.replace(/\s+$/, '') : t.trim();
              if (!v[a.key]) v[a.key] = askDefault(a);
            });
            close(v);
          });
          bar.appendChild(cancel); bar.appendChild(ok); card.appendChild(bar); cover.appendChild(card); document.body.appendChild(cover);
        });
      }
      // 保存した最新の書類を、ダウンロードする前に確認する（外部には送りません）。指摘のリストを返す
      function checkOne(name) {
        var cur = existing[name];
        return (cur.data ? Promise.resolve(cur.data) : download(cur.file.fileKey)).then(function (buf) {
          return extractText(JSZip, buf, extOf(cur.filename || (cur.file && cur.file.name) || ''));
        }).then(function (text) { return checkText(text, ctx.values); });
      }
      function approveNames(names, label) {
        if (!names.length) { setMsg('承認する書類が、ありません', true); return; }
        setMsg('承認の前に、書類をチェックしています…', false);
        names.reduce(function (p, n) { return p.then(function (acc) { return checkOne(n).then(function (f) { acc[n] = f; return acc; }); }); }, Promise.resolve({})).then(function (found) {
          var bad = names.filter(function (n) { return found[n].length; });
          var msgText = label + '（' + names.length + '件）を、承認します。承認者：' + (loginName() || '（不明）') + '\n' +
            (bad.length ? '\n【チェックの指摘があります】\n' + bad.map(function (n) { return '・' + n + '：' + found[n].join(' ／ '); }).join('\n') + '\n\n指摘があっても、承認しますか？' : '\nチェックの指摘は、ありません。承認しますか？');
          if (!confirm(msgText)) { setMsg('承認は、しませんでした', false); return; }
          return approveDocs(ctx, names).then(function (r) {
            names.forEach(function (n) { existing[n].approved = true; existing[n].approver = r.who; existing[n].approvedAt = r.at; });
            setMsg(names.length + '件を承認しました（' + r.who + '）', false); show();
          });
        }).catch(function (e) { setMsg((e && e.message) || '承認できませんでした', true); });
      }
      function makeRow(rec) {
        var name = fieldText(rec[MF.name]), ex = existing[name];
        var ok = hasTemplate(rec), kind = kindOf(rec);
        var row = el('div', 'display:flex;gap:10px;align-items:center;justify-content:space-between;padding:8px 0;border-top:1px solid #eee;flex-wrap:wrap');
        var left = el('div', 'flex:1 1 200px');
        left.appendChild(el('div', 'font-weight:bold', name));
        left.appendChild(el('div', 'font-size:12px;color:#666', (targetOf(rec) ? '【' + targetOf(rec) + '用】' : '') + fieldText(rec[MF.desc])));
        var state = el('div', 'font-size:12px;color:#1f4e79');
        if (ok && kind === 'plain') state.textContent = '雛形のまま使う書類（案件の情報は入れません）';
        else if (ex) {
          state.appendChild(document.createTextNode('最新 v' + ex.version + ' ・ '));
          var a = el('a', 'cursor:pointer;text-decoration:underline;color:#1a73e8', '開く'); a.setAttribute('role', 'link');
          a.addEventListener('click', function () { openStored(existing[name]).catch(function (e) { setMsg((e && e.message) || '開けませんでした', true); }); });
          state.appendChild(a);
          if (ctx.caps && ctx.caps.approval) {
            state.appendChild(document.createTextNode(' ・ '));
            var ap = el('span', ex.approved ? 'color:#1f7a3d;font-weight:bold' : 'color:#a05a00', ex.approved ? '承認済み（' + ex.approver + '・' + String(ex.approvedAt).slice(0, 10) + '）' : '未承認');
            state.appendChild(ap);
          }
        } else state.textContent = '未作成';
        left.appendChild(state);
        var label = !ok ? '雛形が未登録' : !kind ? '確認中…' : kind === 'plain' ? '雛形を開く' : ex ? 'v' + (ex.version + 1) + ' を作成' : '作成して保存';
        var btn = el('button', 'padding:8px 14px;font-size:14px;cursor:pointer', label); btn.type = 'button';
        btn.disabled = !ok || !kind;
        btn.addEventListener('click', function () {
          if (kind === 'plain') {
            btn.disabled = true; setMsg('雛形を取り出しています…', false);
            openAsIs(rec).then(function (n) { btn.disabled = false; setMsg('雛形のまま保存しました：' + n + '（案件には保存していません）', false); })
              .catch(function (e) { btn.disabled = false; setMsg((e && e.message) || '取り出せませんでした', true); });
            return;
          }
          var cur = existing[name];
          if (cur && !confirm('「' + name + '」は v' + cur.version + ' まで作成済みです。v' + (cur.version + 1) + ' を作成しますか？')) return;
          var asks = asksOf(rec);
          (asks.length ? askValues(asks, name) : Promise.resolve({})).then(function (extra) {
            if (!extra) return;
            btn.disabled = true; setMsg('作成しています…', false);
            return makeOne(rec, extra).then(function (r) { saveAs(r.data, r.filename); reloadSoon('保存してダウンロードしました：' + r.filename + unknownNote(r.unknown, r.empty), !!((r.unknown && r.unknown.length) || (r.empty && r.empty.length))); })
              .catch(function (e) { btn.disabled = false; setMsg((e && e.message) || '作成できませんでした', true); });
          });
        });
        var right = el('div', 'display:flex;flex-direction:column;gap:6px;align-items:stretch');
        right.appendChild(btn);
        // 直したファイルを、次の版として登録する（作成済みの書類だけ）
        if (ex && ok && kind !== 'plain') {
          var pick = document.createElement('input'); pick.type = 'file'; pick.style.display = 'none'; pick.setAttribute('data-revision', name);
          var rev = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer', '直したファイルを登録'); rev.type = 'button';
          rev.addEventListener('click', function () { pick.click(); });
          pick.addEventListener('change', function () {
            var f = pick.files && pick.files[0]; if (!f) return;
            if (!confirm('「' + f.name + '」を、「' + name + '」の v' + (ex.version + 1) + ' として登録しますか？')) { pick.value = ''; return; }
            rev.disabled = true; setMsg('登録しています…', false);
            registerRevision(rec, f, ctx).then(function (r) {
              existing[name] = { version: r.version, filename: r.filename, data: r.data };
              setMsg('v' + r.version + ' として登録しました：' + r.filename, false); updateStageLabels(); show();
            }).catch(function (e) { rev.disabled = false; setMsg((e && e.message) || '登録できませんでした', true); });
          });
          // 書類のチェック：保存した最新の書類を、ダウンロードする前に、画面の中だけで確認する（外部には送りません）
          var chk = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer', 'チェック'); chk.type = 'button';
          chk.addEventListener('click', function () {
            chk.disabled = true; setMsg('「' + name + '」をチェックしています…', false);
            checkOne(name).then(function (res) {
              chk.disabled = false;
              setMsg(res.length ? '「' + name + '」のチェック結果（' + res.length + '件）：\n・' + res.join('\n・') : '「' + name + '」：問題は見つかりませんでした', res.length > 0);
            }).catch(function (e) { chk.disabled = false; setMsg((e && e.message) || 'チェックできませんでした', true); });
          });
          right.appendChild(chk); right.appendChild(rev); right.appendChild(pick);
          // 書類関係の承認（承認した人と日時が、書類保管の行に残る。新しい版を作ると、承認は、外れる）
          if (ctx.caps && ctx.caps.approval && !ex.approved) {
            var apb = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer;font-weight:bold', '承認'); apb.type = 'button';
            apb.addEventListener('click', function () { approveNames([name], '「' + name + '」v' + ex.version); });
            right.appendChild(apb);
          }
          // PDF化（承認済みの書類だけ）
          if (ctx.caps && ctx.caps.pdf && ex.approved) {
            if (ex.pdf) {
              var pdfOpen = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer', 'PDFを開く'); pdfOpen.type = 'button';
              pdfOpen.addEventListener('click', function () { download(ex.pdf.fileKey).then(function (buf) { saveAs(buf, ex.pdf.name); }).catch(function (e) { setMsg((e && e.message) || '開けませんでした', true); }); });
              right.appendChild(pdfOpen);
            } else {
              var pdfb = el('button', 'padding:6px 10px;font-size:12px;cursor:pointer', PDF_SERVER_URL ? 'PDF化' : 'PDF化（サーバー未設定）'); pdfb.type = 'button'; pdfb.disabled = !PDF_SERVER_URL;
              pdfb.addEventListener('click', function () {
                pdfb.disabled = true; setMsg('PDF化しています…', false);
                convertToPdf(ctx, name, existing[name]).then(function (r) { existing[name].pdf = { fileKey: null, name: r.name }; saveAs(r.data, r.name); setMsg('PDFを保存しました：' + r.name + '（画面を開き直すと、「PDFを開く」になります）', false); })
                  .catch(function (e) { pdfb.disabled = false; setMsg((e && e.message) || 'PDF化できませんでした', true); });
              });
              right.appendChild(pdfb);
            }
          }
        }
        row.appendChild(left); row.appendChild(right);
        return row;
      }
      function show() {
        list.innerHTML = '';
        var rows = shown();
        if (!rows.length) { list.textContent = 'この条件の書類はありません。'; return; }
        rows.forEach(function (rec) {
          var row = makeRow(rec);
          list.appendChild(row);
          // 雛形に {{ }} があるかを裏で調べ、分かったら、その行の表示とボタンを切り替える
          if (hasTemplate(rec) && !kindOf(rec)) inspect(rec).then(function () { if (row.parentNode === list) list.replaceChild(makeRow(rec), row); });
        });
      }
      // まとめて作成：pick() が返す書類のうち、雛形のあるものを作る。onlyMissing なら、未作成のものだけ（全工程のとき）
      function runBatch(btn, pick, onlyMissing, scope) {
        var withTpl = pick().filter(hasTemplate);
        if (!withTpl.length) { setMsg('雛形が登録された書類がありません', true); return; }
        btn.disabled = true; setMsg('雛形を確認しています…', false);
        Promise.all(withTpl.map(inspect)).then(function () {
          var targets = withTpl.filter(function (r) { return kindOf(r) !== 'plain' && (!onlyMissing || !existing[fieldText(r[MF.name])]); });
          var asIs = withTpl.filter(function (r) { return kindOf(r) === 'plain'; }).length;
          btn.disabled = false; setMsg('', false);
          if (!targets.length) { setMsg(onlyMissing ? scope + 'の書類は、すべて作成済みか、雛形のまま使うものです' : scope + 'の書類は、すべて雛形のまま使うものです（各行の「雛形を開く」で取り出せます）', false); return; }
          if (!confirm(targets.length + '件の書類を作成します。' + (onlyMissing ? '（作成済みの書類は、作りません）' : '作成済みの書類は次の版として作られます。') + (asIs ? '\n（雛形のまま使う ' + asIs + '件は、作成の対象外です）' : '') + '\nよろしいですか？')) return;
          btn.disabled = true;
          var done = [], failed = [], unknown = [], empty = [];
          targets.reduce(function (p, rec) {
            return p.then(function () {
              setMsg('作成中…（' + (done.length + failed.length + 1) + '/' + targets.length + '）', false);
              return makeOne(rec).then(function (r) {
                done.push(r.filename);
                (r.unknown || []).forEach(function (n) { if (unknown.indexOf(n) < 0) unknown.push(n); });
                (r.empty || []).forEach(function (n) { if (empty.indexOf(n) < 0) empty.push(n); });
              }, function (e) { failed.push(fieldText(rec[MF.name]) + '：' + ((e && e.message) || '失敗')); });
            });
          }, Promise.resolve()).then(function () {
            updateStageLabels();
            if (failed.length) { btn.disabled = false; show(); setMsg('作成 ' + done.length + '件、失敗 ' + failed.length + '件\n' + failed.join('\n') + unknownNote(unknown, empty), true); }
            else reloadSoon(done.length + '件を保存しました' + unknownNote(unknown, empty), !!(unknown.length || empty.length));
          });
        });
      }
      batch.addEventListener('click', function () { runBatch(batch, shown, false, 'この工程'); });
      batchAll.addEventListener('click', function () {
        runBatch(batchAll, function () {
          return masters.filter(function (r) {
            return (!audSel.value || ((r[MF.audience] && r[MF.audience].value) || []).indexOf(audSel.value) >= 0) && matchTarget(r, tgtSel.value);
          });
        }, true, '全工程');
      });
      stageSel.addEventListener('change', show);
      tgtSel.addEventListener('change', function () { updateStageLabels(); show(); });
      audSel.addEventListener('change', function () { try { localStorage.setItem(STORE_KEY, audSel.value); } catch (e) {} updateStageLabels(); show(); });
      updateStageLabels();
      show();
      // 全部の雛形を裏で確認して、工程ごとの数を正しくする
      masters.filter(hasTemplate).forEach(function (r) { inspect(r).then(updateStageLabels); });
    }
  }

  // 画面の右下に浮かぶ丸いボタン（書類関係タブの中に置けないときの予備）。マウスを近づけると、文字つきに広がる。いつでも見える
  var FAB_SIZE = 52;
  function floatingButton(label, onClick) {
    var old = document.querySelector('[data-doc-hub]'); if (old && old.parentNode) old.parentNode.removeChild(old);
    var fab = el('button', 'position:fixed;right:16px;bottom:96px;z-index:9999;height:' + FAB_SIZE + 'px;min-width:' + FAB_SIZE + 'px;border:0;border-radius:' + (FAB_SIZE / 2) + 'px;background:#0b5fb8;color:#fff;box-shadow:0 4px 12px rgba(0,0,0,.35);cursor:pointer;display:flex;align-items:center;justify-content:center;padding:0 14px;font-size:14px;font-weight:bold;overflow:hidden;white-space:nowrap');
    fab.type = 'button'; fab.setAttribute('data-doc-hub', '1'); fab.setAttribute('aria-label', label);
    var icon = el('span', 'font-size:22px;line-height:1', '📄');
    var text = el('span', 'display:inline-block;max-width:0;opacity:0;margin-left:0;transition:max-width .25s,opacity .25s,margin-left .25s', label);
    fab.appendChild(icon); fab.appendChild(text);
    function open() { text.style.maxWidth = '160px'; text.style.opacity = '1'; text.style.marginLeft = '8px'; }
    function close() { text.style.maxWidth = '0'; text.style.opacity = '0'; text.style.marginLeft = '0'; }
    fab.addEventListener('mouseenter', open); fab.addEventListener('focus', open);
    fab.addEventListener('mouseleave', close); fab.addEventListener('blur', close);
    fab.addEventListener('click', onClick);
    document.body.appendChild(fab);
    return fab;
  }

  // ---------- 「書類関係」タブの、すぐ下に置くボタン ----------
  // 「タブ表示」プラグインが、どの項目を、どのタブに入れているかに、頼りません。タブの並び（タブ名）を見つけて、
  // 「書類関係」のタブが選ばれている間だけ、その並びの、すぐ下に、ボタンを出します
  var TAB_NAMES = ['書類関係'];
  function bgOf(el) {
    for (var i = 0; el && i < 4; i++, el = el.parentElement) {
      var c = getComputedStyle(el).backgroundColor;
      if (c && c !== 'transparent' && !/rgba\(\s*0,\s*0,\s*0,\s*0\s*\)/.test(c)) return c;
    }
    return '';
  }
  function leafByText(text) {
    var all = document.querySelectorAll('div,span,td,th,li,a,button,p');
    for (var i = 0; i < all.length; i++) { if (!all[i].children.length && all[i].textContent.trim() === text) return all[i]; }
    return null;
  }
  function leavesOf(row) {
    return Array.prototype.filter.call(row.querySelectorAll('div,span,td,th,li,a,button'), function (e) { return !e.children.length && e.textContent.trim(); });
  }
  // タブの並びを探す：「書類関係」の文字を持つ要素と、その並び全体（「工事部」「施工店」も入っている、いちばん内側の枠）
  function findTabBar() {
    for (var n = 0; n < TAB_NAMES.length; n++) {
      var me = leafByText(TAB_NAMES[n]); if (!me) continue;
      var row = me.parentElement;
      for (var i = 0; row && i < 5; i++, row = row.parentElement) {
        var t = row.textContent;
        if (t.indexOf('工事部') >= 0 && t.indexOf('施工店') >= 0 && leavesOf(row).length >= 5) return { me: me, row: row };
      }
    }
    return null;
  }
  // そのタブが、選ばれているか：ほかのタブと、見た目（背景色・文字の色・太さ）が違うものが、選ばれているタブ。
  // 見た目に差がなく、見分けられないときは、選ばれているものとして、扱う（ボタンが、出なくなるのを、避けるため）
  function tabIsActive(bar) {
    try {
      var sig = function (e) { var cs = getComputedStyle(e); return bgOf(e) + '|' + cs.color + '|' + cs.fontWeight; };
      var mine = sig(bar.me), others = leavesOf(bar.row).filter(function (e) { return e !== bar.me && !e.contains(bar.me) && !bar.me.contains(e); });
      if (others.length < 2) return true;
      var same = others.filter(function (e) { return sig(e) === mine; }).length;
      var kinds = {}; others.forEach(function (e) { kinds[sig(e)] = 1; });
      if (Object.keys(kinds).length === 1 && same === others.length) return true;   // 全部同じ見た目：見分けられない
      return same < others.length / 2;
    } catch (e) { return true; }
  }
  function tabBarButton(label, bar, onClick) {
    var olds = document.querySelectorAll('[data-doc-hub]'); Array.prototype.forEach.call(olds, function (o) { if (o.parentNode) o.parentNode.removeChild(o); });
    var wrap = el('div', 'padding:8px 0 4px');
    wrap.setAttribute('data-doc-hub', '1');
    var b = el('button', 'padding:10px 18px;font-size:15px;font-weight:bold;cursor:pointer;background:#0b5fb8;color:#fff;border:0;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,.25)', '📄 ' + label);
    b.type = 'button'; b.addEventListener('click', onClick);
    wrap.appendChild(b);
    bar.row.parentNode.insertBefore(wrap, bar.row.nextSibling);
    var sync = function () { wrap.style.display = tabIsActive(bar) ? 'block' : 'none'; };
    sync(); var timer = setInterval(function () { if (!document.body.contains(wrap) || !document.body.contains(bar.me)) { clearInterval(timer); return; } sync(); }, 400);
    return wrap;
  }

  // 書類関係タブの中（登録画面の、いちばん上）に置くボタン。目印の項目（ANCHOR_FIELD）の行の、すぐ上に入れる。
  // 目印が見えている間（＝書類関係タブを開いている間）だけ見える。目印が見つからなければ、null を返す
  var ANCHOR_FIELD = '安全書類格納完了チェック';   // 案件くんの、書類関係タブの中にある項目（フィールドコード。案件くんのレイアウトで確認済み）
  function inlineButton(label, anchor, onClick) {
    // 目印の項目を入れている枠の、すぐ上に入れる。キントーンの枠の呼び名が違っても、置けるように、見つかる一番外の枠を使う
    var row = (anchor.closest && (anchor.closest('.row-gaia') || anchor.closest('.control-gaia'))) || anchor.parentElement;
    if (!row || !row.parentNode) return null;
    var olds = document.querySelectorAll('[data-doc-hub]'); Array.prototype.forEach.call(olds, function (o) { if (o.parentNode) o.parentNode.removeChild(o); });
    var wrap = el('div', 'padding:6px 0 10px');
    wrap.setAttribute('data-doc-hub', '1');
    var b = el('button', 'padding:10px 18px;font-size:15px;font-weight:bold;cursor:pointer;background:#0b5fb8;color:#fff;border:0;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,.25)', '📄 ' + label);
    b.type = 'button'; b.addEventListener('click', onClick);
    wrap.appendChild(b);
    row.parentNode.insertBefore(wrap, row);
    var sync = function () { wrap.style.display = anchor.offsetParent !== null ? 'block' : 'none'; };
    sync(); var timer = setInterval(function () { if (!document.body.contains(wrap)) { clearInterval(timer); return; } sync(); }, 400);
    return wrap;
  }

  function onShow(event) {
    if (typeof JSZip === 'undefined') { console.error('JSZipが読み込まれていません。JavaScriptの読み込み順を確認してください。'); return event; }
    var record = event.record, projectId = recApi.getId();
    var anchor = null;
    try { anchor = recApi.getFieldElement(ANCHOR_FIELD); } catch (e) { anchor = null; }
    // 置き場所の順番：①タブの並びの、すぐ下（書類関係タブが選ばれている間だけ）→ ②目印の項目の、すぐ上 → ③右下の丸いボタン（いつでも）
    var place = function (label, onClick) {
      var bar = findTabBar();
      if (bar) { tabBarButton(label, bar, onClick); return; }
      // タブの並びが、まだ見つからない：先に、予備（②→③）を出しておき、あとでタブの並びが見つかったら、そちらに移す（約6秒、探す）
      var tries = 0;
      if (!(anchor && inlineButton(label, anchor, onClick))) floatingButton(label, onClick);
      var again = function () {
        var b2 = findTabBar();
        if (b2) { tabBarButton(label, b2, onClick); return; }
        if (++tries < 20) setTimeout(again, 300);
      };
      setTimeout(again, 300);
    };
    place('書類を作る', function () {
      loadLabelMap(record).then(function (lm) {
        var values = buildValues(record, lm);
        var tbl = lm.props && lm.props[SUB.table], tf = (tbl && tbl.fields) || {};
        var caps = { approval: !!(tf[SUB.status] && tf[SUB.by] && tf[SUB.approvedAt]), pdf: !!tf[SUB.pdf] };   // 表に、その列があるときだけ、承認・PDF化を出す
        var projectName = '';
        NAME_LABELS.some(function (l) { if (values[l]) { projectName = values[l]; return true; } return false; });
        openPanel({
          values: values, caps: caps, projectId: projectId, projectName: projectName || ('案件' + projectId),
          currentStage: values[STAGE_LABEL] || '', target: TARGET_LABEL ? values[TARGET_LABEL] || '' : '', guessedTarget: guessTarget(projectName)
        });
      });
    });
    return event;
  }
  kintone.events.on(['app.record.detail.show', 'mobile.app.record.detail.show'], onShow);
})();
