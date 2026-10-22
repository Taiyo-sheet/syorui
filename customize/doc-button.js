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
  var SUB = { table: '書類保管', name: '書類保管_書類名', stage: '書類保管_工程', version: '書類保管_版', file: '書類保管_ファイル', at: '書類保管_作成日' };
  var ACTIVE = '使用中';
  var NAME_LABELS = ['工事名称', '案件名', '現場名', '工事名'];   // ファイル名に使う項目（先に見つかったもの）
  // 案件くんの項目のうち、「橋梁」「一般道路」など、書類マスタの「対象」と同じ言葉が入っているもの。
  // その項目名を入れると、案件を開いたときに、対象が自動で選ばれる（空のままなら、画面で手で選ぶ）
  var TARGET_LABEL = '';
  // ボタンを置く場所：案件くんの「書類関係」タブの中に置いた「スペース」項目の要素ID（scripts/add-button-space.js で足す）。
  // これがあれば、そのタブを開いたときだけボタンが見える。なければ、画面上部のヘッダーに出す
  var BUTTON_SPACE_ID = 'doc_button_space';
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
  var CHECK_LABELS = ['工事名称', '工事場所'];
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
    // 案件くんの値と、書類の値が違う
    CHECK_LABELS.forEach(function (lb) {
      var want = values && values[lb] ? String(values[lb]).trim() : '';
      if (!want) return;
      var m = t.match(new RegExp('(?:^|\\n)[\\s■●◆□・]*' + lb + '[:：]\\s*([^\\n]+)'));
      if (m && m[1].trim() && m[1].trim() !== want) out.push('「' + lb + '」が、案件くんの値と違います（書類：' + m[1].trim() + '／案件くん：' + want + '）');
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


  // Node（テスト）から使うときは、ここまでの部品だけを渡して終わる
  if (typeof kintone === 'undefined') {
    if (typeof module !== 'undefined' && module.exports) {
      module.exports = { checkText: checkText, extractText: extractText, tidyMemo: tidyMemo, BUTTON_SPACE_ID: BUTTON_SPACE_ID, fieldText: fieldText, replacePlaceholders: replacePlaceholders, fillTemplate: fillTemplate, hasPlaceholders: hasPlaceholders, guessTarget: guessTarget, buildFilename: buildFilename, buildValues: buildValues, quote: quote, SUB: SUB, MF: MF, today: today };
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
      if (n && (!map[n] || v >= map[n].version)) map[n] = { version: v, file: files && files[0] ? files[0] : null };
    });
    return map;
  }
  // 書類保管の既存の行を、そのまま送り直せる形にする（送らなかった行は消えてしまうため）
  function keepRow(row) {
    var v = {};
    [SUB.name, SUB.stage, SUB.version, SUB.at].forEach(function (c) { v[c] = { value: cellText(row, c) }; });
    v[SUB.file] = { value: ((row.value[SUB.file] && row.value[SUB.file].value) || []).map(function (f) { return { fileKey: f.fileKey }; }) };
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
    var close = el('button', 'padding:6px 12px;font-size:14px;cursor:pointer', '閉じる'); close.type = 'button';
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
      function askValues(asks, name) {
        return new Promise(function (resolve) {
          var cover = el('div', 'position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:10001;overflow:auto');
          var card = el('div', 'background:#fff;max-width:560px;margin:40px auto;border-radius:8px;padding:16px;font-size:14px;box-sizing:border-box;width:calc(100% - 24px)');
          card.appendChild(el('div', 'font-weight:bold;font-size:16px;margin-bottom:10px', '「' + name + '」の内容'));
          var inputs = {};
          asks.forEach(function (a) {
            card.appendChild(el('div', 'font-size:12px;color:#666;margin-top:8px', a.label));
            var inp = a.type === 'memo' ? el('textarea', 'width:100%;height:160px;box-sizing:border-box;font-size:14px;padding:6px') : el('input', 'width:100%;box-sizing:border-box;font-size:14px;padding:6px');
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
              var t = inputs[a.key].value;
              if (a.type === 'date') { var m = t.match(/^(\d{4})-(\d{2})-(\d{2})$/); t = m ? m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日' : t; }
              v[a.key] = a.type === 'memo' ? tidyMemo(t) : t.trim();
              if (!v[a.key]) v[a.key] = askDefault(a);
            });
            close(v);
          });
          bar.appendChild(cancel); bar.appendChild(ok); card.appendChild(bar); cover.appendChild(card); document.body.appendChild(cover);
        });
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
            var cur = existing[name];
            (cur.data ? Promise.resolve(cur.data) : download(cur.file.fileKey)).then(function (buf) {
              return extractText(JSZip, buf, extOf(cur.filename || (cur.file && cur.file.name) || ''));
            }).then(function (text) {
              var res = checkText(text, ctx.values);
              chk.disabled = false;
              setMsg(res.length ? '「' + name + '」のチェック結果（' + res.length + '件）：\n・' + res.join('\n・') : '「' + name + '」：問題は見つかりませんでした', res.length > 0);
            }).catch(function (e) { chk.disabled = false; setMsg((e && e.message) || 'チェックできませんでした', true); });
          });
          right.appendChild(chk); right.appendChild(rev); right.appendChild(pick);
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

  // 画面の右下に浮かぶ丸いボタン。マウスを近づけると（スマートフォンでは、押すと）、文字つきに広がる。
  // ボタン用スペース（書類関係タブの中）があれば、そのタブを開いている間だけ見える。なければ、いつでも見える
  var FAB_SIZE = 52;
  function floatingButton(label, space, onClick) {
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
    // 書類関係タブを開いている間だけ見せる（タブを閉じると、スペースが隠れる）
    if (space) {
      var sync = function () { fab.style.display = space.offsetParent !== null ? 'flex' : 'none'; };
      sync(); var timer = setInterval(function () { if (!document.body.contains(fab)) { clearInterval(timer); return; } sync(); }, 400);
    }
    return fab;
  }

  function onShow(event) {
    if (typeof JSZip === 'undefined') { console.error('JSZipが読み込まれていません。JavaScriptの読み込み順を確認してください。'); return event; }
    var space = null;
    try { space = recApi.getSpaceElement(BUTTON_SPACE_ID); } catch (e) { space = null; }
    var record = event.record, projectId = recApi.getId();
    floatingButton('書類を作る', space, function () {
      loadLabelMap(record).then(function (lm) {
        var values = buildValues(record, lm);
        var projectName = '';
        NAME_LABELS.some(function (l) { if (values[l]) { projectName = values[l]; return true; } return false; });
        openPanel({
          values: values, projectId: projectId, projectName: projectName || ('案件' + projectId),
          currentStage: values[STAGE_LABEL] || '', target: TARGET_LABEL ? values[TARGET_LABEL] || '' : '', guessedTarget: guessTarget(projectName)
        });
      });
    });
    return event;
  }
  kintone.events.on(['app.record.detail.show', 'mobile.app.record.detail.show'], onShow);
})();
