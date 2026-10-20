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
  var NAME_LABELS = ['案件名', '現場名', '工事名'];   // ファイル名に使う項目（先に見つかったもの）
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
  function replacePlaceholders(xml, values, kind, unknown) {
    return xml.replace(PH, function (whole, inner, offset, str) {
      var label = decodeLabel(inner.replace(/<[^>]*>/g, ''));
      if (Object.prototype.hasOwnProperty.call(values, label)) {
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
    var unknown = [];
    if (!OFFICE[ext]) return Promise.resolve({ data: buf, unknown: unknown });
    return JSZipLib.loadAsync(buf).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return TARGET.test(n); });
      return Promise.all(names.map(function (n) {
        return zip.file(n).async('string').then(function (x) {
          if (n === 'xl/sharedStrings.xml') x = dropPhonetic(x);
          var y = replacePlaceholders(x, values, ext, unknown);
          // 名前が違うものに加えて、{{ }} が段落やセルをまたいでいて置き換えられなかったものも拾う
          (y.replace(/<[^>]*>/g, '').match(/\{\{[^{}]*\}\}/g) || []).forEach(function (t) {
            var l = decodeLabel(t.slice(2, -2));
            if (unknown.indexOf(l) < 0) unknown.push(l);
          });
          zip.file(n, y);
        });
      })).then(function () { return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }); })
        .then(function (data) { return { data: data, unknown: unknown }; });
    });
  }
  // 雛形に {{ }} が1つでもあるか。なければ「案件の情報は入れず、雛形のまま使う書類」として扱う
  function hasPlaceholders(JSZipLib, buf, ext) {
    if (!OFFICE[ext]) return Promise.resolve(false);
    return JSZipLib.loadAsync(buf).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return TARGET.test(n); });
      return Promise.all(names.map(function (n) {
        return zip.file(n).async('string').then(function (x) {
          if (n === 'xl/sharedStrings.xml') x = dropPhonetic(x);
          return /\{\{[^{}]+\}\}/.test(x.replace(/<[^>]*>/g, ''));
        });
      })).then(function (r) { return r.some(Boolean); });
    });
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
      module.exports = { fieldText: fieldText, replacePlaceholders: replacePlaceholders, fillTemplate: fillTemplate, hasPlaceholders: hasPlaceholders, guessTarget: guessTarget, buildFilename: buildFilename, buildValues: buildValues, quote: quote, SUB: SUB, MF: MF, today: today };
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
  function templateOf(rec) { var f = rec[MF.template] && rec[MF.template].value; return f && f[0]; }
  function extOf(name) { return ((String(name).match(/\.([^.]+)$/) || [])[1] || '').toLowerCase(); }
  var pending = {};
  function inspect(rec) {
    var tpl = templateOf(rec);
    if (!tpl) return Promise.resolve('merge');
    if (kinds[tpl.fileKey]) return Promise.resolve(kinds[tpl.fileKey]);
    if (pending[tpl.fileKey]) return pending[tpl.fileKey];
    pending[tpl.fileKey] = download(tpl.fileKey).then(function (buf) { return hasPlaceholders(JSZip, buf, extOf(tpl.name)); })
      .then(function (has) { kinds[tpl.fileKey] = has ? 'merge' : 'plain'; delete pending[tpl.fileKey]; return kinds[tpl.fileKey]; },
        function () { delete pending[tpl.fileKey]; return 'merge'; });   // 調べられないときは、従来どおり「作成して保存」にする
    return pending[tpl.fileKey];
  }
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

  function createDoc(masterRec, ctx) {
    var docName = fieldText(masterRec[MF.name]);
    var files = masterRec[MF.template].value;
    if (!files || !files.length) return Promise.reject(new Error('「' + docName + '」の雛形が書類マスタに登録されていません'));
    var tpl = files[0];
    var ext = ((tpl.name.match(/\.([^.]+)$/) || [])[1] || '').toLowerCase();

    // 保存の直前に、案件を読み直して最新の版を確かめる（同時に別の人が作ったときに、同じ版が2つできないように）
    return Promise.all([projectRecord(ctx.projectId), download(tpl.fileKey)]).then(function (res) {
      var record = res[0], cur = existingOf(record)[docName];
      var version = (cur ? cur.version : 0) + 1;
      var values = Object.assign({}, ctx.values, { '版': 'v' + version, '書類名': docName });
      return fillTemplate(JSZip, res[1], ext, values).then(function (out) {
        var filename = buildFilename(ctx.projectName, docName, version, ext);
        return upload(out.data, filename).then(function (fileKey) {
          var row = { value: {} };
          row.value[SUB.name] = { value: docName };
          row.value[SUB.stage] = { value: fieldText(masterRec[MF.stage]) };
          row.value[SUB.version] = { value: String(version) };
          row.value[SUB.at] = { value: isoToday() };
          row.value[SUB.file] = { value: [{ fileKey: fileKey }] };
          var update = {};
          update[SUB.table] = { value: rowsOf(record).map(keepRow).concat([row]) };
          // revision を付けて更新する：読んだあとに別の人が案件を更新していたら、上書きせずに失敗する
          return api('PUT', '/k/v1/record', { app: appApi.getId(), id: ctx.projectId, revision: record.$revision.value, record: update }).then(function () {
            return { filename: filename, version: version, unknown: out.unknown, data: out.data };
          }, function (e) {
            if (e && e.code === 'GAIA_CO02') throw new Error('別の人が同時に、この案件を更新しました。もう一度押してください');
            throw e;
          });
        });
      });
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
  function unknownNote(list) { return list && list.length ? '\n置き換えできずに残った項目（名前が違う、または {{ }} が段落・セルをまたいでいます）：' + list.map(function (n) { return '{{' + n + '}}'; }).join(' ') : ''; }

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
      bar.appendChild(batch);
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
      function reloadSoon(t) { setMsg(t + '（画面を更新します）', false); setTimeout(function () { location.reload(); }, 2500); }
      function shown() {
        return masters.filter(function (r) {
          if (fieldText(r[MF.stage]) !== stageSel.value) return false;
          return (!audSel.value || ((r[MF.audience] && r[MF.audience].value) || []).indexOf(audSel.value) >= 0) && matchTarget(r, tgtSel.value);
        });
      }
      // 作成に成功したら、画面上の「最新 vN」もその場で更新する（まとめ作成が途中で失敗しても、再実行で二重に作らない）
      function makeOne(rec) {
        return createDoc(rec, ctx).then(function (r) {
          existing[fieldText(rec[MF.name])] = { version: r.version, filename: r.filename, data: r.data };
          return r;
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
          btn.disabled = true; setMsg('作成しています…', false);
          makeOne(rec).then(function (r) { saveAs(r.data, r.filename); reloadSoon('保存してダウンロードしました：' + r.filename + unknownNote(r.unknown)); })
            .catch(function (e) { btn.disabled = false; setMsg((e && e.message) || '作成できませんでした', true); });
        });
        row.appendChild(left); row.appendChild(btn);
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
      batch.addEventListener('click', function () {
        var withTpl = shown().filter(hasTemplate);
        if (!withTpl.length) { setMsg('雛形が登録された書類がありません', true); return; }
        batch.disabled = true; setMsg('雛形を確認しています…', false);
        Promise.all(withTpl.map(inspect)).then(function () {
          var targets = withTpl.filter(function (r) { return kindOf(r) !== 'plain'; });
          var asIs = withTpl.length - targets.length;
          batch.disabled = false; setMsg('', false);
          if (!targets.length) { setMsg('この工程の書類は、すべて雛形のまま使うものです（各行の「雛形を開く」で取り出せます）', false); return; }
          if (!confirm(targets.length + '件の書類を作成します。作成済みの書類は次の版として作られます。' + (asIs ? '\n（雛形のまま使う ' + asIs + '件は、作成の対象外です）' : '') + '\nよろしいですか？')) return;
          batch.disabled = true;
          var done = [], failed = [], unknown = [];
          targets.reduce(function (p, rec) {
            return p.then(function () {
              setMsg('作成中…（' + (done.length + failed.length + 1) + '/' + targets.length + '）', false);
              return makeOne(rec).then(function (r) {
                done.push(r.filename);
                (r.unknown || []).forEach(function (n) { if (unknown.indexOf(n) < 0) unknown.push(n); });
              }, function (e) { failed.push(fieldText(rec[MF.name]) + '：' + ((e && e.message) || '失敗')); });
            });
          }, Promise.resolve()).then(function () {
            updateStageLabels();
            if (failed.length) { batch.disabled = false; show(); setMsg('作成 ' + done.length + '件、失敗 ' + failed.length + '件\n' + failed.join('\n') + unknownNote(unknown), true); }
            else reloadSoon(done.length + '件を保存しました' + unknownNote(unknown));
          });
        });
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

  function onShow(event) {
    if (typeof JSZip === 'undefined') { console.error('JSZipが読み込まれていません。JavaScriptの読み込み順を確認してください。'); return event; }
    var host = isMobile ? kintone.mobile.app.getHeaderSpaceElement() : kintone.app.record.getHeaderMenuSpaceElement();
    if (!host || host.querySelector('[data-doc-hub]')) return event;
    var record = event.record, projectId = recApi.getId();
    var btn = el('button', 'padding:8px 14px;font-size:14px;font-weight:bold;cursor:pointer;margin:4px 8px', '書類を作る');
    btn.type = 'button'; btn.setAttribute('data-doc-hub', '1');
    btn.addEventListener('click', function () {
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
    host.appendChild(btn);
    return event;
  }
  kintone.events.on(['app.record.detail.show', 'mobile.app.record.detail.show'], onShow);
})();
