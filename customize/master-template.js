/* ============================================================
   書類マスタ：「雛形を差し替える」ボタン
   ・書類マスタの画面で、雛形のファイル（Word・Excel など）を選ぶだけで、差し替えられる
     （Claude に頼まなくても、あなたが、いつでも直せる）
   ・古い雛形は、消さずに、「旧版_日付_」を付けて、同じ「雛形」欄に残す（先頭が、使われる雛形）。最新 KEEP 個まで
   ・案件くんは、書類を作るたびに、この「雛形」欄の先頭のファイルを読むので、差し替えると、次の書類作成から、新しい雛形になる

   使い方
     ・レコードの詳細画面：ボタン「雛形を差し替える」→ ファイルを選ぶ → 確認 → 差し替え
     ・一覧画面：ボタン「雛形をまとめて差し替える」→ 複数のファイルを選ぶ
         ファイル名（拡張子をのぞく）が、書類名と同じものを、その書類の雛形にする（例：工事仕様書.docx → 書類名「工事仕様書」）
         書類名と合わないファイルは、差し替えず、画面に出して、知らせる

   設置：書類マスタ（879）のPC用JavaScript（JSZip のあと）。設置は、利用者の許可をもらってから
   ============================================================ */
// 「試しに作る」は、案件くんの書類作成と同じ部品（doc-button.js）を借りる。借りるだけにするための合図（doc-button.js より先に読み込む）
window.__TS_DOC_ENGINE_ONLY = true;
(function () {
  'use strict';

  var KEEP = 3;
  var NAME_FIELD = '書類名', FILE_FIELD = '雛形';
  var ACCEPT = '.docx,.xlsx,.pptx,.pdf';
  var ANKEN_APP = 575;   // 案件くんのアプリ番号（項目名の照合と、「試しに作る」の見本の案件を読むために使う。読み取りだけ）
  var AUTO = ['書類名', '版', '作成日', '打合せ日', '作成者', '議事録本文', '議事録メモ', '場所', '出席者', '議題', '決定事項', 'ToDo', '次回予定'];   // 案件くんの項目ではなく、作成のときに入る言葉

  function el(tag, style, text) { var e = document.createElement(tag); if (style) e.style.cssText = style; if (text) e.textContent = text; return e; }
  function api(method, path, params) { return kintone.api(kintone.api.url(path, true), method, params); }
  function stamp() { var d = new Date(); function z(n) { return (n < 10 ? '0' : '') + n; } return d.getFullYear() + z(d.getMonth() + 1) + z(d.getDate()); }
  function baseName(n) { return String(n).replace(/\.[^.]+$/, ''); }
  function extOf(n) { return ((String(n).match(/\.([^.]+)$/) || [])[1] || '').toLowerCase(); }

  function download(fileKey) {
    return new Promise(function (resolve, reject) {
      var xhr = new XMLHttpRequest();
      xhr.open('GET', kintone.api.url('/k/v1/file', true) + '?fileKey=' + encodeURIComponent(fileKey));
      xhr.setRequestHeader('X-Requested-With', 'XMLHttpRequest');
      xhr.responseType = 'arraybuffer';
      xhr.onload = function () { xhr.status === 200 ? resolve(xhr.response) : reject(new Error('いまの雛形を取り出せませんでした（' + xhr.status + '）')); };
      xhr.onerror = function () { reject(new Error('いまの雛形を取り出せませんでした')); };
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
      xhr.onload = function () { xhr.status === 200 ? resolve(JSON.parse(xhr.responseText).fileKey) : reject(new Error('ファイルをアップロードできませんでした（' + xhr.status + '）')); };
      xhr.onerror = function () { reject(new Error('ファイルをアップロードできませんでした')); };
      xhr.send(fd);
    });
  }
  function readFile(file) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(r.result); };
      r.onerror = function () { reject(new Error('ファイルを読み込めませんでした：' + file.name)); };
      r.readAsArrayBuffer(file);
    });
  }

  // 新しい雛形の {{ }} の一覧（Word・Excel・PowerPoint）。JSZip が無い・読めないときは、null（確認画面に出さない）
  function placeholders(buf, ext) {
    if (!window.JSZip || ['docx', 'xlsx', 'pptx'].indexOf(ext) < 0) return Promise.resolve(null);
    return window.JSZip.loadAsync(buf).then(function (zip) {
      var names = Object.keys(zip.files).filter(function (n) { return /^(word|xl|ppt)\/.*\.xml$/.test(n) && !/\/(styles|theme|settings|fontTable)/.test(n); });
      return Promise.all(names.map(function (n) { return zip.file(n).async('string'); })).then(function (xs) {
        var found = {};
        xs.forEach(function (x) { (x.replace(/<[^>]*>/g, '').match(/\{\{[^{}]+\}\}/g) || []).forEach(function (p) { found[p] = 1; }); });
        return Object.keys(found);
      });
    }).catch(function () { return null; });
  }

  // 1件の差し替え：古い雛形を「旧版」で残し、新しいものを先頭にして、レコードに保存する
  function swap(rec, file) {
    var id = rec.$id.value, old = (rec[FILE_FIELD] && rec[FILE_FIELD].value) || [];
    return readFile(file).then(function (buf) {
      var keep = [];
      return old.reduce(function (p, f) {
        return p.then(function () {
          // すべて取り出して、上げ直す：保存済みのファイルは、新しく上げたファイルより先に並ぶため、上げ直さないと、新しい雛形が先頭に来ない
          var name = /^旧版_/.test(f.name) ? f.name : '旧版_' + stamp() + '_' + f.name;
          return download(f.fileKey).then(function (d) { return upload(d, name); }).then(function (k) { keep.push({ fileKey: k }); });
        });
      }, Promise.resolve()).then(function () { return upload(buf, file.name); }).then(function (newKey) {
        var next = [{ fileKey: newKey }].concat(keep.slice(0, KEEP));
        var body = { app: appId(), id: id, record: {} };
        body.record[FILE_FIELD] = { value: next };
        if (rec.$revision) body.revision = rec.$revision.value;
        return api('PUT', '/k/v1/record', body);
      });
    });
  }

  // 案件くんの項目名の一覧（読めなければ null）。雛形の {{ }} の打ち間違いを見つけるために使う
  var labelsP = null;
  function ankenLabels() {
    if (!labelsP) labelsP = api('GET', '/k/v1/app/form/fields', { app: ANKEN_APP }).then(function (r) {
      var set = {}; Object.keys(r.properties || {}).forEach(function (c) { var l = r.properties[c].label; if (l) set[String(l).trim()] = 1; }); return set;
    }).catch(function () { return null; });
    return labelsP;
  }
  // {{ }} のうち、案件くんにも、自動で入る言葉にも、ない名前
  function unknownNames(ph) {
    if (!ph) return Promise.resolve(null);
    return ankenLabels().then(function (set) {
      if (!set) return null;
      return ph.map(function (x) { return x.slice(2, -2); }).filter(function (n) { return !set[n.trim()] && AUTO.indexOf(n.trim()) < 0; });
    });
  }
  function warn(bad) { return bad && bad.length ? '\n【注意】案件くんにない項目名：' + bad.map(function (n) { return '{{' + n + '}}'; }).join(' ') + '\n　　このままだと、書類の、その場所が、空になります（項目名の打ち間違いが、ないか確認してください）\n' : ''; }

  function summary(file, ph) {
    var s = '「' + file.name + '」を、雛形にします。\n';
    if (ph) s += ph.length ? '差し込む項目：' + ph.join(' ') + '\n' : '差し込む項目（{{ }}）は、ありません。雛形のまま使われます。\n';
    return s;
  }

  function pickFiles(multiple) {
    return new Promise(function (resolve) {
      var inp = document.createElement('input'); inp.type = 'file'; inp.accept = ACCEPT; inp.multiple = !!multiple; inp.style.display = 'none';
      inp.setAttribute('data-master-file', '1');
      inp.addEventListener('change', function () { resolve(Array.prototype.slice.call(inp.files || [])); if (inp.parentNode) inp.parentNode.removeChild(inp); });
      document.body.appendChild(inp);
      inp.click();
    });
  }

  // PC・スマホ（モバイル）の、どちらの画面でも動かす
  var isMobile = /^\/k\/(guest\/\d+\/)?m\//.test(location.pathname);
  function appId() { return isMobile ? kintone.mobile.app.getId() : kintone.app.getId(); }
  // ボタンを置く場所：画面の上の、ボタン置き場。見つからないときは、画面の上に、帯を出す（ボタンが、出ないことが、ないように）
  function hostFor(kind) {
    var h = null;
    try {
      if (isMobile) h = kintone.mobile.app.getHeaderSpaceElement && kintone.mobile.app.getHeaderSpaceElement();
      else if (kind === 'detail') h = kintone.app.record.getHeaderMenuSpaceElement && kintone.app.record.getHeaderMenuSpaceElement();
      else h = kintone.app.getHeaderMenuSpaceElement && kintone.app.getHeaderMenuSpaceElement();
    } catch (e) { h = null; }
    if (h) return h;
    var bar = document.querySelector('[data-master-bar]');
    if (!bar) { bar = el('div', 'position:fixed;left:8px;right:8px;top:8px;z-index:9000;background:#fff;border:1px solid #ccc;border-radius:6px;padding:6px 8px;box-shadow:0 2px 8px rgba(0,0,0,.2)'); bar.setAttribute('data-master-bar', '1'); document.body.appendChild(bar); }
    return bar;
  }
  function guard(fn) { return function (event) { try { return fn(event) || event; } catch (e) { console.error('[master-template]', e); return event; } }; }

  // キントーン標準のボタン（kintoneplugin-button-normal）に近い見た目。primary は、青
  function button(text, attr, onClick, primary) {
    var base = 'display:inline-block;box-sizing:border-box;margin:' + (isMobile ? '4px 8px 4px 0' : '0 8px 0 0') + ';padding:0 ' + (isMobile ? '16px' : '18px') + ';height:' + (isMobile ? '44px' : '40px') + ';line-height:' + (isMobile ? '42px' : '38px') + ';font-size:' + (isMobile ? '15px' : '14px') + ';border-radius:3px;cursor:pointer;vertical-align:middle;font-family:inherit;white-space:nowrap;';
    var look = primary ? 'border:1px solid #2980b9;background:#3498db;color:#fff;box-shadow:1px 1px 1px rgba(255,255,255,.4) inset;' : 'border:1px solid #e3e7e8;background:#f7f9fa;color:#3498db;box-shadow:1px 1px 1px #fff inset;';
    var b = el('button', base + look, text);
    b.type = 'button'; b.setAttribute(attr, '1'); b.addEventListener('click', onClick);
    b.addEventListener('mouseover', function () { b.style.opacity = '.85'; }); b.addEventListener('mouseout', function () { b.style.opacity = '1'; });
    return b;
  }
  function note(host) { var n = el('span', 'font-size:12px;margin-left:4px;color:#1f7a3d'); n.setAttribute('data-master-note', '1'); host.appendChild(n); return n; }

  // ---- 詳細画面：この書類の雛形を差し替える ----
  kintone.events.on(['app.record.detail.show', 'mobile.app.record.detail.show'], guard(function (event) {
    var host = hostFor('detail');
    if (!host || host.querySelector('[data-master-swap]')) return event;
    var rec = event.record, name = rec[NAME_FIELD] && rec[NAME_FIELD].value;
    var msg = null;
    host.appendChild(button('雛形を差し替える', 'data-master-swap', function () {
      msg.style.color = '#1f7a3d'; msg.textContent = '';
      pickFiles(false).then(function (files) {
        if (!files.length) return;
        var f = files[0];
        return readFile(f).then(function (buf) { return placeholders(buf, extOf(f.name)); }).then(function (ph) { return unknownNames(ph).then(function (bad) { return [ph, bad]; }); }).then(function (pb) {
          var ph = pb[0], bad = pb[1];
          var olds = ((rec[FILE_FIELD] && rec[FILE_FIELD].value) || []).map(function (x) { return x.name; });
          if (!confirm('書類「' + name + '」の雛形を差し替えます。\n\n' + summary(f, ph) + warn(bad) + '\nいまの雛形：' + (olds.length ? olds.join('、') : 'なし') + '\n（古いものは「旧版_' + stamp() + '_」を付けて残します。最新' + KEEP + '個まで）\n\nよろしいですか？')) return;
          msg.textContent = '差し替えています…';
          return swap(rec, f).then(function () { msg.textContent = '差し替えました。画面を読み込み直します'; setTimeout(function () { location.reload(); }, 800); });
        });
      }).catch(function (e) { msg.style.color = '#c0392b'; msg.textContent = (e && e.message) || '差し替えに失敗しました（元の雛形は、そのままです）'; });
    }, true));
    host.appendChild(button('試しに作る', 'data-master-try', function () { tryPanel(rec, name, msg); }));
    host.appendChild(button('履歴・前の版に戻す', 'data-master-hist', function () { histPanel(rec, name, msg); }));
    host.appendChild(button('{{ }}の候補', 'data-master-cand', function () { candPanel(); }));
    msg = note(host);
    return event;
  }));

  // ---- 試しに作る：見本の案件を選ぶと、いまの雛形で、記入済みの書類を、その場で作って、点検する（どこにも、保存しない） ----
  function tryPanel(rec, name, msg) {
    var E = window.TSDocEngine;
    var files = (rec[FILE_FIELD] && rec[FILE_FIELD].value) || [];
    if (!E) { msg.style.color = '#c0392b'; msg.textContent = '「試しに作る」には、書類マスタのJavaScriptに doc-button.js が必要です（master-template.js の後ろに登録）'; return; }
    if (!files.length) { msg.style.color = '#c0392b'; msg.textContent = 'この書類には、雛形がありません'; return; }
    var cover = el('div', 'position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:10001;overflow:auto');
    cover.setAttribute('data-try-panel', '1');
    var card = el('div', 'background:#fff;max-width:620px;margin:24px auto;border-radius:8px;padding:16px;font-size:14px;box-sizing:border-box;width:calc(100% - 24px)');
    card.appendChild(el('div', 'font-weight:bold;font-size:16px;margin-bottom:6px', '「' + name + '」を、試しに作る'));
    card.appendChild(el('div', 'font-size:12px;color:#666;margin-bottom:8px', '見本にする案件を選んでください（工事名・場所などを入れるために使います）。議事録の本文など、画面で入力する項目は、見本の文章が入ります。いまの雛形（' + files[0].name + '）で、書類を作って、点検します。案件くんにも、ここにも、保存しません（案件くんは、読むだけです）。'));
    var search = el('input', 'width:100%;box-sizing:border-box;font-size:14px;padding:6px;margin-bottom:6px'); search.type = 'text'; search.placeholder = '案件名の一部で、絞り込む（空なら、新しい順）'; search.setAttribute('data-try-search', '1');
    var list = el('div', 'max-height:220px;overflow:auto;border:1px solid #ddd;margin-bottom:8px');
    var result = el('div', 'margin-top:8px');
    var bar = el('div', 'margin-top:10px;display:flex;justify-content:flex-end'); var close = el('button', 'padding:6px 12px;cursor:pointer', '閉じる'); close.type = 'button'; bar.appendChild(close);
    close.addEventListener('click', function () { document.body.removeChild(cover); });
    card.appendChild(search); card.appendChild(list); card.appendChild(result); card.appendChild(bar); cover.appendChild(card); document.body.appendChild(cover);

    var lm = null;
    function loadLm() {
      if (lm) return Promise.resolve(lm);
      return api('GET', '/k/v1/app/form/fields', { app: ANKEN_APP }).then(function (r) {
        var map = {}, props = {};
        Object.keys(r.properties || {}).forEach(function (c) { var pr = r.properties[c]; props[c] = pr; if (pr.label) map[String(pr.label).trim()] = c; });
        lm = { map: map, props: props }; return lm;
      });
    }
    function title(r, l) { var t = ''; E.NAME_LABELS.some(function (x) { var c = l.map[x]; if (c && r[c] && r[c].value) { t = r[c].value; return true; } return false; }); return t || ('案件 ' + r.$id.value); }
    function loadList() {
      list.textContent = '読み込んでいます…';
      return loadLm().then(function (l) {
        var code = null; E.NAME_LABELS.some(function (x) { if (l.map[x]) { code = l.map[x]; return true; } return false; });
        var q = (search.value.trim() && code ? code + ' like "' + search.value.trim().replace(/["\\]/g, '') + '" ' : '') + 'order by $id desc limit 15';
        return api('GET', '/k/v1/records', { app: ANKEN_APP, query: q }).then(function (res) {
          list.textContent = '';
          if (!res.records.length) { list.textContent = '該当する案件が、ありません'; return; }
          res.records.forEach(function (r) {
            var row = el('div', 'padding:6px 8px;border-bottom:1px solid #eee;cursor:pointer', title(r, l)); row.setAttribute('data-try-case', r.$id.value);
            row.addEventListener('click', function () { run(r, l); });
            list.appendChild(row);
          });
        });
      }).catch(function (e) { list.textContent = '案件くんを読み込めませんでした：' + ((e && e.message) || ''); });
    }
    var timer = null; search.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(loadList, 400); });

    function run(record, l) {
      result.textContent = '作っています…'; result.style.color = '#333';
      var ext = extOf(files[0].name);
      var values = E.buildValues(record, l);
      // 書類作成のとき、画面で入力する項目（議事録の本文など）は、見本の文章を入れて、できあがりの見た目を確認できるようにする。見本がない項目は、目印の言葉
      var d = new Date(), login = ''; try { login = (kintone.getLoginUser() || {}).name || ''; } catch (e) { login = ''; }
      var sample = { 打合せ日: d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日', 作成者: login || '（作成者）', 出席者: '太田（TAIYOSHEET）、國井、藤原（元請）', 場所: '現場事務所', 議題: '・夜間規制の時間について\n・資材の搬入日の調整', 議事録メモ: '・規制時間は、9時から17時で調整する方向\n・資材は金曜午前に搬入したい（元請の了承待ち）', 決定事項: '・規制は9時から17時とする\n・資材の搬入は、金曜午前', ToDo: '・10/25｜太田｜見積書を提出\n・10/27｜國井｜工程表を共有', 次回予定: '11月1日（金）　現場事務所' };
      E.ASK.forEach(function (a) { values[a.key] = sample[a.key] || '【入力欄：' + a.key + '】'; });
      if (E.formatMinutes) values['議事録本文'] = E.formatMinutes(sample);
      values['版'] = 'v0（試し）'; values['書類名'] = name;
      download(files[0].fileKey).then(function (buf) { return E.fillTemplate(window.JSZip, buf, ext, values); }).then(function (out) {
        return E.extractText(window.JSZip, out.data, ext).then(function (text) {
          var hasUnknown = !!(out.unknown && out.unknown.length);
          var issues = E.checkText(text, values).filter(function (x) { return x.indexOf('【入力欄') < 0 && !(hasUnknown && x.indexOf('置き換えられていない項目が残っています') === 0); });   // 項目名の打ち間違いは、下の1件にまとめる
          if (out.unknown && out.unknown.length) issues.unshift('案件くんにない項目名（空になります）：' + out.unknown.map(function (n) { return /^\{\{/.test(n) ? n : '{{' + n + '}}'; }).join(' '));
          if (out.empty && out.empty.length) issues.push('案件くんの値が空の項目：' + out.empty.join('、'));
          result.textContent = '';
          result.appendChild(el('div', 'font-weight:bold;color:' + (issues.length ? '#c0392b' : '#1f7a3d'), issues.length ? '点検の結果：気になるところが、' + issues.length + '件あります' : '点検の結果：問題は、見つかりませんでした'));
          issues.forEach(function (t) { result.appendChild(el('div', 'font-size:13px;margin:2px 0 0 8px', '・' + t)); });
          result.appendChild(el('div', 'font-size:12px;color:#666;margin-top:4px', '※「【入力欄：…】」は、書類作成のとき、画面で入力する項目の場所です。'));
          // できあがりの中身を、文字で見せる（スマホのキントーンのアプリでは、Wordなどのファイルを開けないため）
          var pv = el('pre', 'white-space:pre-wrap;word-break:break-all;max-height:260px;overflow:auto;background:#f7f7f7;border:1px solid #ddd;border-radius:4px;padding:8px;margin:8px 0 0;font-size:13px;line-height:1.5');
          pv.setAttribute('data-try-preview', '1'); pv.textContent = String(text).length > 20000 ? String(text).slice(0, 20000) + '\n…（以下、省略）' : String(text);
          result.appendChild(el('div', 'font-size:12px;color:#666;margin-top:8px', '試しに作った書類の中身（文字だけ。表や体裁は、ここには出ません）'));
          result.appendChild(pv);
          if (!isMobile) {   // PCでは、ファイルとして開ける（保存は、されない）
            var fname = '試し_' + title(record, l) + '_' + name + '.' + ext;
            var dl = el('button', 'margin-top:8px;padding:6px 12px;cursor:pointer', '試しに作ったファイルを、開く（保存はされません）'); dl.type = 'button'; dl.setAttribute('data-try-download', '1');
            dl.addEventListener('click', function () {
              var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([out.data])); a.download = fname; document.body.appendChild(a); a.click(); document.body.removeChild(a);
            });
            result.appendChild(dl);
          }
        });
      }).catch(function (e) { result.style.color = '#c0392b'; result.textContent = '試しに作れませんでした：' + ((e && e.message) || ''); });
    }
    loadList();
  }

  // ---- 共通：画面いっぱいの、小さな窓 ----
  function modal(titleText, noteText) {
    var cover = el('div', 'position:fixed;left:0;top:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:10001;overflow:auto');
    var card = el('div', 'background:#fff;max-width:640px;margin:24px auto;border-radius:8px;padding:16px;font-size:14px;box-sizing:border-box;width:calc(100% - 24px)');
    card.appendChild(el('div', 'font-weight:bold;font-size:16px;margin-bottom:6px', titleText));
    if (noteText) card.appendChild(el('div', 'font-size:12px;color:#666;margin-bottom:8px', noteText));
    var body = el('div'); card.appendChild(body);
    var bar = el('div', 'margin-top:10px;display:flex;justify-content:flex-end'); var close = el('button', 'padding:' + (isMobile ? '10px 16px' : '6px 12px') + ';cursor:pointer', '閉じる'); close.type = 'button'; bar.appendChild(close); card.appendChild(bar);
    close.addEventListener('click', function () { document.body.removeChild(cover); });
    cover.appendChild(card); document.body.appendChild(cover);
    return { cover: cover, card: card, body: body };
  }
  function copyText(t) {
    return new Promise(function (resolve) {
      var fallback = function () { var x = document.createElement('textarea'); x.value = t; x.style.cssText = 'position:fixed;left:0;top:0;opacity:0'; document.body.appendChild(x); x.select(); var ok = false; try { ok = document.execCommand('copy'); } catch (e) { ok = false; } document.body.removeChild(x); resolve(ok); };
      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t).then(function () { resolve(true); }, fallback); else fallback();
    });
  }

  // ---- 雛形に書く {{ }} の候補：案件くんの項目名の一覧（タップでコピー）と、見本の値 ----
  var AUTO_HELP = { '書類名': 'この書類の名前', '版': 'v1, v2…（作るたびに増える）', '作成日': '作った日', '打合せ日': '議事録：打合せの日（画面で入力）', '作成者': '議事録：作った人（自動）', '議事録本文': '議事録：見出し付きの本文（画面で入力）', '出席者': '議事録：画面で入力', '場所': '議事録：画面で入力', '議題': '議事録：画面で入力', '議事録メモ': '議事録：画面で入力', '決定事項': '議事録：画面で入力', 'ToDo': '議事録：画面で入力', '次回予定': '議事録：画面で入力' };
  var SKIP_TYPES = { LABEL: 1, SPACER: 1, HR: 1, SUBTABLE: 1, FILE: 1, GROUP: 1, REFERENCE_TABLE: 1, RECORD_NUMBER: 0, CATEGORY: 1, STATUS: 1, STATUS_ASSIGNEE: 1, CREATOR: 0, MODIFIER: 0 };
  function candPanel() {
    var m = modal('雛形に書く {{ }} の候補', '雛形の中の、案件の値を入れたい場所に、下の「{{ }}」を、そのまま書きます。項目をタップすると、コピーされます。見本の値は、案件くんの、いちばん新しい案件のものです。');
    var search = el('input', 'width:100%;box-sizing:border-box;font-size:14px;padding:6px;margin-bottom:6px'); search.type = 'text'; search.placeholder = '項目名で、絞り込む'; search.setAttribute('data-cand-search', '1');
    var note = el('div', 'font-size:12px;color:#1f7a3d;min-height:16px;margin-bottom:4px'); note.setAttribute('data-cand-note', '1');
    var list = el('div', 'max-height:' + (isMobile ? '55vh' : '380px') + ';overflow:auto;border:1px solid #ddd');
    m.body.appendChild(search); m.body.appendChild(note); m.body.appendChild(list); list.textContent = '読み込んでいます…';
    var items = [];
    var E = window.TSDocEngine;
    Promise.all([
      api('GET', '/k/v1/app/form/fields', { app: ANKEN_APP }),
      api('GET', '/k/v1/records', { app: ANKEN_APP, query: 'order by $id desc limit 1' }).catch(function () { return { records: [] }; })
    ]).then(function (rs) {
      var props = rs[0].properties || {}, rec = (rs[1].records || [])[0] || null, lm = { map: {}, props: {} };
      Object.keys(props).forEach(function (c) { var pr = props[c]; lm.props[c] = pr; if (pr.label) lm.map[String(pr.label).trim()] = c; });
      var vals = rec && E ? E.buildValues(rec, lm) : {};
      Object.keys(props).forEach(function (c) {
        var pr = props[c]; if (!pr.label || SKIP_TYPES[pr.type]) return;
        items.push({ name: String(pr.label).trim(), sample: vals[String(pr.label).trim()] || '', kind: '案件くんの項目' });
      });
      Object.keys(AUTO_HELP).forEach(function (k) { items.push({ name: k, sample: AUTO_HELP[k], kind: '自動で入る言葉' }); });
      render('');
    }).catch(function (e) { list.textContent = '案件くんの項目を読み込めませんでした：' + ((e && e.message) || ''); });
    function render(word) {
      list.textContent = '';
      var shown = items.filter(function (x) { return !word || x.name.indexOf(word) >= 0; });
      if (!shown.length) { list.textContent = '該当する項目が、ありません'; return; }
      var last = '';
      shown.forEach(function (x) {
        if (x.kind !== last) { last = x.kind; list.appendChild(el('div', 'background:#eef3f9;padding:4px 8px;font-size:12px;font-weight:bold', x.kind)); }
        var row = el('div', 'padding:6px 8px;border-bottom:1px solid #eee;cursor:pointer'); row.setAttribute('data-cand-item', x.name);
        row.appendChild(el('div', 'font-family:monospace;font-size:14px', '{{' + x.name + '}}'));
        if (x.sample) row.appendChild(el('div', 'font-size:12px;color:#666;overflow:hidden;text-overflow:ellipsis;white-space:nowrap', (x.kind === '案件くんの項目' ? '見本：' : '') + x.sample));
        row.addEventListener('click', function () { var t = '{{' + x.name + '}}'; copyText(t).then(function (ok) { note.style.color = ok ? '#1f7a3d' : '#c0392b'; note.textContent = ok ? 'コピーしました：' + t : 'コピーできませんでした。上の文字を、選んで、コピーしてください：' + t; }); });
        list.appendChild(row);
      });
    }
    search.addEventListener('input', function () { render(search.value.trim()); });
  }

  // ---- 雛形の履歴と、前の版に戻す（旧版の日付は、ファイル名の「旧版_日付_」から読む） ----
  function histPanel(rec, name, msg) {
    var files = (rec[FILE_FIELD] && rec[FILE_FIELD].value) || [];
    var m = modal('「' + name + '」の雛形の履歴', '先頭が、いま使われている雛形です。「この版に戻す」を押すと、その版が先頭になり、いまの雛形は、旧版として残ります（消えません）。');
    if (!files.length) { m.body.textContent = 'この書類には、雛形が、ありません'; return; }
    files.forEach(function (f, i) {
      var d = (f.name.match(/^旧版_(\d{4})(\d{2})(\d{2})_/) || []);
      var row = el('div', 'padding:8px;border-bottom:1px solid #eee;display:flex;gap:8px;align-items:center;flex-wrap:wrap');
      row.appendChild(el('div', 'flex:1;min-width:200px', (i === 0 ? '【いま使われている】 ' : '') + f.name.replace(/^旧版_\d{8}_/, '') + (d.length ? '（' + d[1] + '/' + d[2] + '/' + d[3] + ' に、旧版になりました）' : '')));
      if (i > 0) {
        var b = el('button', 'padding:' + (isMobile ? '10px 14px' : '5px 10px') + ';cursor:pointer', 'この版に戻す'); b.type = 'button'; b.setAttribute('data-hist-restore', String(i));
        b.addEventListener('click', function () {
          if (!confirm('「' + f.name.replace(/^旧版_\d{8}_/, '') + '」を、いまの雛形に戻します。\nいまの雛形「' + files[0].name + '」は、旧版として残ります。\n\nよろしいですか？')) return;
          b.disabled = true; b.textContent = '戻しています…';
          restore(rec, i).then(function () { m.body.textContent = '戻しました。画面を読み込み直します'; setTimeout(function () { location.reload(); }, 800); },
            function (e) { b.disabled = false; b.textContent = 'この版に戻す'; m.body.insertBefore(el('div', 'color:#c0392b;margin-bottom:6px', (e && e.message) || '戻せませんでした（雛形は、そのままです）'), m.body.firstChild); });
        });
        row.appendChild(b);
      }
      m.body.appendChild(row);
    });
  }
  // 戻す：選んだ版を先頭に、いまの先頭を旧版にして、すべて上げ直す（保存済みのファイルは、新しく上げたファイルより先に並ぶため）
  function restore(rec, index) {
    var files = ((rec[FILE_FIELD] && rec[FILE_FIELD].value) || []).slice();
    var chosen = files[index], head = files[0];
    var order = [{ f: chosen, name: chosen.name.replace(/^旧版_\d{8}_/, '') }, { f: head, name: /^旧版_/.test(head.name) ? head.name : '旧版_' + stamp() + '_' + head.name }];
    files.forEach(function (f, i) { if (i !== 0 && i !== index) order.push({ f: f, name: f.name }); });
    var keys = [];
    return order.reduce(function (p, o) {
      return p.then(function () { return download(o.f.fileKey).then(function (d) { return upload(d, o.name); }).then(function (k) { keys.push({ fileKey: k }); }); });
    }, Promise.resolve()).then(function () {
      var body = { app: appId(), id: rec.$id.value, record: {} }; body.record[FILE_FIELD] = { value: keys.slice(0, KEEP + 1) };
      if (rec.$revision) body.revision = rec.$revision.value;
      return api('PUT', '/k/v1/record', body);
    });
  }

  // ---- 一覧画面：「雛形の状況」（どの書類に雛形があるか・案件で何回作られたか・{{ }}の打ち間違いはないか） ----
  function wideHost() {
    try {
      if (isMobile) return kintone.mobile.app.getHeaderSpaceElement && kintone.mobile.app.getHeaderSpaceElement();
      return kintone.app.getHeaderSpaceElement && kintone.app.getHeaderSpaceElement();
    } catch (e) { return null; }
  }
  // 案件くんの「書類保管」表から、書類名ごとの、作成数と、最終作成日を数える（読むだけ。最新の案件から、最大500件）
  function usageCounts() {
    return api('GET', '/k/v1/records', { app: ANKEN_APP, fields: ['書類保管'], query: 'order by $id desc limit 500' }).then(function (res) {
      var u = {};
      res.records.forEach(function (r) {
        var rows = (r['書類保管'] && r['書類保管'].value) || [];
        rows.forEach(function (row) {
          var n = row.value['書類保管_書類名'] && row.value['書類保管_書類名'].value; if (!n) return;
          var at = (row.value['書類保管_作成日'] && row.value['書類保管_作成日'].value) || '';
          var o = u[n] || (u[n] = { count: 0, last: '' }); o.count++; if (at > o.last) o.last = at;
        });
      });
      return { usage: u, cases: res.records.length };
    }).catch(function () { return null; });
  }
  function dashboard(host) {
    if (!host || host.querySelector('[data-master-dash]')) return;
    var box = el('details', 'margin:8px 0;border:1px solid #cfd8e3;border-radius:6px;background:#fff;padding:6px 10px');
    box.setAttribute('data-master-dash', '1');
    var sum = el('summary', 'cursor:pointer;font-weight:bold;font-size:' + (isMobile ? '15px' : '13px'), '雛形の状況（立場で、絞り込めます）'); box.appendChild(sum);
    var body = el('div', 'margin-top:8px;font-size:12px'); box.appendChild(body); host.appendChild(box);
    var loaded = false;
    var load = function () {
      if (!box.open || loaded) return; loaded = true;
      body.textContent = '読み込んでいます…';
      Promise.all([
        api('GET', '/k/v1/records', { app: appId(), query: 'order by 表示順 asc limit 500' }),
        usageCounts()
      ]).then(function (rs) {
        var recs = rs[0].records, uc = rs[1];
        body.textContent = '';
        var state = { aud: '', old: false, withFile: false };
        var AUDS = [['', 'すべて'], ['施工店', '施工店用'], ['工事部', '工事部用'], ['施工指導', '施工指導用'], ['外部パートナー', '外部パートナー用'], ['TS計画推進', 'TS計画推進']];
        var chipBar = el('div', 'display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px;align-items:center');
        var info = el('div', 'margin-bottom:6px;font-weight:bold');
        var tableHost = el('div');
        var cells = [];
        var chipEls = [];
        var chip = function (label, on, fn) {
          var b = el('button', 'padding:' + (isMobile ? '9px 12px' : '5px 10px') + ';font-size:' + (isMobile ? '14px' : '12px') + ';cursor:pointer;border-radius:16px;border:1px solid #3498db;' + (on ? 'background:#3498db;color:#fff' : 'background:#fff;color:#3498db'), label);
          b.type = 'button'; b.addEventListener('click', fn); return b;
        };
        function drawChips() {
          chipBar.textContent = ''; chipEls = [];
          AUDS.forEach(function (a) { var b = chip(a[1], state.aud === a[0], function () { state.aud = a[0]; drawChips(); draw(); }); b.setAttribute('data-aud', a[0] || 'all'); chipBar.appendChild(b); });
          var f = chip('雛形ありだけ', state.withFile, function () { state.withFile = !state.withFile; drawChips(); draw(); }); f.setAttribute('data-filter-file', '1'); chipBar.appendChild(f);
          var o = chip('廃止も見る', state.old, function () { state.old = !state.old; drawChips(); draw(); }); o.setAttribute('data-filter-old', '1'); chipBar.appendChild(o);
        }
        function draw() {
          var shown = recs.filter(function (r) {
            var old = r['使用状況'] && r['使用状況'].value === '廃止';
            if (old && !state.old) return false;
            if (state.aud && !((r['見られる立場'] && r['見られる立場'].value) || []).some(function (x) { return x === state.aud; })) return false;
            if (state.withFile && !(r[FILE_FIELD].value || []).length) return false;
            return true;
          });
          var none = shown.filter(function (r) { return !(r[FILE_FIELD].value || []).length; }).length;
          info.textContent = (state.aud ? (AUDS.filter(function (a) { return a[0] === state.aud; })[0][1]) + 'の' : '') + '書類 ' + shown.length + ' 件　雛形あり ' + (shown.length - none) + ' 件　雛形なし ' + none + ' 件' + (uc ? '' : '　（案件くんを読めないため、作成数は出せません）');
          tableHost.textContent = ''; cells = [];
          var wrap = el('div', 'overflow-x:auto'); var tb = el('table', 'border-collapse:collapse;width:100%;min-width:' + (isMobile ? '560px' : '720px'));
          var head = el('tr'); ['工程', '書類名', '見られる立場', '雛形', '作成数', '最終作成日', '点検'].forEach(function (h) { head.appendChild(el('th', 'text-align:left;border-bottom:2px solid #cfd8e3;padding:4px 6px;white-space:nowrap', h)); });
          tb.appendChild(head);
          shown.forEach(function (r) {
            var name = r[NAME_FIELD].value, files = r[FILE_FIELD].value || [], used = uc && uc.usage[name], old = r['使用状況'] && r['使用状況'].value === '廃止';
            var tr = el('tr', old ? 'color:#999' : '');
            var td = function (t, st) { var c = el('td', 'border-bottom:1px solid #eee;padding:4px 6px;' + (st || ''), t); tr.appendChild(c); return c; };
            td(r['工程'] ? r['工程'].value : '');
            td(name + (old ? '（廃止）' : ''), 'font-weight:bold');
            td(((r['見られる立場'] && r['見られる立場'].value) || []).join('・'), 'font-size:11px');
            td(files.length ? files[0].name : '未登録', files.length ? 'color:#1f7a3d' : 'color:#c0392b;font-weight:bold');
            td(uc ? String(used ? used.count : 0) : '—', 'text-align:right');
            td(used && used.last ? used.last : '');
            var chk = td(files.length ? '（未点検）' : '', 'white-space:nowrap'); cells.push({ rec: r, cell: chk, files: files, name: name });
            tb.appendChild(tr);
          });
          wrap.appendChild(tb); tableHost.appendChild(wrap);
        }
        drawChips(); body.appendChild(chipBar); body.appendChild(info); body.appendChild(tableHost); draw();
        var go = el('button', 'margin-top:8px;padding:' + (isMobile ? '10px 14px' : '5px 10px') + ';cursor:pointer', '表示中の雛形を、すべて点検する（{{ }}の打ち間違いを探す）'); go.type = 'button'; go.setAttribute('data-master-check-all', '1');
        go.addEventListener('click', function () {
          go.disabled = true; var bad = 0, n = 0;
          cells.filter(function (c) { return c.files.length; }).reduce(function (p, c) {
            return p.then(function () {
              c.cell.textContent = '点検中…';
              return download(c.files[0].fileKey).then(function (buf) { return placeholders(buf, extOf(c.files[0].name)); }).then(unknownNames).then(function (b) {
                n++;
                if (b === null) { c.cell.textContent = '（点検できません）'; c.cell.style.color = '#999'; }
                else if (b.length) { bad++; c.cell.textContent = '注意：' + b.map(function (x) { return '{{' + x + '}}'; }).join(' '); c.cell.style.color = '#c0392b'; c.cell.style.whiteSpace = 'normal'; }
                else { c.cell.textContent = 'OK'; c.cell.style.color = '#1f7a3d'; }
              }, function () { c.cell.textContent = '（読めません）'; });
            });
          }, Promise.resolve()).then(function () { go.disabled = false; sumNote.textContent = n + '件を点検：注意 ' + bad + ' 件'; sumNote.style.color = bad ? '#c0392b' : '#1f7a3d'; });
        });
        var sumNote = el('span', 'margin-left:8px;font-weight:bold'); body.appendChild(go); body.appendChild(sumNote);
      }).catch(function (e) { body.style.color = '#c0392b'; body.textContent = '読み込めませんでした：' + ((e && e.message) || ''); });
    };
    box.addEventListener('toggle', load);
    box.open = true; load();   // 最初から開いて、立場で絞り込めるようにする
  }

  // ---- 一覧画面：複数のファイルを、まとめて差し替える ----
  kintone.events.on(['app.record.index.show', 'mobile.app.record.index.show'], guard(function (event) {
    var host = hostFor('index');
    if (!host || host.querySelector('[data-master-bulk]')) return event;
    var msg = null;
    host.appendChild(button('雛形をまとめて差し替える', 'data-master-bulk', function () {
      msg.style.color = '#1f7a3d'; msg.textContent = '';
      pickFiles(true).then(function (files) {
        if (!files.length) return;
        return api('GET', '/k/v1/records', { app: appId(), query: 'limit 500' }).then(function (res) {
          var byName = {}; res.records.forEach(function (r) { var n = r[NAME_FIELD].value; (byName[n] = byName[n] || []).push(r); });
          var plan = [], skip = [];
          files.forEach(function (f) {
            var hit = byName[baseName(f.name)];
            if (hit && hit.length === 1) plan.push({ file: f, rec: hit[0] });
            else skip.push(f.name + (hit && hit.length > 1 ? '（同じ書類名が複数あります）' : '（同じ書類名が、ありません）'));
          });
          return Promise.all(plan.map(function (p) {
            return readFile(p.file).then(function (buf) { return placeholders(buf, extOf(p.file.name)); }).then(unknownNames).then(function (bad) { p.bad = bad || []; });
          })).then(function () { return [plan, skip]; });
        }).then(function (ps) {
          var plan = ps[0], skip = ps[1];
          if (!plan.length) throw new Error('差し替えられるファイルが、ありません。ファイル名（拡張子をのぞく）を、書類名と同じにしてください。\n' + skip.join('\n'));
          var m = '次の' + plan.length + '件を、差し替えます。\n' + plan.map(function (p) { return '・' + p.rec[NAME_FIELD].value + ' ← ' + p.file.name + (p.bad.length ? '　【注意】案件くんにない項目名：' + p.bad.join('、') : ''); }).join('\n');
          if (skip.length) m += '\n\n差し替えないもの（書類名と合いません）：\n' + skip.map(function (s) { return '・' + s; }).join('\n');
          m += '\n\n古い雛形は「旧版_' + stamp() + '_」を付けて残します。よろしいですか？';
          if (!confirm(m)) return;
          var done = 0, failed = [];
          return plan.reduce(function (p, it) {
            return p.then(function () {
              msg.textContent = '差し替えています… ' + (done + failed.length + 1) + '/' + plan.length;
              return swap(it.rec, it.file).then(function () { done++; }, function (e) { failed.push(it.file.name + '：' + ((e && e.message) || '失敗')); });
            });
          }, Promise.resolve()).then(function () {
            msg.style.color = failed.length ? '#c0392b' : '#1f7a3d';
            msg.textContent = done + '件を、差し替えました。' + (failed.length ? '失敗：' + failed.join('、') : '') + (skip.length ? '　対象外：' + skip.length + '件' : '');
          });
        });
      }).catch(function (e) { msg.style.color = '#c0392b'; msg.textContent = (e && e.message) || '差し替えに失敗しました'; });
    }, true));
    msg = note(host);
    host.appendChild(button('{{ }}の候補', 'data-master-cand', function () { candPanel(); }));
    dashboard(wideHost() || host);
    return event;
  }));
})();
