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
        var body = { app: kintone.app.getId(), id: id, record: {} };
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

  function button(text, attr, onClick) {
    var b = el('button', 'margin:0 8px 0 0;padding:6px 12px;font-size:13px;cursor:pointer', text);
    b.type = 'button'; b.setAttribute(attr, '1'); b.addEventListener('click', onClick);
    return b;
  }
  function note(host) { var n = el('span', 'font-size:12px;margin-left:4px;color:#1f7a3d'); n.setAttribute('data-master-note', '1'); host.appendChild(n); return n; }

  // ---- 詳細画面：この書類の雛形を差し替える ----
  kintone.events.on('app.record.detail.show', function (event) {
    var host = kintone.app.record.getHeaderMenuSpaceElement && kintone.app.record.getHeaderMenuSpaceElement();
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
    }));
    host.appendChild(button('試しに作る', 'data-master-try', function () { tryPanel(rec, name, msg); }));
    msg = note(host);
    return event;
  });

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
    card.appendChild(el('div', 'font-size:12px;color:#666;margin-bottom:8px', '見本にする案件を選んでください。いまの雛形（' + files[0].name + '）で、書類を作って、点検します。案件くんにも、ここにも、保存しません（案件くんは、読むだけです）。'));
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
      E.ASK.forEach(function (a) { values[a.key] = '【入力欄：' + a.key + '】'; });   // 書類作成のとき、画面で入力する項目は、目印の言葉を入れて、位置が分かるようにする
      values['版'] = 'v0（試し）'; values['書類名'] = name;
      download(files[0].fileKey).then(function (buf) { return E.fillTemplate(window.JSZip, buf, ext, values); }).then(function (out) {
        return E.extractText(window.JSZip, out.data, ext).then(function (text) {
          var issues = E.checkText(text, values).filter(function (x) { return x.indexOf('【入力欄') < 0; });
          if (out.unknown && out.unknown.length) issues.unshift('案件くんにない項目名（空になります）：' + out.unknown.map(function (n) { return /^\{\{/.test(n) ? n : '{{' + n + '}}'; }).join(' '));
          if (out.empty && out.empty.length) issues.push('案件くんの値が空の項目：' + out.empty.join('、'));
          result.textContent = '';
          result.appendChild(el('div', 'font-weight:bold;color:' + (issues.length ? '#c0392b' : '#1f7a3d'), issues.length ? '点検の結果：気になるところが、' + issues.length + '件あります' : '点検の結果：問題は、見つかりませんでした'));
          issues.forEach(function (t) { result.appendChild(el('div', 'font-size:13px;margin:2px 0 0 8px', '・' + t)); });
          result.appendChild(el('div', 'font-size:12px;color:#666;margin-top:4px', '※「【入力欄：…】」は、書類作成のとき、画面で入力する項目の場所です。'));
          var fname = '試し_' + title(record, l) + '_' + name + '.' + ext;
          var dl = el('button', 'margin-top:8px;padding:6px 12px;cursor:pointer', '試しに作ったファイルを、開く（保存はされません）'); dl.type = 'button'; dl.setAttribute('data-try-download', '1');
          dl.addEventListener('click', function () {
            var a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([out.data])); a.download = fname; document.body.appendChild(a); a.click(); document.body.removeChild(a);
          });
          result.appendChild(dl);
        });
      }).catch(function (e) { result.style.color = '#c0392b'; result.textContent = '試しに作れませんでした：' + ((e && e.message) || ''); });
    }
    loadList();
  }

  // ---- 一覧画面：複数のファイルを、まとめて差し替える ----
  kintone.events.on('app.record.index.show', function (event) {
    var host = kintone.app.getHeaderMenuSpaceElement && kintone.app.getHeaderMenuSpaceElement();
    if (!host || host.querySelector('[data-master-bulk]')) return event;
    var msg = null;
    host.appendChild(button('雛形をまとめて差し替える', 'data-master-bulk', function () {
      msg.style.color = '#1f7a3d'; msg.textContent = '';
      pickFiles(true).then(function (files) {
        if (!files.length) return;
        return api('GET', '/k/v1/records', { app: kintone.app.getId(), query: 'limit 500' }).then(function (res) {
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
    }));
    msg = note(host);
    return event;
  });
})();
