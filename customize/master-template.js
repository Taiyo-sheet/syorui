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
(function () {
  'use strict';

  var KEEP = 3;
  var NAME_FIELD = '書類名', FILE_FIELD = '雛形';
  var ACCEPT = '.docx,.xlsx,.pptx,.pdf';

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
        return readFile(f).then(function (buf) { return placeholders(buf, extOf(f.name)); }).then(function (ph) {
          var olds = ((rec[FILE_FIELD] && rec[FILE_FIELD].value) || []).map(function (x) { return x.name; });
          if (!confirm('書類「' + name + '」の雛形を差し替えます。\n\n' + summary(f, ph) + '\nいまの雛形：' + (olds.length ? olds.join('、') : 'なし') + '\n（古いものは「旧版_' + stamp() + '_」を付けて残します。最新' + KEEP + '個まで）\n\nよろしいですか？')) return;
          msg.textContent = '差し替えています…';
          return swap(rec, f).then(function () { msg.textContent = '差し替えました。画面を読み込み直します'; setTimeout(function () { location.reload(); }, 800); });
        });
      }).catch(function (e) { msg.style.color = '#c0392b'; msg.textContent = (e && e.message) || '差し替えに失敗しました（元の雛形は、そのままです）'; });
    }));
    msg = note(host);
    return event;
  });

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
          if (!plan.length) throw new Error('差し替えられるファイルが、ありません。ファイル名（拡張子をのぞく）を、書類名と同じにしてください。\n' + skip.join('\n'));
          var m = '次の' + plan.length + '件を、差し替えます。\n' + plan.map(function (p) { return '・' + p.rec[NAME_FIELD].value + ' ← ' + p.file.name; }).join('\n');
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
