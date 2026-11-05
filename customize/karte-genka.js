/* ============================================================
   現場カルテ(826)に「原価管理表を作る」ボタンを出す（PC用）
   ・押すと、書類マスタの「原価管理表」の雛形（Excel）に、この案件の、現場カルテの数字と、請求履歴（請求くん）を入れて、ダウンロードする
   ・現場カルテや、ほかのアプリには、何も書き込まない（ファイルを作るだけ）
   ・入る内容：工事名・工事コード・元請・施工開始/完了・受注金額・予算金額（予算01）・予算原価（発注確定額）・原価明細（請求くんの請求）
   ・同じ部品（doc-button.js の「Excelに入れる」部分）を、借りて使う（ボタンなどは、出さない）
   設置：JSZip → karte-genka.js → doc-button.js の順に、PC用JavaScriptへ登録
   ============================================================ */
window.__TS_DOC_ENGINE_ONLY = true;   // doc-button.js を、「部品を貸すだけ」の形で使う（画面には、何も出さない）
(function () {
  'use strict';
  var MASTER_APP_NAME = '書類マスタ', DOC_NAME = '原価管理表', BILL_APP = 829, BTN_ID = 'ts-karte-genka-btn';
  // 現場カルテの項目コード
  var F = { anken: 'anken', name: 'proj_name', code: 'koji_code', code2: 'proj_code', moto: 'moto', start: 'start_date', end: 'end_date', sales: 'sales', budget: 'budget', ordered: 'ordered' };

  function api(method, path, params) { return kintone.api(kintone.api.url(path, true), method, params); }
  function val(rec, c) { var f = rec && rec[c]; return f && f.value != null ? String(f.value).trim() : ''; }
  function num(rec, c) { var v = Number(val(rec, c).replace(/,/g, '')); return isFinite(v) ? v : 0; }
  function dateJa(iso) { var m = String(iso || '').match(/^(\d{4})-(\d{2})-(\d{2})/); return m ? m[1] + '年' + Number(m[2]) + '月' + Number(m[3]) + '日' : ''; }
  function today() { var d = new Date(); return d.getFullYear() + '年' + (d.getMonth() + 1) + '月' + d.getDate() + '日'; }

  function findMasterTemplate() {
    return api('GET', '/k/v1/apps', { name: MASTER_APP_NAME, limit: 100 }).then(function (r) {
      var m = r.apps.filter(function (a) { return a.name === MASTER_APP_NAME; });
      if (m.length !== 1) throw new Error('「' + MASTER_APP_NAME + '」アプリが、見つからないか、複数あります');
      return api('GET', '/k/v1/records', { app: m[0].appId, query: '書類名 = "' + DOC_NAME + '" limit 1' });
    }).then(function (r) {
      var rec = r.records[0], files = rec && rec['雛形'] && rec['雛形'].value;
      if (!files || !files.length) throw new Error('書類マスタの「' + DOC_NAME + '」に、雛形が、登録されていません');
      return fetch('/k/v1/file.json?fileKey=' + encodeURIComponent(files[0].fileKey), { headers: { 'X-Requested-With': 'XMLHttpRequest' } })
        .then(function (res) { if (!res.ok) throw new Error('雛形を、取り出せませんでした（' + res.status + '）'); return res.arrayBuffer(); });
    });
  }

  function build(rec) {
    var E = window.TSDocEngine;
    if (!E) throw new Error('部品（doc-button.js）が、読み込まれていません。JavaScriptの読み込み順を、確認してください');
    var values = {
      '工事名称': val(rec, F.name), '工事番号': val(rec, F.code) || val(rec, F.code2), '元請 会社名': val(rec, F.moto),
      '📅施工予定日': dateJa(val(rec, F.start)), '🛠️施工完了日': dateJa(val(rec, F.end)), '作成日': today(), '版': 'v1', '書類名': DOC_NAME,
      '工事場所': '', '👷工事担当者': ''   // 現場カルテに、対応する項目がない：空欄にして、あとから入力する
    };
    var anken = val(rec, F.anken);
    var bills = anken ? api('GET', '/k/v1/records', { app: BILL_APP, query: 'anken = "' + anken.replace(/["\\]/g, '') + '" limit 500' }).then(function (r) { return r.records || []; }) : Promise.resolve([]);
    return Promise.all([findMasterTemplate(), bills.catch(function () { return null; })]).then(function (res) {
      var buf = res[0], billRecs = res[1];
      return E.fillTemplate(JSZip, buf, 'xlsx', values).then(function (out) {
        var rows = billRecs ? E.billRowsOf(billRecs) : [];
        var budget = { total: num(rec, F.budget), cost: num(rec, F.ordered) };
        var won = { total: num(rec, F.sales), cost: null };   // 受注原価は、現場カルテに対応する項目がないため、空のまま
        return E.genkaFill(JSZip, out.data, rows, null, won.total ? won : null, budget.total ? budget : null).then(function (r) {
          return { data: r.data, count: r.count, billsOk: billRecs !== null, isGenka: r.isGenka, name: (val(rec, F.name) || '案件') + '_原価管理表.xlsx', sales: won.total, budget: budget };
        });
      });
    });
  }

  function save(data, filename) {
    var url = URL.createObjectURL(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    var a = document.createElement('a'); a.href = url; a.download = filename; document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 10000);
  }

  function addButton(rec) {
    var host = kintone.app.record.getHeaderMenuSpaceElement();
    if (!host || document.getElementById(BTN_ID)) return;
    var wrap = document.createElement('span'); wrap.style.cssText = 'display:inline-flex;align-items:center;gap:8px;margin-left:8px';
    var b = document.createElement('button'); b.id = BTN_ID; b.type = 'button'; b.textContent = '📊 原価管理表を作る';
    b.style.cssText = 'padding:7px 14px;font-size:14px;font-weight:bold;cursor:pointer;background:#1F3864;color:#fff;border:0;border-radius:6px';
    var msg = document.createElement('span'); msg.setAttribute('data-karte-genka-msg', '1'); msg.style.cssText = 'font-size:12px;color:#444;max-width:520px';
    b.addEventListener('click', function () {
      b.disabled = true; msg.style.color = '#444'; msg.textContent = '作っています…';
      build(rec).then(function (r) {
        if (!r.isGenka) throw new Error('書類マスタの「' + DOC_NAME + '」が、原価管理表の雛形ではありません');
        save(r.data, r.name); b.disabled = false;
        msg.textContent = 'ダウンロードしました：' + r.name + '（請求履歴 ' + (r.billsOk ? r.count + ' 行' : '：請求くんを読めませんでした') + (r.sales ? '／受注金額・' : '／') + (r.budget.total ? '予算を入れました）' : '予算は空です）');
      }).catch(function (e) { b.disabled = false; msg.style.color = '#c0392b'; msg.textContent = '作れませんでした：' + ((e && e.message) || e); });
    });
    wrap.appendChild(b); wrap.appendChild(msg); host.appendChild(wrap);
  }

  kintone.events.on('app.record.detail.show', function (event) { try { addButton(event.record); } catch (e) { console.error(e); } return event; });
})();
