/* ============================================================
   案件管理アプリを開くボタン（PC用）
   ・画面の右下に、丸いボタン（📋）が浮かびます。マウスを近づけると、文字つきに広がります
   ・押すと、案件管理アプリ(Render)が新しいタブで開きます
   ・レコード詳細（現場カルテ）／一覧／追加／編集 すべての画面で表示
   ・「書類を作る」ボタン（doc-button.js）と同じ形です。位置は、そのすぐ下です

   設置：このファイルを、アプリ設定 →「JavaScript / CSSでカスタマイズ」→「JavaScript（PC用）」に登録
         （scripts/apply-customize.js が、同じ名前の古い taiyo-open-app.js と入れ替えます）
   ※ このボタンを出したいアプリごとに、設定が必要です。
   ============================================================ */
(function () {
  'use strict';

  // ▼ 開きたいアプリのURL（変えたいときはここだけ直す）
  var APP_URL = 'https://anken-kanri-1.onrender.com/';
  // 鍵（?k=）は、リポジトリには入れない。反映のときだけ、本番のスマホ版ボタンと同じ鍵を入れる（scripts/with-open-key.js）
  var OPEN_KEY = '';
  var LABEL = '案件管理アプリを開く';
  var SIZE = 52;

  function showButton() {
    if (document.getElementById('taiyo-open-app-btn')) return; // 二重表示を防ぐ
    var btn = document.createElement('button');
    btn.id = 'taiyo-open-app-btn';
    btn.type = 'button';
    btn.setAttribute('aria-label', LABEL);
    btn.style.cssText =
      'position:fixed;right:16px;bottom:24px;z-index:2147483647;height:' + SIZE + 'px;min-width:' + SIZE + 'px;' +
      'border:0;border-radius:' + (SIZE / 2) + 'px;background:#1F3864;color:#fff;padding:0 14px;' +
      'display:flex;align-items:center;justify-content:center;overflow:hidden;white-space:nowrap;' +
      'font-weight:700;font-size:14px;cursor:pointer;box-shadow:0 4px 12px rgba(0,0,0,.35);';
    var icon = document.createElement('span');
    icon.style.cssText = 'font-size:22px;line-height:1';
    icon.textContent = '📋';
    var text = document.createElement('span');
    text.style.cssText = 'display:inline-block;max-width:0;opacity:0;margin-left:0;transition:max-width .25s,opacity .25s,margin-left .25s';
    text.textContent = LABEL;
    btn.appendChild(icon); btn.appendChild(text);
    function open() { text.style.maxWidth = '200px'; text.style.opacity = '1'; text.style.marginLeft = '8px'; }
    function close() { text.style.maxWidth = '0'; text.style.opacity = '0'; text.style.marginLeft = '0'; }
    btn.addEventListener('mouseenter', open); btn.addEventListener('focus', open);
    btn.addEventListener('mouseleave', close); btn.addEventListener('blur', close);
    btn.addEventListener('click', function () { window.open(OPEN_KEY ? APP_URL + '?k=' + encodeURIComponent(OPEN_KEY) : APP_URL, '_blank'); });
    document.body.appendChild(btn);
  }

  // レコード詳細（現場カルテ）／一覧／追加／編集 すべてで表示
  var EVENTS = [
    'app.record.detail.show',
    'app.record.index.show',
    'app.record.create.show',
    'app.record.edit.show'
  ];
  kintone.events.on(EVENTS, function (event) {
    showButton();
    return event;
  });
})();
