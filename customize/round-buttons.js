/* ============================================================
   画面の右下のボタンを、丸いボタンにそろえる（近づけると、文字つきに広がる）
   ・元のボタンのJSは、書き換えません。元のボタンが画面に出たあとで、見た目だけを丸い形に直します
     （押したときの動き・開く先のURL・鍵は、元のまま）
   ・対象は、下の TARGETS に書いたボタン（要素ID）。ボタンの文字の先頭の絵文字が、丸いボタンの絵になります
   ・「書類を作る」「案件管理アプリを開く」（PC）と同じ形です

   設置：見積くん・発注くん・請求くん・案件くん（モバイル）の、PC用・モバイル用のJavaScriptに
         元のJSの“あと”に追加（scripts/apply-customize.js --manifest=… が追加します）
   ============================================================ */
(function () {
  'use strict';

  // 丸くするボタンの要素ID → 並びの位置（bottom の px）
  var TARGETS = {
    'open-mitsumori-btn': 24,          // 見積くん（PC）
    'open-mitsumori-btn-m': 24,        // 見積くん（モバイル）
    'ts-hassou-admin-btn': 24,         // 発注くん
    'ts-seikyu-grid-btn': 24,          // 請求くん
    'taiyoOpenAppMobileBtn': 24        // 案件くん（モバイル）
  };
  var SIZE = 52;

  function roundUp(el, bottom) {
    if (el.getAttribute('data-round') === '1') return;
    el.setAttribute('data-round', '1');
    var full = (el.textContent || '').trim();
    var m = full.match(/^(\S+?)\s+(.*)$/);
    var icon = m ? m[1] : '●', label = m ? m[2] : full;
    el.setAttribute('aria-label', label);
    var bg = '';
    try { bg = getComputedStyle(el).backgroundColor; } catch (e) {}
    el.textContent = '';
    // 位置・大きさ・形を、丸いボタンに（色と、押したときの動きは、元のまま）
    el.style.cssText = (el.style.cssText || '') + ';' +
      'position:fixed !important;left:auto !important;transform:none !important;right:16px !important;bottom:' + bottom + 'px !important;' +
      'height:' + SIZE + 'px !important;min-width:' + SIZE + 'px !important;padding:0 14px !important;border-radius:' + (SIZE / 2) + 'px !important;' +
      'display:flex !important;align-items:center;justify-content:center;overflow:hidden;white-space:nowrap;box-sizing:border-box;' + (bg ? 'background:' + bg + ' !important;' : '');
    var i = document.createElement('span'); i.style.cssText = 'font-size:22px;line-height:1'; i.textContent = icon;
    var t = document.createElement('span');
    t.style.cssText = 'display:inline-block;max-width:0;opacity:0;margin-left:0;transition:max-width .25s,opacity .25s,margin-left .25s;font-size:14px;font-weight:bold';
    t.textContent = label;
    el.appendChild(i); el.appendChild(t);
    function open() { t.style.maxWidth = '260px'; t.style.opacity = '1'; t.style.marginLeft = '8px'; }
    function close() { t.style.maxWidth = '0'; t.style.opacity = '0'; t.style.marginLeft = '0'; }
    el.addEventListener('mouseenter', open); el.addEventListener('focus', open);
    el.addEventListener('mouseleave', close); el.addEventListener('blur', close);
  }

  function scan() {
    Object.keys(TARGETS).forEach(function (id) {
      var el = document.getElementById(id);
      if (el) roundUp(el, TARGETS[id]);
    });
  }

  // 元のボタンは、あとから（画面の読み込み後に）出るので、出たら、すぐに丸くする
  function start() {
    scan();
    try { new MutationObserver(scan).observe(document.body, { childList: true }); } catch (e) { setInterval(scan, 500); }
  }
  if (document.body) start(); else document.addEventListener('DOMContentLoaded', start);
})();
