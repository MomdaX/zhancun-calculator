/* ============================================================================
 * 站存-发车流程填表桥 —— content script（ISOLATED world，注入所有页面）
 *
 * 只做两件事（不碰页面逻辑）：
 *   1. 页面 window.postMessage({channel:'__DEP_BRIDGE__', ...})  → 转给 background
 *   2. background 转发过来的消息                                 → window.postMessage 回页面
 *
 * 页面与 content script 处于不同 world，但共享 window，postMessage 可互通。
 * ==========================================================================*/
'use strict';

(function () {
  var CHANNEL = '__DEP_BRIDGE__';
  if (window.__depBridgeInstalled) return;
  window.__depBridgeInstalled = true;

  // 地图页（chinamap.html）开局注入 token：
  // 从扩展本地存储取出 authorization，写进页面的 sessionStorage，
  // 这一步在 document_start 执行，早于页面自身的脚本，所以 GetToken() 能直接读到。
  if (/chinamap\.html/i.test(location.href)) {
    try {
      chrome.runtime.sendMessage({ type: 'getStoredToken' }, function (res) {
        if (res && res.token) {
          try { sessionStorage.setItem('token', res.token); } catch (e) {}
        }
      });
    } catch (e) {}
  }

  // 页面 → background
  window.addEventListener('message', function (ev) {
    if (ev.source !== window) return;               // 只收本页自己发的
    var d = ev.data;
    if (!d || d.channel !== CHANNEL) return;
    if (d.__from === 'dep-bridge-content') return;  // 自己发出的转发，跳过，断开回环
    try {
      var p = chrome.runtime.sendMessage({ channel: CHANNEL, payload: d });
      if (p && typeof p.catch === 'function') p.catch(function () {});
    } catch (e) {}
  });

  // background → 页面（page-fill.js 与页面自身脚本都会收到）
  chrome.runtime.onMessage.addListener(function (msg) {
    if (!msg || msg.channel !== CHANNEL) return;
    // 打标记再发回页面：这样本脚本的 window 监听会识别并跳过，避免无限回环
    window.postMessage(Object.assign({}, msg.payload, { __from: 'dep-bridge-content' }), '*');
  });
})();
