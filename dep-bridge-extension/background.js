/* ============================================================================
 * 站存-发车流程填表桥 —— background（MV3 service worker）
 *
 * 链路：
 *   计算器页面 (window.postMessage)
 *     → content.js（ISOLATED world，收到后转 chrome.runtime）
 *     → background.js（本文件，路由）
 *     → 报表标签页 content.js
 *     → page-fill.js（MAIN world，真正写单元格）
 *
 * 反向：报表页把「编组车次」列表回传 → 计算器页面（用于车次重复校验）。
 * 跨域可用：消息都经由扩展中转，不受同源策略限制。
 * ==========================================================================*/
'use strict';

var CHANNEL = '__DEP_BRIDGE__';

/* 报表地址由 popup 输入框设置，存在 chrome.storage.local（键：depReportUrl）。
 * 这里只留内存缓存，启动时 + 变更时从 storage 载入：
 *   REPORT_URL   = 用户输入的完整地址
 *   REPORT_MATCH = 解析出的 host（如 10.190.136.28:6060），用于判定哪个标签页是报表页 */
var cfg = { reportUrl: '', reportMatch: '' };

/* 地图 token 接口（可在 popup 里改，存在 storage：
 *   depTokenUrl      接口完整地址
 *   depTokenReferer  期望的 Referer（通过 DNR 动态规则附加到页面请求上） */
var mapTokenCfg = {
  url: 'http://10.208.2.72:8080/getToken?ip=10.208.2.72',
  referer: ''
};

/** 从 URL 解析 host，作为「报表页」匹配特征；解析失败则退回原串 */
function hostOf(u) {
  try { return new URL(u).host; } catch (e) { return u; }
}

function loadCfg(cb) {
  chrome.storage.local.get(['depReportUrl', 'depTokenUrl', 'depTokenReferer'], function (r) {
    var url = (r && r.depReportUrl) || '';
    cfg.reportUrl = url;
    cfg.reportMatch = url ? hostOf(url) : '';
    if (r && r.depTokenUrl) mapTokenCfg.url = r.depTokenUrl;
    mapTokenCfg.referer = (r && r.depTokenReferer) || '';
    applyRefererRule();
    if (cb) cb();
  });
}

/** 给「页面发起的」token 请求附加 Referer（DNR 动态规则）。
 *  注意：DNR 不作用于扩展自身发起的请求，所以这条只影响地图页面的请求。 */
function applyRefererRule() {
  var RULE_ID = 9001;
  if (!chrome.declarativeNetRequest || !chrome.declarativeNetRequest.updateDynamicRules) return;
  try {
    if (!mapTokenCfg.referer || !mapTokenCfg.url) {
      chrome.declarativeNetRequest.updateDynamicRules({ removeRuleIds: [RULE_ID] }, function () {});
      return;
    }
    var u = new URL(mapTokenCfg.url);
    chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: [RULE_ID],
      addRules: [{
        id: RULE_ID,
        priority: 1,
        action: {
          type: 'modifyHeaders',
          requestHeaders: [{ header: 'Referer', operation: 'set', value: mapTokenCfg.referer }]
        },
        condition: { urlFilter: u.origin + u.pathname, resourceTypes: ['xmlhttprequest'] }
      }]
    }, function () { if (chrome.runtime.lastError) { /* Referer 可能被限制，忽略 */ } });
  } catch (e) {}
}
loadCfg();
chrome.storage.onChanged.addListener(function (changes, area) {
  if (area === 'local' && changes.depReportUrl) {
    loadCfg();
    clearPendingFill();                            // 报表地址变更，丢弃旧 fill，避免给老标签重发
  }
});

function send(tabId, payload) {
  try {
    var p = chrome.tabs.sendMessage(tabId, { channel: CHANNEL, payload: payload });
    if (p && typeof p.catch === 'function') p.catch(function () {});  // 页面无 content script 时静默
  } catch (e) {}
}

function isReport(url) {
  // 严格按「同 host + 同 pathname」判定，避免同一 host 上其他页面被误认成报表页
  if (!url || !cfg.reportUrl) return false;
  try {
    var a = new URL(url);
    var b = new URL(cfg.reportUrl);
    return a.host === b.host && a.pathname === b.pathname;
  } catch (e) { return false; }
}

/* ====== 待发 fill 队列（用于自动开报表页后延迟投递） =====================
 * 解决：之前用 [2000,4000,7000] 的 setTimeout 重试，会让"上次"点击的旧 fill
 * 在新打开的报表页加载完后被自动填上，看起来像"没点就填了"。
 * 改为：每次新 fill 来时刷新时间戳，30 秒内重试，过期自动丢弃。 */
var pendingFill = null;             // { payload, tabId, ts }
var pendingFillTimer = null;
var FILL_TTL_MS = 30 * 1000;         // 30 秒过期
var FILL_MAX_TRIES = 20;             // 最多重试 20 次（约 6 秒）

function clearPendingFill() {
  if (pendingFillTimer) { clearInterval(pendingFillTimer); pendingFillTimer = null; }
  pendingFill = null;
}

function pumpPendingFill() {
  if (!pendingFill) return;
  if (Date.now() - pendingFill.ts > FILL_TTL_MS) { clearPendingFill(); return; }
  try { send(pendingFill.tabId, pendingFill.payload); } catch (e) {}
}

/* ====== 工具栏图标颜色：已连接(报表页≥1)绿，未连接灰 ====== */
function makeIcon(rgb) {
  var s = 32, data = new Uint8ClampedArray(s * s * 4);
  var cx = s / 2, cy = s / 2, r = s / 2 - 1;
  for (var y = 0; y < s; y++) {
    for (var x = 0; x < s; x++) {
      var idx = (y * s + x) * 4;
      var dx = x - cx + 0.5, dy = y - cy + 0.5;
      if (dx * dx + dy * dy <= r * r) {
        data[idx] = rgb[0]; data[idx + 1] = rgb[1]; data[idx + 2] = rgb[2]; data[idx + 3] = 255;
      } else {
        data[idx + 3] = 0;                       // 圆外透明
      }
    }
  }
  return new ImageData(data, s, s);
}

function updateIcon() {
  chrome.tabs.query({}, function (tabs) {
    var n = (tabs || []).filter(function (t) { return isReport(t.url); }).length;
    var rgb = n > 0 ? [46, 160, 67] : [150, 150, 150];   // 绿 / 灰
    try { chrome.action.setIcon({ imageData: makeIcon(rgb) }); } catch (e) {}
    try { chrome.action.setTitle({ title: n > 0 ? '站存-发车流程填表桥（已连接 ' + n + ' 个报表页）' : '站存-发车流程填表桥（未连接）' }); } catch (e) {}
  });
}

// 标签变化时立即刷新；alarms 定时兜底（可唤醒休眠的 service worker）
['onUpdated', 'onActivated', 'onRemoved', 'onCreated'].forEach(function (ev) {
  if (chrome.tabs[ev]) chrome.tabs[ev].addListener(updateIcon);
});
try {
  chrome.alarms.create('depIconTick', { periodInMinutes: 0.1 });   // 每 6 秒
  chrome.alarms.onAlarm.addListener(function (a) { if (a.name === 'depIconTick') updateIcon(); });
} catch (e) {}
updateIcon();

/* ====== 地图 token 桥接 ==================================================
 * 地图（chinamap.html）自己请求 /getToken 拿不到 token（跨域 / Referer 受限），
 * 改由扩展直接请求真实后端拿 token，再注入地图的 sessionStorage.token。
 * 地图代码里所有 $.ajax 都会 setRequestHeader("Authorization", sessionStorage.token)，
 * 所以只要 sessionStorage.token 是对的，整个地图的 API 就都带上了正确的 Authorization。 */
function fetchMapToken(cb) {
  var opts = { credentials: 'omit', cache: 'no-store' };
  if (mapTokenCfg.referer) opts.headers = { Referer: mapTokenCfg.referer };
  fetchWithRetry(mapTokenCfg.url, opts, function (err, token) {
    if (err) { cb && cb(err); return; }
    var t = (token || '').trim();
    chrome.storage.local.set({ authorization: t, authorizationTs: Date.now() }, function () {});
    cb && cb(null, t);
  });
}

/** 把 token 广播给所有标签页（content.js 收到后转给页面，由 page-fill.js 落地） */
function broadcastMapToken(token) {
  chrome.tabs.query({}, function (tabs) {
    (tabs || []).forEach(function (t) { send(t.id, { type: 'map-token', token: token }); });
  });
}

/** 带 401 自动重试的 fetch 封装
 *  读取 storage 中的 depRetryEnabled / depRetryCount 决定是否重试 */
function fetchWithRetry(url, options, cb) {
  var attempt = 0;
  chrome.storage.local.get(['depRetryEnabled', 'depRetryCount'], function (r) {
    var enabled = !!(r && r.depRetryEnabled);
    var maxRetries = (r && r.depRetryCount != null) ? parseInt(r.depRetryCount, 10) : 2;
    if (isNaN(maxRetries) || maxRetries < 0) maxRetries = 0;
    if (maxRetries > 10) maxRetries = 10;
    var maxAttempts = enabled ? (1 + maxRetries) : 1;

    function doFetch() {
      attempt++;
      fetch(url, options)
        .then(function (resp) {
          if (resp.status === 401 && attempt < maxAttempts) {
            console.warn('[dep-bridge] 401 on attempt ' + attempt + '/' + maxAttempts + ', retrying...');
            doFetch();
            return;
          }
          if (!resp.ok) throw new Error('HTTP ' + resp.status);
          return resp.text().then(function (t) { cb(null, t); });
        })
        .catch(function (e) {
          if (attempt < maxAttempts) {
            console.warn('[dep-bridge] fetch error on attempt ' + attempt + '/' + maxAttempts + ': ' + (e.message || e) + ', retrying...');
            doFetch();
            return;
          }
          cb((e && e.message) || String(e));
        });
    }
    doFetch();
  });
}

chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
  if (!msg || msg.channel !== CHANNEL) {
    // content.js 开局取已保存的 token（写进页面 sessionStorage 用）
    if (msg && msg.type === 'getStoredToken') {
      chrome.storage.local.get(['authorization', 'authorizationTs'], function (r) {
        sendResponse({ token: (r && r.authorization) || '', ts: (r && r.authorizationTs) || 0 });
      });
      return true;
    }
    // popup：测试请求（自定义 URL / Referer），只回显状态码与响应内容，不注入
    if (msg && msg.type === 'testToken') {
      var testUrl = msg.url || mapTokenCfg.url;
      // 顺手把这次填的 URL / Referer 存下来，供后续自动取 token 使用
      chrome.storage.local.set({
        depTokenUrl: testUrl,
        depTokenReferer: msg.referer || ''
      }, function () { loadCfg(); });
      var t0 = Date.now();
      var testOpts = { credentials: 'omit', cache: 'no-store' };
      var testReferer = msg.referer || mapTokenCfg.referer;
      if (testReferer) testOpts.headers = { Referer: testReferer };
      fetch(testUrl, testOpts)
        .then(function (r) {
          return r.text().then(function (text) {
            sendResponse({ ok: r.ok, status: r.status, text: text, ms: Date.now() - t0, url: testUrl });
          });
        })
        .catch(function (e) {
          sendResponse({
            ok: false, status: 0, text: '', url: testUrl, ms: Date.now() - t0,
            error: (e && e.message) || String(e)
          });
        });
      return true;
    }
    // popup：手动取一次地图 token（调试用），取到后顺手广播注入
    if (msg && msg.type === 'fetchMapToken') {
      fetchMapToken(function (err, token) {
        sendResponse({ ok: !err, token: token, error: err });
        if (!err && token) broadcastMapToken(token);
      });
      return true;
    }
    // popup：把已知 token 直接推送注入
    if (msg && msg.type === 'pushMapToken' && msg.token) {
      broadcastMapToken(msg.token);
      sendResponse({ ok: true });
      return true;
    }
    // popup 查询状态
    if (msg && msg.type === 'status') {
      chrome.tabs.query({}, function (tabs) {
        var n = (tabs || []).filter(function (t) { return isReport(t.url); }).length;
        sendResponse({ reportUrl: cfg.reportUrl, reportMatch: cfg.reportMatch, reportTabs: n });
      });
      return true;                                  // 异步 sendResponse
    }
    return;
  }
  var payload = msg.payload || {};
  var fromTabId = sender.tab ? sender.tab.id : null;

  // 地图页索要 token：扩展去真实后端取，取到后回发该标签页（含其中的 iframe）
  if (payload.type === 'map-need-token') {
    fetchMapToken(function (err, token) {
      if (err || !token) return;
      if (fromTabId != null) send(fromTabId, { type: 'map-token', token: token });
    });
    return;
  }

  chrome.tabs.query({}, function (tabs) {
    var hitReport = 0;
    (tabs || []).forEach(function (t) {
      if (t.id === fromTabId) return;               // 不发回来源标签页，避免回环
      if (payload.type === 'fill' || payload.type === 'fillByStrategy') {
        // 填表指令：只发给报表页
        if (isReport(t.url)) { hitReport++; send(t.id, payload); }
      } else {
        // checiList / filled 等回传：只发给非报表页（即计算器页面）
        if (!isReport(t.url)) send(t.id, payload);
      }
    });

    // 报表标签页没开 → 自动开一个，加载完后再补发；但 fill 30 秒过期，到期自动丢弃
    if (payload.type === 'fill' && hitReport === 0 && cfg.reportUrl) {
      try {
        var payloadCopy = Object.assign({}, payload, { ts: Date.now() });
        clearPendingFill();                       // 旧的取消，换成这次新点击的
        pendingFill = { payload: payloadCopy, tabId: null, ts: payloadCopy.ts };
        chrome.tabs.create({ url: cfg.reportUrl }, function (tab) {
          if (!tab || !tab.id) { clearPendingFill(); return; }
          pendingFill.tabId = tab.id;
          // 每 300ms 尝试一次，每次校验 ttl 与次数；超时/过限自动清掉
          var tries = 0;
          pendingFillTimer = setInterval(function () {
            tries++;
            if (!pendingFill || Date.now() - pendingFill.ts > FILL_TTL_MS) { clearPendingFill(); return; }
            if (tries > FILL_MAX_TRIES) { clearPendingFill(); return; }
            pumpPendingFill();
          }, 300);
          pumpPendingFill();                      // 立即试一次（应对报表页其实已 ready 的情形）
        });
      } catch (e) { clearPendingFill(); }
    }
  });
});