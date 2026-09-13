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
/* broadcast：是否启用「广播兜底」。
 * 精确匹配（isReport）没命中任何标签页时：
 *   开 → 广播给所有普通标签页，由页面自行判断是否含 contentPane（popup 里可关）
 *   关 → 只认精确匹配，地址没配对就发不出去（用于验证地址配置是否正确） */
var cfg = { reportUrl: '', reportMatch: '', broadcast: true };

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
  chrome.storage.local.get(['depReportUrl', 'depTokenUrl', 'depTokenReferer', 'depBroadcast'], function (r) {
    var url = (r && r.depReportUrl) || '';
    cfg.reportUrl = url;
    cfg.reportMatch = url ? hostOf(url) : '';
    // 没存过 → 默认开启广播兜底
    cfg.broadcast = (r && r.depBroadcast !== undefined) ? !!r.depBroadcast : true;
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
    loadCfg();                                     // 报表地址变更，重新载入匹配特征
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

/** 消息是否来自「地图页本身」（chinamap.html 所在的 frame）。
 *  content.js 注入在所有页面（<all_urls>）并会无条件转发页面的 postMessage，
 *  而扩展代发请求用的是 host_permissions + 已保存的 token —— 不做来源校验的话，
 *  任意网页只要知道协议就能借这条通道读内网接口、甚至套取 token。
 *  这里只放行 URL 里带 chinamap.html 的 frame，其它一律拒绝。 */
function isMapFrame(sender) {
  var u = (sender && sender.url) || '';
  return /chinamap\.html(\?|#|$)/i.test(u);
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

/** token 是否已过期或即将过期（剩余 < 5 分钟）。
 *  地图 token 有效期只有 8 小时，若不判断，每次打开地图都会先拿一个必然 401 的旧 token
 *  试一次（虽然后续会自动续期，但白等一个往返）。解析不出 exp（非标准 JWT）则视为不过期。 */
function jwtExpiringSoon(token) {
  try {
    var parts = String(token).split('.');
    if (parts.length < 2) return false;
    var b = parts[1].replace(/-/g, '+').replace(/_/g, '/');
    while (b.length % 4) b += '=';
    var payload = JSON.parse(atob(b));
    if (!payload || !payload.exp) return false;
    return (payload.exp * 1000) < (Date.now() + 5 * 60 * 1000);
  } catch (e) { return false; }
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

/* ====== 地图后端接口代理（map-api） ======================================
 * 为什么需要：地图页（chinamap.html）以 file:// 打开，origin 为 null，
 *   1. 带 Authorization 自定义头 → 触发 CORS 预检，而后端不处理 OPTIONS，
 *      预检返回非 2xx → 请求被浏览器拦下；
 *   2. file:// 页面发出的请求 Referer 恒为空（forbidden header，页面改不了）。
 * 因此地图的全部后端接口改由扩展代发：扩展有 host_permissions，既不受 CORS
 * 限制，也能在 fetch 上直接补 Referer / Authorization。
 *   map-api-proxy.js（页面）→ content.js → 本段 → fetch 真实后端
 *   → { type:'map-api-result' } 回传页面，由页面按 jQuery 语义唤起回调。
 * ======================================================================== */
var MAP_API_PATHS = [
  '/getToken', '/getReleaseVersionData',
  '/getMapDftPoint', '/getMapDlnPoint', '/getMapDndPoint',
  '/getMidStationByLj', '/getNode', '/getJl'
];

/** 把 {a:1,b:'x'} 拼成 a=1&b=x（与原 jQuery 的编码方式一致） */
function buildQuery(data) {
  if (!data || typeof data !== 'object') return '';
  return Object.keys(data).map(function (k) {
    var v = data[k];
    if (v == null) return '';
    return encodeURIComponent(k) + '=' + encodeURIComponent(String(v));
  }).filter(function (s) { return s; }).join('&');
}

/** 代发一次地图接口请求；401 时自动续期 token 并重试一次 */
function proxyMapApi(payload, fromTabId) {
  var reqId = payload.reqId || '';
  var base = String(payload.base || '').replace(/\/+$/, '');
  var path = String(payload.path || '');
  var reqBase = { type: 'map-api-result', reqId: reqId, ok: false, status: 0, contentType: '', body: '' };

  function reply(extra) {
    if (fromTabId == null) return;
    send(fromTabId, Object.assign({}, reqBase, extra || {}));
  }

  // 白名单 + 合法 base，避免被页面拿去打任意地址
  if (MAP_API_PATHS.indexOf(path) < 0 || !/^https?:\/\//i.test(base)) {
    reply({ error: '非法代理请求：' + base + path });
    return;
  }

  var qs = buildQuery(payload.data);
  var url = base + path + (qs ? ('?' + qs) : '');
  // Referer：后端认的值是 http://10.208.2.72:8080/cljl（见 map_data 下抓包样张）。
  // 优先用 popup 里配置的 depTokenReferer；没配就按 base + '/cljl' 兜底，保证非空且正确。
  var referer = mapTokenCfg.referer || (base + '/cljl');

  chrome.storage.local.get(['authorization'], function (r) {
    var token = (payload.token && String(payload.token).trim()) || ((r && r.authorization) || '').trim();
    if (token === '3231212') token = '';         // 兜底假 token 不下发
    doApiFetch(url, token, false);
  });

  function doApiFetch(u, token, retried) {
    var headers = { 'X-Requested-With': 'XMLHttpRequest', 'Referer': referer };
    if (token) headers['Authorization'] = token;

    fetch(u, { method: payload.method || 'GET', headers: headers, credentials: 'omit', cache: 'no-store' })
      .then(function (resp) {
        if (resp.status === 401 && !retried) {
          // token 过期：重新取一次、广播给页面，再用新 token 重试
          fetchMapToken(function (err, newToken) {
            if (err || !newToken) {
              reply({ ok: false, status: 401, error: 'token 续期失败：' + (err || 'empty') });
              return;
            }
            broadcastMapToken(newToken);
            doApiFetch(u, newToken, true);
          });
          return;
        }
        return resp.text().then(function (body) {
          reply({
            ok: resp.ok,
            status: resp.status,
            contentType: resp.headers.get('content-type') || '',
            body: body
          });
        });
      })
      .catch(function (e) {
        reply({ ok: false, status: 0, error: (e && e.message) || String(e) });
      });
  }
}

chrome.runtime.onMessage.addListener(function (msg, sender, sendResponse) {
  if (!msg || msg.channel !== CHANNEL) {
    // content.js 开局取已保存的 token（写进页面 sessionStorage 用）
    if (msg && msg.type === 'getStoredToken') {
      // 只有地图页能取 token（content.js 本身也只在地图页调用，这里再校验一次来源做深度防御）
      if (!isMapFrame(sender)) {
        console.warn('[dep-bridge] 拒绝非地图页取 token，来源：' + ((sender && sender.url) || '(空)'));
        sendResponse({ token: '', ts: 0 });
        return;
      }
      chrome.storage.local.get(['authorization', 'authorizationTs'], function (r) {
        var t = (r && r.authorization) || '';
        // 过期/将过期的 token 不注入：让页面自己去 /getToken 取新的（经 map-api 代理），
        // 保证「先取 token → 再取版本号 → 再请求径路」这条链上用的都是有效 token。
        if (t && jwtExpiringSoon(t)) {
          console.log('[dep-bridge] 本地保存的 token 已过期或即将过期，跳过注入，改由页面重新获取');
          t = '';
        }
        sendResponse({ token: t, ts: (r && r.authorizationTs) || 0 });
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

  // ===== 以下两类消息只允许「地图页」发起（避免其它网页借扩展的权限读内网接口）=====
  if (payload.type === 'map-need-token' || payload.type === 'map-api') {
    if (!isMapFrame(sender)) {
      console.warn('[dep-bridge] 拒绝非地图页的 ' + payload.type + ' 请求，来源：' +
                   ((sender && sender.url) || '(空)'));
      return;
    }
  }

  // 地图页索要 token：扩展去真实后端取，取到后回发该标签页（含其中的 iframe）
  if (payload.type === 'map-need-token') {
    fetchMapToken(function (err, token) {
      if (err || !token) return;
      if (fromTabId != null) send(fromTabId, { type: 'map-token', token: token });
    });
    return;
  }

  // 地图接口代理：扩展代发后端请求（绕过 file:// 的 CORS 预检与空 Referer）
  if (payload.type === 'map-api') {
    proxyMapApi(payload, fromTabId);
    return;
  }

  /* 计算器页「刷新报表页」按钮：
   * 给「匹配 popup 报表地址」的标签页下刷新指令（chrome.tabs.reload）。
   * 网页本身没有 tabs 权限，必须由扩展代劳。 */
  if (payload.type === 'reloadReport') {
    chrome.tabs.query({}, function (tabs) {
      var n = 0;
      (tabs || []).forEach(function (t) {
        if (!isReport(t.url)) return;
        try { chrome.tabs.reload(t.id, { bypassCache: false }); n++; }
        catch (e) { /* 忽略个别标签刷新失败 */ }
      });
      console.log('[dep-bridge] 刷新报表页：命中 ' + n + ' 个（匹配 ' + (cfg.reportUrl || '未配置') + '）');
      if (fromTabId != null) {
        send(fromTabId, { type: 'reloadResult', count: n, reportUrl: cfg.reportUrl || '' });
      }
    });
    return;
  }

  /* 发车填表：直接用 chrome.scripting.executeScript 注入 MAIN world 执行。
   * 不再走「常驻 content script + postMessage 转发」那条 5 跳链路
   * （计算器页 → content → background → 报表页 content → page-fill），
   * 任一刻断开都只能静默失败、且难以定位。
   * executeScript 一次性注入，并把每个 frame 的返回值直接带回，链路最短、结果最确定。 */
  if (payload.type === 'fill' || payload.type === 'fillByStrategy') {
    fillByExecute(payload, fromTabId);
    return;
  }

  /* 独立读表：不写任何单元格，只把报表页的表格同步回计算器页。
   * 与「填表」「自动提交」完全解耦 —— 开关关着也照样回执，填表失败也照样回执。 */
  if (payload.type === 'readTable') {
    readTableByExecute(payload, fromTabId);
    return;
  }

  // 其余类型（checiList / filled / reloadResult 等回传）：只发给非报表页（即计算器页面）
  chrome.tabs.query({}, function (tabs) {
    (tabs || []).forEach(function (t) {
      if (t.id === fromTabId) return;               // 不发回来源标签页，避免回环
      if (!isReport(t.url)) send(t.id, payload);
    });
  });
});

/* ====== 发车填表：MAIN world 注入执行 =====================================
 * depFillFunc 会被 chrome.scripting.executeScript 序列化后注入目标页面执行，
 * 所以它【必须自包含】——不能引用本文件作用域里的任何变量（CHANNEL / cfg 等）。
 * 注入到 MAIN world 才能访问 iframe.contentWindow.contentPane（帆软运行时对象）。 */
async function depFillFunc(cells, autoSubmit) {
  /** 元素是 iframe 且其 contentWindow 上有可用的帆软 contentPane → 返回该 contentWindow */
  function paneWin(el) {
    try {
      var w = el && el.contentWindow;
      if (!w) return null;
      var cp = w.contentPane;
      if (cp && typeof cp.setCellValue === 'function') return w;
    } catch (e) { /* 跨域访问 contentWindow.contentPane 会抛错 */ }
    return null;
  }
  /** 探测候选：元素本身是 iframe，或它内部（一层）的 iframe */
  function probe(el) {
    var own = paneWin(el);
    if (own) return own;
    try {
      var inner = el.querySelectorAll('iframe');
      for (var i = 0; i < inner.length; i++) {
        var w2 = paneWin(inner[i]);
        if (w2) return w2;
      }
    } catch (e) {}
    return null;
  }
  /* 目标定位：逐级匹配报表 iframe 的候选选择器（由最精确到最宽松）。
   * 全部落空时，下面还有「本帧即报表页」的兜底 —— 所以单一策略同时覆盖：
   *   ① 报表挂在平台页 iframe 里 → 逐级匹配命中
   *   ② 直接打开报表页本身       → 走兜底分支
   * 旧参数值（self / all / fs_tab_id 等）一律按这套处理，向后兼容。 */
  var SELECTORS = ['.fs-tab-content-item.fs-tab-content-toolbar', 'iframe.fs-tab-content-toolbar',
                   'iframe[id^="fs_tab"]', 'iframe[name^="fs_tab"]', '.fs-tab-content-item', 'iframe'];
  /* 读回报表表格：左半冻结列（#frozen-west）的数据行，col 1~5 = 股道/编组车次/辆数/换长/尾车车号。
     不同报表结构可能有差异，读不到就返回空数组，不影响填表结果本身。 */
  function readTable(win) {
    var out = [];
    try {
      var doc = win.document;
      var box = doc.querySelector('#frozen-west') || doc.querySelector('.frozen-west') || doc;
      var rows = box.querySelectorAll('tbody tr');
      for (var i = 0; i < rows.length && out.length < 20; i++) {
        var tds = rows[i].querySelectorAll('td[col]');
        if (!tds.length) continue;
        var rec = [];
        for (var j = 0; j < tds.length; j++) {
          var col = parseInt(tds[j].getAttribute('col'), 10);
          if (col >= 1 && col <= 5) {
            rec.push((tds[j].textContent || '').replace(/\s+/g, ' ').trim());
          }
        }
        if (rec.length === 5) out.push(rec);
      }
    } catch (e) { /* 跨域 / 结构不同：忽略 */ }
    return out;
  }

  /* 代点报表页的「提交」按钮（等同手工点一次）：
     #fr-btn-Submit > div > em > button —— 报表页自己会处理提交动作 */
  function clickSubmit(win) {
    try {
      var btn = win.document.querySelector('#fr-btn-Submit > div > em > button');
      if (!btn) return false;
      btn.click();
      return true;
    } catch (e) { return false; }
  }

  async function writeInto(win, label) {
    var ok = 0, failed = [];
    for (var id in cells) {
      if (!Object.prototype.hasOwnProperty.call(cells, id)) continue;
      try {
        win.contentPane.setCellValue(id, null, cells[id]);
        ok++;
      } catch (e) { failed.push(id + ':' + (e.message || e)); }
    }
    var submitted = false;
    if (ok > 0 && autoSubmit) {
      submitted = clickSubmit(win);
      // 等报表把提交结果落进表格（本地测试页是同步插行，真实报表可能要一个往返）
      if (submitted) await new Promise(function (r) { setTimeout(r, 450); });
    }
    return {
      ok: ok, failed: failed, target: label, url: location.href,
      submitted: submitted, table: readTable(win)
    };
  }

  // ① 在本 frame 的 document 里找带 contentPane 的报表 iframe（逐级匹配）
  var sels = SELECTORS;
  for (var s = 0; s < sels.length; s++) {
    var list;
    try { list = document.querySelectorAll(sels[s]); } catch (e) { continue; }
    for (var i2 = 0; i2 < list.length; i2++) {
      var w = probe(list[i2]);
      if (w) return writeInto(w, sels[s] + '[' + i2 + ']');
    }
  }

  // ③ 兜底：本 frame 自己就是帆软报表页 → 直接写本帧。
  //   典型场景：不经平台外壳、直接打开报表页本身（如直接打开 departure_flow.html），
  //   此时页面里根本没有 .fs-tab-* 这类 iframe，找 iframe 必然落空，用本帧 contentPane 才对。
  if (window.contentPane && typeof window.contentPane.setCellValue === 'function') {
    return writeInto(window, 'self(本页即报表页)');
  }

  // 本 frame 没有可用目标：无 iframe 的 frame 标记 skip（由调用方忽略，避免多 frame 互相覆盖）
  var hasIframe = false;
  try { hasIframe = !!document.querySelector('iframe'); } catch (e) {}
  return {
    ok: 0, skip: !hasIframe, target: '', url: location.href,
    error: hasIframe
      ? '本 frame 的 iframe 里没有 contentPane（报表可能未加载完，请刷新报表页后重试）'
      : '本 frame 无 iframe'
  };
}

/* ====== 独立读表：只同步报表页表格，不写任何单元格 ========================
 * 与「填表」「自动提交」完全解耦：任何时刻都能单独调用，回执类型 tableData。
 * 同样必须注入 MAIN world —— 只有 MAIN world 才访问得到报表 iframe 的 document。 */
function depReadFunc() {
  /** 元素是 iframe 且其 contentWindow 上有可用的帆软 contentPane → 返回该 contentWindow */
  function paneWin(el) {
    try {
      var w = el && el.contentWindow;
      if (!w) return null;
      var cp = w.contentPane;
      if (cp && typeof cp.setCellValue === 'function') return w;
    } catch (e) {}
    return null;
  }
  /** 探测候选：元素本身是 iframe，或它内部（一层）的 iframe */
  function probe(el) {
    var own = paneWin(el);
    if (own) return own;
    try {
      var inner = el.querySelectorAll('iframe');
      for (var i = 0; i < inner.length; i++) {
        var w2 = paneWin(inner[i]);
        if (w2) return w2;
      }
    } catch (e) {}
    return null;
  }
  /** 读表格：左半冻结列（#frozen-west）的数据行，col 1~5 = 股道/编组车次/辆数/换长/尾车车号 */
  function readTable(win) {
    var out = [];
    try {
      var doc = win.document;
      var box = doc.querySelector('#frozen-west') || doc.querySelector('.frozen-west') || doc;
      var rows = box.querySelectorAll('tbody tr');
      for (var i = 0; i < rows.length && out.length < 20; i++) {
        var tds = rows[i].querySelectorAll('td[col]');
        if (!tds.length) continue;
        var rec = [];
        for (var j = 0; j < tds.length; j++) {
          var col = parseInt(tds[j].getAttribute('col'), 10);
          if (col >= 1 && col <= 5) {
            rec.push((tds[j].textContent || '').replace(/\s+/g, ' ').trim());
          }
        }
        if (rec.length === 5) out.push(rec);
      }
    } catch (e) {}
    return out;
  }

  /* 与 depFillFunc 用同一套候选选择器：逐级匹配，落空时下面还有「本帧即报表页」兜底 */
  var SELECTORS = ['.fs-tab-content-item.fs-tab-content-toolbar', 'iframe.fs-tab-content-toolbar',
                   'iframe[id^="fs_tab"]', 'iframe[name^="fs_tab"]', '.fs-tab-content-item', 'iframe'];

  function pack(win, label) {
    return { ok: 1, target: label, url: location.href, table: readTable(win) };
  }

  // ① 在本 frame 的 document 里找带 contentPane 的报表 iframe（逐级匹配）
  var sels = SELECTORS;
  for (var s = 0; s < sels.length; s++) {
    var list;
    try { list = document.querySelectorAll(sels[s]); } catch (e) { continue; }
    for (var i2 = 0; i2 < list.length; i2++) {
      var w = probe(list[i2]);
      if (w) return pack(w, sels[s] + '[' + i2 + ']');
    }
  }

  // ③ 兜底：本 frame 自己就是报表页
  if (window.contentPane) return pack(window, 'self(本页即报表页)');

  var hasIframe = false;
  try { hasIframe = !!document.querySelector('iframe'); } catch (e) {}
  return {
    ok: 0, skip: !hasIframe, table: [], url: location.href,
    error: hasIframe ? '本 frame 的 iframe 里没有 contentPane（报表可能未加载完）' : '本 frame 无 iframe'
  };
}

/** 向所有候选标签页注入 depReadFunc，取「行数最多」的结果回传（tableData） */
function readTableByExecute(payload, fromTabId) {
  var strategy = payload.strategy || 'all';
  function reply(obj) { if (fromTabId != null) send(fromTabId, obj); }

  chrome.tabs.query({}, function (tabs) {
    var targets = (tabs || []).filter(function (t) { return isReport(t.url); });
    if (!targets.length && cfg.broadcast) {
      targets = (tabs || []).filter(function (t) {
        if (t.id === fromTabId) return false;
        if (!t.url || /^(chrome|chrome-extension|edge|about|devtools|view-source|file):/i.test(t.url)) return false;
        return true;
      });
    }
    if (!targets.length) {
      reply({ type: 'tableData', ok: 0, table: [], error: '未找到报表标签页' });
      return;
    }

    var pending = targets.length;
    var best = null;
    function settle() {
      if (--pending > 0) return;
      var n = best && best.table ? best.table.length : 0;
      console.log('[dep-bridge] 读表完成：' + n + ' 行，target=' + (best ? best.target : '(无)'));
      reply({
        type: 'tableData',
        ok: best ? best.ok : 0,
        table: (best && best.table) || [],
        target: best ? best.target : '',
        error: best ? best.error : '没有任何 frame 返回表格'
      });
    }

    targets.forEach(function (t) {
      try {
        chrome.scripting.executeScript({
          target: { tabId: t.id, allFrames: true },
          world: 'MAIN',
          func: depReadFunc,
          args: []
        }, function (results) {
          (results || []).forEach(function (r) {
            var v = r && r.result;
            if (!v || !v.table) return;
            // 取「行数最多」的那份：真正的报表帧才读得到数据行，空帧会被自然淘汰
            if (!best || v.table.length > ((best.table && best.table.length) || 0)) best = v;
          });
          settle();
        });
      } catch (e) { settle(); }
    });
  });
}

/** 向所有候选标签页的「所有 frame」注入 depFillFunc（MAIN world），取最佳结果回传计算器页 */
function fillByExecute(payload, fromTabId) {
  var cells = payload.cells || {};
  var strategy = payload.strategy || 'all';
  var autoSubmit = !!payload.autoSubmit;      // 写入成功后是否代点报表页的「提交」按钮

  function reply(obj) { if (fromTabId != null) send(fromTabId, obj); }

  chrome.tabs.query({}, function (tabs) {
    // 优先「匹配 popup 报表地址」的标签页；一个都没匹配且开着广播兜底时，注入所有普通标签页
    var targets = (tabs || []).filter(function (t) { return isReport(t.url); });
    if (!targets.length && cfg.broadcast) {
      targets = (tabs || []).filter(function (t) {
        if (t.id === fromTabId) return false;
        if (!t.url || /^(chrome|chrome-extension|edge|about|devtools|view-source|file):/i.test(t.url)) return false;
        return true;
      });
      console.log('[dep-bridge] 未匹配报表页，兜底注入所有普通标签页（' + targets.length + ' 个）');
    }
    if (!targets.length) {
      console.log('[dep-bridge] 无可用目标标签页（报表地址：' + (cfg.reportUrl || '未配置') + '）');
      reply({ type: 'filled', ok: 0, strategy: strategy,
              error: '未找到报表标签页：请检查 popup 里的「报表地址」是否与已打开的页面一致' });
      return;
    }

    var pending = targets.length;
    var best = null;
    function settle() {
      if (--pending > 0) return;
      console.log('[dep-bridge] 注入完成：ok=' + (best ? best.ok : 0) + ' target=' + (best ? best.target : '(无)'));
      reply({
        type: 'filled',
        ok: best ? best.ok : 0,
        failed: best && best.failed,
        error: best ? best.error : '注入后没有任何 frame 返回结果',
        strategy: strategy,
        target: best ? best.target : '',
        submitted: best ? !!best.submitted : false,         // 是否代点了提交
        table: best ? (best.table || null) : null           // 报表页表格内容（供计算器页展示）
      });
    }

    targets.forEach(function (t) {
      try {
        chrome.scripting.executeScript({
          target: { tabId: t.id, allFrames: true },
          world: 'MAIN',
          func: depFillFunc,
          args: [cells, autoSubmit]
        }, function (results) {
          if (chrome.runtime.lastError) {
            console.log('[dep-bridge] 注入失败 tab ' + t.id + '：' + chrome.runtime.lastError.message);
          }
          (results || []).forEach(function (r) {
            var v = r && r.result;
            if (!v) return;
            // 取「写得最多」的 frame；ok 相同则优先取带回表格数据的（避免选到跳过帧的空结果）
            if (!best ||
                (v.ok || 0) > (best.ok || 0) ||
                ((v.ok || 0) === (best.ok || 0) &&
                 !(best.table && best.table.length) && v.table && v.table.length)) {
              best = v;
            }
          });
          settle();
        });
      } catch (e) {
        console.log('[dep-bridge] 注入异常 tab ' + t.id + '：' + (e && e.message || e));
        settle();
      }
    });
  });
}