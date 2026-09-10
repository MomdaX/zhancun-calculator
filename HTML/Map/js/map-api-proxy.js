/* ============================================================================
 * map-api-proxy.js —— 地图后端接口代理（页面侧，MAIN world）
 *
 * 背景：地图页以 file:// 打开（origin 为 null），直连后端会：
 *   1. 因带 Authorization 自定义头而触发 CORS 预检，而后端不处理 OPTIONS，
 *      预检返回非 2xx → 请求被浏览器拦下（见 1005.log）；
 *   2. file:// 页面发出的请求 Referer 恒为空，且 Referer 属于 forbidden header，
 *      页面 JS 无权设置，无法从页面侧修复。
 *
 * 方案：把地图的全部后端接口请求交给浏览器扩展代发（扩展有 host_permissions，
 * 不受 CORS 限制，且可在 fetch 上直接带 Referer / Authorization），结果回传后
 * 由本文件按 jQuery 的 success / error 语义唤起原回调。
 *
 * 链路：
 *   map-api-proxy.js（本文件，覆盖 $.ajax）
 *     → window.postMessage
 *     → content.js（ISOLATED world）转 chrome.runtime
 *     → background.js  fetch(base + path, { Authorization, Referer })
 *     → chrome.tabs.sendMessage 回传 { type:'map-api-result', reqId, status, body }
 *     → content.js postMessage 回页面
 *     → 本文件按 reqId 找到挂起请求，解析后调 success / error
 *
 * 引入位置：必须紧跟 jquery.min.js（及定义了 getApiBase 的内联脚本）之后、
 *          getToken.js 之前，见 chinamap.html。
 * ==========================================================================*/
(function () {
  'use strict';
  if (window.__depApiProxyInstalled) return;
  window.__depApiProxyInstalled = true;

  var CHANNEL = '__DEP_BRIDGE__';
  // 扩展代发超时（毫秒）。内网接口通常 < 1s，这里留足余量；
  // 若扩展未安装/未响应，超时后走 error → 页面回落到本地兜底数据。
  var TIMEOUT = 12000;

  // 需要代理的接口白名单（与地图内所有 $.ajax 的 url 一一对应）
  var API_PATHS = [
    '/getToken',
    '/getReleaseVersionData',
    '/getMapDftPoint',
    '/getMapDlnPoint',
    '/getMapDndPoint',
    '/getMidStationByLj',
    '/getNode',
    '/getJl'
  ];

  var $ = window.jQuery;
  if (!$ || typeof $.ajax !== 'function') {
    console.warn('[dep-api] 未检测到 jQuery，接口代理未启用');
    return;
  }
  var rawAjax = $.ajax;

  var seq = 0;
  var pending = {};                 // reqId -> { settleOK, settleErr }

  /* ---------------- 工具 ---------------- */

  /** 取 url 的 pathname（兼容 /getXxx 与 http://host:port/getXxx） */
  function pathOf(url) {
    var u = String(url || '');
    try {
      if (/^[a-z][a-z0-9+.-]*:\/\//i.test(u)) return new URL(u).pathname;
    } catch (e) {}
    var i = u.indexOf('?');
    var p = i >= 0 ? u.slice(0, i) : u;
    return p.charAt(0) === '/' ? p : '/' + p;
  }

  function isProxyTarget(url) {
    return API_PATHS.indexOf(pathOf(url)) >= 0;
  }

  /** jQuery 的 data 归一化成普通对象（对象 / a=1&b=2 字符串 / [{name,value}] 数组） */
  function normalizeData(d) {
    if (!d) return {};
    if (typeof d === 'string') {
      var out = {};
      d.split('&').forEach(function (kv) {
        if (!kv) return;
        var i = kv.indexOf('=');
        var k = i >= 0 ? kv.slice(0, i) : kv;
        var v = i >= 0 ? kv.slice(i + 1) : '';
        try { out[decodeURIComponent(k)] = decodeURIComponent(v); } catch (e) { out[k] = v; }
      });
      return out;
    }
    if (Object.prototype.toString.call(d) === '[object Array]') {
      var o2 = {};
      d.forEach(function (it) { if (it && it.name != null) o2[it.name] = it.value; });
      return o2;
    }
    if (typeof d === 'object') {
      var o3 = {};
      Object.keys(d).forEach(function (k) { o3[k] = d[k]; });
      return o3;
    }
    return {};
  }

  /** 按 jQuery 的 dataType 智能判断把响应体还原成 JS 值 */
  function parseBody(body, contentType) {
    if (body == null || body === '') return body;
    var ct = String(contentType || '').toLowerCase();
    var looksJson = /^\s*[\[{]/.test(body);
    if (ct.indexOf('json') >= 0 || looksJson) {
      try { return JSON.parse(body); } catch (e) { return body; }
    }
    return body;
  }

  function post(msg) {
    msg.channel = CHANNEL;
    window.postMessage(msg, '*');
  }

  /* ---------------- 极简 jqXHR（页面未使用返回值，仅保证链式语义完整） ---------------- */

  function makeJqXHR() {
    var cbs = { done: [], fail: [], always: [] };
    var settled = false, kind = '', args = [];

    var jq = {
      readyState: 1, status: 0, statusText: '', responseText: '',
      done: function (fn) { return reg('done', fn); },
      fail: function (fn) { return reg('fail', fn); },
      always: function (fn) { return reg('always', fn); },
      then: function (ok, bad) { if (ok) reg('done', ok); if (bad) reg('fail', bad); return jq; },
      abort: function () { if (jq.__cancel) jq.__cancel(); }
    };

    function reg(type, fn) {
      if (typeof fn === 'function') {
        if (settled) {
          if (type === 'done' && kind === 'done') fn.apply(jq, args);
          else if (type === 'fail' && kind === 'fail') fn.apply(jq, args);
          else if (type === 'always') fn.apply(jq, args);
        } else {
          cbs[type].push(fn);
        }
      }
      return jq;
    }

    /** k 取 'done' | 'fail'；无论如何 always 都会触发 */
    jq.__fire = function (k, a) {
      if (settled) return;
      settled = true; kind = k; args = a || [];
      (cbs[k] || []).concat(cbs.always || []).forEach(function (fn) {
        try { fn.apply(jq, args); } catch (e) {}
      });
    };
    return jq;
  }

  /* ---------------- 代理一次 ajax ---------------- */

  function proxyAjax(opts) {
    var reqId = 'mapi-' + (++seq) + '-' + Date.now();
    var jq = makeJqXHR();
    var finished = false;
    var timer = null;

    function settleOK(value, meta) {
      if (finished) return;
      finished = true;
      clearTimeout(timer); delete pending[reqId];
      jq.readyState = 4;
      jq.status = (meta && meta.status) || 200;
      jq.statusText = 'OK';
      jq.responseText = (typeof value === 'string') ? value : '';
      jq.__fire('done', [value, 'success', jq]);
      if (typeof opts.success === 'function') {
        try { opts.success(value, 'success', jq); }
        catch (e) { console.error('[dep-api] success 回调异常：', e); }
      }
      if (typeof opts.complete === 'function') { try { opts.complete(jq, 'success'); } catch (e) {} }
    }

    function settleErr(status, textStatus, errText) {
      if (finished) return;
      finished = true;
      clearTimeout(timer); delete pending[reqId];
      jq.readyState = 4;
      jq.status = status || 0;
      jq.statusText = textStatus || 'error';
      jq.__fire('fail', [jq, textStatus || 'error', errText || '']);
      if (typeof opts.error === 'function') {
        try { opts.error(jq, textStatus || 'error', errText || ''); }
        catch (e) { console.error('[dep-api] error 回调异常：', e); }
      }
      if (typeof opts.complete === 'function') { try { opts.complete(jq, textStatus || 'error'); } catch (e) {} }
    }

    jq.__cancel = function () { settleErr(0, 'abort', ''); };

    // 本地数据集模式：不请求，直接走 error 让页面用本地兜底数据
    // （与 chinamap.html 里 ajaxPrefilter 的 jqXHR.abort() 语义一致）
    if (typeof isLocalMode === 'function' && isLocalMode()) {
      setTimeout(function () { settleErr(0, 'abort', ''); }, 0);
      return jq;
    }

    var base = (typeof getApiBase === 'function') ? getApiBase() : '';
    var path = pathOf(opts.url);
    if (!base) {
      setTimeout(function () { settleErr(0, 'error', '未配置后端地址'); }, 0);
      return jq;
    }

    var token = '';
    try {
      token = window.sessionStorage.getItem('token') || '';
      if (token === '3231212') token = '';       // getToken 失败时的兜底假 token，视为无
    } catch (e) {}

    pending[reqId] = { settleOK: settleOK, settleErr: settleErr };

    console.log('[dep-api] → ' + (opts.type || 'GET') + ' ' + base + path + (token ? '（带 token）' : '（无 token）'));
    post({
      type: 'map-api',
      reqId: reqId,
      method: (opts.type || 'GET').toUpperCase(),
      base: base,
      path: path,
      data: normalizeData(opts.data),
      token: token
    });

    timer = setTimeout(function () {
      console.warn('[dep-api] ✗ 超时（' + TIMEOUT + 'ms）：' + path + '（扩展未响应？）');
      settleErr(0, 'timeout', '扩展代理超时');
    }, TIMEOUT);

    return jq;
  }

  /* ---------------- 覆盖 $.ajax ---------------- */

  $.ajax = function (url, options) {
    var opts;
    if (typeof url === 'string') {
      opts = $.extend({}, options || {}, { url: url });
    } else {
      opts = $.extend({}, url || {});
    }

    // 非地图接口：原样交回 jQuery（原有的 ajaxPrefilter 仍然对它生效）
    if (!isProxyTarget(opts.url)) {
      return rawAjax.apply($, arguments);
    }

    // 代理模式下 beforeSend 里的 setRequestHeader("Authorization", ...) 不再需要：
    // 扩展代发时会统一补上 Authorization 与 Referer。
    return proxyAjax(opts);
  };

  /* ---------------- 接收扩展回传 ---------------- */

  window.addEventListener('message', function (ev) {
    if (ev.source !== window) return;
    var d = ev.data;
    if (!d || d.channel !== CHANNEL || d.type !== 'map-api-result') return;

    var p = pending[d.reqId];
    if (!p) return;                              // 超时后迟到的回包，丢弃
    delete pending[d.reqId];

    if (d.ok) {
      var value = parseBody(d.body, d.contentType);
      console.log('[dep-api] ✓ ' + d.status + ' ' + d.reqId);
      p.settleOK(value, { status: d.status });
    } else {
      console.warn('[dep-api] ✗ ' + (d.status || 0) + ' ' + d.reqId + (d.error ? (' ' + d.error) : ''));
      p.settleErr(d.status || 0, 'error', d.error || ('HTTP ' + (d.status || 0)));
    }
  });

  console.log('[dep-api] 地图接口代理已启用（' + API_PATHS.length + ' 个接口，超时 ' + TIMEOUT + 'ms）');
})();
