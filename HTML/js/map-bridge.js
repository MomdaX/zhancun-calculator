/**
 * map-bridge.js —— 股道存车 ⇄ 铁路地图 桥接层
 * ============================================================================
 * 通过 iframe 嵌入 Map/chinamap.html，用 postMessage 调用其径路功能。
 * Map 侧只需在 chinamap.html 末尾多一行 integration.js 引用，业务代码零改动。
 *
 * 用法（起点固定为钦州港，由 Map 内部决定）：
 *   MapBridge.open({ to: '郑屯', title: '13道 车号1234567' });
 *
 * 依赖：页面中存在 #mapMask 与 #mapFrame 两个元素（见 index.html）
 * ============================================================================
 */
(function (global) {
  'use strict';

  var POST_ORIGIN = '*';      // file:// 下 origin 为 'null'
  var READY_TIMEOUT = 60000;  // Map 初始化超时（毫秒）

  var el = {};                // 缓存 DOM
  var state = {
    ready: false,
    pending: null,            // 就绪前待发的径路请求
    openCount: 0,
    stations: 0,
    loadingShown: false
  };
  var listeners = [];

  function $(id) {
    if (!el[id]) el[id] = document.getElementById(id);
    return el[id];
  }

  /** 创建 iframe（懒加载：首次打开地图时才建，避免拖慢主页面启动） */
  function ensureFrame() {
    var frame = $('mapFrame');
    if (frame.getAttribute('data-src-done')) return frame;
    frame.setAttribute('data-src-done', '1');
    frame.src = 'Map/chinamap.html';
    return frame;
  }

  /** 打开地图并生成径路 */
  function open(opt) {
    opt = opt || {};
    var frame = ensureFrame();

    if ($('mapTitle')) {
      $('mapTitle').textContent = opt.title || ('径路：钦州港 → ' + opt.to);
    }

    showMask(true);
    sendPath(opt.to);
  }

  function close() { showMask(false); }

  function showMask(show) {
    var mask = $('mapMask');
    if (!mask) return;
    mask.style.display = show ? 'flex' : 'none';
  }

  /** 发送径路请求；若 Map 未就绪则暂存，就绪后自动补发 */
  function sendPath(to) {
    var req = { type: 'map:path', to: to || '' };
    if (!state.ready) {
      state.pending = req;
      showLoading(true, '地图初始化中，请稍候…');
      return;
    }
    post(req);
  }

  function post(obj) {
    var frame = $('mapFrame');
    try {
      if (frame && frame.contentWindow) frame.contentWindow.postMessage(obj, POST_ORIGIN);
    } catch (e) { /* 忽略 */ }
  }

  function showLoading(show, text) {
    var box = $('mapLoading');
    if (!box) return;
    if (text) {
      var t = $('mapLoadingText');
      if (t) t.textContent = text;
    }
    box.style.display = show ? 'flex' : 'none';
    state.loadingShown = show;
  }

  function setStatus(text, isError) {
    var s = $('mapStatus');
    if (!s) return;
    s.textContent = text || '';
    s.style.color = isError ? '#c0392b' : '#636e72';
  }

  /* ---------------- 接收 Map 回传 ---------------- */
  function onMessage(e) {
    var m = e.data;
    if (!m || typeof m !== 'object') return;

    switch (m.type) {
      case 'map:ready':
        state.ready = true;
        showLoading(false);
        setStatus('地图已就绪，起点：钦州港');
        if (state.pending) {
          post(state.pending);
          state.pending = null;
        }
        break;

      case 'map:pong':
        state.ready = m.ready;
        break;

      case 'map:result':
        if (m.msg === 'cleared') { setStatus('已清除径路'); break; }
        showLoading(false);
        if (m.ok) {
          setStatus('已生成径路（起点：钦州港）');
        } else {
          setStatus('径路生成失败：' + (m.msg || '未知原因'), true);
        }
        break;
    }

    listeners.forEach(function (fn) {
      try { fn(m); } catch (err) {}
    });
  }

  /* ---------------- 超时保护 ---------------- */
  function startWatchdog() {
    setTimeout(function () {
      if (!state.ready && $('mapMask') && $('mapMask').style.display === 'flex') {
        showLoading(false);
        setStatus('地图加载超时，请检查 Map 目录是否完整', true);
      }
    }, READY_TIMEOUT);
  }

  /* ---------------- 初始化 ---------------- */
  function init() {
    window.addEventListener('message', onMessage);

    var btnClose = $('mapClose');
    if (btnClose) btnClose.addEventListener('click', close);

    var btnClear = $('mapClear');
    if (btnClear) btnClear.addEventListener('click', function () {
      post({ type: 'map:clear' });
    });

    // 点击遮罩空白处关闭（不干扰 iframe 内交互）
    var mask = $('mapMask');
    if (mask) {
      mask.addEventListener('mousedown', function (e) {
        if (e.target === mask) close();
      });
    }

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape' && mask && mask.style.display === 'flex') close();
    });

    startWatchdog();
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else init();

  global.MapBridge = {
    open: open,
    close: close,
    isReady: function () { return state.ready; },
    onMessage: function (fn) { if (typeof fn === 'function') listeners.push(fn); }
  };
})(window);
