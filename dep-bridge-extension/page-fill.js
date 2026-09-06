/* ============================================================================
 * 站存-发车流程填表桥 —— 页面脚本（MAIN world，注入报表页及其内部 iframe）
 *
 * 真正执行填表的部分：
 *   {type:'fill',      cells:{B4:.., C4:.., D4:.., E4:.., F4:..}}  写单元格
 *   {type:'readCheci'}                                              读「编组车次」列
 * 写完 / 页面加载完会回传 {type:'filled'} / {type:'checiList'}，供计算器页查重。
 *
 * 单元格定位沿用帆软规则：id 形如 C4-0-73，按前缀查找。
 * 有 contentPane（帆软填报）优先走官方 API，否则直接写 DOM；
 * 若格内是输入控件（填报 widget），同步写控件值并派发 input/change 事件。
 * ==========================================================================*/
'use strict';

(function () {
  if (window.__depFillInstalled) return;
  window.__depFillInstalled = true;

  var CHANNEL = '__DEP_BRIDGE__';
  var CELLS = ['B4', 'C4', 'D4', 'E4', 'F4'];   // 股道 / 车次 / 辆数 / 换长 / 尾车车号

  function findCell(id) {
    try { return document.querySelector('[id^="' + id + '-"]'); } catch (e) { return null; }
  }

  /** 同步格内输入控件（帆软填报控件常是 input / textarea），并触发事件让控件记住值 */
  function syncWidget(td, val) {
    var ok = false;
    try {
      var els = td.querySelectorAll('input, textarea, select');
      for (var i = 0; i < els.length; i++) {
        var el = els[i];
        if (el.type === 'checkbox' || el.type === 'radio') continue;
        el.value = (val == null) ? '' : val;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
        ok = true;
      }
    } catch (e) {}
    return ok;
  }

  function setCell(id, val) {
    var td = findCell(id);
    if (!td) return false;
    try {
      if (window.contentPane && typeof window.contentPane.setCellValue === 'function') {
        window.contentPane.setCellValue(id, null, val);
        syncWidget(td, val);
        return true;
      }
    } catch (e) { /* 退回 DOM 写法 */ }
    // 纯 DOM 降级：清空格子放纯文本
    td.innerHTML = '';
    var d = document.createElement('div');
    d.style.maxHeight = '28px';
    d.textContent = (val == null) ? '' : val;
    td.appendChild(d);
    syncWidget(td, val);
    return true;
  }

  /** 读「编组车次」列（td[col="2"]），排除正在填写的 C4 */
  function collectCheci() {
    var out = [];
    try {
      var tds = document.querySelectorAll('td[col="2"]');
      for (var i = 0; i < tds.length; i++) {
        var id = tds[i].getAttribute('id') || '';
        if (id.indexOf('C4-') === 0) continue;
        var t = (tds[i].textContent || '').replace(/\s+/g, '');
        if (t) out.push(t);
      }
    } catch (e) {}
    return out;
  }

  function post(msg) {
    msg.channel = CHANNEL;
    window.postMessage(msg, '*');
  }

  /* ====== 地图 token 注入（chinamap.html） ================================
   * 地图自己请求 /getToken 拿不到真 token 时，由扩展代为请求真实后端，
   * 拿到后回传 {type:'map-token'}，这里写入 sessionStorage 并让地图重新加载。
   * 地图的 map.js / jl.js 里每个 $.ajax 都会带
   *   setRequestHeader("Authorization", sessionStorage.token)
   * 所以 token 写对，整个地图的 API 就都带上了正确的 Authorization。 */
  function isMapPage() {
    return !!document.getElementById('status_token');       // chinamap.html 特有节点
  }

  function mapTokenValid() {
    try {
      var t = sessionStorage.getItem('token');
      return !!t && t !== '3231212';        // 3231212 是 getToken.js 失败时写的兜底假 token
    } catch (e) { return false; }
  }

  function applyMapToken(token) {
    try {
      sessionStorage.setItem('token', token);
      window.sessionStorage.token = token;
      window.apiAvailable = true;           // 关键：让 map.js / jl.js 改走真实 API
      var dot = document.getElementById('status_token');
      if (dot) dot.className = 'dot ok';    // 状态灯转绿
      // 重新拉版本 + 初始化地图（initMap 内有防重入，先清掉 map 实例才会重新构建）
      try { if (typeof map !== 'undefined') window.map = null; } catch (e) {}
      try { if (typeof getVersion === 'function') getVersion(); } catch (e) {}
    } catch (e) {}
  }

  // 收到填表 / 读取指令（由 content.js 转进本页）
  window.addEventListener('message', function (ev) {
    if (ev.source !== window) return;
    var d = ev.data;
    if (!d || d.channel !== CHANNEL) return;

    if (d.type === 'fill' && d.cells) {
      var ok = 0;
      for (var i = 0; i < CELLS.length; i++) {
        var id = CELLS[i];
        if (Object.prototype.hasOwnProperty.call(d.cells, id)) {
          if (setCell(id, d.cells[id])) ok++;
        }
      }
      post({ type: 'filled', ok: ok, url: location.href });
      // 填完顺手回传一次车次列表，保证计算器页的查重数据是最新的
      if (document.querySelector('td[col="2"]')) {
        post({ type: 'checiList', list: collectCheci(), url: location.href });
      }
    } else if (d.type === 'readCheci') {
      post({ type: 'checiList', list: collectCheci(), url: location.href });
    } else if (d.type === 'map-token' && d.token) {
      applyMapToken(d.token);               // 扩展代取的真 token 落地
    }
  });

  // 报表页加载完主动上报一次车次列表，让计算器页能做「车次重复」提示
  function report() {
    if (document.querySelector('td[col="2"]')) {
      post({ type: 'checiList', list: collectCheci(), url: location.href });
    }
  }
  if (document.readyState === 'complete') setTimeout(report, 600);
  else window.addEventListener('load', function () { setTimeout(report, 600); });

  // 地图页兜底：页面自己的 GetToken 若没拿到真 token，请扩展代为获取
  setTimeout(function () {
    if (isMapPage() && !mapTokenValid()) post({ type: 'map-need-token' });
  }, 1500);                                  // 等页面自身的 /getToken 先跑完再判断
})();
