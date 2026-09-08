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
    } else if (d.type === 'fillByStrategy' && d.cells) {
      /* 发送按钮用的「按 strategy 选目标」填表：
       *   - 'fs_tab_id' / 'first_iframe' / 'fs_tab_class' / 'name_fs_tab'：仅顶层 frame 处理，
       *     在本页 document 里按对应选择器找 iframe，调其 contentWindow.contentPane.setCellValue
       *   - 'self'：仅报表 iframe frame 自身有 contentPane 时处理
       *   - 'all'：仅顶层 frame 处理（兜底按多种选择器逐个尝试）
       * 帧角色判定避免多 frame 重复执行 setCellValue。 */
      var strategy = d.strategy || 'all';
      console.log('[dep-bridge] fillByStrategy 收到', 'strategy=' + strategy, 'isTop=' + (window === window.top), 'selfHasContentPane=' + !!(window.contentPane), 'cells=' + JSON.stringify(d.cells));
      var targetWin = null;
      var targetLabel = '';
      var skipReason = '';

      /* 遍历「所有」匹配 selector 的 iframe，取第一个真的带 contentPane 的。
       * 必须遍历而不是只取第一个：FS 平台页里有 2 个 iframe，
       * 第 1 个是 display:none 的首页占位（formlet=demo/homepage/finereport.frm，无 contentPane），
       * 第 2 个才是报表（id/name=fs_tab_xxx，class 多了 fs-tab-content-toolbar）。 */
      function trySelect(selector, label) {
        try {
          var list = document.querySelectorAll(selector);
          console.log('[dep-bridge] trySelect "' + selector + '" → 匹配到 ' + list.length + ' 个');
          for (var k = 0; k < list.length; k++) {
            var f = list[k];
            try {
              var cwOk = !!(f && f.contentWindow);
              var cpOk = !!(cwOk && f.contentWindow.contentPane &&
                            typeof f.contentWindow.contentPane.setCellValue === 'function');
              console.log('[dep-bridge]   [' + k + '] src="' + String(f.src || '').slice(0, 70) +
                          '" contentWindow=' + (cwOk ? 'yes' : 'no') +
                          ' contentPane=' + (cpOk ? 'yes' : 'no') +
                          ' display=' + (f.style && f.style.display ? f.style.display : '(visible)'));
              if (cpOk) {
                targetWin = f.contentWindow;
                targetLabel = label + '[' + k + ']';
                return true;
              }
            } catch (e) {
              console.log('[dep-bridge]   [' + k + '] 访问抛错（多半是跨域）：' + (e.message || e));
            }
          }
        } catch (e) {}
        return false;
      }

      if (strategy === 'self') {
        if (window !== window.top && window.contentPane &&
            typeof window.contentPane.setCellValue === 'function') {
          targetWin = window;
          targetLabel = 'self';
        } else {
          skipReason = 'self: this frame has no contentPane';
        }
      } else if (strategy === 'all') {
        // 兜底只在顶层 frame 跑（避免和 self 重复）；按「报表 iframe 特征」优先级从高到低逐个试
        if (window === window.top) {
          if (!trySelect('iframe.fs-tab-content-toolbar', 'iframe.fs-tab-content-toolbar') &&
              !trySelect('iframe[id^="fs_tab"]', 'iframe[id^="fs_tab"]') &&
              !trySelect('iframe[name^="fs_tab"]', 'iframe[name^="fs_tab"]') &&
              !trySelect('iframe.fs-tab-content-item', 'iframe.fs-tab-content-item') &&
              !trySelect('iframe', 'querySelector("iframe")')) {
            skipReason = 'all: no contentPane iframe found';
          }
        } else {
          skipReason = 'all: only top frame handles';
        }
      } else if (strategy === 'fs_tab_toolbar') {
        // 报表 iframe 独有的 class（第 1 个首页占位 iframe 没有这个 class），最精准
        if (window === window.top) {
          if (!trySelect('iframe.fs-tab-content-toolbar', 'iframe.fs-tab-content-toolbar')) {
            skipReason = 'fs_tab_toolbar: no match';
          }
        } else skipReason = 'fs_tab_toolbar: only top frame handles';
      } else if (strategy === 'fs_tab_id') {
        if (window === window.top) {
          if (!trySelect('iframe[id^="fs_tab"]', 'iframe[id^="fs_tab"]')) {
            skipReason = 'fs_tab_id: no match';
          }
        } else skipReason = 'fs_tab_id: only top frame handles';
      } else if (strategy === 'first_iframe') {
        if (window === window.top) {
          if (!trySelect('iframe', 'querySelector("iframe")')) {
            skipReason = 'first_iframe: no match';
          }
        } else skipReason = 'first_iframe: only top frame handles';
      } else if (strategy === 'fs_tab_class') {
        if (window === window.top) {
          if (!trySelect('iframe.fs-tab-content-item', 'iframe.fs-tab-content-item')) {
            skipReason = 'fs_tab_class: no match';
          }
        } else skipReason = 'fs_tab_class: only top frame handles';
      } else if (strategy === 'name_fs_tab') {
        if (window === window.top) {
          if (!trySelect('iframe[name^="fs_tab"]', 'iframe[name^="fs_tab"]')) {
            skipReason = 'name_fs_tab: no match';
          }
        } else skipReason = 'name_fs_tab: only top frame handles';
      } else {
        skipReason = 'unknown strategy: ' + strategy;
      }

      if (!targetWin) {
        console.log('[dep-bridge] ✗ 没找到目标：' + (skipReason || 'no target') + '（strategy=' + strategy + '）');
        post({ type: 'filled', ok: 0, url: location.href, strategy: strategy, target: targetLabel, error: skipReason || 'no target' });
        return;
      }
      console.log('[dep-bridge] ✓ 选中目标：' + targetLabel + '，isTargetSelf=' + (targetWin === window) + '，hasContentPane=' + !!(targetWin.contentPane));

      try {
        // 行激活：真实帆软模板里 setCellValue 必须先「选中行」才生效（否则提交报"请输入内容"）
        // 用 td[idx="0"] 而非 [id^="B4-"]：不依赖 id 模式，DOM 顺序第一个 idx="0" 的 td
        // 点它能激活"第一条新增行"，让后续 setCellValue 落到 dirty 状态、能过校验
        try {
          var b4Doc = (targetWin === window) ? document : targetWin.document;
          var firstIdx = b4Doc && b4Doc.querySelector('td[idx="0"]');
          console.log('[dep-bridge] 行激活 ' + (firstIdx
            ? '找到 td[idx="0"]（id=' + (firstIdx.id || '?') + '），已 click()'
            : '未找到 td[idx="0"]'));
          if (firstIdx) firstIdx.click();
        } catch (e) { /* 行激活失败不影响后续写入 */ }

        var ok = 0, failed = [];
        var keys = Object.keys(d.cells);
        for (var ki = 0; ki < keys.length; ki++) {
          var id = keys[ki];
          try {
            targetWin.contentPane.setCellValue(id, null, d.cells[id]);
            console.log('[dep-bridge] setCellValue ' + id + ' = ' + JSON.stringify(d.cells[id]) + ' → ok');
            ok++;
          } catch (e) {
            console.log('[dep-bridge] setCellValue ' + id + ' → 抛错：' + (e.message || e));
            failed.push(id + ':' + (e.message || e));
          }
        }
        console.log('[dep-bridge] 完成：ok=' + ok + '/5, strategy=' + strategy + ', target=' + targetLabel);
        post({
          type: 'filled', ok: ok, failed: failed,
          url: location.href, strategy: strategy, target: targetLabel
        });
        // 填完顺手回传一次车次列表（与 fill 通道保持一致）
        if (targetWin.document && targetWin.document.querySelector('td[col="2"]')) {
          // 在目标 win 的文档里抓车次（与 collectCheci 同样的逻辑）
          try {
            var tds = targetWin.document.querySelectorAll('td[col="2"]');
            var checiOut = [];
            for (var ti = 0; ti < tds.length; ti++) {
              var tid = tds[ti].getAttribute('id') || '';
              if (tid.indexOf('C4-') === 0) continue;
              var tt = (tds[ti].textContent || '').replace(/\s+/g, '');
              if (tt) checiOut.push(tt);
            }
            post({ type: 'checiList', list: checiOut, url: targetWin.location.href });
          } catch (e) {}
        }
      } catch (e) {
        post({ type: 'filled', ok: 0, url: location.href, strategy: strategy, target: targetLabel, error: e.message });
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
