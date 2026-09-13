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
      /* 发送按钮用的「按 strategy 选目标」填表，现在只有两条路：
       *   - 'fs_tab_toolbar'（默认）：在本帧 document 里逐级匹配报表 iframe，调其
       *     contentWindow.contentPane.setCellValue；一个都没找到时兜底写本帧。
       *   - 'self'：本帧就是报表页（直接打开报表页的场景），直接用本帧 contentPane。
       * 不限制顶层 frame：报表 iframe 可能嵌在中间层，谁在自己的 document 里找得到谁负责写。 */
      var strategy = d.strategy || 'all';

      /* 本页是「站存计算器」自身（指令发起方）→ 直接静默跳过。
       * 计算器页自己也带 iframe（#mapFrame），会被兜底广播扫到，但它不可能是
       * 报表平台页——跑一遍定位只会打出一条容易被误判成"失败"的日志。
       * 用计算器页独有的元素判定（发送按钮 / 明细表）；报表页不会有这些。 */
      if (document.getElementById('btnSendCurrent') || document.getElementById('detailTable')) {
        return;
      }

      /* 前置过滤：本页如果没有「报表平台特征」的 iframe，就不是 FS 平台页，直接静默跳过。
       * 注意不能只看「有没有 iframe」——计算器页自己也带一个 <iframe id="mapFrame">（about:blank），
       * 广播兜底会误发到它，导致刷一堆 "no match"。这里按报表特征判断才准。 */
      if (strategy !== 'self') {
        // 两个放行条件，满足其一就继续（最终以「能否拿到 contentPane」为准）：
        //   ① 本帧有 iframe —— 可能是平台页，报表 iframe 就在里面；
        //   ② 本帧自带 contentPane —— 直接打开报表页本身（无平台外壳、也无 iframe）。
        // 两者都不满足 → 与报表无关，静默跳过（不回执，避免覆盖别的 frame 的成功回执）。
        var anyIframe = false;
        try { anyIframe = !!document.querySelector('iframe'); } catch (e) {}
        var selfPane = !!(window.contentPane && typeof window.contentPane.setCellValue === 'function');
        if (!anyIframe && !selfPane) {
          console.log('[dep-bridge] ⊘ 跳过：本 frame 无 iframe 也无 contentPane（strategy=' + strategy + '）');
          return;
        }
      }

      console.log('[dep-bridge] fillByStrategy 收到', 'strategy=' + strategy, 'isTop=' + (window === window.top), 'selfHasContentPane=' + !!(window.contentPane), 'cells=' + JSON.stringify(d.cells));
      var targetWin = null;
      var targetLabel = '';
      var skipReason = '';
      // 「本 frame 按角色本就不处理该 strategy」= 跳过（不是失败）。
      // 跳过的 frame 绝不能回传 filled，否则会用它把真正写入成功的回执覆盖成 0/5。
      var isSkip = false;

      /* 遍历「所有」匹配 selector 的 iframe，取第一个真的带 contentPane 的。
       * 必须遍历而不是只取第一个：FS 平台页里有 2 个 iframe，
       * 第 1 个是 display:none 的首页占位（formlet=demo/homepage/finereport.frm，无 contentPane），
       * 第 2 个才是报表（id/name=fs_tab_xxx，class 多了 fs-tab-content-toolbar）。 */
      /** 元素是 iframe 且其 contentWindow 上有可用的帆软 contentPane → 返回该 contentWindow，否则 null */
      function paneWin(el) {
        try {
          var w = el && el.contentWindow;
          if (!w) return null;
          var cp = w.contentPane;
          if (cp && typeof cp.setCellValue === 'function') return w;
        } catch (e) { /* 跨域访问 contentWindow.contentPane 会抛错 */ }
        return null;
      }

      function trySelect(selector, label) {
        try {
          var list = document.querySelectorAll(selector);
          console.log('[dep-bridge] trySelect "' + selector + '" → 匹配到 ' + list.length + ' 个');
          for (var k = 0; k < list.length; k++) {
            var f = list[k];
            var win = paneWin(f);
            if (win) {
              console.log('[dep-bridge]   [' + k + '] ✓ 命中（本身是 iframe）');
              targetWin = win;
              targetLabel = label + '[' + k + ']';
              return true;
            }
            // 命中的可能是「容器元素」（帆软有的版本用 div 包 iframe），
            // 在它内部一层找 iframe 再探测；用户实测的选择器就不带 iframe 前缀。
            try {
              var inner = f.querySelectorAll('iframe');
              for (var q = 0; q < inner.length; q++) {
                var w2 = paneWin(inner[q]);
                if (w2) {
                  console.log('[dep-bridge]   [' + k + '] ✓ 命中（容器内 iframe[' + q + ']）');
                  targetWin = w2;
                  targetLabel = label + '[' + k + ']>iframe[' + q + ']';
                  return true;
                }
              }
            } catch (e2) {}
          }
        } catch (e) {}
        return false;
      }

      if (strategy === 'self') {
        // 「自身」：本帧就是报表页（直接打开报表页的场景），不找 iframe
        if (window.contentPane && typeof window.contentPane.setCellValue === 'function') {
          targetWin = window;
          targetLabel = 'self';
        } else {
          isSkip = true;      // 本 frame 不是报表 frame，self 策略不归它处理
          skipReason = 'self: this frame has no contentPane';
        }
      } else {
        /* 找报表 iframe：首选实测最准的「组合类」，再逐级退回其它特征。
         * ★ 不限制 window === window.top：帆软平台页常把报表 iframe 嵌在中间层 frame 内，
         *   「谁在自己的 document 里找得到，谁就负责写」，找到的 frame 唯一、不会重复。
         * 未知 strategy（含历史按钮值 all / fs_tab_id 等）也走这套，等价于原来的兜底。 */
        if (!trySelect('.fs-tab-content-item.fs-tab-content-toolbar', 'fs_tab_toolbar') &&
            !trySelect('iframe.fs-tab-content-toolbar', 'fs_tab_toolbar') &&
            !trySelect('iframe[id^="fs_tab"]', 'fs_tab_toolbar') &&
            !trySelect('iframe[name^="fs_tab"]', 'fs_tab_toolbar') &&
            !trySelect('.fs-tab-content-item', 'fs_tab_toolbar') &&
            !trySelect('iframe', 'fs_tab_toolbar')) {
          skipReason = 'no contentPane iframe found';
        }
      }

      /* 兜底：本 frame 自己就是帆软报表页 → 直接写本帧。
       * 场景：不经平台外壳、直接打开报表页本身（如直接打开 departure_flow.html），
       * 此时页面里没有 .fs-tab-* 这类 iframe，按选择器找必然落空，
       * 但本帧就有 contentPane，直接写才对。（写入逻辑在下面对 targetWin 统一处理。） */
      if (!targetWin && strategy !== 'self' &&
          window.contentPane && typeof window.contentPane.setCellValue === 'function') {
        targetWin = window;
        targetLabel = 'self(本页即报表页)';
      }

      if (!targetWin) {
        /* 静默跳过的两种情况（都不是真失败 → 不回传 filled，避免覆盖成功回执）：
         *   1. isSkip      ：本 frame 按角色本就不处理该 strategy
         *                    （如 iframe 内的 frame 收到 top-only 策略）
         *   2. 本页无 iframe：不是 FS 平台页。广播兜底会发给所有标签页，
         *      无关页面 / 失效页面（如连不上的报表页）会走到这里，静默即可。 */
        var noIframe = (strategy !== 'self') && !document.querySelector('iframe');
        if (isSkip || noIframe) {
          console.log('[dep-bridge] ⊘ 跳过（不回执）：' +
            (noIframe ? '本页无 iframe，非报表平台页' : skipReason));
          return;
        }
        // 走到这里才是真问题：本页有 iframe，但里面没有带 contentPane 的报表
        console.log('[dep-bridge] ⨯ 本 frame 有 iframe，但里面没有带 contentPane 的帆软报表：' +
          (skipReason || 'no target') + '（strategy=' + strategy + '，本 frame=' + location.href + '）' +
          ' —— 若这不是报表标签页，属正常（兜底广播会打扰无关页面）；若是报表页，请刷新报表页后重试。');
        post({ type: 'filled', ok: 0, url: location.href, strategy: strategy, target: targetLabel, error: skipReason || 'no target' });
        return;
      }
      console.log('[dep-bridge] ✓ 选中目标：' + targetLabel + '，isTargetSelf=' + (targetWin === window) + '，hasContentPane=' + !!(targetWin.contentPane));

      try {
        /* 直接写值，无需「激活行」：
         * 已取得准确的 contentPane 对象（等价于
         *   document.querySelector('.fs-tab-content-item.fs-tab-content-toolbar')
         *     .contentWindow.contentPane）
         * 直接 cp.setCellValue(...) 即可写入，实测可行（真实网页已验证）。 */
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
