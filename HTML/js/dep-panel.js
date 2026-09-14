/**
 * dep-panel.js —— 发车作业全流程浮窗
 * ============================================================================
 * 点主表「编好」后弹出：把 5 个值（股道 / 编组车次 / 辆数 / 换长 / 尾车车号）经扩展桥
 * 写进「发车作业全流程」报表页，并把报表表格读回来展示。本模块自持全部内容：
 *   · 「编好」命中后的打开与预填（openForRow，由 grid-events 的钩子调用）
 *   · 三类警示标（车次重复 / 列车超长 / 列车超重）与车次查重
 *   · 车次查重唯一数据源 = 扩展桥同步回的「报表表格」（打开浮窗 / 发送 / 刷新后都同步）
 *   · 发送 / 读表 / 自动提交 / 刷新报表 / 回执聚合 / 时间显示 / 浮窗内状态提示
 *   · 抽屉注册（modalDeparture）与关闭按钮
 *
 * 与 app.js 的关系：模块内部自持全部状态与事件，app.js 只在 bind() 里调一次 init()，
 * 并把 openForRow 交给 grid-events 当「编好」命中后的钩子。**依赖单向、无需注入**。
 *
 * 作用域约定（拆分时踩过坑，务必遵守）：
 *   凡被模块顶层的 openForRow 直接或间接调用的函数/变量（depWarn、depWarnCtx、
 *   depDupState、depWarnList、setDepWarn、refreshDepWarnByTrack、formatDepTime、
 *   refreshDepTime、DEP_BRIDGE），都必须定义在模块顶层，不能留在 init() 内——
 *   留在 init() 内会让顶层函数抛 ReferenceError（事件回调里抛错还会被 try/catch 静默吞掉）。
 *   setDepWarn 内的局部 var on（布尔）会就近遮蔽模块顶层的 on 别名（Utils.on）——
 *   原 app.js 的既有写法，该函数内不再用 Utils.on，原样保留，勿改。
 *
 * 加载顺序：须在 app.js 之前（app.js 的 bind() 会调用 DepPanel.init）。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Utils = global.Utils;
  var $ = Utils.$;
  var on = Utils.on;
  var escapeHtml = Utils.escapeHtml;
  var Store = global.Store;
  var UI = global.UI;
  var YardConfig = global.YardConfig;
  var Checi = global.Checi;
  var COL = global.Aggregate ? global.Aggregate.COL : null;

  var inited = false;

  /* 扩展桥接通道名（dep-bridge-extension）。模块顶层定义：回执监听与指令下发都在
   * init() 里使用它，openForRow 等顶层函数也会引用——放 init() 内会引用不到
   * （ReferenceError 还会被 try/catch 静默吞掉）。 */
  var DEP_BRIDGE = '__DEP_BRIDGE__';

  var depWarn = { dup: false, overlong: false, overweight: false };
  var depWarnCtx = { length: 0, load: 0 };   // 当前股道的 换长 / 重量（点「编好」时记录）
  var depDupState = false;
  var depCurTrack = '';   // 发车浮窗当前对应的股道 id（openDepForRow 设置，车次同步落库用）
  var applyDepCheciRef = null;   // init 装配后指向 applyDepCheci：openDepForRow 预填后要主动跑一次查重
                                 // （applyDepCheci 依赖 depCheciDup 等init 内闭包，只能经此引用）

  /** 车次落库后防抖刷新主表「编好车次」列：输入中高频触发，全表重渲做 150ms 防抖。
   *  render 重建的是 #grid 的 tbody，输入焦点在浮窗输入框里，不受影响。 */
  var syncCheciRender = Utils.debounce(function () {
    if (global.YardApp && global.YardApp.renderGrid) global.YardApp.renderGrid();
  }, 150);

  /** 按固定顺序汇总当前命中的提示文案 */
  function depWarnList() {
    var list = [];
    if (depWarn.dup) list.push('当前车次重复');
    if (depWarn.overlong) list.push('列车超长');
    if (depWarn.overweight) list.push('列车超重');
    return list;
  }

  /** 刷新警示标：有提示 → 显示三角标 + 悬停列表；无提示 → 隐藏 */
  function setDepWarn() {
    var list = depWarnList();
    var w = $('depCheciWarn');
    var tip = $('depWarnTip');
    var inp = $('depInputC4');
    var on = list.length > 0;
    if (w) {
      w.classList.toggle('show', on);
      // 不再写 title：浏览器原生 tooltip 与红色气泡会同时出现（双框），保留自定义气泡即可
      if (!on) w.removeAttribute('title');
    }
    // 输入框红框只跟「车次重复」走（那是输入内容本身的问题）
    if (inp) inp.classList.toggle('dup', !!depWarn.dup);
    // 换长格同步超长红底（视觉与车次格 .dup 一致；超重无对应输入格，仅红三角提示）
    var e4 = $('depInputE4');
    if (e4) e4.classList.toggle('overlong', !!depWarn.overlong);
    if (tip) {
      tip.innerHTML = on
        ? '<ol>' + list.map(function (t, i) {
            return '<li>' + (i + 1) + '. ' + escapeHtml(t) + '</li>';
          }).join('') + '</ol>'
        : '';
    }
  }

  /** 依据当前股道的换长 / 重量刷新「超长 / 超重」提示 */
  function refreshDepWarnByTrack() {
    var thr = (YardConfig && YardConfig.thresholds) || {};
    var maxLen = thr.overlong == null ? 70 : thr.overlong;
    var maxWt = thr.overloadTons == null ? 5000 : thr.overloadTons;
    depWarn.overlong = Number(depWarnCtx.length) > maxLen;
    depWarn.overweight = Number(depWarnCtx.load) > maxWt;
    setDepWarn();
  }

  // ============== 时间显示（每秒刷新）==============
  function formatDepTime() {
    var d = new Date();
    return d.getFullYear() + '-' + Utils.pad2(d.getMonth() + 1) + '-' + Utils.pad2(d.getDate()) +
           ' ' + Utils.pad2(d.getHours()) + ':' + Utils.pad2(d.getMinutes()) + ':' + Utils.pad2(d.getSeconds());
  }
  function refreshDepTime() {
    var t = $('depTime');
    if (t) t.textContent = formatDepTime();
  }

  /** 装配：把原 bind() 内的发车浮窗实现整体跑一遍（绑定与初始化时机与拆分前一致）。 */
  function init() {
    if (inited) return;
    inited = true;

    /* ================= 发车浮窗警示标 =================
     * 三类提示（任一命中即显示红色三角标，悬停在其后方以编号列表全部列出）：
     *   1. 当前车次重复 —— 与 iframe「编组车次」列（td[col="2"]）已有车次相同
     *   2. 列车超长     —— 本股道换长 > YardConfig.thresholds.overlong（70.0）
     *   3. 列车超重     —— 本股道重量 > YardConfig.thresholds.overloadTons（5000）
     * 均为「仅提示」：不拦截填表，车次照常写入 C4。 */
    depWarn = { dup: false, overlong: false, overweight: false };
    depWarnCtx = { length: 0, load: 0 };       // 当前股道的 换长 / 重量（点「编好」时记录）
    // 注：不再保留 depCells 之类的「待填单元格」内存快照——5 个值只在 #depEntry 的 input 里，
    //     发送时由 getDepCells() 现取，关窗即丢，不做任何持久化。

    /* 扩展桥接（dep-bridge-extension）：页面 → content script → background → 报表标签页。
     * 跨域可用（不要求同源），前提：浏览器装了该扩展并已启用。
     * 未装扩展时相关指令静默无效。通道名 DEP_BRIDGE 定义在模块顶层。 */

    /** 车次是否已在报表「编组车次」列中出现。
     *  唯一比对源：扩展桥同步回的「报表表格」depTableRows（与 #depResultBody 展示同一份）。
     *  三个同步时机都会带来最新表格：① 打开浮窗（onOpen）② 点发送 ③ 点「刷新报表页」。
     *  历史包袱已随简化移除：iframe 直读（#depFrame 已不存在）、depRemoteCheci 三条
     *  采集链（localStorage / BroadcastChannel / checiList 监听）、readCheci 独立指令——
     *  报表表格同步本身就携带了全部车次，无需再单独要一份列表。 */
    function depCheciDup(val) {
      var v = String(val == null ? '' : val).replace(/\s+/g, '').toUpperCase();
      if (!v || !depTableRows) return false;
      for (var t = 0; t < depTableRows.length; t++) {
        var cells = depTableRows[t] || [];
        var cv = String(cells[1] == null ? '' : cells[1]).replace(/\s+/g, '').toUpperCase();
        if (cv && cv === v) return true;                    // 编组车次在第 2 列
      }
      return false;
    }

    // 车次输入框：输入中只做查重提示（红三角标），不写表、不广播；
    // 真正把车次写入 C4 并发送到报表页，仅在「回车 / 失焦」时（pushDepCheciToTab）才发生
    depDupState = false;
    function applyDepCheci(showToast) {
      var inp = $('depInputC4');
      if (!inp) return;
      var val = inp.value.trim();
      depWarn.dup = val ? depCheciDup(val) : false;
      setDepWarn();                                  // 刷新警示标
      if (depWarn.dup && !depDupState && showToast) showDepStatus('当前车次重复', 'warn');
      depDupState = depWarn.dup;
      // 注意：此处【不】写内存——车次只在点「发送」时落库一次（见 btnSendCurrent）
    }
    applyDepCheciRef = applyDepCheci;   // 暴露给顶层：openDepForRow 预填后主动查重（Fix：预填不触发 input）

    // 换长框（E4）手动修改 → 实时重判「列车超长」：depWarnCtx.length 原来只在
    // 「编好」预填时记录，用户在浮窗里把换长改大（如 67 → 78）后警示仍按旧值判断。
    on('depInputE4', 'input', function () {
      depWarnCtx.length = Utils.vbVal(this.value) || 0;
      refreshDepWarnByTrack();
    });

    // 提示气泡定位：气泡已改 position:fixed（视口坐标），CSS 不再决定位置，
    // 悬停 / 键盘聚焦时按警示标当前位置摆放——默认右侧垂直居中；
    // 贴近视口右缘放不下时翻到标左侧（.flip 同时翻转箭头方向）。
    // 之前 absolute + left:100% 的弹法会被 .dep-body（overflow:auto）等中间容器裁掉半截。
    (function () {
      var w = $('depCheciWarn'), tip = $('depWarnTip');
      if (!w || !tip) return;
      function place() {
        var r = w.getBoundingClientRect();
        if (!r.width) return;
        tip.style.top = Math.round(r.top + r.height / 2) + 'px';
        tip.style.transform = 'translateY(-50%)';
        tip.style.left = Math.round(r.right + 10) + 'px';
        var tw = tip.offsetWidth || 0;                       // hover/focus 时 tip 已 display:block，能量宽
        if (r.right + 10 + tw > window.innerWidth - 8) {     // 右侧放不下 → 翻左侧
          tip.style.left = Math.round(r.left - 10 - tw) + 'px';
          tip.classList.add('flip');
        } else {
          tip.classList.remove('flip');
        }
      }
      w.addEventListener('mouseenter', place);
      w.addEventListener('focus', place);
    })();
    on('depInputC4', 'input', function () { applyDepCheci(true); });   // 输入中：仅提示
    // 注：车次框按 Enter 不再自动发送——由用户自己点下方发送按钮触发，避免误发。
    // 输入完成后按 Enter → 失焦（blur），方便收起输入法 / 退出当前框
    on('depInputC4', 'keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); this.blur(); }
    });

    // ============== 时间显示（每秒刷新）==============
    refreshDepTime();
    setInterval(refreshDepTime, 1000);

    // ============== 浮窗内右下角状态提示（替代全局 toast）==============
    // 默认 3 秒自动淡出；鼠标悬停时暂停消失（便于看完整回执），移开后 0.8 秒再消失
    var depStatusTimer = null;
    function hideDepStatus(delay) {
      var el = $('depStatus');
      if (!el) return;
      if (depStatusTimer) clearTimeout(depStatusTimer);
      depStatusTimer = setTimeout(function () { el.classList.remove('show'); }, delay);
    }
    function showDepStatus(msg, type) {
      var el = $('depStatus');
      if (!el) return;
      el.textContent = msg;
      el.className = 'dep-status show ' + (type || 'info');
      hideDepStatus(3000);
    }
    // 悬停暂停 / 移开消失（回执文字较长，常含「目标=…」「错误=…」，需要停留查看）
    (function () {
      var el = $('depStatus');
      if (!el) return;
      el.addEventListener('mouseenter', function () {
        if (depStatusTimer) { clearTimeout(depStatusTimer); depStatusTimer = null; }
      });
      el.addEventListener('mouseleave', function () {
        hideDepStatus(800);          // 移开后短延迟再消失，避免误触抖动
      });
    })();

    // ============== 发送按钮（不同 strategy 决定 iframe 定位方式，挑出能用的）==============
    function getDepCells() {
      return {
        B4: ($('depInputB4') || {}).value || '',
        C4: ($('depInputC4') || {}).value.trim() || '',
        D4: ($('depInputD4') || {}).value || '',
        E4: ($('depInputE4') || {}).value || '',
        F4: ($('depInputF4') || {}).value || ''
      };
    }
    /* ============== 发送目标（原两个模式已合并为单一策略）==============
     * 扩展侧 depFillFunc / depReadFunc 都是「先逐级匹配报表 iframe → 匹配不到再兜底写本帧」，
     * 所以一个策略即可覆盖两种场景：
     *   ① 报表挂在平台页 iframe 里        → 逐级匹配命中
     *   ② 直接打开报表页本身（本帧即报表）→ 走兜底分支
     * 这里固定下发 fs_tab_toolbar；扩展侧对未知策略值也一律按它处理，向后兼容旧值。 */
    var DEP_STRATEGY = 'fs_tab_toolbar';
    var depLastTrack = '';   // 最近一次发送所用的股道（道号），用于回执提示
    /* 一次发送会收到「多个 frame」的回执（每个 frame 的 page-fill 都会各自回传），
     * 采「成功优先 + 窗口聚合」：窗口内保留最佳结果，全成功或窗口结束才展示，
     * 避免后到的失败回执把已经成功的提示覆盖成「编好失败」。 */
    var depFillBest = null;
    var depFillTimer = null;
    function showFillResult(res) {
      if (!res) return;
      var total = 5, ok = res.ok || 0;
      // depLastTrack 取自 cells.B4（形如「4道」），本身已含「道」字，不要再拼一次
      var track = depLastTrack || '';
      var msg;
      if (ok >= total)      msg = track + '编好';
      else if (ok > 0)      msg = track + '编好（部分 ' + ok + '/' + total + '）';
      else                  msg = track + '编好失败';
      if (res.error) msg += ' · 原因：' + res.error;
      else if (res.failed && res.failed.length) msg += ' · 未写入：' + res.failed.join('、');
      showDepStatus(msg, ok > 0 ? 'ok' : 'warn');
      renderDepTable();                        // 顺便刷新「报表表格」结果区
    }
    // ============== 「自动提交」开关：勾选后，发送填表成功即自动点报表页的「提交」==============
    // 由扩展在报表帧内执行：#fr-btn-Submit > div > em > button
    var depAutoSubmitEl = $('depAutoSubmit');
    function depAutoSubmitOn() { return !!(depAutoSubmitEl && depAutoSubmitEl.checked); }
    (function () {
      if (!depAutoSubmitEl) return;
      var wrap = depAutoSubmitEl.closest ? depAutoSubmitEl.closest('.dep-auto-submit') : null;
      function sync() { if (wrap) wrap.classList.toggle('is-on', depAutoSubmitEl.checked); }
      depAutoSubmitEl.addEventListener('change', sync);
      sync();
    })();

    // ============== 报表表格结果：展示读自报表标签页的表格内容 ==============
    // depTableRows 由扩展回执带回（报表页 #frozen-west 的数据行），每次发送/提交后刷新。
    var depTableRows = null;
    var depSubmitted = false;
    function escHtml(s) {
      return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
        return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
      });
    }
    function renderDepTable() {
      var box = $('depResult'), body = $('depResultBody'), title = $('depResultTitle');
      if (!box || !body) return;
      if (!depTableRows || !depTableRows.length) {
        body.innerHTML = '<div class="dep-result-empty">报表页未读到数据行（确认报表已打开，且与当前「发送模式」匹配）</div>';
        if (title) title.textContent = '报表表格';
        box.hidden = false;
        return;
      }
      var head = ['股道', '编组车次', '辆数', '换长', '尾车车号'];
      var html = '<table><thead><tr>';
      for (var h = 0; h < head.length; h++) html += '<th>' + head[h] + '</th>';
      html += '</tr></thead><tbody>';
      for (var i = 0; i < depTableRows.length; i++) {
        html += '<tr>';
        for (var j = 0; j < head.length; j++) {
          var row = depTableRows[i] || [];
          html += '<td>' + escHtml(row[j]) + '</td>';
        }
        html += '</tr>';
      }
      html += '</tbody></table>';
      body.innerHTML = html;
      if (title) {
        title.textContent = '报表表格（共 ' + depTableRows.length + ' 行' +
                            (depSubmitted ? ' · 已自动提交' : '') + '）';
      }
      box.hidden = false;
    }
    on('depResultClose', 'click', function () { var b = $('depResult'); if (b) b.hidden = true; });

    /** 同步报表表格：向扩展桥要一份报表页当前表格。
     *  独立指令（type:'readTable'），与「填表」「自动提交」完全无关。
     *  @param {boolean} [showLoading] 是否先显示「正在读取…」占位（打开浮窗时用） */
    function syncDepReportTable(showLoading) {
      if (showLoading) {
        var box = $('depResult'), body = $('depResultBody'), title = $('depResultTitle');
        if (box && body) {
          body.innerHTML = '<div class="dep-result-empty">正在读取报表表格…</div>';
          if (title) title.textContent = '报表表格';
          box.hidden = false;
        }
      }
      try {
        window.postMessage({
          channel: DEP_BRIDGE,
          type: 'readTable',
          strategy: DEP_STRATEGY,
          ts: Date.now()
        }, '*');
      } catch (e) { /* 未装扩展时静默 */ }
    }

    // ============== 右侧「发送」按钮：用当前选中的模式发送 ==============
    on('btnSendCurrent', 'click', function () {
      var cells = getDepCells();
      if (!cells.C4) { showDepStatus('请先输入车次', 'warn'); return; }
      depLastTrack = cells.B4;
      // 车次写内存：只在点「发送」时落库一次（Store.KEYS.readyTrains，主表「编好车次」列 /
      // 31814 报表同源），输入过程中不实时写。
      if (depCurTrack) {
        Checi.set(depCurTrack, cells.C4);
        syncCheciRender();                           // 防抖刷新主表「编好车次」列
      }
      // 重置回执聚合窗口：本次发送的最佳结果由 showFillResult 统一展示
      depFillBest = null;
      depSubmitted = false;
      if (depFillTimer) clearTimeout(depFillTimer);
      depFillTimer = setTimeout(function () {
        showFillResult(depFillBest);
        // 兜底：即便一个回执都没收到，也刷新一次结果区（会显示「未读到数据行」便于排查）
        renderDepTable();
      }, 900);
      var autoSub = depAutoSubmitOn();
      try {
        /* ① 先独立同步一次报表表格：与「填表」「自动提交」完全解耦，
         *    不管开关开没开，都要把报表页当前表格回执到 #depResult */
        window.postMessage({
          channel: DEP_BRIDGE,
          type: 'readTable',
          strategy: DEP_STRATEGY,
          ts: Date.now()
        }, '*');
        /* ② 再下发填表指令（写入后扩展还会另带一份「提交后」的表格回来）*/
        window.postMessage({
          channel: DEP_BRIDGE,
          type: 'fillByStrategy',
          cells: cells,
          strategy: DEP_STRATEGY,
          autoSubmit: autoSub,                      // 扩展写入后是否代点「提交」
          ts: Date.now()
        }, '*');
        showDepStatus('已下发填报指令 · 车次：' + cells.C4 + (autoSub ? ' · 自动提交' : ''), 'ok');
      } catch (e) {
        showDepStatus('发送失败：' + e.message, 'warn');
      }
    });


    // 「刷新报表页」：向扩展桥下刷新指令，由 background 刷新「popup 报表地址」匹配的标签页
    on('btnReloadReport', 'click', function () {
      try {
        window.postMessage({ channel: DEP_BRIDGE, type: 'reloadReport', ts: Date.now() }, '*');
        showDepStatus('已发送刷新指令…', 'info');
      } catch (e) {
        showDepStatus('刷新指令发送失败：' + e.message, 'warn');
      }
    });

    /* 刷新报表页后的「延迟同步表格」定时器：
     * 报表页 reload 后 contentPane 需要重建，立刻读会拿到刷新前的旧内容甚至读空。 */
    var depReloadSyncTimer = null;

    /* 发送成功后的「延迟读表」定时器：1 秒后（扩展代点的「提交」也已完成）再读一次
     * 报表表格，静默刷新 #depResult（不显示「正在读取…」占位）。 */
    var depSendSyncTimer = null;

    // 监听扩展桥回执（filled）—— 哪个 strategy 写成功 / 失败 / 原因都在这里能直接看到
    window.addEventListener('message', function (ev) {
      var d = ev.data;
      if (!d || d.channel !== DEP_BRIDGE) return;
      // 刷新指令回执
      if (d.type === 'reloadResult') {
        var reloaded = (d.count || 0) > 0;
        showDepStatus(
          reloaded
            ? ('已刷新 ' + d.count + ' 个报表页')
            : '未找到报表页：请检查 popup 里的「报表地址」是否与已打开的页面一致',
          reloaded ? 'ok' : 'warn'
        );
        // 刷新成功 → 等报表页重新加载完，再把表格同步回来（会先显示「正在读取报表表格…」）
        if (reloaded) {
          if (depReloadSyncTimer) clearTimeout(depReloadSyncTimer);
          depReloadSyncTimer = setTimeout(function () { syncDepReportTable(true); }, 1800);
        }
        return;
      }
      // 独立读表回执：不管开不开自动提交，同步到的报表表格都要展示
      if (d.type === 'tableData') {
        if (d.table) {
          depTableRows = d.table;
          renderDepTable();
          if (applyDepCheciRef) applyDepCheciRef(false);   // 表格到手即按当前车次重算查重
        }
        return;
      }
      if (d.type === 'filled' && d.strategy) {
        // 多 frame 回执聚合：保留最佳（ok 最大）结果，全成功立即展示，否则等窗口结束再展示
        var ok = d.ok || 0;
        if (!depFillBest || ok > depFillBest.ok) {
          depFillBest = { ok: ok, error: d.error, failed: d.failed };
        }
        // 扩展回传的「报表页表格内容」：只要有就收下并【立即渲染】。
        // 不依赖 ok>=5，也不依赖「自动提交」开关 —— 不开自动提交同样要能看到报表表格。
        if (d.table) {
          depTableRows = d.table;
          if (d.submitted) { depSubmitted = true; }
          renderDepTable();
          if (applyDepCheciRef) applyDepCheciRef(false);   // 表格到手即按当前车次重算查重
        }
        if (ok >= 5) {
          if (depFillTimer) { clearTimeout(depFillTimer); depFillTimer = null; }
          showFillResult(depFillBest);
          // 勾选「提交」：5 格全部写入成功 → 清空本地录入行（B4~F4），便于连续编组下一股道；
          // 不勾选则保留输入内容。清空后警示标随之熄灭。
          if (depAutoSubmitOn()) {
            ['depInputB4', 'depInputC4', 'depInputD4', 'depInputE4', 'depInputF4'].forEach(function (id) {
              var el = $(id); if (el) el.value = '';
            });
            depWarn.dup = false;
            depDupState = false;
            setDepWarn();
          }
          // 发送成功 1 秒后（扩展代点的「提交」也已完成）再读一次报表表格，
          // 静默刷新 #depResult（拿到提交后的最终状态）。
          if (depSendSyncTimer) clearTimeout(depSendSyncTimer);
          depSendSyncTimer = setTimeout(function () { syncDepReportTable(false); }, 1000);
        }
      }
    });

    // 打开浮窗（点「编好」）就先同步一次报表表格，立刻看到报表当前情况；点发送后会再同步一次
    UI.Modal.register('modalDeparture', {
      onOpen: function () { depSubmitted = false; syncDepReportTable(true); }
    });
    on('depClose', 'click', function () { UI.Modal.close('modalDeparture'); });
  }

  /** 「编好」命中后：打开发车作业浮窗，并按当前股道预填（由 grid-events 的钩子调用）。
   *  必须放在模块顶层，否则 global.DepPanel = { openForRow } 导出时引用不到。 */
  function openDepForRow(trackId, row) {
    var cfg = YardConfig.getTrack(trackId);
    var trackName = cfg ? cfg.name : trackId;
    depCurTrack = trackId;   // 记录当前股道：车次框输入/修改时同步落库到该股道
    var count = row.count || 0;
    var length = (row.length === 0 || row.length == null) ? '' : Number(row.length).toFixed(1);
    var lastCar = '';
    if (row.raw && row.raw.length) {
      var last = row.raw[row.raw.length - 1];
      lastCar = String(last[COL.CARNO] == null ? '' : last[COL.CARNO]);
    }

    // 打开浮窗
    UI.Modal.open('modalDeparture');

    // 记录当前股道换长 / 重量，用于「列车超长 / 列车超重」提示
    depWarnCtx.length = Number(row.length) || 0;
    depWarnCtx.load = Number(row.load) || 0;
    depWarn.dup = false;
    depDupState = false;
    refreshDepWarnByTrack();

    function fillLocalTable() {
      var $b4 = $('depInputB4'), $c4 = $('depInputC4'),
          $d4 = $('depInputD4'), $e4 = $('depInputE4'), $f4 = $('depInputF4');
      if ($b4) $b4.value = Utils.depTrackLabel(trackName);
      if ($c4) $c4.value = Checi.of(trackId);
      if ($d4) $d4.value = String(count);
      if ($e4) $e4.value = length;
      if ($f4) $f4.value = lastCar;
      refreshDepTime();
      if ($c4) setTimeout(function () { $c4.focus(); $c4.select(); }, 80);
    }

    fillLocalTable();
    // 预填值可能本身就与报表重复（重复编组同一车次）：程序赋值不触发 input 事件，
    // 这里主动按当前车次跑一次查重亮标（不弹 toast，与手动输入走同一链路）。
    // 报表表格由 onOpen 的 syncDepReportTable 异步带回，到达后还会再重算一次。
    if (applyDepCheciRef) applyDepCheciRef(false);
  }

  global.DepPanel = { init: init, openForRow: openDepForRow };

})(window);
