/**
 * grid-events.js —— 主表（#grid）自身的事件与交互
 * ============================================================================
 * 从 app.js 的 bind() 里整块剥离，收拢「主表本身」的东西：
 *   · 行选中（单击高亮 + 记录 state.selectedIdx）
 *   · 「编好」标记体系（clearBianhao / inBianhaoRange / findRowByTrack / syncBianhao）
 *   · 选中行所属分组的合并单元格加底色（syncGroupHighlight）
 *   · 「编好」伪元素按钮命中 → 通过 openDepForRow 钩子叫开发车浮窗
 *   · 「编好车次」列双击就地录入车次（editCheciCell）
 *   · 空线分组显示 / 隐藏开关
 *   · 注意事项列「收起 / 展开」
 *   · 自适应列宽按钮、上一股道 / 下一股道
 *
 * 与 app.js 的关系（依赖注入 + 单向调用）：
 *   app.js 在 init() 时注入 state / render / syncVirtualBtn / stepDetail / openDepForRow；
 *   本模块不反向触碰 app.js 的私有状态。
 *   openDepForRow 是**唯一的跨模块钩子**：命中判定与行数据读取属于主表，
 *   而「打开浮窗 + 预填 5 格 + 车次查重」属于发车模块（日后归 dep-panel），故由注入方实现。
 *
 * 三个监听必须绑在 #grid 的**捕获阶段**：
 *   编好按钮与注意事项按钮命中都靠坐标/类判定，必须 stopPropagation 抢在
 *   tbody 冒泡的行选中之前，否则刚点出来的「编好」标记会被行选中清掉。
 *
 * 加载顺序：须在 columns.js / track.config.js / col-resize.js / checi-store.js 之后、
 *           app.js 之前（app.js 的 init 会调用 GridEvents.init）。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Utils = global.Utils;
  var $ = Utils.$;
  var on = Utils.on;
  var toast = Utils.toast;
  var Store = global.Store;
  var YardConfig = global.YardConfig;
  var ColResize = global.ColResize;
  var Checi = global.Checi;

  /* app.js 注入的能力，见 init(deps) 的说明 */
  var state, render, syncVirtualBtn, stepDetail, openDepForRow;
  var inited = false;

  /* ==================== 「编好」标记体系 ==================== */

  /** 清除所有「编好」标记 */
  function clearBianhao() {
    var list = $('tbody').querySelectorAll('td.bianhao-on');
    for (var i = 0; i < list.length; i++) list[i].classList.remove('bianhao-on');
  }

  /** 该股道是否落在「编好」生效区间（1道 ~ X15，排除 B1/B2，按 track.config.js 的股道顺序） */
  function inBianhaoRange(trackId) {
    if (BIANHAO_EXCLUDE.indexOf(trackId) >= 0) return false;
    var t = YardConfig.getTrack(trackId);
    var a = YardConfig.getTrack(BIANHAO_FROM);
    var b = YardConfig.getTrack(BIANHAO_TO);
    if (!t || !a || !b) return false;
    var lo = Math.min(a.index, b.index), hi = Math.max(a.index, b.index);
    return t.index >= lo && t.index <= hi;
  }

  /** 根据 trackId 在 state.rows 中查找对应行数据 */
  function findRowByTrack(trackId) {
    if (!state.rows) return null;
    for (var i = 0; i < state.rows.length; i++) {
      if (state.rows[i].track === trackId) return state.rows[i];
    }
    return null;
  }

  /** 选中行后刷新「编好」标记：给该行的目标列单元格打标记 → ::after 浮现。
   *  行级触发——点这一行任意单元格都会显示，无需点到第 9 列本身。 */
  function syncBianhao(tr) {
    clearBianhao();
    if (!tr || tr.classList.contains('area-banner')) return;
    if (tr.classList.contains('blank')) return;   // 空股道不显示「编好」
    if (!inBianhaoRange(tr.getAttribute('data-track'))) return;
    var td = tr.querySelector('td[data-col="' + BIANHAO_COL + '"]');
    if (td) td.classList.add('bianhao-on');
  }

  /** 选中行后：给它所属分组的合并单元格加底色，让用户一眼看出"当前选中的是哪一组"。
   *  分组列用 rowspan 合并，组内只有首行有该单元格，所以按 data-group 反查
   *  （渲染时组名同时打在 tr 与分组 td 上，见 app.js render）。 */
  function syncGroupHighlight(tr) {
    var tbody = $('tbody');
    if (!tbody) return;
    var prev = tbody.querySelectorAll('td.col-b-group.grp-on');
    for (var i = 0; i < prev.length; i++) prev[i].classList.remove('grp-on');
    if (!tr) return;
    var g = tr.getAttribute('data-group');
    if (!g) return;
    var all = tbody.querySelectorAll('td.col-b-group');
    for (var j = 0; j < all.length; j++) {
      if (all[j].getAttribute('data-group') === g) { all[j].classList.add('grp-on'); break; }
    }
  }

  /* ==================== 事件绑定 ==================== */

  /** 单击选中（作业区横幅行不参与选中，否则会被高亮且 selectedIdx 变为 NaN） */
  function bindRowSelect() {
    on('tbody', 'click', function (e) {
      var tr = e.target.closest('tr');
      if (!tr || tr.classList.contains('area-banner')) return;
      var old = $('tbody').querySelector('tr.selected');
      if (old) old.classList.remove('selected');
      tr.classList.add('selected');
      state.selectedIdx = +tr.getAttribute('data-idx');
      // 选中这一行即刷新「编好」标记（无需点击第 9 列）
      syncBianhao(tr);
      syncGroupHighlight(tr);
    });
  }

  /** 空线分组显示 / 隐藏开关（状态本地持久化，按钮文案由 syncVirtualBtn 同步） */
  function bindEmptyGroupToggle() {
    on('btnToggleVirtual', 'click', function () {
      state.showEmptyGroups = !state.showEmptyGroups;
      Store.set(Store.KEYS.showEmptyGroups, state.showEmptyGroups);   // 持久记忆，刷新后保持
      syncVirtualBtn();
      render();
      toast((state.showEmptyGroups ? '已显示' : '已隐藏') + '空线分组', 'ok');
    });
  }

  /** 「编好车次」列：双击就地录入车次（回车 / 失焦保存，Esc 取消）。
   *  已有车次时先清空落库，便于直接改填新车次。 */
  function editCheciCell(td, trackId) {
    if (!td || !document.contains(td) || td.querySelector('.checi-input')) return;
    var old = Checi.of(trackId);        // 调用处已清空过 → 重录时这里是空串
    td.textContent = '';
    var inp = document.createElement('input');
    inp.type = 'text';
    inp.className = 'checi-input';
    inp.maxLength = 10;
    inp.placeholder = '车次';
    inp.value = old;
    td.appendChild(inp);
    inp.focus();
    inp.select();
    var done = false;
    function finish(save) {
      if (done) return;
      done = true;
      if (save) Checi.set(trackId, inp.value.trim());
      render();                        // 复原单元格，并让 31814 / 发车流程拿到最新值
    }
    inp.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    inp.addEventListener('blur', function () { finish(true); });
    // 输入期间的鼠标事件不冒泡到「行选中 / 编好命中 / 双击开明细」
    ['mousedown', 'click', 'dblclick'].forEach(function (ev) {
      inp.addEventListener(ev, function (e) { e.stopPropagation(); });
    });
  }

  /** #grid 上的三个捕获阶段监听（顺序即注册顺序，与拆分前一致）：
   *   ① 注意事项列右侧按钮 → 整表收起 / 展开
   *   ② 「编好」按钮命中 → 钩子转发给发车模块
   *   ③ 「编好车次」列双击 → 就地录入 */
  function bindGridCapture() {
    var gridEl = $('grid');
    if (!gridEl) return;

    // ① 注意事项列：点击行内右侧按钮 → 整表收起/展开（捕获阶段，先于行选中）
    gridEl.addEventListener('click', function (e) {
      var td = e.target.closest ? e.target.closest('td.col-a') : null;
      if (!td) return;
      // 仅当点击在单元格右侧按钮区域（右 32px）时触发
      var rect = td.getBoundingClientRect();
      if (e.clientX - rect.left < rect.width - 32) return;
      e.stopPropagation();
      e.preventDefault();
      var collapsed = gridEl.classList.toggle('notes-collapsed');
      Store.set(Store.KEYS.notesCollapsed, collapsed);
      toast(collapsed ? '已收起注意事项' : '已展开注意事项', 'ok');
    }, true);  // true = 捕获阶段，先于 tbody 行选中触发

    // ② 「编好」伪元素按钮：点击已标记单元格左侧按钮区 → 转发给发车模块。
    // 伪元素不是事件目标，只能用坐标判定左侧命中区；捕获阶段 + stopPropagation，
    // 避免冒泡到 tbody 的行选中把标记清掉。
    gridEl.addEventListener('click', function (e) {
      var td = e.target.closest
        ? e.target.closest('td[data-col="' + BIANHAO_COL + '"].bianhao-on')
        : null;
      if (!td) return;
      var rect = td.getBoundingClientRect();
      // 仅当点击落在单元格左侧按钮区（BIANHAO_HIT_W）时才触发
      if (e.clientX - rect.left > BIANHAO_HIT_W) return;
      e.stopPropagation();
      e.preventDefault();

      // 读取当前行股道数据，交给发车模块去开浮窗并预填
      var trackId = td.parentNode.getAttribute('data-track');
      var row = findRowByTrack(trackId);
      if (!row) { toast('未找到股道数据', 'warn'); return; }
      openDepForRow(trackId, row);
    }, true);

    // ③ 「编好车次」列：双击录入。
    // 捕获阶段 + stopPropagation：抢在 tbody 的「双击行 → 打开明细抽屉」之前，两者不打架。
    gridEl.addEventListener('dblclick', function (e) {
      var td = e.target.closest ? e.target.closest('td[data-col="checi"]') : null;
      if (!td || !document.contains(td)) return;
      e.stopPropagation();
      e.preventDefault();
      var tr = td.closest('tr');
      var trackId = tr ? tr.getAttribute('data-track') : '';
      if (!trackId) return;
      if (Checi.of(trackId)) Checi.set(trackId, '');   // 已有车次 → 先清空落库
      editCheciCell(td, trackId);
    }, true);
  }

  /** 自适应列宽：重置为内容自适应，清除手动拖动记忆 */
  function bindAutoFit() {
    on('btnAutoFitCols', 'click', function () {
      var g = $('grid'), dt = $('detailTable');
      ColResize.safeGet(g).reset();
      ColResize.safeGet(dt).reset();
      var thA = g && g.querySelector('thead th.col-a');
      if (g && thA) g.style.setProperty('--col-a-w', thA.offsetWidth + 'px');
      toast('列宽已重置为自适应', 'ok');
    });
  }

  /** 明细抽屉的「上一股道 / 下一股道」（切股道逻辑在 app.js，这里只转发） */
  function bindTrackStep() {
    on('btnPrevTrack', 'click', function () { stepDetail(-1); });
    on('btnNextTrack', 'click', function () { stepDetail(1); });
  }

  /* ==================== 装配 ==================== */

  /**
   * @param {Object} deps 由 app.js 注入
   * @param {Object}   deps.state           主状态（读 rows / 读写 showEmptyGroups、selectedIdx）
   * @param {Function} deps.render          重渲主表
   * @param {Function} deps.syncVirtualBtn  同步「隐藏/显示空线分组」按钮文案
   * @param {Function} deps.stepDetail      切换上/下一股道
   * @param {Function} deps.openDepForRow   「编好」命中后打开发车浮窗并预填：(trackId, row)
   */
  function init(deps) {
    if (inited) return;
    inited = true;
    state = deps.state;
    render = deps.render;
    syncVirtualBtn = deps.syncVirtualBtn;
    stepDetail = deps.stepDetail;
    openDepForRow = deps.openDepForRow;

    bindRowSelect();
    bindEmptyGroupToggle();
    bindGridCapture();
    bindAutoFit();
    bindTrackStep();
  }

  global.GridEvents = { init: init };

})(window);
