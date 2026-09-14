/**
 * sim-panel.js —— 推演面板（明细抽屉的「分身」）
 * ============================================================================
 * 从左侧滑出，与明细左右分栏（宽度由 CSS 变量 --sim-w 控制，中间分隔条可拖动）。
 * 用途：把各股道挑出的车凑成一列"假想车列"，提前看**长度 / 重量 / 顺序**。
 *
 * 数据模型（全部私有于本模块）：
 *   rows          面板中的车辆（与明细同源的行对象引用，可跨股道累加）
 *   keys          车号去重集合
 *   sel           面板内选中的行下标（供「-」删除）
 *   editExcluded  「计重」编辑态中被删除推算载重的行
 *   isOpen / resume / editMode / drag   交互状态
 *
 * 与 app.js 的关系（刻意保持"依赖注入 + 单向调用"）：
 *   渲染与统计复用 app.js 的 renderDetailRows / computeTotals / totalsSpansHtml，
 *   确保与明细同一口径；这些能力由 app.js 在 init() 时通过 deps 注入，
 *   本模块不反向触碰 app.js 的任何私有状态。
 *   反向，app.js 只在 4 个点调用本模块：
 *     SimPanel.init(deps)              启动装配（注册抽屉 + 绑全部事件）
 *     SimPanel.resumeOnOpen()          openDetail 末尾：曾点「主页」则把面板带回
 *     SimPanel.close()                 closeDetail 开头：关明细时一并收起
 *     SimPanel.toggleEditMode()        「双击计重」的双击派发（明细/推演共用一套手势）
 *     SimPanel.excludeDerivedEst(td)   「点推算值删除」的派发（同上）
 *
 * 交互与明细抽屉对齐：面板内可单击 / 拖选（供「-」删除），点标题栏或表格外则取消选中；
 * 标题统计同样跟随选中——有选中行时显示「选中合计」，无选中行时显示全部行合计。
 * 生命周期：内容只由「重置」清空；X 关闭、ESC、关闭明细都只是收起，内容保留。
 * 加载顺序：须在 app.js 之前（app.js 的 init 会调用 SimPanel.init）。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Utils = global.Utils;
  var $ = Utils.$;
  var on = Utils.on;
  var toast = Utils.toast;
  var UI = global.UI;
  var COL = global.Aggregate ? global.Aggregate.COL : null;

  /* app.js 注入的能力，见 init(deps) 的说明 */
  var state, renderDetailRows, computeTotals, totalsSpansHtml, renderCurrentDetail, closeDetail;
  var inited = false;

  /* ==================== 状态 ==================== */

  var rows = [];                 // 面板中的车辆（保持「加入先后」顺序）
  var keys = {};                 // 车号去重：key → true
  var sel = new Set();           // 面板内选中的行下标（供「-」删除）
  var isOpen = false;
  var resume = false;            // 「主页」隐藏过 → 下次打开明细时把面板一起带回来
  var editMode = false;          // 「计重」编辑态（与明细同一套交互）
  var editExcluded = new Set();  // 编辑态中被删除推算载重的行
  var drag = { active: false, moved: false, anchor: -1, snap: null, mode: 'add' };

  /** 车辆去重键：以车号为准；车号为空时用「股道#顺位」兜底，保证仍能加入且可去重 */
  function keyOf(row) {
    var no = String(row[COL.CARNO] == null ? '' : row[COL.CARNO]).trim();
    if (no) return no;
    return '@' + (row.__track || row[COL.TRACK] || '') + '#' + (row[COL.SEQ] == null ? '' : row[COL.SEQ]);
  }

  /* ==================== 渲染 ==================== */

  /** 重渲面板表格（复用明细同一套列定义与渲染函数）。
   *  rows 本身就是「加入先后」顺序（每次「+」按选中顺序追加，新的一批排在已有内容之后），
   *  所以不做任何排序：序号列取下标 +1 即为「本板序号」，删除后自动重排。 */
  function render() {
    renderDetailRows(rows, {
      head: $('simHead'), body: $('simBody'), table: $('simTable')
    }, {
      // 序号 = 本板顺序，而非车辆在原股道明细中的顺位
      seqText: function (row, i) { return String(i + 1); },
      rowAttr: function (row, i) {
        return ' data-i="' + i + '"' + (sel.has(i) ? ' class="selected"' : '');
      },
      // 「计重」编辑：进入编辑才显示载重列推算值（供删除），与明细同一套渲染分支
      excluded: editExcluded,
      editMode: editMode
    });
    updateTitle();
  }

  /** 刷新标题：股道名恒为「X道」（推演区标志，非真实股道）。
   *  统计口径与明细完全一致——**有选中行时按选中行求和，无选中行时恢复为全部行合计**。
   *  总重复用明细的「计重」机制：双击总重进入编辑，点推算值单元格删除该行推算载重。 */
  function updateTitle() {
    var list = sel.size
      ? rows.filter(function (_, i) { return sel.has(i); })
      : rows;
    var t = computeTotals(list, editMode ? editExcluded : null);
    $('simTitle').innerHTML =
      '<span class="dt-name">X道 - </span>' +
      totalsSpansHtml(t, { calc: editMode, editable: true, editing: editMode });
  }

  /** 切换「计重」编辑态：与明细 setDetailEditMode 同构 */
  function setEditMode(enable) {
    // 面板为空时没有可编辑的推算值，直接提示，不进入编辑态
    if (enable && !rows.length) { toast('推演面板为空', 'error'); return; }
    editMode = enable;
    var tbl = $('simTable');
    if (tbl) tbl.classList.toggle('detail-edit', enable);
    var span = $('simTitle') && $('simTitle').querySelector('.dt-weight-edit');
    if (span) span.classList.toggle('editing', enable);
    if (!enable) editExcluded.clear();
    // 进入 / 退出都要重渲：进入才显示载重列推算值，退出则复原为空
    render();
  }

  /** 供 app.js 的「双击 .dt-weight-edit」派发调用（明细与推演共用一套双击手势） */
  function toggleEditMode() { setEditMode(!editMode); }

  /** 供 app.js 的「点 td.derived-est」派发调用：该行推算载重不计入计重。
   *  （明细 / 搜索表内的同款逻辑在 app.js，两处按 td.closest 各自分流，互不串扰） */
  function excludeDerivedEst(td) {
    if (!editMode) return;
    var tr = td.closest('tr');
    if (!tr) return;
    var row = rows[+tr.getAttribute('data-i')];
    if (!row || editExcluded.has(row)) return;
    editExcluded.add(row);
    td.textContent = '';
    td.classList.remove('derived-est');   // 移除该类 → 编辑态下不再显示「删除」按钮
    td.classList.add('excluded-cell');
    updateTitle();                        // 重算计重（排除该行推算载重）
  }

  /* ==================== 开关 ==================== */

  /** 打开：明细让出左侧空间，触发按钮变为「+」 */
  function open() {
    isOpen = true;
    document.body.classList.add('sim-open');
    UI.Drawer.open('simDrawer');
    var b = $('btnSim');
    if (b) {
      b.textContent = '+';
      b.title = '把明细中选中的车辆加入推演面板';
      b.classList.add('btn-sim-add');
    }
    render();
  }

  /** 收起。真正的状态复位放在抽屉的 onClose 里，
   *  这样 X 按钮 / ESC / 关闭明细 三条路径都会走同一套复位逻辑。 */
  function close() { UI.Drawer.close('simDrawer'); }

  /** openDetail 末尾调用：之前点过「主页」则把面板一并带回（内容保留） */
  function resumeOnOpen() {
    if (!resume) return;
    resume = false;
    open();
  }

  /** 「主页」：同时隐藏面板与明细，回主表挑另一股道；下次打开明细时自动带回。 */
  function goHome() {
    resume = isOpen;
    closeDetail();   // 明细的关闭函数内部会调 SimPanel.close()
  }

  /* ==================== 内容操作 ==================== */

  /** 把明细抽屉当前选中的行加入面板（按车号去重，整组插入）。没选任何行 → 提示不动作。
   *  插入方向由标题栏下拉决定（与"升序/降序"无关，是插在已有内容的上方还是下方）：
   *    S（默认）= 从下方插入 → 追加到末尾
   *    W        = 从上方插入 → 插到最前面
   *  两种都是「整组」插入：本批内部保持选中先后，不拆散、不重排已有内容。 */
  function addSelected() {
    var r = state.rows[state.detailIdx];
    if (!r || !state.detailSel || !state.detailSel.size) {
      toast('未选择车辆', 'error');
      return;
    }
    var list = r.raw || [];
    var added = 0, skip = 0;
    var batch = [];                           // 本批要加入的车（保持选中先后）
    state.detailSel.forEach(function (i) {
      var row = list[i];
      if (!row) return;
      var k = keyOf(row);
      if (keys[k]) { skip++; return; }        // 已存在：按车号去重，不重复加入
      keys[k] = true;
      batch.push(row);
      added++;
    });
    var dir = ($('simDir') && $('simDir').value) || 'S';
    rows = (dir === 'W') ? batch.concat(rows) : rows.concat(batch);
    state.detailSel.clear();                  // 加入后取消明细选中，方便接着挑下一批
    renderCurrentDetail();                    // 明细表重渲 → 选中高亮同步消失
    render();
    if (added || skip) {
      toast('推演面板：已从' + (dir === 'W' ? '上方' : '下方') + '加入 ' + added + ' 辆' +
        (skip ? '，跳过重复 ' + skip + ' 辆' : ''), added ? 'ok' : 'error');
    }
  }

  /** 删除面板中选中的车辆（「-」按钮） */
  function removeSelected() {
    if (!sel.size) return;
    var keep = [];
    for (var i = 0; i < rows.length; i++) {
      if (sel.has(i)) {
        var row = rows[i];
        delete keys[keyOf(row)];
        editExcluded.delete(row);   // 同步清理计重编辑态里的残留
      } else keep.push(rows[i]);
    }
    rows = keep;
    sel.clear();
    render();
  }

  /** 重置：清空全部内容（同时退出计重编辑态） */
  function reset() {
    if (!rows.length) return;
    rows = [];
    keys = {};
    sel.clear();
    editExcluded.clear();
    if (editMode) setEditMode(false);   // 内部会重渲
    else render();
    toast('推演面板已重置', 'ok');
  }

  /* ==================== 面板内行选中（供「-」删除） ==================== */

  /** 清空面板内选中（点标题栏 / 表格外时调用，与明细面板同一交互） */
  function clearSel() {
    if (!sel.size) return;
    sel.clear();
    var trs = $('simBody') ? $('simBody').querySelectorAll('tr.selected') : [];
    for (var i = 0; i < trs.length; i++) trs[i].classList.remove('selected');
    updateTitle();   // 标题回到「全部行合计」
  }

  /** 单行选中状态切换（与明细同样用 .selected 高亮） */
  function toggleRow(i, add) {
    if (add) sel.add(i); else sel.delete(i);
    var tr = $('simBody') && $('simBody').querySelector('tr[data-i="' + i + '"]');
    if (tr) tr.classList.toggle('selected', add);
    updateTitle();   // 标题随选中项求和
  }

  /** 拖选：以按下行为锚点，按快照 + 区间重算，避免"拖出去的行"残留选中 */
  function applyDrag(i) {
    sel = new Set(drag.snap);
    var a = Math.min(drag.anchor, i), b = Math.max(drag.anchor, i);
    for (var j = a; j <= b; j++) {
      if (drag.mode === 'add') sel.add(j); else sel.delete(j);
    }
    var trs = $('simBody') ? $('simBody').querySelectorAll('tr') : [];
    for (var k = 0; k < trs.length; k++) {
      trs[k].classList.toggle('selected', sel.has(+trs[k].getAttribute('data-i')));
    }
    updateTitle();   // 拖动中实时刷新「选中合计」
  }

  function bindTableSelection() {
    on('simBody', 'mousedown', function (e) {
      if (!isOpen) return;
      var tr = e.target.closest('tr');
      if (!tr || tr.querySelector('td.stay') === e.target) return;
      e.preventDefault();
      drag.active = true;
      drag.moved = false;
      drag.anchor = +tr.getAttribute('data-i');
      drag.snap = new Set(sel);
      // 起点已选中 → 取消模式；否则 → 加入模式
      drag.mode = sel.has(drag.anchor) ? 'del' : 'add';
    });
    on('simBody', 'mouseover', function (e) {
      if (!drag.active) return;
      var tr = e.target.closest('tr');
      if (!tr) return;
      drag.moved = true;
      applyDrag(+tr.getAttribute('data-i'));
    });
    document.addEventListener('mouseup', function () {
      if (!drag.active) return;
      drag.active = false;
      drag.snap = null;
    });
    on('simBody', 'click', function (e) {
      if (!isOpen) return;
      if (drag.moved) { drag.moved = false; return; }   // 拖动已实时应用，click 不重复处理
      var tr = e.target.closest('tr');
      if (!tr || tr.querySelector('td.stay') === e.target) return;
      var i = +tr.getAttribute('data-i');
      toggleRow(i, !sel.has(i));
    });
  }

  /** 与明细同一交互：点标题栏、或表格外的任何地方 → 取消面板内选中。
   *  两处放行是为了不误伤自己：
   *    · 表格内——那是框选区域，由 bindTableSelection 负责；
   *    · 方向下拉——边上的控件，点它不应把刚选好的车丢掉。
   *  面板里的按钮无需在此排除：它们在 init 里已对 mousedown 做了 stopPropagation，
   *  document 根本收不到（否则点「-」会先清空选中，删除就失效了）。 */
  function bindClearSelection() {
    document.addEventListener('mousedown', function (e) {
      if (drag.active) return;                       // 正在拖动框选，不清
      var t = e.target;
      if (!t || !t.closest) return;
      if (t.closest('#simTable, #simDir')) return;
      clearSel();
    });
  }

  /* ==================== 分隔条拖动 ==================== */

  /** 拖 .sim-splitter 调整面板宽度（改 CSS 变量 --sim-w，明细自动占剩余） */
  function bindSplitter() {
    var sp = $('simSplitter');
    if (!sp) return;
    sp.addEventListener('mousedown', function (e) {
      if (!isOpen) return;
      e.preventDefault();
      e.stopPropagation();
      sp.classList.add('dragging');
      // 下限：推演至少 260px（或 20%）；上限：明细至少留 360px（或 30%）
      var minW = Math.max(260, window.innerWidth * 0.2);
      var maxW = window.innerWidth - Math.max(360, window.innerWidth * 0.3);
      function move(ev) {
        var w = Math.min(Math.max(ev.clientX, minW), maxW);
        document.body.style.setProperty('--sim-w', w + 'px');
      }
      function up() {
        sp.classList.remove('dragging');
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
      }
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
  }

  /* ==================== 装配 ==================== */

  /**
   * @param {Object} deps 由 app.js 注入（均为 app.js 内的函数/对象引用）
   * @param {Object}   deps.state              主状态（读 rows/detailIdx/detailSel）
   * @param {Function} deps.renderDetailRows   明细同款表格渲染（list, els, opts）
   * @param {Function} deps.computeTotals      辆数/换长/自重/载重/总重/计重
   * @param {Function} deps.totalsSpansHtml    标题统计 span 拼装
   * @param {Function} deps.renderCurrentDetail 重渲当前股道明细（加车后清选中高亮）
   * @param {Function} deps.closeDetail        关闭明细抽屉（「主页」用）
   */
  function init(deps) {
    if (inited) return;
    inited = true;
    state = deps.state;
    renderDetailRows = deps.renderDetailRows;
    computeTotals = deps.computeTotals;
    totalsSpansHtml = deps.totalsSpansHtml;
    renderCurrentDetail = deps.renderCurrentDetail;
    closeDetail = deps.closeDetail;

    // 抽屉注册：不配遮罩——它与明细拼起来铺满视口，点哪一侧都是有效区域。
    // onClose 统一做状态复位，保证 X 按钮 / ESC / 关闭明细 三条路径行为一致。
    UI.Drawer.register('simDrawer', {
      onClose: function () {
        if (!isOpen) return;
        isOpen = false;
        document.body.classList.remove('sim-open');
        var b = $('btnSim');
        if (b) {
          b.textContent = '推演面板';
          b.title = '打开推演面板';
          b.classList.remove('btn-sim-add');
        }
      }
    });

    // 按钮不许让「点表格外清空明细选中」生效：stopPropagation 同时挡住 document 上的
    // 清空监听与 .drawer-head 上的清空监听（btnSim 就挂在明细的标题栏里）。
    ['btnSim', 'btnSimHome', 'btnSimDel', 'btnSimReset', 'btnSimClose'].forEach(function (id) {
      var el = $(id);
      if (!el) return;
      el.addEventListener('mousedown', function (e) { e.stopPropagation(); });
      el.addEventListener('click', function (e) { e.stopPropagation(); });
    });
    on('btnSim', 'click', function () { if (isOpen) addSelected(); else open(); });
    on('btnSimHome', 'click', goHome);
    on('btnSimDel', 'click', removeSelected);
    on('btnSimReset', 'click', reset);
    on('btnSimClose', 'click', close);

    bindTableSelection();
    bindClearSelection();
    bindSplitter();
  }

  global.SimPanel = {
    init: init,
    open: open,
    close: close,
    resumeOnOpen: resumeOnOpen,
    toggleEditMode: toggleEditMode,
    excludeDerivedEst: excludeDerivedEst
  };

})(window);
