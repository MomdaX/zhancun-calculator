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
 * 加入采用「全有或全无」：选中里含重复则整批不加入，并在面板中把重复行闪 3 次（flashByKeys）；
 * 成功加入的那一批会记进 justAdded，渲染时挂 .sim-new（行首蓝条 + 淡蓝底）标出「本次加入」；
 * 该标记不跨交互保留——点表格外、或收起面板（× / ESC / 关闭明细 /「主页」）即清除。
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
  var state, renderDetailRows, computeTotals, totalsSpansHtml, closeDetail;
  var inited = false;

  /* ==================== 状态 ==================== */

  var rows = [];                 // 面板中的车辆（保持「加入先后」顺序）
  var keys = {};                 // 车号去重：key → true
  var sel = new Set();           // 面板内选中的行下标（供「-」删除）
  var justAdded = new Set();     // 最近一次「+」加入的车（行对象），用于行首蓝条标记
  var isOpen = false;
  var resume = false;            // 「主页」隐藏过 → 下次打开明细时把面板一起带回来
  var editMode = false;          // 「计重」编辑态（与明细同一套交互）
  var editExcluded = new Set();  // 编辑态中被删除推算载重的行
  var drag = { active: false, moved: false, anchor: -1, snap: null, mode: 'add' };
  var flashTimer = null;         // 重复车辆闪烁的收尾计时器（保证连点也能重新闪）
  var FLASH_CLASS = 'sim-dup-flash';
  var FLASH_MS = 1200;           // 0.4s × 3 次，与 CSS 动画保持一致
  var NEW_CLASS = 'sim-new';     // 「本次加入」的行标记类

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
        var cls = [];
        if (sel.has(i)) cls.push('selected');
        // 「本次加入」的标记：按行对象判断而非下标——插入方向可上可下、删行还会重排，
        // 只有对象引用是稳定的（同一辆车在面板里始终是同一个引用）
        if (justAdded.has(row)) cls.push(NEW_CLASS);
        return ' data-i="' + i + '"' + (cls.length ? ' class="' + cls.join(' ') + '"' : '');
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

  /** 把明细抽屉当前选中的行加入面板。没选任何行 → 提示不动作。
   *  【全有或全无】选中里只要有一辆重复——面板中已有，或选中内部选到同一车号——
   *  则**整批都不加入**，改为在面板中把重复的那些行闪 3 次提示；明细选中保持不动，
   *  便于用户剔除重复项后重试。
   *  插入方向由标题栏下拉决定（与"升序/降序"无关，是插在已有内容的上方还是下方）：
   *    S（默认）= 从下方插入 → 追加到末尾
   *    W        = 从上方插入 → 插到最前面
   *  插入时整组保持选中先后，不拆散、不重排已有内容。 */
  function addSelected() {
    var r = state.rows[state.detailIdx];
    if (!r || !state.detailSel || !state.detailSel.size) {
      toast('未选择车辆', 'error');
      return;
    }
    var list = r.raw || [];

    // 第一遍：只校验，不落库。因为「含重复则整批不加」，若边校验边写 keys，
    // 整批放弃时还得回滚——两遍走法从根上杜绝脏数据。
    var batch = [];        // 本批要加入的车（保持选中先后）
    var dupKeys = [];      // 重复车的去重键（面板中已有 / 选中内部选重），供闪烁反查
    var seen = {};         // 本批内部去重：明细里同一车号出现两次时同样按重复处理
    state.detailSel.forEach(function (i) {
      var row = list[i];
      if (!row) return;
      var k = keyOf(row);
      if (keys[k] || seen[k]) { dupKeys.push(k); return; }
      seen[k] = true;
      batch.push(row);
    });

    // 有重复 → 整批不加。此时面板内容未变，直接对着现有 DOM 闪即可，无需 render。
    if (dupKeys.length) {
      flashByKeys(dupKeys);
      toast('推演面板：选中里有 ' + dupKeys.length + ' 辆重复，整批 ' +
        (batch.length + dupKeys.length) + ' 辆均未加入', 'error');
      return;
    }

    // 第二遍：确认无重复，才落库并插入
    batch.forEach(function (row) { keys[keyOf(row)] = true; });
    var dir = ($('simDir') && $('simDir').value) || 'S';
    rows = (dir === 'W') ? batch.concat(rows) : rows.concat(batch);
    justAdded = new Set(batch);   // 只有最近一批算「本次加入」：旧标记整体让位（下次 render 即生效）
    // 刻意不清明细的选中：加完仍保留高亮，便于对照/继续追加同一批；
    // 要取消选中按明细既有的交互走——点表格外（或标题栏）即可。
    render();
    toast('推演面板：已从' + (dir === 'W' ? '上方' : '下方') + '加入 ' + batch.length + ' 辆', 'ok');
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
        justAdded.delete(row);      // 以及「本次加入」标记里的残留
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
    justAdded.clear();
    if (editMode) setEditMode(false);   // 内部会重渲
    else render();
    toast('推演面板已重置', 'ok');
  }

  /**
   * 重复车辆提示：让面板中对应行闪烁 3 次。
   * 入参用「去重键」而非行对象——车号相同的车，明细里可能是另一个行对象，
   * 面板里存的是先加入的那个引用，按 key 反查才不会漏闪。
   * 必须在 render() 之后调用：render 会重建 tbody，先加类会被冲掉。
   */
  function flashByKeys(dupKeys) {
    var body = $('simBody');
    if (!body) return;
    var trs = [];
    dupKeys.forEach(function (k) {
      for (var i = 0; i < rows.length; i++) {
        if (keyOf(rows[i]) !== k) continue;
        var tr = body.querySelector('tr[data-i="' + i + '"]');
        if (tr && trs.indexOf(tr) < 0) trs.push(tr);
        break;
      }
    });
    if (!trs.length) return;

    // 先清掉上一批（含未跑完的计时器），再重新触发：连点「+」时才能重新闪而不是被忽略
    if (flashTimer) { clearTimeout(flashTimer); flashTimer = null; }
    var old = body.querySelectorAll('tr.' + FLASH_CLASS);
    for (var j = 0; j < old.length; j++) old[j].classList.remove(FLASH_CLASS);
    void body.offsetWidth;   // 强制回流：让同一个动画类能被重新触发
    trs.forEach(function (tr) { tr.classList.add(FLASH_CLASS); });

    flashTimer = setTimeout(function () {
      trs.forEach(function (tr) { tr.classList.remove(FLASH_CLASS); });
      flashTimer = null;
    }, FLASH_MS);
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

  /** 清「本次加入」标记。标记只对"刚加进来"这一小段时间有意义，
   *  用户一旦把注意力移开（点表格外）或收起面板，它就失去价值，故一并清掉。
   *  注意：光清 Set 不会摘掉行上的类，得手动摘（与 clearSel 同款做法，避免为此整表重渲）。 */
  function clearJustAdded() {
    if (!justAdded.size) return;
    justAdded.clear();
    var body = $('simBody');
    if (!body) return;
    var trs = body.querySelectorAll('tr.' + NEW_CLASS);
    for (var i = 0; i < trs.length; i++) trs[i].classList.remove(NEW_CLASS);
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
   *  面板里的按钮与分隔条也无需在此排除：它们在各自的 mousedown 里已 stopPropagation，
   *  document 根本收不到（否则点「-」会先清空选中、删除就失效；拖分隔条也会把选中一起清掉）。 */
  function bindClearSelection() {
    document.addEventListener('mousedown', function (e) {
      if (drag.active) return;                       // 正在拖动框选，不清
      var t = e.target;
      if (!t || !t.closest) return;
      if (t.closest('#simTable, #simDir')) return;
      clearSel();
      clearJustAdded();
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
   * @param {Function} deps.closeDetail        关闭明细抽屉（「主页」用）
   */
  function init(deps) {
    if (inited) return;
    inited = true;
    state = deps.state;
    renderDetailRows = deps.renderDetailRows;
    computeTotals = deps.computeTotals;
    totalsSpansHtml = deps.totalsSpansHtml;
    closeDetail = deps.closeDetail;

    // 抽屉注册：不配遮罩——它与明细拼起来铺满视口，点哪一侧都是有效区域。
    // onClose 统一做状态复位，保证 X 按钮 / ESC / 关闭明细 三条路径行为一致。
    UI.Drawer.register('simDrawer', {
      onClose: function () {
        if (!isOpen) return;
        isOpen = false;
        document.body.classList.remove('sim-open');
        // 「本次加入」标记不跨收起保留：× 按钮 / ESC / 关闭明细 /「主页」都走这里
        clearJustAdded();
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
