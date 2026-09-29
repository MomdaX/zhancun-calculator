/**
 * flow-table.js —— 车流统计表（左侧无遮罩抽屉）
 * ============================================================================
 * 暴露全局：window.FlowTable
 * 加载顺序：须在 utils / store / ui / report31814 之后（复用 Report31814 的截图/打印）
 *
 * 功能：
 *   · 「功能 ▾」菜单「车流统计表」打开左侧抽屉（无遮罩，主表可继续点）。
 *   · 抽屉内「钦州港车流计划表」：表头字段（预计日期/时段/站调(签)）可编辑；
 *     数据列 序号 / 预计编好时间 / 车流来源(港区|站编) / 编组内容，全部可录入。
 *   · 「预计编好时间」输入自动格式化：1230 → 12:30，230 → 02:30（末两位为分、其余为时）。
 *   · 行操作：新增 / 插入（当前行上方）/ 删除（当前行）/ 拖动排序。
 *   · 数据 localStorage 持久化（Store.KEYS.flowTable），抽屉宽度也记忆。
 *   · 头部按钮：字体放大/缩小、截图、打印、关闭 —— 与 31814 报表头部同一套交互。
 *   · 抽屉打开期间主表隐藏 5 列（有效长/换长/发送/到达车次/载重），由 CSS 类驱动。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Utils = global.Utils;
  var Store = global.Store;
  if (!Utils || !Store) return;

  var $ = Utils.$;
  var on = Utils.on;

  var STORE_KEY = Store.KEYS.flowTable;
  var WIDTH_KEY = Store.KEYS.flowTable + '.width';   // 抽屉宽度记忆（复用前缀键）

  // 车流来源多选选项（悬停「+」展开，勾选多项以「+」拼接）
  var SOURCE_OPTIONS = ['站编', '站内', '货场', '中油', '中粮', '港务局', '天盛', '石化', '国投'];

  /* ==================== 状态 ==================== */
  // data: { date, period, signer, rows:[{time, source, content}] }；序号渲染时按行下标 +1 派生
  var data = defaultData();
  var fs = 15;                  // 表字号（不持久化，刷新重置）
  var FS_MIN = 8, FS_MAX = 22, FS_STEP = 1;

  function defaultData() {
    return { date: todayStr(), period: '', signer: '', note: '', rows: [] };
  }

  function todayStr() {
    var d = new Date();
    return d.getFullYear() + '-' + Utils.pad2(d.getMonth() + 1) + '-' + Utils.pad2(d.getDate());
  }

  /* ==================== 持久化 ==================== */
  function load() {
    var d = defaultData();
    var s = Store.get(STORE_KEY, null);
    if (s && typeof s === 'object') {
      if (s.date != null) d.date = s.date;
      if (s.period != null) d.period = s.period;
      if (s.signer != null) d.signer = s.signer;
      if (s.note != null) d.note = s.note;
      if (Array.isArray(s.rows)) d.rows = s.rows.map(function (r) {
        return { time: r.time || '', source: r.source || '', content: r.content || '' };
      });
    }
    data = d;
  }

  function save() {
    Store.set(STORE_KEY, { date: data.date, period: data.period, signer: data.signer, note: data.note, rows: data.rows });
  }

  /* ==================== 时间格式化 ==================== */
  // 取数字，末两位为「分」，其余为「时」（补零到两位）：1230→12:30，230→02:30
  function fmtTime(raw) {
    var s = String(raw == null ? '' : raw).replace(/\D/g, '');
    if (!s) return '';
    var min = s.slice(-2);
    var hr = s.slice(0, -2);
    if (!hr) hr = '0';
    return String(+hr).padStart(2, '0') + ':' + min.padStart(2, '0');
  }

  function escapeAttr(v) {
    return String(v == null ? '' : v).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
  }

  /* ==================== 渲染 ==================== */
  function render() {
    // 表头字段
    setVal('flowDate', data.date);
    setVal('flowPeriod', data.period);
    setVal('flowSigner', data.signer);
    setVal('flowNote', data.note);
    renderBody();
    applyFs();
  }

  function setVal(id, v) { var el = $(id); if (el) el.value = (v == null ? '' : v); }

  function renderBody() {
    var rows = data.rows;
    var html = rows.map(function (r, i) {
      var chosen = (r.source || '').split('+');
      var opts = SOURCE_OPTIONS.map(function (v) {
        var on = chosen.indexOf(v) >= 0;
        return '<div class="flow-src-opt' + (on ? ' on' : '') + '" data-v="' + v + '">' + v + '</div>';
      }).join('');
      return '<tr data-i="' + i + '" draggable="true">' +
        '<td class="flow-idx-cell">' + (i + 1) + '</td>' +
        '<td><input class="flow-time-input" data-f="time" value="' + escapeAttr(r.time) + '"></td>' +
        '<td class="flow-src-cell" data-i="' + i + '">' +
          '<div class="flow-src-trigger" title="可直接输入；点 + 选择（可多选）">' +
            '<input class="flow-src-manual" data-f="source" value="' + escapeAttr(r.source) + '">' +
            '<span class="flow-src-add">+</span>' +
          '</div>' +
          '<div class="flow-src-pop">' + opts +
            '<div class="flow-src-actions">' +
              '<button type="button" class="btn flow-src-reset">重置</button>' +
              '<button type="button" class="btn flow-src-ok">确定</button>' +
            '</div>' +
          '</div>' +
        '</td>' +
        '<td><textarea class="flow-content-input' + (r.content ? '' : ' flow-content-empty') + '" data-f="content" rows="1">' + escapeAttr(r.content) + '</textarea></td>' +
        '</tr>';
    }).join('');
    $('flowBody').innerHTML = html;
    $('flowEmpty').style.display = rows.length ? 'none' : '';
    refitAll();
  }

  // 按当前字号重算所有「来源输入宽度」「内容文本框高度」（渲染后 / 字号缩放后调用）
  function refitAll() {
    var body = $('flowBody');
    if (!body) return;
    var sis = body.querySelectorAll('.flow-src-manual');
    for (var i = 0; i < sis.length; i++) fitSrcWidth(sis[i]);
    var tas = body.querySelectorAll('textarea.flow-content-input');
    for (var j = 0; j < tas.length; j++) autoGrow(tas[j]);
  }

  function focusRow(i) {
    var tr = $('flowBody').querySelector('tr[data-i="' + i + '"]');
    if (tr) { var inp = tr.querySelector('input'); if (inp) inp.focus(); }
  }

  /* ==================== 字号缩放 ==================== */
  function applyFs() {
    var wrap = $('flowTableWrap');
    if (wrap) wrap.style.setProperty('--flow-fs', fs + 'px');
    var v = $('flowFsVal'); if (v) v.textContent = fs + 'px';
    var zi = $('flowZoomIn'), zo = $('flowZoomOut');
    if (zi) { zi.disabled = fs >= FS_MAX; zi.title = '放大表格字体（当前 ' + fs + 'px）'; }
    if (zo) { zo.disabled = fs <= FS_MIN; zo.title = '缩小表格字体（当前 ' + fs + 'px）'; }
    refitAll();                        // 字号变了，来源列宽 / 内容行高需重算
  }
  function zoom(delta) {
    var next = Math.min(FS_MAX, Math.max(FS_MIN, fs + delta));
    if (next === fs) return;
    fs = next; applyFs();
  }

  /* ==================== 打开 / 关闭 ==================== */
  function positionDrawer() {
    var tb = document.querySelector('.toolbar');
    var sb = document.querySelector('.statusbar');
    var top = tb ? tb.offsetHeight : 0;
    var bot = sb ? sb.offsetHeight : 0;
    var w = parseInt(document.documentElement.style.getPropertyValue('--flow-w'), 10) || 520;
    $('flowDrawer').style.top = top + 'px';
    $('flowDrawer').style.bottom = bot + 'px';
    $('flowSplitter').style.top = top + 'px';
    $('flowSplitter').style.bottom = bot + 'px';
    $('flowSplitter').style.left = w + 'px';
  }

  function open() {
    if (document.body.classList.contains('flow-drawer-open')) { close(); return; }  // 再次点击收起
    load();
    // 打开即保证至少 15 行（仅内存填充，用户编辑后才会写入 localStorage）
    while (data.rows.length < 15) data.rows.push({ time: '', source: '', content: '' });
    var w = Store.get(WIDTH_KEY, 520);
    if (typeof w !== 'number' || w < 360) w = 520;
    document.documentElement.style.setProperty('--flow-w', w + 'px');
    document.body.classList.add('flow-drawer-open');
    // 注意事项列被隐藏后，冻结列（分组/股道）的 sticky 偏移 --col-a-w 必须同步置 0，
    // 否则冻结列左侧残留该列宽度的空白
    var grid = $('grid');
    if (grid) grid.style.setProperty('--col-a-w', '0px');
    $('flowDrawer').style.display = 'flex';
    $('flowSplitter').style.display = 'block';
    positionDrawer();
    render();
  }

  function close() {
    document.body.classList.remove('flow-drawer-open');
    // 恢复冻结列偏移：注意事项列重新显示，按表头实际宽度回填 --col-a-w
    var grid = $('grid');
    var th = document.querySelector('#headRow th[data-col="note"]');
    if (grid) grid.style.setProperty('--col-a-w', (th && th.offsetWidth ? th.offsetWidth + 'px' : '0px'));
    $('flowDrawer').style.display = 'none';
    $('flowSplitter').style.display = 'none';
  }

  /* ==================== 分中线拖动 ==================== */
  var dragging = false, startX = 0, startW = 0;
  function onSplitterDown(e) {
    dragging = true; startX = e.clientX;
    startW = parseInt(document.documentElement.style.getPropertyValue('--flow-w'), 10) || 520;
    $('flowSplitter').classList.add('dragging');
    document.addEventListener('mousemove', onSplitterMove);
    document.addEventListener('mouseup', onSplitterUp);
    e.preventDefault();
  }
  function onSplitterMove(e) {
    if (!dragging) return;
    var w = Math.min(window.innerWidth * 0.7, Math.max(360, startW + (e.clientX - startX)));
    document.documentElement.style.setProperty('--flow-w', w + 'px');
    $('flowSplitter').style.left = w + 'px';
  }
  function onSplitterUp() {
    if (!dragging) return;
    dragging = false;
    $('flowSplitter').classList.remove('dragging');
    document.removeEventListener('mousemove', onSplitterMove);
    document.removeEventListener('mouseup', onSplitterUp);
    var w = parseInt(document.documentElement.style.getPropertyValue('--flow-w'), 10) || 520;
    Store.set(WIDTH_KEY, w);
  }

  /* ==================== 事件绑定（一次性） ==================== */
  function init() {
    if (!$('flowDrawer')) return;

    on('flowClose', 'click', close);
    on('flowZoomIn', 'click', function () { zoom(FS_STEP); });
    on('flowZoomOut', 'click', function () { zoom(-FS_STEP); });

    // 表头字段编辑
    on('flowDate', 'input', function () { data.date = this.value; save(); });
    on('flowPeriod', 'input', function () { data.period = this.value; save(); });
    on('flowSigner', 'input', function () { data.signer = this.value; save(); });
    on('flowNote', 'input', function () { data.note = this.value; save(); });

    // 截图 / 打印（复用 31814 的通用实现）
    on('flowCopyImg', 'click', function () {
      if (global.Report31814 && global.Report31814.copyElementAsImage) global.Report31814.copyElementAsImage($('flowTable'));
    });
    on('flowPrint', 'click', function () {
      if (global.Report31814 && global.Report31814.printElement) global.Report31814.printElement($('flowTable'), '钦州港车流计划表');
    });
    on('flowClear', 'click', clearAll);           // 清空表格内容（保留表头字段）

    // 表格体：录入（input 实时存原文，focusout 时时间列格式化）
    on('flowBody', 'input', onCellInput);
    on('flowBody', 'change', onCellInput);
    on('flowBody', 'focusout', onCellBlur);
    on('flowBody', 'keydown', onCellKey);          // 回车失焦（时间/编组内容）

    // 车流来源多选弹层（点击事件委托到 document，覆盖弹层内部与外部关闭）
    document.addEventListener('click', onDocClick);

    // 拖动排序（HTML5 DnD：dragstart / dragenter 真实换序 / dragover 允许放置 / drop / dragend 回写）
    on('flowBody', 'dragstart', onDragStart);
    on('flowBody', 'dragenter', onDragEnter);
    on('flowBody', 'dragover', function (e) { if (dragRow) e.preventDefault(); });   // 必须，否则不触发 drop
    on('flowBody', 'drop', onDragDrop);
    on('flowBody', 'dragend', onDragEnd);

    // 行右键菜单（在上方插入 / 删除本行 / 在下方插入）
    on('flowBody', 'contextmenu', onRowContextMenu);
    var ctxMenu = $('flowCtxMenu');
    if (ctxMenu) {
      ctxMenu.addEventListener('click', onCtxMenuClick);
      ctxMenu.addEventListener('contextmenu', function (e) { e.preventDefault(); });   // 菜单内不再弹系统菜单
    }
    document.addEventListener('click', hideCtxMenu);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') hideCtxMenu(); });
    var stw = $('flowTableWrap'); if (stw) stw.addEventListener('scroll', hideCtxMenu);

    // 车流表行选中（作为车站列表插入目标行）
    on('flowBody', 'click', onFlowRowClick);
    // 车站列表弹窗
    var picker = $('flowStationPicker');
    if (picker) picker.addEventListener('click', onStationPickerClick);
    var pHead = document.querySelector('.flow-st-head');   // 标题栏拖动窗口
    if (pHead) pHead.addEventListener('pointerdown', onPickerHeadDown);
    document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeStationList(); });
    // 点击窗口外（或点其他主表行）关闭车站列表
    document.addEventListener('pointerdown', function (e) {
      var pk = $('flowStationPicker');
      if (!pk || pk.hidden) return;
      if (e.target.closest && e.target.closest('#flowStationPicker')) return;   // 窗口内不关闭
      closeStationList();
    }, true);

    // 分中线
    on('flowSplitter', 'mousedown', onSplitterDown);
    window.addEventListener('resize', positionDrawer);
  }

  function onCellInput(e) {
    var inp = e.target;
    var f = inp.getAttribute && inp.getAttribute('data-f');
    if (!f) return;
    var tr = inp.closest ? inp.closest('tr') : null;
    if (!tr) return;
    var i = +tr.getAttribute('data-i');
    if (f === 'time') data.rows[i].time = inp.value;            // 输入期间保留原文，失焦再格式化
    else if (f === 'content') {
      data.rows[i].content = inp.value;
      inp.classList.toggle('flow-content-empty', !inp.value);
      autoGrow(inp);
    }
    else if (f === 'source') { data.rows[i].source = inp.value; fitSrcWidth(inp); }
    save();
  }

  // 编组内容（textarea）按内容自动撑高，避免长内容被截断 / 需要横向滚动
  function autoGrow(ta) {
    if (!ta || ta.tagName !== 'TEXTAREA') return;
    ta.style.height = 'auto';
    ta.style.height = ta.scrollHeight + 'px';
  }
  // 车流来源输入：按真实文字像素宽度自适应（中文字符约 2ch，用 canvas 精确测量）
  var _measureCtx = null;
  function measureText(txt, font) {
    if (!_measureCtx) _measureCtx = document.createElement('canvas').getContext('2d');
    _measureCtx.font = font;
    return _measureCtx.measureText(txt).width;
  }
  function fitSrcWidth(inp) {
    if (!inp || inp.tagName !== 'INPUT') return;
    var cs = window.getComputedStyle(inp);
    var font = cs.fontWeight + ' ' + cs.fontSize + ' ' + cs.fontFamily;
    var pad = (parseFloat(cs.paddingLeft) || 0) + (parseFloat(cs.paddingRight) || 0);
    var minW = measureText('站站', font);                       // 空值时的最小宽度（约两字）
    var w = Math.max(measureText(inp.value || '', font), minW);
    inp.style.width = Math.ceil(w + pad + 2) + 'px';
  }

  function onCellBlur(e) {
    var inp = e.target;
    if (!inp.getAttribute || inp.getAttribute('data-f') !== 'time') return;
    var tr = inp.closest ? inp.closest('tr') : null;
    if (!tr) return;
    var i = +tr.getAttribute('data-i');
    inp.value = fmtTime(inp.value);
    data.rows[i].time = inp.value;
    save();
  }

  /* 回车确认：时间列 → 格式化并跳到本行「车流来源」并打开选择器（键盘态）；
   * 来源选择器内按键交给 onSrcPopKey；编组内容是多行文本框，回车换行不拦截 */
  function onCellKey(e) {
    var inPop = e.target.closest ? e.target.closest('.flow-src-pop') : null;
    if (inPop && inPop === kbPop) { onSrcPopKey(e, inPop); return; }
    if (e.key !== 'Enter') return;
    var inp = e.target;
    var f = inp.getAttribute && inp.getAttribute('data-f');
    if (f !== 'time') return;
    e.preventDefault();
    var tr = inp.closest ? inp.closest('tr') : null;
    if (!tr) return;
    inp.value = fmtTime(inp.value);                 // 先格式化（1230 → 12:30）
    var i = +tr.getAttribute('data-i');
    if (data.rows[i]) { data.rows[i].time = inp.value; save(); }
    openSrcPopKb(tr.querySelector('.flow-src-cell'));
  }

  /* ===== 车流来源多选弹层（选项点选高亮，可再点取消；单元格也可直接手写） ===== */
  function srcCellOf(el) { return el.closest ? el.closest('.flow-src-cell') : null; }
  function syncSrcPop(pop, val) {
    var chosen = (val || '').split('+');
    var opts = pop.querySelectorAll('.flow-src-opt');
    for (var k = 0; k < opts.length; k++) {
      opts[k].classList.toggle('on', chosen.indexOf(opts[k].getAttribute('data-v')) >= 0);
    }
  }
  function toggleSrcPop(trig) {
    var cell = srcCellOf(trig);
    if (!cell) return;
    var pop = cell.querySelector('.flow-src-pop');
    if (!pop) return;
    var wasOpen = pop.classList.contains('open');
    closeSrcPop();
    if (!wasOpen) {
      var inp = cell.querySelector('.flow-src-manual');   // 打开时按当前手写值同步高亮
      syncSrcPop(pop, inp ? inp.value : '');
      pop.classList.add('open');
    }
  }
  function closeSrcPop() {
    var pops = $('flowBody').querySelectorAll('.flow-src-pop.open');
    for (var k = 0; k < pops.length; k++) pops[k].classList.remove('open');
    clearKbCursor();                                  // 一并清掉键盘态与光标高亮
  }

  /* ===== 来源选择器·键盘态 =====
   * 时间列回车 → 打开本行选择器并把焦点交给弹层，随后：
   *   ↑↓←→ 移动光标（3 列网格；左右在本行内循环，上下跨行保持列）
   *   Tab   切换光标处选项的选中/反选
   *   回车  确定写入，焦点转到本行「编组内容」
   *   Esc   放弃本次改动并关闭                                             */
  var kbPop = null, kbIdx = 0;

  function kbOpts(pop) { return pop.querySelectorAll('.flow-src-opt'); }

  function clearKbCursor() {
    var els = document.querySelectorAll('.flow-src-opt.kb-on, .flow-src-pop.kb');
    for (var k = 0; k < els.length; k++) els[k].classList.remove('kb-on', 'kb');
    kbPop = null;
  }

  function setKbIdx(pop, i) {
    var opts = kbOpts(pop), n = opts.length;
    if (!n) return;
    kbIdx = ((i % n) + n) % n;                        // 循环取位
    for (var k = 0; k < n; k++) opts[k].classList.toggle('kb-on', k === kbIdx);
    if (opts[kbIdx].scrollIntoView) opts[kbIdx].scrollIntoView({ block: 'nearest' });
  }

  /** 键盘方式打开本行来源选择器：光标停在第一个选项 */
  function openSrcPopKb(cell) {
    if (!cell) return;
    var pop = cell.querySelector('.flow-src-pop');
    if (!pop) return;
    closeSrcPop();
    var inp = cell.querySelector('.flow-src-manual');
    syncSrcPop(pop, inp ? inp.value : '');            // 先按当前值勾好，键盘在其基础上增删
    pop.classList.add('open', 'kb');
    pop.tabIndex = -1;
    pop.focus();                                      // 焦点必须给弹层：否则方向键/Tab 会被输入框吃掉
    kbPop = pop;                                      // 放在 closeSrcPop 之后赋值，避免刚开就被清
    setKbIdx(pop, 0);
  }

  function focusContent(tr) {
    var ta = tr ? tr.querySelector('.flow-content-input') : null;
    if (ta) ta.focus();
  }

  function onSrcPopKey(e, pop) {
    var key = e.key, opts = kbOpts(pop);
    if (key === 'Enter') {                            // 确定 → 写入 → 焦点转本行编组内容
      applySrc(pop);
      var tr = pop.closest ? pop.closest('tr') : null;
      closeSrcPop();
      focusContent(tr);
      e.preventDefault();
      return;
    }
    if (key === 'Escape') {                           // 放弃本次改动
      var cell = pop.closest ? pop.closest('.flow-src-cell') : null;
      closeSrcPop();
      var sin = cell ? cell.querySelector('.flow-src-manual') : null;
      if (sin) sin.focus();                           // 焦点交回来源手写框（否则会掉到 body）
      e.preventDefault();
      e.stopPropagation();                            // 优先关本弹层，不惊动其它 Esc 监听
      return;
    }
    if (key === 'Tab') {                              // 选择 / 反选光标处选项
      if (opts[kbIdx]) opts[kbIdx].classList.toggle('on');
      e.preventDefault();
      return;
    }
    if (key !== 'ArrowUp' && key !== 'ArrowDown' && key !== 'ArrowLeft' && key !== 'ArrowRight') return;
    var cols = 3, n = opts.length, rows = Math.ceil(n / cols);
    var row = Math.floor(kbIdx / cols), col = kbIdx % cols;
    if (key === 'ArrowLeft') col = (col + cols - 1) % cols;                                  // 行内左移
    else if (key === 'ArrowRight') col = (col + 1) % cols;                                   // 行内右移
    else row = (row + (key === 'ArrowUp' ? rows - 1 : 1)) % rows;                            // 跨行、保持列
    setKbIdx(pop, Math.min(row * cols + col, n - 1));
    e.preventDefault();
  }
  function writeSrc(cell, joined) {
    var i = +cell.getAttribute('data-i');
    data.rows[i].source = joined;
    var inp = cell.querySelector('.flow-src-manual');
    if (inp) { inp.value = joined; fitSrcWidth(inp); }
    save();
  }
  function applySrc(pop) {
    var cell = srcCellOf(pop);
    if (!cell) return;
    var on = pop.querySelectorAll('.flow-src-opt.on');
    var vals = [];
    for (var k = 0; k < on.length; k++) vals.push(on[k].getAttribute('data-v'));
    writeSrc(cell, vals.join('+'));
  }
  function resetSrc(pop) {
    var cell = srcCellOf(pop);
    if (!cell) return;
    var on = pop.querySelectorAll('.flow-src-opt.on');
    for (var k = 0; k < on.length; k++) on[k].classList.remove('on');
    writeSrc(cell, '');
  }
  function onDocClick(e) {
    var pop = e.target.closest ? e.target.closest('.flow-src-pop') : null;
    if (pop) {
      var opt = e.target.closest('.flow-src-opt');
      if (opt) { opt.classList.toggle('on'); return; }   // 点选高亮 / 再点取消
      if (e.target.closest('.flow-src-ok')) { applySrc(pop); closeSrcPop(); }
      else if (e.target.closest('.flow-src-reset')) { resetSrc(pop); closeSrcPop(); }
      return;   // 点击弹层内部不关闭
    }
    var trig = e.target.closest ? e.target.closest('.flow-src-trigger') : null;
    // 点输入框区域=手写定位，不展开弹层；点「+」或单元格其余空白处才展开
    if (trig && !e.target.closest('input')) { toggleSrcPop(trig); return; }
    closeSrcPop();
  }

  /* ==================== 行拖动排序（HTML5 DnD：真实换序 + FLIP 位移动画） ====================
   * 与「插入指示线」方案的区别：拖过哪一行就立刻真实换序，行用 FLIP 从旧位置平滑滑到新位置，
   * 被拖行原地留一个虚线空位（.flow-moving）。松手时把 DOM 顺序回写进 data.rows 并重排序号。 */
  var dragRow = null, flip = null, dragMoved = false, dragGhost = null;

  /** 拖拽影像：一个与行等大的纯色条（不显示文字，自然不会与下方文字重影）
   *  放在视口外渲染——不能 display:none，否则浏览器拍不到。 */
  var dragGhost = null;
  function makeDragImage(r) {
    if (dragGhost && dragGhost.parentNode) dragGhost.parentNode.removeChild(dragGhost);
    var el = document.createElement('div');
    el.style.cssText = 'position:fixed;left:-10000px;top:0;pointer-events:none;box-sizing:border-box;' +
      'width:' + r.width + 'px;height:' + r.height + 'px;' +
      'background:#8fb2ff;border:1px solid #3f78ff;border-radius:3px;';
    document.body.appendChild(el);
    dragGhost = el;
    return el;
  }

  function onDragStart(e) {
    var tr = e.target.closest ? e.target.closest('tr') : null;
    if (!tr) return;
    dragRow = tr;
    dragMoved = false;
    flip = global.Flip ? global.Flip.capture($('flowBody').children) : null;   // 记录 FIRST 位置
    if (e.dataTransfer) {
      var r = tr.getBoundingClientRect();
      var img = makeDragImage(r);
      try { e.dataTransfer.setDragImage(img, Math.max(0, e.clientX - r.left), Math.max(0, e.clientY - r.top)); } catch (_) {}
      e.dataTransfer.effectAllowed = 'move';
      try { e.dataTransfer.setData('text/plain', '1'); } catch (_) {}   // 必须 setData 才可拖
    }
    // 被拖行自身也让位：延后加类，避免影响浏览器先拍下的影像
    setTimeout(function () { if (dragRow) dragRow.classList.add('flow-moving'); }, 0);
  }

  /** 拖过某行 → 真实换序（向下拖插到目标之后 / 向上拖插到目标之前）→ FLIP 平滑位移 */
  function onDragEnter(e) {
    if (!dragRow) return;
    var t = e.target.closest ? e.target.closest('tr[data-i]') : null;
    if (!t || t === dragRow) return;
    var body = $('flowBody');
    var kids = body.children;
    var from = Array.prototype.indexOf.call(kids, dragRow);
    var to = Array.prototype.indexOf.call(kids, t);
    if (from < 0 || to < 0 || from === to) return;
    body.insertBefore(dragRow, from < to ? t.nextSibling : t);
    dragMoved = true;
    if (flip) flip.play();
  }

  function onDragDrop(e) { e.preventDefault(); }   // 顺序已在 dragenter 里改好，这里只需允许放置

  function onDragEnd() {
    if (dragRow) dragRow.classList.remove('flow-moving');
    if (dragGhost && dragGhost.parentNode) dragGhost.parentNode.removeChild(dragGhost);   // 影像快照已取完，移除离屏节点
    dragGhost = null;
    if (flip) flip.play();                        // 动画未跑完也先归位
    var moved = dragMoved;
    dragRow = null; flip = null; dragMoved = false;
    if (moved) commitOrder();                     // 有换序才回写并重排序号（顺带清掉内联 transform）
  }

  /** DOM 顺序 → data.rows 顺序（用 data-i 反查原对象），并**就地**重排序号。
   *  刻意不 renderBody()：整表重渲染会重建行元素、打断还在跑的 FLIP 位移动画（松手瞬间会「跳一下」）。 */
  function commitOrder() {
    var nodes = $('flowBody').querySelectorAll('tr[data-i]');
    var next = [];
    for (var i = 0; i < nodes.length; i++) {
      var tr = nodes[i];
      next.push(data.rows[+tr.getAttribute('data-i')]);
      tr.setAttribute('data-i', i);                       // 行号就地更新（输入事件按它索引 data.rows）
      var sc = tr.querySelector('.flow-src-cell');
      if (sc) sc.setAttribute('data-i', i);               // 来源单元格也带 data-i
      if (tr.firstElementChild) tr.firstElementChild.textContent = i + 1;
    }
    data.rows = next;
    save();
  }

  /* ==================== 行右键菜单 ==================== */
  var ctxIdx = -1;
  function onRowContextMenu(e) {
    var tr = e.target.closest ? e.target.closest('tr[data-i]') : null;
    if (!tr) return;                       // 非数据行（表头/表尾）不弹菜单
    e.preventDefault();
    ctxIdx = +tr.getAttribute('data-i');
    var menu = $('flowCtxMenu');
    if (!menu) return;
    menu.hidden = false;
    var mw = menu.offsetWidth, mh = menu.offsetHeight;
    menu.style.left = Math.min(e.clientX, window.innerWidth - mw - 8) + 'px';
    menu.style.top = Math.min(e.clientY, window.innerHeight - mh - 8) + 'px';
  }
  function onCtxMenuClick(e) {
    var item = e.target.closest ? e.target.closest('.flow-ctx-item') : null;
    if (!item) return;
    var act = item.getAttribute('data-act');
    if (ctxIdx >= 0) {
      if (act === 'above') insertRowAt(ctxIdx);
      else if (act === 'below') insertRowAt(ctxIdx + 1);
      else if (act === 'del') delRowAt(ctxIdx);
    }
    hideCtxMenu();
  }
  function hideCtxMenu() {
    var menu = $('flowCtxMenu'); if (menu) menu.hidden = true;
    ctxIdx = -1;
  }
  function insertRowAt(i) {
    data.rows.splice(i, 0, { time: '', source: '', content: '' });
    save(); renderBody(); focusRow(i);
  }
  function delRowAt(i) {
    if (!data.rows.length) { Utils.toast('没有可删除的行', 'error'); return; }
    data.rows.splice(i, 1);
    while (data.rows.length < 15) data.rows.push({ time: '', source: '', content: '' });  // 删除后不足 15 行则补够
    save(); renderBody();
  }

  /** 清空表格内容：行内「时间 / 来源 / 内容」与表尾备注一并清掉，仍保留 15 行空表；
   *  表头字段（预计日期 / 时段 / 站调）属于计划表抬头，不在表格内，故保留。 */
  function clearAll() {
    data.rows = [];
    while (data.rows.length < 15) data.rows.push({ time: '', source: '', content: '' });
    data.note = '';
    setVal('flowNote', '');
    save();
    closeSrcPop();
    renderBody();
    closeStationList();
    Utils.toast('已清空车流统计表内容', 'success');
  }

  /* ==================== 车流表行选中（车站列表插入目标） ==================== */
  function onFlowRowClick(e) {
    var tr = e.target.closest ? e.target.closest('tr[data-i]') : null;
    if (!tr) return;
    var sels = $('flowBody').querySelectorAll('tr.flow-selected');
    for (var k = 0; k < sels.length; k++) sels[k].classList.remove('flow-selected');
    tr.classList.add('flow-selected');
  }

  /* ==================== 车站列表弹窗（主表选中行的到站 → 插入编组内容） ==================== */
  var stationItems = [];
  function openStationList(items, origin) {
    if (!items || !items.length) { Utils.toast('该主表行无到站内容', 'error'); return; }
    stationItems = items;
    var list = $('flowStationList');
    if (!list) return;
    list.innerHTML = items.map(function (it, i) {
      return '<label class="flow-st-item"><input type="checkbox" checked data-i="' + i + '">' +
        Utils.escapeHtml(it) + '</label>';
    }).join('');
    var picker = $('flowStationPicker');
    if (picker) {
      picker.hidden = false;
      // 弹出在鼠标位置（origin 存在时），并夹在视口内避免溢出
      var w = picker.offsetWidth || 320, h = picker.offsetHeight || 360;
      var x = origin ? origin.x + 12 : Math.round((window.innerWidth - w) / 2);
      var y = origin ? origin.y + 12 : 90;
      x = Math.max(8, Math.min(x, window.innerWidth - w - 8));
      y = Math.max(8, Math.min(y, window.innerHeight - h - 8));
      picker.style.left = x + 'px';
      picker.style.top = y + 'px';
    }
  }
  function closeStationList() {
    var picker = $('flowStationPicker');
    if (picker) picker.hidden = true;
  }
  function onStationPickerClick(e) {
    var t = e.target;
    var act = t.getAttribute && t.getAttribute('data-act');
    if (act === 'close') { closeStationList(); return; }   // 仅关闭按钮（无遮罩，点窗口外不关闭）
    if (act === 'all') doStationInsert(true);
    else if (act === 'part') doStationInsert(false);
  }

  /* 标题栏拖动浮动窗（无遮罩，窗口外仍可操作） */
  var pickerDrag = null;
  function onPickerHeadDown(e) {
    if (e.target.closest && e.target.closest('.flow-st-close')) return;   // 关闭按钮不触发拖动
    var picker = $('flowStationPicker');
    if (!picker) return;
    var r = picker.getBoundingClientRect();
    pickerDrag = { dx: e.clientX - r.left, dy: e.clientY - r.top };
    document.addEventListener('pointermove', onPickerMove, true);
    document.addEventListener('pointerup', onPickerUp, true);
    e.preventDefault();
  }
  function onPickerMove(e) {
    if (!pickerDrag) return;
    var picker = $('flowStationPicker'); if (!picker) return;
    var x = e.clientX - pickerDrag.dx, y = e.clientY - pickerDrag.dy;
    x = Math.max(0, Math.min(x, window.innerWidth - picker.offsetWidth));
    y = Math.max(0, Math.min(y, window.innerHeight - picker.offsetHeight));
    picker.style.left = x + 'px';
    picker.style.top = y + 'px';
  }
  function onPickerUp() {
    pickerDrag = null;
    document.removeEventListener('pointermove', onPickerMove, true);
    document.removeEventListener('pointerup', onPickerUp, true);
  }
  function doStationInsert(all) {
    var tr = $('flowBody').querySelector('tr.flow-selected');      // 以当前高亮行为目标，避免索引错位
    if (!tr) { Utils.toast('请先在车流统计表中选择要插入的行', 'error'); return; }
    var idx = +tr.getAttribute('data-i');
    var items = stationItems;
    if (!all) {
      var checked = $('flowStationList').querySelectorAll('input:checked');
      items = [];
      for (var k = 0; k < checked.length; k++) items.push(stationItems[+checked[k].getAttribute('data-i')]);
    }
    if (!items || !items.length) { Utils.toast('没有可插入的项', 'error'); return; }
    var row = data.rows[idx];
    if (!row) { Utils.toast('目标行不存在', 'error'); return; }
    var add = items.join(' ');
    row.content = row.content ? (row.content + ' ' + add) : add;   // 追加到编组内容列
    save();
    renderBody();
    var tr2 = $('flowBody').querySelector('tr[data-i="' + idx + '"]');
    if (tr2) tr2.classList.add('flow-selected');                  // 重渲染后恢复高亮
    closeStationList();
    Utils.toast('已插入到车流表第 ' + (idx + 1) + ' 行编组内容', 'success');
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

  global.FlowTable = { open: open, close: close, openStationList: openStationList, closeStationList: closeStationList };
})(window);
