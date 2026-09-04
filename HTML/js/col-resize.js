/**
 * col-resize.js —— 表格列宽拖拽组件
 * ============================================================================
 * 从 app.js 抽出的通用表格能力，与「股道存车」业务无关，可复用于任意表格。
 * 暴露全局：window.ColResize
 * 加载顺序：必须在 app.js 之前
 *
 * 用法：
 *   var h = ColResize.enable(table, {
 *     persistKey: 'zhancun.grid.cols',   // 传了才记忆列宽（localStorage）
 *     onResize: function (th, w) { ... }
 *   });
 *   h.remount();      // 表头被 innerHTML 重建后重新挂手柄并恢复列宽
 *   h.autoFitAll();   // 全部列按内容自适应
 *   h.reset();        // 清除记忆并重新自适应
 *   ColResize.get(table);   // 取句柄，未启用过返回 null
 *   ColResize.getOrNull(table).remount();  // 调用方无需再写 if (x && x.__xxx)
 * ============================================================================
 *
 * 设计（从表格基础布局模型解决，而非打补丁）：
 *  - 拖拽手柄是 th 内的真实 <div class="col-handle">，而非伪元素 ::after 或坐标遍历。
 *    sticky 左列(col-a/col-b)会浮在普通列上方遮挡 e.target，伪元素/坐标法都会命中错乱
 *    （旧 bug：第三列起选不中）。真实手柄设 z-index 高于 sticky 列，命中 100% 准确。
 *  - 表格 width:max-content + table-layout:fixed，列宽由 th.style.width 决定，
 *    浏览器不会自作主张放大/压缩列（根治"越拖越大、其它列也变"）。
 *  - 主表/明细表均无 transform:scale 缩放，mousemove 用 clientX 差值即像素，天然跟手，
 *    无需 tableScale 归一化。
 *  - document 上的 mousemove/mouseup 只在拖拽期间挂载，拖完即卸载，
 *    不再常驻（旧实现每张表都留一对永久监听，多表叠加永不释放）。
 */
(function (global) {
  'use strict';

  var MIN = 30;            // 最小列宽

  /** 表格 → 句柄。用 WeakMap 而非往 DOM 上挂 __xxx 属性，避免污染与命名冲突 */
  var handles = new WeakMap();

  /** 当前正在拖拽的会话（全局唯一：同一时刻只会拖一根分隔线） */
  var drag = null;

  /* ==================== 拖拽期间的全局监听（按需挂载/卸载） ==================== */

  function onDocMouseMove(e) {
    if (!drag) return;
    // 无缩放，clientX 差值即像素差，分隔线直接跟手
    var dx = e.clientX - drag.startX;
    var w = Math.max(MIN, Math.round(drag.startW + dx));
    drag.th.style.width = w + 'px';
    if (drag.onResize) drag.onResize(drag.th, w);
  }

  function onDocMouseUp() {
    if (!drag) return;
    var h = drag.handle;
    if (h) h.classList.remove('active');
    var persist = drag.persist;
    drag = null;
    document.body.classList.remove('col-resize-active');
    document.removeEventListener('mousemove', onDocMouseMove);
    document.removeEventListener('mouseup', onDocMouseUp);
    persist();          // 拖拽结束即记忆，刷新/重渲染不丢失
  }

  function beginDrag(session) {
    drag = session;
    document.body.classList.add('col-resize-active');
    document.addEventListener('mousemove', onDocMouseMove);
    document.addEventListener('mouseup', onDocMouseUp);
  }

  /* ==================== 启用 ==================== */

  /**
   * 为表格启用列宽拖拽
   * @param {HTMLTableElement} table
   * @param {Object} [opts]
   * @param {string}   [opts.persistKey] 列宽记忆键（localStorage），不传则不记忆
   * @param {Function} [opts.onResize]   宽度变化回调 (th, width)
   * @returns {Object|null} 句柄 { remount, autoFitAll, reset }；表格不合规则返回 null
   */
  function enable(table, opts) {
    opts = opts || {};
    var onResize = opts.onResize || null;
    var persistKey = opts.persistKey || '';

    var thead = table && table.tHead;
    if (!thead) return null;
    var headRow = thead.rows && thead.rows[0];
    if (!headRow) return null;

    /* ---------------- 持久化 ---------------- */

    /** 读取记忆列宽。返回数组或 null（长度不符视为失效） */
    function loadStored() {
      if (!persistKey) return null;
      var arr = (global.Store ? Store.get(persistKey, null) : null);
      return (arr && arr.length === headRow.cells.length) ? arr : null;
    }

    /** 持久化当前列宽（逐列都记，便于完整还原） */
    function persistWidths() {
      if (!persistKey) return;
      var widths = [];
      for (var i = 0; i < headRow.cells.length; i++) {
        widths.push(Math.round(headRow.cells[i].offsetWidth));
      }
      if (global.Store) Store.set(persistKey, widths);
    }

    /* ---------------- 列宽测量 ---------------- */

    /** 测量结果缓存：DOM 结构/内容未变时复用，避免重复强制 reflow。
     *  key 由「行数 + 各列文本指纹」组成，任一处变化即失效（见 invalidateMeasure）。 */
    var _measureCache = null;

    function invalidateMeasure() { _measureCache = null; }

    /** 轻量指纹：用 tbody 各单元格文本拼接的简易哈希，变化即重测 */
    function contentFingerprint() {
      var f = headRow.cells.length + '|';
      var body = table.tBodies && table.tBodies[0];
      if (!body) return f;
      var rows = body.rows;
      var n = Math.min(rows.length, 50);   // 取前 50 行足够反映内容宽度变化
      for (var r = 0; r < n; r++) {
        var cells = rows[r].cells;
        for (var c = 0; c < cells.length; c++) f += (cells[c].textContent || '').length + ',';
      }
      return f;
    }

    /** 临时切到 auto + 全 nowrap，逼浏览器按「每列最长单行内容」算出最真实的列宽。
     *  同步测量，不会触发可见重绘。返回各列 offsetWidth 数组。
     *  测量完后恢复每列原来的内联 width，避免调用方只写一列时其它列丢失宽度。
     *  带结果缓存：同结构同内容仅强制 reflow 一次。 */
    function measureContentWidths() {
      var fp = contentFingerprint();
      if (_measureCache && _measureCache.fp === fp) return _measureCache.widths;

      var prevLayout = table.style.tableLayout;
      var prevWidth = table.style.width;
      var prevWidths = [];
      for (var i = 0; i < headRow.cells.length; i++) prevWidths.push(headRow.cells[i].style.width);
      table.classList.add('measuring');
      table.style.tableLayout = 'auto';
      table.style.width = 'auto';
      for (var i = 0; i < headRow.cells.length; i++) headRow.cells[i].style.width = '';
      var widths = [];
      for (var j = 0; j < headRow.cells.length; j++) widths.push(headRow.cells[j].offsetWidth);
      table.classList.remove('measuring');
      table.style.tableLayout = prevLayout;
      table.style.width = prevWidth;
      for (var i = 0; i < headRow.cells.length; i++) headRow.cells[i].style.width = prevWidths[i];

      _measureCache = { fp: fp, widths: widths };
      return widths;
    }

    /** 「自适应列宽」：所有列各自按真实内容宽重测并写精确 px（一次性全部自适应）。 */
    function autoFitAll() {
      var widths = measureContentWidths();
      for (var m = 0; m < headRow.cells.length; m++) {
        var cell = headRow.cells[m];
        cell.style.width = Math.max(MIN, widths[m]) + 'px';   // 每列独立
        if (onResize) onResize(cell, cell.offsetWidth);
      }
    }

    /** 双击分隔线 → 只有当前这一列恢复为内容自适应（像 Excel：双击列边=该列自适应），
     *  其它列宽度完全不动。 */
    function autoFit(th) {
      var idx = Array.prototype.indexOf.call(th.parentNode.children, th);
      var widths = measureContentWidths();
      th.style.width = Math.max(MIN, widths[idx]) + 'px';   // 只写当前列
      if (onResize) onResize(th, th.offsetWidth);
    }

    /* ---------------- 应用与重置 ---------------- */

    /** 应用记忆列宽到各 th（拖拽/刷新/重渲染后调用，保证不丢失）。
     *  无记忆时执行首次内容自适应（像 Excel 打开即按内容宽），每列都独立精确 px。 */
    function applyStoredWidths() {
      var arr = loadStored();
      if (!arr) { autoFitAll(); return; }    // 首次：内容自适应
      for (var i = 0; i < headRow.cells.length; i++) {
        var cell = headRow.cells[i];
        var w = arr[i];
        if (typeof w === 'number' && w >= MIN) {
          cell.style.width = w + 'px';
          if (onResize) onResize(cell, w);
        }
      }
    }

    /** 重置为自适应（清空拖动记忆，恢复到按内容计算的列宽） */
    function reset() {
      invalidateMeasure();        // 数据/列宽被重置，强制下次重测
      if (persistKey && global.Store) Store.remove(persistKey);
      autoFitAll();
      persistWidths();
    }

    /* ---------------- 拖拽手柄 ---------------- */

    function onHandleDown(e) {
      var handle = e.currentTarget;
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();               // 阻止冒泡到行的单击选中逻辑
      handle.classList.add('active');
      beginDrag({
        th: handle.parentElement,        // 手柄即 th 的直接子元素
        handle: handle,
        startX: e.clientX,
        startW: handle.parentElement.offsetWidth,
        onResize: onResize,
        persist: persistWidths
      });
    }

    function onHandleDbl(e) {
      e.preventDefault();
      e.stopPropagation();
      autoFit(e.currentTarget.parentElement);
    }

    /** 在每个 th 右缘挂载真实拖拽手柄（最后一列不加，无下一列可拖） */
    function mountHandles() {
      var cells = headRow.cells;
      for (var i = 0; i < cells.length; i++) {
        var th = cells[i];
        if (th.classList.contains('no-resize')) continue;          // 禁止单独拖拽的列（如合并表头的左半）
        if (th.querySelector(':scope > .col-handle')) continue;   // 去重
        var h = document.createElement('div');
        h.className = 'col-handle' + (i === cells.length - 1 ? ' last' : '');
        h.addEventListener('mousedown', onHandleDown);
        h.addEventListener('dblclick', onHandleDbl);
        th.appendChild(h);
      }
    }

    /** 表头被 innerHTML 重建后：重新挂手柄 + 恢复记忆列宽 */
    function remount() {
      invalidateMeasure();        // 表头重建后内容指纹变化，必须失效缓存
      mountHandles();
      applyStoredWidths();
    }

    var handle = {
      remount: remount,
      autoFitAll: autoFitAll,
      reset: reset
    };

    remount();               // 首次：挂手柄 + 应用记忆（无记忆则内容自适应）
    handles.set(table, handle);
    return handle;
  }

  /* ==================== 对外接口 ==================== */

  /**
   * 取表格的句柄。
   * 未启用过返回 null；调用方写 ColResize.get(t) && ColResize.get(t).remount() 略啰嗦，
   * 故额外提供 safeGet() 返回一个"什么都不做"的空句柄，调用方可直接链式调用。
   */
  function get(table) { return table ? handles.get(table) || null : null; }

  var NOOP_HANDLE = { remount: function () {}, autoFitAll: function () {}, reset: function () {} };
  function safeGet(table) { return get(table) || NOOP_HANDLE; }

  global.ColResize = {
    enable: enable,
    get: get,
    safeGet: safeGet,
    MIN: MIN
  };
})(window);
