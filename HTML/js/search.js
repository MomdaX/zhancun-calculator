/**
 * search.js —— 全站车辆搜索（独立抽屉）
 * ============================================================================
 * 暴露全局：window.Search
 * 加载顺序：必须在 app.js 之后（依赖 window.YardApp 提供的取数 / 渲染 / 定位能力）
 *
 * 两种模式：
 *   车号  只匹配车号；输入完整车号且唯一命中时，展开该车所在整条股道并定位到该行，
 *         便于查看这辆车在股道中的位置与前后邻车。
 *   内容  匹配 车种 / 到站（处理后的 __dest）/ 品名 / 记事 / 收货人 / 车次。
 *
 * 多关键词按空格拆分，任一命中（OR）即列出。
 * 搜索范围为全部车辆（不受「隐藏空线分组」开关影响）。
 * 结果行点击 → 打开该股道明细并高亮定位到这辆车（搜索抽屉保留在后面，关掉明细即回到结果）。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Utils = global.Utils;
  var UI = global.UI;
  if (!Utils || !UI) return;

  var $ = Utils.$;
  var on = Utils.on;
  var COL = global.Aggregate ? global.Aggregate.COL : null;

  var mode = 'carno';     // 'carno' | 'text'
  var carnoMatch = 'contain';   // 'start'（开头） | 'contain'（包含），仅车号模式生效
  var carnoStartEl = null, carnoContainEl = null;

  /* ==================== 取数 / 匹配 ==================== */

  function cell(row, col) {
    var v = row[col];
    return v == null ? '' : String(v).trim();
  }

  /**
   * 内容模式的待搜文本：车种 / 到站（处理后的 __dest）/ 品名 / 记事 / 收货人 / 车次。
   * 到站按要求只取处理后的 __dest，不使用原始到站。
   */
  function haystack(row) {
    var dest = row.__dest == null ? '' : String(row.__dest).trim();
    return [cell(row, COL.CARTYPE), row.__carType || '', dest,
            cell(row, COL.GOODS), cell(row, COL.NOTE),
            cell(row, COL.CONSIGNEE), cell(row, COL.TRAIN)]
           .join(' ').toLowerCase();
  }

  /** 命中判定：多关键词 OR（任一命中即算匹配）。车号模式支持「开头 / 包含」两种匹配方式 */
  function matched(row, keys) {
    var hay, isCarno = (mode === 'carno');
    if (isCarno) hay = cell(row, COL.CARNO).toLowerCase();
    else hay = haystack(row);
    for (var i = 0; i < keys.length; i++) {
      var k = keys[i];
      if (isCarno && carnoMatch === 'start') {
        if (hay.indexOf(k) === 0) return true;          // 车号以关键词开头
      } else {
        if (hay.indexOf(k) >= 0) return true;            // 包含（车号 / 内容通用）
      }
    }
    return false;
  }

  /** 展平全站车辆 → [{ trackIdx, carIdx, row }]（trackIdx 可直接用于 openDetailAt） */
  function allCars() {
    var rows = global.YardApp ? global.YardApp.getRows() : [];
    var out = [];
    rows.forEach(function (r, ti) {
      var list = r.raw || [];
      for (var i = 0; i < list.length; i++) out.push({ trackIdx: ti, carIdx: i, row: list[i] });
    });
    return out;
  }

  /* ==================== 渲染 ==================== */

  function clearResult() {
    $('searchBody').innerHTML = '';
    $('searchCount').textContent = '共 0 条';
    $('searchHint').classList.remove('hide');
    $('searchTable').style.display = 'none';
  }

  /**
   * 渲染结果表（复用主表的明细渲染，保证列序 / 着色 / 格式与明细抽屉完全一致）
   * @param {Array} list 实际展示的车辆（可能是命中集，也可能是整条股道）
   * @param {Array} hits 命中集（用于给命中行加高亮）
   * @param {number} focusCarIdx 需要滚动定位的车辆下标（-1 为不定位）
   */
  function render(list, hits, focusCarIdx) {
    global.YardApp.renderRows(list.map(function (c) { return c.row; }), {
      head: $('searchHead'), body: $('searchBody'), table: $('searchTable')
    }, {
      destProcessed: false,     // 搜索抽屉显示与明细抽屉一致（原始到站/发站）
      rowAttr: function (row, i) {
        var c = list[i];
        return ' data-track-idx="' + c.trackIdx + '" data-car-i="' + c.carIdx + '"';
      }
    });

    $('searchHint').classList.add('hide');
    $('searchTable').style.display = '';
    $('searchCount').textContent = '共 ' + hits.length + ' 条';

    if (focusCarIdx >= 0) {
      var tr = $('searchBody').querySelector('tr[data-car-i="' + focusCarIdx + '"]');
      if (tr && tr.scrollIntoView) tr.scrollIntoView({ block: 'center' });
    }
  }

  /* ==================== 搜索主流程 ==================== */

  function run() {
    if (!global.YardApp || !global.YardApp.hasData()) return;
    var raw = $('searchInput').value.trim();
    var keys = raw ? raw.toLowerCase().split(/\s+/).filter(Boolean) : [];

    if (!keys.length) { clearResult(); return; }

    var hits = allCars().filter(function (c) { return matched(c.row, keys); });

    // 车号模式 + 唯一命中（通常即输入了完整车号）→ 展示整条股道并定位到该车
    var list = hits, focusCarIdx = -1;
    if (mode === 'carno' && hits.length === 1) {
      var rows = global.YardApp.getRows();
      var ti = hits[0].trackIdx;
      list = (rows[ti].raw || []).map(function (row, i) {
        return { trackIdx: ti, carIdx: i, row: row };
      });
      focusCarIdx = hits[0].carIdx;
    }

    render(list, hits, focusCarIdx);
  }

  /* ==================== 模式 / 开关 ==================== */

  function setMode(m) {
    mode = m;
    var btns = $('searchModes').querySelectorAll('.mode-btn');
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle('active', btns[i].getAttribute('data-mode') === m);
    }
    $('searchInput').placeholder = (m === 'carno')
      ? '输入车号（完整车号可展开整条股道并定位）'
      : '输入 车种 / 到站 / 品名 / 记事 / 收货人 / 车次';
    // 开头 / 包含 仅对车号模式有意义，内容模式禁用
    var carno = (m === 'carno');
    if (carnoStartEl) {
      carnoStartEl.disabled = !carno;
      carnoContainEl.disabled = !carno;
      carnoStartEl.closest('.carno-opt').classList.toggle('disabled', !carno);
      carnoContainEl.closest('.carno-opt').classList.toggle('disabled', !carno);
    }
    run();
  }

  function open() {
    if (!global.YardApp || !global.YardApp.hasData()) {
      Utils.toast('请先加载 xls 数据', 'error');
      return;
    }
    UI.Drawer.open('searchDrawer');
    $('searchInput').focus();
    if ($('searchInput').value.trim()) run();   // 数据可能已刷新，重搜一次保持结果最新
  }

  /* ==================== 初始化 ==================== */

  function init() {
    if (!$('searchDrawer') || !COL) return;

    on('searchModes', 'click', function (e) {
      var btn = e.target.closest ? e.target.closest('.mode-btn') : null;
      if (btn) setMode(btn.getAttribute('data-mode'));
    });

    // 车号匹配方式：开头 / 包含（视觉为复选框，行为互斥且必选其一）
    carnoStartEl = $('carnoOptStart');
    carnoContainEl = $('carnoOptContain');
    function onCarnoOptChange(e) {
      if (e.target === carnoStartEl && carnoStartEl.checked) carnoContainEl.checked = false;
      if (e.target === carnoContainEl && carnoContainEl.checked) carnoStartEl.checked = false;
      // 不允许两者都不勾：取消其中一个时自动保留另一个
      if (!carnoStartEl.checked && !carnoContainEl.checked) {
        (e.target === carnoStartEl ? carnoStartEl : carnoContainEl).checked = true;
      }
      carnoMatch = carnoStartEl.checked ? 'start' : 'contain';
      run();
    }
    on('carnoOptStart', 'change', onCarnoOptChange);
    on('carnoOptContain', 'change', onCarnoOptChange);
    on('searchInput', 'input', Utils.debounce(run, 150));
    on('btnSearchClear', 'click', function () {
      $('searchInput').value = '';
      clearResult();
      $('searchInput').focus();
    });
    on('btnCloseSearch', 'click', function () { UI.Drawer.close('searchDrawer'); });

    // 点结果行 → 打开该股道明细并高亮定位这辆车（搜索抽屉不关闭，仍在后面）
    on('searchBody', 'click', function (e) {
      var tr = e.target.closest ? e.target.closest('tr') : null;
      if (!tr) return;
      // 车号 / 车站单元格的交互走双击（复制 / 开地图），单击不打开明细，避免与双击冲突
      if (e.target.closest('td[data-col="carno"]') || e.target.closest('span.station-link')) return;
      var ti = tr.getAttribute('data-track-idx');
      if (ti == null) return;
      global.YardApp.openDetailAt(+ti, +tr.getAttribute('data-car-i'));
    });

    setMode('carno');

    // 数据刷新（重新加载/切换 xls）后，若搜索抽屉打开且已输入条件，自动重跑保持结果最新
    if (global.YardApp && global.YardApp.onDataChange) {
      global.YardApp.onDataChange(function () {
        if (UI.Drawer.isOpen('searchDrawer') && $('searchInput').value.trim()) run();
      });
    }
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();

  global.Search = { open: open };
})(window);
