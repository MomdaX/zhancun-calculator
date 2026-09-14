/**
 * est-drag.js —— 「计重编辑：按住拖动批量删除推算载重」通用交互（明细表 / 推演面板共用）
 * ============================================================================
 * 交互（两侧效果逐项对齐，修复只改这一处）：
 *   · 双侧范围：锚点固定在按下的推算格所在行，范围 = 锚点 ↔ 当前指针（上/下方都有效），
 *     范围内按 mode 统一应用（起点未删 → 批量删；起点已删 → 批量恢复），范围外按快照复原
 *   · 指针回到锚点：锚点..上次指针行 整段按快照恢复（快速滑动跳行的中间行一并覆盖）
 *   · 实载行（无推算格）自动跳过；拖动结束后的 click 用 moved() 压制，不与单击切换打架
 *
 * 使用（明细 app.js / 推演 sim-panel.js 各建一个实例）：
 *   var dragger = EstDrag.create(cfg);   // 装配时创建，cfg 见下
 *   dragger.bind();                      // tbody 挂 mousedown/mouseover，document 挂 mouseup
 *   if (dragger.moved()) return;         // 单击删除处理前调用：拖动已应用则压掉本次 click
 *
 * cfg = {
 *   body:      function () → HTMLElement  tbody 元素（惰性取：抽屉 DOM 可能重建）
 *   getRow:    function (i) → row|null    行下标 → 行对象（无该行返回 null）
 *   excluded:  Set                        「已删除推算载重」的行对象集合（调用方自有）
 *   estLoadOf: function (row) → number     推算载重值（恢复显示用）
 *   isEditing: function () → boolean       是否处于计重编辑态（非编辑态本交互不启动）
 *   onChanged: function ()                 批量应用后回调（刷新标题合计）
 * }
 * 加载顺序：须在 app.js / sim-panel.js 之前。
 * ============================================================================
 */
(function (global) {
  'use strict';

  function create(cfg) {
    var st = { active: false, moved: false, anchor: -1, last: -1, snap: null, mode: 'exclude' };

    /** 把一行的推算载重格设为 删除(exc=true) / 恢复(exc=false)，同步该格 DOM；
     *  无可编辑推算格（实载行等）跳过。返回是否发生变更。 */
    function setExcRow(tr, row, exc) {
      var td = tr.querySelector('td.derived-est, td.excluded-cell');
      if (!td) return false;
      var isExc = cfg.excluded.has(row);
      if (exc === isExc) return false;                 // 已是目标态
      if (exc) {
        cfg.excluded.add(row);
        td.textContent = '';
        td.classList.remove('derived-est');
        td.classList.add('excluded-cell');
      } else {
        cfg.excluded.delete(row);
        td.textContent = String(Math.round(cfg.estLoadOf(row)));   // 恢复推算值（与渲染口径一致）
        td.classList.remove('excluded-cell');
        td.classList.add('derived-est');
      }
      return true;
    }

    /** 双侧范围应用：锚点..cur 按 mode，范围外按快照；回锚点 → 锚点..last 按快照恢复 */
    function apply(cur) {
      var body = cfg.body();
      if (!body) return;
      if (cur === st.anchor) {
        // 指针回到锚点：锚点..上次指针行 整段按快照恢复——只恢复锚点一格会漏掉
        // 回滑最后一步跨过的行（快速滑动跳行的中间行也一并覆盖）
        var lo = Math.min(st.anchor, st.last), hi = Math.max(st.anchor, st.last);
        var trs0 = body.querySelectorAll('tr');
        for (var i0 = 0; i0 < trs0.length; i0++) {
          var idx0 = +trs0[i0].getAttribute('data-i');
          if (idx0 < lo || idx0 > hi) continue;
          var row0 = cfg.getRow(idx0);
          if (row0) setExcRow(trs0[i0], row0, st.snap.has(row0));
        }
        cfg.onChanged();
        return;
      }
      var a = Math.min(st.anchor, cur), b = Math.max(st.anchor, cur);
      var trs = body.querySelectorAll('tr');
      var changed = false;
      for (var i = 0; i < trs.length; i++) {
        var idx = +trs[i].getAttribute('data-i');
        var row = cfg.getRow(idx);
        if (!row) continue;
        var exc = (idx >= a && idx <= b) ? (st.mode === 'exclude') : st.snap.has(row);
        if (setExcRow(trs[i], row, exc)) changed = true;
      }
      if (changed) cfg.onChanged();
    }

    /** 在 tbody 上挂交互；tbody 内容可被重建（监听挂在 tbody 元素本身，不受影响） */
    function bind() {
      var body = cfg.body();
      if (!body) return;
      body.addEventListener('mousedown', function (e) {
        if (!cfg.isEditing() || e.button !== 0) return;
        var det = e.target.closest && e.target.closest('td.derived-est, td.excluded-cell');
        if (!det) return;
        e.preventDefault();                    // 防止拖动选中单元格文字
        var tr = det.closest('tr');
        var idx = tr ? +tr.getAttribute('data-i') : -1;
        var row = cfg.getRow(idx);
        if (!row) return;
        st.active = true;
        st.moved = false;
        st.anchor = idx;
        st.last = idx;
        st.snap = new Set(cfg.excluded);       // 快照：范围外按快照复原
        st.mode = cfg.excluded.has(row) ? 'restore' : 'exclude';   // 起点状态定模式
      });
      body.addEventListener('mouseover', function (e) {
        if (!st.active) return;
        var tr = e.target.closest('tr');
        if (!tr) return;
        var i = +tr.getAttribute('data-i');
        if (isNaN(i) || i === st.last) return; // 同行去重
        st.moved = true;
        apply(i);
        st.last = i;
      });
      document.addEventListener('mouseup', function () {
        if (!st.active) return;
        st.active = false;                     // moved 保留：松手后的 click 由 moved() 压制
        st.anchor = -1;
        st.last = -1;
        st.snap = null;
      });
    }

    return {
      bind: bind,
      /** 拖动已批量应用时读并清除标志（单击切换处理前调用，防起点格被反向切换） */
      moved: function () { var m = st.moved; st.moved = false; return m; }
    };
  }

  global.EstDrag = { create: create };
})(window);
