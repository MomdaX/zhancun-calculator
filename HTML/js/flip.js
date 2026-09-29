/**
 * flip.js —— FLIP 位移动画小工具（拖拽「真实换序」后的平滑位移）
 * ============================================================================
 * FLIP = First / Last / Invert / Play：
 *   换序前 capture() 记录各元素位置（First）→ 真实改 DOM 顺序（Last）
 *   → play() 先无过渡地反向位移回旧位置（Invert）→ 强制重排 → 打开过渡位移到新位置（Play）
 *
 * 位置用 offsetLeft / offsetTop（布局位置）：
 *   · 不受容器滚动影响（拖到边缘滚动时不会算歪）；
 *   · 不受中途 transform 影响（动画未结束时再次调用也能算对）。
 *
 * 用法（表格行、列表项通用）：
 *   var fl = Flip.capture(container.children);   // 换序前
 *   ...insertBefore / appendChild 真实换序...
 *   fl.play();                                   // 换序后：从旧位置平滑滑到新位置
 *
 * 加载顺序：须在 flow-table.js 等使用方之前。
 * ============================================================================
 */
(function (global) {
  'use strict';

  function FlipItem(el) {
    this.el = el;
    this.x = el.offsetLeft;
    this.y = el.offsetTop;
    this.playing = false;     // 动画进行中：别重复「反转」，让它继续走向真实位置
  }

  /** 基准位置 = 当前布局位置 */
  FlipItem.prototype.record = function () {
    this.x = this.el.offsetLeft;
    this.y = this.el.offsetTop;
  };

  function Flip(items) { this.items = items; }

  /** 换序后播放：先把所有动过的元素反转回旧位置，再统一过渡到真实位置 */
  Flip.prototype.play = function () {
    var i, it, moved = [];
    for (i = 0; i < this.items.length; i++) {
      it = this.items[i];
      if (it.playing) continue;                                  // 动画未结束：跳过一次，避免抖动
      var dx = it.x - it.el.offsetLeft, dy = it.y - it.el.offsetTop;
      if (!dx && !dy) { it.record(); continue; }                 // 没动过
      it.el.style.transition = 'none';                           // Invert：关过渡，立刻回到旧位置
      it.el.style.transform = 'translate3d(' + dx + 'px,' + dy + 'px, 0)';   // 3D 位移：走合成层，更跟手不闪
      moved.push(it);
    }
    if (!moved.length) return;

    void document.body.offsetWidth;                              // 强制重排：让「反转」位置先生效

    for (i = 0; i < moved.length; i++) {
      it = moved[i];
      it.el.style.transition = '';                               // Play：交回样式表里的过渡
      it.el.style.transform = 'none';                            // 过渡到真实位置
      it.playing = true;
      it.record();
      (function (item) {
        item.el.addEventListener('transitionend', function onEnd(ev) {
          if (ev.target !== item.el || ev.propertyName !== 'transform') return;
          item.el.removeEventListener('transitionend', onEnd);
          item.playing = false;
          item.record();                                         // 动画结束：基准更新为真实位置
        });
      })(it);
    }
  };

  /** 记录一组元素的当前位置（First） */
  global.Flip = {
    capture: function (nodes) {
      var items = [];
      for (var i = 0; i < nodes.length; i++) items.push(new FlipItem(nodes[i]));
      return new Flip(items);
    }
  };
})(window);
