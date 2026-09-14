/**
 * grid-layout.js —— 主表布局计算（纯函数，零 DOM）
 * ============================================================================
 * 只与「股道配置 + 可见行」有关，与 DOM / state 无关，故从 app.js 抽出：
 *   groupSpans(vis)   分组「合并列」的跨行信息（谁首行、跨几行）
 *   bannerSlots(vis)  各作业区横幅该插在第几行
 * 抽出的意义：这两处一旦算错，表现是"分组名/横幅错位"这种看不出原因的显示问题，
 * 且原来只藏在一个 3000 行的文件里、无法单测。抽出后可被 tests/run.js 直接覆盖。
 *
 * 依赖：window.YardConfig（track.config.js）——提供 getTrack / mainAreas。
 * 加载顺序：须在 track.config.js 之后、app.js 之前。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var YardConfig = global.YardConfig;

  /**
   * 计算分组「合并列」所需的跨行信息。
   * @param {Array} vis visibleRows() 的结果（[{ r, idx }]）
   * @returns {Array<{group:string, start:boolean, span:number}>} 与 vis 等长
   *   group 该行的分组名（来自 track.config 的 groupName）
   *   start 是否为该分组在当前可见行中的首行（只有首行才输出带 rowspan 的单元格）
   *   span  该分组连续占据的行数（用于 rowspan）
   * 注意：隐藏虚拟股道后，虚拟场分组可能整段消失，span 仅统计可见部分。
   */
  function groupSpans(vis) {
    var arr = vis.map(function (item) {
      var cfg = YardConfig.getTrack(item.r.track);
      return { group: cfg ? cfg.groupName : '', start: false, span: 1 };
    });
    for (var i = 0; i < arr.length; i++) {
      if (i === 0 || arr[i].group !== arr[i - 1].group) {
        arr[i].start = true;
        var j = i + 1;
        while (j < arr.length && arr[j].group === arr[i].group) j++;
        arr[i].span = j - i;
      }
    }
    return arr;
  }

  /**
   * 计算各作业区横幅的插入位置。
   * 主表行顺序取自数据中股道的出现顺序（aggregate 按数据分组），未必等于配置顺序，
   * 故按股道在配置中的序位区间判定：每个作业区取其区间内「第一个出现的可见行」。
   * 该区股道被整段隐藏（空线分组开关）或数据中不存在时，不会产生横幅。
   * @param {Array} vis visibleRows() 的结果
   * @returns {Object} 行号 → 作业区配置（{ name, color, from, to, ids }）
   */
  function bannerSlots(vis) {
    var slots = {};
    var areas = (YardConfig && YardConfig.mainAreas) || [];
    areas.forEach(function (a) {
      for (var n = 0; n < vis.length; n++) {
        var t = YardConfig.getTrack(vis[n].r.track);
        if (t && t.index >= a.from && t.index <= a.to) {
          if (!slots[n]) slots[n] = a;   // 区间互不重叠，理论上不会冲突
          break;
        }
      }
    });
    return slots;
  }

  global.GridLayout = { groupSpans: groupSpans, bannerSlots: bannerSlots };

})(window);
