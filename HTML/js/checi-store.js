/**
 * checi-store.js —— 「编好车次」统一数据源（纯数据层，零 DOM）
 * ============================================================================
 * 三处共用同一份 Store.KEYS.readyTrains（{ 股道id: 车次 }）：
 *   ① 主表「编好车次」列（双击录入）
 *   ② 发车流程浮窗的「车次」框 depInputC4（点「编好」时预填）
 *   ③ 31814 报表「待发股道车次」芯片（report31814.js 读写同一个键）
 *
 * 为什么集中在这里：
 *   车次原本散在三处各写各的读写，任何一处改成别的键或别的结构，另外两处就悄悄不同步。
 *   统一成"只经过本模块"后，键名、去空格规则、空值即删除这三条口径只有一个出处。
 *   且一律走 Store（读写同一份内存缓存），不读单元格 DOM——避免"界面显示值"与
 *   "真实数据"两套真相。
 *
 * 加载顺序：须在 store.js 之后、app.js 之前（index.html 中置于 store.js 下一行）。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Store = global.Store;

  /** 整张 { 股道id: 车次 } 映射。取自 Store 的内存缓存，开销可忽略，故每次现取不缓存。 */
  function map() {
    return (Store && Store.get) ? (Store.get(Store.KEYS.readyTrains, {}) || {}) : {};
  }

  /** 取某股道的编好车次；未录入返回空串 */
  function of(trackId) {
    var m = map();
    return m[trackId] == null ? '' : String(m[trackId]);
  }

  /** 写入某股道的编好车次；传空字符串（或空白）即删除该项 */
  function set(trackId, val) {
    var m = map();
    val = String(val == null ? '' : val).trim();
    if (val) m[trackId] = val; else delete m[trackId];
    if (Store && Store.set) Store.set(Store.KEYS.readyTrains, m);
  }

  /**
   * 车次配色：套用「车辆信息」列同款方向色（沙口蓝 / 南口橙 / 管内紫）。
   * 一股道多方向时按 沙 > 南 > 管 取优先级，与 dest-color.js 的 renderDest 判定顺序一致；
   * 无方向（如"到卸"）返回空串 → 由 CSS 兜底为车次列的默认蓝。
   * @param {string} direction 股道聚合结果里的方向串（可能含换行分隔的多个方向）
   */
  function dirCls(direction) {
    var d = String(direction == null ? '' : direction);
    if (/沙/.test(d)) return 'shakou';
    if (/南/.test(d)) return 'nankou';
    if (/管内/.test(d)) return 'guanna';
    return '';
  }

  global.Checi = { map: map, of: of, set: set, dirCls: dirCls };

})(window);
