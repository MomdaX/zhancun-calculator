/**
 * rpt31814-calc.js —— 31814 十八点站存统计【纯计算层】
 * ============================================================================
 * 从 report31814.js 拆出的**无 DOM 依赖**部分：
 *   · v                 数值取值（复用 Utils.vbVal）
 *   · getDirMap()       方向映射（复用 Aggregate 惰性单例）
 *   · isOpenTopBox(row) 敞顶箱判定
 *   · setCarProperties() 车流属性判定（VBA「设置车流属性.bas」）
 *
 * 【为什么拆】
 * 这段是 31814 的核心业务规则（长 If/ElseIf 链），判错是**静默**的——
 * 不抛异常、不崩溃，只会让站存数字悄悄变错，而本工具正是对账用的。
 * 独立成纯函数模块后，可被 HTML/tests/run.js（纯 Node、无 DOM）直接加载
 * 并回归测试，补上此前 31814 零测试覆盖的盲区。
 *
 * 依赖：window.Aggregate（COL / getDirectionIndex）、window.Utils（vbVal / determineCarType）
 * 暴露：window.Rpt31814Calc
 * ============================================================================
 */
(function (global) {
  'use strict';

  var COL = global.Aggregate ? global.Aggregate.COL : null;
  var Utils = global.Utils || {};

  /**
   * 数值取值：统一复用 Utils.vbVal，与主表聚合引擎共用同一套规则。
   *
   * 原实现为 parseFloat(trim)。实测两者在常规数据上等价
   * （parseFloat 同样会取前导数字，"12.5吨"→12.5、"38 吨"→38、全角空格亦被 trim 吃掉），
   * 替换的主要收益是消除"两套并存"这一隐患本身。
   * 唯一实质差异：旧实现遇 "Infinity"/"-Infinity" 会返回 Infinity 并污染合计，
   * vbVal 返回 0。故此处更稳，且行为变化仅限该极端输入。
   */
  var v = Utils.vbVal;

  /** 方向映射：复用 Aggregate 的惰性单例，全局只解析一次 CSV */
  function getDirMap() {
    if (!global.Aggregate) return {};
    return global.Aggregate.getDirectionIndex().map;
  }

  function isOpenTopBox(row) {
    var note = String(row[COL.NOTE] || '').toUpperCase();
    var train = String(row[COL.TRAIN] || '').toUpperCase();
    if (/敞顶箱|敞车箱/.test(note)) return true;
    if (/\b(86776|86774|49977|20328|34104)\d*\b/.test(note + train)) return true;
    return false;
  }

  /* ========================== 核心：设置车流属性 ==========================
   * 严格对齐 VBA「设置车流属性.bas」的 If / ElseIf 链：
   *
   *   ① 到站="防城港" 且 载重>4 且「品名」含"空"        → d(到站)
   *   ② 到站∈方向库 且 载重>25                          → 记事含[扣修]为空车，否则 d(到站)
   *   ③ 到站∈方向库 且 载重<10 且「品名」="自备" 且 车种=G → 自备
   *   ④ 到站∉方向库 且 载重>25                          → 方向代号 3=南口 / 2=管内 / 其余=沙口
   *   ⑤ 其余                                            → 空车
   *
   * 易错点（原实现的偏差来源）：
   *   · ① ③ 比较的是「品名」arr(i,10)，不是车号 arr(i,4)
   *   · ② VBA 的 Like "*[扣修]*" 是「字符类」，等价于“含『扣』或含『修』”，
   *     不是连续子串 "扣修"
   *   · ④ 必须带「载重>25」，否则到站为空的停留车会被误判成沙口/管内/南口
   *
   * @param {Array}  rawRows  原始数据行（state.rawRows）
   * @param {Object} drrSet   待装股道集合 { trackId: true }
   * @param {Object} crrSet   待发股道集合 { trackId: true }
   * @param {Object} [dirMap] 方向映射覆盖（仅测试注入用；不传则走 getDirMap()）
   * @returns {Array} 每行 → { track, carType, load, dest, dirCode, note, carNo, status, isOpen }
   * ================================================================== */
  function setCarProperties(rawRows, drrSet, crrSet, dirMap) {
    var dirMap2 = dirMap || getDirMap();
    var brr = [];

    rawRows.forEach(function (r) {
      var track = String(r[COL.TRACK] || '').trim();
      if (!track) return;

      var carTypeRaw = r[COL.CARTYPE];
      var carNo = r[COL.CARNO];
      var load = v(r[COL.LOAD]);
      var dest = String(r[COL.DEST] || '').trim();
      var dirCode = String(r[COL.DIR] || '').trim();
      var goods = String(r[COL.GOODS] == null ? '' : r[COL.GOODS]).trim();  // VBA arr(i,10) 品名
      var note = String(r[COL.NOTE] || '');
      var carType = Utils.determineCarType(carTypeRaw, carNo);

      // 等价于 VBA 的 d.Exists(到站)
      var inDir = dest !== '' && Object.prototype.hasOwnProperty.call(dirMap2, dest);

      var status;
      if (dest === '防城港' && load > 4 && goods.indexOf('空') >= 0) {
        // ① 防城港排空箱/重车：品名含"空"
        status = dirMap2[dest];
      } else if (inDir && load > 25) {
        // ② 到站交口（VBA Like "*[扣修]*" 为字符类：含"扣"或含"修"）
        status = /[扣修]/.test(note) ? '空车' : dirMap2[dest];
      } else if (inDir && load < 10 && goods === '自备' && carType === 'G') {
        // ③ 自备罐（品名为"自备"，且车种为 G）
        status = '自备';
      } else if (!inDir && load > 25) {
        // ④ 站名不在方向库时按方向代号识别
        if (dirCode === '3') status = '南口';
        else if (dirCode === '2') status = '管内';
        else status = '沙口';
      } else {
        // ⑤ 其余一律空车
        status = '空车';
      }

      // 罐车记事含「汽油/航煤」保持原判；敞顶箱统一归为待发候选
      var isOpen = isOpenTopBox(r);

      var row = {
        track: track, carType: carType, load: load, dest: dest,
        dirCode: dirCode, note: note, carNo: String(carNo || ''),
        status: status, isOpen: isOpen
      };

      // 待装覆盖
      if (drrSet[track] && load < 25) {
        if (carType === 'G' && row.carNo.charAt(0) === '0') row.status = '待装自备罐';
        else row.status = '待装';
      }

      // 待发覆盖
      if (crrSet[track]) {
        row.status = '待发';
      }

      brr.push(row);
    });

    return brr;
  }

  global.Rpt31814Calc = {
    v: v,
    getDirMap: getDirMap,
    isOpenTopBox: isOpenTopBox,
    setCarProperties: setCarProperties
  };
})(window);
