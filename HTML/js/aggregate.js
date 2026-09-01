/**
 * aggregate.js —— 股道存车聚合引擎
 * 逐条翻译自 VBA 模块「统计股道存车」的 Sub 股道存车(xrr)
 *
 * 【列索引说明（重要）】
 * SMIS 导出的 xls 中，表头行比数据行多一列（表头含"非"列，数据无此列），
 * 因此不能按表头定位，必须使用固定列索引。
 *
 *   0-based   VBA(1-based)   含义
 *   ----------------------------------------
 *   [0]       xrr(i, 1)      股道
 *   [1]       xrr(i, 2)      顺
 *   [2]       xrr(i, 3)      车种
 *   [3]       xrr(i, 4)      车号
 *   [4]       xrr(i, 5)      自重
 *   [5]       xrr(i, 6)      换长
 *   [6]       xrr(i, 7)      载重
 *   [7]       xrr(i, 8)      到站
 *   [8]       xrr(i, 9)      方向
 *   [9]       xrr(i,10)      品名
 *   [10]      xrr(i,11)      发站
 *   [13]      xrr(i,14)      记事
 *   [14]      xrr(i,15)      车次
 *   [15]      xrr(i,16)      到达时间
 */
(function (global) {
  'use strict';

  var COL = {
    TRACK: 0, SEQ: 1, CARTYPE: 2, CARNO: 3, TARE: 4, LEN: 5,
    LOAD: 6, DEST: 7, DIR: 8, GOODS: 9, FROM: 10,
    NOTE: 13, TRAIN: 14, ARRTIME: 15
  };

  /* ==========================================================================
   * VBA 函数模拟
   * ========================================================================== */

  /** VBA InStr：找不到返回 0，找到返回 1 基位置；查找空串返回 1 */
  function vbInStr(haystack, needle) {
    haystack = haystack == null ? '' : String(haystack);
    needle = needle == null ? '' : String(needle);
    if (needle === '') return 1;
    var i = haystack.indexOf(needle);
    return i === -1 ? 0 : i + 1;
  }

  /** VBA Val：从字符串开头解析数字，失败返回 0 */
  function vbVal(v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    var s = String(v).trim().replace(/^[\s　]+/, '');
    var m = /^[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?/.exec(s);
    if (!m) return 0;
    var n = parseFloat(m[0]);
    return isNaN(n) ? 0 : n;
  }

  /** VBA Left(str, n) */
  function vbLeft(s, n) {
    return String(s == null ? '' : s).substring(0, n);
  }

  /** VBA Mid(str, start, len)：start 为 1 基 */
  function vbMid(s, start, len) {
    s = String(s == null ? '' : s);
    var a = Math.max(0, start - 1);
    return len === undefined ? s.substring(a) : s.substring(a, a + len);
  }

  /**
   * 正则_车种：从车种字符串提取车型字母
   * VBA 原逻辑：取 [^\d]+ 部分，再按 C/X/P/G/YW/T/B/D/K 顺序匹配
   */
  function extractCarType(carType) {
    var s = String(carType == null ? '' : carType);
    var m = /[^\d]+/.exec(s);
    var t = m ? m[0] : '';
    var order = ['C', 'X', 'P', 'G', 'YW', 'T', 'B', 'D', 'K'];
    var result = t;
    for (var i = 0; i < order.length; i++) {
      var ls = order[i];
      if (t.indexOf(ls) >= 0) {
        // NX 系列特判为 X
        result = (vbLeft(t, 1) === 'N' && vbMid(t, 2, 1) === 'X') ? 'X' : vbLeft(t, 1);
        break;
      }
      result = t;
    }
    return result;
  }

  /** 解析到达时间字符串 → Date，失败返回 null */
  function parseArriveTime(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    if (!s) return null;
    var m = /(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})[日T\s]*(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?/.exec(s);
    if (!m) return null;
    var d = new Date(
      +m[1], +m[2] - 1, +m[3],
      m[4] ? +m[4] : 0, m[5] ? +m[5] : 0, m[6] ? +m[6] : 0
    );
    return isNaN(d.getTime()) ? null : d;
  }

  /** 小时差：等价于 VBA DateDiff("h", from, to) */
  function hoursDiff(from, to) {
    if (!from) return 0;
    return (to.getTime() - from.getTime()) / 3600000;
  }

  /* ==========================================================================
   * 到站推断 —— 翻译自 VBA 开头的三段 If/ElseIf
   * ========================================================================== */
  function resolveDest(row, dirStations, now) {
    var load = vbVal(row[COL.LOAD]);
    var dest = String(row[COL.DEST] == null ? '' : row[COL.DEST]).trim();
    var note = String(row[COL.NOTE] == null ? '' : row[COL.NOTE]).trim();
    var carType = String(row[COL.CARTYPE] == null ? '' : row[COL.CARTYPE]);
    var carNo = String(row[COL.CARNO] == null ? '' : row[COL.CARNO]);

    function tankJudge() {
      if (vbLeft(carType, 1) === 'G' && vbLeft(carNo, 1) === '6') return '路罐';
      if (vbLeft(carType, 1) === 'G' && vbLeft(carNo, 1) === '0') {
        return vbMid(carNo, 2, 1) === '7' ? '黑罐' : '自备罐';
      }
      return extractCarType(carType);
    }

    // 段1：载重<15 且 (到站为空 或 含"钦州港") 且 记事非空
    if (load < 15 && (dest === '' || dest.indexOf('钦州港') >= 0) && note !== '') {
      /* 先在记事中收集所有命中的站名，再取「最长」的那个。
       * 不能按方向库顺序取第一个命中：库里短名常排在长名之前
       * （如 防城@1 先于 防城港@4、贵阳 先于 贵阳南、昭通 先于 昭通北），
       * 取第一个会把「防城港」截成「防城」。
       * 等长时保留原顺序（先出现者优先），与旧行为一致。 */
      var best = '';
      for (var k = 0; k < dirStations.length; k++) {
        var s = dirStations[k];
        if (!s) continue;
        if (vbInStr(note, s) > 0 && s.length > best.length) best = s;
      }

      if (best) {                                  // 优先到站
        if (vbInStr(note, best + '循环') === 0 &&
            vbInStr(note, '卸空后返回' + best) === 0) {
          dest = best;
        } else {
          dest = tankJudge();
        }
        return dest;
      }

      // 未命中任何站名 → 其次品名（按车种 / 车号判定）
      if (vbLeft(carType, 1) === 'G' && vbLeft(carNo, 1) === '6') {
        if (vbInStr(note, '原装') === 0) {
          if (vbInStr(note, '汽油') > 0) dest = '汽油';
          else if (vbInStr(note, '柴油') > 0) dest = '柴油';
          else dest = '路罐';
        } else {
          dest = '路罐';
        }
      } else if (vbLeft(carType, 1) === 'G' && vbLeft(carNo, 1) === '0') {
        if (vbMid(carNo, 2, 1) === '7') {
          dest = '黑罐';
        } else if (vbInStr(note, '原装') === 0) {
          if (vbInStr(note, '汽油') > 0) dest = '汽油';
          else if (vbInStr(note, '柴油') > 0) dest = '柴油';
          else dest = '自备罐';
        } else {
          dest = '自备罐';
        }
      } else {
        dest = extractCarType(carType);
      }
      return dest;
    }

    // 段2：载重>15 且 到站为"钦州港" → 到卸
    if (load > 15 && dest === '钦州港') return '到卸';

    // 段3：载重、到站、记事均为空
    if (row[COL.LOAD] === '' || row[COL.LOAD] === null || row[COL.LOAD] === undefined) {
      if (dest === '' && note === '') return tankJudge();
    }

    return dest;
  }

  /* ==========================================================================
   * 主聚合函数
   * @param {Array<Array>} rows         原始数据行（不含表头）
   * @param {Object}       directionMap 到站→方向 映射
   * @param {Array<string>} dirStations 方向库全部站名（按顺序，用于到站推断）
   * @param {Object}       cfg          YardConfig.thresholds
   * @param {Date}         now          当前时间（便于测试）
   * ========================================================================== */
  function aggregate(rows, directionMap, dirStations, cfg, now) {
    now = now || new Date();
    cfg = cfg || {};
    var oldCarHours = cfg.oldCarHours != null ? cfg.oldCarHours : 47;
    directionMap = directionMap || {};
    dirStations = dirStations || [];

    // ---- 预处理：修正到站（等价于 VBA 循环体内的三段判断） ----
    var pre = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      var track = String(r[COL.TRACK] == null ? '' : r[COL.TRACK]).trim();
      if (!track) continue;

      var row = r.slice();
      // __destRaw：原始到站（车站名）。明细抽屉如实展示，不参与聚合改写
      row.__destRaw = String(r[COL.DEST] == null ? '' : r[COL.DEST]);

      // 聚合用的到站在副本上计算，避免污染原始值
      var work = r.slice();
      var loadRaw = r[COL.LOAD];
      var destRaw = r[COL.DEST];
      // VBA：载重为空但到站非空 → 清空到站
      if ((loadRaw === '' || loadRaw === null || loadRaw === undefined) &&
          !(destRaw === '' || destRaw === null || destRaw === undefined)) {
        work[COL.DEST] = '';
      }

      // __dest：聚合后的到站分类（可能是车站名，也可能是 路罐/自备罐/黑罐/车种）
      row.__dest = resolveDest(work, dirStations, now);
      row.__carType = extractCarType(r[COL.CARTYPE]);
      row.__track = track;
      pre.push(row);
    }

    // ---- 按股道分组（VBA 依赖数据连续排序，此处改为显式分组，更健壮且结果等价）----
    var groups = {}, order = [];
    for (var j = 0; j < pre.length; j++) {
      var t = pre[j].__track;
      if (!groups[t]) { groups[t] = []; order.push(t); }
      groups[t].push(pre[j]);
    }

    // ---- 逐组聚合 ----
    var result = {};
    for (var g = 0; g < order.length; g++) {
      var trackId = order[g];
      var list = groups[trackId];

      var count = 0, totalLen = 0, totalLoad = 0, oldCar = 0;
      var dirSet = [], destMap = {}, typeMap = {}, trainMap = [], noteWordSet = [];

      // 注意事项「简要提示」关键词字符类（对齐 VBA 显示信息.bas 点后开/重点事项筛选）
      // 含义：暂 不 走 去 向 点 后 开 列 扣 检 无 计 划 坏 超 偏 未 脏 禁 止 有 洗 排 磨 损
      var NOTE_KEY_RE = /[暂不走去向点后开列扣检无计划坏超偏未脏禁止有洗排磨损]/;
      var NOTE_EXCLUDE = '不入扣';

      for (var m = 0; m < list.length; m++) {
        var row2 = list[m];
        count += 1;
        totalLen += vbVal(row2[COL.LEN]);
        totalLoad += vbVal(row2[COL.LOAD]) + vbVal(row2[COL.TARE]);

        // 方向分类
        var dirCode = String(row2[COL.DIR] == null ? '' : row2[COL.DIR]).trim();
        var loadV = vbVal(row2[COL.LOAD]);
        function hasDir(d) {
          for (var x = 0; x < dirSet.length; x++) if (dirSet[x] === d) return true;
          return false;
        }
        if (dirCode === '3' && !hasDir('南口')) dirSet.push('南口');
        else if (dirCode === '2' && !hasDir('管内')) dirSet.push('管内');
        else if (dirCode === '6' && loadV > 20 && !hasDir('到卸')) dirSet.push('到卸');
        else if (vbInStr('236', dirCode) === 0 && !hasDir('沙口')) dirSet.push('沙口');

        // 到站统计（用聚合后的分类值 __dest，原始值 __destRaw 留给明细展示）
        var dst = String(row2.__dest == null ? '' : row2.__dest).trim() || '空车';
        destMap[dst] = (destMap[dst] || 0) + 1;

        // 车种统计
        var ct = row2.__carType || '空车';
        typeMap[ct] = (typeMap[ct] || 0) + 1;

        // 老牌车：停时 > 47h 且车号首位不为 0
        var arrTime = parseArriveTime(row2[COL.ARRTIME]);
        var carNoStr = String(row2[COL.CARNO] == null ? '' : row2[COL.CARNO]);
        if (hoursDiff(arrTime, now) > oldCarHours && vbLeft(carNoStr, 1) !== '0') {
          oldCar += 1;
        }

        // 车次统计
        var train = String(row2[COL.TRAIN] == null ? '' : row2[COL.TRAIN]).trim();
        if (train) trainMap[train] = (trainMap[train] || 0) + 1;

        // 记事（注意事项）简要提示：仅保留含关键词的「词」，与 VBA 显示信息.bas 点后开/重点事项筛选一致
        // VBA：dh(股道) = va & Chr(10) & dh(股道) —— 新词插到字符串「前面」
        // JS 中用 unshift 实现「后处理的词排在最上面」与 VBA 完全一致
        var noteTxt = String(row2[COL.NOTE] == null ? '' : row2[COL.NOTE]).trim();
        if (noteTxt) {
          // 分号、中文分号、空格均作为词分隔符
          var words = noteTxt.split(/[;；\s]+/);
          for (var wi = 0; wi < words.length; wi++) {
            var w = words[wi].trim();
            if (!w) continue;
            if (NOTE_EXCLUDE && vbInStr(NOTE_EXCLUDE, w) > 0) continue;   // 含「不入扣」则跳过
            if (NOTE_KEY_RE.test(w) && noteWordSet.indexOf(w) < 0) noteWordSet.unshift(w);
          }
        }
      }

      // 到站串： " 德保44 到卸23"
      var destStr = '';
      for (var dk in destMap) if (destMap.hasOwnProperty(dk)) destStr += ' ' + dk + destMap[dk];

      // 车种串： " C22 X22"
      var typeStr = '';
      for (var tk in typeMap) if (typeMap.hasOwnProperty(tk)) typeStr += ' ' + tk + typeMap[tk];

      // 车次：出现次数最多者
      var trainStr = '', maxV = 0;
      for (var nk in trainMap) {
        if (trainMap.hasOwnProperty(nk) && trainMap[nk] > maxV) { maxV = trainMap[nk]; trainStr = nk; }
      }

      // 方向串：换行分隔
      var dirStr = dirSet.join('\n');

      // 记事串：仅筛选后的关键词，换行分隔（与方向列一致），用于主表第一列「注意事项」
      var noteStr = noteWordSet.join('\n');

      result[trackId] = {
        track: trackId,
        direction: dirStr,
        count: count,
        carTypes: typeStr,
        length: Math.round(totalLen * 10) / 10,
        dest: destStr,
        train: trainStr,
        note: noteStr,
        load: totalLoad > 0 ? Math.round(totalLoad * 10) / 10 : 0,
        oldCar: oldCar > 0 ? oldCar : 0,
        raw: list
      };
    }

    return result;
  }

  /* ==========================================================================
   * CSV 解析（方向库）
   * ========================================================================== */
  function parseCSV(text) {
    if (!text) return [];
    text = text.replace(/^\uFEFF/, '');   // 去 BOM
    var rows = [], row = [], cur = '', inQ = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (inQ) {
        if (c === '"') {
          if (text[i + 1] === '"') { cur += '"'; i++; }
          else inQ = false;
        } else cur += c;
      } else {
        if (c === '"') inQ = true;
        else if (c === ',') { row.push(cur); cur = ''; }
        else if (c === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
        else if (c === '\r') { /* skip */ }
        else cur += c;
      }
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows.filter(function (r) {
      return r.some(function (cell) { return String(cell).trim() !== ''; });
    });
  }

  /** CSV → { 方向映射, 站名列表 } */
  function buildDirectionIndex(csvText) {
    var rows = parseCSV(csvText);
    if (!rows.length) return { map: {}, stations: [] };
    var header = rows[0].map(function (s) { return String(s).trim().toLowerCase(); });
    var iStation = header.indexOf('station');
    var iDir = header.indexOf('direction');
    if (iStation < 0) iStation = 0;
    if (iDir < 0) iDir = 1;

    var map = {}, stations = [];
    for (var i = 1; i < rows.length; i++) {
      var st = String(rows[i][iStation] == null ? '' : rows[i][iStation]).trim();
      var dr = String(rows[i][iDir] == null ? '' : rows[i][iDir]).trim();
      if (!st) continue;
      map[st] = dr;
      stations.push(st);
    }
    return { map: map, stations: stations };
  }

  global.Aggregate = {
    COL: COL,
    aggregate: aggregate,
    parseCSV: parseCSV,
    buildDirectionIndex: buildDirectionIndex,
    extractCarType: extractCarType,
    // 供调试/测试
    _vbInStr: vbInStr, _vbVal: vbVal, _vbLeft: vbLeft, _vbMid: vbMid,
    _resolveDest: resolveDest, _hoursDiff: hoursDiff
  };
})(window);
