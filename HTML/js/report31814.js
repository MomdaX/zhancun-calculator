/**
 * report31814.js —— 31814 十八点站存统计
 *
 * 复刻 VBA 流程：
 *   1. 设置车流属性（按到站/方向/载重判定状态）
 *   2. 应用「待装股道」「待发股道」覆盖
 *   3. 计算.统计（输出到 Sheet1 的格式）
 *
 * 界面：左侧结果（与原 Sheet1 一致），右侧条件面板（与原 UserForm1 一致）
 *
 * 依赖：window.YardConfig, window.Aggregate, window.DirectionData
 * 调用：Report31814.open(rawRows, fileName)
 */
(function (global) {
  'use strict';

  var COL = global.Aggregate ? global.Aggregate.COL : null;

  // 默认待装区域来自 track.config.js（YardConfig.defaultAreas）。
  // 它是业务配置：现场新增/调整装卸线时只改配置，不必碰报表逻辑。
  var DEFAULT_AREAS = (global.YardConfig && global.YardConfig.defaultAreas) || {};

  /* ========================== 待装股道分组 ==========================
   * 待装股道仅取 DEFAULT_AREAS 中的装卸线，不再混入到发线等其他股道。
   * 每组标题可点击 = 整组全选 / 取消
   * ================================================================== */
  var DZ_GROUPS = (function buildDzGroups() {
    var groups = [];

    Object.keys(DEFAULT_AREAS).forEach(function (name) {
      var ids = DEFAULT_AREAS[name].filter(function (id) {
        return global.YardConfig && global.YardConfig.hasTrack(id);
      });
      if (ids.length) groups.push({ key: 'k' + groups.length, name: name, ids: ids, area: true });
    });

    return groups;
  })();

  var DZ_GROUP_MAP = {};
  DZ_GROUPS.forEach(function (g) { DZ_GROUP_MAP[g.key] = g; });

  // 大组（作业区）→ 子组 key 列表，按配置顺序组织
  var DZ_SECTIONS = (global.YardConfig && global.YardConfig.areaSections) || [];
  var DZ_SECTION_GROUPS = DZ_SECTIONS.map(function (sec) {
    return {
      name: sec.name,
      groups: sec.groups
        .map(function (gn) {
          var hit = null;
          DZ_GROUPS.forEach(function (g) { if (g.name === gn) hit = g.key; });
          return hit;
        })
        .filter(function (k) { return k != null; })
    };
  }).filter(function (s) { return s.groups.length; });

  /** 取股道显示名（如 "1道"、"H1"） */
  function trackLabel(id) {
    return global.YardConfig ? global.YardConfig.trackName(id) : String(id);
  }

  /** 刷新每组标题的选中态与「已选/总数」 */
  function syncGroupState() {
    DZ_GROUPS.forEach(function (g) {
      var n = 0;
      g.ids.forEach(function (id) {
        var cb = document.getElementById('cfgTrack_' + id);
        if (cb && cb.checked) n++;
      });
      var btn = document.querySelector('.cfg-group-btn[data-group="' + g.key + '"]');
      if (btn) btn.classList.toggle('active', n > 0 && n === g.ids.length);
      var cnt = document.querySelector('.cfg-group-count[data-count="' + g.key + '"]');
      if (cnt) cnt.textContent = n + '/' + g.ids.length;
    });
  }

  /** 整组切换（全选 / 取消） */
  function toggleGroup(key) {
    var g = DZ_GROUP_MAP[key];
    if (!g) return;
    var btn = document.querySelector('.cfg-group-btn[data-group="' + key + '"]');
    var on = !(btn && btn.classList.contains('active'));
    g.ids.forEach(function (id) {
      var cb = document.getElementById('cfgTrack_' + id);
      if (cb) cb.checked = on;
    });
    syncGroupState();
  }

  /* ========================== 工具函数 ========================== */

  /**
   * 数值取值：统一复用 Utils.vbVal，与主表聚合引擎共用同一套规则。
   *
   * 原实现为 parseFloat(trim)。实测两者在常规数据上等价
   * （parseFloat 同样会取前导数字，"12.5吨"→12.5、"38 吨"→38、全角空格亦被 trim 吃掉），
   * 本次替换主要收益是消除"两套并存"这一隐患本身。
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
   * ================================================================== */
  function setCarProperties(rawRows, drrSet, crrSet) {
    var dirMap = getDirMap();
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
      var inDir = dest !== '' && Object.prototype.hasOwnProperty.call(dirMap, dest);

      var status;
      if (dest === '防城港' && load > 4 && goods.indexOf('空') >= 0) {
        // ① 防城港排空箱/重车：品名含"空"
        status = dirMap[dest];
      } else if (inDir && load > 25) {
        // ② 到站交口（VBA Like "*[扣修]*" 为字符类：含"扣"或含"修"）
        status = /[扣修]/.test(note) ? '空车' : dirMap[dest];
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

  /* ========================== 核心：计算统计 ==========================
   * 性能优化：原实现对每个 type 都全表扫一遍 brr（9 类 ≈ 9 遍），
   * 管内分支内还对匹配站内再嵌套全表扫（最差 ~18 遍），末尾总数车种再一遍。
   * 现改为【单遍遍历】：一次 brr.forEach，把行按 status/dest/carType/track/isOpen/load
   * 直接累加进所有预建计算器，遍历次数从 ~11 遍降到 1 遍。
   * 语义严格等价（含首次出现顺序、倒序拼接、load>4 过滤等 VBA 细节）。 */
  function calculateStats(brr) {
    var dirMap = getDirMap();
    var types = ['沙口', '南口', '管内', '待卸', '待装', '空车', '待发', '待装自备罐', '自备'];
    var ls = {};
    var 车种顺序 = {};   // 车种按首次出现排序，与 VBA Dictionary 的插入顺序一致
    var gSplit = {};
    var dic = {};
    var openBox = {};  // 敞顶箱车号统计
    var gds = {};
    var totalCars = {};
    var 站内车种 = {};  // 管内：站内 → {carType: n}（load>4），避免嵌套全表扫
    var 站存 = 0;

    // 预建每个 type 的分桶容器（与 gSplit 分离，保持 gSplit 纯净 {road,self}）
    var buckets = {};
    types.forEach(function (type) {
      ls[type] = 0;
      ls[type + '车种'] = '';
      车种顺序[type] = [];
      gSplit[type] = { road: 0, self: 0 };
      buckets[type] = { dis: {}, dcz: {} };
    });

    // ---- 单遍遍历：所有累加在此完成 ----
    brr.forEach(function (row) {
      var type = row.status;
      if (types.indexOf(type) < 0) return;   // 不在九类内的行跳过（原逻辑等同）

      var bucket = buckets[type];
      var disKey = row.dest || '';
      bucket.dis[disKey] = (bucket.dis[disKey] || 0) + 1;
      bucket.dcz[row.carType] = (bucket.dcz[row.carType] || 0) + 1;

      // 罐车 G 分路/自备
      if (row.carType === 'G') {
        if (type === '待装自备罐' || type === '自备') gSplit[type].self += 1;
        else gSplit[type].road += 1;
      }

      // 沙口/南口/管内 的 dic（到站总数）
      if (type === '沙口' || type === '南口' || type === '管内') {
        dic[row.dest] = (dic[row.dest] || 0) + 1;
      }
      // 管内：load>4 的站内车种分桶（替代原嵌套全表扫）
      if (type === '管内' && row.load > 4) {
        var st0 = row.dest;
        if (!站内车种[st0]) 站内车种[st0] = {};
        站内车种[st0][row.carType] = (站内车种[st0][row.carType] || 0) + 1;
      }

      // 待发：股道 + 敞顶箱
      if (type === '待发') {
        gds[row.track] = (gds[row.track] || 0) + 1;
        if (row.isOpen) openBox[row.carNo] = (openBox[row.carNo] || 0) + 1;
      }

      // 总数车种（全量）
      totalCars[row.carType] = (totalCars[row.carType] || 0) + 1;
    });

    // ---- 收口：把分桶结果折算成原输出结构 ----
    types.forEach(function (type) {
      var dis = buckets[type].dis, dcz = buckets[type].dcz;
      ls[type] = Object.keys(dis).reduce(function (a, k) { return a + dis[k]; }, 0);
      // 保留首次出现顺序（VBA Dictionary 的 keys 顺序）：dcz 累加顺序即首次出现顺序
      车种顺序[type] = Object.keys(dcz).map(function (t) { return { t: t, n: dcz[t] }; });
      ls[type + '车种'] = 车种顺序[type].map(function (o) { return o.t + o.n; }).join(' ') + ' ';
      站存 += ls[type];
    });

    ls.gSplit = gSplit;

    // 沙口/南口/管内 字符串
    var 沙口 = '', 南口 = '', 管内 = '', 管内车种合计 = '';
    Object.keys(dic).forEach(function (st) {
      var dir = dirMap[st] || '';
      if (/沙/.test(dir)) 沙口 = st + dic[st] + ' ' + 沙口;
      else if (/南/.test(dir)) 南口 = st + dic[st] + ' ' + 南口;
      else if (/管内/.test(dir)) {
        var stationCars = 站内车种[st] || {};
        // VBA：单项合计 = gr & dcz(gr) & 单项合计 —— 倒序拼接
        var single = '';
        Object.keys(stationCars).forEach(function (t) { single = t + stationCars[t] + single; });
        管内 = st + ':' + dic[st] + ' ' + 管内;
        管内车种合计 = st + ':' + single + ' ' + 管内车种合计;
      }
    });

    // 待发股道合计：按股道顺序升序（数字道在前，X 道在后）
    // 若该股道设了车次，则把「股道名」替换为「车次」（如 2道/58 → 34102/58）
    var 待发股道合计 = Object.keys(gds).sort(Utils.compareTrackId)
      .map(function (tr) {
        var train = readyTrains[tr];
        var prefix = train ? String(train) : (tr + '道');
        return prefix + '/' + gds[tr];
      }).join(' ');

    // 敞顶箱列表：敞顶箱/总数 + 编号/数量
    var openTotal = Object.keys(openBox).reduce(function (a, k) { return a + openBox[k]; }, 0);
    var openList = Object.keys(openBox).map(function (k) { return k + '/' + openBox[k]; }).join(' ');

    // 总车种：VBA 为 总数车种合计 = ar & zs(ar) & " " & 总数车种合计（倒序拼接）
    var 总数车种顺序 = Object.keys(totalCars).map(function (t) { return { t: t, n: totalCars[t] }; });
    var 总数车种合计 = '';
    总数车种顺序.forEach(function (o) { 总数车种合计 = o.t + o.n + ' ' + 总数车种合计; });

    return {
      ls: ls,
      车种顺序: 车种顺序,
      沙口: 沙口.trim(),
      南口: 南口.trim(),
      管内: 管内.trim(),
      管内车种合计: 管内车种合计.trim(),
      待发股道合计: 待发股道合计.trim(),
      总数车种合计: 总数车种合计.trim(),
      总数车种顺序: 总数车种顺序,
      站存: 站存,
      openTotal: openTotal,
      openList: openList
    };
  }

  /* ========================== 结果渲染 ========================== */
  /**
   * 罐车 G 的分数显示：分子 = 自备罐，分母 = 路罐
   * 两者皆为 0 时不显示（对应 Sheet1 中 F 列留空的情形）
   */
  function renderFraction(self, road) {
    if ((self || 0) <= 0 && (road || 0) <= 0) return '';
    return '<span class="fraction"><span class="num">' + (self || 0) +
           '</span><span class="den">' + (road || 0) + '</span></span>';
  }

  /**
   * 拼接车种串，如 "C20 X3 P5"；传入 gRoad/gSelf 时 G 以分数形式附在末尾。
   * ordered 为按首次出现排序的 [{t:'C', n:20}, ...]，与 VBA Dictionary 的迭代顺序一致。
   * 分数中的 G 不再重复输出计数，避免出现 "G587 G 0/10" 这类重复。
   */
  function buildCarLine(ordered, gRoad, gSelf) {
    var gSelf0 = gSelf || 0, gRoad0 = gRoad || 0;
    var hasFrac = gSelf0 > 0 || gRoad0 > 0;
    var parts = [];
    (ordered || []).forEach(function (o) {
      if (o.t === 'G' && hasFrac) return;    // G 由末尾分数呈现
      if (o.n) parts.push(o.t + o.n);
    });
    if (hasFrac) parts.push('G ' + renderFraction(gSelf0, gRoad0));
    return parts.join(' ');
  }

  function renderResult(stats) {
    // 表格内的标题行仍保留原日期标题
    var title = '钦州港 ' + Utils.formatDate(new Date()) + ' 早6点 站存';

    var g = stats.ls.gSplit;
    var ord = stats.车种顺序 || {};

    var html = '<table>';
    html += '<tr><td colspan="3" class="title">' + Utils.escapeHtml(title) + '</td></tr>';

    function row(label, top, bottom, total, cls, ids) {
      var topHtml = top || '&nbsp;';
      // two-row：含 detail-bottom 的两行结构，上半格高度按 CSS 取下半格的 3 倍
      var clsAll = ((cls || '') + (bottom ? ' two-row' : '')).trim();
      var topId = ids && ids.top ? ' id="' + ids.top + '"' : '';
      var totalId = ids && ids.total ? ' id="' + ids.total + '"' : '';
      return '<tr class="' + clsAll + '">' +
        '<td class="label">' + Utils.escapeHtml(label) + '</td>' +
        '<td><div class="detail-top' + (bottom ? '' : ' only') + '"' + topId + '>' + topHtml + '</div>' +
        (bottom ? '<div class="detail-bottom">' + bottom + '</div>' : '') + '</td>' +
        '<td class="total"' + totalId + '>' + total + '</td></tr>';
    }

    // 沙口（B2 到站 / B3 车种，G 不拆分）
    html += row('沙口', Utils.escapeHtml(stats.沙口),
      buildCarLine(ord['沙口'], 0, 0),
      stats.ls['沙口'], 'dir-shakou');
    // 南口（B4 / B5）
    html += row('南口', Utils.escapeHtml(stats.南口),
      buildCarLine(ord['南口'], 0, 0),
      stats.ls['南口'], 'dir-nankou');
    // 管内（B6）：单格，管内车种合计已含站名与车种，如「防城港:X7C56」
    html += row('管内', Utils.escapeHtml(stats.管内车种合计), '',
      stats.ls['管内'], 'dir-guanna');

    // 待卸（B7）：单格，只有车种合计
    html += row('待卸',
      buildCarLine(ord['待卸'], 0, 0),
      '', stats.ls['待卸'], '', { top: 'dxTop', total: 'dxTotal' });

    // 待装（B8）：G 拆分为自备罐(分子) / 路罐(分母)，与车种同行显示
    html += row('待装',
      buildCarLine(ord['待装'], g['待装'].road, g['待装自备罐'].self),
      '',
      stats.ls['待装'] + stats.ls['待装自备罐'], '');

    // 空车（B12）：G 拆分为自备罐(分子) / 路罐(分母)
    // VBA：H12 = ls("空车") + 自备数，F12 = 自备数，F14 = 空车中的路罐数
    html += row('空车',
      buildCarLine(ord['空车'], g['空车'].road, g['自备'].self),
      '',
      stats.ls['空车'] + stats.ls['自备'], '', { top: 'kcTop', total: 'kcTotal' });

    // 待开（B16 股道合计 / B17 车种）
    html += row('待开',
      Utils.escapeHtml(stats.待发股道合计),
      buildCarLine(ord['待发'], 0, 0),
      stats.ls['待发'], '');

    // 站存（B18 总车种合计，G 不拆分；按 VBA zs 字典的插入顺序）
    html += row('站存',
      buildCarLine(stats.总数车种顺序, 0, 0),
      '', stats.站存, '');

    html += '</table>';
    Utils.$('rptBody').innerHTML = html;
    applyFontSize();   // innerHTML 重建后重新套用当前字号
    lastStats = stats; // 供「待卸纠正」依据自动值重算
    // 仅当用户点过「待卸校对」才套用纠正；否则保持自动值（避免打开时立即改值）
    if (corrApplied) applyCorrection();
  }

  /* ===================== 待卸纠正（对应 VBA 待卸纠正） =====================
   * 自动算出的待卸与实况有偏差时，在面板底部 C/X/P 输入框填入正确车数，
   * 点击「待卸校对」后：待卸取输入值；对应车种的偏差从空车中增减（多了放到空车，
   * 少了从空车取），使待卸精确等于输入值、空车总数相应找平。
   * ======================================================================= */
  var lastStats = null;
  var corrApplied = false;   // 仅当用户点击「待卸校对」后才真正套用纠正

  function carMapFromOrdered(ordered) {
    var m = {};
    (ordered || []).forEach(function (o) { m[o.t] = o.n; });
    return m;
  }
  function orderedFromCarMap(map, order) {
    var seen = {}, out = [];
    order.forEach(function (t) {
      if (t in map && !seen[t]) { out.push({ t: t, n: map[t] }); seen[t] = true; }
    });
    Object.keys(map).forEach(function (t) {
      if (!seen[t]) out.push({ t: t, n: map[t] });
    });
    return out;
  }
  function orderedTotal(ordered) {
    return (ordered || []).reduce(function (s, o) { return s + (o.n || 0); }, 0);
  }
  function readCorr(id, autoVal) {
    var el = document.getElementById(id);
    if (!el) return autoVal;
    var v = String(el.value).trim();
    if (v === '') return autoVal;
    var n = parseInt(v, 10);
    return isNaN(n) ? autoVal : n;
  }

  function applyCorrection() {
    if (!lastStats) return;
    var ord = lastStats.车种顺序 || {};
    var autoDx = carMapFromOrdered(ord['待卸']);

    var inpC = readCorr('corrC', autoDx.C || 0);
    var inpX = readCorr('corrX', autoDx.X || 0);
    var inpP = readCorr('corrP', autoDx.P || 0);
    var inpG = readCorr('corrG', autoDx.G || 0);

    // 四个输入都为空 → 不纠正，保持自动值
    if (Utils.$('corrC').value.trim() === '' &&
        Utils.$('corrX').value.trim() === '' &&
        Utils.$('corrP').value.trim() === '' &&
        Utils.$('corrG').value.trim() === '') return;

    var delta = {
      C: inpC - (autoDx.C || 0),
      X: inpX - (autoDx.X || 0),
      P: inpP - (autoDx.P || 0),
      G: inpG - (autoDx.G || 0)
    };
    var g = lastStats.ls.gSplit;
    var corrTypes = ['C', 'X', 'P', 'G'];

    // —— 待卸：在原顺序基础上替换 C/X/P/G 的值为输入值（位置不变）——
    var dxPresent = {}, dxOrdered = [];
    (ord['待卸'] || []).forEach(function (o) {
      dxPresent[o.t] = true;
      if (o.t === 'C') dxOrdered.push({ t: 'C', n: inpC });
      else if (o.t === 'X') dxOrdered.push({ t: 'X', n: inpX });
      else if (o.t === 'P') dxOrdered.push({ t: 'P', n: inpP });
      else if (o.t === 'G') dxOrdered.push({ t: 'G', n: inpG });
      else dxOrdered.push(o);
    });
    // 原顺序中没有、但用户输入了的车种，补到末尾
    corrTypes.forEach(function (t) {
      var el = document.getElementById('corr' + t);
      if (!dxPresent[t] && el && el.value.trim() !== '') {
        dxOrdered.push({ t: t, n: { C: inpC, X: inpX, P: inpP, G: inpG }[t] });
      }
    });
    var dxTop = Utils.$('dxTop');
    if (dxTop) dxTop.innerHTML = buildCarLine(dxOrdered, 0, 0);
    var dxTotal = Utils.$('dxTotal');
    if (dxTotal) dxTotal.textContent = orderedTotal(dxOrdered);

    // —— 空车：C/X/P 按相反偏差增减；G 对应分母(路罐)按相反偏差增减 ——
    var kcPresent = {}, kcOrdered = [];
    (ord['空车'] || []).forEach(function (o) {
      kcPresent[o.t] = true;
      if (o.t === 'C') kcOrdered.push({ t: 'C', n: Math.max(0, (o.n || 0) - delta.C) });
      else if (o.t === 'X') kcOrdered.push({ t: 'X', n: Math.max(0, (o.n || 0) - delta.X) });
      else if (o.t === 'P') kcOrdered.push({ t: 'P', n: Math.max(0, (o.n || 0) - delta.P) });
      else kcOrdered.push(o);
    });
    corrTypes.forEach(function (t) {
      if (t === 'G') return; // G 走分母，不计入分子
      var el = document.getElementById('corr' + t);
      if (!kcPresent[t] && el && el.value.trim() !== '') {
        kcOrdered.push({ t: t, n: Math.max(0, -(delta[t])) });
      }
    });
    var newRoad = Math.max(0, g['空车'].road - delta.G);
    var kcTop = Utils.$('kcTop');
    if (kcTop) kcTop.innerHTML = buildCarLine(kcOrdered, newRoad, g['自备'].self);
    var kcTotal = Utils.$('kcTotal');
    if (kcTotal) {
      var autoKcTotal = lastStats.ls['空车'] + lastStats.ls['自备'];
      kcTotal.textContent = autoKcTotal - (delta.C + delta.X + delta.P + delta.G);
    }
  }

  /* ========================== 结果区截图 ==========================
   * 采用 SVG <foreignObject> + Canvas：不依赖第三方库，可在 file:// 下工作。
   * 由于 file:// 下无法读取 <link> 样式表的 cssRules（SecurityError），
   * 改为用 getComputedStyle 把关键样式内联到克隆节点后再序列化。
   * ================================================================== */

  /** 序列化时需要内联的样式属性白名单 */
  var CAPTURE_PROPS = [
    'box-sizing', 'display', 'position', 'float', 'clear',
    'width', 'height', 'min-width', 'min-height', 'max-width', 'max-height',
    'margin-top', 'margin-right', 'margin-bottom', 'margin-left',
    'padding-top', 'padding-right', 'padding-bottom', 'padding-left',
    'border-top-width', 'border-right-width', 'border-bottom-width', 'border-left-width',
    'border-top-style', 'border-right-style', 'border-bottom-style', 'border-left-style',
    'border-top-color', 'border-right-color', 'border-bottom-color', 'border-left-color',
    'border-collapse', 'border-spacing', 'border-radius', 'outline',
    'background-color', 'background-image', 'color', 'opacity',
    'font-family', 'font-size', 'font-weight', 'font-style', 'font-variant-numeric',
    'line-height', 'letter-spacing', 'text-align', 'text-decoration',
    'vertical-align', 'white-space', 'word-break', 'overflow-wrap',
    'text-overflow', 'overflow', 'overflow-x', 'overflow-y',
    'table-layout', 'list-style',
    'flex', 'flex-direction', 'flex-wrap', 'align-items', 'justify-content', 'gap'
  ];

  /** 把源节点树上每个元素的计算样式内联到目标节点树 */
  function inlineStyles(src, dst) {
    var srcEls = [src].concat(Array.prototype.slice.call(src.querySelectorAll('*')));
    var dstEls = [dst].concat(Array.prototype.slice.call(dst.querySelectorAll('*')));
    for (var i = 0; i < srcEls.length && i < dstEls.length; i++) {
      var cs = getComputedStyle(srcEls[i]);
      var st = dstEls[i].style;
      for (var j = 0; j < CAPTURE_PROPS.length; j++) {
        var p = CAPTURE_PROPS[j];
        var val = cs.getPropertyValue(p);
        if (val) st.setProperty(p, val);
      }
      // sticky/fixed 在 foreignObject 中定位失效，重置为静态
      var pos = st.getPropertyValue('position');
      if (pos === 'sticky' || pos === 'fixed') st.setProperty('position', 'static');
    }
  }

  /* ========================== 结果表格字号缩放 ==========================
   * 只改 #rptBody 上的 --rpt-fs，表格内行高/标题/字号等由 CSS calc 联动。
   * ================================================================== */
  var FS_MIN = 8, FS_MAX = 22, FS_STEP = 1;
  var fsCurrent = 15;   // 与 CSS 中 --rpt-fs 默认值保持一致

  /** 把当前字号写入 #rptBody，并同步中间显示与两侧按钮的可用状态/提示 */
  function applyFontSize() {
    var body = Utils.$('rptBody');
    if (body) body.style.setProperty('--rpt-fs', fsCurrent + 'px');

    // 中间显示当前字号
    var val = Utils.$('rptFsVal');
    if (val) val.textContent = fsCurrent + 'px';

    var inBtn = Utils.$('rptZoomIn'), outBtn = Utils.$('rptZoomOut');
    if (inBtn) {
      inBtn.disabled = (fsCurrent >= FS_MAX);
      inBtn.title = '放大表格字体（当前 ' + fsCurrent + 'px）';
    }
    if (outBtn) {
      outBtn.disabled = (fsCurrent <= FS_MIN);
      outBtn.title = '缩小表格字体（当前 ' + fsCurrent + 'px）';
    }
  }

  /** delta > 0 放大，< 0 缩小 */
  function zoomFontSize(delta) {
    var next = Math.min(FS_MAX, Math.max(FS_MIN, fsCurrent + delta));
    if (next === fsCurrent) return;   // 已到边界
    fsCurrent = next;
    applyFontSize();
  }

  /** 将元素渲染为 PNG Blob（2 倍图，白底） */
  function captureElement(el) {
    return new Promise(function (resolve, reject) {
      if (!el) { reject(new Error('未找到内容区')); return; }

      var rect = el.getBoundingClientRect();
      var w = Math.ceil(Math.max(rect.width, el.scrollWidth));
      var h = Math.ceil(Math.max(rect.height, el.scrollHeight));
      if (!w || !h) { reject(new Error('内容区为空')); return; }

      var clone = el.cloneNode(true);
      inlineStyles(el, clone);
      // 截图取完整内容：解除滚动裁剪，按内容自然高度展开
      clone.style.setProperty('width', w + 'px');
      clone.style.setProperty('height', 'auto');
      clone.style.setProperty('min-height', '0');
      clone.style.setProperty('max-height', 'none');
      clone.style.setProperty('overflow', 'visible');
      clone.style.setProperty('margin', '0');

      // 外部资源在 svg-as-image 中加载不到，且会污染画布，先清理
      Array.prototype.forEach.call(clone.querySelectorAll('img,svg,canvas,iframe,video'), function (n) {
        n.parentNode.removeChild(n);
      });
      Array.prototype.forEach.call([clone].concat(Array.prototype.slice.call(clone.querySelectorAll('*'))), function (n) {
        var bi = n.style.getPropertyValue('background-image');
        if (bi && bi.indexOf('url(') >= 0) n.style.setProperty('background-image', 'none');
      });

      // XHTML 未预定义 &nbsp;，需转为数字实体，否则 XML 解析失败
      var xhtml = new XMLSerializer().serializeToString(clone)
        .replace(/&nbsp;/g, '&#160;');

      // 四周留一圈白边，避免表格最外框（border-collapse 居中描边）被裁掉
      var PAD = 4;
      var W = w + PAD * 2;
      var H = h + PAD * 2;

      var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="' + W + '" height="' + H + '">' +
        '<foreignObject x="0" y="0" width="100%" height="100%">' +
        '<div xmlns="http://www.w3.org/1999/xhtml" style="width:' + w + 'px;' +
        'padding:' + PAD + 'px;box-sizing:content-box;background:#fff">' +
        xhtml + '</div></foreignObject></svg>';

      // file:// 下 blob: 属于 null 源，画到 canvas 会污染画布导致无法导出，
      // 因此必须用 data: URL 承载 SVG
      var img = new Image();

      img.onload = function () {
        try {
          var scale = 2;   // 2 倍图，粘贴后更清晰
          var canvas = document.createElement('canvas');
          canvas.width = Math.ceil(W * scale);
          canvas.height = Math.ceil(H * scale);
          var ctx = canvas.getContext('2d');
          ctx.fillStyle = '#fff';
          ctx.fillRect(0, 0, canvas.width, canvas.height);
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
          canvas.toBlob(function (blob) {
            if (blob) resolve(blob);
            else reject(new Error('生成图片失败'));
          }, 'image/png');
        } catch (err) {
          reject(err);
        }
      };

      img.onerror = function () {
        reject(new Error('渲染失败'));
      };

      img.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
  }

  /** 剪切板不可用时降级为下载图片 */
  function downloadBlob(blob, name) {
    var a = document.createElement('a');
    var url = URL.createObjectURL(blob);
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  /** 截图按钮：把结果表格复制为图片（只取表格本体，不含外层留白） */
  function copyResultAsImage() {
    var body = Utils.$('rptBody');
    var target = (body && body.querySelector('table')) || body;
    if (!target) { Utils.toast('没有可截图的内容', 'error'); return; }

    captureElement(target).then(function (blob) {
      if (!navigator.clipboard || typeof window.ClipboardItem === 'undefined') {
        downloadBlob(blob, '站存计算.png');
        Utils.toast('当前环境不支持剪切板，已下载为图片', 'ok');
        return;
      }
      navigator.clipboard.write([new window.ClipboardItem({ 'image/png': blob })])
        .then(function () { Utils.toast('结果区已复制为图片，可直接粘贴', 'ok'); })
        .catch(function () {
          downloadBlob(blob, '站存计算.png');
          Utils.toast('复制受限，已保存为图片文件', 'error');
        });
    }).catch(function (e) {
      Utils.toast('截图失败：' + (e && e.message ? e.message : e), 'error');
    });
  }

  /* ========================== 右侧条件面板 ========================== */
  var currentRawRows = [];
  var config = { drr: {}, crr: {} };
  var readyTrains = {};   // 待发股道 → 车次 映射（如 {"2":"34102"}），用于替换待开中的股道名

  // 根据 readyTrains 更新某股道「车次芯片」的显隐与文案
  function updateReadyChip(item, id) {
    if (!item) return;
    var chip = item.querySelector('.cfg-ready-chip');
    var text = item.querySelector('.cfg-ready-chip-text');
    if (!chip) return;
    var val = readyTrains[id];
    if (val != null) { text.textContent = val; chip.style.display = 'inline-flex'; }
    else { chip.style.display = 'none'; }
  }

  function ensureConfigUI() {
    if (Utils.$('cfgPanel').innerHTML.trim()) return;

    var html = '';

    // 待装股道 与 待发股道 左右布局
    html += '<div class="cfg-main-row">';
    html += '<div class="cfg-tracks-zone">';
    html += '<div class="cfg-section-title">';
    html += '<span class="cfg-title-text">待装股道</span>';
    html += '<span class="cfg-toggle-row">';
    html += '<button class="btn" id="cfgDefaultDz">默认待装</button>';
    html += '<label class="cfg-all"><input type="checkbox" id="cfgAllDz"> 全选</label>';
    html += '</span>';
    html += '</div>';
    html += '<div class="cfg-tracks" id="cfgTracks"></div>';
    html += '</div>';
    html += '<div class="cfg-ready-zone">';
    html += '<div class="cfg-section-title">待发股道</div>';
    html += '<div class="cfg-ready-tracks" id="cfgReady"></div>';
    html += '</div>';
    html += '</div>';

    html += '<div class="cfg-btns">';
    html += '<button class="btn primary" id="btnStart31814">开始计算</button>';
    html += '<button class="btn" id="btnClear31814">清空数据区域</button>';
    html += '<button class="btn" id="btnCloseCfg31814">关闭</button>';
    html += '</div>';

    Utils.$('cfgPanel').innerHTML = html;

    // 待装股道：按「大组(作业区) / 子组(线别) / 股道」三层渲染
    var tb = Utils.$('cfgTracks');

    function buildGroup(g) {
      var box = document.createElement('div');
      box.className = 'cfg-track-group' + (g.area ? ' is-area' : '');

      var head = document.createElement('div');
      head.className = 'cfg-group-head';
      head.innerHTML = '<button type="button" class="cfg-group-btn" data-group="' + g.key + '">' +
        Utils.escapeHtml(g.name) + '</button>' +
        '<span class="cfg-group-count" data-count="' + g.key + '">0/' + g.ids.length + '</span>';
      box.appendChild(head);

      var wrap = document.createElement('div');
      wrap.className = 'cfg-group-tracks';
      g.ids.forEach(function (id) {
        var label = document.createElement('label');
        label.className = 'cfg-track';
        label.innerHTML = '<input type="checkbox" id="cfgTrack_' + id + '" data-track="' + id + '">' +
          Utils.escapeHtml(trackLabel(id));
        wrap.appendChild(label);
      });
      box.appendChild(wrap);
      return box;
    }

    // 仅渲染股道（不含子组标题/按钮），用于单组大组
    function buildTracksOnly(g) {
      var wrap = document.createElement('div');
      wrap.className = 'cfg-group-tracks';
      g.ids.forEach(function (id) {
        var label = document.createElement('label');
        label.className = 'cfg-track';
        label.innerHTML = '<input type="checkbox" id="cfgTrack_' + id + '" data-track="' + id + '">' +
          Utils.escapeHtml(trackLabel(id));
        wrap.appendChild(label);
      });
      return wrap;
    }

    // 配置中存在大组划分时，按「作业区」分段；否则退回扁平单层
    if (DZ_SECTION_GROUPS.length) {
      DZ_SECTION_GROUPS.forEach(function (sec) {
        var secEl = document.createElement('div');
        secEl.className = 'cfg-section';

        var secHead = document.createElement('div');
        secHead.className = 'cfg-section-head';
        secHead.textContent = sec.name;
        secHead.setAttribute('data-groups', sec.groups.join(','));

        var secBody = document.createElement('div');
        secBody.className = 'cfg-section-groups';

        if (sec.groups.length === 1) {
          // 仅一组：隐藏子组按钮，计数移到段标题后
          var g = DZ_GROUP_MAP[sec.groups[0]];
          if (g) {
            var cnt = document.createElement('span');
            cnt.className = 'cfg-group-count';
            cnt.setAttribute('data-count', g.key);
            cnt.textContent = '0/' + g.ids.length;
            secHead.appendChild(cnt);
            secBody.appendChild(buildTracksOnly(g));
          }
        } else {
          sec.groups.forEach(function (key) {
            var g = DZ_GROUP_MAP[key];
            if (g) secBody.appendChild(buildGroup(g));
          });
        }
        secEl.appendChild(secHead);
        secEl.appendChild(secBody);

        tb.appendChild(secEl);
      });

      // 点击段标题：整段全选 / 清空全选
      tb.querySelectorAll('.cfg-section-head').forEach(function (head) {
        head.addEventListener('click', function () {
          var keys = (head.getAttribute('data-groups') || '').split(',').filter(Boolean);
          var anyUnchecked = false;
          keys.forEach(function (k) {
            var g = DZ_GROUP_MAP[k];
            if (!g) return;
            g.ids.forEach(function (id) {
              var cb = document.getElementById('cfgTrack_' + id);
              if (cb && !cb.checked) anyUnchecked = true;
            });
          });
          var on = anyUnchecked;
          keys.forEach(function (k) {
            var g = DZ_GROUP_MAP[k];
            if (!g) return;
            g.ids.forEach(function (id) {
              var cb = document.getElementById('cfgTrack_' + id);
              if (cb) cb.checked = on;
            });
          });
          syncGroupState();
          runCalculation();
        });
      });
    } else {
      DZ_GROUPS.forEach(function (g) { tb.appendChild(buildGroup(g)); });
    }

    // 待卸纠正：置于最底部，与上方大组用分隔线隔开
    var corr = document.createElement('div');
    corr.className = 'cfg-correction';
    corr.innerHTML =
      '<div class="cfg-divider"></div>' +
      '<div class="cfg-correction-row">' +
        '<label class="cfg-corr-field"><span class="cfg-corr-letter">C</span><input type="number" id="corrC" min="0" step="1" placeholder="0"></label>' +
        '<label class="cfg-corr-field"><span class="cfg-corr-letter">X</span><input type="number" id="corrX" min="0" step="1" placeholder="0"></label>' +
        '<label class="cfg-corr-field"><span class="cfg-corr-letter">P</span><input type="number" id="corrP" min="0" step="1" placeholder="0"></label>' +
        '<label class="cfg-corr-field"><span class="cfg-corr-letter">G</span><input type="number" id="corrG" min="0" step="1" placeholder="0"></label>' +
        '<button type="button" class="btn" id="btnCorrect">待卸校对</button>' +
      '</div>';
    tb.appendChild(corr);

    Utils.$('btnCorrect').addEventListener('click', function () {
      syncGroupState();
      corrApplied = true;     // 只有点击校对后才真正套用纠正
      applyCorrection();
    });

    // 纠正输入框：失焦/回车时本地持久记忆
    ['corrC', 'corrX', 'corrP', 'corrG'].forEach(function (id) {
      var el = document.getElementById(id);
      if (!el) return;
      var save = function () {
        var all = (global.Store && global.Store.get) ? global.Store.get('corrInputs', {}) : {};
        all = all || {};
        all[id] = el.value;
        if (global.Store) global.Store.set('corrInputs', all);
      };
      el.addEventListener('change', save);
      el.addEventListener('blur', save);
    });

    // 全选
    Utils.$('cfgAllDz').addEventListener('change', function () {
      var on = this.checked;
      document.querySelectorAll('#cfgTracks input').forEach(function (cb) { cb.checked = on; });
      var dz = Utils.$('cfgDefaultDz');
      if (dz && !on) dz.classList.remove('active');
      syncGroupState();
      runCalculation();
    });

    // 待发股道复选框（垂直排列在右侧）
    var readyWrap = Utils.$('cfgReady');
    // 待发候选 = 到发线 + X 线，名单取自配置（配置里改了范围，这里自动跟随）
    var readyIds = (global.YardConfig ? global.YardConfig.idsOfGroup('td') : [])
      .concat(global.YardConfig ? global.YardConfig.idsOfGroup('x') : []);
    readyIds.forEach(function (id) {
      var item = document.createElement('div');
      item.className = 'cfg-ready-item';
      item.setAttribute('data-ready-id', id);

      var label = document.createElement('label');
      label.className = 'cfg-ready-track';
      label.innerHTML = '<input type="checkbox" data-ready="' + Utils.escapeHtml(id) + '"> ' +
                        '<span class="cfg-ready-name">' + Utils.escapeHtml(trackLabel(id)) + '</span>';

      // 悬浮出现的「+」按钮（SVG 四方块带加号）
      var addBtn = document.createElement('button');
      addBtn.type = 'button';
      addBtn.className = 'cfg-ready-add';
      addBtn.setAttribute('aria-label', '添加车次');
      addBtn.title = '添加车次';
      addBtn.innerHTML = '<svg viewBox="0 0 14 14" width="10" height="10" aria-hidden="true">' +
        '<line x1="7" y1="2" x2="7" y2="12" stroke="currentColor" stroke-width="2"/>' +
        '<line x1="2" y1="7" x2="12" y2="7" stroke="currentColor" stroke-width="2"/></svg>';

      // 展开的车次输入/已录入车次芯片（默认隐藏）
      var holder = document.createElement('div');
      holder.className = 'cfg-ready-train-holder';
      holder.innerHTML =
        '<span class="cfg-ready-chip" data-chip-for="' + Utils.escapeHtml(id) + '">' +
          '<span class="cfg-ready-chip-text"></span>' +
          '<button type="button" class="cfg-ready-chip-del" title="删除" aria-label="删除">×</button>' +
        '</span>' +
        '<input type="text" class="cfg-ready-train" data-train-for="' +
        Utils.escapeHtml(id) + '" maxlength="8" placeholder="车次">';

      label.appendChild(addBtn);
      item.appendChild(label);
      item.appendChild(holder);
      readyWrap.appendChild(item);
    });

    // 展开编辑：显示输入框（预填当前值）、隐藏芯片，焦点落入输入框
    function openEditor(item, id) {
      var holder = item.querySelector('.cfg-ready-train-holder');
      var inp = item.querySelector('.cfg-ready-train');
      var chip = item.querySelector('.cfg-ready-chip');
      if (holder) holder.style.display = 'block';
      if (inp) { inp.value = readyTrains[id] != null ? readyTrains[id] : ''; inp.style.display = 'inline-block'; inp.focus(); }
      if (chip) chip.style.display = 'none';
    }

    // 「+」按钮 / 芯片删除标：点击处理
    readyWrap.addEventListener('click', function (e) {
      // 芯片右上角红色删除标：点击移除该车次
      var del = e.target.closest ? e.target.closest('.cfg-ready-chip-del') : null;
      if (del) {
        e.preventDefault(); e.stopPropagation();
        var item = del.closest('.cfg-ready-item');
        var id = item.getAttribute('data-ready-id');
        delete readyTrains[id];
        if (global.Store) global.Store.set('readyTrains', readyTrains);
        updateReadyChip(item, id);
        // 删除后收起整个编辑区（含输入框），由「+」按钮重新添加
        var holder = item.querySelector('.cfg-ready-train-holder');
        if (holder) holder.style.display = 'none';
        var inp = item.querySelector('.cfg-ready-train');
        if (inp) inp.style.display = '';
        runCalculation();
        return;
      }
      var addBtn = e.target.closest ? e.target.closest('.cfg-ready-add') : null;
      if (!addBtn) return;
      // 阻止 label 的默认行为，避免点击"+"时误切换股道复选框
      e.preventDefault();
      e.stopPropagation();
      var item = addBtn.closest('.cfg-ready-item');
      var holder = item && item.querySelector('.cfg-ready-train-holder');
      if (!holder) return;
      if (holder.style.display === 'block') holder.style.display = 'none';
      else openEditor(item, item.getAttribute('data-ready-id'));
    });
    // 车次输入：实时更新映射并持久记忆、重算（回车后翻转成芯片）
    readyWrap.addEventListener('input', function (e) {
      var inp = e.target.closest ? e.target.closest('.cfg-ready-train') : null;
      if (!inp) return;
      var id = inp.getAttribute('data-train-for');
      var val = inp.value.trim();
      if (val) readyTrains[id] = val; else delete readyTrains[id];
      if (global.Store) global.Store.set('readyTrains', readyTrains);
      runCalculation();
    });
    // 回车确认：把输入框翻转成「车次芯片」，隐藏输入框
    readyWrap.addEventListener('keydown', function (e) {
      var inp = e.target.closest ? e.target.closest('.cfg-ready-train') : null;
      if (!inp || e.key !== 'Enter') return;
      e.preventDefault();
      var item = inp.closest('.cfg-ready-item');
      var id = inp.getAttribute('data-train-for');
      var val = inp.value.trim();
      if (val) readyTrains[id] = val; else delete readyTrains[id];
      if (global.Store) global.Store.set('readyTrains', readyTrains);
      inp.style.display = 'none';
      updateReadyChip(item, id);
      runCalculation();
    });

    // 默认待装
    Utils.$('cfgDefaultDz').addEventListener('click', function () {
      var on = !this.classList.contains('active');
      this.classList.toggle('active', on);
      document.querySelectorAll('#cfgTracks input').forEach(function (cb) { cb.checked = false; });
      if (on) {
        Object.keys(DEFAULT_AREAS).forEach(function (name) {
          DEFAULT_AREAS[name].forEach(function (tr) {
            var cb = Utils.$('cfgTrack_' + tr);
            if (cb) cb.checked = true;
          });
        });
      }
      syncGroupState();
      runCalculation();
    });

    // 勾选即重算 + 组标题整组切换（事件委托）
    tb.addEventListener('change', function (e) {
      if (!e.target || e.target.type !== 'checkbox') return;
      syncGroupState();
      runCalculation();
    });
    tb.addEventListener('click', function (e) {
      var btn = e.target && e.target.closest ? e.target.closest('.cfg-group-btn') : null;
      if (!btn) return;
      toggleGroup(btn.dataset.group);
      runCalculation();   // 整组选择变化即重算
    });
    Utils.$('cfgReady').addEventListener('change', runCalculation);

    // 按钮
    Utils.$('btnStart31814').addEventListener('click', runCalculation);
    Utils.$('btnClear31814').addEventListener('click', function () {
      clearConfig();
      runCalculation();   // 清空后按"未选择"再算一次
    });
    Utils.$('btnCloseCfg31814').addEventListener('click', function () { UI.Modal.close('modal31814'); });

    // 结果区截图（按钮在弹窗头部，此处绑定一次）
    var copyBtn = Utils.$('rptCopyImg');
    if (copyBtn) copyBtn.addEventListener('click', copyResultAsImage);

    // 表格字号缩放（按钮在弹窗头部，此处绑定一次）
    var zi = Utils.$('rptZoomIn'), zo = Utils.$('rptZoomOut');
    if (zi) zi.addEventListener('click', function () { zoomFontSize(FS_STEP); });
    if (zo) zo.addEventListener('click', function () { zoomFontSize(-FS_STEP); });
    applyFontSize();
  }

  function collectConfig() {
    config.drr = {};
    document.querySelectorAll('#cfgTracks input:checked').forEach(function (cb) {
      config.drr[cb.dataset.track] = 1;
    });
    config.crr = {};
    document.querySelectorAll('#cfgReady input:checked').forEach(function (cb) {
      config.crr[cb.dataset.ready] = 1;
    });
  }

  /** 清空选择条件（不显示占位文本，由调用方决定是否立即重算） */
  function clearConfig() {
    document.querySelectorAll('#cfgTracks input').forEach(function (cb) { cb.checked = false; });
    var dz = Utils.$('cfgDefaultDz');
    if (dz) dz.classList.remove('active');
    var all = Utils.$('cfgAllDz');
    if (all) all.checked = false;
    document.querySelectorAll('#cfgReady input').forEach(function (cb) { cb.checked = false; });
    syncGroupState();
    config.drr = {};
    config.crr = {};
    corrApplied = false;   // 清空数据后，待卸纠正回到未套用状态（需重新点校对）
  }

  function runCalculation() {
    if (!currentRawRows.length) {
      Utils.toast('请先加载 xls 数据', 'error');
      return;
    }
    // 股道变化触发的重算一律按自动值呈现；纠正需用户再次点击「待卸校对」
    corrApplied = false;
    collectConfig();
    var brr = setCarProperties(currentRawRows, config.drr, config.crr);
    var stats = calculateStats(brr);
    renderResult(stats);
  }

  /* ========================== 对外接口 ========================== */
  function open(rawRows, fileName) {
    // state 是 app.js IIFE 内的私有对象，未挂到 window，故不存在全局兜底
    currentRawRows = rawRows || [];
    if (!currentRawRows.length) {
      Utils.toast('请先加载 xls 数据', 'error');
      return;
    }
    UI.Modal.open('modal31814');
    ensureConfigUI();
    clearConfig();
    // 打开时恢复上次本地持久记忆的纠正输入
    var savedCorr = (global.Store && global.Store.get) ? global.Store.get('corrInputs', {}) : {};
    savedCorr = savedCorr || {};
    ['corrC', 'corrX', 'corrP', 'corrG'].forEach(function (id) {
      var el = document.getElementById(id);
      if (el) el.value = (savedCorr[id] != null ? savedCorr[id] : '');
    });
    // 打开时恢复车次输入（待发股道 → 车次映射）
    var savedTrains = (global.Store && global.Store.get) ? global.Store.get('readyTrains', {}) : {};
    readyTrains = savedTrains || {};
    document.querySelectorAll('#cfgReady .cfg-ready-item').forEach(function (item) {
      var id = item.getAttribute('data-ready-id');
      if (readyTrains[id] != null) {
        var holder = item.querySelector('.cfg-ready-train-holder');
        if (holder) holder.style.display = 'block';
        var inp = item.querySelector('.cfg-ready-train');
        if (inp) inp.style.display = 'none';
        updateReadyChip(item, id);
      }
    });
    // 打开即以「未选择任何股道」的状态计算一次，
    // 之后用户勾选待装/待发股道会自动重算
    runCalculation();
  }

  global.Report31814 = {
    open: open,
    _setCarProperties: setCarProperties,
    _calculateStats: calculateStats
  };
})(window);