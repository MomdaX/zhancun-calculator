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
  var $ = function (id) { return document.getElementById(id); };
  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  /* ========================== 默认待装区域 ==========================
   * 仅列出实际可装车的货物线 / 装卸线。
   * 到发线（1-15道）等不用于装车，不在待装股道面板中显示。
   * 顺序决定面板内分组先后。
   * ================================================================== */
  var DEFAULT_AREAS = {
    '货场': ['H1', 'H2', 'H3', 'H4', 'H5'],
    '中粮': ['ZL1', 'ZL2', 'ZL3'],
    '大洋': ['LZ'],
    '港务局': ['G1', 'G2'],
    '永鑫': ['YX2', 'YX3'],
    '天煤': ['TS1', 'TS2'],
    '石化': ['SH1', 'SH2'],
    '国投': ['GT1', 'GT2'],
    '超智': ['CZ3', 'CZ4'],
    '广明': ['GM1', 'GM2', 'GM3', 'GM4'],
    '天油': ['TSY1', 'TSY2', 'TSY3', 'TSY4'],
    '中油': ['Y5', 'Y6', 'Y9', 'Y10', 'Y11', 'Y12', 'Y13', 'Y14', 'Y15', 'Y16']
  };

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

  /** 取股道显示名（如 "1道"、"H1"） */
  function trackLabel(id) {
    var t = global.YardConfig && global.YardConfig.getTrack(id);
    return t ? t.name : id;
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
  function firstChar(s) { return String(s == null ? '' : s).charAt(0); }
  function v(val) {
    if (val == null || val === '') return 0;
    var n = parseFloat(String(val).trim());
    return isNaN(n) ? 0 : n;
  }

  function determineCarType(carTypeRaw, carNo) {
    var ct = String(carTypeRaw || '').toUpperCase();
    var cn = String(carNo || '');
    var firstCN = cn.charAt(0);
    if (firstChar(ct) === 'N' && firstCN === '5') return 'X';
    if (firstChar(ct) === 'B') {
      if (ct === 'BH1') return 'P';
      if (firstCN === '5') return 'X';
      if (firstCN === '6') return 'G';
      return firstChar(ct);
    }
    return firstChar(ct);
  }

  function getDirMap() {
    if (!global.Aggregate || !global.DirectionData) return {};
    return global.Aggregate.buildDirectionIndex(global.DirectionData).map;
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
      var carType = determineCarType(carTypeRaw, carNo);

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

  /* ========================== 核心：计算统计 ========================== */
  function calculateStats(brr) {
    var dirMap = getDirMap();
    var types = ['沙口', '南口', '管内', '待卸', '待装', '空车', '待发', '待装自备罐', '自备'];
    var ls = {};
    var 车种顺序 = {};   // 车种按首次出现排序，与 VBA Dictionary 的插入顺序一致
    var gSplit = {};
    var dic = {};
    var openBox = {};  // 敞顶箱车号统计
    var gds = {};
    var 站存 = 0;

    types.forEach(function (type) {
      var dis = {}, dcz = {};
      gSplit[type] = { road: 0, self: 0 };
      brr.forEach(function (row) {
        if (row.status !== type) return;
        dis[row.dest || ''] = (dis[row.dest || ''] || 0) + 1;
        dcz[row.carType] = (dcz[row.carType] || 0) + 1;
        if (row.carType === 'G') {
          if (type === '待装自备罐' || type === '自备') gSplit[type].self += 1;
          else gSplit[type].road += 1;
        }
        if (type === '沙口' || type === '南口' || type === '管内') {
          dic[row.dest] = (dic[row.dest] || 0) + 1;
        }
        if (type === '待发') {
          gds[row.track] = (gds[row.track] || 0) + 1;
          if (row.isOpen) {
            openBox[row.carNo] = (openBox[row.carNo] || 0) + 1;
          }
        }
      });

      ls[type] = Object.keys(dis).reduce(function (a, k) { return a + dis[k]; }, 0);
      // 保留首次出现顺序（VBA Dictionary 的 keys 顺序），避免与 Excel 展示顺序不一致
      车种顺序[type] = Object.keys(dcz).map(function (t) { return { t: t, n: dcz[t] }; });
      ls[type + '车种'] = 车种顺序[type].map(function (o) { return o.t + o.n; }).join(' ') + ' ';
      站存 += ls[type];
    });

    ls.gSplit = gSplit;
    ls.到卸后四项车种 = ls['待卸车种'];  // 兼容名称

    // 沙口/南口/管内 字符串
    var 沙口 = '', 南口 = '', 管内 = '', 管内车种合计 = '';
    Object.keys(dic).forEach(function (st) {
      var dir = dirMap[st] || '';
      if (/沙/.test(dir)) 沙口 = st + dic[st] + ' ' + 沙口;
      else if (/南/.test(dir)) 南口 = st + dic[st] + ' ' + 南口;
      else if (/管内/.test(dir)) {
        var stationCars = {};
        brr.forEach(function (row) {
          if (row.status === '管内' && row.dest === st && row.load > 4) {
            stationCars[row.carType] = (stationCars[row.carType] || 0) + 1;
          }
        });
        // VBA：单项合计 = gr & dcz(gr) & 单项合计 —— 倒序拼接
        var single = '';
        Object.keys(stationCars).forEach(function (t) { single = t + stationCars[t] + single; });
        管内 = st + ':' + dic[st] + ' ' + 管内;
        管内车种合计 = st + ':' + single + ' ' + 管内车种合计;
      }
    });

    // 待发股道合计：按股道顺序升序（数字道在前，X 道在后）
    var 待发股道合计 = Object.keys(gds).sort(compareTrackId)
      .map(function (tr) { return tr + '道/' + gds[tr]; }).join(' ');

    // 敞顶箱列表：敞顶箱/总数 + 编号/数量
    var openTotal = Object.keys(openBox).reduce(function (a, k) { return a + openBox[k]; }, 0);
    var openList = Object.keys(openBox).map(function (k) { return k + '/' + openBox[k]; }).join(' ');

    // 总车种：VBA 为 总数车种合计 = ar & zs(ar) & " " & 总数车种合计（倒序拼接）
    var totalCars = {};
    brr.forEach(function (row) { totalCars[row.carType] = (totalCars[row.carType] || 0) + 1; });
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
    var title = '钦州港 ' + formatDate(new Date()) + ' 早6点 站存';

    var g = stats.ls.gSplit;
    var ord = stats.车种顺序 || {};

    var html = '<table>';
    html += '<tr><td colspan="3" class="title">' + escapeHtml(title) + '</td></tr>';

    function row(label, top, bottom, total, cls) {
      var topHtml = top || '&nbsp;';
      // two-row：含 detail-bottom 的两行结构，上半格高度按 CSS 取下半格的 3 倍
      var clsAll = ((cls || '') + (bottom ? ' two-row' : '')).trim();
      return '<tr class="' + clsAll + '">' +
        '<td class="label">' + escapeHtml(label) + '</td>' +
        '<td><div class="detail-top' + (bottom ? '' : ' only') + '">' + topHtml + '</div>' +
        (bottom ? '<div class="detail-bottom">' + bottom + '</div>' : '') + '</td>' +
        '<td class="total">' + total + '</td></tr>';
    }

    // 沙口（B2 到站 / B3 车种，G 不拆分）
    html += row('沙口', escapeHtml(stats.沙口),
      buildCarLine(ord['沙口'], 0, 0),
      stats.ls['沙口'], 'dir-shakou');
    // 南口（B4 / B5）
    html += row('南口', escapeHtml(stats.南口),
      buildCarLine(ord['南口'], 0, 0),
      stats.ls['南口'], 'dir-nankou');
    // 管内（B6）：单格，管内车种合计已含站名与车种，如「防城港:X7C56」
    html += row('管内', escapeHtml(stats.管内车种合计), '',
      stats.ls['管内'], 'dir-guanna');

    // 待卸（B7）：单格，只有车种合计
    html += row('待卸',
      buildCarLine(ord['待卸'], 0, 0),
      '', stats.ls['待卸'], '');

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
      stats.ls['空车'] + stats.ls['自备'], '');

    // 待开（B16 股道合计 / B17 车种）
    html += row('待开',
      escapeHtml(stats.待发股道合计),
      buildCarLine(ord['待发'], 0, 0),
      stats.ls['待发'], '');

    // 站存（B18 总车种合计，G 不拆分；按 VBA zs 字典的插入顺序）
    html += row('站存',
      buildCarLine(stats.总数车种顺序, 0, 0),
      '', stats.站存, '');

    html += '</table>';
    $('rptBody').innerHTML = html;
    applyFontSize();   // innerHTML 重建后重新套用当前字号
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
    var body = $('rptBody');
    if (body) body.style.setProperty('--rpt-fs', fsCurrent + 'px');

    // 中间显示当前字号
    var val = $('rptFsVal');
    if (val) val.textContent = fsCurrent + 'px';

    var inBtn = $('rptZoomIn'), outBtn = $('rptZoomOut');
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

  /** 轻量提示（复用页面上的 #toast 元素） */
  function rptToast(msg, isErr) {
    var t = $('toast');
    if (!t) { alert(msg); return; }
    t.textContent = msg;
    t.className = 'toast show' + (isErr ? ' error' : ' ok');
    clearTimeout(rptToast._timer);
    rptToast._timer = setTimeout(function () { t.className = 'toast'; }, 2600);
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
    var body = $('rptBody');
    var target = (body && body.querySelector('table')) || body;
    if (!target) { rptToast('没有可截图的内容', true); return; }

    captureElement(target).then(function (blob) {
      if (!navigator.clipboard || typeof window.ClipboardItem === 'undefined') {
        downloadBlob(blob, '站存计算.png');
        rptToast('当前环境不支持剪切板，已下载为图片');
        return;
      }
      navigator.clipboard.write([new window.ClipboardItem({ 'image/png': blob })])
        .then(function () { rptToast('结果区已复制为图片，可直接粘贴'); })
        .catch(function () {
          downloadBlob(blob, '站存计算.png');
          rptToast('复制受限，已保存为图片文件', true);
        });
    }).catch(function (e) {
      rptToast('截图失败：' + (e && e.message ? e.message : e), true);
    });
  }

  function parseCarTypes(str) {
    var cars = {};
    var parts = String(str || '').trim().split(/\s+/).filter(Boolean);
    parts.forEach(function (p) {
      var type = p.replace(/\d/g, '');
      var num = parseInt(p.replace(/\D/g, ''), 10) || 0;
      cars[type] = (cars[type] || 0) + num;
    });
    return { cars: cars };
  }

  function formatDate(d) {
    return (d.getMonth() + 1) + '月' + d.getDate() + '日';
  }

  /** 股道 id 排序：纯数字道在前按数值升序，字母道（X1…）在后按前缀+数值升序 */
  function compareTrackId(a, b) {
    a = String(a); b = String(b);
    var na = /^\d+$/.test(a), nb = /^\d+$/.test(b);
    if (na && nb) return +a - +b;
    if (na) return -1;
    if (nb) return 1;
    var ma = /^(\D*)(\d+)$/.exec(a), mb = /^(\D*)(\d+)$/.exec(b);
    var pa = ma ? ma[1] : a, pb = mb ? mb[1] : b;
    if (pa !== pb) return pa < pb ? -1 : 1;
    return (ma ? +ma[2] : 0) - (mb ? +mb[2] : 0);
  }

  /* ========================== 右侧条件面板 ========================== */
  var currentRawRows = [];
  var currentFileName = '';
  var config = { drr: {}, crr: {} };

  function ensureConfigUI() {
    if ($('cfgPanel').innerHTML.trim()) return;

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

    $('cfgPanel').innerHTML = html;

    // 待装股道：按「作业区 / 线别」分组渲染
    var tb = $('cfgTracks');
    DZ_GROUPS.forEach(function (g) {
      var box = document.createElement('div');
      box.className = 'cfg-track-group' + (g.area ? ' is-area' : '');

      var head = document.createElement('div');
      head.className = 'cfg-group-head';
      head.innerHTML = '<button type="button" class="cfg-group-btn" data-group="' + g.key + '">' +
        escapeHtml(g.name) + '</button>' +
        '<span class="cfg-group-count" data-count="' + g.key + '">0/' + g.ids.length + '</span>';
      box.appendChild(head);

      var wrap = document.createElement('div');
      wrap.className = 'cfg-group-tracks';
      g.ids.forEach(function (id) {
        var label = document.createElement('label');
        label.className = 'cfg-track';
        label.innerHTML = '<input type="checkbox" id="cfgTrack_' + id + '" data-track="' + id + '">' +
          escapeHtml(trackLabel(id));
        wrap.appendChild(label);
      });
      box.appendChild(wrap);
      tb.appendChild(box);
    });

    // 全选
    $('cfgAllDz').addEventListener('change', function () {
      var on = this.checked;
      document.querySelectorAll('#cfgTracks input').forEach(function (cb) { cb.checked = on; });
      var dz = $('cfgDefaultDz');
      if (dz && !on) dz.classList.remove('active');
      syncGroupState();
      runCalculation();
    });

    // 待发股道复选框（垂直排列在右侧）
    var readyWrap = $('cfgReady');
    for (var i = 1; i <= 15; i++) {
      var label = document.createElement('label');
      label.className = 'cfg-ready-track';
      label.innerHTML = '<input type="checkbox" data-ready="' + i + '"> ' + i + '道';
      readyWrap.appendChild(label);
    }
    for (var i = 1; i <= 15; i++) {
      var label = document.createElement('label');
      label.className = 'cfg-ready-track';
      label.innerHTML = '<input type="checkbox" data-ready="X' + i + '"> X' + i;
      readyWrap.appendChild(label);
    }

    // 默认待装
    $('cfgDefaultDz').addEventListener('click', function () {
      var on = !this.classList.contains('active');
      this.classList.toggle('active', on);
      document.querySelectorAll('#cfgTracks input').forEach(function (cb) { cb.checked = false; });
      if (on) {
        Object.keys(DEFAULT_AREAS).forEach(function (name) {
          DEFAULT_AREAS[name].forEach(function (tr) {
            var cb = $('cfgTrack_' + tr);
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
    $('cfgReady').addEventListener('change', runCalculation);

    // 按钮
    $('btnStart31814').addEventListener('click', runCalculation);
    $('btnClear31814').addEventListener('click', function () {
      clearConfig();
      runCalculation();   // 清空后按"未选择"再算一次
    });
    $('btnCloseCfg31814').addEventListener('click', function () { $('modal31814').classList.remove('show'); });

    // 结果区截图（按钮在弹窗头部，此处绑定一次）
    var copyBtn = $('rptCopyImg');
    if (copyBtn) copyBtn.addEventListener('click', copyResultAsImage);

    // 表格字号缩放（按钮在弹窗头部，此处绑定一次）
    var zi = $('rptZoomIn'), zo = $('rptZoomOut');
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
    var dz = $('cfgDefaultDz');
    if (dz) dz.classList.remove('active');
    var all = $('cfgAllDz');
    if (all) all.checked = false;
    document.querySelectorAll('#cfgReady input').forEach(function (cb) { cb.checked = false; });
    syncGroupState();
    config.drr = {};
    config.crr = {};
  }

  function runCalculation() {
    if (!currentRawRows.length) {
      alert('请先加载 xls 数据');
      return;
    }
    collectConfig();
    var brr = setCarProperties(currentRawRows, config.drr, config.crr);
    var stats = calculateStats(brr);
    renderResult(stats);
  }

  /* ========================== 对外接口 ========================== */
  function open(rawRows, fileName) {
    currentRawRows = rawRows || (global.state ? global.state.rawRows : []);
    currentFileName = fileName || (global.state ? global.state.currentFile : '');
    if (!currentRawRows.length) {
      // 注意：window.toast 会被命名为 id="toast" 的元素覆盖，必须判断是否为函数
      if (typeof global.toast === 'function') global.toast('请先加载 xls 数据');
      else alert('请先加载 xls 数据');
      return;
    }
    $('modal31814').classList.add('show');
    ensureConfigUI();
    clearConfig();
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
