/**
 * app.js —— 股道存车主程序
 *
 * 无服务器设计：
 *   - 方向库、股道配置通过 <script src> 加载（file:// 下唯一可靠方式）
 *   - xls 通过 File System Access API 读取，目录句柄存 IndexedDB 实现"打开即自动读取"
 *   - 不支持该 API 的浏览器自动降级为 <input type="file"> 手动选择
 *
 * ============================ 功能区块索引 ============================
 * 本文件（app.js）按职责划分为以下区块，便于定位（行号为当前快照，后续可能偏移）：
 *
 *   [列定义]              COLUMNS / DETAIL_COLS / BIANHAO_* —— 已抽到 js/columns.js（纯常量，零依赖）
 *   [主表渲染]      render / computeTotals / renderEmpty
 *                        —— 主表 tbody 构建、合计、空态（到站着色 renderDest 已抽到 js/dest-color.js，纯函数零 DOM）
 *   [明细抽屉]      updateDetailTitle / openDetail / closeDetail / stepDetail
 *                        —— 股道明细抽屉的打开、翻页、标题
 *   [明细多选]      selectRow / renderDrag / endDrag / clearDetailSel
 *                        —— 明细行拖拽范围选 + 单击切换（bind 内联，未抽组件）
 *   [事件绑定]      bind —— 全部 DOM 事件绑定
 *   [入口]          init —— 启动编排（含向 global 注入 state/render/loading/syncPickFolderBtn 桥接）
 *   [对外接口]      文件末尾  window.YardApp —— 数据 / 配置 / 渲染 / 交互的统一出口
 *
 * 说明：state 为 IIFE 内私有对象，数据源区块（ensurePerm/pickFolder/loadFromDir/
 *       readAndRender/renderFileSwitcher，原耦合最弱）已抽到 js/data-source.js，
 *       通过 init 注入的 global 桥接访问私有资源；index.html 中 data-source.js 在 app.js 之前加载。
 *       另：列定义（COLUMNS/DETAIL_COLS/BIANHAO_*）抽到 js/columns.js，到站着色（renderDest
 *       等纯函数）抽到 js/dest-color.js；三者均为零 DOM 依赖，可被 node 节点测试覆盖。
 * =====================================================================
 */
(function (global) {
  'use strict';

  /* ============================ 全局状态 ============================ */
  var state = {
    dirIndex: null,        // 方向库
    rows: [],              // 聚合结果数组（按配置顺序）
    rawRows: [],           // 原始数据行
    currentFile: null,     // 当前文件名
    dirHandle: null,       // 文件夹句柄
    selectedIdx: -1,
    detailIdx: -1,
    detailSel: null,       // 明细多选行集合（存 r.raw 的下标）
    printDate: null,       // 数据源打印日期（作为停时基准）
    fileList: null,       // 文件夹内文件列表（多文件时使用）
    showEmptyGroups: true  // 是否显示空线分组（分组内全部股道无车时隐藏）
  };

  var $ = Utils.$;
  var on = Utils.on;
  var escapeHtml = Utils.escapeHtml;
  var toast = Utils.toast;
  // 原始数据的列索引常量（定义见 aggregate.js），避免各处散落魔数
  var COL = Aggregate.COL;

  function loading(show, text) {
    $('loadingText').textContent = text || '正在解析…';
    $('loading').className = show ? 'loading show' : 'loading';
  }

  /* 到站富文本着色 / 车型高亮 / 车站名识别 已抽到 js/dest-color.js（纯函数，挂 global） */

  /* =================== 虚拟股道显示/隐藏 =================== */
  /**
   * 当前应显示的行（连同其在 state.rows 中的原始下标）。
   * 保留原始下标是关键：行上的 data-idx 直接用于 openDetail → state.rows[idx]
   * @returns {Array<{r: Object, idx: number}>}
   */
  function visibleRows() {
    var out = [];
    // 当需要隐藏空线分组时，先找出哪些分组内全部股道均无车
    var emptyGroups = {};
    if (!state.showEmptyGroups) {
      var groupStat = {}; // groupName -> { total, hasCar }
      state.rows.forEach(function (r) {
        var cfg = YardConfig.getTrack(r.track);
        if (!cfg) return;
        var gn = cfg.groupName;
        if (!groupStat[gn]) groupStat[gn] = { total: 0, hasCar: 0 };
        groupStat[gn].total++;
        if (r.count > 0) groupStat[gn].hasCar++;
      });
      Object.keys(groupStat).forEach(function (gn) {
        if (groupStat[gn].hasCar === 0) emptyGroups[gn] = true;
      });
    }
    state.rows.forEach(function (r, idx) {
      if (!state.showEmptyGroups) {
        var cfg = YardConfig.getTrack(r.track);
        if (cfg && emptyGroups[cfg.groupName]) return;
      }
      out.push({ r: r, idx: idx });
    });
    return out;
  }

  /**
   * 同步工具栏「选择数据文件夹」按钮的显隐。
   *
   * 该按钮只在「当前没有可用数据文件夹」时出现：
   *   首次使用 / 句柄丢失 / 上次读取失败 / 浏览器不支持自动读取
   * 用户选过一次后，目录句柄存进 IndexedDB，下次启动自动读取，按钮即收起，
   * 之后要换文件夹走「功能 ▾ → 设置 → 修改」。
   *
   * 所有分支都必须经过这里：分散写 style.display 必然漏掉某条路径，
   * 漏掉就会变成「状态栏提示让你点按钮，但按钮是隐藏的」。
   */
  function syncPickFolderBtn(show) {
    var btn = $('btnPickFolder');
    if (btn) btn.style.display = show ? '' : 'none';
  }

  /** 同步工具栏「空线分组」开关的文案与高亮态 */
  function syncVirtualBtn() {
    var btn = $('btnToggleVirtual');
    if (!btn) return;
    var on = state.showEmptyGroups;
    btn.textContent = on ? '隐藏空线分组' : '显示空线分组';
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-pressed', on ? 'true' : 'false');
  }

  /* =================== 渲染主表 =================== */
  /**
   * 计算分组「合并列」所需的跨行信息。
   * 返回与 vis 等长的数组：{ group, start, span }
   *   group 该行的分组名（来自 track.config 的 groupName）
   *   start 是否为该分组在当前可见行中的首行
   *   span  该分组连续占据的行数（用于 rowspan）
   * 注意：隐藏虚拟股道后，虚拟场分组可能整段消失，span 仅统计可见部分。
   */
  function computeGroupSpans(vis) {
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
  function computeBannerSlots(vis) {
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

  /** 作业区横幅行：通栏单行，仅作分段标识，不参与选中 / 明细 / 合计 */
  function bannerRow(a) {
    return '<tr class="area-banner" data-area="' + escapeHtml(a.name) + '">' +
           '<td colspan="' + COLUMNS.length + '" style="background:' + (a.color || '#2b5cb0') + '">' +
           '<span class="area-banner-text">' + escapeHtml(a.name) + '</span>' +
           '</td></tr>';
  }

  function render() {
    // 表头
    $('headRow').innerHTML = COLUMNS.map(function (c) {
      var cls = c.cls || '';
      // 到站列（表头显示为「车辆信息」）：加 col-flex 类（仅用于允许换行），
      // 仍写内联 width，作为独立可拖拽列
      if (c.key === 'dest') cls += ' col-flex';
      var style = ' style="width:' + c.width + 'px"';
      // 分组合并列（group）与股道列（track）表头合并为一个「股道」：
      //  · group 列表头文字留空、去掉右分隔线，并禁止单独拖拽（no-resize）
      //  · track 列表头用横跨 group+track 两列的标签居中显示「股道」
      // 两个 <th> 仍独立存在，列宽/拖拽/记忆逻辑不受影响。
      if (c.key === 'group') {
        return '<th class="' + cls.trim() + ' no-resize"' + style + '></th>';
      }
      if (c.key === 'track') {
        cls += ' grp-merged';
        return '<th class="' + cls.trim() + '"' + style +
               '><span class="grp-head-merge">股道</span></th>';
      }
      return '<th class="' + cls.trim() + '"' + style + '>' + escapeHtml(c.title) + '</th>';
    }).join('');

    // 表头被 innerHTML 重建，需重新挂载列宽拖拽手柄并恢复记忆列宽
    ColResize.safeGet($('grid')).remount();

    // 同步冻结列宽 → 后续 sticky 列的 left 偏移（避免列间露缝；用实测宽度而非 CSS 兜底值）
    var grid = $('grid');
    var ath = grid.querySelector('th.col-a');
    var gth = grid.querySelector('th.col-b-group');
    if (ath) grid.style.setProperty('--col-a-w', ath.offsetWidth + 'px');
    if (gth) grid.style.setProperty('--col-b-group-w', gth.offsetWidth + 'px');

    var vis = visibleRows();
    var thr = YardConfig.thresholds;
    var tbody = $('tbody');
    var html = [];

    // 分组「合并列」：预先算出每行的所属分组、是否该组首行、跨行数(span)。
    // 渲染时首行输出带 rowspan 的分组单元格，组内其余行不输出该 td（由 rowspan 覆盖）。
    var spans = computeGroupSpans(vis);

    // 作业区横幅：定位每个作业区在可见行中的首行（该区无可见行时自动不显示）
    var bannerAt = computeBannerSlots(vis);

    vis.forEach(function (item, n) {
      var r = item.r, idx = item.idx;
      var track = r.track;

      // 作业区横幅行：通栏单行，把主表按作业区断开
      if (bannerAt[n]) html.push(bannerRow(bannerAt[n]));
      var cfg = YardConfig.getTrack(track);
      var zones = YardConfig.getZones(track);

      // 空股道（无车）
      var isBlank = !r.count;

      // 载重超吨
      var overLoad = r.load > thr.overloadTons;
      // 换长超长
      var overLong = r.length > thr.overlong;

      /* 车种列：作业区禁入告警（对齐 VBA 显示信息.bas「盖车不能进鹰岭/栈桥、
       * 高边不能进中油」）。这是「股道 × 车种」的禁入提示——同一辆车在 5 道
       * 不告警、在 Y5 就告警，与「这个车型该显示什么颜色」是两件事，
       * 故独立于「设置 → 车型高亮」配置，且仅在有股道概念的主表显示。 */
      var carTypeCls = '';
      var types = (r.carTypes || '').trim().split(/\s+/);
      if (types.indexOf('P') >= 0) {
        for (var z = 0; z < zones.length; z++) {
          if (zones[z].forbid.indexOf('P') >= 0) {
            carTypeCls = zones[z].mark === 'pink' ? 'warn-pink' : 'warn-e';
            break;
          }
        }
      }
      if (!carTypeCls && types.indexOf('C') >= 0) {
        for (var z2 = 0; z2 < zones.length; z2++) {
          if (zones[z2].forbid.indexOf('C') >= 0) { carTypeCls = 'warn-e'; break; }
        }
      }

      // 汽油股道 → 到达车次列标黄
      var hasOil = false;
      if (r.raw && !YardConfig.isMarkSuppressed(track)) {
        for (var k = 0; k < r.raw.length; k++) {
          if (/汽油|航煤/.test(String(r.raw[k][COL.GOODS] || ''))) { hasOil = true; break; }
        }
      }

      var cells = [];
      COLUMNS.forEach(function (c) {
        var v = r[c.key];
        // 股道列用显示名：到发线显示为「1道」，其余保持原样
        if (c.key === 'track' && cfg) v = cfg.name;
        // 有效长：股道固有属性，按股道 id 查配置（与车辆数据无关，故不在 aggregate 里算）
        if (c.key === 'effLen') v = YardConfig.trackLength(r.track);
        var cls = c.cls || '';
        var style = '';
        var attrs = '';
        var inner;

        if (c.dest) {
          inner = renderDest(v);
          cls += ' dest';
        } else if (c.num) {
          // 换长保留 1 位小数（如 0.0）；其余数值列维持原样
          if (c.key === 'length') {
            inner = (v === 0 || v === '' || v == null) ? '' : Number(v).toFixed(1);
          } else if (c.key === 'effLen') {
            // 单元格显示「换长」= 有效长 ÷ 11（11m 为 1 换长），保留 1 位、第二位起舍去不进位；
            // 原始米数挂 title，鼠标悬停可见
            inner = (v === 0 || v === '' || v == null) ? '' : effLenToChang(v);
          } else {
            inner = (v === 0 || v === '' || v == null) ? '' : escapeHtml(String(v));
          }
          cls += ' num';
        } else {
          var rawNote = (v == null ? '' : String(v));
          // 注意事项列：聚合时已用 \n 分隔各关键词（超71.86吨 / 扣修 …），
          // 转成 <br> 才能逐条换行；方向列同理（render 内已处理）。
          if (c.key === 'note') {
            // 包一层 .note-body，供表头「收起/展开」控件按状态裁剪高度
            inner = '<div class="note-body">' + escapeHtml(rawNote).replace(/\n/g, '<br>') + '</div>';
          } else {
            inner = escapeHtml(rawNote);
            if (!c.cls && !c.num) cls += ' center';
          }
        }

        // 条件样式
        if (c.key === 'load' && overLoad) cls += ' overload';
        if (c.key === 'length' && overLong) cls += ' overlong';
        if (c.key === 'oldCar' && v) cls += ' oldcar';
        if (c.key === 'carTypes' && carTypeCls) cls += ' ' + carTypeCls;
        if (c.key === 'train' && hasOil) cls += ' oil';
        if (c.key === 'train' && String(v || '').charAt(0) === '6') cls += ' train-loop';
        if (c.key === 'track' && isBlank) cls += ' empty-track';
        // 有效长列：原始米数存入 data-len，由 CSS 在单元格右侧「抽屉抽出」显示
        // （不用原生 title —— 它固定弹在下方，且无法定制样式与动画）
        if (c.key === 'effLen' && !(v === 0 || v === '' || v == null)) {
          attrs = ' data-len="' + escapeHtml(v + 'm') + '"';
        }

        // 分组「合并列」：在股道列之前，按分组跨行合并（rowspan）。
        // 仅每组首行输出带 rowspan 的分组单元格，组内后续行不输出（由 rowspan 覆盖）。
        if (c.key === 'group') {
          var sp = spans[n];
          if (sp.start) {
            // 用分组自带的 color 做左侧色条 + 文字着色，醒目区分到发线/调车线/虚拟场等
            var gc = cfg ? cfg.groupColor : '#888';
            cells.push('<td class="col-b-group grp" data-col="group" rowspan="' + sp.span + '" ' +
                       'style="border-left:3px solid ' + gc + ';color:' + gc + '">' +
                       escapeHtml(sp.group) + '</td>');
          }
          return;   // 非首行：被上方 rowspan 覆盖，不输出 td
        }

        /* data-col 记录列 key：分组列用 rowspan 合并后，各行的 td 个数并不一致
         *（非首行少一个分组单元格），td.cellIndex 会整体前移 1 位而不可靠。
         * 因此「编好」等按列定位的逻辑一律用 data-col，禁用 cellIndex。 */
        cells.push('<td class="' + cls + '" data-col="' + c.key + '"' + style + attrs + '>' + inner + '</td>');
      });

      html.push('<tr data-idx="' + idx + '" data-track="' + escapeHtml(track) + '"' +
                (isBlank ? ' class="blank"' : '') + '>' + cells.join('') + '</tr>');
    });

    tbody.innerHTML = html.join('');

    // 合计行（与可见行保持一致，隐藏虚拟股道后合计同步变化）
    var tc = 0, tl = 0, tw = 0, told = 0;
    vis.forEach(function (item) {
      var r = item.r;
      tc += r.count || 0; tl += r.length || 0; tw += r.load || 0; told += r.oldCar || 0;
    });
    // 由 COLUMNS 逐列生成：增删列时合计行自动跟随，
    // 不再需要手工数着补 <td>（原写死 colspan + 固定个数的空 td，加一列就整体错位）。
    var FOOT_VALUES = {
      count: tc,
      length: (Math.round(tl * 10) / 10).toFixed(1),
      load: Math.round(tw * 10) / 10,
      oldCar: told
    };
    var footCells = [];
    COLUMNS.forEach(function (c, i) {
      // 注意事项(冻结) + 分组合并列(冻结) 合计标签横跨这两列
      if (i === 0) { footCells.push('<td class="col-a" colspan="2">合计</td>'); return; }
      if (i === 1) return;                       // 已被上面的 colspan=2 覆盖（分组合并列）
      if (c.key === 'track') { footCells.push('<td class="col-b"></td>'); return; }
      if (c.key === 'dest') { footCells.push('<td id="footTank"></td>'); return; }
      if (Object.prototype.hasOwnProperty.call(FOOT_VALUES, c.key)) {
        footCells.push('<td class="num mid">' + FOOT_VALUES[c.key] + '</td>');
        return;
      }
      footCells.push('<td></td>');
    });
    $('tfoot').innerHTML = '<tr>' + footCells.join('') + '</tr>';

    // 罐车结存（自备罐/路罐）
    var zb = 0, lg = 0;
    vis.forEach(function (item) {
      var re = /(自备罐|路罐)(\d+)/g, m;
      var d = item.r.dest || '';
      while ((m = re.exec(d))) {
        if (m[1] === '自备罐') zb += +m[2]; else lg += +m[2];
      }
    });
    $('footTank').textContent = zb + '(自)/' + lg + '(路)';

    // 状态栏
    $('stTrack').textContent = vis.filter(function (item) { return item.r.count; }).length;
    $('stCount').textContent = tc;
    $('stLen').textContent = Math.round(tl * 10) / 10;
    $('stLoad').textContent = Math.round(tw * 10) / 10;
    $('stOld').textContent = told;

    // 主表重绘即代表「数据或视图已更新」，统一通知订阅者（供外部模块联动）
    notifyDataChange();
  }

  // 货物推算重量默认值（载重缺失时，记事栏命中货物名称 → 用其预设重量）
  var DEFAULT_EST_GOODS = [
    { name: '汽油', weight: 60 },
    { name: '柴油', weight: 60 },
    { name: '航煤', weight: 60 },
    { name: '煤油', weight: 60 }
  ];

  /* =================== 明细抽屉 =================== */
  // 明细「计重」编辑模式：临时删除某行的推算载重（仅影响计重统计，关闭后复原）
  var detailEditMode = false;
  var detailEditExcluded = new Set();
  var detailEvtBound = false;
  /**
   * 单行推算载重：载重有值→原值；为空→记事命中货物名称用预设重量，否则车种含70取70、其余61。
   * 与 computeTotals 共用，保证明细表「载重列」斜体推算值与标题「计重」口径一致。
   * @param {Array} [goodsList] 可选（循环外取一次传入），缺省从 Store 读取默认列表
   */
  function estLoadOf(row, goodsList) {
    var parseNum = Utils.vbVal;
    var rawLoad = row[COL.LOAD];
    if (rawLoad != null && String(rawLoad).trim() !== '') return parseNum(rawLoad);
    if (!Array.isArray(goodsList)) {
      goodsList = Store.get(Store.KEYS.estLoadGoods, null);
      if (!Array.isArray(goodsList)) goodsList = DEFAULT_EST_GOODS;
    }
    var note = String(row[COL.NOTE] || '');
    if (Array.isArray(goodsList)) {
      for (var gi = 0; gi < goodsList.length; gi++) {
        if (goodsList[gi] && goodsList[gi].name && note.indexOf(goodsList[gi].name) >= 0) {
          return parseNum(goodsList[gi].weight);
        }
      }
    }
    var ct = String(row[COL.CARTYPE] || '');
    return ct.indexOf('70') >= 0 ? 70 : 61;
  }
  /** 计算一组明细行的合计：辆数 / 换长 / 总重（总重 = 自重 + 载重，均 1 位小数）。
   *  rows 为 r.raw 的子数组；为空时返回全 0。 */
  function computeTotals(rows, excluded) {
    var parseNum = Utils.vbVal;   // 与聚合引擎同用一套取数规则，保证明细与主表口径一致
    // 推算载重规则（从设置读取，循环外取一次）：载重缺失时，记事命中货物名称 → 用其预设重量；
    // 否则车种含 70 取 70，其余取 61
    var goodsList = Store.get(Store.KEYS.estLoadGoods, null);
    if (!Array.isArray(goodsList)) goodsList = DEFAULT_EST_GOODS;
    var len = 0, selfW = 0, loadW = 0, estLoad = 0;
    for (var k = 0; k < rows.length; k++) {
      var row = rows[k];
      len += parseNum(row[COL.LEN]);                  // 换长
      selfW += parseNum(row[COL.TARE]);               // 自重
      loadW += parseNum(row[COL.LOAD]);               // 载重（按导出原值直接求和）
      // 编辑模式中被「删除」推算值的行：est 记 0（只影响计重，不动自重/总重）
      var est = (excluded && excluded.has(row)) ? 0 : estLoadOf(row, goodsList);
      estLoad += est;
    }
    return {
      count: rows.length,
      length: Math.round(len * 10) / 10,
      selfW: Math.round(selfW * 10) / 10,
      loadW: Math.round(loadW * 10) / 10,
      estLoad: Math.round(estLoad * 10) / 10,        // 推算载重合计（载重缺失按车型/货物补全）
      weight: Math.round((selfW + loadW) * 10) / 10, // 总重 = 自重 + 载重
      calcW: Math.round((selfW + estLoad) * 10) / 10 // 计重 = 自重 + 推算重量（推算载重）
    };
  }
  /** 动态刷新抽屉标题：有选中行时按选中行求和，无选中行时恢复为全部行合计 */
  function updateDetailTitle() {
    var r = state.rows[state.detailIdx];
    if (!r) return;
    var list = r.raw || [];
    var rows = (state.detailSel && state.detailSel.size)
      ? list.filter(function (_, i) { return state.detailSel.has(i); })
      : list;
    var t = computeTotals(rows, detailEditMode ? detailEditExcluded : null);
    var name = YardConfig.getTrack(r.track);
    $('drawerTitle').innerHTML =
      '<span class="dt-name">' + escapeHtml(name ? name.name : r.track) + ' - </span>' +
      '<span class="dt-total">辆数：' + t.count + '</span>' +
      '<span class="dt-total' + (t.length > 70 ? ' warn' : '') + '">换长：' + t.length.toFixed(1) + '</span>' +
      '<span class="dt-total">自重：' + t.selfW.toFixed(1) + '</span>' +
      '<span class="dt-total">载重：' + t.loadW.toFixed(1) + '</span>' +
      '<span class="dt-total dt-weight' + (t.weight > 5000 ? ' warn' : '') + '">总重：' + t.weight.toFixed(1) + '</span>' +
      '<span class="dt-total dt-weight dt-weight-edit' + (detailEditMode ? ' editing' : '') + (t.calcW > 5000 ? ' warn' : '') + '">计重：' + t.calcW.toFixed(1) + '</span>';
  }
  /** 重新渲染当前股道明细（带入 excluded / 编辑态：进入编辑才显示推算值，退出复原为空） */
  function renderCurrentDetail() {
    var r = state.rows[state.detailIdx];
    if (!r) return;
    renderDetailRows(r.raw || [], {
      head: $('detailHead'), body: $('detailBody'), table: $('detailTable')
    }, {
      // 重渲染会重建 tbody，多选高亮需一并恢复（进入/退出编辑各重渲一次）
      rowAttr: function (row, i) {
        return ' data-i="' + i + '"' +
               (state.detailSel && state.detailSel.has(i) ? ' class="selected"' : '');
      },
      excluded: detailEditExcluded,
      editMode: detailEditMode
    });
    updateDetailTitle();
  }
  /** 切换「计重」编辑模式：on=true 进入（表格可删推算值），false 退出（清空 excluded 并复原显示） */
  function setDetailEditMode(on) {
    detailEditMode = on;
    var tbl = $('detailTable');
    if (tbl) tbl.classList.toggle('detail-edit', on);
    var span = $('drawerTitle') && $('drawerTitle').querySelector('.dt-weight-edit');
    if (span) span.classList.toggle('editing', on);
    if (!on) detailEditExcluded.clear();
    // 进入 / 退出都要重渲：进入才显示载重列推算值（供删除），退出则复原为空
    renderCurrentDetail();
  }
  /** 绑定一次：计重 span 双击进入/退出编辑；编辑模式下点击推算值单元格删除该行推算载重
   *  注意：委托到 document 根节点，规避 UI.Drawer.open 重建抽屉 DOM 后原监听丢失。
   *  双击改用「两次 mousedown 间隔检测」实现（不依赖 dblclick，规避部分环境下 dblclick/closest 偶发失效） */
  function bindDetailEditEvents() {
    if (detailEvtBound) return;
    detailEvtBound = true;
    var _wLastT = 0;
    document.addEventListener('mousedown', function (e) {
      if (e.button !== 0) { _wLastT = 0; return; }
      var t = e.target.closest && e.target.closest('.dt-weight-edit');
      if (!t) { _wLastT = 0; return; }
      var now = Date.now();
      if (_wLastT && (now - _wLastT) < 350) {
        setDetailEditMode(!detailEditMode);
        _wLastT = 0;
      } else {
        _wLastT = now;
      }
    });
    document.addEventListener('click', function (e) {
      if (!detailEditMode) return;
      var td = e.target.closest && e.target.closest('td.derived-est');
      if (!td) return;
      var tr = td.closest('tr'); if (!tr) return;
      var idx = tr.getAttribute('data-i');
      var r = state.rows[state.detailIdx];
      var row = (r && r.raw) ? r.raw[idx] : null;
      if (!row || (detailEditExcluded && detailEditExcluded.has(row))) return;
      detailEditExcluded.add(row);
      td.textContent = '';
      td.classList.remove('derived-est');   // 移除该类 → 编辑态下不再显示「删除」按钮
      td.classList.add('excluded-cell');
      updateDetailTitle();   // 重算计重（排除该行推算载重）
    });
  }
  /** 车种/车号颜色规则（对齐 VBA 显示信息.bas） */
  var carStyle = Utils.carStyle;

  /**
   * 把一批车辆行渲染为「明细表」结构（表头 + 表体 + 停时列）。
   * 明细抽屉与「搜索查询」抽屉共用同一份实现，保证两处的列序、着色、格式完全一致。
   *
   * @param {Array}  list 车辆行数组（Aggregate 产出的行，含 __dest / __carType 等派生字段）
   * @param {Object} els  { head, body, table } 三个目标元素
   * @param {Object} [opts]
   * @param {boolean}  [opts.destProcessed] 到站列优先显示「处理后的到站」__dest。
   *        明细抽屉传 false（沿用原行为：优先原始到站，为空才回退 __dest）；
   *        搜索抽屉按需求传 true（统一显示处理后的到站）。
   * @param {Function} [opts.rowAttr] (row, i) → 字符串，追加到 <tr> 上
   *        （明细挂 data-i 供多选/复制定位；搜索挂 hit 高亮与股道/行下标）
   */
  function renderDetailRows(list, els, opts) {
    opts = opts || {};
    list = list || [];

    // 列定义是常量（DETAIL_COLS），表头与每一行共用同一份
    els.head.innerHTML = DETAIL_COLS.map(function (c) {
      // 记事列：加 col-flex 类（仅用于允许换行），仍写内联 width，作为独立可拖拽列
      var cls = (c.t === '记事') ? ' col-flex' : '';
      return '<th class="' + cls.trim() + '" style="width:' + c.w + 'px">' + escapeHtml(c.t) + '</th>';
    }).join('') + '<th style="width:60px">停时h</th>';

    // 表头被 innerHTML 重建，需重新挂载列宽拖拽手柄并恢复记忆列宽。
    // 注意：此处【不】调用 ColResize.reset()——reset 会清掉记忆并重新自适应，
    // 导致用户拖好的列宽在「关闭再打开抽屉」或「切换股道」后失效。
    // 记忆的恢复已由下面的 remount() 完成（有记忆→恢复，无记忆→首次内容自适应）。
    // 如需强制重置为自适应，用设置里的「列宽自适应」按钮（btnAutoFitCols）。
    ColResize.safeGet(els.table).remount();

    var base = state.printDate || new Date();
    els.body.innerHTML = list.map(function (row, i) {
      var cs = carStyle(row);
      var tds = DETAIL_COLS.map(function (c) {
        var raw = row[c.col];
        // 方向列：只有「轻车/空车」按记事识别到站名时，原方向代码"6"（到卸）才失效；
        // 重车（载重>=15）的方向对应实际到站（如到「田东」卸车的「6」是有效信息），不能清空。
        // 仅渲染层，不改数据：31814 统计读 rawRows 的 COL.DIR，不受影响。
        if (c.col === COL.DIR && row.__destIsStation && row.__load < 15) raw = '';
        if (c.fmt === 'track') {
          var t = YardConfig.getTrack(raw);
          return '<td class="center">' +
                 escapeHtml(t ? t.name : (raw == null ? '' : raw)) + '</td>';
        }
        if (c.disp) {
          // 发站（disp:'raw'）：恒用原始发站——每车必有发站，无需派生回退；
          // 也不套用到站那套方向色与标记，仅挂 station-link 供双击开地图。
          if (c.disp === 'raw') {
            return '<td class="dest">' + renderStationLink(raw) + '</td>';
          }
          // 到站列：显示值由「记事是否匹配到方向库站名」决定（渲染层，不改原始数据）
          //   ① 匹配到站名（__destIsStation）→ 显示该站名，原到站是空/钦州港/湛江 都一样；
          //   ② 未匹配到 → 保持原到站；原到站为空时才显示推算的罐型/车种（沿用原行为）。
          // 显示值 ≠ 原到站 → 视为派生，挂 .derived 斜体（方向色由 renderDest 照常叠加）。
          var rawDest = String(row.__destRaw == null ? '' : row.__destRaw).trim();
          var dv = row.__dest == null ? '' : String(row.__dest).trim();
          var show = opts.destProcessed ? dv
                   : (row.__destIsStation ? dv : (rawDest || dv));
          if (show) {
            return '<td class="dest">' +
                   renderDest(show, true, show !== rawDest ? 'derived' : '') +
                   '</td>';
          }
          return '<td class="dest"></td>';
        }
        // 载重列：原值为空 → 推算值【仅「计重」编辑模式下显示】（供逐行删除），平时留空，
        // 避免把推算值误当实载；推算规则见 estLoadOf，与计重口径一致
        if (c.col === COL.LOAD) {
          var isEmptyLoad = (raw == null || String(raw).trim() === '');
          if (isEmptyLoad) {
            if (!opts.editMode) {
              return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '></td>';   // 默认不显示推算值
            }
            if (opts.excluded && opts.excluded.has(row)) {
              return '<td class="mid excluded-cell"></td>';   // 编辑模式内已删除：显示空
            }
            var ev = estLoadOf(row);
            var clsL = [c.cls, 'derived-est'].filter(Boolean).join(' ');
            return '<td' + (clsL ? ' class="' + clsL + '"' : '') + '>' +
                   String(Math.round(ev)) + '</td>';
          }
          return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '>' +
                 escapeHtml(raw == null ? '' : raw) + '</td>';
        }
        // 车种列
        if (c.col === COL.CARTYPE) {
          var clsA = [c.cls, cs.cls, cs.bg].filter(Boolean).join(' ');
          return '<td' + (clsA ? ' class="' + clsA + '"' : '') + '>' +
                 escapeHtml(raw == null ? '' : raw) + '</td>';
        }
        // 车号列：标记 data-col，供「双击复制车号」识别；title 提示该交互
        if (c.col === COL.CARNO) {
          var clsB = [c.cls, cs.clsN, cs.bgN].filter(Boolean).join(' ');
          return '<td data-col="carno" title="双击复制车号"' + (clsB ? ' class="' + clsB + '"' : '') + '>' +
                 escapeHtml(raw == null ? '' : raw) + '</td>';
        }
        // 品名列：汽油/航煤标黄底（对齐 VBA 显示信息.bas）
        if (c.col === COL.GOODS) {
          var pm = raw == null ? '' : String(raw);
          var jishi = row[COL.NOTE] == null ? '' : String(row[COL.NOTE]);
          if ((pm.indexOf('汽油') !== -1 || pm.indexOf('航煤') !== -1 || jishi.indexOf('汽油') !== -1) && jishi !== '原装汽油') {
            var clsP = c.cls ? c.cls + ' car-yellow-bg' : 'car-yellow-bg';
            return '<td class="' + clsP + '">' + escapeHtml(pm) + '</td>';
          }
        }
        return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '>' +
               escapeHtml(raw == null ? '' : raw) + '</td>';
      }).join('');
      var attr = opts.rowAttr ? (opts.rowAttr(row, i) || '') : '';
      return '<tr' + attr + '>' + tds + '<td class="stay"></td></tr>';
    }).join('');

    // 停时列：需单独计算
    // 复用 Utils.parseArriveTime：兼容 Date 实例、"2026/9/2T08:30:00"、"2026年9月2日" 等写法
    var body = els.body.querySelectorAll('tr');
    list.forEach(function (row, i) {
      // 优先复用 aggregate 预处理已缓存的 __arrTime；缺省时回退解析（兼容非聚合来源的行）
      var d = (row.__arrTime != null) ? row.__arrTime : Utils.parseArriveTime(row[COL.ARRTIME]);
      var hrs = d ? Math.floor((base - d) / 3600000) : '';
      var td = body[i] && body[i].querySelector('td.stay');
      if (td) {
        td.textContent = hrs;
        if (hrs !== '' && hrs > YardConfig.thresholds.bigCarHours) {
          td.style.background = '#fff3cd';
          td.style.fontWeight = '700';
        }
      }
    });
  }

  function openDetail(idx) {
    var r = state.rows[idx];
    if (!r) return;
    state.detailIdx = idx;
    state.detailSel = new Set();   // 重置多选（每次打开明细都清空选中）
    // 切换股道时退出编辑模式，避免 excluded 残留指向旧股道
    detailEditMode = false;
    detailEditExcluded.clear();
    var tbl = $('detailTable'); if (tbl) tbl.classList.remove('detail-edit');
    bindDetailEditEvents();

    renderDetailRows(r.raw || [], {
      head: $('detailHead'),
      body: $('detailBody'),
      table: $('detailTable')
    }, {
      rowAttr: function (row, i) { return ' data-i="' + i + '"'; },
      excluded: detailEditExcluded,
      editMode: detailEditMode   // 打开时已重置为 false：推算值默认不显示
    });

    updateDetailTitle();
    UI.Drawer.open('drawer');
  }

  /**
   * 打开某股道的明细抽屉，并高亮 + 滚动定位到指定车辆行。
   * 供「搜索查询」结果点击时调用（搜索抽屉不关闭，关掉明细后仍在原处）。
   * @param {number} trackIdx 股道在 state.rows 中的下标
   * @param {number} [carIdx] 车辆在该股道 raw 中的下标；缺省则只打开不定位
   */
  function openDetailAt(trackIdx, carIdx) {
    openDetail(trackIdx);
    if (carIdx == null) return;
    var tr = $('detailBody').querySelector('tr[data-i="' + carIdx + '"]');
    if (!tr) return;
    tr.classList.add('hit');
    if (tr.scrollIntoView) tr.scrollIntoView({ block: 'center' });
  }

  function closeDetail() {
    UI.Drawer.close('drawer');
  }

  /** 切换上/下一股道：只在当前可见的行之间移动，跳过被隐藏的虚拟股道 */
  function stepDetail(step) {
    var vis = visibleRows().map(function (item) { return item.idx; });
    var cur = vis.indexOf(state.detailIdx);
    if (cur < 0) {
      // 当前股道已被隐藏：按方向定位到最近的可见位置
      var target = -1;
      for (var i = 0; i < vis.length; i++) {
        if (step > 0) {
          if (vis[i] > state.detailIdx) { target = i; break; }
        } else if (vis[i] < state.detailIdx) {
          target = i;   // 不 break：取最后一个小于当前值的
        }
      }
      // 使下一步 next = cur + step 正好落在 target 上；
      // 找不到 target 时置为边界外，由下方统一提示"已是第一/最后一条"
      cur = step > 0 ? (target >= 0 ? target - 1 : vis.length - 1)
                     : (target >= 0 ? target + 1 : 0);
    }
    var next = cur + step;
    if (next < 0 || next >= vis.length) {
      toast('已经是' + (step < 0 ? '第一' : '最后一') + '条股道');
      return;
    }
    openDetail(vis[next]);
  }

  /* =================== 地图径路 =================== */
  /**
   * 明细中双击车站名 → 打开地图生成径路
   *
   * 明细的到站、发站两列在这里都视为「查询站」（即终到站）：
   * 起点固定钦州港，被点击的站作为终点，交给 Map 自身的搜索流程处理。
   */
  function openStationMap(span) {
    var station = span.getAttribute('data-station');
    if (!station) return;

    // 排除补显的罐车分类（路罐/自备罐等），它们不是车站名。
    // 卸车地点（永鑫/货场 等）与黑罐细化子类（中粮/外运）也不是车站，一并排除。
    if (NOT_STATION[station] || isUnloadSpot(station) || isBlackTankSpot(station) ||
        /^[CXPGYWTBDKN]/.test(station)) {
      toast('「' + station + '」不是车站名，无法生成径路', 'error');
      return;
    }

    if (typeof MapBridge === 'undefined') {
      toast('地图模块未加载', 'error');
      return;
    }

    MapBridge.open({ to: station, title: '钦州港 → ' + station + '（径路）' });
  }

  /* =================== 事件绑定 =================== */

  /**
   * 给「车辆行容器」绑定通用双击交互：
   *   · 双击车站名(span.station-link) → 打开地图径路
   *   · 双击车号(td[data-col="carno"]) → 复制车号到剪贴板
   * 明细抽屉(detailBody)与搜索抽屉(searchBody)共用，保证两处交互一致。
   * 注：行点击/拖拽多选语义不同（明细=多选，搜索=打开并定位），不在此统一。
   * @param {string} elId 容器元素 id
   */
  function bindCarRowEvents(elId) {
    on(elId, 'dblclick', function (e) {
      var sp = e.target.closest('span.station-link');
      if (sp) {
        e.stopPropagation();
        openStationMap(sp);
        return;
      }
      var td = e.target.closest('td[data-col="carno"]');
      if (!td) return;
      var txt = (td.textContent || '').trim();
      if (!txt) return;
      Utils.copyText(txt).then(function () {
        toast('已复制车号：' + txt);
      }).catch(function (err) {
        toast('复制失败：' + (err && err.message ? err.message : err), 'error');
      });
    });
  }

  function bind() {
    on('fileInputMulti', 'change', function (e) {
      var f = e.target.files[0];
      if (f) DataSource.readAndRender(f, f.name, null);
      e.target.value = '';
    });

    // 单文件选择：直接打开系统文件对话框，可读取 Program Files 等
    // 被 showDirectoryPicker 禁止的目录（Chrome 只限制"文件夹"访问路径）
    on('btnPickFile', 'click', function () {
      $('fileInputSingle').click();
    });
    on('fileInputSingle', 'change', function (e) {
      var f = e.target.files[0];
      if (!f) { e.target.value = ''; return; }
      DataSource.readAndRender(f, f.name, null).then(function () {
        // 以文件方式载入时隐藏文件切换器与刷新（无目录可扫描）
        var sw = $('fileSwitcher');
        if (sw) sw.style.display = 'none';
        $('stMsg').textContent = '已载入文件：' + f.name;
      }).catch(function () { /* 错误已在 readAndRender 内提示 */ });
      e.target.value = '';
    });

    on('btnPickFolder', 'click', function () {
      DataSource.pickFolder();
    });

    on('btnReload', 'click', function () {
      if (state.dirHandle) DataSource.loadFromDir(state.dirHandle, false);
      else if (state.fileList && state.fileList.length) {
        DataSource.readAndRender(state.fileList[0].handle.getFile(), state.fileList[0].name, state.fileList);
      } else toast('请先选择文件夹或文件');
    });

    // 搜索查询：打开独立的搜索抽屉（实现在 search.js，通过 window.Search 调用）
    on('btnSearch', 'click', function () {
      if (window.Search && window.Search.open) window.Search.open();
      else toast('搜索模块未加载', 'error');
    });

    // 双击行 → 明细（作业区横幅行不是数据行，跳过）
    on('tbody', 'dblclick', function (e) {
      var tr = e.target.closest('tr');
      if (!tr || tr.classList.contains('area-banner')) return;
      openDetail(+tr.getAttribute('data-idx'));
    });

    // 明细表：双击车站名→地图、双击车号→复制车号（通用交互，见 bindCarRowEvents）
    bindCarRowEvents('detailBody');
    // 搜索表：复用同一套双击交互（之前缺失，导致搜索结果里车号无法复制 / 车站无法开地图）
    bindCarRowEvents('searchBody');

    // 明细中：按下行 → 拖动多选（拖动中实时调整范围，松开确定）；单击 → 切换选中
    var dragSel = { active: false, moved: false, anchor: -1, snap: null, mode: 'add' };
    function selectRow(i, on) {
      if (!state.detailSel) state.detailSel = new Set();
      var tr = $('detailBody').querySelector('tr[data-i="' + i + '"]');
      if (on) { state.detailSel.add(i); if (tr) tr.classList.add('selected'); }
      else { state.detailSel.delete(i); if (tr) tr.classList.remove('selected'); }
    }
    function clearDetailSel() {
      if (detailEditMode) return;   // 编辑模式下禁止清空选中（防止与计重编辑冲突）
      if (state.detailSel) state.detailSel.clear();
      var sels = $('detailBody').querySelectorAll('tr.selected');
      for (var k = 0; k < sels.length; k++) sels[k].classList.remove('selected');
      updateDetailTitle();
    }
    // 拖动中按「快照 + (anchor..cur) 按 mode 应用」实时渲染选中
    function renderDrag(cur) {
      if (!state.detailSel) state.detailSel = new Set();
      var snap = dragSel.snap, add = dragSel.mode === 'add';
      // 还原快照
      var sels = $('detailBody').querySelectorAll('tr');
      for (var k = 0; k < sels.length; k++) {
        var idx = +sels[k].getAttribute('data-i');
        if (snap.has(idx)) { state.detailSel.add(idx); sels[k].classList.add('selected'); }
        else { state.detailSel.delete(idx); sels[k].classList.remove('selected'); }
      }
      if (cur >= 0 && dragSel.anchor >= 0) {
        var a = Math.min(dragSel.anchor, cur), b = Math.max(dragSel.anchor, cur);
        for (var j = a; j <= b; j++) selectRow(j, add);
      }
      updateDetailTitle();
    }
    on('detailBody', 'mousedown', function (e) {
      if (detailEditMode) return;   // 编辑模式下禁止拖选行
      var tr = e.target.closest('tr');
      if (!tr || tr.querySelector('td.stay') === e.target) return;
      e.preventDefault();
      if (!state.detailSel) state.detailSel = new Set();
      dragSel.active = true;
      dragSel.moved = false;
      dragSel.anchor = +tr.getAttribute('data-i');
      dragSel.snap = new Set(state.detailSel);   // 记录拖动前选中快照
      // 起点已选中 → 取消模式；否则 → 加入模式（仅用于拖动，单击在 click 中处理）
      dragSel.mode = state.detailSel.has(dragSel.anchor) ? 'del' : 'add';
    });
    on('detailBody', 'mouseover', function (e) {
      if (!dragSel.active) return;
      var tr = e.target.closest('tr');
      if (!tr) return;
      var i = +tr.getAttribute('data-i');
      dragSel.moved = true;
      renderDrag(i);                             // 实时按当前行调整整段
    });
    function endDrag() {
      if (dragSel.active) {
        dragSel.active = false;
        dragSel.anchor = -1;
        dragSel.snap = null;
      }
    }
    document.addEventListener('mouseup', endDrag);
    // 单击（未发生拖动）时切换该行选中状态；拖动已在 mouseover 中实时应用
    on('detailBody', 'click', function (e) {
      if (detailEditMode) return;   // 编辑模式下禁止点选/多选行
      if (dragSel.moved) { dragSel.moved = false; return; }
      var tr = e.target.closest('tr');
      if (!tr || tr.querySelector('td.stay') === e.target) return;
      var i = +tr.getAttribute('data-i');
      selectRow(i, !state.detailSel.has(i));
      updateDetailTitle();
    });
    // 点击表格外（如抽屉头/标题，或表格任意外部区域）清空选中
    var dh = document.querySelector('.drawer-head');
    if (dh) dh.addEventListener('click', clearDetailSel);
    on('drawerTitle', 'click', clearDetailSel);
    document.addEventListener('mousedown', function (e) {
      // 正在拖动表格内选择时不触发清空
      if (dragSel.active) return;
      var dt = $('detailTable');
      if (dt && !dt.contains(e.target)) clearDetailSel();
    });

    /* ---------- 「编好」伪元素按钮（空箱/空车列 = 表头 th[9]） ---------- */

    /** 清除所有「编好」标记 */
    function clearBianhao() {
      var list = $('tbody').querySelectorAll('td.bianhao-on');
      for (var i = 0; i < list.length; i++) list[i].classList.remove('bianhao-on');
    }

    /** 该股道是否落在「编好」生效区间（1道 ~ X15，排除 B1/B2，按 track.config.js 的股道顺序） */
    function inBianhaoRange(trackId) {
      if (BIANHAO_EXCLUDE.indexOf(trackId) >= 0) return false;
      var t = YardConfig.getTrack(trackId);
      var a = YardConfig.getTrack(BIANHAO_FROM);
      var b = YardConfig.getTrack(BIANHAO_TO);
      if (!t || !a || !b) return false;
      var lo = Math.min(a.index, b.index), hi = Math.max(a.index, b.index);
      return t.index >= lo && t.index <= hi;
    }

    /** 根据 trackId 在 state.rows 中查找对应行数据 */
    function findRowByTrack(trackId) {
      if (!state.rows) return null;
      for (var i = 0; i < state.rows.length; i++) {
        if (state.rows[i].track === trackId) return state.rows[i];
      }
      return null;
    }

    /** 在静态 mockup HTML 中设置单元格文本（按 FineReport ID 模式 B4-0-73 查找） */
    function setCellText(doc, cellId, value) {
      if (!doc) return;
      var td = doc.querySelector('[id^="' + cellId + '-"]');
      if (td) {
        // 清除旧内容，放入纯文本
        td.innerHTML = '';
        var div = doc.createElement('div');
        div.style.maxHeight = '28px';
        div.textContent = value;
        td.appendChild(div);
      }
    }

    /** 选中行后刷新「编好」标记：给该行的目标列单元格打标记 → ::after 浮现。
     *  行级触发——点这一行任意单元格都会显示，无需点到第 9 列本身。 */
    function syncBianhao(tr) {
      clearBianhao();
      if (!tr || tr.classList.contains('area-banner')) return;
      if (tr.classList.contains('blank')) return;   // 空股道不显示「编好」
      if (!inBianhaoRange(tr.getAttribute('data-track'))) return;
      var td = tr.querySelector('td[data-col="' + BIANHAO_COL + '"]');
      if (td) td.classList.add('bianhao-on');
    }

    // 单击选中（作业区横幅行不参与选中，否则会被高亮且 selectedIdx 变为 NaN）
    on('tbody', 'click', function (e) {
      var tr = e.target.closest('tr');
      if (!tr || tr.classList.contains('area-banner')) return;
      var old = $('tbody').querySelector('tr.selected');
      if (old) old.classList.remove('selected');
      tr.classList.add('selected');
      state.selectedIdx = +tr.getAttribute('data-idx');
      // 选中这一行即刷新「编好」标记（无需点击第 9 列）
      syncBianhao(tr);
    });

    // 空线分组显示/隐藏开关
    on('btnToggleVirtual', 'click', function () {
      state.showEmptyGroups = !state.showEmptyGroups;
      Store.set('showEmptyGroups', state.showEmptyGroups);   // 持久记忆，刷新后保持
      syncVirtualBtn();
      render();
      toast((state.showEmptyGroups ? '已显示' : '已隐藏') + '空线分组', 'ok');
    });

    // 注意事项列：点击行内右侧按钮 → 整表收起/展开（捕获阶段，先于行选中）
    var gridEl = $('grid');
    if (gridEl) {
      gridEl.addEventListener('click', function (e) {
        var td = e.target.closest ? e.target.closest('td.col-a') : null;
        if (!td) return;
        // 仅当点击在单元格右侧按钮区域（右 32px）时触发
        var rect = td.getBoundingClientRect();
        if (e.clientX - rect.left < rect.width - 32) return;
        e.stopPropagation();
        e.preventDefault();
        var collapsed = gridEl.classList.toggle('notes-collapsed');
        Store.set('notesCollapsed', collapsed);
        toast(collapsed ? '已收起注意事项' : '已展开注意事项', 'ok');
      }, true);  // true = 捕获阶段，先于 tbody 行选中触发
    }

    // 「编好」伪元素按钮：点击已标记单元格左侧按钮区 → 打开「发车作业全流程」浮窗，
    // 并自动读取当前行股道信息填入 iframe 表格（B4 股道 / D4 辆数 / E4 换长 / F4 尾车车号）。
    // 与 col-a「收起/展开」同一套路：伪元素不是事件目标，只能用坐标判定左侧命中区；
    // 用捕获阶段 + stopPropagation，避免冒泡到 tbody 的行选中把标记清掉。
    if (gridEl) {
      gridEl.addEventListener('click', function (e) {
        var td = e.target.closest
          ? e.target.closest('td[data-col="' + BIANHAO_COL + '"].bianhao-on')
          : null;
        if (!td) return;
        var rect = td.getBoundingClientRect();
        // 仅当点击落在单元格左侧按钮区（BIANHAO_HIT_W）时才触发
        if (e.clientX - rect.left > BIANHAO_HIT_W) return;
        e.stopPropagation();
        e.preventDefault();

        // 读取当前行股道数据
        var tr = td.parentNode;
        var trackId = tr.getAttribute('data-track');
        var row = findRowByTrack(trackId);
        if (!row) { toast('未找到股道数据', 'warn'); return; }
        var cfg = YardConfig.getTrack(trackId);
        var trackName = cfg ? cfg.name : trackId;
        var count = row.count || 0;
        var length = (row.length === 0 || row.length == null) ? '' : Number(row.length).toFixed(1);
        var lastCar = '';
        if (row.raw && row.raw.length) {
          var last = row.raw[row.raw.length - 1];
          lastCar = String(last[COL.CARNO] == null ? '' : last[COL.CARNO]);
        }

        // 打开浮窗
        UI.Modal.open('modalDeparture');

        // 给 iframe 一点时间加载，然后填入单元格
        var depFrame = document.getElementById('depFrame');
        var depCheci = document.getElementById('depCheci');
        if (depCheci) depCheci.value = '';
        // 记录当前股道换长 / 重量，用于「列车超长 / 列车超重」提示
        depWarnCtx.length = Number(row.length) || 0;
        depWarnCtx.load = Number(row.load) || 0;
        depWarn.dup = false;
        depDupState = false;
        refreshDepWarnByTrack();

        function fillLocalTable() {
          // 预填 4 格（股道/辆数/换长/尾车车号）到本地表格；车次留空待用户输入
          // 不再立即向报表发送——由用户点下方任一「发送」按钮触发，便于测试哪种 iframe 定位方式能跑通
          var $b4 = $('depInputB4'), $c4 = $('depInputC4'),
              $d4 = $('depInputD4'), $e4 = $('depInputE4'), $f4 = $('depInputF4');
          if ($b4) $b4.value = trackName;
          if ($c4) $c4.value = '';
          if ($d4) $d4.value = String(count);
          if ($e4) $e4.value = length;
          if ($f4) $f4.value = lastCar;
          refreshDepTime();
          if ($c4) setTimeout(function () { $c4.focus(); $c4.select(); }, 80);
        }

        fillLocalTable();
        // 向报表页要一份最新的「编组车次」列表（扩展桥接通道），用于车次重复校验
        try { window.postMessage({ channel: DEP_BRIDGE, type: 'readCheci', ts: Date.now() }, '*'); } catch (e) {}
      }, true);
    }

    /* ================= 发车浮窗警示标 =================
     * 三类提示（任一命中即显示红色三角标，悬停在其后方以编号列表全部列出）：
     *   1. 当前车次重复 —— 与 iframe「编组车次」列（td[col="2"]）已有车次相同
     *   2. 列车超长     —— 本股道换长 > YardConfig.thresholds.overlong（70.0）
     *   3. 列车超重     —— 本股道重量 > YardConfig.thresholds.overloadTons（5000）
     * 均为「仅提示」：不拦截填表，车次照常写入 C4。 */
    var depWarn = { dup: false, overlong: false, overweight: false };
    var depWarnCtx = { length: 0, load: 0 };   // 当前股道的 换长 / 重量（点「编好」时记录）
    // 注：不再保留 depCells 之类的「待填单元格」内存快照——5 个值只在 #depEntry 的 input 里，
    //     发送时由 getDepCells() 现取，关窗即丢，不做任何持久化。

    /** 报表是否「不可访问」：跨域 / 被 X-Frame-Options 拒绝时，
     *  无法读写 iframe 内容（填表与查重都会静默失败），仅提示一次，避免反复打扰。 */
    var depBlocked = false, depBlockedHinted = false;
    function hintDepBlocked() {
      if (!depBlocked || depBlockedHinted) return;
      depBlockedHinted = true;
      toast('发车流程报表无法自动填表：与页面不同源或被禁止嵌入', 'warn');
    }

    /** 取 iframe 文档；跨域或未加载返回 null */
    function getDepDoc() {
      var f = document.getElementById('depFrame');
      if (!f) return null;
      try { return f.contentDocument || (f.contentWindow && f.contentWindow.document) || null; }
      catch (e) { depBlocked = true; return null; }   // 跨域
    }

    /** 扩展桥接（dep-bridge-extension）：页面 → content script → background → 报表标签页。
     *  跨域可用（不要求同源），前提：浏览器装了该扩展并已启用。
     *  未装扩展时这几行静默无效，不影响其它填表通道。 */
    var DEP_BRIDGE = '__DEP_BRIDGE__';
    /* （已移除）depCells / pushDepBridge / scheduleDepBridge / broadcastDepFill / writeDepCell
     * 原因：新流程是「点「编好」→ 本地表格按当前股道重新读取预填 → 用户点发送按钮
     *       → 经扩展桥一次性发到报表页」。本地录入表的数据【不做任何持久化】：
     *       既不写 localStorage（原 broadcastDepFill 会写 zhancun.depFill），
     *       也不走 BroadcastChannel，每次打开浮窗都按流程重新读取，关窗即丢弃。 */

    /** 独立标签页里的报表回传的「编组车次」列表（同源广播 / localStorage），
     *  用于 iframe 读不到时（跨域或改用独立标签页）仍能查重。 */
    var depRemoteCheci = [];
    function bindDepReceive() {
      try { var s = localStorage.getItem('zhancun.depCheci'); if (s) depRemoteCheci = JSON.parse(s) || []; }
      catch (e) {}
      try {
        if (typeof BroadcastChannel === 'undefined') return;
        var bc = new BroadcastChannel('zhancun-dep');
        bc.onmessage = function (ev) {
          var d = ev.data;
          if (d && d.type === 'checiList' && d.list) depRemoteCheci = d.list;
        };
      } catch (e) {}
      // 跨域回传：① window.opener.postMessage（本页 window.open 的标签页）
      //           ② 浏览器扩展中转（dep-bridge-extension，带 channel 标记）
      window.addEventListener('message', function (ev) {
        var d = ev.data;
        if (!d || d.type !== 'checiList' || !d.list) return;
        if (d.channel && d.channel !== DEP_BRIDGE) return;   // 别的通道的消息不收
        depRemoteCheci = d.list;
        setDepWarn();                                        // 拿到列表后按当前车次重算提示
      });
    }
    bindDepReceive();

    /** 车次是否已在报表「编组车次」列中出现（当前填写行 C4 不计入） */
    function depCheciDup(val) {
      var v = String(val == null ? '' : val).replace(/\s+/g, '').toUpperCase();
      if (!v) return false;
      var doc = getDepDoc();
      if (doc) {
        var tds = doc.querySelectorAll('td[col="2"]');
        for (var i = 0; i < tds.length; i++) {
          var td = tds[i];
          var id = td.getAttribute('id') || '';
          if (id.indexOf('C4-') === 0) continue;            // 当前正在填写的输入行
          var txt = (td.textContent || '').replace(/\s+/g, '').toUpperCase();
          if (txt && txt === v) return true;
        }
      }
      // iframe 读不到时，用报表页回传的车次列表再比对一次
      for (var k = 0; k < depRemoteCheci.length; k++) {
        if (String(depRemoteCheci[k]).replace(/\s+/g, '').toUpperCase() === v) return true;
      }
      return false;
    }

    /** 按固定顺序汇总当前命中的提示文案 */
    function depWarnList() {
      var list = [];
      if (depWarn.dup) list.push('当前车次重复');
      if (depWarn.overlong) list.push('列车超长');
      if (depWarn.overweight) list.push('列车超重');
      return list;
    }

    /** 刷新警示标：有提示 → 显示三角标 + 悬停列表；无提示 → 隐藏 */
    function setDepWarn() {
      var list = depWarnList();
      var w = $('depCheciWarn');
      var tip = $('depWarnTip');
      var inp = $('depInputC4');
      var on = list.length > 0;
      if (w) {
        w.classList.toggle('show', on);
        // 不再写 title：浏览器原生 tooltip 与红色气泡会同时出现（双框），保留自定义气泡即可
        if (!on) w.removeAttribute('title');
      }
      // 输入框红框只跟「车次重复」走（那是输入内容本身的问题）
      if (inp) inp.classList.toggle('dup', !!depWarn.dup);
      if (tip) {
        tip.innerHTML = on
          ? '<ol>' + list.map(function (t, i) {
              return '<li>' + (i + 1) + '. ' + escapeHtml(t) + '</li>';
            }).join('') + '</ol>'
          : '';
      }
    }

    /** 依据当前股道的换长 / 重量刷新「超长 / 超重」提示 */
    function refreshDepWarnByTrack() {
      var thr = (YardConfig && YardConfig.thresholds) || {};
      var maxLen = thr.overlong == null ? 70 : thr.overlong;
      var maxWt = thr.overloadTons == null ? 5000 : thr.overloadTons;
      depWarn.overlong = Number(depWarnCtx.length) > maxLen;
      depWarn.overweight = Number(depWarnCtx.load) > maxWt;
      setDepWarn();
    }

    // 车次输入框：输入中只做查重提示（红三角标），不写表、不广播；
    // 真正把车次写入 C4 并发送到报表页，仅在「回车 / 失焦」时（pushDepCheciToTab）才发生
    var depDupState = false;
    function applyDepCheci(showToast) {
      var inp = $('depInputC4');
      if (!inp) return;
      var val = inp.value.trim();
      depWarn.dup = val ? depCheciDup(val) : false;
      setDepWarn();                                  // 只更新警示标，不写 C4
      if (depWarn.dup && !depDupState && showToast) toast('当前车次重复', 'warn');
      depDupState = depWarn.dup;
    }
    on('depInputC4', 'input', function () { applyDepCheci(true); });   // 输入中：仅提示
    // 注：车次框按 Enter 不再自动发送——由用户自己点下方发送按钮触发，避免误发。
    // 输入完成后按 Enter → 失焦（blur），方便收起输入法 / 退出当前框
    on('depInputC4', 'keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); this.blur(); }
    });

    // ============== 时间显示（每秒刷新）==============
    function pad2(n) { return n < 10 ? '0' + n : '' + n; }
    function formatDepTime() {
      var d = new Date();
      return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()) +
             ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds());
    }
    function refreshDepTime() {
      var t = $('depTime');
      if (t) t.textContent = formatDepTime();
    }
    refreshDepTime();
    setInterval(refreshDepTime, 1000);

    // ============== 浮窗内右下角状态提示（替代全局 toast）==============
    // 默认 3 秒自动淡出；鼠标悬停时暂停消失（便于看完整回执），移开后 0.8 秒再消失
    var depStatusTimer = null;
    function hideDepStatus(delay) {
      var el = $('depStatus');
      if (!el) return;
      if (depStatusTimer) clearTimeout(depStatusTimer);
      depStatusTimer = setTimeout(function () { el.classList.remove('show'); }, delay);
    }
    function showDepStatus(msg, type) {
      var el = $('depStatus');
      if (!el) return;
      el.textContent = msg;
      el.className = 'dep-status show ' + (type || 'info');
      hideDepStatus(3000);
    }
    // 悬停暂停 / 移开消失（回执文字较长，常含「目标=…」「错误=…」，需要停留查看）
    (function () {
      var el = $('depStatus');
      if (!el) return;
      el.addEventListener('mouseenter', function () {
        if (depStatusTimer) { clearTimeout(depStatusTimer); depStatusTimer = null; }
      });
      el.addEventListener('mouseleave', function () {
        hideDepStatus(800);          // 移开后短延迟再消失，避免误触抖动
      });
    })();

    // ============== 发送按钮（不同 strategy 决定 iframe 定位方式，挑出能用的）==============
    function getDepCells() {
      return {
        B4: ($('depInputB4') || {}).value || '',
        C4: ($('depInputC4') || {}).value.trim() || '',
        D4: ($('depInputD4') || {}).value || '',
        E4: ($('depInputE4') || {}).value || '',
        F4: ($('depInputF4') || {}).value || ''
      };
    }
    // ============== 发送模式（下方 7 个按钮：点一下切换当前模式，默认 ⑦）==============
    // strategy 内部标识 → 状态提示里显示的友好名（与下方按钮文字一致）
    var STRATEGY_LABEL = {
      fs_tab_toolbar: '报表Toolbar',
      fs_tab_id: 'id^fs_tab',
      first_iframe: '首个iframe',
      fs_tab_class: 'item类',
      name_fs_tab: 'name^fs_tab',
      self: '自身',
      all: '兜底'
    };
    var depSendBtns = document.querySelectorAll('.dep-send-btns .btn-send');
    var depLastTrack = '';   // 最近一次发送所用的股道（道号），用于回执提示
    var depActiveStrategy = 'fs_tab_toolbar';
    function setActiveStrategy(btn) {
      for (var k = 0; k < depSendBtns.length; k++) depSendBtns[k].classList.remove('active');
      btn.classList.add('active');
      depActiveStrategy = btn.getAttribute('data-strategy') || 'all';
    }
    for (var i = 0; i < depSendBtns.length; i++) {
      (function (btn) {
        btn.addEventListener('click', function () { setActiveStrategy(btn); });
      })(depSendBtns[i]);
    }
    // 同步 HTML 里默认带 active 的那个（⑦ fs_tab_toolbar）
    (function () {
      var def = document.querySelector('.dep-send-btns .btn-send.active');
      if (def) depActiveStrategy = def.getAttribute('data-strategy') || 'fs_tab_toolbar';
    })();

    // ============== 右侧「发送」按钮：用当前选中的模式发送 ==============
    on('btnSendCurrent', 'click', function () {
      var cells = getDepCells();
      if (!cells.C4) { showDepStatus('请先输入车次', 'warn'); return; }
      depLastTrack = cells.B4;
      try {
        window.postMessage({
          channel: DEP_BRIDGE,
          type: 'fillByStrategy',
          cells: cells,
          strategy: depActiveStrategy,
          ts: Date.now()
        }, '*');
        showDepStatus('已下发填报指令 · 模式：' + (STRATEGY_LABEL[depActiveStrategy] || depActiveStrategy) + ' · 车次：' + cells.C4, 'ok');
      } catch (e) {
        showDepStatus('发送失败：' + e.message, 'warn');
      }
    });


    // 「刷新报表页」：向扩展桥下刷新指令，由 background 刷新「popup 报表地址」匹配的标签页
    on('btnReloadReport', 'click', function () {
      try {
        window.postMessage({ channel: DEP_BRIDGE, type: 'reloadReport', ts: Date.now() }, '*');
        showDepStatus('已发送刷新指令…', 'info');
      } catch (e) {
        showDepStatus('刷新指令发送失败：' + e.message, 'warn');
      }
    });

    // 监听扩展桥回执（filled）—— 哪个 strategy 写成功 / 失败 / 原因都在这里能直接看到
    window.addEventListener('message', function (ev) {
      var d = ev.data;
      if (!d || d.channel !== DEP_BRIDGE) return;
      // 刷新指令回执
      if (d.type === 'reloadResult') {
        showDepStatus(
          (d.count || 0) > 0
            ? ('已刷新 ' + d.count + ' 个报表页')
            : '未找到报表页：请检查 popup 里的「报表地址」是否与已打开的页面一致',
          (d.count || 0) > 0 ? 'ok' : 'warn'
        );
        return;
      }
      if (d.type === 'filled' && d.strategy) {
        var ok = d.ok || 0, total = 5;
        var track = depLastTrack ? depLastTrack + '道' : '';
        var msg;
        if (ok >= total)      msg = track + '编好';
        else if (ok > 0)      msg = track + '编好（部分 ' + ok + '/' + total + '）';
        else                  msg = track + '编好失败';
        if (d.error)  msg += ' · 原因：' + d.error;
        if (d.failed && d.failed.length) msg += ' · 未写入：' + d.failed.join('、');
        showDepStatus(msg, ok > 0 ? 'ok' : 'warn');
      }
    });

    // 自适应列宽（重置为内容自适应，清除手动拖动记忆）
    on('btnAutoFitCols', 'click', function () {
      var g = $('grid'), dt = $('detailTable');
      ColResize.safeGet(g).reset();
      ColResize.safeGet(dt).reset();
      var thA = g && g.querySelector('thead th.col-a');
      if (g && thA) g.style.setProperty('--col-a-w', thA.offsetWidth + 'px');
      toast('列宽已重置为自适应', 'ok');
    });

    on('btnPrevTrack', 'click', function () { stepDetail(-1); });
    on('btnNextTrack', 'click', function () { stepDetail(1); });

    /* ---- 浮窗 / 抽屉 / 下拉菜单 -----
     * 注册后自动获得「点空白处关闭 + ESC 关栈顶」，新增面板无需再改 ESC 处理。 */
    UI.Modal.register('modal31814');
    UI.Modal.register('modalSettings', { onOpen: refreshFolderPath });
    UI.Modal.register('modalProductivity');
    UI.Modal.register('modalDeparture');
    UI.Drawer.register('drawer', { maskId: 'drawerMask' });
    UI.Drawer.register('searchDrawer', { maskId: 'searchMask' });

    // 各面板的关闭按钮
    on('btnCloseDrawer', 'click', closeDetail);
    on('rptClose', 'click', function () { UI.Modal.close('modal31814'); });
    on('btnSettingsClose', 'click', function () { UI.Modal.close('modalSettings'); });

    /* ---- 配置备份：导出 / 导入本地持久化（localStorage）---- */
    function tsStamp() {
      var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; };
      return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' +
             p(d.getHours()) + p(d.getMinutes());
    }
    function exportConfig() {
      var map = Store.allSync();
      var json = ConfigIO.toJson(map, {
        app: '站存计算器',
        exportedAt: new Date().toISOString(),
        note: '仅含浏览器 localStorage 持久化项；数据文件夹句柄需重新选择'
      });
      var blob = new Blob([json], { type: 'application/json;charset=utf-8' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url; a.download = 'zhancun-config-' + tsStamp() + '.json';
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      toast('配置已导出（' + Object.keys(map).length + ' 项）', 'ok');
    }
    on('btnExportConfig', 'click', exportConfig);
    on('btnImportConfig', 'click', function () { var f = $('importConfigFile'); if (f) f.click(); });
    on('importConfigFile', 'change', function (e) {
      var file = e.target.files && e.target.files[0];
      if (!file) return;
      var reader = new FileReader();
      reader.onload = function () {
        try {
          var map = ConfigIO.parseJson(String(reader.result));
          var keys = Object.keys(map);
          if (!keys.length) { toast('文件为空或不是有效的配置文件', 'error'); return; }
          if (!confirm('导入将覆盖当前本地配置（共 ' + keys.length + ' 项），确定继续？')) return;
          var n = Store.applySync(map);
          toast('已导入 ' + n + ' 项配置，即将刷新页面生效', 'ok');
          setTimeout(function () { location.reload(); }, 700);
        } catch (err) {
          toast('导入失败：' + (err && err.message ? err.message : err), 'error');
        } finally {
          e.target.value = '';
        }
      };
      reader.onerror = function () { toast('读取文件失败', 'error'); e.target.value = ''; };
      reader.readAsText(file);
    });
    on('prodClose', 'click', function () { UI.Modal.close('modalProductivity'); });
    on('depClose', 'click', function () { UI.Modal.close('modalDeparture'); });
    // 原工具栏「发车流程」按钮（btnDepartureFlow）已移除：
    // 改由主表「空箱/空车」列（表头 th[9]）的「编好」伪元素按钮打开同一浮窗。

    // 功能下拉菜单（组件负责展开 / 收起 / 点外部关闭）
    UI.Dropdown('btnMenu', 'menuList', {
      onSelect: function (item, action) {
        if (action === '31814' && typeof window.Report31814 !== 'undefined') {
          window.Report31814.open(global.YardApp.getRawRows(), state.currentFile);
        } else if (action === 'settings') {
          UI.Modal.open('modalSettings');
        } else if (action === 'productivity' && typeof window.Productivity !== 'undefined') {
          window.Productivity.open();
        }
      }
    });

    // 设置：表格字号滑块（主页表格 grid + 明细抽屉表格 detailTable 同步生效）
    var gridFontSize = $('gridFontSize');
    var gridFontSizeVal = $('gridFontSizeVal');
    function applyTableFontSize(v) {
      $('grid').style.fontSize = v + 'px';
      var detailTable = $('detailTable');
      if (detailTable) {
        detailTable.style.fontSize = v + 'px';   // 覆盖 CSS 中 table.detail 的固定字号
      }
    }
    if (gridFontSize && gridFontSizeVal) {
      var savedFs = Store.get('gridFontSize', '');
      if (savedFs) {
        gridFontSize.value = savedFs;
        gridFontSizeVal.textContent = savedFs + 'px';
        applyTableFontSize(savedFs);
      }
      function highlightTick(v) {
        var ticks = document.querySelectorAll('.range-ticks .tick');
        for (var i = 0; i < ticks.length; i++) {
          ticks[i].classList.toggle('active', ticks[i].textContent === v);
        }
      }
      // 拖动滑块会高频触发 input，写存储走防抖，避免每次都落盘 localStorage
      var saveFontSize = Utils.debounce(function (v) {
        Store.set('gridFontSize', v);
      }, 300);

      gridFontSize.addEventListener('input', function () {
        var v = gridFontSize.value;
        gridFontSizeVal.textContent = v + 'px';
        applyTableFontSize(v);
        saveFontSize(v);
        highlightTick(v);
      });
      highlightTick(gridFontSize.value);
    }

    // 设置：货物推算重量（载重缺失时按记事栏货物名称匹配预设重量；列表可增删）
    initGoodsWeightConfig('estLoadGoods', 'btnAddEstLoadGood', Store.KEYS.estLoadGoods, DEFAULT_EST_GOODS);

    // 设置：默认文件夹
    var folderPath = $('folderPath');
    var btnSettingFolder = $('btnSettingFolder');
    function refreshFolderPath() {
      if (folderPath) {
        var name = Store.get('folderName', '');
        folderPath.textContent = name || '未设置';
        folderPath.title = name || '';
      }
    }
    refreshFolderPath();
    if (btnSettingFolder) {
      btnSettingFolder.addEventListener('click', function () {
        DataSource.pickFolder();
      });
    }

    // 设置：可配置「标签列表」（卸车地点 / 黑罐识别共用同一套交互）
    // 改完任一配置后，用新配置重跑聚合并刷新主表（recompute 是共享的）
    function recompute() {
      if (!state.rawRows || !state.rawRows.length) { render(); return; }
      var base = state.printDate || new Date();
      var agg = Aggregate.aggregate(state.rawRows, state.dirIndex.map,
                                    state.dirIndex.stations, YardConfig.thresholds, base);
      var ordered = [], known = {}, extra = [];
      YardConfig.tracks.forEach(function (t) {
        if (agg[t.id]) ordered.push(agg[t.id]);
        else ordered.push({ track: t.id, direction: '', count: 0, carTypes: '',
                            length: 0, dest: '', train: '', load: 0, oldCar: 0, raw: [] });
        known[t.id] = 1;
      });
      Object.keys(agg).forEach(function (k) { if (!known[k]) extra.push(agg[k]); });
      extra.sort(function (a, b) { return a.track.localeCompare(b.track, 'zh'); });
      state.rows = ordered.concat(extra);
      render();
    }

    /**
     * 初始化一个「可增删标签列表」设置项（行内输入交互）。
     * @param containerId 列表容器 id（内含地点项 + 新增按钮）
     * @param addBtnId   新增按钮 id
     * @param storeKey   持久化键（Store.get/set）
     * @param defaultList 未配置时的默认值
     */
    function initSpotConfig(containerId, addBtnId, storeKey, defaultList) {
      var el = $(containerId);
      var addBtn = $(addBtnId);
      if (!el) return;
      var adding = false;

      function getList() {
        var list = Store.get(storeKey, null);
        return Array.isArray(list) ? list : null;   // null → 用默认
      }
      function renderList() {
        var list = getList() || defaultList;
        var html = list.map(function (name) {
          return '<span class="spot-item" data-name="' + escapeHtml(name) + '">' +
                   escapeHtml(name) +
                   '<button class="spot-del" title="删除" aria-label="删除">×</button>' +
                 '</span>';
        }).join('');
        if (adding) {
          html += '<span class="spot-item spot-editing">' +
                    '<input class="spot-input" type="text" maxlength="20" ' +
                    'placeholder="输入名称" autocomplete="off" />' +
                  '</span>';
        }
        el.innerHTML = html;
        if (addBtn) el.appendChild(addBtn);   // 按钮每次重挂回末尾
        if (adding) {
          var inp = el.querySelector('.spot-input');
          if (inp) inp.focus();
        }
      }
      function commit() {
        var inp = el.querySelector('.spot-input');
        var name = inp ? inp.value.trim() : '';
        adding = false;
        if (!name) { renderList(); return; }
        var list = getList() || defaultList;
        if (list.indexOf(name) >= 0) { UI.toast(name + ' 已存在'); renderList(); return; }
        list.push(name);
        Store.set(storeKey, list);
        renderList();
        recompute();
      }
      function cancel() { if (!adding) return; adding = false; renderList(); }

      renderList();

      el.addEventListener('click', function (e) {
        var del = e.target.closest ? e.target.closest('.spot-del') : null;
        if (!del) return;
        var name = del.parentNode.getAttribute('data-name');
        var list = getList() || defaultList;
        Store.set(storeKey, list.filter(function (n) { return n !== name; }));
        renderList();
        recompute();
      });
      el.addEventListener('keydown', function (e) {
        if (!e.target.classList || !e.target.classList.contains('spot-input')) return;
        if (e.key === 'Enter') { e.preventDefault(); commit(); }
        else if (e.key === 'Escape') { e.preventDefault(); cancel(); }
      });
      el.addEventListener('focusout', function (e) {
        if (e.target.classList && e.target.classList.contains('spot-input')) commit();
      });
      if (addBtn) {
        addBtn.addEventListener('click', function () {
          if (adding) { var i = el.querySelector('.spot-input'); if (i) i.focus(); return; }
          adding = true;
          renderList();
        });
      }
    }

    /**
     * 初始化「货物推算重量」配置项：列表项形如 汽油:60，单击×删除，双击打开弹窗修改；
     * 新增按钮同样打开弹窗。弹窗浮于设置窗口内（#gwModal），确认后写入 Store。
     * 载重缺失时记事栏命中名称即取该重量；值是 { name, weight } 对象而非纯字符串。
     */
    function initGoodsWeightConfig(containerId, addBtnId, storeKey, defaultList) {
      var el = $(containerId);
      var addBtn = $(addBtnId);
      var modal = $('gwModal');
      var nameInp = $('gwName');
      var wInp = $('gwWeight');
      var gwOk = $('gwOk');
      var hint = $('gwHint');
      var mergeArmed = false;   // 编辑改名命中已有项时的二次确认态
      if (!el || !modal) return;
      var editName = null;   // 正在编辑的货物名称（null = 新增）

      function getList() {
        var list = Store.get(storeKey, null);
        return Array.isArray(list) ? list : null;   // null → 用默认
      }
      function renderList() {
        var list = getList() || defaultList;
        el.innerHTML = list.map(function (item) {
          return '<span class="spot-item" data-name="' + escapeHtml(item.name) + '" title="双击修改">' +
                   escapeHtml(item.name) + ':' + escapeHtml(item.weight) +
                   '<button class="spot-del" title="删除" aria-label="删除">×</button>' +
                 '</span>';
        }).join('');
        if (addBtn) el.appendChild(addBtn);   // 新增按钮保持在末尾
      }
      function setHint(txt, cls) {
        if (!hint) return;
        hint.textContent = txt || '';
        hint.className = 'gw-hint' + (cls ? ' ' + cls : '');
        // 重启淡入动画，使每次提示都"冒泡"出现
        hint.style.animation = 'none';
        void hint.offsetWidth;
        hint.style.animation = '';
      }
      function updateHint() {
        if (mergeArmed) return;   // 武装态由 commit 设置冲突提示，不覆盖
        var n = nameInp.value.trim();
        var w = Utils.vbVal(wInp.value);
        if (!n && !(w > 0)) { setHint(''); return; }
        if (!n) { setHint('请填写货物名称', 'warn'); return; }
        if (!(w > 0)) { setHint('请填写重量（>0）', 'warn'); return; }
        setHint(n + ' · ' + (Math.round(w * 10) / 10) + ' 吨');
      }
      function resetOk() {
        mergeArmed = false;
        if (gwOk) { gwOk.textContent = '确定'; gwOk.classList.remove('merge'); }
      }
      function openModal(name) {
        editName = (name == null ? null : name);
        var cur = (getList() || defaultList).filter(function (it) { return it.name === name; })[0];
        nameInp.value = name != null ? name : '';
        wInp.value = (cur && name != null) ? cur.weight : '';
        modal.hidden = false;
        nameInp.focus();
        resetOk();
        updateHint();
      }
      function finish(list) {
        Store.set(storeKey, list);
        renderList();
        closeModal();
        if (state.detailIdx != null) updateDetailTitle();   // 刷新明细计重
      }
      function closeModal() {
        modal.hidden = true;
        editName = null;
        resetOk();
        setHint('');
      }
      function commit() {
        var name = nameInp.value.trim();
        var weight = Utils.vbVal(wInp.value);
        if (!name) { setHint('名称不能为空', 'warn'); return; }
        if (!(weight > 0)) { setHint('重量需大于 0', 'warn'); return; }
        var list = (getList() || defaultList).slice();
        var rounded = Math.round(weight * 10) / 10;
        if (editName == null) {
          // 新增：不允许与已有同名
          if (list.some(function (it) { return it.name === name; })) { setHint(name + ' 已存在', 'warn'); return; }
          list.push({ name: name, weight: rounded });
          finish(list);
        } else {
          var idx = -1, same = -1;
          for (var i = 0; i < list.length; i++) {
            if (list[i].name === editName) idx = i;
            if (list[i].name === name) same = i;
          }
          if (same >= 0 && same !== idx) {
            // 编辑改名命中已有项 → 二次确认：确定按钮变「合并」，点它才合并
            if (!mergeArmed) {
              setHint('「' + name + '」已存在，点“合并”将覆盖其重量并移除「' + editName + '」', 'warn');
              mergeArmed = true;
              if (gwOk) { gwOk.textContent = '合并'; gwOk.classList.add('merge'); }
              return;
            }
            list[same].weight = rounded;
            list.splice(idx, 1);
            finish(list);
          } else if (idx < 0) {
            // 旧项已不在（并发删除等情况），当作 upsert 处理
            if (same >= 0) { list[same].weight = rounded; }
            else { list.push({ name: name, weight: rounded }); }
            finish(list);
          } else {
            // 改名或仅改重（同名命中自己也在此分支）
            list[idx].name = name;
            list[idx].weight = rounded;
            finish(list);
          }
        }
      }

      renderList();

      // 双击标签项 → 打开弹窗并预填（修改内容）；双击×不触发
      el.addEventListener('dblclick', function (e) {
        if (e.target.closest && e.target.closest('.spot-del')) return;
        var item = e.target.closest ? e.target.closest('.spot-item') : null;
        if (item) openModal(item.getAttribute('data-name'));
      });
      // 单击删除按钮
      el.addEventListener('click', function (e) {
        var del = e.target.closest ? e.target.closest('.spot-del') : null;
        if (!del) return;
        var name = del.parentNode.getAttribute('data-name');
        var list = (getList() || defaultList).filter(function (it) { return it.name !== name; });
        Store.set(storeKey, list);
        renderList();
        if (state.detailIdx != null) updateDetailTitle();
      });
      if (addBtn) addBtn.addEventListener('click', function () { openModal(null); });
      $('gwClose').addEventListener('click', closeModal);
      $('gwCancel').addEventListener('click', closeModal);
      $('gwOk').addEventListener('click', commit);
      // 输入变化 → 刷新左侧预览，并解除「合并」武装态（恢复确定按钮）
      nameInp.addEventListener('input', function () { resetOk(); updateHint(); });
      wInp.addEventListener('input', function () { resetOk(); updateHint(); });
      modal.querySelector('.gw-modal-mask').addEventListener('click', closeModal);
      nameInp.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); commit(); }
        else if (e.key === 'Escape') { e.preventDefault(); closeModal(); }
      });
      wInp.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); commit(); }
        else if (e.key === 'Escape') { e.preventDefault(); closeModal(); }
      });
    }

    // 车型高亮：基于 Utils.getCarTypeConfig 在设置面板渲染可编辑列表
    function initCarTypeConfig() {
      if (!Utils.applyCarTypeStyles) return;
      Utils.applyCarTypeStyles(); // 首屏即应用已存配置

      var box = $('carTypeCfg');
      if (!box) return;

      /* 关键 —— 下面会再定义一个同名 render() 用于配置列表自身，
       * JS 作用域会就近解析到内层那个，导致「车型高亮变更后重渲染主表」
       * 静默失效。这里先在外层把主表 render 捕获成别名，persist / 重置
       * 按钮里用这个别名调主表渲染。 */
      var renderGrid = render;

      function colorOptions(sel) {
        return (Utils.CAR_COLORS || []).map(function (c) {
          return '<option value="' + escapeHtml(c.value) + '"' +
            (c.value.toLowerCase() === String(sel).toLowerCase() ? ' selected' : '') + '>' +
            escapeHtml(c.name) + '</option>';
        }).join('');
      }

      function persist(cfg) {
        if (Store && Store.set) Store.set('carTypeStyle', cfg);
        Utils.applyCarTypeStyles();
        if (state.detailIdx != null) openDetail(state.detailIdx); // 实时刷新明细高亮
        // 主表到站列的车型标记同样受该配置控制，否则会「明细变了主表没变」。
        // 必须用外层捕获的 renderGrid：闭包内的 render 是配置列表自身渲染。
        if (state.rows && state.rows.length) renderGrid();
      }

      function render() {
        var cfg = Utils.getCarTypeConfig();
        /* 行一律塞进独立容器 #carTypeRows。
         * 早先写法是 box.insertBefore(row, $('btnAddCarType'))，隐含假设
         * 「新增按钮是 .cartype-cfg 的直接子元素」。一旦给按钮加了包裹层
         * （如 .cartype-cfg-actions），insertBefore 会抛 NotFoundError，
         * 中断整个 render → 后续按钮事件绑定全部跳过 → 按钮点击无反应。
         * 改用独立容器后，按钮怎么摆都不影响渲染。 */
        var rowsBox = $('carTypeRows') || box;   // 找不到容器时退回 box，保证不崩
        rowsBox.innerHTML = '';                  // 清空旧行（按钮在容器外，不受影响）
        cfg.forEach(function (e, idx) {
          var row = document.createElement('div');
          row.className = 'cartype-cfg-row';
          row.setAttribute('data-idx', idx);
          row.innerHTML =
            '<input type="text" class="ct-prefix" value="' + escapeHtml(e.prefix) + '" placeholder="如 DK">' +
            '<input type="text" class="ct-note" value="' + escapeHtml(e.note) + '" placeholder="备注">' +
            '<select class="ct-match">' +
              '<option value="starts"' + (e.match !== 'contains' ? ' selected' : '') + '>开头</option>' +
              '<option value="contains"' + (e.match === 'contains' ? ' selected' : '') + '>包含</option>' +
            '</select>' +
            '<input type="checkbox" class="ct-on" ' + (e.on ? 'checked' : '') + ' title="启用">' +
            '<input type="checkbox" class="ct-bold" ' + (e.bold ? 'checked' : '') + ' title="加粗">' +
            '<select class="ct-color">' + colorOptions(e.color) + '</select>' +
            '<button type="button" class="ct-del" title="删除">×</button>';
          rowsBox.appendChild(row);

          var prefixEl = row.querySelector('.ct-prefix');
          var noteEl = row.querySelector('.ct-note');
          var matchEl = row.querySelector('.ct-match');
          var onEl = row.querySelector('.ct-on');
          var boldEl = row.querySelector('.ct-bold');
          var colorEl = row.querySelector('.ct-color');
          var delEl = row.querySelector('.ct-del');

          function update() {
            var cur = Utils.getCarTypeConfig();
            cur[idx] = {
              prefix: prefixEl.value.trim(),
              note: noteEl.value.trim(),
              match: matchEl.value,
              on: onEl.checked,
              bold: boldEl.checked,
              color: colorEl.value
            };
            persist(cur);
          }
          [prefixEl, noteEl, matchEl, onEl, boldEl, colorEl].forEach(function (el) {
            el.addEventListener('input', update);
            el.addEventListener('change', update);
          });
          delEl.addEventListener('click', function () {
            var cur = Utils.getCarTypeConfig();
            cur.splice(idx, 1);
            persist(cur);
            render();
          });
        });
      }

      render();

      var addBtn = $('btnAddCarType');
      if (addBtn) {
        addBtn.addEventListener('click', function () {
          var cur = Utils.getCarTypeConfig();
          cur.push({ prefix: '', note: '', match: 'starts', on: true, bold: true, color: '#e53e3e' });
          persist(cur);
          render();
        });
      }

      /* 表头「颜色」后的收起 / 展开按钮：只切换数据行容器的显示，
       * 状态存 Store，下次打开设置面板沿用。 */
      var toggleBtn = $('btnCarTypeToggle');
      function applyCarTypeCollapse(collapsed) {
        box.classList.toggle('collapsed', !!collapsed);
        if (toggleBtn) {
          toggleBtn.textContent = collapsed ? '展开' : '收起';
          toggleBtn.title = collapsed ? '展开配置表' : '收起配置表';
          toggleBtn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
        }
        Store.set('carTypeCfgCollapsed', !!collapsed);
      }
      applyCarTypeCollapse(Store.get('carTypeCfgCollapsed', false) === true);
      if (toggleBtn) {
        toggleBtn.addEventListener('click', function () {
          applyCarTypeCollapse(!box.classList.contains('collapsed'));
        });
      }

      /* 「恢复默认」按钮：清掉用户之前保存的自定义车型高亮配置，
       * 回到 defaultCarTypeConfig（包含 NX/X 双 starts 项，平板车全部命中）。
       * 浏览器里的 Store 是独立的，改默认配置对老用户无效——此按钮是面向
       * 「设置被改乱了想一键还原」场景的快捷入口。
       *
       * 选用 location.reload() 而不是即时重渲染的原因：闭包内的 render / 外部
       * renderGrid 在初始化时序差异下可能踩到作用域陷阱（之前 persist 那段
       * 就静默失败了），reload 走整个 IIFE 重新初始化，所有路径都按 default
       * 走，最简单可靠。 */
      var resetBtn = $('btnResetCarType');
      if (resetBtn) {
        resetBtn.addEventListener('click', function () {
          try {
            if (Store && Store.remove) Store.remove('carTypeStyle');
            toast('已恢复默认车型高亮配置');
          } catch (e) {
            console.error('[btnResetCarType] 清', e);
            toast('恢复默认失败：' + (e && e.message || e), 'error');
            return;
          }
          // 给 toast 一点可见时间再刷新
          setTimeout(function () { location.reload(); }, 250);
        });
      } else {
        console.warn('[btnResetCarType] #btnResetCarType 在 DOM 中未找到，按钮事件未绑定');
      }
    }

    // 卸车地点（aggregate 段2 读取）
    initSpotConfig('unloadSpots', 'btnAddUnloadSpot', 'unloadSpots', DEFAULT_UNLOAD_SPOTS);
    // 黑罐识别（aggregate 段1 G7 罐按收货人识别）
    initSpotConfig('blackTankSpots', 'btnAddBlackTankSpot', 'blackTankSpots', DEFAULT_BLACK_TANK_SPOTS);

    // 车型高亮配置（对齐 VBA 显示信息.bas）
    initCarTypeConfig();

    // 表头列宽拖动（persistKey 用于本地记忆，刷新/重渲染不丢失）
    ColResize.enable($('grid'), {
      persistKey: 'zhancun.grid.cols.v2',
      onResize: function (th, w) {
        // 首列 / 分组合并列宽变化 → 同步后续冻结列的偏移，避免列间露缝
        if (th.classList.contains('col-a')) $('grid').style.setProperty('--col-a-w', w + 'px');
        if (th.classList.contains('col-b-group')) $('grid').style.setProperty('--col-b-group-w', w + 'px');
      }
    });
    /* 列宽按「列序索引」记忆，因此调整 DETAIL_COLS 的顺序或增删列后，
     * 旧记忆会整体错位（宽度套到了别的列上）。键名带版本号即可让旧记忆失效、
     * 首次打开重新按内容自适应——改动列序时，把 v1 递增即可。 */
    ColResize.enable($('detailTable'), { persistKey: Store.KEYS.detailCols });
    /* 搜索结果表与明细表列数、列序完全一致（同一份 DETAIL_COLS + 停时列），
     * 故共用同一个记忆键：在任一表中拖好的列宽，另一个表打开时自动沿用。 */
    ColResize.enable($('searchTable'), { persistKey: Store.KEYS.detailCols });
  }

  /* =================== 空框架 =================== */
  /** 按配置清单渲染空行，数据未加载时也能看到完整布局 */
  function renderEmpty() {
    state.rows = YardConfig.tracks.map(function (t) {
      return {
        track: t.id, direction: '', count: 0, carTypes: '', length: 0,
        dest: '', train: '', note: '', load: 0, oldCar: 0, raw: []
      };
    });
    state.currentFile = '未加载数据';
    render();
  }

  /* =================== 浮窗拖动 =================== */
  // 在标题栏(.modal-head)按下可拖动浮窗移动位置；按钮/输入框不触发拖动
  function initModalDrag() {
    document.addEventListener('pointerdown', function (e) {
      var handle = e.target.closest ? e.target.closest('.modal-head') : null;
      if (!handle) return;
      if (e.target.closest('button, input, select, a, .col-handle')) return;
      var modal = handle.closest('.modal');
      if (!modal || !modal.classList.contains('show')) return;
      var content = modal.querySelector('.modal-content');
      if (!content) return;

      var rect = content.getBoundingClientRect();
      var startX = e.clientX, startY = e.clientY;
      var origLeft = rect.left, origTop = rect.top;
      content.style.position = 'fixed';
      content.style.margin = '0';
      content.style.left = origLeft + 'px';
      content.style.top = origTop + 'px';
      content.style.zIndex = '200';
      document.body.style.userSelect = 'none';
      e.preventDefault();

      function move(ev) {
        var nx = origLeft + (ev.clientX - startX);
        var ny = origTop + (ev.clientY - startY);
        nx = Math.max(0, Math.min(nx, window.innerWidth - content.offsetWidth));
        ny = Math.max(0, Math.min(ny, window.innerHeight - content.offsetHeight));
        content.style.left = nx + 'px';
        content.style.top = ny + 'px';
      }
      function up() {
        document.removeEventListener('pointermove', move);
        document.removeEventListener('pointerup', up);
        document.body.style.userSelect = '';
      }
      document.addEventListener('pointermove', move);
      document.addEventListener('pointerup', up);
    });
  }

  /* =================== 浮窗调整大小 =================== */
  // 拖 .dep-resize-handle（右下角）改变浮窗宽高；尺寸记忆到 localStorage，下次打开沿用。
  var DEP_SIZE_KEY = 'depModalSize';
  function applyDepSize() {
    // 尺寸完全由 CSS（.dep-modal）控制，不再用内联样式覆盖。
    // 仅确保默认不残留旧的内联 max 限制即可（拖动时的内联尺寸由拖动逻辑管理）。
    var c = document.querySelector('#modalDeparture .modal-content');
    if (!c) return;
    c.style.maxWidth = '';
    c.style.maxHeight = '';
  }
  function initModalResize() {
    applyDepSize();
    // 双击手柄 → 恢复默认尺寸
    document.addEventListener('dblclick', function (e) {
      var h = e.target.closest ? e.target.closest('.dep-resize-handle') : null;
      if (!h) return;
      var c = h.closest('.modal-content');
      if (!c) return;
      Store.remove(DEP_SIZE_KEY);
      c.style.width = ''; c.style.height = '';
      c.style.maxWidth = ''; c.style.maxHeight = '';
    });
    document.addEventListener('pointerdown', function (e) {
      var handle = e.target.closest ? e.target.closest('.dep-resize-handle') : null;
      if (!handle) return;
      var modal = handle.closest('.modal');
      if (!modal || !modal.classList.contains('show')) return;
      var content = handle.closest('.modal-content');
      if (!content) return;

      var rect = content.getBoundingClientRect();
      var startX = e.clientX, startY = e.clientY;
      var w0 = rect.width, h0 = rect.height;
      var cs = getComputedStyle(content);
      var minW = parseFloat(cs.minWidth) || 320;
      var minH = parseFloat(cs.minHeight) || 200;
      // 允许超过 CSS 里的默认上限（50vh / 96vw）
      content.style.maxWidth = 'none';
      content.style.maxHeight = 'none';
      document.body.style.userSelect = 'none';
      e.preventDefault();
      // 指针捕获：拖动经过内部 iframe 时事件仍回到手柄，不会丢失
      if (e.pointerId != null && handle.setPointerCapture) {
        try { handle.setPointerCapture(e.pointerId); } catch (err) {}
      }

      function move(ev) {
        var maxW = Math.max(minW, window.innerWidth - rect.left - 8);
        var maxH = Math.max(minH, window.innerHeight - rect.top - 8);
        var w = Math.max(minW, Math.min(w0 + (ev.clientX - startX), maxW));
        var h = Math.max(minH, Math.min(h0 + (ev.clientY - startY), maxH));
        content.style.width = w + 'px';
        content.style.height = h + 'px';
      }
      function up() {
        document.removeEventListener('pointermove', move);
        document.removeEventListener('pointerup', up);
        document.body.style.userSelect = '';
        Store.set(DEP_SIZE_KEY, {
          w: Math.round(content.offsetWidth),
          h: Math.round(content.offsetHeight)
        });
      }
      document.addEventListener('pointermove', move);
      document.addEventListener('pointerup', up);
    });
  }

  /* =================== 初始化 =================== */
  function init() {
    // 桥接：把私有共享物挂到 global，供已抽出的 data-source.js 通过全局名访问
    // （state 为 IIFE 私有对象，data-source 仅读取/修改其属性，不重赋值，故引用稳定）
    global.state = state;
    global.render = render;
    global.loading = loading;
    global.syncPickFolderBtn = syncPickFolderBtn;

    initModalDrag();
    initModalResize();
    // 方向库（惰性单例：全局只解析一次，报表模块共用同一份实例）
    if (typeof window.DirectionData !== 'string') {
      toast('方向库未加载，到站着色将不可用', 'error');
    }
    state.dirIndex = Aggregate.getDirectionIndex();

    bind();

    // 恢复「隐藏空线分组」的持久记忆（设置里切换时写入 Store.showEmptyGroups）
    var savedEmpty = Store.get('showEmptyGroups', null);
    if (savedEmpty !== null) state.showEmptyGroups = !!savedEmpty;

    // 恢复「注意事项收起」的持久记忆（表头点击切换时写入 Store.notesCollapsed）
    var savedNotes = Store.get('notesCollapsed', null);
    if (savedNotes === true && $('grid')) $('grid').classList.add('notes-collapsed');

    syncVirtualBtn();

    // 先渲染空框架：让页面一打开就呈现完整股道清单，便于核对配置
    renderEmpty();

    // 护眼色滑块
    (function () {
      var slider = $('eyeProtectSlider');
      var overlay = $('eyeOverlay');
      if (!slider || !overlay) return;
      var saved = Store.get('eyeProtect', 0);
      slider.value = saved;
      overlay.style.opacity = saved / 100;
      if (saved > 0) overlay.style.display = '';
      on('eyeProtectSlider', 'input', function () {
        var v = parseInt(this.value, 10);
        overlay.style.display = v > 0 ? '' : 'none';
        overlay.style.opacity = v / 100;
        Store.set('eyeProtect', v);
      });
    })();

    // 恢复上次选择的文件夹，自动读取最新 xls
    if (!window.showDirectoryPicker) {
      $('stMsg').textContent = '当前浏览器不支持自动读取，请点「选择数据文件夹」手动选取文件';
      syncPickFolderBtn(true);   // 该分支下点击会降级为文件选择
      return;
    }

    Store.async.get('xlsDir').then(function (h) {
      if (!h) {
        $('stMsg').textContent = '首次使用：请点「选择数据文件夹」，之后将自动读取最新文件';
        syncPickFolderBtn(true);   // 显示入口，让用户一眼就能找到
        return;
      }
      syncPickFolderBtn(false);  // 已记住文件夹，按钮收起
      state.dirHandle = h;
      DataSource.rememberFolder(h);
      // 静默恢复：权限未授予时不弹窗，等用户点击
      if (h.queryPermission) {
        return h.queryPermission({ mode: 'read' }).then(function (p) {
          if (p === 'granted') return DataSource.loadFromDir(h, false);
          $('stMsg').textContent = '已记住数据文件夹，点「重新读取」以载入';
        });
      }
      return DataSource.loadFromDir(h, false);
    }).catch(function () {
      $('stMsg').textContent = '读取上次的文件夹失败，请点「选择数据文件夹」重新选择';
      syncPickFolderBtn(true);   // 句柄不可用，重新亮出入口
    });

  }

  /* =================== 对外接口（window.YardApp） ===================
   * 统一出口：后续新增模块（如 search.js）都从这里取数 / 复用渲染 / 触发交互，
   * 不必各自再实现一套遍历或表格渲染，避免重复造轮子和格式漂移。
   *
   * 约定：
   *   · 只暴露「只读数据 + 受控操作」，不把 state 整个开放出去；
   *   · 需要新增能力时在此追加，调用方一律通过 YardApp 访问，不直接碰内部函数。
   * ===================================================================== */

  /** 数据变化订阅者：主表每次重绘后统一回调（加载 / 刷新 / 切换空线分组都会触发） */
  var dataListeners = [];

  /** 注册数据变化回调，返回反注册函数 */
  function onDataChange(fn) {
    if (typeof fn !== 'function') return function () {};
    dataListeners.push(fn);
    return function () {
      var i = dataListeners.indexOf(fn);
      if (i >= 0) dataListeners.splice(i, 1);
    };
  }

  function notifyDataChange() {
    if (!dataListeners) return;          // 理论上不会发生（init 在本文件最后调用）
    for (var i = 0; i < dataListeners.length; i++) {
      // 单个订阅者出错不能拖垮主流程与其它模块
      try { dataListeners[i](); } catch (e) { if (global.console) console.error(e); }
    }
  }

  /** 合计快照（口径与主表状态栏一致） */
  function getSummary() {
    var s = { track: 0, count: 0, length: 0, load: 0, oldCar: 0 };
    (state.rows || []).forEach(function (r) {
      s.track++;
      s.count += r.count || 0;
      s.length += r.length || 0;
      s.load += r.load || 0;
      s.oldCar += r.oldCar || 0;
    });
    return s;
  }

  global.YardApp = {
    /* ---------- 数据 ---------- */
    /** 全部股道聚合数据：每项 { track, count, length, load, oldCar, raw:[车辆行...], ... } */
    getRows: function () { return state.rows || []; },
    /** 原始 xls 数据行（未经聚合） */
    getRawRows: function () { return state.rawRows || []; },
    /** 是否已加载数据 */
    hasData: function () { return !!(state.rows && state.rows.length); },
    /** 合计快照 { track, count, length, load, oldCar } */
    getSummary: getSummary,
    /** 订阅数据变化：fn()，返回反注册函数 */
    onDataChange: onDataChange,

    /* ---------- 配置 ---------- */
    /** 主表列定义 */
    getColumns: function () { return COLUMNS; },
    /** 明细表列定义（明细抽屉与搜索抽屉共用） */
    getDetailCols: function () { return DETAIL_COLS; },
    /** 股道显示名（到发线显示为「1道」等） */
    trackName: function (id) { return YardConfig.trackName(id); },
    /** 股道完整配置（含分组名、颜色、是否虚拟） */
    getTrack: function (id) { return YardConfig.getTrack(id); },

    /* ---------- 渲染 ---------- */
    /**
     * 渲染车辆行为明细表（与明细抽屉完全同款：列序、着色、停时列）
     * @param {Array} list 车辆行
     * @param {Object} els { head, body, table }
     * @param {Object} [opts] { destProcessed, rowAttr } 见 renderDetailRows
     */
    renderRows: function (list, els, opts) { renderDetailRows(list, els, opts); },
    /** 到站着色富文本（主表 / 明细 / 搜索共用同一套着色规则） */
    renderDest: function (text, clickable) { return renderDest(text, clickable); },
    /** 重绘主表 */
    renderGrid: function () { render(); },

    /* ---------- 交互 ---------- */
    /** 打开某股道明细（idx 为 state.rows 下标） */
    openDetail: function (idx) { openDetail(idx); },
    /** 打开某股道明细并高亮定位到指定车辆行 */
    openDetailAt: function (trackIdx, carIdx) { openDetailAt(trackIdx, carIdx); },
    /** 关闭明细抽屉 */
    closeDetail: function () { closeDetail(); },
    /** 右下角提示 */
    toast: function (msg, type) { toast(msg, type); }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else init();
})(window);