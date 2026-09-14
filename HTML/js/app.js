/**
 * app.js —— 股道存车主程序（页面装配层）
 *
 * 无服务器设计：
 *   - 方向库、股道配置通过 <script src> 加载（file:// 下唯一可靠方式）
 *   - xls 通过 File System Access API 读取，目录句柄存 IndexedDB 实现"打开即自动读取"
 *   - 不支持该 API 的浏览器自动降级为 <input type="file"> 手动选择
 *
 * ============================ 本文件职责 ============================
 * 这里是「页面装配层」：把各纯逻辑模块拼装成可交互界面。
 * 凡是能脱离 DOM 独立成立的计算，都不放在本文件——已抽出的模块：
 *
 *   js/columns.js        列定义（COLUMNS / DETAIL_COLS / BIANHAO_*）     纯常量
 *   js/dest-color.js     到站着色 / 车型高亮 / 车站名识别                纯函数
 *   js/checi-store.js    「编好车次」统一数据源（主表 / 发车流程 / 31814 共用）
 *   js/grid-layout.js    主表分组跨行、作业区横幅定位                     纯函数
 *   js/sim-panel.js      推演面板（自持全部状态与事件；本文件注入它需要的渲染能力）
 *   js/grid-events.js    主表事件：行选中 / 编好标记与分组高亮 / 编好车次录入 / 空线开关 /
 *                        注意事项收起 / 自适应列宽 / 上下一股道（本文件注入 openDepForRow 钩子）
 *   js/dep-panel.js      发车作业浮窗（自持全部状态与事件；本文件只调 DepPanel.init）
 *   js/settings.js       设置面板：字号滑块 / 默认文件夹 / 卸车地点 / 黑罐 / 货物推算重量 /
 *                        车型高亮（本文件注入 state / render / updateDetailTitle / openDetail）
 *   js/data-source.js    数据源：FSA 权限 / 文件夹读取 / xls 解析编排
 *   js/rpt31814-calc.js  31814 车流属性等纯计算
 *   js/config-io.js      配置导出/导入的 JSON 编解码                      纯转换
 *
 * 留在本文件的区块（按文件内出现顺序，用 `=== 区块名 ===` 注释分隔；
 * 刻意不写行号——行号改一次就失效，反而误导定位）：
 *
 *   [全局状态]        state / 虚拟股道显隐 / 数据文件夹按钮同步
 *   [渲染主表]        render / renderGridHead / renderGridFoot / computeTotals / bannerRow
 *   [明细抽屉]        openDetail / closeDetail / stepDetail / 计重编辑 / 明细多选
 *   [推演面板]        已抽到 js/sim-panel.js；本文件只留 4 个调用点（见该区块注释）
 *   [地图径路]        双击车站名打开径路
 *   [事件绑定]        bind() —— 剩余的 DOM 事件绑定（文件读写 / 明细多选 / 导入导出…）
 *                     （主表事件已抽到 js/grid-events.js；发车浮窗已抽到 js/dep-panel.js；
 *                       设置面板已抽到 js/settings.js）
 *   [浮窗拖动 / 缩放] 标题栏拖动、右下角手柄改尺寸
 *   [初始化]          init() —— 启动编排；并把 state / render / loading /
 *                     syncPickFolderBtn 注入 global，供子模块桥接访问
 *   [对外接口]        window.YardApp —— 数据 / 配置 / 渲染 / 交互的统一出口
 *
 * 说明：state 为 IIFE 内私有对象，子模块经 init 注入的 global 桥接访问；
 *       index.html 中所有子模块均在 app.js 之前加载。
 * ==================================================================
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

  /* 到站富文本着色 / 车型高亮 / 车站名识别 已抽到 js/dest-color.js（纯函数，挂 global）
   * 「编好车次」统一数据源已抽到 js/checi-store.js（Checi.of / Checi.set / Checi.dirCls） */

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
  /* 分组跨行计算（computeGroupSpans）与作业区横幅定位（computeBannerSlots）
   * 已抽到 js/grid-layout.js（纯函数、可单测），此处改用 GridLayout.groupSpans / bannerSlots。 */

  /* 作业区 → 站场示意图。键 = 横幅上的作业区名，值 = HTML/images/ 下的文件名。
   * 新增或替换示意图：把图片放进 HTML/images/，再在这里登记一行即可，不必改其它代码。
   * 未登记的作业区点图标会提示「暂无示意图」，不会出现破图。 */
  var AREA_IMAGES = {
    '货场作业区': 'images/货场作业区.png',
    '勒沟作业区': 'images/勒沟作业区.png',
    '鹰岭作业区': 'images/鹰岭作业区.png',
    '中油作业区': 'images/中油作业区.png'
  };

  /** 作业区横幅行：通栏单行，仅作分段标识，不参与选中 / 明细 / 合计。
   *  名字前挂一枚路徽图标（点击 → openAreaMap 打开该作业区示意图）。
   *  图标放在 .area-banner-text 内部：与文字共用同一 sticky 定位，
   *  横向滚动时不会与文字分离（放 td 里但 span 外就会分离）。 */
  function bannerRow(a) {
    var name = escapeHtml(a.name);
    return '<tr class="area-banner" data-area="' + name + '">' +
           '<td colspan="' + COLUMNS.length + '" style="background:' + (a.color || '#2b5cb0') + '">' +
           '<span class="area-banner-text">' +
             '<img class="area-banner-icon" src="images/luhui.png" alt="" ' +
                  'data-area="' + name + '" title="查看「' + name + '」示意图">' +
             name +
           '</span>' +
           '</td></tr>';
  }

  /* ---- 作业区示意图缩放 ----
   * areaView.cur：图片显示宽度 ÷ 原始像素宽度（1 = 1:1 原始尺寸）。
   * areaView.fit：「适应窗口」比例——按视口算出的最大不超出比例，上限 1（小图不放大）。
   * 缩放全程由 JS 设定 img.style.width 实现（所以 CSS 里不设 max-width/max-height，
   * 否则会截断放大结果），高度交给 height:auto 按原比例走。
   * 滚轮 = 自由缩放（以鼠标位置为锚点）；点图片 = 在「适应窗口」与「1:1」间切换。 */
  var AREA_ZOOM_MIN = 0.2;      // 最小缩到 fit 的 20%
  var AREA_ZOOM_MAX = 8;        // 最大 8 倍（位图再放大就糊了）
  var areaView = { fit: 1, cur: 1 };

  /** 「适应窗口」比例：与浮窗尺寸约束对应（94vw/92vh 减去内边距与标题栏） */
  function areaFitRatio(nw, nh) {
    if (!nw || !nh) return 1;
    var availW = window.innerWidth * 0.94 - 20;
    var availH = window.innerHeight * 0.92 - 80;
    return Math.min(availW / nw, availH / nh, 1);
  }

  /** 应用缩放。给了鼠标坐标就保持该点不动（滚轮缩放手感的来源）。 */
  function applyAreaZoom(z, clientX, clientY) {
    var img = $('areamapImg'), body = $('areamapBody');
    if (!img || !body || !img.naturalWidth) return;

    var nw = img.naturalWidth;
    z = Math.max(areaView.fit * AREA_ZOOM_MIN, Math.min(AREA_ZOOM_MAX, z));

    // 记录鼠标指向的是「原图上的哪个像素」（在老尺寸下换算）
    var r0 = null, px = 0, py = 0;
    if (clientX != null) {
      r0 = img.getBoundingClientRect();
      px = (clientX - r0.left) / (r0.width / nw);
      py = (clientY - r0.top) / (r0.height / img.naturalHeight);
    }

    img.style.width = (nw * z) + 'px';
    areaView.cur = z;

    if (r0) {
      // 缩放后该像素点跑到新位置，用滚动条把它拉回鼠标底下
      var r1 = img.getBoundingClientRect();
      body.scrollLeft += (r1.left + px * z) - clientX;
      body.scrollTop += (r1.top + py * z) - clientY;
    }

    updateAreaDragState();
  }

  /** 只有图片超出容器才「可拖拽」，据此切换光标（grab / zoom-in）。
   *  缩放后立即调用，光标才能跟着状态走。
   *  注：浮窗未打开时容器 clientWidth/scrollWidth 都是 0，判定为不可拖拽，不会误标。 */
  function updateAreaDragState() {
    var img = $('areamapImg'), body = $('areamapBody');
    if (!img || !body) return;
    var can = !img.hidden &&
              (body.scrollWidth > body.clientWidth + 1 ||
               body.scrollHeight > body.clientHeight + 1);
    img.classList.toggle('can-drag', can);
  }

  /** 打开「作业区示意图」浮窗：按作业区名取 images/ 下登记的图片。
   *  未登记 / 加载失败都在浮窗内给出文字提示，不弹破图。 */
  function openAreaMap(areaName) {
    var title = $('areamapTitle'), img = $('areamapImg'), empty = $('areamapEmpty');
    if (!title || !img || !empty) return;

    var src = AREA_IMAGES[areaName];
    title.textContent = areaName + '示意图';

    // 复位：先摘掉上一张的 src 与尺寸，避免残留旧图闪一下
    img.onload = img.onerror = null;
    img.removeAttribute('src');
    img.style.width = '';
    img.hidden = true;
    empty.hidden = true;

    if (!src) {
      empty.hidden = false;
      empty.textContent = '暂无「' + areaName + '」的示意图。\n' +
                          '把图片放入 HTML/images/ 并在 AREA_IMAGES 中登记即可。';
      UI.Modal.open('modalAreaMap');
      return;
    }

    img.onload = function () {
      var nw = img.naturalWidth, nh = img.naturalHeight;
      areaView.fit = areaFitRatio(nw, nh);
      img.style.width = (nw * areaView.fit) + 'px';   // 先按「适应窗口」显示
      areaView.cur = areaView.fit;
      img.hidden = false;
      img.title = '滚轮缩放 · 按住拖动 · 点击切换 1:1';
      title.textContent = areaName + '示意图（' + nw + '×' + nh + '）';
      updateAreaDragState();
    };
    img.onerror = function () {
      img.onerror = null;                    // 先解绑，避免下面移除 src 时再次触发
      img.hidden = true;
      img.removeAttribute('src');            // 清掉，下次点击可重新尝试加载
      empty.hidden = false;
      empty.textContent = '图片加载失败：' + src + '\n请确认文件已放入 HTML/images/ 目录。';
    };
    img.src = src;
    UI.Modal.open('modalAreaMap');
  }

  /**
   * 主表表头：构建 <th>、重挂列宽拖拽手柄、同步冻结列宽偏移。
   * 从 render() 拆出——表头只依赖 COLUMNS / ColResize，与行渲染无耦合。
   */
  function renderGridHead() {
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
  }

  function render() {
    renderGridHead();

    var vis = visibleRows();
    var thr = YardConfig.thresholds;
    var tbody = $('tbody');
    var html = [];

    // 分组「合并列」：预先算出每行的所属分组、是否该组首行、跨行数(span)。
    // 渲染时首行输出带 rowspan 的分组单元格，组内其余行不输出该 td（由 rowspan 覆盖）。
    var spans = GridLayout.groupSpans(vis);

    // 作业区横幅：定位每个作业区在可见行中的首行（该区无可见行时自动不显示）
    var bannerAt = GridLayout.bannerSlots(vis);

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
        // 编好车次：不是聚合字段，按股道从 Store 读（与 31814 报表「待发股道车次」同一份数据）
        if (c.key === 'checi') v = Checi.of(r.track);
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
          var rawNote = Utils.text(v);
          // 注意事项列：聚合时已用 \n 分隔各关键词（超71.86吨 / 扣修 …），
          // 转成 <br> 才能逐条换行；方向列同理（render 内已处理）。
          if (c.key === 'note') {
            // 包一层 .note-body，供表头「收起/展开」控件按状态裁剪高度
            inner = '<div class="note-body">' + escapeHtml(rawNote).replace(/\n/g, '<br>') + '</div>';
          } else if (c.key === 'checi') {
            // 车次颜色套用「车辆信息」列的方向色（沙口蓝/南口橙/管内紫），无方向走默认蓝
            inner = v ? '<span class="' + Checi.dirCls(r.direction) + '">' +
                       escapeHtml(String(v)) + '</span>' : '';
            if (!c.cls && !c.num) cls += ' center';
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
            cells.push('<td class="col-b-group grp" data-col="group" data-group="' + escapeHtml(sp.group) + '" rowspan="' + sp.span + '" ' +
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

      // data-group：每行都记上所属分组名——分组列用 rowspan 合并（组内只有首行有该单元格），
      // 选中某行时要按它反查"这一行属于哪个分组格"，给该格加底色（见 js/grid-events.js）
      html.push('<tr data-idx="' + idx + '" data-track="' + escapeHtml(track) + '"' +
                ' data-group="' + escapeHtml(spans[n] ? spans[n].group : '') + '"' +
                (isBlank ? ' class="blank"' : '') + '>' + cells.join('') + '</tr>');
    });

    tbody.innerHTML = html.join('');

    renderGridFoot(vis);
  }

  /** 罐车结存 + 状态栏（render 收口）。
   *  原「主表合计行」已移除——股道/车数/换长/载重/老牌车这些合计状态栏本来就有，
   *  罐车结存（自备罐/路罐）也一并从合计行挪到状态栏，避免同一份数据两处显示。
   *  统计与可见行保持一致：隐藏虚拟股道/空线分组后，状态栏同步变化。 */
  function renderGridFoot(vis) {
    var tc = 0, tl = 0, tw = 0, told = 0;
    vis.forEach(function (item) {
      var r = item.r;
      tc += r.count || 0; tl += r.length || 0; tw += r.load || 0; told += r.oldCar || 0;
    });
    // 罐车结存（自备罐/路罐）：原在主表合计行的「到站」列，现随合计行一并移到状态栏
    var zb = 0, lg = 0;
    vis.forEach(function (item) {
      var re = /(自备罐|路罐)(\d+)/g, m;
      var d = item.r.dest || '';
      while ((m = re.exec(d))) {
        if (m[1] === '自备罐') zb += +m[2]; else lg += +m[2];
      }
    });
    $('stTank').textContent = zb + '(自)/' + lg + '(路)';

    // 状态栏
    $('stTrack').textContent = vis.filter(function (item) { return item.r.count; }).length;
    $('stCount').textContent = tc;
    $('stLen').textContent = Utils.round1(tl);
    $('stLoad').textContent = Utils.round1(tw);
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
  /* 计重编辑下的「按住拖动批量删除推算载重」：与拖选行（dragSel）同一套交互范式——
   * 在推算载重格上按下记锚点与快照，mode 由起点格状态决定（未删→批量删 / 已删→批量恢复）；
   * 拖动经过的行按 mode 实时应用，回滑超出范围按快照复原（未松手回滑数值即恢复）。
   * 状态放顶层：bindDetailEditEvents（document 委托）与 bind()（detailBody 监听）都要访问。 */
  var dragExc = { active: false, moved: false, anchor: -1, last: -1, snap: null, mode: 'exclude' };

  /** 把一行的推算载重格设为 删除(excluded) / 恢复(derived)，同步该格 DOM；该行无可编辑
   *  推算格（实载行等）返回 false 跳过。返回是否发生变更。 */
  function setExcRow(tr, row, exc) {
    var td = tr.querySelector('td.derived-est, td.excluded-cell');
    if (!td) return false;
    var isExc = detailEditExcluded.has(row);
    if (exc === isExc) return false;                // 已是目标态
    if (exc) {
      detailEditExcluded.add(row);
      td.textContent = '';
      td.classList.remove('derived-est');
      td.classList.add('excluded-cell');
    } else {
      detailEditExcluded.delete(row);
      td.textContent = String(Math.round(estLoadOf(row)));   // 恢复显示推算值（与渲染口径一致）
      td.classList.remove('excluded-cell');
      td.classList.add('derived-est');
    }
    return true;
  }
  /** 拖动应用：**双侧范围**——锚点固定在按下行，范围 = 锚点到当前指针（cur 在锚点
   *  上/下方都有效），范围内按 mode 统一应用，范围外按快照复原。与拖选行 renderDrag
   *  同一结构：先全表对齐快照，再叠加当前范围。
   *  计重特有：指针回到起拖格 → 锚点..上次指针行 整段恢复快照（删除是破坏性操作，
   *  回到起点即撤销；快速滑动跳行时跳过的中间行一并覆盖）。选行无此分支（保持选中）。 */
  function applyExcRange(cur) {
    var r = state.rows[state.detailIdx];
    if (!r || !r.raw) return;
    var snap = dragExc.snap;
    if (cur === dragExc.anchor) {
      var loA = Math.min(dragExc.anchor, dragExc.last), hiA = Math.max(dragExc.anchor, dragExc.last);
      for (var iA = loA; iA <= hiA; iA++) {
        var rowA = r.raw[iA];
        var trA = $('detailBody').querySelector('tr[data-i="' + iA + '"]');
        if (rowA && trA) setExcRow(trA, rowA, snap.has(rowA));
      }
      updateDetailTitle();
      return;
    }
    var lo = Math.min(dragExc.anchor, cur), hi = Math.max(dragExc.anchor, cur);
    var trs = $('detailBody').querySelectorAll('tr');
    var changed = false;
    for (var i = 0; i < trs.length; i++) {
      var idx = +trs[i].getAttribute('data-i');
      var row = r.raw[idx];
      if (!row) continue;
      var exc = (idx >= lo && idx <= hi) ? (dragExc.mode === 'exclude') : snap.has(row);
      if (setExcRow(trs[i], row, exc)) changed = true;
    }
    if (changed) updateDetailTitle();               // 计重合计实时重算
  }
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
      length: Utils.round1(len),
      selfW: Utils.round1(selfW),
      loadW: Utils.round1(loadW),
      estLoad: Utils.round1(estLoad),        // 推算载重合计（载重缺失按车型/货物补全）
      weight: Utils.round1(selfW + loadW),   // 总重 = 自重 + 载重
      calcW: Utils.round1(selfW + estLoad)   // 计重 = 自重 + 推算重量（推算载重）
    };
  }
  /** 生成「辆数 / 换长 / 自重 / 载重 / 总重」五个统计 span。
   *  明细抽屉与推演面板共用同一口径，避免两处各写一套、日后悄悄走偏。
   * @param {Object} t    computeTotals() 的结果
   * @param {Object} [opts]
   *        calc:true     显示「计重」而非「总重」（明细编辑模式）
   *        editable:true 挂 dt-weight-edit 类（支持双击进入计重编辑）
   *        editing:true  编辑态样式
   */
  function totalsSpansHtml(t, opts) {
    opts = opts || {};
    // 超限阈值统一取 YardConfig.thresholds，避免与主表 render() 各写一套魔数
    var thr = YardConfig.thresholds || { overlong: 70, overloadTons: 5000 };
    // 总重 / 计重合并为一个 span：默认显示「总重」，双击进入编辑模式后切换显示「计重」。
    //（warn 按当前显示的数值判定：非编辑看总重是否超限、编辑看计重是否超限。）
    var showWeight = opts.calc ? t.calcW : t.weight;
    return '<span class="dt-total">辆数：' + t.count + '</span>' +
      '<span class="dt-total' + (t.length > thr.overlong ? ' warn' : '') + '">换长：' + t.length.toFixed(1) + '</span>' +
      '<span class="dt-total">自重：' + t.selfW.toFixed(1) + '</span>' +
      '<span class="dt-total">载重：' + t.loadW.toFixed(1) + '</span>' +
      '<span class="dt-total dt-weight' + (opts.editable ? ' dt-weight-edit' : '') +
        (opts.editing ? ' editing' : '') +
        (showWeight > thr.overloadTons ? ' warn' : '') + '">' +
        (opts.calc ? '计重：' : '总重：') + showWeight.toFixed(1) +
      '</span>';
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
      // 「股道」名前挂路徽小图标（与工具栏路徽呼应）
      '<span class="dt-name"><img class="detail-emblem" src="images/luhui.jpeg" alt="">' +
        escapeHtml(name ? name.name : r.track) + ' - </span>' +
      totalsSpansHtml(t, { calc: detailEditMode, editable: true, editing: detailEditMode });
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
    var _wLast = { t: 0, sim: false };   // 双击检测：上次点击时间 + 目标归属（明细 / 推演）
    document.addEventListener('mousedown', function (e) {
      if (e.button !== 0) { _wLast.t = 0; return; }
      var t = e.target.closest && e.target.closest('.dt-weight-edit');
      if (!t) { _wLast.t = 0; return; }
      // 明细标题与推演标题各有一个 .dt-weight-edit，按归属分发，互不串扰
      var isSim = !!(t.closest && t.closest('#simTitle'));
      var now = Date.now();
      if (_wLast.t && (now - _wLast.t) < 350 && _wLast.sim === isSim) {
        if (isSim) SimPanel.toggleEditMode();   // 推演表的计重编辑态由 sim-panel 自持
        else setDetailEditMode(!detailEditMode);
        _wLast.t = 0;
      } else {
        _wLast.t = now;
        _wLast.sim = isSim;
      }
    });
    document.addEventListener('click', function (e) {
      var td = e.target.closest && e.target.closest('td.derived-est');
      if (!td) return;
      // 明细与推演是两张表，各自的编辑态只作用于各自的推算值单元格，
      // 否则一张表开着编辑、点另一张表的推算值会拿错行。
      // 推演表那一支整段交给 sim-panel（含「非编辑态直接忽略」的判定）。
      if (td.closest('#simTable')) { SimPanel.excludeDerivedEst(td); return; }
      if (!detailEditMode) return;
      if (!td.closest('#detailTable')) return;
      // 按住拖动批量删除已在 mouseover 中按范围应用，松手后的 click 不再切换单格
      if (dragExc.moved) { dragExc.moved = false; return; }
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
        // 「顺」列可被覆盖：推演面板要显示「本板序号」(1、2、3…)，
        // 而车辆在原股道里的顺位跨股道凑车时没有意义（会出现 3、7、12、2…）。
        if (c.col === COL.SEQ && opts.seqText) {
          return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '>' +
                 escapeHtml(opts.seqText(row, i)) + '</td>';
        }
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
          // 到站列：显示值判定抽到 dest-color.js 的 destDisplayValue（纯函数），
          // 明细抽屉与搜索抽屉共用同一口径，避免两处各写一遍、日后悄悄走偏。
          var dd = destDisplayValue(row, opts.destProcessed);
          if (dd.text) {
            return '<td class="dest">' +
                   renderDest(dd.text, true, dd.derived ? 'derived' : '') +
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
          var pm = Utils.text(raw);
          var jishi = Utils.text(row[COL.NOTE]);
          if ((pm.indexOf('汽油') !== -1 || pm.indexOf('航煤') !== -1 || jishi.indexOf('汽油') !== -1) && jishi !== '原装汽油') {
            var clsP = c.cls ? c.cls + ' car-yellow-bg' : 'car-yellow-bg';
            return '<td class="' + clsP + '">' + escapeHtml(pm) + '</td>';
          }
        }
        return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '>' +
               escapeHtml(raw == null ? '' : raw) + '</td>';
      }).join('');
      var attr = opts.rowAttr ? (opts.rowAttr(row, i) || '') : '';
      // 停时列：在拼接阶段直接算好写进 <td>，省掉渲染后的二次全表遍历。
      // （原实现先 innerHTML 重建，再 querySelectorAll('tr') + 每行 querySelector('td.stay')
      //   逐个回填——大股道时是 O(n) 次 DOM 查询，且多一轮样式重算。）
      // 复用 Utils.parseArriveTime：兼容 Date 实例、"2026/9/2T08:30:00"、"2026年9月2日" 等写法
      // 优先复用 aggregate 预处理已缓存的 __arrTime；缺省时回退解析（兼容非聚合来源的行）
      var arrD = (row.__arrTime != null) ? row.__arrTime : Utils.parseArriveTime(row[COL.ARRTIME]);
      var hrs = arrD ? Math.floor((base - arrD) / 3600000) : '';
      var stayStyle = (hrs !== '' && hrs > YardConfig.thresholds.bigCarHours)
        ? ' style="background:#fff3cd;font-weight:700"' : '';
      return '<tr' + attr + '>' + tds + '<td class="stay"' + stayStyle + '>' + hrs + '</td></tr>';
    }).join('');
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
    // 之前点过「主页」：本次打开明细时把推演面板一起带回来（内容保留）
    SimPanel.resumeOnOpen();
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
    SimPanel.close();           // 关闭明细时同时收起推演面板（内容保留，只有「重置」才清空）
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

  /* =================== 推演面板（明细抽屉的「分身」） ===================
   * 已抽到 js/sim-panel.js（依赖注入 + 单向调用）：
   *   state / renderDetailRows / computeTotals / totalsSpansHtml / closeDetail
   *   由本文件在 init() 时注入；本文件只在 4 个点回调它——
   *   init（装配）、resumeOnOpen（openDetail 末尾）、close（closeDetail 开头）、
   *   toggleEditMode 与 excludeDerivedEst（「双击计重」手势的派发）。
   * 面板内容与全部状态由模块内部持有。 */

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

    /* 作业区横幅上的路徽图标 → 打开该作业区的站场示意图。
     * 横幅行由 render() 用 innerHTML 重建，不能逐个绑，故委托到 tbody。
     * 行选中/双击的处理里对 .area-banner 都是直接 return，这里不必阻断冒泡。 */
    on('tbody', 'click', function (e) {
      var icon = e.target.closest && e.target.closest('.area-banner-icon');
      if (!icon) return;
      e.preventDefault();
      openAreaMap(icon.getAttribute('data-area'));
    });

    /* 示意图缩放：滚轮自由缩放（以鼠标位置为锚点），点图片在「适应窗口」与 1:1 间切换。
     * wheel 必须用 { passive:false } 才能 preventDefault（吃掉容器的原生滚动），
     * 而 Utils.on 不支持传 addEventListener 的 options，故这里手动绑。 */
    var areaBody = $('areamapBody');
    if (areaBody) {
      areaBody.addEventListener('wheel', function (e) {
        var img = $('areamapImg');
        if (!img || img.hidden || !img.naturalWidth) return;
        e.preventDefault();
        // 指数步长：滚轮一格与触控板小幅滑动都能得到顺手的缩放量
        applyAreaZoom(areaView.cur * Math.exp(-e.deltaY * 0.0015), e.clientX, e.clientY);
      }, { passive: false });
    }

    /* ---- 示意图拖拽平移 ----
     * 抓图拖动 = 反向滚动容器。mouseup 把「本次是否真拖动过」留给 click 消费，
     * 否则拖完松手会顺带触发「适应窗口 / 1:1」切换。
     * mousedown 里 preventDefault：挡掉 <img> 的原生拖拽（会拖出半透明鬼影）与文本选择。 */
    var areaDrag = null;
    var areaDragMoved = false;

    on('areamapImg', 'mousedown', function (e) {
      if (e.button !== 0 || !this.naturalWidth) return;      // 只响应左键
      var body = $('areamapBody');
      if (!body) return;
      e.preventDefault();
      areaDragMoved = false;
      areaDrag = {
        startX: e.clientX, startY: e.clientY,
        sl: body.scrollLeft, st: body.scrollTop,
        moved: false
      };
      body.classList.add('dragging');
    });

    document.addEventListener('mousemove', function (e) {
      if (!areaDrag) return;
      var body = $('areamapBody');
      if (!body) return;
      var dx = e.clientX - areaDrag.startX;
      var dy = e.clientY - areaDrag.startY;
      // 超过 4px 才算拖动，用于区分「拖动」与「点击」
      if (!areaDrag.moved && (Math.abs(dx) > 4 || Math.abs(dy) > 4)) areaDrag.moved = true;
      body.scrollLeft = areaDrag.sl - dx;                    // 鼠标右移 → 内容左移
      body.scrollTop  = areaDrag.st - dy;
    });

    document.addEventListener('mouseup', function () {
      if (!areaDrag) return;
      areaDragMoved = areaDrag.moved;                        // 交给 click 判断
      areaDrag = null;
      var body = $('areamapBody');
      if (body) body.classList.remove('dragging');
    });

    on('areamapImg', 'click', function (e) {
      if (areaDragMoved) { areaDragMoved = false; return; }  // 刚才是拖动，不是点击
      if (!this.naturalWidth) return;
      // 当前在「适应窗口」→ 切到 1:1；否则切回「适应窗口」
      var atFit = Math.abs(areaView.cur - areaView.fit) < 1e-6;
      applyAreaZoom(atFit ? 1 : areaView.fit, e.clientX, e.clientY);
    });

    // 窗口尺寸变化：只在「适应窗口」态跟随重算，已经手动缩放过的保持不动
    window.addEventListener('resize', function () {
      var img = $('areamapImg');
      if (!img || img.hidden || !img.naturalWidth) return;
      var wasFit = Math.abs(areaView.cur - areaView.fit) < 1e-6;
      areaView.fit = areaFitRatio(img.naturalWidth, img.naturalHeight);
      if (wasFit) applyAreaZoom(areaView.fit);
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
    // 拖动选行：**双侧范围**——锚点固定在按下行，范围 = 锚点到当前指针（cur 在锚点
    // 上/下方都有效），范围内按 mode 应用，范围外按快照复原。与 applyExcRange 同一结构。
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
      // 计重编辑模式：在推算载重格上左键按下 → 启动「拖动批量删除」（锚点 + 快照 + 模式）。
      // 应用发生在 mouseover（拖动）与 click（单击，走原切换逻辑），这里只记录。
      if (detailEditMode) {
        var det = e.target.closest && e.target.closest('td.derived-est, td.excluded-cell');
        if (det && e.button === 0) {
          e.preventDefault();                    // 防止拖动选中单元格文字
          var trE = det.closest('tr');
          var idxE = trE ? +trE.getAttribute('data-i') : -1;
          var rE = state.rows[state.detailIdx];
          var rowE = (rE && rE.raw) ? rE.raw[idxE] : null;
          if (rowE) {
            dragExc.active = true;
            dragExc.moved = false;
            dragExc.anchor = idxE;
            dragExc.last = idxE;
            dragExc.snap = new Set(detailEditExcluded);          // 快照：范围外按快照复原
            dragExc.mode = detailEditExcluded.has(rowE) ? 'restore' : 'exclude';  // 起点状态定模式
          }
          return;
        }
        return;                                // 编辑模式下其他区域不启动拖选行
      }
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
      // 拖动批量删除：经过的行按 mode 实时应用，回滑超出范围按快照复原
      if (dragExc.active) {
        var trX = e.target.closest('tr');
        if (!trX) return;
        var iX = +trX.getAttribute('data-i');
        if (isNaN(iX) || iX === dragExc.last) return;
        dragExc.moved = true;
        applyExcRange(iX);
        dragExc.last = iX;
        return;
      }
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
      // 拖动批量删除结束；moved 保留给随后的 click 压制（click 里清）
      dragExc.active = false;
      dragExc.anchor = -1;
      dragExc.last = -1;
      dragExc.snap = null;
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
      // 推演面板（及分隔条、面板按钮）内的操作不清空明细选中：
      // 否则在推演里点一下，刚挑好的一批车就没了，「+」会读不到选中。
      if (e.target.closest && e.target.closest('#simDrawer, #simSplitter, #btnSim')) return;
      var dt = $('detailTable');
      if (dt && !dt.contains(e.target)) clearDetailSel();
    });

    /* ---- 主表事件：整块交给 js/grid-events.js ----
     * 行选中 / 编好标记与分组高亮 / 编好车次双击录入 / 空线分组开关 /
     * 注意事项收起展开 / 自适应列宽 / 上下一股道，全部由该模块自绑。
     * 这里只注入它需要的能力；「编好」命中后的开窗与预填由发车模块提供（见下）。 */
    // 发车浮窗先装配：grid-events 要拿它的 openForRow 当「编好」命中的钩子
    DepPanel.init();
    GridEvents.init({
      state: state,
      render: render,
      syncVirtualBtn: syncVirtualBtn,
      stepDetail: stepDetail,
      openDepForRow: DepPanel.openForRow
    });
    /* ================= 发车浮窗（点「编好」后弹出） =================
     * 整块已抽到 js/dep-panel.js：打开与预填 / 警示标与查重 / iframe 与扩展桥读写 /
     * 发送与读表 / 自动提交 / 刷新报表 / 回执聚合 / 时间与状态提示 / 浮窗注册。
     * 本文件只在上面 bind() 里调一次 DepPanel.init()。 */
    // 自适应列宽 / 上一股道 / 下一股道：已移入 js/grid-events.js

    /* ---- 推演面板：整块交给 js/sim-panel.js ----
     * 该模块自持全部状态（车辆 / 去重 / 选中 / 计重排除），并自绑按钮、面板内行选中、分隔条拖动。
     * 这里只注入它渲染与统计所需的能力（复用明细同一套，保证两边口径一致）。 */
    SimPanel.init({
      state: state,
      renderDetailRows: renderDetailRows,
      computeTotals: computeTotals,
      totalsSpansHtml: totalsSpansHtml,
      closeDetail: closeDetail
    });

    /* ---- 浮窗 / 抽屉 / 下拉菜单 -----
     * 注册后自动获得「点空白处关闭 + ESC 关栈顶」，新增面板无需再改 ESC 处理。 */
    // 31814 报表里改动过「待发股道车次」后，关窗即重渲主表——
    // 两边读写同一个 Store.KEYS.readyTrains，"编好车次"列随即同步（用户仍可在报表里改/删，行为不变）
    UI.Modal.register('modal31814', { onClose: function () { render(); } });
    // modalSettings（设置浮窗）的注册与 onOpen 已随设置面板移入 js/settings.js
    UI.Modal.register('modalProductivity');
    UI.Modal.register('modalAreaMap');
    UI.Drawer.register('drawer', { maskId: 'drawerMask' });
    UI.Drawer.register('searchDrawer', { maskId: 'searchMask' });
    // 推演抽屉 simDrawer 的注册与状态复位已移入 sim-panel.js（SimPanel.init）

    // 各面板的关闭按钮
    on('btnCloseDrawer', 'click', closeDetail);
    on('rptClose', 'click', function () { UI.Modal.close('modal31814'); });
    on('btnSettingsClose', 'click', function () { UI.Modal.close('modalSettings'); });
    on('areamapClose', 'click', function () { UI.Modal.close('modalAreaMap'); });

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
    // 原工具栏「发车流程」按钮（btnDepartureFlow）已移除：
    // 改由主表「编好车次」列（原「空箱/空车」，表头 th[9]）的「编好」伪元素按钮打开同一浮窗。

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

    /* 设置面板（字号滑块 / 默认文件夹 / 卸车地点 / 黑罐 / 货物推算重量 / 车型高亮）
     * 整块已抽到 js/settings.js；本文件只在下方原 initSpotConfig 调用处调一次 Settings.init。 */

    /* ---- 设置面板：整块交给 js/settings.js ----
     * 字号滑块 / 默认文件夹 / 卸车地点 / 黑罐识别 / 货物推算重量 / 车型高亮，
     * 由该模块自持全部交互与 modalSettings 注册；这里只注入它需要的能力。
     * recompute（配置变更重跑聚合）也已随之搬走，仅供该模块内部使用。 */
    Settings.init({
      state: state,
      render: render,
      updateDetailTitle: updateDetailTitle,
      openDetail: openDetail,
      defaultEstGoods: DEFAULT_EST_GOODS
    });

    // 表头列宽拖动（persistKey 用于本地记忆，刷新/重渲染不丢失）
    ColResize.enable($('grid'), {
      persistKey: Store.KEYS.gridCols,
      onResize: function (th, w) {
        // 首列 / 分组合并列宽变化 → 同步后续冻结列的偏移，避免列间露缝
        if (th.classList.contains('col-a')) $('grid').style.setProperty('--col-a-w', w + 'px');
        if (th.classList.contains('col-b-group')) $('grid').style.setProperty('--col-b-group-w', w + 'px');
      }
    });
    /* 列宽按「列序索引」记忆，因此调整 DETAIL_COLS 的顺序或增删列后，
     * 旧记忆会整体错位（宽度套到了别的列上）。键名带版本号即可让旧记忆失效、
     * 首次打开重新按内容自适应——改动列序时，把 store.js 里 KEYS.detailCols / KEYS.gridCols
     * 的版本号（.v2）递增即可，两处一起改。 */
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
  var DEP_SIZE_KEY = Store.KEYS.depModalSize;
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
    var savedEmpty = Store.get(Store.KEYS.showEmptyGroups, null);
    if (savedEmpty !== null) state.showEmptyGroups = !!savedEmpty;

    // 恢复「注意事项收起」的持久记忆（表头点击切换时写入 Store.notesCollapsed）
    var savedNotes = Store.get(Store.KEYS.notesCollapsed, null);
    if (savedNotes === true && $('grid')) $('grid').classList.add('notes-collapsed');

    syncVirtualBtn();

    // 先渲染空框架：让页面一打开就呈现完整股道清单，便于核对配置
    renderEmpty();

    // 护眼色：图标开关（0 ↔ 100）+ 浓度滑块（精细调节），两者共用同一份 Store.KEYS.eyeProtect
    (function () {
      var slider = $('eyeProtectSlider');
      var overlay = $('eyeOverlay');
      var btn = $('btnEyeProtect');
      if (!slider || !overlay) return;

      /** 统一落点：滑块位置、遮罩浓度、开关外观、持久化 四处一起更新，避免出现"界面和数据不一致" */
      function applyEye(val) {
        var v = parseInt(val, 10) || 0;
        if (v < 0) v = 0; else if (v > 100) v = 100;
        slider.value = v;
        overlay.style.display = v > 0 ? '' : 'none';
        overlay.style.opacity = v / 100;
        Store.set(Store.KEYS.eyeProtect, v);
        if (btn) {
          var on = v > 0;
          btn.classList.toggle('on', on);
          btn.setAttribute('aria-pressed', on ? 'true' : 'false');
          btn.title = on ? '护眼模式：开（点击关闭）' : '护眼模式：关（点击开启）';
        }
      }

      on('eyeProtectSlider', 'input', function () { applyEye(this.value); });
      on('btnEyeProtect', 'click', function () {
        // 开关语义：只在「关 0」与「开 100」之间切换，不记忆中间值
        applyEye(parseInt(slider.value, 10) > 0 ? 0 : 100);
      });

      applyEye(Store.get(Store.KEYS.eyeProtect, 0));   // 打开页面时恢复上次状态
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