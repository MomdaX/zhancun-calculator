/**
 * app.js —— 股道存车主程序
 *
 * 无服务器设计：
 *   - 方向库、股道配置通过 <script src> 加载（file:// 下唯一可靠方式）
 *   - xls 通过 File System Access API 读取，目录句柄存 IndexedDB 实现"打开即自动读取"
 *   - 不支持该 API 的浏览器自动降级为 <input type="file"> 手动选择
 *
 * ============================ 功能区块索引 ============================
 * 本文件约 1016 行，按职责划分为以下区块，便于定位（行号为当前快照，后续可能偏移）：
 *
 *   [列定义]        L13   COLUMNS / DETAIL_COLS —— 主表与明细表的列元数据
 *   [数据源]        L77   ensurePerm / pickFolder / loadFromDir / readAndRender
 *                        / renderFileSwitcher —— FSA 权限、IndexedDB、目录读取、文件切换
 *                        （最独立、state 引用最少；含浏览器原生 API，改后需手动点验）
 *   [主表渲染]      L302  renderDest / render / computeTotals / renderEmpty
 *                        —— 到站着色、主表 tbody 构建、合计、空态
 *   [明细抽屉]      L564  updateDetailTitle / openDetail / closeDetail / stepDetail
 *                        —— 股道明细抽屉的打开、翻页、标题
 *   [明细多选]      L764  selectRow / renderDrag / endDrag / clearDetailSel
 *                        —— 明细行拖拽范围选 + 单击切换（bind 内联，未抽组件）
 *   [事件绑定]      L729  bind —— 全部 DOM 事件绑定
 *   [入口]          L969  init —— 启动编排
 *
 * 说明：state 为 IIFE 内私有对象（约 51 处引用），故未做文件拆分；
 *       如需拆分，数据源区块（L77-301）耦合最弱，最优先。
 * =====================================================================
 */
(function () {
  'use strict';

  /* ============================ 列定义 ============================ */
  var COLUMNS = [
    { key: 'note',      title: '注意事项',       width: 112, cls: 'col-a' },
    { key: 'track',     title: '股道',           width: 66,  cls: 'col-b track' },
    { key: 'direction', title: '方向',           width: 64 },
    { key: 'count',     title: '车数',           width: 48,  num: true, cls: 'mid' },
    { key: 'carTypes',  title: '车种',           width: 124 },
    { key: 'length',    title: '换长',           width: 58,  num: true, cls: 'mid' },
    { key: 'dest',      title: '到站',           width: 340, dest: true },
    { key: 'empty',     title: '空箱/空车',      width: 78,  num: true },
    { key: 'heavy',     title: '重车',           width: 52,  num: true },
    { key: 'train',     title: '到达车次',       width: 76 },
    { key: 'load',      title: '载重',           width: 72,  num: true, cls: 'mid' },
    { key: 'oldCar',    title: '老牌车',         width: 60,  num: true, cls: 'mid' }
  ];

  // fmt: track = 股道显示名（到发线加"道"）  dest = 到站按方向着色
  // col 取 Aggregate.COL 常量：SMIS 导出列序变动时只需改 aggregate.js 一处。
  //
  // 【关于旧注释的更正】此处原写作「aggregate.js 在 app.js 之后才加载，不能在
  // IIFE 顶部引用 COL，必须做成函数」，与 index.html 的实际加载顺序相反：
  // aggregate.js（第 207 行）先于 app.js（第 210 行）。且本文件第 67 行的
  // `var COL = Aggregate.COL;` 本就是顶层引用，早已证明该限制不存在。
  // 列定义是常量，故直接写成常量数组，无需再包一层函数。
  var DETAIL_COLS = [
    { col: Aggregate.COL.TRACK,   t: '股道', w: 54, fmt: 'track' },
    { col: Aggregate.COL.SEQ,     t: '顺', w: 34 },
    { col: Aggregate.COL.CARTYPE, t: '车种', w: 68 }, { col: Aggregate.COL.CARNO, t: '车号', w: 78 },
    { col: Aggregate.COL.TARE,    t: '自重', w: 52 }, { col: Aggregate.COL.LEN,   t: '换长', w: 50, cls: 'mid' },
    { col: Aggregate.COL.LOAD,    t: '载重', w: 56, cls: 'mid' },
    { col: Aggregate.COL.DEST,    t: '到站', w: 120, fmt: 'dest' },
    { col: Aggregate.COL.DIR,     t: '方向', w: 42 }, { col: Aggregate.COL.GOODS, t: '品名', w: 90 },
    { col: Aggregate.COL.FROM,    t: '发站', w: 110, fmt: 'dest' },
    { col: Aggregate.COL.NOTE,    t: '记事', w: 170 },
    { col: Aggregate.COL.TRAIN,   t: '车次', w: 62 }, { col: Aggregate.COL.ARRTIME, t: '到达时间', w: 128 }
  ];

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
    showVirtual: true     // 是否显示虚拟股道（X 线）
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

  /* =================== 目录权限 =================== */
  function ensurePerm(handle, mode) {
    var opts = { mode: mode || 'read' };
    if (!handle.queryPermission) return Promise.resolve(true);
    return handle.queryPermission(opts).then(function (p) {
      if (p === 'granted') return true;
      if (!handle.requestPermission) return false;
      return handle.requestPermission(opts).then(function (p2) { return p2 === 'granted'; });
    });
  }

  /** 列出目录中的 xls 文件，按修改时间倒序（相同则按名称倒序） */
  function listXlsInDir(dirHandle) {
    var files = [];
    if (!dirHandle.entries) return Promise.resolve(files);
    var it = dirHandle.entries(), readNext;
    readNext = function () {
      return it.next().then(function (r) {
        if (r.done) return files;
        var entry = r.value;
        if (!entry) return files;
        var name = entry[0], h = entry[1];
        var step = function () { return readNext(); };
        if (h.kind !== 'file' || !/\.(xls|xlsx)$/i.test(name)) return step();
        return h.getFile().then(function (f) {
          files.push({ name: name, handle: h, lastModified: f.lastModified, size: f.size });
          return step();
        }).catch(function () { return step(); });
      });
    };
    return readNext().then(function (list) {
      list.sort(function (a, b) {
        return (b.lastModified - a.lastModified) || b.name.localeCompare(a.name);
      });
      return list;
    });
  }

  /* =================== 文件加载入口 =================== */

  /** 记住已选定的数据文件夹名（统一两处写入，避免重复直调 Store） */
  function rememberFolder(h) {
    if (h && h.name) Store.set('folderName', h.name);
  }

  /** 选择文件夹（File System Access API） */
  function pickFolder() {
    if (!window.showDirectoryPicker) {
      // 已直接弹出文件选择器，提示语不必再指向某个不存在的按钮
      toast('当前浏览器不支持文件夹选择，请在弹出的窗口中选择 xls 文件', 'error');
      $('fileInputMulti').click();
      return;
    }
    window.showDirectoryPicker({ id: 'yardXls', mode: 'read' })
      .then(function (h) {
        rememberFolder(h);
        syncPickFolderBtn(false);   // 已选定，按钮功成身退
        return Store.async.set('xlsDir', h).then(function () {
          state.dirHandle = h;
          return loadFromDir(h, true);
        });
      })
      .catch(function (e) {
        if (e && e.name === 'AbortError') return;
        toast('选择文件夹失败：' + (e && e.message || e), 'error');
      });
  }

  /** 从文件夹读取最新 xls；showPicker=true 时若无文件则提示 */
  function loadFromDir(dirHandle, showPicker) {
    return ensurePerm(dirHandle, 'read').then(function (ok) {
      if (!ok) { toast('未获得文件夹读取权限', 'error'); return; }
      loading(true, '正在扫描文件夹…');
      return listXlsInDir(dirHandle);
    }).then(function (files) {
      if (!files) return;
      if (!files.length) {
        loading(false);
        toast('该文件夹内没有 xls 文件', 'error');
        return;
      }
      // 自动取最新；若多于 1 个，在状态栏提示可切换
      var target = files[0];
      state.fileList = files;
      return readAndRender(target.handle.getFile(), target.name, files);
    }).catch(function (e) {
      loading(false);
      toast('读取失败：' + (e && e.message || e), 'error');
    });
  }

  /** 读取 File 对象 → 解析 → 聚合 → 渲染 */
  function readAndRender(filePromise, fileName, fileList) {
    loading(true, '正在解析 ' + (fileName || '') + ' …');
    return Promise.resolve(filePromise)
      .then(function (file) {
        return file.arrayBuffer();
      })
      .then(function (buf) {
        // 让出一帧，确保 loading 遮罩先渲染出来再开始同步解析
        return new Promise(function (resolve) {
          requestAnimationFrame(function () {
            requestAnimationFrame(function () { resolve(buf); });
          });
        });
      })
      .then(function (buf) {
        var wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true });
        var ws = wb.Sheets[wb.SheetNames[0]];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });

        // 定位表头行（含"股道"）
        var hIdx = -1;
        for (var i = 0; i < Math.min(aoa.length, 12); i++) {
          if (aoa[i] && aoa[i].indexOf('股道') >= 0) { hIdx = i; break; }
        }
        if (hIdx < 0) throw new Error('未找到含「股道」的表头行');

        // 提取标题行中的打印日期，作为停时基准
        var pd = null;
        for (var j = 0; j < hIdx; j++) {
          var line = (aoa[j] || []).join(' ');
          var m = /(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2})/.exec(line);
          if (m) {
            pd = new Date(+m[1], +m[2] - 1, +m[3], +m[4], +m[5]);
            break;
          }
        }
        state.printDate = pd;

        var rows = aoa.slice(hIdx + 1).filter(function (r) {
          return r && r.some(function (c) { return String(c).trim() !== ''; });
        });

        state.rawRows = rows;
        var base = pd || new Date();
        var agg = Aggregate.aggregate(rows, state.dirIndex.map, state.dirIndex.stations,
                                      YardConfig.thresholds, base);

        // 按配置清单顺序排列；清单内无数据的股道保留空行（与 Excel 的 92 行框架一致）
        // 清单外的股道追加在末尾
        var ordered = [], known = {}, extra = [];
        YardConfig.tracks.forEach(function (t) {
          if (agg[t.id]) { ordered.push(agg[t.id]); }
          else {
            ordered.push({
              track: t.id, direction: '', count: 0, carTypes: '', length: 0,
              dest: '', train: '', load: 0, oldCar: 0, raw: []
            });
          }
          known[t.id] = 1;
        });
        Object.keys(agg).forEach(function (k) {
          if (!known[k]) extra.push(agg[k]);
        });
        extra.sort(function (a, b) { return a.track.localeCompare(b.track, 'zh'); });
        state.rows = ordered.concat(extra);

        state.currentFile = fileName || '手动选择的文件';
        render();
        loading(false);

        var extraMsg = extra.length ? '，其中清单外 ' + extra.length + ' 个已追加' : '';
        var multi = '';
        if (fileList && fileList.length > 1) {
          // 只有当加载的确实是列表首项（最新）时才说"已取最新"
          multi = (fileList[0] && fileList[0].name === state.currentFile)
            ? '（文件夹内共 ' + fileList.length + ' 个文件，已取最新）'
            : '（文件夹内共 ' + fileList.length + ' 个文件）';
        }
        toast('已加载：' + state.currentFile + multi + extraMsg, 'ok');
        if (fileList && fileList.length > 1) {
          renderFileSwitcher(fileList, state.currentFile);
        }
      })
      .catch(function (e) {
        loading(false);
        toast('解析失败：' + (e && e.message || e), 'error');
        console.error(e);
      });
  }

  /**
   * 文件夹内有多个文件时，显示切换列表。
   * activeName：刚加载完成的文件名，用于重建选项后把下拉框还原到对应项
   * （innerHTML 重建会令 selectedIndex 归零，不还原会看起来"切换无效"）。
   */
  function renderFileSwitcher(files, activeName) {
    var sel = $('fileSwitcher');
    if (!sel) {
      var wrap = document.createElement('span');
      wrap.innerHTML = '<select id="fileSwitcher" class="btn" style="max-width:236px"></select> ';
      $('btnReload').parentNode.insertBefore(wrap, $('btnReload'));
      sel = $('fileSwitcher');
      sel.addEventListener('change', function () {
        var f = state.fileList[sel.selectedIndex];
        if (f) readAndRender(f.handle.getFile(), f.name, state.fileList);
      });
    }
    // 先算出目标选中项，写入 innerHTML 后立即还原
    var idx = -1;
    for (var i = 0; i < files.length; i++) {
      if (files[i].name === activeName) { idx = i; break; }
    }
    if (idx < 0) idx = sel.selectedIndex;   // 兜底：保持重建前的选择

    sel.innerHTML = files.map(function (f, i) {
      var match = f.name.match(/共\s*(\d+)\s*辆/);
      var count = match ? match[1] : '?';
      var label = count + '辆-' + Utils.formatDateTime(new Date(f.lastModified));
      var isNew = (i === 0);
      if (isNew) label += ' ● NEW';
      return '<option value="' + i + '"' +
             (isNew ? ' style="font-weight:800;color:var(--accent)"' : '') +
             '>' + label + '</option>';
    }).join('');

    if (idx >= 0 && idx < files.length) sel.selectedIndex = idx;
  }

  /* =================== 到站富文本着色 =================== */
  /**
   * @param {string}  text 到站/发站文本
   * @param {boolean} clickable 是否标记可双击的车站（仅明细页为 true，
   *        主表是汇总串如「德保44」，双击无意义，故不加虚线下划线）
   */
  function renderDest(text, clickable) {
    if (!text) return '';
    var parts = String(text).trim().split(/\s+/).filter(Boolean);
    var map = state.dirIndex.map;
    return parts.map(function (p) {
      var cls = '';
      // 属性标记优先（对应 VBA 的 Array("YW","D","P","到卸","黑罐")）
      if (/^(YW|D|P)/.test(p)) cls = 'danger';
      else if (/^(到卸|黑罐)/.test(p)) cls = 'heavy';
      else {
        var m = /[\u4e00-\u9fa5]+/.exec(p);
        if (m) {
          var dir = map[m[0]] || '';
          if (/沙/.test(dir)) cls = 'shakou';
          else if (/南/.test(dir)) cls = 'nankou';
          else if (/管内/.test(dir)) cls = 'guanna';
        }
      }
      // 车站名标记为可点击（双击查看径路）
      var link = (clickable && isStationName(p)) ? ' station-link' : '';
      return '<span class="' + cls + link + '" data-station="' +
             escapeHtml(stationOf(p)) + '">' + escapeHtml(p) + '</span>';
    }).join(' ');
  }

  /**
   * 判断片段是否为车站名。
   * 方向库仅 583 条，而地图车站库有 7000+ 条，用方向库匹配会大量漏判。
   * 故改用排除法：非分类词的中文片段即视为车站名。
   * 若地图侧仍查不到，桥接层会收到失败回执并提示，不影响使用。
   */
  var NOT_STATION = {
    '路罐': 1, '自备罐': 1, '黑罐': 1, '到卸': 1, '空车': 1,
    '汽油': 1, '柴油': 1, '原装': 1, '卸空': 1, '循环': 1
  };

  function stationOf(part) {
    var m = /[\u4e00-\u9fa5]+/.exec(part);
    if (!m) return '';
    var name = m[0];
    if (NOT_STATION[name]) return '';
    // 车种字母（C/X/P/G/YW/T/B/D/K/N 开头）不是站名
    if (/^[CXPGYWTBDKN]/.test(name)) return '';
    return name;
  }

  function isStationName(part) { return !!stationOf(part); }

  /* =================== 虚拟股道显示/隐藏 =================== */
  /**
   * 当前应显示的行（连同其在 state.rows 中的原始下标）。
   * 保留原始下标是关键：行上的 data-idx 直接用于 openDetail → state.rows[idx]
   * @returns {Array<{r: Object, idx: number}>}
   */
  function visibleRows() {
    var out = [];
    state.rows.forEach(function (r, idx) {
      if (!state.showVirtual && YardConfig.isVirtual(r.track)) return;
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

  /** 同步工具栏「虚拟股道」开关的文案与高亮态 */
  function syncVirtualBtn() {
    var btn = $('btnToggleVirtual');
    if (!btn) return;
    var on = state.showVirtual;
    btn.textContent = on ? '隐藏非常用股道' : '显示非常用股道';
    btn.classList.toggle('active', on);
    btn.setAttribute('aria-pressed', on ? 'true' : 'false');
    btn.title = (on ? '当前显示' : '当前隐藏') + '不常用股道（X线、YX1、YQX、ZXX、DY、TSY、YH、临时），点击切换';
  }

  /* =================== 渲染主表 =================== */
  function render() {
    // 表头
    $('headRow').innerHTML = COLUMNS.map(function (c) {
      var cls = c.cls || '';
      // 到站列：加 col-flex 类（仅用于允许换行），仍写内联 width，作为独立可拖拽列
      if (c.key === 'dest') cls += ' col-flex';
      var style = ' style="width:' + c.width + 'px"';
      return '<th class="' + cls.trim() + '"' + style + '>' + escapeHtml(c.title) + '</th>';
    }).join('');

    // 表头被 innerHTML 重建，需重新挂载列宽拖拽手柄并恢复记忆列宽
    ColResize.safeGet($('grid')).remount();

    var vis = visibleRows();
    var thr = YardConfig.thresholds;
    var tbody = $('tbody');
    var html = [];

    vis.forEach(function (item) {
      var r = item.r, idx = item.idx;
      var track = r.track;
      var cfg = YardConfig.getTrack(track);
      var zones = YardConfig.getZones(track);

      // 空股道（无车）
      var isBlank = !r.count;

      // 载重超吨
      var overLoad = r.load > thr.overloadTons;
      // 换长超长
      var overLong = r.length > thr.overlong;

      // 车种列：作业区禁入标记
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
        var cls = c.cls || '';
        var style = '';
        var inner;

        if (c.dest) {
          inner = renderDest(v);
          cls += ' dest';
        } else if (c.num) {
          inner = (v === 0 || v === '' || v == null) ? '' : escapeHtml(String(v));
          cls += ' num';
        } else {
          var rawNote = (v == null ? '' : String(v));
          // 注意事项列：聚合时已用 \n 分隔各关键词（超71.86吨 / 扣修 …），
          // 转成 <br> 才能逐条换行；方向列同理（render 内已处理）。
          if (c.key === 'note') {
            inner = escapeHtml(rawNote).replace(/\n/g, '<br>');
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
        if (c.key === 'direction') inner = escapeHtml(v || '').replace(/\n/g, '<br>');

        cells.push('<td class="' + cls + '"' + style + '>' + inner + '</td>');
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
      length: Math.round(tl * 10) / 10,
      load: Math.round(tw * 10) / 10,
      oldCar: told
    };
    var footCells = [];
    COLUMNS.forEach(function (c, i) {
      // 前两列（注意事项 / 股道）是冻结列，合计标签横跨这两列
      if (i === 0) { footCells.push('<td class="col-a" colspan="2">合计</td>'); return; }
      if (i === 1) return;                       // 已被上面的 colspan=2 覆盖
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
  }

  /* =================== 明细抽屉 =================== */
  /** 计算一组明细行的合计：辆数 / 换长 / 总重（总重 = 自重 + 载重，均 1 位小数）。
   *  rows 为 r.raw 的子数组；为空时返回全 0。 */
  function computeTotals(rows) {
    var parseNum = Utils.vbVal;   // 与聚合引擎同用一套取数规则，保证明细与主表口径一致
    var len = 0, selfW = 0, loadW = 0;
    for (var k = 0; k < rows.length; k++) {
      var row = rows[k];
      len += parseNum(row[COL.LEN]);                  // 换长
      selfW += parseNum(row[COL.TARE]);               // 自重
      loadW += parseNum(row[COL.LOAD]);               // 载重
    }
    return {
      count: rows.length,
      length: Math.round(len * 10) / 10,
      selfW: Math.round(selfW * 10) / 10,
      loadW: Math.round(loadW * 10) / 10,
      weight: Math.round((selfW + loadW) * 10) / 10   // 总重 = 自重 + 载重
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
    var t = computeTotals(rows);
    var name = YardConfig.getTrack(r.track);
    $('drawerTitle').innerHTML =
      '<span class="dt-name">' + escapeHtml(name ? name.name : r.track) + ' - </span>' +
      '<span class="dt-total">辆数：' + t.count + '</span>' +
      '<span class="dt-total' + (t.length > 70 ? ' warn' : '') + '">换长：' + t.length.toFixed(1) + '</span>' +
      '<span class="dt-total">自重：' + t.selfW.toFixed(1) + '</span>' +
      '<span class="dt-total">载重：' + t.loadW.toFixed(1) + '</span>' +
      '<span class="dt-total dt-weight' + (t.weight > 5000 ? ' warn' : '') + '">总重：' + t.weight.toFixed(1) + '</span>';
  }
  /** 车种/车号颜色规则（对齐 VBA 显示信息.bas） */
  var carStyle = Utils.carStyle;

  function openDetail(idx) {
    var r = state.rows[idx];
    if (!r) return;
    state.detailIdx = idx;
    state.detailSel = new Set();   // 重置多选（每次打开明细都清空选中）
    updateDetailTitle();           // 初始：按全部行求和（无选中）

    // 列定义是常量（DETAIL_COLS），表头与每一行共用同一份
    $('detailHead').innerHTML = DETAIL_COLS.map(function (c) {
      // 记事列：加 col-flex 类（仅用于允许换行），仍写内联 width，作为独立可拖拽列
      var cls = (c.t === '记事') ? ' col-flex' : '';
      return '<th class="' + cls.trim() + '" style="width:' + c.w + 'px">' + escapeHtml(c.t) + '</th>';
    }).join('') + '<th style="width:60px">停时h</th>';

    // 表头被 innerHTML 重建，需重新挂载列宽拖拽手柄并恢复记忆列宽
    var dt = $('detailTable');
    ColResize.safeGet(dt).remount();

    var list = r.raw || [];
    var base = state.printDate || new Date();
    $('detailBody').innerHTML = list.map(function (row, i) {
      var cs = carStyle(row);
      var tds = DETAIL_COLS.map(function (c) {
        var raw = row[c.col];
        if (c.fmt === 'track') {
          var t = YardConfig.getTrack(raw);
          return '<td class="center">' +
                 escapeHtml(t ? t.name : (raw == null ? '' : raw)) + '</td>';
        }
        if (c.fmt === 'dest') {
          var txt = raw == null ? '' : String(raw).trim();
          if (txt) return '<td class="dest">' + renderDest(txt, true) + '</td>';
          var agg = row.__dest == null ? '' : String(row.__dest).trim();
          if (agg) {
            return '<td class="dest"><span class="derived">' +
                   escapeHtml(agg) + '</span></td>';
          }
          return '<td class="dest"></td>';
        }
        // 车种列
        if (c.col === COL.CARTYPE) {
          var clsA = [c.cls, cs.cls, cs.bg].filter(Boolean).join(' ');
          return '<td' + (clsA ? ' class="' + clsA + '"' : '') + '>' +
                 escapeHtml(raw == null ? '' : raw) + '</td>';
        }
        // 车号列
        if (c.col === COL.CARNO) {
          var clsB = [c.cls, cs.clsN, cs.bgN].filter(Boolean).join(' ');
          return '<td' + (clsB ? ' class="' + clsB + '"' : '') + '>' +
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
      return '<tr data-i="' + i + '">' + tds + '<td class="stay"></td></tr>';
    }).join('');

    // 停时列：需单独计算
    // 复用 Utils.parseArriveTime：兼容 Date 实例、"2026/9/2T08:30:00"、"2026年9月2日" 等写法
    var body = $('detailBody').querySelectorAll('tr');
    list.forEach(function (row, i) {
      var d = Utils.parseArriveTime(row[COL.ARRTIME]);
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

    // 明细打开时按内容自适应列宽（渲染完 body 后测量，逐列 autoFit）
    ColResize.safeGet(dt).reset();

    UI.Drawer.open('drawer');
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

    // 排除补显的罐车分类（路罐/自备罐等），它们不是车站名
    if (NOT_STATION[station] || /^[CXPGYWTBDKN]/.test(station)) {
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
  function bind() {
    on('fileInputMulti', 'change', function (e) {
      var f = e.target.files[0];
      if (f) readAndRender(f, f.name, null);
      e.target.value = '';
    });

    on('btnPickFolder', 'click', function () {
      pickFolder();
    });

    on('btnReload', 'click', function () {
      if (state.dirHandle) loadFromDir(state.dirHandle, false);
      else if (state.fileList && state.fileList.length) {
        readAndRender(state.fileList[0].handle.getFile(), state.fileList[0].name, state.fileList);
      } else toast('请先选择文件夹或文件');
    });

    // 双击行 → 明细
    on('tbody', 'dblclick', function (e) {
      var tr = e.target.closest('tr');
      if (!tr) return;
      openDetail(+tr.getAttribute('data-idx'));
    });

    // 明细中双击车站名 → 打开地图生成径路（起点固定钦州港）
    on('detailBody', 'dblclick', function (e) {
      var sp = e.target.closest('span.station-link');
      if (!sp) return;
      e.stopPropagation();
      openStationMap(sp);
    });

    // 明细中：按下行 → 拖动多选（拖动中实时调整范围，松开确定）；单击 → 切换选中
    var dragSel = { active: false, moved: false, anchor: -1, snap: null, mode: 'add' };
    function selectRow(i, on) {
      if (!state.detailSel) state.detailSel = new Set();
      var tr = $('detailBody').querySelector('tr[data-i="' + i + '"]');
      if (on) { state.detailSel.add(i); if (tr) tr.classList.add('selected'); }
      else { state.detailSel.delete(i); if (tr) tr.classList.remove('selected'); }
    }
    function clearDetailSel() {
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

    // 单击选中
    on('tbody', 'click', function (e) {
      var tr = e.target.closest('tr');
      if (!tr) return;
      var old = $('tbody').querySelector('tr.selected');
      if (old) old.classList.remove('selected');
      tr.classList.add('selected');
      state.selectedIdx = +tr.getAttribute('data-idx');
    });

    // 虚拟股道显示/隐藏开关
    on('btnToggleVirtual', 'click', function () {
      var vids = YardConfig.virtualIds || [];
      if (!vids.length) { toast('配置中没有虚拟股道', 'error'); return; }
      state.showVirtual = !state.showVirtual;
      syncVirtualBtn();
      render();
      toast((state.showVirtual ? '已显示' : '已隐藏') + '不常用股道（共 ' + vids.length + ' 条）', 'ok');
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
    UI.Drawer.register('drawer', { maskId: 'drawerMask' });

    // 各面板的关闭按钮
    on('btnCloseDrawer', 'click', closeDetail);
    on('rptClose', 'click', function () { UI.Modal.close('modal31814'); });
    on('btnSettingsClose', 'click', function () { UI.Modal.close('modalSettings'); });

    // 功能下拉菜单（组件负责展开 / 收起 / 点外部关闭）
    UI.Dropdown('btnMenu', 'menuList', {
      onSelect: function (item, action) {
        if (action === '31814' && typeof window.Report31814 !== 'undefined') {
          window.Report31814.open(state.rawRows, state.currentFile);
        } else if (action === 'settings') {
          UI.Modal.open('modalSettings');
        }
      }
    });

    // 设置：表格字号滑块
    var gridFontSize = $('gridFontSize');
    var gridFontSizeVal = $('gridFontSizeVal');
    if (gridFontSize && gridFontSizeVal) {
      var savedFs = Store.get('gridFontSize', '');
      if (savedFs) {
        gridFontSize.value = savedFs;
        gridFontSizeVal.textContent = savedFs + 'px';
        $('grid').style.fontSize = savedFs + 'px';
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
        $('grid').style.fontSize = v + 'px';
        saveFontSize(v);
        highlightTick(v);
      });
      highlightTick(gridFontSize.value);
    }

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
        pickFolder();
      });
    }

    // 表头列宽拖动（persistKey 用于本地记忆，刷新/重渲染不丢失）
    ColResize.enable($('grid'), {
      persistKey: 'zhancun.grid.cols',
      onResize: function (th, w) {
        // 首列宽变化 → 同步「股道」列的冻结偏移，避免两列之间露缝
        if (th.classList.contains('col-a')) $('grid').style.setProperty('--col-a-w', w + 'px');
      }
    });
    ColResize.enable($('detailTable'), { persistKey: 'zhancun.detail.cols' });
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

  /* =================== 初始化 =================== */
  function init() {
    // 方向库（惰性单例：全局只解析一次，报表模块共用同一份实例）
    if (typeof window.DirectionData !== 'string') {
      toast('方向库未加载，到站着色将不可用', 'error');
    }
    state.dirIndex = Aggregate.getDirectionIndex();

    bind();
    syncVirtualBtn();

    // 先渲染空框架：让页面一打开就呈现完整股道清单，便于核对配置
    renderEmpty();

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
      rememberFolder(h);
      // 静默恢复：权限未授予时不弹窗，等用户点击
      if (h.queryPermission) {
        return h.queryPermission({ mode: 'read' }).then(function (p) {
          if (p === 'granted') return loadFromDir(h, false);
          $('stMsg').textContent = '已记住数据文件夹，点「重新读取」以载入';
        });
      }
      return loadFromDir(h, false);
    }).catch(function () {
      $('stMsg').textContent = '读取上次的文件夹失败，请点「选择数据文件夹」重新选择';
      syncPickFolderBtn(true);   // 句柄不可用，重新亮出入口
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else init();
})();
