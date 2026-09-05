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
 *   [对外接口]      文件末尾  window.YardApp —— 数据 / 配置 / 渲染 / 交互的统一出口，
 *                        新增模块一律通过它复用主表能力（详见该区块注释）
 *
 * 说明：state 为 IIFE 内私有对象（约 51 处引用），故未做文件拆分；
 *       如需拆分，数据源区块（L77-301）耦合最弱，最优先。
 * =====================================================================
 */
(function (global) {
  'use strict';

  /* ============================ 列定义 ============================ */
  var COLUMNS = [
    { key: 'note',      title: '注意事项',       width: 112, cls: 'col-a' },
    // 分组合并列：在股道列前，按分组跨行合并（rowspan），仅每组首行输出单元格；文字竖排
    { key: 'group',     title: '',               width: 32,  cls: 'col-b-group' },
    { key: 'track',     title: '股道',           width: 66,  cls: 'col-b track' },
    { key: 'direction', title: '方向',           width: 64 },
    { key: 'count',     title: '车数',           width: 48,  num: true, cls: 'mid' },
    { key: 'carTypes',  title: '车种',           width: 124 },
    { key: 'length',    title: '换长',           width: 58,  num: true, cls: 'mid' },
    { key: 'dest',      title: '车辆信息',       width: 340, dest: true },
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
    // 车号在车种之前（现场按车号点车，先找号再看车型）
    { col: Aggregate.COL.CARNO, t: '车号', w: 78 }, { col: Aggregate.COL.CARTYPE, t: '车种', w: 68 },
    { col: Aggregate.COL.TARE,    t: '自重', w: 52 }, { col: Aggregate.COL.LEN,   t: '换长', w: 50, cls: 'mid' },
    { col: Aggregate.COL.LOAD,    t: '载重', w: 56, cls: 'mid' },
    { col: Aggregate.COL.DEST,    t: '到站', w: 120, disp: 'processed' },
    { col: Aggregate.COL.DIR,     t: '方向', w: 42 }, { col: Aggregate.COL.GOODS, t: '品名', w: 90 },
    { col: Aggregate.COL.FROM,    t: '发站', w: 110, disp: 'raw' },
    { col: Aggregate.COL.NOTE,    t: '记事', w: 170 },
    { col: Aggregate.COL.TRAIN,   t: '车次', w: 62 }, { col: Aggregate.COL.CONSIGNEE, t: '收货人', w: 100 }, { col: Aggregate.COL.ARRTIME, t: '到达时间', w: 128 }
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
  /**
   * 判断片段是否为「卸车地点」（到卸车的细化去向，如 永鑫/货场/天盛/港务局）。
   * 列表来自 Store.unloadSpots（用户在设置里增删），与 aggregate 段2 共用同一份配置。
   * 未配置时回落默认 4 个，保证无设置也能正确加粗。
   *
   * 到站串片段是"永鑫50"（地点+车数），要先去掉尾部数字再匹配地点词。
   */
  function isUnloadSpot(p) {
    if (!p) return false;
    var name = String(p).replace(/\d+$/, '');   // 去尾数：永鑫50 → 永鑫
    if (!name) return false;
    var list = (Store.get && Store.get('unloadSpots', null)) ||
               ['永鑫', '货场', '天盛', '港务局'];
    return list.indexOf(name) >= 0;
  }

  /**
   * 判断片段是否为「黑罐细化子类」（G7 罐车按收货人识别，如 中粮/外运）。
   * 列表来自 Store.blackTankSpots，与 aggregate 段1 共用同一份配置。
   * 未配置时回落默认 中粮/外运。
   */
  function isBlackTankSpot(p) {
    if (!p) return false;
    var name = String(p).replace(/\d+$/, '');
    if (!name) return false;
    var list = (Store.get && Store.get('blackTankSpots', null)) ||
               ['中粮', '外运'];
    return list.indexOf(name) >= 0;
  }

  /**
   * 取片段的车型部分：到站串是「分类+车数」，如 "P5" / "DK2" / "YW3"。
   * 只取前导字母交给车型高亮配置匹配；中文片段（车站名、到卸、黑罐）返回 ''。
   */
  function typeKeyOf(p) {
    var m = /^[A-Za-z]+/.exec(String(p == null ? '' : p));
    return m ? m[0] : '';
  }

  function renderDest(text, clickable) {
    if (!text) return '';
    var parts = String(text).trim().split(/\s+/).filter(Boolean);
    var map = state.dirIndex.map;
    return parts.map(function (p) {
      var cls = '';
      // ① 车型高亮（「设置 → 车型高亮」可编辑的颜色/加粗）。
      //    主表只挂「非平板车」的 ctc 类（颜色=字体色）；平板车（X/NX）的底色
      //    规则仅作用于明细车种列，主表到站列不挂底色 —— 见用户对齐 VBA 的约束。
      var tk = typeKeyOf(p);
      if (tk) {
        var m = Utils.carTypeMatch(tk);
        if (m && !m.isFlatbed) cls = Utils.carTypeClass(tk);
      }
      // ② 未命中配置 → 回落 VBA 原有规则（显示信息.bas「标记到站方向颜色」：
      //    YW/D/P → 红加粗，到卸/黑罐 → 黑加粗）。配置里关掉的项会走到这里。
      if (!cls) {
        if (/^(YW|D|P)/.test(p)) cls = 'danger';
        else if (/^(到卸|黑罐)/.test(p)) cls = 'heavy';
        // 卸车地点（到卸的细化，如 永鑫/货场/天盛/港务局，或用户在设置里增删的）
        // 与"到卸"同款黑加粗。列表来自 Store.unloadSpots，与 aggregate 段2 一致。
        else if (isUnloadSpot(p)) cls = 'heavy';
        // 黑罐细化子类（G7 罐按收货人识别，如 中粮/外运），与"黑罐"同款加粗
        else if (isBlackTankSpot(p)) cls = 'heavy';
        else {
          var m = /[\u4e00-\u9fa5]+/.exec(p);
          if (m) {
            var dir = map[m[0]] || '';
            if (/沙/.test(dir)) cls = 'shakou';
            else if (/南/.test(dir)) cls = 'nankou';
            else if (/管内/.test(dir)) cls = 'guanna';
          }
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
        var cls = c.cls || '';
        var style = '';
        var inner;

        if (c.dest) {
          inner = renderDest(v);
          cls += ' dest';
        } else if (c.num) {
          // 换长保留 1 位小数（如 0.0）；其余数值列维持原样
          if (c.key === 'length') {
            inner = (v === 0 || v === '' || v == null) ? '' : Number(v).toFixed(1);
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
        if (c.key === 'direction') inner = escapeHtml(v || '').replace(/\n/g, '<br>');

        // 分组「合并列」：在股道列之前，按分组跨行合并（rowspan）。
        // 仅每组首行输出带 rowspan 的分组单元格，组内后续行不输出（由 rowspan 覆盖）。
        if (c.key === 'group') {
          var sp = spans[n];
          if (sp.start) {
            // 用分组自带的 color 做左侧色条 + 文字着色，醒目区分到发线/调车线/虚拟场等
            var gc = cfg ? cfg.groupColor : '#888';
            cells.push('<td class="col-b-group grp" rowspan="' + sp.span + '" ' +
                       'style="border-left:3px solid ' + gc + ';color:' + gc + '">' +
                       escapeHtml(sp.group) + '</td>');
          }
          return;   // 非首行：被上方 rowspan 覆盖，不输出 td
        }

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
        if (c.fmt === 'track') {
          var t = YardConfig.getTrack(raw);
          return '<td class="center">' +
                 escapeHtml(t ? t.name : (raw == null ? '' : raw)) + '</td>';
        }
        if (c.disp) {
          // 到站(disp:'processed')默认显示原始站名，仅当 opts.destProcessed 为真才用处理后的 __dest；
          // 发站(disp:'raw')永远用原始发站，从根上杜绝被聚合到站（如「货场」）误填。
          var useProcessed = c.disp === 'processed' && opts.destProcessed;
          var val = useProcessed && row.__dest != null ? String(row.__dest).trim()
                  : (raw == null ? '' : String(raw).trim());
          if (val) return '<td class="dest">' + renderDest(val, true) + '</td>';
          // 原始站名为空 → 回退派生到站（聚合串，如「三街3 麻尾2」），挂 station-link 供双击开地图
          var agg = row.__dest == null ? '' : String(row.__dest).trim();
          if (agg) {
            var dStation = stationOf(agg);
            var dCls = dStation ? 'derived station-link' : 'derived';
            var dAttr = dStation ? ' data-station="' + escapeHtml(dStation) + '"' : '';
            return '<td class="dest"><span class="' + dCls + '"' + dAttr + '>' +
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
    updateDetailTitle();           // 初始：按全部行求和（无选中）

    renderDetailRows(r.raw || [], {
      head: $('detailHead'),
      body: $('detailBody'),
      table: $('detailTable')
    }, {
      rowAttr: function (row, i) { return ' data-i="' + i + '"'; }
    });

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
      if (f) readAndRender(f, f.name, null);
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
      readAndRender(f, f.name, null).then(function () {
        // 以文件方式载入时隐藏文件切换器与刷新（无目录可扫描）
        var sw = $('fileSwitcher');
        if (sw) sw.style.display = 'none';
        $('stMsg').textContent = '已载入文件：' + f.name;
      }).catch(function () { /* 错误已在 readAndRender 内提示 */ });
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

    // 单击选中（作业区横幅行不参与选中，否则会被高亮且 selectedIdx 变为 NaN）
    on('tbody', 'click', function (e) {
      var tr = e.target.closest('tr');
      if (!tr || tr.classList.contains('area-banner')) return;
      var old = $('tbody').querySelector('tr.selected');
      if (old) old.classList.remove('selected');
      tr.classList.add('selected');
      state.selectedIdx = +tr.getAttribute('data-idx');
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
    UI.Drawer.register('drawer', { maskId: 'drawerMask' });
    UI.Drawer.register('searchDrawer', { maskId: 'searchMask' });

    // 各面板的关闭按钮
    on('btnCloseDrawer', 'click', closeDetail);
    on('rptClose', 'click', function () { UI.Modal.close('modal31814'); });
    on('btnSettingsClose', 'click', function () { UI.Modal.close('modalSettings'); });
    on('prodClose', 'click', function () { UI.Modal.close('modalProductivity'); });

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

    // 卸车地点（aggregate 段2 读取，默认 永鑫/货场/天盛/港务局）
    initSpotConfig('unloadSpots', 'btnAddUnloadSpot', 'unloadSpots',
                   ['永鑫', '货场', '天盛', '港务局']);
    // 黑罐识别（aggregate 段1 G7 罐按收货人识别，默认 中粮/外运）
    initSpotConfig('blackTankSpots', 'btnAddBlackTankSpot', 'blackTankSpots',
                   ['中粮', '外运']);

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

  /* =================== 初始化 =================== */
  function init() {
    initModalDrag();
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