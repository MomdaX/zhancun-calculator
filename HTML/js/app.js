/**
 * app.js —— 股道存车主程序
 *
 * 无服务器设计：
 *   - 方向库、股道配置通过 <script src> 加载（file:// 下唯一可靠方式）
 *   - xls 通过 File System Access API 读取，目录句柄存 IndexedDB 实现"打开即自动读取"
 *   - 不支持该 API 的浏览器自动降级为 <input type="file"> 手动选择
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
  var DETAIL_COLS = [
    { i: 0, t: '股道', w: 54, fmt: 'track' },
    { i: 1, t: '顺', w: 34 },
    { i: 2, t: '车种', w: 68 }, { i: 3, t: '车号', w: 78 },
    { i: 4, t: '自重', w: 52 }, { i: 5, t: '换长', w: 50, cls: 'mid' },
    { i: 6, t: '载重', w: 56, cls: 'mid' },
    { i: 7, t: '到站', w: 120, fmt: 'dest' },
    { i: 8, t: '方向', w: 42 }, { i: 9, t: '品名', w: 90 },
    { i: 10, t: '发站', w: 110, fmt: 'dest' },
    { i: 13, t: '记事', w: 170 },
    { i: 14, t: '车次', w: 62 }, { i: 15, t: '到达时间', w: 128 }
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
    printDate: null,       // 数据源打印日期（作为停时基准）
    showVirtual: true      // 是否显示虚拟股道（X 线）
  };

  var $ = function (id) { return document.getElementById(id); };

  /**
   * 安全绑定事件：元素不存在时静默跳过。
   * 避免某个 DOM 缺失导致 bind() 整体中断、后续功能全部失效。
   */
  function on(id, evt, fn) {
    var el = $(id);
    if (el) el.addEventListener(evt, fn);
    return el;
  }

  function escapeHtml(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function toast(msg, type) {
    var t = $('toast');
    t.textContent = msg;
    t.className = 'toast show' + (type ? ' ' + type : '');
    clearTimeout(t._timer);
    t._timer = setTimeout(function () { t.className = 'toast'; }, 2600);
  }

  function loading(show, text) {
    $('loadingText').textContent = text || '正在解析…';
    $('loading').className = show ? 'loading show' : 'loading';
  }

  /* =================== IndexedDB：记住目录句柄 =================== */
  var DB_NAME = 'YardStorageDB', STORE = 'handles';

  function openDB() {
    return new Promise(function (res, rej) {
      var q = indexedDB.open(DB_NAME, 1);
      q.onupgradeneeded = function () {
        if (!q.result.objectStoreNames.contains(STORE)) q.result.createObjectStore(STORE);
      };
      q.onsuccess = function () { res(q.result); };
      q.onerror = function () { rej(q.error); };
    });
  }

  function idbSet(k, v) {
    return openDB().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction(STORE, 'readwrite');
        tx.objectStore(STORE).put(v, k);
        tx.oncomplete = function () { res(); };
        tx.onerror = function () { rej(tx.error); };
      });
    }).catch(function () { /* 存储失败不影响主流程 */ });
  }

  function idbGet(k) {
    return openDB().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction(STORE, 'readonly');
        var rq = tx.objectStore(STORE).get(k);
        rq.onsuccess = function () { res(rq.result); };
        rq.onerror = function () { rej(rq.error); };
      });
    }).catch(function () { return null; });
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

  /** 选择文件夹（File System Access API） */
  function pickFolder() {
    if (!window.showDirectoryPicker) {
      toast('当前浏览器不支持文件夹选择，请点「选择文件」', 'error');
      $('fileInputMulti').click();
      return;
    }
    window.showDirectoryPicker({ id: 'yardXls', mode: 'read' })
      .then(function (h) {
        return idbSet('xlsDir', h).then(function () {
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
      wrap.innerHTML = '<select id="fileSwitcher" class="btn" style="max-width:200px"></select> ';
      $('btnReload').parentNode.insertBefore(wrap, $('btnReload').nextSibling);
      sel = $('fileSwitcher');
      sel.addEventListener('change', function () {
        var f = state.fileList[sel.selectedIndex];
        if (f) readAndRender(f.handle.getFile(), f.name, state.fileList);
      });
    }
    var fmt = function (d) {
      return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
             ('0' + d.getDate()).slice(-2) + ' ' +
             ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2);
    };

    // 先算出目标选中项，写入 innerHTML 后立即还原
    var idx = -1;
    for (var i = 0; i < files.length; i++) {
      if (files[i].name === activeName) { idx = i; break; }
    }
    if (idx < 0) idx = sel.selectedIndex;   // 兜底：保持重建前的选择

    sel.innerHTML = files.map(function (f, i) {
      return '<option value="' + i + '">' + escapeHtml(f.name) +
             '（' + fmt(new Date(f.lastModified)) + '）</option>';
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
      return '<th class="' + (c.cls || '') + '" style="width:' + c.width + 'px">' +
             escapeHtml(c.title) + '</th>';
    }).join('');

    var thr = YardConfig.thresholds;
    var tbody = $('tbody');
    var html = [];

    visibleRows().forEach(function (item) {
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
          if (/汽油|航煤/.test(String(r.raw[k][9] || ''))) { hasOil = true; break; }
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
          inner = escapeHtml(v == null ? '' : v);
          if (!c.cls && !c.num) cls += ' center';
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
    visibleRows().forEach(function (item) {
      var r = item.r;
      tc += r.count || 0; tl += r.length || 0; tw += r.load || 0; told += r.oldCar || 0;
    });
    $('tfoot').innerHTML = '<tr>' +
      '<td class="col-a" colspan="2">合计</td>' +
      '<td></td>' +
      '<td class="num mid">' + tc + '</td>' +
      '<td></td>' +
      '<td class="num mid">' + (Math.round(tl * 10) / 10) + '</td>' +
      '<td id="footTank"></td>' +
      '<td></td><td></td><td></td>' +
      '<td class="num mid">' + (Math.round(tw * 10) / 10) + '</td>' +
      '<td class="num mid">' + told + '</td>' +
      '</tr>';

    // 罐车结存（自备罐/路罐）
    var zb = 0, lg = 0;
    visibleRows().forEach(function (item) {
      var re = /(自备罐|路罐)(\d+)/g, m;
      var d = item.r.dest || '';
      while ((m = re.exec(d))) {
        if (m[1] === '自备罐') zb += +m[2]; else lg += +m[2];
      }
    });
    $('footTank').textContent = zb + '(自)/' + lg + '(路)';

    // 状态栏
    $('stTrack').textContent = visibleRows().filter(function (item) { return item.r.count; }).length;
    $('stCount').textContent = tc;
    $('stLen').textContent = Math.round(tl * 10) / 10;
    $('stLoad').textContent = Math.round(tw * 10) / 10;
    $('stOld').textContent = told;
  }

  /* =================== 明细抽屉 =================== */
  function openDetail(idx) {
    var r = state.rows[idx];
    if (!r) return;
    state.detailIdx = idx;
    var cfgTrack = r.track;
    var name = YardConfig.getTrack(cfgTrack);
    $('drawerTitle').textContent = (name ? name.name : cfgTrack) + ' - 明细（' + (r.count || 0) + ' 辆）';

    $('detailHead').innerHTML = DETAIL_COLS.map(function (c) {
      return '<th style="width:' + c.w + 'px">' + escapeHtml(c.t) + '</th>';
    }).join('') + '<th style="width:60px">停时h</th>';

    var list = r.raw || [];
    var base = state.printDate || new Date();
    $('detailBody').innerHTML = list.map(function (row) {
      var tds = DETAIL_COLS.map(function (c) {
        var raw = row[c.i];
        if (c.fmt === 'track') {
          var t = YardConfig.getTrack(raw);
          return '<td class="center">' +
                 escapeHtml(t ? t.name : (raw == null ? '' : raw)) + '</td>';
        }
        if (c.fmt === 'dest') {
          // 到站列：优先显示原始车站名；原始为空时补显聚合分类（如 路罐/自备罐），
          // 以便空罐车也能看出属性。补显部分加 class 便于区分。
          var txt = raw == null ? '' : String(raw).trim();
          if (txt) return '<td class="dest">' + renderDest(txt, true) + '</td>';
          var agg = row.__dest == null ? '' : String(row.__dest).trim();
          if (agg) {
            return '<td class="dest"><span class="derived">' +
                   escapeHtml(agg) + '</span></td>';
          }
          return '<td class="dest"></td>';
        }
        return '<td' + (c.cls ? ' class="' + c.cls + '"' : '') + '>' +
               escapeHtml(raw == null ? '' : raw) + '</td>';
      }).join('');
      return '<tr>' + tds + '<td class="stay"></td></tr>';
    }).join('');

    // 停时列：需单独计算
    var body = $('detailBody').querySelectorAll('tr');
    list.forEach(function (row, i) {
      var arr = row[15];
      var d = null;
      if (arr instanceof Date) d = arr;
      else if (arr) {
        var m = /(\d{4})[-/](\d{1,2})[-/](\d{1,2})[\sT]*(\d{1,2})?:?(\d{1,2})?/.exec(String(arr));
        if (m) d = new Date(+m[1], +m[2] - 1, +m[3], m[4] ? +m[4] : 0, m[5] ? +m[5] : 0);
      }
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

    $('drawer').classList.add('show');
    $('drawerMask').classList.add('show');
  }

  function closeDetail() {
    $('drawer').classList.remove('show');
    $('drawerMask').classList.remove('show');
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

  /* =================== 表头列宽拖拽 =================== */
  /**
   * 表头列宽可拖动调整。
   * 拖动标记由 CSS 伪元素 ::after 绘制（不新增 DOM），此处只做命中判定与宽度计算。
   * @param {HTMLTableElement} table
   * @param {Function} [onResize] 宽度变化时回调 (th, width)
   */
  function enableColResize(table, onResize) {
    var thead = table && table.tHead;
    if (!thead) return;

    var HIT = 6;   // 拖动命中宽度，需与 CSS 中 ::after 的 width 一致
    var MIN = 30;  // 最小列宽
    var drag = null;

    // Canvas 测文本宽度（复用一个离屏 canvas）
    var measureCvs = document.createElement('canvas');
    var measureCtx = measureCvs.getContext('2d');

    function findTh(e) {
      var th = e.target && e.target.closest ? e.target.closest('th') : null;
      return th && th.closest('thead') === thead ? th : null;
    }

    /** 鼠标是否落在 th 右边缘的拖动区内 */
    function atEdge(th, x) {
      var r = th.getBoundingClientRect();
      return x >= r.right - HIT && x <= r.right + 1;
    }

    function clearMark() {
      var prev = thead.querySelector('th.col-resizing');
      if (prev) prev.classList.remove('col-resizing');
    }

    /** 取消拖拽（双击时可能已由首次单击触发了拖拽） */
    function cancelDrag() {
      if (!drag) return;
      drag.th.classList.remove('col-resizing');
      drag = null;
      document.body.classList.remove('col-resize-active');
    }

    /** 列索引（th 在父行中的位置） */
    function colIdx(th) {
      return Array.prototype.indexOf.call(th.parentNode.children, th);
    }

    /** 测量文本宽度 */
    function textWidth(text, font) {
      measureCtx.font = font;
      return measureCtx.measureText(text).width;
    }

    /** 双击分隔线 → 自动调整为最适列宽 */
    function autoFit(th) {
      var idx = colIdx(th);
      var font = window.getComputedStyle(th).font;
      var maxW = textWidth(th.textContent.trim(), font) + 20;

      // 遍历所有 tbody 行中该列单元格
      var bodies = table.tBodies;
      for (var b = 0; b < bodies.length; b++) {
        var rows = bodies[b].rows;
        for (var r = 0; r < rows.length; r++) {
          var cell = rows[r].cells[idx];
          if (!cell) continue;
          var cf = window.getComputedStyle(cell).font;
          var w = textWidth(cell.textContent.trim(), cf) + 20;
          if (w > maxW) maxW = w;
        }
      }

      // tfoot（无 colspan 的列才纳入，避免索引错位）
      if (table.tFoot) {
        var footRows = table.tFoot.rows;
        for (var fr = 0; fr < footRows.length; fr++) {
          var fc = footRows[fr].cells[idx];
          if (fc && fc.colSpan === 1) {
            var ff = window.getComputedStyle(fc).font;
            var fw = textWidth(fc.textContent.trim(), ff) + 20;
            if (fw > maxW) maxW = fw;
          }
        }
      }

      var finalW = Math.max(MIN, Math.round(maxW));
      th.style.width = finalW + 'px';
      if (onResize) onResize(th, finalW);
    }

    thead.addEventListener('mousemove', function (e) {
      if (drag) return;
      var th = findTh(e);
      clearMark();
      if (th && atEdge(th, e.clientX)) th.classList.add('col-resizing');
    });

    thead.addEventListener('mouseleave', clearMark);

    thead.addEventListener('mousedown', function (e) {
      if (e.button !== 0 || drag) return;
      var th = findTh(e);
      if (!th || !atEdge(th, e.clientX)) return;
      e.preventDefault();
      drag = { th: th, startX: e.clientX, startW: th.getBoundingClientRect().width };
      th.classList.add('col-resizing');
      document.body.classList.add('col-resize-active');
    });

    thead.addEventListener('dblclick', function (e) {
      var th = findTh(e);
      if (!th || !atEdge(th, e.clientX)) return;
      e.preventDefault();
      cancelDrag();
      autoFit(th);
    });

    document.addEventListener('mousemove', function (e) {
      if (!drag) return;
      var w = Math.max(MIN, Math.round(drag.startW + (e.clientX - drag.startX)));
      drag.th.style.width = w + 'px';
      if (onResize) onResize(drag.th, w);
    });

    document.addEventListener('mouseup', function () {
      if (!drag) return;
      drag.th.classList.remove('col-resizing');
      drag = null;
      document.body.classList.remove('col-resize-active');
    });
  }

  /* =================== 事件绑定 =================== */
  function bind() {
    on('btnPickFolder', 'click', pickFolder);

    on('btnPickFile', 'click', function () {
      if (window.showOpenFilePicker) {
        window.showOpenFilePicker({
          multiple: false,
          types: [{ description: 'Excel 文件', accept: { 'application/vnd.ms-excel': ['.xls', '.xlsx'] } }]
        }).then(function (hs) {
          return readAndRender(hs[0].getFile(), hs[0].name, null);
        }).catch(function (e) {
          if (e && e.name === 'AbortError') return;
          $('fileInput').click();
        });
      } else {
        $('fileInput').click();
      }
    });

    on('fileInput', 'change', function (e) {
      var f = e.target.files[0];
      if (f) readAndRender(f, f.name, null);
      e.target.value = '';
    });

    on('fileInputMulti', 'change', function (e) {
      var f = e.target.files[0];
      if (f) readAndRender(f, f.name, null);
      e.target.value = '';
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

    on('btnCloseDrawer', 'click', closeDetail);
    on('drawerMask', 'click', closeDetail);
    on('btnPrevTrack', 'click', function () { stepDetail(-1); });
    on('btnNextTrack', 'click', function () { stepDetail(1); });

    // 功能下拉菜单
    var btnMenu = $('btnMenu');
    var menuList = $('menuList');
    if (btnMenu && menuList) {
      btnMenu.addEventListener('click', function (e) {
        e.stopPropagation();
        menuList.style.display = menuList.style.display === 'none' ? 'block' : 'none';
      });
      menuList.addEventListener('click', function (e) {
        var item = e.target.closest('.menu-item');
        if (!item) return;
        menuList.style.display = 'none';
        var action = item.getAttribute('data-action');
        if (action === '31814' && typeof window.Report31814 !== 'undefined') {
          window.Report31814.open(state.rawRows, state.currentFile);
        }
      });
      document.addEventListener('click', function () { menuList.style.display = 'none'; });
    }

    // 31814 浮窗关闭
    var rptClose = $('rptClose');
    var modal31814 = $('modal31814');
    if (rptClose && modal31814) {
      rptClose.addEventListener('click', function () { modal31814.classList.remove('show'); });
      modal31814.addEventListener('click', function (e) {
        if (e.target === modal31814) modal31814.classList.remove('show');
      });
    }

    document.addEventListener('keydown', function (e) {
      if (e.key === 'Escape') {
        closeDetail();
        if (modal31814) modal31814.classList.remove('show');
      }
    });

    // 表头列宽拖动
    enableColResize($('grid'), function (th, w) {
      // 首列宽变化 → 同步「股道」列的冻结偏移，避免两列之间露缝
      if (th.classList.contains('col-a')) $('grid').style.setProperty('--col-a-w', w + 'px');
    });
    enableColResize($('detailTable'));
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
    // 方向库
    if (typeof window.DirectionData !== 'string') {
      toast('方向库未加载，到站着色将不可用', 'error');
      state.dirIndex = { map: {}, stations: [] };
    } else {
      state.dirIndex = Aggregate.buildDirectionIndex(window.DirectionData);
    }

    bind();
    syncVirtualBtn();

    // 先渲染空框架：让页面一打开就呈现完整股道清单，便于核对配置
    renderEmpty();

    // 恢复上次选择的文件夹，自动读取最新 xls
    if (!window.showDirectoryPicker) {
      $('btnPickFolder').style.display = 'none';
      $('stMsg').textContent = '当前浏览器不支持自动读取，请点「选择文件」';
      return;
    }

    idbGet('xlsDir').then(function (h) {
      if (!h) {
        $('stMsg').textContent = '首次使用：请点「选择数据文件夹」，之后将自动读取最新文件';
        return;
      }
      state.dirHandle = h;
      // 静默恢复：权限未授予时不弹窗，等用户点击
      if (h.queryPermission) {
        return h.queryPermission({ mode: 'read' }).then(function (p) {
          if (p === 'granted') return loadFromDir(h, false);
          $('stMsg').textContent = '已记住数据文件夹，点「重新读取」以载入';
        });
      }
      return loadFromDir(h, false);
    }).catch(function () {
      $('stMsg').textContent = '请点「选择数据文件夹」开始';
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else init();
})();