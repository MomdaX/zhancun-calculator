/**
 * data-source.js —— 数据源区块（从 app.js 抽出，P2-7 拆分第一步）
 *
 * 职责：File System Access API 权限、IndexedDB/文件夹读取、xls 解析与聚合、多文件切换。
 * 无 DOM 渲染逻辑；通过全局桥接访问 app.js 的共享资源：
 *   - state        （全局，app.js 在 init 时挂 global.state）
 *   - render       （全局，app.js 在 init 时挂 global.render）
 *   - loading      （全局，app.js 在 init 时挂 global.loading）
 *   - syncPickFolderBtn（全局，app.js 在 init 时挂 global.syncPickFolderBtn）
 * 其余依赖（Utils / Store / Aggregate / YardConfig / XLSX / document / window）均为已有全局，
 * 本文件仅重声明常用别名，使下方函数体与原 app.js 完全一致、零逻辑改动。
 *
 * 加载顺序：须在 app.js 之前（index.html 中置于 aggregate.js 之后、app.js 之前）。
 */
(function (global) {
  'use strict';

  // 重声明 app.js 顶部的常用别名，使下方函数体与原文件零差异
  var $ = Utils.$;
  var on = Utils.on;
  var escapeHtml = Utils.escapeHtml;
  var toast = Utils.toast;
  var COL = Aggregate.COL;

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

  // 暴露为全局，供 app.js 通过 DataSource.* 调用（依赖注入：state/render/loading/syncPickFolderBtn 由 app.js 挂到 global）
  global.DataSource = {
    ensurePerm: ensurePerm,
    listXlsInDir: listXlsInDir,
    rememberFolder: rememberFolder,
    pickFolder: pickFolder,
    loadFromDir: loadFromDir,
    readAndRender: readAndRender,
    renderFileSwitcher: renderFileSwitcher
  };

})(window);
