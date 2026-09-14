/**
 * settings.js —— 设置面板（「功能 ▾ → 设置」浮窗里的全部配置项）
 * ============================================================================
 * 从 app.js 的 bind() 里整块剥离，收拢「设置」相关的一切：
 *   · 表格字号滑块（主页表格 + 明细抽屉同步生效）
 *   · 默认数据文件夹（显示 / 修改入口）
 *   · 可增删标签列表（卸车地点 / 黑罐识别共用 initSpotConfig）
 *   · 货物推算重量（载重缺失时按记事栏货物名称匹配预设重量，列表可增删改）
 *   · 车型高亮配置（对齐 VBA 显示信息.bas）
 *   · recompute：配置变更后重跑聚合并刷新主表
 *   · modalSettings 浮窗注册（打开即刷新文件夹路径显示）
 *
 * 与 app.js 的关系（依赖注入 + 单向调用）：
 *   app.js 在 bind() 时调用 Settings.init，注入 state / render / updateDetailTitle /
 *   openDetail / defaultEstGoods；本模块不反向触碰 app.js 的私有状态。
 *   DEFAULT_UNLOAD_SPOTS / DEFAULT_BLACK_TANK_SPOTS 由 dest-color.js 挂 global，直接取用。
 *
 * 加载顺序：须在 utils / store / ui / dest-color / data-source / aggregate / track.config
 *           之后、app.js 之前（app.js 的 bind 会调用 Settings.init）。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var Utils = global.Utils;
  var $ = Utils.$;
  var toast = Utils.toast;
  var escapeHtml = Utils.escapeHtml;
  var Store = global.Store;
  var UI = global.UI;
  var DataSource = global.DataSource;
  var Aggregate = global.Aggregate;
  var YardConfig = global.YardConfig;

  var DEFAULT_UNLOAD_SPOTS = global.DEFAULT_UNLOAD_SPOTS || [];
  var DEFAULT_BLACK_TANK_SPOTS = global.DEFAULT_BLACK_TANK_SPOTS || [];

  /* app.js 注入的能力，见 init(deps) 的说明 */
  var state, render, updateDetailTitle, openDetail, DEFAULT_EST_GOODS;
  var inited = false;

  /* =================== 表格字号（设置滑块 → 主表 + 明细同步） =================== */

  function applyTableFontSize(v) {
    $('grid').style.fontSize = v + 'px';
    var detailTable = $('detailTable');
    if (detailTable) {
      detailTable.style.fontSize = v + 'px';   // 覆盖 CSS 中 table.detail 的固定字号
    }
  }

  /* =================== 默认数据文件夹（设置里的路径显示与修改入口） =================== */

  function refreshFolderPath() {
    var folderPath = $('folderPath');
    if (folderPath) {
      var name = Store.get(Store.KEYS.folderName, '');
      folderPath.textContent = name || '未设置';
      folderPath.title = name || '';
    }
  }

  /* =================== 配置变更 → 重跑聚合刷新主表 =================== */

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

  /* =================== 「可增删标签列表」共用实现 =================== */

  /**
   * 读取「列表型」配置：存在且为数组 → 返回；否则 null（调用方 `|| 默认列表` 回落）。
   * 与 Store.getList 的区别：本函数保留「空数组」语义——存了 [] 就返回 []，不回落默认值。
   */
  function readListOrNull(storeKey) {
    var list = Store.get(storeKey, null);
    return Array.isArray(list) ? list : null;
  }

  /**
   * 生成一个「可删除标签」项（卸车地点 / 黑罐识别 / 货物推算重量 三处列表共用同一结构）。
   * @param labelHtml 已转义的标签内容（各处内容形态不同，由调用方转义后传入）
   * @param dataName  该项名称（原样传入，函数内转义后写入 data-name）
   * @param {string} [title] 可选 title 提示（原样传入，函数内转义）
   */
  function spotItemHtml(labelHtml, dataName, title) {
    return '<span class="spot-item" data-name="' + escapeHtml(dataName) + '"' +
             (title ? ' title="' + escapeHtml(title) + '"' : '') + '>' +
             labelHtml +
             '<button class="spot-del" title="删除" aria-label="删除">×</button>' +
           '</span>';
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

    function getList() { return readListOrNull(storeKey); }
    function renderList() {
      var list = getList() || defaultList;
      var html = list.map(function (name) {
        return spotItemHtml(escapeHtml(name), name);
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

  /* =================== 货物推算重量（设置里的可编辑列表） =================== */

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

    function getList() { return readListOrNull(storeKey); }
    function renderList() {
      var list = getList() || defaultList;
      el.innerHTML = list.map(function (item) {
        return spotItemHtml(escapeHtml(item.name) + ':' + escapeHtml(item.weight), item.name, '双击修改');
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
      setHint(n + ' · ' + Utils.round1(w) + ' 吨');
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
      var rounded = Utils.round1(weight);
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

  /* =================== 车型高亮（设置里的可编辑列表） =================== */

  // 车型高亮：基于 Utils.getCarTypeConfig 在设置面板渲染可编辑列表
  function initCarTypeConfig() {
    if (!Utils.applyCarTypeStyles) return;
    Utils.applyCarTypeStyles(); // 首屏即应用已存配置

    var box = $('carTypeCfg');
    if (!box) return;

    /* 关键 —— 下面会再定义一个同名 render() 用于配置列表自身，
     * JS 作用域会就近解析到内层那个，导致「车型高亮变更后重渲染主表」
     * 静默失效。这里先在外层把主表 render 捕获成别名，persist / 重置
     * 按钮里用这个别名调主表渲染。
     * （拆分后 render 是 init 注入的参数，捕获逻辑与拆分前一致。） */
    var renderGrid = render;

    function colorOptions(sel) {
      return (Utils.CAR_COLORS || []).map(function (c) {
        return '<option value="' + escapeHtml(c.value) + '"' +
          (c.value.toLowerCase() === String(sel).toLowerCase() ? ' selected' : '') + '>' +
          escapeHtml(c.name) + '</option>';
      }).join('');
    }

    function persist(cfg) {
      if (Store && Store.set) Store.set(Store.KEYS.carTypeStyle, cfg);
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
      Store.set(Store.KEYS.carTypeCfgCollapsed, !!collapsed);
    }
    applyCarTypeCollapse(Store.get(Store.KEYS.carTypeCfgCollapsed, false) === true);
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
          if (Store && Store.remove) Store.remove(Store.KEYS.carTypeStyle);
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

  /* ==================== 装配 ==================== */

  /**
   * @param {Object} deps 由 app.js 注入
   * @param {Object}   deps.state            主状态（recompute / 货物重量 / 车型高亮读写）
   * @param {Function} deps.render           重渲主表（recompute / 车型高亮 persist）
   * @param {Function} deps.updateDetailTitle 刷新明细标题计重（货物重量变更后）
   * @param {Function} deps.openDetail       重开明细（车型高亮实时刷新）
   * @param {Array}    deps.defaultEstGoods  货物推算重量默认列表（app.js 私有常量，注入共享）
   */
  function init(deps) {
    if (inited) return;
    inited = true;
    state = deps.state;
    render = deps.render;
    updateDetailTitle = deps.updateDetailTitle;
    openDetail = deps.openDetail;
    DEFAULT_EST_GOODS = deps.defaultEstGoods;

    // 设置浮窗：打开即刷新「默认文件夹」路径显示
    UI.Modal.register('modalSettings', { onOpen: refreshFolderPath });

    // 设置：表格字号滑块（主页表格 grid + 明细抽屉表格 detailTable 同步生效）
    var gridFontSize = $(Store.KEYS.gridFontSize);
    var gridFontSizeVal = $('gridFontSizeVal');
    if (gridFontSize && gridFontSizeVal) {
      var savedFs = Store.get(Store.KEYS.gridFontSize, '');
      if (savedFs) {
        gridFontSize.value = savedFs;
        gridFontSizeVal.textContent = savedFs + 'px';
        applyTableFontSize(savedFs);
      }
      // 刻度节点是静态的，缓存一次即可——滑块 input 高频触发，避免每次都 querySelectorAll
      var tickEls = Array.prototype.slice.call(document.querySelectorAll('.range-ticks .tick'));
      function highlightTick(v) {
        for (var i = 0; i < tickEls.length; i++) {
          tickEls[i].classList.toggle('active', tickEls[i].textContent === v);
        }
      }
      // 拖动滑块会高频触发 input，写存储走防抖，避免每次都落盘 localStorage
      var saveFontSize = Utils.debounce(function (v) {
        Store.set(Store.KEYS.gridFontSize, v);
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

    // 设置：默认文件夹（路径显示 + 「修改」入口 → 复用数据源模块的选文件夹流程）
    var btnSettingFolder = $('btnSettingFolder');
    refreshFolderPath();
    if (btnSettingFolder) {
      btnSettingFolder.addEventListener('click', function () {
        DataSource.pickFolder();
      });
    }

    // 卸车地点（aggregate 段2 读取）
    initSpotConfig('unloadSpots', 'btnAddUnloadSpot', 'unloadSpots', DEFAULT_UNLOAD_SPOTS);
    // 黑罐识别（aggregate 段1 G7 罐按收货人识别）
    initSpotConfig('blackTankSpots', 'btnAddBlackTankSpot', 'blackTankSpots', DEFAULT_BLACK_TANK_SPOTS);

    // 车型高亮配置（对齐 VBA 显示信息.bas）
    initCarTypeConfig();
  }

  global.Settings = { init: init };

})(window);
