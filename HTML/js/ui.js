/**
 * ui.js —— 通用界面组件
 * ============================================================================
 * 暴露全局：window.UI
 * 加载顺序：必须在 app.js 之前（依赖 Utils）
 *
 * 收录的组件都是「页面里出现两次以上」的重复模式：
 *   UI.Modal    浮窗：打开/关闭/点遮罩关闭/ESC 关栈顶
 *   UI.Drawer   抽屉：同上，与 Modal 共用同一套开启栈
 *   UI.Dropdown 下拉菜单：展开/收起/点外部关闭/选中回调
 *
 * 为什么需要统一的开启栈：
 *   原来每个浮窗各自写一遍「关闭 + 点遮罩关闭」，ESC 处理则要在一处列举所有浮窗 id
 *   （app.js 里那段 if (modal31814) ... if (modalSettings) ...），
 *   新增浮窗时漏改就会「ESC 关不掉」。改为栈后，注册即自动获得全部行为。
 * ============================================================================
 */
(function (global) {
  'use strict';

  var $ = Utils.$;
  var UI = {};

  /* ==========================================================================
   * 浮窗 / 抽屉
   * ========================================================================== */

  /** 已注册面板：id → { el, maskId, onOpen, onClose, kind } */
  var panels = {};
  /** 开启栈（后进先出），ESC 只关栈顶 */
  var stack = [];

  /**
   * 注册一个面板。注册后自动获得：
   *   点遮罩关闭、ESC 关闭（仅关栈顶）、UI.Modal.open(id)/close(id)
   * @param {string} id         面板元素 id（不含 #）
   * @param {Object} [opts]
   * @param {string}   [opts.maskId]  配套遮罩元素 id（抽屉用，遮罩与面板分离）
   * @param {Function} [opts.onOpen]  打开后回调
   * @param {Function} [opts.onClose] 关闭后回调
   * @param {boolean}  [opts.esc]     是否响应 ESC（默认 true）
   */
  function register(id, opts) {
    opts = opts || {};
    var el = $(id);
    if (!el) return;

    panels[id] = {
      el: el,
      maskId: opts.maskId || '',
      onOpen: opts.onOpen || null,
      onClose: opts.onClose || null,
      esc: opts.esc !== false
    };

    // 点空白处关闭：抽屉有独立遮罩元素，浮窗则是自身外层（面板与遮罩同体）
    var hitEl = (opts.maskId && $(opts.maskId)) ||
                (el.classList.contains('modal') ? el : null);
    if (hitEl) {
      hitEl.addEventListener('click', function (e) {
        if (e.target === hitEl) close(id);
      });
    }
  }

  function open(id) {
    var p = panels[id];
    if (!p) return;
    p.el.classList.add('show');
    var mask = p.maskId ? $(p.maskId) : null;
    if (mask) mask.classList.add('show');
    // 已在栈中则先移除，保证它在栈顶（重复 open 不会破坏栈顺序）
    var i = stack.indexOf(id);
    if (i >= 0) stack.splice(i, 1);
    stack.push(id);
    if (p.onOpen) p.onOpen();
  }

  function close(id) {
    var p = panels[id];
    if (!p) return;
    p.el.classList.remove('show');
    var mask = p.maskId ? $(p.maskId) : null;
    if (mask) mask.classList.remove('show');
    var i = stack.indexOf(id);
    if (i >= 0) stack.splice(i, 1);
    if (p.onClose) p.onClose();
  }

  function closeTop() {
    if (!stack.length) return false;
    close(stack[stack.length - 1]);
    return true;
  }

  function isOpen(id) {
    var p = panels[id];
    return !!(p && p.el.classList.contains('show'));
  }

  /**
   * 是否有任意面板处于打开态（无参版）。
   * isOpen 必须传 id，外部模块（如地图）拿不到具体 id 时用它判断「浮窗是否优先」。
   */
  function anyOpen() { return stack.length > 0; }

  /* ---------------- ESC：只关栈顶，全部关完后再关抽屉 ---------------- */
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    closeTop();
  });

  UI.Modal = {
    register: register,
    open: open,
    close: close,
    closeTop: closeTop,
    isOpen: isOpen,
    anyOpen: anyOpen
  };

  /**
   * 抽屉：与 Modal 完全相同的机制，仅多一个遮罩元素。
   * 单独取名只为调用处语义清晰（openDrawer / closeDrawer）。
   */
  UI.Drawer = {
    register: register,
    open: open,
    close: close,
    isOpen: isOpen,
    anyOpen: anyOpen
  };

  /** 面板无关的顶层出口：外部模块判断「当前是否有浮窗/抽屉打开」 */
  UI.anyOpen = anyOpen;

  /* ==========================================================================
   * 下拉菜单
   * ========================================================================== */

  var openDropdown = null;

  /**
   * 绑定一个下拉菜单
   * @param {string|HTMLElement} btnId  触发按钮
   * @param {string|HTMLElement} listId 菜单容器
   * @param {Object} [opts]
   * @param {Function} [opts.onSelect] 选中回调 (itemEl, action)
   * @param {string}   [opts.itemSel]  菜单项选择器，默认 '.menu-item'
   * @param {string}   [opts.attr]     取值的属性名，默认 'data-action'
   */
  UI.Dropdown = function (btnId, listId, opts) {
    opts = opts || {};
    var btn = typeof btnId === 'string' ? $(btnId) : btnId;
    var list = typeof listId === 'string' ? $(listId) : listId;
    if (!btn || !list) return { show: function () {}, hide: function () {}, toggle: function () {} };

    var itemSel = opts.itemSel || '.menu-item';
    var attr = opts.attr || 'data-action';

    function hide() {
      list.style.display = 'none';
      if (openDropdown === api) openDropdown = null;
    }
    function show() {
      if (openDropdown && openDropdown !== api) openDropdown.hide();
      list.style.display = 'block';
      openDropdown = api;
    }

    btn.addEventListener('click', function (e) {
      e.stopPropagation();
      if (list.style.display === 'none') show(); else hide();
    });

    list.addEventListener('click', function (e) {
      var item = e.target.closest ? e.target.closest(itemSel) : null;
      if (!item) return;
      hide();
      if (opts.onSelect) opts.onSelect(item, item.getAttribute(attr));
    });

    // 点页面任意其它位置收起
    document.addEventListener('click', function () {
      if (list.style.display !== 'none') hide();
    });

    var api = { show: show, hide: hide, toggle: function () { if (list.style.display === 'none') show(); else hide(); } };
    return api;
  };

  global.UI = UI;
})(window);
