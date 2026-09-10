/* ============================================================================
 * 发车流程测试页 shim（供 departure_flow.html 引用）
 *
 * 模拟帆软报表「运行时」的最小环境，仅供本地测试
 *   「插件定位 iframe → 找到 contentPane → 写入 B4~F4」这条链路：
 *
 *   - window.contentPane  setCellValue / getCellValue（插件靠它写入单元格）
 *   - window.$            jQuery 子集（供下面的单元格编辑脚本使用）
 *   - 单元格编辑           点 B4 出股道下拉，C4~F4 出输入框
 *   - 提交按钮             alert 提示，不发起任何服务器请求
 *
 * 注意：
 *   1. 这是测试替身，不是真实帆软。真实报表的 contentPane 由 finereport.js
 *      的 FR.createWidget(...) 创建（见 fclcweb.html 原文件）。
 *   2. 本文件【不做任何持久化】：不写 localStorage、不读 URL 参数，
 *      每次打开 departure_flow.html 都是干净的空表，只有插件写入才会填值。
 *   3. 单元格定位按 id 前缀匹配（B4 → B4-0-55），不依赖固定的报表实例号。
 * ==========================================================================*/
'use strict';

/* ---------- $ shim（jQuery 子集，够编辑脚本用即可） ---------- */
(function () {
  function wrap(arr) {
    arr.first = function () { return wrap(arr.slice(0, 1)); };
    arr.last = function () { return wrap(arr.slice(-1)); };
    arr.eq = function (i) { return wrap([arr[i]]); };
    arr.get = function (i) { return arr[i]; };
    arr.each = function (fn) { arr.forEach(function (el, i) { fn.call(el, i, el); }); return arr; };
    arr.click = function () { arr.forEach(function (el) { el.click(); }); return arr; };
    arr.attr = function (n, v) {
      if (v === undefined) return arr[0] && arr[0].getAttribute(n);
      arr.forEach(function (el) { el.setAttribute(n, v); }); return arr;
    };
    arr.text = function (v) {
      if (v === undefined) return arr[0] && arr[0].textContent;
      arr.forEach(function (el) { el.textContent = v; }); return arr;
    };
    arr.val = function (v) {
      if (v === undefined) { var e = arr[0]; return (e && e.value !== undefined) ? e.value : ''; }
      arr.forEach(function (el) { el.value = v; }); return arr;
    };
    arr.find = function (s) {
      var r = [];
      arr.forEach(function (el) { r = r.concat(Array.prototype.slice.call(el.querySelectorAll(s))); });
      return wrap(r);
    };
    return arr;
  }
  window.$ = function (sel, ctx) {
    if (typeof sel === 'string') {
      return wrap(Array.prototype.slice.call((ctx || document).querySelectorAll(sel)));
    }
    if (sel && sel.nodeType) return wrap([sel]);
    if (sel && sel.length !== undefined) return wrap(Array.prototype.slice.call(sel));
    return wrap([]);
  };
})();

/* ---------- contentPane shim（对齐 FineReport 的 setCellValue / getCellValue）---------- */
window.contentPane = (function () {
  /* 单元格定位：id 精确 → id 前缀（B4 → B4-0-55）→ editor 属性 */
  function findCell(name) {
    if (!name) return null;
    var el = document.getElementById(name);
    if (el) return el;
    try {
      el = document.querySelector('[id^="' + name + '-"]');
      if (el) return el;
      el = document.querySelector('[editor="' + name + '"]');
    } catch (e) { /* ignore */ }
    return el || null;
  }
  return {
    findCell: findCell,
    /* 插件调用形式：setCellValue('B4', null, '3道') */
    setCellValue: function (col, row, value) {
      if (arguments.length >= 3) value = arguments[2];
      var el = findCell(col);
      if (el) el.textContent = (value == null ? '' : value);
      return el;
    },
    getCellValue: function (col, row) {
      var el = findCell(col);
      return el ? el.textContent : null;
    }
  };
})();

/* ---------- 单元格编辑：B4 为股道下拉，其余 editor 单元格为文本输入 ---------- */
(function () {
  var DROPDOWNS = {
    'B4': ['1道', '2道', '3道', '4道', '5道', '6道', '7道', '8道', '9道', '10道']
  };
  var cur = null;

  var cells = document.querySelectorAll('td[editor]');
  for (var i = 0; i < cells.length; i++) { cells[i].style.cursor = 'pointer'; }

  function finish(commit) {
    if (!cur) return;
    var ed = cur.el, td = cur.td, orig = cur.orig, name = cur.name;
    cur = null;
    var v = commit ? String(ed.value == null ? '' : ed.value) : '';
    if (ed.parentNode === td) td.removeChild(ed);
    td.textContent = (commit && v !== '') ? v : orig;
    /* 同步一份到 shim，保证 getCellValue 能读到 */
    if (commit && v !== '' && window.contentPane) {
      try { window.contentPane.setCellValue(name, null, v); } catch (e) { /* ignore */ }
    }
  }

  function open(td) {
    var name = td.getAttribute('editor');
    var opts = DROPDOWNS[name];
    var orig = (td.textContent || '').trim();
    var el;

    if (opts) {
      el = document.createElement('select');
      var o0 = document.createElement('option'); o0.value = ''; o0.textContent = '';
      el.appendChild(o0);
      for (var k = 0; k < opts.length; k++) {
        var o = document.createElement('option');
        o.value = opts[k]; o.textContent = opts[k];
        el.appendChild(o);
      }
    } else {
      el = document.createElement('input');
      el.type = 'text';
    }

    el.style.width = (td.getAttribute('widgetwidth') || 52) + 'px';
    el.style.height = (td.getAttribute('widgetheight') || 29) + 'px';
    el.style.boxSizing = 'border-box';
    el.style.font = 'inherit';
    el.style.border = '0';
    el.style.padding = '0 2px';
    el.style.background = 'transparent';
    el.value = orig;

    td.textContent = '';
    td.appendChild(el);
    cur = { el: el, td: td, orig: orig, name: name };
    el.focus();
    if (el.select) { try { el.select(); } catch (e) { /* ignore */ } }

    el.addEventListener('change', function () { finish(true); });
    el.addEventListener('blur', function () { finish(true); });
    el.addEventListener('keydown', function (ev) {
      if (ev.keyCode === 13) finish(true);          /* Enter 确认 */
      else if (ev.keyCode === 27) finish(false);    /* Esc 取消 */
    });
  }

  document.addEventListener('click', function (e) {
    var td = (e.target && e.target.closest) ? e.target.closest('td[editor]') : null;
    if (!td) return;
    if (cur && cur.td === td) return;   /* 已在本格编辑 */
    finish(true);                        /* 先提交上一格 */
    open(td);
    e.stopPropagation();
  }, true);
})();

/* ---------- 提交按钮：仅提示，绝不发起服务器请求 ---------- */
document.addEventListener('click', function (e) {
  var t = e.target;
  var btn = (t && t.closest) ? t.closest('.fr-btn[widgetname="Submit"]') : null;
  if (!btn && t && t.classList && t.classList.contains('x-emb-submit')) { btn = t; }
  if (btn) {
    if (e.stopImmediatePropagation) { e.stopImmediatePropagation(); }
    if (e.preventDefault) { e.preventDefault(); }
    if (e.stopPropagation) { e.stopPropagation(); }
    alert('提交成功（本地测试，不会真的提交到服务器）');
  }
}, true);
