/* ============================================================================
 * 发车流程测试页 shim（供 departure_flow.html 引用）
 *
 * 模拟帆软报表「运行时」的最小环境，仅供本地测试
 *   「插件定位 iframe → 找到 contentPane → 写入 B4~F4」这条链路：
 *
 *   - window.contentPane  setCellValue / getCellValue（插件靠它写入单元格）
 *   - window.$            jQuery 子集（供下面的单元格编辑脚本使用）
 *   - 单元格编辑           点 B4 出股道下拉，C4~F4 出输入框
 *   - 提交按钮             在 #r-4-0 上方插入一行提交记录（左半同步插一条对齐行），不弹窗、不发请求
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

/* ---------- 提交按钮：向主数据区追加一条记录，不弹窗、绝不发起服务器请求 ---------- */
document.addEventListener('click', function (e) {
  var t = e.target;
  var btn = (t && t.closest) ? t.closest('.fr-btn[widgetname="Submit"]') : null;
  if (!btn && t && t.classList && t.classList.contains('x-emb-submit')) { btn = t; }
  if (btn) {
    if (e.stopImmediatePropagation) { e.stopImmediatePropagation(); }
    if (e.preventDefault) { e.preventDefault(); }
    if (e.stopPropagation) { e.stopPropagation(); }
    appendSubmitRecord();                                   /* ★ 唯一反馈：加一行记录，不再 alert */
  }
}, true);

/**
 * 提交后插入一条记录（不弹窗、不发请求、不做持久化）。
 *
 * 插入位置：#r-4-0 的「上方」—— 即第一条数据行之前（不是追加到末尾）。
 *
 * 冻结表 tr:nth-child(2) 里是左右两个「可滚动数据区」，行必须成对插，否则会错行：
 *   · 左半 #frozen-west   —— 冻结列（col 1~5 = 股道 / 编组车次 / 辆数 / 换长 / 尾车车号）
 *   · 右半 #frozen-center —— 主数据区（col 6~25 = 编成列车 / 通知商检 / …）
 *   （表头和编辑行在 tr:nth-child(1) 的 #frozen-corner / #frozen-north 里，是固定的，不动）
 *
 * 两区里第一条数据行的 id 都是 r-4-0，各自插到它上方即可保持左右对齐。
 * 新行由现有行 clone 而来，保留列宽 / 边框 / 行高（30px），只清掉 id、editor、
 * 帆软超链接和编辑态高亮；左半那条写入本次 B4~F4 的值作为记录内容。
 */
function appendSubmitRecord() {
  var host = document.querySelector('#content-container > div > table > tbody > tr:nth-child(2)');
  if (!host) { return; }

  var westTbody = findTbody(host, '#frozen-west');      /* 左半：冻结列数据 */
  var centerTbody = findTbody(host, '#frozen-center');  /* 右半：主数据 */
  var westAnchor = pickAnchor(westTbody);
  var centerAnchor = pickAnchor(centerTbody);
  if (!westAnchor && !centerAnchor) { return; }

  var cols = ['B4', 'C4', 'D4', 'E4', 'F4'];        /* 股道 / 编组车次 / 辆数 / 换长 / 尾车车号 */
  var vals = [];
  for (var i = 0; i < cols.length; i++) { vals.push(readCell(cols[i])); }

  /* ① 左半：克隆首条数据行 → 写入本次提交的值 → 插到 #r-4-0 上方 */
  var leftTr = null;
  if (westTbody && westAnchor) {
    leftTr = westAnchor.cloneNode(true);
    cleanRow(leftTr);
    var tds = leftTr.querySelectorAll('td[col]');
    for (var j = 0; j < tds.length; j++) {
      var idx = parseInt(tds[j].getAttribute('col'), 10) - 1;   /* col 1~5 → B4~F4 */
      if (idx >= 0 && idx < vals.length) { writeCell(tds[j], vals[idx]); }
    }
    westTbody.insertBefore(leftTr, westAnchor);
  }

  /* ② 右半：克隆首条数据行 → 插到 #r-4-0 上方 */
  var rightTr = null;
  if (centerTbody && centerAnchor) {
    rightTr = centerAnchor.cloneNode(true);
    cleanRow(rightTr);
    centerTbody.insertBefore(rightTr, centerAnchor);
  }

  flash(rightTr, leftTr);       /* 闪一下黄底 = 提交成功的反馈（替代原来的弹窗） */
}

/* 插入锚点：优先 id="r-4-0"（第一条数据行），取不到就退回该区第一行 */
function pickAnchor(tbody) {
  if (!tbody) { return null; }
  return tbody.querySelector('tr#r-4-0') || tbody.rows[0] || null;
}

/* ---------- 提交记录用的小工具 ---------- */

/* 在 host 内按区域选择器找该区域的 x-table > tbody */
function findTbody(host, boxSel) {
  var box = host.querySelector(boxSel);
  if (!box) { return null; }
  return box.querySelector('table.x-table > tbody') || box.querySelector('tbody');
}

/* 读单元格文本：id 精确 → id 前缀（B4 → B4-0-55），与 contentPane.findCell 口径一致 */
function readCell(id) {
  var el = document.getElementById(id);
  if (!el) { el = document.querySelector('[id^="' + id + '-"]'); }
  return el ? (el.textContent || '').trim() : '';
}

/* 写单元格：优先写已有的内层 div，保持帆软 td > div 的结构 */
function writeCell(td, text) {
  var div = td.querySelector('div');
  if (div) { div.textContent = text; } else { td.textContent = text; }
  td.setAttribute('cv', '"' + text + '"');
}

/* clone 来的行要「去身份」：清掉 id / editor / 帆软超链接 / 编辑态高亮，避免与原行撞车 */
function cleanRow(tr) {
  var i, list;
  tr.removeAttribute('id');
  list = tr.querySelectorAll('[id]');
  for (i = 0; i < list.length; i++) { list[i].removeAttribute('id'); }
  list = tr.querySelectorAll('[editor]');
  for (i = 0; i < list.length; i++) {
    list[i].removeAttribute('editor');
    list[i].style.cursor = '';
  }
  /* 冻结列的数据格里带帆软超链接（onclick="_g().stopEditing(),eval(...)"），
     测试页没有 FR / _g，留着的话点新行会抛错，这里直接摘掉 */
  list = tr.querySelectorAll('span.linkspan');
  for (i = 0; i < list.length; i++) {
    list[i].removeAttribute('onclick');
    list[i].removeAttribute('link');
    list[i].style.cursor = '';
  }
  tr.classList.remove('cur-tr-bg');
  list = tr.querySelectorAll('td');
  for (i = 0; i < list.length; i++) {
    list[i].classList.remove('cur-tr-bg');
    /* 只清可编辑行的浅绿底 rgb(204,255,204)，保留主数据区红色的「否」等原有底色 */
    if (/rgb\(\s*204\s*,\s*255\s*,\s*204\s*\)/.test(list[i].style.backgroundColor || '')) {
      list[i].style.backgroundColor = '';
    }
  }
}

/* 新增行闪一下黄底；结束时还原各自原有底色，不破坏主数据区原有的红色等 */
function flash() {
  var rows = arguments, items = [], i, j, k;
  for (i = 0; i < rows.length; i++) {
    if (!rows[i]) { continue; }
    var list = rows[i].querySelectorAll('td');
    for (j = 0; j < list.length; j++) {
      items.push({ el: list[j], bg: list[j].style.backgroundColor || '' });
    }
  }
  for (k = 0; k < items.length; k++) {
    items[k].el.style.transition = 'background-color .45s ease';
    items[k].el.style.backgroundColor = '#ffd666';
  }
  window.setTimeout(function () {
    for (var m = 0; m < items.length; m++) {
      items[m].el.style.backgroundColor = items[m].bg;    /* 还原原有底色 */
    }
  }, 450);
}
