/**
 * utils.js —— 全局工具函数
 * 暴露全局：window.Utils
 * 加载顺序：必须在 aggregate.js、report31814.js、app.js 之前
 */
(function (global) {
  'use strict';

  var Utils = {};

  /* ==================== DOM ==================== */
  Utils.$ = function (id) { return document.getElementById(id); };

  /* ==================== 安全绑定 ==================== */
  Utils.on = function (id, evt, fn) {
    var el = Utils.$(id);
    if (el) el.addEventListener(evt, fn);
    return el;
  };

  /* ==================== HTML 转义 ==================== */
  Utils.escapeHtml = function (s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  };

  /* ==================== Toast 消息 ==================== */
  Utils.toast = function (msg, type) {
    var t = Utils.$('toast');
    if (!t) return;
    t.textContent = msg;
    t.className = 'toast show' + (type ? ' ' + type : '');
    clearTimeout(t._timer);
    t._timer = setTimeout(function () { t.className = 'toast'; }, 2600);
  };

  /* ==================== 日期格式化 ==================== */
  Utils.formatDate = function (d) {
    return (d.getMonth() + 1) + '月' + d.getDate() + '日';
  };

  /** 补零：1 → "01" */
  function pad2(n) { return (n < 10 ? '0' : '') + n; }

  /** 日期时间 → "2026-09-02 08:30" */
  Utils.formatDateTime = function (d) {
    var t = (d instanceof Date) ? d : new Date(d);
    if (isNaN(t.getTime())) return '';
    return t.getFullYear() + '-' + pad2(t.getMonth() + 1) + '-' + pad2(t.getDate()) +
           ' ' + pad2(t.getHours()) + ':' + pad2(t.getMinutes());
  };

  /* ==================== 字符串首字符 ==================== */
  Utils.firstChar = function (s) { return String(s == null ? '' : s).charAt(0); };

  /* ==================== VBA 函数模拟 ==================== */

  /** VBA InStr：找不到返回 0，找到返回 1 基位置；查找空串返回 1 */
  Utils.vbInStr = function (haystack, needle) {
    haystack = haystack == null ? '' : String(haystack);
    needle = needle == null ? '' : String(needle);
    if (needle === '') return 1;
    var i = haystack.indexOf(needle);
    return i === -1 ? 0 : i + 1;
  };

  /** VBA Val：从字符串开头解析数字，失败返回 0 */
  Utils.vbVal = function (v) {
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number') return v;
    var s = String(v).trim().replace(/^[\s　]+/, '');
    var m = /^[-+]?\d*\.?\d+(?:[eE][-+]?\d+)?/.exec(s);
    if (!m) return 0;
    var n = parseFloat(m[0]);
    return isNaN(n) ? 0 : n;
  };

  /** VBA Left(str, n) */
  Utils.vbLeft = function (s, n) {
    return String(s == null ? '' : s).substring(0, n);
  };

  /** VBA Mid(str, start, len)：start 为 1 基 */
  Utils.vbMid = function (s, start, len) {
    s = String(s == null ? '' : s);
    var a = Math.max(0, start - 1);
    return len === undefined ? s.substring(a) : s.substring(a, a + len);
  };

  /* ==================== 日期/时间工具 ==================== */

  /** 解析到达时间字符串 → Date，失败返回 null */
  Utils.parseArriveTime = function (v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    if (!s) return null;
    var m = /(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})[日T\s]*(\d{1,2})?:?(\d{1,2})?:?(\d{1,2})?/.exec(s);
    if (!m) return null;
    var d = new Date(
      +m[1], +m[2] - 1, +m[3],
      m[4] ? +m[4] : 0, m[5] ? +m[5] : 0, m[6] ? +m[6] : 0
    );
    return isNaN(d.getTime()) ? null : d;
  };

  /** 小时差：等价于 VBA DateDiff("h", from, to) */
  Utils.hoursDiff = function (from, to) {
    if (!from) return 0;
    return (to.getTime() - from.getTime()) / 3600000;
  };

  /* ==================== 车种提取 ==================== */

  /**
   * 从车种字符串提取车型字母（VBA 正则 [^\d]+ 逻辑）
   */
  Utils.extractCarType = function (carType) {
    var s = String(carType == null ? '' : carType);
    var m = /[^\d]+/.exec(s);
    var t = m ? m[0] : '';
    var order = ['C', 'X', 'P', 'G', 'YW', 'T', 'B', 'D', 'K'];
    var result = t;
    for (var i = 0; i < order.length; i++) {
      var ls = order[i];
      if (t.indexOf(ls) >= 0) {
        result = (Utils.vbLeft(t, 1) === 'N' && Utils.vbMid(t, 2, 1) === 'X') ? 'X' : Utils.vbLeft(t, 1);
        break;
      }
      result = t;
    }
    return result;
  };

  /**
   * 31814 用：根据车种+车号判定车型（与 extractCarType 不同，需考虑车号）
   */
  Utils.determineCarType = function (carTypeRaw, carNo) {
    var ct = String(carTypeRaw || '').toUpperCase();
    var cn = String(carNo || '');
    var firstCN = cn.charAt(0);
    if (Utils.firstChar(ct) === 'N' && firstCN === '5') return 'X';
    if (Utils.firstChar(ct) === 'B') {
      if (ct === 'BH1') return 'P';
      if (firstCN === '5') return 'X';
      if (firstCN === '6') return 'G';
      return Utils.firstChar(ct);
    }
    return Utils.firstChar(ct);
  };

  /* ==================== 股道排序 ==================== */
  Utils.compareTrackId = function (a, b) {
    a = String(a); b = String(b);
    var na = /^\d+$/.test(a), nb = /^\d+$/.test(b);
    if (na && nb) return +a - +b;
    if (na) return -1;
    if (nb) return 1;
    var ma = /^(\D*)(\d+)$/.exec(a), mb = /^(\D*)(\d+)$/.exec(b);
    var pa = ma ? ma[1] : a, pb = mb ? mb[1] : b;
    if (pa !== pb) return pa < pb ? -1 : 1;
    return (ma ? +ma[2] : 0) - (mb ? +mb[2] : 0);
  };

  /* ==================== 车种/车号颜色规则 ==================== */

  /**
   * 车种/车号颜色规则（对齐 VBA 显示信息.bas）
   *
   * 列索引从 Aggregate.COL 取，不再硬编码 row[2]/row[3]：
   * SMIS 导出列序一旦变动，硬编码会静默取错列（颜色错乱但不报错，极难发现）。
   *
   * utils.js 早于 aggregate.js 加载，模块顶层取不到 COL，故在调用时惰性读取
   * （调用点均在 Aggregate 就绪之后）；取不到时退回原魔数，行为不变。
   */
  Utils.carStyle = function (row) {
    var C = (global.Aggregate && global.Aggregate.COL) || null;
    var iType = C ? C.CARTYPE : 2;
    var iNo = C ? C.CARNO : 3;
    var t = String(row[iType] || '');
    var n = String(row[iNo] || '');
    var cls = '', bg = '', clsN = '', bgN = '';

    if (t.indexOf('DK') === 0) { cls = 'car-red'; }
    else if (t.indexOf('YW') === 0 || t.indexOf('YZ') === 0) { cls = 'car-red'; }
    else if (t.indexOf('B') === 0) { cls = 'car-purple'; }
    else if (t.indexOf('K') === 0) { cls = 'car-blue'; }
    else if (t.indexOf('T') === 0) { cls = 'car-pink'; }

    if (t.indexOf('P') === 0) { bg = 'car-yellow-bg'; }
    else if (/X/.test(t) && cls === '') { bg = 'car-grey-bg'; }

    if (n.charAt(0) === '0') {
      if (n.charAt(1) === '7') { bgN = 'car-self-bold'; clsN = 'car-self-bold'; }
      else { bgN = 'car-grey-bg'; clsN = 'car-grey-bg'; }
    } else {
      clsN = cls;
    }

    return { cls: cls, bg: bg, clsN: clsN, bgN: bgN };
  };

  /* ==================== Toast 悬停暂停 ==================== */
  (function () {
    var t = Utils.$('toast');
    if (t) {
      t.addEventListener('mouseenter', function () { clearTimeout(t._timer); });
      t.addEventListener('mouseleave', function () {
        t._timer = setTimeout(function () { t.className = 'toast'; }, 800);
      });
    }
  })();

  /* ==================== 防抖 / 节流 ==================== */
  /** 防抖：停止触发 wait 毫秒后才执行（用于滑块拖动、输入搜索等高频事件）。
   *  最后一次调用覆盖前一次，适合「结束后才落地」的场景（如写存储）。 */
  Utils.debounce = function (fn, wait) {
    var timer = null;
    return function () {
      var ctx = this, args = arguments;
      if (timer) clearTimeout(timer);
      timer = setTimeout(function () { fn.apply(ctx, args); }, wait || 150);
    };
  };

  /** 节流：每 wait 毫秒最多执行一次（leading 模式，开头立即执行）。
   *  适合「持续触发但需限频」的场景（如 resize、scroll 渲染）。 */
  Utils.throttle = function (fn, wait) {
    var last = 0, timer = null;
    return function () {
      var ctx = this, args = arguments;
      var now = Date.now();
      var remain = wait - (now - last);
      if (remain <= 0) {
        if (timer) { clearTimeout(timer); timer = null; }
        last = now;
        fn.apply(ctx, args);
      } else if (!timer) {
        timer = setTimeout(function () {
          last = Date.now(); timer = null;
          fn.apply(ctx, args);
        }, remain);
      }
    };
  };

  global.Utils = Utils;
})(window);