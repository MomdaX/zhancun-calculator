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
  // 10 种默认高亮颜色（下拉选择用）
  Utils.CAR_COLORS = [
    { name: '红', value: '#e53e3e' },
    { name: '橙', value: '#dd6b20' },
    { name: '黄', value: '#ecc94b' },
    { name: '绿', value: '#38a169' },
    { name: '青', value: '#319795' },
    { name: '蓝', value: '#2b6cb0' },
    { name: '靛', value: '#5a67d8' },
    { name: '紫', value: '#805ad5' },
    { name: '粉', value: '#d53f8c' },
    { name: '灰', value: '#a0aec0' }
  ];

  /* ==========================================================================
   * 车型高亮 —— 「设置 → 车型高亮」的可编辑配置，是车型着色的**唯一来源**
   * --------------------------------------------------------------------------
   * 覆盖范围（改一处，以下位置同步生效）：
   *   · 主表 到站列   汇总串中的车型片段（"P5" / "DK2" / "YW3"）
   *   · 明细 车种列   原始车种（"P64" / "DK"）
   *   · 明细 车号列   车号非 0 开头时继承车种的类
   *   · 明细 到站/发站列 经 renderDest 走同一判定
   *
   * 不在本配置范围内（各自独立，勿并入）：
   *   · 主表 车种列 的红/粉底 = 作业区禁入告警（股道×车种的提示，仅主表显示）
   *   · 车号规则（自备罐灰底 / 中粮罐加粗）按车号判定，且主表本就没有车号列
   *
   * 【历史】VBA 显示信息.bas 里这两处各写了一套硬编码（到站方向 Sub 的车型
   * If-ElseIf 链 / 标记到站方向颜色 Sub 的 Array("YW","D","P")），互不一致，
   * 改一处另一处不跟随。此处统一由本配置驱动，不要再在渲染处写死颜色。
   * --------------------------------------------------------------------------
   * match: 'starts'(前缀开头) / 'contains'(包含)；bold 决定颜色作用于文字(加粗)还是背景
   * ========================================================================== */
  Utils.defaultCarTypeConfig = [
    { prefix: 'DK', note: '大D车',  match: 'starts',   on: true, bold: true,  color: '#e53e3e' },
    { prefix: 'YW', note: '客车',   match: 'starts',   on: true, bold: true,  color: '#e53e3e' },
    { prefix: 'YZ', note: '客车',   match: 'starts',   on: true, bold: true,  color: '#e53e3e' },
    { prefix: 'B',  note: '机保车', match: 'starts',   on: true, bold: true,  color: '#805ad5' },
    { prefix: 'K',  note: '老K车',  match: 'starts',   on: true, bold: true,  color: '#2b6cb0' },
    { prefix: 'T',  note: '检衡车', match: 'starts',   on: true, bold: true,  color: '#d53f8c' },
    /* P 之前是黄底 bold:false，但用户要求"其它车型是字体色 + 加粗"，
     * 改成红字 + 加粗（与 DK/YW/YZ 同色），统一为"字体色"组。 */
    { prefix: 'P',  note: '盖车',   match: 'starts',   on: true, bold: true,  color: '#e53e3e' },
    /* 平板车：VBA 显示信息.bas 用「正则 \w+ 匹配第一段 + InStr("X") > 0」判定，
     * 实际效果是 NX70AF 这种以 NX 开头的也被标灰。
     * 关键区分：底色**只**作用于明细表车种列；主表到站列、车号列都不挂底色。 */
    { prefix: 'NX', note: '平板车(NX)', match: 'starts', on: true, bold: false, color: '#a0aec0' },
    { prefix: 'X',  note: '平板车',     match: 'starts', on: true, bold: false, color: '#a0aec0' }
  ];

  Utils.getCarTypeConfig = function () {
    var s = (global.Store && global.Store.get) ? global.Store.get('carTypeStyle', null) : null;
    if (Array.isArray(s) && s.length) {
      return s.map(function (e) {
        return {
          prefix: e.prefix || '',
          note: e.note || '',
          match: e.match === 'contains' ? 'contains' : 'starts',
          on: e.on !== false,
          bold: !!e.bold,
          color: e.color || '#e53e3e'
        };
      });
    }
    // 深拷贝默认，避免污染
    return Utils.defaultCarTypeConfig.map(function (e) {
      return { prefix: e.prefix, note: e.note, match: e.match, on: e.on, bold: e.bold, color: e.color };
    });
  };

  // 按当前配置生成动态样式表（车型高亮）
  Utils.applyCarTypeStyles = function () {
    var cfg = Utils.getCarTypeConfig();
    var css = '';
    for (var i = 0; i < cfg.length; i++) {
      var e = cfg[i];
      if (!e.on || !e.prefix) continue;
      if (e.bold) css += '.ctc-' + i + '{color:' + e.color + ';font-weight:700;}';
      else css += '.ctc-' + i + '{background-color:' + e.color + ';}';
    }
    // 兜底：cfg 里 X/NX 两项被删光时，carTypeClass 会返回 __ctc_fallback__，
    // 这里同步生成对应灰底样式，确保视觉一致。
    css += '.__ctc_fallback__{background-color:#a0aec0;}';
    var el = document.getElementById('carTypeStyleSheet');
    if (!el) {
      el = document.createElement('style');
      el.id = 'carTypeStyleSheet';
      document.head.appendChild(el);
    }
    el.textContent = css;
  };

  /**
   * 按「设置 → 车型高亮」配置取类名：命中第 i 项 → 'ctc-' + i，未命中返回 ''。
   *
   * 抽成独立函数的目的：主表到站列与明细车种列共用**同一套判定**。
   * 主表到站串是「分类+车数」形式（"P5" / "DK2" / "YW3"），明细车种列是
   * 原始车种（"P64" / "DK"），两者都交给这里匹配，用户在设置里改了颜色/加粗，
   * 两个位置同步生效，不会出现「明细变了主表没变」。
   *
   * @param {string} type 车种字符串，可带数量后缀
   * @returns {string} 'ctc-N' 或 ''
   */
  /**
   * 车种匹配：返回 { idx, prefix, isFlatbed }，未命中返回 null。
   *
   * 相比 carTypeClass，多返回一个 isFlatbed 标记——用于区分"平板车"（X/NX，
   * 底色规则）与"其它车型"（DK/YW/YZ/B/K/T/P，字体色规则）。渲染端根据
   * 上下文决定应用哪一组样式：
   *   · 主表到站列：仅"其它车型"挂 ctc（颜色=字体色）；"平板车"不上底色
   *   · 明细车种列：所有车型都挂 ctc（bold=true → 字体色；bold=false → 底色）
   *   · 明细车号列：不挂 ctc（车号是 7 位数字，盖底色不直观）
   *
   * @returns {{idx: number, prefix: string, isFlatbed: boolean}|null}
   */
  Utils.carTypeMatch = function (type) {
    var t = String(type == null ? '' : type).trim();
    if (!t) return null;
    var cfg = Utils.getCarTypeConfig();
    for (var i = 0; i < cfg.length; i++) {
      var e = cfg[i];
      if (!e.on || !e.prefix) continue;
      var hit = e.match === 'contains' ? (t.indexOf(e.prefix) > -1) : (t.indexOf(e.prefix) === 0);
      if (hit) return { idx: i, prefix: e.prefix, isFlatbed: e.prefix === 'X' || e.prefix === 'NX' };
    }
    /* 硬兜底：含 X 的车种一律视为平板车（X 开头 或 NX 开头）。
     * 不走配置——哪怕用户把 X/NX 两项都关掉/删了/改成不匹配 NX 的 starts，
     * 这条都保底命中，对齐 VBA「正则匹配第一段含 X」的语义。 */
    if (t.charAt(0) === 'X' || t.charAt(1) === 'X') {
      for (var k = 0; k < cfg.length; k++) {
        if ((cfg[k].prefix === 'X' || cfg[k].prefix === 'NX') && cfg[k].on) {
          return { idx: k, prefix: cfg[k].prefix, isFlatbed: true };
        }
      }
      // cfg 里 X/NX 都被删/关掉 → 兜底类（applyCarTypeStyles 同步生成灰底）
      return { idx: -1, prefix: '', isFlatbed: true };
    }
    return null;
  };

  /** 取样式类名（'ctc-N' / '__ctc_fallback__' / ''），carTypeMatch 的便捷封装 */
  Utils.carTypeClass = function (type) {
    var m = Utils.carTypeMatch(type);
    if (!m) return '';
    return m.idx < 0 ? '__ctc_fallback__' : 'ctc-' + m.idx;
  };

  Utils.carStyle = function (row) {
    var C = (global.Aggregate && global.Aggregate.COL) || null;
    var iType = C ? C.CARTYPE : 2;
    var iNo = C ? C.CARNO : 3;
    var t = String(row[iType] || '');
    var n = String(row[iNo] || '');
    var cls = Utils.carTypeClass(t);
    var bg = cls; // 背景与文字用同一类（applyCarTypeStyles 决定颜色/bg）
    var clsN = '', bgN = '';

    // 车号列：按车号首位单独判定（自备罐灰底 / 中粮罐加粗）。
    // **不**再继承车种列的 ctc——车号是 7 位数字，盖底色既不直观也容易与行内
    // 斑马纹混，且平板车的灰底规则不应蔓延到车号列。
    if (n.charAt(0) === '0') {
      if (n.charAt(1) === '7') { bgN = 'car-self-bold'; clsN = 'car-self-bold'; }
      else { bgN = 'car-grey-bg'; clsN = 'car-grey-bg'; }
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

  /* ==================== 复制文本到剪贴板 ==================== */
  /** 降级方案：临时 textarea + execCommand（file:// 打开或浏览器拒绝剪贴板权限时使用） */
  function fallbackCopy(text) {
    return new Promise(function (resolve, reject) {
      var ta = document.createElement('textarea');
      ta.value = text;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.left = '-9999px';
      document.body.appendChild(ta);
      ta.select();
      ta.setSelectionRange(0, ta.value.length);
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (err) { ok = false; }
      document.body.removeChild(ta);
      if (ok) resolve();
      else reject(new Error('浏览器拒绝了复制操作'));
    });
  }

  /**
   * 复制文本到剪贴板
   * @param {string} text 待复制内容
   * @returns {Promise<void>} 成功 resolve；失败 reject(Error)
   */
  Utils.copyText = function (text) {
    text = text == null ? '' : String(text);
    if (!text) return Promise.reject(new Error('内容为空'));
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text).catch(function () { return fallbackCopy(text); });
    }
    return fallbackCopy(text);
  };

  global.Utils = Utils;
})(window);