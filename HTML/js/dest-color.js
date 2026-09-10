/**
 * dest-color.js —— 到站富文本着色与辅助（从 app.js 抽出，P2-7 第三步）
 *
 * 纯函数（输入文本/片段 → 输出 HTML 串或布尔），被主表渲染与明细抽屉共用。
 * 依赖：Utils（carTypeMatch / carTypeClass / escapeHtml）、Store（词表配置）、Aggregate.COL；
 *       renderDest 运行时读 global.state.dirIndex.map（由 app.js 在 init 注入）。
 * 通过 global 暴露，供 app.js 主表/明细按原名字引用。
 */
(function (global) {
  'use strict';

  // 重声明 app.js 顶部的别名，使下方函数体与原文件零差异（stationSpan 用）
  var escapeHtml = Utils.escapeHtml;

  /* =================== 到站富文本着色 =================== */
  /**
   * @param {string}  text 到站/发站文本
   * @param {boolean} clickable 是否标记可双击的车站（仅明细页为 true，
   *        主表是汇总串如「德保44」，双击无意义，故不加虚线下划线）
   */
  /* 词表默认值（本文件内单一来源：isUnloadSpot / isBlackTankSpot / 设置面板初始化共用）。
   * 与 aggregate.js 的 DEFAULT_UNLOAD_SPOTS / DEFAULT_BLACK_TANK_SPOTS 保持一致；
   * 实际生效值以 Store 配置为准（用户在设置里增删），未配置时回落到此处。 */
  var DEFAULT_UNLOAD_SPOTS = ['永鑫', '货场', '天盛', '港务局'];
  var DEFAULT_BLACK_TANK_SPOTS = ['中粮', '外运'];

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
    var list = (Store.getList && Store.getList('unloadSpots', DEFAULT_UNLOAD_SPOTS)) ||
               DEFAULT_UNLOAD_SPOTS;
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
    var list = (Store.getList && Store.getList('blackTankSpots', DEFAULT_BLACK_TANK_SPOTS)) ||
               DEFAULT_BLACK_TANK_SPOTS;
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

  function renderDest(text, clickable, extraCls) {
    if (!text) return '';
    var parts = String(text).trim().split(/\s+/).filter(Boolean);
    var map = state.dirIndex.map;
    // 基础 class（如 'derived'），先于具体方向/标记 class；
    // CSS 中方向色 .nankou 等定义在 .derived 之后，会覆盖默认灰、保留斜体
    var pre = extraCls ? (extraCls + ' ') : '';
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
      var linkSt = (clickable && isStationName(p)) ? stationOf(p) : '';
      return stationSpan(p, pre + cls, linkSt);
    }).join(' ');
  }

  /**
   * 拼一个车站名 span：class + 可选 station-link（双击开地图用 data-station）。
   * renderDest（到站，含着色/标记）与 renderStationLink（发站，纯链接）共用，
   * 避免两处各拼一遍 class / station-link / data-station。
   * @param {string} text        显示文本
   * @param {string} cls         样式 class（可为空）
   * @param {string} linkStation 非空则挂 station-link + data-station（传 stationOf 结果）
   */
  function stationSpan(text, cls, linkStation) {
    var link = linkStation ? ' station-link' : '';
    var attr = linkStation ? ' data-station="' + escapeHtml(linkStation) + '"' : '';
    return '<span class="' + ((cls || '') + link).trim() + '"' + attr + '>' +
           escapeHtml(text) + '</span>';
  }

  /**
   * 发站列渲染：恒用原始发站，不加任何着色/标记，仅挂 station-link 供双击开地图。
   * 发站只是起点站名，到站那套方向色（沙口蓝/南口橙/管内紫）以及
   * YW/D/P 红、到卸/黑罐黑加粗等标记规则都与它无关。
   */
  function renderStationLink(text) {
    var s = String(text == null ? '' : text).trim();
    if (!s) return '';
    var st = stationOf(s);
    if (!st) return escapeHtml(s);   // 非站名：纯文本，不包 span
    return stationSpan(s, '', st);
  }

  /**
   * 有效长(m) → 换长：11m 为 1 个换长单位。
   * 保留 1 位小数，第二位起直接舍去（不四舍五入）：78.36 → 78.3
   */
  function effLenToChang(len) {
    return (Math.floor(len / 11 * 10) / 10).toFixed(1);
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

  global.renderDest = renderDest;
  global.stationSpan = stationSpan;
  global.renderStationLink = renderStationLink;
  global.isUnloadSpot = isUnloadSpot;
  global.isBlackTankSpot = isBlackTankSpot;
  global.typeKeyOf = typeKeyOf;
  global.stationOf = stationOf;
  global.isStationName = isStationName;
  global.effLenToChang = effLenToChang;
  global.NOT_STATION = NOT_STATION;
  global.DEFAULT_UNLOAD_SPOTS = DEFAULT_UNLOAD_SPOTS;
  global.DEFAULT_BLACK_TANK_SPOTS = DEFAULT_BLACK_TANK_SPOTS;

})(window);
