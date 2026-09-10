/**
 * config-io.js —— 本地配置导出/导入（JSON 编解码，纯函数）
 * ============================================================================
 * 暴露：window.ConfigIO
 * 加载顺序：在 app.js 之前（app.js 的「配置备份」按钮依赖此模块）
 *
 * 职责边界：
 *   · 本文件只负责 JSON 文本 ↔ { key: value } 的纯转换，**不碰 localStorage**；
 *   · 实际读写浏览器持久化由 Store.allSync() / Store.applySync() 完成；
 *   · value 为 JS 值（对象/数组/标量），导出即 JSON.stringify，导入即 JSON.parse，
 *     不残留「localStorage 原始 JSON 字符串」的双层转义，文件人类可读；
 *   · 元信息（生成时间/说明）放在顶层 _meta 字段，解析时剔除，不写回 localStorage；
 *   · 不可序列化的 IndexedDB 句柄（xlsDir）不在同步范围，故 JSON 中不会出现。
 *   · 用 .json 而非 .ini：Windows/Edge/Chrome 会拦截 .ini 下载（归入高危扩展名），
 *     .json 任何系统都不拦截，且天然可被任意工具查看/编辑。
 * ============================================================================
 */
(function (global) {
  'use strict';

  /** 序列化为 JSON 文本（带缩进，人类可读）。
   * @param {Object} map  { fullKey: value }（键应带 zhancun. 前缀）
   * @param {Object} [meta] 可选元信息，写入顶层 _meta 字段（不影响数据键） */
  function toJson(map, meta) {
    var out = {};
    if (meta) out._meta = meta;
    if (map && typeof map === 'object') {
      Object.keys(map).forEach(function (k) { out[k] = map[k]; });
    }
    return JSON.stringify(out, null, 2);
  }

  /** 解析 JSON 文本 → { fullKey: value }（剔除顶层 _meta）。
   *  空串 / 非字符串 / 非法 JSON 均安全返回空对象。 */
  function parseJson(text) {
    var s = (text == null ? '' : String(text)).trim();
    if (!s) return {};
    var raw;
    try { raw = JSON.parse(s); } catch (e) { raw = {}; }
    var out = {};
    if (raw && typeof raw === 'object') {
      Object.keys(raw).forEach(function (k) {
        if (k === '_meta') return;        // 元信息不写回 localStorage
        out[k] = raw[k];
      });
    }
    return out;
  }

  global.ConfigIO = {
    toJson: toJson,
    parseJson: parseJson
  };
})(window);
