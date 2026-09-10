/**
 * store.js —— 持久化统一封装
 * ============================================================================
 * 暴露全局：window.Store
 * 加载顺序：必须在 app.js 之前
 *
 * 原来项目里有四套并存、各写各的持久化：
 *   localStorage.getItem/setItem('zhancun.folderName')       —— 文件夹名
 *   localStorage.getItem/setItem('zhancun.gridFontSize')     —— 表格字号
 *   localStorage.setItem(persistKey, JSON.stringify(widths)) —— 列宽（col-resize 内）
 *   indexedDB 手写 openDB/idbSet/idbGet                      —— 目录句柄
 * 每处都自己写 try/catch、自己拼前缀、自己决定要不要 JSON 序列化，
 * 新增一个「记住窗口大小」之类的设置项就得再抄一遍。
 *
 * 统一后：
 *   Store.get('gridFontSize', '13')   // 同步，自动加前缀 + JSON + 容错
 *   Store.set('gridFontSize', '15')
 *   Store.remove('gridFontSize')
 *   await Store.async.set('xlsDir', handle)   // IndexedDB，可存句柄这类不可序列化对象
 *   await Store.async.get('xlsDir')
 * ============================================================================
 */
(function (global) {
  'use strict';

  var PREFIX = 'zhancun.';
  var DB_NAME = 'YardStorageDB';
  var DB_STORE = 'handles';

  /** 内存缓存：与 localStorage 同步（set/remove 同时更新）。
   *  作用：渲染每行（carTypeStyle）/ 聚合每行（blackTankSpots、unloadSpots）都会高频调用
   *  Store.get，原实现每次都走 localStorage 读取 + JSON.parse。缓存后命中只需对象属性查找。
   *  单页应用，无多标签页同步需求，内存态与 localStorage 始终一致（所有读写均经 Store）。 */
  var _mem = {};

  /** 补全前缀；已带前缀的键原样返回（历史键如 zhancun.grid.cols 可直接沿用，不丢记忆） */
  function keyOf(key) {
    return (String(key).indexOf(PREFIX) === 0) ? key : PREFIX + key;
  }

  /* ==================== 同步：localStorage ==================== */

  /**
   * 读取设置项
   * @param {string} key 键名（可省略 zhancun. 前缀）
   * @param {*} [def]    不存在或解析失败时的默认值
   */
  function get(key, def) {
    var k = keyOf(key);
    if (Object.prototype.hasOwnProperty.call(_mem, k)) return _mem[k];
    try {
      var s = localStorage.getItem(k);
      if (s == null) {
        // 原始类型 / null 的默认值可安全缓存；对象 / 数组默认不缓存，避免共享引用被误改。
        if (def === null || typeof def !== 'object') _mem[k] = def;
        return def;
      }
      try {
        var v = JSON.parse(s);
        _mem[k] = v;
        return v;
      } catch (e) {
        // 兼容升级前直接存原始字符串的旧数据（如 '文件夹名' 无引号）：
        // JSON 解析失败时返回原始字符串并缓存，避免老用户升级后显示丢失。
        _mem[k] = s;
        return s;
      }
    } catch (e) {
      return def;
    }
  }

  /**
   * 读取「列表型」配置（如 卸车地点 / 黑罐识别 词表）。
   * 统一处理三种回落：无 Store 环境、未设置、值不是非空数组 → 返回 defaults。
   * 消除调用处各自写 `(Store.get && Store.get(k, null)) || DEFAULTS` 的重复。
   * @param {string} key      键名
   * @param {Array}  defaults 回落默认列表
   */
  function getList(key, defaults) {
    var v = get(key, null);
    return (Array.isArray(v) && v.length) ? v : (defaults || []);
  }

  /** 写入设置项（内部 JSON 序列化，读取时自动还原） */
  function set(key, val) {
    var k = keyOf(key);
    try {
      localStorage.setItem(k, JSON.stringify(val));
      _mem[k] = val;   // 与 localStorage 同步；setItem 抛错（配额/隐私）则不缓存
      return true;
    } catch (e) {
      return false;   // 隐私模式 / 配额满：静默失败，不影响主流程
    }
  }

  function remove(key) {
    var k = keyOf(key);
    try { localStorage.removeItem(k); } catch (e) {}
    delete _mem[k];   // 同步失效内存缓存
  }

  /** 导出：收集所有同步持久化键（zhancun. 前缀）的 JS 值。
   *  返回 { fullKey: value }（value 为 JSON.parse 后的对象/数组/标量，兼容旧版原始字符串）。
   *  IndexedDB 句柄（xlsDir）不可序列化，不在此列。 */
  function allSync() {
    var out = {};
    try {
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (!k || k.indexOf(PREFIX) !== 0) continue;
        var raw = localStorage.getItem(k);
        try { out[k] = JSON.parse(raw); }      // 现代数据：已 JSON 序列化
        catch (e) { out[k] = raw; }             // 兼容升级前直接存原始字符串的旧数据
      }
    } catch (e) {}
    return out;
  }

  /** 导入：把 { fullKey: value } 经 Store.set 写回（自动补前缀 + JSON 序列化）。
   *  与逐项 localStorage.setItem 等价，但复用 Store 的内存缓存同步逻辑；返回实际写入条数。 */
  function applySync(map) {
    var n = 0;
    if (!map || typeof map !== 'object') return n;
    for (var key in map) {
      if (!Object.prototype.hasOwnProperty.call(map, key)) continue;
      try {
        set(key, map[key]);   // keyOf + JSON.stringify + 更新 _mem
        n++;
      } catch (e) {}
    }
    return n;
  }

  /* ==================== 异步：IndexedDB（可存句柄等不可序列化的值） ==================== */

  var dbPromise = null;

  function openDB() {
    if (dbPromise) return dbPromise;
    dbPromise = new Promise(function (res, rej) {
      var q = indexedDB.open(DB_NAME, 1);
      q.onupgradeneeded = function () {
        if (!q.result.objectStoreNames.contains(DB_STORE)) q.result.createObjectStore(DB_STORE);
      };
      q.onsuccess = function () { res(q.result); };
      q.onerror = function () { rej(q.error); };
    }).catch(function (e) {
      dbPromise = null;      // 打开失败时不缓存 Promise，下次可重试
      throw e;
    });
    return dbPromise;
  }

  function idbSet(k, v) {
    return openDB().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).put(v, k);
        tx.oncomplete = function () { res(); };
        tx.onerror = function () { rej(tx.error); };
      });
    }).catch(function () { /* 存储失败不影响主流程 */ });
  }

  function idbGet(k) {
    return openDB().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction(DB_STORE, 'readonly');
        var rq = tx.objectStore(DB_STORE).get(k);
        rq.onsuccess = function () { res(rq.result); };
        rq.onerror = function () { rej(rq.error); };
      });
    }).catch(function () { return null; });
  }

  function idbRemove(k) {
    return openDB().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction(DB_STORE, 'readwrite');
        tx.objectStore(DB_STORE).delete(k);
        tx.oncomplete = function () { res(); };
        tx.onerror = function () { rej(tx.error); };
      });
    }).catch(function () {});
  }

  /** 持久化键集中管理：所有键名在此一处定义，调用处统一用 Store.KEYS.xxx，
   *  避免散写字符串导致拼写错位、静默失败（keyOf 仍统一加前缀）。 */
  var KEYS = {
    folderName: 'folderName',
    gridFontSize: 'gridFontSize',
    showEmptyGroups: 'showEmptyGroups',
    notesCollapsed: 'notesCollapsed',
    unloadSpots: 'unloadSpots',
    blackTankSpots: 'blackTankSpots',
    carTypeStyle: 'carTypeStyle',
    estLoadGoods: 'estLoadGoods',     // 货物推算重量：[{ name, weight }] 列表（载重缺失时记事命中 name → 用 weight）
    xlsDir: 'xlsDir',
    detailCols: 'zhancun.detail.cols.v2',   // 历史列宽记忆键（已带前缀，keyOf 原样返回）
    cfgDzChecked: 'cfgDzChecked',
    corrInputs: 'corrInputs',
    readyTrains: 'readyTrains',
    // 以下 4 个原为散写的字符串键（未集中登记），补入后调用处统一用 Store.KEYS.xxx
    depCheci: 'depCheci',                        // 待发股道车次映射；由外部页写入，读取须直连 localStorage（绕过 Store 内存缓存）
    depModalSize: 'depModalSize',                // 发车作业浮窗尺寸记忆
    carTypeCfgCollapsed: 'carTypeCfgCollapsed',  // 车型高亮配置表收起状态
    eyeProtect: 'eyeProtect'                     // 护眼背景色浓度
  };

  global.Store = {
    KEYS: KEYS,
    get: get,
    getList: getList,
    set: set,
    remove: remove,
    allSync: allSync,
    applySync: applySync,
    async: {
      get: idbGet,
      set: idbSet,
      remove: idbRemove
    },
    PREFIX: PREFIX
  };
})(window);
