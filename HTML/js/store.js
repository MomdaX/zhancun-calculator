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
    try {
      var s = localStorage.getItem(keyOf(key));
      if (s == null) return def;
      try {
        return JSON.parse(s);
      } catch (e) {
        // 兼容升级前直接存原始字符串的旧数据（如 '文件夹名' 无引号）：
        // JSON 解析失败时返回原始字符串，避免老用户升级后显示丢失。
        return s;
      }
    } catch (e) {
      return def;
    }
  }

  /** 写入设置项（内部 JSON 序列化，读取时自动还原） */
  function set(key, val) {
    try {
      localStorage.setItem(keyOf(key), JSON.stringify(val));
      return true;
    } catch (e) {
      return false;   // 隐私模式 / 配额满：静默失败，不影响主流程
    }
  }

  function remove(key) {
    try { localStorage.removeItem(keyOf(key)); } catch (e) {}
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

  global.Store = {
    get: get,
    set: set,
    remove: remove,
    async: {
      get: idbGet,
      set: idbSet,
      remove: idbRemove
    },
    PREFIX: PREFIX
  };
})(window);
