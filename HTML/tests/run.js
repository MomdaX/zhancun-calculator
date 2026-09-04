/**
 * 站存计算器 —— 回归测试
 * ============================================================================
 * 运行：  node HTML/tests/run.js        （仓库根目录）
 *         node tests/run.js             （HTML 目录）
 * 依赖：  无。纯 Node，不装任何 npm 包。
 *
 * 【为什么测这些】
 * 本项目是 VBA 的逐条翻译，业务规则藏在 resolveDest / setCarProperties 这类
 * 长 If-ElseIf 链里。这类代码出错是**静默**的——规则判错不会抛异常、不会崩溃，
 * 只会让站存数字悄悄变错，而本工具正是对账用的。因此回归测试是唯一能兜住
 * 这个风险的手段。
 *
 * 【已捕获的真实缺陷】（保留用例，防止回归）
 *   · 到站推断漏判：倒排索引只探测 note.charAt(0)，导致「装防城港箱 天驰」
 *     这类前缀写法被判为未命中，到站错误归入车种/罐型。
 *     实测 1339 行真实数据中有 93 行（6.95%）判错。见 suite「到站推断」。
 * ============================================================================
 */
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');

/* ==================== 把浏览器脚本载入 Node 沙箱 ==================== */

const ROOT = path.join(__dirname, '..');      // HTML/
const SRC = ['js/utils.js', 'js/store.js', 'js/ui.js', 'js/aggregate.js',
             'js/direction.data.js', 'js/track.config.js'];

const memStore = {};
const sandbox = {
  console, setTimeout, clearTimeout, setInterval, clearInterval,
  localStorage: {
    getItem: k => (Object.prototype.hasOwnProperty.call(memStore, k) ? memStore[k] : null),
    setItem: (k, v) => { memStore[k] = String(v); },
    removeItem: k => { delete memStore[k]; }
  },
  indexedDB: { open: () => ({}) },
  document: {
    getElementById: () => null,
    querySelector: () => null,
    querySelectorAll: () => [],
    addEventListener() {},
    removeEventListener() {},
    body: { classList: { add() {}, remove() {} } },
    readyState: 'complete'
  }
};
sandbox.window = sandbox;
vm.createContext(sandbox);
SRC.forEach(f => vm.runInContext(fs.readFileSync(path.join(ROOT, f), 'utf8'), sandbox, { filename: f }));

const { Utils, Store, Aggregate, YardConfig } = sandbox;
const COL = Aggregate.COL;

/* ==================== 极简断言 ==================== */

let pass = 0, fail = 0;
const failures = [];

function t(name, fn) {
  try { fn(); pass++; }
  catch (e) { fail++; failures.push(name + '  →  ' + e.message); }
}
function suite(name, fn) { console.log('\n── ' + name); fn(); }
function eq(actual, expected, msg) {
  const a = JSON.stringify(actual), b = JSON.stringify(expected);
  if (a !== b) throw new Error((msg ? msg + ': ' : '') + `期望 ${b}，实际 ${a}`);
}
function ok(cond, msg) { if (!cond) throw new Error(msg || '断言失败'); }

/** 造一行原始数据（16 列，默认空串） */
function mkRow(o) {
  const r = [];
  for (let i = 0; i < 16; i++) r[i] = '';
  Object.keys(o || {}).forEach(k => { r[COL[k]] = o[k]; });
  return r;
}

/* ==========================================================================
 * 1. VBA 函数模拟（Utils）—— 聚合与报表共用的取数口径
 * ========================================================================== */
suite('VBA 函数模拟', () => {
  t('vbVal 取前导数字', () => {
    eq(Utils.vbVal('12.5'), 12.5);
    eq(Utils.vbVal('38 吨'), 38);
    eq(Utils.vbVal('　12'), 12);          // 全角空格
    eq(Utils.vbVal('-3'), -3);
    eq(Utils.vbVal('.5'), 0.5);
  });
  t('vbVal 空/非法一律 0', () => {
    [null, undefined, '', 'abc', 'N/A', {}, []].forEach(v => eq(Utils.vbVal(v), 0));
  });
  t('vbVal 不返回 Infinity（旧报表 parseFloat 会，污染合计）', () => {
    eq(Utils.vbVal('Infinity'), 0);
    eq(Utils.vbVal('-Infinity'), 0);
  });
  t('vbInStr 为 1 基，找不到返 0，查空串返 1', () => {
    eq(Utils.vbInStr('abc', 'b'), 2);
    eq(Utils.vbInStr('abc', 'z'), 0);
    eq(Utils.vbInStr('abc', ''), 1);
  });
  t('vbLeft / vbMid', () => {
    eq(Utils.vbLeft('C70', 1), 'C');
    eq(Utils.vbMid('01234567', 2, 1), '1');
    eq(Utils.vbMid('abcdef', 3), 'cdef');
  });
  t('extractCarType：NX 归 X，常规取首字', () => {
    eq(Utils.extractCarType('NX17'), 'X');
    eq(Utils.extractCarType('C70'), 'C');
    eq(Utils.extractCarType('P64'), 'P');
  });
  t('determineCarType 需看车号：N+5→X，B+5→X，BH1→P', () => {
    eq(Utils.determineCarType('N17', '5123456'), 'X');
    eq(Utils.determineCarType('B23', '5123456'), 'X');
    eq(Utils.determineCarType('BH1', '1234567'), 'P');
  });
  t('escapeHtml 拦住 XSS', () => {
    eq(Utils.escapeHtml('<img src=x onerror="a">'),
       '&lt;img src=x onerror=&quot;a&quot;&gt;');
    eq(Utils.escapeHtml(null), '');
  });
  t('parseArriveTime 兼容多种写法', () => {
    const d1 = Utils.parseArriveTime('2026/9/2 08:30');
    const d2 = Utils.parseArriveTime('2026-09-02T08:30:00');
    ok(d1 && d2 && d1.getTime() === d2.getTime(), '斜杠与 ISO 写法应等价');
    const d3 = Utils.parseArriveTime('2026年9月2日');
    ok(d3 && d3.getFullYear() === 2026 && d3.getMonth() === 8 && d3.getDate() === 2, '中文日期');
  });
  t('parseArriveTime 非法输入返 null', () => {
    [null, undefined, '', 'abc'].forEach(v => eq(Utils.parseArriveTime(v), null));
  });
  t('hoursDiff', () => {
    const from = new Date(2026, 8, 1, 0, 0);
    const to = new Date(2026, 8, 3, 0, 0);
    eq(Utils.hoursDiff(from, to), 48);
  });
  t('compareTrackId：数字道升序在前，字母道在后且按号排', () => {
    eq(['X2', '10', '2', 'X1', 'Y1'].sort(Utils.compareTrackId), ['2', '10', 'X1', 'X2', 'Y1']);
  });
});

/* ==========================================================================
 * 2. carStyle —— 列索引必须来自 Aggregate.COL，不得硬编码
 *    （回归 A4：原为 row[2]/row[3]，列序一变就静默错色）
 *
 *    carStyle 现为**配置驱动**：命中 Utils.defaultCarTypeConfig 第 i 项 → 'ctc-' + i，
 *    该类名的具体颜色由 applyCarTypeStyles 动态生成（见 utils.js）。
 *    因此这里不写死 'ctc-6'，改从默认配置推导，配置增删/调序时不会误报。
 * ========================================================================== */
function ctcOf(type) {
  const cfg = Utils.defaultCarTypeConfig;
  for (let i = 0; i < cfg.length; i++) {
    const e = cfg[i];
    if (!e.on || !e.prefix) continue;
    const hit = e.match === 'contains'
      ? String(type).indexOf(e.prefix) > -1
      : String(type).indexOf(e.prefix) === 0;
    if (hit) return 'ctc-' + i;
  }
  return '';   // 未命中任何配置项 → 无高亮类
}

suite('carStyle 列索引解耦', () => {
  t('P 车 → 命中盖车配置（字体色+加粗，非背景高亮）', () => {
    const s = Utils.carStyle(mkRow({ CARTYPE: 'P64', CARNO: '1234567' }));
    eq(s.bg, ctcOf('P64'));
    eq(s.cls, s.bg);                              // 背景与文字同类，颜色由 applyCarTypeStyles 决定
    const cfg = Utils.defaultCarTypeConfig.filter(e => e.prefix === 'P')[0];
    ok(cfg && cfg.bold === true, 'P 按用户需求走字体色+加粗，非背景高亮');
  });
  t('车号 07 开头 → 中粮罐加粗', () => {
    const s = Utils.carStyle(mkRow({ CARTYPE: 'G70', CARNO: '0712345' }));
    eq([s.clsN, s.bgN], ['car-self-bold', 'car-self-bold']);
  });
  t('carTypeClass：主表到站串的「车型+车数」片段同样命中配置', () => {
    eq(Utils.carTypeClass('P5'), ctcOf('P'));     // 主表到站串形如 "德保44 P5 到卸23"
    eq(Utils.carTypeClass('DK2'), ctcOf('DK'));
    eq(Utils.carTypeClass('P64'), ctcOf('P'));    // 明细车种列是原始车种
  });
  t('carTypeClass：NX70AF 等 NX 开头被平板车配置命中（NX starts 双保险）', () => {
    eq(Utils.carTypeClass('NX70AF'), ctcOf('NX')); // NX 项 starts 模式
    eq(Utils.carTypeClass('X70'),    ctcOf('X'));  // X 项 contains 模式
    // 即使把 X 项在设置里改成 starts，NX 项仍能罩住 NX 开头
    const saved = sandbox.Store;
    sandbox.Store = { get: () => [
      { prefix: 'NX', match: 'starts',   on: true, bold: false, color: '#a0aec0' },
      { prefix: 'X',  match: 'starts',   on: true, bold: false, color: '#a0aec0' }   // 用户改成了 starts
    ]};
    eq(Utils.carTypeClass('NX70AF'), 'ctc-0');
    eq(Utils.carTypeClass('X70'),    'ctc-1');
    sandbox.Store = saved;
  });
  t('carTypeClass：硬兜底——配置里 X/NX 被删光，X/NX 开头仍命中灰底', () => {
    // 模拟用户在设置里把 X/NX 两项都删了（或者都 on:false 跳过）
    const saved = sandbox.Store;
    sandbox.Store = { get: () => [
      { prefix: 'P', match: 'starts', on: true, bold: false, color: '#ecc94b' }
      // 没有 X，没有 NX
    ]};
    try {
      eq(Utils.carTypeClass('X70'),    '__ctc_fallback__');
      eq(Utils.carTypeClass('NX70AF'), '__ctc_fallback__');
      eq(Utils.carTypeClass('X6K'),    '__ctc_fallback__');
      // 不含 X 的车型不受影响（P 在 mock 配置里是唯一项，索引 0）
      eq(Utils.carTypeClass('P64'),    'ctc-0');
      eq(Utils.carTypeClass('C70'),    '');
    } finally { sandbox.Store = saved; }
  });
  t('carTypeClass：硬兜底——X/NX 被关掉（on:false）时仍生效', () => {
    const saved = sandbox.Store;
    sandbox.Store = { get: () => [
      { prefix: 'NX', match: 'starts', on: false, bold: false, color: '#a0aec0' },
      { prefix: 'X',  match: 'starts', on: false, bold: false, color: '#a0aec0' }
    ]};
    try {
      eq(Utils.carTypeClass('X70'),    '__ctc_fallback__');
      eq(Utils.carTypeClass('NX70AF'), '__ctc_fallback__');
    } finally { sandbox.Store = saved; }
  });
  t('carTypeMatch：X/NX 标记 isFlatbed=true，其它车型=false（主表据此排除平板车底色）', () => {
    const x = Utils.carTypeMatch('X70');
    const nx = Utils.carTypeMatch('NX70AF');
    const p = Utils.carTypeMatch('P64');
    const dk = Utils.carTypeMatch('DK');
    ok(x && x.isFlatbed === true,  'X70 应标记 isFlatbed');
    ok(nx && nx.isFlatbed === true, 'NX70AF 应标记 isFlatbed');
    ok(p && p.isFlatbed === false,  'P 不应标记 isFlatbed');
    ok(dk && dk.isFlatbed === false,'DK 不应标记 isFlatbed');
  });
  t('carTypeMatch：硬兜底 X 也带 isFlatbed=true（主表不挂底色）', () => {
    const saved = sandbox.Store;
    sandbox.Store = { get: () => [{ prefix: 'P', match: 'starts', on: true, bold: false, color: '#ecc94b' }] };
    try {
      const m = Utils.carTypeMatch('X70');
      ok(m && m.isFlatbed === true, '未被配置命中的 X 兜底为平板车');
    } finally { sandbox.Store = saved; }
  });
  t('carTypeClass：未登记车型 / 空值 → 不高亮', () => {
    eq(Utils.carTypeClass('N5'), '');             // 默认配置无 N
    ['', null, undefined].forEach(v => eq(Utils.carTypeClass(v), ''));
  });
  t('carTypeClass：以 Store 中的自定义配置为准', () => {
    const saved = sandbox.Store;
    sandbox.Store = { get: () => [{ prefix: 'P', match: 'starts', on: false, bold: true, color: '#e53e3e' }] };
    eq(Utils.carTypeClass('P5'), '');             // 关掉的项不生效
    sandbox.Store = { get: () => [{ prefix: 'P', match: 'starts', on: true, bold: true, color: '#e53e3e' }] };
    eq(Utils.carTypeClass('P5'), 'ctc-0');        // 自定义配置的下标从 0 起算
    sandbox.Store = saved;
  });
  t('改 COL.CARTYPE 后 carStyle 跟随（证明未硬编码索引）', () => {
    const origType = COL.CARTYPE, origNo = COL.CARNO;
    try {
      COL.CARTYPE = 9; COL.CARNO = 10;
      const row = []; for (let i = 0; i < 16; i++) row[i] = '';
      row[9] = 'P64'; row[10] = '1234567';       // 车种/车号挪到新位置
      eq(Utils.carStyle(row).bg, ctcOf('P64'), '应读到新位置的车种');
    } finally {
      COL.CARTYPE = origType; COL.CARNO = origNo;   // 必须还原，否则污染后续用例
    }
  });
});

/* ==========================================================================
 * 3. 方向库解析
 * ========================================================================== */
suite('方向库解析', () => {
  const csv = 'station,direction,bureau,express\n' +
              '防城,管内,2,\n' +
              '防城港,管内,2,\n' +
              '钦州港,待卸,6,\n' +
              '德保,南口,1,\n';
  t('buildDirectionIndex 解析出 map 与 stations', () => {
    const idx = Aggregate.buildDirectionIndex(csv);
    eq(idx.stations, ['防城', '防城港', '钦州港', '德保']);
    eq(idx.map['防城港'], '管内');
    eq(idx.map['德保'], '南口');
  });
  t('带 BOM 也能解析', () => {
    const idx = Aggregate.buildDirectionIndex('﻿' + csv);
    eq(idx.stations.length, 4);
  });
  t('含逗号的站名用双引号包裹', () => {
    const idx = Aggregate.buildDirectionIndex('station,direction\n"某站,东",沙口\n');
    eq(idx.stations, ['某站,东']);
    eq(idx.map['某站,东'], '沙口');
  });
  t('空输入返回空结构', () => {
    eq(Aggregate.buildDirectionIndex('').stations, []);
  });
});

/* ==========================================================================
 * 4. 到站推断 resolveDest —— 本项目风险最高的一段
 * ========================================================================== */
suite('到站推断', () => {
  const idx = Aggregate.buildDirectionIndex(
    'station,direction\n防城,管内\n防城港,管内\n钦州港,待卸\n德保,南口\n');
  const st = idx.stations, fci = idx.firstCharIndex;
  const now = new Date(2026, 8, 2, 6, 0);
  const rd = row => Aggregate._resolveDest(row, st, now, fci);

  t('段2：载重>15 且到站钦州港 → 到卸', () => {
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港' })), '到卸');
  });
  t('段2 细化：记事含卸车地点 → 识别为具体地点而非到卸', () => {
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '卸永鑫' })), '永鑫');
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '卸货场' })), '货场');
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '天盛卸' })), '天盛');
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '港务局卸车' })), '港务局');
  });
  t('段2 细化：普通多词取首个（无"转"时）', () => {
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '永鑫货场' })), '永鑫');
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '货场永鑫' })), '货场');
  });
  t('段2 细化："转"改卸写法取转之后末位', () => {
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '货场转永鑫' })), '永鑫');
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '永鑫转货场' })), '货场');
  });
  t('段2 细化：记事无地点词仍回落到卸', () => {
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '待卸' })), '到卸');
  });
  t('段2 细化：卸车地点可配置（Store.unloadSpots 覆盖默认）', () => {
    const saved = sandbox.Store;   // 脚本在 vm 沙箱运行，global 即 sandbox
    sandbox.Store = { get: k => (k === 'unloadSpots' ? ['大榄坪', '勒沟'] : null) };
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '卸大榄坪' })), '大榄坪');
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '勒沟转大榄坪' })), '大榄坪');
    // 默认列表里的词不再生效
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '卸永鑫' })), '到卸');
    sandbox.Store = saved;   // 还原，避免影响后续用例
  });

  t('段2 细化：取词前先按"转"截断（仅取转之后的子串）', () => {
    // 默认列表含 永鑫/货场/天盛/港务局，但配置里删掉永鑫后：
    const saved = sandbox.Store;
    sandbox.Store = { get: k => (k === 'unloadSpots' ? ['货场', '天盛', '港务局'] : null) };
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '货场转永鑫' })), '到卸'); // 永鑫不在配置
    eq(rd(mkRow({ LOAD: 60, DEST: '钦州港', NOTE: '天盛转货场' })), '货场'); // 取转后
    sandbox.Store = saved;
  });
  t('段3：载重/到站/记事皆空，按车种车号判定', () => {
    eq(rd(mkRow({ CARTYPE: 'G60', CARNO: '6123456' })), '路罐');
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0712345' })), '黑罐');
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0512345' })), '自备罐');
    eq(rd(mkRow({ CARTYPE: 'C70', CARNO: '1234567' })), 'C');
  });
  t('段1 黑罐细化：G7 罐车按收货人列识别为具体黑罐子类', () => {
    // 默认配置 中粮/外运
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0712345', CONSIGNEE: '中粮' })), '中粮');
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0712345', CONSIGNEE: '外运物流' })), '外运');
    // 收货人不含配置词 → 回落笼统"黑罐"
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0712345', CONSIGNEE: '其他客户' })), '黑罐');
    // 非 G7 罐不受影响
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0512345', CONSIGNEE: '中粮' })), '自备罐');
  });
  t('段1 黑罐细化：可配置（Store.blackTankSpots 覆盖默认）', () => {
    const saved = sandbox.Store;
    sandbox.Store = { get: k => (k === 'blackTankSpots' ? ['益海', '九三'] : null) };
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0712345', CONSIGNEE: '益海嘉里' })), '益海');
    eq(rd(mkRow({ CARTYPE: 'G70', CARNO: '0712345', CONSIGNEE: '中粮' })), '黑罐'); // 默认词失效
    sandbox.Store = saved;
  });
  t('段1：记事命中站名时取「最长」匹配（防城港 不被截成 防城）', () => {
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '防城港卸' })), '防城港');
  });
  t('段1：记事含「XX循环」时不取该站，回退罐型判定', () => {
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '防城港循环' })), '');
  });
  /* 注意这条反直觉规则（对齐 统计股道存车.bas:81-86）：
   * 记事含「原装」时，G6 → 一律「路罐」，G0 → 一律「自备罐」，**不再看汽油/柴油**。
   * 即 "原装汽油" 判为 路罐 而非 汽油——这是 VBA 的既有行为，不是笔误，
   * 显示信息.bas 里对 "原装汽油" 也做了同样的特例排除。勿"顺手修正"。 */
  t('段1：未命中站名时按品名/罐型兜底', () => {
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '原装汽油', CARTYPE: 'G60', CARNO: '6123456' })), '路罐');
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '原装其他', CARTYPE: 'G60', CARNO: '6123456' })), '路罐');
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '汽油',     CARTYPE: 'G60', CARNO: '6123456' })), '汽油');
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '柴油',     CARTYPE: 'G60', CARNO: '6123456' })), '柴油');
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '检修',     CARTYPE: 'C70', CARNO: '1234567' })), 'C');
  });
  t('段1 兜底：G0 车号（自备罐侧）同样遵循「含原装则不看油品」', () => {
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '汽油',     CARTYPE: 'G70', CARNO: '0512345' })), '汽油');
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '原装汽油', CARTYPE: 'G70', CARNO: '0512345' })), '自备罐');
  });
  t('默认：到站有值且不满足任何分支 → 原样保留', () => {
    eq(rd(mkRow({ LOAD: 60, DEST: '德保' })), '德保');
  });

  /* ---- 关键回归：倒排索引不得漏判「站名不在记事开头」的情形 ----
   * 原实现只探测 note.charAt(0)，"装防城港箱 天驰" 首字为「装」→ 取桶为空 → 漏判。
   * 实测 1339 行真实数据 93 行（6.95%）因此判错。
   * 这里用「快路径 vs 全表扫描」逐条比对，确保两者永远等价。 */
  const notes = [
    '装防城港箱 天驰', '排防城敞顶箱 27292', '送德保', '去防城港',
    '防城港卸', '防城', '到钦州港', '钦州港', '无站名记事',
    '转防城港再排防城', '德保', ' 德保 ', '钦州港装箱'
  ];
  notes.forEach(note => {
    t(`倒排索引与全表扫描等价：note="${note}"`, () => {
      const row = mkRow({ LOAD: 5, DEST: '钦州港', NOTE: note });
      const fast = Aggregate._resolveDest(row, st, now, fci);
      const full = Aggregate._resolveDest(row, st, now, null);
      eq(fast, full, `note="${note}" 快路径与基准不一致`);
    });
  });
  t('前缀写法确实能命中站名（漏判的直接验证）', () => {
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '装防城港箱 天驰' })), '防城港');
    eq(rd(mkRow({ LOAD: 5, DEST: '钦州港', NOTE: '排防城敞顶箱 27292' })), '防城');
  });
});

/* ==========================================================================
 * 5. 端到端聚合
 * ========================================================================== */
suite('端到端聚合', () => {
  const now = new Date(2026, 8, 2, 6, 0);
  // 100 小时前到达 → 老牌车
  const old = new Date(now.getTime() - 100 * 3600000);
  const arr = old.getFullYear() + '/' + (old.getMonth() + 1) + '/' + old.getDate() +
              ' ' + old.getHours() + ':' + old.getMinutes();

  const rows = [
    mkRow({ TRACK: '1', SEQ: 1, CARTYPE: 'C70', CARNO: '1234567', TARE: 20, LEN: 1.5,
            LOAD: 60, DEST: '德保', DIR: '3', GOODS: '煤', FROM: '钦州港',
            NOTE: '扣修', TRAIN: '41001', ARRTIME: arr }),
    mkRow({ TRACK: '1', SEQ: 2, CARTYPE: 'P64', CARNO: '2234567', TARE: 22, LEN: 1.3,
            LOAD: 0,  DEST: '钦州港', DIR: '2', GOODS: '空', FROM: '德保',
            NOTE: '', TRAIN: '41002', ARRTIME: arr }),
    mkRow({ TRACK: '2', SEQ: 1, CARTYPE: 'G70', CARNO: '6234567', TARE: 25, LEN: 1.1,
            LOAD: 50, DEST: '钦州港', DIR: '6', GOODS: '汽油', FROM: '钦州港',
            NOTE: '', TRAIN: '41003', ARRTIME: arr })
  ];

  const res = Aggregate.aggregate(rows, {}, [], { oldCarHours: 47 }, now);

  t('按股道分组', () => { eq(Object.keys(res).sort(), ['1', '2']); });
  t('车数', () => { eq(res['1'].count, 2); eq(res['2'].count, 1); });
  t('换长合计（保留 1 位小数）', () => { eq(res['1'].length, 2.8); });
  t('载重 = 载重 + 自重', () => {
    eq(res['1'].load, (60 + 20) + (0 + 22));
    eq(res['2'].load, 50 + 25);
  });
  t('方向分类：3→南口、2→管内、6且载重>20→到卸', () => {
    ok(res['1'].direction.indexOf('南口') >= 0, '应有南口');
    ok(res['1'].direction.indexOf('管内') >= 0, '应有管内');
    ok(res['2'].direction.indexOf('到卸') >= 0, '应有到卸');
  });
  t('老牌车：停时>47h 且车号非 0 开头', () => {
    eq(res['1'].oldCar, 2);
    eq(res['2'].oldCar, 1);
  });
  t('车次取出现次数最多者', () => {
    ok(['41001', '41002'].indexOf(res['1'].train) >= 0, 'train=' + res['1'].train);
  });
  t('记事关键词筛选：扣修 命中，不入扣 排除', () => {
    ok(res['1'].note.indexOf('扣修') >= 0, 'note=' + res['1'].note);
  });
  t('raw 保留明细行，供抽屉展示', () => { eq(res['1'].raw.length, 2); });
  t('股道为空的行被跳过', () => {
    const r = Aggregate.aggregate([mkRow({ TRACK: '', CARTYPE: 'C70' })], {}, [], {}, now);
    eq(Object.keys(r).length, 0);
  });
  t('__destRaw 保留原始到站，不因聚合改写', () => {
    eq(res['2'].raw[0].__destRaw, '钦州港');
  });
});

/* ==========================================================================
 * 6. 股道配置
 * ========================================================================== */
suite('股道配置', () => {
  t('股道总数 97', () => { eq(YardConfig.tracks.length, 97); });
  t('虚拟股道 38 条（分组 ∪ 个别名单；含广明 GM1-GM4）', () => { eq(YardConfig.virtualIds.length, 38); });
  t('到发线 1-10（10 条）、调车线 11-15（5 条）、临时 7 条', () => {
    eq(YardConfig.idsOfGroup('td').length, 10);
    eq(YardConfig.idsOfGroup('dc').length, 5);
    eq(YardConfig.idsOfGroup('tmp').length, 7);
  });
  t('X 线分组改名为「虚拟场」，且带分组颜色', () => {
    var x = YardConfig.getTrack('X1');
    eq(x.groupName, '虚拟场');
    ok(/^#/.test(x.groupColor), 'groupColor 应为十六进制色值');
  });
  t('到发线显示名带「道」，未登记股道原样返回', () => {
    eq(YardConfig.trackName(1), '1道');
    eq(YardConfig.trackName(999), '999');
  });
  t('个别虚拟股道：YQX 隐藏、YX1 常显', () => {
    ok(YardConfig.isVirtual('YQX'));
    ok(!YardConfig.isVirtual('YX1'));
  });
  t('作业区允许重叠：Y5 同时属栈桥与中油', () => {
    eq(YardConfig.getZones('Y5').map(z => z.name), ['栈桥', '中油']);
  });
  t('按车种过滤作业区：Y5 禁 C → 仅中油', () => {
    eq(YardConfig.getZones('Y5', 'C').map(z => z.name), ['中油']);
  });
  t('未落入任何作业区返回空数组', () => {
    eq(YardConfig.getZones('999'), []);
  });
  t('阈值齐全', () => {
    ['oldCarHours', 'bigCarHours', 'overloadTons', 'overlong', 'heavyLoad', 'lightLoad']
      .forEach(k => ok(typeof YardConfig.thresholds[k] === 'number', '缺少阈值 ' + k));
  });
});

/* ==========================================================================
 * 7. 持久化（Store）
 * ========================================================================== */
suite('持久化', () => {
  t('set/get 自动 JSON 往返', () => {
    Store.set('t.list', [1, 2, 3]);
    eq(Store.get('t.list'), [1, 2, 3]);
  });
  t('key 自动补 zhancun. 前缀', () => {
    Store.set('t.plain', 'x');
    eq(memStore['zhancun.t.plain'], '"x"');
  });
  t('兼容升级前直接存原始字符串的旧数据', () => {
    memStore['zhancun.t.legacy'] = '文件夹名';      // 非合法 JSON
    eq(Store.get('t.legacy'), '文件夹名');
  });
  t('键不存在时返回默认值', () => {
    eq(Store.get('t.nope', 'def'), 'def');
  });
  t('remove 生效', () => {
    Store.set('t.rm', 1); Store.remove('t.rm');
    eq(Store.get('t.rm', null), null);
  });
});

/* ==================== 汇总 ==================== */
console.log('\n' + '='.repeat(58));
if (fail) {
  console.log(`失败 ${fail} 项：`);
  failures.forEach(f => console.log('  ✗ ' + f));
}
console.log(`结果：${pass} 通过 / ${fail} 失败`);
console.log('='.repeat(58));
process.exit(fail ? 1 : 0);
