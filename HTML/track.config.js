/**
 * 股道存车 - 股道配置文件
 * ============================================================================
 * 加载方式： <script src="track.config.js" charset="utf-8"></script>
 * 暴露全局： window.YardConfig
 *
 * 【如何增删股道】见文件底部「快速修改指引」
 * 【注意】数组顺序 = 表格显示顺序，请勿随意调整已有分组的先后
 * ============================================================================
 */
(function (global) {
  'use strict';

  /* ==========================================================================
   * 一、分组定义
   * --------------------------------------------------------------------------
   * id        分组标识，供下方 tracks 引用
   * name      分组名称（可按现场习惯修改）
   * suffix    显示后缀。Excel 中纯数字股道显示为 "1G"，源于数字格式 0"G"
   * occupancy 是否参与「股道占用情况」统计（换长合计 / 容量标色）
   * virtual   是否虚拟股道。虚拟股道可由工具栏「虚拟股道」开关一键显示/隐藏
   * ========================================================================== */
  var GROUPS = [
    { id: 'td',  name: '到发线', suffix: '道', occupancy: true  },
    { id: 'b',   name: 'B线',    suffix: '',  occupancy: true  },
    { id: 'x',   name: 'X线',    suffix: '',  occupancy: true, virtual: true },
    { id: 'h',   name: 'H线',    suffix: '',  occupancy: false },
    { id: 'l',   name: 'L线',    suffix: '',  occupancy: false },
    { id: 'zl',  name: 'ZL线',   suffix: '',  occupancy: false },
    { id: 'lz',  name: 'LZ线',   suffix: '',  occupancy: false },
    { id: 'g',   name: 'G线',    suffix: '',  occupancy: false },
    { id: 'yx',  name: 'YX线',   suffix: '',  occupancy: false },
    { id: 'ts',  name: 'TS线',   suffix: '',  occupancy: false },
    { id: 'sh',  name: 'SH线',   suffix: '',  occupancy: false },
    { id: 'cz',  name: 'CZ线',   suffix: '',  occupancy: false },
    { id: 'gt',  name: 'GT线',   suffix: '',  occupancy: false },
    { id: 'gm',  name: 'GM线',   suffix: '',  occupancy: false },
    { id: 'dy',  name: 'DY线',   suffix: '',  occupancy: false, virtual: true },
    { id: 'tsy', name: 'TSY线',  suffix: '',  occupancy: false, virtual: true },
    { id: 'yh',  name: 'YH线',   suffix: '',  occupancy: false, virtual: true },
    { id: 'y',   name: 'Y线',    suffix: '',  occupancy: false },
    { id: 'yqx', name: 'YQX线',  suffix: '',  occupancy: false },
    { id: 'zxx', name: 'ZXX线',  suffix: '',  occupancy: false },
    { id: 'tmp', name: '临时',   suffix: '',  occupancy: false, virtual: true }
  ];

  /* ==========================================================================
   * 二、股道清单
   * --------------------------------------------------------------------------
   * 支持三种写法，可混用：
   *   ids: '1-15'              → 1,2,3 ... 15
   *   ids: 'Y1-Y16'            → Y1,Y2 ... Y16   （同前缀数字范围）
   *   ids: ['B1','B2']         → B1,B2           （逐个列出）
   *   ids: ['1-3', '7', 'X1']  → 1,2,3,7,X1      （数组可再含范围）
   *
   * id 必须与 SMIS 导出 xls 中「股道」列的取值完全一致
   * ========================================================================== */
  var TRACK_DEFS = [
    { group: 'td',  ids: '1-15'      },
    { group: 'b',   ids: ['B1', 'B2'] },
    { group: 'x',   ids: 'X1-X15'    },
    { group: 'h',   ids: 'H1-H5'     },
    { group: 'l',   ids: 'L1-L3'     },
    { group: 'zl',  ids: 'ZL1-ZL3'   },
    { group: 'lz',  ids: ['LZ']      },
    { group: 'g',   ids: ['G1', 'G2']},
    { group: 'yx',  ids: 'YX1-YX3'   },
    { group: 'ts',  ids: ['TS1', 'TS2'] },
    { group: 'sh',  ids: 'SH1-SH3'   },
    { group: 'cz',  ids: ['CZ3', 'CZ4'] },
    { group: 'gt',  ids: ['GT1', 'GT2'] },
    { group: 'gm',  ids: 'GM1-GM4'   },
    { group: 'dy',  ids: 'DY1-DY4'   },
    { group: 'tsy', ids: 'TSY1-TSY4' },
    { group: 'yh',  ids: ['YH1', 'YH2'] },
    { group: 'y',   ids: 'Y1-Y16'    },
    { group: 'yqx', ids: ['YQX']     },
    { group: 'zxx', ids: ['ZXX']     },
    { group: 'tmp', ids: '##1-##7'   }
  ];

  /* ==========================================================================
   * 三、作业区 / 车型禁入规则
   * --------------------------------------------------------------------------
   * 原 VBA 以硬编码行号实现（B43:B66 等），此处改为按股道区间声明。
   * from / to 取上方清单中的股道 id，表示「从 from 到 to 的连续区间」。
   *
   * forbid   禁止进入的车种字母（P=盖车, C=高边敞车...）
   * mark     E列（车种）标记色：'red'=红  'pink'=粉红
   * markJ    是否同时把 J列（到达车次）标红
   * ========================================================================== */
  var ZONES = [
    { id: 'yingling', name: '鹰岭', from: 'YX1', to: 'TSY4', forbid: ['P'], mark: 'red',  markJ: true  },
    { id: 'zhanqiao', name: '栈桥', from: 'Y1',  to: 'Y14',  forbid: ['P'], mark: 'pink', markJ: false },
    { id: 'zhongyou', name: '中油', from: 'YH1', to: 'YQX',  forbid: ['C'], mark: 'red',  markJ: true  }
  ];

  /* ==========================================================================
   * 四、判定阈值（均来自原 VBA）
   * ========================================================================== */
  var THRESHOLDS = {
    oldCarHours: 47,   // 停时超过此小时数 → 老牌车
    bigCarHours: 49,   // 停时超过此小时数 → 大点车标黄
    overloadTons: 5000,// 载重超过此吨数 → 超吨标红
    overlong: 70,      // 换长超过此值 → 超长标红
    heavyLoad: 20,     // 载重大于此值 → 判为重车
    lightLoad: 15      // 载重小于此值 → 判为轻车/空车
  };

  /* ==========================================================================
   * 五、特殊标记屏蔽名单
   * --------------------------------------------------------------------------
   * 列出后，该股道不参与「汽油股道标黄」与「点后开」标记。
   *
   * 【历史】VBA 显示信息.bas 中曾有：
   *   If Mid(key, 1, 1) = "X" And Mid(key, 2) > 10 Then   '屏蔽11-15道
   * 该分支只是权宜之计——当时 Sheet6 的股道清单只到 X10，
   * .Range("B2:B92").Find("X11") 找不到会返回 Nothing，取 .Row 即报错。
   * 现在清单已扩至 X15，X11~X15 属于正式股道，应与其他股道一样正常标记，
   * 故此处置空。
   *
   * 若日后确有个别股道不需要这两个标记，在此列出 id 即可。
   * ========================================================================== */
  var SUPPRESS_MARK = [
    // 'X11', 'X12', 'X13', 'X14', 'X15'
  ];

  /* ==========================================================================
   * 六、个别虚拟股道名单（股道级隐藏）
   * --------------------------------------------------------------------------
   * 与「分组 virtual」互补：分组 virtual 整组隐藏（如 X1-X15、DY1-4、TSY1-4、
   * YH1-2、##1-#7）；此处列出仅需单独隐藏的个别股道。
   *   · YX1  ：YX 线只需隐藏 YX1（YX2/YX3 仍常显）
   *   · YQX  ：YQX 线（单条）
   *   · ZXX  ：ZXX 线（单条，表格原无，本次新增）
   * 加入此名单即与分组虚拟股道共用工具栏「隐藏非常用股道」开关。
   * ========================================================================== */
  var VIRTUAL_IDS = [
    'YX1', 'YQX', 'ZXX'
  ];

  /* ==========================================================================
   * 以下为展开逻辑，一般无需修改
   * ========================================================================== */

  /** 把 '1-15' / 'Y1-Y16' / 数组 等简写展开为股道 id 字符串数组 */
  function expandIds(spec) {
    var out = [];
    (Array.isArray(spec) ? spec : [spec]).forEach(function (item) {
      if (item === null || item === undefined) return;
      if (Array.isArray(item)) { out = out.concat(expandIds(item)); return; }

      var s = String(item).trim();
      if (!s) return;

      // 纯数字范围：1-15
      var m = /^(\d+)\s*-\s*(\d+)$/.exec(s);
      if (m) { out = out.concat(seq(+m[1], +m[2], '')); return; }

      // 同前缀数字范围：Y1-Y16 / ##1-##7
      var m2 = /^(\D+)(\d+)\s*-\s*\1(\d+)$/.exec(s);
      if (m2) { out = out.concat(seq(+m2[2], +m2[3], m2[1])); return; }

      out.push(s);
    });
    return out;
  }

  function seq(a, b, prefix) {
    var out = [], step = a <= b ? 1 : -1;
    for (var i = a; ; i += step) {
      out.push(prefix + i);
      if (i === b) break;
    }
    return out;
  }

  // 构建分组索引
  var groupMap = {};
  GROUPS.forEach(function (g) { groupMap[g.id] = g; });

  // 展开股道清单
  var tracks = [];
  TRACK_DEFS.forEach(function (def) {
    var g = groupMap[def.group];
    if (!g) throw new Error('[track.config.js] 未定义的分组: ' + def.group);
    expandIds(def.ids).forEach(function (id) {
      var virtual = !!g.virtual || VIRTUAL_IDS.indexOf(id) >= 0;  // 分组虚拟 ∪ 个别名单
      tracks.push({
        id: id,
        name: id + (g.suffix || ''),   // 显示名，如 "1G"
        group: g.id,
        groupName: g.name,
        occupancy: !!g.occupancy,
        virtual: virtual
      });
    });
  });

  // 股道 id → 对象 的快速索引
  var trackMap = {};
  tracks.forEach(function (t, i) { t.index = i; trackMap[t.id] = t; });

  // 展开作业区：把 from/to 解析为实际股道 id 列表
  var zones = ZONES.map(function (z) {
    var a = trackMap[z.from], b = trackMap[z.to];
    if (!a || !b) throw new Error('[track.config.js] 作业区 ' + z.name + ' 的起止股道不存在');
    var from = Math.min(a.index, b.index), to = Math.max(a.index, b.index);
    return {
      id: z.id,
      name: z.name,
      forbid: z.forbid || [],
      mark: z.mark,
      markJ: !!z.markJ,
      ids: tracks.slice(from, to + 1).map(function (t) { return t.id; })
    };
  });

  // 股道 id → 所属作业区列表
  // 注意：作业区之间允许重叠。例如 Y1~Y14 既在「栈桥」（禁盖车P）又在「中油」（禁高边C），
  // 两条规则针对不同车种、各自独立生效，与原 VBA 的两个独立 If 块一致，不可合并。
  var zonesOfTrack = {};
  zones.forEach(function (z) {
    z.ids.forEach(function (id) {
      (zonesOfTrack[id] || (zonesOfTrack[id] = [])).push(z);
    });
  });

  global.YardConfig = {
    version: '1.0',
    tracks: tracks,
    groups: GROUPS,
    zones: zones,
    thresholds: THRESHOLDS,
    suppressMark: SUPPRESS_MARK,

    /** 按 id 取股道定义 */
    getTrack: function (id) { return trackMap[String(id)] || null; },
    /** 判断股道是否在配置清单内 */
    hasTrack: function (id) { return Object.prototype.hasOwnProperty.call(trackMap, String(id)); },
    /**
     * 取股道所属的全部作业区（未落入任何作业区时返回空数组）
     * @param {string} id 股道 id
     * @param {string} [carType] 车种字母（如 'P'、'C'），传入时仅返回禁止该车种进入的作业区
     */
    getZones: function (id, carType) {
      var list = zonesOfTrack[String(id)] || [];
      if (!carType) return list;
      return list.filter(function (z) { return z.forbid.indexOf(carType) >= 0; });
    },
    /** 该股道是否屏蔽汽油/点后开标记 */
    isMarkSuppressed: function (id) { return SUPPRESS_MARK.indexOf(String(id)) >= 0; },
    /** 该股道是否为虚拟股道（可用工具栏开关显示/隐藏） */
    isVirtual: function (id) {
      var t = trackMap[String(id)];
      return !!(t && t.virtual);
    },
    /** 参与占用统计的股道列表 */
    occupancyTracks: tracks.filter(function (t) { return t.occupancy; }),
    /** 虚拟股道 id 列表 */
    virtualIds: tracks.filter(function (t) { return t.virtual; }).map(function (t) { return t.id; }),
    /** 全部股道 id */
    ids: tracks.map(function (t) { return t.id; })
  };

  /* ==========================================================================
   * 快速修改指引
   * --------------------------------------------------------------------------
   * 新增股道（在已有分组内追加）
   *   例：到发线由 1-15 扩为 1-16   →   { group: 'td', ids: '1-16' }
   *   例：Y 线增加 Y17               →   { group: 'y', ids: 'Y1-Y17' }
   *   例：新增零散股道               →   { group: 'tmp', ids: ['##8', '##9'] }
   *
   * 删除股道
   *   缩小范围或直接从数组中移除即可，如 'X1-X15' → 'X1-X10'
   *
   * 新增一个分组
   *   1) 在 GROUPS 中添加 { id:'q', name:'Q线', suffix:'', occupancy:false }
   *   2) 在 TRACK_DEFS 中添加 { group:'q', ids:'Q1-Q5' }
   *   注意插入位置决定表格显示顺序
   *
   * 调整作业区禁入规则
   *   修改 ZONES 中对应项的 from/to/forbid/mark
   *
   * 修改判定阈值
   *   修改 THRESHOLDS，如老牌车由 47 小时改为 36 小时
   * ========================================================================== */
})(window);
