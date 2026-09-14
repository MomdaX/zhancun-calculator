/**
 * columns.js —— 列定义与"编好"列配置（从 app.js 抽出，P2-7 第二步）
 *
 * 纯常量，零逻辑、零 DOM。仅依赖 Aggregate.COL（index.html 中 aggregate.js 已先加载）。
 * 通过 global 暴露，供 app.js 主表渲染 / 明细抽屉直接按原名字引用。
 */
(function (global) {
  'use strict';

  /* ============================ 列定义 ============================ */
  var COLUMNS = [
    { key: 'note',      title: '注意事项',       width: 112, cls: 'col-a' },
    // 分组合并列：在股道列前，按分组跨行合并（rowspan），仅每组首行输出单元格；文字竖排
    { key: 'group',     title: '',               width: 32,  cls: 'col-b-group' },
    { key: 'track',     title: '股道',           width: 66,  cls: 'col-b track' },
    { key: 'effLen',    title: '有效长',         width: 64, num: true, cls: 'mid' },
    { key: 'count',     title: '车数',           width: 48,  num: true, cls: 'mid' },
    { key: 'carTypes',  title: '车种',           width: 124 },
    { key: 'length',    title: '换长',           width: 58,  num: true, cls: 'mid' },
    { key: 'dest',      title: '车辆信息',       width: 340, dest: true },
    // 编好车次：按股道录入的编组车次（双击编辑）。数据源是 Store.KEYS.readyTrains，
    // 与 31814 报表「待发股道车次」、发车流程「车次」共用同一份持久化（见 app.js checiOf/setCheci）
    { key: 'checi',     title: '编好车次',       width: 88 },
    // 发送：只承载「编好」伪元素按钮，不显示任何数据（原「重车」列，数字已取消）
    { key: 'send',      title: '发送',           width: 56 },
    { key: 'train',     title: '到达车次',       width: 76 },
    { key: 'load',      title: '载重',           width: 72,  num: true, cls: 'mid' },
    { key: 'oldCar',    title: '老牌车',         width: 60,  num: true, cls: 'mid' }
  ];

  /* 「编好」按钮：挂在表头 th[10]（XPath 1 起 → COLUMNS[9] = send「发送」）
   * 的合格数据行。选中该行后以 ::after 伪元素浮现，点击打开发车作业全流程。
   * 原先挂在 th[9]（empty「编好车次」），该列改为录入车次后让出按钮位置。
   * 用 COLUMNS 下标推导而非写死 key：将来增删列时自动跟随。
   * 注：只决定"标记/命中哪个单元格"，选中资格（非 blank、股道在区间内）仍在 app.js，未改动。 */
  var BIANHAO_COL     = COLUMNS[9] ? COLUMNS[9].key : 'send';
  /* 生效股道区间（按 track.config.js 的 TRACK_DEFS 顺序取 index 判定）：
   * 1道(1) … X15，含两端。写股道 id 而非行号——行号会随分组/空线显隐而浮动。 */
  var BIANHAO_FROM    = '1';
  var BIANHAO_TO      = 'X15';
  var BIANHAO_EXCLUDE = ['B1', 'B2'];   // 排除边修线（B1/B2 不参与编好）
  var BIANHAO_HIT_W   = 36;   // 伪元素命中宽度：自单元格左缘起算（px）

  // fmt: track = 股道显示名（到发线加"道"）  dest = 到站按方向着色
  // col 取 Aggregate.COL 常量：SMIS 导出列序变动时只需改 aggregate.js 一处。
  //
  // 【关于旧注释的更正】此处原写作「aggregate.js 在 app.js 之后才加载，不能在
  // IIFE 顶部引用 COL，必须做成函数」，与 index.html 的实际加载顺序相反：
  // aggregate.js 先于 app.js。且 app.js 中的 `var COL = Aggregate.COL;` 本就是顶层引用，
  // 早已证明该限制不存在。列定义是常量，故直接写成常量数组，无需再包一层函数。
  var DETAIL_COLS = [
    { col: Aggregate.COL.TRACK,   t: '股道', w: 54, fmt: 'track' },
    { col: Aggregate.COL.SEQ,     t: '顺', w: 34 },
    // 车号在车种之前（现场按车号点车，先找号再看车型）
    { col: Aggregate.COL.CARNO, t: '车号', w: 78 }, { col: Aggregate.COL.CARTYPE, t: '车种', w: 68 },
    { col: Aggregate.COL.TARE,    t: '自重', w: 52 }, { col: Aggregate.COL.LEN,   t: '换长', w: 50, cls: 'mid' },
    { col: Aggregate.COL.LOAD,    t: '载重', w: 56, cls: 'mid' },
    { col: Aggregate.COL.DEST,    t: '到站', w: 120, disp: 'processed' },
    { col: Aggregate.COL.DIR,     t: '方向', w: 42 }, { col: Aggregate.COL.GOODS, t: '品名', w: 90 },
    { col: Aggregate.COL.FROM,    t: '发站', w: 110, disp: 'raw' },
    { col: Aggregate.COL.NOTE,    t: '记事', w: 170 },
    { col: Aggregate.COL.TRAIN,   t: '车次', w: 62 }, { col: Aggregate.COL.CONSIGNEE, t: '收货人', w: 100 }, { col: Aggregate.COL.ARRTIME, t: '到达时间', w: 128 }
  ];

  global.COLUMNS = COLUMNS;
  global.DETAIL_COLS = DETAIL_COLS;
  global.BIANHAO_COL = BIANHAO_COL;
  global.BIANHAO_FROM = BIANHAO_FROM;
  global.BIANHAO_TO = BIANHAO_TO;
  global.BIANHAO_EXCLUDE = BIANHAO_EXCLUDE;
  global.BIANHAO_HIT_W = BIANHAO_HIT_W;

})(window);
