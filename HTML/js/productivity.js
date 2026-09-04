(function (global) {
  if (!global.Utils) return;
  var Utils = global.Utils;
  var Store = global.Store;
  var UI = global.UI;
  var NAMES = ['1调', '2调', '3调'];
  var STORE_KEY = 'productivity';

  function $(id) { return document.getElementById(id); }

  function todayStr() {
    var d = new Date();
    return d.getFullYear() + '-' +
      String(d.getMonth() + 1).padStart(2, '0') + '-' +
      String(d.getDate()).padStart(2, '0');
  }

  function defaultData() {
    var make = function () { return [{h:'',m:''},{h:'',m:''},{h:'',m:''}]; };
    return {
      date: todayStr(),
      safeStart: '2019-03-05',
      nightH: 780,
      dayH: 660,
      night: make(),
      day: make(),
      notes: ''
    };
  }

  function load() {
    var d = defaultData();
    if (Store && Store.get) {
      var s = Store.get(STORE_KEY, null);
      if (s && typeof s === 'object') {
        if (s.date != null) d.date = s.date;
        if (s.safeStart != null) d.safeStart = s.safeStart;
        if (s.dayH != null) d.dayH = s.dayH;
        if (Array.isArray(s.night)) d.night = s.night;
        if (Array.isArray(s.day)) d.day = s.day;
        if (s.safety != null) d.safety = s.safety;
        if (s.notes != null) d.notes = s.notes;
      }
    }
    return d;
  }

  function save(d) {
    if (Store && Store.set) Store.set(STORE_KEY, d);
  }

  function fmt(n, d) {
    if (!isFinite(n) || isNaN(n)) return '';
    return n.toFixed(d);
  }

  function val(id) {
    var el = $(id);
    return el ? parseFloat(el.value) || 0 : 0;
  }
  function str(id) {
    var el = $(id);
    return el ? el.value : '';
  }
  function parseDate(s) {
    if (!s) return null;
    var p = String(s).split('-');
    if (p.length !== 3) return null;
    var dt = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
    if (isNaN(dt.getTime())) return null;
    return dt;
  }
  function diffDays(a, b) {
    var da = parseDate(a), db = parseDate(b);
    if (!da || !db) return null;
    return Math.round((da.getTime() - db.getTime()) / 86400000);
  }
  function setText(id, txt) {
    var el = $(id);
    if (el) el.textContent = (txt == null ? '' : txt);
  }
  function setVal(id, v) {
    var el = $(id);
    if (el) el.value = (v == null ? '' : v);
  }

  function collect() {
    var d = defaultData();
    d.date = str('prodDate');
    d.safeStart = str('prodSafeStart');
    d.nightH = val('prodNightH');
    d.dayH = val('prodDayH');
    d.safety = str('prodSafety');
    d.notes = str('prodNotes');
    for (var i = 0; i < 3; i++) {
      d.night[i] = { h: val('nH' + i), m: val('nM' + i) };
      d.day[i]   = { h: val('dH' + i), m: val('dM' + i) };
    }
    return d;
  }

  // 生成班次区段行（夜班 / 白班）
  function shiftRows(section, prefix, arr) {
    var html = '';
    for (var i = 0; i < 3; i++) {
      var sectionCell = i === 0 ? '<td class="prod-section-label" rowspan="3">' + section + '</td>' : '';
      var totalCell   = i === 0 ? '<td class="prod-calc" rowspan="3" id="' + prefix + 'TotalAll">0</td>' : '';
      html += '<tr>' + sectionCell +
        '<td>' + NAMES[i] + '</td>' +
        '<td><input class="prod-in" type="number" min="0" id="' + prefix + 'H' + i + '"></td>' +
        '<td><input class="prod-in" type="number" min="0" id="' + prefix + 'M' + i + '"></td>' +
        '<td class="prod-calc" id="' + prefix + 'Score' + i + '">0</td>' +
        '<td class="prod-calc" id="' + prefix + 'Util' + i + '">0%</td>' +
        totalCell +
        '<td class="prod-calc" id="' + prefix + 'Remark' + i + '">0</td>' +
        '</tr>';
    }
    return html;
  }

  // 生成合计区段行
  function sumRows() {
    var html = '';
    for (var i = 0; i < 3; i++) {
      var sectionCell = i === 0 ? '<td class="prod-section-label" rowspan="3">合计</td>' : '';
      var totalCell   = i === 0 ? '<td class="prod-calc" rowspan="3" id="sTotalAll">0</td>' : '';
      html += '<tr>' + sectionCell +
        '<td class="prod-sum-label">' + NAMES[i] + '生产时间合计</td>' +
        '<td class="prod-calc" id="sHook' + i + '">0</td>' +
        '<td class="prod-calc" id="sMin' + i + '">0</td>' +
        '<td class="prod-calc" id="sScore' + i + '">0</td>' +
        '<td class="prod-calc" id="sUtil' + i + '">0%</td>' +
        totalCell +
        '<td class="prod-calc" id="sRemark' + i + '">0</td>' +
        '</tr>';
    }
    return html;
  }

  function render() {
    var d = load();
    setVal('prodDate', d.date);
    setVal('prodSafeStart', d.safeStart);
    setVal('prodNightH', d.nightH);
    setVal('prodDayH', d.dayH);
    setVal('prodSafety', d.safety);
    setVal('prodNotes', d.notes);

    var tbody = $('prodMainBody');
    if (tbody) {
      tbody.innerHTML = shiftRows('夜班', 'n', d.night) +
                        shiftRows('白班', 'd', d.day) +
                        sumRows();
    }
    for (var i = 0; i < 3; i++) {
      setVal('nH' + i, d.night[i].h);
      setVal('nM' + i, d.night[i].m);
      setVal('dH' + i, d.day[i].h);
      setVal('dM' + i, d.day[i].m);
    }
    recompute();
  }

  function recompute() {
    var d = collect();
    // 班次时长直接按分钟计算
    var nm = d.nightH || 0;
    var dm = d.dayH || 0;

    var nTotal = d.night[0].h + d.night[1].h + d.night[2].h;
    var dTotal = d.day[0].h + d.day[1].h + d.day[2].h;
    var gTotal = nTotal + dTotal;

    setText('nTotalAll', nTotal);
    setText('dTotalAll', dTotal);
    setText('sTotalAll', gTotal);

    for (var i = 0; i < 3; i++) {
      var nh = d.night[i].h, nmm = d.night[i].m;
      setText('nScore' + i, nh ? fmt(nmm / nh, 1) : '0');
      setText('nUtil' + i, nm ? fmt(nmm / nm * 100, 1) + '%' : '0%');
      setText('nRemark' + i, fmt(nmm / 60, 1));

      var dh = d.day[i].h, dmm = d.day[i].m;
      setText('dScore' + i, dh ? fmt(dmm / dh, 1) : '0');
      setText('dUtil' + i, dm ? fmt(dmm / dm * 100, 1) + '%' : '0%');
      setText('dRemark' + i, fmt(dmm / 60, 1));

      var sm = nmm + dmm, sh = nh + dh;
      setText('sHook' + i, sh);
      setText('sMin' + i, fmt(sm / 60, 1));
      setText('sScore' + i, sh ? fmt(sm / sh, 1) : '0');
      setText('sUtil' + i, (nm + dm) ? fmt(sm / (nm + dm) * 100, 1) + '%' : '0%');
      setText('sRemark' + i, fmt(sm / 60, 1));
    }

    if ($('prodTitleText')) $('prodTitleText').textContent = '钦州港站 ' + (d.date || '') + ' 生产效率表';
    setVal('prodSafety', diffDays(d.date, d.safeStart));
    save(d);
  }

  function open() {
    render();
    if (UI && UI.Modal) UI.Modal.open('modalProductivity');
  }

  function init() {
    var modal = $('modalProductivity');
    if (!modal) return;
    if (UI && UI.Modal) UI.Modal.register('modalProductivity');

    // 任意输入变化（含记事、日期、班次时长、钩数/生产时间）都重算并持久保存
    function onEdit(e) {
      var t = e.target;
      if (!t) return;
      if (t.classList && t.classList.contains('prod-in')) return recompute();
      if (t.id === 'prodNightH' || t.id === 'prodDayH' ||
          t.id === 'prodDate' || t.id === 'prodSafeStart' || t.id === 'prodNotes') {
        recompute();
      }
    }
    modal.addEventListener('input', onEdit);
    modal.addEventListener('change', onEdit);

    global.Productivity = { open: open };
  }

  if (document.readyState !== 'loading') init();
  else document.addEventListener('DOMContentLoaded', init);
})(window);