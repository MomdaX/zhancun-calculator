var ctp = null;
var lastZ1 = "";

function OnLoad(ribbon) {
  console.log('铁路地图加载项已就绪');
  startPollZ1();
}

function OnShowMap() {
  var baseUrl = window.location.origin + window.location.pathname.replace(/\/[^/]*$/, '');
  ctp = wps.CreateTaskPane(baseUrl + '/chinamap.html', '铁路地图');
  ctp.Visible = true;
}

function startPollZ1() {
  setInterval(function () {
    if (!ctp) return;
    try {
      var val = wps.Application.ActiveSheet.Range("Z1").Value2;
      if (!val || String(val).trim() === "" || String(val).trim() === lastZ1) return;
      lastZ1 = String(val).trim();
      console.log("检测到站名: " + lastZ1);
      sendToCpt(lastZ1);
    } catch (e) {}
  }, 500);
}

function sendToCpt(stationName) {
  if (!ctp) { console.log("ctp 为空"); return; }
  var js = "generatePath('" + stationName.replace(/'/g, "\\'") + "')";
  console.log("执行: " + js);
  ctp.Navigate("javascript:" + js);
}

function OnQueryRoute() {
  try {
    if (!ctp) { OnShowMap(); }
    var app = wps.Application;
    var sheet = app.ActiveSheet;
    var sel = app.Selection;
    var guard = function(cond, msg) { return cond || (app.StatusBar = msg, false); };
    return guard(sheet.Name === "站存测试文件", "仅支持在'站存测试文件'工作表中查询")
      && guard(sel && sel.Cells.Count === 1, "请选中一个单元格")
      && guard(sel.Column === 8 || sel.Column === 11, "请选中H列或K列的单元格")
      && guard(sel.Value2 && String(sel.Value2).trim(), "单元格不能为空")
      && guard(String(sel.Value2).trim() !== "钦州港", "起点站不能为自身")
      && (sendToCpt(String(sel.Value2).trim()), true);
  } catch (e) {
    console.log("OnQueryRoute 出错: " + e.message);
    return false;
  }
}

window.OnLoad = OnLoad;
window.OnShowMap = OnShowMap;
window.OnQueryRoute = OnQueryRoute;