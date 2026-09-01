//获取发布版本号（接口返回 ["C260702","2026-07-23","版本说明..."]），并触发地图初始化。
//必须在 Token 就绪后调用（带 Authorization 头），由 GetToken 的 success/error 两处触发。
function getVersion()
{
    $.ajax({
        type : "GET",
        url  : "/getReleaseVersionData",
        data : {},
        beforeSend: function (request) {
            request.setRequestHeader("Authorization", window.sessionStorage.token);
        },
        success : function (msg) {
            if (msg && msg.length && msg[0]) {
                map_version = msg[0];
                $(".right_version").text(map_version);
                //同时填充版本说明表格
                var row = document.createElement("tr");
                var td1 = document.createElement("td");
                var td2 = document.createElement("td");
                var td3 = document.createElement("td");
                td3.style.textAlign = "left";
                td1.appendChild(document.createTextNode(msg[0]));
                td2.appendChild(document.createTextNode(msg[1]));
                td3.appendChild(document.createTextNode(msg[2]));
                row.appendChild(td1); row.appendChild(td2); row.appendChild(td3);
                var t = document.getElementById("rows1");
                if (t) t.append(row);
            }
            initMap();   //版本号就绪，触发地图初始化（initMap 内部自带 map_version 非空保护）
        },
        error : function () {
            //后端不可达：兜底写入本地固定版本（本地兜底数据即 C260702），保证地图仍能加载
            map_version = "C260702";
            $(".right_version").text(map_version);
            initMap();
        }
    });
}

//Token 判定完成后，若地图初始化时因 apiAvailable 尚未确定而挂起了加载，则在此触发
function firePendingMapLoad()
{
    if (window.__pendingMapLoad)
    {
        var fn = window.__pendingMapLoad;
        window.__pendingMapLoad = null;
        fn();
    }
}

function GetToken()
{
    //接收token
    var token = "";
    var ip = "10.208.2.72";
    getTokenFromServer();

    function getTokenFromServer( )
    {
        $.ajax
        ({
            type : "GET",
            url : "/getToken",
            dataType : "",
            data : {ip, ip},
            success : function(msg)
            {
                token =  msg;
                window.sessionStorage.token = token;
                apiAvailable = true;                 //后端正常，后续使用 API 请求
                //正常获取到 Token，点亮状态栏 Token 指示灯（绿色）
                var dot = document.getElementById("status_token");
                if (dot) { dot.className = "dot ok"; }
                getVersion();                        //Token 已就绪，获取发布版本号（带真实 token）
                firePendingMapLoad();                //Token 已判定，触发挂起的地图数据加载
            },
            error : function(errorMsg)
            {
                // console.log("getToken()失败");
                token = "3231212";
                window.sessionStorage.token = token;
                apiAvailable = false;                //后端不可达，后续直接使用本地数据/计算
                //后端不可达，使用本地兜底 Token，指示灯置黄色
                var dot = document.getElementById("status_token");
                if (dot) { dot.className = "dot local"; }
                getVersion();                        //用兜底 token 仍尝试获取版本号（失败则兜底 C260702）
                firePendingMapLoad();                //Token 已判定，触发挂起的地图数据加载
            }
        });
    }
}