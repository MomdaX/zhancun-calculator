//点点径路
function Jl( )
{
    var td;                                                 //特定1，最短2
    var pmdm;                                               //7为品名代码
    var jsb;                                                //记事标
    var train_number = 0;                                   //车次
    var ysfs = 1;
    var cz_xx = "C";                                           //车种
    var cx_xl = "C60"
    var fz_dz_array    = [];                                //发到站列表
    var jlout = null;                                       //径路返回结果
    var loading_index;                                      //转圈
    var jl_state = 0;                                       //1:点点径路;2:输入
    var userToken;
    var cz_dic;
    var cx_dic;
    var xx_dic;
    var xl_dic;
    var user_info;

    //点点径路回退
    this.undoJl = undoJl;
    function undoJl()
    {
        if (fz_dz_array.length > 1)
        {
            if (fz_dz_array.length > 2)
            {
                fz_dz_array.pop();
                d3.selectAll(".map_circle").style("fill", "#FFF");
                getJl();
            }
            else
            {
                fz_dz_array.pop();
                setPrompt("点点径路 : " + fz_dz_array[0].zm);

                d3.selectAll(".map_circle").style("fill","#FFF");
                d3.selectAll(".map_line").style("stroke","#FFF");
                d3.selectAll(".map_circle").on("click", null);
                d3.selectAll(".map_circle").on("click", getCircleInfo);
                d3.selectAll(".map_line").on("click", null);
                d3.selectAll(".map_line").on("click", map.getSegmentInfo);
                d3.selectAll(".map_path").remove();

                //回退时绘制发站的颜色
                if (fz_dz_array[0].id.length < 13)
                    d3.select("#" + fz_dz_array[0].id).style("fill", "#FE4B4B");
                else
                    d3.select("#" + fz_dz_array[0].id).style("stroke", "#FE4B4B");
            }
        }
    }

    this.clear = clear;
    function clear()
    {
        d3.selectAll(".map_path").remove();
        d3.selectAll(".mid_station_node").remove();
        d3.selectAll(".mid_station_text").remove();
        //清理上一条径路补画的中间站端点（切换径路时避免残留）
        d3.selectAll(".jl_mid_endpoint, #jl_mid_dz, #jl_mid_fz").remove();
        jl_state           = 0;
        fz_dz_array.length = 0;
        jlout              = null;
    }

    //清除已生成的径路：移除径路线/中间站/补画端点，并恢复圆点/线路路局本色
    this.clearJl = clearJl;
    function clearJl()
    {
        clear();
        map.drawMapColorByLj();
        setPrompt("径路已清除");
    }

    this.clearParameters = clearParameters;
    function clearParameters( )
    {
        if (jl_state = 1)
            setPrompt("点点径路 : ");

        d3.selectAll(".map_circle").style("fill","#FFF");
        d3.selectAll(".map_line").style("stroke","#FFF");
        d3.selectAll(".map_circle").on("click", null);
        d3.selectAll(".map_circle").on("click", getCircleInfo);
        d3.selectAll(".map_line").on("click", null);
        d3.selectAll(".map_line").on("click", map.getSegmentInfo);
        d3.selectAll(".map_path").remove();
        d3.selectAll(".mid_station_node").remove();
        d3.selectAll(".mid_station_text").remove();
        //清理上一条径路补画的中间站端点（切换径路时避免残留）
        d3.selectAll(".jl_mid_endpoint, #jl_mid_dz, #jl_mid_fz").remove();

        fz_dz_array.length    = 0;
        jlout                 = null;
    }

    this.pointPointJl = pointPointJl;
    function pointPointJl()
    {
        map.clear();
        jl_state = 1;
        clearParameters();
        //径路输入不再依赖后端 jlInput 模态框页：发到站通过点击地图圆点/线段完成
        setPrompt("点点径路 : 请点击地图上发站、到站（及途经站）");
    }

    function getCircleInfo()
    {
        $("#fz_input").val("");
        $("#dz_input").val("");
        var id = this.getAttribute("id");

        //Token 不可用（后端不可达）：跳过 /getNode，直接用本地节点数据取 lh/zm，不发 API
        if (apiAvailable === false)
        {
            var d = d3.select(this).datum();
            if (d && d.lh)
                setFzDzArray(d.id, d.lh, d.lh, 1);   //本地节点无站名 zm，用电报码 lh 占位
            else
                map.setMsgInfo("本地节点数据未加载，无法获取该站信息");
            return;
        }

        //涟漪
        const ripple0 = document.createElement("div");
        ripple0.className = "ripple1";
        document.body.appendChild(ripple0);
        ripple0.style.left = `${d3.event.clientX}px`;
        ripple0.style.top = `${d3.event.clientY}px`;
        ripple0.style.animation = `ripple-effect1 .4s  linear`;
        // ripple0.onanimationend = () => {
        //     document.body.removeChild(ripple0);
        // }
        ripple0.addEventListener('animationend', function () {
            document.body.removeChild(ripple0);
        })

        const ripple1 = document.createElement("div");
        ripple1.className = "ripple2";
        document.body.appendChild(ripple1);
        ripple1.style.left = `${d3.event.clientX}px`;
        ripple1.style.top = `${d3.event.clientY}px`;
        ripple1.style.animation = `ripple-effect1 .8s  linear`;
        // ripple1.onanimationend = () => {
        //     document.body.removeChild(ripple1);
        // }
        ripple1.addEventListener('animationend', function () {
            document.body.removeChild(ripple1);
        })

        const ripple2 = document.createElement("div");
        ripple2.className = "ripple3";
        document.body.appendChild(ripple2);
        ripple2.style.left = `${d3.event.clientX}px`;
        ripple2.style.top = `${d3.event.clientY}px`;
        ripple2.style.animation = `ripple-effect1 1.2s  linear`;
        // ripple2.onanimationend = () => {
        //     document.body.removeChild(ripple2);
        // }
        ripple2.addEventListener('animationend', function () {
            document.body.removeChild(ripple2);
        })

        $.ajax
        ({
            type : "GET",
            url : "/getNode",
            dataType : "",
            data : {map_version : map_version,
                    id          : id},
            beforeSend: function(request) {
                request.setRequestHeader("Authorization", window.sessionStorage.token);
            },
            success : function(msg)
            {
                if (!msg)
                {
                    map.setMsgInfo("登录超时，将返回主界面重新登录!");
                    setTimeout(function () {
                        window.location.href = "/cljl";
                    }, 3000);
                }
                else
                {
                    setFzDzArray( id, msg.lh, msg.zm, 1);
                }
            },
            error : function(errorMsg)
            {
                map.setMsgInfo(errorMsg + ' -> getCircleInfo()失败');
            }
        });
    }

    this.setFzDzArray = setFzDzArray;
    function setFzDzArray( id, lh, zm, if_circle )
    {
        if (fz_dz_array.length == 0)
        {
            clearParameters();
            if (jl_state == 1)
            {
                setPrompt("点点径路 : " + zm);
            }

            //渲染
            if (if_circle == 0)
                d3.select("#" + id).style("stroke", "#FE4B4B");
            else
                d3.select("#" + id).style("fill", "#FE4B4B");
        }
        else if (fz_dz_array.length == 1)
        {
            //渲染
            if (if_circle == 0)
                d3.select("#" + id).style("stroke", "#3462FB");
            else
                d3.select("#" + id).style("fill", "#3462FB");
        }
        else if (fz_dz_array.length > 1 && fz_dz_array.length <= 10)                  //多点经由
        {
            d3.selectAll(".map_circle").style("fill", "#FFFFFF");
            d3.select("#" + id).style("fill", "#22CE65");
            // setPrompt("点点径路 : " + fz_dz_array[0].zm + " -> " + fz_dz_array[1].zm);
        }
        else
            map.setMsgInfo('指定经由个数超限');

        var station = {};
        station.id = id;
        station.lh = lh;
        station.zm = zm;
        fz_dz_array.push( station );
        if (fz_dz_array.length > 1)
            getJl();
    }

    //尝试本地兜底生成径路（后端不可用时调用）。
    // 成功：直接渲染并返回 true；失败：设置提示并返回 false。
    // promptPrefix：右下角提示前缀（点点径路 / 搜索径路）。
    function tryLocalJl(fz_dz_array, promptPrefix)
    {
        if (!(window.RailwayPath && window.getMapDndPoint && window.getMapDlnPoint))
            return false;
        try
        {
            var res = window.RailwayPath.makeJlout(fz_dz_array);
            if (res && res.jlout)
            {
                jlout = res.jlout;
                DrawJl();
                if (jl_state == 1) showJlPrompt();
                //本地计算成功：不弹提示框，仅在右下角提示径路信息
                setPrompt("本地径路 : " + fz_dz_array[0].zm + " -> " + fz_dz_array[1].zm);
                return true;
            }
            var why = (res && res.reason) ? res.reason : "未知原因";
            map.setMsgInfo("后端API不可用时，使用本地计算生成径路（未成功）：" + why);
            setPrompt(promptPrefix + fz_dz_array[0].zm + " -> " + fz_dz_array[1].zm + " 计算有误 - " + why);
            return false;
        }
        catch (e)
        {
            map.setMsgInfo("后端API不可用时，使用本地计算生成径路（未成功）：" + (e && e.message || e));
            setPrompt(promptPrefix + fz_dz_array[0].zm + " -> " + fz_dz_array[1].zm + " 计算有误");
            return false;
        }
    }

    //请求径路：只需版本号 + 发到站（及可选经由，由 fz_dz_array 提供），其余参数固定
    function getJl()
    {
        loading_index = layer.load(1);

        //Token 不可用（后端不可达）：跳过 /getJl，直接本地计算生成径路，不白等 API 超时
        if (apiAvailable === false)
        {
            layer.close(loading_index);
            if (tryLocalJl(fz_dz_array, "径路 : "))
                return;
            map.setMsgInfo("后端API不可用时，使用本地计算生成径路（本地计算未成功：本地数据未加载）");
            setPrompt("径路 : " + fz_dz_array[0].zm + " -> " + fz_dz_array[1].zm + " 计算有误");
            clearParameters();
            return;
        }

        var td_id        = 1;          //特定径路
        var pmdm         = 0;          //品名：无
        var jsb          = 0;          //记事标：无
        var ysfs         = 1;          //运输方式：整车
        var cz_xx        = 0;          //车种：固定
        var cx_xl        = 0;          //车型：固定
        var train_number = 0;          //车次：无

        $.ajax
        ({
            type : "GET",
            url : "/getJl",
            dataType : "",
            data : {map_version  : map_version,
                    td_id        : td_id,
                    zmdm         : getFzDzArrayLh(),
                    pmdm         : pmdm,
                    jsb          : jsb,
                    train_number : train_number,
                    ysfs         : ysfs,
                    cz_xx        : cz_xx,
                    cx_xl        : cx_xl
                    },
            beforeSend: function(request) {
                request.setRequestHeader("Authorization", window.sessionStorage.token);
            },
            success : function(msg)
            {
                if (!msg)
                {
                    map.setMsgInfo("登录超时，将返回主界面重新登录!");
                    setTimeout(function () {
                        window.location.href = "/cljl";
                    }, 3000);
                }
                else
                {
                    jlout = msg;
                    DrawJl();

                    if (jl_state == 1)                          //点点径路
                    {
                        showJlPrompt();
                    }
                }
                },
            error : function(errorMsg)
            {
                layer.close(loading_index);
                //后端不可用时，尝试本地兜底生成径路
                if (tryLocalJl(fz_dz_array, "径路 : "))
                    return;
                //本地兜底能力不可用（库/数据未加载）：兜底提示
                map.setMsgInfo("后端API不可用时，使用本地计算生成径路（本地计算未成功：本地数据未加载）");
                setPrompt("径路 : " + fz_dz_array[0].zm + " -> " + fz_dz_array[1].zm + " 计算有误");
                clearParameters();
            }
        });
    }

    //构造右下角径路提示（站名、经由、里程、分界站、折返）——来自原 getJl 逻辑
    this.showJlPrompt = showJlPrompt;
    function showJlPrompt()
    {
        var jyz = "";
        if (fz_dz_array.length > 2)
        {
            jyz = " 经由( ";
            $.each(fz_dz_array, function(i, obj)
            {
                if (i > 1)
                    jyz += obj.zm + " | ";
            });
            jyz = jyz.slice(0, jyz.length - 2);
            jyz += ") ";
        }
        var fjz_list = " 分界站( ";
        if (jlout.fjz.length != 0)
        {
            $.each(jlout.fjz, function(i, obj)
            {
                if (obj.big_or_small == 1)
                    fjz_list += obj.zm + " | ";
            });
            if (fjz_list.charAt(fjz_list.length - 2) != '(')
                fjz_list = fjz_list.slice(0, fjz_list.length - 2);
            else
                fjz_list = fjz_list.slice(0, fjz_list.length - 1);
            fjz_list += " )";
        }
        else
            fjz_list = " 分界站( )";
        var zf_info = " ";
        if (jlout.zf != "")
        {
            zf_info += jlout.zf;
        }
        else
            zf_info = "";
        var fz_zm = (jlout.fz && jlout.fz.zm) ? jlout.fz.zm : fz_dz_array[0].zm;
        var dz_zm = (jlout.dz && jlout.dz.zm) ? jlout.dz.zm : fz_dz_array[1].zm;
        var jl_info1 = "点点径路 : " + fz_zm + " -> " + dz_zm + " " + jyz + jlout.lc[0] + "公里";
        var jl_info2 = fjz_list + zf_info;
        setPrompt(jl_info1 + jl_info2);
    }

    function getFzDzArrayLh()
    {
        var zmdm = "";
        $.each(fz_dz_array, function(i, obj)
        {
            zmdm += obj.lh + ".";
        });

        return zmdm.slice(0, zmdm.length - 1);
    }

    //精准判定中间站是否真正在径路上：getJl 不返回中间站信息，但中间站数据带
    // node1/node2（所属区段两端站 node）+ lc（里程），而 jlout.path 相邻节点带 node。
    // 仅当某中间站的 {node1,node2} 与径路上某相邻 {node_a,node_b} 集合相等，
    // 且其 lc 落在 [fz.lc, dz.lc] 范围内时，才视为“真正在径路上”，
    // 精准避免平行邻近线、距离近、或“终点延长线”（与 dz 同 node 对但 lc 超出 dz.lc）误染。
    // 终点站本身（zm==dz.zm）若属于中间站（不在 path 圆点列表中），在此直接染蓝。
    function colorMidStationsOnPath(pathPts)
    {
        //若未传入 path（由外部在已有径路后调用），使用当前径路 jlout
        if (!pathPts)
        {
            if (!jlout || !jlout.path) return;
            pathPts = jlout.path;
        }
        if (!pathPts || pathPts.length < 2) return;

        var pairSet = {};
        var endNode = null;
        var dzNode2 = (jlout && jlout.dz && jlout.dz.node2) ? jlout.dz.node2 : null;
        //到站判定源：优先用用户实际选择的到站（fz_dz_array 末项，zm/lh/id 均可靠），
        // fallback 到后端 jlout.dz。避免后端 dz.zm 异常导致到站中间站漏染。
        var dzSel = (fz_dz_array && fz_dz_array.length > 1) ? fz_dz_array[fz_dz_array.length - 1] : null;
        var dzZm    = (dzSel && dzSel.zm)    ? dzSel.zm    : ((jlout && jlout.dz && jlout.dz.zm)    ? jlout.dz.zm    : null);
        var dzLh    = (dzSel && dzSel.lh)    ? dzSel.lh    : ((jlout && jlout.dz && jlout.dz.lh)    ? jlout.dz.lh    : null);
        var dzId    = (dzSel && dzSel.id)    ? dzSel.id    : ((jlout && jlout.dz && jlout.dz.id)    ? jlout.dz.id    : null);
        var hitEnd  = false;
        for (var i = 0; i < pathPts.length - 1; i++)
        {
            var a = pathPts[i].node, b = pathPts[i + 1].node;
            if (a == null || b == null || a === 0 || b === 0) continue;  //跳过 node=0 占位项
            pairSet[a + "-" + b] = true;
            pairSet[b + "-" + a] = true;
            endNode = b;
            if (dzNode2 && (b === dzNode2 || a === dzNode2))
            {
                hitEnd = true;
                break;   //到达真正终点，停止后续延长线对
            }
        }
        //若 dz.node2 未出现在 path 节点中，补齐“path 末真实节点 ↔ dz.node2”终点段，
        //使柳城(1775)→全村(1799) 段上的中间站（宜州/叶茂）能命中变黄。
        if (!hitEnd && dzNode2 && endNode && endNode !== dzNode2)
        {
            pairSet[endNode + "-" + dzNode2] = true;
            pairSet[dzNode2 + "-" + endNode] = true;
        }
        if (endNode == null) return;

        //获取 fz.lc / dz.lc（用于里程范围过滤）：从 window.中间站数据 中按 zm + node 匹配。
        //fz 不在中间站数据时（如港站），按 0 处理（钦州港等起点的 lc 起点为 0）。
        var midData = window["中间站数据"];
        function lookupLc(zm, n1, n2) {
            if (!midData || !zm) return null;
            for (var lj in midData) {
                var arr = midData[lj];
                if (!arr) continue;
                for (var k = 0; k < arr.length; k++) {
                    var m = arr[k];
                    if (m && m.zm === zm && (m.node1 === n1 || m.node2 === n1) && (m.node1 === n2 || m.node2 === n2))
                        return m.lc;
                }
            }
            return null;
        }
        var dzLc = lookupLc((dzZm || (jlout.dz && jlout.dz.zm)), (dzSel && dzSel.node1) || (jlout.dz && jlout.dz.node1), (dzSel && dzSel.node2) || (jlout.dz && jlout.dz.node2));
        if (dzLc == null) dzLc = Number.MAX_VALUE;  //查不到则不限制上限
        var fzLc = lookupLc(jlout.fz.zm, jlout.fz.node1, jlout.fz.node2);
        if (fzLc == null) fzLc = 0;                //fz 不在中间站数据时默认 0

        d3.selectAll(".mid_station_node").each(function(d){
            if (!d || d.node1 == null || d.node2 == null) return;
            var key1 = d.node1 + "-" + d.node2;
            var key2 = d.node2 + "-" + d.node1;
            if (!(pairSet[key1] || pairSet[key2])) return;
            //终点站本身（与 dz 的站名/电报码/id 任一匹配）：它不是 path 圆点，需在此直接染蓝
            if ((dzZm && d.zm === dzZm) || (dzLh && d.lh === dzLh) || (dzId && d.id === dzId))
            {
                d3.select(this).style("fill", "#3462FB");
                return;
            }
            //里程范围过滤：
            // - 当 dz 不在 path 末区段（靠补齐终点段连通）时，用 dzLc 上限剔除“终点延长线”上多出的站；
            // - 当 dz 本身就在 path 末区段内（hitEnd，如广州北→飞来峡同属 1577-2825 区段），
            //   该区段内所有中间站都是真实途经站，不得用 dzLc 上限误杀（否则潖江口/源潭等 lc>dzLc 的站会被漏染）。
            if (d.lc != null && d.lc < fzLc) return;
            if (!hitEnd && d.lc != null && d.lc > dzLc) return;
            //命中且里程在 [fz.lc, 上限] 内 -> 黄色（途经站）
            d3.select(this).style("fill", "#FFFF00");
        });
    }

    function DrawJl( )
    {
        d3.selectAll(".map_path").remove();
        //清理上一条径路补画的中间站端点（切换径路时避免残留），用专属 class+id 双保险
        d3.selectAll(".jl_mid_endpoint, #jl_mid_dz, #jl_mid_fz").remove();
        var lineFunction = d3.line()
                             .x(function(d){ return d.x; })
                             .y(function(d){ return d.y; });

        map.getGroup().insert("path", ".map_circle")
                      .attr("class",         "map_path")
                      .attr("d",             lineFunction(jlout.path))
                      .style("stroke-width", map.getCurrentPathStroke())
                      .on("click",           clickJlPath)
                      .on("mouseover",       function(){ d3.select(this).style("stroke-width", map.getCurrentPathStroke() * 1.5); })
                      .on("mouseout",        function(){ d3.select(this).style("stroke-width", map.getCurrentPathStroke()); });

        //先中断上一次径路可能仍在排队的过渡动画，避免恢复颜色后被旧 transition 重新覆盖
        d3.selectAll(".map_circle").interrupt();
        //按路局恢复所有圆点/线路颜色，清掉上一次径路残留的上色，再重绘本次径路
        map.drawMapColorByLj();

        //先把已绘制的中间站还原为路局本色，避免上一条径路的着色残留
        d3.selectAll(".mid_station_node").style("fill", function(d){
            return lj_color_list[parseInt(d.fq/100)];
        });

        //中间站随径路上色：依据站名判定哪些中间站真正在径路上（getJl 不返回中间站信息，用站名匹配）
        colorMidStationsOnPath(jlout.path);

        //圆点随径路上色（起点红、途经黄、终点蓝）
        for (var i in jlout.path)
        {
            var node_id = jlout.path[i].id;
            var color;
            if (i == 0)
                color = "#FE4B4B";                                 //起点：红
            else if (i == jlout.path.length - 1)
                color = "#3462FB";                                 //终点：蓝
            else
                color = "#FFFF00";                                 //途经：黄

            if (i != 0 && i != jlout.path.length - 1)
                d3.select("#" + node_id).transition().delay(1000).style("fill", color);
            else
                d3.select("#" + node_id).style("fill", color);
        }

        //经由
        var len = fz_dz_array.length;
        if (len > 2)
        {
            for(var j = len - 1; j >= 2; j--)
            {
                d3.select("#" + fz_dz_array[j].id).transition().delay(1000).style("fill", "#22CE65");
            }
        }

        //分界站
        if (jlout !== null && jlout.fjz !== null)
        {
            for (var j = 0; j < jlout.fjz.length; j++)
            {
                if (jlout.fjz[j].big_or_small === 1)
                    d3.select("#" + jlout.fjz[j].id).transition().delay(1000).style("fill", "#FF9900");
            }
        }

        //每次生成径路都自动缩放/平移，使整条径路完整显示在当前视口中
        var minx = d3.min(jlout.path, function(d){return d.x;});
        var miny = d3.min(jlout.path, function(d){return d.y;});
        var maxx = d3.max(jlout.path, function(d){return d.x;});
        var maxy = d3.max(jlout.path, function(d){return d.y;});
        setJlCenter( minx, miny, maxx, maxy );

        //终点/起点为中间站且当前地图上尚未显示该站（如未先点线段加载、且 dz 的 node2 不在 path 节点内）时，
        //单独补画该中间站圆点 + 站名（大站已在地图上显示，无需处理）。
        //优先用用户实际选择的到站/发站（fz_dz_array 末项/首项，zm 可靠），fallback 到后端 jlout.dz/fz。
        var dzSelEp = (fz_dz_array && fz_dz_array.length > 1) ? fz_dz_array[fz_dz_array.length - 1] : jlout.dz;
        var fzSelEp = (fz_dz_array && fz_dz_array.length > 0) ? fz_dz_array[0] : jlout.fz;
        drawMidEndpointIfNeeded(dzSelEp, "#3462FB");
        drawMidEndpointIfNeeded(fzSelEp, "#FE4B4B");

        layer.close(loading_index);
    }

    //若 station 是中间站且地图上尚未显示同名站（大站 map_text 或 已加载中间站 mid_station_text 均无该站名），
    //则从 window.中间站数据 按 zm 命中完整对象，委托 map.drawMidStationOne 补画一个圆点 + 站名
    //（绘制方式/样式与原中间站绘制函数完全一致：class=mid_station_node/mid_station_text，仅补画用专属 class 以免被中间站切换误删）。
    //大站已在地图上显示 -> 同名站已存在 -> 直接跳过，不做任何处理。
    function drawMidEndpointIfNeeded(station, color)
    {
        if (!station || !station.zm) return;

        //专属 id：先判断是否已补画同色端点，避免重复补画同一站点
        var eid = (color === "#3462FB") ? "jl_mid_dz" : "jl_mid_fz";
        if (d3.select("#" + eid).size() > 0) return;

        //按 id 判定：地图上已存在该 id 的圆点（大站或已加载中间站）-> 不补画
        if (station.id && d3.select("#" + station.id).size() > 0) return;

        //兜底按站名(绑定数据 zm)判定：已显示同名站 -> 不补画
        var exist = false;
        d3.selectAll(".map_circle, .mid_station_node").each(function(d){
            if (d && d.zm === station.zm) exist = true;
        });
        if (exist) return;

        //从中间站数据中查找该站完整对象（含 x/y/angle/fq 等）
        var m = null;
        var mid = window["中间站数据"];
        if (mid)
        {
            for (var lj in mid)
            {
                var arr = mid[lj];
                if (!arr) continue;
                for (var k = 0; k < arr.length; k++)
                {
                    if (arr[k] && arr[k].zm === station.zm) { m = arr[k]; break; }
                }
                if (m) break;
            }
        }
        if (!m) return;

        //专属 class：不被“关闭/切换中间站显示”的 selectAll(".mid_station_*").remove() 误删
        map.drawMidStationOne(m, color, eid, "jl_mid_endpoint");
    }

    //点击已生成径路：不再依赖后端 getJloutputPage 片段，直接使用内置 #jloutputModal 渲染
    function clickJlPath()
    {
        if (!jlout) return;
        $("#jloutputModal").modal("show");
        $("#jloutputTable").empty();
        $("#jloutputTable").append($(getJloutMessage()));
    }
    function getJloutMessage()
    {
        var jl_out_message;
        jl_out_message   = "<tr><td style='font-weight:bold;'>发站</td><td width='100px' >" + jlout.fz.zm + "</td><td style='text-align:center;'>" + jlout.fz.lh + "</td><td style='text-align:center;'>" + jlout.fz.ljjc + "局</td></tr>";
        jl_out_message += "<tr><td style='font-weight:bold;'>到站</td><td width='100px' >" + jlout.dz.zm + "</td><td style='text-align:center;'>" + jlout.dz.lh + "</td><td style='text-align:center;'>" + jlout.dz.ljjc + "局</td></tr>";

        jl_out_message += "<tr><td style='font-weight:bold;'>全程里程</td><td style='text-align:center;'>" + jlout.lc[0] + "公里</td><td style='text-align:center;font-weight:bold;'>备注</td><td style='text-align:center;color: #FE4B4B'>" + jlout.zf + "</td></tr>";

        var index;
        jl_out_message += "<tr><td style='text-align:center;font-weight:bold;' colspan='4'>路局里程</td></tr>";
        $.each(jlout.lj, function(i, obj)
        {
            index = i+1;
            jl_out_message += "<tr><td style='text-align:center;'>" + index + "</td><td style='text-align:center;'>" + obj.jc + "局</td><td style='text-align:right;' colspan='2'>" + obj.lc + "公里</td></tr>";
        });

        if (jlout.fjz.length != 0)
        {
            index = 0;
            jl_out_message += "<tr><td style='text-align:center;font-weight:bold;' colspan='4'>分界站</td></tr>";
            $.each(jlout.fjz, function(i, obj)
            {
                index = i+1;
                var c_or_r;
                if (obj.pass == 1)
                    c_or_r = "出";
                else
                    c_or_r = "入";

                var b_or_s;
                if (obj.big_or_small == 1) {
                    b_or_s = "大口";
                    jl_out_message += "<tr><td style='text-align:center;'>" + index + "</td><td style='text-align:left;'>" + obj.zm + "</td><td style='text-align:center;'>" + b_or_s + "</td>><td style='text-align:center;'>" + c_or_r + "</td></tr>";
                }
                });
        }
        return jl_out_message;
    }

    //径路居中
    this.setJlCenter = setJlCenter;
    //暴露中间站上色方法，供地图在生成径路后再加载中间站时复用（按径路规则变色）
    this.colorMidStationsOnPath = colorMidStationsOnPath;
    function setJlCenter( minx, miny, maxx, maxy )
    {
        var px = (minx + maxx)/2;
        var py = (miny + maxy)/2;
        //构造合法的 ZoomTransform（d3.zoomIdentity 是不可变对象，必须用 translate/scale 链式构造）
        var k = Math.min(map_width*0.9/(maxx - minx), map_height*0.9/(maxy - miny));
        if (k > 28) k = 28;                 //缩放最大比列
        if (k < 1)  k = 1;
        var transform = d3.zoomIdentity.translate(-px*k, -py*k).scale(k);

        //定位（map.transformDom 内部用 rAF 缓动做平滑飞入）
        map.transformDom( transform );
    }

    this.getJlState = getJlState;
    function getJlState()
    {
        return jl_state;
    }

    this.getCz = getCz;
    function getCz() {
        return cz_dic;
    }

    this.getCx = getCx;
    function getCx() {
        return cx_dic;
    }

    this.getXx = getXx;
    function getXx() {
        return xx_dic;
    }

    this.getXl = getXl;
    function getXl() {
        return xl_dic;
    }

    function insertStr(str, index, insertStr) {
        const ary = str.split('');
        ary.splice(index, 0, insertStr);
        return ary.join('');
    }

    function setPrompt( prompt ) {
        var jy_begin_index = prompt.indexOf("经由(", 0);
        if (jy_begin_index !== -1) {
            prompt = insertStr(prompt, jy_begin_index + 3, "<span class=\"jyzText\">");
            var jy_end_index = prompt.indexOf(")", 0);
            prompt = insertStr(prompt, jy_end_index, "</span>");
        }

        var fjz_begin_index = prompt.lastIndexOf("分界站(");
        if (fjz_begin_index !== -1) {
            prompt = insertStr(prompt, fjz_begin_index + 4, "<span class=\"fjzText\">");
            var fjz_end_index = prompt.lastIndexOf(")");
            prompt = insertStr(prompt, fjz_end_index, "</span>");
        }

        var zf_begin_index = prompt.lastIndexOf("有折返");
        if (zf_begin_index !== -1) {
            prompt = insertStr(prompt, zf_begin_index, "<span class=\"zfText\">");
            var zf_end_index = prompt.lastIndexOf("返");
            prompt = insertStr(prompt, zf_end_index + 1, "</span>");
            alert("请注意：\n" +
                "当前径路存在折返，您所选经由站顺序必须和径路顺序一致！\n" +
                "（可同时按住Ctrl+Z，使用回退功能删除上一个选择的车站。）")
        }
        $(".right_svg_prompt").html(prompt +
            "<style>" +
            ".jyzText {" +
            "color: #22CE65" +
            "}" +
            ".fjzText {" +
            "color: #FF9900" +
            "}" +
            ".zfText {" +
            "color: #FF0033" +
            "}" +
            "</style>");
    }

    //从本地中间站数据中按 zm / lh / id 查找完整中间站对象
    function findMidStation(row) {
        var mid = window["中间站数据"];
        if (!mid || !row) return null;
        for (var lj in mid) {
            var arr = mid[lj];
            if (!arr) continue;
            for (var k = 0; k < arr.length; k++) {
                var m = arr[k];
                if (!m) continue;
                if ((row.id && m.id === row.id) ||
                    (row.lh && m.lh === row.lh) ||
                    (row.zm && m.zm === row.zm))
                    return m;
            }
        }
        return null;
    }

    //钦州港 -> 选中站 径路生成（搜索框下拉选站后调用）
    // row: {zm, lh, id, ljjc, ...} 来自 /findStation 结果
    this.generateJlByStation = generateJlByStation;
    function generateJlByStation(row)
    {
        //只清除旧径路，保留已绘制的中间站（点击线段绘制的叠加层不清理）
        d3.selectAll(".map_path").remove();
        //清理上一条径路补画的中间站端点（本次径路重新生成前避免残留）
        d3.selectAll(".jl_mid_endpoint, #jl_mid_dz, #jl_mid_fz").remove();
        jl_state = 1;

        //若到站是中间站（不在大站节点数据中），先在地图上单站绘制它，
        //否则本地算法无法在图中定位该站、也就无法计算径路。
        var midStation = findMidStation(row);
        if (midStation) {
            map.drawMidStationOne(midStation, "#FE4B4B", row.id, "jl_mid_endpoint");
        }

        //填充发到站数组，供 DrawJl 上色起终点
        fz_dz_array.length = 0;
        fz_dz_array.push({id : "QVZ", lh : "QVZ", zm : "钦州港"});
        fz_dz_array.push({id : row.id, lh : row.lh, zm : row.zm});

        //复用 getJl() 统一请求（含 Token 不可用直接本地、error 本地兜底）
        getJl();
    }
}
