function Map(height, width)
{
    var map_nodes           = null;
    var zoom                = null;
    var group               = null;
    var transform           = {};
    var loading_index;                          //转圈
    var scale;                                  //比例
    var circle_r            = 2.0;              //初始圆半径
    var line_stroke_width   = 0.8;              //初始线宽
    var path_stroke_width   = 1.0;              //初始PATH线宽
    var font_size           = 1.2;              //初始字体大小
    var mid_font_size       = 0.9;              //中间站初始字体大小

    var current_circle_r;                       //当前圆半径
    var current_mid_circle_r;                   //当前中间站圆半径
    var current_line_stroke_width;              //当前线宽
    var current_path_stroke_width;              //当前线宽

    var transform_k_svg = 1;                    //svg当前缩放比例scale
    var mousedown_flag = false;                 //用于拖拽svg时鼠标手势的变化,检测鼠标是否为按下状态

    var current_ljh = 0;
    this.clear = clear;
    function clear()
    {
        jl.clear();
        d3.selectAll(".mid_station_node").remove();
        d3.selectAll(".mid_station_text").remove();
    }

    this.createMap = createMap;
    function createMap(e)
    {
        setPrompt("路局显示->");
        loading_index = layer.load(1);
        var svg = d3.select("svg")
                    .attr("width",  width)
                    .attr("height", height)
                    .attr("viewBox",  + "" + (-map_width/2) + " " + (-map_height/2) + " " + map_width + " " + map_height)

        //鼠标按下鼠标手势变为grabbing
        svg.on("mousedown", function () {
            mousedown_flag = true;
            d3.select('.svg_map').style('cursor','grabbing');
        })
        svg.on("mousemove", function () {
            if (mousedown_flag){
                d3.select('.svg_map').style('cursor','grab');
                mousedown_flag = false;
            }
        })
        svg.on("mouseup", function () {
            d3.select('.svg_map').style('cursor','grab');
        })

        scale = Math.sqrt(map_width * map_width + map_height * map_height) / Math.sqrt(width * width + height * height);
        zoom = d3.zoom()
                 .scaleExtent([1,100])							        //用于设置最小和最大的缩放比例
                 .on("zoom", zoomed)

        d3.select(".svg_map").call(zoom)
                        .on("dblclick.zoom", null)							//取消双击ZOOM
                        .on("dblclick", svgDbclick) 					    //改为图形还原

        //创建GROUP
        group = svg.append("g");
        group.attr("class", "mapGroup");

        //加载地图数据：根据 Token 可用性决定走 API 还是本地（第一次加载即确定，不轮询、不白等 API 超时）
        loadMapData();
    }

    //根据 apiAvailable 决定地图数据来源：undefined→挂起等 Token 判定；true→API；false→本地
    function loadMapData()
    {
        if (apiAvailable === undefined)
        {
            window.__pendingMapLoad = loadMapData;   //Token 尚未判定，挂起，待 getToken 回调触发
            return;
        }
        if (!apiAvailable)
        {
            loadMapLocal();                          //后端不可达：直接用本地兜底数据，不发 API
            return;
        }
        //后端正常：走 API 请求
        getMapFontPoint();
        getMapLinePoint();
        getMapNodePoint();
    }

    //本地兜底绘制站名/线路/节点（等价于三个 API 的 success 分支，不发请求）
    //绘制顺序须与 API 模式一致：font → line → node（node 最后画，圆点在线路之上），否则本地模式图层会错乱
    function loadMapLocal()
    {
        if (window.getMapDftPoint) { drawMapFont(window.getMapDftPoint); setDataStatus("font", "local"); }
        if (window.getMapDndPoint) { map_nodes = window.getMapDndPoint; }   //先备好节点数据，供 drawMapLine 内部同步补画节点
        if (window.getMapDlnPoint) { drawMapLine(window.getMapDlnPoint); setDataStatus("line", "local"); }
        if (window.getMapDndPoint) { setDataStatus("node", "local"); }     //节点由 drawMapLine 在 map_nodes 就绪时绘制
    }

    //设置接口状态指示灯：type 取值 font/line/node；status 取值 ok/api正常 local/本地兜底
    function setDataStatus(type, status)
    {
        var idMap = { font : "status_font", line : "status_line", node : "status_node" };
        var el = document.getElementById(idMap[type]);
        if (!el) return;
        //用 classList 增量管理，避免覆盖元素上已有的其它 class
        el.classList.remove("ok", "local");
        el.classList.add("dot");
        if (status === "ok")        el.classList.add("ok");
        else if (status === "local") el.classList.add("local");
    }

    //获得站名的坐标
    function getMapFontPoint( )
    {
        $.ajax
        ({
            type : "GET",
            url : "/getMapDftPoint",
            dataType : "",
            data : {map_version : map_version},
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
                    drawMapFont( msg );
                    setDataStatus("font", "ok");
                }

            },
            error : function(errorMsg)
            {
                layer.close(loading_index);
                setDataStatus("font", "local");
                //后端不可用时，从本地 JS 兜底加载站名数据（window.getMapDftPoint）
                if (window.getMapDftPoint)
                    drawMapFont(window.getMapDftPoint);
            }
        });
    }
    //站名铺画
    function drawMapFont( fonts )
    {
        var text = d3.select("g").selectAll("text")
            .data(fonts)
            .enter()
            .append("text")
            .attr("x", function(d) { return d.x; })
            .attr("y", function(d) { return d.y; })
            .text(function(d) { return d.zm;})
            .attr("text-anchor", (d,i) => {
                if (d.angle == 90)
                    return "middle";
                else
                    return "middle";
            })
            .attr('transform', function(d, i) { return 'rotate(' + (d.angle) + ' ' + (d.x) + ',' + (d.y) + ')' })
            .attr('rotate', (d,i) => {
                if (d.angle == 90)
                    return -90;
                else
                    return 0;
            })
            .attr("id", function(d) { return d.id; })
            .attr("class", "map_text")
            .style("font-size", font_size * scale)
            .attr("dx", (d,i) => {
                if (d.angle == 90)
                    return "3";
                else
                    return 1.75;
            })
            .attr("dy", (d,i) => {
                if (d.angle == 90)
                    return "1.25";
                else
                    return 1.25;
            })
            .on("mouseover", function(){ d3.select(this).style("font-size", font_size * scale * 1.2); })
            .on("mouseout",  function(){ d3.select(this).style("font-size", font_size * scale); });
    }
    //获得线段的坐标
    function getMapLinePoint( )
    {
        $.ajax
        ({
            type : "GET",
            url : "/getMapDlnPoint",
            dataType : "",
            data : {map_version : map_version},
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
                    drawMapLine( msg );
                    setDataStatus("line", "ok");
                }

            },
            error : function(errorMsg)
            {
                layer.close(loading_index);
                setDataStatus("line", "local");
                //后端不可用时，从本地 JS 兜底加载线路数据（window.getMapDlnPoint）
                if (window.getMapDlnPoint)
                    drawMapLine(window.getMapDlnPoint);
            }
        });
    }

    function drawMapLine( lines )
    {
        //线段铺画
        current_line_stroke_width = line_stroke_width * scale;
        var line = d3.select("g").selectAll("line")
            .data(lines)
            .enter()
            .append("line")
            .attr("x1", function(d) { return d.x1; })
            .attr("y1", function(d) { return d.y1; })
            .attr("x2", function(d) { return d.x2; })
            .attr("y2", function(d) { return d.y2; })
            .attr("id", function(d) { return d.id; })
            .attr("lj", function(d){ return parseInt(d.fq/100); })
            .style("stroke", function(d){ return lj_color_list[parseInt(d.fq/100)]; })
            .style("stroke-width", function(){ return current_line_stroke_width; })
            .attr("class", "map_line")
            //点击线段：按路局编号发起中间站请求，后端不可用时由 getMidStations 用 window.中间站数据 本地兜底绘制
            .on("click", function(){ getMidStations(parseInt(this.getAttribute("lj"))); })
            .on("mouseover", function(){ d3.select(this).style("stroke-width", current_line_stroke_width * 1.5); })
            .on("mouseout",  function(){ d3.select(this).style("stroke-width", current_line_stroke_width); });

        //铺画 NODE：等待节点数据就绪（与 line 接口并发返回，顺序不确定）
        if (map_nodes && map_nodes.length)
        {
            drawMapNode( map_nodes );
        }
        else
        {
            var _waitNode = setInterval(function () {
                if (map_nodes && map_nodes.length)
                {
                    clearInterval(_waitNode);
                    drawMapNode( map_nodes );
                }
            }, 50);
        }

        //获得线路车站信息
        //GetMapLineStation( group );
    }
    //获得节点坐标
    function getMapNodePoint( )
    {
        $.ajax
        ({
            type : "GET",
            url : "/getMapDndPoint",
            dataType : "",
            data : {map_version : map_version},
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
                    map_nodes =  msg;
                    setDataStatus("node", "ok");
                }

            },
            error : function(errorMsg)
            {
                layer.close(loading_index);
                setDataStatus("node", "local");
                //后端不可用时，从本地 JS 兜底加载节点数据（window.getMapDndPoint），并供 drawMapLine 的等待逻辑消费
                if (window.getMapDndPoint)
                {
                    map_nodes = window.getMapDndPoint;
                    drawMapNode(window.getMapDndPoint);
                }
                else
                {
                    map_nodes = null;
                }
            }
        });
    }
    function drawMapNode( nodes )
    {
        //圆铺画
        current_circle_r = circle_r * scale;
        var circle = d3.select("g").selectAll("circle")
            .data(nodes)
            .enter()
            .append("circle")
            .attr("cx", function(d) { return d.x; })
            .attr("cy", function(d) { return d.y; })
            .attr("id", function(d) { return d.id; })
            .attr("lj", function(d){ return parseInt(d.fq/100); })
            .style("fill", function(d){ return lj_color_list[parseInt(d.fq/100)]; })
            .attr("class", "map_circle")
            .style("r", current_circle_r)
            .on("click", getNodeInfo)
            .on("mouseover", function(){ d3.select(this).style("r", current_circle_r * 1.3); })
            .on("mouseout",  function(){ d3.select(this).style("r", current_circle_r); });

        layer.close(loading_index);
    }

    //缩放
    this.zoomed = zoomed;
    function zoomed()
    {
        d3.selectAll("g").attr("transform",	d3.event.transform);
        transform.x = d3.event.transform.x;
        transform.y = d3.event.transform.y;
        transform.k = d3.event.transform.k;

        transform_x_svg = transform.x;
        transform_y_svg = transform.y;
        transform_k_svg = transform.k;

        //按缩放级别统一计算系数（原 4 段重复分支合并为单一系数，避免与 transformDom 的职责重复）
        //圆点系数 rF：k<=5→1.0，5<k<10→0.8，10<=k<20→0.65，k>=20→0.4（与备份一致）
        //中间站半径恒为圆点半径的 0.65 倍；线宽/径路线宽系数 lwF：1.0 / 0.8 / 0.7 / 0.6
        var k = transform.k;
        var rF, lwF;
        if (k <= 5)            { rF = 1.00; lwF = 1.00; }
        else if (k < 10)       { rF = 0.80; lwF = 0.80; }
        else if (k < 20)       { rF = 0.65; lwF = 0.70; }
        else                   { rF = 0.40; lwF = 0.60; }

        current_circle_r     = circle_r * scale * rF;
        current_mid_circle_r = circle_r * scale * rF * 0.65;
        current_line_stroke_width = line_stroke_width * scale * lwF;
        current_path_stroke_width = path_stroke_width * scale * lwF;

        //动态变化圆点大小（大站 + 补画中间站端点一并同步）
        d3.selectAll(".map_circle").style("r", current_circle_r );
        d3.selectAll(".mid_station_node, .jl_mid_endpoint").style("r", current_mid_circle_r );

        //动态变换线宽度
        d3.selectAll(".map_line").style("stroke-width", current_line_stroke_width);
        d3.selectAll(".map_path").style("stroke-width", current_path_stroke_width);
    }
    //还原图形
    function svgDbclick()
    {
        //双击还原也走 transformDom，与径路飞入一致的 rAF 缓动
        var transform = d3.zoomIdentity.translate(0, 0).scale(1);
        transformDom( transform );
    }

    //坐标转换
    function corrdinateTransX(x)
    {
        return ( width/2 ) - ( ( map_width * height / map_height ) / 2) + (x * height / map_height );
    }

    //坐标转换
    function corrdinateTransY(y)
    {
        return y * height / map_height;
    }

    //获得点击距离远近
    this.getClickDistence = getClickDistence;
    function getClickDistence( dom )
    {
        var x1 = parseFloat( dom.getAttribute("x1") );
        var y1 = parseFloat( dom.getAttribute("y1") );
        var x2 = parseFloat( dom.getAttribute("x2") );
        var y2 = parseFloat( dom.getAttribute("y2") );
        var s_point = d3.mouse( dom );
        var x = parseFloat( s_point[0] );
        var y = parseFloat( s_point[1] );
        var d1 = Math.sqrt((x1 - x)*(x1 - x) + (y1 - y)*(y1 - y));
        var d2 = Math.sqrt((x2 - x)*(x2 - x) + (y2 - y)*(y2 - y));

        if (d1 > d2)
            return 2;
        else
            return 1;
    }

    this.getSegmentInfo = getSegmentInfo;
    function getSegmentInfo( )
    {
        //点击线段：按路局编号发起中间站请求，后端不可用时由 getMidStations 用 window.中间站数据 本地兜底绘制
        //不再加载 /getSegmentPage 模态框页
        var lj = parseInt(this.getAttribute("lj"));
        getMidStations(lj);
    }

    this.getNodeInfo = getNodeInfo;
    function getNodeInfo()
    {
        //直接从圆点绑定的节点数据取 id 与 lh，无需再请求 /getNode
        var d = this.__data__;
        if (!d || !d.lh)
            return;

        //点击车站即生成 钦州港(QVZ) -> 该站 径路，zm 兜底用 lh（提示站名由后端 dz.zm 提供）
        jl.generateJlByStation({id: d.id, lh: d.lh, zm: d.zm || d.lh});
    }

    //通过路局属性渲染节点线段
    this.drawMapColorByLj = drawMapColorByLj;
    function drawMapColorByLj()
    {
        d3.selectAll(".map_circle").each(function( )
        {
            var lj = d3.select(this).attr("lj");
            d3.select(this).style("fill", function(){ return lj_color_list[lj]; });
        });

        d3.selectAll(".map_line").each(function( )
        {
            var lj = d3.select(this).attr("lj");
            d3.select(this).style("stroke", function(){ return lj_color_list[lj]; });
        });
    }

    //获得数据版本
    this.getMapVersion = getMapVersion;
    function getMapVersion()
    {
        return map_version;
    }

    //获得ZOOM
    this.getMapZoom = getMapZoom;
    function getMapZoom()
    {
        return zoom;
    }

    //获得Group
    this.getGroup = getGroup;
    function getGroup()
    {
        return group;
    }

    //获得当前节点的r
    this.getCurrentCircleR = getCurrentCircleR;
    function getCurrentCircleR()
    {
        return current_circle_r;
    }

    //获得当前线路的stroke
    this.getCurrentLineStroke = getCurrentLineStroke;
    function getCurrentLineStroke()
    {
        return current_line_stroke_width;
    }

    //获得当前径路的stroke
    this.getCurrentPathStroke = getCurrentPathStroke;
    function getCurrentPathStroke()
    {
        return current_path_stroke_width;
    }

    this.transformDom = transformDom;
    var _jlTransformTimer = null;          //当前飞入动画句柄，避免快速连点冲突
    function transformDom( t )
    {
        //方案 B：手动 rAF 缓动飞入（缩放+平移同步插值），不依赖 d3.zoom 的 transition（实测易失效）
        var node = d3.select(".svg_map").node();
        if (!node) return;
        var start = node.__zoom ? node.__zoom : d3.zoomIdentity;
        var sx = start.x, sy = start.y, sk = start.k;
        var ex = t.x,    ey = t.y,    ek = t.k;
        var dur = 800;
        var t0 = null;

        //缓动函数：cubic ease-in-out
        function ease(p){ return p < 0.5 ? 4*p*p*p : 1 - Math.pow(-2*p + 2, 3)/2; }

        if (_jlTransformTimer) cancelAnimationFrame(_jlTransformTimer);
        //计算起止视口中心对应的地图坐标，避免缩放与平移节奏不一致导致的“镜头晃动”
        var mc0x = (-sx) / sk, mc0y = (-sy) / sk;
        var mc1x = (-ex) / ek, mc1y = (-ey) / ek;
        function step(ts)
        {
            if (t0 === null) t0 = ts;
            var p = Math.min(1, (ts - t0) / dur);
            var e = ease(p);
            //缩放在对数空间插值（人眼对缩放的感知是对数的），过渡更自然均匀
            var lk0 = Math.log(sk <= 0 ? 1e-6 : sk);
            var lk1 = Math.log(ek <= 0 ? 1e-6 : ek);
            var ck = Math.exp(lk0 + (lk1 - lk0) * e);
            //视口中心（地图坐标）线性插值，再换算回平移量，使缩放与平移同步、镜头不晃
            var mcx = mc0x + (mc1x - mc0x) * e;
            var mcy = mc0y + (mc1y - mc0y) * e;
            var cx = -mcx * ck;
            var cy = -mcy * ck;
            var interp = d3.zoomIdentity.translate(cx, cy).scale(ck);
            d3.select(".svg_map").call( zoom.transform, interp );
            if (p < 1) _jlTransformTimer = requestAnimationFrame(step);
            else _jlTransformTimer = null;
        }
        _jlTransformTimer = requestAnimationFrame(step);
    }

    //信息提示
    this.setMsgInfo = setMsgInfo;
    function setMsgInfo(msg)
    {
        layer.msg(msg, {
            time: 2000 //20s后自动关闭
        });
    }

    function setPrompt( prompt ) {
        $(".right_svg_prompt").html(prompt);
    }

    //关闭中间站显示：清除已绘制的该局中间站节点/站名，并复位当前局标记
    this.hideMidStation = hideMidStation;
    function hideMidStation() {
        d3.selectAll(".mid_station_node").remove();
        d3.selectAll(".mid_station_text").remove();
        current_ljh = 0;
        setPrompt("中间站显示已关闭");
    }
    //获得对应路局中间站信息（暴露给外部 jl 等直接调用，不再依赖中间站模态框）
    this.getMidStations = getMidStations;
    function getMidStations( lj )
    {
        //同一局再次点击：切换关闭
        if (current_ljh == lj)
        {
            d3.selectAll(".mid_station_node").remove();
            d3.selectAll(".mid_station_text").remove();
            current_ljh = 0;
            return;
        }
        //Token 不可用（后端不可达）：直接用本地中间站数据绘制，不发 /getMidStationByLj 请求
        if (apiAvailable === false)
        {
            if (window["中间站数据"] && window["中间站数据"][lj])
            {
                current_ljh = lj;
                drawMidStationFont(window["中间站数据"][lj]);
                drawMidStationNode(window["中间站数据"][lj]);
            }
            else
            {
                setMsgInfo("未找到该局(" + lj + ")的中间站数据");
            }
            return;
        }
        //Token 可用：走后端 API 获取中间站
        $.ajax
        ({
            type : "GET",
            url : "/getMidStationByLj",
            dataType : "",
            data : {map_version : map_version, lj : lj},
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
                    current_ljh = lj;
                    drawMidStationFont(msg);
                    drawMidStationNode(msg);
                }
            },
            error : function(errorMsg)
            {
                layer.close(loading_index);
                //后端不可用时，使用本地中间站数据兜底绘制
                if (window["中间站数据"] && window["中间站数据"][lj])
                {
                    current_ljh = lj;
                    drawMidStationFont(window["中间站数据"][lj]);
                    drawMidStationNode(window["中间站数据"][lj]);
                }
                else
                {
                    setMsgInfo(errorMsg + ' -> getMidStations()失败');
                }
            }
        });
    }

    //中间站站名铺画
    function drawMidStationFont( fonts )
    {
        d3.select("g").selectAll(".mid_station_text").remove();

        var text = d3.select("g").selectAll(".mid_station_text")
            .data(fonts)
            .enter()
            .append("text")
            .attr("x", function(d) { return d.x; })
            .attr("y", function(d) { return d.y; })
            .text( function(d) { return d.zm; })
            .attr("id", function(d) { return d.id; })
            .attr("class", "mid_station_text")
            .attr("text-anchor", (d,i) => {
                if (d.angle == 0)
                    return "start";
                else if (d.angle <= 90)
                    return "start";
                else if (d.angle > 90 && d.angle < 135)
                    return "start";
                else if (d.angle >= 135)
                    return "end";
                else
                    return "middle";
            })
            .style("font-size", mid_font_size * scale)
            .attr('transform', (d,i) => {
                if (d.angle == 90)
                    return 'rotate(' + (d.angle) + ',' + (d.x+1) + ',' + (d.y +1) + ')';
                else if (d.angle > 90 && d.angle < 135)
                    return 'rotate(' + (d.angle) + ',' + (d.x + 1) + ',' + (d.y + 1) + ')';
                else if (d.angle >= 135)
                    return 'rotate(' + (d.angle -180) + ',' + (d.x) + ',' + (d.y) + ')';
                else
                    return 'rotate(' + (d.angle) + ',' + (d.x) + ',' + (d.y) + ')';
            })
            .attr('rotate', (d,i) => {
                if (d.angle == 90)
                    return -90;
                else if (d.angle < 90 && d.angle >= 45)
                    return -90;
                else if (d.angle > 90 && d.angle < 135)
                    return -90;
                else
                    return 0;
            })
            .attr("dx", (d,i) => {
                if (d.angle == 0)
                    return "0.75";
                else if (d.angle >= 135)
                    return "-0.75";
                else if (d.angle > 90 && d.angle < 135)
                    return 2.25;
                else if (d.angle < 45)
                    return "-3.5";
                else if (d.angle < 90 && d.angle >= 45)
                    return "1.5";
                else if (d.angle == 90)
                    return "1.5";
                else
                    return 0;
            })
            .attr("dy", (d,i) => {
                if (d.angle <= 45)
                    return "0.40";
                else if (d.angle < 90 && d.angle > 45)
                    return "0.5";
                else if (d.angle == 90)
                    return "2.35";
                else if (d.angle > 90 && d.angle < 135)
                    return "2.75";
                else if (d.angle >= 135)
                    return "0.40";
                else
                    return 0;
            })
            .on("mouseover", function(){ d3.select(this).style("font-size", mid_font_size * scale * 1.2); })
            .on("mouseout",  function(){ d3.select(this).style("font-size", mid_font_size * scale); });
    }

    //中间站节点铺画
    function drawMidStationNode( nodes )
    {
        d3.select("g").selectAll(".mid_station_node").remove();

        //圆铺画
        var circle = d3.select("g").selectAll(".mid_station_node")
            .data(nodes)
            .enter()
            .append("circle")
            .attr("cx", function(d) { return d.x; })
            .attr("cy", function(d) { return d.y; })
            .attr("id", function(d) { return d.id; })
            .attr("lj", function(d) { return parseInt(d.fq/100); })
            .style("fill", function(d) { return lj_color_list[parseInt(d.fq/100)]; })
            .attr("class", "mid_station_node")
            .style("r", current_mid_circle_r)
            //中间站圆点点击生成径路：与大站圆点一致，发站固定钦州港(QVZ)，点击站为到站
            .on("click", function()
            {
                var d = d3.select(this).datum();
                if (jl && jl.generateJlByStation && d)
                    jl.generateJlByStation({id : d.id, lh : d.lh, zm : d.zm});
            })
            .on("mouseover", function() { d3.select(this).style("r", current_mid_circle_r * 1.3); })
            .on("mouseout",  function() { d3.select(this).style("r", current_mid_circle_r); });

        //若当前已有径路，按径路规则对新加载的中间站上色（途经站黄、终点站蓝）
        if (jl && jl.colorMidStationsOnPath) jl.colorMidStationsOnPath();

    }

    //单站补画：仅绘制一个中间站圆点 + 站名（绘制方式 / 样式与 drawMidStationNode / drawMidStationFont 完全一致）
    //用于在生成径路时，端点为中间站但尚未在地图上显示的情况下补画出来（fill 用传入的端点色）。
    //cls 可选：默认按普通中间站 class 绘制；补画径路端点时传 "jl_mid_endpoint"，使其不被
    //“关闭 / 切换中间站显示”的 selectAll(".mid_station_*").remove() 误删（样式完全一致，仅 class 不同）。
    this.drawMidStationOne = drawMidStationOne;
    function drawMidStationOne( station, color, gid, cls )
    {
        if (!station || !station.x || !station.zm) return;

        var id = gid ? gid : (station.id ? station.id : ("mid_" + station.zm));
        var nodeCls = cls ? cls : "mid_station_node";
        var textCls = cls ? cls : "mid_station_text";

        //圆点：与 drawMidStationNode 一致（半径随缩放同步）
        d3.select("g").append("circle")
            .attr("cx", station.x)
            .attr("cy", station.y)
            .attr("id", id)
            .attr("lj", station.fq ? parseInt(station.fq/100) : 0)
            .style("fill", color)
            .attr("class", nodeCls)
            .style("r", current_mid_circle_r)
            .on("click", function()
            {
                if (jl && jl.generateJlByStation)
                    jl.generateJlByStation({id : station.id, lh : station.lh, zm : station.zm});
            })
            .on("mouseover", function() { d3.select(this).style("r", current_mid_circle_r * 1.3); })
            .on("mouseout",  function() { d3.select(this).style("r", current_mid_circle_r); });

        //站名：与 drawMidStationFont 一致（旋转 / 偏移 / 字号同步）
        //注意：不在这里内联设置 fill，让 CSS .mid_station_text { fill:#E6E6E6 } 生效（与原中间站文字白色一致）
        var angle = (station.angle != null) ? station.angle : 0;
        d3.select("g").append("text")
            .attr("x", station.x)
            .attr("y", station.y)
            .text(station.zm)
            .attr("id", id)
            .attr("class", textCls)
            //复用中间站文字白色：原 drawMidStationFont 依赖 CSS .mid_station_text{fill:#E6E6E6}，
            //但补画用专属 class 以不被切换中间站误删，故这里内联设置同色（与中间站一致）
            .style("fill", "#E6E6E6")
            .attr("text-anchor", (function()
            {
                if (angle == 0 || angle <= 90)          return "start";
                else if (angle > 90 && angle < 135)      return "start";
                else if (angle >= 135)                   return "end";
                else                                     return "middle";
            })())
            .style("font-size", mid_font_size * scale)
            .attr('transform', (function()
            {
                if (angle == 90)                                   return 'rotate(' + angle + ',' + (station.x+1) + ',' + (station.y+1) + ')';
                else if (angle > 90 && angle < 135)                return 'rotate(' + angle + ',' + (station.x+1) + ',' + (station.y+1) + ')';
                else if (angle >= 135)                             return 'rotate(' + (angle-180) + ',' + station.x + ',' + station.y + ')';
                else                                               return 'rotate(' + angle + ',' + station.x + ',' + station.y + ')';
            })())
            .attr('rotate', (function()
            {
                if (angle == 90)                                   return -90;
                else if (angle < 90 && angle >= 45)                return -90;
                else if (angle > 90 && angle < 135)                return -90;
                else                                               return 0;
            })())
            .attr("dx", (function()
            {
                if (angle == 0)                                    return "0.75";
                else if (angle >= 135)                             return "-0.75";
                else if (angle > 90 && angle < 135)                return 2.25;
                else if (angle < 45)                               return "-3.5";
                else if (angle < 90 && angle >= 45)                return "1.5";
                else if (angle == 90)                              return "1.5";
                else                                               return 0;
            })())
            .attr("dy", (function()
            {
                if (angle <= 45)                                   return "0.40";
                else if (angle < 90 && angle > 45)                 return "0.5";
                else if (angle == 90)                              return "2.35";
                else if (angle > 90 && angle < 135)                return "2.75";
                else if (angle >= 135)                             return "0.40";
                else                                               return 0;
            })())
            .on("mouseover", function(){ d3.select(this).style("font-size", mid_font_size * scale * 1.2); })
            .on("mouseout",  function(){ d3.select(this).style("font-size", mid_font_size * scale); });
    }
}



