function Search( ) {

    var loading_index;                          //转圈

    //本地车站总库（大站 + 中间站），首次搜索时构建并预计算拼音，供纯前端匹配
    var localStationCache = null;
    function buildLocalStationList()
    {
        if (localStationCache) return localStationCache;
        var list = [];
        //大站：window.getMapDftPoint（数组，每项含 zm/lh）
        if (window.getMapDftPoint && window.getMapDftPoint.length)
        {
            window.getMapDftPoint.forEach(function(d){
                if (d && d.zm && d.lh) list.push({ zm : d.zm, lh : d.lh, id : d.id });
            });
        }
        //中间站：window.中间站数据 为 {路局编号:[{zm,lh,id}...], ...}
        if (window.中间站数据)
        {
            for (var lj in window.中间站数据)
            {
                var arr = window.中间站数据[lj];
                if (!arr) continue;
                arr.forEach(function(d){
                    if (d && d.zm && d.lh) list.push({ zm : d.zm, lh : d.lh, id : d.id });
                });
            }
        }
        //预计算拼音：全拼（无空格）与首字母，便于匹配
        if (window.pinyinUtil)
        {
            list.forEach(function(s){
                try {
                    s.py  = window.pinyinUtil.getPinyin(s.zm, '', false) || "";       //如 "qinzhougang"
                    s.py1 = window.pinyinUtil.getFirstLetter(s.zm) || "";             //如 "qzh"
                } catch (e) { s.py = ""; s.py1 = ""; }
            });
        }
        localStationCache = list;
        return list;
    }

    // -------- 最近搜索记录（localStorage，最多 10 条，新查置顶去重） --------
    var RECENT_KEY = "recentStations";
    var RECENT_MAX = 10;
    function loadRecent() {
        try {
            var arr = JSON.parse(localStorage.getItem(RECENT_KEY) || "[]");
            return Array.isArray(arr) ? arr : [];
        } catch (e) { return []; }
    }
    function saveRecent(row) {
        if (!row || !row.zm || !row.lh) return;
        var arr = loadRecent().filter(function (r) {
            return !(r.lh === row.lh && r.zm === row.zm);   //去重
        });
        arr.unshift({ zm: row.zm, lh: row.lh, id: row.id });
        if (arr.length > RECENT_MAX) arr = arr.slice(0, RECENT_MAX);
        try { localStorage.setItem(RECENT_KEY, JSON.stringify(arr)); } catch (e) {}
    }

    //当前唯一匹配项（输入收敛到唯一时暂存，等待用户回车确认）
    var _pendingUnique = null;
    this.submitIfUnique = function () {
        if (_pendingUnique) {
            var r = _pendingUnique;
            _pendingUnique = null;
            selectStation(r);
            return true;
        }
        return false;
    };

    //车站输入检索（工具栏搜索框回车/输入/聚焦调用）：
    //  - 空输入：显示最近 10 次搜索过的车站
    //  - 有输入：纯前端匹配本地车站库，得到电报码后生成径路
    this.stationInputSearch = stationInputSearch;
    function stationInputSearch()
    {
        var input = $("#stationSearchInput").val().trim();

        //空输入：展示最近搜索记录（聚焦或清空内容时触发）
        if (!input)
        {
            var recent = loadRecent();
            if (recent.length)
            {
                $("#stationMatchCount").text(recent.length);
                showStationDropdown(recent);
            }
            else
            {
                $("#stationMatchCount").text("0");
            }
            return;
        }

        var list = buildLocalStationList();
        var kw   = input.toLowerCase();
        var isLetter = /^[a-z]+$/.test(kw);     //纯字母：按电报码/拼音/首字母匹配
        var hits = [];
        for (var i = 0; i < list.length; i++)
        {
            var s = list[i];
            if (isLetter)
            {
                //电报码前缀/包含，或拼音首字母包含，或全拼包含
                if ((s.lh && s.lh.toLowerCase().indexOf(kw) >= 0) ||
                    (s.py1 && s.py1.toLowerCase().indexOf(kw) >= 0) ||
                    (s.py  && s.py.toLowerCase().indexOf(kw) >= 0))
                    hits.push(s);
            }
            else
            {
                //含汉字：站名包含 或 拼音/首字母包含（支持拼音串输入）
                if ((s.zm && s.zm.indexOf(input) >= 0) ||
                    (s.py  && s.py.toLowerCase().indexOf(kw) >= 0) ||
                    (s.py1 && s.py1.toLowerCase().indexOf(kw) >= 0))
                    hits.push(s);
            }
        }

        //按 电报码+站名 去重（与下拉一致），得到唯一候选项
        var uniq = [], seenKey = {};
        for (var j = 0; j < hits.length; j++) {
            var r = hits[j];
            if (!r || !r.lh || !r.zm) continue;
            var k = r.lh + "_" + r.zm;
            if (seenKey[k]) continue;
            seenKey[k] = true;
            uniq.push(r);
        }

        //实时更新匹配数量：找到显示数量，找不到显示 0（下拉保留上一次内容）
        $("#stationMatchCount").text(uniq.length);

        if (!uniq.length)
        {
            //输入内容但无匹配：兜底显示最近搜索记录，便于快速重选
            var recent = loadRecent();
            if (recent.length)
            {
                showStationDropdown(recent);
            }
            return;
        }

        //仅 1 个候选项时：正常显示下拉，暂存为待确认项，等待用户回车再生成径路
        if (uniq.length === 1)
        {
            _pendingUnique = uniq[0];
            showStationDropdown(uniq);
            return;
        }

        _pendingUnique = null;   //有多项匹配，清空待确认项（回车走正常重匹配）
        showStationDropdown(hits);
    }

    //选中某个车站：回填站名、记录最近搜索、生成径路（下拉项点击与“唯一直选”共用）
    function selectStation(row)
    {
        if (!row) return;
        $(".station_dropdown").remove();
        $("#stationSearchInput").val(row.zm);   //回填站名，便于确认刚选择了哪个站
        saveRecent(row);                         //记录到最近搜索
        if (jl && jl.generateJlByStation) jl.generateJlByStation(row);
    }

    //渲染候选下拉（选项：站名 电报码），选中调用 jl.generateJlByStation（固定钦州港为发站）
    function showStationDropdown(list)
    {
        $(".station_dropdown").remove();
        $(document).off("click.dropdown");   //清掉上一次可能残留的外部点击监听
        var ul = $('<ul class="station_dropdown"></ul>');
        //按 电报码+站名 去重，避免大站与中间站重复项
        var seen = {};
        $.each(list, function(i, row)
        {
            if (!row || !row.lh || !row.zm) return;
            var key = row.lh + "_" + row.zm;
            if (seen[key]) return;
            seen[key] = true;
            var li = $('<li></li>')
                .append($('<span class="sd_zm"></span>').text(row.zm))
                .append($('<span class="sd_lh"></span>').text(row.lh))
                .on("click", function(e)
                {
                    e.stopPropagation();
                    selectStation(row);
                });
            ul.append(li);
        });

        var pos = $("#stationSearchInput").offset();
        var inputW = $("#stationSearchInput").outerWidth();
        ul.css({
            top      : pos.top + $("#stationSearchInput").outerHeight(),
            left     : pos.left,
            minWidth : inputW,                 //至少与输入框等宽
            maxWidth : Math.max(inputW, 320)   //上限，避免过宽（宽度由 css fit-content 自适应）
        });
        $("body").append(ul);

        //点击外部关闭：延迟注册，且忽略点击输入框本身（避免聚焦时的点击误关刚弹出的下拉）
        setTimeout(function () {
            $(document).on("click.dropdown", function (e) {
                if ($(e.target).closest("#stationSearchInput, .station_dropdown").length) return;
                $(".station_dropdown").remove();
                $(document).off("click.dropdown");
            });
        }, 0);
    }





    function setPrompt( prompt ) {
        $(".right_svg_prompt").html(prompt);
    }
}