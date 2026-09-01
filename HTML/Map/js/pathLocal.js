// 本地径路生成（后端 /getJl 不可用时兜底）
// 数据来源：window.getMapDndPoint（节点/车站）、window.getMapDlnPoint（线段）。
// 算法：基于节点-边图的 Dijkstra（A* 几何启发），边权=地理距离，里程=线段 nnlc 累加。
// 优先队列使用原生二叉最小堆（不依赖任何第三方库）。
(function (global) {
    "use strict";

    // -------- 图构建（带缓存） --------
    var _graph = null;
    var _midSeq = 0;          // 中间站虚拟节点编号计数器（负数，避免与大站 node 冲突）
    var _midRegistered = {};  // 已注册中间站 id 缓存，避免重复注册

    function buildGraph() {
        if (_graph) return _graph;

        var nodes = global.getMapDndPoint;   // [{id, node, fq, lh, x, y, ...}]
        var lines = global.getMapDlnPoint;   // [{id, node1, node2, fq, x1,y1,x2,y2, nnlc, ...}]

        if (!nodes || !lines) {
            throw new Error("本地地图数据未加载（getMapDndPoint / getMapDlnPoint）");
        }

        // 索引：id -> 节点对象；node(数字) -> 节点对象
        var byId = {}, byNode = {};
        nodes.forEach(function (n) {
            byId[n.id] = n;
            byNode[n.node] = n;
        });

        // 邻接表：node(数字) -> [{to:node, w:距离, line:线段}]
        var adj = {};
        function addEdge(a, b, line) {
            var na = byNode[a], nb = byNode[b];
            if (!na || !nb) return;            // 端点无坐标则跳过
            var dx = na.x - nb.x, dy = na.y - nb.y;
            var w = Math.sqrt(dx * dx + dy * dy);
            (adj[a] = adj[a] || []).push({ to: b, w: w, line: line });
        }
        lines.forEach(function (ln) {
            var a = ln.node1, b = ln.node2;
            if (a == null || b == null) return;
            addEdge(a, b, ln);
            addEdge(b, a, ln);
        });

        _graph = { byId: byId, byNode: byNode, adj: adj, nodes: nodes, lines: lines };
        return _graph;
    }

    // 把单个中间站懒注册进图（按需，不一次性建全部 5845 个）：
    // 用负数 __negNode 作为图 key，通过 node1/node2（所属区段两端大站）连入铁路网。
    // 注册后才能作为径路起/终点被寻路算法搜索到。
    function registerMidStation(m) {
        if (!m || m.x == null || m.y == null) return null;
        var g = buildGraph();
        if (_midRegistered[m.id]) return g.byId[m.id];   // 已注册，直接返回

        var negNode = -1 - (++_midSeq);     // 负数编号，唯一
        m.__negNode = negNode;
        g.byId[m.id] = m;
        g.byNode[negNode] = m;

        function geoDist(a, b) {
            var dx = a.x - b.x, dy = a.y - b.y;
            return Math.sqrt(dx * dx + dy * dy);
        }
        // 连到相邻大站 node1（必存在）
        if (m.node1 != null && g.byNode[m.node1]) {
            var w1 = geoDist(m, g.byNode[m.node1]);
            (g.adj[negNode] = g.adj[negNode] || []).push({ to: m.node1, w: w1, line: null });
            (g.adj[m.node1] = g.adj[m.node1] || []).push({ to: negNode, w: w1, line: null });
        }
        // 连到相邻大站 node2（若存在且有效）
        if (m.node2 != null && m.node2 >= 0 && g.byNode[m.node2]) {
            var w2 = geoDist(m, g.byNode[m.node2]);
            (g.adj[negNode] = g.adj[negNode] || []).push({ to: m.node2, w: w2, line: null });
            (g.adj[m.node2] = g.adj[m.node2] || []).push({ to: negNode, w: w2, line: null });
        }
        _midRegistered[m.id] = true;
        return m;
    }

    // 从本地中间站数据中按 zm / lh / id 查找完整中间站对象
    function findMidData(stop) {
        var midData = global["中间站数据"];
        if (!midData || !stop) return null;
        for (var lj in midData) {
            var arr = midData[lj];
            if (!arr) continue;
            for (var k = 0; k < arr.length; k++) {
                var m = arr[k];
                if (!m) continue;
                if ((stop.id && m.id === stop.id) ||
                    (stop.lh && m.lh === stop.lh) ||
                    (stop.zm && m.zm === stop.zm))
                    return m;
            }
        }
        return null;
    }

    // -------- 定位起/终点节点对象 --------
    // 优先级：① id 精确命中（大站或已注册中间站）；② 大站 lh；
    //         ③ 中间站 lh / zm（找到后懒注册进图，使算法可搜索）。
    function locate(stop) {
        if (!stop) return null;
        var g = buildGraph();
        if (stop.id && g.byId[stop.id]) return g.byId[stop.id];
        if (stop.lh) {
            for (var i = 0; i < g.nodes.length; i++) {
                if (g.nodes[i].lh === stop.lh) return g.nodes[i];
            }
        }
        // 中间站：从中间站数据定位后懒注册进图
        var m = findMidData(stop);
        if (m) return registerMidStation(m);
        return null;
    }

    // -------- 取某站相邻的一条线段（用于填充 fz/dz 的 node1/node2） --------
    function adjacentLine(nodeObj) {
        if (!nodeObj) return null;
        var g = buildGraph();
        var list = g.adj[nodeObj.node] || [];
        return list.length ? list[0].line : null;
    }

    // -------- 原生二叉最小堆（按 .f 比较） --------
    function MinHeap() { this._a = []; }
    MinHeap.prototype._less = function (i, j) { return this._a[i].f < this._a[j].f; };
    MinHeap.prototype.push = function (x) {
        var a = this._a, i = a.length;
        a.push(x);
        while (i > 0) {
            var p = (i - 1) >> 1;
            if (this._less(i, p)) { var t = a[i]; a[i] = a[p]; a[p] = t; i = p; }
            else break;
        }
    };
    MinHeap.prototype.pop = function () {
        var a = this._a, n = a.length;
        if (n === 0) return null;
        var top = a[0], last = a.pop();
        if (n > 1) {
            a[0] = last;
            var i = 0, sz = a.length;
            while (true) {
                var l = 2 * i + 1, r = 2 * i + 2, m = i;
                if (l < sz && this._less(l, m)) m = l;
                if (r < sz && this._less(r, m)) m = r;
                if (m === i) break;
                var t = a[i]; a[i] = a[m]; a[m] = t; i = m;
            }
        }
        return top;
    };
    MinHeap.prototype.empty = function () { return this._a.length === 0; };

    // -------- Dijkstra（A* 几何启发） --------
    function findPath(startId, endId) {
        var g = buildGraph();
        var sN = g.byId[startId], eN = g.byId[endId];
        if (!sN || !eN) return null;

        // 图节点 key：中间站用 __negNode（负数），大站用 node（正整数），二者在 byNode/adj 中一致
        function keyOf(o) { return (o.__negNode != null) ? o.__negNode : o.node; }
        var sk = keyOf(sN), ek = keyOf(eN);
        if (sk == null || ek == null) return null;

        var ex = eN.x, ey = eN.y;
        function h(n) { var dx = n.x - ex, dy = n.y - ey; return Math.sqrt(dx * dx + dy * dy); }

        var open = new MinHeap();   // 原生二叉最小堆，不依赖第三方库

        var gScore = {};            // node(图key) -> 已得最短距离
        var prev = {};              // node(图key) -> 前驱 {node, line}
        var closed = {};

        gScore[sk] = 0;
        open.push({ node: sk, f: h(sN) });

        while (!open.empty()) {
            var cur = open.pop();
            var cn = cur.node;
            if (closed[cn]) continue;
            closed[cn] = true;
            if (cn === ek) break;

            var edges = g.adj[cn] || [];
            for (var i = 0; i < edges.length; i++) {
                var e = edges[i];
                if (closed[e.to]) continue;
                var tentative = gScore[cn] + e.w;
                if (gScore[e.to] == null || tentative < gScore[e.to]) {
                    gScore[e.to] = tentative;
                    prev[e.to] = { node: cn, line: e.line };
                    var nodeObj = g.byNode[e.to];
                    open.push({ node: e.to, f: tentative + h(nodeObj) });
                }
            }
        }

        if (prev[ek] == null && ek !== sk) return null;   // 不可达

        // 回溯路径（终点 -> 起点）
        var seq = [];
        var n = ek;
        while (n != null) {
            var obj = g.byNode[n];
            seq.unshift(obj);
            var p = prev[n];
            n = p ? p.node : null;
        }
        if (seq[0] !== sN) seq.unshift(sN);   // 兜底保证起点在列
        // 给每个节点补 node 字段（统一为图 key），供里程累加与下游消费
        seq.forEach(function (o) { if (o.node == null) o.node = keyOf(o); });

        // 计算里程（累加线段 nnlc）
        var lc = 0;
        for (var k = 1; k < seq.length; k++) {
            var a = seq[k - 1].node, b = seq[k].node;
            var es = (g.adj[a] || []).concat(g.adj[b] || []);
            for (var m = 0; m < es.length; m++) {
                if ((es[m].to === b && es[m].line) || (es[m].to === a && es[m].line)) {
                    var nn = es[m].line.nnlc;
                    if (typeof nn === "number") { lc += nn; break; }
                }
            }
        }
        return { nodes: seq, lc: lc };
    }

    // -------- 组装与后端兼容的 jlout --------
    // 返回 { jlout: object|null, reason: string|null }，失败时 reason 给出精确原因。
    function makeJlout(fz_dz_array) {
        if (!fz_dz_array || fz_dz_array.length < 2) {
            return { jlout: null, reason: "起/到站列表为空（需至少 2 个站点）" };
        }

        var nodes = global.getMapDndPoint, lines = global.getMapDlnPoint;
        if (!nodes || !nodes.length) return { jlout: null, reason: "本地节点数据未加载（getMapDndPoint）" };
        if (!lines || !lines.length) return { jlout: null, reason: "本地线段数据未加载（getMapDlnPoint）" };

        var stops = fz_dz_array.slice();   // 支持多点经由：逐段拼接
        var pathNodes = [];
        var totalLc = 0;

        function stopDesc(stop, idx) {
            return "第" + (idx + 1) + "站(" + ((stop && (stop.zm || stop.lh || stop.id)) || "?") + ")";
        }

        for (var s = 0; s < stops.length - 1; s++) {
            var sn = locate(stops[s]), en = locate(stops[s + 1]);
            if (!sn) {
                return { jlout: null,
                    reason: "本地数据中未找到" + stopDesc(stops[s], s) +
                            "（id=" + (stops[s] && stops[s].id) + ", lh=" + (stops[s] && stops[s].lh) + "）" };
            }
            if (!en) {
                return { jlout: null,
                    reason: "本地数据中未找到" + stopDesc(stops[s + 1], s + 1) +
                            "（id=" + (stops[s + 1] && stops[s + 1].id) + ", lh=" + (stops[s + 1] && stops[s + 1].lh) + "）" };
            }
            var res = findPath(sn.id, en.id);
            if (!res) {
                return { jlout: null,
                    reason: stopDesc(stops[s], s) + " 到 " + stopDesc(stops[s + 1], s + 1) +
                            " 在本地铁路图中不可达（图不连通）" };
            }
            totalLc += res.lc;
            // 拼接：首段全量，后续段跳过首节点避免重复
            var part = res.nodes;
            for (var i = (s === 0 ? 0 : 1); i < part.length; i++) {
                var nd = part[i];
                pathNodes.push({
                    id: nd.id, x: nd.x, y: nd.y, lh: nd.lh, fq: nd.fq,
                    node: nd.node, zm: nd.lh
                });
            }
        }
        if (pathNodes.length < 2) {
            return { jlout: null, reason: "拼装后径路节点不足 2 个" };
        }

        function stopInfo(stop) {
            var nd = locate(stop);
            var ln = nd ? adjacentLine(nd) : null;
            return {
                id: stop.id,
                lh: stop.lh || (nd && nd.lh),
                zm: stop.zm || (nd && nd.lh),
                node1: ln ? ln.node1 : null,
                node2: ln ? ln.node2 : null,
                fq: nd ? nd.fq : null
            };
        }

        return {
            jlout: {
                fz: stopInfo(stops[0]),
                dz: stopInfo(stops[stops.length - 1]),
                path: pathNodes,
                fjz: [],                       // 本地无法可靠判定分界站
                lc: [Math.round(totalLc)],     // [全程公里]
                lj: [],                        // 本地不拆分路局里程
                zf: "本地兜底径路"
            },
            reason: null
        };
    }

    global.RailwayPath = {
        buildGraph: buildGraph,
        findPath: findPath,
        makeJlout: makeJlout,
        locate: locate,
        reset: function () { _graph = null; }
    };
})(window);
