/**
 * integration.js —— 铁路地图对外集成接口
 * ============================================================================
 * 让外部页面（如「股道存车」明细）通过 postMessage 调用本图的径路功能，
 * 不改动 Map 现有的任何业务代码。
 *
 * 接入方式（对 chinamap.html 的唯一改动）：
 *       <script src="./js/integration.js"></script>
 *
 * 实现思路（尽量简单）：
 *   不重新实现径路算法，而是直接复用地图工具栏已有的搜索流程 ——
 *   把站名填进 #stationSearchInput，调用搜索并确认唯一项，
 *   后续全部交给 Map 自身的 selectStation() → jl.generateJlByStation()。
 *   起点固定为钦州港，由 generateJlByStation 内部决定，此处无需关心。
 *
 * 通信协议（file:// 下 origin 为 'null'，targetOrigin 只能用 '*'）：
 *   父 → 子：  { type: 'map:ping' }
 *              { type: 'map:path', to: '到站名' }
 *              { type: 'map:clear' }
 *   子 → 父：  { type: 'map:ready' }
 *              { type: 'map:pong',   ready: bool }
 *              { type: 'map:result', ok: bool, msg: '' }
 * ============================================================================
 */
(function () {
    'use strict';

    var ready = false;

    function post(obj) {
        try {
            if (window.parent && window.parent !== window) {
                window.parent.postMessage(obj, '*');
            }
        } catch (e) { /* 父页面未监听或不可达时忽略 */ }
    }

    /**
     * 按站名生成径路（起点：钦州港）
     * @param {string} name 到站名/电报码/拼音
     */
    function queryStation(name) {
        if (!window.search || !window.jl) {
            return { ok: false, msg: '地图尚未初始化完成' };
        }
        var kw = (name || '').trim();
        if (!kw) return { ok: false, msg: '站名为空' };

        try {
            // 填入搜索框并触发检索
            $('#stationSearchInput').val(kw);
            window.search.stationInputSearch();

            // 唯一直达：确认后由 Map 内部完成回填、记录、生成径路
            if (window.search.submitIfUnique && window.search.submitIfUnique()) {
                return { ok: true, msg: '' };
            }
            return { ok: false, msg: '未唯一匹配到车站：' + kw };
        } catch (e) {
            return { ok: false, msg: '生成径路出错：' + (e && e.message || e) };
        }
    }

    window.addEventListener('message', function (e) {
        var m = e.data;
        if (!m || typeof m !== 'object') return;

        if (m.type === 'map:ping') {
            post({ type: 'map:pong', ready: ready });
            return;
        }

        if (m.type === 'map:path') {
            var res = queryStation(m.to);
            post({ type: 'map:result', ok: !!res.ok, msg: res.msg || '' });
            return;
        }

        if (m.type === 'map:clear') {
            try { if (window.jl && window.jl.clearJl) window.jl.clearJl(); } catch (err) {}
            post({ type: 'map:result', ok: true, msg: 'cleared' });
        }
    });

    // 就绪检测：search / jl / map 均由 initMap 创建，需等待
    var tries = 0;
    var timer = setInterval(function () {
        if (ready) { clearInterval(timer); return; }
        if (window.search && window.jl && window.map &&
            window.$ && $('#stationSearchInput').length) {
            ready = true;
            post({ type: 'map:ready' });
            clearInterval(timer);
            return;
        }
        if (++tries > 600) clearInterval(timer);   // 最多等 60 秒
    }, 100);

    // 也留一个直接调用入口，便于将来不用 iframe 时改用 script 直引
    window.RailwayMapBridge = {
        queryStation: queryStation,
        isReady: function () { return ready; }
    };
})();
