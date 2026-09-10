/* ============================================================================
 * map-version-badge.js —— 版本升级标记（页面侧，零侵入）
 *
 * 背景：本地兜底数据（map_data/*.js）里的每个元素都带 version 字段（如 "C260702"），
 * 而 /getJl 返回的径路节点 id 也带同一版本后缀，页面用 d3.select("#" + id) 去点亮
 * 团点 —— 所以「后端版本 ≠ 本地数据版本」时，径路着色会静默失效。
 *
 * 本脚本只做一件事：当后端版本号（getVersion 写入 .right_version 的值）与本地数据
 * 版本（window.getMapDftPoint[0].version，回退到页面初始文本）不一致时，
 * 在版本号前面插入一个「可升级」SVG 图标并给出 title 提示；一致则移除图标。
 *
 * 离线（本地数据集 __local__）模式下 map_version 与本地一致，故不会出现图标。
 *
 * 引入位置：chinamap.html 的 </body> 之前（此时 #cx1 已解析、map_data 已加载）。
 * ==========================================================================*/
(function () {
  'use strict';
  if (window.__depVersionBadgeInstalled) return;
  window.__depVersionBadgeInstalled = true;

  var ICON_ID = 'depVersionUpgradeIcon';
  var SVG_NS = 'http://www.w3.org/2000/svg';
  var COLOR = '#ffb300';                       // 琥珀色，深色地图顶栏上清晰可见

  var host = document.getElementById('cx1');   // <div id="cx1" class="right_version">C260702</div>
  if (!host) return;

  // 脚本加载时（早于 $(document).ready 里的 GetToken）记下页面初始文本，
  // 它即 HTML 里硬编码的本地兜底数据版本，作为取不到数据文件 version 时的回退值。
  var htmlVersion = String(host.textContent || '').trim();

  /** 本地兜底数据的版本：优先取数据文件自带的 version 字段（最真实），回退页面初始文本 */
  function localVersion() {
    try {
      var arr = window.getMapDftPoint;
      if (arr && arr.length && arr[0] && arr[0].version) {
        return String(arr[0].version).trim();
      }
    } catch (e) {}
    return htmlVersion;
  }

  /** 当前显示的版本号（由 getVersion() 经 $('.right_version').text(...) 写入） */
  function shownVersion() {
    return String(host.textContent || '').trim();   // SVG 不贡献文本，textContent 即纯版本号
  }

  /** 构造「圆圈 + 向上箭头」的升级图标 */
  function makeIcon() {
    var svg = document.createElementNS(SVG_NS, 'svg');
    svg.setAttribute('id', ICON_ID);
    svg.setAttribute('viewBox', '0 0 16 16');
    svg.setAttribute('width', '11');
    svg.setAttribute('height', '11');
    svg.setAttribute('style', 'vertical-align:-1px;margin-right:2px;flex:none;');

    var circle = document.createElementNS(SVG_NS, 'circle');
    circle.setAttribute('cx', '8'); circle.setAttribute('cy', '8'); circle.setAttribute('r', '7');
    circle.setAttribute('fill', 'none');
    circle.setAttribute('stroke', COLOR); circle.setAttribute('stroke-width', '1.8');

    var stem = document.createElementNS(SVG_NS, 'path');
    stem.setAttribute('d', 'M8 11.6V5.4');
    stem.setAttribute('stroke', COLOR); stem.setAttribute('stroke-width', '1.8');
    stem.setAttribute('stroke-linecap', 'round');

    var head = document.createElementNS(SVG_NS, 'path');
    head.setAttribute('d', 'M5.2 8.2 8 5.1l2.8 3.1');
    head.setAttribute('fill', 'none');
    head.setAttribute('stroke', COLOR); head.setAttribute('stroke-width', '1.8');
    head.setAttribute('stroke-linecap', 'round'); head.setAttribute('stroke-linejoin', 'round');

    svg.appendChild(circle); svg.appendChild(stem); svg.appendChild(head);
    return svg;
  }

  function sync() {
    var local = localVersion();
    var shown = shownVersion();
    if (!local || !shown) return;

    var needBadge = (shown !== local);
    var hasBadge = !!document.getElementById(ICON_ID);

    if (needBadge && !hasBadge) {
      // insertBefore 放在最前 = 版本号前面
      host.insertBefore(makeIcon(), host.firstChild);
      host.setAttribute('title', '后端已升级到 ' + shown + '，本地兜底数据为 ' + local +
                                 '，建议同步 Map/map_data 下的本地数据');
      // 图标会占一点宽度，让容器自适应但不换行（宽度仅「按需变宽」，不影响其它布局）
      host.style.whiteSpace = 'nowrap';
      host.style.width = 'auto';
      host.style.minWidth = '60px';
      console.log('[dep-badge] 后端版本 ' + shown + ' ≠ 本地数据版本 ' + local + ' → 已标记升级');
    } else if (!needBadge && hasBadge) {
      var ic = document.getElementById(ICON_ID);
      if (ic && ic.parentNode) ic.parentNode.removeChild(ic);
      host.removeAttribute('title');
      host.style.whiteSpace = '';
      host.style.width = '60px';
      host.style.minWidth = '';
      console.log('[dep-badge] 后端版本与本地数据一致（' + shown + '）→ 已移除升级标记');
    }
  }

  /* getVersion() 用 $('.right_version').text(map_version) 整体替换子节点，
   * 会把图标一起清掉 —— 用 MutationObserver 监听后在 sync() 里自动补回，幂等无死循环。 */
  try {
    new MutationObserver(sync).observe(host, { childList: true, characterData: true, subtree: true });
  } catch (e) {}
  sync();
})();
