/* 状态面板：
 *   - 显示扩展是否就绪、报表标签页是否已被识别；
 *   - 提供输入框，让用户粘贴报表页完整地址；保存时写入 chrome.storage，
 *     由 background 解析出 REPORT_URL（原样）与 REPORT_MATCH（URL 的 host）。 */
function hostOf(u) {
  try { return new URL(u).host; } catch (e) { return ''; }
}

function refresh() {
  chrome.runtime.sendMessage({ type: 'status' }, function (res) {
    var st = document.getElementById('st');
    var tabsEl = document.getElementById('tabs');
    if (!res) { st.innerHTML = '<span class="no">扩展未响应，请刷新扩展</span>'; return; }
    if (res.reportTabs > 0) {
      st.innerHTML = '<span class="ok">就绪</span>';
      tabsEl.innerHTML = '<span class="ok">已打开 ' + res.reportTabs + ' 个</span>';
    } else {
      st.innerHTML = '<span class="no">未发现报表页</span>';
      tabsEl.innerHTML = '<span class="no">0（点「编好」时自动开）</span>';
    }
  });
}

document.addEventListener('DOMContentLoaded', function () {
  // 主题切换
  var themeToggle = document.getElementById('themeToggle');
  function applyTheme(isDark) {
    document.documentElement.setAttribute('data-theme', isDark ? 'dark' : '');
    if (themeToggle) themeToggle.textContent = isDark ? '☾' : '☀';
  }
  chrome.storage.local.get(['depTheme'], function (r) {
    applyTheme(r && r.depTheme === 'dark');
  });
  if (themeToggle) {
    themeToggle.addEventListener('click', function () {
      var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
      var next = !isDark;
      applyTheme(next);
      chrome.storage.local.set({ depTheme: next ? 'dark' : 'light' });
    });
  }

  // 标签切换：站存-发车流程填表桥 / 地图 Token 管理
  var tabs = document.querySelectorAll('.tab');
  tabs.forEach(function (tab) {
    tab.addEventListener('click', function () {
      tabs.forEach(function (t) { t.classList.remove('active'); });
      tab.classList.add('active');
      var target = tab.getAttribute('data-tab');
      document.getElementById('panel-fill').style.display = (target === 'fill') ? '' : 'none';
      document.getElementById('panel-token').style.display = (target === 'token') ? '' : 'none';
    });
  });

  var input = document.getElementById('urlInput');
  var parse = document.getElementById('parse');
  var btn = document.getElementById('saveBtn');

  // 回填已保存的地址
  chrome.storage.local.get(['depReportUrl'], function (r) {
    if (r && r.depReportUrl) input.value = r.depReportUrl;
    showParse();
  });

  function showParse() {
    var u = input.value.trim();
    if (!u) { parse.textContent = ''; return; }
    var h = hostOf(u);
    parse.textContent = h ? ('匹配特征（自动）：' + h) : '⚠ 地址无法解析，请检查格式';
  }

  input.addEventListener('input', showParse);

  // 地图 token：输入 URL / Referer 测试，或取回后注入地图
  var tokenBtn = document.getElementById('tokenBtn');
  var testBtn = document.getElementById('testBtn');
  var tokenInfo = document.getElementById('tokenInfo');
  var versionInfo = document.getElementById('versionInfo');
  var tokenUrl = document.getElementById('tokenUrl');
  var tokenReferer = document.getElementById('tokenReferer');

  // 默认地址常量（打开即填充，可手改）
  var DEFAULT_TOKEN_URL = 'http://10.208.2.72:8080/getToken?ip=10.208.2.72';
  var TEST_TOKEN_URL = 'http://localhost:3000/getToken?ip=10.208.2.72';
  var DEFAULT_REFERER = 'http://10.208.2.72:8080/cljl';

  // 回填上次填的 Token 地址 / Referer；没存过则用默认值（打开即有值，可手改）
  chrome.storage.local.get(['depTokenUrl', 'depTokenReferer'], function (r) {
    tokenUrl.value = (r && r.depTokenUrl) ? r.depTokenUrl : DEFAULT_TOKEN_URL;
    tokenReferer.value = (r && r.depTokenReferer) ? r.depTokenReferer : DEFAULT_REFERER;
  });

  // 默认 / 测试 Token 地址按钮
  var defaultTokenBtn = document.getElementById('defaultTokenBtn');
  var testTokenBtn = document.getElementById('testTokenBtn');
  if (defaultTokenBtn) {
    defaultTokenBtn.addEventListener('click', function () {
      tokenUrl.value = DEFAULT_TOKEN_URL;
    });
  }
  if (testTokenBtn) {
    testTokenBtn.addEventListener('click', function () {
      tokenUrl.value = TEST_TOKEN_URL;
    });
  }
  // Referer 默认按钮
  var defaultRefererBtn = document.getElementById('defaultRefererBtn');
  if (defaultRefererBtn) {
    defaultRefererBtn.addEventListener('click', function () {
      tokenReferer.value = DEFAULT_REFERER;
    });
  }

  function showResult(cls, text) {
    tokenInfo.className = 'v ' + cls;
    tokenInfo.textContent = text;
  }

  // 测试请求：先取 Token，再用该 Token 请求 /getReleaseVersionData 获取版本号并显示
  if (testBtn) {
    testBtn.addEventListener('click', function () {
      var url = tokenUrl.value.trim();
      if (!url) { showResult('no', '请填写 Token 地址'); return; }
      showResult('', '请求中…');
      if (versionInfo) versionInfo.textContent = '请求中…';
      chrome.runtime.sendMessage({
        type: 'testTokenAndVersion', url: url, referer: tokenReferer.value.trim()
      }, function (res) {
        if (!res) { showResult('no', '✗ 扩展无响应，请在扩展页点「重新加载」'); return; }

        // Token 结果（单行）
        if (!res.tokenOk) {
          showResult('no', '✗ 失败（' + (res.tokenStatus || 0) + '）');
        } else {
          var tk = String(res.token || '');
          showResult('ok', tk.length > 40 ? tk.slice(0, 40) + '…' : (tk || '✓ 成功'));
        }

        // 版本号结果（单行）
        if (versionInfo) {
          if (!res.versionOk) {
            versionInfo.className = 'v no';
            versionInfo.textContent = '✗ 失败（' + (res.versionStatus || 0) + '）';
          } else if (res.version) {
            versionInfo.className = 'v ok';
            versionInfo.textContent = res.version;
          } else {
            versionInfo.className = 'v';
            versionInfo.textContent = '✓ 空';
          }
        }
      });
    });
  }

  // 取回 token 并注入地图
  if (tokenBtn) {
    tokenBtn.addEventListener('click', function () {
      var url = tokenUrl.value.trim();
      var referer = tokenReferer.value.trim();
      chrome.storage.local.set({
        depTokenUrl: url, depTokenReferer: referer
      }, function () {
        chrome.runtime.sendMessage({ type: 'fetchMapToken', url: url, referer: referer }, showInject);
      });
    });
  }

  function showInject(res) {
    if (!res || !res.ok) {
      showResult('no', '✗ 失败：' + ((res && res.error) || '扩展无响应'));
      return;
    }
    showResult('ok', '✓ 已保存到本地');
  }

  // 打开面板时显示已保存的 token 状态（单行）
  chrome.storage.local.get(['authorization', 'authorizationTs'], function (r) {
    if (r && r.authorization) {
      var st = String(r.authorization);
      showResult('ok', st.length > 40 ? st.slice(0, 40) + '…' : st);
    } else {
      showResult('', '尚未获取 token');
    }
  });

  // 广播兜底开关：关掉后只发给「精确匹配」的报表页，用于验证地址是否配对
  var broadcastToggle = document.getElementById('broadcastToggle');
  var broadcastHint = document.getElementById('broadcastHint');
  if (broadcastToggle) {
    chrome.storage.local.get(['depBroadcast'], function (r) {
      // 默认开启（没存过就是开）
      broadcastToggle.checked = (r && r.depBroadcast !== undefined) ? !!r.depBroadcast : true;
      if (broadcastHint) {
        broadcastHint.textContent = broadcastToggle.checked
          ? '开：地址没配对也能找到报表页'
          : '关：只发给精确匹配的报表页';
      }
    });
    broadcastToggle.addEventListener('change', function () {
      chrome.storage.local.set({ depBroadcast: this.checked });
      if (broadcastHint) {
        broadcastHint.textContent = this.checked
          ? '开：地址没配对也能找到报表页'
          : '关：只发给精确匹配的报表页';
      }
    });
  }

  // 401 自动重试开关与重试次数
  var retryToggle = document.getElementById('retryToggle');
  var retryCount = document.getElementById('retryCount');

  chrome.storage.local.get(['depRetryEnabled', 'depRetryCount'], function (r) {
    if (retryToggle) retryToggle.checked = !!(r && r.depRetryEnabled);
    if (retryCount) retryCount.value = (r && r.depRetryCount != null) ? r.depRetryCount : 2;
  });

  if (retryToggle) {
    retryToggle.addEventListener('change', function () {
      chrome.storage.local.set({ depRetryEnabled: this.checked });
    });
  }
  if (retryCount) {
    retryCount.addEventListener('change', function () {
      var v = parseInt(this.value, 10);
      if (isNaN(v) || v < 0) v = 0;
      if (v > 10) v = 10;
      this.value = v;
      chrome.storage.local.set({ depRetryCount: v });
    });
  }

  btn.addEventListener('click', function () {
    var u = input.value.trim();
    if (!u) { parse.textContent = '⚠ 请先粘贴报表地址'; return; }
    if (!hostOf(u)) { parse.textContent = '⚠ 地址无法解析，请检查格式'; return; }
    chrome.storage.local.set({ depReportUrl: u }, function () {
      parse.textContent = '已保存，匹配特征：' + hostOf(u);
      refresh();
    });
  });

  refresh();
});