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
  var tokenUrl = document.getElementById('tokenUrl');
  var tokenReferer = document.getElementById('tokenReferer');

  // 回填上次填的 Token 地址 / Referer
  chrome.storage.local.get(['depTokenUrl', 'depTokenReferer'], function (r) {
    if (r && r.depTokenUrl) tokenUrl.value = r.depTokenUrl;
    if (r && r.depTokenReferer) tokenReferer.value = r.depTokenReferer;
  });

  // 默认 Token 地址按钮
  var defaultTokenBtn = document.getElementById('defaultTokenBtn');
  if (defaultTokenBtn) {
    defaultTokenBtn.addEventListener('click', function () {
      tokenUrl.value = 'http://localhost:3000/getToken?ip=10.208.2.72';
    });
  }

  function showResult(cls, text) {
    tokenInfo.className = 'v ' + cls;
    tokenInfo.textContent = text;
  }

  // 测试请求：只回显状态码和响应内容，不注入地图
  if (testBtn) {
    testBtn.addEventListener('click', function () {
      var url = tokenUrl.value.trim();
      if (!url) { showResult('no', '请填写 Token 地址'); return; }
      showResult('', '请求中…');
      chrome.runtime.sendMessage({
        type: 'testToken', url: url, referer: tokenReferer.value.trim()
      }, function (res) {
        if (!res) { showResult('no', '✗ 扩展无响应，请在扩展页点「重新加载」'); return; }
        if (!res.ok) {
          showResult('no', '✗ 失败（' + (res.status || 0) + '）' + (res.error ? '\n' + res.error : '') +
                          '\n' + res.url + '\n耗时 ' + res.ms + 'ms');
          return;
        }
        var body = String(res.text || '');
        showResult('ok', '✓ 成功 HTTP ' + res.status + '（' + res.ms + 'ms）\n' +
                        '返回内容：' + (body.length > 300 ? body.slice(0, 300) + '…' : body));
      });
    });
  }

  // 取回 token 并注入地图
  if (tokenBtn) {
    tokenBtn.addEventListener('click', function () {
      var url = tokenUrl.value.trim();
      if (url) {
        chrome.storage.local.set({
          depTokenUrl: url, depTokenReferer: tokenReferer.value.trim()
        }, function () {
          chrome.runtime.sendMessage({ type: 'fetchMapToken' }, showInject);
        });
      } else {
        chrome.runtime.sendMessage({ type: 'fetchMapToken' }, showInject);
      }
    });
  }

  function showInject(res) {
    if (!res || !res.ok) {
      showResult('no', '✗ 失败：' + ((res && res.error) || '扩展无响应'));
      return;
    }
    var t = String(res.token || '');
    showResult('ok', '✓ 已保存到本地：' + (t.length > 40 ? t.slice(0, 40) + '…' : t) +
                    '\n获取时间：' + new Date().toLocaleString('zh-CN') +
                    '\n（地图打开/刷新后自动生效）');
  }

  // 打开面板时显示已保存的 token 状态
  chrome.storage.local.get(['authorization', 'authorizationTs'], function (r) {
    if (r && r.authorization) {
      var t = String(r.authorization);
      showResult('ok', '已保存 token：' + (t.length > 40 ? t.slice(0, 40) + '…' : t) +
                       '\n获取时间：' + new Date(r.authorizationTs || Date.now()).toLocaleString('zh-CN'));
    } else {
      showResult('', '尚未获取 token，点「刷新 Token」');
    }
  });

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