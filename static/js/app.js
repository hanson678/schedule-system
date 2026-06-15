// ===== ZURU 排期智能录入 v4 =====
const App = {
    _progressTimer: null,
    showLoading(text = '处理中...', opts = {}) {
        document.getElementById('loading-text').textContent = text;
        const wrap = document.getElementById('loading-progress-wrap');
        const sub = document.getElementById('loading-sub');
        if (wrap) wrap.style.display = opts.progress ? 'block' : 'none';
        if (sub) sub.textContent = opts.sub || '通过WPS COM写入，不影响其他行格式';
        if (opts.progress) App.setProgress(0, '');
        document.getElementById('loading-overlay').style.display = 'flex';
    },
    hideLoading() {
        if (App._progressTimer) { clearInterval(App._progressTimer); App._progressTimer = null; }
        document.getElementById('loading-overlay').style.display = 'none';
    },
    setProgress(pct, label = '') {
        const bar = document.getElementById('loading-progress-bar');
        const pctEl = document.getElementById('loading-progress-pct');
        const labelEl = document.getElementById('loading-progress-label');
        if (bar) bar.style.width = pct + '%';
        if (pctEl) pctEl.textContent = pct + '%';
        if (labelEl && label) labelEl.textContent = label;
    },
    startRealProgress() {
        const stageMap = {'读取文件...': 15, '搜索排期...': 40, '匹配货号...': 70, '整理结果...': 90, '完成': 100};
        App._progressTimer = setInterval(async () => {
            try {
                const resp = await fetch((window.BASE_URL || '') + '/api/batch-progress');
                const d = await resp.json();
                if (!d.running && d.done === 0) return; // 还没开始
                const pct = stageMap[d.current] || Math.round((d.done / Math.max(d.total, 1)) * 100);
                App.setProgress(Math.min(pct, 95), d.current || '处理中...');
            } catch(e) {} // 网络错误静默忽略
        }, 600);
    },
    alert(message, type = 'success') {
        const el = document.getElementById('global-alert');
        const text = document.getElementById('global-alert-text');
        el.className = `alert alert-${type} alert-dismissible fade show`;
        text.textContent = message;
        el.style.display = 'block';
        if (type === 'success') setTimeout(() => { el.style.display = 'none'; }, 4000);
    },
    async api(url, options = {}) {
        // 自动加上反向代理前缀（如 /schedule）
        const baseUrl = window.BASE_URL || '';
        if (url.startsWith('/') && baseUrl) url = baseUrl + url;
        const defaults = { headers: { 'Content-Type': 'application/json' } };
        if (options.body && typeof options.body === 'object' && !(options.body instanceof FormData)) {
            options.body = JSON.stringify(options.body);
        }
        if (options.body instanceof FormData) {
            delete defaults.headers['Content-Type'];
        }
        // 超时保护：防止服务器重启导致请求永远挂起
        const timeout = options.timeout || 180000; // 默认3分钟
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), timeout);
        try {
            const resp = await fetch(url, { ...defaults, ...options, signal: controller.signal });
            clearTimeout(timer);
            const ct = resp.headers.get('content-type') || '';
            if (!ct.includes('json')) {
                throw new Error(`服务器返回了非JSON响应(${resp.status})，请刷新页面重试`);
            }
            const data = await resp.json();
            if (!resp.ok) throw new Error(data.error || `请求失败 (${resp.status})`);
            return data;
        } catch (e) {
            clearTimeout(timer);
            if (e.name === 'AbortError') throw new Error('请求超时，请刷新页面重试');
            throw e;
        }
    }
};

// 时钟
setInterval(() => {
    const el = document.getElementById('clock');
    if (el) el.textContent = new Date().toLocaleString('zh-CN');
}, 1000);
