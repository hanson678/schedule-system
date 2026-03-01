// ===== ZURU 排期智能录入 v4 =====
const App = {
    showLoading(text = '处理中...') {
        document.getElementById('loading-text').textContent = text;
        document.getElementById('loading-overlay').style.display = 'flex';
    },
    hideLoading() {
        document.getElementById('loading-overlay').style.display = 'none';
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
        const resp = await fetch(url, { ...defaults, ...options });
        const data = await resp.json();
        if (!resp.ok) throw new Error(data.error || `请求失败 (${resp.status})`);
        return data;
    }
};

// 时钟
setInterval(() => {
    const el = document.getElementById('clock');
    if (el) el.textContent = new Date().toLocaleString('zh-CN');
}, 1000);
