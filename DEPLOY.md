# ZURU 排期录入系统 v4.1 — 完整部署指南

## 目录结构

```
schedule-system/
├── app.py              # Flask 主程序
├── excel_handler.py    # Excel 读写核心（WPS COM）
├── pdf_parser.py       # PDF 解析
├── excel_po_parser.py  # Excel PO 解析
├── email_handler.py    # 邮件收取
├── requirements.txt    # Python 依赖包
├── 启动系统.bat         # 一键启动脚本
├── templates/          # 网页模板
├── static/             # 前端静态文件
├── schedules/          # 排期 Excel 文件（初始数据）
└── data/               # 运行数据（config.json 等）
```

---

## 一、前提条件（必须满足）

| 条件 | 说明 |
|------|------|
| **Windows 系统** | 写入Excel 用 WPS COM，仅 Windows 可用 |
| **WPS Office 已安装** | 金山 WPS，需包含表格组件 |
| **Python 3.12** | 官网下载：https://www.python.org/downloads/ |
| **网络访问 GitHub** | 用于下载代码 |

---

## 二、安装步骤

### 1. 下载代码

```bash
git clone https://github.com/hanson678/schedule-system.git
cd schedule-system
```

或直接在 GitHub 页面点击 **Code → Download ZIP** 解压。

### 2. 安装依赖包

打开命令提示符（CMD）或 Git Bash，在项目目录执行：

```bash
pip install -r requirements.txt
```

如果 pip 找不到，用完整路径：
```bash
C:\Users\你的用户名\AppData\Local\Programs\Python\Python312\python.exe -m pip install -r requirements.txt
```

主要依赖：`flask`, `openpyxl`, `pdfplumber`, `pywin32`

### 3. 启动系统

**方式A**（推荐）：直接双击 `启动系统.bat`

**方式B**：命令行启动
```bash
python app.py
```

启动后浏览器访问：**http://localhost:5000**

---

## 三、首次使用配置

### 配置排期文件路径

1. 打开系统首页（http://localhost:5000）
2. 顶部有"**当前排期路径**"输入框
3. 点击"**切换**"按钮 → 弹出文件夹浏览器
4. 可以自由导航到**任意文件夹**（本地盘、网络共享盘均可）
5. 找到存放排期 Excel 文件的文件夹后，点击"**使用此目录**"

> 初始排期文件在 `schedules/` 文件夹内（37个排期文件），可复制到你想要的路径。

路径设置后会保存到 `data/config.json`，重启系统后自动记忆。

---

## 四、如果系统被挂在子路径下（如 /schedule/）

当系统通过反向代理部署在 `/schedule/` 路径下（不是直接在根路径），需要额外配置：

**方式A**（推荐）— 启动时设置环境变量：
```bat
set APP_PREFIX=/schedule
python app.py
```

或在 `启动系统.bat` 中加入 `set APP_PREFIX=/schedule` 一行。

**方式B** — 在上层 nginx 中加一行：
```nginx
proxy_set_header X-Script-Name /schedule;
```

---

## 五、常见问题

### Q: 提示"COM 错误"或"WPS 未找到"
A: 确认 WPS Office 已安装，且版本包含 WPS 表格（Spreadsheet）。

### Q: pip install 某个包失败
A: 单独安装 pywin32：
```bash
pip install pywin32==311
python -m pywin32_postinstall -install
```

### Q: 浏览排期路径时看不到文件
A: 确认目标文件夹内有 `.xlsx` 文件，且文件名不以 `~$` 开头（临时文件会被过滤）。

### Q: 系统启动后访问 5000 端口被占用
A: 修改 `app.py` 最后一行的 `port=5000` 为其他端口（如 `port=8080`）。

---

## 六、本次代码修改说明（v4.1 更新内容）

以下是从原版到 v4.1 做的所有改动，供部署维护参考：

### 1. 反向代理 URL 前缀支持
- **文件**: `app.py`
- **改动**: 新增 `ReverseProxyMiddleware` 中间件
- **效果**: 支持挂在 `/schedule/` 子路径下正常运行（静态文件、API、页面链接全部自动加前缀）
- **配置**: 通过环境变量 `APP_PREFIX=/schedule` 或 nginx 头 `X-Script-Name` 激活

### 2. 模板路径全部改为动态生成
- **文件**: `templates/base.html`, `templates/index.html`, `templates/history.html`
- **改动**: 所有 `/static/...`、`/history` 等硬编码路径改为 Flask `url_for()` 动态生成
- **效果**: 无论部署在什么路径下，链接都正确

### 3. JavaScript API 调用自动加前缀
- **文件**: `static/js/app.js`
- **改动**: `App.api()` 函数自动读取 `window.BASE_URL` 并加到 API 路径前
- **效果**: 所有 `/api/...` 调用在子路径部署时自动变成 `/schedule/api/...`

### 4. 跨 Sheet 行号偏移 Bug 修复
- **文件**: `excel_handler.py`（`batch_process` 函数）
- **改动**: `deleted_rows` 改为 `[(sheet名, 行号)]` 元组列表；`inserted_positions` 改为 `{sheet名: [行号列表]}` 字典
- **效果**: 同一 Excel 文件内多个 Sheet 同时操作时，行号偏移不再互相影响

  **修改前（有 Bug）：**
  ```python
  deleted_rows = []
  inserted_positions = []
  # ...
  shift_del = sum(1 for d in deleted_rows if d < ref)  # 错误：不区分sheet
  ```

  **修改后（已修复）：**
  ```python
  deleted_rows = []  # [(sheet_name, row)] 元组
  inserted_positions = {}  # {sheet_name: [pos_list]}
  last_insert_pos = {}     # {sheet_name: pos}
  # ...
  shift_del = sum(1 for dsn, d in deleted_rows if dsn == sn and d < ref)  # 只统计同sheet
  ```

### 5. 参考行精确匹配（SKU-SPEC 三级优先级）
- **文件**: `excel_handler.py`（`_search_sku_in_file`、`_search_sku_com`、`auto_find` 函数）
- **改动**: 新增 `_sku_spec()` 函数，搜索时三级优先级：① SKU-SPEC 精确（如 92105-S001）→ ② 基础货号（92105）→ ③ 前缀匹配
- **效果**: 同货号不同变体（如 92105-S001 和 92105-S004）能正确区分，中文名/内箱/外箱从正确的参考行复制

### 6. 货号（ITEM#）用 PDF 数据覆写
- **文件**: `excel_handler.py`（`_do_new_com` 函数，约第 1791 行）
- **改动**: 原来"不覆写货号"→ 改为明确写入 PDF 的 `sku_spec` 值
- **效果**: 写入行的货号和 PDF 完全一致（如"9296-S001"）

### 7. 货号大小写跟随排期
- **文件**: `excel_handler.py`（`_do_new_com` 函数）
- **改动**: 写入货号前先读参考行已有货号大小写，若忽略大小写后相同则沿用排期的格式
- **效果**: PDF 给"9296-S001"但排期已有"9296-s001"，则写入"9296-s001"

### 8. 新增功能：删除已入单行（可撤销）
- **文件**: `excel_handler.py`（`delete_entries_com` 方法）、`app.py`（`/api/delete-entries`）、`templates/index.html`
- **效果**: 在"已完成"区域可选中已入单行删除，支持一键撤销

### 9. 新增功能：一键重试 + 定时入单
- **文件**: `excel_handler.py`（`reentry_batch` 方法）、`app.py`（`/api/reentry`、`/api/schedule-retry`）、`templates/index.html`
- **效果**: 可立即重试入单，或设定时间自动入单，结果反馈与正常入单相同

---

## 七、技术架构说明

| 组件 | 技术 | 说明 |
|------|------|------|
| Web 框架 | Flask 3.1.3 | Python |
| Excel 写入 | WPS COM（Ket.Application v12.0）| Windows 专属，保持格式不破坏其他行 |
| Excel 读取 | openpyxl 3.1.5 | 只读搜索用 |
| PDF 解析 | pdfplumber | 解析 ZURU PO PDF |
| 前端 | Bootstrap 5.3 + 原生 JS | 无框架依赖 |
| 数据存储 | JSON 文件（data/ 目录）| 不依赖数据库 |

---

*最后更新：2026-03-01 | 华登玩具集团 排期系统 v4.1*
