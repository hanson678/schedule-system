# 总排期文件上传功能设计

## 需求
在现有路径浏览器选择方式基础上，新增"上传总排期"功能，用户可以直接上传本地 .xlsx 文件作为总排期。

## 方案
在顶部工具栏"选择总排期"和"恢复默认"按钮旁，新增"上传总排期"按钮。

## 后端 (app.py)
- 新增路由 `POST /api/master-schedule-upload-file`
- 接收 multipart/form-data 上传的 .xlsx 文件
- 保存到 `data/uploaded_master.xlsx`（覆盖旧文件）
- 自动将 `_custom_path` 切换到该文件
- 返回新路径和文件状态

## 前端 (templates/master.html)
- 在"选择总排期"按钮旁新增"上传总排期"按钮
- 隐藏 `<input type="file" accept=".xlsx,.xls">`
- 点击按钮触发文件选择 → 上传 → 更新路径显示和状态徽标

## 行为
- 上传成功后，系统自动使用上传的文件
- "恢复默认"仍可回到 Z 盘默认路径
- 路径显示框显示上传文件的完整路径
- 状态徽标正常刷新（绿色=可用）
