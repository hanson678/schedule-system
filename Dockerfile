# ============================================================
# 排期录入系统 - 生产环境 Dockerfile (multi-stage build)
# ============================================================

# ---------- 第一阶段：安装依赖 ----------
FROM python:3.12-slim AS builder

WORKDIR /build

# 系统依赖（pdfplumber 需要的库）
RUN apt-get update && \
    apt-get install -y --no-install-recommends gcc && \
    rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --prefix=/install -r requirements.txt gunicorn

# ---------- 第二阶段：生产镜像 ----------
FROM python:3.12-slim

LABEL maintainer="sales11@hanson2.com"
LABEL description="ZURU 排期录入系统"

# 安装运行时依赖 + cifs-utils（挂载网络共享盘用）
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
        cifs-utils \
        curl \
        tini && \
    rm -rf /var/lib/apt/lists/*

# 从 builder 阶段复制已安装的 Python 包
COPY --from=builder /install /usr/local

# 创建非 root 用户
RUN groupadd -r appuser && useradd -r -g appuser -d /app -s /sbin/nologin appuser

WORKDIR /app

# 复制应用代码
COPY app.py .
COPY excel_handler.py .
COPY excel_po_parser.py .
COPY pdf_parser.py .
COPY email_handler.py .
COPY templates/ templates/
COPY static/ static/

# 创建数据目录并设置权限
RUN mkdir -p data uploads /mnt/schedules && \
    chown -R appuser:appuser /app /mnt/schedules

# 环境变量
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    FLASK_ENV=production \
    CONTAINER_MODE=1 \
    Z_DRIVE_PATH=/mnt/schedules \
    APP_PORT=5000

# 健康检查
HEALTHCHECK --interval=30s --timeout=5s --start-period=10s --retries=3 \
    CMD curl -f http://localhost:${APP_PORT}/ || exit 1

EXPOSE ${APP_PORT}

USER appuser

# 用 tini 作为 PID 1，处理僵尸进程
ENTRYPOINT ["tini", "--"]

# 用 gunicorn 运行（生产环境不用 Flask 内置服务器）
CMD ["gunicorn", \
     "--bind", "0.0.0.0:5000", \
     "--workers", "2", \
     "--threads", "4", \
     "--timeout", "120", \
     "--access-logfile", "-", \
     "--error-logfile", "-", \
     "app:app"]
