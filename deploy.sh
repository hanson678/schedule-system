#!/usr/bin/env bash
# ============================================================
# 排期录入系统 - 一键部署脚本
# 用法:
#   ./deploy.sh --server 1.2.3.4
#   ./deploy.sh --server 1.2.3.4 --tag v1.0.0
#   ./deploy.sh --server 1.2.3.4 --registry registry.example.com
#   ./deploy.sh --rollback --server 1.2.3.4
# ============================================================
set -euo pipefail

# =================== 默认配置 ===================
SERVER=""
SSH_USER="root"
SSH_PORT="22"
REGISTRY=""
TAG=""
DEPLOY_PATH="/opt/schedule-system"
ROLLBACK=false
ENV_FILE=".env"

# 颜色输出
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

log_info()  { echo -e "${GREEN}[信息]${NC} $1"; }
log_warn()  { echo -e "${YELLOW}[警告]${NC} $1"; }
log_error() { echo -e "${RED}[错误]${NC} $1"; }

# =================== 参数解析 ===================
while [[ $# -gt 0 ]]; do
    case $1 in
        --server)   SERVER="$2"; shift 2 ;;
        --user)     SSH_USER="$2"; shift 2 ;;
        --port)     SSH_PORT="$2"; shift 2 ;;
        --registry) REGISTRY="$2"; shift 2 ;;
        --tag)      TAG="$2"; shift 2 ;;
        --path)     DEPLOY_PATH="$2"; shift 2 ;;
        --rollback) ROLLBACK=true; shift ;;
        --env)      ENV_FILE="$2"; shift 2 ;;
        *)          log_error "未知参数: $1"; exit 1 ;;
    esac
done

# 从 .env 文件读取默认值（如果命令行没指定）
if [[ -f "$ENV_FILE" ]]; then
    log_info "读取配置文件: $ENV_FILE"
    source "$ENV_FILE" 2>/dev/null || true
    SERVER="${SERVER:-${DEPLOY_SERVER:-}}"
    SSH_USER="${SSH_USER:-${DEPLOY_USER:-root}}"
    SSH_PORT="${SSH_PORT:-${DEPLOY_SSH_PORT:-22}}"
    DEPLOY_PATH="${DEPLOY_PATH:-${DEPLOY_PATH:-/opt/schedule-system}}"
fi

# 验证必要参数
if [[ -z "$SERVER" ]]; then
    log_error "必须指定服务器地址: --server <IP或域名>"
    exit 1
fi

# 自动生成 tag（用 git commit hash 或时间戳）
if [[ -z "$TAG" ]]; then
    if git rev-parse --short HEAD &>/dev/null; then
        TAG="$(git rev-parse --short HEAD)"
    else
        TAG="$(date +%Y%m%d-%H%M%S)"
    fi
fi

IMAGE_NAME="schedule-system"
if [[ -n "$REGISTRY" ]]; then
    FULL_IMAGE="${REGISTRY}/${IMAGE_NAME}:${TAG}"
else
    FULL_IMAGE="${IMAGE_NAME}:${TAG}"
fi

SSH_CMD="ssh -p ${SSH_PORT} ${SSH_USER}@${SERVER}"

# =================== 回滚 ===================
if [[ "$ROLLBACK" == true ]]; then
    log_info "开始回滚到上一个版本..."
    $SSH_CMD << 'ROLLBACK_EOF'
        set -e
        cd /opt/schedule-system
        PREV_TAG=$(cat .previous_tag 2>/dev/null || echo "")
        if [[ -z "$PREV_TAG" ]]; then
            echo "[错误] 没有找到上一个版本记录，无法回滚"
            exit 1
        fi
        echo "[信息] 回滚到版本: $PREV_TAG"
        export IMAGE_TAG="$PREV_TAG"
        docker compose -f docker-compose.prod.yml up -d
        echo "[信息] 等待服务启动..."
        sleep 5
        if docker compose -f docker-compose.prod.yml ps | grep -q "healthy"; then
            echo "[信息] 回滚成功！"
        else
            echo "[警告] 服务可能未完全启动，请检查日志"
            docker compose -f docker-compose.prod.yml logs --tail=20
        fi
ROLLBACK_EOF
    exit 0
fi

# =================== 正常部署流程 ===================
log_info "=============================="
log_info "  排期录入系统 - 部署开始"
log_info "=============================="
log_info "服务器: ${SSH_USER}@${SERVER}:${SSH_PORT}"
log_info "镜像: ${FULL_IMAGE}"
log_info "部署路径: ${DEPLOY_PATH}"

# ---------- 第1步：本地构建镜像 ----------
log_info "[1/6] 构建 Docker 镜像..."
docker build -t "${FULL_IMAGE}" -t "${IMAGE_NAME}:latest" .

# ---------- 第2步：推送镜像到 Registry ----------
if [[ -n "$REGISTRY" ]]; then
    log_info "[2/6] 推送镜像到 Registry..."
    docker push "${FULL_IMAGE}"
else
    log_info "[2/6] 没有配置 Registry，通过 SSH 直接传输镜像..."
    docker save "${FULL_IMAGE}" | gzip | \
        $SSH_CMD "gunzip | docker load"
fi

# ---------- 第3步：上传配置文件到远端 ----------
log_info "[3/6] 同步配置文件到服务器..."
$SSH_CMD "mkdir -p ${DEPLOY_PATH}/nginx ${DEPLOY_PATH}/scripts"

# 上传 compose 文件和配置
scp -P "${SSH_PORT}" \
    docker-compose.prod.yml \
    "${ENV_FILE}" \
    "${SSH_USER}@${SERVER}:${DEPLOY_PATH}/"

scp -P "${SSH_PORT}" \
    nginx/nginx.conf \
    "${SSH_USER}@${SERVER}:${DEPLOY_PATH}/nginx/"

scp -P "${SSH_PORT}" \
    scripts/backup.sh \
    scripts/restore.sh \
    "${SSH_USER}@${SERVER}:${DEPLOY_PATH}/scripts/"

# ---------- 第4步：在远端启动服务 ----------
log_info "[4/6] 在远端启动服务..."
$SSH_CMD << DEPLOY_EOF
    set -e
    cd ${DEPLOY_PATH}

    # 记录当前版本（用于回滚）
    CURRENT_TAG=\$(grep "^IMAGE_TAG=" .env 2>/dev/null | cut -d= -f2 || echo "")
    if [[ -n "\$CURRENT_TAG" ]]; then
        echo "\$CURRENT_TAG" > .previous_tag
    fi

    # 更新镜像标签
    sed -i "s/^IMAGE_TAG=.*/IMAGE_TAG=${TAG}/" .env 2>/dev/null || \
        echo "IMAGE_TAG=${TAG}" >> .env

    # 拉取镜像（如果用了 Registry）
    if [[ -n "${REGISTRY}" ]]; then
        docker pull "${FULL_IMAGE}"
    fi

    # 停止旧服务并启动新服务
    docker compose -f docker-compose.prod.yml down --timeout 30
    docker compose -f docker-compose.prod.yml up -d

    echo "[信息] 等待服务启动..."
    sleep 8
DEPLOY_EOF

# ---------- 第5步：健康检查 ----------
log_info "[5/6] 检查服务状态..."
HEALTH_OK=false
for i in {1..6}; do
    if $SSH_CMD "curl -sf http://localhost/ > /dev/null 2>&1"; then
        HEALTH_OK=true
        break
    fi
    log_warn "第 ${i} 次检查失败，等待 5 秒..."
    sleep 5
done

if [[ "$HEALTH_OK" == true ]]; then
    log_info "[6/6] 部署成功！"
    $SSH_CMD "docker compose -f ${DEPLOY_PATH}/docker-compose.prod.yml ps"
else
    log_error "[6/6] 服务未正常启动！正在自动回滚..."
    $SSH_CMD << 'AUTO_ROLLBACK'
        set -e
        cd /opt/schedule-system
        PREV_TAG=$(cat .previous_tag 2>/dev/null || echo "")
        if [[ -n "$PREV_TAG" ]]; then
            export IMAGE_TAG="$PREV_TAG"
            docker compose -f docker-compose.prod.yml down --timeout 10
            docker compose -f docker-compose.prod.yml up -d
            echo "[信息] 已回滚到版本: $PREV_TAG"
        else
            echo "[错误] 无法回滚（没有历史版本），请手动检查"
        fi
        docker compose -f docker-compose.prod.yml logs --tail=30
AUTO_ROLLBACK
    exit 1
fi

log_info "=============================="
log_info "  部署完成！版本: ${TAG}"
log_info "=============================="
