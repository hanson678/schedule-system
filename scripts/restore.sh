#!/usr/bin/env bash
# ============================================================
# 排期录入系统 - 数据还原脚本
# 用法:
#   ./scripts/restore.sh backup-20260228-120000.tar.gz
#   ./scripts/restore.sh --file /opt/backups/schedule-system/backup-20260228-120000.tar.gz
#   ./scripts/restore.sh --list  # 列出所有可用备份
# ============================================================
set -euo pipefail

BACKUP_DIR="/opt/backups/schedule-system"
BACKUP_FILE=""
LIST_ONLY=false

# 颜色
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

log_info()  { echo -e "${GREEN}[还原]${NC} $1"; }
log_warn()  { echo -e "${YELLOW}[警告]${NC} $1"; }
log_error() { echo -e "${RED}[错误]${NC} $1"; }

# 参数解析
while [[ $# -gt 0 ]]; do
    case $1 in
        --file)  BACKUP_FILE="$2"; shift 2 ;;
        --dir)   BACKUP_DIR="$2"; shift 2 ;;
        --list)  LIST_ONLY=true; shift ;;
        *)
            # 如果直接传文件名
            if [[ -f "$1" ]]; then
                BACKUP_FILE="$1"
            elif [[ -f "${BACKUP_DIR}/$1" ]]; then
                BACKUP_FILE="${BACKUP_DIR}/$1"
            else
                log_error "找不到备份文件: $1"
                exit 1
            fi
            shift ;;
    esac
done

# 列出备份
if [[ "$LIST_ONLY" == true ]]; then
    echo "可用的备份文件："
    echo "=============================="
    ls -lh "${BACKUP_DIR}"/backup-*.tar.gz 2>/dev/null | \
        awk '{print NR". "$NF" ("$5")"}' || \
        echo "  没有找到任何备份文件"
    exit 0
fi

# 验证
if [[ -z "$BACKUP_FILE" ]]; then
    log_error "请指定备份文件："
    echo "  ./scripts/restore.sh backup-20260228-120000.tar.gz"
    echo "  ./scripts/restore.sh --list  # 查看可用备份"
    exit 1
fi

if [[ ! -f "$BACKUP_FILE" ]]; then
    log_error "备份文件不存在: $BACKUP_FILE"
    exit 1
fi

log_info "=============================="
log_info "  开始还原"
log_info "  文件: $(basename "$BACKUP_FILE")"
log_info "=============================="

# 确认操作
echo ""
echo -e "${RED}警告：还原操作会覆盖当前数据！${NC}"
read -p "确定要继续吗？(输入 yes 确认): " CONFIRM
if [[ "$CONFIRM" != "yes" ]]; then
    log_info "已取消还原操作"
    exit 0
fi

# 创建临时目录
TEMP_DIR=$(mktemp -d)
trap "rm -rf ${TEMP_DIR}" EXIT

# ---------- 1. 解压备份 ----------
log_info "[1/4] 解压备份文件..."
tar -xzf "$BACKUP_FILE" -C "$TEMP_DIR"

# 找到解压后的目录
RESTORE_DIR=$(find "$TEMP_DIR" -maxdepth 1 -type d -name "backup-*" | head -1)
if [[ -z "$RESTORE_DIR" ]]; then
    log_error "备份文件格式不正确"
    exit 1
fi

# ---------- 2. 停止应用 ----------
log_info "[2/4] 停止应用服务..."
cd /opt/schedule-system
docker compose -f docker-compose.prod.yml stop app

# ---------- 3. 还原数据 ----------
log_info "[3/4] 还原数据..."

# 还原应用数据
if [[ -d "${RESTORE_DIR}/data" ]]; then
    log_info "  还原应用数据（JSON 配置、日志）..."
    docker run --rm \
        -v schedule-system_app-data:/target \
        -v "${RESTORE_DIR}/data":/source:ro \
        alpine sh -c "rm -rf /target/* && cp -r /source/* /target/"
    log_info "  应用数据还原完成"
fi

# 还原上传文件
if [[ -d "${RESTORE_DIR}/uploads" ]]; then
    log_info "  还原上传文件..."
    docker run --rm \
        -v schedule-system_app-uploads:/target \
        -v "${RESTORE_DIR}/uploads":/source:ro \
        alpine sh -c "rm -rf /target/* && cp -r /source/* /target/"
    log_info "  上传文件还原完成"
fi

# ---------- 4. 重启服务 ----------
log_info "[4/4] 重启应用服务..."
docker compose -f docker-compose.prod.yml up -d

# 等待健康检查
log_info "等待服务启动..."
sleep 8

if docker compose -f docker-compose.prod.yml ps | grep -q "healthy"; then
    log_info "服务已正常启动"
else
    log_warn "服务可能未完全启动，请检查："
    docker compose -f docker-compose.prod.yml logs --tail=20 app
fi

log_info "=============================="
log_info "  还原完成！"
log_info "=============================="
