#!/usr/bin/env bash
# ============================================================
# 排期录入系统 - 数据备份脚本
# 用法:
#   ./scripts/backup.sh                    # 备份到本地 /opt/backups/
#   ./scripts/backup.sh --scp user@host:/path  # 备份后 SCP 到远端
# ============================================================
set -euo pipefail

# 配置
BACKUP_DIR="/opt/backups/schedule-system"
COMPOSE_FILE="/opt/schedule-system/docker-compose.prod.yml"
TIMESTAMP="$(date +%Y%m%d-%H%M%S)"
BACKUP_NAME="backup-${TIMESTAMP}"
BACKUP_PATH="${BACKUP_DIR}/${BACKUP_NAME}"
SCP_TARGET=""
KEEP_DAYS=30  # 保留最近30天的备份

# 颜色
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

log_info()  { echo -e "${GREEN}[备份]${NC} $1"; }
log_warn()  { echo -e "${YELLOW}[警告]${NC} $1"; }
log_error() { echo -e "${RED}[错误]${NC} $1"; }

# 参数解析
while [[ $# -gt 0 ]]; do
    case $1 in
        --scp)     SCP_TARGET="$2"; shift 2 ;;
        --dir)     BACKUP_DIR="$2"; shift 2 ;;
        --keep)    KEEP_DAYS="$2"; shift 2 ;;
        *)         log_error "未知参数: $1"; exit 1 ;;
    esac
done

BACKUP_PATH="${BACKUP_DIR}/${BACKUP_NAME}"

log_info "=============================="
log_info "  开始备份 - ${TIMESTAMP}"
log_info "=============================="

# 创建备份目录
mkdir -p "${BACKUP_PATH}"

# ---------- 1. 备份应用数据（JSON 配置、日志、SKU映射）----------
log_info "[1/4] 备份应用数据..."
docker cp schedule-app:/app/data "${BACKUP_PATH}/data" 2>/dev/null || {
    log_warn "无法从容器复制数据目录，尝试从 volume 复制..."
    # 通过临时容器复制 volume 数据
    docker run --rm \
        -v schedule-system_app-data:/source:ro \
        -v "${BACKUP_PATH}:/backup" \
        alpine sh -c "cp -r /source /backup/data"
}

# ---------- 2. 备份上传的 PDF 文件 ----------
log_info "[2/4] 备份上传文件..."
docker run --rm \
    -v schedule-system_app-uploads:/source:ro \
    -v "${BACKUP_PATH}:/backup" \
    alpine sh -c "cp -r /source /backup/uploads" 2>/dev/null || \
    log_warn "上传目录为空或不存在，跳过"

# ---------- 3. 备份排期文件快照（可选，文件可能很大）----------
log_info "[3/4] 备份排期文件列表（不备份文件本身，太大）..."
docker run --rm \
    -v schedule-system_schedules:/source:ro \
    -v "${BACKUP_PATH}:/backup" \
    alpine sh -c "ls -la /source/ > /backup/schedules-filelist.txt 2>/dev/null" || \
    log_warn "排期目录不可访问，跳过"

# ---------- 4. 压缩 ----------
log_info "[4/4] 压缩备份文件..."
cd "${BACKUP_DIR}"
tar -czf "${BACKUP_NAME}.tar.gz" "${BACKUP_NAME}"
rm -rf "${BACKUP_PATH}"

BACKUP_FILE="${BACKUP_DIR}/${BACKUP_NAME}.tar.gz"
BACKUP_SIZE=$(du -sh "${BACKUP_FILE}" | cut -f1)
log_info "备份文件: ${BACKUP_FILE} (${BACKUP_SIZE})"

# ---------- 可选：SCP 到远端 ----------
if [[ -n "$SCP_TARGET" ]]; then
    log_info "传输备份到: ${SCP_TARGET}"
    scp "${BACKUP_FILE}" "${SCP_TARGET}/"
    log_info "远端传输完成"
fi

# ---------- 清理旧备份 ----------
log_info "清理 ${KEEP_DAYS} 天前的旧备份..."
find "${BACKUP_DIR}" -name "backup-*.tar.gz" -mtime +${KEEP_DAYS} -delete 2>/dev/null || true
REMAINING=$(ls -1 "${BACKUP_DIR}"/backup-*.tar.gz 2>/dev/null | wc -l)
log_info "当前保留 ${REMAINING} 个备份文件"

log_info "=============================="
log_info "  备份完成！"
log_info "=============================="
