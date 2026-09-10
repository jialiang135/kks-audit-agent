#!/usr/bin/env bash
# KKS 审核智能体 - 安全部署脚本
#
# 用法：bash deploy.sh /tmp/kks-audit-agent-deploy-YYYYMMDD.tar.gz
#
# 为什么需要它：AI 配置有两处持久化位置，都不在部署包内——
#   1) .env                       （隐藏文件，docker compose 自动插值注入）
#   2) config/app_config.json     （Web「AI 配置」页保存）
# 直接用 `rm -rf *` 再解压会把第 2 个删掉，导致"AI 配置又没了"。
# 本脚本用"覆盖解压 + 备份恢复"代替 `rm -rf *`，配置不会丢。
set -euo pipefail

PKG="${1:-}"
APP_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [ -z "$PKG" ] || [ ! -f "$PKG" ]; then
  echo "用法: bash deploy.sh <部署包.tar.gz>"
  echo "例:   bash deploy.sh /tmp/kks-audit-agent-deploy-20260910.tar.gz"
  exit 1
fi

cd "$APP_DIR"
echo "==> 应用目录: $APP_DIR"

# 1. 备份用户配置（这两个文件不在包内，重新部署必须保留）
BACKUP="$(mktemp -d)"
for f in .env config/app_config.json; do
  if [ -f "$f" ]; then
    mkdir -p "$BACKUP/$(dirname "$f")"
    cp -a "$f" "$BACKUP/$f"
    echo "    备份 $f"
  fi
done

# 2. 覆盖解压（不使用 rm -rf *，避免误删配置与 runs 数据）
echo "==> 解压 $PKG"
tar -xzf "$PKG" -C "$APP_DIR"

# 3. 恢复配置
for f in .env config/app_config.json; do
  if [ -f "$BACKUP/$f" ]; then
    mkdir -p "$(dirname "$f")"
    cp -a "$BACKUP/$f" "$f"
    echo "    恢复 $f"
  fi
done
rm -rf "$BACKUP"

# 4. 清理 Python 字节码缓存，避免旧代码残留
find "$APP_DIR" -name '__pycache__' -type d -prune -exec rm -rf {} + 2>/dev/null || true

# 5. 重建并启动
echo "==> docker compose up -d --build"
docker compose up -d --build

echo "==> 完成。状态: docker compose ps   ｜ 日志: docker compose logs -f"
