#!/usr/bin/env bash
# sales-dashboard-builder 一键安装脚本
# 用法（安装到默认 .user_skills）：
#   git clone --depth 1 https://github.com/wwyets-cmd/sales-dashboard-builder /tmp/sdb && /tmp/sdb/install.sh
# 也可手动指定目标目录：
#   ./install.sh /path/to/workspace/.user_skills
set -e

NAME="sales-dashboard-builder"
SRC="$(cd "$(dirname "$0")" && pwd)"
TARGET="${1:-}"

if [ -z "$TARGET" ]; then
  for p in "$HOME/.doubao/agent_mode/workspace/.user_skills"; do
    if [ -d "$p" ]; then TARGET="$p"; break; fi
  done
fi
if [ -z "$TARGET" ]; then
  TARGET="$(find "$HOME" -type d -path '*/workspace/.user_skills' 2>/dev/null | head -n1 || true)"
fi
if [ -z "$TARGET" ]; then
  echo "未找到 workspace/.user_skills 目录。"
  echo "请手动传入路径: ./install.sh /path/to/workspace/.user_skills"
  exit 1
fi

DEST="$TARGET/$NAME"
mkdir -p "$DEST"
if command -v rsync >/dev/null 2>&1; then
  rsync -a --exclude '.git' --exclude 'install.sh' "$SRC"/ "$DEST"/
else
  cp -R "$SRC"/. "$DEST"/
  rm -rf "$DEST/.git" "$DEST/install.sh"
fi

echo "✅ 已安装到: $DEST"
echo "重新加载后即可使用技能 $NAME。"
