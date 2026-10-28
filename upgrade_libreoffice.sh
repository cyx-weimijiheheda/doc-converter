#!/usr/bin/env bash
# ============================================================================
# upgrade_libreoffice.sh —— 把 LibreOffice 升级到 26.8+，解决老式
#                            Microsoft Equation 3.0 公式对象导致的转换崩溃
#                            （exit=134 / stl_vector 断言）。
#
# 不使用 flatpak！升级路径（按顺序尝试）：
#   1) dnf 仓库升级——若仓库已提供 26.8+，最省事。
#   2) 仓库版本不够时，下载官方 RPM 包 → 本地提取 → 在 /usr/local/bin
#      生成 soffice / libreoffice 包装脚本（有 sudo 时）；无 sudo 则装到
#      ~/.local/bin。已提取过的会直接复用，不重复下载。
#
# 用法：
#   bash upgrade_libreoffice.sh               # 升级 + 清理冗余下载包
#   bash upgrade_libreoffice.sh --no-clean    # 只升级，保留下载包
#   bash upgrade_libreoffice.sh --clean-only  # 只清理冗余下载包
#
# 可用环境变量覆盖：
#   MIN_VERSION=26.8  LO_VERSION=26.8.1  PORTABLE_DIR=/home/wmjhhd/lo_portable
#   MIRROR=<rpm 目录地址>
# ============================================================================
set -euo pipefail

MIN_VERSION="${MIN_VERSION:-26.8}"
LO_VERSION="${LO_VERSION:-26.8.1}"
PORTABLE_DIR="${PORTABLE_DIR:-/home/wmjhhd/lo_portable}"
MIRROR="${MIRROR:-https://mirrors.tuna.tsinghua.edu.cn/libreoffice/libreoffice/stable/${LO_VERSION}/rpm/x86_64}"
TARBALL="LibreOffice_${LO_VERSION}_Linux_x86-64_rpm.tar.gz"

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[0;33m'; NC='\033[0m'
info(){ echo -e "${GREEN}[+]${NC} $*"; }
warn(){ echo -e "${YELLOW}[!]${NC} $*"; }
fail(){ echo -e "${RED}[x]${NC} $*"; }

do_upgrade=true; do_clean=true
for a in "$@"; do
  case "$a" in
    --no-clean)   do_clean=false ;;
    --clean-only) do_upgrade=false ;;
    -h|--help)
      echo "用法: bash $(basename "$0") [--no-clean|--clean-only]"
      echo "  默认       : 升级 + 清理冗余下载包"
      echo "  --no-clean : 只升级，保留下载包"
      echo "  --clean-only: 只清理冗余下载包，不升级"
      exit 0 ;;
  esac
done

# 版本比较：把 "26.8.1.1" 归并成 "26.8"，用 sort -V 判定 $1 >= $2
version_ge(){
  local a b
  a=$(echo "$1" | grep -oE '[0-9]+\.[0-9]+' | head -1)
  b="$2"
  [[ "$(printf '%s\n%s\n' "$b" "$a" | sort -V | head -1)" == "$b" ]]
}

# 当前 PATH 上 soffice 的版本号
path_soffice_version(){
  local b
  b=$(command -v soffice 2>/dev/null) || { echo "0.0.0"; return; }
  "$b" --version 2>/dev/null | grep -oE '[0-9]+\.[0-9]+\.[0-9]+' | head -1 || echo "0.0.0"
}

# 若已提取过官方二进制，返回其路径
portable_bin(){
  local p
  for p in "$PORTABLE_DIR"/root/opt/libreoffice*/program/soffice; do
    [ -x "$p" ] && { echo "$p"; return 0; }
  done
  return 1
}

# 是否可提权
SUDO=""
if [ "$(id -u)" -ne 0 ] && command -v sudo >/dev/null 2>&1; then SUDO="sudo"; fi

bin_version(){ "$1" --version 2>/dev/null | grep -oE '[0-9]+\.[0-9]+\.[0-9]+' | head -1; }

# 在 PATH 靠前的目录生成 soffice/libreoffice 包装脚本
install_wrapper(){
  local bin="$1" ver dir
  ver=$(bin_version "$bin")
  if [ "$(id -u)" -eq 0 ] || [ -n "$SUDO" ]; then
    dir="/usr/local/bin"
    $SUDO mkdir -p "$dir"
    printf '#!/bin/sh\nexec "%s" "$@"\n' "$bin" | $SUDO tee "$dir/soffice" >/dev/null
    $SUDO chmod +x "$dir/soffice"
    printf '#!/bin/sh\nexec "%s" "$@"\n' "$bin" | $SUDO tee "$dir/libreoffice" >/dev/null
    $SUDO chmod +x "$dir/libreoffice"
    info "已生成包装脚本：$dir/soffice -> $bin （版本 $ver）"
  else
    dir="$HOME/.local/bin"; mkdir -p "$dir"
    printf '#!/bin/sh\nexec "%s" "$@"\n' "$bin" > "$dir/soffice"; chmod +x "$dir/soffice"
    printf '#!/bin/sh\nexec "%s" "$@"\n' "$bin" > "$dir/libreoffice"; chmod +x "$dir/libreoffice"
    info "已生成包装脚本：$dir/soffice （版本 $ver）"
    case ":$PATH:" in
      *":$dir:"*) ;;
      *) warn "请把 $dir 加入 PATH：echo 'export PATH=\"\$HOME/.local/bin:\$PATH\"' >> ~/.bashrc && source ~/.bashrc" ;;
    esac
  fi
}

ensure_tools(){
  local ok=0
  command -v curl     >/dev/null 2>&1 || { fail "缺少 curl"; ok=1; }
  command -v rpm2cpio >/dev/null 2>&1 || { fail "缺少 rpm2cpio（dnf install rpm）"; ok=1; }
  command -v cpio     >/dev/null 2>&1 || { fail "缺少 cpio（dnf install cpio）"; ok=1; }
  return $ok
}

# 官方 RPM 方案（不使用 flatpak）
install_official_rpm(){
  local bin v f rpmdir
  if bin=$(portable_bin) && v=$(bin_version "$bin") && version_ge "$v" "$MIN_VERSION"; then
    info "发现已提取好的官方版本 $v：$bin（跳过下载）"
    install_wrapper "$bin"
    return 0
  fi

  ensure_tools || return 1
  mkdir -p "$PORTABLE_DIR"
  local tgz="$PORTABLE_DIR/$TARBALL"
  if [ -s "$tgz" ]; then
    info "复用已有下载包：$tgz"
  else
    info "下载 $TARBALL（约 250MB）..."
    curl -fL --retry 3 -o "$tgz" "$MIRROR/$TARBALL" || { fail "下载失败：$MIRROR/$TARBALL"; return 1; }
  fi

  info "解压 tar 包 ..."
  rm -rf "$PORTABLE_DIR/src" "$PORTABLE_DIR/root"
  mkdir -p "$PORTABLE_DIR/src"
  tar xzf "$tgz" -C "$PORTABLE_DIR/src" 2>/dev/null || true
  rpmdir=$(find "$PORTABLE_DIR/src" -maxdepth 3 -type d -name RPMS | head -1)
  [ -n "$rpmdir" ] || { fail "解压后未找到 RPMS 目录"; return 1; }

  info "提取 RPM 包（rpm2cpio）..."
  mkdir -p "$PORTABLE_DIR/root"
  for f in "$rpmdir"/*.rpm; do
    ( cd "$PORTABLE_DIR/root" && rpm2cpio "$f" | cpio -idmu --quiet )
  done

  bin=$(portable_bin) || { fail "提取后未找到 soffice"; return 1; }
  install_wrapper "$bin"
}

# ===========================================================================
# 1) 升级
# ===========================================================================
if $do_upgrade; then
  CUR=$(path_soffice_version)
  info "当前 soffice 版本：$CUR"

  if version_ge "$CUR" "$MIN_VERSION"; then
    info "已满足 $MIN_VERSION+，无需升级。"
  else
    if command -v dnf >/dev/null 2>&1; then
      info "先尝试 dnf 仓库升级（若仓库已有 26.8+ 则最省事）..."
      $SUDO dnf install -y libreoffice libreoffice-calc libreoffice-impress libreoffice-writer 2>/dev/null \
        || warn "dnf 仓库安装未成功（仓库可能仍是 26.2，属正常）"
      CUR=$(path_soffice_version)
    fi

    if version_ge "$CUR" "$MIN_VERSION"; then
      info "dnf 升级成功，版本 $CUR"
    else
      warn "仓库版本 $CUR 低于 $MIN_VERSION，改用官方 RPM 包安装（不使用 flatpak）"
      install_official_rpm
    fi
  fi

  NEW=$(path_soffice_version)
  if version_ge "$NEW" "$MIN_VERSION"; then
    info "升级完成，最终 soffice 版本：$NEW ✓"
  else
    warn "当前 PATH 上版本仍为 $NEW；若包装脚本装在 ~/.local/bin，请确认它在 PATH 中后重新打开终端。"
  fi
fi

# ===========================================================================
# 2) 清理冗余下载包（保留可用的 root/，因为包装脚本可能指向它）
# ===========================================================================
if $do_clean; then
  if [ -d "$PORTABLE_DIR" ]; then
    artifacts=()
    [ -d "$PORTABLE_DIR/src" ] && artifacts+=("$PORTABLE_DIR/src")
    for t in "$PORTABLE_DIR"/*.tar.gz; do [ -f "$t" ] && artifacts+=("$t"); done

    if [ ${#artifacts[@]} -eq 0 ]; then
      info "没有冗余下载包需要清理（当前 $PORTABLE_DIR：$(du -sh "$PORTABLE_DIR" 2>/dev/null | cut -f1)）"
    else
      warn "将删除以下冗余文件（保留可用的 root/）："
      printf '    %s\n' "${artifacts[@]}"
      if [ -t 0 ]; then
        read -r -p "确认删除？[y/N] " ans
        if [[ "${ans:-}" =~ ^[Yy]$ ]]; then
          rm -rf "${artifacts[@]}"
          info "已清理，当前占用：$(du -sh "$PORTABLE_DIR" 2>/dev/null | cut -f1)"
        else
          info "已取消删除。"
        fi
      else
        warn "非交互环境，跳过删除。如需清理请手动执行："
        printf '    rm -rf %s\n' "${artifacts[@]}"
      fi
    fi
  else
    info "便携目录不存在：$PORTABLE_DIR，无需清理。"
  fi
fi

echo
info "可执行以下命令自检："
echo "    soffice --version"
echo "    python office2pdf.py --check"
