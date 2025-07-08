#!/bin/bash
set -e

# カラー出力
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# 設定
REPO_OWNER="takayamaekawa"
REPO_NAME="discord-exporter-tui"
VERSION="v1.1"
BINARY_NAME="discord_exporter"
INSTALL_DIR="/usr/local/bin"
BINARY_URL="https://github.com/${REPO_OWNER}/${REPO_NAME}/releases/download/${VERSION}/${BINARY_NAME}"
CHECKSUM_URL="https://github.com/${REPO_OWNER}/${REPO_NAME}/releases/download/${VERSION}/hashes.sha256"

# ヘルパー関数
print_status() {
  echo -e "${GREEN}[INFO]${NC} $1"
}

print_error() {
  echo -e "${RED}[ERROR]${NC} $1"
}

print_warning() {
  echo -e "${YELLOW}[WARNING]${NC} $1"
}

print_header() {
  echo -e "${BLUE}================================${NC}"
  echo -e "${BLUE}  Discord Exporter TUI Installer${NC}"
  echo -e "${BLUE}================================${NC}"
  echo ""
}

# システムチェック
check_system() {
  print_status "Checking system compatibility..."

  # Linux チェック
  if [ "$(uname -s)" != "Linux" ]; then
    print_error "This installer is only for Linux systems"
    exit 1
  fi

  # 依存関係チェック
  if ! command -v curl &>/dev/null; then
    print_error "curl is required but not installed"
    print_error "Please install curl: sudo apt install curl (Ubuntu/Debian) or sudo yum install curl (CentOS/RHEL)"
    exit 1
  fi

  if ! command -v sha256sum &>/dev/null; then
    print_error "sha256sum is required but not installed"
    exit 1
  fi

  print_status "System compatibility check passed"
}

# チェックサムの検証
verify_checksum() {
  local file=$1
  local expected_hash_file=$2

  print_status "Verifying checksum..."

  # チェックサムファイルから期待値を取得
  local expected_hash=$(grep "${BINARY_NAME}" "$expected_hash_file" | cut -d' ' -f1)

  if [ -z "$expected_hash" ]; then
    print_error "Could not find checksum for ${BINARY_NAME}"
    return 1
  fi

  # 実際のハッシュを計算
  local actual_hash=$(sha256sum "$file" | cut -d' ' -f1)

  if [ "$actual_hash" = "$expected_hash" ]; then
    print_status "✓ Checksum verified"
    return 0
  else
    print_error "✗ Checksum mismatch!"
    print_error "Expected: $expected_hash"
    print_error "Actual:   $actual_hash"
    return 1
  fi
}

# バイナリのダウンロード
download_binary() {
  print_status "Downloading discord-exporter-tui ${VERSION}..."

  # 一時ディレクトリを作成
  local tmp_dir=$(mktemp -d)
  local binary_path="$tmp_dir/$BINARY_NAME"
  local checksum_path="$tmp_dir/hashes.sha256"

  # バイナリをダウンロード
  if ! curl -fsSL "$BINARY_URL" -o "$binary_path"; then
    print_error "Failed to download binary"
    rm -rf "$tmp_dir"
    exit 1
  fi

  # チェックサムをダウンロード
  if ! curl -fsSL "$CHECKSUM_URL" -o "$checksum_path"; then
    print_error "Failed to download checksum file"
    rm -rf "$tmp_dir"
    exit 1
  fi

  # チェックサムを検証
  if ! verify_checksum "$binary_path" "$checksum_path"; then
    print_error "Checksum verification failed"
    rm -rf "$tmp_dir"
    exit 1
  fi

  # 実行権限を付与
  chmod +x "$binary_path"

  # パスを標準出力に出力（print_statusを使わない）
  printf "%s" "$binary_path"
}

# インストール
install_binary() {
  local binary_path=$1

  print_status "Installing discord-exporter-tui to $INSTALL_DIR..."

  # インストールディレクトリが存在しない場合は作成
  if [ ! -d "$INSTALL_DIR" ]; then
    print_status "Creating $INSTALL_DIR directory..."
    if ! sudo mkdir -p "$INSTALL_DIR"; then
      print_error "Failed to create $INSTALL_DIR"
      exit 1
    fi
  fi

  # 既存のバイナリをバックアップ
  if [ -f "$INSTALL_DIR/$BINARY_NAME" ]; then
    print_warning "Existing installation found, creating backup..."
    sudo mv "$INSTALL_DIR/$BINARY_NAME" "$INSTALL_DIR/$BINARY_NAME.backup"
  fi

  # バイナリをインストール
  if [ -w "$INSTALL_DIR" ]; then
    mv "$binary_path" "$INSTALL_DIR/$BINARY_NAME"
  else
    if ! sudo mv "$binary_path" "$INSTALL_DIR/$BINARY_NAME"; then
      print_error "Failed to install binary"
      exit 1
    fi
  fi

  print_status "✓ Binary installed successfully"
}

# 環境変数の設定
setup_environment() {
  print_status "Setting up environment..."

  # PATHに既に含まれているかチェック
  if echo "$PATH" | grep -q "$INSTALL_DIR"; then
    print_status "✓ $INSTALL_DIR is already in PATH"
    return 0
  fi

  # シェルプロファイルを判定
  local shell_profile=""
  if [ -n "$BASH_VERSION" ]; then
    shell_profile="$HOME/.bashrc"
  elif [ -n "$ZSH_VERSION" ]; then
    shell_profile="$HOME/.zshrc"
  elif [ -f "$HOME/.bash_profile" ]; then
    shell_profile="$HOME/.bash_profile"
  else
    shell_profile="$HOME/.profile"
  fi

  # PATHに追加
  print_status "Adding $INSTALL_DIR to PATH in $shell_profile"
  echo "" >>"$shell_profile"
  echo "# Discord Exporter TUI" >>"$shell_profile"
  echo "export PATH=\"$INSTALL_DIR:\$PATH\"" >>"$shell_profile"

  print_warning "Please run 'source $shell_profile' or restart your terminal to update PATH"
}

# バージョン確認
verify_installation() {
  print_status "Verifying installation..."

  if [ -f "$INSTALL_DIR/$BINARY_NAME" ]; then
    print_status "✓ Installation verified"
    return 0
  else
    print_error "Installation verification failed"
    return 1
  fi
}

# 使用方法の表示
show_usage() {
  echo ""
  echo -e "${BLUE}Usage:${NC}"
  echo "  $BINARY_NAME [options]"
  echo ""
  echo -e "${BLUE}Examples:${NC}"
  echo "  $BINARY_NAME --help              # Show help"
  echo "  $BINARY_NAME --version           # Show version"
  echo ""
  echo -e "${BLUE}For more information:${NC}"
  echo "  https://github.com/${REPO_OWNER}/${REPO_NAME}"
}

# クリーンアップ
cleanup() {
  if [ -n "$tmp_dir" ] && [ -d "$tmp_dir" ]; then
    rm -rf "$tmp_dir"
  fi
}

# エラーハンドリング
trap cleanup EXIT

# メイン処理
main() {
  print_header

  check_system

  # バイナリをダウンロード
  local binary_path
  binary_path=$(download_binary)

  # ダウンロードが成功したかチェック
  if [ ! -f "$binary_path" ]; then
    print_error "Downloaded binary not found: $binary_path"
    exit 1
  fi

  install_binary "$binary_path"
  setup_environment

  if verify_installation; then
    echo ""
    echo -e "${GREEN}🎉 Installation completed successfully!${NC}"
    echo ""
    echo -e "${BLUE}Next steps:${NC}"
    echo "1. Restart your terminal or run: source ~/.bashrc (or ~/.zshrc)"
    echo "2. Run: $BINARY_NAME --help"
    echo ""
    show_usage
  else
    print_error "Installation failed"
    exit 1
  fi
}

# ヘルプメッセージ
if [ "$1" = "--help" ] || [ "$1" = "-h" ]; then
  print_header
  echo "Discord Exporter TUI Installer"
  echo ""
  echo "This script will:"
  echo "  - Download the latest discord-exporter-tui binary"
  echo "  - Verify the checksum"
  echo "  - Install to $INSTALL_DIR"
  echo "  - Setup PATH environment variable"
  echo ""
  echo "Usage: $0 [--help]"
  echo ""
  echo "Requirements:"
  echo "  - Linux system"
  echo "  - curl"
  echo "  - sha256sum"
  echo ""
  exit 0
fi

# 実行
main "$@"
