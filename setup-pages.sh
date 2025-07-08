#!/bin/bash
set -e

# カラー出力
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

print_status() {
  echo -e "${GREEN}[INFO]${NC} $1"
}

print_step() {
  echo -e "${BLUE}[STEP]${NC} $1"
}

print_warning() {
  echo -e "${YELLOW}[WARNING]${NC} $1"
}

# 現在のブランチを保存
CURRENT_BRANCH=$(git branch --show-current)

print_step "Setting up GitHub Pages for discord-exporter-tui"
echo ""

# 1. install.shをscriptsディレクトリに移動（既に存在する場合）
if [ -f "install.sh" ]; then
  print_status "Moving install.sh to scripts/ directory"
  mv install.sh scripts/
else
  print_warning "install.sh not found in root directory"
fi

# 2. scriptsディレクトリの確認と作成
if [ ! -d "scripts" ]; then
  print_status "Creating scripts/ directory"
  mkdir -p scripts
fi

# install.shの確認
if [ ! -f "scripts/install.sh" ]; then
  if [ -f "install.sh" ]; then
    print_status "Moving install.sh to scripts/ directory"
    mv install.sh scripts/
  else
    print_status "Creating placeholder install.sh in scripts/ directory"
    echo "Please add your install.sh to scripts/ directory and run this script again"
    exit 1
  fi
fi

# 3. gh-pagesブランチが既に存在するかチェック
if git show-ref --verify --quiet refs/heads/gh-pages; then
  print_warning "gh-pages branch already exists"
  read -p "Do you want to recreate it? (y/N): " -n 1 -r
  echo
  if [[ $REPLY =~ ^[Yy]$ ]]; then
    print_status "Deleting existing gh-pages branch"
    git branch -D gh-pages 2>/dev/null || true
    git push origin --delete gh-pages 2>/dev/null || true
  else
    print_status "Switching to existing gh-pages branch"
    git checkout gh-pages
    # scriptsディレクトリとinstall.shをコピー
    git checkout $CURRENT_BRANCH -- scripts/
    git add scripts/
    git commit -m "Update install script" || print_warning "No changes to commit"
    git push origin gh-pages
    git checkout $CURRENT_BRANCH
    print_status "✓ Updated existing gh-pages branch"
    echo ""
    echo "GitHub Pages URL: https://takayamaekawa.github.io/discord-exporter-tui/scripts/install.sh"
    echo "Install command: curl -fsSL https://takayamaekawa.github.io/discord-exporter-tui/scripts/install.sh | bash"
    exit 0
  fi
fi

# 4. 変更をコミット（まだの場合）
if ! git diff --cached --quiet 2>/dev/null; then
  print_status "Committing current changes to $CURRENT_BRANCH"
  git add .
  git commit -m "Add install script to scripts directory"
  git push origin $CURRENT_BRANCH
elif ! git diff --quiet scripts/ 2>/dev/null; then
  print_status "Committing scripts changes to $CURRENT_BRANCH"
  git add scripts/
  git commit -m "Update install script"
  git push origin $CURRENT_BRANCH
fi

# 5. gh-pagesブランチを作成
print_step "Creating gh-pages branch"

# orphanブランチを作成
git checkout --orphan gh-pages

# 不要なファイルを削除（.gitignoreに基づいて）
print_status "Cleaning up gh-pages branch"
git rm -rf . 2>/dev/null || true

# 必要なファイルのみを復元
if git ls-tree $CURRENT_BRANCH scripts/ >/dev/null 2>&1; then
  git checkout $CURRENT_BRANCH -- scripts/
else
  print_warning "scripts/ directory not tracked in git on $CURRENT_BRANCH"
  print_status "Creating scripts/ directory and adding install.sh"
  mkdir -p scripts
  git checkout $CURRENT_BRANCH -- install.sh 2>/dev/null && mv install.sh scripts/ || true
fi

# GitHub Pages用のindex.htmlを作成
print_status "Creating index.html for GitHub Pages"
cat >index.html <<'EOF'
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Discord Exporter TUI - Installation</title>
    <meta name="description" content="Easy installation for Discord Exporter TUI - A terminal-based Discord chat exporter">
    <script
        async
        src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client=ca-pub-2877828132102103"
        crossorigin="anonymous">
    </script>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 2rem;
            line-height: 1.6;
            color: #333;
            background: #fff;
        }
        .header {
            text-align: center;
            margin-bottom: 3rem;
            padding: 2rem 0;
            border-bottom: 1px solid #eee;
        }
        .header h1 {
            color: #1a1a1a;
            margin-bottom: 0.5rem;
        }
        .header p {
            color: #666;
            font-size: 1.1rem;
        }
        .install-section {
            background: #f8f9fa;
            border-radius: 12px;
            padding: 2rem;
            margin: 2rem 0;
            border: 1px solid #e9ecef;
        }
        .install-section h2 {
            color: #343a40;
            margin-top: 0;
        }
        .code {
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 1.5rem;
            border-radius: 8px;
            font-family: 'Monaco', 'Consolas', 'SF Mono', monospace;
            overflow-x: auto;
            margin: 1rem 0;
            border: 1px solid #333;
        }
        .button {
            display: inline-block;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 0.75rem 1.5rem;
            text-decoration: none;
            border-radius: 8px;
            margin: 0.5rem;
            transition: transform 0.2s;
            font-weight: 500;
        }
        .button:hover {
            transform: translateY(-2px);
        }
        .requirements {
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 8px;
            padding: 1.5rem;
        }
        .ad-container {
            margin: 2rem 0;
            text-align: center;
            padding: 1rem;
            background: #f8f9fa;
            border-radius: 8px;
            border: 1px solid #dee2e6;
        }
        .footer {
            text-align: center;
            margin-top: 3rem;
            padding-top: 2rem;
            border-top: 1px solid #eee;
            color: #666;
        }
    </style>
</head>
<body>
    <div class="header">
        <h1>🚀 Discord Exporter TUI</h1>
        <p>A powerful terminal-based Discord chat exporter</p>
        <a href="https://github.com/takayamaekawa/discord-exporter-tui" class="button">📁 View on GitHub</a>
        <a href="https://github.com/takayamaekawa/discord-exporter-tui/releases" class="button">📦 Releases</a>
    </div>

    <div class="ad-container">
        <!-- AdSense Ad Unit -->
        <ins class="adsbygoogle"
             style="display:block"
             data-ad-client="ca-pub-2877828132102103"
             data-ad-slot="auto"
             data-ad-format="auto"
             data-full-width-responsive="true"></ins>
        <script>
             (adsbygoogle = window.adsbygoogle || []).push({});
        </script>
    </div>

    <div class="install-section">
        <h2>⚡ Quick Install (Linux)</h2>
        <p>Run this one-liner command in your terminal:</p>
        <div class="code">
curl -fsSL https://provider.maekawa.dev/scripts/install.sh | bash
        </div>
        <p><small>This script will automatically download, verify, and install the latest version.</small></p>
    </div>

    <div class="install-section">
        <h2>🔍 Manual Installation</h2>
        <p>If you prefer to review the script before running it:</p>
        <div class="code">
# Download and review the script
curl -fsSL https://provider.maekawa.dev/scripts/install.sh -o install.sh

# Make it executable and review
chmod +x install.sh
cat install.sh

# Run the script
./install.sh
        </div>
    </div>

    <div class="ad-container">
        <!-- Second AdSense Ad Unit -->
        <ins class="adsbygoogle"
             style="display:block"
             data-ad-client="ca-pub-2877828132102103"
             data-ad-slot="auto"
             data-ad-format="auto"
             data-full-width-responsive="true"></ins>
        <script>
             (adsbygoogle = window.adsbygoogle || []).push({});
        </script>
    </div>

    <div class="requirements">
        <h2>📋 Requirements</h2>
        <ul>
            <li>🐧 Linux system (x86_64 or ARM64)</li>
            <li>📡 curl (for downloading)</li>
            <li>🔐 sha256sum (for verification)</li>
            <li>👤 sudo access (for installation to /usr/local/bin)</li>
        </ul>
    </div>

    <div class="install-section">
        <h2>🛠️ Features</h2>
        <ul>
            <li>Export Discord channels to various formats</li>
            <li>Terminal-based user interface</li>
            <li>Secure token-based authentication</li>
            <li>Multiple export formats supported</li>
            <li>Fast and lightweight</li>
        </ul>
    </div>

    <div class="footer">
        <p>Made with ❤️ by <a href="https://github.com/takayamaekawa">takayamaekawa</a></p>
        <p>Licensed under MIT License</p>
    </div>
</body>
</html>
EOF

# CNAMEファイルを作成
print_status "Creating CNAME file for custom domain"
echo "provider.maekawa.dev" >CNAME

# README.mdを作成
print_status "Creating README.md for GitHub Pages"
cat >README.md <<'EOF'
# Discord Exporter TUI - Installation

This branch contains the installation scripts and GitHub Pages content for Discord Exporter TUI.

## Quick Install

```bash
curl -fsSL https://provider.maekawa.dev/scripts/install.sh | bash
```

## Manual Install

```bash
# Download the script
curl -fsSL https://provider.maekawa.dev/scripts/install.sh -o install.sh

# Review the script
cat install.sh

# Run the script
bash install.sh
```

## Files

- `scripts/install.sh` - Installation script
- `index.html` - GitHub Pages landing page

## Main Repository

[https://github.com/takayamaekawa/discord-exporter-tui](https://github.com/takayamaekawa/discord-exporter-tui)
EOF

# 6. コミットしてプッシュ
print_status "Committing and pushing gh-pages branch"
git add .
git commit -m "Initial GitHub Pages setup with install script"
git push origin gh-pages

# 7. 元のブランチに戻る
print_status "Switching back to $CURRENT_BRANCH branch"
git checkout $CURRENT_BRANCH

print_step "✓ GitHub Pages setup complete!"
echo ""
echo -e "${BLUE}Next steps:${NC}"
echo "1. Go to your GitHub repository settings"
echo "2. Navigate to Settings > Pages"
echo "3. Set Source to: Deploy from a branch"
echo "4. Select branch: gh-pages"
echo "5. Select folder: / (root)"
echo ""
echo -e "${BLUE}URLs (will be available after DNS propagation and GitHub Pages deployment):${NC}"
echo "Install script: https://provider.maekawa.dev/scripts/install.sh"
echo "Landing page:   https://provider.maekawa.dev/"
echo ""
echo -e "${BLUE}Install command for users:${NC}"
echo "curl -fsSL https://provider.maekawa.dev/scripts/install.sh | bash"
echo ""
echo -e "${YELLOW}Note:${NC} It may take a few minutes for GitHub Pages to deploy."
