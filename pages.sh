# 1. gh-pagesブランチを作成
git checkout --orphan gh-pages
git rm -rf .
echo "# Discord Exporter TUI" >README.md
mkdir scripts
# install.shを scripts/に配置
git add .
git commit -m "Add install script"
git push origin gh-pages

# 2. GitHub Pagesを有効化
# Settings > Pages > Source: Deploy from a branch > gh-pages

# 3. アクセスURL
# https://takayamaekawa.github.io/discord-exporter-tui/scripts/install.sh
