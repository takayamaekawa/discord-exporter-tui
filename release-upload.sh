#!/bin/bash
# release-upload.sh

REPO_OWNER="takayamaekawa"
REPO_NAME="discord-exporter-tui"
TAG_NAME="v1.1"
GITHUB_TOKEN="${discord_exporter_tui_token}" # Personal Access Token

# リリースを作成
create_release() {
  curl -X POST \
    -H "Authorization: token $GITHUB_TOKEN" \
    -H "Content-Type: application/json" \
    -d '{
        "tag_name": "'$TAG_NAME'",
        "name": "'$TAG_NAME'",
        "body": "Release notes",
        "draft": false,
        "prerelease": false
      }' \
    "https://api.github.com/repos/$REPO_OWNER/$REPO_NAME/releases"
}

# アセットをアップロード
upload_asset() {
  local file=$1
  local filename=$(basename "$file")

  # リリースIDを取得
  local release_id=$(curl -s \
    -H "Authorization: token $GITHUB_TOKEN" \
    "https://api.github.com/repos/$REPO_OWNER/$REPO_NAME/releases/tags/$TAG_NAME" |
    grep '"id":' | head -1 | cut -d':' -f2 | cut -d',' -f1 | tr -d ' ')

  # アセットをアップロード
  curl -X POST \
    -H "Authorization: token $GITHUB_TOKEN" \
    -H "Content-Type: application/octet-stream" \
    --data-binary @"$file" \
    "https://uploads.github.com/repos/$REPO_OWNER/$REPO_NAME/releases/$release_id/assets?name=$filename"
}

# 使用例
create_release
upload_asset "./dist/discord_exporter"
upload_asset "./dist/hashes.sha256"
