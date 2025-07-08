# Discord Exporter TUI Makefile
# プラットフォーム判別型インストーラー

# 設定
REPO_OWNER := takayamaekawa
REPO_NAME := discord-exporter-tui
VERSION := v1.1
BINARY_NAME := discord_exporter
INSTALL_DIR := /usr/local/bin
PYTHON := python3
VENV_DIR := venv

# プラットフォーム判別
UNAME_S := $(shell uname -s)
ifeq ($(UNAME_S),Linux)
	PLATFORM := linux
else ifeq ($(UNAME_S),Darwin)
	PLATFORM := darwin
else ifeq ($(findstring CYGWIN,$(UNAME_S)),CYGWIN)
	PLATFORM := windows
else ifeq ($(findstring MINGW,$(UNAME_S)),MINGW)
	PLATFORM := windows
else
	PLATFORM := unknown
endif

# デフォルトターゲット
.PHONY: help
help:
	@echo "Discord Exporter TUI Makefile"
	@echo ""
	@echo "Available targets:"
	@echo "  install      - Install discord-exporter-tui"
	@echo "  build        - Build binary using PyInstaller"
	@echo "  clean        - Clean build artifacts"
	@echo "  test         - Test the installation"
	@echo "  help         - Show this help"
	@echo ""
	@echo "Platform: $(PLATFORM)"

# インストール（プラットフォーム判別）
.PHONY: install
install:
ifeq ($(PLATFORM),linux)
	@echo "Installing prebuilt binary for Linux..."
	@$(MAKE) install-binary
else
	@echo "Building from source for $(PLATFORM)..."
	@$(MAKE) install-build
endif

# Linuxバイナリインストール
.PHONY: install-binary
install-binary:
	@echo "Downloading prebuilt binary from GitHub Releases..."
	@if ! command -v curl >/dev/null 2>&1; then \
		echo "Error: curl is required but not installed"; \
		echo "Please install curl: sudo apt install curl (Ubuntu/Debian) or sudo yum install curl (CentOS/RHEL)"; \
		exit 1; \
	fi
	@if ! command -v sha256sum >/dev/null 2>&1; then \
		echo "Error: sha256sum is required but not installed"; \
		exit 1; \
	fi
	@tmp_dir=$$(mktemp -d); \
	binary_path="$$tmp_dir/$(BINARY_NAME)"; \
	checksum_path="$$tmp_dir/hashes.sha256"; \
	echo "Downloading $(BINARY_NAME) $(VERSION)..."; \
	if ! curl -fsSL "https://github.com/$(REPO_OWNER)/$(REPO_NAME)/releases/download/$(VERSION)/$(BINARY_NAME)" -o "$$binary_path"; then \
		echo "Error: Failed to download binary"; \
		rm -rf "$$tmp_dir"; \
		exit 1; \
	fi; \
	if ! curl -fsSL "https://github.com/$(REPO_OWNER)/$(REPO_NAME)/releases/download/$(VERSION)/hashes.sha256" -o "$$checksum_path"; then \
		echo "Error: Failed to download checksum file"; \
		rm -rf "$$tmp_dir"; \
		exit 1; \
	fi; \
	expected_hash=$$(grep "$(BINARY_NAME)" "$$checksum_path" | cut -d' ' -f1); \
	if [ -z "$$expected_hash" ]; then \
		echo "Error: Could not find checksum for $(BINARY_NAME)"; \
		rm -rf "$$tmp_dir"; \
		exit 1; \
	fi; \
	actual_hash=$$(sha256sum "$$binary_path" | cut -d' ' -f1); \
	if [ "$$actual_hash" != "$$expected_hash" ]; then \
		echo "Error: Checksum mismatch!"; \
		echo "Expected: $$expected_hash"; \
		echo "Actual: $$actual_hash"; \
		rm -rf "$$tmp_dir"; \
		exit 1; \
	fi; \
	echo "Checksum verified"; \
	chmod +x "$$binary_path"; \
	if [ ! -d "$(INSTALL_DIR)" ]; then \
		echo "Creating $(INSTALL_DIR) directory..."; \
		sudo mkdir -p "$(INSTALL_DIR)"; \
	fi; \
	if [ -f "$(INSTALL_DIR)/$(BINARY_NAME)" ]; then \
		echo "Backing up existing installation..."; \
		sudo mv "$(INSTALL_DIR)/$(BINARY_NAME)" "$(INSTALL_DIR)/$(BINARY_NAME).backup"; \
	fi; \
	echo "Installing $(BINARY_NAME) to $(INSTALL_DIR)..."; \
	sudo mv "$$binary_path" "$(INSTALL_DIR)/$(BINARY_NAME)"; \
	rm -rf "$$tmp_dir"; \
	echo "Installation completed successfully!"

# ソースビルドインストール
.PHONY: install-build
install-build:
	@echo "Installing from source..."
	@if ! command -v $(PYTHON) >/dev/null 2>&1; then \
		echo "Error: $(PYTHON) is required but not installed"; \
		echo "Would you like to install Python? (y/n)"; \
		read -r answer; \
		if [ "$$answer" != "y" ] && [ "$$answer" != "Y" ]; then \
			echo "Installation cancelled"; \
			exit 1; \
		fi; \
		echo "Please install Python manually and run 'make install' again"; \
		exit 1; \
	fi
	@$(MAKE) setup-venv
	@$(MAKE) build
	@$(MAKE) install-local-binary

# 仮想環境セットアップ
.PHONY: setup-venv
setup-venv:
	@if [ ! -d "$(VENV_DIR)" ]; then \
		echo "Creating virtual environment..."; \
		$(PYTHON) -m venv $(VENV_DIR); \
	fi
	@echo "Installing dependencies..."
	@. $(VENV_DIR)/bin/activate && pip install -r requirements.txt

# ビルド
.PHONY: build
build: setup-venv
	@echo "Building binary with PyInstaller..."
	@. $(VENV_DIR)/bin/activate && pyinstaller --onefile --name $(BINARY_NAME) discord_exporter.py

# ローカルバイナリインストール
.PHONY: install-local-binary
install-local-binary:
	@if [ ! -f "dist/$(BINARY_NAME)" ]; then \
		echo "Error: Binary not found. Please run 'make build' first"; \
		exit 1; \
	fi
	@if [ ! -d "$(INSTALL_DIR)" ]; then \
		echo "Creating $(INSTALL_DIR) directory..."; \
		sudo mkdir -p "$(INSTALL_DIR)"; \
	fi
	@if [ -f "$(INSTALL_DIR)/$(BINARY_NAME)" ]; then \
		echo "Backing up existing installation..."; \
		sudo mv "$(INSTALL_DIR)/$(BINARY_NAME)" "$(INSTALL_DIR)/$(BINARY_NAME).backup"; \
	fi
	@echo "Installing $(BINARY_NAME) to $(INSTALL_DIR)..."
	@sudo cp "dist/$(BINARY_NAME)" "$(INSTALL_DIR)/$(BINARY_NAME)"
	@sudo chmod +x "$(INSTALL_DIR)/$(BINARY_NAME)"
	@echo "Installation completed successfully!"

# テスト
.PHONY: test
test:
	@if [ -f "$(INSTALL_DIR)/$(BINARY_NAME)" ]; then \
		echo "Testing installation..."; \
		$(INSTALL_DIR)/$(BINARY_NAME) --help; \
		echo "Installation test passed!"; \
	else \
		echo "Error: $(BINARY_NAME) not found in $(INSTALL_DIR)"; \
		echo "Please run 'make install' first"; \
		exit 1; \
	fi

# クリーンアップ
.PHONY: clean
clean:
	@echo "Cleaning build artifacts..."
	@rm -rf build/
	@rm -rf dist/
	@rm -rf $(VENV_DIR)/
	@rm -f *.spec
	@echo "Clean completed!"

# アンインストール
.PHONY: uninstall
uninstall:
	@if [ -f "$(INSTALL_DIR)/$(BINARY_NAME)" ]; then \
		echo "Uninstalling $(BINARY_NAME)..."; \
		sudo rm -f "$(INSTALL_DIR)/$(BINARY_NAME)"; \
		if [ -f "$(INSTALL_DIR)/$(BINARY_NAME).backup" ]; then \
			echo "Removing backup..."; \
			sudo rm -f "$(INSTALL_DIR)/$(BINARY_NAME).backup"; \
		fi; \
		echo "Uninstallation completed!"; \
	else \
		echo "$(BINARY_NAME) is not installed"; \
	fi
