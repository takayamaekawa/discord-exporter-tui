#!/usr/bin/env python3
"""
Discord Chat to XLSX Exporter
Discordチャンネルから直接XLSXファイルにエクスポートするスクリプト
"""

import argparse
import asyncio
import json
import os
import sys
import warnings
from datetime import datetime, timezone

# SSL関連の警告を抑制
warnings.filterwarnings("ignore", category=RuntimeWarning, message=".*Event loop is closed.*")
warnings.filterwarnings("ignore", category=ResourceWarning, message=".*unclosed.*")

import discord
import pandas as pd
import curses
from rich.console import Console
from rich.table import Table

# クロスプラットフォーム対応の一文字入力
def getch():
    """一文字入力を取得（クロスプラットフォーム対応）"""
    try:
        # Windows
        import msvcrt
        return msvcrt.getch().decode('utf-8')
    except ImportError:
        try:
            # Unix/Linux/Mac
            import termios, tty
            fd = sys.stdin.fileno()
            old_settings = termios.tcgetattr(fd)
            try:
                tty.setcbreak(fd)  # tty.cbreak → tty.setcbreak に修正
                ch = sys.stdin.read(1)
            finally:
                termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)
            return ch
        except Exception:
            # フォールバック: 通常の入力
            return input().strip()[:1] if input().strip() else '\n'

def get_single_key_input(prompt):
    """
    一文字入力のラッパー関数（エラーハンドリング付き）
    """
    print(prompt, end="", flush=True)
    try:
        return getch().lower()
    except Exception:
        # エラー時はEnter待ちの通常入力にフォールバック
        print("\n[一文字入力に失敗しました。Enterキーを押してください]")
        response = input().strip().lower()
        return response[0] if response else '\n'

# 必要なライブラリのインストール
# pip install discord.py pandas openpyxl


class DiscordExporter:
    def __init__(self, token):
        intents = discord.Intents.default()
        intents.message_content = True
        self.client = discord.Client(intents=intents)
        self.token = token
        self.channels_file = "channels.json"
        self.config_file = "config.json"

    async def fetch_and_save_channels(self):
        """
        全チャンネルを取得してJSONファイルに保存
        """
        await self.client.wait_until_ready()

        channels_data = []

        print("🔍 チャンネル情報を取得中...")

        for guild in self.client.guilds:
            print(f"サーバー: {guild.name}")

            for channel in guild.channels:
                # メッセージ履歴を持つチャンネルのみ対象
                if hasattr(channel, "history"):
                    try:
                        # メッセージ数を推定（最新10件から推定）
                        recent_messages = []
                        # 型チェック対応
                        from typing import cast
                        import discord
                        messageable_channel = cast(discord.abc.Messageable, channel)
                        
                        print(f"  チャンネル {channel.name} のメッセージを取得中...")
                        try:
                            async for message in messageable_channel.history(limit=10):
                                try:
                                    # メッセージの基本情報をデバッグ出力
                                    print(f"    メッセージID: {message.id} (type: {type(message.id)})")
                                    print(f"    created_at: {message.created_at} (type: {type(message.created_at)})")
                                    
                                    # created_atが正しいdatetimeオブジェクトかどうかチェック
                                    if not hasattr(message.created_at, 'year'):
                                        print(f"    警告: created_atが正しいdatetimeオブジェクトではありません")
                                        continue
                                    
                                    recent_messages.append(message)
                                except Exception as msg_error:
                                    print(f"    メッセージ処理エラー: {msg_error}")
                                    continue
                        except Exception as history_error:
                            print(f"  メッセージ履歴取得エラー: {history_error}")
                            estimated_messages = 0
                            continue

                        # 推定総メッセージ数（簡易計算）
                        print(f"  取得したメッセージ数: {len(recent_messages)}")
                        if recent_messages and len(recent_messages) > 0:
                            print(f"  メッセージの日付比較を開始...")
                            try:
                                # 比較前に各メッセージの日付をチェック
                                for i, msg in enumerate(recent_messages):
                                    print(f"    メッセージ{i}: {msg.created_at} (type: {type(msg.created_at)})")
                                
                                oldest_message = min(
                                    recent_messages, key=lambda m: m.created_at
                                )
                                newest_message = max(
                                    recent_messages, key=lambda m: m.created_at
                                )
                                print(f"  最旧: {oldest_message.created_at}, 最新: {newest_message.created_at}")
                            except (TypeError, ValueError) as e:
                                print(f"  メッセージ比較エラー: {e}")
                                print(f"  エラー詳細: recent_messagesの型 = {type(recent_messages)}")
                                for i, msg in enumerate(recent_messages):
                                    print(f"    エラーメッセージ{i}: id={msg.id}, created_at={msg.created_at} (type: {type(msg.created_at)})")
                                estimated_messages = len(recent_messages)
                            else:
                                # 正常に比較できた場合の処理
                                if len(recent_messages) >= 10:
                                    time_diff = (
                                        newest_message.created_at
                                        - oldest_message.created_at
                                    ).total_seconds()
                                    if time_diff > 0:
                                        messages_per_second = (
                                            len(recent_messages) / time_diff
                                        )
                                        channel_age = (
                                            datetime.now(timezone.utc) - channel.created_at
                                        ).total_seconds()
                                        estimated_messages = int(
                                            messages_per_second * channel_age
                                        )
                                    else:
                                        estimated_messages = len(recent_messages)
                                else:
                                    estimated_messages = len(recent_messages)
                        else:
                            estimated_messages = 0

                        # カテゴリー情報を取得
                        category_name = None
                        if hasattr(channel, 'category') and channel.category:
                            category_name = channel.category.name
                        
                        channels_data.append(
                            {
                                "guild_name": guild.name,
                                "guild_id": guild.id,
                                "channel_name": channel.name,
                                "channel_id": channel.id,
                                "channel_type": str(channel.type),
                                "category_name": category_name,
                                "estimated_messages": estimated_messages,
                                "created_at": channel.created_at.isoformat(),
                            }
                        )

                        print(
                            f"  #{channel.name} (推定メッセージ数: {estimated_messages})"
                        )

                    except Exception as e:
                        print(f"  #{channel.name} - エラー: {e}")
                        continue

        # JSONファイルに保存
        with open(self.channels_file, "w", encoding="utf-8") as f:
            json.dump(channels_data, f, ensure_ascii=False, indent=2)

        print(f"✅ チャンネル情報を {self.channels_file} に保存しました")
        await self.client.close()
        return True

    def load_channels(self):
        """
        保存されたチャンネル情報を読み込み
        """
        if os.path.exists(self.channels_file):
            with open(self.channels_file, "r", encoding="utf-8") as f:
                return json.load(f)
        return []

    def select_channels_interactive(self, use_tui=True):
        """
        インタラクティブなチャンネル選択
        """
        channels = self.load_channels()

        if not channels:
            print(
                "❌ チャンネル情報が見つかりません。最初に --fetch-channels を実行してください。"
            )
            return []

        if use_tui:
            return self._select_channels_tui(channels)
        else:
            return self._select_channels_cli(channels)

    def _select_channels_tui(self, channels):
        """
        TUI（Terminal User Interface）でチェックボックス形式のチャンネル選択
        """
        console = Console()
        
        total_estimated = sum(channel.get("estimated_messages", 0) for channel in channels)
        
        # まず統計情報を表示
        console.print("\n" + "="*80)
        console.print(f"📊 [bold]チャンネル統計[/bold]")
        console.print(f"   総チャンネル数: [cyan]{len(channels)}[/cyan]")
        console.print(f"   総推定メッセージ数: [yellow]{total_estimated:,}[/yellow]")
        console.print("="*80)
        
        # 警告表示を先に行う（設定で無効化されていない場合）
        config = self.load_config()
        should_show_warning = config.get("show_message_count_warning", True)
        
        if total_estimated > 50000 and should_show_warning:
            console.print("\n⚠️  [bold red]警告: 総メッセージ数が多すぎます！[/bold red]")
            console.print("   - 処理に非常に時間がかかる可能性があります")
            console.print("   - メモリ使用量が大きくなる可能性があります")
            console.print("   - 日付範囲やメッセージ数制限の使用を検討してください")
            console.print("")
            console.print("[dim]y: 継続 | N: 中止 | s: 今後この警告を表示しない[/dim]")

            # 一文字入力での確認
            while True:
                try:
                    response = get_single_key_input("続行しますか？ (y/N/s): ")
                    if response in ['y']:
                        print("y")
                        break
                    elif response in ['s']:
                        print("s")
                        console.print("[yellow]今後この警告を表示しないように設定しました。[/yellow]")
                        # 設定を更新
                        config["show_message_count_warning"] = False
                        self.save_config(config)
                        break
                    elif response in ['n', '\n', '\r', '\x1b']:  # n, Enter, ESC
                        print("n" if response == 'n' else "")
                        console.print("[red]処理を中止しました。[/red]")
                        return []
                    # その他のキーは無視して再入力待ち
                    print(f"\r続行しますか？ (y/N/s): ", end="", flush=True)
                except (KeyboardInterrupt, EOFError):
                    print("\n")
                    console.print("[red]処理を中止しました。[/red]")
                    return []

        # チェックボックス形式のTUIを起動
        try:
            selected_indices = curses.wrapper(self._checkbox_ui, channels)
            if selected_indices is None:
                console.print("[red]選択がキャンセルされました。[/red]")
                return []
            
            selected_channels = [channels[i] for i in selected_indices]
            return self._finalize_selection(selected_channels, console)
            
        except Exception as e:
            console.print(f"[red]TUIエラー: {e}[/red]")
            console.print("[yellow]CLIモードにフォールバックします...[/yellow]")
            return self._select_channels_cli(channels)

    def _organize_channels_by_category(self, channels):
        """
        チャンネルをカテゴリー別に整理
        """
        from collections import defaultdict
        
        categories = defaultdict(list)
        
        for channel in channels:
            category_name = channel.get('category_name', None)
            if category_name is None:
                category_name = "未分類"
            categories[category_name].append(channel)
        
        # カテゴリーをソート（未分類を最後に）
        sorted_categories = []
        for category_name in sorted(categories.keys()):
            if category_name != "未分類":
                sorted_categories.append((category_name, categories[category_name]))
        
        # 未分類を最後に追加
        if "未分類" in categories:
            sorted_categories.append(("未分類", categories["未分類"]))
        
        return sorted_categories

    def _create_display_list(self, channels):
        """
        カテゴリー別表示用のリストを作成
        """
        categories = self._organize_channels_by_category(channels)
        display_items = []
        channel_map = {}  # display_index -> channel_index
        
        for category_name, category_channels in categories:
            # カテゴリーヘッダーを追加
            display_items.append({
                "type": "category_header",
                "text": f"📁 {category_name}",
                "category_name": category_name
            })
            
            # カテゴリー内のチャンネルを追加
            for channel in category_channels:
                original_index = channels.index(channel)
                display_index = len(display_items)
                
                estimated = channel.get("estimated_messages", 0)
                text = f"  ☐ #{channel['channel_name']} ({estimated:,})"
                
                display_items.append({
                    "type": "channel",
                    "text": text,
                    "channel": channel,
                    "original_index": original_index
                })
                
                channel_map[display_index] = original_index
        
        return display_items, channel_map

    def _checkbox_ui(self, stdscr, channels):
        """
        cursesを使用したチェックボックスUI（カテゴリー別表示対応）
        """
        curses.curs_set(0)  # カーソルを非表示
        stdscr.keypad(1)    # 特殊キーを有効化
        
        # カラーペアの定義
        curses.start_color()
        curses.init_pair(1, curses.COLOR_WHITE, curses.COLOR_BLUE)   # 選択行
        curses.init_pair(2, curses.COLOR_GREEN, curses.COLOR_BLACK)  # チェック済み
        curses.init_pair(3, curses.COLOR_YELLOW, curses.COLOR_BLACK) # ヘッダー
        curses.init_pair(4, curses.COLOR_RED, curses.COLOR_BLACK)    # 警告
        curses.init_pair(5, curses.COLOR_CYAN, curses.COLOR_BLACK)   # カテゴリー
        
        # カテゴリー別表示リストを作成
        display_items, channel_map = self._create_display_list(channels)
        
        # ページング設定
        ITEMS_PER_PAGE = 15
        FIRST_PAGE_ITEMS = 12  # 1ページ目はAllオプションがあるので少なめ
        current_pos = 0
        current_page = 0
        
        # ページ数計算
        remaining_items = len(display_items) - FIRST_PAGE_ITEMS
        if remaining_items <= 0:
            total_pages = 1
        else:
            total_pages = 1 + ((remaining_items + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
        
        # チェック状態を管理
        checked = [False] * len(channels)
        all_checked = False
        
        while True:
            stdscr.clear()
            height, width = stdscr.getmaxyx()
            
            # ヘッダー情報
            title = "Discord チャンネル選択"
            stdscr.addstr(0, (width - len(title)) // 2, title, curses.color_pair(3) | curses.A_BOLD)
            
            page_info = f"ページ {current_page + 1}/{total_pages} | 選択済み: {sum(checked)}/{len(channels)}"
            stdscr.addstr(1, (width - len(page_info)) // 2, page_info, curses.color_pair(3))
            
            # 操作説明
            help_text = "↑↓/jk: 移動 | SPACE: チェック/カテゴリ選択 | A: 全選択/解除 | ENTER: 確定 | Q: キャンセル | ←→/hl: ページ移動"
            if len(help_text) < width:
                stdscr.addstr(2, (width - len(help_text)) // 2, help_text)
            
            stdscr.addstr(3, 0, "="*min(width-1, 80))
            
            # 現在ページの表示範囲を計算
            if current_page == 0:
                start_idx = 0
                end_idx = min(FIRST_PAGE_ITEMS, len(display_items))
            else:
                start_idx = FIRST_PAGE_ITEMS + (current_page - 1) * ITEMS_PER_PAGE
                end_idx = min(start_idx + ITEMS_PER_PAGE, len(display_items))
            
            y_offset = 5
            
            # "All"オプション（1ページ目のみ）
            if current_page == 0:
                all_symbol = "☑" if all_checked else "☐"
                all_text = f"{all_symbol} All ({len(channels)} channels)"
                
                if current_pos == 0:
                    stdscr.addstr(y_offset, 2, all_text, curses.color_pair(1) | curses.A_BOLD)
                else:
                    stdscr.addstr(y_offset, 2, all_text, curses.color_pair(2) if all_checked else 0)
                y_offset += 1
                
                # 区切り線
                stdscr.addstr(y_offset, 2, "-" * min(width-4, 40), curses.color_pair(3))
                y_offset += 2
            
            # 表示アイテム一覧（カテゴリー + チャンネル）
            display_pos = 0  # 表示アイテムの位置カウンター
            
            for i in range(start_idx, end_idx):
                if i >= len(display_items):
                    break
                    
                item = display_items[i]
                
                # 選択位置の計算（Allオプションがあるかどうかを考慮）
                is_selected = False
                if current_page == 0:
                    # 1ページ目：「All」オプション(pos=0)があるので、アイテム選択は pos-1 と比較
                    is_selected = (current_pos - 1) == display_pos
                else:
                    # 2ページ目以降：「All」オプションがないので、直接比較
                    is_selected = current_pos == display_pos
                
                if item["type"] == "category_header":
                    # カテゴリーヘッダー（選択可能）
                    category_name = item["category_name"]
                    
                    # カテゴリ内のチャンネル選択状態を確認
                    category_channels = [ch for ch in channels if ch.get('category_name', '未分類') == category_name]
                    selected_in_category = 0
                    try:
                        selected_in_category = sum(1 for ch in category_channels if checked[channels.index(ch)])
                    except (ValueError, IndexError):
                        # インデックスが見つからない場合は0として継続
                        selected_in_category = 0
                    total_in_category = len(category_channels)
                    
                    # カテゴリーヘッダーの表示テキストを作成
                    if selected_in_category == total_in_category:
                        symbol = "☑"  # 全選択
                    elif selected_in_category > 0:
                        symbol = "▣"  # 部分選択
                    else:
                        symbol = "☐"  # 未選択
                    
                    display_text = f"{symbol} {item['text']} ({selected_in_category}/{total_in_category})"
                    
                    # ハイライト表示
                    if is_selected:
                        stdscr.addstr(y_offset, 0, display_text, curses.color_pair(1) | curses.A_BOLD)
                    else:
                        color = curses.color_pair(2) if selected_in_category > 0 else curses.color_pair(5) | curses.A_BOLD
                        stdscr.addstr(y_offset, 0, display_text, color)
                    
                    y_offset += 1
                    
                elif item["type"] == "channel":
                    # チャンネル（選択可能）
                    original_idx = item["original_index"]
                    symbol = "☑" if checked[original_idx] else "☐"
                    text = item["text"].replace("☐", symbol)
                    
                    # 画面幅に合わせてカット
                    if len(text) > width - 4:
                        text = text[:width-7] + "..."
                    
                    # ハイライト表示
                    if is_selected:
                        stdscr.addstr(y_offset, 0, text, curses.color_pair(1) | curses.A_BOLD)
                    else:
                        color = curses.color_pair(2) if checked[original_idx] else 0
                        stdscr.addstr(y_offset, 0, text, color)
                    
                    y_offset += 1
                
                display_pos += 1  # 表示アイテムの位置をインクリメント
                
                # 画面の高さを超えないようにチェック
                if y_offset >= height - 2:
                    break
            
            # フッター
            footer = f"選択中: {sum(checked)} チャンネル"
            if current_page < total_pages - 1:
                footer += " | 次ページ: →"
            if current_page > 0:
                footer += " | 前ページ: ←"
            
            stdscr.addstr(height-1, 0, footer[:width-1], curses.color_pair(3))
            
            stdscr.refresh()
            
            # キー入力処理
            key = stdscr.getch()
            
            if key == ord('q') or key == ord('Q'):
                return None  # キャンセル
            
            elif key == ord('\n') or key == 10:  # Enter
                if sum(checked) == 0:
                    # 何も選択されていない場合のメッセージ
                    stdscr.addstr(height-2, 0, "少なくとも1つのチャンネルを選択してください", curses.color_pair(4))
                    stdscr.refresh()
                    stdscr.getch()
                    continue
                return [i for i, c in enumerate(checked) if c]
            
            elif key == ord(' '):  # スペースキー
                if current_page == 0 and current_pos == 0:
                    # "All"を切り替え
                    all_checked = not all_checked
                    checked = [all_checked] * len(channels)
                else:
                    # 現在の表示範囲を取得
                    if current_page == 0:
                        start_idx = 0
                        end_idx = min(FIRST_PAGE_ITEMS, len(display_items))
                    else:
                        start_idx = FIRST_PAGE_ITEMS + (current_page - 1) * ITEMS_PER_PAGE
                        end_idx = min(start_idx + ITEMS_PER_PAGE, len(display_items))
                    
                    # 現在選択されているアイテムを見つける
                    display_pos = 0  # 表示アイテムの位置カウンター
                    
                    for i in range(start_idx, end_idx):
                        if i >= len(display_items):
                            break
                            
                        item = display_items[i]
                        
                        # ハイライト判定（表示ロジックと同じ）
                        is_selected = False
                        if current_page == 0:
                            is_selected = (current_pos - 1) == display_pos
                        else:
                            is_selected = current_pos == display_pos
                        
                        if is_selected:
                            if item["type"] == "category_header":
                                # カテゴリー選択：そのカテゴリ内の全チャンネルを切り替え
                                category_name = item["category_name"]
                                category_channels = [ch for ch in channels if ch.get('category_name', '未分類') == category_name]
                                
                                # カテゴリ内の選択状態を確認
                                try:
                                    category_indices = [channels.index(ch) for ch in category_channels]
                                    selected_in_category = sum(1 for idx in category_indices if checked[idx])
                                except (ValueError, IndexError):
                                    # インデックスが見つからない場合はスキップ
                                    break
                                else:
                                    # 全選択なら全解除、そうでなければ全選択
                                    new_state = selected_in_category < len(category_channels)
                                    for idx in category_indices:
                                        checked[idx] = new_state
                                
                                all_checked = all(checked)
                                break
                                
                            elif item["type"] == "channel":
                                # 個別チャンネルを切り替え
                                original_idx = item["original_index"]
                                checked[original_idx] = not checked[original_idx]
                                all_checked = all(checked)
                                break
                        
                        display_pos += 1
            
            elif key == ord('a') or key == ord('A'):
                # 全選択/全解除切り替え
                all_checked = not all_checked
                checked = [all_checked] * len(channels)
            
            elif key == curses.KEY_UP or key == ord('k') or key == ord('K'):
                if current_pos > 0:
                    current_pos -= 1
                elif current_page > 0:
                    current_page -= 1
                    # 前ページの最後の項目に移動
                    if current_page == 0:
                        # 1ページ目の場合：Allオプション + 全アイテム
                        search_start = 0
                        search_end = min(FIRST_PAGE_ITEMS, len(display_items))
                        item_count = search_end - search_start
                        current_pos = item_count  # Allオプション(0) + 全アイテム数
                    else:
                        # 2ページ目以降の場合
                        search_start = FIRST_PAGE_ITEMS + (current_page - 1) * ITEMS_PER_PAGE
                        search_end = min(search_start + ITEMS_PER_PAGE, len(display_items))
                        item_count = search_end - search_start
                        current_pos = item_count - 1
            
            elif key == curses.KEY_DOWN or key == ord('j') or key == ord('J'):
                # 現在ページでのアイテム数を計算（カテゴリー + チャンネル）
                if current_page == 0:
                    search_start = 0
                    search_end = min(FIRST_PAGE_ITEMS, len(display_items))
                else:
                    search_start = FIRST_PAGE_ITEMS + (current_page - 1) * ITEMS_PER_PAGE
                    search_end = min(search_start + ITEMS_PER_PAGE, len(display_items))
                
                # ページ内の全アイテム数をカウント
                item_count = search_end - search_start
                
                max_pos_on_page = item_count
                if current_page == 0:
                    max_pos_on_page += 1  # Allオプションの分を追加
                
                if current_pos < max_pos_on_page - 1:
                    current_pos += 1
                elif current_page < total_pages - 1:
                    current_page += 1
                    current_pos = 0
            
            elif key == curses.KEY_LEFT or key == ord('h') or key == ord('H'):
                if current_page > 0:
                    current_page -= 1
                    current_pos = 0
            
            elif key == curses.KEY_RIGHT or key == ord('l') or key == ord('L'):
                if current_page < total_pages - 1:
                    current_page += 1
                    current_pos = 0

    def _select_channels_with_paging(self, channels, console):
        """
        ページング機能付きチャンネル選択
        """
        CHANNELS_PER_PAGE = 20
        total_pages = (len(channels) + CHANNELS_PER_PAGE - 1) // CHANNELS_PER_PAGE
        selected_channels = []
        
        console.print(f"\n🎯 [bold]チャンネル選択（ページング機能付き）[/bold]")
        console.print(f"   総ページ数: {total_pages}")
        console.print("   各ページでチャンネルを選択し、最後に確定します\n")
        
        for page in range(total_pages):
            start_idx = page * CHANNELS_PER_PAGE
            end_idx = min(start_idx + CHANNELS_PER_PAGE, len(channels))
            page_channels = channels[start_idx:end_idx]
            
            console.print(f"--- ページ {page + 1}/{total_pages} ({start_idx + 1}-{end_idx}) ---")
            
            # ページ内のチャンネルを表示
            choices = []
            for i, channel in enumerate(page_channels):
                estimated = channel.get("estimated_messages", 0)
                display_name = f"[{channel['guild_name']}] #{channel['channel_name']} (推定: {estimated:,} メッセージ)"
                choices.append((display_name, start_idx + i))
                console.print(f"  {start_idx + i + 1:2d}. {display_name}")
            
            # このページでの選択
            console.print(f"\nページ {page + 1} での選択:")
            print("選択方法:")
            print("  - 番号をカンマ区切りで入力 (例: 1,3,5)")
            print("  - 範囲指定可能 (例: 1-5)")
            print("  - 'all' でページ内全選択")
            print("  - 'skip' でページをスキップ")
            print("  - 'done' で選択完了")
            
            while True:
                try:
                    selection = input(f"ページ {page + 1} 選択: ").strip()
                    
                    if selection.lower() == 'skip':
                        break
                    elif selection.lower() == 'done':
                        if selected_channels:
                            return self._finalize_selection(selected_channels, console)
                        else:
                            print("まだチャンネルが選択されていません。")
                            continue
                    elif selection.lower() == 'all':
                        selected_channels.extend(page_channels)
                        console.print(f"[green]ページ {page + 1} の全チャンネルを選択しました[/green]")
                        break
                    else:
                        # 番号解析
                        page_selected = self._parse_selection(selection, len(page_channels))
                        if page_selected:
                            for idx in page_selected:
                                if 0 <= idx < len(page_channels):
                                    selected_channels.append(page_channels[idx])
                            console.print(f"[green]ページ {page + 1} から {len(page_selected)} チャンネルを選択しました[/green]")
                            break
                        else:
                            print("無効な選択です。")
                            continue
                            
                except (ValueError, KeyboardInterrupt):
                    print("無効な入力です。")
                    continue
            
            # 現在の選択状況を表示
            if selected_channels:
                console.print(f"\n現在の選択: {len(selected_channels)} チャンネル")
        
        return self._finalize_selection(selected_channels, console)

    def _select_channels_simple(self, channels, console):
        """
        シンプルなチャンネル選択（少数の場合）
        """
        # チャンネル一覧を表示
        table = Table(show_header=True, header_style="bold magenta")
        table.add_column("No.", style="dim", width=4)
        table.add_column("サーバー", style="cyan")
        table.add_column("チャンネル", style="green")
        table.add_column("推定メッセージ数", justify="right", style="yellow")

        for i, channel in enumerate(channels, 1):
            estimated = channel.get("estimated_messages", 0)
            table.add_row(
                str(i),
                channel['guild_name'],
                f"#{channel['channel_name']}",
                f"{estimated:,}"
            )

        console.print("\n📋 利用可能なチャンネル:")
        console.print(table)
        
        console.print("\n🎯 エクスポートしたいチャンネルを選択してください:")
        print("選択方法:")
        print("  - 番号をカンマ区切りで入力 (例: 1,3,5)")
        print("  - 範囲指定可能 (例: 1-5)")
        print("  - 'all' で全選択")
        
        while True:
            try:
                selection = input("\n選択: ").strip()
                
                if selection.lower() == 'all':
                    selected_channels = channels[:]
                    break
                else:
                    selected_indices = self._parse_selection(selection, len(channels))
                    if selected_indices:
                        selected_channels = [channels[i] for i in selected_indices]
                        break
                    else:
                        print("無効な選択です。")
                        continue
                        
            except (ValueError, KeyboardInterrupt):
                print("無効な入力です。")
                continue
        
        return self._finalize_selection(selected_channels, console)

    def _parse_selection(self, selection, max_count):
        """
        選択文字列を解析してインデックスリストを返す
        """
        selected_indices = []
        
        for part in selection.split(','):
            part = part.strip()
            if '-' in part:
                try:
                    start, end = map(int, part.split('-'))
                    selected_indices.extend(range(start-1, end))
                except ValueError:
                    return None
            else:
                try:
                    selected_indices.append(int(part) - 1)
                except ValueError:
                    return None
        
        # 重複削除と範囲チェック
        selected_indices = list(set(selected_indices))
        selected_indices = [i for i in selected_indices if 0 <= i < max_count]
        
        return selected_indices if selected_indices else None

    def _finalize_selection(self, selected_channels, console):
        """
        選択結果の確認と表示
        """
        if not selected_channels:
            console.print("[red]チャンネルが選択されていません。[/red]")
            return []

        # 結果表示
        console.print(f"\n✅ [bold green]{len(selected_channels)} チャンネルを選択しました:[/bold green]")
        selected_total = 0
        for channel in selected_channels:
            estimated = channel.get("estimated_messages", 0)
            selected_total += estimated
            console.print(
                f"   - [{channel['guild_name']}] #{channel['channel_name']} "
                f"(推定: {estimated:,} メッセージ)"
            )

        console.print(f"\n選択チャンネルの総推定メッセージ数: [bold yellow]{selected_total:,}[/bold yellow]")

        return selected_channels

    def _select_channels_cli(self, channels):
        """
        従来のCLIでチャンネル選択
        """
        print("\n📋 利用可能なチャンネル:")
        print("=" * 80)

        total_estimated = 0
        for i, channel in enumerate(channels, 1):
            estimated = channel.get("estimated_messages", 0)
            total_estimated += estimated
            print(
                f"{i:2d}. [{channel['guild_name']}] #{channel['channel_name']} "
                f"(推定: {estimated:,} メッセージ)"
            )

        print("=" * 80)
        print(f"総推定メッセージ数: {total_estimated:,}")

        # 警告表示
        if total_estimated > 50000:
            print("\n⚠️  警告: 総メッセージ数が多すぎます！")
            print("   - 処理に非常に時間がかかる可能性があります")
            print("   - メモリ使用量が大きくなる可能性があります")
            print("   - 日付範囲やメッセージ数制限の使用を検討してください")

            confirm = input("\n続行しますか？ (y/N): ").strip().lower()
            if confirm != "y":
                print("処理を中止しました。")
                return []

        print("\n🎯 エクスポートしたいチャンネルを選択してください:")
        print("   - 番号をカンマ区切りで入力 (例: 1,3,5-7)")
        print("   - 範囲指定可能 (例: 1-5)")
        print("   - 'all' で全選択")

        while True:
            try:
                selection = input("\n選択: ").strip()

                if selection.lower() == "all":
                    selected_channels = channels[:]
                    break

                selected_indices = []

                for part in selection.split(","):
                    part = part.strip()
                    if "-" in part:
                        start, end = map(int, part.split("-"))
                        selected_indices.extend(range(start - 1, end))
                    else:
                        selected_indices.append(int(part) - 1)

                # 重複削除と範囲チェック
                selected_indices = list(set(selected_indices))
                selected_indices = [
                    i for i in selected_indices if 0 <= i < len(channels)
                ]

                if not selected_indices:
                    print("❌ 有効な選択がありません。")
                    continue

                selected_channels = [channels[i] for i in selected_indices]
                break

            except ValueError:
                print("❌ 無効な入力です。数字とカンマ、ハイフンのみ使用してください。")
                continue

        print(f"\n✅ {len(selected_channels)} チャンネルを選択しました:")
        selected_total = 0
        for channel in selected_channels:
            estimated = channel.get("estimated_messages", 0)
            selected_total += estimated
            print(
                f"   - [{channel['guild_name']}] #{channel['channel_name']} "
                f"(推定: {estimated:,} メッセージ)"
            )

        print(f"\n選択チャンネルの総推定メッセージ数: {selected_total:,}")

        return selected_channels

    def save_config(self, config):
        """
        設定をJSONファイルに保存
        """
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)

    def load_config(self):
        """
        設定をJSONファイルから読み込み
        """
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                pass
        
        # デフォルト設定
        return {
            "token": "",
            "output_file": "discord_export.xlsx",
            "after_date": "",
            "before_date": "",
            "limit": "",
            "mode": "interactive",
            "show_message_count_warning": True
        }

    def get_channels_info(self):
        """
        チャンネル情報の状態を取得
        """
        if not os.path.exists(self.channels_file):
            return None
        
        try:
            stat = os.stat(self.channels_file)
            last_modified = datetime.fromtimestamp(stat.st_mtime)
            
            with open(self.channels_file, 'r', encoding='utf-8') as f:
                channels_data = json.load(f)
            
            return {
                "count": len(channels_data),
                "last_modified": last_modified,
                "data": channels_data
            }
        except Exception:
            return None

    def main_menu(self):
        """
        メインメニューを表示
        """
        try:
            return curses.wrapper(self._main_menu_ui)
        except Exception as e:
            console = Console()
            console.print(f"[red]メニューUIエラー: {e}[/red]")
            return self._main_menu_cli()

    def _main_menu_ui(self, stdscr):
        """
        cursesを使用したメインメニューUI
        """
        curses.curs_set(0)
        stdscr.keypad(1)
        
        # カラーペアの定義
        curses.start_color()
        curses.init_pair(1, curses.COLOR_WHITE, curses.COLOR_BLUE)   # 選択行
        curses.init_pair(2, curses.COLOR_GREEN, curses.COLOR_BLACK)  # 成功
        curses.init_pair(3, curses.COLOR_YELLOW, curses.COLOR_BLACK) # ヘッダー
        curses.init_pair(4, curses.COLOR_RED, curses.COLOR_BLACK)    # 警告
        curses.init_pair(5, curses.COLOR_CYAN, curses.COLOR_BLACK)   # 情報
        
        current_pos = 0
        
        while True:
            stdscr.clear()
            height, width = stdscr.getmaxyx()
            
            # ヘッダー
            title = "Discord Exporter メインメニュー"
            stdscr.addstr(0, (width - len(title)) // 2, title, curses.color_pair(3) | curses.A_BOLD)
            
            help_text = "↑↓/jk: 移動 | ENTER/SPACE: 選択 | Q: 終了"
            if len(help_text) < width:
                stdscr.addstr(1, (width - len(help_text)) // 2, help_text, curses.color_pair(5))
            
            stdscr.addstr(2, 0, "="*min(width-1, 80))
            
            # チャンネル情報の状態を表示
            y_offset = 4
            channels_info = self.get_channels_info()
            
            if channels_info:
                stdscr.addstr(y_offset, 2, f"📊 チャンネル情報: {channels_info['count']} チャンネル", curses.color_pair(2))
                stdscr.addstr(y_offset + 1, 2, f"最終更新: {channels_info['last_modified'].strftime('%Y-%m-%d %H:%M:%S')}", curses.color_pair(5))
                
                # 更新が古い場合の警告
                days_old = (datetime.now() - channels_info['last_modified']).days
                if days_old > 7:
                    stdscr.addstr(y_offset + 2, 2, f"⚠️  {days_old}日前の情報です（更新を推奨）", curses.color_pair(4))
                    y_offset += 1
            else:
                stdscr.addstr(y_offset, 2, "❌ チャンネル情報がありません", curses.color_pair(4))
                stdscr.addstr(y_offset + 1, 2, "最初にチャンネル情報を取得してください", curses.color_pair(5))
            
            y_offset += 4
            
            # メニュー項目
            menu_items = [
                ("1. チャンネル情報を更新", "update_channels"),
                ("2. チャンネルを選択してエクスポート", "export_interactive"),
                ("3. 設定を変更", "config"),
                ("4. 終了", "exit")
            ]
            
            for i, (label, action) in enumerate(menu_items):
                if i == current_pos:
                    stdscr.addstr(y_offset + i * 2, 4, f"→ {label}", curses.color_pair(1) | curses.A_BOLD)
                else:
                    stdscr.addstr(y_offset + i * 2, 4, f"  {label}")
            
            # フッター情報
            footer_y = height - 3
            config = self.load_config()
            if config.get("token"):
                stdscr.addstr(footer_y, 2, f"Bot Token: 設定済み", curses.color_pair(2))
            else:
                stdscr.addstr(footer_y, 2, f"Bot Token: 未設定", curses.color_pair(4))
            
            stdscr.addstr(footer_y + 1, 2, f"出力ファイル: {config.get('output_file', '未設定')}", curses.color_pair(5))
            
            stdscr.refresh()
            
            # キー入力処理
            key = stdscr.getch()
            
            if key == ord('q') or key == ord('Q'):
                return "exit"
            
            elif key == ord('\n') or key == 10 or key == ord(' '):  # Enter or Space
                return menu_items[current_pos][1]
            
            elif key == curses.KEY_UP or key == ord('k') or key == ord('K'):
                current_pos = (current_pos - 1) % len(menu_items)
            
            elif key == curses.KEY_DOWN or key == ord('j') or key == ord('J'):
                current_pos = (current_pos + 1) % len(menu_items)
            
            elif key in [ord('1'), ord('2'), ord('3'), ord('4')]:
                # 数字キーで直接選択
                num = int(chr(key)) - 1
                if 0 <= num < len(menu_items):
                    return menu_items[num][1]

    def _main_menu_cli(self):
        """
        CLIでのメインメニュー
        """
        console = Console()
        
        while True:
            console.print("\n" + "="*60)
            console.print("🎯 [bold]Discord Exporter メインメニュー[/bold]")
            console.print("="*60)
            
            # チャンネル情報の状態を表示
            channels_info = self.get_channels_info()
            
            if channels_info:
                console.print(f"📊 チャンネル情報: [green]{channels_info['count']} チャンネル[/green]")
                console.print(f"最終更新: [cyan]{channels_info['last_modified'].strftime('%Y-%m-%d %H:%M:%S')}[/cyan]")
                
                days_old = (datetime.now() - channels_info['last_modified']).days
                if days_old > 7:
                    console.print(f"⚠️  [yellow]{days_old}日前の情報です（更新を推奨）[/yellow]")
            else:
                console.print("❌ [red]チャンネル情報がありません[/red]")
                console.print("最初にチャンネル情報を取得してください")
            
            console.print()
            
            # メニュー項目
            console.print("[bold]メニュー:[/bold]")
            console.print("1. チャンネル情報を更新")
            console.print("2. チャンネルを選択してエクスポート")  
            console.print("3. 設定を変更")
            console.print("4. 終了")
            
            choice = input("\n選択 (1-4): ").strip()
            
            if choice == "1":
                return "update_channels"
            elif choice == "2":
                return "export_interactive"
            elif choice == "3":
                return "config"
            elif choice == "4" or choice.lower() == "q":
                return "exit"
            else:
                console.print("[red]無効な選択です。1-4の数字を入力してください。[/red]")

    def ask_continue(self):
        """
        メインメニューに戻るかどうかを確認
        """
        console = Console()
        console.print("\n" + "="*50)
        console.print("🎯 [bold]操作完了[/bold]")
        console.print("="*50)
        
        while True:
            try:
                response = get_single_key_input("メインメニューに戻りますか？ (y/N): ")
                if response in ['y']:
                    print("y")
                    return True
                elif response in ['n', '\n', '\r', '\x1b']:  # n, Enter, ESC
                    print("n" if response == 'n' else "")
                    return False
                # その他のキーは無視して再入力待ち
                print(f"\r続行しますか？ (y/N): ", end="", flush=True)
            except (KeyboardInterrupt, EOFError):
                print("\n")
                return False

    async def cleanup_client(self):
        """
        Discordクライアントの適切なクリーンアップ
        """
        try:
            import asyncio
            # イベントループが開いているかどうかチェック
            loop = asyncio.get_event_loop()
            if loop.is_closed():
                return
            
            # クライアントが接続されているかどうかチェック
            if hasattr(self.client, '_ready') and self.client._ready.is_set():
                print("接続をクローズ中...")
                await self.client.close()
                print("接続クローズ完了")
            elif not self.client.is_closed():
                await self.client.close()
                
        except Exception as close_error:
            # クリーンアップエラーは静かに無視
            pass

    def config_ui(self):
        """
        設定UIを表示してパラメータを収集
        """
        try:
            config = curses.wrapper(self._config_form_ui)
            if config is None:
                print("設定がキャンセルされました。")
                return None
            
            # 設定を保存
            self.save_config(config)
            return config
            
        except Exception as e:
            console = Console()
            console.print(f"[red]設定UIエラー: {e}[/red]")
            console.print("[yellow]CLIモードで設定を入力してください...[/yellow]")
            return self._config_cli()

    def _config_form_ui(self, stdscr):
        """
        cursesを使用した設定フォームUI
        """
        curses.curs_set(1)  # カーソルを表示
        stdscr.keypad(1)
        
        # カラーペアの定義
        curses.start_color()
        curses.init_pair(1, curses.COLOR_WHITE, curses.COLOR_BLUE)   # 選択行
        curses.init_pair(2, curses.COLOR_GREEN, curses.COLOR_BLACK)  # 完了済み
        curses.init_pair(3, curses.COLOR_YELLOW, curses.COLOR_BLACK) # ヘッダー
        curses.init_pair(4, curses.COLOR_RED, curses.COLOR_BLACK)    # エラー
        curses.init_pair(5, curses.COLOR_CYAN, curses.COLOR_BLACK)   # 説明
        
        # 既存の設定を読み込み
        config = self.load_config()
        
        # フォームフィールド
        fields = [
            {"name": "token", "label": "Discord Bot Token", "value": config.get("token", ""), "type": "password"},
            {"name": "output_file", "label": "出力ファイル名", "value": config.get("output_file", "discord_export.xlsx"), "type": "text"},
            {"name": "after_date", "label": "開始日 (YYYY-MM-DD)", "value": config.get("after_date", ""), "type": "text"},
            {"name": "before_date", "label": "終了日 (YYYY-MM-DD)", "value": config.get("before_date", ""), "type": "text"},
            {"name": "limit", "label": "メッセージ数制限", "value": str(config.get("limit", "")), "type": "number"},
            {"name": "mode", "label": "実行モード", "value": config.get("mode", "interactive"), "type": "select", "options": ["fetch-channels", "interactive", "cli"]}
        ]
        
        current_field = 0
        editing = False
        edit_text = ""
        
        while True:
            stdscr.clear()
            height, width = stdscr.getmaxyx()
            
            # ヘッダー
            title = "Discord Exporter 設定"
            stdscr.addstr(0, (width - len(title)) // 2, title, curses.color_pair(3) | curses.A_BOLD)
            
            help_text = "↑↓/jk: 移動 | ENTER: 編集/選択 | TAB: 次へ | F10: 保存してメニューに戻る | ESC: キャンセル"
            if len(help_text) < width:
                stdscr.addstr(1, (width - len(help_text)) // 2, help_text, curses.color_pair(5))
            
            stdscr.addstr(2, 0, "="*min(width-1, 80))
            
            # フォームフィールドを表示
            y_offset = 4
            for i, field in enumerate(fields):
                label = field["label"] + ":"
                value = field["value"]
                
                # ラベル表示
                if i == current_field:
                    stdscr.addstr(y_offset, 2, label, curses.color_pair(1) | curses.A_BOLD)
                else:
                    stdscr.addstr(y_offset, 2, label)
                
                # 値の表示
                display_value = value
                if field["type"] == "password" and value:
                    display_value = "*" * len(value)
                elif field["type"] == "select":
                    display_value = f"[{value}]"
                
                if editing and i == current_field:
                    # 編集中
                    if field["type"] == "select":
                        # セレクトボックスの選択肢を表示
                        stdscr.addstr(y_offset + 1, 4, "選択肢:", curses.color_pair(5))
                        for j, option in enumerate(field["options"]):
                            marker = ">" if option == value else " "
                            color = curses.color_pair(1) if option == value else 0
                            stdscr.addstr(y_offset + 2 + j, 6, f"{marker} {option}", color)
                    else:
                        display_value = edit_text + "_"
                        stdscr.addstr(y_offset + 1, 4, display_value, curses.color_pair(1))
                else:
                    stdscr.addstr(y_offset + 1, 4, display_value)
                
                # フィールドの説明
                if i == current_field:
                    descriptions = {
                        "token": "Discord Developer PortalでBotを作成してTokenを取得",
                        "output_file": "エクスポートするExcelファイル名 (.xlsx)",
                        "after_date": "この日付以降のメッセージのみ (例: 2024-01-01)",
                        "before_date": "この日付以前のメッセージのみ (例: 2024-12-31)",
                        "limit": "チャンネル毎の最大メッセージ数 (空白=制限なし)",
                        "mode": "fetch-channels: チャンネル情報取得, interactive: TUI選択, cli: CLI選択"
                    }
                    desc = descriptions.get(field["name"], "")
                    if desc and len(desc) < width - 6:
                        stdscr.addstr(y_offset + 2 + (len(field.get("options", [])) if editing and field["type"] == "select" else 0), 
                                    4, desc, curses.color_pair(5))
                
                y_offset += 4 + (len(field.get("options", [])) if editing and field["type"] == "select" and i == current_field else 0)
                
                if y_offset >= height - 5:
                    break
            
            # フッター
            footer = "F10: 保存してメニューに戻る"
            if any(not field["value"] for field in fields[:2]):  # tokenとoutput_fileは必須
                footer = "必須項目を入力してください"
            
            stdscr.addstr(height-1, 0, footer[:width-1], curses.color_pair(3))
            stdscr.refresh()
            
            # キー入力処理
            key = stdscr.getch()
            
            if key == 27:  # ESC
                return None
            
            elif key == curses.KEY_F10:  # F10で実行
                # 必須項目チェック
                if not fields[0]["value"] or not fields[1]["value"]:
                    continue
                
                # 設定を辞書に変換
                result_config = {}
                for field in fields:
                    if field["value"]:
                        if field["type"] == "number":
                            try:
                                result_config[field["name"]] = int(field["value"])
                            except ValueError:
                                result_config[field["name"]] = ""
                        else:
                            result_config[field["name"]] = field["value"]
                    else:
                        result_config[field["name"]] = ""
                
                return result_config
            
            elif key == ord('\t'):  # TAB
                if not editing:
                    current_field = (current_field + 1) % len(fields)
            
            elif key == curses.KEY_UP or key == ord('k') or key == ord('K'):
                if not editing:
                    current_field = (current_field - 1) % len(fields)
                elif fields[current_field]["type"] == "select":
                    options = fields[current_field]["options"]
                    current_idx = options.index(fields[current_field]["value"])
                    new_idx = (current_idx - 1) % len(options)
                    fields[current_field]["value"] = options[new_idx]
            
            elif key == curses.KEY_DOWN or key == ord('j') or key == ord('J'):
                if not editing:
                    current_field = (current_field + 1) % len(fields)
                elif fields[current_field]["type"] == "select":
                    options = fields[current_field]["options"]
                    current_idx = options.index(fields[current_field]["value"])
                    new_idx = (current_idx + 1) % len(options)
                    fields[current_field]["value"] = options[new_idx]
            
            elif key == ord('\n') or key == 10:  # Enter
                if editing:
                    if fields[current_field]["type"] != "select":
                        fields[current_field]["value"] = edit_text
                    editing = False
                else:
                    if fields[current_field]["type"] == "select":
                        editing = True
                    else:
                        editing = True
                        edit_text = fields[current_field]["value"]
            
            elif editing and fields[current_field]["type"] != "select":
                if key == curses.KEY_BACKSPACE or key == 127:
                    edit_text = edit_text[:-1]
                elif 32 <= key <= 126:  # 印刷可能文字
                    edit_text += chr(key)

    def _config_cli(self):
        """
        CLIでの設定入力
        """
        config = self.load_config()
        
        print("\n=== Discord Exporter 設定 ===")
        
        # Discord Bot Token
        token = input(f"Discord Bot Token [{config.get('token', '')}]: ").strip()
        if not token:
            token = config.get('token', '')
        
        # 出力ファイル名
        output_file = input(f"出力ファイル名 [{config.get('output_file', 'discord_export.xlsx')}]: ").strip()
        if not output_file:
            output_file = config.get('output_file', 'discord_export.xlsx')
        
        # 開始日
        after_date = input(f"開始日 (YYYY-MM-DD) [{config.get('after_date', '')}]: ").strip()
        if not after_date:
            after_date = config.get('after_date', '')
        
        # 終了日
        before_date = input(f"終了日 (YYYY-MM-DD) [{config.get('before_date', '')}]: ").strip()
        if not before_date:
            before_date = config.get('before_date', '')
        
        # メッセージ数制限
        limit = input(f"メッセージ数制限 [{config.get('limit', '')}]: ").strip()
        if not limit:
            limit = config.get('limit', '')
        
        # 実行モード
        print("実行モード:")
        print("  1. fetch-channels (チャンネル情報取得)")
        print("  2. interactive (TUI選択)")
        print("  3. cli (CLI選択)")
        
        mode_map = {"1": "fetch-channels", "2": "interactive", "3": "cli"}
        mode_choice = input(f"選択 [2]: ").strip()
        mode = mode_map.get(mode_choice, "interactive")
        
        return {
            "token": token,
            "output_file": output_file,
            "after_date": after_date,
            "before_date": before_date,
            "limit": int(limit) if limit.isdigit() else "",
            "mode": mode
        }

    async def export_channel_to_xlsx(
        self, channel_id, output_file, after_date=None, before_date=None, limit=None
    ):
        """
        DiscordチャンネルをXLSXファイルにエクスポート

        Args:
            channel_id (int): チャンネルID
            output_file (str): 出力ファイル名
            after_date (datetime): この日付以降のメッセージ
            before_date (datetime): この日付以前のメッセージ
            limit (int): メッセージ数の上限
        """

        await self.client.wait_until_ready()

        try:
            channel = self.client.get_channel(channel_id)
            if not channel:
                print(f"チャンネルID {channel_id} が見つかりません")
                return

            channel_name = getattr(channel, "name", f"Channel {channel.id}")
            print(f"チャンネル '{channel_name}' からメッセージを取得中...")

            messages_data = []
            message_count = 0

            # メッセージ履歴を取得
            if not hasattr(channel, "history"):
                print(f"チャンネル '{channel_name}' はメッセージ履歴を持っていません")
                return

            # 型チェック対応: channelを適切な型にキャスト
            from typing import cast

            import discord

            # メッセージ履歴を持つチャンネルの型を定義
            messageable_channel = cast(discord.abc.Messageable, channel)

            async for message in messageable_channel.history(
                limit=limit, after=after_date, before=before_date, oldest_first=False
            ):
                # 添付ファイルのURL取得
                attachments = []
                for attachment in message.attachments:
                    attachments.append(
                        {
                            "filename": attachment.filename,
                            "url": attachment.url,
                            "size": attachment.size,
                        }
                    )

                # リアクション情報
                reactions = []
                for reaction in message.reactions:
                    reactions.append(
                        {"emoji": str(reaction.emoji), "count": reaction.count}
                    )

                # メンション情報
                mentions = [user.display_name for user in message.mentions]

                # メッセージデータを構築
                message_data = {
                    "timestamp": message.created_at.strftime("%Y-%m-%d %H:%M:%S"),
                    "author_name": message.author.display_name,
                    "author_id": str(message.author.id),
                    "content": message.content,
                    "message_id": str(message.id),
                    "channel_name": channel_name,
                    "edited_at": message.edited_at.strftime("%Y-%m-%d %H:%M:%S")
                    if message.edited_at
                    else "",
                    "reply_to": str(message.reference.message_id)
                    if message.reference
                    else "",
                    "attachments_count": len(attachments),
                    "attachments_urls": "; ".join([att["url"] for att in attachments]),
                    "reactions_count": len(reactions),
                    "reactions": "; ".join(
                        [f"{r['emoji']}({r['count']})" for r in reactions]
                    ),
                    "mentions": "; ".join(mentions),
                    "is_bot": message.author.bot,
                    "message_type": str(message.type),
                }

                messages_data.append(message_data)
                message_count += 1

                if message_count % 100 == 0:
                    print(f"取得済み: {message_count} メッセージ")

            # DataFrameに変換
            df = pd.DataFrame(messages_data)

            # 時系列順にソート（古い順）
            df = df.sort_values("timestamp")

            # XLSXファイルに保存
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                # メインシート
                df.to_excel(writer, sheet_name="Messages", index=False)

                # 統計シート
                stats_data = {
                    "メトリック": [
                        "総メッセージ数",
                        "ユニークユーザー数",
                        "ボットメッセージ数",
                        "添付ファイル数",
                        "リアクション付きメッセージ数",
                        "エクスポート日時",
                        "チャンネル名",
                    ],
                    "値": [
                        len(df),
                        df["author_name"].nunique(),
                        len(df[df["is_bot"] == True]),
                        df["attachments_count"].sum(),
                        len(df[df["reactions_count"] > 0]),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        channel_name,
                    ],
                }

                stats_df = pd.DataFrame(stats_data)
                stats_df.to_excel(writer, sheet_name="Statistics", index=False)

                # ユーザー別統計
                user_stats = (
                    df.groupby("author_name")
                    .agg(
                        {
                            "message_id": "count",
                            "attachments_count": "sum",
                            "reactions_count": "sum",
                        }
                    )
                    .rename(
                        {
                            "message_id": "メッセージ数",
                            "attachments_count": "添付ファイル数",
                            "reactions_count": "リアクション数",
                        },
                        axis=1,
                    )
                    .reset_index()
                )

                user_stats.to_excel(writer, sheet_name="User_Statistics", index=False)

            print("✅ エクスポート完了!")
            print(f"   ファイル: {output_file}")
            print(f"   メッセージ数: {len(df)}")
            print(f"   ユーザー数: {df['author_name'].nunique()}")
            
            # エクスポート完了後の継続確認
            return True

        except Exception as e:
            print(f"❌ エラー: {e}")
            return False

        finally:
            await self.client.close()

    async def export_multiple_channels(
        self,
        selected_channels,
        output_file,
        after_date=None,
        before_date=None,
        limit=None,
    ):
        """
        複数チャンネルを一つのExcelファイルにエクスポート
        """
        await self.client.wait_until_ready()

        all_messages_data = []

        try:
            for channel_info in selected_channels:
                channel_id = channel_info["channel_id"]
                channel_name = channel_info["channel_name"]
                guild_name = channel_info["guild_name"]

                print(f"\n🔄 [{guild_name}] #{channel_name} を処理中...")

                channel = self.client.get_channel(channel_id)
                if not channel:
                    print(f"❌ チャンネルID {channel_id} が見つかりません")
                    continue

                if not hasattr(channel, "history"):
                    print(
                        f"❌ チャンネル '{channel_name}' はメッセージ履歴を持っていません"
                    )
                    continue

                # 型チェック対応
                from typing import cast

                import discord

                messageable_channel = cast(discord.abc.Messageable, channel)

                message_count = 0
                
                # パラメータの型安全性を確保
                safe_limit = None if limit is None or limit == "" else int(limit) if str(limit).isdigit() else None
                safe_after = after_date if after_date is not None else None
                safe_before = before_date if before_date is not None else None
                
                print(f"  メッセージ取得を開始...")
                print(f"    limit: {safe_limit} (type: {type(safe_limit)})")
                print(f"    after: {safe_after} (type: {type(safe_after)})")
                print(f"    before: {safe_before} (type: {type(safe_before)})")
                
                try:
                    # 段階的にパラメータを追加してエラーを特定
                    print("    基本的なhistory()を試行...")
                    
                    # 最もシンプルな形から開始
                    if safe_limit is None and safe_after is None and safe_before is None:
                        print("    パラメータなしで実行")
                        message_iter = messageable_channel.history()
                    elif safe_after is None and safe_before is None:
                        print(f"    limit={safe_limit}のみで実行")
                        message_iter = messageable_channel.history(limit=safe_limit)
                    else:
                        print(f"    全パラメータで実行")
                        message_iter = messageable_channel.history(
                            limit=safe_limit,
                            after=safe_after,
                            before=safe_before
                        )
                    
                    async for message in message_iter:
                        try:
                            # メッセージの基本情報をデバッグ出力
                            if message_count < 5:  # 最初の5件だけ詳細出力
                                print(f"    メッセージ{message_count}: ID={message.id} (type: {type(message.id)})")
                                print(f"      created_at: {message.created_at} (type: {type(message.created_at)})")
                                print(f"      author: {message.author.display_name} (id: {message.author.id}, type: {type(message.author.id)})")
                            
                            # 添付ファイル情報
                            attachments = []
                            for attachment in message.attachments:
                                attachments.append(
                                    {
                                        "filename": attachment.filename,
                                        "url": attachment.url,
                                        "size": attachment.size,
                                    }
                                )
    
                            # リアクション情報
                            reactions = []
                            for reaction in message.reactions:
                                reactions.append(
                                    {"emoji": str(reaction.emoji), "count": reaction.count}
                                )
    
                            # メンション情報
                            mentions = [user.display_name for user in message.mentions]
    
                            # メッセージデータ
                            message_data = {
                                "timestamp": message.created_at.strftime("%Y-%m-%d %H:%M:%S"),
                                "author_name": message.author.display_name,
                                "author_id": str(message.author.id),
                                "content": message.content,
                                "message_id": str(message.id),
                                "guild_name": guild_name,
                                "channel_name": channel_name,
                                "edited_at": message.edited_at.strftime("%Y-%m-%d %H:%M:%S")
                                if message.edited_at
                                else "",
                                "reply_to": str(message.reference.message_id)
                                if message.reference
                                else "",
                                "attachments_count": len(attachments),
                                "attachments_urls": "; ".join(
                                    [att["url"] for att in attachments]
                                ),
                                "reactions_count": len(reactions),
                                "reactions": "; ".join(
                                    [f"{r['emoji']}({r['count']})" for r in reactions]
                                ),
                                "mentions": "; ".join(mentions),
                                "is_bot": message.author.bot,
                                "message_type": str(message.type),
                            }
    
                            all_messages_data.append(message_data)
                            message_count += 1
    
                            if message_count % 100 == 0:
                                print(f"  取得済み: {message_count} メッセージ")
                                
                        except Exception as msg_error:
                            print(f"  メッセージ処理エラー (message_count={message_count}): {msg_error}")
                            print(f"    問題のメッセージID: {getattr(message, 'id', 'Unknown')}")
                            print(f"    created_at: {getattr(message, 'created_at', 'Unknown')} (type: {type(getattr(message, 'created_at', None))})")
                            continue
                            
                except Exception as history_error:
                    print(f"  メッセージ履歴取得エラー: {history_error}")
                    print(f"  エラーの種類: {type(history_error).__name__}")
                    
                    # 代替手法を試行
                    print("  代替手法で再試行...")
                    try:
                        # 最もシンプルなメッセージ取得
                        print("    シンプルなhistory(limit=10)で再試行")
                        simple_count = 0
                        async for simple_message in messageable_channel.history(limit=10):
                            try:
                                simple_data = {
                                    "timestamp": simple_message.created_at.strftime("%Y-%m-%d %H:%M:%S"),
                                    "author_name": simple_message.author.display_name,
                                    "author_id": str(simple_message.author.id),
                                    "content": simple_message.content,
                                    "message_id": str(simple_message.id),
                                    "guild_name": guild_name,
                                    "channel_name": channel_name,
                                    "edited_at": "",
                                    "reply_to": "",
                                    "attachments_count": 0,
                                    "attachments_urls": "",
                                    "reactions_count": 0,
                                    "reactions": "",
                                    "mentions": "",
                                    "is_bot": simple_message.author.bot,
                                    "message_type": str(simple_message.type),
                                }
                                all_messages_data.append(simple_data)
                                simple_count += 1
                            except Exception as simple_error:
                                print(f"    シンプルメッセージ処理エラー: {simple_error}")
                                continue
                        
                        print(f"  代替手法で{simple_count}件取得成功")
                        message_count = simple_count
                        
                    except Exception as fallback_error:
                        print(f"  代替手法も失敗: {fallback_error}")
                        print(f"  チャンネル '{channel_name}' をスキップします")
                        continue

                print(f"✅ {channel_name}: {message_count} メッセージ取得完了")

            if not all_messages_data:
                print("❌ エクスポートするメッセージがありません")
                return False

            print(f"\nデータ処理を開始... (総メッセージ数: {len(all_messages_data)})")
            
            # DataFrameに変換
            try:
                df = pd.DataFrame(all_messages_data)
                print(f"DataFrame作成完了: {len(df)} rows, {len(df.columns)} columns")
                
                # ソート前にデータ型を確認
                print(f"timestamp列のデータ型: {df['timestamp'].dtype}")
                if len(df) > 0:
                    print(f"timestampサンプル: {df['timestamp'].iloc[0]} (type: {type(df['timestamp'].iloc[0])})")
                
                # ソート処理
                print("DataFrameをソート中...")
                df = df.sort_values(["guild_name", "channel_name", "timestamp"])
                print("ソート完了")
                
            except Exception as df_error:
                print(f"データ処理エラー: {df_error}")
                print(f"エラーの種類: {type(df_error).__name__}")
                
                # データの一部を出力してデバッグ
                if all_messages_data:
                    print("最初のメッセージデータ:")
                    first_msg = all_messages_data[0]
                    for key, value in first_msg.items():
                        print(f"  {key}: {value} (type: {type(value)})")
                
                return False

            # Excelファイルに保存
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                # メインシート
                df.to_excel(writer, sheet_name="All_Messages", index=False)

                # チャンネル別統計
                channel_stats = (
                    df.groupby(["guild_name", "channel_name"])
                    .agg(
                        {
                            "message_id": "count",
                            "author_name": "nunique",
                            "attachments_count": "sum",
                            "reactions_count": "sum",
                        }
                    )
                    .rename(
                        {
                            "message_id": "メッセージ数",
                            "author_name": "ユニークユーザー数",
                            "attachments_count": "添付ファイル数",
                            "reactions_count": "リアクション数",
                        },
                        axis=1,
                    )
                    .reset_index()
                )

                channel_stats.to_excel(
                    writer, sheet_name="Channel_Statistics", index=False
                )

                # ユーザー別統計
                user_stats = (
                    df.groupby(["guild_name", "channel_name", "author_name"])
                    .agg(
                        {
                            "message_id": "count",
                            "attachments_count": "sum",
                            "reactions_count": "sum",
                        }
                    )
                    .rename(
                        {
                            "message_id": "メッセージ数",
                            "attachments_count": "添付ファイル数",
                            "reactions_count": "リアクション数",
                        },
                        axis=1,
                    )
                    .reset_index()
                )

                user_stats.to_excel(writer, sheet_name="User_Statistics", index=False)

                # 全体統計
                total_stats = {
                    "メトリック": [
                        "総メッセージ数",
                        "総チャンネル数",
                        "総ユーザー数",
                        "総添付ファイル数",
                        "リアクション付きメッセージ数",
                        "エクスポート日時",
                    ],
                    "値": [
                        len(df),
                        df["channel_name"].nunique(),
                        df["author_name"].nunique(),
                        df["attachments_count"].sum(),
                        len(df[df["reactions_count"] > 0]),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ],
                }

                total_stats_df = pd.DataFrame(total_stats)
                total_stats_df.to_excel(
                    writer, sheet_name="Total_Statistics", index=False
                )

            print("\n✅ エクスポート完了!")
            print(f"   ファイル: {output_file}")
            print(f"   総メッセージ数: {len(df):,}")
            print(f"   チャンネル数: {df['channel_name'].nunique()}")
            print(f"   ユーザー数: {df['author_name'].nunique()}")
            
            # エクスポート完了後の継続確認
            return True

        except Exception as e:
            print(f"❌ エラー: {e}")
            print(f"エラーの種類: {type(e).__name__}")
            import traceback
            print(f"スタックトレース: {traceback.format_exc()}")
            return False
        finally:
            try:
                if hasattr(self.client, '_ready') and self.client._ready.is_set():
                    print("接続をクローズ中...")
                    await self.client.close()
                    print("接続クローズ完了")
            except Exception as close_error:
                print(f"接続クローズエラー: {close_error}")
                pass


async def main():
    try:
        parser = argparse.ArgumentParser(description="Discord Chat to XLSX Exporter")
        parser.add_argument("-t", "--token", help="Discord Bot Token")
        parser.add_argument(
            "--fetch-channels", action="store_true", help="Fetch and save all channels"
        )
        parser.add_argument(
            "--interactive", action="store_true", help="Interactive channel selection (TUI)"
        )
        parser.add_argument(
            "--cli", action="store_true", help="Interactive channel selection (CLI mode)"
        )
        parser.add_argument(
            "-c", "--channel", type=int, help="Single channel ID (legacy mode)"
        )
        parser.add_argument("-o", "--output", help="Output XLSX file")
        parser.add_argument("--after", help="After date (YYYY-MM-DD)")
        parser.add_argument("--before", help="Before date (YYYY-MM-DD)")
        parser.add_argument("--limit", type=int, help="Message limit per channel")
        parser.add_argument("--config", action="store_true", help="Launch configuration UI")

        args = parser.parse_args()

        # 引数が何も指定されていない場合、または--configが指定された場合はメインメニューを起動
        if not any(vars(args).values()) or args.config:
            try:
                print("Discord Exporter を起動中...")
                
                # メインメニューループ
                temp_exporter = DiscordExporter("")
                
                while True:
                    try:
                        if args.config:
                            # 設定UIを直接起動
                            action = "config"
                            args.config = False  # 一度だけ実行
                        else:
                            # メインメニューを表示
                            action = temp_exporter.main_menu()
                        
                        if action == "exit":
                            print("終了します。")
                            return
                        
                        elif action == "update_channels":
                            # チャンネル情報を更新
                            config = temp_exporter.load_config()
                            if not config.get("token"):
                                print("Bot Tokenが設定されていません。設定を変更してください。")
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                                continue
                            
                            print("チャンネル情報を更新中...")
                            temp_exporter.token = config["token"]
                            
                            # Discord接続してチャンネル情報を更新
                            @temp_exporter.client.event
                            async def on_ready():
                                print(f"ログイン: {temp_exporter.client.user}")
                                await temp_exporter.fetch_and_save_channels()
                            
                            try:
                                await temp_exporter.client.start(config["token"])
                                
                                # 更新完了後の継続確認
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                                    
                            except KeyboardInterrupt:
                                print("\n操作がキャンセルされました。")
                                await temp_exporter.cleanup_client()
                                return
                            except Exception as e:
                                print(f"エラー: {e}")
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                            
                            # 新しいクライアントインスタンスを作成（接続をリセット）
                            temp_exporter = DiscordExporter("")
                            continue
                        
                        elif action == "export_interactive":
                            # エクスポート実行
                            config = temp_exporter.load_config()
                            if not config.get("token"):
                                print("Bot Tokenが設定されていません。設定を変更してください。")
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                                continue
                            
                            channels_info = temp_exporter.get_channels_info()
                            if not channels_info:
                                print("チャンネル情報がありません。最初にチャンネル情報を更新してください。")
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                                continue
                            
                            # エクスポート実行
                            try:
                                selected_channels = temp_exporter.select_channels_interactive(use_tui=True)
                                
                                if not selected_channels:
                                    print("チャンネルが選択されていません。")
                                    if not temp_exporter.ask_continue():
                                        await temp_exporter.cleanup_client()
                                        return
                                    continue
                                
                                # エクスポート実行
                                temp_exporter.token = config["token"]
                                
                                after_date = None
                                before_date = None
                                
                                if config.get("after_date"):
                                    try:
                                        after_date = datetime.strptime(config["after_date"], "%Y-%m-%d").replace(tzinfo=timezone.utc)
                                    except ValueError:
                                        pass
                                
                                if config.get("before_date"):
                                    try:
                                        before_date = datetime.strptime(config["before_date"], "%Y-%m-%d").replace(tzinfo=timezone.utc)
                                    except ValueError:
                                        pass
                                
                                @temp_exporter.client.event
                                async def on_ready():
                                    print(f"ログイン: {temp_exporter.client.user}")
                                    return await temp_exporter.export_multiple_channels(
                                        selected_channels, config["output_file"], after_date, before_date, config.get("limit")
                                    )
                                
                                await temp_exporter.client.start(config["token"])
                                
                                # エクスポート完了後の継続確認
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                                    
                            except KeyboardInterrupt:
                                print("\n操作がキャンセルされました。")
                                await temp_exporter.cleanup_client()
                                return
                            except Exception as e:
                                print(f"エラー: {e}")
                                if not temp_exporter.ask_continue():
                                    await temp_exporter.cleanup_client()
                                    return
                            
                            # 新しいクライアントインスタンスを作成
                            temp_exporter = DiscordExporter("")
                            continue
                        
                        elif action == "config":
                            # 設定UI
                            try:
                                config = temp_exporter.config_ui()
                                if config:
                                    print("設定が保存されました。メニューに戻ります...")
                                else:
                                    print("設定がキャンセルされました。メニューに戻ります...")
                                    
                            except KeyboardInterrupt:
                                print("\n設定がキャンセルされました。メニューに戻ります...")
                            
                            # 設定完了後は自動的にメニューに戻る
                            continue
                            
                    except KeyboardInterrupt:
                        print("\n操作がキャンセルされました。")
                        if 'temp_exporter' in locals():
                            await temp_exporter.cleanup_client()
                        return
                        
            except KeyboardInterrupt:
                print("\nプログラムを終了します。")
                if 'temp_exporter' in locals():
                    await temp_exporter.cleanup_client()
                return
        else:
            # 必須項目チェック
            if not args.token:
                print("エラー: Discord Bot Tokenが必要です。")
                print("--config オプションで設定UIを使用するか、-t でTokenを指定してください。")
                return

            if not args.output and not args.fetch_channels:
                print("エラー: 出力ファイル名が必要です。")
                return

            # 日付の解析
            after_date = None
            before_date = None

            if args.after:
                try:
                    after_date = datetime.strptime(args.after, "%Y-%m-%d").replace(
                        tzinfo=timezone.utc
                    )
                except ValueError:
                    print(f"エラー: 無効な開始日形式: {args.after}")
                    return

            if args.before:
                try:
                    before_date = datetime.strptime(args.before, "%Y-%m-%d").replace(
                        tzinfo=timezone.utc
                    )
                except ValueError:
                    print(f"エラー: 無効な終了日形式: {args.before}")
                    return

            # エクスポーター実行
            exporter = DiscordExporter(args.token)

            if args.fetch_channels:
                # チャンネル情報を取得してJSONに保存
                @exporter.client.event
                async def on_ready():
                    print(f"ログイン: {exporter.client.user}")
                    await exporter.fetch_and_save_channels()

                await exporter.client.start(args.token)

            elif args.interactive:
                # インタラクティブモード（TUI）
                selected_channels = exporter.select_channels_interactive(use_tui=True)

                if not selected_channels:
                    print("チャンネルが選択されていません。")
                    return

                @exporter.client.event
                async def on_ready():
                    print(f"ログイン: {exporter.client.user}")
                    await exporter.export_multiple_channels(
                        selected_channels, args.output, after_date, before_date, args.limit
                    )

                await exporter.client.start(args.token)
            
            elif args.cli:
                # インタラクティブモード（CLI）
                selected_channels = exporter.select_channels_interactive(use_tui=False)

                if not selected_channels:
                    print("チャンネルが選択されていません。")
                    return

                @exporter.client.event
                async def on_ready():
                    print(f"ログイン: {exporter.client.user}")
                    await exporter.export_multiple_channels(
                        selected_channels, args.output, after_date, before_date, args.limit
                    )

                await exporter.client.start(args.token)

            elif args.channel:
                # 従来のシングルチャンネルモード
                @exporter.client.event
                async def on_ready():
                    print(f"ログイン: {exporter.client.user}")
                    await exporter.export_channel_to_xlsx(
                        args.channel, args.output, after_date, before_date, args.limit
                    )

                await exporter.client.start(args.token)

            else:
                print("使用方法:")
                print("  1. メインメニュー: python discord_exporter.py")
                print("  2. 設定UI: python discord_exporter.py --config")
                print("  3. 従来通り: python discord_exporter.py -t TOKEN --interactive -o FILE")
                parser.print_help()

    except KeyboardInterrupt:
        print("\nプログラムを終了します。")
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nプログラムを終了します。")
    except Exception as e:
        print(f"予期しないエラーが発生しました: {e}")

# 使用例:
# python discord_exporter.py -t "YOUR_BOT_TOKEN" -c 123456789 -o "chat.xlsx"
# python discord_exporter.py -t "YOUR_BOT_TOKEN" -c 123456789 -o "chat.xlsx" --after 2024-01-01 --before 2024-12-31
# python discord_exporter.py -t "YOUR_BOT_TOKEN" -c 123456789 -o "chat.xlsx" --limit 1000
