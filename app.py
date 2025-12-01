import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import subprocess
import csv
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build
import threading
import sys
import json
from datetime import datetime

# ----- UI設定 -----
BG_COLOR = "#F0F5FF"
PRIMARY_COLOR = "#6B8E23"
ACCENT_COLOR = "#87CEEB"
FONT_FAMILY = "Yu Gothic UI"

class AmazonRankingApp(tk.Tk):
    def __init__(self):
        super().__init__()

        # Window dimensions
        window_width = 830
        window_height = 650

        self.title("アマゾンランキング取得ツール")
        self.geometry(f"{window_width}x{window_height}")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)

        # Get screen width and height
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Calculate x and y coordinates for the center
        center_x = int(screen_width / 2 - window_width / 2)
        center_y = int(screen_height / 2 - window_height / 2)

        # Set the window position
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")
        
        self.create_widgets()

    def create_widgets(self):
        # メインフレーム
        main_frame = tk.Frame(self, bg=BG_COLOR, padx=20, pady=20)
        main_frame.pack(expand=True, fill="both")

        # タイトル
        title_label = tk.Label(
            main_frame, 
            text="アマゾンランキング取得ツール", 
            font=(FONT_FAMILY, 24, "bold"), 
            bg=BG_COLOR, 
            fg=PRIMARY_COLOR
        )
        title_label.pack(pady=(0, 10))

        # 説明
        description_label = tk.Label(
            main_frame,
            text="このツールは、keywords.csv および asins.csv ファイルに記載されたデータを読み込み、対応する商品のランキングを取得します。\n取得した結果は、自動的に指定されたスプレッドシートに保存されるため、データ管理や分析を効率的に行うことができます。",
            font=(FONT_FAMILY, 12),
            bg=BG_COLOR,
            justify="left"
        )
        description_label.pack(pady=(0, 20), anchor="w")

        # スプレッドシートの指定フレーム
        spreadsheet_frame = tk.Frame(main_frame, bg=BG_COLOR)
        spreadsheet_frame.pack(fill="x", pady=(0, 20))

        spreadsheet_label = tk.Label(
            spreadsheet_frame,
            text="📝 スプレッドシートの指定",
            font=(FONT_FAMILY, 12, "bold"),
            bg=BG_COLOR,
            fg=ACCENT_COLOR
        )
        spreadsheet_label.pack(anchor="w")

        # ドロップダウンフレーム
        dropdown_frame = tk.Frame(spreadsheet_frame, bg=BG_COLOR, pady=10)
        dropdown_frame.pack(fill="x")

        # ID ドロップダウン
        id_frame = tk.Frame(dropdown_frame, bg=BG_COLOR)
        id_frame.pack(side="left", padx=(0, 20))
        id_label = tk.Label(id_frame, text="ID", font=(FONT_FAMILY, 12, "bold"), bg=BG_COLOR)
        id_label.pack(anchor="w")
        self.id_dropdown = ttk.Combobox(id_frame, state="readonly", width=60)
        self.id_dropdown.pack(pady=(5, 0))
        # CSV から ID を読み込み
        spreadsheet_ids = self.load_spreadsheet_ids()
        if spreadsheet_ids:
            self.id_dropdown["values"] = tuple(["IDを選択してください", *spreadsheet_ids])
            self.id_dropdown.set("IDを選択してください")
        else:
            self.id_dropdown["values"] = ("IDを選択してください",)
            self.id_dropdown.set("IDを選択してください")
        # ID 選択時にシート一覧を読み込む
        self.id_dropdown.bind("<<ComboboxSelected>>", self.load_sheets_for_selected_id)

        # シート ドロップダウン
        sheet_frame = tk.Frame(dropdown_frame, bg=BG_COLOR)
        sheet_frame.pack(side="left")
        sheet_label = tk.Label(sheet_frame, text="シート", font=(FONT_FAMILY, 12, "bold"), bg=BG_COLOR)
        sheet_label.pack(anchor="w")
        self.sheet_dropdown = ttk.Combobox(sheet_frame, state="readonly", width=60)
        self.sheet_dropdown.pack(pady=(5, 0))
        # 初期値（ID 選択後に更新）
        self.sheet_dropdown["values"] = ("シートを選択してください",)
        self.sheet_dropdown.set("シートを選択してください")

        # 開始ボタン
        self.start_button = tk.Button(
            main_frame,
            text="🚀 開始",
            font=(FONT_FAMILY, 12, "bold"),
            bg=ACCENT_COLOR,
            fg="white",
            bd=0,
            relief="flat",
            command=self.start_scraping
        )
        self.start_button.pack(pady=(8, 8), fill="x", ipady=5)
        
        # エラーメッセージラベル
        self.error_label = tk.Label(
            main_frame, 
            text="", 
            font=(FONT_FAMILY, 12, "bold"), 
            bg=BG_COLOR, 
            fg="red"
        )
        self.error_label.pack(pady=(0, 10))

        # ログフレーム
        log_frame = tk.Frame(main_frame, bg="white", relief="groove", bd=2)
        log_frame.pack(fill="both", expand=True)

        log_label = tk.Label(
            log_frame,
            text="📜 ログ",
            font=(FONT_FAMILY, 12, "bold"),
            bg="white",
            fg=ACCENT_COLOR
        )
        log_label.pack(anchor="w", padx=10, pady=5)
        
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled", font=(FONT_FAMILY, 10))
        self.log_text.pack(expand=True, fill="both", padx=10, pady=(0, 10))

    def get_google_sheets_service(self):
        """Google Sheets API サービスクライアントを返す。"""
        # Use executable directory when frozen so users can replace the JSON
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        creds_path = os.path.join(base_dir, "weighty-vertex-464012-u4-7cd9bab1166b.json")
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        credentials = service_account.Credentials.from_service_account_file(creds_path, scopes=scopes)
        return build("sheets", "v4", credentials=credentials)

    def fetch_sheet_titles(self, spreadsheet_id):
        """スプレッドシート ID からシート名一覧を取得して返す。"""
        service = self.get_google_sheets_service()
        meta = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = meta.get("sheets", [])
        return [s["properties"]["title"] for s in sheets if "properties" in s and "title" in s["properties"]]

    def load_sheets_for_selected_id(self, event=None):
        selected_id = self.id_dropdown.get()
        if selected_id == "IDを選択してください":
            self.sheet_dropdown["values"] = ("シートを選択してください",)
            self.sheet_dropdown.set("シートを選択してください")
            return

        # 先にログを表示し、バックグラウンドで取得
        self.update_log("▶️ シート一覧を取得中...")
        threading.Thread(target=self._fetch_and_set_sheets, args=(selected_id,), daemon=True).start()

    def _fetch_and_set_sheets(self, spreadsheet_id):
        try:
            titles = self.fetch_sheet_titles(spreadsheet_id)
            if titles:
                def _apply_success():
                    self.sheet_dropdown["values"] = tuple(["シートを選択してください", *titles])
                    self.sheet_dropdown.set("シートを選択してください")
                    self.update_log(f"✅ {len(titles)} 件のシートを取得しました。")
                self.after(0, _apply_success)
            else:
                def _apply_empty():
                    self.sheet_dropdown["values"] = ("シートを選択してください",)
                    self.sheet_dropdown.set("シートを選択してください")
                    self.update_log("⚠️ シートが見つかりませんでした。")
                self.after(0, _apply_empty)
        except FileNotFoundError:
            self.after(0, self.update_log, "エラー: 認証ファイルが見つかりません。weighty-vertex-464012-u4-7cd9bab1166b.json を配置してください。")
            self.after(0, messagebox.showerror, "エラー", "認証ファイルが見つかりません。weighty-vertex-464012-u4-7cd9bab1166b.json を同じフォルダに配置してください。")
        except Exception as e:
            self.after(0, self.update_log, f"シート取得中にエラーが発生しました: {e}")
            self.after(0, messagebox.showerror, "エラー", f"シート取得中にエラーが発生しました: {e}")

    def load_spreadsheet_ids(self):
        """spreadsheetIDs.csv の先頭列から ID を読み込み、リストで返す。"""
        ids = []
        try:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            csv_path = os.path.join(base_dir, "spreadsheetIDs.csv")
            with open(csv_path, newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    if not row:
                        continue
                    value = row[0].strip()
                    if not value:
                        continue
                    header_like = {"id", "ids", "spreadsheetid", "spreadsheet_id"}
                    if value.lower() in header_like:
                        continue
                    ids.append(value)
        except FileNotFoundError:
            # ファイルが無い場合は空リストを返す（UI 側でプレースホルダのみ表示）
            pass
        except Exception as e:
            # 予期しないエラーはログに出す
            try:
                self.update_log(f"ID 読み込み中にエラーが発生しました: {e}")
            except Exception:
                pass
        return ids

    def start_scraping(self):
        self.error_label.config(text="")
        selected_id = self.id_dropdown.get()
        selected_sheet = self.sheet_dropdown.get()

        if selected_id == "IDを選択してください" or selected_sheet == "シートを選択してください":
            self.error_label.config(text="⚠️ スプレッドシートを選択してください。")
            return

        # 開始ボタンを無効化
        self.start_button.config(state="disabled", text="🔄 実行中...")

        self.update_log("スクリプトを開始します...")
        self.update_log(f"選択されたID: {selected_id}")
        self.update_log(f"選択されたシート: {selected_sheet}")

        # scrap.py をバックグラウンドで実行し、出力を逐次ログに反映
        thread = threading.Thread(target=self._run_scrap_and_stream_logs, args=(selected_id, selected_sheet), daemon=True)
        thread.start()

    def _run_scrap_and_stream_logs(self, spreadsheet_id, sheet_name):
        try:
            base_dir_script = os.path.dirname(os.path.abspath(__file__))
            # When frozen (exe), re-invoke the same executable with a special flag to run scraping
            if getattr(sys, 'frozen', False):
                cmd = [sys.executable, "--run-scrap", "--spreadsheet-id", spreadsheet_id, "--sheet", sheet_name]
                run_cwd = os.path.dirname(sys.executable)
            else:
                script_path = os.path.join(base_dir_script, "scrap.py")
                if not os.path.exists(script_path):
                    self.after(0, self.update_log, "エラー: scrap.py が見つかりません。")
                    self.after(0, messagebox.showerror, "エラー", "scrap.py が見つかりません。ファイルが同じディレクトリにあるか確認してください。")
                    return
                cmd = [sys.executable, "-u", script_path, "--spreadsheet-id", spreadsheet_id, "--sheet", sheet_name]
                run_cwd = base_dir_script
            env = os.environ.copy()
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUNBUFFERED"] = "1"
            env["PYTHONLEGACYWINDOWSSTDIO"] = "1"
            proc = subprocess.Popen(
                cmd,
                cwd=run_cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=1,
                universal_newlines=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )

            result_data = None
            if proc.stdout is not None:
                while True:
                    line = proc.stdout.readline()
                    if not line:
                        if proc.poll() is not None:
                            break
                        continue
                    line = line.rstrip("\n")
                    if line:
                        # 結果データをキャプチャ
                        if line.startswith("RESULT_DATA:"):
                            try:
                                result_json = line[12:]  # "RESULT_DATA:" を除去
                                result_data = json.loads(result_json)
                                self.after(0, self.update_log, f"📊 結果データを取得しました: {len(result_data)}件")
                            except json.JSONDecodeError as e:
                                self.after(0, self.update_log, f"⚠️ 結果データの解析に失敗: {e}")
                        else:
                            self.after(0, self.update_log, line)

            return_code = proc.wait()
            if return_code == 0:
                self.after(0, self.update_log, "✅ スクレイピングが完了しました。")
                # 結果データをスプレッドシートに書き込み
                if result_data:
                    self.after(0, self._write_to_spreadsheet, spreadsheet_id, sheet_name, result_data)
                else:
                    self.after(0, self.update_log, "⚠️ 結果データがありません。")
                    # 結果データがない場合もボタンを再有効化
                    self.after(0, self._enable_start_button)
            else:
                self.after(0, self.update_log, f"⚠️ scrap.py が異常終了しました (exit {return_code})")
                # スクレイピングが異常終了した場合もボタンを再有効化
                self.after(0, self._enable_start_button)
        except Exception as e:
            self.after(0, self.update_log, f"実行中にエラーが発生しました: {e}")
            self.after(0, messagebox.showerror, "エラー", f"実行中にエラーが発生しました: {e}")
            # エラーが発生した場合もボタンを再有効化
            self.after(0, self._enable_start_button)

    def _write_to_spreadsheet(self, spreadsheet_id, sheet_name, result_data):
        """スプレッドシートに結果データを書き込む"""
        try:
            self.update_log("📝 スプレッドシートに書き込み中...")
            
            # バックグラウンドでスプレッドシート操作を実行
            thread = threading.Thread(
                target=self._write_to_spreadsheet_thread, 
                args=(spreadsheet_id, sheet_name, result_data), 
                daemon=True
            )
            thread.start()
            
        except Exception as e:
            self.update_log(f"スプレッドシート書き込みエラー: {e}")
            messagebox.showerror("エラー", f"スプレッドシート書き込みエラー: {e}")

    def _write_to_spreadsheet_thread(self, spreadsheet_id, sheet_name, result_data):
        """スプレッドシート書き込みのバックグラウンド処理"""
        try:
            service = self.get_google_sheets_service()
            
            # 現在の日付を取得
            current_date = datetime.now().strftime("%Y/%m/%d")
            
            # キーワードを読み込み
            keywords = self.load_keywords()
            if not keywords:
                self.after(0, self.update_log, "⚠️ keywords.csv が見つかりません。")
                return
            
            # スプレッドシートの現在の内容を取得
            range_name = f"{sheet_name}!A1:Z100"  # 十分な範囲を指定
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, 
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            # ヘッダーを検証・設定
            if not self._validate_and_set_headers(service, spreadsheet_id, sheet_name, values, keywords):
                self.after(0, self.update_log, "⚠️ ヘッダーの設定に失敗しました。")
                return
            
            # データ行を追加/更新
            self._add_or_update_data_row(service, spreadsheet_id, sheet_name, current_date, result_data, keywords)
            
            self.after(0, self.update_log, "✅ スプレッドシートへの書き込みが完了しました。")
            # スプレッドシート書き込み完了後にボタンを再有効化
            self.after(0, self._enable_start_button)
            
        except Exception as e:
            self.after(0, self.update_log, f"スプレッドシート書き込みエラー: {e}")
            self.after(0, messagebox.showerror, "エラー", f"スプレッドシート書き込みエラー: {e}")
            # エラーが発生した場合もボタンを再有効化
            self.after(0, self._enable_start_button)

    def load_keywords(self):
        """keywords.csvからキーワードを読み込む"""
        keywords = []
        try:
            if getattr(sys, 'frozen', False):
                base_dir = os.path.dirname(sys.executable)
            else:
                base_dir = os.path.dirname(os.path.abspath(__file__))
            csv_path = os.path.join(base_dir, "keywords.csv")
            with open(csv_path, newline='', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                for row in reader:
                    if not row:
                        continue
                    value = row[0].strip().lstrip('\ufeff')
                    if not value:
                        continue
                    header_like = {"keyword", "keywords", "キーワード"}
                    if value.lower() in header_like:
                        continue
                    keywords.append(value)
        except FileNotFoundError:
            pass
        except Exception as e:
            self.after(0, self.update_log, f"キーワード読み込みエラー: {e}")
        return keywords

    def _validate_and_set_headers(self, service, spreadsheet_id, sheet_name, values, keywords):
        """ヘッダーを検証し、必要に応じて設定する"""
        try:
            # キーワード数に基づいて列数を計算
            keyword_count = len(keywords)
            total_columns = 1 + (keyword_count * 3)  # 日付列 + (キーワード数 × 3カテゴリ)
            
            # 期待するヘッダー構造を動的に生成
            expected_header1 = ["日付/キーワード"]
            expected_header2 = [""]
            
            # 各カテゴリ（自然検索、SP、SB）のヘッダーを生成
            categories = ["自然検索", "SP", "SB"]
            for category in categories:
                expected_header1.append(category)
                expected_header2.extend(keywords)  # 各カテゴリに全キーワードを配置
                # カテゴリ名の後の空列を追加
                for _ in range(keyword_count - 1):
                    expected_header1.append("")
            
            # 現在のヘッダーをチェック
            if len(values) >= 2:
                current_header1 = values[0][:total_columns] if len(values[0]) >= total_columns else values[0]
                current_header2 = values[1][:total_columns] if len(values[1]) >= total_columns else values[1]
                
                # ヘッダーが一致するかチェック
                if (current_header1 == expected_header1 and 
                    current_header2 == expected_header2):
                    return True
            
            # ヘッダーが一致しない場合はクリアして再設定
            self.after(0, self.update_log, "🔄 ヘッダーを再設定します...")
            
            # シート全体をクリア
            service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:Z"
            ).execute()
            
            # 新しいヘッダーを設定
            header1 = expected_header1
            header2 = expected_header2
            
            # ヘッダーを書き込み
            end_column = chr(ord('A') + total_columns - 1)  # 最後の列を計算
            body = {
                'values': [header1, header2]
            }
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A1:{end_column}2",
                valueInputOption='RAW',
                body=body
            ).execute()
            
            return True
            
        except Exception as e:
            self.after(0, self.update_log, f"ヘッダー設定エラー: {e}")
            return False

    def _add_or_update_data_row(self, service, spreadsheet_id, sheet_name, current_date, result_data, keywords):
        """データ行を追加または更新する"""
        try:
            # 既存のデータを取得
            range_name = f"{sheet_name}!A3:Z100"
            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, 
                range=range_name
            ).execute()
            
            values = result.get('values', [])
            
            # 同じ日付の行を探す
            row_index = None
            for i, row in enumerate(values):
                if row and len(row) > 0 and row[0] == current_date:
                    row_index = i + 3  # 1ベースの行番号（ヘッダーが2行あるため+3）
                    break
            
            # キーワード数に基づいて列数を計算
            keyword_count = len(keywords)
            total_columns = 1 + (keyword_count * 3)  # 日付列 + (キーワード数 × 3カテゴリ)
            
            # 結果データを整理（動的な列数のデータ行を作成）
            # 各カテゴリ（自然検索、SP、SB）にキーワード数の結果を配置
            organic_results = ["-"] * keyword_count  # 自然検索のキーワード結果
            sp_results = ["-"] * keyword_count       # SPのキーワード結果
            sb_results = ["-"] * keyword_count       # SBのキーワード結果
            
            # 結果データから各キーワードの結果を取得
            for item in result_data:
                # BOM文字を除去して比較
                result_keyword = item["keyword"].lstrip('\ufeff').strip()
                
                # キーワードのインデックスを取得
                keyword_index = None
                for i, kw in enumerate(keywords):
                    if result_keyword == kw.lstrip('\ufeff').strip():
                        keyword_index = i
                        break
                
                if keyword_index is not None:
                    # 各カテゴリの結果を設定
                    organic_results[keyword_index] = item["自然検索"]
                    sp_results[keyword_index] = item["SP"]
                    sb_results[keyword_index] = item["SB"]
            
            # データ行を作成（動的な列数）
            new_row = [current_date] + organic_results + sp_results + sb_results
            
            if row_index:
                # 既存行を更新
                end_column = chr(ord('A') + total_columns - 1)  # 最後の列を計算
                range_name = f"{sheet_name}!A{row_index}:{end_column}{row_index}"
                body = {'values': [new_row]}
                service.spreadsheets().values().update(
                    spreadsheetId=spreadsheet_id,
                    range=range_name,
                    valueInputOption='RAW',
                    body=body
                ).execute()
                self.after(0, self.update_log, f"📝 日付 {current_date} の行を更新しました。")
            else:
                # 新しい行を追加
                end_column = chr(ord('A') + total_columns - 1)  # 最後の列を計算
                body = {'values': [new_row]}
                service.spreadsheets().values().append(
                    spreadsheetId=spreadsheet_id,
                    range=f"{sheet_name}!A:{end_column}",
                    valueInputOption='RAW',
                    insertDataOption='INSERT_ROWS',
                    body=body
                ).execute()
                self.after(0, self.update_log, f"📝 日付 {current_date} の新しい行を追加しました。")
                
        except Exception as e:
            self.after(0, self.update_log, f"データ行追加エラー: {e}")
            raise

    def _enable_start_button(self):
        """開始ボタンを再有効化する"""
        self.start_button.config(state="normal", text="🚀 開始")

    def update_log(self, message):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

if __name__ == "__main__":
    # Support headless scrap mode when running the built exe to execute scraping in a child process
    if "--run-scrap" in sys.argv:
        # Lazy imports to avoid impacting the GUI startup path
        import argparse
        import asyncio
        from scrap import scraping

        parser = argparse.ArgumentParser()
        parser.add_argument("--run-scrap", action="store_true")
        parser.add_argument("--spreadsheet-id", required=True)
        parser.add_argument("--sheet", required=True)
        args = parser.parse_args()

        # Run scraping and stream prints to stdout for the parent GUI process to capture
        # Ensure UTF-8 stdout for proper Japanese logs
        try:
            import sys as _sys
            if hasattr(_sys.stdout, "reconfigure"):
                _sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass
        asyncio.run(scraping(args.spreadsheet_id, args.sheet))
    else:
        app = AmazonRankingApp()
        app.mainloop()