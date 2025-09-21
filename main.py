"""
VOICE_CORRECTOR - 音声入力テキスト修正ツール

音声入力された文字列を文法的に正しい文章に変換するGUIアプリケーション
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import os
import pyperclip
import requests
import threading
import winsound
from typing import Dict, List, Optional
import glob
import ctypes
import platform


class VoiceCorrector:
    def __init__(self):
        # DPI対応の設定
        self.setup_dpi_awareness()
        
        self.root = tk.Tk()
        self.root.title("VOICE_CORRECTOR")
        
        # DPIスケーリングファクターを取得
        self.scale_factor = self.get_scale_factor()
        
        # ウィンドウサイズをスケーリングに合わせて調整
        base_width = 800
        base_height = 600
        scaled_width = int(base_width * self.scale_factor)
        scaled_height = int(base_height * self.scale_factor)
        self.root.geometry(f"{scaled_width}x{scaled_height}")
        
        # フォントサイズの調整
        self.setup_scaled_fonts()
        
        # 設定ファイルのパス
        self.config_file = "settings.json"
        self.reference_folder = "reference"
        
        # 設定の初期値
        self.settings = {
            "conversion_policy": "",
            "reference_text": "",
            "selected_reference_file": ""
        }
        
        # 参考用ファイルリスト
        self.reference_files = []
        
        # GUI構成要素の初期化
        self.setup_gui()
        
        # 設定の読み込み
        self.load_settings()
        
        # 参考用ファイルリストの更新
        self.update_reference_files()
        
    def setup_dpi_awareness(self):
        """DPI認識を設定"""
        try:
            if platform.system() == "Windows":
                # Windows 8.1以降でDPI認識を有効化
                ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            # 古いWindowsバージョンでは何もしない
            pass
    
    def get_scale_factor(self) -> float:
        """現在のDPIスケーリングファクターを取得"""
        try:
            if platform.system() == "Windows":
                # Windows DPIスケールを取得
                import tkinter as tk
                root = tk.Tk()
                root.withdraw()  # ウィンドウを表示しない
                dpi = root.winfo_fpixels('1i')
                root.destroy()
                # 標準DPI(96)との比率を計算
                scale_factor = dpi / 96.0
                # スケールファクターを1.0〜2.0の範囲に制限
                return max(1.0, min(2.0, scale_factor))
            else:
                return 1.0
        except Exception:
            return 1.0
    
    def setup_scaled_fonts(self):
        """スケーリングに対応したフォントを設定"""
        # 基本フォントサイズ
        base_font_size = 9
        scaled_font_size = int(base_font_size * self.scale_factor)
        
        # デフォルトフォントを設定
        default_font = ("Yu Gothic UI", scaled_font_size)
        self.root.option_add("*Font", default_font)
        
        # ttk.Styleでより詳細な設定
        style = ttk.Style()
        style.configure(".", font=default_font)
        style.configure("TLabel", font=default_font)
        style.configure("TButton", font=default_font)
        style.configure("TCombobox", font=default_font)
        
    def scale_size(self, size: int) -> int:
        """サイズをスケーリングファクターに基づいて調整"""
        return int(size * self.scale_factor)
        
    def setup_gui(self):
        """GUI要素のセットアップ"""
        # メインフレーム
        padding = self.scale_size(10)
        main_frame = ttk.Frame(self.root, padding=str(padding))
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # グリッドの重みを設定
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # 入力ボックス
        ttk.Label(main_frame, text="入力ボックス (改行3連続で自動変換):").grid(row=0, column=0, sticky=tk.W, pady=(0, self.scale_size(5)))
        input_height = self.scale_size(6)
        input_width = self.scale_size(70)
        self.input_text = scrolledtext.ScrolledText(main_frame, height=input_height, width=input_width)
        self.input_text.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, self.scale_size(10)))
        
        # 入力ボックスのキーイベントをバインド（改行3連続での自動変換）
        self.input_text.bind('<KeyRelease>', self.on_input_key_release)
        
        # 右クリックメニューを明示的に有効化
        self.setup_text_context_menu(self.input_text)
        
        # 変換の方針ボックス
        ttk.Label(main_frame, text="変換の方針:").grid(row=2, column=0, sticky=tk.W, pady=(0, self.scale_size(5)))
        policy_height = self.scale_size(3)
        self.policy_text = scrolledtext.ScrolledText(main_frame, height=policy_height, width=input_width)
        self.policy_text.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, self.scale_size(10)))
        
        # 右クリックメニューを設定
        self.setup_text_context_menu(self.policy_text)
        
        # 参考用ボックスセレクターと参考用ボックス
        ttk.Label(main_frame, text="参考用ファイル:").grid(row=4, column=0, sticky=tk.W, pady=(0, self.scale_size(5)))
        
        reference_frame = ttk.Frame(main_frame)
        reference_frame.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, self.scale_size(10)))
        reference_frame.columnconfigure(1, weight=1)
        
        combo_width = self.scale_size(30)
        self.reference_selector = ttk.Combobox(reference_frame, width=combo_width, state="readonly")
        self.reference_selector.grid(row=0, column=0, padx=(0, self.scale_size(10)))
        self.reference_selector.bind("<<ComboboxSelected>>", self.on_reference_selected)
        
        refresh_btn = ttk.Button(reference_frame, text="更新", command=self.update_reference_files)
        refresh_btn.grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(main_frame, text="参考用ボックス:").grid(row=6, column=0, sticky=tk.W, pady=(0, self.scale_size(5)))
        reference_height = self.scale_size(4)
        self.reference_text = scrolledtext.ScrolledText(main_frame, height=reference_height, width=input_width)
        self.reference_text.grid(row=7, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, self.scale_size(10)))
        
        # 右クリックメニューを設定
        self.setup_text_context_menu(self.reference_text)
        
        # ボタンフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=self.scale_size(10))
        
        # 変換ボタン
        self.convert_btn = ttk.Button(button_frame, text="変換", command=self.convert_text)
        self.convert_btn.pack(side=tk.LEFT, padx=(0, self.scale_size(10)))
        
        # コピーボタン
        self.copy_btn = ttk.Button(button_frame, text="コピー", command=self.copy_output)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, self.scale_size(10)))
        
        # クリアボタン
        self.clear_btn = ttk.Button(button_frame, text="クリア", command=self.clear_text)
        self.clear_btn.pack(side=tk.LEFT)
        
        # 出力ボックス
        ttk.Label(main_frame, text="出力ボックス:").grid(row=9, column=0, sticky=tk.W, pady=(self.scale_size(10), self.scale_size(5)))
        output_height = self.scale_size(6)
        self.output_text = scrolledtext.ScrolledText(main_frame, height=output_height, width=input_width)
        self.output_text.grid(row=10, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, self.scale_size(10)))
        
        # 右クリックメニューを設定
        self.setup_text_context_menu(self.output_text)
        
        # ステータスバー
        self.status_var = tk.StringVar()
        self.status_var.set("準備完了")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.grid(row=11, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(self.scale_size(10), 0))
        
        # 行の重みを設定（テキストボックスが拡張できるように）
        main_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(7, weight=1)
        main_frame.rowconfigure(10, weight=1)
        
    def on_input_key_release(self, event):
        """入力ボックスのキーリリースイベント処理"""
        # 変換中は何もしない
        if self.convert_btn['state'] == 'disabled':
            return
        
        # 特定のキー（Ctrl、Alt、Shift等）は無視
        if event.keysym in ['Control_L', 'Control_R', 'Alt_L', 'Alt_R', 
                           'Shift_L', 'Shift_R', 'Menu', 'Super_L', 'Super_R']:
            return
            
        # 入力テキスト全体を取得
        current_text = self.input_text.get(1.0, tk.END)
        
        # 末尾の連続改行をチェック
        if current_text.endswith('\n\n\n'):
            # 3連続改行を検出したら変換を開始
            # まず末尾の改行を削除
            clean_text = current_text.rstrip('\n')
            self.input_text.delete(1.0, tk.END)
            self.input_text.insert(1.0, clean_text)
            
            # 変換を開始
            self.convert_text()
    
    def setup_text_context_menu(self, text_widget):
        """テキストウィジェットに右クリックメニューを設定"""
        context_menu = tk.Menu(text_widget, tearoff=0)
        
        # メニュー項目を追加
        context_menu.add_command(label="切り取り", command=lambda: self.text_cut(text_widget))
        context_menu.add_command(label="コピー", command=lambda: self.text_copy(text_widget))
        context_menu.add_command(label="貼り付け", command=lambda: self.text_paste(text_widget))
        context_menu.add_separator()
        context_menu.add_command(label="全て選択", command=lambda: self.text_select_all(text_widget))
        
        # 右クリックイベントをバインド
        def show_context_menu(event):
            try:
                context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                context_menu.grab_release()
        
        text_widget.bind("<Button-3>", show_context_menu)  # 右クリック
        
    def text_cut(self, text_widget):
        """テキストを切り取り"""
        try:
            text_widget.event_generate("<<Cut>>")
        except tk.TclError:
            pass
    
    def text_copy(self, text_widget):
        """テキストをコピー"""
        try:
            text_widget.event_generate("<<Copy>>")
        except tk.TclError:
            pass
    
    def text_paste(self, text_widget):
        """テキストを貼り付け"""
        try:
            text_widget.event_generate("<<Paste>>")
        except tk.TclError:
            pass
    
    def text_select_all(self, text_widget):
        """全てのテキストを選択"""
        try:
            text_widget.tag_add(tk.SEL, "1.0", tk.END)
            text_widget.mark_set(tk.INSERT, "1.0")
            text_widget.see(tk.INSERT)
        except tk.TclError:
            pass
        
    def update_reference_files(self):
        """参考用ファイルリストを更新"""
        try:
            if not os.path.exists(self.reference_folder):
                os.makedirs(self.reference_folder)
            
            # テキストファイルを検索
            pattern = os.path.join(self.reference_folder, "*.txt")
            files = glob.glob(pattern)
            
            # ファイル名のみを取得
            self.reference_files = [os.path.basename(f) for f in files]
            
            # コンボボックスを更新
            self.reference_selector['values'] = [""] + self.reference_files
            
            self.status_var.set(f"参考用ファイル {len(self.reference_files)} 件を読み込みました")
            
        except Exception as e:
            messagebox.showerror("エラー", f"参考用ファイルの読み込みに失敗しました: {str(e)}")
            
    def on_reference_selected(self, event=None):
        """参考用ファイルが選択されたときの処理"""
        selected_file = self.reference_selector.get()
        if selected_file:
            try:
                file_path = os.path.join(self.reference_folder, selected_file)
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                self.reference_text.delete(1.0, tk.END)
                self.reference_text.insert(1.0, content)
                
                self.settings["selected_reference_file"] = selected_file
                
            except Exception as e:
                messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {str(e)}")
        else:
            self.reference_text.delete(1.0, tk.END)
            self.settings["selected_reference_file"] = ""
            
    def convert_text(self):
        """テキストの変換処理"""
        input_text = self.input_text.get(1.0, tk.END).strip()
        if not input_text:
            # ダイアログではなくステータス表示と警告音のみ
            self.status_var.set("警告: 入力テキストが空です")
            try:
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            except Exception:
                pass
            return
            
        # 設定を保存
        self.save_settings()
        
        # 参考用ファイルリストを更新
        self.update_reference_files()
        
        # 変換ボタンを無効化
        self.convert_btn.config(state='disabled')
        self.status_var.set("変換中...")
        
        # 別スレッドで変換処理を実行
        thread = threading.Thread(target=self._convert_text_async, args=(input_text,))
        thread.daemon = True
        thread.start()
        
    def _convert_text_async(self, input_text: str):
        """非同期でテキスト変換を実行"""
        try:
            # OpenRouter APIを呼び出し
            corrected_text = self.call_openrouter_api(input_text)
            
            # UIを更新（メインスレッドで実行）
            self.root.after(0, self._update_output, corrected_text)
            
        except Exception as e:
            error_msg = f"変換に失敗しました: {str(e)}"
            self.root.after(0, self._show_error, error_msg)
            
    def _update_output(self, corrected_text: str):
        """出力テキストを更新"""
        self.output_text.delete(1.0, tk.END)
        self.output_text.insert(1.0, corrected_text)
        
        # クリップボードにコピー
        try:
            pyperclip.copy(corrected_text)
        except Exception as e:
            print(f"クリップボードへのコピーに失敗: {e}")
            
        # 完了音を再生
        try:
            winsound.MessageBeep(winsound.MB_OK)
        except Exception as e:
            print(f"音声再生に失敗: {e}")
            
        # ボタンを有効化
        self.convert_btn.config(state='normal')
        self.status_var.set("変換完了")
        
    def _show_error(self, error_msg: str):
        """エラーメッセージを表示"""
        messagebox.showerror("エラー", error_msg)
        self.convert_btn.config(state='normal')
        self.status_var.set("エラーが発生しました")
        
    def call_openrouter_api(self, input_text: str) -> str:
        """OpenRouter APIを呼び出して文章を修正"""
        api_key = os.getenv('OPENROUTER_API_KEY')
        if not api_key:
            raise ValueError("OPENROUTER_API_KEY環境変数が設定されていません")
            
        # プロンプトデータを作成
        conversion_policy = self.policy_text.get(1.0, tk.END).strip()
        reference_text = self.reference_text.get(1.0, tk.END).strip()
        
        prompt_data = {
            "input_text": input_text,
            "conversion_policy": conversion_policy,
            "reference_text": reference_text
        }
        
        # デバッグ用：prompt_dataをコンソールに出力
        print("=== デバッグ：prompt_data ===")
        print(json.dumps(prompt_data, ensure_ascii=False, indent=2))
        print("===========================")
        
        # システムプロンプト
        system_prompt = """あなたは、音声入力されたテキストを修正・校正する専門家です。
あなたのタスクは、与えられた入力テキストを、文法的かつ文脈的に自然で正しい日本語の文章に変換することです。
入力されている情報には誤認識や余計な記号などが含まれる可能性があります。

以下のテンプレートに従って、入力を解釈し、出力を生成してください。

-----

### **入力テンプレート**

```json
{
  "input_text": "ここに音声入力された文字列が入ります。",
  "conversion_policy": "ここには変換の方針が入ります。",
  "reference_text": "ここに文のスタイルの参考となる文章が入ります。この内容を直接的に出力に反映してはいけません。"
}
```

-----

### **実行指示**

1.  **テキストの解析**: まず、`input_text` を読み込み、誤字、脱字、文法的な誤り、不自然な言い回しを特定します。
2.  **方針の適用**:
      * `conversion_policy` に指示がある場合は、その方針に厳密に従ってテキストを修正します。（例：「地名は漢字にする」「プロダクト名は英字にする」など）
      * `conversion_policy` が空欄の場合は、文脈に沿った日本語として最も自然な文章になるように修正してください。
3.  **文体の参照**: `reference_text` が提供されている場合は、その文章のスタイル、トーン、語彙、句読点の使い方を参考にして、出力する文章の自然さを高めてください。
4.  **修正の実行**: 上記の解析、方針、参照に基づき、`input_text` の元の意図を絶対に損なわないように注意しながら、句読点、助詞、接続詞などを適切に補い、自然で流暢な文章を作成します。
5.  **出力の生成**: 修正が完了した文章を、以下の出力テンプレートの `corrected_text` の値として生成します。

-----

### **出力テンプレート**

```json
{
  "corrected_text": "ここに修正・校正が完了した文章を生成します。"
}
```

**【重要】**

  * 出力は、指定された**JSON形式の出力テンプレート**を厳守してください。
  * `corrected_text` の値以外に、いかなる説明、前置き、後書きも追加してはいけません。
  * 文体や文のトーンを厳守してください。特に参考用のテキストに含まれるトーンは重視してください。
"""
        
        # APIリクエストを作成
        headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "HTTP-Referer": "https://github.com/gpsnmeajp/VOICE_CORRECTOR",
            "X-Title": "VOICE_CORRECTOR"
        }
        
        data = {
            "model": "openai/gpt-5",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(prompt_data, ensure_ascii=False)}
            ],
            "temperature": 0.5
        }
        
        # APIを呼び出し
        response = requests.post(
            "https://openrouter.ai/api/v1/chat/completions",
            headers=headers,
            json=data,
            timeout=30
        )
        
        if response.status_code != 200:
            raise Exception(f"API呼び出しに失敗しました: {response.status_code}")
            
        try:
            result = response.json()
        except json.JSONDecodeError:
            raise Exception("APIからの応答が無効なJSON形式です")
        
        if 'choices' not in result or not result['choices']:
            raise Exception("APIからの応答にchoicesが含まれていません")
            
        content = result['choices'][0]['message']['content']
        
        # JSON応答をパース（マークダウンやコードブロック内のJSONも対応）
        corrected_text = self.extract_json_response(content)
        return corrected_text
    
    def extract_json_response(self, content: str) -> str:
        """様々な形式の応答からJSON部分を抽出してcorrected_textを取得"""
        import re
        
        # 1. 直接JSONとして解析を試行
        try:
            parsed_result = json.loads(content.strip())
            if isinstance(parsed_result, dict) and 'corrected_text' in parsed_result:
                return parsed_result['corrected_text']
        except json.JSONDecodeError:
            pass
        
        # 2. マークダウンコードブロック内のJSONを検索
        # ```json ... ``` 形式
        json_pattern = r'```(?:json)?\s*\n?(.*?)\n?```'
        matches = re.findall(json_pattern, content, re.DOTALL | re.IGNORECASE)
        
        for match in matches:
            try:
                parsed_result = json.loads(match.strip())
                if isinstance(parsed_result, dict) and 'corrected_text' in parsed_result:
                    return parsed_result['corrected_text']
            except json.JSONDecodeError:
                continue
        
        # 3. { } で囲まれたJSON部分を検索
        brace_pattern = r'\{[^{}]*"corrected_text"[^{}]*\}'
        matches = re.findall(brace_pattern, content, re.DOTALL)
        
        for match in matches:
            try:
                parsed_result = json.loads(match)
                if isinstance(parsed_result, dict) and 'corrected_text' in parsed_result:
                    return parsed_result['corrected_text']
            except json.JSONDecodeError:
                continue
        
        # 4. より複雑なネストしたJSONパターンを検索
        nested_pattern = r'\{(?:[^{}]|{[^{}]*})*"corrected_text"(?:[^{}]|{[^{}]*})*\}'
        matches = re.findall(nested_pattern, content, re.DOTALL)
        
        for match in matches:
            try:
                parsed_result = json.loads(match)
                if isinstance(parsed_result, dict) and 'corrected_text' in parsed_result:
                    return parsed_result['corrected_text']
            except json.JSONDecodeError:
                continue
        
        # 5. "corrected_text": "..." の値を直接抽出
        text_pattern = r'"corrected_text"\s*:\s*"([^"]*(?:\\.[^"]*)*)"'
        match = re.search(text_pattern, content)
        if match:
            # エスケープ文字を処理
            text = match.group(1)
            text = text.replace('\\"', '"').replace('\\n', '\n').replace('\\t', '\t').replace('\\\\', '\\')
            return text
        
        # 6. すべて失敗した場合、元のコンテンツを返す（フォールバック）
        # ただし、明らかにJSON形式でない場合は説明として扱う
        stripped_content = content.strip()
        if stripped_content.startswith('{') or stripped_content.startswith('```'):
            raise Exception(f"JSON形式の応答を解析できませんでした: {content[:200]}...")
        else:
            # プレーンテキストとして返す
            return stripped_content
            
    def copy_output(self):
        """出力テキストをクリップボードにコピー"""
        output = self.output_text.get(1.0, tk.END).strip()
        if output:
            try:
                pyperclip.copy(output)
                self.status_var.set("クリップボードにコピーしました")
            except Exception as e:
                messagebox.showerror("エラー", f"クリップボードへのコピーに失敗しました: {str(e)}")
        else:
            # ダイアログではなくステータス表示と警告音のみ
            self.status_var.set("警告: 出力テキストが空です")
            try:
                winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
            except Exception:
                pass
            
    def clear_text(self):
        """入力と出力をクリア"""
        self.input_text.delete(1.0, tk.END)
        self.output_text.delete(1.0, tk.END)
        self.status_var.set("クリアしました")
        
    def save_settings(self):
        """設定をJSONファイルに保存"""
        try:
            self.settings["conversion_policy"] = self.policy_text.get(1.0, tk.END).strip()
            self.settings["reference_text"] = self.reference_text.get(1.0, tk.END).strip()
            self.settings["selected_reference_file"] = self.reference_selector.get()
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, ensure_ascii=False, indent=2)
                
        except Exception as e:
            print(f"設定の保存に失敗: {e}")
            
    def load_settings(self):
        """設定をJSONファイルから読み込み"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.settings = json.load(f)
                
                # GUIに設定を反映
                self.policy_text.delete(1.0, tk.END)
                self.policy_text.insert(1.0, self.settings.get("conversion_policy", ""))
                
                self.reference_text.delete(1.0, tk.END)
                self.reference_text.insert(1.0, self.settings.get("reference_text", ""))
                
                # 参考ファイルが指定されている場合、それを選択
                selected_file = self.settings.get("selected_reference_file", "")
                if selected_file:
                    # 参考ファイルリストが更新された後に設定する必要があるため、
                    # update_reference_files()呼び出し後に設定
                    self.root.after(100, lambda: self._set_reference_file(selected_file))
                    
        except Exception as e:
            print(f"設定の読み込みに失敗: {e}")
    
    def _set_reference_file(self, selected_file: str):
        """参考ファイルを設定する補助メソッド"""
        if selected_file in self.reference_files:
            self.reference_selector.set(selected_file)
            self.on_reference_selected()  # ファイル内容も読み込む
            
    def run(self):
        """アプリケーションを実行"""
        # 終了時に設定を保存
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.root.mainloop()
        
    def on_closing(self):
        """アプリケーション終了時の処理"""
        self.save_settings()
        self.root.destroy()


def main():
    """メイン関数"""
    app = VoiceCorrector()
    app.run()


if __name__ == "__main__":
    main()