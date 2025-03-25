import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pyperclip
from datetime import datetime
import json
import re
from functools import partial

class FileManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ファイル管理ユーティリティ")
        self.root.geometry("800x600")
        
        # 変数の初期化
        self.sequence_number = tk.StringVar(value="1")
        self.sequence_digits = tk.StringVar(value="4")  # 連番の桁数設定用
        self.auto_increment = tk.BooleanVar(value=False)  # 連番自動増加用
        self.custom_text = tk.StringVar()
        self.date_format = tk.StringVar(value="%Y%m%d")
        self.filename_pattern = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.selected_file_path = tk.StringVar()
        self.batch_mode = tk.BooleanVar(value=False)  # 一括処理モード用
        self.selected_files = []  # 一括処理用のファイルリスト
        
        # メインノートブックとタブの作成
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # タブの作成
        self.create_filename_tab()
        self.create_destination_tab()
        self.create_operations_tab()
        
        # ステータスバー
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_var.set("準備完了")
        
        # ショートカットキーの設定
        self.setup_shortcuts()
        
        # 保存されたテンプレートがあれば読み込む
        self.load_settings()

    def setup_shortcuts(self):
        # コピー: Ctrl+C
        self.root.bind("<Control-c>", lambda event: self.copy_to_clipboard())
        
        # 名前変更: Ctrl+R
        self.root.bind("<Control-r>", lambda event: self.rename_file())
        
        # ファイル移動: Ctrl+M
        self.root.bind("<Control-m>", lambda event: self.move_file())
        
        # 名前変更と移動: Ctrl+Shift+M - 複数の表記方法で対応
        self.root.bind("<Control-Shift-M>", lambda event: self.rename_and_move_file())
        self.root.bind("<Control-Shift-m>", lambda event: self.rename_and_move_file())
        self.root.bind("<Control-M>", lambda event: self.rename_and_move_file())
        
        # プレビュー更新: F5
        self.root.bind("<F5>", lambda event: self.update_filename_preview())
        
        # ファイル選択: Ctrl+O
        self.root.bind("<Control-o>", lambda event: self.browse_file())
        
        # 保存先選択: Ctrl+D
        self.root.bind("<Control-d>", lambda event: self.browse_destination())
        
        # テンプレート保存: Ctrl+S
        self.root.bind("<Control-s>", lambda event: self.save_filename_template())
        
        # テンプレート読込: Ctrl+L
        self.root.bind("<Control-l>", lambda event: self.load_filename_template())
        
        # キーボードイベントデバッグ用（開発時のみ使用）
        # self.root.bind_all("<Key>", self.debug_key_event)

    def create_filename_tab(self):
        filename_frame = ttk.Frame(self.notebook)
        self.notebook.add(filename_frame, text="ファイル名設定")
        
        # パターンプレビュー
        preview_frame = ttk.LabelFrame(filename_frame, text="ファイル名プレビュー")
        preview_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.preview_label = ttk.Label(preview_frame, textvariable=self.filename_pattern, font=("Courier", 12))
        self.preview_label.pack(pady=10)
        
        # パターンコンポーネント
        components_frame = ttk.LabelFrame(filename_frame, text="ファイル名コンポーネント")
        components_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 日付フォーマット
        date_frame = ttk.Frame(components_frame)
        date_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(date_frame, text="日付フォーマット:").pack(side=tk.LEFT, padx=5)
        date_formats = ["%Y%m%d", "%Y-%m-%d", "%d%m%Y", "%d-%m-%Y", "%m%d%Y", "%m-%d-%Y"]
        self.date_combo = ttk.Combobox(date_frame, textvariable=self.date_format, values=date_formats)
        self.date_combo.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Button(date_frame, text="日付を挿入", command=lambda: self.insert_placeholder("{date}")).pack(side=tk.LEFT, padx=5)
        
        # 連番
        seq_frame = ttk.Frame(components_frame)
        seq_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(seq_frame, text="連番:").pack(side=tk.LEFT, padx=5)
        ttk.Spinbox(seq_frame, from_=1, to=9999, textvariable=self.sequence_number, width=5).pack(side=tk.LEFT, padx=5)
        
        # 連番の桁数設定
        ttk.Label(seq_frame, text="桁数:").pack(side=tk.LEFT, padx=5)
        ttk.Spinbox(seq_frame, from_=1, to=10, textvariable=self.sequence_digits, width=3).pack(side=tk.LEFT, padx=5)
        
        # 連番自動増加チェックボックス
        ttk.Checkbutton(seq_frame, text="コピー時に自動増加", variable=self.auto_increment).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(seq_frame, text="連番を挿入", command=lambda: self.insert_placeholder("{seq}")).pack(side=tk.LEFT, padx=5)
        
        # カスタムテキスト
        text_frame = ttk.Frame(components_frame)
        text_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(text_frame, text="カスタムテキスト:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(text_frame, textvariable=self.custom_text).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        ttk.Button(text_frame, text="テキストを挿入", command=lambda: self.insert_placeholder("{text}")).pack(side=tk.LEFT, padx=5)
        
        # 区切り文字
        separator_frame = ttk.Frame(components_frame)
        separator_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(separator_frame, text="区切り文字:").pack(side=tk.LEFT, padx=5)
        separators = ["_", "-", ".", " "]
        for sep in separators:
            ttk.Button(separator_frame, text=f'"{sep}"', width=3, 
                      command=partial(self.insert_text, sep)).pack(side=tk.LEFT, padx=2)
        
        # カスタム区切り文字用のエントリー
        self.custom_separator = ttk.Entry(separator_frame, width=5)
        self.custom_separator.pack(side=tk.LEFT, padx=5)
        ttk.Button(separator_frame, text="挿入", 
                  command=lambda: self.insert_text(self.custom_separator.get())).pack(side=tk.LEFT, padx=2)
        
        # パターン編集
        pattern_frame = ttk.LabelFrame(filename_frame, text="ファイル名パターン")
        pattern_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.pattern_entry = ttk.Entry(pattern_frame, font=("Courier", 12))
        self.pattern_entry.pack(fill=tk.X, padx=5, pady=5, expand=True)
        
        button_frame = ttk.Frame(pattern_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(button_frame, text="プレビュー更新 (F5)", command=self.update_filename_preview).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="クリップボードにコピー (Ctrl+C)", command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        
        # テンプレート管理
        template_frame = ttk.LabelFrame(filename_frame, text="テンプレート管理")
        template_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.filename_templates = {}
        self.template_name_var = tk.StringVar()
        
        template_entry_frame = ttk.Frame(template_frame)
        template_entry_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(template_entry_frame, text="テンプレート名:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(template_entry_frame, textvariable=self.template_name_var).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        template_button_frame = ttk.Frame(template_frame)
        template_button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(template_button_frame, text="テンプレート保存 (Ctrl+S)", command=self.save_filename_template).pack(side=tk.LEFT, padx=5)
        
        self.template_combo = ttk.Combobox(template_button_frame, width=30)
        self.template_combo.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Button(template_button_frame, text="テンプレート読込 (Ctrl+L)", command=self.load_filename_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="テンプレート削除", command=self.delete_filename_template).pack(side=tk.LEFT, padx=5)

    def create_destination_tab(self):
        destination_frame = ttk.Frame(self.notebook)
        self.notebook.add(destination_frame, text="保存先設定")
        
        # 現在の保存先
        current_dest_frame = ttk.LabelFrame(destination_frame, text="現在の保存先")
        current_dest_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.dest_label = ttk.Label(current_dest_frame, textvariable=self.destination_path, font=("Courier", 10))
        self.dest_label.pack(pady=10, fill=tk.X)
        
        # 保存先選択
        select_frame = ttk.Frame(current_dest_frame)
        select_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(select_frame, text="参照... (Ctrl+D)", command=self.browse_destination).pack(side=tk.LEFT, padx=5)
        
        # 保存先テンプレート
        template_frame = ttk.LabelFrame(destination_frame, text="保存先テンプレート")
        template_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.dest_templates = {}
        self.dest_template_name_var = tk.StringVar()
        
        template_entry_frame = ttk.Frame(template_frame)
        template_entry_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(template_entry_frame, text="テンプレート名:").pack(side=tk.LEFT, padx=5)
        ttk.Entry(template_entry_frame, textvariable=self.dest_template_name_var).pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        template_button_frame = ttk.Frame(template_frame)
        template_button_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(template_button_frame, text="テンプレート保存", command=self.save_destination_template).pack(side=tk.LEFT, padx=5)
        
        self.dest_template_combo = ttk.Combobox(template_button_frame, width=30)
        self.dest_template_combo.pack(side=tk.LEFT, padx=5, expand=True, fill=tk.X)
        
        ttk.Button(template_button_frame, text="テンプレート読込", command=self.load_destination_template).pack(side=tk.LEFT, padx=5)
        ttk.Button(template_button_frame, text="テンプレート削除", command=self.delete_destination_template).pack(side=tk.LEFT, padx=5)

    def create_operations_tab(self):
        operations_frame = ttk.Frame(self.notebook)
        self.notebook.add(operations_frame, text="ファイル操作")
        
        # ファイル選択
        file_frame = ttk.LabelFrame(operations_frame, text="ファイル選択")
        file_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # 一括処理モード用チェックボックス
        batch_frame = ttk.Frame(file_frame)
        batch_frame.pack(fill=tk.X, pady=5)
        ttk.Checkbutton(batch_frame, text="一括処理モード", variable=self.batch_mode, 
                       command=self.toggle_batch_mode).pack(side=tk.LEFT, padx=5)
        
        # 単一ファイル用表示
        self.single_file_frame = ttk.Frame(file_frame)
        self.single_file_frame.pack(fill=tk.X, pady=5)
        
        self.file_label = ttk.Label(self.single_file_frame, textvariable=self.selected_file_path, font=("Courier", 10))
        self.file_label.pack(pady=10, fill=tk.X)
        
        ttk.Button(self.single_file_frame, text="ファイル参照... (Ctrl+O)", command=self.browse_file).pack(pady=5)
        
        # 複数ファイル用リストボックス (初期状態では非表示)
        self.batch_file_frame = ttk.Frame(file_frame)
        
        self.files_listbox = tk.Listbox(self.batch_file_frame, height=6)
        scrollbar = ttk.Scrollbar(self.batch_file_frame, orient="vertical", command=self.files_listbox.yview)
        self.files_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=2, pady=5)
        
        batch_button_frame = ttk.Frame(self.batch_file_frame)
        batch_button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(batch_button_frame, text="ファイル追加", command=self.add_files).pack(side=tk.LEFT, padx=5)
        ttk.Button(batch_button_frame, text="選択ファイル削除", command=self.remove_selected_file).pack(side=tk.LEFT, padx=5)
        ttk.Button(batch_button_frame, text="すべて削除", command=self.clear_files).pack(side=tk.LEFT, padx=5)
        
        # 操作
        op_frame = ttk.LabelFrame(operations_frame, text="操作")
        op_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(op_frame, text="名前をクリップボードにコピー (Ctrl+C)", command=self.copy_to_clipboard).pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(op_frame, text="ファイル名変更 (Ctrl+R)", command=self.rename_file).pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(op_frame, text="ファイル移動 (Ctrl+M)", command=self.move_file).pack(fill=tk.X, padx=5, pady=5)
        ttk.Button(op_frame, text="ファイル名変更と移動 (Ctrl+Shift+M)", command=self.rename_and_move_file).pack(fill=tk.X, padx=5, pady=5)
        
        # ショートカットキーリスト
        shortcut_frame = ttk.LabelFrame(operations_frame, text="ショートカットキー")
        shortcut_frame.pack(fill=tk.X, padx=10, pady=10)
        
        shortcuts = [
            "F5: プレビュー更新",
            "Ctrl+C: クリップボードにコピー",
            "Ctrl+R: ファイル名変更",
            "Ctrl+M: ファイル移動",
            "Ctrl+Shift+M: ファイル名変更と移動",
            "Ctrl+O: ファイル選択",
            "Ctrl+D: 保存先選択",
            "Ctrl+S: テンプレート保存",
            "Ctrl+L: テンプレート読込"
        ]
        
        for i, shortcut in enumerate(shortcuts):
            if i % 3 == 0:
                row_frame = ttk.Frame(shortcut_frame)
                row_frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(row_frame, text=shortcut, width=25).pack(side=tk.LEFT, padx=5)

    def toggle_batch_mode(self):
        if self.batch_mode.get():
            self.single_file_frame.pack_forget()
            self.batch_file_frame.pack(fill=tk.X, pady=5)
        else:
            self.batch_file_frame.pack_forget()
            self.single_file_frame.pack(fill=tk.X, pady=5)

    def add_files(self):
        filenames = filedialog.askopenfilenames()
        if filenames:
            for file in filenames:
                if file not in self.selected_files:
                    self.selected_files.append(file)
                    self.files_listbox.insert(tk.END, os.path.basename(file))
            self.status_var.set(f"{len(filenames)}個のファイルを追加しました")

    def remove_selected_file(self):
        selected_indices = self.files_listbox.curselection()
        if selected_indices:
            # 選択項目を逆順で削除（インデックスが変わるのを防ぐため）
            for index in sorted(selected_indices, reverse=True):
                del self.selected_files[index]
                self.files_listbox.delete(index)
            self.status_var.set("選択したファイルを削除しました")

    def clear_files(self):
        self.selected_files.clear()
        self.files_listbox.delete(0, tk.END)
        self.status_var.set("すべてのファイルをリストから削除しました")

    def insert_text(self, text):
        current_pos = self.pattern_entry.index(tk.INSERT)
        current_text = self.pattern_entry.get()
        
        new_text = current_text[:current_pos] + text + current_text[current_pos:]
        self.pattern_entry.delete(0, tk.END)
        self.pattern_entry.insert(0, new_text)
        self.update_filename_preview()

    def update_filename_preview(self):
        pattern = self.pattern_entry.get()
        if not pattern:
            self.filename_pattern.set("")
            return
        
        # プレースホルダーを実際の値に置き換え
        filename = pattern
        if "{date}" in pattern:
            current_date = datetime.now().strftime(self.date_format.get())
            filename = filename.replace("{date}", current_date)
        
        if "{seq}" in pattern:
            # 連番の桁数に合わせてゼロ埋め
            digits = int(self.sequence_digits.get())
            if digits <= 1:
                # 桁数が1以下の場合はゼロ埋めしない
                seq_num = str(int(self.sequence_number.get()))
            else:
                seq_num = self.sequence_number.get().zfill(digits)
            filename = filename.replace("{seq}", seq_num)
        
        if "{text}" in pattern:
            custom = self.custom_text.get()
            filename = filename.replace("{text}", custom)
        
        self.filename_pattern.set(filename)

    def insert_placeholder(self, placeholder):
        current_pos = self.pattern_entry.index(tk.INSERT)
        current_text = self.pattern_entry.get()
        
        new_text = current_text[:current_pos] + placeholder + current_text[current_pos:]
        self.pattern_entry.delete(0, tk.END)
        self.pattern_entry.insert(0, new_text)
        self.update_filename_preview()

    def copy_to_clipboard(self):
        current_name = self.filename_pattern.get()
        if current_name:
            pyperclip.copy(current_name)
            self.status_var.set(f"クリップボードにコピーしました: {current_name}")
            
            # 自動増加が有効の場合、連番を増加
            if self.auto_increment.get():
                try:
                    current_seq = int(self.sequence_number.get())
                    self.sequence_number.set(str(current_seq + 1))
                    self.update_filename_preview()
                except ValueError:
                    pass
        else:
            messagebox.showwarning("警告", "ファイル名パターンが指定されていません")

    def save_filename_template(self):
        template_name = self.template_name_var.get()
        if not template_name:
            messagebox.showwarning("警告", "テンプレート名を入力してください")
            return
        
        pattern = self.pattern_entry.get()
        if not pattern:
            messagebox.showwarning("警告", "保存するパターンがありません")
            return
        
        self.filename_templates[template_name] = {
            "pattern": pattern,
            "date_format": self.date_format.get(),
            "custom_text": self.custom_text.get(),
            "sequence_digits": self.sequence_digits.get()  # 連番桁数も保存
        }
        
        self.update_filename_templates_combo()
        self.save_settings()
        self.status_var.set(f"ファイル名テンプレートを保存しました: {template_name}")

    def update_filename_templates_combo(self):
        template_names = list(self.filename_templates.keys())
        self.template_combo['values'] = template_names
        if template_names:
            self.template_combo.current(0)

    def load_filename_template(self):
        template_name = self.template_combo.get()
        if not template_name or template_name not in self.filename_templates:
            messagebox.showwarning("警告", "テンプレートが選択されていないか存在しません")
            return
        
        template = self.filename_templates[template_name]
        self.pattern_entry.delete(0, tk.END)
        self.pattern_entry.insert(0, template["pattern"])
        self.date_format.set(template["date_format"])
        self.custom_text.set(template["custom_text"])
        
        # 連番桁数があれば設定
        if "sequence_digits" in template:
            self.sequence_digits.set(template["sequence_digits"])
        
        self.update_filename_preview()
        self.status_var.set(f"ファイル名テンプレートを読み込みました: {template_name}")

    def delete_filename_template(self):
        template_name = self.template_combo.get()
        if not template_name or template_name not in self.filename_templates:
            messagebox.showwarning("警告", "テンプレートが選択されていないか存在しません")
            return
        
        del self.filename_templates[template_name]
        self.update_filename_templates_combo()
        self.save_settings()
        self.status_var.set(f"ファイル名テンプレートを削除しました: {template_name}")

    def browse_destination(self):
        directory = filedialog.askdirectory()
        if directory:
            self.destination_path.set(directory)
            self.status_var.set(f"保存先を選択しました: {directory}")

    def save_destination_template(self):
        template_name = self.dest_template_name_var.get()
        if not template_name:
            messagebox.showwarning("警告", "テンプレート名を入力してください")
            return
        
        destination = self.destination_path.get()
        if not destination:
            messagebox.showwarning("警告", "保存する保存先がありません")
            return
        
        self.dest_templates[template_name] = destination
        
        self.update_destination_templates_combo()
        self.save_settings()
        self.status_var.set(f"保存先テンプレートを保存しました: {template_name}")

    def update_destination_templates_combo(self):
        template_names = list(self.dest_templates.keys())
        self.dest_template_combo['values'] = template_names
        if template_names:
            self.dest_template_combo.current(0)

    def load_destination_template(self):
        template_name = self.dest_template_combo.get()
        if not template_name or template_name not in self.dest_templates:
            messagebox.showwarning("警告", "テンプレートが選択されていないか存在しません")
            return
        
        destination = self.dest_templates[template_name]
        self.destination_path.set(destination)
        self.status_var.set(f"保存先テンプレートを読み込みました: {template_name}")

    def delete_destination_template(self):
        template_name = self.dest_template_combo.get()
        if not template_name or template_name not in self.dest_templates:
            messagebox.showwarning("警告", "テンプレートが選択されていないか存在しません")
            return
        
        del self.dest_templates[template_name]
        self.update_destination_templates_combo()
        self.save_settings()
        self.status_var.set(f"保存先テンプレートを削除しました: {template_name}")

    def browse_file(self):
        filename = filedialog.askopenfilename()
        if filename:
            self.selected_file_path.set(filename)
            self.status_var.set(f"ファイルを選択しました: {filename}")

    def get_formatted_filename(self, with_extension=True):
        if not self.filename_pattern.get():
            return None
        
        # 現在のファイル名パターンを取得
        new_filename = self.filename_pattern.get()
        
        if with_extension:
            # ファイル拡張子を追加
            if self.batch_mode.get():
                # 一括処理モードではダミー拡張子
                return new_filename + ".ext"
            else:
                source_path = self.selected_file_path.get()
                _, file_extension = os.path.splitext(source_path)
                return new_filename + file_extension
        
        return new_filename

    def rename_file(self):
        if self.batch_mode.get():
            self.batch_rename_files()
        else:
            self.single_rename_file()

    def single_rename_file(self):
        if not self.selected_file_path.get():
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not self.filename_pattern.get():
            messagebox.showwarning("警告", "ファイル名パターンが指定されていません")
            return
        
        source_path = self.selected_file_path.get()
        
        # ファイルの存在確認
        if not os.path.exists(source_path):
            messagebox.showerror("エラー", f"ファイルが見つかりません: {source_path}\n\nファイルが移動または削除された可能性があります。")
            return
            
        source_dir = os.path.dirname(source_path)
        _, file_extension = os.path.splitext(source_path)
        
        new_filename = self.filename_pattern.get() + file_extension
        destination_path = os.path.join(source_dir, new_filename)
        
        # 絶対パスに変換して正規化
        source_path = os.path.abspath(os.path.normpath(source_path))
        destination_path = os.path.abspath(os.path.normpath(destination_path))
        
        # 既に同名のファイルが存在するかチェック
        if os.path.exists(destination_path) and source_path != destination_path:
            result = messagebox.askyesno("確認", f"既に同じ名前のファイルが存在します: {new_filename}\n上書きしますか？")
            if not result:
                return
        
        try:
            shutil.move(source_path, destination_path)
            self.selected_file_path.set(destination_path)
            self.status_var.set(f"ファイル名を変更しました: {new_filename}")
            
            # 自動増加が有効の場合、連番を増加
            if self.auto_increment.get():
                try:
                    current_seq = int(self.sequence_number.get())
                    self.sequence_number.set(str(current_seq + 1))
                    self.update_filename_preview()
                except ValueError:
                    pass
        except PermissionError:
            messagebox.showerror("エラー", f"ファイル '{new_filename}' へのアクセス権限がありません。\nファイルが他のプログラムで使用中でないか確認してください。")
        except FileNotFoundError:
            messagebox.showerror("エラー", f"ファイルが見つかりません: {source_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル名変更中にエラーが発生しました: {str(e)}")

    def batch_rename_files(self):
        if not self.selected_files:
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not self.filename_pattern.get():
            messagebox.showwarning("警告", "ファイル名パターンが指定されていません")
            return
        
        # 現在の連番を保存
        start_seq = int(self.sequence_number.get())
        success_count = 0
        
        # 処理前にファイルの存在を確認
        non_existent_files = []
        for source_path in self.selected_files:
            if not os.path.exists(source_path):
                non_existent_files.append(source_path)
        
        if non_existent_files:
            missing_files = "\n".join([os.path.basename(f) for f in non_existent_files])
            result = messagebox.askyesno("警告", f"以下のファイルが見つかりません。これらをスキップして続行しますか？\n\n{missing_files}")
            if not result:
                return
        
        # 処理を開始
        for i, source_path in enumerate(self.selected_files):
            # ファイルが存在しない場合はスキップ
            if not os.path.exists(source_path):
                continue
                
            # 連番を設定
            current_seq = start_seq + i
            self.sequence_number.set(str(current_seq))
            self.update_filename_preview()
            
            # 名前変更処理
            source_dir = os.path.dirname(source_path)
            _, file_extension = os.path.splitext(source_path)
            
            new_filename = self.filename_pattern.get() + file_extension
            destination_path = os.path.join(source_dir, new_filename)
            
            # 絶対パスに変換して正規化
            source_path = os.path.abspath(os.path.normpath(source_path))
            destination_path = os.path.abspath(os.path.normpath(destination_path))
            
            # 既に同名のファイルが存在するかチェック（自分自身以外）
            if os.path.exists(destination_path) and source_path != destination_path:
                result = messagebox.askyesno("確認", f"既に同じ名前のファイルが存在します: {new_filename}\n上書きしますか？")
                if not result:
                    continue
            
            try:
                shutil.move(source_path, destination_path)
                # リストとリストボックスの更新
                self.selected_files[i] = destination_path
                self.files_listbox.delete(i)
                self.files_listbox.insert(i, os.path.basename(destination_path))
                success_count += 1
            except PermissionError:
                messagebox.showerror("エラー", f"ファイル '{os.path.basename(source_path)}' へのアクセス権限がありません。")
            except FileNotFoundError:
                messagebox.showerror("エラー", f"ファイルが見つかりません: {os.path.basename(source_path)}")
            except Exception as e:
                messagebox.showerror("エラー", f"ファイル '{os.path.basename(source_path)}' の名前変更中にエラーが発生しました: {str(e)}")
        
        # 最終連番 + 1 を設定
        if self.auto_increment.get():
            self.sequence_number.set(str(start_seq + len(self.selected_files)))
            self.update_filename_preview()
        else:
            # 元の連番に戻す
            self.sequence_number.set(str(start_seq))
            self.update_filename_preview()
        
        self.status_var.set(f"{len(self.selected_files)}個中{success_count}個のファイル名を変更しました")

    def move_file(self):
        if self.batch_mode.get():
            self.batch_move_files()
        else:
            self.single_move_file()

    def single_move_file(self):
        if not self.selected_file_path.get():
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not self.destination_path.get():
            messagebox.showwarning("警告", "保存先が指定されていません")
            return
        
        source_path = self.selected_file_path.get()
        
        # ファイルの存在確認
        if not os.path.exists(source_path):
            messagebox.showerror("エラー", f"ファイルが見つかりません: {source_path}\n\nファイルが移動または削除された可能性があります。")
            return
            
        # 保存先ディレクトリの存在確認
        dest_dir = self.destination_path.get()
        if not os.path.exists(dest_dir):
            try:
                os.makedirs(dest_dir)
            except Exception as e:
                messagebox.showerror("エラー", f"保存先ディレクトリを作成できません: {dest_dir}\n{str(e)}")
                return
                
        filename = os.path.basename(source_path)
        destination_path = os.path.join(dest_dir, filename)
        
        # 絶対パスに変換して正規化
        source_path = os.path.abspath(os.path.normpath(source_path))
        destination_path = os.path.abspath(os.path.normpath(destination_path))
        
        # 既に同名のファイルが存在するかチェック
        if os.path.exists(destination_path) and source_path != destination_path:
            result = messagebox.askyesno("確認", f"保存先に同じ名前のファイルが既に存在します: {filename}\n上書きしますか？")
            if not result:
                return
        
        try:
            shutil.move(source_path, destination_path)
            self.selected_file_path.set(destination_path)
            self.status_var.set(f"ファイルを移動しました: {destination_path}")
        except PermissionError:
            messagebox.showerror("エラー", f"ファイル '{filename}' へのアクセス権限がありません。\nファイルが他のプログラムで使用中でないか確認してください。")
        except FileNotFoundError:
            messagebox.showerror("エラー", f"ファイルが見つかりません: {source_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル移動中にエラーが発生しました: {str(e)}")

    def batch_move_files(self):
        if not self.selected_files:
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not self.destination_path.get():
            messagebox.showwarning("警告", "保存先が指定されていません")
            return
        
        # 保存先ディレクトリの存在確認
        dest_dir = self.destination_path.get()
        if not os.path.exists(dest_dir):
            try:
                os.makedirs(dest_dir)
            except Exception as e:
                messagebox.showerror("エラー", f"保存先ディレクトリを作成できません: {dest_dir}\n{str(e)}")
                return
                
        # 処理前にファイルの存在を確認
        non_existent_files = []
        for source_path in self.selected_files:
            if not os.path.exists(source_path):
                non_existent_files.append(source_path)
        
        if non_existent_files:
            missing_files = "\n".join([os.path.basename(f) for f in non_existent_files])
            result = messagebox.askyesno("警告", f"以下のファイルが見つかりません。これらをスキップして続行しますか？\n\n{missing_files}")
            if not result:
                return
                
        success_count = 0
        new_files = []
        
        for i, source_path in enumerate(self.selected_files):
            # ファイルが存在しない場合はスキップ
            if not os.path.exists(source_path):
                new_files.append(source_path)  # 元のパスを保持
                continue
                
            filename = os.path.basename(source_path)
            destination_path = os.path.join(dest_dir, filename)
            
            # 絶対パスに変換して正規化
            source_path = os.path.abspath(os.path.normpath(source_path))
            destination_path = os.path.abspath(os.path.normpath(destination_path))
            
            # 既に同名のファイルが存在するかチェック
            if os.path.exists(destination_path) and source_path != destination_path:
                result = messagebox.askyesno("確認", f"保存先に同じ名前のファイルが既に存在します: {filename}\n上書きしますか？")
                if not result:
                    new_files.append(source_path)  # 元のパスを保持
                    continue
            
            try:
                shutil.move(source_path, destination_path)
                # 成功したら新しいパスをリストに追加
                new_files.append(destination_path)
                success_count += 1
            except PermissionError:
                messagebox.showerror("エラー", f"ファイル '{filename}' へのアクセス権限がありません。")
                new_files.append(source_path)  # 元のパスを保持
            except FileNotFoundError:
                messagebox.showerror("エラー", f"ファイルが見つかりません: {filename}")
                new_files.append(source_path)  # 元のパスを保持
            except Exception as e:
                messagebox.showerror("エラー", f"ファイル '{filename}' の移動中にエラーが発生しました: {str(e)}")
                new_files.append(source_path)  # 元のパスを保持
        
        # 成功したファイルをリストに更新
        self.selected_files = new_files
        
        # リストボックスを更新
        self.files_listbox.delete(0, tk.END)
        for file in self.selected_files:
            self.files_listbox.insert(tk.END, os.path.basename(file))
        
        self.status_var.set(f"{len(self.selected_files)}個中{success_count}個のファイルを移動しました")

    def rename_and_move_file(self):
        if self.batch_mode.get():
            self.batch_rename_and_move_files()
        else:
            self.single_rename_and_move_file()

    def single_rename_and_move_file(self):
        if not self.selected_file_path.get():
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not self.filename_pattern.get():
            messagebox.showwarning("警告", "ファイル名パターンが指定されていません")
            return
        
        if not self.destination_path.get():
            messagebox.showwarning("警告", "保存先が指定されていません")
            return
        
        source_path = self.selected_file_path.get()
        
        # ファイルの存在確認
        if not os.path.exists(source_path):
            messagebox.showerror("エラー", f"ファイルが見つかりません: {source_path}\n\nファイルが移動または削除された可能性があります。")
            return
            
        # 保存先ディレクトリの存在確認
        dest_dir = self.destination_path.get()
        if not os.path.exists(dest_dir):
            try:
                os.makedirs(dest_dir)
            except Exception as e:
                messagebox.showerror("エラー", f"保存先ディレクトリを作成できません: {dest_dir}\n{str(e)}")
                return
                
        _, file_extension = os.path.splitext(source_path)
        
        new_filename = self.filename_pattern.get() + file_extension
        destination_path = os.path.join(dest_dir, new_filename)
        
        # 絶対パスに変換して正規化
        source_path = os.path.abspath(os.path.normpath(source_path))
        destination_path = os.path.abspath(os.path.normpath(destination_path))
        
        # 既に同名のファイルが存在するかチェック
        if os.path.exists(destination_path) and source_path != destination_path:
            result = messagebox.askyesno("確認", f"保存先に同じ名前のファイルが既に存在します: {new_filename}\n上書きしますか？")
            if not result:
                return
        
        try:
            shutil.move(source_path, destination_path)
            self.selected_file_path.set(destination_path)
            self.status_var.set(f"ファイル名を変更し移動しました: {destination_path}")
            
            # 自動増加が有効の場合、連番を増加
            if self.auto_increment.get():
                try:
                    current_seq = int(self.sequence_number.get())
                    self.sequence_number.set(str(current_seq + 1))
                    self.update_filename_preview()
                except ValueError:
                    pass
        except PermissionError:
            messagebox.showerror("エラー", f"ファイル '{new_filename}' へのアクセス権限がありません。\nファイルが他のプログラムで使用中でないか確認してください。")
        except FileNotFoundError:
            messagebox.showerror("エラー", f"ファイルが見つかりません: {source_path}")
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル名変更と移動中にエラーが発生しました: {str(e)}")

    def batch_rename_and_move_files(self):
        if not self.selected_files:
            messagebox.showwarning("警告", "ファイルが選択されていません")
            return
        
        if not self.filename_pattern.get():
            messagebox.showwarning("警告", "ファイル名パターンが指定されていません")
            return
        
        if not self.destination_path.get():
            messagebox.showwarning("警告", "保存先が指定されていません")
            return
            
        # 保存先ディレクトリの存在確認
        dest_dir = self.destination_path.get()
        if not os.path.exists(dest_dir):
            try:
                os.makedirs(dest_dir)
            except Exception as e:
                messagebox.showerror("エラー", f"保存先ディレクトリを作成できません: {dest_dir}\n{str(e)}")
                return
        
        # 処理前にファイルの存在を確認
        non_existent_files = []
        for source_path in self.selected_files:
            if not os.path.exists(source_path):
                non_existent_files.append(source_path)
        
        if non_existent_files:
            missing_files = "\n".join([os.path.basename(f) for f in non_existent_files])
            result = messagebox.askyesno("警告", f"以下のファイルが見つかりません。これらをスキップして続行しますか？\n\n{missing_files}")
            if not result:
                return
                
        # 現在の連番を保存
        start_seq = int(self.sequence_number.get())
        success_count = 0
        new_files = []
        
        # 保存先ファイル名の重複をチェック
        file_names_to_create = []
        for i, source_path in enumerate(self.selected_files):
            if not os.path.exists(source_path):
                continue
                
            # 連番を設定
            current_seq = start_seq + i
            self.sequence_number.set(str(current_seq))
            self.update_filename_preview()
            
            # ファイル名生成
            _, file_extension = os.path.splitext(source_path)
            new_filename = self.filename_pattern.get() + file_extension
            file_names_to_create.append(new_filename)
        
        # 重複ファイル名の検出
        duplicate_files = [item for item, count in 
                          {file_name: file_names_to_create.count(file_name) 
                           for file_name in file_names_to_create}.items() 
                          if count > 1]
        if duplicate_files:
            result = messagebox.askyesno("警告", 
                    f"生成されるファイル名に重複があります。このままでは一部のファイルが上書きされます。\n\n重複ファイル: {', '.join(duplicate_files)}\n\n続行しますか？")
            if not result:
                return
        
        # 本処理開始
        for i, source_path in enumerate(self.selected_files):
            # ファイルが存在しない場合はスキップ
            if not os.path.exists(source_path):
                new_files.append(source_path)  # 元のパスを保持
                continue
                
            # 連番を設定
            current_seq = start_seq + i
            self.sequence_number.set(str(current_seq))
            self.update_filename_preview()
            
            # 名前変更と移動処理
            _, file_extension = os.path.splitext(source_path)
            
            new_filename = self.filename_pattern.get() + file_extension
            destination_path = os.path.join(dest_dir, new_filename)
            
            # 絶対パスに変換して正規化
            source_path = os.path.abspath(os.path.normpath(source_path))
            destination_path = os.path.abspath(os.path.normpath(destination_path))
            
            # 既に同名のファイルが存在するかチェック（自分自身以外）
            if os.path.exists(destination_path) and source_path != destination_path:
                result = messagebox.askyesno("確認", f"既に同じ名前のファイルが存在します: {new_filename}\n上書きしますか？")
                if not result:
                    new_files.append(source_path)  # 元のパスを保持
                    continue
            
            try:
                shutil.move(source_path, destination_path)
                new_files.append(destination_path)
                success_count += 1
            except PermissionError:
                messagebox.showerror("エラー", f"ファイル '{new_filename}' へのアクセス権限がありません。")
                new_files.append(source_path)  # 元のパスを保持
            except FileNotFoundError:
                messagebox.showerror("エラー", f"ファイルが見つかりません: {os.path.basename(source_path)}")
                new_files.append(source_path)  # 元のパスを保持
            except Exception as e:
                new_files.append(source_path)  # 失敗した場合は元のパスを保持
                messagebox.showerror("エラー", f"ファイル '{os.path.basename(source_path)}' の名前変更と移動中にエラーが発生しました: {str(e)}")
        
        # 成功したファイルをリストに更新
        self.selected_files = new_files
        
        # リストボックスを更新
        self.files_listbox.delete(0, tk.END)
        for file in self.selected_files:
            self.files_listbox.insert(tk.END, os.path.basename(file))
        
        # 最終連番 + 1 を設定
        if self.auto_increment.get():
            self.sequence_number.set(str(start_seq + len(self.selected_files)))
            self.update_filename_preview()
        else:
            # 元の連番に戻す
            self.sequence_number.set(str(start_seq))
            self.update_filename_preview()
        
        self.status_var.set(f"{len(self.selected_files)}個中{success_count}個のファイル名を変更し移動しました")

    def save_settings(self):
        settings = {
            "filename_templates": self.filename_templates,
            "dest_templates": self.dest_templates,
            "sequence_digits": self.sequence_digits.get(),
            "auto_increment": self.auto_increment.get()
        }
        
    # キーボードイベントのデバッグ用メソッド
    def debug_key_event(self, event):
        """キーボードイベントのデバッグ情報を表示する"""
        key_info = f"キー情報: char={event.char}, keysym={event.keysym}, keycode={event.keycode}, state={event.state}"
        print(key_info)
        self.status_var.set(key_info)
        
        try:
            with open('file_manager_settings.json', 'w') as f:
                json.dump(settings, f)
        except Exception as e:
            messagebox.showerror("エラー", f"設定保存中にエラーが発生しました: {str(e)}")

    def load_settings(self):
        try:
            if os.path.exists('file_manager_settings.json'):
                with open('file_manager_settings.json', 'r') as f:
                    settings = json.load(f)
                
                if "filename_templates" in settings:
                    self.filename_templates = settings["filename_templates"]
                    self.update_filename_templates_combo()
                
                if "dest_templates" in settings:
                    self.dest_templates = settings["dest_templates"]
                    self.update_destination_templates_combo()
                
                if "sequence_digits" in settings:
                    self.sequence_digits.set(settings["sequence_digits"])
                
                if "auto_increment" in settings:
                    self.auto_increment.set(settings["auto_increment"])
        except Exception as e:
            messagebox.showerror("エラー", f"設定読み込み中にエラーが発生しました: {str(e)}")

def main():
    root = tk.Tk()
    app = FileManagerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()