# ファイル管理ユーティリティ / File Management Utility

## 概要 / Overview

このアプリケーションは、ファイル名の変更や移動を効率的に行うためのユーティリティツールです。日付、連番、カスタムテキストなどを組み合わせたファイル名パターンを作成し、単一または複数のファイルに適用することができます。

This application is a utility tool for efficiently renaming and moving files. It allows you to create filename patterns combining date, sequence numbers, and custom text, which can be applied to single or multiple files.

## システム要件 / System Requirements

- Windows 7以降
- 最小メモリ: 256MB
- ディスク容量: 50MB以上の空き容量

## インストール方法 / Installation


### ソースコードからのビルド / Building from Source

1. リポジトリをクローンまたはダウンロードします。
2. 必要なパッケージをインストールします：
   ```
   pip install pyperclip
   ```
3. PyInstallerをインストールします：
   ```
   pip install pyinstaller
   ```
4. 以下のコマンドでexeファイルを作成します：
   ```
   pyinstaller --onefile --noconsole file_manager.py
   ```
5. 生成された`dist`フォルダ内の`file_manager.exe`が実行ファイルです。

## 主な機能 / Main Features

- **ファイル名パターン作成**: 日付、連番、カスタムテキストを組み合わせたファイル名パターンを作成
- **テンプレート管理**: よく使うファイル名パターンを保存・読み込み
- **一括処理**: 複数ファイルの名前変更・移動を一度に実行
- **保存先管理**: 頻繁に使用する保存先をテンプレートとして保存
- **自動連番**: ファイルのコピーや名前変更時に連番を自動的に増加
- **ショートカットキー**: 主要な機能にショートカットキーを割り当て

## 使用方法 / How to Use

### ファイル名パターンの作成 / Creating Filename Patterns

1. 「ファイル名設定」タブを選択します。
2. 日付フォーマット、連番、カスタムテキストなどを設定します。
3. 「日付を挿入」「連番を挿入」「テキストを挿入」ボタンをクリックして、パターンエディタに挿入します。
4. 「プレビュー更新」をクリックするか F5 キーを押して、結果を確認します。

### ファイル操作 / File Operations

1. 「ファイル操作」タブを選択します。
2. 「ファイル参照」ボタンをクリックして操作対象のファイルを選択します。
3. 必要に応じて「保存先設定」タブで保存先を選択します。
4. 以下の操作を実行できます：
   - ファイル名をクリップボードにコピー (Ctrl+C)
   - ファイル名変更 (Ctrl+R)
   - ファイル移動 (Ctrl+M)
   - ファイル名変更と移動 (Ctrl+Shift+M)

### 一括処理 / Batch Processing

1. 「ファイル操作」タブで「一括処理モード」にチェックを入れます。
2. 「ファイル追加」ボタンをクリックして、複数のファイルを選択します。
3. ファイル名パターンと保存先を設定します。
4. 必要な操作を実行します。

## トラブルシューティング / Troubleshooting

### 設定ファイルが保存されない場合 / Settings File Not Saving

このアプリケーションは設定ファイル（`file_manager_settings.json`）をexeファイルと同じディレクトリに保存します。以下の点を確認してください：

1. アプリケーションがあるフォルダに書き込み権限があることを確認します。
2. 管理者権限で実行してみてください（右クリック→「管理者として実行」）。
3. フォルダが読み取り専用になっていないか確認します。

### 一般的な問題 / General Issues

- **アプリケーションが起動しない**: .NET Frameworkがインストールされていることを確認してください。
- **ファイル操作ができない**: ファイルへのアクセス権限を確認してください。
- **日本語が文字化けする**: システムの言語設定を確認してください。

## ショートカットキー一覧 / Shortcut Keys

| キー | 機能 |
|------|------|
| F5 | プレビュー更新 |
| Ctrl+C | クリップボードにコピー |
| Ctrl+R | ファイル名変更 |
| Ctrl+M | ファイル移動 |
| Ctrl+Shift+M | ファイル名変更と移動 |
| Ctrl+O | ファイル選択 |
| Ctrl+D | 保存先選択 |
| Ctrl+S | テンプレート保存 |
| Ctrl+L | テンプレート読込 |

## ライセンス / License

このソフトウェアはMITライセンスの下で公開されています。詳細はLICENSEファイルを参照してください。

## 作者 / Author

[inata169] - 最終更新: 2025年3月25日
