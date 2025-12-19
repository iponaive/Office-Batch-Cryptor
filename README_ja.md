# Office-Batch-Cryptor

[English](README_en.md) | [日本語](README_ja.md) | [繁體中文](README.md) | [한국어](README_ko.md)

---

PowerShell ベースの Excel & Word 一括自動暗号化・復号ツールです。カスタムパスワードロジックに対応し、大量のファイルを同時に処理することが可能です。

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg) ![Office](https://img.shields.io/badge/Microsoft-Office-red.svg)

https://github.com/user-attachments/assets/fb2fbc2c-6dc2-40bb-85fb-841e0b1918b9

## 主な機能
* **効率的な一括処理:** 同じ暗号化パスワード、または同じ復号パスワードを使用する場合、大量のファイルを一度に処理できます。
* **同時処理:** 暗号化タスクと復号タスクを同時に実行できます。
* **カスタムキー:** カスタムの `.txt` パスワードファイルを優先的に読み込みます。ファイルがない場合は、デフォルトのパスワード（Kpmg123）が自動的に使用されます。
* **グループキーロジック:** 暗号化と復号のタスクにそれぞれ異なるキーを設定可能です。
* **対応フォーマット:** Excel (.xlsx, .xlsm, .xls) および Word (.docx, .doc) に対応しています。

## 技術的な特徴
* **ハングアップ防止機能:** ファイルを開く際に二重検出（空パスワードテスト）を行い、Office の認証ポップアップによる自動化プログラムの停止を回避します。
* **システムリソースの最適化:** `Marshal.ReleaseComObject` と .NET のガベージコレクションを統合し、大量処理後もシステムに残存プロセスが発生しないように設計されています。
* **安全な操作:** 読み取り専用属性を自動的に検出して解除し、`SaveAs2` が失敗した際のフォールバック（切り戻し）メカزمを実装しています。

## 使用方法
1. 対象ファイルを対応するフォルダに入れます。
2. `run.bat` を実行して一括処理を開始します。完了後、処理統計が表示されます。

## プロジェクト構成
```text
.
├── AutoEncryptDecrypt.ps1    # コア自動化ロジック
├── run.bat                   # クイックスタートバッチファイル
├── input_to_encrypt/         # 暗号化待ちファイルパス
├── input_to_decrypt/         # 復号待ちファイルパス
├── encrypt_password.txt      # (任意) カスタム暗号化パスワード。空の場合はデフォルトを使用
└── decrypt_password.txt      # (任意) カスタム復号パスワード。空の場合はデフォルトを使用
