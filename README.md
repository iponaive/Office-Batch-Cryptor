# Office-Batch-Cryptor
一款基於 PowerShell 實現的 Excel &amp; Word 批次自動化加解密工具。支援自定義密碼邏輯，並可「同時」執行大量檔案的加密與解密。
簡而言之，一次同時大量加密、大量解密檔案，密碼自選或預設。

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg) ![Office](https://img.shields.io/badge/Microsoft-Office-red.svg)

[English](README_en.md) | [日本語](README_ja.md) [繁體中文](README.md) |

## 核心功能
* 高效批次：當該批次檔案的加密密碼相同、或解密密碼相同，即可一次處理大量檔案。
* 同步處理：支援加密與解密任務同時進行。
* 密鑰自訂義：優先讀取自定義 .txt 密碼檔，缺失時自動回退至預設密碼 (Kpmg123)。
* 分組密鑰邏輯：可分別為加密與解密任務設定不同的金鑰。
* 檔案支援：相容於 Excel (.xlsx, .xlsm, .xls) 與 Word (.docx, .doc) 格式。

## 技術實現亮點
* 異常中斷防護：開啟檔案時採雙重偵測（空密碼測試），避免 Office 認證彈窗導致自動化程序掛起。
* 系統資源優化：整合 Marshal.ReleaseComObject 與 .NET 垃圾回收機制，確保大量處理後系統無殘留進程。
* 操作安全：自動檢測並解除唯讀屬性，並實作 SaveAs2 失敗時的 Fallback 回退機制。

## 使用方法
1. 將檔案放入對應資料夾。
2. 執行 run.bat 開始批次處理，完成後會顯示處理統計。

## 專案結構
```text
.
├── AutoEncryptDecrypt.ps1    # 核心自動化邏輯
├── run.bat                   # 快速啟動批次檔
├── input_to_encrypt/         # 待加密檔案路徑
├── input_to_decrypt/         # 待解密檔案路徑
├── encrypt_password.txt      # (選填) 自定義加密密碼；空白則使用預設密碼
└── decrypt_password.txt      # (選填) 自定義解密密碼；空白則使用預設密碼
