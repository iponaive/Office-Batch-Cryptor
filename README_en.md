# Office-Batch-Cryptor

[English](README_en.md) | [日本語](README_ja.md) | [繁體中文](README.md) | [한국어](README_ko.md)

A PowerShell-based automation tool for batch encrypting and decrypting Excel & Word files. Supports custom password logic and simultaneous batch processing.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg) ![Office](https://img.shields.io/badge/Microsoft-Office-red.svg)

https://github.com/user-attachments/assets/fb2fbc2c-6dc2-40bb-85fb-841e0b1918b9







## Core Features
* **High Efficiency:** Process a large number of files at once when they share the same encryption or decryption password.
* **Simultaneous Processing:** Supports running encryption and decryption tasks at the same time.
* **Custom Keys:** Prioritizes custom `.txt` password files; automatically falls back to the default password (Kpmg123) if missing.
* **Group Key Logic:** Separate keys can be set for encryption and decryption tasks.
* **Format Support:** Compatible with Excel (.xlsx, .xlsm, .xls) and Word (.docx, .doc).

## Technical Highlights
* **Anti-Hanging Protection:** Double-detection (empty password test) during file opening to prevent Office authentication pop-ups from freezing the automation.
* **Resource Optimization:** Integrated `Marshal.ReleaseComObject` and .NET Garbage Collection to ensure no residual processes after heavy batch processing.
* **Security & Safety:** Auto-detects and removes Read-Only attributes, and implements a fallback mechanism for `SaveAs2` failures.

## Usage
1. Place files into the corresponding folders.
2. Run `run.bat` to start batch processing. A summary will be displayed upon completion.

## Project Structure
```text
.
├── AutoEncryptDecrypt.ps1    # Core automation logic
├── run.bat                   # Quick start batch file
├── input_to_encrypt/         # Path for files to be encrypted
├── input_to_decrypt/         # Path for files to be decrypted
├── encrypt_password.txt      # (Optional) Custom encryption password; defaults if empty
└── decrypt_password.txt      # (Optional) Custom decryption password; defaults if empty
