# Office-Batch-Cryptor

[English](README.md) | [日本語](README_ja.md) | [繁體中文](README_zh.md) | [한국어](README_ko.md)

---

PowerShell 기반의 Excel & Word 배치 자동 암호화/복호화 도구입니다. 사용자 정의 비밀번호 로직을 지원하며, 대량의 파일을 동시에 처리할 수 있습니다.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg) ![Office](https://img.shields.io/badge/Microsoft-Office-red.svg)

https://github.com/user-attachments/assets/fb2fbc2c-6dc2-40bb-85fb-841e0b1918b9

## 핵심 기능
* **고효율 배치 처리:** 동일한 암호화 또는 복호화 비밀번호를 사용하는 경우, 대량의 파일을 한 번에 처리할 수 있습니다.
* **동시 처리 지원:** 암호화 및 복호화 작업을 동시에 수행할 수 있습니다.
* **사용자 정의 키:** 사용자 정의 `.txt` 비밀번호 파일을 우선적으로 읽으며, 파일이 없을 경우 기본 비밀번호(Kpmg123)로 자동 대체됩니다.
* **그룹별 키 로직:** 암호화와 복호화 작업에 각각 다른 키를 설정할 수 있습니다.
* **지원 형식:** Excel(.xlsx, .xlsm, .xls) 및 Word(.docx, .doc) 형식을 지원합니다.

## 기술적 특징
* **응답 없음 방지:** 파일 오픈 시 이중 감지(빈 비밀번호 테스트)를 통해 Office 인증 팝업으로 인한 자동화 프로그램 중단을 방지합니다.
* **시스템 리소스 최적화:** `Marshal.ReleaseComObject`와 .NET 가비지 컬렉션(GC)을 통합하여 대량 처리 후에도 시스템에 잔류 프로세스가 남지 않도록 설계되었습니다.
* **안전한 작업:** 읽기 전용 속성을 자동으로 감지하여 해제하며, `SaveAs2` 실패 시 복구(Fallback) 메커니즘을 구현했습니다.

## 사용 방법
1. 파일을 대응하는 폴더에 넣습니다.
2. `run.bat`를 실행하여 배치 처리를 시작합니다. 완료 후 처리 통계가 표시됩니다.

## 프로젝트 구조
```text
.
├── AutoEncryptDecrypt.ps1    # 핵심 자동화 로직
├── run.bat                   # 빠른 실행 배치 파일
├── input_to_encrypt/         # 암호화 대상 파일 경로
├── input_to_decrypt/         # 복호화 대상 파일 경로
├── encrypt_password.txt      # (선택 사항) 사용자 정의 암호화 비밀번호, 비어있을 경우 기본값 사용
└── decrypt_password.txt      # (선택 사항) 사용자 정의 복호화 비밀번호, 비어있을 경우 기본값 사용
