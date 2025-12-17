
<#
批次加密/解密 Excel/Word（以資料夾決定動作；覆蓋原檔）
- input_to_encrypt：加密（套用 encrypt 密碼）
- input_to_decrypt：解密（移除密碼）
- 密碼檔（只讀第一行；缺失或空白 → 預設 Kpmg123）：
  * encrypt_password.txt（加密用；可選）
  * decrypt_password.txt（解密用；可選）
- 避免彈窗：Open 時先以空密碼 "" 嘗試，失敗再用解密密碼；DisplayAlerts 關閉
- 相容：SaveAs2 不支援時回退 SaveAs
- 支援：Excel .xlsx/.xlsm/.xls；Word .docx/.doc
#>

$ErrorActionPreference = "Stop"

# === 路徑 ===
$basePath        = Split-Path -Parent $MyInvocation.MyCommand.Path
$encFolder       = Join-Path $basePath "input_to_encrypt"
$decFolder       = Join-Path $basePath "input_to_decrypt"
$pwdEncFile      = Join-Path $basePath "encrypt_password.txt"
$pwdDecFile      = Join-Path $basePath "decrypt_password.txt"

# === 確保資料夾存在 ===
foreach ($p in @($encFolder,$decFolder)) {
    if (-not (Test-Path $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

# === 密碼讀取（缺失/空白 → 預設 Kpmg123） ===
function Get-FirstLineOrDefault([string]$path, [string]$defaultVal) {
    if (Test-Path $path) {
        $firstLine = (Get-Content -Path $path -Encoding UTF8 -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($null -ne $firstLine) {
            $value = $firstLine.Trim()
            if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
        }
    }
    return $defaultVal
}
$DEFAULT_PWD       = "Kpmg123"   # ← 只有 K 大寫
$encryptPassword   = Get-FirstLineOrDefault $pwdEncFile $DEFAULT_PWD
$decryptPassword   = Get-FirstLineOrDefault $pwdDecFile $DEFAULT_PWD

Write-Host "加密密碼：$encryptPassword（來源：encrypt_password.txt 或預設）" -ForegroundColor Cyan
Write-Host "解密密碼：$decryptPassword（來源：decrypt_password.txt 或預設）" -ForegroundColor Cyan

# === 取得檔案清單 ===
$encFilesAll = Get-ChildItem -Path (Join-Path $encFolder '*') -File -ErrorAction SilentlyContinue
$decFilesAll = Get-ChildItem -Path (Join-Path $decFolder '*') -File -ErrorAction SilentlyContinue

$encExcelFiles = $encFilesAll | Where-Object { $_.Extension.ToLower() -in @('.xlsx','.xlsm','.xls') }
$encWordFiles  = $encFilesAll | Where-Object { $_.Extension.ToLower() -in @('.docx','.doc') }
$decExcelFiles = $decFilesAll | Where-Object { $_.Extension.ToLower() -in @('.xlsx','.xlsm','.xls') }
$decWordFiles  = $decFilesAll | Where-Object { $_.Extension.ToLower() -in @('.docx','.doc') }

if (($encExcelFiles.Count + $encWordFiles.Count + $decExcelFiles.Count + $decWordFiles.Count) -eq 0) {
    Write-Host "兩個資料夾皆無可處理的 Excel/Word 檔案。" -ForegroundColor Yellow
    pause
    exit
}

# === 解除唯讀屬性（避免另存失敗） ===
foreach ($f in ($encExcelFiles + $encWordFiles + $decExcelFiles + $decWordFiles)) {
    try { if ($f.IsReadOnly) { attrib -R $f.FullName } } catch {}
}

# === 輔助：Excel/Word 格式常數 ===
function Get-ExcelFormatCode([string]$extension) {
    switch ($extension.ToLower()) {
        ".xlsx" { return 51 } # xlOpenXMLWorkbook
        ".xlsm" { return 52 } # xlOpenXMLWorkbookMacroEnabled
        ".xls"  { return 56 } # xlExcel8
        default { return $null }
    }
}
function Get-WordFormatCode([string]$extension) {
    switch ($extension.ToLower()) {
        ".docx" { return 16 } # wdFormatDocumentDefault
        ".doc"  { return 0 }  # wdFormatDocument
        default { return $null }
    }
}

# === 統計 ===
$cntEncXls=0; $cntEncDoc=0; $cntDecXls=0; $cntDecDoc=0; $cntErrXls=0; $cntErrDoc=0

# ======================
# 👉 ENCRYPT：Excel（覆蓋原檔）
# ======================
if ($encExcelFiles.Count -gt 0) {
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible        = $false
        $excel.DisplayAlerts  = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents   = $false

        foreach ($f in $encExcelFiles) {
            $wb = $null
            $format = Get-ExcelFormatCode $f.Extension
            if ($null -eq $format) {
                Write-Host "略過未知 Excel 格式（加密）：$($f.Name)" -ForegroundColor Yellow
                continue
            }

            Write-Host "加密 Excel：$($f.FullName)"

            # 先以空密碼開啟；失敗再用 decryptPassword（避免彈窗）
            try {
                $wb = $excel.Workbooks.Open($f.FullName, 0, $false, 5, "", "", $true)
            } catch {
                try {
                    $wb = $excel.Workbooks.Open($f.FullName, 0, $false, 5, $decryptPassword, "", $true)
                } catch {
                    Write-Host "→ 無法開啟（密碼錯誤或檔案毀損）：$($f.Name)" -ForegroundColor Red
                    $cntErrXls++; continue
                }
            }

            # 覆蓋原檔（套用加密密碼）
            try {
                $saveOk = $false
                try {
                    $wb.SaveAs2($f.FullName, $format, $encryptPassword, [Type]::Missing, $false, $false, 1, 2, $false, [Type]::Missing, [Type]::Missing, $true)
                    $saveOk = $true
                } catch {
                    $wb.SaveAs($f.FullName, $format, $encryptPassword, [Type]::Missing, $false, $false, 1, 2, $false, [Type]::Missing, [Type]::Missing, $true)
                    $saveOk = $true
                }
                if ($saveOk) { $cntEncXls++; Write-Host "→ 已加密覆蓋：$($f.Name)" -ForegroundColor Green }
            } catch {
                Write-Host "Excel 加密失敗：$($f.Name)｜$($_.Exception.Message)" -ForegroundColor Red
                $cntErrXls++
            } finally {
                try { if ($wb) { $wb.Close($false) } } catch {}
            }
        }
    } finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

# ======================
# 👉 ENCRYPT：Word（覆蓋原檔）
# ======================
if ($encWordFiles.Count -gt 0) {
    $word = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible        = $false
        $word.DisplayAlerts  = 0
        $word.ScreenUpdating = $false

        foreach ($f in $encWordFiles) {
            $doc = $null
            $format = Get-WordFormatCode $f.Extension
            if ($null -eq $format) {
                Write-Host "略過未知 Word 格式（加密）：$($f.Name)" -ForegroundColor Yellow
                continue
            }

            Write-Host "加密 Word：$($f.FullName)"

            # 先以空密碼開啟；失敗再用 decryptPassword
            try {
                $doc = $word.Documents.Open($f.FullName, $false, $false, $false, "", "", $false)
            } catch {
                try {
                    $doc = $word.Documents.Open($f.FullName, $false, $false, $false, $decryptPassword, "", $false)
                } catch {
                    Write-Host "→ 無法開啟（密碼錯誤或檔案毀損）：$($f.Name)" -ForegroundColor Red
                    $cntErrDoc++; continue
                }
            }

            # 覆蓋原檔（套用加密密碼）
            try {
                $saveOk = $false
                try {
                    $doc.SaveAs2($f.FullName, $format, $false, $encryptPassword)
                    $saveOk = $true
                } catch {
                    $doc.SaveAs($f.FullName, $format, $false, $encryptPassword)
                    $saveOk = $true
                }
                if ($saveOk) { $cntEncDoc++; Write-Host "→ 已加密覆蓋：$($f.Name)" -ForegroundColor Green }
            } catch {
                Write-Host "Word 加密失敗：$($f.Name)｜$($_.Exception.Message)" -ForegroundColor Red
                $cntErrDoc++
            } finally {
                try { if ($doc) { $doc.Close($false) } } catch {}
            }
        }
    } finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

# ======================
# 👉 DECRYPT：Excel（覆蓋原檔）
# ======================
if ($decExcelFiles.Count -gt 0) {
    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible        = $false
        $excel.DisplayAlerts  = $false
        $excel.ScreenUpdating = $false
        $excel.EnableEvents   = $false

        foreach ($f in $decExcelFiles) {
            $wb = $null
            $format = Get-ExcelFormatCode $f.Extension
            if ($null -eq $format) {
                Write-Host "略過未知 Excel 格式（解密）：$($f.Name)" -ForegroundColor Yellow
                continue
            }

            Write-Host "解密 Excel：$($f.FullName)"

            # 先以解密密碼開啟；若未加密亦可用空密碼開
            try {
                try {
                    # 嘗試用解密密碼
                    $wb = $excel.Workbooks.Open($f.FullName, 0, $false, 5, $decryptPassword, "", $true)
                } catch {
                    # 若未加密，改用空密碼
                    $wb = $excel.Workbooks.Open($f.FullName, 0, $false, 5, "", "", $true)
                }
            } catch {
                Write-Host "→ 無法開啟（密碼錯誤或檔案毀損）：$($f.Name)" -ForegroundColor Red
                $cntErrXls++; continue
            }

            # 覆蓋原檔（移除密碼）
            try {
                $saveOk = $false
                try {
                    $wb.SaveAs2($f.FullName, $format, "", [Type]::Missing, $false, $false, 1, 2, $false, [Type]::Missing, [Type]::Missing, $true)
                    $saveOk = $true
                } catch {
                    $wb.SaveAs($f.FullName, $format, "", [Type]::Missing, $false, $false, 1, 2, $false, [Type]::Missing, [Type]::Missing, $true)
                    $saveOk = $true
                }
                if ($saveOk) { $cntDecXls++; Write-Host "→ 已解密覆蓋：$($f.Name)" -ForegroundColor Green }
            } catch {
                Write-Host "Excel 解密失敗：$($f.Name)｜$($_.Exception.Message)" -ForegroundColor Red
                $cntErrXls++
            } finally {
                try { if ($wb) { $wb.Close($false) } } catch {}
            }
        }
    } finally {
        if ($excel) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

# ======================
# 👉 DECRYPT：Word（覆蓋原檔）
# ======================
if ($decWordFiles.Count -gt 0) {
    $word = $null
    try {
        $word = New-Object -ComObject Word.Application
        $word.Visible        = $false
        $word.DisplayAlerts  = 0
        $word.ScreenUpdating = $false

        foreach ($f in $decWordFiles) {
            $doc = $null
            $format = Get-WordFormatCode $f.Extension
            if ($null -eq $format) {
                Write-Host "略過未知 Word 格式（解密）：$($f.Name)" -ForegroundColor Yellow
                continue
            }

            Write-Host "解密 Word：$($f.FullName)"

            # 先以解密密碼開啟；未加密則用空密碼
            try {
                try {
                    $doc = $word.Documents.Open($f.FullName, $false, $false, $false, $decryptPassword, "", $false)
                } catch {
                    $doc = $word.Documents.Open($f.FullName, $false, $false, $false, "", "", $false)
                }
            } catch {
                Write-Host "→ 無法開啟（密碼錯誤或檔案毀損）：$($f.Name)" -ForegroundColor Red
                $cntErrDoc++; continue
            }

            # 覆蓋原檔（移除密碼）
            try {
                $saveOk = $false
                try {
                    $doc.SaveAs2($f.FullName, $format, $false, "")
                    $saveOk = $true
                } catch {
                    $doc.SaveAs($f.FullName, $format, $false, "")
                    $saveOk = $true
                }
                if ($saveOk) { $cntDecDoc++; Write-Host "→ 已解密覆蓋：$($f.Name)" -ForegroundColor Green }
            } catch {
                Write-Host "Word 解密失敗：$($f.Name)｜$($_.Exception.Message)" -ForegroundColor Red
                $cntErrDoc++
            } finally {
                try { if ($doc) { $doc.Close($false) } } catch {}
            }
        }
    } finally {
        if ($word) {
            $word.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
    }
}

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host ""
Write-Host "全部處理完成。" -ForegroundColor Cyan
Write-Host "加密覆蓋：Excel $cntEncXls，Word $cntEncDoc" -ForegroundColor Cyan
Write-Host "解密覆蓋：Excel $cntDecXls，Word $cntDecDoc" -ForegroundColor Cyan
Write-Host "錯誤：Excel $cntErrXls，Word $cntErrDoc" -ForegroundColor Cyan
pause
