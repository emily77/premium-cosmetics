# finalize_docx.ps1 — 用本機 Word 更新目錄並（可選）匯出 PDF
#
# build_docx.py 產生的 .docx 內含 TOC 功能變數，但功能變數要由 Word 更新才會有內容。
# 本腳本用 Word COM 分兩段處理：
#   1. 開啟（可寫）→ 更新所有功能變數與目錄 → 存檔 → 關閉
#   2. 開啟（唯讀）→ 匯出 PDF → 關閉
# 分成兩段是因為在同一個 Word 工作階段裡「存檔後立刻匯出」實測會卡住。
# 匯出用 SaveAs2(路徑, 17) 而非 ExportAsFixedFormat —— 後者在本機實測會無限卡住。
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File scripts\finalize_docx.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\finalize_docx.ps1 -Pdf
#
# 注意：需要本機安裝 Microsoft Word。若沒有 Word，腳本會提示並直接結束（不視為失敗）。

param(
    [switch]$Pdf
)

$ErrorActionPreference = 'Stop'

$root    = Split-Path -Parent $PSScriptRoot
$docx    = Join-Path $root 'docs\07_投資與獲利分析報告.docx'
$pdfPath = Join-Path $root 'docs\07_投資與獲利分析報告.pdf'

if (-not (Test-Path $docx)) {
    Write-Host "找不到 $docx，請先執行 python scripts\build_docx.py" -ForegroundColor Red
    exit 1
}

function Get-WordApp {
    try {
        $w = New-Object -ComObject Word.Application
        $w.Visible = $false
        $w.DisplayAlerts = 0
        return $w
    } catch {
        return $null
    }
}

# ── 第 1 段：更新目錄並存檔 ────────────────────────────────────────────
$word = Get-WordApp
if ($null -eq $word) {
    Write-Host "本機沒有 Microsoft Word，略過目錄更新（Word 開啟時會自動提示更新）。" -ForegroundColor Yellow
    exit 0
}

try {
    $doc = $word.Documents.Open($docx, $false, $false)
    $doc.Fields.Update() | Out-Null
    foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
    $doc.Repaginate()
    $pages = $doc.ComputeStatistics(2)
    $doc.Save()
    $doc.Close(0)
    Write-Host "目錄已更新，共 $pages 頁" -ForegroundColor Green
} finally {
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    [GC]::Collect()
    Start-Sleep -Milliseconds 800
}

# ── 第 2 段：匯出 PDF ──────────────────────────────────────────────────
if ($Pdf) {
    $word2 = Get-WordApp
    if ($null -eq $word2) {
        Write-Host "沒有 Word，無法匯出 PDF。" -ForegroundColor Yellow
        exit 0
    }
    try {
        if (Test-Path $pdfPath) { Remove-Item $pdfPath -Force }
        $word2.Options.UpdateFieldsAtPrint = $false
        $doc2 = $word2.Documents.Open($docx, $false, $true, $false)
        $doc2.SaveAs2($pdfPath, 17)
        $doc2.Close(0)
        Write-Host "PDF → $pdfPath" -ForegroundColor Green
    } finally {
        $word2.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word2) | Out-Null
        [GC]::Collect()
    }
}

Write-Host "完成 → $docx" -ForegroundColor Green
