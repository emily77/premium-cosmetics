# 頂級化妝品線 — 原料成本分析一鍵執行
#
# 流程：L1 自動抓價 → 三層價格回退 → 成本計算 → 產出報告
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File scripts\run_cost_analysis.ps1
#   powershell -ExecutionPolicy Bypass -File scripts\run_cost_analysis.ps1 -Refresh -UsdTwd 32.5
#
# 參數：
#   -Refresh   忽略既有抓取結果，全部重抓（約 7 分鐘）
#   -UsdTwd    指定 USD→TWD 匯率（預設 32.0）
#   -SkipFetch 跳過抓價，只用現有資料算成本

[CmdletBinding()]
param(
    [switch]$Refresh,
    [double]$UsdTwd = 32.0,
    [switch]$SkipFetch
)

$ErrorActionPreference = 'Stop'
$env:PYTHONIOENCODING = 'utf-8'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root

function Write-Step($msg) {
    Write-Host ""
    Write-Host "=== $msg ===" -ForegroundColor Cyan
}

Write-Host "頂級化妝品線 — 原料成本分析" -ForegroundColor Green
Write-Host "工作目錄：$Root"
Write-Host "匯率：USD 1 = TWD $UsdTwd"

# ---------- 前置檢查 ----------
Write-Step "前置檢查"
$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) { throw "找不到 python，請先安裝 Python 3.10+" }
Write-Host "Python：$((python --version 2>&1))"

$required = @(
    'data\products.json',
    'data\ingredients.csv',
    'data\ingredient_master.json'
)
foreach ($f in $required) {
    if (-not (Test-Path $f)) { throw "缺少必要檔案：$f" }
    Write-Host "OK  $f"
}

# ---------- 步驟 1：抓價 ----------
if (-not $SkipFetch) {
    Write-Step "步驟 1／2：抓取原料公開價格（L1）"
    $fetchArgs = @('scripts\fetch_prices.py')
    if ($Refresh) { $fetchArgs += '--refresh' }
    & python @fetchArgs
    if ($LASTEXITCODE -ne 0) {
        Write-Host "警告：抓價未完全成功（exit $LASTEXITCODE），將以現有資料繼續" -ForegroundColor Yellow
    }
} else {
    Write-Step "步驟 1／2：跳過抓價（-SkipFetch）"
}

# ---------- 步驟 2：算成本 ----------
Write-Step "步驟 2／2：計算成本並產出報告"
& python scripts\cost_engine.py --usd-twd $UsdTwd
if ($LASTEXITCODE -ne 0) { throw "成本計算失敗（exit $LASTEXITCODE）" }

# ---------- 結果摘要 ----------
Write-Step "產出檔案"
$outputs = @(
    'data\ingredient_prices.json',
    'data\cost_detail.csv',
    'data\cost_report.csv',
    'data\cost_report.md'
)
foreach ($f in $outputs) {
    if (Test-Path $f) {
        $size = (Get-Item $f).Length
        Write-Host ("OK  {0,-32} {1,8:N0} bytes" -f $f, $size) -ForegroundColor Green
    } else {
        Write-Host "MISSING  $f" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "完成。請開啟 data\cost_report.md 檢視結果。" -ForegroundColor Green
Write-Host "提醒：本報告為估算值，實際成本以代工廠正式報價為準。" -ForegroundColor Yellow
