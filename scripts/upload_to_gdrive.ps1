# 上傳「頂級化妝品」專案到 Google Drive
# 建立日：2026-10-25
# 用法：powershell -ExecutionPolicy Bypass -File .\upload_to_gdrive.ps1
# 前置：rclone 已設定 gdrive: remote（type = drive）

$ErrorActionPreference = 'Stop'

# rclone 會把 NOTICE 寫到 stderr，Windows PowerShell 5.1 會當成終止錯誤。
# 因此原生指令一律加 --quiet，並在呼叫期間放寬 ErrorActionPreference。
function Invoke-Rclone {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$Args)
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & rclone @Args --quiet 2>&1 | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] }
        return @{ Code = $LASTEXITCODE; Output = $out }
    } finally {
        $ErrorActionPreference = $prev
    }
}

$LocalRoot  = 'C:\Users\emily\GroperBot\out\premium-cosmetics'
$RemoteRoot = 'gdrive:頂級化妝品'

$SubDirs = @(
    '00_法遵與智財',
    '01_產品規格',
    '02_成分與配方',
    '03_代工廠估價',
    '04_STO_RWA',
    '05_資料檔',
    '06_網頁',
    '07_投資分析',
    '07_投資分析/charts'
)

Write-Host '=== 1/4 建立雲端資料夾 ===' -ForegroundColor Cyan
Invoke-Rclone mkdir $RemoteRoot | Out-Null
foreach ($d in $SubDirs) {
    Invoke-Rclone mkdir "$RemoteRoot/$d" | Out-Null
    Write-Host "  建立 $RemoteRoot/$d"
}

Write-Host '=== 2/4 上傳文件 ===' -ForegroundColor Cyan

$Map = @(
    @{ Src = 'README.md';                              Dst = '' },
    @{ Src = 'docs\00_法遵與智財邊界.md';               Dst = '00_法遵與智財' },
    @{ Src = 'docs\01_品類對標分析.md';                 Dst = '01_產品規格' },
    @{ Src = 'docs\02_產品線規格_原創.md';              Dst = '01_產品規格' },
    @{ Src = 'docs\03_成分明細與比例_打樣版.md';        Dst = '02_成分與配方' },
    @{ Src = 'docs\04_代工廠估價需求單_RFQ.md';         Dst = '03_代工廠估價' },
    @{ Src = 'docs\05_STO_RWA_對映.md';                 Dst = '04_STO_RWA' },
    @{ Src = 'docs\06_成本分析.md';                     Dst = '03_代工廠估價' },
    @{ Src = 'data\products.json';                      Dst = '05_資料檔' },
    @{ Src = 'data\ingredients.csv';                    Dst = '05_資料檔' },
    @{ Src = 'data\rfq_items.csv';                      Dst = '05_資料檔' },
    @{ Src = 'data\ingredient_master.json';             Dst = '05_資料檔' },
    @{ Src = 'data\ingredient_prices.json';             Dst = '05_資料檔' },
    @{ Src = 'data\price_overrides.json';               Dst = '05_資料檔' },
    @{ Src = 'data\cost_detail.csv';                    Dst = '05_資料檔' },
    @{ Src = 'data\cost_report.csv';                    Dst = '05_資料檔' },
    @{ Src = 'data\cost_report.md';                     Dst = '03_代工廠估價' },
    @{ Src = 'scripts\fetch_prices.py';                 Dst = '05_資料檔' },
    @{ Src = 'scripts\cost_engine.py';                  Dst = '05_資料檔' },
    @{ Src = 'scripts\run_cost_analysis.ps1';           Dst = '05_資料檔' },
    @{ Src = 'data\finance_model.json';                 Dst = '05_資料檔' },
    @{ Src = 'scripts\finance_model.py';                Dst = '05_資料檔' },
    @{ Src = 'scripts\build_web.py';                    Dst = '05_資料檔' },
    @{ Src = 'scripts\build_docx.py';                   Dst = '05_資料檔' },
    @{ Src = 'scripts\finalize_docx.ps1';               Dst = '05_資料檔' },
    @{ Src = 'web\index.html';                          Dst = '06_網頁' },
    @{ Src = 'docs\07_投資與獲利分析報告.docx';         Dst = '07_投資分析' },
    @{ Src = 'docs\07_投資與獲利分析報告.pdf';          Dst = '07_投資分析' },
    @{ Src = 'docs\charts\01_cost_structure.png';       Dst = '07_投資分析/charts' },
    @{ Src = 'docs\charts\02_investment.png';           Dst = '07_投資分析/charts' },
    @{ Src = 'docs\charts\03_ai_team.png';              Dst = '07_投資分析/charts' },
    @{ Src = 'docs\charts\04_waterfall.png';            Dst = '07_投資分析/charts' },
    @{ Src = 'docs\charts\05_pricing.png';              Dst = '07_投資分析/charts' },
    @{ Src = 'docs\charts\06_pnl.png';                  Dst = '07_投資分析/charts' }
)

$Failed = @()
foreach ($m in $Map) {
    $src = Join-Path $LocalRoot $m.Src
    if (-not (Test-Path $src)) {
        Write-Warning "  找不到 $src"
        $Failed += $m.Src
        continue
    }
    $dst = if ([string]::IsNullOrEmpty($m.Dst)) { $RemoteRoot } else { "$RemoteRoot/$($m.Dst)" }
    $r = Invoke-Rclone copy $src $dst --no-traverse
    if ($r.Code -ne 0) {
        Write-Warning "  上傳失敗：$($m.Src)"
        $Failed += $m.Src
    } else {
        Write-Host "  OK  $($m.Src)  ->  $dst"
    }
}

Write-Host '=== 3/4 驗證雲端內容 ===' -ForegroundColor Cyan
$tree = Invoke-Rclone tree $RemoteRoot --dirs-only
$tree.Output | ForEach-Object { Write-Host "  $_" }

Write-Host '=== 4/4 逐檔比對大小 ===' -ForegroundColor Cyan
$ok = 0; $bad = 0
foreach ($m in $Map) {
    $src = Join-Path $LocalRoot $m.Src
    if (-not (Test-Path $src)) { continue }
    $name = Split-Path $src -Leaf
    $dst = if ([string]::IsNullOrEmpty($m.Dst)) { $RemoteRoot } else { "$RemoteRoot/$($m.Dst)" }
    $r = Invoke-Rclone size "$dst/$name" --json
    $remoteSize = $null
    try { $remoteSize = ($r.Output | Out-String | ConvertFrom-Json).bytes } catch { $remoteSize = $null }
    $localSize  = (Get-Item $src).Length
    if ($remoteSize -eq $localSize) {
        Write-Host "  OK  $name  ($localSize bytes)"
        $ok++
    } else {
        Write-Warning "  大小不符 $name  本機=$localSize  雲端=$remoteSize"
        $bad++
    }
}

Write-Host ''
Write-Host "完成：$ok 個檔案一致，$bad 個不一致，$($Failed.Count) 個失敗" -ForegroundColor Green
if ($bad -gt 0 -or $Failed.Count -gt 0) { exit 1 }
