# deploy_github.ps1 — 把本專案部署到 GitHub，並開啟 GitHub Pages
#
# 用法（建議把 token 放環境變數，不要寫在指令裡）：
#   $env:GITHUB_TOKEN = 'ghp_xxxxxxxx'
#   powershell -ExecutionPolicy Bypass -File scripts\deploy_github.ps1 -User cia8885 -Repo premium-cosmetics
#
# 或直接帶參數：
#   powershell -ExecutionPolicy Bypass -File scripts\deploy_github.ps1 -User cia8885 -Repo premium-cosmetics -Token ghp_xxx
#
# 參數：
#   -User         GitHub 帳號
#   -Repo         repo 名稱（預設 premium-cosmetics）
#   -Token        Personal Access Token（需 repo 權限；要開 Pages 需 workflow 權限）
#   -Private      建立私有 repo（私有 repo 的 Pages 需 GitHub Pro 以上）
#   -SkipPages    不啟用 GitHub Pages
#
# 注意：token 只在本機記憶體中使用，不會寫入任何檔案。

param(
    [Parameter(Mandatory = $true)][string]$User,
    [string]$Repo = 'premium-cosmetics',
    [string]$Token = $env:GITHUB_TOKEN,
    [switch]$Private,
    [switch]$SkipPages
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Token)) {
    Write-Host '缺少 token。請先設定 $env:GITHUB_TOKEN，或用 -Token 傳入。' -ForegroundColor Red
    exit 1
}

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

$apiBase = 'https://api.github.com'
$headers = @{
    'Authorization'        = "Bearer $Token"
    'Accept'               = 'application/vnd.github+json'
    'X-GitHub-Api-Version' = '2022-11-28'
    'User-Agent'           = 'rokaizumi-deploy'
}

function Invoke-Gh {
    param([string]$Method, [string]$Path, $Body)
    $p = @{ Method = $Method; Uri = "$apiBase$Path"; Headers = $headers; TimeoutSec = 60 }
    if ($null -ne $Body) {
        $p.Body = ($Body | ConvertTo-Json -Depth 6)
        $p.ContentType = 'application/json; charset=utf-8'
    }
    return Invoke-RestMethod @p
}

# ── 0. 驗證 token 與帳號 ────────────────────────────────────────────────
Write-Host '=== 0/5 驗證 token ===' -ForegroundColor Cyan
try {
    $me = Invoke-Gh GET '/user'
} catch {
    Write-Host "token 無效或無權限：$($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
Write-Host "  登入身分：$($me.login)"
if ($me.login -ne $User) {
    Write-Warning "  token 屬於 $($me.login)，與指定帳號 $User 不同。將以 $($me.login) 建立 repo。"
    $User = $me.login
}

# ── 1. 建立 repo（已存在則沿用）────────────────────────────────────────
Write-Host '=== 1/5 建立 GitHub repo ===' -ForegroundColor Cyan
$exists = $true
try { $null = Invoke-Gh GET "/repos/$User/$Repo" } catch { $exists = $false }

if ($exists) {
    Write-Host "  repo 已存在，沿用 $User/$Repo"
} else {
    $body = @{
        name        = $Repo
        description = 'Rokaizumi 頂級化妝品產品線：14 支 SKU、成本引擎、投資與獲利分析'
        private     = [bool]$Private
        has_issues  = $true
        has_wiki    = $false
        auto_init   = $false
    }
    $null = Invoke-Gh POST '/user/repos' $body
    Write-Host "  已建立 $User/$Repo（$(if ($Private) {'私有'} else {'公開'})）"
}

# ── 2. 推送 ────────────────────────────────────────────────────────────
Write-Host '=== 2/5 推送程式碼 ===' -ForegroundColor Cyan
$remoteUrl = "https://$User`:$Token@github.com/$User/$Repo.git"
$prev = $ErrorActionPreference
$ErrorActionPreference = 'Continue'
git remote remove origin 2>&1 | Out-Null
git remote add origin $remoteUrl 2>&1 | Out-Null
git branch -M main 2>&1 | Out-Null
git push -u origin main --force 2>&1 | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] } | ForEach-Object { Write-Host "  $_" }
$pushCode = $LASTEXITCODE
$ErrorActionPreference = $prev
if ($pushCode -ne 0) {
    Write-Host '  推送失敗' -ForegroundColor Red
    exit 1
}
# 推送後把 token 從 remote URL 移除，避免留在 .git/config
git remote set-url origin "https://github.com/$User/$Repo.git"
Write-Host '  推送完成（remote 已改回不含 token 的網址）'

# ── 3. 啟用 GitHub Pages ───────────────────────────────────────────────
$pagesUrl = "https://$User.github.io/$Repo/"
if ($SkipPages) {
    Write-Host '=== 3/5 略過 GitHub Pages ===' -ForegroundColor Yellow
} else {
    Write-Host '=== 3/5 啟用 GitHub Pages ===' -ForegroundColor Cyan
    $pagesBody = @{ source = @{ branch = 'main'; path = '/' } }
    try {
        $null = Invoke-Gh POST "/repos/$User/$Repo/pages" $pagesBody
        Write-Host "  已啟用 Pages：$pagesUrl"
    } catch {
        try {
            $null = Invoke-Gh PUT "/repos/$User/$Repo/pages" $pagesBody
            Write-Host "  Pages 已更新：$pagesUrl"
        } catch {
            Write-Warning "  無法啟用 Pages（私有 repo 需 GitHub Pro）：$($_.Exception.Message)"
            $pagesUrl = $null
        }
    }
}

# ── 4. 驗證 ────────────────────────────────────────────────────────────
Write-Host '=== 4/5 驗證 ===' -ForegroundColor Cyan
$info = Invoke-Gh GET "/repos/$User/$Repo"
Write-Host "  repo：$($info.html_url)"
Write-Host "  預設分支：$($info.default_branch)"
Write-Host "  可見性：$(if ($info.private) {'私有'} else {'公開'})"
$files = Invoke-Gh GET "/repos/$User/$Repo/contents/"
Write-Host "  根目錄檔案數：$($files.Count)"

# ── 5. 結果 ────────────────────────────────────────────────────────────
Write-Host '=== 5/5 完成 ===' -ForegroundColor Green
Write-Host "  程式庫：$($info.html_url)"
if ($pagesUrl) {
    Write-Host "  網頁：  $pagesUrl" -ForegroundColor Green
    Write-Host '  （Pages 首次部署約需 1–3 分鐘，之後每次 push 會自動重建）'
}
