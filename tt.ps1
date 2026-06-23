<#
.SYNOPSIS
    TimeTrackerNX 実績登録スクリプト（Outlookカレンダー正）

.USAGE
    # 本日分を登録（Outlookソース・デフォルト）
    .\tt.ps1

    # 日付指定（YYMMDD）
    .\tt.ps1 260624

    # dry-runで確認のみ
    .\tt.ps1 260624 -Dry

    # 日報ファイルから登録（Outlookに入っていない残業・集中作業がある場合）
    .\tt.ps1 260624 -Diary
    .\tt.ps1 260624 -Diary -Dry

    # 既登録分を削除して再登録
    python output\tt_delete_date.py 2026-06-24
    .\tt.ps1 260624
#>
param(
    [string]$Date  = (Get-Date -Format "yyMMdd"),
    [switch]$Dry,
    [switch]$Diary   # 日報ソースを使う場合のみ指定。省略時はOutlook
)

$ScriptDir = $PSScriptRoot
$KeyFile   = Join-Path $ScriptDir "tt_apikey.txt"
$Python    = Join-Path $ScriptDir ".venv\Scripts\python.exe"
$Register  = Join-Path $ScriptDir "timetracker_register.py"

# APIキー読み込み
if (-not (Test-Path $KeyFile)) {
    Write-Error "APIキーファイルが見つかりません: $KeyFile"
    Write-Host "  tt_apikey.txt にAPIキーを1行目に貼り付けてください。"
    exit 1
}
$env:TT_API_KEY = (Get-Content $KeyFile -First 1).Trim()
if (-not $env:TT_API_KEY) {
    Write-Error "tt_apikey.txt が空です。APIキーを書き込んでください。"
    exit 1
}

# 実行
if ($Dry) {
    $mode = if ($Diary) { "diary" } else { "outlook" }
    Write-Host "[DRY-RUN] --date $Date --source $mode" -ForegroundColor Cyan
    & $Python $Register --date $Date --source $mode
} else {
    $mode = if ($Diary) { "diary" } else { "outlook" }
    Write-Host "[EXECUTE] --date $Date --source $mode" -ForegroundColor Green
    & $Python $Register --date $Date --source $mode --execute
    if ($LASTEXITCODE -ne 0) {
        $d = "20$($Date.Substring(0,2))-$($Date.Substring(2,2))-$($Date.Substring(4,2))"
        Write-Host ""
        Write-Host "[HINT] 重複エラーの場合は先に削除:" -ForegroundColor Yellow
        Write-Host "  python output\tt_delete_date.py $d" -ForegroundColor Yellow
    }
}
