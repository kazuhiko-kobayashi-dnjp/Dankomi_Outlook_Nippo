# st.ps1 - Wrapper for schedule_tasks.py
#
# Usage:
#   .\st.ps1              # dry-run (current week, 露出制御業務進捗管理[Web]から読込・デフォルト)
#   .\st.ps1 260623       # dry-run (week containing 260623)
#   .\st.ps1 -Execute     # apply to Outlook (current week)
#   .\st.ps1 260623 -Execute
#   .\st.ps1 -Clear            # dry-run: show tasks to delete
#   .\st.ps1 -Clear -Execute   # delete auto-registered tasks, then reschedule
#   .\st.ps1 -Excel            # 旧・進捗管理Excelを直接読み込む（Web版を使わない場合）
#   .\st.ps1 -Excel -Execute
#
# 配布先の人向け（担当者名が「小林」固定ではなく自分の名前でタスクを拾いたい場合）:
#   .\st.ps1 -Person "自分の名字"                    # その場限りで担当者名を上書き
#   .\st.ps1 -Person "自分の名字" -SavePerson          # tools\.user_config.json に保存（次回から省略可）
#   .\st.ps1 -Json "C:\path\to\tasks.json"            # tasks.json の場所を明示指定（省略時はリポジトリ相対パス）
#   .\st.ps1 -Url "http://10.41.55.204:3100/api/tasks" # OneDrive同期パスが無くてもHTTP経由でサーバーから直接取得
#   .\st.ps1 -Excel -File "C:\path\to\進捗.xlsx"      # -Excel使用時、Excelパスも明示指定可能

param(
    [string]$Week    = '',
    [switch]$Execute,
    [switch]$Clear,
    [switch]$Excel,
    [string]$Person      = '',
    [switch]$SavePerson,
    [string]$Json        = '',
    [string]$Url         = '',
    [string]$File        = ''
)

Set-Location $PSScriptRoot

$venv_py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venv_py)) {
    Write-Error '.venv not found. Run: python -m venv .venv'
    exit 1
}

$args_list = @()
if ($Week)    { $args_list += '--week'; $args_list += $Week }
if ($Execute) { $args_list += '--execute' }
if ($Clear)   { $args_list += '--clear' }
if ($Excel)   { $args_list += '--source'; $args_list += 'excel' }
if ($Person)     { $args_list += '--person'; $args_list += $Person }
if ($SavePerson)  { $args_list += '--save-person' }
if ($Json)     { $args_list += '--json'; $args_list += $Json }
if ($Url)      { $args_list += '--url'; $args_list += $Url }
if ($File)     { $args_list += '--file'; $args_list += $File }

& $venv_py schedule_tasks.py @args_list