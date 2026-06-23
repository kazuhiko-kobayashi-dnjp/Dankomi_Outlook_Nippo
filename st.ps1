# st.ps1 — schedule_tasks.py のワンコマンド起動スクリプト
#
# 使い方:
#   .\st.ps1              # 今週をdry-run確認
#   .\st.ps1 260623       # 260623 を含む週をdry-run確認
#   .\st.ps1 -Execute     # 今週に実際にOutlookへ登録
#   .\st.ps1 260623 -Execute

param(
    [string]$Week    = '',
    [switch]$Execute
)

Set-Location $PSScriptRoot

$venv_py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venv_py)) {
    Write-Error '.venv が見つかりません。先に python -m venv .venv を実行してください。'
    exit 1
}

$args_list = @()
if ($Week)    { $args_list += '--week'; $args_list += $Week }
if ($Execute) { $args_list += '--execute' }

& $venv_py schedule_tasks.py @args_list
