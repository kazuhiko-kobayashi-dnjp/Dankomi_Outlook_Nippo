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

param(
    [string]$Week    = '',
    [switch]$Execute,
    [switch]$Clear,
    [switch]$Excel
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

& $venv_py schedule_tasks.py @args_list