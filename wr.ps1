# wr.ps1 -- Teams chat weekly report generator
#
# Usage:
#   .\wr.ps1                           # 直近 7 日, チャットのみ（最速）
#   .\wr.ps1 -Days 14                  # 直近 14 日
#   .\wr.ps1 -ListTeams                # 参加 Teams + チャンネル一覧を表示して終了
#   .\wr.ps1 -Channels                 # Teams チャンネルも含む（全 Teams スキャン）
#   .\wr.ps1 -Channels -TeamsAllow "BEV,FCM,DCAP"   # 指定 Teams のみスキャン（高速）
#   .\wr.ps1 -Channels -TeamsAllow "BEV" -SaveTeamsAllow  # フィルタ保存（次回以降省略可）
#   .\wr.ps1 -Summarize                # AI 要約を自動実行（要 API キー保存済み）
#   .\wr.ps1 -ApiKey "ghp_xxx" -SaveKey -Summarize  # 初回: キー保存 + 要約
#   .\wr.ps1 -Model gpt-4o -Summarize  # モデル指定（省略時: gpt-4o-mini）
#   .\wr.ps1 -Prompt                   # プロンプトのみ出力（API 不要、手貼り用）
#   .\wr.ps1 -Auth                     # トークン切れ時の再認証
#
# 月曜日に先週分を取得したいとき:
#   .\wr.ps1 -Days 7 -Summarize        # 月曜朝に実行 → 先週月〜日をカバー
#
# 初回セットアップ（任意）:
#   tools\.user_config.example.json を tools\.user_config.json にコピーし、
#   自分の担当者名(task_owner)・業務進捗Excelパス(task_excel_path)を記入すると、
#   -TaskOwner / -TaskExcel を毎回指定しなくてもよくなる（このファイルはgit管理対象外）。
#   もしくは .\wr.ps1 -TaskOwner "自分の名前" -SaveTaskOwner -Prompt で担当者名だけ保存も可能。
#
# Output: output\weekly_report_YYYYMMDD_HHMM.md
#         output\summary_prompt_YYYYMMDD_HHMM.txt  (-Prompt 時)

param(
    [int]    $Days          = 7,
    [switch] $Channels,             # Teams チャンネルを含む
    [switch] $ListTeams,            # 参加 Teams 一覧を表示して終了
    [string] $TeamsAllow   = "",    # スキャン対象 Teams 名（カンマ区切り部分一致）
    [switch] $SaveTeamsAllow,       # TeamsAllow を tools\.teams_allow に保存
    [switch] $Prompt,               # プロンプトのみ出力（API 不要）
    [switch] $Summarize,            # AI 要約を自動実行
    [string] $ApiKey       = "",    # GitHub Models / OpenAI API キー（初回のみ指定）
    [switch] $SaveKey,              # ApiKey を tools\.gh_models_token に保存
    [string] $Model        = "",    # モデル名（省略時: gpt-4o-mini）
    [switch] $Auth,                 # 再認証（Device Code Flow）
    [string] $NippoDir    = "",     # 日報ディレクトリ（省略時: 自動検出）
    [string] $TaskExcel   = "",     # 業務進捗Excelパス（省略時: tools\.user_config.json の task_excel_path）
    [string] $TaskOwner   = "",     # 業務進捗表の担当者フィルタ（省略時: tools\.user_config.json の task_owner）
    [switch] $SaveTaskOwner          # TaskOwner を tools\.user_config.json に保存
)

Set-Location $PSScriptRoot

$py = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $py)) {
    Write-Error '.venv not found. Run: python -m venv .venv'
    exit 1
}

$a = @('tools\teams_weekly_report.py', '--days', $Days)
if ($ListTeams)      { $a += '--list-teams' }
if (-not $Channels -and -not $ListTeams) { $a += '--no-channels' }
if ($TeamsAllow)     { $a += '--teams-allow'; $a += $TeamsAllow }
if ($SaveTeamsAllow) { $a += '--save-teams-allow' }
if ($Prompt)         { $a += '--prompt-only' }
if ($Summarize)      { $a += '--summarize' }
if ($ApiKey)         { $a += '--api-key'; $a += $ApiKey }
if ($SaveKey)        { $a += '--save-key' }
if ($Model)          { $a += '--model'; $a += $Model }
if ($Auth)           { $a += '--device-code' }
if ($NippoDir)       { $a += '--nippo-dir'; $a += $NippoDir }
if ($TaskExcel)      { $a += '--task-excel'; $a += $TaskExcel }
if ($TaskOwner)      { $a += '--task-owner'; $a += $TaskOwner }
if ($SaveTaskOwner)  { $a += '--save-task-owner' }
# TaskExcel/TaskOwner 省略時のデフォルトは tools\teams_weekly_report.py 側が
# tools\.user_config.json（git管理対象外・個人設定）から読み込む。
# ファイルが無ければ何も自動設定されず、業務進捗表セクションは単に省略される
# （第三者配布時も安全に動作する）。設定例は tools\.user_config.example.json を参照。

Write-Host "Weekly Report (last $Days days)" -ForegroundColor Cyan
& $py @a
