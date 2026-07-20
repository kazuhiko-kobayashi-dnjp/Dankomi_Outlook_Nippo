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
#   .\wr.ps1 -Summarize -TaskWeb       # 業務進捗表をExcelの代わりにmanage_exp_progressのtasks.jsonから読む
#
# ※すでに一度実行済みで、後からAI要約だけ付けたいとき(Teamsログの再取得は遷いためスキップ):
#   .\wr.ps1                           # 1回目: 事前収集(何も付けなくても自動でキャッシュ保存される)
#   .\wr.ps1 -FromCache -Summarize     # 2回目以降: キャッシュを使って再収集なしでAI要約だけ実行(高速)
#   .\wr.ps1 -FromCache -Prompt        # 同様にプロンプトのみもキャッシュから高速生成可能
#
# 月曜日に先週分を取得したいとき:
#   .\wr.ps1 -Days 7 -Summarize        # 月曜朝に実行 → 先週月〜日をカバー
#
# 初回セットアップ（任意）:
#   tools\.user_config.example.json を tools\.user_config.json にコピーし、
#   自分の担当者名(task_owner)・業務進捗Excelパス(task_excel_path)を記入すると、
#   -TaskOwner / -TaskExcel を毎回指定しなくてもよくなる（このファイルはgit管理対象外）。
#   もしくは .\wr.ps1 -TaskOwner "自分の名前" -SaveTaskOwner -Prompt で担当者名だけ保存も可能。
#   Excelの代わりに manage_exp_progress (Web化システム) を使う場合は -TaskWeb を付ける。
#
# Output: output\weekly_report_YYYYMMDD_HHMM.md
#         output\summary_prompt_YYYYMMDD_HHMM.txt  (-Prompt 時)
#         output\.records_cache.json  (毎回自動上書き保存、-FromCacheで次回再利用)

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
    [switch] $FromCache,            # 直前の収集結果(output\.records_cache.json)を再利用し、Graph API再収集をスキップ
    [string] $NippoDir    = "",     # 日報ディレクトリ（省略時: 自動検出）
    [string] $TaskExcel   = "",     # 業務進捗Excelパス（省略時: tools\.user_config.json の task_excel_path）
    [switch] $TaskWeb,               # 業務進捗表の取得元をExcelの代わりにmanage_exp_progressのtasks.jsonにする
    [string] $TaskJson    = "",     # manage_exp_progressのtasks.jsonパス（省略時: tools\.user_config.json の task_json_path、それも無ければ既定の相対パス）
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
if ($FromCache)      { $a += '--from-cache' }
if ($NippoDir)       { $a += '--nippo-dir'; $a += $NippoDir }
if ($TaskExcel)      { $a += '--task-excel'; $a += $TaskExcel }
if ($TaskWeb)        { $a += '--task-source'; $a += 'web' }
if ($TaskJson)       { $a += '--task-json'; $a += $TaskJson }
if ($TaskOwner)      { $a += '--task-owner'; $a += $TaskOwner }
if ($SaveTaskOwner)  { $a += '--save-task-owner' }
# TaskExcel/TaskOwner/TaskJson 省略時のデフォルトは tools\teams_weekly_report.py 側が
# tools\.user_config.json（git管理対象外・個人設定）から読み込む。
# ファイルが無ければ何も自動設定されず、業務進捗表セクションは単に省略される
# （第三者配布時も安全に動作する）。設定例は tools\.user_config.example.json を参照。

Write-Host "Weekly Report (last $Days days)" -ForegroundColor Cyan
& $py @a
