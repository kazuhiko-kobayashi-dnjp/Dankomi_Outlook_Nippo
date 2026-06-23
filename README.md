# outlook_to_timetracker

Outlook カレンダーの予定を TimeTrackerNX に自動登録するツール。

## 動作環境

- Windows 10/11
- Python 3.10+
- Outlook（デスクトップ版）がインストール済みであること
- TimeTrackerNX アカウント・API キーがあること

## セットアップ（初回のみ）

### 1. Python 仮想環境を作成

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
# pywin32 の後処理（必須）
python .venv\Scripts\pywin32_postinstall.py -install
```

### 2. API キーファイルを作成

`tt_apikey.txt` を作成し、1行目に API キーを記入する。  
（TimeTrackerNX のユーザー設定画面で発行）

```
e03bc76e-xxxx-xxxx-xxxx-xxxxxxxxxxxx
```

### 3. `timetracker_config.json` を自分の環境に合わせて編集

- `project_rules` — 予定タイトルのキーワード → プロジェクトコード対応
- `workitem_map` — プロジェクト × 作業種別 → workItemId 対応

## 日常の使い方

```powershell
# 本日分を登録（Outlook ソース・デフォルト）
.\tt.ps1

# 日付指定（YYMMDD 形式）
.\tt.ps1 260624

# 内容確認のみ（dry-run）
.\tt.ps1 260624 -Dry

# 日報ファイルをソースにする場合
.\tt.ps1 260624 -Diary
```

詳細は `timetracker_manual.html` を参照。

## ファイル構成

| ファイル | 役割 |
|---|---|
| `timetracker_register.py` | メインスクリプト |
| `timetracker_config.json` | プロジェクト・WorkItem 設定 |
| `tt.ps1` | ワンコマンド起動スクリプト |
| `timetracker_manual.html` | 使い方マニュアル |
| `output/tt_delete_date.py` | 指定日の登録済み実績を全削除するユーティリティ |
| `requirements.txt` | Python 依存パッケージ |
| `tt_apikey.txt` | **自分で作成**（Git 管理外・要秘密保持） |

## Git 管理外のファイル

- `tt_apikey.txt` — API キー（`.gitignore` で除外済み）
- `.venv/` — 仮想環境
- 日報ファイル・個人作業ログ
