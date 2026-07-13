# 露出制御業務進捗管理 (manage_exp_progress)

`◎250217_露出制御業務進捗及び報告.xlsx` (Sheet1) をWeb化した進捗管理アプリ。
参考: [manage_swreq_expisp](https://github.com/kazuhiko-kobayashi-dnjp/manage_swreq_expisp) (BEV要件織り込み管理) と同じ構成。

## 特長
- フィルタ: プロジェクト / 担当者 / ステータス(◎○●▲済完) / キーワード検索
- 案件クリックで詳細をモーダル表示・編集（Excelのように行高さを都度調整する必要なし）
- 期限は超過/2日以内/7日以内/それ以降で自動色分け
- Excelへのエクスポート（フィルタ後のみ、または全件）

## 起動方法

このPCにはNode.js/npmが未インストールのため、参考アプリ(manage_swreq_expisp、Express構成)と
同等のAPIを **Flask (Python)** で実装している。事前に `pip install flask openpyxl` が必要（既に導入済みの場合は不要）。

```powershell
cd app
python server.py
```

ブラウザで `http://localhost:3100` を開く。

## ファイル構成

```
app/
  server.py            # Flask APIサーバー
  import_excel.py       # 元Excel(Sheet1) → data/tasks.json 変換スクリプト
  export_excel.py        # data/tasks.json → Excel 書き出し (build_workbook関数をserver.pyから直接呼び出し)
  data/
    tasks.json            # タスクデータ（449件、初回インポート済み）
  public/
    index.html            # フロントエンド(SPA)
```

## データ再取込

元Excelが更新された場合、以下で再取込できる（**注意: 現在のdata/tasks.jsonのWeb上の編集内容は上書きされる**）。

```powershell
python app/import_excel.py "<元Excelパスxlsx>" app/data/tasks.json
```

## データ永続化
- 編集・追加・削除は `data/tasks.json` にリアルタイム書き込み。
- Excelへの書き戻しはエクスポートAPI経由（`export_excel.py`）で別ファイルとして生成。元Excelへの直接上書きは行わない。

## バックアップ・バージョン管理

`manage_exp_progress` フォルダはこのリポジトリ（`memo`）と同じ場所、つまり
**小林個人のOneDrive（`OneDrive - DENSO`）配下**にある。そのため `data/tasks.json` も
含めて自動的にクラウドへバックアップされ、複数PC（本番機 10.41.55.204 含む）間で同期される。
エクセル版の時と同様、原本は既に個人OneDrive上にある（ローカルPC限定ではない）。

保存内容を元に戻したい場合、以下の2段構えで対応できる。

1. **OneDriveのバージョン履歴（第一の手段）**
   `data/tasks.json` を右クリック → 「以前のバージョン」（またはOneDrive Web上の
   「バージョン履歴」）から、過去の任意の保存時点に戻せる。DENSOの法人OneDrive
   （SharePoint Online基盤）は既定で多数の世代を自動保持するため、追加設定は不要。

2. **アプリ内蔵の自動バックアップ（第二の手段・ローカル/OneDrive両対応の保険）**
   `server.py` の保存処理（`write_data()`）は、上書き前に `data/backups/tasks_YYYYMMDD_HHMMSS.json`
   としてスナップショットを自動保存する（直近500世代を保持、5分以内の連続保存はスキップして
   世代が乱立しないようにしている）。1世代あたりのファイルサイズはtasks.json全体
   （現状 約700KB）なので、500世代でも約350MB程度でありローカルディスク容量としては
   問題にならない。複数人が短時間に連続保存しても、消費される世代数は「保存回数」ではなく
   「実際に5分以上間隔が空いた回数」で決まる（5分以内の保存はまとめて1世代分として扱われる）。
   OneDriveの同期が一時的に止まっていても機能する。
   戻したい場合は、該当のバックアップファイルを `data/tasks.json` に上書きコピーすればよい
   （アプリ内の「🕐 バージョン履歴」からも一覧・差分確認・復元ができる）。

3. **チームSharePointへの自動バックアップ（第三の手段・組織管理下の保管先）**
   保存のたびに `data/tasks.json` の内容を、チームSharePoint
   （`https://globaldenso.sharepoint.com/teams/TMS_o365_jp103832/Shared Documents/dankomi/`）
   へ自動アップロードする（[sharepoint_backup.py](app/sharepoint_backup.py)）。2種類の形で保存する:
   - `dankomi/tasks_backup.json` … 常に最新版に上書き（すぐ参照したい時用）
   - `dankomi/backups/tasks_YYYYMMDD_HHMMSS.json` … 世代管理用。**直近100世代を明示的に保持**
     （SharePoint自体のバージョン履歴の保持設定・上限に依存しない。5分以内の連続保存は
     新規世代作成をスキップして乱立を防止。101世代目以降は古い方から自動削除）

   アップロードは非同期（別スレッド）で行われ、失敗しても保存処理自体はブロックされない。

   **初回セットアップ（1回だけ、ブラウザでの手動操作が必要）:**
   ```powershell
   cd app
   pip install msal requests    # 未導入の場合のみ
   python sharepoint_backup.py --auth
   ```
   表示された認証URLをブラウザで開き、表示されたコードを入力してサインインする。
   以後はトークンがキャッシュされ（`app/.msal_token_cache.json`、Git管理対象外）、
   自動更新されるため再認証は基本不要（動作確認: `python sharepoint_backup.py --test`）。

### アプリのバージョン履歴機能について

アプリ内の「🕐 バージョン履歴」ボタンから確認・復元できるのは、**ローカルの
`data/backups/`（上記2番）のみ**。SharePoint側の世代（上記3番）は現時点でアプリ内から
一覧・復元する機能はなく、必要であればSharePointのWeb画面から直接参照すること。

### 「保存」とみなされるタイミングについて

バージョン履歴・バックアップでいう「保存」は、`server.py`の`write_data()`が呼ばれた
タイミング全てを指す。具体的には: Web画面でのタスク作成・編集・削除、および
「🔄 Excelから再取込」の実行。**元Excelの内容を再取込した場合も1つの保存イベントとして
記録される**（差分ビューでも通常の編集と同じ扱いで表示される。取込元が何であったかの
ラベル付けは現時点では無い）。

### 何か起きた際の復旧範囲について

上記3つの手段はいずれも**データ（`data/tasks.json`の中身）のバックアップ**であり、
アプリ本体のコード（`server.py`・`public/index.html`等）はカバーしない。

コード自体は、`Dankomi_Outlook_Nippo`リポジトリ（`wr.ps1`と同じリポジトリ）にGit管理下で
追跡されている（`server.py`/`import_excel.py`/`export_excel.py`/`richtext_bridge.py`/
`sharepoint_backup.py`/`public/index.html`/`README.md`）。`app/data/`（業務データ本体・
ローカルバックアップ・元Excelパス設定）と`app/.msal_token_cache.json`（SharePoint認証
キャッシュ）は個人情報・認証情報のため引き続きGit管理対象外（`.gitignore`で除外）。
これらはOneDrive同期のみでカバーされ、OneDriveのバージョン履歴・ごみ箱機能で復旧する。

### 既知の注意点（複数PCから同時に書き込むと競合コピーが発生する）

`data/tasks.json` は同一ファイルとして本番機（10.41.55.204）と開発用PCの両方から
OneDrive経由でアクセスされる。**2台以上のPCで同時にFlaskサーバー（`python server.py`）を
起動して書き込みが重なると、OneDriveが `tasks-<デバイス名>.json` という競合コピーを
自動生成してしまう**（2026-07-09に実際に発生し、`data/_archive/`に退避済み）。
動作確認のためにサーバーを一時的に起動する場合は、確認が終わったらすぐに停止すること。
