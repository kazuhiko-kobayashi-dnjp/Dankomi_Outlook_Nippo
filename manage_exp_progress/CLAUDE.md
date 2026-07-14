# CLAUDE.md — manage_exp_progress 引き継ぎメモ

このファイルは、別の開発環境(WSL上のClaude Code等)でこのプロジェクトの開発を
引き継ぐ際に、経緯や既知の注意点を素早く把握できるようにするための引き継ぎ資料。
アーキテクチャの詳細・運用手順は [README.md](README.md) を参照。ここでは
「直近のセッションで何をしたか」「繰り返しハマった罠」を中心にまとめる。

## プロジェクト概要(1行)
Excelで管理していた「露出制御業務進捗管理」表を、複数人で同時編集できるWebアプリ
(Flask + 単一JSONファイル + vanilla JS SPA)に置き換えたもの。詳細はREADME.md参照。

## 重要: このコードとデータの置き場所の関係
- **アプリコード**(`server.py`/`import_excel.py`/`export_excel.py`/`richtext_bridge.py`/
  `sharepoint_backup.py`/`public/index.html`/`README.md`)はGit管理下
  (`Dankomi_Outlook_Nippo`リポジトリ、PUBLIC)。
- **業務データ本体**(`app/data/tasks.json`、`app/data/backups/`、
  `app/data/excel_source_path.txt`、`app/data/template_source.xlsx`)と
  **認証キャッシュ**(`app/.msal_token_cache.json`)は、個人情報・機密情報のため
  **意図的にGit管理対象外**(`.gitignore`で除外)。これらは元PC側の
  `OneDrive - DENSO`同期フォルダにのみ存在する。
- そのため、このプロジェクトを新しい環境(WSL等)にcloneしただけでは`app/data/`が
  存在せず、アプリは起動できない。元PCの実データを参照する場合は、Windows側の
  OneDriveフォルダを `/mnt/c/...` 経由でsymlinkするなどして、**データは1箇所
  (元PCのOneDrive)に一元化したまま、コードだけ新環境で編集する**運用を推奨。
  データを新環境にコピーして分岐させると、どちらが最新か分からなくなり
  OneDrive競合コピー同様の事故が起きやすいので避けること。

## 本番環境について(絶対に踏み抜かないこと)
- 本番は元PC(社内LAN上、`10.41.55.204:3100`)で**常時稼働中**。
- `app/data/tasks.json`はOneDrive同期フォルダ上にあるため、**別プロセスから
  同時に書き込むとOneDriveが`tasks-<デバイス名>.json`という競合コピーを
  自動生成する**(2026-07-09に実際に発生済み。データ損失はなかったが要注意)。
- 新環境で動作確認する際は、本番と同じ`tasks.json`を指すsymlink経由で
  サーバーを起動する運用にする場合、**動作確認が終わったら必ずプロセスを停止する**。
  長時間放置しないこと。
- `server.py`は`debug=False`(自動リロード無し)。コードを変更したら
  **必ず手動でFlaskプロセスを再起動**しないと、古いAPIのままで
  「Unexpected token '<'」のようなエラーになる。

## 開発運用: 「WSLで編集 → Windowsで実行」(2026-07-14 確立)
Claude CodeがWSL上でしか動かないため、**コード編集はWSL、Flask本番はWindowsで実行**
という分担にした。サーバーを起動するのは常にWindowsの1プロセスだけなので、
tasks.jsonへの二重書き込み(OneDrive競合コピー事故)が構造的に発生しない。
- WSL側checkout: `/home/<user>/workspace/Dankomi_Outlook_Nippo/manage_exp_progress`
  (`app/data`と`app/.msal_token_cache.json`は`/mnt/f/...OneDrive`へのsymlink)
- Windows側checkout(本番): `F:\OneDrive_F\OneDrive - DENSO\2017\else\memo`
  配下の`manage_exp_progress\app`。両者とも同じGitHubリポジトリのclone。
- **反映は手動コピー禁止・git経由**: WSLで編集→commit→push、Windowsで`git pull`→再起動。
- この一連(停止→pull→再起動→稼働確認)をWSLからワンコマンドで回す
  **`deploy_restart.sh`**(git管理外のローカル運用ツール)を用意した。
  `./deploy_restart.sh` 実行だけでデバッグループが人間の介在なしで回る。
  ログは`app/server_out.log`(git管理外)へ出るので`tail -f`で追える。

### deploy_restart.sh 実装時に潰した3つの罠(再発防止)
1. **Windows側checkoutが全ファイルCRLF、WSL側がLF**。git上「全行変更」に見え
   (`--ignore-all-space`で差分ゼロ)、放置すると`git pull`がブロックされる。
   → pull直前に`manage_exp_progress`サブツリーだけ`git checkout -- <subtree>`で
   作業ツリーを捨ててから`merge --ff-only`。他パス(timetracker等の進行中作業)には触れない。
2. **バックグラウンド起動時stdoutがcp932になる** → 起動メッセージの絵文字🚀で
   `UnicodeEncodeError`が出てFlaskが即クラッシュ(手動起動のUTF-8端末では顕在化しない)。
   → `PYTHONUTF8=1`を渡す。
3. **WSL interopからWindowsプロセスを完全デタッチ起動する方法**:
   `cmd /c start /b`はcmd終了時に子が道連れで死ぬ、PowerShell`Start-Process
   -RedirectStandardOutput`はハンドルを掴んだままブロックして戻らない。
   → **VBScript(`WScript.Shell.Run cmd, 0, False`)**が唯一ブロックせず生き残る。
   `deploy_restart.sh`が`_launch_server.vbs`(git管理外)を生成してwscriptで叩いている。
- 注: 3100プロセスの特定/停止/起動確認はWSLから`powershell.exe Get-NetTCPConnection`
  /`Stop-Process`で行う。WindowsのpythonパスとWSL IPはスクリプト冒頭にハードコード
  してあるので、環境が変わったら`PY_WIN`と(portproxy運用に戻すなら)connectaddressを更新する。

## 直近のセッション(2026-07-14)で対応した内容
1. **Git管理化**: `Dankomi_Outlook_Nippo`リポジトリの`.gitignore`にホワイトリストを
   追加し、アプリコードのみ追跡開始(データ・認証情報は対象外のまま)。
2. **バージョン履歴で複数バージョンの差分が同一に見えるバグ修正**: `_backup_before_write()`
   が`shutil.copy2()`でmtimeをコピー元から引き継いでいたため、バックアップ間隔の
   スロットリング判定(5分)がファイル名ではなく古いmtimeを見てしまい、内容不変の
   バックアップが短時間に何個も作られていた。ファイル名からタイムスタンプを解析する
   `_parse_backup_stamp()`を追加し、さらに直前バックアップと内容が完全一致する場合は
   新規作成しない冪等チェックも追加。
3. **差分ビューで「書式のみの変更」を分離表示**: Excel再取込時、セル全体太字の
   キャプチャ修正(後述)により`<b>`タグだけが追加される行が大量発生し、「差分が
   あると言われるが見た目は同じで分からない」と混乱を招いていた。プレーンテキスト
   化後に完全一致するフィールドは`formatOnly:true`として分離し、実質的な内容変更
   (✏️)と書式のみの変更(📝、`<details>`で折りたたみ)を別々に表示するようにした。
4. **差分ビューの表記見直し**: 「復元後」という表現が復元前提の印象を与えるため、
   「📜 {バージョン日時} 時点」という表現に変更。色も現在=青系・過去=グレーに
   変更し、「現在の方が正しくて復元後が理想」というバイアスを排除した。
5. **ローカルバックアップの保持世代数を60→500に拡大**: 1世代あたり約700KB
   (tasks.json全体のスナップショット)なので500世代でも約350MBで問題なし。
   5分間隔のスロットリングにより「保存回数」ではなく「実際に5分以上間隔が
   空いた回数」だけ世代を消費する仕組みも要説明(ユーザーが誤解しやすい点)。

## それより前のセッションで対応した主な内容(詳細は各ファイルのコメント・README参照)
- Excelインポート時にセル全体の太字/下線/文字色が反映されない問題の修正
  (`richtext_bridge.py`の`celltext_to_html()`にfont引数追加)。
- Excelエクスポートで「修復が必要」エラーが出る問題の修正(openpyxlのVML/drawing
  再シリアライズ崩れを、テンプレートの元バイト列で復元する`_restore_legacy_drawing_parts()`)。
- 保存すると絞り込み条件がリセットされる問題の修正(`populateFilters()`)。
- テキスト欄への画像貼り付け・挿入機能(data URL埋め込み方式)。
- 担当者別の行色付け(元Excelのdxf条件付き書式から色コードを抽出)。
- バージョン履歴GUI(一覧・差分・復元)、SharePointへの自動バックアップ
  (`sharepoint_backup.py`、複数世代保持)。

## 繰り返し発生した典型的なハマりどころ
- **openpyxlのround-trip問題**: VMLコメント・シート直貼り画像・チャートシートを
  含むExcelを「読み込み→保存」すると、値を変えていない部分でも壊れることがある
  (`export_excel.py`参照)。同様の症状(Excelの「修復されたレコード」ログ)が
  出たら、まず変更前後でどのXMLパーツが不一致か比較する。
- **`shutil.copy2`はmtimeをコピー元から引き継ぐ**ため、「ファイルのmtime=作成時刻」
  という前提のロジックは壊れることがある(今回のバックアップ間隔バグの根本原因)。
  ファイル名に埋め込んだタイムスタンプなど、自分で制御できる情報を基準にする方が安全。
- **OneDrive同期フォルダ上でのローカルテストサーバー起動は要注意**(競合コピー発生済み)。
- **既存CSSのクラス名を新しいUI要素に流用すると、意図せず`color`等が漏れて継承される**
  ことがある(`.badge.xxx`のような親クラス限定か、`.xxx`単体セレクタで汎用的に
  効いてしまうか要確認)。

## 「ローカル」と「SharePoint」の違い(混同しやすいので明記)
- 元PCのプロジェクトフォルダ自体が`OneDrive - DENSO`という個人OneDrive同期フォルダの
  中にあるため、`app/data/backups/`(ローカルバックアップ)も**結果的にはOneDrive経由で
  クラウドにも同期される**。ただし、あくまで**個人のOneDrive**であり、ファイルシステムに
  直接読み書きするので高速・オフラインでも動作する。
- 一方`sharepoint_backup.py`が行うアップロードは、**チーム共有のSharePointサイト**
  (`globaldenso.sharepoint.com/teams/TMS_o365_jp103832`)への、Graph API経由の
  明示的なコピー。個人OneDriveとは別の保管場所であり、チームメンバーも参照できる点、
  個人OneDriveアカウント側に問題があっても独立して残る点が価値。
