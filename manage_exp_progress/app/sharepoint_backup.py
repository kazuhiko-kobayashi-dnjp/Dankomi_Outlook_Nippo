#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
data/tasks.json をチームSharePointへバックアップするモジュール。

server.py の write_data() から保存のたびに呼び出され、SharePoint上に2種類の形で保存する。
  1) dankomi/tasks_backup.json          … 常に最新版に上書き(すぐ参照したい時用)
  2) dankomi/backups/tasks_YYYYMMDD_HHMMSS.json … 世代管理用。直近 BACKUP_KEEP 件を明示的に保持
     (SharePoint自体のバージョン履歴の保持設定に依存せず、こちら側で世代数を担保する)

初回セットアップ(認証。ブラウザでの手動操作が必要):
    python sharepoint_backup.py --auth

単体テスト(現在の data/tasks.json を1回だけアップロード):
    python sharepoint_backup.py --test

必要パッケージ: pip install msal requests
"""
import sys
import os
import subprocess
import threading
import time
from pathlib import Path

# ── プロキシ自動検出 (tools/teams_weekly_report.py と同じロジックを踏襲) ──
def _auto_detect_proxy():
    if os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY"):
        return
    try:
        ps_cmd = (
            '$p=[System.Net.WebRequest]::GetSystemWebProxy();'
            '$u=[Uri]"https://login.microsoftonline.com";'
            '$r=$p.GetProxy($u);'
            'if($r.Authority -ne $u.Authority){Write-Output $r.AbsoluteUri}'
        )
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", ps_cmd],
            capture_output=True, text=True, timeout=10,
        )
        proxy = r.stdout.strip().rstrip("/")
        if proxy:
            os.environ["HTTP_PROXY"] = proxy
            os.environ["HTTPS_PROXY"] = proxy
            return
    except Exception:
        pass
    try:
        r = subprocess.run(["netsh", "winhttp", "show", "proxy"],
                           capture_output=True, text=True, timeout=5)
        for line in r.stdout.splitlines():
            if "Proxy Server" in line and ":" in line:
                proxy = line.split(":", 1)[1].strip()
                if proxy and proxy != "(null)":
                    os.environ["HTTP_PROXY"] = proxy
                    os.environ["HTTPS_PROXY"] = proxy
                    return
    except Exception:
        pass

_auto_detect_proxy()

try:
    import msal
    import requests
except ImportError:
    msal = None
    requests = None

AUTHORITY = "https://login.microsoftonline.com/69405920-b673-4f7c-8845-e124e9d08af2"
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"   # Microsoft Office (1st party)
SCOPES    = ["https://graph.microsoft.com/.default"]
GRAPH     = "https://graph.microsoft.com/v1.0"

# 対象SharePoint (「BEV要件織り込み管理」等が置かれているのと同じチームサイト直下)
SITE_HOSTNAME   = "globaldenso.sharepoint.com"
SITE_PATH       = "/teams/TMS_o365_jp103832"
TARGET_FOLDER   = "dankomi"
TARGET_FILENAME = "tasks_backup.json"          # 常に最新を上書き(すぐ参照したい時用)
BACKUP_SUBFOLDER = f"{TARGET_FOLDER}/backups"  # 世代管理用(日付付きファイルを蝹積)
BACKUP_KEEP = 100                              # 直近この件数のみ保持(SharePoint自体のバージョン履歴設定に依存しない)
BACKUP_MIN_INTERVAL_SEC = 300                  # 前回世代バックアップからこの秒数未満なら新規作成をスキップ

_CACHE_FILE = Path(__file__).with_name(".msal_token_cache.json")


def _load_cache():
    cache = msal.SerializableTokenCache()
    if _CACHE_FILE.exists():
        cache.deserialize(_CACHE_FILE.read_text(encoding="utf-8"))
    return cache


def _save_cache(cache):
    if cache.has_state_changed:
        _CACHE_FILE.write_text(cache.serialize(), encoding="utf-8")


def get_token(interactive: bool = False):
    """キャッシュから無言取得。取得できず interactive=True の場合のみ device code 認証を行う。
    server.py からの自動呼び出し時は interactive=False とし、未認証ならNoneを返して
    静かにスキップする(保存処理そのものは妨げない)。"""
    if msal is None:
        return None
    cache = _load_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _save_cache(cache)
            return result["access_token"]
    if not interactive:
        return None

    flow = app.initiate_device_flow(scopes=SCOPES)
    print("=" * 55, flush=True)
    print(f"  認証URL : {flow['verification_uri']}", flush=True)
    print(f"  コード  : {flow['user_code']}", flush=True)
    print("=" * 55, flush=True)
    print("ブラウザで上記URLを開きコードを入力してください...", flush=True)
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        print(f"[ERROR] 認証失敗: {result.get('error_description', '')[:200]}", flush=True)
        return None
    _save_cache(cache)
    print("認証成功。トークンをキャッシュしました（以後は自動更新されます）。", flush=True)
    return result["access_token"]


_site_drive_cache = {}  # プロセス内キャッシュ(サイトID/ドライブID解決結果の使い回し)


def _resolve_drive_id(token: str):
    if "drive_id" in _site_drive_cache:
        return _site_drive_cache["drive_id"]
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(f"{GRAPH}/sites/{SITE_HOSTNAME}:{SITE_PATH}", headers=headers, timeout=15)
    if r.status_code != 200:
        print(f"[WARN] SharePointサイト解決失敗: {r.status_code} {r.text[:200]}", flush=True)
        return None
    site_id = r.json()["id"]
    r2 = requests.get(f"{GRAPH}/sites/{site_id}/drive", headers=headers, timeout=15)
    if r2.status_code != 200:
        print(f"[WARN] ドキュメントライブラリ解決失敗: {r2.status_code} {r2.text[:200]}", flush=True)
        return None
    drive_id = r2.json()["id"]
    _site_drive_cache["drive_id"] = drive_id
    return drive_id


def upload_backup(json_bytes: bytes, interactive: bool = False) -> bool:
    """data/tasks.json の内容をSharePointへバックアップする。
    1) dankomi/tasks_backup.json に常に上書き(最新版をすぐ参照したい時用)
    2) dankomi/backups/tasks_YYYYMMDD_HHMMSS.json として世代付きでも保存し、
       直近BACKUP_KEEP件のみを明示的に保持する(SharePoint自体のバージョン履歴の
       保持設定に依存せず、こちら側で世代数を担保するため)。
    失敗しても例外を送出せず False を返すだけ(呼び出し元の保存処理を妨げない)。"""
    if requests is None or msal is None:
        return False
    token = get_token(interactive=interactive)
    if not token:
        return False
    drive_id = _resolve_drive_id(token)
    if not drive_id:
        return False
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    ok_latest = False
    try:
        url_latest = f"{GRAPH}/drives/{drive_id}/root:/{TARGET_FOLDER}/{TARGET_FILENAME}:/content"
        r = requests.put(url_latest, headers=headers, data=json_bytes, timeout=30)
        ok_latest = r.status_code in (200, 201)
        if not ok_latest:
            print(f"[WARN] SharePointバックアップ(最新)失敗: {r.status_code} {r.text[:200]}", flush=True)
    except Exception as e:
        print(f"[WARN] SharePointバックアップ(最新)例外: {e}", flush=True)

    global _last_backup_time
    now = time.time()
    if now - _last_backup_time >= BACKUP_MIN_INTERVAL_SEC:
        try:
            stamp = time.strftime("%Y%m%d_%H%M%S")
            url_dated = f"{GRAPH}/drives/{drive_id}/root:/{BACKUP_SUBFOLDER}/tasks_{stamp}.json:/content"
            r2 = requests.put(url_dated, headers=headers, data=json_bytes, timeout=30)
            if r2.status_code in (200, 201):
                _last_backup_time = now
                _prune_backups(token, drive_id)
            else:
                print(f"[WARN] SharePointバックアップ(世代)失敗: {r2.status_code} {r2.text[:200]}", flush=True)
        except Exception as e:
            print(f"[WARN] SharePointバックアップ(世代)例外: {e}", flush=True)

    return ok_latest


_last_backup_time = 0.0  # プロセス内キャッシュ(世代バックアップの連続作成を抱制)


def _list_dated_backups(token: str, drive_id: str):
    headers = {"Authorization": f"Bearer {token}"}
    r = requests.get(
        f"{GRAPH}/drives/{drive_id}/root:/{BACKUP_SUBFOLDER}:/children",
        headers=headers, params={"$select": "id,name", "$top": "999"}, timeout=15,
    )
    if r.status_code == 404:
        return []  # フォルダ未作成(初回)
    if r.status_code != 200:
        print(f"[WARN] SharePointバックアップ一覧取得失敗: {r.status_code} {r.text[:200]}", flush=True)
        return []
    items = [it for it in r.json().get("value", []) if it.get("name", "").startswith("tasks_")]
    return sorted(items, key=lambda x: x["name"])


def _prune_backups(token: str, drive_id: str):
    items = _list_dated_backups(token, drive_id)
    if len(items) <= BACKUP_KEEP:
        return
    headers = {"Authorization": f"Bearer {token}"}
    for old in items[:-BACKUP_KEEP]:
        try:
            requests.delete(f"{GRAPH}/drives/{drive_id}/items/{old['id']}", headers=headers, timeout=15)
        except Exception as e:
            print(f"[WARN] 古いSharePointバックアップの削除に失敗: {e}", flush=True)


def upload_backup_async(json_bytes: bytes):
    """バックグラウンドスレッドで非同期アップロードする(保存APIのレスポンスを遅延させない)。"""
    def _run():
        upload_backup(json_bytes, interactive=False)
    threading.Thread(target=_run, daemon=True).start()


if __name__ == "__main__":
    if "--auth" in sys.argv:
        token = get_token(interactive=True)
        sys.exit(0 if token else 1)
    if "--test" in sys.argv:
        data_file = Path(__file__).parent / "data" / "tasks.json"
        content = data_file.read_bytes()
        ok = upload_backup(content, interactive=False)
        print("アップロード成功" if ok else "アップロード失敗（先に --auth を実行してください）", flush=True)
        sys.exit(0 if ok else 1)
    print("Usage: python sharepoint_backup.py --auth | --test", flush=True)
