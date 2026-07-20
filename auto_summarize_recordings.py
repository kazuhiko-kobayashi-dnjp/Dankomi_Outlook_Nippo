# -*- coding: utf-8 -*-
"""
OneDrive/SharePoint フォルダURL -> MP4列挙 -> transcript(VTT)取得 -> Copilot(API)要約 -> 保存
- transcript API: SharePoint REST v2.1 /_api/v2.1/drives/{driveId}/items/{itemId}/media/transcripts
  (temporaryDownloadUrl で VTT をダウンロード)
- Copilot API: Graph beta /copilot/conversations -> /copilot/conversations/{id}
"""

import os
import re
import json
import csv
import time
import base64
import argparse
import subprocess
from urllib.parse import urlparse, unquote

import requests
import msal


def _auto_detect_proxy():
    """
    .NET GetSystemWebProxy() を使って実際のプロキシを自動検出。
    WinHTTP/WinINET + PAC 対応のため、netsh より正確。
    """
    if os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY"):
        print(f"[INFO] 既存のProxy環境変数を使用: {os.environ.get('HTTPS_PROXY')}", flush=True)
        return
    try:
        # PowerShell 経由で .NET のシステムプロキシを問い合わせ
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
            print(f"[INFO] Proxy自動検出(Win): {proxy}", flush=True)
            return
    except Exception as e:
        print(f"[WARN] Proxy自動検出失敗: {e}", flush=True)

    # フォールバック
    proxy_fallback = "http://in-proxy-o.denso.co.jp:8080"
    os.environ["HTTP_PROXY"] = proxy_fallback
    os.environ["HTTPS_PROXY"] = proxy_fallback
    print(f"[INFO] Proxy設定(Fallback): {proxy_fallback}", flush=True)


_auto_detect_proxy()


GRAPH = "https://graph.microsoft.com"
GRAPH_BETA = "https://graph.microsoft.com/beta"


def b64url_no_pad(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).decode("utf-8").rstrip("=")


def encode_share_url(url: str) -> str:
    # Graph /shares/{shareIdOrEncodedSharingUrl} 用: "u!{base64url(url)}"
    return "u!" + b64url_no_pad(url.encode("utf-8"))


def ensure_dir(path: str):
    os.makedirs(path, exist_ok=True)


def request_with_retry(method, url, headers=None, params=None, json_body=None, stream=False, timeout=60):
    # 429/5xx を簡易リトライ（指数バックオフ）
    for attempt in range(8):
        resp = requests.request(
            method, url,
            headers=headers,
            params=params,
            json=json_body,
            stream=stream,
            timeout=timeout,
        )
        if resp.status_code in (429, 500, 502, 503, 504):
            wait = min(60, (2 ** attempt))  # 上限60秒
            ra = resp.headers.get("Retry-After")
            if ra and ra.isdigit():
                wait = min(120, int(ra))
            time.sleep(wait)
            continue
        return resp
    return resp


def get_token(scope=None):
    # 「実績」重視：常に .default スコープを使用してキャッシュを最大限活用する。
    # 特定のスコープを要求するとブラウザ認証が走るため、リソースベースの .default に統一。
    if not scope:
        scopes = ["https://graph.microsoft.com/.default"]
    elif isinstance(scope, str):
        if "graph.microsoft.com" in scope:
            scopes = ["https://graph.microsoft.com/.default"]
        elif "sharepoint.com" in scope:
            u = urlparse(scope.replace("/.default", ""))
            scopes = [f"{u.scheme}://{u.netloc}/.default"]
        else:
            scopes = [scope]
    else:
        # リストで渡された場合も Graph 向けであれば .default に丸める
        # これによりブラウザ認証を回避し、キャッシュされたトークンを使い回す
        is_graph = any("graph.microsoft.com" in str(s) for s in scope)
        if is_graph:
            scopes = ["https://graph.microsoft.com/.default"]
        else:
            scopes = scope

    print(f"[DEBUG] get_token(scopes={scopes}) called", flush=True)
    # 過去実績の王者 (Microsoft Office) を使用。
    tenant_id = os.environ.get("TENANT_ID", "69405920-b673-4f7c-8845-e124e9d08af2").strip()
    client_id = os.environ.get("CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c").strip()
    client_secret = os.environ.get("CLIENT_SECRET", "").strip()
    token_cache_path = os.path.join("tools", ".msal_token_cache.json")

    if not tenant_id or not client_id:
        raise SystemExit("TENANT_ID と CLIENT_ID を環境変数に設定してください。")

    authority = f"https://login.microsoftonline.com/{tenant_id}"

    # トークンキャッシュの設定
    cache = msal.SerializableTokenCache()
    if os.path.exists(token_cache_path):
        with open(token_cache_path, "r") as f:
            cache.deserialize(f.read())

    # キャッシュを保存する関数
    def save_cache():
        if cache.has_state_changed:
            with open(token_cache_path, "w") as f:
                f.write(cache.serialize())

    # 可能ならアプリ権限（client credentials）、なければ device code（委任）
    if client_secret:
        app = msal.ConfidentialClientApplication(
            client_id=client_id,
            authority=authority,
            client_credential=client_secret,
            token_cache=cache
        )
        # まずキャッシュから取得
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(scopes, account=accounts[0])
            if result:
                return result["access_token"]
        
        result = app.acquire_token_for_client(scopes=scopes)
        if "access_token" not in result:
            raise SystemExit(f"トークン取得失敗(client credentials): {result}")
        save_cache()
        return result["access_token"]
    # 可能ならブラウザでインタラクティブ認証
    app = msal.PublicClientApplication(
        client_id=client_id, 
        authority=authority,
        token_cache=cache
    )
    
    # まずキャッシュから
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        if result:
            return result["access_token"]

    print(f"ブラウザを開いて認証してください ({scopes})...", flush=True)
    result = app.acquire_token_interactive(
        scopes=scopes
    )
    if "access_token" in result:
        save_cache()
        return result["access_token"]
    
    raise SystemExit(f"トークン取得失敗: {result}")


def graph_get(access_token: str, path: str, params=None):
    url = GRAPH + path
    headers = {"Authorization": f"Bearer {access_token}"}
    resp = request_with_retry("GET", url, headers=headers, params=params)
    if resp.status_code >= 400:
        raise RuntimeError(f"Graph GET failed {resp.status_code}: {url}\n{resp.text}")
    return resp.json()


def graph_post_beta(access_token: str, path: str, body: dict):
    url = GRAPH_BETA + path
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    resp = request_with_retry("POST", url, headers=headers, json_body=body)
    if resp.status_code >= 400:
        raise RuntimeError(f"Graph POST failed {resp.status_code}: {url}\n{resp.text}")
    return resp.json()


def list_mp4_in_folder_url(access_token: str, folder_url: str):
    # 共有リンクからの解決を試行
    print(f"[INFO] 共有リンクからアイテムを解決中: {folder_url}", flush=True)
    
    # 1. URLパラメータ 'id' をチェック (OneDriveブラウザURL対策)
    parsed = urlparse(folder_url)
    import urllib.parse
    qs = urllib.parse.parse_qs(parsed.query)
    id_param = qs.get("id", [None])[0]
    
    host = parsed.hostname
    drive_id = None
    folder_item_id = None

    if id_param:
        path = unquote(id_param)
        print(f"[DEBUG] URLの 'id' パラメータ検出: {path}", flush=True)
        # /personal/user_id/Documents/Path ... 形式を想定
        m = re.match(r"/personal/([^/]+)/Documents/(.*)$", path)
        if m:
            user_id = m.group(1)
            rel_path = m.group(2).rstrip('/')
            print(f"[INFO] 'id'パラメータ解釈成功: user={user_id}, path={rel_path}", flush=True)
            
            # 手道解決の試行 (複数の手法)
            upn_candidates = [
                user_id.replace("_", "."), # sadahiro_senda... -> sadahiro.senda...
                user_id.replace("_denso_com", "@denso.com").replace("_", "."), # sadahiro.senda.j6y.jp@denso.com
                user_id.replace("_", ".") + "@denso.com"
            ]
            
            for upn in upn_candidates:
                try:
                    print(f"[DEBUG] UPN候補試行: {upn}", flush=True)
                    drive = graph_get(access_token, f"/v1.0/users/{upn}/drive")
                    drive_id = drive["id"]
                    folder_item = graph_get(access_token, f"/v1.0/drives/{drive_id}/root:/{rel_path}")
                    folder_item_id = folder_item["id"]
                    print(f"[INFO] UPN解決成功: {upn} -> {folder_item_id}", flush=True)
                    break
                except Exception:
                    continue
            
            if not folder_item_id:
                try:
                    site = graph_get(access_token, f"/v1.0/sites/{host}:/personal/{user_id}")
                    drive = graph_get(access_token, f"/v1.0/sites/{site['id']}/drive")
                    drive_id = drive["id"]
                    folder_item = graph_get(access_token, f"/v1.0/drives/{drive_id}/root:/{rel_path}")
                    folder_item_id = folder_item["id"]
                    print(f"[INFO] 'id'パス解決成功: {folder_item_id}", flush=True)
                except Exception as e:
                    print(f"[WARN] 'id'パス解決失敗: {e}", flush=True)

    if not folder_item_id:
        share_id = encode_share_url(folder_url)
        try:
            # shares エンドポイントを使用
            di = graph_get(access_token, f"/v1.0/shares/{share_id}/driveItem")
            drive_id = di["parentReference"]["driveId"]
            folder_item_id = di["id"]
            print(f"[INFO] 共有リンク解決成功: Drive={drive_id}, Item={folder_item_id}", flush=True)
        except Exception as e:
            print(f"[WARN] sharesエンドポイント失敗: {e}", flush=True)
            # URL解析
            path = unquote(parsed.path)
            path = re.sub(r"^/:[a-z]:/r/", "/", path)
            m = re.match(r"/personal/([^/]+)/Documents/(.*)$", path)
            if not m:
                 raise RuntimeError(f"URL解析失敗: {folder_url}")
            
            user_id = m.group(1)
            # クエリを完全に排除
            rel_path = m.group(2).split('?')[0].split('&')[0].rstrip('/')
            
            print(f"[INFO] 手動解決開始: host={host}, user={user_id}, path={rel_path}", flush=True)
        # ここで例外をキャッチして続行するように修正
        try:
            site = graph_get(access_token, f"/v1.0/sites/{host}:/personal/{user_id}")
            drive = graph_get(access_token, f"/v1.0/sites/{site['id']}/drive")
            drive_id = drive["id"]
            
            candidates = [f"/{rel_path}", f"/Documents/{rel_path}"]
            folder_item = None
            for cand in candidates:
                try:
                    print(f"[DEBUG] 候補パス試行: {cand}", flush=True)
                    folder_item = graph_get(access_token, f"/v1.0/drives/{drive_id}/root:{cand}")
                    if folder_item: break
                except Exception: continue
            
            if not folder_item:
                print(f"[DEBUG] パス直接指定失敗。ルート直下を確認します。", flush=True)
                # ドライブ全体を確認
                root_info = graph_get(access_token, f"/v1.0/drives/{drive_id}/root")
                print(f"[DEBUG] Root Name: {root_info.get('name')}", flush=True)
                
                root_children = graph_get(access_token, f"/v1.0/drives/{drive_id}/root/children")
                child_names = [c['name'] for c in root_children.get("value", [])]
                print(f"[DEBUG] Root children: {child_names}", flush=True)
                
                # 'Documents' があるか確認
                if 'Documents' in child_names:
                    doc_folder = next(c for c in root_children.get("value", []) if c['name'] == 'Documents')
                    doc_children = graph_get(access_token, f"/v1.0/drives/{drive_id}/items/{doc_folder['id']}/children")
                    print(f"[DEBUG] Documents children: {[c['name'] for c in doc_children.get('value', [])]}", flush=True)
            
            if not folder_item:
                raise RuntimeError(f"フォルダ特定不能: {rel_path}")
            
            folder_item_id = folder_item["id"]
        except Exception as inner_e:
            print(f"[ERROR] 手動解決も失敗: {inner_e}", flush=True)
            raise

    print(f"[INFO] フォルダ内アイテム取得中 (Drive: {drive_id})", flush=True)
    # 子アイテム列挙 (MP4のみ)
    children = graph_get(
        access_token,
        f"/v1.0/drives/{drive_id}/items/{folder_item_id}/children",
        params={"$top": "999"}
    )

    mp4s = []
    for it in children.get("value", []):
        name = it.get("name", "")
        if name.lower().endswith(".mp4"):
            mp4s.append({
                "name": name,
                "driveId": drive_id,
                "itemId": it["id"],
                "webUrl": it.get("webUrl", ""),
            })
    return mp4s


def get_sharepoint_ids(access_token: str, drive_id: str, item_id: str):
    # sharepointIds の中に listItemUniqueId 等が入ることが多い
    it = graph_get(
        access_token,
        f"/v1.0/drives/{drive_id}/items/{item_id}",
        params={"$select": "id,name,webUrl,sharepointIds"}
    )
    return it.get("sharepointIds", {}), it.get("webUrl", ""), it.get("name", "")


def extract_personal_site_root(web_url: str) -> str:
    """
    例:
    https://globaldenso-my.sharepoint.com/personal/sadahiro_senda_j6y_jp_denso_com/Documents/Recordings/...
    -> https://globaldenso-my.sharepoint.com/personal/sadahiro_senda_j6y_jp_denso_com
    """
    u = urlparse(web_url)
    parts = [p for p in u.path.split("/") if p]
    # ["personal", "{user_upn_like}", "Documents", ...]
    if len(parts) < 2 or parts[0] != "personal":
        # チャネル会議など sites/xxx の可能性があるため、最低限 origin を返す
        return f"{u.scheme}://{u.netloc}"
    return f"{u.scheme}://{u.netloc}/personal/{parts[1]}"


def sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]+', "_", name).strip()


def download_transcripts_vtt(drive_id: str, item_id: str, site_root: str, out_dir: str):
    """
    SharePoint用トークンを取得して SharePoint REST v2.1 経由で列挙
    """
    # SharePoint用のトークンを明示的に取得 (Audienceエラー回避)
    # MySiteの場合は netloc に基づくリソースURLが必要
    u = urlparse(site_root)
    sp_scope = f"{u.scheme}://{u.netloc}/.default"
    sp_token = get_token(sp_scope)
    
    headers = {"Authorization": f"Bearer {sp_token}", "Accept": "application/json"}
    # SharePoint REST v2.1 エンドポイント
    url = f"{site_root}/_api/v2.1/drives/{drive_id}/items/{item_id}/media/transcripts"
    resp = request_with_retry("GET", url, headers=headers)
    if resp.status_code >= 400:
        raise RuntimeError(f"transcripts list failed {resp.status_code}: {url}\n{resp.text}")

    data = resp.json()
    vtts = []
    for idx, t in enumerate(data.get("value", []), start=1):
        dl = t.get("temporaryDownloadUrl")
        if not dl:
            continue
        vtt_path = os.path.join(out_dir, f"{idx}.vtt")
        # temporaryDownloadUrl は VTT(バイナリ)を返すため Accept: application/json は不可
        # また tempauth が付いているため Authorization ヘッダーなしでも通るはずだが
        # 406エラー回避のためヘッダーを最小限にする
        dl_headers = {}
        if sp_token:
             # 安全のため Authorization だけは残す（tempauthがあれば不要なことが多い）
             dl_headers["Authorization"] = f"Bearer {sp_token}"
        
        r2 = request_with_retry("GET", dl, headers=dl_headers, stream=True)
        if r2.status_code >= 400:
            raise RuntimeError(f"VTT download failed {r2.status_code}: {dl}\n{r2.text}")
        with open(vtt_path, "wb") as f:
            for chunk in r2.iter_content(chunk_size=1024 * 256):
                if chunk:
                    f.write(chunk)
        vtts.append(vtt_path)
    return vtts


def vtt_to_text(vtt_path: str) -> str:
    # WEBVTT とタイムコード行を落とす簡易整形
    out = []
    with open(vtt_path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            if line.startswith("WEBVTT"):
                continue
            if re.match(r"^\d{2}:\d{2}:\d{2}\.\d{3}\s+-->\s+\d{2}:\d{2}:\d{2}\.\d{3}", line):
                continue
            out.append(line)
    return "\n".join(out)


def chunk_text(text: str, size: int):
    return [text[i:i + size] for i in range(0, len(text), size)]


def copilot_summarize(access_token: str, text: str, title: str, tz="Asia/Tokyo", locale="ja-JP", chunk_chars=12000):
    """
    Copilot 会話API（Graph beta）:
      POST /beta/copilot/conversations で会話作成
      POST /beta/copilot/conversations/{id} に message.text を送信
    """
    # 会話作成
    conv = graph_post_beta(access_token, "/copilot/conversations", {})
    conv_id = conv["id"]

    def ask(prompt: str) -> str:
        body = {
            "message": {"text": prompt},
            "locationHint": {"timeZone": tz}
        }
        resp = graph_post_beta(access_token, f"/copilot/conversations/{conv_id}", body)
        # 返却messagesの末尾を採用（環境差があるので保険的に走査）
        msgs = resp.get("messages", [])
        for m in reversed(msgs):
            t = m.get("text")
            if t:
                return t
        return json.dumps(resp, ensure_ascii=False)

    base_prompt = f"""以下の会議トランスクリプトを要約してください（言語: {locale}）。
出力形式（この順で）:
- 会議名
- 目的
- 主要議題（3〜7点）
- 決定事項
- 未決事項
- 次アクション（担当/期限があれば明記）
会議名: {title}
トランスクリプト:
"""

    chunks = chunk_text(text, chunk_chars)
    partials = []
    for i, ch in enumerate(chunks, start=1):
        p = base_prompt + f"\n--- part {i}/{len(chunks)} ---\n" + ch
        partials.append(ask(p))

    if len(partials) == 1:
        return partials[0]

    merge_prompt = f"""以下は同一会議の分割要約です。重複を除去し、最終要約を作成してください。
出力形式:
- 会議名
- 目的
- 主要議題（3〜7点）
- 決定事項
- 未決事項
- 次アクション（担当/期限）
会議名: {title}
分割要約:
""" + "\n\n".join(partials)

    return ask(merge_prompt)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--folder-url", required=True, help="OneDriveのフォルダURL（onedrive.aspx?id=...）")
    ap.add_argument("--out-dir", required=True, help="出力先フォルダ")
    ap.add_argument("--locale", default="ja-JP")
    ap.add_argument("--chunk-chars", type=int, default=12000)
    args = ap.parse_args()

    out_dir = args.out_dir
    ensure_dir(out_dir)
    transcripts_root = os.path.join(out_dir, "transcripts")
    summaries_root = os.path.join(out_dir, "summaries")
    ensure_dir(transcripts_root)
    ensure_dir(summaries_root)

    # 1. 基本的な権限でファイル一覧を取得
    token = get_token()
    
    # 2. Copilot権限の取得を試みる (失敗しても続行)
    copilot_token = None
    try:
        # Copilot APIに必要な権限群
        cp_scopes = [
            f"{GRAPH}/Sites.Read.All",
            f"{GRAPH}/Mail.Read",
            f"{GRAPH}/People.Read.All",
            f"{GRAPH}/OnlineMeetingTranscript.Read.All",
            f"{GRAPH}/Chat.Read",
            f"{GRAPH}/ChannelMessage.Read.All",
            f"{GRAPH}/ExternalItem.Read.All",
            f"{GRAPH}/User.Read",
            f"{GRAPH}/Files.Read.All"
        ]
        copilot_token = get_token(cp_scopes)
        print("[INFO] Copilot API権限を取得しました。自動要約が有効です。", flush=True)
    except Exception as e:
        print(f"[WARN] Copilot API権限の取得に失敗しました (政策制限の可能性があります): {e}", flush=True)
        print("[INFO] 自動要約はスキップし、トランスクリプトの取得のみ継続します。", flush=True)

    mp4s = list_mp4_in_folder_url(token, args.folder_url)

    index_rows = []

    for mp4 in mp4s:
        name = mp4["name"]
        safe = sanitize_filename(os.path.splitext(name)[0])
        work_tr = os.path.join(transcripts_root, safe)
        work_sm = os.path.join(summaries_root, safe)

        # 既に transcript.txt を取得済みならスキップ（再ダウンロード・再要約しない）
        existing_transcript = os.path.join(work_sm, "transcript.txt")
        if os.path.exists(existing_transcript) and os.path.getsize(existing_transcript) > 0:
            index_rows.append({
                "name": name,
                "status": "ok",
                "out_dir": work_sm,
                "webUrl": mp4.get("webUrl", "")
            })
            print(f"[SKIP] {name} (既に取得済み)")
            continue

        ensure_dir(work_tr)
        ensure_dir(work_sm)

        try:
            spids, weburl, real_name = get_sharepoint_ids(token, mp4["driveId"], mp4["itemId"])
            site_root = extract_personal_site_root(weburl or mp4.get("webUrl", ""))

            vtt_files = download_transcripts_vtt(
                mp4["driveId"],
                mp4["itemId"],
                site_root,
                work_tr
            )
            if not vtt_files:
                raise RuntimeError("transcript(VTT)が取得できません（未生成/権限/ポリシーの可能性）")

            # 複数VTTがある場合は連結
            texts = [vtt_to_text(p) for p in vtt_files]
            transcript_text = "\n".join(texts).strip()

            with open(os.path.join(work_sm, "transcript.txt"), "w", encoding="utf-8") as f:
                f.write(transcript_text)

            if copilot_token:
                try:
                    summary = copilot_summarize(
                        copilot_token,
                        transcript_text,
                        title=real_name or name,
                        locale=args.locale,
                        chunk_chars=args.chunk_chars
                    )
                except Exception as ex:
                    print(f"[WARN] Copilot summary failed for {name}: {ex}", flush=True)
                    summary = "※Copilot APIエラーのため自動生成に失敗しました。後ほど手動で要約を追加してください。"
            else:
                summary = "※Copilot API権限不足（組織ポリシー等）のため、自動要約をスキップしました。\ntranscript.txt を参照し、手動でCopilotに要約を依頼してください。"

            with open(os.path.join(work_sm, "summary.md"), "w", encoding="utf-8") as f:
                f.write(summary)

            with open(os.path.join(work_sm, "meta.json"), "w", encoding="utf-8") as f:
                json.dump({
                    "name": real_name or name,
                    "webUrl": weburl or mp4.get("webUrl", ""),
                    "driveId": mp4["driveId"],
                    "itemId": mp4["itemId"],
                    "sharepointIds": spids,
                    "transcripts": vtt_files
                }, f, ensure_ascii=False, indent=2)

            index_rows.append({
                "name": real_name or name,
                "status": "ok",
                "out_dir": work_sm,
                "webUrl": weburl or mp4.get("webUrl", "")
            })
            print(f"[OK] {name}")

        except Exception as e:
            index_rows.append({
                "name": name,
                "status": "error",
                "error": str(e),
                "out_dir": work_sm,
                "webUrl": mp4.get("webUrl", "")
            })
            print(f"[NG] {name}: {e}")

    # index.csv 出力
    index_path = os.path.join(out_dir, "index.csv")
    with open(index_path, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["name", "status", "error", "out_dir", "webUrl"])
        w.writeheader()
        for r in index_rows:
            if "error" not in r:
                r["error"] = ""
            w.writerow(r)

    print(f"done: {index_path}")


if __name__ == "__main__":
    main()