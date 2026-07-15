# -*- coding: utf-8 -*-
"""
Teams会議録画URL(またはOneDrive/SharePoint録画フォルダURL) -> transcript(VTT)取得 -> 保存

vtt-only モード (Power Automate/Copilot Studio連携向け単体部品):
  python auto_summarize_recordings_v3.py --recording-url "<録画URL>" --out-dir "<出力先>" --mode vtt-only

dr-json モード (vtt-onlyに加えて、transcript.txtからDR勧告表用JSONを生成):
  python auto_summarize_recordings_v3.py --recording-url "<録画URL>" --out-dir "<出力先>" --mode dr-json

出力: result.json / transcript.vtt / transcript.txt / run.log (dr-jsonモードはさらに dr_kankokuhyo.json)

認証・URL解決・VTTダウンロードの実処理は auto_summarize_recordings.py (v1) の
実績のある関数をそのまま再利用する（v1側は一切変更しない）。
このファイルでは v1に無い「単一録画URLの解決」「vtt-only向けのVTT整形」
「Power Automate向けの result.json / 終了コード契約」「dr-json生成(GitHub Models API)」を追加する。

dr-json生成には GitHub Models / OpenAI互換 Chat Completions API を使う
(Graph beta /copilot/conversations は使わない。Teams Copilot/Excel Copilotの呼び出しでもない)。
Excel書き込みはこのファイルでは実装しない(範囲外、後段のPower Automateが担当)。
"""

import os
import re
import sys
import json
import math
import html
import shutil
import tempfile
import argparse
import traceback
from datetime import datetime, timedelta
from urllib.parse import urlparse, unquote, parse_qs

import requests

import auto_summarize_recordings as base


# ============================================================
# エラー分類用の例外
# ============================================================
class AuthError(Exception):
    """トークン取得(MSAL)に失敗した場合"""


class GraphPermissionError(Exception):
    """Graph/SharePointが401/403を返した場合(権限不足)"""


class TransientNetworkError(Exception):
    """429/5xx等、リトライしても解消しなかった場合"""


class RecordingNotFoundError(Exception):
    """録画URLからmp4を特定できなかった場合"""


class TranscriptNotFoundError(Exception):
    """mp4は特定できたが、トランスクリプト(VTT)が存在しない場合"""


class AIGenerationError(Exception):
    """dr-json生成に失敗した場合(APIキー未設定/入力サイズ超過/通信エラー/JSON解析失敗/スキーマ不一致/finish_reason=length等)"""

    def __init__(self, code: str, message: str, finish_reason: str = "", ai_model: str = ""):
        super().__init__(message)
        self.code = code
        self.finish_reason = finish_reason
        self.ai_model = ai_model


def _classify_runtime_error(e: Exception):
    """base.graph_get 等が投げる RuntimeError のメッセージからHTTPステータスを推測して分類する"""
    msg = str(e)
    m = re.search(r"\b(\d{3})\b", msg)
    code = int(m.group(1)) if m else None
    if code in (401, 403):
        return GraphPermissionError(msg)
    if code == 404:
        return RecordingNotFoundError(msg)
    if code in (429, 500, 502, 503, 504):
        return TransientNetworkError(msg)
    return None


def _call_classified(fn, *args, **kwargs):
    """v1(base)の関数を呼び、RuntimeErrorを分類済み例外に変換して再送出する"""
    try:
        return fn(*args, **kwargs)
    except RuntimeError as e:
        classified = _classify_runtime_error(e)
        if classified:
            raise classified from e
        raise


# ============================================================
# 標準出力/標準エラーを run.log にも同時書き込みするTee
# ============================================================
class Tee:
    def __init__(self, log_path):
        self.original_stdout = sys.stdout
        self.original_stderr = sys.stderr
        self.log_file = open(log_path, "a", encoding="utf-8")

    def write(self, data):
        self.original_stdout.write(data)
        try:
            self.log_file.write(data)
            self.log_file.flush()
        except ValueError:
            pass

    def flush(self):
        self.original_stdout.flush()
        try:
            self.log_file.flush()
        except ValueError:
            pass

    def close(self):
        self.log_file.close()


# ============================================================
# 単一録画URLの解決 (v1にはフォルダURL専用の解決しかないため新規追加)
# ============================================================
def _pick_single(candidates, warnings):
    if not candidates:
        raise RecordingNotFoundError("指定URL配下にmp4の録画が見つかりませんでした。")
    if len(candidates) > 1:
        names = ", ".join(c["name"] for c in candidates)
        warnings.append(
            f"複数の録画({len(candidates)}件: {names})が見つかったため、先頭の1件を対象にしました。"
        )
    return candidates[0], warnings


def _unwrap_meetingrecap_url(url: str) -> str:
    """
    Teamsの会議レコード共有リンク(https://teams.microsoft.com/l/meetingrecap?...)は
    Graphの/sharesエンドポイントで解決できる共有リンクではない。
    クエリパラメータ`fileUrl`に実際のSharePoint/OneDrive上のmp4パスが入っているので、
    存在すればそちらを実際の解決対象URLとして使う。
    """
    parsed = urlparse(url)
    if "meetingrecap" not in parsed.path.lower():
        return url
    qs = parse_qs(parsed.query)
    file_url = qs.get("fileUrl", [None])[0]
    if not file_url:
        return url
    file_url = unquote(file_url)
    print(f"[INFO] Teams会議録画リンク(meetingrecap)からfileUrlを抽出しました: {file_url}", flush=True)
    return file_url


def _extract_meetingrecap_params(url: str) -> dict:
    """
    meetingrecapリンクのクエリパラメータ(fileUrl/iCalUid/organizerId/threadId)をまとめて抽出する。
    meetingrecapリンクでなければ空 dict を返す。
    注意: iCalUidは以前はカレンダー予定を引き引く目的で使っていたが、
    「明示的な入力でない情報をbasicInfoに混ぜない」方針(2026-07-15)により、
    現在dr-json生成パイプラインからは使用していない。
    """
    parsed = urlparse(url)
    if "meetingrecap" not in parsed.path.lower():
        return {}
    qs = parse_qs(parsed.query)

    def _get(key):
        v = qs.get(key, [None])[0]
        return unquote(v) if v else None

    return {
        "fileUrl": _get("fileUrl"),
        "iCalUid": _get("iCalUid"),
        "organizerId": _get("organizerId"),
        "threadId": _get("threadId"),
    }


_FILENAME_DT_RE = re.compile(r"^(?P<title>.+)-(?P<date>\d{8})_(?P<time>\d{6})-(?P<suffix>.+)$")


def _parse_filename_metadata(name: str) -> dict:
    """
    Teamsの録画ファイル名は多くの場合
    "{会議名}-{YYYYMMDD}_{HHMMSS}-{種別}.mp4" の形式なので、
    ここから会議名(DR対象)を確実に抽出する(LLMに推測させない)。
    開催日時はここからは使わない(録画開始時刻≠会議開始時刻であり、
    meetingMeta.jsonのような明示入力がない限り「要確認」とする方針のため)。
    合致しなければ空 dict を返す。
    """
    base_name = os.path.splitext(name)[0]
    m = _FILENAME_DT_RE.match(base_name)
    if not m:
        return {}
    title = m.group("title").strip()
    return {"title": title}


def _try_resolve_personal_file_path(token: str, url: str):
    """
    共有リンク(shares API)でなく、`https://{tenant}-my.sharepoint.com/personal/{user}/Documents/{path}.mp4`
    のような個人サイトのパスを直接指す単一mp4のURLを解決する。
    v1のlist_mp4_in_folder_urlはフォルダ前提(最後に必ずchildren列挙する)のため、
    ファイル単体を指すURLには対応できない。そのためitem-by-path (root:/{path}) で直接解決する。
    該当しない/解決できない場合はNoneを返す(例外は投げない、呼び出し側でフォールバックする)。
    """
    parsed = urlparse(url)
    host = parsed.hostname
    if not host:
        return None
    path = unquote(parsed.path)
    m = re.match(r"^/personal/([^/]+)/Documents/(.*)$", path)
    if not m:
        return None
    user_id = m.group(1)
    rel_path = m.group(2).split("?")[0].split("&")[0].rstrip("/")
    if not rel_path.lower().endswith(".mp4"):
        return None

    site = _call_classified(base.graph_get, token, f"/v1.0/sites/{host}:/personal/{user_id}")
    drive = _call_classified(base.graph_get, token, f"/v1.0/sites/{site['id']}/drive")
    drive_id = drive["id"]
    item = _call_classified(
        base.graph_get, token, f"/v1.0/drives/{drive_id}/root:/{rel_path}",
        {"$select": "id,name,webUrl"},
    )
    return {
        "driveId": drive_id,
        "itemId": item["id"],
        "name": item.get("name", os.path.basename(rel_path)),
        "webUrl": item.get("webUrl", ""),
    }


def resolve_recording(token: str, recording_url: str):
    """
    recording_url が
      - Teams会議レコーディングリンク(meetingrecap) -> fileUrlクエリパラメータを抽出して以降の解決に使う
      - 単一mp4への共有リンク -> そのまま対象にする
      - 単一mp4への個人サイトパス直リンク(共有リンク形式でない) -> item-by-pathで直接解決する
      - フォルダへの共有リンク/ブラウザURL -> 配下のmp4を列挙し、1件ならそれを、複数なら先頭を対象にする(warningあり)
    のいずれにも対応する。
    戻り値: ({"driveId":..., "itemId":..., "name":..., "webUrl":...}, warnings)
    """
    warnings = []
    recording_url = _unwrap_meetingrecap_url(recording_url)
    item = None

    # 1) 共有リンクとして解決を試みる(ファイル/フォルダ両対応)
    try:
        share_id = base.encode_share_url(recording_url)
        item = _call_classified(
            base.graph_get,
            token,
            f"/v1.0/shares/{share_id}/driveItem",
            {"$select": "id,name,webUrl,parentReference,file,folder"},
        )
    except (GraphPermissionError, TransientNetworkError):
        raise
    except Exception as e:
        print(f"[INFO] 共有リンクとしての解決に失敗、他の解決方法を試します: {e}", flush=True)
        item = None

    if item is not None:
        if "file" in item:
            name = item.get("name", "")
            if not name.lower().endswith(".mp4"):
                raise RecordingNotFoundError(f"共有リンク先はmp4ではありません: {name}")
            return {
                "driveId": item["parentReference"]["driveId"],
                "itemId": item["id"],
                "name": name,
                "webUrl": item.get("webUrl", ""),
            }, warnings
        if "folder" in item:
            drive_id = item["parentReference"]["driveId"]
            children = _call_classified(
                base.graph_get,
                token,
                f"/v1.0/drives/{drive_id}/items/{item['id']}/children",
                {"$top": "999"},
            )
            candidates = [
                {"driveId": drive_id, "itemId": c["id"], "name": c["name"], "webUrl": c.get("webUrl", "")}
                for c in children.get("value", [])
                if c.get("name", "").lower().endswith(".mp4")
            ]
            return _pick_single(candidates, warnings)

    # 2) 共有リンクでない個人サイトパス直リンク(単一mp4)としての解決を試みる
    try:
        direct_item = _try_resolve_personal_file_path(token, recording_url)
    except (GraphPermissionError, TransientNetworkError):
        raise
    except Exception as e:
        print(f"[INFO] 個人サイトパス直リンクとしての解決にも失敗、フォルダURL解決にフォールバックします: {e}", flush=True)
        direct_item = None
    if direct_item is not None:
        return direct_item, warnings

    # 3) フォールバック: v1の list_mp4_in_folder_url (?id=パラメータ/personal site手動解決等) を再利用
    try:
        candidates = base.list_mp4_in_folder_url(token, recording_url)
    except RuntimeError as e:
        classified = _classify_runtime_error(e)
        if classified:
            raise classified from e
        raise RecordingNotFoundError(f"録画URLを解決できませんでした: {e}") from e
    except Exception as e:
        raise RecordingNotFoundError(f"録画URLを解決できませんでした: {e}") from e

    return _pick_single(candidates, warnings)


# ============================================================
# vtt-only モードのVTT整形 (v1のvtt_to_textはNOTE行除去に未対応のため新規実装)
#
# WebVTTのcue識別子は数字連番とは限らず、実データではGUID形式
# (例: "cb5b632b-20a7-4f07-965f-cb04e1e25a0b/30-0") も使われる。
# 識別子の文字列パターンを決め打ちで判定するのは頑健でないため、
# 空行区切りのブロック単位でパースし「タイムコード行より前にある行は
# すべて識別子として無視し、タイムコード行より後をcue本文とする」方式にする。
# ============================================================
_TIME_RE = re.compile(r"^\d{2}:\d{2}:\d{2}[.,]\d{3}\s+-->\s+\d{2}:\d{2}:\d{2}[.,]\d{3}")
_V_TAG_RE = re.compile(r"<v\s+([^>]+)>(.*?)(?:</v>)?$", re.IGNORECASE)


def _iter_vtt_cue_texts(raw_vtt: str):
    """WEBVTTをcue(空行区切りブロック)単位でパースし、各cueの本文行リストを順に返す。
    WEBVTTヘッダー・NOTEブロック・cue識別子行(数字/GUID等どんな形式でも)は自動的に除外する。"""
    text = raw_vtt.lstrip("\ufeff")  # 念のためBOM除去(呼び出し側で読み込み時に処理済みでも二重に安全策)
    blocks = re.split(r"\r?\n\r?\n+", text)
    for block in blocks:
        lines = [ln for ln in block.splitlines() if ln.strip() != ""]
        if not lines:
            continue
        head = lines[0].strip().upper()
        if head.startswith("WEBVTT") or head.startswith("NOTE") or head.startswith("STYLE") or head.startswith("REGION"):
            continue
        time_idx = None
        for i, ln in enumerate(lines):
            if _TIME_RE.match(ln.strip()):
                time_idx = i
                break
        if time_idx is None:
            continue  # タイムコードが見つからないブロックはヘッダ等とみなし無視
        cue_lines = [ln.strip() for ln in lines[time_idx + 1:] if ln.strip()]
        if cue_lines:
            yield cue_lines


def format_transcript_text(raw_vtt: str) -> str:
    """WEBVTTヘッダー/NOTE/タイムコード/cue識別子を除去し、話者名を保持したテキストにする"""
    lines_out = []
    for cue_lines in _iter_vtt_cue_texts(raw_vtt):
        # HTML実体参照(&lt;v ...&gt;)でエスケープされた話者タグも不備えだけ対応できるように先にアンエスケープする
        joined = html.unescape(" ".join(cue_lines))
        m = _V_TAG_RE.match(joined)
        if m:
            speaker = m.group(1).strip()
            spoken = re.sub(r"</?v[^>]*>", "", m.group(2)).strip()
            if spoken:
                lines_out.append(f"{speaker}: {spoken}")
            continue
        stripped = re.sub(r"<[^>]+>", "", joined).strip()
        if stripped:
            lines_out.append(stripped)
    return "\n".join(lines_out)


def analyze_vtt(raw_vtt: str):
    """話者集合と発言(utterance)件数を数える"""
    speakers = set()
    utterance_count = 0
    for cue_lines in _iter_vtt_cue_texts(raw_vtt):
        joined = html.unescape(" ".join(cue_lines))
        m = _V_TAG_RE.match(joined)
        if m:
            speaker = m.group(1).strip()
            spoken = re.sub(r"</?v[^>]*>", "", m.group(2)).strip()
            if spoken:
                speakers.add(speaker)
                utterance_count += 1
            continue
        stripped = re.sub(r"<[^>]+>", "", joined).strip()
        if stripped:
            utterance_count += 1
    return speakers, utterance_count


def download_transcripts_vtt_tagged(drive_id: str, item_id: str, site_root: str, out_dir: str):
    """
    v1のbase.download_transcripts_vtt()と同じ処理だが、temporaryDownloadUrlに
    `is=1&applymediaedits=false` を付与してから本体をダウンロードする点だけが異なる。

    背景: SharePoint REST v2.1 media/transcripts のtemporaryDownloadUrlをそのまま取得すると
    話者タグ(`<v 話者名>`)が剥がされたVTTが返る(本文/タイムコード/cue識別子は同一)。
    実際の手動ダウンロード時のNetworkリクエストを解析した結果、上記クエリを付与するだけで
    話者タグ付きのVTT(手動DL版とバイト完全一致)が取得できることを実証済み(2026-07-15)。
    v1(base.download_transcripts_vtt)は一切変更しない。
    """
    u = urlparse(site_root)
    sp_scope = f"{u.scheme}://{u.netloc}/.default"
    sp_token = base.get_token(sp_scope)

    headers = {"Authorization": f"Bearer {sp_token}", "Accept": "application/json"}
    list_url = f"{site_root}/_api/v2.1/drives/{drive_id}/items/{item_id}/media/transcripts"
    resp = _call_classified(base.request_with_retry, "GET", list_url, headers=headers)
    if resp is None or resp.status_code >= 400:
        status = resp.status_code if resp is not None else "timeout"
        text = resp.text if resp is not None else ""
        raise RuntimeError(f"transcripts list failed {status}: {list_url}\n{text}")

    data = resp.json()
    vtts = []
    for idx, t in enumerate(data.get("value", []), start=1):
        dl = t.get("temporaryDownloadUrl")
        if not dl:
            continue

        # 話者タグを保持させるクエリを tempauth の前に挿入する
        if "?" in dl:
            base_no_query, query_part = dl.split("?", 1)
            dl_tagged = f"{base_no_query}?is=1&applymediaedits=false&{query_part}"
        else:
            dl_tagged = f"{dl}?is=1&applymediaedits=false"

        vtt_path = os.path.join(out_dir, f"{idx}.vtt")
        dl_headers = {"Authorization": f"Bearer {sp_token}"} if sp_token else {}
        r2 = _call_classified(base.request_with_retry, "GET", dl_tagged, headers=dl_headers, stream=True)
        if r2 is None or r2.status_code >= 400:
            status = r2.status_code if r2 is not None else "timeout"
            text = r2.text if r2 is not None else ""
            raise RuntimeError(f"VTT download failed {status}: {dl_tagged}\n{text}")
        with open(vtt_path, "wb") as f:
            for chunk in r2.iter_content(chunk_size=1024 * 256):
                if chunk:
                    f.write(chunk)
        vtts.append(vtt_path)
    return vtts


def download_vtt_only(token: str, item: dict, out_dir: str):
    """v1のget_sharepoint_ids/extract_personal_site_rootと、話者タグを保持するdownload_transcripts_vtt_tagged()を使い、
    transcript.vtt / transcript.txt を out_dir 直下に生成する。"""
    spids, weburl, real_name = _call_classified(base.get_sharepoint_ids, token, item["driveId"], item["itemId"])
    site_root = base.extract_personal_site_root(weburl or item.get("webUrl", ""))

    tmp_dir = tempfile.mkdtemp(prefix="vtt_raw_")
    try:
        vtt_files = _call_classified(download_transcripts_vtt_tagged, item["driveId"], item["itemId"], site_root, tmp_dir)
        if not vtt_files:
            raise TranscriptNotFoundError("Transcript was not found for the recording.")

        raw_parts = []
        for p in vtt_files:
            # utf-8-sig: 実データでUTF-8 BOM付きVTTが確認されたため、BOMを自動除去して読み込む
            with open(p, "r", encoding="utf-8-sig", errors="ignore") as f:
                raw_parts.append(f.read())
        combined_raw = "\n\n".join(raw_parts)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    with open(os.path.join(out_dir, "transcript.vtt"), "w", encoding="utf-8") as f:
        f.write(combined_raw)

    text = format_transcript_text(combined_raw)
    with open(os.path.join(out_dir, "transcript.txt"), "w", encoding="utf-8") as f:
        f.write(text)

    if not text.strip():
        raise TranscriptNotFoundError("Transcript was downloaded but produced empty text.")

    speakers, utterance_count = analyze_vtt(combined_raw)
    return (real_name or item.get("name", "")), speakers, utterance_count


# ============================================================
# dr-json モード: GitHub Models / OpenAI互換 Chat Completions API での DR勧告表JSON生成
# ============================================================
GH_MODELS_ENDPOINT = "https://models.inference.ai.azure.com/chat/completions"
DEFAULT_DR_JSON_MODEL = "gpt-4o"

# 保守的な初期しきい値(実測で調整可能)。gpt-4oの128kコンテキストに対し、
# プロンプト定型文+出力分の余裕を見て文字数ベースでガードする。
MAX_TRANSCRIPT_CHARS = 60000

DR_JSON_REQUIRED_KEYS = [
    "basicInfo", "meetingInfo", "findings", "confirmations", "warnings", "sourceEvidence",
]

_DR_JSON_SCHEMA_EXAMPLE = """{
  "basicInfo": {
    "引当車両": "要確認",
    "DR対象": "要確認",
    "開催日時": "要確認",
    "開催場所": "要確認",
    "出席者": "要確認"
  },
  "meetingInfo": {
    "会議目的": "要確認",
    "会議結論": "要確認",
    "会議議事": "要確認",
    "会議資料": "要確認"
  },
  "findings": [
    {
      "No.": "1",
      "機能": "要確認",
      "指摘理由／指摘内容": "要確認",
      "検討省略": "要確認",
      "担当部署": "要確認",
      "担当者": "要確認",
      "実施期限": "要確認",
      "処置内容": "要確認",
      "完了日": "要確認"
    }
  ],
  "confirmations": [
    {"No.": "1", "確認事項": "要確認"}
  ],
  "warnings": [],
  "sourceEvidence": [
    {"item": "会議目的", "evidence": "発言: 『...』"},
    {"item": "DR対象", "evidence": "recordingFileName: {ファイル名}"}
  ]
}"""


def _load_gh_models_api_key():
    """
    GH_MODELS_TOKEN 環境変数を優先。無ければ tools/.gh_models_token (PoC用フォールバック) を読む。
    共有実行環境(Power Automate Desktop等)では GH_MODELS_TOKEN の設定を前提とし、
    個人用ファイルへの依存は避けること。
    戻り値: (api_key, key_source) 見つからなければ ("", "")
    """
    key = os.environ.get("GH_MODELS_TOKEN", "").strip()
    if key:
        return key, "env:GH_MODELS_TOKEN"

    token_path = os.path.join("tools", ".gh_models_token")
    if os.path.exists(token_path):
        try:
            with open(token_path, "r", encoding="utf-8") as f:
                key = f.read().strip()
            if key:
                return key, "file:tools/.gh_models_token (PoC用フォールバック)"
        except Exception:
            pass
    return "", ""


# sourceEvidence/basicInfoで使用禁止の根拠表現(meetingMeta.json等の明示入力が無い限り)
_FORBIDDEN_EVIDENCE_PHRASES = ["カレンダー予定", "カレンダー招待者", "招待者", "既知情報"]


def _build_dr_json_prompt(transcript_text: str, recording_file_name: str, speaker_count: int) -> str:
    return f"""あなたはDR(デザインレビュー)勧告表の下書き作成を支援するアシスタントです。
以下の情報**だけ**を根拠に、DR勧告表用のJSONを生成してください。

利用可能な情報源(これ以外の情報は一切提供されていません):
1. 会議トランスクリプト全文(下部に記載)
2. 録画ファイル名: {recording_file_name}
3. トランスクリプトから検出された話者数: {speaker_count}

カレンダー予定、出席者名簿、招待者一覧などの情報は一切提供されていません。存在しないものとして扱ってください。

最重要ルール:
- 出力はJSONのみとしてください。前置き、解説、コードフェンス等、JSON以外の文章を一切含めないでください。
- トランスクリプトと録画ファイル名に明示されていない情報は絶対に推測・補完しないでください。不明な項目は必ず「要確認」にしてください。
- 人名の漢字を推測しないでください。発言者表記、トランスクリプト内に存在する表記だけを使ってください。
- 担当者、実施期限、処置内容は、会議内で明示された場合のみ記入してください。
- 指摘事項はDR勧告表に転記しやすい粒度にしてください。指摘が無い場合はfindingsを空配列にしてください。
- 評価DRではない会議(状況共有・相談等)を無理に評価DR化しないでください。
- 会議目的、結論、議事、資料、指摘事項、確認事項は、それぞれ混同せず分けて記載してください。

basicInfoの各項目のルール(最重要、厳守):
- 開催日時、開催場所は、必ず「要確認」としてください(この処理では取得できません)。
- 出席者は、トランスクリプト検出話者数が{speaker_count}件です。
  - 話者数が0件の場合、出席者は必ず「要確認」としてください。トランスクリプトの発言内容から出席者名を推測しないでください。
  - 話者数が1件以上の場合でも、出席者欄はこちら側で確定的に設定するため、あなたは出席者欄に何を書いても構いませんが、「要確認」以外にする場合は必ずトランスクリプト内に実際に登場する話者表記のみを使ってください。カレンダーや名簿からの補完は禁止です。
- DR対象は、録画ファイル名(上記2)から読み取れる会議名をそのまま使って構いません。その場合のsourceEvidenceは「recordingFileName: {recording_file_name}」としてください。
- 引当車両は、トランスクリプトに明示的な記載が無い限り「要確認」としてください。

sourceEvidenceの形式(重要):
- sourceEvidenceは文字列配列ではなく、{{"item": "対象項目名", "evidence": "根拠"}} のオブジェクト配列にしてください。
- evidenceには、発言の短い引用、または「recordingFileName: {recording_file_name}」のみを使ってください。
- 「カレンダー予定」「招待者」「既知情報」という語句は絶対に使用しないでください(そのような情報は提供されていません)。
- meetingInfo・findings・confirmationsのうち「要確認」以外の値を持つ項目には、対応するsourceEvidenceを必ず1件以上含めてください。

出力するJSONは、必ず次のトップレベルキーを全て含めてください:
basicInfo, meetingInfo, findings, confirmations, warnings, sourceEvidence

スキーマ例(値は例示であり、実際は会議内容に基づいて埋めること):
{_DR_JSON_SCHEMA_EXAMPLE}

トランスクリプト:
{transcript_text}
"""


def _validate_dr_json_grounding(dr_json: dict, speaker_count: int, transcript_text: str):
    """
    生成されたJSONに根拠のない情報(カレンダー/招待者など)が混入していないかを検証する。
    違反を検知した場合は AIGenerationError(ai_generation_failed) を送出する
    (黙って修正するのではなく、モデルの違反を明示的に失敗扱いにする)。
    """
    basic_info = dr_json.get("basicInfo") or {}
    source_evidence = dr_json.get("sourceEvidence") or []

    # 1) speakerCount==0 の場合、出席者は必ず「要確認」
    attendees_val = str(basic_info.get("出席者", ""))
    if speaker_count == 0 and attendees_val != "要確認":
        raise AIGenerationError(
            "ai_generation_failed",
            f"speakerCount=0にも関わらずbasicInfo.出席者が「要確認」以外です(補完の疑い): {attendees_val!r}",
        )

    # 2) 開催日時・開催場所はmeetingMeta.json無しでは常に「要確認」
    for field in ("開催日時", "開催場所"):
        val = str(basic_info.get(field, ""))
        if val != "要確認":
            raise AIGenerationError(
                "ai_generation_failed",
                f"meetingMeta.json無しでbasicInfo.{field}が「要確認」以外です(補完の疑い): {val!r}",
            )

    # 3) 禁止根拠表現のチェック(sourceEvidence全体 + basicInfoの値全体)
    texts_to_check = []
    for ev in source_evidence:
        texts_to_check.append(str(ev.get("evidence", "")) if isinstance(ev, dict) else str(ev))
    for v in basic_info.values():
        texts_to_check.append(str(v))
    for phrase in _FORBIDDEN_EVIDENCE_PHRASES:
        for text in texts_to_check:
            if phrase in text:
                raise AIGenerationError(
                    "ai_generation_failed",
                    f"禁止された根拠表現「{phrase}」が含まれています: {text!r}",
                )

    # 4) speakerCount>0の場合、出席者に記載された人名がtranscript.txt内に実在するか確認(簡易チェック)
    #    LLMが稀にPython/JSON風のリスト表記("['A', 'B']"等)で返すことがあるため、
    #    括弧・引用符ノイズを除去してから区切り文字で分割する。
    if speaker_count > 0 and attendees_val not in ("", "要確認"):
        cleaned = re.sub(r"[\[\]'\"]", "", attendees_val)
        candidates = [c.strip() for c in re.split(r"[、,，/;；・]", cleaned) if c.strip()]
        candidates = [re.sub(r"\(.*?\)", "", c).strip() for c in candidates]
        candidates = [c for c in candidates if c]
        for name in candidates:
            if name not in transcript_text:
                raise AIGenerationError(
                    "ai_generation_failed",
                    f"出席者「{name}」がtranscript.txt内に見つかりません(補完の疑い)",
                )


def _finalize_dr_json(dr_json: dict, recording_file_name: str, speaker_count: int, speaker_names) -> dict:
    """
    検証済みのJSONに対し、確定的に分かっている情報(録画ファイル名/実際の検出話者)を
    確認・上書きし、必須のwarnings(speaker_not_detected/non_dr_meeting)を付与する。
    """
    basic_info = dr_json.setdefault("basicInfo", {})
    source_evidence = dr_json.get("sourceEvidence")
    if not isinstance(source_evidence, list):
        source_evidence = []

    filename_meta = _parse_filename_metadata(recording_file_name)
    if filename_meta.get("title"):
        basic_info["DR対象"] = filename_meta["title"]
        source_evidence = [e for e in source_evidence if not (isinstance(e, dict) and e.get("item") == "DR対象")]
        source_evidence.append({"item": "DR対象", "evidence": f"recordingFileName: {recording_file_name}"})

    if speaker_count > 0 and speaker_names:
        basic_info["出席者"] = "、".join(sorted(speaker_names))
        source_evidence = [e for e in source_evidence if not (isinstance(e, dict) and e.get("item") == "出席者")]
        source_evidence.append({"item": "出席者", "evidence": "transcript.vtt内の話者タグより"})

    dr_json["sourceEvidence"] = source_evidence
    dr_json["basicInfo"] = basic_info

    warnings_list = dr_json.get("warnings")
    if not isinstance(warnings_list, list):
        warnings_list = []
    codes_present = {w.get("code") for w in warnings_list if isinstance(w, dict)}
    if speaker_count == 0 and "speaker_not_detected" not in codes_present:
        warnings_list.append({
            "code": "speaker_not_detected",
            "message": "transcript.txtから話者名を抽出できなかったため、出席者は要確認としました。",
        })
    if not dr_json.get("findings") and "non_dr_meeting" not in codes_present:
        warnings_list.append({
            "code": "non_dr_meeting",
            "message": "DR指摘として扱える明示的な発言が少ないため、findingsは空配列としました。",
        })
    dr_json["warnings"] = warnings_list
    return dr_json


def generate_dr_json(
    transcript_text: str,
    recording_file_name: str,
    speaker_count: int,
    speaker_names=None,
    model: str = DEFAULT_DR_JSON_MODEL,
):
    """
    GitHub Models / OpenAI互換 Chat Completions API に transcript_text を送信し、
    DR勧告表用JSON(dict)を生成する。
    生成後に厄格なバリデーション(_validate_dr_json_grounding)を行い、カレンダー予定等
    提供されていない情報を根拠にした形跡があれば AIGenerationError で失敗扱いにする。
    検証合格後、DR対象(録画ファイル名より)と出席者(実際の検出話者より)を
    確定的に上書きし、必須warnings(speaker_not_detected/non_dr_meeting)を付与する(_finalize_dr_json)。
    戻り値: (dr_json_dict, meta) meta = {"finishReason": str, "aiModel": str}
    失敗時は AIGenerationError を送出する(finish_reason/ai_modelを可能な範囲で保持)。
    """
    speaker_names = speaker_names or set()
    api_key, key_source = _load_gh_models_api_key()
    if not api_key:
        raise AIGenerationError(
            "ai_generation_not_available",
            "GH_MODELS_TOKEN環境変数もtools/.gh_models_tokenも見つかりません。APIキーを設定してください。",
            ai_model=model,
        )

    proxies = {}
    proxy_url = os.environ.get("HTTPS_PROXY") or os.environ.get("HTTP_PROXY")
    if proxy_url:
        proxies = {"http": proxy_url, "https": proxy_url}

    prompt = _build_dr_json_prompt(transcript_text, recording_file_name, speaker_count)
    system_msg = "あなたはJSON以外を一切出力しないアシスタントです。応答は必ずJSONのみで返してください。"

    print(f"[INFO] dr-json生成: GitHub Models API呼び出し中 (model={model}, key_source={key_source})", flush=True)
    try:
        resp = requests.post(
            GH_MODELS_ENDPOINT,
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            json={
                "model": model,
                "messages": [
                    {"role": "system", "content": system_msg},
                    {"role": "user", "content": prompt},
                ],
                "temperature": 0.2,
                "max_tokens": 4000,
                "response_format": {"type": "json_object"},
            },
            proxies=proxies or None,
            timeout=180,
        )
    except requests.exceptions.RequestException as e:
        raise AIGenerationError(
            "ai_generation_failed", f"GitHub Models APIへの通信に失敗しました: {e}", ai_model=model
        ) from e

    if resp.status_code != 200:
        raise AIGenerationError(
            "ai_generation_failed",
            f"GitHub Models APIがエラーを返しました({resp.status_code}): {resp.text[:500]}",
            ai_model=model,
        )

    try:
        body = resp.json()
        choice = body["choices"][0]
        content = choice["message"]["content"]
        finish_reason = choice.get("finish_reason", "") or ""
    except Exception as e:
        raise AIGenerationError(
            "ai_generation_failed", f"GitHub Models APIの応答形式が不正です: {e}", ai_model=model
        ) from e

    if finish_reason == "length":
        raise AIGenerationError(
            "ai_generation_failed",
            "AI応答がmax_tokensで打ち切られました(finish_reason=length)。",
            finish_reason=finish_reason,
            ai_model=model,
        )

    try:
        dr_json = json.loads(content)
    except json.JSONDecodeError as e:
        raise AIGenerationError(
            "ai_generation_failed",
            f"AI応答をJSONとして解析できませんでした: {e}",
            finish_reason=finish_reason,
            ai_model=model,
        ) from e

    if not isinstance(dr_json, dict):
        raise AIGenerationError(
            "ai_generation_failed", "AI応答がJSONオブジェクトではありません。",
            finish_reason=finish_reason, ai_model=model,
        )

    missing = [k for k in DR_JSON_REQUIRED_KEYS if k not in dr_json]
    if missing:
        raise AIGenerationError(
            "ai_generation_failed",
            f"AI応答に必須キーが不足しています: {missing}",
            finish_reason=finish_reason, ai_model=model,
        )
    if not isinstance(dr_json.get("findings"), list):
        raise AIGenerationError(
            "ai_generation_failed", "findingsが配列ではありません。",
            finish_reason=finish_reason, ai_model=model,
        )
    if not isinstance(dr_json.get("confirmations"), list):
        raise AIGenerationError(
            "ai_generation_failed", "confirmationsが配列ではありません。",
            finish_reason=finish_reason, ai_model=model,
        )

    # 厄格バリデーション: 根拠のない情報(カレンダー/招待者等)の混入を検知したら失敗扱いにする
    try:
        _validate_dr_json_grounding(dr_json, speaker_count, transcript_text)
    except AIGenerationError as e:
        e.finish_reason = e.finish_reason or finish_reason
        e.ai_model = e.ai_model or model
        raise

    # 検証合格後、確定的に分かっている情報(DR対象/出席者)を上書きし、必須warningsを付与する
    dr_json = _finalize_dr_json(dr_json, recording_file_name, speaker_count, speaker_names)

    return dr_json, {"finishReason": finish_reason, "aiModel": model}


# ============================================================
# CLI エントリポイント
# ============================================================
def main():
    ap = argparse.ArgumentParser(description="Teams会議録画(MP4)からトランスクリプト(VTT)を取得する")
    ap.add_argument("--recording-url", dest="recording_url", help="録画mp4の共有リンク、または録画フォルダのURL")
    ap.add_argument("--folder-url", dest="folder_url", help="(互換用) 録画フォルダのURL。--recording-urlと同義")
    ap.add_argument("--out-dir", required=True, help="出力先フォルダ")
    ap.add_argument(
        "--mode", choices=["vtt-only", "dr-json"], default="vtt-only",
        help="vtt-only: VTT/テキスト取得のみ / dr-json: さらにDR勧告表用JSONを生成",
    )
    ap.add_argument("--ai-model", dest="ai_model", default=DEFAULT_DR_JSON_MODEL, help="dr-jsonで使うモデル名(既定: gpt-4o)")
    args = ap.parse_args()

    recording_url = args.recording_url or args.folder_url

    os.makedirs(args.out_dir, exist_ok=True)

    result = {
        "status": "error",
        "recordingUrl": recording_url or "",
        "driveId": "",
        "itemId": "",
        "recordingFileName": "",
        "vttFilePath": "",
        "textFilePath": "",
        "speakerCount": 0,
        "utteranceCount": 0,
        "warnings": [],
        "errors": [],
    }

    tee = Tee(os.path.join(args.out_dir, "run.log"))
    sys.stdout = tee
    sys.stderr = tee
    exit_code = 0
    try:
        if not recording_url:
            raise ValueError("--recording-url を指定してください。")

        try:
            token = base.get_token()
        except SystemExit as e:
            raise AuthError(str(e)) from e
        except Exception as e:
            raise AuthError(str(e)) from e

        item, warnings = resolve_recording(token, recording_url)
        result["warnings"].extend(warnings)
        result["driveId"] = item["driveId"]
        result["itemId"] = item["itemId"]

        name, speakers, utterance_count = download_vtt_only(token, item, args.out_dir)
        speaker_count = len(speakers)
        result["status"] = "success"
        result["recordingFileName"] = name
        result["vttFilePath"] = "transcript.vtt"
        result["textFilePath"] = "transcript.txt"
        result["speakerCount"] = speaker_count
        result["utteranceCount"] = utterance_count
        print(f"[OK] transcript.vtt / transcript.txt を生成しました ({name})", flush=True)

        if args.mode == "dr-json":
            result["mode"] = "dr-json"
            result["transcriptStatus"] = "success"
            result["drJsonStatus"] = "error"
            result["drJsonFilePath"] = ""
            result["aiModel"] = args.ai_model
            result["finishReason"] = ""

            transcript_path = os.path.join(args.out_dir, "transcript.txt")
            with open(transcript_path, "r", encoding="utf-8") as f:
                transcript_text = f.read()
            char_count = len(transcript_text)
            estimated_tokens = math.ceil(char_count / 1.5)
            result["transcriptCharCount"] = char_count
            result["estimatedInputTokens"] = estimated_tokens

            try:
                if char_count == 0:
                    raise AIGenerationError(
                        "ai_generation_not_available",
                        "transcript.txtが空のためAI生成をスキップしました。",
                        ai_model=args.ai_model,
                    )
                if char_count > MAX_TRANSCRIPT_CHARS:
                    raise AIGenerationError(
                        "ai_input_too_large",
                        f"transcriptが{char_count}文字でしきい値({MAX_TRANSCRIPT_CHARS}文字)を超えるため、AI呼び出しをスキップしました。",
                        ai_model=args.ai_model,
                    )

                dr_json, meta = generate_dr_json(
                    transcript_text,
                    recording_file_name=name,
                    speaker_count=speaker_count,
                    speaker_names=speakers,
                    model=args.ai_model,
                )

                dr_json_path = os.path.join(args.out_dir, "dr_kankokuhyo.json")
                with open(dr_json_path, "w", encoding="utf-8") as f:
                    json.dump(dr_json, f, ensure_ascii=False, indent=2)

                result["drJsonFilePath"] = "dr_kankokuhyo.json"
                result["drJsonStatus"] = "success"
                result["finishReason"] = meta.get("finishReason", "")
                result["aiModel"] = meta.get("aiModel", args.ai_model)
                print("[OK] dr_kankokuhyo.json を生成しました", flush=True)

            except AIGenerationError as e:
                result["status"] = "partial_success"
                result["drJsonStatus"] = "error"
                result["drJsonFilePath"] = ""
                result["finishReason"] = e.finish_reason
                if e.ai_model:
                    result["aiModel"] = e.ai_model
                result["warnings"].append({"code": e.code, "message": str(e)})
                exit_code = 6
                print(f"[NG] dr-json生成失敗({e.code}): {e}", flush=True)
            except Exception as e:
                result["status"] = "partial_success"
                result["drJsonStatus"] = "error"
                result["drJsonFilePath"] = ""
                result["warnings"].append({"code": "ai_generation_failed", "message": f"{type(e).__name__}: {e}"})
                exit_code = 6
                print(f"[NG] dr-json生成中に予期しないエラー: {e}", flush=True)
                traceback.print_exc()

    except RecordingNotFoundError as e:
        exit_code = 1
        result["errors"].append({"code": "recording_not_found", "message": str(e)})
        print(f"[NG] recording_not_found: {e}", flush=True)
    except TranscriptNotFoundError as e:
        exit_code = 1
        result["errors"].append({"code": "transcript_not_found", "message": str(e)})
        print(f"[NG] transcript_not_found: {e}", flush=True)
    except AuthError as e:
        exit_code = 2
        result["errors"].append({"code": "auth_error", "message": str(e)})
        print(f"[NG] auth_error: {e}", flush=True)
    except GraphPermissionError as e:
        exit_code = 3
        result["errors"].append({"code": "permission_error", "message": str(e)})
        print(f"[NG] permission_error: {e}", flush=True)
    except (TransientNetworkError, requests.exceptions.RequestException) as e:
        exit_code = 4
        result["errors"].append({"code": "network_error", "message": str(e)})
        print(f"[NG] network_error: {e}", flush=True)
    except Exception as e:
        exit_code = 5
        result["errors"].append({"code": "unexpected_error", "message": f"{type(e).__name__}: {e}"})
        print(f"[NG] unexpected_error: {e}", flush=True)
        traceback.print_exc()
    finally:
        result_path = os.path.join(args.out_dir, "result.json")
        with open(result_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        print(f"[DONE] result.json -> {result_path}", flush=True)
        sys.stdout = tee.original_stdout
        sys.stderr = tee.original_stderr
        tee.close()

    sys.exit(exit_code)


if __name__ == "__main__":
    main()
