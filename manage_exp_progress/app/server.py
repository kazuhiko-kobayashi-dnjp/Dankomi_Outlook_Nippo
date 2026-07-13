#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
露出制御業務進捗管理 - Flask APIサーバー

起動:
    python server.py
    ブラウザで http://localhost:3100 を開く

(このPCにNode.js/npmが未インストールのため、参考アプリ(manage_swreq_expisp)の
 Express構成と同等のAPIをFlaskで実装している)
"""
import io
import json
import re
import time
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory, send_file

from export_excel import export_workbook_bytes
from import_excel import refresh_tasks_from_excel

try:
    # SharePointへの自動バックアップ(任意機能)。msal未インストールでも本体機能に影響しないようにする。
    from sharepoint_backup import upload_backup_async
except Exception as _e:
    upload_backup_async = None
    print(f"[INFO] SharePointバックアップ機能は無効です(未セットアップ or msal未インストール): {_e}")

BASE_DIR = Path(__file__).parent
DATA_FILE = BASE_DIR / "data" / "tasks.json"
BACKUP_DIR = BASE_DIR / "data" / "backups"
BACKUP_KEEP = 60          # 直近この件数のみ保持(古いものから自動削除)
BACKUP_MIN_INTERVAL_SEC = 300  # 前回バックアップからこの秒数未満ならスキップ(連続保存で乗立しないように)
EXCEL_SOURCE_CONFIG = BASE_DIR / "data" / "excel_source_path.txt"
PUBLIC_DIR = BASE_DIR / "public"
PORT = 3100

app = Flask(__name__, static_folder=None)


def get_excel_source_path():
    """元Excel(他の人が継続メンテするxlsx)のパスを返す。
    PCごとにOneDriveのパスが異なるため、data/excel_source_path.txt で設定できる。"""
    if EXCEL_SOURCE_CONFIG.exists():
        # utf-8-sig: メモ帳等でBOM付きUTF-8として保存されても文字化けしないようにする
        p = EXCEL_SOURCE_CONFIG.read_text(encoding="utf-8-sig").strip()
        # エクスプローラーの「パスのコピー」はダブルクォートで囲まれるため、
        # そのまま貼り付けられても動くように前後の引用符を除去する
        p = p.strip('"').strip("'").strip()
        if p:
            return Path(p)
    return None


def read_data():
    with open(DATA_FILE, encoding="utf-8") as f:
        return json.load(f)


# バックアップファイル名(tasks_YYYYMMDD_HHMMSS.json)からタイムスタンプを取り出すための正規表現。
# パストラバーサル対策(このパターン以外のファイル名は拒否する)にも使う。
_BACKUP_NAME_RE = re.compile(r"^tasks_(\d{8})_(\d{6})\.json$")


def _parse_backup_stamp(filename):
    """バックアップファイル名に埋め込まれたタイムスタンプをUNIX時刻に変換する。
    ファイルのmtime(stat().st_mtime)はshutil.copy2でコピー元の更新日時が
    そのまま引き継がれてしまい「バックアップを作成した時刻」としては信頼できないため、
    ファイル名の方を正としてバックアップ間隔を判定する。"""
    m = _BACKUP_NAME_RE.match(filename)
    if not m:
        return None
    try:
        return time.mktime(time.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S"))
    except ValueError:
        return None


def _backup_before_write():
    """tasks.jsonを上書きする前に、変更前の内容をタイムスタンプ付きでdata/backups/に退避する。
    OneDriveのバージョン履歴機能とは別に、ローカルでもすぐに差分を確認・戻せるようにするための保险。
    失敗しても保存本体は続行する(バックアップ失敗で本来の保存をブロックしない)。"""
    if not DATA_FILE.exists():
        return
    try:
        BACKUP_DIR.mkdir(parents=True, exist_ok=True)
        existing = sorted(BACKUP_DIR.glob("tasks_*.json"))
        current_bytes = DATA_FILE.read_bytes()
        if existing:
            last = existing[-1]
            # 直前のバックアップと内容が完全に同じなら、中身が同じバージョンを
            # 何個も作らないようにスキップする(「バージョン履歴の中身が全部同じ」対策)。
            if last.read_bytes() == current_bytes:
                return
            last_stamp = _parse_backup_stamp(last.name)
            if last_stamp is not None and (time.time() - last_stamp) < BACKUP_MIN_INTERVAL_SEC:
                return
        stamp = time.strftime("%Y%m%d_%H%M%S")
        (BACKUP_DIR / f"tasks_{stamp}.json").write_bytes(current_bytes)
        all_backups = sorted(BACKUP_DIR.glob("tasks_*.json"))
        for old in all_backups[:-BACKUP_KEEP]:
            old.unlink(missing_ok=True)
    except Exception as e:
        print(f"[WARN] バックアップ作成に失敗しました: {e}")


def write_data(data):
    _backup_before_write()
    json_bytes = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    with open(DATA_FILE, "wb") as f:
        f.write(json_bytes)
    if upload_backup_async:
        try:
            upload_backup_async(json_bytes)
        except Exception as e:
            print(f"[WARN] SharePointバックアップの起動に失敗しました: {e}")


def refresh_meta(data):
    projects, categories, persons = set(), set(), set()
    for t in data["tasks"]:
        if t.get("project"):
            projects.add(t["project"])
        if t.get("category"):
            categories.add(t["category"])
        if t.get("person"):
            persons.add(t["person"])
    data.setdefault("meta", {})
    data["meta"]["projects"] = sorted(projects)
    data["meta"]["categories"] = sorted(categories)
    data["meta"]["persons"] = sorted(persons)
    return data


def to_number(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ========== 静的ファイル ==========
@app.route("/")
def index():
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.route("/<path:filename>")
def static_files(filename):
    fpath = PUBLIC_DIR / filename
    if fpath.exists():
        return send_from_directory(PUBLIC_DIR, filename)
    return send_from_directory(PUBLIC_DIR, "index.html")


# ========== API: 一覧取得 ==========
@app.route("/api/tasks", methods=["GET"])
def get_tasks():
    return jsonify(read_data())


# ========== API: 元Excelから再取込 ==========
@app.route("/api/import/refresh", methods=["POST"])
def import_refresh():
    source_path = get_excel_source_path()
    if source_path is None:
        return jsonify({
            "error": f"元Excelのパスが未設定です。{EXCEL_SOURCE_CONFIG} に元Excelのフルパスを1行記載してください。"
        }), 400
    if not source_path.exists():
        return jsonify({"error": f"元Excelが見つかりません: {source_path}"}), 404

    try:
        data = read_data()
        new_data, summary = refresh_tasks_from_excel(source_path, data)
        refresh_meta(new_data)
        write_data(new_data)
        return jsonify({"ok": True, "source": str(source_path), **summary})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ========== API: バックアップ一覧・復元 ==========
# _backup_before_write() が保存のたびに data/backups/tasks_YYYYMMDD_HHMMSS.json として
# 自動保存している世代を、画面から一覧・復元できるようにする(OneDrive/SharePointの
# バージョン履歴とは別に、アプリ内で完結する簡易版バージョン管理)。
# (_BACKUP_NAME_RE は _parse_backup_stamp() と共用するためファイル冒頭で定義済み)


@app.route("/api/backups", methods=["GET"])
def list_backups():
    if not BACKUP_DIR.exists():
        return jsonify({"backups": []})
    items = []
    for p in sorted(BACKUP_DIR.glob("tasks_*.json"), reverse=True):
        m = _BACKUP_NAME_RE.match(p.name)
        if not m:
            continue
        d, t = m.group(1), m.group(2)
        label = f"{d[0:4]}/{d[4:6]}/{d[6:8]} {t[0:2]}:{t[2:4]}:{t[4:6]}"
        items.append({"filename": p.name, "label": label, "size": p.stat().st_size})
    return jsonify({"backups": items})


# 差分表示対象のフィールド(created_at/updated_at/rowは対象外)
_DIFF_FIELDS = ("id", "project", "category", "person", "plan", "recent",
                "deadline", "status_mark", "status", "manhours", "note2", "comment")


def _diff_for_restore(backup_tasks, current_tasks):
    """backup_tasksに復元した場合、現在(current_tasks)から何がどう変わるかを返す。"""
    backup_by_row = {t["row"]: t for t in backup_tasks}
    current_by_row = {t["row"]: t for t in current_tasks}
    will_be_removed = []  # 現在にはあるがこのバージョンには無い → 復元すると削除される
    will_be_added = []    # このバージョンにはあるが現在は無い → 復元すると復活する
    will_be_changed = []  # 内容が変わる
    for row in sorted(set(backup_by_row) | set(current_by_row)):
        b = backup_by_row.get(row)
        c = current_by_row.get(row)
        if c and not b:
            will_be_removed.append({"row": row, "id": c.get("id"), "project": c.get("project"), "category": c.get("category"), "person": c.get("person")})
        elif b and not c:
            will_be_added.append({"row": row, "id": b.get("id"), "project": b.get("project"), "category": b.get("category"), "person": b.get("person")})
        else:
            field_diffs = {}
            for f in _DIFF_FIELDS:
                if b.get(f) != c.get(f):
                    field_diffs[f] = {"current": c.get(f), "restored": b.get(f)}
            if field_diffs:
                will_be_changed.append({
                    "row": row, "id": b.get("id"), "project": b.get("project"), "category": b.get("category"),
                    "person": b.get("person") or c.get("person"),
                    "fields": field_diffs,
                })
    return {
        "will_be_removed": will_be_removed,
        "will_be_added": will_be_added,
        "will_be_changed": will_be_changed,
    }


@app.route("/api/backups/<filename>/diff", methods=["GET"])
def diff_backup(filename):
    if not _BACKUP_NAME_RE.match(filename):
        return jsonify({"error": "不正なファイル名です"}), 400
    backup_path = BACKUP_DIR / filename
    if not backup_path.exists():
        return jsonify({"error": "バックアップファイルが見つかりません"}), 404
    try:
        with open(backup_path, encoding="utf-8") as f:
            backup_data = json.load(f)
        current_data = read_data()
        diff = _diff_for_restore(backup_data.get("tasks", []), current_data.get("tasks", []))
        return jsonify(diff)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/backups/<filename>/restore", methods=["POST"])
def restore_backup(filename):
    # ファイル名検証(パストラバーサル対策。自動生成される命名規則以外は拒否)
    if not _BACKUP_NAME_RE.match(filename):
        return jsonify({"error": "不正なファイル名です"}), 400
    backup_path = BACKUP_DIR / filename
    if not backup_path.exists():
        return jsonify({"error": "バックアップファイルが見つかりません"}), 404
    try:
        with open(backup_path, encoding="utf-8") as f:
            restored_data = json.load(f)
        refresh_meta(restored_data)
        write_data(restored_data)  # 復元前の現在の状態も_backup_before_write()で自動退避される
        return jsonify({"ok": True, "restored_from": filename, "task_count": len(restored_data.get("tasks", []))})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ========== API: 1件取得 (row=Excel上の行番号。管理番号(id)は重複しうるため一意キーとしてrowを使う) ==========
@app.route("/api/tasks/<int:row>", methods=["GET"])
def get_task(row):
    data = read_data()
    item = next((t for t in data["tasks"] if t["row"] == row), None)
    if not item:
        return jsonify({"error": "Not found"}), 404
    return jsonify(item)


# ========== API: 新規作成 ==========
@app.route("/api/tasks", methods=["POST"])
def create_task():
    data = read_data()
    body = request.get_json(force=True) or {}
    max_row = max((t.get("row", 0) for t in data["tasks"]), default=0)
    new_id = (body.get("id") or "").strip() or f"new{int(time.time() * 1000)}"
    now = time_iso()
    new_task = {
        "id": new_id,
        "row": max_row + 1,
        "project": body.get("project", "") or "",
        "category": body.get("category", "") or "",
        "person": body.get("person", "") or "",
        "plan": body.get("plan", "") or "",
        "recent": body.get("recent", "") or "",
        "deadline": body.get("deadline") or None,
        "status_mark": body.get("status_mark", "") or "",
        "status": body.get("status", "") or "",
        "manhours": to_number(body.get("manhours")),
        "note2": body.get("note2", "") or "",
        "comment": body.get("comment", "") or "",
        "created_at": now,
        "updated_at": now,
    }
    data["tasks"].append(new_task)
    refresh_meta(data)
    write_data(data)
    return jsonify(new_task)


# ========== API: 更新 ==========
@app.route("/api/tasks/<int:row>", methods=["PUT"])
def update_task(row):
    data = read_data()
    idx = next((i for i, t in enumerate(data["tasks"]) if t["row"] == row), -1)
    if idx < 0:
        return jsonify({"error": "Not found"}), 404
    body = request.get_json(force=True) or {}
    body.pop("row", None)  # rowは一意キーのため上書き不可

    # 楽観的排他制御: 自分が画面を開いた時点のupdated_atと現在のupdated_atが
    # 異なる場合、他の人が先に更新済み（上書き事故の可能性）とみなして409を返す
    expected_updated_at = body.pop("expected_updated_at", None)
    current_updated_at = data["tasks"][idx].get("updated_at")
    if expected_updated_at and current_updated_at and expected_updated_at != current_updated_at:
        return jsonify({
            "error": "conflict",
            "message": "他の人がこのタスクを更新しました。最新の内容を確認してから保存し直してください。",
            "latest": data["tasks"][idx],
        }), 409

    updated = {**data["tasks"][idx], **body}
    updated["manhours"] = to_number(body.get("manhours", data["tasks"][idx].get("manhours")))
    updated["updated_at"] = time_iso()
    data["tasks"][idx] = updated
    refresh_meta(data)
    write_data(data)
    return jsonify(updated)


# ========== API: 削除 ==========
@app.route("/api/tasks/<int:row>", methods=["DELETE"])
def delete_task(row):
    data = read_data()
    before = len(data["tasks"])
    data["tasks"] = [t for t in data["tasks"] if t["row"] != row]
    if len(data["tasks"]) == before:
        return jsonify({"error": "Not found"}), 404
    refresh_meta(data)
    write_data(data)
    return jsonify({"deleted": row})


# ========== API: Excelエクスポート ==========
@app.route("/api/export/excel", methods=["GET"])
def export_excel_get():
    data = read_data()
    return _send_excel(data["tasks"])


@app.route("/api/export/excel", methods=["POST"])
def export_excel_post():
    data = read_data()
    body = request.get_json(force=True) or {}
    ids = body.get("ids")
    tasks = data["tasks"]
    if ids:
        tasks = [t for t in tasks if t["id"] in ids]
    return _send_excel(tasks)


def _send_excel(tasks):
    xlsx_bytes = export_workbook_bytes(tasks)
    buf = io.BytesIO(xlsx_bytes)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    resp = send_file(
        buf,
        as_attachment=True,
        download_name=f"露出制御業務進捗_{timestamp}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    # ブラウザ側でエクスポート結果がキャッシュされ古い内容が返り続けることを防ぐ
    resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"] = "no-cache"
    return resp


def time_iso():
    return time.strftime("%Y-%m-%dT%H:%M:%S")


if __name__ == "__main__":
    print("\n🚀 露出制御業務進捗管理サーバー起動")
    print(f"   http://localhost:{PORT}")
    print(f"   データ: {DATA_FILE}\n")
    app.run(host="0.0.0.0", port=PORT, debug=False)
