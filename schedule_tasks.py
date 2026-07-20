#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
schedule_tasks.py
露出制御業務進捗管理(manage_exp_progress, http://10.41.55.204:3100/ / tasks.json)
から担当タスクを読み取り、Outlook の今週スケジュールの空き時間帯に自動挿入する。
フォーカス時間は削除して空き枠を確保。
（--source excel を指定すると旧・進捗管理Excelを直接読み込む方式にも切替可能）

Usage:
    # dry-run（確認のみ）
    python schedule_tasks.py
    python schedule_tasks.py --week 260623   # 260623 を含む週（月曜起算）

    # 実際に Outlook を操作
    python schedule_tasks.py --execute
    python schedule_tasks.py --week 260623 --execute
"""

import sys
import re
import json
import argparse
import unicodedata
import warnings
from pathlib import Path
from datetime import datetime, timedelta

# ──────────────────────────────────────────────────────────────────────
# 個人設定 (tools/.user_config.json) 読み込み
# ──────────────────────────────────────────────────────────────────────
# wr.ps1 (tools/teams_weekly_report.py) と同じ設定ファイルを共有する。
# 配布先の人はこのファイルをコピーして task_owner / task_json_path / task_excel_path を
# 自分の値に書き換えれば、下記のハードコード値を変更せずに使える(このファイルはgit管理対象外)。
_USER_CONFIG_FILE = Path(__file__).parent / 'tools' / '.user_config.json'


def _load_user_config() -> dict:
    if not _USER_CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(_USER_CONFIG_FILE.read_text(encoding='utf-8'))
    except Exception as e:
        print(f'[WARN] {_USER_CONFIG_FILE} の読み込みに失敗しました: {e}')
        return {}


def _save_user_config(**updates):
    """既存設定に updates をマージして .user_config.json に保存"""
    cfg = _load_user_config()
    cfg.update(updates)
    _USER_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    _USER_CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'[INFO] 個人設定を保存しました: {_USER_CONFIG_FILE}')


# ──────────────────────────────────────────────────────────────────────
# 設定
# ──────────────────────────────────────────────────────────────────────
# 以下は小林個人の環境向けフォールバック値。配布先の人は tools/.user_config.json の
# task_owner / task_json_path / task_excel_path で上書きするか、--person / --json / --file 引数を使う。
_LEGACY_EXCEL_PATH = (r'C:\Users\10001179776\OneDrive - DENSO\2019\Else\99_temp'
                      r'\◎250217_露出制御業務進捗及び報告.xlsx')
# manage_exp_progress/app/data/tasks.json はリポジトリ相対パスで解決する
# (このリポジトリを自分のOneDrive配下にクローン/同期している人なら誰でもそのまま使える)。
DEFAULT_JSON_PATH = Path(__file__).parent / 'manage_exp_progress' / 'app' / 'data' / 'tasks.json'
EXCEL_PATH   = _LEGACY_EXCEL_PATH  # 後方互換のため維持。実値は main() で解決される
JSON_PATH    = str(DEFAULT_JSON_PATH)
SHEET_NAME   = 'Sheet1'
PERSON       = '小林'  # 実際の値は main() で --person / .user_config.json の task_owner により上書きされる
WORK_START   = 6.5  # 06:30
WORK_END     = 20   # 18:00
MIN_SLOT_MIN = 30   # 30 分未満のスロットは使わない
MAX_CHUNK_H  = 1.0  # 1 スロットに入れる最大時間（これを超えるタスクは分割）
FOCUS_PATTERN    = 'フォーカス時間'
AUTO_SCHED_MARKER = '[auto-sched]'  # 自動登録判別用 Body マーカー

# Outlook に登録されている分類名（完全一致でそのまま使用）
OUTLOOK_CATEGORIES = {
    'BEV', 'Blue category', 'DCAP', 'GSP3', 'GSP4',
    'PostGSP4', 'SA4', 'SA5', '管理業務', '休桪',
}
# Excel B列値 → Outlook 分類名 のキーワードマッピング（登録順に評価）
CATEGORY_MAP = [
    ('BEV',      'BEV'),        # BEV3 等
    ('一体型',  'DCAP'),       # GSP4一体型
    ('ミニス4',  'SA4'),        # ミニステーサ4
    ('ミニス5',  'SA5'),        # ミニステーサ5
    ('GSP5',     'PostGSP4'),   # GSP5
    ('その他',   '管理業務'),   # その他
    ('GSP3',     'GSP3'),
    ('GSP4',     'GSP4'),
    ('SA4',      'SA4'),
    ('SA5',      'SA5'),
    ('DCAP',     'DCAP'),
    ('PostGSP4', 'PostGSP4'),
]

# ──────────────────────────────────────────────────────────────────────
# 依存チェック
# ──────────────────────────────────────────────────────────────────────
try:
    import openpyxl
except ImportError:
    print('[ERROR] openpyxl がインストールされていません: pip install openpyxl')
    sys.exit(1)

try:
    import win32com.client
    import pywintypes
except ImportError:
    print('[ERROR] pywin32 がインストールされていません: pip install pywin32')
    sys.exit(1)


# ──────────────────────────────────────────────────────────────────────
# Excel タスク読み込み
# ──────────────────────────────────────────────────────────────────────
PRIORITY_ORDER = {'◎': 0, '○': 1, '●': 2}


def _normalize_category(raw: str) -> str:
    """NFKC正規化 + 対応表キーワードマッチで Excel B列値を Outlook 分類名に変換する。"""
    if not raw:
        return ''
    raw_n = unicodedata.normalize('NFKC', raw.strip())
    if raw_n in OUTLOOK_CATEGORIES:        # 完全一致を優先
        return raw_n
    for keyword, category in CATEGORY_MAP:
        kw_n = unicodedata.normalize('NFKC', keyword)
        if kw_n in raw_n:
            return category
    return ''                              # 一致なし → 分類なし


def load_tasks(excel_path: str = EXCEL_PATH) -> list:
    """
    D列=小林 かつ H列≠空白・済 かつ I列≠0 の行を抽出。

    Returns:
        list of {'title': str, 'duration_h': float}
    """
    warnings.filterwarnings('ignore')
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]

    tasks = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        d     = str(row[3] or '')
        h_val = row[7]
        i_val = row[8]

        # 高速化フィルタ: H列が空白 or 済 は除外
        if h_val in (None, '', '済'):
            continue
        # D列=小林 かつ I列≠0
        if PERSON not in d:
            continue
        if not i_val or i_val == 0:
            continue

        e_text     = str(row[4] or '').strip()
        f_text     = str(row[5] or '').strip()
        first_line = f_text.split('\n')[0].strip()
        if not first_line:
            continue

        # I列の値を作業時間（h）として使用（0.5〜4h にクランプ）
        try:
            dur_h = float(i_val)
        except (TypeError, ValueError):
            dur_h = 1.0
        dur_h = max(0.5, min(dur_h, 4.0))

        category = (_normalize_category(str(row[1] or '').strip())   # B列優先
                   or _normalize_category(first_line))               # タイトルからフォールバック

        tasks.append({'title': first_line, 'duration_h': dur_h,
                      'priority': h_val, 'category': category,
                      'e_col': e_text, 'f_col': f_text})

    wb.close()
    # H列の優先度順（◎ > ○ > ●）でソート
    tasks.sort(key=lambda t: PRIORITY_ORDER.get(str(t['priority'] or ''), 99))
    return tasks


_TAG_BLOCK_RE = re.compile(r'</?(?:div|p|br)[^>]*>', re.IGNORECASE)
_TAG_RE       = re.compile(r'<[^>]+>')


def _html_to_plain(text: str) -> str:
    """Web版(manage_exp_progress)のリッチテキスト(HTML)欄をプレーンテキストに変換する。"""
    if not text:
        return ''
    if '<' not in text:
        return text
    import html as _html
    t = _TAG_BLOCK_RE.sub('\n', text)
    t = _TAG_RE.sub('', t)
    return _html.unescape(t)


def _extract_tasks_from_web_data(data: dict) -> list:
    """
    manage_exp_progress の tasks.json 相当のデータ(dict)から抽出する共通ロジック。
    load_tasks_from_json() / load_tasks_from_url() の両方から使う。
    担当者=PERSON かつ 状況≠空白・済 かつ 工数≠0 の案件。
    """
    tasks = []
    for t in data.get('tasks', []):
        person = t.get('person') or ''
        status_mark = t.get('status_mark')
        manhours = t.get('manhours')

        if status_mark in (None, '', '済'):
            continue
        if PERSON not in person:
            continue
        if not manhours or manhours == 0:
            continue

        e_text = _html_to_plain(t.get('plan') or '').strip()
        f_text = _html_to_plain(t.get('recent') or '').strip()
        first_line = f_text.split('\n')[0].strip()
        if not first_line:
            continue

        try:
            dur_h = float(manhours)
        except (TypeError, ValueError):
            dur_h = 1.0
        dur_h = max(0.5, min(dur_h, 4.0))

        category = (_normalize_category(t.get('project') or '')
                   or _normalize_category(first_line))

        tasks.append({'title': first_line, 'duration_h': dur_h,
                      'priority': status_mark, 'category': category,
                      'e_col': e_text, 'f_col': f_text})

    tasks.sort(key=lambda t: PRIORITY_ORDER.get(str(t['priority'] or ''), 99))
    return tasks


def load_tasks_from_json(json_path: str = JSON_PATH) -> list:
    """
    manage_exp_progress の data/tasks.json をローカルファイルとして直接読み込む版。
    OneDrive同期等でこのファイルにローカルアクセスできる環境向け(既定)。
    """
    with open(json_path, encoding='utf-8') as f:
        data = json.load(f)
    return _extract_tasks_from_web_data(data)


def load_tasks_from_url(url: str) -> list:
    """
    manage_exp_progress の GET /api/tasks をHTTP経由で取得する版。
    OneDrive同期パスにアクセスできない配布先の人向け(サーバーに直接アクセスできれば
    ローカルファイルコピーが無くても使える)。
    """
    import requests
    resp = requests.get(url, timeout=10)
    resp.raise_for_status()
    return _extract_tasks_from_web_data(resp.json())


# ──────────────────────────────────────────────────────────────────────
# Outlook 操作ヘルパー
# ──────────────────────────────────────────────────────────────────────
def get_week_dates(week_start: datetime) -> list:
    """月曜日〜金曜日の5日間を返す。"""
    return [week_start + timedelta(days=i) for i in range(5)]


def get_calendar_events(calendar, date: datetime) -> list:
    """
    指定日の Outlook イベントを [(start_dt, end_dt, subject)] で返す。
    終日イベントは除外。BusyStatus=0(Free) は除外。
    """
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort('[Start]')

    d_str  = date.strftime('%Y/%m/%d')
    d_next = (date + timedelta(days=1)).strftime('%Y/%m/%d')
    flt    = f"[Start] >= '{d_str} 00:00' AND [Start] < '{d_next} 00:00'"
    found  = items.Restrict(flt)

    events = []
    for item in found:
        try:
            if item.AllDayEvent:
                continue
            if item.BusyStatus == 0:   # Free（Canceled 等）
                continue
            st = item.Start
            en = item.End
            s  = datetime(st.year, st.month, st.day, st.hour, st.minute)
            e  = datetime(en.year, en.month, en.day, en.hour, en.minute)
            if e > s:
                events.append((s, e, item.Subject or ''))
        except Exception:
            continue

    return sorted(events, key=lambda x: x[0])


def find_free_slots(events: list, date: datetime) -> list:
    """
    作業時間（WORK_START〜WORK_END）内の空きスロット (start_dt, end_dt) を返す。
    フォーカス時間イベントは空きとして扱う（別途削除する）。
    """
    _ws_h, _ws_m = int(WORK_START), int((WORK_START % 1) * 60)
    _we_h, _we_m = int(WORK_END),   int((WORK_END   % 1) * 60)
    day_start = date.replace(hour=_ws_h, minute=_ws_m, second=0, microsecond=0)
    day_end   = date.replace(hour=_we_h, minute=_we_m, second=0, microsecond=0)

    # フォーカス時間を除いた有効イベント（作業時間内のもの）
    effective = [
        (max(s, day_start), min(e, day_end))
        for s, e, sub in events
        if FOCUS_PATTERN not in sub and e > day_start and s < day_end
    ]
    effective = [(s, e) for s, e in effective if e > s]
    effective.sort()

    slots = []
    cur = day_start
    for s, e in effective:
        if cur < s:
            slots.append((cur, s))
        cur = max(cur, e)
    if cur < day_end:
        slots.append((cur, day_end))

    # MIN_SLOT_MIN 未満のスロットは除外
    return [(s, e) for s, e in slots
            if (e - s).total_seconds() / 60 >= MIN_SLOT_MIN]


def collect_focus_events(calendar, date: datetime) -> list:
    """指定日のフォーカス時間イベントオブジェクトのリストを返す。"""
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort('[Start]')

    d_str  = date.strftime('%Y/%m/%d')
    d_next = (date + timedelta(days=1)).strftime('%Y/%m/%d')
    flt    = f"[Start] >= '{d_str} 00:00' AND [Start] < '{d_next} 00:00'"
    found  = items.Restrict(flt)

    result = []
    for item in found:
        try:
            if FOCUS_PATTERN in (item.Subject or ''):
                st = item.Start
                start_dt = datetime(st.year, st.month, st.day, st.hour, st.minute)
                result.append((start_dt, item))
        except Exception:
            continue
    return result


def clear_scheduled_tasks(calendar, week_dates: list, task_titles: set,
                          dry_run: bool) -> int:
    """
    以下のいずれかに該当する予定を週全体から削除する。
      1. Body に AUTO_SCHED_MARKER が含まれる（今回以降の登録）
      2. Subject がタスクタイトル（末尾の "(n/m)" を除いた基底部分）と一致（過去登録分）
    dry_run=True のとき削除予定の表示のみ。削除（予定）件数を返す。
    """
    _suffix_re = re.compile(r'\s*\(\d+/\d+\)\s*$')

    count = 0
    for day in week_dates:
        day_label = day.strftime('%m/%d(%a)')
        items = calendar.Items
        items.IncludeRecurrences = True
        items.Sort('[Start]')
        d_str  = day.strftime('%Y/%m/%d')
        d_next = (day + timedelta(days=1)).strftime('%Y/%m/%d')
        flt    = f"[Start] >= '{d_str} 00:00' AND [Start] < '{d_next} 00:00'"
        found  = items.Restrict(flt)

        to_delete = []
        for item in found:
            try:
                sub  = item.Subject or ''
                body = item.Body or ''
                by_marker = AUTO_SCHED_MARKER in body
                by_title  = _suffix_re.sub('', sub).strip() in task_titles
                if by_marker or by_title:
                    st = item.Start
                    start_dt = datetime(st.year, st.month, st.day, st.hour, st.minute)
                    to_delete.append((start_dt, sub, item))
            except Exception:
                continue

        if to_delete:
            print(f'── {day_label} ─────────────────────')
            for start_dt, sub, item in to_delete:
                if dry_run:
                    print(f'  [削除予定] {start_dt.strftime("%H:%M")} {sub}')
                else:
                    print(f'  [削除]     {start_dt.strftime("%H:%M")} {sub}')
                    item.Delete()
                count += 1
    return count


def categorize_existing_events(calendar, date: datetime, dry_run: bool) -> None:
    """
    指定日の既存予定のうち、カテゴリ未設定かつ件名がキーワードに一致するものに
    自動で分類を付ける。dry_run=True のとき表示のみ。
    """
    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort('[Start]')
    d_str  = date.strftime('%Y/%m/%d')
    d_next = (date + timedelta(days=1)).strftime('%Y/%m/%d')
    flt    = f"[Start] >= '{d_str} 00:00' AND [Start] < '{d_next} 00:00'"
    found  = items.Restrict(flt)

    for item in found:
        try:
            if item.AllDayEvent:
                continue
            if item.BusyStatus == 0:        # Free（キャンセル等）
                continue
            if (item.Categories or '').strip():  # 既に分類済み → スキップ
                continue
            sub = item.Subject or ''
            cat = _normalize_category(sub)
            if not cat:
                continue
            st = item.Start
            t  = datetime(st.year, st.month, st.day, st.hour, st.minute)
            if dry_run:
                print(f'  [分類予定] {t.strftime("%H:%M")} {sub}  →  [{cat}]')
            else:
                item.Categories = cat
                item.Save()
                print(f'  [分類済み] {t.strftime("%H:%M")} {sub}  →  [{cat}]')
        except Exception:
            continue


# ──────────────────────────────────────────────────────────────────────
# メイン
# ──────────────────────────────────────────────────────────────────────
def main():
    global PERSON

    parser = argparse.ArgumentParser(
        description='進捗管理 Excel タスクを Outlook に自動スケジュール')
    parser.add_argument('--week', default=None,
                        help='対象週の任意の日 YYMMDD（省略=今週）')
    parser.add_argument('--execute', action='store_true',
                        help='実際に Outlook を操作する（省略=dry-run）')
    parser.add_argument('--clear', action='store_true',
                        help='自動登録タスクを削除してから再スケジュール（--execute と組み合わせ可）')
    parser.add_argument('--file', default='',
                        help='Excel ファイルパス (--source excel の時のみ使用。'
                             '省略時: tools/.user_config.json の task_excel_path)')
    parser.add_argument('--source', choices=['excel', 'web'], default='web',
                        help='タスクの取得元: web=manage_exp_progress(露出制御業務進捗管理, '
                             'http://10.41.55.204:3100/)のtasks.jsonから読み込み(デフォルト) / '
                             'excel=旧進捗管理Excelを直接読み込み')
    parser.add_argument('--json', default='',
                        help='tasks.json のローカルパス (--source web かつ --url 未指定の時のみ使用。'
                             '省略時: tools/.user_config.json の task_json_path、'
                             'それも無ければリポジトリ相対の manage_exp_progress/app/data/tasks.json)')
    parser.add_argument('--url', default='',
                        help='manage_exp_progress の GET /api/tasks のURL (--source web の時のみ使用)。'
                             '指定するとローカルファイル(--json)の代わりにHTTP経由でタスクを取得する。'
                             'OneDrive同期でtasks.jsonにローカルアクセスできない配布先の人向け '
                             '(例: http://10.41.55.204:3100/api/tasks)。'
                             '省略時: tools/.user_config.json の task_json_url')
    parser.add_argument('--person', default='',
                        help='担当者フィルタ文字列（省略時: tools/.user_config.json の task_owner、'
                             'それも無ければ既定値「小林」）')
    parser.add_argument('--save-person', action='store_true',
                        help='--person で指定した値を tools/.user_config.json に保存し、'
                             '次回から --person 省略可能にする')
    args = parser.parse_args()

    # ── 個人設定の解決: コマンドライン引数 > tools/.user_config.json > ハードコード既定値 ──
    user_cfg = _load_user_config()
    PERSON = args.person or user_cfg.get('task_owner') or PERSON
    excel_path = args.file or user_cfg.get('task_excel_path') or _LEGACY_EXCEL_PATH
    json_path = args.json or user_cfg.get('task_json_path') or str(DEFAULT_JSON_PATH)
    task_url = args.url or user_cfg.get('task_json_url') or ''

    if args.save_person and args.person:
        _save_user_config(task_owner=args.person)

    dry_run = not args.execute

    # 対象週の月曜日を決定
    if args.week:
        yw = args.week.zfill(6)
        base = datetime(2000 + int(yw[:2]), int(yw[2:4]), int(yw[4:6]))
    else:
        base = datetime.today()
    week_start = (base - timedelta(days=base.weekday())).replace(
        hour=0, minute=0, second=0, microsecond=0)

    week_end = week_start + timedelta(days=4)

    if args.source == 'web':
        source_label = (f'manage_exp_progress (HTTP: {task_url})' if task_url
                        else f'manage_exp_progress (tasks.json: {json_path})')
    else:
        source_label = f'進捗管理Excel（直接: {excel_path}）'
    print(f'[INFO] モード      : {"dry-run（確認のみ）" if dry_run else "EXECUTE（Outlook 操作）"}')
    print(f'[INFO] データ取得元  : {source_label}')
    print(f'[INFO] 担当者フィルタ: {PERSON}')
    print(f'[INFO] 対象週      : {week_start.strftime("%Y-%m-%d")} (月) ～ {week_end.strftime("%Y-%m-%d")} (金)')
    print()

    # ── タスク読み込み ──
    if args.source == 'web':
        tasks = load_tasks_from_url(task_url) if task_url else load_tasks_from_json(json_path)
    else:
        tasks = load_tasks(excel_path)
    if not tasks:
        print(f'[WARN] タスクなし（担当者={PERSON} かつ 状況≠空白/済 かつ 工数≠0）')
        return


    total_h = sum(t['duration_h'] for t in tasks)
    print(f'[INFO] タスク {len(tasks)} 件（合計 {total_h:.1f}h）')
    for t in tasks:
        h = int(t['duration_h'])
        m = int((t['duration_h'] - h) * 60)
        dur_str  = f'{h}h{m:02d}m' if m else f'{h}h'
        cat_str  = f'  [{t["category"]}]' if t.get('category') else ''
        prio_str = str(t.get('priority') or '')
        print(f'  {prio_str}[{dur_str}] {t["title"]}{cat_str}')
    print()

    # ── Outlook 接続 ──
    app      = win32com.client.Dispatch('Outlook.Application')
    ns       = app.GetNamespace('MAPI')
    calendar = ns.GetDefaultFolder(9)

    week_dates = get_week_dates(week_start)

    # ── --clear モード: 自動登録タスクを削除 ──
    if args.clear:
        print(f'[INFO] 自動登録タスクを{"削除予定（dry-run）" if dry_run else "削除"}します...')
        print()
        task_titles = {t['title'] for t in tasks}
        n_deleted = clear_scheduled_tasks(calendar, week_dates, task_titles, dry_run)
        print()
        print('=' * 60)
        print(f'{"[DRY-RUN] " if dry_run else ""}削除{"予定" if dry_run else "済み"}: {n_deleted} 件')
        if dry_run:
            print('→ 実際に削除するには --execute を付けて実行してください。')
            return
        print()
        # dry_run=False → 削除後にそのまま再スケジューリングへ続行

    # ── フォーカス時間削除 & 空きスロット収集（週全体） ──
    # week_slots: [start_dt, end_dt, cur_dt]  cur は未使用領域の先頭
    week_slots = []
    for day in week_dates:
        day_label = day.strftime('%m/%d(%a)')
        print(f'── {day_label} ─────────────────────')

        focus_items = collect_focus_events(calendar, day)
        for start_dt, item in focus_items:
            if dry_run:
                print(f'  [削除予定] {start_dt.strftime("%H:%M")} {item.Subject}')
            else:
                print(f'  [削除]     {start_dt.strftime("%H:%M")} {item.Subject}')
                item.Delete()

        events = get_calendar_events(calendar, day)
        slots  = find_free_slots(events, day)
        if not slots:
            print('  (空きスロットなし)')
        else:
            for s, e in slots:
                mins = int((e - s).total_seconds() / 60)
                print(f'  [空き] {s.strftime("%H:%M")}-{e.strftime("%H:%M")} ({mins}min)')
                week_slots.append([s, e, s])   # [orig_start, end, cur_pos]

        categorize_existing_events(calendar, day, dry_run)
        print()

    # ── スロットをラウンドロビン順に並べ替え（曜日間で均等分散） ──
    # 例: [Mon_0, Tue_0, Wed_0, Thu_0, Fri_0, Mon_1, Tue_1, ...]
    slots_by_day = {}
    for slot in week_slots:
        key = slot[0].date()
        slots_by_day.setdefault(key, []).append(slot)
    days_sorted = sorted(slots_by_day)
    max_per_day = max((len(v) for v in slots_by_day.values()), default=0)
    week_slots = [
        slots_by_day[day][i]
        for i in range(max_per_day)
        for day in days_sorted
        if i < len(slots_by_day[day])
    ]

    # ── タスクを優先度順に「週全体のスロット」へ割り当て ──
    # ◎ : 全量確保できなければロールバック（現行ロジック維持）
    # ○● : ラウンドロビン（MIN_SLOT_MIN ずつ公平割り当て。部分登録あり）
    scheduled           = []
    unscheduled         = []
    partially_scheduled = []   # ○● で枠不足・部分登録になったタスク

    def _register_chunks(task, chunks):
        """チャンクを日付順に並べ、表示・Outlook 登録・scheduled へ追記。"""
        chunks_sorted = sorted(chunks)
        total = len(chunks_sorted)
        for j, (start, end) in enumerate(chunks_sorted):
            suffix = f' ({j+1}/{total})' if total > 1 else ''
            title = task['title'] + suffix
            if dry_run:
                cat_str = f'  [{task["category"]}]' if task.get('category') else ''
                print(f'[登録予定] {start.strftime("%m/%d %H:%M")}-{end.strftime("%H:%M")}  {title}{cat_str}')
            else:
                appt             = app.CreateItem(1)
                appt.Subject     = title
                appt.Start       = start.strftime('%Y-%m-%d %H:%M')
                appt.End         = end.strftime('%Y-%m-%d %H:%M')
                appt.BusyStatus  = 2
                appt.ReminderSet = False
                body_parts = [AUTO_SCHED_MARKER]
                if task.get('e_col'):
                    body_parts.append(f'【E】{task["e_col"]}')
                if task.get('f_col') and task['f_col'] != task['title']:
                    body_parts.append('【内容】')
                    body_parts.append(task['f_col'])
                appt.Body        = '\n'.join(body_parts)
                if task.get('category'):
                    appt.Categories = task['category']
                appt.Save()
                cat_str = f'  [{task["category"]}]' if task.get('category') else ''
                print(f'[登録済み] {start.strftime("%m/%d %H:%M")}-{end.strftime("%H:%M")}  {title}{cat_str}')
            scheduled.append({'task': task, 'start': start, 'end': end, 'title': title})

    prio_key    = lambda t: PRIORITY_ORDER.get(str(t['priority'] or ''), 99)
    top_tasks   = [t for t in tasks if prio_key(t) == 0]   # ◎
    other_tasks = [t for t in tasks if prio_key(t) >  0]   # ○ ●

    # --- ◎ タスク: 全量 first-fit（現行ロジック） ---
    for task in top_tasks:
        remaining_m = int(task['duration_h'] * 60)
        chunk_m     = int(MAX_CHUNK_H * 60)
        task_chunks = []
        slot_backup = {id(s): s[2] for s in week_slots}

        for slot in week_slots:
            if remaining_m <= 0:
                break
            _, slot_end, cur = slot
            avail_m = int((slot_end - cur).total_seconds() / 60)
            if avail_m < MIN_SLOT_MIN:
                continue
            place_m = min(remaining_m, chunk_m, avail_m)
            start   = cur
            end     = cur + timedelta(minutes=place_m)
            slot[2] = end
            remaining_m -= place_m
            task_chunks.append((start, end))

        if remaining_m > 0:
            for slot in week_slots:
                slot[2] = slot_backup[id(slot)]
            unscheduled.append(task)
            continue

        _register_chunks(task, task_chunks)

    # --- ○● タスク: ラウンドロビン（MIN_SLOT_MIN ずつ公平割り当て） ---
    # 1ラウンドで各タスクに30分ずつ割り当て → 全タスクに最低1枠を保証してから追加分を配分
    rr_remaining = [int(t['duration_h'] * 60) for t in other_tasks]
    rr_chunks    = [[] for _ in other_tasks]

    made_progress = True
    while made_progress:
        made_progress = False
        for i in range(len(other_tasks)):
            if rr_remaining[i] <= 0:
                continue
            for slot in week_slots:
                _, slot_end, cur = slot
                avail_m = int((slot_end - cur).total_seconds() / 60)
                if avail_m < MIN_SLOT_MIN:
                    continue
                place_m = min(rr_remaining[i], MIN_SLOT_MIN, avail_m)
                if place_m < MIN_SLOT_MIN:
                    continue
                start   = cur
                end     = cur + timedelta(minutes=place_m)
                slot[2] = end
                rr_remaining[i] -= place_m
                rr_chunks[i].append((start, end))
                made_progress = True
                break

    for i, task in enumerate(other_tasks):
        if not rr_chunks[i]:
            unscheduled.append(task)
            continue
        if rr_remaining[i] > 0:
            partially_scheduled.append(
                {'task': task, 'remaining_m': rr_remaining[i]}
            )
        _register_chunks(task, rr_chunks[i])

    n_scheduled = len(scheduled)

    # ── サマリー ──
    print('=' * 60)
    print(f'{"[DRY-RUN] " if dry_run else ""}スケジュール: {n_scheduled} 件 / '
          f'部分登録: {len(partially_scheduled)} 件 / '
          f'未スケジュール: {len(unscheduled)} 件')
    if partially_scheduled:
        print('[INFO] 空き不足により一部のみ登録（残り時間あり）:')
        for p in partially_scheduled:
            h = p['remaining_m'] // 60
            m = p['remaining_m'] % 60
            rem_str = f'{h}h{m:02d}m' if h else f'{m}min'
            print(f'  - {p["task"]["title"]}  残 {rem_str}')
    if unscheduled:
        print('[WARN] 空き時間が不足してスケジュールできなかったタスク:')
        for t in unscheduled:
            print(f'  - {t["title"]}')
    if dry_run:
        print()
        print('→ 実際に登録するには --execute を付けて実行してください。')


if __name__ == '__main__':
    main()
