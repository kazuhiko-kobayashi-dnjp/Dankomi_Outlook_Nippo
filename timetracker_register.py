#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
timetracker_register.py
日報の＜業務計画＞セクションをパースして TimeTrackerNX に登録する。

Usage:
    # dry-run（デフォルト）：APIを叩かず内容確認
    python timetracker_register.py
    python timetracker_register.py --date 260622
    python timetracker_register.py "C:/path/to/260622_日報.txt"

    # 実際に登録
    python timetracker_register.py --execute

    # WorkItemId 一覧表示（登録前に config に設定する）
    python timetracker_register.py --list-workitems
    python timetracker_register.py --list-workitems --project-id 5

    # カスタム設定ファイル
    python timetracker_register.py --config my_config.json

認証:
    TimeTracker の「ユーザー設定 → APIキーを生成」で取得したキーを
    環境変数 TT_API_KEY にセットしてください。
    （config の api.api_key_env で変数名を変更可）

    例: $env:TT_API_KEY = "your-api-key-here"
"""

import re
import json
import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

try:
    import requests
    _REQUESTS_AVAILABLE = True
except ImportError:
    _REQUESTS_AVAILABLE = False

try:
    import win32com.client
    _WIN32COM_AVAILABLE = True
except ImportError:
    _WIN32COM_AVAILABLE = False


# ──────────────────────────────────────────────────────────────────────
# 定数
# ──────────────────────────────────────────────────────────────────────
TIME_RE = re.compile(r'^(\d{2}:\d{2})')
ACT_RE  = re.compile(r'＜([^＞]*)＞')


# ──────────────────────────────────────────────────────────────────────
# 設定ファイル
# ──────────────────────────────────────────────────────────────────────
def load_config(config_path: Path) -> dict:
    if not config_path.exists():
        raise FileNotFoundError(f"設定ファイルが見つかりません: {config_path}")
    with open(config_path, encoding='utf-8') as f:
        return json.load(f)


# ──────────────────────────────────────────────────────────────────────
# 日報パーサー
# ──────────────────────────────────────────────────────────────────────
def parse_plan_section(filepath: Path, date) -> list:
    """
    ＜業務計画＞〜＜業務実績＞ 間の時刻エントリを返す。
    Returns: [(datetime, [activity_str, ...]), ...]
    """
    text = filepath.read_text(encoding='utf-8-sig')
    m = re.search(r'＜業務計画＞(.+?)＜業務実績＞', text, re.DOTALL)
    if not m:
        raise ValueError(f"＜業務計画＞/＜業務実績＞ セクションが見つかりません: {filepath}")
    section = m.group(1)

    entries = []
    current_time = None
    current_acts = []

    for line in section.splitlines():
        time_m = TIME_RE.match(line)
        if time_m:
            if current_time is not None:
                entries.append((current_time, current_acts))
            current_time = datetime.combine(
                date, datetime.strptime(time_m.group(1), '%H:%M').time()
            )
            current_acts = [a.strip() for a in ACT_RE.findall(line) if a.strip()]
        elif current_time is not None:
            extras = [a.strip() for a in ACT_RE.findall(line) if a.strip()]
            current_acts.extend(extras)

    if current_time is not None:
        entries.append((current_time, current_acts))
    return entries


# ──────────────────────────────────────────────────────────────────────
# Outlook カレンダーパーサー
# ──────────────────────────────────────────────────────────────────────
def parse_outlook_events(target_date, config: dict) -> list:
    """
    Outlook カレンダーから予定を読み取り、parse_plan_section と同じ形式で返す。

    処理:
    - BusyStatus=0 (Free / Canceled) はスキップ
    - 終日イベントはスキップ
    - 時間が重複する予定は subject を結合してマージ
    - skip_patterns に一致する予定はスキップ

    Returns: [(datetime_start, [activity_str, ...]), ...]
             ※ 最後のエントリに end_time を示すダミーエントリを追加
    """
    if not _WIN32COM_AVAILABLE:
        raise ImportError('pywin32 がインストールされていません: pip install pywin32')

    from datetime import timedelta

    skip_patterns = config.get('skip_patterns', [])

    outlook  = win32com.client.Dispatch('Outlook.Application')
    ns       = outlook.GetNamespace('MAPI')
    calendar = ns.GetDefaultFolder(9)   # olFolderCalendar
    items    = calendar.Items
    items.IncludeRecurrences = True
    items.Sort('[Start]')

    d_str  = target_date.strftime('%Y/%m/%d')
    d_next = (target_date + timedelta(days=1)).strftime('%Y/%m/%d')
    flt    = f"[Start] >= '{d_str} 00:00' AND [Start] < '{d_next} 00:00'"
    found  = items.Restrict(flt)

    # 生イベントを収集
    raw_events = []
    for item in found:
        try:
            subject   = (item.Subject or '').strip()
            busy      = item.BusyStatus   # 0=Free,1=Tentative,2=Busy,3=OOO,4=WFH
            is_allday = item.AllDayEvent
            if is_allday or busy == 0:   # 終日イベント・Free(Canceled) はスキップ
                continue
            if should_skip(subject, skip_patterns):
                continue
            st = item.Start
            en = item.End
            start_dt = datetime(st.year, st.month, st.day, st.hour, st.minute)
            end_dt   = datetime(en.year, en.month, en.day, en.hour, en.minute)
            if end_dt <= start_dt:
                continue
            raw_events.append({'subject': subject, 'start': start_dt, 'end': end_dt, 'busy': busy})
        except Exception:
            continue

    raw_events.sort(key=lambda x: x['start'])

    if not raw_events:
        return []

    # 同一開始時刻のイベントのみマージ（並行会議の subject を結合）
    # ※ 時間帯が重なるだけの別イベントは独立エントリとして保持する
    merged = []
    for ev in raw_events:
        if merged and ev['start'] == merged[-1]['start']:
            # 全く同じ開始時刻 → subject を追加、end を延長
            prev = merged[-1]
            if ev['subject'] not in prev['subjects']:
                prev['subjects'].append(ev['subject'])
            prev['end'] = max(prev['end'], ev['end'])
        else:
            merged.append({'start': ev['start'], 'end': ev['end'], 'subjects': [ev['subject']]})

    # [{'start': datetime, 'end': datetime, 'subjects': [str]}, ...] をそのまま返す
    return merged


# ──────────────────────────────────────────────────────────────────────
# 分類・エントリ構築
# ──────────────────────────────────────────────────────────────────────
def build_entries_outlook(events: list, config: dict) -> list:
    """
    Outlook イベントリスト [{start, end, subjects}, ...] から登録エントリを構築。
    kyuka_patterns にマッチするイベントは JPRF0208_有休 として登録し、
    その時間帯と被った他のイベントはスキップする。
    """
    skip_patterns  = config.get('skip_patterns', [])
    kyuka_patterns = config.get('kyuka_patterns', [])

    def is_kyuka(subjects):
        return bool(kyuka_patterns) and any(
            any(re.search(p, s) for p in kyuka_patterns)
            for s in subjects
        )

    # Step 1: 有休イベントを特定
    kyuka_blocks = [blk for blk in events if is_kyuka(blk['subjects'])]

    def overlaps_kyuka(blk):
        return any(
            blk['start'] < k['end'] and blk['end'] > k['start']
            for k in kyuka_blocks
        )

    result = []
    for blk in sorted(events, key=lambda x: x['start']):
        subjects = blk['subjects']
        start = blk['start']
        end   = blk['end']
        dur_h = (end - start).total_seconds() / 3600.0
        if dur_h <= 0:
            continue

        if is_kyuka(subjects):
            # 有休イベント → JPRF0208_有休 として登録
            wid = config.get('workitem_map', {}).get('JPRF0208', {}).get('有休', '')
            result.append({
                'startTime':  start,
                'finishTime': end,
                'project':    'JPRF0208',
                'worktype':   '有休',
                'category':   'JPRF0208_有休',
                'workItemId': wid,
                'memo':       ' / '.join(subjects),
                'hours_raw':  round(dur_h, 2),
            })
        elif kyuka_blocks and overlaps_kyuka(blk):
            # 有休時間帯と被る → スキップ
            continue
        else:
            meaningful = [s for s in subjects if s and not should_skip(s, skip_patterns)]
            if not meaningful:
                continue
            project, worktype, wid, label = classify_entry(meaningful, config)
            result.append({
                'startTime':  start,
                'finishTime': end,
                'project':    project,
                'worktype':   worktype,
                'category':   label,
                'workItemId': wid,
                'memo':       ' / '.join(meaningful),
                'hours_raw':  round(dur_h, 2),
            })
    return result


def should_skip(activity: str, skip_patterns: list) -> bool:
    return any(re.search(p, activity) for p in skip_patterns)


def classify_entry(acts: list, config: dict):
    """
    2段階分類: プロジェクト決定 → 作業種別決定 → workItemId 解決

    Returns:
        (project_code, worktype, workitem_id, display_label)
    """
    combined = ' '.join(acts)

    # ── Step 1: プロジェクト決定 ──
    project = None
    for rule in config.get('project_rules', []):
        if re.search(rule['pattern'], combined):
            project = rule['project']
            break
    if project is None:
        project = config.get('default_project', 'JPRF0208')

    # ── Step 2: 作業種別決定 ──
    worktype = None
    for rule in config.get('worktype_rules', []):
        if re.search(rule['pattern'], combined):
            worktype = rule['worktype']
            break
    if worktype is None:
        worktype = config.get('default_worktype', '設計業務')

    # ── Step 3: workItemId 解決 ──
    proj_map  = config.get('workitem_map', {}).get(project, {})
    wid = proj_map.get(worktype)
    if not wid:
        # フォールバック: fallback_worktype（デフォルト"その他"）
        fallback = config.get('fallback_worktype', 'その他')
        wid = proj_map.get(fallback, '')

    label = f"{project}_{worktype}"
    return project, worktype, wid or '', label


def build_entries(raw_entries: list, config: dict) -> list:
    """
    空スロットを直前有効エントリに吸収し、startTime/finishTime 付きエントリを返す。

    Returns:
        list of {
            "startTime":   datetime,
            "finishTime":  datetime,
            "project":     str,
            "worktype":    str,
            "category":    str,   # 表示用ラベル (project_worktype)
            "workItemId":  str,
            "memo":        str,
            "hours_raw":   float,
        }
    """
    skip_patterns = config.get('skip_patterns', [])

    # ステップ1: 全スロットの startTime/finishTime を計算
    raw = []
    for i, (start, acts) in enumerate(raw_entries):
        end   = raw_entries[i + 1][0] if i + 1 < len(raw_entries) else start
        dur_h = (end - start).total_seconds() / 3600.0
        raw.append([start, end, dur_h, acts])

    # ステップ2: 空スロットを直前有効エントリに吸収
    effective = []
    for start, end, dur_h, acts in raw:
        meaningful = [a for a in acts if a and not should_skip(a, skip_patterns)]
        if meaningful:
            effective.append([start, end, dur_h, meaningful])
        else:
            if effective:
                prev = effective[-1]
                prev[2] += dur_h
                prev[1] = end

    # ステップ3: 分類・workItemId 解決
    result = []
    for start, end, dur_h, acts in effective:
        if dur_h <= 0:
            continue
        project, worktype, wid, label = classify_entry(acts, config)

        result.append({
            'startTime':  start,
            'finishTime': end,
            'project':    project,
            'worktype':   worktype,
            'category':   label,
            'workItemId': wid,
            'memo':       ' / '.join(acts),
            'hours_raw':  round(dur_h, 2),
        })
    return result


# ──────────────────────────────────────────────────────────────────────
# TimeTrackerNX API ヘルパー
# ──────────────────────────────────────────────────────────────────────
def _check_requests():
    if not _REQUESTS_AVAILABLE:
        print('[ERROR] requests がインストールされていません: pip install requests')
        sys.exit(1)


def make_headers(config: dict) -> dict:
    """X-TT-ApiKey ヘッダを返す。"""
    api_cfg     = config.get('api', {})
    env_name    = api_cfg.get('api_key_env', 'TT_API_KEY')
    api_key     = os.environ.get(env_name, '')
    if not api_key:
        print(f'[ERROR] 環境変数 {env_name} が未設定です。')
        print(f'        TimeTracker の「ユーザー設定 → APIキーを生成」で取得し、')
        print(f'        PowerShell で  $env:{env_name} = "your-key"  とセットしてください。')
        sys.exit(1)
    return {
        'X-TT-ApiKey': api_key,
        'Content-Type': 'application/json',
    }


def api_request(method: str, url: str, headers: dict, json_body=None, timeout=30):
    """GET/POST/PUT/DELETE の薄いラッパー。エラー時は例外を raise する。"""
    resp = requests.request(method, url, headers=headers, json=json_body, timeout=timeout)
    if not resp.ok:
        try:
            detail = resp.json()
        except Exception:
            detail = resp.text
        raise RuntimeError(f"HTTP {resp.status_code}: {detail}")
    return resp


def get_my_user_id(base_url: str, headers: dict) -> str:
    """GET /system/users/me → ユーザーID を返す。"""
    resp = api_request('GET', f'{base_url}/api/system/users/me', headers)
    data = resp.json()
    # レスポンスは User オブジェクト（または配列）
    if isinstance(data, list):
        data = data[0]
    return str(data['id'])


def list_workitems(base_url: str, headers: dict, project_id: str = None):
    """
    ワークアイテム一覧を返す。
    project_id 指定時: GET /project/projects/{id}/... で絞り込む。
    未指定時: GET /workitem/workItems（全件 or 上限あり）
    """
    if project_id:
        url = f'{base_url}/api/workitem/workItems/{project_id}/subItems'
    else:
        url = f'{base_url}/api/workitem/workItems'
    resp = api_request('GET', url, headers)
    return resp.json()


# ──────────────────────────────────────────────────────────────────────
# 登録処理
# ──────────────────────────────────────────────────────────────────────
def register_entries(
    user_id: str,
    entries: list,
    base_url: str,
    headers: dict,
    dry_run: bool,
) -> None:
    """
    エントリ一覧を TimeTrackerNX に登録する。
    dry_run=True の場合は POST せず内容を表示するだけ。
    """
    endpoint = f'{base_url}/api/system/users/{user_id}/timeEntries'

    ok_count  = 0
    err_count = 0

    for entry in entries:
        payload = {
            'workItemId': entry['workItemId'],
            'startTime':  entry['startTime'].strftime('%Y-%m-%dT%H:%M:%S'),
            'finishTime': entry['finishTime'].strftime('%Y-%m-%dT%H:%M:%S'),
            'memo':       entry['memo'],
        }

        start_s = entry['startTime'].strftime('%H:%M')
        end_s   = entry['finishTime'].strftime('%H:%M')
        label   = f"{start_s}-{end_s} [{entry['category']}] {entry['memo']}"

        if not entry['workItemId']:
            print(f'[SKIP]    workItemId 未設定 → {label}')
            continue

        if dry_run:
            print(f'[DRY-RUN] POST {endpoint}')
            print(f'          {json.dumps(payload, ensure_ascii=False)}')
        else:
            try:
                resp_data = api_request('POST', endpoint, headers, json_body=payload).json()
                print(f'[OK]      id={resp_data.get("id", "?")}  {label}')
                ok_count += 1
            except Exception as exc:
                print(f'[ERROR]   {label}')
                print(f'          {exc}')
                err_count += 1

    if not dry_run:
        print(f'\n登録完了: {ok_count} 件成功 / {err_count} 件失敗')


# ──────────────────────────────────────────────────────────────────────
# エントリポイント
# ──────────────────────────────────────────────────────────────────────
def find_diary(date_str: str, base_dir: Path) -> Path:
    year = f'20{date_str[:2]}'
    return base_dir / year / f'{date_str}_日報.txt'


def main():
    parser = argparse.ArgumentParser(
        description='日報＜業務計画＞→ TimeTrackerNX 登録スクリプト',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        'file', nargs='?',
        help='日報ファイルパス（省略時: --date または本日の日報を自動検索）'
    )
    parser.add_argument(
        '--date', metavar='YYMMDD',
        help='対象日 (例: 260622)。file 省略時に使用。省略時は本日。'
    )
    parser.add_argument(
        '--execute', action='store_true',
        help='実際に API を叩く（デフォルトは dry-run）'
    )
    parser.add_argument(
        '--list-workitems', action='store_true',
        help='ワークアイテム一覧を表示して終了（config の category_workitems を設定するために使用）'
    )
    parser.add_argument(
        '--project-id', metavar='ID',
        help='--list-workitems 時にプロジェクト ID で絞り込む'
    )
    parser.add_argument(
        '--config', default=None,
        help='設定ファイルパス（省略: スクリプトと同じディレクトリの timetracker_config.json）'
    )
    parser.add_argument(
        '--source', choices=['diary', 'outlook'], default='outlook',
        help='予定の取得元: outlook=Outlookカレンダー（デフォルト）/ diary=日報ファイル'
    )
    args = parser.parse_args()

    dry_run    = not args.execute
    script_dir = Path(__file__).parent
    config_path = (
        Path(args.config) if args.config
        else script_dir / 'timetracker_config.json'
    )
    config   = load_config(config_path)
    api_cfg  = config.get('api', {})
    base_url = api_cfg.get('base_url', '').rstrip('/')

    if not base_url and not dry_run:
        print('[ERROR] api.base_url が未設定です。timetracker_config.json を確認してください。')
        print('        例: "base_url": "http://ttserver01/TimeTrackerNX"')
        sys.exit(1)
    if not base_url:
        print('[WARN] api.base_url 未設定 — dry-run のみ実行可能です。')
        print()

    # ── ワークアイテム一覧表示モード ──
    if args.list_workitems:
        _check_requests()
        headers = make_headers(config)
        items = list_workitems(base_url, headers, project_id=args.project_id)
        if isinstance(items, list):
            print(f'{"ID":<10}  {"名前"}')
            print('─' * 60)
            for item in items:
                print(f'{item.get("id", ""):<10}  {item.get("name", "")}')
        else:
            print(json.dumps(items, ensure_ascii=False, indent=2))
        return

    # ── 日付・ファイルパス決定 ──
    if args.file:
        diary_path = Path(args.file)
        date_str   = diary_path.stem[:6]
    else:
        date_str = args.date or datetime.today().strftime('%y%m%d')
        base_dir = Path(config.get('diary_base_dir', script_dir))
        diary_path = find_diary(date_str, base_dir)

    try:
        date = datetime.strptime(date_str, '%y%m%d').date()
    except ValueError:
        print(f'[ERROR] 日付フォーマット不正: {date_str}  (YYMMDD 形式)')
        sys.exit(1)

    # ── ソース決定 ──
    use_outlook = (args.source == 'outlook')

    if use_outlook:
        print(f'[INFO] ソース       : Outlook カレンダー')
    else:
        if not diary_path.exists():
            print(f'[ERROR] 日報ファイルが見つかりません: {diary_path}')
            sys.exit(1)
        print(f'[INFO] 対象ファイル : {diary_path}')

    print(f'[INFO] 対象日       : {date}')
    print(f'[INFO] モード       : {"dry-run（確認のみ）" if dry_run else "EXECUTE（API登録）"}')
    print()

    # ── パース・分類 ──
    if use_outlook:
        outlook_events = parse_outlook_events(date, config)
        entries = build_entries_outlook(outlook_events, config)
    else:
        raw_entries = parse_plan_section(diary_path, date)
        entries = build_entries(raw_entries, config)

    # ── 結果表示 ──
    col_cat = 18
    print(f"{'カテゴリ':<{col_cat}}  {'WorkItemID':<12}  {'時間':>5}  内容")
    print('─' * 80)
    for e in entries:
        wid = e['workItemId'] or '(未設定)'
        h   = e['hours_raw']
        print(
            f"{e['category']:<{col_cat}}  {wid:<12}  {h:>4.1f}h  "
            f"{e['startTime'].strftime('%H:%M')}-{e['finishTime'].strftime('%H:%M')} {e['memo']}"
        )
    total_h = sum(e['hours_raw'] for e in entries)
    print('─' * 80)
    print(f"{'合計':<{col_cat}}  {'':12}  {total_h:>4.1f}h")
    print()

    # workItemId 未設定カテゴリの警告
    missing = {e['category'] for e in entries if not e['workItemId']}
    if missing:
        print('[WARN] 以下のカテゴリに workItemId が未設定です。')
        print('       --list-workitems で ID を確認し、timetracker_config.json に追加してください。')
        for cat in sorted(missing):
            print(f'         - {cat}')
        print()

    # ── API 登録 ──
    if not dry_run:
        _check_requests()
    headers = make_headers(config) if not dry_run else {}
    user_id = get_my_user_id(base_url, headers) if not dry_run else '<dry-run>'

    register_entries(user_id, entries, base_url, headers, dry_run=dry_run)


if __name__ == '__main__':
    main()
