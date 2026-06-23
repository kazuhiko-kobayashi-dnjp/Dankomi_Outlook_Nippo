#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
schedule_tasks.py
進捗管理 Excel から担当タスクを読み取り、Outlook の今週スケジュールの
空き時間帯に自動挿入する。フォーカス時間は削除して空き枠を確保。

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
import argparse
import warnings
from datetime import datetime, timedelta

# ──────────────────────────────────────────────────────────────────────
# 設定
# ──────────────────────────────────────────────────────────────────────
EXCEL_PATH   = (r'C:\Users\10001179776\OneDrive - DENSO\2019\Else\99_temp'
                r'\◎250217_露出制御業務進捗及び報告.xlsx')
SHEET_NAME   = 'Sheet1'
PERSON       = '小林'
WORK_START   = 9    # 09:00
WORK_END     = 18   # 18:00
MIN_SLOT_MIN = 30   # 30 分未満のスロットは使わない
FOCUS_PATTERN = 'フォーカス時間'

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

        tasks.append({'title': first_line, 'duration_h': dur_h})

    wb.close()
    return tasks


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
    day_start = date.replace(hour=WORK_START, minute=0, second=0, microsecond=0)
    day_end   = date.replace(hour=WORK_END,   minute=0, second=0, microsecond=0)

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


# ──────────────────────────────────────────────────────────────────────
# メイン
# ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description='進捗管理 Excel タスクを Outlook に自動スケジュール')
    parser.add_argument('--week', default=None,
                        help='対象週の任意の日 YYMMDD（省略=今週）')
    parser.add_argument('--execute', action='store_true',
                        help='実際に Outlook を操作する（省略=dry-run）')
    parser.add_argument('--file', default=EXCEL_PATH,
                        help='Excel ファイルパス')
    args = parser.parse_args()

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

    print(f'[INFO] モード      : {"dry-run（確認のみ）" if dry_run else "EXECUTE（Outlook 操作）"}')
    print(f'[INFO] 対象週      : {week_start.strftime("%Y-%m-%d")} (月) ～ {week_end.strftime("%Y-%m-%d")} (金)')
    print()

    # ── タスク読み込み ──
    tasks = load_tasks(args.file)
    if not tasks:
        print('[WARN] タスクなし（D列=小林 かつ H列≠空白/済 かつ I列≠0）')
        return

    total_h = sum(t['duration_h'] for t in tasks)
    print(f'[INFO] タスク {len(tasks)} 件（合計 {total_h:.1f}h）')
    for t in tasks:
        h = int(t['duration_h'])
        m = int((t['duration_h'] - h) * 60)
        dur_str = f'{h}h{m:02d}m' if m else f'{h}h'
        print(f'  [{dur_str}] {t["title"]}')
    print()

    # ── Outlook 接続 ──
    app      = win32com.client.Dispatch('Outlook.Application')
    ns       = app.GetNamespace('MAPI')
    calendar = ns.GetDefaultFolder(9)

    week_dates = get_week_dates(week_start)
    task_queue = list(tasks)
    n_scheduled = 0

    for day in week_dates:
        if not task_queue:
            break

        day_label = day.strftime('%m/%d(%a)')
        print(f'── {day_label} ─────────────────────')

        # フォーカス時間を削除（空き枠確保）
        focus_items = collect_focus_events(calendar, day)
        for start_dt, item in focus_items:
            if dry_run:
                print(f'  [削除予定] {start_dt.strftime("%H:%M")} {item.Subject}')
            else:
                print(f'  [削除]     {start_dt.strftime("%H:%M")} {item.Subject}')
                item.Delete()

        # 空きスロット取得（フォーカス時間削除後の状態で計算）
        events = get_calendar_events(calendar, day)
        slots  = find_free_slots(events, day)

        if not slots:
            print('  (空きスロットなし)')
            print()
            continue

        # タスクをスロットに詰める
        for slot_s, slot_e in slots:
            if not task_queue:
                break
            cur = slot_s
            while task_queue and cur < slot_e:
                task  = task_queue[0]
                dur_m = int(task['duration_h'] * 60)
                end   = cur + timedelta(minutes=dur_m)
                if end > slot_e:
                    break   # このスロットに収まらない → 次のスロットへ

                task_queue.pop(0)
                n_scheduled += 1

                if dry_run:
                    print(f'  [登録予定] {cur.strftime("%H:%M")}-{end.strftime("%H:%M")}  {task["title"]}')
                else:
                    appt            = app.CreateItem(1)   # olAppointmentItem
                    appt.Subject    = task['title']
                    appt.Start      = cur.strftime('%Y-%m-%d %H:%M')
                    appt.End        = end.strftime('%Y-%m-%d %H:%M')
                    appt.BusyStatus = 2   # olBusy
                    appt.ReminderSet = False
                    appt.Save()
                    print(f'  [登録済み] {cur.strftime("%H:%M")}-{end.strftime("%H:%M")}  {task["title"]}')

                cur = end

        print()

    # ── サマリー ──
    print('=' * 60)
    print(f'{"[DRY-RUN] " if dry_run else ""}スケジュール: {n_scheduled} 件 / '
          f'未スケジュール: {len(task_queue)} 件')
    if task_queue:
        print('[WARN] 空き時間が不足してスケジュールできなかったタスク:')
        for t in task_queue:
            print(f'  - {t["title"]}')
    if dry_run:
        print()
        print('→ 実際に登録するには --execute を付けて実行してください。')


if __name__ == '__main__':
    main()
