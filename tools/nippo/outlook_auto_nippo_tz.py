# -*- coding: cp932 -*-
import win32com.client
from datetime import datetime, timedelta
import sys

"""
タイムゾーン対応版: 既存の機能は保持しつつ、Outlook から取得した開始/終了時刻をローカルタイムに変換して
ログ表示とスロットマッチングに用います。
使い方は既存スクリプトと同じ:
  python outlook_auto_nippo_tz.py [yymmdd] [noempty]
  または
  python outlook_auto_nippo_tz.py [yyyy] [m] [d] [noempty]
"""


def get_outlook_schedule(date):
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    calendar = namespace.GetDefaultFolder(9)  # 9: olFolderCalendar
    print(f"[DEBUG] calendar name: {calendar.Name}")
    try:
        print(f"[DEBUG] account name: {calendar.Parent.Name}")
    except Exception:
        pass

    items = calendar.Items
    items.IncludeRecurrences = True
    items.Sort("Start")
    print(f"[DEBUG] items.Count: {items.Count}")

    # 指定日の予定のみ取得
    start = datetime(date.year, date.month, date.day, 0, 0, 0)
    end = start + timedelta(days=1)
    restrict_str = f"[Start] >= '{start.strftime('%Y/%m/%d %H:%M')}' AND [Start] < '{end.strftime('%Y/%m/%d %H:%M')}'"
    items = items.Restrict(restrict_str)

    count = 0
    schedule = []
    if getattr(items, 'Count', 0) == 0:
        print("[DEBUG] Restrict returned 0 items。アイテムの走査を続行します。")

    for item in items:
        if item is None:
            continue
        start_raw = item.Start
        end_raw = item.End
        # ローカルタイムに変換して tzinfo を剥がす（比較を単純にするため）
        start_local = start_raw
        end_local = end_raw
        try:
            if hasattr(start_raw, 'tzinfo') and start_raw.tzinfo is not None:
                start_local = start_raw.astimezone().replace(tzinfo=None)
            if hasattr(end_raw, 'tzinfo') and end_raw.tzinfo is not None:
                end_local = end_raw.astimezone().replace(tzinfo=None)
        except Exception:
            # 変換失敗でも元データを使って続行
            start_local = start_raw
            end_local = end_raw

        # 既存のスロットマッチングを壊さないため、schedule に格納する start/end は
        # tz 情報を剥がすのみ（日時そのものは変えない）。この値が create_nippou の比較に使われる。
        start_naive = start_raw
        end_naive = end_raw
        try:
            if hasattr(start_raw, 'tzinfo') and start_raw.tzinfo is not None:
                start_naive = start_raw.replace(tzinfo=None)
            if hasattr(end_raw, 'tzinfo') and end_raw.tzinfo is not None:
                end_naive = end_raw.replace(tzinfo=None)
        except Exception:
            start_naive = start_raw
            end_naive = end_raw

        entry = {
            # 互換性のため、'start'/'end' は naive の日時（数値は変えない）を格納
            "start": start_naive,
            "end": end_naive,
            # デバッグ用に raw と local を保持
            "start_raw": start_raw,
            "end_raw": end_raw,
            "start_local": start_local,
            "end_local": end_local,
            "subject": item.Subject
        }
        # ログには raw と local の両方を表示してトラブルシュートしやすくする
        try:
            print(f"[LOG] 予定: {entry['start_raw']} (raw) -> {entry['start_local']} (local) ～ {entry['end_raw']} (raw) -> {entry['end_local']} (local) | 件名: {entry['subject']}")
        except Exception:
            # 出力で失敗しても処理は続行
            print(f"[LOG] 予定: {entry['start']} ～ {entry['end']} | 件名: {entry['subject']}")

        schedule.append(entry)
        count += 1

    print(f"[LOG] 指定日の予定件数: {count}")
    return schedule


def create_nippou(date, schedule):
    lines = []
    lines.append(f"{date.strftime('%y%m%d')}_日報\n")
    lines.append("＜特記事項＞")
    lines.append("    ・")
    lines.append("    ・")
    lines.append("＜業務計画＞")
    time_slots = [
        "04:20", "05:25", "06:15", "06:30", "07:00", "07:30", "08:00", "08:30",
        "09:00", "09:30", "10:00", "10:15", "10:30", "11:00", "11:30",
        "12:15 　～昼休憩～", "13:00", "13:15", "13:30", "14:00", "14:30",
        "15:00", "15:30", "16:00", "16:30", "17:00", "17:30", "18:00", "18:30", "19:30", "20:00　業務終了"
    ]
    noempty = False
    if 'noempty' in sys.argv:
        noempty = True

    # 準備: スロットの日時リストを作る
    slots = []
    for slot in time_slots:
        slot_hour, slot_minute = map(int, slot[:5].split(":"))
        slot_time = datetime(date.year, date.month, date.day, slot_hour, slot_minute)
        slots.append((slot, slot_time))

    # イベントを時刻でソート（念のため）
    events = list(schedule)
    try:
        events.sort(key=lambda e: e.get('start') if isinstance(e.get('start'), datetime) else datetime.max)
    except Exception:
        pass

    assigned = set()  # イベントのインデックスをマーク
    inserted_events = []  # マージ時に差し込んだ未割当イベントを記録

    # まずスロットごとにマッチするイベントを判定して assigned を決め、matched_subjects_by_slot を作る
    matched_subjects_by_slot = {}
    for si, (slot_label, slot_time) in enumerate(slots):
        matched_subjects = []
        matched_indices = []
        for ei, event in enumerate(events):
            event_start = event.get('start')
            if isinstance(event_start, str):
                try:
                    event_start = datetime.strptime(event_start, "%m/%d/%Y %H:%M:%S")
                except Exception:
                    try:
                        event_start = datetime.strptime(event_start, "%Y-%m-%d %H:%M:%S")
                    except Exception:
                        continue
            try:
                if abs((event_start - slot_time).total_seconds()) < 60*15:
                    matched_subjects.append(event.get('subject'))
                    matched_indices.append(ei)
            except Exception:
                continue
        matched_subjects_by_slot[slot_time] = (slot_label, matched_subjects, matched_indices)
        for idx in matched_indices:
            assigned.add(idx)

    # 未割当イベントを収集
    unmatched_events = []
    for ei, event in enumerate(events):
        if ei in assigned:
            continue
        event_start = event.get('start')
        if isinstance(event_start, str):
            try:
                event_start = datetime.strptime(event_start, "%m/%d/%Y %H:%M:%S")
            except Exception:
                try:
                    event_start = datetime.strptime(event_start, "%Y-%m-%d %H:%M:%S")
                except Exception:
                    continue
        unmatched_events.append((ei, event_start, event.get('subject')))

    # スロット順に出力し、各スロットの後にその区間に含まれる未割当イベントを挿入する
    out_lines = []
    for i, (slot_label, slot_time) in enumerate(slots):
        slot_label, matched_subjects, _ = matched_subjects_by_slot.get(slot_time, (slot_label, [], []))
        if matched_subjects:
            out_lines.append(f"{slot_label}\t＜{matched_subjects[0]}＞")
            for sub in matched_subjects[1:]:
                out_lines.append(f"\t\t＜{sub}＞")
        elif not noempty:
            out_lines.append(f"{slot_label}\t＜＞")

        # このスロットの区間に含まれる未割当イベントを探して挿入
        next_time = slots[i+1][1] if i+1 < len(slots) else None
        to_insert = []
        for ei, estart, subj in list(unmatched_events):
            if estart is None:
                continue
            if next_time is None:
                cond = estart >= slot_time
            else:
                cond = slot_time <= estart < next_time
            if cond:
                to_insert.append((ei, estart, subj))
                # 挿入と同時に記録
                inserted_events.append((ei, estart, subj))
                unmatched_events = [u for u in unmatched_events if u[0] != ei]
        to_insert.sort(key=lambda x: x[1])
        for ei, estart, subj in to_insert:
            timestr = estart.strftime('%H:%M') if isinstance(estart, datetime) else ''
            out_lines.append(f"{timestr}\t＜{subj}＞")

    # もし slots が空であれば、未割当イベントのみを時系列で出す
    if not slots:
        for ei, estart, subj in sorted(unmatched_events, key=lambda x: x[1] or datetime.max):
            timestr = estart.strftime('%H:%M') if isinstance(estart, datetime) else ''
            out_lines.append(f"{timestr}\t＜{subj}＞")

    # 出力 lines にマージ結果を追加
    for l in out_lines:
        lines.append(l)

    # 実績欄を業務計画から転記（従来の振る舞いを維持）。これに加え、挿入したイベントも該当区間に追記する
    lines.append("\n＜業務実績＞")
    for i, (slot_label, slot_time) in enumerate(slots):
        matched_subjects = matched_subjects_by_slot.get(slot_time, (slot_label, [], []))[1]
        if matched_subjects:
            entry = f"{slot_label}\t＜{matched_subjects[0]}＞"
            lines.append(entry)
            for sub in matched_subjects[1:]:
                lines.append(f"\t\t＜{sub}＞")
            lines.append("\t\t・")
            lines.append("\t\t・")
        elif not noempty:
            entry = f"{slot_label}\t＜＞"
            lines.append(entry)
            lines.append("\t\t・")
            lines.append("\t\t・")

        # ここでも未割当だった挿入イベントを同じ区間に追記
        next_time = slots[i+1][1] if i+1 < len(slots) else None
        to_insert_perf = [e for e in inserted_events if (e[1] is not None and ((next_time is None and e[1] >= slot_time) or (next_time is not None and slot_time <= e[1] < next_time)))]
        to_insert_perf.sort(key=lambda x: x[1])
        for ei, estart, subj in to_insert_perf:
            timestr = estart.strftime('%H:%M') if isinstance(estart, datetime) else ''
            lines.append(f"{timestr}\t＜{subj}＞")
            lines.append("\t\t・")
            lines.append("\t\t・")

    return "\n".join(lines)


if __name__ == "__main__":
    print(f"[DEBUG] sys.argv: {sys.argv}")
    if any(arg.lower() in ["help", "helpu", "-h", "--help"] for arg in sys.argv):
        print("使い方: python outlook_auto_nippo_tz.py [yymmdd] [noempty]  または  python outlook_auto_nippo_tz.py [yyyy] [m] [d] [noempty]")
        print("例: python outlook_auto_nippo_tz.py 250729 noempty")
        sys.exit(0)

    if len(sys.argv) >= 2 and len(sys.argv[1]) == 6 and sys.argv[1].isdigit():
        ymd = sys.argv[1]
        year = 2000 + int(ymd[:2])
        month = int(ymd[2:4])
        day = int(ymd[4:6])
        target_date = datetime(year, month, day)
    elif len(sys.argv) == 4:
        year = int(sys.argv[1])
        month = int(sys.argv[2])
        day = int(sys.argv[3])
        target_date = datetime(year, month, day)
    else:
        target_date = datetime(2025, 7, 28)

    print(f"[DEBUG] target_date: {target_date}")
    schedule = get_outlook_schedule(target_date)
    nippou_text = create_nippou(target_date, schedule)
    filename = f"{target_date.strftime('%y%m%d')}_日報.txt"
    with open(filename, "w", encoding="utf-8") as f:
        f.write(nippou_text)
