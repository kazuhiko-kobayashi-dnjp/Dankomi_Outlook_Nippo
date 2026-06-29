import os
import win32com.client
from datetime import datetime, timedelta
from pathlib import Path

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

    # 指定日の前後を含めて取得し、後でローカルタイムでフィルタリングする
    # (RestrictがUTCで動作する場合に朝の予定が漏れるのを防ぐため)
    search_start = date - timedelta(days=1)
    search_end = date + timedelta(days=2)
    restrict_str = f"[Start] >= '{search_start.strftime('%Y/%m/%d %H:%M')}' AND [Start] < '{search_end.strftime('%Y/%m/%d %H:%M')}'"
    items = items.Restrict(restrict_str)

    count = 0
    schedule = []
    if getattr(items, 'Count', 0) == 0:
        print("[DEBUG] Restrict returned 0 items。")

    target_date_only = date.date()

    for item in items:
        if item is None:
            continue
        
        start_raw = item.Start
        end_raw = item.End
        
        # タイムゾーン変換を行わず、数値上の時刻をそのまま使用する (Numerical JST as UTC)
        # 環境によって(OutlookのMAPI経由など)JSTの時刻がUTCとして返ってくるケースがあるため
        start_local_dt = start_raw.replace(tzinfo=None) if hasattr(start_raw, 'tzinfo') else start_raw
        end_local_dt = end_raw.replace(tzinfo=None) if hasattr(end_raw, 'tzinfo') else end_raw

        # 指定日の予定かチェック
        try:
            if start_local_dt.date() != target_date_only:
                continue
        except Exception:
            continue

        entry = {
            "start": start_local_dt,
            "end": end_local_dt,
            "subject": item.Subject,
            "body": (item.Body or '').strip()
        }
        print(f"[LOG] 予定: {start_local_dt} ～ {end_local_dt} | 件名: {entry['subject']}")
        schedule.append(entry)
        count += 1
    print(f"[LOG] 指定日の予定件数: {count}")
    return schedule


def _append_body_detail(lines: list, body: str) -> None:
    """[auto-sched]ボディの詳細行（E列/F列情報）を日報実績欄に追記する。"""
    if not body or '[auto-sched]' not in body:
        return
    past_marker = False
    for line in body.split('\n'):
        if '[auto-sched]' in line:
            past_marker = True
            continue
        if past_marker and line.strip():
            lines.append(f"\t\t{line.strip()}")


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
    import sys
    if 'noempty' in sys.argv:
        noempty = True

    # 準備: スロットの日時リストを作る
    slots = []
    for slot in time_slots:
        slot_hour, slot_minute = map(int, slot[:5].split(":"))
        slot_time = datetime(date.year, date.month, date.day, slot_hour, slot_minute)
        slots.append((slot, slot_time))

    # イベントを時刻でソート
    events = list(schedule)
    try:
        events.sort(key=lambda e: e.get('start') if isinstance(e.get('start'), datetime) else datetime.max)
    except Exception:
        pass

    assigned = set()
    inserted_events = []

    # スロットごとにマッチ判定して assigned を決める
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
                    matched_subjects.append((event.get('subject', ''), event.get('body', '')))
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

    # スロット順に出力し、各スロットの後に区間内の未割当イベントを挿入
    out_lines = []
    
    # 【追加】最初のスロットより前のイベント
    if slots:
        first_slot_time = slots[0][1]
        to_insert_early = []
        for ei, estart, subj in list(unmatched_events):
            if estart < first_slot_time:
                to_insert_early.append((ei, estart, subj))
                inserted_events.append((ei, estart, subj))
                unmatched_events = [u for u in unmatched_events if u[0] != ei]
        to_insert_early.sort(key=lambda x: x[1])
        for ei, estart, subj in to_insert_early:
            timestr = estart.strftime('%H:%M')
            out_lines.append(f"{timestr}\t＜{subj}＞")

    for i, (slot_label, slot_time) in enumerate(slots):
        slot_label, matched_subjects, _ = matched_subjects_by_slot.get(slot_time, (slot_label, [], []))
        if matched_subjects:
            out_lines.append(f"{slot_label}\t＜{matched_subjects[0][0]}＞")
            for sub, _b in matched_subjects[1:]:
                out_lines.append(f"\t\t＜{sub}＞")
        elif not noempty:
            out_lines.append(f"{slot_label}\t＜＞")

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
                inserted_events.append((ei, estart, subj))
                unmatched_events = [u for u in unmatched_events if u[0] != ei]
        to_insert.sort(key=lambda x: x[1])
        for ei, estart, subj in to_insert:
            timestr = estart.strftime('%H:%M') if isinstance(estart, datetime) else ''
            out_lines.append(f"{timestr}\t＜{subj}＞")

    if not slots:
        for ei, estart, subj in sorted(unmatched_events, key=lambda x: x[1] or datetime.max):
            timestr = estart.strftime('%H:%M') if isinstance(estart, datetime) else ''
            out_lines.append(f"{timestr}\t＜{subj}＞")

    for l in out_lines:
        lines.append(l)

    # 実績欄: スロット毎に従来の記載を行い、同区間の挿入イベントも追記
    lines.append("\n＜業務実績＞")
    
    # 【追加】最初のスロットより前の実績
    if slots:
        first_slot_time = slots[0][1]
        to_insert_early_perf = [e for e in inserted_events if e[1] < first_slot_time]
        to_insert_early_perf.sort(key=lambda x: x[1])
        for ei, estart, subj in to_insert_early_perf:
            timestr = estart.strftime('%H:%M')
            lines.append(f"{timestr}\t＜{subj}＞")
            lines.append("\t\t・")
            lines.append("\t\t・")

    for i, (slot_label, slot_time) in enumerate(slots):
        matched_subjects = matched_subjects_by_slot.get(slot_time, (slot_label, [], []))[1]
        if matched_subjects:
            entry = f"{slot_label}\t＜{matched_subjects[0][0]}＞"
            lines.append(entry)
            for sub, _b in matched_subjects[1:]:
                lines.append(f"\t\t＜{sub}＞")
            _append_body_detail(lines, matched_subjects[0][1])
            lines.append("\t\t・")
            lines.append("\t\t・")
        elif not noempty:
            entry = f"{slot_label}\t＜＞"
            lines.append(entry)
            lines.append("\t\t・")
            lines.append("\t\t・")

        next_time = slots[i+1][1] if i+1 < len(slots) else None
        to_insert_perf = [e for e in inserted_events if (e[1] is not None and ((next_time is None and e[1] >= slot_time) or (next_time is not None and slot_time <= e[1] < next_time)))]
        to_insert_perf.sort(key=lambda x: x[1])
        for ei, estart, subj in to_insert_perf:
            timestr = estart.strftime('%H:%M') if isinstance(estart, datetime) else ''
            lines.append(f"{timestr}\t＜{subj}＞")
            lines.append("\t\t・")
            lines.append("\t\t・")

    return "\n".join(lines)
import sys
if __name__ == "__main__":

    print(f"[DEBUG] sys.argv: {sys.argv}")
    # ヘルプ表示
    if any(arg.lower() in ["help", "helpu", "-h", "--help"] for arg in sys.argv):
        print("使い方: python outlook_auto_nippo.py [yymmdd] [noempty]  または  python outlook_auto_nippo.py [yyyy] [m] [d] [noempty]")
        print("例: python outlook_auto_nippo.py 250729 noempty")
        print("例: python outlook_auto_nippo.py 2025 7 29 noempty")
        print("[noempty] オプションで空白行を省略")
        sys.exit(0)
    # sys.argv[1]が6桁数字なら日付として扱う（オプションがあってもOK）
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
        # デフォルトは今日
        target_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    print(f"[DEBUG] target_date: {target_date}")
    schedule = get_outlook_schedule(target_date)
    nippou_text = create_nippou(target_date, schedule)
    # Save to 日報/20YY/ subfolder (workspace root = 2 levels up from tools/nippo/)
    workspace_root = Path(__file__).resolve().parent.parent.parent
    year_str = target_date.strftime('%Y')
    nippo_dir = workspace_root / "日報" / year_str
    nippo_dir.mkdir(parents=True, exist_ok=True)
    filename = nippo_dir / f"{target_date.strftime('%y%m%d')}_日報.txt"
    if filename.exists():
        backup_name = nippo_dir / f"{target_date.strftime('%y%m%d')}_日報_bak_{datetime.now().strftime('%H%M%S')}.txt"
        filename.rename(backup_name)
        print(f"[INFO] 既存の日報をバックアップしました: {backup_name}")
    with open(filename, "w", encoding="utf-8-sig") as f:
        f.write(nippou_text)
    print(f"[INFO] 保存先: {filename}")