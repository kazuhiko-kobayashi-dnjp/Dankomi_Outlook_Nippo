"""
指定日の TimeTrackerNX 実績を全削除する
Usage: python tt_delete_date.py YYYY-MM-DD
"""
import urllib.request, urllib.error, json, os, sys

base = 'https://ttnx3.sys.globaldenso.com/TimeTrackerNX_05006'
key  = os.environ.get('TT_API_KEY', '')
h    = {'X-TT-ApiKey': key}
uid  = '309'

date = sys.argv[1] if len(sys.argv) > 1 else '2026-06-22'

req = urllib.request.Request(
    f'{base}/api/system/users/{uid}/timeEntries?startDate={date}&finishDate={date}',
    headers=h
)
with urllib.request.urlopen(req, timeout=15) as r:
    data = json.loads(r.read())
entries = data if isinstance(data, list) else data.get('data', [])

print(f'{date} のエントリ数: {len(entries)}')
ok = ng = 0
for e in entries:
    eid   = e['id']
    wname = e.get('workItemName','')
    pcode = e.get('projectCode','')
    memo  = e.get('memo','')[:30]
    dreq  = urllib.request.Request(
        f'{base}/api/system/users/{uid}/timeEntries/{eid}',
        headers=h, method='DELETE'
    )
    try:
        with urllib.request.urlopen(dreq, timeout=15) as r:
            print(f'  DELETE OK  id={eid}  [{pcode}_{wname}]  {memo}')
            ok += 1
    except urllib.error.HTTPError as ex:
        print(f'  DELETE NG  id={eid}  {ex.code}')
        ng += 1

print(f'\n削除完了: {ok} 件成功 / {ng} 件失敗')
