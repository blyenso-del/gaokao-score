import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
import requests, json, os, time, threading
from concurrent.futures import ThreadPoolExecutor

B = 'https://static-data.gaokao.cn'
OUT = r'E:\ai\second\prov_cache'
os.makedirs(OUT, exist_ok=True)
H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
     'Referer': 'https://www.gaokao.cn/'}
YEARS = ['2023', '2024', '2025']
WANT = {'2073', '2074'}  # 物理类 / 历史类

# school_id 列表
with open(r'E:\ai\second\school_code.json', encoding='utf-8') as f:
    scd = json.load(f)['data']
sids = {}
for v in scd.values():
    sid = str(v['school_id'])
    if sid not in sids:
        sids[sid] = v['name']
sid_list = list(sids.keys())

cnt = [0]; hit = [0]; lock = threading.Lock(); t0 = time.time()

def work(sid):
    fp = os.path.join(OUT, sid + '.json')
    if os.path.exists(fp):
        with lock: cnt[0] += 1
        return
    try:
        d = requests.get(B + f'/www/2.0/school/{sid}/info.json', headers=H, timeout=20).json()['data']
    except Exception:
        with lock: cnt[0] += 1
        return
    ptm = d.get('pro_type_min') or {}
    scores = {}  # pid -> {year: {'2073':x,'2074':y}}
    for pid, arr in ptm.items():
        if pid == 'list' or not isinstance(arr, list): continue
        ys = {}
        for item in arr:
            y = str(item.get('year'))
            if y not in YEARS: continue
            tp = item.get('type') or {}
            picked = {k: tp[k] for k in tp if k in WANT}
            if picked: ys[y] = picked
        if ys: scores[pid] = ys
    tags = [x.get('name') for x in (d.get('label_list') or []) if x.get('name')]
    rec = {'sid': sid, 'name': d.get('name') or sids.get(sid),
           'zs_code': d.get('zs_code'), 'level': d.get('level_name'),
           'tags': ','.join(tags), 'scores': scores}
    with open(fp, 'w', encoding='utf-8') as f:
        json.dump(rec, f, ensure_ascii=False)
    with lock:
        if scores: hit[0] += 1
    with lock:
        cnt[0] += 1
        if cnt[0] % 200 == 0:
            print(f'[progress] {cnt[0]}/{len(sid_list)} 有分数={hit[0]} elapsed={int(time.time()-t0)}s', flush=True)

print(f'[start] 下载 {len(sid_list)} 所学校 info.json, 8并发', flush=True)
with ThreadPoolExecutor(max_workers=8) as ex:
    ex.map(work, sid_list)
print(f'[DONE] {cnt[0]} 有分数={hit[0]} elapsed={int(time.time()-t0)}s', flush=True)
