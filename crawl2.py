# -*- coding: utf-8 -*-
"""全国各省·专业级录取数据爬虫 (schoolprovinceindex 全字段)。
用法:
  python crawl2.py 44      只爬广东(area=44), 单省测试
  python crawl2.py all     爬全部新高考省份
输出: sp_full/{area}.json
"""
import sys, io, json, os, time, threading
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
import requests
from concurrent.futures import ThreadPoolExecutor

BASE = r'E:\ai\second'
SD = 'https://static-data.gaokao.cn'
OUT = os.path.join(BASE, 'sp_full')
os.makedirs(OUT, exist_ok=True)
H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)',
     'Referer': 'https://www.gaokao.cn/'}
YEARS = ['2025', '2024', '2023', '2022']
KEEP_TYPES = {'3', '2073', '2074'}  # 综合改革 / 物理类 / 历史类

def get_json(url, tries=6):
    for _ in range(tries):
        try:
            r = requests.get(url, headers=H, timeout=20)
            if r.status_code == 200:
                return r.json()
        except Exception:
            pass
        time.sleep(0.5)
    return None

def load_prov_types():
    for sid in ['140', '100', '105']:
        d = get_json(f'{SD}/www/2.0/school/{sid}/info.json')
        if d and d.get('data', {}).get('pro_type', {}).get('area_2025'):
            a = d['data']['pro_type']['area_2025']
            m = {}
            for area, types in a.items():
                if area == 'area_id' or not isinstance(types, list):
                    continue
                ts = [str(t) for t in types if str(t) in KEEP_TYPES]
                if ts:
                    m[str(area)] = ts
            return m
    raise SystemExit('cannot load province-type map')

with open(os.path.join(BASE, 'school_code.json'), encoding='utf-8') as f:
    scd = json.load(f)['data']
SIDS = []
seen = set()
for v in scd.values():
    sid = str(v['school_id'])
    if sid not in seen:
        seen.add(sid); SIDS.append(sid)

def parse_row(it):
    g = lambda k: it.get(k) if it.get(k) not in (None, '') else None
    row = {'spname': g('spname'), 'sp_code': g('sp_code'),
           'group': g('special_group'), 'select': g('scsub_select_data'),
           'length': g('length'), 'fee': g('fee'), 'info': g('sp_info'),
           'zslx': g('zslx_name'), 'batch': g('local_batch_name')}
    for y in YEARS:
        row['min_' + y] = g('min_' + y)
        row['sec_' + y] = g('min_section_' + y)
    return row

def fetch_school_area_type(sid, area, typ):
    rows = []
    first = get_json(f'{SD}/www/2.0/schoolprovinceindex/2025/{sid}/{area}/{typ}/1.json')
    if not first or not first.get('data'):
        return rows
    data = first['data']
    pages = int(data.get('data_pages') or 1)
    for it in (data.get('data') or []):
        rows.append(parse_row(it))
    for p in range(2, pages + 1):
        d = get_json(f'{SD}/www/2.0/schoolprovinceindex/2025/{sid}/{area}/{typ}/{p}.json')
        if d and d.get('data', {}).get('data'):
            for it in d['data']['data']:
                rows.append(parse_row(it))
    return rows

def crawl_area(area, types):
    fp = os.path.join(OUT, f'{area}.json')
    schools = {}
    cnt = [0]; hit = [0]; lock = threading.Lock(); t0 = time.time()
    def work(sid):
        allrows = []
        for typ in types:
            allrows += fetch_school_area_type(sid, area, typ)
        with lock:
            cnt[0] += 1
            if allrows:
                schools[sid] = {'rows': allrows}; hit[0] += 1
            if cnt[0] % 300 == 0:
                print(f'  [{area}] {cnt[0]}/{len(SIDS)} schools_hit={hit[0]} {int(time.time()-t0)}s', flush=True)
    with ThreadPoolExecutor(max_workers=8) as ex:
        list(ex.map(work, SIDS))
    obj = {'area': area, 'types': types, 'schools': schools}
    with open(fp, 'w', encoding='utf-8') as f:
        json.dump(obj, f, ensure_ascii=False)
    nrows = sum(len(s['rows']) for s in schools.values())
    print(f'[DONE area={area}] schools={hit[0]} rows={nrows} {int(time.time()-t0)}s -> {fp}', flush=True)

def main():
    arg = sys.argv[1] if len(sys.argv) > 1 else 'all'
    pmap = load_prov_types()
    print(f'provinces={len(pmap)} school_ids={len(SIDS)}', flush=True)
    if arg == 'all':
        targets = sorted(pmap.keys(), key=lambda x: int(x))
    else:
        if arg not in pmap:
            raise SystemExit(f'area {arg} not in map; options: {",".join(sorted(pmap))}')
        targets = [arg]
    for area in targets:
        crawl_area(area, pmap[area])

if __name__ == '__main__':
    main()
