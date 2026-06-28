import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', line_buffering=True)
import requests, json, os, time, re, threading
from concurrent.futures import ThreadPoolExecutor

OUT = r'E:\ai\second\mnzy_cache_all'
os.makedirs(OUT, exist_ok=True)
API = 'https://mnzy.gaokao.cn/api/pc/v2/v2/query/universityAnalysis'
SIGN = 'a45687483114c02627a052ba7245482d'
H = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
     'Referer': 'https://mnzy.gaokao.cn/'}

def fetch(code, gt, cl, bt, sc):
    subj = '物理,化学,生物' if cl == '物理' else '历史,政治,地理'
    p = {'province':'广东','classify':cl,'subjects':subj,'gradeType':gt,
         'score':sc,'rank':'26988','recruitCode':str(code),'batch':bt,'signsafe':SIGN}
    for _ in range(3):
        try:
            d = requests.get(API, params=p, headers=H, timeout=15).json()
            return d.get('body') or []
        except Exception:
            time.sleep(0.6)
    return None

def name_of(body):
    if not body: return None
    for g in body:
        for b in g.get('branches', []):
            lg = b.get('logo') or ''
            if 'images/' in lg:
                nm = lg.split('images/')[-1].rsplit('.', 1)[0]
                if nm and not re.fullmatch(r'[0-9a-fA-F]{16,}', nm):
                    return nm
    return None

cnt = [0]; hit = [0]; fail = [0]; lock = threading.Lock(); t0 = time.time()

def work(code):
    fp = os.path.join(OUT, f'{code}.json')
    if os.path.exists(fp):
        with lock: cnt[0] += 1
        return
    level = phys = hist = None
    bp = fetch(code, '本科', '物理', '本科批', '600')
    if bp is None:
        with lock:
            fail[0] += 1; cnt[0] += 1
        return
    if bp:
        level, phys = '本科', bp
        hist = fetch(code, '本科', '历史', '本科批', '600')
    else:
        sp = fetch(code, '专科', '物理', '专科批', '400')
        if sp:
            level, phys = '专科', sp
            hist = fetch(code, '专科', '历史', '专科批', '400')
    if level:
        nm = name_of(phys) or name_of(hist) or f'院校{code}'
        rec = {'code': str(code), 'level': level, 'name': nm,
               'phys': phys or [], 'hist': hist or []}
        with open(fp, 'w', encoding='utf-8') as f:
            json.dump(rec, f, ensure_ascii=False)
        with lock: hit[0] += 1
    with lock:
        cnt[0] += 1
        if cnt[0] % 200 == 0:
            el = int(time.time() - t0)
            print(f'[progress] scanned={cnt[0]}/9999 hit={hit[0]} fail={fail[0]} elapsed={el}s', flush=True)

codes = list(range(10001, 20000))
print(f'[start] 枚举 {len(codes)} 个国标码, 6并发', flush=True)
with ThreadPoolExecutor(max_workers=6) as ex:
    ex.map(work, codes)
el = int(time.time() - t0)
print(f'[DONE] scanned={cnt[0]} hit={hit[0]} fail={fail[0]} elapsed={el}s', flush=True)
