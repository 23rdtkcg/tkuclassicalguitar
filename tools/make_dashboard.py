"""產生「專案追蹤網站」靜態快照。
用法：python3 tools/make_dashboard.py <活動總表.xlsx> <輸出.html>
資料來源：活動總表（Google 試算表匯出 xlsx）。總表沒有、但原規劃有的工作，會標「總表缺漏」。
"""
import sys, json, datetime, re, pathlib
from openpyxl import load_workbook
sys.path.insert(0, str(pathlib.Path(__file__).parent))

ROOT = pathlib.Path(__file__).parent
src, out = sys.argv[1], sys.argv[2]
wb = load_workbook(src, data_only=True)

def d(v):
    return v.strftime('%Y-%m-%d') if isinstance(v, (datetime.datetime, datetime.date)) else None
def s(v):
    return '' if v is None else str(v).strip()

# ---- 原規劃 56 項（從 make_schedule_xlsx 的 T 清單取得）----
spec = (ROOT / 'make_schedule_xlsx.py').read_text()
# 直接用正則擷取 T 清單，避免執行整支腳本
block = spec.split('T=[', 1)[1].split('\n]\n', 1)[0]
T = eval('[' + block + ']')
G = {"總": "總召、副召", "B2B": "廠商對接 B2B", "賽": "賽事運行", "B2C": "行銷公關 B2C", "行": "事務行政"}
def D(md):
    m, dd = map(int, md.split('/'))
    return datetime.date(2026 if m >= 11 else 2027, m, dd).isoformat()
plan, cnt = {}, {}
for g, name, st, en, pre, note in T:
    cnt[g] = cnt.get(g, 0) + 1
    code = f"{g}-{cnt[g]:02d}"
    plan[code] = dict(code=code, title=name, group=G[g], owner='', start=D(st), end=D(en),
                      status='未開始', pre=pre, note=note, missing=True)

# ---- 工作總表（現行）----
ws = wb['工作總表']
hdr = [s(c.value) for c in ws[1]]
ix = {h: i for i, h in enumerate(hdr)}
sheet_tasks = {}
for r in ws.iter_rows(min_row=2, values_only=True):
    code = s(r[ix['編號']])
    if not code:
        continue
    full = s(r[ix['完整原文／備註']])
    note = full.split('｜', 1)[1] if '｜' in full else ''
    sheet_tasks[code] = dict(code=code, title=s(r[ix['工作事項']]), group=s(r[ix['負責組']]),
        owner=s(r[ix['負責人']]), start=d(r[ix['開始日']]), end=d(r[ix['截止日']]),
        status=s(r[ix['狀態']]) or '未開始', pre=s(r[ix['前置工作']]), note=note,
        done=d(r[ix['實際完成日']]), missing=False)
tasks = []
for code, p in plan.items():
    tasks.append(sheet_tasks.pop(code, p))
tasks.extend(sheet_tasks.values())  # 總表新增、原規劃沒有的

# ---- 里程碑（前置／備註欄位錯位，不採用）----
ms = []
ws = wb['里程碑']
for r in ws.iter_rows(min_row=2, values_only=True):
    if not s(r[0]):
        continue
    ms.append(dict(code=s(r[0]), date=d(r[1]), title=s(r[2]), group=s(r[3]), cond=s(r[4]),
                   fallback=s(r[5]), status=s(r[6]) or '未開始', stage=s(r[13]) if len(r) > 13 else ''))

# ---- 決策 ----
dec = []
for r in wb['決策追蹤'].iter_rows(min_row=2, values_only=True):
    if not s(r[0]):
        continue
    dec.append(dict(code=s(r[0]), q=s(r[1]), rec=s(r[2]), type=s(r[3]), due=d(r[4]),
                    answer=s(r[5]), opts=s(r[7]), task=s(r[8]), ms=s(r[9])))

# ---- 名冊（只取姓名與角色，不取任何聯絡資料）----
people = []
for r in wb['組別與人員'].iter_rows(min_row=41, max_row=66, values_only=True):
    if s(r[0]) and s(r[1]) and s(r[0]) != '組別':
        people.append(dict(group=s(r[0]), name=s(r[1]), role=s(r[2]), note=s(r[3])))

# ---- 經費 ----
wsb = wb['經費總覽']
rows = {s(r[0]): r for r in wsb.iter_rows(values_only=True) if s(r[0])}
def num(v):
    try: return round(float(v))
    except (TypeError, ValueError): return 0
def trio(key):
    r = rows.get(key)
    return [num(r[1]), num(r[2]), num(r[3])] if r else [0, 0, 0]
income = []
for r in wsb.iter_rows(values_only=True):
    if re.fullmatch(r'I\d\d', s(r[0])):
        income.append(dict(code=s(r[0]), name=s(r[1]), amount=num(r[2]), status=s(r[3]), note=s(r[4])))
budget = dict(plans=['方案一｜原案全戶外', '方案二｜戶外輕量', '方案三｜戶外音樂祭＋室內比賽'],
    total=trio('總支出（含10%預備金）'), target=100000, lean=trio('精簡後需籌措的規模'),
    gapConfirmed=trio('缺口｜僅已確認'), gapAll=trio('缺口｜若預估都到位'), income=income)

data = dict(snapshot=datetime.date.today().isoformat(), eventStart='2027-04-29', eventEnd='2027-04-30',
            tasks=tasks, milestones=ms, decisions=dec, people=people, budget=budget,
            sheetCount=len([t for t in tasks if not t['missing']]),
            missingCount=len([t for t in tasks if t['missing']]))
tpl = (ROOT / 'dashboard_template.html').read_text()
pathlib.Path(out).write_text(tpl.replace('/*__DATA__*/null', json.dumps(data, ensure_ascii=False)))
print('tasks', len(tasks), 'in sheet', data['sheetCount'], 'missing', data['missingCount'],
      'milestones', len(ms), 'decisions', len(dec), 'people', len(people), 'income', len(income))
