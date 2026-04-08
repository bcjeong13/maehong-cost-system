#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
매홍엘앤에프 통합 원가 관리 시스템 - 웹 대시보드 v2
클릭 시 BOM 상세 / 원부재료 계산식 / 인건비 계산식 모달 표시
"""
import sys, os, json, openpyxl
from collections import defaultdict
from datetime import datetime
from flask import Flask, render_template_string, jsonify, request, session, redirect, url_for
from functools import wraps
import shutil

try:
    sys.stdout.reconfigure(encoding='utf-8')
except:
    pass
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
os.chdir(BASE_DIR)

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB
app.secret_key = os.environ.get('SECRET_KEY', 'maehong-cost-system-2026-secret')

# ============================================================
# 사용자 계정 (추후 DB로 전환 가능)
# ============================================================
USERS = {
    'admin': {'password': 'admin1234', 'role': 'admin', 'name': '관리자'},
    'user':  {'password': 'user1234',  'role': 'user',  'name': '사용자'},
}

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'user' not in session:
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'ok':False,'msg':'로그인이 필요합니다'}), 401
            return redirect('/login')
        return f(*args, **kwargs)
    return decorated

def admin_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get('role') != 'admin':
            if request.is_json or request.path.startswith('/api/'):
                return jsonify({'ok':False,'msg':'관리자 권한이 필요합니다'}), 403
            return redirect('/')
        return f(*args, **kwargs)
    return decorated

def is_admin():
    return session.get('role') == 'admin'

# ============================================================
# 데이터 로드
# ============================================================
SRC = 'pricing_base.xlsx' if os.path.exists(os.path.join(BASE_DIR, 'pricing_base.xlsx')) else '단가기준정리.xlsx'
PRD = 'production.xlsx' if os.path.exists(os.path.join(BASE_DIR, 'production.xlsx')) else '생산실적.xlsx'
wb_src = openpyxl.load_workbook(SRC, data_only=True)
wb_prd = openpyxl.load_workbook(PRD, data_only=True)

# 자사 생산제품
ws = wb_src['자사 생산제품']
products = {}
for r in range(2, ws.max_row+1):
    pn = ws.cell(r,3).value
    if not pn: continue
    products[str(pn).strip()] = {
        'pn': str(pn).strip(), 'name': ws.cell(r,4).value or '',
        'category': ws.cell(r,5).value or '', 'weight_g': ws.cell(r,6).value or 0,
        'type': ws.cell(r,7).value or '',
    }

# BOM (원본 보존 - 반제품 포함)
ws = wb_src['bom현황']
bom = defaultdict(list)
bom_raw = defaultdict(list)  # 반제품 포함 원본
for r in range(2, ws.max_row+1):
    mo = ws.cell(r,2).value
    if not mo: continue
    item = {
        'ja_pn': str(ws.cell(r,11).value).strip() if ws.cell(r,11).value else '',  # col11: 자품번
        'ja_name': ws.cell(r,12).value or '',   # col12: 자품명
        'ja_type': ws.cell(r,13).value or '',   # col13: 규격
        'ja_unit': ws.cell(r,14).value or '',   # col14: 재고단위
        'qty_net': ws.cell(r,15).value or 0,    # col15: 정미수량
        'loss_pct': ws.cell(r,16).value or 0,   # col16: LOSS(%)
        'qty_req': ws.cell(r,17).value or 0,    # col17: 필요수량
    }
    mo_s = str(mo).strip()
    bom[mo_s].append(item)
    bom_raw[mo_s].append(item)

# 반제품 메타 (E코드 → 품명, 규격, 카테고리)
semi_products = {}
for r in range(2, ws.max_row+1):
    mo = ws.cell(r,2).value
    if not mo: continue
    mo_s = str(mo).strip()
    if mo_s.startswith('E') and mo_s not in semi_products:
        cat_col = ws.cell(r,4).value or ''    # col4: 카테고리 (고구마류 / 고구마바류)
        spec = ws.cell(r,5).value or ''       # col5: 규격 (반제품 / 반제품(2))
        semi_products[mo_s] = {
            'pn': mo_s,
            'name': ws.cell(r,3).value or '',
            'spec': spec,
            'cat': '고구마바류' if '고구마바' in str(cat_col) else '고구마류',
        }

# 인건비
ws = wb_src['인건비 현황']
employee_wages = {}
for r in range(2, ws.max_row+1):
    name = ws.cell(r,3).value
    pay = ws.cell(r,6).value
    col7 = ws.cell(r,7).value or ''
    if name:
        employee_wages[str(name).strip()] = {
            'pay': pay if pay else 0, 'common': '공통배부' in str(col7),
        }

# 원부재료비
ws = wb_src['원부재료비']
material_prices = {}
for r in range(2, ws.max_row+1):
    pn = ws.cell(r,10).value; price = ws.cell(r,17).value; dt = ws.cell(r,2).value
    if pn and price:
        pn_s = str(pn).strip()
        if pn_s not in material_prices or (dt and dt > (material_prices[pn_s]['date'] or datetime.min)):
            material_prices[pn_s] = {'price': price, 'date': dt}

# 원가보고서 로드
ws_cost_report = wb_src['원가보고서'] if '원가보고서' in wb_src.sheetnames else None
cost_report = {}  # {'1월': {'노무비':X, '급여':Y, '퇴직급여':Z}, ...}
if ws_cost_report:
    # 월 컬럼 매핑 (Col3=1월, Col4=2월, ...)
    for c in range(3, ws_cost_report.max_column+1):
        month_name = ws_cost_report.cell(1, c).value
        if not month_name: continue
        month_name = str(month_name).strip()
        labor_total = ws_cost_report.cell(14, c).value or 0  # Row14: 노무비 합계
        salary = ws_cost_report.cell(15, c).value or 0        # Row15: 급여
        misc = ws_cost_report.cell(16, c).value or 0          # Row16: 잡급
        retire = ws_cost_report.cell(17, c).value or 0         # Row17: 퇴직급여
        cost_report[month_name] = {
            'labor_total': labor_total,
            'salary': salary,
            'misc': misc,
            'retire': retire,
        }

# 생산실적 (누적 저장)
prod_records = []

def load_prod_xlsx(filepath):
    """생산실적 xlsx를 파싱하여 레코드 리스트 반환"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in range(2, ws.max_row+1):
        dt = ws.cell(r,4).value; pn = ws.cell(r,10).value
        if not dt or not pn: continue
        rows.append({
            'date': str(dt)[:10], 'pn': str(pn).strip(),
            'name': ws.cell(r,11).value or '', 'qty': ws.cell(r,15).value or 0,
            'erp_price': ws.cell(r,23).value or 0, 'bigo': str(ws.cell(r,30).value or ''),
        })
    return rows

# 초기 파일 로드
if os.path.exists(PRD):
    prod_records = load_prod_xlsx(PRD)

# ============================================================
# 계산 엔진
# ============================================================
MIN_WAGE = 10300; H = 209
RETIRE_RATE = 1/12  # 퇴직급여 비율 (8.33%)

def hw(name):
    """시급 = (지급총액 + 퇴직급여) ÷ 209h = 지급총액 × (1+1/12) ÷ 209"""
    e = employee_wages.get(name.strip())
    if e and e['pay']>0:
        return e['pay'] * (1 + RETIRE_RATE) / H
    return MIN_WAGE * (1 + RETIRE_RATE)

def hw_detail(name):
    """시급 + 계산식 상세 반환 (퇴직급여 포함)"""
    e = employee_wages.get(name.strip())
    if e and e['pay']>0:
        pay = e['pay']
        retire = round(pay * RETIRE_RATE)
        total = pay + retire
        hourly = round(total / H, 2)
        return {'name':name,'pay':pay,'retire':retire,'total_pay':total,'hourly':hourly,
                'formula':f"({pay:,} + 퇴직{retire:,}) ÷ {H}h",'is_min':False}
    min_hw = round(MIN_WAGE * (1 + RETIRE_RATE), 2)
    return {'name':name,'pay':0,'retire':0,'total_pay':0,'hourly':min_hw,
            'formula':f"최저시급 {MIN_WAGE:,} + 퇴직8.33%",'is_min':True}

def wsum(s, hc):
    min_hw = MIN_WAGE * (1 + RETIRE_RATE)  # 최저시급에도 퇴직급여 포함
    if not s: return min_hw * hc
    names = [n.strip() for n in s.replace(' ',',').split(',') if n.strip()]
    t = sum(hw(n) for n in names)
    rem = hc - len(names)
    return t + (min_hw*rem if rem>0 else 0)

def mat_price(pn):
    return material_prices.get(pn,{}).get('price',0)

def mat_date(pn):
    d = material_prices.get(pn,{}).get('date')
    return str(d)[:10] if d else ''

def explode(pn, qty=1):
    res = []
    for it in bom.get(pn,[]):
        ja = it['ja_pn']; req = (it['qty_req'] or 0)*qty
        if ja.startswith('E'):
            res.extend(explode(ja, req))
        else:
            p = mat_price(ja); tp = '원재료' if ja.startswith('A') else '부재료'
            res.append({'pn':ja,'name':it['ja_name'],'type':tp,'unit':it['ja_unit'],'qty':req,'price':p,'cost':p*req,'date':mat_date(ja)})
    return res

def find_kg(pn, qty=1):
    t = 0
    for it in bom.get(pn,[]):
        ja = it['ja_pn']; req = (it['qty_req'] or 0)*qty
        if ja.startswith('E'):
            if it['ja_unit']=='KG': t += req
            else: t += find_kg(ja, req)
    return t

def inner_ea(pn):
    for it in bom.get(pn,[]):
        if it['ja_pn'].startswith('E') and it['qty_req']: return it['qty_req']
    return 1

# 공정 정의 (단가 + 메타)
PROC_META = {
    '절단':     {'workers':'김흥수,한승엽,박미영,송선임','hc':4,'hours':8,'capa':1000,'unit':'KG'},
    '큐브':     {'workers':'정미혜,아다치에리,이소현','hc':4,'hours':8,'capa':1000,'unit':'KG'},
    '로터리':   {'workers':'서혜진,김하윤','hc':3,'hours':8,'capa':12000,'unit':'EA'},
    '수작업':   {'workers':'','hc':7,'hours':8,'capa':9310,'unit':'EA'},
    '나라시':   {'workers':'최엘라','hc':1,'hours':8,'capa':12000,'unit':'EA'},
    '살균':     {'workers':'사만,선티조이','hc':2,'hours':8,'capa':8960,'unit':'EA'},
    '선날인':   {'workers':'김영미','hc':1,'hours':8,'capa':8960,'unit':'EA'},
    '번들':     {'workers':'권미정,박예민,조명희','hc':3,'hours':8,'capa':13230,'unit':'EA'},
    '낱봉':     {'workers':'권미정,박예민,조명희','hc':3,'hours':8,'capa':14700,'unit':'EA'},
    '절단바':   {'workers':'김흥수,한승엽,박미영,송선임','hc':4,'hours':8,'capa':860,'unit':'KG'},
    '선별바':   {'workers':'정미혜,아다치에리,서혜진','hc':8,'hours':8,'capa':21728,'unit':'EA'},
    '살균바':   {'workers':'사만,선티조이','hc':2,'hours':8,'capa':22400,'unit':'EA'},
    '번들바':   {'workers':'권미정,박예민,조명희','hc':3,'hours':8,'capa':22386,'unit':'EA'},
    '오트밀':   {'workers':'이근용,김영숙,권해조','hc':5,'hours':8,'capa':2200,'unit':'EA'},
    '스틱':     {'workers':'이영주,고은하,김수민,강병윤','hc':5,'hours':8,'capa':7000,'unit':'EA'},
    '공통':     {'workers':'정병찬,지우철,김진천,김영윤,이진호,심민호,박진주','hc':7,'hours':8,'capa':9310,'unit':'EA'},
}
P = {}
for k,m in PROC_META.items():
    P[k] = wsum(m['workers'],m['hc'])*m['hours']/m['capa']

# ============================================================
# 공통배부 인건비 (관리자 7명 → 전 공정에 배부)
# ============================================================
def calc_common_rate():
    """공통배부율 = 공통인원 월급합(퇴직포함) ÷ 생산직 월급합(퇴직포함)
    직접인건비 비율 배부 방식 (원가회계 표준)"""
    common_total = 0
    prod_total = 0
    for name, e in employee_wages.items():
        if e['pay'] > 0:
            pay_with_retire = e['pay'] * (1 + RETIRE_RATE)
            if e['common']:
                common_total += pay_with_retire
            else:
                prod_total += pay_with_retire
    return common_total / prod_total if prod_total > 0 else 0

COMMON_RATE = calc_common_rate()  # 약 39.7%

def calc_labor(pn, manual=True):
    pr = products.get(pn)
    if not pr: return 0, []
    cat, pt = pr['category'], pr['type']
    items = []; kg = find_kg(pn); ea = inner_ea(pn)
    if '고구마' in cat and '바' not in cat and '스틱' not in cat:
        items.append(('절단', P['절단']*kg, '절단', kg))
        if '큐브' in cat: items.append(('선별(큐브)', P['큐브']*kg, '큐브', kg))
        pk_key = '수작업' if manual else '로터리'
        items.append(('내포장(수작업)' if manual else '내포장(로터리)', P[pk_key]*ea, pk_key, ea))
        items.append(('대차나라시', P['나라시']*ea, '나라시', ea))
        items.append(('살균', P['살균']*ea, '살균', ea))
        items.append(('선날인', P['선날인']*ea, '선날인', ea))
        ok = '번들' if '번들' in pt else '낱봉'
        items.append(('외포장('+('번들' if '번들' in pt else '낱봉')+')', P[ok]*1, ok, 1))
    elif '고구마' in cat and ('바' in cat or '바' in pr['name']):
        items.append(('절단', P['절단바']*kg, '절단바', kg))
        items.append(('선별(바)', P['선별바']*kg, '선별바', kg))
        pk_key = '수작업' if manual else '로터리'
        items.append(('내포장', P[pk_key]*ea, pk_key, ea))
        items.append(('나라시', P['나라시']*ea, '나라시', ea))
        items.append(('살균', P['살균바']*ea, '살균바', ea))
        items.append(('선날인', P['선날인']*ea, '선날인', ea))
        ok = '번들바' if '번들' in pt else '낱봉'
        items.append(('외포장', P[ok]*1, ok, 1))
    elif '오트밀' in cat or '오트밀' in pr['name']:
        items.append(('생산', P['오트밀']*ea, '오트밀', ea))
        if '번들' in pt: items.append(('외포장', P['번들']*1, '번들', 1))
    elif '스틱' in cat or '스틱' in pr['name']:
        items.append(('생산', P['스틱']*ea, '스틱', ea))
        if '번들' in pt: items.append(('외포장', P['번들']*1, '번들', 1))
    else:
        if kg>0: items.append(('절단', P['절단']*kg, '절단', kg))
        if ea>1:
            pk_key = '수작업' if manual else '로터리'
            items.append(('내포장', P[pk_key]*ea, pk_key, ea))
            items.append(('살균', P['살균']*ea, '살균', ea))
        if '번들' in pt: items.append(('외포장', P['번들']*1, '번들', 1))
    # 공통배부 = 직접인건비 합계 × 공통배부율
    direct_labor = sum(c for _,c,_,_ in items)
    common = direct_labor * COMMON_RATE
    if common > 0:
        items.append(('공통배부(관리)', common, '공통', 1))
    return sum(c for _,c,_,_ in items), items

def calc_semi_labor(pn, manual=True):
    """반제품(E코드) 인건비 계산.
    - 절단&건조 반제품: 절단 공정만
    - 살균후반제품: 내포장~살균 (절단 제외)
    카테고리 판별: BOM 규격의 '고구마바류' → 바류 CAPA, 그 외 → 고구마류 CAPA
    """
    sp = semi_products.get(pn)
    if not sp:
        return 0, []
    name = sp['name']
    cat = sp['cat']  # '고구마류' or '고구마바류'
    is_bar = cat == '고구마바류'

    items = []
    kg = find_kg(pn)

    is_cut = '절단' in name and '건조' in name
    is_sal = '살균' in name or '반제품(2)' in sp['spec']

    if is_cut:
        # 절단&건조: 절단만 (1KG 기준)
        cut_key = '절단바' if is_bar else '절단'
        items.append(('절단', P[cut_key] * 1, cut_key, 1))
    elif is_sal:
        # 살균후반제품 (절단 제외)
        if is_bar:
            # 고구마바류: 선별(바타입) 860KG + 살균 22,400EA
            items.append(('선별(바타입)', P['선별바'] * 1, '선별바', 1))  # 1KG 기준
        else:
            # 고구마류: 큐브 선별 or 내포장 + 나라시
            if '큐브' in name:
                items.append(('선별(큐브)', P['큐브'] * kg, '큐브', kg))
            pk_key = '수작업' if manual else '로터리'
            pk_label = '내포장(수작업)' if manual else '내포장(로터리)'
            items.append((pk_label, P[pk_key] * 1, pk_key, 1))
            items.append(('대차나라시', P['나라시'] * 1, '나라시', 1))

        # 살균: 고구마바류 22,400 / 고구마류 8,960
        sal_key = '살균바' if is_bar else '살균'
        items.append(('살균', P[sal_key] * 1, sal_key, 1))
        items.append(('선날인', P['선날인'] * 1, '선날인', 1))

    # 반제품도 직접인건비 비율 배부
    direct_labor = sum(c for _, c, _, _ in items)
    common = direct_labor * COMMON_RATE
    if common > 0:
        items.append(('공통배부(관리)', common, '공통', 1))
    total = sum(c for _, c, _, _ in items)
    return total, items

def calc_cost(pn, manual=True):
    mi = explode(pn)
    raw = sum(m['cost'] for m in mi if m['type']=='원재료')
    sub = sum(m['cost'] for m in mi if m['type']=='부재료')
    lab, li = calc_labor(pn, manual)
    return {'raw':raw,'sub':sub,'mat':raw+sub,'labor':lab,'labor_items':li,'total':raw+sub+lab,'mat_items':mi}

# 반제품 인건비도 사전 계산
semi_costs = {}
for epn in semi_products:
    lab_m, li_m = calc_semi_labor(epn, True)
    lab_r, li_r = calc_semi_labor(epn, False)
    semi_costs[epn] = {'m': {'labor': lab_m, 'labor_items': li_m}, 'r': {'labor': lab_r, 'labor_items': li_r}}

all_costs = {}
for pn in products:
    all_costs[pn] = {'m': calc_cost(pn, True), 'r': calc_cost(pn, False)}

def get_daily_prod_wage():
    return sum((e['pay']/H)*8 for e in employee_wages.values() if e['pay']>0 and not e['common'])

def get_std_ea_per_mh(pn):
    """기준 EA/MH + 산출 근거 반환.
    returns (ea_per_mh, {'proc':공정명, 'capa':수량, 'hc':인원, 'hours':시간, 'unit':단위})"""
    def _result(key, proc_name):
        m = PROC_META[key]
        ea_mh = m['capa'] / (m['hc'] * m['hours'])
        return ea_mh, {'proc': proc_name, 'capa': m['capa'], 'hc': m['hc'], 'hours': m['hours'], 'unit': m['unit']}

    if pn in products:
        pr = products[pn]
        cat = pr['category']
        if '고구마' in cat and '바' not in cat:
            return _result('수작업', '내포장(수작업)')
        elif '고구마' in cat and '바' in cat:
            return _result('선별바', '선별(바타입)')
        elif '오트밀' in cat or '오트밀' in pr['name']:
            return _result('오트밀', '오트밀생산')
        elif '스틱' in cat or '스틱' in pr['name']:
            return _result('스틱', '스틱생산')
    elif pn in semi_products:
        sp = semi_products[pn]
        is_bar = sp['cat'] == '고구마바류'
        is_cut = '절단' in sp['name'] and '건조' in sp['name']
        is_sal = '살균' in sp['name'] or '반제품(2)' in sp['spec']
        if is_cut:
            key = '절단바' if is_bar else '절단'
            return _result(key, '절단')
        elif is_sal:
            if is_bar:
                return _result('선별바', '선별(바타입)')
            else:
                return _result('수작업', '내포장(수작업)')
    return 0, {}

# 실제인건비 데이터 저장 (key: "날짜|품번|수량" → [{name, hours},...])
actual_labor_data = {}

def actual_labor_key(date, pn, qty):
    return f"{date}|{pn}|{qty}"

LUNCH_START = 12*60  # 12:00
LUNCH_END = 13*60    # 13:00
LUNCH_MIN = LUNCH_END - LUNCH_START  # 60분

def _calc_work_minutes(start_str, end_str):
    """작업시간(분) 계산 — 점심시간(12:00~13:00) 자동 차감"""
    sh,sm = map(int, start_str.split(':'))
    eh,em = map(int, end_str.split(':'))
    s_min = sh*60+sm
    e_min = eh*60+em
    if e_min < s_min: e_min += 24*60  # 야간
    total = e_min - s_min
    # 점심시간 겹침 계산
    overlap_start = max(s_min, LUNCH_START)
    overlap_end = min(e_min, LUNCH_END)
    if overlap_start < overlap_end:
        total -= (overlap_end - overlap_start)
    return max(total, 0)

def calc_actual_labor(data):
    """실제인건비 계산 (점심시간 자동 차감, 공통배부 포함).
    data = {timeSlots:[{start,end},...], grid:{사원명:[bool,bool,...]}}
    """
    ts = data.get('timeSlots', [])
    grid = data.get('grid', {})
    if not ts or not grid:
        slots = data if isinstance(data, list) else []
        direct = 0
        for s in slots:
            name = s.get('name','').strip()
            hours = float(s.get('hours',0))
            direct += hw(name) * hours
        total = direct * (1 + COMMON_RATE)
        return round(total, 0)
    # 시간대별 실작업시간(분→시간)
    slot_hours = []
    for s in ts:
        mins = _calc_work_minutes(s['start'], s['end'])
        slot_hours.append(round(mins/60, 4))
    direct = 0
    for name, checks in grid.items():
        h = sum(slot_hours[i] for i,c in enumerate(checks) if c and i < len(slot_hours))
        direct += hw(name) * h
    total = direct * (1 + COMMON_RATE)
    return round(total, 0)

# ============================================================
# BOM 트리 생성 (모달용)
# ============================================================
def build_bom_tree(pn, qty=1, depth=0):
    """BOM을 트리 구조로 반환 (반제품 포함)"""
    nodes = []
    for it in bom_raw.get(pn,[]):
        ja = it['ja_pn']; req = (it['qty_req'] or 0)*qty
        price = mat_price(ja); cost = price * req
        node = {
            'pn': ja, 'name': it['ja_name'], 'type': it['ja_type'],
            'unit': it['ja_unit'], 'qty_net': it['qty_net'], 'loss': it['loss_pct'],
            'qty_req': it['qty_req'], 'qty_total': round(req, 6),
            'price': price, 'cost': round(cost, 2), 'depth': depth,
            'children': [],
        }
        if ja.startswith('E'):
            node['children'] = build_bom_tree(ja, req, depth+1)
            node['cost'] = sum(c['cost'] for c in explode(ja, req))
        nodes.append(node)
    return nodes

# ============================================================
# API 엔드포인트
# ============================================================
@app.route('/api/bom/<pn>')
@login_required
def api_bom(pn):
    """BOM 트리 + 제품 기본정보"""
    pr = products.get(pn, {})
    tree = build_bom_tree(pn)
    return jsonify({'product': pr, 'bom_tree': tree})

@app.route('/api/material/<pn>')
@login_required
def api_material(pn):
    """원부재료 상세 계산식"""
    pr = products.get(pn, {})
    mat_items = explode(pn)
    # BOM 경로 추적
    tree = build_bom_tree(pn)
    raw_total = sum(m['cost'] for m in mat_items if m['type']=='원재료')
    sub_total = sum(m['cost'] for m in mat_items if m['type']=='부재료')
    return jsonify({
        'product': pr,
        'mat_items': mat_items,
        'raw_total': round(raw_total, 2),
        'sub_total': round(sub_total, 2),
        'total': round(raw_total + sub_total, 2),
        'bom_tree': tree,
    })

@app.route('/api/labor/<pn>')
@login_required
@admin_required
def api_labor(pn):
    """인건비 상세 계산식 (admin 전용)"""
    pr = products.get(pn, {})
    if not pr and pn in semi_products:
        sp = semi_products[pn]
        pr = {'pn':pn, 'name':sp['name'], 'category':sp['spec'], 'weight_g':0, 'type':'반제품'}
    kg = find_kg(pn); ea = inner_ea(pn)
    # 제품이면 calc_labor, 반제품이면 calc_semi_labor
    if pn in products:
        _, items_m = calc_labor(pn, True)
        _, items_r = calc_labor(pn, False)
    elif pn in semi_products:
        _, items_m = calc_semi_labor(pn, True)
        _, items_r = calc_semi_labor(pn, False)
    else:
        items_m = items_r = []
    details = []
    direct_total = sum(c for _,c,k,_ in items_m if k != '공통')
    for label, cost, proc_key, usage in items_m:
        meta = PROC_META.get(proc_key, {})
        if proc_key == '공통':
            # 공통배부: 직접인건비 × 배부율 방식
            ws_str = meta.get('workers','')
            names = [n.strip() for n in ws_str.replace(' ',',').split(',') if n.strip()] if ws_str else []
            worker_details = [hw_detail(n) for n in names]
            hourly_sum = sum(w['hourly'] for w in worker_details)
            rate_pct = round(COMMON_RATE * 100, 1)
            details.append({
                'label': label, 'proc_key': proc_key,
                'workers': worker_details,
                'hourly_sum': round(hourly_sum, 2),
                'hours': 0, 'capa': 0, 'unit': '',
                'per_unit': 0, 'usage': 0,
                'cost': round(cost, 2),
                'is_common': True,
                'common_rate': rate_pct,
                'direct_labor': round(direct_total, 2),
                'formula': f"직접인건비 {direct_total:,.2f}원 × 배부율 {rate_pct}% = {cost:,.2f}원",
            })
        else:
            ws_str = meta.get('workers','')
            hc = meta.get('hc',0)
            names = [n.strip() for n in ws_str.replace(' ',',').split(',') if n.strip()] if ws_str else []
            worker_details = [hw_detail(n) for n in names]
            for i in range(hc - len(names)):
                min_hw = round(MIN_WAGE * (1 + RETIRE_RATE), 2)
                worker_details.append({'name':f'미등록{i+1}','pay':0,'hourly':min_hw,'formula':f'최저시급 {MIN_WAGE:,} + 퇴직8.33%','is_min':True})
            hourly_sum = sum(w['hourly'] for w in worker_details)
            capa = meta.get('capa',0)
            hours = meta.get('hours',8)
            unit = meta.get('unit','')
            per_unit = (hourly_sum * hours) / capa if capa > 0 else 0
            details.append({
                'label': label, 'proc_key': proc_key,
                'workers': worker_details,
                'hourly_sum': round(hourly_sum, 2),
                'hours': hours, 'capa': capa, 'unit': unit,
                'per_unit': round(per_unit, 4),
                'usage': round(usage, 4),
                'cost': round(cost, 2),
                'is_common': False,
                'formula': f"({hourly_sum:,.2f} × {hours}h) ÷ {capa:,}{unit} × {usage:.4f}{unit} = {cost:,.2f}원",
            })
    total_m = sum(d['cost'] for d in details)
    total_r = sum(c for _,c,_,_ in items_r)
    return jsonify({
        'product': pr,
        'cut_kg': round(kg, 4), 'inner_ea': ea,
        'details': details,
        'total_manual': round(total_m, 2),
        'total_rotary': round(total_r, 2),
    })

# ============================================================
# 파일 업로드 API
# ============================================================
@app.route('/api/upload', methods=['POST'])
@login_required
@admin_required
def api_upload():
    """단가기준정리.xlsx 업로드 → 백업 후 교체 → 서버 재로드"""
    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'ok':False,'msg':'xlsx 파일만 업로드 가능합니다'}), 400
    # 백업
    backup = SRC.replace('.xlsx', f'_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    if os.path.exists(SRC):
        shutil.copy2(SRC, backup)
    f.save(SRC)
    return jsonify({'ok':True,'msg':f'업로드 완료. 백업: {backup}\n페이지를 새로고침하면 반영됩니다.\n(서버 재시작 필요)'})

@app.route('/api/upload_prod', methods=['POST'])
@login_required
def api_upload_prod():
    """생산실적 xlsx 업로드 (누적 추가 or 교체)"""
    global prod_records
    f = request.files.get('file')
    mode = request.form.get('mode', 'append')  # append or replace
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'ok':False,'msg':'xlsx 파일만 업로드 가능합니다'}), 400
    # 임시 저장 후 파싱
    tmp = f'_tmp_prod_{datetime.now().strftime("%H%M%S")}.xlsx'
    f.save(tmp)
    try:
        new_rows = load_prod_xlsx(tmp)
    except Exception as e:
        os.remove(tmp)
        return jsonify({'ok':False,'msg':f'파일 파싱 오류: {e}'}), 400
    os.remove(tmp)
    if not new_rows:
        return jsonify({'ok':False,'msg':'유효한 생산실적 데이터가 없습니다'}), 400
    dates = sorted(set(r['date'] for r in new_rows))
    if mode == 'replace':
        prod_records = new_rows
        msg = f'교체 완료: {len(new_rows)}건 ({dates[0]} ~ {dates[-1]})'
    else:
        # 누적: 같은 날짜+품번+수량의 중복 제거
        existing = set((r['date'],r['pn'],r['qty']) for r in prod_records)
        added = 0
        for r in new_rows:
            key = (r['date'],r['pn'],r['qty'])
            if key not in existing:
                prod_records.append(r)
                existing.add(key)
                added += 1
        msg = f'누적 추가: {added}건 (중복 {len(new_rows)-added}건 제외) | 기간: {dates[0]} ~ {dates[-1]}'
    return jsonify({'ok':True,'msg':msg,'total':len(prod_records),'dates':dates})

@app.route('/api/prod_records')
@login_required
def api_prod_records():
    """생산실적 조회 (필터 지원)"""
    date_from = request.args.get('from','')
    date_to = request.args.get('to','')
    search = (request.args.get('q','') or '').strip().lower()
    filtered = []
    for r in prod_records:
        if date_from and r['date'] < date_from: continue
        if date_to and r['date'] > date_to: continue
        if search and search not in r['pn'].lower() and search not in r['name'].lower(): continue
        pn = r['pn']
        is_prod = pn in products
        is_semi = pn.startswith('E') and pn in semi_products
        if is_prod and pn in all_costs:
            le = all_costs[pn]['m']['labor']; le_r = all_costs[pn]['r']['labor']
            me = all_costs[pn]['m']['mat']; te = all_costs[pn]['m']['total']
        elif is_semi and pn in semi_costs:
            le = semi_costs[pn]['m']['labor']; le_r = semi_costs[pn]['r']['labor']
            me = 0; te = le
        else:
            le = le_r = me = te = 0
        akey = actual_labor_key(r['date'], r['pn'], r['qty'])
        adata = actual_labor_data.get(akey, None)
        actual_total = calc_actual_labor(adata) if adata else None
        actual_ea = round(actual_total / r['qty'], 2) if actual_total and r['qty'] else None
        # 투입 Man-Hour 계산
        actual_mh = 0
        if adata and isinstance(adata, dict) and adata.get('timeSlots') and adata.get('grid'):
            for ti, slot in enumerate(adata['timeSlots']):
                mins = _calc_work_minutes(slot['start'], slot['end'])
                cnt = sum(1 for checks in adata['grid'].values() if ti < len(checks) and checks[ti])
                actual_mh += cnt * mins/60
        # 생산성: 기준 EA/MH vs 실제 EA/MH
        std_ea_mh, std_info = get_std_ea_per_mh(pn)
        actual_ea_mh = round(r['qty'] / actual_mh, 2) if actual_mh > 0 else None
        productivity = round(actual_ea_mh / std_ea_mh * 100, 1) if actual_ea_mh and std_ea_mh > 0 else None
        filtered.append({**r, 'is_prod':is_prod, 'labor_ea':round(le,2), 'labor_ea_r':round(le_r,2), 'mat_ea':round(me,2),
                         'total_ea':round(te,2), 'labor_sub':round(le*r['qty'],0), 'total_sub':round(te*r['qty'],0),
                         'actual_labor':actual_total, 'actual_labor_ea':actual_ea,
                         'actual_mh':round(actual_mh,1), 'std_ea_mh':round(std_ea_mh,1),
                         'std_info':std_info,
                         'actual_ea_mh':actual_ea_mh, 'productivity':productivity})
    # 날짜 목록
    all_dates = sorted(set(r['date'] for r in prod_records))
    # 일별 소계
    daily_summary = defaultdict(lambda:{'count':0,'labor':0,'total':0})
    for r in filtered:
        d = r['date']
        daily_summary[d]['count'] += 1
        daily_summary[d]['labor'] += r['labor_sub']
        daily_summary[d]['total'] += r['total_sub']
    return jsonify({
        'records': filtered,
        'all_dates': all_dates,
        'daily_summary': dict(daily_summary),
        'daily_wage': round(get_daily_prod_wage(),0),
        'total_count': len(prod_records),
    })

@app.route('/api/prod_detail/<pn>')
@login_required
def api_prod_detail(pn):
    """생산실적 품번 클릭 시 상세 (원가 + 해당 품번 실적 이력)"""
    pr = products.get(pn, {})
    if not pr and pn in semi_products:
        sp = semi_products[pn]
        pr = {'pn':pn, 'name':sp['name'], 'category':sp['spec'], 'weight_g':0, 'type':'반제품'}
    ac = all_costs.get(pn, {}).get('m', {})
    if not ac and pn in semi_costs:
        sc = semi_costs[pn]['m']
        ac = {'raw':0,'sub':0,'labor':sc['labor'],'total':sc['labor'],'mat_items':[],'labor_items':sc['labor_items']}
    history = [r for r in prod_records if r['pn'] == pn]
    history.sort(key=lambda x: x['date'], reverse=True)
    total_qty = sum(r['qty'] for r in history)
    return jsonify({
        'product': pr,
        'cost': {'raw':ac.get('raw',0),'sub':ac.get('sub',0),'labor':ac.get('labor',0),'total':ac.get('total',0)},
        'history': history[:50],
        'total_qty': total_qty,
        'history_count': len(history),
    })

# ============================================================
# 실제인건비 (투입인원) API
# ============================================================
@app.route('/api/actual_labor/get', methods=['POST'])
@login_required
def api_actual_labor_get():
    """특정 실적의 투입인원 슬롯 조회"""
    d = request.get_json()
    akey = actual_labor_key(d.get('date',''), d.get('pn',''), d.get('qty',0))
    slots = actual_labor_data.get(akey, [])
    total = calc_actual_labor(slots) if slots else 0
    # 생산직 사원목록 (공통배부 제외)
    emp_list = []
    for name, e in employee_wages.items():
        if not e['common']:
            emp_list.append({'name':name, 'pay':e['pay'], 'hourly':round(hw(name),0)})
    emp_list.sort(key=lambda x: x['name'])
    return jsonify({'data':slots, 'total':total, 'emp_list':emp_list, 'min_wage':MIN_WAGE})

@app.route('/api/actual_labor/save', methods=['POST'])
@login_required
def api_actual_labor_save():
    """투입인원 출석부 저장. data = {date,pn,qty,timeSlots,grid}"""
    d = request.get_json()
    akey = actual_labor_key(d.get('date',''), d.get('pn',''), d.get('qty',0))
    payload = {'timeSlots': d.get('timeSlots',[]), 'grid': d.get('grid',{})}
    # grid에서 한 곳이라도 체크된 사원만 남기기
    clean_grid = {name: checks for name, checks in payload['grid'].items() if any(checks)}
    payload['grid'] = clean_grid
    if clean_grid and payload['timeSlots']:
        actual_labor_data[akey] = payload
    else:
        actual_labor_data.pop(akey, None)
    total = calc_actual_labor(payload) if clean_grid else 0
    ppl = len(clean_grid)
    return jsonify({'ok':True, 'total':total, 'msg':f'{ppl}명 저장 (실제인건비: {total:,.0f}원)'})

# ============================================================
# 인건비 관리 API
# ============================================================
@app.route('/api/employees')
@login_required
@admin_required
def api_employees():
    """전체 인건비 목록"""
    rows = []
    for name, e in employee_wages.items():
        hourly = round(e['pay']*(1+RETIRE_RATE)/H, 2) if e['pay']>0 else round(MIN_WAGE*(1+RETIRE_RATE), 2)
        rows.append({'name':name,'pay':e['pay'],'common':e['common'],'hourly':hourly})
    rows.sort(key=lambda x: (-x['pay'], x['name']))
    return jsonify({'employees':rows,'min_wage':MIN_WAGE,'hours':H})

@app.route('/api/employees/add', methods=['POST'])
@login_required
@admin_required
def api_employee_add():
    """신규 인원 추가"""
    d = request.get_json()
    name = (d.get('name') or '').strip()
    pay = int(d.get('pay') or 0)
    common = bool(d.get('common', False))
    if not name:
        return jsonify({'ok':False,'msg':'이름을 입력하세요'}), 400
    if name in employee_wages:
        return jsonify({'ok':False,'msg':f'{name}은(는) 이미 등록되어 있습니다'}), 400
    employee_wages[name] = {'pay':pay,'common':common}
    _recalc_all()
    return jsonify({'ok':True,'msg':f'{name} 추가 완료 (지급총액: {pay:,}원)'})

@app.route('/api/employees/update', methods=['POST'])
@login_required
@admin_required
def api_employee_update():
    """인건비 금액 변경"""
    d = request.get_json()
    name = (d.get('name') or '').strip()
    pay = int(d.get('pay') or 0)
    common = bool(d.get('common', False))
    if name not in employee_wages:
        return jsonify({'ok':False,'msg':f'{name}을(를) 찾을 수 없습니다'}), 404
    old_pay = employee_wages[name]['pay']
    employee_wages[name] = {'pay':pay,'common':common}
    _recalc_all()
    return jsonify({'ok':True,'msg':f'{name} 변경 완료 ({old_pay:,}원 → {pay:,}원)'})

@app.route('/api/employees/delete', methods=['POST'])
@login_required
@admin_required
def api_employee_delete():
    """인원 삭제"""
    d = request.get_json()
    name = (d.get('name') or '').strip()
    if name not in employee_wages:
        return jsonify({'ok':False,'msg':f'{name}을(를) 찾을 수 없습니다'}), 404
    del employee_wages[name]
    _recalc_all()
    return jsonify({'ok':True,'msg':f'{name} 삭제 완료'})

def _recalc_all():
    """인건비 변경 후 공정단가 & 전체 원가 재계산"""
    global P, all_costs, semi_costs, COMMON_RATE
    for k,m in PROC_META.items():
        P[k] = wsum(m['workers'],m['hc'])*m['hours']/m['capa']
    COMMON_RATE = calc_common_rate()
    for pn in products:
        all_costs[pn] = {'m': calc_cost(pn, True), 'r': calc_cost(pn, False)}
    for epn in semi_products:
        lab_m, li_m = calc_semi_labor(epn, True)
        lab_r, li_r = calc_semi_labor(epn, False)
        semi_costs[epn] = {'m': {'labor': lab_m, 'labor_items': li_m}, 'r': {'labor': lab_r, 'labor_items': li_r}}

# ============================================================
# 원부재료 단가 관리 API
# ============================================================
@app.route('/api/materials')
@login_required
def api_materials():
    """원부재료 단가 목록"""
    rows = []
    for pn, info in sorted(material_prices.items()):
        rows.append({'pn':pn, 'price':info['price'],
                     'date':str(info['date'])[:10] if info['date'] else '',
                     'name': _get_mat_name(pn)})
    return jsonify({'materials':rows})

@app.route('/api/materials/update', methods=['POST'])
@login_required
@admin_required
def api_material_update():
    """원부재료 단가 수정"""
    d = request.get_json()
    pn = (d.get('pn') or '').strip()
    price = float(d.get('price', 0))
    if not pn:
        return jsonify({'ok':False,'msg':'품번을 입력하세요'}), 400
    old = material_prices.get(pn, {}).get('price', 0)
    material_prices[pn] = {'price': price, 'date': datetime.now()}
    _recalc_all()
    return jsonify({'ok':True,'msg':f'{pn} 단가 변경: {old:,.0f} → {price:,.0f}원'})

@app.route('/api/materials/add', methods=['POST'])
@login_required
@admin_required
def api_material_add():
    """원부재료 신규 등록"""
    d = request.get_json()
    pn = (d.get('pn') or '').strip()
    price = float(d.get('price', 0))
    name = (d.get('name') or '').strip()
    if not pn:
        return jsonify({'ok':False,'msg':'품번을 입력하세요'}), 400
    material_prices[pn] = {'price': price, 'date': datetime.now()}
    _recalc_all()
    return jsonify({'ok':True,'msg':f'{pn} 등록 완료 (단가: {price:,.0f}원)'})

def _get_mat_name(pn):
    """BOM에서 해당 품번의 품명 찾기"""
    for items in bom_raw.values():
        for it in items:
            if it['ja_pn'] == pn:
                return it['ja_name']
    return ''

# ============================================================
# 인건비 검증 API
# ============================================================
verify_data = {'cost_report': {}}  # 원가보고서만 별도, 생산실적은 prod_records 공유

@app.route('/api/verify/upload_report', methods=['POST'])
@login_required
def api_verify_upload_report():
    """원가보고서 업로드 (단가기준정리.xlsx 내 원가보고서 시트 or 별도 파일)"""
    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'ok':False,'msg':'xlsx 파일만 업로드 가능합니다'}), 400
    tmp = f'_tmp_report_{datetime.now().strftime("%H%M%S")}.xlsx'
    f.save(tmp)
    try:
        wb = openpyxl.load_workbook(tmp, data_only=True)
        # 원가보고서 시트 찾기
        ws = None
        for sn in wb.sheetnames:
            if '원가' in sn and '보고' in sn:
                ws = wb[sn]; break
        if not ws:
            ws = wb[wb.sheetnames[0]]  # 첫 시트 사용
        report = {}
        for c in range(3, ws.max_column+1):
            month_name = ws.cell(1, c).value
            if not month_name: continue
            month_name = str(month_name).strip()
            report[month_name] = {
                'labor_total': ws.cell(14, c).value or 0,
                'salary': ws.cell(15, c).value or 0,
                'misc': ws.cell(16, c).value or 0,
                'retire': ws.cell(17, c).value or 0,
            }
        verify_data['cost_report'] = report
        os.remove(tmp)
        months = sorted(report.keys())
        return jsonify({'ok':True,'msg':f'원가보고서 로드 완료: {", ".join(months)}','months':months})
    except Exception as e:
        os.remove(tmp)
        return jsonify({'ok':False,'msg':f'파일 파싱 오류: {e}'}), 400

@app.route('/api/verify/run')
@login_required
def api_verify_run():
    """검증 실행"""
    month = request.args.get('month', '')
    cr = verify_data['cost_report'].get(month, {})
    recs = prod_records  # 공용 생산실적 사용

    # 월 필터
    month_num = month.replace('월','').strip()
    try: mn = int(month_num)
    except: mn = 0
    years = set(r['date'][:4] for r in recs if r['date'])
    year = max(years) if years else '2026'
    prefix = f"{year}-{mn:02d}" if mn > 0 else ''

    # 품번별 집계
    prod_by_pn = defaultdict(float)
    for r in recs:
        if prefix and not r['date'].startswith(prefix): continue
        prod_by_pn[r['pn']] += r['qty']

    # 실제인건비 집계 (일자별 생산실적 탭에서 입력한 actual_labor)
    total_actual = 0
    actual_count = 0
    for r in recs:
        if prefix and not r['date'].startswith(prefix): continue
        akey = actual_labor_key(r['date'], r['pn'], r['qty'])
        adata = actual_labor_data.get(akey)
        if adata:
            total_actual += calc_actual_labor(adata)
            actual_count += 1

    # 기준인건비 × 수량
    items = []
    total_std_m = total_std_r = 0
    for pn, qty in sorted(prod_by_pn.items()):
        name = ''; labor_m = labor_r = 0
        if pn in products:
            name = products[pn]['name']
            labor_m = all_costs.get(pn,{}).get('m',{}).get('labor',0)
            labor_r = all_costs.get(pn,{}).get('r',{}).get('labor',0)
        elif pn in semi_products:
            name = semi_products[pn]['name']
            labor_m = semi_costs.get(pn,{}).get('m',{}).get('labor',0)
            labor_r = semi_costs.get(pn,{}).get('r',{}).get('labor',0)
        sub_m = labor_m * qty; sub_r = labor_r * qty
        total_std_m += sub_m; total_std_r += sub_r
        items.append({'pn':pn,'name':name[:35],'qty':qty,
                      'labor_m':round(labor_m,2),'labor_r':round(labor_r,2),
                      'sub_m':round(sub_m,0),'sub_r':round(sub_r,0)})

    report_labor = cr.get('labor_total', 0)
    return jsonify({
        'month': month, 'report': cr, 'report_labor': report_labor,
        'items': items,
        'total_std_manual': round(total_std_m,0), 'total_std_rotary': round(total_std_r,0),
        'diff_manual': round(total_std_m - report_labor,0), 'diff_rotary': round(total_std_r - report_labor,0),
        'pass_manual': total_std_m >= report_labor if report_labor > 0 else None,
        'pass_rotary': total_std_r >= report_labor if report_labor > 0 else None,
        'total_actual': round(total_actual,0), 'actual_count': actual_count,
        'diff_actual': round(total_actual - report_labor,0) if total_actual > 0 else None,
        'pass_actual': total_actual >= report_labor if total_actual > 0 and report_labor > 0 else None,
        'available_months': sorted(verify_data['cost_report'].keys()),
        'prod_total': len(prod_records),
    })

# ============================================================
# 생산 진척도 & 월 결산 API
# ============================================================
report_store = {'plan_data':[], 'sales_data':[], 'plan_agg':{}, 'sales_agg':{}}

def _categorize(name):
    n = str(name).lower()
    if '오트밀' in n: return '오트밀류'
    if '누룽지' in n: return '누룽지류'
    if '카사바' in n: return '카사바류'
    if '바' in n and '고구마' in n and ('스틱' in n or '바 ' in n or '바20' in n or '바22' in n): return '고구마바류'
    if '스틱' in n and '고구마' in n: return '고구마스틱류'
    if '고구마' in n: return '고구마말랭이류'
    return '기타'

CAPA_MAP = {'고구마말랭이류':9310,'고구마스틱류':7000,'고구마바류':21728,'오트밀류':2200,'카사바류':7000,'누룽지류':0,'기타':0}

@app.route('/api/report/upload_plan', methods=['POST'])
@login_required
def api_upload_plan():
    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'ok':False,'msg':'xlsx 파일만 가능'}),400
    tmp = f'_tmp_plan.xlsx'; f.save(tmp)
    try:
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 7: continue
            pn = row[2]; qty = row[6]  # col3=idx2, col7=idx6
            if not pn or not isinstance(qty,(int,float)): continue
            rows.append({'pn':str(pn).strip(),'name':str(row[3] or '')[:40],'qty':qty,
                         'customer':str(row[1] or '')[:25]})
        wb.close()
        report_store['plan_data'] = rows
        # 사전 집계
        agg = defaultdict(lambda:{'name':'','qty':0})
        for r in rows:
            agg[r['pn']]['name'] = r['name']; agg[r['pn']]['qty'] += r['qty']
        report_store['plan_agg'] = dict(agg)
        os.remove(tmp)
        return jsonify({'ok':True,'msg':f'판매계획 {len(rows)}건 로드 ({len(agg)}품번)','count':len(rows)})
    except Exception as e:
        os.remove(tmp)
        return jsonify({'ok':False,'msg':str(e)}),400

@app.route('/api/report/upload_sales', methods=['POST'])
@login_required
def api_upload_sales():
    f = request.files.get('file')
    if not f or not f.filename.endswith('.xlsx'):
        return jsonify({'ok':False,'msg':'xlsx 파일만 가능'}),400
    tmp = f'_tmp_sales.xlsx'; f.save(tmp)
    try:
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        ws = wb[wb.sheetnames[0]]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if len(row) < 17: continue
            pn = row[10]; qty = row[16]  # 0-indexed: col11=idx10, col17=idx16
            if not pn or not isinstance(qty,(int,float)): continue
            rows.append({'pn':str(pn).strip(),'name':str(row[11] or '')[:40],'qty':qty,
                         'date':str(row[1] or '')[:10]})  # col2=idx1
        wb.close()
        report_store['sales_data'] = rows
        # 사전 집계 (월별)
        agg = defaultdict(lambda: defaultdict(lambda:{'name':'','qty':0}))  # {month: {pn: {name,qty}}}
        agg_all = defaultdict(lambda:{'name':'','qty':0})
        for r in rows:
            pn = r['pn']; m = r.get('date','')[:7]
            agg[m][pn]['name'] = r['name']; agg[m][pn]['qty'] += r['qty']
            agg_all[pn]['name'] = r['name']; agg_all[pn]['qty'] += r['qty']
        report_store['sales_agg'] = {'by_month':dict(agg), 'all':dict(agg_all)}
        os.remove(tmp)
        return jsonify({'ok':True,'msg':f'매출현황 {len(rows)}건 로드 ({len(agg_all)}품번)','count':len(rows)})
    except Exception as e:
        os.remove(tmp)
        return jsonify({'ok':False,'msg':str(e)}),400

@app.route('/api/report/progress')
@login_required
def api_report_progress():
    """생산 진척도: 판매계획 vs 생산실적"""
    remain_days = int(request.args.get('days',17))
    month = request.args.get('month','')  # YYYY-MM
    # 판매계획 (사전 집계 사용)
    plan_by_pn = report_store.get('plan_agg', {})
    # 생산실적 집계 (월 필터)
    prod_by_pn = defaultdict(float)
    for r in prod_records:
        if month and not r['date'].startswith(month): continue
        prod_by_pn[r['pn']] += r['qty']
    # 품목별
    items = []
    cat_agg = defaultdict(lambda:{'plan':0,'prod':0})
    for pn in sorted(plan_by_pn.keys()):
        if pn not in products: continue  # 자사 생산제품만
        p = plan_by_pn[pn]
        if p['qty'] <= 0: continue
        produced = prod_by_pn.get(pn,0)
        remain = max(p['qty'] - produced, 0)
        rate = produced / p['qty'] * 100 if p['qty']>0 else 0
        daily = remain / remain_days if remain_days>0 else 0
        cat = _categorize(p['name'])
        capa = CAPA_MAP.get(cat,0)
        status = 'done' if rate>=100 else ('ok' if rate>=50 else ('delay' if produced>0 else 'none'))
        danger = daily > capa if capa>0 else False
        cat_agg[cat]['plan'] += p['qty']
        cat_agg[cat]['prod'] += produced
        labor_m = all_costs.get(pn,{}).get('m',{}).get('labor',0)
        items.append({'pn':pn,'name':p['name'],'cat':cat,'plan':round(p['qty']),'prod':round(produced),
                      'remain':round(remain),'rate':round(rate,1),'daily':round(daily),
                      'capa':capa,'status':status,'danger':danger,'labor_ea':round(labor_m,2)})
    # 카테고리 요약
    cats = []
    for cat in ['고구마말랭이류','고구마스틱류','고구마바류','오트밀류','카사바류','누룽지류','기타']:
        d = cat_agg.get(cat)
        if not d or d['plan']==0: continue
        remain = d['plan']-d['prod']
        daily = remain/remain_days if remain_days>0 else 0
        capa = CAPA_MAP.get(cat,0)
        cats.append({'cat':cat,'plan':round(d['plan']),'prod':round(d['prod']),'remain':round(remain),
                     'rate':round(d['prod']/d['plan']*100,1),'daily':round(daily),'capa':capa,
                     'over':daily>capa if capa>0 else False})
    all_months = sorted(set(r['date'][:7] for r in prod_records if r['date'] and len(r['date'])>=7))
    return jsonify({'items':items,'categories':cats,'remain_days':remain_days,
                    'plan_count':len(report_store.get('plan_data',[])),'months':all_months,'selected_month':month})

@app.route('/api/report/settlement')
@login_required
def api_report_settlement():
    """월 결산: 매출현황 vs 생산실적"""
    month = request.args.get('month','')  # YYYY-MM
    # 매출 집계 (사전 집계 사용)
    sa = report_store.get('sales_agg', {})
    if month and month in sa.get('by_month',{}):
        sales_by_pn = {k:v for k,v in sa['by_month'][month].items() if k in products}
    elif not month and sa.get('all'):
        sales_by_pn = {k:v for k,v in sa['all'].items() if k in products}
    else:
        sales_by_pn = {}
    # 생산실적 집계 (월 필터)
    prod_by_pn = defaultdict(lambda:{'name':'','qty':0,'erp':0})
    for r in prod_records:
        if month and not r['date'].startswith(month): continue
        prod_by_pn[r['pn']]['name'] = r['name']
        prod_by_pn[r['pn']]['qty'] += r['qty']
        prod_by_pn[r['pn']]['erp'] += r['erp_price'] * r['qty']
    # 합집합
    all_pn = sorted(set(list(sales_by_pn.keys()) + [k for k in prod_by_pn if k in products]))
    items = []; over=[]; under=[]
    for pn in all_pn:
        s = sales_by_pn.get(pn,{}).get('qty',0)
        p = prod_by_pn.get(pn,{}).get('qty',0)
        name = sales_by_pn.get(pn,{}).get('name','') or prod_by_pn.get(pn,{}).get('name','')
        diff = p - s
        ratio = p/s*100 if s>0 else (999 if p>0 else 0)
        cat = _categorize(name)
        labor_m = all_costs.get(pn,{}).get('m',{}).get('labor',0)
        std_total = round(labor_m * p)
        status = 'over' if ratio>150 else ('under' if 0<ratio<80 else 'normal')
        it = {'pn':pn,'name':name[:35],'cat':cat,'prod':round(p),'sales':round(s),'diff':round(diff),
              'ratio':round(ratio,1),'status':status,'labor_ea':round(labor_m,2),'std_labor_total':std_total}
        items.append(it)
        if status=='over' and s>0: over.append(it)
        if status=='under' and s>100: under.append(it)
    over.sort(key=lambda x:-x['ratio']); under.sort(key=lambda x:x['ratio'])
    all_months_p = sorted(set(r['date'][:7] for r in prod_records if r['date'] and len(r['date'])>=7))
    all_months_s = sorted(set(r['date'][:7] for r in report_store.get('sales_data',[]) if r.get('date','') and len(r['date'])>=7))
    all_months = sorted(set(all_months_p + all_months_s))
    return jsonify({'items':items,'over_stock':over[:10],'under_stock':under[:10],
                    'total_prod':sum(i['prod'] for i in items),'total_sales':sum(i['sales'] for i in items),
                    'sales_count':len(report_store.get('sales_data',[])),'months':all_months,'selected_month':month})

# ============================================================
# HTML
# ============================================================
HTML = r"""
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>매홍엘앤에프 통합 원가 관리 시스템</title>
<style>
:root{--primary:#1a365d;--accent:#2b6cb0;--good:#38a169;--bad:#e53e3e;--warn:#d69e2e;--bg:#f7fafc;--card:#fff;--border:#e2e8f0;--text:#2d3748;--muted:#718096}
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Pretendard','맑은 고딕',sans-serif;background:var(--bg);color:var(--text);font-size:14px}
.header{background:linear-gradient(135deg,#0055a5 0%,#0077cc 50%,#5cb85c 85%,#7BC142 100%);color:#fff;padding:18px 32px;display:flex;align-items:center;justify-content:space-between}
.header h1{font-size:22px;font-weight:700;letter-spacing:-0.5px}
.header .sub{font-size:12px;opacity:.8}
.tabs{display:flex;gap:0;background:#fff;border-bottom:2px solid var(--border);padding:0 32px;position:sticky;top:0;z-index:100;height:50px;align-items:stretch}
.tab{padding:0 24px;cursor:pointer;font-weight:600;color:var(--muted);border-bottom:3px solid transparent;transition:.2s;display:flex;align-items:center;line-height:50px}
.tab:hover{color:var(--text)}.tab.active{color:var(--accent);border-bottom-color:var(--accent)}
.panel{display:none;padding:24px 32px;max-width:1600px;margin:0 auto}.panel.active{display:block}
.card{background:var(--card);border-radius:12px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:20px 24px;margin-bottom:20px}
.card h3{font-size:15px;color:var(--primary);margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid var(--border)}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:16px;margin-bottom:24px}
.stat{background:var(--card);border-radius:10px;padding:18px 20px;box-shadow:0 1px 3px rgba(0,0,0,.06);border-left:4px solid var(--accent)}
.stat.good{border-left-color:var(--good)}.stat.bad{border-left-color:var(--bad)}.stat.warn{border-left-color:var(--warn)}
.stat .label{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.5px}
.stat .value{font-size:24px;font-weight:700;margin-top:4px}
.stat .detail{font-size:11px;color:var(--muted);margin-top:2px}
table{width:100%;border-collapse:collapse;font-size:13px}
th{background:var(--primary);color:#fff;padding:10px 12px;text-align:center;font-weight:600}
td{padding:8px 12px;border-bottom:1px solid var(--border)}
tr:hover{background:#edf2f7}
.r{text-align:right}.c{text-align:center}
.highlight{background:#fefcbf !important}
.good-bg{background:#f0fff4}.bad-bg{background:#fff5f5}
.badge{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
.badge-good{background:#c6f6d5;color:#22543d}.badge-bad{background:#fed7d7;color:#9b2c2c}.badge-warn{background:#fefcbf;color:#744210}
.search{padding:8px 14px;border:1px solid var(--border);border-radius:8px;width:300px;font-size:13px;margin-bottom:16px}
.compare-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px}
@media(max-width:900px){.compare-grid{grid-template-columns:1fr}.stats{grid-template-columns:1fr 1fr}}
.footer{text-align:center;padding:20px;color:var(--muted);font-size:12px}

/* 클릭 가능 셀 */
.clickable{cursor:pointer;transition:.15s}
.clickable:hover{background:#ebf5fb;color:var(--accent);text-decoration:underline}

/* 모달 */
.modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:1000;justify-content:center;align-items:flex-start;padding:30px 20px;overflow-y:auto}
.modal-overlay.open{display:flex}
.modal{background:#fff;border-radius:12px;max-width:960px;width:100%;box-shadow:0 20px 60px rgba(0,0,0,.3);animation:slideIn .2s ease}
@keyframes slideIn{from{opacity:0;transform:translateY(-16px)}to{opacity:1;transform:translateY(0)}}
.modal-header{padding:20px 28px 16px;border-bottom:1px solid var(--border)}
.modal-header h2{font-size:15px;font-weight:700;color:var(--text)}
.modal-header .mh-sub{font-size:12px;color:var(--muted);margin-top:2px}
.modal-close{position:absolute;right:20px;top:18px;background:none;border:none;font-size:22px;cursor:pointer;color:var(--muted);padding:4px 8px;border-radius:6px}
.modal-close:hover{background:#f0f0f0;color:var(--text)}
.modal-body{padding:0 28px 28px;max-height:75vh;overflow-y:auto}
.modal-body table{font-size:12.5px;margin-top:0}
.modal-body th{position:static;font-size:11.5px;padding:9px 12px;background:#f8f9fa;color:var(--text);font-weight:600;border-bottom:2px solid var(--border)}
.modal-body td{padding:8px 12px;border-bottom:1px solid #f0f0f0}
.modal-body tr:last-child td{border-bottom:none}

/* 모달 상단 요약 카드 */
.m-stats{display:flex;gap:0;margin:20px 0 16px;border:1px solid var(--border);border-radius:10px;overflow:hidden}
.m-stat{flex:1;padding:14px 16px;border-right:1px solid var(--border);text-align:center}
.m-stat:last-child{border-right:none}
.m-stat .ms-label{font-size:10.5px;color:var(--muted);letter-spacing:.3px}
.m-stat .ms-value{font-size:18px;font-weight:700;margin-top:3px}
.m-stat .ms-unit{font-size:11px;color:var(--muted);font-weight:400}
.m-stat.accent{background:#f0f7ff}
.m-stat.total{background:var(--primary)}.m-stat.total .ms-label,.m-stat.total .ms-value,.m-stat.total .ms-unit{color:#fff}

/* 그룹 헤더 */
.group-row td{background:#f8f9fa;font-weight:700;font-size:12px;color:var(--primary);border-bottom:1px solid var(--border);padding:10px 12px}
.total-row td{background:#f8f9fa;font-weight:700;border-top:2px solid var(--border);padding:10px 12px}

/* 인건비 공정 카드 */
.proc-card{background:#fafbfc;border:1px solid #edf0f2;border-radius:10px;padding:16px 18px;margin:12px 0}
.proc-card .pc-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px}
.proc-card .pc-title{font-size:13px;font-weight:700;color:var(--primary)}
.proc-card .pc-result{font-size:15px;font-weight:700;color:var(--accent)}
.proc-card .pc-workers{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px}
.worker-tag{display:inline-flex;align-items:center;gap:4px;background:#fff;border:1px solid #e0e4e8;border-radius:6px;padding:3px 8px;font-size:11px}
.worker-tag .wt-name{font-weight:600;color:var(--text)}.worker-tag .wt-rate{color:var(--accent)}
.worker-tag.min{background:#fffbeb;border-color:#f0d78c}.worker-tag.min .wt-name{color:#92710c}
.proc-card .pc-formula{font-size:12px;color:var(--muted);line-height:1.7;font-family:'Consolas','맑은 고딕',monospace;background:#fff;border:1px solid #eee;border-radius:6px;padding:10px 12px}
.pc-formula .pf-line{margin:1px 0}
.pc-formula .pf-result{color:var(--primary);font-weight:700;font-size:13px;margin-top:4px}

/* 합계 박스 */
.labor-total{background:var(--primary);color:#fff;border-radius:10px;padding:18px 20px;margin-top:16px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px}
.labor-total .lt-items{font-size:11px;opacity:.8;flex:1}
.labor-total .lt-sum{font-size:26px;font-weight:700}
.labor-total .lt-sub{font-size:12px;opacity:.7;margin-top:2px}

/* 비율 바 */
.ratio-bar{display:flex;height:28px;border-radius:6px;overflow:hidden;margin-top:12px}
.ratio-bar div{display:flex;align-items:center;justify-content:center;font-size:10px;color:#fff;font-weight:600;min-width:30px}
</style>
</head>
<body>

<div class="header">
  <div style="display:flex;align-items:center;gap:16px">
    <img src="/logo.png" alt="Maehong L&F" style="height:48px;border-radius:6px;background:#fff;padding:4px 8px">
    <div>
      <h1>통합 원가 관리 시스템</h1>
      <div class="sub">제조원가 분석 대시보드 | 기준일 {{ now }} | 최저시급 {{ min_wage }}원</div>
    </div>
  </div>
  <div style="text-align:right;display:flex;align-items:center;gap:16px">
    <div>
      <div style="font-size:28px;font-weight:700">{{ total_products }}개</div>
      <div class="sub">관리 품목</div>
    </div>
    <div style="border-left:1px solid rgba(255,255,255,.3);padding-left:16px">
      <div style="font-size:13px">{{ user_name }} {% if is_admin %}<span style="background:rgba(255,255,255,.2);padding:1px 6px;border-radius:4px;font-size:10px">ADMIN</span>{% endif %}</div>
      <a href="/logout" style="color:rgba(255,255,255,.7);font-size:11px;text-decoration:none">로그아웃</a>
    </div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" onclick="showTab(0)">품목별 기준원가표</div>
  <div class="tab" onclick="showTab(1)">일자별 생산실적</div>
  <div class="tab" onclick="showTab(2)">이슈 분석</div>
  <div class="tab" onclick="showTab(3)">G0010 검증</div>
  <div class="tab" onclick="showTab(4)">인건비 검증</div>
  <div class="tab" onclick="showTab(5)">생산 진척도</div>
  <div class="tab" onclick="showTab(6)">월 결산</div>
  {% if is_admin %}<div class="tab" onclick="showTab(7)">설정/관리</div>{% endif %}
</div>

<!-- 탭1: 품목별 기준원가표 -->
<div class="panel active" id="p0">
  <div class="stats">
    <div class="stat"><div class="label">총 품목수</div><div class="value">{{ total_products }}</div></div>
    {% if is_admin %}<div class="stat warn"><div class="label">평균 인건비(수작업)</div><div class="value">{{ avg_labor }}원</div></div>{% endif %}
    <div class="stat"><div class="label">평균 원부재료비</div><div class="value">{{ avg_mat }}원</div></div>
    <div class="stat good"><div class="label">평균 기준원가</div><div class="value">{{ avg_total }}원</div></div>
  </div>
  <div class="card">
    <h3>전 품목 기준원가표 ({{ total_products }}개)  <span style="font-size:12px;color:var(--muted);font-weight:400">— 품번/품명, 원재료비, 부재료비, 인건비를 클릭하면 상세 계산식이 표시됩니다</span></h3>
    <input class="search" type="text" placeholder="품번 또는 품명 검색..." onkeyup="filterTable(this,'t1')">
    <div style="overflow-x:auto">
    <table id="t1">
      <thead><tr>
        <th>품번</th><th>품명</th><th>카테고리</th><th>중량</th><th>유형</th>
        <th>원재료비</th><th>부재료비</th>
        {% if is_admin %}<th>인건비(수작업)</th><th>인건비(로터리)</th>{% endif %}
        <th>기준원가(수작업)</th><th>기준원가(로터리)</th>
      </tr></thead>
      <tbody>
      {% for p in cost_rows %}
      <tr class="{{ 'highlight' if p.pn == 'G0010' else '' }}">
        <td class="c clickable" onclick="openBom('{{p.pn}}')" title="클릭: BOM 상세">{{ p.pn }}</td>
        <td class="clickable" onclick="openBom('{{p.pn}}')" title="클릭: BOM 상세">{{ p.name[:42] }}</td>
        <td class="c">{{ p.cat }}</td>
        <td class="r">{{ "{:,}".format(p.weight|int) }}g</td>
        <td class="c"><span class="badge {{ 'badge-warn' if p.ptype=='번들' else 'badge-good' }}">{{ p.ptype }}</span></td>
        <td class="r clickable" onclick="openMaterial('{{p.pn}}')" title="클릭: 원재료비 계산식" style="color:#c0392b">{{ "{:,.0f}".format(p.raw) }}</td>
        <td class="r clickable" onclick="openMaterial('{{p.pn}}')" title="클릭: 부재료비 계산식" style="color:#8e44ad">{{ "{:,.0f}".format(p.sub) }}</td>
        {% if is_admin %}
        <td class="r clickable" onclick="openLabor('{{p.pn}}')" title="클릭: 인건비 계산식" style="font-weight:600;color:#2471a3">{{ "{:,.0f}".format(p.labor_m) }}</td>
        <td class="r clickable" onclick="openLabor('{{p.pn}}')" title="클릭: 인건비 계산식" style="color:#2471a3">{{ "{:,.0f}".format(p.labor_r) }}</td>
        {% endif %}
        <td class="r" style="font-weight:700;color:var(--primary)">{{ "{:,.0f}".format(p.total_m) }}</td>
        <td class="r">{{ "{:,.0f}".format(p.total_r) }}</td>
      </tr>
      {% endfor %}
      </tbody>
    </table>
    </div>
  </div>
</div>

<!-- 탭2: 일자별 생산실적 -->
<div class="panel" id="p1">
  <!-- 업로드 + 필터 -->
  <div class="card">
    <h3>생산실적 데이터 관리</h3>
    <div style="display:flex;gap:16px;flex-wrap:wrap;align-items:flex-end">
      <!-- 업로드 -->
      <div style="flex:1;min-width:280px">
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">생산실적 파일 업로드 (.xlsx)</label>
        <div style="display:flex;gap:8px;align-items:center">
          <input type="file" id="prodFileInput" accept=".xlsx" style="font-size:12px;flex:1">
          <select id="prodMode" style="padding:7px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
            <option value="append">누적 추가</option>
            <option value="replace">전체 교체</option>
          </select>
          <button onclick="uploadProd()" style="padding:7px 16px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;white-space:nowrap">업로드</button>
        </div>
        <div id="prodUploadStatus" style="margin-top:6px;font-size:12px"></div>
      </div>
      <!-- 필터 -->
      <div style="display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap">
        <div>
          <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">시작일</label>
          <input type="date" id="prodFrom" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
        </div>
        <div>
          <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">종료일</label>
          <input type="date" id="prodTo" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
        </div>
        <div>
          <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">품번/품명 검색</label>
          <input type="text" id="prodSearch" placeholder="G0010, 고구마..." style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px;width:160px">
        </div>
        <button onclick="loadProdRecords()" style="padding:7px 16px;background:var(--primary);color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">조회</button>
        <button onclick="document.getElementById('prodFrom').value='';document.getElementById('prodTo').value='';document.getElementById('prodSearch').value='';loadProdRecords()" style="padding:7px 12px;background:#eee;color:var(--text);border:none;border-radius:6px;font-size:12px;cursor:pointer">초기화</button>
      </div>
    </div>
  </div>

  <!-- 요약 -->
  <div class="stats" id="prodStats"></div>

  <!-- 실적 테이블 (일자별 그룹) -->
  <div class="card">
    <h3 id="prodTableTitle">생산실적</h3>
    <div style="overflow-x:auto">
      <table id="prodTable">
        <thead><tr>
          <th>일자</th><th>품번</th><th>품명</th><th>구분</th><th>실적수량</th><th>ERP단가</th>
          {% if is_admin %}<th>인건비/EA(수작업)</th><th>인건비/EA(로터리)</th><th>실제인건비/EA</th>{% endif %}<th>기준 EA/MH</th><th>실제 EA/MH</th><th>생산성</th><th>비고</th>
        </tr></thead>
        <tbody id="prodBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- 탭3: 이슈 분석 -->
<div class="panel" id="p2">
  {% if is_admin %}
  <div class="compare-grid">
    <div class="card">
      <h3>수작업 vs 로터리 내포장</h3>
      <table>
        <tr><th>구분</th><th>인원</th><th>CAPA</th><th>인건비/EA</th><th>비율</th></tr>
        <tr><td class="c">로터리</td><td class="c">3명</td><td class="r">12,000</td>
          <td class="r">{{ "{:,.2f}".format(p_rot) }}원</td><td class="c">기준</td></tr>
        <tr class="bad-bg"><td class="c"><b>수작업</b></td><td class="c">7명</td><td class="r">9,310</td>
          <td class="r"><b>{{ "{:,.2f}".format(p_man) }}원</b></td>
          <td class="c"><span class="badge badge-bad">{{ "{:.0f}".format(p_man/p_rot*100) }}%</span></td></tr>
      </table>
      <div style="margin-top:12px;padding:12px;background:#fff5f5;border-radius:8px;font-size:13px">
        수작업 전환 시 내포장 인건비 <b>+{{ "{:.1f}".format((p_man/p_rot-1)*100) }}%</b> 상승
      </div>
    </div>
    <div class="card">
      <h3>낱봉 vs 번들 외포장</h3>
      <table>
        <tr><th>구분</th><th>CAPA</th><th>인건비/EA</th><th>차이</th></tr>
        <tr><td class="c">낱봉</td><td class="r">14,700</td>
          <td class="r">{{ "{:,.2f}".format(p_naet) }}원</td><td class="c">기준</td></tr>
        <tr><td class="c">번들</td><td class="r">13,230</td>
          <td class="r">{{ "{:,.2f}".format(p_bund) }}원</td>
          <td class="c"><span class="badge badge-warn">+{{ "{:.1f}".format((p_bund/p_naet-1)*100) }}%</span></td></tr>
      </table>
    </div>
  </div>
  <div class="card">
    <h3>주요 고구마류 품목 - 수작업 전환 시 인건비 영향</h3>
    <table>
      <thead><tr><th>품번</th><th>품명</th><th>유형</th><th>로터리</th><th>수작업</th><th>차이</th><th>상승률</th></tr></thead>
      <tbody>
      {% for p in issue_rows %}
      <tr class="{{ 'bad-bg' if p.diff > 100 else '' }}">
        <td class="c clickable" onclick="openLabor('{{p.pn}}')">{{ p.pn }}</td>
        <td>{{ p.name[:32] }}</td><td class="c">{{ p.ptype }}</td>
        <td class="r">{{ "{:,.0f}".format(p.labor_r) }}</td>
        <td class="r"><b>{{ "{:,.0f}".format(p.labor_m) }}</b></td>
        <td class="r" style="color:var(--bad);font-weight:700">+{{ "{:,.0f}".format(p.diff) }}</td>
        <td class="c"><span class="badge badge-bad">+{{ "{:.0f}".format(p.pct) }}%</span></td>
      </tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
  {% endif %}
  {% if is_admin %}
  <div class="card">
    <h3>공정별 인원 시급 현황</h3>
    <table>
      <thead><tr><th>공정</th><th>인원명</th><th>지급총액</th><th>시급</th><th>비고</th></tr></thead>
      <tbody>
      {% for w in wage_rows %}
      <tr><td class="c">{{ w.proc }}</td><td class="c">{{ w.name }}</td>
        <td class="r">{{ "{:,.0f}".format(w.pay) }}원</td>
        <td class="r">{{ "{:,.0f}".format(w.hourly) }}원</td>
        <td>{{ w.note }}</td></tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
  {% endif %}
</div>

<!-- 탭4: G0010 검증 -->
<div class="panel" id="p3">
  <div class="stats">
    <div class="stat clickable" onclick="openMaterial('G0010')"><div class="label">원재료비</div><div class="value">{{ "{:,.0f}".format(g_raw) }}원</div><div class="detail">클릭: 상세보기</div></div>
    <div class="stat clickable" onclick="openMaterial('G0010')"><div class="label">부재료비</div><div class="value">{{ "{:,.0f}".format(g_sub) }}원</div><div class="detail">클릭: 상세보기</div></div>
    <div class="stat warn clickable" onclick="openLabor('G0010')"><div class="label">인건비(수작업)</div><div class="value">{{ "{:,.2f}".format(g_labor) }}원</div><div class="detail">목표 676.48원 | 클릭: 상세보기</div></div>
    <div class="stat good"><div class="label">기준원가 합계</div><div class="value">{{ "{:,.0f}".format(g_total) }}원</div></div>
  </div>
  <div class="compare-grid">
    <div class="card">
      <h3>BOM 3단계 폭발 <span style="font-size:12px;color:var(--muted);cursor:pointer" onclick="openBom('G0010')">[트리 보기]</span></h3>
      <table>
        <tr><th>구분</th><th>품번</th><th>품명</th><th>수량</th><th>단가</th><th>금액</th></tr>
        {% for m in g_mats %}
        <tr><td class="c">{{ m.type }}</td><td class="c">{{ m.pn }}</td><td>{{ m.name[:28] }}</td>
          <td class="r">{{ "{:.4f}".format(m.qty) }}{{ m.unit }}</td>
          <td class="r">{{ "{:,.0f}".format(m.price) }}</td>
          <td class="r" style="font-weight:600">{{ "{:,.2f}".format(m.cost) }}</td></tr>
        {% endfor %}
      </table>
    </div>
    <div class="card">
      <h3>공정별 인건비 <span style="font-size:12px;color:var(--muted);cursor:pointer" onclick="openLabor('G0010')">[계산식 보기]</span></h3>
      <table>
        <tr><th>공정</th><th>인건비</th><th>비율</th></tr>
        {% for li in g_labor_items %}
        <tr><td>{{ li.proc }}</td>
          <td class="r" style="font-weight:600">{{ "{:,.2f}".format(li.cost) }}원</td>
          <td class="r">{{ "{:.1f}".format(li.pct) }}%</td></tr>
        {% endfor %}
        <tr style="background:var(--primary);color:#fff;font-weight:700">
          <td>합계</td><td class="r">{{ "{:,.2f}".format(g_labor) }}원</td><td class="r">100%</td></tr>
      </table>
    </div>
  </div>
</div>

<!-- 탭5: 인건비 검증 -->
<div class="panel" id="p4">
  <!-- 파일 업로드 -->
  <div class="compare-grid">
    <div class="card">
      <h3>1. 원가보고서 업로드</h3>
      <p style="font-size:12px;color:var(--muted);margin-bottom:10px">단가기준정리.xlsx (원가보고서 시트 포함) 또는 원가보고서 전용 파일</p>
      <div style="display:flex;gap:8px;align-items:center">
        <input type="file" id="vReportFile" accept=".xlsx" style="font-size:12px;flex:1">
        <button onclick="uploadVReport()" style="padding:7px 14px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer;font-weight:600">업로드</button>
      </div>
      <div id="vReportStatus" style="margin-top:6px;font-size:12px"></div>
    </div>
    <div class="card">
      <h3>2. 생산실적</h3>
      <p style="font-size:12px;color:var(--muted)">[일자별 생산실적] 탭에서 업로드한 데이터가 자동으로 사용됩니다. (현재 {{ prod_count }}건)</p>
    </div>
  </div>

  <!-- 검증 실행 -->
  <div class="card">
    <h3>3. 검증 실행</h3>
    <div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap">
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">검증 월</label>
        <select id="verifyMonth" style="padding:7px 12px;border:1px solid var(--border);border-radius:6px;font-size:13px">
          <option value="">-- 원가보고서 업로드 후 선택 --</option>
        </select>
      </div>
      <button onclick="runVerify()" style="padding:8px 24px;background:var(--primary);color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">검증 실행</button>
    </div>
  </div>

  <!-- 검증 결과 -->
  <div id="verifyResult"></div>
</div>

<!-- 탭6: 생산 진척도 -->
<div class="panel" id="p5">
  <div class="card">
    <h3>생산 진척도</h3>
    <p style="font-size:12px;color:var(--muted);margin-bottom:12px">생산실적은 [일자별 생산실적] 탭에서 업로드하면 자동 반영됩니다.</p>
    <div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap">
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">판매계획현황 (.xlsx)</label>
        <input type="file" id="planFile" accept=".xlsx" style="font-size:12px">
      </div>
      <button onclick="uploadPlan()" style="padding:7px 14px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer;font-weight:600">업로드</button>
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">연도</label>
        <select id="progressYear" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
          <option value="2025">2025</option><option value="2026" selected>2026</option><option value="2027">2027</option>
        </select>
      </div>
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">월</label>
        <select id="progressMonth" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
          <option value="">전체</option>
          <option value="01">1월</option><option value="02">2월</option><option value="03">3월</option>
          <option value="04">4월</option><option value="05">5월</option><option value="06">6월</option>
          <option value="07">7월</option><option value="08">8월</option><option value="09">9월</option>
          <option value="10">10월</option><option value="11">11월</option><option value="12">12월</option>
        </select>
      </div>
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">남은 영업일</label>
        <input type="number" id="remainDays" value="17" min="1" max="30" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:13px;width:70px">
      </div>
      <button onclick="runProgress()" style="padding:7px 18px;background:var(--primary);color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer;font-weight:700">분석 실행</button>
      <span id="planStatus" style="font-size:12px"></span>
    </div>
  </div>
  <div id="progressResult"></div>
</div>

<!-- 탭7: 월 결산 -->
<div class="panel" id="p6">
  <div class="card">
    <h3>월 결산</h3>
    <p style="font-size:12px;color:var(--muted);margin-bottom:12px">생산실적은 [일자별 생산실적] 탭에서 업로드하면 자동 반영됩니다.</p>
    <div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap">
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">매출마감현황 (.xlsx)</label>
        <input type="file" id="salesFile" accept=".xlsx" style="font-size:12px">
      </div>
      <button onclick="uploadSales()" style="padding:7px 14px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer;font-weight:600">업로드</button>
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">연도</label>
        <select id="settlementYear" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
          <option value="2025">2025</option><option value="2026" selected>2026</option><option value="2027">2027</option>
        </select>
      </div>
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">월</label>
        <select id="settlementMonth" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px">
          <option value="">전체</option>
          <option value="01">1월</option><option value="02">2월</option><option value="03">3월</option>
          <option value="04">4월</option><option value="05">5월</option><option value="06">6월</option>
          <option value="07">7월</option><option value="08">8월</option><option value="09">9월</option>
          <option value="10">10월</option><option value="11">11월</option><option value="12">12월</option>
        </select>
      </div>
      <button onclick="runSettlement()" style="padding:7px 18px;background:var(--primary);color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer;font-weight:700">분석 실행</button>
      <span id="salesStatus" style="font-size:12px"></span>
    </div>
  </div>
  <div id="settlementResult"></div>
</div>

<!-- 탭8: 설정/관리 (admin only) -->
{% if is_admin %}
<div class="panel" id="p7">
  <div class="compare-grid">
    <!-- 파일 업로드 -->
    <div class="card">
      <h3>단가기준정리 파일 업데이트</h3>
      <p style="font-size:13px;color:var(--muted);margin-bottom:16px">
        단가기준정리.xlsx 파일을 새로 업로드하면 기존 파일은 자동 백업됩니다.<br>
        업로드 후 서버를 재시작해야 데이터가 반영됩니다.
      </p>
      <div id="uploadArea" style="border:2px dashed var(--border);border-radius:10px;padding:32px;text-align:center;cursor:pointer;transition:.2s;background:#fafbfc"
           ondragover="event.preventDefault();this.style.borderColor='var(--accent)';this.style.background='#f0f7ff'"
           ondragleave="this.style.borderColor='var(--border)';this.style.background='#fafbfc'"
           ondrop="handleDrop(event)" onclick="document.getElementById('fileInput').click()">
        <div style="font-size:36px;margin-bottom:8px;opacity:.4">📁</div>
        <div style="font-size:14px;font-weight:600;color:var(--text)">파일을 드래그하거나 클릭하여 선택</div>
        <div style="font-size:12px;color:var(--muted);margin-top:4px">.xlsx 파일만 업로드 가능</div>
      </div>
      <input type="file" id="fileInput" accept=".xlsx" style="display:none" onchange="handleFile(this.files[0])">
      <div id="uploadStatus" style="margin-top:12px"></div>
    </div>

    <!-- 인건비 관리 -->
    <div class="card">
      <h3>인건비 신규 등록</h3>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px">
        <div>
          <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">사원명</label>
          <input id="empName" type="text" placeholder="이름 입력" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:8px;font-size:13px">
        </div>
        <div>
          <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">지급총액 (원/월)</label>
          <input id="empPay" type="number" placeholder="예: 2500000" style="width:100%;padding:8px 12px;border:1px solid var(--border);border-radius:8px;font-size:13px">
        </div>
      </div>
      <div style="display:flex;align-items:center;gap:12px;margin-bottom:16px">
        <label style="font-size:13px;display:flex;align-items:center;gap:6px;cursor:pointer">
          <input id="empCommon" type="checkbox" style="width:16px;height:16px"> 공통배부
        </label>
        <button onclick="addEmployee()" style="padding:8px 20px;background:var(--accent);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer">추가</button>
      </div>
      <div id="empAddStatus" style="margin-bottom:8px"></div>
    </div>
  </div>

  <!-- 인건비 현황 테이블 -->
  <div class="card">
    <h3>인건비 현황 ({{ emp_count }}명) <span style="font-size:12px;color:var(--muted);font-weight:400">— 지급총액을 직접 수정하고 저장 버튼을 누르세요</span></h3>
    <div style="overflow-x:auto">
    <table id="empTable" style="table-layout:fixed;width:100%">
      <colgroup>
        <col style="width:60px"><col style="width:140px"><col style="width:180px">
        <col style="width:140px"><col style="width:100px"><col style="width:160px">
      </colgroup>
      <thead><tr>
        <th>No</th><th>사원명</th><th>지급총액(원)</th>
        <th>시급(원)</th><th>공통배부</th><th>관리</th>
      </tr></thead>
      <tbody id="empBody"></tbody>
    </table>
    </div>
  </div>

  <!-- 원부재료 단가 관리 -->
  <div class="card">
    <h3>원부재료 단가 관리 <span style="font-size:12px;color:var(--muted);font-weight:400">— 단가를 직접 수정하고 저장 버튼을 누르세요</span></h3>
    <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;padding:12px;background:#f8f9fa;border-radius:8px">
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">품번</label>
        <input id="matPn" type="text" placeholder="예: A0027" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:13px;width:100px">
      </div>
      <div>
        <label style="font-size:12px;color:var(--muted);display:block;margin-bottom:4px">단가 (원)</label>
        <input id="matPrice" type="number" placeholder="예: 3345" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:13px;width:120px">
      </div>
      <div style="display:flex;align-items:flex-end">
        <button onclick="addMaterial()" style="padding:7px 16px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">추가/수정</button>
      </div>
      <div style="display:flex;align-items:flex-end;flex:1">
        <input id="matSearch" type="text" placeholder="품번/품명 검색..." oninput="filterMatTable()" style="padding:6px 10px;border:1px solid var(--border);border-radius:6px;font-size:12px;width:200px">
      </div>
      <div id="matStatus" style="display:flex;align-items:flex-end;font-size:12px"></div>
    </div>
    <div style="overflow-x:auto;max-height:400px;overflow-y:auto">
    <table id="matTable" style="table-layout:fixed;width:100%">
      <colgroup>
        <col style="width:50px"><col style="width:80px"><col style="width:auto">
        <col style="width:130px"><col style="width:90px"><col style="width:110px">
      </colgroup>
      <thead><tr>
        <th>No</th><th>품번</th><th>품명</th><th>단가(원)</th><th>단가일</th><th>관리</th>
      </tr></thead>
      <tbody id="matBody"></tbody>
    </table>
    </div>
  </div>
</div>
{% endif %}

<div class="footer">매홍엘앤에프 통합 원가 관리 시스템 v2.0</div>

<!-- 모달 -->
<div class="modal-overlay" id="modalOverlay" onclick="if(event.target===this)closeModal()">
  <div class="modal" style="position:relative">
    <div class="modal-header">
      <h2 id="modalTitle">상세 정보</h2>
      <div class="mh-sub" id="modalSub"></div>
      <button class="modal-close" onclick="closeModal()">&times;</button>
    </div>
    <div class="modal-body" id="modalBody">로딩 중...</div>
  </div>
</div>

<script>
const _isAdmin = {{ 'true' if is_admin else 'false' }};
const fmt = n => n==null||n===''?'-':Number(n).toLocaleString('ko-KR',{maximumFractionDigits:2});
const fmt0 = n => n==null?'-':Number(n).toLocaleString('ko-KR',{maximumFractionDigits:0});
const fmt4 = n => n==null?'-':Number(n).toLocaleString('ko-KR',{minimumFractionDigits:4,maximumFractionDigits:4});

function showTab(n){
  document.querySelectorAll('.tab').forEach((t,i)=>{t.classList.toggle('active',i===n)});
  document.querySelectorAll('.panel').forEach((p,i)=>{p.classList.toggle('active',i===n)});
}
function filterTable(input,tid){
  const v=input.value.toLowerCase();
  document.querySelectorAll('#'+tid+' tbody tr').forEach(r=>{
    r.style.display=r.textContent.toLowerCase().includes(v)?'':'none';
  });
}
function openModal(title,sub,html){
  document.getElementById('modalTitle').textContent=title;
  document.getElementById('modalSub').textContent=sub||'';
  document.getElementById('modalBody').innerHTML=html;
  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow='hidden';
}
function closeModal(){
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow='';
}
document.addEventListener('keydown',e=>{if(e.key==='Escape')closeModal()});

// =========================================
// BOM 모달
// =========================================
async function openBom(pn){
  openModal(pn,'로딩 중...','<div style="text-align:center;padding:60px;color:var(--muted)">데이터 조회 중...</div>');
  const [bomRes, matRes, labRes] = await Promise.all([
    fetch('/api/bom/'+pn).then(r=>r.json()),
    fetch('/api/material/'+pn).then(r=>r.json()),
    fetch('/api/labor/'+pn).then(r=>r.json()),
  ]);
  const d=bomRes, md=matRes, ld=labRes, p=d.product;
  const rawItems=md.mat_items.filter(m=>m.type==='원재료');
  const subItems=md.mat_items.filter(m=>m.type==='부재료');
  const bomCount=md.mat_items.length;

  document.getElementById('modalTitle').textContent=`[${pn}] ${p.name||''}`;
  document.getElementById('modalSub').textContent=`${p.category||''} · 자재 ${bomCount}건${_isAdmin?' · 기준인건비 '+fmt(ld.total_manual)+'원/EA':''}`;

  // 상단 요약 카드
  let html=`<div class="m-stats">
    <div class="m-stat accent"><div class="ms-label">원재료비(A)</div><div class="ms-value" style="color:#c0392b">${fmt0(md.raw_total)}<span class="ms-unit">원</span></div></div>
    <div class="m-stat"><div class="ms-label">부재료비</div><div class="ms-value" style="color:#8e44ad">${fmt0(md.sub_total)}<span class="ms-unit">원</span></div></div>
    ${_isAdmin?`<div class="m-stat"><div class="ms-label">기준인건비/EA</div><div class="ms-value" style="color:#2471a3">${fmt(ld.total_manual)}<span class="ms-unit">원</span></div></div>`:''}
    <div class="m-stat total"><div class="ms-label">합계</div><div class="ms-value">${fmt0(md.total + ld.total_manual)}<span class="ms-unit">원</span></div></div>
  </div>`;

  // 테이블
  html+=`<table><thead><tr>
    <th style="width:70px">자품번</th><th>자재명</th><th style="width:55px">구분</th><th style="width:45px">단위</th>
    <th style="width:70px">전개수량</th><th style="width:65px">단가</th><th style="width:75px">원가</th><th style="width:80px">단가일</th>
  </tr></thead><tbody>`;

  // 원재료 그룹
  if(rawItems.length){
    html+=`<tr class="group-row"><td colspan="6" style="color:#c0392b">원재료(A) (${rawItems.length}건)</td><td class="r" colspan="2" style="color:#c0392b;font-size:13px">${fmt0(md.raw_total)}</td></tr>`;
    for(const m of rawItems){
      html+=`<tr>
        <td style="font-weight:600">${m.pn}</td><td>${m.name}</td>
        <td class="c"><span class="badge badge-bad" style="font-size:10px">원재료</span></td>
        <td class="c">${m.unit}</td>
        <td class="r">${fmt4(m.qty)}</td><td class="r">${fmt0(m.price)}</td>
        <td class="r" style="font-weight:700">${fmt0(m.cost)}</td>
        <td class="c" style="font-size:11px;color:var(--muted)">${m.date||''}</td>
      </tr>`;
    }
  }

  // 부재료 그룹
  if(subItems.length){
    html+=`<tr class="group-row"><td colspan="6" style="color:#8e44ad">부재료(B/C/D) (${subItems.length}건)</td><td class="r" colspan="2" style="color:#8e44ad;font-size:13px">${fmt0(md.sub_total)}</td></tr>`;
    for(const m of subItems){
      html+=`<tr>
        <td style="font-weight:600">${m.pn}</td><td>${m.name}</td>
        <td class="c"><span class="badge badge-warn" style="font-size:10px">부재료</span></td>
        <td class="c">${m.unit}</td>
        <td class="r">${fmt4(m.qty)}</td><td class="r">${fmt0(m.price)}</td>
        <td class="r" style="font-weight:700">${fmt0(m.cost)}</td>
        <td class="c" style="font-size:11px;color:var(--muted)">${m.date||''}</td>
      </tr>`;
    }
  }

  // 합계
  html+=`<tr class="total-row"><td colspan="6" style="text-align:right;font-size:13px">합계</td><td class="r" style="font-size:14px" colspan="2">${fmt0(md.total)}</td></tr>`;
  html+=`</tbody></table>`;

  document.getElementById('modalBody').innerHTML=html;
}

// =========================================
// 원부재료 모달 (BOM과 동일 - openBom으로 통합)
// =========================================
async function openMaterial(pn){ openBom(pn); }

// =========================================
// 인건비 모달
// =========================================
async function openLabor(pn){
  openModal(pn,'로딩 중...','<div style="text-align:center;padding:60px;color:var(--muted)">데이터 조회 중...</div>');
  const res=await fetch('/api/labor/'+pn);
  const d=await res.json();
  const p=d.product;

  document.getElementById('modalTitle').textContent=`[${pn}] ${p.name||''} — 인건비 계산식`;
  document.getElementById('modalSub').textContent=`${p.category||''} · ${p.type||''} · 절단투입 ${fmt4(d.cut_kg)}KG · 내포장 ${d.inner_ea}EA`;

  // 상단 요약
  let html=`<div class="m-stats">
    <div class="m-stat"><div class="ms-label">절단 투입량</div><div class="ms-value">${fmt4(d.cut_kg)}<span class="ms-unit"> KG</span></div></div>
    <div class="m-stat"><div class="ms-label">내포장 EA</div><div class="ms-value">${d.inner_ea}<span class="ms-unit"> EA</span></div></div>
    <div class="m-stat accent"><div class="ms-label">인건비(수작업)</div><div class="ms-value" style="color:var(--accent)">${fmt(d.total_manual)}<span class="ms-unit">원</span></div></div>
    <div class="m-stat"><div class="ms-label">인건비(로터리)</div><div class="ms-value">${fmt(d.total_rotary)}<span class="ms-unit">원</span></div></div>
  </div>`;

  const colors=['#2b6cb0','#27ae60','#d4a017','#e74c3c','#8e44ad','#e67e22','#16a085','#c0392b'];
  let totalCost=d.total_manual;

  // 공정별 카드
  for(let i=0;i<d.details.length;i++){
    const proc=d.details[i];
    const color=colors[i%colors.length];
    const pct=totalCost>0?(proc.cost/totalCost*100).toFixed(1):'0';

    if(proc.is_common){
      // 공통배부: 비율 배부 방식으로 표시
      html+=`<div class="proc-card" style="border-left:4px solid ${color};background:#f8f0ff">
        <div class="pc-head">
          <div class="pc-title" style="color:${color}">${proc.label} <span style="font-size:11px;color:var(--muted);font-weight:400">배부율 ${proc.common_rate}% · ${proc.workers.length}명</span></div>
          <div class="pc-result" style="color:${color}">${fmt(proc.cost)}원 <span style="font-size:11px;color:var(--muted)">(${pct}%)</span></div>
        </div>
        <div class="pc-workers">`;
      for(const w of proc.workers){
        html+=`<span class="worker-tag">
          <span class="wt-name">${w.name}</span>
          <span class="wt-rate">${fmt0(w.hourly)}원/h</span>
        </span>`;
      }
      html+=`</div>
        <div class="pc-formula">
          <div class="pf-line">공통배부율 = 공통인원 월급합(퇴직포함) ÷ 생산직 월급합(퇴직포함) = <b>${proc.common_rate}%</b></div>
          <div class="pf-result">직접인건비 ${fmt(proc.direct_labor)}원 × ${proc.common_rate}% = <b>${fmt(proc.cost)}원</b></div>
        </div>
      </div>`;
    } else {
      // 일반 공정
      html+=`<div class="proc-card" style="border-left:4px solid ${color}">
        <div class="pc-head">
          <div class="pc-title" style="color:${color}">${proc.label} <span style="font-size:11px;color:var(--muted);font-weight:400">${proc.capa.toLocaleString()}${proc.unit}/일 · ${proc.workers.length}명</span></div>
          <div class="pc-result" style="color:${color}">${fmt(proc.cost)}원 <span style="font-size:11px;color:var(--muted)">(${pct}%)</span></div>
        </div>
        <div class="pc-workers">`;
      for(const w of proc.workers){
        html+=`<span class="worker-tag ${w.is_min?'min':''}">
          <span class="wt-name">${w.name}</span>
          <span class="wt-rate">${fmt0(w.hourly)}원/h</span>
          <span style="font-size:9.5px;color:var(--muted)">${w.is_min?'최저시급':fmt0(w.pay)+'÷'+209}</span>
        </span>`;
      }
      html+=`</div>
        <div class="pc-formula">
          <div class="pf-line">시급합 = ${proc.workers.map(w=>fmt0(w.hourly)).join(' + ')} = <b>${fmt(proc.hourly_sum)}원/h</b></div>
          <div class="pf-line">단가 = (${fmt(proc.hourly_sum)} × ${proc.hours}h) ÷ ${proc.capa.toLocaleString()}${proc.unit} = <b>${fmt4(proc.per_unit)}원/${proc.unit}</b></div>
          <div class="pf-result">${fmt4(proc.per_unit)} × ${fmt4(proc.usage)}${proc.unit} = ${fmt(proc.cost)}원</div>
        </div>
      </div>`;
    }
  }

  // 합계
  html+=`<div class="labor-total">
    <div class="lt-items">${d.details.map(p=>`${p.label} ${fmt(p.cost)}원`).join(' + ')}</div>
    <div style="text-align:right">
      <div class="lt-sum">= ${fmt(d.total_manual)}원</div>
      <div class="lt-sub">로터리 기준: ${fmt(d.total_rotary)}원 (차이 ${fmt(d.total_manual-d.total_rotary)}원)</div>
    </div>
  </div>`;

  // 비율 바
  html+=`<div class="ratio-bar">`;
  for(let i=0;i<d.details.length;i++){
    const proc=d.details[i];
    const pct=totalCost>0?(proc.cost/totalCost*100):0;
    if(pct<1.5)continue;
    html+=`<div style="width:${pct}%;background:${colors[i%colors.length]}" title="${proc.label}: ${fmt(proc.cost)}원">${pct>=8?proc.label:''} ${pct.toFixed(0)}%</div>`;
  }
  html+=`</div>`;

  document.getElementById('modalBody').innerHTML=html;
}

// =========================================
// 일자별 생산실적 탭
// =========================================

async function uploadProd(){
  const fileInput=document.getElementById('prodFileInput');
  const file=fileInput.files[0];
  if(!file){document.getElementById('prodUploadStatus').innerHTML='<span style="color:var(--bad)">파일을 선택하세요</span>';return;}
  const mode=document.getElementById('prodMode').value;
  document.getElementById('prodUploadStatus').innerHTML='<span style="color:var(--accent)">업로드 중...</span>';
  const fd=new FormData();
  fd.append('file',file);
  fd.append('mode',mode);
  const res=await fetch('/api/upload_prod',{method:'POST',body:fd});
  const d=await res.json();
  if(d.ok){
    document.getElementById('prodUploadStatus').innerHTML=`<span style="color:var(--good)">${d.msg} (총 ${d.total}건)</span>`;
    fileInput.value='';
    loadProdRecords();
  } else {
    document.getElementById('prodUploadStatus').innerHTML=`<span style="color:var(--bad)">${d.msg}</span>`;
  }
}

async function loadProdRecords(){
  const from=document.getElementById('prodFrom')?.value||'';
  const to=document.getElementById('prodTo')?.value||'';
  const q=document.getElementById('prodSearch')?.value||'';
  const params=new URLSearchParams();
  if(from)params.set('from',from);
  if(to)params.set('to',to);
  if(q)params.set('q',q);
  const res=await fetch('/api/prod_records?'+params);
  const d=await res.json();

  // 요약 카드
  const recs=d.records;
  const dates=Object.keys(d.daily_summary).sort();
  // 생산성 집계 (입력된 것만)
  const withActual=recs.filter(r=>r.productivity!=null);
  const avgProd=withActual.length?Math.round(withActual.reduce((s,r)=>s+r.productivity,0)/withActual.length*10)/10:null;
  document.getElementById('prodStats').innerHTML=`
    <div class="stat"><div class="label">총 실적건수</div><div class="value">${recs.length}<span style="font-size:14px;color:var(--muted)"> / ${d.total_count}건</span></div><div class="detail">${dates.length}일</div></div>
    <div class="stat warn"><div class="label">실적 입력</div><div class="value">${withActual.length}<span style="font-size:14px;color:var(--muted)"> / ${recs.length}건</span></div></div>
    <div class="stat ${avgProd!=null?(avgProd>=100?'good':'bad'):''}"><div class="label">평균 생산성</div>
      <div class="value" style="color:${avgProd!=null?(avgProd>=100?'var(--good)':'var(--bad)'):'var(--muted)'}">${avgProd!=null?avgProd+'%':'—'}</div>
      <div class="detail">${avgProd!=null?(avgProd>=100?'기준 이상':'기준 미달'):'실적 입력 후 표시'}</div></div>
  `;

  // 테이블 제목
  const periodStr=dates.length?` (${dates[0]} ~ ${dates[dates.length-1]})`:'';
  document.getElementById('prodTableTitle').textContent=`생산실적 ${recs.length}건${periodStr}`;

  // 테이블 렌더
  let html='';
  let prevDate='';
  let dayCount=0;

  function flushDay(){
    if(!prevDate)return '';
    return `<tr class="total-row">
      <td colspan="12" class="r" style="font-size:12px">${prevDate} 소계 (${dayCount}건)</td></tr>`;
  }

  for(const r of recs){
    if(r.date!==prevDate){
      html+=flushDay();
      prevDate=r.date; dayCount=0;
      html+=`<tr class="group-row"><td colspan="12" style="font-size:13px">${r.date}</td></tr>`;
    }
    dayCount++;
    const badge=r.is_prod?'badge-good':'badge-warn';
    const typeStr=r.is_prod?'제품':'반제품';
    const alJson=JSON.stringify({date:r.date,pn:r.pn,qty:r.qty}).replace(/"/g,'&quot;');

    // 실제인건비/EA 셀 (클릭→입력 모달)
    let actualEaCell;
    if(r.actual_labor_ea!=null){
      actualEaCell=`<td class="r clickable" data-al="${alJson}" onclick="openActualLabor(JSON.parse(this.dataset.al))" style="font-weight:600" title="총액:${fmt0(r.actual_labor)}원 / ${fmt0(r.actual_mh)}MH\n클릭: 수정">${fmt(r.actual_labor_ea)}</td>`;
    } else {
      actualEaCell=`<td class="c clickable" data-al="${alJson}" onclick="openActualLabor(JSON.parse(this.dataset.al))" title="클릭: 투입인원 입력"><span style="color:var(--accent);font-size:11px;border:1px dashed var(--border);padding:2px 8px;border-radius:4px">입력</span></td>`;
    }

    // 기준 EA/MH (툴팁에 산출 근거)
    let stdTip='';
    if(r.std_info&&r.std_info.proc){
      const si=r.std_info;
      stdTip=`[${si.proc}] 기준\nCAPA: ${si.capa.toLocaleString()}${si.unit} / (${si.hc}명 × ${si.hours}h)\n= ${si.capa.toLocaleString()} ÷ ${si.hc*si.hours} = ${fmt(r.std_ea_mh)} ${si.unit}/MH\n\n1 Man-Hour당 ${fmt(r.std_ea_mh)}${si.unit} 생산 가능`;
    }
    const stdCell=r.std_ea_mh?`<td class="r" style="color:var(--muted);cursor:help" title="${stdTip}">${fmt(r.std_ea_mh)}</td>`:`<td class="c" style="color:var(--muted)">-</td>`;

    // 실제 EA/MH (툴팁에 산출 근거)
    let actMhCell;
    if(r.actual_ea_mh!=null){
      const better=r.actual_ea_mh>=r.std_ea_mh;
      const actTip=`실적수량: ${fmt0(r.qty)}\n투입 Man-Hour: ${r.actual_mh}MH\n\n${fmt0(r.qty)} ÷ ${r.actual_mh}MH = ${fmt(r.actual_ea_mh)} EA/MH\n\n기준(${fmt(r.std_ea_mh)}) 대비 ${r.productivity}%`;
      actMhCell=`<td class="r" style="font-weight:600;color:${better?'var(--good)':'var(--bad)'};cursor:help" title="${actTip}">${fmt(r.actual_ea_mh)}</td>`;
    } else {
      actMhCell=`<td class="c" style="color:var(--muted)">-</td>`;
    }

    // 생산성
    let prodCell;
    if(r.productivity!=null){
      const color=r.productivity>=100?'var(--good)':r.productivity>=80?'var(--warn)':'var(--bad)';
      const bg=r.productivity>=100?'#c6f6d5':r.productivity>=80?'#fefcbf':'#fed7d7';
      prodCell=`<td class="c"><span class="badge" style="background:${bg};color:${color};font-size:11px;font-weight:700">${r.productivity}%</span></td>`;
    } else {
      prodCell=`<td class="c" style="color:var(--muted)">-</td>`;
    }

    html+=`<tr>
      <td class="c" style="font-size:11px;color:var(--muted)">${r.date}</td>
      <td class="c clickable" onclick="openProdDetail('${r.pn}')" style="font-weight:600">${r.pn}</td>
      <td class="clickable" onclick="openProdDetail('${r.pn}')">${r.name.substring(0,32)}</td>
      <td class="c"><span class="badge ${badge}" style="font-size:10px">${typeStr}</span></td>
      <td class="r">${fmt0(r.qty)}</td>
      <td class="r">${fmt0(r.erp_price)}</td>
      ${_isAdmin?`<td class="r clickable" onclick="openLabor('${r.pn}')" style="color:var(--accent)">${r.labor_ea?fmt0(r.labor_ea):'-'}</td>
      <td class="r" style="color:var(--muted)">${r.labor_ea_r?fmt0(r.labor_ea_r):'-'}</td>
      ${actualEaCell}`:''}

      ${stdCell}
      ${actMhCell}
      ${prodCell}
      <td style="font-size:11px;color:var(--muted)">${(r.bigo||'').substring(0,16)}</td>
    </tr>`;
  }
  html+=flushDay();

  document.getElementById('prodBody').innerHTML=html;
}

// 생산실적 품번 클릭 → 상세 모달
async function openProdDetail(pn){
  openModal(pn,'로딩 중...','<div style="text-align:center;padding:60px;color:var(--muted)">데이터 조회 중...</div>');
  const [detRes, matRes, labRes] = await Promise.all([
    fetch('/api/prod_detail/'+pn).then(r=>r.json()),
    fetch('/api/material/'+pn).then(r=>r.json()),
    fetch('/api/labor/'+pn).then(r=>r.json()),
  ]);
  const d=detRes, md=matRes, ld=labRes, p=d.product;
  const c=d.cost;

  document.getElementById('modalTitle').textContent=`[${pn}] ${p.name||d.history[0]?.name||''}`;
  document.getElementById('modalSub').textContent=`${p.category||''} · 생산이력 ${d.history_count}건 · 총 생산수량 ${fmt0(d.total_qty)}`;

  let html=`<div class="m-stats">
    <div class="m-stat accent"><div class="ms-label">원재료비</div><div class="ms-value" style="color:#c0392b">${fmt0(c.raw)}<span class="ms-unit">원</span></div></div>
    <div class="m-stat"><div class="ms-label">부재료비</div><div class="ms-value" style="color:#8e44ad">${fmt0(c.sub)}<span class="ms-unit">원</span></div></div>
    ${_isAdmin?`<div class="m-stat"><div class="ms-label">인건비</div><div class="ms-value" style="color:#2471a3">${fmt(c.labor)}<span class="ms-unit">원</span></div></div>`:''}
    <div class="m-stat total"><div class="ms-label">기준원가</div><div class="ms-value">${fmt0(c.total)}<span class="ms-unit">원</span></div></div>
    <div class="m-stat"><div class="ms-label">총 생산수량</div><div class="ms-value">${fmt0(d.total_qty)}<span class="ms-unit">EA</span></div></div>
  </div>`;

  // 원부재료 내역 (간략)
  if(md.mat_items.length){
    html+=`<div style="font-size:13px;font-weight:700;color:var(--primary);margin:16px 0 8px;padding:6px 0;border-bottom:2px solid var(--accent)">원부재료 내역</div>
    <table><thead><tr><th>자품번</th><th>자재명</th><th>구분</th><th>단위</th><th>수량</th><th>단가</th><th>원가</th></tr></thead><tbody>`;
    for(const m of md.mat_items){
      const color=m.type==='원재료'?'#c0392b':'#8e44ad';
      html+=`<tr><td style="font-weight:600">${m.pn}</td><td>${m.name}</td>
        <td class="c"><span class="badge" style="background:${m.type==='원재료'?'#fadbd8':'#e8daef'};color:${color};font-size:10px">${m.type}</span></td>
        <td class="c">${m.unit}</td><td class="r">${fmt4(m.qty)}</td><td class="r">${fmt0(m.price)}</td>
        <td class="r" style="font-weight:700">${fmt0(m.cost)}</td></tr>`;
    }
    html+=`</tbody></table>`;
  }

  // 인건비 요약
  if(ld.details.length){
    html+=`<div style="font-size:13px;font-weight:700;color:var(--primary);margin:16px 0 8px;padding:6px 0;border-bottom:2px solid var(--accent)">인건비 내역 (수작업) <span style="font-size:11px;color:var(--muted);cursor:pointer;font-weight:400" onclick="closeModal();setTimeout(()=>openLabor('${pn}'),200)">[상세 계산식 보기]</span></div>
    <table><thead><tr><th>공정</th><th>인건비</th><th>비율</th></tr></thead><tbody>`;
    for(const proc of ld.details){
      const pct=ld.total_manual>0?(proc.cost/ld.total_manual*100).toFixed(1):'0';
      html+=`<tr><td>${proc.label}</td><td class="r" style="font-weight:600">${fmt(proc.cost)}원</td><td class="r">${pct}%</td></tr>`;
    }
    html+=`<tr style="background:var(--primary);color:#fff;font-weight:700"><td>합계</td><td class="r">${fmt(ld.total_manual)}원</td><td class="r">100%</td></tr>`;
    html+=`</tbody></table>`;
  }

  // 생산이력
  if(d.history.length){
    html+=`<div style="font-size:13px;font-weight:700;color:var(--primary);margin:16px 0 8px;padding:6px 0;border-bottom:2px solid var(--accent)">생산 이력 (최근 ${Math.min(d.history.length,50)}건)</div>
    <table><thead><tr><th>일자</th><th>실적수량</th><th>ERP단가</th><th>인건비소계</th><th>원가소계</th><th>비고</th></tr></thead><tbody>`;
    for(const h of d.history){
      const lsub=c.labor*h.qty, tsub=c.total*h.qty;
      html+=`<tr><td class="c">${h.date}</td><td class="r">${fmt0(h.qty)}</td><td class="r">${fmt0(h.erp_price)}</td>
        <td class="r">${fmt0(lsub)}</td><td class="r" style="font-weight:600">${fmt0(tsub)}</td>
        <td style="font-size:11px;color:var(--muted)">${(h.bigo||'').substring(0,20)}</td></tr>`;
    }
    html+=`</tbody></table>`;
  }

  document.getElementById('modalBody').innerHTML=html;
}

// =========================================
// 실제인건비 입력 모달
// =========================================

let _alCtx=null;
let _alRows=[];      // [{start,end,names:['김흥수','한승엽',...]}]
let _alEmpList=[];   // [{name,hourly}]
let _alEmpMap={};    // name→hourly

async function openActualLabor(ctx){
  _alCtx=ctx;
  openModal('실제인건비 입력','로딩 중...','<div style="text-align:center;padding:60px;color:var(--muted)">데이터 조회 중...</div>');
  const res=await fetch('/api/actual_labor/get',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(ctx)});
  const d=await res.json();
  _alEmpList=d.emp_list||[];
  _alEmpMap={};
  _alEmpList.forEach(e=>{_alEmpMap[e.name]=e.hourly});
  // 복원
  const saved=d.data||{};
  if(saved.timeSlots&&saved.timeSlots.length){
    // grid→rows 변환
    _alRows=saved.timeSlots.map((ts,i)=>{
      const names=[];
      for(const [name,checks] of Object.entries(saved.grid||{})){
        if(checks[i]) names.push(name);
      }
      return {start:ts.start,end:ts.end,names};
    });
  } else {
    _alRows=[{start:'09:00',end:'12:00',names:[]},{start:'13:00',end:'17:00',names:[]}];
  }
  document.getElementById('modalTitle').textContent=`실제인건비 — ${ctx.date} / ${ctx.pn}`;
  document.getElementById('modalSub').textContent=`시간대별 투입인원을 입력하세요 (점심 12:00~13:00 자동 차감)`;
  renderAlRows();
}

function calcH(start,end){
  if(!start||!end) return 0;
  const [sh,sm]=start.split(':').map(Number);
  const [eh,em]=end.split(':').map(Number);
  let sMin=sh*60+sm, eMin=eh*60+em;
  if(eMin<sMin) eMin+=24*60;
  let total=eMin-sMin;
  // 점심시간(12:00~13:00) 자동 차감
  const lunchS=720, lunchE=780;
  const overS=Math.max(sMin,lunchS), overE=Math.min(eMin,lunchE);
  if(overS<overE) total-=(overE-overS);
  return Math.round(Math.max(total,0)/60*100)/100;
}
function getH(name){return _alEmpMap[name]||10300;}

function fmtTime(t){
  if(!t) return '';
  const [h,m]=t.split(':').map(Number);
  const ampm=h<12?'오전':'오후';
  const dh=h===0?12:h>12?h-12:h;
  return ampm+' '+String(dh).padStart(2,'0')+':'+String(m).padStart(2,'0');
}

function timeSelectHtml(val,rowIdx,field){
  const [vh,vm]=(val||'09:00').split(':').map(Number);
  const sid=field+rowIdx;
  let hOpts='',mOpts='';
  for(let h=0;h<24;h++){
    const sel=h===vh?'selected':'';
    const label=(h<12?'오전 ':'오후 ')+(h===0?'12':h>12?String(h-12).padStart(2,'0'):String(h).padStart(2,'0'));
    hOpts+=`<option value="${h}" ${sel}>${label}시</option>`;
  }
  for(let m=0;m<60;m+=5){
    const sel=m===vm?'selected':'';
    mOpts+=`<option value="${m}" ${sel}>${String(m).padStart(2,'0')}분</option>`;
  }
  return `<select onchange="onTimeSelect(${rowIdx},'${field}')" id="${sid}h" style="padding:4px 2px;border:1px solid var(--border);border-radius:5px;font-size:12px">${hOpts}</select>
  <select onchange="onTimeSelect(${rowIdx},'${field}')" id="${sid}m" style="padding:4px 2px;border:1px solid var(--border);border-radius:5px;font-size:12px">${mOpts}</select>`;
}
function onTimeSelect(rowIdx,field){
  const h=document.getElementById(field+rowIdx+'h').value;
  const m=document.getElementById(field+rowIdx+'m').value;
  _alRows[rowIdx][field]=String(h).padStart(2,'0')+':'+String(m).padStart(2,'0');
  renderAlRows();
}

function renderAlRows(){
  let html='';
  let grandTotal=0, totalMH=0;

  for(let i=0;i<_alRows.length;i++){
    const row=_alRows[i];
    const h=calcH(row.start,row.end);
    const rowCost=row.names.reduce((s,n)=>s+getH(n)*h,0);
    grandTotal+=rowCost;
    totalMH+=row.names.length*h;

    html+=`<div style="border:1px solid var(--border);border-radius:10px;margin-bottom:12px;overflow:hidden">`;
    // 시간대 헤더
    html+=`<div style="display:flex;align-items:center;gap:8px;padding:10px 14px;background:#f4f6f8;border-bottom:1px solid var(--border);flex-wrap:wrap">
      ${timeSelectHtml(row.start,i,'start')}
      <span style="color:var(--muted);font-size:13px">~</span>
      ${timeSelectHtml(row.end,i,'end')}
      <span style="background:var(--primary);color:#fff;padding:2px 10px;border-radius:4px;font-size:13px;font-weight:700">${fmtTime(row.start)} ~ ${fmtTime(row.end)} (${h}h)</span>
      <span style="color:var(--accent);font-size:12px;font-weight:600">${row.names.length}명</span>
      <span style="color:var(--muted);font-size:12px;margin-left:auto">${fmt0(rowCost)}원</span>
      <button onclick="_alRows.splice(${i},1);renderAlRows()" style="background:none;border:none;color:var(--bad);cursor:pointer;font-size:16px;margin-left:4px" title="시간대 삭제">×</button>
    </div>`;
    // 인원 영역
    html+=`<div style="padding:10px 14px;display:flex;flex-wrap:wrap;gap:4px;align-items:center;min-height:40px">`;
    // 기존 사원 태그 (체크된)
    for(let j=0;j<row.names.length;j++){
      const n=row.names[j];
      const hr=getH(n);
      const isReg=_alEmpMap[n]!=null;
      html+=`<span style="display:inline-flex;align-items:center;gap:4px;background:${isReg?'#e8f4fd':'#fff3cd'};border:1px solid ${isReg?'#b3d9f2':'#f0d78c'};border-radius:6px;padding:3px 8px;font-size:12px">
        <span style="font-weight:600">${n}</span>
        <span style="color:var(--accent);font-size:10px">${fmt0(hr)}원/h</span>
        <button onclick="_alRows[${i}].names.splice(${j},1);renderAlRows()" style="background:none;border:none;color:var(--bad);cursor:pointer;font-size:14px;padding:0 2px;line-height:1">×</button>
      </span>`;
    }
    // 추가 버튼 (드롭다운)
    html+=`<div style="position:relative;display:inline-block">
      <button onclick="toggleEmpPicker(${i})" style="padding:3px 10px;background:#fff;border:1px dashed var(--accent);border-radius:6px;color:var(--accent);font-size:12px;cursor:pointer;font-weight:600">+ 인원추가</button>
      <div id="empPicker${i}" style="display:none;position:absolute;top:100%;left:0;z-index:200;background:#fff;border:1px solid var(--border);border-radius:8px;box-shadow:0 4px 16px rgba(0,0,0,.15);padding:8px;width:260px;max-height:250px;overflow-y:auto;margin-top:4px">
        <input type="text" id="empSearch${i}" placeholder="이름 검색/직접입력..." oninput="filterEmpPicker(${i})" onkeydown="if(event.key==='Enter'){addManualName(${i});}"
          style="width:100%;padding:6px 8px;border:1px solid var(--border);border-radius:6px;font-size:12px;margin-bottom:6px">
        <div id="empPickerList${i}">`;
    // 등록 사원 목록
    for(const emp of _alEmpList){
      const already=row.names.includes(emp.name);
      html+=`<div onclick="${already?'':`addEmpToRow(${i},'${emp.name}')`}" style="padding:5px 8px;cursor:${already?'default':'pointer'};border-radius:4px;display:flex;justify-content:space-between;align-items:center;${already?'opacity:.4':''}${already?'':'hover:background:#f0f7ff'}" class="ep-item">
        <span style="font-size:12px">${emp.name}</span>
        <span style="font-size:11px;color:var(--accent)">${fmt0(emp.hourly)}원/h</span>
        ${already?'<span style="font-size:10px;color:var(--good)">추가됨</span>':''}
      </div>`;
    }
    html+=`</div>
        <button onclick="addManualName(${i})" style="margin-top:4px;width:100%;padding:5px;background:var(--accent);color:#fff;border:none;border-radius:5px;font-size:11px;cursor:pointer">입력한 이름 추가 (미등록자)</button>
      </div>
    </div>`;
    html+=`</div></div>`;
  }

  // 시간대 추가
  const lastEnd=_alRows.length?_alRows[_alRows.length-1].end:'09:00';
  const [lh,lm]=lastEnd.split(':').map(Number);
  const nEnd=String(Math.min(lh+2,23)).padStart(2,'0')+':'+String(lm).padStart(2,'0');
  html+=`<button onclick="_alRows.push({start:'${lastEnd}',end:'${nEnd}',names:[]});renderAlRows()"
    style="padding:8px 20px;background:#f0f7ff;color:var(--accent);border:1px solid var(--accent);border-radius:8px;font-size:13px;cursor:pointer;font-weight:600;width:100%">+ 시간대 추가</button>`;

  // 합계 + 공통배부 + 저장
  const pplSet=new Set(); _alRows.forEach(r=>r.names.forEach(n=>pplSet.add(n)));
  const commonRate=39.7;
  const commonAmt=grandTotal*commonRate/100;
  const withCommon=grandTotal+commonAmt;
  const eaQty=_alCtx?_alCtx.qty:0;
  const perEa=eaQty>0?withCommon/eaQty:0;
  html+=`<div style="margin-top:16px;background:var(--primary);color:#fff;border-radius:10px;padding:16px 20px;display:flex;justify-content:space-between;align-items:center">
    <div>
      <div style="font-size:12px;opacity:.7">실제인건비 (공통배부 ${commonRate}% 포함)</div>
      <div style="display:flex;align-items:baseline;gap:16px">
        <div><div style="font-size:24px;font-weight:700">${fmt(perEa)}원/EA</div></div>
        <div style="opacity:.7"><div style="font-size:13px">직접 ${fmt0(grandTotal)} + 공통 ${fmt0(commonAmt)} = ${fmt0(withCommon)}원 ÷ ${fmt0(eaQty)}EA</div></div>
      </div>
      <div style="font-size:11px;opacity:.7;margin-top:2px">총 ${totalMH.toFixed(1)} Man-Hour | 투입인원 ${pplSet.size}명</div>
    </div>
    <button onclick="saveActualLabor()" style="padding:10px 28px;background:#fff;color:var(--primary);border:none;border-radius:8px;font-size:14px;font-weight:700;cursor:pointer">저장</button>
  </div>`;
  html+=`<div id="alSaveStatus" style="margin-top:8px"></div>`;
  document.getElementById('modalBody').innerHTML=html;
}

function toggleEmpPicker(rowIdx){
  const el=document.getElementById('empPicker'+rowIdx);
  el.style.display=el.style.display==='none'?'block':'none';
  if(el.style.display==='block'){
    const inp=document.getElementById('empSearch'+rowIdx);
    setTimeout(()=>inp.focus(),50);
  }
}
function filterEmpPicker(rowIdx){
  const q=document.getElementById('empSearch'+rowIdx).value.toLowerCase();
  document.querySelectorAll('#empPickerList'+rowIdx+' .ep-item').forEach(el=>{
    el.style.display=el.textContent.toLowerCase().includes(q)?'':'none';
  });
}
function addEmpToRow(rowIdx,name){
  if(!_alRows[rowIdx].names.includes(name)){
    _alRows[rowIdx].names.push(name);
  }
  renderAlRows();
}
function addManualName(rowIdx){
  const inp=document.getElementById('empSearch'+rowIdx);
  const name=(inp?inp.value:'').trim();
  if(!name) return;
  if(!_alRows[rowIdx].names.includes(name)){
    _alRows[rowIdx].names.push(name);
  }
  renderAlRows();
}

async function saveActualLabor(){
  // rows→timeSlots+grid 변환
  const timeSlots=_alRows.map(r=>({start:r.start,end:r.end}));
  const grid={};
  const allNames=new Set(); _alRows.forEach(r=>r.names.forEach(n=>allNames.add(n)));
  for(const name of allNames){
    grid[name]=_alRows.map(r=>r.names.includes(name));
  }
  const body={..._alCtx, timeSlots, grid};
  const res=await fetch('/api/actual_labor/save',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const d=await res.json();
  if(d.ok){
    document.getElementById('alSaveStatus').innerHTML=`<span style="color:var(--good);font-size:13px">${d.msg}</span>`;
    setTimeout(()=>{closeModal();loadProdRecords();},800);
  } else {
    document.getElementById('alSaveStatus').innerHTML=`<span style="color:var(--bad);font-size:13px">${d.msg||'저장 실패'}</span>`;
  }
}

// =========================================
// 설정/관리 탭
// =========================================

// 파일 업로드
function handleDrop(e){
  e.preventDefault();
  e.currentTarget.style.borderColor='var(--border)';
  e.currentTarget.style.background='#fafbfc';
  const file=e.dataTransfer.files[0];
  if(file) handleFile(file);
}
async function handleFile(file){
  if(!file) return;
  if(!file.name.endsWith('.xlsx')){
    document.getElementById('uploadStatus').innerHTML='<div style="color:var(--bad);font-size:13px">xlsx 파일만 업로드 가능합니다</div>';
    return;
  }
  document.getElementById('uploadStatus').innerHTML='<div style="color:var(--accent);font-size:13px">업로드 중...</div>';
  const fd=new FormData();
  fd.append('file',file);
  const res=await fetch('/api/upload',{method:'POST',body:fd});
  const d=await res.json();
  if(d.ok){
    document.getElementById('uploadStatus').innerHTML=`<div style="padding:12px;background:#f0fff4;border-radius:8px;border:1px solid #c6f6d5;font-size:13px">
      <b style="color:var(--good)">업로드 성공</b><br><span style="color:var(--muted)">${d.msg}</span>
      <br><button onclick="location.reload()" style="margin-top:8px;padding:6px 16px;background:var(--accent);color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:12px">새로고침하여 반영</button>
    </div>`;
  } else {
    document.getElementById('uploadStatus').innerHTML=`<div style="color:var(--bad);font-size:13px">${d.msg}</div>`;
  }
}

// 인건비 관리
async function loadEmployees(){
  const res=await fetch('/api/employees');
  const d=await res.json();
  const tbody=document.getElementById('empBody');
  if(!tbody) return;
  let html='';
  d.employees.forEach((e,i)=>{
    const hwStr=e.pay>0?fmt0(e.hourly):`${d.min_wage.toLocaleString()} (최저)`;
    html+=`<tr id="emp-${i}">
      <td class="c">${i+1}</td>
      <td class="c" style="font-weight:600">${e.name}</td>
      <td class="c"><input type="number" value="${e.pay}" id="pay-${i}" data-name="${e.name}"
        style="width:100%;padding:6px 10px;border:1px solid var(--border);border-radius:6px;text-align:right;font-size:13px;box-sizing:border-box"
        onchange="this.style.borderColor='var(--warn)';this.style.background='#fffbeb'"></td>
      <td class="r" style="color:var(--accent)">${hwStr}</td>
      <td class="c"><input type="checkbox" ${e.common?'checked':''} id="com-${i}" style="width:16px;height:16px;cursor:pointer"></td>
      <td class="c" style="white-space:nowrap">
        <button onclick="saveEmployee(${i})" style="padding:5px 14px;background:var(--accent);color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer;margin-right:4px">저장</button>
        <button onclick="deleteEmployee('${e.name}')" style="padding:5px 14px;background:var(--bad);color:#fff;border:none;border-radius:6px;font-size:12px;cursor:pointer">삭제</button>
      </td>
    </tr>`;
  });
  tbody.innerHTML=html;
}

async function addEmployee(){
  const name=document.getElementById('empName').value.trim();
  const pay=parseInt(document.getElementById('empPay').value)||0;
  const common=document.getElementById('empCommon').checked;
  if(!name){document.getElementById('empAddStatus').innerHTML='<span style="color:var(--bad);font-size:12px">이름을 입력하세요</span>';return;}
  const res=await fetch('/api/employees/add',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,pay,common})});
  const d=await res.json();
  document.getElementById('empAddStatus').innerHTML=`<span style="color:${d.ok?'var(--good)':'var(--bad)'};font-size:12px">${d.msg}</span>`;
  if(d.ok){document.getElementById('empName').value='';document.getElementById('empPay').value='';loadEmployees();}
}

async function saveEmployee(idx){
  const input=document.getElementById('pay-'+idx);
  const name=input.dataset.name;
  const pay=parseInt(input.value)||0;
  const common=document.getElementById('com-'+idx).checked;
  const res=await fetch('/api/employees/update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name,pay,common})});
  const d=await res.json();
  if(d.ok){input.style.borderColor='var(--good)';input.style.background='#f0fff4';setTimeout(()=>{input.style.borderColor='var(--border)';input.style.background='#fff'},1500);loadEmployees();}
  else{alert(d.msg);}
}

async function deleteEmployee(name){
  if(!confirm(name+'을(를) 삭제하시겠습니까?')) return;
  const res=await fetch('/api/employees/delete',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({name})});
  const d=await res.json();
  if(d.ok) loadEmployees();
  else alert(d.msg);
}

// =========================================
// 인건비 검증
// =========================================
async function uploadVReport(){
  const file=document.getElementById('vReportFile').files[0];
  if(!file){document.getElementById('vReportStatus').innerHTML='<span style="color:var(--bad)">파일을 선택하세요</span>';return;}
  const fd=new FormData(); fd.append('file',file);
  document.getElementById('vReportStatus').innerHTML='<span style="color:var(--accent)">업로드 중...</span>';
  const res=await fetch('/api/verify/upload_report',{method:'POST',body:fd});
  const d=await res.json();
  document.getElementById('vReportStatus').innerHTML=`<span style="color:${d.ok?'var(--good)':'var(--bad)'}">${d.msg}</span>`;
  if(d.ok&&d.months){
    const sel=document.getElementById('verifyMonth');
    sel.innerHTML='<option value="">-- 월 선택 --</option>';
    d.months.forEach(m=>{sel.innerHTML+=`<option value="${m}">${m}</option>`;});
  }
}
function initVerify(){}

async function runVerify(){
  const month=document.getElementById('verifyMonth').value;
  if(!month){alert('검증 월을 선택하세요');return;}
  document.getElementById('verifyResult').innerHTML='<div style="text-align:center;padding:40px;color:var(--muted)">검증 중...</div>';
  const res=await fetch('/api/verify/run?month='+encodeURIComponent(month));
  const d=await res.json();
  const rpt=d.report;

  if(d.report_labor===0&&d.items.length===0){
    document.getElementById('verifyResult').innerHTML='<div class="card" style="text-align:center;padding:30px;color:var(--muted)">해당 월의 원가보고서 또는 생산실적 데이터가 없습니다.</div>';
    return;
  }

  // 요약 카드
  let html=`<div class="stats">
    <div class="stat"><div class="label">검증 월</div><div class="value">${d.month}</div><div class="detail">생산실적 ${d.prod_total}건</div></div>
    <div class="stat"><div class="label">원가보고서 노무비</div><div class="value">${fmt0(d.report_labor)}원</div>
      <div class="detail">급여 ${fmt0(rpt.salary||0)} + 잡급 ${fmt0(rpt.misc||0)} + 퇴직 ${fmt0(rpt.retire||0)}</div></div>
    <div class="stat ${d.pass_manual?'good':'bad'}"><div class="label">검증1: 기준인건비</div>
      <div class="value" style="color:${d.pass_manual?'var(--good)':'var(--bad)'}">${d.pass_manual===null?'—':d.pass_manual?'PASS':'FAIL'}</div>
      <div class="detail">${fmt0(d.total_std_manual)}원 (차이 ${d.diff_manual>=0?'+':''}${fmt0(d.diff_manual)})</div></div>
    <div class="stat ${d.pass_actual===null?'':(d.pass_actual?'good':'bad')}"><div class="label">검증2: 실제인건비</div>
      <div class="value" style="color:${d.pass_actual===null?'var(--muted)':(d.pass_actual?'var(--good)':'var(--bad)')}">${d.pass_actual===null?'미입력':d.pass_actual?'PASS':'FAIL'}</div>
      <div class="detail">${d.total_actual>0?fmt0(d.total_actual)+'원 ('+d.actual_count+'건 입력)':'일자별 생산실적에서 입력'}</div></div>
  </div>`;

  // 비교 테이블
  html+=`<div class="card">
    <h3>${d.month} 인건비 검증 비교표</h3>
    <table>
      <thead><tr><th>구분</th><th>금액</th><th>차이</th><th>검증</th></tr></thead>
      <tbody>
        <tr><td style="font-weight:700">원가보고서 노무비</td><td class="r" style="font-weight:700">${fmt0(d.report_labor)}원</td><td></td><td class="c">기준</td></tr>
        <tr style="${d.pass_manual?'background:#f0fff4':'background:#fff5f5'}">
          <td>검증1: 기준인건비 × 생산수량 (수작업)</td><td class="r" style="font-weight:700">${fmt0(d.total_std_manual)}원</td>
          <td class="r" style="font-weight:700;color:${d.diff_manual>=0?'var(--good)':'var(--bad)'}">${d.diff_manual>=0?'+':''}${fmt0(d.diff_manual)}원</td>
          <td class="c"><span class="badge ${d.pass_manual?'badge-good':'badge-bad'}" style="font-size:13px;padding:4px 12px">${d.pass_manual===null?'—':d.pass_manual?'PASS':'FAIL'}</span></td></tr>
        <tr style="${d.pass_rotary?'background:#f0fff4':'background:#fff5f5'}">
          <td>검증1: 기준인건비 × 생산수량 (로터리)</td><td class="r" style="font-weight:700">${fmt0(d.total_std_rotary)}원</td>
          <td class="r" style="font-weight:700;color:${d.diff_rotary>=0?'var(--good)':'var(--bad)'}">${d.diff_rotary>=0?'+':''}${fmt0(d.diff_rotary)}원</td>
          <td class="c"><span class="badge ${d.pass_rotary?'badge-good':'badge-bad'}" style="font-size:13px;padding:4px 12px">${d.pass_rotary===null?'—':d.pass_rotary?'PASS':'FAIL'}</span></td></tr>`;
  if(d.total_actual>0){
    html+=`<tr style="${d.pass_actual?'background:#f0fff4':'background:#fff5f5'}">
          <td>검증2: 실제인건비 합계 (${d.actual_count}건)</td><td class="r" style="font-weight:700">${fmt0(d.total_actual)}원</td>
          <td class="r" style="font-weight:700;color:${d.diff_actual>=0?'var(--good)':'var(--bad)'}">${d.diff_actual>=0?'+':''}${fmt0(d.diff_actual)}원</td>
          <td class="c"><span class="badge ${d.pass_actual?'badge-good':'badge-bad'}" style="font-size:13px;padding:4px 12px">${d.pass_actual?'PASS':'FAIL'}</span></td></tr>`;
  } else {
    html+=`<tr><td>검증2: 실제인건비</td><td class="c" colspan="3" style="color:var(--muted)">일자별 생산실적 탭에서 투입인원 입력 후 검증 가능</td></tr>`;
  }
  html+=`</tbody></table></div>`;

  // 품번별 상세
  if(d.items.length){
    html+=`<div class="card">
      <details><summary style="cursor:pointer;font-size:15px;font-weight:700;color:var(--primary);padding:4px 0">품번별 기준인건비 산출 내역 (${d.items.length}건) ▸ 클릭하여 펼치기</summary>
      <div style="overflow-x:auto;max-height:500px;overflow-y:auto">
      <table><thead><tr><th>품번</th><th>품명</th><th>생산수량</th><th>기준/EA(수작업)</th><th>기준/EA(로터리)</th><th>소계(수작업)</th><th>소계(로터리)</th></tr></thead><tbody>`;
    for(const it of d.items){
      html+=`<tr><td class="c" style="font-weight:600">${it.pn}</td><td>${it.name}</td>
        <td class="r">${fmt0(it.qty)}</td><td class="r">${fmt(it.labor_m)}</td><td class="r">${fmt(it.labor_r)}</td>
        <td class="r" style="font-weight:600">${fmt0(it.sub_m)}</td><td class="r">${fmt0(it.sub_r)}</td></tr>`;
    }
    html+=`<tr style="background:var(--primary);color:#fff;font-weight:700">
      <td colspan="5" class="r">합계</td><td class="r">${fmt0(d.total_std_manual)}원</td><td class="r">${fmt0(d.total_std_rotary)}원</td></tr>`;
    html+=`</tbody></table></div></details></div>`;
  }
  document.getElementById('verifyResult').innerHTML=html;
}

// =========================================
// 생산 진척도
// =========================================
async function uploadPlan(){
  const file=document.getElementById('planFile').files[0];
  if(!file){document.getElementById('planStatus').innerHTML='<span style="color:var(--bad)">파일 선택</span>';return;}
  document.getElementById('planStatus').innerHTML='<span style="color:var(--accent)">업로드 중...</span>';
  const t=performance.now();
  const fd=new FormData();fd.append('file',file);
  const res=await fetch('/api/report/upload_plan',{method:'POST',body:fd});
  const d=await res.json();
  const sec=((performance.now()-t)/1000).toFixed(1);
  document.getElementById('planStatus').innerHTML=`<span style="color:${d.ok?'var(--good)':'var(--bad)'}">${d.msg} (${sec}초)</span>`;
}
let _progressData=null;
async function runProgress(){
  const days=document.getElementById('remainDays')?.value||17;
  const yr=document.getElementById('progressYear')?.value||'';
  const mn=document.getElementById('progressMonth')?.value||'';
  const month=yr&&mn?yr+'-'+mn:'';
  // 로딩 표시 (애니메이션)
  document.getElementById('progressResult').innerHTML=`
    <div style="text-align:center;padding:60px">
      <div style="width:40px;height:40px;border:4px solid var(--border);border-top:4px solid var(--accent);border-radius:50%;animation:spin 1s linear infinite;margin:0 auto"></div>
      <div style="color:var(--accent);margin-top:12px;font-weight:600">서버에서 분석 중...</div>
    </div>
    <style>@keyframes spin{to{transform:rotate(360deg)}}</style>`;
  const params=new URLSearchParams({days});
  if(month) params.set('month',month);
  const res=await fetch('/api/report/progress?'+params);
  const d=await res.json();
  _progressData=d;
  if(!d.categories.length){document.getElementById('progressResult').innerHTML='<div class="card" style="text-align:center;padding:30px;color:var(--muted)">판매계획을 먼저 업로드하세요</div>';return;}

  // 요약만 먼저 렌더 (빠름)
  let html='';
  html+=`<div class="stats">`;
  for(const c of d.categories){
    const color=c.rate>=50?'good':(c.rate>0?'warn':'bad');
    html+=`<div class="stat ${color}"><div class="label">${c.cat}</div><div class="value">${c.rate}%</div>
      <div class="detail">계획:${fmt0(c.plan)} 실적:${fmt0(c.prod)}</div>
      <div class="detail">일필요:${fmt0(c.daily)}/일${c.over?' ⚠️CAPA초과':''}</div></div>`;
  }
  html+=`</div>`;

  const overs=d.categories.filter(c=>c.over);
  if(overs.length){
    html+=`<div class="card" style="border-left:4px solid var(--bad)"><h3>⚠️ CAPA 초과 카테고리</h3><table>
      <tr><th>카테고리</th><th>일필요</th><th>기준CAPA</th><th>초과분</th><th>의견</th></tr>`;
    for(const c of overs){
      html+=`<tr class="bad-bg"><td>${c.cat}</td><td class="r">${fmt0(c.daily)}/일</td><td class="r">${fmt0(c.capa)}/일</td>
        <td class="r" style="color:var(--bad);font-weight:700">+${fmt0(c.daily-c.capa)}</td>
        <td>특근 필요 또는 인원 추가 배부</td></tr>`;
    }
    html+=`</table></div>`;
  }

  // 위험 품목 (상위 10건만)
  const dangers=d.items.filter(i=>i.status==='none'&&i.plan>=1000).sort((a,b)=>b.plan-a.plan);
  if(dangers.length){
    html+=`<div class="card"><h3>🔴 미착수 대량 품목 (${dangers.length}건 중 상위 10)</h3><table>
      <tr><th>품번</th><th>품명</th><th>카테고리</th><th>계획</th><th>잔여</th><th>일필요</th></tr>`;
    for(const i of dangers.slice(0,10)){
      html+=`<tr class="bad-bg"><td class="c" style="font-weight:600">${i.pn}</td><td>${i.name.substring(0,28)}</td><td class="c">${i.cat}</td>
        <td class="r">${fmt0(i.plan)}</td><td class="r" style="font-weight:700">${fmt0(i.remain)}</td>
        <td class="r" style="font-weight:700">${fmt0(i.daily)}/일</td></tr>`;
    }
    html+=`</table></div>`;
  }

  // 전체 테이블은 버튼 클릭 시 lazy 로드
  html+=`<div class="card"><button onclick="renderProgressTable()" id="btnProgressTable"
    style="padding:10px 20px;background:var(--accent);color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer;font-weight:600;width:100%">
    전체 품목 진척도 (${d.items.length}건) 보기</button><div id="progressTableArea"></div></div>`;

  document.getElementById('progressResult').innerHTML=html;
}

function renderProgressTable(){
  const d=_progressData;
  if(!d) return;
  const btn=document.getElementById('btnProgressTable');
  if(btn) btn.style.display='none';
  let html=`<div style="overflow-x:auto;max-height:500px;overflow-y:auto;margin-top:12px"><table>
    <tr><th>품번</th><th>품명</th><th>카테고리</th><th>계획</th><th>실적</th><th>달성률</th><th>잔여</th><th>일필요</th><th>상태</th></tr>`;
  for(const i of d.items){
    const sc=i.status==='done'?'var(--good)':i.status==='ok'?'var(--warn)':'var(--bad)';
    const label=i.status==='done'?'완료':i.status==='ok'?'진행':i.status==='delay'?'지연':'미착수';
    html+=`<tr><td class="c" style="font-weight:600">${i.pn}</td><td>${i.name.substring(0,28)}</td><td class="c">${i.cat}</td>
      <td class="r">${fmt0(i.plan)}</td><td class="r">${fmt0(i.prod)}</td>
      <td class="c"><span class="badge" style="background:${sc}20;color:${sc};font-weight:700">${i.rate}%</span></td>
      <td class="r">${fmt0(i.remain)}</td><td class="r">${fmt0(i.daily)}/일</td>
      <td class="c"><span class="badge" style="background:${sc}20;color:${sc}">${label}</span></td></tr>`;
  }
  html+=`</table></div>`;
  document.getElementById('progressTableArea').innerHTML=html;
}

// =========================================
// 월 결산
// =========================================
async function uploadSales(){
  const file=document.getElementById('salesFile').files[0];
  if(!file){document.getElementById('salesStatus').innerHTML='<span style="color:var(--bad)">파일 선택</span>';return;}
  document.getElementById('salesStatus').innerHTML='<span style="color:var(--accent)">업로드 중... (대용량 파일은 수 초 소요)</span>';
  const fd=new FormData();fd.append('file',file);
  const res=await fetch('/api/report/upload_sales',{method:'POST',body:fd});
  const d=await res.json();
  document.getElementById('salesStatus').innerHTML=`<span style="color:${d.ok?'var(--good)':'var(--bad)'}">${d.msg}</span>`;
}
let _settlementData=null;
async function runSettlement(){
  const yr=document.getElementById('settlementYear')?.value||'';
  const mn=document.getElementById('settlementMonth')?.value||'';
  const month=yr&&mn?yr+'-'+mn:'';
  document.getElementById('settlementResult').innerHTML=`
    <div style="text-align:center;padding:60px">
      <div style="width:40px;height:40px;border:4px solid var(--border);border-top:4px solid var(--accent);border-radius:50%;animation:spin 1s linear infinite;margin:0 auto"></div>
      <div style="color:var(--accent);margin-top:12px;font-weight:600">서버에서 분석 중...</div>
    </div>`;
  const params=new URLSearchParams();
  if(month) params.set('month',month);
  const res=await fetch('/api/report/settlement?'+params);
  const d=await res.json();
  _settlementData=d;
  if(!d.items.length){document.getElementById('settlementResult').innerHTML='<div class="card" style="text-align:center;padding:30px;color:var(--muted)">매출현황을 먼저 업로드하세요</div>';return;}

  let html=`<div class="stats">
    <div class="stat"><div class="label">분석 품목</div><div class="value">${d.items.length}종</div></div>
    <div class="stat"><div class="label">총 생산량</div><div class="value">${fmt0(d.total_prod)}EA</div></div>
    <div class="stat"><div class="label">총 판매량</div><div class="value">${fmt0(d.total_sales)}EA</div></div>
    <div class="stat bad"><div class="label">재고 과잉</div><div class="value">${d.over_stock.length}건</div><div class="detail">생산>판매 150% 초과</div></div>
    <div class="stat warn"><div class="label">품절 주의</div><div class="value">${d.under_stock.length}건</div><div class="detail">생산<판매 80% 미만</div></div>
  </div>`;

  // 재고 과잉
  if(d.over_stock.length){
    html+=`<div class="card" style="border-left:4px solid var(--bad)"><h3>📦 재고 과잉 (생산 > 판매 150%)</h3><table>
      <tr><th>품번</th><th>품명</th><th>생산</th><th>판매</th><th>차이</th><th>비율</th></tr>`;
    for(const i of d.over_stock){
      html+=`<tr class="bad-bg"><td class="c" style="font-weight:600">${i.pn}</td><td>${i.name}</td>
        <td class="r">${fmt0(i.prod)}</td><td class="r">${fmt0(i.sales)}</td>
        <td class="r" style="color:var(--bad);font-weight:700">+${fmt0(i.diff)}</td>
        <td class="c"><span class="badge badge-bad">${i.ratio}%</span></td></tr>`;
    }
    html+=`</table></div>`;
  }

  // 품절 주의
  if(d.under_stock.length){
    html+=`<div class="card" style="border-left:4px solid var(--warn)"><h3>⚠️ 품절 주의 (생산 < 판매 80%)</h3><table>
      <tr><th>품번</th><th>품명</th><th>생산</th><th>판매</th><th>부족</th><th>비율</th></tr>`;
    for(const i of d.under_stock){
      html+=`<tr style="background:#fffbeb"><td class="c" style="font-weight:600">${i.pn}</td><td>${i.name}</td>
        <td class="r">${fmt0(i.prod)}</td><td class="r">${fmt0(i.sales)}</td>
        <td class="r" style="color:var(--warn);font-weight:700">${fmt0(i.diff)}</td>
        <td class="c"><span class="badge badge-warn">${i.ratio}%</span></td></tr>`;
    }
    html+=`</table></div>`;
  }

  // 전체 테이블은 버튼 클릭 시 lazy 로드
  html+=`<div class="card"><button onclick="renderSettlementTable()" id="btnSettlementTable"
    style="padding:10px 20px;background:var(--accent);color:#fff;border:none;border-radius:8px;font-size:13px;cursor:pointer;font-weight:600;width:100%">
    생산 vs 판매 통합 실적표 (${d.items.length}건) 보기</button><div id="settlementTableArea"></div></div>`;

  document.getElementById('settlementResult').innerHTML=html;
}

function renderSettlementTable(){
  const d=_settlementData;
  if(!d) return;
  const btn=document.getElementById('btnSettlementTable');
  if(btn) btn.style.display='none';
  let html=`<div style="overflow-x:auto;max-height:500px;overflow-y:auto;margin-top:12px"><table>
    <tr><th>품번</th><th>품명</th><th>카테고리</th><th>생산량</th><th>판매량</th><th>재고변동</th><th>비율</th><th>기준인건비/EA</th><th>인건비합계</th><th>상태</th></tr>`;
  for(const i of d.items){
    const sc=i.status==='over'?'var(--bad)':i.status==='under'?'var(--warn)':'var(--good)';
    const label=i.status==='over'?'과잉':i.status==='under'?'부족':'정상';
    html+=`<tr><td class="c" style="font-weight:600">${i.pn}</td><td>${i.name}</td><td class="c">${i.cat}</td>
      <td class="r">${fmt0(i.prod)}</td><td class="r">${fmt0(i.sales)}</td>
      <td class="r" style="font-weight:700;color:${i.diff>=0?'var(--good)':'var(--bad)'}">${i.diff>=0?'+':''}${fmt0(i.diff)}</td>
      <td class="c">${i.ratio>0?i.ratio+'%':'-'}</td>
      <td class="r">${i.labor_ea?fmt(i.labor_ea):'-'}</td>
      <td class="r">${i.std_labor_total?fmt0(i.std_labor_total):'-'}</td>
      <td class="c"><span class="badge" style="background:${sc}20;color:${sc}">${label}</span></td></tr>`;
  }
  html+=`</table></div>`;
  document.getElementById('settlementTableArea').innerHTML=html;
}

// =========================================
// 원부재료 단가 관리
// =========================================
async function loadMaterials(){
  const res=await fetch('/api/materials');
  const d=await res.json();
  const tbody=document.getElementById('matBody');
  if(!tbody) return;
  let html='';
  d.materials.forEach((m,i)=>{
    html+=`<tr class="mat-row">
      <td class="c">${i+1}</td>
      <td class="c" style="font-weight:600">${m.pn}</td>
      <td>${m.name}</td>
      <td class="c"><input type="number" value="${m.price}" id="matP-${i}" data-pn="${m.pn}"
        style="width:100%;padding:5px 8px;border:1px solid var(--border);border-radius:6px;text-align:right;font-size:12px"
        onchange="this.style.borderColor='var(--warn)';this.style.background='#fffbeb'"></td>
      <td class="c" style="font-size:11px;color:var(--muted)">${m.date||''}</td>
      <td class="c"><button onclick="saveMaterial(${i})" style="padding:4px 12px;background:var(--accent);color:#fff;border:none;border-radius:5px;font-size:11px;cursor:pointer">저장</button></td>
    </tr>`;
  });
  tbody.innerHTML=html;
}

async function saveMaterial(idx){
  const input=document.getElementById('matP-'+idx);
  const pn=input.dataset.pn;
  const price=parseFloat(input.value)||0;
  const res=await fetch('/api/materials/update',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pn,price})});
  const d=await res.json();
  if(d.ok){
    input.style.borderColor='var(--good)';input.style.background='#f0fff4';
    document.getElementById('matStatus').innerHTML=`<span style="color:var(--good)">${d.msg}</span>`;
    setTimeout(()=>{input.style.borderColor='var(--border)';input.style.background='#fff'},1500);
  } else {alert(d.msg);}
}

async function addMaterial(){
  const pn=document.getElementById('matPn').value.trim();
  const price=parseFloat(document.getElementById('matPrice').value)||0;
  if(!pn){document.getElementById('matStatus').innerHTML='<span style="color:var(--bad)">품번을 입력하세요</span>';return;}
  const res=await fetch('/api/materials/add',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({pn,price})});
  const d=await res.json();
  document.getElementById('matStatus').innerHTML=`<span style="color:${d.ok?'var(--good)':'var(--bad)'}">${d.msg}</span>`;
  if(d.ok){document.getElementById('matPn').value='';document.getElementById('matPrice').value='';loadMaterials();}
}

function filterMatTable(){
  const q=(document.getElementById('matSearch')?.value||'').toLowerCase();
  document.querySelectorAll('#matTable .mat-row').forEach(r=>{
    r.style.display=r.textContent.toLowerCase().includes(q)?'':'none';
  });
}

// 탭 전환 시 데이터 로드
const origShowTab=showTab;
showTab=function(n){
  origShowTab(n);
  if(n===1) loadProdRecords();
  if(n===4) initVerify();
  if(n===7){loadEmployees();loadMaterials();}
};
document.addEventListener('DOMContentLoaded',()=>{
  if(document.querySelector('#p7.active')){loadEmployees();loadMaterials();}
});
</script>
</body>
</html>
"""

# ============================================================
# 로고 서빙
# ============================================================
LOGO_PATH = os.path.join(BASE_DIR, 'logo.png') if os.path.exists(os.path.join(BASE_DIR, 'logo.png')) else os.path.join(BASE_DIR, '기업CI.png')

@app.route('/logo.png')
def serve_logo():
    from flask import send_file
    return send_file(LOGO_PATH, mimetype='image/png')

# ============================================================
# 로그인/로그아웃
# ============================================================
LOGIN_HTML = r"""
<!DOCTYPE html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>로그인 - Maehong L&F</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Pretendard','맑은 고딕',sans-serif;background:linear-gradient(135deg,#0055a5 0%,#0077cc 40%,#5cb85c 80%,#7BC142 100%);min-height:100vh;display:flex;align-items:center;justify-content:center}
.login-box{background:#fff;border-radius:20px;padding:48px 40px 40px;width:400px;box-shadow:0 24px 80px rgba(0,0,0,.25);text-align:center}
.login-box .logo{margin-bottom:16px}
.login-box .logo img{height:80px}
.login-box .system-name{font-size:14px;color:#555;margin-bottom:32px;letter-spacing:1px}
.login-box form{text-align:left}
.login-box label{display:block;font-size:12px;color:#888;margin-bottom:4px;margin-top:16px;font-weight:600}
.login-box input{width:100%;padding:11px 14px;border:1px solid #ddd;border-radius:10px;font-size:14px;transition:.2s}
.login-box input:focus{outline:none;border-color:#0066cc;box-shadow:0 0 0 3px rgba(0,102,204,.12)}
.login-box button{width:100%;padding:13px;background:linear-gradient(135deg,#0066cc,#7BC142);color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:700;cursor:pointer;margin-top:28px;transition:.2s;letter-spacing:.5px}
.login-box button:hover{opacity:.9;transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,102,204,.3)}
.err{color:#e53e3e;font-size:13px;text-align:center;margin-top:14px}
.login-footer{font-size:11px;color:#aaa;margin-top:20px;text-align:center}
</style></head><body>
<div class="login-box">
  <div class="logo"><img src="/logo.png" alt="Maehong L&F"></div>
  <div class="system-name">통합 원가 관리 시스템</div>
  <form method="POST" action="/login">
    <label>아이디</label>
    <input name="username" type="text" placeholder="아이디를 입력하세요" autofocus required>
    <label>비밀번호</label>
    <input name="password" type="password" placeholder="비밀번호를 입력하세요" required>
    <button type="submit">로그인</button>
  </form>
  {% if error %}<div class="err">{{ error }}</div>{% endif %}
  <div class="login-footer">Maehong L&F · Living & Food</div>
</div>
</body></html>
"""

@app.route('/login', methods=['GET','POST'])
def login():
    if request.method == 'POST':
        uid = request.form.get('username','').strip()
        pwd = request.form.get('password','')
        u = USERS.get(uid)
        if u and u['password'] == pwd:
            session['user'] = uid
            session['role'] = u['role']
            session['name'] = u['name']
            return redirect('/')
        return render_template_string(LOGIN_HTML, error='아이디 또는 비밀번호가 일치하지 않습니다')
    return render_template_string(LOGIN_HTML, error=None)

@app.route('/logout')
def logout():
    session.clear()
    return redirect('/login')

# ============================================================
# 메인 라우트
# ============================================================
@app.route('/')
@login_required
def index():
    cost_rows = []
    for pn in sorted(products.keys(), key=lambda x: (products[x]['category'], x)):
        p = products[pn]; ac = all_costs[pn]
        cost_rows.append(type('R',(),{
            'pn':pn,'name':p['name'],'cat':p['category'],'weight':p['weight_g'] or 0,'ptype':p['type'],
            'raw':ac['m']['raw'],'sub':ac['m']['sub'],
            'labor_m':ac['m']['labor'],'labor_r':ac['r']['labor'],
            'total_m':ac['m']['total'],'total_r':ac['r']['total'],
        })())

    labors = [c.labor_m for c in cost_rows if c.labor_m > 0]
    mats = [c.raw + c.sub for c in cost_rows if c.raw + c.sub > 0]
    tots = [c.total_m for c in cost_rows if c.total_m > 0]

    issue_rows = []
    for pn in sorted(products.keys()):
        p = products[pn]
        if '고구마' not in p['category'] or '바' in p['category']: continue
        ac = all_costs[pn]
        if ac['r']['labor'] <= 0: continue
        diff = ac['m']['labor'] - ac['r']['labor']
        pct = diff / ac['r']['labor'] * 100
        issue_rows.append(type('R',(),{'pn':pn,'name':p['name'],'ptype':p['type'],'labor_r':ac['r']['labor'],'labor_m':ac['m']['labor'],'diff':diff,'pct':pct})())
    issue_rows.sort(key=lambda x: -x.diff)

    wage_rows = []
    for proc_name, workers in [('절단',['김흥수','한승엽','박미영','송선임']),('선별',['정미혜','아다치에리','이소현']),
             ('내포장',['서혜진','김하윤']),('나라시',['최엘라']),('살균',['사만','선티조이']),('선날인',['김영미']),('외포장',['권미정'])]:
        for w in workers:
            e = employee_wages.get(w,{}); pay = e.get('pay',0) if e else 0
            hourly_w = round(pay*(1+RETIRE_RATE)/H, 2) if pay>0 else round(MIN_WAGE*(1+RETIRE_RATE), 2)
            wage_rows.append(type('W',(),{'proc':proc_name,'name':w,'pay':pay,'hourly':hourly_w,'note':'최저시급' if pay==0 else ''})())

    g = all_costs.get('G0010',{}).get('m',{})
    g_mats = [type('M',(),m)() for m in g.get('mat_items',[])]
    colors = ['#2b6cb0','#38a169','#d69e2e','#e53e3e','#805ad5','#dd6b20']
    g_labor_items = []
    for i,(proc,cost,_,_) in enumerate(g.get('labor_items',[])):
        pct = cost/g['labor']*100 if g.get('labor',0)>0 else 0
        g_labor_items.append(type('L',(),{'proc':proc,'cost':cost,'pct':pct,'color':colors[i%len(colors)]})())

    return render_template_string(HTML,
        now=datetime.now().strftime('%Y-%m-%d'), min_wage=f'{MIN_WAGE:,}',
        total_products=len(products), cost_rows=cost_rows,
        avg_labor=f"{sum(labors)/len(labors):,.0f}" if labors else '0',
        avg_mat=f"{sum(mats)/len(mats):,.0f}" if mats else '0',
        avg_total=f"{sum(tots)/len(tots):,.0f}" if tots else '0',
        p_rot=P['로터리'], p_man=P['수작업'], p_naet=P['낱봉'], p_bund=P['번들'],
        issue_rows=issue_rows, wage_rows=wage_rows,
        g_raw=g.get('raw',0), g_sub=g.get('sub',0),
        g_labor=g.get('labor',0), g_total=g.get('total',0),
        g_mats=g_mats, g_labor_items=g_labor_items,
        emp_count=len(employee_wages),
        prod_count=len(prod_records),
        is_admin=is_admin(), user_name=session.get('name',''),
    )

if __name__ == '__main__':
    print('\n' + '='*50)
    print('  매홍엘앤에프 통합 원가 관리 시스템 v2')
    print('  http://localhost:5000')
    print('='*50 + '\n')
    app.run(host='0.0.0.0', port=5000, debug=True, use_reloader=False)
