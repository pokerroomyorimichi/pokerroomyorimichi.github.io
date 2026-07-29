#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
YORIMICHI ダッシュボード 静的データ生成エンジン
  旧 GAS Web App (Code.gs) の doGet 8アクションを Python に忠実移植し、
  data.json を1本生成する。GitHub Actions から毎朝実行される。

  出力: dashboard/data.json（index.html の api() が fetch する）
"""
import os, json, warnings, re, math, calendar, datetime as dt
warnings.filterwarnings("ignore")
from zoneinfo import ZoneInfo
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# ===== 定数（Code.gs と一致させる）=====
MASTER_SS_ID   = '1QLu2OVgN66MBXbNjtodLSrefnnEnatLQdeAnV0jnsVw'
SURVEY_SS_ID   = '17iU4i2FQjdMJ-_ccSJ-fYJDZ5kUkM7fOasAlQNbtMyM'
TIMECARD_SS_ID = '1iTwt9GHkI-41sODIPI1rqGm6wbaEasYKljGsrbq_iXI'
BEP = 1771671
TZ  = ZoneInfo('Asia/Tokyo')

TARGETS = {
    '2026/04': {'sales': 2900000, 'customers': 400},
    '2026/05': {'sales': 3000000, 'customers': 430},
    '2026/06': {'sales': 2700000, 'customers': 450},
    '2026/07': {'sales': 2900000, 'customers': None},
    '2026/08': {'sales': 3200000, 'customers': None},
    '2026/09': {'sales': 3100000, 'customers': None},
}

# import_収支記録 の列インデックス（0始まり・Code.gs COL と一致）
COL = {
    'date':0,'dow':1,'kind':2,'total':3,'cash':4,'card':5,'emoney':6,'cust':7,'unit':8,
    'nyujo':9,'taiken':10,'chip':11,'toname':12,
    'tanpin_al':13,'tanpin_na':14,'free_al':15,'free_na':16,'bottle':17,'staff_dr':18,'food':19,
    'other':20,'sale':21,'rake':22,'commission':23,'pd':24,'tokuten':25,
}
CATEGORIES = {
    'game':  {'cols': ['nyujo','taiken','chip','toname']},
    'drink': {'cols': ['tanpin_al','tanpin_na','free_al','free_na','bottle','staff_dr','food']},
    'other': {'cols': ['other','sale']},
    'rake':  {'cols': ['rake','commission','pd']},
}
DOW = ['日','月','火','水','木','金','土']

# ===== ユーティリティ =====
def parse_num(v):
    # Code.gs の parseInt(String(v).replace(/[¥,\s人]/g,'')) を再現。
    # 'pts' 等の接尾辞が付いても JS parseInt は先頭の数値だけを拾うため、
    # 先頭の符号付き整数だけを抽出する（例 '4,666 pts' → 4666, '-500 pts' → -500）。
    if v is None or v == '':
        return 0
    s = str(v)
    for ch in '¥,　 人':
        s = s.replace(ch, '')
    m = re.match(r'^\s*(-?\d+)', s)
    return int(m.group(1)) if m else 0

def rhu(x):
    # JS Math.round（四捨五入・0.5は+∞方向）を再現。Python 標準 round は銀行丸めのため使わない。
    return math.floor(x + 0.5)

def sub_months(d, n):
    # JS Date.setMonth(getMonth()-n) 相当。カレンダー月で n ヶ月戻し、日はクランプ。
    y, m = d.year, d.month - n
    while m <= 0:
        m += 12; y -= 1
    day = min(d.day, calendar.monthrange(y, m)[1])
    return dt.date(y, m, day)

def now_jst():
    return dt.datetime.now(TZ)

def fmt_date(d):
    return d.strftime('%Y/%m/%d')

def parse_cell_date(v):
    """セルの日付文字列を date に。'2026/01/03' or '2026-01-03' 想定。失敗時 None"""
    if v is None or v == '':
        return None
    s = str(v).strip().replace('-', '/')
    parts = s.split('/')
    if len(parts) < 3:
        return None
    try:
        return dt.date(int(parts[0]), int(parts[1]), int(parts[2][:2]))
    except (ValueError, IndexError):
        return None

def round1(x):
    return rhu(x * 10) / 10

# ===== Sheets 読み取り =====
def get_service():
    tok = os.environ.get('TOKEN_JSON_PATH',
                         os.path.expanduser('~/.beebloom/07_Development/credentials/token.json'))
    creds = Credentials.from_authorized_user_file(tok)
    return build('sheets', 'v4', credentials=creds, cache_discovery=False), \
           build('drive', 'v3', credentials=creds, cache_discovery=False)

def read_values(svc, sid, rng, value_render='FORMATTED_VALUE'):
    # 数値の多いシート（収支記録・来客数）は UNFORMATTED_VALUE で生値を読む。
    # 表示用フォーマットは小数を丸めるため、GAS の getValues()（生値）+ parseInt と
    # 一致させるには生値が必要。日付は FORMATTED_STRING で '2026/01/03' 形式に保つ。
    try:
        r = svc.spreadsheets().values().get(
            spreadsheetId=sid, range=rng,
            valueRenderOption=value_render,
            dateTimeRenderOption='FORMATTED_STRING').execute()
        return r.get('values', [])
    except Exception:
        return []

# ===== 方針B：月ブロック自動追記（Drive自動発見）=====
# 店舗は毎月「別スプレッドシート」を新規作成する。統合マスタの import タブへ
# 当月・翌月ぶんの IMPORTRANGE ブロックが未反映なら、列Aの最終データ行の直下へ
# 追記する。既存ブロックは一切変更しない（末尾追記のみ）。
# 実書き込みは環境変数 APPEND_MONTHS=1 のときのみ（既定は dry-run でログ出力）。
IMPORT_TABS = {
    'import_収支記録_2026': {
        'src_name': '収支記録 {y}年{m}月',        # Drive 命名の安定部分（接頭辞 ★G は揺れるため除外）
        'src_tab':  '売上管理',
        'gid':      '296529933',
        'wrap':     True,   # ={ ... } の配列括り・カンマ後スペースなし
    },
    'import_来客数_2026': {
        'src_name': '来客数管理シート {y}年{m}月',  # 接頭辞 ★H/★F は揺れるため除外
        'src_tab':  '来客数管理シート',
        'gid':      '2092809373',
        'wrap':     False,  # 平文・カンマ後スペースあり
    },
}

def _find_source_id(drive, name_contains):
    q = (f"name contains '{name_contains}' "
         "and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false")
    res = drive.files().list(q=q, fields="files(id,name)", pageSize=10,
                             includeItemsFromAllDrives=True, supportsAllDrives=True).execute()
    files = res.get('files', [])
    return (files[0]['id'], files[0]['name']) if files else (None, None)

def _months_present(sheets, tab):
    col = sheets.spreadsheets().values().get(
        spreadsheetId=MASTER_SS_ID, range=f"'{tab}'!A1:A400",
        valueRenderOption='FORMATTED_VALUE', dateTimeRenderOption='FORMATTED_STRING'
    ).execute().get('values', [])
    present, last = set(), 0
    for i, r in enumerate(col, start=1):
        s = str(r[0]).strip() if r else ''
        if s:
            last = i
        d = parse_cell_date(s)
        if d:
            present.add((d.year, d.month))
    return present, last

def _build_import_formula(cfg, sid, end_row):
    url = f"https://docs.google.com/spreadsheets/d/{sid}/edit?gid={cfg['gid']}#gid={cfg['gid']}"
    if cfg['wrap']:
        return ('={\n  IMPORTRANGE("' + url + '","' + cfg['src_tab'] + f'!B4:AE{end_row}")\n}}')
    return ('=IMPORTRANGE("' + url + '", "' + cfg['src_tab'] + f'!B4:AE{end_row}")')

def ensure_month_blocks(sheets, drive, dry_run=True):
    """当月・翌月の未反映ブロックを検出し、末尾に追記（dry_run時はログのみ）。"""
    now = now_jst()
    targets = [(now.year, now.month)]
    ny, nm = (now.year, now.month + 1) if now.month < 12 else (now.year + 1, 1)
    targets.append((ny, nm))
    logs = []
    for tab, cfg in IMPORT_TABS.items():
        try:
            present, last = _months_present(sheets, tab)
        except Exception as e:
            logs.append(f"[{tab}] present取得失敗: {str(e)[:60]} → スキップ")
            continue
        for (y, m) in targets:
            if (y, m) in present:
                continue
            name = cfg['src_name'].format(y=y, m=m)
            try:
                sid, sname = _find_source_id(drive, name)
            except Exception as e:
                logs.append(f"[{tab}] {y}/{m}: Drive検索失敗 {str(e)[:50]}")
                continue
            if not sid:
                logs.append(f"[{tab}] {y}/{m}: ソース未発見（店舗が未作成）→ スキップ")
                continue
            days = calendar.monthrange(y, m)[1]
            end = days + 3
            cell = f"'{tab}'!A{last + 1}"
            formula = _build_import_formula(cfg, sid, end)
            if dry_run:
                logs.append(f"[{tab}] {y}/{m}: 追記予定 {cell} ← '{sname}' (B4:AE{end})")
            else:
                sheets.spreadsheets().values().update(
                    spreadsheetId=MASTER_SS_ID, range=cell,
                    valueInputOption='USER_ENTERED', body={'values': [[formula]]}
                ).execute()
                logs.append(f"[{tab}] {y}/{m}: ✅追記完了 {cell} ← '{sname}'")
                last += days  # 次の追記位置を更新（IMPORTRANGEは非同期のため手動加算）
    return logs

# ===== 生データ取得（Code.gs getDailyRows_ / getCustomerMap_ 相当）=====
def get_daily_rows(rows, months):
    """収支記録 → 営業日のみ [{date,sales,customers}]（Code.gs getDailyRows_ 相当）"""
    cutoff = sub_months(now_jst().date(), months)
    out = []
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        d = parse_cell_date(r[0])
        if d is None or d <= cutoff:
            continue
        kind = str(r[COL['kind']]) if len(r) > COL['kind'] else ''
        if kind != '営':
            continue
        def cell(k):
            i = COL[k]
            return r[i] if len(r) > i else ''
        out.append({
            'date': fmt_date(d),
            'sales': parse_num(cell('total')),
            'customers': parse_num(cell('cust')),
        })
    return out

def get_customer_map(rows, months):
    cutoff = sub_months(now_jst().date(), months)
    m = {}
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        d = parse_cell_date(r[0])
        if d is None or d <= cutoff:
            continue
        key = fmt_date(d)
        total  = parse_num(r[3]) if len(r) > 3 else 0
        repeat = parse_num(r[4]) if len(r) > 4 else 0
        newc   = parse_num(r[5]) if len(r) > 5 else 0
        m[key] = {'total': total, 'repeat': repeat, 'newC': newc}
    return m

# ===== SUMMARY =====
def compute_summary(daily_all, cmap_all):
    now = now_jst()
    ym = now.strftime('%Y/%m')
    target = TARGETS.get(ym, {})
    days_in_month = ((now.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
                     - dt.timedelta(days=1)).day

    mSales = mCust = mNew = mRepeat = mDays = 0
    for r in daily_all:
        if not r['date'].startswith(ym):
            continue
        mSales += r['sales']; mCust += r['customers']; mDays += 1
        c = cmap_all.get(r['date'], {})
        mNew += c.get('newC', 0); mRepeat += c.get('repeat', 0)

    projected = rhu(mSales / mDays * days_in_month) if mDays > 0 else 0
    trend7 = sorted([r for r in daily_all if r['date'].startswith(ym)],
                    key=lambda r: r['date'], reverse=True)[:7]
    trend7 = [{'date': r['date'][5:], 'sales': r['sales'], 'customers': r['customers']}
              for r in reversed(trend7)]
    unit = rhu(mSales / mCust) if mCust > 0 else 0

    return {
        'ok': True,
        'updatedAt': now.strftime('%Y/%m/%d %H:%M'),
        'month': ym,
        'sales': {
            'actual': mSales,
            'target': target.get('sales'),
            'projected': projected,
            'bep': BEP,
            'rate': round1(mSales / target['sales'] * 100) if target.get('sales') else None,
            'bepRate': round1(mSales / BEP * 100),
            'remaining': (target['sales'] - mSales) if target.get('sales') else None,
        },
        'customers': {
            'total': mCust,
            'target': target.get('customers'),
            'new': mNew,
            'repeat': mRepeat,
            'rate': round1(mCust / target['customers'] * 100) if target.get('customers') else None,
        },
        'unitPrice': {'actual': unit, 'target': 7000},
        'trend7': trend7,
    }

# ===== HISTORY =====
def compute_history(daily_all, cmap_all):
    monthly = {}
    for r in daily_all:
        ym = r['date'][:7]
        m = monthly.setdefault(ym, {'sales':0,'customers':0,'new':0,'repeat':0,'days':0})
        m['sales'] += r['sales']; m['customers'] += r['customers']; m['days'] += 1
        c = cmap_all.get(r['date'], {})
        m['new'] += c.get('newC',0); m['repeat'] += c.get('repeat',0)

    months = sorted(monthly.keys())
    cumS = cumC = 0
    data = []
    for idx, ym in enumerate(months):
        m = monthly[ym]; tgt = TARGETS.get(ym, {})
        cumS += m['sales']; cumC += m['customers']
        prev = monthly[months[idx-1]]['sales'] if idx > 0 else None
        growth = round1((m['sales']-prev)/prev*100) if prev else None
        data.append({
            'label': ym, 'sales': m['sales'], 'customers': m['customers'],
            'newC': m['new'], 'repeat': m['repeat'], 'days': m['days'],
            'unitPrice': rhu(m['sales']/m['customers']) if m['customers']>0 else 0,
            'target': tgt.get('sales'), 'custTarget': tgt.get('customers'),
            'rate': round1(m['sales']/tgt['sales']*100) if tgt.get('sales') else None,
            'bepRate': round1(m['sales']/BEP*100),
            'growthRate': growth, 'cumSales': cumS, 'cumCustomers': cumC,
        })
    best  = max(data, key=lambda d: d['sales']) if data else {}
    worst = min(data, key=lambda d: d['sales']) if data else {}
    return {
        'ok': True, 'data': data,
        'summary': {
            'totalMonths': len(data), 'cumSales': cumS, 'cumCustomers': cumC,
            'avgSales': rhu(cumS/len(data)) if data else 0,
            'avgCust': rhu(cumC/len(data)) if data else 0,
            'bestMonth': best.get('label','-'), 'bestSales': best.get('sales',0),
            'worstMonth': worst.get('label','-'), 'worstSales': worst.get('sales',0),
            'bepMonths': len([d for d in data if d['bepRate']>=100]),
        },
    }

# ===== SALES =====
def compute_sales_monthly(daily_all):
    monthly = {}
    for r in daily_all:
        ym = r['date'][:7]
        m = monthly.setdefault(ym, {'sales':0,'customers':0})
        m['sales'] += r['sales']; m['customers'] += r['customers']
    data = [{'label':ym,'sales':monthly[ym]['sales'],'customers':monthly[ym]['customers'],
             'target':(TARGETS.get(ym,{}) or {}).get('sales'),'bep':BEP}
            for ym in sorted(monthly)]
    return {'ok': True, 'period': 'monthly', 'data': data}

def compute_sales_daily(daily_all, cmap_all):
    data = []
    for r in sorted(daily_all, key=lambda r: r['date']):
        d = parse_cell_date(r['date'])
        c = cmap_all.get(r['date'], {})
        data.append({
            'date': r['date'], 'dow': DOW[d.weekday()!=6 and d.isoweekday()%7 or 0],
            'sales': r['sales'], 'customers': r['customers'],
            'newC': c.get('newC',0), 'repeat': c.get('repeat',0),
            'unitPrice': rhu(r['sales']/r['customers']) if r['customers']>0 else 0,
        })
    return {'ok': True, 'period': 'daily', 'data': data}

# ===== BREAKDOWN =====
def compute_breakdown(rows, months):
    cutoff = sub_months(now_jst().date(), min(months, 24))
    detail_keys = ['nyujo','taiken','chip','toname','tanpin_al','tanpin_na','free_al','free_na',
                   'bottle','staff_dr','food','sale','commission','pd','cash','card','emoney']
    parsed = []
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        d = parse_cell_date(r[0])
        if d is None or d <= cutoff:
            continue
        if (str(r[COL['kind']]) if len(r)>COL['kind'] else '') != '営':
            continue
        row = {'date': fmt_date(d), 'ym': fmt_date(d)[:7]}
        for k,i in COL.items():
            row[k] = parse_num(r[i]) if len(r)>i else 0
        for cat,defn in CATEGORIES.items():
            row[cat] = sum(row.get(c,0) for c in defn['cols'])
        parsed.append(row)

    monthly = {}
    for r in parsed:
        m = monthly.get(r['ym'])
        if m is None:
            m = {'label':r['ym'],'total':0,'cust':0,'days':0,'game':0,'drink':0,'other':0,'rake':0}
            for k in detail_keys: m[k]=0
            monthly[r['ym']] = m
        m['total'] += r['total']; m['cust'] += r['cust']; m['days'] += 1
        for c in CATEGORIES: m[c] += r[c]
        for k in detail_keys: m[k] += r.get(k,0)

    mData = [monthly[ym] for ym in sorted(monthly)]
    totals = {'game':0,'drink':0,'other':0,'rake':0,'total':0}
    for m in mData:
        for k in totals: totals[k] += m[k]
    return {'ok': True, 'period':'monthly', 'data': mData, 'totals': totals, 'months': months}

# ===== YESTERDAY =====
def compute_yesterday(daily_all, cmap_all, tc_service):
    now = now_jst()
    yday = now.date() - dt.timedelta(days=1)
    yd = fmt_date(yday)
    dow = yday.isoweekday() % 7  # 0=日

    ydRow = next((r for r in daily_all if r['date']==yd), {'sales':0,'customers':0})
    ydCust = cmap_all.get(yd, {'total':0,'newC':0,'repeat':0})

    last7 = sorted([r for r in daily_all if r['date']<=yd],
                   key=lambda r:r['date'], reverse=True)[:7][::-1]
    same = [r for r in daily_all
            if parse_cell_date(r['date']).isoweekday()%7==dow and r['date']<yd]
    same = sorted(same, key=lambda r:r['date'], reverse=True)[:4]
    avgS = rhu(sum(r['sales'] for r in same)/len(same)) if same else 0
    avgC = rhu(sum(r['customers'] for r in same)/len(same)) if same else 0
    vsS = round1((ydRow['sales']-avgS)/avgS*100) if avgS>0 else None
    vsC = round1((ydRow['customers']-avgC)/avgC*100) if avgC>0 else None

    staffYd = compute_staff_punches_for_date(tc_service, yd)
    return {
        'ok': True, 'date': yd, 'dowJa': DOW[dow],
        'sales': ydRow['sales'], 'customers': ydRow['customers'],
        'newC': ydCust.get('newC',0), 'repeat': ydCust.get('repeat',0),
        'unitPrice': rhu(ydRow['sales']/ydRow['customers']) if ydRow['customers']>0 else 0,
        'bepRate': rhu(ydRow['sales']/(BEP/30)*10)/10,
        'avgSales': avgS, 'avgCust': avgC, 'vsAvgSales': vsS, 'vsAvgCust': vsC,
        'last7': last7,
        'staffYd': sorted(staffYd, key=lambda s:s['workMin'], reverse=True),
        'sameDow': [{'date':r['date'],'sales':r['sales'],'customers':r['customers']} for r in same],
    }

# ===== タイムカード（現行の月別タブ 打刻ログ_YYYY-MM に対応）=====
def tc_time_to_sec(t):
    p = str(t).split(':')
    h = int(p[0]) if p[0].isdigit() else 0
    m = int(p[1]) if len(p)>1 and p[1].isdigit() else 0
    s = h*3600 + m*60
    if h < 8:
        s += 86400
    return s

def tc_work_min(punches):
    clockIn=None; brk=None; work=0
    for p in punches:
        s=tc_time_to_sec(p['time']); t=p['type']
        if t=='出勤': clockIn=s; brk=None
        elif t=='休憩開始' and clockIn is not None: work+=s-clockIn; clockIn=None; brk=s
        elif t=='休憩終了' and brk is not None: brk=None; clockIn=s
        elif t=='退勤' and clockIn is not None: work+=s-clockIn; clockIn=None
    return rhu(work/60)

def tc_late_min(punches):
    LATE=22*3600; late=0; clockIn=None; brk=None
    for p in punches:
        s=tc_time_to_sec(p['time']); t=p['type']
        if t=='出勤': clockIn=s; brk=None
        elif t=='休憩開始' and clockIn is not None:
            if s>LATE: late+=rhu((s-max(clockIn,LATE))/60)
            clockIn=None; brk=s
        elif t=='休憩終了' and brk is not None: brk=None; clockIn=s
        elif t=='退勤' and clockIn is not None:
            if s>LATE: late+=rhu((s-max(clockIn,LATE))/60)
            clockIn=None
    return late

def tc_month_tab(ym_slash):
    """'2026/07' → '打刻ログ_2026-07'"""
    return '打刻ログ_' + ym_slash.replace('/', '-')

def read_timecard_rows(svc, ym_slash):
    tab = tc_month_tab(ym_slash)
    rows = read_values(svc, TIMECARD_SS_ID, f"'{tab}'!A1:F2000")
    return rows

def compute_staff_punches_for_date(svc, date_str):
    ym = date_str[:7]
    rows = read_timecard_rows(svc, ym)
    byStaff = {}
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        date = parse_cell_date(r[3]) if len(r)>3 else None
        if date is None or fmt_date(date) != date_str:
            continue
        staff = str(r[1]) if len(r)>1 else ''
        typ   = str(r[2]) if len(r)>2 else ''
        time  = str(r[4])[:5] if len(r)>4 else ''
        byStaff.setdefault(staff, []).append({'type':typ,'time':time})
    out=[]
    for name,punches in byStaff.items():
        inT  = next((p['time'] for p in punches if p['type']=='出勤'), '-')
        outT = next((p['time'] for p in reversed(punches) if p['type']=='退勤'), '-')
        out.append({'name':name,'workMin':tc_work_min(punches),'lateMin':tc_late_min(punches),
                    'inT':inT,'outT':outT})
    return out

def compute_staff(svc):
    now = now_jst(); ym = now.strftime('%Y/%m'); today = fmt_date(now.date())
    rows = read_timecard_rows(svc, ym)
    byStaffDate = {}
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        staff = str(r[1]) if len(r)>1 else ''
        date = parse_cell_date(r[3]) if len(r)>3 else None
        if date is None or not fmt_date(date).startswith(ym):
            continue
        typ  = str(r[2]) if len(r)>2 else ''
        time = str(r[4])[:5] if len(r)>4 else ''
        byStaffDate.setdefault(staff, {}).setdefault(fmt_date(date), []).append({'type':typ,'time':time})

    stats=[]
    for staff, dates in byStaffDate.items():
        totalWork=totalLate=days=0
        for date, ps in dates.items():
            w=tc_work_min(ps); l=tc_late_min(ps)
            if w>0: days+=1; totalWork+=w; totalLate+=l
        todayPs = dates.get(today, [])
        last = todayPs[-1] if todayPs else None
        state = 'absent' if not last else ('done' if last['type']=='退勤'
                else 'break' if last['type']=='休憩開始' else 'working')
        stats.append({'name':staff,'days':days,'totalWorkMin':totalWork,'totalLateMin':totalLate,
                      'normalMin':max(0,totalWork-totalLate),'todayState':state,
                      'todayIn':next((p['time'] for p in todayPs if p['type']=='出勤'),'-')})
    stats.sort(key=lambda s:s['totalWorkMin'], reverse=True)
    return {'ok': True, 'month': ym, 'today': today, 'staffStats': stats}

# ===== CUSTOMERS（アンケート）=====
def compute_customers(sheets, cmap12):
    monthly={}
    for date,c in cmap12.items():
        ym=date[:7]
        m=monthly.setdefault(ym,{'total':0,'newC':0,'repeat':0})
        m['total']+=c.get('total',0); m['newC']+=c.get('newC',0); m['repeat']+=c.get('repeat',0)
    trend=[{'label':ym,'total':monthly[ym]['total'],'newC':monthly[ym]['newC'],
            'repeat':monthly[ym]['repeat'],
            'newRate':rhu(monthly[ym]['newC']/monthly[ym]['total']*100) if monthly[ym]['total']>0 else 0}
           for ym in sorted(monthly)]

    referral={}; motive={}; exp={}; total=0
    rows = read_values(sheets, SURVEY_SS_ID, "'アンケート回答'!A1:D5000")
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        total+=1
        for v in (str(r[1]) if len(r)>1 else '').split('、'):
            if v: referral[v]=referral.get(v,0)+1
        for v in (str(r[2]) if len(r)>2 else '').split('、'):
            if v: motive[v]=motive.get(v,0)+1
        e=(str(r[3]) if len(r)>3 else '').strip()
        if e: exp[e]=exp.get(e,0)+1
    sort_obj=lambda o:[{'label':k,'count':v} for k,v in sorted(o.items(),key=lambda x:-x[1])]
    return {'ok':True,'monthlyTrend':trend,'referral':sort_obj(referral),
            'motive':sort_obj(motive),'experience':sort_obj(exp),'totalSurveys':total}

# ===== PREDICTION（Anthropic API）=====
def compute_prediction(daily_all):
    now=now_jst(); ym=now.strftime('%Y/%m')
    days_in_month=((now.replace(day=28)+dt.timedelta(days=4)).replace(day=1)-dt.timedelta(days=1)).day
    target=TARGETS.get(ym,{})
    monthly={}; mS=mC=mD=0
    for r in daily_all:
        m=r['date'][:7]
        mm=monthly.setdefault(m,{'sales':0,'customers':0,'days':0})
        mm['sales']+=r['sales']; mm['customers']+=r['customers']; mm['days']+=1
        if m==ym: mS+=r['sales']; mC+=r['customers']; mD+=1
    projected=rhu(mS/mD*days_in_month) if mD>0 else 0
    unit=rhu(mS/mC) if mC>0 else 0
    hist='\n'.join(
        f"  {m}: 売上¥{monthly[m]['sales']:,} / 来客{monthly[m]['customers']}人 / 目標¥{(TARGETS.get(m,{}) or {}).get('sales') or '-'}"
        for m in sorted(monthly) if m!=ym)[-3:] if monthly else ''
    hist_lines='\n'.join(
        f"  {m}: 売上¥{monthly[m]['sales']:,} / 来客{monthly[m]['customers']}人 / 目標¥{((TARGETS.get(m,{}) or {}).get('sales')) or '-'}"
        for m in sorted([k for k in monthly if k!=ym])[-3:])

    prompt=f"""あなたはポーカールーム「YORIMICHI」（静岡県沼津市）の経営アドバイザーです。
以下のデータをもとに、経営者向けの簡潔で実用的な分析と提案を行ってください。

【今月の状況】({ym}・{now.day}日時点 / 全{days_in_month}日)
- 累計売上: ¥{mS:,} / 目標¥{(target.get('sales') or 0):,} (達成率{rhu(mS/target['sales']*100) if target.get('sales') else '-'}%)
- 累計来客: {mC}人 / 目標{target.get('customers') or '-'}人
- 客単価: ¥{unit:,}
- 月末着地予測: ¥{projected:,}
- 損益分岐点: ¥1,771,671

【直近3ヶ月の実績】
{hist_lines}

以下のJSON形式のみで回答してください（前置き・説明文不要）:
{{
  "forecast": "月末着地予測の根拠と見通し（1〜2文）",
  "trend": "直近トレンドの分析（1〜2文）",
  "actions": ["今すぐ実行できる施策（具体的に）","今週中にできる施策","今月中にできる施策"],
  "risk": "現時点の最大リスク（1文）",
  "opportunity": "見逃せないチャンス（1文）"
}}"""

    api_key=os.environ.get('ANTHROPIC_API_KEY')
    ai={}
    if not api_key:
        ai={'forecast':'ANTHROPIC_API_KEY 未設定のためAI分析はスキップされました'}
    else:
        try:
            import urllib.request
            req=urllib.request.Request('https://api.anthropic.com/v1/messages',
                data=json.dumps({'model':'claude-haiku-4-5-20251001','max_tokens':1024,
                    'messages':[{'role':'user','content':prompt}]}).encode('utf-8'),
                headers={'x-api-key':api_key,'anthropic-version':'2023-06-01',
                         'content-type':'application/json'})
            with urllib.request.urlopen(req, timeout=60) as resp:
                raw=json.loads(resp.read().decode('utf-8'))
            text=raw['content'][0]['text']
            import re
            mm=re.search(r'\{[\s\S]*\}', text)
            ai=json.loads(mm.group(0)) if mm else {'forecast':text}
        except Exception as e:
            ai={'forecast':'AI分析エラー: '+str(e)}
    data={'generatedAt':now.strftime('%Y/%m/%d %H:%M'),'month':ym,
          'currentSales':mS,'projectedSales':projected,'targetSales':target.get('sales'),
          'projectionRate':rhu(projected/target['sales']*100) if target.get('sales') else None,
          'ai':ai}
    return {'ok':True,'cached':False,'data':data}

# ===== メイン =====
def main():
    sheets, drive = get_service()

    # 方針B：当月・翌月の未反映ブロックを検出し追記（APPEND_MONTHS=1 のときのみ実書込）。
    # データ生成の前段。失敗してもデータ生成は続行する。
    try:
        dry = os.environ.get('APPEND_MONTHS') != '1'
        for line in ensure_month_blocks(sheets, drive, dry_run=dry):
            print(f"[ensure_blocks]{'(dry)' if dry else ''} {line}")
    except Exception as e:
        print(f"[ensure_blocks] 例外のためスキップ: {str(e)[:100]}")

    # 収支記録・来客数を1回ずつ全読み
    sales_rows = read_values(sheets, MASTER_SS_ID, "'import_収支記録_2026'!A1:AZ400", 'UNFORMATTED_VALUE')
    cust_rows  = read_values(sheets, MASTER_SS_ID, "'import_来客数_2026'!A1:H400", 'UNFORMATTED_VALUE')

    daily24 = get_daily_rows(sales_rows, 24)
    cmap24  = get_customer_map(cust_rows, 24)
    daily3  = [r for r in daily24]  # summary は月フィルタするので全期間でも可
    cmap3   = cmap24
    cmap12  = cmap24
    # sales(daily) は GAS getSalesData_(period=daily, months=2) と同じ2ヶ月窓で読む
    daily2  = get_daily_rows(sales_rows, 2)
    cmap2   = get_customer_map(cust_rows, 2)

    out = {
        '_generatedAt': now_jst().strftime('%Y/%m/%d %H:%M'),
        'summary':        compute_summary(daily3, cmap3),
        'sales_monthly':  compute_sales_monthly(daily24),
        'sales_daily_2':  compute_sales_daily(daily2, cmap2),
        'history':        compute_history(daily24, cmap24),
        'breakdown_monthly_3':   compute_breakdown(sales_rows, 3),
        'breakdown_monthly_6':   compute_breakdown(sales_rows, 6),
        'breakdown_monthly_12':  compute_breakdown(sales_rows, 12),
        'breakdown_monthly_all': compute_breakdown(sales_rows, 999),
        'customers':      compute_customers(sheets, cmap12),
        'staff':          compute_staff(sheets),
        'yesterday':      compute_yesterday(daily24, cmap24, sheets),
        'prediction':     compute_prediction(daily24),
    }

    here = os.path.dirname(os.path.abspath(__file__))
    dst = os.environ.get('DATA_JSON_OUT', os.path.join(here, 'data.json'))
    with open(dst, 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
    print(f"[build_dashboard] wrote {dst}")
    s = out['summary']['sales']
    print(f"  当月売上 ¥{s['actual']:,} / BEP達成 {s['bepRate']}% / 月数 {out['history']['summary']['totalMonths']}")

if __name__ == '__main__':
    main()
