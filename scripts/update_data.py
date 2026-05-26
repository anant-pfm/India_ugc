#!/usr/bin/env python3
"""
UGC Tracker — regenerates index.html with fresh data from data.xlsx
Usage: python scripts/update_data.py path/to/data.xlsx
"""
import sys, os, json, math, re
import pandas as pd
from datetime import datetime, date, timedelta

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def clean(v):
    if v is None: return ''
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return ''
    s = str(v).strip()
    if s.endswith('.0') and s[:-2].lstrip('-').isdigit(): s = s[:-2]
    return '' if s in ('nan','None','NaT') else s

def num(v):
    if v is None: return 0
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return 0
    try: f=float(v); return int(f) if f==int(f) else round(f,4)
    except: return 0

def tgt(v):
    if v is None: return ''
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)): return ''
    s=str(v).strip()
    if s in ('nan','None','NaT',''): return ''
    if s=='??': return '??'
    try: f=float(v); return int(f) if f==int(f) else round(f,4)
    except: return s

def excel_date(v):
    if v is None: return ''
    if isinstance(v,(datetime,pd.Timestamp)):
        try: return v.strftime('%Y-%m-%d')
        except: return ''
    if isinstance(v,float) and not (math.isnan(v) or math.isinf(v)):
        try: return (date(1899,12,30)+timedelta(days=int(v))).strftime('%Y-%m-%d')
        except: return ''
    return ''

def to_ds(v):
    if pd.isna(v) if not isinstance(v,str) else not v: return ''
    if isinstance(v,(datetime,pd.Timestamp)):
        try: return v.strftime('%Y-%m-%d')
        except: return ''
    return str(v)[:10]

WD  = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
MN  = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def day_name(ds):
    try: return WD[datetime.strptime(ds,'%Y-%m-%d').weekday()]
    except: return ''
def mo_lbl(ds):
    try: d=datetime.strptime(ds,'%Y-%m-%d'); return f"{MN[d.month-1]} {d.year}"
    except: return ''
def wk_lbl(ds):
    try:
        d=datetime.strptime(ds,'%Y-%m-%d'); fd=d.replace(day=1); off=fd.weekday()
        wn=(d.day+off-1)//7+1; sy=str(d.year)[2:]
        return f"W{wn} - {MN[d.month-1]}'{sy}"
    except: return ''

def process(xl):
    df=pd.read_excel(xl,sheet_name='Monthwise',header=None)
    data=df.iloc[4:].reset_index(drop=True)
    weekly=[]
    for _,r in data.iterrows():
        wk=clean(r.iloc[0])
        if not wk or wk=='nan': continue
        show=clean(r.iloc[3])
        if not show: continue
        weekly.append([clean(r.iloc[2]),wk,show,clean(r.iloc[4]),clean(r.iloc[5]),
            clean(r.iloc[6]),clean(r.iloc[7]),num(r.iloc[8]),num(r.iloc[9]),
            clean(r.iloc[10]) or '0',clean(r.iloc[11]) or '0',clean(r.iloc[12]) or '0',
            tgt(r.iloc[13]),num(r.iloc[14]),num(r.iloc[15]),num(r.iloc[16]),
            tgt(r.iloc[18]),num(r.iloc[19]),num(r.iloc[20]),
            tgt(r.iloc[22]),num(r.iloc[23]),
            excel_date(r.iloc[27]) if len(r)>27 else ''])

    df2=pd.read_excel(xl,sheet_name='Weekly Dump',header=0)
    df2=df2.dropna(how='all').reset_index(drop=True)
    for c in ['Under Review (scripts)','Approved (scripts)','Released (eps)','Under Review (word count)','Released (hr)']:
        if c in df2.columns: df2[c]=pd.to_numeric(df2[c],errors='coerce').fillna(0)
    yearly=[]
    for _,r in df2.iterrows():
        ds=to_ds(r.get('Period',''))
        if not ds: continue
        ttl=clean(r.get('Show Title',''))
        if not ttl: continue
        yearly.append([clean(r.get('Show ID','')),ttl,mo_lbl(ds),ds[:7],
            num(r.get('Under Review (scripts)',0)),num(r.get('Approved (scripts)',0)),
            num(r.get('Released (eps)',0)),num(r.get('Under Review (word count)',0)),
            round(float(r.get('Released (hr)',0) or 0),4),wk_lbl(ds),ds])

    df3=pd.read_excel(xl,sheet_name='Daily Dump',header=0)
    df3=df3.dropna(how='all').reset_index(drop=True)
    for c in ['Under Review (scripts)','Approved (scripts)','Released (eps)','Under Review (word count)','Released (hr)']:
        if c in df3.columns: df3[c]=pd.to_numeric(df3[c],errors='coerce').fillna(0)
    daily=[]
    for _,r in df3.iterrows():
        ds=to_ds(r.get('Period',''))
        if not ds: continue
        ttl=clean(r.get('Show Title',''))
        if not ttl: continue
        daily.append([ds,day_name(ds),clean(r.get('Show ID','')),ttl,
            num(r.get('Under Review (scripts)',0)),num(r.get('Approved (scripts)',0)),
            num(r.get('Released (eps)',0)),num(r.get('Under Review (word count)',0)),
            round(float(r.get('Released (hr)',0) or 0),4)])
    return weekly, yearly, daily

def main():
    if len(sys.argv) < 2:
        print("Usage: python scripts/update_data.py path/to/data.xlsx")
        sys.exit(1)
    xl = sys.argv[1]
    if not os.path.exists(xl):
        print(f"File not found: {xl}"); sys.exit(1)

    print(f"Reading {xl} ...")
    weekly, yearly, daily = process(xl)
    print(f"  WDATA: {len(weekly)} rows")
    print(f"  YDATA: {len(yearly)} rows")
    print(f"  DDATA: {len(daily)} rows")

    today = datetime.today().strftime('%d %b %Y %H:%M')

    new_data = (
        f"// Auto-generated {today} from {os.path.basename(xl)}\n"
        f"// WDATA:{len(weekly)} rows  YDATA:{len(yearly)} rows  DDATA:{len(daily)} rows\n"
        f"window.WDATA={json.dumps(weekly,ensure_ascii=False,separators=(',',':'))};\n"
        f"window.YDATA={json.dumps(yearly,ensure_ascii=False,separators=(',',':'))};\n"
        f"window.DDATA={json.dumps(daily,ensure_ascii=False,separators=(',',':'))};"
    )

    index_path = os.path.join(ROOT, 'index.html')
    with open(index_path,'r',encoding='utf-8') as f:
        html = f.read()

    # Replace the data block: from "// Auto-generated" line through end of window.DDATA=...;
    pattern = r'// Auto-generated.*?window\.DDATA=\[.*?\];'
    if not re.search(pattern, html, re.DOTALL):
        print("ERROR: Could not find data block in index.html")
        print("Make sure index.html contains the '// Auto-generated' comment")
        sys.exit(1)

    new_html = re.sub(pattern, new_data, html, count=1, flags=re.DOTALL)

    with open(index_path,'w',encoding='utf-8') as f:
        f.write(new_html)

    print(f"✓ index.html updated ({len(new_html):,} chars)")
    print("Done.")

if __name__ == '__main__':
    main()
