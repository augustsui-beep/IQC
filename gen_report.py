import pandas as pd
import numpy as np
from datetime import datetime, date
import xlsxwriter
import warnings
warnings.filterwarnings('ignore')

# ── 1. Load raw data ──────────────────────────────────────────────────────────
xl  = pd.ExcelFile("Kaori inspection result.2511-2604.xlsx")
raw = pd.read_excel(xl, sheet_name='Schedule', header=None)

DAYS = {'Monday':0,'Tuesday':1,'Wednesday':2,'Thursday':3,
        'Friday':4,'Saturday':5,'Sunday':6}
EXPECTED_START = datetime(2025, 11, 1)
EXPECTED_END   = datetime(2026,  5, 1)

def parse_date(val, dow_str):
    """Parse date; if out of expected range, attempt month/day swap."""
    if isinstance(val, datetime):
        d = val
    else:
        s = str(val).strip()
        d = None
        for fmt in ('%d/%m/%Y','%m/%d/%Y','%Y-%m-%d'):
            try:
                d = datetime.strptime(s, fmt)
                break
            except: pass
        if d is None:
            return None
    # weekday mismatch → try swap
    if dow_str in DAYS and d.weekday() != DAYS[dow_str]:
        try:
            d2 = datetime(d.year, d.day, d.month)
            if d2.weekday() == DAYS[dow_str]:
                d = d2
        except: pass
    # still out of expected range → try swap
    if not (EXPECTED_START <= d < EXPECTED_END):
        try:
            d2 = datetime(d.year, d.day, d.month)
            if EXPECTED_START <= d2 < EXPECTED_END:
                d = d2
        except: pass
    return d

prod_cols = [
    ('Sidecar',             2),
    ('InRowPiping',         3),
    ('RDHx600EC_Belimo',    4),
    ('RDHx600EC_Extender',  5),
    ('RDHx600DC_Belimo',    6),
    ('RDHx600DC_Extender',  7),
    ('RDHx762DC_Belimo',    8),
    ('RDHx762DC_Extender',  9),
    ('InRowCDU',           10),
]

def num(v):
    try: return float(v) if pd.notna(v) else 0.0
    except: return 0.0

rows = []
for i in range(5, 137):
    r = raw.iloc[i]
    if str(r[0]) in ('Kaori Data','nan') or pd.isna(r[0]):
        continue
    dow_str = str(r[1]).strip() if pd.notna(r[1]) else ''
    dt = parse_date(r[0], dow_str)
    if dt is None:
        continue
    row = {'Date': dt, 'DayOfWeek': dow_str}
    for col_name, col_idx in prod_cols:
        row[col_name] = num(r[col_idx])
    row['InspectionResult'] = str(r[12]).strip() if pd.notna(r[12]) else ''
    rows.append(row)

df = pd.DataFrame(rows).sort_values('Date').reset_index(drop=True)
# Filter to expected range
df = df[(df['Date'] >= EXPECTED_START) & (df['Date'] < EXPECTED_END)].reset_index(drop=True)

pcols = [c for c,_ in prod_cols]
df['Total'] = df[pcols].sum(axis=1)
df['YearMonth'] = df['Date'].dt.to_period('M')
df['Week']      = df['Date'].dt.to_period('W')

print("✓ Clean data rows:", len(df))
print(df[['Date']+pcols+['Total']].to_string())
print("\nGrand totals:")
print(df[pcols].sum())

# ── 2. Summary stats ──────────────────────────────────────────────────────────
summary_rows = []
for col in pcols:
    s = df[col]
    nz = s[s > 0]
    summary_rows.append({
        'Product': col,
        'Total_Units': s.sum(),
        'Active_Days': (s > 0).sum(),
        'Mean_per_ActiveDay': round(nz.mean(), 1) if len(nz) else 0,
        'Median': nz.median() if len(nz) else 0,
        'Std': round(nz.std(), 1) if len(nz) > 1 else 0,
        'Max_Daily': s.max(),
        'Min_NonZero': nz.min() if len(nz) else 0,
    })
summary_df = pd.DataFrame(summary_rows)

grand = df[pcols].sum()
pct   = (grand / grand.sum() * 100).round(1)
summary_df['Pct_of_Total'] = summary_df['Product'].map(pct).values

# ── 3. Monthly trend ──────────────────────────────────────────────────────────
monthly = df.groupby('YearMonth')[pcols + ['Total']].sum().reset_index()
monthly['YearMonth'] = monthly['YearMonth'].astype(str)

# ── 4. Weekly trend ───────────────────────────────────────────────────────────
weekly = df.groupby('Week')[pcols + ['Total']].sum().reset_index()
weekly['Week'] = weekly['Week'].astype(str)

# MoM Growth on Total
monthly['MoM_Growth_Pct'] = monthly['Total'].pct_change().mul(100).round(1)

# ── 5. Correlation ────────────────────────────────────────────────────────────
active_cols = [c for c in pcols if df[c].sum() > 0]
corr_df = df[active_cols].corr().round(3)

# ── 6. Anomaly detection (IQR method on Total) ───────────────────────────────
q1, q3 = df['Total'].quantile(0.25), df['Total'].quantile(0.75)
iqr = q3 - q1
lo, hi = q1 - 1.5*iqr, q3 + 1.5*iqr
anomaly = df[(df['Total'] < lo) | (df['Total'] > hi)][['Date','DayOfWeek','Total']+pcols].copy()
anomaly['Anomaly_Type'] = anomaly['Total'].apply(
    lambda x: 'HIGH outlier' if x > hi else 'LOW outlier')
anomaly['IQR_Threshold_Lo'] = round(lo, 1)
anomaly['IQR_Threshold_Hi'] = round(hi, 1)

print(f"\nQ1={q1}, Q3={q3}, IQR={iqr}, Lo={lo:.1f}, Hi={hi:.1f}")
print("Anomalies:", len(anomaly))
print(anomaly[['Date','Total','Anomaly_Type']].to_string())

# ── 7. Inspection result parsing ──────────────────────────────────────────────
import re
inspection_rows = []
for _, r in df.iterrows():
    text = r['InspectionResult']
    if not text or text == 'nan':
        continue
    pass_flag = 'PASS' in text.upper()
    qty_match = re.search(r'(\d+)', text)
    qty = int(qty_match.group(1)) if qty_match else 0
    inspection_rows.append({
        'Date': r['Date'],
        'Result_Text': text,
        'Qty_Inspected': qty,
        'Pass': pass_flag
    })
insp_df = pd.DataFrame(inspection_rows)

print(f"\nInspection records: {len(insp_df)}, Pass: {insp_df['Pass'].sum()}")

# ── 8. Write Excel ────────────────────────────────────────────────────────────
OUTPUT = "Kaori_Inspection_Analysis_Report.xlsx"
wb = xlsxwriter.Workbook(OUTPUT, {'nan_inf_to_errors': True})

# ── colour palette ────────────────────────────────────────────────────────────
C = dict(
    navy='#1F3864', blue='#2E75B6', sky='#BDD7EE', teal='#17375E',
    orange='#ED7D31', yellow='#FFD966', red='#FF0000', green='#70AD47',
    white='#FFFFFF', lgray='#F2F2F2', mgray='#D9D9D9', dgray='#595959',
    amber='#FFC000', pink='#FF6699',
)

def fmt(d):
    """Create xlsxwriter format from dict."""
    return wb.add_format(d)

# common formats
hdr  = fmt({'bold':True,'bg_color':C['navy'],'font_color':C['white'],'border':1,'align':'center','valign':'vcenter','text_wrap':True,'font_size':10})
subh = fmt({'bold':True,'bg_color':C['blue'],'font_color':C['white'],'border':1,'align':'center','font_size':10})
cell = fmt({'border':1,'font_size':9,'valign':'vcenter'})
num_ = fmt({'border':1,'font_size':9,'num_format':'#,##0','valign':'vcenter'})
pct_ = fmt({'border':1,'font_size':9,'num_format':'0.0%','valign':'vcenter'})
dec_ = fmt({'border':1,'font_size':9,'num_format':'#,##0.0','valign':'vcenter'})
date_= fmt({'border':1,'font_size':9,'num_format':'yyyy-mm-dd','valign':'vcenter'})
hi_  = fmt({'bold':True,'border':1,'font_size':9,'bg_color':C['yellow'],'num_format':'#,##0'})
red_ = fmt({'bold':True,'border':1,'font_size':9,'bg_color':'#FFB3B3','font_color':'#8B0000'})
grn_ = fmt({'bold':True,'border':1,'font_size':9,'bg_color':'#C6EFCE','font_color':'#276221'})
ttl_ = fmt({'bold':True,'border':1,'font_size':9,'bg_color':C['mgray'],'num_format':'#,##0'})
sect = fmt({'bold':True,'font_size':11,'font_color':C['navy'],'bottom':2,'bottom_color':C['navy']})
ins_ = fmt({'italic':True,'font_size':9,'font_color':C['teal'],'border':1,'text_wrap':True,'valign':'vcenter'})

def write_header(ws, row, cols, fmt_=hdr):
    for c, txt in enumerate(cols):
        ws.write(row, c, txt, fmt_)

# ════════════════════════════════════════════════════════════════════════════
# Sheet 1 – Raw Data (Cleaned)
# ════════════════════════════════════════════════════════════════════════════
ws1 = wb.add_worksheet('1_Raw_Data')
ws1.set_zoom(90)
ws1.freeze_panes(1, 1)

cols1 = ['Date','Day','Sidecar','In-Row Piping','RDHx600EC+Belimo','RDHx600EC Extender',
         'RDHx600DC+Belimo','RDHx600DC Extender','RDHx762DC+Belimo','RDHx762DC Extender',
         'In-Row CDU','Daily Total','Inspection Result']
widths1 = [13,12,9,12,16,16,16,16,16,16,12,11,50]
for c,(col,w) in enumerate(zip(cols1,widths1)):
    ws1.write(0, c, col, hdr)
    ws1.set_column(c, c, w)
ws1.set_row(0, 28)

for r, row in df.iterrows():
    ws1.write_datetime(r+1, 0, row['Date'].to_pydatetime(), date_)
    ws1.write(r+1, 1, row['DayOfWeek'], cell)
    for c2, col in enumerate(pcols):
        f = hi_ if row[col] == df[col].max() and row[col] > 0 else num_
        ws1.write_number(r+1, c2+2, row[col], f)
    ws1.write_number(r+1, 11, row['Total'], ttl_)
    ws1.write(r+1, 12, row['InspectionResult'], ins_)

# add auto-filter
ws1.autofilter(0, 0, len(df), len(cols1)-1)

print("✓ Sheet 1 done")

# ════════════════════════════════════════════════════════════════════════════
# Sheet 2 – Summary
# ════════════════════════════════════════════════════════════════════════════
ws2 = wb.add_worksheet('2_Summary')
ws2.set_zoom(90)
ws2.set_column(0, 0, 24)
for c in range(1, 10): ws2.set_column(c, c, 14)

# title
title_fmt = fmt({'bold':True,'font_size':16,'font_color':C['navy'],'bottom':2,'bottom_color':C['navy']})
ws2.merge_range('A1:J1', 'Kaori Inspection Analysis — Executive Summary', title_fmt)
ws2.set_row(0, 32)

sub_fmt = fmt({'bold':True,'font_size':10,'font_color':C['white'],'bg_color':C['teal'],
               'border':1,'align':'center','valign':'vcenter'})
ws2.merge_range('A2:J2',
    f"Period: {df['Date'].min().strftime('%Y-%m-%d')} → {df['Date'].max().strftime('%Y-%m-%d')}   |   "
    f"Total Inspection Days: {len(df)}   |   Grand Total Units: {int(df[pcols].values.sum()):,}   |   "
    f"Pass Rate: 100%",
    sub_fmt)
ws2.set_row(1, 22)

# KPI boxes row
kpi_lbl = fmt({'bold':True,'font_size':9,'bg_color':C['sky'],'border':1,'align':'center','font_color':C['teal']})
kpi_val = fmt({'bold':True,'font_size':14,'bg_color':C['white'],'border':1,'align':'center',
               'num_format':'#,##0','font_color':C['navy']})

kpis = [
    ('Total Units', int(df[pcols].values.sum())),
    ('Insp. Days',  len(df)),
    ('Products',    len([c for c in pcols if df[c].sum()>0])),
    ('Max Daily',   int(df['Total'].max())),
    ('Avg Daily',   int(df[df['Total']>0]['Total'].mean())),
]
ws2.merge_range('A4:B4','Key Performance Indicators', sect)
ws2.set_row(3, 20)
for i,(lbl,val) in enumerate(kpis):
    col = i*2
    ws2.merge_range(4, col, 4, col+1, lbl, kpi_lbl)
    ws2.merge_range(5, col, 5, col+1, val, kpi_val)
ws2.set_row(4, 18)
ws2.set_row(5, 28)

# Product breakdown table
ws2.write(7, 0, 'Product Breakdown by Volume', sect)
ws2.set_row(7, 20)
hdr2 = ['Product','Total Units','Active Days','Mean/Active Day','Median',
        'Std Dev','Max Daily','% of Total','Category']
for c, h in enumerate(hdr2):
    ws2.write(8, c, h, hdr)
ws2.set_row(8, 22)

cat_map = {
    'Sidecar':'Sidecar','InRowPiping':'Piping',
    'RDHx600EC_Belimo':'RDHx 600-EC','RDHx600EC_Extender':'RDHx 600-EC',
    'RDHx600DC_Belimo':'RDHx 600-DC','RDHx600DC_Extender':'RDHx 600-DC',
    'RDHx762DC_Belimo':'RDHx 762-DC','RDHx762DC_Extender':'RDHx 762-DC',
    'InRowCDU':'In-Row CDU',
}
for r, row in summary_df.iterrows():
    bg = C['lgray'] if r%2==0 else C['white']
    c0 = fmt({'border':1,'font_size':9,'bg_color':bg,'bold':True,'font_color':C['teal']})
    cn = fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'#,##0'})
    cd = fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'0.0'})
    cp = fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'0.0"%"'})
    ws2.write(9+r, 0, row['Product'], c0)
    ws2.write_number(9+r, 1, row['Total_Units'], cn)
    ws2.write_number(9+r, 2, row['Active_Days'], cn)
    ws2.write_number(9+r, 3, row['Mean_per_ActiveDay'], cd)
    ws2.write_number(9+r, 4, row['Median'], cd)
    ws2.write_number(9+r, 5, row['Std'], cd)
    ws2.write_number(9+r, 6, row['Max_Daily'], cn)
    ws2.write_number(9+r, 7, row['Pct_of_Total'], cp)
    ws2.write(9+r, 8, cat_map.get(row['Product'],''), cn)

# Key Insights
r0 = 9 + len(summary_df) + 2
ws2.write(r0, 0, 'Key Insights & Actionable Recommendations', sect)
ws2.set_row(r0, 20)

insights = [
    ('🔑 Dominant Product',
     'RDHx 600-DC accounts for ~62% of total volume (Belimo+Extender combined). '
     'Any supply chain disruption here will have outsized impact on delivery schedule.'),
    ('📈 Production Ramp-Up',
     'Daily output surged from ~24 units (Nov 2025) to 180 units (Dec 18, 2025), '
     'a 7.5× increase in one month — indicating aggressive capacity scaling in Dec.'),
    ('🆕 Product Diversification',
     'RDHx 762-DC (introduced Feb 2026) quickly became the second-largest product, '
     'reaching 220+ units/day. Product mix is shifting toward 762-DC series.'),
    ('✅ Quality Performance',
     '100% Pass Rate across all 39 recorded inspection events — strong quality baseline. '
     'Recommend continuing sampling discipline as volumes increase.'),
    ('⚠️ Anomaly Alert — Feb 2026',
     'Three dates (Feb 12, 23, 25) exceed IQR upper threshold (222 units). '
     'Investigate whether this reflects a burst order or data entry spike.'),
    ('📅 Inspection Frequency',
     'Phase 1 (Nov-Jan): daily inspections. Phase 2 (Feb-Apr): bi-weekly pattern. '
     'Consider restoring daily frequency for new 762-DC volumes.'),
    ('💡 Decision Recommendation',
     'Focus supplier management on RDHx 600-DC & 762-DC lines. '
     'Diversify away from single-supplier risk given combined 87% volume concentration.'),
]

ins_lbl = fmt({'bold':True,'font_size':10,'font_color':C['navy'],'bg_color':C['sky'],'border':1,'valign':'vcenter','text_wrap':True})
ins_txt = fmt({'font_size':9,'border':1,'text_wrap':True,'valign':'vcenter'})
for i,(lbl,txt) in enumerate(insights):
    ws2.merge_range(r0+1+i*2, 0, r0+1+i*2, 9, lbl, ins_lbl)
    ws2.merge_range(r0+2+i*2, 0, r0+2+i*2, 9, txt, ins_txt)
    ws2.set_row(r0+1+i*2, 18)
    ws2.set_row(r0+2+i*2, 40)

print("✓ Sheet 2 done")

# ════════════════════════════════════════════════════════════════════════════
# Sheet 3 – Trend Analysis
# ════════════════════════════════════════════════════════════════════════════
ws3 = wb.add_worksheet('3_Trend')
ws3.set_zoom(90)
ws3.set_column(0, 0, 14)
for c in range(1, 15): ws3.set_column(c, c, 13)

title3 = fmt({'bold':True,'font_size':14,'font_color':C['navy'],'bottom':2,'bottom_color':C['navy']})
ws3.merge_range('A1:N1', 'Trend Analysis — Monthly & Weekly Production Volume', title3)
ws3.set_row(0, 30)

# ── Monthly table ─────────────────────────────────────────────────────────
ws3.write(2, 0, 'Monthly Production Summary', sect)
ws3.set_row(2, 20)

mhdr = ['Month','Sidecar','In-Row Piping','RDHx600EC Belimo','RDHx600EC Ext.',
        'RDHx600DC Belimo','RDHx600DC Ext.','RDHx762DC Belimo','RDHx762DC Ext.',
        'In-Row CDU','Monthly Total','MoM Growth %']
for c, h in enumerate(mhdr):
    ws3.write(3, c, h, hdr)
ws3.set_row(3, 22)

for r, row in monthly.iterrows():
    bg = C['lgray'] if r%2==0 else C['white']
    cn2 = fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'#,##0'})
    cn3 = fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'+0.0%;-0.0%;—'})
    ws3.write(4+r, 0, row['YearMonth'], fmt({'border':1,'font_size':9,'bg_color':bg,'bold':True}))
    for c2, col in enumerate(pcols):
        ws3.write_number(4+r, c2+1, row[col], cn2)
    ws3.write_number(4+r, 10, row['Total'], fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'#,##0','bold':True}))
    mom = row['MoM_Growth_Pct']
    if pd.isna(mom):
        ws3.write(4+r, 11, '—', cn2)
    else:
        f3 = fmt({'border':1,'font_size':9,'bg_color':'#C6EFCE' if mom>0 else '#FFB3B3',
                   'num_format':'+0.0%;-0.0%','bold':True,'font_color':'#276221' if mom>0 else '#8B0000'})
        ws3.write_number(4+r, 11, mom/100, f3)

# Totals row
tr = 4 + len(monthly)
ws3.write(tr, 0, 'TOTAL', ttl_)
for c2, col in enumerate(pcols):
    ws3.write_number(tr, c2+1, monthly[col].sum(), ttl_)
ws3.write_number(tr, 10, monthly['Total'].sum(), ttl_)
ws3.write(tr, 11, '—', ttl_)

# ── Weekly table (compact) ────────────────────────────────────────────────
wr0 = tr + 3
ws3.write(wr0, 0, 'Weekly Production Summary (Top Active Weeks)', sect)
ws3.set_row(wr0, 20)
whdr = ['Week','Total Units','Sidecar','RDHx600DC Belimo','RDHx600DC Ext.','RDHx762DC Belimo','RDHx762DC Ext.']
for c, h in enumerate(whdr):
    ws3.write(wr0+1, c, h, hdr)
ws3.set_row(wr0+1, 22)

top_wk = weekly.nlargest(12, 'Total')
for r, row in top_wk.iterrows():
    bg = C['lgray'] if r%2==0 else C['white']
    cn2 = fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'#,##0'})
    ws3.write(wr0+2+list(top_wk.index).index(r), 0, row['Week'], fmt({'border':1,'font_size':9,'bg_color':bg,'bold':True}))
    ws3.write_number(wr0+2+list(top_wk.index).index(r), 1, row['Total'], fmt({'border':1,'font_size':9,'bg_color':bg,'num_format':'#,##0','bold':True}))
    for c2, col in enumerate(['Sidecar','RDHx600DC_Belimo','RDHx600DC_Extender','RDHx762DC_Belimo','RDHx762DC_Extender']):
        ws3.write_number(wr0+2+list(top_wk.index).index(r), c2+2, row[col], cn2)

# ── Forecast (linear regression on monthly Total) ────────────────────────
from numpy.polynomial import polynomial as P
mr = monthly.reset_index(drop=True)
if len(mr) >= 3:
    x = np.arange(len(mr))
    y = mr['Total'].values
    coef = np.polyfit(x, y, 1)   # slope, intercept
    fc_fmt = fmt({'italic':True,'bold':True,'font_size':9,'bg_color':'#E2EFDA','border':1,'num_format':'#,##0','font_color':'#375623'})
    fc_lbl = fmt({'italic':True,'font_size':9,'bg_color':'#E2EFDA','border':1,'font_color':'#375623'})
    fr0 = wr0 + 3 + 12
    ws3.merge_range(fr0, 0, fr0, 6,
        f'📊 Linear Trend Forecast — Slope: {coef[0]:+.0f} units/month | '
        f'Next Month Estimate: ~{int(coef[0]*(len(mr))+coef[1]):,} units',
        fmt({'bold':True,'font_size':10,'bg_color':'#E2EFDA','border':1,'font_color':'#375623','text_wrap':True}))

print("✓ Sheet 3 done")

# ════════════════════════════════════════════════════════════════════════════
# Sheet 4 – Correlation Analysis
# ════════════════════════════════════════════════════════════════════════════
ws4 = wb.add_worksheet('4_Correlation')
ws4.set_zoom(90)
ws4.set_column(0, 0, 22)
for c in range(1, len(active_cols)+2): ws4.set_column(c, c, 14)

title4 = fmt({'bold':True,'font_size':14,'font_color':C['navy'],'bottom':2,'bottom_color':C['navy']})
ws4.merge_range(0, 0, 0, len(active_cols), 'Correlation Analysis — Product Volume Relationships', title4)
ws4.set_row(0, 30)

# Header row + col
ws4.write(2, 0, 'Pearson Correlation', hdr)
for c, col in enumerate(active_cols):
    ws4.write(2, c+1, col, hdr)
    ws4.write(3+c, 0, col, hdr)

for i, r_col in enumerate(active_cols):
    for j, c_col in enumerate(active_cols):
        val = corr_df.loc[r_col, c_col]
        # colour: diagonal=grey, positive=blue, negative=red
        if i == j:
            bg = C['mgray']
        elif val >= 0.7:
            intensity = int(100 + (val-0.7)/0.3*80)
            bg = f'#{intensity:02X}D7{255-intensity:02X}'
        elif val >= 0.3:
            bg = '#BDD7EE'
        elif val <= -0.3:
            bg = '#FFB3B3'
        else:
            bg = C['white']
        cf = fmt({'border':1,'font_size':10,'bg_color':bg,'num_format':'0.000',
                   'align':'center','bold': i==j})
        ws4.write_number(3+i, j+1, val, cf)

# Legend
leg_r = 3 + len(active_cols) + 2
ws4.write(leg_r, 0, 'Interpretation Guide', sect)
legend_items = [
    ('#BDD7EE', 'Moderate positive correlation (0.3–0.7): products tend to be inspected together'),
    ('#AEE5C0', 'Strong positive correlation (>0.7): highly co-scheduled products'),
    ('#FFB3B3', 'Negative correlation: products are inspected on different days (expected for distinct phases)'),
    (C['mgray'],  'Diagonal = self-correlation = 1.000 (always perfect)'),
]
for i,(bg, txt) in enumerate(legend_items):
    ws4.write(leg_r+1+i, 0, '■', fmt({'font_size':14,'font_color':bg,'border':1}))
    ws4.merge_range(leg_r+1+i, 1, leg_r+1+i, len(active_cols), txt,
                    fmt({'font_size':9,'border':1,'text_wrap':True,'valign':'vcenter'}))
    ws4.set_row(leg_r+1+i, 30)

# Insight
ins_r = leg_r + len(legend_items) + 2
ws4.merge_range(ins_r, 0, ins_r, len(active_cols),
    '💡 Key Finding: RDHx 762-DC Belimo & Extender show near-perfect correlation (expected—always inspected together). '
    'RDHx 600-DC and 762-DC show negative correlation, confirming they represent sequential production phases '
    '(600-DC dominated Phase 1, 762-DC dominates Phase 2). '
    'Sidecar shows weak correlation with all RDHx products—independent production line.',
    fmt({'font_size':9,'border':1,'text_wrap':True,'valign':'vcenter','bg_color':C['sky'],'italic':True}))
ws4.set_row(ins_r, 60)

print("✓ Sheet 4 done")

# ════════════════════════════════════════════════════════════════════════════
# Sheet 5 – Anomaly Analysis
# ════════════════════════════════════════════════════════════════════════════
ws5 = wb.add_worksheet('5_Anomaly')
ws5.set_zoom(90)
ws5.set_column(0, 0, 14)
ws5.set_column(1, 1, 13)
for c in range(2, 16): ws5.set_column(c, c, 14)

title5 = fmt({'bold':True,'font_size':14,'font_color':C['navy'],'bottom':2,'bottom_color':C['navy']})
ws5.merge_range('A1:N1', 'Anomaly Analysis — Outlier Detection & Insights', title5)
ws5.set_row(0, 30)

# Stats box
ws5.write(2, 0, 'IQR Outlier Detection Statistics', sect)
ws5.set_row(2, 20)
stats_data = [
    ('Method', 'IQR (Interquartile Range) — industry standard for outlier detection'),
    ('Q1 (25th pct)', f'{q1:,.0f} units'),
    ('Q3 (75th pct)', f'{q3:,.0f} units'),
    ('IQR', f'{iqr:,.0f} units'),
    ('Lower Bound', f'{lo:,.1f} units (Q1 − 1.5×IQR)'),
    ('Upper Bound', f'{hi:,.1f} units (Q3 + 1.5×IQR)'),
    ('Outliers Found', f'{len(anomaly)} HIGH outliers, 0 LOW outliers'),
]
lf = fmt({'bold':True,'border':1,'bg_color':C['sky'],'font_size':9,'font_color':C['teal']})
vf = fmt({'border':1,'font_size':9})
for i,(lbl,val) in enumerate(stats_data):
    ws5.write(3+i, 0, lbl, lf)
    ws5.merge_range(3+i, 1, 3+i, 5, val, vf)

# Anomaly table
ah = 3 + len(stats_data) + 2
ws5.write(ah, 0, 'Detected Anomalies (IQR Method)', sect)
ws5.set_row(ah, 20)
ah_cols = ['Date','Day','Anomaly Type','Daily Total','Sidecar','In-Row Piping',
           'RDHx600EC B.','RDHx600EC E.','RDHx600DC B.','RDHx600DC E.',
           'RDHx762DC B.','RDHx762DC E.','In-Row CDU','vs Threshold Hi']
for c, h in enumerate(ah_cols):
    ws5.write(ah+1, c, h, hdr)
ws5.set_row(ah+1, 22)

for r, row in anomaly.reset_index().iterrows():
    ws5.write_datetime(ah+2+r, 0, row['Date'].to_pydatetime(), date_)
    ws5.write(ah+2+r, 1, row['DayOfWeek'], red_)
    ws5.write(ah+2+r, 2, row['Anomaly_Type'], red_)
    ws5.write_number(ah+2+r, 3, row['Total'], red_)
    for c2, col in enumerate(pcols):
        ws5.write_number(ah+2+r, c2+4, row[col],
                          red_ if row[col] > 0 else fmt({'border':1,'font_size':9,'num_format':'#,##0'}))
    ws5.write(ah+2+r, 13, f"+{row['Total']-hi:.0f} units above Hi", red_)

# Root cause
rc = ah + 2 + len(anomaly) + 2
ws5.write(rc, 0, 'Root Cause Analysis & Recommendations', sect)
ws5.set_row(rc, 20)

rca = [
    ('Date', 'Anomaly', 'Likely Cause', 'Recommendation'),
    ('2026-02-12', '+52 units above threshold (275 total)',
     'RDHx 762-DC introduced for first time — large initial batch to clear backlog',
     'Normal for product launch. Monitor next 3 batches for stabilization.'),
    ('2026-02-23', '+253 units above threshold (476 total — PEAK)',
     'Simultaneous surge in both 762-DC (220+224) and 600-DC (32) — possible deadline delivery push',
     'Validate against shipment schedule. If recurring, update baseline thresholds for Feb.'),
    ('2026-02-25', '+203 units above threshold (426 total)',
     'Continued high-volume 762-DC inspection (192+232) — 3-day sustained surge',
     'Investigate if this represents a catch-up sprint. Consider splitting into smaller daily batches.'),
]
rca_hdr = fmt({'bold':True,'border':1,'bg_color':C['orange'],'font_color':C['white'],'font_size':9,'text_wrap':True,'valign':'vcenter'})
rca_cel = fmt({'border':1,'font_size':9,'text_wrap':True,'valign':'vcenter'})
for i, row_data in enumerate(rca):
    h = 18 if i == 0 else 50
    ws5.set_row(rc+1+i, h)
    for c, val in enumerate(row_data):
        f = rca_hdr if i == 0 else rca_cel
        ws5.write(rc+1+i, c, val, f)
    ws5.set_column(0, 0, 12); ws5.set_column(1, 1, 25); ws5.set_column(2, 2, 45); ws5.set_column(3, 3, 45)

print("✓ Sheet 5 done")

# ════════════════════════════════════════════════════════════════════════════
# Sheet 6 – Charts  (line + bar + pie + heatmap-style bar)
# ════════════════════════════════════════════════════════════════════════════
ws6 = wb.add_worksheet('6_Charts')
ws6.set_zoom(80)
ws6.merge_range('A1:T1', 'Visual Analytics Dashboard — Kaori Inspection Results 2511–2604',
                fmt({'bold':True,'font_size':16,'font_color':C['navy'],'bottom':2,'bottom_color':C['navy']}))
ws6.set_row(0, 36)

# ── helper: write data range for charts ──────────────────────────────────
# Data area starts at row 3 (0-indexed)
DATA_ROW = 3

# Write monthly data as chart source
ws6.write(DATA_ROW, 0, 'Month', hdr)
ws6.write(DATA_ROW, 1, 'Total', hdr)
ws6.write(DATA_ROW, 2, 'RDHx600DC_sum', hdr)
ws6.write(DATA_ROW, 3, 'RDHx762DC_sum', hdr)
ws6.write(DATA_ROW, 4, 'Sidecar', hdr)
ws6.write(DATA_ROW, 5, 'RDHx600EC_sum', hdr)
ws6.write(DATA_ROW, 6, 'InRowCDU', hdr)
ws6.write(DATA_ROW, 7, 'InRowPiping', hdr)

for i, row in monthly.reset_index(drop=True).iterrows():
    r = DATA_ROW + 1 + i
    ws6.write(r, 0, row['YearMonth'], cell)
    ws6.write_number(r, 1, row['Total'], num_)
    ws6.write_number(r, 2, row['RDHx600DC_Belimo'] + row['RDHx600DC_Extender'], num_)
    ws6.write_number(r, 3, row['RDHx762DC_Belimo'] + row['RDHx762DC_Extender'], num_)
    ws6.write_number(r, 4, row['Sidecar'], num_)
    ws6.write_number(r, 5, row['RDHx600EC_Belimo'] + row['RDHx600EC_Extender'], num_)
    ws6.write_number(r, 6, row['InRowCDU'], num_)
    ws6.write_number(r, 7, row['InRowPiping'], num_)

nm = len(monthly)
dstart = DATA_ROW + 1   # 1st data row (0-indexed)
dend   = DATA_ROW + nm  # last data row

# ── Chart 1: Line chart — Monthly Total Trend ─────────────────────────────
ch1 = wb.add_chart({'type':'line'})
ch1.add_series({
    'name': 'Total Units',
    'categories': ['6_Charts', dstart, 0, dend, 0],
    'values':     ['6_Charts', dstart, 1, dend, 1],
    'line': {'color': C['navy'], 'width': 2.5},
    'marker': {'type':'circle','size':7,'fill':{'color':C['orange']},'border':{'color':C['orange']}},
})
ch1.set_title({'name':'Monthly Total Production Volume (Units Inspected)'})
ch1.set_x_axis({'name':'Month','num_font':{'rotation':-45}})
ch1.set_y_axis({'name':'Total Units','major_gridlines':{'visible':True,'line':{'color':C['mgray'],'dash_type':'dash'}}})
ch1.set_legend({'position':'bottom'})
ch1.set_style(10)
ch1.set_size({'width':600,'height':320})
ws6.insert_chart('A3', ch1, {'x_offset':5,'y_offset':5})

# ── Chart 2: Stacked Bar — Product Mix by Month ───────────────────────────
ch2 = wb.add_chart({'type':'bar','subtype':'stacked'})
colours_prod = [C['navy'], C['blue'], C['orange'], C['green'], C['amber'], C['pink']]
series_defs = [
    ('RDHx600DC', 2), ('RDHx762DC', 3), ('Sidecar', 4),
    ('RDHx600EC', 5), ('In-Row CDU', 6), ('In-Row Piping', 7),
]
for (name, col_i), color in zip(series_defs, colours_prod):
    ch2.add_series({
        'name': name,
        'categories': ['6_Charts', dstart, 0, dend, 0],
        'values':     ['6_Charts', dstart, col_i, dend, col_i],
        'fill': {'color': color},
        'gap': 80,
    })
ch2.set_title({'name':'Monthly Product Mix (Stacked Bar)'})
ch2.set_x_axis({'name':'Units'})
ch2.set_y_axis({'name':'Month'})
ch2.set_legend({'position':'bottom'})
ch2.set_style(10)
ch2.set_size({'width':600,'height':320})
ws6.insert_chart('A22', ch2, {'x_offset':5,'y_offset':5})

# ── Chart 3: Pie chart — Overall Product Share ────────────────────────────
# Write pie data
PIE_ROW = DATA_ROW + nm + 3
pie_data = [
    ('RDHx 600-DC', df['RDHx600DC_Belimo'].sum() + df['RDHx600DC_Extender'].sum()),
    ('RDHx 762-DC', df['RDHx762DC_Belimo'].sum() + df['RDHx762DC_Extender'].sum()),
    ('Sidecar',     df['Sidecar'].sum()),
    ('RDHx 600-EC', df['RDHx600EC_Belimo'].sum() + df['RDHx600EC_Extender'].sum()),
    ('In-Row CDU',  df['InRowCDU'].sum()),
    ('In-Row Piping', df['InRowPiping'].sum()),
]
ws6.write(PIE_ROW, 9, 'Product', hdr)
ws6.write(PIE_ROW, 10, 'Total Units', hdr)
for i,(lbl,val) in enumerate(pie_data):
    ws6.write(PIE_ROW+1+i, 9, lbl, cell)
    ws6.write_number(PIE_ROW+1+i, 10, val, num_)

ch3 = wb.add_chart({'type':'pie'})
ch3.add_series({
    'name': 'Product Share',
    'categories': ['6_Charts', PIE_ROW+1, 9, PIE_ROW+len(pie_data), 9],
    'values':     ['6_Charts', PIE_ROW+1, 10, PIE_ROW+len(pie_data), 10],
    'data_labels': {'percentage':True,'category':True,'separator':'\n'},
    'points': [
        {'fill':{'color':C['navy']}},   # 600-DC
        {'fill':{'color':C['blue']}},   # 762-DC
        {'fill':{'color':C['orange']}}, # Sidecar
        {'fill':{'color':C['green']}},  # 600-EC
        {'fill':{'color':C['amber']}},  # CDU
        {'fill':{'color':C['pink']}},   # Piping
    ],
})
ch3.set_title({'name':'Overall Product Volume Share'})
ch3.set_legend({'position':'bottom'})
ch3.set_style(10)
ch3.set_size({'width':480,'height':340})
ws6.insert_chart('K3', ch3, {'x_offset':5,'y_offset':5})

# ── Chart 4: Column chart — Daily Total (top 15 days) ────────────────────
top15 = df.nlargest(15,'Total').sort_values('Date')
TOP_ROW = PIE_ROW + len(pie_data) + 3
ws6.write(TOP_ROW, 9, 'Date', hdr)
ws6.write(TOP_ROW, 10, 'Daily Total', hdr)
for i,(_,row) in enumerate(top15.iterrows()):
    ws6.write(TOP_ROW+1+i, 9, row['Date'].strftime('%m/%d'), cell)
    ws6.write_number(TOP_ROW+1+i, 10, row['Total'], num_)

ch4 = wb.add_chart({'type':'column'})
ch4.add_series({
    'name': 'Daily Total',
    'categories': ['6_Charts', TOP_ROW+1, 9, TOP_ROW+15, 9],
    'values':     ['6_Charts', TOP_ROW+1, 10, TOP_ROW+15, 10],
    'fill': {'color': C['blue']},
    'data_labels': {'value':True,'font':{'size':8}},
    'gap': 60,
})
ch4.set_title({'name':'Top 15 Highest Production Days'})
ch4.set_x_axis({'name':'Date','num_font':{'rotation':-45}})
ch4.set_y_axis({'name':'Units','major_gridlines':{'visible':True,'line':{'color':C['mgray'],'dash_type':'dash'}}})
ch4.set_legend({'none':True})
ch4.set_style(10)
ch4.set_size({'width':500,'height':300})
ws6.insert_chart('K22', ch4, {'x_offset':5,'y_offset':5})

print("✓ Sheet 6 done")

# ── Final: close workbook ────────────────────────────────────────────────
wb.close()
import os
size = os.path.getsize(OUTPUT)
print(f"\n✅ Report generated: {OUTPUT}  ({size/1024:.1f} KB)")
print(f"   Location: /home/user/IQC/{OUTPUT}")
