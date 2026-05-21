import subprocess
import sys

# Auto-install required packages if missing
def install(pkg):
    subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

try:
    import pandas as pd
except ImportError:
    print("Installing pandas...")
    install("pandas")
    import pandas as pd

try:
    import openpyxl
except ImportError:
    print("Installing openpyxl...")
    install("openpyxl")

import json
import re
import os

# ============================================================
# CONFIGURATION
# ============================================================
EXCEL_FILE = "Indicadores_BCRP_MultiFrecuencia.xlsx"
OUTPUT_HTML = "Peru_Macro_Dashboard.html"

# ============================================================
# READ EXCEL
# ============================================================
print(f"\nReading {EXCEL_FILE}...")

try:
    bcrp_q = pd.read_excel(EXCEL_FILE, sheet_name='Trimestral')
    bcrp_m = pd.read_excel(EXCEL_FILE, sheet_name='Mensual')
    bcrp_a = pd.read_excel(EXCEL_FILE, sheet_name='Anual')
    print(f"  Trimestral: {len(bcrp_q)} rows")
    print(f"  Mensual:    {len(bcrp_m)} rows")
    print(f"  Anual:      {len(bcrp_a)} rows")
except FileNotFoundError:
    print(f"\nERROR: Could not find '{EXCEL_FILE}'")
    print("Make sure the Excel file is in the same folder as this script.")
    input("\nPress Enter to exit...")
    sys.exit()
except Exception as e:
    print(f"\nERROR reading Excel: {e}")
    input("\nPress Enter to exit...")
    sys.exit()

# ============================================================
# DATE PARSERS
# ============================================================
MONTH_ESP = {
    'Ene':1,'Feb':2,'Mar':3,'Abr':4,'May':5,'Jun':6,
    'Jul':7,'Ago':8,'Sep':9,'Oct':10,'Nov':11,'Dic':12
}

def parse_quarter(s):
    try:
        m = re.match(r'T(\d)\.(\d+)', str(s))
        if m:
            q, y = int(m.group(1)), int(m.group(2))
            y = 2000 + y if y < 100 else y
            return f'{y} Q{q}'
    except:
        pass
    return None

def parse_month(s):
    try:
        parts = str(s).split('.')
        mo = MONTH_ESP.get(parts[0])
        y = int(parts[1])
        if mo:
            return f'{y}-{mo:02d}'
    except:
        pass
    return None

# ============================================================
# SERIES EXTRACTORS
# ============================================================
def q_series(col):
    if col not in bcrp_q.columns:
        print(f"  WARNING: '{col}' not found in Trimestral sheet")
        return []
    df = bcrp_q[['Fechas', col]].copy()
    df['x'] = df['Fechas'].apply(parse_quarter)
    df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna()
    return [{'x': r['x'], 'y': round(float(r[col]), 2)} for _, r in df.iterrows()]

def m_series(col):
    if col not in bcrp_m.columns:
        print(f"  WARNING: '{col}' not found in Mensual sheet")
        return []
    df = bcrp_m[['Fechas', col]].copy()
    df['x'] = df['Fechas'].apply(parse_month)
    df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna()
    return [{'x': r['x'], 'y': round(float(r[col]), 2)} for _, r in df.iterrows()]

def a_series(col):
    if col not in bcrp_a.columns:
        print(f"  WARNING: '{col}' not found in Anual sheet")
        return []
    df = bcrp_a[['Fechas', col]].copy()
    df[col] = pd.to_numeric(df[col], errors='coerce')
    df = df.dropna()
    return [{'x': str(int(r['Fechas'])), 'y': round(float(r[col]), 2)} for _, r in df.iterrows()]

# ============================================================
# EXTRACT ALL SERIES
# ============================================================
print("\nExtracting data...")

gdp    = q_series('Economic_Growth_GDP_var%')
cons   = q_series('Private_Consumption_var%')
inv    = q_series('Fixed_Investment_var%')
govcon = q_series('Government_Consumption_var%')
domdem = q_series('Domestic_Demand_var%')
exports= q_series('Exports_var%')
imports= q_series('Imports_var%')
ca     = q_series('Current_Account_%GDP')

tb = []
tb_cols = ['Trade_Balance_USD_bn','Merchandise_Exports_USD_bn','Merchandise_Imports_USD_bn']
if all(c in bcrp_q.columns for c in tb_cols):
    tdf = bcrp_q[['Fechas'] + tb_cols].copy()
    tdf['x'] = tdf['Fechas'].apply(parse_quarter)
    tdf = tdf.dropna()
    for _, r in tdf.iterrows():
        try:
            tb.append({
                'x':       r['x'],
                'balance': round(float(r['Trade_Balance_USD_bn']), 2),
                'exports': round(float(r['Merchandise_Exports_USD_bn']), 2),
                'imports': round(float(r['Merchandise_Imports_USD_bn']), 2),
            })
        except:
            pass

cpi     = m_series('Inflation_CPI_annvar%_eop')
cpi_mom = m_series('Inflation_CPI_momvar%')
rate    = m_series('Policy_Interest_Rate_%eop')
bond    = m_series('10Year_Bond Yield_%eop')
fx      = m_series('Exchange _Rate_PENperUSD_eop')
unemp   = m_series('Unemployment_%of_active_population_eop')
activity= m_series('Economic_Activity_annvar%')
retail  = m_series('Retail_Sales_annvar%')
manuf_m = m_series('Manufacturing_Production_annvar%')
banking = m_series('Official_Banking_Rate_%eop')

pubdebt  = a_series('Public_Debt_%GDP')
intres   = a_series('International_Reserves_USD_MM')
cpi_eop  = a_series('Inflation_CPI_eop%')
priv_inv = a_series('Fixed_Private_Investment_var%')
pub_inv  = a_series('Fixed_Public_Investment_var%')

sec = []
sec_cols = ['Economic_Growth_GDP_var%','Agriculture_var%','Minning_var%',
            'Construction_var%','Services_var%','Manufacturing_var%']
if all(c in bcrp_a.columns for c in sec_cols):
    for _, r in bcrp_a.iterrows():
        try:
            sec.append({
                'x':            str(int(r['Fechas'])),
                'gdp':          round(float(r['Economic_Growth_GDP_var%']), 2),
                'agri':         round(float(r['Agriculture_var%']), 2),
                'mining':       round(float(r['Minning_var%']), 2),
                'construction': round(float(r['Construction_var%']), 2),
                'services':     round(float(r['Services_var%']), 2),
                'manufacturing':round(float(r['Manufacturing_var%']), 2),
            })
        except:
            pass

# ============================================================
# KPIs
# ============================================================
def latest_val(s): return s[-1]['y'] if s else 0
def latest_lbl(s): return s[-1]['x'] if s else ''

kpis = {
    'gdp':        latest_val(gdp),
    'cpi':        latest_val(cpi),
    'rate':       latest_val(rate),
    'fx':         latest_val(fx),
    'unemp':      latest_val(unemp),
    'ca':         latest_val(ca),
    'gdp_label':  latest_lbl(gdp),
    'cpi_label':  latest_lbl(cpi),
    'rate_label': latest_lbl(rate),
}

print(f"\n  Latest values:")
print(f"  GDP:          {kpis['gdp']}%  ({kpis['gdp_label']})")
print(f"  CPI:          {kpis['cpi']}%  ({kpis['cpi_label']})")
print(f"  Policy Rate:  {kpis['rate']}%")
print(f"  FX:           {kpis['fx']} PEN/USD")
print(f"  Unemployment: {kpis['unemp']}%")
print(f"  Curr Account: {kpis['ca']}% GDP")

# ============================================================
# PACK DATA
# ============================================================
data = {
    'gdp':gdp,'cons':cons,'inv':inv,'govcon':govcon,'domdem':domdem,
    'exports':exports,'imports':imports,'ca':ca,'tb':tb,
    'cpi':cpi,'cpi_mom':cpi_mom,'rate':rate,'bond':bond,'fx':fx,
    'unemp':unemp,'activity':activity,'retail':retail,
    'manuf_m':manuf_m,'banking':banking,
    'sec':sec,'pubdebt':pubdebt,'intres':intres,
    'cpi_eop':cpi_eop,'priv_inv':priv_inv,'pub_inv':pub_inv,
    'kpis':kpis,
}
DATA_JSON = json.dumps(data, separators=(',',':'))

# ============================================================
# HTML TEMPLATE (complete dashboard with all legends)
# ============================================================
print(f"\nBuilding {OUTPUT_HTML}...")

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Peru Macro Dashboard - BCRP</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;500;600;700&family=IBM+Plex+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{--navy:#0f1f3d;--navy2:#1b2a4a;--blue:#2e5090;--blueL:#4a73c4;--red:#c0392b;--gold:#d4a017;--green:#27ae60;--purple:#8e44ad;--orange:#e67e22;--bg:#f4f6fb;--border:#dde3f0;--text:#1a1a2e;--muted:#6b7a99;--sw:210px;}
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'IBM Plex Sans',sans-serif;background:var(--bg);color:var(--text);display:flex;min-height:100vh;}
#sidebar{width:var(--sw);min-height:100vh;background:var(--navy);display:flex;flex-direction:column;position:fixed;left:0;top:0;z-index:100;}
.sb-brand{padding:22px 18px 18px;border-bottom:1px solid rgba(255,255,255,0.08);}
.sb-logo{width:34px;height:34px;background:var(--red);border-radius:6px;display:flex;align-items:center;justify-content:center;font-weight:700;color:#fff;font-size:13px;margin-bottom:9px;}
.sb-title{color:#fff;font-size:13px;font-weight:600;}
.sb-sub{color:rgba(255,255,255,0.4);font-size:10px;margin-top:2px;text-transform:uppercase;letter-spacing:.5px;}
.sb-sec{padding:18px 12px 6px;}
.sb-sec-lbl{color:rgba(255,255,255,0.28);font-size:9px;font-weight:600;letter-spacing:1.5px;text-transform:uppercase;margin-bottom:7px;padding:0 6px;}
.nav-btn{display:flex;align-items:center;gap:9px;width:100%;padding:9px 10px;background:none;border:none;border-radius:7px;cursor:pointer;color:rgba(255,255,255,0.55);font-family:inherit;font-size:12px;font-weight:500;transition:all .15s;text-align:left;margin-bottom:2px;}
.nav-btn:hover{background:rgba(255,255,255,0.07);color:rgba(255,255,255,0.9);}
.nav-btn.active{background:var(--blue);color:#fff;}
.nav-ic{width:15px;text-align:center;font-size:12px;flex-shrink:0;}
.sb-foot{margin-top:auto;padding:14px 18px;border-top:1px solid rgba(255,255,255,0.06);}
.sb-src{color:rgba(255,255,255,0.28);font-size:9px;letter-spacing:.4px;line-height:1.8;}
#main{margin-left:var(--sw);flex:1;display:flex;flex-direction:column;}
.topbar{background:#fff;border-bottom:1px solid var(--border);padding:0 26px;height:56px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:50;}
.tb-title{font-size:15px;font-weight:700;color:var(--navy);}
.tb-sub{font-size:10px;color:var(--muted);margin-top:1px;}
.tb-right{display:flex;align-items:center;gap:10px;}
.badge{background:var(--bg);border:1px solid var(--border);border-radius:20px;padding:4px 11px;font-size:10px;color:var(--muted);}
.flag{background:var(--navy2);color:#fff;border-radius:20px;padding:4px 13px;font-size:10px;font-weight:600;}
.page{padding:22px 26px;display:none;flex-direction:column;gap:18px;}
.page.active{display:flex;}
.kpi-row{display:grid;grid-template-columns:repeat(3,1fr);gap:13px;}
.kc{background:#fff;border-radius:10px;border:1px solid var(--border);overflow:hidden;position:relative;}
.kc::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--blue);z-index:1;}
.kc.red::before{background:var(--red);}
.kc.gold::before{background:var(--gold);}
.kc.green::before{background:var(--green);}
.kc.purple::before{background:var(--purple);}
.kc-top{padding:13px 15px 6px;}
.kc-lbl{font-size:10px;color:var(--muted);font-weight:600;text-transform:uppercase;letter-spacing:.8px;margin-bottom:5px;}
.kc-val{font-size:25px;font-weight:700;color:var(--navy);font-family:'IBM Plex Mono',monospace;line-height:1;}
.kc-unit{font-size:10px;color:var(--muted);margin-top:3px;}
.kc-per{font-size:9px;color:var(--muted);font-style:italic;margin-top:2px;}
.kc-spark{height:65px;padding:0 0 6px;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:15px;}
.cc{background:#fff;border-radius:10px;border:1px solid var(--border);padding:16px 18px;}
.ch{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:13px;}
.ct{font-size:13px;font-weight:600;color:var(--navy);}
.cs{font-size:10px;color:var(--muted);margin-top:2px;}
.cbadge{font-size:9px;background:var(--bg);border:1px solid var(--border);border-radius:4px;padding:2px 6px;color:var(--muted);white-space:nowrap;}
.cw{position:relative;height:195px;}
.cw.t{height:230px;}
.leg{display:flex;gap:14px;margin-top:10px;flex-wrap:wrap;padding:7px 12px;background:#eef1f8;border-radius:6px;border-left:3px solid #2e5090;}
.li{display:flex;align-items:center;gap:7px;font-size:11px;font-weight:600;color:#1a1a2e;}
.ld{width:13px;height:13px;border-radius:3px;flex-shrink:0;box-shadow:0 1px 2px rgba(0,0,0,0.2);}
.ll{width:24px;height:3px;flex-shrink:0;border-radius:2px;box-shadow:0 1px 2px rgba(0,0,0,0.2);}
.sec-row{display:flex;align-items:center;gap:11px;}
.sec-lbl{font-size:10px;font-weight:700;color:var(--blue);text-transform:uppercase;letter-spacing:1px;white-space:nowrap;}
.sec-hr{flex:1;height:1px;background:var(--border);}
</style>
</head>
<body>
<nav id="sidebar">
  <div class="sb-brand"><div class="sb-logo">PE</div><div class="sb-title">Macro Dashboard</div><div class="sb-sub">BCRP - Peru</div></div>
  <div class="sb-sec">
    <div class="sb-sec-lbl">Pages</div>
    <button class="nav-btn active" onclick="go('overview',this)"><span class="nav-ic">&#9672;</span>Overview</button>
    <button class="nav-btn" onclick="go('activity',this)"><span class="nav-ic">&#9638;</span>National Accounts</button>
    <button class="nav-btn" onclick="go('prices',this)"><span class="nav-ic">&#9673;</span>Prices</button>
    <button class="nav-btn" onclick="go('external',this)"><span class="nav-ic">&#9680;</span>External Sector</button>
    <button class="nav-btn" onclick="go('labor',this)"><span class="nav-ic">&#9681;</span>Labor &amp; Sectoral</button>
    <button class="nav-btn" onclick="go('rates',this)"><span class="nav-ic">&#9682;</span>Rates &amp; FX</button>
    <button class="nav-btn" onclick="go('fiscal',this)"><span class="nav-ic">&#9683;</span>Fiscal</button>
  </div>
  <div class="sb-foot"><div class="sb-src">Source: BCRP<br>Coverage: 2010-2025<br>Works offline</div></div>
</nav>
<div id="main">
  <div class="topbar">
    <div><div class="tb-title" id="tt">Overview</div><div class="tb-sub" id="ts">Key macroeconomic indicators - Peru</div></div>
    <div class="tb-right"><span class="badge">2010-2025</span><span class="flag">Peru BCRP</span></div>
  </div>

  <!-- OVERVIEW -->
  <div class="page active" id="page-overview">
    <div class="kpi-row">
      <div class="kc"><div class="kc-top"><div class="kc-lbl">GDP Growth</div><div class="kc-val" id="v-gdp">-</div><div class="kc-unit">% year-on-year</div><div class="kc-per" id="p-gdp"></div></div><div class="kc-spark"><canvas id="sp-gdp"></canvas></div></div>
      <div class="kc red"><div class="kc-top"><div class="kc-lbl">CPI Inflation</div><div class="kc-val" id="v-cpi">-</div><div class="kc-unit">% year-on-year</div><div class="kc-per" id="p-cpi"></div></div><div class="kc-spark"><canvas id="sp-cpi"></canvas></div></div>
      <div class="kc"><div class="kc-top"><div class="kc-lbl">Policy Rate (BCR)</div><div class="kc-val" id="v-rate">-</div><div class="kc-unit">% end of period</div><div class="kc-per" id="p-rate"></div></div><div class="kc-spark"><canvas id="sp-rate"></canvas></div></div>
      <div class="kc gold"><div class="kc-top"><div class="kc-lbl">Exchange Rate</div><div class="kc-val" id="v-fx">-</div><div class="kc-unit">PEN per USD</div><div class="kc-per">End of period</div></div><div class="kc-spark"><canvas id="sp-fx"></canvas></div></div>
      <div class="kc purple"><div class="kc-top"><div class="kc-lbl">Unemployment</div><div class="kc-val" id="v-unemp">-</div><div class="kc-unit">% active population</div><div class="kc-per">Lima Metropolitan</div></div><div class="kc-spark"><canvas id="sp-unemp"></canvas></div></div>
      <div class="kc green"><div class="kc-top"><div class="kc-lbl">Current Account</div><div class="kc-val" id="v-ca">-</div><div class="kc-unit">% of GDP</div><div class="kc-per">Quarterly</div></div><div class="kc-spark"><canvas id="sp-ca"></canvas></div></div>
    </div>
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">GDP Growth</div><div class="cs">% YoY - Quarterly</div></div><span class="cbadge">% YoY</span></div><div class="cw"><canvas id="c-gdp-ov"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--red)"></div>GDP var%</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>4-quarter moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">CPI Inflation</div><div class="cs">% YoY - Monthly</div></div><span class="cbadge">% YoY</span></div><div class="cw"><canvas id="c-cpi-ov"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--red)"></div>CPI year-on-year %</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>3-month moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Policy Rate and 10Y Bond</div><div class="cs">% - Monthly</div></div><span class="cbadge">%</span></div><div class="cw"><canvas id="c-rate-ov"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--red)"></div>BCR policy rate</div><div class="li"><div class="ll" style="background:var(--blue)"></div>10Y government bond</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Exchange Rate PEN/USD</div><div class="cs">End of period - Monthly</div></div><span class="cbadge">PEN/USD</span></div><div class="cw"><canvas id="c-fx-ov"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--gold)"></div>PEN per USD (higher = weaker sol)</div></div></div>
    </div>
  </div>

  <!-- NATIONAL ACCOUNTS -->
  <div class="page" id="page-activity">
    <div class="sec-row"><span class="sec-lbl">GDP and Demand Components</span><div class="sec-hr"></div></div>
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">GDP Growth</div><div class="cs">% YoY - Quarterly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-gdp-na"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--red)"></div>GDP growth %</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>4-quarter moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Domestic Demand vs Private Consumption</div><div class="cs">% YoY - Quarterly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-dom-na"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--blue)"></div>Domestic demand</div><div class="li"><div class="ll" style="background:var(--gold)"></div>Private consumption</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Fixed Investment vs Govt Consumption</div><div class="cs">% YoY - Quarterly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-inv-na"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--blue)"></div>Fixed investment</div><div class="li"><div class="ll" style="background:var(--green)"></div>Govt consumption</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Economic Activity - Monthly PBI</div><div class="cs">% YoY - Monthly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-act-na"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--red)"></div>Monthly activity % YoY</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>3-month moving avg</div></div></div>
    </div>
  </div>

  <!-- PRICES -->
  <div class="page" id="page-prices">
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">CPI Inflation YoY</div><div class="cs">% YoY - Monthly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-cpi-p"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--red)"></div>CPI year-on-year %</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>3-month moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">CPI Inflation MoM</div><div class="cs">% MoM - Monthly</div></div><span class="cbadge">% MoM</span></div><div class="cw t"><canvas id="c-cpimom-p"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--orange)"></div>CPI month-on-month %</div><div class="li"><div class="ll" style="background:var(--blue)"></div>3-month moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Policy Rate vs 10Y Bond vs Banking Rate</div><div class="cs">% - Monthly</div></div><span class="cbadge">%</span></div><div class="cw t"><canvas id="c-rb-p"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--red)"></div>BCR policy rate</div><div class="li"><div class="ll" style="background:var(--blue)"></div>10Y bond yield</div><div class="li"><div class="ll" style="background:var(--gold)"></div>Banking rate</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">CPI End of Period Annual</div><div class="cs">% - Annual</div></div><span class="cbadge">Annual</span></div><div class="cw t"><canvas id="c-cpieop-p"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--red)"></div>CPI end of year %</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>3-year trend</div></div></div>
    </div>
  </div>

  <!-- EXTERNAL -->
  <div class="page" id="page-external">
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">Trade Balance</div><div class="cs">USD billions - Quarterly</div></div><span class="cbadge">USD bn</span></div><div class="cw t"><canvas id="c-tb-ex"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--blue)"></div>Exports USD bn</div><div class="li"><div class="ld" style="background:var(--red)"></div>Imports USD bn</div><div class="li"><div class="ll" style="background:var(--green)"></div>Trade balance</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Current Account</div><div class="cs">% of GDP - Quarterly</div></div><span class="cbadge">% GDP</span></div><div class="cw t"><canvas id="c-ca-ex"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--green)"></div>Surplus (positive)</div><div class="li"><div class="ld" style="background:var(--red)"></div>Deficit (negative)</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Exports and Imports Growth</div><div class="cs">% YoY - Quarterly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-exim-ex"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--blue)"></div>Exports growth % YoY</div><div class="li"><div class="ll" style="background:var(--red)"></div>Imports growth % YoY</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">International Reserves</div><div class="cs">USD millions - Annual</div></div><span class="cbadge">USD MM</span></div><div class="cw t"><canvas id="c-res-ex"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--green)"></div>International reserves USD MM</div></div></div>
    </div>
  </div>

  <!-- LABOR -->
  <div class="page" id="page-labor">
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">Unemployment Rate</div><div class="cs">% active population - Monthly</div></div><span class="cbadge">%</span></div><div class="cw t"><canvas id="c-unemp-l"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--purple)"></div>Unemployment rate % (Lima)</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Retail Sales</div><div class="cs">% YoY - Monthly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-retail-l"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--blue)"></div>Retail sales % YoY</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>3-month moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Sectoral Growth Annual</div><div class="cs">% YoY by sector</div></div><span class="cbadge">Annual</span></div><div class="cw t"><canvas id="c-sec-l"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--blue)"></div>Agriculture</div><div class="li"><div class="ld" style="background:var(--red)"></div>Mining</div><div class="li"><div class="ld" style="background:var(--green)"></div>Construction</div><div class="li"><div class="ld" style="background:var(--gold)"></div>Services</div><div class="li"><div class="ld" style="background:var(--purple)"></div>Manufacturing</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Manufacturing Production</div><div class="cs">% YoY - Monthly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-manuf-l"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--navy)"></div>Manufacturing % YoY</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>3-month moving avg</div></div></div>
    </div>
  </div>

  <!-- RATES -->
  <div class="page" id="page-rates">
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">Policy Rate BCR</div><div class="cs">% end of period - Monthly</div></div><span class="cbadge">% eop</span></div><div class="cw t"><canvas id="c-rate-r"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--red)"></div>BCR reference rate %</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">10-Year Bond Yield</div><div class="cs">% end of period - Monthly</div></div><span class="cbadge">%</span></div><div class="cw t"><canvas id="c-bond-r"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--blue)"></div>10Y government bond yield %</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Exchange Rate PEN/USD</div><div class="cs">End of period - Monthly</div></div><span class="cbadge">PEN/USD</span></div><div class="cw t"><canvas id="c-fx-r"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--gold)"></div>PEN per USD (higher = weaker sol)</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">All Rates Overlaid</div><div class="cs">Policy - 10Y Bond - Banking - Monthly</div></div><span class="cbadge">%</span></div><div class="cw t"><canvas id="c-all-r"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--red)"></div>BCR policy rate</div><div class="li"><div class="ll" style="background:var(--blue)"></div>10Y bond yield</div><div class="li"><div class="ll" style="background:var(--gold)"></div>Banking rate</div></div></div>
    </div>
  </div>

  <!-- FISCAL -->
  <div class="page" id="page-fiscal">
    <div class="g2">
      <div class="cc"><div class="ch"><div><div class="ct">Public Debt</div><div class="cs">% of GDP - Annual</div></div><span class="cbadge">% GDP</span></div><div class="cw t"><canvas id="c-debt-f"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--red)"></div>Public debt % of GDP</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Government Consumption</div><div class="cs">% YoY - Quarterly</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-gov-f"></canvas></div><div class="leg"><div class="li"><div class="ld" style="background:var(--blue)"></div>Govt consumption % YoY</div><div class="li"><div class="ll" style="background:var(--blueL)"></div>4-quarter moving avg</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">Private vs Public Investment</div><div class="cs">% YoY - Annual</div></div><span class="cbadge">% YoY</span></div><div class="cw t"><canvas id="c-inv-f"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--blue)"></div>Private investment % YoY</div><div class="li"><div class="ll" style="background:var(--red)"></div>Public investment % YoY</div></div></div>
      <div class="cc"><div class="ch"><div><div class="ct">International Reserves</div><div class="cs">USD millions - Annual</div></div><span class="cbadge">USD MM</span></div><div class="cw t"><canvas id="c-res-f"></canvas></div><div class="leg"><div class="li"><div class="ll" style="background:var(--green)"></div>International reserves USD MM</div></div></div>
    </div>
  </div>

</div>
<script>
const D=__DATA__;
const C={red:'#c0392b',redA:'rgba(192,57,43,0.10)',blue:'#2e5090',blueA:'rgba(46,80,144,0.09)',blueL:'#4a73c4',gold:'#d4a017',goldA:'rgba(212,160,23,0.10)',green:'#27ae60',greenA:'rgba(39,174,96,0.09)',purple:'#8e44ad',purpleA:'rgba(142,68,173,0.09)',orange:'#e67e22',navy:'#0f1f3d'};
function ma(a,n){return a.map((_,i)=>{if(i<n-1)return null;return a.slice(i-n+1,i+1).reduce((s,v)=>s+(v||0),0)/n;});}
function sortQ(a){return[...a].sort((a,b)=>{const[ya,qa]=a.x.split(' Q'),[yb,qb]=b.x.split(' Q');return ya!=yb?+ya-+yb:+qa-+qb;});}
const bO=(n=8)=>({responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{mode:'index',intersect:false,backgroundColor:'rgba(15,31,61,0.92)',titleFont:{size:10},bodyFont:{size:10},padding:9,cornerRadius:5}},scales:{x:{ticks:{font:{size:9},maxTicksLimit:n,color:'#6b7a99'},grid:{display:false}},y:{ticks:{font:{size:9},color:'#6b7a99'},grid:{color:'rgba(0,0,0,0.04)'}}}});
function spark(id,lb,vl,col){const el=document.getElementById(id);if(!el)return;const n=Math.min(24,lb.length),l=lb.slice(-n),v=vl.slice(-n),r=parseInt(col.slice(1,3),16),g=parseInt(col.slice(3,5),16),b=parseInt(col.slice(5,7),16);new Chart(el,{type:'line',data:{labels:l,datasets:[{data:v,borderColor:col,borderWidth:1.5,pointRadius:0,tension:0.4,fill:true,backgroundColor:`rgba(${r},${g},${b},0.08)`}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},tooltip:{enabled:false}},scales:{x:{display:false},y:{display:false}}}});}
function barLine(id,lb,bars,line,bc,lc){const el=document.getElementById(id);if(!el)return;new Chart(el,{type:'bar',data:{labels:lb,datasets:[{type:'bar',data:bars,backgroundColor:bars.map(v=>v>=0?bc+'bb':C.red+'bb'),borderRadius:2,borderSkipped:false},{type:'line',data:line,borderColor:lc,borderWidth:2,pointRadius:0,tension:0.4,fill:false}]},options:bO()});}
function lines(id,lb,sets){const el=document.getElementById(id);if(!el)return;new Chart(el,{type:'line',data:{labels:lb,datasets:sets},options:bO()});}
function stacked(id,lb,sets){const el=document.getElementById(id);if(!el)return;const o=bO(16);o.scales.x.stacked=true;o.scales.y.stacked=true;new Chart(el,{type:'bar',data:{labels:lb,datasets:sets},options:o});}
function singleBar(id,lb,vl,fn){const el=document.getElementById(id);if(!el)return;new Chart(el,{type:'bar',data:{labels:lb,datasets:[{data:vl,backgroundColor:vl.map(fn),borderRadius:3}]},options:bO(16)});}
function areaLine(id,lb,vl,col,colA){const el=document.getElementById(id);if(!el)return;new Chart(el,{type:'line',data:{labels:lb,datasets:[{data:vl,borderColor:col,borderWidth:2.5,pointRadius:0,tension:0.3,fill:true,backgroundColor:colA}]},options:bO()});}
document.getElementById('v-gdp').textContent=D.kpis.gdp.toFixed(2);
document.getElementById('v-cpi').textContent=D.kpis.cpi.toFixed(2);
document.getElementById('v-rate').textContent=D.kpis.rate.toFixed(2);
document.getElementById('v-fx').textContent=D.kpis.fx.toFixed(2);
document.getElementById('v-unemp').textContent=D.kpis.unemp.toFixed(2);
document.getElementById('v-ca').textContent=D.kpis.ca.toFixed(2);
document.getElementById('p-gdp').textContent=D.kpis.gdp_label;
document.getElementById('p-cpi').textContent=D.kpis.cpi_label;
document.getElementById('p-rate').textContent=D.kpis.rate_label;
const gS=sortQ(D.gdp),cS=sortQ(D.ca);
spark('sp-gdp',gS.map(d=>d.x),gS.map(d=>d.y),C.blue);
spark('sp-cpi',D.cpi.map(d=>d.x),D.cpi.map(d=>d.y),C.red);
spark('sp-rate',D.rate.map(d=>d.x),D.rate.map(d=>d.y),C.blue);
spark('sp-fx',D.fx.map(d=>d.x),D.fx.map(d=>d.y),C.gold);
spark('sp-unemp',D.unemp.map(d=>d.x),D.unemp.map(d=>d.y),C.purple);
spark('sp-ca',cS.map(d=>d.x),cS.map(d=>d.y),C.green);
const META={overview:{t:'Overview',s:'Key macroeconomic indicators - Peru'},activity:{t:'National Accounts',s:'GDP, demand components, economic activity'},prices:{t:'Prices and Monetary',s:'CPI inflation, policy rate, exchange rate'},external:{t:'External Sector',s:'Trade balance, current account, exports and imports'},labor:{t:'Labor and Sectoral',s:'Employment, retail, sectoral growth'},rates:{t:'Interest Rates and FX',s:'BCR policy rate, bond yields, exchange rate'},fiscal:{t:'Fiscal',s:'Public debt, government consumption, investment'}};
const done=new Set(['overview']);
function go(id,btn){document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));document.querySelectorAll('.nav-btn').forEach(b=>b.classList.remove('active'));document.getElementById('page-'+id).classList.add('active');btn.classList.add('active');document.getElementById('tt').textContent=META[id].t;document.getElementById('ts').textContent=META[id].s;if(!done.has(id)){done.add(id);rP(id);}}
function rP(id){if(id==='overview')rOv();else if(id==='activity')rAc();else if(id==='prices')rPr();else if(id==='external')rEx();else if(id==='labor')rLa();else if(id==='rates')rRa();else if(id==='fiscal')rFi();}
function rOv(){const g=sortQ(D.gdp),gV=g.map(d=>d.y);barLine('c-gdp-ov',g.map(d=>d.x),gV,ma(gV,4),C.red,C.blueL);const cV=D.cpi.map(d=>d.y);barLine('c-cpi-ov',D.cpi.map(d=>d.x),cV,ma(cV,3),C.red,C.blueL);lines('c-rate-ov',D.rate.map(d=>d.x),[{label:'Policy Rate',data:D.rate.map(d=>d.y),borderColor:C.red,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'10Y Bond',data:D.bond.map(d=>d.y),borderColor:C.blue,borderWidth:2,pointRadius:0,tension:0.3,fill:false}]);areaLine('c-fx-ov',D.fx.map(d=>d.x),D.fx.map(d=>d.y),C.gold,C.goldA);}
function rAc(){const g=sortQ(D.gdp),gV=g.map(d=>d.y);barLine('c-gdp-na',g.map(d=>d.x),gV,ma(gV,4),C.red,C.blueL);const dom=sortQ(D.domdem),cons=sortQ(D.cons);lines('c-dom-na',dom.map(d=>d.x),[{label:'Dom Demand',data:dom.map(d=>d.y),borderColor:C.blue,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'Private Cons',data:cons.map(d=>d.y),borderColor:C.gold,borderWidth:2,pointRadius:0,tension:0.3,fill:false}]);const inv=sortQ(D.inv),gov=sortQ(D.govcon);lines('c-inv-na',inv.map(d=>d.x),[{label:'Fixed Inv',data:inv.map(d=>d.y),borderColor:C.blue,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'Govt Cons',data:gov.map(d=>d.y),borderColor:C.green,borderWidth:2,pointRadius:0,tension:0.3,fill:false}]);const aV=D.activity.map(d=>d.y);barLine('c-act-na',D.activity.map(d=>d.x),aV,ma(aV,3),C.red,C.blueL);}
function rPr(){const cV=D.cpi.map(d=>d.y);barLine('c-cpi-p',D.cpi.map(d=>d.x),cV,ma(cV,3),C.red,C.blueL);const cmV=D.cpi_mom.map(d=>d.y);barLine('c-cpimom-p',D.cpi_mom.map(d=>d.x),cmV,ma(cmV,3),C.orange,C.blue);lines('c-rb-p',D.rate.map(d=>d.x),[{label:'Policy',data:D.rate.map(d=>d.y),borderColor:C.red,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'10Y',data:D.bond.map(d=>d.y),borderColor:C.blue,borderWidth:2,pointRadius:0,tension:0.3,fill:false},{label:'Banking',data:D.banking.map(d=>d.y),borderColor:C.gold,borderWidth:1.5,pointRadius:0,tension:0.3,fill:false}]);const eV=D.cpi_eop.map(d=>d.y);barLine('c-cpieop-p',D.cpi_eop.map(d=>d.x),eV,ma(eV,3),C.red,C.blueL);}
function rEx(){const tb=sortQ(D.tb),el=document.getElementById('c-tb-ex');new Chart(el,{type:'bar',data:{labels:tb.map(d=>d.x),datasets:[{type:'bar',label:'Exports',data:tb.map(d=>d.exports),backgroundColor:C.blue+'aa',borderRadius:2},{type:'bar',label:'Imports',data:tb.map(d=>d.imports),backgroundColor:C.red+'aa',borderRadius:2},{type:'line',label:'Balance',data:tb.map(d=>d.balance),borderColor:C.green,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false}]},options:bO()});const ca=sortQ(D.ca);singleBar('c-ca-ex',ca.map(d=>d.x),ca.map(d=>d.y),v=>v>=0?C.green+'aa':C.red+'aa');const ex=sortQ(D.exports),im=sortQ(D.imports);lines('c-exim-ex',ex.map(d=>d.x),[{label:'Exports',data:ex.map(d=>d.y),borderColor:C.blue,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'Imports',data:im.map(d=>d.y),borderColor:C.red,borderWidth:2,pointRadius:0,tension:0.3,fill:false}]);areaLine('c-res-ex',D.intres.map(d=>d.x),D.intres.map(d=>d.y),C.green,C.greenA);}
function rLa(){areaLine('c-unemp-l',D.unemp.map(d=>d.x),D.unemp.map(d=>d.y),C.purple,C.purpleA);const rV=D.retail.map(d=>d.y);barLine('c-retail-l',D.retail.map(d=>d.x),rV,ma(rV,3),C.blue,C.blueL);const s=D.sec;stacked('c-sec-l',s.map(d=>d.x),[{label:'Agriculture',data:s.map(d=>d.agri),backgroundColor:C.blue+'cc',borderRadius:1},{label:'Mining',data:s.map(d=>d.mining),backgroundColor:C.red+'cc',borderRadius:1},{label:'Construction',data:s.map(d=>d.construction),backgroundColor:C.green+'cc',borderRadius:1},{label:'Services',data:s.map(d=>d.services),backgroundColor:C.gold+'cc',borderRadius:1},{label:'Manufacturing',data:s.map(d=>d.manufacturing),backgroundColor:C.purple+'cc',borderRadius:1}]);const mV=D.manuf_m.map(d=>d.y);barLine('c-manuf-l',D.manuf_m.map(d=>d.x),mV,ma(mV,3),C.navy,C.blueL);}
function rRa(){areaLine('c-rate-r',D.rate.map(d=>d.x),D.rate.map(d=>d.y),C.red,C.redA);areaLine('c-bond-r',D.bond.map(d=>d.x),D.bond.map(d=>d.y),C.blue,C.blueA);areaLine('c-fx-r',D.fx.map(d=>d.x),D.fx.map(d=>d.y),C.gold,C.goldA);lines('c-all-r',D.rate.map(d=>d.x),[{label:'Policy',data:D.rate.map(d=>d.y),borderColor:C.red,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'10Y Bond',data:D.bond.map(d=>d.y),borderColor:C.blue,borderWidth:2,pointRadius:0,tension:0.3,fill:false},{label:'Banking',data:D.banking.map(d=>d.y),borderColor:C.gold,borderWidth:1.5,pointRadius:0,tension:0.3,fill:false}]);}
function rFi(){areaLine('c-debt-f',D.pubdebt.map(d=>d.x),D.pubdebt.map(d=>d.y),C.red,C.redA);const gV=sortQ(D.govcon);barLine('c-gov-f',gV.map(d=>d.x),gV.map(d=>d.y),ma(gV.map(d=>d.y),4),C.blue,C.blueL);lines('c-inv-f',D.priv_inv.map(d=>d.x),[{label:'Private Inv',data:D.priv_inv.map(d=>d.y),borderColor:C.blue,borderWidth:2.5,pointRadius:0,tension:0.3,fill:false},{label:'Public Inv',data:D.pub_inv.map(d=>d.y),borderColor:C.red,borderWidth:2,pointRadius:0,tension:0.3,fill:false}]);areaLine('c-res-f',D.intres.map(d=>d.x),D.intres.map(d=>d.y),C.green,C.greenA);}
rOv();
</script>
</body>
</html>"""

# Inject data into template
html = HTML_TEMPLATE.replace('const D=__DATA__;', f'const D={DATA_JSON};')

with open(OUTPUT_HTML, 'w', encoding='utf-8') as f:
    f.write(html)

print(f"\nSUCCESS! Dashboard saved as: {OUTPUT_HTML}")
print(f"Open it in Chrome to see your updated dashboard.")
print(f"\nNext quarter: update the Excel, run this script again, done!")
input("\nPress Enter to close...")
