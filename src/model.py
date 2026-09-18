"""Transparent operating model and dated DCF; USD millions except per share."""
from __future__ import annotations
from datetime import date
from copy import deepcopy
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASOF = date(2026, 9, 18)
ANCHOR = date(2026, 6, 26)
PRICE = 380.92
SHARES = 36.3
CASH = 875.055
NEW_DEBT = 75.0
NONOP_ASSETS = 89.103
MIN_CASH_PCT = .02
TAX_PAYABLE_SETTLEMENT = 57.447
WACC = .10
G = .03
YEARS = list(range(2027, 2037))
# June 30 convention is explicitly used for projected fiscal year ends.
PERIOD_DAYS = (date(2027,6,30) - ANCHOR).days
ELAPSED = (ASOF - ANCHOR).days / PERIOD_DAYS

HIST = {
  2024: dict(dc=1529.693, comm=767.368, other=585.906, rev=2882.967, cogs=2526.849, gp=356.118, sga=78.481, restructuring=.032, ebit=277.605, interest=33.204, pbt=311.354, tax=15.173, ni=296.181, da=49.017, cfo=413.146, capex=47.528, accrued_capex=49.270, sbc=28.374, diluted_shares=36.564, gaap_eps=8.10),
  2025: dict(dc=1579.915, comm=1049.397, other=790.015, rev=3419.327, cogs=3005.978, gp=413.349, sga=87.466, restructuring=1.436, ebit=324.447, interest=40.162, pbt=355.180, tax=22.653, ni=332.527, da=53.433, cfo=328.365, capex=121.078, accrued_capex=130.658, sbc=33.004, diluted_shares=36.267, gaap_eps=9.17, cash=934.244, ar=758.894, inv=581.015, ap=637.417, prepaid=38.476, other_ca=116.210, payroll=24.566, accrued=30.630, other_payable=66.717, sev_current=0, ppe=380.640, tax_payable=7.939),
  2026: dict(dc=2225.081, comm=1546.403, other=869.613, rev=4641.097, cogs=4084.586, gp=556.511, sga=93.507, restructuring=.117, ebit=462.887, interest=32.418, pbt=555.113, tax=82.086, ni=473.027, da=68.371, cfo=256.725, capex=252.503, accrued_capex=298.921, sbc=34.630, diluted_shares=36.252, gaap_eps=13.05, cash=CASH, ar=1017.933, inv=1021.235, ap=1005.761, prepaid=49.915, other_ca=195.571, payroll=29.850, accrued=56.549, other_payable=157.470, sev_current=2.319, ppe=615.067, tax_payable=63.469),
}
for h in HIST.values():
    h['fcf'] = h['cfo'] - h['capex']
    h['gm'] = h['gp']/h['rev']
    h['om'] = h['ebit']/h['rev']
    if 'ar' in h:
        h['other_nwc'] = h['prepaid']+h['other_ca']-h['payroll']-h['accrued']-h['other_payable']-h['sev_current']
        h['nwc'] = h['ar']+h['inv']-h['ap']+h['other_nwc']
        h['dso'] = h['ar']/h['rev']*365
        h['dio'] = h['inv']/h['cogs']*365
        h['dpo'] = h['ap']/h['cogs']*365

CASES = {
 'Bear': dict(prob=.20, dc=[.22,.10,.08,.07,.06], comm=[.15,.06,.05,.04,.03], other=[.05,.03,.03,.03,.03], gm=[.115,.113,.114,.115,.116], sga=[110,116,122,128,134], da=[90,115,130,145,155], capex=[400,375,330,290,270], dso=[82,83,81,79,78], dio=[100,102,98,94,90], dpo=[83,82,82,82,82], tax=.19, roic=.15, pe=18),
 'Base': dict(prob=.60, dc=[.38,.25,.20,.15,.10], comm=[.27,.16,.12,.08,.05], other=[.12,.07,.06,.05,.04], gm=[.120,.120,.121,.122,.122], sga=[108,119,131,143,155], da=[90,115,135,150,165], capex=[350,350,330,320,315], dso=[77,76,75,74,73], dio=[92,90,87,84,82], dpo=[86,85,84,83,82], tax=.175, roic=.20, pe=20),
 'Bull': dict(prob=.20, dc=[.50,.40,.30,.24,.18], comm=[.35,.25,.20,.15,.10], other=[.15,.10,.08,.06,.05], gm=[.124,.128,.131,.133,.133], sga=[112,128,145,164,181], da=[95,120,150,175,200], capex=[380,400,430,450,460], dso=[72,70,69,68,67], dio=[85,80,76,73,70], dpo=[86,85,84,83,82], tax=.15, roic=.25, pe=26),
}

def case_assumptions(name):
    s=deepcopy(CASES[name])
    for key in ['dc','comm','other']:
        last=s[key][-1]
        s[key] += [G+(last-G)*(4-k)/5 for k in range(5)]
    s['gm'] += [s['gm'][-1]]*5
    s['sga'] += [s['sga'][-1]*(1.05**k) for k in range(1,6)]
    s['da'] += [0]*5 # Years 6-10 use a percentage of revenue in forecast.
    s['capex'] += [0]*5
    for key in ['dso','dio','dpo']:
        s[key] += [s[key][-1]]*5
    return s

def forecast(name, growth_shift=0.0):
    s=case_assumptions(name)
    old=HIST[2026]
    oldrev=old['rev']; oldwc=old['nwc']
    seg={k:old[k] for k in ['dc','comm','other']}
    rows=[]
    for i,fy in enumerate(YEARS):
        for k in seg: seg[k] *= 1+s[k][i]+(growth_shift if i<5 else growth_shift*(9-i)/5)
        rev=sum(seg.values()); gm=s['gm'][i]
        gp=rev*gm; cogs=rev-gp; sga=s['sga'][i]; ebit=gp-sga
        tax=ebit*s['tax']; nopat=ebit-tax
        da=s['da'][i] if i<5 else rev*.018
        capex=s['capex'][i] if i<5 else max(rev*.024, s['capex'][4])
        ar=rev/365*s['dso'][i]; inv=cogs/365*s['dio'][i]; ap=cogs/365*s['dpo'][i]
        otherwc=rev*(HIST[2026]['other_nwc']/HIST[2026]['rev'])
        wc=ar+inv-ap+otherwc
        min_cash=rev*MIN_CASH_PCT; old_min_cash=oldrev*MIN_CASH_PCT
        dwc=wc-oldwc; dmin=min_cash-old_min_cash
        settlement=TAX_PAYABLE_SETTLEMENT if i==0 else 0.0
        fcff=nopat+da-capex-dwc-dmin-settlement
        netinterest=(CASH*.035-NEW_DEBT*.055)*(1-s['tax'])
        eps=(nopat+netinterest)/SHARES
        adj_eps=eps+1.24+(1.60 if i==0 else 1.60*rev/rows[0]['rev'])
        t=(date(fy,6,30)-ASOF).days/365
        rows.append(dict(fy=fy, **seg, rev=rev, growth=rev/oldrev-1, gm=gm, gp=gp,cogs=cogs,sga=sga,ebit=ebit,om=ebit/rev,tax_rate=s['tax'],tax=tax,nopat=nopat,da=da,capex=capex,ar=ar,inv=inv,ap=ap,other_nwc=otherwc,nwc=wc,dnwc=dwc,min_cash=min_cash,dmin_cash=dmin,tax_settlement=settlement,fcff=fcff,eps=eps,adj_eps=adj_eps,t=t,dso=s['dso'][i],dio=s['dio'][i],dpo=s['dpo'][i]))
        oldwc=wc; oldrev=rev
    return rows

def valuation(name, wacc=WACC, g=G, growth_shift=0.0):
    rows=forecast(name,growth_shift); roic=CASES[name]['roic']
    # First fiscal year's cash flow and cash roll are straight-line estimates.
    # The loan adds cash and debt equally; expenditures remain in forecast capex.
    cash_roll=rows[0]['fcff']*ELAPSED
    excess_cash=CASH+NEW_DEBT+cash_roll-HIST[2026]['rev']*MIN_CASH_PCT
    debt=NEW_DEBT
    nonop=NONOP_ASSETS*.50 # Explicit 50% haircut to nonmarketable equity holdings.
    stub=1-ELAPSED
    pvs=[r['fcff']*(stub if i==0 else 1)/(1+wacc)**r['t'] for i,r in enumerate(rows)]
    terminal_nopat=rows[-1]['nopat']*(1+g)
    terminal_reinvestment=terminal_nopat*g/roic
    terminal_fcf=terminal_nopat-terminal_reinvestment
    tv=terminal_fcf/(wacc-g)
    pv_tv=tv/(1+wacc)**rows[-1]['t']
    ev=sum(pvs)+pv_tv
    equity=ev+excess_cash-debt+nonop
    fair=equity/SHARES
    # Roll EV to exit and subtract FCFF delivered before exit; add cash retained.
    exits={}
    for months in [3,6,12]:
        h=months/12
        before=[(i,r) for i,r in enumerate(rows) if r['t']<=h]
        delivered=sum(r['fcff']*(stub if i==0 else 1) for i,r in before)
        ev_exit=ev*(1+wacc)**h-sum(r['fcff']*(stub if i==0 else 1)*(1+wacc)**(h-r['t']) for i,r in before)
        # Assume zero interest on incremental retained FCFF and debt constant.
        cash_exit=excess_cash+delivered
        exits[months]=(ev_exit+cash_exit-debt+nonop)/SHARES
    fwd_eps=rows[1]['eps'] # FY2028 as a clearly labelled exit cross-check.
    return dict(name=name, rows=rows,wacc=wacc,g=g,roic=roic,pvs=pvs,pv_tv=pv_tv,terminal_nopat=terminal_nopat,terminal_reinvestment=terminal_reinvestment,terminal_fcf=terminal_fcf,terminal_value=tv,ev=ev,excess_cash=excess_cash,cash_roll=cash_roll,debt=debt,nonop=nonop,equity=equity,fair_value=fair,exits=exits,tv_share=pv_tv/ev,upside=fair/PRICE-1,exit_return=exits[12]/PRICE-1,eps_crosscheck=fwd_eps*CASES[name]['pe'],prob=CASES[name]['prob'])

def reverse_growth():
    lo,hi=-.10,.40
    for _ in range(70):
        mid=(lo+hi)/2
        if valuation('Base',growth_shift=mid)['fair_value']<PRICE: lo=mid
        else: hi=mid
    return (lo+hi)/2

def outputs():
    cases={n:valuation(n) for n in CASES}
    shift=reverse_growth()
    return dict(asof=ASOF.isoformat(),price_date='2026-09-17 close',price=PRICE,shares=SHARES,anchor=ANCHOR.isoformat(),elapsed=ELAPSED,cases=cases,reverse_growth_shift=shift,reverse_rows=forecast('Base',shift),expected_exit=sum(v['exits'][12]*v['prob'] for v in cases.values()))

if __name__=='__main__':
    data=outputs(); out=ROOT/'research'/'model_outputs.json'; out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(data,indent=2),encoding='utf-8')
    for n,c in data['cases'].items():
        print(n, 'FY27 revenue',round(c['rows'][0]['rev'],1),'FCFF',round(c['rows'][0]['fcff'],1),'DCF now',round(c['fair_value'],2),'12m',round(c['exits'][12],2),'return',round(c['exit_return']*100,1),'FY28 EPS',round(c['rows'][1]['eps'],2),'P/E check',round(c['eps_crosscheck'],1))
    print('Market price requires extra growth pp',round(data['reverse_growth_shift']*100,2))
    print('Expected exit',round(data['expected_exit'],2))
