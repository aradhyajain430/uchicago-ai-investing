"""Final valuation engine: 3–6 month fear-unwind thesis, independently tested DCF."""
from pathlib import Path
from datetime import date
import json
import statistics

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'deliverables'
OUT.mkdir(exist_ok=True)
PRICE=893.61
PRE_EVENT=927.03
EVENT=835.03
ANALYST_LOW,ANALYST_AVG,ANALYST_HIGH=820,1149,1400
STUB=(date(2027,6,26)-date(2026,9,18)).days/365
YEARS=list(range(2027,2037))
CAP=dict(cash=2043.5,investments=694.9,other_debt=92.8,common=88.6,preferred=2.9,
         rsu=2.0,psu=.8,options=.2,strike=8.1,nwc=553.4,minority=0,preferred_claim=0)
NOTES=[dict(year=2026,p=54.8,k=99.29),dict(year=2028,p=179.6,k=131.03),
       dict(year=2029,p=54.9,k=69.54),dict(year=2032,p=1265,k=187.77)]
TSM=CAP['options']*max(1-CAP['strike']/PRICE,0)
CORE=CAP['common']+CAP['preferred']+CAP['rsu']+CAP['psu']+TSM
ALL_SHARES=CORE+sum(n['p']/n['k'] for n in NOTES)
NET_CASH=CAP['cash']+CAP['investments']-CAP['other_debt']
DEBT=CAP['other_debt']+sum(n['p'] for n in NOTES)
PEERS=[dict(ticker='COHR',price=295.98,shares=195.83,debt=3550,cash=1990,beta=2.10,
            ebitda=1420,adj_margin=.218,forward_pe=31.43),
       dict(ticker='MRVL',price=240.76,shares=876.93,debt=5290,cash=3930,beta=2.25,
            ebitda=2850,adj_margin=.366,forward_pe=44.12),
       dict(ticker='FN',price=380.92,shares=35.83,debt=4.01,cash=875.06,beta=1.21,
            ebitda=530.73,adj_margin=.0998,forward_pe=20.71)]
TAX=.20
for p in PEERS:
    p['market_equity']=p['price']*p['shares']
    p['de']=p['debt']/p['market_equity']
    p['unlevered_beta']=p['beta']/(1+(1-.25)*p['de'])
    p['ev']=p['market_equity']+p['debt']-p['cash']
    p['ev_ebitda']=p['ev']/p['ebitda']
RF=.0494
ERP=.045
ERP_OBS=.0414
PEER_BETA=statistics.median(p['unlevered_beta'] for p in PEERS)
# Equity-equivalent notes are excluded from debt weights to match the valuation bridge.
# Sensitivity separately exposes using all principal as debt. Preferred participates in equity.
MARKET_EQUITY=ALL_SHARES*PRICE
WACC_DEBT=CAP['other_debt']
DE=WACC_DEBT/MARKET_EQUITY
BETA=PEER_BETA*(1+(1-TAX)*DE)
KE=RF+BETA*ERP
KD=RF+.025
DW=WACC_DEBT/(WACC_DEBT+MARKET_EQUITY)
WACC=KE*(1-DW)+KD*(1-TAX)*DW
OWN_BETA=1.54
OWN_WACC=(RF+OWN_BETA*ERP)*(1-DW)+KD*(1-TAX)*DW
G=.03
EXIT=16.0

SOURCES=[
 ('S1','FY2026 results / Q1 guide','2026-08-11','lite_release','FY26 revenue $3,014m; Q1 guide $1,225–1,275m, adjusted margin 39.5–40.5%; accounting reconciliation.'),
 ('S2','FY2026 10-K','2026-08-17','lite_fy26_10k','Cash, NWC, PPE, debt, preferred, options and RSU/PSU; June 27 snapshot.'),
 ('S3','LITE price history','2026-09-17 close','lite_price','Sept 11 $927.03; Sept 14 $835.03; Sept 17 $893.61. Event attribution remains a thesis.'),
 ('S4','LITE public estimates','accessed 2026-09-18','lite_forecast','FY27 revenue $6.32bn; FY28 $9.61bn; FY27 adjusted EPS $21.67; analyst range $820–1,400, avg $1,149. Secondary vendor.'),
 ('S5','FY2025 results / former segments','2025-08','lite_fy25_release','Cloud & Networking $1,410.8m; Industrial Tech $234.2m. FY26 no longer reports these segments.'),
 ('S6','Treasury yield curve','2026-09-17','treasury_2026','10-year Treasury 4.94%; dated nominal USD risk-free proxy.'),
 ('S7','Damodaran implied ERP','2026-09-01','damodaran_erp','Observed trailing adjusted payout ERP 4.14%; model deliberately uses 4.50%.'),
 ('S8','Coherent statistics','accessed 2026-09-18','cohr_stats','Beta 2.10; debt $3.55bn; cash $1.99bn; EBITDA $1.42bn. Price ratios re-anchored to Sept 17.'),
 ('S9','Marvell statistics','accessed 2026-09-18','mrvl_stats','Beta 2.25; debt $5.29bn; cash $3.93bn; EBITDA $2.85bn; forward P/E 44.12x.'),
 ('S10','Fabrinet statistics','accessed 2026-09-18','fn_stats','Beta 1.21; debt $4.01m; cash $875.06m; EBITDA $530.73m; forward P/E re-anchored to Sept 17.'),
 ('S11','Coherent FY26 results','2026-08-12','cohr_sec_release','Q4 adjusted operating margin 21.8%; mature margin cross-check, not like-for-like economics.'),
 ('S12','Marvell Q2 FY27 results','2026-08-27','mrvl_q2_fy27','Q2 adjusted operating margin 36.6%; higher-margin fabless comparison.'),
 ('S13','Meta Q2 CY26 results','2026-07-29','meta_q2_2026','2026 capex including finance-lease principal $130–145bn. Broader than optics spending.'),
 ('S14','Amazon Q2 CY26 results','2026-07-30','amazon_q2_2026','TTM net PPE purchases $169.007bn, +64%; infrastructure evidence, not an optics TAM.'),
 ('S15','LITE statistics','accessed 2026-09-18','lite_stats','Own observed beta 1.54; lower-WACC cross-check, not the primary beta.'),
 ('S16','Amodei essay','September 2026','amodei_essay','Team thesis: fear discount unwinds over 3–6 months if execution remains intact. Causality unproven.'),
]
manifest={x['id']:x for x in json.loads((ROOT/'research/sources/manifest.json').read_text(encoding='utf-8'))}
SOURCES=[dict(id=a,title=b,date=c,key=d,usage=e,url=manifest.get(d,{}).get('url',''),
              path=manifest.get(d,{}).get('path','')) for a,b,c,d,e in SOURCES]

IND26=234.2*1.05 # estimated former end-market allocation, not reported segment revenue
IND=[IND26*1.06**(i+1) for i in range(5)]
for growth in [.05,.04,.03,.03,.03]:IND.append(IND[-1]*(1+growth))
BASE_TOTAL=[6320,9610]
for growth in [.30,.22,.15,.12,.09,.07,.05,.03]:BASE_TOTAL.append(BASE_TOTAL[-1]*(1+growth))
BASE_CLOUD=[r-i for r,i in zip(BASE_TOTAL,IND)]
EML=[.01,.015,.02,.025,.03,.03,.025,.02,.015,.01]
MIX=[.02,.03,.04,.045,.05,.05,.045,.04,.035,.03]
OCS=[150,300,550,800,1100,1250,1350,1400,1442,1485.26]
SOURCING=[.005,.01,.015,.015,.015,.01,.0075,.005,.0025,0]
BASE_MARGIN=[.4025,.415,.41,.40,.38,.35,.32,.295,.27,.25]
COMMON=dict(sbc=[.045]*10,tax=TAX,dep=[.04,.045,.05,.055,.055,.055,.055,.055,.055,.055],
            capex=[.13,.11,.095,.08,.075,.07,.065,.06,.06,.06],wacc=WACC,g=G,exit=EXIT)
CASES={
 'Bull':dict(**COMMON,cloud=[c*(1+e+m)+o for c,e,m,o in zip(BASE_CLOUD,EML,MIX,OCS)],
             industrial=IND,margin=[m+s for m,s in zip(BASE_MARGIN,SOURCING)],nwc=.18,
             description='Team conviction: no near-term spending pause; LITE share/mix/OCS wins and internal sourcing improve cash generation.'),
 'Base':dict(**COMMON,cloud=BASE_CLOUD,industrial=IND,margin=BASE_MARGIN,nwc=.18,
             description='Published FY27/FY28 revenue anchors; subsequent datacom growth and margins fade. No fear-induced collapse.'),
 'Bear':dict(**COMMON,cloud=[5340],industrial=[260],
             margin=[.35,.34,.33,.32,.31,.29,.275,.26,.25,.25],nwc=.22,
             description='Orders slow, working capital stretches, but cash flows remain positive and principal is funded. Equity downside is not a solvency prediction.'),
}
for g in [.20,.15,.10,.08,.06,.05,.04,.03,.03]:CASES['Bear']['cloud'].append(CASES['Bear']['cloud'][-1]*(1+g))
for g in [0,.02,.03,.03,.03,.03,.03,.03,.03]:CASES['Bear']['industrial'].append(CASES['Bear']['industrial'][-1]*(1+g))

def bridge(ev):
    ordered=sorted(NOTES,key=lambda n:n['k'])
    for n in range(5):
        active=ordered[:n]; inactive=ordered[n:]
        debt=sum(x['p'] for x in inactive)
        shares=CORE+sum(x['p']/x['k'] for x in active)
        eq=ev+NET_CASH-CAP['minority']-CAP['preferred_claim']-debt
        value=eq/shares
        if n==4 or value<ordered[n]['k']:
            converted_principal=sum(x['p'] for x in active)
            net_shares=shares-converted_principal/value
            cash_eq=eq-converted_principal
            return dict(value=value,equity=eq,shares=shares,unconverted_debt=debt,
                        equity_after_all_principal=cash_eq,net_shares=net_shares,
                        preferred_claim=CAP['preferred_claim'],minority=CAP['minority'])
    raise AssertionError('No consistent conversion regime')

def evaluate(s,wacc=None,g=None,exit_multiple=None,reverse_growth=None,reverse_margin=None):
    w=s['wacc'] if wacc is None else wacc
    tg=s['g'] if g is None else g
    mult=s['exit'] if exit_multiple is None else exit_multiple
    cloud=list(s['cloud']);margin=list(s['margin'])
    if reverse_growth is not None:
        for i in range(1,10):
            rate=reverse_growth if i<=4 else reverse_growth+(tg-reverse_growth)*(i-4)/5
            cloud[i]=cloud[i-1]*(1+rate)
    if reverse_margin is not None:
        for i in range(1,10):
            margin[i]=reverse_margin if i<=4 else reverse_margin+(.25-reverse_margin)*(i-4)/5
    prev_nwc=CAP['nwc'];rows=[]
    for i,year in enumerate(YEARS):
        r=cloud[i]+s['industrial'][i]
        adj=r*margin[i];sbc=r*s['sbc'][i];ebit=adj-sbc
        nopat=ebit*(1-s['tax']);depr=r*s['dep'][i];capex=r*s['capex'][i]
        nwc=r*s['nwc'];dnwc=nwc-prev_nwc
        fcf=nopat+depr-capex-dnwc
        fraction=STUB if i==0 else 1;time=STUB+i
        rows.append(dict(year=year,cloud=cloud[i],industrial=s['industrial'][i],revenue=r,
             margin=margin[i],adj_ebit=adj,sbc=sbc,ebit=ebit,tax=ebit*s['tax'],nopat=nopat,
             da=depr,capex=capex,nwc=nwc,dnwc=dnwc,fcf=fcf,ebitda=ebit+depr,
             fraction=fraction,time=time,pv=fcf*fraction/(1+w)**time))
        prev_nwc=nwc
    last=rows[-1];tr=last['revenue']*(1+tg)
    tf=tr*((last['margin']-s['sbc'][-1])*(1-s['tax'])+s['dep'][-1]-s['capex'][-1])-(tr*s['nwc']-last['nwc'])
    pv=sum(r['pv'] for r in rows)
    tv=tf/(w-tg);etv=last['ebitda']*mult
    tvpv=tv/(1+w)**last['time'];epv=etv/(1+w)**last['time']
    ev=pv+tvpv;eev=pv+epv
    return dict(rows=rows,terminal_revenue=tr,terminal_fcf=tf,pv_explicit=pv,tv=tv,tvpv=tvpv,ev=ev,
       gordon=bridge(ev),exit_ev=eev,exit_tv=etv,exit_pv=epv,exit=bridge(eev),
       tv_share=tvpv/ev,exit_tv_share=epv/eev,wacc=w,g=tg,exit_multiple=mult,
       implied_terminal_multiple=tv/last['ebitda'])

RESULTS={n:evaluate(s) for n,s in CASES.items()}
def solve(fn,target,lo,hi):
    assert fn(lo)<=target<=fn(hi),(fn(lo),target,fn(hi))
    for _ in range(100):
        mid=(lo+hi)/2
        if fn(mid)<target:lo=mid
        else:hi=mid
    return (lo+hi)/2
REV_GROWTH=solve(lambda x:evaluate(CASES['Bull'],reverse_growth=x)['gordon']['value'],PRICE,0,2)
REV_MARGIN=solve(lambda x:evaluate(CASES['Bull'],reverse_margin=x)['gordon']['value'],PRICE,0,3)
REVERSE_G=evaluate(CASES['Bull'],reverse_growth=REV_GROWTH)
REVERSE_M=evaluate(CASES['Bull'],reverse_margin=REV_MARGIN)
# Price-calibrated risk premium: a diagnostic, never the independently sourced WACC.
def implied_discount(price):
    lo,hi=G+.0001,.50
    for _ in range(100):
        mid=(lo+hi)/2
        if evaluate(CASES['Bull'],wacc=mid)['gordon']['value']>price:lo=mid
        else:hi=mid
    return (lo+hi)/2
IMPLIED_CURRENT=implied_discount(PRICE)
IMPLIED_PRE=implied_discount(PRE_EVENT)
IMPLIED_SHOCK=implied_discount(EVENT)
IMPLIED_ANALYST=implied_discount(ANALYST_AVG)
BEAR_LIQ=[];cash=CAP['cash']+CAP['investments']
principal={2027:54.8+52.4,2028:179.6+10.8,2029:54.9+29.6,2032:1265}
remaining=DEBT
for r in RESULTS['Bear']['rows']:
    repay=principal.get(r['year'],0)
    interest=remaining*KD*(1-TAX)*r['fraction']
    cash+=r['fcf']*r['fraction']-repay-interest
    remaining-=repay
    BEAR_LIQ.append(dict(year=r['year'],repayment=repay,interest=interest,cash=cash,debt=remaining))

ASSUMPTIONS=[
 ('Horizon','3–6 months','Next 1–2 earnings reports test execution; DCF duration is separate.','Team thesis'),
 ('Forecast length','5-year ramp + 5-year fade','FY27–31 captures ramp; FY32–36 fades growth to 3% and adjusted margin to 25%.','Team model'),
 ('Cloud / Industrial history','FY25: 1,410.8 / 234.2','FY26 Industrial = FY25 × 1.05 estimate; Cloud is total less Industrial. Not reported FY26 segments.','S2 / S5 + estimate'),
 ('Base revenue FY27 / FY28','6,320 / 9,610','Public vendor consensus; later growth 30%,22%,15%,12%,9%,7%,5%,3%. No later consensus claim.','S4 + team fade'),
 ('Bull EML share effect','+1% → +3% → +1%','Revenue uplift on baseline Cloud, a proxy for share/content gains; not a claimed measured EML market share.','Team assumption'),
 ('Bull 1.6T mix effect','+2% → +5% → +3%','Separate external module mix premium; excludes internal component transfers and avoids intercompany double count.','Team assumption; S1 direction'),
 ('Bull OCS increment','150 → 1,100 → 1,485','USD m added over the baseline Cloud path. OCS is part of Cloud; no claim this is disclosed backlog/revenue.','Team assumption; S1 direction'),
 ('Internal sourcing margin','+50 → +150 → 0 bps','Increment to adjusted margin only; no extra revenue for internal laser transfers.','Team assumption'),
 ('Adjusted EBIT margin','~41% → 42.5% peak → 25%','Mature terminal 25% sits between COHR 21.8% and MRVL 36.6%; FN ~10% GAAP margin provides manufacturing context.','S10/S11/S12; team terminal choice'),
 ('Recurring SBC + payroll','4.5% revenue','Economic expense retained. Existing awards included separately; no recurring share-count escalator.','S1 FY26 191.3/3,014=6.35%; assumed leverage'),
 ('Tax','20%','Higher than FY26 adjusted 16.5%; common across scenarios.','S1 + conservative estimate'),
 ('D&A add-back','4% → 5.5% revenue','Depreciation only. Acquired amortization already removed from adjusted EBIT; no second add-back.','S1/S2 + estimate'),
 ('Capex','13% → 6% revenue','Gross capacity investment; terminal capex exceeds depreciation. FY26 cash capex 451.3m.','S2 + estimate'),
 ('Operating NWC','18%; bear 22%','Broad operating-current-asset less operating-current-liability proxy; opening 553.4m.','S2 + estimate'),
 ('WACC',f'{WACC:.2%}','4.94% Treasury + relevered peer beta × 4.5% ERP; marginal debt cost 7.44%; market weights.','S6–S10'),
 ('Terminal growth','3.0%','Below roughly 4% nominal GDP benchmark (Fed June 2026: 2% real + 2% inflation); no perpetual datacom hypergrowth.','S17 Fed benchmark / team'),
 ('Exit multiple','16× economic EBITDA','Mature terminal cross-check, below current 24–74× peer TTM range; not an observed future market multiple.','S8–S10 + judgment'),
 ('Capital snapshot','June 27, 2026','Common 88.6m; preferred 2.9m; RSU 2m, PSU 0.8m; options 0.2m @8.10 treasury method.','S2'),
 ('Stub convention',f'{STUB:.3f} years','Uniform remaining FY27 cash generation; June cash held fixed. Approximation disclosed.','Team convention'),
]

def validate():
    for name,r in RESULTS.items():
        assert r['wacc']>r['g']
        assert abs(r['rows'][-1]['margin']-.25)<1e-9
        for x in r['rows']:
            assert abs(x['fcf']-(x['nopat']+x['da']-x['capex']-x['dnwc']))<1e-8
        for method in ['gordon','exit']:
            b=r[method]
            assert abs(b['equity']/b['shares']-b['value'])<1e-8
            assert abs(b['equity_after_all_principal']/b['net_shares']-b['value'])<1e-8
        assert evaluate(CASES[name],wacc=WACC+.01)['gordon']['value']<r['gordon']['value']
    assert min(r['cash'] for r in BEAR_LIQ)>0
    assert abs(REVERSE_G['gordon']['value']-PRICE)<1e-7
    assert abs(REVERSE_M['gordon']['value']-PRICE)<1e-7
    return 'PASS'

if __name__=='__main__':
    validate()
    print(json.dumps({'wacc':WACC,'own_beta_wacc':OWN_WACC,'beta':BETA,
      'cases':{n:dict(gordon=r['gordon']['value'],exit=r['exit']['value'],fy27=r['rows'][0]['revenue'],fy36=r['rows'][-1]['revenue']) for n,r in RESULTS.items()},
      'reverse_growth':REV_GROWTH,'reverse_margin':REV_MARGIN,'minimum_bear_cash':min(r['cash'] for r in BEAR_LIQ),
      'fear_remaining_upside':PRE_EVENT/PRICE-1,'implied_wacc_now':IMPLIED_CURRENT,'implied_wacc_pre':IMPLIED_PRE},indent=2))
