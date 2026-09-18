"""Archive public sources used in the competition research; no credentials."""
from pathlib import Path
import json
import sys
from concurrent.futures import ThreadPoolExecutor
import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'research' / 'sources'
OUT.mkdir(parents=True, exist_ok=True)
SOURCES = {
    'lite_forecast': 'https://stockanalysis.com/stocks/lite/forecast/',
    'lite_stats': 'https://stockanalysis.com/stocks/lite/statistics/',
    'cohr_stats': 'https://stockanalysis.com/stocks/cohr/statistics/',
    'mrvl_stats': 'https://stockanalysis.com/stocks/mrvl/statistics/',
    'fn_stats': 'https://stockanalysis.com/stocks/fn/statistics/',
    'lite_fy25_release': 'https://investor.lumentum.com/financial-news-releases/news-details/2025/Lumentum-Announces-Fourth-Quarter-and-Full-Fiscal-Year-2025-Results/',
    'treasury_2026': 'https://home.treasury.gov/resource-center/data-chart-center/interest-rates/TextView?type=daily_treasury_yield_curve&field_tdr_date_value=2026',
    'damodaran_erp': 'https://pages.stern.nyu.edu/adamodar/New_Home_Page/home.htm',
    'meta_q2_2026': 'https://investor.atmeta.com/investor-news/press-release-details/2026/Meta-Reports-Second-Quarter-2026-Results/default.aspx',
    'amazon_q2_2026': 'https://ir.aboutamazon.com/news-release/news-release-details/2026/Amazon-com-Announces-Second-Quarter-Results/default.aspx',
    'cohr_sec_release': 'https://www.sec.gov/Archives/edgar/data/820318/000119312526346860/d128030dex991.htm',
    'mrvl_q2_fy27': 'https://investor.marvell.com/news-events/press-releases/detail/1031/marvell-technology-inc-reports-second-quarter-of-fiscal-year-2027-financial-results',
    'lite_fy26_10k': 'https://www.sec.gov/Archives/edgar/data/1633978/000162828026057358/lite-20260627.htm',
    'lite_price': 'https://stockanalysis.com/stocks/lite/history/',
    'cohr_fy26_10k': 'https://www.sec.gov/Archives/edgar/data/820318/000082031826000020/iivi-20260630.htm',
    'cohr_full_release': 'https://ir.coherent.com/news-releases/news-release-details/coherent-corp-reports-fourth-quarter-and-full-year-fiscal-2026',
    'cohr_transcript': 'https://www.fool.com/earnings/call-transcripts/2026/08/19/coherent-cohr-q4-2026-earnings-call-transcript/',
    'cohr_price': 'https://stockanalysis.com/stocks/cohr/history/',
    'fn_fy26_release': 'https://investor.fabrinet.com/node/13666',
    'fn_fy26_10k': 'https://www.sec.gov/Archives/edgar/data/1408710/000140871026000028/fn-20260626.htm',
    'fn_ir': 'https://investor.fabrinet.com/',
    'fn_price': 'https://stockanalysis.com/stocks/fn/history/',
    'fn_forecast': 'https://stockanalysis.com/stocks/fn/forecast/',
    'lite_release': 'https://investor.lumentum.com/financial-news-releases/news-details/2026/Lumentum-Announces-Fourth-Quarter-and-Full-Fiscal-Year-2026-Results/default.aspx',
    'cohr_release': 'https://www.coherent.com/news/press-releases/fourth-quarter-and-fiscal-year-2026-results',
    'cohr_presentation': 'https://www.coherent.com/content/dam/coherent/site/en/documents/investors/investor-presentations/2026/august-12/investor-presentation-20260812.pdf',
    'amodei_essay': 'https://darioamodei.com/post/we-must-pace-the-frontier',
}

def fetch(item):
    key, url = item
    try:
        r = requests.get(url, timeout=35, headers={'User-Agent': 'InvestmentCompetitionResearch research@example.com'})
        r.raise_for_status()
        suffix = '.pdf' if r.content[:4] == b'%PDF' else '.html'
        p = OUT / (key + suffix)
        p.write_bytes(r.content)
        if suffix == '.html':
            soup = BeautifulSoup(r.content, 'html.parser')
            for tag in soup(['script', 'style']): tag.decompose()
            (OUT / (key + '.txt')).write_text(soup.get_text('\n', strip=True), encoding='utf-8')
            tables = [[[c.get_text(' ', strip=True) for c in tr.find_all(['td','th'])] for tr in t.find_all('tr')] for t in soup.find_all('table')]
            (OUT / (key + '_tables.json')).write_text(json.dumps(tables, indent=2), encoding='utf-8')
            links = [{'text':a.get_text(' ', strip=True), 'href':a.get('href')} for a in soup.find_all('a', href=True)]
            (OUT / (key + '_links.json')).write_text(json.dumps(links, indent=2), encoding='utf-8')
        return {'id':key, 'url':url, 'status':'archived', 'bytes':len(r.content), 'path':str(p.relative_to(ROOT))}
    except Exception as e:
        return {'id':key, 'url':url, 'status':'failed', 'error':str(e)}

if __name__ == '__main__':
    selected={k:v for k,v in SOURCES.items() if not sys.argv[1:] or k in sys.argv[1:]}
    with ThreadPoolExecutor(max_workers=5) as pool:
        results = list(pool.map(fetch, selected.items()))
    manifest=OUT/'manifest.json'
    old=json.loads(manifest.read_text(encoding='utf-8')) if manifest.exists() else []
    combined=[x for x in old if x['id'] not in selected]+results
    manifest.write_text(json.dumps(combined, indent=2), encoding='utf-8')
    for result in results: print(json.dumps(result))
