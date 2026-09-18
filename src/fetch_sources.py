"""Archive public sources used in the competition research; no credentials."""
from pathlib import Path
import json
from concurrent.futures import ThreadPoolExecutor
import requests
from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'research' / 'sources'
OUT.mkdir(parents=True, exist_ok=True)
SOURCES = {
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
    with ThreadPoolExecutor(max_workers=5) as pool:
        results = list(pool.map(fetch, SOURCES.items()))
    (OUT / 'manifest.json').write_text(json.dumps(results, indent=2), encoding='utf-8')
    for result in results: print(json.dumps(result))
