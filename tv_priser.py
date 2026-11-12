# -*- coding: utf-8 -*-
"""Best-effort daily 75-inch TV price collector v5.1. No fabricated observations."""
import csv, json, logging, re, time, urllib.parse, os
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo
import requests
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright, TimeoutError as BrowserTimeout
from openpyxl import Workbook, load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment

ROOT = Path(os.environ.get('TV_PRISER_DIR', str(Path(__file__).resolve().parent / 'data')))
ROOT.mkdir(parents=True, exist_ok=True)
LOG = ROOT / 'innhenting.log'
logging.basicConfig(filename=LOG, level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s', encoding='utf-8')
STORES = {'Elkjøp':'elkjop.no', 'Power':'power.no', 'Elon':'elon.no', 'Komplett':'komplett.no'}
HEADERS = {'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124.0 Safari/537.36'}
COLUMNS = ['Dato','Klokkeslett','Butikk','Modell','Pris_NOK','Produkt_URL','Kilde','Status']
SESSION = requests.Session(); SESSION.headers.update(HEADERS)
CATEGORY_URLS = {
    'Power': ['https://www.power.no/c/16642/tv-og-lyd/tv/70-115-tv/'],
    'Elkjøp': ['https://www.elkjop.no/tv-lyd-og-smarte-hjem/tv-og-tilbehor/tv/75-98-tommers-tv-er'],
    'Komplett': ['https://www.komplett.no/category/325918/tv-lyd-bilde/tv-er/70-79-tommer-tv'],
    'Elon': ['https://www.elon.no/tv-lyd/tv', 'https://www.elon.no/kampanjer/black-week/lyd-bilde/tv'],
}
# URLs for additional in-store search attempts. These are candidates, not verified APIs.
SEARCH_URLS = {
    'Elkjøp': ['https://www.elkjop.no/search?query=75%20tommer%20tv'],
    'Power': ['https://www.power.no/search/?q=75%20tv'],
    'Komplett': [],
    'Elon': [],
}
# These are product links supplied by the user in an earlier conversation.
SEED_PRODUCTS = {
    'Elon': [
        'https://www.elon.no/samsung-75-tu75m72hbuxna-154668',
    ],
    'Komplett': [
        'https://www.komplett.no/product/1328473/tv-lyd-bilde/tv-er/70-79-tommer-tv/samsung-75-q6f-qled-4k-smart-tv-2025',
        'https://www.komplett.no/product/1321465/tv-lyd-bilde/tv-er/70-79-tommer-tv/samsung-75-q8fa-qled-4k-smart-tv-2025',
        'https://www.komplett.no/product/1339474/tv-lyd-bilde/tv-er/70-79-tommer-tv/lg-75-nu90-nano-uhd-smart-tv-2026',
        'https://www.komplett.no/product/1336816/tv-lyd-bilde/tv-er/70-79-tommer-tv/samsung-75-m72h-mini-led-4k-smart-tv-2026',
    ],
}
PRODUCT_PATTERNS = {'Power': r'/p-\d+/?', 'Komplett': r'/product/\d+', 'Elkjøp': r'/product/|/product-|/[^/]+/\d{6,}(?:/|$)', 'Elon': r'/(?:[^/?]+-)?\d{5,7}/?$|/produkt/|/product/|/produkter/|/p/|/p-\d+'}


def is_75_label(label):
    return bool(re.search(r'(?<!\d)75\s*(?:[\"”″]|-?tommer|inch|[- ]|(?=[a-z]))', label, re.I))


def browser_links(page, store):
    """Collect exact-size product links from rendered pages, including nested card labels."""
    pattern = PRODUCT_PATTERNS[store]
    found = {}
    items = page.locator('a[href]').evaluate_all('''(nodes) => nodes.map(a => ({
        url:a.href,
        text:(a.innerText||a.getAttribute('aria-label')||a.getAttribute('title')||'').trim(),
        card:(a.closest('article,li,[data-testid*="product"],[class*="product-card"],[class*="productCard"]')?.innerText||'').slice(0,350)
    }))''')
    for item in items:
        url = item['url'].split('#')[0].split('?')[0]
        parsed = urllib.parse.urlsplit(url)
        host = (parsed.hostname or '').lower()
        if host != STORES[store] and not host.endswith('.'+STORES[store]):
            continue
        if not re.search(pattern, parsed.path, re.I):
            continue
        label = ' '.join((item['text'], item['card'], urllib.parse.unquote(parsed.path)))
        if is_75_label(label):
            found[url] = label[:180]
    return found


def dismiss_cookies(page):
    for text in ('Godta alle', 'Godta alle informasjonskapsler', 'Aksepter alle', 'Accept all', 'Tillat alle'):
        try:
            button = page.get_by_role('button', name=text, exact=True)
            if button.count() and button.first.is_visible(timeout=800):
                button.first.click(timeout=1200)
                return
        except Exception:
            pass


def collect_browser_links(browser, store):
    """Same browser strategy for every store; different entry pages and URL patterns."""
    urls = {url: '' for url in SEED_PRODUCTS.get(store, [])}
    start_pages = list(dict.fromkeys(CATEGORY_URLS.get(store, []) + SEARCH_URLS.get(store, [])))
    for category in start_pages:
        page = browser.new_page(locale='nb-NO', timezone_id='Europe/Oslo')
        try:
            response = page.goto(category, wait_until='domcontentloaded', timeout=12000 if store == 'Komplett' else 60000)
            if response and response.status >= 400:
                logging.warning('%s: kategori HTTP %s: %s', store, response.status, category)
                continue
            page.wait_for_timeout(2500)
            dismiss_cookies(page)
            previous = -1
            for scroll in range(12):
                urls.update(browser_links(page, store))
                count = len(urls)
                if scroll >= 5 and count == previous:
                    break
                previous = count
                page.evaluate('window.scrollBy(0, Math.max(1000, window.innerHeight * 2))')
                page.wait_for_timeout(900)
                for text in ('Vis flere', 'Last inn flere', 'Load more'):
                    try:
                        btn = page.get_by_role('button', name=text, exact=False)
                        if btn.count() and btn.first.is_visible(timeout=300):
                            btn.first.click(timeout=1000)
                    except Exception:
                        pass
            logging.info('%s: kategori %s, HTTP %s, lenker i DOM %s, produktkandidater %s',
                         store, category, response.status if response else '?',
                         page.locator('a[href]').count(), len(urls))
        except Exception as exc:
            logging.warning('%s: kategori feilet %s: %s', store, category, exc)
        finally:
            page.close()
    for row in load_rows():
        if row.get('Butikk') == store:
            url = row.get('Produkt_URL','')
            if re.search(PRODUCT_PATTERNS[store], urllib.parse.urlsplit(url).path, re.I):
                urls.setdefault(url, '')
    manual = ROOT / 'produktlenker.txt'
    if manual.exists():
        for line in manual.read_text(encoding='utf-8-sig').splitlines():
            url = line.strip()
            host = (urllib.parse.urlsplit(url).hostname or '').lower()
            if (host == STORES[store] or host.endswith('.'+STORES[store])) and re.search(PRODUCT_PATTERNS[store], urllib.parse.urlsplit(url).path, re.I):
                urls.setdefault(url.split('?')[0], '')
    logging.info('%s: %s unike produktkandidater', store, len(urls))
    return urls


def browser_product(browser, store, url):
    page = browser.new_page(locale='nb-NO', timezone_id='Europe/Oslo')
    try:
        response = page.goto(url, wait_until='domcontentloaded', timeout=12000 if store == 'Komplett' else 60000)
        if response and response.status >= 400:
            logging.warning('%s: produktside HTTP %s: %s', store, response.status, url)
            return None, url
        page.wait_for_timeout(2200)
        dismiss_cookies(page)
        html = page.content()
        result = parse_product(html, page.url)
        logging.info('%s: produkt HTTP %s, pris %s, URL %s', store,
                     response.status if response else 'ukjent', result[1] if result else 'IKKE FUNNET', page.url)
        return result, page.url
    except Exception as exc:
        logging.warning('%s: produktside feilet %s: %s', store, url, exc)
        return None, url
    finally:
        page.close()


def flatten_jsonld(obj):
    if isinstance(obj,list):
        for x in obj: yield from flatten_jsonld(x)
    elif isinstance(obj,dict):
        yield obj
        if '@graph' in obj: yield from flatten_jsonld(obj['@graph'])


def parse_price(value):
    if value is None: return None
    s = re.sub(r'[^\d,.]','',str(value))
    if not s: return None
    if ',' in s and '.' in s:
        s = s.replace('.','').replace(',','.') if s.rfind(',') > s.rfind('.') else s.replace(',','')
    elif ',' in s: s = s.replace(',','.') if len(s.split(',')[-1])<=2 else s.replace(',','')
    elif '.' in s and len(s.split('.')[-1])==3: s=s.replace('.','')
    try: return round(float(s))
    except ValueError: return None


def parse_product(html, url):
    soup=BeautifulSoup(html,'html.parser')
    candidates=[]
    for tag in soup.select('script[type="application/ld+json"]'):
        try:
            for obj in flatten_jsonld(json.loads(tag.string or tag.get_text())):
                typ=obj.get('@type',[])
                if isinstance(typ,str): typ=[typ]
                if 'Product' not in typ: continue
                name=obj.get('name','')
                offers=obj.get('offers',{})
                if isinstance(offers,list): offers=next((x for x in offers if isinstance(x,dict)),{})
                if not isinstance(offers,dict): offers={}
                price=parse_price(offers.get('price') or offers.get('lowPrice') or (offers.get('priceSpecification', {}) or {}).get('price'))
                currency=offers.get('priceCurrency','NOK')
                candidates.append((name,price,currency))
        except (ValueError,TypeError): pass
    if not any(price for _, price, _ in candidates):
        title=soup.title.get_text(' ',strip=True) if soup.title else ''
        price=None
        for sel in ['meta[property="product:price:amount"]','meta[property="og:price:amount"]','meta[itemprop="price"]']:
            node=soup.select_one(sel)
            if node: price=parse_price(node.get('content')); break
        if price is None:
            for sel in ('[itemprop="price"][content]', '[data-testid="product-price"]',
                        '[data-testid="price"]', '.price-box .price', '.product-info-price .price',
                        '.special-price .price', '[class*="salesPrice"]'):
                node=soup.select_one(sel)
                if node:
                    price=parse_price(node.get('content') or node.get_text(' ',strip=True))
                    if price is not None: break
        candidates=[(title,price,'NOK')]
    for name,price,currency in candidates:
        if not isinstance(name,str) or not price or currency not in ('NOK','',None): continue
        text=(name+' '+urllib.parse.unquote(url)).lower()
        if not is_75_label(text): continue
        if not (1000<=price<=20000): continue
        if not any(x in text for x in ('tv','qled','oled','led','samsung','lg ','tcl','hisense','philips','sony','panasonic','cepter')): continue
        return name.strip()[:180],price
    return None


def load_rows():
    p=ROOT/'prisobservasjoner.csv'
    if not p.exists(): return []
    with p.open(encoding='utf-8-sig',newline='') as f: return list(csv.DictReader(f,delimiter=';'))


def write_rows(rows):
    target = ROOT / 'prisobservasjoner.csv'
    temporary = ROOT / 'prisobservasjoner.tmp.csv'
    with temporary.open('w',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=COLUMNS,delimiter=';');writer.writeheader();writer.writerows(rows)
    temporary.replace(target)


def excel(rows, stats):
    wb=Workbook(); raw=wb.active;raw.title='Prisobservasjoner';raw.append(COLUMNS)
    for r in rows: raw.append([r.get(k,'') if k!='Pris_NOK' else int(r[k]) for k in COLUMNS])
    raw.freeze_panes='A2';raw.auto_filter.ref=raw.dimensions
    latest={}
    for r in sorted(rows, key=lambda r:(r.get('Dato',''),r.get('Klokkeslett',''))):
        latest[(r['Butikk'],r['Produkt_URL'])]=r
    summary=wb.create_sheet('Siste priser');summary.append(COLUMNS)
    for r in sorted(latest.values(),key=lambda x:int(x['Pris_NOK'])):
        summary.append([r.get(k,'') if k!='Pris_NOK' else int(r[k]) for k in COLUMNS])
    summary.freeze_panes='A2';summary.auto_filter.ref=summary.dimensions
    history=wb.create_sheet('Prisutvikling')
    history.append(['Dato'] + list(STORES))
    daily = {}
    for r in rows:
        try:
            key=(r['Dato'],r['Butikk'])
            daily[key]=min(daily.get(key, 10**9), int(r['Pris_NOK']))
        except (ValueError, KeyError):
            continue
    for date in sorted({d for d,_ in daily}):
        history.append([date] + [daily.get((date,store)) for store in STORES])
    chart=LineChart()
    chart.title='Laveste observerte pris per butikk og dag'
    chart.y_axis.title='NOK'; chart.x_axis.title='Dato'
    chart.width=23; chart.height=12
    if history.max_row>1:
        chart.add_data(Reference(history,min_col=2,max_col=5,min_row=1,max_row=history.max_row),titles_from_data=True)
        chart.set_categories(Reference(history,min_col=1,min_row=2,max_row=history.max_row))
        history.add_chart(chart,'G2')
    deals=wb.create_sheet('Under 10 000 kr')
    deals.append(COLUMNS)
    for r in sorted(latest.values(), key=lambda x:int(x['Pris_NOK'])):
        if int(r['Pris_NOK']) < 10000:
            deals.append([r.get(k,'') if k!='Pris_NOK' else int(r[k]) for k in COLUMNS])
    deals.freeze_panes='A2'; deals.auto_filter.ref=deals.dimensions
    status=wb.create_sheet('Kjøringsstatus');status.append(['Butikk','Oppdagede lenker','Gyldige priser kontrollert','Merknad'])
    for store,found,valid in stats: status.append([store,found,valid,'Ufullstendig dekning mulig; ingen pris funnet betyr ikke utsolgt'])
    status.append(['Metode','','','Playwright for Elkjøp, Power og Elon. Komplett er deaktivert som standard grunnet nettverksfeil.'])
    for ws in wb:
        ws.sheet_view.showGridLines=False
        for cell in ws[1]: cell.font=Font(color='FFFFFF',bold=True);cell.fill=PatternFill('solid',fgColor='17365D');cell.alignment=Alignment(wrap_text=True)
        for col,width in {'A':15,'B':16,'C':18,'D':65,'E':16,'F':65,'G':20,'H':20}.items(): ws.column_dimensions[col].width=width
    for ws in (raw,summary,deals):
        for row in ws.iter_rows(min_row=2): row[4].number_format='#,##0 "kr"'
    tmp=ROOT/'TV_priser.tmp.xlsx';wb.save(tmp);tmp.replace(ROOT/'TV_priser.xlsx')


def main():
    now=datetime.now(ZoneInfo('Europe/Oslo'))
    rows=load_rows();seen={(r['Dato'],r['Butikk'],r['Produkt_URL']) for r in rows}
    stats=[]
    enable_komplett = os.environ.get('TV_TEST_KOMPLETT', '').strip() == '1'
    logging.info('Oppstart TV-priser v5.1. Komplett: %s', 'test aktivert' if enable_komplett else 'hoppes over')
    print('Starter TV-priser Cloud (Komplett er deaktivert som standard)', flush=True)
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # Separate Chromium process for Komplett: HTTP/2 is disabled to address
        # the ERR_HTTP2_PROTOCOL_ERROR observed on this user's Windows PC.
        komplett_browser = None
        try:
            for store,domain in STORES.items():
                if store == 'Komplett' and not enable_komplett:
                    logging.info('Komplett: hoppes over pga. dokumenterte nettverksfeil. Sett TV_TEST_KOMPLETT=1 for test.')
                    stats.append((store, 0, 0))
                    continue
                print('Kontrollerer ' + store + ' ...', flush=True)
                logging.info('%s: starter innhenting', store)
                active_browser = browser
                if store == 'Komplett':
                    try:
                        komplett_browser = p.chromium.launch(headless=True, args=['--disable-http2'])
                        active_browser = komplett_browser
                        logging.info('Komplett: Chromium med --disable-http2 aktivert')
                    except Exception as exc:
                        logging.warning('Komplett: alternativ Chromium feilet: %s', exc)
                try:
                    urls = collect_browser_links(active_browser, store)
                except Exception as exc:
                    logging.exception('%s: lenkeinnhenting feilet: %s', store, exc)
                    stats.append((store, 0, 0))
                    continue
                valid = 0
                for url in list(urls)[:120]:
                    try:
                        result, final_url = browser_product(active_browser, store, url)
                    except Exception as exc:
                        logging.exception('%s: uventet feil for %s: %s', store, url, exc)
                        continue
                    if result is None:
                        continue
                    name,price = result
                    key=(now.date().isoformat(),store,final_url.split('?')[0])
                    valid+=1
                    if key in seen:
                        # Keep one observation per day; update if a later run sees a new price.
                        for existing in rows:
                            if (existing['Dato'],existing['Butikk'],existing['Produkt_URL']) == key:
                                existing.update({'Klokkeslett':now.strftime('%H:%M'),'Pris_NOK':str(price),'Modell':name})
                                break
                        continue
                    rows.append(dict(zip(COLUMNS,[key[0],now.strftime('%H:%M'),store,name,price,key[2],'Playwright produktside','Bekreftet ved innhenting'])))
                    seen.add(key)
                    logging.info('%s: lagret %s kr: %s',store,price,name)
                stats.append((store,len(urls),valid))
                logging.info('%s: %s produktlenker, %s gyldige priser kontrollert',store,len(urls),valid)
                write_rows(rows)  # Save each completed store, even if a later store fails.
                print(f'{store}: {valid} gyldige priser kontrollert', flush=True)
        finally:
            write_rows(rows)
            if komplett_browser is not None:
                komplett_browser.close()
            browser.close()
    write_rows(rows);excel(rows,stats)
    print('Ferdig:',len(rows),'totale observasjoner; gyldige priser kontrollert:',sum(x[2] for x in stats))
    print('Excel:',ROOT/'TV_priser.xlsx')
    if not sum(x[2] for x in stats):
        print('MERK: Ingen gyldige priser i denne kjoringen. Se logg.')
        if not any(x[1] for x in stats): raise SystemExit(2)

if __name__=='__main__': main()
