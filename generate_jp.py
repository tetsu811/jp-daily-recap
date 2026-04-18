#!/usr/bin/env python3
"""日股復盤儀表板生成器

抓東証 17 業種代表個股日線資料,輸出單頁靜態 HTML,包含:
- 大盤概況(日經平均、TOPIX)
- 17 業種漲跌排行(等權重、中位數)
- 每板塊證據瓦片:放量突破、強勢、弱勢 TOP 3

相依: yfinance, pandas, numpy
"""
import json
import sys
import time
import warnings
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib import parse as urlparse
from urllib import request as urlrequest
from xml.etree import ElementTree as ET

import numpy as np
import pandas as pd
import yfinance as yf

warnings.filterwarnings('ignore')

HERE = Path(__file__).parent
SECTORS_JSON = HERE / 'sectors.json'
OUTPUT_HTML = HERE / 'jp_dashboard.html'
OUTPUT_EMBED = HERE / 'jp_dashboard_embed.html'
MCAP_CACHE = HERE / 'market_caps.json'
NEWS_CACHE = HERE / 'news_cache.json'
MCAP_REFRESH_DAYS = 30
JST = timezone(timedelta(hours=9))


def load_sectors():
    with open(SECTORS_JSON, 'r', encoding='utf-8') as f:
        return json.load(f)


def yf_ticker(code: str) -> str:
    if code.startswith('^') or code.endswith('.T'):
        return code
    return f"{code}.T"


def build_ticker_map(sectors: dict):
    """Return {yf_ticker: (sector, name)}. Skips _meta and 指数 keys.
    Dedupes across sectors (last-write wins, which is rare in curated list)."""
    m = {}
    for sector, stocks in sectors.items():
        if sector.startswith('_') or sector == '指数':
            continue
        for stk in stocks:
            m[yf_ticker(stk['code'])] = (sector, stk['name'])
    return m


def fetch_indices(index_list):
    """Fetch index values one-by-one (small list, handles ^ tickers)."""
    out = []
    for item in index_list:
        code = item['code']
        try:
            hist = yf.Ticker(code).history(period='10d', auto_adjust=False)
            if len(hist) < 2:
                continue
            last = float(hist['Close'].iloc[-1])
            prev = float(hist['Close'].iloc[-2])
            out.append({
                'name': item['name'],
                'code': code,
                'close': last,
                'chg': (last - prev) / prev * 100,
                'date': hist.index[-1].date(),
            })
        except Exception as e:
            print(f"  [index err] {code}: {e}", file=sys.stderr)
    return out


def fetch_all(tickers):
    """Batch-download ~90d OHLCV. Returns multi-index DataFrame."""
    print(f"Downloading {len(tickers)} tickers (yfinance batch)...")
    df = yf.download(
        tickers=' '.join(tickers),
        period='120d',
        group_by='ticker',
        threads=True,
        progress=False,
        auto_adjust=False,
    )
    return df


def stock_metrics(df, ticker):
    try:
        sub = df[ticker].dropna(subset=['Close'])
    except (KeyError, ValueError):
        return None
    if len(sub) < 21:
        return None
    close = sub['Close'].astype(float)
    vol = sub['Volume'].astype(float)
    last_close = float(close.iloc[-1])
    prev_close = float(close.iloc[-2])
    if prev_close <= 0:
        return None
    chg = (last_close - prev_close) / prev_close * 100
    # 5-day change (close vs 5 sessions ago)
    chg_5d = None
    if len(sub) >= 6:
        prev_5 = float(close.iloc[-6])
        if prev_5 > 0:
            chg_5d = (last_close - prev_5) / prev_5 * 100
    last_vol = float(vol.iloc[-1])
    avg_vol_20 = float(vol.iloc[-21:-1].mean())
    vol_ratio = last_vol / avg_vol_20 if avg_vol_20 > 0 else 0.0
    high_20 = float(close.iloc[-21:-1].max())
    ma20 = float(close.iloc[-20:].mean())
    # RSI14
    delta = close.diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = -delta.clip(upper=0).rolling(14).mean()
    rs = gain / loss.replace(0, np.nan)
    rsi_series = 100 - 100 / (1 + rs)
    rsi = float(rsi_series.iloc[-1]) if pd.notna(rsi_series.iloc[-1]) else None
    # 52w-ish low (using all available data, typically ~120d)
    low_min = float(close.min())
    near_low = (last_close - low_min) / low_min * 100 if low_min > 0 else 0
    return {
        'chg': chg,
        'chg_5d': chg_5d,
        'close': last_close,
        'vol_ratio': vol_ratio,
        'broke_20d_high': last_close > high_20,
        'ma20': ma20,
        'above_ma20': last_close > ma20,
        'rsi': rsi,
        'near_low_pct': near_low,
        'last_vol': last_vol,
        'date': sub.index[-1].date(),
    }


# ---------- B. News RSS (Google News JP) ----------

SECTOR_NEWS_QUERIES = {
    "食品": ["食品株", "食料品 決算"],
    "エネルギー資源": ["原油", "石油株", "エネルギー株"],
    "建設・資材": ["建設株", "ゼネコン"],
    "素材・化学": ["化学株", "素材株"],
    "医薬品": ["製薬 株", "医薬品株"],
    "自動車・輸送機": ["自動車株", "トヨタ 日産"],
    "鉄鋼・非鉄": ["鉄鋼株", "非鉄金属"],
    "機械": ["機械株", "工作機械"],
    "電機・精密": ["半導体株", "電機 精密"],
    "情報通信・サービスその他": ["IT株 日本", "通信株"],
    "電気・ガス": ["電力株", "ガス会社 株"],
    "運輸・物流": ["海運株", "鉄道株"],
    "商社・卸売": ["商社株", "総合商社"],
    "小売": ["小売株", "百貨店"],
    "銀行": ["銀行株", "メガバンク"],
    "金融(除く銀行)": ["証券株", "保険株"],
    "不動産": ["不動産株", "REIT"],
}


def _fetch_news_one(sector, queries, max_items=3):
    headlines = []
    q = ' OR '.join(f'"{k}"' for k in queries)
    url = f"https://news.google.com/rss/search?q={urlparse.quote(q)}&hl=ja&gl=JP&ceid=JP:ja"
    try:
        req = urlrequest.Request(url, headers={'User-Agent': 'jp-recap/1.0'})
        with urlrequest.urlopen(req, timeout=15) as r:
            xml = r.read()
        root = ET.fromstring(xml)
        today = datetime.now(JST).date()
        for item in root.findall('.//item')[:12]:
            title_el = item.find('title')
            link_el = item.find('link')
            pub_el = item.find('pubDate')
            if title_el is None:
                continue
            title = title_el.text or ''
            link = link_el.text if link_el is not None else ''
            pub = pub_el.text if pub_el is not None else ''
            # Parse pubDate to filter today+yesterday
            pub_date = None
            try:
                pub_date = datetime.strptime(pub, '%a, %d %b %Y %H:%M:%S %Z').astimezone(JST).date()
            except Exception:
                pass
            if pub_date and (today - pub_date).days > 1:
                continue
            # Strip trailing " - source" Google News adds
            title_clean = title.rsplit(' - ', 1)[0].strip()
            headlines.append({'title': title_clean, 'link': link, 'date': str(pub_date) if pub_date else ''})
            if len(headlines) >= max_items:
                break
    except Exception as e:
        print(f"  news err ({sector}): {e}", file=sys.stderr)
    return sector, headlines


def fetch_news_all(sector_names):
    """Parallel fetch Google News RSS for each sector."""
    out = {}
    to_fetch = [(s, SECTOR_NEWS_QUERIES.get(s, [])) for s in sector_names]
    with ThreadPoolExecutor(max_workers=8) as ex:
        futures = [ex.submit(_fetch_news_one, s, qs) for s, qs in to_fetch if qs]
        for f in as_completed(futures):
            sector, items = f.result()
            out[sector] = items
    # Cache for debug
    try:
        NEWS_CACHE.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:
        pass
    return out


def _fetch_stock_news_one(code, name, max_items=2):
    """個股新聞 — 用公司名 + 股價/決算 關鍵字。"""
    # 去掉 HD、グループ 等 noise 讓搜尋更準
    short = name.replace('ホールディングス', '').replace('HD', '').replace('グループ', '').strip()
    query = f'"{short}" 株価 OR 決算 OR 業績'
    url = f"https://news.google.com/rss/search?q={urlparse.quote(query)}&hl=ja&gl=JP&ceid=JP:ja"
    headlines = []
    try:
        req = urlrequest.Request(url, headers={'User-Agent': 'jp-recap/1.0'})
        with urlrequest.urlopen(req, timeout=15) as r:
            xml = r.read()
        root = ET.fromstring(xml)
        today = datetime.now(JST).date()
        for item in root.findall('.//item')[:8]:
            title_el = item.find('title')
            link_el = item.find('link')
            pub_el = item.find('pubDate')
            if title_el is None:
                continue
            title = (title_el.text or '').rsplit(' - ', 1)[0].strip()
            link = link_el.text if link_el is not None else ''
            pub = pub_el.text if pub_el is not None else ''
            pub_date = None
            try:
                pub_date = datetime.strptime(pub, '%a, %d %b %Y %H:%M:%S %Z').astimezone(JST).date()
            except Exception:
                pass
            # 個股新聞僅取今日/昨日(<=1 天)
            if pub_date and (today - pub_date).days > 1:
                continue
            headlines.append({'title': title, 'link': link, 'date': str(pub_date) if pub_date else ''})
            if len(headlines) >= max_items:
                break
    except Exception as e:
        print(f"  stock-news err ({code} {name}): {e}", file=sys.stderr)
    return code, headlines


def fetch_stock_news_for_report(report):
    """為每個板塊的 top_contributors 抓個股新聞"""
    # Collect all (code, name) pairs
    tasks = []
    seen = set()
    for r in report:
        for s in r.get('top_contributors', [])[:3]:
            key = s['code']
            if key in seen:
                continue
            seen.add(key)
            tasks.append((s['code'], s['name']))
    print(f"Fetching stock-level news for {len(tasks)} contributors...")
    out = {}
    with ThreadPoolExecutor(max_workers=10) as ex:
        futures = [ex.submit(_fetch_stock_news_one, c, n) for c, n in tasks]
        for f in as_completed(futures):
            code, items = f.result()
            out[code] = items
    # Attach back to stocks
    for r in report:
        for s in r.get('top_contributors', [])[:3]:
            s['news'] = out.get(s['code'], [])
    total_headlines = sum(len(v) for v in out.values())
    print(f"Stock news fetched: {total_headlines} headlines for {sum(1 for v in out.values() if v)}/{len(tasks)} stocks")
    return out


# ---------- C. Claude LLM one-sentence summaries ----------

CLAUDE_MODEL = 'claude-sonnet-4-5'


def build_summary_prompt(report):
    """把 17 個板塊的今日資料組成一個 prompt,一次叫 Claude 產出所有 summary。
    餵 sector 新聞 + 權重貢獻股的個股新聞,讓總結能 reference 具體催化劑。"""
    lines = [
        "你是日股分析師。下面是今日東証 17 業種表現。每個板塊給我一句 30 字內的繁體中文解釋,",
        "專注於「為什麼今天動這樣」 — 優先引用具體事件(法人調升/降評、配息變動、業績預警、併購、新品、訴訟、產業政策等),",
        "不要只說「權重股拖累」這種廢話。沒明確原因再寫「廣泛式,無具體催化劑」。",
        "",
        "輸出格式(一行一個板塊,嚴格遵守):",
        "板塊名|解釋句",
        "",
        "---資料---",
    ]
    for r in report:
        main_chg = r['mcap_chg'] if r.get('mcap_chg') is not None else r['avg_chg']
        gainer_str = ', '.join(f"{g['name']}{g['chg']:+.1f}%" for g in r['top_gainers'][:2])
        loser_str = ', '.join(f"{l['name']}{l['chg']:+.1f}%" for l in r['top_losers'][:2])
        sector_news_str = '; '.join(n['title'][:60] for n in r.get('news', [])[:2])
        lines.append(f"\n【{r['sector']}】{main_chg:+.2f}% (成分 {r['n']} 檔)")
        # 每個權重主導股 + 它的個股新聞(最多 1 條)
        contribs = r.get('top_contributors', [])[:3]
        if contribs:
            lines.append("  權重主導 + 個股催化劑:")
            for c in contribs:
                row = f"    {c['name']} {c['chg']:+.1f}% (貢獻{c['contribution_pp']:+.2f}pp, 權{c.get('weight',0)*100:.0f}%)"
                stock_news = c.get('news', [])
                if stock_news:
                    # 去掉那種 "AIが解説" 制式文章,留真新聞
                    real = [n for n in stock_news if 'AI' not in n['title'] and '解説' not in n['title']]
                    if real:
                        row += f" — 新聞:{real[0]['title'][:80]}"
                lines.append(row)
        if gainer_str:
            lines.append(f"  其他領漲: {gainer_str}")
        if loser_str:
            lines.append(f"  其他領跌: {loser_str}")
        if sector_news_str:
            lines.append(f"  產業新聞: {sector_news_str}")
    lines.append("")
    lines.append("請現在產出,嚴格格式 `板塊名|解釋句`,一行一個。不要加任何前後說明。")
    return '\n'.join(lines)


def fetch_summaries(report, api_key):
    """呼叫 Claude API,回傳 {sector_name: summary_sentence}"""
    import os
    prompt = build_summary_prompt(report)
    body = json.dumps({
        'model': CLAUDE_MODEL,
        'max_tokens': 2000,
        'messages': [{'role': 'user', 'content': prompt}],
    }).encode('utf-8')
    req = urlrequest.Request(
        'https://api.anthropic.com/v1/messages',
        data=body,
        method='POST',
        headers={
            'x-api-key': api_key,
            'anthropic-version': '2023-06-01',
            'content-type': 'application/json',
        },
    )
    try:
        with urlrequest.urlopen(req, timeout=60) as r:
            data = json.loads(r.read())
    except Exception as e:
        print(f"  LLM err: {e}", file=sys.stderr)
        return {}
    text = data.get('content', [{}])[0].get('text', '')
    out = {}
    for line in text.splitlines():
        line = line.strip()
        if '|' not in line:
            continue
        name, _, sentence = line.partition('|')
        name = name.strip()
        sentence = sentence.strip()
        if name and sentence:
            out[name] = sentence
    return out


def load_market_caps(tickers):
    """載入/更新市值快取。快取 > 30 天才重抓,避免每天都打 .info (慢)。"""
    import time
    now_ts = time.time()
    cache = {}
    if MCAP_CACHE.exists():
        try:
            cache = json.loads(MCAP_CACHE.read_text(encoding='utf-8'))
        except Exception:
            cache = {}
    fetched_ts = cache.get('_fetched_at', 0)
    age_days = (now_ts - fetched_ts) / 86400
    if age_days < MCAP_REFRESH_DAYS and all(t in cache for t in tickers):
        print(f"  mcap cache hit (age {age_days:.1f}d)")
        return {t: cache[t] for t in tickers}
    print(f"  mcap cache stale or incomplete, refetching (age {age_days:.1f}d)")
    caps = {}
    for i, t in enumerate(tickers):
        try:
            info = yf.Ticker(t).fast_info
            mc = getattr(info, 'market_cap', None)
            if mc:
                caps[t] = float(mc)
        except Exception as e:
            pass
        if (i + 1) % 50 == 0:
            print(f"  mcap {i+1}/{len(tickers)}")
    caps['_fetched_at'] = now_ts
    MCAP_CACHE.write_text(json.dumps(caps, ensure_ascii=False, indent=2), encoding='utf-8')
    return {t: caps[t] for t in tickers if t in caps}


def compute_contributions(stocks, mcap_map):
    """Add mcap + weight + contribution to each stock; sort by contribution magnitude."""
    total_mcap = 0
    for s in stocks:
        mc = mcap_map.get(s['ticker'], 0)
        s['mcap'] = mc
        total_mcap += mc
    for s in stocks:
        if total_mcap > 0 and s['mcap']:
            s['weight'] = s['mcap'] / total_mcap
            s['contribution_pp'] = s['weight'] * s['chg']  # 百分點貢獻
        else:
            s['weight'] = 0
            s['contribution_pp'] = 0
    return total_mcap


def build_sector_report(ticker_map, df, mcap_map=None):
    by_sector = {}
    for ticker, (sector, name) in ticker_map.items():
        m = stock_metrics(df, ticker)
        if m is None:
            continue
        by_sector.setdefault(sector, []).append({
            'ticker': ticker,
            'code': ticker.replace('.T', ''),
            'name': name,
            **m,
        })
    report = []
    for sector, stocks in by_sector.items():
        if not stocks:
            continue
        chgs = np.array([s['chg'] for s in stocks])
        breakouts = sorted(
            [s for s in stocks if s['vol_ratio'] >= 2.0 and s['broke_20d_high']],
            key=lambda x: x['vol_ratio'], reverse=True,
        )[:3]
        bottom_volume = sorted(
            [s for s in stocks if s['vol_ratio'] >= 2.0 and s['near_low_pct'] <= 10],
            key=lambda x: x['vol_ratio'], reverse=True,
        )[:3]
        top_gainers = sorted(stocks, key=lambda x: x['chg'], reverse=True)[:3]
        top_losers = sorted(stocks, key=lambda x: x['chg'])[:3]
        # 市值加權(A):
        if mcap_map:
            compute_contributions(stocks, mcap_map)
            top_contributors = sorted(stocks, key=lambda x: abs(x['contribution_pp']), reverse=True)[:3]
            mcap_weighted_chg = sum(s['contribution_pp'] for s in stocks)
            # 5d mcap-weighted chg
            total_mcap = sum(s.get('mcap', 0) for s in stocks)
            if total_mcap > 0:
                mcap_chg_5d = sum(
                    (s.get('chg_5d') or 0) * (s.get('mcap', 0) / total_mcap) for s in stocks
                )
            else:
                mcap_chg_5d = None
        else:
            top_contributors = []
            mcap_weighted_chg = None
            mcap_chg_5d = None
        # 5d equal-weight
        chgs_5d = [s['chg_5d'] for s in stocks if s.get('chg_5d') is not None]
        avg_chg_5d = float(np.mean(chgs_5d)) if chgs_5d else None
        report.append({
            'sector': sector,
            'avg_chg': float(chgs.mean()),
            'avg_chg_5d': avg_chg_5d,
            'mcap_chg': mcap_weighted_chg,
            'mcap_chg_5d': mcap_chg_5d,
            'median_chg': float(np.median(chgs)),
            'advancers': int((chgs > 0).sum()),
            'decliners': int((chgs < 0).sum()),
            'n': len(stocks),
            'volume_breakouts': breakouts,
            'bottom_volume': bottom_volume,
            'top_gainers': top_gainers,
            'top_losers': top_losers,
            'top_contributors': top_contributors,
        })
    # Sort by market-cap weighted change if available, else equal-weighted
    report.sort(key=lambda x: x.get('mcap_chg') if x.get('mcap_chg') is not None else x['avg_chg'], reverse=True)
    return report


# ---------- HTML rendering ----------

def _stock_row(s, highlight_key=None):
    """Render a single stock row inside an evidence tile."""
    chg_cls = 'up' if s['chg'] > 0 else ('down' if s['chg'] < 0 else '')
    chg_str = f"{s['chg']:+.2f}%"
    extra = ''
    if highlight_key == 'vol':
        extra = f"<span class='extra'>量 {s['vol_ratio']:.1f}×</span>"
    elif highlight_key == 'rsi' and s.get('rsi') is not None:
        extra = f"<span class='extra'>RSI {s['rsi']:.0f}</span>"
    elif highlight_key == 'contrib':
        c = s.get('contribution_pp', 0)
        w = s.get('weight', 0) * 100
        extra = f"<span class='extra'>{c:+.2f}pp 權{w:.0f}%</span>"
    news_html = ''
    if highlight_key == 'contrib' and s.get('news'):
        news_html = "<div class='stock-news'>" + ''.join(
            f'<a href="{n["link"]}" target="_blank" rel="noopener">↳ {n["title"]}</a>'
            for n in s['news'][:2]
        ) + "</div>"
    return (
        f"<div class='stock-row'>"
        f"<span class='code'>{s['code']}</span>"
        f"<span class='name'>{s['name']}</span>"
        f"<span class='chg {chg_cls}'>{chg_str}</span>"
        f"{extra}"
        f"</div>"
        f"{news_html}"
    )


def _tile(title, stocks, key=None, empty_text='— 無 —'):
    if not stocks:
        body = f"<div class='tile-empty'>{empty_text}</div>"
    else:
        body = ''.join(_stock_row(s, highlight_key=key) for s in stocks)
    return (
        f"<div class='tile'>"
        f"<div class='tile-title'>{title}</div>"
        f"{body}"
        f"</div>"
    )


def _sector_card(s):
    main_chg = s['mcap_chg'] if s.get('mcap_chg') is not None else s['avg_chg']
    chg_cls = 'up' if main_chg > 0 else ('down' if main_chg < 0 else '')
    width = min(abs(main_chg) * 15, 100)
    bar_side = 'bar-up' if main_chg > 0 else 'bar-down'
    meta_parts = [f"成分 {s['n']} 檔", f"漲 {s['advancers']} 跌 {s['decliners']}"]
    if s.get('mcap_chg') is not None:
        meta_parts.append(f"等權 {s['avg_chg']:+.2f}%")
    meta_parts.append(f"中位 {s['median_chg']:+.2f}%")
    chg5 = s.get('mcap_chg_5d') if s.get('mcap_chg_5d') is not None else s.get('avg_chg_5d')
    if chg5 is not None:
        meta_parts.append(f"5日 {chg5:+.2f}%")
    if s.get('alpha_pp') is not None:
        meta_parts.append(f"vs TOPIX {s['alpha_pp']:+.2f}pp")
    summary_html = f"<div class='sec-summary'>💬 {s['summary']}</div>" if s.get('summary') else ''
    return (
        f"<details class='sector-card' open>"
        f"<summary class='sector-summary'>"
        f"<div class='sec-head'>"
        f"<span class='sec-name'>{s['sector']}</span>"
        f"<span class='sec-chg {chg_cls}'>{main_chg:+.2f}%</span>"
        f"</div>"
        f"<div class='sec-meta'>{' | '.join(meta_parts)}</div>"
        f"<div class='sec-bar'><div class='{bar_side}' style='width:{width}%'></div></div>"
        f"</summary>"
        f"{summary_html}"
        f"<div class='tiles'>"
        f"{_tile('🎯 權重貢獻 TOP3', s.get('top_contributors', []), key='contrib', empty_text='無市值資料')}"
        f"{_tile('🔥 放量突破 TOP3', s['volume_breakouts'], key='vol', empty_text='今日無放量創新高')}"
        f"{_tile('📈 漲幅 TOP3', s['top_gainers'])}"
        f"{_tile('📉 跌幅 TOP3', s['top_losers'])}"
        f"{_tile('💡 底部放量 TOP3', s['bottom_volume'], key='vol', empty_text='今日無底部放量')}"
        f"{_news_tile(s.get('news', []))}"
        f"</div>"
        f"</details>"
    )


def _index_card(idx):
    chg_cls = 'up' if idx['chg'] > 0 else ('down' if idx['chg'] < 0 else '')
    if idx.get('is_breadth'):
        # breadth card: 顯示等權%、漲跌家數比
        adv = idx.get('advancers', 0)
        dec = idx.get('decliners', 0)
        return (
            f"<div class='index-card breadth'>"
            f"<div class='idx-name'>{idx['name']}</div>"
            f"<div class='idx-close {chg_cls}'>{idx['chg']:+.2f}%</div>"
            f"<div class='idx-chg'>漲 <span class='up'>{adv}</span> / 跌 <span class='down'>{dec}</span></div>"
            f"</div>"
        )
    return (
        f"<div class='index-card'>"
        f"<div class='idx-name'>{idx['name']}</div>"
        f"<div class='idx-close'>{idx['close']:,.2f}</div>"
        f"<div class='idx-chg {chg_cls}'>{idx['chg']:+.2f}%</div>"
        f"</div>"
    )


EMBED_STYLE = """<style>
  .jpr-root {
    --jpr-bg: #fafaf7;
    --jpr-card: #ffffff;
    --jpr-border: #e5e2dc;
    --jpr-text: #2a2a2a;
    --jpr-muted: #888;
    --jpr-up: #c0392b;
    --jpr-down: #27874b;
    background: var(--jpr-bg);
    color: var(--jpr-text);
    font-family: -apple-system, "Hiragino Sans", "Noto Sans JP", "Noto Sans TC", sans-serif;
    font-size: 14px;
    line-height: 1.5;
    padding: 16px;
    box-sizing: border-box;
  }
  .jpr-root * { box-sizing: border-box; }
  .jpr-root header { margin-bottom: 16px; }
  .jpr-root h1 { font-size: 18px; margin: 0 0 4px; font-weight: 600; }
  .jpr-root .sub { color: var(--jpr-muted); font-size: 12px; }
  .jpr-root .indices { display: flex; gap: 10px; margin: 12px 0 20px; flex-wrap: wrap; }
  .jpr-root .index-card { flex: 1 1 140px; background: var(--jpr-card); border: 1px solid var(--jpr-border); border-radius: 6px; padding: 10px 12px; }
  .jpr-root .idx-name { font-size: 12px; color: var(--jpr-muted); }
  .jpr-root .idx-close { font-size: 18px; font-weight: 600; margin-top: 2px; }
  .jpr-root .idx-chg { font-size: 13px; margin-top: 2px; }
  .jpr-root .sector-card { background: var(--jpr-card); border: 1px solid var(--jpr-border); border-radius: 6px; margin-bottom: 10px; overflow: hidden; }
  .jpr-root .sector-summary { cursor: pointer; list-style: none; padding: 12px 14px; user-select: none; }
  .jpr-root .sector-summary::-webkit-details-marker { display: none; }
  .jpr-root .sec-head { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; }
  .jpr-root .sec-name { font-size: 15px; font-weight: 600; }
  .jpr-root .sec-chg { font-size: 15px; font-weight: 600; font-variant-numeric: tabular-nums; }
  .jpr-root .sec-meta { font-size: 11px; color: var(--jpr-muted); margin-top: 2px; }
  .jpr-root .sec-bar { height: 3px; background: #f0ede8; margin-top: 8px; border-radius: 2px; overflow: hidden; }
  .jpr-root .bar-up { background: var(--jpr-up); height: 100%; border-radius: 2px; }
  .jpr-root .bar-down { background: var(--jpr-down); height: 100%; border-radius: 2px; }
  .jpr-root .tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 8px; padding: 0 14px 14px; }
  .jpr-root .tile { background: #fbfaf6; border: 1px solid #ece8df; border-radius: 4px; padding: 8px 10px; }
  .jpr-root .tile-title { font-size: 12px; color: var(--jpr-muted); margin-bottom: 4px; font-weight: 500; }
  .jpr-root .tile-empty { font-size: 12px; color: #bbb; padding: 4px 0; }
  .jpr-root .stock-row { display: grid; grid-template-columns: 48px 1fr auto auto; gap: 6px; padding: 3px 0; font-size: 13px; align-items: baseline; }
  .jpr-root .stock-row .code { color: var(--jpr-muted); font-variant-numeric: tabular-nums; font-size: 12px; }
  .jpr-root .stock-row .name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
  .jpr-root .stock-row .chg { font-variant-numeric: tabular-nums; font-weight: 500; }
  .jpr-root .stock-row .extra { font-size: 11px; color: var(--jpr-muted); min-width: 52px; text-align: right; }
  .jpr-root .up { color: var(--jpr-up); }
  .jpr-root .down { color: var(--jpr-down); }
  .jpr-root footer { margin-top: 24px; padding-top: 12px; border-top: 1px solid var(--jpr-border); color: var(--jpr-muted); font-size: 11px; }
</style>"""


# ---------- inline-style renderers (WP embed 用,避開 <style>/class 過濾) ----------

_ROW_STYLE = 'display:grid;grid-template-columns:48px 1fr auto 100px;gap:6px;padding:3px 0;font-size:13px;align-items:baseline;'
_CODE_STYLE = 'color:#888;font-size:12px;font-variant-numeric:tabular-nums;'
_NAME_STYLE = 'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#2a2a2a;'
_CHG_STYLE = 'font-variant-numeric:tabular-nums;font-weight:500;'
_EXTRA_STYLE = 'font-size:11px;color:#888;text-align:right;'


def _color(chg):
    return '#c0392b' if chg > 0 else ('#27874b' if chg < 0 else '#888')


def _stock_row_inline(s, highlight_key=None):
    chg_color = _color(s['chg'])
    chg_str = f"{s['chg']:+.2f}%"
    extra = ''
    if highlight_key == 'vol':
        extra = f'<div style="{_EXTRA_STYLE}">量 {s["vol_ratio"]:.1f}×</div>'
    elif highlight_key == 'rsi' and s.get('rsi') is not None:
        extra = f'<div style="{_EXTRA_STYLE}">RSI {s["rsi"]:.0f}</div>'
    elif highlight_key == 'contrib':
        contrib = s.get('contribution_pp', 0)
        weight = s.get('weight', 0) * 100
        extra = f'<div style="{_EXTRA_STYLE}">{contrib:+.2f}pp 權{weight:.0f}%</div>'
    else:
        extra = '<div></div>'
    row = (
        f'<div style="{_ROW_STYLE}">'
        f'<div style="{_CODE_STYLE}">{s["code"]}</div>'
        f'<div style="{_NAME_STYLE}">{s["name"]}</div>'
        f'<div style="{_CHG_STYLE}color:{chg_color};">{chg_str}</div>'
        f'{extra}'
        f'</div>'
    )
    # Per-stock news (只 contrib tile 顯示)
    if highlight_key == 'contrib' and s.get('news'):
        news_style = 'padding-left:54px;padding-bottom:4px;'
        link_style = 'display:block;font-size:11px;color:#1a4d8c;text-decoration:none;padding:1px 0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
        news_html = ''.join(
            f'<a href="{n["link"]}" target="_blank" rel="noopener" style="{link_style}">↳ {n["title"]}</a>'
            for n in s['news'][:2]
        )
        row += f'<div style="{news_style}">{news_html}</div>'
    return row


def _news_tile_inline(news_items):
    tile_style = 'background:#fbfaf6;border:1px solid #ece8df;border-radius:4px;padding:8px 10px;grid-column:span 2;'
    summary_style = 'font-size:12px;color:#888;cursor:pointer;font-weight:500;list-style:none;'
    empty_style = 'font-size:12px;color:#bbb;padding:4px 0;'
    link_style = 'display:block;padding:3px 0;font-size:12px;color:#1a4d8c;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
    count = len(news_items) if news_items else 0
    if not news_items:
        return f'<div style="{tile_style}"><div style="{summary_style}">📰 今日相關新聞 (0)</div></div>'
    body = ''.join(
        f'<a href="{n["link"]}" target="_blank" rel="noopener" style="{link_style}">• {n["title"]}</a>'
        for n in news_items[:3]
    )
    # 預設收合,要點才展開
    return (
        f'<details style="{tile_style}">'
        f'<summary style="{summary_style}">📰 今日相關新聞 ({count}) ▸</summary>'
        f'<div style="margin-top:4px;">{body}</div>'
        f'</details>'
    )


def _news_tile(news_items):
    """Non-inline (desktop) version"""
    count = len(news_items) if news_items else 0
    if not news_items:
        return "<div class='tile news-tile'><div class='tile-title'>📰 今日相關新聞 (0)</div></div>"
    body = ''.join(
        f'<a href="{n["link"]}" target="_blank" rel="noopener" class="news-link">• {n["title"]}</a>'
        for n in news_items[:3]
    )
    return (
        f"<details class='tile news-tile'>"
        f"<summary class='tile-title'>📰 今日相關新聞 ({count}) ▸</summary>"
        f"<div>{body}</div>"
        f"</details>"
    )


def _tile_inline(title, stocks, key=None, empty_text='— 無 —'):
    tile_style = 'background:#fbfaf6;border:1px solid #ece8df;border-radius:4px;padding:8px 10px;'
    title_style = 'font-size:12px;color:#888;margin-bottom:4px;font-weight:500;'
    empty_style = 'font-size:12px;color:#bbb;padding:4px 0;'
    if not stocks:
        body = f'<div style="{empty_style}">{empty_text}</div>'
    else:
        body = ''.join(_stock_row_inline(s, highlight_key=key) for s in stocks)
    return f'<div style="{tile_style}"><div style="{title_style}">{title}</div>{body}</div>'


def _sector_card_inline(s):
    # 用市值加權當主要展示,等權重當對照
    main_chg = s['mcap_chg'] if s.get('mcap_chg') is not None else s['avg_chg']
    chg_color = _color(main_chg)
    bar_color = '#c0392b' if main_chg > 0 else '#27874b'
    width = min(abs(main_chg) * 15, 100)
    card_style = 'background:#fff;border:1px solid #e5e2dc;border-radius:6px;margin-bottom:10px;overflow:hidden;'
    summary_style = 'cursor:pointer;list-style:none;padding:12px 14px;'
    head_style = 'display:flex;justify-content:space-between;align-items:baseline;gap:12px;'
    name_style = 'font-size:15px;font-weight:600;color:#2a2a2a;'
    chg_style = f'font-size:15px;font-weight:600;font-variant-numeric:tabular-nums;color:{chg_color};'
    meta_style = 'font-size:11px;color:#888;margin-top:2px;'
    bar_bg = 'height:3px;background:#f0ede8;margin-top:8px;border-radius:2px;overflow:hidden;'
    bar_fg = f'background:{bar_color};height:100%;width:{width}%;border-radius:2px;'
    tiles_style = 'display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:8px;padding:0 14px 14px;'

    meta_parts = [f'成分 {s["n"]} 檔', f'漲 {s["advancers"]} 跌 {s["decliners"]}']
    if s.get('mcap_chg') is not None:
        meta_parts.append(f'等權 {s["avg_chg"]:+.2f}%')
    meta_parts.append(f'中位 {s["median_chg"]:+.2f}%')
    # 5d 輪動
    chg5 = s.get('mcap_chg_5d') if s.get('mcap_chg_5d') is not None else s.get('avg_chg_5d')
    if chg5 is not None:
        meta_parts.append(f'5日 {chg5:+.2f}%')
    # alpha vs TOPIX
    if s.get('alpha_pp') is not None:
        meta_parts.append(f'vs TOPIX {s["alpha_pp"]:+.2f}pp')
    meta_text = ' | '.join(meta_parts)

    tiles_html = (
        _tile_inline('🎯 權重貢獻 TOP3', s.get('top_contributors', []), key='contrib', empty_text='無市值資料') +
        _tile_inline('🔥 放量突破 TOP3', s['volume_breakouts'], key='vol', empty_text='今日無放量創新高') +
        _tile_inline('📈 漲幅 TOP3', s['top_gainers']) +
        _tile_inline('📉 跌幅 TOP3', s['top_losers']) +
        _tile_inline('💡 底部放量 TOP3', s['bottom_volume'], key='vol', empty_text='今日無底部放量') +
        _news_tile_inline(s.get('news', []))
    )
    summary_html = ''
    if s.get('summary'):
        summary_inline_style = 'background:#f5f2ea;border-left:3px solid #c2a25a;padding:6px 10px;margin:6px 14px 0;font-size:13px;color:#3a3a3a;border-radius:2px;'
        summary_html = f'<div style="{summary_inline_style}">💬 {s["summary"]}</div>'
    return (
        f'<details open style="{card_style}">'
        f'<summary style="{summary_style}">'
        f'<div style="{head_style}">'
        f'<div style="{name_style}">{s["sector"]}</div>'
        f'<div style="{chg_style}">{main_chg:+.2f}%</div>'
        f'</div>'
        f'<div style="{meta_style}">{meta_text}</div>'
        f'<div style="{bar_bg}"><div style="{bar_fg}"></div></div>'
        f'</summary>'
        f'{summary_html}'
        f'<div style="{tiles_style}">{tiles_html}</div>'
        f'</details>'
    )


def _index_card_inline(idx):
    chg_color = _color(idx['chg'])
    card_style = 'flex:1 1 140px;background:#fff;border:1px solid #e5e2dc;border-radius:6px;padding:10px 12px;'
    name_style = 'font-size:12px;color:#888;'
    close_style = 'font-size:18px;font-weight:600;margin-top:2px;color:#2a2a2a;'
    chg_style_block = f'font-size:13px;margin-top:2px;color:{chg_color};'
    if idx.get('is_breadth'):
        adv = idx.get('advancers', 0)
        dec = idx.get('decliners', 0)
        return (
            f'<div style="{card_style}">'
            f'<div style="{name_style}">{idx["name"]}</div>'
            f'<div style="font-size:18px;font-weight:600;margin-top:2px;color:{chg_color};">{idx["chg"]:+.2f}%</div>'
            f'<div style="font-size:13px;margin-top:2px;color:#2a2a2a;">'
            f'漲 <span style="color:#c0392b;">{adv}</span> / 跌 <span style="color:#27874b;">{dec}</span>'
            f'</div>'
            f'</div>'
        )
    return (
        f'<div style="{card_style}">'
        f'<div style="{name_style}">{idx["name"]}</div>'
        f'<div style="{close_style}">{idx["close"]:,.2f}</div>'
        f'<div style="{chg_style_block}">{idx["chg"]:+.2f}%</div>'
        f'</div>'
    )


def render_embed(indices, report):
    """嵌入 WordPress 用的版本 — 全 inline style。

    WP KSES 會剝掉 <style> 區塊和 <script>,快取 plugin 也會動 iframe
    屬性。唯一保證過關的是 inline `style=""` 屬性 (admins 可用)。"""
    gen_ts = datetime.now(JST).strftime('%Y-%m-%d %H:%M JST')
    data_date = indices[0]['date'].strftime('%Y-%m-%d') if indices else '—'
    idx_html = ''.join(_index_card_inline(i) for i in indices)
    sec_html = ''.join(_sector_card_inline(s) for s in report)
    # 關鍵:第一條 `.jpr-root, .jpr-root *` 用 all:revert 清掉主題繼承
    # 然後再 layer 我們的 styles,所有規則 !important 防主題 override
    # 全 inline style,用 font-family 在 root 繼承,其他要手動設
    # 注意:font-family 內用單引號 — 因為外層 style="" 是雙引號
    root_style = (
        "background:#fafaf7;color:#2a2a2a;"
        "font-family:-apple-system,'Hiragino Sans','Noto Sans JP','Noto Sans TC',sans-serif;"
        "font-size:14px;line-height:1.5;padding:16px;border-radius:6px;"
    )
    header_style = 'margin:0 0 16px;'
    h1_style = 'font-size:18px;margin:0 0 4px;font-weight:600;color:#2a2a2a;'
    sub_style = 'color:#888;font-size:12px;'
    indices_style = 'display:flex;gap:10px;margin:12px 0 20px;flex-wrap:wrap;'
    footer_style = 'margin-top:24px;padding-top:12px;border-top:1px solid #e5e2dc;color:#888;font-size:11px;'
    return f"""<div style="{root_style}">
<div style="{header_style}">
  <div style="{h1_style}">日股復盤・板塊地圖</div>
  <div style="{sub_style}">資料日 {data_date} | 產生時間 {gen_ts}</div>
  <div style="{indices_style}">{idx_html}</div>
</div>
<div>{sec_html}</div>
<div style="{footer_style}">資料來源 Yahoo Finance。板塊漲跌 = 該業種成分股等權重日漲幅平均。放量突破 = 量 ≥ 20MA × 2 且收盤創 20 日新高。底部放量 = 放量且距 120 日低點 10% 內。僅供參考,非投資建議。</div>
</div>"""


def _heat_color(chg):
    """JP convention: red=up, green=down.
    Return (bg hex, text color) based on chg in %.
    Range ~[-5, +5] mapped to gradient."""
    # clamp
    c = max(-5.0, min(5.0, chg))
    if c >= 0:
        # 0% → very pale red; +5% → deep red (127,29,29)
        t = c / 5.0  # 0..1
        r = int(255 - (255 - 127) * t)
        g = int(255 - (255 - 29) * t)
        b = int(255 - (255 - 29) * t)
        # interpolate from #fff5f5 (255,245,245) to #7f1d1d (127,29,29) -- use pale base
        r = int(255 - (255 - 127) * t)
        g = int(245 - (245 - 29) * t)
        b = int(245 - (245 - 29) * t)
    else:
        t = -c / 5.0  # 0..1
        # #f0fdf4 (240,253,244) → #14532d (20,83,45)
        r = int(240 - (240 - 20) * t)
        g = int(253 - (253 - 83) * t)
        b = int(244 - (244 - 45) * t)
    # text: dark if bg is pale; white if bg is deep
    pale = (r + g + b) / 3 > 180
    return f'rgb({r},{g},{b})', '#0f172a' if pale else '#fff'


def _us_stock_row(s, sector_weight_info=True):
    """US-style expandable stock row with tech + news inside."""
    chg = s.get('chg', 0)
    chg_cls = 'up' if chg > 0 else ('dn' if chg < 0 else 'mu')
    weight_span = ''
    if sector_weight_info and s.get('weight') is not None:
        w = s['weight'] * 100
        c = s.get('contribution_pp', 0)
        weight_span = (
            f'<span class="val mu" style="color:var(--mu);font-weight:500;min-width:82px;font-size:11px" '
            f'title="權重 {w:.2f}% / 貢獻 {c:+.2f}pp">貢 {c:+.2f}pp</span>'
        )
    vol_pill = ''
    vr = s.get('vol_ratio', 0)
    if vr >= 1.5:
        vol_pill = f' <span class="pill am" title="量/20日均 = {vr:.2f}">放量</span>'
    # Tech block
    rsi = s.get('rsi')
    rsi_txt = f'{rsi:.0f}'
    rsi_cls = 'mu'
    if rsi is not None:
        if rsi >= 70: rsi_cls, rsi_txt = 'dn', f'{rsi:.0f} 超買'
        elif rsi <= 30: rsi_cls, rsi_txt = 'up', f'{rsi:.0f} 超賣'
        else: rsi_cls, rsi_txt = 'mu', f'{rsi:.0f} 中性'
    else:
        rsi_txt, rsi_cls = '—', 'mu'
    chg_5d = s.get('chg_5d')
    chg5d_cls = 'up' if (chg_5d or 0) > 0 else ('dn' if (chg_5d or 0) < 0 else 'mu')
    chg5d_txt = f'{chg_5d:+.2f}%' if chg_5d is not None else '—'
    ma20 = s.get('ma20')
    ma_cls = 'up' if (ma20 and s.get('close', 0) > ma20) else 'dn'
    ma_arrow = '↑' if ma_cls == 'up' else '↓'
    tech = f"""<div class="sd-block">
  <h5>📐 技術面</h5>
  <div class="sd-grid">
    <div><span class="k">當日</span><span class="v {chg_cls}">{chg:+.2f}%</span></div>
    <div><span class="k">近 5 日</span><span class="v {chg5d_cls}">{chg5d_txt}</span></div>
    <div><span class="k">量 / 20 日均</span><span class="v {'am' if vr >= 1.5 else 'mu'}">×{vr:.2f}</span></div>
    <div><span class="k">RSI(14)</span><span class="v {rsi_cls}">{rsi_txt}</span></div>
  </div>
  <div class="sd-ma"><span class="pill {ma_cls}">MA20 {ma_arrow} ¥{ma20:,.0f}</span></div>
</div>"""
    # News block
    news_items = s.get('news', [])
    if news_items:
        news_rows = ''.join(
            f'<div class="sd-news-item"><a href="{n["link"]}" target="_blank" rel="noopener">{n["title"]}</a>'
            f'<div class="sd-news-meta">{n.get("date","")}</div></div>'
            for n in news_items[:3]
        )
    else:
        news_rows = '<div class="sd-news-meta">無當日相關新聞</div>'
    news = f'<div class="sd-block"><h5>📰 新聞</h5>{news_rows}</div>'
    return f"""<details class="stock-row"><summary class="row-item">
  <span class="sym">{s['code']}</span>
  <span class="nm" title="{s['name']}">{s['name']}{vol_pill}</span>
  <span class="val {chg_cls}">{chg:+.2f}%</span>{weight_span}
  <span class="chev">▾</span>
</summary><div class="stock-detail">{tech}{news}</div></details>"""


def _us_simple_row(s):
    """Compact row (no details, for 放量/底部 column)."""
    chg = s.get('chg', 0)
    chg_cls = 'up' if chg > 0 else ('dn' if chg < 0 else 'mu')
    vr = s.get('vol_ratio', 0)
    return f"""<div class="row-item">
  <span class="sym">{s['code']}</span>
  <span class="nm" title="{s['name']}">{s['name']}</span>
  <span class="val {chg_cls}">{chg:+.2f}%</span>
  <span class="val am" style="min-width:52px">×{vr:.2f}</span>
</div>"""


def _us_drill(r):
    """Drill-down panel per sector."""
    main_chg = r['mcap_chg'] if r.get('mcap_chg') is not None else r['avg_chg']
    chg_cls = 'up' if main_chg > 0 else ('dn' if main_chg < 0 else 'mu')
    # Meta line
    meta_parts = [f'成分 {r["n"]} 檔']
    if r.get('mcap_chg') is not None:
        meta_parts.append(f'等權 {r["avg_chg"]:+.2f}%')
    meta_parts.append(f'中位 {r["median_chg"]:+.2f}%')
    chg5 = r.get('mcap_chg_5d') or r.get('avg_chg_5d')
    if chg5 is not None:
        meta_parts.append(f'5日 {chg5:+.2f}%')
    if r.get('alpha_pp') is not None:
        meta_parts.append(f'vs TOPIX {r["alpha_pp"]:+.2f}pp')
    # AI summary
    summary_html = ''
    if r.get('summary'):
        summary_html = f'<div class="ai-summary">💬 {r["summary"]}</div>'
    # 4 columns:
    # pos contributors + gainers
    pos_contribs = [s for s in r.get('top_contributors', []) if s.get('contribution_pp', 0) >= 0][:3]
    neg_contribs = [s for s in r.get('top_contributors', []) if s.get('contribution_pp', 0) < 0][:3]
    # merge with top_gainers/losers excluding dupes
    pos_codes = {s['code'] for s in pos_contribs}
    neg_codes = {s['code'] for s in neg_contribs}
    extra_gainers = [s for s in r['top_gainers'] if s['code'] not in pos_codes and s['chg'] > 0][:3]
    extra_losers = [s for s in r['top_losers'] if s['code'] not in neg_codes and s['chg'] < 0][:3]
    pos_rows = ''.join(_us_stock_row(s) for s in pos_contribs + extra_gainers)
    neg_rows = ''.join(_us_stock_row(s) for s in neg_contribs + extra_losers)
    # 放量
    vol_items = r.get('volume_breakouts', []) + r.get('bottom_volume', [])
    seen_v = set()
    vol_rows_list = []
    for v in vol_items:
        if v['code'] not in seen_v:
            seen_v.add(v['code'])
            vol_rows_list.append(v)
    vol_rows = ''.join(_us_simple_row(v) for v in vol_rows_list[:5])
    if not vol_rows:
        vol_rows = '<div class="row-item" style="color:var(--mu);font-size:12px">今日無放量突破</div>'
    # 新聞 sector-level
    news_items = r.get('news', [])
    if news_items:
        news_rows = ''.join(
            f'<div class="news-item"><a href="{n["link"]}" target="_blank" rel="noopener">{n["title"]}</a>'
            f'<div class="meta">{n.get("date","")}</div></div>'
            for n in news_items[:4]
        )
    else:
        news_rows = '<div class="meta" style="padding:8px 0">暫無產業新聞</div>'
    return f"""<div class="drill" data-sector="{r['sector']}">
  <h2>{r['sector']} <span class="val {chg_cls}" style="font-size:18px;margin-left:8px">{main_chg:+.2f}%</span></h2>
  <div class="meta">{' · '.join(meta_parts)}</div>
  {summary_html}
  <div class="grid-4">
    <div class="col pos">
      <h4>▲ 今日推升</h4>
      {pos_rows or '<div class="row-item" style="color:var(--mu)">無</div>'}
    </div>
    <div class="col neg">
      <h4>▼ 今日拖累</h4>
      {neg_rows or '<div class="row-item" style="color:var(--mu)">無</div>'}
    </div>
    <div class="col vol">
      <h4>⚡ 放量 / 底部放量</h4>
      {vol_rows}
    </div>
    <div class="col news">
      <h4>📰 相關新聞</h4>
      {news_rows}
    </div>
  </div>
</div>"""


def _us_heatmap_card(r):
    main_chg = r['mcap_chg'] if r.get('mcap_chg') is not None else r['avg_chg']
    bg, txt = _heat_color(main_chg)
    vr_avg = None
    # summary meta: avg vol ratio of top contributors (approx)
    contribs = r.get('top_contributors', [])
    if contribs:
        vr_avg = sum(c.get('vol_ratio', 0) for c in contribs) / len(contribs)
    chg5 = r.get('mcap_chg_5d') or r.get('avg_chg_5d')
    sub_line = []
    if vr_avg is not None:
        sub_line.append(f'量×{vr_avg:.2f}')
    if chg5 is not None:
        arrow = '↑' if chg5 > 0 else '↓'
        sub_line.append(f'{arrow}5日 {chg5:+.2f}%')
    return f"""<div class="sector-card" data-sector="{r['sector']}" style="background:{bg};color:{txt}" onclick="selectSector('{r['sector']}')">
  <div class="name">{r['sector']}</div>
  <div class="chg">{main_chg:+.2f}%</div>
  <div class="sub">{' · '.join(sub_line)}</div>
</div>"""


def _us_market_card(idx):
    chg_cls = 'up' if idx['chg'] > 0 else ('dn' if idx['chg'] < 0 else 'mu')
    if idx.get('is_breadth'):
        adv = idx.get('advancers', 0)
        dec = idx.get('decliners', 0)
        return f"""<div class="mk">
  <div class="lbl">{idx['name']}</div>
  <div class="px {chg_cls}" style="font-size:20px">{idx['chg']:+.2f}%</div>
  <div class="chg">漲 <span class="up">{adv}</span> / 跌 <span class="dn">{dec}</span></div>
</div>"""
    return f"""<div class="mk">
  <div class="lbl">{idx['name']}</div>
  <div class="px">{idx['close']:,.2f}</div>
  <div class="chg {chg_cls}">{idx['chg']:+.2f}%</div>
</div>"""


def render_html(indices, report):
    gen_ts = datetime.now(JST).strftime('%Y-%m-%d %H:%M JST')
    data_date = indices[0]['date'].strftime('%Y-%m-%d') if indices else datetime.now(JST).strftime('%Y-%m-%d')
    idx_html = ''.join(_us_market_card(i) for i in indices)
    heat_html = ''.join(_us_heatmap_card(r) for r in report)
    drill_html = ''.join(_us_drill(r) for r in report)

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>日股板塊復盤 — {data_date}</title>
<style>
:root{{--bl:#2563eb;--gr:#16a34a;--rd:#dc2626;--am:#d97706;--bg:#f8fafc;--brd:#e2e8f0;--txt:#1e293b;--mu:#64748b;--card:#fff;--hover:#eef2ff;--up:#c0392b;--dn:#27874b}}
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Hiragino Sans','Noto Sans JP','Noto Sans TC',system-ui,sans-serif;background:var(--bg);color:var(--txt);font-size:14px;line-height:1.5}}
.hdr{{background:linear-gradient(135deg,#7f1d1d 0%,#991b1b 40%,#b91c1c 100%);color:#fff;padding:22px 28px 16px}}
.hdr h1{{font-size:20px;font-weight:800;margin-bottom:4px;letter-spacing:-0.3px}}
.hdr .sub{{font-size:11.5px;opacity:.9}}
.nav-link{{color:#fecaca;font-size:11.5px;text-decoration:none;margin-left:14px;border-bottom:1px dashed #fecaca;padding-bottom:1px}}
.nav-link:hover{{opacity:.7}}
.market{{display:flex;gap:10px;padding:14px 28px;background:var(--card);border-bottom:1px solid var(--brd);flex-wrap:wrap}}
.mk{{background:var(--bg);border:1px solid var(--brd);border-radius:8px;padding:10px 16px;min-width:140px}}
.mk .lbl{{font-size:10.5px;color:var(--mu);font-weight:600;letter-spacing:0.3px}}
.mk .px{{font-size:18px;font-weight:700;margin-top:2px}}
.mk .chg{{font-size:12px;font-weight:600;margin-top:1px}}
.chg.up,.val.up,.px.up{{color:var(--up)}}
.chg.dn,.val.dn,.px.dn{{color:var(--dn)}}
.chg.mu,.val.mu{{color:var(--mu)}}
.up{{color:var(--up)}}.dn{{color:var(--dn)}}
.pane{{padding:20px 28px}}
.ttl{{font-size:15px;font-weight:700;margin-bottom:4px}}
.desc{{font-size:12.5px;color:var(--mu);margin-bottom:16px;line-height:1.7}}
.heat{{display:grid;grid-template-columns:repeat(auto-fill,minmax(170px,1fr));gap:10px}}
.sector-card{{border-radius:10px;padding:14px 16px;cursor:pointer;position:relative;transition:transform .12s,box-shadow .12s;border:1px solid rgba(0,0,0,.08)}}
.sector-card:hover{{transform:translateY(-2px);box-shadow:0 6px 18px rgba(0,0,0,.14)}}
.sector-card.active{{outline:3px solid #111;outline-offset:2px}}
.sector-card .name{{font-size:15px;font-weight:800}}
.sector-card .chg{{font-size:22px;font-weight:800;margin-top:8px;letter-spacing:-0.5px}}
.sector-card .sub{{font-size:10.5px;opacity:.85;margin-top:3px}}
.drill{{margin-top:20px;padding:18px 20px;background:var(--card);border:1px solid var(--brd);border-radius:12px;display:none}}
.drill.active{{display:block}}
.drill h2{{font-size:17px;font-weight:800;margin-bottom:4px}}
.drill .meta{{font-size:12px;color:var(--mu);margin-bottom:12px;line-height:1.7}}
.ai-summary{{background:#fff7ed;border-left:3px solid #f59e0b;padding:10px 14px;margin-bottom:14px;font-size:13px;line-height:1.7;color:#78350f;border-radius:2px}}
.grid-4{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:14px}}
.col h4{{font-size:12px;font-weight:700;color:var(--mu);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;padding-bottom:6px;border-bottom:2px solid var(--brd)}}
.col.pos h4{{color:var(--up);border-bottom-color:#fecaca}}
.col.neg h4{{color:var(--dn);border-bottom-color:#bbf7d0}}
.col.vol h4{{color:var(--am);border-bottom-color:#fed7aa}}
.col.news h4{{color:var(--bl);border-bottom-color:#bfdbfe}}
.row-item{{display:flex;align-items:center;justify-content:space-between;padding:6px 0;border-bottom:1px dashed #eef2f7;font-size:12.5px}}
.row-item:last-child{{border-bottom:none}}
.row-item .sym{{font-weight:700;color:var(--txt);min-width:46px}}
.row-item .nm{{flex:1;color:var(--mu);font-size:11.5px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;margin:0 8px}}
.row-item .val{{font-weight:700;font-variant-numeric:tabular-nums;min-width:60px;text-align:right}}
.row-item .chev{{color:var(--mu);font-size:10px;margin-left:6px}}
.pill{{display:inline-block;padding:2px 8px;border-radius:10px;font-size:10.5px;font-weight:700}}
.pill.up{{background:#fee2e2;color:var(--up)}}
.pill.dn{{background:#dcfce7;color:var(--dn)}}
.pill.am{{background:#fef3c7;color:#b45309}}
details.stock-row{{border-bottom:1px dashed #eef2f7;margin:0}}
details.stock-row:last-child{{border-bottom:none}}
details.stock-row > summary{{list-style:none;cursor:pointer;padding:7px 0;display:flex;align-items:center;justify-content:space-between;font-size:12.5px;transition:background .1s}}
details.stock-row > summary::-webkit-details-marker{{display:none}}
details.stock-row > summary:hover{{background:#f8fafc}}
details.stock-row[open] > summary .chev{{transform:rotate(180deg)}}
.stock-detail{{background:#f8fafc;padding:10px 12px;border-radius:6px;margin:4px 0 8px}}
.sd-block{{padding:6px 0;border-bottom:1px dashed #e2e8f0}}
.sd-block:last-child{{border-bottom:none}}
.sd-block h5{{font-size:11.5px;font-weight:700;color:var(--mu);margin-bottom:6px;text-transform:uppercase;letter-spacing:0.4px}}
.sd-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:6px 14px;margin-bottom:6px}}
.sd-grid .k{{font-size:10.5px;color:var(--mu);display:block}}
.sd-grid .v{{font-size:12.5px;font-weight:700;font-variant-numeric:tabular-nums}}
.sd-ma{{margin-top:4px}}
.sd-news-item{{padding:4px 0;border-bottom:1px dashed #eef2f7;font-size:12px}}
.sd-news-item:last-child{{border-bottom:none}}
.sd-news-item a{{color:var(--txt);text-decoration:none}}
.sd-news-item a:hover{{color:var(--bl);text-decoration:underline}}
.sd-news-meta{{font-size:10.5px;color:var(--mu);margin-top:1px}}
.news-item{{padding:8px 0;border-bottom:1px dashed #eef2f7;font-size:13px}}
.news-item:last-child{{border-bottom:none}}
.news-item a{{color:var(--txt);text-decoration:none;line-height:1.5}}
.news-item a:hover{{color:var(--bl);text-decoration:underline}}
.news-item .meta{{font-size:10.5px;color:var(--mu);margin-top:2px}}
.ft{{text-align:center;color:var(--mu);font-size:11.5px;padding:20px 28px;border-top:1px solid var(--brd);line-height:1.8;background:var(--card)}}
</style>
</head>
<body>
<div class="hdr">
  <h1>日股板塊復盤</h1>
  <div class="sub">更新：{data_date}（每個交易日收盤後自動更新,JST 15:30）
    <a class="nav-link" href="https://tetsu811.github.io/cb-dashboard/us_index.html">→ 美股板塊</a>
    <a class="nav-link" href="https://tetsu811.github.io/cb-dashboard/etf_index.html">→ ETF 資金流向</a>
    <a class="nav-link" href="https://tetsu811.github.io/cb-dashboard/index.html">→ 可轉債儀表板</a>
  </div>
</div>
<div class="market">{idx_html}</div>
<div class="pane">
  <div class="ttl">板塊熱力圖</div>
  <div class="desc">東証 17 業種 × 469 檔大中型股 · 市值加權漲跌排序 · 顏色越深表示漲/跌幅越大(紅漲綠跌,日股慣例)。點擊任一板塊查看「為什麼」——包含推升股、拖累股、放量個股、AI 分析與相關新聞。</div>
  <div class="heat">{heat_html}</div>
  {drill_html}
</div>
<div class="ft">
  資料來源：Yahoo Finance (yfinance) + Google News (日本語) + Anthropic Claude  &nbsp;|&nbsp; 僅供研究參考,不構成投資建議  &nbsp;|&nbsp; 產生時間 {gen_ts}
</div>
<script>
function selectSector(sector){{
  document.querySelectorAll('.sector-card').forEach(function(c){{
    c.classList.toggle('active', c.dataset.sector===sector);
  }});
  document.querySelectorAll('.drill').forEach(function(d){{
    d.classList.toggle('active', d.dataset.sector===sector);
  }});
  var drill = document.querySelector('.drill.active');
  if(drill) drill.scrollIntoView({{behavior:'smooth', block:'start'}});
}}
// auto-select first sector on load
window.addEventListener('DOMContentLoaded', function(){{
  var first = document.querySelector('.sector-card');
  if(first) selectSector(first.dataset.sector);
}});
</script>
</body>
</html>
"""


def synthesize_breadth(report):
    """所有成分股等權平均 + 漲跌家數 — 當 N225 / TOPIX 的對照組。"""
    all_chgs = []
    advancers = decliners = unchanged = 0
    last_date = None
    for sec in report:
        for src in (sec['top_gainers'], sec['top_losers']):
            for s in src:
                last_date = s.get('date', last_date)
        # use full sector lists via re-traverse: top_gainers/losers only have 3 each
    # Better: compute from advancers/decliners totals already in report
    for sec in report:
        advancers += sec['advancers']
        decliners += sec['decliners']
        # avg_chg already weighted within sector; recompute total via n*avg_chg
        all_chgs.extend([sec['avg_chg']] * sec['n'])
    if not all_chgs:
        return None
    import numpy as np
    return {
        'name': '成分股等權平均',
        'code': 'BREADTH',
        'close': float(np.mean(all_chgs)),  # actually % change, label as such in card
        'chg': float(np.mean(all_chgs)),
        'date': last_date or datetime.now(JST).date(),
        'is_breadth': True,
        'advancers': advancers,
        'decliners': decliners,
    }


def main():
    sectors = load_sectors()
    ticker_map = build_ticker_map(sectors)
    print(f"Sectors: {len([k for k in sectors if not k.startswith('_') and k != '指数'])}, tickers: {len(ticker_map)}")
    idx_data = fetch_indices(sectors.get('指数', []))
    df = fetch_all(list(ticker_map.keys()))
    mcap_map = load_market_caps(list(ticker_map.keys()))
    print(f"Market caps loaded: {len(mcap_map)}/{len(ticker_map)}")
    report = build_sector_report(ticker_map, df, mcap_map=mcap_map)
    # #4 — alpha vs TOPIX (use TOPIX 1306 ETF as benchmark)
    topix = next((i for i in idx_data if 'TOPIX (1306)' in i.get('name', '')), None)
    topix_chg = topix['chg'] if topix else None
    for r in report:
        if topix_chg is not None and r.get('mcap_chg') is not None:
            r['alpha_pp'] = r['mcap_chg'] - topix_chg
        else:
            r['alpha_pp'] = None
    # B. News per sector + per top-contributor stock
    print("Fetching news...")
    news_map = fetch_news_all([r['sector'] for r in report])
    for r in report:
        r['news'] = news_map.get(r['sector'], [])
    print(f"News fetched: {sum(len(r['news']) for r in report)} items across {sum(1 for r in report if r['news'])} sectors")
    fetch_stock_news_for_report(report)
    # C. LLM summaries (optional, requires ANTHROPIC_API_KEY)
    import os
    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if api_key:
        print("Generating LLM summaries...")
        summaries = fetch_summaries(report, api_key)
        for r in report:
            r['summary'] = summaries.get(r['sector'], '')
        print(f"Summaries generated: {sum(1 for r in report if r.get('summary'))}/{len(report)}")
    else:
        print("ANTHROPIC_API_KEY not set — skipping LLM summaries")
        for r in report:
            r['summary'] = ''
    ok = sum(r['n'] for r in report)
    print(f"Processed {ok}/{len(ticker_map)} tickers across {len(report)} sectors")
    breadth = synthesize_breadth(report)
    if breadth:
        idx_data.append(breadth)
    html = render_html(idx_data, report)
    OUTPUT_HTML.write_text(html, encoding='utf-8')
    print(f"✓ Wrote {OUTPUT_HTML}")
    embed = render_embed(idx_data, report)
    OUTPUT_EMBED.write_text(embed, encoding='utf-8')
    print(f"✓ Wrote {OUTPUT_EMBED}")


if __name__ == '__main__':
    main()
