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


# ---------- C. Claude LLM one-sentence summaries ----------

CLAUDE_MODEL = 'claude-sonnet-4-5'


def build_summary_prompt(report):
    """把 17 個板塊的今日資料組成一個 prompt,一次叫 Claude 產出所有 summary。"""
    lines = [
        "你是日股分析師。下面是今日東証 17 業種表現。每個板塊給我一句 30 字內的繁體中文解釋,專注於「為什麼今天動這樣」— 可從新聞、領漲/領跌個股、市值權重貢獻推論。",
        "沒明確原因就寫「廣泛式,無具體催化劑」。",
        "",
        "輸出格式(一行一個板塊,嚴格遵守):",
        "板塊名|解釋句",
        "",
        "---資料---",
    ]
    for r in report:
        main_chg = r['mcap_chg'] if r.get('mcap_chg') is not None else r['avg_chg']
        contrib_str = ', '.join(
            f"{c['name']}{c['chg']:+.1f}%(貢獻{c['contribution_pp']:+.2f}pp)"
            for c in r.get('top_contributors', [])[:3]
        )
        gainer_str = ', '.join(f"{g['name']}{g['chg']:+.1f}%" for g in r['top_gainers'][:2])
        loser_str = ', '.join(f"{l['name']}{l['chg']:+.1f}%" for l in r['top_losers'][:2])
        news_str = '; '.join(n['title'][:60] for n in r.get('news', [])[:3])
        lines.append(f"\n【{r['sector']}】{main_chg:+.2f}% (成分 {r['n']} 檔)")
        if contrib_str:
            lines.append(f"  權重主導: {contrib_str}")
        if gainer_str:
            lines.append(f"  領漲: {gainer_str}")
        if loser_str:
            lines.append(f"  領跌: {loser_str}")
        if news_str:
            lines.append(f"  新聞: {news_str}")
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
        else:
            top_contributors = []
            mcap_weighted_chg = None
        report.append({
            'sector': sector,
            'avg_chg': float(chgs.mean()),
            'mcap_chg': mcap_weighted_chg,
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
    return (
        f"<div class='stock-row'>"
        f"<span class='code'>{s['code']}</span>"
        f"<span class='name'>{s['name']}</span>"
        f"<span class='chg {chg_cls}'>{chg_str}</span>"
        f"{extra}"
        f"</div>"
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
        # contribution 以 pp 顯示(百分點),e.g. -0.31pp 代表把板塊拉低 0.31 個百分點
        extra = f'<div style="{_EXTRA_STYLE}">{contrib:+.2f}pp 權{weight:.0f}%</div>'
    else:
        extra = '<div></div>'
    return (
        f'<div style="{_ROW_STYLE}">'
        f'<div style="{_CODE_STYLE}">{s["code"]}</div>'
        f'<div style="{_NAME_STYLE}">{s["name"]}</div>'
        f'<div style="{_CHG_STYLE}color:{chg_color};">{chg_str}</div>'
        f'{extra}'
        f'</div>'
    )


def _news_tile_inline(news_items):
    tile_style = 'background:#fbfaf6;border:1px solid #ece8df;border-radius:4px;padding:8px 10px;grid-column:span 2;'
    title_style = 'font-size:12px;color:#888;margin-bottom:4px;font-weight:500;'
    empty_style = 'font-size:12px;color:#bbb;padding:4px 0;'
    link_style = 'display:block;padding:3px 0;font-size:12px;color:#1a4d8c;text-decoration:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
    if not news_items:
        body = f'<div style="{empty_style}">暫無相關新聞</div>'
    else:
        body = ''.join(
            f'<a href="{n["link"]}" target="_blank" rel="noopener" style="{link_style}">• {n["title"]}</a>'
            for n in news_items[:3]
        )
    return f'<div style="{tile_style}"><div style="{title_style}">📰 今日相關新聞</div>{body}</div>'


def _news_tile(news_items):
    """Non-inline (desktop) version"""
    if not news_items:
        body = "<div class='tile-empty'>暫無相關新聞</div>"
    else:
        body = ''.join(
            f'<a href="{n["link"]}" target="_blank" rel="noopener" class="news-link">• {n["title"]}</a>'
            for n in news_items[:3]
        )
    return f"<div class='tile news-tile'><div class='tile-title'>📰 今日相關新聞</div>{body}</div>"


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
    else:
        meta_parts.append(f'中位 {s["median_chg"]:+.2f}%')
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


def render_html(indices, report):
    gen_ts = datetime.now(JST).strftime('%Y-%m-%d %H:%M JST')
    data_date = indices[0]['date'].strftime('%Y-%m-%d') if indices else '—'
    idx_html = ''.join(_index_card(i) for i in indices)
    sec_html = ''.join(_sector_card(s) for s in report)

    return f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>日股復盤 {data_date}</title>
<style>
  :root {{
    --bg: #fafaf7;
    --card: #ffffff;
    --border: #e5e2dc;
    --text: #2a2a2a;
    --muted: #888;
    --up: #c0392b;
    --down: #27874b;
    --accent: #1a4d8c;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0;
    padding: 16px;
    font-family: -apple-system, "Hiragino Sans", "Noto Sans JP", "Noto Sans TC", sans-serif;
    background: var(--bg);
    color: var(--text);
    font-size: 14px;
    line-height: 1.5;
  }}
  header {{ margin-bottom: 16px; }}
  h1 {{
    font-size: 18px;
    margin: 0 0 4px;
    font-weight: 600;
    letter-spacing: 0.02em;
  }}
  .sub {{ color: var(--muted); font-size: 12px; }}
  .indices {{
    display: flex;
    gap: 10px;
    margin: 12px 0 20px;
    flex-wrap: wrap;
  }}
  .index-card {{
    flex: 1 1 140px;
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 10px 12px;
  }}
  .idx-name {{ font-size: 12px; color: var(--muted); }}
  .idx-close {{ font-size: 18px; font-weight: 600; margin-top: 2px; }}
  .idx-chg {{ font-size: 13px; margin-top: 2px; }}
  .sector-card {{
    background: var(--card);
    border: 1px solid var(--border);
    border-radius: 6px;
    margin-bottom: 10px;
    overflow: hidden;
  }}
  .sector-summary {{
    cursor: pointer;
    list-style: none;
    padding: 12px 14px;
    user-select: none;
  }}
  .sector-summary::-webkit-details-marker {{ display: none; }}
  .sec-head {{
    display: flex;
    justify-content: space-between;
    align-items: baseline;
    gap: 12px;
  }}
  .sec-name {{ font-size: 15px; font-weight: 600; }}
  .sec-chg {{ font-size: 15px; font-weight: 600; font-variant-numeric: tabular-nums; }}
  .sec-meta {{ font-size: 11px; color: var(--muted); margin-top: 2px; }}
  .sec-bar {{
    height: 3px;
    background: #f0ede8;
    margin-top: 8px;
    border-radius: 2px;
    overflow: hidden;
    position: relative;
  }}
  .bar-up {{ background: var(--up); height: 100%; border-radius: 2px; }}
  .bar-down {{ background: var(--down); height: 100%; border-radius: 2px; }}
  .tiles {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 8px;
    padding: 0 14px 14px;
  }}
  .tile {{
    background: #fbfaf6;
    border: 1px solid #ece8df;
    border-radius: 4px;
    padding: 8px 10px;
  }}
  .tile-title {{
    font-size: 12px;
    color: var(--muted);
    margin-bottom: 4px;
    font-weight: 500;
  }}
  .tile-empty {{ font-size: 12px; color: #bbb; padding: 4px 0; }}
  .stock-row {{
    display: grid;
    grid-template-columns: 48px 1fr auto auto;
    gap: 6px;
    padding: 3px 0;
    font-size: 13px;
    align-items: baseline;
  }}
  .stock-row .code {{ color: var(--muted); font-variant-numeric: tabular-nums; font-size: 12px; }}
  .stock-row .name {{ overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
  .stock-row .chg {{ font-variant-numeric: tabular-nums; font-weight: 500; }}
  .stock-row .extra {{ font-size: 11px; color: var(--muted); min-width: 52px; text-align: right; }}
  .up {{ color: var(--up); }}
  .down {{ color: var(--down); }}
  footer {{
    margin-top: 24px;
    padding-top: 12px;
    border-top: 1px solid var(--border);
    color: var(--muted);
    font-size: 11px;
  }}
  .sec-summary {{
    background: #f5f2ea;
    border-left: 3px solid #c2a25a;
    padding: 6px 10px;
    margin: 6px 14px 0;
    font-size: 13px;
    color: #3a3a3a;
    border-radius: 2px;
  }}
  .news-link {{
    display: block;
    padding: 3px 0;
    font-size: 12px;
    color: var(--accent);
    text-decoration: none;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }}
  .news-link:hover {{ text-decoration: underline; }}
  .news-tile {{ grid-column: span 2; }}
</style>
</head>
<body>
<header>
  <h1>日股復盤・板塊地圖</h1>
  <div class="sub">資料日 {data_date} | 產生時間 {gen_ts}</div>
  <div class="indices">{idx_html}</div>
</header>
<main>
  {sec_html}
</main>
<footer>
  資料來源 Yahoo Finance。板塊漲跌為該板塊成分股等權重日漲幅平均。放量突破 = 成交量 ≥ 20 日均量 2 倍 且 收盤創 20 日新高。底部放量 = 放量且價位距區間低點 10% 以內。僅供參考,非投資建議。
</footer>
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
    # B. News per sector
    print("Fetching news...")
    news_map = fetch_news_all([r['sector'] for r in report])
    for r in report:
        r['news'] = news_map.get(r['sector'], [])
    print(f"News fetched: {sum(len(r['news']) for r in report)} items across {sum(1 for r in report if r['news'])} sectors")
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
